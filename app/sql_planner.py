from __future__ import annotations

import json
from typing import TYPE_CHECKING

from ollama import AsyncClient
from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator

from app.excel_planner import CellValue, FilterOperator
from app.intent import Action, IntentDraft
from app.sql_service import SqlTableMetadata, SqliteMetadata

if TYPE_CHECKING:
    from app.config import Settings


SUPPORTED_SQL_ACTIONS = {Action.SELECT, Action.UPDATE_ROWS, Action.DELETE_ROWS}


class SqlColumnRef(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str
    declared_type: str = ""


class SqlFilterCondition(BaseModel):
    model_config = ConfigDict(extra="forbid")

    column: SqlColumnRef
    operator: FilterOperator
    value: CellValue = None

    @model_validator(mode="after")
    def validate_value(self) -> "SqlFilterCondition":
        no_value = {FilterOperator.IS_EMPTY, FilterOperator.NOT_EMPTY}
        if self.operator in no_value and self.value is not None:
            raise ValueError("For is_empty/not_empty, value must be null.")
        if self.operator not in no_value and self.value is None:
            raise ValueError(f"Operator {self.operator.value} requires a value.")
        return self


class SqlAssignment(BaseModel):
    model_config = ConfigDict(extra="forbid")

    column: SqlColumnRef
    value: CellValue = None


class SqlOperationPlan(BaseModel):
    model_config = ConfigDict(extra="forbid")

    action: Action
    resolved: bool
    table_name: str | None = None
    filters: list[SqlFilterCondition] = Field(default_factory=list)
    selected_columns: list[SqlColumnRef] = Field(default_factory=list)
    assignments: list[SqlAssignment] = Field(default_factory=list)
    confidence: float = Field(ge=0.0, le=1.0)
    resolution_note: str
    alternative_matches: list[str] = Field(default_factory=list)
    clarification_question: str | None = None

    @model_validator(mode="after")
    def validate_shape(self) -> "SqlOperationPlan":
        if not self.resolved:
            return self
        if self.action not in SUPPORTED_SQL_ACTIONS:
            raise ValueError("This SQL action is not supported.")
        if not self.table_name:
            raise ValueError("An exact table is required.")
        if self.action in {Action.UPDATE_ROWS, Action.DELETE_ROWS} and not self.filters:
            raise ValueError("SQL UPDATE/DELETE without a WHERE-like filter is not allowed.")
        if self.action == Action.UPDATE_ROWS and not self.assignments:
            raise ValueError("UPDATE requires new values.")
        return self


class SqlRequiredAssignmentPlan(BaseModel):
    """Catalog-bound recovery for UPDATE values already identified by the AI intent layer."""

    model_config = ConfigDict(extra="forbid")

    resolved: bool
    table_name: str | None = None
    assignments: list[SqlAssignment] = Field(default_factory=list)
    confidence: float = Field(ge=0.0, le=1.0)
    resolution_note: str
    clarification_question: str | None = None

    @model_validator(mode="after")
    def validate_shape(self) -> "SqlRequiredAssignmentPlan":
        if self.resolved and not self.assignments:
            raise ValueError(
                "A resolved UPDATE assignment recovery requires at least one assignment."
            )
        return self


class SqlRequiredFilterPlan(BaseModel):
    """Catalog-bound recovery for row filters already identified by the AI intent layer."""

    model_config = ConfigDict(extra="forbid")

    resolved: bool
    table_name: str | None = None
    filters: list[SqlFilterCondition] = Field(default_factory=list)
    confidence: float = Field(ge=0.0, le=1.0)
    resolution_note: str
    clarification_question: str | None = None

    @model_validator(mode="after")
    def validate_shape(self) -> "SqlRequiredFilterPlan":
        if self.resolved and not self.filters:
            raise ValueError("A resolved filter recovery requires at least one filter.")
        return self


SQL_PLANNER_PROMPT = """
Ты — строгий SQL-планировщик DataOps Commander для реальной SQLite-базы.
Верни только структурированный план по JSON Schema. Никогда не генерируй raw SQL.

Разрешены:
- select: показать строки;
- update_rows: изменить значения только по обязательному фильтру;
- delete_rows: удалить строки только по обязательному фильтру.

Правила:
1. action строго совпадает с parsed_intent.action.
   SQLite-файл уже выбран пользователем и является активным: никогда не
   спрашивай, какой источник или файл использовать.
2. table_name и каждый столбец копируй точно из каталога.
3. Нельзя выдумывать таблицы, столбцы, фильтры или значения.
4. Все filters объединяются через AND.
5. UPDATE и DELETE без фильтра: resolved=false.
   Если parsed_intent.filter_hints НЕ пустой, каждый пользовательский критерий
   отбора должен быть отражён в filters. Нельзя возвращать filters=[] для SELECT,
   если пользователь явно запросил условие (например, stock < 5).
6. Для update_rows assignments ОБЯЗАТЕЛЕН и должен содержать каждое новое
   значение как точную пару column + value. Нельзя оставлять assignments=[] и
   нельзя прятать новое значение только в parsed_intent.value_hints.
7. selected_columns=[] означает вывести все столбцы.
8. Для eq/ne/contains/not_contains/gt/gte/lt/lte/older_than_days нужен value.
9. Для is_empty/not_empty value=null.
10. Понимай опечатки, сленг, русский/английский и очевидные синонимы. Если
   таблица одна — выбирай её. Если один реальный столбец очевидно подходит по
   смыслу — выбирай его.
11. Только если два реально существующих объекта одинаково подходят или не
   указано критичное значение, resolved=false и задай один короткий вопрос.
12. resolution_note — короткое объяснение совпадения без скрытых рассуждений.
13. User-facing fields clarification_question and resolution_note must always be in English. Preserve exact table names, column names and data values from the source.
""".strip()


SQL_ASSIGNMENT_RECOVERY_PROMPT = """
Ты — узкий UPDATE Assignment Resolver для DataOps Commander.

AI Command Interpreter уже понял, что пользователь хочет UPDATE, и выделил
column_hints / value_hints / normalized_request. Твоя единственная задача —
привязать НОВЫЕ записываемые значения к точным столбцам SQLite schema и вернуть
structured assignments. Никаких raw SQL, никаких filters и никаких новых действий.

Правила:
1. resolved=true допустим только если каждое требуемое новое значение можно
   однозначно привязать к существующему столбцу.
2. table_name и column.name копируй точно из sqlite_catalog.
3. Если expected_table_name задан и существует в sqlite_catalog — используй его.
4. Не путай значение фильтра с новым записываемым значением.
   Например: «поставь остаток 77 для категории Периферия» означает assignment
   stock=77; «Периферия» относится к row filter и НЕ является assignment.
5. Используй user_command, normalized_request, column_hints, filter_hints и
   value_hints только как уже выделенный AI-контекст. Не добавляй новых действий.
6. Не придумывай таблицы, столбцы или значения.
7. Если безопасно привязать новое значение невозможно, resolved=false и один
   короткий clarification_question на английском.
8. resolution_note и clarification_question всегда на английском. Точные имена
   таблиц, столбцов и значения из данных не переводи.
9. Верни только JSON по schema.
""".strip()


SQL_FILTER_RECOVERY_PROMPT = """
Ты — узкий Filter Resolver для DataOps Commander.

AI Command Interpreter уже понял пользовательский смысл и выделил
parsed_intent.filter_hints. Твоя единственная задача — привязать ЭТИ уже
выделенные условия к точной SQLite schema и вернуть структурированные filters.
Никаких raw SQL и никаких новых условий.

Правила:
1. Если required_filter_hints не пустой, resolved=true допустим только когда
   КАЖДОЕ пользовательское условие отражено в filters.
2. table_name и column.name копируй точно из sqlite_catalog.
3. Операторы: eq, ne, contains, not_contains, gt, gte, lt, lte,
   is_empty, not_empty, older_than_days.
4. Не теряй числовые границы и направление сравнения: «меньше 5» = lt 5,
   «не меньше 5» = gte 5, «больше 5» = gt 5 и т.д.
5. Если таблица одна или resource_name_hint однозначно совпадает — выбери её.
6. Если column_hints однозначно совпадает с реальным столбцом — используй его.
7. Не придумывай таблицы, столбцы, значения или дополнительные фильтры.
8. Если безопасно привязать условие невозможно, resolved=false и один короткий
   clarification_question на английском.
9. resolution_note и clarification_question всегда на английском. Точные имена
   таблиц, столбцов и значения из данных не переводи.
10. Верни только JSON по schema.
""".strip()


def _catalog(metadata: SqliteMetadata) -> dict:
    return {
        "database": metadata.original_name,
        "tables": [
            {
                "name": table.name,
                "row_count": table.row_count,
                "supports_rowid": table.supports_rowid,
                "foreign_key_count": table.foreign_key_count,
                "referenced_by_foreign_keys": table.referenced_by_foreign_keys,
                "trigger_count": table.trigger_count,
                "columns": [column.model_dump(mode="json") for column in table.columns],
            }
            for table in metadata.tables
        ],
    }


def _table(metadata: SqliteMetadata, name: str) -> SqlTableMetadata:
    table = next((item for item in metadata.tables if item.name == name), None)
    if table is None:
        raise ValueError(f"Table “{name}” is not present in the catalog.")
    return table


def _validate_column(table: SqlTableMetadata, ref: SqlColumnRef) -> None:
    actual = next((item for item in table.columns if item.name == ref.name), None)
    if actual is None:
        raise ValueError(f"Column “{ref.name}” does not match the catalog.")
    if ref.declared_type and actual.declared_type != ref.declared_type:
        raise ValueError(f"Column type for “{ref.name}” does not match the catalog.")
    ref.declared_type = actual.declared_type


def validate_sql_plan(
    plan: SqlOperationPlan,
    intent: IntentDraft,
    metadata: SqliteMetadata,
) -> SqlOperationPlan:
    if plan.action != intent.action:
        raise ValueError("The SQL planner changed the action.")
    if not plan.resolved:
        return plan
    if intent.filter_hints and not plan.filters:
        raise ValueError(
            "The SQL planner dropped row filters that were already present in "
            "parsed_intent.filter_hints."
        )
    table = _table(metadata, plan.table_name or "")
    for ref in [
        *plan.selected_columns,
        *(item.column for item in plan.filters),
        *(item.column for item in plan.assignments),
    ]:
        _validate_column(table, ref)
    return plan



def validate_required_assignment_plan(
    recovered: SqlRequiredAssignmentPlan,
    intent: IntentDraft,
    metadata: SqliteMetadata,
    *,
    expected_table_name: str | None = None,
) -> SqlRequiredAssignmentPlan:
    if intent.action != Action.UPDATE_ROWS:
        return recovered
    if not recovered.resolved:
        return recovered
    if not recovered.assignments:
        raise ValueError("UPDATE assignment recovery returned no assignments.")

    table_name = recovered.table_name
    if not table_name and expected_table_name:
        _table(metadata, expected_table_name)
        table_name = expected_table_name
        recovered = recovered.model_copy(update={"table_name": table_name})

    if not table_name:
        raise ValueError("UPDATE assignment recovery did not provide an exact table.")

    table = _table(metadata, table_name)
    for assignment in recovered.assignments:
        _validate_column(table, assignment.column)
    return recovered


def _merge_recovered_assignments_payload(
    payload: dict,
    recovered: SqlRequiredAssignmentPlan,
) -> dict:
    if not recovered.resolved or not recovered.assignments:
        raise ValueError(
            "The AI could not safely bind the requested UPDATE value to the SQLite schema."
        )
    merged = dict(payload)
    if recovered.table_name:
        merged["table_name"] = recovered.table_name
    merged["assignments"] = [
        item.model_dump(mode="json") for item in recovered.assignments
    ]
    return merged


def _candidate_object(candidate: str) -> dict:
    payload = json.loads(candidate)
    if not isinstance(payload, dict):
        raise ValueError("SQL planner response must be a JSON object.")
    return payload


def _is_missing_update_assignments_error(
    error: Exception,
    intent: IntentDraft,
) -> bool:
    return (
        intent.action == Action.UPDATE_ROWS
        and isinstance(error, ValidationError)
        and "UPDATE requires new values." in str(error)
    )


def validate_required_filter_plan(
    recovered: SqlRequiredFilterPlan,
    intent: IntentDraft,
    metadata: SqliteMetadata,
    *,
    expected_table_name: str | None = None,
) -> SqlRequiredFilterPlan:
    if not intent.filter_hints:
        return recovered
    if not recovered.resolved:
        return recovered
    if not recovered.filters:
        raise ValueError("Filter recovery returned no row filters.")

    table_name = recovered.table_name
    if not table_name and expected_table_name:
        _table(metadata, expected_table_name)
        table_name = expected_table_name
        recovered = recovered.model_copy(update={"table_name": table_name})

    if not table_name:
        raise ValueError("Filter recovery did not provide an exact table.")

    table = _table(metadata, table_name)
    for condition in recovered.filters:
        _validate_column(table, condition.column)
    return recovered


def _merge_recovered_filters_payload(
    payload: dict,
    recovered: SqlRequiredFilterPlan,
) -> dict:
    if not recovered.resolved or not recovered.filters:
        raise ValueError(
            "The AI could not safely bind the requested row filter to the SQLite schema."
        )
    merged = dict(payload)
    if recovered.table_name:
        merged["table_name"] = recovered.table_name
    merged["filters"] = [
        item.model_dump(mode="json") for item in recovered.filters
    ]
    return merged


def _merge_recovered_filters(
    plan: SqlOperationPlan,
    recovered: SqlRequiredFilterPlan,
    intent: IntentDraft,
    metadata: SqliteMetadata,
) -> SqlOperationPlan:
    if not recovered.resolved or not recovered.table_name or not recovered.filters:
        raise ValueError(
            "The AI could not safely bind the requested row filter to the SQLite schema."
        )
    # The recovery model is another structured AI result, scoped only to filters.
    # Python does not reinterpret user language here; it only applies the already
    # resolved catalog-bound filter objects deterministically.
    merged = plan.model_copy(
        update={
            "table_name": recovered.table_name,
            "filters": recovered.filters,
        }
    )
    return validate_sql_plan(merged, intent, metadata)


def _is_dropped_filter_error(error: Exception) -> bool:
    message = str(error)
    return (
        (isinstance(error, ValueError) and "dropped row filters" in message)
        or (
            isinstance(error, ValidationError)
            and "SQL UPDATE/DELETE without a WHERE-like filter is not allowed." in message
        )
    )


def _promote_complete_unresolved_plan(
    plan: SqlOperationPlan,
    intent: IntentDraft,
    metadata: SqliteMetadata,
) -> SqlOperationPlan | None:
    if plan.resolved:
        return plan
    payload = plan.model_dump(mode="python")
    payload["resolved"] = True
    if not payload.get("table_name") and len(metadata.tables) == 1:
        payload["table_name"] = metadata.tables[0].name
    try:
        promoted = SqlOperationPlan.model_validate(payload)
        return validate_sql_plan(promoted, intent, metadata)
    except (ValidationError, ValueError):
        return None


class SqlPlanner:
    def __init__(self, settings: Settings) -> None:
        self._model = settings.ollama_model
        self._client = AsyncClient(host=settings.ollama_host, timeout=120.0)

    async def _recover_required_assignments(
        self,
        user_text: str,
        intent: IntentDraft,
        metadata: SqliteMetadata,
        *,
        expected_table_name: str | None = None,
    ) -> SqlRequiredAssignmentPlan:
        payload = {
            "user_command": user_text,
            "normalized_request": intent.normalized_request,
            "action": intent.action.value,
            "resource_name_hint": intent.resource_name_hint,
            "column_hints": intent.column_hints,
            "filter_hints": intent.filter_hints,
            "value_hints": intent.value_hints,
            "sqlite_catalog": _catalog(metadata),
            "expected_table_name": expected_table_name,
        }
        messages = [
            {"role": "system", "content": SQL_ASSIGNMENT_RECOVERY_PROMPT},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False, indent=2)},
        ]
        candidate = ""
        last_error: ValidationError | ValueError | None = None
        for attempt in range(3):
            response = await self._client.chat(
                model=self._model,
                messages=(
                    messages
                    if attempt == 0
                    else [
                        *messages,
                        {"role": "assistant", "content": candidate},
                        {
                            "role": "user",
                            "content": (
                                f"UPDATE assignment recovery validation failed: {last_error}. "
                                "Repair only the structured column/value assignments. "
                                "Do not create filters or raw SQL. Use exact sqlite_catalog "
                                "column names and return only JSON."
                            ),
                        },
                    ]
                ),
                format=SqlRequiredAssignmentPlan.model_json_schema(),
                options={"temperature": 0},
                think=False,
            )
            candidate = response.message.content or ""
            try:
                return validate_required_assignment_plan(
                    SqlRequiredAssignmentPlan.model_validate_json(candidate),
                    intent,
                    metadata,
                    expected_table_name=expected_table_name,
                )
            except (ValidationError, ValueError) as error:
                last_error = error
        if last_error is not None:
            raise last_error
        raise ValueError("SQL UPDATE assignment recovery returned no validated result.")

    async def _recover_required_filters(
        self,
        user_text: str,
        intent: IntentDraft,
        metadata: SqliteMetadata,
        *,
        expected_table_name: str | None = None,
    ) -> SqlRequiredFilterPlan:
        payload = {
            "user_command": user_text,
            "normalized_request": intent.normalized_request,
            "action": intent.action.value,
            "resource_name_hint": intent.resource_name_hint,
            "column_hints": intent.column_hints,
            "required_filter_hints": intent.filter_hints,
            "value_hints": intent.value_hints,
            "sqlite_catalog": _catalog(metadata),
            "expected_table_name": expected_table_name,
        }
        messages = [
            {"role": "system", "content": SQL_FILTER_RECOVERY_PROMPT},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False, indent=2)},
        ]
        candidate = ""
        last_error: ValidationError | ValueError | None = None
        for attempt in range(3):
            response = await self._client.chat(
                model=self._model,
                messages=(
                    messages
                    if attempt == 0
                    else [
                        *messages,
                        {"role": "assistant", "content": candidate},
                        {
                            "role": "user",
                            "content": (
                                f"Filter recovery validation failed: {last_error}. "
                                "Repair only the structured filter binding. Preserve every "
                                "required_filter_hint, use exact sqlite_catalog names, and "
                                "return only JSON."
                            ),
                        },
                    ]
                ),
                format=SqlRequiredFilterPlan.model_json_schema(),
                options={"temperature": 0},
                think=False,
            )
            candidate = response.message.content or ""
            try:
                return validate_required_filter_plan(
                    SqlRequiredFilterPlan.model_validate_json(candidate),
                    intent,
                    metadata,
                    expected_table_name=expected_table_name,
                )
            except (ValidationError, ValueError) as error:
                last_error = error
        if last_error is not None:
            raise last_error
        raise ValueError("SQL filter recovery returned no validated result.")

    async def resolve(
        self,
        user_text: str,
        intent: IntentDraft,
        metadata: SqliteMetadata,
    ) -> SqlOperationPlan:
        payload = {
            "user_command": user_text,
            "parsed_intent": intent.model_dump(mode="json"),
            "selected_source": {
                "name": metadata.original_name,
                "is_active": True,
                "must_be_used": True,
            },
            "sqlite_catalog": _catalog(metadata),
        }
        messages = [
            {"role": "system", "content": SQL_PLANNER_PROMPT},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False, indent=2)},
        ]
        response = await self._client.chat(
            model=self._model,
            messages=messages,
            format=SqlOperationPlan.model_json_schema(),
            options={"temperature": 0},
            think=False,
        )
        raw = response.message.content or ""
        parsed: SqlOperationPlan | None = None
        last_error: ValidationError | ValueError | None = None
        candidate = raw

        for repair_attempt in range(3):
            try:
                parsed = validate_sql_plan(
                    SqlOperationPlan.model_validate_json(candidate), intent, metadata
                )
                break
            except (ValidationError, ValueError) as error:
                last_error = error

                if _is_missing_update_assignments_error(error, intent):
                    # The broad planner understood UPDATE but emitted an invalid
                    # structured shape with assignments=[]. Recover only the
                    # requested new column/value binding using a narrow AI schema.
                    # Python does not infer UPDATE semantics from user wording.
                    base_payload = _candidate_object(candidate)
                    expected_table_name = base_payload.get("table_name")
                    recovered_assignments = await self._recover_required_assignments(
                        user_text,
                        intent,
                        metadata,
                        expected_table_name=expected_table_name,
                    )
                    base_payload = _merge_recovered_assignments_payload(
                        base_payload,
                        recovered_assignments,
                    )
                    candidate = json.dumps(base_payload, ensure_ascii=False)

                    # Re-validate the complete plan immediately. This also lets the
                    # existing dropped-filter guard run if the same broad plan lost
                    # both assignments and filters.
                    try:
                        parsed = validate_sql_plan(
                            SqlOperationPlan.model_validate_json(candidate),
                            intent,
                            metadata,
                        )
                        break
                    except (ValidationError, ValueError) as recovered_error:
                        last_error = recovered_error
                        error = recovered_error

                if _is_dropped_filter_error(error) and intent.filter_hints:
                    # Recover the already-detected row condition against the real
                    # catalog. Work from the raw structured payload because UPDATE
                    # and DELETE plans with filters=[] are intentionally invalid at
                    # the Pydantic model level.
                    base_payload = _candidate_object(candidate)
                    expected_table_name = base_payload.get("table_name")
                    recovered = await self._recover_required_filters(
                        user_text,
                        intent,
                        metadata,
                        expected_table_name=expected_table_name,
                    )
                    base_payload = _merge_recovered_filters_payload(
                        base_payload,
                        recovered,
                    )
                    candidate = json.dumps(base_payload, ensure_ascii=False)

                    # Re-enter the validation loop. A malformed UPDATE can lose
                    # both its WHERE-like filters and its assignments in the same
                    # broad AI response; recovering one structural component must
                    # not bypass recovery of the other.
                    continue
                if repair_attempt >= 2:
                    raise
                repaired = await self._client.chat(
                    model=self._model,
                    messages=[
                        *messages,
                        {"role": "assistant", "content": candidate},
                        {
                            "role": "user",
                            "content": (
                                f"Plan validation failed: {error}. Исправь JSON без "
                                "изменения смысла запроса. Не меняй action и используй "
                                "только sqlite_catalog. Верни только JSON."
                            ),
                        },
                    ],
                    format=SqlOperationPlan.model_json_schema(),
                    options={"temperature": 0},
                    think=False,
                )
                candidate = repaired.message.content or ""

        if parsed is None:
            if last_error is not None:
                raise last_error
            raise ValueError("SQL planner returned no validated plan.")

        if parsed.resolved:
            return parsed
        promoted = _promote_complete_unresolved_plan(parsed, intent, metadata)
        if promoted is not None:
            return promoted

        forced = await self._client.chat(
            model=self._model,
            messages=[
                *messages,
                {"role": "assistant", "content": parsed.model_dump_json()},
                {
                    "role": "user",
                    "content": (
                        "Повтори как best-effort. База уже выбрана, вопрос об "
                        "источнике запрещён. Если таблица одна — выбери её. "
                        "Очевидный единственный столбец выбирай по смыслу. "
                        "Пользователь увидит preview и подтвердит запись. "
                        "resolved=false допустим только без критичного значения "
                        "или при двух реально равнозначных объектах. Только JSON."
                    ),
                },
            ],
            format=SqlOperationPlan.model_json_schema(),
            options={"temperature": 0},
            think=False,
        )
        return validate_sql_plan(
            SqlOperationPlan.model_validate_json(forced.message.content or ""),
            intent,
            metadata,
        )
