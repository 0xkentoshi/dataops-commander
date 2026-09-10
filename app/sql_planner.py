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
            raise ValueError("Для is_empty/not_empty value должен быть null.")
        if self.operator not in no_value and self.value is None:
            raise ValueError(f"Для {self.operator.value} нужен value.")
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
            raise ValueError("SQL-действие не поддерживается.")
        if not self.table_name:
            raise ValueError("Нужна точная таблица.")
        if self.action in {Action.UPDATE_ROWS, Action.DELETE_ROWS} and not self.filters:
            raise ValueError("SQL UPDATE/DELETE без WHERE-подобного фильтра запрещён.")
        if self.action == Action.UPDATE_ROWS and not self.assignments:
            raise ValueError("Для UPDATE нужны новые значения.")
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
6. selected_columns=[] означает вывести все столбцы.
7. Для eq/ne/contains/not_contains/gt/gte/lt/lte/older_than_days нужен value.
8. Для is_empty/not_empty value=null.
9. Понимай опечатки, сленг, русский/английский и очевидные синонимы. Если
   таблица одна — выбирай её. Если один реальный столбец очевидно подходит по
   смыслу — выбирай его.
10. Только если два реально существующих объекта одинаково подходят или не
   указано критичное значение, resolved=false и задай один короткий вопрос.
11. resolution_note — короткое объяснение совпадения без скрытых рассуждений.
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
        raise ValueError(f"Таблица «{name}» отсутствует в каталоге.")
    return table


def _validate_column(table: SqlTableMetadata, ref: SqlColumnRef) -> None:
    actual = next((item for item in table.columns if item.name == ref.name), None)
    if actual is None:
        raise ValueError(f"Столбец «{ref.name}» не совпал с каталогом.")
    if ref.declared_type and actual.declared_type != ref.declared_type:
        raise ValueError(f"Тип столбца «{ref.name}» не совпал с каталогом.")
    ref.declared_type = actual.declared_type


def validate_sql_plan(
    plan: SqlOperationPlan,
    intent: IntentDraft,
    metadata: SqliteMetadata,
) -> SqlOperationPlan:
    if plan.action != intent.action:
        raise ValueError("SQL-планировщик изменил action.")
    if not plan.resolved:
        return plan
    table = _table(metadata, plan.table_name or "")
    for ref in [
        *plan.selected_columns,
        *(item.column for item in plan.filters),
        *(item.column for item in plan.assignments),
    ]:
        _validate_column(table, ref)
    return plan


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
        try:
            parsed = validate_sql_plan(
                SqlOperationPlan.model_validate_json(raw), intent, metadata
            )
        except (ValidationError, ValueError) as error:
            repaired = await self._client.chat(
                model=self._model,
                messages=[
                    *messages,
                    {"role": "assistant", "content": raw},
                    {
                        "role": "user",
                        "content": (
                            f"План не прошёл проверку: {error}. Исправь JSON. "
                            "Не меняй action и используй только каталог."
                        ),
                    },
                ],
                format=SqlOperationPlan.model_json_schema(),
                options={"temperature": 0},
                think=False,
            )
            parsed = validate_sql_plan(
                SqlOperationPlan.model_validate_json(repaired.message.content or ""),
                intent,
                metadata,
            )

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
