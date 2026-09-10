from __future__ import annotations

import json
from enum import Enum
from typing import TYPE_CHECKING, TypeAlias

from ollama import AsyncClient
from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator

from app.excel_service import SheetMetadata, WorkbookMetadata
from app.intent import Action, IntentDraft

if TYPE_CHECKING:
    from app.config import Settings


MAX_COLUMNS_IN_LLM_CATALOG = 120
CellValue: TypeAlias = str | int | float | bool | None

SUPPORTED_EXCEL_ACTIONS = {
    Action.SELECT,
    Action.ADD_COLUMN,
    Action.RENAME_COLUMN,
    Action.DROP_COLUMN,
    Action.UPDATE_ROWS,
    Action.REPLACE_ALL_VALUES,
    Action.DELETE_ROWS,
    Action.CLEAR_VALUES,
    Action.DEDUPLICATE,
}


class FilterOperator(str, Enum):
    EQ = "eq"
    NE = "ne"
    CONTAINS = "contains"
    NOT_CONTAINS = "not_contains"
    GT = "gt"
    GTE = "gte"
    LT = "lt"
    LTE = "lte"
    IS_EMPTY = "is_empty"
    NOT_EMPTY = "not_empty"
    OLDER_THAN_DAYS = "older_than_days"


class ColumnRef(BaseModel):
    model_config = ConfigDict(extra="forbid")

    index: int = Field(ge=1)
    header: str


class FilterCondition(BaseModel):
    model_config = ConfigDict(extra="forbid")

    column: ColumnRef
    operator: FilterOperator
    value: CellValue = None

    @model_validator(mode="after")
    def validate_filter_value(self) -> "FilterCondition":
        no_value_operators = {FilterOperator.IS_EMPTY, FilterOperator.NOT_EMPTY}
        if self.operator in no_value_operators and self.value is not None:
            raise ValueError("For is_empty/not_empty, value must be null.")
        if self.operator not in no_value_operators and self.value is None:
            raise ValueError(f"Operator {self.operator.value} requires a value.")
        return self


class ColumnAssignment(BaseModel):
    model_config = ConfigDict(extra="forbid")

    column: ColumnRef
    value: CellValue = None


class ExcelOperationPlan(BaseModel):
    model_config = ConfigDict(extra="forbid")

    action: Action
    resolved: bool
    sheet_name: str | None = None
    filters: list[FilterCondition] = Field(default_factory=list)
    selected_columns: list[ColumnRef] = Field(default_factory=list)
    target_columns: list[ColumnRef] = Field(default_factory=list)
    match_cell_value: CellValue = None
    match_all_cells: bool = False
    assignments: list[ColumnAssignment] = Field(default_factory=list)
    apply_to_all_rows: bool = False
    replacement_value: CellValue = None
    deduplicate_columns: list[ColumnRef] = Field(default_factory=list)
    new_column_name: str | None = None
    new_column_default: CellValue = None
    fill_new_column: bool = False
    confidence: float = Field(ge=0.0, le=1.0)
    resolution_note: str
    alternative_matches: list[str] = Field(default_factory=list)
    clarification_question: str | None = None

    @model_validator(mode="after")
    def validate_action_shape(self) -> "ExcelOperationPlan":
        if not self.resolved:
            return self
        if self.action not in SUPPORTED_EXCEL_ACTIONS:
            raise ValueError("This action is not supported by the Excel engine yet.")
        if self.action != Action.REPLACE_ALL_VALUES and not self.sheet_name:
            raise ValueError("A resolved plan requires sheet_name.")
        if self.action in {Action.DROP_COLUMN, Action.RENAME_COLUMN}:
            if len(self.target_columns) != 1:
                raise ValueError("Exactly one target column is required.")
        if self.action == Action.RENAME_COLUMN and not self.new_column_name:
            raise ValueError("rename_column requires a new name.")
        if self.action == Action.ADD_COLUMN and not self.new_column_name:
            raise ValueError("add_column requires a new column name.")
        if self.action == Action.UPDATE_ROWS:
            if not self.filters and not self.apply_to_all_rows:
                raise ValueError(
                    "UPDATE without filters is allowed only with explicit apply_to_all_rows=true."
                )
            if not self.assignments:
                raise ValueError("UPDATE requires assignments.")
        if self.action == Action.REPLACE_ALL_VALUES:
            if self.replacement_value is None:
                raise ValueError("replace_all_values requires replacement_value.")
            if self.filters or self.assignments or self.target_columns:
                raise ValueError(
                    "replace_all_values does not use a sheet, filters, or target columns."
                )
        if self.action == Action.DELETE_ROWS and not self.filters:
            raise ValueError("DELETE ROWS without a filter is not allowed.")
        if self.action == Action.CLEAR_VALUES:
            if not self.target_columns and self.match_cell_value is None:
                raise ValueError("CLEAR requires a target column or an exact cell value.")
            if self.match_cell_value is not None and self.target_columns:
                raise ValueError("Cell search by value cannot be combined with column clearing.")
            if self.match_cell_value is not None and self.filters:
                raise ValueError("Cell search by value does not use row filters.")
        if self.action == Action.DEDUPLICATE and not self.deduplicate_columns:
            raise ValueError("Deduplication requires key columns.")
        return self


EXCEL_PLANNER_PROMPT = """
Ты — AI-планировщик DataOps Commander для УЖЕ выбранной реальной Excel-книги.

На входе:
- исходная самостоятельная data-задача пользователя;
- parsed_intent от Command Interpreter;
- точный каталог реальных листов и столбцов.

Python НЕ будет пытаться повторно понимать русский текст regexp'ами. Именно ты
должен сопоставить смысл задачи со схемой и вернуть исполнимый ExcelOperationPlan.

Поддерживаемые действия:
- select: показать подходящие строки;
- rename_column: переименовать один существующий столбец;
- add_column: добавить новый столбец в конец таблицы;
- drop_column: полностью удалить один столбец;
- update_rows: изменить существующие ячейки в строках по фильтру;
- replace_all_values: заменить все непустые ячейки на всех листах книги;
- delete_rows: удалить строки только по обязательному фильтру;
- clear_values: очистить значения в целевых столбцах ИЛИ конкретные ячейки по
  точному содержимому, не удаляя строки;
- deduplicate: удалить повторные строки, оставив первое вхождение.

ПРАВИЛА
1. action обязан совпадать с parsed_intent.action.
2. Источник уже выбран. Никогда не спрашивай, какой файл использовать.
3. Используй только реальные sheet_name и ColumnRef из каталога, копируя index
   и header точно.
4. Не выдумывай существующие столбцы/листы/значения.
5. Если parsed_intent.column_hints содержит точное имя реального столбца,
   предпочти именно его.
6. Все filters объединяются AND.
7. update_rows без filters разрешён только при явном запросе изменить все строки:
   тогда apply_to_all_rows=true. Иначе resolved=false. Для update_rows ОБЯЗАТЕЛЬНО
   заполняй assignments: каждый изменяемый столбец должен быть ColumnAssignment с
   реальным ColumnRef и value. Не кодируй UPDATE через target_columns +
   replacement_value — эти поля не заменяют assignments.
8. delete_rows без filters всегда запрещён.
9. clear_values по конкретному содержимому: match_cell_value=X. Если пользователь
   явно просит все совпадения, match_all_cells=true; иначе false.
10. deduplicate: если пользователь назвал ключ/ключи, они обязаны попасть в
    deduplicate_columns. Полные дубли по всем столбцам используй только когда
    пользователь действительно просит полные дубли.
11. «оставить только X» обычно реализуется delete_rows с фильтром NE для строк,
    которые надо убрать, если parsed_intent.action=delete_rows.
12. Для is_empty/not_empty value=null. Для остальных filter operator нужен value.
13. add_column: fill_new_column=true только если пользователь просит сразу
    заполнить новый столбец значением.
14. Не задавай уточнение из-за сленга, регистра, склонения или очевидной опечатки,
    если по реальному каталогу есть одно однозначное совпадение.
15. resolved=false только если отсутствует критичное значение или есть минимум
    два действительно равнозначных объекта каталога.
16. resolution_note — коротко, без скрытого chain-of-thought.
17. Возвращай только JSON по схеме.
18. User-facing fields clarification_question and resolution_note must always be in English. Preserve exact sheet names, column names and data values from the source.
""".strip()


def _build_catalog(metadata: WorkbookMetadata) -> dict:
    sheets: list[dict] = []
    columns_left = MAX_COLUMNS_IN_LLM_CATALOG
    for sheet in metadata.sheets:
        if columns_left <= 0:
            break
        columns = [
            {
                "index": column.index,
                "letter": column.letter,
                "header": column.header,
                "non_empty_cells": column.non_empty_cells,
                "formula_cells": column.formula_cells,
                "samples": column.samples[:3],
            }
            for column in sheet.columns[:columns_left]
        ]
        columns_left -= len(columns)
        sheets.append(
            {
                "name": sheet.name,
                "header_row": sheet.header_row,
                "data_rows": sheet.data_rows,
                "columns": columns,
            }
        )
    return {
        "file_name": metadata.original_name,
        "sheets": sheets,
        "catalog_truncated": columns_left <= 0,
    }


def _sheet(metadata: WorkbookMetadata, name: str) -> SheetMetadata:
    result = next((sheet for sheet in metadata.sheets if sheet.name == name), None)
    if result is None:
        raise ValueError(f"Sheet “{name}” is not present in the catalog.")
    return result


def _validate_ref(sheet: SheetMetadata, ref: ColumnRef) -> None:
    actual = next((column for column in sheet.columns if column.index == ref.index), None)
    if actual is None or actual.header != ref.header:
        raise ValueError(f"Column {ref.index} / “{ref.header}” does not match the catalog.")


def _exact_hint_refs(intent: IntentDraft, sheet: SheetMetadata) -> set[int]:
    """Resolve only exact schema names; this is validation, not language parsing."""
    by_name = {column.header.strip().casefold(): column.index for column in sheet.columns}
    return {
        by_name[hint.strip().casefold()]
        for hint in intent.column_hints
        if hint.strip().casefold() in by_name
    }


def _validate_against_catalog(
    plan: ExcelOperationPlan,
    intent: IntentDraft,
    metadata: WorkbookMetadata,
) -> ExcelOperationPlan:
    if plan.action != intent.action:
        raise ValueError("The planner changed the action from the parsed intent.")
    if not plan.resolved:
        return plan
    if plan.action == Action.REPLACE_ALL_VALUES:
        return plan
    if plan.sheet_name is None:
        raise ValueError("No sheet was specified.")
    sheet = _sheet(metadata, plan.sheet_name)
    refs = [
        *plan.selected_columns,
        *plan.target_columns,
        *plan.deduplicate_columns,
        *(condition.column for condition in plan.filters),
        *(assignment.column for assignment in plan.assignments),
    ]
    for ref in refs:
        _validate_ref(sheet, ref)

    # Safety invariant: when the interpreter already resolved exact schema keys
    # for deduplication, the planner may not silently replace them with others.
    if plan.action == Action.DEDUPLICATE:
        expected = _exact_hint_refs(intent, sheet)
        if expected:
            actual = {ref.index for ref in plan.deduplicate_columns}
            if actual != expected:
                raise ValueError(
                    "The deduplication key from the intent does not match the Excel plan."
                )

    # Safety invariant based on STRUCTURED intent, not re-parsing user language.
    if (
        intent.filter_hints
        and plan.action in {Action.SELECT, Action.UPDATE_ROWS, Action.CLEAR_VALUES}
        and plan.match_cell_value is None
        and not plan.filters
        and not plan.apply_to_all_rows
    ):
        raise ValueError("The intent contains a filter, but the Excel plan lost it.")

    if plan.new_column_name:
        normalized = plan.new_column_name.strip().casefold()
        if any(column.header.strip().casefold() == normalized for column in sheet.columns):
            raise ValueError(f"Column “{plan.new_column_name}” already exists.")
    return plan



def _normalize_structural_plan_payload(payload: dict) -> dict:
    """Repair safe field-shape aliases without re-interpreting user language.

    Some local models occasionally represent a single-column UPDATE as
    target_columns + replacement_value even though the executor schema requires
    assignments. When that mapping is unambiguous, normalize only the structure.
    No column/value is inferred from raw user text here.
    """
    normalized = dict(payload)
    action = normalized.get("action")
    assignments = normalized.get("assignments") or []
    target_columns = normalized.get("target_columns") or []
    replacement_value = normalized.get("replacement_value")

    if (
        action == Action.UPDATE_ROWS.value
        and not assignments
        and len(target_columns) == 1
        and replacement_value is not None
    ):
        normalized["assignments"] = [
            {
                "column": target_columns[0],
                "value": replacement_value,
            }
        ]
        normalized["target_columns"] = []
        normalized["replacement_value"] = None

    return normalized


def _parse_and_validate_plan(
    raw: str,
    intent: IntentDraft,
    metadata: WorkbookMetadata,
) -> ExcelOperationPlan:
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as error:
        raise ValueError(f"Planner returned invalid JSON: {error}") from error
    if not isinstance(payload, dict):
        raise ValueError("Planner JSON must be an object.")
    payload = _normalize_structural_plan_payload(payload)
    parsed = ExcelOperationPlan.model_validate(payload)
    return _validate_against_catalog(parsed, intent, metadata)


def _action_shape_contract(action: Action) -> str:
    if action == Action.UPDATE_ROWS:
        return (
            "For update_rows: assignments MUST contain at least one object with "
            "the real target ColumnRef and the value to write. Row-selection "
            "conditions belong in filters. Do not encode an update only as "
            "target_columns/replacement_value."
        )
    if action == Action.RENAME_COLUMN:
        return (
            "For rename_column: target_columns must contain exactly one real "
            "ColumnRef and new_column_name must contain the requested new name."
        )
    if action == Action.DROP_COLUMN:
        return "For drop_column: target_columns must contain exactly one real ColumnRef."
    if action == Action.DELETE_ROWS:
        return "For delete_rows: filters must be non-empty."
    if action == Action.DEDUPLICATE:
        return "For deduplicate: deduplicate_columns must contain the real key columns."
    if action == Action.CLEAR_VALUES:
        return (
            "For clear_values: use target_columns for column clearing OR "
            "match_cell_value for exact-cell-content clearing."
        )
    return "Follow the exact conditional field requirements of ExcelOperationPlan."

def _promote_complete_unresolved_plan(
    plan: ExcelOperationPlan,
    intent: IntentDraft,
    metadata: WorkbookMetadata,
) -> ExcelOperationPlan | None:
    if plan.resolved:
        return plan
    payload = plan.model_dump(mode="python")
    payload["resolved"] = True
    if not payload.get("sheet_name") and len(metadata.sheets) == 1:
        payload["sheet_name"] = metadata.sheets[0].name
    try:
        promoted = ExcelOperationPlan.model_validate(payload)
        return _validate_against_catalog(promoted, intent, metadata)
    except (ValidationError, ValueError):
        return None


# Kept as a compatibility hook for older imports. It intentionally performs no
# NLP: v5.3 routes natural language through the LLM planner only.
def try_fast_excel_plan(
    user_text: str,
    intent: IntentDraft,
    metadata: WorkbookMetadata,
) -> ExcelOperationPlan | None:
    del user_text, intent, metadata
    return None


class ExcelPlanner:
    def __init__(self, settings: Settings) -> None:
        self._model = settings.ollama_model
        self._client = AsyncClient(host=settings.ollama_host, timeout=120.0)

    async def resolve(
        self,
        user_text: str,
        intent: IntentDraft,
        metadata: WorkbookMetadata,
    ) -> ExcelOperationPlan:
        payload = {
            "user_command": user_text,
            "parsed_intent": intent.model_dump(mode="json"),
            "selected_source": {
                "name": metadata.original_name,
                "is_active": True,
                "must_be_used": True,
            },
            "excel_catalog": _build_catalog(metadata),
        }
        messages = [
            {"role": "system", "content": EXCEL_PLANNER_PROMPT},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False, indent=2)},
        ]
        response = await self._client.chat(
            model=self._model,
            messages=messages,
            format=ExcelOperationPlan.model_json_schema(),
            options={"temperature": 0},
            think=False,
        )
        raw = response.message.content or ""
        last_error: ValidationError | ValueError | None = None
        parsed: ExcelOperationPlan | None = None
        candidate_raw = raw

        # Local structured-output models can occasionally satisfy the JSON schema
        # while missing an action-specific invariant implemented by Pydantic
        # validators (for example UPDATE with no assignments). Validate every
        # candidate, repair safe structural aliases, then allow two bounded LLM
        # corrections before surfacing a planning error.
        for repair_attempt in range(3):
            try:
                parsed = _parse_and_validate_plan(candidate_raw, intent, metadata)
                break
            except (ValidationError, ValueError) as error:
                last_error = error
                if repair_attempt >= 2:
                    break
                contract = _action_shape_contract(intent.action)
                repair = await self._client.chat(
                    model=self._model,
                    messages=[
                        *messages,
                        {"role": "assistant", "content": candidate_raw},
                        {
                            "role": "user",
                            "content": (
                                f"Plan validation failed: {error}. "
                                "Перестрой только ЭТУ data-задачу. Не меняй action и "
                                "не теряй значения/условия из parsed_intent. Используй "
                                "только реальные объекты excel_catalog. "
                                f"Обязательный контракт действия: {contract} "
                                "Верни полный ExcelOperationPlan JSON без текста вокруг."
                            ),
                        },
                    ],
                    format=ExcelOperationPlan.model_json_schema(),
                    options={"temperature": 0},
                    think=False,
                )
                candidate_raw = repair.message.content or ""

        if parsed is None:
            raise ValueError(
                "LLM three times returned an invalid Excel plan for "
                f"{intent.action.value}: {last_error}"
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
                        "Retry best-effort. The file is already selected. If there is one sheet, "
                        "select it. If exactly one real column matches the schema, "
                        "select it. The user will still see an exact preview before "
                        "writing. Use resolved=false only for genuine "
                        "ambiguity or a missing critical value. JSON only."
                    ),
                },
            ],
            format=ExcelOperationPlan.model_json_schema(),
            options={"temperature": 0},
            think=False,
        )
        return _parse_and_validate_plan(
            forced.message.content or "", intent, metadata
        )
