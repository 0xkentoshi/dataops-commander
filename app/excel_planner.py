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
            raise ValueError("Для is_empty/not_empty value должен быть null.")
        if self.operator not in no_value_operators and self.value is None:
            raise ValueError(f"Для оператора {self.operator.value} нужен value.")
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
            raise ValueError("Действие пока не поддерживается Excel engine.")
        if self.action != Action.REPLACE_ALL_VALUES and not self.sheet_name:
            raise ValueError("Для resolved-плана обязателен sheet_name.")
        if self.action in {Action.DROP_COLUMN, Action.RENAME_COLUMN}:
            if len(self.target_columns) != 1:
                raise ValueError("Нужен ровно один целевой столбец.")
        if self.action == Action.RENAME_COLUMN and not self.new_column_name:
            raise ValueError("Для rename_column нужно новое имя.")
        if self.action == Action.ADD_COLUMN and not self.new_column_name:
            raise ValueError("Для add_column нужно имя нового столбца.")
        if self.action == Action.UPDATE_ROWS:
            if not self.filters and not self.apply_to_all_rows:
                raise ValueError(
                    "UPDATE без фильтра разрешён только при явном apply_to_all_rows=true."
                )
            if not self.assignments:
                raise ValueError("Для UPDATE нужны присваивания.")
        if self.action == Action.REPLACE_ALL_VALUES:
            if self.replacement_value is None:
                raise ValueError("Для замены всех ячеек нужно replacement_value.")
            if self.filters or self.assignments or self.target_columns:
                raise ValueError(
                    "replace_all_values не использует лист, фильтры или столбцы."
                )
        if self.action == Action.DELETE_ROWS and not self.filters:
            raise ValueError("DELETE ROWS без фильтра запрещён.")
        if self.action == Action.CLEAR_VALUES:
            if not self.target_columns and self.match_cell_value is None:
                raise ValueError("Для очистки нужен столбец или точное значение ячейки.")
            if self.match_cell_value is not None and self.target_columns:
                raise ValueError("Поиск ячеек по значению не смешивается с очисткой столбцов.")
            if self.match_cell_value is not None and self.filters:
                raise ValueError("Поиск ячеек по значению не использует построчные filters.")
        if self.action == Action.DEDUPLICATE and not self.deduplicate_columns:
            raise ValueError("Для удаления дублей нужны ключевые столбцы.")
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
   тогда apply_to_all_rows=true. Иначе resolved=false.
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
        raise ValueError(f"Лист «{name}» отсутствует в каталоге.")
    return result


def _validate_ref(sheet: SheetMetadata, ref: ColumnRef) -> None:
    actual = next((column for column in sheet.columns if column.index == ref.index), None)
    if actual is None or actual.header != ref.header:
        raise ValueError(f"Столбец {ref.index} / «{ref.header}» не совпал с каталогом.")


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
        raise ValueError("Планировщик изменил action из разобранного намерения.")
    if not plan.resolved:
        return plan
    if plan.action == Action.REPLACE_ALL_VALUES:
        return plan
    if plan.sheet_name is None:
        raise ValueError("Не указан лист.")
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
                    "Ключ дедупликации из intent не совпал с ключом Excel-плана."
                )

    # Safety invariant based on STRUCTURED intent, not re-parsing user language.
    if (
        intent.filter_hints
        and plan.action in {Action.SELECT, Action.UPDATE_ROWS, Action.CLEAR_VALUES}
        and plan.match_cell_value is None
        and not plan.filters
        and not plan.apply_to_all_rows
    ):
        raise ValueError("Intent содержит фильтр, но Excel-план потерял его.")

    if plan.new_column_name:
        normalized = plan.new_column_name.strip().casefold()
        if any(column.header.strip().casefold() == normalized for column in sheet.columns):
            raise ValueError(f"Столбец «{plan.new_column_name}» уже существует.")
    return plan


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
        try:
            parsed = ExcelOperationPlan.model_validate_json(raw)
            parsed = _validate_against_catalog(parsed, intent, metadata)
        except (ValidationError, ValueError) as error:
            repair = await self._client.chat(
                model=self._model,
                messages=[
                    *messages,
                    {"role": "assistant", "content": raw},
                    {
                        "role": "user",
                        "content": (
                            "План не прошёл проверку: "
                            f"{error}. Исправь только план. Не меняй action. "
                            "Используй лишь реальные объекты excel_catalog и сохрани "
                            "все фильтры/ключи из parsed_intent. Верни только JSON."
                        ),
                    },
                ],
                format=ExcelOperationPlan.model_json_schema(),
                options={"temperature": 0},
                think=False,
            )
            parsed = ExcelOperationPlan.model_validate_json(repair.message.content or "")
            parsed = _validate_against_catalog(parsed, intent, metadata)

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
                        "Сделай best-effort повторно. Файл уже выбран. Если лист один — "
                        "выбери его. Если по schema подходит один реальный столбец — "
                        "выбери его. Пользователь всё равно увидит точный preview перед "
                        "записью. resolved=false оставляй только при реальной "
                        "неоднозначности/отсутствующем критичном значении. Только JSON."
                    ),
                },
            ],
            format=ExcelOperationPlan.model_json_schema(),
            options={"temperature": 0},
            think=False,
        )
        final = ExcelOperationPlan.model_validate_json(forced.message.content or "")
        return _validate_against_catalog(final, intent, metadata)
