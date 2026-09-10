from __future__ import annotations

import json
from enum import Enum
from typing import TYPE_CHECKING, Any

from ollama import AsyncClient
from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator

if TYPE_CHECKING:
    from app.config import Settings


class SourceType(str, Enum):
    SQL = "sql"
    EXCEL = "excel"
    CSV = "csv"
    GOOGLE_SHEETS = "google_sheets"
    UNKNOWN = "unknown"


class Action(str, Enum):
    OPEN_SOURCE = "open_source"
    SELECT = "select"
    INSERT_ROWS = "insert_rows"
    UPDATE_ROWS = "update_rows"
    REPLACE_ALL_VALUES = "replace_all_values"
    DELETE_ROWS = "delete_rows"
    ADD_COLUMN = "add_column"
    RENAME_COLUMN = "rename_column"
    DROP_COLUMN = "drop_column"
    CLEAR_VALUES = "clear_values"
    COPY_ROWS = "copy_rows"
    CREATE_SHEET = "create_sheet"
    DELETE_SHEET = "delete_sheet"
    DEDUPLICATE = "deduplicate"
    FIX_FORMAT = "fix_format"
    RESTORE_FORMULAS = "restore_formulas"
    COMPARE = "compare"
    REPORT = "report"
    UNKNOWN = "unknown"


class CommandMode(str, Enum):
    """High-level user intent before a data plan is built."""

    DATA = "data"
    UNDO = "undo"
    REPEAT_LAST = "repeat_last"
    SEND_LAST_FILE = "send_last_file"
    RENAME_LAST_FILE = "rename_last_file"
    UPDATE_PENDING = "update_pending"
    CONFIRM_PENDING = "confirm_pending"
    CANCEL_PENDING = "cancel_pending"
    UNKNOWN = "unknown"


class ExecutionDirective(BaseModel):
    model_config = ConfigDict(extra="forbid")

    copy_original: bool = False
    output_name: str | None = Field(
        default=None,
        description="Requested result filename. Null when user did not request one.",
    )


class IntentDraft(BaseModel):
    model_config = ConfigDict(extra="forbid")

    normalized_request: str = Field(
        description="Нормальная и однозначная формулировка одной data-задачи"
    )
    action: Action
    source_type_hint: SourceType
    source_name_hint: str | None = Field(
        description="Название файла/подключения, только если пользователь назвал его"
    )
    resource_name_hint: str | None = Field(
        description="Название SQL-таблицы или Excel-листа, если оно указано"
    )
    column_hints: list[str] = Field(
        description=(
            "Столбцы задачи. Если schema активного источника передана и совпадение "
            "однозначно, используй точные реальные имена столбцов из schema."
        )
    )
    filter_hints: list[str] = Field(
        description="Условия отбора строк обычным, но однозначным языком"
    )
    value_hints: list[str] = Field(
        description="Значения, которые требуется записать, найти или сравнить"
    )
    is_write_operation: bool
    is_destructive: bool
    needs_discovery: bool = Field(
        description="Нужно ли ещё исследовать источник/схему до построения плана"
    )
    confidence: float = Field(ge=0.0, le=1.0)
    interpretation_note: str = Field(
        description="Короткое объяснение интерпретации без скрытых рассуждений"
    )
    clarification_question: str | None = Field(
        description="Вопрос только при реальной неоднозначности/нехватке данных"
    )


class IntentBatch(BaseModel):
    """All independent data tasks from one message, preserving original order."""

    model_config = ConfigDict(extra="forbid")

    tasks: list[IntentDraft] = Field(min_length=1, max_length=12)

    @model_validator(mode="after")
    def propagate_shared_context(self) -> "IntentBatch":
        source_names = {task.source_name_hint for task in self.tasks if task.source_name_hint}
        resources = {task.resource_name_hint for task in self.tasks if task.resource_name_hint}
        source_types = {
            task.source_type_hint
            for task in self.tasks
            if task.source_type_hint != SourceType.UNKNOWN
        }
        shared_source = next(iter(source_names)) if len(source_names) == 1 else None
        shared_resource = next(iter(resources)) if len(resources) == 1 else None
        shared_type = next(iter(source_types)) if len(source_types) == 1 else SourceType.UNKNOWN
        self.tasks = [
            task.model_copy(
                update={
                    "source_name_hint": task.source_name_hint or shared_source,
                    "source_type_hint": (
                        task.source_type_hint
                        if task.source_type_hint != SourceType.UNKNOWN
                        else shared_type
                    ),
                    "resource_name_hint": task.resource_name_hint or shared_resource,
                }
            )
            for task in self.tasks
        ]
        return self


class CommandEnvelope(BaseModel):
    """Single AI-native interpretation of the whole Telegram message."""

    model_config = ConfigDict(extra="forbid")

    mode: CommandMode
    tasks: list[IntentDraft] = Field(default_factory=list, max_length=12)
    execution: ExecutionDirective = Field(default_factory=ExecutionDirective)
    target_source_name: str | None = Field(
        default=None,
        description=(
            "Target filename/source for repeat/open semantics when explicitly or "
            "unambiguously referenced by the user."
        ),
    )
    rename_to: str | None = Field(
        default=None,
        description="New filename for rename_last_file mode.",
    )
    confidence: float = Field(ge=0.0, le=1.0)
    interpretation_note: str
    clarification_question: str | None = None

    @model_validator(mode="after")
    def validate_mode_shape(self) -> "CommandEnvelope":
        if self.mode == CommandMode.DATA and not self.tasks:
            raise ValueError("data mode requires at least one task")
        if self.mode != CommandMode.DATA and self.tasks:
            # Non-data commands must never accidentally execute data tasks.
            self.tasks = []
        if self.mode == CommandMode.RENAME_LAST_FILE and not (
            self.rename_to or self.execution.output_name
        ):
            raise ValueError("rename_last_file requires a target filename")
        return self


COMMAND_SYSTEM_PROMPT = """
Ты — AI Command Interpreter для DataOps Commander.

Твоя задача — понять ВСЁ сообщение пользователя по смыслу и вернуть один
CommandEnvelope. Python после тебя НЕ будет распознавать русские фразы regexp'ами:
поэтому именно ты отвечаешь за семантику, разбиение составной команды и микро-команды.
Ты ничего не выполняешь сам.

Тебе передаётся runtime_context. В нём могут быть:
- active_source: уже выбранный файл/БД с реальными листами, таблицами и столбцами;
- pending_operation: есть ли операция, ожидающая подтверждения;
- last_completed_operation: последняя успешная операция и её исходная команда;
- available_sources: известные файлы пользователя.

COMMAND MODES
- data: пользователь просит прочитать/изменить данные;
- undo: вернуть/откатить последнее выполненное изменение;
- repeat_last: повторить прошлую успешную data-операцию на текущем или другом файле;
- send_last_file: прислать/скинуть/отправить последний обработанный файл;
- rename_last_file: переименовать уже готовый последний файл;
- update_pending: изменить параметры ожидающего preview, например работать на копии
  или задать имя результата, НЕ меняя сам data-план;
- confirm_pending: обычным текстом подтвердить ожидающую операцию;
- cancel_pending: обычным текстом отменить ожидающую операцию;
- unknown: запрос нельзя безопасно классифицировать.

ГЛАВНЫЕ ПРАВИЛА
1. Понимай смысл, а не конкретные фразы. Сленг, опечатки, сокращения и свободный
   порядок слов нормальны.
2. Если active_source передан, считай его основным источником. Не спрашивай имя
   файла повторно. Используй реальные schema names из runtime_context.
3. Одно сообщение может содержать много data-задач. Для mode=data верни КАЖДОЕ
   самостоятельное действие отдельным IntentDraft в исходном порядке. Ничего не
   схлопывай и не теряй.
4. Например смысл «убери записи без email и дубли по email» — ДВЕ задачи:
   delete_rows с email is empty, затем deduplicate по email.
5. «оставь только ...» означает удалить строки, которые НЕ удовлетворяют
   указанному условию. Сформулируй filter_hints однозначно, например
   «city != Новосибирск», чтобы executor удалил лишние строки.
6. «удали/очисти ячейку со значением X» — clear_values, НЕ delete_rows.
7. «удали строки/записи ...» — delete_rows.
8. «дубли по X» — deduplicate, а column_hints должен содержать X. Если schema
   активного источника известна, используй точное имя реального столбца.
9. «переименуй столбец A в B» — rename_column: column_hints=[A], value_hints=[B].
10. «сохрани отдельно», «не трогай оригинал», «сделай копию» и любые
    семантически эквивалентные просьбы -> execution.copy_original=true.
11. Просьба назвать результат/итоговый файл -> execution.output_name.
    Если это единственная просьба и pending_operation=true -> update_pending.
    Если pending_operation=false и речь о уже готовом файле -> rename_last_file.
12. Если пользователь просит повторить прошлое действие, mode=repeat_last.
    Если назван новый файл — заполни target_source_name. Сам прошлый data-план
    не выдумывай: runtime_context содержит last_completed_operation, а приложение
    заново интерпретирует его на схеме нового файла.
13. Если пользователь просит последний обработанный файл -> send_last_file.
14. Если пользователь передумал после выполнения («верни как было», «откати») -> undo.
15. Если pending_operation=true и пользователь просто соглашается/подтверждает ->
    confirm_pending; если отказывается/отменяет -> cancel_pending.
16. Не превращай произвольное «да» в confirm_pending, если pending_operation=false.
17. Не придумывай существующие листы/таблицы/столбцы. При известной schema выбирай
    точные имена из неё. Если пользователь назвал понятный существующий столбец,
    не задавай бессмысленное уточнение.
18. clarification_question нужен только при реальном противоречии или когда без
    уточнения нельзя безопасно определить действие/значение.
19. source_name_hint не заполняй старым именем файла при repeat_last, если
    пользователь повторяет операцию на новом источнике.
20. Возвращай только JSON по схеме. Никакого текста вокруг JSON.

ACTION SEMANTICS
- open_source: выбрать/открыть источник;
- select: показать/найти строки;
- update_rows: изменить значения в строках;
- delete_rows: удалить строки по условию;
- clear_values: очистить ячейки/значения без удаления строк;
- deduplicate: удалить дубли;
- rename_column: переименовать столбец;
- add_column: добавить столбец;
- drop_column: удалить столбец;
- replace_all_values: заменить все заполненные ячейки во всей книге;
- прочие Action используй только если они действительно соответствуют запросу.

Для каждой data-задачи normalized_request делай самостоятельным и однозначным:
он должен содержать все нужные условия этой конкретной задачи, даже если в
исходной фразе контекст был указан один раз на весь список.
""".strip()


class IntentParser:
    def __init__(self, settings: Settings) -> None:
        self._model = settings.ollama_model
        self._client = AsyncClient(host=settings.ollama_host, timeout=120.0)

    async def ping(self) -> None:
        await self._client.show(self._model)

    async def parse_command(
        self,
        user_text: str,
        runtime_context: dict[str, Any] | None = None,
    ) -> CommandEnvelope:
        payload = {
            "user_message": user_text,
            "runtime_context": runtime_context or {},
        }
        messages = [
            {"role": "system", "content": COMMAND_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": json.dumps(payload, ensure_ascii=False, indent=2),
            },
        ]
        response = await self._client.chat(
            model=self._model,
            messages=messages,
            format=CommandEnvelope.model_json_schema(),
            options={"temperature": 0},
            think=False,
        )
        raw_content = response.message.content or ""
        try:
            envelope = CommandEnvelope.model_validate_json(raw_content)
        except ValidationError as error:
            repair = await self._client.chat(
                model=self._model,
                messages=[
                    *messages,
                    {"role": "assistant", "content": raw_content},
                    {
                        "role": "user",
                        "content": (
                            "Ответ не прошёл JSON Schema: "
                            f"{error}. Исправь только структуру/пропущенные поля. "
                            "Не теряй ни одной задачи пользователя и не меняй их смысл. "
                            "Верни только CommandEnvelope JSON."
                        ),
                    },
                ],
                format=CommandEnvelope.model_json_schema(),
                options={"temperature": 0},
                think=False,
            )
            envelope = CommandEnvelope.model_validate_json(repair.message.content or "")

        if envelope.mode == CommandMode.DATA:
            envelope = envelope.model_copy(
                update={"tasks": IntentBatch(tasks=envelope.tasks).tasks}
            )
        return envelope

    async def parse_many(
        self,
        user_text: str,
        runtime_context: dict[str, Any] | None = None,
    ) -> list[IntentDraft]:
        """Backward-compatible API used by older tests/integrations."""
        envelope = await self.parse_command(user_text, runtime_context)
        if envelope.mode != CommandMode.DATA:
            return []
        return envelope.tasks

    async def parse(
        self,
        user_text: str,
        runtime_context: dict[str, Any] | None = None,
    ) -> IntentDraft:
        tasks = await self.parse_many(user_text, runtime_context)
        if not tasks:
            raise ValueError("Команда не содержит data-задачи.")
        return tasks[0]
