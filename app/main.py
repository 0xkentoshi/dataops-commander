from __future__ import annotations

import asyncio
import html
import shutil
import tempfile
from dataclasses import dataclass, field, replace
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import TypeAlias
from uuid import uuid4

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.types import CallbackQuery, FSInputFile, InlineKeyboardButton, InlineKeyboardMarkup, Message

from app.batch_operations import (
    BatchExecutionOutcome,
    copy_sqlite_database,
    execute_excel_batch,
    execute_sqlite_batch,
)
from app.config import get_settings
from app.database import AuditEventRecord, DataRepository, OperationRecord
from app.excel_operations import (
    ExcelExecutionOutcome,
    ExcelOperationPreview,
    WRITE_ACTIONS,
    build_operation_preview,
    execute_operation,
)
from app.excel_planner import (
    ExcelOperationPlan,
    ExcelPlanner,
    SUPPORTED_EXCEL_ACTIONS,
)
from app.excel_service import WorkbookMetadata, inspect_workbook
from app.intent import Action, CommandMode, IntentDraft, IntentParser
from app.source_catalog import (
    DataSource,
    SourceCatalog,
    SourceKind,
    SourceOrigin,
    match_sources,
)
from app.sql_operations import (
    SQL_WRITE_ACTIONS,
    SqlExecutionOutcome,
    SqlOperationPreview,
    build_sql_preview,
    execute_sql_operation,
)
from app.sql_planner import (
    SUPPORTED_SQL_ACTIONS,
    SqlOperationPlan,
    SqlPlanner,
)
from app.sql_service import SqliteMetadata, inspect_sqlite
from app.rollback import restore_snapshot, sha256_file
from app.micro_features import (
    ExecutionOptions,
    normalized_output_filename,
)


DATABASE_PATH = Path("data/dataops.sqlite3")
TELEGRAM_SOURCES_DIRECTORY = Path("data/telegram_sources")
MAX_TELEGRAM_FILE_SIZE_BYTES = 20 * 1024 * 1024
FILES_PER_PAGE = 7
OPERATION_TTL = timedelta(minutes=15)
CURRENT_ONBOARDING_VERSION = 1
SOURCE_INSPECTION_TIMEOUT_SECONDS = 30.0

settings = get_settings()
intent_parser = IntentParser(settings)
excel_planner = ExcelPlanner(settings)
sql_planner = SqlPlanner(settings)
repository = DataRepository(DATABASE_PATH)
catalog = SourceCatalog(
    settings.workspace_directory,
    TELEGRAM_SOURCES_DIRECTORY,
    excluded_paths={DATABASE_PATH.resolve()},
    legacy_telegram_root=Path("data/uploads"),
)
router = Router()
file_locks: dict[str, asyncio.Lock] = {}


SourceMetadata: TypeAlias = WorkbookMetadata | SqliteMetadata


@dataclass(frozen=True, slots=True)
class ActiveSource:
    source: DataSource
    metadata: SourceMetadata


@dataclass(frozen=True, slots=True)
class PendingExcelOperation:
    operation_id: str
    user_id: int
    source: DataSource
    metadata: WorkbookMetadata
    plan: ExcelOperationPlan
    preview: ExcelOperationPreview
    created_at: datetime
    execution: ExecutionOptions = field(default_factory=ExecutionOptions)


@dataclass(frozen=True, slots=True)
class PendingSqlOperation:
    operation_id: str
    user_id: int
    source: DataSource
    metadata: SqliteMetadata
    plan: SqlOperationPlan
    preview: SqlOperationPreview
    created_at: datetime
    execution: ExecutionOptions = field(default_factory=ExecutionOptions)


SinglePendingOperation: TypeAlias = PendingExcelOperation | PendingSqlOperation


@dataclass(frozen=True, slots=True)
class PendingBatchOperation:
    operation_id: str
    user_id: int
    source: DataSource
    steps: list[SinglePendingOperation]
    created_at: datetime
    execution: ExecutionOptions = field(default_factory=ExecutionOptions)


PendingOperation: TypeAlias = SinglePendingOperation | PendingBatchOperation


ACTION_LABELS = {
    Action.OPEN_SOURCE: "select source",
    Action.SELECT: "show matching rows",
    Action.ADD_COLUMN: "add column",
    Action.RENAME_COLUMN: "rename column",
    Action.DROP_COLUMN: "drop column",
    Action.UPDATE_ROWS: "update rows",
    Action.REPLACE_ALL_VALUES: "replace all filled cells",
    Action.DELETE_ROWS: "delete rows",
    Action.CLEAR_VALUES: "clear values",
    Action.DEDUPLICATE: "remove duplicates",
}

STATUS_LABELS = {
    "received": "🧠 Analyzing",
    "planned": "🟡 Awaiting confirmation",
    "previewed": "👁 Previewed",
    "executing": "🔵 Executing",
    "completed": "✅ Completed",
    "cancelled": "❌ Cancelled",
    "expired": "⌛ Expired",
    "needs_clarification": "❓ Needs clarification",
    "rejected": "⛔ Rejected",
    "failed": "⚠️ Failed",
    "undone": "↩️ Undone",
}

EVENT_LABELS = {
    "request_received": "Request received",
    "intent_parsed": "Intent parsed",
    "plan_created": "Structured plan created",
    "preview_shown": "Preview shown",
    "confirmation_received": "Confirmation received",
    "snapshot_created": "Snapshot created",
    "change_applied": "Change applied",
    "verification_passed": "Result verified",
    "source_updated": "Source updated in place",
    "result_sent": "Result sent to Telegram (legacy event)",
    "active_source_selected": "Source selected as active",
    "active_source_cleared": "Active source cleared",
    "operation_completed": "Operation completed",
    "read_completed": "Read completed",
    "no_changes_found": "No matching changes found",
    "operation_cancelled": "Operation cancelled",
    "operation_replaced": "Replaced by a new request",
    "confirmation_expired": "Confirmation window expired",
    "operation_rejected": "Operation blocked",
    "clarification_required": "Clarification required",
    "planning_failed": "Planning failed",
    "execution_failed": "Execution failed",
    "execution_interrupted": "Execution interrupted by restart",
    "operation_undone": "Change reverted from snapshot",
    "rollback_completed": "Rollback completed",
}


def esc(value: object) -> str:
    return html.escape(str(value))


def planning_failure_text() -> str:
    return (
        "<b>Could not prepare a safe plan</b>\n"
        "The AI response did not pass validation. No data was changed. "
        "Please try the request again."
    )


def allowed(user_id: int) -> bool:
    return user_id in settings.allowed_user_ids


def shorten(value: str, limit: int = 80) -> str:
    compact = " ".join(value.split())
    return compact if len(compact) <= limit else f"{compact[:limit - 1]}…"


def format_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).strftime("%d.%m.%Y %H:%M:%S UTC")


def human_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def _lock_for(path: Path) -> asyncio.Lock:
    key = str(path.resolve()).casefold()
    lock = file_locks.get(key)
    if lock is None:
        lock = asyncio.Lock()
        file_locks[key] = lock
    return lock


def main_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📂 Open files", callback_data="files:page:0")],
            [
                InlineKeyboardButton(text="🧾 History", callback_data="menu:history"),
                InlineKeyboardButton(text="🩺 Health", callback_data="menu:health"),
            ],
        ]
    )


def onboarding_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="✅ Got it, start",
                    callback_data="onboarding:accept",
                )
            ]
        ]
    )


def onboarding_text() -> str:
    return (
        "<b>DataOps Commander</b>\n"
        "Excel + SQLite · manage data with natural language.\n\n"
        "<b>How it works</b>\n"
        "• Send one or multiple tasks in a single message.\n"
        "• Read operations run immediately; write operations are shown as a plan first.\n"
        "• Writes run only after your explicit confirmation.\n"
        "• A source snapshot is created before a write; the latest change can be reverted with Undo or a natural-language request.\n\n"
        "<b>Important</b>\n"
        "• Review the plan before confirming: natural language can be "
        "interpreted differently than expected.\n"
        "• Ambiguous or unsupported actions will require clarification "
        "or will be rejected.\n"
        "• One batch operates on one source; tasks inside the batch "
        "run sequentially.\n\n"
        "Technical history is available through /history and /audit. "
        "These onboarding notes will not be repeated in the working interface."
    )


def onboarding_accepted(user_id: int) -> bool:
    return repository.get_onboarding_version(user_id) >= CURRENT_ONBOARDING_VERSION


async def require_onboarding_message(message: Message, user_id: int) -> bool:
    if onboarding_accepted(user_id):
        return True
    await message.answer(
        "Send /start first and accept the startup rules."
    )
    return False


async def require_onboarding_callback(callback: CallbackQuery) -> bool:
    if onboarding_accepted(callback.from_user.id):
        return True
    await callback.answer(
        "Send /start first and accept the startup rules.",
        show_alert=True,
    )
    return False


def operation_keyboard(operation_id: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="👁 Show data",
                    callback_data=f"op:preview:{operation_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    text="✅ Confirm",
                    callback_data=f"op:confirm:{operation_id}",
                ),
                InlineKeyboardButton(
                    text="❌ Cancel",
                    callback_data=f"op:cancel:{operation_id}",
                ),
            ],
        ]
    )


def files_keyboard(
    sources: list[DataSource],
    page: int,
    active_path: Path | None,
) -> InlineKeyboardMarkup:
    pages = max((len(sources) - 1) // FILES_PER_PAGE + 1, 1)
    safe_page = min(max(page, 0), pages - 1)
    start = safe_page * FILES_PER_PAGE
    rows: list[list[InlineKeyboardButton]] = []
    for source in sources[start : start + FILES_PER_PAGE]:
        active = "✅ " if active_path and source.path == active_path.resolve() else ""
        shown_name = source.relative_name if source.origin == SourceOrigin.LOCAL else source.display_name
        rows.append(
            [
                InlineKeyboardButton(
                    text=(
                        f"{active}{source.origin_icon} {shorten(shown_name, 34)} "
                        f"· {source.extension.removeprefix('.').upper() or 'FILE'}"
                    ),
                    callback_data=f"files:open:{source.source_id}:{safe_page}",
                )
            ]
        )
    navigation: list[InlineKeyboardButton] = []
    if safe_page > 0:
        navigation.append(
            InlineKeyboardButton(text="⬅️", callback_data=f"files:page:{safe_page - 1}")
        )
    navigation.append(
        InlineKeyboardButton(text=f"{safe_page + 1}/{pages}", callback_data="files:noop")
    )
    if safe_page + 1 < pages:
        navigation.append(
            InlineKeyboardButton(text="➡️", callback_data=f"files:page:{safe_page + 1}")
        )
    rows.append(navigation)
    rows.append(
        [
            InlineKeyboardButton(
                text="🔄 Refresh list",
                callback_data=f"files:refresh:{safe_page}",
            )
        ]
    )
    return InlineKeyboardMarkup(inline_keyboard=rows)


def back_to_files_keyboard(page: int = 0) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="⬅️ Back to list", callback_data=f"files:page:{page}")],
            [InlineKeyboardButton(text="🔄 Refresh", callback_data=f"files:refresh:{page}")],
        ]
    )


def completed_operation_keyboard(operation_id: str, page: int = 0) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📤 Download file",
                    callback_data=f"sendfile:{operation_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    text="↩️ Undo changes",
                    callback_data=f"undo:{operation_id}",
                )
            ],
            [InlineKeyboardButton(text="⬅️ Back to list", callback_data=f"files:page:{page}")],
        ]
    )



def _merge_execution_options(
    current: ExecutionOptions,
    requested: ExecutionOptions,
) -> ExecutionOptions:
    return ExecutionOptions(
        copy_original=current.copy_original or requested.copy_original,
        output_name=requested.output_name or current.output_name,
        repeated_from_operation_id=(
            requested.repeated_from_operation_id
            or current.repeated_from_operation_id
        ),
    )


def _with_execution_options(
    operation: PendingOperation,
    options: ExecutionOptions,
) -> PendingOperation:
    return replace(operation, execution=options)


def _unique_output_path(
    directory: Path,
    filename: str,
    source_path: Path,
    *,
    allow_same: bool = True,
) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    requested = directory / filename
    same_as_source = requested.resolve() == source_path.resolve()
    if same_as_source and allow_same:
        return requested.resolve()
    if not requested.exists() and not same_as_source:
        return requested.resolve()
    stem = requested.stem
    suffix = requested.suffix
    for index in range(2, 1000):
        candidate = directory / f"{stem} ({index}){suffix}"
        if not candidate.exists():
            return candidate.resolve()
    raise FileExistsError("Could not allocate a unique output filename.")


def _copy_target_path(source: DataSource, options: ExecutionOptions) -> Path:
    filename = normalized_output_filename(options.output_name, source.path)
    if not filename:
        filename = f"{source.path.stem}_result{source.path.suffix}"
    directory = source.path.parent if source.origin == SourceOrigin.LOCAL else catalog.workspace_root
    return _unique_output_path(
        directory, filename, source.path, allow_same=False
    )


def _rename_target_path(source_path: Path, requested_name: str) -> Path:
    filename = normalized_output_filename(requested_name, source_path)
    if not filename:
        raise ValueError("Could not parse the new filename.")
    return _unique_output_path(
        source_path.parent, filename, source_path, allow_same=True
    )


def _copy_source_file(source: DataSource, target: Path) -> None:
    if source.kind == SourceKind.SQLITE:
        copy_sqlite_database(source.path, target)
    else:
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source.path, target)


def _rename_source_file(source_path: Path, target_path: Path) -> None:
    if source_path.resolve() == target_path.resolve():
        return
    source_path.replace(target_path)
    # SQLite sidecars normally do not exist because the agent closes connections,
    # but keep a rename safe if one is present.
    for suffix in ("-wal", "-shm", "-journal"):
        old_sidecar = Path(str(source_path) + suffix)
        if old_sidecar.exists():
            old_sidecar.replace(Path(str(target_path) + suffix))


def _operation_before_metadata(operation: PendingOperation) -> SourceMetadata:
    if isinstance(operation, PendingBatchOperation):
        if not operation.steps:
            raise ValueError("An empty batch has no source metadata.")
        return operation.steps[0].metadata
    return operation.metadata


def _metadata_rows(metadata: SourceMetadata) -> int:
    if isinstance(metadata, WorkbookMetadata):
        return sum(sheet.data_rows for sheet in metadata.sheets)
    return sum(table.row_count for table in metadata.tables)


def _summary_steps(operation: PendingOperation) -> list[SinglePendingOperation]:
    if isinstance(operation, PendingBatchOperation):
        return operation.steps
    return [operation]


def _step_result_summary(step: SinglePendingOperation) -> str | None:
    action = step.plan.action
    preview = step.preview
    if action == Action.DEDUPLICATE:
        return f"Duplicates removed: <b>{preview.matched_rows}</b>"
    if action == Action.DELETE_ROWS:
        filters = getattr(step.plan, "filters", [])
        operator = getattr(filters[0], "operator", None) if len(filters) == 1 else None
        if len(filters) == 1 and getattr(operator, "value", None) == "is_empty":
            column = getattr(filters[0], "column", None)
            header = getattr(column, "header", None) or getattr(column, "name", None) or "values"
            return f"Removed empty {esc(str(header))}: <b>{preview.matched_rows}</b>"
        return f"Rows deleted: <b>{preview.matched_rows}</b>"
    if action == Action.RENAME_COLUMN and isinstance(step, PendingExcelOperation):
        old = step.plan.target_columns[0].header if step.plan.target_columns else "—"
        new = step.plan.new_column_name or "—"
        return f"Renamed: <code>{esc(old)} → {esc(new)}</code>"
    if action == Action.UPDATE_ROWS:
        assignments = getattr(step.plan, "assignments", [])
        rendered: list[str] = []
        for assignment in assignments[:3]:
            column = getattr(assignment, "column", None)
            name = getattr(column, "header", None) or getattr(column, "name", None) or "field"
            rendered.append(f"{esc(str(name))} → {esc(str(assignment.value))}")
        tail = " · " + ", ".join(rendered) if rendered else ""
        return f"Rows updated: <b>{preview.matched_rows}</b>{tail}"
    if action == Action.CLEAR_VALUES:
        return f"Cells cleared: <b>{preview.affected_cells}</b>"
    if action == Action.DROP_COLUMN and isinstance(step, PendingExcelOperation):
        name = step.plan.target_columns[0].header if step.plan.target_columns else "—"
        return f"Column dropped: <code>{esc(name)}</code>"
    if action == Action.ADD_COLUMN and isinstance(step, PendingExcelOperation):
        return f"Column added: <code>{esc(step.plan.new_column_name or '—')}</code>"
    if action == Action.REPLACE_ALL_VALUES:
        return f"Cells replaced: <b>{preview.affected_cells}</b>"
    return None


def completed_summary_text(
    operation: PendingOperation,
    refreshed_source: DataSource,
    refreshed_metadata: SourceMetadata,
    outcome: BatchExecutionOutcome | ExcelExecutionOutcome | SqlExecutionOutcome,
) -> str:
    before_rows = _metadata_rows(_operation_before_metadata(operation))
    after_rows = _metadata_rows(refreshed_metadata)
    lines = [
        "<b>Done</b>",
        f"{refreshed_source.origin_icon} <code>{esc(refreshed_source.display_name)}</code>",
        "",
        f"Rows before: <b>{before_rows}</b>",
        f"Rows after: <b>{after_rows}</b>",
    ]
    details = [item for item in (_step_result_summary(step) for step in _summary_steps(operation)) if item]
    if details:
        lines.extend(["", *details])
    if operation.execution.copy_original:
        lines.extend(["", "Original unchanged · a separate copy was created."])
    if operation.execution.repeated_from_operation_id:
        lines.append(
            "Repeated operation: "
            f"<code>#{esc(operation.execution.repeated_from_operation_id)}</code>"
        )
    lines.extend(
        [
            "",
            f"Snapshot: <code>#{esc(operation.operation_id)}</code>",
        ]
    )
    return "\n".join(lines)


async def send_latest_processed_file(message: Message, user_id: int) -> bool:
    record = repository.get_latest_completed_file_operation(user_id)
    if record is None or not record.result_path:
        await message.answer("There is no processed file to send yet.")
        return True
    path = Path(record.result_path).resolve()
    if not path.is_file():
        await message.answer("The latest processed file is no longer available on disk.")
        return True
    await message.answer_document(
        document=FSInputFile(path, filename=path.name),
        caption=f"📎 <b>Latest processed file</b>\n<code>{esc(path.name)}</code>",
    )
    try:
        repository.add_event(
            record.operation_id,
            user_id,
            "result_file_sent",
            {"path": str(path), "filename": path.name, "requested_as_latest": True},
        )
    except Exception:
        pass
    return True


async def rename_latest_processed_file(
    message: Message,
    user_id: int,
    requested_name: str,
    bot: Bot,
) -> bool:
    record = repository.get_latest_completed_file_operation(user_id)
    if record is None or not record.result_path:
        await message.answer("There is no result file to rename yet.")
        return True
    old_path = Path(record.result_path).resolve()
    if not old_path.is_file():
        await message.answer("The latest result file is no longer available on disk.")
        return True
    new_path = _rename_target_path(old_path, requested_name)
    if new_path == old_path:
        await message.answer(f"The file is already named <code>{esc(old_path.name)}</code>.")
        return True
    async with _lock_for(old_path):
        await asyncio.to_thread(_rename_source_file, old_path, new_path)
        repository.relocate_file_path(user_id, old_path, new_path, record.operation_id)
        source = catalog.find_by_path(user_id, new_path)
        if source is None:
            raise RuntimeError("The renamed file did not appear in the catalog.")
        metadata = await inspect_source(source)
        repository.set_active_source(
            user_id,
            source.path,
            source.display_name,
            _catalog_metadata(source, metadata),
            record.operation_id,
        )
    await message.answer(
        "<b>File renamed</b>\n"
        f"<code>{esc(old_path.name)} → {esc(new_path.name)}</code>",
        reply_markup=completed_operation_keyboard(record.operation_id),
    )
    try:
        await ensure_files_dashboard(bot, message.chat.id, user_id)
    except Exception:
        pass
    return True


def undo_more_keyboard(page: int = 0) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="↩️ Undo another", callback_data="undo:latest")],
            [InlineKeyboardButton(text="⬅️ Back to list", callback_data=f"files:page:{page}")],
        ]
    )


def source_choice_keyboard(sources: list[DataSource]) -> InlineKeyboardMarkup:
    rows = [
        [
            InlineKeyboardButton(
                text=f"{source.origin_icon} {shorten(source.relative_name, 38)}",
                callback_data=f"files:open:{source.source_id}:0",
            )
        ]
        for source in sources[:8]
    ]
    rows.append([InlineKeyboardButton(text="📂 All files", callback_data="files:page:0")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def files_text(
    sources: list[DataSource],
    page: int,
    active_name: str | None,
    refreshed: bool = False,
) -> str:
    pages = max((len(sources) - 1) // FILES_PER_PAGE + 1, 1)
    safe_page = min(max(page, 0), pages - 1)
    prefix = "✅ Refreshed\n" if refreshed else ""
    active = f"<code>{esc(active_name)}</code>" if active_name else "not selected"
    return (
        prefix
        + "<b>Files</b>\n"
        + f"Active: {active}\n"
        + f"{len(sources)} sources · {safe_page + 1}/{pages}\n\n"
        + "Choose a file."
    )


def _catalog_metadata(source: DataSource, metadata: SourceMetadata) -> dict:
    return {
        "source_id": source.source_id,
        "origin": source.origin.value,
        "kind": source.kind.value,
        "size_bytes": source.size_bytes,
        "modified_ns": source.modified_ns,
        "content": metadata.model_dump(mode="json"),
    }


def _cached_metadata(source: DataSource, payload: dict) -> SourceMetadata | None:
    if (
        payload.get("kind") != source.kind.value
        or payload.get("size_bytes") != source.size_bytes
        or payload.get("modified_ns") != source.modified_ns
    ):
        return None
    content = payload.get("content")
    if not isinstance(content, dict):
        return None
    try:
        if source.kind == SourceKind.EXCEL:
            return WorkbookMetadata.model_validate(content)
        if source.kind == SourceKind.SQLITE:
            return SqliteMetadata.model_validate(content)
    except Exception:
        return None
    return None


async def inspect_source(source: DataSource) -> SourceMetadata:
    if source.kind == SourceKind.EXCEL:
        worker = asyncio.to_thread(inspect_workbook, source.path, source.display_name)
    elif source.kind == SourceKind.SQLITE:
        worker = asyncio.to_thread(inspect_sqlite, source.path, source.display_name)
    else:
        raise ValueError(
            f"The {source.kind_label} adapter currently lists this file, "
            "but does not manage its data yet."
        )
    try:
        return await asyncio.wait_for(worker, timeout=SOURCE_INSPECTION_TIMEOUT_SECONDS)
    except TimeoutError as error:
        raise TimeoutError(
            "Schema inspection took more than 30 seconds. "
            "The file is too large or damaged; the operation was stopped."
        ) from error


async def activate_source(
    user_id: int,
    source: DataSource,
    operation_id: str | None = None,
    cancel_pending: bool = True,
) -> ActiveSource:
    metadata = await inspect_source(source)
    if cancel_pending:
        repository.cancel_open_operations(
            user_id,
            "The user selected another source.",
        )
    repository.set_active_source(
        user_id,
        source.path,
        source.display_name,
        _catalog_metadata(source, metadata),
        operation_id,
    )
    return ActiveSource(source=source, metadata=metadata)


async def get_active_source(user_id: int) -> ActiveSource | None:
    record = repository.get_active_source(user_id)
    if record is None:
        return None
    source = catalog.find_by_path(user_id, Path(record.file_path))
    if source is None:
        repository.clear_active_source(
            user_id,
            "The file was deleted, renamed, or moved outside the allowed catalog.",
        )
        return None

    cached = _cached_metadata(source, record.metadata)
    if cached is not None:
        return ActiveSource(source=source, metadata=cached)

    try:
        metadata = await inspect_source(source)
    except Exception:
        repository.clear_active_source(user_id, "The source is no longer readable.")
        raise
    repository.set_active_source(
        user_id,
        source.path,
        source.display_name,
        _catalog_metadata(source, metadata),
    )
    return ActiveSource(source=source, metadata=metadata)


def source_header(source: DataSource) -> list[str]:
    return [
        f"<b>{source.origin_icon} {esc(source.display_name)}</b>",
        f"{esc(source.kind_label)} · {esc(human_size(source.size_bytes))}",
        "",
    ]


def workbook_text(source: DataSource, metadata: WorkbookMetadata) -> str:
    lines = source_header(source)
    lines.append(f"Sheets: {len(metadata.sheets)}")
    for sheet in metadata.sheets[:5]:
        columns = ", ".join(esc(item.header[:45]) for item in sheet.columns[:10])
        lines.extend(
            [
                "",
                f"<b>{esc(sheet.name)}</b> · {sheet.data_rows} rows",
                f"{columns or 'No columns found'}",
            ]
        )
    lines.extend(["", "You can send one or multiple tasks in one message."])
    return "\n".join(lines)


def sqlite_text(source: DataSource, metadata: SqliteMetadata) -> str:
    lines = source_header(source)
    lines.append(f"Tables: {len(metadata.tables)}")
    for table in metadata.tables[:8]:
        columns = ", ".join(esc(item.name) for item in table.columns[:10])
        lines.extend(
            [
                "",
                f"<b>{esc(table.name)}</b> · {table.row_count} rows",
                f"{columns or 'No columns found'}",
            ]
        )
    lines.extend(["", "You can send one or multiple tasks in one message."])
    return "\n".join(lines)


def active_source_text(active: ActiveSource) -> str:
    if isinstance(active.metadata, WorkbookMetadata):
        return workbook_text(active.source, active.metadata)
    return sqlite_text(active.source, active.metadata)













def unsupported_source_text(source: DataSource) -> str:
    return "\n".join(
        [
            "<b>File found</b>",
            "",
            f"{source.origin_icon} {esc(source.display_name)}",
            f"Type: <code>{esc(source.kind_label)}</code>",
            f"Size: {esc(human_size(source.size_bytes))}",
            "",
            "It is visible in the file dashboard, but the write adapter for this format "
            "is not implemented yet. Excel (.xlsx) and SQLite "
            "(.db/.sqlite/.sqlite3) are currently supported.",
        ]
    )


def _single_proposal_lines(operation: SinglePendingOperation) -> list[str]:
    if isinstance(operation, PendingExcelOperation):
        target = (
            "entire workbook"
            if operation.plan.action == Action.REPLACE_ALL_VALUES
            else f"sheet “{esc(operation.plan.sheet_name)}”"
        )
    else:
        target = f"table “{esc(operation.plan.table_name)}”"
    return [
        f"<b>{esc(ACTION_LABELS.get(operation.plan.action, operation.plan.action.value))}</b> · {target}",
        esc(operation.preview.summary),
    ]


def _execution_note_lines(options: ExecutionOptions) -> list[str]:
    lines: list[str] = []
    if options.copy_original:
        lines.append("Result: separate copy · original remains unchanged.")
    if options.output_name:
        lines.append(f"Output name: <code>{esc(options.output_name)}</code>")
    if options.repeated_from_operation_id:
        lines.append(
            "Repeat of operation: "
            f"<code>#{esc(options.repeated_from_operation_id)}</code>"
        )
    return lines


def proposal_text(operation: PendingOperation) -> str:
    if isinstance(operation, PendingBatchOperation):
        affected = sum(step.preview.affected_cells for step in operation.steps)
        matched = sum(step.preview.matched_rows for step in operation.steps)
        lines = [
            f"<b>Plan ready · {len(operation.steps)} tasks</b>",
            f"{operation.source.origin_icon} <code>{esc(operation.source.display_name)}</code>",
            "",
        ]
        for index, step in enumerate(operation.steps, start=1):
            summary = _single_proposal_lines(step)
            lines.append(f"<b>{index}.</b> {summary[0].removeprefix('<b>').replace('</b>', '', 1)}")
            lines.append(f"   {summary[1]}")
        lines.extend(
            [
                "",
                f"Total: rows {matched}, cells {affected}.",
            ]
        )
        notes = _execution_note_lines(operation.execution)
        if notes:
            lines.extend(["", *notes])
        return "\n".join(lines)

    lines = [
        "<b>Plan ready</b>",
        f"{operation.source.origin_icon} <code>{esc(operation.source.display_name)}</code>",
        "",
        *_single_proposal_lines(operation),
    ]
    notes = _execution_note_lines(operation.execution)
    if notes:
        lines.extend(["", *notes])
    return "\n".join(lines)


def _append_preview_rows(lines: list[str], operation: SinglePendingOperation, limit: int = 6) -> None:
    preview = operation.preview
    for row in preview.rows[:limit]:
        cells: list[str] = []
        if isinstance(preview, ExcelOperationPreview):
            row_label = (
                f"{row.sheet_name} · row {row.row_number}"
                if row.sheet_name
                else f"Row {row.row_number}"
            )
            for cell in row.cells:
                value = f"{cell.column_header}: {cell.before}"
                if cell.after is not None:
                    value += f" → {cell.after}"
                cells.append(esc(value))
        else:
            row_label = f"rowid {row.rowid}"
            for cell in row.cells:
                value = f"{cell.column_name}: {cell.before}"
                if cell.after is not None:
                    value += f" → {cell.after}"
                cells.append(esc(value))
        candidate = f"<b>{esc(row_label)}</b>: " + "; ".join(cells)
        if len("\n".join([*lines, candidate])) > 3600:
            lines.append("…preview truncated.")
            return
        lines.append(candidate)
    if preview.matched_rows > min(len(preview.rows), limit):
        lines.append(f"…and {preview.matched_rows - min(len(preview.rows), limit)} more rows.")


def preview_text(operation: PendingOperation) -> str:
    if isinstance(operation, PendingBatchOperation):
        lines = [
            f"<b>Preview · {len(operation.steps)} tasks</b>",
            f"{operation.source.origin_icon} {esc(operation.source.display_name)}",
        ]
        for index, step in enumerate(operation.steps, start=1):
            lines.extend(["", f"<b>{index}. {esc(ACTION_LABELS.get(step.plan.action, step.plan.action.value))}</b>"])
            lines.append(esc(step.preview.summary))
            _append_preview_rows(lines, step, limit=3)
            if len("\n".join(lines)) > 3600:
                lines.append("…remaining tasks are hidden due to Telegram message limits.")
                break
        return "\n".join(lines)

    preview = operation.preview
    target_label = "Sheet" if isinstance(operation, PendingExcelOperation) else "Table"
    target_name = preview.sheet_name if isinstance(preview, ExcelOperationPreview) else preview.table_name
    lines = [
        "<b>Preview</b>",
        f"{operation.source.origin_icon} {esc(operation.source.display_name)} · {target_label.lower()} {esc(target_name)}",
        esc(preview.summary),
        "",
    ]
    _append_preview_rows(lines, operation)
    return "\n".join(lines)


def _single_pending_to_payload(operation: SinglePendingOperation) -> dict:
    source = operation.source
    common = {
        "source_id": source.source_id,
        "source_path": str(source.path),
        "source_origin": source.origin.value,
        "source_kind": source.kind.value,
        "display_name": source.display_name,
        "execution": operation.execution.model_dump(mode="json"),
    }
    if isinstance(operation, PendingExcelOperation):
        primary = next(
            iter(
                operation.plan.target_columns
                or operation.plan.deduplicate_columns
                or operation.plan.selected_columns
            ),
            None,
        )
        return {
            **common,
            "engine": "excel",
            "sheet_name": operation.plan.sheet_name,
            "column_header": primary.header if primary else None,
            "metadata": operation.metadata.model_dump(mode="json"),
            "plan": operation.plan.model_dump(mode="json"),
            "preview": operation.preview.model_dump(mode="json"),
        }
    return {
        **common,
        "engine": "sqlite",
        "table_name": operation.plan.table_name,
        "metadata": operation.metadata.model_dump(mode="json"),
        "plan": operation.plan.model_dump(mode="json"),
        "preview": operation.preview.model_dump(mode="json"),
    }


def pending_to_payload(operation: PendingOperation) -> dict:
    if not isinstance(operation, PendingBatchOperation):
        return _single_pending_to_payload(operation)
    return {
        "engine": "batch",
        "source_id": operation.source.source_id,
        "source_path": str(operation.source.path),
        "source_origin": operation.source.origin.value,
        "source_kind": operation.source.kind.value,
        "display_name": operation.source.display_name,
        "task_count": len(operation.steps),
        "execution": operation.execution.model_dump(mode="json"),
        "steps": [_single_pending_to_payload(step) for step in operation.steps],
    }


def _single_pending_from_payload(
    record: OperationRecord,
    payload: dict,
    source: DataSource,
    operation_id: str,
) -> SinglePendingOperation:
    engine = payload.get("engine")
    if engine == "excel" and source.kind == SourceKind.EXCEL:
        return PendingExcelOperation(
            operation_id=operation_id,
            user_id=record.user_id,
            source=source,
            metadata=WorkbookMetadata.model_validate(payload["metadata"]),
            plan=ExcelOperationPlan.model_validate(payload["plan"]),
            preview=ExcelOperationPreview.model_validate(payload["preview"]),
            created_at=record.created_at,
            execution=ExecutionOptions.model_validate(payload.get("execution", {})),
        )
    if engine == "sqlite" and source.kind == SourceKind.SQLITE:
        return PendingSqlOperation(
            operation_id=operation_id,
            user_id=record.user_id,
            source=source,
            metadata=SqliteMetadata.model_validate(payload["metadata"]),
            plan=SqlOperationPlan.model_validate(payload["plan"]),
            preview=SqlOperationPreview.model_validate(payload["preview"]),
            created_at=record.created_at,
            execution=ExecutionOptions.model_validate(payload.get("execution", {})),
        )
    raise ValueError("The saved step uses an unknown or incompatible engine.")


def pending_from_record(record: OperationRecord) -> PendingOperation:
    payload = record.plan
    source_id = str(payload.get("source_id", ""))
    source = catalog.get(record.user_id, source_id)
    if source is None:
        raise ValueError("The source was deleted, renamed, or is no longer allowed.")
    if source.path != Path(str(payload.get("source_path", ""))).resolve():
        raise ValueError("The source path does not match the saved plan.")
    if payload.get("engine") == "batch":
        raw_steps = payload.get("steps")
        if not isinstance(raw_steps, list) or not raw_steps:
            raise ValueError("The saved batch is empty.")
        steps = [
            _single_pending_from_payload(record, item, source, f"{record.operation_id}.{index}")
            for index, item in enumerate(raw_steps, start=1)
            if isinstance(item, dict)
        ]
        if len(steps) != len(raw_steps):
            raise ValueError("The saved batch is corrupted.")
        return PendingBatchOperation(
            operation_id=record.operation_id,
            user_id=record.user_id,
            source=source,
            steps=steps,
            created_at=record.created_at,
            execution=ExecutionOptions.model_validate(payload.get("execution", {})),
        )
    return _single_pending_from_payload(record, payload, source, record.operation_id)


def history_text(records: list[OperationRecord]) -> str:
    if not records:
        return "<b>History is empty</b>\nNo text commands have been processed yet."
    lines = ["<b>Recent operations</b>", ""]
    for record in records:
        payload = record.plan
        legacy_plan = payload.get("excel_plan", {})
        target = (
            payload.get("sheet_name")
            or payload.get("table_name")
            or legacy_plan.get("sheet_name")
        )
        display_name = (
            payload.get("display_name")
            or payload.get("original_name")
            or "—"
        )
        lines.extend(
            [
                f"<code>{esc(record.operation_id)}</code> — "
                f"{esc(STATUS_LABELS.get(record.status, record.status))}",
                f"{format_utc(record.created_at)} · <code>{esc(record.action)}</code>",
                f"Command: {esc(shorten(record.command_text, 65))}",
                f"Source: {esc(display_name)}"
                + (f" → {esc(target)}" if target else ""),
                "",
            ]
        )
    lines.append("Details: <code>/audit ID</code>")
    return "\n".join(lines)


def event_extra(event: AuditEventRecord) -> str:
    if event.event_type == "intent_parsed":
        return f" — {event.details.get('action', 'unknown')}"
    if event.event_type == "plan_created":
        target = event.details.get("sheet_name") or event.details.get("table_name") or ""
        return f" — {target}"
    if event.event_type in {
        "operation_rejected",
        "clarification_required",
        "planning_failed",
        "execution_failed",
        "execution_interrupted",
    }:
        reason = event.details.get("reason") or event.details.get("error")
        return f" — {shorten(str(reason), 100)}" if reason else ""
    return ""


def audit_text(record: OperationRecord, events: list[AuditEventRecord]) -> str:
    lines = [
        "<b>Operation audit log</b>",
        "",
        f"ID: <code>{esc(record.operation_id)}</code>",
        f"Status: {esc(STATUS_LABELS.get(record.status, record.status))}",
        f"Action: <code>{esc(record.action)}</code>",
        f"Command: {esc(record.command_text)}",
        f"Created: {format_utc(record.created_at)}",
        "",
        "<b>Events</b>",
    ]
    lines.extend(
        f"{event.created_at.astimezone(timezone.utc).strftime('%H:%M:%S')} — "
        f"{esc(EVENT_LABELS.get(event.event_type, event.event_type))}"
        f"{esc(event_extra(event))}"
        for event in events
    )
    if record.error_message:
        lines.extend(["", f"Reason: {esc(record.error_message)}"])
    return "\n".join(lines)


async def ensure_files_dashboard(
    bot: Bot,
    chat_id: int,
    user_id: int,
    page: int | None = None,
    refreshed: bool = False,
    target_message_id: int | None = None,
    force_new: bool = False,
) -> int:
    sources = catalog.list_sources(user_id)
    active_record = repository.get_active_source(user_id)
    active_path = Path(active_record.file_path) if active_record else None
    dashboard = repository.get_dashboard(user_id)
    requested_page = page if page is not None else (dashboard.page if dashboard else 0)
    pages = max((len(sources) - 1) // FILES_PER_PAGE + 1, 1)
    safe_page = min(max(requested_page, 0), pages - 1)
    text = files_text(
        sources,
        safe_page,
        active_record.original_name if active_record else None,
        refreshed,
    )
    markup = files_keyboard(sources, safe_page, active_path)

    message_id: int | None = None
    editable_message_id = target_message_id
    if (
        editable_message_id is None
        and not force_new
        and dashboard
        and dashboard.chat_id == chat_id
    ):
        editable_message_id = dashboard.message_id

    if editable_message_id is not None:
        try:
            await bot.edit_message_text(
                chat_id=chat_id,
                message_id=editable_message_id,
                text=text,
                reply_markup=markup,
            )
            message_id = editable_message_id
        except Exception as error:
            if "message is not modified" in str(error).casefold():
                message_id = editable_message_id
            elif dashboard and dashboard.message_id == editable_message_id:
                repository.clear_dashboard(user_id)

    if message_id is None:
        sent = await bot.send_message(chat_id=chat_id, text=text, reply_markup=markup)
        message_id = sent.message_id

    if (
        dashboard
        and dashboard.chat_id == chat_id
        and dashboard.message_id != message_id
    ):
        try:
            await bot.unpin_chat_message(
                chat_id=chat_id,
                message_id=dashboard.message_id,
            )
        except Exception:
            pass

    repository.set_dashboard(user_id, chat_id, message_id, safe_page)
    try:
        await bot.pin_chat_message(
            chat_id=chat_id,
            message_id=message_id,
            disable_notification=True,
        )
    except Exception:
        # Дисплей остаётся рабочим, даже если Telegram отклонил повторный pin.
        pass
    return message_id


@router.message(CommandStart())
async def start(message: Message, bot: Bot) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer(
            "Access denied.\n"
            f"Your Telegram ID: <code>{user_id}</code>"
        )
        return
    if not onboarding_accepted(user_id):
        await message.answer(
            onboarding_text(),
            reply_markup=onboarding_keyboard(),
        )
        return
    await ensure_files_dashboard(bot, message.chat.id, user_id, force_new=True)


@router.callback_query(F.data == "onboarding:accept")
async def onboarding_callback(callback: CallbackQuery, bot: Bot) -> None:
    user_id = callback.from_user.id
    if not allowed(user_id):
        await callback.answer("Access denied.", show_alert=True)
        return
    repository.accept_onboarding(user_id, CURRENT_ONBOARDING_VERSION)
    await callback.answer("Working mode enabled.")
    if callback.message is None:
        return
    try:
        await callback.message.edit_text("✅ <b>Ready</b> · rules accepted.")
    except Exception:
        pass
    await ensure_files_dashboard(
        bot,
        callback.message.chat.id,
        user_id,
        force_new=True,
    )


@router.message(Command("rules"))
async def rules(message: Message) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    await message.answer(onboarding_text())


@router.message(Command("whoami"))
async def whoami(message: Message) -> None:
    user_id = message.from_user.id if message.from_user else 0
    await message.answer(f"Your Telegram ID: <code>{user_id}</code>")


@router.message(Command("health"))
async def health(message: Message) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    status = await message.answer("Checking services…")
    try:
        repository.healthcheck()
        await intent_parser.ping()
        sources = catalog.list_sources(user_id)
        await status.edit_text(
            "Telegram: <b>OK</b>\nOllama: <b>OK</b>\nSQLite audit: <b>OK</b>\n"
            f"Catalog: <b>OK</b> ({len(sources)} files)\n"
            f"Model: <code>{esc(settings.ollama_model)}</code>"
        )
    except Exception as error:
        await status.edit_text(
            "<b>Health check failed</b>\n"
            f"<code>{esc(type(error).__name__)}: {esc(error)}</code>"
        )


@router.message(Command("folder"))
async def folder(message: Message) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    await message.answer(
        "<b>Managed workspace</b>\n\n"
        f"<code>{esc(catalog.workspace_root)}</code>\n\n"
        "To use a different folder, set this in <code>.env</code>:\n"
        "<code>DATA_WORKSPACE_DIR=C:\\path\\to\\workspace</code>\n"
        "and restart the bot. The LLM cannot access files outside this folder."
    )


@router.message(Command("files"))
async def files(message: Message, bot: Bot) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    await ensure_files_dashboard(bot, message.chat.id, user_id, force_new=True)


@router.message(Command("refresh"))
async def refresh(message: Message, bot: Bot) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    await ensure_files_dashboard(
        bot,
        message.chat.id,
        user_id,
        refreshed=True,
        force_new=True,
    )


@router.message(Command("source"))
async def source(message: Message) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    try:
        active = await get_active_source(user_id)
    except Exception as error:
        await message.answer(
            f"The source is no longer readable: <code>{esc(type(error).__name__)}: "
            f"{esc(error)}</code>",
            reply_markup=main_menu_keyboard(),
        )
        return
    if active is None:
        await message.answer(
            "No source selected. Open the file dashboard.",
            reply_markup=main_menu_keyboard(),
        )
        return
    await message.answer(active_source_text(active), reply_markup=back_to_files_keyboard())


@router.message(Command("history"))
async def history(message: Message) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    await message.answer(history_text(repository.list_operations(user_id, limit=10)))


@router.message(Command("audit"))
async def audit(message: Message) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    parts = (message.text or "").split(maxsplit=1)
    operation_id = parts[1].strip() if len(parts) == 2 else ""
    record = (
        repository.get_operation(operation_id)
        if operation_id
        else repository.get_latest_operation(user_id)
    )
    if record is None or record.user_id != user_id:
        await message.answer("Operation not found. Use <code>/history</code>.")
        return
    await message.answer(
        audit_text(record, repository.list_events(record.operation_id, user_id))
    )


@router.callback_query(F.data.startswith("menu:"))
async def menu_callback(callback: CallbackQuery) -> None:
    if not allowed(callback.from_user.id):
        await callback.answer("Access denied.", show_alert=True)
        return
    if not await require_onboarding_callback(callback):
        return
    action = (callback.data or "").split(":", 1)[-1]
    if action == "history":
        await callback.answer()
        if callback.message:
            await callback.message.answer(
                history_text(repository.list_operations(callback.from_user.id, limit=10))
            )
        return
    if action == "health":
        await callback.answer("Use /health", show_alert=True)
        return
    await callback.answer("Unknown button.", show_alert=True)








async def execute_undo(
    user_id: int,
    status: Message,
    bot: Bot,
    target_operation_id: str | None = None,
) -> None:
    target: OperationRecord | None = None
    if target_operation_id and target_operation_id != "latest":
        candidate = repository.get_operation(target_operation_id)
        if candidate is not None and candidate.user_id == user_id:
            target = candidate
    if target is None:
        active = await get_active_source(user_id)
        if active is None:
            await status.edit_text("Select the file whose latest change you want to undo first.")
            return
        target = repository.get_latest_undoable_operation(user_id, active.source.path)

    if target is None or not target.result_path or not target.snapshot_path:
        await status.edit_text("There is no change for this file that can be safely undone.")
        return
    source_path = Path(target.result_path).resolve()
    latest = repository.get_latest_undoable_operation(user_id, source_path)
    if latest is None or latest.operation_id != target.operation_id:
        await status.edit_text("This button is stale. Only the latest change to this file can be undone.")
        return
    if not target.result_sha256:
        await status.edit_text(
            "This change predates safe Undo and has no result checksum. Automatic rollback is blocked."
        )
        return

    undo_operation_id = uuid4().hex[:12]
    repository.cancel_open_operations(user_id, "Undo of the latest change started.")
    repository.create_request(undo_operation_id, user_id, f"undo operation {target.operation_id}")
    source_type = "excel" if source_path.suffix.casefold() == ".xlsx" else "sql"
    repository.save_intent(
        undo_operation_id, user_id, "undo", source_type,
        {"target_operation_id": target.operation_id},
    )
    repository.save_plan(
        undo_operation_id, user_id,
        {
            "action": "undo",
            "target_operation_id": target.operation_id,
            "snapshot_path": target.snapshot_path,
            "result_path": str(source_path),
        },
    )
    repository.claim_operation(undo_operation_id, user_id)

    try:
        await status.edit_text("↩️ Restoring the state before the latest change…")
        async with _lock_for(source_path):
            outcome = await asyncio.to_thread(
                restore_snapshot,
                source_path,
                Path(target.snapshot_path),
                expected_current_sha256=target.result_sha256,
            )
            source = next(
                (item for item in catalog.list_sources(user_id) if item.path.resolve() == source_path),
                None,
            )
            if source is None:
                raise RuntimeError("The restored file was not found in the catalog.")
            metadata = await inspect_source(source)
            repository.set_active_source(
                user_id, source.path, source.display_name, _catalog_metadata(source, metadata), undo_operation_id
            )
            repository.complete_operation(
                undo_operation_id,
                user_id,
                outcome.snapshot_path,
                source.path,
                outcome.result_sha256,
            )
            if not repository.mark_undone(target.operation_id, user_id, undo_operation_id):
                raise RuntimeError("The file was restored, but the original operation could not be marked as undone.")
            repository.add_event(
                undo_operation_id,
                user_id,
                "rollback_completed",
                {
                    "target_operation_id": target.operation_id,
                    "restored_from": Path(target.snapshot_path).name,
                },
            )

        await status.edit_text(
            "<b>Restored previous state</b>\n"
            f"<code>{esc(source.display_name)}</code> · before operation <code>{esc(target.operation_id)}</code>.",
            reply_markup=undo_more_keyboard(),
        )
        try:
            await ensure_files_dashboard(bot, status.chat.id, user_id)
        except Exception:
            pass
    except Exception as error:
        repository.fail_operation(undo_operation_id, user_id, f"{type(error).__name__}: {error}")
        await status.edit_text(
            "<b>Undo failed</b>\n"
            f"<code>{esc(type(error).__name__)}: {esc(error)}</code>"
        )


@router.callback_query(F.data.startswith("sendfile:"))
async def send_file_callback(callback: CallbackQuery) -> None:
    user_id = callback.from_user.id
    if not allowed(user_id):
        await callback.answer("Access denied.", show_alert=True)
        return
    if not await require_onboarding_callback(callback):
        return

    operation_id = (callback.data or "").split(":", 1)[-1].strip()
    record = repository.get_operation(operation_id)
    if record is None or record.user_id != user_id:
        await callback.answer("The file for this operation was not found.", show_alert=True)
        return
    if not record.result_path:
        await callback.answer("This operation has no completed file.", show_alert=True)
        return

    path = Path(record.result_path).resolve()
    if not path.is_file():
        await callback.answer("The file is no longer available on disk.", show_alert=True)
        return
    if callback.message is None:
        await callback.answer()
        return

    try:
        # Отправляем именно текущую версию файла по этому пути. Если после
        # операции файл редактировали ещё раз, пользователь получает свежий
        # вариант, что удобнее для мобильной работы и пересылки дальше.
        await callback.message.answer_document(
            document=FSInputFile(path, filename=path.name),
            caption=(
                "📎 <b>Current file version</b>\n"
                f"<code>{esc(path.name)}</code>"
            ),
        )
    except Exception as error:
        # Ошибка отправки документа не меняет статус уже выполненной операции.
        try:
            await callback.answer(
                f"Could not send file: {type(error).__name__}",
                show_alert=True,
            )
        except Exception:
            pass
        return

    try:
        await callback.answer("File sent.")
    except Exception:
        pass
    try:
        repository.add_event(
            operation_id,
            user_id,
            "result_file_sent",
            {"path": str(path), "filename": path.name},
        )
    except Exception:
        pass


@router.callback_query(F.data.startswith("undo:"))
async def undo_callback(callback: CallbackQuery, bot: Bot) -> None:
    user_id = callback.from_user.id
    if not allowed(user_id):
        await callback.answer("Access denied.", show_alert=True)
        return
    if not await require_onboarding_callback(callback):
        return
    if callback.message is None:
        await callback.answer()
        return
    target = (callback.data or "").split(":", 1)[-1]
    await callback.answer("Checking whether Undo is safe…")
    await execute_undo(user_id, callback.message, bot, target)


@router.callback_query(F.data.startswith("files:"))
async def files_callback(callback: CallbackQuery, bot: Bot) -> None:
    if not allowed(callback.from_user.id):
        await callback.answer("Access denied.", show_alert=True)
        return
    if not await require_onboarding_callback(callback):
        return
    parts = (callback.data or "").split(":")
    action = parts[1] if len(parts) > 1 else ""
    if action == "noop":
        await callback.answer()
        return
    if action in {"page", "refresh"} and len(parts) == 3:
        try:
            page = int(parts[2])
        except ValueError:
            page = 0
        await callback.answer("Scanning workspace…" if action == "refresh" else "")
        if callback.message:
            await ensure_files_dashboard(
                bot,
                callback.message.chat.id,
                callback.from_user.id,
                page,
                refreshed=action == "refresh",
                target_message_id=callback.message.message_id,
            )
        return
    if action != "open" or len(parts) != 4:
        await callback.answer("Invalid button.", show_alert=True)
        return
    source_id = parts[2]
    try:
        page = int(parts[3])
    except ValueError:
        page = 0
    selected = catalog.get(callback.from_user.id, source_id)
    if selected is None:
        await callback.answer("The file disappeared. Press “Refresh”.", show_alert=True)
        return
    if not selected.editable:
        await callback.answer("This format is currently read-only in the dashboard.")
        if callback.message:
            await callback.message.answer(unsupported_source_text(selected))
        return
    await callback.answer("Inspecting schema…")
    try:
        active = await activate_source(callback.from_user.id, selected)
        if callback.message:
            await ensure_files_dashboard(
                bot,
                callback.message.chat.id,
                callback.from_user.id,
                page,
                target_message_id=callback.message.message_id,
            )
            await callback.message.answer(active_source_text(active))
    except Exception as error:
        if callback.message:
            await callback.message.answer(
                "<b>Could not open file</b>\n"
                f"<code>{esc(type(error).__name__)}: {esc(error)}</code>"
            )


@router.message(F.document)
async def receive_file(message: Message, bot: Bot) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    document = message.document
    if document is None:
        return
    original_name = document.file_name or "uploaded_file"
    if document.file_size and document.file_size > MAX_TELEGRAM_FILE_SIZE_BYTES:
        await message.answer("Files larger than 20 MB are not supported yet.")
        return
    destination = catalog.allocate_telegram_upload(user_id, original_name)
    status = await message.answer("Saving as a <code>📨 TG</code> source…")
    try:
        await bot.download(document, destination=destination)
        selected = catalog.find_by_path(user_id, destination)
        if selected is None:
            raise RuntimeError("The uploaded file did not appear in the catalog.")
        if selected.editable:
            active = await activate_source(user_id, selected)
            text = (
                "<b>File added to the dashboard with the 📨 TG marker</b>\n\n"
                "It was not copied into the managed local workspace.\n\n"
                + active_source_text(active)
            )
        else:
            text = (
                "<b>File added to the dashboard with the 📨 TG marker</b>\n\n"
                "It was not copied into the managed local workspace.\n\n"
                + unsupported_source_text(selected)
            )
        await status.edit_text(text)
        await ensure_files_dashboard(
            bot,
            message.chat.id,
            user_id,
            refreshed=True,
            force_new=True,
        )
    except Exception as error:
        catalog.discard_telegram_upload(destination)
        await status.edit_text(
            "Could not save source.\n"
            f"<code>{esc(type(error).__name__)}: {esc(error)}</code>"
        )


async def execute_confirmed_operation(
    operation: PendingOperation,
    status_message: Message,
    bot: Bot,
) -> None:
    operation_id = operation.operation_id
    if datetime.now(timezone.utc) - operation.created_at > OPERATION_TTL:
        repository.expire_operation(operation_id, operation.user_id)
        await status_message.edit_text("The plan expired. Send the command again.")
        return
    if not repository.claim_operation(operation_id, operation.user_id):
        await status_message.edit_text("This operation has already been processed.")
        return

    created_copy_path: Path | None = None
    completed = False
    try:
        async with _lock_for(operation.source.path):
            task_count = len(operation.steps) if isinstance(operation, PendingBatchOperation) else 1
            options = operation.execution
            await status_message.edit_text(
                f"<b>Executing</b> · {task_count} "
                + ("task" if task_count == 1 else "tasks")
                + (" · on a copy…" if options.copy_original else "…")
            )

            execution_source = operation.source
            execution_operation = operation
            if options.copy_original:
                target_path = _copy_target_path(operation.source, options)
                await asyncio.to_thread(_copy_source_file, operation.source, target_path)
                created_copy_path = target_path
                execution_source = catalog.find_by_path(operation.user_id, target_path)
                if execution_source is None:
                    raise RuntimeError("The created copy did not appear in the file catalog.")
                execution_operation = replace(operation, source=execution_source)
                repository.add_event(
                    operation_id,
                    operation.user_id,
                    "result_copy_created",
                    {
                        "source_path": str(operation.source.path),
                        "copy_path": str(target_path),
                    },
                )

            if isinstance(execution_operation, PendingBatchOperation):
                if all(isinstance(step, PendingExcelOperation) for step in execution_operation.steps):
                    excel_steps = [
                        (step.plan, step.preview, step.metadata)
                        for step in execution_operation.steps
                        if isinstance(step, PendingExcelOperation)
                    ]
                    outcome: BatchExecutionOutcome | ExcelExecutionOutcome | SqlExecutionOutcome = await asyncio.to_thread(
                        execute_excel_batch,
                        execution_source.path,
                        execution_source.display_name,
                        execution_operation.operation_id,
                        excel_steps,
                    )
                    engine = "excel_batch"
                elif all(isinstance(step, PendingSqlOperation) for step in execution_operation.steps):
                    sql_steps = [
                        (step.plan, step.preview, step.metadata)
                        for step in execution_operation.steps
                        if isinstance(step, PendingSqlOperation)
                    ]
                    outcome = await asyncio.to_thread(
                        execute_sqlite_batch,
                        execution_source.path,
                        execution_source.display_name,
                        execution_operation.operation_id,
                        sql_steps,
                    )
                    engine = "sqlite_batch"
                else:
                    raise ValueError("A single batch cannot mix execution engines.")
                details = {
                    "engine": engine,
                    "task_count": task_count,
                    "changed_rows": outcome.changed_rows,
                    "affected_cells": outcome.affected_cells,
                }
                verify_details = {"result_verified": outcome.result_verified}
            elif isinstance(execution_operation, PendingExcelOperation):
                outcome = await asyncio.to_thread(
                    execute_operation,
                    execution_source.path,
                    execution_source.display_name,
                    execution_operation.operation_id,
                    execution_operation.plan,
                    execution_operation.preview,
                    execution_operation.metadata,
                )
                details = {
                    "engine": "excel",
                    "changed_rows": outcome.changed_rows,
                    "affected_cells": outcome.affected_cells,
                }
                verify_details = {
                    "source_updated": outcome.source_updated,
                    "result_verified": outcome.result_verified,
                }
            else:
                outcome = await asyncio.to_thread(
                    execute_sql_operation,
                    execution_source.path,
                    execution_operation.operation_id,
                    execution_operation.plan,
                    execution_operation.preview,
                    execution_operation.metadata,
                )
                details = {
                    "engine": "sqlite",
                    "changed_rows": outcome.changed_rows,
                    "affected_cells": outcome.affected_cells,
                }
                verify_details = {
                    "transaction_committed": outcome.transaction_committed,
                    "result_verified": outcome.result_verified,
                }

            repository.add_event(
                operation_id,
                operation.user_id,
                "snapshot_created",
                {"snapshot_name": outcome.snapshot_path.name},
            )
            repository.add_event(
                operation_id,
                operation.user_id,
                "change_applied",
                details,
            )
            repository.add_event(
                operation_id,
                operation.user_id,
                "verification_passed",
                verify_details,
            )

            final_path = execution_source.path.resolve()
            # If the user requested a human result name without copy mode, rename
            # only after the confirmed data operation succeeded and was verified.
            if options.output_name and not options.copy_original:
                renamed_path = _rename_target_path(final_path, options.output_name)
                if renamed_path != final_path:
                    await asyncio.to_thread(_rename_source_file, final_path, renamed_path)
                    repository.relocate_file_path(
                        operation.user_id,
                        final_path,
                        renamed_path,
                        operation_id,
                    )
                    final_path = renamed_path

            refreshed_source = catalog.find_by_path(operation.user_id, final_path)
            if refreshed_source is None:
                raise RuntimeError("The result source disappeared from the catalog.")
            refreshed_metadata = await inspect_source(refreshed_source)
            repository.set_active_source(
                operation.user_id,
                refreshed_source.path,
                refreshed_source.display_name,
                _catalog_metadata(refreshed_source, refreshed_metadata),
                operation_id,
            )
            result_sha = await asyncio.to_thread(sha256_file, refreshed_source.path)
            repository.complete_operation(
                operation_id,
                operation.user_id,
                outcome.snapshot_path,
                refreshed_source.path,
                result_sha,
            )
            repository.add_event(
                operation_id,
                operation.user_id,
                "source_updated",
                {
                    "path": str(refreshed_source.path),
                    "origin": refreshed_source.origin.value,
                    "copy_original": options.copy_original,
                    "output_name": options.output_name,
                },
            )
            completed = True

            await status_message.edit_text(
                completed_summary_text(
                    operation,
                    refreshed_source,
                    refreshed_metadata,
                    outcome,
                ),
                reply_markup=completed_operation_keyboard(operation_id),
            )
            try:
                await ensure_files_dashboard(
                    bot,
                    status_message.chat.id,
                    operation.user_id,
                )
            except Exception:
                pass
    except Exception as error:
        if created_copy_path is not None and not completed:
            try:
                created_copy_path.unlink(missing_ok=True)
                for suffix in ("-wal", "-shm", "-journal"):
                    Path(str(created_copy_path) + suffix).unlink(missing_ok=True)
            except Exception:
                pass
        error_message = f"{type(error).__name__}: {error}"
        repository.fail_operation(operation_id, operation.user_id, error_message)
        hint = (
            "\n\nIf the Excel file is open in Microsoft Excel, close it and confirm again."
            if isinstance(error, PermissionError)
            else ""
        )
        await status_message.edit_text(
            "<b>Execution failed</b>\n"
            f"<code>{esc(error_message)}</code>{hint}",
            reply_markup=operation_keyboard(operation_id),
        )


@router.callback_query(F.data.startswith("op:"))
async def operation_callback(callback: CallbackQuery, bot: Bot) -> None:
    if not allowed(callback.from_user.id):
        await callback.answer("Access denied.", show_alert=True)
        return
    if not await require_onboarding_callback(callback):
        return
    parts = (callback.data or "").split(":", maxsplit=2)
    if len(parts) != 3:
        await callback.answer("Invalid command.", show_alert=True)
        return
    _, callback_action, operation_id = parts
    record = repository.get_operation(operation_id)
    if record is None:
        await callback.answer("Operation not found.", show_alert=True)
        return
    if callback.from_user.id != record.user_id:
        await callback.answer("This operation belongs to another user.", show_alert=True)
        return
    if record.status not in {"planned", "previewed", "failed"}:
        label = STATUS_LABELS.get(record.status, record.status)
        await callback.answer(f"Operation is no longer available: {label}", show_alert=True)
        return
    try:
        operation = pending_from_record(record)
    except Exception as error:
        await callback.answer(f"Plan unavailable: {error}", show_alert=True)
        return
    if datetime.now(timezone.utc) - operation.created_at > OPERATION_TTL:
        repository.expire_operation(operation_id, operation.user_id)
        await callback.answer("The 15-minute confirmation window expired.", show_alert=True)
        if callback.message:
            await callback.message.edit_text("Operation cancelled due to timeout.")
        return
    if callback_action == "preview":
        repository.mark_previewed(operation_id, operation.user_id)
        await callback.answer()
        if callback.message:
            await callback.message.answer(preview_text(operation))
        return
    if callback_action == "cancel":
        cancelled = repository.cancel_operation(
            operation_id,
            operation.user_id,
            "The user pressed Cancel.",
        )
        if not cancelled:
            await callback.answer("The operation has already been processed.", show_alert=True)
            return
        await callback.answer("Cancelled.")
        if callback.message:
            await callback.message.edit_text("<b>Cancelled</b>")
        return
    if callback_action != "confirm":
        await callback.answer("Unknown action.", show_alert=True)
        return
    await callback.answer("Confirmation accepted.")
    if callback.message is None:
        return
    await execute_confirmed_operation(operation, callback.message, bot)


async def _resolve_command_source(
    user_id: int,
    intent: IntentDraft,
    command_text: str,
    operation_id: str,
) -> tuple[ActiveSource | None, list[DataSource]]:
    sources = catalog.list_sources(user_id)
    hint = intent.source_name_hint
    if intent.action == Action.OPEN_SOURCE and not hint:
        hint = intent.resource_name_hint

    if intent.action == Action.OPEN_SOURCE and hint:
        matches = match_sources(hint, sources)
        if len(matches) == 1:
            if not matches[0].editable:
                return None, matches
            return await activate_source(
                user_id,
                matches[0],
                operation_id,
                cancel_pending=False,
            ), []
        return None, matches

    active = await get_active_source(user_id)
    if active is not None:
        # Выбор в закреплённом интерфейсе — главный контекст. Переключаемся
        # внутри составной команды только при буквальном имени другого файла.
        if hint:
            matches = match_sources(hint, sources)
            if len(matches) == 1 and matches[0].editable:
                candidate = matches[0]
                command = command_text.casefold().replace("/", "\\")
                explicit_names = {
                    candidate.display_name.casefold(),
                    candidate.relative_name.casefold().replace("/", "\\"),
                }
                explicitly_named = any(
                    name and name in command and "." in Path(name).name
                    for name in explicit_names
                )
                if explicitly_named and candidate.path != active.source.path:
                    return await activate_source(
                        user_id,
                        candidate,
                        operation_id,
                        cancel_pending=False,
                    ), []
        return active, []

    if hint:
        matches = match_sources(hint, sources)
        if len(matches) == 1 and matches[0].editable:
            return await activate_source(
                user_id,
                matches[0],
                operation_id,
                cancel_pending=False,
            ), []
        return None, matches

    fallback_matches = match_sources(command_text, sources)
    if len(fallback_matches) == 1 and fallback_matches[0].editable:
        return await activate_source(
            user_id,
            fallback_matches[0],
            operation_id,
            cancel_pending=False,
        ), []
    return None, fallback_matches


class BatchClarificationError(ValueError):
    def __init__(self, task_number: int, reason: str, alternatives: list[str] | None = None) -> None:
        super().__init__(reason)
        self.task_number = task_number
        self.reason = reason
        self.alternatives = alternatives or []


def _replace_all_plan(intent: IntentDraft, task_text: str) -> ExcelOperationPlan | None:
    del task_text
    if intent.action != Action.REPLACE_ALL_VALUES or not intent.value_hints:
        return None
    return ExcelOperationPlan(
        action=Action.REPLACE_ALL_VALUES,
        resolved=True,
        sheet_name=None,
        replacement_value=intent.value_hints[0],
        confidence=max(intent.confidence, 0.95),
        resolution_note="The replacement value came from the structured AI intent.",
    )


async def _plan_excel_intent(
    intent: IntentDraft,
    task_text: str,
    metadata: WorkbookMetadata,
) -> ExcelOperationPlan:
    if intent.action not in SUPPORTED_EXCEL_ACTIONS:
        raise ValueError(f"Excel action {intent.action.value} is not implemented yet.")
    special = _replace_all_plan(intent, task_text)
    if special is not None:
        return special
    return await excel_planner.resolve(task_text, intent, metadata)


async def _plan_sql_intent(
    intent: IntentDraft,
    task_text: str,
    metadata: SqliteMetadata,
) -> SqlOperationPlan:
    if intent.action not in SUPPORTED_SQL_ACTIONS:
        raise ValueError(
            "SQLite currently supports select/update_rows/delete_rows, "
            f"but the interpreted action is {intent.action.value}."
        )
    return await sql_planner.resolve(task_text, intent, metadata)


def _batch_tasks_and_source_selector(
    intents: list[IntentDraft],
) -> tuple[list[IntentDraft], IntentDraft | None]:
    if len(intents) <= 1:
        return intents, None
    selectors = [intent for intent in intents if intent.action == Action.OPEN_SOURCE]
    tasks = [intent for intent in intents if intent.action != Action.OPEN_SOURCE]
    if not selectors or not tasks:
        return intents, None
    selector = selectors[0]
    source_hint = selector.source_name_hint or selector.resource_name_hint
    if source_hint:
        tasks = [
            task.model_copy(
                update={"source_name_hint": task.source_name_hint or source_hint}
            )
            for task in tasks
        ]
    return tasks, selector


async def _build_batch_operation(
    active: ActiveSource,
    intents: list[IntentDraft],
    operation_id: str,
    user_id: int,
    execution: ExecutionOptions | None = None,
) -> PendingBatchOperation:
    created_at = datetime.now(timezone.utc)
    steps: list[SinglePendingOperation] = []

    with tempfile.TemporaryDirectory(prefix="dataops-plan-batch-") as temporary:
        root = Path(temporary)
        if active.source.kind == SourceKind.EXCEL:
            stage_path = root / "stage.xlsx"
            shutil.copy2(active.source.path, stage_path)
            for index, intent in enumerate(intents, start=1):
                metadata = await asyncio.to_thread(
                    inspect_workbook, stage_path, active.source.display_name
                )
                task_text = intent.normalized_request
                plan = await _plan_excel_intent(intent, task_text, metadata)
                if not plan.resolved or (
                    plan.action != Action.REPLACE_ALL_VALUES and not plan.sheet_name
                ):
                    raise BatchClarificationError(
                        index,
                        plan.clarification_question
                        or plan.resolution_note
                        or "No unambiguous match was found.",
                        plan.alternative_matches,
                    )
                preview = await asyncio.to_thread(
                    build_operation_preview,
                    stage_path,
                    plan,
                    metadata,
                )
                step = PendingExcelOperation(
                    operation_id=f"{operation_id}.{index}",
                    user_id=user_id,
                    source=active.source,
                    metadata=metadata,
                    plan=plan,
                    preview=preview,
                    created_at=created_at,
                )
                steps.append(step)
                if plan.action in WRITE_ACTIONS and preview.has_changes:
                    await asyncio.to_thread(
                        execute_operation,
                        stage_path,
                        active.source.display_name,
                        f"{operation_id}-plan-{index}",
                        plan,
                        preview,
                        metadata,
                        root / "snapshots",
                    )
        elif active.source.kind == SourceKind.SQLITE:
            stage_path = root / "stage.sqlite3"
            await asyncio.to_thread(copy_sqlite_database, active.source.path, stage_path)
            for index, intent in enumerate(intents, start=1):
                metadata = await asyncio.to_thread(
                    inspect_sqlite, stage_path, active.source.display_name
                )
                task_text = intent.normalized_request
                plan = await _plan_sql_intent(intent, task_text, metadata)
                if not plan.resolved or not plan.table_name:
                    raise BatchClarificationError(
                        index,
                        plan.clarification_question
                        or plan.resolution_note
                        or "No unambiguous match was found.",
                    )
                preview = await asyncio.to_thread(
                    build_sql_preview,
                    stage_path,
                    plan,
                    metadata,
                )
                step = PendingSqlOperation(
                    operation_id=f"{operation_id}.{index}",
                    user_id=user_id,
                    source=active.source,
                    metadata=metadata,
                    plan=plan,
                    preview=preview,
                    created_at=created_at,
                )
                steps.append(step)
                if plan.action in SQL_WRITE_ACTIONS and preview.has_changes:
                    await asyncio.to_thread(
                        execute_sql_operation,
                        stage_path,
                        f"{operation_id}-plan-{index}",
                        plan,
                        preview,
                        metadata,
                        root / "snapshots",
                    )
        else:
            raise ValueError("The selected format has no execution engine yet.")

    return PendingBatchOperation(
        operation_id=operation_id,
        user_id=user_id,
        source=active.source,
        steps=steps,
        created_at=created_at,
        execution=execution or ExecutionOptions(),
    )


async def _handle_multi_intent_command(
    status: Message,
    user_id: int,
    operation_id: str,
    intents: list[IntentDraft],
    command_text: str,
    execution: ExecutionOptions,
) -> bool:
    tasks, selector = _batch_tasks_and_source_selector(intents)
    if len(tasks) <= 1:
        return False

    representative = selector or tasks[0]
    active, candidates = await _resolve_command_source(
        user_id,
        representative,
        command_text,
        operation_id,
    )
    if active is None:
        reason = (
            "Multiple similar files were found."
            if len(candidates) > 1
            else "Source not found."
        )
        repository.stop_operation(
            operation_id,
            user_id,
            "needs_clarification",
            reason,
            "clarification_required",
        )
        await status.edit_text(
            f"<b>Select a file</b>\n{esc(reason)}",
            reply_markup=(
                source_choice_keyboard(candidates) if candidates else main_menu_keyboard()
            ),
        )
        return True

    sources = catalog.list_sources(user_id)
    for task in tasks:
        if not task.source_name_hint:
            continue
        matches = match_sources(task.source_name_hint, sources)
        if len(matches) == 1 and matches[0].editable and matches[0].path != active.source.path:
            raise ValueError(
                "A batch currently operates on one source. "
                "Send separate messages for different files."
            )

    try:
        batch = await _build_batch_operation(
            active, tasks, operation_id, user_id, execution
        )
    except BatchClarificationError as error:
        repository.stop_operation(
            operation_id,
            user_id,
            "needs_clarification",
            error.reason,
            "clarification_required",
        )
        alternatives = "\n".join(f"• {esc(item)}" for item in error.alternatives)
        await status.edit_text(
            f"<b>Clarification required {error.task_number}</b>\n{esc(error.reason)}"
            + (f"\n\n{alternatives}" if alternatives else "")
        )
        return True
    repository.save_plan(operation_id, user_id, pending_to_payload(batch))

    has_write = any(
        (isinstance(step, PendingExcelOperation) and step.plan.action in WRITE_ACTIONS)
        or (isinstance(step, PendingSqlOperation) and step.plan.action in SQL_WRITE_ACTIONS)
        for step in batch.steps
    )
    has_changes = any(step.preview.has_changes for step in batch.steps if step.preview.is_write)

    if not has_write:
        repository.finish_without_file(
            operation_id,
            user_id,
            "read_completed",
            {"task_count": len(batch.steps)},
        )
        await status.edit_text(preview_text(batch))
        return True
    if not has_changes:
        repository.finish_without_file(
            operation_id,
            user_id,
            "no_changes_found",
            {"task_count": len(batch.steps)},
        )
        await status.edit_text(
            f"<b>No changes needed</b> · {len(batch.steps)} tasks\n"
            + "\n".join(
                f"{index}. {esc(step.preview.summary)}"
                for index, step in enumerate(batch.steps, start=1)
            )
        )
        return True

    await status.edit_text(
        proposal_text(batch),
        reply_markup=operation_keyboard(operation_id),
    )
    return True


def _runtime_source_context(active: ActiveSource | None) -> dict | None:
    if active is None:
        return None
    if isinstance(active.metadata, WorkbookMetadata):
        return {
            "name": active.source.display_name,
            "kind": "excel",
            "sheets": [
                {
                    "name": sheet.name,
                    "data_rows": sheet.data_rows,
                    "header_row": sheet.header_row,
                    "columns": [
                        {
                            "name": column.header,
                            "index": column.index,
                            "samples": column.samples[:3],
                        }
                        for column in sheet.columns[:120]
                    ],
                }
                for sheet in active.metadata.sheets[:20]
            ],
        }
    if isinstance(active.metadata, SqliteMetadata):
        return {
            "name": active.source.display_name,
            "kind": "sql",
            "tables": [
                {
                    "name": table.name,
                    "row_count": table.row_count,
                    "columns": [
                        {
                            "name": column.name,
                            "type": column.declared_type,
                            "samples": column.samples[:3],
                        }
                        for column in table.columns[:120]
                    ],
                }
                for table in active.metadata.tables[:30]
            ],
        }
    return {
        "name": active.source.display_name,
        "kind": active.source.kind.value,
    }


def _operation_context(record: OperationRecord | None) -> dict | None:
    if record is None:
        return None
    return {
        "operation_id": record.operation_id,
        "command_text": record.command_text,
        "action": record.action,
        "status": record.status,
        "result_file": Path(record.result_path).name if record.result_path else None,
    }


async def _command_runtime_context(
    user_id: int,
    active: ActiveSource | None = None,
) -> dict:
    if active is None:
        try:
            active = await get_active_source(user_id)
        except Exception:
            active = None
    pending = repository.get_latest_confirmable_operation(user_id)
    previous = repository.get_latest_completed_file_operation(user_id)
    sources = catalog.list_sources(user_id)
    return {
        "active_source": _runtime_source_context(active),
        "pending_operation": _operation_context(pending),
        "last_completed_operation": _operation_context(previous),
        "available_sources": [
            {
                "name": source.display_name,
                "kind": source.kind.value,
                "editable": source.editable,
            }
            for source in sources[:40]
        ],
    }


def _execution_from_envelope(envelope) -> ExecutionOptions:
    return ExecutionOptions(
        copy_original=envelope.execution.copy_original,
        output_name=envelope.execution.output_name,
    )


@router.message(F.text)
async def text_command(message: Message, bot: Bot) -> None:
    user_id = message.from_user.id if message.from_user else 0
    if not allowed(user_id):
        await message.answer("Access denied.")
        return
    if not await require_onboarding_message(message, user_id):
        return
    if not message.text or message.text.startswith("/"):
        await message.answer("Send a request in natural language.")
        return

    original_text = message.text.strip()

    # The LLM is the only natural-language interpreter. Python below only
    # dispatches a validated structured CommandEnvelope.
    active_before = await get_active_source(user_id)
    runtime_context = await _command_runtime_context(user_id, active_before)
    try:
        envelope = await intent_parser.parse_command(original_text, runtime_context)
    except Exception as error:
        await message.answer(
            "<b>Could not understand the command</b>\n"
            f"<code>{esc(type(error).__name__)}: {esc(error)}</code>"
        )
        return

    requested_execution = _execution_from_envelope(envelope)

    if envelope.mode == CommandMode.UNDO:
        status = await message.answer("↩️ Checking the latest change…")
        await execute_undo(user_id, status, bot)
        return

    if envelope.mode == CommandMode.SEND_LAST_FILE:
        await send_latest_processed_file(message, user_id)
        return

    if envelope.mode == CommandMode.RENAME_LAST_FILE:
        target_name = envelope.rename_to or envelope.execution.output_name
        if not target_name:
            await message.answer("Specify the new result filename.")
            return
        await rename_latest_processed_file(message, user_id, target_name, bot)
        return

    if envelope.mode == CommandMode.CONFIRM_PENDING:
        record = repository.get_latest_confirmable_operation(user_id)
        if record is None:
            await message.answer(
                "There is no operation awaiting confirmation. "
                "First describe what you want to change."
            )
            return
        try:
            operation = pending_from_record(record)
        except Exception as error:
            await message.answer(
                "The saved plan is no longer available: "
                f"<code>{esc(type(error).__name__)}: {esc(error)}</code>"
            )
            return
        status = await message.answer("Executing…")
        await execute_confirmed_operation(operation, status, bot)
        return

    if envelope.mode == CommandMode.CANCEL_PENDING:
        record = repository.get_latest_confirmable_operation(user_id)
        if record is None:
            await message.answer("There is no operation to cancel.")
            return
        repository.cancel_operation(
            record.operation_id,
            user_id,
            "The user cancelled the operation with a text message.",
        )
        await message.answer("<b>Cancelled</b>")
        return

    if envelope.mode == CommandMode.UPDATE_PENDING:
        record = repository.get_latest_confirmable_operation(user_id)
        if record is None:
            # If the model understood an output-name request but there is no
            # preview, treat it as a rename of the latest completed file.
            target_name = envelope.rename_to or envelope.execution.output_name
            if target_name:
                await rename_latest_processed_file(message, user_id, target_name, bot)
                return
            await message.answer(
                "There is no pending operation whose output options can be changed."
            )
            return
        try:
            pending = pending_from_record(record)
            merged = _merge_execution_options(pending.execution, requested_execution)
            pending = _with_execution_options(pending, merged)
            repository.save_plan(record.operation_id, user_id, pending_to_payload(pending))
            notes: list[str] = []
            if merged.copy_original:
                notes.append("The original will remain unchanged — the result will be created as a separate copy.")
            if merged.output_name:
                notes.append(f"Output filename: <code>{esc(merged.output_name)}</code>.")
            await message.answer(
                "<b>Result options updated</b>\n"
                + ("\n".join(notes) if notes else "Options saved.")
                + "\n\n"
                + proposal_text(pending),
                reply_markup=operation_keyboard(record.operation_id),
            )
        except Exception as error:
            await message.answer(
                "Could not update the pending operation: "
                f"<code>{esc(type(error).__name__)}: {esc(error)}</code>"
            )
        return

    repeat_mode = envelope.mode == CommandMode.REPEAT_LAST
    effective_request = original_text
    intents: list[IntentDraft]

    if repeat_mode:
        previous = repository.get_latest_completed_file_operation(user_id)
        if previous is None:
            await message.answer("There is no successful operation to repeat yet.")
            return

        active = active_before
        if envelope.target_source_name:
            matches = match_sources(
                envelope.target_source_name,
                catalog.list_sources(user_id),
            )
            if len(matches) != 1 or not matches[0].editable:
                await message.answer(
                    "Could not identify the new file unambiguously. Select it in the dashboard "
                    "or name it more precisely."
                )
                return
            active = await activate_source(
                user_id,
                matches[0],
                cancel_pending=False,
            )
        if active is None:
            await message.answer(
                "Select a new file in the dashboard or name it in the command first."
            )
            return
        if previous.result_path and active.source.path == Path(previous.result_path).resolve():
            await message.answer(
                "The same file is still selected. Choose a different file or specify its name."
            )
            return

        replay_context = await _command_runtime_context(user_id, active)
        replay_context["replaying_operation_id"] = previous.operation_id
        replay = await intent_parser.parse_command(previous.command_text, replay_context)
        if replay.mode != CommandMode.DATA or not replay.tasks:
            await message.answer("The previous operation no longer contains a repeatable data task.")
            return
        intents = [
            task.model_copy(update={"source_name_hint": None})
            for task in replay.tasks
            if task.action != Action.OPEN_SOURCE
        ]
        if not intents:
            await message.answer("The previous operation no longer contains a repeatable data task.")
            return
        effective_request = previous.command_text
        requested_execution = requested_execution.model_copy(
            update={"repeated_from_operation_id": previous.operation_id}
        )
    elif envelope.mode == CommandMode.DATA:
        intents = envelope.tasks
    else:
        await message.answer(
            envelope.clarification_question
            or "I could not determine the requested action. Please be a bit more specific."
        )
        return

    if not intents:
        await message.answer(
            envelope.clarification_question or "No data task was found in the message."
        )
        return

    repository.cancel_open_operations(user_id, "The user sent a new command.")
    operation_id = uuid4().hex[:12]
    repository.create_request(operation_id, user_id, effective_request)
    if repeat_mode and requested_execution.repeated_from_operation_id:
        repository.add_event(
            operation_id,
            user_id,
            "operation_repeated",
            {
                "from_operation_id": requested_execution.repeated_from_operation_id,
                "user_command": original_text,
            },
        )

    selected_record = repository.get_active_source(user_id)
    task_count = len(intents)
    status = await message.answer(
        (
            f"Building plan · {task_count} tasks…"
            if task_count > 1
            else (
                f"🎯 Working with <code>{esc(selected_record.original_name)}</code>. "
                "Building plan…"
                if selected_record
                else "Finding source and building plan…"
            )
        )
    )

    try:
        compact_tasks, selector = _batch_tasks_and_source_selector(intents)
        if selector is not None and len(compact_tasks) == 1:
            intents = compact_tasks

        if len(intents) > 1:
            repository.save_intent(
                operation_id,
                user_id,
                "batch",
                intents[0].source_type_hint.value,
                {"tasks": [item.model_dump(mode="json") for item in intents]},
            )
            if await _handle_multi_intent_command(
                status,
                user_id,
                operation_id,
                intents,
                effective_request,
                requested_execution,
            ):
                return

        intent = intents[0]
        repository.save_intent(
            operation_id,
            user_id,
            intent.action.value,
            intent.source_type_hint.value,
            intent.model_dump(mode="json"),
        )
        active, candidates = await _resolve_command_source(
            user_id,
            intent,
            effective_request,
            operation_id,
        )
        if active is None:
            reason = (
                "Multiple similar files were found."
                if len(candidates) > 1
                else "Source not found or its adapter is not available yet."
            )
            repository.stop_operation(
                operation_id,
                user_id,
                "needs_clarification",
                reason,
                "clarification_required",
            )
            await status.edit_text(
                f"<b>Select a file</b>\n{esc(reason)}",
                reply_markup=(
                    source_choice_keyboard(candidates)
                    if candidates
                    else main_menu_keyboard()
                ),
            )
            return
        if intent.action == Action.OPEN_SOURCE:
            repository.finish_without_file(
                operation_id,
                user_id,
                "read_completed",
                {"source_id": active.source.source_id, "action": "open_source"},
            )
            await status.edit_text(
                active_source_text(active),
                reply_markup=back_to_files_keyboard(),
            )
            return

        if active.source.kind == SourceKind.EXCEL:
            if intent.action not in SUPPORTED_EXCEL_ACTIONS:
                raise ValueError(f"Excel action {intent.action.value} is not implemented yet.")
            if not isinstance(active.metadata, WorkbookMetadata):
                raise TypeError("Excel metadata has an invalid type.")
            plan = await _plan_excel_intent(
                intent,
                intent.normalized_request,
                active.metadata,
            )
            if not plan.resolved or (
                plan.action != Action.REPLACE_ALL_VALUES and not plan.sheet_name
            ):
                reason = (
                    plan.clarification_question
                    or plan.resolution_note
                    or "No unambiguous match was found."
                )
                alternatives = "\n".join(
                    f"• {esc(item)}" for item in plan.alternative_matches
                )
                repository.stop_operation(
                    operation_id,
                    user_id,
                    "needs_clarification",
                    reason,
                    "clarification_required",
                )
                await status.edit_text(
                    "<b>Clarification required</b>\n"
                    + esc(reason)
                    + (f"\n\n{alternatives}" if alternatives else "")
                )
                return
            preview = await asyncio.to_thread(
                build_operation_preview,
                active.source.path,
                plan,
                active.metadata,
            )
            operation: PendingOperation = PendingExcelOperation(
                operation_id=operation_id,
                user_id=user_id,
                source=active.source,
                metadata=active.metadata,
                plan=plan,
                preview=preview,
                created_at=datetime.now(timezone.utc),
                execution=requested_execution,
            )
            is_read = plan.action == Action.SELECT
            is_write = plan.action in WRITE_ACTIONS
        elif active.source.kind == SourceKind.SQLITE:
            if intent.action not in SUPPORTED_SQL_ACTIONS:
                raise ValueError(
                    "SQLite currently supports select/update_rows/delete_rows, "
                    f"but the interpreted action is {intent.action.value}."
                )
            if not isinstance(active.metadata, SqliteMetadata):
                raise TypeError("SQLite metadata has an invalid type.")
            plan = await sql_planner.resolve(
                intent.normalized_request,
                intent,
                active.metadata,
            )
            if not plan.resolved or not plan.table_name:
                reason = (
                    plan.clarification_question
                    or plan.resolution_note
                    or "No unambiguous match was found."
                )
                repository.stop_operation(
                    operation_id,
                    user_id,
                    "needs_clarification",
                    reason,
                    "clarification_required",
                )
                await status.edit_text("<b>Clarification required</b>\n" + esc(reason))
                return
            preview = await asyncio.to_thread(
                build_sql_preview,
                active.source.path,
                plan,
                active.metadata,
            )
            operation = PendingSqlOperation(
                operation_id=operation_id,
                user_id=user_id,
                source=active.source,
                metadata=active.metadata,
                plan=plan,
                preview=preview,
                created_at=datetime.now(timezone.utc),
                execution=requested_execution,
            )
            is_read = plan.action == Action.SELECT
            is_write = plan.action in SQL_WRITE_ACTIONS
        else:
            raise ValueError("The selected format has no execution engine yet.")

        repository.save_plan(operation_id, user_id, pending_to_payload(operation))
        if is_read:
            repository.finish_without_file(
                operation_id,
                user_id,
                "read_completed",
                {"matched_rows": preview.matched_rows},
            )
            await status.edit_text(preview_text(operation))
            return
        if is_write and not preview.has_changes:
            repository.finish_without_file(
                operation_id,
                user_id,
                "no_changes_found",
                {"summary": preview.summary},
            )
            await status.edit_text("<b>No changes needed</b>\n" + esc(preview.summary))
            return
        await status.edit_text(
            proposal_text(operation),
            reply_markup=operation_keyboard(operation_id),
        )
    except Exception as error:
        error_message = f"{type(error).__name__}: {error}"
        repository.stop_operation(
            operation_id,
            user_id,
            "failed",
            error_message,
            "planning_failed",
        )
        await status.edit_text(planning_failure_text())


async def main() -> None:
    catalog.initialize()
    repository.initialize()
    bot = Bot(
        token=settings.telegram_bot_token.get_secret_value(),
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
    )
    try:
        for user_id in settings.allowed_user_ids:
            if not onboarding_accepted(user_id):
                continue
            try:
                await ensure_files_dashboard(bot, user_id, user_id)
            except Exception as error:
                print(
                    "Could not restore the file dashboard for "
                    f"Telegram ID {user_id}: {type(error).__name__}: {error}"
                )
        dispatcher = Dispatcher()
        dispatcher.include_router(router)
        await bot.delete_webhook(drop_pending_updates=True)
        print(
            "DataOps Commander started. "
            f"Ollama: {settings.ollama_model}. "
            f"Workspace: {catalog.workspace_root}. "
            f"Audit DB: {DATABASE_PATH.resolve()}"
        )
        await dispatcher.start_polling(
            bot,
            allowed_updates=dispatcher.resolve_used_update_types(),
        )
    finally:
        await bot.session.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("Bot stopped")
