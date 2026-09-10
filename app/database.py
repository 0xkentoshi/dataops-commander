import json
import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


@dataclass(frozen=True, slots=True)
class ActiveSourceRecord:
    user_id: int
    file_path: str
    original_name: str
    metadata: dict[str, Any]
    updated_at: datetime


@dataclass(frozen=True, slots=True)
class DashboardRecord:
    user_id: int
    chat_id: int
    message_id: int
    page: int
    updated_at: datetime


@dataclass(frozen=True, slots=True)
class OperationRecord:
    operation_id: str
    user_id: int
    command_text: str
    action: str
    source_type: str
    status: str
    intent: dict[str, Any]
    plan: dict[str, Any]
    created_at: datetime
    updated_at: datetime
    confirmed_at: datetime | None
    completed_at: datetime | None
    snapshot_path: str | None
    result_path: str | None
    result_sha256: str | None
    error_message: str | None


@dataclass(frozen=True, slots=True)
class AuditEventRecord:
    event_id: int
    operation_id: str | None
    user_id: int
    event_type: str
    details: dict[str, Any]
    created_at: datetime


_UNSET = object()


def _now() -> datetime:
    return datetime.now(timezone.utc)


def _iso(value: datetime) -> str:
    return value.astimezone(timezone.utc).isoformat()


def _datetime(value: str | None) -> datetime | None:
    if value is None:
        return None
    result = datetime.fromisoformat(value)
    if result.tzinfo is None:
        result = result.replace(tzinfo=timezone.utc)
    return result.astimezone(timezone.utc)


def _json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, default=str)


def _dict(value: str | None) -> dict[str, Any]:
    loaded = json.loads(value or "{}")
    return loaded if isinstance(loaded, dict) else {}


class DataRepository:
    def __init__(self, database_path: Path) -> None:
        self.database_path = database_path

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(self.database_path, timeout=10)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA busy_timeout = 5000")
        try:
            with connection:
                yield connection
        finally:
            connection.close()

    def initialize(self) -> None:
        with self._connect() as connection:
            connection.execute("PRAGMA journal_mode = WAL")
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS active_sources (
                    user_id INTEGER PRIMARY KEY,
                    file_path TEXT NOT NULL,
                    original_name TEXT NOT NULL,
                    metadata_json TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS file_dashboards (
                    user_id INTEGER PRIMARY KEY,
                    chat_id INTEGER NOT NULL,
                    message_id INTEGER NOT NULL,
                    page INTEGER NOT NULL DEFAULT 0,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS user_settings (
                    user_id INTEGER PRIMARY KEY,
                    onboarding_version INTEGER NOT NULL DEFAULT 0,
                    onboarding_accepted_at TEXT
                );

                CREATE TABLE IF NOT EXISTS operations (
                    operation_id TEXT PRIMARY KEY,
                    user_id INTEGER NOT NULL,
                    command_text TEXT NOT NULL,
                    action TEXT NOT NULL DEFAULT 'unknown',
                    source_type TEXT NOT NULL DEFAULT 'unknown',
                    status TEXT NOT NULL,
                    intent_json TEXT NOT NULL DEFAULT '{}',
                    plan_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    confirmed_at TEXT,
                    completed_at TEXT,
                    snapshot_path TEXT,
                    result_path TEXT,
                    result_sha256 TEXT,
                    error_message TEXT
                );

                CREATE INDEX IF NOT EXISTS idx_operations_user_created
                    ON operations(user_id, created_at DESC);

                CREATE TABLE IF NOT EXISTS audit_events (
                    event_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    operation_id TEXT,
                    user_id INTEGER NOT NULL,
                    event_type TEXT NOT NULL,
                    details_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(operation_id)
                        REFERENCES operations(operation_id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_audit_operation
                    ON audit_events(operation_id, event_id);
                """
            )
            operation_columns = {
                row["name"] for row in connection.execute("PRAGMA table_info(operations)").fetchall()
            }
            if "result_sha256" not in operation_columns:
                connection.execute("ALTER TABLE operations ADD COLUMN result_sha256 TEXT")
            self._recover_interrupted(connection)

    def get_onboarding_version(self, user_id: int) -> int:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT onboarding_version FROM user_settings WHERE user_id = ?",
                (user_id,),
            ).fetchone()
        if row is None:
            return 0
        return int(row["onboarding_version"] or 0)

    def accept_onboarding(self, user_id: int, version: int) -> None:
        moment = _now()
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO user_settings (
                    user_id, onboarding_version, onboarding_accepted_at
                ) VALUES (?, ?, ?)
                ON CONFLICT(user_id) DO UPDATE SET
                    onboarding_version = excluded.onboarding_version,
                    onboarding_accepted_at = excluded.onboarding_accepted_at
                """,
                (user_id, max(int(version), 0), _iso(moment)),
            )

    def healthcheck(self) -> None:
        with self._connect() as connection:
            row = connection.execute("SELECT 1").fetchone()
        if row is None or row[0] != 1:
            raise RuntimeError("SQLite did not respond to the health-check query.")

    def _event(
        self,
        connection: sqlite3.Connection,
        operation_id: str | None,
        user_id: int,
        event_type: str,
        details: dict[str, Any] | None = None,
        moment: datetime | None = None,
    ) -> None:
        connection.execute(
            """
            INSERT INTO audit_events (
                operation_id, user_id, event_type, details_json, created_at
            ) VALUES (?, ?, ?, ?, ?)
            """,
            (
                operation_id,
                user_id,
                event_type,
                _json(details or {}),
                _iso(moment or _now()),
            ),
        )

    def _recover_interrupted(self, connection: sqlite3.Connection) -> None:
        rows = connection.execute(
            "SELECT operation_id, user_id FROM operations WHERE status = 'executing'"
        ).fetchall()
        if not rows:
            return

        moment = _now()
        reason = (
            "The bot restarted during execution. Re-read the source before retrying because "
            "the write may have completed before shutdown."
        )
        for row in rows:
            connection.execute(
                """
                UPDATE operations
                SET status = 'failed', updated_at = ?, error_message = ?
                WHERE operation_id = ?
                """,
                (_iso(moment), reason, row["operation_id"]),
            )
            self._event(
                connection,
                row["operation_id"],
                row["user_id"],
                "execution_interrupted",
                {"reason": reason},
                moment,
            )

    def _change(
        self,
        operation_id: str,
        user_id: int,
        *,
        status: str | object = _UNSET,
        action: str | object = _UNSET,
        source_type: str | object = _UNSET,
        intent: dict[str, Any] | object = _UNSET,
        plan: dict[str, Any] | object = _UNSET,
        error: str | None | object = _UNSET,
        snapshot_path: Path | object = _UNSET,
        result_path: Path | object = _UNSET,
        result_sha256: str | None | object = _UNSET,
        confirmed: bool = False,
        completed: bool = False,
        clear_completed: bool = False,
        allowed_statuses: tuple[str, ...] = (),
        event_type: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> bool:
        moment = _now()
        fields = ["updated_at = ?"]
        values: list[Any] = [_iso(moment)]

        for column, value in (
            ("status", status),
            ("action", action),
            ("source_type", source_type),
        ):
            if value is not _UNSET:
                fields.append(f"{column} = ?")
                values.append(value)
        if intent is not _UNSET:
            fields.append("intent_json = ?")
            values.append(_json(intent))
        if plan is not _UNSET:
            fields.append("plan_json = ?")
            values.append(_json(plan))
        if error is not _UNSET:
            fields.append("error_message = ?")
            values.append(error)
        if snapshot_path is not _UNSET:
            fields.append("snapshot_path = ?")
            values.append(str(Path(snapshot_path).resolve()))
        if result_path is not _UNSET:
            fields.append("result_path = ?")
            values.append(str(Path(result_path).resolve()))
        if result_sha256 is not _UNSET:
            fields.append("result_sha256 = ?")
            values.append(result_sha256)
        if confirmed:
            fields.append("confirmed_at = ?")
            values.append(_iso(moment))
        if completed:
            fields.append("completed_at = ?")
            values.append(_iso(moment))
        elif clear_completed:
            fields.append("completed_at = NULL")

        sql = (
            f"UPDATE operations SET {', '.join(fields)} "
            "WHERE operation_id = ? AND user_id = ?"
        )
        values.extend((operation_id, user_id))
        if allowed_statuses:
            placeholders = ",".join("?" for _ in allowed_statuses)
            sql += f" AND status IN ({placeholders})"
            values.extend(allowed_statuses)

        with self._connect() as connection:
            cursor = connection.execute(sql, values)
            changed = cursor.rowcount == 1
            if changed and event_type:
                self._event(
                    connection,
                    operation_id,
                    user_id,
                    event_type,
                    details,
                    moment,
                )
        return changed

    def set_active_source(
        self,
        user_id: int,
        file_path: Path,
        original_name: str,
        metadata: dict[str, Any],
        operation_id: str | None = None,
    ) -> None:
        moment = _now()
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO active_sources (
                    user_id, file_path, original_name, metadata_json, updated_at
                ) VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(user_id) DO UPDATE SET
                    file_path = excluded.file_path,
                    original_name = excluded.original_name,
                    metadata_json = excluded.metadata_json,
                    updated_at = excluded.updated_at
                """,
                (
                    user_id,
                    str(file_path.resolve()),
                    original_name,
                    _json(metadata),
                    _iso(moment),
                ),
            )
            self._event(
                connection,
                operation_id,
                user_id,
                "active_source_selected",
                {"original_name": original_name},
                moment,
            )

    def get_active_source(self, user_id: int) -> ActiveSourceRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM active_sources WHERE user_id = ?",
                (user_id,),
            ).fetchone()
        if row is None:
            return None
        updated_at = _datetime(row["updated_at"])
        if updated_at is None:
            raise ValueError("The source record has no updated_at value.")
        return ActiveSourceRecord(
            user_id=row["user_id"],
            file_path=row["file_path"],
            original_name=row["original_name"],
            metadata=_dict(row["metadata_json"]),
            updated_at=updated_at,
        )

    def clear_active_source(self, user_id: int, reason: str) -> None:
        with self._connect() as connection:
            connection.execute(
                "DELETE FROM active_sources WHERE user_id = ?",
                (user_id,),
            )
            self._event(
                connection,
                None,
                user_id,
                "active_source_cleared",
                {"reason": reason},
            )

    def set_dashboard(
        self,
        user_id: int,
        chat_id: int,
        message_id: int,
        page: int,
    ) -> None:
        moment = _now()
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO file_dashboards (
                    user_id, chat_id, message_id, page, updated_at
                ) VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(user_id) DO UPDATE SET
                    chat_id = excluded.chat_id,
                    message_id = excluded.message_id,
                    page = excluded.page,
                    updated_at = excluded.updated_at
                """,
                (user_id, chat_id, message_id, max(page, 0), _iso(moment)),
            )

    def get_dashboard(self, user_id: int) -> DashboardRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM file_dashboards WHERE user_id = ?",
                (user_id,),
            ).fetchone()
        if row is None:
            return None
        updated_at = _datetime(row["updated_at"])
        if updated_at is None:
            raise ValueError("The file-dashboard record has no updated_at value.")
        return DashboardRecord(
            user_id=row["user_id"],
            chat_id=row["chat_id"],
            message_id=row["message_id"],
            page=row["page"],
            updated_at=updated_at,
        )

    def clear_dashboard(self, user_id: int) -> None:
        with self._connect() as connection:
            connection.execute(
                "DELETE FROM file_dashboards WHERE user_id = ?",
                (user_id,),
            )

    def create_request(
        self,
        operation_id: str,
        user_id: int,
        command_text: str,
    ) -> None:
        moment = _now()
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO operations (
                    operation_id, user_id, command_text,
                    status, created_at, updated_at
                ) VALUES (?, ?, ?, 'received', ?, ?)
                """,
                (
                    operation_id,
                    user_id,
                    command_text,
                    _iso(moment),
                    _iso(moment),
                ),
            )
            self._event(
                connection,
                operation_id,
                user_id,
                "request_received",
                {"command_text": command_text},
                moment,
            )

    def save_intent(
        self,
        operation_id: str,
        user_id: int,
        action: str,
        source_type: str,
        intent: dict[str, Any],
    ) -> None:
        self._change(
            operation_id,
            user_id,
            action=action,
            source_type=source_type,
            intent=intent,
            event_type="intent_parsed",
            details={"action": action, "source_type": source_type},
        )

    def save_plan(
        self,
        operation_id: str,
        user_id: int,
        plan: dict[str, Any],
    ) -> None:
        self._change(
            operation_id,
            user_id,
            status="planned",
            plan=plan,
            error=None,
            clear_completed=True,
            event_type="plan_created",
            details={
                "sheet_name": plan.get("sheet_name"),
                "table_name": plan.get("table_name"),
                "column_header": plan.get("column_header"),
            },
        )

    def stop_operation(
        self,
        operation_id: str,
        user_id: int,
        status: str,
        reason: str,
        event_type: str,
    ) -> None:
        self._change(
            operation_id,
            user_id,
            status=status,
            error=reason,
            completed=True,
            event_type=event_type,
            details={"reason": reason},
        )

    def mark_previewed(self, operation_id: str, user_id: int) -> None:
        changed = self._change(
            operation_id,
            user_id,
            status="previewed",
            allowed_statuses=("planned",),
            event_type="preview_shown",
        )
        if not changed:
            self.add_event(operation_id, user_id, "preview_shown")

    def claim_operation(self, operation_id: str, user_id: int) -> bool:
        return self._change(
            operation_id,
            user_id,
            status="executing",
            error=None,
            confirmed=True,
            clear_completed=True,
            allowed_statuses=("planned", "previewed", "failed"),
            event_type="confirmation_received",
        )

    def cancel_operation(
        self,
        operation_id: str,
        user_id: int,
        reason: str,
    ) -> bool:
        return self._change(
            operation_id,
            user_id,
            status="cancelled",
            error=reason,
            completed=True,
            allowed_statuses=("received", "planned", "previewed", "failed"),
            event_type="operation_cancelled",
            details={"reason": reason},
        )

    def expire_operation(self, operation_id: str, user_id: int) -> bool:
        reason = "The 15-minute confirmation window expired."
        return self._change(
            operation_id,
            user_id,
            status="expired",
            error=reason,
            completed=True,
            allowed_statuses=("planned", "previewed", "failed"),
            event_type="confirmation_expired",
            details={"reason": reason},
        )

    def cancel_open_operations(self, user_id: int, reason: str) -> int:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT operation_id FROM operations
                WHERE user_id = ? AND (
                    status IN ('received', 'planned', 'previewed')
                    OR (status = 'failed' AND plan_json <> '{}')
                )
                """,
                (user_id,),
            ).fetchall()
            if not rows:
                return 0

            moment = _now()
            for row in rows:
                connection.execute(
                    """
                    UPDATE operations
                    SET status = 'cancelled', updated_at = ?,
                        completed_at = ?, error_message = ?
                    WHERE operation_id = ?
                    """,
                    (
                        _iso(moment),
                        _iso(moment),
                        reason,
                        row["operation_id"],
                    ),
                )
                self._event(
                    connection,
                    row["operation_id"],
                    user_id,
                    "operation_replaced",
                    {"reason": reason},
                    moment,
                )
            return len(rows)

    def add_event(
        self,
        operation_id: str,
        user_id: int,
        event_type: str,
        details: dict[str, Any] | None = None,
    ) -> None:
        with self._connect() as connection:
            self._event(
                connection,
                operation_id,
                user_id,
                event_type,
                details,
            )

    def fail_operation(
        self,
        operation_id: str,
        user_id: int,
        error_message: str,
    ) -> None:
        self._change(
            operation_id,
            user_id,
            status="failed",
            error=error_message,
            event_type="execution_failed",
            details={"error": error_message},
        )

    def complete_operation(
        self,
        operation_id: str,
        user_id: int,
        snapshot_path: Path,
        result_path: Path,
        result_sha256: str | None = None,
    ) -> None:
        self._change(
            operation_id,
            user_id,
            status="completed",
            error=None,
            snapshot_path=snapshot_path,
            result_path=result_path,
            result_sha256=result_sha256,
            completed=True,
            event_type="operation_completed",
            details={
                "snapshot_name": snapshot_path.name,
                "result_name": result_path.name,
                "result_sha256": result_sha256,
            },
        )

    def mark_undone(
        self,
        operation_id: str,
        user_id: int,
        undo_operation_id: str,
    ) -> bool:
        return self._change(
            operation_id,
            user_id,
            status="undone",
            allowed_statuses=("completed",),
            event_type="operation_undone",
            details={"undo_operation_id": undo_operation_id},
        )

    def finish_without_file(
        self,
        operation_id: str,
        user_id: int,
        event_type: str,
        details: dict[str, Any] | None = None,
    ) -> None:
        self._change(
            operation_id=operation_id,
            user_id=user_id,
            status="completed",
            error=None,
            completed=True,
            event_type=event_type,
            details=details,
        )

    @staticmethod
    def _operation(row: sqlite3.Row) -> OperationRecord:
        created_at = _datetime(row["created_at"])
        updated_at = _datetime(row["updated_at"])
        if created_at is None or updated_at is None:
            raise ValueError("The operation has no timestamps.")
        return OperationRecord(
            operation_id=row["operation_id"],
            user_id=row["user_id"],
            command_text=row["command_text"],
            action=row["action"],
            source_type=row["source_type"],
            status=row["status"],
            intent=_dict(row["intent_json"]),
            plan=_dict(row["plan_json"]),
            created_at=created_at,
            updated_at=updated_at,
            confirmed_at=_datetime(row["confirmed_at"]),
            completed_at=_datetime(row["completed_at"]),
            snapshot_path=row["snapshot_path"],
            result_path=row["result_path"],
            result_sha256=row["result_sha256"],
            error_message=row["error_message"],
        )

    def get_operation(self, operation_id: str) -> OperationRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        return None if row is None else self._operation(row)

    def get_latest_operation(self, user_id: int) -> OperationRecord | None:
        records = self.list_operations(user_id, limit=1)
        return records[0] if records else None

    def get_latest_undoable_operation(
        self,
        user_id: int,
        result_path: Path | str,
    ) -> OperationRecord | None:
        resolved = str(Path(result_path).resolve())
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT * FROM operations
                WHERE user_id = ?
                  AND status = 'completed'
                  AND action <> 'undo'
                  AND snapshot_path IS NOT NULL
                  AND result_path = ?
                ORDER BY completed_at DESC, created_at DESC
                LIMIT 1
                """,
                (user_id, resolved),
            ).fetchone()
        return None if row is None else self._operation(row)

    def get_latest_completed_file_operation(
        self,
        user_id: int,
    ) -> OperationRecord | None:
        """Latest successful operation that produced/updated a real file."""
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT * FROM operations
                WHERE user_id = ?
                  AND status = 'completed'
                  AND action <> 'undo'
                  AND result_path IS NOT NULL
                ORDER BY completed_at DESC, created_at DESC
                LIMIT 1
                """,
                (user_id,),
            ).fetchone()
        return None if row is None else self._operation(row)

    def relocate_file_path(
        self,
        user_id: int,
        old_path: Path | str,
        new_path: Path | str,
        operation_id: str | None = None,
    ) -> None:
        """Keep active-source and undo history valid after a user-requested rename."""
        old_resolved = str(Path(old_path).resolve())
        new_resolved = str(Path(new_path).resolve())
        moment = _now()
        with self._connect() as connection:
            connection.execute(
                """
                UPDATE operations
                SET result_path = ?, updated_at = ?
                WHERE user_id = ? AND result_path = ?
                """,
                (new_resolved, _iso(moment), user_id, old_resolved),
            )
            connection.execute(
                """
                UPDATE active_sources
                SET file_path = ?, original_name = ?, updated_at = ?
                WHERE user_id = ? AND file_path = ?
                """,
                (new_resolved, Path(new_resolved).name, _iso(moment), user_id, old_resolved),
            )
            self._event(
                connection,
                operation_id,
                user_id,
                "result_file_renamed",
                {"old_path": old_resolved, "new_path": new_resolved},
                moment,
            )

    def get_latest_confirmable_operation(
        self,
        user_id: int,
    ) -> OperationRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT * FROM operations
                WHERE user_id = ?
                ORDER BY updated_at DESC, created_at DESC
                LIMIT 1
                """,
                (user_id,),
            ).fetchone()
        if (
            row is None
            or row["status"] not in {"planned", "previewed", "failed"}
            or row["plan_json"] == "{}"
        ):
            return None
        return self._operation(row)

    def list_operations(
        self,
        user_id: int,
        limit: int = 10,
    ) -> list[OperationRecord]:
        safe_limit = min(max(limit, 1), 50)
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT * FROM operations
                WHERE user_id = ?
                ORDER BY created_at DESC
                LIMIT ?
                """,
                (user_id, safe_limit),
            ).fetchall()
        return [self._operation(row) for row in rows]

    def list_events(
        self,
        operation_id: str,
        user_id: int,
    ) -> list[AuditEventRecord]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT * FROM audit_events
                WHERE operation_id = ? AND user_id = ?
                ORDER BY event_id
                """,
                (operation_id, user_id),
            ).fetchall()

        result: list[AuditEventRecord] = []
        for row in rows:
            created_at = _datetime(row["created_at"])
            if created_at is not None:
                result.append(
                    AuditEventRecord(
                        event_id=row["event_id"],
                        operation_id=row["operation_id"],
                        user_id=row["user_id"],
                        event_type=row["event_type"],
                        details=_dict(row["details_json"]),
                        created_at=created_at,
                    )
                )
        return result
