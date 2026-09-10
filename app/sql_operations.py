from __future__ import annotations

import hashlib
import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from pydantic import BaseModel, ConfigDict, Field

from app.excel_planner import FilterOperator
from app.intent import Action
from app.sql_planner import SqlColumnRef, SqlFilterCondition, SqlOperationPlan
from app.sql_service import SqlTableMetadata, SqliteMetadata, open_read_only, quote_identifier


MAX_PREVIEW_ROWS = 10
MAX_PREVIEW_COLUMNS = 8
MAX_MUTATED_ROWS = 5_000
SQL_WRITE_ACTIONS = {Action.UPDATE_ROWS, Action.DELETE_ROWS}


class SqlPreviewCell(BaseModel):
    model_config = ConfigDict(extra="forbid")

    column_name: str
    before: str
    after: str | None = None


class SqlPreviewRow(BaseModel):
    model_config = ConfigDict(extra="forbid")

    rowid: int
    cells: list[SqlPreviewCell] = Field(default_factory=list)


class SqlOperationPreview(BaseModel):
    model_config = ConfigDict(extra="forbid")

    action: Action
    table_name: str
    is_write: bool
    has_changes: bool
    matched_rows: int
    affected_cells: int
    target_rowids: list[int] = Field(default_factory=list)
    rows: list[SqlPreviewRow] = Field(default_factory=list)
    summary: str
    target_signature: str


@dataclass(frozen=True, slots=True)
class SqlExecutionOutcome:
    snapshot_path: Path
    database_path: Path
    changed_rows: int
    affected_cells: int
    transaction_committed: bool
    result_verified: bool


def _table(metadata: SqliteMetadata, name: str) -> SqlTableMetadata:
    result = next((item for item in metadata.tables if item.name == name), None)
    if result is None:
        raise ValueError(f"Table “{name}” is not present in the schema.")
    return result


def _display(value: Any) -> str:
    if value is None:
        return "<NULL>"
    if isinstance(value, bytes):
        return f"<BLOB {len(value)} bytes>"
    text = str(value).replace("\n", " ").strip()
    return text if len(text) <= 100 else f"{text[:99]}…"


def _safe(value: Any) -> Any:
    if isinstance(value, bytes):
        return {"blob_sha256": hashlib.sha256(value).hexdigest(), "bytes": len(value)}
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _like_value(value: object) -> str:
    text = str(value).replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    return f"%{text}%"


def _where(filters: list[SqlFilterCondition]) -> tuple[str, list[Any]]:
    clauses: list[str] = []
    parameters: list[Any] = []
    for condition in filters:
        column = quote_identifier(condition.column.name)
        operator = condition.operator
        if operator == FilterOperator.IS_EMPTY:
            clauses.append(f"({column} IS NULL OR TRIM(CAST({column} AS TEXT)) = '')")
        elif operator == FilterOperator.NOT_EMPTY:
            clauses.append(f"({column} IS NOT NULL AND TRIM(CAST({column} AS TEXT)) <> '')")
        elif operator == FilterOperator.EQ:
            clauses.append(f"{column} = ?")
            parameters.append(condition.value)
        elif operator == FilterOperator.NE:
            clauses.append(f"{column} <> ?")
            parameters.append(condition.value)
        elif operator == FilterOperator.CONTAINS:
            clauses.append(f"CAST({column} AS TEXT) LIKE ? ESCAPE '\\'")
            parameters.append(_like_value(condition.value))
        elif operator == FilterOperator.NOT_CONTAINS:
            clauses.append(f"CAST({column} AS TEXT) NOT LIKE ? ESCAPE '\\'")
            parameters.append(_like_value(condition.value))
        elif operator in {
            FilterOperator.GT,
            FilterOperator.GTE,
            FilterOperator.LT,
            FilterOperator.LTE,
        }:
            signs = {
                FilterOperator.GT: ">",
                FilterOperator.GTE: ">=",
                FilterOperator.LT: "<",
                FilterOperator.LTE: "<=",
            }
            clauses.append(f"{column} {signs[operator]} ?")
            parameters.append(condition.value)
        elif operator == FilterOperator.OLDER_THAN_DAYS:
            clauses.append(f"date({column}) < date('now', ?)")
            parameters.append(f"-{int(float(condition.value))} days")
        else:
            raise ValueError(f"SQL filter {operator.value} is not supported.")
    return (" AND ".join(clauses) if clauses else "1 = 1"), parameters


def _preview_refs(plan: SqlOperationPlan, table: SqlTableMetadata) -> list[SqlColumnRef]:
    if plan.action == Action.SELECT:
        refs = plan.selected_columns or [
            SqlColumnRef(name=item.name, declared_type=item.declared_type)
            for item in table.columns
        ]
    elif plan.action == Action.UPDATE_ROWS:
        refs = [
            *(item.column for item in plan.filters),
            *(item.column for item in plan.assignments),
        ]
    else:
        refs = [
            SqlColumnRef(name=item.name, declared_type=item.declared_type)
            for item in table.columns
        ]
    unique: dict[str, SqlColumnRef] = {}
    for ref in refs:
        unique.setdefault(ref.name, ref)
    return list(unique.values())[:MAX_PREVIEW_COLUMNS]


def _collect(
    connection: sqlite3.Connection,
    plan: SqlOperationPlan,
    metadata: SqliteMetadata,
) -> SqlOperationPreview:
    if not plan.resolved or not plan.table_name:
        raise ValueError("Cannot build an SQL preview for an unresolved plan.")
    table = _table(metadata, plan.table_name)
    if not table.supports_rowid or not table.rowid_alias:
        raise ValueError("WITHOUT ROWID tables are currently available for schema inspection only.")
    if plan.action in SQL_WRITE_ACTIONS and table.trigger_count:
        raise ValueError("Write blocked: the table has SQL triggers.")
    if plan.action in SQL_WRITE_ACTIONS and (
        table.foreign_key_count or table.referenced_by_foreign_keys
    ):
        raise ValueError("Write blocked: the table has foreign keys.")
    if plan.action == Action.UPDATE_ROWS:
        primary_keys = {
            column.name for column in table.columns if column.primary_key_position > 0
        }
        if any(item.column.name in primary_keys for item in plan.assignments):
            raise ValueError("PRIMARY KEY changes are blocked by the safety guard.")

    quoted_table = quote_identifier(table.name)
    where_sql, parameters = _where(plan.filters)
    matched_rows = int(
        connection.execute(
            f"SELECT COUNT(*) FROM {quoted_table} WHERE {where_sql}", parameters
        ).fetchone()[0]
    )
    if plan.action in SQL_WRITE_ACTIONS and matched_rows > MAX_MUTATED_ROWS:
        raise ValueError(
            f"The operation affects {matched_rows} rows; the limit is {MAX_MUTATED_ROWS}."
        )

    all_names = [item.name for item in table.columns]
    all_columns_sql = ", ".join(quote_identifier(name) for name in all_names)
    limit = MAX_MUTATED_ROWS + 1 if plan.action in SQL_WRITE_ACTIONS else MAX_PREVIEW_ROWS
    target_rows = connection.execute(
        f"SELECT {table.rowid_alias}, {all_columns_sql} FROM {quoted_table} "
        f"WHERE {where_sql} ORDER BY {table.rowid_alias} LIMIT ?",
        [*parameters, limit],
    ).fetchall()
    target_rowids = [int(row[0]) for row in target_rows]
    signature_payload = [
        [int(row[0]), *(_safe(value) for value in row[1:])] for row in target_rows
    ]
    signature = hashlib.sha256(
        json.dumps(signature_payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()

    refs = _preview_refs(plan, table)
    positions = {name: index + 1 for index, name in enumerate(all_names)}
    assignments = {item.column.name: item.value for item in plan.assignments}
    preview_rows: list[SqlPreviewRow] = []
    for row in target_rows[:MAX_PREVIEW_ROWS]:
        cells: list[SqlPreviewCell] = []
        for ref in refs:
            before = row[positions[ref.name]]
            after = assignments.get(ref.name, ...)
            cells.append(
                SqlPreviewCell(
                    column_name=ref.name,
                    before=_display(before),
                    after=None if after is ... else _display(after),
                )
            )
        preview_rows.append(SqlPreviewRow(rowid=int(row[0]), cells=cells))

    if plan.action == Action.SELECT:
        affected_cells = 0
        summary = f"Rows found: {matched_rows}. The database will not be modified."
    elif plan.action == Action.UPDATE_ROWS:
        affected_cells = sum(
            not _equivalent(
                row[positions[assignment.column.name]],
                assignment.value,
            )
            for row in target_rows
            for assignment in plan.assignments
        )
        summary = f"Rows to update: {matched_rows}."
    elif plan.action == Action.DELETE_ROWS:
        affected_cells = matched_rows * max(len(table.columns), 1)
        summary = f"Rows to delete: {matched_rows}."
    else:
        raise ValueError(f"SQL action {plan.action.value} is not supported.")
    return SqlOperationPreview(
        action=plan.action,
        table_name=table.name,
        is_write=plan.action in SQL_WRITE_ACTIONS,
        has_changes=plan.action in SQL_WRITE_ACTIONS and affected_cells > 0,
        matched_rows=matched_rows,
        affected_cells=affected_cells,
        target_rowids=target_rowids if plan.action in SQL_WRITE_ACTIONS else [],
        rows=preview_rows,
        summary=summary,
        target_signature=signature,
    )


def build_sql_preview(
    database_path: Path,
    plan: SqlOperationPlan,
    metadata: SqliteMetadata,
) -> SqlOperationPreview:
    connection = open_read_only(database_path)
    try:
        return _collect(connection, plan, metadata)
    finally:
        connection.close()


def _snapshot_path(database_path: Path, operation_id: str, root: Path) -> Path:
    root.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    safe_stem = "".join(
        character if character.isalnum() or character in "._-" else "_"
        for character in database_path.stem
    )
    return root / f"{safe_stem}_{stamp}_{operation_id}_snapshot.sqlite3"


def _backup_database(source_path: Path, snapshot_path: Path) -> None:
    source = sqlite3.connect(source_path, timeout=10)
    destination = sqlite3.connect(snapshot_path)
    try:
        source.backup(destination)
        check = destination.execute("PRAGMA integrity_check").fetchone()
        if check is None or check[0] != "ok":
            raise RuntimeError(f"The SQLite snapshot is corrupted: {check}")
    finally:
        destination.close()
        source.close()


def _chunks(values: list[int], size: int = 800) -> list[list[int]]:
    return [values[index : index + size] for index in range(0, len(values), size)]


def _equivalent(actual: Any, expected: Any) -> bool:
    if actual is None or expected is None:
        return actual is expected
    if isinstance(actual, (int, float)) and not isinstance(actual, bool):
        try:
            return float(actual) == float(expected)
        except (TypeError, ValueError):
            return False
    return str(actual) == str(expected)


def execute_sql_operation(
    database_path: Path,
    operation_id: str,
    plan: SqlOperationPlan,
    preview: SqlOperationPreview,
    metadata: SqliteMetadata,
    snapshots_root: Path = Path("data/snapshots/sqlite"),
) -> SqlExecutionOutcome:
    if plan.action not in SQL_WRITE_ACTIONS:
        raise ValueError("Read-only SQL does not require a write transaction.")
    snapshot_path = _snapshot_path(database_path, operation_id, snapshots_root)
    _backup_database(database_path, snapshot_path)
    table = _table(metadata, plan.table_name or "")
    quoted_table = quote_identifier(table.name)
    connection = sqlite3.connect(database_path, timeout=10, isolation_level=None)
    connection.row_factory = sqlite3.Row
    committed = False
    try:
        connection.execute("PRAGMA busy_timeout = 5000")
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("BEGIN IMMEDIATE")
        fresh = _collect(connection, plan, metadata)
        if (
            fresh.target_rowids != preview.target_rowids
            or fresh.target_signature != preview.target_signature
            or fresh.affected_cells != preview.affected_cells
        ):
            raise ValueError("The data changed after the preview. Build the plan again.")

        rowids = preview.target_rowids
        changed_rows = len(rowids)
        if plan.action == Action.UPDATE_ROWS:
            set_sql = ", ".join(
                f"{quote_identifier(item.column.name)} = ?" for item in plan.assignments
            )
            set_values = [item.value for item in plan.assignments]
            for chunk in _chunks(rowids):
                placeholders = ",".join("?" for _ in chunk)
                connection.execute(
                    f"UPDATE {quoted_table} SET {set_sql} "
                    f"WHERE {table.rowid_alias} IN ({placeholders})",
                    [*set_values, *chunk],
                )
        else:
            for chunk in _chunks(rowids):
                placeholders = ",".join("?" for _ in chunk)
                connection.execute(
                    f"DELETE FROM {quoted_table} "
                    f"WHERE {table.rowid_alias} IN ({placeholders})",
                    chunk,
                )

        remaining = 0
        for chunk in _chunks(rowids):
            placeholders = ",".join("?" for _ in chunk)
            remaining += int(
                connection.execute(
                    f"SELECT COUNT(*) FROM {quoted_table} "
                    f"WHERE {table.rowid_alias} IN ({placeholders})",
                    chunk,
                ).fetchone()[0]
            )
        if plan.action == Action.DELETE_ROWS and remaining != 0:
            raise RuntimeError("DELETE verification failed.")
        if plan.action == Action.UPDATE_ROWS and remaining != len(rowids):
            raise RuntimeError("UPDATE verification failed: rows disappeared.")
        if plan.action == Action.UPDATE_ROWS:
            assignment_names = [item.column.name for item in plan.assignments]
            assignment_sql = ", ".join(
                quote_identifier(name) for name in assignment_names
            )
            expected_values = [item.value for item in plan.assignments]
            for chunk in _chunks(rowids):
                placeholders = ",".join("?" for _ in chunk)
                rows = connection.execute(
                    f"SELECT {assignment_sql} FROM {quoted_table} "
                    f"WHERE {table.rowid_alias} IN ({placeholders})",
                    chunk,
                ).fetchall()
                if any(
                    not all(
                        _equivalent(row[index], expected_values[index])
                        for index in range(len(expected_values))
                    )
                    for row in rows
                ):
                    raise RuntimeError("Verification of the new UPDATE values failed.")
        check = connection.execute("PRAGMA quick_check").fetchone()
        if check is None or check[0] != "ok":
            raise RuntimeError(f"SQLite quick_check failed: {check}")
        connection.commit()
        committed = True
    except Exception:
        if connection.in_transaction:
            connection.rollback()
        raise
    finally:
        connection.close()

    return SqlExecutionOutcome(
        snapshot_path=snapshot_path,
        database_path=database_path.resolve(),
        changed_rows=len(preview.target_rowids),
        affected_cells=preview.affected_cells,
        transaction_committed=committed,
        result_verified=committed,
    )
