from __future__ import annotations

import sqlite3
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field


MAX_TABLES = 50
MAX_COLUMNS_PER_TABLE = 120
MAX_SAMPLE_VALUES = 4


class SqlColumnMetadata(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str
    declared_type: str
    not_null: bool
    primary_key_position: int
    samples: list[str] = Field(default_factory=list)


class SqlTableMetadata(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str
    row_count: int
    columns: list[SqlColumnMetadata] = Field(default_factory=list)
    supports_rowid: bool
    rowid_alias: str | None = None
    foreign_key_count: int = 0
    referenced_by_foreign_keys: int = 0
    trigger_count: int = 0


class SqliteMetadata(BaseModel):
    model_config = ConfigDict(extra="forbid")

    original_name: str
    tables: list[SqlTableMetadata] = Field(default_factory=list)
    journal_mode: str
    warnings: list[str] = Field(default_factory=list)


def quote_identifier(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def open_read_only(database_path: Path) -> sqlite3.Connection:
    uri = database_path.resolve().as_uri() + "?mode=ro"
    connection = sqlite3.connect(uri, uri=True, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA busy_timeout = 5000")
    connection.execute("PRAGMA query_only = ON")
    return connection


def _display(value: object) -> str:
    if value is None:
        return "<NULL>"
    if isinstance(value, bytes):
        return f"<BLOB {len(value)} bytes>"
    text = str(value).replace("\n", " ").strip()
    return text if len(text) <= 80 else f"{text[:79]}…"


def _rowid_alias(connection: sqlite3.Connection, table_name: str, names: set[str]) -> str | None:
    for candidate in ("rowid", "_rowid_", "oid"):
        if candidate.casefold() in names:
            continue
        try:
            connection.execute(
                f"SELECT {candidate} FROM {quote_identifier(table_name)} LIMIT 1"
            ).fetchone()
            return candidate
        except sqlite3.DatabaseError:
            continue
    return None


def inspect_sqlite(database_path: Path, original_name: str) -> SqliteMetadata:
    if not database_path.is_file():
        raise ValueError("SQLite-файл больше не существует.")
    warnings: list[str] = []
    tables: list[SqlTableMetadata] = []
    try:
        connection = open_read_only(database_path)
    except sqlite3.DatabaseError as error:
        raise ValueError(f"Файл не является читаемой SQLite-базой: {error}") from error

    try:
        integrity = connection.execute("PRAGMA quick_check").fetchone()
        if integrity is None or integrity[0] != "ok":
            raise ValueError(f"SQLite quick_check не пройден: {integrity}")
        journal_row = connection.execute("PRAGMA journal_mode").fetchone()
        journal_mode = str(journal_row[0] if journal_row else "unknown")
        table_rows = connection.execute(
            """
            SELECT name
            FROM sqlite_schema
            WHERE type = 'table' AND name NOT LIKE 'sqlite_%'
            ORDER BY name
            LIMIT ?
            """,
            (MAX_TABLES + 1,),
        ).fetchall()
        if len(table_rows) > MAX_TABLES:
            warnings.append(f"Показаны первые {MAX_TABLES} таблиц.")

        visible_table_names = [str(row["name"]) for row in table_rows[:MAX_TABLES]]
        inbound_foreign_keys = {name: 0 for name in visible_table_names}
        inbound_name_lookup = {
            name.casefold(): name for name in visible_table_names
        }
        for source_table in visible_table_names:
            quoted_source = quote_identifier(source_table)
            for foreign_key in connection.execute(
                f"PRAGMA foreign_key_list({quoted_source})"
            ).fetchall():
                referenced_table = str(foreign_key["table"])
                canonical_name = inbound_name_lookup.get(referenced_table.casefold())
                if canonical_name is not None:
                    inbound_foreign_keys[canonical_name] += 1

        for table_row in table_rows[:MAX_TABLES]:
            table_name = str(table_row["name"])
            quoted_table = quote_identifier(table_name)
            column_rows = connection.execute(
                f"PRAGMA table_xinfo({quoted_table})"
            ).fetchall()
            visible_rows = [row for row in column_rows if int(row["hidden"]) == 0]
            if len(visible_rows) > MAX_COLUMNS_PER_TABLE:
                warnings.append(
                    f"Таблица «{table_name}»: показаны первые "
                    f"{MAX_COLUMNS_PER_TABLE} столбцов."
                )
            names = {str(row["name"]).casefold() for row in visible_rows}
            rowid_alias = _rowid_alias(connection, table_name, names)
            row_count = int(
                connection.execute(f"SELECT COUNT(*) FROM {quoted_table}").fetchone()[0]
            )
            foreign_key_count = len(
                connection.execute(f"PRAGMA foreign_key_list({quoted_table})").fetchall()
            )
            trigger_count = int(
                connection.execute(
                    "SELECT COUNT(*) FROM sqlite_schema "
                    "WHERE type = 'trigger' AND tbl_name = ?",
                    (table_name,),
                ).fetchone()[0]
            )
            columns: list[SqlColumnMetadata] = []
            for column_row in visible_rows[:MAX_COLUMNS_PER_TABLE]:
                column_name = str(column_row["name"])
                samples = [
                    _display(row[0])
                    for row in connection.execute(
                        f"SELECT {quote_identifier(column_name)} FROM {quoted_table} "
                        f"WHERE {quote_identifier(column_name)} IS NOT NULL LIMIT ?",
                        (MAX_SAMPLE_VALUES,),
                    ).fetchall()
                ]
                columns.append(
                    SqlColumnMetadata(
                        name=column_name,
                        declared_type=str(column_row["type"] or ""),
                        not_null=bool(column_row["notnull"]),
                        primary_key_position=int(column_row["pk"]),
                        samples=samples,
                    )
                )
            tables.append(
                SqlTableMetadata(
                    name=table_name,
                    row_count=row_count,
                    columns=columns,
                    supports_rowid=rowid_alias is not None,
                    rowid_alias=rowid_alias,
                    foreign_key_count=foreign_key_count,
                    referenced_by_foreign_keys=inbound_foreign_keys[table_name],
                    trigger_count=trigger_count,
                )
            )
    except sqlite3.DatabaseError as error:
        raise ValueError(f"Ошибка чтения SQLite: {error}") from error
    finally:
        connection.close()

    if not tables:
        warnings.append("Пользовательские таблицы не найдены.")
    return SqliteMetadata(
        original_name=original_name,
        tables=tables,
        journal_mode=journal_mode,
        warnings=warnings,
    )
