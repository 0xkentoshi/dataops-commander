from __future__ import annotations

import hashlib
import os
import re
import shutil
import sqlite3
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from app.excel_operations import (
    ExcelOperationPreview,
    WRITE_ACTIONS,
    build_operation_preview,
    execute_operation,
)
from app.excel_planner import ExcelOperationPlan
from app.excel_service import WorkbookMetadata, inspect_workbook
from app.sql_operations import (
    SQL_WRITE_ACTIONS,
    SqlOperationPreview,
    build_sql_preview,
    execute_sql_operation,
)
from app.sql_planner import SqlOperationPlan
from app.sql_service import SqliteMetadata, inspect_sqlite


@dataclass(frozen=True, slots=True)
class BatchExecutionOutcome:
    snapshot_path: Path
    result_path: Path
    changed_rows: int
    affected_cells: int
    result_verified: bool


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_stem(name: str, fallback: str) -> str:
    value = re.sub(
        r"[^0-9A-Za-zА-Яа-яЁё._() -]+",
        "_",
        Path(name).stem,
    ).strip(" .")
    return value or fallback


def _snapshot_name(original_name: str, batch_id: str, suffix: str) -> str:
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    stem = _safe_stem(original_name, "source")
    return f"{stem}_{timestamp}_{batch_id}_snapshot{suffix}"


def copy_sqlite_database(source_path: Path, destination_path: Path) -> None:
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    source = sqlite3.connect(source_path, timeout=10)
    destination = sqlite3.connect(destination_path, timeout=10)
    try:
        source.backup(destination)
        check = destination.execute("PRAGMA quick_check").fetchone()
        if check is None or check[0] != "ok":
            raise RuntimeError(f"SQLite backup повреждён: {check}")
    finally:
        destination.close()
        source.close()


def execute_excel_batch(
    source_path: Path,
    original_name: str,
    batch_id: str,
    steps: list[tuple[ExcelOperationPlan, ExcelOperationPreview, WorkbookMetadata]],
    snapshots_root: Path = Path("data/snapshots/excel"),
) -> BatchExecutionOutcome:
    if not steps:
        raise ValueError("Batch Excel пуст.")
    write_steps = [item for item in steps if item[0].action in WRITE_ACTIONS and item[1].has_changes]
    if not write_steps:
        raise ValueError("Batch не содержит изменений.")

    initial_hash = steps[0][1].source_sha256
    if _sha256(source_path) != initial_hash:
        raise ValueError("Исходный файл изменился после preview. Создайте план заново.")

    snapshots_root.mkdir(parents=True, exist_ok=True)
    snapshot_path = snapshots_root / _snapshot_name(original_name, batch_id, ".xlsx")
    shutil.copy2(source_path, snapshot_path)
    if _sha256(snapshot_path) != initial_hash:
        raise RuntimeError("Snapshot Excel не совпал с исходником.")

    fd, temporary_name = tempfile.mkstemp(
        prefix=f".dataops-batch-{source_path.stem}-",
        suffix=".xlsx",
        dir=source_path.parent,
    )
    os.close(fd)
    temporary_path = Path(temporary_name)
    restore_path = source_path.parent / f".dataops-{source_path.stem}-{batch_id}.restore.xlsx"
    shutil.copy2(source_path, temporary_path)

    changed_rows = 0
    affected_cells = 0
    replaced = False
    try:
        with tempfile.TemporaryDirectory(prefix="dataops-batch-snapshots-") as temp_snapshots:
            temp_root = Path(temp_snapshots)
            for index, (plan, preview, _metadata) in enumerate(steps, start=1):
                current_metadata = inspect_workbook(temporary_path, original_name)
                fresh = build_operation_preview(temporary_path, plan, current_metadata)
                stored_semantic = preview.model_dump(mode="json")
                fresh_semantic = fresh.model_dump(mode="json")
                # XLSX — zip-контейнер: одинаковая логическая книга после save может
                # иметь другой byte-level SHA из-за служебных timestamps. Сравниваем
                # подтверждённое содержимое, а execute_operation получает свежий SHA.
                stored_semantic.pop("source_sha256", None)
                fresh_semantic.pop("source_sha256", None)
                if fresh_semantic != stored_semantic:
                    raise ValueError(
                        f"Данные для задачи {index} изменились после preview. "
                        "Создайте план заново."
                    )
                if plan.action in WRITE_ACTIONS and fresh.has_changes:
                    outcome = execute_operation(
                        temporary_path,
                        original_name,
                        f"{batch_id}-{index}",
                        plan,
                        fresh,
                        current_metadata,
                        temp_root,
                    )
                    changed_rows += outcome.changed_rows
                    affected_cells += outcome.affected_cells

        # Внешнее изменение источника между planning и confirm не должно быть затёрто.
        if _sha256(source_path) != initial_hash:
            raise ValueError("Исходный файл изменился после preview. Создайте план заново.")

        os.replace(temporary_path, source_path)
        replaced = True
        inspect_workbook(source_path, original_name)
    except Exception as error:
        if replaced:
            try:
                shutil.copy2(snapshot_path, restore_path)
                os.replace(restore_path, source_path)
            except Exception as restore_error:
                raise RuntimeError(
                    "Batch Excel записан, проверка не прошла и восстановление "
                    f"не удалось: {restore_error}"
                ) from error
        raise
    finally:
        temporary_path.unlink(missing_ok=True)
        restore_path.unlink(missing_ok=True)

    return BatchExecutionOutcome(
        snapshot_path=snapshot_path,
        result_path=source_path.resolve(),
        changed_rows=changed_rows,
        affected_cells=affected_cells,
        result_verified=True,
    )


def execute_sqlite_batch(
    source_path: Path,
    original_name: str,
    batch_id: str,
    steps: list[tuple[SqlOperationPlan, SqlOperationPreview, SqliteMetadata]],
    snapshots_root: Path = Path("data/snapshots/sqlite"),
) -> BatchExecutionOutcome:
    if not steps:
        raise ValueError("Batch SQLite пуст.")
    write_steps = [
        item for item in steps if item[0].action in SQL_WRITE_ACTIONS and item[1].has_changes
    ]
    if not write_steps:
        raise ValueError("Batch не содержит изменений.")

    snapshots_root.mkdir(parents=True, exist_ok=True)
    snapshot_path = snapshots_root / _snapshot_name(original_name, batch_id, ".sqlite3")
    copy_sqlite_database(source_path, snapshot_path)

    with tempfile.TemporaryDirectory(prefix="dataops-sqlite-batch-") as temporary:
        root = Path(temporary)
        stage_path = root / "stage.sqlite3"
        copy_sqlite_database(source_path, stage_path)

        changed_rows = 0
        affected_cells = 0
        for index, (plan, preview, _metadata) in enumerate(steps, start=1):
            current_metadata = inspect_sqlite(stage_path, original_name)
            fresh = build_sql_preview(stage_path, plan, current_metadata)
            if (
                fresh.target_rowids != preview.target_rowids
                or fresh.target_signature != preview.target_signature
                or fresh.affected_cells != preview.affected_cells
                or fresh.matched_rows != preview.matched_rows
            ):
                raise ValueError(
                    f"Данные для задачи {index} изменились после preview. "
                    "Создайте план заново."
                )
            if plan.action in SQL_WRITE_ACTIONS and fresh.has_changes:
                outcome = execute_sql_operation(
                    stage_path,
                    f"{batch_id}-{index}",
                    plan,
                    fresh,
                    current_metadata,
                    root / "snapshots",
                )
                changed_rows += outcome.changed_rows
                affected_cells += outcome.affected_cells

        # SQLite backup API даёт согласованную замену содержимого БД и не требует raw SQL.
        try:
            copy_sqlite_database(stage_path, source_path)
            inspect_sqlite(source_path, original_name)
        except Exception as error:
            try:
                copy_sqlite_database(snapshot_path, source_path)
            except Exception as restore_error:
                raise RuntimeError(
                    "Batch SQLite не прошёл проверку и восстановление snapshot "
                    f"не удалось: {restore_error}"
                ) from error
            raise

    return BatchExecutionOutcome(
        snapshot_path=snapshot_path,
        result_path=source_path.resolve(),
        changed_rows=changed_rows,
        affected_cells=affected_cells,
        result_verified=True,
    )
