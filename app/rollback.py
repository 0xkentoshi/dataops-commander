from __future__ import annotations

import hashlib
import os
import shutil
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

from openpyxl import load_workbook


@dataclass(frozen=True, slots=True)
class UndoOutcome:
    snapshot_path: Path
    restored_from: Path
    result_sha256: str


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _verify_file(path: Path) -> None:
    suffix = path.suffix.casefold()
    if suffix == ".xlsx":
        workbook = load_workbook(path, read_only=True, data_only=False)
        workbook.close()
        return
    if suffix in {".db", ".sqlite", ".sqlite3"}:
        connection = sqlite3.connect(f"file:{path.resolve()}?mode=ro", uri=True)
        try:
            result = connection.execute("PRAGMA quick_check").fetchone()
        finally:
            connection.close()
        if not result or str(result[0]).casefold() != "ok":
            raise RuntimeError("The SQLite snapshot failed PRAGMA quick_check.")
        return
    raise ValueError(f"Undo currently supports Excel and SQLite; received {path.suffix or 'a file without an extension'}.")


def restore_snapshot(
    source_path: Path,
    snapshot_path: Path,
    *,
    expected_current_sha256: str,
    snapshots_root: Path = Path("data/snapshots/undo"),
) -> UndoOutcome:
    source_path = source_path.resolve()
    snapshot_path = snapshot_path.resolve()
    if not source_path.is_file():
        raise ValueError("The current file no longer exists.")
    if not snapshot_path.is_file():
        raise ValueError("The rollback snapshot no longer exists.")
    if source_path.suffix.casefold() != snapshot_path.suffix.casefold():
        raise ValueError("The snapshot has a different file format.")
    current_sha = sha256_file(source_path)
    if current_sha != expected_current_sha256:
        raise ValueError(
            "The file changed after the latest operation. Automatic rollback is blocked to protect newer changes."
        )

    snapshots_root.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    backup = snapshots_root / f"{source_path.stem}_{timestamp}_{uuid4().hex[:10]}_before_undo{source_path.suffix}"
    temporary = source_path.parent / f".dataops-{source_path.stem}-{uuid4().hex[:10]}.undo.tmp{source_path.suffix}"
    restore = source_path.parent / f".dataops-{source_path.stem}-{uuid4().hex[:10]}.undo.restore{source_path.suffix}"

    shutil.copy2(source_path, backup)
    shutil.copy2(snapshot_path, temporary)
    _verify_file(temporary)

    replaced = False
    try:
        os.replace(temporary, source_path)
        replaced = True
        _verify_file(source_path)
    except Exception as error:
        if replaced:
            shutil.copy2(backup, restore)
            os.replace(restore, source_path)
        raise error
    finally:
        temporary.unlink(missing_ok=True)
        restore.unlink(missing_ok=True)

    return UndoOutcome(
        snapshot_path=backup,
        restored_from=snapshot_path,
        result_sha256=sha256_file(source_path),
    )
