from __future__ import annotations

import hashlib
import os
import re
import shutil
from dataclasses import dataclass
from difflib import SequenceMatcher
from enum import Enum
from pathlib import Path
from uuid import uuid4


MAX_CATALOG_FILES = 250
IGNORED_DIRECTORIES = {
    ".git",
    ".idea",
    ".venv",
    ".vscode",
    "__pycache__",
    "data",
    "results",
    "snapshots",
}


class SourceOrigin(str, Enum):
    LOCAL = "local"
    TELEGRAM = "telegram"


class SourceKind(str, Enum):
    EXCEL = "excel"
    SQLITE = "sqlite"
    CSV = "csv"
    TSV = "tsv"
    SQL_SCRIPT = "sql_script"
    JSON = "json"
    OTHER = "other"


SOURCE_KIND_LABELS = {
    SourceKind.EXCEL: "Excel",
    SourceKind.SQLITE: "SQLite",
    SourceKind.CSV: "CSV",
    SourceKind.TSV: "TSV",
    SourceKind.SQL_SCRIPT: "SQL-скрипт",
    SourceKind.JSON: "JSON",
    SourceKind.OTHER: "Файл",
}

EDITABLE_SOURCE_KINDS = {SourceKind.EXCEL, SourceKind.SQLITE}


@dataclass(frozen=True, slots=True)
class DataSource:
    source_id: str
    origin: SourceOrigin
    kind: SourceKind
    path: Path
    display_name: str
    relative_name: str
    extension: str
    size_bytes: int
    modified_ns: int

    @property
    def origin_icon(self) -> str:
        return "📁" if self.origin == SourceOrigin.LOCAL else "📨"

    @property
    def origin_label(self) -> str:
        return "Локальный" if self.origin == SourceOrigin.LOCAL else "ТГ"

    @property
    def kind_label(self) -> str:
        return SOURCE_KIND_LABELS[self.kind]

    @property
    def editable(self) -> bool:
        return self.kind in EDITABLE_SOURCE_KINDS


def clean_filename(filename: str) -> str:
    name = Path(filename).name
    cleaned = re.sub(
        r"[^0-9A-Za-zА-Яа-яЁё._() -]+",
        "_",
        name,
    ).strip(" .")
    return cleaned[:140] or "uploaded_file"


def detect_source_kind(path: Path) -> SourceKind:
    suffix = path.suffix.casefold()
    if suffix == ".xlsx":
        return SourceKind.EXCEL
    if suffix in {".sqlite", ".sqlite3", ".db"}:
        return SourceKind.SQLITE
    if suffix == ".csv":
        return SourceKind.CSV
    if suffix in {".tsv", ".tab"}:
        return SourceKind.TSV
    if suffix == ".sql":
        return SourceKind.SQL_SCRIPT
    if suffix == ".json":
        return SourceKind.JSON
    return SourceKind.OTHER


def _inside(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _stable_id(origin: SourceOrigin, path: Path) -> str:
    normalized = os.path.normcase(str(path.resolve()))
    payload = f"{origin.value}|{normalized}".encode("utf-8")
    return hashlib.sha256(payload).hexdigest()[:16]


def _source(
    path: Path,
    origin: SourceOrigin,
    display_name: str,
    relative_name: str,
) -> DataSource | None:
    try:
        stat = path.stat()
    except (FileNotFoundError, OSError):
        return None
    if not path.is_file():
        return None
    return DataSource(
        source_id=_stable_id(origin, path),
        origin=origin,
        kind=detect_source_kind(path),
        path=path.resolve(),
        display_name=display_name,
        relative_name=relative_name,
        extension=path.suffix.casefold() or "без расширения",
        size_bytes=stat.st_size,
        modified_ns=stat.st_mtime_ns,
    )


class SourceCatalog:
    def __init__(
        self,
        workspace_root: Path,
        telegram_root: Path,
        excluded_paths: set[Path] | None = None,
        legacy_telegram_root: Path | None = None,
    ) -> None:
        self.workspace_root = workspace_root.resolve()
        self.telegram_root = telegram_root.resolve()
        self.legacy_telegram_root = (
            legacy_telegram_root.resolve() if legacy_telegram_root else None
        )
        self.excluded_paths = {
            path.resolve() for path in (excluded_paths or set())
        }

    def initialize(self) -> None:
        self.workspace_root.mkdir(parents=True, exist_ok=True)
        self.telegram_root.mkdir(parents=True, exist_ok=True)

    def _local_sources(self) -> list[DataSource]:
        result: list[DataSource] = []
        if not self.workspace_root.is_dir():
            return result
        for path in self.workspace_root.rglob("*"):
            if len(result) >= MAX_CATALOG_FILES:
                break
            try:
                relative = path.relative_to(self.workspace_root)
            except ValueError:
                continue
            if any(part in IGNORED_DIRECTORIES for part in relative.parts[:-1]):
                continue
            if path.name.startswith(("~$", ".dataops-")):
                continue
            if path.name.casefold().endswith(("-wal", "-shm", "-journal")):
                continue
            if path.resolve() in self.excluded_paths:
                continue
            if path.is_symlink() and not _inside(path, self.workspace_root):
                continue
            item = _source(
                path,
                SourceOrigin.LOCAL,
                path.name,
                str(relative),
            )
            if item is not None:
                result.append(item)
        return result

    def _telegram_sources(self, user_id: int) -> list[DataSource]:
        user_root = self.telegram_root / str(user_id)
        result: list[DataSource] = []
        if not user_root.is_dir():
            return result
        for upload_directory in user_root.iterdir():
            if not upload_directory.is_dir():
                continue
            files = sorted(
                (
                    item
                    for item in upload_directory.iterdir()
                    if item.is_file()
                    and not item.name.casefold().endswith(("-wal", "-shm", "-journal"))
                ),
                key=lambda item: item.name.casefold(),
            )
            if not files:
                continue
            path = files[0]
            item = _source(
                path,
                SourceOrigin.TELEGRAM,
                path.name,
                path.name,
            )
            if item is not None:
                result.append(item)
        return result

    def _legacy_telegram_sources(self, user_id: int) -> list[DataSource]:
        if self.legacy_telegram_root is None:
            return []
        user_root = self.legacy_telegram_root / str(user_id)
        if not user_root.is_dir():
            return []
        result: list[DataSource] = []
        for path in sorted(user_root.iterdir(), key=lambda item: item.name.casefold()):
            if not path.is_file() or path.name.casefold().endswith(
                ("-wal", "-shm", "-journal")
            ):
                continue
            display_name = path.name
            prefix, separator, remainder = display_name.partition("_")
            if separator and len(prefix) == 32 and all(
                character in "0123456789abcdefABCDEF" for character in prefix
            ):
                display_name = remainder or display_name
            item = _source(
                path,
                SourceOrigin.TELEGRAM,
                display_name,
                display_name,
            )
            if item is not None:
                result.append(item)
        return result

    def list_sources(self, user_id: int) -> list[DataSource]:
        sources = [
            *self._local_sources(),
            *self._telegram_sources(user_id),
            *self._legacy_telegram_sources(user_id),
        ]
        return sorted(
            sources,
            key=lambda item: (
                0 if item.origin == SourceOrigin.LOCAL else 1,
                item.relative_name.casefold(),
            ),
        )

    def get(self, user_id: int, source_id: str) -> DataSource | None:
        return next(
            (
                source
                for source in self.list_sources(user_id)
                if source.source_id == source_id
            ),
            None,
        )

    def find_by_path(self, user_id: int, path: Path) -> DataSource | None:
        resolved = path.resolve()
        return next(
            (
                source
                for source in self.list_sources(user_id)
                if source.path == resolved
            ),
            None,
        )

    def allocate_telegram_upload(self, user_id: int, filename: str) -> Path:
        directory = self.telegram_root / str(user_id) / uuid4().hex
        directory.mkdir(parents=True, exist_ok=False)
        return directory / clean_filename(filename)

    def discard_telegram_upload(self, path: Path) -> None:
        parent = path.resolve().parent
        user_root = parent.parent
        if _inside(parent, self.telegram_root) and _inside(user_root, self.telegram_root):
            shutil.rmtree(parent, ignore_errors=True)


_ALIASES = {
    "прайс": "price",
    "цены": "price",
    "цена": "price",
    "база": "database",
    "бд": "database",
    "склад": "warehouse",
}


def _normalized(value: str) -> str:
    text = value.casefold().replace("ё", "е")
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    tokens = [_ALIASES.get(token, token) for token in text.split()]
    return " ".join(tokens)


def _match_score(hint: str, source: DataSource) -> float:
    query = _normalized(hint)
    full_name = _normalized(source.display_name)
    stem = _normalized(Path(source.display_name).stem)
    relative = _normalized(source.relative_name)
    if not query:
        return 0.0
    if query in {full_name, stem, relative}:
        return 1.0
    if relative and relative in query:
        return 0.98
    if stem and stem in query:
        return 0.94
    if query in full_name or query in relative:
        return 0.91
    query_tokens = set(query.split())
    name_tokens = set(full_name.split())
    overlap = len(query_tokens & name_tokens) / max(len(query_tokens), 1)
    similarity = max(
        SequenceMatcher(None, query, full_name).ratio(),
        SequenceMatcher(None, query, stem).ratio(),
    )
    return max(overlap * 0.86, similarity * 0.82)


def match_sources(hint: str, sources: list[DataSource]) -> list[DataSource]:
    ranked = sorted(
        ((_match_score(hint, source), source) for source in sources),
        key=lambda pair: pair[0],
        reverse=True,
    )
    if not ranked or ranked[0][0] < 0.58:
        return []
    best = ranked[0][0]
    return [source for score, source in ranked if score >= best - 0.07 and score >= 0.58]
