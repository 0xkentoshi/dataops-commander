from __future__ import annotations

import re
from enum import Enum
from typing import TypeAlias


CellLiteral: TypeAlias = str | int | float | bool


class ReplyKind(str, Enum):
    CONFIRM = "confirm"
    CANCEL = "cancel"
    OTHER = "other"


_CONFIRM_PHRASES = {
    "yes",
    "ага",
    "верно",
    "вперед",
    "го",
    "да",
    "давай",
    "давай делай",
    "давай запускай",
    "да оно",
    "делай",
    "запускай",
    "именно",
    "можно",
    "ну да",
    "ок",
    "окей",
    "подтверждаю",
    "применяй",
    "согласен",
    "согласна",
    "точно",
    "угу",
    "выполняй",
}

_CANCEL_PHRASES = {
    "cancel",
    "stop",
    "не делай",
    "не надо",
    "нет",
    "отбой",
    "отмена",
    "отмени",
    "стоп",
}


def normalize_reply(text: str) -> str:
    normalized = text.casefold().replace("ё", "е")
    normalized = re.sub(r"[^0-9a-zа-я]+", " ", normalized)
    return " ".join(normalized.split())


def classify_reply(text: str) -> ReplyKind:
    """Распознаёт только короткий ответ, а не полноценную новую команду."""
    normalized = normalize_reply(text)
    if not normalized or len(normalized.split()) > 5:
        return ReplyKind.OTHER
    if normalized in _CONFIRM_PHRASES:
        return ReplyKind.CONFIRM
    if normalized in _CANCEL_PHRASES:
        return ReplyKind.CANCEL
    if normalized.startswith(
        (
            "да давай",
            "да делай",
            "да выполняй",
            "да запускай",
            "да применяй",
            "да сделай",
        )
    ):
        return ReplyKind.CONFIRM
    if normalized.startswith(("нет отмен", "нет не надо")):
        return ReplyKind.CANCEL
    return ReplyKind.OTHER


def is_replace_all_command(text: str) -> bool:
    normalized = normalize_reply(text)
    has_scope = bool(re.search(r"\b(все|всех|каждую|каждой|целиком)\b", normalized))
    has_cells = bool(re.search(r"\bячейк\w*\b", normalized))
    has_verb = bool(re.search(r"\b(замен\w*|измени\w*|поменя\w*|заполни\w*)\b", normalized))
    has_value_marker = bool(re.search(r"\b(на|значением)\b", normalized))
    whole_file = bool(
        re.search(
            r"\b(вообще|весь\s+файл|всю\s+книгу|во\s+всем\s+файле|"
            r"во\s+всей\s+книге|все\s+заполненн\w*\s+ячейк\w*)\b",
            normalized,
        )
    )
    narrow_column = bool(re.search(r"\b(?:в|из)\s+столбц\w*\b", normalized))
    return (
        has_scope
        and has_cells
        and has_verb
        and has_value_marker
        and whole_file
        and not narrow_column
    )


def parse_cell_literal(value: str) -> CellLiteral:
    cleaned = value.strip().strip("\"'«»„“” ").rstrip(".!?")
    lowered = cleaned.casefold().replace("ё", "е")
    if lowered in {"true", "истина"}:
        return True
    if lowered in {"false", "ложь"}:
        return False
    if re.fullmatch(r"[-+]?\d+", cleaned):
        return int(cleaned)
    if re.fullmatch(r"[-+]?(?:\d+[.,]\d+|\d+\.)", cleaned):
        return float(cleaned.replace(",", "."))
    return cleaned


def replacement_value_from_text(text: str) -> CellLiteral | None:
    """Достаёт X из фраз вида «замени все ячейки ... на X»."""
    markers = list(
        re.finditer(
            r"\b(?:на|значением)\b\s+",
            text,
            flags=re.IGNORECASE,
        )
    )
    if not markers:
        return None
    candidate = text[markers[-1].end() :].strip()
    candidate = re.split(
        r",\s*(?:ну\s+)?(?:вообще\s+)?(?:прям\s+)?все\b",
        candidate,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0]
    candidate = candidate.strip()
    return parse_cell_literal(candidate) if candidate else None
