from __future__ import annotations

import re
from pathlib import Path

from pydantic import BaseModel, ConfigDict


class ExecutionOptions(BaseModel):
    """Execution behavior already understood by the AI command interpreter.

    This module deliberately contains no natural-language recognition. Language
    belongs to the LLM; Python only stores/validates execution parameters.
    """

    model_config = ConfigDict(extra="forbid")

    copy_original: bool = False
    output_name: str | None = None
    repeated_from_operation_id: str | None = None


def normalized_output_filename(requested: str | None, source_path: Path) -> str | None:
    """Sanitize a model/user supplied filename for the local filesystem.

    This regex is filesystem validation, not NLP.
    """
    if requested is None:
        return None
    name = Path(requested.strip()).name
    cleaned = re.sub(r"[^0-9A-Za-zА-Яа-яЁё._() -]+", "_", name).strip(" .")[:140]
    if not cleaned:
        return None
    suffix = source_path.suffix
    if Path(cleaned).suffix.casefold() != suffix.casefold():
        cleaned = f"{Path(cleaned).stem or cleaned}{suffix}"
    return cleaned
