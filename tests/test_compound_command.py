from __future__ import annotations

import sys
import types
import unittest

try:
    import ollama  # noqa: F401
except (ModuleNotFoundError, ImportError):
    ollama_stub = types.ModuleType("ollama")
    ollama_stub.AsyncClient = object
    sys.modules["ollama"] = ollama_stub

from app.intent import (
    Action,
    CommandEnvelope,
    CommandMode,
    ExecutionDirective,
    IntentDraft,
    SourceType,
)


def _draft(
    action: Action,
    normalized: str,
    *,
    columns: list[str] | None = None,
    filters: list[str] | None = None,
    values: list[str] | None = None,
) -> IntentDraft:
    return IntentDraft(
        normalized_request=normalized,
        action=action,
        source_type_hint=SourceType.EXCEL,
        source_name_hint=None,
        resource_name_hint="clients",
        column_hints=columns or [],
        filter_hints=filters or [],
        value_hints=values or [],
        is_write_operation=action != Action.SELECT,
        is_destructive=action in {Action.DELETE_ROWS, Action.DEDUPLICATE},
        needs_discovery=False,
        confidence=0.99,
        interpretation_note="test",
        clarification_question=None,
    )


class AiNativeCompoundCommandTests(unittest.TestCase):
    def test_full_demo_command_is_represented_as_four_ai_tasks(self) -> None:
        envelope = CommandEnvelope(
            mode=CommandMode.DATA,
            execution=ExecutionDirective(copy_original=True),
            tasks=[
                _draft(
                    Action.DELETE_ROWS,
                    "Удалить строки, где email пустой",
                    columns=["email"],
                    filters=["email is empty"],
                ),
                _draft(
                    Action.DEDUPLICATE,
                    "Удалить дубли по email",
                    columns=["email"],
                ),
                _draft(
                    Action.RENAME_COLUMN,
                    "Переименовать client в client_name",
                    columns=["client"],
                    values=["client_name"],
                ),
                _draft(
                    Action.DELETE_ROWS,
                    "Удалить строки, где city != Новосибирск",
                    columns=["city"],
                    filters=["city != Новосибирск"],
                ),
            ],
            confidence=0.99,
            interpretation_note="four tasks",
        )
        self.assertEqual(len(envelope.tasks), 4)
        self.assertEqual(envelope.tasks[0].action, Action.DELETE_ROWS)
        self.assertEqual(envelope.tasks[1].action, Action.DEDUPLICATE)
        self.assertEqual(envelope.tasks[1].column_hints, ["email"])
        self.assertEqual(envelope.tasks[2].action, Action.RENAME_COLUMN)
        self.assertEqual(envelope.tasks[3].action, Action.DELETE_ROWS)
        self.assertTrue(envelope.execution.copy_original)

    def test_micro_commands_are_modes_not_regex_helpers(self) -> None:
        self.assertEqual(
            CommandEnvelope(
                mode=CommandMode.SEND_LAST_FILE,
                confidence=0.99,
                interpretation_note="send",
            ).mode,
            CommandMode.SEND_LAST_FILE,
        )
        self.assertEqual(
            CommandEnvelope(
                mode=CommandMode.REPEAT_LAST,
                target_source_name="october.xlsx",
                confidence=0.99,
                interpretation_note="repeat",
            ).target_source_name,
            "october.xlsx",
        )
        self.assertTrue(
            CommandEnvelope(
                mode=CommandMode.UPDATE_PENDING,
                execution=ExecutionDirective(copy_original=True),
                confidence=0.99,
                interpretation_note="copy",
            ).execution.copy_original
        )


if __name__ == "__main__":
    unittest.main()
