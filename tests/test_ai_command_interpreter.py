from __future__ import annotations

import asyncio
import json
import sys
import types
import unittest
from pathlib import Path
from types import SimpleNamespace

try:
    import ollama  # noqa: F401
except (ModuleNotFoundError, ImportError):
    ollama_stub = types.ModuleType("ollama")

    class AsyncClientStub:
        def __init__(self, *args, **kwargs) -> None:
            pass

    ollama_stub.AsyncClient = AsyncClientStub
    sys.modules["ollama"] = ollama_stub

from app.intent import Action, CommandMode, IntentParser


class FakeClient:
    def __init__(self, response_payload: dict) -> None:
        self.response_payload = response_payload
        self.calls: list[dict] = []

    async def chat(self, **kwargs):
        self.calls.append(kwargs)
        return SimpleNamespace(
            message=SimpleNamespace(content=json.dumps(self.response_payload, ensure_ascii=False))
        )


class AiCommandInterpreterTests(unittest.TestCase):
    def _parser(self, response_payload: dict) -> IntentParser:
        settings = SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        parser = IntentParser(settings)
        parser._client = FakeClient(response_payload)
        return parser

    def test_active_schema_is_sent_to_ai_and_four_tasks_survive(self) -> None:
        response = {
            "mode": "data",
            "tasks": [
                {
                    "normalized_request": "Удалить строки, где email пустой",
                    "action": "delete_rows",
                    "source_type_hint": "excel",
                    "source_name_hint": None,
                    "resource_name_hint": "clients",
                    "column_hints": ["email"],
                    "filter_hints": ["email is empty"],
                    "value_hints": [],
                    "is_write_operation": True,
                    "is_destructive": True,
                    "needs_discovery": False,
                    "confidence": 0.99,
                    "interpretation_note": "empty email",
                    "clarification_question": None,
                },
                {
                    "normalized_request": "Удалить дубли по email",
                    "action": "deduplicate",
                    "source_type_hint": "excel",
                    "source_name_hint": None,
                    "resource_name_hint": "clients",
                    "column_hints": ["email"],
                    "filter_hints": [],
                    "value_hints": [],
                    "is_write_operation": True,
                    "is_destructive": True,
                    "needs_discovery": False,
                    "confidence": 0.99,
                    "interpretation_note": "dedup email",
                    "clarification_question": None,
                },
                {
                    "normalized_request": "Переименовать client в client_name",
                    "action": "rename_column",
                    "source_type_hint": "excel",
                    "source_name_hint": None,
                    "resource_name_hint": "clients",
                    "column_hints": ["client"],
                    "filter_hints": [],
                    "value_hints": ["client_name"],
                    "is_write_operation": True,
                    "is_destructive": False,
                    "needs_discovery": False,
                    "confidence": 0.99,
                    "interpretation_note": "rename",
                    "clarification_question": None,
                },
                {
                    "normalized_request": "Удалить строки, где city != Новосибирск",
                    "action": "delete_rows",
                    "source_type_hint": "excel",
                    "source_name_hint": None,
                    "resource_name_hint": "clients",
                    "column_hints": ["city"],
                    "filter_hints": ["city != Новосибирск"],
                    "value_hints": ["Новосибирск"],
                    "is_write_operation": True,
                    "is_destructive": True,
                    "needs_discovery": False,
                    "confidence": 0.99,
                    "interpretation_note": "keep city",
                    "clarification_question": None,
                },
            ],
            "execution": {"copy_original": True, "output_name": None},
            "target_source_name": None,
            "rename_to": None,
            "confidence": 0.99,
            "interpretation_note": "compound",
            "clarification_question": None,
        }
        parser = self._parser(response)
        context = {
            "active_source": {
                "name": "clients_september.xlsx",
                "kind": "excel",
                "sheets": [
                    {
                        "name": "clients",
                        "columns": [
                            {"name": "client_id"},
                            {"name": "client"},
                            {"name": "email"},
                            {"name": "city"},
                        ],
                    }
                ],
            }
        }
        envelope = asyncio.run(
            parser.parse_command(
                "убери записи без email и дубли по email, переименуй client в client_name, "
                "оставь только Новосибирск и сохрани отдельно",
                context,
            )
        )
        self.assertEqual(envelope.mode, CommandMode.DATA)
        self.assertEqual(len(envelope.tasks), 4)
        self.assertEqual(envelope.tasks[1].action, Action.DEDUPLICATE)
        self.assertEqual(envelope.tasks[1].column_hints, ["email"])
        self.assertTrue(envelope.execution.copy_original)
        payload = json.loads(parser._client.calls[0]["messages"][1]["content"])
        self.assertEqual(payload["runtime_context"]["active_source"]["name"], "clients_september.xlsx")
        self.assertIn({"name": "email"}, payload["runtime_context"]["active_source"]["sheets"][0]["columns"])

    def test_semantic_python_regexes_are_removed_from_runtime_modules(self) -> None:
        root = Path(__file__).resolve().parents[1] / "app"
        intent_source = (root / "intent.py").read_text(encoding="utf-8")
        micro_source = (root / "micro_features.py").read_text(encoding="utf-8")
        planner_source = (root / "excel_planner.py").read_text(encoding="utf-8")
        self.assertNotIn("import re", intent_source)
        self.assertNotIn("_COPY_PATTERNS", micro_source)
        self.assertNotIn("_REPEAT_PATTERNS", micro_source)
        self.assertNotIn("re.search", planner_source)
        self.assertNotIn("re.compile", planner_source)


if __name__ == "__main__":
    unittest.main()
