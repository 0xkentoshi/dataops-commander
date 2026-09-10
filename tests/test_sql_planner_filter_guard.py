from __future__ import annotations

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
    ollama_stub.AsyncClient = object
    sys.modules["ollama"] = ollama_stub

from app.excel_planner import FilterOperator
from app.intent import Action, IntentDraft, SourceType
from app.sql_operations import build_sql_preview
from app.sql_planner import SqlOperationPlan, SqlPlanner, validate_sql_plan
from app.sql_service import inspect_sqlite


class FakeClient:
    def __init__(self, payloads: list[dict]) -> None:
        self.payloads = list(payloads)
        self.calls = 0

    async def chat(self, **kwargs):
        payload = self.payloads[self.calls]
        self.calls += 1
        return SimpleNamespace(
            message=SimpleNamespace(content=json.dumps(payload, ensure_ascii=False))
        )


def make_database(path: Path) -> None:
    import sqlite3

    connection = sqlite3.connect(path)
    try:
        connection.execute(
            "CREATE TABLE products ("
            "id INTEGER PRIMARY KEY, article TEXT, name TEXT, category TEXT, "
            "stock INTEGER, price REAL, status TEXT)"
        )
        connection.executemany(
            "INSERT INTO products VALUES (?, ?, ?, ?, ?, ?, ?)",
            [
                (1, "KB-001", "Клавиатура", "Периферия", 12, 3490.0, "active"),
                (2, "MS-002", "Мышь", "Периферия", 4, 1890.0, "active"),
                (3, "MN-003", "Монитор 24", "Мониторы", 7, 17990.0, "active"),
                (4, "HD-004", "Наушники", "Аудио", 0, 5290.0, "out_of_stock"),
                (5, "CM-005", "Веб-камера", "Периферия", 9, 4190.0, "active"),
                (6, "MC-006", "Микрофон", "Аудио", 2, 8990.0, "active"),
            ],
        )
        connection.commit()
    finally:
        connection.close()


def filtered_select_intent() -> IntentDraft:
    return IntentDraft(
        normalized_request="Покажи товары, у которых остаток меньше 5.",
        action=Action.SELECT,
        source_type_hint=SourceType.SQL,
        source_name_hint=None,
        resource_name_hint="products",
        column_hints=["stock"],
        filter_hints=["stock < 5"],
        value_hints=["5"],
        is_write_operation=False,
        is_destructive=False,
        needs_discovery=False,
        confidence=0.99,
        interpretation_note="test",
        clarification_question=None,
    )


class SqlPlannerFilterGuardTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self) -> None:
        import tempfile

        self.temp = tempfile.TemporaryDirectory()
        self.database = Path(self.temp.name) / "warehouse.sqlite3"
        make_database(self.database)
        self.metadata = inspect_sqlite(self.database, self.database.name)
        self.stock = next(
            column for column in self.metadata.tables[0].columns if column.name == "stock"
        )

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _bad_select(self) -> dict:
        return {
            "action": "select",
            "resolved": True,
            "table_name": "products",
            "filters": [],
            "selected_columns": [],
            "assignments": [],
            "confidence": 0.99,
            "resolution_note": "Selected products.",
            "alternative_matches": [],
            "clarification_question": None,
        }

    def _filter_recovery(self) -> dict:
        return {
            "resolved": True,
            "table_name": "products",
            "filters": [
                {
                    "column": {
                        "name": "stock",
                        "declared_type": self.stock.declared_type,
                    },
                    "operator": "lt",
                    "value": 5,
                }
            ],
            "confidence": 0.99,
            "resolution_note": "Bound the requested stock threshold to products.stock.",
            "clarification_question": None,
        }

    async def test_filtered_select_recovers_dropped_filter_with_narrow_ai_resolver(self) -> None:
        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient([self._bad_select(), self._filter_recovery()])

        plan = await planner.resolve(
            "Покажи товары, у которых остаток меньше 5.",
            filtered_select_intent(),
            self.metadata,
        )

        self.assertEqual(planner._client.calls, 2)
        self.assertEqual(len(plan.filters), 1)
        self.assertEqual(plan.filters[0].column.name, "stock")
        self.assertEqual(plan.filters[0].operator, FilterOperator.LT)
        self.assertEqual(plan.filters[0].value, 5)
        self.assertEqual(build_sql_preview(self.database, plan, self.metadata).matched_rows, 3)

    async def test_filter_recovery_inherits_exact_table_when_resolver_omits_it(self) -> None:
        recovery_without_table = self._filter_recovery()
        recovery_without_table["table_name"] = None

        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient([self._bad_select(), recovery_without_table])

        plan = await planner.resolve(
            "Покажи товары, у которых остаток меньше 5.",
            filtered_select_intent(),
            self.metadata,
        )

        self.assertEqual(planner._client.calls, 2)
        self.assertEqual(plan.table_name, "products")
        self.assertEqual(len(plan.filters), 1)
        self.assertEqual(plan.filters[0].column.name, "stock")
        self.assertEqual(plan.filters[0].operator, FilterOperator.LT)
        self.assertEqual(plan.filters[0].value, 5)
        self.assertEqual(build_sql_preview(self.database, plan, self.metadata).matched_rows, 3)

    async def test_filter_recovery_retries_its_small_schema_instead_of_broad_sql_plan(self) -> None:
        invalid_recovery = {
            "resolved": True,
            "table_name": "products",
            "filters": [],
            "confidence": 0.95,
            "resolution_note": "Missing filter by mistake.",
            "clarification_question": None,
        }
        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient(
            [self._bad_select(), invalid_recovery, self._filter_recovery()]
        )

        plan = await planner.resolve(
            "Покажи товары, у которых остаток меньше 5.",
            filtered_select_intent(),
            self.metadata,
        )

        self.assertEqual(planner._client.calls, 3)
        self.assertEqual(plan.filters[0].operator, FilterOperator.LT)
        self.assertEqual(build_sql_preview(self.database, plan, self.metadata).matched_rows, 3)

    async def test_correct_filtered_select_does_not_add_extra_llm_call(self) -> None:
        correct = self._bad_select()
        correct["filters"] = self._filter_recovery()["filters"]
        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient([correct])

        plan = await planner.resolve(
            "Покажи товары, у которых остаток меньше 5.",
            filtered_select_intent(),
            self.metadata,
        )

        self.assertEqual(planner._client.calls, 1)
        self.assertEqual(build_sql_preview(self.database, plan, self.metadata).matched_rows, 3)

    def test_validator_rejects_dropped_filter_for_select(self) -> None:
        plan = SqlOperationPlan(
            action=Action.SELECT,
            resolved=True,
            table_name="products",
            filters=[],
            selected_columns=[],
            assignments=[],
            confidence=0.99,
            resolution_note="Selected products.",
        )
        with self.assertRaisesRegex(ValueError, "dropped row filters"):
            validate_sql_plan(plan, filtered_select_intent(), self.metadata)

    def test_unfiltered_select_is_still_allowed_when_user_requested_all_rows(self) -> None:
        intent = filtered_select_intent().model_copy(update={"filter_hints": [], "value_hints": []})
        plan = SqlOperationPlan(
            action=Action.SELECT,
            resolved=True,
            table_name="products",
            filters=[],
            selected_columns=[],
            assignments=[],
            confidence=0.99,
            resolution_note="Selected all products.",
        )
        validated = validate_sql_plan(plan, intent, self.metadata)
        self.assertEqual(validated.filters, [])
        self.assertEqual(build_sql_preview(self.database, validated, self.metadata).matched_rows, 6)


if __name__ == "__main__":
    unittest.main()
