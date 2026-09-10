from __future__ import annotations

import json
import sqlite3
import sys
import tempfile
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

from app.intent import Action, IntentDraft, SourceType
from app.sql_operations import build_sql_preview
from app.sql_planner import SqlPlanner
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


def live_update_intent() -> IntentDraft:
    return IntentDraft(
        normalized_request=(
            "Поставить значение 77 в столбце stock для строк, "
            "где category = Периферия."
        ),
        action=Action.UPDATE_ROWS,
        source_type_hint=SourceType.SQL,
        source_name_hint=None,
        resource_name_hint="products",
        column_hints=["stock", "category"],
        filter_hints=["category = Периферия"],
        value_hints=["77", "Периферия"],
        is_write_operation=True,
        is_destructive=False,
        needs_discovery=False,
        confidence=0.99,
        interpretation_note="Update stock for Peripheral products.",
        clarification_question=None,
    )


class SqlUpdateAssignmentRecoveryTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.database = Path(self.temp.name) / "warehouse.sqlite3"
        make_database(self.database)
        self.metadata = inspect_sqlite(self.database, self.database.name)
        columns = {c.name: c for c in self.metadata.tables[0].columns}
        self.stock_type = columns["stock"].declared_type
        self.category_type = columns["category"].declared_type

    def tearDown(self) -> None:
        self.temp.cleanup()

    def bad_update(self, *, keep_filter: bool = True) -> dict:
        filters = []
        if keep_filter:
            filters = [
                {
                    "column": {
                        "name": "category",
                        "declared_type": self.category_type,
                    },
                    "operator": "eq",
                    "value": "Периферия",
                }
            ]
        return {
            "action": "update_rows",
            "resolved": True,
            "table_name": "products",
            "filters": filters,
            "selected_columns": [],
            "assignments": [],
            "confidence": 0.99,
            "resolution_note": "Update matching products.",
            "alternative_matches": [],
            "clarification_question": None,
        }

    def assignment_recovery(self, *, include_table: bool = True) -> dict:
        return {
            "resolved": True,
            "table_name": "products" if include_table else None,
            "assignments": [
                {
                    "column": {
                        "name": "stock",
                        "declared_type": self.stock_type,
                    },
                    "value": 77,
                }
            ],
            "confidence": 0.99,
            "resolution_note": "Bound the requested new value to products.stock.",
            "clarification_question": None,
        }

    def filter_recovery(self) -> dict:
        return {
            "resolved": True,
            "table_name": None,
            "filters": [
                {
                    "column": {
                        "name": "category",
                        "declared_type": self.category_type,
                    },
                    "operator": "eq",
                    "value": "Периферия",
                }
            ],
            "confidence": 0.99,
            "resolution_note": "Bound the category condition.",
            "clarification_question": None,
        }

    async def test_live_update_recovers_missing_assignment(self) -> None:
        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient([self.bad_update(), self.assignment_recovery()])

        plan = await planner.resolve(
            "Поставь остаток 77 для товаров категории Периферия.",
            live_update_intent(),
            self.metadata,
        )

        self.assertEqual(planner._client.calls, 2)
        self.assertEqual(plan.table_name, "products")
        self.assertEqual(len(plan.assignments), 1)
        self.assertEqual(plan.assignments[0].column.name, "stock")
        self.assertEqual(plan.assignments[0].value, 77)
        self.assertEqual(plan.filters[0].column.name, "category")
        self.assertEqual(plan.filters[0].value, "Периферия")

        preview = build_sql_preview(self.database, plan, self.metadata)
        self.assertEqual(preview.matched_rows, 3)
        self.assertEqual(preview.affected_cells, 3)

    async def test_assignment_recovery_can_inherit_exact_table(self) -> None:
        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient(
            [self.bad_update(), self.assignment_recovery(include_table=False)]
        )

        plan = await planner.resolve(
            "Поставь остаток 77 для товаров категории Периферия.",
            live_update_intent(),
            self.metadata,
        )

        self.assertEqual(plan.table_name, "products")
        self.assertEqual(plan.assignments[0].column.name, "stock")
        self.assertEqual(plan.assignments[0].value, 77)

    async def test_same_bad_plan_can_recover_assignment_and_filter(self) -> None:
        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient(
            [
                self.bad_update(keep_filter=False),
                self.filter_recovery(),
                self.assignment_recovery(include_table=False),
            ]
        )

        plan = await planner.resolve(
            "Поставь остаток 77 для товаров категории Периферия.",
            live_update_intent(),
            self.metadata,
        )

        self.assertEqual(planner._client.calls, 3)
        self.assertEqual(plan.assignments[0].column.name, "stock")
        self.assertEqual(plan.assignments[0].value, 77)
        self.assertEqual(plan.filters[0].column.name, "category")
        self.assertEqual(plan.filters[0].value, "Периферия")
        self.assertEqual(
            build_sql_preview(self.database, plan, self.metadata).matched_rows,
            3,
        )

    async def test_update_with_assignment_but_missing_filter_recovers_filter(self) -> None:
        bad = self.bad_update(keep_filter=False)
        bad["assignments"] = self.assignment_recovery()["assignments"]

        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient([bad, self.filter_recovery()])

        plan = await planner.resolve(
            "Поставь остаток 77 для товаров категории Периферия.",
            live_update_intent(),
            self.metadata,
        )

        self.assertEqual(planner._client.calls, 2)
        self.assertEqual(plan.assignments[0].column.name, "stock")
        self.assertEqual(plan.assignments[0].value, 77)
        self.assertEqual(plan.filters[0].column.name, "category")
        self.assertEqual(plan.filters[0].value, "Периферия")
        self.assertEqual(
            build_sql_preview(self.database, plan, self.metadata).matched_rows,
            3,
        )

    async def test_correct_update_uses_only_the_broad_planner(self) -> None:
        correct = self.bad_update()
        correct["assignments"] = self.assignment_recovery()["assignments"]

        planner = SqlPlanner(
            SimpleNamespace(ollama_model="qwen3:8b", ollama_host="http://localhost:11434")
        )
        planner._client = FakeClient([correct])

        plan = await planner.resolve(
            "Поставь остаток 77 для товаров категории Периферия.",
            live_update_intent(),
            self.metadata,
        )

        self.assertEqual(planner._client.calls, 1)
        self.assertEqual(plan.assignments[0].value, 77)


if __name__ == "__main__":
    unittest.main()
