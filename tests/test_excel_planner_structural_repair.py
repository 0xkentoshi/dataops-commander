from __future__ import annotations

import json
import sys
import types
import unittest
from pathlib import Path

try:
    import ollama  # noqa: F401
except (ModuleNotFoundError, ImportError):
    ollama_stub = types.ModuleType("ollama")
    ollama_stub.AsyncClient = object
    sys.modules["ollama"] = ollama_stub

from app.excel_planner import _parse_and_validate_plan
from app.excel_service import inspect_workbook
from app.intent import Action, IntentDraft, SourceType


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEMO_FILE = PROJECT_ROOT / "demo_data" / "price.xlsx"


class ExcelPlannerStructuralRepairTests(unittest.TestCase):
    def test_update_target_columns_plus_replacement_value_is_normalized_to_assignment(self) -> None:
        metadata = inspect_workbook(DEMO_FILE, DEMO_FILE.name)
        sheet = metadata.sheets[0]
        category = next(column for column in sheet.columns if column.header == "Категория")
        stock = next(column for column in sheet.columns if column.header == "Остаток")

        intent = IntentDraft(
            normalized_request="Поставь остаток 77 для категории Периферия",
            action=Action.UPDATE_ROWS,
            source_type_hint=SourceType.EXCEL,
            source_name_hint=None,
            resource_name_hint=sheet.name,
            column_hints=["Остаток", "Категория"],
            filter_hints=["Категория = Периферия"],
            value_hints=["77", "Периферия"],
            is_write_operation=True,
            is_destructive=False,
            needs_discovery=False,
            confidence=0.99,
            interpretation_note="test",
            clarification_question=None,
        )

        # This is the exact structural failure observed in live QA: the model
        # chose the correct UPDATE semantics but represented the assignment via
        # target_columns + replacement_value, leaving assignments empty.
        raw = json.dumps(
            {
                "action": "update_rows",
                "resolved": True,
                "sheet_name": sheet.name,
                "filters": [
                    {
                        "column": {"index": category.index, "header": category.header},
                        "operator": "eq",
                        "value": "Периферия",
                    }
                ],
                "selected_columns": [],
                "target_columns": [
                    {"index": stock.index, "header": stock.header}
                ],
                "match_cell_value": None,
                "match_all_cells": False,
                "assignments": [],
                "apply_to_all_rows": False,
                "replacement_value": 77,
                "deduplicate_columns": [],
                "new_column_name": None,
                "new_column_default": None,
                "fill_new_column": False,
                "confidence": 0.99,
                "resolution_note": "update stock",
                "alternative_matches": [],
                "clarification_question": None,
            },
            ensure_ascii=False,
        )

        plan = _parse_and_validate_plan(raw, intent, metadata)

        self.assertEqual(plan.action, Action.UPDATE_ROWS)
        self.assertEqual(len(plan.assignments), 1)
        self.assertEqual(plan.assignments[0].column.header, "Остаток")
        self.assertEqual(plan.assignments[0].value, 77)
        self.assertEqual(plan.filters[0].column.header, "Категория")
        self.assertEqual(plan.filters[0].value, "Периферия")
