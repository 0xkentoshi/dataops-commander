from __future__ import annotations

import shutil
import sys
import tempfile
import types
import unittest
from pathlib import Path

try:
    import ollama  # noqa: F401
except (ModuleNotFoundError, ImportError):
    ollama_stub = types.ModuleType("ollama")
    ollama_stub.AsyncClient = object
    sys.modules["ollama"] = ollama_stub

from openpyxl import load_workbook

from app.excel_operations import build_operation_preview, execute_operation
from app.excel_planner import ExcelOperationPlan
from app.excel_service import inspect_workbook
from app.intent import Action


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEMO_FILE = PROJECT_ROOT / "demo_data" / "price.xlsx"
if not DEMO_FILE.is_file():
    DEMO_FILE = PROJECT_ROOT / "workspace_files" / "price.xlsx"


class ClearCellTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.excel = self.root / "price.xlsx"
        shutil.copy2(DEMO_FILE, self.excel)
        self.snapshots = self.root / "snapshots"

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def test_clear_single_matching_cell_can_include_header(self) -> None:
        metadata = inspect_workbook(self.excel, self.excel.name)
        plan = ExcelOperationPlan(
            action=Action.CLEAR_VALUES,
            resolved=True,
            sheet_name="Прайс",
            match_cell_value="Цена",
            match_all_cells=False,
            confidence=0.99,
            resolution_note="AI planner resolved exact cell content",
        )

        preview = build_operation_preview(self.excel, plan, metadata)
        self.assertEqual(preview.affected_cells, 1)
        self.assertEqual(preview.matched_cell_addresses, ["D3"])
        self.assertIn("Строки и столбцы сохранятся", preview.summary)

        execute_operation(
            self.excel,
            self.excel.name,
            "clear-header-cell",
            plan,
            preview,
            metadata,
            self.snapshots,
        )
        workbook = load_workbook(self.excel)
        try:
            sheet = workbook["Прайс"]
            self.assertIsNone(sheet["D3"].value)
            self.assertEqual(sheet["D4"].value, 3490)
            self.assertEqual(sheet["C3"].value, "Категория")
        finally:
            workbook.close()


if __name__ == "__main__":
    unittest.main()
