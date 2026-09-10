from __future__ import annotations

import ast
import re
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
CYRILLIC = re.compile(r"[А-Яа-яЁё]")


class EnglishInterfaceTests(unittest.TestCase):
    def test_main_user_facing_string_literals_are_english(self) -> None:
        source = (PROJECT_ROOT / "app" / "main.py").read_text(encoding="utf-8")
        tree = ast.parse(source)
        offenders: list[tuple[int, str]] = []
        for node in ast.walk(tree):
            if not isinstance(node, ast.Constant) or not isinstance(node.value, str):
                continue
            if CYRILLIC.search(node.value):
                offenders.append((getattr(node, "lineno", 0), node.value))
        self.assertEqual(offenders, [])

    def test_planners_request_english_user_facing_explanations(self) -> None:
        for relative in (
            "app/intent.py",
            "app/excel_planner.py",
            "app/sql_planner.py",
        ):
            source = (PROJECT_ROOT / relative).read_text(encoding="utf-8")
            self.assertIn("must always be in English", source, relative)

    def test_catalog_ui_labels_are_english(self) -> None:
        source = (PROJECT_ROOT / "app" / "source_catalog.py").read_text(encoding="utf-8")
        self.assertIn('SourceKind.SQL_SCRIPT: "SQL script"', source)
        self.assertIn('SourceKind.OTHER: "File"', source)
        self.assertIn('return "Local" if self.origin == SourceOrigin.LOCAL else "TG"', source)


if __name__ == "__main__":
    unittest.main()
