from __future__ import annotations

import unittest
from pathlib import Path

from app.micro_features import ExecutionOptions, normalized_output_filename


class MicroFeatureInfrastructureTests(unittest.TestCase):
    def test_execution_options_are_data_only_not_language_parser(self) -> None:
        options = ExecutionOptions(copy_original=True, output_name="Клиенты чистые")
        self.assertTrue(options.copy_original)
        self.assertEqual(options.output_name, "Клиенты чистые")

    def test_output_extension_is_preserved(self) -> None:
        self.assertEqual(
            normalized_output_filename("Клиенты чистые", Path("source.xlsx")),
            "Клиенты чистые.xlsx",
        )
        self.assertEqual(
            normalized_output_filename("warehouse.xlsx", Path("warehouse.sqlite3")),
            "warehouse.sqlite3",
        )

    def test_output_filename_is_sanitized(self) -> None:
        self.assertEqual(
            normalized_output_filename("../boss:final?.xlsx", Path("source.xlsx")),
            "boss_final_.xlsx",
        )


if __name__ == "__main__":
    unittest.main()
