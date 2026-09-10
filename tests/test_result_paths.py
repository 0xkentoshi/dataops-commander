from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from app.database import DataRepository


class ResultPathTests(unittest.TestCase):
    def test_latest_completed_file_and_relocate_keep_history(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            repo = DataRepository(root / "audit.sqlite3")
            repo.initialize()
            old_path = root / "old.xlsx"
            new_path = root / "Clients clean.xlsx"
            snapshot = root / "snapshot.xlsx"
            old_path.write_bytes(b"result")
            snapshot.write_bytes(b"before")

            for operation_id in ("op1", "op2"):
                repo.create_request(operation_id, 42, "удали дубли")
                repo.save_intent(operation_id, 42, "deduplicate", "excel", {})
                repo.save_plan(operation_id, 42, {"engine": "excel"})
                repo.complete_operation(
                    operation_id,
                    42,
                    snapshot,
                    old_path,
                    "sha",
                )

            latest = repo.get_latest_completed_file_operation(42)
            self.assertIsNotNone(latest)
            self.assertEqual(latest.operation_id, "op2")

            repo.relocate_file_path(42, old_path, new_path, "op2")
            self.assertEqual(Path(repo.get_operation("op1").result_path), new_path.resolve())
            self.assertEqual(Path(repo.get_operation("op2").result_path), new_path.resolve())


if __name__ == "__main__":
    unittest.main()
