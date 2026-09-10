from __future__ import annotations

import sys
import types
import unittest

# app.main imports aiogram; if dependencies are unavailable in a minimal test
# environment, existing project test infrastructure provides them. This file
# simply exercises the user-facing helper when imports succeed.
try:
    from app.main import planning_failure_text
except ModuleNotFoundError as error:
    if error.name == "aiogram":
        raise
    raise


class PlanningErrorUiTests(unittest.TestCase):
    def test_internal_validation_details_are_not_exposed_to_user(self) -> None:
        text = planning_failure_text()
        self.assertIn("Could not prepare a safe plan", text)
        self.assertIn("No data was changed", text)
        self.assertNotIn("ValidationError", text)
        self.assertNotIn("ValueError", text)
        self.assertNotIn("pydantic", text.casefold())


if __name__ == "__main__":
    unittest.main()
