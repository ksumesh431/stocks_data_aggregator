"""Tests for clear destination-workbook error handling."""

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from portfolio_report.errors import WorkbookOutputError
from portfolio_report.excel_writer import ensure_output_available


class OutputHandlingTests(unittest.TestCase):
    def test_locked_workbook_has_actionable_message(self):
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "portfolio.xlsx"
            output_path.touch()

            with (
                patch.object(Path, "open", side_effect=PermissionError("locked")),
                self.assertRaises(WorkbookOutputError) as raised,
            ):
                ensure_output_available(output_path)

            message = str(raised.exception)
            self.assertIn("probably open in Excel or LibreOffice", message)
            self.assertIn("Close the workbook", message)
            self.assertIn("--output", message)
            self.assertIn("portfolio_new.xlsx", message)


if __name__ == "__main__":
    unittest.main()
