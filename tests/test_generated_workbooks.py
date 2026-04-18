from __future__ import annotations

import unittest
from pathlib import Path

import pandas as pd

from excel_describer_lib.drawings import extract_sheet_floating_text
from excel_describer_lib.test_workbooks import build_all_test_workbooks
from support import workspace_temp_dir


class GeneratedWorkbookTests(unittest.TestCase):
    def test_build_all_test_workbooks_creates_a_broad_suite(self) -> None:
        with workspace_temp_dir("generated_suite") as tmp_dir:
            workbook_paths = build_all_test_workbooks(tmp_dir)

            self.assertGreaterEqual(len(workbook_paths), 9)
            for workbook_path in workbook_paths:
                self.assertTrue(workbook_path.exists(), workbook_path.name)
                excel_file = pd.ExcelFile(workbook_path)
                self.assertGreaterEqual(len(excel_file.sheet_names), 1)

    def test_generated_textbox_workbook_contains_extractable_floating_text(self) -> None:
        with workspace_temp_dir("generated_textbox") as tmp_dir:
            workbook_paths = build_all_test_workbooks(tmp_dir)
            textbox_workbook = next(
                path for path in workbook_paths if path.name == "textbox_sheet.xlsx"
            )
            excel_file = pd.ExcelFile(textbox_workbook)

            text_by_sheet = extract_sheet_floating_text(textbox_workbook, excel_file.sheet_names)

        self.assertIn("TextboxSheet", text_by_sheet)
        self.assertEqual(len(text_by_sheet["TextboxSheet"]), 1)
        self.assertIn("Floating business rule", text_by_sheet["TextboxSheet"][0]["text"])
        self.assertEqual(text_by_sheet["TextboxSheet"][0]["anchor"], "A5")


if __name__ == "__main__":
    unittest.main()
