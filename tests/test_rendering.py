from __future__ import annotations

import unittest

import openpyxl
import pandas as pd

from excel_describer_lib.app import build_workbook_markdown, write_workbook_markdown
from excel_describer_lib.progress import noop_progress
from excel_describer_lib.rendering import tabularize_sheet
from excel_describer_lib.test_workbooks import (
    create_empty_and_sparse_workbook,
    create_formula_workbook,
    create_multi_sheet_workbook,
)
from support import workspace_temp_dir


class RenderingTests(unittest.TestCase):
    def test_tabularize_sheet_keeps_formula_when_cached_value_is_missing(self) -> None:
        with workspace_temp_dir("render_formula") as tmp_dir:
            workbook_path = create_formula_workbook(tmp_dir / "formula.xlsx")
            xl_pd = pd.ExcelFile(workbook_path)
            workbook = openpyxl.load_workbook(workbook_path, data_only=False)

            markdown_table = tabularize_sheet(
                workbook["Formulas"],
                xl_pd,
                "Formulas",
                False,
                source_path=workbook_path,
                progress_fn=noop_progress,
            )

        self.assertIn("`=A2*B2`", markdown_table)
        self.assertIn("| Row | A | B | C |", markdown_table)

    def test_build_workbook_markdown_includes_intro_sheet_sections_and_tables(self) -> None:
        with workspace_temp_dir("render_multi") as tmp_dir:
            workbook_path = create_multi_sheet_workbook(tmp_dir / "multi.xlsx")
            markdown = build_workbook_markdown(
                workbook_path,
                tab_sheets=["Overview", "FormulaSheet"],
                progress_fn=noop_progress,
            )

        self.assertIn("## How to Read This Markdown", markdown)
        self.assertIn("## Sheet: Overview", markdown)
        self.assertIn("## Sheet: NotesBeforeTable", markdown)
        self.assertIn("## Table: Overview", markdown)
        self.assertIn("## Table: FormulaSheet", markdown)

    def test_write_workbook_markdown_creates_output_file(self) -> None:
        with workspace_temp_dir("render_write") as tmp_dir:
            workbook_path = create_empty_and_sparse_workbook(tmp_dir / "sparse.xlsx")
            output_path = tmp_dir / "result.md"

            written_path = write_workbook_markdown(
                workbook_path,
                output_path=output_path,
                progress_fn=noop_progress,
            )

            content = written_path.read_text(encoding="utf-8")

        self.assertEqual(written_path, output_path)
        self.assertIn("## Sheet: Empty", content)
        self.assertIn("_[empty sheet grid — skipped]_", content)
        self.assertIn("## Sheet: Sparse", content)


if __name__ == "__main__":
    unittest.main()
