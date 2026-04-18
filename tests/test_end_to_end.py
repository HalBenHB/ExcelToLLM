from __future__ import annotations

import unittest

from excel_describer_lib.app import build_workbook_markdown
from excel_describer_lib.progress import noop_progress
from excel_describer_lib.test_workbooks import create_manual_override_workbook
from support import workspace_temp_dir


class EndToEndTests(unittest.TestCase):
    def test_manual_override_and_table_rendering_work_together(self) -> None:
        with workspace_temp_dir("end_to_end") as tmp_dir:
            workbook_path = create_manual_override_workbook(tmp_dir / "manual.xlsx")
            markdown = build_workbook_markdown(
                workbook_path,
                tab_sheets=["ManualHeader"],
                header_overrides={"ManualHeader": 2},
                progress_fn=noop_progress,
            )

        self.assertIn("## Sheet: ManualHeader", markdown)
        self.assertIn("Manual header override applied", markdown)
        self.assertIn("## Table: ManualHeader", markdown)
        self.assertIn("| Row | A | B | C | D | E |", markdown)
        self.assertIn("Marka", markdown)
        self.assertIn("Berivan Ozturk", markdown)


if __name__ == "__main__":
    unittest.main()
