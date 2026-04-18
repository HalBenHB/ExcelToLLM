from __future__ import annotations

import unittest

from excel_describer_lib.prompts import parse_header_overrides, parse_sheet_selection


class PromptParsingTests(unittest.TestCase):
    def test_parse_sheet_selection_accepts_ranges_and_individuals(self) -> None:
        self.assertEqual(parse_sheet_selection("1,3-4,6", 6), [1, 3, 4, 6])

    def test_parse_sheet_selection_supports_all_and_skip(self) -> None:
        self.assertEqual(parse_sheet_selection("all", 4), [1, 2, 3, 4])
        self.assertEqual(parse_sheet_selection("0", 4), [])

    def test_parse_sheet_selection_rejects_invalid_tokens(self) -> None:
        self.assertIsNone(parse_sheet_selection("2-x", 5))
        self.assertIsNone(parse_sheet_selection("7", 5))

    def test_parse_header_overrides_accepts_valid_mapping(self) -> None:
        sheet_names = ["Overview", "Manual", "Pivot"]
        overrides = parse_header_overrides("2:3,3:5", sheet_names)
        self.assertEqual(overrides, {"Manual": 2, "Pivot": 4})

    def test_parse_header_overrides_rejects_duplicates_and_invalid_rows(self) -> None:
        sheet_names = ["Overview", "Manual", "Pivot"]
        self.assertIsNone(parse_header_overrides("2:3,2:4", sheet_names))
        self.assertIsNone(parse_header_overrides("1:0", sheet_names))
        self.assertIsNone(parse_header_overrides("4:2", sheet_names))


if __name__ == "__main__":
    unittest.main()
