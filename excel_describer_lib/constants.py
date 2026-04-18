from __future__ import annotations

MAX_UNIQUE_DISPLAY = 70

MARKDOWN_STRUCTURE_NOTE = """## How to Read This Markdown

This file is designed for both humans and LLMs.

- Each `## Sheet: ...` section describes one worksheet.
- `Table starts at` tells you which Excel row is treated as the real header row.
- Rows shown in blockquotes above the table description are metadata, notes, filters, fast-calculation rows, or other pre-table content that appeared before the header.
- Floating textboxes and other text-bearing drawing shapes are listed separately when present.
- Each `####` subsection describes one detected column, including its Excel column letter, inferred type, null counts, and sample values.
- If formulas are present, they are shown inline using backticks.
- If `## Table: ...` sections exist near the end, they contain row-by-row Markdown table exports of the selected sheets.

When a sheet uses a manual header override, that row is trusted as the header even if automatic detection would choose a different starting row.
"""

XML_NS = {
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "office_rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

PIVOT_AGG_PREFIXES = (
    "sum of ",
    "count of ",
    "average of ",
    "avg of ",
    "max of ",
    "min of ",
    "product of ",
    "stddev of ",
)

PIVOT_FILTER_KEYWORDS = ("(multiple items)", "(all)", "(blank)")
