# ExcelDescriber

`excel_describer.py` turns an Excel workbook into Markdown that is easier for both humans and LLMs to read. It describes every sheet, tries to find the real table header, summarizes each surviving column, preserves important metadata above the table, and can optionally append full Markdown table exports for selected sheets.

The goal is not pixel-perfect spreadsheet reconstruction. The goal is to give an LLM a faithful, structured, compact view of what the workbook contains.

## Version

Current documented version: `v1.4`

## What v1.4 adds

- The app is now modularized into a reusable package instead of living in one large script.
- Added a broad automated workbook suite generator covering many table structures and edge cases.
- Added multiple `unittest` files that exercise prompt parsing, header detection, rendering, floating text extraction, and end-to-end workbook output.
- Added a generated test-workbook flow via `create_test_excel.py`.
- Fixed an important edge case where formula-only columns could disappear if Excel had not cached displayed values yet.

## What v1.3 adds

- Floating textbox extraction for `.xlsx` sheets when text is stored in DrawingML shapes.
- Sheet-level reporting of textbox names, approximate anchor positions, and extracted text.

## What v1.2 adds

- Manual per-sheet header-row overrides.
- A new interactive prompt where you can enter overrides like `2:3,5:7`.
- Metadata rows above a manually specified header are still preserved and described.
- Every generated Markdown file now starts with a short structure guide for LLMs.

## What v1.1 adds

- Multi-sheet tabularization.
- More flexible sheet selection when rendering raw Markdown tables.
- The same descriptive pass as before, but now you can append multiple `## Table:` sections in one run.

## What the script produces

For each workbook run, the script writes one Markdown file next to the source workbook:

- `MyWorkbook.xlsx` -> `MyWorkbook.md`

The Markdown file contains:

1. A workbook title
2. A short preamble explaining how to read the Markdown structure
3. A `## Sheet: ...` section for every worksheet
4. Sheet-level structural notes such as skipped leading rows, pivot-like metadata blocks, or manual-header notices
5. Floating textboxes or other text-bearing drawing shapes, when present on `.xlsx` sheets
6. Column-by-column summaries including:
   - inferred type
   - null counts
   - unique counts
   - sample values
   - numeric stats where applicable
   - formula notes for `.xlsx` / `.xlsm`
7. Optional `## Table: ...` sections with full Markdown tables for the selected sheets

## Project Structure

`v1.4` introduces a small package layout so the CLI, workbook analysis, and tests can all reuse the same logic.

Key files and folders:

- `excel_describer.py`: thin compatibility wrapper and CLI entry point
- `excel_describer_lib/app.py`: top-level workbook-to-Markdown orchestration
- `excel_describer_lib/analysis.py`: header detection, skipped-row reporting, and sheet summaries
- `excel_describer_lib/rendering.py`: raw Markdown table rendering
- `excel_describer_lib/drawings.py`: floating textbox extraction from `.xlsx` package parts
- `excel_describer_lib/prompts.py`: interactive menu parsing and prompting helpers
- `excel_describer_lib/workbook_io.py`: workbook discovery and loading helpers
- `excel_describer_lib/test_workbooks.py`: generated workbook suite used by tests and the generator script
- `create_test_excel.py`: creates a folder of automated scenario workbooks
- `tests/`: `unittest` coverage for parsing, analysis, rendering, generated workbooks, and end-to-end flow

## Quick Start

1. Put `excel_describer.py` in the same folder as the Excel workbook you want to inspect.
2. Install dependencies:

```bash
pip install pandas openpyxl xlrd
```

3. Run the script:

```bash
py -3 excel_describer.py
```

4. Pick a workbook from the menu.
5. Pick which sheets to tabularize, if any.
6. Enter any manual header overrides for sheets whose real header row is known.
7. Open the generated `.md` file.

## Programmatic Usage

The new package layout also makes it easier to call the describer from Python:

```python
from pathlib import Path
from excel_describer_lib.app import write_workbook_markdown

write_workbook_markdown(
    file_path=Path("report.xlsx"),
    tab_sheets=["Summary"],
    header_overrides={"Raw Export": 2},
)
```

## Automated Test Workbooks

`v1.4` adds an automated workbook generator so you can produce many scenario files quickly:

```bash
py create_test_excel.py
```

This writes a `generated_test_excels/` folder containing a broad suite such as:

- simple clean tables
- metadata rows followed by a blank separator and real header
- manual-header-override cases
- pivot-like exports
- formula-driven tables
- unnamed header cells
- empty and sparse sheets
- multi-sheet mixed workbooks
- floating-textbox cases

## Running Tests

The project now includes built-in `unittest` coverage and does not require `pytest`.

Run the full suite with:

```bash
py -m unittest discover -s tests -v
```

Current coverage areas include:

- sheet-selection parsing
- manual header override parsing
- table-start detection
- pivot-like metadata detection
- formula-driven column reporting
- unnamed-header normalization
- raw Markdown table rendering
- workbook-level Markdown generation
- floating textbox extraction
- end-to-end output with header overrides and tabularization

## Why this project exists

LLMs are much better with structured text than with opaque Excel binaries. A raw workbook often contains:

- title blocks before the real header
- floating textboxes with business rules or notes
- pivot filter rows
- formulas mixed with displayed values
- blank spacer rows
- wide sheets with many mostly-empty columns

This project converts that into stable Markdown that preserves useful intent while staying compact enough for prompts, audits, or downstream tooling.

## Design Decisions

### 1. `pandas` for table-shaped reading, `openpyxl` for workbook-aware inspection

`pandas` is useful for profiling sheet structure and data types. `openpyxl` is better for formulas and workbook-level context.

That split is deliberate:

- `pandas` gives fast structural analysis
- `openpyxl` preserves formula strings
- a second `data_only=True` workbook load gives displayed values for formulas when cached values exist

### 2. Header detection is heuristic by default, but overrideable

Real business workbooks often begin with notes, titles, merged headers, pivot metadata, or helper rows. The script uses a heuristic `find_table_start(...)` instead of assuming row 1 is the true header, but it also lets you override the header row for specific sheets when you already know the correct answer.

That balance is practical:

- it works well on messy operational files
- it avoids hardcoding one workbook layout
- it accepts that some sheets are not clean tables
- it gives you a manual escape hatch when heuristics are not enough

### 3. Floating text matters too

Some workbooks store critical context in floating textboxes instead of cells. Those objects sit above the grid, so a plain sheet read misses them entirely.

The `.xlsx` drawing-package pass adds that missing context back into the Markdown output.

### 4. Markdown is the output format on purpose

Markdown is:

- easy to diff
- easy to read in Git
- easy to paste into LLM prompts
- simple to archive alongside the workbook

The goal is explainability, not perfect spreadsheet reconstruction.

### 5. Tests are workbook-driven, not only function-driven

Table heuristics can look correct in isolation while still failing on real files. `v1.4` therefore adds generated workbook scenarios, not just direct unit tests on small helper functions.

That keeps the suite closer to the real job of the tool:

- interpreting messy spreadsheets
- preserving metadata
- not losing formulas
- not missing floating notes

## How the analysis works

At a high level, each sheet goes through this pipeline:

1. Load the raw sheet with `header=None`
2. Extract floating textboxes from `.xlsx` drawing parts when available
3. Guess the header row with `find_table_start(...)`, unless a manual override was provided
4. Detect pivot-like leading rows with `_looks_like_pivot(...)`
5. Re-parse the sheet using the detected or overridden header row
6. Preserve columns that are formula-driven even when cached displayed values are missing
7. Drop fully empty rows and non-surviving columns
8. Normalize unnamed headers to `Unnamed_col_<n>`
9. Summarize each remaining column
10. Optionally render selected sheets as raw Markdown tables

## Changelog

### v1.4

- Modularized the codebase into `excel_describer_lib`.
- Added generated workbook fixtures covering many real-world structures.
- Added multiple `unittest` modules for parsing, analysis, rendering, generation, and end-to-end flow.
- Added `create_test_excel.py` as a test-workbook generator script.
- Fixed formula-only columns so they are still reported when formulas exist but cached displayed values are missing.

### v1.3

- Added extraction of floating textbox content from `.xlsx` drawing shapes.
- Added sheet-level reporting of textbox names, approximate anchor positions, and extracted text.

### v1.2

- Added manual per-sheet header overrides using `sheet_number:header_row` input.
- Preserved descriptive reporting for metadata rows that appear above a manually forced header row.
- Added a reusable Markdown structure preamble to each generated output file.

### v1.1

- Added support for tabularizing more than one sheet in a single run.
- Added `a` / `all` support to render every sheet as a Markdown table.
- Added support for mixed sheet selections such as `1,3,5` and `2-4,7`.

### v1.0

- Initial workbook-to-Markdown description flow.
- Sheet-by-sheet descriptive summary output.
- Table-start detection for messy worksheets.
- Pivot-like leading-row detection and reporting.
- Formula-aware summaries for `.xlsx` workbooks.
- Optional single-sheet tabularization.

## Gotchas

### Formula values may be stale

For `.xlsx`, the script loads:

- one workbook with formulas preserved
- one workbook with `data_only=True`

If Excel has not recalculated and saved the workbook recently, cached displayed values may be outdated or missing. `v1.4` now keeps formula-driven columns visible even in that case, but displayed values may still be absent.

### `.xls` is value-first only

Legacy `.xls` files are still useful for descriptive analysis, but formula introspection is unavailable there.

### Header detection can still be wrong on highly irregular sheets

The header row is inferred, not guaranteed. Sheets with:

- several title bands
- multi-row headers
- merged cells
- decorative spacing
- multiple unrelated tables on one sheet

can still confuse the heuristic. Manual header overrides are the fallback for those cases.

### Floating textbox extraction is `.xlsx`-only and shape-specific

The textbox pass reads DrawingML text-bearing shapes from modern `.xlsx` packages. That means:

- it helps when the workbook stores notes in floating textboxes
- it does not apply to legacy `.xls`
- it does not guarantee extraction from every possible Excel drawing object type
- comment infrastructure and non-text graphics may still be out of scope

### Multiple tables on one sheet are still not modeled separately

The current descriptive pass assumes one dominant table-like region per sheet. A sheet containing several unrelated data islands may still be summarized as one structure.

## Suggested Workflow With an LLM

1. Run `excel_describer.py`
2. Review the generated `.md`
3. Keep the descriptive sections for broad context
4. Tabularize only the sheets the LLM must reason over in detail
5. Paste the relevant Markdown sections into your prompt

This usually works better than pasting the raw workbook blindly.

## Summary

`v1.4` moves ExcelDescriber from a handy one-off script toward a maintainable tool:

- modular code instead of one long file
- generated workbook scenarios for regression coverage
- repeatable `unittest` verification
- better confidence that messy real-world sheets are being interpreted correctly
