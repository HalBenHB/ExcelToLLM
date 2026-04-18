# ExcelDescriber

`excel_describer.py` turns an Excel workbook into Markdown that is easy for a human or an LLM to read. It describes every sheet, tries to find where the real table starts, summarizes each surviving column, and can optionally append sheet-level Markdown tables for deeper inspection.

The main goal is simple: give an LLM a faithful, compact picture of what the workbook looks like without forcing it to reason over a binary `.xlsx` file.

## Version

Current documented version: `v1.1`

## What v1.1 adds

- Multi-sheet tabularization.
- More flexible sheet selection when rendering raw Markdown tables.
- The same descriptive pass as before, but now you can append multiple `## Table:` sections in one run.

You can now choose:

- `0` to skip tabularization
- `a` or `all` to tabularize every sheet
- `1,3,5` for specific sheets
- `2-4,7` for ranges plus specific sheets

## What the script produces

For each workbook run, the script writes one Markdown file next to the source workbook:

- `MyWorkbook.xlsx` -> `MyWorkbook.md`

The Markdown file contains:

1. A workbook title
2. A `## Sheet: ...` section for every worksheet
3. Sheet-level structural notes such as skipped leading rows or pivot-like metadata blocks
4. Column-by-column summaries including:
   - inferred type
   - null counts
   - unique counts
   - sample values
   - numeric stats where applicable
   - formula notes for `.xlsx` / `.xlsm`
5. Optional `## Table: ...` sections with full Markdown tables for the selected sheets

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
6. Open the generated `.md` file.

## Why this project exists

LLMs are much better with structured text than with opaque Excel binaries. A raw workbook often contains:

- title blocks before the real header
- pivot filter rows
- formulas mixed with displayed values
- blank spacer rows
- wide sheets with many mostly-empty columns

This script converts that into a stable Markdown representation that preserves useful intent while staying compact enough to feed into an LLM prompt or an analysis pipeline.

## Design Decisions

### 1. `pandas` for table-shaped reading, `openpyxl` for formulas

`pandas` is good at quickly loading sheets into DataFrames for profiling and type checks. `openpyxl` is better for workbook-aware inspection such as formula discovery.

That split is deliberate:

- `pandas` gives fast structural analysis
- `openpyxl` preserves formula strings
- a second `data_only=True` workbook load gives displayed values for formula cells when cached values exist

### 2. Describe all sheets first, tabularize only when asked

Dumping every sheet as a full Markdown table can create huge outputs. The script therefore always produces summaries for every sheet, then lets you choose which raw sheets should be fully rendered.

This keeps the default output usable for LLMs while still allowing deep inspection where needed.

### 3. Header detection is heuristic, not schema-driven

Real business workbooks often begin with notes, titles, merged headers, or pivot metadata. The script uses a heuristic `find_table_start(...)` instead of assuming row 1 is the true header.

That is intentionally pragmatic:

- it works well on messy operational files
- it avoids hardcoding one workbook layout
- it accepts that some sheets are not clean tables

### 4. Markdown is the output format on purpose

Markdown is:

- easy to diff
- easy to read in Git
- easy to paste into LLM prompts
- simple to archive alongside the workbook

The goal is explainability, not perfect spreadsheet reconstruction.

### 5. Legacy `.xls` support is best-effort

Old `.xls` files are supported through `pandas`/`xlrd` for value-level analysis, but not full formula introspection. This tradeoff keeps older files usable without pretending they provide the same fidelity as modern `.xlsx`.

## How the analysis works

At a high level, each sheet goes through this pipeline:

1. Load the raw sheet with `header=None`
2. Guess the header row with `find_table_start(...)`
3. Detect pivot-like leading rows with `_looks_like_pivot(...)`
4. Re-parse the sheet using the detected header row
5. Drop fully empty rows and columns
6. Normalize unnamed headers to `Unnamed_col_<n>`
7. Summarize each remaining column
8. Optionally render selected sheets as raw Markdown tables

## Changelog

### v1.1

- Added support for tabularizing more than one sheet in a single run.
- Added `a` / `all` support to render every sheet as a Markdown table.
- Added support for mixed sheet selections such as `1,3,5` and `2-4,7`.
- Kept output format backward-friendly by continuing to place tabularized sheets after the descriptive sections.

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

If Excel has not recalculated and saved the workbook recently, the cached displayed values may be outdated or missing. In that case, you may see formula strings without reliable evaluated results.

### `.xls` is value-first only

Legacy `.xls` files do not provide the same formula inspection path in this implementation. They are still useful for descriptive analysis, but you should treat formula reporting as unavailable there.

### Header detection can be wrong on highly irregular sheets

The header row is inferred, not guaranteed. Sheets with:

- several title bands
- multi-row headers
- merged cells
- decorative spacing
- multiple unrelated tables on one sheet

can confuse the heuristic. When that happens, the generated description is still useful, but the inferred table start may need a manual sanity check.

### Wide sheets can create huge Markdown

If you choose to tabularize a very large sheet, the resulting Markdown can become enormous. That is expected. For LLM usage, it is usually better to tabularize only the sheets you truly need.

### Unicode display can depend on the terminal

The script writes UTF-8 Markdown, but terminal output can still look odd on some Windows console configurations. If the console shows garbled characters while the `.md` file looks fine in an editor, the issue is usually console encoding rather than the output file itself.

### Formula samples are representative, not exhaustive

Column summaries show sample formulas and mention when multiple distinct formulas exist. They are meant to explain the column, not to serialize every formula in every row.

### Multiple tables on one sheet are not modeled separately

The current approach assumes one dominant table-like region per sheet for the descriptive summary. A sheet containing three unrelated data islands may still be summarized as one structure.

## Output Philosophy

This project intentionally prefers:

- useful summaries over pixel-perfect spreadsheet fidelity
- LLM readability over Excel feature completeness
- pragmatic heuristics over brittle assumptions

If a workbook is messy, the output should still tell you:

- where the likely data starts
- what each column means
- whether formulas are involved
- which sheets deserve closer inspection

## When this tool works especially well

- Business workbooks with one dominant table per sheet
- Operational reports with a few leading note rows
- Pivot exports that include filter metadata above the real headers
- Workbooks you want to summarize before sending to an LLM

## When to be careful

- Highly formatted presentation sheets
- Sheets with merged, multi-row headers
- Sheets with multiple disconnected table regions
- Workbooks that depend on fresh Excel recalculation
- Extremely large sheets selected for full tabularization

## Suggested workflow with an LLM

1. Run `excel_describer.py`
2. Review the generated `.md`
3. Keep the descriptive sections for broad context
4. Tabularize only the sheets the LLM must reason over in detail
5. Paste the relevant Markdown sections into your prompt

This usually gives better results than pasting the whole workbook blindly.

## Repository Notes

Relevant files today:

- `excel_describer.py`: workbook-to-Markdown descriptor
- `excel_comparator.py`: workbook comparison report generator
- `create_test_excel.py`: tiny helper that generates a simple test workbook

## Future Ideas

- CLI arguments instead of interactive prompts
- row and column limits for table rendering
- per-sheet output files
- better multi-table detection within one worksheet
- optional JSON output for downstream tooling
- configurable header-detection strategies

## Summary

`v1.1` keeps the original spirit of the tool intact: explain a workbook in a way an LLM can actually use. The main improvement is that raw table rendering is no longer limited to one sheet, which makes the script much more practical on real multi-sheet business workbooks.
