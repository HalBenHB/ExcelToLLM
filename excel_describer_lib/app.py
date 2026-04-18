from __future__ import annotations

from pathlib import Path

import pandas as pd

from .analysis import describe_sheet
from .constants import MARKDOWN_STRUCTURE_NOTE, MAX_UNIQUE_DISPLAY
from .drawings import extract_sheet_floating_text
from .progress import noop_progress, progress_bar
from .prompts import prompt_file_selection, prompt_header_overrides, prompt_tabularize
from .rendering import tabularize_sheet
from .workbook_io import list_excel_files, load_workbook_safe


def build_workbook_markdown(
    file_path: Path,
    tab_sheets: list[str] | None = None,
    header_overrides: dict[str, int] | None = None,
    max_unique_display: int = MAX_UNIQUE_DISPLAY,
    progress_fn=noop_progress,
) -> str:
    xl_pd = pd.ExcelFile(file_path)
    workbook, is_xls = load_workbook_safe(file_path)

    sheet_names = xl_pd.sheet_names
    selected_tab_sheets = tab_sheets or []
    selected_header_overrides = header_overrides or {}
    floating_text_by_sheet = (
        {} if is_xls else extract_sheet_floating_text(file_path, sheet_names)
    )

    lines: list[str] = [f"# {file_path.name}", "", MARKDOWN_STRUCTURE_NOTE, ""]

    total_sheets = len(sheet_names)
    for sheet_idx, sheet in enumerate(sheet_names, 1):
        lines.append(f"## Sheet: {sheet}")
        lines.append("")

        ws = workbook[sheet] if workbook is not None else None
        description = describe_sheet(
            xl_pd,
            ws,
            sheet,
            is_xls,
            floating_text_items=floating_text_by_sheet.get(sheet, []),
            manual_header_row_idx=selected_header_overrides.get(sheet),
            max_unique_display=max_unique_display,
            progress_fn=progress_fn,
        )
        lines.append(description)
        lines.append("")

        progress_fn(sheet_idx, total_sheets, f"Sheet '{sheet}' done")

    if selected_tab_sheets:
        lines.append("---")
        lines.append("")
        total_tables = len(selected_tab_sheets)
        for table_idx, tab_sheet in enumerate(selected_tab_sheets, 1):
            progress_fn(table_idx, total_tables, f"Preparing table '{tab_sheet}'")
            lines.append(f"## Table: {tab_sheet}")
            lines.append("")
            ws = workbook[tab_sheet] if workbook is not None else None
            lines.append(
                tabularize_sheet(
                    ws,
                    xl_pd,
                    tab_sheet,
                    is_xls,
                    source_path=file_path,
                    progress_fn=progress_fn,
                ).rstrip()
            )
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def write_workbook_markdown(
    file_path: Path,
    output_path: Path | None = None,
    tab_sheets: list[str] | None = None,
    header_overrides: dict[str, int] | None = None,
    max_unique_display: int = MAX_UNIQUE_DISPLAY,
    progress_fn=noop_progress,
) -> Path:
    final_output_path = output_path or file_path.with_suffix(".md")
    markdown = build_workbook_markdown(
        file_path=file_path,
        tab_sheets=tab_sheets,
        header_overrides=header_overrides,
        max_unique_display=max_unique_display,
        progress_fn=progress_fn,
    )
    final_output_path.write_text(markdown, encoding="utf-8")
    return final_output_path


def main() -> None:
    script_dir = Path(__file__).resolve().parent.parent
    excel_files = list_excel_files(script_dir)

    if not excel_files:
        print(f"No Excel files found in {script_dir}")
        raise SystemExit(0)

    file_path = prompt_file_selection(excel_files)
    if file_path is None:
        print("  Exiting.")
        raise SystemExit(0)

    print(f"\n  Processing: {file_path.name}")

    try:
        workbook, is_xls = load_workbook_safe(file_path)
        if is_xls:
            print("  [i] Legacy .xls detected - formula introspection disabled.")

        sheet_names = pd.ExcelFile(file_path).sheet_names
        tab_sheets = prompt_tabularize(sheet_names)
        header_overrides = prompt_header_overrides(sheet_names)
        output_path = file_path.with_suffix(".md")

        print(f"\n  Writing -> {output_path.name}")
        write_workbook_markdown(
            file_path=file_path,
            output_path=output_path,
            tab_sheets=tab_sheets,
            header_overrides=header_overrides,
            max_unique_display=MAX_UNIQUE_DISPLAY,
            progress_fn=progress_bar,
        )
        print(f"\n  [ok] Written to {output_path.name}")
        del workbook
    except Exception as exc:
        print(f"\n  [x] Error processing {file_path.name}: {exc}")
        raise
