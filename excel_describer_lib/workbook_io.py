from __future__ import annotations

from pathlib import Path

import openpyxl


def load_workbook_safe(file_path: Path):
    if file_path.suffix.lower() == ".xls":
        return None, True
    workbook = openpyxl.load_workbook(file_path, data_only=False)
    workbook.path = str(file_path)
    return workbook, False


def list_excel_files(script_dir: Path) -> list[Path]:
    excel_files = list(script_dir.glob("*.xlsx")) + list(script_dir.glob("*.xls"))
    return sorted(excel_files, key=lambda path: path.name.lower())
