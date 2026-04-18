from __future__ import annotations

import re
from pathlib import Path


def prompt_file_selection(excel_files: list[Path]) -> Path | None:
    print("\n+-- Excel files found " + "-" * 40)
    for i, file_path in enumerate(excel_files, 1):
        print(f"|  [{i}] {file_path.name}")
    print("|  [0] Exit")
    print("+" + "-" * 60)
    while True:
        raw = input("  Select a file (number): ").strip()
        if raw == "0":
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(excel_files):
            return excel_files[int(raw) - 1]
        print(f"  x Enter a number between 0 and {len(excel_files)}.")


def parse_sheet_selection(raw: str, max_index: int) -> list[int] | None:
    raw = raw.strip().lower()
    if raw in {"0", ""}:
        return []
    if raw in {"a", "all"}:
        return list(range(1, max_index + 1))

    selected: list[int] = []
    seen: set[int] = set()

    for token in re.split(r"[\s,]+", raw):
        if not token:
            continue

        if "-" in token:
            parts = token.split("-", maxsplit=1)
            if len(parts) != 2 or not all(part.isdigit() for part in parts):
                return None

            start, end = map(int, parts)
            if start > end or start < 1 or end > max_index:
                return None

            for idx in range(start, end + 1):
                if idx not in seen:
                    selected.append(idx)
                    seen.add(idx)
            continue

        if not token.isdigit():
            return None

        idx = int(token)
        if idx < 1 or idx > max_index:
            return None
        if idx not in seen:
            selected.append(idx)
            seen.add(idx)

    return selected


def prompt_tabularize(sheet_names: list[str]) -> list[str]:
    print("\n+-- Tabularize sheet tables " + "-" * 33)
    for i, sheet_name in enumerate(sheet_names, 1):
        print(f"|  [{i}] {sheet_name}")
    print("|  [a] All sheets")
    print("|  [0] Skip - no tabularization")
    print("+" + "-" * 60)
    while True:
        raw = input(
            "  Select sheet(s) to render as Markdown tables "
            "(e.g. 1,3-4 / a / 0): "
        )
        selected = parse_sheet_selection(raw, len(sheet_names))
        if selected is not None:
            return [sheet_names[i - 1] for i in selected]
        print(
            f"  x Enter 0, 'a', or sheet numbers/ranges between 1 and {len(sheet_names)}."
        )


def parse_header_overrides(raw: str, sheet_names: list[str]) -> dict[str, int] | None:
    raw = raw.strip().lower()
    if raw in {"", "0"}:
        return {}

    overrides: dict[str, int] = {}
    seen_sheet_numbers: set[int] = set()

    for token in re.split(r"[\s,]+", raw):
        if not token:
            continue

        if ":" not in token:
            return None

        sheet_part, row_part = token.split(":", maxsplit=1)
        if not sheet_part.isdigit() or not row_part.isdigit():
            return None

        sheet_number = int(sheet_part)
        excel_row = int(row_part)

        if sheet_number < 1 or sheet_number > len(sheet_names) or excel_row < 1:
            return None
        if sheet_number in seen_sheet_numbers:
            return None

        seen_sheet_numbers.add(sheet_number)
        overrides[sheet_names[sheet_number - 1]] = excel_row - 1

    return overrides


def prompt_header_overrides(sheet_names: list[str]) -> dict[str, int]:
    print("\n+-- Manual header row overrides " + "-" * 28)
    for i, sheet_name in enumerate(sheet_names, 1):
        print(f"|  [{i}] {sheet_name}")
    print("|  Format: sheet_number:header_row")
    print("|  Example: 2:3,5:7")
    print("|  [0] No manual overrides")
    print("+" + "-" * 60)
    while True:
        raw = input("  Enter overrides for sheets whose header row is known: ")
        overrides = parse_header_overrides(raw, sheet_names)
        if overrides is not None:
            return overrides
        print("  x Enter values like 2:3,5:7 or 0.")
