from __future__ import annotations

from pathlib import Path

from excel_describer_lib.test_workbooks import build_all_test_workbooks


def main() -> None:
    output_dir = Path(__file__).resolve().parent / "generated_test_excels"
    workbook_paths = build_all_test_workbooks(output_dir)
    print(f"Generated {len(workbook_paths)} workbook(s) in {output_dir}:")
    for workbook_path in workbook_paths:
        print(f"- {workbook_path.name}")


if __name__ == "__main__":
    main()
