import argparse
import builtins
import shutil
import sys
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Compare"))

from export_comparison_to_excel import (
    content_types,
    doc_props_app,
    doc_props_core,
    root_rels,
    styles_xml,
    workbook_rels,
    workbook_xml,
    write_sheet_from_csv,
)


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if args:
        args = (f"{timestamp} {args[0]}", *args[1:])
    else:
        args = (timestamp,)
    builtins.print(*args, **kwargs)


def export_duplicate_keys_workbook(duplicate_keys_csv, output_xlsx):
    duplicate_keys_csv = Path(duplicate_keys_csv)
    output_xlsx = Path(output_xlsx)
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)

    temp_directory = output_xlsx.parent / ".excel-temp"
    if temp_directory.exists():
        shutil.rmtree(temp_directory, ignore_errors=True)
    temp_directory.mkdir(parents=True, exist_ok=True)

    sheet_name = "DuplicateKeys"
    sheet_path = temp_directory / f"{output_xlsx.stem}-{uuid.uuid4().hex}-sheet1.xml.tmp"

    try:
        stats = write_sheet_from_csv(duplicate_keys_csv, sheet_path)
        with zipfile.ZipFile(output_xlsx, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            archive.writestr("[Content_Types].xml", content_types(1))
            archive.writestr("_rels/.rels", root_rels())
            archive.writestr("xl/workbook.xml", workbook_xml([sheet_name]))
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(1))
            archive.writestr("xl/styles.xml", styles_xml())
            archive.writestr("docProps/app.xml", doc_props_app([sheet_name]))
            archive.writestr("docProps/core.xml", doc_props_core())
            archive.write(sheet_path, "xl/worksheets/sheet1.xml")
    finally:
        try:
            sheet_path.unlink()
        except (FileNotFoundError, PermissionError):
            pass
        try:
            temp_directory.rmdir()
        except (FileNotFoundError, OSError):
            pass

    return stats


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--comparison-directory", required=True)
    parser.add_argument("--output-xlsx")
    args = parser.parse_args()

    comparison_directory = Path(args.comparison_directory)
    duplicate_keys_csv = comparison_directory / "DuplicateKeys.csv"
    if not duplicate_keys_csv.exists():
        raise FileNotFoundError(f"DuplicateKeys.csv not found: {duplicate_keys_csv}")

    output_xlsx = Path(args.output_xlsx) if args.output_xlsx else comparison_directory / "DuplicateKeys.xlsx"
    row_count, column_count, truncated = export_duplicate_keys_workbook(duplicate_keys_csv, output_xlsx)

    print(f"Excel file created: {output_xlsx}")
    print(f"DuplicateKeys: {row_count} rows, {column_count} columns")
    if truncated:
        print("WARNING: DuplicateKeys was truncated to the Excel row limit.")


if __name__ == "__main__":
    main()
