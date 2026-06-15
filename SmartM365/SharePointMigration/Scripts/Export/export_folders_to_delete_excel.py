import argparse
import builtins
import csv
import shutil
import sys
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Compare"))

from export_comparison_to_excel import (
    cell_xml,
    content_types,
    detect_csv_dialect,
    doc_props_app,
    doc_props_core,
    root_rels,
    sheet_xml_end,
    sheet_xml_start,
    styles_xml,
    workbook_rels,
    workbook_xml,
)


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if args:
        args = (f"{timestamp} {args[0]}", *args[1:])
    else:
        args = (timestamp,)
    builtins.print(*args, **kwargs)


FOLDER_DELETE_COLUMNS = [
    "Action",
    "Reason",
    "TargetWebUrl",
    "TargetLibraryTitle",
    "TargetFolderServerRelativeUrl",
    "FolderName",
    "Depth",
    "ExtraFileCount",
    "ExtraFileBytes",
    "ExtraFileMB",
    "LibraryKey",
    "WebPath",
    "LibraryPath",
    "NormalizedFolderKey",
]

NUMERIC_COLUMNS = {
    "Depth",
    "ExtraFileCount",
    "ExtraFileBytes",
    "ExtraFileMB",
}


def build_delete_folder_rows(comparison_directory):
    extra_folders_path = Path(comparison_directory) / "ExtraFoldersInTarget.csv"
    if not extra_folders_path.exists():
        raise FileNotFoundError(f"ExtraFoldersInTarget.csv not found: {extra_folders_path}")

    rows = []
    with extra_folders_path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle, dialect=detect_csv_dialect(handle)):
            if row.get("TargetWebUrl", "").strip() and row.get("TargetFolderServerRelativeUrl", "").strip():
                rows.append({column: row.get(column, "") for column in FOLDER_DELETE_COLUMNS})

    rows.sort(
        key=lambda item: (
            item["TargetWebUrl"].lower(),
            -int(item["Depth"] or 0),
            item["TargetFolderServerRelativeUrl"].lower(),
        )
    )
    return rows


def write_sheet(rows, xml_path):
    with xml_path.open("w", encoding="utf-8", newline="") as handle:
        sheet_xml_start(handle, FOLDER_DELETE_COLUMNS)
        handle.write('<row r="1">')
        for column_index, column in enumerate(FOLDER_DELETE_COLUMNS, start=1):
            handle.write(cell_xml(1, column_index, column))
        handle.write("</row>\n")

        for row_index, row in enumerate(rows, start=2):
            handle.write(f'<row r="{row_index}">')
            for column_index, column in enumerate(FOLDER_DELETE_COLUMNS, start=1):
                handle.write(cell_xml(row_index, column_index, row.get(column, ""), column in NUMERIC_COLUMNS))
            handle.write("</row>\n")

        sheet_xml_end(handle, len(rows) + 1, len(FOLDER_DELETE_COLUMNS))


def export_workbook(rows, output_xlsx):
    output_xlsx = Path(output_xlsx)
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    temp_directory = output_xlsx.parent / ".excel-temp"
    if temp_directory.exists():
        shutil.rmtree(temp_directory, ignore_errors=True)
    temp_directory.mkdir(parents=True, exist_ok=True)
    sheet_path = temp_directory / f"{output_xlsx.stem}-{uuid.uuid4().hex}-sheet1.xml.tmp"

    try:
        write_sheet(rows, sheet_path)
        with zipfile.ZipFile(output_xlsx, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            archive.writestr("[Content_Types].xml", content_types(1))
            archive.writestr("_rels/.rels", root_rels())
            archive.writestr("xl/workbook.xml", workbook_xml(["FoldersToDelete"]))
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(1))
            archive.writestr("xl/styles.xml", styles_xml())
            archive.writestr("docProps/app.xml", doc_props_app(["FoldersToDelete"]))
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


def write_csv(rows, output_csv):
    output_csv = Path(output_csv)
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FOLDER_DELETE_COLUMNS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--comparison-directory", required=True)
    parser.add_argument("--output-xlsx")
    parser.add_argument("--output-csv")
    args = parser.parse_args()

    comparison_directory = Path(args.comparison_directory)
    output_xlsx = Path(args.output_xlsx) if args.output_xlsx else comparison_directory / "Folders-To-Delete.xlsx"
    output_csv = Path(args.output_csv) if args.output_csv else comparison_directory / "Folders-To-Delete.csv"

    rows = build_delete_folder_rows(comparison_directory)
    if not rows:
        for stale_path in (output_xlsx, output_csv):
            if stale_path.exists():
                stale_path.unlink()
        print("Folders-To-Delete export skipped.")
        print("Reason: no SPO folder path was found missing from the SP2019 source file paths.")
        print(f"Excel file not created: {output_xlsx}")
        return

    write_csv(rows, output_csv)
    export_workbook(rows, output_xlsx)
    print(f"Excel file created: {output_xlsx}")
    print(f"CSV file created: {output_csv}")
    print(f"Folders to delete: {len(rows)}")


if __name__ == "__main__":
    main()
