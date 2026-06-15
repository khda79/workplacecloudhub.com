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


DELETE_COLUMNS = [
    "Action",
    "Reason",
    "TargetWebUrl",
    "TargetLibraryTitle",
    "FileName",
    "TargetFileUrl",
    "TargetServerRelativeUrl",
    "SizeBytes",
    "SizeMB",
    "Modified",
    "ModifiedBy",
    "WebTitle",
    "LibraryKey",
    "WebPath",
    "LibraryPath",
    "Key",
]

NUMERIC_COLUMNS = {
    "SizeBytes",
    "SizeMB",
}


def as_int(value):
    if value is None or str(value).strip() == "":
        return 0
    return int(str(value).strip())


def bytes_to_mb(value):
    return round(as_int(value) / 1024 / 1024, 4)


def build_delete_file_rows(extra_in_target_path):
    rows = []
    with extra_in_target_path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle, dialect=detect_csv_dialect(handle)):
            target_web_url = row.get("WebUrl", "").strip()
            server_relative_url = row.get("ServerRelativeUrl", "").strip()
            file_url = row.get("FileUrl", "").strip()
            if not target_web_url or not server_relative_url:
                raise ValueError(
                    "Cannot build delete candidate for file "
                    f"'{row.get('FileName', '')}' because WebUrl or ServerRelativeUrl is empty."
                )

            rows.append(
                {
                    "Action": "Delete SPO file permanently, then rerun ShareGate incremental copy if needed",
                    "Reason": "SPO contains this file but it does not exist in the SP2019 source",
                    "TargetWebUrl": target_web_url,
                    "TargetLibraryTitle": row.get("LibraryTitle", ""),
                    "FileName": row.get("FileName", ""),
                    "TargetFileUrl": file_url,
                    "TargetServerRelativeUrl": server_relative_url,
                    "SizeBytes": row.get("SizeBytes", ""),
                    "SizeMB": bytes_to_mb(row.get("SizeBytes")),
                    "Modified": row.get("Modified", ""),
                    "ModifiedBy": row.get("ModifiedBy", ""),
                    "WebTitle": row.get("WebTitle", ""),
                    "LibraryKey": row.get("LibraryKey", ""),
                    "WebPath": row.get("WebPath", ""),
                    "LibraryPath": row.get("LibraryPath", ""),
                    "Key": row.get("Key", ""),
                }
            )

    rows.sort(
        key=lambda item: (
            item["TargetWebUrl"].lower(),
            item["TargetLibraryTitle"].lower(),
            item["TargetServerRelativeUrl"].lower(),
        )
    )
    return rows


def write_sheet(rows, xml_path):
    with xml_path.open("w", encoding="utf-8", newline="") as handle:
        sheet_xml_start(handle, DELETE_COLUMNS)
        handle.write('<row r="1">')
        for column_index, column in enumerate(DELETE_COLUMNS, start=1):
            handle.write(cell_xml(1, column_index, column))
        handle.write("</row>\n")

        for row_index, row in enumerate(rows, start=2):
            handle.write(f'<row r="{row_index}">')
            for column_index, column in enumerate(DELETE_COLUMNS, start=1):
                handle.write(
                    cell_xml(
                        row_index,
                        column_index,
                        row.get(column, ""),
                        column in NUMERIC_COLUMNS,
                        False,
                        False,
                    )
                )
            handle.write("</row>\n")

        sheet_xml_end(handle, len(rows) + 1, len(DELETE_COLUMNS))


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
            archive.writestr("xl/workbook.xml", workbook_xml(["FilesToDelete"]))
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(1))
            archive.writestr("xl/styles.xml", styles_xml())
            archive.writestr("docProps/app.xml", doc_props_app(["FilesToDelete"]))
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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--comparison-directory", required=True)
    parser.add_argument("--output-xlsx")
    args = parser.parse_args()

    comparison_directory = Path(args.comparison_directory)
    extra_in_target_path = comparison_directory / "ExtraInTarget.csv"
    if not extra_in_target_path.exists():
        raise FileNotFoundError(f"ExtraInTarget.csv not found: {extra_in_target_path}")

    output_xlsx = Path(args.output_xlsx) if args.output_xlsx else comparison_directory / "Files-To-Delete.xlsx"
    rows = build_delete_file_rows(extra_in_target_path)
    if not rows:
        if output_xlsx.exists():
            output_xlsx.unlink()
        print("Files-To-Delete Excel export skipped.")
        print("Reason: no extra files were found in ExtraInTarget.csv.")
        print(f"Excel file not created: {output_xlsx}")
        return

    export_workbook(rows, output_xlsx)

    print(f"Excel file created: {output_xlsx}")
    print(f"Files to delete: {len(rows)}")


if __name__ == "__main__":
    main()
