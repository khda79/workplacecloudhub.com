import argparse
import builtins
import csv
import html
import shutil
import sys
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Compare"))

from export_comparison_to_excel import (
    cell_xml,
    column_name,
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
    "TargetLibraryTitle",
    "TargetLibraryUrl",
    "TargetFiles",
    "TargetSizeBytes",
    "TargetSizeGB",
    "ExtraInTarget",
    "ExtraInTargetPercent",
    "ExtraInTargetBytes",
    "ExtraInTargetGB",
    "SourceWebUrl",
    "TargetWebUrl",
    "MissingInTarget",
    "MatchedFiles",
    "SourceLibraryTitle",
    "SourceFiles",
    "SourceSizeBytes",
    "SourceSizeGB",
    "LibraryPath",
    "Status",
]

NUMERIC_COLUMNS = {
    "TargetFiles",
    "TargetSizeBytes",
    "TargetSizeGB",
    "ExtraInTarget",
    "ExtraInTargetBytes",
    "ExtraInTargetGB",
    "MissingInTarget",
    "MatchedFiles",
    "SourceFiles",
    "SourceSizeBytes",
    "SourceSizeGB",
}
PERCENT_COLUMNS = {
    "ExtraInTargetPercent",
}


def as_int(value):
    if value is None or str(value).strip() == "":
        return 0
    return int(str(value).strip())


def bytes_to_gb(value):
    return round(as_int(value) / 1024 / 1024 / 1024, 4)


def build_target_library_url(target_web_url, web_path, library_path):
    if not target_web_url or not library_path:
        return ""

    clean_web_path = (web_path or "").strip().strip("/").lower()
    clean_library_path = library_path.strip().strip("/")
    lower_library_path = clean_library_path.lower()

    if clean_web_path and lower_library_path.startswith(clean_web_path + "/"):
        relative_library_path = clean_library_path[len(clean_web_path) + 1 :]
    else:
        relative_library_path = clean_library_path

    if not relative_library_path or "|" in relative_library_path:
        return ""

    encoded_relative_path = "/".join(quote(part) for part in relative_library_path.split("/"))
    return target_web_url.rstrip("/") + "/" + encoded_relative_path


def build_delete_rows(library_summary_path):
    rows = []
    with library_summary_path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle, dialect=detect_csv_dialect(handle)):
            # ShareGate updates changed files during the final copy; only extra SPO files trigger deletion.
            if as_int(row.get("ExtraInTarget")) <= 0:
                continue

            target_web_url = row.get("TargetWebUrl", "").strip()
            target_library_url = build_target_library_url(
                target_web_url,
                row.get("WebPath", ""),
                row.get("LibraryPath", ""),
            )
            if not target_web_url or not target_library_url:
                raise ValueError(
                    "Cannot build delete candidate for library "
                    f"'{row.get('TargetLibraryTitle', '')}' because TargetWebUrl or TargetLibraryUrl is empty."
                )

            rows.append(
                {
                    "Action": "Delete SPO library, then rerun ShareGate copy",
                    "Reason": "SPO contains files that do not exist in the SP2019 source",
                    "TargetWebUrl": target_web_url,
                    "TargetLibraryTitle": row.get("TargetLibraryTitle", ""),
                    "TargetLibraryUrl": target_library_url,
                    "TargetFiles": row.get("TargetFiles", ""),
                    "TargetSizeBytes": row.get("TargetBytes", ""),
                    "TargetSizeGB": bytes_to_gb(row.get("TargetBytes")),
                    "ExtraInTarget": row.get("ExtraInTarget", ""),
                    "ExtraInTargetPercent": row.get("ExtraInTargetPercent", ""),
                    "ExtraInTargetBytes": row.get("ExtraInTargetBytes", ""),
                    "ExtraInTargetGB": bytes_to_gb(row.get("ExtraInTargetBytes")),
                    "MissingInTarget": row.get("MissingInTarget", ""),
                    "MatchedFiles": row.get("MatchedFiles", ""),
                    "SourceWebUrl": row.get("SourceWebUrl", ""),
                    "SourceLibraryTitle": row.get("SourceLibraryTitle", ""),
                    "SourceFiles": row.get("SourceFiles", ""),
                    "SourceSizeBytes": row.get("SourceBytes", ""),
                    "SourceSizeGB": bytes_to_gb(row.get("SourceBytes")),
                    "LibraryPath": row.get("LibraryPath", ""),
                    "Status": row.get("Status", ""),
                }
            )

    rows.sort(
        key=lambda item: (
            -as_int(item["ExtraInTarget"]),
            item["TargetWebUrl"].lower(),
            item["TargetLibraryTitle"].lower(),
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
                        column in PERCENT_COLUMNS,
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
            archive.writestr("xl/workbook.xml", workbook_xml(["LibrariesToDelete"]))
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(1))
            archive.writestr("xl/styles.xml", styles_xml())
            archive.writestr("docProps/app.xml", doc_props_app(["LibrariesToDelete"]))
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
    library_summary_path = comparison_directory / "LibrarySummary.csv"
    if not library_summary_path.exists():
        raise FileNotFoundError(f"LibrarySummary.csv not found: {library_summary_path}")

    output_xlsx = Path(args.output_xlsx) if args.output_xlsx else comparison_directory / "Libraries-To-Delete.xlsx"
    rows = build_delete_rows(library_summary_path)
    if not rows:
        if output_xlsx.exists():
            output_xlsx.unlink()
        print("Libraries-To-Delete Excel export skipped.")
        print("Reason: no libraries with extra files were found in LibrarySummary.csv.")
        print(f"Excel file not created: {output_xlsx}")
        return

    export_workbook(rows, output_xlsx)

    print(f"Excel file created: {output_xlsx}")
    print(f"Libraries to delete: {len(rows)}")


if __name__ == "__main__":
    main()
