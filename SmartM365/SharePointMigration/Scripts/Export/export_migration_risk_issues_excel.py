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

from compare_sp_source_target_file_inventories import (
    CSV_OUTPUT_DELIMITER,
    inventory_exclusion_reason,
    inventory_web_key,
    load_web_url_filter,
)
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


RISK_COLUMNS = [
    "RiskType",
    "Reason",
    "FileName",
    "Extension",
    "SizeBytes",
    "SourceWebUrl",
    "LibraryTitle",
    "ServerRelativeUrl",
]
NUMERIC_COLUMNS = {"SizeBytes"}


def file_extension(file_name):
    text = str(file_name or "").strip()
    if not text or "." not in text:
        return ""
    return "." + text.rsplit(".", 1)[-1].lower()


def build_risk_rows(source_csv_path, source_web_urls_file, source_prefixes):
    rows = []
    source_allowed_webs = load_web_url_filter(source_web_urls_file, source_prefixes)

    with source_csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle, dialect=detect_csv_dialect(handle)):
            if source_allowed_webs is not None and inventory_web_key(row, source_prefixes) not in source_allowed_webs:
                continue

            risk_type, reason = inventory_exclusion_reason(row)
            if not risk_type:
                continue

            file_name = row.get("FileName", "")
            rows.append(
                {
                    "RiskType": risk_type,
                    "Reason": reason,
                    "FileName": file_name,
                    "Extension": file_extension(file_name),
                    "SizeBytes": row.get("SizeBytes", ""),
                    "SourceWebUrl": row.get("WebUrl", ""),
                    "LibraryTitle": row.get("LibraryTitle", ""),
                    "ServerRelativeUrl": row.get("ServerRelativeUrl", ""),
                }
            )

    rows.sort(
        key=lambda item: (
            item["RiskType"].lower(),
            item["SourceWebUrl"].lower(),
            item["LibraryTitle"].lower(),
            item["ServerRelativeUrl"].lower(),
        )
    )
    return rows


def write_csv(rows, output_csv):
    output_csv = Path(output_csv)
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=RISK_COLUMNS, delimiter=CSV_OUTPUT_DELIMITER)
        writer.writeheader()
        writer.writerows(rows)


def write_sheet(rows, xml_path):
    with xml_path.open("w", encoding="utf-8", newline="") as handle:
        sheet_xml_start(handle, RISK_COLUMNS)
        handle.write('<row r="1">')
        for column_index, column in enumerate(RISK_COLUMNS, start=1):
            handle.write(cell_xml(1, column_index, column))
        handle.write("</row>\n")

        for row_index, row in enumerate(rows, start=2):
            handle.write(f'<row r="{row_index}">')
            for column_index, column in enumerate(RISK_COLUMNS, start=1):
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

        sheet_xml_end(handle, len(rows) + 1, len(RISK_COLUMNS))


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
            archive.writestr("xl/workbook.xml", workbook_xml(["MigrationRiskIssues"]))
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(1))
            archive.writestr("xl/styles.xml", styles_xml())
            archive.writestr("docProps/app.xml", doc_props_app(["MigrationRiskIssues"]))
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
    parser.add_argument("--source-csv", required=True)
    parser.add_argument("--output-xlsx", required=True)
    parser.add_argument("--output-csv")
    parser.add_argument("--source-web-urls-file")
    parser.add_argument("--source-prefix", action="append", default=[])
    args = parser.parse_args()

    source_csv_path = Path(args.source_csv)
    if not source_csv_path.exists():
        raise FileNotFoundError(f"Source CSV not found: {source_csv_path}")

    rows = build_risk_rows(
        source_csv_path=source_csv_path,
        source_web_urls_file=args.source_web_urls_file,
        source_prefixes=args.source_prefix,
    )

    output_xlsx = Path(args.output_xlsx)
    export_workbook(rows, output_xlsx)

    if args.output_csv:
        write_csv(rows, Path(args.output_csv))

    print(f"Excel file created: {output_xlsx}")
    if args.output_csv:
        print(f"CSV file created: {args.output_csv}")
    print(f"Migration risk issues: {len(rows)}")


if __name__ == "__main__":
    main()
