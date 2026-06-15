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
    decode_sharegate_path,
    inventory_web_key,
    is_excluded_inventory_file,
    load_web_url_filter,
    normalize_sharegate_path_characters,
    strip_query_from_path,
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


DEFAULT_PATH_LIMIT = 400
DEFAULT_SEGMENT_LIMIT = 255
ISSUE_COLUMNS = [
    "Status",
    "PathLength",
    "PathLimit",
    "OverLimitBy",
    "MaxSegmentLength",
    "SegmentLimit",
    "SegmentOverLimitBy",
    "LongestSegment",
    "TargetServerRelativeUrl",
    "SourceServerRelativeUrl",
    "SourceWebUrl",
    "SourceWebTitle",
    "SourceLibraryTitle",
    "FileName",
    "SizeBytes",
    "Modified",
    "ModifiedBy",
    "SourceFileUrl",
]
NUMERIC_COLUMNS = {
    "PathLength",
    "PathLimit",
    "OverLimitBy",
    "MaxSegmentLength",
    "SegmentLimit",
    "SegmentOverLimitBy",
    "SizeBytes",
}


def normalize_source_path(value):
    if not value:
        return ""

    text = normalize_sharegate_path_characters(decode_sharegate_path(strip_query_from_path(value)))
    if not text.startswith("/"):
        text = "/" + text

    while "//" in text:
        text = text.replace("//", "/")

    return text.rstrip("/")


def build_target_server_relative_url(source_server_relative_url, target_prefix):
    source_path = normalize_source_path(source_server_relative_url)
    if not source_path:
        return ""

    clean_target_prefix = normalize_source_path(target_prefix).rstrip("/")
    if not clean_target_prefix:
        return source_path

    if source_path.lower() == clean_target_prefix.lower():
        return source_path

    if source_path.lower().startswith((clean_target_prefix + "/").lower()):
        return source_path

    return (clean_target_prefix + "/" + source_path.strip("/")).replace("//", "/")


def build_issue_rows(source_csv_path, source_web_urls_file, source_prefixes, target_prefix, path_limit, segment_limit):
    rows = []
    seen_target_paths = set()
    source_allowed_webs = load_web_url_filter(source_web_urls_file, source_prefixes)

    with source_csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle, dialect=detect_csv_dialect(handle)):
            if source_allowed_webs is not None and inventory_web_key(row, source_prefixes) not in source_allowed_webs:
                continue
            if is_excluded_inventory_file(row):
                continue

            source_server_relative_url = row.get("ServerRelativeUrl") or row.get("FileUrl") or ""
            target_server_relative_url = build_target_server_relative_url(source_server_relative_url, target_prefix)
            if not target_server_relative_url:
                continue

            target_key = target_server_relative_url.lower()
            if target_key in seen_target_paths:
                continue
            seen_target_paths.add(target_key)

            path_for_length_check = target_server_relative_url.lstrip("/")
            path_length = len(path_for_length_check)
            segments = [segment for segment in path_for_length_check.split("/") if segment]
            longest_segment = max(segments, key=len, default="")
            max_segment_length = len(longest_segment)

            status_parts = []
            if path_length > path_limit:
                status_parts.append("Path too long for SharePoint Online")
            if max_segment_length > segment_limit:
                status_parts.append("Path segment too long for SharePoint Online")

            if not status_parts:
                continue

            rows.append(
                {
                    "Status": "; ".join(status_parts),
                    "PathLength": path_length,
                    "PathLimit": path_limit,
                    "OverLimitBy": path_length - path_limit,
                    "MaxSegmentLength": max_segment_length,
                    "SegmentLimit": segment_limit,
                    "SegmentOverLimitBy": max(max_segment_length - segment_limit, 0),
                    "LongestSegment": longest_segment,
                    "TargetServerRelativeUrl": target_server_relative_url,
                    "SourceServerRelativeUrl": source_server_relative_url,
                    "SourceWebUrl": row.get("WebUrl", ""),
                    "SourceWebTitle": row.get("WebTitle", ""),
                    "SourceLibraryTitle": row.get("LibraryTitle", ""),
                    "FileName": row.get("FileName", ""),
                    "SizeBytes": row.get("SizeBytes", ""),
                    "Modified": row.get("Modified", ""),
                    "ModifiedBy": row.get("ModifiedBy", ""),
                    "SourceFileUrl": row.get("FileUrl", ""),
                }
            )

    rows.sort(
        key=lambda item: (
            -int(item["OverLimitBy"]),
            -int(item["SegmentOverLimitBy"]),
            -int(item["PathLength"]),
            item["TargetServerRelativeUrl"].lower(),
        )
    )
    return rows


def write_csv(rows, output_csv):
    output_csv = Path(output_csv)
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=ISSUE_COLUMNS, delimiter=CSV_OUTPUT_DELIMITER)
        writer.writeheader()
        writer.writerows(rows)


def write_sheet(rows, xml_path):
    with xml_path.open("w", encoding="utf-8", newline="") as handle:
        sheet_xml_start(handle, ISSUE_COLUMNS)
        handle.write('<row r="1">')
        for column_index, column in enumerate(ISSUE_COLUMNS, start=1):
            handle.write(cell_xml(1, column_index, column))
        handle.write("</row>\n")

        for row_index, row in enumerate(rows, start=2):
            handle.write(f'<row r="{row_index}">')
            for column_index, column in enumerate(ISSUE_COLUMNS, start=1):
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

        sheet_xml_end(handle, len(rows) + 1, len(ISSUE_COLUMNS))


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
            archive.writestr("xl/workbook.xml", workbook_xml(["PathLengthIssues"]))
            archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(1))
            archive.writestr("xl/styles.xml", styles_xml())
            archive.writestr("docProps/app.xml", doc_props_app(["PathLengthIssues"]))
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
    parser.add_argument("--target-prefix", required=True)
    parser.add_argument("--path-limit", type=int, default=DEFAULT_PATH_LIMIT)
    parser.add_argument("--segment-limit", type=int, default=DEFAULT_SEGMENT_LIMIT)
    args = parser.parse_args()

    if args.path_limit < 1:
        raise ValueError("--path-limit must be greater than zero.")
    if args.segment_limit < 1:
        raise ValueError("--segment-limit must be greater than zero.")

    source_csv_path = Path(args.source_csv)
    if not source_csv_path.exists():
        raise FileNotFoundError(f"Source CSV not found: {source_csv_path}")

    rows = build_issue_rows(
        source_csv_path=source_csv_path,
        source_web_urls_file=args.source_web_urls_file,
        source_prefixes=args.source_prefix,
        target_prefix=args.target_prefix,
        path_limit=args.path_limit,
        segment_limit=args.segment_limit,
    )

    output_xlsx = Path(args.output_xlsx)
    export_workbook(rows, output_xlsx)

    if args.output_csv:
        write_csv(rows, Path(args.output_csv))

    print(f"Excel file created: {output_xlsx}")
    if args.output_csv:
        print(f"CSV file created: {args.output_csv}")
    print(f"Path length issues: {len(rows)}")
    print(f"Path length limit: {args.path_limit}")
    print(f"Path segment length limit: {args.segment_limit}")


if __name__ == "__main__":
    main()
