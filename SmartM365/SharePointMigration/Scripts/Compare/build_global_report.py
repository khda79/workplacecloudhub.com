"""Summarize the latest available file comparison for each migration, offline."""
__version__ = "1.0.0"

import argparse
import csv
from datetime import datetime
from html import escape
from io import StringIO
import os
from pathlib import Path
import re
import sys
from urllib.parse import quote
from uuid import uuid4
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "Export"))
from report_html import metric_card, render_report
from export_comparison_to_excel import (
    COLUMN_WIDTHS,
    cell_xml,
    column_name,
    content_types,
    doc_props_app,
    doc_props_core,
    root_rels,
    sheet_xml_end,
    sheet_xml_start,
    styles_xml,
    workbook_rels,
    workbook_xml,
)


STAMP = re.compile(r"-(\d{8})-(\d{6})(?:-[^.]+)?(?:\.csv)?$", re.I)
COUNT_FIELDS = {
    "SourceKeys", "TargetKeys", "MatchedFiles", "MissingInTarget", "ExtraInTarget",
    "ExtraFoldersInTarget", "DifferentSize", "ChangedModifiedDate", "TargetOlderThanSource",
    "ChangedVersion", "SourceFilteredRows", "TargetFilteredRows", "SourceExcludedRows", "TargetExcludedRows",
}
DATE_FIELDS = {"ComparedAt", "SourceScannedAt", "TargetScannedAt"}
DECIMAL_FIELDS = {"ScanGapHours", "OldestScanAgeAtCompareHours"}


def write_global_workbook(rows, fields, headings, xlsx_path, *, count_fields=None, decimal_fields=None,
                          definitions=None, sheet_name="Latest comparisons"):
    """Use the project's dependency-free OOXML writer and preserve Excel value types."""
    count_fields = COUNT_FIELDS if count_fields is None else count_fields
    decimal_fields = DECIMAL_FIELDS if decimal_fields is None else decimal_fields
    COLUMN_WIDTHS.update({
        "Migration": 24, "Status": 29, "Evidence": 64, "Source": 58,
        "Destination": 58, "Detail": 65, "Compared": 22, "Source scan": 22,
        "Target scan": 22, "Success %": 15, "Definition": 85,
    })
    sheet = StringIO()
    sheet_xml_start(sheet, headings)
    sheet.write('<row r="1" ht="24" customHeight="1">')
    for column, heading in enumerate(headings, start=1):
        cell = cell_xml(1, column, heading)
        sheet.write(cell.replace(f'<c r="{column_name(column)}1"', f'<c r="{column_name(column)}1" s="3"', 1))
    sheet.write('</row>\n')
    for excel_row, row in enumerate(rows, start=2):
        sheet.write(f'<row r="{excel_row}">')
        for column, field in enumerate(fields, start=1):
            value = row.get(field, "")
            if field == "SuccessPercent" and isinstance(value, str) and value.endswith("%"):
                sheet.write(cell_xml(excel_row, column, f"{float(value[:-1]) / 100:.12g}", is_percent=True))
            else:
                sheet.write(cell_xml(excel_row, column, value,
                                     is_numeric=field in count_fields or field in decimal_fields,
                                     is_date=field in DATE_FIELDS))
        sheet.write('</row>\n')
    sheet_xml_end(sheet, len(rows) + 1, len(fields))

    if definitions is None:
        definitions = [
        ("Success %", "Matched files / source unique keys in the compared scope. Extra target files are separate."),
        ("Scan dates", "Derived from inventory CSV filenames in this overview; verify scan receipts before migration acceptance."),
        ("Evidence alerts", "The overview flags scan gaps above 12 hours and an oldest scan above 24 hours at comparison."),
        ("Filtered rows", "A 100% success rate does not cover source or destination rows excluded by the scope filter."),
        ]
    note_sheet = StringIO()
    sheet_xml_start(note_sheet, ["Field", "Definition"])
    for excel_row, values in enumerate([("Field", "Definition"), *definitions], start=1):
        note_sheet.write(f'<row r="{excel_row}"' + (' ht="24" customHeight="1"' if excel_row == 1 else '') + '>')
        for column, value in enumerate(values, start=1):
            cell = cell_xml(excel_row, column, value)
            if excel_row == 1:
                cell = cell.replace(f'<c r="{column_name(column)}1"', f'<c r="{column_name(column)}1" s="3"', 1)
            note_sheet.write(cell)
        note_sheet.write('</row>\n')
    sheet_xml_end(note_sheet, len(definitions) + 1, 2)

    sheets = [sheet_name, "Definitions"]
    comparison_xml = sheet.getvalue().replace(
        "<sheetViews>", f'<dimension ref="A1:{column_name(len(fields))}{len(rows) + 1}"/>\n<sheetViews>', 1)
    definitions_xml = note_sheet.getvalue().replace(
        "<sheetViews>", f'<dimension ref="A1:B{len(definitions) + 1}"/>\n<sheetViews>', 1)
    default_styles = styles_xml().replace('<fonts count="1">', '<fonts count="2">', 1).replace(
        '</fonts>', '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>', 1).replace(
        '<fills count="1">', '<fills count="2">', 1).replace(
        '</fills>', '<fill><patternFill patternType="solid"><fgColor rgb="FF0D2747"/><bgColor indexed="64"/></patternFill></fill></fills>', 1).replace(
        '<cellXfs count="3">', '<cellXfs count="4">', 1).replace(
        '</cellXfs>', '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>', 1).replace(
        "</styleSheet>", '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        archive.writestr("[Content_Types].xml", content_types(len(sheets)))
        archive.writestr("_rels/.rels", root_rels())
        archive.writestr("xl/workbook.xml", workbook_xml(sheets))
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels(len(sheets)))
        archive.writestr("xl/styles.xml", default_styles)
        archive.writestr("docProps/app.xml", doc_props_app(sheets))
        archive.writestr("docProps/core.xml", doc_props_core())
        archive.writestr("xl/worksheets/sheet1.xml", comparison_xml)
        archive.writestr("xl/worksheets/sheet2.xml", definitions_xml)


def download_bar(xlsx_path, csv_path):
    excel_link = escape(quote(Path(xlsx_path).name), quote=True)
    csv_link = escape(quote(Path(csv_path).name), quote=True)
    return (f'<div class="download-bar" aria-label="Report downloads"><strong>Download this report</strong>'
            f'<a class="download-primary" href="{excel_link}">Open Excel report</a>'
            f'<a class="download-secondary" href="{csv_link}">Download CSV</a></div>')


def stamp(value):
    found = STAMP.search(Path(str(value).replace("\\", "/")).name)
    if not found:
        return "Unknown"
    try:
        return datetime.strptime("".join(found.groups()), "%Y%m%d%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        return "Unknown"


def roots(mapping):
    if not mapping.is_file():
        return "Unknown", "Unknown"
    source, target = [], []
    for line in mapping.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        pair = re.split(r"[\s;,]+", line, maxsplit=2)
        if len(pair) >= 2:
            source.append(pair[0]); target.append(pair[1])
    def label(items):
        items = list(dict.fromkeys(items))
        return ", ".join(items[:3]) + (f" (+{len(items)-3})" if len(items) > 3 else "") if items else "Unknown"
    return label(source), label(target)


def newest_summary(migration):
    comparison_root = migration / "comparisons" / "files"
    if not comparison_root.is_dir():
        return None
    candidates = []
    for path in comparison_root.rglob("Summary.csv"):
        when = stamp(path.parent.name)
        candidates.append((when if when != "Unknown" else "", path))
    for _, path in sorted(candidates, key=lambda entry: (entry[0], str(entry[1])), reverse=True):
        try:
            with path.open(encoding="utf-8-sig", newline="") as handle:
                row = next(csv.DictReader(handle, delimiter=";"), None)
            if row and "MatchedKeys" in row:
                return path, row
        except (OSError, csv.Error, UnicodeError):
            continue
    return None


def collect(migrations_root):
    rows = []
    for migration in sorted(migrations_root.iterdir()):
        if not migration.is_dir() or migration.name.startswith("_") or not (migration / "migration.config.psd1").is_file():
            continue
        latest = newest_summary(migration)
        if not latest:
            continue
        path, summary = latest
        source, target = roots(migration / "migration.mapping.txt")
        try:
            matched = int(summary.get("MatchedKeys") or 0)
            source_count = int(summary.get("SourceUniqueKeys") or 0)
            target_count = int(summary.get("TargetUniqueKeys") or 0)
        except ValueError:
            continue
        percent = f"{matched / source_count * 100:.2f}%" if source_count and target_count else "N/A"
        html_report = next(path.parent.glob("*-summary-*.html"), None)
        compared_at = stamp(path.parent.name)
        source_at = stamp(summary.get("SourceCsv", ""))
        target_at = stamp(summary.get("TargetCsv", ""))
        if "Unknown" not in (compared_at, source_at, target_at):
            compared_dt, source_dt, target_dt = (datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
                                                 for value in (compared_at, source_at, target_at))
            scan_gap = abs((source_dt - target_dt).total_seconds()) / 3600
            oldest_age = (compared_dt - min(source_dt, target_dt)).total_seconds() / 3600
        else:
            scan_gap = oldest_age = None
        filtered = int(summary.get("SourceFilteredRows") or 0) + int(summary.get("TargetFilteredRows") or 0)
        difference_fields = ("MissingInTarget", "ExtraInTarget", "ExtraFoldersInTarget", "DifferentSize",
                             "ChangedModifiedDate", "TargetOlderThanSource", "ChangedVersion")
        differences = sum(int(summary.get(field) or 0) for field in difference_fields)
        validation = summary.get("ValidationStatus") or (
            "InconclusiveEmptyInventory" if not source_count or not target_count else
            "ReviewNeeded" if differences else "ReviewScopeFilter" if filtered else "NoRelevantDifference")
        concerns = []
        if not source_count or not target_count:
            concerns.append("Empty inventory")
        if filtered:
            concerns.append("Filtered rows")
        if scan_gap is None:
            concerns.append("Scan date unknown")
        else:
            if scan_gap > 12:
                concerns.append("Scan gap >12h")
            if oldest_age > 24:
                concerns.append("Scan age >24h at comparison")
        concerns.append("Filename date only in overview")
        rows.append({
            "Migration": migration.name,
            "SuccessPercent": percent,
            "ValidationStatus": validation,
            "EvidenceNotes": "; ".join(concerns),
            "ComparedAt": compared_at,
            "SourceScannedAt": source_at,
            "TargetScannedAt": target_at,
            "ScanGapHours": f"{scan_gap:.1f}" if scan_gap is not None else "",
            "OldestScanAgeAtCompareHours": f"{oldest_age:.1f}" if oldest_age is not None else "",
            "Source": source, "Destination": target,
            "SourceKeys": source_count, "TargetKeys": target_count,
            "MatchedFiles": matched,
            **{key: summary.get(key, "") for key in (
                "MissingInTarget", "ExtraInTarget", "ExtraFoldersInTarget", "DifferentSize",
                "ChangedModifiedDate", "TargetOlderThanSource", "ChangedVersion",
                "SourceFilteredRows", "TargetFilteredRows", "SourceExcludedRows", "TargetExcludedRows")},
            "ComparisonReport": str(html_report or ""),
        })
    return rows


def build(migrations_root, output_directory):
    migrations_root = Path(migrations_root)
    output_directory = Path(output_directory)
    rows = collect(migrations_root)
    output_directory.mkdir(parents=True, exist_ok=True)
    name = "Global-File-Comparison-" + datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid4().hex[:8]
    csv_path = output_directory / (name + ".csv")
    html_path = output_directory / (name + ".html")
    xlsx_path = output_directory / (name + ".xlsx")
    fields = ["Migration", "SuccessPercent", "ValidationStatus", "EvidenceNotes", "ComparedAt", "SourceScannedAt", "TargetScannedAt", "ScanGapHours", "OldestScanAgeAtCompareHours",
              "Source", "Destination", "SourceKeys", "TargetKeys", "MatchedFiles", "MissingInTarget", "ExtraInTarget",
              "ExtraFoldersInTarget", "DifferentSize", "ChangedModifiedDate", "TargetOlderThanSource", "ChangedVersion",
              "SourceFilteredRows", "TargetFilteredRows", "SourceExcludedRows", "TargetExcludedRows", "ComparisonReport"]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, delimiter=";")
        writer.writeheader(); writer.writerows(rows)
    headings = ["Migration", "Success %", "Status", "Evidence", "Compared", "Source scan", "Target scan", "Gap (h)", "Oldest age (h)", "Source", "Destination",
                "Source keys", "Target keys", "Matched", "Missing", "Extra", "Extra folders", "Size", "Modified", "Target older",
                "Version", "Source filtered", "Target filtered", "Source excluded", "Target excluded", "Detail"]
    write_global_workbook(rows, fields, headings, xlsx_path)
    body_rows = []
    for row in rows:
        cells = []
        for field in fields:
            value = row.get(field, "")
            if field == "ComparisonReport":
                link = Path(value) if value else None
                href = quote(Path(os.path.relpath(link, html_path.parent)).as_posix()) if link and link.is_file() else ""
                value = f'<a href="{escape(href, quote=True)}">Open</a>' if href else "—"
            else:
                value = escape(str(value))
            cells.append(f"<td>{value}</td>")
        body_rows.append("<tr>" + "".join(cells) + "</tr>")
    table = '<section class="section"><div class="section-heading"><h2>Latest file comparisons</h2></div>'
    table += '<p>Each row uses the newest available Summary.csv by comparison timestamp. Scan dates come from CSV filenames and are not verified by a receipt in this overview. Evidence notes flag a gap above 12 hours or an oldest scan above 24 hours at comparison. Success % = matched files / source keys in scope; extra target files and filtered rows are separate.</p>'
    table += '<div class="table-scroll" role="region" tabindex="0"><table><thead><tr>'
    table += "".join(f"<th scope=\"col\">{escape(item)}</th>" for item in headings) + "</tr></thead><tbody>"
    table += "".join(body_rows) if body_rows else f'<tr><td colspan="{len(fields)}">No file comparisons available.</td></tr>'
    table += '</tbody></table></div></section>'
    cards = metric_card("Migrations with comparisons", len(rows), "ok")
    cards += metric_card("Inconclusive inventories", sum(row["ValidationStatus"] == "InconclusiveEmptyInventory" for row in rows), "note")
    document = render_report("Global file comparison report", "Migration portfolio", datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                             "Review scan freshness", "note", cards, table, "Generated from local comparison summaries",
                             download_html=download_bar(xlsx_path, csv_path))
    html_path.write_text(document, encoding="utf-8")
    return html_path


def _console_main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--migrations-root", required=True)
    parser.add_argument("--output-directory", required=True)
    args = parser.parse_args()
    print(build(args.migrations_root, args.output_directory))


if __name__ == "__main__":
    from pathlib import Path as _ScriptPath
    import sys as _script_sys
    _script_sys.path.insert(0, str(_ScriptPath(__file__).resolve().parents[1]))
    from console_lifecycle import run_console_script
    run_console_script(_console_main, __file__, __version__)
