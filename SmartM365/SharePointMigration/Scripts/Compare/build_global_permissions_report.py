"""Summarize the latest available permission comparison for each migration, offline."""
__version__ = "1.0.0"

import argparse
import csv
from datetime import datetime
from html import escape
import os
from pathlib import Path
import sys
from urllib.parse import quote
from uuid import uuid4

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_global_report import download_bar, roots, stamp, write_global_workbook
from report_html import metric_card, render_report


COUNT_FIELDS = {
    "SourceRows", "TargetRows", "SourceKeys", "TargetKeys", "MatchedPermissions",
    "MissingInSPO", "DisabledEntraUsersNotInSPO", "ExtraInSPO", "TargetHasMorePermissions",
    "TargetHasLessPermissions", "PermissionLevelDifferent", "SourceLimitedAccessOnlyIgnored",
    "TargetLimitedAccessOnlyIgnored", "SourceUsersNotInEntraIgnored", "SourceDuplicateKeysIgnored",
    "TargetDuplicateKeysIgnored", "EntraUserAliasesLoaded", "DisabledEntraUsersLoaded",
    "SharePointGroupMappingsLoaded",
}


def read_summary(path):
    try:
        with path.open(encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            return next(csv.DictReader(handle, dialect=dialect), None)
    except (OSError, UnicodeError, csv.Error):
        return None


def latest_summary(migration):
    root = migration / "comparisons" / "permissions"
    if not root.is_dir():
        return None
    runs = sorted((item for item in root.iterdir() if item.is_dir()),
                  key=lambda item: (stamp(item.name), item.name), reverse=True)
    for run in runs:
        path = run / "Summary.csv"
        if not path.is_file():
            continue
        row = read_summary(path)
        if row and "MatchedPermissions" in row and "SourceUniqueKeys" in row:
            return path, row
    return None


def count(row, key):
    try:
        return int(row.get(key) or 0)
    except ValueError:
        return 0


def collect(migrations_root):
    rows = []
    for migration in sorted(Path(migrations_root).iterdir()):
        if not migration.is_dir() or migration.name.startswith("_") or not (migration / "migration.config.psd1").is_file():
            continue
        latest = latest_summary(migration)
        if not latest:
            continue
        path, summary = latest
        source, destination = roots(migration / "migration.mapping.txt")
        source_keys = count(summary, "SourceUniqueKeys")
        target_keys = count(summary, "TargetUniqueKeys")
        matched = count(summary, "MatchedPermissions")
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
        difference_fields = ("MissingInSPO", "DisabledEntraUsersNotInSPO", "ExtraInSPO",
                             "TargetHasMorePermissions", "TargetHasLessPermissions", "PermissionLevelDifferent")
        differences = sum(count(summary, field) for field in difference_fields)
        validation = summary.get("ValidationStatus") or (
            "InconclusiveEmptyInventory" if not source_keys or not target_keys else
            "ReviewNeeded" if differences else "NoRelevantDifference")
        concerns = []
        if not source_keys or not target_keys:
            concerns.append("Empty inventory")
        if scan_gap is None:
            concerns.append("Scan date unknown")
        else:
            if scan_gap > 24:
                concerns.append("Scan gap >24h")
            if oldest_age > 48:
                concerns.append("Scan age >48h at comparison")
        if summary.get("ScopeWarning"):
            concerns.append("Scope warning")
        concerns.append("Filename date only in overview")
        detail = next(path.parent.glob("*-summary-*.html"), None)
        row = {
            "Migration": migration.name,
            "SuccessPercent": f"{matched / source_keys * 100:.2f}%" if source_keys and target_keys else "N/A",
            "ValidationStatus": validation,
            "EvidenceNotes": "; ".join(concerns),
            "ComparedAt": compared_at, "SourceScannedAt": source_at, "TargetScannedAt": target_at,
            "ScanGapHours": f"{scan_gap:.1f}" if scan_gap is not None else "",
            "OldestScanAgeAtCompareHours": f"{oldest_age:.1f}" if oldest_age is not None else "",
            "Source": source, "Destination": destination,
            "SourceRootPath": summary.get("SourceRootPath", ""),
            "TargetRootPath": summary.get("TargetRootPath", ""),
            "ScopeWarning": summary.get("ScopeWarning", ""),
            "SourceRows": count(summary, "SourceRows"), "TargetRows": count(summary, "TargetRows"),
            "SourceKeys": source_keys, "TargetKeys": target_keys,
            "MatchedPermissions": matched,
            **{field: summary.get(field, "") for field in (
                "MissingInSPO", "DisabledEntraUsersNotInSPO", "ExtraInSPO",
                "TargetHasMorePermissions", "TargetHasLessPermissions", "PermissionLevelDifferent",
                "SourceLimitedAccessOnlyIgnored", "TargetLimitedAccessOnlyIgnored",
                "SourceUsersNotInEntraIgnored", "SourceDuplicateKeysIgnored", "TargetDuplicateKeysIgnored",
                "EntraUserAliasesLoaded", "DisabledEntraUsersLoaded", "SharePointGroupMappingsLoaded")},
            "ComparisonReport": str(detail or ""),
        }
        rows.append(row)
    return rows


FIELDS = [
    "Migration", "SuccessPercent", "ValidationStatus", "EvidenceNotes", "ComparedAt",
    "SourceScannedAt", "TargetScannedAt", "ScanGapHours", "OldestScanAgeAtCompareHours",
    "Source", "Destination", "SourceRootPath", "TargetRootPath", "ScopeWarning",
    "SourceRows", "TargetRows", "SourceKeys", "TargetKeys", "MatchedPermissions",
    "MissingInSPO", "DisabledEntraUsersNotInSPO", "ExtraInSPO", "TargetHasMorePermissions",
    "TargetHasLessPermissions", "PermissionLevelDifferent", "SourceLimitedAccessOnlyIgnored",
    "TargetLimitedAccessOnlyIgnored", "SourceUsersNotInEntraIgnored", "SourceDuplicateKeysIgnored",
    "TargetDuplicateKeysIgnored", "EntraUserAliasesLoaded", "DisabledEntraUsersLoaded",
    "SharePointGroupMappingsLoaded", "ComparisonReport",
]
HEADINGS = [
    "Migration", "Success %", "Status", "Evidence", "Compared", "Source scan", "Target scan",
    "Gap (h)", "Oldest age (h)", "Source", "Destination", "Source root", "Target root", "Scope warning",
    "Source rows", "Target rows", "Source keys", "Target keys", "Matched", "Missing in SPO",
    "Disabled users missing", "Extra in SPO", "Target more", "Target less", "Different levels",
    "Source limited access ignored", "Target limited access ignored", "Source users not in Entra",
    "Source duplicate keys", "Target duplicate keys", "Entra aliases", "Disabled Entra users",
    "SharePoint group mappings", "Detail",
]


def build(migrations_root, output_directory):
    rows = collect(migrations_root)
    output_directory = Path(output_directory)
    output_directory.mkdir(parents=True, exist_ok=True)
    name = "Global-Permissions-Comparison-" + datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid4().hex[:8]
    csv_path = output_directory / (name + ".csv")
    xlsx_path = output_directory / (name + ".xlsx")
    html_path = output_directory / (name + ".html")
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDS, delimiter=";")
        writer.writeheader(); writer.writerows(rows)
    write_global_workbook(rows, FIELDS, HEADINGS, xlsx_path, count_fields=COUNT_FIELDS,
                          definitions=[
                              ("Success %", "Matched permissions / source unique permission keys in the compared scope."),
                              ("Scan dates", "Derived from inventory CSV filenames in this overview; verify scan receipts before migration acceptance."),
                              ("Evidence alerts", "The overview flags scan gaps above 24 hours and an oldest scan above 48 hours at comparison."),
                              ("Scope warning", "Review the source comparison report when this field is populated."),
                          ], sheet_name="Latest permissions")
    body_rows = []
    for row in rows:
        cells = []
        for field in FIELDS:
            value = row.get(field, "")
            if field == "ComparisonReport":
                link = Path(value) if value else None
                href = quote(Path(os.path.relpath(link, html_path.parent)).as_posix()) if link and link.is_file() else ""
                value = f'<a href="{escape(href, quote=True)}">Open</a>' if href else "—"
            else:
                value = escape(str(value))
            cells.append(f"<td>{value}</td>")
        body_rows.append("<tr>" + "".join(cells) + "</tr>")
    table = '<section class="section"><div class="section-heading"><h2>Latest permission comparisons</h2></div>'
    table += '<p>One row per migration uses the newest top-level permission Summary.csv. Success % = matched permissions / source unique permission keys. Scan dates come from filenames and require receipt verification. Review ignored principals and scope warnings before accepting results.</p>'
    table += '<div class="table-scroll" role="region" tabindex="0"><table><thead><tr>'
    table += "".join(f'<th scope="col">{escape(item)}</th>' for item in HEADINGS) + '</tr></thead><tbody>'
    table += "".join(body_rows) if body_rows else f'<tr><td colspan="{len(FIELDS)}">No permission comparisons available.</td></tr>'
    table += '</tbody></table></div></section>'
    cards = metric_card("Migrations with comparisons", len(rows), "ok")
    cards += metric_card("Matched permissions", sum(count(row, "MatchedPermissions") for row in rows), "ok")
    cards += metric_card("Missing in SPO", sum(count(row, "MissingInSPO") for row in rows), "bad")
    cards += metric_card("Inconclusive inventories", sum(row["ValidationStatus"] == "InconclusiveEmptyInventory" for row in rows), "note")
    document = render_report("Global permissions comparison report", "Migration portfolio",
                             datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Review scan freshness", "note",
                             cards, table, "Generated from local permission comparison summaries",
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
