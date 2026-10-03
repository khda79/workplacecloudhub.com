"""Compare two permission inventory scans from the same SharePoint endpoint."""

__version__ = "1.0.0"

import argparse
import csv
import html
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
from report_html import metric_card, render_report


DETAIL_FIELDS = (
    "PermissionLevels",
    "HasUniqueRoleAssignments",
    "InheritedFrom",
    "IsLimitedAccessOnly",
    "PrincipalMemberLoginNames",
    "PrincipalMemberCount",
    "PrincipalUserMemberCount",
    "PrincipalDomainGroupMemberCount",
    "PrincipalMemberLookupStatus",
)
OUTPUT_COLUMNS = (
    "Status", "ObjectScope", "ObjectUrl", "PrincipalType", "Principal",
    "ChangedFields", "PreviousPermissionLevels", "CurrentPermissionLevels",
    "PreviousMembers", "CurrentMembers", "PreviousUniquePermissions",
    "CurrentUniquePermissions", "PreviousInheritedFrom", "CurrentInheritedFrom",
)
STAMP_RE = re.compile(r"-(\d{8})-(\d{6})(?:-[0-9a-f]{32})?\.csv$", re.I)


def scan_stamp(path: Path) -> datetime:
    match = STAMP_RE.search(path.name)
    if not match:
        raise ValueError(f"Inventory filename has no scan timestamp: {path.name}")
    return datetime.strptime("".join(match.groups()), "%Y%m%d%H%M%S")


def read_inventory(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        header = handle.readline()
        if not header:
            raise ValueError(f"Inventory CSV is empty: {path}")
        delimiter = max((";", ",", "\t"), key=header.count)
        handle.seek(0)
        reader = csv.DictReader(handle, delimiter=delimiter)
        required = {"ObjectScope", "ObjectUrl", "PrincipalType", "PrincipalName", "PermissionLevels"}
        missing = sorted(required - set(reader.fieldnames or ()))
        if missing:
            raise ValueError(f"Inventory CSV lacks columns {missing}: {path}")
        rows = list(reader)
    if any(None in row for row in rows):
        raise ValueError(f"Inventory CSV contains rows with extra columns: {path}")
    return rows


def object_path(row: dict[str, str]) -> str:
    value = row.get("ObjectServerRelativeUrl") or row.get("ObjectUrl") or ""
    return (urlparse(value).path or value).rstrip("/").casefold()


def grant_key(row: dict[str, str]) -> tuple[str, ...]:
    object_url = row.get("ObjectUrl") or row.get("WebUrl") or ""
    host = urlparse(object_url).netloc.casefold()
    principal = (row.get("PrincipalLoginName") or row.get("PrincipalName") or "").strip().casefold()
    path = object_path(row)
    if not path or not principal:
        raise ValueError(f"Permission row has no object path or principal: {object_url}")
    return (
        host,
        (row.get("ObjectScope") or "").strip().casefold(),
        path,
        (row.get("ItemId") or "").strip().casefold(),
        (row.get("PrincipalType") or "").strip().casefold(),
        principal,
    )


def normalized(value: str, field: str) -> str:
    text = (value or "").strip()
    if field == "PermissionLevels":
        return "|".join(sorted(part.strip().casefold() for part in text.split("|") if part.strip()))
    if field == "PrincipalMemberLoginNames":
        return "||".join(sorted(part.strip().casefold() for part in text.split("||") if part.strip()))
    return text.casefold()


def signature(row: dict[str, str]) -> tuple[str, ...]:
    return tuple(normalized(row.get(field, ""), field) for field in DETAIL_FIELDS)


def index_grants(rows: list[dict[str, str]]) -> dict[tuple[str, ...], list[dict[str, str]]]:
    result = defaultdict(list)
    for row in rows:
        result[grant_key(row)].append(row)
    return result


def change_row(status: str, previous: dict[str, str], current: dict[str, str], changed: str) -> dict[str, str]:
    row = current or previous
    return {
        "Status": status,
        "ObjectScope": row.get("ObjectScope", ""),
        "ObjectUrl": row.get("ObjectUrl", ""),
        "PrincipalType": row.get("PrincipalType", ""),
        "Principal": row.get("PrincipalLoginName") or row.get("PrincipalName") or "",
        "ChangedFields": changed,
        "PreviousPermissionLevels": previous.get("PermissionLevels", ""),
        "CurrentPermissionLevels": current.get("PermissionLevels", ""),
        "PreviousMembers": previous.get("PrincipalMemberLoginNames", ""),
        "CurrentMembers": current.get("PrincipalMemberLoginNames", ""),
        "PreviousUniquePermissions": previous.get("HasUniqueRoleAssignments", ""),
        "CurrentUniquePermissions": current.get("HasUniqueRoleAssignments", ""),
        "PreviousInheritedFrom": previous.get("InheritedFrom", ""),
        "CurrentInheritedFrom": current.get("InheritedFrom", ""),
    }


def compare(previous_rows: list[dict[str, str]], current_rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], Counter]:
    previous = index_grants(previous_rows)
    current = index_grants(current_rows)
    result = []
    counts = Counter()
    for key in sorted(previous.keys() | current.keys()):
        old_rows = previous.get(key, [])
        new_rows = current.get(key, [])
        old = old_rows[0] if old_rows else {}
        new = new_rows[0] if new_rows else {}
        old_signatures = {signature(row) for row in old_rows}
        new_signatures = {signature(row) for row in new_rows}
        if len(old_signatures) > 1 or len(new_signatures) > 1:
            status, changed = "Ambiguous", "Conflicting grant rows in an inventory"
        elif not old_rows:
            status, changed = "Added", "Grant absent from previous scan"
        elif not new_rows:
            status, changed = "Removed", "Grant absent from current scan"
        elif old_signatures == new_signatures:
            status, changed = "Unchanged", ""
        else:
            status = "Changed"
            changed = "|".join(field for field in DETAIL_FIELDS
                               if normalized(old.get(field, ""), field) != normalized(new.get(field, ""), field))
        counts[status] += 1
        result.append(change_row(status, old, new, changed))
    return result, counts


def write_csv(path: Path, columns: tuple[str, ...], rows: list[dict[str, str]]) -> None:
    temporary = path.with_name(path.name + ".tmp")
    try:
        with temporary.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def make_report(name: str, previous: Path, current: Path, rows: list[dict[str, str]], counts: Counter) -> str:
    changed = [row for row in rows if row["Status"] != "Unchanged"]
    cards = "".join(metric_card(label, counts[label], tone) for label, tone in (
        ("Unchanged", "ok"), ("Added", "note"), ("Removed", "bad"),
        ("Changed", "note"), ("Ambiguous", "bad")))
    table_rows = "".join(
        "<tr>" + "".join(f"<td>{html.escape(str(row[column]))}</td>" for column in
                            ("Status", "ObjectScope", "ObjectUrl", "Principal", "ChangedFields")) + "</tr>"
        for row in changed[:200]
    ) or '<tr><td colspan="5" class="empty">No permission changes detected.</td></tr>'
    body = (
        '<section class="section"><h2>Selected scans</h2><dl class="context">'
        f'<dt>Previous</dt><dd>{html.escape(str(previous))}</dd>'
        f'<dt>Current</dt><dd>{html.escape(str(current))}</dd></dl></section>'
        '<section class="section"><h2>Changed grants</h2>'
        f'<p class="section-note">Showing {min(len(changed), 200)} of {len(changed)} changed or ambiguous grants. '
        'The CSV contains every grant, including unchanged grants.</p>'
        '<div class="table-scroll"><table><thead><tr><th>Status</th><th>Scope</th><th>Object URL</th>'
        '<th>Principal</th><th>Changed fields</th></tr></thead><tbody>' + table_rows + '</tbody></table></div></section>'
    )
    download = '<div class="download-bar"><strong>Complete results</strong><a class="download-primary" href="PermissionChanges.csv">Download permission changes CSV</a><a class="download-secondary" href="Summary.csv">Summary CSV</a></div>'
    return render_report(name, "Permission scan history", datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                         "Review changes" if changed else "No changes", "note" if changed else "ok",
                         cards, body, "Read-only comparison of two permission inventory CSV files",
                         download_html=download)


def run(previous: Path, current: Path, output: Path, name: str) -> Counter:
    if previous.resolve() == current.resolve():
        raise ValueError("Previous and current scans must be different files")
    if scan_stamp(previous) >= scan_stamp(current):
        raise ValueError("Previous scan timestamp must be earlier than current scan timestamp")
    previous_rows = read_inventory(previous)
    current_rows = read_inventory(current)
    rows, counts = compare(previous_rows, current_rows)
    output.mkdir(parents=True, exist_ok=True)
    write_csv(output / "PermissionChanges.csv", OUTPUT_COLUMNS, rows)
    summary = {
        "ComparisonName": name,
        "PreviousCsv": str(previous), "CurrentCsv": str(current),
        "PreviousRows": len(previous_rows), "CurrentRows": len(current_rows),
        "PreviousUniqueGrants": len(index_grants(previous_rows)),
        "CurrentUniqueGrants": len(index_grants(current_rows)),
        **{status: counts[status] for status in ("Unchanged", "Added", "Removed", "Changed", "Ambiguous")},
    }
    write_csv(output / "Summary.csv", tuple(summary), [summary])
    report = make_report(name, previous, current, rows, counts)
    temporary = output / "PermissionHistory-Report.html.tmp"
    try:
        temporary.write_text(report, encoding="utf-8")
        os.replace(temporary, output / "PermissionHistory-Report.html")
    finally:
        temporary.unlink(missing_ok=True)
    return counts


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--old-csv", type=Path, required=True)
    parser.add_argument("--new-csv", type=Path, required=True)
    parser.add_argument("--output-directory", type=Path, required=True)
    parser.add_argument("--comparison-name", required=True)
    args = parser.parse_args()
    counts = run(args.old_csv, args.new_csv, args.output_directory, args.comparison_name)
    print(f"{datetime.now():%Y-%m-%d %H:%M:%S} Permission history: "
          + "; ".join(f"{status}={counts[status]}" for status in ("Unchanged", "Added", "Removed", "Changed", "Ambiguous")))
    print(f"{datetime.now():%Y-%m-%d %H:%M:%S} Report: {args.output_directory / 'PermissionHistory-Report.html'}")


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from console_lifecycle import run_console_script

    run_console_script(main, __file__, __version__)
