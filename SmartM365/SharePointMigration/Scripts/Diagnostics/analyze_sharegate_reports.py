"""Read-only ShareGate report analysis and private HTML/CSV output."""

__version__ = "1.0.1"

import argparse
import collections
import csv
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import uuid
from contextlib import contextmanager

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "Compare"))
from report_html import escape, metric_card, render_report  # noqa: E402


BASE = Path(__file__).resolve().parents[2]
CONFIG = BASE / "Config"
STATES = {"To fix", "Accepted", "Fixed"}
STATUS_ORDER = {"Success": 0, "Skipped": 1, "Warning": 2, "Error": 3}
ITEM_TYPES = {"file", "file version", "folder", "list item", "content"}
URL = re.compile(r"https?://[^\s<>\"']+", re.I)
GUID = re.compile(r"\b[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}\b", re.I)
EMAIL = re.compile(r"\b[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+\b")
PATH = re.compile(r"(?:[A-Za-z]:\\|\\\\)[^\s<>\"']+")


def atomic_text(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        temporary.write_text(value, encoding="utf-8")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def atomic_csv(path, fieldnames, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with temporary.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def merge_missing(default, current):
    if isinstance(default, dict) and isinstance(current, dict):
        result = dict(current)
        for key, value in default.items():
            result[key] = merge_missing(value, result[key]) if key in result else value
        return result
    return current


def load_config(name):
    template = CONFIG / (name + ".json.template")
    runtime = CONFIG / (name + ".json.txt")
    default = json.loads(template.read_text(encoding="utf-8-sig"))
    if runtime.exists():
        current = json.loads(runtime.read_text(encoding="utf-8-sig"))
        merged = merge_missing(default, current)
        if name.endswith("rules"):
            existing = {rule.get("Id") for rule in merged.get("Rules", [])}
            merged["Rules"] += [rule for rule in default["Rules"] if rule["Id"] not in existing]
    else:
        merged = default
    if not runtime.exists() or merged != current:
        atomic_text(runtime, json.dumps(merged, ensure_ascii=False, indent=2) + "\n")
    return merged


def status_of(value):
    value = str(value or "").strip().casefold()
    values = {"success": "Success", "successful": "Success", "réussite": "Success",
              "succès": "Success", "warning": "Warning", "avertissement": "Warning",
              "error": "Error", "failed": "Error", "erreur": "Error", "échec": "Error",
              "skipped": "Skipped", "ignoré": "Skipped", "ignore": "Skipped"}
    return values.get(value, "Unknown")


def normalize_message(value):
    value = str(value or "")
    value = re.split(r"(?:={8,}|\bVersion\s+\d+\.\d+|\n\s*at\s+[A-Za-z])", value, maxsplit=1, flags=re.I)[0]
    value = URL.sub("<url>", value)
    value = GUID.sub("<guid>", value)
    value = EMAIL.sub("<account>", value)
    value = PATH.sub("<path>", value)
    value = re.sub(r"'[^']{1,200}'", "'<value>'", value)
    value = re.sub(r"\b\d+\b", "<number>", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value[:500] or "No diagnostic message"


def resolve_columns(headers, aliases):
    available = {re.sub(r"\s+", " ", name.strip()).casefold(): name for name in headers}
    mapping = {}
    for field, names in aliases.items():
        for alias in names:
            actual = available.get(re.sub(r"\s+", " ", alias.strip()).casefold())
            if actual:
                mapping[field] = actual
                break
    if "Status" not in mapping or "SessionId" not in mapping:
        raise ValueError("The report needs mapped Status and Session ID columns.")
    return mapping


def read_csv(path):
    csv.field_size_limit(min(sys.maxsize, 2**31 - 1))
    try:
        with path.open(encoding="utf-8-sig", newline="") as stream:
            reader = csv.DictReader(stream)
            if not reader.fieldnames:
                raise ValueError(f"Empty report: {path}")
            return list(reader), list(reader.fieldnames)
    except UnicodeDecodeError:
        with path.open(encoding="cp1252", newline="") as stream:
            reader = csv.DictReader(stream)
            if not reader.fieldnames:
                raise ValueError(f"Empty report: {path}")
            return list(reader), list(reader.fieldnames)


def inspect_csv(path):
    """Return session IDs without loading report rows or changing configuration."""
    csv.field_size_limit(min(sys.maxsize, 2**31 - 1))
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            with path.open(encoding=encoding, newline="") as stream:
                reader = csv.DictReader(stream)
                if not reader.fieldnames:
                    raise ValueError("Empty report")
                names = {re.sub(r"\s+", " ", name.strip()).casefold(): name for name in reader.fieldnames}
                session_column = next((names[alias] for alias in ("session id", "id de session", "identifiant de session") if alias in names), None)
                if not session_column:
                    raise ValueError("Session ID column not found")
                sessions = {str(row.get(session_column) or "").strip() for row in reader}
                sessions.discard("")
                return sorted(sessions)
        except UnicodeDecodeError:
            if encoding == "cp1252":
                raise


def load_rules():
    catalog = load_config("sharegate-diagnostics.rules")
    compiled = []
    seen = set()
    for rule in catalog.get("Rules", []):
        if not all(rule.get(key) for key in ("Id", "Regex", "Category", "Severity", "ProbableCause", "Remediation", "ShareGateAction")):
            raise ValueError("Every diagnostic rule needs all required fields.")
        if rule["Id"] in seen:
            raise ValueError(f"Duplicate rule ID: {rule['Id']}")
        seen.add(rule["Id"])
        compiled.append((rule, re.compile(rule["Regex"])))
    return compiled


def state_path(project_root, key):
    digest = hashlib.sha256(key.encode("utf-8")).hexdigest()
    return project_root / "ShareGate" / "DiagnosticsState" / (digest + ".json.txt")


def state_for(project_root, key, default):
    path = state_path(project_root, key)
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if data.get("PatternKey") == key and data.get("State") in STATES:
            return data["State"]
    return default if default in STATES else "To fix"


@contextmanager
def state_lock(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+b") as stream:
        stream.seek(0)
        if os.name == "nt":
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_LOCK, 1)
            try:
                yield
            finally:
                stream.seek(0)
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            fcntl.flock(stream.fileno(), fcntl.LOCK_EX)
            try:
                yield
            finally:
                fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def set_state(project_root, key, value, expected=""):
    if value not in STATES:
        raise ValueError("State must be To fix, Accepted or Fixed.")
    actor = "\\".join(part for part in (os.environ.get("USERDOMAIN"), os.environ.get("USERNAME")) if part)
    record = {"PatternKey": key, "State": value, "UpdatedAtUtc": dt.datetime.now(dt.timezone.utc).isoformat(),
              "Actor": actor, "Machine": os.environ.get("COMPUTERNAME", "")}
    path = state_path(project_root, key)
    with state_lock(path.with_name(path.name + ".lock")):
        if path.exists() and expected:
            current = json.loads(path.read_text(encoding="utf-8-sig"))
            if current.get("State") != expected:
                raise ValueError("Pattern state was changed by another operator. Reanalyze before saving.")
        atomic_text(path, json.dumps(record, ensure_ascii=False, indent=2) + "\n")


def normalize_row(raw, mapping, source):
    def get(name):
        value = str(raw.get(mapping.get(name, ""), "") or "").strip()
        return "" if value.casefold() in {"nan", "none", "null"} else value
    status = status_of(get("Status"))
    if status == "Unknown":
        raise ValueError(f"Unrecognized ShareGate status '{get('Status')}' in {source}.")
    message = get("Errors") if status == "Error" else get("Warnings") if status == "Warning" else get("Messages")
    if not message:
        message = get("Errors") or get("Warnings") or get("Messages") or get("Details")
    source_id = get("SourceItemId")
    source_site = get("SourceUrl").rstrip("/").casefold()
    source_list = (get("SourceListId") or get("SourceList")).casefold()
    item_key = ""
    if get("ObjectType").casefold() in ITEM_TYPES and source_site and source_list and source_id.isdecimal() and int(source_id) > 0:
        item_key = "|".join((source_site, source_list, str(int(source_id))))
    item_name = get("ItemName") or re.split(r"[/\\]", get("SourcePath").rstrip("/\\"))[-1]
    return {"SessionId": get("SessionId"), "RowId": get("RowId"), "Timestamp": get("Timestamp"),
            "Status": status, "ObjectType": get("ObjectType"), "ItemName": item_name,
            "SourceUrl": get("SourceUrl"), "SourceList": get("SourceList"), "SourceListId": get("SourceListId"),
            "SourceItemId": source_id, "DestinationUrl": get("DestinationUrl"), "DestinationList": get("DestinationList"),
            "Message": message, "Details": get("Details"), "HelpLinks": get("HelpLinks"),
            "CopyOptions": get("CopyOptions"), "ImportStatus": get("ImportStatus"),
            "ThrottlingStatistics": get("ThrottlingStatistics"), "ItemKey": item_key,
            "InputFile": str(source), "Raw": raw}


def classify(row, rules, project_root):
    if row["Status"] not in ("Error", "Warning"):
        row.update({"PatternKey": "", "Pattern": "", "Category": "", "RuleId": "", "State": "", "Action": ""})
        return
    text = "\n".join((row["Message"], row["Details"],
                      "Import Status: " + row["ImportStatus"] if row["ImportStatus"] else "",
                      row["ThrottlingStatistics"] if row["ThrottlingStatistics"] else ""))[:6000]
    found = next((rule for rule, expression in rules if expression.search(text)), None)
    if found is None:
        found = {"Id": "UNKNOWN", "Category": "Unknown - review", "Severity": row["Status"],
                 "ProbableCause": "Cause not classified.", "Remediation": "Review the original report row.",
                 "ShareGateAction": "Review", "DefaultState": "To fix"}
    pattern = normalize_message(row["Message"] or row["Details"])
    key = hashlib.sha256((found["Id"] + "|" + pattern.casefold()).encode("utf-8")).hexdigest()
    row.update({"PatternKey": key, "Pattern": pattern, "Category": found["Category"],
                "RuleId": found["Id"], "State": state_for(project_root, key, found.get("DefaultState", "To fix")),
                "Action": found["ShareGateAction"], "ProbableCause": found["ProbableCause"],
                "Remediation": found["Remediation"]})


def summarize(rows, duplicates, conflicts, inputs, project_root):
    line_status = collections.Counter(row["Status"] for row in rows)
    issue_rows = [row for row in rows if row["Status"] in ("Error", "Warning")]
    issue_state = collections.Counter(row["State"] for row in issue_rows)
    items = collections.defaultdict(list)
    for row in rows:
        if row["ItemKey"]:
            items[row["ItemKey"]].append(row)
    item_status = collections.Counter(max((row["Status"] for row in group), key=lambda x: STATUS_ORDER.get(x, -1)) for group in items.values())
    item_states = collections.Counter()
    for group in items.values():
        states = {row["State"] for row in group if row["State"]}
        if "To fix" in states:
            item_states["To fix"] += 1
        elif "Fixed" in states:
            item_states["Fixed"] += 1
        elif "Accepted" in states:
            item_states["Accepted"] += 1
    patterns = {}
    for row in issue_rows:
        key = row["PatternKey"]
        if key not in patterns:
            patterns[key] = {"PatternKey": key, "Pattern": row["Pattern"], "Category": row["Category"],
                             "RuleId": row["RuleId"], "State": row["State"], "Status": row["Status"], "Action": row["Action"],
                             "ProbableCause": row["ProbableCause"], "Remediation": row["Remediation"],
                             "Lines": 0, "Items": set(), "HelpLink": ""}
        entry = patterns[key]
        entry["Lines"] += 1
        if STATUS_ORDER[row["Status"]] > STATUS_ORDER[entry["Status"]]:
            entry["Status"] = row["Status"]
        if row["ItemKey"]:
            entry["Items"].add(row["ItemKey"])
        if not entry["HelpLink"]:
            match = URL.search(row["HelpLinks"])
            if match:
                entry["HelpLink"] = match.group(0).rstrip(".,;)")
    pattern_rows = []
    for value in patterns.values():
        value["Items"] = len(value["Items"])
        pattern_rows.append(value)
    pattern_rows.sort(key=lambda value: (-value["Lines"], value["Category"], value["Pattern"]))
    eligible_lines = len(rows) - issue_state["Accepted"]
    eligible_items = len(items) - item_states["Accepted"]
    breakdowns = {}
    for label, field in (("ByTypeStatus", "ObjectType"), ("BySourceSiteStatus", "SourceUrl"),
                         ("BySourceListStatus", "SourceList")):
        grouped = collections.defaultdict(collections.Counter)
        for row in rows:
            grouped[row[field] or "(unknown)"][row["Status"]] += 1
        breakdowns[label] = {name: dict(counts) for name, counts in grouped.items()}
    return {"Project": project_root.name, "GeneratedAtUtc": dt.datetime.now(dt.timezone.utc).isoformat(),
            "Inputs": [str(value) for value in inputs], "Sessions": sorted({row["SessionId"] for row in rows}),
            "Lines": len(rows), "LineStatus": dict(line_status), "IssueLineState": dict(issue_state),
            "DistinctItems": len(items), "ItemStatus": dict(item_status), "IssueItemState": dict(item_states),
            "UnkeyedRows": len(rows) - sum(len(group) for group in items.values()),
            "ResidualLineRate": round(100 * issue_state["To fix"] / eligible_lines, 2) if eligible_lines else None,
            "ResidualItemRate": round(100 * item_states["To fix"] / eligible_items, 2) if eligible_items else None,
            "DuplicateRowsSuppressed": duplicates, "ConflictingDuplicateRows": conflicts,
            "ByType": dict(collections.Counter(row["ObjectType"] or "(unknown)" for row in rows)),
            "BySourceSite": dict(collections.Counter(row["SourceUrl"] or "(unknown)" for row in rows)),
            "BySourceList": dict(collections.Counter(row["SourceList"] or "(unknown)" for row in rows)),
            "ManualActionCategories": dict(collections.Counter(row["Category"] for row in rows if row.get("Action") == "Manual")),
            "Patterns": pattern_rows, **breakdowns}


def table(headers, records):
    head = "".join(f"<th>{escape(item)}</th>" for item in headers)
    body = "".join("<tr>" + "".join(f"<td>{escape(value)}</td>" for value in record) + "</tr>" for record in records)
    return f'<div class="table-scroll"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


def make_html(summary):
    status = summary["LineStatus"]
    issue = summary["IssueLineState"]
    cards = "".join((metric_card("Report lines", summary["Lines"], "ok"),
                     metric_card("Successful lines", f"{status.get('Success', 0)} ({100*status.get('Success',0)/summary['Lines']:.1f}%)", "ok"),
                     metric_card("Error / warning lines", status.get("Error", 0) + status.get("Warning", 0), "bad"),
                     metric_card("Distinct keyed items", summary["DistinctItems"], "ok"),
                     metric_card("Unkeyed lines", summary["UnkeyedRows"], "muted"),
                     metric_card("Residual issue lines", issue.get("To fix", 0), "bad"),
                     metric_card("Residual line rate", f"{summary['ResidualLineRate']}%" if summary["ResidualLineRate"] is not None else "n/a", "bad"),
                     metric_card("Residual item rate", f"{summary['ResidualItemRate']}%" if summary["ResidualItemRate"] is not None else "n/a", "bad")))
    pattern_head = "".join(f"<th>{escape(value)}</th>" for value in ("Category", "State", "Lines", "Items", "Pattern", "Action", "Help"))
    pattern_body = ""
    for pattern in summary["Patterns"]:
        link = pattern["HelpLink"]
        help_cell = f'<a href="{escape(link)}" rel="noopener noreferrer">Open help</a>' if link.startswith(("https://", "http://")) else ""
        cells = (pattern["Category"], pattern["State"], pattern["Lines"], pattern["Items"], pattern["Pattern"], pattern["Remediation"])
        pattern_body += "<tr>" + "".join(f"<td>{escape(value)}</td>" for value in cells) + f"<td>{help_cell}</td></tr>"
    pattern_table = f'<div class="table-scroll"><table><thead><tr>{pattern_head}</tr></thead><tbody>{pattern_body}</tbody></table></div>'
    def breakdown(name, values):
        records = []
        for label, counts in values.items():
            records.append((label, sum(counts.values()), counts.get("Success", 0), counts.get("Warning", 0),
                            counts.get("Error", 0), counts.get("Skipped", 0)))
        records.sort(key=lambda item: -item[1])
        return table((name, "Lines", "Success", "Warning", "Error", "Skipped"), records)
    type_table = breakdown("Object type", summary["ByTypeStatus"])
    site_table = breakdown("Source site", summary["BySourceSiteStatus"])
    list_table = breakdown("Source library or list", summary["BySourceListStatus"])
    manual_table = table(("Manual action category", "Lines"),
                         sorted(summary["ManualActionCategories"].items(), key=lambda item: (-item[1], item[0])))
    body = ("<section class='section'><h2>Interpretation</h2><p>Raw success is based on report lines. Distinct item counts require a source site, list and positive Source ID; version rows share an item. Residual rates exclude Accepted issues from the denominator and count To fix issues in the numerator. Fixed is a tracking state, not proof that a new migration run succeeded. Unknown access side remains undetermined.</p></section>"
            + f"<section class='section'><h2>Manual actions by category</h2>{manual_table}</section>"
            + f"<section class='section'><h2>Issue patterns</h2>{pattern_table}</section>"
            + f"<section class='section'><h2>Object types</h2>{type_table}</section>"
            + f"<section class='section'><h2>Source sites</h2>{site_table}</section>"
            + f"<section class='section'><h2>Source libraries and lists</h2>{list_table}</section>")
    downloads = ('<div class="download-bar"><strong>Full evidence</strong>'
                 '<a class="download-primary" href="ClassifiedRows.csv">Classified rows (CSV)</a>'
                 '<a class="download-secondary" href="PatternSummary.csv">Patterns (CSV)</a>'
                 '<a class="download-secondary" href="UnknownPatterns.csv">Unknown patterns (CSV)</a>'
                 '<a class="download-secondary" href="Remediation-Actions.csv">Manual actions (CSV)</a>'
                 '<a class="download-secondary" href="UserMapping-Candidates.csv">User mapping candidates (CSV)</a>'
                 '<a class="download-secondary" href="DestinationMatches-Review.csv">Ambiguous matches (CSV)</a></div>')
    alert = ""
    if summary["ConflictingDuplicateRows"]:
        alert = f'<div class="callout warn">{summary["ConflictingDuplicateRows"]} conflicting duplicate report rows were detected. Review input provenance before accepting this report.</div>'
    return render_report("ShareGate migration diagnostics", "Migration diagnostics", summary["GeneratedAtUtc"],
                         "Review required" if issue.get("To fix", 0) else "No outstanding classified issues", "note", cards,
                         body, "Source reports are read only; generated outputs are private migration evidence", alert, downloads,
                         "This report is based on local ShareGate export files. It does not validate migration completeness or live site state.")


def analyze(inputs, output_dir, project_root, selected_session="", source_labels=None):
    if source_labels is None:
        source_labels = inputs
    if len(source_labels) != len(inputs):
        raise ValueError("Each input needs one source label.")
    aliases = load_config("sharegate-diagnostics.columns")
    rules = load_rules()
    seen = {}
    rows = []
    duplicates = conflicts = 0
    raw_headers = []
    for path, source_label in zip(inputs, source_labels):
        data, headers = read_csv(path)
        mapping = resolve_columns(headers, aliases)
        raw_headers.extend(header for header in headers if header not in raw_headers)
        for raw in data:
            row = normalize_row(raw, mapping, source_label)
            if not row["SessionId"]:
                raise ValueError(f"A report row has no Session ID: {path}")
            if selected_session and row["SessionId"] != selected_session:
                continue
            key = (row["SessionId"], row["RowId"]) if row["RowId"] else (row["SessionId"], hashlib.sha256(json.dumps(raw, sort_keys=True).encode("utf-8")).hexdigest())
            previous = seen.get(key)
            if previous is not None:
                duplicates += 1
                if any(previous[field] != row[field] for field in ("Status", "ObjectType", "ItemName", "SourceItemId", "SourceUrl", "DestinationUrl")):
                    conflicts += 1
                continue
            seen[key] = row
            classify(row, rules, project_root)
            rows.append(row)
    if not rows:
        raise ValueError("No ShareGate report rows were found.")
    summary = summarize(rows, duplicates, conflicts, source_labels, project_root)
    output_dir.mkdir(parents=True, exist_ok=True)
    fields = ["SessionId", "RowId", "Timestamp", "Status", "ObjectType", "ItemName", "SourceUrl", "SourceList", "SourceListId", "SourceItemId", "DestinationUrl", "DestinationList", "Message", "Details", "HelpLinks", "CopyOptions", "ImportStatus", "ThrottlingStatistics", "ItemKey", "PatternKey", "Pattern", "Category", "RuleId", "State", "Action", "InputFile"]
    raw_fields = ["Raw: " + header for header in raw_headers]
    atomic_csv(output_dir / "ClassifiedRows.csv", fields + raw_fields,
               [{**{field: row.get(field, "") for field in fields}, **{"Raw: " + name: row["Raw"].get(name, "") for name in raw_headers}} for row in rows])
    pattern_fields = ("PatternKey", "Category", "RuleId", "Status", "State", "Lines", "Items", "Pattern", "ProbableCause", "Remediation", "Action", "HelpLink")
    atomic_csv(output_dir / "PatternSummary.csv", pattern_fields, summary["Patterns"])
    atomic_csv(output_dir / "UnknownPatterns.csv", pattern_fields,
               [pattern for pattern in summary["Patterns"] if pattern["RuleId"] == "UNKNOWN"])
    manual_fields = ("SessionId", "RowId", "PatternKey", "Category", "State", "SourceUrl", "SourceList",
                     "SourceItemId", "DestinationUrl", "DestinationList", "ItemName", "Remediation", "HelpLinks")
    atomic_csv(output_dir / "Remediation-Actions.csv", manual_fields,
               [row for row in rows if row.get("Action") == "Manual"])
    user_rows = {}
    destination_rows = []
    for row in rows:
        if row.get("RuleId") == "SG-INACTIVE-USER" and row["ObjectType"].casefold() == "user":
            identity = row["ItemName"]
            if identity:
                key = identity.casefold()
                user_rows.setdefault(key, {"SourceIdentity": identity, "TargetUPN": "", "FallbackAccount": "",
                                           "Decision": "", "SourceRows": 0})["SourceRows"] += 1
        elif row.get("RuleId") == "SG-DEST-MULTIPLE":
            destination_rows.append({"SessionId": row["SessionId"], "SourceUrl": row["SourceUrl"],
                                     "SourceList": row["SourceList"], "SourceItemId": row["SourceItemId"],
                                     "DestinationUrl": row["DestinationUrl"], "DestinationList": row["DestinationList"],
                                     "Decision": "", "Notes": ""})
    atomic_csv(output_dir / "UserMapping-Candidates.csv",
               ("SourceIdentity", "TargetUPN", "FallbackAccount", "Decision", "SourceRows"), user_rows.values())
    atomic_csv(output_dir / "DestinationMatches-Review.csv",
               ("SessionId", "SourceUrl", "SourceList", "SourceItemId", "DestinationUrl", "DestinationList", "Decision", "Notes"), destination_rows)
    summary["ReportPath"] = str(output_dir / "MigrationDiagnostics-Report.html")
    summary["RowsPath"] = str(output_dir / "ClassifiedRows.csv")
    summary["PatternsPath"] = str(output_dir / "PatternSummary.csv")
    atomic_text(output_dir / "Summary.json.txt", json.dumps(summary, ensure_ascii=False, indent=2) + "\n")
    atomic_text(output_dir / "MigrationDiagnostics-Report.html", make_html(summary))
    print(summary["ReportPath"])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", action="append", type=Path)
    parser.add_argument("--source-label", action="append", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--project-root", type=Path)
    parser.add_argument("--set-state", choices=sorted(STATES))
    parser.add_argument("--pattern-key")
    parser.add_argument("--expected-state", default="")
    parser.add_argument("--session", default="")
    parser.add_argument("--inspect", action="append", type=Path)
    args = parser.parse_args()
    if args.inspect:
        result = []
        for path in args.inspect:
            try:
                result.append({"Path": str(path), "Sessions": inspect_csv(path), "Error": ""})
            except (OSError, ValueError, csv.Error) as exc:
                result.append({"Path": str(path), "Sessions": [], "Error": str(exc)})
        print(json.dumps(result, ensure_ascii=False))
    elif args.set_state:
        if not args.project_root:
            parser.error("--project-root is required")
        if not args.pattern_key:
            parser.error("--pattern-key is required with --set-state")
        set_state(args.project_root, args.pattern_key, args.set_state, args.expected_state)
    else:
        if not args.project_root or not args.input or not args.output_dir:
            parser.error("--project-root, --input and --output-dir are required for analysis")
        analyze(args.input, args.output_dir, args.project_root, args.session, args.source_label)


if __name__ == "__main__":
    main()
