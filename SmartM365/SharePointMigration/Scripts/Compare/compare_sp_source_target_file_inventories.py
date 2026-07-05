import argparse
import builtins
import csv
import html
import re
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote, unquote, urlparse


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if args:
        args = (f"{timestamp} {args[0]}", *args[1:])
    else:
        args = (timestamp,)
    builtins.print(*args, **kwargs)


INVENTORY_COLUMNS = [
    "Key",
    "LibraryKey",
    "WebPath",
    "LibraryPath",
    "SiteCollectionUrl",
    "WebUrl",
    "WebTitle",
    "LibraryTitle",
    "FileName",
    "ServerRelativeUrl",
    "FileUrl",
    "SizeBytes",
    "Modified",
    "ModifiedBy",
    "Version",
    "VersionsCount",
]

DUPLICATE_KEY_COLUMNS = ["Side"] + INVENTORY_COLUMNS

MISSING_IN_TARGET_COLUMNS = INVENTORY_COLUMNS + [
    "MissingReason",
    "TargetMatchingPathCount",
    "TargetSizeBytes",
    "TargetVersions",
    "TargetVersionComparisons",
]

MISSING_DIAGNOSTIC_COLUMNS = [
    "Key",
    "DiagnosticCategory",
    "DiagnosticSignal",
    "ActionHint",
    "SourceFileName",
    "SourceSizeBytes",
    "SourceVersion",
    "SourceModified",
    "SourceWebUrl",
    "SourceLibraryTitle",
    "SourceServerRelativeUrl",
    "TargetCandidateCount",
    "TargetCandidateFileName",
    "TargetCandidateSizeBytes",
    "TargetCandidateVersion",
    "TargetCandidateModified",
    "TargetCandidateWebUrl",
    "TargetCandidateLibraryTitle",
    "TargetCandidateServerRelativeUrl",
]

DIFFERENT_SIZE_COLUMNS = [
    "Key",
    "SourceSizeBytes",
    "TargetSizeBytes",
    "DeltaBytes",
    "DeltaBytesPercent",
    "SourceWebUrl",
    "TargetWebUrl",
    "SourceLibraryTitle",
    "TargetLibraryTitle",
    "SourceServerRelativeUrl",
    "TargetServerRelativeUrl",
    "SourceModified",
    "TargetModified",
    "SourceVersion",
    "TargetVersion",
]

CHANGED_MODIFIED_DATE_COLUMNS = [
    "Key",
    "SourceModified",
    "TargetModified",
    "SourceModifiedNormalizedUtc",
    "TargetModifiedNormalizedUtc",
    "DeltaModifiedMinutes",
    "ModifiedDirection",
    "SourceSizeBytes",
    "TargetSizeBytes",
    "SourceVersion",
    "TargetVersion",
    "SourceWebUrl",
    "TargetWebUrl",
    "SourceLibraryTitle",
    "TargetLibraryTitle",
    "SourceServerRelativeUrl",
    "TargetServerRelativeUrl",
]

TARGET_OLDER_THAN_SOURCE_COLUMNS = [
    "Key",
    "SourceModified",
    "TargetModified",
    "SourceModifiedNormalizedUtc",
    "TargetModifiedNormalizedUtc",
    "TargetOlderByMinutes",
    "SourceSizeBytes",
    "TargetSizeBytes",
    "SourceVersion",
    "TargetVersion",
    "SourceWebUrl",
    "TargetWebUrl",
    "SourceLibraryTitle",
    "TargetLibraryTitle",
    "SourceServerRelativeUrl",
    "TargetServerRelativeUrl",
]

CHANGED_VERSION_COLUMNS = [
    "Key",
    "SourceVersion",
    "TargetVersion",
    "VersionComparison",
    "SourceModified",
    "TargetModified",
    "SourceSizeBytes",
    "TargetSizeBytes",
    "SourceWebUrl",
    "TargetWebUrl",
    "SourceLibraryTitle",
    "TargetLibraryTitle",
    "SourceServerRelativeUrl",
    "TargetServerRelativeUrl",
]

EXTRA_FOLDER_COLUMNS = [
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

LIBRARY_SUMMARY_COLUMNS = [
    "Status",
    "SourceFiles",
    "TargetFiles",
    "MatchedFiles",
    "MatchedFilesPercent",
    "MissingInTarget",
    "MissingInTargetPercent",
    "ExtraInTarget",
    "ExtraInTargetPercent",
    "ExtraInTargetBytes",
    "SourceWebUrl",
    "TargetWebUrl",
    "LibraryKey",
    "WebPath",
    "LibraryPath",
    "SourceWebTitle",
    "TargetWebTitle",
    "SourceLibraryTitle",
    "TargetLibraryTitle",
    "DifferentSize",
    "DifferentSizePercent",
    "ChangedModifiedDate",
    "ChangedModifiedDatePercent",
    "TargetOlderThanSource",
    "TargetOlderThanSourcePercent",
    "ChangedVersion",
    "ChangedVersionPercent",
    "SourceBytes",
    "TargetBytes",
    "DifferentSizeSourceBytes",
    "DifferentSizeTargetBytes",
    "DifferentSizeDeltaBytes",
]
DEFAULT_DOCUMENT_LIBRARY_SEGMENTS = {
    "documents",
    "documents partages",
    "shared documents",
}
SHAREGATE_REPLACED_CHARACTERS = {"&", "#", "%", '"', "*", ":", ";", "<", ">", "?", "\\", "|", "{", "}", "~"}
SHAREGATE_REPLACEMENT_CHARACTER = "_"
COMPARISON_VERSION_SEPARATOR = "|version="
CSV_OUTPUT_DELIMITER = ";"
EXCLUDED_EXACT_FILE_NAMES = {"desktop.ini", ".ds_store"}
EXCLUDED_FILE_EXTENSIONS = {".db"}
PERCENT_SEQUENCE_PATTERN = re.compile(r"%([0-9A-Fa-f]{2})")
MODIFIED_DATE_FORMATS = (
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%d",
)
MODIFIED_TIME_ZONE_MODES = {
    "": "Raw",
    "none": "Raw",
    "raw": "Raw",
    "local": "Local",
    "utc": "UTC",
}


def detect_csv_dialect(handle):
    sample = handle.read(65536)
    handle.seek(0)
    if not sample:
        return csv.excel

    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        return csv.excel


def to_int(value):
    try:
        return int(str(value or "0").strip())
    except ValueError:
        return 0


def html_escape(value):
    return html.escape(str(value or ""), quote=True)


def html_file_link(path):
    path = Path(path)
    return f'<a href="{html_escape(quote(path.name))}">{html_escape(path.name)}</a>'


def format_integer(value):
    return f"{to_int(value):,}".replace(",", " ")


def top_file_difference_rows(library_summary_rows, limit=20):
    def difference_count(row):
        return (
            to_int(row.get("MissingInTarget"))
            + to_int(row.get("ExtraInTarget"))
            + to_int(row.get("DifferentSize"))
            + to_int(row.get("ChangedModifiedDate"))
            + to_int(row.get("TargetOlderThanSource"))
            + to_int(row.get("ChangedVersion"))
        )

    rows = [row for row in library_summary_rows if difference_count(row)]
    rows.sort(
        key=lambda row: (
            -difference_count(row),
            str(row.get("WebPath") or ""),
            str(row.get("SourceLibraryTitle") or row.get("TargetLibraryTitle") or ""),
        )
    )
    return rows[:limit]


def create_file_html_summary(path, title, summary, library_summary_rows, report_links):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    matched = to_int(summary.get("MatchedKeys"))
    missing = to_int(summary.get("MissingInTarget"))
    extra = to_int(summary.get("ExtraInTarget"))
    extra_folders = to_int(summary.get("ExtraFoldersInTarget"))
    different_size = to_int(summary.get("DifferentSize"))
    changed_modified = to_int(summary.get("ChangedModifiedDate"))
    target_older = to_int(summary.get("TargetOlderThanSource"))
    changed_version = to_int(summary.get("ChangedVersion"))
    review_count = missing + extra + different_size + changed_modified + target_older + changed_version
    status_text = "Review needed" if review_count or extra_folders else "No relevant difference"
    status_class = "warn" if review_count else ("note" if extra_folders else "ok")

    cards = [
        ("Matched files", matched, "ok"),
        ("Missing in target", missing, "bad" if missing else "ok"),
        ("Extra in target", extra, "bad" if extra else "ok"),
        ("Changed version", changed_version, "bad" if changed_version else "ok"),
        ("Different size", different_size, "bad" if different_size else "ok"),
        ("Target older", target_older, "bad" if target_older else "ok"),
        ("Modified date", changed_modified, "note" if changed_modified else "ok"),
        ("Extra folders", extra_folders, "note" if extra_folders else "ok"),
    ]
    card_html = [
        f'<div class="metric {css_class}"><div class="metric-label">{html_escape(label)}</div><div class="metric-value">{format_integer(value)}</div></div>'
        for label, value, css_class in cards
    ]

    report_link_rows = []
    for label, target_path, description in report_links:
        target_path = Path(target_path) if target_path else None
        if not target_path or not target_path.exists():
            continue
        report_link_rows.append(
            "<tr>"
            f"<td>{html_escape(label)}</td>"
            f"<td>{html_file_link(target_path)}</td>"
            f"<td>{html_escape(description)}</td>"
            "</tr>"
        )
    if not report_link_rows:
        report_link_rows.append('<tr><td colspan="3" class="empty">No report files found.</td></tr>')

    top_rows_html = []
    for row in top_file_difference_rows(library_summary_rows):
        top_rows_html.append(
            "<tr>"
            f"<td>{html_escape(row.get('Status'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('MissingInTarget'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('ExtraInTarget'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('ChangedVersion'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('DifferentSize'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('TargetOlderThanSource'))}</td>"
            f"<td>{html_escape(row.get('WebPath'))}</td>"
            f"<td>{html_escape(row.get('SourceLibraryTitle') or row.get('TargetLibraryTitle') or row.get('LibraryPath'))}</td>"
            "</tr>"
        )
    if not top_rows_html:
        top_rows_html.append('<tr><td colspan="8" class="empty">No library-level differences.</td></tr>')

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    document = f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{html_escape(title)}</title>
<style>
:root {{ color-scheme: light; --bg:#F5F8FB; --card:#FFFFFF; --text:#1F2937; --muted:#5F6B7A; --line:#DDE7F0; --accent:#0078D4; --ok:#107C10; --bad:#C50F1F; --note:#8A6A00; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--text); font-family:"Segoe UI", Arial, sans-serif; font-size:14px; line-height:1.45; }}
main {{ max-width:1320px; margin:0 auto; padding:28px; }}
.header {{ background:var(--card); border:1px solid var(--line); border-radius:8px; padding:22px 24px; margin-bottom:18px; display:flex; justify-content:space-between; gap:18px; align-items:flex-start; }}
h1 {{ margin:0 0 6px; font-size:25px; font-weight:650; letter-spacing:0; }}
.subtitle {{ color:var(--muted); }}
.badge {{ display:inline-block; border:1px solid var(--line); border-radius:999px; padding:5px 10px; font-weight:600; background:#fff; white-space:nowrap; }}
.badge.ok {{ color:var(--ok); border-color:#B8DAB8; background:#F1FAF1; }}
.badge.warn {{ color:var(--bad); border-color:#F1B7BC; background:#FFF4F5; }}
.badge.note {{ color:var(--note); border-color:#E7D99B; background:#FFF9DF; }}
.metrics {{ display:grid; grid-template-columns:repeat(4, minmax(150px, 1fr)); gap:12px; margin-bottom:18px; }}
.metric {{ background:var(--card); border:1px solid var(--line); border-radius:8px; padding:15px; min-height:92px; }}
.metric-label {{ color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }}
.metric-value {{ font-size:28px; font-weight:700; margin-top:8px; }}
.metric.ok .metric-value {{ color:var(--ok); }} .metric.bad .metric-value {{ color:var(--bad); }} .metric.note .metric-value {{ color:var(--note); }}
.section {{ background:var(--card); border:1px solid var(--line); border-radius:8px; padding:18px; margin-bottom:18px; }}
h2 {{ margin:0 0 12px; font-size:17px; }}
.grid {{ display:grid; grid-template-columns:220px minmax(0, 1fr); gap:8px 14px; }}
.key {{ color:var(--muted); }}
a {{ color:var(--accent); text-decoration:none; }} a:hover {{ text-decoration:underline; }}
table {{ width:100%; border-collapse:collapse; }}
th, td {{ border-bottom:1px solid var(--line); padding:9px 10px; text-align:left; vertical-align:top; }}
th {{ background:#F8FBFE; color:#334155; font-size:12px; text-transform:uppercase; letter-spacing:.04em; }}
.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
.empty {{ color:var(--muted); text-align:center; padding:18px; }}
.footer {{ color:var(--muted); font-size:12px; margin-top:16px; }}
@media (max-width:1000px) {{ .metrics {{ grid-template-columns:repeat(2, minmax(150px, 1fr)); }} .header {{ display:block; }} .badge {{ margin-top:12px; }} }}
</style>
</head>
<body>
<main>
  <div class="header">
    <div>
      <h1>{html_escape(title)}</h1>
      <div class="subtitle">Generated at {html_escape(generated_at)}</div>
    </div>
    <div class="badge {status_class}">{html_escape(status_text)}</div>
  </div>
  <div class="metrics">{''.join(card_html)}</div>
  <div class="section">
    <h2>Run context</h2>
    <div class="grid">
      <div class="key">Source CSV</div><div>{html_escape(summary.get('SourceCsv'))}</div>
      <div class="key">Target CSV</div><div>{html_escape(summary.get('TargetCsv'))}</div>
      <div class="key">Source unique keys</div><div>{format_integer(summary.get('SourceUniqueKeys'))}</div>
      <div class="key">Target unique keys</div><div>{format_integer(summary.get('TargetUniqueKeys'))}</div>
      <div class="key">Size tolerance bytes</div><div>{format_integer(summary.get('SizeToleranceBytes'))}</div>
      <div class="key">Modified date tolerance minutes</div><div>{html_escape(summary.get('ModifiedDateToleranceMinutes'))}</div>
      <div class="key">Comparison key includes version</div><div>{html_escape(summary.get('ComparisonKeyIncludesVersion'))}</div>
    </div>
  </div>
  <div class="section">
    <h2>Report files</h2>
    <table><thead><tr><th>Report</th><th>File</th><th>Description</th></tr></thead><tbody>{''.join(report_link_rows)}</tbody></table>
  </div>
  <div class="section">
    <h2>Top libraries with differences</h2>
    <table><thead><tr><th>Status</th><th>Missing</th><th>Extra</th><th>Version</th><th>Size</th><th>Older</th><th>Web path</th><th>Library</th></tr></thead><tbody>{''.join(top_rows_html)}</tbody></table>
  </div>
  <div class="footer">SmartM365 SharePoint migration file comparison summary.</div>
</main>
</body>
</html>
'''
    path.write_text(document, encoding="utf-8", newline="\n")
    return path


def strip_query_from_path(value):
    text = str(value).strip()
    if text.lower().startswith(("http://", "https://")):
        return urlparse(text).path
    return text.split("?", 1)[0]


def replace_encoded_sharegate_character(match):
    try:
        character = bytes.fromhex(match.group(1)).decode("utf-8")
    except UnicodeDecodeError:
        return match.group(0)

    if character in SHAREGATE_REPLACED_CHARACTERS or character in {"/", "\\"} or character == "\t" or not character.isprintable():
        if character == "%":
            return SHAREGATE_REPLACEMENT_CHARACTER
        return SHAREGATE_REPLACEMENT_CHARACTER + match.group(1)

    return match.group(0)


def replace_literal_percent_sequence(match):
    if match.group(1).lower() == "25":
        return SHAREGATE_REPLACEMENT_CHARACTER
    return SHAREGATE_REPLACEMENT_CHARACTER + match.group(1)


def decode_sharegate_path(path, literal_percent_sequences=False):
    if literal_percent_sequences:
        path = PERCENT_SEQUENCE_PATTERN.sub(replace_literal_percent_sequence, path)
        return unquote(path)

    path = PERCENT_SEQUENCE_PATTERN.sub(replace_encoded_sharegate_character, path)
    return unquote(path)


def replace_edge_dots_and_spaces(value):
    value = re.sub(r"^[ .]+", lambda match: SHAREGATE_REPLACEMENT_CHARACTER * len(match.group(0)), value)
    return re.sub(r"[ .]+$", lambda match: SHAREGATE_REPLACEMENT_CHARACTER * len(match.group(0)), value)


def replace_consecutive_dots(value):
    return re.sub(
        r"\.{2,}",
        lambda match: SHAREGATE_REPLACEMENT_CHARACTER * (len(match.group(0)) - 1) + ".",
        value,
    )


def replace_edge_dots_and_spaces_before_extension(segment):
    match = re.match(r"^(?P<stem>.+?)(?P<extension>\.[A-Za-z0-9]{1,12})$", segment)
    if not match:
        return replace_edge_dots_and_spaces(segment)

    stem = replace_edge_dots_and_spaces(match.group("stem"))
    stem = re.sub(r"(?:[._]*_+[._]*)+$", SHAREGATE_REPLACEMENT_CHARACTER, stem)
    return stem + match.group("extension")


def normalize_sharegate_segment(segment):
    if segment == "":
        return segment

    clean_segment = segment
    clean_segment = re.sub("_vti_", SHAREGATE_REPLACEMENT_CHARACTER, clean_segment, flags=re.IGNORECASE)
    clean_segment = "".join(
        SHAREGATE_REPLACEMENT_CHARACTER
        if character in SHAREGATE_REPLACED_CHARACTERS or character == "\t" or not character.isprintable()
        else character
        for character in clean_segment
    )
    clean_segment = replace_consecutive_dots(clean_segment)
    clean_segment = replace_edge_dots_and_spaces_before_extension(clean_segment)
    return clean_segment


def normalize_sharegate_path_characters(path):
    if not path:
        return path

    segments = []
    for segment in path.split("/"):
        segments.append(normalize_sharegate_segment(segment))

    return "/".join(segments)


def normalize_mapping_path(value):
    text = normalize_sharegate_path_characters(decode_sharegate_path(strip_query_from_path(value)))
    if not text.startswith("/"):
        text = "/" + text
    return text.rstrip("/") or "/"


def load_path_mappings(path):
    if not path:
        return []

    mappings = []
    with Path(path).open("r", encoding="utf-8-sig", errors="replace") as handle:
        for line_number, raw_line in enumerate(handle, start=1):
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue

            parts = [part for part in re.split(r"[\t;, ]+", line) if part]
            if len(parts) < 2:
                raise ValueError(f"Invalid path mapping at {path}:{line_number}. Expected: <source-url-or-path> <target-url-or-path>.")

            source = normalize_mapping_path(parts[0])
            target = normalize_mapping_path(parts[1])
            mappings.append((source, target))

    mappings.sort(key=lambda item: len(item[0]), reverse=True)
    return mappings


def apply_path_mappings(path, mappings):
    if not path or not mappings:
        return path

    path_compare = path.lower()
    for source, target in mappings:
        source_compare = source.lower()
        if path_compare == source_compare:
            return target
        if path_compare.startswith((source.rstrip("/") + "/").lower()):
            suffix = path[len(source.rstrip("/")):]
            return (target.rstrip("/") + suffix) or "/"

    return path


def normalize_path(value, prefixes, literal_percent_sequences=False, mappings=None):
    if not value:
        return None

    text = normalize_sharegate_path_characters(
        decode_sharegate_path(strip_query_from_path(value), literal_percent_sequences=literal_percent_sequences)
    )

    if not text.startswith("/"):
        text = "/" + text

    text = text.rstrip("/") or "/"
    text = apply_path_mappings(text, mappings)

    for prefix in prefixes:
        if not prefix:
            continue
        clean = normalize_sharegate_path_characters(decode_sharegate_path(strip_query_from_path(prefix)))
        if not clean.startswith("/"):
            clean = "/" + clean
        clean = clean.rstrip("/")
        if not clean:
            continue
        if text.lower() == clean.lower():
            text = "/"
            break
        if text.lower().startswith((clean + "/").lower()):
            text = text[len(clean):]
            break

    if not text.startswith("/"):
        text = "/" + text
    while "//" in text:
        text = text.replace("//", "/")
    return text.lower()


def parse_modified_datetime(value):
    text = str(value or "").strip()
    if not text:
        return None

    text = text.replace("Z", "").split(".", 1)[0]
    for date_format in MODIFIED_DATE_FORMATS:
        try:
            return datetime.strptime(text, date_format)
        except ValueError:
            continue

    return None


def normalize_modified_time_zone_mode(value):
    key = str(value or "").strip().lower()
    if key in MODIFIED_TIME_ZONE_MODES:
        return MODIFIED_TIME_ZONE_MODES[key]
    valid_values = ", ".join(sorted(set(MODIFIED_TIME_ZONE_MODES.values())))
    raise ValueError(f"Unsupported modified date time zone mode '{value}'. Use one of: {valid_values}.")


def normalize_modified_datetime(value, time_zone_mode):
    parsed = parse_modified_datetime(value)
    if parsed is None:
        return None

    mode = normalize_modified_time_zone_mode(time_zone_mode)
    if mode == "UTC":
        return parsed.replace(tzinfo=timezone.utc)
    if mode == "Local":
        return parsed.astimezone().astimezone(timezone.utc)
    return parsed


def comparable_modified_datetimes(source_datetime, target_datetime):
    if source_datetime is None or target_datetime is None:
        return source_datetime, target_datetime

    source_aware = source_datetime.tzinfo is not None and source_datetime.utcoffset() is not None
    target_aware = target_datetime.tzinfo is not None and target_datetime.utcoffset() is not None
    if source_aware == target_aware:
        return source_datetime, target_datetime

    if source_aware:
        source_datetime = source_datetime.astimezone(timezone.utc).replace(tzinfo=None)
    if target_aware:
        target_datetime = target_datetime.astimezone(timezone.utc).replace(tzinfo=None)
    return source_datetime, target_datetime


def format_normalized_utc(value, time_zone_mode):
    parsed = normalize_modified_datetime(value, time_zone_mode)
    if parsed is None:
        return ""

    is_aware = parsed.tzinfo is not None and parsed.utcoffset() is not None
    if is_aware:
        parsed = parsed.astimezone(timezone.utc).replace(tzinfo=None)
    return parsed.strftime("%Y-%m-%d %H:%M:%S")


def modified_date_delta_minutes(source_value, target_value, source_time_zone="Raw", target_time_zone="Raw"):
    source_text = str(source_value or "").strip()
    target_text = str(target_value or "").strip()
    if not source_text and not target_text:
        return None

    source_datetime = normalize_modified_datetime(source_text, source_time_zone)
    target_datetime = normalize_modified_datetime(target_text, target_time_zone)
    if source_datetime is None or target_datetime is None:
        if source_text == target_text:
            return None
        return ""

    source_datetime, target_datetime = comparable_modified_datetimes(source_datetime, target_datetime)
    return abs((target_datetime - source_datetime).total_seconds()) / 60


def signed_modified_date_delta_minutes(source_value, target_value, source_time_zone="Raw", target_time_zone="Raw"):
    source_datetime = normalize_modified_datetime(source_value, source_time_zone)
    target_datetime = normalize_modified_datetime(target_value, target_time_zone)
    if source_datetime is None or target_datetime is None:
        return None

    source_datetime, target_datetime = comparable_modified_datetimes(source_datetime, target_datetime)
    return (target_datetime - source_datetime).total_seconds() / 60


def modified_direction(delta_minutes):
    if delta_minutes is None:
        return ""
    if delta_minutes < 0:
        return "TargetOlderThanSource"
    if delta_minutes > 0:
        return "TargetNewerThanSource"
    return "Same"


def normalize_version(value):
    text = str(value or "").strip()
    if not text:
        return ""

    parts = [part.strip() for part in text.split(".")]
    normalized_parts = []
    for part in parts:
        if re.fullmatch(r"\d+", part):
            normalized_parts.append(str(int(part)))
        else:
            normalized_parts.append(part.lower())
    return ".".join(normalized_parts)


def parse_version_tuple(value):
    text = normalize_version(value)
    if not text:
        return None

    parts = text.split(".")
    if not all(re.fullmatch(r"\d+", part) for part in parts):
        return None

    return tuple(int(part) for part in parts)


def compare_versions(source_value, target_value):
    source_text = normalize_version(source_value)
    target_text = normalize_version(target_value)
    if not source_text and not target_text:
        return "Same"
    if source_text == target_text:
        return "Same"

    source_tuple = parse_version_tuple(source_value)
    target_tuple = parse_version_tuple(target_value)
    if source_tuple is None or target_tuple is None:
        return "Different"

    width = max(len(source_tuple), len(target_tuple))
    source_tuple = source_tuple + (0,) * (width - len(source_tuple))
    target_tuple = target_tuple + (0,) * (width - len(target_tuple))
    if target_tuple < source_tuple:
        return "TargetOlderVersion"
    if target_tuple > source_tuple:
        return "TargetNewerVersion"
    return "Same"


def split_normalized_path(value):
    text = str(value or "").strip().strip("/")
    if not text:
        return []
    return [segment for segment in text.split("/") if segment]


def join_path_segments(segments):
    if not segments:
        return "/"
    return "/" + "/".join(segments)


def parent_path(value):
    segments = split_normalized_path(value)
    if len(segments) <= 1:
        return "/"
    return join_path_segments(segments[:-1])


def ancestor_paths(value):
    segments = split_normalized_path(value)
    return [join_path_segments(segments[:index]) for index in range(1, len(segments) + 1)]


def path_depth(value):
    return len(split_normalized_path(value))


def path_name(value):
    segments = split_normalized_path(value)
    if not segments:
        return ""
    return segments[-1]


def inventory_file_name(row):
    file_name = str(row.get("FileName") or "").strip()
    if file_name:
        return file_name

    path = decode_sharegate_path(strip_query_from_path(row.get("ServerRelativeUrl") or row.get("FileUrl") or ""))
    if not path:
        return ""

    return path.rstrip("/").rsplit("/", 1)[-1].strip()


def inventory_file_name_has_literal_percent_sequence(row):
    return bool(PERCENT_SEQUENCE_PATTERN.search(str(row.get("FileName") or "")))


def replace_last_path_segment_with_file_name(path, file_name):
    clean_path = strip_query_from_path(path or "")
    clean_file_name = str(file_name or "").strip()
    if not clean_path or not clean_file_name:
        return path

    trimmed_path = clean_path.rstrip("/\\")
    slash_index = max(trimmed_path.rfind("/"), trimmed_path.rfind("\\"))
    if slash_index < 0:
        return clean_file_name

    return trimmed_path[: slash_index + 1] + clean_file_name


def is_excluded_inventory_file(row):
    risk_type, reason = inventory_exclusion_reason(row)
    return bool(risk_type or reason)


def inventory_exclusion_reason(row):
    file_name = inventory_file_name(row)
    if not file_name:
        return "", ""

    lower_file_name = file_name.lower()
    if lower_file_name.startswith("~$"):
        return "Temporary file", "Office temporary files starting with ~$ are not valid migration candidates."
    if lower_file_name == "desktop.ini":
        return "Blocked or system file", "desktop.ini is a system file name blocked by SharePoint Online."
    if lower_file_name == ".ds_store":
        return "System metadata file", ".DS_Store is a macOS metadata file and is excluded from comparison."
    if any(lower_file_name.endswith(extension) for extension in EXCLUDED_FILE_EXTENSIONS):
        return "System or cache file type", "Files with the .db extension, including Thumbs.db, are excluded from comparison."

    return "", ""


def inventory_path_key(row, prefixes, mappings=None):
    path = row.get("ServerRelativeUrl") or row.get("FileUrl")
    literal_percent_sequences = inventory_file_name_has_literal_percent_sequence(row)
    if literal_percent_sequences:
        path = replace_last_path_segment_with_file_name(path, row.get("FileName"))

    key = normalize_path(path, prefixes, literal_percent_sequences=literal_percent_sequences, mappings=mappings)
    return normalize_default_document_library_key(row, key, prefixes, mappings=mappings)


def inventory_key(row, prefixes, mappings=None):
    path_key = inventory_path_key(row, prefixes, mappings=mappings)
    if not path_key:
        return path_key

    return f"{path_key}{COMPARISON_VERSION_SEPARATOR}{normalize_version(row.get('Version'))}"


def comparison_path_key(key):
    return str(key or "").split(COMPARISON_VERSION_SEPARATOR, 1)[0]


def inventory_web_key(row, prefixes, mappings=None):
    return normalize_path(row.get("WebUrl"), prefixes, mappings=mappings)


def normalize_library_title(value):
    return " ".join(str(value or "").strip().lower().split())


def normalize_default_document_library_key(row, key, prefixes, mappings=None):
    if not key:
        return key

    web_key = inventory_web_key(row, prefixes, mappings=mappings) or ""
    web_prefix = web_key.rstrip("/")
    relative_path = ""
    base_path = ""

    if web_prefix and web_prefix != "/" and key.startswith(web_prefix + "/"):
        relative_path = key[len(web_prefix) :].strip("/")
        base_path = web_prefix
    elif web_key == "/" and key.startswith("/"):
        relative_path = key.strip("/")
    else:
        return key

    if not relative_path:
        return key

    first_segment, separator, remaining_path = relative_path.partition("/")
    if first_segment.lower() not in DEFAULT_DOCUMENT_LIBRARY_SEGMENTS:
        return key

    normalized_relative_path = "documents"
    if separator:
        normalized_relative_path += "/" + remaining_path

    if base_path:
        return f"{base_path}/{normalized_relative_path}".replace("//", "/")
    return "/" + normalized_relative_path


def library_key(row, prefixes, mappings=None):
    web_key = inventory_web_key(row, prefixes, mappings=mappings) or ""
    file_key = inventory_path_key(row, prefixes, mappings=mappings) or ""
    library_path = ""

    if file_key and web_key and (file_key == web_key or file_key.startswith((web_key.rstrip("/") + "/"))):
        relative_path = file_key[len(web_key.rstrip("/")):].strip("/")
    else:
        relative_path = file_key.strip("/")

    if relative_path:
        first_segment = relative_path.split("/", 1)[0]
        library_path = f"{web_key.rstrip('/')}/{first_segment}".replace("//", "/")

    if not library_path:
        library_path = f"{web_key}|{normalize_library_title(row.get('LibraryTitle'))}"

    return library_path.lower(), web_key, library_path


def ensure_library_summary(stats, key, web_path, library_path=""):
    if key not in stats:
        stats[key] = {
            "Status": "",
            "LibraryKey": key,
            "WebPath": web_path,
            "LibraryPath": library_path,
            "SourceWebUrl": "",
            "TargetWebUrl": "",
            "SourceWebTitle": "",
            "TargetWebTitle": "",
            "SourceLibraryTitle": "",
            "TargetLibraryTitle": "",
            "SourceFiles": 0,
            "TargetFiles": 0,
            "MatchedFiles": 0,
            "MatchedFilesPercent": 0,
            "MissingInTarget": 0,
            "MissingInTargetPercent": 0,
            "ExtraInTarget": 0,
            "ExtraInTargetPercent": 0,
            "ExtraInTargetBytes": 0,
            "DifferentSize": 0,
            "DifferentSizePercent": 0,
            "ChangedModifiedDate": 0,
            "ChangedModifiedDatePercent": 0,
            "TargetOlderThanSource": 0,
            "TargetOlderThanSourcePercent": 0,
            "ChangedVersion": 0,
            "ChangedVersionPercent": 0,
            "SourceBytes": 0,
            "TargetBytes": 0,
            "DifferentSizeSourceBytes": 0,
            "DifferentSizeTargetBytes": 0,
            "DifferentSizeDeltaBytes": 0,
        }
    return stats[key]


def percent(numerator, denominator):
    denominator = int(denominator or 0)
    if denominator <= 0:
        return "0.000000"
    return f"{int(numerator or 0) / denominator:.6f}"


def library_status(entry):
    missing = int(entry.get("MissingInTarget") or 0)
    extra = int(entry.get("ExtraInTarget") or 0)
    different_size = int(entry.get("DifferentSize") or 0)
    changed_modified = int(entry.get("ChangedModifiedDate") or 0)
    target_older = int(entry.get("TargetOlderThanSource") or 0)
    changed_version = int(entry.get("ChangedVersion") or 0)

    if missing == 0 and extra == 0 and different_size == 0 and changed_modified == 0 and target_older == 0 and changed_version == 0:
        return "OK - no difference"

    issues = []
    if missing:
        issues.append("missing in SPO")
    if extra:
        issues.append("extra in SPO")
    if different_size:
        issues.append("size differences")
    if changed_modified:
        issues.append("modified date differences")
    if target_older:
        issues.append("target older than source")
    if changed_version:
        issues.append("version differences")

    if len(issues) == 1:
        labels = {
            "missing in SPO": "Missing in SPO",
            "extra in SPO": "Extra in SPO",
            "size differences": "Size differences",
            "modified date differences": "Modified date differences",
            "target older than source": "Target older than source",
            "version differences": "Version differences",
        }
        return labels[issues[0]]

    return "Mixed differences: " + ", ".join(issues)


def update_source_library(stats, key, web_path, library_path, row, size):
    entry = ensure_library_summary(stats, key, web_path, library_path)
    entry["SourceFiles"] += 1
    entry["SourceBytes"] += size or 0
    if not entry["SourceWebUrl"]:
        entry["SourceWebUrl"] = row.get("WebUrl", "")
    if not entry["SourceWebTitle"]:
        entry["SourceWebTitle"] = row.get("WebTitle", "")
    if not entry["SourceLibraryTitle"]:
        entry["SourceLibraryTitle"] = row.get("LibraryTitle", "")
    return entry


def update_target_library(stats, key, web_path, row, size):
    library_path = row.get("LibraryPath", "")
    entry = ensure_library_summary(stats, key, web_path, library_path)
    entry["TargetFiles"] += 1
    entry["TargetBytes"] += size or 0
    if not entry["TargetWebUrl"]:
        entry["TargetWebUrl"] = row.get("WebUrl", "")
    if not entry["TargetWebTitle"]:
        entry["TargetWebTitle"] = row.get("WebTitle", "")
    if not entry["TargetLibraryTitle"]:
        entry["TargetLibraryTitle"] = row.get("LibraryTitle", "")
    return entry


def bytes_to_mb(value):
    return f"{int(value or 0) / 1024 / 1024:.4f}"


def is_under_library(folder_key, library_path):
    clean_folder = str(folder_key or "").rstrip("/").lower()
    clean_library = str(library_path or "").rstrip("/").lower()
    if not clean_folder or not clean_library:
        return True
    return clean_folder.startswith(clean_library + "/")


def update_extra_folder_candidates(candidates, source_folder_keys, row, file_key, file_size):
    normalized_folder_ancestors = ancestor_paths(parent_path(file_key))
    actual_folder_ancestors = ancestor_paths(parent_path(row.get("ServerRelativeUrl", "")))
    if not normalized_folder_ancestors or not actual_folder_ancestors:
        return

    offset = len(actual_folder_ancestors) - len(normalized_folder_ancestors)
    if offset < 0:
        return

    library_path = str(row.get("LibraryPath") or "").rstrip("/").lower()
    web_path = str(row.get("WebPath") or "").rstrip("/").lower()

    for index, normalized_folder_key in enumerate(normalized_folder_ancestors):
        normalized_folder_key = normalized_folder_key.lower().rstrip("/") or "/"
        if normalized_folder_key in source_folder_keys:
            continue
        if normalized_folder_key in {"/", web_path, library_path}:
            continue
        if not is_under_library(normalized_folder_key, library_path):
            continue

        actual_folder_url = actual_folder_ancestors[index + offset]
        candidate = candidates.setdefault(
            normalized_folder_key,
            {
                "Action": "Delete SPO folder after extra files are deleted, only if the folder is empty",
                "Reason": "SPO folder path is inferred from extra files and does not exist in the SP2019 source file paths",
                "TargetWebUrl": row.get("WebUrl", ""),
                "TargetLibraryTitle": row.get("LibraryTitle", ""),
                "TargetFolderServerRelativeUrl": actual_folder_url,
                "FolderName": path_name(actual_folder_url),
                "Depth": path_depth(actual_folder_url),
                "ExtraFileCount": 0,
                "ExtraFileBytes": 0,
                "ExtraFileMB": "0.0000",
                "LibraryKey": row.get("LibraryKey", ""),
                "WebPath": row.get("WebPath", ""),
                "LibraryPath": row.get("LibraryPath", ""),
                "NormalizedFolderKey": normalized_folder_key,
            },
        )
        candidate["ExtraFileCount"] += 1
        candidate["ExtraFileBytes"] += file_size or 0
        candidate["ExtraFileMB"] = bytes_to_mb(candidate["ExtraFileBytes"])


def missing_reason_details(source_record, target_path_records):
    if not target_path_records:
        return {
            "MissingReason": "FileAbsent",
            "TargetMatchingPathCount": 0,
            "TargetSizeBytes": "",
            "TargetVersions": "",
            "TargetVersionComparisons": "",
        }

    source_size = as_int(source_record.get("SizeBytes"))
    same_key_records = [record for record in target_path_records if record.get("Key") == source_record.get("Key")]
    zero_size_records = [record for record in same_key_records if (as_int(record.get("SizeBytes")) or 0) == 0]
    missing_reason = "TargetZeroBytes" if source_size and zero_size_records else "VersionMismatch"
    relevant_records = zero_size_records if zero_size_records else target_path_records
    target_sizes = sorted({str(as_int(record.get("SizeBytes")) if as_int(record.get("SizeBytes")) is not None else "") for record in relevant_records})
    target_versions = sorted({normalize_version(record.get("Version")) for record in target_path_records})
    version_comparisons = []
    for version in target_versions:
        comparison = compare_versions(source_record.get("Version"), version)
        version_label = version if version else "<blank>"
        version_comparisons.append(f"{version_label}={comparison}")

    return {
        "MissingReason": missing_reason,
        "TargetMatchingPathCount": len(target_path_records),
        "TargetSizeBytes": ", ".join(size for size in target_sizes if size),
        "TargetVersions": ", ".join(version if version else "<blank>" for version in target_versions),
        "TargetVersionComparisons": ", ".join(version_comparisons),
    }


def normalized_text(value):
    return " ".join(str(value or "").strip().lower().split())


def source_has_sharegate_sensitive_text(record):
    source_path = str(record.get("ServerRelativeUrl") or record.get("FileUrl") or record.get("FileName") or "")
    return bool(
        re.search(r"%(?:26|23|25|22|2A|3A|3B|3C|3E|3F|5C|7B|7D|7E)", source_path, flags=re.IGNORECASE)
        or any(character in source_path for character in "&#*:;<>?\\{}~")
    )


def target_candidate_score(source_record, target_record):
    score = 0
    if normalized_text(source_record.get("FileName")) == normalized_text(target_record.get("FileName")):
        score += 20
    if as_int(source_record.get("SizeBytes")) == as_int(target_record.get("SizeBytes")):
        score += 15
    if normalize_version(source_record.get("Version")) == normalize_version(target_record.get("Version")):
        score += 15
    if normalized_text(source_record.get("LibraryTitle")) == normalized_text(target_record.get("LibraryTitle")):
        score += 8
    if normalized_text(source_record.get("WebTitle")) == normalized_text(target_record.get("WebTitle")):
        score += 4
    if normalized_text(source_record.get("Modified")) == normalized_text(target_record.get("Modified")):
        score += 3
    return score


def best_target_candidate(source_record, target_records, predicate):
    candidates = [record for record in target_records if predicate(record)]
    if not candidates:
        return [], None
    candidates.sort(key=lambda record: (-target_candidate_score(source_record, record), str(record.get("ServerRelativeUrl") or "").lower()))
    return candidates, candidates[0]


def missing_diagnostic_row(source_record, category, signal, action_hint, candidates=None, candidate=None):
    candidates = candidates or []
    candidate = candidate or (candidates[0] if candidates else {})
    return {
        "Key": source_record.get("Key", ""),
        "DiagnosticCategory": category,
        "DiagnosticSignal": signal,
        "ActionHint": action_hint,
        "SourceFileName": source_record.get("FileName", ""),
        "SourceSizeBytes": source_record.get("SizeBytes", ""),
        "SourceVersion": source_record.get("Version", ""),
        "SourceModified": source_record.get("Modified", ""),
        "SourceWebUrl": source_record.get("WebUrl", ""),
        "SourceLibraryTitle": source_record.get("LibraryTitle", ""),
        "SourceServerRelativeUrl": source_record.get("ServerRelativeUrl", ""),
        "TargetCandidateCount": len(candidates),
        "TargetCandidateFileName": candidate.get("FileName", ""),
        "TargetCandidateSizeBytes": candidate.get("SizeBytes", ""),
        "TargetCandidateVersion": candidate.get("Version", ""),
        "TargetCandidateModified": candidate.get("Modified", ""),
        "TargetCandidateWebUrl": candidate.get("WebUrl", ""),
        "TargetCandidateLibraryTitle": candidate.get("LibraryTitle", ""),
        "TargetCandidateServerRelativeUrl": candidate.get("ServerRelativeUrl", ""),
    }


def missing_diagnostics(source_record, missing_details, target_path_records, target_records):
    rows = []
    reason = missing_details.get("MissingReason")
    if reason == "VersionMismatch":
        rows.append(
            missing_diagnostic_row(
                source_record,
                "WrongVersion",
                "Target contains the same normalized path, but not the source version.",
                "Remigrate or restore the expected source version.",
                target_path_records,
            )
        )
    elif reason == "TargetZeroBytes":
        rows.append(
            missing_diagnostic_row(
                source_record,
                "ZeroByteTarget",
                "Target contains the same normalized path/version with a zero-byte file.",
                "Delete the zero-byte target file and recopy it.",
                target_path_records,
            )
        )
    else:
        source_name = normalized_text(source_record.get("FileName"))
        source_size = as_int(source_record.get("SizeBytes"))
        source_version = normalize_version(source_record.get("Version"))
        source_modified = normalized_text(source_record.get("Modified"))
        candidates, candidate = best_target_candidate(
            source_record,
            target_records,
            lambda record: normalized_text(record.get("FileName")) == source_name
            and as_int(record.get("SizeBytes")) == source_size
            and normalize_version(record.get("Version")) == source_version,
        )
        if candidates:
            rows.append(
                missing_diagnostic_row(
                    source_record,
                    "PossibleRelocatedInTarget",
                    "Target contains the same file name, size, and version at another path.",
                    "Review whether the file was intentionally moved or should be recopied to the expected path.",
                    candidates,
                    candidate,
                )
            )

        candidates, candidate = best_target_candidate(
            source_record,
            target_records,
            lambda record: normalized_text(record.get("FileName")) == source_name
            and as_int(record.get("SizeBytes")) == source_size
            and normalized_text(record.get("Modified")) == source_modified,
        )
        if candidates:
            rows.append(
                missing_diagnostic_row(
                    source_record,
                    "PossibleSameContentDifferentVersion",
                    "Target contains the same file name, size, and modified date, but not necessarily the same version.",
                    "Review target version metadata before deciding whether to recopy.",
                    candidates,
                    candidate,
                )
            )

        if source_has_sharegate_sensitive_text(source_record):
            rows.append(
                missing_diagnostic_row(
                    source_record,
                    "PossibleShareGateRename",
                    "Source path contains characters or encoded sequences that ShareGate may rename.",
                    "Check the target for a ShareGate-renamed path before treating it as truly absent.",
                )
            )
    return rows


def load_web_url_filter(path, prefixes, mappings=None):
    if not path:
        return None

    filter_path = Path(path)
    if not filter_path.exists():
        raise FileNotFoundError(f"Web URL filter file not found: {filter_path}")

    allowed = set()
    with filter_path.open("r", encoding="utf-8-sig") as handle:
        for line in handle:
            value = line.strip()
            if not value or value.startswith("#"):
                continue
            normalized = normalize_path(value, prefixes, mappings=mappings)
            if normalized:
                allowed.add(normalized)
    return allowed


def inventory_record(row, key):
    return {
        "Key": key,
        "LibraryKey": row.get("LibraryKey", ""),
        "WebPath": row.get("WebPath", ""),
        "LibraryPath": row.get("LibraryPath", ""),
        "SiteCollectionUrl": row.get("SiteCollectionUrl", ""),
        "WebUrl": row.get("WebUrl", ""),
        "WebTitle": row.get("WebTitle", ""),
        "LibraryTitle": row.get("LibraryTitle", ""),
        "FileName": row.get("FileName", ""),
        "ServerRelativeUrl": row.get("ServerRelativeUrl", ""),
        "FileUrl": row.get("FileUrl", ""),
        "SizeBytes": row.get("SizeBytes", ""),
        "Modified": row.get("Modified", ""),
        "ModifiedBy": row.get("ModifiedBy", ""),
        "Version": row.get("Version", ""),
        "VersionsCount": row.get("VersionsCount", ""),
    }


def as_int(value):
    if value is None or str(value).strip() == "":
        return None
    return int(str(value).strip())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-csv", required=True)
    parser.add_argument("--target-csv", required=True)
    parser.add_argument("--output-directory", required=True)
    parser.add_argument("--source-prefix", action="append", default=[])
    parser.add_argument("--target-prefix", action="append", default=[])
    parser.add_argument("--path-mapping-file")
    parser.add_argument("--source-web-urls-file")
    parser.add_argument("--target-web-urls-file")
    parser.add_argument("--comparison-name", default="SharePointInventoryComparison")
    parser.add_argument("--size-tolerance-bytes", type=int, default=10240)
    parser.add_argument(
        "--modified-date-tolerance-minutes",
        type=float,
        default=-1,
        help="Compare Modified dates when set to 0 or more. Negative value disables this comparison.",
    )
    parser.add_argument(
        "--source-modified-time-zone",
        default="Raw",
        help="How source Modified values are interpreted for comparison: Raw, Local, or UTC.",
    )
    parser.add_argument(
        "--target-modified-time-zone",
        default="Raw",
        help="How target Modified values are interpreted for comparison: Raw, Local, or UTC.",
    )
    parser.add_argument("--sharegate-replacement-character", default="_")
    args = parser.parse_args()
    if len(args.sharegate_replacement_character) != 1:
        raise ValueError("--sharegate-replacement-character must contain exactly one character.")
    args.source_modified_time_zone = normalize_modified_time_zone_mode(args.source_modified_time_zone)
    args.target_modified_time_zone = normalize_modified_time_zone_mode(args.target_modified_time_zone)

    global SHAREGATE_REPLACEMENT_CHARACTER
    SHAREGATE_REPLACEMENT_CHARACTER = args.sharegate_replacement_character

    source_csv = Path(args.source_csv)
    target_csv = Path(args.target_csv)
    output_directory = Path(args.output_directory)
    output_directory.mkdir(parents=True, exist_ok=True)
    compare_modified_dates = args.modified_date_tolerance_minutes >= 0

    missing_path = output_directory / "MissingInTarget.csv"
    missing_diagnostics_path = output_directory / "MissingDiagnostics.csv"
    extra_path = output_directory / "ExtraInTarget.csv"
    extra_folders_path = output_directory / "ExtraFoldersInTarget.csv"
    different_size_path = output_directory / "DifferentSize.csv"
    changed_modified_date_path = output_directory / "ChangedModifiedDate.csv"
    target_older_than_source_path = output_directory / "TargetOlderThanSource.csv"
    changed_version_path = output_directory / "ChangedVersion.csv"
    duplicate_keys_path = output_directory / "DuplicateKeys.csv"
    summary_path = output_directory / "Summary.csv"

    source_index = {}
    source_folder_keys = set()
    extra_folder_candidates = {}
    library_stats = {}
    path_mappings = load_path_mappings(args.path_mapping_file)
    if path_mappings:
        print(f"Path mappings loaded: {len(path_mappings)}")

    source_allowed_webs = load_web_url_filter(args.source_web_urls_file, args.source_prefix, mappings=path_mappings)
    target_allowed_webs = load_web_url_filter(args.target_web_urls_file, args.target_prefix)
    source_total_rows = 0
    source_filtered_rows = 0
    source_excluded_rows = 0
    source_duplicate_keys = 0
    source_path_index = {}

    print(f"Source CSV: {source_csv}")
    print(f"Target CSV: {target_csv}")
    print(f"Output directory: {output_directory}")
    if compare_modified_dates:
        print(
            "Modified date comparison: "
            f"source={args.source_modified_time_zone}; target={args.target_modified_time_zone}; "
            f"tolerance={args.modified_date_tolerance_minutes} minute(s)"
        )
    print("Indexing source inventory...")

    with duplicate_keys_path.open("w", encoding="utf-8", newline="") as duplicate_handle:
        duplicate_writer = csv.DictWriter(
            duplicate_handle,
            fieldnames=DUPLICATE_KEY_COLUMNS,
            delimiter=CSV_OUTPUT_DELIMITER,
        )
        duplicate_writer.writeheader()

        with source_csv.open("r", encoding="utf-8-sig", newline="") as handle:
            for row in csv.DictReader(handle, dialect=detect_csv_dialect(handle)):
                source_total_rows += 1
                if source_allowed_webs is not None and inventory_web_key(row, args.source_prefix, mappings=path_mappings) not in source_allowed_webs:
                    source_filtered_rows += 1
                    continue
                if is_excluded_inventory_file(row):
                    source_excluded_rows += 1
                    continue

                path_key = inventory_path_key(row, args.source_prefix, mappings=path_mappings)
                key = inventory_key(row, args.source_prefix, mappings=path_mappings)
                if not key:
                    continue

                lib_key, web_path, library_path = library_key(row, args.source_prefix, mappings=path_mappings)
                row["LibraryKey"] = lib_key
                row["WebPath"] = web_path
                row["LibraryPath"] = library_path

                if key in source_index:
                    source_duplicate_keys += 1
                    duplicate_writer.writerow({"Side": "Source", **inventory_record(row, key)})
                    continue

                for folder_key in ancestor_paths(parent_path(path_key)):
                    source_folder_keys.add(folder_key.lower())

                update_source_library(library_stats, lib_key, web_path, library_path, row, as_int(row.get("SizeBytes")) or 0)
                source_record = inventory_record(row, key)
                source_index[key] = source_record
                source_path_index.setdefault(path_key, []).append(source_record)

        print(
            f"Source rows: {source_total_rows}; unique keys: {len(source_index)}; "
            f"duplicate keys ignored: {source_duplicate_keys}; filtered rows: {source_filtered_rows}; "
            f"excluded rows: {source_excluded_rows}"
        )
        print("Comparing target inventory...")

        target_seen = set()
        target_valid_seen = set()
        target_records = []
        target_path_index = {}
        target_total_rows = 0
        target_filtered_rows = 0
        target_excluded_rows = 0
        target_duplicate_keys = 0
        matched = 0
        extra = 0
        different_size = 0
        changed_modified_date = 0
        target_older_than_source = 0
        changed_version = 0

        with extra_path.open("w", encoding="utf-8", newline="") as extra_handle, different_size_path.open(
            "w", encoding="utf-8", newline=""
        ) as diff_handle, changed_modified_date_path.open(
            "w", encoding="utf-8", newline=""
        ) as changed_modified_handle, target_older_than_source_path.open(
            "w", encoding="utf-8", newline=""
        ) as target_older_handle, changed_version_path.open(
            "w", encoding="utf-8", newline=""
        ) as changed_version_handle, target_csv.open("r", encoding="utf-8-sig", newline="") as target_handle:
            extra_writer = csv.DictWriter(extra_handle, fieldnames=INVENTORY_COLUMNS, delimiter=CSV_OUTPUT_DELIMITER)
            diff_writer = csv.DictWriter(diff_handle, fieldnames=DIFFERENT_SIZE_COLUMNS, delimiter=CSV_OUTPUT_DELIMITER)
            changed_modified_writer = csv.DictWriter(
                changed_modified_handle,
                fieldnames=CHANGED_MODIFIED_DATE_COLUMNS,
                delimiter=CSV_OUTPUT_DELIMITER,
            )
            target_older_writer = csv.DictWriter(
                target_older_handle,
                fieldnames=TARGET_OLDER_THAN_SOURCE_COLUMNS,
                delimiter=CSV_OUTPUT_DELIMITER,
            )
            changed_version_writer = csv.DictWriter(
                changed_version_handle,
                fieldnames=CHANGED_VERSION_COLUMNS,
                delimiter=CSV_OUTPUT_DELIMITER,
            )
            extra_writer.writeheader()
            diff_writer.writeheader()
            changed_modified_writer.writeheader()
            target_older_writer.writeheader()
            changed_version_writer.writeheader()

            for row in csv.DictReader(target_handle, dialect=detect_csv_dialect(target_handle)):
                target_total_rows += 1
                if target_allowed_webs is not None and inventory_web_key(row, args.target_prefix) not in target_allowed_webs:
                    target_filtered_rows += 1
                    continue
                if is_excluded_inventory_file(row):
                    target_excluded_rows += 1
                    continue

                path_key = inventory_path_key(row, args.target_prefix)
                key = inventory_key(row, args.target_prefix)
                if not key:
                    continue

                target_lib_key, target_web_path, target_library_path = library_key(row, args.target_prefix)
                row["LibraryKey"] = target_lib_key
                row["WebPath"] = target_web_path
                row["LibraryPath"] = target_library_path

                if key in target_seen:
                    target_duplicate_keys += 1
                    duplicate_writer.writerow({"Side": "Target", **inventory_record(row, key)})
                    continue

                target_seen.add(key)
                update_target_library(library_stats, target_lib_key, target_web_path, row, as_int(row.get("SizeBytes")) or 0)
                target_record = inventory_record(row, key)
                target_records.append(target_record)
                target_path_index.setdefault(path_key, []).append(target_record)
                source_record = source_index.get(key)
                source_size = as_int(source_record["SizeBytes"]) if source_record is not None else None
                target_size = as_int(target_record["SizeBytes"])
                target_effective_size = target_size or 0
                if source_record is not None and source_size and target_effective_size == 0:
                    extra += 1
                    extra_library_entry = ensure_library_summary(library_stats, target_lib_key, target_web_path, target_library_path)
                    extra_library_entry["ExtraInTarget"] += 1
                    extra_library_entry["ExtraInTargetBytes"] += target_effective_size
                    extra_writer.writerow(target_record)
                    continue

                if source_record is None:
                    extra += 1
                    extra_library_entry = ensure_library_summary(library_stats, target_lib_key, target_web_path, target_library_path)
                    extra_library_entry["ExtraInTarget"] += 1
                    extra_library_entry["ExtraInTargetBytes"] += target_effective_size
                    source_path_records = source_path_index.get(path_key, [])
                    source_version_record = next(
                        (
                            candidate
                            for candidate in source_path_records
                            if compare_versions(candidate["Version"], target_record["Version"]) != "Same"
                        ),
                        None,
                    )
                    if source_version_record is not None:
                        version_comparison = compare_versions(source_version_record["Version"], target_record["Version"])
                        changed_version += 1
                        version_library_entry = ensure_library_summary(
                            library_stats,
                            source_version_record.get("LibraryKey") or target_lib_key,
                            source_version_record.get("WebPath") or target_web_path,
                            source_version_record.get("LibraryPath") or target_library_path,
                        )
                        version_library_entry["ChangedVersion"] += 1
                        changed_version_writer.writerow(
                            {
                                "Key": key,
                                "SourceVersion": source_version_record["Version"],
                                "TargetVersion": target_record["Version"],
                                "VersionComparison": version_comparison,
                                "SourceModified": source_version_record["Modified"],
                                "TargetModified": target_record["Modified"],
                                "SourceSizeBytes": source_version_record["SizeBytes"],
                                "TargetSizeBytes": target_effective_size if target_size is not None else "",
                                "SourceWebUrl": source_version_record["WebUrl"],
                                "TargetWebUrl": target_record["WebUrl"],
                                "SourceLibraryTitle": source_version_record["LibraryTitle"],
                                "TargetLibraryTitle": target_record["LibraryTitle"],
                                "SourceServerRelativeUrl": source_version_record["ServerRelativeUrl"],
                                "TargetServerRelativeUrl": target_record["ServerRelativeUrl"],
                            }
                        )
                    update_extra_folder_candidates(extra_folder_candidates, source_folder_keys, row, path_key, target_effective_size)
                    extra_writer.writerow(target_record)
                    continue

                target_valid_seen.add(key)
                matched += 1
                source_lib_key = source_record.get("LibraryKey") or target_lib_key
                source_web_path = source_record.get("WebPath") or target_web_path
                source_library_path = source_record.get("LibraryPath") or target_library_path
                source_library_entry = ensure_library_summary(library_stats, source_lib_key, source_web_path, source_library_path)
                source_library_entry["MatchedFiles"] += 1
                if not source_library_entry["TargetWebUrl"]:
                    source_library_entry["TargetWebUrl"] = target_record["WebUrl"]
                if not source_library_entry["TargetWebTitle"]:
                    source_library_entry["TargetWebTitle"] = target_record["WebTitle"]
                if not source_library_entry["TargetLibraryTitle"]:
                    source_library_entry["TargetLibraryTitle"] = target_record["LibraryTitle"]
                if (
                    source_size is not None
                    and target_size is not None
                    and abs(source_size - target_size) > args.size_tolerance_bytes
                ):
                    different_size += 1
                    source_library_entry["DifferentSize"] += 1
                    source_library_entry["DifferentSizeSourceBytes"] += source_size
                    source_library_entry["DifferentSizeTargetBytes"] += target_size
                    source_library_entry["DifferentSizeDeltaBytes"] += target_size - source_size
                    diff_writer.writerow(
                        {
                            "Key": key,
                            "SourceSizeBytes": source_size,
                            "TargetSizeBytes": target_size,
                            "DeltaBytes": target_size - source_size,
                            "DeltaBytesPercent": f"{((target_size - source_size) / source_size):.6f}" if source_size else "",
                            "SourceWebUrl": source_record["WebUrl"],
                            "TargetWebUrl": target_record["WebUrl"],
                            "SourceLibraryTitle": source_record["LibraryTitle"],
                            "TargetLibraryTitle": target_record["LibraryTitle"],
                            "SourceServerRelativeUrl": source_record["ServerRelativeUrl"],
                            "TargetServerRelativeUrl": target_record["ServerRelativeUrl"],
                            "SourceModified": source_record["Modified"],
                            "TargetModified": target_record["Modified"],
                            "SourceVersion": source_record["Version"],
                            "TargetVersion": target_record["Version"],
                        }
                    )

                signed_modified_delta_minutes = (
                    signed_modified_date_delta_minutes(
                        source_record["Modified"],
                        target_record["Modified"],
                        args.source_modified_time_zone,
                        args.target_modified_time_zone,
                    )
                    if compare_modified_dates
                    else None
                )
                modified_delta_minutes = (
                    abs(signed_modified_delta_minutes)
                    if isinstance(signed_modified_delta_minutes, (int, float))
                    else (
                        modified_date_delta_minutes(
                            source_record["Modified"],
                            target_record["Modified"],
                            args.source_modified_time_zone,
                            args.target_modified_time_zone,
                        )
                        if compare_modified_dates
                        else None
                    )
                )
                if compare_modified_dates and (
                    modified_delta_minutes == "" or (
                        modified_delta_minutes is not None
                        and modified_delta_minutes > args.modified_date_tolerance_minutes
                    )
                ):
                    changed_modified_date += 1
                    source_library_entry["ChangedModifiedDate"] += 1
                    changed_modified_writer.writerow(
                        {
                            "Key": key,
                            "SourceModified": source_record["Modified"],
                            "TargetModified": target_record["Modified"],
                            "SourceModifiedNormalizedUtc": format_normalized_utc(
                                source_record["Modified"], args.source_modified_time_zone
                            ),
                            "TargetModifiedNormalizedUtc": format_normalized_utc(
                                target_record["Modified"], args.target_modified_time_zone
                            ),
                            "DeltaModifiedMinutes": (
                                f"{modified_delta_minutes:.2f}" if isinstance(modified_delta_minutes, (int, float)) else ""
                            ),
                            "ModifiedDirection": modified_direction(signed_modified_delta_minutes),
                            "SourceSizeBytes": source_size if source_size is not None else "",
                            "TargetSizeBytes": target_effective_size if target_size is not None else "",
                            "SourceVersion": source_record["Version"],
                            "TargetVersion": target_record["Version"],
                            "SourceWebUrl": source_record["WebUrl"],
                            "TargetWebUrl": target_record["WebUrl"],
                            "SourceLibraryTitle": source_record["LibraryTitle"],
                            "TargetLibraryTitle": target_record["LibraryTitle"],
                            "SourceServerRelativeUrl": source_record["ServerRelativeUrl"],
                            "TargetServerRelativeUrl": target_record["ServerRelativeUrl"],
                        }
                    )
                    if (
                        isinstance(signed_modified_delta_minutes, (int, float))
                        and signed_modified_delta_minutes < -args.modified_date_tolerance_minutes
                    ):
                        target_older_than_source += 1
                        source_library_entry["TargetOlderThanSource"] += 1
                        target_older_writer.writerow(
                            {
                                "Key": key,
                                "SourceModified": source_record["Modified"],
                                "TargetModified": target_record["Modified"],
                                "SourceModifiedNormalizedUtc": format_normalized_utc(
                                    source_record["Modified"], args.source_modified_time_zone
                                ),
                                "TargetModifiedNormalizedUtc": format_normalized_utc(
                                    target_record["Modified"], args.target_modified_time_zone
                                ),
                                "TargetOlderByMinutes": f"{abs(signed_modified_delta_minutes):.2f}",
                                "SourceSizeBytes": source_size if source_size is not None else "",
                                "TargetSizeBytes": target_effective_size if target_size is not None else "",
                                "SourceVersion": source_record["Version"],
                                "TargetVersion": target_record["Version"],
                                "SourceWebUrl": source_record["WebUrl"],
                                "TargetWebUrl": target_record["WebUrl"],
                                "SourceLibraryTitle": source_record["LibraryTitle"],
                                "TargetLibraryTitle": target_record["LibraryTitle"],
                                "SourceServerRelativeUrl": source_record["ServerRelativeUrl"],
                                "TargetServerRelativeUrl": target_record["ServerRelativeUrl"],
                            }
                        )

                version_comparison = compare_versions(source_record["Version"], target_record["Version"])
                if version_comparison != "Same":
                    changed_version += 1
                    source_library_entry["ChangedVersion"] += 1
                    changed_version_writer.writerow(
                        {
                            "Key": key,
                            "SourceVersion": source_record["Version"],
                            "TargetVersion": target_record["Version"],
                            "VersionComparison": version_comparison,
                            "SourceModified": source_record["Modified"],
                            "TargetModified": target_record["Modified"],
                            "SourceSizeBytes": source_size if source_size is not None else "",
                            "TargetSizeBytes": target_effective_size if target_size is not None else "",
                            "SourceWebUrl": source_record["WebUrl"],
                            "TargetWebUrl": target_record["WebUrl"],
                            "SourceLibraryTitle": source_record["LibraryTitle"],
                            "TargetLibraryTitle": target_record["LibraryTitle"],
                            "SourceServerRelativeUrl": source_record["ServerRelativeUrl"],
                            "TargetServerRelativeUrl": target_record["ServerRelativeUrl"],
                        }
                    )

    if not compare_modified_dates:
        try:
            changed_modified_date_path.unlink()
        except FileNotFoundError:
            pass
        try:
            target_older_than_source_path.unlink()
        except FileNotFoundError:
            pass

    print("Finding source files missing from target...")
    missing = 0
    missing_diagnostic_count = 0
    with missing_path.open("w", encoding="utf-8", newline="") as missing_handle, missing_diagnostics_path.open(
        "w", encoding="utf-8", newline=""
    ) as missing_diagnostics_handle:
        missing_writer = csv.DictWriter(missing_handle, fieldnames=MISSING_IN_TARGET_COLUMNS, delimiter=CSV_OUTPUT_DELIMITER)
        missing_diagnostics_writer = csv.DictWriter(
            missing_diagnostics_handle,
            fieldnames=MISSING_DIAGNOSTIC_COLUMNS,
            delimiter=CSV_OUTPUT_DELIMITER,
        )
        missing_writer.writeheader()
        missing_diagnostics_writer.writeheader()
        for key, record in source_index.items():
            if key not in target_valid_seen:
                missing += 1
                lib_key = record.get("LibraryKey") or ""
                web_path = record.get("WebPath") or ""
                library_path = record.get("LibraryPath") or ""
                ensure_library_summary(library_stats, lib_key, web_path, library_path)["MissingInTarget"] += 1
                target_path_records = target_path_index.get(comparison_path_key(key), [])
                missing_details = missing_reason_details(record, target_path_records)
                missing_row = dict(record)
                missing_row.update(missing_details)
                missing_writer.writerow(missing_row)
                for diagnostic_row in missing_diagnostics(record, missing_details, target_path_records, target_records):
                    missing_diagnostics_writer.writerow(diagnostic_row)
                    missing_diagnostic_count += 1

    report_timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    html_summary_path = output_directory / f"{args.comparison_name}-summary-{report_timestamp}.html"

    summary = {
        "ComparisonName": args.comparison_name,
        "SourceCsv": str(source_csv),
        "TargetCsv": str(target_csv),
        "SourceRows": source_total_rows,
        "SourceFilteredRows": source_filtered_rows,
        "SourceExcludedRows": source_excluded_rows,
        "SourceUniqueKeys": len(source_index),
        "SourceDuplicateKeysIgnored": source_duplicate_keys,
        "TargetRows": target_total_rows,
        "TargetFilteredRows": target_filtered_rows,
        "TargetExcludedRows": target_excluded_rows,
        "TargetUniqueKeys": len(target_seen),
        "TargetDuplicateKeysIgnored": target_duplicate_keys,
        "MatchedKeys": matched,
        "MissingInTarget": missing,
        "ExtraInTarget": extra,
        "ExtraFoldersInTarget": len(extra_folder_candidates),
        "DifferentSize": different_size,
        "ChangedModifiedDate": changed_modified_date,
        "TargetOlderThanSource": target_older_than_source,
        "ChangedVersion": changed_version,
        "MissingDiagnostics": missing_diagnostic_count,
        "SizeToleranceBytes": args.size_tolerance_bytes,
        "ModifiedDateComparisonEnabled": compare_modified_dates,
        "ModifiedDateToleranceMinutes": args.modified_date_tolerance_minutes if compare_modified_dates else "",
        "SourceModifiedTimeZone": args.source_modified_time_zone if compare_modified_dates else "",
        "TargetModifiedTimeZone": args.target_modified_time_zone if compare_modified_dates else "",
        "ShareGateReplacementCharacter": SHAREGATE_REPLACEMENT_CHARACTER,
        "ComparisonKeyIncludesVersion": True,
        "DuplicateKeysCsv": str(duplicate_keys_path),
        "ExtraFoldersInTargetCsv": str(extra_folders_path),
        "MissingDiagnosticsCsv": str(missing_diagnostics_path),
        "ChangedModifiedDateCsv": str(changed_modified_date_path) if compare_modified_dates else "",
        "TargetOlderThanSourceCsv": str(target_older_than_source_path) if compare_modified_dates else "",
        "ChangedVersionCsv": str(changed_version_path),
        "HtmlSummary": str(html_summary_path),
        "OutputDirectory": str(output_directory),
    }

    with summary_path.open("w", encoding="utf-8", newline="") as summary_handle:
        writer = csv.DictWriter(summary_handle, fieldnames=list(summary.keys()), delimiter=CSV_OUTPUT_DELIMITER)
        writer.writeheader()
        writer.writerow(summary)

    with extra_folders_path.open("w", encoding="utf-8", newline="") as extra_folders_handle:
        writer = csv.DictWriter(extra_folders_handle, fieldnames=EXTRA_FOLDER_COLUMNS, delimiter=CSV_OUTPUT_DELIMITER)
        writer.writeheader()
        for entry in sorted(
            extra_folder_candidates.values(),
            key=lambda value: (
                str(value["TargetWebUrl"]).lower(),
                -int(value["Depth"] or 0),
                str(value["TargetFolderServerRelativeUrl"]).lower(),
            ),
        ):
            writer.writerow({column: entry.get(column, "") for column in EXTRA_FOLDER_COLUMNS})

    library_summary_path = output_directory / "LibrarySummary.csv"
    library_summary_rows = []
    with library_summary_path.open("w", encoding="utf-8", newline="") as library_summary_handle:
        writer = csv.DictWriter(library_summary_handle, fieldnames=LIBRARY_SUMMARY_COLUMNS, delimiter=CSV_OUTPUT_DELIMITER)
        writer.writeheader()
        for entry in sorted(
            library_stats.values(),
            key=lambda value: (
                -int(value["MissingInTarget"]) - int(value["ExtraInTarget"]) - int(value["DifferentSize"]) - int(value["ChangedModifiedDate"]) - int(value["TargetOlderThanSource"]) - int(value["ChangedVersion"]),
                str(value["WebPath"]).lower(),
                str(value["SourceLibraryTitle"] or value["TargetLibraryTitle"]).lower(),
            ),
        ):
            entry["MatchedFilesPercent"] = percent(entry["MatchedFiles"], entry["SourceFiles"])
            entry["MissingInTargetPercent"] = percent(entry["MissingInTarget"], entry["SourceFiles"])
            entry["ExtraInTargetPercent"] = percent(entry["ExtraInTarget"], entry["TargetFiles"])
            entry["DifferentSizePercent"] = percent(entry["DifferentSize"], entry["MatchedFiles"])
            entry["ChangedModifiedDatePercent"] = percent(entry["ChangedModifiedDate"], entry["MatchedFiles"])
            entry["TargetOlderThanSourcePercent"] = percent(entry["TargetOlderThanSource"], entry["MatchedFiles"])
            entry["ChangedVersionPercent"] = percent(entry["ChangedVersion"], entry["SourceFiles"])
            entry["Status"] = library_status(entry)
            row = {column: entry.get(column, "") for column in LIBRARY_SUMMARY_COLUMNS}
            library_summary_rows.append(row)
            writer.writerow(row)

    create_file_html_summary(
        html_summary_path,
        args.comparison_name,
        summary,
        library_summary_rows,
        [
            ("Summary", summary_path, "Run counters and comparison inputs."),
            ("Library summary", library_summary_path, "Library-level file comparison summary."),
            ("Missing in target", missing_path, "Source files absent from target or not usable at target."),
            ("Missing diagnostics", missing_diagnostics_path, "Diagnostic hints for missing files: wrong version, zero-byte target, relocated candidates, and ShareGate rename suspects."),
            ("Extra in target", extra_path, "Target files not found in source."),
            ("Different size", different_size_path, "Matched files with size differences beyond tolerance."),
            ("Changed modified date", changed_modified_date_path, "Matched files with modified-date differences beyond tolerance."),
            ("Target older than source", target_older_than_source_path, "Matched files where target is older than source."),
            ("Changed version", changed_version_path, "Same-path files whose SharePoint version differs."),
            ("Extra folders in target", extra_folders_path, "Target folders containing extra files."),
            ("Duplicate keys", duplicate_keys_path, "Source or target duplicate comparison keys ignored."),
        ],
    )

    print("Comparison completed.")
    print(
        f"Matched: {matched}; Missing in target: {missing}; "
        f"Extra in target: {extra}; Extra folders in target: {len(extra_folder_candidates)}; "
        f"Different size: {different_size}; Changed modified date: {changed_modified_date}; "
        f"Target older than source: {target_older_than_source}; Changed version: {changed_version}"
    )
    print(f"Summary: {summary_path}")
    print(f"Library summary: {library_summary_path}")
    print(f"Extra folders in target: {extra_folders_path}")
    print(f"Missing diagnostics: {missing_diagnostics_path}")
    print(f"Changed modified date: {changed_modified_date_path}")
    print(f"Target older than source: {target_older_than_source_path}")
    print(f"Changed version: {changed_version_path}")
    print(f"Duplicate keys: {duplicate_keys_path}")
    print(f"HTML summary: {html_summary_path}")


if __name__ == "__main__":
    main()
