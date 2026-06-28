import argparse
import builtins
import csv
import html
import re
import unicodedata
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if args:
        args = (f"{timestamp} {args[0]}", *args[1:])
    else:
        args = (timestamp,)
    builtins.print(*args, **kwargs)


CSV_DELIMITERS = ",;\t"
MAX_EXCEL_ROWS = 1_048_576
INVALID_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

PERMISSION_SUMMARY_COLUMNS = [
    "Status",
    "ObjectScope",
    "ComparisonObjectPath",
    "ObjectTitle",
    "ListTitle",
    "WebTitle",
    "SourcePermissions",
    "TargetPermissions",
    "MatchedPermissions",
    "MatchedPermissionsPercent",
    "MissingInSPO",
    "MissingInSPOPercent",
    "ExtraInSPO",
    "ExtraInSPOPercent",
    "SourceWebUrl",
    "TargetWebUrl",
    "SourceObjectUrl",
    "TargetObjectUrl",
]

COLUMN_WIDTHS = {
    "Status": 28,
    "ObjectScope": 16,
    "ComparisonObjectPath": 76,
    "ObjectTitle": 34,
    "ListTitle": 26,
    "WebTitle": 24,
    "SourcePermissions": 18,
    "TargetPermissions": 18,
    "MatchedPermissions": 20,
    "MatchedPermissionsPercent": 27,
    "MissingInSPO": 16,
    "MissingInSPOPercent": 22,
    "ExtraInSPO": 14,
    "ExtraInSPOPercent": 20,
    "SourceWebUrl": 58,
    "TargetWebUrl": 58,
    "SourceObjectUrl": 76,
    "TargetObjectUrl": 76,
    "SiteCollectionUrl": 58,
    "WebUrl": 58,
    "ObjectUrl": 76,
    "ObjectServerRelativeUrl": 76,
    "ParentObjectUrl": 76,
    "ListUrl": 76,
    "PrincipalName": 34,
    "PrincipalLoginName": 42,
    "PermissionLevels": 32,
    "ComparisonPrincipal": 42,
    "ComparisonPermissionLevels": 34,
}

PERMISSION_LEVEL_ALIASES = {
    "acces limite": "limited access",
    "accès limité": "limited access",
    "lecture": "read",
    "controle total": "full control",
    "contrôle total": "full control",
    "modification": "edit",
    "contribution": "contribute",
    "conception": "design",
    "lecture restreinte": "restricted view",
    "affichage seul": "view only",
    "approuver": "approve",
    "gerer la hierarchie": "manage hierarchy",
    "gérer la hiérarchie": "manage hierarchy",
}


def sniff_dialect(path: Path):
    sample = path.read_text(encoding="utf-8-sig", errors="replace")[:8192]
    try:
        return csv.Sniffer().sniff(sample, delimiters=CSV_DELIMITERS)
    except csv.Error:
        class SemiColon(csv.excel):
            delimiter = ";"
        return SemiColon


def read_rows(path: Path):
    dialect = sniff_dialect(path)
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as handle:
        yield from csv.DictReader(handle, dialect=dialect)


def write_csv(path: Path, rows, fieldnames):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def clean_cell(value):
    if value is None:
        return ""
    return INVALID_XML_RE.sub("", str(value))


def xlsx_col_name(index):
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def excel_dimension(row_count, column_count):
    if row_count <= 0 or column_count <= 0:
        return "A1"
    return f"A1:{xlsx_col_name(column_count)}{row_count}"


def column_width(header):
    if header in COLUMN_WIDTHS:
        return COLUMN_WIDTHS[header]
    return min(max(len(str(header or "")) + 3, 14), 80)


def sheet_xml(rows):
    row_count = len(rows)
    column_count = max((len(row) for row in rows), default=0)
    dimension = excel_dimension(row_count, column_count)
    headers = rows[0] if rows else []
    output = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        f'<dimension ref="{dimension}"/>',
    ]
    if column_count:
        output.append("<cols>")
        for column_index, header in enumerate(headers, start=1):
            width = column_width(header)
            output.append(f'<col min="{column_index}" max="{column_index}" width="{width}" customWidth="1"/>')
        output.append("</cols>")
    output.append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    output.append("<sheetData>")
    for row_index, row in enumerate(rows, start=1):
        output.append(f'<row r="{row_index}">')
        for column_index, value in enumerate(row, start=1):
            ref = f"{xlsx_col_name(column_index)}{row_index}"
            text = html.escape(clean_cell(value), quote=True)
            output.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
        output.append("</row>")
    output.append("</sheetData>")
    if row_count > 1 and column_count:
        output.append(f'<autoFilter ref="{dimension}"/>')
    output.append("</worksheet>")
    return "".join(output)


def create_xlsx(path: Path, sheets):
    path.parent.mkdir(parents=True, exist_ok=True)
    safe_sheets = [(name[:31], rows[:MAX_EXCEL_ROWS]) for name, rows in sheets]

    workbook_sheets = []
    workbook_rels = []
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    ]

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""")

        for index, (name, rows) in enumerate(safe_sheets, start=1):
            workbook_sheets.append(f'<sheet name="{html.escape(name, quote=True)}" sheetId="{index}" r:id="rId{index}"/>')
            workbook_rels.append(f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>')
            content_types.append(f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            archive.writestr(f"xl/worksheets/sheet{index}.xml", sheet_xml(rows))

        content_types.append("</Types>")
        archive.writestr("[Content_Types].xml", "".join(content_types))
        archive.writestr("xl/workbook.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{''.join(workbook_sheets)}</sheets></workbook>""")
        archive.writestr("xl/_rels/workbook.xml.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{''.join(workbook_rels)}</Relationships>""")


def normalize_path(value):
    if not value:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if re.match(r"^https?://", text, flags=re.I):
        parsed = urlparse(text)
        text = parsed.path
    text = re.sub(r"[?#].*$", "", text).strip()
    if not text.startswith("/"):
        text = "/" + text
    text = re.sub(r"/+", "/", text).rstrip("/")
    return text.lower() or "/"


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

            mappings.append((normalize_path(parts[0]), normalize_path(parts[1])))

    mappings.sort(key=lambda item: len(item[0]), reverse=True)
    return mappings


def replace_path_mapping(path, mappings):
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


def replace_path_prefix(path, source_prefix, target_prefix):
    source_prefix = normalize_path(source_prefix)
    target_prefix = normalize_path(target_prefix)
    if path == source_prefix:
        return target_prefix
    if path.startswith(source_prefix + "/"):
        return target_prefix + path[len(source_prefix):]
    return path


def normalize_principal(value):
    text = (value or "").strip().lower()
    if "|" in text:
        text = text.split("|")[-1]
    text = re.sub(r"^i:0[#.a-z0-9]*", "", text)
    return text


def normalize_label(value):
    text = (value or "").strip().lower()
    text = re.sub(r"\s+", " ", text)
    plain = "".join(
        character
        for character in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(character)
    )
    return PERMISSION_LEVEL_ALIASES.get(text, PERMISSION_LEVEL_ALIASES.get(plain, text))


def normalize_permissions(value):
    permissions = [normalize_label(part) for part in (value or "").split("|") if part.strip()]
    return "|".join(sorted(set(permissions)))


def row_path(row, source_prefix=None, target_prefix=None, path_mappings=None):
    raw = row.get("ObjectServerRelativeUrl") or row.get("ObjectUrl") or ""
    path = normalize_path(raw)
    if path_mappings:
        path = replace_path_mapping(path, path_mappings)
    elif source_prefix and target_prefix:
        path = replace_path_prefix(path, source_prefix, target_prefix)
    return path


def make_key(row, source_prefix=None, target_prefix=None, path_mappings=None):
    return (
        (row.get("ObjectScope") or "").strip().lower(),
        row_path(row, source_prefix, target_prefix, path_mappings=path_mappings),
        normalize_principal(row.get("PrincipalLoginName") or row.get("PrincipalName")),
        normalize_permissions(row.get("PermissionLevels")),
    )


def attach_key(row, key):
    new_row = dict(row)
    new_row["ComparisonObjectScope"] = key[0]
    new_row["ComparisonObjectPath"] = key[1]
    new_row["ComparisonPrincipal"] = key[2]
    new_row["ComparisonPermissionLevels"] = key[3]
    return new_row


def percent(numerator, denominator):
    denominator = int(denominator or 0)
    if denominator <= 0:
        return "0.000000"
    return f"{int(numerator or 0) / denominator:.6f}"


def permission_summary_status(entry):
    missing = int(entry.get("MissingInSPO") or 0)
    extra = int(entry.get("ExtraInSPO") or 0)
    if missing == 0 and extra == 0:
        return "OK - no difference"
    if missing and extra:
        return "Mixed differences: missing in SPO, extra in SPO"
    if missing:
        return "Missing in SPO"
    return "Extra in SPO"


def ensure_permission_summary(stats, object_scope, object_path):
    key = (object_scope or "", object_path or "")
    if key not in stats:
        stats[key] = {
            "Status": "",
            "ObjectScope": key[0],
            "ComparisonObjectPath": key[1],
            "ObjectTitle": "",
            "ListTitle": "",
            "WebTitle": "",
            "SourcePermissions": 0,
            "TargetPermissions": 0,
            "MatchedPermissions": 0,
            "MatchedPermissionsPercent": "0.000000",
            "MissingInSPO": 0,
            "MissingInSPOPercent": "0.000000",
            "ExtraInSPO": 0,
            "ExtraInSPOPercent": "0.000000",
            "SourceWebUrl": "",
            "TargetWebUrl": "",
            "SourceObjectUrl": "",
            "TargetObjectUrl": "",
        }
    return stats[key]


def update_permission_summary_metadata(entry, row, side):
    for field in ("ObjectTitle", "ListTitle", "WebTitle"):
        if not entry.get(field) and row.get(field):
            entry[field] = row.get(field, "")
    if side == "Source":
        if not entry.get("SourceWebUrl"):
            entry["SourceWebUrl"] = row.get("WebUrl", "")
        if not entry.get("SourceObjectUrl"):
            entry["SourceObjectUrl"] = row.get("ObjectUrl", "")
    else:
        if not entry.get("TargetWebUrl"):
            entry["TargetWebUrl"] = row.get("WebUrl", "")
        if not entry.get("TargetObjectUrl"):
            entry["TargetObjectUrl"] = row.get("ObjectUrl", "")


def build_permission_summary(source_by_key, target_by_key, matched_keys, missing_keys, extra_keys):
    stats = {}
    for row in source_by_key.values():
        entry = ensure_permission_summary(stats, row.get("ComparisonObjectScope"), row.get("ComparisonObjectPath"))
        entry["SourcePermissions"] += 1
        update_permission_summary_metadata(entry, row, "Source")
    for row in target_by_key.values():
        entry = ensure_permission_summary(stats, row.get("ComparisonObjectScope"), row.get("ComparisonObjectPath"))
        entry["TargetPermissions"] += 1
        update_permission_summary_metadata(entry, row, "Target")
    for key in matched_keys:
        object_scope, object_path = key[0], key[1]
        ensure_permission_summary(stats, object_scope, object_path)["MatchedPermissions"] += 1
    for key in missing_keys:
        object_scope, object_path = key[0], key[1]
        ensure_permission_summary(stats, object_scope, object_path)["MissingInSPO"] += 1
    for key in extra_keys:
        object_scope, object_path = key[0], key[1]
        ensure_permission_summary(stats, object_scope, object_path)["ExtraInSPO"] += 1

    rows = []
    for entry in stats.values():
        entry["MatchedPermissionsPercent"] = percent(entry["MatchedPermissions"], entry["SourcePermissions"])
        entry["MissingInSPOPercent"] = percent(entry["MissingInSPO"], entry["SourcePermissions"])
        entry["ExtraInSPOPercent"] = percent(entry["ExtraInSPO"], entry["TargetPermissions"])
        entry["Status"] = permission_summary_status(entry)
        rows.append(entry)
    return sorted(
        rows,
        key=lambda item: (
            -int(item.get("MissingInSPO") or 0) - int(item.get("ExtraInSPO") or 0),
            str(item.get("ObjectScope") or ""),
            str(item.get("ComparisonObjectPath") or ""),
        ),
    )


def list_rows_for_excel(path: Path, max_rows=MAX_EXCEL_ROWS):
    rows = list(read_rows(path))
    if not rows:
        return [[]]
    header = list(rows[0].keys())
    return [header] + [[row.get(column, "") for column in header] for row in rows[: max_rows - 1]]


def main():
    parser = argparse.ArgumentParser(description="Compare SharePoint permission inventories.")
    parser.add_argument("--source-csv", required=True)
    parser.add_argument("--target-csv", required=True)
    parser.add_argument("--output-directory", required=True)
    parser.add_argument("--source-root-path", default="/FR")
    parser.add_argument("--target-root-path", default="/")
    parser.add_argument("--path-mapping-file")
    parser.add_argument("--comparison-name", default="SP2019-vs-SPO-Permissions")
    args = parser.parse_args()

    source_csv = Path(args.source_csv)
    target_csv = Path(args.target_csv)
    output_dir = Path(args.output_directory)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_rows = list(read_rows(source_csv))
    target_rows = list(read_rows(target_csv))

    path_mappings = load_path_mappings(args.path_mapping_file)
    if path_mappings:
        print(f"Path mappings loaded: {len(path_mappings)}")

    source_by_key = {}
    target_by_key = {}
    source_duplicates = []
    target_duplicates = []

    for row in source_rows:
        key = make_key(row, args.source_root_path, args.target_root_path, path_mappings=path_mappings)
        keyed = attach_key(row, key)
        if key in source_by_key:
            source_duplicates.append(keyed)
            continue
        source_by_key[key] = keyed

    for row in target_rows:
        key = make_key(row)
        keyed = attach_key(row, key)
        if key in target_by_key:
            target_duplicates.append(keyed)
            continue
        target_by_key[key] = keyed

    source_keys = set(source_by_key)
    target_keys = set(target_by_key)
    matched_keys = sorted(source_keys & target_keys)
    missing_keys = sorted(source_keys - target_keys)
    extra_keys = sorted(target_keys - source_keys)

    summary_rows = [
        {
            "SourceCsv": str(source_csv),
            "TargetCsv": str(target_csv),
            "SourceRows": len(source_rows),
            "TargetRows": len(target_rows),
            "SourceUniqueKeys": len(source_by_key),
            "TargetUniqueKeys": len(target_by_key),
            "MatchedPermissions": len(matched_keys),
            "MissingInSPO": len(missing_keys),
            "ExtraInSPO": len(extra_keys),
            "SourceDuplicateKeysIgnored": len(source_duplicates),
            "TargetDuplicateKeysIgnored": len(target_duplicates),
            "SourceRootPath": args.source_root_path,
            "TargetRootPath": args.target_root_path,
        }
    ]

    def rows_from_keys(keys, lookup):
        return [lookup[key] for key in keys]

    matched_rows = rows_from_keys(matched_keys, source_by_key)
    missing_rows = rows_from_keys(missing_keys, source_by_key)
    extra_rows = rows_from_keys(extra_keys, target_by_key)

    scope_counter = Counter()
    for key in matched_keys:
        scope_counter[(key[0], "Matched")] += 1
    for key in missing_keys:
        scope_counter[(key[0], "MissingInSPO")] += 1
    for key in extra_keys:
        scope_counter[(key[0], "ExtraInSPO")] += 1
    scope_rows = [
        {"ObjectScope": scope, "Status": status, "Count": count}
        for (scope, status), count in sorted(scope_counter.items())
    ]
    permission_summary_rows = build_permission_summary(source_by_key, target_by_key, matched_keys, missing_keys, extra_keys)

    summary_csv = output_dir / "Summary.csv"
    scope_csv = output_dir / "ScopeSummary.csv"
    permission_summary_csv = output_dir / "PermissionSummary.csv"
    missing_csv = output_dir / "MissingInSPO.csv"
    extra_csv = output_dir / "ExtraInSPO.csv"
    matched_csv = output_dir / "Matched.csv"
    duplicate_source_csv = output_dir / "DuplicateKeys-Source.csv"
    duplicate_target_csv = output_dir / "DuplicateKeys-Target.csv"

    write_csv(summary_csv, summary_rows, list(summary_rows[0].keys()))
    write_csv(scope_csv, scope_rows, ["ObjectScope", "Status", "Count"])
    write_csv(permission_summary_csv, permission_summary_rows, PERMISSION_SUMMARY_COLUMNS)

    output_fields = list(source_rows[0].keys()) if source_rows else []
    for extra_field in ["ComparisonObjectScope", "ComparisonObjectPath", "ComparisonPrincipal", "ComparisonPermissionLevels"]:
        if extra_field not in output_fields:
            output_fields.append(extra_field)
    write_csv(missing_csv, missing_rows, output_fields)
    write_csv(matched_csv, matched_rows, output_fields)

    target_fields = list(target_rows[0].keys()) if target_rows else output_fields
    for extra_field in ["ComparisonObjectScope", "ComparisonObjectPath", "ComparisonPrincipal", "ComparisonPermissionLevels"]:
        if extra_field not in target_fields:
            target_fields.append(extra_field)
    write_csv(extra_csv, extra_rows, target_fields)
    write_csv(duplicate_source_csv, source_duplicates, output_fields)
    write_csv(duplicate_target_csv, target_duplicates, target_fields)

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    xlsx_path = output_dir / f"{args.comparison_name}-{timestamp}.xlsx"
    create_xlsx(
        xlsx_path,
        [
            ("Summary", list_rows_for_excel(summary_csv)),
            ("PermissionSummary", list_rows_for_excel(permission_summary_csv)),
            ("ScopeSummary", list_rows_for_excel(scope_csv)),
            ("MissingInSPO", list_rows_for_excel(missing_csv)),
            ("ExtraInSPO", list_rows_for_excel(extra_csv)),
            ("Matched", list_rows_for_excel(matched_csv)),
            ("DuplicateSource", list_rows_for_excel(duplicate_source_csv)),
            ("DuplicateTarget", list_rows_for_excel(duplicate_target_csv)),
        ],
    )

    print(f"Comparison completed.")
    print(f"Matched permissions: {len(matched_keys)}")
    print(f"Missing in SPO: {len(missing_keys)}")
    print(f"Extra in SPO: {len(extra_keys)}")
    print(f"Summary: {summary_csv}")
    print(f"Excel: {xlsx_path}")


if __name__ == "__main__":
    main()
