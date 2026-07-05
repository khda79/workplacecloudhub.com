import argparse
import builtins
import csv
import html
import re
import unicodedata
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, unquote, urlparse


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
SHAREGATE_REPLACED_CHARACTERS = {"&", "#", "%", '"', "*", ":", ";", "<", ">", "?", "\\", "|", "{", "}", "~"}
SHAREGATE_REPLACEMENT_CHARACTER = "_"
PERCENT_SEQUENCE_PATTERN = re.compile(r"%([0-9A-Fa-f]{2})")
DEFAULT_DOCUMENT_LIBRARY_SEGMENTS = {
    "documents",
    "documents partages",
    "shared documents",
}

PERMISSION_SUMMARY_COLUMNS = [
    "Status",
    "SourcePermissions",
    "TargetPermissions",
    "MatchedPermissions",
    "MatchedPermissionsPercent",
    "MissingInSPO",
    "MissingInSPOPercent",
    "DisabledEntraUsersNotInSPO",
    "DisabledEntraUsersNotInSPOPercent",
    "ExtraInSPO",
    "ExtraInSPOPercent",
    "SourceWebUrl",
    "TargetWebUrl",
    "ObjectScope",
    "ComparisonObjectPath",
    "ObjectTitle",
    "ListTitle",
    "WebTitle",
    "SourceObjectUrl",
    "TargetObjectUrl",
]

NUMERIC_COLUMNS = {
    "SourceRows",
    "TargetRows",
    "SourceUniqueKeys",
    "TargetUniqueKeys",
    "MatchedPermissions",
    "MissingInSPO",
    "ExtraInSPO",
    "SourceDuplicateKeysIgnored",
    "TargetDuplicateKeysIgnored",
    "SourceLimitedAccessOnlyIgnored",
    "TargetLimitedAccessOnlyIgnored",
    "EntraUserAliasesLoaded",
    "DisabledEntraUsersLoaded",
    "SourceUsersNotInEntraIgnored",
    "DisabledEntraUsersNotInSPO",
    "SharePointGroupMappingsLoaded",
    "SourceRowsWithoutLibraryScopeMetadata",
    "TargetRowsWithoutLibraryScopeMetadata",
    "SourcePermissions",
    "TargetPermissions",
    "Count",
}
PERCENT_COLUMNS = {
    "MatchedPermissionsPercent",
    "MissingInSPOPercent",
    "DisabledEntraUsersNotInSPOPercent",
    "ExtraInSPOPercent",
}
INTEGER_RE = re.compile(r"^-?\d+$")
DECIMAL_RE = re.compile(r"^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$")

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
    "DisabledEntraUsersNotInSPO": 28,
    "DisabledEntraUsersNotInSPOPercent": 34,
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
    "MappingSourceWebPath": 58,
    "MappingTargetWebPath": 58,
    "MappingSourceObjectPath": 76,
    "MappingRole": 14,
    "MappingSourcePrincipalName": 34,
    "MappingTargetPrincipalName": 34,
}

PERMISSION_LEVEL_ALIASES = {
    "acces limite": "limited access",
    "accès limité": "limited access",
    "lecture": "read",
    "controle total": "full control",
    "contrôle total": "full control",
    "modification": "edit",
    "contribution": "contribute",
    "collaboration": "contribute",
    "conception": "design",
    "lecture restreinte": "restricted read",
    "affichage seul": "view only",
    "approuver": "approve",
    "gerer la hierarchie": "manage hierarchy",
    "gérer la hiérarchie": "manage hierarchy",
    "gestion de la hierarchie": "manage hierarchy",
    "gestion de la hiérarchie": "manage hierarchy",
    "interfaces restreintes pour la traduction": "restricted interfaces for translation",
}

SHAREPOINT_ASSOCIATED_GROUP_SUFFIXES = {
    "owners": "owners",
    "owner": "owners",
    "proprietaires": "owners",
    "propriétaires": "owners",
    "members": "members",
    "member": "members",
    "membres": "members",
    "visitors": "visitors",
    "visitor": "visitors",
    "visiteurs": "visitors",
}

EXPLICIT_SHAREPOINT_GROUP_MAPPINGS = [
    {
        "source_web_path": "/sites/corp-newrealestate",
        "target_web_path": "/sites/corp-newrealestate",
        "role": "owners",
        "source_group_name": "Real Estate Owners - do not delete",
        "target_group_name": "CORP-newrealestate Owners",
        "method": "ConfirmedManualMapping",
    },
    {
        "source_web_path": "/sites/corp-finance",
        "target_web_path": "/sites/corp-finance",
        "role": "owners",
        "source_group_name": "Finance CORP - Owners",
        "target_group_name": "CORP-Finance Owners",
        "source_object_path": "/sites/corp-finance/documents/financial controlling/ifrs 16 leasing/training presentation",
        "method": "ConfirmedManualMapping",
    },
]


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


def split_identity_values(value):
    if not value:
        return []
    values = []
    for part in re.split(r"[;,]", str(value)):
        text = part.strip()
        if not text:
            continue
        if ":" in text and text.split(":", 1)[0].lower() in {"smtp", "sip"}:
            text = text.split(":", 1)[1].strip()
        if text:
            values.append(text)
    return values


def add_entra_alias(alias_map, alias, canonical):
    alias_key = normalize_principal_text(alias)
    canonical_value = normalize_principal_text(canonical)
    if not alias_key or not canonical_value:
        return
    alias_map.setdefault(alias_key, canonical_value)


def load_entra_user_aliases(path):
    if not path:
        return {}, set()

    csv_path = Path(path)
    if not csv_path.exists():
        raise FileNotFoundError(f"Entra users cache not found: {csv_path}")

    alias_map = {}
    disabled_canonical_users = set()
    for row in read_rows(csv_path):
        canonical = row.get("UserPrincipalName") or row.get("User principal name") or row.get("Mail") or row.get("mail")
        if not canonical:
            continue

        identity_values = [
            canonical,
            row.get("Mail"),
            row.get("mail"),
            row.get("OnPremisesUserPrincipalName"),
            row.get("onPremisesUserPrincipalName"),
        ]
        for field in ("OtherMails", "otherMails"):
            identity_values.extend(split_identity_values(row.get(field)))

        for value in identity_values:
            add_entra_alias(alias_map, value, canonical)

        account_enabled = row.get("AccountEnabled")
        if account_enabled is None:
            account_enabled = row.get("accountEnabled")
        if parse_bool(account_enabled) is False:
            disabled_canonical_users.add(normalize_principal_text(canonical))

    return alias_map, disabled_canonical_users


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


def cell_xml(row_index, column_index, value, header=None, is_header=False):
    ref = f"{xlsx_col_name(column_index)}{row_index}"
    text = clean_cell(value).strip()
    if is_header:
        escaped = html.escape(text, quote=True)
        return f'<c r="{ref}" t="inlineStr" s="1"><is><t>{escaped}</t></is></c>'
    if header in PERCENT_COLUMNS and DECIMAL_RE.fullmatch(text):
        return f'<c r="{ref}" s="2"><v>{text}</v></c>'
    if header in NUMERIC_COLUMNS and INTEGER_RE.fullmatch(text):
        return f'<c r="{ref}"><v>{text}</v></c>'
    escaped = html.escape(text, quote=True)
    return f'<c r="{ref}" t="inlineStr"><is><t>{escaped}</t></is></c>'


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
    output.append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    if column_count:
        output.append("<cols>")
        for column_index, header in enumerate(headers, start=1):
            width = column_width(header)
            output.append(f'<col min="{column_index}" max="{column_index}" width="{width}" customWidth="1"/>')
        output.append("</cols>")
    output.append("<sheetData>")
    for row_index, row in enumerate(rows, start=1):
        output.append(f'<row r="{row_index}">')
        for column_index, value in enumerate(row, start=1):
            header = headers[column_index - 1] if column_index <= len(headers) else ""
            output.append(cell_xml(row_index, column_index, value, header=header, is_header=(row_index == 1)))
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
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""")

        for index, (name, rows) in enumerate(safe_sheets, start=1):
            workbook_sheets.append(f'<sheet name="{html.escape(name, quote=True)}" sheetId="{index}" r:id="rId{index}"/>')
            workbook_rels.append(f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>')
            content_types.append(f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            archive.writestr(f"xl/worksheets/sheet{index}.xml", sheet_xml(rows))

        workbook_rels.append('<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
        content_types.append("</Types>")
        archive.writestr("[Content_Types].xml", "".join(content_types))
        archive.writestr("xl/workbook.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{''.join(workbook_sheets)}</sheets></workbook>""")
        archive.writestr("xl/_rels/workbook.xml.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{''.join(workbook_rels)}</Relationships>""")
        archive.writestr("xl/styles.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0078D4"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFDDE7F0"/></left><right style="thin"><color rgb="FFDDE7F0"/></right><top style="thin"><color rgb="FFDDE7F0"/></top><bottom style="thin"><color rgb="FFDDE7F0"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>""")


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


def top_permission_difference_rows(permission_summary_rows, limit=20):
    rows = [
        row for row in permission_summary_rows
        if to_int(row.get("MissingInSPO")) or to_int(row.get("DisabledEntraUsersNotInSPO")) or to_int(row.get("ExtraInSPO"))
    ]
    rows.sort(
        key=lambda row: (
            -(to_int(row.get("MissingInSPO")) + to_int(row.get("DisabledEntraUsersNotInSPO")) + to_int(row.get("ExtraInSPO"))),
            str(row.get("ComparisonObjectPath") or ""),
        )
    )
    return rows[:limit]


def create_permission_html_summary(path, title, summary_row, permission_summary_rows, scope_rows, report_links, source_csv=None, target_csv=None):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    matched = to_int(summary_row.get("MatchedPermissions"))
    missing = to_int(summary_row.get("MissingInSPO"))
    disabled_missing = to_int(summary_row.get("DisabledEntraUsersNotInSPO"))
    extra = to_int(summary_row.get("ExtraInSPO"))
    not_entra = to_int(summary_row.get("SourceUsersNotInEntraIgnored"))
    source_la = to_int(summary_row.get("SourceLimitedAccessOnlyIgnored"))
    target_la = to_int(summary_row.get("TargetLimitedAccessOnlyIgnored"))
    real_difference_count = missing + extra
    status_text = "Review needed" if real_difference_count or disabled_missing else "No relevant difference"
    status_class = "warn" if real_difference_count else ("note" if disabled_missing else "ok")

    cards = [
        ("Matched", matched, "ok"),
        ("Missing in SPO", missing, "bad" if missing else "ok"),
        ("Disabled Entra users not in SPO", disabled_missing, "note" if disabled_missing else "ok"),
        ("Extra in SPO", extra, "bad" if extra else "ok"),
        ("Source users not in Entra", not_entra, "note" if not_entra else "ok"),
        ("Limited Access ignored", source_la + target_la, "muted"),
    ]
    card_html = [
        f'<div class="metric {css_class}"><div class="metric-label">{html_escape(label)}</div><div class="metric-value">{format_integer(value)}</div></div>'
        for label, value, css_class in cards
    ]

    source_csv_html = html_escape(source_csv) if source_csv else html_escape(summary_row.get("SourceCsv"))
    target_csv_html = html_escape(target_csv) if target_csv else html_escape(summary_row.get("TargetCsv"))
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    warning = summary_row.get("ScopeWarning") or ""
    warning_html = f'<div class="callout warn"><strong>Warning</strong><br>{html_escape(warning)}</div>' if warning else ""

    report_link_rows = []
    for label, target_path, description in report_links:
        report_link_rows.append(
            "<tr>"
            f"<td>{html_escape(label)}</td>"
            f"<td>{html_file_link(target_path)}</td>"
            f"<td>{html_escape(description)}</td>"
            "</tr>"
        )

    top_rows_html = []
    for row in top_permission_difference_rows(permission_summary_rows):
        top_rows_html.append(
            "<tr>"
            f"<td>{html_escape(row.get('Status'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('MissingInSPO'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('DisabledEntraUsersNotInSPO'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('ExtraInSPO'))}</td>"
            f"<td>{html_escape(row.get('ObjectScope'))}</td>"
            f"<td>{html_escape(row.get('ComparisonObjectPath'))}</td>"
            f"<td>{html_escape(row.get('ObjectTitle') or row.get('ListTitle') or row.get('WebTitle'))}</td>"
            "</tr>"
        )
    if not top_rows_html:
        top_rows_html.append('<tr><td colspan="7" class="empty">No object-level differences.</td></tr>')

    scope_rows_html = []
    for row in scope_rows:
        scope_rows_html.append(
            "<tr>"
            f"<td>{html_escape(row.get('ObjectScope'))}</td>"
            f"<td>{html_escape(row.get('Status'))}</td>"
            f"<td class=\"num\">{format_integer(row.get('Count'))}</td>"
            "</tr>"
        )
    if not scope_rows_html:
        scope_rows_html.append('<tr><td colspan="3" class="empty">No scope summary rows.</td></tr>')

    document = f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{html_escape(title)}</title>
<style>
:root {{ color-scheme: light; --bg:#F5F8FB; --card:#FFFFFF; --text:#1F2937; --muted:#5F6B7A; --line:#DDE7F0; --accent:#0078D4; --ok:#107C10; --bad:#C50F1F; --note:#8A6A00; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--text); font-family:"Segoe UI", Arial, sans-serif; font-size:14px; line-height:1.45; }}
main {{ max-width:1280px; margin:0 auto; padding:28px; }}
.header {{ background:var(--card); border:1px solid var(--line); border-radius:8px; padding:22px 24px; margin-bottom:18px; display:flex; justify-content:space-between; gap:18px; align-items:flex-start; }}
h1 {{ margin:0 0 6px; font-size:25px; font-weight:650; letter-spacing:0; }}
.subtitle {{ color:var(--muted); }}
.badge {{ display:inline-block; border:1px solid var(--line); border-radius:999px; padding:5px 10px; font-weight:600; background:#fff; white-space:nowrap; }}
.badge.ok {{ color:var(--ok); border-color:#B8DAB8; background:#F1FAF1; }}
.badge.warn {{ color:var(--bad); border-color:#F1B7BC; background:#FFF4F5; }}
.badge.note {{ color:var(--note); border-color:#E7D99B; background:#FFF9DF; }}
.metrics {{ display:grid; grid-template-columns:repeat(6, minmax(140px, 1fr)); gap:12px; margin-bottom:18px; }}
.metric {{ background:var(--card); border:1px solid var(--line); border-radius:8px; padding:15px; min-height:92px; }}
.metric-label {{ color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }}
.metric-value {{ font-size:28px; font-weight:700; margin-top:8px; }}
.metric.ok .metric-value {{ color:var(--ok); }} .metric.bad .metric-value {{ color:var(--bad); }} .metric.note .metric-value {{ color:var(--note); }} .metric.muted .metric-value {{ color:var(--muted); }}
.section {{ background:var(--card); border:1px solid var(--line); border-radius:8px; padding:18px; margin-bottom:18px; }}
h2 {{ margin:0 0 12px; font-size:17px; }}
.grid {{ display:grid; grid-template-columns:190px minmax(0, 1fr); gap:8px 14px; }}
.key {{ color:var(--muted); }}
a {{ color:var(--accent); text-decoration:none; }} a:hover {{ text-decoration:underline; }}
table {{ width:100%; border-collapse:collapse; }}
th, td {{ border-bottom:1px solid var(--line); padding:9px 10px; text-align:left; vertical-align:top; }}
th {{ background:#F8FBFE; color:#334155; font-size:12px; text-transform:uppercase; letter-spacing:.04em; }}
.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
.empty {{ color:var(--muted); text-align:center; padding:18px; }}
.callout {{ border-radius:8px; padding:12px 14px; margin-bottom:18px; }}
.callout.warn {{ border:1px solid #E7D99B; background:#FFF9DF; color:#5F4B00; }}
.footer {{ color:var(--muted); font-size:12px; margin-top:16px; }}
@media (max-width:1000px) {{ .metrics {{ grid-template-columns:repeat(2, minmax(140px, 1fr)); }} .header {{ display:block; }} .badge {{ margin-top:12px; }} }}
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
  {warning_html}
  <div class="metrics">{''.join(card_html)}</div>
  <div class="section">
    <h2>Run context</h2>
    <div class="grid">
      <div class="key">Source CSV</div><div>{source_csv_html}</div>
      <div class="key">Target CSV</div><div>{target_csv_html}</div>
      <div class="key">Source root path</div><div>{html_escape(summary_row.get('SourceRootPath'))}</div>
      <div class="key">Target root path</div><div>{html_escape(summary_row.get('TargetRootPath'))}</div>
      <div class="key">Source rows</div><div>{format_integer(summary_row.get('SourceRows'))}</div>
      <div class="key">Target rows</div><div>{format_integer(summary_row.get('TargetRows'))}</div>
    </div>
  </div>
  <div class="section">
    <h2>Report files</h2>
    <table><thead><tr><th>Report</th><th>File</th><th>Description</th></tr></thead><tbody>{''.join(report_link_rows)}</tbody></table>
  </div>
  <div class="section">
    <h2>Top objects with differences</h2>
    <table><thead><tr><th>Status</th><th>Missing</th><th>Disabled users</th><th>Extra</th><th>Scope</th><th>Object path</th><th>Title</th></tr></thead><tbody>{''.join(top_rows_html)}</tbody></table>
  </div>
  <div class="section">
    <h2>Scope summary</h2>
    <table><thead><tr><th>Object scope</th><th>Status</th><th>Count</th></tr></thead><tbody>{''.join(scope_rows_html)}</tbody></table>
  </div>
  <div class="footer">SmartM365 SharePoint migration permission comparison summary.</div>
</main>
</body>
</html>
'''
    path.write_text(document, encoding="utf-8", newline="\n")
    return path


def strip_query_from_path(value):
    text = str(value or "").strip()
    if text.lower().startswith(("http://", "https://")):
        return urlparse(text).path
    return text.split("?", 1)[0]


def replace_encoded_sharegate_character(match):
    try:
        character = bytes.fromhex(match.group(1)).decode("utf-8")
    except UnicodeDecodeError:
        return match.group(0)

    if character in SHAREGATE_REPLACED_CHARACTERS or character in {"/", "\\"} or character == "\t" or not character.isprintable():
        return SHAREGATE_REPLACEMENT_CHARACTER + match.group(1)

    return match.group(0)


def decode_sharegate_path(path):
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
    return "/".join(normalize_sharegate_segment(segment) for segment in path.split("/"))


def normalize_path(value):
    if not value:
        return ""
    text = normalize_sharegate_path_characters(decode_sharegate_path(strip_query_from_path(value)))
    if not text.startswith("/"):
        text = "/" + text
    text = re.sub(r"/+", "/", text).rstrip("/") or "/"
    return text.lower()


def normalize_mapping_path(value):
    return normalize_path(value)


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

            mappings.append((normalize_mapping_path(parts[0]), normalize_mapping_path(parts[1])))

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


def replace_path_prefix(path, source_prefix, target_prefix):
    source_prefix = normalize_path(source_prefix)
    target_prefix = normalize_path(target_prefix)
    if path == source_prefix:
        return target_prefix
    if path.startswith(source_prefix + "/"):
        return target_prefix + path[len(source_prefix):]
    return path


def row_web_path(row, source_prefix=None, target_prefix=None, path_mappings=None):
    web_path = normalize_path(row.get("WebUrl"))
    if path_mappings:
        web_path = apply_path_mappings(web_path, path_mappings)
    elif source_prefix and target_prefix:
        web_path = replace_path_prefix(web_path, source_prefix, target_prefix)
    return web_path


def normalize_default_document_library_key(row, key, source_prefix=None, target_prefix=None, path_mappings=None):
    if not key:
        return key

    web_key = row_web_path(row, source_prefix, target_prefix, path_mappings=path_mappings) or ""
    web_prefix = web_key.rstrip("/")
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

def is_web_relative_path(value):
    text = (value or "").strip()
    if not text:
        return False
    parsed = urlparse(text)
    if parsed.scheme and parsed.netloc:
        return False
    return not text.startswith(("/", "\\"))


def normalize_web_relative_key(row, key, source_prefix=None, target_prefix=None, path_mappings=None):
    if not key:
        return key
    raw = row.get("ObjectServerRelativeUrl") or row.get("ObjectUrl") or ""
    if not is_web_relative_path(raw):
        return key

    web_key = row_web_path(row, source_prefix, target_prefix, path_mappings=path_mappings)
    if not web_key or web_key == "/" or key == web_key or key.startswith(web_key.rstrip("/") + "/"):
        return key
    return f"{web_key.rstrip('/')}/{key.lstrip('/')}"


def normalize_events_list_key(key):
    if not key or re.search(r"/lists/events($|/)", key):
        return key
    return re.sub(r"^(/sites/[^/]+|/[^/]+)/events($|/)", r"\1/lists/events\2", key, count=1)

def normalize_principal_text(value):
    text = (value or "").strip().lower()
    if "|" in text:
        text = text.split("|")[-1]
    text = re.sub(r"^i:0[#.a-z0-9]*", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_sharepoint_group_name(value):
    text = normalize_label(value)
    prefix_patterns = [
        ("proprietaires de ", "owners"),
        ("propriétaires de ", "owners"),
        ("owners of ", "owners"),
        ("membres de ", "members"),
        ("members of ", "members"),
        ("visiteurs de ", "visitors"),
        ("visitors of ", "visitors"),
    ]
    for prefix, role in prefix_patterns:
        if text.startswith(prefix):
            return text[len(prefix) :].strip(" -"), role

    match = re.search(r"(?:\s+-\s+|\s+)([^\s-]+)$", text)
    if not match:
        return None, None
    suffix = SHAREPOINT_ASSOCIATED_GROUP_SUFFIXES.get(match.group(1))
    if not suffix:
        return None, None
    base = text[: match.start()].strip(" -")
    return base, suffix


def associated_group_values(row):
    return {
        "members": row.get("AssociatedMemberGroup"),
        "owners": row.get("AssociatedOwnerGroup"),
        "visitors": row.get("AssociatedVisitorGroup"),
    }


def group_role_from_permission_row(row):
    if normalize_label(row.get("PrincipalType")) != "sharepointgroup":
        return None
    if normalize_label(row.get("ObjectScope")) != "web":
        return None
    _, role = parse_sharepoint_group_name(row.get("PrincipalName"))
    return role


def default_group_candidate_score(row, web_path):
    group_base, role = parse_sharepoint_group_name(row.get("PrincipalName"))
    if not role:
        return -1
    base = canonical_group_base(group_base)
    candidates = web_group_base_candidates(row, web_path)
    if base in candidates:
        return 1000
    if any(base == f"{candidate}corp" or base == f"corp{candidate}" for candidate in candidates):
        return 900
    if any(base.startswith(candidate) or candidate.startswith(base) for candidate in candidates):
        return 600 - min(len(base), 200)
    if any(candidate and candidate in base for candidate in candidates):
        return 250 - min(len(base), 200)
    return 0


def add_default_group_candidate(groups, row, web_path):
    role = group_role_from_permission_row(row)
    if not role or not web_path:
        return
    score = default_group_candidate_score(row, web_path)
    current = groups[web_path].get(role)
    if current is None or score > current[0]:
        groups[web_path][role] = (score, row.get("PrincipalName"))


def target_principal_for_group(target_web_row, target_web_path, target_group_name):
    fake_row = dict(target_web_row or {})
    fake_row["PrincipalType"] = "SharePointGroup"
    fake_row["PrincipalName"] = target_group_name
    fake_row["PrincipalLoginName"] = target_group_name
    return normalize_principal(fake_row, web_path=target_web_path)


def add_sharepoint_group_mapping(mappings, rows, source_web_path, target_web_path, role, source_group_name, target_group_name, method, target_web_row, target_group_names_by_web=None, source_object_path=None):
    source_group_name = (source_group_name or "").strip()
    target_group_name = (target_group_name or "").strip()
    if not source_web_path or not target_web_path or not role or not source_group_name or not target_group_name:
        return

    source_principal = normalize_principal_text(source_group_name)
    source_object_path = normalize_path(source_object_path) if source_object_path else ""
    if method != "ConfirmedManualMapping" and source_principal in (target_group_names_by_web or {}).get(target_web_path, set()):
        return
    target_principal = target_principal_for_group(target_web_row, target_web_path, target_group_name)
    key = (source_web_path, source_object_path, source_principal) if source_object_path else (source_web_path, source_principal)
    if key in mappings:
        return

    mappings[key] = target_principal
    rows.append(
        {
            "MappingSourceWebPath": source_web_path,
            "MappingTargetWebPath": target_web_path,
            "MappingSourceObjectPath": source_object_path,
            "MappingRole": role,
            "MappingSourcePrincipalName": source_group_name,
            "MappingTargetPrincipalName": target_group_name,
            "MappingTargetComparisonPrincipal": target_principal,
            "MappingMethod": method,
        }
    )


def build_sharepoint_group_mappings(source_rows, target_rows, source_prefix=None, target_prefix=None, path_mappings=None):
    mappings = {}
    mapping_rows = []
    source_web_rows = {}
    target_web_rows = {}
    target_group_names_by_web = defaultdict(set)
    source_associated = defaultdict(dict)
    target_associated = defaultdict(dict)

    for row in source_rows:
        source_web_path = row_web_path(row, source_prefix, target_prefix, path_mappings=path_mappings)
        if not source_web_path:
            continue
        if normalize_label(row.get("ObjectScope")) == "web":
            source_web_rows.setdefault(source_web_path, row)
        for role, group_name in associated_group_values(row).items():
            if group_name and role not in source_associated[source_web_path]:
                source_associated[source_web_path][role] = group_name

    for row in target_rows:
        target_web_path = row_web_path(row)
        if not target_web_path:
            continue
        if normalize_label(row.get("ObjectScope")) == "web":
            target_web_rows.setdefault(target_web_path, row)
        if normalize_label(row.get("PrincipalType")) == "sharepointgroup":
            target_group_names_by_web[target_web_path].add(normalize_principal_text(row.get("PrincipalName") or row.get("PrincipalLoginName")))
        for role, group_name in associated_group_values(row).items():
            if group_name and role not in target_associated[target_web_path]:
                target_associated[target_web_path][role] = group_name

    for explicit_mapping in EXPLICIT_SHAREPOINT_GROUP_MAPPINGS:
        add_sharepoint_group_mapping(
            mappings,
            mapping_rows,
            explicit_mapping["source_web_path"],
            explicit_mapping["target_web_path"],
            explicit_mapping["role"],
            explicit_mapping["source_group_name"],
            explicit_mapping["target_group_name"],
            explicit_mapping["method"],
            target_web_rows.get(explicit_mapping["target_web_path"]),
            target_group_names_by_web=target_group_names_by_web,
            source_object_path=explicit_mapping.get("source_object_path"),
        )
    for source_web_path, role_map in source_associated.items():
        target_web_path = source_web_path
        target_role_map = target_associated.get(target_web_path, {})
        target_web_row = target_web_rows.get(target_web_path)
        for role, source_group_name in role_map.items():
            add_sharepoint_group_mapping(
                mappings,
                mapping_rows,
                source_web_path,
                target_web_path,
                role,
                source_group_name,
                target_role_map.get(role),
                "AssociatedWebGroup",
                target_web_row,
                target_group_names_by_web=target_group_names_by_web,
            )

    source_web_role_groups = defaultdict(dict)
    target_web_role_groups = defaultdict(dict)
    for row in source_rows:
        add_default_group_candidate(
            source_web_role_groups,
            row,
            row_web_path(row, source_prefix, target_prefix, path_mappings=path_mappings),
        )
    for row in target_rows:
        add_default_group_candidate(target_web_role_groups, row, row_web_path(row))

    for source_web_path, role_map in source_web_role_groups.items():
        target_web_path = source_web_path
        target_role_map = target_web_role_groups.get(target_web_path, {})
        target_web_row = target_web_rows.get(target_web_path)
        for role, source_candidate in role_map.items():
            target_candidate = target_role_map.get(role)
            add_sharepoint_group_mapping(
                mappings,
                mapping_rows,
                source_web_path,
                target_web_path,
                role,
                source_candidate[1],
                target_candidate[1] if target_candidate else None,
                "WebPermissionFallback",
                target_web_row,
                target_group_names_by_web=target_group_names_by_web,
            )

    return mappings, mapping_rows

def canonical_group_base(value):
    text = normalize_label(value)
    text = re.sub(r"^corp[\s-]+", "", text)
    return re.sub(r"[^a-z0-9]+", "", text)


def web_group_base_candidates(row, web_path):
    candidates = {canonical_group_base(row.get("WebTitle"))}
    web_segment = (web_path or "").rstrip("/").split("/")[-1]
    if web_segment:
        candidates.add(canonical_group_base(web_segment))
        candidates.add(canonical_group_base(re.sub(r"^corp-", "", web_segment, flags=re.IGNORECASE)))
    return {candidate for candidate in candidates if candidate}


def normalize_principal(row, web_path="", object_path="", entra_user_aliases=None, sharepoint_group_mappings=None):
    principal = row.get("PrincipalLoginName") or row.get("PrincipalName")
    normalized = normalize_principal_text(principal)
    principal_type = normalize_label(row.get("PrincipalType"))
    if principal_type == "sharepointgroup":
        if sharepoint_group_mappings:
            mapped_principal = sharepoint_group_mappings.get((web_path, object_path, normalized)) or sharepoint_group_mappings.get((web_path, normalized))
            if mapped_principal:
                return mapped_principal
        group_base, suffix = parse_sharepoint_group_name(normalized)
        if suffix and canonical_group_base(group_base) in web_group_base_candidates(row, web_path):
            return f"sharepointgroup:{web_path}:{suffix}"
    if principal_type == "domaingroup" and normalized == "true" and normalize_label(row.get("PrincipalName")) in {"everyone", "tout le monde"}:
        return "everyone"
    if principal_type == "user" and entra_user_aliases:
        return entra_user_aliases.get(normalized, normalized)
    return normalized


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
    permissions = [permission for permission in permissions if permission != "limited access"]
    return "|".join(sorted(set(permissions)))


def row_path(row, source_prefix=None, target_prefix=None, path_mappings=None):
    raw = row.get("ObjectServerRelativeUrl") or row.get("ObjectUrl") or ""
    path = normalize_path(raw)
    if path_mappings:
        path = apply_path_mappings(path, path_mappings)
    elif source_prefix and target_prefix:
        path = replace_path_prefix(path, source_prefix, target_prefix)
    path = normalize_web_relative_key(row, path, source_prefix, target_prefix, path_mappings=path_mappings)
    path = normalize_events_list_key(path)
    return normalize_default_document_library_key(row, path, source_prefix, target_prefix, path_mappings=path_mappings)


def make_key(row, source_prefix=None, target_prefix=None, path_mappings=None, entra_user_aliases=None, sharepoint_group_mappings=None):
    web_path = row_web_path(row, source_prefix, target_prefix, path_mappings=path_mappings)
    object_path = row_path(row, source_prefix, target_prefix, path_mappings=path_mappings)
    return (
        (row.get("ObjectScope") or "").strip().lower(),
        object_path,
        normalize_principal(row, web_path=web_path, object_path=object_path, entra_user_aliases=entra_user_aliases, sharepoint_group_mappings=sharepoint_group_mappings),
        normalize_permissions(row.get("PermissionLevels")),
    )

def attach_key(row, key):
    new_row = dict(row)
    new_row["ComparisonObjectScope"] = key[0]
    new_row["ComparisonObjectPath"] = key[1]
    new_row["ComparisonPrincipal"] = key[2]
    new_row["ComparisonPermissionLevels"] = key[3]
    return new_row


def source_user_not_found_in_entra(row, entra_user_aliases):
    if not entra_user_aliases:
        return False
    if normalize_label(row.get("PrincipalType")) != "user":
        return False
    principal = row.get("PrincipalLoginName") or row.get("PrincipalName")
    normalized = normalize_principal_text(principal)
    return bool(normalized) and normalized not in entra_user_aliases


def source_key_is_disabled_entra_user(row, key, disabled_entra_users):
    if not disabled_entra_users:
        return False
    if normalize_label(row.get("PrincipalType")) != "user":
        return False
    comparison_principal = key[2] if len(key) > 2 else ""
    return bool(comparison_principal) and comparison_principal in disabled_entra_users


def with_ignore_reason(row, reason):
    updated = dict(row)
    updated["ComparisonIgnoreReason"] = reason
    return updated

def key_has_comparable_permissions(key):
    return bool(key[3])


def percent(numerator, denominator):
    denominator = int(denominator or 0)
    if denominator <= 0:
        return "0.000000"
    return f"{int(numerator or 0) / denominator:.6f}"


def permission_summary_status(entry):
    differences = []
    if int(entry.get("MissingInSPO") or 0):
        differences.append("missing in SPO")
    if int(entry.get("DisabledEntraUsersNotInSPO") or 0):
        differences.append("disabled Entra users not in SPO")
    if int(entry.get("ExtraInSPO") or 0):
        differences.append("extra in SPO")
    if not differences:
        return "OK - no difference"
    if len(differences) == 1:
        return differences[0][:1].upper() + differences[0][1:]
    return "Mixed differences: " + ", ".join(differences)


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
            "DisabledEntraUsersNotInSPO": 0,
            "DisabledEntraUsersNotInSPOPercent": "0.000000",
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


def build_permission_summary(source_by_key, target_by_key, matched_keys, disabled_entra_missing_keys, missing_keys, extra_keys):
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
    for key in disabled_entra_missing_keys:
        object_scope, object_path = key[0], key[1]
        ensure_permission_summary(stats, object_scope, object_path)["DisabledEntraUsersNotInSPO"] += 1
    for key in extra_keys:
        object_scope, object_path = key[0], key[1]
        ensure_permission_summary(stats, object_scope, object_path)["ExtraInSPO"] += 1

    rows = []
    for entry in stats.values():
        entry["MatchedPermissionsPercent"] = percent(entry["MatchedPermissions"], entry["SourcePermissions"])
        entry["MissingInSPOPercent"] = percent(entry["MissingInSPO"], entry["SourcePermissions"])
        entry["DisabledEntraUsersNotInSPOPercent"] = percent(entry["DisabledEntraUsersNotInSPO"], entry["SourcePermissions"])
        entry["ExtraInSPOPercent"] = percent(entry["ExtraInSPO"], entry["TargetPermissions"])
        entry["Status"] = permission_summary_status(entry)
        rows.append(entry)
    return sorted(
        rows,
        key=lambda item: (
            -int(item.get("MissingInSPO") or 0) - int(item.get("DisabledEntraUsersNotInSPO") or 0) - int(item.get("ExtraInSPO") or 0),
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



def parse_bool(value):
    text = str(value or "").strip().lower()
    if text in {"true", "1", "yes", "y"}:
        return True
    if text in {"false", "0", "no", "n"}:
        return False
    return None


def has_library_scope_metadata(row):
    return parse_bool(row.get("IsDocumentLibrary")) is not None


def is_document_library_permission(row):
    object_scope = (row.get("ObjectScope") or "").strip().lower()
    if object_scope not in {"list", "item"}:
        return False
    return parse_bool(row.get("IsDocumentLibrary")) is True


def library_scope_metadata_missing_count(rows):
    return sum(
        1
        for row in rows
        if (row.get("ObjectScope") or "").strip().lower() in {"list", "item"}
        and not has_library_scope_metadata(row)
    )


def build_scope_warning(report_scope, source_rows, target_rows, source_scan_document_libraries_only=False, target_scan_document_libraries_only=False):
    warnings = []
    source_missing = library_scope_metadata_missing_count(source_rows)
    target_missing = library_scope_metadata_missing_count(target_rows)
    if source_missing or target_missing:
        warnings.append(
            "Permission inventory rows are missing IsDocumentLibrary metadata. Rerun source and target permission scans with the updated inventory scripts for reliable split reports."
        )
    if report_scope == "DocumentLibraries" and not source_rows and not target_rows:
        warnings.append(
            "No document library permission rows were identified. If these scans were created before IsDocumentLibrary metadata existed, rerun source and target permission scans."
        )
    if report_scope == "OtherScopes" and (source_scan_document_libraries_only or target_scan_document_libraries_only):
        limited_sides = []
        if source_scan_document_libraries_only:
            limited_sides.append("source")
        if target_scan_document_libraries_only:
            limited_sides.append("target")
        warnings.append(
            "OtherScopes is incomplete because the {0} permission scan was run with DocumentLibrariesOnly.".format(
                " and ".join(limited_sides)
            )
        )
    return " ".join(warnings)


def write_permission_comparison_report(
    source_rows,
    target_rows,
    output_dir,
    comparison_name,
    source_root_path,
    target_root_path,
    path_mappings=None,
    report_scope="AllScopes",
    source_scan_document_libraries_only=False,
    target_scan_document_libraries_only=False,
    entra_user_aliases=None,
    disabled_entra_users=None,
    sharepoint_group_mappings=None,
    sharepoint_group_mapping_rows=None,
):
    output_dir.mkdir(parents=True, exist_ok=True)
    sharepoint_group_mappings = sharepoint_group_mappings or {}
    sharepoint_group_mapping_rows = sharepoint_group_mapping_rows or []

    source_by_key = {}
    target_by_key = {}
    source_duplicates = []
    target_duplicates = []
    source_limited_access_only = []
    target_limited_access_only = []
    source_users_not_in_entra = []
    disabled_entra_users = disabled_entra_users or set()

    for row in source_rows:
        key = make_key(row, source_root_path, target_root_path, path_mappings=path_mappings, entra_user_aliases=entra_user_aliases, sharepoint_group_mappings=sharepoint_group_mappings)
        keyed = attach_key(row, key)
        if not key_has_comparable_permissions(key):
            source_limited_access_only.append(keyed)
            continue
        if source_user_not_found_in_entra(row, entra_user_aliases):
            source_users_not_in_entra.append(with_ignore_reason(keyed, "Source user principal not found in Entra cache"))
            continue
        if key in source_by_key:
            source_duplicates.append(keyed)
            continue
        source_by_key[key] = keyed

    for row in target_rows:
        key = make_key(row, entra_user_aliases=entra_user_aliases)
        keyed = attach_key(row, key)
        if not key_has_comparable_permissions(key):
            target_limited_access_only.append(keyed)
            continue
        if key in target_by_key:
            target_duplicates.append(keyed)
            continue
        target_by_key[key] = keyed

    source_keys = set(source_by_key)
    target_keys = set(target_by_key)
    matched_keys = sorted(source_keys & target_keys)
    raw_missing_keys = sorted(source_keys - target_keys)
    disabled_entra_missing_keys = [
        key for key in raw_missing_keys
        if source_key_is_disabled_entra_user(source_by_key[key], key, disabled_entra_users)
    ]
    disabled_entra_missing_key_set = set(disabled_entra_missing_keys)
    missing_keys = [key for key in raw_missing_keys if key not in disabled_entra_missing_key_set]
    extra_keys = sorted(target_keys - source_keys)

    summary_rows = [
        {
            "ReportScope": report_scope,
            "SourceRows": len(source_rows),
            "TargetRows": len(target_rows),
            "SourceUniqueKeys": len(source_by_key),
            "TargetUniqueKeys": len(target_by_key),
            "MatchedPermissions": len(matched_keys),
            "MissingInSPO": len(missing_keys),
            "DisabledEntraUsersNotInSPO": len(disabled_entra_missing_keys),
            "ExtraInSPO": len(extra_keys),
            "SourceDuplicateKeysIgnored": len(source_duplicates),
            "TargetDuplicateKeysIgnored": len(target_duplicates),
            "SourceLimitedAccessOnlyIgnored": len(source_limited_access_only),
            "TargetLimitedAccessOnlyIgnored": len(target_limited_access_only),
            "EntraUserAliasesLoaded": len(entra_user_aliases or {}),
            "DisabledEntraUsersLoaded": len(disabled_entra_users),
            "SourceUsersNotInEntraIgnored": len(source_users_not_in_entra),
            "SharePointGroupMappingsLoaded": len(sharepoint_group_mappings),
            "SourceRowsWithoutLibraryScopeMetadata": library_scope_metadata_missing_count(source_rows),
            "TargetRowsWithoutLibraryScopeMetadata": library_scope_metadata_missing_count(target_rows),
            "SourceScanDocumentLibrariesOnly": str(bool(source_scan_document_libraries_only)),
            "TargetScanDocumentLibrariesOnly": str(bool(target_scan_document_libraries_only)),
            "SourceRootPath": source_root_path,
            "TargetRootPath": target_root_path,
            "ScopeWarning": build_scope_warning(
                report_scope,
                source_rows,
                target_rows,
                source_scan_document_libraries_only=source_scan_document_libraries_only,
                target_scan_document_libraries_only=target_scan_document_libraries_only,
            ),
        }
    ]

    def rows_from_keys(keys, lookup):
        return [lookup[key] for key in keys]

    matched_rows = rows_from_keys(matched_keys, source_by_key)
    missing_rows = rows_from_keys(missing_keys, source_by_key)
    disabled_entra_missing_rows = [
        with_ignore_reason(source_by_key[key], "Source user principal is disabled in Entra and missing in SPO")
        for key in disabled_entra_missing_keys
    ]
    extra_rows = rows_from_keys(extra_keys, target_by_key)

    scope_counter = Counter()
    for key in matched_keys:
        scope_counter[(key[0], "Matched")] += 1
    for key in missing_keys:
        scope_counter[(key[0], "MissingInSPO")] += 1
    for key in disabled_entra_missing_keys:
        scope_counter[(key[0], "DisabledEntraUsersNotInSPO")] += 1
    for key in extra_keys:
        scope_counter[(key[0], "ExtraInSPO")] += 1
    scope_rows = [
        {"ObjectScope": scope, "Status": status, "Count": count}
        for (scope, status), count in sorted(scope_counter.items())
    ]
    permission_summary_rows = build_permission_summary(source_by_key, target_by_key, matched_keys, disabled_entra_missing_keys, missing_keys, extra_keys)

    summary_csv = output_dir / "Summary.csv"
    scope_csv = output_dir / "ScopeSummary.csv"
    permission_summary_csv = output_dir / "PermissionSummary.csv"
    missing_csv = output_dir / "MissingInSPO.csv"
    disabled_entra_missing_csv = output_dir / "DisabledEntraUsersNotInSPO.csv"
    extra_csv = output_dir / "ExtraInSPO.csv"
    matched_csv = output_dir / "Matched.csv"
    duplicate_source_csv = output_dir / "DuplicateKeys-Source.csv"
    duplicate_target_csv = output_dir / "DuplicateKeys-Target.csv"
    limited_access_source_csv = output_dir / "LimitedAccessOnly-Source.csv"
    limited_access_target_csv = output_dir / "LimitedAccessOnly-Target.csv"
    source_users_not_in_entra_csv = output_dir / "SourceUsersNotInEntra.csv"
    sharepoint_group_mappings_csv = output_dir / "SharePointGroupMappings.csv"

    write_csv(summary_csv, summary_rows, list(summary_rows[0].keys()))
    write_csv(scope_csv, scope_rows, ["ObjectScope", "Status", "Count"])
    write_csv(permission_summary_csv, permission_summary_rows, PERMISSION_SUMMARY_COLUMNS)

    output_fields = list(source_rows[0].keys()) if source_rows else (list(target_rows[0].keys()) if target_rows else [])
    for extra_field in ["ComparisonObjectScope", "ComparisonObjectPath", "ComparisonPrincipal", "ComparisonPermissionLevels", "ComparisonIgnoreReason"]:
        if extra_field not in output_fields:
            output_fields.append(extra_field)
    write_csv(missing_csv, missing_rows, output_fields)
    write_csv(disabled_entra_missing_csv, disabled_entra_missing_rows, output_fields)
    write_csv(matched_csv, matched_rows, output_fields)

    target_fields = list(target_rows[0].keys()) if target_rows else output_fields
    for extra_field in ["ComparisonObjectScope", "ComparisonObjectPath", "ComparisonPrincipal", "ComparisonPermissionLevels", "ComparisonIgnoreReason"]:
        if extra_field not in target_fields:
            target_fields.append(extra_field)
    write_csv(extra_csv, extra_rows, target_fields)
    write_csv(duplicate_source_csv, source_duplicates, output_fields)
    write_csv(duplicate_target_csv, target_duplicates, target_fields)
    write_csv(limited_access_source_csv, source_limited_access_only, output_fields)
    write_csv(limited_access_target_csv, target_limited_access_only, target_fields)
    write_csv(source_users_not_in_entra_csv, source_users_not_in_entra, output_fields)
    write_csv(sharepoint_group_mappings_csv, sharepoint_group_mapping_rows, ["MappingSourceWebPath", "MappingTargetWebPath", "MappingSourceObjectPath", "MappingRole", "MappingSourcePrincipalName", "MappingTargetPrincipalName", "MappingTargetComparisonPrincipal", "MappingMethod"])

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    xlsx_path = output_dir / f"{comparison_name}-{timestamp}.xlsx"
    create_xlsx(
        xlsx_path,
        [
            ("Summary", list_rows_for_excel(summary_csv)),
            ("PermissionSummary", list_rows_for_excel(permission_summary_csv)),
            ("ScopeSummary", list_rows_for_excel(scope_csv)),
            ("MissingInSPO", list_rows_for_excel(missing_csv)),
            ("DisabledUsers", list_rows_for_excel(disabled_entra_missing_csv)),
            ("ExtraInSPO", list_rows_for_excel(extra_csv)),
            ("Matched", list_rows_for_excel(matched_csv)),
            ("DuplicateSource", list_rows_for_excel(duplicate_source_csv)),
            ("DuplicateTarget", list_rows_for_excel(duplicate_target_csv)),
            ("IgnoredSourceLA", list_rows_for_excel(limited_access_source_csv)),
            ("IgnoredTargetLA", list_rows_for_excel(limited_access_target_csv)),
            ("IgnoredSourceNotEntra", list_rows_for_excel(source_users_not_in_entra_csv)),
            ("SPGroupMappings", list_rows_for_excel(sharepoint_group_mappings_csv)),
        ],
    )

    html_path = output_dir / f"{comparison_name}-summary-{timestamp}.html"
    create_permission_html_summary(
        html_path,
        f"{comparison_name} - {report_scope}",
        summary_rows[0],
        permission_summary_rows,
        scope_rows,
        [
            ("Summary", summary_csv, "Run counters and comparison inputs."),
            ("Permission summary", permission_summary_csv, "Object-level summary of permission differences."),
            ("Scope summary", scope_csv, "Counts grouped by object scope and status."),
            ("Missing in SPO", missing_csv, "Source permissions absent from SPO, excluding disabled Entra users."),
            ("Disabled Entra users", disabled_entra_missing_csv, "Source disabled Entra user permissions absent from SPO."),
            ("Extra in SPO", extra_csv, "Target permissions not found in source."),
            ("Source users not in Entra", source_users_not_in_entra_csv, "Source users ignored because they are not in the Entra cache."),
            ("SharePoint group mappings", sharepoint_group_mappings_csv, "Detected and confirmed SharePoint group mappings."),
            ("Excel workbook", xlsx_path, "Full comparison workbook."),
        ],
    )

    return {
        "ReportScope": report_scope,
        "MatchedPermissions": len(matched_keys),
        "MissingInSPO": len(missing_keys),
        "DisabledEntraUsersNotInSPO": len(disabled_entra_missing_keys),
        "ExtraInSPO": len(extra_keys),
        "Summary": summary_csv,
        "Excel": xlsx_path,
        "Html": html_path,
    }
def main():
    parser = argparse.ArgumentParser(description="Compare SharePoint permission inventories.")
    parser.add_argument("--source-csv", required=True)
    parser.add_argument("--target-csv", required=True)
    parser.add_argument("--output-directory", required=True)
    parser.add_argument("--source-root-path", default="/FR")
    parser.add_argument("--target-root-path", default="/")
    parser.add_argument("--path-mapping-file")
    parser.add_argument("--sharegate-replacement-character", default="_")
    parser.add_argument("--comparison-name", default="SP2019-vs-SPO-Permissions")
    parser.add_argument("--source-scan-document-libraries-only", action="store_true")
    parser.add_argument("--target-scan-document-libraries-only", action="store_true")
    parser.add_argument("--entra-users-csv", required=True)
    args = parser.parse_args()

    if len(args.sharegate_replacement_character) != 1:
        raise ValueError("--sharegate-replacement-character must contain exactly one character.")

    global SHAREGATE_REPLACEMENT_CHARACTER
    SHAREGATE_REPLACEMENT_CHARACTER = args.sharegate_replacement_character

    source_csv = Path(args.source_csv)
    target_csv = Path(args.target_csv)
    output_dir = Path(args.output_directory)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_rows = list(read_rows(source_csv))
    target_rows = list(read_rows(target_csv))

    path_mappings = load_path_mappings(args.path_mapping_file)
    if path_mappings:
        print(f"Path mappings loaded: {len(path_mappings)}")

    entra_user_aliases, disabled_entra_users = load_entra_user_aliases(args.entra_users_csv)
    print(f"Entra user aliases loaded: {len(entra_user_aliases)} from {args.entra_users_csv}")
    print(f"Disabled Entra users loaded: {len(disabled_entra_users)} from {args.entra_users_csv}")

    sharepoint_group_mappings, sharepoint_group_mapping_rows = build_sharepoint_group_mappings(
        source_rows,
        target_rows,
        source_prefix=args.source_root_path,
        target_prefix=args.target_root_path,
        path_mappings=path_mappings,
    )
    print(f"SharePoint group mappings loaded: {len(sharepoint_group_mappings)}")

    source_by_key = {}
    target_by_key = {}
    source_duplicates = []
    target_duplicates = []
    source_limited_access_only = []
    target_limited_access_only = []
    source_users_not_in_entra = []

    for row in source_rows:
        key = make_key(row, args.source_root_path, args.target_root_path, path_mappings=path_mappings, entra_user_aliases=entra_user_aliases, sharepoint_group_mappings=sharepoint_group_mappings)
        keyed = attach_key(row, key)
        if not key_has_comparable_permissions(key):
            source_limited_access_only.append(keyed)
            continue
        if source_user_not_found_in_entra(row, entra_user_aliases):
            source_users_not_in_entra.append(with_ignore_reason(keyed, "Source user principal not found in Entra cache"))
            continue
        if key in source_by_key:
            source_duplicates.append(keyed)
            continue
        source_by_key[key] = keyed

    for row in target_rows:
        key = make_key(row, entra_user_aliases=entra_user_aliases)
        keyed = attach_key(row, key)
        if not key_has_comparable_permissions(key):
            target_limited_access_only.append(keyed)
            continue
        if key in target_by_key:
            target_duplicates.append(keyed)
            continue
        target_by_key[key] = keyed

    source_keys = set(source_by_key)
    target_keys = set(target_by_key)
    matched_keys = sorted(source_keys & target_keys)
    raw_missing_keys = sorted(source_keys - target_keys)
    disabled_entra_missing_keys = [
        key for key in raw_missing_keys
        if source_key_is_disabled_entra_user(source_by_key[key], key, disabled_entra_users)
    ]
    disabled_entra_missing_key_set = set(disabled_entra_missing_keys)
    missing_keys = [key for key in raw_missing_keys if key not in disabled_entra_missing_key_set]
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
            "DisabledEntraUsersNotInSPO": len(disabled_entra_missing_keys),
            "ExtraInSPO": len(extra_keys),
            "SourceDuplicateKeysIgnored": len(source_duplicates),
            "TargetDuplicateKeysIgnored": len(target_duplicates),
            "SourceLimitedAccessOnlyIgnored": len(source_limited_access_only),
            "TargetLimitedAccessOnlyIgnored": len(target_limited_access_only),
            "EntraUserAliasesLoaded": len(entra_user_aliases or {}),
            "DisabledEntraUsersLoaded": len(disabled_entra_users),
            "SourceUsersNotInEntraIgnored": len(source_users_not_in_entra),
            "SharePointGroupMappingsLoaded": len(sharepoint_group_mappings),
            "SourceRootPath": args.source_root_path,
            "TargetRootPath": args.target_root_path,
        }
    ]

    def rows_from_keys(keys, lookup):
        return [lookup[key] for key in keys]

    matched_rows = rows_from_keys(matched_keys, source_by_key)
    missing_rows = rows_from_keys(missing_keys, source_by_key)
    disabled_entra_missing_rows = [
        with_ignore_reason(source_by_key[key], "Source user principal is disabled in Entra and missing in SPO")
        for key in disabled_entra_missing_keys
    ]
    extra_rows = rows_from_keys(extra_keys, target_by_key)

    scope_counter = Counter()
    for key in matched_keys:
        scope_counter[(key[0], "Matched")] += 1
    for key in missing_keys:
        scope_counter[(key[0], "MissingInSPO")] += 1
    for key in disabled_entra_missing_keys:
        scope_counter[(key[0], "DisabledEntraUsersNotInSPO")] += 1
    for key in extra_keys:
        scope_counter[(key[0], "ExtraInSPO")] += 1
    scope_rows = [
        {"ObjectScope": scope, "Status": status, "Count": count}
        for (scope, status), count in sorted(scope_counter.items())
    ]
    permission_summary_rows = build_permission_summary(source_by_key, target_by_key, matched_keys, disabled_entra_missing_keys, missing_keys, extra_keys)

    summary_csv = output_dir / "Summary.csv"
    scope_csv = output_dir / "ScopeSummary.csv"
    permission_summary_csv = output_dir / "PermissionSummary.csv"
    missing_csv = output_dir / "MissingInSPO.csv"
    disabled_entra_missing_csv = output_dir / "DisabledEntraUsersNotInSPO.csv"
    extra_csv = output_dir / "ExtraInSPO.csv"
    matched_csv = output_dir / "Matched.csv"
    duplicate_source_csv = output_dir / "DuplicateKeys-Source.csv"
    duplicate_target_csv = output_dir / "DuplicateKeys-Target.csv"
    limited_access_source_csv = output_dir / "LimitedAccessOnly-Source.csv"
    limited_access_target_csv = output_dir / "LimitedAccessOnly-Target.csv"
    source_users_not_in_entra_csv = output_dir / "SourceUsersNotInEntra.csv"
    sharepoint_group_mappings_csv = output_dir / "SharePointGroupMappings.csv"

    write_csv(summary_csv, summary_rows, list(summary_rows[0].keys()))
    write_csv(scope_csv, scope_rows, ["ObjectScope", "Status", "Count"])
    write_csv(permission_summary_csv, permission_summary_rows, PERMISSION_SUMMARY_COLUMNS)

    output_fields = list(source_rows[0].keys()) if source_rows else []
    for extra_field in ["ComparisonObjectScope", "ComparisonObjectPath", "ComparisonPrincipal", "ComparisonPermissionLevels", "ComparisonIgnoreReason"]:
        if extra_field not in output_fields:
            output_fields.append(extra_field)
    write_csv(missing_csv, missing_rows, output_fields)
    write_csv(disabled_entra_missing_csv, disabled_entra_missing_rows, output_fields)
    write_csv(matched_csv, matched_rows, output_fields)

    target_fields = list(target_rows[0].keys()) if target_rows else output_fields
    for extra_field in ["ComparisonObjectScope", "ComparisonObjectPath", "ComparisonPrincipal", "ComparisonPermissionLevels", "ComparisonIgnoreReason"]:
        if extra_field not in target_fields:
            target_fields.append(extra_field)
    write_csv(extra_csv, extra_rows, target_fields)
    write_csv(duplicate_source_csv, source_duplicates, output_fields)
    write_csv(duplicate_target_csv, target_duplicates, target_fields)
    write_csv(limited_access_source_csv, source_limited_access_only, output_fields)
    write_csv(limited_access_target_csv, target_limited_access_only, target_fields)
    write_csv(source_users_not_in_entra_csv, source_users_not_in_entra, output_fields)
    write_csv(sharepoint_group_mappings_csv, sharepoint_group_mapping_rows, ["MappingSourceWebPath", "MappingTargetWebPath", "MappingSourceObjectPath", "MappingRole", "MappingSourcePrincipalName", "MappingTargetPrincipalName", "MappingTargetComparisonPrincipal", "MappingMethod"])

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    xlsx_path = output_dir / f"{args.comparison_name}-{timestamp}.xlsx"
    create_xlsx(
        xlsx_path,
        [
            ("Summary", list_rows_for_excel(summary_csv)),
            ("PermissionSummary", list_rows_for_excel(permission_summary_csv)),
            ("ScopeSummary", list_rows_for_excel(scope_csv)),
            ("MissingInSPO", list_rows_for_excel(missing_csv)),
            ("DisabledUsers", list_rows_for_excel(disabled_entra_missing_csv)),
            ("ExtraInSPO", list_rows_for_excel(extra_csv)),
            ("Matched", list_rows_for_excel(matched_csv)),
            ("DuplicateSource", list_rows_for_excel(duplicate_source_csv)),
            ("DuplicateTarget", list_rows_for_excel(duplicate_target_csv)),
            ("IgnoredSourceLA", list_rows_for_excel(limited_access_source_csv)),
            ("IgnoredTargetLA", list_rows_for_excel(limited_access_target_csv)),
            ("IgnoredSourceNotEntra", list_rows_for_excel(source_users_not_in_entra_csv)),
            ("SPGroupMappings", list_rows_for_excel(sharepoint_group_mappings_csv)),
        ],
    )

    html_path = output_dir / f"{args.comparison_name}-summary-{timestamp}.html"
    create_permission_html_summary(
        html_path,
        args.comparison_name,
        summary_rows[0],
        permission_summary_rows,
        scope_rows,
        [
            ("Summary", summary_csv, "Run counters and comparison inputs."),
            ("Permission summary", permission_summary_csv, "Object-level summary of permission differences."),
            ("Scope summary", scope_csv, "Counts grouped by object scope and status."),
            ("Missing in SPO", missing_csv, "Source permissions absent from SPO, excluding disabled Entra users."),
            ("Disabled Entra users", disabled_entra_missing_csv, "Source disabled Entra user permissions absent from SPO."),
            ("Extra in SPO", extra_csv, "Target permissions not found in source."),
            ("Source users not in Entra", source_users_not_in_entra_csv, "Source users ignored because they are not in the Entra cache."),
            ("SharePoint group mappings", sharepoint_group_mappings_csv, "Detected and confirmed SharePoint group mappings."),
            ("Excel workbook", xlsx_path, "Full comparison workbook."),
        ],
        source_csv=source_csv,
        target_csv=target_csv,
    )

    document_library_report = write_permission_comparison_report(
        [row for row in source_rows if is_document_library_permission(row)],
        [row for row in target_rows if is_document_library_permission(row)],
        output_dir / "DocumentLibraries",
        f"{args.comparison_name}-DocumentLibraries",
        args.source_root_path,
        args.target_root_path,
        path_mappings=path_mappings,
        report_scope="DocumentLibraries",
        source_scan_document_libraries_only=args.source_scan_document_libraries_only,
        target_scan_document_libraries_only=args.target_scan_document_libraries_only,
        entra_user_aliases=entra_user_aliases,
        disabled_entra_users=disabled_entra_users,
        sharepoint_group_mappings=sharepoint_group_mappings,
        sharepoint_group_mapping_rows=sharepoint_group_mapping_rows,
    )
    other_scope_report = write_permission_comparison_report(
        [row for row in source_rows if not is_document_library_permission(row)],
        [row for row in target_rows if not is_document_library_permission(row)],
        output_dir / "OtherScopes",
        f"{args.comparison_name}-OtherScopes",
        args.source_root_path,
        args.target_root_path,
        path_mappings=path_mappings,
        report_scope="OtherScopes",
        source_scan_document_libraries_only=args.source_scan_document_libraries_only,
        target_scan_document_libraries_only=args.target_scan_document_libraries_only,
        entra_user_aliases=entra_user_aliases,
        disabled_entra_users=disabled_entra_users,
        sharepoint_group_mappings=sharepoint_group_mappings,
        sharepoint_group_mapping_rows=sharepoint_group_mapping_rows,
    )
    print(f"Comparison completed.")
    print(f"Matched permissions: {len(matched_keys)}")
    print(f"Missing in SPO: {len(missing_keys)}")
    print(f"Disabled Entra users not in SPO: {len(disabled_entra_missing_keys)}")
    print(f"Extra in SPO: {len(extra_keys)}")
    print(f"Summary: {summary_csv}")
    print(f"Excel: {xlsx_path}")
    print(f"HTML summary: {html_path}")
    print(f"Document library permission report: {document_library_report['Excel']}")
    print(f"Document library HTML summary: {document_library_report['Html']}")
    print(f"Other SharePoint scope permission report: {other_scope_report['Excel']}")
    print(f"Other SharePoint scope HTML summary: {other_scope_report['Html']}")


if __name__ == "__main__":
    main()

