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
from urllib.parse import unquote, urlparse


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
    "SourceRowsWithoutLibraryScopeMetadata",
    "TargetRowsWithoutLibraryScopeMetadata",
    "SourcePermissions",
    "TargetPermissions",
    "Count",
}
PERCENT_COLUMNS = {
    "MatchedPermissionsPercent",
    "MissingInSPOPercent",
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
    "collaboration": "contribute",
    "conception": "design",
    "lecture restreinte": "restricted view",
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
        return {}

    csv_path = Path(path)
    if not csv_path.exists():
        raise FileNotFoundError(f"Entra users cache not found: {csv_path}")

    alias_map = {}
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
        for field in ("ProxyAddresses", "Proxy addresses", "OtherMails", "otherMails"):
            identity_values.extend(split_identity_values(row.get(field)))

        for value in identity_values:
            add_entra_alias(alias_map, value, canonical)

    return alias_map


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
        return SHAREGATE_REPLACEMENT_CHARACTER

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


def normalize_principal_text(value):
    text = (value or "").strip().lower()
    if "|" in text:
        text = text.split("|")[-1]
    text = re.sub(r"^i:0[#.a-z0-9]*", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_sharepoint_group_name(value):
    text = normalize_label(value)
    match = re.search(r"(?:\s+-\s+|\s+)([^\s-]+)$", text)
    if not match:
        return None, None
    suffix = SHAREPOINT_ASSOCIATED_GROUP_SUFFIXES.get(match.group(1))
    if not suffix:
        return None, None
    base = text[: match.start()].strip(" -")
    return base, suffix


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


def normalize_principal(row, web_path="", entra_user_aliases=None):
    principal = row.get("PrincipalLoginName") or row.get("PrincipalName")
    normalized = normalize_principal_text(principal)
    principal_type = normalize_label(row.get("PrincipalType"))
    if principal_type == "sharepointgroup":
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
    return normalize_default_document_library_key(row, path, source_prefix, target_prefix, path_mappings=path_mappings)


def make_key(row, source_prefix=None, target_prefix=None, path_mappings=None, entra_user_aliases=None):
    web_path = row_web_path(row, source_prefix, target_prefix, path_mappings=path_mappings)
    return (
        (row.get("ObjectScope") or "").strip().lower(),
        row_path(row, source_prefix, target_prefix, path_mappings=path_mappings),
        normalize_principal(row, web_path=web_path, entra_user_aliases=entra_user_aliases),
        normalize_permissions(row.get("PermissionLevels")),
    )


def attach_key(row, key):
    new_row = dict(row)
    new_row["ComparisonObjectScope"] = key[0]
    new_row["ComparisonObjectPath"] = key[1]
    new_row["ComparisonPrincipal"] = key[2]
    new_row["ComparisonPermissionLevels"] = key[3]
    return new_row

def key_has_comparable_permissions(key):
    return bool(key[3])


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
):
    output_dir.mkdir(parents=True, exist_ok=True)
    source_by_key = {}
    target_by_key = {}
    source_duplicates = []
    target_duplicates = []
    source_limited_access_only = []
    target_limited_access_only = []

    for row in source_rows:
        key = make_key(row, source_root_path, target_root_path, path_mappings=path_mappings, entra_user_aliases=entra_user_aliases)
        keyed = attach_key(row, key)
        if not key_has_comparable_permissions(key):
            source_limited_access_only.append(keyed)
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
    missing_keys = sorted(source_keys - target_keys)
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
            "ExtraInSPO": len(extra_keys),
            "SourceDuplicateKeysIgnored": len(source_duplicates),
            "TargetDuplicateKeysIgnored": len(target_duplicates),
            "SourceLimitedAccessOnlyIgnored": len(source_limited_access_only),
            "TargetLimitedAccessOnlyIgnored": len(target_limited_access_only),
            "EntraUserAliasesLoaded": len(entra_user_aliases or {}),
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
    limited_access_source_csv = output_dir / "LimitedAccessOnly-Source.csv"
    limited_access_target_csv = output_dir / "LimitedAccessOnly-Target.csv"

    write_csv(summary_csv, summary_rows, list(summary_rows[0].keys()))
    write_csv(scope_csv, scope_rows, ["ObjectScope", "Status", "Count"])
    write_csv(permission_summary_csv, permission_summary_rows, PERMISSION_SUMMARY_COLUMNS)

    output_fields = list(source_rows[0].keys()) if source_rows else (list(target_rows[0].keys()) if target_rows else [])
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
    write_csv(limited_access_source_csv, source_limited_access_only, output_fields)
    write_csv(limited_access_target_csv, target_limited_access_only, target_fields)

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    xlsx_path = output_dir / f"{comparison_name}-{timestamp}.xlsx"
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
            ("IgnoredSourceLA", list_rows_for_excel(limited_access_source_csv)),
            ("IgnoredTargetLA", list_rows_for_excel(limited_access_target_csv)),
        ],
    )

    return {
        "ReportScope": report_scope,
        "MatchedPermissions": len(matched_keys),
        "MissingInSPO": len(missing_keys),
        "ExtraInSPO": len(extra_keys),
        "Summary": summary_csv,
        "Excel": xlsx_path,
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
    parser.add_argument("--entra-users-csv")
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

    entra_user_aliases = load_entra_user_aliases(args.entra_users_csv)
    if args.entra_users_csv:
        print(f"Entra user aliases loaded: {len(entra_user_aliases)} from {args.entra_users_csv}")

    source_by_key = {}
    target_by_key = {}
    source_duplicates = []
    target_duplicates = []
    source_limited_access_only = []
    target_limited_access_only = []

    for row in source_rows:
        key = make_key(row, args.source_root_path, args.target_root_path, path_mappings=path_mappings, entra_user_aliases=entra_user_aliases)
        keyed = attach_key(row, key)
        if not key_has_comparable_permissions(key):
            source_limited_access_only.append(keyed)
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
            "SourceLimitedAccessOnlyIgnored": len(source_limited_access_only),
            "TargetLimitedAccessOnlyIgnored": len(target_limited_access_only),
            "EntraUserAliasesLoaded": len(entra_user_aliases or {}),
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
    limited_access_source_csv = output_dir / "LimitedAccessOnly-Source.csv"
    limited_access_target_csv = output_dir / "LimitedAccessOnly-Target.csv"

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
    write_csv(limited_access_source_csv, source_limited_access_only, output_fields)
    write_csv(limited_access_target_csv, target_limited_access_only, target_fields)

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
            ("IgnoredSourceLA", list_rows_for_excel(limited_access_source_csv)),
            ("IgnoredTargetLA", list_rows_for_excel(limited_access_target_csv)),
        ],
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
    )
    print(f"Comparison completed.")
    print(f"Matched permissions: {len(matched_keys)}")
    print(f"Missing in SPO: {len(missing_keys)}")
    print(f"Extra in SPO: {len(extra_keys)}")
    print(f"Summary: {summary_csv}")
    print(f"Excel: {xlsx_path}")
    print(f"Document library permission report: {document_library_report['Excel']}")
    print(f"Other SharePoint scope permission report: {other_scope_report['Excel']}")


if __name__ == "__main__":
    main()
