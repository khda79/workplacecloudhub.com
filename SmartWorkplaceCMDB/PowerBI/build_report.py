"""Build a private frozen native Power BI report (six or nine pages), stable V1.

Python 3.10+ standard library only. Existing destinations are never overwritten.
Run: python build_report.py --data-root <DATA-LAST> --output <NEW-DIRECTORY>
"""
import argparse
import csv
import datetime as dt
import hashlib
import importlib.util
import json
import re
from pathlib import Path

VERSION = "1.0.0-local-report"
BASE = "https://developer.microsoft.com/json-schemas/"
IDENTITY = ["TenantKey", "OrganizationKey", "EnvironmentKey", "TenantId"]
PRODUCT = Path(__file__).resolve().parents[1]
PRESENTATION_COLUMNS = {
    "DimUser": ["DisplayName", "UserPrincipalName", "Department", "UserType"],
    "DimDevice": ["DeviceName", "OperatingSystem", "OperatingSystemVersion", "Ownership", "ComplianceState", "EncryptionState", "ManagementState"],
    "DimLicenseSku": ["SkuPartNumber"],
    "FactUserLicense": ["AssignmentState"],
    "FactMailbox": ["DisplayName", "PrimarySmtpAddress", "RecipientTypeDetails", "ArchiveStatus"],
    "FactDataQuality": ["FindingType", "Description", "RecommendedAction"],
    "SourceHealth": ["SourceRows", "CompletedDateTime"],
}
RELATIONSHIPS = [
    ("FactUserLicense", "TenantUserKey", "DimUser", "TenantUserKey"),
    ("FactUserLicense", "TenantSkuKey", "DimLicenseSku", "TenantSkuKey"),
    ("FactDeviceCompliance", "TenantDeviceKey", "DimDevice", "TenantDeviceKey"),
    ("FactUserDeviceRelationship", "TenantUserKey", "DimUser", "TenantUserKey"),
    ("FactUserDeviceRelationship", "TenantDeviceKey", "DimDevice", "TenantDeviceKey"),
    ("FactMailbox", "TenantUserKey", "DimUser", "TenantUserKey"),
]


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        if any(None in r or any(v is None for v in r.values()) for r in rows):
            raise ValueError("Malformed CSV row: " + path.name)
        return reader.fieldnames, rows


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def write_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def type_of(column):
    if column in ("AccountEnabled", "MailEnabled", "SecurityEnabled"):
        return "boolean", "type logical"
    if column == "ConfidenceScore":
        return "double", "type number"
    if column in ("Year", "Month", "Day", "ConsumedUnits", "EnabledUnits", "SourceRows", "MaxItems", "ObservedPathCount"):
        return "int64", "Int64.Type"
    if column == "Date":
        return "dateTime", "type date"
    if column.endswith("DateTime"):
        return "dateTime", "type datetime"
    return "string", "type text"


def validate_types(name, columns, rows):
    for c in columns:
        ty, _ = type_of(c)
        for row in rows:
            v = row[c]
            if not v:
                continue
            if ty == "boolean" and v.lower() not in ("true", "false"):
                raise ValueError("Invalid boolean: " + name + "." + c)
            if ty == "int64":
                int(v)
            if ty == "double":
                float(v)
            if ty == "dateTime":
                parsed = dt.datetime.fromisoformat(v.replace("Z", "+00:00"))
                if c != "Date" and parsed.tzinfo is None:
                    raise ValueError("Missing timestamp timezone: " + name + "." + c)


def m_string(value):
    return '"' + str(value).replace('"', '""') + '"'


def m_query(path, columns, identity, tenant_scoped=True):
    header = "{" + ",".join(map(m_string, columns)) + "}"
    guard = " and ".join("Record.Field(_, " + m_string(k) + ") = " + m_string(v)
                         for k, v in identity.items()) if tenant_scoped else "true"
    transforms = []
    for c in columns:
        ty, mt = type_of(c)
        convert = {"boolean": "Logical.FromText(_)", "double": 'Number.FromText(_, "en-US")',
                   "int64": 'Int64.From(_, "en-US")', "string": "Text.From(_)"}.get(ty)
        if c == "Date":
            convert = 'Date.FromText(_, [Format="yyyy-MM-dd", Culture="en-US"])'
        elif ty == "dateTime":
            convert = 'DateTimeZone.RemoveZone(DateTimeZone.ToUtc(DateTimeZone.FromText(_, [Culture="en-US"])))'
        transforms.append("{" + m_string(c) + ', each if _ = null or _ = "" then null else ' + convert + ", " + mt + "}")
    return ["let", ' Source = Csv.Document(File.Contents(' + m_string(path) + '), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),',
            " Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),",
            " Checked = if Table.ColumnNames(Headers) = " + header + ' then Headers else error "CSV schema mismatch",',
            " Isolated = if Table.RowCount(Table.SelectRows(Checked, each not (" + guard + '))) = 0 then Checked else error "Unexpected tenant identity",',
            " Typed = Table.TransformColumns(Isolated, {" + ",".join(transforms) + "})", "in Typed"]


def prepare_data(root):
    contract = json.loads((PRODUCT / "Schema/SmartWorkplaceCMDB.tables.json").read_text(encoding="utf-8-sig"))
    needed = [t for t in contract["tables"] if t["area"] == "PowerBI" or t["name"] in ("CMDB_Mailboxes.csv", "CMDB_DataQuality.csv")]
    _, tenants = read_csv(root / "PowerBI/DimTenant.csv")
    if len(tenants) != 1 or any(not tenants[0].get(k) for k in IDENTITY):
        raise ValueError("Exactly one complete tenant identity is required")
    identity = {k: tenants[0][k] for k in IDENTITY}
    data, columns, source_hashes = {}, {}, {}
    for t in needed:
        name = Path(t["name"]).stem
        path = root / t["area"] / t["name"]
        cols, rows = read_csv(path)
        if cols != t["columns"]:
            raise ValueError("CSV contract mismatch: " + name)
        if t.get("tenantScoped", True) and any(any(r[k] != v for k, v in identity.items()) for r in rows):
            raise ValueError("Tenant identity mismatch: " + name)
        validate_types(name, cols, rows)
        data[name], columns[name] = rows, cols[:]
        source_hashes[str(path.relative_to(root))] = sha(path)
    keys = {"DimUser": "TenantUserKey", "DimDevice": "TenantDeviceKey", "DimLicenseSku": "TenantSkuKey",
            "DimGroup": "TenantGroupKey", "DimDate": "Date", "DimTenant": "TenantKey",
            "FactMailbox": "TenantMailboxKey", "FactDataQuality": "TenantFindingKey",
            "FactDeviceCompliance": "TenantDeviceKey", "FactUserDeviceRelationship": "TenantRelationshipKey",
            "CMDB_Mailboxes": "CmdbMailboxId", "CMDB_DataQuality": "FindingId"}
    for name, key in keys.items():
        vals = [r[key].casefold() for r in data[name]]
        if not all(vals) or len(vals) != len(set(vals)):
            raise ValueError("Blank or duplicate key: " + name)
    pairs = [(r["TenantUserKey"].casefold(), r["TenantSkuKey"].casefold()) for r in data["FactUserLicense"]]
    if len(pairs) != len(set(pairs)) or any(not all(p) for p in pairs):
        raise ValueError("Blank or duplicate license pair")
    for f, fk, d, dk in RELATIONSHIPS:
        parent = {r[dk].casefold() for r in data[d]}
        if any(r[fk] and r[fk].casefold() not in parent for r in data[f]):
            raise ValueError("Orphan model relationship: " + f + "." + fk)
    for fact, child, fk, ck, extras in [
        ("FactMailbox", "CMDB_Mailboxes", "CmdbMailboxId", "CmdbMailboxId", ["DisplayName", "PrimarySmtpAddress"]),
        ("FactDataQuality", "CMDB_DataQuality", "FindingId", "FindingId", ["Description", "RecommendedAction"]),
    ]:
        lookup = {r[ck].casefold(): r for r in data[child]}
        if any(r[fk].casefold() not in lookup for r in data[fact]):
            raise ValueError("Missing detail row: " + fact)
        columns[fact].extend(extras)
        for r in data[fact]:
            r.update({c: lookup[r[fk].casefold()][c] for c in extras})
        del data[child], columns[child]
    decorations = {
        "DimUser": {"AccountStatusLabel": lambda r: {"true": "Enabled", "false": "Disabled"}.get(r["AccountEnabled"].lower(), "Not provided")},
        "FactMailbox": {"LinkStatusLabel": lambda r: "Linked to a user" if r["TenantUserKey"] else "No linked user"},
        "FactDataQuality": {"SeverityLabel": lambda r: {"Warning": "Warning", "Information": "Information", "Critical": "Critical"}.get(r["Severity"], r["Severity"] or "Not provided")},
    }
    for name, additions in decorations.items():
        columns[name].extend(additions)
        for r in data[name]:
            r.update({c: fn(r) for c, fn in additions.items()})
    raw_contract = json.loads((PRODUCT / "Schema/SmartWorkplaceCMDB.raw.tables.json").read_text(encoding="utf-8-sig"))
    health = []
    for t in raw_contract["tables"]:
        path = root.joinpath(*t["area"].replace("\\", "/").split("/"), t["name"])
        sidecar = path.with_name(path.name + ".status.json")
        h = dict(identity, SourceName=path.stem, Status="Not collected", Coverage="Unknown", SourceRows="", MaxItems="", StartedDateTime="", CompletedDateTime="", Evidence="Missing")
        if path.exists():
            raw_columns, raw = read_csv(path)
            if raw_columns != t["columns"]:
                raise ValueError("Raw CSV contract mismatch: " + path.name)
            if any(any(r[k] != v for k, v in identity.items()) for r in raw):
                raise ValueError("Raw CSV tenant mismatch: " + path.name)
            h.update(SourceRows=str(len(raw)), Status="CSV without evidence")
            source_hashes[str(path.relative_to(root))] = sha(path)
        if sidecar.exists():
            s = json.loads(sidecar.read_text(encoding="utf-8-sig"))
            if any(s.get(k) != v for k, v in identity.items()):
                raise ValueError("Source evidence tenant mismatch: " + path.name)
            valid = path.exists() and s.get("SHA256") == sha(path) and s.get("RowCount") == len(raw)
            h.update(Status=s.get("Status", "Unknown"), Coverage=s.get("Coverage", "Unknown"),
                     MaxItems=str(s.get("MaxItems", "")), StartedDateTime=s.get("StartedUtc", ""),
                     CompletedDateTime=s.get("CompletedUtc", ""), Evidence="Verified" if valid else "Inconsistent")
            source_hashes[str(sidecar.relative_to(root))] = sha(sidecar)
        health.append(h)
    data["SourceHealth"] = health
    columns["SourceHealth"] = list(health[0])
    validate_types("SourceHealth", columns["SourceHealth"], health)
    # Display-only columns keep source nulls, numeric types and keys unchanged.
    for name, fields in PRESENTATION_COLUMNS.items():
        for field in fields:
            columns[name].append(field + "Label")
            for row in data[name]:
                value = row[field]
                if not value or not value.strip():
                    label = "Not provided"
                elif name == "SourceHealth" and field == "SourceRows":
                    label = format(int(value), ",d")
                elif name == "SourceHealth" and field == "CompletedDateTime":
                    label = dt.datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
                else:
                    label = value
                row[field + "Label"] = label
    return identity, data, columns, source_hashes


def enrich_360(root, identity, data, columns, hashes):
    spec = importlib.util.spec_from_file_location('cmdb_report_360', Path(__file__).with_name('report_360.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    relationships = module.enrich(root, identity, data, columns, hashes, read_csv, sha, PRODUCT)
    for name in data:
        validate_types(name, columns[name], data[name])
    return relationships


def measures(data):
    result = []
    def add(table, name, expression, expected, definition, fmt="#,0"):
        result.append(dict(table=table, name=name, expression=expression, expected=expected, description=definition, formatString=fmt))
    def count(table, name, description):
        add(table, name, "COALESCE(COUNTROWS('" + table + "'),0)", len(data[table]), description)
    count("DimUser", "Users", "Distinct accounts in the Entra export, subject to page filters.")
    count("DimDevice", "Devices", "Unique devices after Entra and Intune reconciliation; a device may not be managed by Intune.")
    count("DimGroup", "Groups", "Exported Entra groups; no membership metric.")
    count("DimLicenseSku", "License SKUs", "Subscribed product SKUs. A SKU does not represent a person.")
    count("FactUserLicense", "License assignments", "Unique user/SKU pairs across all assignment states.")
    count("FactMailbox", "Mailboxes", "All exported mailboxes, including technical and shared mailboxes.")
    count("FactUserDeviceRelationship", "User-device links", "Resolved relationships between a user and a device.")
    count("FactDataQuality", "Quality findings", "Retained findings, including technical information.")
    def filtered(table, base, name, column, value, description):
        is_bool = isinstance(value, bool)
        literal = ("TRUE()" if value else "FALSE()") if is_bool else '"' + value + '"'
        expected = sum((r[column].lower() == str(value).lower()) for r in data[table])
        add(table, name, f"CALCULATE([{base}], KEEPFILTERS('{table}'[{column}] == {literal}))", expected, description)
    filtered("DimUser", "Users", "Enabled accounts", "AccountEnabled", True, "AccountEnabled=true; this is not a measure of activity or sign-ins.")
    filtered("DimUser", "Users", "Disabled accounts", "AccountEnabled", False, "AccountEnabled=false; unknown values remain separate.")
    filtered("DimDevice", "Devices", "Compliant devices", "ComplianceState", "compliant", "State explicitly reported as compliant in the export.")
    filtered("DimDevice", "Devices", "Noncompliant devices", "ComplianceState", "noncompliant", "State explicitly reported as noncompliant; missing states are not classified as noncompliant.")
    filtered("FactDataQuality", "Quality findings", "Warnings", "Severity", "Warning", "Findings classified as Warning by the quality normalizer.")
    filtered("FactDataQuality", "Quality findings", "Information", "Severity", "Information", "Retained information, including technical DiscoveryMailbox findings classified by the local rule.")
    filtered("FactDataQuality", "Quality findings", "Critical findings", "Severity", "Critical", "Findings classified as Critical by the quality normalizer.")
    for table, key, name, definition in [
        ("FactUserLicense", "TenantUserKey", "Users with assignments", "Distinct users with at least one assignment across all states; does not establish an active or used license."),
        ("FactUserDeviceRelationship", "TenantDeviceKey", "Devices with a resolved user", "Distinct devices with a resolved user relationship."),
        ("FactUserDeviceRelationship", "TenantUserKey", "Users with a resolved device", "Distinct users with at least one resolved device relationship."),
    ]:
        add(table, name, f"COALESCE(DISTINCTCOUNTNOBLANK('{table}'[{key}]),0)", len({r[key] for r in data[table] if r[key]}), definition)
    add("FactMailbox", "Mailboxes without a resolved user", 'CALCULATE([Mailboxes], KEEPFILTERS(ISBLANK(\'FactMailbox\'[TenantUserKey])))', sum(not r["TenantUserKey"] for r in data["FactMailbox"]), "Includes technical mailboxes; a missing user link is not always an issue.")
    conform = sum(r["ComplianceState"].lower() == "compliant" for r in data["DimDevice"])
    add("DimDevice", "Compliant device share", "DIVIDE([Compliant devices], [Devices])", conform / len(data["DimDevice"]) if data["DimDevice"] else None, "Compliant devices / all filtered devices, including unknown states in the denominator.", "0.0%")
    enabled = sum(r["AccountEnabled"].lower() == "true" for r in data["DimUser"])
    add("DimUser", "Enabled account share", "DIVIDE([Enabled accounts], [Users])", enabled / len(data["DimUser"]) if data["DimUser"] else None, "Enabled accounts / all filtered users; no activity information is implied.", "0.0%")
    managed = sum(r["ManagementState"].lower() == "managed" for r in data["DimDevice"])
    unmanaged = sum(r["ManagementState"].lower() == "unmanaged" for r in data["DimDevice"])
    missing_management = sum(not r["ManagementState"] for r in data["DimDevice"])
    other_compliance = max(0, len(data["DimDevice"]) - conform - sum(r["ComplianceState"].lower() == "noncompliant" for r in data["DimDevice"]))
    cockpit = [
        ("DimDevice", "Managed devices", "CALCULATE([Devices], KEEPFILTERS('DimDevice'[ManagementState] == \"Managed\"))", managed, "Devices whose exported management state is explicitly Managed. This does not prove current Intune enrollment or activity.", "#,0"),
        ("DimDevice", "Unmanaged devices", "CALCULATE([Devices], KEEPFILTERS('DimDevice'[ManagementState] == \"Unmanaged\"))", unmanaged, "Devices whose exported management state is explicitly Unmanaged.", "#,0"),
        ("DimDevice", "Devices with missing management state", "CALCULATE([Devices], KEEPFILTERS(ISBLANK('DimDevice'[ManagementState])))", missing_management, "Devices without a management state in the frozen snapshot; they are not classified as unmanaged.", "#,0"),
        ("DimDevice", "Managed device share", "DIVIDE([Managed devices], [Devices])", managed / len(data["DimDevice"]) if data["DimDevice"] else None, "Explicitly managed devices divided by all filtered devices. Missing management states remain in the denominator.", "0.0%"),
        ("DimDevice", "Other or missing compliance devices", "MAX(0, [Devices] - [Compliant devices] - [Noncompliant devices])", other_compliance, "Devices not explicitly classified as compliant or noncompliant, including missing, unknown and grace-period states.", "#,0"),
        ("FactUserDeviceRelationship", "Device user-link coverage", "DIVIDE([Devices with a resolved user], [Devices])", None, "Devices with at least one resolved user relationship divided by all filtered devices. A relationship does not prove exclusive ownership.", "0.0%"),
        ("FactUserDeviceRelationship", "User device-link coverage", "DIVIDE([Users with a resolved device], [Users])", None, "Users with at least one resolved device relationship divided by all filtered users.", "0.0%"),
        ("FactUserLicense", "User license assignment coverage", "DIVIDE([Users with assignments], [Users])", None, "Users with at least one observed license assignment divided by all filtered users. This is not license usage.", "0.0%"),
        ("SourceHealth", "Source feeds", "COALESCE(COUNTROWS('SourceHealth'), 0)", len(data.get("SourceHealth", [])), "Configured source feeds represented in SourceHealth for the frozen snapshot.", "#,0"),
        ("SourceHealth", "Collected source feeds", "CALCULATE([Source feeds], KEEPFILTERS('SourceHealth'[Status] == \"Completed\"))", sum(r["Status"].lower() == "completed" for r in data.get("SourceHealth", [])), "Source feeds whose collection status is explicitly Completed. Completed is not tenant certification.", "#,0"),
    ]
    corporate = sum(r["Ownership"].lower() == "corporate" for r in data["DimDevice"])
    cockpit += [
        ("DimDevice", "Corporate devices", "CALCULATE([Devices], REMOVEFILTERS('DimDevice'[Ownership]), 'DimDevice'[Ownership] == \"Corporate\")", corporate, "Corporate devices in the current country and device context. The ownership slicer is intentionally ignored so this remains a corporate-only metric.", "#,0"),
        ("DimDevice", "Corporate device country share", "VAR _selectedCountryCorporateDevices = [Corporate devices] VAR _allCountryCorporateDevices = CALCULATE([Corporate devices], REMOVEFILTERS('DimDevice'[CountryCode], 'DimDevice'[CountryLabel], 'DimDevice'[CountryStatus])) RETURN DIVIDE(_selectedCountryCorporateDevices, _allCountryCorporateDevices)", 1 if corporate else None, "Corporate devices in the selected country divided by all corporate devices. Unknown or unassigned devices remain in the denominator.", "0.0%"),
    ]
    source_total = len(data.get("SourceHealth", []))
    source_done = sum(r["Status"].lower() == "completed" for r in data.get("SourceHealth", []))
    cockpit += [
        ("SourceHealth", "Missing source feeds", "MAX(0, [Source feeds] - [Collected source feeds])", max(0, source_total - source_done), "Source feeds not explicitly completed in the frozen snapshot.", "#,0"),
        ("SourceHealth", "Source collection coverage", "DIVIDE([Collected source feeds], [Source feeds])", source_done / source_total if source_total else None, "Completed source feeds divided by represented source feeds. This is collection coverage, not a composite CMDB quality score.", "0.0%"),
    ]
    for c, name in [("EnabledUnits", "SKU enabled units"), ("ConsumedUnits", "SKU consumed units")]:
        expected = int(data["DimLicenseSku"][0][c]) if len(data["DimLicenseSku"]) == 1 and data["DimLicenseSku"][0][c] else None
        add("DimLicenseSku", name, f"IF(HASONEVALUE('DimLicenseSku'[TenantSkuKey]), SUM('DimLicenseSku'[{c}]))", expected, "Available only for a single SKU; do not sum different products as a count of people.")
    if 'LicenseAssignmentPath' in data:
        count('LicenseAssignmentPath', 'Assignment paths', 'Individual direct/group assignment paths, not distinct user/SKU pairs.')
        filtered('LicenseAssignmentPath', 'Assignment paths', 'Group assignment paths', 'AssignmentRoute', 'Group', 'Observed group-based paths, not membership or effective access.')
        filtered('LicenseAssignmentPath', 'Assignment paths', 'Direct assignment paths', 'AssignmentRoute', 'Direct', 'Observed direct assignment paths.')
        filtered('LicenseAssignmentPath', 'Assignment paths', 'Assignment paths with errors', 'ErrorStatus', 'Error reported', 'Paths reporting an error; another path may still supply the entitlement.')
        add('LicenseAssignmentPath', 'Assignment error rate', 'DIVIDE([Assignment paths with errors], [Assignment paths])', None,
            'Assignment paths reporting an error divided by all observed assignment paths. Another path may still supply the entitlement.', '0.000%')
        add('FactMailbox', 'Mailbox link gap rate', 'DIVIDE([Mailboxes without a resolved user], [Mailboxes])', None,
            'Mailboxes without a resolved user divided by all mailboxes. Technical and shared mailboxes can be legitimate.', '0.000%')
        for table, name, expression, expected, definition, fmt in cockpit:
            add(table, name, expression, expected, definition, fmt)
        count('DeviceSource', 'Device source records', 'Raw Entra and Intune evidence records; multiple records may describe one CMDB device.')
        for entity, dim, pk, table, name in [
            ('Device','DimDevice','TenantDeviceKey','DimDevice','Device details'),
            ('Device','DimDevice','TenantDeviceKey','DeviceSource','Device source details'),
            ('Device','DimDevice','TenantDeviceKey','EntityFinding','Device finding details'),
            ('User','DimUser','TenantUserKey','DimUser','User details'),
            ('User','DimUser','TenantUserKey','FactUserDeviceRelationship','User device details'),
            ('User','DimUser','TenantUserKey','FactMailbox','User mailbox details'),
            ('User','DimUser','TenantUserKey','LicenseAssignmentPath','User assignment details'),
            ('User','DimUser','TenantUserKey','EntityFinding','User finding details'),
            ('Group','DimGroup','TenantGroupKey','DimGroup','Group details'),
            ('Group','DimGroup','TenantGroupKey','LicenseAssignmentPath','Group assignment details'),
            ('Group','DimGroup','TenantGroupKey','EntityFinding','Group finding details'),
        ]:
            expected = None
            if len(data[dim]) == 1:
                selected = data[dim][0][pk]
                matched = data[table] if table == dim else [r for r in data[table] if r.get(pk) == selected]
                expected = len(matched) or None
            # A slicer filters the selection label, not the key column. Using
            # ALLSELECTED(key) here loses that indirect restriction and hides
            # valid rows. Restore the selected dimension as a whole, also
            # removing all of the detail visual's row-grouping columns.
            add(table, name, f"IF(COUNTROWS(ALLSELECTED('{dim}')) = 1, COUNTROWS('{table}'))", expected,
                f"Detail rows only when exactly one {entity.lower()} is selected outside visual row grouping. Blank otherwise.")
    return result


def lit(value):
    v = ("true" if value else "false") if isinstance(value, bool) else str(value) + "D" if isinstance(value, (int, float)) else "'" + value.replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": v}}}


def color(value):
    return {"solid": {"color": lit(value)}}


def projection(table, property_name, label=None, measure=False):
    original_name = property_name
    if not measure and property_name in PRESENTATION_COLUMNS.get(table, []):
        property_name += "Label"
    return {"field": {"Measure" if measure else "Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": property_name}},
            "queryRef": table + "." + property_name, "nativeQueryRef": label or original_name,
            "displayName": label or original_name}


def exclude_synthetic_blank(visual, table, column):
    """Hide the relationship-generated blank member, not missing source values.

    Every physical row has a nonblank display label, including Not provided.
    This filter cannot remove an actual source row or change a KPI denominator.
    """
    field = projection(table, column)["field"]
    actual_column = field["Column"]["Property"]
    visual["filterConfig"] = {"filters": [{
        "name": visual["name"] + "nonblank", "field": field,
        "type": "Categorical", "howCreated": "User",
        "filter": {"Version": 2, "From": [{"Name": "s", "Entity": table, "Type": 0}],
                   "Where": [{"Condition": {"Not": {"Expression": {"In": {
                       "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": actual_column}}],
                       "Values": [[{"Literal": {"Value": "null"}}]]
                   }}}}}]}
    }]}


def report_pages(measure_defs, include_360=False):
    home = {m["name"]: m["table"] for m in measure_defs}
    pages = []
    navy, blue, teal, muted = "#1F2937", "#005A9E", "#16877B", "#526577"
    def page(name, title, subtitle):
        p = {"name": name, "title": title, "visuals": []}
        pages.append(p)
        short_title = re.sub(r"^\d+\s+", "", title)
        text(p, "Smart Workplace CMDB — " + short_title, 24, 14, 900, 42, 26, navy,
             font="Segoe UI Semibold", weight="bold")
        text(p, "VERSION 1.0.0", 24, 60, 900, 26, 11, muted)
        text(p, subtitle, 24, 96, 1232, 44, 13, muted)
        text(p, "Frozen snapshot · Navigation: tabs below · Selections filter related visuals", 24, 864, 1232, 28, 10, muted)
        return p
    def visual(p, kind, title, x, y, w, h, query=None, objects=None):
        n = len(p["visuals"])
        container = {"title": [{"properties": {"show": lit(True), "text": lit(title), "fontSize": lit(12), "fontColor": color(navy)}}],
                     "background": [{"properties": {"show": lit(True), "color": color("#FFFFFF"), "transparency": lit(0)}}],
                     "border": [{"properties": {"show": lit(True), "color": color("#DDE7F0"), "radius": lit(8)}}],
                     "visualHeader": [{"properties": {"show": lit(False)}}],
                     "padding": [{"properties": {"top": lit(8), "bottom": lit(8), "left": lit(10), "right": lit(10)}}]}
        v = {"$schema": BASE + "fabric/item/report/definition/visualContainer/2.9.0/schema.json", "name": p["name"] + "v" + str(n),
             "position": {"x": x, "y": y, "width": w, "height": h, "z": n, "tabOrder": n},
             "visual": {"visualType": kind, "objects": objects or {}, "visualContainerObjects": container, "drillFilterOtherVisuals": True}}
        if query:
            v["visual"]["query"] = {"queryState": query}
        p["visuals"].append(v)
        return v
    def text(p, value, x, y, w, h, size=14, ink=muted, font="Segoe UI", weight=None):
        def run(line):
            style = {"fontFamily": font, "fontSize": str(size) + "px", "color": ink}
            if weight:
                style["fontWeight"] = weight
            return {"value": line, "textStyle": style}
        paragraphs = [{"textRuns": [run(line)], "horizontalTextAlignment": "left"} for line in value.splitlines()]
        v = visual(p, "textbox", "", x, y, w, h, objects={"general": [{"properties": {"paragraphs": paragraphs}}]})
        v["visual"]["visualContainerObjects"] = {"title": [{"properties": {"show": lit(False)}}], "background": [{"properties": {"show": lit(False)}}], "border": [{"properties": {"show": lit(False)}}], "padding": [{"properties": {side: lit(0) for side in ("top", "bottom", "left", "right")}}]}
    def card(p, name, title, i, y=224):
        visual(p, "cardVisual", title, 24 + i * 312, y, 296, 116,
               {"Data": {"projections": [projection(home[name], name, title, True)]}},
               {"value": [{"properties": {"show": lit(True), "fontSize": lit(24), "bold": lit(True), "fontColor": color(blue), "labelDisplayUnits": lit(1), "labelPrecision": lit(1 if name.endswith("share") else 0), "showBlankAs": lit("Select one SKU" if name.startswith("SKU ") else "No value")}, "selector": {"id": "default"}}],
                "label": [{"properties": {"show": lit(False), "fontSize": lit(12)}, "selector": {"id": "default"}}],
                "outline": [{"properties": {"show": lit(False)}, "selector": {"id": "default"}}],
                "padding": [{"properties": {"paddingUniform": lit(4)}, "selector": {"id": "default"}}],
                "layout": [{"properties": {"paddingUniform": lit(0)}, "selector": {"id": "default"}}],
                "spacing": [{"properties": {"verticalSpacing": lit(0)}, "selector": {"id": "default"}}]})
    def slicer(p, table, c, title, i, total=3):
        w = (1232 - 16 * (total - 1)) / total
        # Reserve a complete dropdown row below the container title. At the
        # previous 60px height Desktop clipped the selected value and arrow.
        v = visual(p, "slicer", title, 24 + i * (w + 16), 144, w, 76,
               {"Values": {"projections": [projection(table, c, title)]}},
               {"data": [{"properties": {"mode": lit("Dropdown")}}], "header": [{"properties": {"show": lit(False)}}]})
        v["visual"]["visualContainerObjects"]["padding"][0]["properties"].update(top=lit(4), bottom=lit(4))
        exclude_synthetic_blank(v, table, c)
    def bar(p, table, c, measure, title, x=24, y=358, w=608, h=232):
        v = visual(p, "barChart", title, x, y, w, h,
                    {"Category": {"projections": [projection(table, c)]}, "Y": {"projections": [projection(home[measure], measure, measure, True)]}},
                    {"dataPoint": [{"properties": {"defaultColor": color(blue)}}], "labels": [{"properties": {"show": lit(True), "fontSize": lit(11), "labelDisplayUnits": lit(1), "labelPrecision": lit(0)}}],
                     "categoryAxis": [{"properties": {"fontSize": lit(11)}}], "valueAxis": [{"properties": {"fontSize": lit(10), "labelDisplayUnits": lit(1), "labelPrecision": lit(0)}}]})
        exclude_synthetic_blank(v, table, c)
        v["visual"]["query"]["sortDefinition"] = {"sort": [{"field": projection(home[measure], measure, measure, True)["field"], "direction": "Descending"}], "isDefaultSort": True}
    def table(p, fields, title, x=24, y=608, w=1232, h=244):
        visual(p, "tableEx", title, x, y, w, h, {"Values": {"projections": [projection(t, c, label, is_measure) for t, c, label, is_measure in fields]}},
               {"columnHeaders": [{"properties": {"fontColor": color("#FFFFFF"), "backColor": color(navy), "bold": lit(True), "fontSize": lit(11), "wordWrap": lit(True)}}],
                "values": [{"properties": {"fontSize": lit(11), "wordWrap": lit(True), "backColorPrimary": color("#FFFFFF"), "backColorSecondary": color("#F2F6FA")}}]})
    def cols(t, fields):
        return [(t, c, label, False) for c, label in fields]
    p = page("overview", "01  Overview", "Cloud inventory and findings. Counts cover different types of objects and must not be added together.")
    text(p, "See Quality for source scope and status. No global tenant filter: this project contains one tenant, checked on import.", 24, 154, 1232, 48)
    for i, (m, title) in enumerate([("Users", "Entra accounts"), ("Devices", "Reconciled devices"), ("Mailboxes", "Mailboxes"), ("Warnings", "Warnings to review")]):
        card(p, m, title, i)
    bar(p, "DimDevice", "ComplianceState", "Devices", "Devices by compliance state")
    bar(p, "DimUser", "AccountStatusLabel", "Users", "Accounts by enabled status", x=648)
    for i, (m, title) in enumerate([("Groups", "Entra groups"), ("Users with assignments", "Users with assignments¹"), ("User-device links", "User-device links"), ("Information", "Quality information")]):
        card(p, m, title, i, y=608)
    text(p, "¹ At least one license assignment across all states. This count does not measure usage or potential savings.\nMissing compliance remains unknown. No time trend: the exports represent a single collection.", 24, 740, 1232, 84)
    p = page("devices", "02  Devices and compliance", "Review compliance, management and user links separately. The compliant share uses all filtered devices as its denominator.")
    for i, (c, title) in enumerate([("OperatingSystem", "Operating system"), ("ComplianceState", "Compliance"), ("ManagementState", "Management")]):
        slicer(p, "DimDevice", c, title, i)
    for i, (m, title) in enumerate([("Devices", "Devices"), ("Compliant devices", "Explicitly compliant"), ("Compliant device share", "Compliant device share"), ("Devices with a resolved user", "With a resolved user")]):
        card(p, m, title, i)
    bar(p, "DimDevice", "OperatingSystem", "Devices", "Devices by operating system")
    bar(p, "DimDevice", "ComplianceState", "Devices", "Devices by compliance state", x=648)
    table(p, cols("DimDevice", [("DeviceName", "Device"), ("OperatingSystem", "System"), ("OperatingSystemVersion", "Version"), ("Ownership", "Ownership"), ("ManagementState", "Management"), ("ComplianceState", "Compliance")]) + [("DimDevice", "Devices", "Count", True)], "Filtered inventory — one row per combination of displayed values")
    p = page("users", "03  Users and relationships", "An enabled account does not prove usage. Relationships below include only resolved matches.")
    for i, (c, title) in enumerate([("Department", "Department"), ("UserType", "Account type"), ("AccountStatusLabel", "Account status")]):
        slicer(p, "DimUser", c, title, i)
    for i, (m, title) in enumerate([("Users", "Users"), ("Enabled accounts", "Enabled accounts"), ("Users with assignments", "With a license assignment"), ("Users with a resolved device", "With a resolved device")]):
        card(p, m, title, i)
    bar(p, "DimUser", "Department", "Users", "Users by department — scroll as needed")
    bar(p, "DimUser", "UserType", "Users", "Users by type", x=648)
    table(p, cols("DimUser", [("DisplayName", "Name"), ("UserPrincipalName", "Account"), ("Department", "Department"), ("AccountStatusLabel", "Status")]) + [("FactUserLicense", "License assignments", "Assignments", True), ("FactUserDeviceRelationship", "User-device links", "Device links", True)], "Filtered users and relationships — personal data, private use")
    p = page("licenses", "04  Licenses and assignments", "Assignments and reported consumption are different metrics. All states are included unless filtered. No price or usage is inferred.")
    slicer(p, "DimLicenseSku", "SkuPartNumber", "Product / SKU", 0, 2)
    slicer(p, "FactUserLicense", "AssignmentState", "Assignment state (assignments only)", 1, 2)
    for i, (m, title) in enumerate([("License SKUs", "License SKUs¹"), ("License assignments", "Assignments"), ("Users with assignments", "Users with assignments"), ("SKU consumed units", "Consumed — select one SKU¹")]):
        card(p, m, title, i)
    bar(p, "DimLicenseSku", "SkuPartNumber", "License assignments", "Assignments by product — scroll as needed")
    bar(p, "FactUserLicense", "AssignmentState", "License assignments", "Assignments by state", x=648)
    table(p, cols("DimLicenseSku", [("SkuPartNumber", "Product")]) + [("DimLicenseSku", "SKU enabled units", "Enabled units¹", True), ("DimLicenseSku", "SKU consumed units", "Consumed units¹", True), ("FactUserLicense", "License assignments", "Filtered assignments", True)], "Capacity by SKU — ¹ unaffected by the assignment-state filter")
    p = page("mailboxes", "05  Mailboxes", "Technical and shared mailboxes are included in the inventory. Interpret mailboxes without a resolved user in the context of their type.")
    for i, (c, title) in enumerate([("RecipientTypeDetails", "Mailbox type"), ("ArchiveStatus", "Archive status"), ("LinkStatusLabel", "User link")]):
        slicer(p, "FactMailbox", c, title, i)
    card(p, "Mailboxes", "Mailboxes", 0)
    card(p, "Mailboxes without a resolved user", "Without a resolved user", 1)
    text(p, "DiscoveryMailbox: a technical mailbox without an external identifier may produce an informational finding. See Quality for its classification.", 648, 232, 608, 90)
    bar(p, "FactMailbox", "RecipientTypeDetails", "Mailboxes", "Mailboxes by type")
    bar(p, "FactMailbox", "ArchiveStatus", "Mailboxes", "Mailboxes by archive status", x=648)
    table(p, cols("FactMailbox", [("DisplayName", "Name"), ("PrimarySmtpAddress", "Primary address"), ("RecipientTypeDetails", "Type"), ("ArchiveStatus", "Archive"), ("LinkStatusLabel", "User link")]), "Filtered mailboxes — personal data, private use")
    p = page("quality", "06  Quality and coverage", "All findings remain visible, including information. Dates below are UTC and describe the frozen export, not a tenant refresh.")
    slicer(p, "FactDataQuality", "SeverityLabel", "Severity (findings only)", 0, 2)
    slicer(p, "FactDataQuality", "FindingType", "Finding type (findings only)", 1, 2)
    for i, (m, title) in enumerate([("Quality findings", "Retained findings"), ("Critical findings", "Critical findings"), ("Warnings", "Warnings"), ("Information", "Information")]):
        card(p, m, title, i)
    table(p, cols("FactDataQuality", [("SeverityLabel", "Severity"), ("FindingType", "Type"), ("Description", "Finding"), ("RecommendedAction", "Suggested action")]), "Filtered findings — review before taking any tenant action", y=348, h=226)
    table(p, cols("SourceHealth", [("SourceName", "Source"), ("Status", "Status"), ("Coverage", "Coverage"), ("SourceRows", "Rows"), ("CompletedDateTime", "Completed UTC"), ("Evidence", "CSV evidence")]), "All source coverage — unaffected by finding filters", y=592, h=220)
    text(p, "Complete = collection reported complete for granted permissions, not tenant certification. Missing sources remain Not collected; see the manifest for collection start times.", 24, 818, 1232, 42, 12)
    if include_360:
        # Put the same unique selector field in source visuals and drillthrough
        # bindings; identical display names cannot merge unrelated entities.
        for source_page, table_name, selector in [('devices','DimDevice','DeviceSelection'),('users','DimUser','UserSelection')]:
            p = next(p for p in pages if p['name'] == source_page)
            v = next(v for v in p['visuals'] if v['visual']['visualType'] == 'tableEx')
            v['visual']['query']['queryState']['Values']['projections'].insert(0, projection(table_name,selector,'Open 360 / unique selection'))

        def detail_page(name, title, dim, selector, note):
            p = page(name,title,note)
            slicer(p,dim,selector,'Search and select one object',0,1)
            v=p['visuals'][-1]
            v['visual']['objects'].update(selection=[{'properties':{'strictSingleSelect':lit(True),'selectAllCheckboxEnabled':lit(False)}}],
                                          general=[{'properties':{'selfFilterEnabled':lit(True)}}])
            field=projection(dim,selector)['field']
            filter_name=name+'selection'
            p['filterConfig']={'filters':[{'name':filter_name,'field':field,'type':'Categorical','howCreated':'User'}], 'filterSortOrder':'Custom'}
            p['pageBinding']={'name':name+'binding','type':'Drillthrough','referenceScope':'Default','acceptsFilterContext':'None',
                              'parameters':[{'name':name+'parameter','boundFilter':filter_name,'fieldExpr':field}]}
            return p

        def detail_table(p,t,fields,title,gate,y,h,x=24,w=1232):
            table(p,cols(t,fields)+[(home[gate],gate,'Detail rows',True)],title,x=x,y=y,w=w,h=h)

        p=detail_page('device360','07  Device 360','DimDevice','DeviceSelection',
                      'Select one device. Source records retain their own dates. An associated account does not prove exclusive ownership.')
        detail_table(p,'DimDevice', [('DeviceName','Device'),('OperatingSystem','System'),('OperatingSystemVersion','Version'),
            ('Ownership','Ownership'),('ComplianceState','Compliance'),('AssociatedAccount','Associated account'),('AssociationStatus','Association')],
            'Identity and association','Device details',224,116)
        detail_table(p,'DeviceSource',[('SourceSystem','Source'),('ManagementAgent','Agent'),('EnrollmentUtcDateTime','Enrollment UTC'),
            ('EnrollmentType','Enrollment type'),('SelectionStatus','Candidate'),('SourceObjectId','Source record ID')],
            'Enrollment and source identity — scroll horizontally for all columns','Device source details',350,158)
        detail_table(p,'DeviceSource',[('SourceSystem','Source'),('ActivityKind','Signal'),('ActivityRaw','Original date text'),
            ('ActivityStatus','Date qualification'),('ActivityUtcDateTime','Qualified UTC'),('SourceCollectedDateTime','Collected UTC')],
            'Activity evidence — ambiguous dates remain unqualified','Device source details',518,160)
        detail_table(p,'EntityFinding',[('Severity','Severity'),('FindingType','Finding'),('Description','Evidence'),('RecommendedAction','Suggested review')],
            'Findings linked to this device','Device finding details',688,160)

        p=detail_page('user360','08  User 360','DimUser','UserSelection',
                      'Select one account. Paths explain assigned licenses, not usage. Managers and user activity have not been collected.')
        detail_table(p,'DimUser',[('DisplayName','Name'),('UserPrincipalName','Account'),('JobTitle','Job title'),('Department','Department'),
            ('AccountStatusLabel','Status'),('CreationRaw','Creation — original text'),('CreationStatus','Date qualification'),
            ('SourceCollectedDateTime','Collected UTC')], 'Account profile — scroll horizontally for all columns','User details',224,116)
        detail_table(p,'FactUserDeviceRelationship',[('DeviceSelection','Device and identity'),('DeviceCompliance','Compliance'),('DeviceSyncStatus','Sync qualification')],
            'Resolved device associations','User device details',350,158,w=608)
        p['visuals'][-1]['visual']['query']['queryState']['Values']['projections'][0] = projection('DimDevice','DeviceSelection','Open Device 360')
        detail_table(p,'FactMailbox',[('PrimarySmtpAddress','Mailbox'),('RecipientTypeDetails','Type'),('ArchiveStatus','Archive')],
            'Associated mailboxes','User mailbox details',350,158,x=648,w=608)
        detail_table(p,'LicenseAssignmentPath',[('Product','Product'),('AssignmentRoute','Route'),('GroupName','Group'),('AssignmentState','State'),
            ('AssignmentError','Reported error'),('AssignmentUpdatedUtcDateTime','Last updated UTC'),('DisabledPlanIds','Disabled plan IDs'),
            ('TenantAssignmentPathKey','Path identity')], 'License paths — every route retained; last updated is not initial assignment date','User assignment details',518,160)
        p['visuals'][-1]['visual']['query']['queryState']['Values']['projections'].insert(3,projection('DimGroup','GroupSelection','Open Group 360'))
        detail_table(p,'EntityFinding',[('Severity','Severity'),('FindingType','Finding'),('Description','Evidence'),('RecommendedAction','Suggested review')],
            'Findings directly linked to this account','User finding details',688,160)

        p=detail_page('group360','09  Group 360','DimGroup','GroupSelection',
                      'Select one group. License paths are not membership. No observed path does not mean unused or safe to delete.')
        p['visuals'][-1]['position']['width']=904
        slicer(p,'DimGroup','ObservedPathStatus','License path coverage',1,2)
        p['visuals'][-1]['position'].update(x=944,width=312)
        detail_table(p,'DimGroup',[('DisplayName','Group'),('MailEnabled','Mail enabled'),('SecurityEnabled','Security enabled'),('GroupTypes','Type flags'),
            ('MemberStatus','Members'),('OwnerStatus','Owners'),('ObservedPathCount','Observed paths'),('SourceCollectedDateTime','Collected UTC')],
            'Group profile','Group details',224,116)
        detail_table(p,'LicenseAssignmentPath',[('Account','Account'),('Product','Product'),('AssignmentState','State'),('AssignmentError','Reported error'),
            ('AssignmentUpdatedUtcDateTime','Last updated UTC'),('DisabledPlanIds','Disabled plan IDs'),('TenantAssignmentPathKey','Path identity')],
            'Observed license assignment paths from this group','Group assignment details',350,298)
        p['visuals'][-1]['visual']['query']['queryState']['Values']['projections'][0] = projection('DimUser','UserSelection','Open User 360')
        detail_table(p,'EntityFinding',[('Severity','Severity'),('FindingType','Finding'),('Description','Evidence'),('RecommendedAction','Suggested review')],
            'Findings directly linked to this group','Group finding details',660,188)

        # Keep full identity and gate fields in the query, but use native
        # projection visibility to remove technical columns from the canvas.
        # Hidden fields still preserve row grain and drillthrough bindings.
        layouts = {
            'device360': [(224,110),(344,190),(544,198),(752,100)],
            'user360': [(224,160),(394,164),(394,164),(568,170),(748,104)],
            'group360': [(224,160),(394,304),(708,144)],
        }
        widths = {
            'Device':180,'System':110,'Version':150,'Ownership':110,
            'Compliance':130,'Associated account':270,'Association':140,
            'Name':160,'Account':240,'Job title':250,'Department':150,'Status':90,
            'Creation — original text':190,'Date qualification':180,'Collected UTC':170,
            'Group':230,'Mail enabled':110,'Security enabled':130,'Type flags':130,
            'Members':130,'Owners':130,'Observed paths':130,
            'Source':100,'Agent':170,'Enrollment UTC':170,'Enrollment type':170,
            'Candidate':140,'Source record ID':260,'Signal':190,'Original date text':210,
            'Qualified UTC':170,'Mailbox':310,'Type':160,'Archive':90,
            'Sync qualification':190,'Product':200,'Route':90,'State':100,
            'Reported error':160,'Last updated UTC':170,'Disabled plan IDs':200,
            'Severity':100,'Finding':230,'Evidence':440,'Suggested review':410,
        }
        for p in pages:
            if p['name'] not in layouts:
                continue
            tables=[v for v in p['visuals'] if v['visual']['visualType']=='tableEx']
            for v,(y,h) in zip(tables,layouts[p['name']],strict=True):
                v['position'].update(y=y,height=h)
                vis=v['visual']
                projections=vis['query']['queryState']['Values']['projections']
                if p['name']=='user360' and any(x['queryRef']=='DimDevice.DeviceSelection' for x in projections):
                    projections.insert(0,projection('DimDevice','DeviceName','Device'))
                if p['name']=='group360' and any(x['queryRef']=='DimUser.UserSelection' for x in projections):
                    projections[:0]=[projection('DimUser','DisplayName','Name'),projection('DimUser','UserPrincipalName','Account')]
                for pr in projections:
                    kind,f=next(iter(pr['field'].items()))
                    if (kind=='Measure' and f['Property'].endswith('details')) or f['Property'] in (
                        'DeviceSelection','UserSelection','GroupSelection','TenantAssignmentPathKey'):
                        pr['hidden']=True
                objects=vis['objects']
                objects['total']=[{'properties':{'totals':lit(False)}}]
                objects['columnHeaders'][0]['properties']['autoSizeColumnWidth']=lit(False)
                objects['grid']=[{'properties':{'rowPadding':lit(3)}}]
                objects['columnWidth']=[{'selector':{'metadata':pr['queryRef']},
                    'properties':{'value':lit(widths.get(pr['displayName'],170))}}
                    for pr in projections if not pr.get('hidden')]
                if p['name']=='device360':
                    detail_widths={'DeviceSource.SourceObjectId':310,
                        'DeviceSource.EnrollmentType':210,'DeviceSource.ActivityKind':220,
                        'DeviceSource.ActivityRaw':240,'DeviceSource.ActivityStatus':230}
                    for column in objects['columnWidth']:
                        if column['selector']['metadata'] in detail_widths:
                            column['properties']['value']=lit(detail_widths[column['selector']['metadata']])
    return pages


def build(root, output, include_360=False):
    root, output = root.resolve(), output.resolve()
    if output.exists() or output == root or root in output.parents or output in root.parents:
        raise ValueError("Output must be a new directory outside the source DATA-LAST")
    source_baseline = {p:sha(p) for p in root.rglob('*') if p.is_file()}
    identity, data, columns, inputs = prepare_data(root)
    relationships = RELATIONSHIPS[:]
    if include_360:
        relationships += enrich_360(root, identity, data, columns, inputs)
    if source_baseline != {p:sha(p) for p in root.rglob('*') if p.is_file()}:
        raise ValueError('Source changed during report preparation')
    ms = measures(data)
    pages = report_pages(ms, include_360)
    for p in pages:
        for v in p["visuals"]:
            box = v["position"]
            if box["x"] < 0 or box["y"] < 0 or box["x"] + box["width"] > 1280 or box["y"] + box["height"] > 900:
                raise ValueError("Visual outside page: " + v["name"])
            for role in v["visual"].get("query", {}).get("queryState", {}).values():
                for proj in role["projections"]:
                    typ, field = next(iter(proj["field"].items()))
                    t, c = field["Expression"]["SourceRef"]["Entity"], field["Property"]
                    if (typ == "Column" and c not in columns[t]) or (typ == "Measure" and (t, c) not in {(m["table"], m["name"]) for m in ms}):
                        raise ValueError("Unbound visual field")
    model_dir = output / "CMDB-REPORTS.SemanticModel"
    report_dir = output / "CMDB-REPORTS.Report"
    tables = []
    for name, rows in data.items():
        path = output / "ReportData" / (name + ".csv")
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=columns[name])
            writer.writeheader()
            writer.writerows(rows)
        tc = []
        for c in columns[name]:
            col = dict(name=c, dataType=type_of(c)[0], sourceColumn=c, summarizeBy="none")
            if c in IDENTITY or c.startswith("Tenant") and c.endswith("Key") or c.startswith("Cmdb") or c in ("FindingId", "SkuId"):
                col["isHidden"] = True
            if type_of(c)[0] == "dateTime":
                col["formatString"] = "yyyy-MM-dd" if c == "Date" else "yyyy-MM-dd HH:mm:ss"
            tc.append(col)
        table = {"name": name, "columns": tc, "partitions": [{"name": name, "mode": "import", "source": {"type": "m", "expression": m_query(path, columns[name], identity, name != "DimDate")}}],
                 "measures": [{k: m[k] for k in ("name", "expression", "description", "formatString")} for m in ms if m["table"] == name]}
        if name in ("DimDate", "DimTenant"):
            table["isHidden"] = True
        tables.append(table)
    write_json(model_dir / "model.bim", {"compatibilityLevel": 1600, "model": {"culture": "en-US", "defaultPowerBIDataSourceVersion": "powerBI_V3", "tables": tables,
               "relationships": [dict(name="cmdb-report-" + str(i), fromTable=f, fromColumn=fk, toTable=d, toColumn=dk, fromCardinality="many", toCardinality="one", crossFilteringBehavior="oneDirection", isActive=True) for i, (f, fk, d, dk) in enumerate(relationships)]}})
    write_json(model_dir / "definition.pbism", {"$schema": BASE + "fabric/item/semanticModel/definitionProperties/1.0.0/schema.json", "version": "4.2", "settings": {"qnaEnabled": False}})
    write_json(output / "SmartWorkplaceCMDB.pbip", {"$schema": BASE + "fabric/pbip/pbipProperties/1.0.0/schema.json", "version": "1.0", "artifacts": [{"report": {"path": report_dir.name}}], "settings": {"enableAutoRecovery": True}})
    write_json(report_dir / "definition.pbir", {"$schema": BASE + "fabric/item/report/definitionProperties/2.0.0/schema.json", "version": "4.0", "datasetReference": {"byPath": {"path": "../" + model_dir.name}}})
    definition = report_dir / "definition"
    write_json(definition / "version.json", {"$schema": BASE + "fabric/item/report/definition/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    write_json(definition / "report.json", {"$schema": BASE + "fabric/item/report/definition/report/3.2.0/schema.json", "themeCollection": {}, "settings": {"defaultDrillFilterOtherVisuals": True, "pagesPosition": "Bottom"}, "annotations": [{"name": "defaultPage", "value": pages[0]["name"]}]})
    write_json(definition / "pages/pages.json", {"$schema": BASE + "fabric/item/report/definition/pagesMetadata/1.0.0/schema.json", "pageOrder": [p["name"] for p in pages], "activePageName": pages[0]["name"]})
    for p in pages:
        path = definition / "pages" / p["name"]
        page_json = {"$schema": BASE + "fabric/item/report/definition/page/2.1.0/schema.json", "name": p["name"], "displayName": p["title"], "displayOption": "FitToPage", "height": 900, "width": 1280,
                     "objects": {
                         "background": [{"properties": {"color": color("#F5F8FB"), "transparency": lit(0)}}],
                         "outspace": [{"properties": {"color": color("#E8EEF5"), "transparency": lit(0)}}],
                     },
                     "filterConfig": p.get('filterConfig', {"filters": [], "filterSortOrder": "Custom"})}
        if 'pageBinding' in p:
            page_json['pageBinding'] = p['pageBinding']
        write_json(path / "page.json", page_json)
        for v in p["visuals"]:
            write_json(path / "visuals" / v["name"] / "visual.json", v)
    dax = "DEFINE\n" + "\n".join(" MEASURE '" + m["table"] + "'[" + m["name"] + "] = " + m["expression"] for m in ms)
    dax += "\nEVALUATE ROW(\n" + ",\n".join(' "' + m["name"] + '", [' + m["name"] + "]" for m in ms) + "\n)\n"
    (output / "VALIDATION-MEASURES.dax").write_text(dax, encoding="utf-8")
    write_json(output / "REPORT-MANIFEST.json", {"version": VERSION, "preparedUtc": dt.datetime.now(dt.timezone.utc).isoformat(), "sourceRoot": str(root), "pages": [p["title"] for p in pages],
               "visualCount": sum(len(p["visuals"]) for p in pages), "counts": {n: len(r) for n, r in data.items()}, "measures": ms, "inputSha256": inputs,
               "presentationColumns": PRESENTATION_COLUMNS, "include360": include_360,
               "relationships": relationships,
               "reportCsvSha256": {str(p.relative_to(output)): sha(p) for p in (output / "ReportData").glob("*.csv")},
               "limitations": ["Stable V1 local preparation; Desktop visual rendering and refresh still require validation for each generated snapshot.", "One tenant only; no RLS or multitenant report security claim.", "Frozen export; refreshing this project does not collect new tenant data.", "No historical trend, price, usage, or AD qualification.", "SourceHealth is disconnected and ignores finding filters."]})
    (output / "OPEN-REPORT.txt").write_text("SmartWorkplaceCMDB — Version 1.0.0\n\n1. Open SmartWorkplaceCMDB.pbip in Power BI Desktop.\n2. Refresh: only local ReportData CSV files are read. No tenant access.\n3. Check all pages, filters and totals against REPORT-MANIFEST.json.\n4. On 360 pages, search and select one object. Detail tables stay blank without a single selection. Right-click a unique selection in the inventory to drill through; page tabs return to the overview.\n5. Save As the current SmartWorkplaceCMDB.pbix after review, replacing the previous version if no longer needed.\n\nThis project contains personal data. Keep it in the authorized private data folder; do not publish, share or add it to Git.\nFrozen snapshot: generate a new folder for a new authorized collection. No validated history.\nThe operator still needs to verify the new pages in Desktop.\n", encoding="utf-8")
    print(json.dumps({"version": VERSION, "pages": len(pages), "visuals": sum(len(p["visuals"]) for p in pages), "measures": len(ms), "tables": len(tables)}, ensure_ascii=False))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--include-360", action='store_true', help='Add Device/User/Group 360 from existing raw and CMDB CSVs')
    args = parser.parse_args()
    build(args.data_root, args.output, args.include_360)
