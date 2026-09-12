"""Create the compact decision-oriented CMDB cockpit without rebuilding data.

The deterministic brownfield transformation keeps the reviewed semantic
bindings and three drill-through identities while consolidating the report into
seven decision pages plus three visible 360 detail pages. Unsupported lifecycle
and business-service metrics remain explicitly unavailable instead of being
fabricated.
"""

from __future__ import annotations

import argparse
import copy
import csv
import hashlib
import json
import shutil
from pathlib import Path

from build_report import exclude_synthetic_blank, lit, m_query, projection, type_of


MANAGED_PAGES = {"overview", "risk", "lifecycle", "businessservices", "devices", "users", "device360"}
RETIRED_PAGES = {"quality", "transformation", "impact", "mailboxes", "hardwarefleet", "hardware"}
DRILLTHROUGH_PAGES = {"device360", "user360", "group360"}
FLEET_TABLE_FIELDS = [
    ("DimDevice", "DeviceSelection", "Open Device 360", False),
    ("DimDevice", "DeviceNameLabel", "Device", False),
    ("DimDevice", "OperatingSystemLabel", "System", False),
    ("DimDevice", "OperatingSystemVersionLabel", "Version", False),
    ("DimDevice", "OwnershipLabel", "Ownership", False),
    ("DimDevice", "ManagementStateLabel", "Management", False),
    ("DimDevice", "ComplianceStateLabel", "Compliance", False),
    ("DimDevice", "Devices", "Count", True),
]
PAGE_ORDER = [
    "overview", "risk", "lifecycle", "licenses", "devices", "users",
    "businessservices", "device360", "user360", "group360",
]
DISPLAY_NAMES = {
    "overview": "01  Executive Overview",
    "risk": "02  Workplace Health",
    "lifecycle": "03  Transformation & Lifecycle",
    "licenses": "04  Licensing & Assignments",
    "devices": "05  Fleet & Hardware",
    "users": "06  People & Messaging",
    "businessservices": "07  Services & Impact",
    "device360": "08  Device 360",
    "user360": "09  User 360",
    "group360": "10  Group 360",
}
HEADER_NAMES = {
    "overview": "Executive Overview",
    "risk": "Workplace Health",
    "lifecycle": "Transformation & Lifecycle",
    "licenses": "Licensing & Assignments",
    "devices": "Fleet & Hardware",
    "users": "People & Messaging",
    "businessservices": "Services & Impact",
    "device360": "Device 360",
    "user360": "User 360",
    "group360": "Group 360",
}

IDENTITY_COLUMNS = ["TenantKey", "OrganizationKey", "EnvironmentKey", "TenantId"]
MAILBOX_HOSTING_COLUMNS = [
    *IDENTITY_COLUMNS,
    "MailboxHostingKey",
    "CountryLabel",
    "HostingLocation",
    "RecipientTypeDetails",
    "EvidenceSource",
]
LICENSE_SUMMARY_SKUS = [
    ("Microsoft 365 F1", "M365_F1"),
    ("Microsoft 365 F3", "SPE_F1"),
    ("Microsoft 365 E3", "SPE_E3"),
    ("Microsoft 365 E5", "SPE_E5"),
    ("Microsoft 365 Copilot", "Microsoft_365_Copilot"),
]


def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write(path: Path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False), encoding="utf-8")


def clone_visual(pages: Path, page: str, visual: str):
    return copy.deepcopy(load(pages / page / "visuals" / visual / "visual.json"))


def set_text(visual, value: str):
    props = visual["visual"]["objects"]["general"][0]["properties"]
    paragraphs = props.get("paragraphs") or []
    run = paragraphs[0]["textRuns"][0] if paragraphs else {"value": ""}
    style = copy.deepcopy(run.get("textStyle", {}))
    props["paragraphs"] = [
        {"textRuns": [{"value": line, "textStyle": copy.deepcopy(style)}],
         "horizontalTextAlignment": "left"}
        for line in value.splitlines()
    ]


def set_title(visual, title: str):
    visual["visual"]["visualContainerObjects"]["title"][0]["properties"]["text"] = lit(title)


def put(visuals, visual, page_id: str, x, y, w, h):
    visual["name"] = f"{page_id}v{len(visuals)}"
    for index, item in enumerate(visual.get("filterConfig", {}).get("filters", [])):
        item["name"] = f"{visual['name']}filter{index}"
    visual["position"] = {
        "x": x, "y": y, "width": w, "height": h,
        "z": len(visuals), "tabOrder": len(visuals),
    }
    visuals.append(visual)
    return visual


def clear_filter(visual):
    visual.pop("filterConfig", None)
    for entry in visual.get("visual", {}).get("objects", {}).get("general", []):
        entry.get("properties", {}).pop("filter", None)
        entry.get("properties", {}).pop("selfFilter", None)


def read_csv(path: Path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        rows = list(reader)
        if not reader.fieldnames or any(None in row for row in rows):
            raise ValueError(f"Malformed CSV: {path}")
        return reader.fieldnames, rows


def write_csv(path: Path, columns, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def normalized_address(row):
    for name in ("PrimarySmtpAddress", "WindowsEmailAddress", "UserPrincipalName", "EmailAddress"):
        value = (row.get(name) or "").strip().lower()
        if value and "@" in value:
            return value
    return ""


def recipient_type(row):
    return (row.get("RecipientTypeDetails") or row.get("RecipientType") or "Unknown").strip() or "Unknown"


def set_categorical_selection(visual, table, column, value):
    props = visual["visual"].setdefault("objects", {}).setdefault("general", [{"properties": {}}])[0].setdefault("properties", {})
    props["filter"] = {
        "filter": {
            "Version": 2,
            "From": [{"Name": "s", "Entity": table, "Type": 0}],
            "Where": [{"Condition": {"In": {
                "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": column}}],
                "Values": [[{"Literal": {"Value": "'" + value.replace("'", "''") + "'"}}]],
            }}}],
        }
    }


def set_single_value_filter(visual, table, column, value, suffix="selection"):
    field = projection(table, column)["field"]
    actual_column = field["Column"]["Property"]
    visual["filterConfig"] = {"filters": [{
        "name": visual["name"] + suffix,
        "field": field,
        "type": "Categorical",
        "howCreated": "User",
        "filter": {
            "Version": 2,
            "From": [{"Name": "s", "Entity": table, "Type": 0}],
            "Where": [{"Condition": {"In": {
                "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": actual_column}}],
                "Values": [[{"Literal": {"Value": "'" + value.replace("'", "''") + "'"}}]],
            }}}],
        },
    }]}


def set_donut(visual, category_table, category_column, measure_table, measure, title):
    clear_filter(visual)
    visual["visual"]["visualType"] = "donutChart"
    visual["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [projection(category_table, category_column)]},
            "Y": {"projections": [projection(measure_table, measure, measure, True)]},
        },
        "sortDefinition": {
            "sort": [{"field": projection(measure_table, measure, measure, True)["field"], "direction": "Descending"}],
            "isDefaultSort": True,
        },
    }
    set_title(visual, title)
    exclude_synthetic_blank(visual, category_table, category_column)


def add_donut(pages, visuals, page_id, category_table, category_column, measure_table, measure, title, x, y, w=240, h=200):
    visual = clone_visual(pages, "devices", "devicesv11")
    set_donut(visual, category_table, category_column, measure_table, measure, title)
    return put(visuals, visual, page_id, x, y, w, h)


def add_country_bar(pages, visuals, page_id, x, y, w, h):
    visual = clone_visual(pages, "licenses", "licensesv10")
    clear_filter(visual)
    visual["visual"]["visualType"] = "clusteredBarChart"
    visual["visual"].get("objects", {}).pop("dataPoint", None)
    measures = [
        ("DimCountry", "Executive corporate country share", "Corporate devices"),
        ("DimCountry", "Executive user country share", "Users"),
        ("DimCountry", "Executive mailbox country share", "Mailboxes"),
    ]
    visual["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [projection("DimCountry", "CountryLabel", "Country")]},
            "Y": {"projections": [projection(table, measure, label, True) for table, measure, label in measures]},
        },
        "sortDefinition": {
            "sort": [{"field": projection("DimCountry", "Executive user country share", "Users", True)["field"], "direction": "Descending"}],
            "isDefaultSort": True,
        },
    }
    set_title(visual, "Country footprint — population distribution")
    exclude_synthetic_blank(visual, "DimCountry", "CountryLabel")
    for entry in visual["visual"].get("objects", {}).get("labels", []):
        entry.get("properties", {})["labelPrecision"] = lit(1)
    return put(visuals, visual, page_id, x, y, w, h)


def add_license_summary_card(pages, visuals, page_id, title, sku, x, y, w=240, h=104):
    visual = clone_visual(pages, "devices", "devicesv7")
    clear_filter(visual)
    visual["visual"]["query"] = {
        "queryState": {"Data": {"projections": [
            projection("DimLicenseSku", "SKU allocation summary", "Assigned / enabled · allocation", True)
        ]}}
    }
    for entry in visual["visual"].get("objects", {}).get("value", []):
        entry.setdefault("properties", {})["fontSize"] = lit(14)
        entry["properties"]["textWrap"] = lit(False)
    set_title(visual, title)
    put(visuals, visual, page_id, x, y, w, h)
    set_single_value_filter(visual, "DimLicenseSku", "SkuPartNumber", sku, "sku")
    return visual


def set_page_navigation(visual, page_name, tooltip):
    visual["visual"].setdefault("visualContainerObjects", {})["visualLink"] = [{
        "properties": {
            "show": lit(True),
            "type": lit("PageNavigation"),
            "navigationSection": lit(page_name),
            "tooltip": lit(tooltip),
            "showDefaultTooltip": lit(False),
        }
    }]


def reconcile_mailboxes(fact_mailboxes, remote_rows, local_rows, identity, user_country_by_key, user_country_by_address):
    records = {}

    def add(row, hosting, evidence, replace=False):
        address = normalized_address(row)
        if not address:
            return
        country = (
            user_country_by_key.get(row.get("TenantUserKey", ""))
            or user_country_by_address.get(address)
            or "Unknown / unassigned"
        )
        candidate = {
            **identity,
            "MailboxHostingKey": hashlib.sha256((identity["TenantKey"] + "|" + address).encode("utf-8")).hexdigest().upper(),
            "CountryLabel": country,
            "HostingLocation": hosting,
            "RecipientTypeDetails": recipient_type(row),
            "EvidenceSource": evidence,
        }
        if replace or address not in records:
            records[address] = candidate
        elif records[address]["CountryLabel"] == "Unknown / unassigned" and country != "Unknown / unassigned":
            records[address]["CountryLabel"] = country

    # The report's FactMailbox export is authoritative Exchange Online evidence.
    for row in fact_mailboxes:
        add(row, "Exchange Online", "CMDB FactMailbox / Exchange Online")

    for row in remote_rows:
        if recipient_type(row).lower().startswith("remote"):
            add(row, "Exchange Online", "Exchange RemoteMailbox")

    for row in local_rows:
        if not recipient_type(row).lower().startswith("remote"):
            # Online/Remote evidence wins when an address exists in both exports.
            add(row, "Exchange On-premises", "Exchange local Mailbox")

    return sorted(records.values(), key=lambda row: (row["HostingLocation"], row["MailboxHostingKey"]))


def mailbox_hosting_rows(report: Path, local_path: Path | None, remote_path: Path | None):
    data_dir = report.parent / "ReportData"
    _, tenants = read_csv(data_dir / "DimTenant.csv")
    if len(tenants) != 1 or any(not tenants[0].get(column) for column in IDENTITY_COLUMNS):
        raise ValueError("Exactly one complete tenant identity is required")
    identity = {column: tenants[0][column] for column in IDENTITY_COLUMNS}
    _, users = read_csv(data_dir / "DimUser.csv")
    user_country_by_key = {
        row.get("TenantUserKey", ""): row.get("CountryLabel", "") or "Unknown / unassigned"
        for row in users if row.get("TenantUserKey")
    }
    user_country_by_address = {
        (row.get("UserPrincipalName") or "").strip().lower(): row.get("CountryLabel", "") or "Unknown / unassigned"
        for row in users if row.get("UserPrincipalName")
    }
    _, fact_mailboxes = read_csv(data_dir / "FactMailbox.csv")
    remote_rows = read_csv(remote_path)[1] if remote_path else []
    local_rows = read_csv(local_path)[1] if local_path else []
    rows = reconcile_mailboxes(
        fact_mailboxes, remote_rows, local_rows, identity,
        user_country_by_key, user_country_by_address,
    )

    metadata = {
        "online": sum(row["HostingLocation"] == "Exchange Online" for row in rows),
        "onPremises": sum(row["HostingLocation"] == "Exchange On-premises" for row in rows),
        "total": len(rows),
        "localEvidenceDate": local_path.stat().st_mtime if local_path else None,
        "remoteEvidenceDate": remote_path.stat().st_mtime if remote_path else None,
    }
    return identity, rows, metadata


def add_or_replace_measure(table, name, expression, description, format_string):
    measures = table.setdefault("measures", [])
    measures[:] = [measure for measure in measures if measure.get("name") != name]
    measures.append({
        "name": name,
        "expression": expression,
        "description": description,
        "formatString": format_string,
    })


def enrich_semantic_model(report: Path, local_path: Path | None, remote_path: Path | None):
    identity, hosting_rows, metadata = mailbox_hosting_rows(report, local_path, remote_path)
    data_dir = report.parent / "ReportData"
    hosting_path = data_dir / "FactMailboxHosting.csv"
    write_csv(hosting_path, MAILBOX_HOSTING_COLUMNS, hosting_rows)

    _, users = read_csv(data_dir / "DimUser.csv")
    _, devices = read_csv(data_dir / "DimDevice.csv")
    countries = sorted({
        row.get("CountryLabel") or "Unknown / unassigned"
        for row in [*users, *devices, *hosting_rows]
    })
    country_path = data_dir / "DimCountry.csv"
    write_csv(country_path, ["CountryLabel"], [{"CountryLabel": country} for country in countries])

    model_path = next(report.parent.glob("*.SemanticModel/model.bim"), None)
    if not model_path:
        raise ValueError("Expected one BIM semantic model beside the report")
    model_json = load(model_path)
    model = model_json["model"]
    tables = model["tables"]

    def import_table(name, path, columns, tenant_scoped, measures=None):
        table_columns = []
        for column in columns:
            item = {"name": column, "dataType": type_of(column)[0], "sourceColumn": column, "summarizeBy": "none"}
            if column in IDENTITY_COLUMNS or column.endswith("Key"):
                item["isHidden"] = True
            table_columns.append(item)
        return {
            "name": name,
            "columns": table_columns,
            "partitions": [{
                "name": name,
                "mode": "import",
                "source": {"type": "m", "expression": m_query(path, columns, identity, tenant_scoped)},
            }],
            "measures": measures or [],
        }

    hosting_table = import_table("FactMailboxHosting", hosting_path, MAILBOX_HOSTING_COLUMNS, True)
    add_or_replace_measure(hosting_table, "Hosted mailboxes", "COALESCE(COUNTROWS('FactMailboxHosting'), 0)", "Distinct mailbox addresses after reconciliation of Exchange Online, RemoteMailbox and local Mailbox evidence.", "#,0")
    add_or_replace_measure(hosting_table, "Exchange Online mailboxes", "CALCULATE([Hosted mailboxes], KEEPFILTERS('FactMailboxHosting'[HostingLocation] == \"Exchange Online\"))", "Mailboxes observed in Exchange Online or as RemoteMailbox; duplicate addresses are counted once.", "#,0")
    add_or_replace_measure(hosting_table, "Exchange On-premises mailboxes", "CALCULATE([Hosted mailboxes], KEEPFILTERS('FactMailboxHosting'[HostingLocation] == \"Exchange On-premises\"))", "Local Mailbox recipients not already observed as Exchange Online or RemoteMailbox; duplicate addresses are counted once.", "#,0")
    add_or_replace_measure(hosting_table, "Exchange Online share", "DIVIDE([Exchange Online mailboxes], [Hosted mailboxes])", "Exchange Online mailboxes divided by the reconciled mailbox footprint.", "0.0%")
    add_or_replace_measure(hosting_table, "Exchange On-premises share", "DIVIDE([Exchange On-premises mailboxes], [Hosted mailboxes])", "Exchange On-premises mailboxes divided by the reconciled mailbox footprint.", "0.0%")

    country_table = import_table("DimCountry", country_path, ["CountryLabel"], False)
    add_or_replace_measure(country_table, "Executive workplace devices", "CALCULATE([Devices], TREATAS(VALUES('DimCountry'[CountryLabel]), 'DimDevice'[CountryLabel]))", "Devices in the selected shared country context; the ownership slicer may further restrict the population.", "#,0")
    add_or_replace_measure(country_table, "Executive managed device share", "CALCULATE([Managed device share], TREATAS(VALUES('DimCountry'[CountryLabel]), 'DimDevice'[CountryLabel]))", "Managed-device share in the selected shared country context.", "0.0%")
    add_or_replace_measure(country_table, "Executive compliant device share", "CALCULATE([Compliant device share], TREATAS(VALUES('DimCountry'[CountryLabel]), 'DimDevice'[CountryLabel]))", "Compliant-device share in the selected shared country context.", "0.0%")
    add_or_replace_measure(country_table, "Executive corporate devices", "CALCULATE([Corporate devices], TREATAS(VALUES('DimCountry'[CountryLabel]), 'DimDevice'[CountryLabel]))", "Corporate devices in the selected shared country context; ownership selection is intentionally ignored by the base measure.", "#,0")
    add_or_replace_measure(country_table, "Executive corporate country share", "VAR _selected = [Executive corporate devices] VAR _all = CALCULATE([Executive corporate devices], REMOVEFILTERS('DimCountry'[CountryLabel])) RETURN DIVIDE(_selected, _all)", "Corporate devices for each country divided by all corporate devices.", "0.0%")
    add_or_replace_measure(country_table, "Executive user country share", "DIVIDE([Users], CALCULATE([Users], REMOVEFILTERS('DimCountry'[CountryLabel])))", "Users for each country divided by all users, including unknown or unassigned country.", "0.0%")
    add_or_replace_measure(country_table, "Executive mailbox country share", "DIVIDE([Hosted mailboxes], CALCULATE([Hosted mailboxes], REMOVEFILTERS('DimCountry'[CountryLabel])))", "Reconciled mailboxes for each country divided by all reconciled mailboxes, including unknown or unassigned country.", "0.0%")

    tables[:] = [table for table in tables if table.get("name") not in {"DimCountry", "FactMailboxHosting"}]
    tables.extend([country_table, hosting_table])
    relationships = model.setdefault("relationships", [])
    relationships[:] = [
        relationship for relationship in relationships
        if relationship.get("name") not in {"cmdb-executive-country-user", "cmdb-executive-country-mailbox"}
    ]
    relationships.extend([
        {
            "name": "cmdb-executive-country-user",
            "fromTable": "DimUser",
            "fromColumn": "CountryLabel",
            "toTable": "DimCountry",
            "toColumn": "CountryLabel",
            "fromCardinality": "many",
            "toCardinality": "one",
            "crossFilteringBehavior": "oneDirection",
            "isActive": True,
        },
        {
            "name": "cmdb-executive-country-mailbox",
            "fromTable": "FactMailboxHosting",
            "fromColumn": "CountryLabel",
            "toTable": "DimCountry",
            "toColumn": "CountryLabel",
            "fromCardinality": "many",
            "toCardinality": "one",
            "crossFilteringBehavior": "oneDirection",
            "isActive": True,
        },
    ])
    write(model_path, model_json)
    return metadata


def set_card(visual, table, measure, title, precision=None):
    visual["visual"]["query"] = {
        "queryState": {"Data": {"projections": [
            projection(table, measure, title, True)
        ]}}
    }
    set_title(visual, title)
    props = visual["visual"]["objects"]["value"][0]["properties"]
    props["labelPrecision"] = lit(1 if precision is None and ("share" in measure.lower() or "rate" in measure.lower() or "coverage" in measure.lower()) else (precision or 0))
    props["showBlankAs"] = lit("No value")


def set_bar(visual, table, column, measure_table, measure, title):
    clear_filter(visual)
    visual["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [projection(table, column)]},
            "Y": {"projections": [projection(measure_table, measure, measure, True)]},
        },
        "sortDefinition": {
            "sort": [{"field": projection(measure_table, measure, measure, True)["field"],
                      "direction": "Descending"}],
            "isDefaultSort": True,
        },
    }
    set_title(visual, title)
    exclude_synthetic_blank(visual, table, column)


def add_text(pages, visuals, page_id, value, x, y, w, h, template="overviewv4"):
    visual = clone_visual(pages, "devices", "devicesv2")
    set_text(visual, value)
    return put(visuals, visual, page_id, x, y, w, h)


def add_card(pages, visuals, page_id, table, measure, title, x, y, w=296, h=96, precision=None):
    visual = clone_visual(pages, "devices", "devicesv7")
    set_card(visual, table, measure, title, precision)
    return put(visuals, visual, page_id, x, y, w, h)


def add_bar(pages, visuals, page_id, table, column, measure_table, measure, title, x, y, w=608, h=220):
    visual = clone_visual(pages, "devices", "devicesv11")
    set_bar(visual, table, column, measure_table, measure, title)
    return put(visuals, visual, page_id, x, y, w, h)


def add_slicer(pages, visuals, page_id, table, column, title, x, y=144, w=608, h=76):
    visual = clone_visual(pages, "devices", "devicesv4")
    clear_filter(visual)
    visual["visual"]["query"] = {
        "queryState": {"Values": {"projections": [projection(table, column, title)]}}
    }
    set_title(visual, title)
    exclude_synthetic_blank(visual, table, column)
    return put(visuals, visual, page_id, x, y, w, h)


def add_table(pages, visuals, page_id, fields, title, x, y, w, h):
    # Licensing remains a stable, visible page before and after consolidation,
    # so its table is the durable template for repeatable brownfield runs.
    visual = clone_visual(pages, "licenses", "licensesv12")
    clear_filter(visual)
    visual["visual"]["query"] = {
        "queryState": {"Values": {"projections": [
            projection(table, field, label, is_measure)
            for table, field, label, is_measure in fields
        ]}}
    }
    set_title(visual, title)
    objects = visual["visual"]["objects"]
    objects.pop("columnWidth", None)
    objects["columnHeaders"][0]["properties"].update(
        autoSizeColumnWidth=lit(True), columnAdjustment=lit("growToFit")
    )
    return put(visuals, visual, page_id, x, y, w, h)


def new_page(pages: Path, page_id: str, title: str, subtitle: str):
    page = copy.deepcopy(load(pages / "devices" / "page.json"))
    page.update(name=page_id, displayName=DISPLAY_NAMES[page_id])
    page.pop("pageBinding", None)
    page.pop("filterConfig", None)
    visuals = []
    for template, value, y, h in [
        ("devicesv0", f"Smart Workplace CMDB — {title}", 14, 42),
        ("devicesv1", "VERSION 1.0.0", 60, 26),
        ("devicesv2", subtitle, 96, 44),
        ("devicesv3", "Frozen snapshot · Navigation: ordered tabs below · Selections filter related visuals", 864, 28),
    ]:
        visual = clone_visual(pages, "devices", template)
        set_text(visual, value)
        put(visuals, visual, page_id, 24, y, 1232, h)
    return page, visuals


def build_risk(pages):
    page, visuals = new_page(
        pages, "risk", "Workplace Health",
        "Prioritize compliance and CMDB-quality exceptions together. Missing states remain separate from explicit noncompliance; no composite health score is inferred.",
    )
    add_slicer(pages, visuals, "risk", "DimDevice", "ComplianceStateLabel", "Compliance", 24, w=296)
    add_slicer(pages, visuals, "risk", "DimDevice", "ManagementStateLabel", "Management", 336, w=296)
    severity_slicer = add_slicer(pages, visuals, "risk", "FactDataQuality", "SeverityLabel", "Finding severity", 648, w=296)
    set_categorical_selection(severity_slicer, "FactDataQuality", "SeverityLabel", "Warning")
    add_slicer(pages, visuals, "risk", "FactDataQuality", "FindingTypeLabel", "Finding type", 960, w=296)
    for index, item in enumerate([
        ("DimDevice", "Noncompliant devices", "Explicitly noncompliant"),
        ("DimDevice", "Other or missing compliance devices", "Other / missing compliance"),
        ("FactDataQuality", "Critical findings", "Critical findings"),
        ("FactDataQuality", "Warnings", "Data-quality warnings"),
    ]):
        add_card(pages, visuals, "risk", *item, 24 + index * 312, 232)
    add_bar(pages, visuals, "risk", "DimDevice", "ComplianceStateLabel", "DimDevice", "Devices", "Devices by compliance state", 24, 340, h=212)
    add_bar(pages, visuals, "risk", "DimDevice", "ManagementStateLabel", "DimDevice", "Devices", "Devices by management state", 648, 340, h=212)
    add_table(
        pages, visuals, "risk",
        [
            ("FactDataQuality", "SeverityLabel", "Severity", False),
            ("FactDataQuality", "FindingTypeLabel", "Finding", False),
            ("FactDataQuality", "DescriptionLabel", "Evidence", False),
            ("FactDataQuality", "RecommendedActionLabel", "Suggested review", False),
        ],
        "Quality findings — review evidence before taking tenant action", 24, 568, 1232, 276,
    )
    return page, visuals


def build_lifecycle(pages):
    page, visuals = new_page(
        pages, "lifecycle", "Transformation & Lifecycle",
        "Assess the observed estate, management coverage and source readiness. No migration plan, purchase date, warranty, support end date or application lifecycle source is present.",
    )
    add_slicer(pages, visuals, "lifecycle", "DimDevice", "OperatingSystemLabel", "Operating system", 24)
    add_slicer(pages, visuals, "lifecycle", "DimDevice", "ManagementStateLabel", "Management", 648)
    for index, item in enumerate([
        ("DimDevice", "Devices", "Observed devices"),
        ("DimDevice", "Managed device share", "Managed device rate"),
        ("DeviceHardware", "Hardware coverage rate", "Hardware coverage"),
        ("SourceHealth", "Source collection coverage", "Source collection coverage"),
    ]):
        add_card(pages, visuals, "lifecycle", *item, 24 + index * 312, 232)
    add_bar(pages, visuals, "lifecycle", "DimDevice", "OperatingSystemLabel", "DimDevice", "Devices", "Observed operating systems — not an EOL classification", 24, 340, h=212)
    add_bar(pages, visuals, "lifecycle", "DimDevice", "ManagementStateLabel", "DimDevice", "Devices", "Observed devices by management state", 648, 340, h=212)
    add_table(
        pages, visuals, "lifecycle",
        [
            ("SourceHealth", "SourceName", "Source", False),
            ("SourceHealth", "Status", "Collection status", False),
            ("SourceHealth", "Coverage", "Coverage", False),
            ("SourceHealth", "SourceRowsLabel", "Rows", False),
            ("SourceHealth", "CompletedDateTimeLabel", "Completed UTC", False),
        ],
        "Source readiness for transformation and future lifecycle decisions", 24, 568, 1232, 276,
    )
    return page, visuals


def build_transformation(pages):
    page, visuals = new_page(
        pages, "transformation", "Workplace transformation",
        "Describe the observed estate and management coverage. The snapshot does not contain migration plans, eligibility, Autopilot, VDI or Golden Image data.",
    )
    add_slicer(pages, visuals, "transformation", "DimDevice", "OperatingSystem", "Operating system", 24)
    add_slicer(pages, visuals, "transformation", "DimDevice", "ManagementState", "Management state", 648)
    for index, item in enumerate([
        ("DimDevice", "Devices", "Observed devices"),
        ("DimDevice", "Managed devices", "Explicitly managed"),
        ("DimDevice", "Managed device share", "Managed device rate"),
        ("DimDevice", "Devices with missing management state", "Missing management state"),
    ]):
        add_card(pages, visuals, "transformation", *item, 24 + index * 312, 232)
    add_bar(pages, visuals, "transformation", "DimDevice", "OperatingSystem", "DimDevice", "Devices", "Observed devices by operating system", 24, 340)
    add_bar(pages, visuals, "transformation", "DimDevice", "ManagementState", "DimDevice", "Devices", "Observed devices by management state", 648, 340)
    for index, item in enumerate([
        ("FactUserDeviceRelationship", "Devices with a resolved user", "Devices with a user link"),
        ("FactUserDeviceRelationship", "Device user-link coverage", "Device / user link coverage"),
        ("DeviceHardware", "Hardware covered devices", "Devices with hardware"),
        ("DeviceHardware", "Hardware data gap rate", "Hardware data gap"),
    ]):
        add_card(pages, visuals, "transformation", *item, 24 + index * 312, 572)
    add_text(
        pages, visuals, "transformation",
        "These are inventory and coverage indicators, not migration progress or readiness. Add authoritative campaign, eligibility, deployment-ring and VDI/application sources before introducing transformation targets.",
        24, 684, 1232, 150,
    )
    return page, visuals


def build_business_services(pages):
    page, visuals = new_page(
        pages, "businessservices", "Services & Impact",
        "Trace the relationships the current CMDB can prove while keeping business-service ownership and criticality explicitly unavailable.",
    )
    add_slicer(pages, visuals, "businessservices", "FactUserDeviceRelationship", "RelationshipType", "Relationship type", 24)
    add_slicer(pages, visuals, "businessservices", "LicenseAssignmentPath", "AssignmentRoute", "Assignment route", 648)
    for index, item in enumerate([
        ("FactUserDeviceRelationship", "User-device links", "Resolved user / device links"),
        ("LicenseAssignmentPath", "Assignment paths", "Observed license paths"),
        ("FactUserDeviceRelationship", "Device user-link coverage", "Device / user link coverage"),
        ("SourceHealth", "Source collection coverage", "Source collection coverage"),
    ]):
        add_card(pages, visuals, "businessservices", *item, 24 + index * 312, 232)
    add_bar(pages, visuals, "businessservices", "FactUserDeviceRelationship", "RelationshipType", "FactUserDeviceRelationship", "User-device links", "User-device links by relationship type", 24, 340, h=212)
    add_bar(pages, visuals, "businessservices", "LicenseAssignmentPath", "AssignmentRoute", "LicenseAssignmentPath", "Assignment paths", "License assignment paths by route", 648, 340, h=212)
    add_table(
        pages, visuals, "businessservices",
        [
            ("LicenseAssignmentPath", "Account", "Account", False),
            ("LicenseAssignmentPath", "Product", "Product", False),
            ("LicenseAssignmentPath", "AssignmentRoute", "Route", False),
            ("LicenseAssignmentPath", "GroupName", "Group", False),
            ("LicenseAssignmentPath", "ErrorStatus", "Error status", False),
        ],
        "Observed paths — no authoritative business-service or application dependency source is present", 24, 568, 1232, 276,
    )
    return page, visuals


def build_fleet_hardware(pages):
    page, visuals = new_page(
        pages, "devices", "Fleet & Hardware",
        "Explore the reconciled device estate and source-reported hardware together. Serial values never define identity or ownership.",
    )
    add_slicer(pages, visuals, "devices", "DimDevice", "OperatingSystemLabel", "Operating system", 24, w=400)
    add_slicer(pages, visuals, "devices", "DimDevice", "OwnershipLabel", "Ownership", 440, w=400)
    add_slicer(pages, visuals, "devices", "DimDevice", "ComplianceStateLabel", "Compliance", 856, w=400)
    for index, item in enumerate([
        ("DimDevice", "Devices", "Workplace devices"),
        ("DimDevice", "Compliant device share", "Device compliance rate"),
        ("DeviceHardware", "Hardware coverage rate", "Hardware coverage"),
        ("DeviceHardware", "Hardware uncovered devices", "Devices without hardware"),
    ]):
        add_card(pages, visuals, "devices", *item, 24 + index * 312, 232)
    add_bar(pages, visuals, "devices", "DimDevice", "ComplianceStateLabel", "DimDevice", "Devices", "Devices by compliance state", 24, 340, h=212)
    add_bar(pages, visuals, "devices", "DeviceHardware", "Manufacturer", "DeviceHardware", "Hardware records", "Hardware records by manufacturer", 648, 340, h=212)
    add_table(
        pages, visuals, "devices",
        FLEET_TABLE_FIELDS,
        "Fleet inventory — right-click a unique device for hardware detail in Device 360", 24, 568, 1232, 276,
    )
    return page, visuals


def build_people_messaging(pages):
    page, visuals = new_page(
        pages, "users", "People & Messaging",
        "Review accounts, observed assignments, device relationships and mailboxes together. Enabled accounts and assigned licenses do not prove activity.",
    )
    add_slicer(pages, visuals, "users", "DimUser", "DepartmentLabel", "Department", 24, w=400)
    add_slicer(pages, visuals, "users", "DimUser", "AccountStatusLabel", "Account status", 440, w=400)
    add_slicer(pages, visuals, "users", "FactMailbox", "RecipientTypeDetailsLabel", "Mailbox type", 856, w=400)
    for index, item in enumerate([
        ("DimUser", "Users", "Users"),
        ("DimUser", "Enabled accounts", "Enabled accounts"),
        ("FactUserLicense", "Users with assignments", "With a license assignment"),
        ("FactMailbox", "Mailboxes", "Mailboxes"),
    ]):
        add_card(pages, visuals, "users", *item, 24 + index * 312, 232)
    add_bar(pages, visuals, "users", "DimUser", "DepartmentLabel", "DimUser", "Users", "Users by department — scroll as needed", 24, 340, h=212)
    add_bar(pages, visuals, "users", "FactMailbox", "RecipientTypeDetailsLabel", "FactMailbox", "Mailboxes", "Mailboxes by type", 648, 340, h=212)
    add_table(
        pages, visuals, "users",
        [
            ("DimUser", "UserSelection", "Open User 360", False),
            ("DimUser", "DisplayNameLabel", "Name", False),
            ("DimUser", "UserPrincipalNameLabel", "Account", False),
            ("DimUser", "AccountStatusLabel", "Status", False),
            ("FactUserLicense", "License assignments", "Assignments", True),
        ],
        "Filtered users — personal data, private use", 24, 568, 608, 276,
    )
    add_table(
        pages, visuals, "users",
        [
            ("FactMailbox", "DisplayNameLabel", "Mailbox", False),
            ("FactMailbox", "PrimarySmtpAddressLabel", "Primary address", False),
            ("FactMailbox", "RecipientTypeDetailsLabel", "Type", False),
            ("FactMailbox", "LinkStatusLabel", "User link", False),
        ],
        "Filtered mailboxes — personal data, private use", 648, 568, 608, 276,
    )
    return page, visuals


def build_device_360(pages):
    page, visuals = new_page(
        pages, "device360", "Device 360",
        "Select one device. Identity, enrollment, activity and hardware retain their own source evidence and dates.",
    )
    page["filterConfig"] = copy.deepcopy(load(pages / "device360" / "page.json").get("filterConfig", {}))
    page["pageBinding"] = copy.deepcopy(load(pages / "device360" / "page.json").get("pageBinding", {}))
    add_slicer(pages, visuals, "device360", "DimDevice", "DeviceSelection", "Search and select one device", 24, w=1232)
    add_table(
        pages, visuals, "device360",
        [
            ("DimDevice", "DeviceNameLabel", "Device", False),
            ("DimDevice", "OperatingSystemLabel", "System", False),
            ("DimDevice", "OperatingSystemVersionLabel", "Version", False),
            ("DimDevice", "OwnershipLabel", "Ownership", False),
            ("DimDevice", "ComplianceStateLabel", "Compliance", False),
            ("DimDevice", "AssociatedAccount", "Associated account", False),
            ("DimDevice", "AssociationStatus", "Association", False),
            ("DimDevice", "Device details", "Detail rows", True),
        ],
        "Identity and association", 24, 232, 1232, 100,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "device360",
        [
            ("DeviceSource", "SourceSystem", "Source", False),
            ("DeviceSource", "ManagementAgent", "Agent", False),
            ("DeviceSource", "EnrollmentUtcDateTime", "Enrollment UTC", False),
            ("DeviceSource", "EnrollmentType", "Enrollment type", False),
            ("DeviceSource", "SelectionStatus", "Candidate", False),
            ("DeviceSource", "SourceObjectId", "Source record ID", False),
            ("DeviceSource", "Device source details", "Detail rows", True),
        ],
        "Enrollment and source identity", 24, 344, 608, 176,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "device360",
        [
            ("DeviceHardware", "ManagedDeviceId", "Source record", False),
            ("DeviceHardware", "SerialNumber", "Serial number", False),
            ("DeviceHardware", "Manufacturer", "Manufacturer", False),
            ("DeviceHardware", "Model", "Model", False),
            ("DeviceHardware", "Storage", "Storage", False),
            ("DeviceHardware", "Hardware detail rows", "Detail rows", True),
        ],
        "Hardware equipment", 648, 344, 608, 176,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "device360",
        [
            ("DeviceSource", "SourceSystem", "Source", False),
            ("DeviceSource", "ActivityKind", "Signal", False),
            ("DeviceSource", "ActivityRaw", "Original date text", False),
            ("DeviceSource", "ActivityStatus", "Date qualification", False),
            ("DeviceSource", "ActivityUtcDateTime", "Qualified UTC", False),
            ("DeviceSource", "SourceCollectedDateTime", "Collected UTC", False),
            ("DeviceSource", "Device source details", "Detail rows", True),
        ],
        "Activity evidence — ambiguous dates remain unqualified", 24, 532, 608, 176,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "device360",
        [
            ("DeviceHardware", "ManagedDeviceId", "Source record", False),
            ("DeviceHardware", "HardwareCollectedDateTime", "Hardware collected UTC", False),
            ("DeviceHardware", "InventoryCollectedDateTime", "Inventory collected UTC", False),
            ("DeviceHardware", "CollectionCoverage", "Coverage", False),
            ("DeviceHardware", "CollectionMode", "Mode", False),
            ("DeviceHardware", "Hardware detail rows", "Detail rows", True),
        ],
        "Hardware source evidence", 648, 532, 608, 176,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "device360",
        [
            ("EntityFinding", "Severity", "Severity", False),
            ("EntityFinding", "FindingType", "Finding", False),
            ("EntityFinding", "Description", "Evidence", False),
            ("EntityFinding", "RecommendedAction", "Suggested review", False),
            ("EntityFinding", "Device finding details", "Detail rows", True),
        ],
        "Findings linked to this device", 24, 720, 1232, 124,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    return page, visuals


def build_impact(pages):
    page, visuals = new_page(
        pages, "impact", "Impact analysis",
        "Trace the relationships the current CMDB can prove. User-device and license assignment paths are evidence; they are not a complete business dependency graph.",
    )
    add_slicer(pages, visuals, "impact", "FactUserDeviceRelationship", "RelationshipType", "Relationship type", 24)
    add_slicer(pages, visuals, "impact", "LicenseAssignmentPath", "AssignmentRoute", "Assignment route", 648)
    for index, item in enumerate([
        ("FactUserDeviceRelationship", "User-device links", "User / device links"),
        ("FactUserDeviceRelationship", "Devices with a resolved user", "Devices with a user"),
        ("FactUserDeviceRelationship", "Users with a resolved device", "Users with a device"),
        ("FactUserLicense", "Users with assignments", "Users with assignments"),
    ]):
        add_card(pages, visuals, "impact", *item, 24 + index * 312, 232)
    add_bar(pages, visuals, "impact", "FactUserDeviceRelationship", "RelationshipType", "FactUserDeviceRelationship", "User-device links", "User-device links by relationship type", 24, 340)
    add_bar(pages, visuals, "impact", "LicenseAssignmentPath", "AssignmentRoute", "LicenseAssignmentPath", "Assignment paths", "License assignment paths by route", 648, 340)
    add_table(
        pages, visuals, "impact",
        [
            ("LicenseAssignmentPath", "Account", "Account", False),
            ("LicenseAssignmentPath", "Product", "Product", False),
            ("LicenseAssignmentPath", "AssignmentRoute", "Route", False),
            ("LicenseAssignmentPath", "GroupName", "Group", False),
            ("LicenseAssignmentPath", "ErrorStatus", "Error status", False),
        ],
        "Observed assignment paths — use User 360 and Group 360 for focused evidence", 24, 576, 1232, 250,
    )
    return page, visuals


def update_overview(pages: Path, mailbox_metadata):
    page = copy.deepcopy(load(pages / "overview" / "page.json"))
    page.update(name="overview", displayName=DISPLAY_NAMES["overview"])
    page.pop("filterConfig", None)
    visuals = []

    for template, value, x, y, w, h in [
        ("devicesv0", "Smart Workplace CMDB — Executive Overview", 24, 14, 680, 42),
        ("devicesv1", "VERSION 1.0.0", 24, 60, 680, 24),
        ("devicesv2", "Workforce, workplace devices, messaging and Microsoft 365 licensing — one shared country context.", 24, 88, 680, 34),
        ("devicesv3", "V1 · Frozen private snapshot · Selections filter related visuals", 24, 864, 1232, 28),
    ]:
        visual = clone_visual(pages, "devices", template)
        set_text(visual, value)
        put(visuals, visual, "overview", x, y, w, h)

    corporate_share = add_card(
        pages, visuals, "overview", "DimCountry", "Executive corporate country share",
        "Corporate country rate", 720, 14, 216, 80,
    )
    corporate_share["visual"]["objects"]["value"][0]["properties"]["labelPrecision"] = lit(1)

    ownership = clone_visual(pages, "devices", "devicesv4")
    clear_filter(ownership)
    ownership["visual"]["query"] = {
        "queryState": {"Values": {"projections": [projection("DimDevice", "Ownership", "Device ownership")]}}
    }
    set_title(ownership, "Device ownership")
    ownership = put(visuals, ownership, "overview", 944, 14, 152, 80)
    # The slicer projection resolves the source field "Ownership" to the
    # modeled display column "OwnershipLabel".  The persisted selection must
    # target that same modeled column or Desktop renders the data as filtered
    # while the dropdown misleadingly displays "All".
    set_categorical_selection(ownership, "DimDevice", "OwnershipLabel", "Corporate")

    country = clone_visual(pages, "devices", "devicesv4")
    clear_filter(country)
    country["visual"]["query"] = {
        "queryState": {"Values": {"projections": [projection("DimCountry", "CountryLabel", "Country")]}}
    }
    set_title(country, "Country")
    exclude_synthetic_blank(country, "DimCountry", "CountryLabel")
    country = put(visuals, country, "overview", 1104, 14, 152, 80)

    kpis = [
        ("DimCountry", "Executive workplace devices", "Workplace devices"),
        ("DimCountry", "Executive managed device share", "Managed device rate"),
        ("DimCountry", "Executive compliant device share", "Device compliance rate"),
        ("DimUser", "Users", "Users"),
        ("FactMailboxHosting", "Hosted mailboxes", "Mailboxes"),
        ("FactDataQuality", "Warnings", "Data-quality warnings"),
    ]
    warning = None
    for index, item in enumerate(kpis):
        card = add_card(pages, visuals, "overview", *item, 24 + index * 208, 136, 192, 96)
        if item[1] == "Warnings":
            warning = card
            props = card["visual"]["objects"]["value"][0]["properties"]
            props["fontColor"] = {"solid": {"color": lit("#B45309")}}
            set_page_navigation(card, "risk", "Open Workplace Health filtered to Warning findings")

    form_factor = add_donut(pages, visuals, "overview", "DimDevice", "Device form factor", "DimDevice", "Devices", "PC vs mobile devices", 24, 240)
    ownership_donut = add_donut(pages, visuals, "overview", "DimDevice", "Device ownership group", "DimDevice", "Devices", "Corporate vs personal", 272, 240)
    windows = add_donut(pages, visuals, "overview", "DimDevice", "Windows version group", "DimDevice", "Devices", "Windows 11 adoption", 520, 240)
    set_single_value_filter(windows, "DimDevice", "OperatingSystem", "Windows", "windows")
    accounts = add_donut(pages, visuals, "overview", "DimUser", "AccountStatusLabel", "DimUser", "Users", "Enabled vs disabled users", 768, 240)
    mailboxes = add_donut(pages, visuals, "overview", "FactMailboxHosting", "HostingLocation", "FactMailboxHosting", "Hosted mailboxes", "Exchange Online vs on-premises", 1016, 240)

    country_bar = add_country_bar(pages, visuals, "overview", 24, 448, 1232, 280)

    for index, (title, sku) in enumerate(LICENSE_SUMMARY_SKUS):
        add_license_summary_card(pages, visuals, "overview", title, sku, 24 + index * 248, 736)

    page["visualInteractions"] = [
        {"source": ownership["name"], "target": ownership_donut["name"], "type": "NoFilter"},
        {"source": country["name"], "target": country_bar["name"], "type": "NoFilter"},
    ]
    persist_page(pages, page, visuals)


def persist_page(pages: Path, page, visuals):
    page_id = page["name"]
    target = pages / page_id
    if target.exists():
        if target.parent != pages or page_id not in MANAGED_PAGES:
            raise ValueError(f"Refusing to replace unmanaged page: {target}")
        shutil.rmtree(target)
    write(target / "page.json", page)
    for visual in visuals:
        write(target / "visuals" / visual["name"] / "visual.json", visual)


def prepare(report: Path, exchange_onprem_local: Path | None = None, exchange_onprem_remote: Path | None = None):
    report = report.resolve()
    pages = report / "definition" / "pages"
    if not (pages / "pages.json").is_file():
        raise ValueError("Expected a PBIR report definition")
    required = {"overview", "devices", "users", "licenses", "device360", "user360", "group360"}
    if not required.issubset({p.name for p in pages.iterdir() if p.is_dir()}):
        raise ValueError("The existing CMDB report is incomplete")

    exchange_onprem_local = exchange_onprem_local.resolve() if exchange_onprem_local else None
    exchange_onprem_remote = exchange_onprem_remote.resolve() if exchange_onprem_remote else None
    for source in (exchange_onprem_local, exchange_onprem_remote):
        if source and not source.is_file():
            raise ValueError(f"Exchange evidence file not found: {source}")
    mailbox_metadata = enrich_semantic_model(report, exchange_onprem_local, exchange_onprem_remote)
    for builder in [build_risk, build_lifecycle, build_fleet_hardware,
                    build_people_messaging, build_business_services, build_device_360]:
        page, visuals = builder(pages)
        persist_page(pages, page, visuals)
    update_overview(pages, mailbox_metadata)

    for page_id in RETIRED_PAGES:
        target = pages / page_id
        if target.exists():
            if target.parent != pages or target.name != page_id:
                raise ValueError(f"Refusing to retire unsafe page path: {target}")
            shutil.rmtree(target)

    for page_id, display_name in DISPLAY_NAMES.items():
        page_path = pages / page_id / "page.json"
        page = load(page_path)
        page["displayName"] = display_name
        page.pop("visibility", None)
        write(page_path, page)
        if page_id in HEADER_NAMES:
            visual_path = pages / page_id / "visuals" / f"{page_id}v0" / "visual.json"
            if visual_path.is_file():
                visual = load(visual_path)
                set_text(visual, f"Smart Workplace CMDB — {HEADER_NAMES[page_id]}")
                write(visual_path, visual)
        footer_path = pages / page_id / "visuals" / f"{page_id}v3" / "visual.json"
        if footer_path.is_file():
            footer = load(footer_path)
            footer_text = (
                "V1 · Detail page · Select one item or use drill-through navigation"
                if page_id in DRILLTHROUGH_PAGES else
                "V1 · Frozen snapshot · Selections filter related visuals"
            )
            set_text(footer, footer_text)
            write(footer_path, footer)

    metadata_path = pages / "pages.json"
    metadata = load(metadata_path)
    metadata["pageOrder"] = PAGE_ORDER
    metadata["activePageName"] = "overview"
    write(metadata_path, metadata)

    report_path = report / "definition" / "report.json"
    report_json = load(report_path)
    report_json.setdefault("settings", {})["pagesPosition"] = "Bottom"
    for annotation in report_json.get("annotations", []):
        if annotation.get("name") == "defaultPage":
            annotation["value"] = "overview"
    write(report_path, report_json)
    return {
        "status": "Prepared", "pages": len(PAGE_ORDER), "visiblePages": len(PAGE_ORDER),
        "hiddenPages": 0, "retiredPages": len(RETIRED_PAGES),
        "activePage": "overview", "mailboxHosting": mailbox_metadata,
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--exchange-onprem-local", type=Path)
    parser.add_argument("--exchange-onprem-remote", type=Path)
    args = parser.parse_args()
    print(json.dumps(prepare(args.report, args.exchange_onprem_local, args.exchange_onprem_remote), ensure_ascii=False))
