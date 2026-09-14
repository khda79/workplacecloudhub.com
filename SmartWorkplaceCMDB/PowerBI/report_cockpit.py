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
import re
import shutil
from pathlib import Path

from build_report import exclude_synthetic_blank, lit, m_query, projection, type_of


MANAGED_PAGES = {
    "overview", "risk", "lifecycle", "licenses", "businessservices", "devices", "users",
    "device360", "user360", "group360",
}
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
    "devices": "05  Fleet, Hardware & Apps",
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
    "devices": "Fleet, Hardware & Apps",
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
    "MailboxTypeGroup",
    "EvidenceSource",
]
UPGRADE_ELIGIBILITY_COLUMNS = [
    *IDENTITY_COLUMNS,
    "TenantUpgradeEligibilityDeviceKey",
    "MetricId",
    "MetricDeviceId",
    "DeviceId",
    "DeviceIdSource",
    "DeviceName",
    "UpgradeEligibility",
    "SourceCollectedDateTime",
    "SourceSystem",
]
UPGRADE_ELIGIBILITY_STATES = {
    "upgraded", "unknown", "notCapable", "capable", "unknownFutureValue",
}
TOP_APPLICATION_COLUMNS = [
    "ApplicationProduct", "DisplayName", "Publisher", "Platform",
    "VersionCount", "ReportedDeviceCount",
]
DIM_DETECTED_APPLICATION_COLUMNS = [
    *IDENTITY_COLUMNS, "TenantApplicationKey", "AppId", "SourceApplicationKey",
    "DisplayName", "Version", "Publisher", "DeviceCount", "ReportedDeviceCount",
    "ExactRelatedDeviceCount", "RelationshipCoverageStatus", "Platform",
    "SourceCollectedDateTime",
]
RELATIONSHIP_OVERVIEW_COLUMNS = [
    "RelationshipType", "RelationshipCount", "EvidenceSource",
]
FACT_DEVICE_APPLICATION_COLUMNS = [
    *IDENTITY_COLUMNS, "TenantDeviceApplicationKey", "TenantApplicationKey",
    "ManagedDeviceId", "AppId", "SourceCollectedDateTime",
]
INTUNE_DEVICE_BRIDGE_COLUMNS = [
    *IDENTITY_COLUMNS, "TenantIntuneDeviceKey", "TenantDeviceKey", "ManagedDeviceId",
]
FACT_AD_INTUNE_COVERAGE_COLUMNS = [
    *IDENTITY_COLUMNS, "TenantADComputerKey", "CmdbAdComputerId", "DeviceName",
    "Enabled", "OperatingSystem", "OperatingSystemVersion", "TenantDeviceKey",
    "EntraDeviceId", "IntuneManagedDeviceId", "CoverageState", "MatchMethod",
    "SourceCollectedDateTime",
]
DIM_SHAREPOINT_SITE_COLUMNS = [
    *IDENTITY_COLUMNS, "TenantSiteKey", "SiteId", "SiteUrl", "SiteName",
    "OwnerPrincipalName", "LastActivityDate", "ActivityState", "StorageUsedBytes",
    "StorageAllocatedBytes", "RootWebTemplate", "IsDeleted", "SourceCollectedDateTime",
]
DIM_TEAM_COLUMNS = [
    *IDENTITY_COLUMNS, "TenantTeamKey", "TeamId", "DisplayName", "Visibility",
    "CreatedDateTime", "LastActivityDate", "ActivityState", "OwnerCount",
    "MemberCount", "GuestCount", "UnresolvedMemberCount", "MembershipCoverageStatus",
    "IsArchived", "SourceCollectedDateTime",
]
FACT_TEAM_MEMBER_COLUMNS = [
    *IDENTITY_COLUMNS, "TenantTeamMemberKey", "TenantTeamKey", "TeamId",
    "TenantUserKey", "UserId", "UserPrincipalName", "UserType", "Role",
    "SourceCollectedDateTime",
]
FACT_USER_ACTIVITY_COLUMNS = [
    *IDENTITY_COLUMNS, "TenantUserActivityKey", "TenantUserKey", "UserPrincipalName",
    "MatchStatus", "ReportRefreshDate", "IsDeleted", "ExchangeLastActivityDate",
    "OneDriveLastActivityDate", "SharePointLastActivityDate", "TeamsLastActivityDate",
    "LastActivityDate", "LastActivityWorkload", "HasAnyM365Activity",
    "AssignedProducts", "SourceCollectedDateTime",
]
LICENSE_SUMMARY_SKUS = [
    ("Microsoft 365 F1", "M365_F1"),
    ("Microsoft 365 F3", "SPE_F1"),
    ("Microsoft 365 E3", "SPE_E3"),
    ("Microsoft 365 E5", "SPE_E5"),
    ("Microsoft 365 Copilot", "Microsoft_365_Copilot"),
]
LICENSE_COUNTRY_SERIES = [
    ("Microsoft 365 F1", "M365_F1", "Executive M365 F1 country share", "F1", "#7567A8"),
    ("Microsoft 365 F3", "SPE_F1", "Executive M365 F3 country share", "F3", "#2A8C8C"),
    ("Microsoft 365 E3", "SPE_E3", "Executive M365 E3 country share", "E3", "#4D7FB8"),
    ("Microsoft 365 E5", "SPE_E5", "Executive M365 E5 country share", "E5", "#8C6FB1"),
    ("Microsoft 365 Copilot", "Microsoft_365_Copilot", "Executive Copilot country share", "Copilot", "#A6824A"),
]
QUALITY_INDICATORS = [
    (
        "Integrity issues",
        "CALCULATE([Quality findings], KEEPFILTERS('FactDataQuality'[FindingType] == \"OrphanPrimaryUserReference\")) + "
        "CALCULATE([Quality findings], KEEPFILTERS('FactDataQuality'[FindingType] == \"ObservedLicenseAssignmentError\"))",
        "Orphan primary-user references plus observed Microsoft 365 license-assignment errors.",
    ),
    (
        "Coverage gaps",
        "CALCULATE([Quality findings], KEEPFILTERS('FactDataQuality'[FindingType] == \"UserCountryUnknown\")) + "
        "CALCULATE([Quality findings], KEEPFILTERS('FactDataQuality'[FindingType] == \"DeviceWithoutPrimaryUser\"))",
        "User records without a country plus devices without a primary user; these are coverage gaps, not integrity failures.",
    ),
    (
        "Derived country gaps",
        "CALCULATE([Quality findings], KEEPFILTERS('FactDataQuality'[FindingType] == \"DeviceCountryUnknown\"))",
        "Devices whose country cannot be derived. This overlaps device primary-user findings and must not be added to Coverage gaps.",
    ),
]
POPULATION_COUNTRY_SERIES = [
    ("DimCountry", "Executive corporate country share", "Corporate devices", "#2A8C8C"),
    ("DimCountry", "Executive user country share", "Users", "#46637F"),
    ("DimCountry", "Executive mailbox country share", "Mailboxes", "#6E8FB3"),
]
EXECUTIVE_RATIO_CATEGORY_COLORS = {
    "form_factor": {
        "Mobile": "#7567A8",
        "PC": "#4D7FB8",
        "Unclassified": "#AAB7C4",
    },
    "ownership": {
        "Corporate": "#2A8C8C",
        "Personal": "#6DB7B7",
        "Unknown / missing": "#AAB7C4",
    },
    "windows": {
        "Windows 11": "#3F7FC4",
        "Other Windows": "#7BAFD4",
        "Unknown version": "#AAB7C4",
    },
    "accounts": {
        "Enabled": "#3F8F62",
        "Disabled": "#C96C6C",
    },
    "mailbox_types": {
        "User mailbox": "#4D7FB8",
        "Shared mailbox": "#2A8C8C",
        "Other mailbox types": "#AAB7C4",
    },
}

OVERVIEW_ICON_RESOURCE = "smartworkplace-overview-20260913.svg"
PAGE_ICON_RESOURCES = {
    "overview": OVERVIEW_ICON_RESOURCE,
    "risk": "smartworkplace-health-20260913.svg",
    "lifecycle": "smartworkplace-lifecycle-20260913.svg",
    "licenses": "smartworkplace-licenses-20260913.svg",
    "devices": "smartworkplace-device-20260913.svg",
    "users": "smartworkplace-people-20260913.svg",
    "businessservices": "smartworkplace-services-20260913.svg",
    "device360": "smartworkplace-device-20260913.svg",
    "user360": "smartworkplace-people-20260913.svg",
    "group360": "smartworkplace-people-20260913.svg",
}
OVERVIEW_RATIO_ICON_RESOURCES = {
    "formfactor": "smartworkplace-ratio-formfactor-20260913.svg",
    "ownership": "smartworkplace-ratio-ownership-20260913.svg",
    "windows": "smartworkplace-ratio-windows-20260913.svg",
    "accounts": "smartworkplace-ratio-accounts-20260913.svg",
    "mailboxes": "smartworkplace-ratio-mailboxes-20260913.svg",
}
KPI_ICON_RESOURCES = {
    "devices": "smartworkplace-kpi-devices-20260913.svg",
    "management": "smartworkplace-kpi-management-20260913.svg",
    "compliance": "smartworkplace-kpi-compliance-20260913.svg",
    "users": "smartworkplace-kpi-users-20260913.svg",
    "mailboxes": "smartworkplace-kpi-mailboxes-20260913.svg",
    "licenses": "smartworkplace-kpi-licenses-20260913.svg",
    "quality": "smartworkplace-kpi-quality-20260913.svg",
    "analytics": "smartworkplace-kpi-analytics-20260913.svg",
    "relationships": "smartworkplace-kpi-relationships-20260913.svg",
}
SECTION_ICON_RESOURCES = {
    "identity": "smartworkplace-section-identity-20260913.svg",
    "enrollment": "smartworkplace-section-enrollment-20260913.svg",
    "hardware": "smartworkplace-section-hardware-20260913.svg",
    "activity": "smartworkplace-section-activity-20260913.svg",
    "findings": "smartworkplace-section-findings-20260913.svg",
}
COPYRIGHT_TEXT = "© 2026 WorkplaceCloudHub — https://workplacecloudhub.com/"
COPYRIGHT_URL = "https://workplacecloudhub.com/"
DEFAULT_OWNERSHIP = "Corporate"
DEFAULT_OWNERSHIP_FILTER_NAME = "defaultdeviceownership"
COUNTRY_FOOTPRINT_COLUMN = "Country footprint label"
COUNTRY_FOOTPRINT_UNKNOWN = "Unknown"
ENDPOINT_ANALYTICS_CARD_TITLE = "Endpoint Analytics Score"
_VISUAL_TEMPLATE_CACHE = {}

BAR_CATEGORY_COLORS = {
    ("DimDevice", "ComplianceStateLabel"): {
        "Compliant": "#3F8F62",
        "NonCompliant": "#C96C6C",
        "InGracePeriod": "#A6824A",
        "Not provided": "#AAB7C4",
        "Unknown": "#7F8C9A",
    },
    ("DimDevice", "ManagementStateLabel"): {
        "Managed": "#2A8C8C",
        "Unmanaged": "#C96C6C",
        "Not provided": "#AAB7C4",
    },
    ("DimDevice", "OperatingSystemLabel"): {
        "Windows": "#4D7FB8",
        "Android": "#2A8C8C",
        "iOS": "#7567A8",
        "Unknown": "#AAB7C4",
        "Not provided": "#AAB7C4",
    },
    ("FactWindowsUpdateAlert", "AggregateState"): {
        "Success": "#3F8F62",
        "In progress": "#4D7FB8",
        "Error": "#C96C6C",
        "Cancelled": "#AAB7C4",
        "Rollback initiated": "#A6824A",
    },
    ("FactAutopilotDevice", "EnrollmentState"): {
        "enrolled": "#2A8C8C",
        "notContacted": "#AAB7C4",
        "failed": "#C96C6C",
    },
    ("DimUser", "ActivityState"): {
        "Active": "#3F8F62",
        "Inactive": "#C96C6C",
        "Unknown": "#AAB7C4",
    },
    ("DimUser", "AccountStatusLabel"): EXECUTIVE_RATIO_CATEGORY_COLORS["accounts"],
    ("FactUserLicense", "AssignmentStateLabel"): {
        "Active": "#3F8F62",
        "Error": "#C96C6C",
        "Unknown": "#AAB7C4",
    },
    ("LicenseAssignmentPath", "AssignmentRoute"): {
        "Direct": "#4D7FB8",
        "Group": "#7567A8",
        "Unknown": "#AAB7C4",
    },
    ("DimLicenseServicePlan", "ProvisioningStatus"): {
        "Success": "#3F8F62",
        "PendingInput": "#A6824A",
        "PendingActivation": "#4D7FB8",
        "PendingProvisioning": "#4D7FB8",
        "Disabled": "#AAB7C4",
        "Error": "#C96C6C",
    },
    ("FactEndpointAnalyticsUpgradeEligibility", "UpgradeEligibility"): {
        "capable": "#3F8F62",
        "notCapable": "#C96C6C",
        "upgraded": "#4D7FB8",
        "unknown": "#AAB7C4",
        "unknownFutureValue": "#7F8C9A",
    },
    ("FactMailboxHosting", "HostingLocation"): {
        "Exchange Online": "#4D7FB8",
        "Exchange On-premises": "#A6824A",
    },
    ("FactMailboxHosting", "MailboxTypeGroup"): EXECUTIVE_RATIO_CATEGORY_COLORS["mailbox_types"],
    ("FactRelationshipOverview", "RelationshipType"): {
        "PrimaryUser": "#4D7FB8",
        "HasMailbox": "#2A8C8C",
        "AssignedLicense": "#7567A8",
        "MemberOfGroup": "#2A8C8C",
        "DeviceHasApplication": "#A6824A",
        "DeviceInAutopilot": "#3F8F62",
    },
    ("DimSharePointSite", "ActivityState"): {
        "Active (90d)": "#3F8F62", "Inactive (>90d)": "#C96C6C", "Unknown": "#AAB7C4",
    },
    ("DimTeam", "ActivityState"): {
        "Active (90d)": "#3F8F62", "Inactive (>90d)": "#C96C6C", "Unknown": "#AAB7C4",
    },
    ("FactTeamMember", "Role"): {"Owner": "#4D7FB8", "Member": "#2A8C8C"},
}


def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write(path: Path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False), encoding="utf-8")


def clone_visual(pages: Path, page: str, visual: str):
    key = (str(pages.resolve()), page, visual)
    if key not in _VISUAL_TEMPLATE_CACHE:
        _VISUAL_TEMPLATE_CACHE[key] = load(pages / page / "visuals" / visual / "visual.json")
    return copy.deepcopy(_VISUAL_TEMPLATE_CACHE[key])


def clone_visual_by_type(pages: Path, page: str, candidates, visual_types):
    for candidate in candidates:
        path = pages / page / "visuals" / candidate / "visual.json"
        if not path.is_file():
            continue
        visual = load(path)
        if visual.get("visual", {}).get("visualType") in set(visual_types):
            return copy.deepcopy(visual)
    raise ValueError(f"No {visual_types} template found in {page}: {candidates}")


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


def count_csv_rows(path: Path, predicate=lambda _row: True):
    """Count large fact rows without materializing the source in memory."""
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if not reader.fieldnames:
            raise ValueError(f"Malformed CSV: {path}")
        count = 0
        for row in reader:
            if None in row:
                raise ValueError(f"Malformed CSV: {path}")
            if predicate(row):
                count += 1
        return count


def build_intune_device_bridge(rows, identity):
    """Build one exact Intune managed-device key per report device."""
    result = []
    seen = set()
    for row in rows:
        if (row.get("SourceSystem") or "").strip() != "Intune":
            continue
        if any((row.get(column) or "") != identity[column] for column in IDENTITY_COLUMNS):
            raise ValueError("Intune device bridge tenant identity mismatch")
        managed_device_id = (row.get("SourceObjectId") or "").strip()
        tenant_device_key = (row.get("TenantDeviceKey") or "").strip()
        if not managed_device_id or not tenant_device_key:
            raise ValueError("Intune device bridge requires exact managed-device and CMDB device keys")
        tenant_intune_device_key = (
            f"{identity['TenantKey']}|intune-device|{managed_device_id.lower()}"
        )
        folded = tenant_intune_device_key.casefold()
        if folded in seen:
            raise ValueError(f"Duplicate Intune device bridge key: {tenant_intune_device_key}")
        seen.add(folded)
        result.append({
            **identity,
            "TenantIntuneDeviceKey": tenant_intune_device_key,
            "TenantDeviceKey": tenant_device_key,
            "ManagedDeviceId": managed_device_id,
        })
    if not result:
        raise ValueError("No exact Intune device bridge rows were found")
    return result


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


def mailbox_type_group(row):
    """Normalize local and remote Exchange recipient types into executive families."""
    value = recipient_type(row).casefold()
    if value in {"usermailbox", "remoteusermailbox"}:
        return "User mailbox"
    if value in {"sharedmailbox", "remotesharedmailbox"}:
        return "Shared mailbox"
    return "Other mailbox types"


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


def set_page_categorical_filter(page, name, table, column, value):
    """Add or replace one unlocked page filter with an explicit default value."""
    field = projection(table, column)["field"]
    actual_column = field["Column"]["Property"]
    item = {
        "name": name,
        "field": field,
        "type": "Categorical",
        "howCreated": "User",
        "filter": {
            "Version": 2,
            "From": [{"Name": "s", "Entity": table, "Type": 0}],
            "Where": [{"Condition": {"In": {
                "Expressions": [{"Column": {
                    "Expression": {"SourceRef": {"Source": "s"}},
                    "Property": actual_column,
                }}],
                "Values": [[{"Literal": {
                    "Value": "'" + value.replace("'", "''") + "'",
                }}]],
            }}}],
        },
    }
    config = page.setdefault("filterConfig", {})
    filters = config.setdefault("filters", [])
    def is_same_ownership_filter(current):
        current_column = current.get("field", {}).get("Column", {})
        current_entity = (
            current_column.get("Expression", {}).get("SourceRef", {}).get("Entity")
        )
        return (
            current.get("name") in {name, DEFAULT_OWNERSHIP_FILTER_NAME}
            or (current_entity == table and current_column.get("Property") == actual_column)
        )

    config["filters"] = [current for current in filters if not is_same_ownership_filter(current)]
    config["filters"].append(item)
    config.setdefault("filterSortOrder", "Custom")


def set_categorical_values_filter(visual, table, column, values, suffix="selection"):
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
                "Values": [[{"Literal": {"Value": "'" + value.replace("'", "''") + "'"}}] for value in values],
            }}}],
        },
    }]}


def set_donut(visual, category_table, category_column, measure_table, measure, title):
    clear_filter(visual)
    visual["visual"]["visualType"] = "donutChart"
    # The chart templates are cloned from the existing fleet page.  Bar-axis
    # objects are invalid on a donut visual and must not survive the type swap.
    objects = visual["visual"].setdefault("objects", {})
    objects.pop("categoryAxis", None)
    objects.pop("valueAxis", None)
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
    visual = clone_visual_by_type(pages, "devices", ["devicesv12", "devicesv11", "devicesv13"], {"barChart", "clusteredBarChart"})
    set_donut(visual, category_table, category_column, measure_table, measure, title)
    return put(visuals, visual, page_id, x, y, w, h)


def set_ratio_bar(visual, category_table, category_column, measure_table, ratio_measure, count_measure, title):
    """Render a compact ratio-first horizontal bar with counts in the tooltip."""
    clear_filter(visual)
    visual["visual"]["visualType"] = "clusteredBarChart"
    visual["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [projection(category_table, category_column)]},
            "Y": {"projections": [projection(measure_table, ratio_measure, ratio_measure, True)]},
            "Tooltips": {"projections": [projection(measure_table, count_measure, count_measure, True)]},
        },
        "sortDefinition": {
            "sort": [{
                "field": projection(measure_table, ratio_measure, ratio_measure, True)["field"],
                "direction": "Descending",
            }],
            "isDefaultSort": True,
        },
    }
    set_title(visual, title)
    exclude_synthetic_blank(visual, category_table, category_column)
    objects = visual["visual"].setdefault("objects", {})
    objects["labels"] = [{"properties": {
        "show": lit(True),
        "labelPosition": lit("OutsideEnd"),
        "fontSize": lit(10),
        "bold": lit(True),
        "color": {"solid": {"color": lit("#25364A")}},
        "labelDisplayUnits": lit(1),
        "labelPrecision": lit(1),
    }}]
    objects["categoryAxis"] = [{"properties": {
        "show": lit(True),
        "fontSize": lit(9),
        "maxMarginFactor": lit(42),
        "innerPadding": lit(28),
        "showAxisTitle": lit(False),
        "gridlineShow": lit(False),
    }}]
    objects["valueAxis"] = [{"properties": {
        "show": lit(False),
        "start": lit(0),
        "end": lit(1),
        "showAxisTitle": lit(False),
        "gridlineShow": lit(False),
    }}]
    objects["legend"] = [{"properties": {"show": lit(False)}}]
    category_colors = BAR_CATEGORY_COLORS.get((category_table, category_column))
    if category_colors:
        set_category_colors(visual, category_table, category_column, category_colors)


def add_ratio_bar(pages, visuals, page_id, category_table, category_column, measure_table,
                  ratio_measure, count_measure, title, x, y, w=240, h=200):
    visual = clone_visual_by_type(pages, "devices", ["devicesv12", "devicesv11", "devicesv13"], {"barChart", "clusteredBarChart"})
    set_ratio_bar(
        visual, category_table, category_column, measure_table,
        ratio_measure, count_measure, title,
    )
    return put(visuals, visual, page_id, x, y, w, h)


def category_scope_selector(table, column, value):
    return {
        "data": [{
            "scopeId": {
                "Comparison": {
                    "ComparisonKind": 0,
                    "Left": {
                        "Column": {
                            "Expression": {"SourceRef": {"Entity": table}},
                            "Property": column,
                        }
                    },
                    "Right": {"Literal": {"Value": f"'{value.replace(chr(39), chr(39) * 2)}'"}},
                }
            }
        }]
    }


def set_category_colors(visual, table, column, value_colors, fallback="#94A3B8"):
    entries = [{
        "properties": {"defaultColor": {"solid": {"color": lit(fallback)}}}
    }]
    entries.extend(
        {
            "properties": {"fill": {"solid": {"color": lit(hex_color)}}},
            "selector": category_scope_selector(table, column, value),
        }
        for value, hex_color in value_colors.items()
    )
    visual["visual"].setdefault("objects", {})["dataPoint"] = entries


def set_series_colors(visual, series):
    visual["visual"].setdefault("objects", {})["dataPoint"] = [
        {
            "properties": {"fill": {"solid": {"color": lit(hex_color)}}},
            "selector": {"metadata": f"{table}.{measure}"},
        }
        for table, measure, _label, hex_color in series
    ]


def add_country_bar(pages, visuals, page_id, x, y, w, h):
    visual = clone_visual_by_type(pages, "devices", ["devicesv12", "devicesv11", "devicesv13"], {"barChart", "clusteredBarChart"})
    clear_filter(visual)
    visual["visual"]["visualType"] = "clusteredBarChart"
    measures = [(table, measure, label) for table, measure, label, _color in POPULATION_COUNTRY_SERIES]
    visual["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [projection("DimCountry", COUNTRY_FOOTPRINT_COLUMN, "Country")]},
            "Y": {"projections": [projection(table, measure, label, True) for table, measure, label in measures]},
        },
        "sortDefinition": {
            "sort": [{"field": projection("DimCountry", "Executive user country share", "Users", True)["field"], "direction": "Descending"}],
            "isDefaultSort": True,
        },
    }
    set_title(visual, "Country footprint — population distribution")
    set_series_colors(visual, POPULATION_COUNTRY_SERIES)
    exclude_synthetic_blank(visual, "DimCountry", COUNTRY_FOOTPRINT_COLUMN)
    for entry in visual["visual"].get("objects", {}).get("labels", []):
        entry.get("properties", {})["labelPrecision"] = lit(1)
    return put(visuals, visual, page_id, x, y, w, h)


def add_license_country_bar(pages, visuals, page_id, x, y, w, h):
    visual = clone_visual_by_type(pages, "devices", ["devicesv12", "devicesv11", "devicesv13"], {"barChart", "clusteredBarChart"})
    clear_filter(visual)
    visual["visual"]["visualType"] = "clusteredBarChart"
    measures = [
        ("DimCountry", measure, label, color)
        for _title, _sku, measure, label, color in LICENSE_COUNTRY_SERIES
    ]
    visual["visual"]["query"] = {
        "queryState": {
            "Category": {"projections": [projection("DimCountry", COUNTRY_FOOTPRINT_COLUMN, "Country")]},
            "Y": {"projections": [projection(table, measure, label, True) for table, measure, label, _color in measures]},
        },
        "sortDefinition": {
            "sort": [{"field": projection("DimCountry", "Executive M365 E3 country share", "E3", True)["field"], "direction": "Descending"}],
            "isDefaultSort": True,
        },
    }
    set_title(visual, "Country footprint — Microsoft 365 licenses")
    set_series_colors(visual, measures)
    exclude_synthetic_blank(visual, "DimCountry", COUNTRY_FOOTPRINT_COLUMN)
    for entry in visual["visual"].get("objects", {}).get("labels", []):
        entry.get("properties", {})["labelPrecision"] = lit(1)
    return put(visuals, visual, page_id, x, y, w, h)


def add_quality_summary_card(pages, visuals, page_id, x, y, w, h):
    visual = clone_visual_by_type(pages, "devices", ["devicesv8", "devicesv7", "devicesv9"], {"cardVisual"})
    clear_filter(visual)
    visual["visual"]["query"] = {
        "queryState": {"Data": {"projections": [
            projection("FactDataQuality", "Data quality health", "Quality status", True)
        ]}}
    }
    value_props = visual["visual"]["objects"]["value"][0]["properties"]
    value_props["fontSize"] = lit(16)
    value_props["bold"] = lit(True)
    value_props["textWrap"] = lit(False)
    value_props["labelDisplayUnits"] = lit(1)
    value_props["fontColor"] = {"solid": {"color": {
        "expr": {"Conditional": {
            "Cases": [
                {
                    "Condition": {"Comparison": {
                        "ComparisonKind": 2,
                        "Left": projection("FactDataQuality", "Data quality score", measure=True)["field"],
                        "Right": {"Literal": {"Value": "0.9D"}},
                    }},
                    "Value": {"Literal": {"Value": "'#16803A'"}},
                },
                {
                    "Condition": {"Comparison": {
                        "ComparisonKind": 3,
                        "Left": projection("FactDataQuality", "Data quality score", measure=True)["field"],
                        "Right": {"Literal": {"Value": "0.6D"}},
                    }},
                    "Value": {"Literal": {"Value": "'#C7352A'"}},
                },
            ],
            "DefaultValue": {"Literal": {"Value": "'#9A6700'"}},
        }}
    }}}
    label_props = visual["visual"]["objects"]["label"][0]["properties"]
    label_props["show"] = lit(True)
    label_props["fontSize"] = lit(9)
    label_props["textWrap"] = lit(False)
    label_props["text"] = lit("Open quality details")
    set_title(visual, "Data quality")
    set_page_navigation(visual, "risk", "Open Workplace Health for finding details")
    return put(visuals, visual, page_id, x, y, w, h)


def add_compact_card(pages, visuals, page_id, table, measure, title, x, y, w, h,
                     show_blank_as="Not collected"):
    """Render a compact supporting metric beneath a primary visual."""
    visual = clone_visual_by_type(pages, "devices", ["devicesv8", "devicesv7", "devicesv9"], {"cardVisual"})
    set_card(visual, table, measure, title)
    value_props = visual["visual"]["objects"]["value"][0]["properties"]
    value_props["fontSize"] = lit(10)
    value_props["showBlankAs"] = lit(show_blank_as)
    title_props = visual["visual"]["visualContainerObjects"]["title"][0]["properties"]
    title_props["fontSize"] = lit(8)
    visual["visual"]["visualContainerObjects"]["padding"] = [{"properties": {
        "top": lit(2), "bottom": lit(2), "left": lit(6), "right": lit(6),
    }}]
    return put(visuals, visual, page_id, x, y, w, h)


def add_license_summary_card(pages, visuals, page_id, title, sku, x, y, w=240, h=104):
    visual = clone_visual_by_type(pages, "devices", ["devicesv8", "devicesv7", "devicesv9"], {"cardVisual"})
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
            "MailboxTypeGroup": mailbox_type_group(row),
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


def mailbox_hosting_rows(report: Path, local_path: Path | None, remote_path: Path | None, data_dir: Path | None = None):
    data_dir = data_dir or (report.parent / "ReportData")
    _, tenants = read_csv(data_dir / "DimTenant.csv")
    if len(tenants) != 1 or any(not tenants[0].get(column) for column in IDENTITY_COLUMNS):
        raise ValueError("Exactly one complete tenant identity is required")
    identity = {column: tenants[0][column] for column in IDENTITY_COLUMNS}
    existing_hosting_path = data_dir / "FactMailboxHosting.csv"
    if not local_path and not remote_path and existing_hosting_path.is_file():
        columns, rows = read_csv(existing_hosting_path)
        if columns != MAILBOX_HOSTING_COLUMNS:
            raise ValueError("Existing mailbox-hosting CSV schema mismatch")
        metadata = {
            "online": sum(row["HostingLocation"] == "Exchange Online" for row in rows),
            "onPremises": sum(row["HostingLocation"] == "Exchange On-premises" for row in rows),
            "total": len(rows),
            "localEvidenceDate": None,
            "remoteEvidenceDate": None,
        }
        return identity, rows, metadata
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


def add_or_replace_calculated_column(table, name, expression, description, data_type="string"):
    columns = table.setdefault("columns", [])
    columns[:] = [column for column in columns if column.get("name") != name]
    columns.append({
        "name": name,
        "dataType": data_type,
        "type": "calculated",
        "expression": expression,
        "description": description,
        "summarizeBy": "none",
    })


def add_operational_measures(tables):
    by_name = {}
    for table in tables:
        by_name.setdefault(table.get("name"), table)

    device = by_name["DimDevice"]
    for name, column, description in [
        ("Compliance state rate", "ComplianceStateLabel", "Share of each compliance state in the filtered device population."),
        ("Operating system rate", "OperatingSystemLabel", "Share of each operating-system family in the filtered device population."),
        ("Management state rate", "ManagementStateLabel", "Share of each management state in the filtered device population."),
        ("Ownership rate", "OwnershipLabel", "Share of each ownership state in the filtered device population."),
    ]:
        add_or_replace_measure(device, name, f"DIVIDE([Devices], CALCULATE([Devices], REMOVEFILTERS('DimDevice'[{column}])))", description, "0.0%")

    endpoint = by_name.get("FactEndpointAnalyticsDevice")
    if endpoint:
        id_scope = (
            "VAR _deviceKeys = VALUES('DimDevice'[TenantDeviceKey]) "
            "VAR _intuneIds = CALCULATETABLE(VALUES('DeviceSource'[SourceObjectId]), "
            "TREATAS(_deviceKeys, 'DeviceSource'[TenantDeviceKey]), "
            "KEEPFILTERS('DeviceSource'[SourceSystem] = \"Intune\")) "
        )
        for name, column, description in [
            ("Selected Endpoint Analytics score", "EndpointAnalyticsScore", "Average Endpoint Analytics score for exact Intune device identifiers in the filtered device population."),
            ("Selected startup performance score", "StartupPerformanceScore", "Average startup score for exact Intune device identifiers in the filtered device population."),
            ("Selected app reliability score", "AppReliabilityScore", "Average app reliability score for exact Intune device identifiers in the filtered device population."),
            ("Selected work from anywhere score", "WorkFromAnywhereScore", "Average work-from-anywhere score for exact Intune device identifiers in the filtered device population."),
        ]:
            add_or_replace_measure(endpoint, name, id_scope + f"RETURN CALCULATE(AVERAGE('FactEndpointAnalyticsDevice'[{column}]), TREATAS(_intuneIds, 'FactEndpointAnalyticsDevice'[DeviceId]))", description, "0.0")

    upgrade = by_name.get("FactEndpointAnalyticsUpgradeEligibility")
    if upgrade:
        add_or_replace_measure(
            upgrade, "Selected upgrade eligibility devices",
            "VAR _deviceKeys = VALUES('DimDevice'[TenantDeviceKey]) VAR _intuneIds = CALCULATETABLE(VALUES('DeviceSource'[SourceObjectId]), TREATAS(_deviceKeys, 'DeviceSource'[TenantDeviceKey]), KEEPFILTERS('DeviceSource'[SourceSystem] = \"Intune\")) RETURN CALCULATE(DISTINCTCOUNT('FactEndpointAnalyticsUpgradeEligibility'[DeviceId]), TREATAS(_intuneIds, 'FactEndpointAnalyticsUpgradeEligibility'[DeviceId]))",
            "Distinct readiness records linked by exact Intune device identifier to the filtered device population.", "#,0",
        )
        add_or_replace_measure(upgrade, "Upgrade eligibility rate", "DIVIDE([Selected upgrade eligibility devices], CALCULATE([Selected upgrade eligibility devices], REMOVEFILTERS('FactEndpointAnalyticsUpgradeEligibility'[UpgradeEligibility])))", "Share of each reported Windows upgrade-eligibility state within exact matched Intune devices.", "0.0%")

    autopilot = by_name.get("FactAutopilotDevice")
    if autopilot:
        add_or_replace_measure(autopilot, "Selected Autopilot devices", "VAR _deviceKeys = VALUES('DimDevice'[TenantDeviceKey]) VAR _intuneIds = CALCULATETABLE(VALUES('DeviceSource'[SourceObjectId]), TREATAS(_deviceKeys, 'DeviceSource'[TenantDeviceKey]), KEEPFILTERS('DeviceSource'[SourceSystem] = \"Intune\")) RETURN CALCULATE(DISTINCTCOUNT('FactAutopilotDevice'[ManagedDeviceId]), TREATAS(_intuneIds, 'FactAutopilotDevice'[ManagedDeviceId]))", "Autopilot records linked by exact Intune managed-device identifier to the filtered device population.", "#,0")
        add_or_replace_measure(autopilot, "Autopilot enrollment rate", "DIVIDE([Selected Autopilot devices], CALCULATE([Selected Autopilot devices], REMOVEFILTERS('FactAutopilotDevice'[EnrollmentState])))", "Share of each Autopilot enrollment state within exact matched Intune devices.", "0.0%")

    apps = by_name.get("DimDetectedApplication")
    if apps:
        add_or_replace_calculated_column(apps, "Application product", "COALESCE('DimDetectedApplication'[DisplayName], \"Unknown application\") & \" · \" & COALESCE('DimDetectedApplication'[Platform], \"Unknown platform\")", "Readable product label used to group version and publisher-metadata variants in visuals; exact application-device identity remains AppId.")
        add_or_replace_measure(apps, "Application products", "DISTINCTCOUNT('DimDetectedApplication'[SourceApplicationKey])", "Distinct source application product identifiers. Versions and publisher-metadata variants sharing the same exact Intune ApplicationKey are counted once.", "#,0")
        add_or_replace_measure(apps, "Application version rows", "COUNTROWS('DimDetectedApplication')", "Source-reported application product-version rows.", "#,0")
        add_or_replace_measure(apps, "Reported application-device occurrences", "SUM('DimDetectedApplication'[ReportedDeviceCount])", "Source-reported application-device occurrences. Exact device-level reporting uses FactDeviceApplication instead.", "#,0")
        add_or_replace_measure(apps, "Application occurrence rate", "DIVIDE([Reported application-device occurrences], CALCULATE([Reported application-device occurrences], REMOVEFILTERS('DimDetectedApplication'[Application product])))", "Share of source-reported application-device occurrences for each product; occurrences are not distinct devices.", "0.0%")

    device_source = by_name.get("DeviceSource")
    device_app = by_name.get("FactDeviceApplication")
    if device_source:
        columns = device_source.setdefault("columns", [])
        columns[:] = [
            column for column in device_source.get("columns", [])
            if column.get("name") != "Tenant Intune device key"
        ]
    if device_source and device_app:
        add_or_replace_calculated_column(device_app, "Tenant Intune device key", "'FactDeviceApplication'[TenantKey] & \"|intune-device|\" & LOWER('FactDeviceApplication'[ManagedDeviceId])", "Exact tenant-scoped Intune managed-device identifier.")
        add_or_replace_measure(device_app, "Installed application products", "VAR _applicationKeys = VALUES('FactDeviceApplication'[TenantApplicationKey]) RETURN CALCULATE(DISTINCTCOUNT('DimDetectedApplication'[SourceApplicationKey]), TREATAS(_applicationKeys, 'DimDetectedApplication'[TenantApplicationKey]))", "Distinct Intune source application products installed on exact matched devices in the current country and ownership context. Versions and publisher-metadata variants are counted once.", "#,0")
        add_or_replace_measure(device_app, "Application-device installations", "COUNTROWS('FactDeviceApplication')", "Exact application-to-managed-device relations in the current filter context.", "#,0")
        add_or_replace_measure(device_app, "Devices reporting applications", "DISTINCTCOUNT('FactDeviceApplication'[ManagedDeviceId])", "Distinct exact Intune managed-device identifiers with the selected application evidence.", "#,0")
        add_or_replace_measure(device_app, "Application device rate", "DIVIDE([Devices reporting applications], CALCULATE([Devices reporting applications], REMOVEFILTERS('DimDetectedApplication'[Application product])))", "Distinct devices with each normalized application product divided by devices with any collected application relation in the same device context.", "0.0%")

    ad_intune = by_name.get("FactADIntuneCoverage")
    if ad_intune:
        add_or_replace_measure(ad_intune, "Enabled AD Windows workstations", "CALCULATE(COUNTROWS('FactADIntuneCoverage'), KEEPFILTERS('FactADIntuneCoverage'[Enabled] = TRUE()))", "Enabled Active Directory Windows 7 through Windows 11 workstation objects. Servers are excluded.", "#,0")
        add_or_replace_measure(ad_intune, "AD Windows workstations managed in Intune", "CALCULATE([Enabled AD Windows workstations], KEEPFILTERS('FactADIntuneCoverage'[CoverageState] = \"Managed in Intune\"))", "Enabled AD workstations linked exactly through AD ObjectSid to Entra onPremisesSecurityIdentifier, then Entra deviceId to Intune azureADDeviceId.", "#,0")
        add_or_replace_measure(ad_intune, "AD to Intune coverage rate", "DIVIDE([AD Windows workstations managed in Intune], [Enabled AD Windows workstations])", "Exact Intune coverage of enabled AD Windows workstations. The measure is tenant-wide because unmatched AD devices do not have a defensible country or ownership attribution.", "0.0%")
        add_or_replace_measure(ad_intune, "AD workstation coverage state rate", "DIVIDE([Enabled AD Windows workstations], CALCULATE([Enabled AD Windows workstations], REMOVEFILTERS('FactADIntuneCoverage'[CoverageState])))", "Share of enabled AD Windows workstations by exact coverage state.", "0.0%")

    user = by_name["DimUser"]
    user_activity = by_name.get("FactUserActivity")
    if user_activity:
        add_or_replace_measure(user_activity, "Users with exact M365 activity evidence", "CALCULATE(DISTINCTCOUNT('FactUserActivity'[TenantUserKey]), KEEPFILTERS('FactUserActivity'[MatchStatus] = \"ExactUPN\"))", "Users whose Microsoft 365 workload report row was linked to DimUser by exact normalized UPN.", "#,0")
        review_expression = "VAR _activityMatch = LOOKUPVALUE('FactUserActivity'[MatchStatus], 'FactUserActivity'[TenantUserKey], 'DimUser'[TenantUserKey]) VAR _workloadLast = LOOKUPVALUE('FactUserActivity'[LastActivityDate], 'FactUserActivity'[TenantUserKey], 'DimUser'[TenantUserKey]) VAR _days = IF(NOT ISBLANK(_workloadLast), DATEDIFF(_workloadLast, TODAY(), DAY)) RETURN SWITCH(TRUE(), 'DimUser'[AccountEnabled] = FALSE(), \"Priority · disabled account\", _activityMatch <> \"ExactUPN\", \"Coverage gap · no exact workload row\", ISBLANK(_workloadLast), \"Priority · no M365 activity in D180\", _days > 90, \"Priority · workload over 90 days\", _days > 30, \"Watch · workload 31–90 days\", \"Recent M365 workload activity\")"
        review_description = "License-review prioritization from exact Office 365 workload activity evidence and account status. It is evidence for review, not proof that a license is unused."
    else:
        review_expression = "\"Coverage gap · M365 activity not collected\""
        review_description = "Microsoft 365 workload activity has not been collected; no license non-use inference is made."
    add_or_replace_calculated_column(user, "License review band", review_expression, review_description)
    add_or_replace_measure(user, "Activity state rate", "DIVIDE([Users], CALCULATE([Users], REMOVEFILTERS('DimUser'[ActivityState])))", "Share of each observed activity state within the filtered user population.", "0.0%")
    add_or_replace_measure(user, "License review candidates", "CALCULATE(DISTINCTCOUNT('LicenseAssignmentPath'[TenantUserKey]), KEEPFILTERS(FILTER('DimUser', LEFT('DimUser'[License review band], 8) = \"Priority\" || LEFT('DimUser'[License review band], 5) = \"Watch\")))", "Distinct licensed users in priority or watch bands based on exact M365 workload evidence; coverage gaps are excluded and non-use is not asserted.", "#,0")

    sites = by_name.get("DimSharePointSite")
    if sites:
        add_or_replace_measure(sites, "SharePoint sites", "DISTINCTCOUNT('DimSharePointSite'[TenantSiteKey])", "Distinct SharePoint sites from the Microsoft 365 D180 site-usage report; OneDrive personal sites are excluded.", "#,0")
        add_or_replace_measure(sites, "Active SharePoint sites", "CALCULATE([SharePoint sites], KEEPFILTERS('DimSharePointSite'[ActivityState] = \"Active (90d)\"))", "SharePoint sites with activity reported within 90 days.", "#,0")
        add_or_replace_measure(sites, "SharePoint active rate", "DIVIDE([Active SharePoint sites], [SharePoint sites])", "SharePoint sites active within 90 days divided by all reported SharePoint sites.", "0.0%")
        add_or_replace_measure(sites, "SharePoint activity state rate", "DIVIDE([SharePoint sites], CALCULATE([SharePoint sites], REMOVEFILTERS('DimSharePointSite'[ActivityState])))", "Share of SharePoint sites by explicit 90-day activity state.", "0.0%")
        add_or_replace_measure(sites, "SharePoint storage used GB", "DIVIDE(SUM('DimSharePointSite'[StorageUsedBytes]), 1073741824)", "Total reported SharePoint site storage in GiB.", "#,0.0")
        add_or_replace_measure(sites, "SharePoint storage utilization", "DIVIDE(SUM('DimSharePointSite'[StorageUsedBytes]), SUM('DimSharePointSite'[StorageAllocatedBytes]))", "Reported SharePoint storage used divided by reported allocated storage.", "0.0%")
    teams = by_name.get("DimTeam")
    team_members = by_name.get("FactTeamMember")
    if teams:
        add_or_replace_measure(teams, "Teams", "DISTINCTCOUNT('DimTeam'[TenantTeamKey])", "Distinct Microsoft 365 groups provisioned as Teams.", "#,0")
        add_or_replace_measure(teams, "Active Teams", "CALCULATE([Teams], KEEPFILTERS('DimTeam'[ActivityState] = \"Active (90d)\"))", "Teams with activity reported within 90 days.", "#,0")
        add_or_replace_measure(teams, "Teams active rate", "DIVIDE([Active Teams], [Teams])", "Teams active within 90 days divided by all Teams.", "0.0%")
        add_or_replace_measure(teams, "Teams activity state rate", "DIVIDE([Teams], CALCULATE([Teams], REMOVEFILTERS('DimTeam'[ActivityState])))", "Share of Teams by explicit 90-day activity state.", "0.0%")
        add_or_replace_measure(teams, "Average members per Team", "CALCULATE(AVERAGE('DimTeam'[MemberCount]), KEEPFILTERS('DimTeam'[MembershipCoverageStatus] = \"Complete\"))", "Average exact enumerated group-member count across Teams with complete membership coverage. Partial Teams are excluded.", "#,0.0")
        add_or_replace_measure(teams, "Teams with partial membership coverage", "CALCULATE([Teams], KEEPFILTERS('DimTeam'[MembershipCoverageStatus] = \"Partial\"))", "Teams for which Graph returned one or more member objects without an immutable identifier; exact links are preserved but the member total is incomplete.", "#,0")
        add_or_replace_measure(teams, "Teams without owner", "CALCULATE([Teams], KEEPFILTERS('DimTeam'[OwnerCount] = 0))", "Teams whose exact owner enumeration returned no owner.", "#,0")
    if team_members:
        add_or_replace_measure(team_members, "Team memberships", "COUNTROWS('FactTeamMember')", "Exact group membership edges for Teams.", "#,0")
        add_or_replace_measure(team_members, "Teams with guests", "CALCULATE(DISTINCTCOUNT('FactTeamMember'[TenantTeamKey]), KEEPFILTERS('FactTeamMember'[UserType] = \"Guest\"))", "Teams containing at least one exact guest membership.", "#,0")
        add_or_replace_measure(team_members, "Team membership role rate", "DIVIDE([Team memberships], CALCULATE([Team memberships], REMOVEFILTERS('FactTeamMember'[Role])))", "Share of exact Team membership edges by owner/member role.", "0.0%")

    relationship_overview = by_name.get("FactRelationshipOverview")
    if relationship_overview:
        add_or_replace_measure(relationship_overview, "Relationship edges", "SUM('FactRelationshipOverview'[RelationshipCount])", "Exact observed relationship edges by supported source type.", "#,0")
        add_or_replace_measure(relationship_overview, "Relationship type rate", "DIVIDE([Relationship edges], CALCULATE([Relationship edges], REMOVEFILTERS('FactRelationshipOverview'[RelationshipType])))", "Share of exact observed edges by relationship type.", "0.0%")

    hosting = by_name["FactMailboxHosting"]
    add_or_replace_measure(hosting, "Mailbox hosting rate", "DIVIDE([Hosted mailboxes], CALCULATE([Hosted mailboxes], REMOVEFILTERS('FactMailboxHosting'[HostingLocation])))", "Share of each mailbox-hosting location within the filtered reconciled mailbox population.", "0.0%")
    assignment = by_name.get("LicenseAssignmentPath")
    if assignment:
        add_or_replace_measure(assignment, "Assignment route rate", "DIVIDE([Assignment paths], CALCULATE([Assignment paths], REMOVEFILTERS('LicenseAssignmentPath'[AssignmentRoute])))", "Share of each observed license-assignment route.", "0.0%")
    user_license = by_name.get("FactUserLicense")
    if user_license:
        add_or_replace_measure(user_license, "Assignment state rate", "DIVIDE([License assignments], CALCULATE([License assignments], REMOVEFILTERS('FactUserLicense'[AssignmentStateLabel])))", "Share of license assignments by normalized assignment state.", "0.0%")
    plans = by_name.get("DimLicenseServicePlan")
    if plans:
        add_or_replace_measure(plans, "Service plan provisioning rate", "DIVIDE([License service plans], CALCULATE([License service plans], REMOVEFILTERS('DimLicenseServicePlan'[ProvisioningStatus])))", "Share of each service-plan provisioning status.", "0.0%")
    updates = by_name.get("FactWindowsUpdateAlert")
    if updates:
        add_or_replace_measure(updates, "Update aggregate state rate", "DIVIDE([Windows update alert records], CALCULATE([Windows update alert records], REMOVEFILTERS('FactWindowsUpdateAlert'[AggregateState])))", "Share of Windows Update evidence records by aggregate state.", "0.0%")
    hardware = by_name.get("DeviceHardware")
    if hardware:
        add_or_replace_measure(hardware, "Manufacturer record rate", "DIVIDE([Hardware records], CALCULATE([Hardware records], REMOVEFILTERS('DeviceHardware'[Manufacturer])))", "Share of source-reported hardware records by manufacturer.", "0.0%")
    top_apps = by_name.get("TopApplication")
    if top_apps:
        add_or_replace_measure(top_apps, "Top application occurrence rate", "DIVIDE([Top application occurrences], CALCULATE([Top application occurrences], REMOVEFILTERS('TopApplication'[ApplicationProduct])))", "Share within the five highest-volume normalized application products; source-reported occurrences are not distinct devices.", "0.0%")
    user_device = by_name.get("FactUserDeviceRelationship")
    if user_device:
        user_device["measures"][:] = [
            measure for measure in user_device.get("measures", [])
            if measure.get("name") != "Relationship type rate"
        ]
        add_or_replace_measure(user_device, "User-device relationship type rate", "DIVIDE([User-device links], CALCULATE([User-device links], REMOVEFILTERS('FactUserDeviceRelationship'[RelationshipType])))", "Share of each observed user-device relationship type.", "0.0%")
    findings = by_name.get("EntityFinding")
    if findings:
        add_or_replace_measure(findings, "Finding records", "COUNTROWS('EntityFinding')", "Finding evidence records in the current entity and filter context.", "#,0")
        add_or_replace_measure(findings, "Affected devices", "CALCULATE(DISTINCTCOUNT('EntityFinding'[TenantDeviceKey]), KEEPFILTERS(NOT ISBLANK('EntityFinding'[TenantDeviceKey])))", "Distinct devices referenced by the current finding context.", "#,0")
        add_or_replace_measure(findings, "Finding severity rate", "DIVIDE([Finding records], CALCULATE([Finding records], REMOVEFILTERS('EntityFinding'[Severity])))", "Share of entity findings by severity.", "0.0%")


def add_executive_measures(tables):
    country_table = next((table for table in tables if table.get("name") == "DimCountry"), None)
    if not country_table:
        raise ValueError("Expected DimCountry in the semantic model")
    add_or_replace_measure(
        country_table,
        "Executive Windows 10 devices",
        "VAR _countries = VALUES('DimCountry'[CountryLabel]) "
        "RETURN CALCULATE([Devices], TREATAS(_countries, 'DimDevice'[CountryLabel]), "
        "KEEPFILTERS(FILTER('DimDevice', "
        "VAR _os = COALESCE('DimDevice'[OperatingSystem], \"\") "
        "VAR _version = COALESCE('DimDevice'[OperatingSystemVersion], \"\") "
        "VAR _isNumericVersion = LEFT(_version, 5) = \"10.0.\" "
        "VAR _build = IF(_isNumericVersion, IFERROR(VALUE(PATHITEM(SUBSTITUTE(_version, \".\", \"|\"), 3, INTEGER)), BLANK())) "
        "RETURN _os = \"Windows\" && (_version = \"Windows 10\" || (_build > 0 && _build < 22000)))))",
        "Windows 10 devices identified from the reported operating-system version in the selected country and ownership context. This is an OS count, not a hardware-readiness assessment.",
        "#,0",
    )
    add_or_replace_measure(
        country_table,
        "Executive Endpoint Analytics score",
        "VAR _countries = VALUES('DimCountry'[CountryLabel]) "
        "VAR _deviceKeys = CALCULATETABLE(VALUES('DimDevice'[TenantDeviceKey]), TREATAS(_countries, 'DimDevice'[CountryLabel])) "
        "VAR _intuneIds = CALCULATETABLE(VALUES('DeviceSource'[SourceObjectId]), TREATAS(_deviceKeys, 'DeviceSource'[TenantDeviceKey]), KEEPFILTERS('DeviceSource'[SourceSystem] = \"Intune\")) "
        "RETURN CALCULATE(AVERAGE('FactEndpointAnalyticsDevice'[EndpointAnalyticsScore]), TREATAS(_intuneIds, 'FactEndpointAnalyticsDevice'[DeviceId]))",
        "Average Endpoint Analytics score from 0 to 100, linked by exact Intune device identifier in the selected country and ownership context; it is a normalized score, not a population percentage.",
        "0.0",
    )
    add_or_replace_measure(
        country_table,
        "Executive Windows 10 not compatible devices",
        "VAR _countries = VALUES('DimCountry'[CountryLabel]) "
        "VAR _windows10DeviceKeys = CALCULATETABLE(VALUES('DimDevice'[TenantDeviceKey]), "
        "TREATAS(_countries, 'DimDevice'[CountryLabel]), "
        "KEEPFILTERS(FILTER('DimDevice', "
        "VAR _os = COALESCE('DimDevice'[OperatingSystem], \"\") "
        "VAR _version = COALESCE('DimDevice'[OperatingSystemVersion], \"\") "
        "VAR _isNumericVersion = LEFT(_version, 5) = \"10.0.\" "
        "VAR _build = IF(_isNumericVersion, IFERROR(VALUE(PATHITEM(SUBSTITUTE(_version, \".\", \"|\"), 3, INTEGER)), BLANK())) "
        "RETURN _os = \"Windows\" && (_version = \"Windows 10\" || (_build > 0 && _build < 22000))))) "
        "VAR _intuneDeviceIds = CALCULATETABLE(VALUES('DeviceSource'[SourceObjectId]), "
        "TREATAS(_windows10DeviceKeys, 'DeviceSource'[TenantDeviceKey]), "
        "KEEPFILTERS('DeviceSource'[SourceSystem] = \"Intune\")) "
        "VAR _observed = CALCULATE(COUNTROWS('FactEndpointAnalyticsUpgradeEligibility'), "
        "TREATAS(_intuneDeviceIds, 'FactEndpointAnalyticsUpgradeEligibility'[DeviceId]), "
        "KEEPFILTERS('FactEndpointAnalyticsUpgradeEligibility'[UpgradeEligibility] IN {\"upgraded\", \"unknown\", \"notCapable\", \"capable\", \"unknownFutureValue\"})) "
        "VAR _notCapable = CALCULATE(DISTINCTCOUNT('FactEndpointAnalyticsUpgradeEligibility'[DeviceId]), "
        "TREATAS(_intuneDeviceIds, 'FactEndpointAnalyticsUpgradeEligibility'[DeviceId]), "
        "KEEPFILTERS('FactEndpointAnalyticsUpgradeEligibility'[UpgradeEligibility] = \"notCapable\")) "
        "RETURN IF(_observed = 0, BLANK(), _notCapable)",
        "Distinct Windows 10 Intune devices explicitly reported as notCapable by Endpoint Analytics in the selected country and ownership context. Returns blank when no readiness evidence exists; unknown is never treated as incompatible.",
        "#,0",
    )
    for title, sku, measure, _label, _color in LICENSE_COUNTRY_SERIES:
        add_or_replace_measure(
            country_table,
            measure,
            "VAR _countryAssignments = CALCULATE([License assignments], "
            f"KEEPFILTERS('DimLicenseSku'[SkuPartNumber] == \"{sku}\")) "
            "VAR _allAssignments = CALCULATE([License assignments], "
            f"KEEPFILTERS('DimLicenseSku'[SkuPartNumber] == \"{sku}\"), "
            "REMOVEFILTERS('DimCountry')) "
            "RETURN DIVIDE(_countryAssignments, _allAssignments)",
            f"{title} assignments for each country divided by all observed {title} assignments. Assignment does not prove usage.",
            "0.0%",
        )

    quality_table = next((table for table in tables if table.get("name") == "FactDataQuality"), None)
    if not quality_table:
        raise ValueError("Expected FactDataQuality in the semantic model")
    for name, expression, description in QUALITY_INDICATORS:
        add_or_replace_measure(quality_table, name, expression, description, "#,0")
    add_or_replace_measure(
        quality_table,
        "Data quality score",
        "VAR _users = CALCULATE([Users], REMOVEFILTERS('DimUser'), REMOVEFILTERS('DimCountry')) "
        "VAR _devices = CALCULATE([Devices], REMOVEFILTERS('DimDevice'), REMOVEFILTERS('DimCountry')) "
        "VAR _userCountryCoverage = IF(_users > 0, MAX(0, DIVIDE(_users - CALCULATE([Quality findings], KEEPFILTERS('FactDataQuality'[FindingType] == \"UserCountryUnknown\")), _users))) "
        "VAR _primaryUserCoverage = IF(_devices > 0, MAX(0, DIVIDE(_devices - CALCULATE([Quality findings], KEEPFILTERS('FactDataQuality'[FindingType] == \"DeviceWithoutPrimaryUser\")), _devices))) "
        "VAR _deviceCountryCoverage = IF(_devices > 0, MAX(0, DIVIDE(_devices - CALCULATE([Quality findings], KEEPFILTERS('FactDataQuality'[FindingType] == \"DeviceCountryUnknown\")), _devices))) "
        "VAR _population = _users + _devices "
        "VAR _integrityCoverage = IF(_population > 0, MAX(0, DIVIDE(_population - [Integrity issues], _population))) "
        "VAR _rates = {_userCountryCoverage, _primaryUserCoverage, _deviceCountryCoverage, _integrityCoverage} "
        "RETURN MINX(FILTER(_rates, NOT ISBLANK([Value])), [Value])",
        "Lowest non-blank global completeness rate across user country, device primary-user, device country and integrity checks. Overlapping findings are assessed separately, never added together. Country and ownership selections are intentionally ignored because quality findings are not country-grained.",
        "0.0%",
    )
    add_or_replace_measure(
        quality_table,
        "Data quality status",
        "VAR _score = [Data quality score] RETURN SWITCH(TRUE(), ISBLANK(_score), \"Not assessed\", _score >= 0.9, \"Good\", _score >= 0.6, \"Average\", \"Poor\")",
        "Traffic-light assessment of the data quality score: Good at 90% or above, Average from 60% to below 90%, and Poor below 60%.",
        "General",
    )
    add_or_replace_measure(
        quality_table,
        "Data quality health",
        "VAR _score = [Data quality score] VAR _status = [Data quality status] RETURN IF(ISBLANK(_score), _status, _status & \" · \" & FORMAT(_score, \"0.0%\"))",
        "Compact status and score intended for the Executive Overview; open Workplace Health for the underlying findings.",
        "General",
    )

    user_table = next((table for table in tables if table.get("name") == "DimUser"), None)
    if not user_table:
        raise ValueError("Expected DimUser in the semantic model")
    add_or_replace_measure(
        user_table,
        "Account status share",
        "DIVIDE([Users], CALCULATE([Users], REMOVEFILTERS('DimUser'[AccountStatusLabel])))",
        "Share of the current account-status category within the filtered user population.",
        "0.0%",
    )


def validate_upgrade_eligibility_source(path: Path):
    columns, rows = read_csv(path)
    if columns != UPGRADE_ELIGIBILITY_COLUMNS:
        raise ValueError(
            "Upgrade eligibility CSV schema mismatch. "
            f"Expected {UPGRADE_ELIGIBILITY_COLUMNS}; received {columns}."
        )
    observed = {}
    for index, row in enumerate(rows, start=2):
        device_id = (row.get("DeviceId") or "").strip()
        state = (row.get("UpgradeEligibility") or "").strip()
        if not device_id:
            raise ValueError(f"Upgrade eligibility row {index} is missing DeviceId")
        if state not in UPGRADE_ELIGIBILITY_STATES:
            raise ValueError(
                f"Upgrade eligibility row {index} has unsupported state {state!r}"
            )
        previous = observed.setdefault(device_id.casefold(), state)
        if previous != state:
            raise ValueError(
                "Conflicting upgrade eligibility states for the same exact DeviceId: "
                f"{previous!r} and {state!r}"
            )
    return rows


def enrich_semantic_model(
    report: Path,
    local_path: Path | None,
    remote_path: Path | None,
    upgrade_eligibility_path: Path | None = None,
):
    model_path = next(report.parent.glob("*.SemanticModel/model.bim"), None)
    if not model_path:
        raise ValueError("Expected one BIM semantic model beside the report")
    model_json = load(model_path)
    model = model_json["model"]
    derived_data_dir = report.parent / "ReportData"
    data_dir = derived_data_dir
    parameter = next((item for item in model.get("expressions", []) if item.get("name") == "CMDBDataRoot"), None)
    if parameter and isinstance(parameter.get("expression"), str):
        parts = parameter["expression"].split('"')
        if len(parts) >= 3 and parts[1]:
            candidate = Path(parts[1]) / "PowerBI" / "CMDB-REPORTS" / "ReportData"
            if candidate.is_dir():
                data_dir = candidate
    identity, hosting_rows, metadata = mailbox_hosting_rows(
        report, local_path, remote_path, data_dir
    )
    upgrade_target = data_dir / "FactEndpointAnalyticsUpgradeEligibility.csv"
    if upgrade_eligibility_path:
        upgrade_eligibility_path = upgrade_eligibility_path.resolve()
        if not upgrade_eligibility_path.is_file():
            raise ValueError(
                f"Upgrade eligibility evidence file not found: {upgrade_eligibility_path}"
            )
        validate_upgrade_eligibility_source(upgrade_eligibility_path)
        if upgrade_eligibility_path != upgrade_target.resolve():
            shutil.copy2(upgrade_eligibility_path, upgrade_target)
    if upgrade_target.is_file():
        validate_upgrade_eligibility_source(upgrade_target)
    # Keep every refreshable report CSV under the configured private data root.
    # The PBIR folder itself can then remain in Git without carrying tenant data.
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

    _, application_rows = read_csv(data_dir / "DimDetectedApplication.csv")
    application_groups = {}
    for row in application_rows:
        key = (
            (row.get("DisplayName") or "Unknown application").strip(),
            (row.get("Publisher") or "Unknown publisher").strip(),
            (row.get("Platform") or "Unknown platform").strip(),
        )
        item = application_groups.setdefault(key, {"versions": set(), "occurrences": 0})
        item["versions"].add((row.get("Version") or "Unknown version").strip())
        try:
            item["occurrences"] += int(float(row.get("DeviceCount") or 0))
        except ValueError as exc:
            raise ValueError(f"Invalid detected-application DeviceCount: {row.get('DeviceCount')!r}") from exc
    top_applications = [
        {
            "ApplicationProduct": f"{display_name} · {publisher} · {platform}",
            "DisplayName": display_name,
            "Publisher": publisher,
            "Platform": platform,
            "VersionCount": len(item["versions"]),
            "ReportedDeviceCount": item["occurrences"],
        }
        for (display_name, publisher, platform), item in application_groups.items()
    ]
    top_applications.sort(key=lambda item: (-item["ReportedDeviceCount"], item["ApplicationProduct"].casefold()))
    top_application_path = data_dir / "TopApplication.csv"
    write_csv(top_application_path, TOP_APPLICATION_COLUMNS, top_applications[:5])

    tables = model["tables"]

    def import_table(name, path, columns, tenant_scoped, measures=None, type_overrides=None):
        table_columns = []
        for column in columns:
            item = {"name": column, "dataType": (type_overrides or {}).get(column, type_of(column))[0], "sourceColumn": column, "summarizeBy": "none"}
            if column in IDENTITY_COLUMNS or column.endswith("Key"):
                item["isHidden"] = True
            table_columns.append(item)
        source_expression = m_query(
            path, columns, identity, tenant_scoped, type_overrides
        )
        if parameter and path.resolve().parent == data_dir.resolve():
            source_expression[1] = (
                ' Source = Csv.Document(File.Contents(CMDBDataRoot & '
                f'"\\\\PowerBI\\\\CMDB-REPORTS\\\\ReportData\\\\{path.name}"), '
                '[Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),'
            )
        return {
            "name": name,
            "columns": table_columns,
            "partitions": [{
                "name": name,
                "mode": "import",
                "source": {"type": "m", "expression": source_expression},
            }],
            "measures": measures or [],
        }

    hosting_table = import_table("FactMailboxHosting", hosting_path, MAILBOX_HOSTING_COLUMNS, True)
    add_or_replace_measure(hosting_table, "Hosted mailboxes", "COALESCE(COUNTROWS('FactMailboxHosting'), 0)", "Distinct mailbox addresses after reconciliation of Exchange Online, RemoteMailbox and local Mailbox evidence.", "#,0")
    add_or_replace_measure(hosting_table, "Exchange Online mailboxes", "CALCULATE([Hosted mailboxes], KEEPFILTERS('FactMailboxHosting'[HostingLocation] == \"Exchange Online\"))", "Mailboxes observed in Exchange Online or as RemoteMailbox; duplicate addresses are counted once.", "#,0")
    add_or_replace_measure(hosting_table, "Exchange On-premises mailboxes", "CALCULATE([Hosted mailboxes], KEEPFILTERS('FactMailboxHosting'[HostingLocation] == \"Exchange On-premises\"))", "Local Mailbox recipients not already observed as Exchange Online or RemoteMailbox; duplicate addresses are counted once.", "#,0")
    add_or_replace_measure(hosting_table, "Exchange Online share", "DIVIDE([Exchange Online mailboxes], [Hosted mailboxes])", "Exchange Online mailboxes divided by the reconciled mailbox footprint.", "0.0%")
    add_or_replace_measure(hosting_table, "Exchange On-premises share", "DIVIDE([Exchange On-premises mailboxes], [Hosted mailboxes])", "Exchange On-premises mailboxes divided by the reconciled mailbox footprint.", "0.0%")
    add_or_replace_measure(hosting_table, "User mailboxes", "CALCULATE([Hosted mailboxes], KEEPFILTERS('FactMailboxHosting'[MailboxTypeGroup] == \"User mailbox\"))", "UserMailbox and RemoteUserMailbox recipients after mailbox-hosting reconciliation.", "#,0")
    add_or_replace_measure(hosting_table, "Shared mailboxes", "CALCULATE([Hosted mailboxes], KEEPFILTERS('FactMailboxHosting'[MailboxTypeGroup] == \"Shared mailbox\"))", "SharedMailbox and RemoteSharedMailbox recipients after mailbox-hosting reconciliation.", "#,0")
    add_or_replace_measure(hosting_table, "Other mailbox types", "CALCULATE([Hosted mailboxes], KEEPFILTERS('FactMailboxHosting'[MailboxTypeGroup] == \"Other mailbox types\"))", "Room, scheduling, discovery and other recipient types after mailbox-hosting reconciliation.", "#,0")
    add_or_replace_measure(hosting_table, "User mailbox share", "DIVIDE([User mailboxes], [Hosted mailboxes])", "User mailbox family divided by all reconciled mailboxes.", "0.0%")
    add_or_replace_measure(hosting_table, "Shared mailbox share", "DIVIDE([Shared mailboxes], [Hosted mailboxes])", "Shared mailbox family divided by all reconciled mailboxes.", "0.0%")
    add_or_replace_measure(hosting_table, "Other mailbox type share", "DIVIDE([Other mailbox types], [Hosted mailboxes])", "Other mailbox types divided by all reconciled mailboxes.", "0.0%")
    add_or_replace_measure(hosting_table, "Mailbox type share", "DIVIDE([Hosted mailboxes], CALCULATE([Hosted mailboxes], REMOVEFILTERS('FactMailboxHosting'[MailboxTypeGroup])))", "Share of the current mailbox-type category within the filtered reconciled mailbox population.", "0.0%")

    optional_tables = [
        ("FactDeviceApplication", FACT_DEVICE_APPLICATION_COLUMNS, {"SourceCollectedDateTime": ("dateTime", "type datetime")}),
        ("FactADIntuneCoverage", FACT_AD_INTUNE_COVERAGE_COLUMNS, {"Enabled": ("boolean", "type logical"), "SourceCollectedDateTime": ("dateTime", "type datetime")}),
        ("DimSharePointSite", DIM_SHAREPOINT_SITE_COLUMNS, {"LastActivityDate": ("dateTime", "type datetime"), "StorageUsedBytes": ("int64", "Int64.Type"), "StorageAllocatedBytes": ("int64", "Int64.Type"), "IsDeleted": ("boolean", "type logical"), "SourceCollectedDateTime": ("dateTime", "type datetime")}),
        ("DimTeam", DIM_TEAM_COLUMNS, {"CreatedDateTime": ("dateTime", "type datetime"), "LastActivityDate": ("dateTime", "type datetime"), "OwnerCount": ("int64", "Int64.Type"), "MemberCount": ("int64", "Int64.Type"), "GuestCount": ("int64", "Int64.Type"), "UnresolvedMemberCount": ("int64", "Int64.Type"), "IsArchived": ("boolean", "type logical"), "SourceCollectedDateTime": ("dateTime", "type datetime")}),
        ("FactTeamMember", FACT_TEAM_MEMBER_COLUMNS, {"SourceCollectedDateTime": ("dateTime", "type datetime")}),
        ("FactUserActivity", FACT_USER_ACTIVITY_COLUMNS, {"ReportRefreshDate": ("dateTime", "type datetime"), "ExchangeLastActivityDate": ("dateTime", "type datetime"), "OneDriveLastActivityDate": ("dateTime", "type datetime"), "SharePointLastActivityDate": ("dateTime", "type datetime"), "TeamsLastActivityDate": ("dateTime", "type datetime"), "LastActivityDate": ("dateTime", "type datetime"), "HasAnyM365Activity": ("boolean", "type logical"), "SourceCollectedDateTime": ("dateTime", "type datetime")}),
    ]
    for table_name, columns, overrides in optional_tables:
        source_path = data_dir / f"{table_name}.csv"
        if not source_path.is_file():
            # A header-only private source keeps the PBIR refreshable before the
            # first run of a newly introduced collector. Blank means not
            # collected; no metric is inferred from another grain.
            write_csv(source_path, columns, [])
        imported = import_table(table_name, source_path, columns, True, type_overrides=overrides)
        tables[:] = [table for table in tables if table.get("name") != table_name]
        tables.append(imported)

    detected_application_path = data_dir / "DimDetectedApplication.csv"
    existing_detected_application = next(
        (table for table in tables if table.get("name") == "DimDetectedApplication"),
        None,
    )
    detected_application_table = import_table(
        "DimDetectedApplication",
        detected_application_path,
        DIM_DETECTED_APPLICATION_COLUMNS,
        False,
        type_overrides={
            "DeviceCount": ("int64", "Int64.Type"),
            "ReportedDeviceCount": ("int64", "Int64.Type"),
            "ExactRelatedDeviceCount": ("int64", "Int64.Type"),
            "SourceCollectedDateTime": ("dateTime", "type datetime"),
        },
    )
    if existing_detected_application:
        detected_application_table["measures"] = existing_detected_application.get("measures", [])
    tables[:] = [
        table for table in tables
        if table.get("name") != "DimDetectedApplication"
    ]
    tables.append(detected_application_table)

    _, device_source_rows = read_csv(data_dir / "DeviceSource.csv")
    intune_device_bridge_rows = build_intune_device_bridge(device_source_rows, identity)
    intune_device_bridge_path = data_dir / "DimIntuneManagedDevice.csv"
    write_csv(
        intune_device_bridge_path,
        INTUNE_DEVICE_BRIDGE_COLUMNS,
        intune_device_bridge_rows,
    )
    intune_device_bridge_table = import_table(
        "DimIntuneManagedDevice",
        intune_device_bridge_path,
        INTUNE_DEVICE_BRIDGE_COLUMNS,
        True,
    )
    tables[:] = [
        table for table in tables
        if table.get("name") != "DimIntuneManagedDevice"
    ]
    tables.append(intune_device_bridge_table)

    relationship_sources = [
        ("PrimaryUser", "FactUserDeviceRelationship.csv", "Exact curated user-device links", lambda row: True),
        ("HasMailbox", "FactMailbox.csv", "Exact curated user-mailbox links", lambda row: bool((row.get("CmdbUserId") or "").strip())),
        ("AssignedLicense", "FactUserLicense.csv", "Exact curated user-SKU assignments", lambda row: True),
        ("MemberOfGroup", "FactTeamMember.csv", "Exact Microsoft Teams membership enumeration", lambda row: True),
        ("DeviceHasApplication", "FactDeviceApplication.csv", "Exact Intune detected-app to managed-device links", lambda row: True),
        ("DeviceInAutopilot", "FactAutopilotDevice.csv", "Exact Intune managed-device to Autopilot links", lambda row: bool((row.get("ManagedDeviceId") or "").strip())),
    ]
    relationship_rows = []
    for relationship_type, file_name, evidence, predicate in relationship_sources:
        source_path = data_dir / file_name
        count = 0
        if source_path.is_file():
            count = count_csv_rows(source_path, predicate)
        relationship_rows.append({
            "RelationshipType": relationship_type,
            "RelationshipCount": count,
            "EvidenceSource": evidence,
        })
    relationship_overview_path = data_dir / "FactRelationshipOverview.csv"
    write_csv(relationship_overview_path, RELATIONSHIP_OVERVIEW_COLUMNS, relationship_rows)
    relationship_overview_table = import_table(
        "FactRelationshipOverview",
        relationship_overview_path,
        RELATIONSHIP_OVERVIEW_COLUMNS,
        False,
        type_overrides={"RelationshipCount": ("int64", "Int64.Type")},
    )
    tables[:] = [table for table in tables if table.get("name") != "FactRelationshipOverview"]
    tables.append(relationship_overview_table)

    country_table = import_table("DimCountry", country_path, ["CountryLabel"], False)
    add_or_replace_calculated_column(
        country_table,
        COUNTRY_FOOTPRINT_COLUMN,
        (
            "IF('DimCountry'[CountryLabel] = \"Unknown / unassigned\", "
            f'\"{COUNTRY_FOOTPRINT_UNKNOWN}\", \'DimCountry\'[CountryLabel])'
        ),
        "Compact country label used only by Executive Overview footprint charts; the full country key remains unchanged.",
    )
    add_or_replace_measure(country_table, "Executive workplace devices", "CALCULATE([Devices], TREATAS(VALUES('DimCountry'[CountryLabel]), 'DimDevice'[CountryLabel]))", "Devices in the selected shared country context; the ownership slicer may further restrict the population.", "#,0")
    add_or_replace_measure(country_table, "Executive managed device share", "CALCULATE([Managed device share], TREATAS(VALUES('DimCountry'[CountryLabel]), 'DimDevice'[CountryLabel]))", "Managed-device share in the selected shared country context.", "0.0%")
    add_or_replace_measure(country_table, "Executive compliant device share", "CALCULATE([Compliant device share], TREATAS(VALUES('DimCountry'[CountryLabel]), 'DimDevice'[CountryLabel]))", "Compliant-device share in the selected shared country context.", "0.0%")
    add_or_replace_measure(country_table, "Executive corporate devices", "CALCULATE([Corporate devices], TREATAS(VALUES('DimCountry'[CountryLabel]), 'DimDevice'[CountryLabel]))", "Corporate devices in the selected shared country context; ownership selection is intentionally ignored by the base measure.", "#,0")
    add_or_replace_measure(country_table, "Executive corporate country share", "VAR _selected = [Executive corporate devices] VAR _all = CALCULATE([Executive corporate devices], REMOVEFILTERS('DimCountry')) RETURN DIVIDE(_selected, _all)", "Corporate devices for each country divided by all corporate devices.", "0.0%")
    add_or_replace_measure(country_table, "Executive user country share", "DIVIDE([Users], CALCULATE([Users], REMOVEFILTERS('DimCountry')))", "Users for each country divided by all users, including unknown or unassigned country.", "0.0%")
    add_or_replace_measure(country_table, "Executive mailbox country share", "DIVIDE([Hosted mailboxes], CALCULATE([Hosted mailboxes], REMOVEFILTERS('DimCountry')))", "Reconciled mailboxes for each country divided by all reconciled mailboxes, including unknown or unassigned country.", "0.0%")
    for name, category_column, description in [
        ("Executive device form factor share", "Device form factor", "Share of each device form-factor category in the selected shared country and ownership context."),
        ("Executive device ownership share", "Device ownership group", "Share of each device-ownership category in the selected shared country context."),
        ("Executive Windows version share", "Windows version group", "Share of each Windows-version category in the selected shared country and ownership context."),
    ]:
        add_or_replace_measure(
            country_table,
            name,
            "VAR _countries = VALUES('DimCountry'[CountryLabel]) "
            "VAR _selected = CALCULATE([Devices], TREATAS(_countries, 'DimDevice'[CountryLabel])) "
            f"VAR _all = CALCULATE([Devices], REMOVEFILTERS('DimDevice'[{category_column}]), TREATAS(_countries, 'DimDevice'[CountryLabel])) "
            "RETURN DIVIDE(_selected, _all)",
            description,
            "0.0%",
        )
    if upgrade_target.is_file():
        upgrade_table = import_table(
            "FactEndpointAnalyticsUpgradeEligibility",
            upgrade_target,
            UPGRADE_ELIGIBILITY_COLUMNS,
            True,
        )
        tables[:] = [
            table for table in tables
            if table.get("name") != "FactEndpointAnalyticsUpgradeEligibility"
        ]
        tables.append(upgrade_table)

    top_application_table = import_table(
        "TopApplication", top_application_path, TOP_APPLICATION_COLUMNS, False,
        type_overrides={"VersionCount": ("int64", "Int64.Type"), "ReportedDeviceCount": ("int64", "Int64.Type")},
    )
    add_or_replace_measure(top_application_table, "Top application occurrences", "SUM('TopApplication'[ReportedDeviceCount])", "Source-reported application-device occurrences for the five highest-volume normalized products; not distinct devices.", "#,0")
    tables[:] = [table for table in tables if table.get("name") != "TopApplication"]
    tables.append(top_application_table)

    add_operational_measures([hosting_table, *tables])
    add_executive_measures([country_table, hosting_table, *tables])

    tables[:] = [table for table in tables if table.get("name") not in {"DimCountry", "FactMailboxHosting"}]
    tables.extend([country_table, hosting_table])
    relationships = model.setdefault("relationships", [])
    relationships[:] = [
        relationship for relationship in relationships
        if relationship.get("name") not in {"cmdb-executive-country-user", "cmdb-executive-country-mailbox", "cmdb-app-device", "cmdb-app-device-parent", "cmdb-app-product", "cmdb-user-activity", "cmdb-team-member-team", "cmdb-team-member-user"}
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
    table_names = {table.get("name") for table in tables}
    if {"FactDeviceApplication", "DimIntuneManagedDevice", "DimDevice", "DimDetectedApplication"}.issubset(table_names):
        relationships.extend([
            {"name":"cmdb-app-device","fromTable":"FactDeviceApplication","fromColumn":"Tenant Intune device key","toTable":"DimIntuneManagedDevice","toColumn":"TenantIntuneDeviceKey","fromCardinality":"many","toCardinality":"one","crossFilteringBehavior":"oneDirection","isActive":True},
            {"name":"cmdb-app-device-parent","fromTable":"DimIntuneManagedDevice","fromColumn":"TenantDeviceKey","toTable":"DimDevice","toColumn":"TenantDeviceKey","fromCardinality":"many","toCardinality":"one","crossFilteringBehavior":"oneDirection","isActive":True},
            {"name":"cmdb-app-product","fromTable":"FactDeviceApplication","fromColumn":"TenantApplicationKey","toTable":"DimDetectedApplication","toColumn":"TenantApplicationKey","fromCardinality":"many","toCardinality":"one","crossFilteringBehavior":"oneDirection","isActive":True},
        ])
    if "FactUserActivity" in table_names:
        relationships.append({"name":"cmdb-user-activity","fromTable":"FactUserActivity","fromColumn":"TenantUserKey","toTable":"DimUser","toColumn":"TenantUserKey","fromCardinality":"many","toCardinality":"one","crossFilteringBehavior":"oneDirection","isActive":True})
    if {"FactTeamMember", "DimTeam"}.issubset(table_names):
        relationships.append({"name":"cmdb-team-member-team","fromTable":"FactTeamMember","fromColumn":"TenantTeamKey","toTable":"DimTeam","toColumn":"TenantTeamKey","fromCardinality":"many","toCardinality":"one","crossFilteringBehavior":"oneDirection","isActive":True})
        relationships.append({"name":"cmdb-team-member-user","fromTable":"FactTeamMember","fromColumn":"TenantUserKey","toTable":"DimUser","toColumn":"TenantUserKey","fromCardinality":"many","toCardinality":"one","crossFilteringBehavior":"oneDirection","isActive":True})
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
    props["labelDisplayUnits"] = lit(1)
    is_ratio = any(
        re.search(rf"\b{word}\b", measure, flags=re.IGNORECASE)
        for word in ("share", "rate", "coverage")
    )
    props["labelPrecision"] = lit(
        1 if precision is None and is_ratio else (precision or 0)
    )
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
    category_colors = BAR_CATEGORY_COLORS.get((table, column))
    if category_colors:
        set_category_colors(visual, table, column, category_colors)


def add_text(pages, visuals, page_id, value, x, y, w, h, template="overviewv4"):
    visual = clone_visual(pages, "devices", "devicesv2")
    set_text(visual, value)
    return put(visuals, visual, page_id, x, y, w, h)


def kpi_icon_key(table, measure):
    """Return a restrained icon family for headline KPIs only."""
    lowered = measure.casefold()
    if "application" in lowered or "app " in lowered:
        return "devices"
    if "mailbox" in lowered:
        return "mailboxes"
    if "license" in lowered or "assignment" in lowered:
        return "licenses"
    if "quality" in lowered or "finding" in lowered or "warning" in lowered or "alert" in lowered:
        return "quality"
    if "compliant" in lowered or "compliance" in lowered:
        return "compliance"
    if "managed" in lowered or "management" in lowered:
        return "management"
    if "endpoint analytics" in lowered or "source collection" in lowered:
        return "analytics"
    if "relationship" in lowered or "resolved" in lowered or "user-link" in lowered or "device-link" in lowered:
        return "relationships"
    if table == "DimUser" or "user" in lowered or "account" in lowered:
        return "users"
    if "device" in lowered or table in {"DimDevice", "DeviceHardware", "FactAutopilotDevice"}:
        return "devices"
    return None


def add_kpi_icon(pages, visuals, page_id, key, x, y, w):
    resource = KPI_ICON_RESOURCES.get(key)
    if not resource:
        return None
    icon = add_registered_image(pages, visuals, page_id, resource, x + w - 34, y + 8, 22, 22)
    icon["name"] = f"{page_id}kpi{len(visuals)}icon"
    icon["position"]["z"] = 850
    icon["position"]["tabOrder"] = 850
    return icon


def add_card(pages, visuals, page_id, table, measure, title, x, y, w=296, h=96, precision=None):
    visual = clone_visual_by_type(pages, "devices", ["devicesv8", "devicesv7", "devicesv9"], {"cardVisual"})
    set_card(visual, table, measure, title, precision)
    return put(visuals, visual, page_id, x, y, w, h)


def add_bar(pages, visuals, page_id, table, column, measure_table, measure, title, x, y, w=608, h=220):
    visual = clone_visual_by_type(pages, "devices", ["devicesv12", "devicesv11", "devicesv13"], {"barChart", "clusteredBarChart"})
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
    configure_searchable_slicer(visual)
    return put(visuals, visual, page_id, x, y, w, h)


def add_registered_image(pages, visuals, page_id, resource_name, x, y, w, h):
    visual = clone_visual(pages, "devices", "devicesv0")
    clear_filter(visual)
    visual["visual"] = {
        "visualType": "image",
        "objects": {
            "general": [{"properties": {
                "imageUrl": {"expr": {"ResourcePackageItem": {
                    "PackageName": "RegisteredResources",
                    "PackageType": 1,
                    "ItemName": resource_name,
                }}}
            }}]
        },
        "drillFilterOtherVisuals": True,
    }
    return put(visuals, visual, page_id, x, y, w, h)


def decorate_page_kpis(pages, page_id):
    """Add KPI icons only after every page builder has used its stable templates."""
    visual_root = pages / page_id / "visuals"
    for card_path in sorted(visual_root.glob("*/visual.json")):
        card = load(card_path)
        if card.get("visual", {}).get("visualType") != "cardVisual":
            continue
        position = card.get("position", {})
        if position.get("height", 0) < 80:
            continue
        projections = (
            card.get("visual", {})
            .get("query", {})
            .get("queryState", {})
            .get("Data", {})
            .get("projections", [])
        )
        if not projections:
            continue
        measure_field = projections[0].get("field", {}).get("Measure", {})
        table = measure_field.get("Expression", {}).get("SourceRef", {}).get("Entity", "")
        measure = measure_field.get("Property", "")
        key = "licenses" if table == "DimLicenseSku" else kpi_icon_key(table, measure)
        resource = KPI_ICON_RESOURCES.get(key)
        if not resource:
            continue
        icon = add_registered_image(
            pages, [], page_id, resource,
            position["x"] + position["width"] - 34,
            position["y"] + 8,
            22, 22,
        )
        icon["name"] = f"{card['name']}kpiicon"
        icon["position"]["z"] = 850
        icon["position"]["tabOrder"] = 850
        write(visual_root / icon["name"] / "visual.json", icon)


def section_icon_key(title):
    lowered = title.casefold()
    if "finding" in lowered:
        return "findings"
    if "activity" in lowered:
        return "activity"
    if "hardware" in lowered:
        return "hardware"
    if "enrollment" in lowered or "source identity" in lowered:
        return "enrollment"
    if "identity" in lowered or "profile" in lowered or "association" in lowered:
        return "identity"
    return None


def add_section_icon(pages, visuals, page_id, title, x, y, w):
    key = section_icon_key(title)
    resource = SECTION_ICON_RESOURCES.get(key)
    if not resource:
        return None
    icon = add_registered_image(pages, visuals, page_id, resource, x + w - 32, y + 5, 20, 20)
    icon["name"] = f"{page_id}section{len(visuals)}icon"
    icon["position"]["z"] = 840
    icon["position"]["tabOrder"] = 840
    return icon


def decorate_detail_sections(pages, visuals, page_id):
    """Add section icons after the detail builder has finished editing its tables."""
    tables = [visual for visual in visuals if visual.get("visual", {}).get("visualType") == "tableEx"]
    for table in tables:
        title = (
            table.get("visual", {})
            .get("visualContainerObjects", {})
            .get("title", [{}])[0]
            .get("properties", {})
            .get("text", {})
            .get("expr", {})
            .get("Literal", {})
            .get("Value", "")
        ).strip("'")
        if section_icon_key(title):
            position = table["position"]
            add_section_icon(
                pages, visuals, page_id, title,
                position["x"], position["y"], position["width"],
            )


def add_copyright(pages, page_id):
    """Persist the approved compact, clickable copyright on every page."""
    visual = clone_visual(pages, "devices", "devicesv3")
    set_text(visual, COPYRIGHT_TEXT)
    paragraphs = visual["visual"]["objects"]["general"][0]["properties"].get("paragraphs", [])
    for paragraph in paragraphs:
        paragraph["horizontalTextAlignment"] = "right"
        for run in paragraph.get("textRuns", []):
            run.setdefault("textStyle", {}).update({
                "fontFamily": "Segoe UI",
                "fontSize": "10px",
                "color": "#526577",
            })
    visual["visual"].setdefault("visualContainerObjects", {})["visualLink"] = [{
        "properties": {
            "show": lit(True),
            "type": lit("WebUrl"),
            "webUrl": lit(COPYRIGHT_URL),
            "tooltip": lit("Open WorkplaceCloudHub"),
            "showDefaultTooltip": lit(False),
        }
    }]
    visual["name"] = f"{page_id}copyright"
    visual["position"] = {
        "x": 736, "y": 864, "width": 520, "height": 28,
        "z": 1001, "tabOrder": 1001,
    }
    write(pages / page_id / "visuals" / visual["name"] / "visual.json", visual)
    return visual


def position_version_in_footer(pages, page_id):
    """Place the report version in the visual center of the shared footer."""
    path = pages / page_id / "visuals" / f"{page_id}v1" / "visual.json"
    if not path.is_file():
        return
    visual = load(path)
    set_text(visual, "VERSION 1.0.0")
    visual["position"].update(x=536, y=864, width=208, height=28)
    paragraphs = (
        visual.get("visual", {}).get("objects", {}).get("general", [{}])[0]
        .get("properties", {}).get("paragraphs", [])
    )
    for paragraph in paragraphs:
        paragraph["horizontalTextAlignment"] = "center"
    write(path, visual)


def persist_page_header_icon(pages, page_id, resource_name):
    """Add one non-interactive icon without depending on builder visual order."""
    icon = add_registered_image(pages, [], page_id, resource_name, 24, 18, 40, 40)
    icon["name"] = f"{page_id}icon"
    icon["position"]["z"] = 1000
    icon["position"]["tabOrder"] = 1000
    write(pages / page_id / "visuals" / icon["name"] / "visual.json", icon)


def add_overview_ratio_icon(pages, visuals, key, x, y):
    """Place a small, non-interactive illustration inside an overview ratio card."""
    icon = add_registered_image(
        pages, visuals, "overview", OVERVIEW_RATIO_ICON_RESOURCES[key],
        x + 198, y + 8, 30, 30,
    )
    icon["name"] = f"overview{key}icon"
    icon["position"]["z"] = 900
    icon["position"]["tabOrder"] = 900
    return icon


def configure_searchable_slicer(visual):
    """Use the native dropdown search without persisting a data selection."""
    objects = visual["visual"].setdefault("objects", {})
    objects["data"] = [{"properties": {"mode": lit("Dropdown")}}]
    general = objects.setdefault("general", [{"properties": {}}])
    general[0].setdefault("properties", {})["selfFilterEnabled"] = lit(True)
    # Strict single-select forces the first entity to be selected when the page
    # opens directly. Keep All as the neutral state; the gated detail measures
    # render only after one value is selected or supplied by drill-through.
    objects.pop("selection", None)


def compact_detail_header(visuals, width):
    """Reserve the upper-right canvas for compact detail-page controls."""
    for visual in visuals[:3]:
        visual["position"]["width"] = width


def compact_page_header(visuals, width=560):
    """Reserve the upper-right canvas for compact page filters."""
    for visual in visuals[:3]:
        visual["position"]["width"] = width
    paragraphs = visuals[0].get("visual", {}).get("objects", {}).get("general", [{}])[0].get("properties", {}).get("paragraphs", [])
    for paragraph in paragraphs:
        for run in paragraph.get("textRuns", []):
            # Long page titles otherwise render from their trailing edge in
            # Power BI Desktop when compact slicers reserve the right side.
            font_size = "18px" if len(run.get("value", "")) > 45 else "20px"
            run.setdefault("textStyle", {})["fontSize"] = font_size


def add_table(pages, visuals, page_id, fields, title, x, y, w, h):
    # Cache the fleet table before that page is rebuilt so later page builders
    # do not depend on visual identifiers produced earlier in the same run.
    visual = clone_visual_by_type(pages, "devices", ["devicesv15", "devicesv13", "devicesv16"], {"tableEx"})
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
    compact_page_header(visuals, 560)
    add_slicer(pages, visuals, "risk", "DimDevice", "CountryLabel", "Country", 600, y=14, w=120, h=80)
    add_slicer(pages, visuals, "risk", "DimDevice", "ComplianceStateLabel", "Compliance", 728, y=14, w=120, h=80)
    add_slicer(pages, visuals, "risk", "DimDevice", "ManagementStateLabel", "Management", 856, y=14, w=120, h=80)
    add_slicer(pages, visuals, "risk", "EntityFinding", "Severity", "Finding severity", 984, y=14, w=128, h=80)
    add_slicer(pages, visuals, "risk", "EntityFinding", "FindingType", "Finding type", 1120, y=14, w=136, h=80)
    for index, item in enumerate([
        ("DimDevice", "Noncompliant devices", "Explicitly noncompliant"),
        ("FactWindowsUpdateAlert", "Windows update alerts needing attention", "Update alerts needing attention"),
        ("FactDataQuality", "Critical findings", "Critical findings"),
        ("FactDataQuality", "Warnings", "Warning finding records"),
    ]):
        add_card(pages, visuals, "risk", *item, 24 + index * 312, 136, h=88)
    add_ratio_bar(pages, visuals, "risk", "DimDevice", "ComplianceStateLabel", "DimDevice", "Compliance state rate", "Devices", "Compliance — rate and device count", 24, 240, w=608, h=184)
    add_ratio_bar(pages, visuals, "risk", "FactWindowsUpdateAlert", "AggregateState", "FactWindowsUpdateAlert", "Update aggregate state rate", "Windows update alert records", "Windows Update evidence by state", 648, 240, w=608, h=184)
    add_table(
        pages, visuals, "risk",
        [
            ("EntityFinding", "Severity", "Severity", False),
            ("EntityFinding", "FindingType", "Finding", False),
            ("EntityFinding", "Affected devices", "Affected devices", True),
            ("EntityFinding", "Finding records", "Records", True),
            ("EntityFinding", "Description", "Evidence", False),
            ("EntityFinding", "RecommendedAction", "Suggested review", False),
        ],
        "Quality findings — affected devices are distinct exact device keys", 24, 440, 1232, 404,
    )
    return page, visuals


def build_lifecycle(pages):
    page, visuals = new_page(
        pages, "lifecycle", "Transformation & Lifecycle",
        "Assess the observed estate, Autopilot footprint, Endpoint Analytics and source readiness. No migration plan, purchase date, warranty, support end date or application lifecycle source is present.",
    )
    compact_page_header(visuals, 712)
    add_slicer(pages, visuals, "lifecycle", "DimDevice", "CountryLabel", "Country", 752, y=14, w=160, h=80)
    add_slicer(pages, visuals, "lifecycle", "DimDevice", "OperatingSystemLabel", "Operating system", 920, y=14, w=160, h=80)
    add_slicer(pages, visuals, "lifecycle", "DimDevice", "ManagementStateLabel", "Management", 1088, y=14, w=168, h=80)
    for index, item in enumerate([
        ("DimDevice", "Devices", "Observed devices"),
        ("FactADIntuneCoverage", "AD to Intune coverage rate", "AD → Intune coverage"),
        ("FactAutopilotDevice", "Selected Autopilot devices", "Autopilot · exact matches"),
        ("FactEndpointAnalyticsDevice", "Selected Endpoint Analytics score", "Endpoint Analytics score"),
    ]):
        add_card(pages, visuals, "lifecycle", *item, 24 + index * 312, 136, h=88)
    add_ratio_bar(pages, visuals, "lifecycle", "DimDevice", "OperatingSystemLabel", "DimDevice", "Operating system rate", "Devices", "Operating systems — observed, not EOL", 24, 240, w=296, h=204)
    add_ratio_bar(pages, visuals, "lifecycle", "DimDevice", "ManagementStateLabel", "DimDevice", "Management state rate", "Devices", "Management coverage", 336, 240, w=296, h=204)
    add_ratio_bar(pages, visuals, "lifecycle", "FactEndpointAnalyticsUpgradeEligibility", "UpgradeEligibility", "FactEndpointAnalyticsUpgradeEligibility", "Upgrade eligibility rate", "Selected upgrade eligibility devices", "Windows 11 readiness evidence", 648, 240, w=296, h=204)
    add_ratio_bar(pages, visuals, "lifecycle", "FactAutopilotDevice", "EnrollmentState", "FactAutopilotDevice", "Autopilot enrollment rate", "Selected Autopilot devices", "Autopilot — exact Intune matches", 960, 240, w=296, h=204)
    for index, item in enumerate([
        ("FactEndpointAnalyticsDevice", "Selected startup performance score", "Startup"),
        ("FactEndpointAnalyticsDevice", "Selected app reliability score", "App reliability"),
        ("FactEndpointAnalyticsDevice", "Selected work from anywhere score", "Work from anywhere"),
    ]):
        add_compact_card(pages, visuals, "lifecycle", *item, 648 + index * 204, 452, 194, 58)
    add_table(
        pages, visuals, "lifecycle",
        [
            ("SourceHealth", "SourceName", "Source", False),
            ("SourceHealth", "Status", "Collection status", False),
            ("SourceHealth", "Coverage", "Coverage", False),
            ("SourceHealth", "SourceRowsLabel", "Rows", False),
            ("SourceHealth", "CompletedDateTimeLabel", "Completed UTC", False),
        ],
        "Source readiness — lifecycle decisions remain bounded by collected evidence", 24, 526, 1232, 318,
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


def build_licensing(pages):
    page, visuals = new_page(
        pages, "licenses", "Licensing & Assignments",
        "Separate capacity, assignment paths and review candidates. Sign-in evidence prioritizes review but does not prove Microsoft 365 license non-use.",
    )
    compact_page_header(visuals, 712)
    add_slicer(pages, visuals, "licenses", "DimUser", "CountryLabel", "Country", 752, y=14, w=160, h=80)
    add_slicer(pages, visuals, "licenses", "DimLicenseSku", "SkuPartNumberLabel", "Product / SKU", 920, y=14, w=160, h=80)
    add_slicer(pages, visuals, "licenses", "FactUserLicense", "AssignmentStateLabel", "Assignment state", 1088, y=14, w=168, h=80)
    for index, item in enumerate([
        ("DimLicenseSku", "License SKUs", "Observed SKUs"),
        ("FactUserLicense", "License assignments", "Assignments"),
        ("FactUserLicense", "Users with assignments", "Licensed users"),
        ("DimUser", "License review candidates", "Review candidates"),
    ]):
        add_card(pages, visuals, "licenses", *item, 24 + index * 312, 136, h=88)
    add_ratio_bar(pages, visuals, "licenses", "FactUserLicense", "AssignmentStateLabel", "FactUserLicense", "Assignment state rate", "License assignments", "Assignments by state", 24, 240, w=400, h=184)
    add_ratio_bar(pages, visuals, "licenses", "DimLicenseServicePlan", "ProvisioningStatus", "DimLicenseServicePlan", "Service plan provisioning rate", "License service plans", "Service plans by provisioning status", 440, 240, w=400, h=184)
    add_ratio_bar(pages, visuals, "licenses", "LicenseAssignmentPath", "AssignmentRoute", "LicenseAssignmentPath", "Assignment route rate", "Assignment paths", "Direct vs group assignment", 856, 240, w=400, h=184)
    candidates = add_table(
        pages, visuals, "licenses",
        [
            ("DimLicenseSku", "SkuPartNumberLabel", "SKU", False),
            ("DimLicenseSku", "SKU enabled units", "Enabled", True),
            ("DimLicenseSku", "SKU consumed units", "Consumed", True),
            ("DimLicenseSku", "SKU allocation rate", "Allocation", True),
        ],
        "Capacity by SKU — unaffected by assignment-state selection", 24, 440, 760, 404,
    )
    add_table(
        pages, visuals, "licenses",
        [
            ("DimUser", "DisplayNameLabel", "User", False),
            ("DimUser", "CountryLabel", "Country", False),
            ("DimUser", "AccountStatusLabel", "Account", False),
            ("DimUser", "LastSuccessfulSignInDateTime", "Last successful sign-in", False),
            ("DimUser", "License review band", "Review band", False),
            ("LicenseAssignmentPath", "Product", "Product", False),
            ("LicenseAssignmentPath", "AssignmentRoute", "Route", False),
        ],
        "License review candidates — evidence for review, not proof of non-use", 800, 440, 456, 404,
    )
    set_categorical_values_filter(
        candidates, "DimUser", "License review band",
        [
            "Priority · disabled account",
            "Priority · no M365 activity in D180",
            "Priority · workload over 90 days",
            "Watch · workload 31–90 days",
        ],
        "review",
    )
    return page, visuals


def build_business_services(pages):
    page, visuals = new_page(
        pages, "businessservices", "Services & Impact",
        "Trace exact collaboration, application, device and licensing evidence. Business-service ownership and criticality remain explicitly unavailable.",
    )
    compact_page_header(visuals, 712)
    add_slicer(pages, visuals, "businessservices", "DimUser", "CountryLabel", "Country", 752, y=14, w=160, h=80)
    add_slicer(pages, visuals, "businessservices", "FactRelationshipOverview", "RelationshipType", "Relationship type", 920, y=14, w=160, h=80)
    add_slicer(pages, visuals, "businessservices", "LicenseAssignmentPath", "AssignmentRoute", "Assignment route", 1088, y=14, w=168, h=80)
    for index, item in enumerate([
        ("DimSharePointSite", "SharePoint sites", "SharePoint sites"),
        ("DimTeam", "Teams", "Teams"),
        ("FactRelationshipOverview", "Relationship edges", "Exact relationships"),
        ("DimTeam", "Average members per Team", "Avg members / Team"),
    ]):
        add_card(pages, visuals, "businessservices", *item, 24 + index * 312, 136, h=88)
    add_ratio_bar(pages, visuals, "businessservices", "DimSharePointSite", "ActivityState", "DimSharePointSite", "SharePoint activity state rate", "SharePoint sites", "SharePoint activity — site count and rate", 24, 240, w=400, h=184)
    add_ratio_bar(pages, visuals, "businessservices", "DimTeam", "ActivityState", "DimTeam", "Teams activity state rate", "Teams", "Teams activity — Team count and rate", 440, 240, w=400, h=184)
    add_ratio_bar(pages, visuals, "businessservices", "FactRelationshipOverview", "RelationshipType", "FactRelationshipOverview", "Relationship type rate", "Relationship edges", "Exact relationship types", 856, 240, w=400, h=184)
    add_table(
        pages, visuals, "businessservices",
        [
            ("LicenseAssignmentPath", "Account", "Account", False),
            ("LicenseAssignmentPath", "Product", "Product", False),
            ("LicenseAssignmentPath", "AssignmentRoute", "Route", False),
            ("LicenseAssignmentPath", "GroupName", "Group", False),
            ("LicenseAssignmentPath", "ErrorStatus", "Error status", False),
        ],
        "Observed license paths — exact collaboration and application edges are summarized above", 24, 440, 1232, 320,
    )
    add_text(
        pages, visuals, "businessservices",
        "Exact modeled edges now include PrimaryUser, HasMailbox, AssignedLicense, Teams MemberOfGroup, DeviceHasApplication and DeviceInAutopilot. SharePoint member counts are deliberately not inferred from site-usage reports.",
        24, 776, 1232, 68,
    )
    return page, visuals


def build_fleet_hardware(pages):
    page, visuals = new_page(
        pages, "devices", "Fleet, Hardware & Apps",
        "Explore reconciled devices, hardware and exact Intune device-application relations. Product counts collapse versions only for presentation; every source version remains available in detail.",
    )
    compact_page_header(visuals, 560)
    add_slicer(pages, visuals, "devices", "DimDevice", "CountryLabel", "Country", 600, y=14, w=152, h=80)
    add_slicer(pages, visuals, "devices", "DimDevice", "OperatingSystemLabel", "Operating system", 760, y=14, w=152, h=80)
    ownership = add_slicer(
        pages, visuals, "devices", "DimDevice", "OwnershipLabel",
        "Device ownership", 920, y=14, w=152, h=80,
    )
    set_categorical_selection(
        ownership, "DimDevice", "OwnershipLabel", DEFAULT_OWNERSHIP,
    )
    add_slicer(pages, visuals, "devices", "DimDevice", "ComplianceStateLabel", "Compliance", 1080, y=14, w=176, h=80)
    for index, item in enumerate([
        ("DimDevice", "Devices", "Workplace devices"),
        ("DimDevice", "Compliant device share", "Device compliance rate"),
        ("DeviceHardware", "Hardware coverage rate", "Hardware coverage"),
        ("FactDeviceApplication", "Installed application products", "Application products"),
    ]):
        add_card(pages, visuals, "devices", *item, 24 + index * 312, 136, h=88)
    add_ratio_bar(pages, visuals, "devices", "DimDevice", "ComplianceStateLabel", "DimDevice", "Compliance state rate", "Devices", "Compliance — rate and device count", 24, 240, w=400, h=184)
    add_ratio_bar(pages, visuals, "devices", "DeviceHardware", "Manufacturer", "DeviceHardware", "Manufacturer record rate", "Hardware records", "Hardware records by manufacturer", 440, 240, w=400, h=184)
    add_ratio_bar(
        pages, visuals, "devices", "TopApplication", "ApplicationProduct",
        "TopApplication", "Top application occurrence rate", "Top application occurrences",
        "Top 5 apps — global collected occurrences", 856, 240, w=400, h=184,
    )
    add_table(
        pages, visuals, "devices",
        FLEET_TABLE_FIELDS,
        "Fleet inventory — right-click a unique device for Device 360", 24, 440, 760, 404,
    )
    add_table(
        pages, visuals, "devices",
        [
            ("DimDetectedApplication", "DisplayName", "Application", False),
            ("DimDetectedApplication", "Version", "Version", False),
            ("DimDetectedApplication", "Publisher", "Publisher", False),
            ("DimDetectedApplication", "Platform", "Platform", False),
            ("FactDeviceApplication", "Devices reporting applications", "Exact devices", True),
            ("FactDeviceApplication", "Application device rate", "Device rate", True),
        ],
        "Application detail — versions retained; device links are exact", 800, 440, 456, 404,
    )
    return page, visuals


def build_people_messaging(pages):
    page, visuals = new_page(
        pages, "users", "People & Messaging",
        "Review accounts, exact M365 workload activity, observed assignments, device relationships and mailboxes together. Workload activity is evidence for license review, not proof of non-use.",
    )
    compact_page_header(visuals, 560)
    add_slicer(pages, visuals, "users", "DimUser", "CountryLabel", "Country", 600, y=14, w=152, h=80)
    add_slicer(pages, visuals, "users", "DimUser", "DepartmentLabel", "Department", 760, y=14, w=152, h=80)
    add_slicer(pages, visuals, "users", "DimUser", "AccountStatusLabel", "Account status", 920, y=14, w=152, h=80)
    add_slicer(pages, visuals, "users", "FactMailboxHosting", "MailboxTypeGroup", "Mailbox type", 1080, y=14, w=176, h=80)
    for index, item in enumerate([
        ("DimUser", "Users", "Users"),
        ("DimUser", "Enabled account share", "Enabled account rate"),
        ("FactMailboxHosting", "Hosted mailboxes", "Mailboxes"),
        ("FactMailboxHosting", "Exchange Online share", "Exchange Online rate"),
    ]):
        add_card(pages, visuals, "users", *item, 24 + index * 312, 136, h=88)
    add_ratio_bar(pages, visuals, "users", "DimUser", "AccountStatusLabel", "DimUser", "Account status share", "Users", "Enabled vs disabled users", 24, 240, w=296, h=184)
    add_ratio_bar(pages, visuals, "users", "DimUser", "ActivityState", "DimUser", "Activity state rate", "Users", "Observed sign-in activity", 336, 240, w=296, h=184)
    add_ratio_bar(pages, visuals, "users", "FactMailboxHosting", "MailboxTypeGroup", "FactMailboxHosting", "Mailbox type share", "Hosted mailboxes", "User vs shared vs other mailboxes", 648, 240, w=296, h=184)
    add_ratio_bar(pages, visuals, "users", "FactMailboxHosting", "HostingLocation", "FactMailboxHosting", "Mailbox hosting rate", "Hosted mailboxes", "Exchange Online vs on-premises", 960, 240, w=296, h=184)
    add_table(
        pages, visuals, "users",
        [
            ("DimUser", "UserSelection", "Open User 360", False),
            ("DimUser", "DisplayNameLabel", "Name", False),
            ("DimUser", "UserPrincipalNameLabel", "Account", False),
            ("DimUser", "AccountStatusLabel", "Status", False),
            ("FactUserLicense", "License assignments", "Assignments", True),
        ],
        "Filtered users — personal data, private use", 24, 440, 608, 404,
    )
    add_table(
        pages, visuals, "users",
        [
            ("FactMailboxHosting", "RecipientTypeDetails", "Recipient type", False),
            ("FactMailboxHosting", "MailboxTypeGroup", "Mailbox family", False),
            ("FactMailboxHosting", "HostingLocation", "Hosting", False),
            ("FactMailboxHosting", "CountryLabel", "Country", False),
            ("FactMailboxHosting", "EvidenceSource", "Evidence", False),
        ],
        "Reconciled mailbox hosting — addresses remain hidden", 648, 440, 608, 404,
    )
    return page, visuals


def update_licensing_service_plans(pages: Path):
    """Use the second comparison chart for the newly collected service plans."""
    visual_path = pages / "licenses" / "visuals" / "licensesv11" / "visual.json"
    if not visual_path.is_file():
        raise ValueError("Expected the reviewed licensing comparison visual")
    visual = load(visual_path)
    set_bar(
        visual,
        "DimLicenseServicePlan",
        "ProvisioningStatus",
        "DimLicenseServicePlan",
        "License service plans",
        "License service plans by provisioning status",
    )
    write(visual_path, visual)


def build_device_360(pages):
    page, visuals = new_page(
        pages, "device360", "Device 360",
        "Select one device. Identity, enrollment, activity and hardware retain their own source evidence and dates.",
    )
    page["filterConfig"] = copy.deepcopy(load(pages / "device360" / "page.json").get("filterConfig", {}))
    page["pageBinding"] = copy.deepcopy(load(pages / "device360" / "page.json").get("pageBinding", {}))
    compact_detail_header(visuals, 720)
    selector = add_slicer(
        pages, visuals, "device360", "DimDevice", "DeviceSelection",
        "Search and select one device", 776, y=14, w=296, h=80,
    )
    configure_searchable_slicer(selector)
    country = add_slicer(
        pages, visuals, "device360", "DimDevice", "CountryLabel",
        "Country", 1088, y=14, w=168, h=80,
    )
    configure_searchable_slicer(country)
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
        "Identity and association", 24, 144, 1232, 112,
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
        "Enrollment and source identity", 24, 268, 608, 184,
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
        "Hardware equipment", 648, 268, 608, 184,
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
        "Activity evidence — ambiguous dates remain unqualified", 24, 464, 608, 184,
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
        "Hardware source evidence", 648, 464, 608, 184,
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
        "Findings linked to this device", 24, 660, 1232, 184,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    decorate_detail_sections(pages, visuals, "device360")
    return page, visuals


def build_user_360(pages):
    page, visuals = new_page(
        pages, "user360", "User 360",
        "Select one account. Paths explain assigned licenses, not usage. Activity uses collected sign-in timestamps; manager evidence is not collected.",
    )
    current = load(pages / "user360" / "page.json")
    page["filterConfig"] = copy.deepcopy(current.get("filterConfig", {}))
    page["pageBinding"] = copy.deepcopy(current.get("pageBinding", {}))
    compact_detail_header(visuals, 720)
    selector = add_slicer(
        pages, visuals, "user360", "DimUser", "UserSelection",
        "Search and select one user", 776, y=14, w=296, h=80,
    )
    configure_searchable_slicer(selector)
    country = add_slicer(
        pages, visuals, "user360", "DimUser", "CountryLabel",
        "Country", 1088, y=14, w=168, h=80,
    )
    configure_searchable_slicer(country)
    add_table(
        pages, visuals, "user360",
        [
            ("DimUser", "DisplayNameLabel", "Name", False),
            ("DimUser", "UserPrincipalNameLabel", "Account", False),
            ("DimUser", "JobTitle", "Job title", False),
            ("DimUser", "DepartmentLabel", "Department", False),
            ("DimUser", "AccountStatusLabel", "Status", False),
            ("DimUser", "CreationRaw", "Creation — original text", False),
            ("DimUser", "CreationStatus", "Date qualification", False),
            ("DimUser", "SourceCollectedDateTime", "Collected UTC", False),
            ("DimUser", "User details", "Detail rows", True),
        ],
        "Account profile", 24, 144, 1232, 150,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "user360",
        [
            ("DimDevice", "DeviceNameLabel", "Device", False),
            ("DimDevice", "DeviceSelection", "Open Device 360", False),
            ("FactUserDeviceRelationship", "DeviceCompliance", "Compliance", False),
            ("FactUserDeviceRelationship", "DeviceSyncStatus", "Sync qualification", False),
            ("FactUserDeviceRelationship", "User device details", "Detail rows", True),
        ],
        "Resolved device associations", 24, 306, 608, 170,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][1]["hidden"] = True
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "user360",
        [
            ("FactMailbox", "PrimarySmtpAddressLabel", "Mailbox", False),
            ("FactMailbox", "RecipientTypeDetailsLabel", "Type", False),
            ("FactMailbox", "ArchiveStatusLabel", "Archive", False),
            ("FactMailbox", "User mailbox details", "Detail rows", True),
        ],
        "Associated mailboxes", 648, 306, 608, 170,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "user360",
        [
            ("LicenseAssignmentPath", "Product", "Product", False),
            ("LicenseAssignmentPath", "AssignmentRoute", "Route", False),
            ("LicenseAssignmentPath", "GroupName", "Group", False),
            ("DimGroup", "GroupSelection", "Open Group 360", False),
            ("LicenseAssignmentPath", "AssignmentState", "State", False),
            ("LicenseAssignmentPath", "AssignmentError", "Reported error", False),
            ("LicenseAssignmentPath", "AssignmentUpdatedUtcDateTime", "Last updated UTC", False),
            ("LicenseAssignmentPath", "DisabledPlanIds", "Disabled plan IDs", False),
            ("LicenseAssignmentPath", "TenantAssignmentPathKey", "Path identity", False),
            ("LicenseAssignmentPath", "User assignment details", "Detail rows", True),
        ],
        "License paths — every route retained; last updated is not initial assignment date", 24, 488, 1232, 202,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][3]["hidden"] = True
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-2]["hidden"] = True
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "user360",
        [
            ("EntityFinding", "Severity", "Severity", False),
            ("EntityFinding", "FindingType", "Finding", False),
            ("EntityFinding", "Description", "Evidence", False),
            ("EntityFinding", "RecommendedAction", "Suggested review", False),
            ("EntityFinding", "User finding details", "Detail rows", True),
        ],
        "Findings directly linked to this account", 24, 702, 1232, 142,
    )
    visuals[-1]["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    decorate_detail_sections(pages, visuals, "user360")
    return page, visuals


def build_group_360(pages):
    page, visuals = new_page(
        pages, "group360", "Group 360",
        "Select one group. License paths are not membership. No observed path does not mean unused or safe to delete.",
    )
    current = load(pages / "group360" / "page.json")
    page["filterConfig"] = copy.deepcopy(current.get("filterConfig", {}))
    page["pageBinding"] = copy.deepcopy(current.get("pageBinding", {}))
    compact_detail_header(visuals, 584)
    selector = add_slicer(
        pages, visuals, "group360", "DimGroup", "GroupSelection",
        "Search and select one group", 624, y=14, w=240, h=80,
    )
    configure_searchable_slicer(selector)
    coverage = add_slicer(
        pages, visuals, "group360", "DimGroup", "ObservedPathStatus",
        "License path coverage", 880, y=14, w=200, h=80,
    )
    configure_searchable_slicer(coverage)
    country = add_slicer(
        pages, visuals, "group360", "DimUser", "CountryLabel",
        "User country", 1096, y=14, w=160, h=80,
    )
    configure_searchable_slicer(country)
    add_table(
        pages, visuals, "group360",
        [
            ("DimGroup", "DisplayName", "Group", False),
            ("DimGroup", "MailEnabled", "Mail enabled", False),
            ("DimGroup", "SecurityEnabled", "Security enabled", False),
            ("DimGroup", "GroupTypes", "Type flags", False),
            ("DimGroup", "MemberStatus", "Members", False),
            ("DimGroup", "OwnerStatus", "Owners", False),
            ("DimGroup", "ObservedPathCount", "Observed paths", False),
            ("DimGroup", "SourceCollectedDateTime", "Collected UTC", False),
            ("DimGroup", "Group details", "Detail rows", True),
        ],
        "Group profile", 24, 144, 1232, 144,
    )
    profile = visuals[-1]
    profile["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "group360",
        [
            ("DimUser", "DisplayNameLabel", "Name", False),
            ("DimUser", "UserPrincipalNameLabel", "Account", False),
            ("DimUser", "UserSelection", "Open User 360", False),
            ("LicenseAssignmentPath", "Product", "Product", False),
            ("LicenseAssignmentPath", "AssignmentState", "State", False),
            ("LicenseAssignmentPath", "AssignmentError", "Reported error", False),
            ("LicenseAssignmentPath", "AssignmentUpdatedUtcDateTime", "Last updated UTC", False),
            ("LicenseAssignmentPath", "DisabledPlanIds", "Disabled plan IDs", False),
            ("LicenseAssignmentPath", "TenantAssignmentPathKey", "Path identity", False),
            ("LicenseAssignmentPath", "Group assignment details", "Detail rows", True),
        ],
        "Observed license assignment paths from this group", 24, 300, 1232, 356,
    )
    paths = visuals[-1]
    paths["visual"]["query"]["queryState"]["Values"]["projections"][2]["hidden"] = True
    paths["visual"]["query"]["queryState"]["Values"]["projections"][-2]["hidden"] = True
    paths["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    add_table(
        pages, visuals, "group360",
        [
            ("EntityFinding", "Severity", "Severity", False),
            ("EntityFinding", "FindingType", "Finding", False),
            ("EntityFinding", "Description", "Evidence", False),
            ("EntityFinding", "RecommendedAction", "Suggested review", False),
            ("EntityFinding", "Group finding details", "Detail rows", True),
        ],
        "Findings directly linked to this group", 24, 668, 1232, 176,
    )
    findings = visuals[-1]
    findings["visual"]["query"]["queryState"]["Values"]["projections"][-1]["hidden"] = True
    page["visualInteractions"] = [
        {"source": country["name"], "target": selector["name"], "type": "NoFilter"},
        {"source": country["name"], "target": coverage["name"], "type": "NoFilter"},
        {"source": country["name"], "target": profile["name"], "type": "NoFilter"},
        {"source": country["name"], "target": findings["name"], "type": "NoFilter"},
    ]
    decorate_detail_sections(pages, visuals, "group360")
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
        ("devicesv0", "Smart Workplace CMDB — Executive Overview", 76, 14, 628, 42),
        ("devicesv1", "VERSION 1.0.0", 76, 60, 628, 24),
        ("devicesv2", "Workforce, workplace devices, messaging and Microsoft 365 licensing — one shared country context.", 76, 88, 628, 34),
        ("devicesv3", "V1 · Frozen private snapshot · Selections filter related visuals", 24, 864, 1232, 28),
    ]:
        visual = clone_visual(pages, "devices", template)
        set_text(visual, value)
        put(visuals, visual, "overview", x, y, w, h)

    endpoint_score = add_card(
        pages, visuals, "overview", "DimCountry", "Executive Endpoint Analytics score",
        ENDPOINT_ANALYTICS_CARD_TITLE, 720, 14, 216, 80,
    )
    endpoint_score["visual"]["objects"]["value"][0]["properties"]["labelPrecision"] = lit(1)

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
    set_categorical_selection(
        ownership, "DimDevice", "OwnershipLabel", DEFAULT_OWNERSHIP,
    )

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
        ("FactADIntuneCoverage", "AD to Intune coverage rate", "AD → Intune coverage"),
        ("DimCountry", "Executive compliant device share", "Compliance rate"),
        ("DimUser", "Users", "Users"),
        ("FactMailboxHosting", "Hosted mailboxes", "Mailboxes"),
    ]
    for index, item in enumerate(kpis):
        add_card(pages, visuals, "overview", *item, 24 + index * 208, 136, 192, 96)
    add_quality_summary_card(pages, visuals, "overview", 1064, 136, 192, 96)

    form_factor = add_ratio_bar(
        pages, visuals, "overview", "DimDevice", "Device form factor", "DimCountry",
        "Executive device form factor share", "Executive workplace devices",
        "PC vs mobile devices", 24, 240,
    )
    ownership_ratio = add_ratio_bar(
        pages, visuals, "overview", "DimDevice", "Device ownership group", "DimCountry",
        "Executive device ownership share", "Executive workplace devices",
        "Corporate vs personal", 272, 240,
    )
    windows = add_ratio_bar(
        pages, visuals, "overview", "DimDevice", "Windows version group", "DimCountry",
        "Executive Windows version share", "Executive workplace devices",
        "Windows 11 adoption", 520, 240, h=136,
    )
    set_single_value_filter(windows, "DimDevice", "OperatingSystem", "Windows", "windows")
    add_compact_card(
        pages, visuals, "overview", "DimCountry", "Executive Windows 10 devices",
        "Windows 10", 520, 382, 116, 58,
    )
    add_compact_card(
        pages, visuals, "overview", "DimCountry", "Executive Windows 10 not compatible devices",
        "Not capable", 642, 382, 118, 58,
    )
    accounts = add_ratio_bar(
        pages, visuals, "overview", "DimUser", "AccountStatusLabel", "DimUser",
        "Account status share", "Users", "Enabled vs disabled users", 768, 240,
    )
    mailboxes = add_ratio_bar(
        pages, visuals, "overview", "FactMailboxHosting", "MailboxTypeGroup", "FactMailboxHosting",
        "Mailbox type share", "Hosted mailboxes", "Mailbox types", 1016, 240,
    )
    set_category_colors(
        form_factor, "DimDevice", "Device form factor",
        EXECUTIVE_RATIO_CATEGORY_COLORS["form_factor"],
    )
    set_category_colors(
        ownership_ratio, "DimDevice", "Device ownership group",
        EXECUTIVE_RATIO_CATEGORY_COLORS["ownership"],
    )
    set_category_colors(
        windows, "DimDevice", "Windows version group",
        EXECUTIVE_RATIO_CATEGORY_COLORS["windows"],
    )
    set_category_colors(
        accounts, "DimUser", "AccountStatusLabel",
        EXECUTIVE_RATIO_CATEGORY_COLORS["accounts"],
    )
    set_category_colors(
        mailboxes, "FactMailboxHosting", "MailboxTypeGroup",
        EXECUTIVE_RATIO_CATEGORY_COLORS["mailbox_types"],
    )
    for key, x in [
        ("formfactor", 24),
        ("ownership", 272),
        ("windows", 520),
        ("accounts", 768),
        ("mailboxes", 1016),
    ]:
        add_overview_ratio_icon(pages, visuals, key, x, 240)

    country_bar = add_country_bar(pages, visuals, "overview", 24, 448, 608, 280)
    license_country_bar = add_license_country_bar(pages, visuals, "overview", 648, 448, 608, 280)

    for index, (title, sku) in enumerate(LICENSE_SUMMARY_SKUS):
        add_license_summary_card(pages, visuals, "overview", title, sku, 24 + index * 248, 736)

    page["visualInteractions"] = [
        {"source": ownership["name"], "target": ownership_ratio["name"], "type": "NoFilter"},
        {"source": country["name"], "target": country_bar["name"], "type": "NoFilter"},
        {"source": country["name"], "target": license_country_bar["name"], "type": "NoFilter"},
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


def prepare(
    report: Path,
    exchange_onprem_local: Path | None = None,
    exchange_onprem_remote: Path | None = None,
    upgrade_eligibility: Path | None = None,
):
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
    mailbox_metadata = enrich_semantic_model(
        report,
        exchange_onprem_local,
        exchange_onprem_remote,
        upgrade_eligibility,
    )
    for builder in [build_risk, build_lifecycle, build_licensing, build_fleet_hardware,
                    build_people_messaging, build_business_services,
                    build_device_360, build_user_360, build_group_360]:
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
        if page_id not in {"overview", "devices"}:
            set_page_categorical_filter(
                page,
                f"{page_id}{DEFAULT_OWNERSHIP_FILTER_NAME}",
                "DimDevice",
                "OwnershipLabel",
                DEFAULT_OWNERSHIP,
            )
        write(page_path, page)
        if page_id in HEADER_NAMES:
            visual_path = pages / page_id / "visuals" / f"{page_id}v0" / "visual.json"
            if visual_path.is_file():
                visual = load(visual_path)
                set_text(visual, f"Smart Workplace CMDB — {HEADER_NAMES[page_id]}")
                right = visual["position"]["x"] + visual["position"]["width"]
                visual["position"]["x"] = 76
                visual["position"]["width"] = max(1, right - 76)
                # Keep the title above slicer/background layers. Power BI can
                # otherwise mask the leading text on compact 360 headers even
                # though the complete text run remains present in PBIR.
                visual["position"]["z"] = 1002
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
            footer["position"]["width"] = 496
            write(footer_path, footer)
        position_version_in_footer(pages, page_id)
        persist_page_header_icon(pages, page_id, PAGE_ICON_RESOURCES[page_id])
        decorate_page_kpis(pages, page_id)
        add_copyright(pages, page_id)

    user_context_path = pages / "user360" / "visuals" / "user360v2" / "visual.json"
    if user_context_path.is_file():
        user_context = load(user_context_path)
        set_text(
            user_context,
            "Select one account. Paths explain assigned licenses, not usage. "
            "Activity uses collected sign-in timestamps; manager evidence is not collected.",
        )
        write(user_context_path, user_context)

    metadata_path = pages / "pages.json"
    metadata = load(metadata_path)
    metadata["pageOrder"] = PAGE_ORDER
    metadata["activePageName"] = "overview"
    write(metadata_path, metadata)

    report_path = report / "definition" / "report.json"
    resource_names = sorted(
        set(PAGE_ICON_RESOURCES.values())
        | set(OVERVIEW_RATIO_ICON_RESOURCES.values())
        | set(KPI_ICON_RESOURCES.values())
        | set(SECTION_ICON_RESOURCES.values())
    )
    resource_root = report / "StaticResources" / "RegisteredResources"
    resource_root.mkdir(parents=True, exist_ok=True)
    for resource_name in resource_names:
        icon_source = Path(__file__).resolve().parent / "Assets" / resource_name
        if not icon_source.is_file():
            raise ValueError(f"Page icon asset not found: {icon_source}")
        shutil.copy2(icon_source, resource_root / resource_name)
    report_json = load(report_path)
    packages = report_json.setdefault("resourcePackages", [])
    registered = next((package for package in packages if package.get("name") == "RegisteredResources"), None)
    if registered is None:
        registered = {"name": "RegisteredResources", "type": "RegisteredResources", "items": []}
        packages.append(registered)
    items = registered.setdefault("items", [])
    items[:] = [item for item in items if item.get("name") not in resource_names]
    items.extend({"name": name, "path": name, "type": "Image"} for name in resource_names)
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
    parser.add_argument("--upgrade-eligibility", type=Path)
    args = parser.parse_args()
    print(json.dumps(prepare(
        args.report,
        args.exchange_onprem_local,
        args.exchange_onprem_remote,
        args.upgrade_eligibility,
    ), ensure_ascii=False))
