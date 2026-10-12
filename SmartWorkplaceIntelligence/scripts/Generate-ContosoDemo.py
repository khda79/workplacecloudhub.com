#!/usr/bin/env python3
# Version: 1.2.0
"""Generate and validate an isolated, deterministic Contoso demonstration dataset."""

from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import json
import math
import random
import re
import uuid
from collections import Counter, defaultdict
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import urlparse


PRODUCT = Path(__file__).resolve().parents[1]
CONTRACT = PRODUCT / "config" / "prepared-evidence-contract.json.txt"
FIXTURE_HEADERS = PRODUCT / "fixtures" / "valid"
MODEL_TABLES = PRODUCT / "pbip" / "SmartWorkplaceIntelligence.SemanticModel" / "definition" / "tables"
TENANT = "contoso-demo"
TENANT_ID = "00000000-0000-4000-8000-00000000c026"
COUNTRIES = [("France", 42, 8), ("United Kingdom", 20, 4), ("Germany", 15, 3),
             ("Spain", 10, 2), ("United States", 8, 2), ("Netherlands", 5, 1)]
DEPARTMENTS = [("Operations", 25), ("Sales", 20), ("Technology", 15),
               ("Customer Support", 12), ("Corporate Services", 10),
               ("Field Services", 8), ("Finance", 5), ("People", 5)]
FIRST = ["Aven", "Bela", "Cerin", "Dara", "Elian", "Fara", "Galen", "Hana",
         "Ilan", "Jora", "Kian", "Lena", "Maren", "Nila", "Oren", "Pera",
         "Quin", "Rina", "Soren", "Talia", "Ulan", "Vera", "Wilan", "Yara"]
LAST = ["Aldren", "Bellan", "Corven", "Darian", "Esmar", "Ferran", "Galenor",
        "Havren", "Ivaran", "Jorren", "Kellan", "Lorian", "Mavren", "Nordan",
        "Orlian", "Pevran", "Quellan", "Ravern", "Seloran", "Tavren", "Uldran",
        "Valoren", "Wendar", "Yeloran"]
LAST_END = ["berg", "brook", "den", "field", "ford", "holm", "ley", "lin",
            "mere", "mont", "more", "nor", "quist", "ridge", "sen", "son",
            "stead", "ton", "vale", "ward", "well", "wick", "wood", "wyn"]
OS_COUNTS = [("Windows 11", "windows11"), ("Windows 10", "windows10"),
             ("macOS", "mac"), ("Android", "android"), ("iOS", "iphone")]
LICENSES = [("LIC-001", "Microsoft 365 E3", "M365_E3", "Workplace Suite"),
            ("LIC-002", "Microsoft 365 F3", "M365_F3", "Frontline Suite"),
            ("LIC-003", "Microsoft 365 E5", "M365_E5", "Workplace Suite"),
            ("LIC-004", "Office 365 E3", "O365_E3", "Microsoft 365"),
            ("LIC-005", "Microsoft 365 Copilot", "COPILOT", "Add-on"),
            ("LIC-006", "Power BI Pro", "POWER_BI_PRO", "Add-on"),
            ("LIC-007", "Teams Phone", "TEAMS_PHONE", "Add-on"),
            ("LIC-008", "Project Plan 3", "PROJECT_P3", "Add-on")]

SP_COMMON = ["TenantKey", "FarmId", "RunId", "CollectedAtUtc"]
SP_SCHEMAS = {
    "Farms": ["FarmVersion", "BuildVersion", "ConfigurationDatabase", "ServerCount", "ServiceApplicationCount", "CollectionStatus"],
    "Servers": ["ServerId", "ServerName", "Role", "Status", "CollectionStatus"],
    "ServiceApplications": ["ServiceApplicationId", "Name", "TypeName", "Status", "ApplicationPool", "CollectionStatus"],
    "WebApplications": ["WebApplicationId", "Name", "DefaultUrl", "ApplicationPool", "ContentDatabaseCount", "CollectionStatus"],
    "WebApplicationZones": ["WebApplicationId", "Zone", "PublicUrl", "AuthenticationMode", "ClaimsAuthentication", "AuthenticationProviders", "CollectionStatus"],
    "ContentDatabases": ["WebApplicationId", "ContentDatabaseId", "Name", "DatabaseServer", "Status", "SiteCollectionCount", "IsReadOnly", "CollectionStatus", "DatabaseSizeBytes", "DatabaseSizeStatus", "NeedsUpgradeIncludeChildren", "NeedsUpgradeStatus"],
    "SiteCollections": ["WebApplicationId", "ContentDatabaseId", "SiteCollectionId", "Url", "OwnerLogin", "SecondaryOwnerLogin", "StorageBytes", "QuotaLimitBytes", "QuotaUsagePercent", "QuotaUsageStatus", "LastContentModifiedUtc", "RootWebTemplate", "SiteCategory", "SiteCategoryStatus", "LanguageLcid", "LockState", "IsReadOnly", "ReadLocked", "WriteLocked", "LockIssue", "ContentDatabaseIsReadOnly", "LockStatus", "CollectionStatus"],
    "SiteAdministrators": ["WebApplicationId", "ContentDatabaseId", "SiteCollectionId", "AdministratorLogin", "IsPrimary", "IsSecondary"],
    "Webs": ["WebApplicationId", "ContentDatabaseId", "SiteCollectionId", "WebId", "Url", "Title", "WebTemplate", "LanguageLcid", "ListCount", "LibraryCount", "HasUniqueRoleAssignments"],
    "CollectionCoverage": ["WebApplicationId", "ContentDatabaseId", "SiteCollectionId", "Url", "Status", "WebsCollected", "ErrorStage", "ErrorMessage"],
}
SP_INFRA = ("Farms", "Servers", "ServiceApplications", "WebApplications", "WebApplicationZones", "ContentDatabases")
SP_CONTENT = ("SiteCollections", "SiteAdministrators", "Webs", "CollectionCoverage")


def stable_hex(*parts: object, length: int = 16) -> str:
    return hashlib.sha256("|".join(map(str, parts)).encode("utf-8")).hexdigest()[:length]


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def iso_day(value: date) -> str:
    return value.isoformat()


def iso_time(value: date, hour: int = 8) -> str:
    return f"{value.isoformat()}T{hour:02d}:00:00Z"


def apportioned(total: int, weights: list[int]) -> list[int]:
    if total < 0 or sum(weights) <= 0:
        raise ValueError("Invalid total or weights")
    raw = [total * weight / sum(weights) for weight in weights]
    result = [math.floor(value) for value in raw]
    for index in sorted(range(len(raw)), key=lambda i: (raw[i] - result[i], -i), reverse=True)[:total - sum(result)]:
        result[index] += 1
    return result


def weighted_labels(total: int, labels: list[str], weights: list[int], rng: random.Random) -> list[str]:
    values = [label for label, count in zip(labels, apportioned(total, weights)) for _ in range(count)]
    rng.shuffle(values)
    return values


def luhn_imei(seed: int, ordinal: int) -> str:
    stem = "35" + f"{int(stable_hex(seed, ordinal, 'imei', length=12), 16) % 10**12:012d}"
    digits = [int(digit) for digit in stem]
    score = sum((digit if index % 2 == 0 else (digit * 2 // 10 + digit * 2 % 10))
                for index, digit in enumerate(digits))
    return stem + str((-score) % 10)


def write_csv(path: Path, headers: list[str], rows: list[dict] | tuple[dict, ...]) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=headers, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in headers})
            count += 1
    return count


def write_json(path: Path, data: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False, sort_keys=True) + "\n", encoding="utf-8")


def write_sharepoint_csv(path: Path, columns: list[str], rows: list[dict]) -> None:
    """Match the on-prem collectors: UTF-8 BOM, semicolons, quoted fields and CRLF."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=columns, extrasaction="raise",
                                delimiter=";", quoting=csv.QUOTE_ALL, lineterminator="\r\n")
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in columns})


def sp_guid(seed: int, kind: str, ordinal: object) -> str:
    return str(uuid.uuid5(uuid.NAMESPACE_DNS, f"contoso.com:{seed}:{kind}:{ordinal}"))


def sharepoint_config(args: argparse.Namespace) -> dict[str, int]:
    return {"MemberServers": args.sp_member_servers, "ExternalServers": args.sp_external_servers,
            "ServiceApplications": args.sp_service_applications, "WebApplications": args.sp_web_applications,
            "ContentDatabases": args.sp_content_databases, "PublishedCollections": args.sp_published_collections,
            "ExcludedPersonal": args.sp_excluded_personal, "MetadataIncomplete": args.sp_metadata_incomplete,
            "DatabaseCoverageMismatch": args.sp_database_mismatches,
            "WebsPerCollection": args.sp_webs_per_collection,
            "AdministratorsPerCollection": args.sp_administrators_per_collection}


def write_sharepoint_onprem(root: Path, people: list[dict], as_of: date, seed: int,
                            config: dict[str, int]) -> dict:
    if config["WebApplications"] < 2 or config["ContentDatabases"] < config["WebApplications"]:
        raise ValueError("SharePoint requires at least two web applications and one database per web application")
    if config["PublishedCollections"] < config["MetadataIncomplete"] or config["PublishedCollections"] < config["WebApplications"]:
        raise ValueError("Invalid published collection volume")
    if config["DatabaseCoverageMismatch"] > config["ContentDatabases"] or config["WebsPerCollection"] < 1 or config["AdministratorsPerCollection"] < 1:
        raise ValueError("Invalid SharePoint relationship volume")
    if any(value < 0 for value in config.values()) or not people:
        raise ValueError("SharePoint volumes must be nonnegative and users must exist")
    rng = random.Random(seed ^ 0x5A17C026)
    infra_run = uuid.uuid5(uuid.NAMESPACE_DNS, f"contoso.com:{seed}:sp-infra:{as_of}").hex
    content_run = uuid.uuid5(uuid.NAMESPACE_DNS, f"contoso.com:{seed}:sp-content:{as_of}").hex
    farm_id = sp_guid(seed, "farm", 1)
    common = {"TenantKey": TENANT, "FarmId": farm_id}
    infra = {**common, "RunId": infra_run, "CollectedAtUtc": iso_time(as_of, 8)}
    content = {**common, "RunId": content_run, "CollectedAtUtc": iso_time(as_of, 10)}
    rows: dict[str, list[dict]] = {name: [] for name in SP_SCHEMAS}
    rows["Farms"].append({**infra, "FarmVersion": "SharePoint Server 2019", "BuildVersion": "16.0.10396.20000",
                           "ConfigurationDatabase": "Contoso_SP_Config", "ServerCount": config["MemberServers"],
                           "ServiceApplicationCount": config["ServiceApplications"], "CollectionStatus": "Collected"})
    roles = ["WebFrontEnd", "Application", "ApplicationWithSearch", "DistributedCache"]
    for number in range(config["MemberServers"] + config["ExternalServers"]):
        external = number >= config["MemberServers"]
        rows["Servers"].append({**infra, "ServerId": sp_guid(seed, "server", number),
                                "ServerName": f"sp-{'external' if external else 'member'}-{number + 1:02d}.contoso.com",
                                "Role": "Invalid" if external else roles[number % len(roles)],
                                "Status": "Online" if number % 11 else "Offline", "CollectionStatus": "Collected"})
    service_types = ["Managed Metadata Service", "Search Service Application", "User Profile Service Application",
                     "Secure Store Service", "Business Data Connectivity Service", "App Management Service"]
    for number in range(config["ServiceApplications"]):
        rows["ServiceApplications"].append({**infra, "ServiceApplicationId": sp_guid(seed, "service", number),
                                            "Name": f"Contoso {service_types[number % len(service_types)]} {number + 1:02d}",
                                            "TypeName": service_types[number % len(service_types)], "Status": "Online",
                                            "ApplicationPool": "Contoso Service Pool", "CollectionStatus": "Collected"})
    web_count = config["WebApplications"]
    db_counts = apportioned(config["ContentDatabases"], [6, 5, 5, 4, 5][:web_count] if web_count <= 5 else [1] * web_count)
    published_counts = apportioned(config["PublishedCollections"], [32, 25, 20, 13, 10][:web_count]
                                   if web_count <= 5 else [1] * web_count)
    hosts = ["portal", "collab", "projects", "services"]
    webapps = []
    databases = []
    for app_index in range(web_count):
        host = "mysites" if app_index == web_count - 1 else hosts[app_index] if app_index < len(hosts) else f"workspace{app_index + 1}"
        url = f"https://{host}.contoso.com"
        app_id = sp_guid(seed, "webapp", app_index)
        webapps.append((app_id, url, app_index))
        rows["WebApplications"].append({**infra, "WebApplicationId": app_id,
                                        "Name": f"Contoso {host.title()}", "DefaultUrl": url,
                                        "ApplicationPool": f"Contoso {host.title()} Pool",
                                        "ContentDatabaseCount": db_counts[app_index], "CollectionStatus": "Collected"})
        rows["WebApplicationZones"].append({**infra, "WebApplicationId": app_id, "Zone": "Default",
                                             "PublicUrl": url, "AuthenticationMode": "Claims",
                                             "ClaimsAuthentication": "True", "AuthenticationProviders": "Windows",
                                             "CollectionStatus": "Collected"})
        if app_index < min(4, web_count - 1):
            rows["WebApplicationZones"].append({**infra, "WebApplicationId": app_id, "Zone": "Intranet",
                                                 "PublicUrl": f"https://intranet-{host}.contoso.com",
                                                 "AuthenticationMode": "Claims", "ClaimsAuthentication": "True",
                                                 "AuthenticationProviders": "Windows", "CollectionStatus": "Collected"})
        for db_index in range(db_counts[app_index]):
            db_id = sp_guid(seed, "database", f"{app_index}:{db_index}")
            databases.append((app_index, db_index, app_id, db_id))
            rows["ContentDatabases"].append({**infra, "WebApplicationId": app_id, "ContentDatabaseId": db_id,
                                             "Name": f"Contoso_SP_Content_{app_index + 1:02d}_{db_index + 1:02d}",
                                             "DatabaseServer": f"sql-{1 + (app_index + db_index) % 3:02d}.contoso.com",
                                             "Status": "Online", "SiteCollectionCount": 0, "IsReadOnly": "False",
                                             "CollectionStatus": "Collected", "DatabaseSizeBytes": (12 + db_index * 3) * 1024**3,
                                             "DatabaseSizeStatus": "Collected", "NeedsUpgradeIncludeChildren": "False",
                                             "NeedsUpgradeStatus": "Collected"})
    db_by_app = {index: [(db_id, db_index) for app, db_index, _, db_id in databases if app == index]
                 for index in range(web_count)}
    enabled = [person for person in people if person.get("AccountState") == "Enabled"] or people
    incomplete = set(rng.sample(range(config["PublishedCollections"]), config["MetadataIncomplete"]))
    collection_number = 0
    db_counts_observed = Counter()
    for app_index, (app_id, base_url, _) in enumerate(webapps):
        for local_index in range(published_counts[app_index]):
            db_id, _ = db_by_app[app_index][local_index % len(db_by_app[app_index])]
            db_counts_observed[db_id] += 1
            collection_id = sp_guid(seed, "collection", collection_number)
            path = f"/sites/{'team' if app_index < web_count - 1 else 'community'}-{collection_number + 1:04d}"
            url = base_url + path
            primary = enabled[collection_number % len(enabled)]["UPN"]
            secondary = enabled[(collection_number + 137) % len(enabled)]["UPN"]
            issue = collection_number in incomplete
            issue_kind = collection_number % 3
            storage = (5 + rng.randrange(75)) * 1024**3
            quota = (100 + rng.randrange(300)) * 1024**3
            modified = iso_time(as_of - timedelta(days=rng.randrange(0, 900)), 7)
            site = {**content, "WebApplicationId": app_id, "ContentDatabaseId": db_id,
                    "SiteCollectionId": collection_id, "Url": url, "OwnerLogin": primary,
                    "SecondaryOwnerLogin": secondary, "StorageBytes": storage, "QuotaLimitBytes": quota,
                    "QuotaUsagePercent": round(storage / quota * 100, 2), "QuotaUsageStatus": "Observed",
                    "LastContentModifiedUtc": modified, "RootWebTemplate": "STS#0",
                    "SiteCategory": "OtherOnMyHost" if app_index == web_count - 1 else "Other",
                    "SiteCategoryStatus": "Observed", "LanguageLcid": 1033,
                    "LockState": "Unlock", "IsReadOnly": "False", "ReadLocked": "False", "WriteLocked": "False",
                    "LockIssue": "", "ContentDatabaseIsReadOnly": "False", "LockStatus": "Observed",
                    "CollectionStatus": "MetadataIncomplete" if issue else "Collected"}
            if issue:
                if issue_kind == 0:
                    site["LastContentModifiedUtc"] = ""
                    missing = "LastContentModifiedUtc"
                elif issue_kind == 1:
                    site["QuotaLimitBytes"] = ""
                    site["QuotaUsagePercent"] = ""
                    site["QuotaUsageStatus"] = "QuotaUnavailable"
                    missing = "QuotaLimitBytes"
                else:
                    site["LockStatus"] = "Unverified"
                    site["LockIssue"] = "Synthetic lock metadata unavailable"
                    missing = "LockState"
            rows["SiteCollections"].append(site)
            for admin_index in range(config["AdministratorsPerCollection"]):
                admin = primary if admin_index == 0 else secondary if admin_index == 1 else enabled[(collection_number + 137 + admin_index) % len(enabled)]["UPN"]
                rows["SiteAdministrators"].append({**content, "WebApplicationId": app_id,
                                                   "ContentDatabaseId": db_id, "SiteCollectionId": collection_id,
                                                   "AdministratorLogin": admin, "IsPrimary": str(admin_index == 0),
                                                   "IsSecondary": str(admin_index == 1)})
            for web_index in range(config["WebsPerCollection"]):
                list_count = 4 + rng.randrange(14)
                rows["Webs"].append({**content, "WebApplicationId": app_id, "ContentDatabaseId": db_id,
                                    "SiteCollectionId": collection_id, "WebId": sp_guid(seed, "web", f"{collection_number}:{web_index}"),
                                    "Url": url if web_index == 0 else f"{url}/subweb-{web_index:02d}",
                                    "Title": f"Contoso Team {collection_number + 1:04d}" if web_index == 0 else f"Workspace {web_index:02d}",
                                    "WebTemplate": "STS#0", "LanguageLcid": 1033,
                                    "ListCount": list_count, "LibraryCount": 2 + rng.randrange(min(8, list_count - 1)),
                                    "HasUniqueRoleAssignments": str(web_index > 0 and collection_number % 7 == 0)})
            rows["CollectionCoverage"].append({**content, "WebApplicationId": app_id,
                                               "ContentDatabaseId": db_id, "SiteCollectionId": collection_id,
                                               "Url": url, "Status": site["CollectionStatus"],
                                               "WebsCollected": config["WebsPerCollection"],
                                               "ErrorStage": "SiteCollectionMetadata" if issue else "",
                                               "ErrorMessage": f"Missing or unverified observations: {missing}" if issue else ""})
            collection_number += 1
    my_app_id, my_url, my_index = webapps[-1]
    for number in range(config["ExcludedPersonal"]):
        db_id, _ = db_by_app[my_index][number % len(db_by_app[my_index])]
        db_counts_observed[db_id] += 1
        rows["CollectionCoverage"].append({**content, "WebApplicationId": my_app_id,
                                           "ContentDatabaseId": db_id,
                                           "SiteCollectionId": sp_guid(seed, "personal", number),
                                           "Url": f"{my_url}/personal/user-{number + 1:05d}",
                                           "Status": "ExcludedPersonal", "WebsCollected": 0,
                                           "ErrorStage": "", "ErrorMessage": ""})
    mismatched = {db_id for _, _, _, db_id in databases[:config["DatabaseCoverageMismatch"]]}
    for db_row in rows["ContentDatabases"]:
        db_id = db_row["ContentDatabaseId"]
        observed = db_counts_observed[db_id]
        expected = observed + 1 if db_id in mismatched else observed
        db_row["SiteCollectionCount"] = expected
        if db_id in mismatched:
            rows["CollectionCoverage"].append({**content, "WebApplicationId": db_row["WebApplicationId"],
                                               "ContentDatabaseId": db_id, "SiteCollectionId": "", "Url": "",
                                               "Status": "DatabaseCoverageMismatch", "WebsCollected": "",
                                               "ErrorStage": "ContentDatabase",
                                               "ErrorMessage": f"DatabaseCoverageMismatch: expected={expected}; observed={observed}"})
    for name, data in rows.items():
        write_sharepoint_csv(root / "DATA-LAST" / f"SharePoint_OnPrem_{name}.csv", SP_COMMON + SP_SCHEMAS[name], data)
    return {"SyntheticOnly": True, "InfrastructureRunId": infra_run, "ContentRunId": content_run,
            "InfrastructureStatus": "Success",
            "ContentStatus": "Partial" if config["MetadataIncomplete"] + config["DatabaseCoverageMismatch"] else "Complete",
            "MetadataIncomplete": config["MetadataIncomplete"],
            "DatabaseCoverageMismatch": config["DatabaseCoverageMismatch"],
            "PublishedCollections": config["PublishedCollections"],
            "ExcludedPersonal": config["ExcludedPersonal"],
            "Files": {f"SharePoint_OnPrem_{name}.csv": len(data) for name, data in rows.items()},
            "Parameters": config}


def validate_sharepoint_onprem(root: Path, people: list[dict], as_of: date,
                               config: dict[str, int]) -> dict:
    tables = {}
    report_root = root / "DATA-LAST"
    for name, columns in SP_SCHEMAS.items():
        path = report_root / f"SharePoint_OnPrem_{name}.csv"
        raw = path.read_bytes()
        if not raw.startswith(b"\xef\xbb\xbf") or b"\r\n" not in raw or raw.replace(b"\r\n", b"").find(b"\n") >= 0:
            raise ValueError(f"SharePoint CSV encoding or line ending mismatch: {name}")
        with path.open(encoding="utf-8-sig", newline="") as stream:
            reader = csv.DictReader(stream, delimiter=";")
            if reader.fieldnames != SP_COMMON + columns:
                raise ValueError(f"SharePoint header mismatch: {name}")
            tables[name] = list(reader)
        if not raw[3:].startswith(b'"TenantKey";'):
            raise ValueError(f"SharePoint quoted delimiter mismatch: {name}")
    farm = tables["Farms"]
    if len(farm) != 1:
        raise ValueError("Expected one synthetic SharePoint farm")
    farm_id = farm[0]["FarmId"]
    infra_run = farm[0]["RunId"]
    content_run = tables["SiteCollections"][0]["RunId"]
    uuid.UUID(farm_id)
    if not re.fullmatch(r"[0-9a-f]{32}", infra_run) or not re.fullmatch(r"[0-9a-f]{32}", content_run) or infra_run == content_run:
        raise ValueError("Invalid or shared SharePoint run identifiers")
    for name, records in tables.items():
        run = infra_run if name in SP_INFRA else content_run
        timestamp = iso_time(as_of, 8 if name in SP_INFRA else 10)
        for record in records:
            if any(record[key] != value for key, value in (("TenantKey", TENANT), ("FarmId", farm_id),
                                                            ("RunId", run), ("CollectedAtUtc", timestamp))):
                raise ValueError(f"SharePoint common columns mismatch: {name}")
        if len({tuple(item.items()) for item in records}) != len(records):
            raise ValueError(f"Duplicate SharePoint row: {name}")
    if len(tables["Servers"]) != config["MemberServers"] + config["ExternalServers"] or len(tables["ServiceApplications"]) != config["ServiceApplications"]:
        raise ValueError("SharePoint infrastructure volume mismatch")
    if sum(row["Role"] != "Invalid" for row in tables["Servers"]) != int(farm[0]["ServerCount"]):
        raise ValueError("Farm ServerCount must exclude external Invalid servers")
    if int(farm[0]["ServiceApplicationCount"]) != len(tables["ServiceApplications"]):
        raise ValueError("Farm service application count mismatch")
    for name, key in (("Servers", "ServerId"), ("ServiceApplications", "ServiceApplicationId")):
        values = [row[key] for row in tables[name]]
        if len(values) != len(set(values)):
            raise ValueError(f"Duplicate SharePoint identity: {name}")
        for value in values:
            uuid.UUID(value)
    if len(tables["WebApplications"]) != config["WebApplications"] or len(tables["ContentDatabases"]) != config["ContentDatabases"]:
        raise ValueError("SharePoint application or database volume mismatch")
    if len(tables["WebApplicationZones"]) != config["WebApplications"] + min(4, config["WebApplications"] - 1):
        raise ValueError("SharePoint zone volume mismatch")
    apps = {row["WebApplicationId"]: row for row in tables["WebApplications"]}
    dbs = {row["ContentDatabaseId"]: row for row in tables["ContentDatabases"]}
    if len(apps) != len(tables["WebApplications"]) or len(dbs) != len(tables["ContentDatabases"]):
        raise ValueError("Duplicate SharePoint application or database key")
    for value in (*apps, *dbs):
        uuid.UUID(value)
    for row in tables["ContentDatabases"]:
        if row["WebApplicationId"] not in apps or int(row["DatabaseSizeBytes"]) < 0:
            raise ValueError("Invalid SharePoint content database")
    for app_id, app in apps.items():
        if int(app["ContentDatabaseCount"]) != sum(row["WebApplicationId"] == app_id for row in dbs.values()):
            raise ValueError("Web application database count mismatch")
        zones = [row for row in tables["WebApplicationZones"] if row["WebApplicationId"] == app_id]
        if len([row for row in zones if row["Zone"] == "Default" and row["PublicUrl"] == app["DefaultUrl"]]) != 1:
            raise ValueError("Missing default public zone")
    upns = {person["UPN"] for person in people}
    published = {row["SiteCollectionId"]: row for row in tables["SiteCollections"]}
    if len(published) != config["PublishedCollections"] or len(published) != len(tables["SiteCollections"]):
        raise ValueError("Published collection key or volume mismatch")
    for site_id in published:
        uuid.UUID(site_id)
    zone_keys = [(row["WebApplicationId"], row["Zone"]) for row in tables["WebApplicationZones"]]
    if len(zone_keys) != len(set(zone_keys)) or any(app_id not in apps for app_id, _ in zone_keys):
        raise ValueError("Invalid SharePoint zone keys")
    coverage = tables["CollectionCoverage"]
    status = Counter(row["Status"] for row in coverage)
    expected_status = {"Collected": config["PublishedCollections"] - config["MetadataIncomplete"],
                       "MetadataIncomplete": config["MetadataIncomplete"],
                       "ExcludedPersonal": config["ExcludedPersonal"],
                       "DatabaseCoverageMismatch": config["DatabaseCoverageMismatch"]}
    if status != +Counter(expected_status):
        raise ValueError("SharePoint coverage status mismatch")
    coverage_by_site = {row["SiteCollectionId"]: row for row in coverage if row["Status"] in ("Collected", "MetadataIncomplete")}
    if len(coverage_by_site) != len(published) or set(coverage_by_site) != set(published):
        raise ValueError("Published collection coverage must be one-to-one")
    personal_ids = [row["SiteCollectionId"] for row in coverage if row["Status"] == "ExcludedPersonal"]
    if len(personal_ids) != len(set(personal_ids)) or set(personal_ids) & set(published):
        raise ValueError("Duplicate or published personal collection identity")
    for site_id in personal_ids:
        uuid.UUID(site_id)
    for site_id, row in published.items():
        db = dbs.get(row["ContentDatabaseId"])
        if not db or row["WebApplicationId"] != db["WebApplicationId"] or coverage_by_site[site_id]["Status"] != row["CollectionStatus"]:
            raise ValueError("Published collection relationship or status mismatch")
        if row["OwnerLogin"] not in upns or row["SecondaryOwnerLogin"] not in upns:
            raise ValueError("SharePoint owner not present in synthetic users")
        if (row["LastContentModifiedUtc"] and date.fromisoformat(row["LastContentModifiedUtc"][:10]) > as_of) or (not row["LastContentModifiedUtc"] and row["CollectionStatus"] != "MetadataIncomplete"):
            raise ValueError("Invalid last content modification date")
        if int(row["StorageBytes"]) < 0 or (row["QuotaLimitBytes"] and int(row["QuotaLimitBytes"]) <= 0):
            raise ValueError("Invalid site storage or quota")
        if row["QuotaUsagePercent"] and (not row["QuotaLimitBytes"] or not math.isclose(
                float(row["QuotaUsagePercent"]), round(int(row["StorageBytes"]) / int(row["QuotaLimitBytes"]) * 100, 2), abs_tol=0.01)):
            raise ValueError("Invalid site quota percentage")
        if row["LastContentModifiedUtc"] and not re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z", row["LastContentModifiedUtc"]):
            raise ValueError("Invalid site modification timestamp")
        if row["SiteCategory"] not in ("Other", "OtherOnMyHost") or row["RootWebTemplate"] != "STS#0":
            raise ValueError("Invalid published site category")
        if int(coverage_by_site[site_id]["WebsCollected"]) != config["WebsPerCollection"]:
            raise ValueError("Collection web count mismatch")
    for name, key in (("SiteAdministrators", "AdministratorLogin"), ("Webs", "WebId")):
        for row in tables[name]:
            site = published.get(row["SiteCollectionId"])
            if not site or row["ContentDatabaseId"] != site["ContentDatabaseId"] or row["WebApplicationId"] != site["WebApplicationId"]:
                raise ValueError(f"Invalid {name} relationship")
            if name == "SiteAdministrators" and row[key] not in upns:
                raise ValueError("SharePoint administrator not present in synthetic users")
            if name == "Webs" and (int(row["ListCount"]) < int(row["LibraryCount"]) or int(row["LibraryCount"]) < 0):
                raise ValueError("Invalid SharePoint web list counts")
    if len(tables["Webs"]) != len(published) * config["WebsPerCollection"] or len(tables["SiteAdministrators"]) != len(published) * config["AdministratorsPerCollection"]:
        raise ValueError("SharePoint web or administrator volume mismatch")
    web_ids = [row["WebId"] for row in tables["Webs"]]
    admin_keys = [(row["SiteCollectionId"], row["AdministratorLogin"]) for row in tables["SiteAdministrators"]]
    if len(web_ids) != len(set(web_ids)) or len(admin_keys) != len(set(admin_keys)):
        raise ValueError("Duplicate SharePoint web or administrator key")
    for web_id in web_ids:
        uuid.UUID(web_id)
    db_coverage = Counter(row["ContentDatabaseId"] for row in coverage if row["Status"] != "DatabaseCoverageMismatch")
    mismatch_dbs = set()
    for row in coverage:
        db = dbs.get(row["ContentDatabaseId"])
        if not db or row["WebApplicationId"] != db["WebApplicationId"]:
            raise ValueError("Coverage references unknown database or web application")
        if row["Status"] == "ExcludedPersonal" and (row["SiteCollectionId"] in published or "/personal/" not in row["Url"]):
            raise ValueError("Personal coverage must not appear in published collections")
        if row["Status"] == "DatabaseCoverageMismatch":
            if any(row[field] for field in ("SiteCollectionId", "Url", "WebsCollected")) or row["ErrorStage"] != "ContentDatabase":
                raise ValueError("Database coverage must have database grain")
            mismatch_dbs.add(row["ContentDatabaseId"])
        elif row["Status"] == "MetadataIncomplete" and (row["ErrorStage"] != "SiteCollectionMetadata" or not row["ErrorMessage"]):
            raise ValueError("Incomplete metadata requires an explanation")
    if len(mismatch_dbs) != config["DatabaseCoverageMismatch"]:
        raise ValueError("Duplicate database mismatch coverage")
    for db_id, db in dbs.items():
        expected_delta = 1 if db_id in mismatch_dbs else 0
        if int(db["SiteCollectionCount"]) != db_coverage[db_id] + expected_delta:
            raise ValueError("Content database site count mismatch")
    for name, records in tables.items():
        for record in records:
            for column, value in record.items():
                if column in ("Url", "DefaultUrl", "PublicUrl") and value and (urlparse(value).scheme != "https" or not urlparse(value).hostname.endswith(".contoso.com")):
                    raise ValueError(f"Non-Contoso URL in {name}")
                if column in ("ServerName", "DatabaseServer") and not value.endswith(".contoso.com"):
                    raise ValueError(f"Non-Contoso server in {name}")
                if column in ("IsReadOnly", "ReadLocked", "WriteLocked", "ContentDatabaseIsReadOnly",
                              "IsPrimary", "IsSecondary", "HasUniqueRoleAssignments", "ClaimsAuthentication",
                              "NeedsUpgradeIncludeChildren") and value and value not in ("True", "False"):
                    raise ValueError(f"Invalid Boolean value in {name}.{column}")
                if column in ("LanguageLcid", "ListCount", "LibraryCount", "StorageBytes", "QuotaLimitBytes",
                              "DatabaseSizeBytes", "SiteCollectionCount", "WebsCollected", "ContentDatabaseCount") and value:
                    if not re.fullmatch(r"\d+", value):
                        raise ValueError(f"Invalid integer value in {name}.{column}")
                if column in ("Name", "ConfigurationDatabase") and value and name in ("ServiceApplications", "WebApplications", "ContentDatabases", "Farms") and not value.startswith("Contoso"):
                    raise ValueError(f"Non-Contoso object name in {name}")
    return {"Files": len(tables), "Rows": {name: len(records) for name, records in tables.items()},
            "Coverage": dict(sorted(status.items())), "ServerCount": int(farm[0]["ServerCount"]),
            "ContentStatus": "Partial" if status["MetadataIncomplete"] + status["DatabaseCoverageMismatch"] else "Complete",
            "InfrastructureRunId": infra_run, "ContentRunId": content_run}


def make_people(count: int, as_of: date, rng: random.Random) -> tuple[list[dict], list[dict]]:
    site_catalog = []
    for country, _, number in COUNTRIES:
        prefix = re.sub("[^A-Z]", "", country.upper())[:3]
        for index in range(number):
            site_catalog.append({"SiteId": f"CT-{prefix}-{index+1:02d}",
                                 "SiteName": f"Contoso {prefix} Campus {index+1:02d}",
                                 "Country": country, "SiteType": "HQ" if country == "France" and index == 0 else "Office"})
    countries = weighted_labels(count, [item[0] for item in COUNTRIES], [item[1] for item in COUNTRIES], rng)
    departments = weighted_labels(count, [item[0] for item in DEPARTMENTS], [item[1] for item in DEPARTMENTS], rng)
    people = []
    for index in range(count):
        ordinal = index + 1
        country = countries[index]
        sites = [site for site in site_catalog if site["Country"] == country]
        site = sites[int(rng.random() ** 2 * len(sites))]
        first = FIRST[(index + rng.randrange(len(FIRST))) % len(FIRST)]
        last = LAST[rng.randrange(len(LAST))] + LAST_END[rng.randrange(len(LAST_END))]
        join = as_of - timedelta(days=rng.randint(60, 3200))
        last_activity = as_of - timedelta(days=(rng.randint(91, 400) if rng.random() < .03 else rng.randint(0, 28)))
        people.append({"PersonId": f"CT-PER-{ordinal:06d}", "FirstName": first, "LastName": last,
                       "DisplayName": f"{first} {last}", "UPN": f"{first.lower()}.{last.lower()}.{ordinal:05d}@contoso.com",
                       "Department": departments[index], "Country": country, "SiteId": site["SiteId"],
                       "JoinedDate": iso_day(join), "LastActivityUtc": iso_time(last_activity),
                       "AccountState": "Disabled" if rng.random() < .012 else "Enabled",
                       "ManagerPersonId": ""})
    by_group = defaultdict(list)
    for person in people:
        by_group[(person["Country"], person["Department"])].append(person)
    for group in by_group.values():
        manager = group[0]["PersonId"]
        for person in group[1:]:
            person["ManagerPersonId"] = manager
    return people, site_catalog


def device_specs(kind: str, rng: random.Random) -> tuple[str, str, str, str, str]:
    if kind in ("windows11", "windows10"):
        manufacturer = rng.choices(["Dell", "HP", "Lenovo", "Microsoft"], [38, 28, 27, 7])[0]
        current_models = {"Dell": "Latitude 7450", "HP": "EliteBook 840 G11", "Lenovo": "ThinkPad T14 Gen 5",
                          "Microsoft": "Surface Laptop 6"}
        legacy_models = {"Dell": "Latitude 5420", "HP": "EliteBook 840 G7", "Lenovo": "ThinkPad T14 Gen 1",
                         "Microsoft": "Surface Laptop 3"}
        model = (current_models if kind == "windows11" else legacy_models)[manufacturer]
        version = rng.choices(["25H2", "24H2", "23H2"], [40, 50, 10])[0] if kind == "windows11" else rng.choices(["22H2", "LTSC"], [90, 10])[0]
        build = {"25H2": "10.0.26200", "24H2": "10.0.26100", "23H2": "10.0.22631",
                 "22H2": "10.0.19045", "LTSC": "10.0.19044"}[version]
        return manufacturer, model, version, build, "Windows"
    if kind == "mac":
        return "Apple", rng.choice(["MacBook Air M3", "MacBook Pro M4"]), rng.choices(["macOS 15", "macOS 14"], [70, 30])[0], "", "macOS"
    if kind == "android":
        manufacturer = rng.choices(["Samsung", "Google", "Motorola", "Zebra"], [58, 20, 12, 10])[0]
        return manufacturer, {"Samsung": "Galaxy A55", "Google": "Pixel 9", "Motorola": "ThinkPhone", "Zebra": "TC58"}[manufacturer], rng.choices(["Android 16", "Android 15", "Android 14"], [55, 35, 10])[0], "", "Android"
    if kind == "tablets":
        if rng.random() < .65:
            return "Apple", "iPad Air", "iPadOS 26", "", "iPadOS"
        return "Samsung", "Galaxy Tab S9", "Android 15", "", "Android"
    return "Apple", rng.choice(["iPhone 15", "iPhone 16", "iPhone 17"]), rng.choices(["iOS 26", "iOS 18", "iOS 17"], [60, 30, 10])[0], "", "iOS"


def make_devices(people: list[dict], volumes: dict[str, int], as_of: date, seed: int, rng: random.Random) -> list[dict]:
    kinds = [kind for kind, count in volumes.items() for _ in range(count)]
    rng.shuffle(kinds)
    compliance = weighted_labels(len(kinds), ["Compliant", "Noncompliant", "Unknown"], [89, 7, 4], rng)
    orphan = set(rng.sample(range(len(kinds)), round(len(kinds) * .015)))
    stale = set(rng.sample(range(len(kinds)), round(len(kinds) * .08)))
    devices = []
    for index, kind in enumerate(kinds):
        person = people[index] if index < len(people) else people[rng.randrange(len(people))]
        manufacturer, model, version, build, family = device_specs(kind, rng)
        purchase = as_of - timedelta(days=rng.randint(1200, 2200) if kind == "windows10" else rng.randint(60, 1450))
        enrolled = purchase + timedelta(days=rng.randint(1, 30))
        last_sync = as_of - timedelta(days=rng.randint(31, 180) if index in stale else rng.randint(0, 14))
        if last_sync < enrolled:
            last_sync = enrolled
        device_id = f"CT-DEV-{index+1:06d}"
        eligibility = rng.choices(["Capable", "Not capable", "Undetermined"], [68, 22, 10])[0] if kind == "windows10" else "Adopted" if kind == "windows11" else "Not applicable"
        capacity = 512 if kind.startswith("windows") or kind == "mac" else 128
        devices.append({"DeviceId": device_id, "PersonId": person["PersonId"], "PrimaryUserUPN": person["UPN"],
                        "SiteId": person["SiteId"], "Country": person["Country"], "Department": person["Department"],
                        "Kind": kind, "Platform": "Windows" if kind.startswith("windows") else family,
                        "DeviceCategory": "PC Windows" if kind.startswith("windows") else "Mac" if kind == "mac" else "Tablet" if kind == "tablets" else "Mobile",
                        "Manufacturer": manufacturer, "Model": model, "OSFamily": family if not kind.startswith("windows") else "Windows 11" if kind == "windows11" else "Windows 10",
                        "OSVersion": version, "OSBuild": build, "SerialNumber": "CT" + stable_hex(seed, device_id, "serial", length=14).upper(),
                        "IMEI": luhn_imei(seed, index + 1) if kind in ("android", "iphone") or (kind == "tablets" and index % 2 == 0) else "",
                        "PurchaseDate": iso_day(purchase), "EnrollmentDateTime": iso_time(enrolled),
                        "LastSyncDateTime": iso_time(last_sync), "Compliance": compliance[index],
                        "DirectoryMatch": "Missing" if index in orphan else "Matched",
                        "UpgradeEligibility": eligibility, "FreeDiskGB": round(rng.uniform(8, capacity * .8), 1),
                        "StorageGB": capacity})
    return devices


def core_rows(people: list[dict], sites: list[dict], devices: list[dict], root: Path) -> None:
    write_csv(root / "reference" / "contoso_people.csv", list(people[0]), people)
    write_csv(root / "reference" / "contoso_sites.csv", list(sites[0]), sites)
    write_csv(root / "reference" / "contoso_assets.csv", list(devices[0]), devices)


def fixture_rows(people: list[dict], devices: list[dict], as_of: date, rng: random.Random) -> dict[str, list[dict]]:
    day, timestamp = iso_day(as_of), iso_time(as_of)
    rows: dict[str, list[dict]] = defaultdict(list)
    for person in people:
        uid = person["PersonId"]
        rows["users.csv"].append({"UserId": uid, "Country": person["Country"], "Department": person["Department"],
                                   "UserType": "Member", "AccountState": person["AccountState"], "WorkforceStatus": "Active" if person["AccountState"] == "Enabled" else "Departed",
                                   "LastActivityUtc": person["LastActivityUtc"]})
        rows["identity_matches.csv"].append({"UserId": uid, "SourceScope": "Hybrid", "MatchStatus": "Matched", "MatchMethod": "ImmutableId", "CandidateCount": 1})
        rows["identity_security_status.csv"].append({"ObservationDate": day, "UserId": uid, "SecurityPopulation": "Human", "MfaEligible": "true",
                                                "MfaCapable": "true" if rng.random() < .9 else "false", "PasswordlessCapable": "true" if rng.random() < .68 else "false",
                                                "PrivilegedAccount": "false", "PimProtected": "false", "MfaChallengeCount": 4, "MfaSuccessfulChallengeCount": 4, "EvidenceStatus": "Observed"})
        primary_license = ("LIC-002" if person["Department"] == "Field Services" else
                           "LIC-003" if len(rows["users.csv"]) % 25 == 0 else
                           "LIC-004" if len(rows["users.csv"]) % 37 == 0 else "LIC-001")
        rows["license_assignments.csv"].append({"UserId": uid, "LicenseProductId": primary_license,
                                                "AssignmentState": "Assigned", "InactivityEvidenceState": "Observed", "ApprovedExclusion": "false"})
        for license_id, interval in (("LIC-005", 20), ("LIC-006", 13), ("LIC-007", 10), ("LIC-008", 80)):
            if len(rows["users.csv"]) % interval == 0:
                rows["license_assignments.csv"].append({"UserId": uid, "LicenseProductId": license_id,
                                                        "AssignmentState": "Assigned", "InactivityEvidenceState": "Observed",
                                                        "ApprovedExclusion": "false"})
        if person["AccountState"] == "Enabled":
            rows["user_activity.csv"].append({"UserId": uid, "ServiceName": "Collaboration", "ActivityDateUtc": person["LastActivityUtc"], "ActivitySignal": "ActiveUse"})
    for device in devices:
        did = device["DeviceId"]
        rows["devices.csv"].append({"DeviceId": did, "Country": device["Country"], "Department": device["Department"],
                                    "Platform": device["Platform"], "DeviceCategory": device["DeviceCategory"], "Ownership": "Corporate",
                                    "ManagementState": "Managed", "ComplianceState": device["Compliance"], "Manufacturer": device["Manufacturer"],
                                    "Model": device["Model"], "OperatingSystemFamily": device["OSFamily"], "OperatingSystemVersion": device["OSVersion"],
                                    "ReadinessState": device["UpgradeEligibility"], "LastActivityUtc": device["LastSyncDateTime"],
                                    "SystemDriveCapacityGB": device["StorageGB"], "FreeDiskSpaceGB": device["FreeDiskGB"]})
        rows["device_matches.csv"].append({"DeviceId": did, "SourceScope": "DirectoryAndManagement", "MatchStatus": device["DirectoryMatch"],
                                           "MatchMethod": "DeviceId" if device["DirectoryMatch"] == "Matched" else "None",
                                           "CandidateCount": 1 if device["DirectoryMatch"] == "Matched" else 0})
        rows["device_compliance.csv"].append({"DeviceId": did, "ObservationDate": day, "ComplianceState": device["Compliance"],
                                              "BlockedState": "true" if device["Compliance"] == "Noncompliant" else "false"})
        rows["endpoint_security_status.csv"].append({"ObservationDate": day, "DeviceId": did, "SecurityEligibilityState": "Eligible",
                                                    "DefenderHealthy": "true" if device["Compliance"] == "Compliant" else "false",
                                                    "DiskEncrypted": "true" if device["Compliance"] == "Compliant" else "false",
                                                    "FirewallHealthy": "true", "AntivirusHealthy": "true", "EvidenceStatus": "Observed"})
        rows["relationships.csv"].append({"RelationshipId": f"CT-REL-{did}", "RelationshipType": "PrimaryUser",
                                          "FromEntityType": "User", "FromEntityId": device["PersonId"], "ToEntityType": "Device", "ToEntityId": did,
                                          "ServiceName": "Device Management", "EvidenceStatus": "Observed"})
        if device["Kind"].startswith("windows"):
            rows["windows_lifecycle.csv"].append({"DeviceId": did, "AssessmentDate": day, "OperatingSystemFamily": device["OSFamily"],
                                                  "ReadinessState": device["UpgradeEligibility"],
                                                  "BlockerCategory": "Hardware" if device["UpgradeEligibility"] == "Not capable" else "None", "EndOfSupportDate": "2025-10-14" if device["Kind"] == "windows10" else "2027-11-09"})
            rows["endpoint_experience.csv"].append({"DeviceId": did, "MeasurementDateUtc": device["LastSyncDateTime"],
                                                   "MetricName": "EndpointAnalyticsScore", "MetricValue": rng.randint(35, 98), "MetricUnit": "Score", "MetricVersion": "1.0"})
    for lid, name, _, family in LICENSES:
        assigned = sum(row["LicenseProductId"] == lid for row in rows["license_assignments.csv"])
        rows["license_capacity.csv"].append({"LicenseProductId": lid, "LicenseProduct": name, "LicenseFamily": family,
                                             "EnabledUnits": assigned + max(20, round(assigned * .08)), "ConsumedUnits": assigned})
    for i in range(180):
        app_id = f"CT-APP-{i+1:03d}"
        for device in rng.sample(devices, min(len(devices), 20 + (180 - i) * 3)):
            rows["application_installations.csv"].append({"DeviceId": device["DeviceId"], "ApplicationId": app_id,
                                                            "NormalizedProduct": f"Contoso Workplace App {i+1:03d}", "Publisher": "Contoso Software",
                                                            "NormalizedVersion": f"{1+i%5}.{rng.randrange(1,4)}.0", "Platform": device["Platform"], "VersionStatus": "Current"})
    for i, person in enumerate(people[:min(len(people), 11000)]):
        mailbox_id = f"CT-MBX-{i+1:06d}"
        rows["mailboxes.csv"].append({"MailboxId": mailbox_id, "PrimarySmtpAddress": person["UPN"], "Country": person["Country"],
                                      "Department": person["Department"], "MailboxType": "UserMailbox", "RecipientType": "UserMailbox",
                                      "OperationalState": "Active", "HostingLocation": "Exchange Online", "ExchangeVersion": "Exchange Online",
                                      "MailboxSizeGb": round(rng.uniform(.2, 48), 2), "ArchiveState": "Enabled" if i % 8 == 0 else "Disabled",
                                      "ArchiveSizeGb": round(rng.uniform(.1, 10), 2) if i % 8 == 0 else 0,
                                      "DelegationCount": 0, "DelegationTypes": "None", "MigrationState": "Completed", "QuotaRaw": "50 GB", "QuotaGb": 50, "QuotaParseStatus": "Parsed"})
        rows["mailbox_matches.csv"].append({"MailboxId": mailbox_id, "MatchStatus": "Matched", "MatchMethod": "ImmutableId", "CandidateCount": 1})
        rows["backup_coverage.csv"].append({"EntityId": mailbox_id, "EntityType": "Mailbox", "ServiceName": "Exchange",
                                            "EligibilityState": "Eligible", "CoverageState": "Protected" if i % 20 else "Unprotected",
                                            "PolicyScope": "InScope", "ExclusionReason": ""})
    for i in range(900):
        rows["collaboration_usage.csv"].append({"UsageDate": day, "ServiceName": "Teams", "ObjectId": f"CT-TEAM-{i+1:04d}",
                                                "ObjectType": "Team", "EligibleUserCount": 8 + i % 50, "ActiveUserCount": 5 + i % 34, "ObjectCount": 1})
    for i in range(1200):
        rows["content_usage.csv"].append({"UsageDate": day, "SiteId": f"CT-SPO-{i+1:05d}", "SiteType": "TeamSite",
                                           "Workload": "SharePoint", "ActivityState": "Inactive" if i % 12 == 0 else "Active",
                                           "StorageUsedGb": round(rng.uniform(1, 900), 2), "OwnerEvidenceStatus": "Missing" if i % 30 == 0 else "Observed"})
    rows["conditional_access_coverage.csv"].append({"ObservationDate": day, "ApplicationName": "Microsoft Teams", "ApplicationCategory": "Collaboration",
                                                    "ActiveSignInCount": len(people) * 3, "CoverageState": "Covered", "EvidenceStatus": "Observed"})
    rows["security_posture.csv"].append({"ObservationDate": day, "ControlId": "CT-CTRL-001", "ControlName": "Multifactor Authentication Coverage",
                                         "SecurityDomain": "Identity", "ServiceName": "Identity", "SeverityLabel": "Medium",
                                         "EvidenceStatus": "Observed", "ObservedScore": 89, "MaximumScore": 100, "GapState": "Gap"})
    rows["source_runs.csv"].append({"DataSourceName": "Contoso Demo Generator", "RequiredSource": "true", "FreshnessTargetHours": 48,
                                    "LatestAcceptedRunUtc": timestamp, "ActualRowCount": len(people)+len(devices), "RejectRowCount": 0})
    rows["data_quality_results.csv"].append({"RuleId": "CT-RULE-001", "RuleName": "Valid entity references", "SeverityLabel": "Critical",
                                             "StatusLabel": "Passed", "AffectedRowCount": 0})
    rows["findings.csv"].append({"FindingId": "CT-FIND-001", "FindingCategory": "Device Compliance", "FindingTitle": "Noncompliant devices",
                                   "ServiceName": "Device Management", "SeverityLabel": "Medium", "EvidenceStatus": "Observed",
                                   "AffectedEntityType": "Device", "AffectedEntityId": next(d["DeviceId"] for d in devices if d["Compliance"] == "Noncompliant"),
                                   "RecommendedAction": "Review device compliance", "SourceContract": "ManagedDevices/1.0.0"})
    return rows


def write_fixtures(root: Path, rows: dict[str, list[dict]], as_of: date) -> dict[str, int]:
    counts = {}
    entries = []
    for source in sorted(FIXTURE_HEADERS.glob("*.csv")):
        with source.open(encoding="utf-8-sig", newline="") as stream:
            header = next(csv.reader(stream))
        source_row = {}
        with source.open(encoding="utf-8-sig", newline="") as stream:
            source_row = next(csv.DictReader(stream))
        common = {key: source_row[key] for key in header[:11]}
        common.update({"TenantId": TENANT_ID, "OrganizationId": "ORG-CONTOSO-DEMO", "EnvironmentId": "ENV-CONTOSO-DEMO",
                       "RunId": f"RUN-CONTOSO-{as_of.strftime('%Y%m%d')}", "CollectedAtUtc": iso_time(as_of),
                       "SourceScriptVersion": "contoso-demo-1.0.0"})
        output = [{**common, **row} for row in rows.get(source.name, [])]
        if not output:
            raise ValueError(f"No generated rows for fixture {source.name}")
        path = root / "fixtures" / "valid" / source.name
        counts[source.name] = write_csv(path, header, output)
        header_text = path.open(encoding="utf-8").readline().rstrip("\r\n")
        entries.append({"path": f"valid/{source.name}", "sourceFamily": common["SourceFamily"],
                        "sourceEntity": common["SourceEntity"], "grain": "One accepted Contoso demo observation per source grain",
                        "expectedOutcome": "Accept", "expectedRows": counts[source.name],
                        "schemaSha256": hashlib.sha256(header_text.encode("utf-8")).hexdigest().upper(),
                        "fileSha256": sha256(path)})
    (root / "fixtures" / "quality-cases").mkdir(parents=True, exist_ok=True)
    (root / "fixtures" / "source-contract.json").write_bytes((PRODUCT / "fixtures" / "source-contract.json").read_bytes())
    write_json(root / "fixtures" / "fixture-manifest.json",
               {"contractVersion": "1.0.0", "fixtureSet": "contoso-demo-v1", "generatedAtUtc": iso_time(as_of),
                "syntheticOnly": True,
                "defaultContext": {"tenantId": TENANT_ID, "organizationId": "ORG-CONTOSO-DEMO",
                                   "environmentId": "ENV-CONTOSO-DEMO", "runId": f"RUN-CONTOSO-{as_of.strftime('%Y%m%d')}",
                                   "collectedAtUtc": iso_time(as_of), "sourceScriptVersion": "contoso-demo-1.0.0"},
                "files": entries})
    return counts


def base_prepared_row(columns: list[dict], as_of: date) -> dict:
    row = {}
    for column in columns:
        name, kind = column["name"], column["type"]
        if kind in ("int64", "double", "decimal"):
            row[name] = 0
        elif kind == "boolean":
            row[name] = "false"
        elif kind == "dateTime":
            row[name] = iso_time(as_of)
        elif "Evidence Status" in name:
            row[name] = "Observed"
        elif "Snapshot Date" == name or "Evidence Date" == name:
            row[name] = iso_day(as_of)
        elif name == "Date Key":
            row[name] = as_of.strftime("%Y%m%d")
        elif name in ("Snapshot Week", "Week Label"):
            row[name] = f"{as_of.isocalendar().year}-W{as_of.isocalendar().week:02d}"
        else:
            row[name] = ""
    return row


def prepared_rows(name: str, columns: list[dict], people: list[dict], devices: list[dict],
                  fixtures: dict[str, list[dict]], as_of: date, rng: random.Random) -> list[dict]:
    base = base_prepared_row(columns, as_of)
    day = iso_day(as_of)
    week = f"{as_of.isocalendar().year}-W{as_of.isocalendar().week:02d}"
    def row(**values):
        return {**base, **values}
    if name == "WindowsMigrationEvidence.csv":
        return [row(**{"Device Source ID": d["DeviceId"], "Device Name": d["DeviceId"], "Country": d["Country"],
                       "Primary User UPN": d["PrimaryUserUPN"], "OS Version": d["OSBuild"], "Windows Release": d["OSVersion"],
                       "Intune State": "Managed", "Upgrade Eligibility": d["UpgradeEligibility"],
                       "Hardware Reasons": "TPM 2.0" if d["UpgradeEligibility"] == "Not capable" else "",
                       "Free Disk GB": d["FreeDiskGB"], "Disk Review": "Low space" if d["FreeDiskGB"] < 20 else "Sufficient",
                       "Migration Status": "Ready" if d["UpgradeEligibility"] == "Capable" else "Blocked" if d["UpgradeEligibility"] == "Not capable" else "Undetermined",
                       "Risk Priority": "High" if d["UpgradeEligibility"] != "Capable" else "Low",
                       "Next Action": "Plan Windows 11 upgrade" if d["UpgradeEligibility"] == "Capable" else "Review hardware or missing evidence",
                       "Evidence Date": day, "Inventory Source": "Synthetic Intune", "Last Intune Sync": d["LastSyncDateTime"],
                       "Inventory Observation Date": day, "Eligibility Observation Date": day, "Eligibility OS Version": d["OSBuild"],
                       "Eligibility Match Status": "Matched", "Country Basis": "Primary user"})
                for d in devices if d["Kind"] == "windows10"]
    if name == "WindowsMigrationADObservations.csv":
        return []
    if name == "DeviceInventoryEvidence.csv":
        return [row(**{"Device Source ID": d["DeviceId"], "Device Name": d["DeviceId"], "Primary User UPN": d["PrimaryUserUPN"],
                       "Country": d["Country"], "Platform": d["Platform"], "Device Category": d["DeviceCategory"], "Ownership": "Corporate",
                       "Management State": "Managed", "Compliance State": d["Compliance"], "Manufacturer": d["Manufacturer"], "Model": d["Model"],
                       "Operating System Family": d["OSFamily"], "Operating System Version": d["OSBuild"] or d["OSVersion"],
                       "Windows Release": d["OSVersion"] if d["Kind"].startswith("windows") else "",
                       "System Drive Capacity (GB)": d["StorageGB"], "Free Disk Space (GB)": d["FreeDiskGB"],
                       "Free Disk Space (%)": round(d["FreeDiskGB"] / d["StorageGB"] * 100, 1),
                       "Disk Space State": "Low" if d["FreeDiskGB"] < 20 else "Healthy", "AD to Intune Match": d["DirectoryMatch"],
                       "Entra Match": d["DirectoryMatch"], "Enrollment DateTime": d["EnrollmentDateTime"], "Last Sync DateTime": d["LastSyncDateTime"],
                       "Observed Since": d["EnrollmentDateTime"], "Source Scope": "Synthetic Contoso", "Windows Generation": d["OSFamily"] if d["Kind"].startswith("windows") else "",
                       "Windows 11 Upgrade Eligibility": d["UpgradeEligibility"], "Endpoint Analytics Score": rng.randint(35, 98) if d["Kind"].startswith("windows") else "",
                       "Endpoint Analytics State": "Observed" if d["Kind"].startswith("windows") else "Not applicable",
                       "Intune Encryption State": "Encrypted" if d["Compliance"] == "Compliant" else "Unknown"}) for d in devices]
    if name == "UserInventoryEvidence.csv":
        return [row(**{"User Source ID": p["PersonId"], "User Principal Name": p["UPN"], "Display Name": p["DisplayName"],
                       "Country": p["Country"], "Department": p["Department"], "Department Evidence State": "Observed",
                       "Manager": p["ManagerPersonId"], "Manager Evidence State": "Observed" if p["ManagerPersonId"] else "Missing",
                       "Account State": p["AccountState"], "User Type": "Member", "Workforce Status": "Active" if p["AccountState"] == "Enabled" else "Departed",
                       "Activity State": "Active" if (as_of - date.fromisoformat(p["LastActivityUtc"][:10])).days <= 30 else "Inactive",
                       "Activity Evidence Status": "Observed", "Days Since Last Activity": (as_of - date.fromisoformat(p["LastActivityUtc"][:10])).days,
                       "Last Activity Date": p["LastActivityUtc"], "Identity Match Status": "Matched", "Identity Match Method": "Synthetic person ID",
                       "M365 License Count": 1, "Has M365 License": "true", "Duplicate UPN State": "Unique", "Evidence Date": day,
                       "Account Type": "Human", "Account Population": "Workforce", "Job Title": p["Department"] + " Specialist",
                       "AD Last Activity Date": p["LastActivityUtc"], "M365 Last Activity Date": p["LastActivityUtc"],
                       "Activity Source": "Synthetic AD and M365", "Workforce Persona": "Field worker" if p["Department"] == "Field Services" else "Office worker",
                       "Persona Match State": "Matched", "MFA Registration State": "Registered", "Has Microsoft 365 F3": "true" if p["Department"] == "Field Services" else "false",
                       "Has Microsoft 365 E3": "false" if p["Department"] == "Field Services" else "true", "Directory Site Code": p["SiteId"],
                       "Site Type": "HQ" if p["SiteId"] == "CT-FRA-01" else "Office", "Site Match State": "Matched",
                       "Persona Classification Source": "Contoso demo rules"}) for p in people]
    if name == "UserActivityHistoryEvidence.csv":
        output = []
        latest_monday = as_of - timedelta(days=as_of.weekday())
        for weeks_back in range(13):
            snapshot = latest_monday - timedelta(days=7 * weeks_back)
            for p in people:
                if date.fromisoformat(p["JoinedDate"]) > snapshot:
                    continue
                last_activity = min(date.fromisoformat(p["LastActivityUtc"][:10]), snapshot)
                days_since = (snapshot - last_activity).days
                output.append(row(**{"Snapshot Date": iso_day(snapshot), "Date Key": snapshot.strftime("%Y%m%d"),
                                     "Week Label": f"{snapshot.isocalendar().year}-W{snapshot.isocalendar().week:02d}",
                                     "User Source ID": p["PersonId"], "User Principal Name": p["UPN"],
                                     "Activity State": "Active" if days_since <= 30 else "Inactive", "Activity Evidence Status": "Observed",
                                     "Days Since Last Activity": days_since, "Last Activity Date": iso_time(last_activity),
                                     "AD Last Activity Date": iso_time(last_activity), "M365 Last Activity Date": iso_time(last_activity),
                                     "Activity Source": "Synthetic", "Confirmed Human": "true"}))
        return output
    if name == "LicenseEvidence.csv":
        by_id = {item["LicenseProductId"]: item for item in fixtures["license_capacity.csv"]}
        return [row(**{"License Product": label, "SKU Part Number": sku, "Microsoft 365 Plan": label,
                       "Capacity Class": "Governed capacity", "Enabled Units": by_id[lid]["EnabledUnits"],
                       "Consumed Units": by_id[lid]["ConsumedUnits"], "Available Units": int(by_id[lid]["EnabledUnits"]) - int(by_id[lid]["ConsumedUnits"]),
                       "Capacity Utilization": round(int(by_id[lid]["ConsumedUnits"]) / int(by_id[lid]["EnabledUnits"]), 4),
                       "Assigned Users": by_id[lid]["ConsumedUnits"], "Snapshot Date": day, "Risk Priority": "Low"})
                for lid, label, sku, _ in LICENSES]
    if name == "LicenseOptimizationEvidence.csv":
        assigned = {item["UserId"]: item["LicenseProductId"] for item in fixtures["license_assignments.csv"]
                    if item["LicenseProductId"] in {"LIC-001", "LIC-002", "LIC-003", "LIC-004"}}
        product_by_id = {lid: (label, sku) for lid, label, sku, _ in LICENSES}
        return [row(**{"User Principal Name": p["UPN"], "SKU Part Number": product_by_id[assigned[p["PersonId"]]][1],
                       "License Product": product_by_id[assigned[p["PersonId"]]][0],
                       "Microsoft 365 Plan": product_by_id[assigned[p["PersonId"]]][0],
                       "Account State": p["AccountState"], "Activity State": "Inactive" if (as_of-date.fromisoformat(p["LastActivityUtc"][:10])).days > 90 else "Active",
                       "Last Activity Date": p["LastActivityUtc"], "Optimization Opportunity": "Review" if p["AccountState"] == "Disabled" else "None",
                       "Recommended Action": "Review license" if p["AccountState"] == "Disabled" else "Retain", "Snapshot Date": day}) for p in people]
    if name == "MailboxEvidence.csv":
        return [row(**{"Mailbox Key": m["MailboxId"], "Primary SMTP Address": m["PrimarySmtpAddress"], "Display Name": m["PrimarySmtpAddress"],
                       "Country": m["Country"], "Hosting Location": m["HostingLocation"], "Exchange Version": m["ExchangeVersion"],
                       "Mailbox Type": m["MailboxType"], "Recipient Type": m["RecipientType"], "Operational State": m["OperationalState"],
                       "Mailbox Size GB": m["MailboxSizeGb"], "Archive State": m["ArchiveState"], "Archive Size GB": m["ArchiveSizeGb"],
                       "Delegation Count": 0, "Delegation Types": "None", "Forwarding State": "Disabled", "Last Activity Date": day,
                       "Hosting State": "Hosted", "Risk Priority": "Low", "Recommended Action": "Retain", "Total Storage GB": round(float(m["MailboxSizeGb"]) + float(m["ArchiveSizeGb"]), 2)})
                for m in fixtures["mailboxes.csv"]]
    if name == "ApplicationInventoryEvidence.csv":
        grouped = defaultdict(list)
        for item in fixtures["application_installations.csv"]:
            grouped[item["ApplicationId"]].append(item)
        return [row(**{"Application Source ID": app, "Application Product Key": app, "Application Name": items[0]["NormalizedProduct"],
                       "Application Version": items[0]["NormalizedVersion"], "Publisher": items[0]["Publisher"], "Publisher State": "Known",
                       "Platform": "Mixed", "Device Count": len({item["DeviceId"] for item in items}), "Product Distinct Devices": len({item["DeviceId"] for item in items}),
                       "Product Total Installs": len(items), "Product Version Count": len({item["NormalizedVersion"] for item in items}),
                       "Metric Definition": "Distinct device installations", "Collection Scope": "Synthetic Contoso", "Snapshot Date": day, "Snapshot Week": week,
                       "Standardization State": "Review" if len(items) > 100 else "Healthy", "Recommended Action": "Review version spread"}) for app, items in sorted(grouped.items())]
    if name == "CollaborationObjectEvidence.csv":
        return [row(**{"Object ID": item["ObjectId"], "Object Name": item["ObjectId"], "Service": "Teams", "Object Type": "Team",
                       "Visibility": "Private", "Activity State": "Active", "Last Activity Date": day,
                       "Owner State": "Missing" if i % 30 == 0 else "Known", "Owner Count": 0 if i % 30 == 0 else 1,
                       "Member Count": item["EligibleUserCount"], "Guest Count": 0, "External Access State": "Disabled", "Risk Priority": "Low",
                       "Recommended Action": "Assign owner" if i % 30 == 0 else "Retain", "Snapshot Date": day})
                for i, item in enumerate(fixtures["collaboration_usage.csv"])]
    if name == "ContentStorageEvidence.csv":
        return [row(**{"Content Object Key": item["SiteId"], "Content Object Name": item["SiteId"], "Service": "SharePoint",
                       "Object Type": "TeamSite", "Owner": "" if i % 30 == 0 else people[i % len(people)]["UPN"],
                       "Owner State": "Missing" if i % 30 == 0 else "Known",
                       "Activity State": item["ActivityState"], "Last Activity Date": day, "Storage Used GB": item["StorageUsedGb"],
                       "Storage Allocated GB": 1000, "Storage Utilization (%)": round(float(item["StorageUsedGb"]) / 10, 1),
                       "Storage Pressure": "High" if float(item["StorageUsedGb"]) > 800 else "Normal", "File Count": 1000 + i,
                       "Active File Count": 800 + i, "Deleted State": "Active", "Risk Priority": "Medium" if i % 30 == 0 else "Low",
                       "Recommended Action": "Assign owner" if i % 30 == 0 else "Retain", "Snapshot Date": day})
                for i, item in enumerate(fixtures["content_usage.csv"])]
    if name == "BackupMailboxEvidence.csv":
        mailbox_by_id = {item["MailboxId"]: item for item in fixtures["mailboxes.csv"]}
        return [row(**{"Backup Entity Key": item["EntityId"], "Display Name": item["EntityId"],
                       "Primary Address": mailbox_by_id[item["EntityId"]]["PrimarySmtpAddress"],
                       "Entity Type": "Mailbox", "Mailbox Type": "UserMailbox", "Account State": "Enabled", "Expected Scope State": "InScope",
                       "Protection State": item["CoverageState"], "Reconciliation State": "Matched", "Protection Status": item["CoverageState"],
                       "Risk Priority": "Low" if item["CoverageState"] == "Protected" else "High",
                       "Recommended Action": "Review protection" if item["CoverageState"] != "Protected" else "Retain",
                       "Protection Snapshot Date": day, "Policy Snapshot Date": day}) for item in fixtures["backup_coverage.csv"]]
    if name == "DataTrustEvidence.csv":
        header_count = {}
        for source in FIXTURE_HEADERS.glob("*.csv"):
            with source.open(encoding="utf-8-sig", newline="") as stream:
                header_count[source.name] = len(next(csv.reader(stream)))
        return [row(**{"Source Key": f"CT-SRC-{i+1:03d}", "Evidence Domain": "Synthetic", "Source Name": source,
                       "File Name": source, "File State": "Present", "Schema State": "Valid", "Expected Columns": header_count[source],
                       "Missing Columns": 0, "Row Count": len(values), "Column Count": header_count[source], "Snapshot Date": day,
                       "Snapshot Basis": "Deterministic Contoso generator", "Freshness State": "Current", "Decision Readiness": "Demo only",
                       "Evidence Status": "Synthetic", "Recommended Action": "Do not use as production evidence"})
                for i, (source, values) in enumerate(sorted(fixtures.items()))]
    if name == "ExecutiveKPITrends.csv":
        output = []
        for month in range(12):
            moment = as_of - timedelta(days=month * 30)
            for metric, value in [("# Users", len(people)), ("# Devices", len(devices)),
                                  ("Windows 11 Adoption (%)", round(sum(d["Kind"] == "windows11" for d in devices) / max(1, sum(d["Kind"].startswith("windows") for d in devices)), 4)),
                                  ("M365 Active Use (30D) (%)", .88), ("Exchange Online Adoption (%)", .96), ("Secure Score (%)", .79)]:
                output.append(row(**{"Date Key": moment.strftime("%Y%m%d"), "Tenant Key": 1, "Environment Key": 1,
                                     "Metric Name": metric, "Metric Value": value, "Evidence Status": "Synthetic"}))
        return output
    if name in ("WorkforceTrendEvidence.csv", "WindowsLifecycleTrendEvidence.csv", "EndpointExperienceTrendEvidence.csv",
                "ApplicationTrendEvidence.csv", "CollaborationTrendEvidence.csv", "ContentStorageTrendEvidence.csv"):
        output = []
        latest_monday = as_of - timedelta(days=as_of.weekday())
        for offset in range(13):
            moment = latest_monday - timedelta(days=offset * 7)
            present_people = [p for p in people if date.fromisoformat(p["JoinedDate"]) <= moment]
            present_devices = [d for d in devices if date.fromisoformat(d["PurchaseDate"]) <= moment]
            windows_count = sum(d["Kind"].startswith("windows") for d in present_devices)
            win11_count = sum(d["Kind"] == "windows11" for d in present_devices)
            active_people = sum((moment - date.fromisoformat(p["LastActivityUtc"][:10])).days <= 30 for p in present_people)
            site_storage = round(sum(float(x["StorageUsedGb"]) for x in fixtures["content_usage.csv"]), 2)
            values = {"Snapshot Date": iso_day(moment), "Date Key": moment.strftime("%Y%m%d"),
                      "Week Label": f"{moment.isocalendar().year}-W{moment.isocalendar().week:02d}",
                      "Snapshot Week": f"{moment.isocalendar().year}-W{moment.isocalendar().week:02d}",
                      "Enabled Workforce Accounts": sum(p["AccountState"] == "Enabled" for p in present_people),
                      "Activity Covered Users": len(present_people), "Activity Evidence Coverage": 1,
                      "Active Workforce": active_people, "Accounts Without Activity >30D": len(present_people) - active_people,
                      "Confirmed Human Users": len(present_people), "Managed Devices": len(present_devices),
                      "Windows Devices": windows_count,
                      "Windows 10 Devices": sum(d["Kind"] == "windows10" for d in present_devices),
                      "Windows 11 Devices": win11_count, "Windows 11 Adoption (%)": round(win11_count / windows_count * 100, 1) if windows_count else 0,
                      "Endpoint Analytics Covered Devices": sum(d["Kind"].startswith("windows") for d in present_devices),
                      "Endpoint Analytics Score": 78, "Startup Covered Devices": sum(d["Kind"].startswith("windows") for d in present_devices),
                      "Startup Score": 76, "App Reliability Covered Devices": sum(d["Kind"].startswith("windows") for d in present_devices),
                      "App Reliability Score": 81, "Applications": 180, "Application Versions": 450,
                      "Publishers": 1, "Application Install Observations": len(fixtures["application_installations.csv"]),
                      "Service": "Teams", "Total Objects": 900, "Active Objects 30D": 810,
                      "Eligible Users": len(present_people), "Active Users 30D": round(len(present_people) * .88),
                      "Adoption Rate (%)": 88, "SharePoint Sites": 1200, "SharePoint Storage Used GB": site_storage,
                      "SharePoint Storage Allocated GB": 1200000,
                      "SharePoint Storage Utilization (%)": round(site_storage / 1200000 * 100, 1)}
            output.append(row(**values))
        return output
    if name == "DeviceRiskEvidenceSummary.csv":
        return [row(**{"Signal Sort": 1, "Signal": "Noncompliant devices", "Affected Devices": sum(d["Compliance"] == "Noncompliant" for d in devices),
                       "Covered Devices": len(devices), "Managed Devices Affected": sum(d["Compliance"] == "Noncompliant" for d in devices),
                       "Managed Devices Covered": len(devices), "Source Findings": 1})]
    if name == "IdentityReconciliationEvidence.csv":
        return [row(**{"Evidence Date": day, "Match Status": "Matched", "Match Method": "Synthetic person ID",
                       "Identities": len(people), "Eligible Identities": len(people), "Evidence Status": "Synthetic", "Recommended Action": "None"})]
    if name == "WorkforceOperationalSignals.csv":
        return [row(**{"Severity Label": "Medium", "Signal": "Inactive accounts", "Risk Category": "Identity Hygiene",
                       "Affected Users": sum((as_of-date.fromisoformat(p["LastActivityUtc"][:10])).days > 90 for p in people),
                       "Covered Population": len(people), "Affected Estate (%)": 3, "Evidence Status": "Synthetic",
                       "Recommended Action": "Review inactive accounts", "Source Scope": "Contoso demo"})]
    if name == "CollaborationAdoptionEvidence.csv":
        return [row(**{"Service": "Teams", "Eligible Users": len(people), "Active Users 30D": round(len(people) * .88),
                       "Active User Rate (%)": 88, "Service Objects": 900, "Active Objects 30D": 810, "Inactive Objects": 90,
                       "Orphaned Objects": 30, "Objects with External Access": 45, "Storage Used GB": 50000,
                       "Usage Refresh Date": day, "Freshness State": "Current", "Evidence Status": "Synthetic"})]
    if name == "SecurityControlEvidence.csv":
        return [row(**{"Control Sort": 1, "Security Domain": "Endpoint", "Control Name": "Device compliance",
                       "Severity Label": "Medium", "Evidence Status": "Synthetic", "Coverage State": "Observed",
                       "Covered Entities": len(devices), "Healthy Entities": sum(d["Compliance"] == "Compliant" for d in devices),
                       "Affected Entities": sum(d["Compliance"] == "Noncompliant" for d in devices),
                       "Health Rate": round(sum(d["Compliance"] == "Compliant" for d in devices) / len(devices), 4),
                       "Gap State": "Gap", "Evidence Source": "Contoso demo", "Recommended Action": "Review noncompliant devices",
                       "Snapshot Date": day, "Metric Unit": "Devices"})]
    if name == "DeviceDirectorySummary.csv":
        return [row(**{"Snapshot Date": day, "Enabled AD Windows PCs": sum(d["Kind"].startswith("windows") and d["DirectoryMatch"] == "Matched" for d in devices),
                       "Enabled AD Windows PCs Managed by Intune": sum(d["Kind"].startswith("windows") and d["DirectoryMatch"] == "Matched" for d in devices),
                       "Enabled AD Windows PCs Unmanaged": 0, "Disabled AD Computer Objects": 0,
                       "Disabled Entra Devices": 0, "Country Definition": "Primary user", "Evidence Status": "Synthetic",
                       "Enabled AD Windows 10 PCs": sum(d["Kind"] == "windows10" and d["DirectoryMatch"] == "Matched" for d in devices),
                       "Intune Windows 10 Eligibility Undetermined": sum(d["Kind"] == "windows10" and d["UpgradeEligibility"] == "Undetermined" for d in devices)})]
    raise ValueError(f"No generator for prepared file {name}")


def write_prepared(root: Path, contract: dict, people: list[dict], devices: list[dict],
                   fixtures: dict[str, list[dict]], as_of: date, seed: int, rng: random.Random) -> dict:
    batch_id = as_of.strftime("%Y%m%dT000000000Z") + "-" + stable_hex(seed, as_of, "batch", length=8)
    batch = root / "DATA-POWERBI" / "batches" / batch_id
    reports = []
    for table in contract["tables"]:
        headers = ["TenantKey"] + [column["name"] for column in table["columns"]]
        values = prepared_rows(table["file"], table["columns"], people, devices, fixtures, as_of, rng)
        if not values and table.get("allowEmpty") is not True:
            raise ValueError(f"Unexpected empty prepared file {table['file']}")
        for value in values:
            value["TenantKey"] = TENANT
        path = batch / table["file"]
        count = write_csv(path, headers, values)
        reports.append({"File": table["file"], "Rows": count, "Columns": len(headers), "SHA256": sha256(path)})
    validation = {"Passed": True, "SchemaOnly": False, "Files": reports, "Errors": []}
    write_json(batch / "validation.json.txt", validation)
    history_versions = {table["table"]: table.get("historyKeyVersion", 1) for table in contract["tables"] if "History" in table["table"] or "Trend" in table["table"]}
    manifest = {"SchemaVersion": 1, "BatchId": batch_id, "TenantKey": TENANT, "CreatedUtc": iso_time(as_of, 0),
                "Provenance": {"Kind": "SyntheticContosoDemo", "Seed": seed, "AsOf": iso_day(as_of), "NoCustomerData": True},
                "ContractSHA256": sha256(CONTRACT), "HistoryKeyVersions": history_versions, "HistoryKeyResets": [],
                "Files": [{"File": item["File"], "Rows": item["Rows"], "Bytes": (batch / item["File"]).stat().st_size,
                           "SHA256": item["SHA256"]} for item in reports]}
    write_json(batch / "batch.json.txt", manifest)
    pointer = {"SchemaVersion": 1, "BatchId": batch_id, "TenantKey": TENANT,
               "ManifestSHA256": sha256(batch / "batch.json.txt"), "PreviousBatchId": None}
    write_json(batch / "current.json.txt", pointer)
    write_json(root / "DATA-POWERBI" / "current.json.txt", pointer)
    return {"BatchId": batch_id, "Reports": reports}


def license_headers(file_name: str) -> list[str]:
    for path in MODEL_TABLES.glob("License Report *.tmdl"):
        match = re.search(r'GetLicenseReportCsv\("' + re.escape(file_name) + r'",\s*\{([^}]*)\}', path.read_text(encoding="utf-8-sig"))
        if match:
            return re.findall(r'"([^"]+)"', match.group(1))
    raise ValueError(f"Model schema missing for {file_name}")


def write_license_report(root: Path, people: list[dict], fixtures: dict[str, list[dict]], as_of: date, seed: int) -> None:
    folder = root / "DATA-LAST"
    snapshot_id = stable_hex(seed, as_of, "license", length=32)
    assignments = Counter(item["LicenseProductId"] for item in fixtures["license_assignments.csv"])
    primary_by_user = {item["UserId"]: item["LicenseProductId"] for item in fixtures["license_assignments.csv"]
                       if item["LicenseProductId"] in {"LIC-001", "LIC-002", "LIC-003", "LIC-004"}}
    label_by_id = {lid: label for lid, label, _, _ in LICENSES}
    dormant = [p for p in people if (as_of-date.fromisoformat(p["LastActivityUtc"][:10])).days > 90]
    candidates_by_product = Counter(label_by_id[primary_by_user[p["PersonId"]]] for p in dormant)
    four = [(label, assignments[lid]) for lid, label, _, _ in LICENSES[:4]]
    other = [(label, assignments[lid]) for lid, label, _, _ in LICENSES[4:7]]
    all_products = [(name, count, "Suite") for name, count in four] + [(name, count, "Other") for name, count in other]
    summary = []
    for name, assigned, category in all_products:
        summary.append({"SnapshotId": snapshot_id, "TenantKey": TENANT, "GeneratedAtUtc": iso_time(as_of),
                        "LicenseCollectedAtUtc": iso_time(as_of), "Product": name, "Category": category,
                        "EnabledUnits": assigned + 100, "ConsumedUnits": assigned, "AvailableUnits": 100,
                        "Subscribed": "true", "AssignedUsers": assigned, "RecoveryCandidates": candidates_by_product[name],
                        "RecoveryUnknown": 0, "NoM365Activity": candidates_by_product[name],
                        "EvidenceStatus": "Synthetic", "WorkbookQualification": "Demo"})
    candidates = [{"SnapshotId": snapshot_id, "TenantKey": TENANT, "CandidateType": "InactiveUser",
                   "Product": label_by_id[primary_by_user[p["PersonId"]]],
                   "UserId": p["PersonId"], "TenantUserKey": p["PersonId"], "IdentityJoinStatus": "Matched",
                   "UserPrincipalName": p["UPN"], "DisplayName": p["DisplayName"], "Reason": "Inactive over 90 days",
                   "LastAdActivityDate": p["LastActivityUtc"], "LastM365ActivityDate": p["LastActivityUtc"],
                   "PrimaryOnIntuneWindowsPc": "true", "TargetSuites": label_by_id[primary_by_user[p["PersonId"]]], "MailboxSizeGB": 1.2,
                   "OneDriveUsedGB": 2.5, "OneDriveUsedBytes": 2684354560, "ArchiveStatus": "Disabled",
                   "LitigationHoldEnabled": "false", "RetentionHoldEnabled": "false", "WindowsAppsUse180d": "false", "MacAppsUse180d": "false"}
                  for p in dormant]
    gaps = [{"SnapshotId": snapshot_id, "TenantKey": TENANT, "Section": "Demo", "Metric": f"Metric {i+1:02d}",
             "Value": i, "EvidenceStatus": "Synthetic"} for i in range(25)]
    sources = [{"SnapshotId": snapshot_id, "TenantKey": TENANT, "Name": "Contoso Demo Generator", "Ready": "true",
                "Provisional": "false", "Date": iso_day(as_of), "Reason": "Synthetic demo", "SHA256": stable_hex(seed, "source", length=64).upper()}]
    for suffix, values in [("Summary", summary), ("Candidates", candidates), ("Gaps", gaps), ("Sources", sources)]:
        name = f"M365_Licenses_Report{suffix}.csv"
        write_csv(folder / name, license_headers(name), values)
    product_json = [{"Product": name, "Enabled": count + 100, "Consumed": count, "Subscribed": True,
                     "Counts": {"RecoveryCandidates": candidates_by_product[name], "M365Unknown": 0}} for name, count in four]
    other_json = [{"Product": name, "Enabled": count + 100, "Consumed": count, "Subscribed": True} for name, count in other]
    write_json(folder / "M365_Licenses_ReportSnapshot.json.txt",
               {"SchemaVersion": 1, "SnapshotId": snapshot_id, "TenantKey": TENANT,
                "GeneratedAtUtc": iso_time(as_of), "LicenseCollectedAtUtc": iso_time(as_of),
                "Products": product_json, "OtherProducts": other_json, "SyntheticOnly": True})


def validate(root: Path, expected_people: int, volumes: dict[str, int], contract: dict,
             sp_config: dict[str, int] | None = None) -> dict:
    def read(path: Path) -> list[dict]:
        with path.open(encoding="utf-8", newline="") as stream:
            return list(csv.DictReader(stream))
    people = read(root / "reference" / "contoso_people.csv")
    devices = read(root / "reference" / "contoso_assets.csv")
    if len(people) != expected_people or len(devices) != sum(volumes.values()):
        raise ValueError("Requested volume mismatch")
    person_ids = {person["PersonId"] for person in people}
    upns = {person["UPN"] for person in people}
    if len(person_ids) != len(people) or len(upns) != len(people) or any(not item.endswith("@contoso.com") for item in upns):
        raise ValueError("Invalid person IDs or UPNs")
    device_ids = {device["DeviceId"] for device in devices}
    if len(device_ids) != len(devices) or any(device["PersonId"] not in person_ids or device["PrimaryUserUPN"] not in upns for device in devices):
        raise ValueError("Invalid device ownership")
    serials = {device["SerialNumber"] for device in devices}
    imeis = [device["IMEI"] for device in devices if device["IMEI"]]
    if len(serials) != len(devices) or len(imeis) != len(set(imeis)):
        raise ValueError("Duplicate serial number or IMEI")
    for imei in imeis:
        digits = [int(value) for value in imei]
        score = sum((digit if index % 2 == 0 else (digit * 2 // 10 + digit * 2 % 10))
                    for index, digit in enumerate(digits[:-1]))
        if len(imei) != 15 or (score + digits[-1]) % 10:
            raise ValueError("Invalid IMEI checksum")
    counts = Counter(device["Kind"] for device in devices)
    if dict(counts) != {kind: count for kind, count in volumes.items() if count}:
        raise ValueError("Device category mismatch")
    for device in devices:
        if not (device["PurchaseDate"] <= device["EnrollmentDateTime"][:10] <= device["LastSyncDateTime"][:10]):
            raise ValueError("Invalid device chronology")
        if not (0 <= float(device["FreeDiskGB"]) <= float(device["StorageGB"])):
            raise ValueError("Invalid storage capacity")
    batch_root = root / "DATA-POWERBI"
    pointer = json.loads((batch_root / "current.json.txt").read_text(encoding="utf-8"))
    batch = batch_root / "batches" / pointer["BatchId"]
    manifest_path = batch / "batch.json.txt"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if pointer["ManifestSHA256"] != sha256(manifest_path) or manifest["ContractSHA256"] != sha256(CONTRACT):
        raise ValueError("Prepared metadata hash mismatch")
    if len(manifest["Files"]) != len(contract["tables"]) or len(contract["tables"]) != 26:
        raise ValueError("Prepared file count mismatch")
    for table, entry in zip(contract["tables"], manifest["Files"]):
        path = batch / table["file"]
        if entry["File"] != table["file"] or entry["SHA256"] != sha256(path) or entry["Bytes"] != path.stat().st_size:
            raise ValueError(f"Prepared manifest mismatch: {table['file']}")
        with path.open(encoding="utf-8", newline="") as stream:
            reader = csv.DictReader(stream)
            if reader.fieldnames != ["TenantKey"] + [column["name"] for column in table["columns"]]:
                raise ValueError(f"Prepared header mismatch: {table['file']}")
            keys = set()
            count = 0
            for record in reader:
                count += 1
                if record["TenantKey"] != TENANT:
                    raise ValueError(f"Tenant mismatch: {table['file']}")
                if table.get("uniqueKey"):
                    key = tuple(record[field].casefold() for field in table["uniqueKey"])
                    if not all(key) or key in keys:
                        raise ValueError(f"Invalid unique key: {table['file']}")
                    keys.add(key)
                for column in table["columns"]:
                    value = record[column["name"]]
                    if value and column["type"] in ("int64", "double", "decimal"):
                        number = float(value)
                        if not math.isfinite(number) or (column["type"] == "int64" and not number.is_integer()):
                            raise ValueError(f"Invalid numeric value: {table['file']}")
            if count != entry["Rows"] or (count == 0 and not table.get("allowEmpty")):
                raise ValueError(f"Prepared row count mismatch: {table['file']}")
    migration = read(batch / "WindowsMigrationEvidence.csv")
    if len(migration) != volumes["windows10"] or any(row["Primary User UPN"] not in upns for row in migration):
        raise ValueError("Windows 10 migration coverage mismatch")
    if read(batch / "WindowsMigrationADObservations.csv"):
        raise ValueError("AD migration observations should be header-only")
    if len(read(batch / "DeviceInventoryEvidence.csv")) != len(devices) or len(read(batch / "UserInventoryEvidence.csv")) != len(people):
        raise ValueError("Prepared entity coverage mismatch")
    prepared_devices = read(batch / "DeviceInventoryEvidence.csv")
    if any(row["Primary User UPN"] not in upns for row in prepared_devices):
        raise ValueError("Prepared device user mismatch")
    people_by_upn = {person["UPN"]: person for person in people}
    if any(row["Country"] != people_by_upn[row["Primary User UPN"]]["Country"] for row in prepared_devices):
        raise ValueError("Prepared device geography mismatch")
    fixture_root = root / "fixtures" / "valid"
    fixture_users = read(fixture_root / "users.csv")
    fixture_devices = read(fixture_root / "devices.csv")
    fixture_user_ids = {item["UserId"] for item in fixture_users}
    fixture_device_ids = {item["DeviceId"] for item in fixture_devices}
    if fixture_user_ids != person_ids or fixture_device_ids != device_ids:
        raise ValueError("Fixture entity identity mismatch")
    license_ids = {item["LicenseProductId"] for item in read(fixture_root / "license_capacity.csv")}
    for item in read(fixture_root / "license_assignments.csv"):
        if item["UserId"] not in person_ids or item["LicenseProductId"] not in license_ids:
            raise ValueError("Invalid fixture license assignment")
    for item in read(fixture_root / "application_installations.csv"):
        if item["DeviceId"] not in device_ids:
            raise ValueError("Invalid fixture application installation")
    relationships = read(fixture_root / "relationships.csv")
    if len(relationships) != len(devices) or any(item["FromEntityId"] not in person_ids or item["ToEntityId"] not in device_ids for item in relationships):
        raise ValueError("Invalid device-user relationships")
    for source in FIXTURE_HEADERS.glob("*.csv"):
        with source.open(encoding="utf-8-sig", newline="") as stream:
            expected = next(csv.reader(stream))
        with (root / "fixtures" / "valid" / source.name).open(encoding="utf-8", newline="") as stream:
            reader = csv.DictReader(stream)
            if reader.fieldnames != expected or not any(True for _ in reader):
                raise ValueError(f"Fixture schema or population mismatch: {source.name}")
    fixture_manifest = json.loads((fixture_root.parent / "fixture-manifest.json").read_text(encoding="utf-8"))
    if fixture_manifest["syntheticOnly"] is not True or len(fixture_manifest["files"]) != 24:
        raise ValueError("Invalid fixture manifest")
    for entry in fixture_manifest["files"]:
        path = fixture_root.parent / entry["path"]
        header_text = path.open(encoding="utf-8").readline().rstrip("\r\n")
        if entry["fileSha256"] != sha256(path) or entry["schemaSha256"] != hashlib.sha256(header_text.encode("utf-8")).hexdigest().upper():
            raise ValueError(f"Fixture hash mismatch: {entry['path']}")
    report_root = root / "DATA-LAST"
    snapshot = json.loads((report_root / "M365_Licenses_ReportSnapshot.json.txt").read_text(encoding="utf-8"))
    if len(snapshot["Products"]) != 4 or len(snapshot["OtherProducts"]) != 3:
        raise ValueError("License snapshot product mismatch")
    for suffix in ("Summary", "Candidates", "Gaps", "Sources"):
        name = f"M365_Licenses_Report{suffix}.csv"
        with (report_root / name).open(encoding="utf-8", newline="") as stream:
            reader = csv.DictReader(stream)
            if reader.fieldnames != license_headers(name) or any(row["SnapshotId"] != snapshot["SnapshotId"] or row["TenantKey"] != TENANT for row in reader):
                raise ValueError(f"License report mismatch: {name}")
    candidate_counts = Counter(item["Product"] for item in read(report_root / "M365_Licenses_ReportCandidates.csv"))
    summary = read(report_root / "M365_Licenses_ReportSummary.csv")
    if any(int(item["RecoveryCandidates"]) != candidate_counts[item["Product"]] for item in summary):
        raise ValueError("License candidate totals do not reconcile")
    compliance_counts = Counter(device["Compliance"] for device in devices)
    batch_date = date(int(pointer["BatchId"][:4]), int(pointer["BatchId"][4:6]), int(pointer["BatchId"][6:8]))
    result = {"Passed": True, "Users": len(people), "Devices": len(devices), "DeviceKinds": dict(sorted(counts.items())),
            "Compliance": dict(sorted(compliance_counts.items())),
            "DirectoryOrphans": sum(device["DirectoryMatch"] == "Missing" for device in devices),
            "InactiveDevicesOver30Days": sum((batch_date - date.fromisoformat(device["LastSyncDateTime"][:10])).days > 30 for device in devices),
            "Windows10Eligibility": dict(sorted(Counter(item["Upgrade Eligibility"] for item in migration).items())),
            "PreparedFiles": len(contract["tables"]), "Windows10MigrationRows": len(migration), "Fixtures": len(list(FIXTURE_HEADERS.glob('*.csv')))}
    if sp_config is not None:
        result["SharePointOnPrem"] = validate_sharepoint_onprem(root, people, batch_date, sp_config)
    return result


def manifest_files(root: Path) -> list[dict]:
    folders = ("reference", "fixtures", "DATA-POWERBI", "DATA-LAST")
    paths = sorted(path for folder in folders for path in (root / folder).rglob("*") if path.is_file())
    return [{"Path": path.relative_to(root).as_posix(), "Bytes": path.stat().st_size, "SHA256": sha256(path)}
            for path in paths]


def verify_demo_manifest(root: Path, allow_extra_sharepoint: bool = False) -> dict:
    manifest = json.loads((root / "demo-manifest.json.txt").read_text(encoding="utf-8"))
    actual = manifest_files(root)
    if allow_extra_sharepoint and not manifest.get("SharePointOnPrem"):
        actual = [entry for entry in actual if not entry["Path"].startswith("DATA-LAST/SharePoint_OnPrem_")]
    if manifest["Files"] != actual or not manifest["SyntheticOnly"]:
        raise ValueError("Contoso demo manifest file hashes or paths mismatch")
    return manifest


RAW_IDENTITY = ["TenantKey", "OrganizationKey", "EnvironmentKey", "TenantId"]
RAW_LOCAL_SYSTEM = RAW_IDENTITY + [
    "DeviceName", "DeviceId", "AzureADDeviceId", "UserPrincipalName", "SecureBootStatus",
    "BIOSVersion", "BIOSDate", "FirmwareType", "PlatformScriptLastUpdate", "LastRebootDate",
    "Last Reboot Date", "LastBootUpTime", "LastBootTime", "IsEncrypted", "DeviceGuardVBSState",
    "DeviceGuardVBSHardwareRequirement", "DeviceGuardCredentialGuardState", "ManagementState",
    "ComplianceState", "LastSyncDateTime", "ThirdPartyAVStatus", "ThirdPartyAVProducts",
    "ExpectedThirdPartyAVName", "ExpectedThirdPartyAVStatus", "WindowsUpdateRebootRequired",
    "BitLockerOSProtectionStatus", "BitLockerOSConversionStatus"]
RAW_MANAGED = RAW_IDENTITY + [
    "ManagedDeviceId", "AzureADDeviceId", "DeviceName", "OperatingSystem", "OsVersion",
    "ManagedDeviceOwnerType", "ComplianceState", "ManagementAgent", "DeviceEnrollmentType",
    "EnrolledDateTime", "LastSyncDateTime", "UserId", "UserPrincipalName", "Manufacturer",
    "Model", "SerialNumber", "IsEncrypted", "TotalStorageSpaceInBytes", "FreeStorageSpaceInBytes",
    "CollectedAtUtc"]
RAW_DIRECTORY = RAW_IDENTITY + ["DomainName", "ObjectClass", "DistinguishedName", "ObjectGUID",
                                "ObjectSID", "PrimaryGroupID"]
RAW_BATTERY = RAW_IDENTITY + ["RunId", "ReportName", "ReportRefreshDate", "DeviceId", "DeviceName",
                              "Manufacturer", "Model", "BatteryId", "BatteryCount", "BatteryHealthScore",
                              "BatteryCapacityScore", "BatteryRuntimeScore", "MaximumCapacity", "CycleCount",
                              "DesignCapacity", "FullChargeCapacity", "NeedsAttention"]
RAW_BATTERY_QUALITY = RAW_IDENTITY + ["RunId", "ReportName", "ApiVersion", "Status", "RowCount",
                                      "ExportJobStatus", "IsAdvancedAnalytics", "RequiredPermission", "ErrorCode",
                                      "ErrorMessage", "CollectedAtUtc", "RawRowCount", "ExcludedRowCount",
                                      "ExcludedDeviceCount"]


def source_headers() -> dict[str, list[str]]:
    """Extract exact ordered headers from producer source; no operational data is read."""
    inventory = PRODUCT.parent / "SmartM365" / "SmartInventory"
    ad = (inventory / "ActiveDirectoryInventory" / "SmartM365-ActiveDirectory-Inventory.ps1").read_text(encoding="utf-8-sig")
    users = (inventory / "ActiveDirectoryInventory" / "SmartM365-ActiveDirectory-UsersEnrichment.ps1").read_text(encoding="utf-8-sig")
    computer_extra = (inventory / "ActiveDirectoryInventory" / "SmartM365-ActiveDirectory-EnrichedColumns.ps1").read_text(encoding="utf-8-sig")
    m365 = (inventory / "M365Inventory" / "Users" / "SmartM365-ActiveUsers-Inventory.ps1").read_text(encoding="utf-8-sig")
    intune = (inventory / "M365Inventory" / "IntuneInventory" / "Devices" / "SmartM365-Devices-Inventory.ps1").read_text(encoding="utf-8-sig")
    def quoted(block: str) -> list[str]:
        return re.findall(r"['\"]([^'\"]+)['\"]", block)
    def unique(columns: list[str]) -> list[str]:
        result, seen = [], set()
        for column in columns:
            if column.casefold() not in seen:
                result.append(column)
                seen.add(column.casefold())
        return result
    native = {}
    for kind in ("Users", "Computers"):
        match = re.search(r"'" + kind + r"'\s*\{\s*@\((.*?)\)\s*\}", ad, re.S)
        if not match:
            raise ValueError(f"Cannot read AD {kind} producer schema")
        native[kind] = quoted(match.group(1))
    user_calculated = re.search(r"\$calculatedColumns\s*=\s*@\((.*?)\)\s*WriteLog", users, re.S)
    active_block = re.search(r"\$obj\s*=\s*\[ordered\]@\{(.*?)\n\s*\}", m365, re.S)
    inventory_block = re.search(r"function Get-InventoryColumns\s*\{\s*\[string\[\]\]@\((.*?)\)\s*\}", intune, re.S)
    if not all((user_calculated, active_block, inventory_block)):
        raise ValueError("Cannot read current source producer schemas")
    active_columns = re.findall(r'^\s*"([^"]+)"\s*=', active_block.group(1), re.M)
    intune_columns = quoted(inventory_block.group(1))
    schemas = {
        "M365_Users_Active.csv": RAW_IDENTITY + active_columns,
        "AD_Users_AllDomains.csv": RAW_IDENTITY + unique(native["Users"] + quoted(user_calculated.group(1))),
        "AD_Computers_AllDomains.csv": RAW_IDENTITY + unique(native["Computers"] + quoted(computer_extra)),
        "Intune_Devices_Inventory.csv": RAW_IDENTITY + intune_columns,
        "Intune_Devices_LocalSystem.csv": RAW_LOCAL_SYSTEM,
        "AD_DirectoryObjects_AllDomains.csv": RAW_DIRECTORY,
        "Intune_ManagedDevices_All.csv": RAW_MANAGED,
        "Intune_EndpointAnalytics_BatteryHealth.csv": RAW_BATTERY,
        "Intune_EndpointAnalytics_DataQuality.csv": RAW_BATTERY_QUALITY,
    }
    if len(active_columns) != 37 or len(intune_columns) != 40 or len(native["Users"]) < 60 or len(native["Computers"]) < 50:
        raise ValueError("Producer schema extraction returned an unexpected column count")
    for name, columns in schemas.items():
        if len(columns) != len(set(columns)):
            raise ValueError(f"Duplicate source header: {name}")
    return schemas


def write_source_csv(path: Path, headers: list[str], rows: list[dict]) -> int:
    """Match Export-SmartM365Csv: comma, quoted fields, UTF-8 and CRLF."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=headers, extrasaction="ignore",
                                quoting=csv.QUOTE_ALL, lineterminator="\r\n")
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


def deterministic_guid(seed: int, kind: str, alias: str) -> str:
    return str(uuid.uuid5(uuid.NAMESPACE_DNS, f"contoso.com:{seed}:{kind}:{alias}"))


def versioned_entities(people: list[dict], devices: list[dict], seed: int) -> tuple[list[dict], list[dict]]:
    """Keep CT aliases in the legacy reference while using strong IDs in batch 1.2."""
    person_ids = {p["PersonId"]: deterministic_guid(seed, "entra-user", p["PersonId"]) for p in people}
    new_people = []
    for person in people:
        value = dict(person)
        value["Alias"] = value["PersonId"]
        value["PersonId"] = person_ids[value["Alias"]]
        value["ADGuid"] = deterministic_guid(seed, "ad-user", value["Alias"])
        value["ManagerPersonId"] = person_ids.get(value["ManagerPersonId"], "")
        if value["LastActivityUtc"][:10] < value["JoinedDate"]:
            value["LastActivityUtc"] = iso_time(date.fromisoformat(value["JoinedDate"]))
        new_people.append(value)
    new_devices = []
    for device in devices:
        value = dict(device)
        value["Alias"] = value["DeviceId"]
        value["DeviceId"] = deterministic_guid(seed, "intune-device", value["Alias"])
        value["AzureId"] = deterministic_guid(seed, "entra-device", value["Alias"])
        value["PersonId"] = person_ids[value["PersonId"]]
        new_devices.append(value)
    return new_people, new_devices


def synthetic_identity() -> dict[str, str]:
    return {"TenantKey": TENANT, "OrganizationKey": "contoso", "EnvironmentKey": "demo",
            "TenantId": TENANT_ID}


def source_rows(people: list[dict], devices: list[dict], as_of: date, seed: int,
                include_battery: bool, battery_coverage: int) -> dict[str, list[dict]]:
    stamp = iso_time(as_of)
    identity = synthetic_identity()
    country_code = {"France": "FR", "United Kingdom": "GB", "Germany": "DE", "Spain": "ES",
                    "United States": "US", "Netherlands": "NL"}
    output: dict[str, list[dict]] = defaultdict(list)
    people_by_id = {p["PersonId"]: p for p in people}
    for index, person in enumerate(people):
        ad_guid = uuid.UUID(person["ADGuid"])
        immutable = base64.b64encode(ad_guid.bytes_le).decode("ascii")
        sid = f"S-1-5-21-20061009-20261009-12000-{1000 + index}"
        common = {**identity, "Country": person["Country"], "Department": person["Department"]}
        output["M365_Users_Active.csv"].append({**common, "Object Id": person["PersonId"],
            "User principal name": person["UPN"], "Display name": person["DisplayName"],
            "First name": person["FirstName"], "Last name": person["LastName"],
            "CountryOrRegion": person["Country"], "Usage location": country_code[person["Country"]],
            "AccountEnabled": str(person["AccountState"] == "Enabled").lower(), "UserType": "Member",
            "OnPremisesImmutableId": immutable, "OnPremisesSecurityIdentifier": sid,
            "OnPremisesSyncEnabled": "true", "DirSyncEnabled": "true",
            "When created": iso_time(date.fromisoformat(person["JoinedDate"])),
            "LastSignInDateTime": person["LastActivityUtc"],
            "LastSuccessfulSignInDateTime": person["LastActivityUtc"], "CollectedAtUtc": stamp})
        output["AD_Users_AllDomains.csv"].append({**common, "DomainName": "contoso.com", "ObjectType": "User",
            "ObjectGUID": person["ADGuid"], "ObjectSID": sid, "ImmutableId_AD": immutable,
            "SamAccountName": person["UPN"].split("@")[0], "Name": person["DisplayName"],
            "DisplayName": person["DisplayName"], "GivenName": person["FirstName"],
            "Surname": person["LastName"], "UserPrincipalName": person["UPN"],
            "Enabled": str(person["AccountState"] == "Enabled").lower(),
            "LastLogonDate": person["LastActivityUtc"], "WhenCreated": iso_time(date.fromisoformat(person["JoinedDate"])),
            "DistinguishedName": f"CN={person['DisplayName']},OU=People,DC=contoso,DC=com",
            "AccountType": "Human", "AccountPopulation": "Workforce"})
        output["AD_DirectoryObjects_AllDomains.csv"].append({**identity, "DomainName": "contoso.com",
            "ObjectClass": "user", "DistinguishedName": f"CN={person['DisplayName']},OU=People,DC=contoso,DC=com",
            "ObjectGUID": person["ADGuid"], "ObjectSID": sid, "PrimaryGroupID": "513"})
    for index, device in enumerate(devices):
        person = people_by_id[device["PersonId"]]
        common = {**identity, "Manufacturer": device["Manufacturer"], "Model": device["Model"]}
        os_name = "Windows" if device["Kind"].startswith("windows") else device["Platform"]
        is_encrypted = "false" if device["Compliance"] == "Noncompliant" else "true"
        output["Intune_ManagedDevices_All.csv"].append({**common, "ManagedDeviceId": device["DeviceId"],
            "AzureADDeviceId": device["AzureId"], "DeviceName": device["Alias"], "OperatingSystem": os_name,
            "OsVersion": device["OSBuild"] or device["OSVersion"], "ManagedDeviceOwnerType": "company",
            "ComplianceState": device["Compliance"].lower(), "ManagementAgent": "mdm",
            "DeviceEnrollmentType": "windowsAzureADJoin" if os_name == "Windows" else "userEnrollment",
            "EnrolledDateTime": device["EnrollmentDateTime"], "LastSyncDateTime": device["LastSyncDateTime"],
            "UserId": person["PersonId"], "UserPrincipalName": person["UPN"],
            "SerialNumber": device["SerialNumber"], "IsEncrypted": is_encrypted,
            "TotalStorageSpaceInBytes": int(device["StorageGB"] * 1024**3),
            "FreeStorageSpaceInBytes": int(float(device["FreeDiskGB"]) * 1024**3), "CollectedAtUtc": stamp})
        if device["Kind"].startswith("windows"):
            output["Intune_Devices_Inventory.csv"].append({**common, "Device ID": device["DeviceId"],
                "Device name": device["Alias"], "Enrollment date": device["EnrollmentDateTime"],
                "DeviceEnrollmentType": "windowsAzureADJoin", "Last check-in": device["LastSyncDateTime"],
                "LastSyncDateTime": device["LastSyncDateTime"], "Intune registered": "true",
                "Azure AD Device ID": device["AzureId"], "Azure AD registered": "true", "OS": "Windows",
                "OS version": device["OSBuild"], "Serial number": device["SerialNumber"],
                "Encrypted": is_encrypted, "Compliance": device["Compliance"].lower(),
                "IsCompliant": str(device["Compliance"] == "Compliant").lower(), "Managed by": "mdm",
                "Ownership": "company", "Primary user UPN": person["UPN"],
                "Primary user email address": person["UPN"], "Primary user display name": person["DisplayName"],
                "UserId": person["PersonId"], "Total storage": int(device["StorageGB"] * 1024**3),
                "Free storage": int(float(device["FreeDiskGB"]) * 1024**3), "Category": "PC Windows",
                "PhysicalMemoryGB": 16 if device["Kind"] == "windows11" else 8,
                "Entra DeviceId": device["AzureId"], "Entra ObjectId": deterministic_guid(seed, "entra-object", device["Alias"])})
            observed = min(as_of, date.fromisoformat(device["LastSyncDateTime"][:10]) + timedelta(days=1))
            boot = iso_time(max(date.fromisoformat(device["PurchaseDate"]),
                                observed - timedelta(days=(index % 21) + 1)))
            output["Intune_Devices_LocalSystem.csv"].append({**identity, "DeviceName": device["Alias"],
                "DeviceId": device["DeviceId"], "AzureADDeviceId": device["AzureId"],
                "UserPrincipalName": person["UPN"], "SecureBootStatus": "Disabled" if index % 97 == 0 else "Enabled",
                "BIOSVersion": f"CTBIOS-{1 + index % 8}.0", "BIOSDate": device["PurchaseDate"],
                "FirmwareType": "Uefi", "PlatformScriptLastUpdate": iso_time(observed),
                "LastRebootDate": boot, "Last Reboot Date": boot, "LastBootUpTime": boot, "LastBootTime": boot,
                "IsEncrypted": is_encrypted, "DeviceGuardVBSState": "running" if index % 13 else "notEnabled",
                "DeviceGuardVBSHardwareRequirement": "meetHardwareRequirements",
                "DeviceGuardCredentialGuardState": "running" if index % 17 else "notEnabled",
                "ManagementState": "managed", "ComplianceState": device["Compliance"].lower(),
                "LastSyncDateTime": device["LastSyncDateTime"], "ThirdPartyAVStatus": "NotDetected",
                "ExpectedThirdPartyAVStatus": "NotDetected", "WindowsUpdateRebootRequired": str(index % 11 == 0),
                "BitLockerOSProtectionStatus": "On" if is_encrypted == "true" else "Off",
                "BitLockerOSConversionStatus": "FullyEncrypted" if is_encrypted == "true" else "FullyDecrypted"})
            if device["DirectoryMatch"] == "Matched":
                sid = f"S-1-5-21-20061009-20261009-13000-{1000 + index}"
                dn = f"CN={device['Alias']},OU=Workstations,DC=contoso,DC=com"
                output["AD_Computers_AllDomains.csv"].append({**identity, "DomainName": "contoso.com",
                    "ObjectType": "Computer", "ObjectGUID": device["AzureId"], "Name": device["Alias"],
                    "SamAccountName": device["Alias"] + "$", "DistinguishedName": dn,
                    "Enabled": "true", "OperatingSystem": device["OSFamily"],
                    "operatingSystemVersion": device["OSBuild"], "LastLogonDate": device["LastSyncDateTime"],
                    "WhenCreated": device["EnrollmentDateTime"], "SID": sid,
                    "IntuneDeviceId": device["DeviceId"], "EntraDeviceIdNormalized": device["AzureId"],
                    "ExistsInIntune": "true", "EntitySiteType": "Office"})
                output["AD_DirectoryObjects_AllDomains.csv"].append({**identity, "DomainName": "contoso.com",
                    "ObjectClass": "computer", "DistinguishedName": dn, "ObjectGUID": device["AzureId"],
                    "ObjectSID": sid, "PrimaryGroupID": "515"})
    for index in range(max(1, round(len(people) * .01))):
        guid = deterministic_guid(seed, "ad-service", str(index))
        output["AD_DirectoryObjects_AllDomains.csv"].append({**identity, "DomainName": "contoso.com",
            "ObjectClass": "msDS-GroupManagedServiceAccount", "DistinguishedName": f"CN=CT-SVC-{index+1:04d},OU=Services,DC=contoso,DC=com",
            "ObjectGUID": guid, "ObjectSID": f"S-1-5-21-20061009-20261009-14000-{1000+index}"})
    if include_battery:
        windows = [d for d in devices if d["Kind"].startswith("windows")]
        covered_ids = {d["DeviceId"] for d in sorted(windows, key=lambda d: stable_hex(seed, "battery-coverage", d["DeviceId"]))
                       [:round(len(windows) * battery_coverage / 100)]}
        covered = [d for d in windows if d["DeviceId"] in covered_ids]
        for index, device in enumerate(covered):
            score = 42 if index % 13 == 0 else 74 + index % 24
            output["Intune_EndpointAnalytics_BatteryHealth.csv"].append({**identity,
                "RunId": f"RUN-CONTOSO-{as_of:%Y%m%d}", "ReportName": "BRDeviceBatteryAgg",
                "ReportRefreshDate": stamp, "DeviceId": device["DeviceId"], "DeviceName": device["Alias"],
                "Manufacturer": device["Manufacturer"], "Model": device["Model"],
                "BatteryId": deterministic_guid(seed, "battery", device["Alias"]), "BatteryCount": 1,
                "BatteryHealthScore": score, "BatteryCapacityScore": score,
                "BatteryRuntimeScore": max(0, score - 4), "MaximumCapacity": 88 if score >= 60 else 52,
                "CycleCount": 80 + index % 700, "DesignCapacity": 50000,
                "FullChargeCapacity": 44000 if score >= 60 else 26000,
                "NeedsAttention": str(score < 60).lower()})
        output["Intune_EndpointAnalytics_DataQuality.csv"].append({**identity,
            "RunId": f"RUN-CONTOSO-{as_of:%Y%m%d}", "ReportName": "BRDeviceBatteryAgg",
            "ApiVersion": "beta", "Status": "Collected",
            "RowCount": len(covered), "ExportJobStatus": "completed", "IsAdvancedAnalytics": "true",
            "RequiredPermission": "DeviceManagementManagedDevices.ReadWrite.All", "CollectedAtUtc": stamp,
            "RawRowCount": len(covered), "ExcludedRowCount": 0, "ExcludedDeviceCount": 0})
    return output


def new_prepared_rows(name: str, columns: list[dict], people: list[dict], devices: list[dict],
                      raw: dict[str, list[dict]], as_of: date) -> list[dict]:
    base = base_prepared_row(columns, as_of)
    stamp = iso_time(as_of)
    by_upn = {p["UPN"]: p for p in people}
    by_device = {d["DeviceId"]: d for d in devices}
    if name == "AllAccountsEvidence.csv":
        entra_by_immutable = {r["OnPremisesImmutableId"]: r for r in raw["M365_Users_Active.csv"]}
        rows = []
        for ad in raw["AD_Users_AllDomains.csv"]:
            cloud = entra_by_immutable.get(ad["ImmutableId_AD"])
            if cloud is None:
                raise ValueError("Generated AD user has no strong Entra link")
            rows.append({**base, **synthetic_identity(), "Account Key": "ad:" + ad["ObjectGUID"],
                "Inventory Source": "AD + Entra", "Account Kind": "user", "AD Object GUID": ad["ObjectGUID"],
                "Entra Object ID": cloud["Object Id"], "AD State": "Enabled" if ad["Enabled"] == "true" else "Disabled",
                "Entra State": "Enabled" if cloud["AccountEnabled"] == "true" else "Disabled",
                "Entra User Type": "Member", "Country": cloud["CountryOrRegion"], "Country Basis": "Entra country",
                "Match Basis": "SID + immutable ID", "UPN Comparison": "Same", "AD Observation UTC": stamp,
                "Entra Observation UTC": stamp})
        for item in raw["AD_DirectoryObjects_AllDomains.csv"]:
            if item["ObjectClass"] != "msDS-GroupManagedServiceAccount":
                continue
            rows.append({**base, **synthetic_identity(), "Account Key": "ad:" + item["ObjectGUID"],
                "Inventory Source": "AD only", "Account Kind": item["ObjectClass"],
                "AD Object GUID": item["ObjectGUID"], "AD State": "Unknown", "Entra State": "Unknown",
                "Country": "Unknown", "Country Basis": "Not observed", "Match Basis": "AD only",
                "UPN Comparison": "Not compared", "AD Observation UTC": stamp, "Entra Observation UTC": ""})
        return rows
    if name == "AllDevicesEvidence.csv":
        matched = {r["ObjectGUID"]: r for r in raw["AD_Computers_AllDomains.csv"]}
        output = []
        for row in raw["Intune_ManagedDevices_All.csv"]:
            device = by_device[row["ManagedDeviceId"]]
            ad = matched.get(row["AzureADDeviceId"])
            output.append({**base, **synthetic_identity(),
                "Device Key": ("ad:" + ad["ObjectGUID"]) if ad else ("intune:" + row["ManagedDeviceId"]),
                "Inventory Source": "AD + Intune" if ad else "Intune only",
                "AD Object GUID": ad["ObjectGUID"] if ad else "", "Intune Device ID": row["ManagedDeviceId"],
                "Azure AD Device ID": row["AzureADDeviceId"], "Platform": row["OperatingSystem"],
                "Country": device["Country"], "Country Basis": "Primary-user proxy: Entra country",
                "Match Basis": "AD GUID = Azure AD device ID" if ad else "Intune only",
                "AD Observation UTC": stamp if ad else "", "Intune Observation UTC": stamp})
        return output
    if name == "DeviceSystemEvidence.csv":
        detailed = {r["Device ID"]: r for r in raw["Intune_Devices_Inventory.csv"]}
        output = []
        for system in raw["Intune_Devices_LocalSystem.csv"]:
            did = system["DeviceId"]
            inventory = detailed[did]
            person = by_upn[system["UserPrincipalName"]]
            observation = system["PlatformScriptLastUpdate"]
            age = (as_of - date.fromisoformat(observation[:10])).days * 24
            encrypted = system["IsEncrypted"].title()
            compliance = system["ComplianceState"]
            output.append({**base, "TenantKey": TENANT, "Device System Key": TENANT + "|" + did,
                "Device Source ID": did, "Device Name": system["DeviceName"],
                "Azure AD Device ID": system["AzureADDeviceId"], "Primary User UPN": person["UPN"],
                "Country": person["Country"], "Country Basis": "Primary user CountryOrRegion",
                "Inventory Match": "Matched", "Prepared Inventory Match": "Matched", "Observation Raw": observation,
                "Observation UTC": observation, "Observation Time Basis": "Explicit offset",
                "Observation Age Hours": age, "Freshness": "Within policy" if age <= 720 else "Stale",
                "Freshness Max Age Hours": 720, "Prepared UTC": stamp,
                "Secure Boot State": system["SecureBootStatus"], "Secure Boot Raw": system["SecureBootStatus"],
                "BIOS Version": system["BIOSVersion"], "BIOS Date": system["BIOSDate"],
                "BIOS Date Quality": "Observed date; not a health verdict",
                "Firmware Type": system["FirmwareType"], "Firmware Raw": system["FirmwareType"],
                "Last Boot Time": system["LastRebootDate"], "Last Boot Quality": "Observed local time; time zone unspecified",
                "Local Encryption State": encrypted, "Inventory Encryption State": inventory["Encrypted"].title(),
                "Encryption Comparison": "Agreement", "VBS State": system["DeviceGuardVBSState"],
                "VBS Raw": system["DeviceGuardVBSState"],
                "VBS Hardware Requirements": system["DeviceGuardVBSHardwareRequirement"],
                "VBS Hardware Raw": system["DeviceGuardVBSHardwareRequirement"],
                "Credential Guard State": system["DeviceGuardCredentialGuardState"],
                "Credential Guard Raw": system["DeviceGuardCredentialGuardState"],
                "Management State": system["ManagementState"], "Management Raw": system["ManagementState"],
                "Compliance State": compliance, "Compliance Raw": compliance,
                "Inventory Compliance State": inventory["Compliance"], "Compliance Comparison": "Agreement",
                "Last Sync Raw": system["LastSyncDateTime"], "Last Sync UTC": system["LastSyncDateTime"],
                "Last Sync Time Basis": "Explicit offset", "Third Party AV State": "NotDetected",
                "Third Party AV Raw": "NotDetected", "Expected AV State": "NotDetected",
                "Expected AV Raw": "NotDetected", "AV Qualification": "Expected product not configured",
                "AV Comparison": "No contradiction observed",
                "Windows Update Reboot Required": system["WindowsUpdateRebootRequired"],
                "Windows Update Reboot Raw": system["WindowsUpdateRebootRequired"],
                "BitLocker OS Protection": system["BitLockerOSProtectionStatus"],
                "BitLocker OS Protection Raw": system["BitLockerOSProtectionStatus"],
                "BitLocker OS Conversion": system["BitLockerOSConversionStatus"],
                "BitLocker OS Conversion Raw": system["BitLockerOSConversionStatus"]})
        return output
    raise ValueError(f"Unexpected 1.2 prepared file: {name}")


def write_prepared_v12(root: Path, contract: dict, people: list[dict], devices: list[dict],
                       fixtures: dict[str, list[dict]], raw: dict[str, list[dict]],
                       as_of: date, seed: int, rng: random.Random, previous: str | None) -> dict:
    batch_id = as_of.strftime("%Y%m%dT000000000Z") + "-" + stable_hex(seed, as_of, "batch-v1.2.0", length=8)
    batch = root / "DATA-POWERBI" / "batches" / batch_id
    if previous == batch_id:
        pointer = json.loads((root / "DATA-POWERBI" / "current.json.txt").read_text(encoding="utf-8"))
        previous = pointer.get("PreviousBatchId")
    reports = []
    for table in contract["tables"]:
        name = table["file"]
        headers = (["TenantKey"] if not table.get("stagingIncludesTenantKey") else []) + [c["name"] for c in table["columns"]]
        if name in ("AllAccountsEvidence.csv", "AllDevicesEvidence.csv"):
            headers = RAW_IDENTITY + [c["name"] for c in table["columns"]]
        values = (new_prepared_rows(name, table["columns"], people, devices, raw, as_of)
                  if name in ("DeviceSystemEvidence.csv", "AllAccountsEvidence.csv", "AllDevicesEvidence.csv")
                  else prepared_rows(name, table["columns"], people, devices, fixtures, as_of, rng))
        if not values and not table.get("allowEmpty"):
            raise ValueError(f"Unexpected empty prepared file: {name}")
        for value in values:
            value["TenantKey"] = TENANT
        path = batch / name
        count = write_csv(path, headers, values)
        reports.append({"File": name, "Rows": count, "Columns": len(headers), "Bytes": path.stat().st_size,
                        "SHA256": sha256(path)})
    write_json(batch / "validation.json.txt", {"Passed": True, "SchemaOnly": False, "Files": reports, "Errors": []})
    versions = {table["table"]: table.get("historyKeyVersion", 1) for table in contract["tables"]
                if "History" in table["table"] or "Trend" in table["table"]}
    manifest = {"SchemaVersion": 1, "BatchId": batch_id, "TenantKey": TENANT, "CreatedUtc": iso_time(as_of, 0),
        "Provenance": {"Kind": "SyntheticContosoDemo", "GeneratorVersion": "1.2.0", "Seed": seed,
                       "AsOf": iso_day(as_of), "NoCustomerData": True},
        "ContractSHA256": sha256(CONTRACT), "HistoryKeyVersions": versions, "HistoryKeyResets": [],
        "Files": [{key: item[key] for key in ("File", "Rows", "Bytes", "SHA256")} for item in reports]}
    write_json(batch / "batch.json.txt", manifest)
    pointer = {"SchemaVersion": 1, "BatchId": batch_id, "TenantKey": TENANT,
               "ManifestSHA256": sha256(batch / "batch.json.txt"), "PreviousBatchId": previous}
    write_json(batch / "current.json.txt", pointer)
    write_json(root / "DATA-POWERBI" / "current.json.txt", pointer)
    return {"BatchId": batch_id, "Reports": reports, "PreviousBatchId": previous}


def read_demo_csv(path: Path, encoding: str = "utf-8") -> list[dict]:
    with path.open(encoding=encoding, newline="") as stream:
        return list(csv.DictReader(stream))


def validate_legacy_batch(root: Path, batch_id: str | None, legacy_sha: str | None,
                          contract: dict) -> dict | None:
    if not batch_id:
        return None
    batch = root / "DATA-POWERBI" / "batches" / batch_id
    manifest_path = batch / "batch.json.txt"
    if not manifest_path.is_file() or (legacy_sha and sha256(manifest_path) != legacy_sha):
        raise ValueError("Archived 26-file batch manifest changed")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    tables = contract["tables"][3:]
    if len(manifest["Files"]) != 26 or len(tables) != 26:
        raise ValueError("Archived prepared batch must contain 26 files")
    for table, entry in zip(tables, manifest["Files"]):
        path = batch / entry["File"]
        if entry["File"] != table["file"] or sha256(path) != entry["SHA256"] or path.stat().st_size != entry["Bytes"]:
            raise ValueError(f"Archived prepared hash mismatch: {entry['File']}")
        with path.open(encoding="utf-8", newline="") as stream:
            reader = csv.reader(stream)
            if next(reader) != ["TenantKey"] + [c["name"] for c in table["columns"]]:
                raise ValueError(f"Archived prepared header mismatch: {entry['File']}")
            if sum(1 for _ in reader) != entry["Rows"]:
                raise ValueError(f"Archived prepared count mismatch: {entry['File']}")
    pointer_path = batch / "current.json.txt"
    pointer = json.loads(pointer_path.read_text(encoding="utf-8"))
    if pointer["BatchId"] != batch_id or pointer["ManifestSHA256"] != sha256(manifest_path):
        raise ValueError("Archived 26-file batch pointer changed")
    return {"BatchId": batch_id, "Files": 26, "ContractSHA256": manifest["ContractSHA256"],
            "ManifestSHA256": sha256(manifest_path), "Validated": True}


def validate_v12(root: Path, users: int, volumes: dict[str, int], contract: dict,
                 sp_config: dict[str, int], include_battery: bool, battery_coverage: int,
                 legacy_id: str | None, legacy_sha: str | None) -> dict:
    if len(contract["tables"]) != 29:
        raise ValueError("Current prepared contract must contain 29 files")
    demo_metadata = json.loads((root / "demo-manifest.json.txt").read_text(encoding="utf-8"))
    if (demo_metadata.get("GeneratorVersion") != "1.2.0" or
            demo_metadata.get("PreparedContractSHA256") != sha256(CONTRACT) or
            demo_metadata.get("SourceCoverage", {}).get("SourceContractFiles") != 50 or
            demo_metadata.get("SourceCoverage", {}).get("FullSourceContractCoverage") is not False):
        raise ValueError("Demo manifest version, contract hash or source scope mismatch")
    reference_people = read_demo_csv(root / "reference" / "contoso_people.csv")
    reference_devices = read_demo_csv(root / "reference" / "contoso_assets.csv")
    if len(reference_people) != users or len(reference_devices) != sum(volumes.values()):
        raise ValueError("Reference volume mismatch")
    seed = int(demo_metadata["Seed"])
    people, devices = versioned_entities(reference_people, reference_devices, seed)
    person_ids = {p["PersonId"] for p in people}
    device_ids = {d["DeviceId"] for d in devices}
    upns = {p["UPN"] for p in people}
    countries = {p["Country"] for p in people}
    if len(person_ids) != users or len(device_ids) != len(devices) or len(upns) != users:
        raise ValueError("Duplicate deterministic identity")
    for value in person_ids | device_ids:
        if uuid.UUID(value).version != 5:
            raise ValueError("Expected deterministic version-5 GUID")
    if any(p["LastActivityUtc"][:10] < p["JoinedDate"] for p in people):
        raise ValueError("User activity predates joining")
    if any(d["PersonId"] not in person_ids or d["PrimaryUserUPN"] not in upns or
           d["Country"] not in countries or
           not (d["PurchaseDate"] <= d["EnrollmentDateTime"][:10] <= d["LastSyncDateTime"][:10])
           for d in devices):
        raise ValueError("Device ownership, country or chronology mismatch")
    raw_root = root / "DATA-LAST"
    schemas = source_headers()
    raw = {}
    mandatory = tuple(name for name in schemas if name not in
                      ("Intune_EndpointAnalytics_BatteryHealth.csv", "Intune_EndpointAnalytics_DataQuality.csv"))
    for name in mandatory + (("Intune_EndpointAnalytics_BatteryHealth.csv", "Intune_EndpointAnalytics_DataQuality.csv")
                             if include_battery else ()):
        path = raw_root / name
        data = path.read_bytes()
        if data.startswith(b"\xef\xbb\xbf") or not data.startswith(b'"TenantKey",') or b"\r\n" not in data:
            raise ValueError(f"Raw source format mismatch: {name}")
        with path.open(encoding="utf-8", newline="") as stream:
            reader = csv.DictReader(stream)
            if reader.fieldnames != schemas[name]:
                raise ValueError(f"Raw source header mismatch: {name}")
            records = list(reader)
        if any(any(row[key] != value for key, value in synthetic_identity().items()) for row in records):
            raise ValueError(f"Raw source tenant mismatch: {name}")
        raw[name] = records
    if not include_battery and any((raw_root / name).exists() for name in
                                   ("Intune_EndpointAnalytics_BatteryHealth.csv", "Intune_EndpointAnalytics_DataQuality.csv")):
        raise ValueError("Battery files exist but battery option is disabled")
    if len(raw["M365_Users_Active.csv"]) != users or len(raw["AD_Users_AllDomains.csv"]) != users or len(raw["Intune_ManagedDevices_All.csv"]) != len(devices):
        raise ValueError("Primary source volume mismatch")
    for name, field in (("M365_Users_Active.csv", "Object Id"),
                        ("AD_Users_AllDomains.csv", "ObjectGUID"),
                        ("AD_DirectoryObjects_AllDomains.csv", "ObjectGUID"),
                        ("Intune_ManagedDevices_All.csv", "ManagedDeviceId"),
                        ("Intune_ManagedDevices_All.csv", "AzureADDeviceId"),
                        ("Intune_Devices_LocalSystem.csv", "DeviceId")):
        if any(uuid.UUID(row[field]).version != 5 for row in raw[name]):
            raise ValueError(f"Invalid synthetic GUID: {name}.{field}")
    if (any(not row["User principal name"].endswith("@contoso.com") or
            row["CountryOrRegion"] not in countries for row in raw["M365_Users_Active.csv"]) or
            any(row["DomainName"] != "contoso.com" or not row["DistinguishedName"].endswith("DC=contoso,DC=com")
                for row in raw["AD_DirectoryObjects_AllDomains.csv"])):
        raise ValueError("Non-Contoso source domain or country")
    for row in raw["Intune_ManagedDevices_All.csv"]:
        if (not row["EnrolledDateTime"] <= row["LastSyncDateTime"] <= row["CollectedAtUtc"] or
                int(row["TotalStorageSpaceInBytes"]) <= 0 or
                not 0 <= int(row["FreeStorageSpaceInBytes"]) <= int(row["TotalStorageSpaceInBytes"])):
            raise ValueError("Managed-device source date or storage mismatch")
    windows = [d for d in devices if d["Kind"].startswith("windows")]
    if len(raw["Intune_Devices_Inventory.csv"]) != len(windows) or len(raw["Intune_Devices_LocalSystem.csv"]) != len(windows):
        raise ValueError("Windows source coverage mismatch")
    ad_computers = [d for d in windows if d["DirectoryMatch"] == "Matched"]
    if len(raw["AD_Computers_AllDomains.csv"]) != len(ad_computers) or len(raw["AD_DirectoryObjects_AllDomains.csv"]) != users + len(ad_computers) + max(1, round(users * .01)):
        raise ValueError("AD directory coverage mismatch")
    cloud_by_id = {r["Object Id"]: r for r in raw["M365_Users_Active.csv"]}
    ad_by_id = {r["ObjectGUID"]: r for r in raw["AD_Users_AllDomains.csv"]}
    managed_by_id = {r["ManagedDeviceId"]: r for r in raw["Intune_ManagedDevices_All.csv"]}
    detailed_by_id = {r["Device ID"]: r for r in raw["Intune_Devices_Inventory.csv"]}
    local_by_id = {r["DeviceId"]: r for r in raw["Intune_Devices_LocalSystem.csv"]}
    if (len(cloud_by_id) != users or len(ad_by_id) != users or len(managed_by_id) != len(devices) or
            len(detailed_by_id) != len(windows) or len(local_by_id) != len(windows)):
        raise ValueError("Duplicate source identity")
    for person in people:
        cloud = cloud_by_id[person["PersonId"]]
        ad = ad_by_id[person["ADGuid"]]
        expected_immutable = base64.b64encode(uuid.UUID(person["ADGuid"]).bytes_le).decode("ascii")
        if (cloud["OnPremisesImmutableId"] != ad["ImmutableId_AD"] or ad["ImmutableId_AD"] != expected_immutable or
                cloud["OnPremisesSecurityIdentifier"] != ad["ObjectSID"] or cloud["User principal name"] != person["UPN"] or
                cloud["CountryOrRegion"] != person["Country"]):
            raise ValueError("AD and Entra strong user join mismatch")
    for device in devices:
        managed = managed_by_id[device["DeviceId"]]
        if (managed["AzureADDeviceId"] != device["AzureId"] or managed["UserId"] != device["PersonId"] or
                managed["UserPrincipalName"] != device["PrimaryUserUPN"]):
            raise ValueError("Intune managed device join mismatch")
        if device["Kind"].startswith("windows"):
            detail = detailed_by_id[device["DeviceId"]]
            local = local_by_id[device["DeviceId"]]
            if detail["Azure AD Device ID"] != managed["AzureADDeviceId"] or local["AzureADDeviceId"] != managed["AzureADDeviceId"]:
                raise ValueError("Intune detailed and local-system join mismatch")
            if not (device["PurchaseDate"] <= local["BIOSDate"] <= local["LastRebootDate"][:10] <= local["PlatformScriptLastUpdate"][:10]):
                raise ValueError("Local-system chronology mismatch")
    if include_battery:
        expected_battery = round(len(windows) * battery_coverage / 100)
        battery = raw["Intune_EndpointAnalytics_BatteryHealth.csv"]
        quality = raw["Intune_EndpointAnalytics_DataQuality.csv"]
        if (len(battery) != expected_battery or len({r["DeviceId"] for r in battery}) != len(battery) or
                any(r["DeviceId"] not in detailed_by_id for r in battery) or len(quality) != 1 or
                quality[0]["ReportName"] != "BRDeviceBatteryAgg" or quality[0]["Status"] != "Collected" or
                int(quality[0]["RowCount"]) != len(battery) or int(quality[0]["ExcludedDeviceCount"]) != 0):
            raise ValueError("Battery coverage or DataQuality mismatch")
    pointer = json.loads((root / "DATA-POWERBI" / "current.json.txt").read_text(encoding="utf-8"))
    batch = root / "DATA-POWERBI" / "batches" / pointer["BatchId"]
    manifest_path = batch / "batch.json.txt"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if (pointer["ManifestSHA256"] != sha256(manifest_path) or manifest["ContractSHA256"] != sha256(CONTRACT) or
            len(manifest["Files"]) != 29 or pointer["PreviousBatchId"] != legacy_id):
        raise ValueError("Current 29-file batch metadata mismatch")
    prepared_counts = {}
    for table, entry in zip(contract["tables"], manifest["Files"]):
        name = table["file"]
        path = batch / name
        headers = (["TenantKey"] if not table.get("stagingIncludesTenantKey") else []) + [c["name"] for c in table["columns"]]
        if name in ("AllAccountsEvidence.csv", "AllDevicesEvidence.csv"):
            headers = RAW_IDENTITY + [c["name"] for c in table["columns"]]
        if entry["File"] != name or entry["SHA256"] != sha256(path) or entry["Bytes"] != path.stat().st_size:
            raise ValueError(f"Prepared file hash mismatch: {name}")
        count, keys = 0, set()
        with path.open(encoding="utf-8", newline="") as stream:
            reader = csv.DictReader(stream)
            if reader.fieldnames != headers:
                raise ValueError(f"Prepared header mismatch: {name}")
            for row in reader:
                count += 1
                if row["TenantKey"] != TENANT:
                    raise ValueError(f"Prepared tenant mismatch: {name}")
                if table.get("uniqueKey"):
                    key = tuple(row[field].casefold() for field in table["uniqueKey"])
                    if not all(key) or key in keys:
                        raise ValueError(f"Prepared unique key mismatch: {name}")
                    keys.add(key)
                for col in table["columns"]:
                    value = row[col["name"]]
                    if not value:
                        continue
                    kind = col["type"]
                    if kind in ("int64", "double", "decimal"):
                        number = float(value)
                        if not math.isfinite(number) or kind == "int64" and not number.is_integer():
                            raise ValueError(f"Prepared numeric type mismatch: {name}")
                    elif kind == "dateTime":
                        try:
                            date.fromisoformat(value[:10])
                        except ValueError as exc:
                            raise ValueError(f"Prepared datetime type mismatch: {name}") from exc
                    elif kind == "boolean" and value.lower() not in ("true", "false"):
                        raise ValueError(f"Prepared boolean type mismatch: {name}")
        if count != entry["Rows"] or count == 0 and not table.get("allowEmpty"):
            raise ValueError(f"Prepared row count mismatch: {name}")
        prepared_counts[name] = count
    if (prepared_counts["AllAccountsEvidence.csv"] != users + max(1, round(users * .01)) or
            prepared_counts["AllDevicesEvidence.csv"] != len(devices) or
            prepared_counts["DeviceSystemEvidence.csv"] != len(windows) or
            prepared_counts["WindowsMigrationEvidence.csv"] != volumes["windows10"]):
        raise ValueError("New prepared population mismatch")
    accounts = read_demo_csv(batch / "AllAccountsEvidence.csv")
    all_devices = read_demo_csv(batch / "AllDevicesEvidence.csv")
    system = read_demo_csv(batch / "DeviceSystemEvidence.csv")
    if (any(r["Entra Object ID"] and r["Entra Object ID"] not in cloud_by_id for r in accounts) or
            any(r["Intune Device ID"] not in managed_by_id for r in all_devices) or
            any(r["Device Source ID"] not in local_by_id for r in system)):
        raise ValueError("Prepared-to-source identity mismatch")
    legacy = validate_legacy_batch(root, legacy_id, legacy_sha, contract)
    sharepoint = validate_sharepoint_onprem(root, reference_people,
                                           date.fromisoformat(demo_metadata["AsOf"]), sp_config)
    with (raw_root / "SharePoint_OnPrem_SiteCollections.csv").open(encoding="utf-8-sig", newline="") as stream:
        site_reader = csv.DictReader(stream, delimiter=";")
        if any(row["SiteCategoryStatus"] != "Observed" for row in site_reader):
            raise ValueError("SharePoint SiteCategoryStatus must be Observed")
    return {"Passed": True, "Users": users, "Devices": len(devices), "PreparedFiles": 29,
            "RawSources": {name: len(rows) for name, rows in sorted(raw.items())},
            "Windows10MigrationRows": prepared_counts["WindowsMigrationEvidence.csv"],
            "DirectoryOrphans": sum(d["DirectoryMatch"] == "Missing" for d in devices),
            "LegacyPrepared": legacy, "SharePointOnPrem": sharepoint,
            "BatteryHealth": {"Enabled": include_battery, "CoveredDevices": len(raw.get("Intune_EndpointAnalytics_BatteryHealth.csv", [])),
                              "WindowsDevices": len(windows)}}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--users", type=int, default=12000)
    parser.add_argument("--windows11", type=int, default=12000)
    parser.add_argument("--windows10", type=int, default=1000)
    parser.add_argument("--mac", type=int, default=500)
    parser.add_argument("--android", type=int, default=3000)
    parser.add_argument("--iphone", type=int, default=1500)
    parser.add_argument("--tablets", type=int, default=0)
    parser.add_argument("--seed", type=int, default=20261009)
    parser.add_argument("--as-of", type=date.fromisoformat, default=date(2026, 10, 1))
    parser.add_argument("--output-root", type=Path, default=PRODUCT / "_private" / "demo" / "Contoso")
    parser.add_argument("--validate-only", action="store_true")
    parser.add_argument("--upgrade-v1-2", action="store_true",
                        help="Extend the existing Contoso demo while preserving its archived 26-file batch")
    parser.add_argument("--include-battery-health", action="store_true",
                        help="Emit optional BatteryHealth and DataQuality sources outside the prepared batch")
    parser.add_argument("--battery-coverage-percent", type=int, default=70)
    parser.add_argument("--extend-sharepoint-onprem", action="store_true",
                        help="Add SharePoint CSVs to an existing Contoso demo without rewriting the prepared evidence")
    parser.add_argument("--sp-member-servers", type=int, default=8)
    parser.add_argument("--sp-external-servers", type=int, default=2)
    parser.add_argument("--sp-service-applications", type=int, default=10)
    parser.add_argument("--sp-web-applications", type=int, default=5)
    parser.add_argument("--sp-content-databases", type=int, default=25)
    parser.add_argument("--sp-published-collections", type=int, default=800)
    parser.add_argument("--sp-excluded-personal", type=int, default=350)
    parser.add_argument("--sp-metadata-incomplete", type=int, default=10)
    parser.add_argument("--sp-database-mismatches", type=int, default=2)
    parser.add_argument("--sp-webs-per-collection", type=int, default=2)
    parser.add_argument("--sp-administrators-per-collection", type=int, default=2)
    args = parser.parse_args()
    volumes = {kind: getattr(args, kind) for _, kind in OS_COUNTS}
    volumes["tablets"] = args.tablets
    if args.users <= 0 or any(count < 0 for count in volumes.values()):
        raise SystemExit("User volume must be positive and device volumes must be nonnegative")
    if not 0 <= args.battery_coverage_percent <= 100:
        raise SystemExit("Battery coverage percent must be between 0 and 100")
    contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
    root = args.output_root.resolve()
    allowed = (PRODUCT / "_private" / "demo" / "Contoso").resolve()
    if root != allowed and allowed not in root.parents:
        raise SystemExit(f"Output root must be inside the private Contoso demo folder: {allowed}")
    sp_config = sharepoint_config(args)
    if args.validate_only:
        manifest = verify_demo_manifest(root)
        legacy = manifest.get("LegacyPrepared") or {}
        result = validate_v12(root, args.users, volumes, contract, sp_config,
                              bool(manifest.get("BatteryHealth", {}).get("Enabled")),
                              int(manifest.get("BatteryHealth", {}).get("CoveragePercent", 70)),
                              legacy.get("BatchId"), legacy.get("ManifestSHA256"))
        if manifest["Validation"] != result or manifest.get("SharePointOnPrem", {}).get("ContentStatus") != result["SharePointOnPrem"]["ContentStatus"]:
            raise ValueError("Contoso demo manifest validation or SharePoint status mismatch")
        print(json.dumps(result, sort_keys=True))
        return
    if args.extend_sharepoint_onprem:
        raise SystemExit("Use --upgrade-v1-2 for the current 29-file contract")
    if root.exists() and not args.upgrade_v1_2:
        raise SystemExit(f"Output root already exists; use --upgrade-v1-2 or choose a new empty path: {root}")
    old_manifest = verify_demo_manifest(root) if root.exists() else None
    if old_manifest and (old_manifest["Seed"] != args.seed or old_manifest["AsOf"] != iso_day(args.as_of)):
        raise ValueError("Existing demo seed or date differs from requested parameters")
    if old_manifest and old_manifest.get("BatteryHealth", {}).get("Enabled") and not args.include_battery_health:
        raise ValueError("Existing battery sources require --include-battery-health on regeneration")
    root.mkdir(parents=True, exist_ok=True)
    rng = random.Random(args.seed)
    people, sites = make_people(args.users, args.as_of, rng)
    devices = make_devices(people, volumes, args.as_of, args.seed, rng)
    if old_manifest:
        recorded_people = read_demo_csv(root / "reference" / "contoso_people.csv")
        recorded_devices = read_demo_csv(root / "reference" / "contoso_assets.csv")
        if [p["PersonId"] for p in people] != [p["PersonId"] for p in recorded_people] or [d["DeviceId"] for d in devices] != [d["DeviceId"] for d in recorded_devices]:
            raise ValueError("Existing CT aliases differ from requested seed or volumes")
        legacy = old_manifest.get("LegacyPrepared")
        if old_manifest.get("GeneratorVersion") == "1.2.0" and not legacy:
            legacy_id, legacy_sha = None, None
        elif not legacy:
            old_pointer = json.loads((root / "DATA-POWERBI" / "current.json.txt").read_text(encoding="utf-8"))
            legacy = validate_legacy_batch(root, old_pointer["BatchId"], old_pointer["ManifestSHA256"], contract)
            legacy_id, legacy_sha = legacy["BatchId"], legacy["ManifestSHA256"]
        else:
            validate_legacy_batch(root, legacy["BatchId"], legacy["ManifestSHA256"], contract)
            legacy_id, legacy_sha = legacy["BatchId"], legacy["ManifestSHA256"]
    else:
        core_rows(people, sites, devices, root)
        legacy_id, legacy_sha = None, None
    original_fixtures = fixture_rows(people, devices, args.as_of, rng)
    if not old_manifest:
        write_fixtures(root, original_fixtures, args.as_of)
        write_license_report(root, people, original_fixtures, args.as_of, args.seed)
    new_people, new_devices = versioned_entities(people, devices, args.seed)
    new_rng = random.Random(args.seed)
    new_fixtures = fixture_rows(new_people, new_devices, args.as_of, new_rng)
    raw = source_rows(new_people, new_devices, args.as_of, args.seed,
                      args.include_battery_health, args.battery_coverage_percent)
    schemas = source_headers()
    for name, rows in sorted(raw.items()):
        write_source_csv(root / "DATA-LAST" / name, schemas[name], rows)
    prepared = write_prepared_v12(root, contract, new_people, new_devices, new_fixtures, raw,
                                   args.as_of, args.seed, new_rng, legacy_id)
    sharepoint = write_sharepoint_onprem(root, people, args.as_of, args.seed, sp_config)
    manifest = {"SyntheticOnly": True, "Company": "Contoso", "GeneratorVersion": "1.2.0",
                "Seed": args.seed, "AsOf": iso_day(args.as_of), "BatchId": prepared["BatchId"],
                "PreparedContractSHA256": sha256(CONTRACT), "LegacyPrepared":
                    {"BatchId": legacy_id, "ManifestSHA256": legacy_sha, "Files": 26, "Validated": True} if legacy_id else None,
                "BatteryHealth": {"Enabled": args.include_battery_health,
                                  "CoveragePercent": args.battery_coverage_percent},
                "SourceCoverage": {"GeneratedFiles": sorted(raw), "SourceContractFiles": 50,
                                   "FullSourceContractCoverage": False},
                "SharePointOnPrem": sharepoint, "Validation": {}, "Files": manifest_files(root)}
    write_json(root / "demo-manifest.json.txt", manifest)
    result = validate_v12(root, args.users, volumes, contract, sp_config, args.include_battery_health,
                          args.battery_coverage_percent, legacy_id, legacy_sha)
    manifest["Validation"] = result
    write_json(root / "demo-manifest.json.txt", manifest)
    verify_demo_manifest(root)
    print(json.dumps({"OutputRoot": str(root), **result}, sort_keys=True))


if __name__ == "__main__":
    main()
