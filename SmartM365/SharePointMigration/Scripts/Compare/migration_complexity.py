__version__ = "1.0.0"
"""Score the intrinsic complexity of a migration from its source scans only.

The score is independent from comparison progress: it describes the source estate
(volume, history, paths, risky files and permissions). Thresholds and weights come
from Config/migration-complexity.json.txt, self-healed from its committed template.
"""
import argparse
import builtins
import csv
import hashlib
import json
import math
import os
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

# The embedded Portable Python omits the script directory from sys.path.
sys.path.insert(0, str(Path(__file__).resolve().parent))
from compare_sp_source_target_file_inventories import (
    apply_path_mappings,
    decode_sharegate_path,
    detect_csv_dialect,
    inventory_exclusion_reason,
    load_path_mappings,
    normalize_sharegate_path_characters,
    strip_query_from_path,
)
from compare_sp_source_target_permissions import (
    load_entra_user_aliases,
    normalize_principal_text,
    parse_bool,
    split_principal_member_values,
)

SCHEMA_VERSION = 1
CONFIG_NAME = "migration-complexity"
FILE_CRITERIA = {"FileCount", "TotalSizeGiB", "SubsiteCount", "LibraryCount", "AverageVersions",
                 "PathLimitFiles", "LargeFiles", "RiskyFiles"}
PERMISSION_CRITERIA = {"BrokenInheritance", "DirectUsers"}
ENTRA_CRITERIA = {"InactiveAccounts"}
SYSTEM_IDENTITY_PREFIXES = ("sharepoint\\", "nt authority\\")
DOTNET_EPOCH_TICKS = 621355968000000000


def print(*args, **kwargs):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    args = (f"{timestamp} {args[0]}", *args[1:]) if args else (timestamp,)
    kwargs.setdefault("flush", True)
    builtins.print(*args, **kwargs)


def atomic_text(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_text(value, encoding="utf-8")
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


def merge_keyed_list(default_items, current_items, key):
    """Keep local entries, self-heal their missing keys and append template entries they lack."""
    defaults = {item.get(key): item for item in default_items}
    merged = [merge_missing(defaults.get(item.get(key), {}), item) for item in current_items]
    existing = {item.get(key) for item in merged}
    return merged + [item for item in default_items if item.get(key) not in existing]


def load_config(config_dir):
    config_dir = Path(config_dir)
    template = config_dir / f"{CONFIG_NAME}.json.template"
    runtime = config_dir / f"{CONFIG_NAME}.json.txt"
    default = json.loads(template.read_text(encoding="utf-8-sig"))
    current = json.loads(runtime.read_text(encoding="utf-8-sig")) if runtime.exists() else None
    if current is None:
        merged = default
    else:
        merged = merge_missing(default, current)
        for name, key in (("Criteria", "Id"), ("Levels", "Name"), ("FloorRules", "CriterionId")):
            merged[name] = merge_keyed_list(default.get(name, []), current.get(name, []), key)
    if merged != current:
        atomic_text(runtime, json.dumps(merged, ensure_ascii=False, indent=2) + "\n")
    validate_config(merged)
    return merged, runtime


def validate_config(config):
    criteria = config.get("Criteria") or []
    if not criteria:
        raise ValueError("Complexity configuration defines no criteria.")
    for criterion in criteria:
        bounds = [criterion.get(name) for name in ("OnePointFrom", "TwoPointsAbove", "ThreePointsAbove")]
        if any(not isinstance(value, (int, float)) for value in bounds) or not bounds[0] <= bounds[1] <= bounds[2]:
            raise ValueError(f"Complexity criterion {criterion.get('Id')} needs ordered numeric thresholds.")
        if not isinstance(criterion.get("Weight"), (int, float)) or criterion["Weight"] < 0:
            raise ValueError(f"Complexity criterion {criterion.get('Id')} needs a non-negative weight.")
    levels = config.get("Levels") or []
    if not levels or min(level.get("MinScore", 0) for level in levels) != 0:
        raise ValueError("Complexity levels must start at MinScore 0.")


def points_for(value, criterion):
    if value < criterion["OnePointFrom"]:
        return 0
    if value <= criterion["TwoPointsAbove"]:
        return 1
    if value <= criterion["ThreePointsAbove"]:
        return 2
    return 3


def format_number(value):
    if isinstance(value, float) and not value.is_integer():
        return f"{value:,.2f}"
    return f"{int(value):,}"


def band_texts(criterion):
    one, two, three = (format_number(criterion[name]) for name in ("OnePointFrom", "TwoPointsAbove", "ThreePointsAbove"))
    return [f"< {one}", f"{one} - {two}", f"> {two} - {three}", f"> {three}"]


def read_rows(path):
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        yield from csv.DictReader(handle, dialect=detect_csv_dialect(handle))


def to_number(value):
    text = str(value or "").strip().replace(",", ".")
    try:
        return float(text) if text else 0.0
    except ValueError:
        return 0.0


def site_path(url):
    path = urlparse(str(url or "").strip()).path if url else ""
    return "/" + path.strip("/") if path.strip("/") else ""


def projected_target_path(server_relative_url, mappings, target_site_path):
    text = normalize_sharegate_path_characters(decode_sharegate_path(strip_query_from_path(server_relative_url)))
    if not text.startswith("/"):
        text = "/" + text
    mapped = apply_path_mappings(text, mappings)
    if mapped == text and target_site_path and not text.lower().startswith(target_site_path.lower() + "/"):
        mapped = target_site_path + text
    while "//" in mapped:
        mapped = mapped.replace("//", "/")
    return mapped.lstrip("/")


def file_metrics(path, settings, mappings, target_site_path):
    large_bytes = float(settings.get("LargeFileGiB", 2)) * 1024 ** 3
    path_limit = int(settings.get("PathLimit", 400))
    segment_limit = int(settings.get("SegmentLimit", 255))
    files = 0
    size = 0.0
    versions = 0.0
    webs = set()
    libraries = set()
    large = checked_out = excluded = 0
    target_paths = set()
    over_limit = 0
    for row in read_rows(path):
        files += 1
        row_size = to_number(row.get("SizeBytes"))
        size += row_size
        versions += to_number(row.get("VersionsCount"))
        web = str(row.get("WebUrl") or "").strip().rstrip("/").lower()
        if web:
            webs.add(web)
        library = str(row.get("LibraryUrl") or "").strip().rstrip("/").lower()
        if library:
            libraries.add((web, library))
        if row_size > large_bytes:
            large += 1
        risk_type, _ = inventory_exclusion_reason(row)
        if risk_type:
            excluded += 1
        elif str(row.get("CheckedOutBy") or "").strip():
            checked_out += 1
        if risk_type:
            continue
        target = projected_target_path(row.get("ServerRelativeUrl") or row.get("FileUrl") or "", mappings, target_site_path)
        if not target or target.lower() in target_paths:
            continue
        target_paths.add(target.lower())
        if len(target) > path_limit or max((len(part) for part in target.split("/")), default=0) > segment_limit:
            over_limit += 1
    return {
        "FileCount": (files, f"{files:,} inventory rows"),
        "TotalSizeGiB": (round(size / 1024 ** 3, 2), f"{size:,.0f} bytes (current versions only)"),
        "SubsiteCount": (len(webs), "Distinct WebUrl values"),
        "LibraryCount": (len(libraries), "Distinct LibraryUrl values"),
        "AverageVersions": (round(versions / files, 2) if files else 0.0, "AVG(VersionsCount)"),
        "PathLimitFiles": (over_limit, f"Projected target path > {path_limit} characters or segment > {segment_limit}"),
        "LargeFiles": (large, f"SizeBytes > {format_number(float(settings.get('LargeFileGiB', 2)))} GiB"),
        "RiskyFiles": (checked_out + excluded, f"Checked out: {checked_out:,}; blocked or excluded: {excluded:,}"),
    }


def is_person_identity(value):
    return bool(value) and ("@" in value or "\\" in value) and not value.startswith(SYSTEM_IDENTITY_PREFIXES)


def permission_metrics(path, settings, entra_users_csv):
    scopes = {str(scope).strip().lower() for scope in settings.get("BrokenInheritanceScopes", ["Item"])}
    broken = set()
    direct_users = set()
    identities = set()
    for row in read_rows(path):
        scope = str(row.get("ObjectScope") or "").strip().lower()
        if scope in scopes and parse_bool(row.get("HasUniqueRoleAssignments")) is True:
            broken.add(str(row.get("ObjectId") or row.get("ObjectUrl") or "").strip().lower())
        principal_type = str(row.get("PrincipalType") or "").strip().lower()
        if principal_type == "user":
            identity = normalize_principal_text(row.get("PrincipalLoginName") or row.get("PrincipalName"))
            if is_person_identity(identity):
                identities.add(identity)
                if parse_bool(row.get("IsLimitedAccessOnly")) is not True:
                    direct_users.add(identity)
        elif principal_type == "sharepointgroup":
            for login in split_principal_member_values(row.get("PrincipalMemberLoginNames")):
                identity = normalize_principal_text(login)
                if is_person_identity(identity):
                    identities.add(identity)
    broken.discard("")
    metrics = {
        "BrokenInheritance": (len(broken), f"Distinct objects with unique permissions in scopes: {', '.join(sorted(scopes))}"),
        "DirectUsers": (len(direct_users), "Distinct user principals assigned outside groups (limited access excluded)"),
    }
    if entra_users_csv:
        aliases, disabled_users = load_entra_user_aliases(entra_users_csv)
        unresolved = sum(1 for identity in identities if identity not in aliases)
        disabled = sum(1 for identity in identities if identity in aliases and aliases[identity] in disabled_users)
        metrics["InactiveAccounts"] = (unresolved + disabled,
                                       f"Disabled in Entra: {disabled:,}; not found in Entra: {unresolved:,}; identities checked: {len(identities):,}")
    return metrics


def file_fingerprint(path):
    if not path:
        return None
    item = Path(path)
    stat = item.stat()
    return {"Name": item.name, "Length": stat.st_size,
            "LastWriteTimeUtcTicks": stat.st_mtime_ns // 100 + DOTNET_EPOCH_TICKS}


def level_for(score, levels):
    ordered = sorted(levels, key=lambda level: level["MinScore"])
    chosen = ordered[0]
    for level in ordered:
        if score >= level["MinScore"]:
            chosen = level
    return chosen


def score_migration(config, values):
    criteria_rows = []
    weighted = 0.0
    evaluated_weight = 0.0
    for criterion in config["Criteria"]:
        measured = values.get(criterion["Id"])
        bands = band_texts(criterion)
        row = {
            "Id": criterion["Id"], "Axis": criterion.get("Axis", ""), "Name": criterion.get("Name", criterion["Id"]),
            "Unit": criterion.get("Unit", ""), "Weight": criterion["Weight"], "Bands": bands,
            "Evaluated": measured is not None, "Value": None, "ValueDisplay": "Not evaluated",
            "Points": None, "WeightedPoints": None, "MaxWeightedPoints": criterion["Weight"] * 3,
            "Band": "Not evaluated", "Detail": "",
        }
        if measured is not None:
            value, detail = measured
            points = points_for(value, criterion)
            row.update({
                "Value": value, "ValueDisplay": f"{format_number(value)} {row['Unit']}".strip(),
                "Points": points, "WeightedPoints": points * criterion["Weight"],
                "Band": f"{points} pt ({bands[points]})", "Detail": detail,
            })
            weighted += points * criterion["Weight"]
            evaluated_weight += criterion["Weight"]
        criteria_rows.append(row)

    score = int(math.floor(weighted / (3 * evaluated_weight) * 100 + 0.5)) if evaluated_weight else 0
    level = level_for(score, config["Levels"])
    floors = []
    level_rank = {item["Name"]: item["MinScore"] for item in config["Levels"]}
    by_id = {row["Id"]: row for row in criteria_rows}
    for rule in config.get("FloorRules", []):
        row = by_id.get(rule.get("CriterionId"))
        minimum = rule.get("MinimumLevel")
        if not row or row["Points"] is None or minimum not in level_rank or row["Points"] < rule.get("MinPoints", 3):
            continue
        floors.append(row["Id"])
        if level_rank[minimum] > level["MinScore"]:
            level = next(item for item in config["Levels"] if item["Name"] == minimum)

    drivers = sorted((row for row in criteria_rows if row["WeightedPoints"]),
                     key=lambda row: (-row["WeightedPoints"], -row["Weight"], row["Id"]))[:3]
    return {
        "Score": score, "Level": level["Name"], "LevelColor": level.get("Color", "#1F2937"),
        "LevelBackground": level.get("Background", "#F0F4F8"),
        "ScoreLevel": level_for(score, config["Levels"])["Name"], "FloorApplied": floors,
        "EvaluatedWeight": evaluated_weight, "TotalWeight": sum(item["Weight"] for item in config["Criteria"]),
        "TopDrivers": [row["Id"] for row in drivers], "Criteria": criteria_rows,
    }


def compute(args):
    config, runtime = load_config(args.config_dir)
    settings = config.get("Settings", {})
    mappings = load_path_mappings(args.path_mapping_file) if args.path_mapping_file else []
    values = dict(file_metrics(args.file_inventory, settings, mappings, site_path(args.target_site_url)))
    missing = []
    if args.permission_inventory:
        values.update(permission_metrics(args.permission_inventory, settings, args.entra_users_csv))
        if not args.entra_users_csv:
            missing.append("Entra users cache")
    else:
        missing += ["Source permission inventory", "Entra users cache"] if not args.entra_users_csv else ["Source permission inventory"]
    result = {
        "SchemaVersion": SCHEMA_VERSION,
        "ToolVersion": __version__,
        "Migration": args.migration_name,
        "ComputedAtUtc": datetime.now(timezone.utc).isoformat(),
        "Inputs": {
            "FileInventory": file_fingerprint(args.file_inventory),
            "PermissionInventory": file_fingerprint(args.permission_inventory),
            "EntraUsers": file_fingerprint(args.entra_users_csv) if args.permission_inventory else None,
            "PathMapping": file_fingerprint(args.path_mapping_file),
            "TargetSiteUrl": args.target_site_url or "",
            "ConfigSha256": hashlib.sha256(runtime.read_bytes()).hexdigest().upper(),
        },
        "Status": "Partial" if missing else "Complete",
        "MissingInputs": missing,
    }
    result.update(score_migration(config, values))
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description="Score migration complexity from source scans.")
    parser.add_argument("--migration-name", required=True)
    parser.add_argument("--file-inventory", required=True)
    parser.add_argument("--permission-inventory")
    parser.add_argument("--entra-users-csv")
    parser.add_argument("--path-mapping-file")
    parser.add_argument("--target-site-url")
    parser.add_argument("--config-dir", default=str(Path(__file__).resolve().parents[2] / "Config"))
    parser.add_argument("--output", required=True)
    args = parser.parse_args(argv)
    result = compute(args)
    atomic_text(args.output, json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(f"Complexity {result['Level']} ({result['Score']}/100) for {args.migration_name}: {args.output}")


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from console_lifecycle import run_console_script

    run_console_script(main, __file__, __version__)
