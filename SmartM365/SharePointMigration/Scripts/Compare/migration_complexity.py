__version__ = "1.2.1"
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
    if current is not None and current.get("ScoringRevision", 1) < 2:
        # Upgrade inherited old weights once; explicit non-default weights remain local.
        old_weights = {"TotalSizeGiB": 15, "BrokenInheritance": 15, "DirectUsers": 10}
        old_names = {"FileCount": "Number of files", "AverageVersions": "Average versions per file"}
        defaults_by_id = {item["Id"]: item for item in default["Criteria"]}
        for item in current.get("Criteria", []):
            if item.get("Id") in old_weights and item.get("Weight") == old_weights[item["Id"]]:
                item["Weight"] = defaults_by_id[item["Id"]]["Weight"]
            if item.get("Id") in old_names and item.get("Name") == old_names[item["Id"]]:
                item["Name"] = defaults_by_id[item["Id"]]["Name"]
        current["ScoringRevision"] = 2
        # Keep a separate snapshot so migrated values are persisted even if no keys are missing.
        original = json.loads(runtime.read_text(encoding="utf-8-sig"))
    else:
        original = current
    if current is None:
        merged = default
    else:
        merged = merge_missing(default, current)
        for name, key in (("Criteria", "Id"), ("Levels", "Name"), ("FloorRules", "CriterionId")):
            merged[name] = merge_keyed_list(default.get(name, []), current.get(name, []), key)
    if merged != original:
        # The shared file can be read by another scan or GUI at the same time; the merged
        # values are used in memory and the next calculation retries the self-heal.
        try:
            atomic_text(runtime, json.dumps(merged, ensure_ascii=False, indent=2) + "\n")
        except OSError as error:
            builtins.print(f"WARNING Could not update {runtime}: {error}", file=sys.stderr, flush=True)
            if not runtime.exists():
                raise
    validate_config(merged)
    return merged, runtime


def validate_config(config):
    criteria = config.get("Criteria") or []
    if not criteria:
        raise ValueError("Complexity configuration defines no criteria.")
    for criterion in criteria + [alternative for item in criteria for alternative in item.get("Alternatives", [])]:
        bounds = [criterion.get(name) for name in ("OnePointFrom", "TwoPointsAbove", "ThreePointsAbove")]
        if any(not isinstance(value, (int, float)) for value in bounds) or not bounds[0] <= bounds[1] <= bounds[2]:
            raise ValueError(f"Complexity criterion {criterion.get('Id')} needs ordered numeric thresholds.")
        if not isinstance(criterion.get("Weight", 0), (int, float)) or criterion.get("Weight", 0) < 0:
            raise ValueError(f"Complexity criterion {criterion.get('Id')} needs a non-negative weight.")
    levels = config.get("Levels") or []
    if not levels or min(level.get("MinScore", 0) for level in levels) != 0:
        raise ValueError("Complexity levels must start at MinScore 0.")


def points_for(value, criterion):
    if criterion.get("ZeroIsZero") and value == 0:
        return 0
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
    if criterion.get("BandLabels"):
        return criterion["BandLabels"]
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


def is_mapped_path(path, mappings):
    lower = path.lower()
    return any(lower == source.lower() or lower.startswith(source.rstrip("/").lower() + "/") for source, _ in mappings)


def projected_target_path(server_relative_url, mappings, target_site_path):
    """Destination path of a source file, or "" when a mapping file exists but excludes it."""
    text = normalize_sharegate_path_characters(decode_sharegate_path(strip_query_from_path(server_relative_url)))
    if not text.startswith("/"):
        text = "/" + text
    if mappings:
        if not is_mapped_path(text, mappings):
            return ""
        mapped = apply_path_mappings(text, mappings)
    elif target_site_path and not text.lower().startswith(target_site_path.lower() + "/"):
        mapped = target_site_path + text
    else:
        mapped = text
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
    library_counts = {}
    heavy_versions = maximum_versions = 0
    versions_complete = libraries_complete = True
    large = checked_out = excluded = 0
    target_paths = set()
    over_limit = unmapped = 0
    for row in read_rows(path):
        files += 1
        row_size = to_number(row.get("SizeBytes"))
        size += row_size
        versions += to_number(row.get("VersionsCount"))
        raw_versions = str(row.get("VersionsCount") or "").strip()
        try:
            version_count = float(raw_versions)
            if not math.isfinite(version_count) or version_count < 0:
                raise ValueError("Invalid versions")
        except ValueError:
            versions_complete = False
            version_count = 0
        heavy_versions += int(version_count >= settings.get("HeavyVersionsFrom", 100))
        maximum_versions = max(maximum_versions, version_count)
        web = str(row.get("WebUrl") or "").strip().rstrip("/").lower()
        if web:
            webs.add(web)
        library = str(row.get("LibraryUrl") or "").strip().rstrip("/").lower()
        if library:
            libraries.add((web, library))
            library_counts[(web, library)] = library_counts.get((web, library), 0) + 1
        else:
            libraries_complete = False
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
        if not target:
            unmapped += 1
            continue
        if target.lower() in target_paths:
            continue
        target_paths.add(target.lower())
        if len(target) > path_limit or max((len(part) for part in target.split("/")), default=0) > segment_limit:
            over_limit += 1
    metrics = {
        "FileCount": (files, f"{files:,} inventory rows"),
        "TotalSizeGiB": (round(size / 1024 ** 3, 2), f"{size:,.0f} bytes (current versions only)"),
        "SubsiteCount": (len(webs), "Distinct WebUrl values"),
        "LibraryCount": (len(libraries), "Distinct LibraryUrl values"),
        "AverageVersions": (versions / files if files else 0.0, "AVG(VersionsCount)"),
        "PathLimitFiles": (over_limit, f"Projected target path > {path_limit} characters or segment > {segment_limit}; "
                                       f"files outside the path mapping skipped: {unmapped:,}"),
        "LargeFiles": (large, f"SizeBytes > {format_number(float(settings.get('LargeFileGiB', 2)))} GiB"),
        "RiskyFiles": (checked_out + excluded, f"Checked out: {checked_out:,}; blocked or excluded: {excluded:,}"),
    }
    if versions_complete:
        metrics["VersionConcentration"] = (100 * heavy_versions / files if files else 0,
            f"Files with >= {settings.get('HeavyVersionsFrom', 100)} historical versions: {heavy_versions:,}/{files:,}; maximum: {maximum_versions:g}")
    else:
        metrics.pop("AverageVersions", None)
    if libraries_complete:
        metrics["LibraryConcentration"] = (max(library_counts.values(), default=0), "Maximum source files in one library")
    return metrics


def sha256_file(path):
    if not path:
        return ""
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def assessment_metrics(path, file_inventory, permission_inventory, excluded=()):
    """Consume a reviewed source assessment; never infer compatibility from component names."""
    if not path:
        return {}
    assessment = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if assessment.get("SchemaVersion") != 1:
        raise ValueError("Unsupported source complexity assessment schema.")
    inputs = assessment.get("Inputs", {})
    values = {}
    for section, criterion in (("FunctionalComponents", "FunctionalComplexity"), ("CustomPermissions", "CustomPermissionLevels")):
        if criterion in excluded:
            continue
        data = assessment.get(section, {})
        if data.get("Complete") is not True:
            continue
        expected_path = permission_inventory if section == "CustomPermissions" else file_inventory
        expected_key = "PermissionInventorySha256" if section == "CustomPermissions" else "FileInventorySha256"
        if not expected_path or inputs.get(expected_key, "").upper() != sha256_file(expected_path):
            continue
        if not data.get("Evidence") or not data.get("ReviewedAtUtc"):
            raise ValueError(f"{section} needs review date and source evidence.")
        reviewed = datetime.fromisoformat(data["ReviewedAtUtc"].replace("Z", "+00:00"))
        if reviewed.utcoffset() != timezone.utc.utcoffset(reviewed) or reviewed > datetime.now(timezone.utc):
            raise ValueError(f"{section} needs a valid UTC review date, not in the future.")
        items = data.get("Items")
        if not isinstance(items, list):
            raise ValueError(f"{section} needs an explicit Items list (empty only after a complete assessment).")
        if section == "FunctionalComponents":
            efforts = {}
            for item in items:
                effort = item.get("Effort")
                if not item.get("Id") or item.get("Category") not in ("Workflow", "Form", "Customization") or type(effort) is not int or not 0 <= effort <= 3 or not item.get("Evidence"):
                    raise ValueError("Each functional component needs Id, Category, Effort 0..3 and Evidence.")
                efforts[item["Id"]] = max(effort, efforts.get(item["Id"], 0))
            values[criterion] = (max(efforts.values(), default=0), f"Maximum remediation effort; {len(efforts)} distinct components. {data['Evidence']}")
        else:
            definitions = set()
            for item in items:
                if type(item.get("IsBuiltIn")) is not bool or type(item.get("IsUsed")) is not bool:
                    raise ValueError("Permission definitions need explicit IsBuiltIn and IsUsed booleans.")
                if item["IsBuiltIn"] or not item["IsUsed"]:
                    continue
                mask = str(item.get("EffectivePermissions", ""))
                if not mask.isdecimal() or not 0 <= int(mask) < 2 ** 64 or not item.get("Evidence"):
                    raise ValueError("Used custom permission definitions need an unsigned 64-bit effective permission mask and Evidence.")
                definitions.add(int(mask))
            values[criterion] = (len(definitions), f"Distinct effective permission masks, used custom definitions only. {data['Evidence']}")
    return values


def automatic_assessment_metrics(file_inventory, permission_inventory, target_platform="SPO"):
    """Read per-scan sidecars; partial automatic evidence never falls back to old manual values."""
    values, evidence, present, hashes = {}, {}, set(), {}
    for key, section, criterion, source in (
            ("File", "FunctionalComponents", "FunctionalComplexity", file_inventory),
            ("Permission", "CustomPermissions", "CustomPermissionLevels", permission_inventory)):
        sidecar = Path(str(source) + ".assessment.json.txt") if source else None
        hashes[key] = sha256_file(sidecar) if sidecar and sidecar.exists() else ""
        hashes[key + "Details"] = ""
        if not sidecar or not sidecar.exists():
            continue
        present.add(criterion)
        record = {"Source": "Automatic", "File": sidecar.name, "Valid": False,
                  "CollectionComplete": False, "QualificationComplete": False, "Estimated": key == "File"}
        evidence[criterion] = record
        try:
            assessment = json.loads(sidecar.read_text(encoding="utf-8-sig"))
            if not isinstance(assessment, dict):
                raise ValueError("Automatic assessment root must be an object.")
            detail_name = assessment.get("DetailsFile", "") if isinstance(assessment, dict) else ""
            safe_detail = bool(detail_name) and Path(detail_name).name == detail_name and "/" not in detail_name and "\\" not in detail_name
            detail = sidecar.parent / detail_name if safe_detail else None
            if detail and detail.is_file():
                hashes[key + "Details"] = sha256_file(detail)
            if assessment.get("SchemaVersion") != 1 or section not in assessment:
                raise ValueError("Unsupported automatic assessment schema or section.")
            data = assessment[section]
            if not isinstance(data, dict):
                raise ValueError("Automatic assessment section must be an object.")
            if data.get("Mode") != "Automatic":
                raise ValueError("Automatic assessment marker is missing.")
            if key == "File" and (target_platform != "SPO" or assessment.get("TargetPlatform", "SPO") != "SPO"):
                raise ValueError("Automatic functional rules apply to an SPO destination only.")
            hash_key = "FileInventorySha256" if key == "File" else "PermissionInventorySha256"
            inputs = assessment.get("Inputs")
            if not isinstance(inputs, dict) or not isinstance(inputs.get(hash_key), str):
                raise ValueError("Automatic assessment Inputs must contain a source SHA256 string.")
            if inputs[hash_key].upper() != sha256_file(source):
                raise ValueError("Automatic assessment does not match the source CSV SHA256.")
            if not safe_detail:
                raise ValueError("Automatic assessment detail reference must be a local file name.")
            if not isinstance(assessment.get("DetailsSha256"), str) or not hashes[key + "Details"] or hashes[key + "Details"] != assessment["DetailsSha256"].upper():
                raise ValueError("Automatic assessment detail CSV SHA256 mismatch.")
            items = data.get("Items")
            if not isinstance(items, list) or any(not isinstance(item, dict) for item in items):
                raise ValueError("Automatic assessment Items must be a list of objects.")
            for field in ("Errors", "Detectors"):
                if not isinstance(data.get(field, []), list) or any(not isinstance(item, dict) for item in data.get(field, [])):
                    raise ValueError(f"Automatic assessment {field} must be a list of objects.")
            if data.get("Complete") is True and any(not isinstance(data.get(field), str) for field in ("Evidence", "ReviewedAtUtc")):
                raise ValueError("Complete automatic assessment needs evidence and UTC date strings.")
            record.update({"Valid": True, "CollectionComplete": data.get("CollectionComplete") is True,
                           "QualificationComplete": data.get("QualificationComplete") is True,
                           "UnqualifiedCount": sum(item.get("Classification") == "Unqualified" for item in items),
                           "KnownRisks": [{"Id": item.get("Id"), "Technology": item.get("Technology"),
                                           "Effort": item.get("Effort"), "Evidence": item.get("Evidence"),
                                           "Reason": item.get("Reason")} for item in items
                                          if isinstance(item.get("Effort"), int) and item["Effort"] >= 2],
                           "Errors": data.get("Errors", []), "Detectors": data.get("Detectors", []),
                           "RulesSha256": assessment.get("RulesSha256", ""), "DetailFile": detail_name})
            if data.get("Complete") is True and record["CollectionComplete"] and record["QualificationComplete"] and not record["UnqualifiedCount"]:
                values.update(assessment_metrics(sidecar, file_inventory, permission_inventory))
        except (ValueError, TypeError, KeyError, OSError) as error:
            record.update(Valid=False, Error=str(error))
    return values, evidence, present, hashes


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
        # Template self-heal re-adds deleted entries, so "Enabled": false is the way to disable one.
        enabled = criterion.get("Enabled", True) is not False
        measured = values.get(criterion["Id"]) if enabled else None
        bands = band_texts(criterion)
        components = []
        if enabled and criterion.get("Alternatives"):
            for component in [criterion] + criterion["Alternatives"]:
                metric = values.get(component["Id"])
                components.append({"Id": component["Id"], "Name": component.get("MetricName", component["Name"]),
                    "Value": metric[0] if metric else None, "Bands": band_texts(component),
                    "Points": points_for(metric[0], component) if metric else None,
                    "Detail": metric[1] if metric else "Not evaluated"})
            if any(component["Points"] is None for component in components):
                measured = None
        row = {
            "Id": criterion["Id"], "Axis": criterion.get("Axis", ""), "Name": criterion.get("Name", criterion["Id"]),
            "Unit": criterion.get("Unit", ""), "Weight": criterion["Weight"], "Bands": bands,
            "Evaluated": measured is not None, "Value": None,
            "ValueDisplay": "Not evaluated" if enabled else "Disabled",
            "Points": None, "WeightedPoints": None, "MaxWeightedPoints": criterion["Weight"] * 3,
            "Band": "Not evaluated" if enabled else "Disabled in configuration", "Detail": "",
            "Components": components,
        }
        if components and measured is None:
            row["Detail"] = "Missing source metrics: " + ", ".join(component["Name"] for component in components if component["Points"] is None)
        if measured is not None:
            value, detail = measured
            points = points_for(value, criterion)
            if components:
                points = max(component["Points"] for component in components)
                detail = "MAX (no additive points): " + "; ".join(
                    f"{component['Name']} = {component['Value']:g}, {component['Points']}/3 ({component['Detail']})" for component in components)
            row.update({
                "Value": value, "ValueDisplay": f"{format_number(value)} {row['Unit']}".strip(),
                "Points": points, "WeightedPoints": points * criterion["Weight"],
                "Band": f"{points} pt (MAX of component bands)" if components else f"{points} pt ({bands[points]})", "Detail": detail,
            })
            if components:
                row["ValueDisplay"] = " / ".join(format_number(component["Value"]) for component in components)
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
        if rule.get("Enabled", True) is False or not row or row["Points"] is None or minimum not in level_rank or row["Points"] < rule.get("MinPoints", 3):
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
        "EvaluatedWeight": evaluated_weight,
        "CoveragePercent": round(100 * evaluated_weight / sum(item["Weight"] for item in config["Criteria"] if item.get("Enabled", True) is not False), 2) if any(item["Weight"] for item in config["Criteria"] if item.get("Enabled", True) is not False) else 0,
        "TotalWeight": sum(item["Weight"] for item in config["Criteria"] if item.get("Enabled", True) is not False),
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
    assessment_path = getattr(args, "source_assessment", None)
    automatic, assessment_evidence, present, automatic_hashes = automatic_assessment_metrics(
        args.file_inventory, args.permission_inventory, getattr(args, "assessment_target_platform", "SPO"))
    if len(present) < 2:
        values.update(assessment_metrics(assessment_path, args.file_inventory, args.permission_inventory, excluded=present))
    values.update(automatic)
    scored = score_migration(config, values)
    for row in scored["Criteria"]:
        record = assessment_evidence.get(row["Id"])
        if record:
            row["Assessment"] = record
            state = f"Automatic: collection {'complete' if record['CollectionComplete'] else 'partial'}, qualification {'complete' if record['QualificationComplete'] else 'partial'}"
            if record.get("Error"):
                state += "; " + record["Error"]
            if record.get("KnownRisks"):
                risk_counts = {}
                for item in record["KnownRisks"]:
                    label = f"{item['Technology']} (effort {item['Effort']})"
                    risk_counts[label] = risk_counts.get(label, 0) + 1
                state += "; known risks: " + "; ".join(f"{label} x{count}" for label, count in sorted(risk_counts.items()))
            row["Detail"] = state + ". " + row["Detail"]
            if row["Evaluated"] and record["Estimated"]:
                row["Band"] += " - estimated"
    missing = [row["Name"] for row in scored["Criteria"] if not row["Evaluated"] and next(item for item in config["Criteria"] if item["Id"] == row["Id"]).get("Enabled", True) is not False]
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
            "SourceAssessmentSha256": sha256_file(assessment_path),
            "AutomaticAssessmentHashes": automatic_hashes,
        },
        "Status": "Partial" if missing else "Complete",
        "MissingInputs": missing,
        "AssessmentEvidence": assessment_evidence,
        "FunctionalEstimate": "FunctionalComplexity" in automatic,
    }
    result.update(scored)
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description="Score migration complexity from source scans.")
    parser.add_argument("--migration-name", required=True)
    parser.add_argument("--file-inventory", required=True)
    parser.add_argument("--permission-inventory")
    parser.add_argument("--source-assessment", help="Reviewed source-complexity-assessment.json.txt; incomplete/stale sections remain unevaluated.")
    parser.add_argument("--assessment-target-platform", default="SPO", choices=("SPO", "Other"))
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
