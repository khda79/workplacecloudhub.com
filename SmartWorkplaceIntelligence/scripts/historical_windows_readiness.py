"""Historical Windows 11 readiness: explicit tenant attestation and conservative projection.

Version 1.0.0. This module does not perform collection, remote I/O, or publication.
"""
from __future__ import annotations
import argparse
import csv
import datetime as dt
import hashlib
import json
import os
import re
import tempfile
from collections import defaultdict
from pathlib import Path

RAW_COLUMNS = ("SourceFile", "pc.Date", "pc.name", "pc.Manufacturer", "pc.Model", "pc.WinVersion", "pc.WINEDT", "pc.Storage", "pc.Storage_Check", "pc.Memory", "pc.Memory_Check", "pc.TPM", "pc.TPM_Check", "pc.Processor", "pc.Processor_Check", "pc.SecureBoot", "pc.SecureBoot_Check", "pc.CPU", "pc.CPU_Check", "pc.GLOBALResult")
QUALIFIED_COLUMNS = ("TenantKey", "HistoricalSourceSHA256", "HistoricalObservationUtc") + RAW_COLUMNS
CHECKS = ("pc.Storage_Check", "pc.Memory_Check", "pc.TPM_Check", "pc.Processor_Check", "pc.SecureBoot_Check")
DETAIL = ("pc.GLOBALResult", "pc.Manufacturer", "pc.Model", "pc.WinVersion", "pc.WINEDT", "pc.Storage", "pc.Memory", "pc.TPM", "pc.Processor", "pc.SecureBoot") + CHECKS
OUTPUT_COLUMNS = ("TenantKey", "ObservationKey", "Device Name", "Observation UTC", "Historical Result", "Storage Check", "Memory Check", "TPM Check", "Processor Check", "Secure Boot Check", "Manufacturer", "Model", "Source Files", "Source SHA256", "Duplicate Count", "Observation Status", "Match Status", "Candidate Device Source ID")


def digest(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest().upper()


def name_key(value):
    return (value or "").strip().upper().removesuffix("$").split(".")[0]


def observation_utc(value):
    value = (value or "").strip()
    if value.endswith(":"):
        value = value[:-1].rstrip()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}Z", value):
        raise ValueError("Observation date is missing, malformed, or not UTC")
    return dt.datetime.strptime(value, "%Y-%m-%d %H:%M:%SZ").strftime("%Y-%m-%dT%H:%M:%SZ")


def _reader(path, columns):
    handle = open(path, encoding="utf-8-sig", newline="")
    reader = csv.DictReader(handle, strict=True)
    if reader.fieldnames != list(columns):
        handle.close()
        raise ValueError("Unexpected CSV schema or column order")
    return handle, reader


def _row(row):
    if None in row or any(v is None for v in row.values()):
        raise ValueError("Malformed CSV row")
    if not name_key(row["pc.name"]):
        raise ValueError("Missing device name")
    observation_utc(row["pc.Date"])
    if row["pc.GLOBALResult"].strip().upper() not in {"CAPABLE", "NOT CAPABLE", "UNDETERMINED"}:
        raise ValueError("Unrecognized readiness result")
    for col in CHECKS:
        if row[col].strip().upper() not in {"PASS", "FAIL", "UNDETERMINED", ""}:
            raise ValueError("Unrecognized check status")


def qualify(source, output, tenant, expected_sha256):
    """Create a new tenant-attested source, never overwrite inputs or outputs."""
    source, output = Path(source).resolve(strict=True), Path(output).resolve()
    if not tenant or not tenant.strip() or tenant != tenant.strip():
        raise ValueError("Explicit nonempty tenant key required")
    if not re.fullmatch(r"[0-9A-Fa-f]{64}", expected_sha256 or ""):
        raise ValueError("Explicit expected SHA-256 required")
    if source == output or output.exists():
        raise ValueError("Destination must be new and distinct from the source")
    expected = expected_sha256.upper()
    if digest(source) != expected:
        raise ValueError("Source SHA-256 differs from attested source")
    handle, reader = _reader(source, RAW_COLUMNS)
    tmp = None
    count = 0
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", newline="", dir=output.parent, delete=False, prefix=".readiness-", suffix=".tmp") as out:
            tmp = Path(out.name)
            writer = csv.DictWriter(out, fieldnames=QUALIFIED_COLUMNS, lineterminator="\n")
            writer.writeheader()
            for row in reader:
                _row(row)
                writer.writerow({"TenantKey": tenant, "HistoricalSourceSHA256": expected, "HistoricalObservationUtc": observation_utc(row["pc.Date"]), **row})
                count += 1
            out.flush()
            os.fsync(out.fileno())
        if not count:
            raise ValueError("Source contains no observations")
        if digest(source) != expected:
            raise ValueError("Source changed during qualification")
        # Atomic create: a concurrent file is never replaced.
        if os.name == "nt":
            os.rename(tmp, output)  # Windows rename fails when destination exists.
        else:
            os.link(tmp, output)
        return {"rows": count, "source_sha256": expected, "qualified_sha256": digest(output)}
    finally:
        handle.close()
        if tmp is not None:
            tmp.unlink(missing_ok=True)


def project(rows, tenant, migration, inventory, ad):
    """Latest observation only for a currently eligible fallback device.

    A unique name is only a candidate; identity conflicts or model differences
    quarantine the observation. A usable Intune assessment always takes priority.
    """
    migration_names = defaultdict(list)
    device_ids = set()
    for row in migration:
        device_id = (row.get("Device Source ID") or "").strip()
        if not device_id or device_id in device_ids:
            raise ValueError("Missing or duplicate migration device identity")
        device_ids.add(device_id)
        key = name_key(row.get("Device Name", ""))
        if key:
            migration_names[key].append(row)
    intune_names, ad_names = defaultdict(list), defaultdict(list)
    for source, index, name_col in ((inventory, intune_names, "Device name"), (ad, ad_names, "Name")):
        for row in source:
            if row.get("TenantKey") != tenant:
                raise ValueError("Foreign or missing tenant on identity source")
            key = name_key(row.get(name_col, ""))
            if key:
                index[key].append(row)
    needed = {name for name, candidates in migration_names.items()
              if any((r.get("Upgrade Eligibility") or "").strip().lower() in {"unknown", "not assessed", "undetermined"} for r in candidates)}
    latest = {}
    hashes = set()
    for row in rows:
        if row.get("TenantKey") != tenant or not re.fullmatch(r"[0-9A-F]{64}", row.get("HistoricalSourceSHA256", "")):
            raise ValueError("Unqualified or foreign historical source")
        _row(row)
        stamp = observation_utc(row["pc.Date"])
        if row.get("HistoricalObservationUtc") != stamp:
            raise ValueError("Qualified observation date differs from raw source")
        hashes.add(row["HistoricalSourceSHA256"])
        name = name_key(row["pc.name"])
        if name not in needed:
            continue
        item = latest.get(name)
        if item is None or stamp > item["stamp"]:
            item = {"stamp": stamp, "variants": set(), "sources": set(), "count": 0}
            latest[name] = item
        if stamp == item["stamp"]:
            item["variants"].add(tuple(row[x].strip() for x in DETAIL))
            item["sources"].add(row["SourceFile"].strip())
            item["count"] += 1
    if len(hashes) > 1:
        raise ValueError("Historical rows have inconsistent source hashes")
    source_hash = next(iter(hashes), "")
    result = []
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    for name in sorted(needed):
        item = latest.get(name)
        if item is None:
            continue
        candidates = migration_names[name]
        for candidate in candidates:
            if (candidate.get("Upgrade Eligibility") or "").strip().lower() not in {"unknown", "not assessed", "undetermined"}:
                continue
            device_id = candidate["Device Source ID"].strip()
            matches = intune_names.get(name, [])
            ad_matches = ad_names.get(name, [])
            conflict = len(item["variants"]) != 1
            values = dict(zip(DETAIL, next(iter(item["variants"])) if not conflict else ("",) * len(DETAIL)))
            reason = "Candidate by unique name; identity unverified"
            if len(candidates) != 1 or len(matches) > 1 or len(ad_matches) > 1:
                reason = "Ambiguous current name"
            elif item["stamp"] > now:
                reason = "Future historical observation"
            elif conflict:
                reason = "Conflicting latest observations"
            elif candidate.get("Inventory Source") == "Intune":
                if len(matches) != 1 or (matches[0].get("Device ID") or "").strip().lower() != device_id.lower():
                    reason = "Intune identity mismatch"
                elif (matches[0].get("Identity Status") or "").strip().lower() == "identity conflict":
                    reason = "Intune identity conflict"
                else:
                    old_model = values["pc.Model"].strip().upper()
                    new_model = (matches[0].get("Model") or "").strip().upper()
                    if old_model and new_model and old_model not in {"UNKNOWN", "N/A"} and new_model not in {"UNKNOWN", "N/A"} and old_model != new_model:
                        reason = "Model mismatch"
            elif candidate.get("Inventory Source") == "AD only":
                if len(ad_matches) != 1 or len(matches) != 0:
                    reason = "AD identity ambiguous"
                else:
                    expected_ad_id = "ad:" + tenant + ":" + (ad_matches[0].get("ObjectGUID") or "").strip().lower()
                    if not (ad_matches[0].get("ObjectGUID") or "").strip() or device_id.lower() != expected_ad_id.lower():
                        reason = "AD identity mismatch"
            else:
                reason = "Unknown inventory source"
            usable = reason == "Candidate by unique name; identity unverified"
            if not usable:
                values = dict.fromkeys(DETAIL, "")
            status = "Conflicting observations" if conflict else ("Observed" if usable else "Not qualified")
            observation_key = hashlib.sha256((tenant + "\x1f" + device_id).encode("utf-8")).hexdigest()
            result.append(dict(zip(OUTPUT_COLUMNS, (tenant, observation_key, name, item["stamp"], values["pc.GLOBALResult"], values["pc.Storage_Check"], values["pc.Memory_Check"], values["pc.TPM_Check"], values["pc.Processor_Check"], values["pc.SecureBoot_Check"], values["pc.Manufacturer"], values["pc.Model"], " | ".join(sorted(item["sources"])), source_hash, item["count"], status, reason, device_id))))
    return result


def _identity_csv(path, required, tenant):
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, strict=True)
        if not reader.fieldnames or len(reader.fieldnames) != len(set(reader.fieldnames)) or not set(required).issubset(reader.fieldnames):
            raise ValueError("Invalid current identity schema")
        result = list(reader)
    if any(r.get("TenantKey") != tenant for r in result):
        raise ValueError("Current identity tenant mismatch")
    return result

def validate_qualified(source, tenant):
    """Validate optional tenant-bound historical data without producing output."""
    before = digest(source)
    handle, reader = _reader(source, QUALIFIED_COLUMNS)
    count = 0
    def counted_rows():
        nonlocal count
        for row in reader:
            count += 1
            yield row
    try:
        project(counted_rows(), tenant, [], [], [])
    finally:
        handle.close()
    if not count:
        raise ValueError("Historical source is empty")
    if digest(source) != before:
        raise ValueError("Historical source changed during validation")
    return {"rows": count, "qualified_sha256": before}


def build_projection(source, migration_path, output, tenant, inventory_path, ad_path):
    """Prepare an optional source; absent means an empty observation table."""
    output = Path(output).resolve()
    input_paths = [Path(x).resolve() for x in (source, migration_path, inventory_path, ad_path)]
    if output in input_paths:
        raise ValueError("Projection output must not overwrite an input")
    with open(migration_path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, strict=True)
        if not reader.fieldnames or not {"Device Source ID", "Device Name", "Inventory Source", "Upgrade Eligibility"}.issubset(reader.fieldnames):
            raise ValueError("Invalid migration schema")
        migration = list(reader)
    inventory = _identity_csv(inventory_path, ("TenantKey", "Device ID", "Device name"), tenant)
    ad = _identity_csv(ad_path, ("TenantKey", "ObjectGUID", "Name"), tenant)
    observations = []
    source_path = Path(source)
    if source_path.exists():
        before = digest(source_path)
        handle, reader = _reader(source_path, QUALIFIED_COLUMNS)
        try:
            observations = project(reader, tenant, migration, inventory, ad)
        finally:
            handle.close()
        if digest(source_path) != before:
            raise ValueError("Qualified source changed during projection")
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", newline="", dir=output.parent, delete=False, prefix=".readiness-", suffix=".tmp") as out:
            tmp = Path(out.name)
            writer = csv.DictWriter(out, fieldnames=OUTPUT_COLUMNS, lineterminator="\n")
            writer.writeheader()
            writer.writerows(observations)
            out.flush()
            os.fsync(out.fileno())
        os.replace(tmp, output)
        return {"observations": len(observations), "source_available": source_path.exists()}
    finally:
        if tmp is not None:
            tmp.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True)
    parser.add_argument("--output")
    parser.add_argument("--tenant-key", required=True)
    parser.add_argument("--expected-sha256")
    parser.add_argument("--migration")
    parser.add_argument("--inventory")
    parser.add_argument("--ad")
    parser.add_argument("--validate", action="store_true")
    args = parser.parse_args()
    if args.validate:
        result = validate_qualified(args.source, args.tenant_key)
    elif args.migration:
        if not args.inventory or not args.ad or not args.output:
            parser.error("--inventory, --ad and --output are required with --migration")
        result = build_projection(args.source, args.migration, args.output, args.tenant_key, args.inventory, args.ad)
    else:
        if not args.output:
            parser.error("--output is required for qualification")
        result = qualify(args.source, args.output, args.tenant_key, args.expected_sha256)
    print(json.dumps(result))


if __name__ == "__main__":
    main()
