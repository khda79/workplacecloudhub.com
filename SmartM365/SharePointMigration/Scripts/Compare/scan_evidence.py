"""Create a portable, copy-resistant receipt for a completed inventory CSV."""
__version__ = "1.0.0"

import argparse
import csv
import hashlib
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4


def manifest_path(csv_path):
    return Path(str(csv_path) + ".manifest.json.txt")


def inspect_csv(path):
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        columns = next(reader, [])
        rows = sum(1 for _ in reader)
    if not columns:
        raise ValueError(f"Inventory has no CSV header: {path}")
    return digest.hexdigest(), rows


def write_manifest(path, side, kind, scope):
    path = Path(path)
    digest, rows = inspect_csv(path)
    scope_path = Path(scope)
    if scope_path.is_file():
        scope_roots = [line.strip() for line in scope_path.read_text(encoding="utf-8-sig").splitlines()
                       if line.strip() and not line.lstrip().startswith("#")]
    else:
        scope_roots = [scope] if scope else []
    receipt = {
        "SchemaVersion": 1,
        "InventoryFile": path.name,
        "CompletedAtUtc": datetime.now(timezone.utc).isoformat(),
        "Side": side,
        "Kind": kind,
        "Scope": scope,
        "ScopeRoots": scope_roots,
        "Rows": rows,
        "Sha256": digest,
    }
    target = manifest_path(path)
    temporary = target.with_name(target.name + "." + uuid4().hex + ".tmp")
    try:
        temporary.write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)
    return target


def read_scan_time(path):
    path = Path(path)
    receipt = manifest_path(path)
    if receipt.is_file():
        data = json.loads(receipt.read_text(encoding="utf-8"))
        digest, _ = inspect_csv(path)
        if data.get("InventoryFile") != path.name or data.get("Sha256", "").lower() != digest:
            raise ValueError(f"Scan manifest does not match inventory: {path}")
        return datetime.fromisoformat(data["CompletedAtUtc"]).astimezone(timezone.utc), "Manifest"
    found = re.search(r"-(\d{8})-(\d{6})(?:-[^.]+)?\.csv$", path.name, re.I)
    if found:
        local_time = datetime.strptime("".join(found.groups()), "%Y%m%d%H%M%S").astimezone()
        return local_time.astimezone(timezone.utc), "Filename"
    return None, "Unknown"


def describe_pair(source, target, max_gap_hours=12, max_age_hours=24):
    source_at, source_kind = read_scan_time(source)
    target_at, target_kind = read_scan_time(target)
    if not source_at or not target_at:
        return {"ScanEvidenceStatus": "Unverified", "SourceScanCompletedUtc": "", "TargetScanCompletedUtc": ""}
    now = datetime.now(timezone.utc)
    gap = abs((source_at - target_at).total_seconds()) / 3600
    oldest_age = (now - min(source_at, target_at)).total_seconds() / 3600
    stale = gap > max_gap_hours or oldest_age > max_age_hours or oldest_age < -0.25
    return {
        "ScanEvidenceStatus": "Stale" if stale else ("Verified" if source_kind == target_kind == "Manifest" else "LegacyFilename"),
        "SourceScanCompletedUtc": source_at.isoformat(),
        "TargetScanCompletedUtc": target_at.isoformat(),
        "ScanGapHours": f"{gap:.2f}",
        "OldestScanAgeHours": f"{oldest_age:.2f}",
        "SourceScanEvidence": source_kind,
        "TargetScanEvidence": target_kind,
    }


if __name__ == "__main__":
    from pathlib import Path as _ScriptPath
    import sys as _script_sys
    print(f"{_ScriptPath(__file__).name} v{__version__}", file=_script_sys.stderr)
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", required=True)
    parser.add_argument("--side", required=True, choices=("Source", "Target"))
    parser.add_argument("--kind", required=True, choices=("File", "Permission"))
    parser.add_argument("--scope", required=True)
    args = parser.parse_args()
    print(write_manifest(args.csv, args.side, args.kind, args.scope))
