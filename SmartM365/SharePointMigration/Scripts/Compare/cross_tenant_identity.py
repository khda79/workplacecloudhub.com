"""Strict, explicit identity mapping for cross-tenant permission evidence.

This module only validates and resolves mappings. It does not decide whether
permissions or group membership are equivalent, or change any inventory row.
"""
__version__ = "1.0.1"

import csv
import argparse
import hashlib
import io
import json
import os
from pathlib import Path
import re
import tempfile
from urllib.parse import urlsplit


REQUIRED_COLUMNS = (
    "SourceWebUrl", "SourcePrincipalType", "SourcePrincipalLoginName",
    "TargetWebUrl", "TargetPrincipalType", "TargetPrincipalLoginName", "Status",
)


def _identity(value):
    return str(value or "").strip().casefold()


def _web_url(value):
    value = str(value or "").strip()
    try:
        parsed = urlsplit(value)
        parsed.port  # Force invalid-port validation.
    except ValueError as exc:
        raise ValueError("Invalid source web URL") from exc
    if (parsed.scheme.lower() != "https" or not parsed.hostname or
            parsed.username or parsed.password or parsed.query or parsed.fragment or
            any(character.isspace() or ord(character) < 32 for character in value)):
        raise ValueError("Web URL must be an absolute HTTPS URL without credentials, query or fragment")
    return value.rstrip("/").casefold()


def source_key(web_url, principal_type, login_name):
    kind, login = _identity(principal_type), _identity(login_name)
    if not kind or not login:
        raise ValueError("Source principal type and login name must be non-empty")
    return _web_url(web_url), kind, login


def load_confirmed_mappings(path):
    """Load confirmed, unambiguous mappings; never infer identities by name."""
    mappings = {}
    targets = {}
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or not set(REQUIRED_COLUMNS).issubset(reader.fieldnames):
            raise ValueError("Identity mapping CSV must contain: " + ", ".join(REQUIRED_COLUMNS))
        if len(reader.fieldnames) != len(set(reader.fieldnames)):
            raise ValueError("Identity mapping CSV contains duplicate column names")
        for number, row in enumerate(reader, start=2):
            if None in row:
                raise ValueError(f"Identity mapping row {number} has extra fields")
            if _identity(row.get("Status")) != "confirmed":
                raise ValueError(f"Identity mapping row {number} is not confirmed")
            key = source_key(row.get("SourceWebUrl"), row.get("SourcePrincipalType"), row.get("SourcePrincipalLoginName"))
            target_type = str(row.get("TargetPrincipalType") or "").strip()
            target_login = str(row.get("TargetPrincipalLoginName") or "").strip()
            target_web = _web_url(row.get("TargetWebUrl"))
            if not target_type or not target_login:
                raise ValueError(f"Identity mapping row {number} has an empty target principal")
            if key in mappings:
                raise ValueError(f"Identity mapping row {number} duplicates a source principal in the same web")
            target_key = (target_web, _identity(target_type), _identity(target_login))
            source_identity = (key[1], key[2])
            if target_key in targets and targets[target_key] != source_identity:
                raise ValueError(f"Identity mapping row {number} maps distinct source principals to the same target principal")
            targets[target_key] = source_identity
            mappings[key] = {"WebUrl": target_web, "PrincipalType": target_type, "PrincipalLoginName": target_login}
    if not mappings:
        raise ValueError("Identity mapping CSV is empty")
    return mappings


def resolve_source_row(row, mappings):
    """Return a confirmed mapping or None; preserve the original evidence."""
    key = source_key(row.get("WebUrl"), row.get("PrincipalType"), row.get("PrincipalLoginName"))
    return mappings.get(key)


def resolve_inventory(rows, mappings):
    """Partition rows without discarding unresolved identities or modifying evidence."""
    resolved, unresolved = [], []
    for row in rows:
        evidence = dict(row)
        try:
            target = resolve_source_row(evidence, mappings)
        except ValueError:
            unresolved.append((evidence, "Invalid source principal or web URL"))
            continue
        if target is None:
            unresolved.append((evidence, "No confirmed mapping"))
            continue
        resolved.append((evidence, dict(target)))
    return resolved, unresolved


def matches_target_row(mapping, target_row):
    """Match explicit target web, principal type and login, without aliases."""
    expected = source_key(mapping.get("WebUrl"), mapping.get("PrincipalType"), mapping.get("PrincipalLoginName"))
    observed = source_key(target_row.get("WebUrl"), target_row.get("PrincipalType"), target_row.get("PrincipalLoginName"))
    return expected == observed


def load_web_mappings(path):
    """Load full HTTPS web URL pairs for cross-tenant host-aware routing."""
    routes = {}
    for number, raw in enumerate(Path(path).read_text(encoding="utf-8-sig").splitlines(), start=1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = re.split(r"[\t;, ]+", line)
        if len(parts) != 2:
            raise ValueError(f"Site mapping row {number} must have exactly two URLs")
        source, target = _web_url(parts[0]), _web_url(parts[1])
        if source in routes and routes[source] != target:
            raise ValueError(f"Site mapping row {number} conflicts with an earlier source web")
        routes[source] = target
    if not routes:
        raise ValueError("Cross-tenant site mapping is empty")
    return routes


def expected_target_web(source_web, routes):
    source = _web_url(source_web)
    candidates = [root for root in routes if source == root or source.startswith(root + "/")]
    if not candidates:
        raise ValueError("Source web is outside the configured cross-tenant site mapping")
    root = max(candidates, key=len)
    return routes[root] + source[len(root):]


def prepare_comparison_rows(source_rows, target_rows, mappings, mapped_source_web_path, normalize_web_path, expected_target_web_for_source=None):
    """Prepare isolated exact identity tokens; reject incomplete or inconsistent evidence.

    The caller supplies its existing path-mapping functions. The returned rows
    retain original identity fields and never infer aliases or group membership.
    """
    if not source_rows or not target_rows:
        raise ValueError("Cross-tenant identity comparison requires non-empty source and target inventories")
    observed_target_webs = set()
    prepared_targets = []
    for row in target_rows:
        target_key = source_key(row.get("WebUrl"), row.get("PrincipalType"), row.get("PrincipalLoginName"))
        observed_target_webs.add(target_key[0])
        prepared = dict(row)
        prepared["CrossTenantComparisonPrincipal"] = json.dumps(target_key, ensure_ascii=True, separators=(",", ":"))
        prepared_targets.append(prepared)
    prepared_sources = []
    unresolved_count = 0
    for row in source_rows:
        try:
            mapped = resolve_source_row(row, mappings)
        except ValueError:
            mapped = None
        if mapped is None:
            unresolved_count += 1
            continue
        target_key = source_key(mapped["WebUrl"], mapped["PrincipalType"], mapped["PrincipalLoginName"])
        if target_key[0] not in observed_target_webs:
            raise ValueError("A mapped target web is not evidenced by the target permission inventory")
        if expected_target_web_for_source is not None and target_key[0] != expected_target_web_for_source(row.get("WebUrl")):
            raise ValueError("Identity mapping target URL conflicts with the configured host-aware site mapping")
        if mapped_source_web_path(row) != normalize_web_path(mapped["WebUrl"]):
            raise ValueError("Identity mapping target web conflicts with the configured site path mapping")
        prepared = dict(row)
        prepared["CrossTenantComparisonPrincipal"] = json.dumps(target_key, ensure_ascii=True, separators=(",", ":"))
        prepared_sources.append(prepared)
    if unresolved_count:
        raise ValueError(f"Cross-tenant identity mappings are incomplete for {unresolved_count} source assignment rows")
    return prepared_sources, prepared_targets


def _sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def preflight_inventory(source_path, mapping_path, unresolved_path=None, target_path=None):
    """Validate mapping coverage; this is not a permission comparison."""
    from scan_evidence import assert_inventory_complete, read_scan_time

    source_path, mapping_path = Path(source_path), Path(mapping_path)
    assert_inventory_complete(source_path)
    _, evidence_kind = read_scan_time(source_path, "Source", "Permission")
    source_hash, mapping_hash = _sha256(source_path), _sha256(mapping_path)
    mappings = load_confirmed_mappings(mapping_path)
    with source_path.open("r", encoding="utf-8-sig", newline="") as handle:
        sample = handle.read(8192)
        handle.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(handle, dialect=dialect)
        if not reader.fieldnames or not {"WebUrl", "PrincipalType", "PrincipalLoginName"}.issubset(reader.fieldnames):
            raise ValueError("Source permission inventory lacks required identity columns")
        if len(reader.fieldnames) != len(set(reader.fieldnames)):
            raise ValueError("Source permission inventory contains duplicate columns")
        rows = []
        for row in reader:
            if None in row:
                raise ValueError("Source permission inventory contains a malformed row")
            rows.append(row)
    if not rows:
        raise ValueError("Source permission inventory is empty")
    resolved, unresolved = resolve_inventory(rows, mappings)
    target_summary = {"TargetInventoryChecked": False}
    if target_path is not None:
        target_path = Path(target_path)
        assert_inventory_complete(target_path)
        _, target_evidence = read_scan_time(target_path, "Target", "Permission")
        target_hash = _sha256(target_path)
        with target_path.open("r", encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(8192)
            handle.seek(0)
            try:
                target_dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                target_dialect = csv.excel
            reader = csv.DictReader(handle, dialect=target_dialect)
            if not reader.fieldnames or not {"WebUrl", "PrincipalType", "PrincipalLoginName"}.issubset(reader.fieldnames):
                raise ValueError("Target permission inventory lacks required identity columns")
            if len(reader.fieldnames) != len(set(reader.fieldnames)):
                raise ValueError("Target permission inventory contains duplicate columns")
            target_keys = set()
            target_rows = invalid_targets = 0
            for target_row in reader:
                if None in target_row:
                    raise ValueError("Target permission inventory contains a malformed row")
                target_rows += 1
                try:
                    target_keys.add(source_key(target_row.get("WebUrl"), target_row.get("PrincipalType"), target_row.get("PrincipalLoginName")))
                except ValueError:
                    invalid_targets += 1
        if _sha256(target_path) != target_hash:
            raise ValueError("Target input changed during mapping validation")
        assert_inventory_complete(target_path)
        _, final_target_evidence = read_scan_time(target_path, "Target", "Permission")
        if final_target_evidence != target_evidence:
            raise ValueError("Target scan evidence changed during mapping validation")
        observed = sum(source_key(target["WebUrl"], target["PrincipalType"], target["PrincipalLoginName"]) in target_keys for _, target in resolved)
        target_summary = {
            "TargetInventoryChecked": True,
            "TargetSha256": target_hash,
            "TargetScanEvidence": target_evidence,
            "TargetRows": target_rows,
            "TargetRowsWithInvalidIdentity": invalid_targets,
            "MappedSourceRowsWithTargetPrincipalObserved": observed,
            "MappedSourceRowsWithTargetPrincipalNotObserved": len(resolved) - observed,
        }
    if _sha256(source_path) != source_hash or _sha256(mapping_path) != mapping_hash:
        raise ValueError("Input changed during mapping validation")
    assert_inventory_complete(source_path)
    _, final_evidence_kind = read_scan_time(source_path, "Source", "Permission")
    if final_evidence_kind != evidence_kind:
        raise ValueError("Source scan evidence changed during mapping validation")
    if unresolved_path is not None:
        output = io.StringIO(newline="")
        columns = ("WebUrl", "PrincipalType", "PrincipalLoginName", "Reason")
        writer = csv.DictWriter(output, fieldnames=columns)
        writer.writeheader()
        for evidence, reason in unresolved:
            writer.writerow({name: reason if name == "Reason" else evidence.get(name, "") for name in columns})
        destination = Path(unresolved_path)
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8-sig", newline="", dir=destination.parent, prefix=".identity-review-", suffix=".tmp", delete=False) as handle:
                temporary = Path(handle.name)
                handle.write(output.getvalue())
                handle.flush()
                os.fsync(handle.fileno())
            # Windows rename refuses an existing destination; POSIX rename does not.
            if os.name == "nt":
                os.rename(temporary, destination)
            else:
                os.link(temporary, destination)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
    return {
        "PreflightOnly": True,
        "SourceSha256": source_hash,
        "MappingSha256": mapping_hash,
        "SourceScanEvidence": evidence_kind,
        "SourceRows": len(rows),
        "ResolvedRows": len(resolved),
        "UnresolvedRows": len(unresolved),
        "GroupMembershipCompared": False,
        "PermissionsCompared": False,
        **target_summary,
    }


def main():
    parser = argparse.ArgumentParser(description="Check explicit cross-tenant identity mapping coverage without comparing permissions.")
    parser.add_argument("--source-csv", required=True)
    parser.add_argument("--mapping-csv", required=True)
    parser.add_argument("--target-csv", help="Optionally check whether mapped principals occur in a target permission inventory.")
    parser.add_argument("--unresolved-csv", help="Create a new private CSV containing unresolved principals; existing files are never overwritten.")
    arguments = parser.parse_args()
    result = preflight_inventory(arguments.source_csv, arguments.mapping_csv, arguments.unresolved_csv, arguments.target_csv)
    print(json.dumps(result, sort_keys=True))
    return 2 if result["UnresolvedRows"] else 0


if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from console_lifecycle import run_console_script
    raise SystemExit(run_console_script(main, __file__, __version__))
