"""Read-only, current-only CMDB snapshot consumer (Python 3.10+, stdlib).

Candidate infrastructure, NOT a Power Query connector. Consumers must read
Snapshot.iter_rows(), not reopen the synchronized CSVs after validation.
No source collection, output copies, history, report edits or external actions.
"""
import argparse
import collections
import csv
import datetime as dt
import hashlib
import io
import importlib.util
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from types import MappingProxyType

VERSION = '0.1.3'
MANIFEST = 'current.json.txt'
OWNER = 'SmartInventory-CMDB-Prepared'
UTC = dt.timezone.utc
IDENTITY_FIELDS = ('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')
CSV_FIELD_LIMIT = 64 * 1024 * 1024
DEFAULT_MAX_BYTES = 2 * 1024 * 1024 * 1024
PREPARATION = Path(__file__).resolve().parents[2] / 'SmartM365/SmartInventory/PreparedEvidence'
CONTRACT = PREPARATION / 'cmdb-prepared-contract.json.txt'
REGISTRY = PREPARATION.parents[1] / 'Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt'
# Use the same pure policy implementation as preparation, without importing its CLI.
_freshness_spec = importlib.util.spec_from_file_location('cmdb_reader_freshness', PREPARATION / 'cmdb_freshness.py')
freshness = importlib.util.module_from_spec(_freshness_spec)
_freshness_spec.loader.exec_module(freshness)
PARENTS = {
    'TenantDeviceKey': 'DimDevice', 'TenantUserKey': 'DimUser',
    'TenantGroupKey': 'DimGroup', 'TenantSkuKey': 'DimLicenseSku',
    'TenantServicePlanKey': 'DimLicenseServicePlan',
    'TenantApplicationKey': 'DimDetectedApplication', 'TenantTeamKey': 'DimTeam',
    'TenantADGroupKey': 'ADGroupSource', 'TenantADObjectKey': 'ADDirectoryObject',
}


def digest(data):
    return hashlib.sha256(data).hexdigest().upper()


def _json(data):
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError('Duplicate JSON property: ' + key)
            result[key] = value
        return result
    def invalid_constant(_):
        raise ValueError('Non-finite JSON value')
    result = json.loads(data.decode('utf-8-sig'), object_pairs_hook=unique,
                        parse_constant=invalid_constant)
    if not isinstance(result, dict):
        raise ValueError('Expected a JSON document object')
    return result


def _utc(value):
    try:
        result = dt.datetime.fromisoformat(value.replace('Z', '+00:00'))
        if result.tzinfo is None:
            raise ValueError()
        return result.astimezone(UTC)
    except (ValueError, TypeError, AttributeError):
        raise ValueError('Timestamp requires an explicit timezone') from None


def _hash(value):
    if not isinstance(value, str) or not re.fullmatch(r'[0-9a-fA-F]{64}', value):
        raise ValueError('Invalid SHA256 evidence')
    return value.upper()


def _count(value):
    if type(value) is not int or value < 0:
        raise ValueError('Row count must be a non-negative integer')
    return value


def _named(records, field):
    if not isinstance(records, list):
        raise ValueError('Evidence list is required')
    result = {}
    for record in records:
        name = record[field]
        if not isinstance(name, str) or not name or name in result:
            raise ValueError('Missing or repeated evidence name: ' + field)
        result[name] = record
    return result


def _safe_name(name):
    if not isinstance(name, str) or not re.fullmatch(r'[A-Za-z0-9_.-]+', name) or name in ('.', '..'):
        raise ValueError('Unsafe artifact name')
    return name


def _read(path, limit):
    if path.is_symlink() or not path.is_file():
        raise ValueError('Missing, linked or non-file artifact: ' + path.name)
    if path.stat().st_size > limit:
        raise ValueError('Snapshot memory budget exceeded; no partial load accepted')
    chunks, size = [], 0
    with path.open('rb') as stream:
        while True:
            # Never preallocate the entire (potentially GiB-sized) budget
            # for a small CSV. Also bound files that grow after stat().
            chunk = stream.read(min(4 * 1024 * 1024, limit - size + 1))
            if not chunk:
                break
            size += len(chunk)
            if size > limit:
                raise ValueError('Snapshot memory budget exceeded; no partial load accepted')
            chunks.append(chunk)
    return b''.join(chunks)


def _csv_rows(data, name, columns=None):
    previous = csv.field_size_limit()
    csv.field_size_limit(max(previous, CSV_FIELD_LIMIT))
    try:
        with io.TextIOWrapper(io.BytesIO(data), encoding='utf-8-sig', newline='') as stream:
            reader = csv.DictReader(stream, strict=True)
            fields = reader.fieldnames
            if not fields or len(fields) != len(set(fields)) or any(not field for field in fields):
                raise ValueError('Missing or duplicate CSV header: ' + name)
            if columns is not None and fields != columns:
                raise ValueError('CSV schema mismatch: ' + name)
            for row in reader:
                if None in row or any(value is None for value in row.values()):
                    raise ValueError('Malformed logical CSV row: ' + name)
                yield row
    except (csv.Error, UnicodeError) as error:
        raise ValueError('Invalid CSV encoding or syntax: ' + name) from error
    finally:
        csv.field_size_limit(previous)


def _evidence(manifest, contract, registry, registry_hash, now):
    evidence = manifest['SourceEvidence']
    if contract['producerRegistry'] != REGISTRY.name or _hash(evidence['RegistrySHA256']) != registry_hash:
        raise ValueError('Producer registry mismatch')
    freshness.policies(contract)
    future = now + dt.timedelta(minutes=5)
    producers = _named(registry['Producers'], 'Script')
    receipts = _named(evidence['ProducerReceipts'], 'Producer')
    files = _named(evidence['Files'], 'File')
    expected_files = {source['file'] for source in contract['sources']}
    source_definitions = {source['file']: source for source in contract['sources']}
    if len(source_definitions) != len(contract['sources']):
        raise ValueError('Repeated source in preparation contract')
    if set(receipts) != set(producers) or set(files) != expected_files:
        raise ValueError('Missing or unexpected source/producer evidence')
    starts, ends, registered = [], [], set()
    for name, producer in producers.items():
        receipt = receipts[name]
        coverage = receipt.get('DomainCoverage')
        if coverage is not None:
            # Same pure policy as preparation; no blanket partial-source bypass.
            freshness.partial_ad_coverage(name, coverage)
        if receipt['File'] != producer['Receipt'] or receipt['Scope'] != producer['Scope']:
            raise ValueError('Producer receipt or full scope mismatch: ' + name)
        _hash(receipt['SHA256'])
        if any(not isinstance(receipt.get(field), str) or not receipt[field].strip()
               for field in ('RunId', 'ScriptVersion')):
            raise ValueError('Missing producer lineage: ' + name)
        start, end = _utc(receipt['StartedAtUtc']), _utc(receipt['CompletedAtUtc'])
        if start > end or end > future:
            raise ValueError('Invalid producer acquisition interval: ' + name)
        rule = freshness.producer_policy(contract, producer['Files'])
        if now - start > dt.timedelta(hours=rule['MaxAgeHours']):
            raise ValueError('Stale acquisition evidence: ' + name)
        local = set(producer['Files'])
        if len(local) != len(producer['Files']) or registered & local:
            raise ValueError('Repeated source in producer registry')
        registered.update(local)
        for file in local:
            record = files[file]
            if (record.get('Status') != 'Success' or record.get('IsPartialInventory') is not bool(coverage)
                    or type(record.get('Errors')) is not int or record['Errors'] != 0):
                raise ValueError('Incomplete source result: ' + file)
            if record.get('DomainCoverage') != coverage:
                raise ValueError('Source domain coverage mismatch: ' + file)
            if any(record.get(field) != receipt[field]
                   for field in ('Producer', 'RunId', 'ScriptVersion', 'Scope', 'StartedAtUtc')):
                raise ValueError('Source lineage mismatch: ' + file)
            completed = _utc(record['CompletedAtUtc'])
            if not start <= completed <= end:
                raise ValueError('Invalid source acquisition interval: ' + file)
            _hash(record['SHA256'])
            if _count(record['Rows']) == 0 and not source_definitions[file]['allowEmpty']:
                raise ValueError('Unexpected empty source evidence: ' + file)
            starts.append(start)
            ends.append(completed)
    if registered != expected_files or not starts:
        raise ValueError('Producer registry and source contract disagree')
    assessed = freshness.evaluate(contract, evidence['Files'], now)
    if (_utc(evidence['StartedAtUtc']) != min(starts)
            or _utc(evidence['CompletedAtUtc']) != max(ends)):
        raise ValueError('Aggregate acquisition evidence mismatch')
    generated = _utc(manifest['GeneratedAtUtc'])
    if generated < max(ends) or generated > future:
        raise ValueError('Invalid preparation timestamp')
    return assessed


def _validate_tables(buffers, definitions, identity, manifest):
    parent_keys = {field: set() for field in PARENTS}
    managed_ids, mailbox_keys = set(), set()
    applications = {}
    for name, definition in definitions.items():
        seen, count = set(), 0
        for row in _csv_rows(buffers[name], name, definition['columns']):
            count += 1
            if any(row[field] != value for field, value in identity.items() if field in row):
                raise ValueError('Reporting identity mismatch: ' + name)
            key = tuple(row[field].strip().casefold() for field in definition['key'])
            if any(not part for part in key) or key in seen:
                raise ValueError('Blank or duplicate immutable key: ' + name)
            seen.add(key)
            for field, parent in PARENTS.items():
                if name == parent + '.csv':
                    parent_keys[field].add(row[field])
            if name == 'DimIntuneManagedDevice.csv':
                managed_ids.add(row['ManagedDeviceId'])
            if name == 'FactMailbox.csv':
                mailbox_keys.add(row['TenantMailboxKey'])
            if name == 'DimDetectedApplication.csv':
                applications[row['AppId'].strip().casefold()] = row
        if count == 0 and not definition['allowEmpty']:
            raise ValueError('Unexpected empty table: ' + name)
        if count != _count(manifest['OutputFiles'][name]['Rows']):
            raise ValueError('Output row count mismatch: ' + name)
    qualified_apps = 'DeviceLinkStatus' in definitions['FactDeviceApplication.csv']['columns']
    application_counts, unresolved_devices = collections.Counter(), collections.Counter()
    # Same parent/FK rules as the preparation guard; applied to retained bytes.
    for name in definitions:
        for row in _csv_rows(buffers[name], name):
            for field, parent in PARENTS.items():
                if name != parent + '.csv' and row.get(field) and row[field] not in parent_keys[field]:
                    raise ValueError('Orphan output relationship: ' + name + '.' + field)
            if (name == 'DeviceHardware.csv' or (name == 'FactDeviceApplication.csv' and not qualified_apps)) and row['ManagedDeviceId'] not in managed_ids:
                raise ValueError('Orphan output managed device: ' + name)
            if name == 'FactDeviceApplication.csv' and qualified_apps:
                aid = row['AppId'].strip().casefold()
                application = applications.get(aid)
                if (row['ApplicationLinkStatus'] != ('Resolved' if application else 'Unresolved')
                        or row['TenantApplicationKey'] != (application['TenantApplicationKey'] if application else '')):
                    raise ValueError('Inconsistent application link qualification')
                resolved = row['ManagedDeviceId'] in managed_ids
                if row['DeviceLinkStatus'] != ('Resolved' if resolved else 'Unresolved'):
                    raise ValueError('Inconsistent application device link qualification')
                application_counts[aid] += 1
                unresolved_devices[aid] += not resolved
            if name == 'FactMailboxHosting.csv' and row['MailboxHostingKey'] not in mailbox_keys:
                raise ValueError('Orphan output mailbox hosting')
            for field in ('ManagerADObjectKey', 'ManagedByADObjectKey'):
                if row.get(field) and row[field] not in parent_keys['TenantADObjectKey']:
                    raise ValueError('Orphan AD source owner/manager relationship')
            if row.get('NestedTenantGroupKey') and row['NestedTenantGroupKey'] not in parent_keys['TenantGroupKey']:
                raise ValueError('Orphan nested Entra group relationship')
    if qualified_apps:
        for aid, row in applications.items():
            count, missing = application_counts[aid], unresolved_devices[aid]
            reported = int(row['ReportedDeviceCount'])
            parts = []
            if reported != count:
                parts.append('Relation count differs')
            if missing:
                parts.append('Device links unresolved')
            status = '; '.join(parts) or 'Complete'
            if (row['RelationshipCoverageStatus'] != status or row['DeviceCount'] != str(count)
                    or row['ExactRelatedDeviceCount'] != str(count)
                    or row['ResolvedDeviceCount'] != str(count - missing)
                    or row['UnresolvedDeviceCount'] != str(missing)):
                raise ValueError('Inconsistent application relationship coverage')


@dataclass(frozen=True)
class Snapshot:
    """Validated immutable bytes, usable even if synchronized files later change."""
    manifest_bytes: bytes
    tables: object
    contract_bytes: bytes = b''

    @property
    def batch_sha256(self):
        return digest(self.manifest_bytes)

    @property
    def manifest(self):
        return _json(self.manifest_bytes)  # A fresh copy, not mutable internal state.

    def iter_rows(self, table):
        return _csv_rows(self.tables[_safe_name(table) + '.csv'], table)


def load_snapshot(root, identity, *, contract_path=CONTRACT, registry_path=REGISTRY,
                  now=None, max_bytes=DEFAULT_MAX_BYTES, expected_manifest_sha256=None):
    """Fail closed; one retained byte snapshot, no writes or fallback to old data."""
    if (not isinstance(identity, dict) or set(identity) != set(IDENTITY_FIELDS)
            or any(not isinstance(value, str) or not value.strip() for value in identity.values())):
        raise ValueError('Complete expected reporting identity is required')
    if type(max_bytes) is not int or max_bytes <= 0:
        raise ValueError('Snapshot memory budget must be a positive integer')
    root = Path(root)
    manifest_bytes = _read(root / MANIFEST, min(max_bytes, 16 * 1024 * 1024))
    if expected_manifest_sha256 and digest(manifest_bytes) != _hash(expected_manifest_sha256):
        raise ValueError('Pinned manifest hash mismatch')
    manifest = _json(manifest_bytes)
    contract_bytes = _read(Path(contract_path), 16 * 1024 * 1024)
    registry_bytes = _read(Path(registry_path), 16 * 1024 * 1024)
    contract, registry = _json(contract_bytes), _json(registry_bytes)
    if (manifest.get('Owner') != OWNER or manifest.get('Status') != 'Validated'
            or manifest.get('Identity') != identity or manifest.get('TenantKey') != identity['TenantKey']):
        raise ValueError('Output is not the owned, validated expected tenant snapshot')
    if (manifest.get('ContractVersion') != contract['version']
            or _hash(manifest.get('ContractSHA256')) != digest(contract_bytes)):
        raise ValueError('Preparation contract mismatch')
    definitions = {}
    for definition in contract['tables']:
        name = _safe_name(definition['name']) + '.csv'
        columns, key = definition['columns'], definition['key']
        if (name in definitions or not columns or len(set(columns)) != len(columns)
                or not key or len(set(key)) != len(key) or not set(key).issubset(columns)
                or type(definition['allowEmpty']) is not bool):
            raise ValueError('Invalid table contract')
        definitions[name] = definition
    expected = set(definitions) | {MANIFEST}
    if set(manifest['OutputFiles']) != set(definitions):
        raise ValueError('Missing or unexpected output manifest entry')
    if {item.name for item in root.iterdir()} != expected:
        raise ValueError('Unexpected or missing output artifacts')
    clock = _utc((now or dt.datetime.now(UTC)).isoformat())
    _evidence(manifest, contract, registry, digest(registry_bytes), clock)
    buffers, total = {}, len(manifest_bytes)
    for name in definitions:
        data = _read(root / name, max_bytes - total)
        total += len(data)
        if digest(data) != _hash(manifest['OutputFiles'][name]['SHA256']):
            raise ValueError('Output hash mismatch: ' + name)
        buffers[name] = data
    _validate_tables(buffers, definitions, identity, manifest)
    # Detect publication during loading; data validation never reopens CSVs.
    if (_read(root / MANIFEST, len(manifest_bytes)) != manifest_bytes
            or {item.name for item in root.iterdir()} != expected):
        raise ValueError('Snapshot publication changed during loading')
    if (_read(Path(contract_path), len(contract_bytes)) != contract_bytes
            or _read(Path(registry_path), len(registry_bytes)) != registry_bytes):
        raise ValueError('Consumer contract changed during loading')
    # Validation can be lengthy: do not accept evidence that expired meanwhile.
    if now is None:
        _evidence(manifest, contract, registry, digest(registry_bytes), dt.datetime.now(UTC))
    return Snapshot(manifest_bytes, MappingProxyType(buffers), contract_bytes)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', required=True)
    for field in IDENTITY_FIELDS:
        parser.add_argument('--' + re.sub(r'(?<!^)(?=[A-Z])', '-', field).lower(), required=True)
    parser.add_argument('--contract', type=Path, default=CONTRACT)
    parser.add_argument('--registry', type=Path, default=REGISTRY)
    parser.add_argument('--max-bytes', type=int, default=DEFAULT_MAX_BYTES)
    parser.add_argument('--expected-manifest-sha256')
    args = parser.parse_args()
    identity = {field: getattr(args, re.sub(r'(?<!^)(?=[A-Z])', '_', field).lower()) for field in IDENTITY_FIELDS}
    try:
        snapshot = load_snapshot(args.root, identity, contract_path=args.contract, registry_path=args.registry,
                                 max_bytes=args.max_bytes, expected_manifest_sha256=args.expected_manifest_sha256)
        print(json.dumps({'Status': 'ValidatedBufferedSnapshot', 'ReaderVersion': VERSION,
                          'BatchSHA256': snapshot.batch_sha256, 'Tables': len(snapshot.tables),
                          'Bytes': sum(map(len, snapshot.tables.values())), 'ReportModified': False,
                          'FreshnessWarnings':freshness.evaluate(_json(args.contract.read_bytes()),
                              snapshot.manifest['SourceEvidence']['Files'], dt.datetime.now(UTC))['Warnings']}))
    except (ValueError, OSError, KeyError, TypeError) as error:
        parser.exit(1, 'Snapshot rejected: ' + str(error) + '\n')


if __name__ == '__main__':
    main()
