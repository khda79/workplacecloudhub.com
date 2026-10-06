"""Read-only transfer plan for a pinned, already validated current CMDB cohort.

No CSV recalculation, raw-source access, authentication or publication. Hashes
bind this transfer to the generator's validated bytes; this is not BI qualification.
"""
import argparse
import datetime as dt
import hashlib
import json
import re
from pathlib import Path

import cmdb_freshness

VERSION = '0.1.0'
CONTRACT = Path(__file__).with_name('cmdb-prepared-contract.json.txt')
MANIFEST = 'current.json.txt'


def digest(path):
    result = hashlib.sha256()
    with path.open('rb') as stream:
        for chunk in iter(lambda: stream.read(4 * 1024 * 1024), b''):
            result.update(chunk)
    return result.hexdigest().upper()


def check(condition, message):
    # Explicit exceptions: validation must also work when Python is optimized.
    if not condition:
        raise ValueError(message)


def plan(root, identity, expected_hash='', now=None, contract_path=CONTRACT):
    root = Path(root).absolute()
    now = now or dt.datetime.now(dt.timezone.utc)
    check(root.name == 'DATA-POWERBI-CMDB' and root.is_dir()
          and not root.is_symlink() and root.resolve() == root,
          'Use the unlinked current DATA-POWERBI-CMDB directory')
    contract_path = Path(contract_path)
    contract_hash = digest(contract_path)
    contract = json.loads(contract_path.read_text(encoding='utf-8-sig'))
    names = [table['name'] + '.csv' for table in contract['tables']]
    check(len(names) == len(set(names)) == 46
          and all(Path(name).name == name and '/' not in name and '\\' not in name for name in names),
          'Unexpected CMDB table contract')
    pointer = root / MANIFEST
    entries = list(root.iterdir())
    check(all(entry.is_file() and not entry.is_symlink() for entry in entries)
          and {entry.name for entry in entries} == set(names) | {MANIFEST},
          'Missing, extra or linked prepared artifacts')
    manifest_hash = digest(pointer)
    check(not expected_hash or (re.fullmatch(r'[A-Fa-f0-9]{64}', expected_hash)
          and manifest_hash == expected_hash.upper()), 'Expected manifest SHA256 mismatch')
    manifest = json.loads(pointer.read_text(encoding='utf-8-sig'))
    check(manifest.get('Owner') == 'SmartInventory-CMDB-Prepared'
          and manifest.get('Status') == 'Validated', 'Unvalidated or foreign prepared snapshot')
    fields = ('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')
    check(all(isinstance(identity.get(field), str) and identity[field].strip() for field in fields)
          and all(manifest.get('Identity', {}).get(field) == identity[field] for field in fields)
          and manifest.get('TenantKey') == identity['TenantKey'], 'Prepared reporting identity mismatch')
    check(manifest.get('ContractVersion') == contract['version']
          and manifest.get('ContractSHA256') == contract_hash, 'Prepared contract mismatch')
    generated = cmdb_freshness.utc(manifest['GeneratedAtUtc'])
    check(generated <= now + dt.timedelta(minutes=5), 'Future prepared generation time')
    records = manifest['OutputFiles']
    check(isinstance(records, dict) and set(records) == set(names), 'Manifest output set mismatch')
    files = []
    for name in sorted(names):
        record = records[name]
        check(type(record.get('Rows')) is int and record['Rows'] >= 0
              and isinstance(record.get('SHA256'), str)
              and re.fullmatch(r'[A-Fa-f0-9]{64}', record['SHA256']), 'Invalid output declaration: ' + name)
        path = root / name
        check(digest(path) == record['SHA256'].upper(), 'CSV hash mismatch: ' + name)
        files.append(dict(Name=name, SHA256=record['SHA256'].upper(), Bytes=path.stat().st_size))
    freshness = cmdb_freshness.evaluate(contract, manifest['SourceEvidence']['Files'], now)
    check(digest(pointer) == manifest_hash and digest(contract_path) == contract_hash,
          'Manifest or contract changed during transfer verification')
    files.append(dict(Name=MANIFEST, SHA256=manifest_hash, Bytes=pointer.stat().st_size))
    return dict(Status='VerifiedForTransfer', TransferVersion=VERSION,
                PreparedRoot=str(root), ManifestSHA256=manifest_hash,
                ContractVersion=contract['version'], GeneratedAtUtc=manifest['GeneratedAtUtc'],
                EarliestSourceExpiryUtc=freshness['ExpiresAtUtc'], FreshnessWarnings=freshness['Warnings'],
                Files=files, CsvFiles=len(names),
                Scope='Validated byte integrity and acquisition age only; no recalculation or Power BI qualification')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', required=True)
    parser.add_argument('--tenant-key', required=True)
    parser.add_argument('--organization-key', required=True)
    parser.add_argument('--environment-key', required=True)
    parser.add_argument('--tenant-id', required=True)
    parser.add_argument('--expected-manifest-sha256', default='')
    args = parser.parse_args()
    identity = dict(TenantKey=args.tenant_key, OrganizationKey=args.organization_key,
                    EnvironmentKey=args.environment_key, TenantId=args.tenant_id)
    print(json.dumps(plan(args.root, identity, args.expected_manifest_sha256), indent=2))


if __name__ == '__main__':
    main()
