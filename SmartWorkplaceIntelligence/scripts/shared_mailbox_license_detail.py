"""Version 0.1.0: reconcile shared mailbox detail with a license snapshot.

The input files are accepted only when their exact digests match the snapshot.
The result is review evidence, never an instruction to change an assignment.
"""
import csv
import argparse
import hashlib
import json
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path


def sha256(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest().upper()


def normalized(value):
    return (value or '').strip().lower()


def size_gb(value):
    try:
        value = (value or '').strip()
        if not re.fullmatch(r'\d+(?:[.,]\d+)?', value):
            return None
        value = value.replace(',', '.')
        number = Decimal(value)
        return number if number.is_finite() and number >= 0 else None
    except (InvalidOperation, ValueError):
        return None


def evaluate(snapshot, tenant, license_rows, mailbox_rows):
    """Build details with the collector's native identity and guard rules."""
    if snapshot.get('TenantKey') != tenant:
        raise ValueError('Snapshot tenant mismatch')
    products = {'spe_e3': 'Microsoft 365 E3', 'spe_f1': 'Microsoft 365 F3', 'm365_f1': 'Microsoft 365 F1', 'spe_e5': 'Microsoft 365 E5'}
    e3_guids = {'dcf0408c-aaec-446c-afd4-43e3683943ea', 'c2fe850d-fbbb-4858-b67d-bd0c6e746da3', '602e6573-55a3-46b1-a1a0-cc267991501a'}
    f3_guids = {'f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87'}
    e5_guids = {'18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e', '3271cf8e-2be5-4a09-a549-70fd05baaa17', '1e988bf3-8b7c-4731-bec0-4e2a2946600c', '6ee4114a-9b2d-4577-9e7a-49fa43d222d3', '90277bc7-a6fe-4181-99d8-712b08b8d32b', 'a640eead-25f6-4bec-97e3-23cfd382d7c2'}
    assignments = set()
    for row in license_rows:
        if row.get('TenantKey') != tenant:
            raise ValueError('Cross-tenant license row')
        user, sku, sku_id = normalized(row.get('UserId')), normalized(row.get('SkuPartNumber')), normalized(row.get('SkuId'))
        # Match collector precedence: F1, F3, E3, then E5.
        if sku in ('m365_f1', 'm365_f1_comm') or sku_id == '0666269f-b167-4c5b-a76f-fc574f2b1118':
            product = products['m365_f1']
        elif sku == 'spe_f1' or sku_id in f3_guids:
            product = products['spe_f1']
        elif sku == 'spe_e3' or sku_id in e3_guids:
            product = products['spe_e3']
        elif sku == 'spe_e5' or sku_id in e5_guids:
            product = products['spe_e5']
        else:
            product = None
        if product:
            if not user:
                raise ValueError('Missing licensed user identity')
            assignments.add((user, product))
    users = {user for user, _ in assignments}
    mailboxes = {}
    for row in mailbox_rows:
        if row.get('TenantKey') != tenant:
            raise ValueError('Cross-tenant mailbox row')
        user = normalized(row.get('ExternalDirectoryObjectId'))
        if user not in users:
            continue
        if user in mailboxes:
            raise ValueError('Duplicate mailbox identity')
        mailboxes[user] = row
    details = []
    counts = {p: {'SharedLicensed': 0, 'SharedUnder50': 0, 'SharedEligible': 0, 'SharedE3Plan2Review': 0} for p in products.values()}
    for user, product in sorted(assignments):
        row = mailboxes.get(user)
        if not row or row.get('NativeIdentityStatus', '').strip().lower() != 'observed' or row.get('RecipientTypeDetails', '').strip().lower() != 'sharedmailbox':
            continue
        size = size_gb(row.get('TotalItemSizeGB'))
        archive = normalized(row.get('ArchiveStatus'))
        litigation = normalized(row.get('LitigationHoldEnabled'))
        retention = normalized(row.get('RetentionHoldEnabled'))
        valid = size is not None and archive in ('none', 'disabled', 'false', 'active', 'enabled', 'true') and litigation in ('true', 'false') and retention in ('true', 'false')
        eligible = valid and size < 50 and archive in ('none', 'disabled', 'false') and litigation == 'false' and retention == 'false'
        plan2 = product == 'Microsoft 365 E3' and size is not None and Decimal(50) < size < Decimal(100)
        c = counts[product]
        c['SharedLicensed'] += 1
        c['SharedUnder50'] += int(size is not None and size < 50)
        c['SharedEligible'] += int(eligible)
        c['SharedE3Plan2Review'] += int(plan2)
        details.append({'UserId': user, 'DisplayName': row.get('DisplayName', ''), 'PrimarySmtpAddress': row.get('PrimarySmtpAddress', ''), 'License': product, 'MailboxSizeGB': str(size) if size is not None else None, 'ArchiveStatus': row.get('ArchiveStatus', ''), 'LitigationHoldEnabled': row.get('LitigationHoldEnabled', ''), 'RetentionHoldEnabled': row.get('RetentionHoldEnabled', ''), 'Review': 'E3 to Plan 2' if plan2 else 'Removal candidate' if eligible else 'Review required'})
    for product, actual in counts.items():
        matches = [x for x in snapshot.get('Products', []) if x.get('Product') == product]
        if len(matches) != 1:
            raise ValueError('Missing or duplicate snapshot product')
        expected = matches[0].get('Counts', {})
        if matches[0].get('UsageAvailable') is not True or expected.get('SharedUnknown') != 0:
            raise ValueError('Snapshot mailbox evidence is incomplete: ' + product)
        for metric, value in actual.items():
            if expected.get(metric) != value:
                raise ValueError('Snapshot does not reconcile: ' + product + ' ' + metric)
    return details


def from_files(snapshot_path, license_path, mailbox_path, tenant):
    snapshot_path = Path(snapshot_path)
    snapshot_hash = sha256(snapshot_path)
    with snapshot_path.open(encoding='utf-8-sig') as stream:
        snapshot = json.load(stream)
    if snapshot.get('Rendering', {}).get('SharedSourceReady') is not True:
        raise ValueError('Snapshot shared mailbox source is not ready')
    input_hashes = {}
    for path in (Path(license_path), Path(mailbox_path)):
        entries = [x for x in snapshot.get('Sources', []) if x.get('Name') == path.name]
        current_hash = sha256(path)
        if len(entries) != 1 or entries[0].get('Ready') is not True or entries[0].get('Provisional') is not False or entries[0].get('SHA256', '').upper() != current_hash:
            raise ValueError('Source is not qualified or does not match snapshot: ' + path.name)
        input_hashes[path] = current_hash
    with Path(license_path).open(encoding='utf-8-sig', newline='') as licenses, Path(mailbox_path).open(encoding='utf-8-sig', newline='') as mailboxes:
        lic_reader, mb_reader = csv.DictReader(licenses), csv.DictReader(mailboxes)
        for reader, required in ((lic_reader, {'TenantKey', 'UserId', 'SkuPartNumber', 'SkuId'}), (mb_reader, {'TenantKey', 'ExternalDirectoryObjectId', 'NativeIdentityStatus', 'RecipientTypeDetails', 'TotalItemSizeGB', 'ArchiveStatus', 'LitigationHoldEnabled', 'RetentionHoldEnabled'})):
            names = reader.fieldnames or []
            if len(names) != len(set(names)) or not required.issubset(names):
                raise ValueError('Source schema is missing or duplicates required columns')
        result = evaluate(snapshot, tenant, lic_reader, mb_reader)
    if sha256(snapshot_path) != snapshot_hash or any(sha256(path) != digest for path, digest in input_hashes.items()):
        raise ValueError('Source changed while reading')
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--snapshot', required=True)
    parser.add_argument('--licenses', required=True)
    parser.add_argument('--mailboxes', required=True)
    parser.add_argument('--tenant', required=True)
    args = parser.parse_args()
    print(json.dumps(from_files(args.snapshot, args.licenses, args.mailboxes, args.tenant), ensure_ascii=False, separators=(',', ':')))
