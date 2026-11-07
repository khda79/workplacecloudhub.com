"""Version 0.2.0: read-only Plan 2 migration evidence from one prepared batch.

The output is an estimate, not an assignment or purchase instruction. The
prepared pointer, manifest and every input are verified before and after read.
Unknown identity, classification, size and license evidence remain review.
"""
import argparse
import csv
import hashlib
import json
import re
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path


FILES = (
    'Source-Exchange_OnPrem_Mailboxes_AllDomains.csv',
    'MailboxMigrationExceptions.csv', 'ADUserSource.csv',
    'FactUserLicense.csv', 'FactUserServicePlan.csv', 'DimLicenseSku.csv',
)
KEY_PATTERN = re.compile(r'^\d{8}T\d{9}Z-[0-9a-f]{8}$')


def digest(path):
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest().upper()


def key(value):
    return (value or '').strip().lower().strip('{}')


def read_rows(path, tenant, required, expected_rows=None):
    with path.open('r', encoding='utf-8-sig', newline='') as stream:
        reader = csv.DictReader(stream)
        if not reader.fieldnames or len(reader.fieldnames) != len(set(reader.fieldnames)) or not set(required).issubset(reader.fieldnames):
            raise ValueError('Missing or duplicate columns: ' + path.name)
        count = 0
        for row in reader:
            count += 1
            if None in row or row.get('TenantKey') != tenant:
                raise ValueError('Cross-tenant row: ' + path.name)
            yield row
        if expected_rows is not None and count != expected_rows:
            raise ValueError('Prepared row count mismatch: ' + path.name)


def evaluate(mailboxes, exceptions, ad_rows, licenses, plans, skus):
    retained, pending_ids, pending_addresses = set(), set(), set()
    for row in exceptions:
        identity, candidate = key(row['Mailbox Identity']), key(row['Candidate Identity'])
        if row['Retained'] == '1' and row['Requires Review'] == '0' and identity.startswith('local_mailboxes|'):
            retained.add(identity)
        elif row['Requires Review'] == '1' or row['Retained'] == '1':
            pending_ids.update(x for x in (identity, candidate) if x)
            if key(row['Primary SMTP Address']):
                pending_addresses.add(key(row['Primary SMTP Address']))
        elif row['Retained'] not in ('0', '1') or row['Requires Review'] not in ('0', '1'):
            raise ValueError('Invalid exception state')

    ad = {}
    for row in ad_rows:
        obj = key(row['ObjectGUID'])
        if not obj or obj in ad:
            raise ValueError('Missing or duplicate AD object identity')
        ad[obj] = row
    sku_by_key = {}
    for row in skus:
        sku = key(row['TenantSkuKey'])
        if not sku or sku in sku_by_key:
            raise ValueError('Missing or duplicate SKU identity')
        part, sku_id = key(row['SkuPartNumber']), key(row.get('SkuId'))
        if sku_id in ('dcf0408c-aaec-446c-afd4-43e3683943ea','c2fe850d-fbbb-4858-b67d-bd0c6e746da3','602e6573-55a3-46b1-a1a0-cc267991501a'):
            part = 'spe_e3'
        elif sku_id == 'f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87':
            part = 'spe_f1'
        sku_by_key[sku] = part

    selected, seen_mail, cloud_to_mail = [], set(), {}
    for row in mailboxes:
        if key(row['RecipientType']) != 'sharedmailbox':
            continue
        obj = key(row['ObjectGUID'])
        mailbox = key(row['ExchangeGuid']) or obj
        if not mailbox or mailbox in seen_mail:
            raise ValueError('Missing or duplicate shared mailbox identity')
        seen_mail.add(mailbox)
        ident = 'local_mailboxes|' + obj if obj else ''
        address = key(row['PrimarySMTPaddress'])
        entry = {'mailbox': mailbox, 'object': obj, 'identity': ident, 'row': row, 'cloud': ''}
        if ident in retained:
            entry['state'] = 'retained'
        elif ident in pending_ids or (address and address in pending_addresses):
            entry['state'] = 'review'
        elif key(row['NativeIdentityStatus']) != 'observed' or not obj or obj not in ad:
            entry['state'] = 'review'
        else:
            account = ad[obj]
            classification = key(account['AccountType'])
            if classification in ('service', 'service account'):
                entry['state'] = 'retained'
            elif classification not in ('named account', 'ext account', 'generic account', 'admin account', 'system account', 'resource account', 'shared mailbox'):
                entry['state'] = 'review'
            elif account['CloudMatchStatus'] != 'Unique native SID' or not key(account['TenantUserKey']):
                entry['state'] = 'review'
            else:
                entry['state'] = 'candidate'
                entry['cloud'] = key(account['TenantUserKey'])
                cloud_to_mail.setdefault(entry['cloud'], []).append(entry)
        selected.append(entry)
    for entries in cloud_to_mail.values():
        if len(entries) != 1:
            for entry in entries:
                entry['state'] = 'review'

    users = {x['cloud'] for x in selected if x['state'] == 'candidate'}
    by_user, by_plan = {u: set() for u in users}, {u: [] for u in users}
    for row in licenses:
        user = key(row['TenantUserKey'])
        if user in users:
            by_user[user].add((key(row['TenantSkuKey']), key(row['AssignmentState']), key(row['UserLinkStatus'])))
    for row in plans:
        user = key(row['TenantUserKey'])
        if user in users:
            by_plan[user].append((key(row['TenantSkuKey']), key(row['ServicePlanName']), key(row['IsEnabled']), key(row['AssignmentState']), key(row['UserLinkStatus'])))

    result = dict(Required=0, AssignedCoverage=0, ToComplete=0, CoverageReview=0,
                  Excluded=0, IdentityOrRegisterReview=0, SizeReview=0, Under50=0,
                  E3ToPlan2Review=0, E3Under50Review=0, F3Under50Review=0)
    for entry in selected:
        if entry['state'] == 'retained':
            result['Excluded'] += 1
            continue
        if entry['state'] != 'candidate':
            result['IdentityOrRegisterReview'] += 1
            continue
        raw = (entry['row']['TotalItemSize-In-MB'] or '').strip().replace(',', '.')
        try:
            size = Decimal(raw)
            if not size.is_finite() or size < 0:
                raise InvalidOperation
        except InvalidOperation:
            result['SizeReview'] += 1
            continue
        if size in (Decimal(51200), Decimal(102400)) or size > 102400:
            result['SizeReview'] += 1
            continue
        user = entry['cloud']
        assigned = by_user[user]
        active_keys = {sku for sku, state, link in assigned if state in ('active', 'success') and link == 'resolved' and sku in sku_by_key}
        effective = {sku_by_key.get(sku, '') for sku, state, link in assigned if state in ('active', 'success') and link == 'resolved'}
        uncertain = any(state not in ('active', 'success') or link != 'resolved' or sku not in sku_by_key for sku, state, link in assigned)
        if size < 51200:
            result['Under50'] += 1
            result['E3Under50Review'] += int('spe_e3' in effective)
            result['F3Under50Review'] += int('spe_f1' in effective)
            continue
        result['Required'] += 1
        result['E3ToPlan2Review'] += int('spe_e3' in effective)
        states = by_plan[user]
        if any(sku in active_keys and name == 'exchange_s_enterprise' and enabled == 'true' and state == 'success' and link == 'resolved' for sku, name, enabled, state, link in states):
            result['AssignedCoverage'] += 1
        elif uncertain or not states or any(sku not in active_keys or enabled not in ('true','false') or link != 'resolved' or state not in ('success', 'active') for sku, _, enabled, state, link in states):
            result['CoverageReview'] += 1
        else:
            result['ToComplete'] += 1
    assert result['Required'] == result['AssignedCoverage'] + result['ToComplete'] + result['CoverageReview']
    return result


def collect(root, tenant):
    root = Path(root)
    pointer_path = root / 'current.json.txt'
    pointer_hash = digest(pointer_path)
    pointer = json.loads(pointer_path.read_text(encoding='utf-8-sig'))
    batch = pointer.get('BatchId', '')
    if pointer.get('SchemaVersion') != 1 or pointer.get('TenantKey') != tenant or not KEY_PATTERN.fullmatch(batch):
        raise ValueError('Invalid prepared pointer identity')
    folder = root / 'batches' / batch
    manifest_path = folder / 'batch.json.txt'
    manifest_hash = digest(manifest_path)
    if manifest_hash != str(pointer.get('ManifestSHA256', '')).upper():
        raise ValueError('Prepared manifest does not match pointer')
    result = collect_manifest(manifest_path, tenant)
    if result['BatchId'] != batch:
        raise ValueError('Prepared batch identity mismatch')
    if digest(pointer_path) != pointer_hash or digest(manifest_path) != manifest_hash:
        raise ValueError('Prepared pointer or manifest changed during calculation')
    return result


def collect_manifest(manifest_path, tenant):
    """Qualify an in-progress batch from its frozen draft manifest, without a pointer."""
    manifest_path = Path(manifest_path)
    folder = manifest_path.parent
    manifest_hash = digest(manifest_path)
    manifest = json.loads(manifest_path.read_text(encoding='utf-8-sig'))
    batch = manifest.get('BatchId', '')
    if not KEY_PATTERN.fullmatch(batch):
        raise ValueError('Invalid prepared batch identity')
    if manifest.get('BatchId') != batch or manifest.get('TenantKey') != tenant:
        raise ValueError('Prepared batch identity mismatch')
    provenance = manifest.get('Provenance', {})
    audit = provenance.get('MailboxMigrationExceptions', {}).get('Audit', {})
    native = provenance.get('NativeEvidence', {}).get('Audit', {})
    if audit.get('Source Status') != 'Qualified' or audit.get('Identity', {}).get('TenantKey') != tenant or native.get('Identity', {}).get('TenantKey') != tenant or manifest.get('LicenseReportQualification', {}).get('State') != 'Qualified':
        raise ValueError('Prepared exception, native or license evidence is not qualified')
    native_files = native.get('Evidence', {}).get('Files', [])
    source_names = ('Exchange_OnPrem_Mailboxes_AllDomains.csv','M365_Licenses_Tenant.csv',
                    'M365_Licenses_Users.csv','M365_Licenses_ServicePlans_Catalog.csv',
                    'M365_Licenses_UserServicePlanStates.csv','AD_Users_AllDomains.csv')
    source_proofs = {}
    for name in source_names:
        proof = [x for x in native_files if x.get('File') == name]
        if len(proof) != 1 or proof[0].get('Status') != 'Success' or proof[0].get('IsPartialInventory') is not False or proof[0].get('Errors') != 0:
            raise ValueError('Prepared source receipt is incomplete: ' + name)
        source_proofs[name] = proof[0]
    run_id = manifest.get('LicenseReportQualification', {}).get('ReceiptRunId')
    if not run_id or any(source_proofs[name].get('RunId') != run_id for name in source_names[1:5]):
        raise ValueError('Prepared license source runs do not match')
    metadata = {}
    for name in FILES:
        entries = [x for x in manifest.get('Files', []) + manifest.get('Artifacts', []) if x.get('File') == name]
        if len(entries) != 1 or not re.fullmatch(r'[0-9a-fA-F]{64}', str(entries[0].get('SHA256', ''))):
            raise ValueError('Prepared file record missing or ambiguous: ' + name)
        path = folder / name
        if digest(path) != entries[0]['SHA256'].upper():
            raise ValueError('Prepared file hash mismatch: ' + name)
        metadata[name] = entries[0]['SHA256'].upper()
    if metadata[FILES[0]] != str(source_proofs['Exchange_OnPrem_Mailboxes_AllDomains.csv'].get('SHA256', '')).upper():
        raise ValueError('Prepared mailbox artifact differs from its native source receipt')
    requirements = (
        ('TenantKey','RecipientType','ObjectGUID','ExchangeGuid','NativeIdentityStatus','PrimarySMTPaddress','TotalItemSize-In-MB'),
        ('TenantKey','Mailbox Identity','Candidate Identity','Primary SMTP Address','Retained','Requires Review','Status'),
        ('TenantKey','ObjectGUID','TenantUserKey','CloudMatchStatus','AccountType'),
        ('TenantKey','TenantUserKey','TenantSkuKey','AssignmentState','UserLinkStatus'),
        ('TenantKey','TenantUserKey','TenantSkuKey','ServicePlanName','IsEnabled','AssignmentState','UserLinkStatus'),
        ('TenantKey','TenantSkuKey','SkuPartNumber','SkuId'),
    )
    row_counts = {}
    for record in manifest.get('Files', []):
        if record.get('File') in FILES:
            if type(record.get('Rows')) is not int or record['Rows'] < 0:
                raise ValueError('Prepared file row count missing')
            row_counts[record['File']] = record['Rows']
    result = evaluate(*(read_rows(folder / name, tenant, columns, row_counts.get(name)) for name, columns in zip(FILES, requirements)))
    for name, expected in metadata.items():
        if digest(folder / name) != expected:
            raise ValueError('Prepared file changed during calculation: ' + name)
    if digest(manifest_path) != manifest_hash:
        raise ValueError('Prepared manifest changed during calculation')
    return dict(BatchId=batch, BatchCreatedUtc=manifest.get('CreatedUtc', ''), ExceptionSnapshotDate=audit.get('Snapshot Date', ''), Counts=result)


def main():
    parser = argparse.ArgumentParser()
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument('--root')
    source.add_argument('--batch-manifest')
    parser.add_argument('--tenant', required=True)
    args = parser.parse_args()
    try:
        result = collect(args.root, args.tenant) if args.root else collect_manifest(args.batch_manifest, args.tenant)
        print(json.dumps(result, separators=(',', ':')))
    except (OSError, ValueError, KeyError, TypeError) as exc:
        print('Shared mailbox evidence not qualified: ' + str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
