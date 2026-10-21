"""Mailbox retention decisions on a receipt-qualified, immutable planning snapshot.

No collection, cloud writes, automatic binding approval, or publication. SMTP
bootstrap candidates require an operator-confirmed private binding registry.
"""
import argparse
import collections
import csv
import datetime as dt
import hashlib
import json
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from workplace_mailbox_identity import reconcile, guid, value

WORKBOOK = 'SmartWorkplaceIntelligence-MailboxMigrationExceptions.xlsx'
HEADERS = ['Primary SMTP Address', 'Retain On-premises', 'Reason', 'Decision Owner',
           'Decision Date', 'Review Date', 'Exchange GUID', 'AD Object GUID',
           'Display Name', 'Source Reference']
SCOPE = {'usermailbox', 'sharedmailbox', 'roommailbox', 'equipmentmailbox', 'linkedmailbox'}
IDENTITY = ('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')
DECISION_COLUMNS = list(IDENTITY) + ['Decision Key', 'Source Rows', 'Primary SMTP Address',
    'Retain On-premises', 'Reason', 'Decision Owner', 'Decision Date', 'Review Date',
    'Source Reference', 'Mailbox Identity', 'Candidate Identity', 'Source Domain',
    'Domain Key', 'Status', 'Requires Review', 'Retained', 'Documentation Warnings',
    'Snapshot Date']
SUMMARY_COLUMNS = list(IDENTITY) + ['Domain Key', 'Source Domain', 'Source Status',
    'Snapshot Date', 'Inventory Count', 'Retained Count', 'Review Count',
    'Global Review Count', 'Remaining Count', 'Documentation Warning Count',
    'Workbook SHA256', 'Workbook Modified UTC', 'Captured UTC', 'Binding SHA256',
    'Source Rows']


def key(raw):
    return str(raw or '').strip().casefold()


def digest(raw):
    return hashlib.sha256(raw).hexdigest()


def read_register(path):
    """Read the actual structured table, never UsedRange or a fixed row limit."""
    with zipfile.ZipFile(path) as archive:
        def xml(name):
            if archive.getinfo(name).file_size > 32 * 1024 * 1024:
                raise ValueError('Oversized workbook XML')
            raw = archive.read(name)
            if len(raw) > 32 * 1024 * 1024 or b'<!DOCTYPE' in raw.upper():
                raise ValueError('Unsafe or oversized workbook XML')
            return ET.fromstring(raw)

        if any(n.endswith('vbaProject.bin') or n.startswith('xl/externalLinks/') for n in archive.namelist()):
            raise ValueError('Macros and external links are not supported')
        workbook = xml('xl/workbook.xml')
        if workbook.find('.//{*}workbookPr') is not None and workbook.find('.//{*}workbookPr').get('date1904') in ('1', 'true'):
            raise ValueError('Use the template Excel 1900 date system')
        rels = {r.get('Id'): r.get('Target') for r in xml('xl/_rels/workbook.xml.rels')}
        sheets = [s for s in workbook.findall('.//{*}sheet') if s.get('name') == 'Exceptions']
        if len(sheets) != 1:
            raise ValueError('Exactly one Exceptions sheet is required')
        relation = next(v for k, v in sheets[0].attrib.items() if k.endswith('}id'))
        def target(base, relative):
            resolved = posixpath.normpath(posixpath.join(posixpath.dirname(base), relative)) if not relative.startswith('/') else relative.lstrip('/')
            if not resolved.startswith('xl/') or '..' in resolved.split('/'):
                raise ValueError('Unsafe workbook relationship')
            return resolved
        sheet_path = target('xl/workbook.xml', rels[relation])
        sheet = xml(sheet_path)
        sheet_rels = xml(posixpath.dirname(sheet_path) + '/_rels/' + posixpath.basename(sheet_path) + '.rels')
        linked = {r.get('Id'): r.get('Target') for r in sheet_rels if r.get('TargetMode') != 'External'}
        tables = []
        for part in sheet.findall('.//{*}tablePart'):
            rid = next(v for k, v in part.attrib.items() if k.endswith('}id'))
            table = xml(target(sheet_path, linked[rid]))
            if table.get('displayName') == 'MailboxMigrationExceptions':
                tables.append(table)
        if len(tables) != 1:
            raise ValueError('Exactly one MailboxMigrationExceptions table is required')
        table = tables[0]
        input_headers = [c.get('name') for c in table.findall('.//{*}tableColumn')]
        if input_headers not in (HEADERS, HEADERS[:-1]) or table.get('totalsRowCount', '0') != '0':
            raise ValueError('Invalid exception table columns or totals row')
        last_column = 'J' if len(input_headers) == 10 else 'I'
        match = re.fullmatch(r'A7:' + last_column + r'(\d+)', table.get('ref', ''))
        if not match or not 8 <= int(match[1]) <= 1048576:
            raise ValueError('Exception table must start at A7:' + last_column + '8')
        last_row = int(match[1])
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            strings = [''.join(e.itertext()) for e in xml('xl/sharedStrings.xml')]
        result, header_seen = [], False
        for row in sheet.findall('.//{*}sheetData/{*}row'):
            number = int(row.get('r'))
            if not 7 <= number <= last_row:
                continue
            cells = {}
            for cell in row:
                ref = re.fullmatch(r'([A-' + last_column + r'])(\d+)', cell.get('r', ''))
                if not ref:
                    continue
                if cell.find('{*}f') is not None:
                    raise ValueError('Decision table cells must contain values, not formulas')
                raw = cell.findtext('{*}v', '')
                kind = cell.get('t')
                if kind == 's':
                    raw = strings[int(raw)]
                elif kind == 'inlineStr':
                    raw = ''.join(cell.find('{*}is').itertext())
                elif kind == 'e':
                    raise ValueError('Workbook error cell')
                column = ord(ref[1]) - ord('A')
                if number > 7 and column in (4, 5) and raw and kind not in ('s', 'inlineStr', 'str'):
                    serial = float(raw)
                    if not 61 <= serial < 2958466 or serial != int(serial):
                        raise ValueError('Invalid Excel decision date')
                    raw = (dt.date(1899, 12, 30) + dt.timedelta(days=int(serial))).isoformat()
                cells[column] = raw
            values = [cells.get(i, '') for i in range(len(input_headers))]
            if number == 7:
                header_seen = True
                if values != input_headers:
                    raise ValueError('Worksheet headers differ from table metadata')
            elif any(str(v).strip() for v in values):
                result.append(dict(dict.fromkeys(HEADERS, ''), **dict(zip(input_headers, values)),
                                   **{'Source Row': number}))
        if not header_seen:
            raise ValueError('Worksheet header row is missing')
        return result


def validate_bindings(document, identity):
    if document.get('Version') != 1 or document.get('Identity') != identity:
        raise ValueError('Binding registry tenant/environment identity mismatch')
    bindings = {}
    for row in document.get('Bindings', []):
        address, native = key(row.get('PrimarySmtpAddress')), row.get('NativeIdentity', '')
        if (not address or address in bindings or not re.fullmatch(r'(local_mailboxes|mailboxes)\|[0-9a-f-]{36}', native)
                or not guid(native.split('|')[1]) or not str(row.get('ConfirmedBy', '')).strip()):
            raise ValueError('Invalid, unconfirmed or repeated binding')
        confirmed = dt.datetime.fromisoformat(row.get('ConfirmedAtUtc', '').replace('Z', '+00:00'))
        if confirmed.tzinfo is None:
            raise ValueError('Binding confirmation must carry a UTC offset')
        bindings[address] = native
    return bindings


def resolve(rows, populations, identity, bindings=None, as_of=None):
    """Pure, tenant-scoped matching. All unresolved groups block only their scope."""
    as_of = as_of or dt.date.today()
    bindings = bindings or {}
    for records in populations.values():
        for row in records:
            if value(row, 'TenantKey') != identity['TenantKey'] or any(
                    value(row, f) and value(row, f) != identity[f] for f in IDENTITY[1:]):
                raise ValueError('Mailbox inventory tenant/environment identity mismatch')
    _, observations, _ = reconcile(populations)
    local_targets = {item['target'] for item in observations.values()
                     if item['identity'][0] == 'local_mailboxes'}
    by_smtp, by_exchange, by_ad = (collections.defaultdict(set) for _ in range(3))
    native = {'|'.join(k): v for k, v in observations.items()}
    for token, item in native.items():
        # Resolve to the same native physical identity only through existing technical links.
        target = '|'.join(item['target']) if item['target'] else token
        if item['smtp']:
            by_smtp[item['smtp']].add(target)
        if item['exchange']:
            by_exchange[item['exchange']].add(target)
        if item['identity'][0] != 'mailboxes' and guid(item['identity'][1]):
            by_ad[guid(item['identity'][1])].add(target)
    # N has exactly the existing local detailed migration inventory scope and key rule.
    inventory = collections.Counter()
    population_keys = set()
    for row in populations['local_mailboxes']:
        identifier = key(value(row, 'ExchangeGuid')) or key(value(row, 'ObjectGUID'))
        if not identifier or identifier in population_keys:
            raise ValueError('Missing or duplicate detailed local mailbox identity')
        population_keys.add(identifier)
        if key(value(row, 'RecipientType')) in SCOPE:
            inventory[key(value(row, 'DomainName')) or '__unknown__'] += 1
    pending = []
    for row in rows:
        address, decision = key(row['Primary SMTP Address']), key(row['Retain On-premises'])
        resolved, candidate, status = '', '', ''
        pins = []
        for column, index in [('Exchange GUID', by_exchange), ('AD Object GUID', by_ad)]:
            raw = str(row[column]).strip()
            if raw:
                token = guid(raw)
                found = index.get(token, set()) if token else set()
                if not token or len(found) != 1:
                    status = 'Invalid or unmatched technical identity'
                    break
                pins.append(next(iter(found)))
        if not status and pins:
            if len(set(pins)) != 1:
                status = 'Identity conflict'
            else:
                resolved = pins[0]
                if address in bindings and bindings[address] != resolved:
                    status = 'Identity conflict with confirmed binding'
        elif not status and address in bindings:
            resolved = bindings[address]
            if resolved not in native:
                status = 'Confirmed identity no longer observed'
            elif by_smtp.get(address, set()) - {resolved}:
                status = 'Address reassigned / identity conflict'
        elif not status:
            found = by_smtp.get(address, set()) if re.fullmatch(r'[^\s@]+@[^\s@]+', address) else set()
            if len(found) == 1:
                candidate = next(iter(found))
                status = 'Binding confirmation required'
            else:
                status = 'Ambiguous address' if found else 'Unmatched address'
        item = native.get(resolved or candidate)
        domain = '__unknown__'
        if item and item['identity'][0] == 'local_mailboxes':
            domain = key(value(item['row'], 'DomainName')) or '__unknown__'
        if not status and item:
            if item['issues']:
                status = 'Identity conflict'
            elif item['identity'][0] == 'mailboxes':
                if item['identity'] in local_targets:
                    status = 'Conflicting hosting observations'
                elif decision == 'no' and by_smtp.get(address, set()) - {resolved}:
                    status = 'Address reassigned / identity conflict'
                else:
                    status = 'Withdrawn - verified EXO' if decision == 'no' else 'Already in Exchange Online'
            elif item['identity'][0] == 'local_mailboxes' and item['target'] != item['identity']:
                status = 'Conflicting hosting observations'
            elif item['identity'][0] == 'remote_mailboxes':
                status = 'Remote recipient only'
            elif key(value(item['row'], 'RecipientType')) not in SCOPE:
                status = 'Outside planning scope'
            else:
                status = 'Matched local'
                if address and address != item['smtp']:
                    status += ' / Address changed'
        if decision not in ('yes', 'no'):
            status = 'Invalid retention decision'
        warnings = []
        for field in ('Reason', 'Decision Owner'):
            if not str(row[field]).strip():
                warnings.append(field + ' missing')
        for field in ('Decision Date', 'Review Date'):
            if str(row[field]).strip():
                try:
                    date = dt.date.fromisoformat(row[field])
                    if field == 'Review Date' and date < as_of:
                        warnings.append('Review overdue')
                except ValueError:
                    warnings.append(field + ' invalid')
        # A contradictory supplied GUID cannot be hidden by grouping on SMTP.
        group = resolved or ('candidate|' + candidate if candidate else 'input|' + digest(json.dumps(
            [address, key(row['Exchange GUID']), key(row['AD Object GUID'])]).encode()))
        pending.append(dict(row, _group=group, _resolved=resolved, _candidate=candidate,
                            _domain=domain, _decision=decision, _status=status, _warnings=warnings))
    groups = collections.defaultdict(list)
    for row in pending:
        groups[row['_group']].append(row)
    decisions = []
    for group, items in sorted(groups.items()):
        states = {r['_decision'] for r in items}
        statuses = {r['_status'] for r in items}
        status = ('Conflicting retention decisions' if len(states) > 1 else
                  'Identity conflict' if len(statuses) > 1 else items[0]['_status'])
        first = items[0]
        review = not (status.startswith('Matched local') or status == 'Withdrawn - verified EXO')
        record = dict(identity)
        for field in ('Primary SMTP Address', 'Reason', 'Decision Owner', 'Decision Date', 'Review Date', 'Source Reference'):
            record[field] = '\n'.join(dict.fromkeys(str(r[field]) for r in items))
        record.update({'Decision Key': identity['TenantKey'] + '|' + digest(group.encode()),
            'Source Rows': ','.join(str(r['Source Row']) for r in items),
            'Retain On-premises': ' / '.join(sorted(states)), 'Mailbox Identity': first['_resolved'],
            'Candidate Identity': first['_candidate'], 'Source Domain':
                'Unknown / unassigned' if first['_domain'] == '__unknown__' else first['_domain'],
            'Domain Key': first['_domain'], 'Status': status, 'Requires Review': int(review),
            'Retained': int(not review and states == {'yes'}),
            'Documentation Warnings': '; '.join(sorted({w for r in items for w in r['_warnings']})) +
                ('; Duplicate input rows' if len(items) > 1 else ''), 'Snapshot Date': as_of.isoformat()})
        decisions.append(record)
    return inventory, decisions


def summarize(identity, inventory, decisions, metadata):
    result = []
    domains = sorted(set(inventory) | {r['Domain Key'] for r in decisions})
    global_reviews = sum(r['Requires Review'] for r in decisions if r['Domain Key'] == '__unknown__')
    for domain in ['__all__'] + domains:
        selected = decisions if domain == '__all__' else [r for r in decisions if r['Domain Key'] == domain]
        n = sum(inventory.values()) if domain == '__all__' else inventory[domain]
        e = sum(r['Retained'] for r in selected)
        r = sum(row['Requires Review'] for row in selected)
        if e > n:
            raise ValueError('Retained decisions exceed the same-snapshot local inventory')
        qualified = metadata['Source Status'] == 'Qualified'
        result.append(dict(identity, **metadata, **{'Domain Key': domain, 'Source Domain': domain,
            'Inventory Count': n if qualified else '', 'Retained Count': e if qualified else '',
            'Review Count': r if qualified else '', 'Global Review Count': global_reviews if qualified else '',
            'Remaining Count': n - e if qualified and not r and not global_reviews else '',
            'Documentation Warning Count': sum(bool(row['Documentation Warnings']) for row in selected) if qualified else ''}))
    return result


def write_csv(path, columns, records):
    with path.open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=columns, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(records)


def build(data_root, output, tenant, audit_path=None, enabled=False, binding_path=None, now=None):
    from workplace_native_build import identity_from_receipts
    from workplace_native_sources import load_json, rows as csv_rows, recheck_sources, sha, utc, CONTRACT
    now = now or dt.datetime.now(dt.timezone.utc)
    data_root, output = data_root.resolve(strict=True), output.resolve()
    if output == data_root or output in data_root.parents or data_root in output.parents:
        raise ValueError('Planning staging must be outside the captured source tree')
    if any((output / name).exists() for name in ('MailboxMigrationExceptions.csv', 'MailboxMigrationExceptionSummary.csv')):
        raise ValueError('Planning staging must not overwrite existing outputs')
    identity = identity_from_receipts(data_root / 'DATA-LAST', tenant) if enabled else {
        'TenantKey': tenant, 'OrganizationKey': '', 'EnvironmentKey': '', 'TenantId': ''}
    metadata = {'Source Status': 'Not configured', 'Snapshot Date': now.date().isoformat(),
        'Workbook SHA256': '', 'Workbook Modified UTC': '', 'Captured UTC': now.isoformat(),
        'Binding SHA256': '', 'Source Rows': 0}
    decisions, inventory = [], collections.Counter()
    if enabled:
        if not audit_path:
            raise ValueError('Receipt-qualified native capture audit is required')
        audit = load_json(audit_path)
        if audit.get('Version') != 1 or audit.get('ContractSHA256') != sha(CONTRACT) or audit['Identity'] != identity:
            raise ValueError('Native capture identity mismatch')
        contract = load_json(CONTRACT)
        recheck_sources(data_root / 'DATA-LAST', audit['Evidence'], contract)
        workbook = data_root / WORKBOOK
        before = workbook.read_bytes()
        rows = read_register(workbook)
        bindings = {}
        if binding_path:
            binding_bytes = binding_path.read_bytes()
            bindings = validate_bindings(json.loads(binding_bytes.decode('utf-8-sig')), identity)
            metadata['Binding SHA256'] = digest(binding_bytes)
        specs = {s['name']: s for s in contract['sources']}
        populations = {family: list(csv_rows(data_root / 'DATA-LAST' / specs[family]['file']))
                       for family in ('local_mailboxes', 'remote_mailboxes', 'mailboxes')}
        inventory, decisions = resolve(rows, populations, identity, bindings, now.date())
        if workbook.read_bytes() != before:
            raise ValueError('Exception workbook changed during preparation')
        recheck_sources(data_root / 'DATA-LAST', audit['Evidence'], contract)
        # One qualified file acquisition interval can cross midnight or contain
        # several domain reading dates. Its receipt completion defines this snapshot.
        proof = next(r for r in audit['Evidence']['Files'] if r['File'] == specs['local_mailboxes']['file'])
        snapshot = utc(proof['CompletedAtUtc']).date().isoformat()
        for row in decisions:
            row['Snapshot Date'] = snapshot
        metadata.update({'Source Status': 'Qualified', 'Snapshot Date': snapshot,
            'Workbook SHA256': digest(before), 'Workbook Modified UTC':
                dt.datetime.fromtimestamp(workbook.stat().st_mtime, dt.timezone.utc).isoformat(), 'Source Rows': len(rows)})
    summary = summarize(identity, inventory, decisions, metadata)
    output.mkdir(parents=True, exist_ok=True)
    write_csv(output / 'MailboxMigrationExceptions.csv', DECISION_COLUMNS, decisions)
    write_csv(output / 'MailboxMigrationExceptionSummary.csv', SUMMARY_COLUMNS, summary)
    (output / 'MailboxMigrationExceptionsAudit.json.txt').write_text(json.dumps(
        dict(Version=1, Identity=identity, **metadata, DecisionGroups=len(decisions), Publication=False), indent=2), encoding='utf-8')
    return summary


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--validate-workbook', type=Path)
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-root', type=Path)
    parser.add_argument('--tenant')
    parser.add_argument('--audit', type=Path)
    parser.add_argument('--enabled', action='store_true')
    parser.add_argument('--bindings', type=Path)
    args = parser.parse_args()
    if args.validate_workbook:
        print('Validated exception workbook rows: ' + str(len(read_register(args.validate_workbook))))
        return
    if not args.data_root or not args.output_root or not args.tenant:
        parser.error('--data-root, --output-root and --tenant are required for preparation')
    build(args.data_root, args.output_root, args.tenant, args.audit, args.enabled, args.bindings)
    print('Mailbox migration exception evidence prepared; no refresh or publication.')


if __name__ == '__main__':
    main()
