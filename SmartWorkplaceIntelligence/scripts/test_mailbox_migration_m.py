"""Mailbox Inventory: non-identity M differential tests and historical source guards, isolated from Power BI Desktop.

Set PQTEST_PATH to the SDK's PQTest.exe and run:
  python -B -m unittest test_mailbox_migration_m -v
All file/folder/SharePoint access is replaced with synthetic in-memory binaries.
The before partition is frozen in fixtures/mailbox-migration-before.pq.
No production data, model connection, refresh, credentials or publication.
"""
import base64
import copy
import csv
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import textwrap
import unittest

ROOT = Path(__file__).resolve().parents[1]
DEFINITION = ROOT / 'pbip/SmartWorkplaceIntelligence.SemanticModel/definition'
PQTEST = os.environ.get('PQTEST_PATH') or shutil.which('PQTest.exe')
BASELINE = Path(__file__).parent / 'fixtures/mailbox-migration-before.pq'
TENANT = 'synthetic'
BATCH = 'synthetic-batch'
PREP = 'MailboxEvidence.csv'
MAIL = 'Source-Exchange_OnPrem_Mailboxes_AllDomains.csv'
IDENTITY = 'Source-Exchange_HybridIdentity_Issues.csv'
PERMISSION = 'Source-Exchange_Mailboxes_AllSources_PermissionsByUser.csv'
JOBS = 'Source-Exchange_EXO_MigrationJobs.csv'
REQUIRED = {
    PREP: ['TenantKey', 'Primary SMTP Address', 'Country', 'Exchange Version'],
    MAIL: ['TenantKey', 'ExchangeGuid', 'ObjectGUID', 'RecipientType', 'PrimarySMTPaddress',
           'DisplayName', 'TotalItemSize-In-MB', 'ArchiveTotalItemSize-In-MB',
           'FullAccessCount', 'SendAsCount', 'DomainName'],
    IDENTITY: ['TenantKey', 'ObjectGUID', 'Potential_Issue', 'IssueCategory'],
    PERMISSION: ['TenantKey', 'PrimarySMTPaddress', 'HasCrossPremisesPermissions'],
    JOBS: ['TenantKey', 'EmailAddress', 'MigrationUser', 'BatchName', 'BatchStatus', 'UserStatus'],
}
OPTIONAL = ['LitigationHoldEnabled', 'RetentionHoldEnabled', 'InPlaceHolds',
            'LargeItemCount-Over-35MB', 'GrantSendOnBehalfTo', 'CollectedAtUtc',
            'ArchiveStatus', 'ArchiveState', 'LargeItemCollectionStatus']


def m_text(value):
    return '"' + value.replace('#', '#(#)').replace('"', '""').replace(
        '\r', '#(cr)').replace('\n', '#(lf)') + '"'


def expression(name):
    from test_prepared_latest_m import expression as read_expression
    return read_expression(name)


def partition():
    source = (DEFINITION / 'tables/Mailbox Migration Inventory.tmdl').read_text(encoding='utf-8-sig')
    match = re.search(r'\t\tsource = ```\n(.*?)\n\t+```', source, re.S)
    if not match:
        match = re.search(r'\t\tsource =\n(.*?)(?=\n\tannotation|\Z)', source, re.S)
    if not match:
        raise AssertionError('Mailbox partition not found')
    return textwrap.dedent(match.group(1)).strip()


def fixture():
    return {
        PREP: [dict(zip(REQUIRED[PREP], [TENANT, ' person@example.invalid ', 'Sample', '2016']))],
        MAIL: [dict(zip(REQUIRED[MAIL], [TENANT, 'mail-1', 'ad-1', 'UserMailbox',
               'PERSON@example.invalid', 'Synthetic user', '51200', '1024', '1', '0', 'example.invalid']))],
        IDENTITY: [dict(zip(REQUIRED[IDENTITY], [TENANT, 'AD-1', 'Synthetic identity issue', 'Critical']))],
        PERMISSION: [dict(zip(REQUIRED[PERMISSION], [TENANT, 'person@example.invalid', 'true']))],
        JOBS: [dict(zip(REQUIRED[JOBS], [TENANT, 'person@example.invalid', '', 'Synthetic batch', 'Synced', 'Failed']))],
    }


def binaries(rows=None, delimiter=',', bom=False, newline='\r\n', optional=False):
    rows = fixture() if rows is None else rows
    files = {}
    for name, required in REQUIRED.items():
        headers = list(required)
        if name == MAIL and optional:
            headers += OPTIONAL
        # Include deliberately unused columns; projection must preserve output.
        headers += ['Unused payload']
        stream = io.StringIO(newline='')
        writer = csv.writer(stream, delimiter=delimiter, quoting=csv.QUOTE_ALL, lineterminator=newline)
        writer.writerow(headers)
        for row in rows[name]:
            writer.writerow([row.get(col, '') for col in headers])
        files[name] = stream.getvalue().encode('utf-8-sig' if bom else 'utf-8')
    return files


def batch_files(payload, mutate=None):
    manifest = {'SchemaVersion': 1, 'BatchId': BATCH, 'TenantKey': TENANT,
                'Files': [{'File': PREP, 'Bytes': len(payload[PREP])}],
                'Artifacts': [{'File': n, 'Bytes': len(b)} for n, b in payload.items() if n != PREP]}
    pointer = {'SchemaVersion': 1, 'BatchId': BATCH, 'TenantKey': TENANT}
    if mutate:
        mutate(manifest, pointer)
    files = {'batches/' + BATCH + '/' + n: b for n, b in payload.items()}
    files['batches/' + BATCH + '/batch.json.txt'] = json.dumps(manifest).encode()
    files['current.json.txt'] = json.dumps(pointer).encode()
    return files


def binary_literal(value):
    return 'Binary.FromText(' + m_text(base64.b64encode(value).decode()) + ', BinaryEncoding.Base64)'


def source_bindings(files, mode):
    rows = ['{' + m_text('C:\\synthetic\\' + n.replace('/', '\\')) + ',' + binary_literal(b) + '}'
            for n, b in files.items()]
    tree = {}
    for name, value in files.items():
        node = tree
        parts = ['Documents', 'DATA', 'DATA-POWERBI'] + name.split('/')
        for part in parts[:-1]:
            node = node.setdefault(part, {})
        node[parts[-1]] = value

    def folder(node):
        return '#table({"Name","Content"},{' + ','.join(
            '{' + m_text(n) + ',' + (folder(v) if isinstance(v, dict) else binary_literal(v)) + '}'
            for n, v in node.items()) + '})'

    return [
        'PreparedSourceMode=' + m_text(mode), 'PreparedEvidenceRoot="C:\\synthetic"',
        'PreparedBatchId=' + m_text(BATCH), 'PreparedExpectedTenantKey=' + m_text(TENANT),
        'SharePointSiteUrl="https://example.invalid"', 'SharePointLibraryName="Documents"',
        'SharePointDataFolder="DATA"', 'PreparedSharePointFolder="DATA-POWERBI"',
        'SyntheticFiles=#table({"Path","Content"},{' + ','.join(rows) + '})',
        'Folder.Contents=(path as text)=>let files=Table.SelectRows(SyntheticFiles, each '
        'Text.BeforeDelimiter([Path],"\\",{0,RelativePosition.FromEnd})=path) '
        'in Table.AddColumn(files,"Name",each Text.AfterDelimiter([Path],"\\",{0,RelativePosition.FromEnd}))',
        'File.Contents=(path as text)=>Table.SelectRows(SyntheticFiles,each [Path]=path){0}[Content]',
        'SharePoint.Contents=(url as text, options as record)=>' + folder(tree),
        'Web.Contents=(url as text, options as record)=>'
        'if options[RelativePath]="_api/web/lists" then Json.FromValue([value={[Title="Documents",RootFolder=[Name="Documents",ServerRelativeUrl="/Documents"]]}]) '
        'else if Text.EndsWith(options[RelativePath],"/RootFolder") then Json.FromValue([ServerRelativeUrl="/Documents"]) '
        'else let name=if Text.Contains(options[RelativePath],"current.json.txt") then "current.json.txt" else "current.json", '
        'found=Table.SelectRows(SyntheticFiles,each [Path]="C:\\synthetic\\" & name) '
        'in if Table.RowCount(found)=1 then found{0}[Content] else Value.ReplaceMetadata(Text.ToBinary(""),[Response.Status=404])',
    ]


def differential(files, check='true', error=False, mode='Local', overrides=None):
    bindings = source_bindings(files, mode)
    for name, replacement in (overrides or {}).items():
        bindings = [replacement if b.startswith(name + '=') else b for b in bindings]
    names = ['PreparedReadFolder', 'CreatePreparedReaders', 'NativeEvidenceTypes',
             'GetPreparedEvidenceFile', 'GetPreparedBatchManifest', 'ReadMigrationCsv',
             'CreateMailboxMigrationReader', 'ReadMailboxMigrationCsv']
    bindings += [n + '=(' + expression(n) + ')' for n in names]
    bindings += ['Before=(' + BASELINE.read_text(encoding='utf-8') + ')', 'After=(' + partition() + ')',
                 'Serialize=(t)=>let business=Table.RemoveColumns(t,{"Loaded Batch ID","Identity Issues","Migration Status","Risk Reasons","Next Action"},MissingField.Ignore) '
                 'in Binary.Buffer(Json.FromValue([Columns=Table.ColumnNames(business),Rows=Table.ToRows(business)]))',
                 'Old=try Serialize(Before)', 'New=try Serialize(After)']
    if error:
        result = 'Old[HasError] and New[HasError] and Old[Error][Message]=New[Error][Message]'
    else:
        result = 'not Old[HasError] and not New[HasError] and Old[Value]=New[Value] and (' + check + ')'
    return 'let ' + ',\n'.join(bindings) + '\nin ' + result


def cases():
    result = []

    def add(name, payload=None, check='true', error=False, mode='Local', mutate=None, files=None, overrides=None):
        files = batch_files(binaries() if payload is None else payload, mutate) if files is None else files
        result.append((name, differential(files, check, error, mode, overrides)))

    add('All output columns/values/order and source joins', check='Table.RowCount(After)=1 and '
        'After{0}[Country]="Sample" and After{0}[Migration Status]="Observed blocker" and '
        'After{0}[#"Cross-premises Permissions"]="Observed" and After{0}[Mailbox Size GB]=50')
    rows = fixture()
    rows[MAIL].append({**rows[MAIL][0], 'ExchangeGuid': 'mail-2', 'ObjectGUID': 'ad-2',
                      'PrimarySMTPaddress': 'second@example.invalid'})
    rows[PREP].insert(0, {**rows[PREP][0], 'Primary SMTP Address': 'second@example.invalid',
                          'Country': 'Second'})
    # Reversed index order and unmatched keys expose incorrect key/value pairing.
    rows[IDENTITY].insert(0, {**rows[IDENTITY][0], 'ObjectGUID': 'unmatched'})
    rows[PERMISSION].insert(0, {**rows[PERMISSION][0], 'PrimarySMTPaddress': ''})
    add('Multiple output rows, reversed index order, blank and unmatched keys', binaries(rows),
        check='Table.RowCount(After)=2 and After{0}[Country]="Sample" and '
        'After{1}[Country]="Second" and After{1}[Batch Name]="Not observed" and '
        'After{1}[#"Cross-premises Permissions"]="Unknown"')
    for delimiter, bom, newline in [(';', False, '\n'), (',', True, '\r\n'), (';', True, '\r\n')]:
        add(f'Delimiter {delimiter}, BOM {bom}, newline {repr(newline)}', binaries(delimiter=delimiter, bom=bom, newline=newline))
    rows = fixture()
    rows[MAIL][0].update(dict(zip(OPTIONAL, ['true', 'true', '[hold-1]', '2', 'delegate@example.invalid',
        '2024-06-18T09:08:07Z', 'Active', 'Fallback', 'Collected'])))
    rows[MAIL][0]['DisplayName'] = 'Synthetic "quoted"; name\nsecond line \u00e9'
    rows[MAIL][0]['Unused payload'] = 'x' * 10000
    add('All optional values, quotes, embedded newline and Unicode', binaries(rows, optional=True),
        check='After{0}[Archive State]="Active" and After{0}[Hold and Retention]="Review - preserve settings" '
        'and After{0}[Large Item Count]="2" and After{0}[Evidence Date]="2024-06-18" '
        'and After{0}[Send On Behalf]="delegate@example.invalid" and After{0}[Large Item Evidence]="Collected"')
    # Every optional field independently affects a tested output or risk decision.
    optional_checks = {
        'LitigationHoldEnabled': ('true', 'After{0}[Hold and Retention]="Review - preserve settings"'),
        'RetentionHoldEnabled': ('true', 'After{0}[Hold and Retention]="Review - preserve settings"'),
        'InPlaceHolds': ('[hold]', 'After{0}[Hold and Retention]="Review - preserve settings"'),
        'LargeItemCount-Over-35MB': ('3', 'After{0}[Large Item Count]="3"'),
        'GrantSendOnBehalfTo': ('delegate', 'After{0}[Send On Behalf]="delegate"'),
        'CollectedAtUtc': ('2024-06-18T00:00:00Z', 'After{0}[Evidence Date]="2024-06-18"'),
        'ArchiveStatus': ('Active', 'After{0}[Archive State]="Active"'),
        'ArchiveState': ('Fallback', 'After{0}[Archive State]="Fallback"'),
        'LargeItemCollectionStatus': ('NotCollected', 'After{0}[Large Item Evidence]="NotCollected"'),
    }
    for field, (value, check) in optional_checks.items():
        rows = fixture()
        rows[MAIL][0][field] = value
        add('Optional field: ' + field, binaries(rows, optional=True), check=check)
    add('Missing optional columns retain unknown/default semantics',
        check='After{0}[Archive State]="Unknown" and After{0}[Evidence Date]="Unknown" '
        'and After{0}[Large Item Count]=null')
    rows = fixture()
    rows[MAIL][0]['ExchangeGuid'] = ''
    add('Identity falls back to ObjectGUID', binaries(rows), check='After{0}[Mailbox GUID]="ad-1"')
    rows = fixture()
    rows[MAIL][0]['TotalItemSize-In-MB'] = 'not measured'
    rows[MAIL][0]['FullAccessCount'] = ''
    add('Invalid numeric evidence stays null', binaries(rows), check='After{0}[Mailbox Size GB]=null and After{0}[Full Access Count]=null')
    for name, check in [(PREP, 'After{0}[Country]="Unknown"'),
                        (PERMISSION, 'After{0}[#"Cross-premises Permissions"]="Unknown"'),
                        (JOBS, 'After{0}[Batch Name]="Multiple - review" and After{0}[Migration User Status]="Not observed"')]:
        rows = fixture()
        rows[name].append(copy.deepcopy(rows[name][0]))
        add('Ambiguous duplicate: ' + name, binaries(rows), check=check)
    rows = fixture()
    rows[IDENTITY].append({**rows[IDENTITY][0], 'Potential_Issue': 'Second issue'})
    add('Stale legacy identity findings are ignored', binaries(rows), check='After{0}[Identity Issues]="Hybrid identity evidence unavailable in selected batch"')
    rows = fixture()
    rows[JOBS][0]['EmailAddress'] = ''
    rows[JOBS][0]['MigrationUser'] = 'PERSON@example.invalid'
    add('Job identity fallback and case normalization', binaries(rows), check='After{0}[Batch Name]="Synthetic batch"')
    rows = fixture()
    for name in [IDENTITY, PERMISSION, JOBS]:
        rows[name] = []
    add('Empty optional source tables', binaries(rows), check='After{0}[#"Cross-premises Permissions"]="Unknown" and After{0}[Batch Name]="Not observed"')
    rows = fixture()
    rows[MAIL] += [{**rows[MAIL][0], 'ExchangeGuid': '', 'ObjectGUID': '', 'RecipientType': 'RemoteUserMailbox'}]
    add('Remote account excluded before identity check', binaries(rows), check='Table.RowCount(After)=1')
    rows = fixture()
    rows[MAIL] = []
    add('Header-only mailbox source', binaries(rows), check='Table.RowCount(After)=0 and '
        'List.Count(Table.ColumnNames(After))=27 and List.Contains(Table.ColumnNames(After),"Loaded Batch ID")')
    for name in (n for n in REQUIRED if n != IDENTITY):
        rows = fixture()
        rows[name][0]['TenantKey'] = 'different-synthetic-tenant'
        add('Tenant mismatch: ' + name, binaries(rows), error=True)
    rows = fixture()
    rows[PREP].append({**rows[PREP][0], 'TenantKey': 'other'})
    add('Mixed prepared tenant', binaries(rows), error=True)
    rows = fixture()
    rows[MAIL].append({**rows[MAIL][0], 'TenantKey': 'other', 'RecipientType': 'RemoteUserMailbox'})
    add('Remote rows still checked for tenant mismatch', binaries(rows), error=True)
    rows = fixture()
    rows[MAIL].append({**rows[MAIL][0], 'ExchangeGuid': ' MAIL-1 '})
    add('Duplicate normalized mailbox identity', binaries(rows), error=True)
    rows = fixture()
    rows[MAIL][0].update(ExchangeGuid='', ObjectGUID='')
    add('Missing mailbox identity', binaries(rows), error=True)
    for name, content in [('empty', b''), ('missing required', b'"TenantKey","Other"\n"synthetic","x"\n'),
                          ('duplicate unused header', b'"TenantKey","Unused","Unused"\n'),
                          ('empty header', b'"TenantKey",""\n')]:
        payload = binaries()
        payload[MAIL] = content
        add('CSV guard: ' + name, payload, error=True)
    for label, mutate in [
        ('manifest tenant', lambda m, p: m.update(TenantKey='other')),
        ('pointer tenant', lambda m, p: p.update(TenantKey='other')),
        ('manifest batch ID', lambda m, p: m.update(BatchId='wrong')),
        ('manifest schema', lambda m, p: m.update(SchemaVersion=2)),
        ('file byte length', lambda m, p: m['Files'][0].update(Bytes=1)),
        ('undeclared file', lambda m, p: m.update(Files=[])),
        ('duplicate declaration', lambda m, p: m['Files'].append(dict(m['Files'][0]))),
    ]:
        add('Batch guard: ' + label, mutate=mutate, error=True)
    files = batch_files(binaries())
    del files['batches/' + BATCH + '/' + MAIL]
    add('Missing physical file', files=files, error=True)
    for name, replacement in [('PreparedBatchId', 'PreparedBatchId=""'),
                              ('PreparedExpectedTenantKey', 'PreparedExpectedTenantKey=""'),
                              ('PreparedBatchId', 'PreparedBatchId="../escape"'),
                              ('PreparedSourceMode', 'PreparedSourceMode="Invalid"')]:
        add('Required parameter: ' + replacement, overrides={name: replacement}, error=True)
    add('Pinned batch independent of newer current pointer', mutate=lambda m, p: p.update(BatchId='newer'))
    for mode in ['Local', 'SharePoint']:
        files = batch_files(binaries())
        for name in list(files):
            if name.endswith('.json.txt'):
                files[name[:-4]] = files.pop(name)
        add('Legacy metadata absence fallback: ' + mode, files=files, mode=mode)
        files = batch_files(binaries())
        files['current.json'] = files['current.json.txt']
        files['current.json.txt'] = b'invalid json'
        add('No fallback after metadata parse error: ' + mode, files=files, mode=mode, error=True)
    add('Synthetic SharePoint navigation', mode='SharePoint')
    return result


class MailboxMigrationMTests(unittest.TestCase):
    def test_scope_and_materialization_contract(self):
        candidate = partition()
        business = candidate.split('BusinessRows = (', 1)[1].rsplit(')\nin\n', 1)[0]
        old = BASELINE.read_text(encoding='utf-8')
        # Unrelated mailbox projection stays equivalent; readiness semantics are tested separately.
        self.assertIn('Readers[GetReadinessSource]("HybridIdentityFindings.csv")', candidate)
        self.assertNotIn('Exchange_HybridIdentity_Issues.csv', candidate)
        self.assertEqual(candidate.count('CreateMailboxMigrationReader()'), 1)
        self.assertNotIn('GetPreparedEvidenceFile(', candidate)
        factory = expression('CreatePreparedReaders')
        self.assertIn('ManifestEntries = List.Buffer(', factory)
        self.assertIn('CreateMailboxMigrationReader = () as function => GetPreparedEvidenceFile', factory)
        reader = factory.split('    ReadMailboxMigrationCsv = ', 1)[1].split(',\n    ReadMigrationWeeklyHistory = ', 1)[0]
        self.assertNotIn('Text.FromBinary', reader)
        self.assertIn('List.First(Lines.FromBinary(', reader)
        self.assertIn('Table.Buffer(Projected)', reader)
        for field in OPTIONAL:
            self.assertIn('"' + field + '"', candidate[:candidate.index('    Project =')])

    @unittest.skipUnless(PQTEST, 'Set PQTEST_PATH to run actual Power Query M tests')
    def test_actual_m_before_after(self):
        tests = cases()
        definitions = ['[Name=' + m_text(name) + ',Run=()=>(' + body + ')]' for name, body in tests]
        query = 'let Cases={' + ',\n'.join(definitions) + '},\nResults=List.Transform(Cases,(c)=>'
        query += 'let a=try c[Run]() in [Name=c[Name],Passed=if a[HasError] then false else a[Value]=true,'
        query += 'Detail=if a[HasError] then a[Error][Message] else null]) in Table.FromRecords(Results)'
        with tempfile.TemporaryDirectory(prefix='mailbox-migration-m-') as scratch:
            path = Path(scratch) / 'mailbox.query.pq'
            path.write_text(query, encoding='utf-8')
            completed = subprocess.run([PQTEST, 'run-test', '--queryFile', str(path)], cwd=scratch,
                                       capture_output=True, text=True, encoding='utf-8-sig', timeout=180)
        self.assertEqual(completed.returncode, 0, completed.stdout[-5000:] + completed.stderr[-2000:])
        result = json.loads(completed.stdout)
        self.assertEqual(result[0]['Status'], 'Passed', completed.stdout[-5000:])
        self.assertEqual(result[0]['RowCount'], len(tests), completed.stdout[-5000:])
        for row in result[0]['Output']:
            with self.subTest(case=row['Name']):
                self.assertTrue(row['Passed'], row['Detail'])
        print(f'Executed {len(tests)} synthetic before/after cases through the actual M engine.')


if __name__ == '__main__':
    unittest.main()
