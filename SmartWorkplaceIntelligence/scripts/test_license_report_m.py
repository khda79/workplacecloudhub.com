"""Run the actual license report M expressions on a synthetic prepared batch.

Set PQTEST_PATH to Microsoft.PowerQuery.SdkTools/tools/PQTest.exe, then run
python -B -m unittest test_license_report_m -v from this directory.
Covers the batch qualification (Qualified, Not qualified, Not verified), the
empty-table gate and the Licences 1.49 SKU capacity header. Synthetic files
only; no Desktop connection, production source, refresh or publication.
"""
import base64
import csv
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import textwrap
import unittest


PQTEST = os.environ.get('PQTEST_PATH') or shutil.which('PQTest.exe')
DEFINITION = Path(__file__).resolve().parents[1] / 'pbip' / 'SmartWorkplaceIntelligence.SemanticModel' / 'definition'
EXPRESSIONS = DEFINITION / 'expressions.tmdl'
SNAPSHOT = 'a' * 32
RUN = 'synthetic-licensing-run'
TENANT = 'synthetic'
SUMMARY = ['SnapshotId', 'TenantKey', 'GeneratedAtUtc', 'LicenseCollectedAtUtc', 'Product', 'Category', 'EnabledUnits', 'ConsumedUnits',
           'AvailableUnits', 'Subscribed', 'AssignedUsers']
CAPACITY_149 = ['SnapshotId', 'TenantKey', 'ReceiptRunId', 'SkuId', 'SkuPartNumber', 'SkuDisplayName', 'Excluded', 'ExclusionReason',
                'EnabledUnits', 'WarningUnits', 'SuspendedUnits', 'LockedOutUnits', 'ConsumedUnits', 'AvailableUnits', 'AssignedUsers',
                'DisabledAccountUsers', 'ActiveUsersUsageNotMeasured', 'RecommendedCandidates', 'ReviewCandidates', 'BlockedCandidates',
                'ReferencedCandidates', 'CopilotGraceUsers', 'UsageEvidence', 'UsageEvidenceReason']
REPORT_FILES = ['M365_Licenses_ReportSummary.csv', 'M365_Licenses_ReportCandidates.csv', 'M365_Licenses_ReportGaps.csv',
                'M365_Licenses_ReportSources.csv', 'M365_Licenses_ReportAssignmentIssues.csv']
DECISION_FILES = ['M365_Licenses_ReportSkuCandidates.csv', 'M365_Licenses_ReportSkuCapacity.csv']


def m_text(value):
    return '"' + value.replace('#', '#(#)').replace('"', '""').replace('\r', '#(cr)').replace('\n', '#(lf)') + '"'


def expression(name):
    match = re.search(r'^expression ' + re.escape(name) + r' =\r?\n(.*?)(?=^\tlineageTag:)',
                      EXPRESSIONS.read_text(encoding='utf-8-sig'), re.M | re.S)
    if not match:
        raise AssertionError(name + ' expression not found')
    return textwrap.dedent(match.group(1)).strip()


def partition(table):
    text = (DEFINITION / 'tables' / (table + '.tmdl')).read_text(encoding='utf-8-sig')
    match = re.search(r'\t\tsource =\r?\n(.*?)(?=\r?\n\r?\n\tannotation)', text, re.S)
    if not match:
        raise AssertionError(table + ' partition not found')
    return textwrap.dedent(match.group(1)).strip()


def write_csv(path, header, rows):
    with open(path, 'w', encoding='utf-8', newline='') as handle:
        writer = csv.writer(handle, quoting=csv.QUOTE_ALL)
        writer.writerow(header)
        writer.writerows(rows)


def build_batch(root, batch_id, qualification, decision_files=True):
    folder = root / 'batches' / batch_id
    folder.mkdir(parents=True)
    write_csv(folder / 'Source-M365_Licenses_ReportSummary.csv', SUMMARY,
              [[SNAPSHOT, TENANT, '', '', 'Microsoft 365 E3', 'Suite', '10', '8', '2', 'True', '8'] for _ in range(7)])
    for name in REPORT_FILES[1:]:
        write_csv(folder / ('Source-' + name), ['SnapshotId', 'TenantKey'], [])
    snapshot = {'SnapshotId': SNAPSHOT, 'TenantKey': TENANT, 'SchemaVersion': 2, 'Products': [{'Product': 'Microsoft 365 E3'}] * 4,
                'OtherProducts': [{}, {}, {}], 'AssignmentIssues': {'Available': False, 'Rows': []}}
    (folder / 'Source-M365_Licenses_ReportSnapshot.json.txt').write_text(json.dumps(snapshot), encoding='utf-8')
    artifacts = ['Source-' + n for n in REPORT_FILES] + ['Source-M365_Licenses_ReportSnapshot.json.txt']
    if decision_files:
        # Licences 1.49 header: no ReportProduct column.
        write_csv(folder / 'Source-M365_Licenses_ReportSkuCapacity.csv', CAPACITY_149,
                  [[SNAPSHOT, TENANT, RUN, 'sku-1', 'SPE_E3', 'Microsoft 365 E3', 'False', '', '10', '0', '0', '0', '8', '2', '8', '1', '0', '1', '0', '0', '1', '0', 'Not applicable', 'x']])
        write_csv(folder / 'Source-M365_Licenses_ReportSkuCandidates.csv', ['SnapshotId', 'TenantKey'], [])
        artifacts += ['Source-' + n for n in DECISION_FILES]
    entries = [{'File': name, 'Bytes': (folder / name).stat().st_size, 'SHA256': '0' * 64} for name in artifacts]
    manifest = {'SchemaVersion': 1, 'BatchId': batch_id, 'TenantKey': TENANT, 'Files': [], 'Artifacts': entries}
    if qualification is not None:
        manifest['LicenseReportQualification'] = qualification
    (folder / 'batch.json.txt').write_text(json.dumps(manifest), encoding='utf-8')
    (root / 'current.json.txt').write_text(json.dumps({'SchemaVersion': 1, 'BatchId': batch_id, 'TenantKey': TENANT}), encoding='utf-8')


def qualification(state, reason='', capacity_state='Qualified'):
    files = [{'File': n, 'Kind': 'Report', 'State': 'Qualified' if state == 'Qualified' else 'Not qualified', 'Reason': reason} for n in REPORT_FILES]
    files += [{'File': 'M365_Licenses_ReportSkuCandidates.csv', 'Kind': 'Decision', 'State': 'Qualified' if state == 'Qualified' else 'Not qualified', 'Reason': ''},
              {'File': 'M365_Licenses_ReportSkuCapacity.csv', 'Kind': 'Decision', 'State': capacity_state if state == 'Qualified' else 'Not qualified', 'Reason': ''}]
    return {'Version': 1, 'State': state, 'Reason': reason, 'SnapshotId': SNAPSHOT, 'ReceiptRunId': RUN, 'Files': files}


@unittest.skipUnless(PQTEST, 'Set PQTEST_PATH to run actual Power Query M regression tests')
class LicenseReportMTests(unittest.TestCase):
    def evaluate(self, root, batch_id, body):
        names = ['GetPreparedEvidenceFile', 'GetPreparedBatchManifest', 'GetLicenseReportQualification', 'LicenseReportFileUsable',
                 'GetLicenseReportSnapshot', 'GetLicenseReportCsv']
        bindings = [
            'PreparedSourceMode = "Local"', 'PreparedEvidenceRoot = ' + m_text(str(root)), 'PreparedBatchId = ' + m_text(batch_id),
            'PreparedExpectedTenantKey = ' + m_text(TENANT), 'SharePointSiteUrl = "https://example.invalid"', 'SharePointLibraryName = "Documents"',
            'SharePointDataFolder = "DATA"', 'PreparedSharePointFolder = "DATA-POWERBI"']
        # In-memory synthetic folder: the real expressions run unchanged, with no file or credential access.
        rows = []
        for path in sorted(Path(root).rglob('*')):
            if path.is_file():
                rows.append('{' + m_text(str(path)) + ', Binary.FromText(' + m_text(base64.b64encode(path.read_bytes()).decode()) + ', BinaryEncoding.Base64)}')
        bindings += ['SyntheticFiles = #table({"Path", "Content"}, {' + ', '.join(rows) + '})',
                     'Folder.Contents = (path as text) => let parent = Text.TrimEnd(path, {"\\", "/"}), files = Table.SelectRows(SyntheticFiles, each Text.BeforeDelimiter([Path], "\\", {0, RelativePosition.FromEnd}) = parent) in Table.AddColumn(files, "Name", each Text.AfterDelimiter([Path], "\\", {0, RelativePosition.FromEnd}))',
                     'File.Contents = (path as text) => Table.SelectRows(SyntheticFiles, each [Path] = path){0}[Content]']
        bindings += [name + ' = ' + expression(name) for name in names]
        bindings += ['SkuCapacity = ' + partition('License SKU Capacity'), 'Qualification = ' + partition('License Report Qualification')]
        query = 'let\n' + ',\n'.join(bindings) + ',\nResult = ' + body + '\nin\nResult'
        with tempfile.TemporaryDirectory() as scratch:
            query_file = Path(scratch) / 'license.query.pq'
            query_file.write_text(query, encoding='utf-8')
            completed = subprocess.run([PQTEST, 'run-test', '--queryFile', str(query_file)], capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=600)
        self.assertEqual(completed.returncode, 0, completed.stdout[-2000:] + completed.stderr[-2000:])
        return completed.stdout

    def run_case(self, qualification_value, body, decision_files=True):
        with tempfile.TemporaryDirectory() as data:
            root = Path(data) / 'DATA-POWERBI'
            batch_id = '20261018T000000000Z-0000abcd'
            build_batch(root, batch_id, qualification_value, decision_files)
            return self.evaluate(root, batch_id, body)

    def assert_result(self, output, expected):
        # PQTest prints the evaluated record as [{"Name": ..., "Value": ...}]; compare one field.
        name, value = expected.split(':', 1)
        values = {match.group(1): match.group(2) for match in re.finditer(r'\{"Name":"([^"]+)","Value":("[^"]*"|true|false|null|-?[0-9.]+)\}', output)}
        self.assertIn('"Status":"Passed"', output.replace(' ', ''), output[-3000:])
        self.assertEqual(values.get(name.strip('"')), value, output[-3000:])

    def test_qualified_report_reads_rows_and_legacy_capacity_header(self):
        output = self.run_case(qualification('Qualified'), '''[
            Capacity = Table.RowCount(SkuCapacity),
            ReportProduct = SkuCapacity{0}[ReportProduct] = null,
            Class = SkuCapacity{0}[Capacity Class],
            State = GetLicenseReportQualification()[State],
            Rows = Table.RowCount(Qualification)
        ]''')
        self.assert_result(output, '"Capacity":1')
        self.assert_result(output, '"ReportProduct":true')
        self.assert_result(output, '"Class":"Governed"')
        self.assert_result(output, '"State":"Qualified"')
        self.assert_result(output, '"Rows":8')

    def test_not_qualified_report_is_empty_and_never_read(self):
        output = self.run_case(qualification('Not qualified', 'Report snapshot belongs to another licensing receipt run.'), '''[
            Capacity = Table.RowCount(SkuCapacity),
            Columns = List.Count(Table.ColumnNames(SkuCapacity)),
            Usable = LicenseReportFileUsable("M365_Licenses_ReportSummary.csv"),
            Reason = GetLicenseReportQualification()[Reason]
        ]''')
        self.assert_result(output, '"Capacity":0')
        self.assert_result(output, '"Usable":false')
        self.assert_result(output, '"Reason":"Report snapshot belongs to another licensing receipt run."')

    def test_decision_file_not_qualified_alone(self):
        output = self.run_case(qualification('Qualified', capacity_state='Not qualified'), '''[
            Capacity = Table.RowCount(SkuCapacity),
            Summary = LicenseReportFileUsable("M365_Licenses_ReportSummary.csv")
        ]''')
        self.assert_result(output, '"Capacity":0')
        self.assert_result(output, '"Summary":true')

    def test_batch_before_qualification_is_not_verified(self):
        output = self.run_case(None, '''[
            State = GetLicenseReportQualification()[State],
            Summary = LicenseReportFileUsable("M365_Licenses_ReportSummary.csv"),
            Capacity = Table.RowCount(SkuCapacity),
            CapacityState = Table.SelectRows(Qualification, each [Scope] = "M365_Licenses_ReportSkuCapacity.csv"){0}[State]
        ]''', decision_files=False)
        self.assert_result(output, '"State":"Not verified"')
        self.assert_result(output, '"Summary":true')
        self.assert_result(output, '"Capacity":0')
        self.assert_result(output, '"CapacityState":"Absent"')

    def test_invalid_qualification_fails_closed(self):
        broken = qualification('Qualified')
        broken['Version'] = 2
        output = self.run_case(broken, '(try GetLicenseReportQualification())[HasError]')
        self.assertIn('true', output.lower())


if __name__ == '__main__':
    unittest.main()
