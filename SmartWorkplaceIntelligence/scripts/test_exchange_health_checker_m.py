"""Execute the Health Checker table partitions against synthetic binaries only.

The partition M is read from the checkout TMDL, so the test runs the shipped
expression. No model, Desktop connection or production data is used.
"""
import json
from pathlib import Path
import subprocess
import tempfile
import textwrap
import unittest

import test_exchange_infrastructure_m as infra
from test_prepared_latest_m import DEFINITION

SUMMARY = 'Source-Exchange_OnPrem_HealthChecker_Summary.csv'
FINDINGS = 'Source-Exchange_OnPrem_HealthChecker_Findings.csv'
SUMMARY_COLUMNS = ['TenantKey', 'ServerName', 'RunStatus', 'RunMessage', 'DurationSeconds', 'HealthCheckerVersion',
                   'GenerationTime', 'ExchangeVersion', 'BuildNumber', 'CumulativeUpdate', 'SecurityUpdate',
                   'LatestInstallTime', 'SecurityUpdateStatus', 'CumulativeUpdateOutOfDate', 'CumulativeUpdateDaysOld',
                   'EndOfLifeStatus', 'VulnerabilityCount', 'Vulnerabilities', 'ErrorCount', 'WarningCount', 'OkCount',
                   'InfoCount', 'ServersProcessed', 'ServersTotal']
FINDING_COLUMNS = ['TenantKey', 'ServerName', 'Category', 'CategoryOrder', 'RowNumber', 'Check', 'Item', 'Status',
                   'Value', 'ValueMasked', 'Recommendation', 'HealthCheckerVersion', 'GenerationTime']


def partition(table):
    text = (DEFINITION / 'tables' / (table + '.tmdl')).read_text(encoding='utf-8-sig').replace('\r\n', '\n')
    lines = text.split('\n')
    start = next(i for i, line in enumerate(lines) if line.strip() == 'source =') + 1
    body = []
    for line in lines[start:]:
        if line and not line.startswith('\t\t\t'):
            break
        body.append(line)
    return textwrap.dedent('\n'.join(body)).strip()


def summary_row(server, status, vulnerabilities, errors, warnings, su='MissingLatestSU', days='400'):
    values = {'TenantKey': 'synthetic', 'ServerName': server, 'RunStatus': status, 'RunMessage': '' if status == 'Completed' else 'stopped',
              'DurationSeconds': '42', 'HealthCheckerVersion': '26.10.13.1644', 'GenerationTime': '10/21/2026 06:00:00',
              'ExchangeVersion': 'Exchange 2016 CU23 Nov24SUv2' if status == 'Completed' else '', 'BuildNumber': '15.1.2507.44' if status == 'Completed' else '',
              'CumulativeUpdate': 'CU23' if status == 'Completed' else '', 'SecurityUpdate': 'Nov24SUv2' if status == 'Completed' else '',
              'LatestInstallTime': '', 'SecurityUpdateStatus': su if status == 'Completed' else 'Unknown',
              'CumulativeUpdateOutOfDate': 'True' if status == 'Completed' else '', 'CumulativeUpdateDaysOld': days if status == 'Completed' else '',
              'EndOfLifeStatus': 'Supported', 'VulnerabilityCount': vulnerabilities, 'Vulnerabilities': 'CVE-2024-21410' if vulnerabilities == '1' else '',
              'ErrorCount': errors, 'WarningCount': warnings, 'OkCount': '10' if status == 'Completed' else '', 'InfoCount': '20' if status == 'Completed' else '',
              'ServersProcessed': '1', 'ServersTotal': '2'}
    return [values[c] for c in SUMMARY_COLUMNS]


def finding_row(server, row, status, value, masked='False', tenant='synthetic'):
    values = {'TenantKey': tenant, 'ServerName': server, 'Category': 'Security Settings', 'CategoryOrder': '4', 'RowNumber': str(row),
              'Check': 'Check ' + str(row), 'Item': '', 'Status': status, 'Value': value, 'ValueMasked': masked,
              'Recommendation': 'https://aka.ms/HC-Example' if status == 'Error' else '', 'HealthCheckerVersion': '26.10.13.1644',
              'GenerationTime': '10/21/2026 06:00:00'}
    return [values[c] for c in FINDING_COLUMNS]


def fixtures():
    summary = infra.csv_bytes(SUMMARY_COLUMNS, [summary_row('mail-two', 'Timeout', '', '', ''),
                                                summary_row('mail-one', 'Completed', '1', '3', '2')])
    findings = infra.csv_bytes(FINDING_COLUMNS, [finding_row('mail-one', 2, 'Error', 'Out of date; More Information: https://aka.ms/HC-Example'),
                                                 finding_row('mail-one', 1, 'Info', '[masked]', 'True')])
    files = {SUMMARY: summary, FINDINGS: findings}
    servers = partition('Exchange Health Checker Servers')
    found = partition('Exchange Health Checker Findings')
    tests = []

    def add(name, source, check='true', error=False, **options):
        bindings = infra.fixture(source, **options)
        bindings.append('Servers = try (' + servers + ')')
        bindings.append('Findings = try (' + found + ')')
        if error:
            result = 'Servers[HasError] and Findings[HasError]'
        else:
            result = 'not Servers[HasError] and not Findings[HasError] and (let S = Servers[Value], F = Findings[Value] in ' + check + ')'
        tests.append((name, 'let\n' + ',\n'.join(bindings) + '\nin ' + result))

    add('Servers and findings are rebuilt from the selected batch', files, '''let
        One = Table.SelectRows(S, each [Server Name] = "mail-one"){0},
        Two = Table.SelectRows(S, each [Server Name] = "mail-two"){0}
        in Table.RowCount(S) = 2 and One[Run Status] = "Completed" and One[Vulnerability Count] = 1 and
           One[Error Count] = 3 and One[Warning Count] = 2 and One[Cumulative Update Days Old] = 400 and
           One[Security Update Status] = "MissingLatestSU" and One[Cumulative Update] = "CU23" and One[Is Server] and
           Two[Run Status] = "Timeout" and Two[Vulnerability Count] = null and Two[Error Count] = null and
           Value.Is(One[Vulnerability Count], type number) and List.Distinct(S[Loaded Batch ID]) = {"batch-a"} and
           Table.RowCount(F) = 2 and F{0}[Row Number] = 1 and F{0}[Value Masked] = true and F{0}[Value] = "[masked]" and
           F{1}[Status] = "Error" and F{1}[Value] = "Out of date; More Information: https://aka.ms/HC-Example" and
           F{1}[Recommendation] = "https://aka.ms/HC-Example" and F{1}[Value Masked] = false and
           List.Distinct(F[Loaded Batch ID]) = {"batch-a"}''')
    add('Health Checker collection status in the infrastructure table', files, '''let
        R = Table.SelectRows(Infrastructure, each [Dataset] = "Health Checker summary"),
        One = List.Distinct(Table.SelectRows(R, each [Server Name] = "mail-one")[Collection Status]),
        Two = List.Distinct(Table.SelectRows(R, each [Server Name] = "mail-two")[Collection Status]),
        Found = List.Distinct(Table.SelectRows(Infrastructure, each [Dataset] = "Health Checker findings")[Collection Status])
        in One = {"OK"} and Two = {"WARNING"} and Found = {"OK"}''')
    add('Not captured is an explicit unknown marker, never zero findings', {}, '''
        Table.RowCount(S) = 1 and S{0}[Is Server] = false and S{0}[Source Available] = false and
        S{0}[Run Status] = "Not captured in selected batch" and S{0}[Vulnerability Count] = null and Table.RowCount(F) = 0''')
    add('Captured without rows is distinct from absence',
        {SUMMARY: infra.csv_bytes(SUMMARY_COLUMNS, []), FINDINGS: infra.csv_bytes(FINDING_COLUMNS, [])}, '''
        Table.RowCount(S) = 1 and S{0}[Source Available] = true and S{0}[Run Status] = "Captured - no rows" and Table.RowCount(F) = 0''')
    bad = dict(files)
    bad[FINDINGS] = infra.csv_bytes(FINDING_COLUMNS, [finding_row('mail-one', 1, 'Info', 'x', tenant='foreign')])
    add('Foreign findings tenant is rejected', bad, error=True)
    bad = dict(files)
    bad[SUMMARY] = infra.csv_bytes([c for c in SUMMARY_COLUMNS if c != 'RunStatus'], [[v for c, v in zip(SUMMARY_COLUMNS, summary_row('mail-one', 'Completed', '1', '3', '2')) if c != 'RunStatus']])
    add('Missing required summary column is rejected', bad, error=True)
    add('Changed pointer is rejected', files, error=True, change='verify-infrastructure')
    return tests


@unittest.skipUnless(infra.PQTEST, 'Set PQTEST_PATH for actual M execution')
class HealthCheckerReaderTests(unittest.TestCase):
    def test_actual_m(self):
        tests = fixtures()
        cases = ','.join('[Name=' + infra.m_text(n) + ',Run=()=>(' + q + ')]' for n, q in tests)
        query = 'let C={' + cases + '},R=List.Transform(C,(c)=>let r=try c[Run]() in [Name=c[Name],Passed=if r[HasError] then false else r[Value]=true,Detail=if r[HasError] then r[Error][Message] else null]) in Table.FromRecords(R)'
        with tempfile.TemporaryDirectory(prefix='exchange-health-checker-test-') as directory:
            path = Path(directory) / 'healthchecker.query.pq'
            path.write_text(query, encoding='utf-8')
            run = subprocess.run([infra.PQTEST, 'run-test', '--queryFile', str(path)], cwd=directory, capture_output=True,
                                 text=True, encoding='utf-8-sig', timeout=300)
        self.assertEqual(run.returncode, 0, run.stderr)
        result = json.loads(run.stdout)[0]
        self.assertEqual(result['Status'], 'Passed', run.stdout[-1500:])
        for case in result['Output']:
            with self.subTest(name=case['Name']):
                self.assertTrue(case['Passed'], case['Detail'])
        print('Executed ' + str(len(tests)) + ' synthetic Exchange Health Checker M cases.')


if __name__ == '__main__':
    unittest.main()
