"""Execute the Exchange hybrid free/busy table partitions against synthetic binaries only.

The partition M and the shared reader are read from the checkout TMDL, so the test runs the
shipped expressions. No model, Desktop connection or production data is used.
"""
import json
from pathlib import Path
import subprocess
import tempfile
import textwrap
import unittest

import test_exchange_infrastructure_m as infra
from test_prepared_latest_m import DEFINITION

SUMMARY = 'Source-Exchange_Hybrid_FreeBusy_Summary.csv'
FINDINGS = 'Source-Exchange_Hybrid_FreeBusy_Findings.csv'
SUMMARY_COLUMNS = ['TenantKey', 'ServerName', 'Side', 'Auth', 'RunStatus', 'SideStatus', 'RunMessage', 'NotAuthorizedCmdlets',
                   'MissingRoles', 'DurationSeconds', 'FreeBusyCheckerVersion', 'GenerationTime', 'OkCount', 'WarningCount',
                   'ErrorCount', 'InfoCount', 'NotAuthorizedCount']
FINDING_COLUMNS = ['TenantKey', 'ServerName', 'Side', 'Auth', 'Category', 'RowNumber', 'Item', 'Status', 'Value', 'ValueMasked',
                   'FreeBusyCheckerVersion', 'GenerationTime']


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


def summary_row(side, auth, side_status, ok, warn, err, info, not_authorized, cmdlets='', roles='', tenant='synthetic'):
    values = {'TenantKey': tenant, 'ServerName': 'mail-one', 'Side': side, 'Auth': auth, 'RunStatus': 'EXONotAuthorized',
              'SideStatus': side_status, 'RunMessage': '', 'NotAuthorizedCmdlets': cmdlets, 'MissingRoles': roles,
              'DurationSeconds': '95', 'FreeBusyCheckerVersion': '26.06.06.1238', 'GenerationTime': '10/25/2026 05:00:00',
              'OkCount': ok, 'WarningCount': warn, 'ErrorCount': err, 'InfoCount': info, 'NotAuthorizedCount': not_authorized}
    return [values[c] for c in SUMMARY_COLUMNS]


def finding_row(side, auth, row, status, value, masked='False', tenant='synthetic'):
    values = {'TenantKey': tenant, 'ServerName': 'mail-one', 'Side': side, 'Auth': auth, 'Category': 'Summary - Get-OrganizationRelationship',
              'RowNumber': str(row), 'Item': 'Item ' + str(row), 'Status': status, 'Value': value, 'ValueMasked': masked,
              'FreeBusyCheckerVersion': '26.06.06.1238', 'GenerationTime': '10/25/2026 05:00:00'}
    return [values[c] for c in FINDING_COLUMNS]


def fixtures():
    summary = infra.csv_bytes(SUMMARY_COLUMNS, [
        summary_row('ExchangeOnPremises', 'DAuth', 'Completed', '9', '1', '0', '3', '0'),
        summary_row('ExchangeOnPremises', 'OAuth', 'Completed', '12', '0', '1', '2', '0'),
        summary_row('ExchangeOnline', 'DAuth', 'Completed', '6', '0', '1', '1', '0'),
        summary_row('ExchangeOnline', 'OAuth', 'NotAuthorizedPartial', '4', '0', '0', '1', '1', 'Test-EOOAuthConnectivity', 'Mail Recipients; Organization Transport Settings')])
    findings = infra.csv_bytes(FINDING_COLUMNS, [
        finding_row('ExchangeOnline', 'OAuth', 2, 'NotAuthorized', 'Not authorized in app-only mode'),
        finding_row('ExchangeOnPremises', 'DAuth', 1, 'Warning', '[masked]', 'True')])
    files = {SUMMARY: summary, FINDINGS: findings}
    summary_m = partition('Exchange Hybrid Free Busy Summary')
    findings_m = partition('Exchange Hybrid Free Busy Findings')
    tests = []

    def add(name, source, check='true', error=False, summary_only=False, **options):
        bindings = infra.fixture(source, **options)
        bindings.append('Summary = try (' + summary_m + ')')
        bindings.append('Findings = try (' + findings_m + ')')
        if error:
            result = 'Summary[HasError]' if summary_only else 'Summary[HasError] and Findings[HasError]'
        else:
            result = 'not Summary[HasError] and not Findings[HasError] and (let S = Summary[Value], F = Findings[Value] in ' + check + ')'
        tests.append((name, 'let\n' + ',\n'.join(bindings) + '\nin ' + result))

    add('Summary and findings are rebuilt from the selected batch', files, '''let
        Oauth = Table.SelectRows(S, each [Side] = "ExchangeOnline" and [Auth] = "OAuth"){0},
        OnPrem = Table.SelectRows(S, each [Side] = "ExchangeOnPremises" and [Auth] = "OAuth"){0}
        in Table.RowCount(S) = 4 and List.AllTrue(S[Is Result]) and List.AllTrue(S[Source Available]) and
           Oauth[Side Status] = "NotAuthorizedPartial" and Oauth[Not Authorized Count] = 1 and
           Oauth[Not Authorized Cmdlets] = "Test-EOOAuthConnectivity" and Text.Contains(Oauth[Missing Roles], "Mail Recipients") and
           OnPrem[Error Count] = 1 and OnPrem[OK Count] = 12 and Value.Is(OnPrem[OK Count], type number) and
           List.Distinct(S[Loaded Batch ID]) = {"batch-a"} and
           Table.RowCount(F) = 2 and F{0}[Side] = "ExchangeOnPremises" and F{0}[Value] = "[masked]" and F{0}[Value Masked] = true and
           F{1}[Side] = "ExchangeOnline" and F{1}[Status] = "NotAuthorized" and F{1}[Value Masked] = false and
           List.Distinct(F[Loaded Batch ID]) = {"batch-a"}''')
    add('Not captured is an explicit unknown marker, never zero findings', {}, '''
        Table.RowCount(S) = 1 and S{0}[Is Result] = false and S{0}[Source Available] = false and
        S{0}[Run Status] = "Not captured in selected batch" and S{0}[Error Count] = null and Table.RowCount(F) = 0''')
    add('Captured without rows is distinct from absence',
        {SUMMARY: infra.csv_bytes(SUMMARY_COLUMNS, []), FINDINGS: infra.csv_bytes(FINDING_COLUMNS, [])}, '''
        Table.RowCount(S) = 1 and S{0}[Source Available] = true and S{0}[Run Status] = "Captured - no rows" and Table.RowCount(F) = 0''')
    bad = dict(files)
    bad[SUMMARY] = infra.csv_bytes(SUMMARY_COLUMNS, [summary_row('ExchangeOnline', 'DAuth', 'Completed', '1', '0', '0', '0', '0', tenant='foreign')])
    bad[FINDINGS] = infra.csv_bytes(FINDING_COLUMNS, [finding_row('ExchangeOnline', 'DAuth', 1, 'OK', 'x', tenant='foreign')])
    add('Foreign tenant is rejected', bad, error=True)
    bad = dict(files)
    bad[SUMMARY] = infra.csv_bytes([c for c in SUMMARY_COLUMNS if c != 'SideStatus'], [[v for c, v in zip(SUMMARY_COLUMNS, summary_row('ExchangeOnline', 'DAuth', 'Completed', '1', '0', '0', '0', '0')) if c != 'SideStatus']])
    bad[FINDINGS] = infra.csv_bytes([c for c in FINDING_COLUMNS if c != 'Status'], [[v for c, v in zip(FINDING_COLUMNS, finding_row('ExchangeOnline', 'DAuth', 1, 'OK', 'x')) if c != 'Status']])
    add('Missing required columns are rejected', bad, error=True)
    add('Changed pointer is rejected', files, error=True, change='verify-freebusy')
    # The fixture duplicates the first declared artifact (the summary).
    add('Duplicate declaration is rejected', files, error=True, summary_only=True, duplicate=True)
    return tests


@unittest.skipUnless(infra.PQTEST, 'Set PQTEST_PATH for actual M execution')
class HybridFreeBusyReaderTests(unittest.TestCase):
    def test_actual_m(self):
        tests = fixtures()
        cases = ','.join('[Name=' + infra.m_text(n) + ',Run=()=>(' + q + ')]' for n, q in tests)
        query = 'let C={' + cases + '},R=List.Transform(C,(c)=>let r=try c[Run]() in [Name=c[Name],Passed=if r[HasError] then false else r[Value]=true,Detail=if r[HasError] then r[Error][Message] else null]) in Table.FromRecords(R)'
        with tempfile.TemporaryDirectory(prefix='exchange-hybrid-freebusy-test-') as directory:
            path = Path(directory) / 'freebusy.query.pq'
            path.write_text(query, encoding='utf-8')
            run = subprocess.run([infra.PQTEST, 'run-test', '--queryFile', str(path)], cwd=directory, capture_output=True,
                                 text=True, encoding='utf-8-sig', timeout=300)
        self.assertEqual(run.returncode, 0, run.stderr)
        result = json.loads(run.stdout)[0]
        self.assertEqual(result['Status'], 'Passed', run.stdout[-1500:])
        for case in result['Output']:
            with self.subTest(name=case['Name']):
                self.assertTrue(case['Passed'], case['Detail'])
        print('Executed ' + str(len(tests)) + ' synthetic Exchange hybrid free/busy M cases.')


if __name__ == '__main__':
    unittest.main()
