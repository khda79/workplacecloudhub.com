"""Native Source Health status mapping: actual M on synthetic statuses plus offline model/report contract.

Set PQTEST_PATH to Microsoft.PowerQuery.SdkTools/tools/PQTest.exe to run the M cases, then run
python -B -m unittest test_native_source_health -v from this directory.
No Desktop connection, production source, refresh, or publication is used.
"""
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
PBIP = Path(__file__).resolve().parents[1] / 'pbip'
DEFINITION = Path(os.environ.get('NATIVE_EVIDENCE_EXPRESSIONS') or (
    PBIP / 'SmartWorkplaceIntelligence.SemanticModel' / 'definition' / 'expressions.tmdl')).parent
TABLE = DEFINITION / 'tables' / 'Native Source Health.tmdl'
KPI = (PBIP / 'SmartWorkplaceIntelligence.Report' / 'definition' / 'pages' / '6665fb035f086688c22f' /
       'visuals' / '995cc6913e44c6b9e887' / 'visual.json')
STATES = {'OK', 'Warning', 'Failed', 'Unknown'}
# Status -> expected state. Receipt-level and near-miss spellings must never become OK.
CASES = [('Success', 'OK'), ('SuccessWithWarnings', 'Warning'), ('Failed', 'Failed'), ('Cancelled', 'Failed'),
         ('Completed', 'Unknown'), ('CompletedWithWarnings', 'Unknown'), ('success', 'Unknown'),
         (' Success', 'Unknown'), ('Success ', 'Unknown'), ('Partial', 'Unknown'), ('', 'Unknown'), (None, 'Unknown')]


def m_text(value):
    return '"' + value.replace('#', '#(#)').replace('"', '""') + '"'


def state_expression():
    match = re.search(r'^expression NativeSourceHealthState =\r?\n(.*?)(?=^\tlineageTag:)',
                      (DEFINITION / 'expressions.tmdl').read_text(encoding='utf-8-sig'), re.M | re.S)
    if not match:
        raise AssertionError('NativeSourceHealthState expression not found')
    return textwrap.dedent(match.group(1)).strip()


def mapping():
    return dict(re.findall(r'(\w+) = "(\w+)"', re.search(r'States = \[(.*?)\]', state_expression())[1]))


def measures():
    source = TABLE.read_text(encoding='utf-8-sig')
    return dict(re.findall(r"^\tmeasure '([^']+)' = (.*)$", source, re.M)), source


class NativeSourceHealthModelTests(unittest.TestCase):
    def test_mapping_is_explicit_and_closed(self):
        self.assertEqual(mapping(), {'Success': 'OK', 'SuccessWithWarnings': 'Warning',
                                     'Failed': 'Failed', 'Cancelled': 'Failed'})
        self.assertIn('"Unknown"', state_expression())

    def test_measures_use_health_state_not_a_hard_coded_status(self):
        defined, source = measures()
        self.assertNotRegex(source, r"\[Status\]\s*==?\s*\"")
        self.assertEqual(defined['Native Collected source feeds'].strip(),
                         '([Native OK source feeds] + [Native Warning source feeds])')
        for name, state in (('Native OK source feeds', 'OK'), ('Native Warning source feeds', 'Warning'),
                            ('Native Failed source feeds', 'Failed'), ('Native Unknown-status source feeds', 'Unknown')):
            self.assertIn("KEEPFILTERS('Native Source Health'[Health State] == \"" + state + '")', defined[name])
        used = set(re.findall(r"\[Health State\] == \"(\w+)\"", source))
        self.assertEqual(used, STATES)
        self.assertIn('[Native Source feeds] - [Native Collected source feeds]', defined['Native Missing source feeds'])

    def test_health_state_column_comes_from_the_shared_function(self):
        _, source = measures()
        column = re.search(r"\tcolumn 'Health State'\r?\n(.*?)(?=\r?\n\r?\n)", source, re.S)
        self.assertIsNotNone(column)
        self.assertIn('sourceColumn: Health State', column[1])
        self.assertIn('dataType: string', column[1])
        self.assertIn('"Health State", each NativeSourceHealthState([Status]), type text', source)
        self.assertIn('Readers[Finish](BusinessRows)', source)

    def test_kpi_table_shows_every_health_state(self):
        visual = json.loads(KPI.read_text(encoding='utf-8-sig'))
        bound = [p['field']['Measure']['Property'] for p in visual['visual']['query']['queryState']['Values']['projections']]
        self.assertEqual(bound, ['Native Source feeds', 'Native Collected source feeds', 'Native OK source feeds',
                                 'Native Warning source feeds', 'Native Failed source feeds',
                                 'Native Unknown-status source feeds', 'Native Missing source feeds',
                                 'Native Source collection coverage'])
        defined, _ = measures()
        self.assertTrue(set(bound) <= set(defined))


class NativeSourceHealthMTests(unittest.TestCase):
    @unittest.skipUnless(PQTEST, 'Set PQTEST_PATH to run actual Power Query M regression tests')
    def test_actual_m_mapping(self):
        rows = ',\n'.join('[Name=' + m_text(repr(status)) + ', Status=' + ('null' if status is None else m_text(status))
                          + ', Expected=' + m_text(expected) + ']' for status, expected in CASES)
        query = ('let NativeSourceHealthState = (' + state_expression() + '),\n'
                 'Cases={' + rows + '},\n'
                 'Results=List.Transform(Cases,(Case)=>let Attempt=try NativeSourceHealthState(Case[Status]) in '
                 '[Name=Case[Name],Passed=not Attempt[HasError] and Attempt[Value]=Case[Expected],'
                 'Detail=if Attempt[HasError] then Attempt[Error][Message] else Attempt[Value]])\n'
                 'in Table.FromRecords(Results)')
        with tempfile.TemporaryDirectory(prefix='native-source-health-m-') as scratch:
            query_file = Path(scratch) / 'synthetic.query.pq'
            query_file.write_text(query, encoding='utf-8')
            completed = subprocess.run(
                [PQTEST, 'run-test', '--queryFile', str(query_file)],
                capture_output=True, text=True, encoding='utf-8-sig', timeout=90, cwd=scratch)
        self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)
        results = json.loads(completed.stdout)
        self.assertEqual(results[0]['Status'], 'Passed', completed.stdout)
        self.assertEqual(results[0]['RowCount'], len(CASES), completed.stdout)
        for result in results[0]['Output']:
            with self.subTest(case=result['Name']):
                self.assertTrue(result['Passed'], result['Detail'])
        print(f'Executed {len(CASES)} synthetic status cases through the actual M engine.')


if __name__ == '__main__':
    unittest.main()
