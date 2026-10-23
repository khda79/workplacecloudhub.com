"""Run the actual NativeEvidenceTypes M expression on synthetic inputs only.

Set PQTEST_PATH to Microsoft.PowerQuery.SdkTools/tools/PQTest.exe, then run
python -B -m unittest test_native_evidence_m -v from this directory.
NATIVE_EVIDENCE_EXPRESSIONS optionally selects another expressions.tmdl.
NATIVE_EVIDENCE_FACTORY_PATH optionally selects a staged MCP-exported M function.
No Desktop connection, production source, refresh, or publication is used.
"""
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

from workplace_native_types import source_datetime


PQTEST = os.environ.get('PQTEST_PATH') or shutil.which('PQTest.exe')
EXPRESSIONS = Path(os.environ.get('NATIVE_EVIDENCE_EXPRESSIONS') or (
    Path(__file__).resolve().parents[1] / 'pbip' /
    'SmartWorkplaceIntelligence.SemanticModel' / 'definition' / 'expressions.tmdl'))


def m_text(value):
    return '"' + value.replace('#', '#(#)').replace('"', '""').replace(
        '\r', '#(cr)').replace('\n', '#(lf)') + '"'


def native_expression():
    staged = os.environ.get('NATIVE_EVIDENCE_FACTORY_PATH')
    if staged:
        return Path(staged).read_text(encoding='utf-8-sig')
    match = re.search(
        r'^expression NativeEvidenceTypes =\r?\n(.*?)(?=^\tlineageTag:)',
        EXPRESSIONS.read_text(encoding='utf-8-sig'), re.M | re.S)
    if not match:
        raise AssertionError('NativeEvidenceTypes expression not found')
    return textwrap.dedent(match.group(1)).strip()


def synthetic_cases():
    cases = []

    def check(name, expression):
        cases.append('[Name=' + m_text(name) + ', Run=() => ' + expression + ']')

    def accepts(name, value, expected, wall_clock='true', empty_contact='false'):
        raw = 'null' if value is None else m_text(value)
        check(name, f'Convert({raw}, {wall_clock}, {empty_contact}) = {expected}')

    def rejects(name, value, wall_clock='true', empty_contact='false'):
        check(name, '(try Convert(' + m_text(value) + ', ' + wall_clock + ', ' +
              empty_contact + '))[HasError]')

    accepts('ISO timestamp', '2024-06-18T09:08:07', '#datetime(2024,6,18,9,8,7)')
    accepts('ISO date only', '2024-06-18', '#datetime(2024,6,18,0,0,0)')
    accepts('legacy day first', '05/06/2024 09:08:07', '#datetime(2024,6,5,9,8,7)')
    accepts('leap day', '2024-02-29T00:00:00', '#datetime(2024,2,29,0,0,0)')
    accepts('minimum source year', '1900-01-01T00:00:00', '#datetime(1900,1,1,0,0,0)')
    accepts('null stays null', None, 'null')
    accepts('blank stays null', '', 'null')
    accepts('UTC remains UTC', '2024-06-18T09:08:07Z', '#datetime(2024,6,18,9,8,7)', 'false')
    accepts('positive offset normalized', '2024-06-18T09:08:07+02:00',
            '#datetime(2024,6,18,7,8,7)', 'false')
    accepts('negative offset normalized', '2024-06-18T09:08:07-05:30',
            '#datetime(2024,6,18,14,38,7)', 'false')
    accepts('Autopilot sentinel opted in', '01/01/0001 00:00:00', 'null', 'true', 'true')
    accepts('Autopilot blank', '', 'null', 'true', 'true')
    for digits in range(1, 8):
        fraction = '1234567'[:digits]
        value = '2024-06-18T09:08:07.' + fraction
        format_string = "yyyy-MM-dd'T'HH:mm:ss." + 'f' * digits
        check(f'ISO fraction {digits}', 'DateTime.ToText(Convert(' + m_text(value) +
              ', true, false), [Format=' + m_text(format_string) +
              ', Culture="en-US"]) = ' + m_text(value))
    for name, value in (
        ('ISO requires opt in', '2024-06-18T09:08:07'),
        ('ISO date requires opt in', '2024-06-18'),
        ('legacy requires opt in', '05/06/2024 09:08:07')):
        rejects(name, value, 'false')
    for name, value in (
        ('invalid calendar date', '2024-02-30T09:08:07'),
        ('invalid legacy date', '31/02/2024 09:08:07'),
        ('month first forbidden', '06/18/2024 09:08:07'),
        ('before minimum year', '1899-12-31T23:59:59'),
        ('ISO zero year', '0001-01-01T00:00:00'),
        ('no fraction digits', '2024-06-18T09:08:07.'),
        ('excessive fraction', '2024-06-18T09:08:07.12345678'),
        ('non numeric fraction', '2024-06-18T09:08:07.12x'),
        ('leading whitespace forbidden', ' 2024-06-18T09:08:07'),
        ('trailing whitespace forbidden', '2024-06-18T09:08:07 '),
        ('non padded date', '2024-6-18T09:08:07'),
        ('ambiguous short date', '05/06/2024'),
        ('invalid hour', '2024-06-18T25:08:07'),
        ('sentinel not opted in', '01/01/0001 00:00:00'),
        ('arbitrary text', 'not a date'),
        ('invalid explicit offset', '2024-06-18T09:08:07+99:00')):
        rejects(name, value)
    check('wall clock remains timezone free',
          'Value.Is(Convert("2024-06-18T09:08:07", true, false), type datetime) '
          'and not Value.Is(Convert("2024-06-18T09:08:07", true, false), type datetimezone)')
    check('missing column remains an error',
          '(try NativeEvidenceTypes(#table({"Other"}, {{"x"}}), '
          '{[name="Value", dataType="DateTime", allowSourceWallClock=true]}){0}[Value])[HasError]')
    check('default timezone guard remains strict',
          '(try NativeEvidenceTypes(#table({"Value"}, {{"2024-06-18T09:08:07"}}), '
          '{[name="Value", dataType="DateTime"]}){0}[Value])[HasError]')

    # Feed the real Python producer's CSV representation to the actual M consumer.
    for table, columns, inputs in (
        ('Native User Evidence', ['LastSignInDateTime', 'LastNonInteractiveSignInDateTime',
                                  'LastSuccessfulSignInDateTime'],
         ['18/06/2024 09:08:07', '2024-06-18T10:08:07.1234567', '2024-06-18T11:08:07Z']),
        ('Autopilot Evidence', ['LastContactedDateTime'], ['18/06/2024 09:08:07']),
        ('Native Configuration Catalog', ['CreatedDateTime', 'LastModifiedDateTime'],
         ['18/06/2024 09:08:07', '2024-06-19T10:08:07']),
        ('Native Update Policy Evidence', ['CreatedDateTime', 'LastModifiedDateTime'],
         ['18/06/2024 09:08:07', '2024-06-19T10:08:07'])):
        stream = io.StringIO(newline='')
        writer = csv.writer(stream)
        writer.writerow(columns)
        produced = [source_datetime(value) for value in inputs]
        writer.writerow(produced)
        writer.writerow([source_datetime('01/01/0001 00:00:00')] * len(columns))
        definitions = ','.join('[name=' + m_text(column) +
                               ',dataType="DateTime",allowSourceWallClock=true]' for column in columns)
        assertions = []
        for column, value in zip(columns, produced):
            expected = value.removesuffix('Z')
            fraction = expected.partition('.')[2]
            format_string = "yyyy-MM-dd'T'HH:mm:ss" + ('.' + 'f' * len(fraction) if fraction else '')
            assertions.append('DateTime.ToText(Projected{0}[' + column + '], [Format=' +
                              m_text(format_string) + ',Culture="en-US"]) = ' + m_text(expected) +
                              ' and Projected{1}[' + column + '] = null')
        check('producer CSV to M: ' + table,
              'let Rows=Table.PromoteHeaders(Csv.Document(Text.ToBinary(' +
              m_text(stream.getvalue()) + ', TextEncoding.Utf8), '
              '[Delimiter=",",Encoding=65001,QuoteStyle=QuoteStyle.Csv])), '
              'Projected=NativeEvidenceTypes(Rows, {' + definitions + '}) '
              'in List.AllTrue({' + ','.join(assertions) + '})')

    lifecycle = [('ModificationRaw', 'String'), ('ModificationUtcDateTime', 'DateTime'),
                 ('ModificationStatus', 'String')]
    password = [('PasswordExpired', 'Boolean'), ('PasswordExpiryStatus', 'String')]
    creation = [('CreationRaw', 'String'), ('CreationUtcDateTime', 'DateTime'), ('CreationStatus', 'String')]
    for table, additions in (
        ('AD User Evidence', lifecycle + password),
        ('AD Computer Evidence', lifecycle + password),
        ('Native User Evidence', lifecycle),
        ('Native Device Evidence', creation + lifecycle),
    ):
        definitions = ','.join('[name=' + m_text(name) + ',dataType=' + m_text(kind) +
                               ',allowHistoricalMissing=true]' for name, kind in additions)
        assertions = ' and '.join('Projected{0}[' + name + '] = null' for name, _ in additions)
        check('old batch nullable additions: ' + table,
              'let Projected=NativeEvidenceTypes(#table({"Existing"},{{"preserved"}}), '
              '{[name="Existing",dataType="String"],' + definitions + '}) '
              'in Projected{0}[Existing] = "preserved" and ' + assertions)
        future_values = ['"2024-06-18T09:08:07Z"' if kind == 'DateTime' else
                         '"false"' if kind == 'Boolean' else '"observed raw/status"'
                         for _, kind in additions]
        future_checks = ' and '.join('Projected{0}[' + name + '] = ' +
                                     ('#datetime(2024,6,18,9,8,7)' if kind == 'DateTime' else
                                      'false' if kind == 'Boolean' else '"observed raw/status"')
                                     for name, kind in additions)
        check('future batch strict additions: ' + table,
              'let Projected=NativeEvidenceTypes(#table({"Existing",' +
              ','.join(m_text(name) for name, _ in additions) + '},{{"preserved",' +
              ','.join(future_values) + '}}), {[name="Existing",dataType="String"],' + definitions + '}) '
              'in Projected{0}[Existing] = "preserved" and ' + future_checks)

    check('nullable PasswordExpired type and strict values', '''let
        Projected=NativeEvidenceTypes(#table({"PasswordExpired"},
            {{true},{false},{null},{""},{"true"},{"false"},{"TRUE"},{"False"}}),
            {[name="PasswordExpired",dataType="Boolean",allowHistoricalMissing=true]})
        in Projected[PasswordExpired] = {true,false,null,null,true,false,true,false} and
            Type.Is(Type.TableColumn(Value.Type(Projected), "PasswordExpired"), type nullable logical)''')
    check('empty legacy rows gain nullable schema without fabricated rows', '''let
        Projected=NativeEvidenceTypes(#table({"Existing"},{}),
            {[name="Existing",dataType="String"],
             [name="PasswordExpired",dataType="Boolean",allowHistoricalMissing=true],
             [name="ModificationUtcDateTime",dataType="DateTime",allowHistoricalMissing=true]})
        in Table.RowCount(Projected) = 0 and Table.HasColumns(Projected, {"PasswordExpired","ModificationUtcDateTime"}) and
            Type.Is(Type.TableColumn(Value.Type(Projected), "PasswordExpired"), type nullable logical)''')
    for bad in ('yes', '0', 'true ', 'unknown'):
        check('malformed lifecycle Boolean: ' + bad,
              '(try NativeEvidenceTypes(#table({"PasswordExpired"},{{' + m_text(bad) + '}}), '
              '{[name="PasswordExpired",dataType="Boolean",allowHistoricalMissing=true]}){0}[PasswordExpired])[HasError]')
    check('new lifecycle dates keep strict UTC conversion', '''let
        Projected=NativeEvidenceTypes(#table({"ModificationUtcDateTime"},{{"2024-06-18T09:08:07+02:00"}}),
            {[name="ModificationUtcDateTime",dataType="DateTime",allowHistoricalMissing=true]})
        in Projected{0}[ModificationUtcDateTime] = #datetime(2024,6,18,7,8,7)''')
    check('new lifecycle malformed date stays error', '''(try NativeEvidenceTypes(
        #table({"ModificationUtcDateTime"},{{"not a date"}}),
        {[name="ModificationUtcDateTime",dataType="DateTime",allowHistoricalMissing=true]}){0}[ModificationUtcDateTime])[HasError]''')
    check('lifecycle optional flag is required for absent columns', '''(try NativeEvidenceTypes(
        #table({"Existing"},{{"x"}}), {[name="PasswordExpired",dataType="Boolean"]}){0}[PasswordExpired])[HasError]''')
    check('non-lifecycle columns cannot suppress missing-field errors', '''(try NativeEvidenceTypes(
        #table({"Existing"},{{"x"}}), {[name="Other",dataType="String",allowHistoricalMissing=true]}){0}[Other])[HasError]''')
    check('historical Boolean kind cannot be weakened to text', '''(try NativeEvidenceTypes(
        #table({"Existing"},{{"x"}}), {[name="PasswordExpired",dataType="String",allowHistoricalMissing=true]}){0}[PasswordExpired])[HasError]''')
    return cases


class NativeEvidenceMTests(unittest.TestCase):
    @unittest.skipUnless(PQTEST, 'Set PQTEST_PATH to run actual Power Query M regression tests')
    def test_actual_m_consumer_and_python_producer(self):
        cases = synthetic_cases()
        query = ('let NativeEvidenceTypes = (' + native_expression() + '),\n'
                 'Convert=(Value,WallClock,EmptyContact)=>NativeEvidenceTypes('
                 '#table({"Value"},{{Value}}),{[name="Value",dataType="DateTime",'
                 'allowSourceWallClock=WallClock,allowEmptyAutopilotContact=EmptyContact]}){0}[Value],\n'
                 'Cases={' + ',\n'.join(cases) + '},\n'
                 'Results=List.Transform(Cases,(Case)=>let Attempt=try Case[Run]() in '
                 '[Name=Case[Name],Passed=if Attempt[HasError] then false else Attempt[Value]=true,'
                 'Detail=if Attempt[HasError] then Attempt[Error][Message] else null])\n'
                 'in Table.FromRecords(Results)')
        with tempfile.TemporaryDirectory(prefix='native-evidence-m-') as scratch:
            query_file = Path(scratch) / 'synthetic.query.pq'
            query_file.write_text(query, encoding='utf-8')
            completed = subprocess.run(
                [PQTEST, 'run-test', '--queryFile', str(query_file)],
                capture_output=True, text=True, encoding='utf-8-sig', timeout=90, cwd=scratch)
        self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)
        results = json.loads(completed.stdout)
        self.assertEqual(len(results), 1, completed.stdout)
        self.assertEqual(results[0]['Status'], 'Passed', completed.stdout)
        self.assertEqual(results[0]['RowCount'], len(cases), completed.stdout)
        for result in results[0]['Output']:
            with self.subTest(case=result['Name']):
                self.assertTrue(result['Passed'], result['Detail'])
        print(f'Executed {len(cases)} synthetic cases through the actual M engine.')


if __name__ == '__main__':
    unittest.main()
