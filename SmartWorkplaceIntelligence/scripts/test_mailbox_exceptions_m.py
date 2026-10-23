"""Execute the actual optional snapshot M queries on synthetic manifest/CSV data."""
import base64
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

import workplace_mailbox_exceptions as subject

PQTEST = os.environ.get('PQTEST_PATH') or shutil.which('PQTest.exe')
MODEL = Path(os.environ.get('SMARTWORKPLACE_PBIP') or Path(__file__).resolve().parents[1] / 'pbip') / 'SmartWorkplaceIntelligence.SemanticModel/definition/tables'


def m_text(s):
    return '"' + s.replace('#', '#(#)').replace('"', '""').replace('\n', '#(lf)').replace('\r', '#(cr)') + '"'


def model_expression(name):
    source = (MODEL.parent / 'expressions.tmdl').read_text(encoding='utf-8-sig')
    match = re.search(r'^expression ' + re.escape(name) + r' =([^\n]*(?:\n(?!\tlineageTag:|expression |///).*)*)', source, re.M)
    if not match:
        raise AssertionError('Missing expression: ' + name)
    body = re.split(r'\n\t(?:lineageTag:|annotation |queryGroup:)', match[1])[0]
    return textwrap.dedent(body.strip('\r\n')).strip()


def binary(value):
    if not isinstance(value, bytes):
        value = json.dumps(value).encode('utf-8')
    return 'Binary.FromText(' + m_text(base64.b64encode(value).decode('ascii')) + ', BinaryEncoding.Base64)'


class SnapshotMTests(unittest.TestCase):
    @unittest.skipUnless(PQTEST, 'Set PQTEST_PATH for actual Power Query tests')
    def test_optional_tables_and_fail_closed_guards(self):
        cases = []
        names = ['MailboxMigrationExceptions.csv', 'MailboxMigrationExceptionSummary.csv']
        factory = model_expression('CreatePreparedReaders')
        native_types = model_expression('NativeEvidenceTypes')
        for table, columns in [('Mailbox Migration Exceptions', subject.DECISION_COLUMNS),
                               ('Mailbox Migration Exception Summary', subject.SUMMARY_COLUMNS)]:
            source = (MODEL / (table + '.tmdl')).read_text(encoding='utf-8-sig')
            match = re.search(r'^\s+source =\s*\n(.*?)(?=\n\s+annotation |\Z)', source, re.M | re.S)
            self.assertIsNotNone(match)
            expression = textwrap.dedent(match[1]).strip()
            for case in ['old batch', 'header only', 'one row', 'foreign tenant', 'wrong header', 'partial pair', 'declared missing']:
                row = dict.fromkeys(columns, '')
                row['TenantKey'] = 'foreign' if case == 'foreign tenant' else 'test'
                output = io.StringIO()
                writer = csv.DictWriter(output, columns if case != 'wrong header' else columns[:-1]);writer.writeheader()
                if case in ('one row', 'foreign tenant'): writer.writerow(row)
                payload = output.getvalue().encode('utf-8')
                present = [] if case in ('old batch', 'declared missing') else names[:1 if case == 'partial pair' else 2]
                pointer = {'SchemaVersion': 1, 'BatchId': 'batch-test', 'TenantKey': 'test'}
                manifest = dict(pointer, Files=[{'File': n, 'Bytes': len(payload)} for n in present],
                                Provenance={} if case == 'old batch' else {'MailboxMigrationExceptions': {}})
                folder_rows = [('batch.json.txt', binary(manifest))] + [(n, binary(payload)) for n in present]
                folder = '(segments as list, phase as text) as table => if List.IsEmpty(segments) then '
                folder += '#table({"Name","Content"}, {{"current.json.txt",' + binary(pointer) + '}}) else '
                folder += '#table({"Name","Content"}, {' + ','.join('{' + m_text(n) + ',' + content + '}' for n, content in folder_rows) + '})'
                # Only folder I/O is fictional; execute the shipped factory and table queries.
                prefix = ('let PreparedExpectedTenantKey="test", PreparedBatchId="Latest", '
                          'PreparedReadFolder=' + folder + ', NativeEvidenceTypes=' + native_types + ', '
                          'CreatePreparedReaders=' + factory + ', Result=(' + expression + ')')
                error = case in ('foreign tenant', 'wrong header', 'partial pair', 'declared missing')
                check = '(try Table.RowCount(Result))[HasError]' if error else 'Table.RowCount(Result)='+('1' if case=='one row' else '0')
                # Force every column/type conversion, not just lazy row enumeration.
                if case == 'foreign tenant': check = '(try Table.ToRows(Table.Buffer(Result)))[HasError]'
                cases.append('[Name='+m_text(table+'/'+case)+',Run=()=>('+prefix+' in '+check+')]')
        query='let Cases={'+','.join(cases)+'}, Results=List.Transform(Cases,(c)=>let r=try c[Run]() in [Name=c[Name],Passed=if r[HasError] then false else r[Value]=true,Detail=if r[HasError] then r[Error][Message] else null]) in Table.FromRecords(Results)'
        with tempfile.TemporaryDirectory(prefix='mailbox-exception-m-') as temp:
            path=Path(temp)/'test.query.pq';path.write_text(query,encoding='utf-8')
            completed=subprocess.run([PQTEST,'run-test','--queryFile',str(path)],cwd=temp,capture_output=True,text=True,encoding='utf-8-sig',timeout=180)
        self.assertEqual(completed.returncode,0,completed.stdout[-3000:]+completed.stderr[-1000:])
        result=json.loads(completed.stdout)[0]
        self.assertEqual(result['Status'],'Passed',completed.stdout[-3000:])
        for row in result['Output']:
            with self.subTest(case=row['Name']):self.assertTrue(row['Passed'],row['Detail'])
        self.assertEqual(result['RowCount'],14)
        print('Executed 14 mailbox exception scenarios through the actual reader factory and M engine.')


if __name__ == '__main__':
    unittest.main()
