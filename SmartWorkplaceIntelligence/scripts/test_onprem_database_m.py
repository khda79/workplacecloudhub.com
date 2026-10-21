"""Execute same-batch database reader against synthetic binaries only."""
import copy
import json
from pathlib import Path
import subprocess
import tempfile
import unittest

import test_exchange_infrastructure_m as infra
from onprem_database_model import database_expression


def fixtures():
    readiness = 'Source-Exchange_OnPrem_MigrationReadiness_Config.csv'
    mailbox = 'Source-Exchange_OnPrem_Mailboxes_AllDomains.csv'
    columns = ['TenantKey','Category','ObjectName','Setting','Value','CollectionStatus']
    rows = [['synthetic','MailboxDatabase','DB-A',key,value,'OK'] for key,value in
            [('Mounted','True'),('Server','mail-one'),('DatabaseSize','2 GB (2,147,483,648 bytes)'),
             ('AvailableNewMailboxSpace','1 GB (1,073,741,824 bytes)')]]
    rows += [['synthetic','MailboxDatabase','DB-B',key,value,'OK'] for key,value in
             [('Mounted','False'),('Server','mail-two'),('DatabaseSize','Unknown')]]
    mailcols = ['TenantKey','ObjectGUID','RecipientType','Database']
    mails = [['synthetic','object-a','UserMailbox','db-a'],['synthetic','object-b','SharedMailbox','']]
    files = {readiness:infra.csv_bytes(columns,rows), mailbox:infra.csv_bytes(mailcols,mails)}
    tests=[]
    def add(name,source,check='true',error=False,**options):
        bindings=infra.fixture(source,**options)
        bindings.append('Databases=try ('+database_expression()+')')
        result='Databases[HasError]' if error else 'not Databases[HasError] and (let D=Databases[Value] in '+check+')'
        tests.append((name,'let\n'+',\n'.join(bindings)+'\nin '+result))
    add('Names status bytes server and mailbox counts',files,
        '''let A=Table.SelectRows(D,each [Database]="DB-A"){0}, B=Table.SelectRows(D,each [Database]="DB-B"){0}
        in Table.RowCount(D)=3 and A[Database Status]="Mounted" and A[Server Name]="mail-one" and
        A[Observed Mailboxes]=1 and A[Database Size GB]=2 and A[Reusable Internal Space GB]=1 and
        B[Database Status]="Unmounted" and B[Observed Mailboxes]=0 and B[Database Size GB]=null and
        Table.SelectRows(D,each not [Is Database]){0}[Observed Mailboxes]=1 and
        List.Distinct(D[Loaded Batch ID])={"batch-a"}''')
    add('Absent infrastructure retains references but marks unknown',{mailbox:files[mailbox]},
        'List.AllTrue(List.Transform(D[Database Source Available],each _=false)) and D{0}[Database Status]="Unknown"')
    add('No sources does not manufacture a database',{},
        'Table.RowCount(D)=1 and D{0}[Is Database]=false and D{0}[Observed Mailboxes]=null')
    add('Absent mailbox source keeps database counts unknown',{readiness:files[readiness]},
        'List.AllTrue(List.Transform(D[Observed Mailboxes],each _=null))')
    missing=dict(files);missing[mailbox]=infra.csv_bytes(mailcols[:3],[m[:3] for m in mails])
    add('Absent Database column remains unknown',missing,
        'List.AllTrue(List.Transform(D[Observed Mailboxes],each _=null))')
    bad=dict(files);bad[mailbox]=infra.csv_bytes(mailcols,[mails[0],mails[0]])
    add('Duplicate native mailbox identity rejected',bad,error=True)
    bad=dict(files);bad[mailbox]=infra.csv_bytes(mailcols,[['foreign','object-a','UserMailbox','DB-A']])
    add('Foreign mailbox tenant rejected',bad,error=True)
    add('Declared file size mismatch rejected',files,error=True,bad_size=True)
    add('Changed pointer rejected',files,error=True,change='verify-manifest')
    bad=dict(files);bad[readiness]=infra.csv_bytes(columns,rows+[rows[0][:4]+['False','OK']])
    add('Conflicting mount states are explicit',bad,
        'Table.SelectRows(D,each [Database]="DB-A"){0}[Database Status]="Conflicting"')
    return tests


@unittest.skipUnless(infra.PQTEST,'Set PQTEST_PATH for actual M execution')
class DatabaseReaderTests(unittest.TestCase):
    def test_actual_m(self):
        tests=fixtures()
        cases=','.join('[Name='+infra.m_text(n)+',Run=()=>('+q+')]' for n,q in tests)
        query='let C={'+cases+'},R=List.Transform(C,(c)=>let r=try c[Run]() in [Name=c[Name],Passed=if r[HasError] then false else r[Value]=true,Detail=if r[HasError] then r[Error][Message] else null]) in Table.FromRecords(R)'
        with tempfile.TemporaryDirectory(prefix='exchange-database-test-') as directory:
            p=Path(directory)/'database.query.pq';p.write_text(query,encoding='utf-8')
            r=subprocess.run([infra.PQTEST,'run-test','--queryFile',str(p)],cwd=directory,capture_output=True,text=True,encoding='utf-8-sig',timeout=180)
        self.assertEqual(r.returncode,0,r.stderr)
        result=json.loads(r.stdout)[0]
        self.assertEqual(result['Status'],'Passed',r.stdout[-1500:])
        for case in result['Output']:
            with self.subTest(name=case['Name']):self.assertTrue(case['Passed'],case['Detail'])
        print('Executed '+str(len(tests))+' synthetic Exchange database M cases.')
