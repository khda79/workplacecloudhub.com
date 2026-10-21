"""Execute the isolated server-reference reader using synthetic binaries only."""
import copy
import csv
import io
import json
import subprocess
import tempfile
from pathlib import Path
import unittest

import test_mailbox_migration_m as h
from exchange_reconciliation_model import server_expression


def cases():
    row=dict(TenantKey=h.TENANT,ObjectGUID=' OBJECT-A ',RecipientType='UserMailbox',ServerName='synthetic-server')
    result=[]
    def add(name,rows,check='true',error=False,mode='Local',without_server=False):
        fields=['TenantKey','ObjectGUID','RecipientType']+([] if without_server else ['ServerName'])
        stream=io.StringIO();writer=csv.DictWriter(stream,fieldnames=fields,extrasaction='ignore');writer.writeheader();writer.writerows(rows)
        payload=h.binaries();payload[h.MAIL]=stream.getvalue().encode();files=h.batch_files(payload)
        bindings=h.source_bindings(files,mode)
        for n in ['PreparedReadFolder','CreatePreparedReaders','NativeEvidenceTypes','GetPreparedEvidenceFile','GetPreparedBatchManifest','ReadMigrationCsv','CreateMailboxMigrationReader','ReadMailboxMigrationCsv']:
            bindings.append(n+'=('+h.expression(n)+')')
        bindings.append('Value=try ('+server_expression()+')')
        expr='Value[HasError]' if error else 'not Value[HasError] and (let After=Value[Value] in '+check+')'
        result.append((name,'let '+',\n'.join(bindings)+'\nin '+expr))
    add('Native object key and server retained',[row],check='After{0}[Object GUID]="object-a" and After{0}[Server Name]="synthetic-server" and After{0}[Loaded Batch ID]='+h.m_text(h.BATCH))
    add('SharePoint guarded read',[row],mode='SharePoint',check='Table.RowCount(After)=1')
    add('Historical optional server absent',[row],without_server=True,check='After{0}[Server Name]="Not observed"')
    add('Blank server remains unknown',[dict(row,ServerName='')],check='After{0}[Server Name]="Not observed"')
    add('Duplicate native object rejected',[row,row],error=True)
    add('Missing native object rejected',[dict(row,ObjectGUID='')],error=True)
    add('Cross tenant rejected',[dict(row,TenantKey='other-synthetic-tenant')],error=True)
    add('Remote recipients excluded',[row,dict(row,RecipientType='RemoteUserMailbox')],check='Table.RowCount(After)=1')
    add('Empty source remains empty',[],check='Table.RowCount(After)=0 and List.Contains(Table.ColumnNames(After),"Loaded Batch ID")')
    return result


class ServerReferenceTests(unittest.TestCase):
    @unittest.skipUnless(h.PQTEST,'Set PQTEST_PATH for the actual M engine')
    def test_actual_m(self):
        tests=cases();entries=['[Name='+h.m_text(name)+',Run=()=>('+body+')]' for name,body in tests]
        query='let Cases={'+','.join(entries)+'}, Results=List.Transform(Cases,(c)=>let r=try c[Run]() in [Name=c[Name],Passed=if r[HasError] then false else r[Value]=true,Detail=if r[HasError] then r[Error][Message] else null]) in Table.FromRecords(Results)'
        with tempfile.TemporaryDirectory(prefix='server-reference-m-') as directory:
            p=Path(directory)/'server.query.pq';p.write_text(query)
            r=subprocess.run([h.PQTEST,'run-test','--queryFile',str(p)],cwd=directory,capture_output=True,text=True,encoding='utf-8-sig',timeout=180)
        self.assertEqual(r.returncode,0,r.stderr)
        output=json.loads(r.stdout)[0];self.assertEqual(output['Status'],'Passed',r.stdout[-1000:])
        for case in output['Output']:
            with self.subTest(name=case['Name']):self.assertTrue(case['Passed'],case['Detail'])
        print('Executed '+str(len(tests))+' isolated server-reference M cases.')


if __name__=='__main__':unittest.main()
