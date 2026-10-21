"""Version 1.0.0: actual M checks over synthetic AD classification payloads."""
import csv,io,json,subprocess,tempfile
from pathlib import Path
import unittest
import test_mailbox_migration_m as h
from exchange_account_classification import reader_expression,EXPRESSION


def cases():
    base=dict(TenantKey=h.TENANT,ObjectGUID=' OBJECT-A ',TenantADUserKey='ad-a',AccountType='Service Account')
    result=[]
    def add(name,rows,check='true',error=False,mode='Local',without_type=False,mutate=None):
        fields=['TenantKey','ObjectGUID','TenantADUserKey']+([] if without_type else ['AccountType'])
        stream=io.StringIO();writer=csv.DictWriter(stream,fieldnames=fields,extrasaction='ignore');writer.writeheader();writer.writerows(rows)
        payload=h.binaries();payload['ADUserSource.csv']=stream.getvalue().encode();files=h.batch_files(payload,mutate)
        bindings=h.source_bindings(files,mode)
        for n in ['PreparedReadFolder','CreatePreparedReaders','NativeEvidenceTypes','GetPreparedEvidenceFile','GetPreparedBatchManifest','ReadMigrationCsv','CreateMailboxMigrationReader','ReadMailboxMigrationCsv']:
            bindings.append(n+'=('+h.expression(n)+')')
        bindings.append('Reader=('+h.expression(EXPRESSION)+')');bindings.append('Value=try Reader()')
        expr='Value[HasError]' if error else 'not Value[HasError] and (let After=Value[Value] in '+check+')'
        result.append((name,'let '+',\n'.join(bindings)+'\nin '+expr))
    add('Native identity and batch',[base],check='After{0}[Object GUID]="object-a" and After{0}[AD Account Key]="ad-a" and After{0}[Loaded Batch ID]='+h.m_text(h.BATCH))
    add('SharePoint guarded reader',[base],mode='SharePoint',check='After{0}[Service Classification]="Service"')
    add('Historical field absent',[base],without_type=True,check='After{0}[Account Type]="Not captured" and After{0}[Service Classification]="Not captured"')
    add('Blank field unknown',[dict(base,AccountType='')],check='After{0}[Service Classification]="Not captured"')
    types=['Named Account','Ext Account','Generic Account','Admin Account','System Account','Resource Account','Service Account','Unclassified Account','Future Custom Type']
    add('Every raw type preserved',[dict(base,ObjectGUID='object-'+str(i),TenantADUserKey='ad-'+str(i),AccountType=t) for i,t in enumerate(types)],
        check='Table.RowCount(After)=9 and List.Contains(After[Account Type],"Future Custom Type") and Table.SelectRows(After,each [Account Type]="Generic Account"){0}[Service Classification]<>"Service"')
    add('Duplicate native object rejected',[base,dict(base,TenantADUserKey='ad-b')],error=True)
    add('Duplicate account key rejected',[base,dict(base,ObjectGUID='object-b')],error=True)
    add('Blank native identity rejected',[dict(base,ObjectGUID='')],error=True)
    add('Cross tenant rejected',[dict(base,TenantKey='different')],error=True)
    add('Empty source remains typed',[],check='Table.RowCount(After)=0 and List.Contains(Table.ColumnNames(After),"Loaded Batch ID")')
    add('Wrong manifest batch rejected',[base],error=True,mutate=lambda manifest,pointer:manifest.update(BatchId='other'))
    add('Wrong byte size rejected',[base],error=True,mutate=lambda manifest,pointer:next(x for x in manifest['Artifacts'] if x['File']=='ADUserSource.csv').update(Bytes=1))
    return result


class AccountClassificationTests(unittest.TestCase):
    def test_committed_reader_matches_reviewed_definition(self):
        self.assertEqual(h.expression(EXPRESSION).strip(),reader_expression().strip())

    @unittest.skipUnless(h.PQTEST,'Set PQTEST_PATH for the actual M engine')
    def test_actual_m(self):
        tests=cases();entries=['[Name='+h.m_text(name)+',Run=()=>('+body+')]' for name,body in tests]
        query='let Cases={'+','.join(entries)+'}, Results=List.Transform(Cases,(c)=>let r=try c[Run]() in [Name=c[Name],Passed=if r[HasError] then false else r[Value]=true,Detail=if r[HasError] then r[Error][Message] else null]) in Table.FromRecords(Results)'
        with tempfile.TemporaryDirectory(prefix='account-classification-m-') as directory:
            p=Path(directory)/'classification.query.pq';p.write_text(query)
            r=subprocess.run([h.PQTEST,'run-test','--queryFile',str(p)],cwd=directory,capture_output=True,text=True,encoding='utf-8-sig',timeout=180)
        self.assertEqual(r.returncode,0,r.stderr)
        output=json.loads(r.stdout)[0];self.assertEqual(output['Status'],'Passed',r.stdout[-1000:])
        for case in output['Output']:
            with self.subTest(name=case['Name']):self.assertTrue(case['Passed'],case['Detail'])
        print('Executed '+str(len(tests))+' isolated AD classification M cases.')


if __name__=='__main__':unittest.main()
