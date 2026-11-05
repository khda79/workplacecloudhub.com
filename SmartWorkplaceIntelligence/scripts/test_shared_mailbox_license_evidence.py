"""Version 0.1.0: synthetic reconciliation and qualification tests."""
import unittest
import csv
import hashlib
import json
import tempfile
from copy import deepcopy
from pathlib import Path
from shared_mailbox_license_evidence import evaluate, collect, collect_manifest, FILES


class SharedMailboxEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.mail = [{'RecipientType':'SharedMailbox','ObjectGUID':'obj1','ExchangeGuid':'mail1',
                      'NativeIdentityStatus':'Observed','PrimarySMTPaddress':'a@example.invalid','TotalItemSize-In-MB':'60000'}]
        self.ad = [{'ObjectGUID':'obj1','TenantUserKey':'user1','CloudMatchStatus':'Unique native SID','AccountType':'Generic Account'}]
        self.sku = [{'TenantSkuKey':'sku1','SkuPartNumber':'SPE_E3'}]
        self.licenses = [{'TenantUserKey':'user1','TenantSkuKey':'sku1','AssignmentState':'Active','UserLinkStatus':'Resolved'}]
        self.plans = [{'TenantUserKey':'user1','TenantSkuKey':'sku1','ServicePlanName':'EXCHANGE_S_ENTERPRISE','IsEnabled':'true','AssignmentState':'Success','UserLinkStatus':'Resolved'}]
        self.exceptions = []

    def result(self):
        return evaluate(self.mail, self.exceptions, self.ad, self.licenses, self.plans, self.sku)

    def test_covered_e3_is_counted_once(self):
        self.ad[0]['AccountType'] = 'Shared Mailbox'
        self.plans.append(deepcopy(self.plans[0]))
        result = self.result()
        self.assertEqual((result['Required'], result['AssignedCoverage'], result['ToComplete'], result['E3ToPlan2Review']), (1, 1, 0, 1))

    def test_retained_exclusion_and_pending_hint(self):
        self.exceptions = [{'Mailbox Identity':'local_mailboxes|obj1','Candidate Identity':'', 'Primary SMTP Address':'a@example.invalid',
                            'Retained':'1','Requires Review':'0','Status':'Matched local'}]
        self.assertEqual((self.result()['Excluded'], self.result()['Required']), (1, 0))
        self.exceptions[0].update({'Retained':'0','Requires Review':'1','Status':'Review'})
        self.assertEqual((self.result()['IdentityOrRegisterReview'], self.result()['Required']), (1, 0))

    def test_unknown_assignment_and_boundary_are_not_zero_coverage(self):
        self.plans = []
        self.assertEqual(self.result()['CoverageReview'], 1)
        self.mail[0]['TotalItemSize-In-MB'] = '51200'
        self.assertEqual((self.result()['SizeReview'], self.result()['Required']), (1, 0))

    def test_uncovered_and_small_recovery_are_distinct(self):
        self.plans[0]['ServicePlanName'] = 'EXCHANGE_S_STANDARD'
        self.assertEqual(self.result()['ToComplete'], 1)
        self.mail[0]['TotalItemSize-In-MB'] = '10000'
        self.assertEqual((self.result()['Under50'], self.result()['E3Under50Review'], self.result()['Required']), (1, 1, 0))

    def test_ambiguous_links_and_missing_classification_are_review(self):
        self.mail.append(dict(self.mail[0], ObjectGUID='obj2', ExchangeGuid='mail2'))
        self.ad.append(dict(self.ad[0], ObjectGUID='obj2'))
        self.assertEqual((self.result()['IdentityOrRegisterReview'], self.result()['Required']), (2, 0))
        self.mail.pop(); self.ad.pop(); self.ad[0]['AccountType']=''
        self.assertEqual(self.result()['IdentityOrRegisterReview'], 1)

    def test_duplicate_mailbox_or_ad_identity_fails(self):
        self.mail.append(deepcopy(self.mail[0]))
        with self.assertRaises(ValueError): self.result()
        self.mail.pop(); self.ad.append(deepcopy(self.ad[0]))
        with self.assertRaises(ValueError): self.result()

    def test_prepared_batch_integrity_and_tenant(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); batch='20261103T070757019Z-f084240f'; folder=root/'batches'/batch
            folder.mkdir(parents=True)
            data=[self.mail, [], self.ad, self.licenses, self.plans, [dict(self.sku[0],SkuId='sku-id')]]
            fields=[['TenantKey','RecipientType','ObjectGUID','ExchangeGuid','NativeIdentityStatus','PrimarySMTPaddress','TotalItemSize-In-MB'],
                    ['TenantKey','Mailbox Identity','Candidate Identity','Primary SMTP Address','Retained','Requires Review','Status'],
                    ['TenantKey','ObjectGUID','TenantUserKey','CloudMatchStatus','AccountType'],
                    ['TenantKey','TenantUserKey','TenantSkuKey','AssignmentState','UserLinkStatus'],
                    ['TenantKey','TenantUserKey','TenantSkuKey','ServicePlanName','IsEnabled','AssignmentState','UserLinkStatus'],
                    ['TenantKey','TenantSkuKey','SkuPartNumber','SkuId']]
            entries=[]
            for name, rows, columns in zip(FILES,data,fields):
                path=folder/name
                with path.open('w',encoding='utf-8',newline='') as stream:
                    writer=csv.DictWriter(stream,fieldnames=columns);writer.writeheader()
                    for row in rows: writer.writerow(dict(row,TenantKey='test'))
                entries.append({'File':name,'SHA256':hashlib.sha256(path.read_bytes()).hexdigest(),'Rows':len(rows)})
            source_names=['Exchange_OnPrem_Mailboxes_AllDomains.csv','M365_Licenses_Tenant.csv','M365_Licenses_Users.csv',
                          'M365_Licenses_ServicePlans_Catalog.csv','M365_Licenses_UserServicePlanStates.csv','AD_Users_AllDomains.csv']
            proofs=[{'File':name,'Status':'Success','IsPartialInventory':False,'Errors':0,'RunId':'run1',
                     'SHA256':entries[0]['SHA256'] if i==0 else '0'*64} for i,name in enumerate(source_names)]
            manifest={'BatchId':batch,'TenantKey':'test','CreatedUtc':'2026-11-03T07:10:00Z',
                      'Files':entries,'Artifacts':[], 'LicenseReportQualification':{'State':'Qualified','ReceiptRunId':'run1'},
                      'Provenance':{'MailboxMigrationExceptions':{'Audit':{'Source Status':'Qualified','Identity':{'TenantKey':'test'}}},
                                    'NativeEvidence':{'Audit':{'Identity':{'TenantKey':'test'},'Evidence':{'Files':proofs}}}}}
            manifest_path=folder/'batch.json.txt';manifest_path.write_text(json.dumps(manifest),encoding='utf-8')
            pointer={'SchemaVersion':1,'BatchId':batch,'TenantKey':'test','ManifestSHA256':hashlib.sha256(manifest_path.read_bytes()).hexdigest()}
            (root/'current.json.txt').write_text(json.dumps(pointer),encoding='utf-8')
            self.assertEqual(collect(root,'test')['Counts']['AssignedCoverage'],1)
            self.assertEqual(collect_manifest(manifest_path, 'test'), collect(root, 'test'))
            with self.assertRaises(ValueError): collect(root,'other')
            with self.assertRaises(ValueError): collect_manifest(manifest_path,'other')
            with (folder/FILES[0]).open('a',encoding='utf-8') as stream: stream.write('tamper')
            with self.assertRaises(ValueError): collect(root,'test')
            with self.assertRaises(ValueError): collect_manifest(manifest_path,'test')


if __name__ == '__main__':
    unittest.main()
