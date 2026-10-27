"""Actual Power Query readiness retirement regressions; synthetic Contoso evidence only."""
import csv, io, json, unittest
from test_mailbox_migration_m import (ROOT, REQUIRED, PREP, MAIL, PERMISSION, JOBS, TENANT,
    batch_files, binaries, source_bindings, expression, partition)
import test_prepared_latest_m as prepared_tests

HYBRID = 'HybridIdentityFindings.csv'
GAP = 'DeviceReadinessGapEvidence.csv'

def readiness_payload(severity='Blocker', blocker='true', tenant=TENANT, empty=False):
    contract = json.loads((ROOT/'config/prepared-evidence-contract.json.txt').read_text(encoding='utf-8-sig'))
    payload = binaries()
    for name in (HYBRID, GAP):
        spec = next(t for t in contract['tables'] if t['file'] == name)
        cols = [c['name'] for c in spec['columns']]
        output = io.StringIO(newline='')
        writer = csv.DictWriter(output, fieldnames=cols, lineterminator='\r\n')
        writer.writeheader()
        if name == HYBRID and not empty:
            row = dict.fromkeys(cols, '')
            row.update({'TenantKey': tenant, 'Finding Key': 'contoso|HYB-001|ad-1',
                        'AD Object GUID': 'ad-1', 'Rule ID': 'HYB-001', 'Legacy Issue Number': '1',
                        'Severity': severity, 'Is Blocker': blocker, 'Observed Value': 'Contoso observation',
                        'Evidence Date': '2026-10-25T12:00:00Z'})
            writer.writerow(row)
        payload[name] = output.getvalue().encode()
    return payload


def bindings(payload):
    values = source_bindings(batch_files(payload), 'Local')
    names = ['PreparedReadFolder', 'CreatePreparedReaders', 'NativeEvidenceTypes',
             'GetPreparedEvidenceFile', 'GetPreparedBatchManifest', 'ReadMigrationCsv',
             'CreateMailboxMigrationReader', 'ReadMailboxMigrationCsv']
    return values + [n+'=('+expression(n)+')' for n in names] + [
        'Readers=CreatePreparedReaders()', 'After=('+partition()+')']


class ReadinessRetirementMTests(unittest.TestCase):
    evaluate = prepared_tests.PreparedLatestMTests.evaluate
    def test_complete_findings_use_blocker_and_not_legacy_category(self):
        payload=readiness_payload()
        # Remove unrelated operational risks from the mailbox fixture.
        for name in (MAIL, JOBS):
            rows=list(csv.DictReader(io.StringIO(payload[name].decode())))
            for row in rows:
                if name == MAIL:
                    row.update({'TotalItemSize-In-MB':'1024','FullAccessCount':'0','SendAsCount':'0'})
                else: row['UserStatus']='Synced'
            out=io.StringIO(newline='');w=csv.DictWriter(out,fieldnames=rows[0].keys());w.writeheader();w.writerows(rows)
            payload[name]=out.getvalue().encode()
        self.evaluate(bindings(payload), '[OK=After{0}[Migration Status]="Observed blocker" and Text.Contains(After{0}[Identity Issues],"HYB-001 - Contoso observation")]')
        payload.update({HYBRID:readiness_payload('Review','false')[HYBRID]})
        self.evaluate(bindings(payload), '[OK=After{0}[Migration Status]<>"Observed blocker" and not Text.Contains(After{0}[Risk Reasons],"Identity issue")]')

    def test_old_batch_has_explicit_unavailable_evidence(self):
        self.evaluate(bindings(binaries()), '[OK=not Readers[GetReadinessSource]("HybridIdentityFindings.csv")[Available] and Text.Contains(After{0}[Next Action],"obtain hybrid identity evidence")]')

    def test_empty_complete_pair_is_available(self):
        self.evaluate(bindings(readiness_payload(empty=True)), '[OK=Readers[GetReadinessSource]("HybridIdentityFindings.csv")[Available] and Readers[GetReadinessSource]("HybridIdentityFindings.csv")[RowCount]=0]')

    def test_partial_pair_foreign_tenant_and_inconsistent_blocker_are_rejected(self):
        partial=readiness_payload();del partial[GAP]
        for payload in (partial,readiness_payload(tenant='foreign'),readiness_payload('Review','true')):
            with self.subTest(case=list(payload)):
                self.evaluate(bindings(payload),'[OK=(try Readers[GetReadinessSource]("HybridIdentityFindings.csv")[RowCount])[HasError]]')

    def test_numeric_backlog_columns_are_explicitly_typed(self):
        model=ROOT/'pbip/SmartWorkplaceIntelligence.SemanticModel/definition/tables'
        for table,column in [('Hybrid Identity Findings','Severity Sort'),('Mailbox Account License Coverage','M365 License Count')]:
            source=(model/(table+'.tmdl')).read_text(encoding='utf-8-sig')
            self.assertIn('"'+column+'", Int64.Type',source)
            self.assertIn("column '"+column+"'\n\t\tdataType: int64",source)
