"""Qualification tests for shared mailbox license detail."""
import csv
import json
import tempfile
import unittest
from pathlib import Path
from shared_mailbox_license_detail import evaluate, from_files, sha256, size_gb


class SharedMailboxDetailsTests(unittest.TestCase):
    def setUp(self):
        self.snapshot = {'TenantKey': 'test', 'Products': [
            {'Product': 'Microsoft 365 E3', 'UsageAvailable': True, 'Counts': {'SharedUnknown': 0, 'SharedLicensed': 1, 'SharedUnder50': 0, 'SharedEligible': 0, 'SharedE3Plan2Review': 1}},
            {'Product': 'Microsoft 365 F3', 'UsageAvailable': True, 'Counts': {'SharedUnknown': 0, 'SharedLicensed': 1, 'SharedUnder50': 1, 'SharedEligible': 1, 'SharedE3Plan2Review': 0}},
            {'Product': 'Microsoft 365 F1', 'UsageAvailable': True, 'Counts': {'SharedUnknown': 0, 'SharedLicensed': 0, 'SharedUnder50': 0, 'SharedEligible': 0, 'SharedE3Plan2Review': 0}},
            {'Product': 'Microsoft 365 E5', 'UsageAvailable': True, 'Counts': {'SharedUnknown': 0, 'SharedLicensed': 0, 'SharedUnder50': 0, 'SharedEligible': 0, 'SharedE3Plan2Review': 0}},
        ]}
        self.licenses = [dict(TenantKey='test', UserId='one', SkuPartNumber='SPE_E3'), dict(TenantKey='test', UserId='two', SkuPartNumber='SPE_F1')]
        self.mailboxes = [self.mailbox('one', '65,2'), self.mailbox('two', '10')]

    @staticmethod
    def mailbox(user, size):
        return dict(TenantKey='test', ExternalDirectoryObjectId=user, NativeIdentityStatus='Observed', RecipientTypeDetails='SharedMailbox', TotalItemSizeGB=size, ArchiveStatus='None', LitigationHoldEnabled='False', RetentionHoldEnabled='False')

    def test_reconciled_detail(self):
        rows = evaluate(self.snapshot, 'test', self.licenses, self.mailboxes)
        self.assertEqual([x['Review'] for x in rows], ['E3 to Plan 2', 'Removal candidate'])

    def test_changed_guard_fails_reconciliation(self):
        self.mailboxes[1]['LitigationHoldEnabled'] = 'True'
        with self.assertRaisesRegex(ValueError, 'SharedEligible'):
            evaluate(self.snapshot, 'test', self.licenses, self.mailboxes)

    def test_duplicate_identity_fails_closed(self):
        with self.assertRaisesRegex(ValueError, 'Duplicate mailbox'):
            evaluate(self.snapshot, 'test', self.licenses, self.mailboxes + [self.mailboxes[0]])

    def test_tenant_and_boundary(self):
        with self.assertRaisesRegex(ValueError, 'tenant'):
            evaluate(self.snapshot, 'wrong', self.licenses, self.mailboxes)
        self.assertIsNone(size_gb('1e2'))
        self.assertIsNone(size_gb('-1'))
        self.mailboxes[0]['TotalItemSizeGB'] = '100'
        with self.assertRaisesRegex(ValueError, 'SharedE3Plan2Review'):
            evaluate(self.snapshot, 'test', self.licenses, self.mailboxes)

    def test_source_hash_and_schema_fail_closed(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            licenses = root / 'M365_Licenses_Users.csv'
            mailboxes = root / 'Exchange_EXO_Mailboxes_AllDomains.csv'
            # SkuId is part of the collector mapping even when it is empty.
            license_rows = [dict(row, SkuId='') for row in self.licenses]
            for path, rows in ((licenses, license_rows), (mailboxes, self.mailboxes)):
                with path.open('w', encoding='utf-8', newline='') as stream:
                    writer = csv.DictWriter(stream, fieldnames=rows[0].keys())
                    writer.writeheader()
                    writer.writerows(rows)
            self.snapshot['Rendering'] = {'SharedSourceReady': True}
            self.snapshot['Sources'] = [dict(Name=p.name, Ready=True, Provisional=False, SHA256=sha256(p)) for p in (licenses, mailboxes)]
            snapshot = root / 'snapshot.json'
            snapshot.write_text(json.dumps(self.snapshot), encoding='utf-8')
            self.assertEqual(len(from_files(snapshot, licenses, mailboxes, 'test')), 2)
            with licenses.open('a', encoding='utf-8') as stream:
                stream.write('\n')
            with self.assertRaisesRegex(ValueError, 'does not match'):
                from_files(snapshot, licenses, mailboxes, 'test')


if __name__ == '__main__':
    unittest.main()
