"""Synthetic mailbox retention tests. No customer files, APIs or Desktop."""
import collections
import datetime as dt
import json
import re
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile
from xml.etree import ElementTree as ET

import workplace_mailbox_exceptions as subject

IDENTITY = dict(TenantKey='test', OrganizationKey='sample', EnvironmentKey='sandbox', TenantId='test-tenant')
AD = '11111111-1111-1111-1111-111111111111'
EXCHANGE = '22222222-2222-2222-2222-222222222222'
OTHER = '33333333-3333-3333-3333-333333333333'
NOW = dt.date(2026, 10, 19)


def local(ad=AD, exchange=EXCHANGE, smtp='room@example.invalid', domain='example.invalid', kind='UserMailbox'):
    return dict(IDENTITY, ObjectGUID=ad, ExchangeGuid=exchange, PrimarySMTPaddress=smtp,
                DomainName=domain, RecipientType=kind, CollectedAtUtc='2026-10-18T08:00:00Z')


def decision(**updates):
    row = dict.fromkeys(subject.HEADERS, '')
    row.update({'Source Row': 8, 'Primary SMTP Address': 'room@example.invalid',
                'Retain On-premises': 'Yes', 'Exchange GUID': EXCHANGE})
    row.update(updates)
    return row


class MatchingTests(unittest.TestCase):
    def run_case(self, rows=None, locals=None, cloud=None, remote=None, bindings=None):
        return subject.resolve(rows if rows is not None else [decision()],
            dict(local_mailboxes=locals if locals is not None else [local()],
                 mailboxes=cloud or [], remote_mailboxes=remote or []), IDENTITY, bindings, NOW)

    def summary(self, inventory, decisions, status='Qualified'):
        return subject.summarize(IDENTITY, inventory, decisions, {'Source Status': status})

    def test_guid_and_all_five_types(self):
        for kind in subject.SCOPE:
            inventory, rows = self.run_case(locals=[local(kind=kind)])
            self.assertEqual(sum(inventory.values()), 1)
            self.assertEqual(rows[0]['Retained'], 1)

    def test_address_only_requires_confirmation(self):
        _, rows = self.run_case([decision(**{'Exchange GUID': ''})])
        self.assertEqual(rows[0]['Retained'], 0)
        self.assertEqual(rows[0]['Requires Review'], 1)
        self.assertEqual(rows[0]['Candidate Identity'], 'local_mailboxes|' + AD)

    def test_confirmed_address_only(self):
        _, rows = self.run_case([decision(**{'Exchange GUID': '', 'Primary SMTP Address': ' ROOM@EXAMPLE.INVALID '})],
                                bindings={'room@example.invalid': 'local_mailboxes|' + AD})
        self.assertEqual(rows[0]['Retained'], 1)

    def test_unknown_or_malformed_pin_never_falls_back(self):
        for raw in [OTHER, 'bad-guid', '00000000-0000-0000-0000-000000000000']:
            _, rows = self.run_case([decision(**{'Exchange GUID': raw})])
            self.assertEqual(rows[0]['Retained'], 0)
            self.assertEqual(rows[0]['Requires Review'], 1)

    def test_two_pins_must_converge(self):
        _, rows = self.run_case([decision(**{'AD Object GUID': OTHER})],
                                locals=[local(), local(ad=OTHER, exchange=OTHER, smtp='other@example.invalid')])
        self.assertEqual(rows[0]['Retained'], 0)

    def test_address_changed_guid_preserves_identity(self):
        _, rows = self.run_case(locals=[local(smtp='new@example.invalid')])
        self.assertEqual(rows[0]['Retained'], 1)
        self.assertIn('Address changed', rows[0]['Status'])

    def test_address_reassigned_never_transfers_decision(self):
        _, rows = self.run_case([decision(**{'Exchange GUID': ''})],
            locals=[local(smtp='new@example.invalid'), local(ad=OTHER, exchange=OTHER)],
            bindings={'room@example.invalid': 'local_mailboxes|' + AD})
        self.assertEqual(rows[0]['Retained'], 0)
        self.assertIn('reassigned', rows[0]['Status'])

    def test_ambiguous_primary_address(self):
        _, rows = self.run_case([decision(**{'Exchange GUID': ''})], locals=[local(), local(ad=OTHER, exchange=OTHER)])
        self.assertEqual(rows[0]['Retained'], 0)
        self.assertIn('Ambiguous', rows[0]['Status'])

    def test_duplicate_input_once_provenance_preserved(self):
        _, rows = self.run_case([decision(), decision(**{'Source Row': 9})])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['Retained'], 1)
        self.assertEqual(rows[0]['Source Rows'], '8,9')

    def test_yes_no_conflict_and_no_retires(self):
        inventory, rows = self.run_case([decision(), decision(**{'Retain On-premises': 'No'})])
        summary = self.summary(inventory, rows)[0]
        self.assertEqual((summary['Retained Count'], summary['Review Count'], summary['Remaining Count']), (0, 1, ''))
        _, rows = self.run_case([decision(**{'Retain On-premises': 'No'})])
        self.assertEqual((rows[0]['Retained'], rows[0]['Requires Review']), (0, 0))

    def test_exo_only_and_remote_only_are_review(self):
        cloud = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        _, rows = self.run_case(locals=[], cloud=[cloud])
        self.assertEqual(rows[0]['Status'], 'Already in Exchange Online')
        remote = dict(IDENTITY, ObjectGuid=AD, ExchangeGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        _, rows = self.run_case(locals=[], remote=[remote])
        self.assertEqual(rows[0]['Retained'], 0)
        self.assertEqual(rows[0]['Requires Review'], 1)

    def test_linked_hybrid_observations_do_not_double_count(self):
        cloud = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        _, rows = self.run_case(cloud=[cloud])
        self.assertEqual(rows[0]['Retained'], 0)
        self.assertEqual(rows[0]['Status'], 'Conflicting hosting observations')

    def test_verified_exo_withdrawal_clears_review_and_global_block(self):
        cloud = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        inventory, rows = self.run_case([decision(**{'Retain On-premises': 'No'})],
            locals=[local(ad=OTHER, exchange=OTHER, smtp='other@example.invalid')], cloud=[cloud])
        self.assertEqual(rows[0]['Status'], 'Withdrawn - verified EXO')
        self.assertEqual((rows[0]['Retained'], rows[0]['Requires Review']), (0, 0))
        self.assertIn('Decision Owner missing', rows[0]['Documentation Warnings'])
        summary = self.summary(inventory, rows)[0]
        self.assertEqual((summary['Retained Count'], summary['Review Count'],
            summary['Global Review Count'], summary['Remaining Count']), (0, 0, 0, 1))

    def test_exo_withdrawal_requires_verified_identity_and_no_decision_conflict(self):
        cloud = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        cases = [
            [decision(**{'Retain On-premises': 'No', 'Exchange GUID': ''})],
            [decision(**{'Retain On-premises': 'No', 'Exchange GUID': 'bad-guid'})],
            [decision(**{'Retain On-premises': 'No', 'Exchange GUID': OTHER})],
            [decision(**{'Retain On-premises': 'Yes'})],
            [decision(**{'Retain On-premises': ''})],
            [decision(**{'Retain On-premises': 'No'}), decision(**{'Source Row': 9})]]
        for case in cases:
            with self.subTest(case=case):
                _, rows = self.run_case(case, locals=[], cloud=[cloud])
                self.assertTrue(all(r['Requires Review'] == 1 and r['Retained'] == 0 for r in rows))

    def test_confirmed_exo_binding_and_remote_native_link_can_withdraw(self):
        cloud = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        _, rows = self.run_case([decision(**{'Retain On-premises': 'No', 'Exchange GUID': ''})],
            locals=[], cloud=[cloud], bindings={'room@example.invalid': 'mailboxes|' + EXCHANGE})
        self.assertEqual(rows[0]['Status'], 'Withdrawn - verified EXO')
        remote = dict(IDENTITY, ObjectGuid=AD, ExchangeGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        _, rows = self.run_case([decision(**{'Retain On-premises': 'No', 'Exchange GUID': '', 'AD Object GUID': AD})],
            locals=[], cloud=[cloud], remote=[remote])
        self.assertEqual((rows[0]['Status'], rows[0]['Retained'], rows[0]['Requires Review']),
            ('Withdrawn - verified EXO', 0, 0))

    def test_exo_withdrawal_never_clears_hosting_smtp_or_pin_conflicts(self):
        cloud = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        no = decision(**{'Retain On-premises': 'No'})
        for locals in ([local()], [local(exchange=OTHER)]):
            with self.subTest(locals=locals):
                _, rows = self.run_case([no], locals=locals, cloud=[cloud])
                self.assertEqual((rows[0]['Retained'], rows[0]['Requires Review']), (0, 1))
        _, rows = self.run_case([decision(**{'Retain On-premises': 'No', 'AD Object GUID': AD})],
            locals=[local(exchange=OTHER, smtp='other@example.invalid')], cloud=[cloud])
        self.assertEqual(rows[0]['Status'], 'Identity conflict')
        _, rows = self.run_case([no], locals=[], cloud=[cloud],
            bindings={'room@example.invalid': 'local_mailboxes|' + AD})
        self.assertEqual(rows[0]['Status'], 'Identity conflict with confirmed binding')

    def test_exo_withdrawal_reassigned_address_and_unresolved_remote_stay_review(self):
        original = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='new@example.invalid')
        reassigned = dict(IDENTITY, MailboxGuid=OTHER, PrimarySmtpAddress='room@example.invalid')
        no = decision(**{'Retain On-premises': 'No'})
        _, rows = self.run_case([no], locals=[], cloud=[original, reassigned])
        self.assertEqual(rows[0]['Status'], 'Address reassigned / identity conflict')
        remote = dict(IDENTITY, ObjectGuid=AD, ExchangeGuid=OTHER, PrimarySmtpAddress='room@example.invalid')
        cloud = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='room@example.invalid')
        _, rows = self.run_case([no], locals=[], cloud=[cloud], remote=[remote])
        self.assertEqual(rows[0]['Requires Review'], 1)
        _, rows = self.run_case([no], locals=[], remote=[remote])
        self.assertEqual(rows[0]['Requires Review'], 1)

    def test_exo_withdrawals_duplicate_once_with_provenance_and_changed_address(self):
        cloud = dict(IDENTITY, MailboxGuid=EXCHANGE, PrimarySmtpAddress='new@example.invalid')
        no = decision(**{'Retain On-premises': 'No', 'Source Reference': 'First source'})
        duplicate = dict(no, **{'Source Row': 9, 'Source Reference': 'Second source'})
        _, rows = self.run_case([no, duplicate], locals=[], cloud=[cloud])
        self.assertEqual(len(rows), 1)
        self.assertEqual((rows[0]['Status'], rows[0]['Retained'], rows[0]['Requires Review']),
            ('Withdrawn - verified EXO', 0, 0))
        self.assertEqual(rows[0]['Source Rows'], '8,9')
        self.assertEqual(rows[0]['Source Reference'], 'First source\nSecond source')

    def test_corrected_aliases_count_native_mailbox_once_and_keep_all_rows(self):
        rows = [decision(**{'Primary SMTP Address': 'old@example.invalid', 'AD Object GUID': AD}),
                decision(**{'Source Row': 9, 'AD Object GUID': AD})]
        _, result = self.run_case(rows)
        self.assertEqual(result[0]['Requires Review'], 1)
        # Normalize the reviewed primary address while retaining both source rows.
        rows[0]['Primary SMTP Address'] = 'room@example.invalid'
        _, result = self.run_case(rows)
        self.assertEqual(len(result), 1)
        self.assertEqual((result[0]['Retained'], result[0]['Requires Review']), (1, 0))
        self.assertEqual(result[0]['Source Rows'], '8,9')

    def test_empty_valid_and_unavailable(self):
        inventory, decisions = self.run_case([])
        row = self.summary(inventory, decisions)[0]
        self.assertEqual((row['Retained Count'], row['Review Count'], row['Remaining Count']), (0, 0, 1))
        row = self.summary(inventory, decisions, 'Not configured')[0]
        self.assertEqual((row['Retained Count'], row['Review Count'], row['Remaining Count']), ('', '', ''))

    def test_documentation_does_not_cancel_yes(self):
        _, rows = self.run_case([decision(**{'Review Date': '2026-01-01'})])
        self.assertEqual(rows[0]['Retained'], 1)
        self.assertIn('Review overdue', rows[0]['Documentation Warnings'])
        self.assertEqual(rows[0]['Reason'], '')

    def test_scope_and_invalid_decision(self):
        for case in [[decision(**{'Retain On-premises': ''})], [decision(**{'Retain On-premises': 'perhaps'})]]:
            _, rows = self.run_case(case)
            self.assertEqual(rows[0]['Requires Review'], 1)
        _, rows = self.run_case(locals=[local(kind='ArbitrationMailbox')])
        self.assertEqual(rows[0]['Status'], 'Outside planning scope')

    def test_foreign_tenant_or_environment_rejected(self):
        for field in ('TenantKey', 'EnvironmentKey'):
            row = local(); row[field] = 'foreign'
            with self.assertRaisesRegex(ValueError, 'identity mismatch'):
                self.run_case(locals=[row])

    def test_domain_filters_do_not_infer_suffix(self):
        inventory, rows = self.run_case(locals=[local(domain='directory.example.invalid')])
        self.assertEqual(rows[0]['Domain Key'], 'directory.example.invalid')
        inventory['other.example.invalid'] = 3
        # A localized review must not block the other domain.
        rows[0]['Requires Review'] = 1; rows[0]['Retained'] = 0
        result = {r['Domain Key']: r for r in self.summary(inventory, rows)}
        self.assertEqual(result['other.example.invalid']['Remaining Count'], 3)
        rows[0]['Domain Key'] = '__unknown__'
        result = {r['Domain Key']: r for r in self.summary(inventory, rows)}
        self.assertEqual(result['other.example.invalid']['Remaining Count'], '')

    def test_negative_net_and_duplicate_inventory_rejected(self):
        _, rows = self.run_case()
        with self.assertRaises(ValueError): self.summary(collections.Counter(), rows)
        with self.assertRaises(ValueError): self.run_case(locals=[local(), local()])

    def test_bindings_are_scoped_and_confirmed(self):
        document = dict(Version=1, Identity=IDENTITY, Bindings=[dict(PrimarySmtpAddress='room@example.invalid',
            NativeIdentity='local_mailboxes|' + AD, ConfirmedBy='Example operator', ConfirmedAtUtc='2026-10-19T08:00:00Z')])
        self.assertEqual(len(subject.validate_bindings(document, IDENTITY)), 1)
        with self.assertRaises(ValueError): subject.validate_bindings(document, dict(IDENTITY, TenantKey='foreign'))
        document['Bindings'][0]['ConfirmedBy'] = ''
        with self.assertRaises(ValueError): subject.validate_bindings(document, IDENTITY)


class ProducerTests(unittest.TestCase):
    def test_qualified_capture_and_disabled_backwards_compatibility(self):
        from workplace_native_sources import CONTRACT, sha
        contract = json.loads(CONTRACT.read_text(encoding='utf-8-sig'))
        definitions = {s['name']: s for s in contract['sources']}
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp); data = root / 'capture'; last = data / 'DATA-LAST'; last.mkdir(parents=True)
            for family in ('local_mailboxes', 'remote_mailboxes', 'mailboxes'):
                records = [local()] if family == 'local_mailboxes' else []
                columns = list(local()) if records else definitions[family]['columns']
                subject.write_csv(last / definitions[family]['file'], columns, records)
            template = Path(__file__).resolve().parents[1] / 'config/templates' / subject.WORKBOOK
            (data / subject.WORKBOOK).write_bytes(template.read_bytes())
            audit = dict(Version=1, Identity=IDENTITY, ContractSHA256=sha(CONTRACT), Evidence={'Files':[
                {'File': definitions['local_mailboxes']['file'], 'CompletedAtUtc':'2026-10-19T00:05:00Z'}]})
            audit_path = root / 'audit.json.txt';audit_path.write_text(json.dumps(audit))
            with patch('workplace_native_build.identity_from_receipts', return_value=IDENTITY), patch('workplace_native_sources.recheck_sources') as recheck:
                result = subject.build(data, root/'prepared', 'test', audit_path, enabled=True,
                    now=dt.datetime(2026,10,19,8,tzinfo=dt.timezone.utc))
                self.assertEqual(recheck.call_count, 2)
            self.assertEqual(result[0]['Snapshot Date'], '2026-10-19')
            self.assertEqual((result[0]['Retained Count'],result[0]['Review Count'],result[0]['Remaining Count']), (0,0,1))
            self.assertEqual(result[0]['Workbook SHA256'], subject.digest(template.read_bytes()))
            self.assertEqual((data / subject.WORKBOOK).read_bytes(),template.read_bytes())
            disabled = subject.build(data,root/'disabled','test')
            self.assertEqual(disabled[0]['Source Status'],'Not configured')
            with self.assertRaisesRegex(ValueError,'overwrite'):subject.build(data,root/'disabled','test')
            with self.assertRaisesRegex(ValueError,'outside'):subject.build(data,data/'bad-output','test')

    def test_bad_native_audit_never_produces_a_measured_zero(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);data=root/'capture';data.mkdir()
            audit=root/'audit.json.txt';audit.write_text(json.dumps({'Identity':IDENTITY}))
            with patch('workplace_native_build.identity_from_receipts',return_value=IDENTITY):
                with self.assertRaisesRegex(ValueError,'identity mismatch'):
                    subject.build(data,root/'prepared','test',audit,enabled=True)
            self.assertFalse((root/'prepared/MailboxMigrationExceptionSummary.csv').exists())


class WorkbookTests(unittest.TestCase):
    def make_register(self, path, headers, values, last_column=None, worksheet_headers=None):
        """Small synthetic XLSX fixture with the production table relationship."""
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        last_column = last_column or chr(64 + len(headers))
        sheet = ET.Element('worksheet', xmlns=ns)
        data = ET.SubElement(sheet, 'sheetData')
        for number, cells in [(7, worksheet_headers or headers), (8, values)]:
            row = ET.SubElement(data, 'row', r=str(number))
            for i, value in enumerate(cells):
                cell = ET.SubElement(row, 'c', r=f'{chr(65+i)}{number}', t='inlineStr')
                ET.SubElement(ET.SubElement(cell, 'is'), 't').text = value
        ET.SubElement(ET.SubElement(sheet, 'tableParts', count='1'), 'tablePart', {f'{{{rel}}}id':'rId1'})
        table = ET.Element('table', xmlns=ns, displayName='MailboxMigrationExceptions', ref=f'A7:{last_column}8')
        columns = ET.SubElement(table, 'tableColumns', count=str(len(headers)))
        for i, header in enumerate(headers):
            ET.SubElement(columns, 'tableColumn', id=str(i+1), name=header)
        with zipfile.ZipFile(path, 'w') as z:
            z.writestr('xl/workbook.xml', f'<workbook xmlns="{ns}" xmlns:r="{rel}"><sheets><sheet name="Exceptions" r:id="rId1"/></sheets></workbook>')
            z.writestr('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
            z.writestr('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships><Relationship Id="rId1" Target="../tables/table1.xml"/></Relationships>')
            z.writestr('xl/worksheets/sheet1.xml', ET.tostring(sheet))
            z.writestr('xl/tables/table1.xml', ET.tostring(table))

    def test_nine_and_ten_columns_preserve_identity_decisions_and_dates(self):
        values = ['room@example.invalid', 'Yes', 'Reviewed reason', 'Operator',
                  '2026-10-19', '', EXCHANGE, AD, 'Room']
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'register.xlsx'
            for headers, cells in [(subject.HEADERS[:-1], values),
                                   (subject.HEADERS, values + [''])]:
                self.make_register(path, headers, cells)
                rows = subject.read_register(path)
                self.assertEqual(rows, [dict(zip(subject.HEADERS, values + ['']), **{'Source Row':8})])
                _, resolved = subject.resolve(rows, dict(local_mailboxes=[local()], mailboxes=[], remote_mailboxes=[]), IDENTITY, as_of=NOW)
                self.assertEqual(resolved[0]['Status'], 'Matched local')
            self.make_register(path, subject.HEADERS, values + ['Legacy provenance'])
            self.assertEqual(subject.read_register(path)[0]['Source Reference'], 'Legacy provenance')

    def test_nine_column_reader_keeps_strict_schema_and_range_guards(self):
        headers = subject.HEADERS[:-1]
        values = ['room@example.invalid', 'Yes', '', '', '', '', EXCHANGE, AD, 'Room']
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'register.xlsx'
            cases = [
                (headers, 'J', None),
                (subject.HEADERS, 'I', None),
                (headers[:-1], 'H', None),
                (headers[:-1] + ['Unexpected'], 'I', None),
                (headers[:6] + [headers[7], headers[6], headers[8]], 'I', None),
                (headers, 'I', headers[:-1] + ['Wrong worksheet header']),
            ]
            for columns, last_column, sheet_headers in cases:
                with self.subTest(columns=columns, last_column=last_column, sheet_headers=sheet_headers):
                    self.make_register(path, columns, values, last_column, sheet_headers)
                    with self.assertRaises(ValueError):
                        subject.read_register(path)

    def test_neutral_template_is_valid_and_empty(self):
        template = Path(__file__).resolve().parents[1] / 'config/templates' / subject.WORKBOOK
        self.assertEqual(subject.read_register(template), [])
        with zipfile.ZipFile(template) as z:
            table = ET.fromstring(z.read('xl/tables/table1.xml'))
            self.assertEqual(table.get('ref'), 'A7:I8')
            self.assertEqual([c.get('name') for c in table.findall('.//{*}tableColumn')], subject.HEADERS[:-1])
            text = ''.join(z.read(n).decode() for n in z.namelist() if n.endswith('.xml'))
            self.assertNotIn('Source Reference', text)
            self.assertIsNone(re.search(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+', text))
            self.assertNotIn('SourceData', text)

    def test_absent_corrupt_and_wrong_schema_fail(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'bad.xlsx'
            with self.assertRaises(FileNotFoundError): subject.read_register(path)
            path.write_text('not a workbook')
            with self.assertRaises(zipfile.BadZipFile): subject.read_register(path)
            with zipfile.ZipFile(path, 'w') as z:
                z.writestr('xl/workbook.xml', '<workbook><sheets/></workbook>')
                z.writestr('xl/_rels/workbook.xml.rels', '<Relationships/>')
            with self.assertRaises(ValueError): subject.read_register(path)


if __name__ == '__main__':
    unittest.main()
