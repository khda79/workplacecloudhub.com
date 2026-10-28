"""Synthetic offline checks; no client fixtures or production preparation."""
import copy
import csv
import datetime as dt
import json
from pathlib import Path
import tempfile
import unittest
from types import SimpleNamespace
import workplace_spo_migration as m

NOW = dt.datetime(2026, 10, 19, 12, tzinfo=dt.timezone.utc)
PARAMS = {'TenantKey': 'sample-prod', 'ProjectKey': 'sample-project', 'ReviewDate': '2026-10-19'}


def web(path='', ident='web-root', collection='collection-a'):
    return dict(TenantKey='sample-prod', FarmId='farm-a', SiteCollectionId=collection,
                WebId=ident, WebApplicationId='app-a', RunId='run-a', CollectedAtUtc=NOW.isoformat(),
                Url='https://portal.example.invalid' + path, Title=ident)


def operation(path='', **overrides):
    return dict(MigrationID='MIG-000001', SourceURL='https://portal.example.invalid' + path,
                SourceName='Sample root', Disposition='MIGRATE', ExecutionStatus='DONE',
                ValidationStatus='NOT_RECORDED', TargetURL='https://sample.sharepoint.com/sites/sample',
                TargetReviewStatus='CONFIRMED', SourceObjectType='WEB', SourceScope='SELF',
                ActualDate='2026-10-12', _row=2, **overrides)


def tables(rows=None, shared=None):
    return {'MigrationSP': rows if rows is not None else [operation()], 'MigrationSharedFolders': shared or [],
            'LegacyWebs': [], 'LegacyCollections': []}


def resolve(data=None, native=None, **kwargs):
    return m.resolve(PARAMS, data or tables(), 'sample-prod', native if native is not None else [web()], [],
                     [{'Source': 'Content', 'Source SHA256': 'sample-sha', 'Qualification': 'Observed only'}], NOW, **kwargs)


class MigrationTests(unittest.TestCase):
    def test_business_url_protocol_and_default_ports(self):
        self.assertEqual(m.url_key('http://PORTAL.example.invalid:80/x/'), m.url_key('https://portal.example.invalid:443/X'))
        self.assertNotEqual(m.url_key('http://portal.example.invalid:443/x'), m.url_key('https://portal.example.invalid/x'))
        self.assertNotEqual(m.url_key('https://portal.example.invalid/item'), m.url_key('https://portal.example.invalid/items'))
        self.assertNotEqual(m.url_key('https://portal.example.invalid/a%2Fb'), m.url_key('https://portal.example.invalid/a/b'))

    def test_protocol_match_does_not_change_raw_url(self):
        row = operation(); row['SourceURL'] = 'http://portal.example.invalid/'
        result = resolve(tables([row]))
        self.assertEqual(result['Plan'][0]['Source Qualification'], 'Direct native match')
        self.assertEqual(result['Plan'][0]['Source Reference'], row['SourceURL'])

    def test_no_implicit_domain_replacement(self):
        row = operation(); row['SourceURL'] = 'https://old.example.invalid'
        self.assertEqual(resolve(tables([row]))['Plan'][0]['Source Qualification'], 'Not observed')

    def test_approved_alias_preserves_multiple_native_keys(self):
        row = operation()
        old = web(ident='web-old'); old['Url'] = 'http://old.example.invalid'
        pairs = {(row['MigrationID'], m.url_key(row['SourceURL'])): {m.url_key(old['Url'])}}
        result = resolve(tables([row]), [web(), old], aliases=pairs)
        self.assertEqual(result['Plan'][0]['Native Web Count'], 2)
        self.assertEqual(len(result['Web']), 2)

    def test_cross_flow_ids_are_distinct(self):
        shared = operation(); shared['SourcePath'] = r'\\files.example.invalid\share\folder'
        result = resolve(tables(shared=[shared]))
        self.assertEqual(len({r['Operation Key'] for r in result['Plan']}), 2)
        self.assertEqual(result['Plan'][1]['Source Qualification'], 'Declared path only')

    def test_same_flow_duplicate_is_fatal(self):
        with self.assertRaisesRegex(ValueError, 'Duplicate'):
            resolve(tables([operation(), operation()]))

    def test_same_source_two_destinations_are_operations(self):
        first, second = operation(), operation()
        second['MigrationID'] = 'MIG-000002'; second['TargetURL'] += '-other'
        result = resolve(tables([first, second]))
        self.assertEqual(len(result['Plan']), 2)
        self.assertEqual(len(result['Web']), 1)

    def test_same_native_id_different_collections(self):
        other = web('/other', collection='collection-b')
        self.assertEqual(len(resolve(native=[web(), other])['Web']), 2)

    def test_native_identity_duplicates_are_fatal(self):
        with self.assertRaisesRegex(ValueError, 'Duplicate native'):
            resolve(native=[web(), web()])

    def test_child_override_and_collection_boundary(self):
        root, child = operation(), operation('/child')
        root['SourceScope'] = 'INCLUDE_CHILDREN'
        child['MigrationID'] = 'MIG-000002'; child['Disposition'] = 'DELETE'
        result = resolve(tables([root, child]), [web(), web('/child', 'web-child'), web('/sibling', 'web-sibling'), web('/other', 'web-other', 'collection-b')])
        by_key = {r['Web Key'].split('|')[-1]: r for r in result['Web']}
        self.assertEqual(by_key['web-other']['Coverage Kind'], 'NOTINPLAN')
        effective = [r for r in result['Scope'] if r['Web Key'].endswith('|web-child') and r['Is Effective']]
        self.assertEqual([r['Effective Disposition'] for r in effective], ['DELETE'])

    def test_partial_content_does_not_cover_web(self):
        row = operation('/library'); row['SourceObjectType'] = 'LIBRARY'; row['SourceScope'] = 'INCLUDE_CHILDREN'
        result = resolve(tables([row]), [web(), web('/child', 'web-child')])
        self.assertEqual(result['Plan'][0]['Source Qualification'], 'Parent context only')
        self.assertFalse(any(r['Is Effective'] for r in result['Scope']))

    def test_probable_content_types_remain_partial(self):
        for kind in ['LIBRARY_PROBABLE','FOLDER_PROBABLE']:
            row=operation('/content');row['SourceObjectType']=kind;row['SourceScope']='INCLUDE_CHILDREN'
            result=resolve(tables([row]),[web(),web('/content','web-content')])
            self.assertFalse(any(r['Is Effective'] for r in result['Scope']))
            self.assertTrue(all(r['Coverage Kind']=='PartialContent' for r in result['Scope']))

    def test_unobserved_delete_preserves_unknown_date(self):
        row = operation('/absent'); row['Disposition'] = 'DELETE'; row['ActualDate'] = ''
        result = resolve(tables([row]))
        op = result['Plan'][0]
        self.assertEqual((op['Disposition'], op['Execution Status'], op['Actual Date']), ('DELETE', 'DONE', ''))
        self.assertTrue(any(f['Finding Rule'] == 'Source identity not qualified' for f in result['Finding']))

    def test_observed_delete_is_review_not_rewrite(self):
        row = operation(); row['Disposition'] = 'DELETE'
        result = resolve(tables([row]))
        self.assertTrue(any(f['Finding Rule'] == 'Declared deletion still observed' for f in result['Finding']))
        self.assertEqual(result['Plan'][0]['Actual Date'], '2026-10-12')

    def test_validation_needs_coherent_date(self):
        row = operation(); row['ValidationStatus'] = 'VALIDATED'
        for invalid in ['', '2026-10-10', '2026-10-20', 'ambiguous']:
            row['ValidationDate'] = invalid
            self.assertFalse(resolve(tables([row]))['Plan'][0]['Is Validated'])
        row['ValidationDate'] = '2026-10-13'
        self.assertTrue(resolve(tables([row]))['Plan'][0]['Is Validated'])

    def test_relative_target_resolution(self):
        row = operation(); row['TargetURL'] = '/sites/sample'
        result = resolve(tables([row]), target_tenant_url='https://sample.sharepoint.com')
        self.assertEqual(result['Plan'][0]['Resolved Target URL'], 'https://sample.sharepoint.com/sites/sample')
        self.assertEqual(result['Plan'][0]['Target URL'], '/sites/sample')

    def test_external_target_not_qualified(self):
        row = operation(); row['TargetURL'] = 'https://other.example.invalid/sites/sample'
        self.assertEqual(resolve(tables([row]))['Plan'][0]['Resolved Target URL'], '')

    def test_parent_approval_does_not_promote_raw_status(self):
        root, child = operation(), operation('/child')
        child.update(MigrationID='MIG-000002', TargetReviewStatus='NOT_RECORDED', TargetObjectType='SITE')
        key = m.token('sample-prod', 'sample-project', 'SHAREPOINT-ONPREM', child['MigrationID'])
        proof = {key: {'SourceReference': child['SourceURL'], 'TargetURL': child['TargetURL'], 'ParentMigrationID': root['MigrationID']}}
        result = resolve(tables([root, child]), [web(), web('/child', 'web-child')], target_bindings=proof)
        op = result['Plan'][1]
        self.assertEqual(op['Target Approval Evidence'], 'Approved parent mapping')
        self.assertEqual(op['Target Review Status'], 'NOT_RECORDED')
        self.assertFalse(op['Is Validated'])
        proof[key]['TargetURL'] += '/changed'
        self.assertEqual(resolve(tables([root, child]), target_bindings=proof)['Plan'][1]['Target Approval Evidence'], 'Not recorded')

    def test_history_preserves_attributes_and_disappearance(self):
        first = resolve()['History']
        second = copy.deepcopy(first)
        second[0].update({'Snapshot ID': 'second', 'History Key': 'new-key', 'Snapshot Date': '2026-10-20', 'Business Unit': 'New unit'})
        merged = m.merge_history(second, first, 'sample-prod', 'sample-project', 'second', '2026-10-20')
        self.assertEqual(len(merged), 2)
        self.assertEqual(merged[0]['Business Unit'], 'Unknown')
        gone = m.merge_history([], merged, 'sample-prod', 'sample-project', 'third', '2026-10-21')
        self.assertFalse(gone[-1]['Present In Register'])
        again = m.merge_history(second, merged, 'sample-prod', 'sample-project', 'second', '2026-10-20')
        self.assertEqual(len(again), 2)

    def test_history_rejects_cross_tenant(self):
        prior = resolve()['History']; prior[0]['TenantKey'] = 'other'
        with self.assertRaises(ValueError):
            m.merge_history([], prior, 'sample-prod', 'sample-project', 'new', '2026-10-20')

    def test_empty_register_snapshot_retains_the_latest_observation_time(self):
        first = resolve()['History']
        afternoon = '2026-10-19T23:00:00+00:00'
        gone = m.merge_history([], first, 'sample-prod', 'sample-project', 'removed', '2026-10-19', afternoon)
        latest = max(gone, key=lambda row: (row['Capture UTC'], row['Snapshot ID']))
        self.assertEqual(latest['Snapshot ID'], 'removed')
        self.assertFalse(latest['Present In Register'])
        self.assertEqual(latest['Business Unit'], first[0]['Business Unit'])
        self.assertEqual(latest['Actual Date'], first[0]['Actual Date'])
        self.assertEqual(first[0]['Snapshot ID'], resolve()['History'][0]['Snapshot ID'])

    def test_template_has_no_operations(self):
        p = Path(__file__).resolve().parents[1] / 'config/templates/SmartWorkplaceIntelligence-Migration-to-SPO.xlsx'
        parameters, data = m.read_workbook(p)
        self.assertFalse(any(data.values()))
        self.assertFalse(parameters['TenantKey'])
        self.assertFalse(parameters['ProjectKey'])

    def test_approved_host_mapping_is_exact_path_and_scoped(self):
        aliases=m.ApprovedAliases({'old.example.invalid':'portal.example.invalid'})
        row=operation('/child');row['SourceURL']='http://old.example.invalid/child'
        self.assertEqual(resolve(tables([row]),[web('/child','web-child')],aliases=aliases)['Plan'][0]['Source Qualification'],'Direct native match')
        self.assertEqual(resolve(tables([row]),[web('/children','web-child')],aliases=aliases)['Plan'][0]['Source Qualification'],'Not observed')

    def test_exact_cross_host_target_edge_is_not_a_source_alias(self):
        parent,child=operation('/branch'),operation('/branch/child')
        child.update(MigrationID='MIG-000002',SourceURL='https://other.example.invalid/branch/child',TargetReviewStatus='NOT_RECORDED')
        key=m.token('sample-prod','sample-project','SHAREPOINT-ONPREM',child['MigrationID'])
        proof={key:{'SourceReference':child['SourceURL'],'ParentSourceReference':parent['SourceURL'],'TargetURL':child['TargetURL'],'ParentMigrationID':parent['MigrationID']}}
        result=resolve(tables([parent,child]),target_bindings=proof)
        self.assertEqual(result['Plan'][1]['Target Approval Evidence'],'Approved parent mapping')
        self.assertEqual(result['Plan'][1]['Source Qualification'],'Not observed')

    def test_disabled_outputs_are_typed_and_not_configured(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            args = SimpleNamespace(enabled=False, tenant='sample-prod', as_of=NOW.isoformat(), output_root=str(root), previous_history='')
            m.build(args)
            contract = json.loads(m.CONTRACT.read_text())
            self.assertEqual(len(list(root.glob('*.csv'))), 9)
            self.assertEqual(m.read_csv(root/'SharePointMigrationSourceSummary.csv')[0]['Qualification'], 'Not configured')
            for spec in contract['tables']:
                with (root/spec['file']).open(encoding='utf-8-sig') as stream:
                    self.assertEqual(next(m.csv.reader(stream)), [c['name'] for c in spec['columns']])
                # The batch validator rejects a blank key part, so a disabled feature must still publish valid keys.
                for row in m.read_csv(root/spec['file']):
                    self.assertTrue(all(row[k].strip() for k in spec['uniqueKey']), spec['file'])


class InventoryReceiptTests(unittest.TestCase):
    producer = 'SmartM365-SharePoint-OnPrem-Content-Inventory'

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.last = Path(self.temp.name)
        self.receipt_path = self.last / ('SmartInventory_' + self.producer + '.current.json.txt')

    def write(self, groups, entry_changes=None):
        entries = []
        for index, rows in enumerate(groups):
            path = self.last / f'SharePoint_OnPrem_Sample{index}.csv'
            with path.open('w', encoding='utf-8-sig', newline='') as stream:
                writer = csv.DictWriter(stream, fieldnames=['TenantKey', 'RunId'], delimiter=';')
                writer.writeheader()
                writer.writerows(rows)
            entry = {'File': path.name, 'Rows': len(rows), 'SHA256': m.sha(path), 'RunId': 'publication-run', 'Status': 'Success'}
            entry.update(entry_changes or {})
            entries.append(entry)
        self.receipt = {'Owner': 'SmartInventory-SourceReceipt', 'ContractVersion': '1.2', 'TenantKey': 'sample-prod',
                        'Producer': self.producer + '.ps1', 'Status': 'Completed', 'RunId': 'publication-run',
                        'CompletedAtUtc': NOW.isoformat(), 'Files': entries, 'RequiredFiles': [e['File'] for e in entries]}
        self.receipt_path.write_text(json.dumps(self.receipt))

    def check(self):
        return m.inventory(self.last, self.producer, 'sample-prod', NOW)

    def test_distinct_publication_and_collector_runs_are_valid(self):
        rows = [{'TenantKey': 'sample-prod', 'RunId': 'collector-run'}]
        self.write([rows, rows, []])
        files, state = self.check()
        self.assertEqual(len(files), 3)
        self.assertEqual(state['Run ID'], 'publication-run')
        self.assertEqual(state['Qualification'], 'Observed only')

    def test_rejects_mixed_csv_runs_across_files(self):
        self.write([[{'TenantKey': 'sample-prod', 'RunId': run}] for run in ['collector-a', 'collector-b']])
        with self.assertRaisesRegex(ValueError, 'Mixed SharePoint CSV inventory RunId.*Sample1'):
            self.check()

    def test_rejects_mixed_csv_runs_within_a_file(self):
        self.write([[{'TenantKey': 'sample-prod', 'RunId': run} for run in ['collector-a', 'collector-b']]])
        with self.assertRaisesRegex(ValueError, 'Mixed SharePoint CSV inventory RunId.*row 3'):
            self.check()

    def test_rejects_cross_tenant_rows_and_missing_inventory_run(self):
        for row, expected in [({'TenantKey': 'other', 'RunId': 'collector-run'}, 'CSV tenant mismatch'),
                              ({'TenantKey': 'sample-prod', 'RunId': ''}, 'Missing SharePoint CSV inventory RunId')]:
            with self.subTest(row=row):
                self.write([[row]])
                with self.assertRaisesRegex(ValueError, expected):
                    self.check()

    def test_receipt_entries_must_match_the_publication(self):
        for change in [{'RunId': 'other-publication'}, {'RunId': ''}, {'Status': 'Failed'}]:
            with self.subTest(change=change):
                self.write([[{'TenantKey': 'sample-prod', 'RunId': 'collector-run'}]], change)
                with self.assertRaisesRegex(ValueError, 'receipt entry publication identity or status'):
                    self.check()

    def test_still_rejects_tampered_payload_and_wrong_row_count(self):
        row = [{'TenantKey': 'sample-prod', 'RunId': 'collector-run'}]
        self.write([row])
        with (self.last / self.receipt['Files'][0]['File']).open('a') as stream:
            stream.write('tamper')
        with self.assertRaisesRegex(ValueError, 'payload mismatch'):
            self.check()
        self.write([row], {'Rows': 2})
        with self.assertRaisesRegex(ValueError, 'row count mismatch'):
            self.check()


if __name__ == '__main__':
    unittest.main()
