"""Offline weekly application tolerance and strict output qualification tests."""
import csv
import json
import unittest
import test_cmdb_preparation as fixtures

pipeline = fixtures.pipeline


class ApplicationCoverageTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.PreparationTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)

    def update(self):
        self.fixture.write_inputs()

    def assess(self):
        return pipeline.assess_application_coverage(self.fixture.source, self.fixture.contract)

    def test_complete_baseline_has_no_warning_and_resolved_links(self):
        coverage, warnings = self.assess()
        self.assertEqual(coverage['RelationRows'], 2)
        self.assertFalse(warnings)
        self.fixture.prepare()
        self.assertTrue(all(r['DeviceLinkStatus'] == r['ApplicationLinkStatus'] == 'Resolved'
                            for r in self.fixture.table('FactDeviceApplication')))
        self.assertTrue(all(r['RelationshipCoverageStatus'] == 'Complete'
                            for r in self.fixture.table('DimDetectedApplication')))

    def test_validate_only_reports_all_orphans_without_output_or_source_changes(self):
        self.fixture.inputs['app_relations'] = [
            {'AppId': 'a1', 'DeviceId': 'missing'}, {'AppId': 'a2', 'DeviceId': 'missing'},
            {'AppId': 'outside', 'DeviceId': 'other'}]
        self.update()
        before = {p.name: pipeline.sha(p) for p in self.fixture.source.iterdir()}
        result = self.fixture.prepare(validate_only=True)
        coverage = result['ApplicationCoverage']
        self.assertEqual(coverage['UnresolvedDeviceRelationRows'], 3)
        self.assertEqual(coverage['DistinctUnresolvedDeviceIds'], 2)
        self.assertEqual(coverage['UnresolvedApplicationRelationRows'], 1)
        self.assertEqual(coverage['DistinctUnresolvedApplicationIds'], 1)
        self.assertEqual(len(result['ApplicationWarnings']), 2)
        self.assertEqual(before, {p.name: pipeline.sha(p) for p in self.fixture.source.iterdir()})
        self.assertFalse(self.fixture.output.exists())
        self.assertFalse(list(self.fixture.root.glob('.cmdb-*')))

    def test_all_relations_retained_without_fabricated_parent_or_current_device_count(self):
        self.fixture.inputs['app_relations'].append({'AppId': 'outside', 'DeviceId': 'outside'})
        self.fixture.inputs['app_relations'][0]['DeviceId'] = 'missing'
        self.update()
        self.fixture.prepare()
        relations = self.fixture.table('FactDeviceApplication')
        self.assertEqual(len(relations), 3)
        outside = next(r for r in relations if r['AppId'] == 'outside')
        self.assertEqual(outside['TenantApplicationKey'], '')
        self.assertEqual(outside['ApplicationLinkStatus'], 'Unresolved')
        self.assertEqual(outside['DeviceLinkStatus'], 'Unresolved')
        self.assertEqual(len(self.fixture.table('DimIntuneManagedDevice')), 2)
        self.assertEqual(len(self.fixture.table('DimDevice')), 2)
        self.assertEqual(len(self.fixture.table('DimDetectedApplication')), 2)
        self.assertEqual(self.fixture.table('TopApplication')[0]['ReportedDeviceCount'], '2')
        app = next(r for r in self.fixture.table('DimDetectedApplication') if r['AppId'] == 'a1')
        self.assertEqual(app['ResolvedDeviceCount'], '0')
        self.assertEqual(app['UnresolvedDeviceCount'], '1')
        self.assertEqual(app['RelationshipCoverageStatus'], 'Device links unresolved')
        health = next(r for r in self.fixture.table('SourceHealth')
                      if r['SourceName'] == 'Intune_DiscoveredApps_AppDeviceRelations.csv')
        self.assertEqual(health['Status'], 'SuccessWithWarnings')
        self.assertIn('Weekly application evidence', health['Coverage'])
        self.assertIn('acquisition age', health['Evidence'])
        manifest = pipeline.load_json(self.fixture.output / pipeline.MANIFEST)
        self.assertEqual(manifest['PreparationQualifications']['ApplicationCoverage'],
                         manifest['SourceEvidence']['ApplicationCoverage'])
        findings = [r for r in self.fixture.table('FactDataQuality')
                    if r['FindingType'] == 'ApplicationCoverageWarning']
        self.assertEqual(len(findings), 2)
        self.assertTrue(all(r['Severity'] == 'Information' for r in findings))

    def test_mixed_device_resolution_and_count_mismatch_have_separate_counts(self):
        self.fixture.inputs['app_relations'].append({'AppId': 'a1', 'DeviceId': 'missing'})
        self.fixture.inputs['apps'][0]['DeviceCount'] = '7'
        self.update()
        result = self.fixture.prepare()
        app = next(r for r in self.fixture.table('DimDetectedApplication') if r['AppId'] == 'a1')
        self.assertEqual(app['DeviceCount'], '2')
        self.assertEqual(app['ReportedDeviceCount'], '7')
        self.assertEqual(app['ResolvedDeviceCount'], '1')
        self.assertEqual(app['UnresolvedDeviceCount'], '1')
        self.assertEqual(app['RelationshipCoverageStatus'], 'Relation count differs; Device links unresolved')
        self.assertEqual(result['ApplicationCoverage']['CountMismatchApplications'], 1)

    def test_zero_count_and_no_relations_are_complete(self):
        self.fixture.inputs['app_relations'] = []
        for row in self.fixture.inputs['apps']:
            row['DeviceCount'] = '0'
        self.update()
        self.assertFalse(self.assess()[1])
        self.fixture.prepare()
        self.assertFalse(self.fixture.table('FactDeviceApplication'))
        self.assertTrue(all(r['RelationshipCoverageStatus'] == 'Complete'
                            for r in self.fixture.table('DimDetectedApplication')))

    def test_native_keys_are_case_and_space_normalized(self):
        self.fixture.inputs['app_relations'][0] = {'AppId': ' A1 ', 'DeviceId': ' MD1 '}
        self.update()
        self.assertFalse(self.assess()[1])
        self.fixture.prepare()
        self.assertTrue(all(r['DeviceLinkStatus'] == 'Resolved' for r in self.fixture.table('FactDeviceApplication')))

    def test_unknown_device_can_resolve_on_next_inventory_without_losing_relation(self):
        self.fixture.inputs['app_relations'][0]['DeviceId'] = 'missing'
        self.update()
        self.fixture.prepare()
        first = next(r for r in self.fixture.table('FactDeviceApplication') if r['AppId'] == 'a1')
        device = dict(self.fixture.inputs['managed'][0], ManagedDeviceId='missing', AzureADDeviceId='ed3')
        hardware = dict(self.fixture.inputs['hardware'][0], ManagedDeviceId='missing', azureADDeviceId='ed3')
        self.fixture.inputs['managed'].append(device)
        self.fixture.inputs['hardware'].append(hardware)
        self.update()
        self.fixture.prepare()
        second = next(r for r in self.fixture.table('FactDeviceApplication') if r['AppId'] == 'a1')
        self.assertEqual(second['TenantDeviceApplicationKey'], first['TenantDeviceApplicationKey'])
        self.assertEqual(second['DeviceLinkStatus'], 'Resolved')

    def test_invalid_reported_count_still_rejects_validation(self):
        for bad in ('', '-1', 'NaN', '1.5', 'unknown'):
            with self.subTest(value=bad):
                self.fixture.inputs['apps'][0]['DeviceCount'] = bad
                self.update()
                with self.assertRaisesRegex(ValueError, 'Invalid application reported device count'):
                    self.fixture.prepare(validate_only=True)

    def test_foreign_tenant_blank_keys_duplicates_and_wrong_hash_still_reject(self):
        for mutation, error in (
                (lambda r: r.update(TenantKey='foreign'), 'TenantKey'),
                (lambda r: r.update(DeviceId=''), 'Blank immutable'),
                (lambda r: self.fixture.inputs['app_relations'].append(dict(r)), 'Duplicate immutable')):
            self.fixture.inputs['app_relations'] = [{'AppId': 'a1', 'DeviceId': 'md1'}]
            mutation(self.fixture.inputs['app_relations'][0]); self.update()
            if error == 'TenantKey':
                # The standard fixture stamps its tenant; change the native file
                # explicitly and bind that byte change to its synthetic receipt.
                path = self.fixture.source / 'Intune_DiscoveredApps_AppDeviceRelations.csv'
                path.write_text(path.read_text(encoding='utf-8-sig').replace('synthetic,', 'foreign,'),
                                encoding='utf-8-sig')
                record = next(r for r in self.fixture.proof['Files'] if r['File'] == path.name)
                record['SHA256'] = pipeline.sha(path)
                self.fixture.write_proof()
            with self.subTest(error=error), self.assertRaisesRegex(ValueError, error):
                self.fixture.prepare(validate_only=True)
        self.fixture.inputs['app_relations'] = [{'AppId': 'a1', 'DeviceId': 'md1'}]
        self.update()
        with (self.fixture.source / 'Intune_DiscoveredApps_AppDeviceRelations.csv').open('a') as stream:
            stream.write('changed\n')
        with self.assertRaisesRegex(ValueError, 'hash mismatch'):
            self.fixture.prepare(validate_only=True)

    def tamper_output(self, table, mutate):
        path = self.fixture.output / (table + '.csv')
        data = self.fixture.table(table)
        mutate(data)
        columns = next(d['columns'] for d in self.fixture.contract['tables'] if d['name'] == table)
        with path.open('w', encoding='utf-8-sig', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=columns)
            writer.writeheader(); writer.writerows(data)
        manifest_path = self.fixture.output / pipeline.MANIFEST
        manifest = pipeline.load_json(manifest_path)
        manifest['OutputFiles'][path.name].update(SHA256=pipeline.sha(path), Rows=len(data))
        manifest_path.write_text(json.dumps(manifest), encoding='utf-8')

    def test_output_cannot_hide_missing_device_as_resolved(self):
        self.fixture.inputs['app_relations'][0]['DeviceId'] = 'missing'
        self.update(); self.fixture.prepare()
        self.tamper_output('FactDeviceApplication', lambda data: data[0].update(DeviceLinkStatus='Resolved'))
        with self.assertRaisesRegex(ValueError, 'device link qualification'):
            pipeline.validate_current(self.fixture.output, self.fixture.contract, 'synthetic')

    def test_output_cannot_mark_existing_device_as_unresolved(self):
        self.fixture.prepare()
        self.tamper_output('FactDeviceApplication', lambda data: data[0].update(DeviceLinkStatus='Unresolved'))
        with self.assertRaisesRegex(ValueError, 'device link qualification'):
            pipeline.validate_current(self.fixture.output, self.fixture.contract, 'synthetic')

    def test_output_cannot_fabricate_application_parent(self):
        self.fixture.inputs['app_relations'][0]['AppId'] = 'outside'
        self.update(); self.fixture.prepare()
        other = self.fixture.table('DimDetectedApplication')[0]['TenantApplicationKey']
        self.tamper_output('FactDeviceApplication', lambda data: data[0].update(TenantApplicationKey=other))
        with self.assertRaisesRegex(ValueError, 'application link qualification'):
            pipeline.validate_current(self.fixture.output, self.fixture.contract, 'synthetic')

    def test_output_cannot_claim_complete_coverage_after_missing_devices(self):
        self.fixture.inputs['app_relations'][0]['DeviceId'] = 'missing'
        self.update(); self.fixture.prepare()
        self.tamper_output('DimDetectedApplication', lambda data: data[0].update(RelationshipCoverageStatus='Complete'))
        with self.assertRaisesRegex(ValueError, 'relationship coverage'):
            pipeline.validate_current(self.fixture.output, self.fixture.contract, 'synthetic')


if __name__ == '__main__':
    unittest.main(verbosity=2)
