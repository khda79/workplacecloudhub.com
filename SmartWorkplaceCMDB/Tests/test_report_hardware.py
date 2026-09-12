"""Stable V1 hardware report adapter tests. Synthetic data, no Desktop or tenant."""
import csv
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'PowerBI'))
import report_hardware as hw


class HardwareReport(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='cmdb-report-hardware-')
        self.root = Path(self.temp.name)
        self.identity = dict(TenantKey='example-test', OrganizationKey='example', EnvironmentKey='test', TenantId='')
        self.dim = self.root / 'DimDevice.csv'
        self.device = dict(self.identity, CmdbDeviceId='example-test|device|1', TenantDeviceKey='example-test|key|1', SourceDeviceId='azure-1')
        self.write(self.dim, [self.device])

    def tearDown(self):
        self.temp.cleanup()

    def write(self, path, rows, columns=None):
        with path.open('w', encoding='utf-8', newline='') as stream:
            writer = csv.DictWriter(stream, columns or list(rows[0])); writer.writeheader(); writer.writerows(rows)

    def fixture(self):
        self.raw_path = self.root / 'Intune_DeviceHardware.csv'
        self.inv_path = self.root / 'Intune_ManagedDevices.csv'
        self.ci_path = self.root / 'CMDB_CIDeviceHardware.csv'
        self.raw = dict(self.identity, SourceSystem='MicrosoftIntune', ManagedDeviceId='native-1', AzureAdDeviceId='azure-1',
            SerialNumber='SYNTHETIC', SerialNumberStatus='Reported', Manufacturer='', ManufacturerStatus='Missing',
            Model='Example', ModelStatus='Reported', TotalStorageSpaceInBytes='512110190592', StorageStatus='Reported', SourceCollectedDateTime='2026-01-02T00:00:00Z')
        self.inv = dict(self.identity, ManagedDeviceId='native-1', SourceCollectedDateTime='2026-01-01T00:00:00Z')
        self.ci = dict(self.raw, CI_ID=self.device['CmdbDeviceId'], HardwareCollectedDateTime=self.raw['SourceCollectedDateTime'], InventoryCollectedDateTime=self.inv['SourceCollectedDateTime'], CollectionCoverage='Bounded', CollectionMode='Fixture')
        self.ci.pop('SourceCollectedDateTime')
        self.save_fixture()

    def save_fixture(self):
        self.write(self.raw_path, [self.raw]); self.write(self.inv_path, [self.inv])
        cols = hw.load_json(hw.PRODUCT / 'Schema/SmartWorkplaceCMDB.ci.hardware.json')['columns']
        self.write(self.ci_path, [self.ci], cols)
        manifest = dict(Channel='stable', Status='Exported', Hardware=dict(Status='Validated', RowCount=1, Coverage='Bounded', Mode='Fixture', InputHashes={str(self.raw_path): hw.sha(self.raw_path)}), SourceEvidence=dict(InputHashes={str(self.inv_path): hw.sha(self.inv_path)}))
        (self.root / 'CIRegistry.manifest.json').write_text(json.dumps(manifest))

    def test_missing_source_has_no_fake_rows(self):
        identity, rows, coverage, _ = hw.prepare_data(self.dim)
        self.assertEqual(identity, self.identity); self.assertEqual(rows, [])
        self.assertEqual(coverage['Status'], 'NotProvided')

    def test_maps_ci_and_separate_dates_and_missing_values(self):
        self.fixture()
        _, rows, coverage, _ = hw.prepare_data(self.dim, self.ci_path)
        self.assertEqual(rows[0]['TenantDeviceKey'], self.device['TenantDeviceKey'])
        self.assertEqual(rows[0]['Manufacturer'], 'Not provided')
        self.assertEqual(rows[0]['Storage'], '476.94 GiB')
        self.assertNotEqual(rows[0]['HardwareCollectedDateTime'], rows[0]['InventoryCollectedDateTime'])
        self.assertEqual(coverage['Coverage'], 'Bounded')

    def test_zero_and_missing_capacity(self):
        self.fixture()
        for value, status, expected in [('0', 'ZeroReported', 'Unknown (reported 0)'), ('', 'Missing', 'Not provided')]:
            self.raw.update(TotalStorageSpaceInBytes=value, StorageStatus=status)
            self.ci.update(TotalStorageSpaceInBytes=value, StorageStatus=status)
            self.save_fixture()
            self.assertEqual(hw.prepare_data(self.dim, self.ci_path)[1][0]['Storage'], expected)

    def test_modified_raw_hash_rejected(self):
        self.fixture(); self.raw_path.write_text(self.raw_path.read_text() + '\n')
        with self.assertRaisesRegex(ValueError, 'evidence changed'): hw.prepare_data(self.dim, self.ci_path)

    def test_wrong_mapping_rejected(self):
        self.fixture(); self.device['SourceDeviceId'] = 'different'; self.write(self.dim, [self.device])
        with self.assertRaisesRegex(ValueError, 'correlation'): hw.prepare_data(self.dim, self.ci_path)

    def test_foreign_row_rejected(self):
        self.fixture(); self.ci['TenantId'] = 'foreign'; self.save_fixture()
        with self.assertRaisesRegex(ValueError, 'tenant'): hw.prepare_data(self.dim, self.ci_path)

    def test_changed_ci_attribute_rejected(self):
        self.fixture(); self.ci['SerialNumber'] = 'changed'; self.save_fixture()
        with self.assertRaisesRegex(ValueError, 'differs'): hw.prepare_data(self.dim, self.ci_path)

    def test_invalid_capacity_rejected(self):
        self.fixture(); self.ci['TotalStorageSpaceInBytes'] = self.raw['TotalStorageSpaceInBytes'] = '-1'; self.save_fixture()
        with self.assertRaisesRegex(ValueError, 'capacity'): hw.prepare_data(self.dim, self.ci_path)

    def test_empty_validated_snapshot_remains_empty(self):
        self.fixture()
        self.write(self.raw_path, [], list(self.raw)); self.write(self.ci_path, [], hw.load_json(hw.PRODUCT / 'Schema/SmartWorkplaceCMDB.ci.hardware.json')['columns'])
        path = self.root / 'CIRegistry.manifest.json'; manifest = hw.load_json(path)
        manifest['Hardware'].update(RowCount=0, InputHashes={str(self.raw_path): hw.sha(self.raw_path)})
        path.write_text(json.dumps(manifest))
        _, rows, coverage, _ = hw.prepare_data(self.dim, self.ci_path)
        self.assertEqual(rows, []); self.assertEqual(coverage['Status'], 'Validated')

    def test_model_payloads_are_typed_and_filter_safe(self):
        definitions = hw.model_definitions(self.root, self.identity)
        self.assertEqual(definitions[0]['name'], 'DeviceHardware')
        self.assertIn('Unexpected tenant identity', definitions[0]['mExpression'])
        self.assertIn('Int64.Type', definitions[1]['mExpression'])
        measure = next(m for m in hw.measures() if m['name'] == 'Hardware detail rows')
        self.assertIn('ALLSELECTED', measure['expression'])
        coverage = next(m for m in hw.measures() if m['name'] == 'Hardware covered devices')
        self.assertIn('REMOVEFILTERS', coverage['expression'])
        self.assertIn("'DeviceHardware'[Manufacturer]", coverage['expression'])
        rate = next(m for m in hw.measures() if m['name'] == 'Hardware coverage rate')
        self.assertEqual(rate['formatString'], '0.0%')
        gap = next(m for m in hw.measures() if m['name'] == 'Hardware data gap rate')
        self.assertEqual(gap['expression'], 'DIVIDE([Hardware uncovered devices], [Devices])')
        repeated = next(m for m in hw.measures() if m['name'] == 'Repeated serial values')
        self.assertIn('Not provided', repeated['expression'])
        self.assertIn('> 1', repeated['expression'])
        fleet_gate = next(m for m in hw.measures() if m['name'] == 'Hardware fleet rows')
        self.assertEqual(fleet_gate['expression'], "COUNTROWS('DeviceHardware')")
        self.assertTrue(fleet_gate['isHidden'])
        self.assertTrue(all(m['formatString'] for m in hw.measures()))

    def test_fleet_measure_contract_is_stable_and_filter_safe(self):
        names = {m['name'] for m in hw.measures()}
        self.assertTrue({'Hardware covered devices', 'Hardware uncovered devices', 'Hardware coverage rate',
                         'Missing serial records', 'Missing manufacturer records', 'Missing model records',
                         'Zero-reported storage records', 'Repeated serial values'} <= names)
        self.assertEqual(len(hw.measures()), 15)


if __name__ == '__main__':
    unittest.main()
