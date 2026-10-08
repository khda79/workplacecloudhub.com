"""Synthetic/offline checks; no tenant access, real exports or report writes."""
import copy
import csv
import datetime as dt
import io
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import uuid
import shutil
from unittest.mock import patch

POWERBI = Path(__file__).resolve().parents[1] / 'PowerBI'
sys.path.insert(0, str(POWERBI))
import load_prepared_snapshot as consumer


class SnapshotTests(unittest.TestCase):
    def setUp(self):
        # Python 3.13+ mkdtemp's Windows owner-only DACL excludes sandbox
        # identities. A unique mkdir inherits the existing temp-root ACL.
        temporary = Path(tempfile.gettempdir()) / ('cmdb-consumer-' + uuid.uuid4().hex)
        temporary.mkdir()
        self.addCleanup(shutil.rmtree, temporary)
        self.root = temporary / 'DATA-POWERBI-CMDB'
        self.root.mkdir()
        self.contract = consumer._json(consumer.CONTRACT.read_bytes())
        self.registry = consumer._json(consumer.REGISTRY.read_bytes())
        self.now = dt.datetime(2026, 1, 15, 12, tzinfo=dt.timezone.utc)
        self.start = (self.now - dt.timedelta(hours=1)).isoformat()
        self.end = (self.now - dt.timedelta(minutes=30)).isoformat()
        self.identity = dict(TenantKey='example-test', OrganizationKey='example',
                             EnvironmentKey='test', TenantId='00000000-0000-0000-0000-000000000001')
        self.data = {}
        for definition in self.contract['tables']:
            name = definition['name']
            table_rows = []
            if not definition['allowEmpty']:
                row = {field: self.identity.get(field, '') for field in definition['columns']}
                for field in definition['key']:
                    row[field] = self.identity.get(field, 'synthetic')
                table_rows.append(row)
            self.write_table(name, table_rows)
        receipts, files = [], []
        for index, producer in enumerate(self.registry['Producers']):
            lineage = dict(Producer=producer['Script'], RunId='synthetic-' + str(index),
                           ScriptVersion='0.0.0', Scope=producer['Scope'],
                           StartedAtUtc=self.start, CompletedAtUtc=self.end)
            receipts.append(dict(lineage, File=producer['Receipt'], SHA256='A' * 64, Qualifications=[]))
            for file in producer['Files']:
                files.append(dict(lineage, File=file, SHA256='B' * 64, Rows=1,
                                  Status='Success', IsPartialInventory=False, Errors=0))
        self.manifest = dict(Owner=consumer.OWNER, Status='Validated', ScriptVersion='0.0.0',
                             ContractVersion=self.contract['version'],
                             ContractSHA256=consumer.digest(consumer.CONTRACT.read_bytes()),
                             TenantKey=self.identity['TenantKey'], Identity=self.identity.copy(),
                             GeneratedAtUtc=self.now.isoformat(),
                             # Producer machine paths need not match the synchronized consumer.
                             SourceRoot='remote-source', OutputRoot='remote-output',
                             SourceEvidence=dict(ProducerReceipts=receipts, Files=files,
                                                 RegistrySHA256=consumer.digest(consumer.REGISTRY.read_bytes()),
                                                 StartedAtUtc=self.start, CompletedAtUtc=self.end),
                             OutputFiles={})
        self.save_manifest()

    def definition(self, table):
        return next(item for item in self.contract['tables'] if item['name'] == table)

    def row(self, table, **values):
        row = {field: self.identity.get(field, '') for field in self.definition(table)['columns']}
        row.update(values)
        return row

    def write_table(self, table, table_rows):
        stream = io.StringIO(newline='')
        writer = csv.DictWriter(stream, fieldnames=self.definition(table)['columns'])
        writer.writeheader()
        writer.writerows(table_rows)
        data = stream.getvalue().encode('utf-8-sig')
        (self.root / (table + '.csv')).write_bytes(data)
        self.data[table + '.csv'] = (data, len(table_rows))

    def save_manifest(self):
        self.manifest['OutputFiles'] = {
            name: dict(SHA256=consumer.digest(data), Rows=count)
            for name, (data, count) in self.data.items()
        }
        self.write_manifest()

    def write_manifest(self):
        (self.root / consumer.MANIFEST).write_text(json.dumps(self.manifest), encoding='utf-8')

    def load(self, **options):
        return consumer.load_snapshot(self.root, self.identity, now=self.now, **options)

    def reject(self, pattern):
        with self.assertRaisesRegex(ValueError, pattern):
            self.load()

    def test_all_46_contract_tables_and_only_retained_bytes_are_returned(self):
        snapshot = self.load()
        self.assertEqual(len(snapshot.tables), 46)
        original = list(snapshot.iter_rows('DimCountry'))
        (self.root / 'DimCountry.csv').write_text('replaced', encoding='utf-8')
        self.assertEqual(list(snapshot.iter_rows('DimCountry')), original)
        self.assertEqual(snapshot.batch_sha256, consumer.digest(snapshot.manifest_bytes))
        with self.assertRaises(TypeError):
            snapshot.tables['DimCountry.csv'] = b'other'
        manifest = snapshot.manifest
        manifest['Identity']['TenantKey'] = 'other'
        self.assertEqual(snapshot.manifest['Identity'], self.identity)

    def test_no_files_are_written_by_loader(self):
        before = {item.name: item.read_bytes() for item in self.root.iterdir()}
        self.load()
        self.assertEqual(before, {item.name: item.read_bytes() for item in self.root.iterdir()})

    def test_mixed_synchronized_file_is_rejected(self):
        with (self.root / 'DimCountry.csv').open('ab') as stream:
            stream.write(b'changed')
        self.reject('Output hash mismatch')

    def test_missing_extra_and_directory_artifacts(self):
        for kind in ('missing', 'extra', 'directory'):
            with self.subTest(kind=kind):
                path = self.root / ('extra.csv' if kind == 'extra' else 'DimCountry.csv')
                if kind != 'extra':
                    path.unlink()
                if kind == 'extra':
                    path.write_bytes(b'')
                elif kind == 'directory':
                    path.mkdir()
                self.reject('artifacts|artifact')
                if path.is_dir():
                    path.rmdir()
                elif path.exists():
                    path.unlink()
                (self.root / 'DimCountry.csv').write_bytes(self.data['DimCountry.csv'][0])

    def test_identity_requires_all_four_expected_fields(self):
        for field in consumer.IDENTITY_FIELDS:
            with self.subTest(field=field):
                identity = self.identity.copy()
                identity.pop(field)
                with self.assertRaisesRegex(ValueError, 'Complete expected'):
                    consumer.load_snapshot(self.root, identity, now=self.now)

    def test_foreign_tenant_manifest_and_rows(self):
        self.manifest['Identity']['TenantId'] = 'different'
        self.write_manifest()
        self.reject('expected tenant')
        self.manifest['Identity'] = self.identity.copy()
        self.write_table('DimTenant', [self.row('DimTenant', TenantId='different')])
        self.save_manifest()
        self.reject('Reporting identity mismatch')

    def test_owner_status_contract_and_registry_mismatch(self):
        for field, value in (('Owner', 'other'), ('Status', 'Running'),
                             ('ContractVersion', 'old'), ('ContractSHA256', 'C' * 64)):
            with self.subTest(field=field):
                original = self.manifest[field]
                self.manifest[field] = value
                self.write_manifest()
                self.reject('snapshot|contract mismatch')
                self.manifest[field] = original
        self.manifest['SourceEvidence']['RegistrySHA256'] = 'C' * 64
        self.write_manifest()
        self.reject('registry mismatch')

    def test_row_counts_reject_bool_negative_fraction_and_mismatch(self):
        for count in (True, -1, 1.5, 99):
            with self.subTest(count=count):
                self.manifest['OutputFiles']['DimCountry.csv']['Rows'] = count
                self.write_manifest()
                self.reject('Row count|row count')

    def test_schema_and_malformed_csv_rows(self):
        original = self.data['DimCountry.csv'][0]
        for data in (b'TenantKey,TenantKey\n', b'wrong\n',
                     original + b'too,few\n', original + b'"unterminated\n'):
            with self.subTest(data=data[:30]):
                (self.root / 'DimCountry.csv').write_bytes(data)
                self.data['DimCountry.csv'] = (data, 2)
                self.save_manifest()
                self.reject('CSV')

    def test_empty_required_and_duplicate_blank_keys(self):
        for rows, pattern in (([], 'Unexpected empty'),
                              ([self.row('DimCountry', CountryLabel='')], 'immutable key'),
                              ([self.row('DimCountry', CountryLabel='A'),
                                self.row('DimCountry', CountryLabel=' a ')], 'immutable key')):
            with self.subTest(pattern=pattern):
                self.write_table('DimCountry', rows)
                self.save_manifest()
                self.reject(pattern)

    def test_quoted_newline_is_one_logical_row(self):
        self.write_table('DimCountry', [self.row('DimCountry', CountryLabel='one\ncountry')])
        self.save_manifest()
        self.assertEqual(list(self.load().iter_rows('DimCountry'))[0]['CountryLabel'], 'one\ncountry')

    def test_partial_source_scope_lineage_and_missing_producer(self):
        evidence = copy.deepcopy(self.manifest['SourceEvidence'])
        for field, value, pattern in (('IsPartialInventory', True, 'Incomplete'),
                                      ('Errors', True, 'Incomplete'), ('Status', 'Failed', 'Incomplete'),
                                      ('RunId', 'other', 'lineage mismatch')):
            with self.subTest(field=field):
                self.manifest['SourceEvidence'] = copy.deepcopy(evidence)
                self.manifest['SourceEvidence']['Files'][0][field] = value
                self.write_manifest()
                self.reject(pattern)
        self.manifest['SourceEvidence'] = copy.deepcopy(evidence)
        self.manifest['SourceEvidence']['ProducerReceipts'][0]['Scope'] = 'limited'
        self.write_manifest()
        self.reject('scope mismatch')
        self.manifest['SourceEvidence'] = copy.deepcopy(evidence)
        self.manifest['SourceEvidence']['ProducerReceipts'].pop()
        self.write_manifest()
        self.reject('Missing or unexpected')

    def test_refresh_timestamp_does_not_hide_expired_acquisition(self):
        with self.assertRaisesRegex(ValueError, 'Stale acquisition'):
            consumer.load_snapshot(self.root, self.identity, now=self.now + dt.timedelta(hours=48))

    def age_apps(self, hours):
        files = {'Intune_DiscoveredApps_Summary.csv', 'Intune_DiscoveredApps_AppDeviceRelations.csv'}
        start = (self.now - dt.timedelta(hours=hours)).isoformat()
        end = (self.now - dt.timedelta(hours=hours - 1)).isoformat()
        for record in self.manifest['SourceEvidence']['Files']:
            if record['File'] in files:
                record.update(StartedAtUtc=start, CompletedAtUtc=end)
                producer = record['Producer']
        for receipt in self.manifest['SourceEvidence']['ProducerReceipts']:
            if receipt['Producer'] == producer:
                receipt.update(StartedAtUtc=start, CompletedAtUtc=end)
        self.manifest['SourceEvidence']['StartedAtUtc'] = start
        self.write_manifest()

    def test_weekly_apps_do_not_extend_core_freshness(self):
        self.age_apps(216)
        snapshot = self.load()
        assessed = consumer.freshness.evaluate(self.contract, snapshot.manifest['SourceEvidence']['Files'], self.now)
        self.assertEqual(len(assessed['Warnings']), 2)
        self.assertEqual(consumer._utc(assessed['ExpiresAtUtc']), self.now + dt.timedelta(hours=24))
        self.age_apps(241)
        self.reject('Stale acquisition')

    def test_future_reversed_naive_and_aggregate_timestamps(self):
        for value in ((self.now + dt.timedelta(hours=1)).isoformat(),
                      (self.now - dt.timedelta(hours=2)).isoformat(), '2026-01-15T12:00:00'):
            with self.subTest(value=value):
                original = self.manifest['SourceEvidence']['ProducerReceipts'][0]['CompletedAtUtc']
                self.manifest['SourceEvidence']['ProducerReceipts'][0]['CompletedAtUtc'] = value
                self.write_manifest()
                self.reject('interval|timezone')
                self.manifest['SourceEvidence']['ProducerReceipts'][0]['CompletedAtUtc'] = original
        self.manifest['SourceEvidence']['CompletedAtUtc'] = self.now.isoformat()
        self.write_manifest()
        self.reject('Aggregate acquisition')

    def test_orphan_and_valid_device_relationship(self):
        self.write_table('FactDeviceApplication', [self.row('FactDeviceApplication',
                         TenantDeviceApplicationKey='edge',
                         TenantApplicationKey='app', ManagedDeviceId='managed')])
        self.save_manifest()
        self.reject('Orphan')
        self.write_table('DimDevice', [self.row('DimDevice', TenantDeviceKey='device')])
        app = self.row('DimDetectedApplication', TenantApplicationKey='app', AppId='native-app',
                       DeviceCount='1', ReportedDeviceCount='1', ExactRelatedDeviceCount='1',
                       RelationshipCoverageStatus='Complete')
        if 'ResolvedDeviceCount' in app:
            app.update(ResolvedDeviceCount='1', UnresolvedDeviceCount='0')
        self.write_table('DimDetectedApplication', [app])
        link = self.row('FactDeviceApplication', TenantDeviceApplicationKey='edge',
                        TenantApplicationKey='app', AppId='native-app', ManagedDeviceId='managed')
        if 'DeviceLinkStatus' in link:
            link.update(ApplicationLinkStatus='Resolved', DeviceLinkStatus='Resolved')
        self.write_table('FactDeviceApplication', [link])
        self.write_table('DimIntuneManagedDevice', [self.row('DimIntuneManagedDevice',
                         TenantIntuneDeviceKey='managed-key', ManagedDeviceId='managed', TenantDeviceKey='device')])
        self.save_manifest()
        self.assertEqual(len(list(self.load().iter_rows('FactDeviceApplication'))), 1)

    def test_manifest_change_during_load_is_rejected(self):
        original = consumer._validate_tables
        def change(*args):
            original(*args)
            self.manifest['ScriptVersion'] = 'changed'
            self.write_manifest()
        with patch.object(consumer, '_validate_tables', side_effect=change):
            self.reject('changed during loading|memory budget')

    def qualified_application_fixture(self):
        catalog = self.definition('DimDetectedApplication')
        catalog['columns'] += [c for c in ('ResolvedDeviceCount', 'UnresolvedDeviceCount')
                              if c not in catalog['columns']]
        links = self.definition('FactDeviceApplication')
        links['columns'] += [c for c in ('ApplicationLinkStatus', 'DeviceLinkStatus')
                            if c not in links['columns']]
        self.write_table('DimDetectedApplication', [self.row('DimDetectedApplication',
            TenantApplicationKey='app', AppId='native-app', DeviceCount='1',
            ReportedDeviceCount='1', ExactRelatedDeviceCount='1', ResolvedDeviceCount='0',
            UnresolvedDeviceCount='1', RelationshipCoverageStatus='Device links unresolved')])
        self.write_table('FactDeviceApplication', [self.row('FactDeviceApplication',
            TenantDeviceApplicationKey='edge', TenantApplicationKey='app', AppId='native-app',
            ManagedDeviceId='not-in-current-devices', ApplicationLinkStatus='Resolved',
            DeviceLinkStatus='Unresolved')])
        self.save_manifest()
        definitions = {d['name'] + '.csv': d for d in self.contract['tables']}
        buffers = {name: data[0] for name, data in self.data.items()}
        return buffers, definitions

    def test_qualified_weekly_application_device_gap_is_retained(self):
        buffers, definitions = self.qualified_application_fixture()
        consumer._validate_tables(buffers, definitions, self.identity, self.manifest)

    def test_qualified_weekly_application_gap_cannot_claim_resolution(self):
        buffers, definitions = self.qualified_application_fixture()
        buffers['FactDeviceApplication.csv'] = buffers['FactDeviceApplication.csv'].replace(b',Unresolved', b',Resolved')
        with self.assertRaisesRegex(ValueError, 'device link qualification'):
            consumer._validate_tables(buffers, definitions, self.identity, self.manifest)

    def test_qualified_application_coverage_counts_are_reconciled(self):
        buffers, definitions = self.qualified_application_fixture()
        buffers['DimDetectedApplication.csv'] = buffers['DimDetectedApplication.csv'].replace(b'Device links unresolved', b'Complete')
        with self.assertRaisesRegex(ValueError, 'relationship coverage'):
            consumer._validate_tables(buffers, definitions, self.identity, self.manifest)

    def test_pinned_manifest_and_memory_budget(self):
        with self.assertRaisesRegex(ValueError, 'Pinned manifest'):
            self.load(expected_manifest_sha256='F' * 64)
        pinned = consumer.digest((self.root / consumer.MANIFEST).read_bytes())
        self.assertEqual(self.load(expected_manifest_sha256=pinned).batch_sha256, pinned)
        with self.assertRaisesRegex(ValueError, 'memory budget'):
            self.load(max_bytes=1)

    def test_duplicate_json_properties_are_rejected(self):
        with (self.root / consumer.MANIFEST).open('w', encoding='utf-8') as stream:
            stream.write('{"Owner":"first","Owner":"second"}')
        self.reject('Duplicate JSON')

    def test_non_object_nonfinite_and_invalid_manifest_hash(self):
        for data, pattern in ((b'[]', 'document object'), (b'{"Rows":NaN}', 'Non-finite')):
            with self.subTest(data=data):
                (self.root / consumer.MANIFEST).write_bytes(data)
                self.reject(pattern)
        self.manifest['OutputFiles']['DimCountry.csv']['SHA256'] = 'not-a-hash'
        self.write_manifest()
        self.reject('Invalid SHA256')

    def test_manifest_exact_output_file_set(self):
        self.manifest['OutputFiles']['unexpected.csv'] = dict(Rows=0, SHA256='A' * 64)
        self.write_manifest()
        self.reject('output manifest entry')
        self.manifest['OutputFiles'].pop('unexpected.csv')
        self.manifest['OutputFiles'].pop('DimCountry.csv')
        self.write_manifest()
        self.reject('output manifest entry')

    def test_source_evidence_and_receipt_names_are_unique(self):
        for field in ('Files', 'ProducerReceipts'):
            with self.subTest(field=field):
                self.manifest['SourceEvidence'][field].append(copy.deepcopy(self.manifest['SourceEvidence'][field][0]))
                self.write_manifest()
                self.reject('repeated evidence name')
                self.manifest['SourceEvidence'][field].pop()

    def test_empty_required_source_evidence_is_rejected(self):
        contract = copy.deepcopy(self.contract)
        contract['sources'][0]['allowEmpty'] = False
        required = contract['sources'][0]['file']
        contract_path = self.root.parent / 'required-source-contract.json.txt'
        contract_path.write_text(json.dumps(contract), encoding='utf-8')
        self.manifest['ContractSHA256'] = consumer.digest(contract_path.read_bytes())
        next(record for record in self.manifest['SourceEvidence']['Files'] if record['File'] == required)['Rows'] = 0
        self.write_manifest()
        with self.assertRaisesRegex(ValueError, 'empty source evidence'):
            self.load(contract_path=contract_path)

    def test_evidence_expiring_during_validation_is_rejected(self):
        dates = iter((self.now, self.now + dt.timedelta(hours=48)))
        class Clock(dt.datetime):
            @classmethod
            def now(cls, tz=None):
                return next(dates)
        with patch.object(consumer.dt, 'datetime', Clock):
            with self.assertRaisesRegex(ValueError, 'Stale acquisition'):
                consumer.load_snapshot(self.root, self.identity)

    def test_contract_change_during_load_is_rejected(self):
        contract_path = self.root.parent / 'contract.json.txt'
        contract_path.write_bytes(consumer.CONTRACT.read_bytes())
        original = consumer._validate_tables
        def change(*args):
            original(*args)
            data = contract_path.read_bytes()
            contract_path.write_bytes(data.replace(b'"version"', b'"Version"', 1))
        with patch.object(consumer, '_validate_tables', side_effect=change):
            with self.assertRaisesRegex(ValueError, 'contract changed'):
                self.load(contract_path=contract_path)

    def test_read_is_chunked_not_preallocated_to_budget(self):
        sizes = []
        class TracedBytes(io.BytesIO):
            def read(self, size=-1):
                sizes.append(size)
                return super().read(size)
        with patch.object(Path, 'open', side_effect=lambda *args, **kwargs: TracedBytes(b'small')):
            self.assertEqual(consumer._read(self.root / 'DimCountry.csv', consumer.DEFAULT_MAX_BYTES), b'small')
        self.assertTrue(sizes)
        self.assertLessEqual(max(sizes), 4 * 1024 * 1024)

    def test_relationship_parent_mapping_matches_producer(self):
        # Review alarm if the generator's relationship guard evolves independently.
        import ast
        tree = ast.parse((consumer.PREPARATION / 'cmdb_prepare.py').read_text(encoding='utf-8-sig'))
        function = next(node for node in tree.body if isinstance(node, ast.FunctionDef)
                        and node.name == 'validate_relationships')
        self.assertEqual(consumer.PARENTS, ast.literal_eval(function.body[0].value))

    def test_real_generator_synthetic_output_is_accepted_without_source_access(self):
        fixture_path = consumer.PREPARATION.parents[1] / 'Tests/test_cmdb_preparation.py'
        spec = importlib.util.spec_from_file_location('cmdb_generator_fixture', fixture_path)
        fixtures = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(fixtures)
        fixture = fixtures.PreparationTests('test_all_46_tables_and_one_current_manifest')
        fixture.setUp()
        try:
            fixture.prepare()
            # The consumer must not reopen raw sources or producer receipts.
            shutil.rmtree(fixture.source)
            snapshot = consumer.load_snapshot(fixture.output, fixtures.IDENTITY, now=fixtures.NOW)
            self.assertEqual(len(snapshot.tables), 46)
            self.assertEqual(len(list(snapshot.iter_rows('DeviceHardware'))), 2)
            hardware = list(snapshot.iter_rows('DeviceHardware'))[0]
            self.assertEqual(hardware['PhysicalMemoryGiB'], '8.0')
        finally:
            fixture.tearDown()

    def test_cli_failure_is_nonzero_without_report_mutation(self):
        command = [sys.executable, '-B', str(POWERBI / 'load_prepared_snapshot.py'), '--root', str(self.root)]
        for field, value in self.identity.items():
            import re
            command.extend(['--' + re.sub(r'(?<!^)(?=[A-Z])', '-', field).lower(), value])
        # The synthetic reference clock is intentionally stale compared to now.
        result = subprocess.run(command, capture_output=True, text=True, timeout=20)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('Snapshot rejected:', result.stderr)


if __name__ == '__main__':
    unittest.main()
