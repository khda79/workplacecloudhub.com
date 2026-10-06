"""Offline CMDB transfer-plan tests using the real synthetic preparation fixture."""
import datetime as dt
import json
import shutil
import subprocess
import sys
import unittest
from pathlib import Path
from unittest import mock

import test_cmdb_preparation as fixtures
import cmdb_transfer as transfer


def create_case():
    fixtures.NOW = dt.datetime.now(dt.timezone.utc) - dt.timedelta(minutes=1)
    case = fixtures.PreparationTests()
    case.setUp()
    case.prepare()
    return case


class TransferPlanTests(unittest.TestCase):
    def setUp(self):
        self.case = create_case()
        self.root = self.case.output
        self.pointer = self.root / transfer.MANIFEST
        self.hash = transfer.digest(self.pointer)

    def tearDown(self):
        self.case.tearDown()

    def plan(self, **changes):
        return transfer.plan(self.root, fixtures.IDENTITY, changes.pop('expected_hash', self.hash), **changes)

    def change_manifest(self, change):
        manifest = json.loads(self.pointer.read_text())
        change(manifest)
        self.pointer.write_text(json.dumps(manifest))
        self.hash = transfer.digest(self.pointer)

    def test_real_generator_cohort_has_46_csvs_and_manifest_last_without_writes(self):
        before = {p.name: transfer.digest(p) for p in self.root.iterdir()}
        plan = self.plan()
        self.assertEqual(plan['CsvFiles'], 46)
        self.assertEqual(len(plan['Files']), 47)
        self.assertEqual(plan['Files'][-1]['Name'], 'current.json.txt')
        self.assertEqual(before, {p.name: transfer.digest(p) for p in self.root.iterdir()})

    def test_expected_hash_is_required_to_match(self):
        for value in ('0' * 64, 'invalid'):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, 'Expected manifest'):
                self.plan(expected_hash=value)

    def test_lowercase_pin_is_accepted(self):
        self.assertEqual(self.plan(expected_hash=self.hash.lower())['ManifestSHA256'], self.hash)

    def test_all_identity_fields_are_enforced(self):
        for field in fixtures.IDENTITY:
            identity = dict(fixtures.IDENTITY, **{field: 'foreign'})
            with self.subTest(field=field), self.assertRaisesRegex(ValueError, 'identity mismatch'):
                transfer.plan(self.root, identity, self.hash)

    def test_unvalidated_foreign_owner_and_contract_are_refused(self):
        original = self.pointer.read_bytes()
        for field, value in (('Status', 'Running'), ('Owner', 'Foreign'),
                             ('ContractVersion', 'obsolete'), ('ContractSHA256', '0' * 64)):
            self.pointer.write_bytes(original)
            self.change_manifest(lambda m: m.update({field: value}))
            with self.subTest(field=field), self.assertRaises(ValueError):
                self.plan()

    def test_altered_csv_is_refused(self):
        csv = next(self.root.glob('*.csv'))
        csv.write_bytes(csv.read_bytes() + b'changed')
        with self.assertRaisesRegex(ValueError, 'CSV hash'):
            self.plan()

    def test_missing_csv_is_refused(self):
        next(self.root.glob('*.csv')).unlink()
        with self.assertRaisesRegex(ValueError, 'artifacts'):
            self.plan()

    def test_extra_file_directory_and_legacy_json_are_refused(self):
        for name in ('extra.csv', 'current.json', 'extra-directory'):
            path = self.root / name
            if name == 'extra-directory':
                path.mkdir()
            else:
                path.write_text('extra')
            with self.subTest(name=name), self.assertRaisesRegex(ValueError, 'artifacts'):
                self.plan()
            if path.is_dir():
                path.rmdir()
            else:
                path.unlink()

    def test_invalid_row_count_and_output_hash_declarations_are_refused(self):
        original = self.pointer.read_bytes()
        name = next(iter(json.loads(original)['OutputFiles']))
        for field, value in (('Rows', -1), ('Rows', True), ('Rows', '1'), ('SHA256', 'invalid')):
            self.pointer.write_bytes(original)
            self.change_manifest(lambda m: m['OutputFiles'][name].update({field: value}))
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, 'declaration'):
                self.plan()

    def test_manifest_output_set_is_exact(self):
        self.change_manifest(lambda m: m['OutputFiles'].update({'extra.csv': {'Rows': 0, 'SHA256': '0' * 64}}))
        with self.assertRaisesRegex(ValueError, 'output set'):
            self.plan()

    def test_core_expiry_is_based_on_acquisition_not_generation_or_mtime(self):
        with self.assertRaisesRegex(ValueError, 'Stale acquisition'):
            self.plan(now=fixtures.NOW + dt.timedelta(hours=49))

    def test_weekly_apps_warn_after_168h_and_reject_after_240h(self):
        app_files = {s['file'] for s in self.case.contract['sources'] if s['name'] in ('apps', 'app_relations')}
        original = self.pointer.read_bytes()
        for age, rejected in ((169, False), (240, False), (241, True)):
            self.pointer.write_bytes(original)
            def change(m):
                for record in m['SourceEvidence']['Files']:
                    if record['File'] in app_files:
                        record['StartedAtUtc'] = (fixtures.NOW - dt.timedelta(hours=age)).isoformat()
                        record['CompletedAtUtc'] = (fixtures.NOW - dt.timedelta(hours=age-1)).isoformat()
            self.change_manifest(change)
            with self.subTest(age=age):
                if rejected:
                    with self.assertRaisesRegex(ValueError, 'Stale acquisition'):
                        self.plan(now=fixtures.NOW)
                else:
                    self.assertEqual(len(self.plan(now=fixtures.NOW)['FreshnessWarnings']), 2)

    def test_missing_or_repeated_source_proof_is_refused(self):
        self.change_manifest(lambda m: m['SourceEvidence']['Files'].pop())
        with self.assertRaisesRegex(ValueError, 'Freshness source set'):
            self.plan()

    def test_future_generation_is_refused(self):
        self.change_manifest(lambda m: m.update(GeneratedAtUtc=(fixtures.NOW + dt.timedelta(hours=1)).isoformat()))
        with self.assertRaisesRegex(ValueError, 'Future prepared'):
            self.plan()

    def test_manifest_change_during_checks_is_refused(self):
        original = transfer.digest
        def mutate(path):
            value = original(path)
            if Path(path).suffix == '.csv':
                self.pointer.write_bytes(self.pointer.read_bytes() + b' ')
            return value
        with mock.patch.object(transfer, 'digest', side_effect=mutate), self.assertRaisesRegex(ValueError, 'changed during'):
            self.plan()

    def test_optimized_python_still_rejects_corruption(self):
        result = subprocess.run([sys.executable, '-O', '-B', str(Path(transfer.__file__)), '--root', str(self.root),
            '--tenant-key', 'synthetic', '--organization-key', 'test', '--environment-key', 'test',
            '--tenant-id', 'synthetic-tenant', '--expected-manifest-sha256', '0' * 64], capture_output=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(b'Expected manifest SHA256 mismatch', result.stderr)


if __name__ == '__main__':
    if len(sys.argv) == 3 and sys.argv[1] == '--fixture-root':
        # Called only by the offline PowerShell harness, never by production.
        root = Path(sys.argv[2]).absolute()
        import tempfile
        if root.parent.resolve() != Path(tempfile.gettempdir()).resolve() or not root.name.startswith('SmartM365-CmdbTransferTest-'):
            raise ValueError('Expected an isolated system-temp fixture root')
        case = create_case()
        try:
            shutil.copytree(case.output, root / 'DATA-POWERBI-CMDB')
        finally:
            case.tearDown()
    else:
        unittest.main()
