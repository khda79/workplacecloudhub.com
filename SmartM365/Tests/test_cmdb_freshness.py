"""Offline acquisition-age boundaries; no tenant data, collectors or report edits."""
import copy
import datetime as dt
import json
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'SmartInventory/PreparedEvidence'))
import cmdb_freshness as policy
import cmdb_prepare as preparation
import test_cmdb_preparation as fixtures

NOW = fixtures.NOW
APP_FILES = {'Intune_DiscoveredApps_Summary.csv', 'Intune_DiscoveredApps_AppDeviceRelations.csv'}


class PolicyTests(unittest.TestCase):
    def setUp(self):
        self.contract = preparation.load_json(preparation.CONTRACT)

    def evidence(self, apps=24, core=1):
        return [dict(File=item['file'],
                     StartedAtUtc=(NOW - dt.timedelta(hours=apps if item['file'] in APP_FILES else core)).isoformat(),
                     CompletedAtUtc=(NOW - dt.timedelta(hours=apps - 0.5 if item['file'] in APP_FILES else 0.5)).isoformat())
                for item in self.contract['sources']]

    def test_exact_targets_and_hard_boundaries(self):
        for age, warning in ((48, False), (168, False), (168.000001, True), (239, True), (240, True)):
            with self.subTest(age=age):
                assessed = policy.evaluate(self.contract, self.evidence(apps=age), NOW)
                self.assertEqual(len(assessed['Warnings']), 2 if warning else 0)
        with self.assertRaisesRegex(ValueError, 'Stale acquisition'):
            policy.evaluate(self.contract, self.evidence(apps=240.000001), NOW)
        policy.evaluate(self.contract, self.evidence(core=48), NOW)
        with self.assertRaisesRegex(ValueError, 'Stale acquisition'):
            policy.evaluate(self.contract, self.evidence(core=48.000001), NOW)

    def test_apps_do_not_expand_core_span_and_expiry_is_per_source(self):
        for age, remaining in ((216, 24), (180, 47)):
            assessed = policy.evaluate(self.contract, self.evidence(apps=age), NOW)
            self.assertEqual(len(assessed['Groups']), 2)
            self.assertEqual(policy.utc(assessed['ExpiresAtUtc']), NOW + dt.timedelta(hours=remaining))
        records = self.evidence(apps=200, core=48)
        records[0]['CompletedAtUtc'] = (NOW + dt.timedelta(minutes=1)).isoformat()
        with self.assertRaisesRegex(ValueError, 'interval exceeds.*Core'):
            policy.evaluate(self.contract, records, NOW)

    def test_invalid_policies_never_silently_fall_back(self):
        for bad in (True, 0, -1, float('nan'), float('inf'), 8761, '240'):
            contract = copy.deepcopy(self.contract)
            contract['freshnessGroups'][0]['maxAgeHours'] = bad
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                policy.policies(contract)
        for mutation in ('unknown', 'duplicate', 'overlap', 'warning', 'half-producer'):
            contract = copy.deepcopy(self.contract)
            group = contract['freshnessGroups'][0]
            if mutation == 'unknown': group['sources'].append('unknown.csv')
            elif mutation == 'duplicate': group['sources'].append(group['sources'][0])
            elif mutation == 'overlap': contract['freshnessGroups'].append(dict(group, name='Other'))
            elif mutation == 'warning': group['warningAgeHours'] = 241
            else: group['sources'].pop()
            with self.subTest(mutation=mutation), self.assertRaises(ValueError):
                policy.producer_policy(contract, sorted(APP_FILES))

    def test_missing_duplicate_reversed_future_and_naive_evidence(self):
        for kind in ('missing', 'duplicate', 'reverse', 'future', 'naive'):
            records = self.evidence()
            if kind == 'missing': records.pop()
            elif kind == 'duplicate': records.append(records[0])
            elif kind == 'reverse': records[0]['StartedAtUtc'] = NOW.isoformat()
            elif kind == 'future': records[0]['CompletedAtUtc'] = (NOW + dt.timedelta(minutes=6)).isoformat()
            else: records[0]['StartedAtUtc'] = '2026-01-15T00:00:00'
            with self.subTest(kind=kind), self.assertRaises(ValueError):
                policy.evaluate(self.contract, records, NOW)


class PreparationFreshnessTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.PreparationTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)

    def apps_age(self, age):
        for record in self.fixture.proof['Files']:
            if record['File'] in APP_FILES:
                record['StartedAtUtc'] = (NOW - dt.timedelta(hours=age)).isoformat()
                record['CompletedAtUtc'] = (NOW - dt.timedelta(hours=age - 1)).isoformat()
        self.fixture.write_proof()

    def test_weekly_apps_have_visible_evidence_and_findings(self):
        self.apps_age(216)
        result = self.fixture.prepare()
        self.assertEqual(result['Status'], 'Prepared')
        self.assertEqual(len(result['FreshnessWarnings']), 2)
        manifest = preparation.load_json(self.fixture.output / preparation.MANIFEST)
        self.assertEqual(len(manifest['SourceEvidence']['Freshness']['Warnings']), 2)
        findings = [row for row in self.fixture.table('FactDataQuality') if row['FindingType'] == 'SourceFreshnessWarning']
        self.assertEqual(len(findings), 2)
        self.assertTrue(all(row['EntityType'] == 'Source' for row in findings))
        health = [row for row in self.fixture.table('SourceHealth') if row['SourceName'] in APP_FILES]
        self.assertTrue(all('Aging' in row['Evidence'] and '240h' in row['Evidence'] for row in health))
        self.assertEqual(self.fixture.prepare(validate_only=True)['GeneratedTables'], 0)

    def test_expired_apps_preserve_current_output(self):
        self.fixture.unchanged_after(lambda: (self.apps_age(241), self.fixture.prepare()), 'Stale acquisition')

    def test_running_failed_partial_apps_are_not_rescued(self):
        self.apps_age(180)
        producer = next(item for item in preparation.load_json(preparation.REGISTRY)['Producers'] if set(item['Files']) == APP_FILES)
        path = self.fixture.source / producer['Receipt']
        original = path.read_bytes()
        for mutation in (dict(Status='Running'), dict(Status='Failed'), dict(IsPartialInventory=True), dict(Errors=1)):
            proof = json.loads(original)
            proof.update(mutation)
            path.write_text(json.dumps(proof), encoding='utf-8')
            with self.subTest(mutation=mutation), self.assertRaisesRegex(ValueError, 'Incomplete producer'):
                self.fixture.prepare(validate_only=True)
        path.write_bytes(original)


if __name__ == '__main__':
    unittest.main(verbosity=2)
