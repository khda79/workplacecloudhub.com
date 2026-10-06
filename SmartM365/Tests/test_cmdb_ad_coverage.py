"""Synthetic AD partial coverage through the real preparation transaction."""
import copy
import datetime as dt
import json
import importlib.util
import os
from pathlib import Path
import unittest
from unittest import mock

import test_cmdb_preparation as fixtures
from test_cmdb_preparation import IDENTITY, NOW, pipeline


class AdCoverageTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.PreparationTests('runTest')
        self.fixture.setUp()
        # Use the operational registry/contract, including the sixth Licences source.
        self.contract_path = pipeline.CONTRACT
        self.fixture.inputs['ad_domains'] = [dict(DNSRoot='good.synthetic.invalid', DomainSID='S-1-5-21-1-2-3')]
        self.fixture.inputs['ad_computers'][0]['DomainName'] = 'good.synthetic.invalid'
        self.fixture.write_inputs()
        self.definition = next(p for p in pipeline.load_json(pipeline.REGISTRY)['Producers']
                               if p['Script'] == 'SmartM365-ActiveDirectory-Inventory.ps1')
        self.path = self.fixture.source / self.definition['Receipt']
        self.coverage = dict(Kind='ADDomainCoverage', Reason='NonBlockingDomainErrors', Status='PartialAccepted',
            ExpectedDomains=['good.synthetic.invalid', 'missing.synthetic.invalid'],
            CollectedDomains=['good.synthetic.invalid'], UnavailableDomains=['missing.synthetic.invalid'],
            NonBlockingDomainErrors=['missing.synthetic.invalid'])
        self.attach_coverage()

    def tearDown(self):
        self.fixture.tearDown()

    def prepare(self, **options):
        return self.fixture.prepare(contract_path=self.contract_path, **options)

    def attach_coverage(self):
        proof = pipeline.load_json(self.path)
        proof.update(Owner='SmartInventory-SourceReceipt', ContractVersion='1.2', ScopeQualification='ConsumerScope',
                     RequiredFiles=self.definition['Files'], ConsumerScopeQualified=True, IsPartialInventory=True,
                     DomainCoverage=copy.deepcopy(self.coverage))
        for record in proof['Files']:
            record.update(IsPartialInventory=True, DomainCoverage=copy.deepcopy(self.coverage))
        self.save(proof)

    def save(self, proof):
        self.path.write_text(json.dumps(proof), encoding='utf-8')

    def test_partial_is_visible_in_manifest_and_all_six_health_rows(self):
        result = self.prepare()
        self.assertEqual(len(result['CoverageWarnings']), 1)
        manifest = pipeline.load_json(self.fixture.output / pipeline.MANIFEST)
        self.assertEqual(manifest['PreparationQualifications']['ADDomainCoverage'], [self.coverage])
        ad_health = [r for r in self.fixture.table('SourceHealth') if r['SourceName'] in self.definition['Files']]
        self.assertEqual(len(ad_health), 6)
        self.assertTrue(all(r['Status'] == 'SuccessWithWarnings' and 'missing.synthetic.invalid' in r['Coverage'] for r in ad_health))
        self.assertEqual(self.fixture.table('ADDomainSource')[0]['DNSRoot'], 'good.synthetic.invalid')

    def test_validate_only_returns_warning_without_creating_output(self):
        result = self.prepare(validate_only=True)
        self.assertEqual(len(result['CoverageWarnings']), 1)
        self.assertFalse(self.fixture.output.exists())

    def test_invalid_parent_coverage_never_replaces_existing_output(self):
        cases = [dict(NonBlockingDomainErrors=['other.synthetic.invalid']), dict(CollectedDomains=[]),
                 dict(ExpectedDomains=['third.synthetic.invalid']),
                 dict(CollectedDomains=['good.synthetic.invalid', 'missing.synthetic.invalid']),
                 dict(CollectedDomains=['good.synthetic.invalid', 'GOOD.synthetic.invalid']),
                 dict(UnavailableDomains=['']), dict(Reason='IgnoredErrors')]
        original = pipeline.load_json(self.path)
        self.prepare()
        before = pipeline.sha(self.fixture.output / pipeline.MANIFEST)
        for change in cases:
            with self.subTest(change=change):
                proof = copy.deepcopy(original)
                proof['DomainCoverage'].update(change)
                self.save(proof)
                with self.assertRaisesRegex(ValueError, 'partial AD coverage|Partial AD coverage'):
                    self.prepare()
                self.assertEqual(pipeline.sha(self.fixture.output / pipeline.MANIFEST), before)

    def test_file_must_repeat_exact_parent_coverage(self):
        proof = pipeline.load_json(self.path)
        proof['Files'][0]['DomainCoverage']['CollectedDomains'] = ['other.synthetic.invalid']
        self.save(proof)
        with self.assertRaisesRegex(ValueError, 'domain coverage differs'):
            self.prepare()

    def test_failed_or_error_receipt_stays_blocking(self):
        original = pipeline.load_json(self.path)
        for change in (dict(Status='Failed'), dict(Errors=1), dict(Errors='0'), dict(ConsumerScopeQualified=False)):
            with self.subTest(change=change):
                proof = copy.deepcopy(original); proof.update(change); self.save(proof)
                with self.assertRaises(ValueError):
                    self.prepare()

    def test_generic_partial_producer_is_not_accepted(self):
        users = pipeline.load_json(pipeline.REGISTRY)['Producers'][0]
        path = self.fixture.source / users['Receipt']
        proof = pipeline.load_json(path)
        proof.update(Owner='SmartInventory-SourceReceipt', ContractVersion='1.2', ScopeQualification='ConsumerScope',
                     RequiredFiles=users['Files'], ConsumerScopeQualified=True, IsPartialInventory=True,
                     DomainCoverage=self.coverage)
        path.write_text(json.dumps(proof), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'Invalid partial AD coverage'):
            self.prepare()

    def test_unavailable_domain_rows_are_not_accepted_as_current(self):
        self.fixture.inputs['ad_computers'][0]['DomainName'] = 'missing.synthetic.invalid'
        self.fixture.write_inputs(); self.attach_coverage()
        with self.assertRaisesRegex(ValueError, 'unavailable or unqualified domain'):
            self.prepare()

    def test_declared_collected_domains_must_match_native_domain_export(self):
        self.fixture.inputs['ad_domains'][0]['DNSRoot'] = 'other.synthetic.invalid'
        self.fixture.write_inputs(); self.attach_coverage()
        with self.assertRaisesRegex(ValueError, 'domain export differs'):
            self.prepare()

    def test_partial_coverage_cannot_bypass_hash_or_freshness(self):
        proof = pipeline.load_json(self.path)
        proof['Files'][0]['SHA256'] = '0' * 64
        self.save(proof)
        with self.assertRaisesRegex(ValueError, 'hash mismatch'):
            self.prepare()
        self.fixture.write_inputs(); self.attach_coverage()
        with self.assertRaisesRegex(ValueError, 'Stale acquisition'):
            pipeline.validate_sources(self.fixture.source, self.fixture.contract, 'synthetic',
                now=NOW + dt.timedelta(hours=49), identity=IDENTITY)

    @unittest.skipUnless(os.environ.get('CMDB_CANDIDATE_READER'), 'Unpublished buffered-reader candidate is optional')
    def test_buffered_reader_preserves_partial_coverage_and_rejects_tampering(self):
        # Optional integration with the separately reviewed, unpublished reader.
        spec = importlib.util.spec_from_file_location('cmdb_ad_candidate_reader',
            Path(os.environ['CMDB_CANDIDATE_READER']))
        reader = importlib.util.module_from_spec(spec)
        import sys
        with mock.patch.dict(sys.modules, {spec.name: reader}):
            spec.loader.exec_module(reader)
        # Use this worktree's shared policy, not unrelated working-tree changes.
        with mock.patch.object(reader, 'freshness', pipeline.cmdb_freshness), \
                mock.patch.object(reader, 'REGISTRY', pipeline.REGISTRY):
            self.prepare()
            snapshot = reader.load_snapshot(self.fixture.output, IDENTITY, now=NOW,
                contract_path=self.contract_path, registry_path=pipeline.REGISTRY)
            ad_rows = [r for r in snapshot.iter_rows('SourceHealth') if r['SourceName'] in self.definition['Files']]
            self.assertTrue(all(r['Status'] == 'SuccessWithWarnings' for r in ad_rows))
            path = self.fixture.output / pipeline.MANIFEST
            manifest = pipeline.load_json(path)
            record = next(r for r in manifest['SourceEvidence']['Files'] if r['Producer'] == self.definition['Script'])
            record['IsPartialInventory'] = False
            path.write_text(json.dumps(manifest), encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'Incomplete source result'):
                reader.load_snapshot(self.fixture.output, IDENTITY, now=NOW,
                    contract_path=self.contract_path, registry_path=pipeline.REGISTRY)
            record['IsPartialInventory'] = True
            record['DomainCoverage']['NonBlockingDomainErrors'] = ['other.synthetic.invalid']
            path.write_text(json.dumps(manifest), encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'domain coverage mismatch'):
                reader.load_snapshot(self.fixture.output, IDENTITY, now=NOW,
                    contract_path=self.contract_path, registry_path=pipeline.REGISTRY)


if __name__ == '__main__':
    unittest.main()
