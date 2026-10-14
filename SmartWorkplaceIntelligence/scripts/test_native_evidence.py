"""Offline same-input migration tests; no production CSVs, APIs or publication."""
import datetime as dt
import csv
import json
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'SmartM365/Tests'))
from test_cmdb_preparation import PreparationTests, NOW, IDENTITY
import cmdb_prepare as baseline
import workplace_native_build as unified
import workplace_native_sources as sources


class NativeEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.fixture = PreparationTests(methodName='runTest')
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        self.output = self.fixture.root / 'unified-staging'
        self.output.mkdir()

    def audit(self):
        return unified.validate(self.fixture.source, IDENTITY['TenantKey'], NOW, IDENTITY)

    def compare(self):
        self.fixture.prepare()
        result = unified.build(self.fixture.source, self.output, self.audit())
        for table in self.fixture.contract['tables']:
            name = table['name'] + '.csv'
            self.assertEqual(baseline.sha(self.fixture.output / name), sources.sha(self.output / name), name)
        self.assertEqual(len(result['OutputRows']), 46)
        self.assertFalse(result['Publication'])
        self.assertFalse((self.output / 'current.json.txt').exists())
        return result

    def test_same_capture_all_native_columns(self):
        self.compare()

    def graph_rows(self, name):
        with (self.output / (name+'.csv')).open(encoding='utf-8-sig', newline='') as stream:
            return list(csv.DictReader(stream))

    def test_graph_keys_endpoints_and_native_counts(self):
        result=unified.build(self.fixture.source,self.output,self.audit())
        entities=self.graph_rows('CMDBEntityEvidence')
        links=self.graph_rows('CMDBRelationshipEvidence')
        endpoints=self.graph_rows('CMDBRelationshipEndpoint')
        self.assertEqual(len(entities),len({r['Entity Key'] for r in entities}))
        self.assertEqual(len(links),len({r['Link Key'] for r in links}))
        self.assertEqual(len(endpoints),len({(r['Link Key'],r['Entity Key']) for r in endpoints}))
        expected={(r['Link Key'],r[c]) for r in links for c in ('From Entity Key','To Entity Key') if r[c]}
        self.assertEqual(expected,{(r['Link Key'],r['Entity Key']) for r in endpoints})
        self.assertTrue({r['Entity Key'] for r in endpoints}<={r['Entity Key'] for r in entities})
        self.assertEqual(len(links),result['GraphRows']['CMDBRelationshipEvidence'])
        for source in ('FactDeviceApplication','FactUserServicePlan','DimLicenseServicePlan'):
            self.assertEqual(len(self.graph_rows(source)),len([r for r in links if r['Source Table']==source]))

    def test_graph_preserves_direct_and_primary_group_without_transitive_edges(self):
        self.fixture.seed_ad();self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        links=self.graph_rows('CMDBRelationshipEvidence')
        memberships=[r for r in links if r['Source Table']=='ADMembership']
        self.assertEqual({r['Relationship Type'] for r in memberships},{'Direct','PrimaryGroup'})
        self.assertEqual(len(memberships),2)
        self.assertTrue(all(r['Resolution']=='Resolved' for r in memberships))
        self.assertFalse(any('transitive' in r['Relationship Type'].lower() for r in links))

    def test_graph_unknown_parent_is_not_a_fabricated_entity(self):
        self.fixture.inputs['app_relations'].append({'AppId':'missing-app','DeviceId':'missing-device'})
        self.fixture.write_inputs();unified.build(self.fixture.source,self.output,self.audit())
        missing=[r for r in self.graph_rows('CMDBRelationshipEvidence') if r['From Name']=='missing-device']
        self.assertEqual(len(missing),1)
        self.assertEqual(missing[0]['Resolution'],'Unresolved endpoint')
        self.assertEqual(missing[0]['From Entity Key'],'')
        self.assertEqual(missing[0]['To Entity Key'],'')
        self.assertNotIn('missing-device',{r['Native ID'] for r in self.graph_rows('CMDBEntityEvidence')})

    def test_graph_manager_is_observed(self):
        self.fixture.seed_ad();self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        links=[r for r in self.graph_rows('CMDBRelationshipEvidence') if r['Relationship Type']=='Directory manager']
        self.assertEqual(len(links),1)
        self.assertEqual(links[0]['Resolution'],'Resolved')
        self.assertEqual(links[0]['Source Table'],'ADUserSource')

    def test_graph_missing_manager_stays_unresolved(self):
        self.fixture.seed_ad()
        self.fixture.inputs['ad_users'][0]['manager']='CN=Absent,DC=synthetic,DC=invalid'
        self.fixture.write_inputs();unified.build(self.fixture.source,self.output,self.audit())
        links=[r for r in self.graph_rows('CMDBRelationshipEvidence') if r['Relationship Type']=='Directory manager']
        self.assertEqual(len(links),1)
        self.assertEqual(links[0]['Resolution'],'Unresolved endpoint')
        self.assertEqual(links[0]['To Entity Key'],'')
        self.assertEqual(links[0]['To Name'],'CN=Absent,DC=synthetic,DC=invalid')

    def test_graph_nested_group_and_non_user_members_are_not_dropped(self):
        self.fixture.inputs['group_scope'][0]['MemberCount']='4'
        self.fixture.inputs['group_members']=[dict(GroupId='g1',MemberId=mid,MemberType=kind,MembershipKind='Direct',CollectionStatus='Collected')
            for mid,kind in [('u1','#microsoft.graph.user'),('eo1','#microsoft.graph.device'),('g1','#microsoft.graph.group'),('sp1','#microsoft.graph.servicePrincipal')]]
        self.fixture.write_inputs();unified.build(self.fixture.source,self.output,self.audit())
        links=[r for r in self.graph_rows('CMDBRelationshipEvidence') if r['Source Table']=='EntraGroupMembership']
        self.assertEqual(len(links),4)
        self.assertEqual(sum(r['Resolution']=='Resolved' for r in links),3)
        self.assertEqual(sum(r['From Entity Key']==r['To Entity Key'] for r in links),1)
        self.assertTrue(any(r['To Name']=='sp1' and r['Resolution']=='Unresolved endpoint' for r in links))
        aliases=self.graph_rows('CMDBEntityAlias')
        self.assertEqual(len(aliases),len({(r['Source Table'],r['Source Key']) for r in aliases}))

    def test_graph_reuses_canonical_estate_strong_links(self):
        path=self.output/'AllDevicesEvidence.csv'
        with path.open('w',encoding='utf-8',newline='') as stream:
            writer=csv.DictWriter(stream,fieldnames=['Device Key','AD Object GUID','Intune Device ID'])
            writer.writeheader();writer.writerow({'Device Key':'intune:md1','Intune Device ID':'md1'})
        unified.build(self.fixture.source,self.output,self.audit())
        bridges=[r for r in self.graph_rows('CMDBEntityAlias') if r['Canonical Device Key']]
        self.assertEqual(len(bridges),1)
        self.assertEqual(bridges[0]['Canonical Device Key'],'intune:md1')
        devices=[r for r in self.graph_rows('CMDBEntityEvidence') if r['Entity Type']=='Device']
        # Identical display names never merge the two managed identities.
        self.assertEqual(len(devices),2)

    def test_unresolved_application_links_are_preserved(self):
        self.fixture.inputs['app_relations'].append({'AppId': 'missing-app', 'DeviceId': 'missing-device'})
        self.fixture.write_inputs()
        self.compare()

    def test_unresolved_license_users_are_preserved(self):
        self.fixture.inputs['license_paths'].append({'UserId':'missing-user','SkuId':'s1','AssignedByGroupId':'g1','AssignmentState':'Error'})
        self.fixture.write_inputs()
        self.compare()

    def test_endpoint_score_sentinels_are_not_zero(self):
        self.fixture.inputs['analytics'][0]['EndpointAnalyticsScore'] = '-1'
        self.fixture.write_inputs()
        self.compare()

    def test_wrong_effective_identity_rejected(self):
        with self.assertRaisesRegex(ValueError, 'effective tenant'):
            unified.validate(self.fixture.source, 'synthetic', NOW, dict(IDENTITY, TenantId='other'))

    def test_changed_source_after_validation_rejected(self):
        audit = self.audit()
        path = self.fixture.source / self.fixture.contract['sources'][0]['file']
        with path.open('a', encoding='utf-8') as stream:
            stream.write('\n')
        with self.assertRaisesRegex(ValueError, 'Source changed'):
            unified.build(self.fixture.source, self.output, audit)

    def test_stale_acquisition_is_not_replaced_by_file_time(self):
        with self.assertRaisesRegex(ValueError, 'acquisition|Stale'):
            unified.validate(self.fixture.source, 'synthetic', NOW + dt.timedelta(days=30), IDENTITY)

    def test_no_overwrite_of_other_domain_output(self):
        target = self.output / 'DimUser.csv'
        target.write_text('sentinel', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'overwrite'):
            unified.build(self.fixture.source, self.output, self.audit())
        self.assertEqual(target.read_text(), 'sentinel')

    def test_ctrl_c_propagates_to_single_owner_cleanup(self):
        audit = self.audit()
        with mock.patch.object(unified, 'build_tables', side_effect=KeyboardInterrupt):
            with self.assertRaises(KeyboardInterrupt):
                unified.build(self.fixture.source, self.output, audit)
        self.assertFalse((self.output / 'current.json.txt').exists())

    def test_no_legacy_preparation_or_publication_dependency(self):
        code = Path(unified.__file__).read_text(encoding='utf-8')
        self.assertNotIn('import cmdb_', code)
        self.assertNotIn('baseline.prepare', code)
        self.assertFalse(hasattr(sources, 'prepare'))
        self.assertFalse(hasattr(sources, 'PublicationLock'))

    def test_recovery_mapping_is_exhaustive(self):
        root = Path(__file__).resolve().parents[1] / 'config'
        mapping = sources.load_json(root / 'consolidation-recovery-map.json.txt')
        contract = sources.load_json(sources.CONTRACT)
        self.assertEqual({t['name'] for t in contract['tables']}, {t['sourceTable'] for t in mapping['tables']})
        self.assertEqual(len(mapping['measures']), 249)
        self.assertEqual(len({m['sourceMeasure'] for m in mapping['measures']}), 249)
        self.assertEqual(sum(m['disposition'] == 'Reuse' for m in mapping['measures']), 59)
        self.assertEqual(len(mapping['pages']), 13)
        for table in contract['tables']:
            mapped = next(t for t in mapping['tables'] if t['sourceTable'] == table['name'])
            self.assertEqual(table['columns'], [c['source'] for c in mapped['columns']])
            self.assertTrue(mapped['test'])


if __name__ == '__main__':
    unittest.main()
