"""Synthetic-only Teams / SPO tests; never reads runtime DATA or publishes."""
import copy
import csv
import datetime as dt
import json
import tempfile
import unittest
from pathlib import Path

import workplace_teams_sites as module

NOW = dt.datetime(2026, 1, 10, 12, tzinfo=dt.timezone.utc)
IDENTITY = dict(TenantKey='synthetic', OrganizationKey='example', EnvironmentKey='test', TenantId='tenant-test')


class TeamsSitesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='workplace-teams-sites-test-')
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.team = dict(TeamId='team1', TeamDisplayName='Same label', SharePointSiteId='host,s1,w1',
                         SharePointSiteLinkStatus='Collected', IsPartialInventory='False', ChannelCollectionStatus='Collected',
                         StandardChannelCount='1', PrivateChannelCount='1', SharedChannelCount='0')
        self.channels = [dict(TeamId='team1', ChannelId='standard', MembershipType='standard',
                              SharePointSiteId='', SharePointSiteLinkStatus='NotApplicable'),
                         dict(TeamId='team1', ChannelId='private', MembershipType='private',
                              SharePointSiteId='host,s2,w2', SharePointSiteLinkStatus='Collected')]
        self.sites = [dict(SiteId=f'host,s{i},w{i}', Title='Same label', IsOneDrive='False', StorageUsedGB='10',
                           CollectionScope='UsageReportSitesExcludingOneDrive', ReportRefreshDate='2026-01-08',
                           CollectedAtUtc=NOW.isoformat()) for i in (1,2,3)]
        self.cap = dict(StorageUsedGB='30', StorageCapacityGB='', CapacitySource='Get-PnPTenantUnauthorized', IsPartialInventory='False')
        self.write()

    def write(self):
        groups = {'Teams': [[self.team], self.channels], 'SPO': [self.sites, [self.cap]]}
        for producer, datasets in groups.items():
            proof = dict(IDENTITY, Status='Completed', Errors=0, IsPartialInventory=False,
                         ConsumerScopeQualified=True, ScriptVersion='test', RunId=producer+'-receipt',
                         StartedAtUtc=(NOW-dt.timedelta(minutes=5)).isoformat(), CompletedAtUtc=NOW.isoformat(), Files=[])
            for name, data in zip(module.FILES[producer], datasets):
                values = [dict(r, **IDENTITY, RunId=producer+'-business') for r in data]
                fields = list(dict.fromkeys(k for row in values for k in row))
                with (self.root/name).open('w',newline='',encoding='utf-8') as stream:
                    writer=csv.DictWriter(stream,fieldnames=fields);writer.writeheader();writer.writerows(values)
                proof['Files'].append(dict(File=name, Status='Success', Errors=0, IsPartialInventory=False,
                    Delimiter=',', RunId=proof['RunId'], SHA256=module.digest(self.root/name), Rows=len(values)))
            self.save_proof(producer,proof)
            (self.root/f'SmartInventory_SmartM365-{producer}-Inventory.run.json.txt').write_text(
                json.dumps(dict(Status='Completed',RunId=proof['RunId'],UnqualifiedPublication=False)))

    def save_proof(self, producer, proof):
        (self.root/f'SmartInventory_SmartM365-{producer}-Inventory.current.json.txt').write_text(json.dumps(proof))

    def audit(self):
        return module.validate(self.root,IDENTITY,NOW)

    def project(self, audit=None):
        return module.project([self.team],self.channels,self.sites,[self.cap],'synthetic',audit or self.audit())

    def test_hash_counts_lineage_and_qualified_ratio(self):
        audit=self.audit();data=self.project(audit)
        self.assertTrue(audit['Coherent'])
        sites=data['SharePointSiteLinkEvidence']
        self.assertEqual(sum(r['Teams Link State']=='Observed Teams link' for r in sites),2)
        self.assertEqual(len(sites),3)
        self.assertTrue(all(r['Link Coverage Qualified'] for r in sites))
        self.assertEqual(sum(r['Storage Used GB'] for r in sites),30)
        self.assertEqual(data['TeamsSiteLinkEvidence'][1]['Resolution'],'Not applicable')
        self.assertEqual(data['SharePointTenantCapacityEvidence'][0]['Capacity GB'],'')

    def test_unmatched_sites_preserved_no_name_join(self):
        self.team['SharePointSiteId']='host,outside,w';self.write()
        data=self.project()
        self.assertEqual(data['TeamsSiteLinkEvidence'][0]['Resolution'],'Outside observed SPO scope')
        self.assertEqual(len(data['TeamsSiteLinkEvidence']),3)

    def test_old_schema_unknown_not_negative(self):
        del self.team['SharePointSiteLinkStatus'];self.write()
        self.assertFalse(self.project()['SharePointSiteLinkEvidence'][0]['Link Coverage Qualified'])
        self.assertEqual(self.project()['SharePointSiteLinkEvidence'][2]['Teams Link State'],'Unknown')

    def test_case_normalization_and_no_double_count(self):
        self.channels[1]['SharePointSiteId']='HOST,S1,W1';self.write()
        data=self.project()['SharePointSiteLinkEvidence']
        self.assertEqual(data[0]['Link Count'],2)
        self.assertEqual(sum(r['Teams Link State']=='Observed Teams link' for r in data),1)

    def test_missing_channel_prevents_qualified_denominator(self):
        self.team['PrivateChannelCount']='2';self.write()
        self.assertFalse(self.project()['SharePointSiteLinkEvidence'][0]['Link Coverage Qualified'])

    def test_channel_status_and_blank_capacity_stay_unknown(self):
        self.team['ChannelCollectionStatus']='Unavailable';self.write()
        self.assertFalse(self.project()['SharePointSiteLinkEvidence'][0]['Link Coverage Qualified'])
        self.assertEqual(self.project()['SharePointTenantCapacityEvidence'][0]['Capacity State'],'Not measured')

    def test_required_extra_schema_is_checked_even_when_empty(self):
        path=self.root/'M365_Teams_Channels.csv';path.write_text('TeamId,ChannelId\n')
        proof=self.audit_proof('Teams')
        record=next(r for r in proof['Files'] if r['File']==path.name)
        record.update(Rows=0,SHA256=module.digest(path));self.save_proof('Teams',proof)
        with self.assertRaisesRegex(ValueError,'schema'):self.audit()

    def audit_proof(self, producer):
        return json.loads((self.root/f'SmartInventory_SmartM365-{producer}-Inventory.current.json.txt').read_text())

    def test_duplicate_keys_and_orphans_rejected(self):
        self.sites.append(copy.deepcopy(self.sites[0]))
        with self.assertRaisesRegex(ValueError,'duplicate'):self.project()
        self.sites.pop();self.channels[0]['TeamId']='unknown'
        with self.assertRaisesRegex(ValueError,'parent'):self.project()

    def test_measured_capacity_not_site_quota_or_estimate(self):
        self.cap.update(StorageCapacityGB='100',CapacitySource='Get-PnPTenant.StorageQuota');self.write()
        self.assertEqual(self.project()['SharePointTenantCapacityEvidence'][0]['Utilization'],.3)
        self.cap['CapacitySource']='Estimated'
        self.assertEqual(self.project()['SharePointTenantCapacityEvidence'][0]['Capacity GB'],'')

    def test_hash_mutation_and_missing_file(self):
        path=self.root/'M365_Teams_Channels.csv';path.write_text(path.read_text()+'\n')
        with self.assertRaisesRegex(ValueError,'hash'):self.audit()
        path.unlink()
        with self.assertRaises(FileNotFoundError):self.audit()

    def test_stale_wrong_tenant_and_partial_proofs(self):
        for field,value in [('TenantKey','other'),('IsPartialInventory',True),('StartedAtUtc',(NOW-dt.timedelta(days=10)).isoformat())]:
            self.write();proof=self.audit()['Proofs']['Teams'];proof[field]=value;self.save_proof('Teams',proof)
            with self.assertRaises(ValueError):self.audit()

    def test_skew_disables_ratio(self):
        audit=self.audit();audit['Coherent']=False
        self.assertFalse(self.project(audit)['SharePointSiteLinkEvidence'][0]['Link Coverage Qualified'])

    def test_recheck_and_output_schema(self):
        audit=self.audit();output=self.root/'synthetic-output';output.mkdir()
        counts=module.build(self.root,output,'synthetic',audit)
        self.assertEqual(counts['TeamsSiteLinkEvidence'],3)
        for name,(_,fields) in module.SCHEMAS.items():
            with (output/(name+'.csv')).open() as stream:self.assertEqual(next(csv.reader(stream)),list(fields))
        with self.assertRaises(FileExistsError):module.build(self.root,output,'synthetic',audit)
        self.save_proof('Teams',{})
        with self.assertRaisesRegex(ValueError,'changed'):module.recheck(self.root,audit)


if __name__=='__main__':unittest.main()
