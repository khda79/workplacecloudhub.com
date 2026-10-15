"""Offline same-input migration tests; no production CSVs, APIs or publication."""
import datetime as dt
import base64
import csv
import json
import sys
import unittest
import uuid
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
        original_write_proof = self.fixture.write_proof
        def write_qualified_proofs():
            original_write_proof()
            self.extend_collaboration_fixture()
        self.fixture.write_proof = write_qualified_proofs
        self.extend_collaboration_fixture()

    def extend_collaboration_fixture(self):
        """Upgrade only synthetic Teams/SPO proofs to the extended capture."""
        for producer, names in unified.teams_sites.FILES.items():
            path = self.fixture.source / f'SmartInventory_SmartM365-{producer}-Inventory.current.json.txt'
            proof = json.loads(path.read_text(encoding='utf-8-sig'))
            definition = next(p for p in sources.load_json(sources.REGISTRY)['Producers'] if p['Script'] == proof['Producer'])
            proof.update(Owner='SmartInventory-SourceReceipt', ContractVersion='1.2',
                         ScopeQualification='ConsumerScope', RequiredFiles=definition['Files'],
                         FullInventoryQualified=False, ConsumerScopeQualified=True)
            for name in names:
                target = self.fixture.source / name
                if target.exists():
                    with target.open(encoding='utf-8-sig', newline='') as stream:
                        reader = csv.DictReader(stream); fields = reader.fieldnames; rows = list(reader)
                else:
                    rows = [] if producer == 'Teams' else [dict(StorageUsedGB='', StorageCapacityGB='', CapacitySource='Not measured', IsPartialInventory='False')]
                    fields = unified.teams_sites.EXTRA_HEADERS[name] if producer == 'Teams' else list(rows[0])
                rows = [dict(r, **IDENTITY, RunId='synthetic-business') for r in rows]
                fields = list(dict.fromkeys(fields + list(IDENTITY) + ['RunId']))
                with target.open('w', encoding='utf-8-sig', newline='') as stream:
                    writer = csv.DictWriter(stream, fieldnames=fields); writer.writeheader(); writer.writerows(rows)
                record = next((r for r in proof['Files'] if r['File'] == name), None)
                if record is None:
                    record = dict(proof['Files'][0], File=name); proof['Files'].append(record)
                record.update(SHA256=sources.sha(target), Rows=len(rows), Delimiter=',')
                original = next((r for r in self.fixture.proof['Files'] if r['File'] == name), None)
                if original is not None:
                    original.update(SHA256=sources.sha(target), Rows=len(rows))
            path.write_text(json.dumps(proof), encoding='utf-8')
            path.with_name(path.name.replace('.current.', '.run.')).write_text(json.dumps(dict(
                Owner='SmartInventory-SourceRun', ContractVersion='1.0', **IDENTITY,
                Producer=proof['Producer'], Status='Completed', RunId=proof['RunId'],
                PublicationStarted=True, UnqualifiedPublication=False)), encoding='utf-8')

    def audit(self):
        return unified.validate(self.fixture.source, IDENTITY['TenantKey'], NOW, IDENTITY)

    def compare(self):
        self.fixture.prepare()
        result = unified.build(self.fixture.source, self.output, self.audit())
        # The legacy SMTP collapse is deliberately NOT the mailbox oracle.
        # Dedicated tests below cover keys, counts, findings and graph relations.
        identity_outputs = {'FactMailbox','FactMailboxHosting','FactDataQuality','EntityFinding','FactRelationshipOverview'}
        for table in self.fixture.contract['tables']:
            if table['name'] in identity_outputs:
                continue
            name = table['name'] + '.csv'
            self.assertEqual(baseline.sha(self.fixture.output / name), sources.sha(self.output / name), name)
        for name in ('FactDataQuality','EntityFinding','FactRelationshipOverview'):
            with (self.fixture.output/(name+'.csv')).open(encoding='utf-8-sig',newline='') as stream:
                previous=list(csv.DictReader(stream))
            def unrelated(row):
                return row.get('EntityType','').casefold() not in ('mailbox','mailbox recipient') and row.get('RelationshipType')!='HasMailbox'
            self.assertEqual([r for r in previous if unrelated(r)], [r for r in self.graph_rows(name) if unrelated(r)],name)
        self.assertEqual(len(self.graph_rows('FactMailbox')),2)
        self.assertEqual(len(self.graph_rows('FactMailboxHosting')),2)
        self.assertEqual({r['SourceSystem'] for r in self.graph_rows('FactMailbox')},{'local_mailboxes','mailboxes'})
        self.assertEqual(len(result['OutputRows']), 46)
        self.assertFalse(result['Publication'])
        self.assertFalse((self.output / 'current.json.txt').exists())
        return result

    def seed_smtp_conflict(self, cloud_user='u1'):
        ids = ['00000000-0000-0000-0000-00000000000'+str(i) for i in range(1,5)]
        anchor = base64.b64encode(uuid.UUID(ids[1]).bytes_le).decode()
        self.fixture.inputs['local_mailboxes']=[]
        self.fixture.write_inputs()
        self.fixture.write_native_rows('remote_mailboxes',[
            dict(ObjectGuid=ids[0],PrimarySmtpAddress='user@synthetic.invalid',ExchangeGuid=ids[3],
                 RecipientTypeDetails='RemoteUserMailbox',CollectedAtUtc=NOW.isoformat(),ImmutableId=''),
            dict(ObjectGuid=ids[1],PrimarySmtpAddress='USER@synthetic.invalid',ExchangeGuid=str(uuid.UUID(int=0)),
                 RecipientTypeDetails='RemoteUserMailbox',CollectedAtUtc=NOW.isoformat(),ImmutableId=anchor)])
        self.fixture.write_native_rows('mailboxes',[
            dict(MailboxGuid=ids[2],PrimarySmtpAddress='user@synthetic.invalid',ExternalDirectoryObjectId=cloud_user,
                 RecipientTypeDetails='UserMailbox',CollectedAtUtc=NOW.isoformat(),OnPremisesImmutableId=anchor)])
        return ids

    def test_mailbox_conflict_counts_findings_keys_and_graph_endpoints(self):
        ids=self.seed_smtp_conflict()
        digests={p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()}
        result=unified.build(self.fixture.source,self.output,self.audit())
        mailboxes=self.graph_rows('FactMailbox');hosting=self.graph_rows('FactMailboxHosting')
        self.assertEqual(len(mailboxes),1);self.assertEqual(len(hosting),1)
        self.assertEqual(mailboxes[0]['SourceSystem'],'mailboxes')
        self.assertEqual(hosting[0]['MailboxHostingKey'],mailboxes[0]['TenantMailboxKey'])
        audit=result['Details']['MailboxIdentity']
        self.assertEqual(len(audit['Observations']),3)
        self.assertEqual(audit['UnresolvedRemoteRecipients'],1)
        self.assertEqual({o['NativeId'] for o in audit['Observations']},set(ids[:3]))
        remote={o['NativeId']:o for o in audit['Observations'] if o['SourceSystem']=='remote_mailboxes'}
        self.assertEqual(remote[ids[0]]['TenantMailboxKey'],'')
        self.assertEqual(remote[ids[1]]['TenantMailboxKey'],mailboxes[0]['TenantMailboxKey'])
        findings=self.graph_rows('FactDataQuality')
        conflicts=[r for r in findings if r['FindingType']=='ConflictingPrimarySmtpAddress']
        self.assertEqual(len(conflicts),2)
        self.assertTrue(any(r['FindingType']=='UnresolvedRemoteMailboxRecipient' for r in findings))
        self.assertEqual(len([r for r in self.graph_rows('CMDBEntityEvidence') if r['Entity Type']=='Mailbox']),1)
        links=[r for r in self.graph_rows('CMDBRelationshipEvidence') if r['Source Table']=='FactMailbox']
        self.assertEqual(len(links),1);self.assertEqual(links[0]['Resolution'],'Resolved')
        overview=next(r for r in self.graph_rows('FactRelationshipOverview') if r['RelationshipType']=='HasMailbox')
        self.assertEqual(overview['RelationshipCount'],'1')
        self.assertEqual(digests,{p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()})

    def test_ambiguous_smtp_cannot_fallback_to_user(self):
        self.seed_smtp_conflict(cloud_user='')
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertEqual(self.graph_rows('FactMailbox')[0]['TenantUserKey'],'')
        self.assertEqual(self.graph_rows('FactMailbox')[0]['LinkStatusLabel'],'Unresolved')

    def test_unknown_technical_user_cannot_fallback_to_known_smtp(self):
        self.seed_smtp_conflict(cloud_user='unknown-user')
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertEqual(self.graph_rows('FactMailbox')[0]['TenantUserKey'],'')

    def test_unqualified_publication_still_blocks_even_with_preserved_conflict(self):
        self.seed_smtp_conflict()
        self.fixture.source_run('Failed',unqualified=True)
        with self.assertRaisesRegex(ValueError,'publication is in progress or remains unqualified'):
            self.audit()
        self.assertFalse(list(self.output.iterdir()))

    def test_mailbox_keys_are_stable_after_smtp_rename(self):
        ids=self.seed_smtp_conflict()
        unified.build(self.fixture.source,self.output,self.audit())
        before=self.graph_rows('FactMailbox')[0]['TenantMailboxKey']
        self.fixture.write_native_rows('mailboxes',[
            dict(MailboxGuid=ids[2],PrimarySmtpAddress='renamed@synthetic.invalid',ExternalDirectoryObjectId='u1',
                 RecipientTypeDetails='UserMailbox',CollectedAtUtc=NOW.isoformat())])
        # A separate synthetic staging tree, never overwrite the first batch.
        self.output=self.fixture.root/'renamed-staging';self.output.mkdir()
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertEqual(self.graph_rows('FactMailbox')[0]['TenantMailboxKey'],before)

    def test_same_capture_all_native_columns(self):
        self.compare()

    def test_teams_channels_sites_flow_into_native_graph(self):
        self.fixture.inputs['teams']=[dict(TeamId='t1',TeamDisplayName='Synthetic team',
            MemberCollectionStatus='Collected',OwnerCollectionStatus='Collected',ChannelCollectionStatus='Collected',
            MemberCount='0',OwnerCount='0',GuestCount='0',IsArchived='false',IsPartialInventory='false',
            StandardChannelCount='1',PrivateChannelCount='2',SharedChannelCount='0',
            SharePointSiteId='host,s1,w1',SharePointSiteLinkStatus='Collected',CollectedAtUtc=NOW.isoformat())]
        self.fixture.inputs['sites']=[dict(SiteId='host,s1,w1',SiteIdentitySource='Graph',IsOneDrive='False',
            SiteUrl='https://synthetic.invalid/site',Title='Synthetic site',StorageUsedGB='10',StorageUsedMB='10240',
            CollectedAtUtc=NOW.isoformat(),ReportRefreshDate=NOW.date().isoformat())]
        self.fixture.write_inputs()
        channels=[dict(TeamId='t1',ChannelId=cid,ChannelDisplayName=cid,MembershipType=kind,
            SharePointSiteId=site,SharePointSiteLinkStatus='Collected' if site else 'NotApplicable')
            for cid,kind,site in [('c1','standard',''),('c2','private','host,s1,w1'),('c3','private','host,outside,w')]]
        with (self.fixture.source/'M365_Teams_Channels.csv').open('w',newline='',encoding='utf-8') as stream:
            writer=csv.DictWriter(stream,fieldnames=list(channels[0]));writer.writeheader();writer.writerows(channels)
        self.fixture.write_proof()
        before={p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()}
        unified.build(self.fixture.source,self.output,self.audit())
        links=[r for r in self.graph_rows('CMDBRelationshipEvidence') if r['Source Table']=='TeamsSiteLinkEvidence']
        self.assertEqual(len(links),6)
        self.assertEqual(sum(r['Resolution']=='Resolved' for r in links),5)
        self.assertEqual(len({r['Link Key'] for r in links}),6)
        self.assertEqual(len([r for r in self.graph_rows('CMDBEntityEvidence') if r['Entity Type']=='Channel']),3)
        self.assertEqual(self.graph_rows('SharePointSiteLinkEvidence')[0]['Link Count'],'2')
        self.assertEqual(before,{p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()})

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
        self.assertEqual({r['Source System'] for r in entities if r['Entity Type']=='Mailbox'},
                         {'Exchange Online','Exchange On-premises'})
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

    def test_native_dates_and_license_counts_use_output_types(self):
        self.fixture.inputs['users'][0].update(LastSignInDateTime='05/10/2026 09:08:07',
            LastNonInteractiveSignInDateTime='31/07/2026 08:00:34', LastSuccessfulSignInDateTime='2026-10-05T09:08:07Z')
        self.fixture.inputs['skus'][0].update(TenantConsumedUnits='8794', TenantPrepaidEnabled='9000')
        self.fixture.inputs['policies'] = [dict(PolicyId='p'+str(i), PolicyFamily=family,
            CreatedDateTime='17/07/2025 12:00:23', LastModifiedDateTime='05/10/2026 09:08:07',
            NativeEvidenceJson='{}', AssignmentCollectionStatus='Collected') for i, family in enumerate(
            ('SettingsCatalog', 'DeviceConfiguration', 'DeviceCompliance', 'WindowsFeatureUpdate', 'WindowsQualityUpdate'))]
        self.fixture.inputs['autopilot'] = [{'Autopilot ID':'ap1', 'Last contacted':'14/10/2026 09:00:00'}]
        self.fixture.write_inputs()
        unified.build(self.fixture.source, self.output, self.audit())
        self.assertEqual(self.graph_rows('DimUser')[0]['LastSignInDateTime'], '2026-10-05T09:08:07')
        self.assertEqual(self.graph_rows('DimLicenseSku')[0]['ConsumedUnits'], '8794')
        self.assertEqual(self.graph_rows('DimLicenseSku')[0]['EnabledUnits'], '9000')
        for table in ('DimIntuneConfigurationPolicy', 'DimWindowsUpdatePolicy'):
            self.assertEqual(self.graph_rows(table)[0]['CreatedDateTime'], '2025-07-17T12:00:23')
            self.assertEqual(self.graph_rows(table)[0]['LastModifiedDateTime'], '2026-10-05T09:08:07')
        self.assertEqual(self.graph_rows('FactAutopilotDevice')[0]['LastContactedDateTime'], '2026-10-14T09:00:00')

    def test_type_failure_precedes_graph_and_leaves_sources_unchanged(self):
        self.fixture.inputs['skus'][0]['TenantConsumedUnits'] = '1.5'
        self.fixture.write_inputs()
        before = {p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()}
        with mock.patch.object(unified, 'build_graph') as graph:
            with self.assertRaisesRegex(ValueError, 'int64'):
                unified.build(self.fixture.source, self.output, self.audit())
            graph.assert_not_called()
        self.assertEqual(before, {p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()})
        self.assertFalse((self.output / 'current.json.txt').exists())

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
