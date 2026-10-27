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
        # Explicit, independent directory country in this synthetic baseline.
        self.fixture.inputs['users'][0]['CountryOrRegion'] = 'France'
        self.fixture.write_inputs()

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
            if table['name'] in {'ADUserSource','ADComputerSource','DimUser','DimDevice'}:
                # Lifecycle attributes are additive. Every pre-existing cell,
                # including immutable keys, must still equal the legacy oracle,
                # except the deliberately corrected country provenance.
                expected = list(baseline.rows(self.fixture.output / name))
                actual = self.output_rows(table['name'])
                self.assertEqual(len(expected), len(actual), name)
                ignored = {'CountryStatus'} if table['name'] == 'DimDevice' else set()
                self.assertEqual([{k:v for k,v in r.items() if k not in ignored} for r in expected],
                                 [{k:row[k] for k in old if k not in ignored} for old,row in zip(expected,actual)], name)
                continue
            self.assertEqual(baseline.sha(self.fixture.output / name), sources.sha(self.output / name), name)
        for name in ('FactDataQuality','EntityFinding','FactRelationshipOverview'):
            with (self.fixture.output/(name+'.csv')).open(encoding='utf-8-sig',newline='') as stream:
                previous=list(csv.DictReader(stream))
            def unrelated(row):
                return row.get('EntityType','').casefold() not in ('mailbox','mailbox recipient') and row.get('RelationshipType')!='HasMailbox'
            for r in previous:
                for field in ('Description', 'DescriptionLabel'):
                    if r.get(field) == 'Device country comes from the primary user usage location, not device geolocation.':
                        r[field] = 'Device country comes from the primary user Entra country, not device geolocation.'
            self.assertEqual([r for r in previous if unrelated(r)], [r for r in self.output_rows(name) if unrelated(r)],name)
        self.assertEqual(len(self.output_rows('FactMailbox')),2)
        self.assertEqual(len(self.output_rows('FactMailboxHosting')),2)
        self.assertEqual({r['SourceSystem'] for r in self.output_rows('FactMailbox')},{'local_mailboxes','mailboxes'})
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
        mailboxes=self.output_rows('FactMailbox');hosting=self.output_rows('FactMailboxHosting')
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
        findings=self.output_rows('FactDataQuality')
        conflicts=[r for r in findings if r['FindingType']=='ConflictingPrimarySmtpAddress']
        self.assertEqual(len(conflicts),2)
        self.assertTrue(any(r['FindingType']=='UnresolvedRemoteMailboxRecipient' for r in findings))
        self.assertTrue(mailboxes[0]['TenantUserKey'])
        overview=next(r for r in self.output_rows('FactRelationshipOverview') if r['RelationshipType']=='HasMailbox')
        self.assertEqual(overview['RelationshipCount'],'1')
        self.assertEqual(digests,{p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()})

    def test_ambiguous_smtp_cannot_fallback_to_user(self):
        self.seed_smtp_conflict(cloud_user='')
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertEqual(self.output_rows('FactMailbox')[0]['TenantUserKey'],'')
        self.assertEqual(self.output_rows('FactMailbox')[0]['LinkStatusLabel'],'Unresolved')

    def test_unknown_technical_user_cannot_fallback_to_known_smtp(self):
        self.seed_smtp_conflict(cloud_user='unknown-user')
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertEqual(self.output_rows('FactMailbox')[0]['TenantUserKey'],'')

    def test_unqualified_publication_still_blocks_even_with_preserved_conflict(self):
        self.seed_smtp_conflict()
        self.fixture.source_run('Failed',unqualified=True)
        with self.assertRaisesRegex(ValueError,'publication is in progress or remains unqualified'):
            self.audit()
        self.assertFalse(list(self.output.iterdir()))

    def test_mailbox_keys_are_stable_after_smtp_rename(self):
        ids=self.seed_smtp_conflict()
        unified.build(self.fixture.source,self.output,self.audit())
        before=self.output_rows('FactMailbox')[0]['TenantMailboxKey']
        self.fixture.write_native_rows('mailboxes',[
            dict(MailboxGuid=ids[2],PrimarySmtpAddress='renamed@synthetic.invalid',ExternalDirectoryObjectId='u1',
                 RecipientTypeDetails='UserMailbox',CollectedAtUtc=NOW.isoformat())])
        # A separate synthetic staging tree, never overwrite the first batch.
        self.output=self.fixture.root/'renamed-staging';self.output.mkdir()
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertEqual(self.output_rows('FactMailbox')[0]['TenantMailboxKey'],before)

    def test_same_capture_all_native_columns(self):
        self.compare()

    def test_entra_country_drives_user_device_and_hosting_not_usage_location(self):
        self.fixture.inputs['users'][0].update(CountryOrRegion='Spain', **{'Usage location': 'FR'})
        self.fixture.write_inputs()
        audit = self.audit()
        source_hashes = {p.name: sources.sha(p) for p in self.fixture.source.glob('*.csv')}
        result = unified.build(self.fixture.source, self.output, audit)
        user = self.output_rows('DimUser')[0]
        self.assertEqual((user['CountryCode'], user['CountryRaw'], user['CountryName'], user['CountrySource'], user['UsageLocation']),
                         ('ES', 'Spain', 'Spain', 'Entra country', 'FR'))
        device = next(r for r in self.output_rows('DimDevice') if r['PrimaryUserSourceId'] == user['SourceUserId'])
        self.assertEqual((device['CountryCode'], device['CountryStatus']), ('ES', 'Primary user Entra country'))
        mailbox_keys = {r['TenantMailboxKey'] for r in self.output_rows('FactMailbox') if r['TenantUserKey'] == user['TenantUserKey']}
        hosting = [r for r in self.output_rows('FactMailboxHosting') if r['MailboxHostingKey'] in mailbox_keys]
        self.assertTrue(hosting)
        self.assertTrue(all(r['CountryLabel'] == 'ES' for r in hosting))
        self.assertEqual(result['NativeContractVersion'], '0.5.0')
        self.assertEqual(source_hashes, {p.name: sources.sha(p) for p in self.fixture.source.glob('*.csv')})

    def test_missing_entra_country_never_uses_licensing_or_ad_country(self):
        self.fixture.seed_ad()
        self.fixture.inputs['ad_users'][0]['Country'] = 'DE'
        self.fixture.inputs['users'][0].pop('CountryOrRegion', None)
        self.fixture.inputs['users'][0]['Usage location'] = 'FR'
        self.fixture.write_inputs()
        unified.build(self.fixture.source, self.output, self.audit())
        user = self.output_rows('DimUser')[0]
        self.assertEqual((user['CountryCode'], user['CountryName'], user['CountryStatus']), ('', 'Unknown', 'Not provided'))
        self.assertEqual(user['UsageLocation'], 'FR')
        self.assertTrue(all(not r['CountryCode'] for r in self.output_rows('DimDevice')))

    def test_country_reference_changed_after_validation_is_rejected(self):
        audit = self.audit()
        audit['CountryReferenceSHA256'] = '0' * 64
        with self.assertRaisesRegex(ValueError, 'Country reference changed'):
            unified.build(self.fixture.source, self.output, audit)
        self.assertFalse(list(self.output.glob('*.csv')))

    def test_explicit_ad_password_expiry_retains_true_and_false(self):
        self.fixture.seed_ad()
        self.fixture.inputs['ad_users'][0]['PasswordExpired'] = 'True'
        self.fixture.inputs['ad_computers'][0]['PasswordExpired'] = 'False'
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        user = self.output_rows('ADUserSource')[0]
        computer = self.output_rows('ADComputerSource')[0]
        self.assertEqual((user['PasswordExpired'], user['PasswordExpiryStatus']), ('true','Observed'))
        self.assertEqual((computer['PasswordExpired'], computer['PasswordExpiryStatus']), ('false','Observed'))

    def test_missing_password_expiry_is_not_inferred_from_other_flags(self):
        self.fixture.seed_ad()
        for name in ('ad_users','ad_computers'):
            self.fixture.inputs[name][0].update(PasswordNeverExpires='False',
                PasswordMustChangeNextLogon='True', PasswordLastSet='2020-01-01T00:00:00Z')
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        for name in ('ADUserSource','ADComputerSource'):
            row = self.output_rows(name)[0]
            self.assertEqual((row['PasswordExpired'], row['PasswordExpiryStatus']), ('','Not collected'))

    def test_blank_password_expiry_is_unknown_and_invalid_boolean_is_rejected(self):
        self.fixture.seed_ad()
        self.fixture.inputs['ad_users'][0]['PasswordExpired'] = ''
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        row = self.output_rows('ADUserSource')[0]
        self.assertEqual((row['PasswordExpired'], row['PasswordExpiryStatus']), ('','Not provided'))
        self.fixture.inputs['ad_users'][0]['PasswordExpired'] = 'yes'
        self.fixture.write_inputs()
        output = self.fixture.root / 'invalid-expiry-staging'
        output.mkdir()
        with self.assertRaisesRegex(ValueError,'Invalid boolean evidence'):
            unified.build(self.fixture.source,output,self.audit())
        self.assertFalse((output / 'current.json.txt').exists())

    def test_ad_modification_raw_retained_without_naive_utc_inference(self):
        self.fixture.seed_ad()
        self.fixture.inputs['ad_users'][0]['WhenChanged'] = '2026-01-14T15:00:00'
        self.fixture.inputs['ad_computers'][0]['WhenChanged'] = '2026-01-14T15:00:00+02:00'
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        user = self.output_rows('ADUserSource')[0]
        computer = self.output_rows('ADComputerSource')[0]
        self.assertEqual(user['ModificationRaw'], '2026-01-14T15:00:00')
        self.assertEqual(user['ModificationUtcDateTime'], '')
        self.assertEqual(user['ModificationStatus'], 'Timezone not established')
        self.assertEqual(computer['ModificationRaw'], '2026-01-14T15:00:00+02:00')
        self.assertEqual(computer['ModificationUtcDateTime'], '2026-01-14T13:00:00+00:00')
        self.assertEqual(computer['ModificationStatus'], 'UTC qualified')

    def test_entra_lifecycle_uses_only_explicit_object_dates(self):
        self.fixture.inputs['users'][0]['lastModifiedDateTime'] = '2026-01-13T15:00:00Z'
        self.fixture.inputs['entra_devices'][0].update(createdDateTime='2020-01-01T10:00:00Z',
            ModifiedDateTime='2026-01-13T17:00:00+02:00')
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        user = self.output_rows('DimUser')[0]
        device = next(row for row in self.output_rows('DimDevice') if row['SourceDeviceId'] == 'ed1')
        self.assertEqual(user['ModificationRaw'], '2026-01-13T15:00:00Z')
        self.assertEqual(user['ModificationUtcDateTime'], '2026-01-13T15:00:00+00:00')
        self.assertEqual(device['CreationRaw'], '2020-01-01T10:00:00Z')
        self.assertEqual(device['CreationUtcDateTime'], '2020-01-01T10:00:00+00:00')
        self.assertEqual(device['ModificationUtcDateTime'], '2026-01-13T15:00:00+00:00')

    def test_roundtrip_companions_reach_lifecycle_and_activity_without_replacing_raw(self):
        self.fixture.seed_ad()
        self.fixture.inputs['users'][0].update({
            'When created':'05/10/2020 09:08:07',
            'When createdIso8601':'2020-10-05T09:08:07.1234567+02:00',
            'LastSignInDateTime':'14/01/2026 15:00:00',
            'LastSignInDateTimeIso8601':'2026-01-14T15:00:00+02:00'})
        for name in ('ad_users','ad_computers'):
            self.fixture.inputs[name][0].update(
                WhenCreated='05/10/2020 09:08:07', WhenCreatedIso8601='2020-10-05T09:08:07Z',
                WhenChanged='14/01/2026 15:00:00', WhenChangedIso8601='2026-01-14T15:00:00+02:00')
        self.fixture.inputs['ad_users'][0]['PasswordExpired'] = 'false'
        self.fixture.inputs['ad_computers'][0]['LastLogonTimestamp'] = '134127828000000000'
        self.fixture.inputs['autopilot'] = [{'Autopilot ID':'ap1',
            'Last contacted':'14/01/2026 15:00:00',
            'Last contactedIso8601':'2026-01-14T15:00:00+02:00'}]
        self.fixture.inputs['policies'] = [dict(PolicyId='p'+str(i), PolicyFamily=family,
            CreatedDateTime='05/10/2020 09:08:07', LastModifiedDateTime='14/01/2026 15:00:00',
            CreatedDateTimeIso8601='2020-10-05T09:08:07+02:00',
            LastModifiedDateTimeIso8601='2026-01-14T15:00:00Z',
            NativeEvidenceJson='{}', AssignmentCollectionStatus='Collected') for i, family in enumerate(
            ('SettingsCatalog', 'DeviceConfiguration', 'DeviceCompliance', 'WindowsFeatureUpdate', 'WindowsQualityUpdate'))]
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        user = self.output_rows('DimUser')[0]
        self.assertEqual(user['CreationRaw'], '05/10/2020 09:08:07')
        self.assertEqual(user['CreationUtcDateTime'], '2020-10-05T07:08:07.123456+00:00')
        self.assertEqual(user['LastSignInDateTime'], '2026-01-14T15:00:00+02:00')
        for name in ('ADUserSource','ADComputerSource'):
            row = self.output_rows(name)[0]
            self.assertEqual(row['CreationUtcDateTime'], '2020-10-05T09:08:07+00:00')
            self.assertEqual(row['ModificationRaw'], '14/01/2026 15:00:00')
            self.assertEqual(row['ModificationUtcDateTime'], '2026-01-14T13:00:00+00:00')
            self.assertEqual(row['ModificationStatus'], 'UTC qualified')
        self.assertEqual(self.output_rows('ADUserSource')[0]['PasswordExpiryStatus'], 'Observed')
        self.assertEqual(self.output_rows('ADUserSource')[0]['PasswordExpired'], 'false')
        self.assertTrue(self.output_rows('ADComputerSource')[0]['LastLogonUtcDateTime'])
        self.assertEqual(self.output_rows('FactAutopilotDevice')[0]['LastContactedDateTime'],
                         '2026-01-14T15:00:00+02:00')
        for table in ('DimIntuneConfigurationPolicy','DimWindowsUpdatePolicy'):
            for policy in self.output_rows(table):
                self.assertEqual(policy['CreatedDateTime'], '2020-10-05T09:08:07+02:00')
                self.assertEqual(policy['LastModifiedDateTime'], '2026-01-14T15:00:00Z')

    def test_lifecycle_companion_failures_do_not_fall_back_to_qualified_legacy_dates(self):
        from workplace_native_tables import lifecycle_date, qualified
        for raw, status in [('2026-01-14T15:00:00','Timezone not established'),
                            ('14/01/2026 15:00:00','Timezone not established'),
                            ('31/02/2026 15:00:00','Invalid date format'),
                            ('0001-01-01T00:00:00','Sentinel / unknown')]:
            with self.subTest(raw=raw):
                result = lifecycle_date({'WhenChanged':'2026-01-14T15:00:00Z',
                                         'WhenChangedIso8601':raw}, 'Modification', 'WhenChanged')
                self.assertEqual(result['ModificationRaw'], '2026-01-14T15:00:00Z')
                self.assertEqual(result['ModificationUtcDateTime'], '')
                self.assertEqual(result['ModificationStatus'], status)
        self.assertEqual(qualified(''), ('','Not provided'))
        self.assertEqual(qualified('01/01/1601 00:00:00'), ('','Sentinel / unknown'))
        # Acquisition parsing must remain strict even though business dates accept legacy text.
        for value in ('2026-01-14T15:00:00','14/01/2026 15:00:00'):
            with self.assertRaises(ValueError):
                sources.utc(value)

    def test_entra_missing_lifecycle_does_not_substitute_sync_enrollment_or_password_dates(self):
        self.fixture.inputs['users'][0]['LastPasswordChangeDateTime'] = '2026-01-13T15:00:00Z'
        self.fixture.inputs['entra_devices'][0].update(RegistrationDateTime='2020-01-01T10:00:00Z',
            ApproximateLastSignInDateTime='2026-01-13T15:00:00Z')
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        user = self.output_rows('DimUser')[0]
        self.assertEqual((user['ModificationRaw'],user['ModificationUtcDateTime'],user['ModificationStatus']), ('','','Not provided'))
        for device in self.output_rows('DimDevice'):
            for prefix in ('Creation','Modification'):
                self.assertEqual((device[prefix+'Raw'],device[prefix+'UtcDateTime'],device[prefix+'Status']), ('','','Not provided'))

    def test_lifecycle_changes_keep_native_keys_and_source_files_unchanged(self):
        self.fixture.seed_ad()
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        keys = {'ADUserSource':'TenantADUserKey','ADComputerSource':'TenantADComputerKey',
                'DimUser':'TenantUserKey','DimDevice':'TenantDeviceKey'}
        previous = {name:[row[field] for row in self.output_rows(name)] for name,field in keys.items()}
        for name in ('ad_users','ad_computers'):
            self.fixture.inputs[name][0].update(WhenChanged='2026-01-13T15:00:00Z',PasswordExpired='true')
        self.fixture.inputs['users'][0]['lastModifiedDateTime'] = '2026-01-13T15:00:00Z'
        self.fixture.inputs['entra_devices'][0]['CreatedDateTime'] = '2020-01-01T00:00:00Z'
        self.fixture.write_inputs()
        before = {path.name:sources.sha(path) for path in self.fixture.source.iterdir() if path.is_file()}
        self.output = self.fixture.root / 'changed-lifecycle-staging'
        self.output.mkdir()
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertEqual(previous,{name:[row[field] for row in self.output_rows(name)] for name,field in keys.items()})
        self.assertEqual(before,{path.name:sources.sha(path) for path in self.fixture.source.iterdir() if path.is_file()})

    def test_teams_channels_sites_preserve_direct_links(self):
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
        links=self.output_rows('TeamsSiteLinkEvidence')
        self.assertEqual(len(links),4)
        self.assertEqual(len({r['Link Key'] for r in links}),4)
        self.assertTrue(any(r['Site ID']=='host,outside,w' for r in links))
        self.assertEqual(self.output_rows('SharePointSiteLinkEvidence')[0]['Link Count'],'2')
        self.assertEqual(before,{p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()})

    def output_rows(self, name):
        with (self.output / (name+'.csv')).open(encoding='utf-8-sig', newline='') as stream:
            return list(csv.DictReader(stream))

    def test_no_generic_graph_outputs_and_native_keys_remain(self):
        result=unified.build(self.fixture.source,self.output,self.audit())
        self.assertNotIn('GraphRows',result)
        self.assertFalse(list(self.output.glob('CMDB*.csv')))
        for table in sources.load_json(sources.CONTRACT)['tables']:
            rows=self.output_rows(table['name'])
            if table.get('key'):
                self.assertEqual(len(rows),len({tuple(r[k] for k in table['key']) for r in rows}),table['name'])

    def test_direct_and_primary_group_membership_preserved(self):
        self.fixture.seed_ad();self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        members=self.output_rows('ADMembership')
        self.assertEqual({r['MembershipKind'] for r in members},{'Direct','PrimaryGroup'})
        self.assertEqual(len(members),2)
        self.assertFalse(any('transitive' in r['MembershipKind'].lower() for r in members))

    def test_unknown_application_and_device_remain_observed_not_fabricated(self):
        self.fixture.inputs['app_relations'].append({'AppId':'missing-app','DeviceId':'missing-device'})
        self.fixture.write_inputs();unified.build(self.fixture.source,self.output,self.audit())
        missing=[r for r in self.output_rows('FactDeviceApplication') if r['ManagedDeviceId']=='missing-device']
        self.assertEqual(len(missing),1)
        self.assertEqual(missing[0]['AppId'],'missing-app')
        self.assertFalse(any(r['ManagedDeviceId']=='missing-device' for r in self.output_rows('DimIntuneManagedDevice')))

    def test_manager_identity_preserved(self):
        self.fixture.seed_ad();self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertTrue(any(r['ManagerADObjectKey'] for r in self.output_rows('ADUserSource')))

    def test_missing_manager_stays_unresolved(self):
        self.fixture.seed_ad()
        self.fixture.inputs['ad_users'][0]['manager']='CN=Absent,DC=synthetic,DC=invalid'
        self.fixture.write_inputs();unified.build(self.fixture.source,self.output,self.audit())
        users=[r for r in self.output_rows('ADUserSource') if r['ManagerDistinguishedName']=='CN=Absent,DC=synthetic,DC=invalid']
        self.assertEqual(len(users),1)
        self.assertEqual(users[0]['ManagerADObjectKey'],'')

    def test_nested_groups_and_non_user_members_preserved(self):
        self.fixture.inputs['group_scope'][0]['MemberCount']='4'
        self.fixture.inputs['group_members']=[dict(GroupId='g1',MemberId=mid,MemberType=kind,MembershipKind='Direct',CollectionStatus='Collected')
            for mid,kind in [('u1','#microsoft.graph.user'),('eo1','#microsoft.graph.device'),('g1','#microsoft.graph.group'),('sp1','#microsoft.graph.servicePrincipal')]]
        self.fixture.write_inputs();unified.build(self.fixture.source,self.output,self.audit())
        members=self.output_rows('EntraGroupMembership')
        self.assertEqual(len(members),4)
        self.assertEqual({r['MemberId'] for r in members},{'u1','eo1','g1','sp1'})
        self.assertTrue(all(r['MembershipKind']=='Direct' for r in members))
        self.assertEqual(len(members),len({r['TenantMembershipKey'] for r in members}))

    def test_managed_device_ids_not_merged_by_display_name(self):
        unified.build(self.fixture.source,self.output,self.audit())
        devices=self.output_rows('DimIntuneManagedDevice')
        self.assertEqual(len(devices),2)
        self.assertEqual(len({r['ManagedDeviceId'] for r in devices}),2)

    def test_unresolved_application_links_are_preserved(self):
        self.fixture.inputs['app_relations'].append({'AppId': 'missing-app', 'DeviceId': 'missing-device'})
        self.fixture.write_inputs()
        self.compare()

    def test_unresolved_license_users_are_preserved(self):
        self.fixture.inputs['license_paths'].append({'UserId':'missing-user','SkuId':'s1','AssignedByGroupId':'g1','AssignmentState':'Error'})
        self.fixture.write_inputs()
        self.compare()

    def test_source_health_statuses_are_mapped_to_collected_states(self):
        from test_native_source_health import mapping
        self.fixture.inputs['license_paths'].append({'UserId':'missing-user','SkuId':'s1','AssignedByGroupId':'g1','AssignmentState':'Error'})
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        statuses={r['Status'] for r in self.output_rows('SourceHealth')}
        self.assertEqual(statuses,{'Success','SuccessWithWarnings'})
        self.assertEqual({mapping()[s] for s in statuses},{'OK','Warning'})

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
        self.assertEqual(self.output_rows('DimUser')[0]['LastSignInDateTime'], '2026-10-05T09:08:07')
        self.assertEqual(self.output_rows('DimLicenseSku')[0]['ConsumedUnits'], '8794')
        self.assertEqual(self.output_rows('DimLicenseSku')[0]['EnabledUnits'], '9000')
        for table in ('DimIntuneConfigurationPolicy', 'DimWindowsUpdatePolicy'):
            self.assertEqual(self.output_rows(table)[0]['CreatedDateTime'], '2025-07-17T12:00:23')
            self.assertEqual(self.output_rows(table)[0]['LastModifiedDateTime'], '2026-10-05T09:08:07')
        self.assertEqual(self.output_rows('FactAutopilotDevice')[0]['LastContactedDateTime'], '2026-10-14T09:00:00')

    def test_ad_account_classification_is_carried_without_reclassification(self):
        self.fixture.seed_ad()
        self.fixture.inputs['ad_users'][0]['AccountType']='Service Account'
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertEqual(self.output_rows('ADUserSource')[0]['AccountType'],'Service Account')

    def test_absent_ad_account_classification_stays_unknown(self):
        self.fixture.seed_ad()
        for row in self.fixture.inputs['ad_users']:row.pop('AccountType',None)
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        self.assertTrue(all(row['AccountType']=='' for row in self.output_rows('ADUserSource')))

    def test_all_account_types_and_disabled_accounts_are_retained(self):
        self.fixture.seed_ad()
        template=self.fixture.inputs['ad_users'][0]
        types=['Named Account','Ext Account','Generic Account','Admin Account','System Account',
            'Resource Account','Service Account','Unclassified Account','Future Custom Type']
        self.fixture.inputs['ad_users']=[dict(template,ObjectGUID='ad-type-'+str(i),
            ObjectSID='S-1-5-21-1-2-3-'+str(500+i),Enabled='false',AccountType=value)
            for i,value in enumerate(types)]
        self.fixture.write_inputs()
        unified.build(self.fixture.source,self.output,self.audit())
        actual=self.output_rows('ADUserSource')
        self.assertCountEqual([row['AccountType'] for row in actual],types)
        self.assertTrue(all(row['Enabled']=='false' for row in actual))

    def test_type_failure_precedes_collaboration_and_leaves_sources_unchanged(self):
        self.fixture.inputs['skus'][0]['TenantConsumedUnits'] = '1.5'
        self.fixture.write_inputs()
        before = {p.name:sources.sha(p) for p in self.fixture.source.iterdir() if p.is_file()}
        with mock.patch.object(unified.teams_sites, 'build') as collaboration:
            with self.assertRaisesRegex(ValueError, 'int64'):
                unified.build(self.fixture.source, self.output, self.audit())
            collaboration.assert_not_called()
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
            captured=[c['source'] for c in mapped.get('capturedOnlyColumns',[])]
            self.assertEqual(len(captured),len(set(captured)))
            self.assertTrue(set(captured).issubset(set(table['columns'])))
            self.assertEqual([name for name in table['columns'] if name not in captured],
                [c['source'] for c in mapped['columns']])
            for column in mapped.get('capturedOnlyColumns',[]):
                self.assertTrue(column['reader']);self.assertTrue(column['reason'])
            self.assertTrue(mapped['test'])


if __name__ == '__main__':
    unittest.main()
