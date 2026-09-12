"""Synthetic V1 tests of 360 grain, isolation, provenance and date handling."""
import contextlib
import csv
import io
import json
import unittest
from pathlib import Path
import test_report_builder as fixtures
report, PRODUCT = fixtures.report, fixtures.PRODUCT


class Report360Tests(unittest.TestCase):
    setUpBase = fixtures.ReportTests.setUp
    tearDown = fixtures.ReportTests.tearDown
    path = fixtures.ReportTests.path
    write = fixtures.ReportTests.write

    def setUp(self):
        self.setUpBase()
        self.raw_contract = {Path(t['name']).stem:t for t in json.loads((PRODUCT/'Schema/SmartWorkplaceCMDB.raw.tables.json').read_text())['tables']}
        self.now='2026-09-09T20:00:00Z'
        self.write('DimUser',[
            dict(TenantUserKey='fictional-prod|u1',CmdbUserId='fictional-prod|u1',DisplayName='Same name',UserPrincipalName='one@example.invalid',AccountEnabled='true',CountryCode='FR',CountryLabel='FR',CountryStatus='Reported'),
            dict(TenantUserKey='fictional-prod|u2',CmdbUserId='fictional-prod|u2',DisplayName='Same name',UserPrincipalName='two@example.invalid',AccountEnabled='false',CountryLabel='Unknown / unassigned',CountryStatus='Not provided')])
        self.write('CMDB_Users',[
            dict(CmdbUserId='fictional-prod|u1',SourceUserId='u1',UsageLocation='FR',UsageLocationStatus='Reported',CreatedDateTime='09/01/2026 10:00:00',SourceCollectedDateTime=self.now),
            dict(CmdbUserId='fictional-prod|u2',SourceUserId='u2',UsageLocationStatus='Not provided',CreatedDateTime='2026-09-01T12:00:00+02:00',SourceCollectedDateTime=self.now)])
        self.write('DimDevice',[
            dict(TenantDeviceKey='fictional-prod|d1',CmdbDeviceId='fictional-prod|d1',DeviceName='Same device',ComplianceState='compliant'),
            dict(TenantDeviceKey='fictional-prod|d2',CmdbDeviceId='fictional-prod|d2',DeviceName='Same device')])
        self.write('CMDB_Devices',[
            dict(CmdbDeviceId='fictional-prod|d1',SourceDeviceId='d1',PrimaryUserId='u1',LastSyncDateTime=self.now,SourceCollectedDateTime=self.now),
            dict(CmdbDeviceId='fictional-prod|d2',SourceDeviceId='d2',PrimaryUserId='absent-user',LastSyncDateTime='0001-01-01T00:00:00Z',SourceCollectedDateTime=self.now)])
        self.write('DimGroup',[dict(TenantGroupKey='fictional-prod|g1',CmdbGroupId='fictional-prod|g1',DisplayName='Fictional group',MailEnabled='false',SecurityEnabled='true')])
        self.write('CMDB_Groups',[dict(CmdbGroupId='fictional-prod|g1',SourceGroupId='g1',SourceCollectedDateTime=self.now)])
        self.write('DimLicenseSku',[
            dict(TenantSkuKey='fictional-prod|s1',SkuId='s1',SkuPartNumber='SKU-A',EnabledUnits='20',ConsumedUnits='1'),
            dict(TenantSkuKey='fictional-prod|s2',SkuId='s2',SkuPartNumber='SKU-B',EnabledUnits='30',ConsumedUnits='0')])
        self.write('FactUserDeviceRelationship',[dict(TenantRelationshipKey='fictional-prod|r1',TenantUserKey='fictional-prod|u1',TenantDeviceKey='fictional-prod|d1')])
        self.write_raw('Entra_Devices',[
            dict(SourceObjectId='entra1',SourceDeviceId='d1',DeviceName='Same device',ApproximateLastSignInDateTime='09/01/2026 10:00:00',SourceCollectedDateTime=self.now),
            dict(SourceObjectId='entra2',SourceDeviceId='d2',DeviceName='Same device',SourceCollectedDateTime=self.now)])
        self.write_raw('Intune_ManagedDevices',[
            dict(ManagedDeviceId='intune1',AzureAdDeviceId='d1',DeviceName='Same device',UserId='u1',LastSyncDateTime=self.now,
                 EnrolledDateTime='2026-09-01T12:00:00+02:00',DeviceEnrollmentType='windowsAzureADJoin',ManagementAgent='mdm',SourceCollectedDateTime=self.now)])
        self.paths=[dict(RawAssignmentKey='path1',SourceUserId='u1',SkuId='s1',AssignmentState='Active',AssignmentError='None',SourceCollectedDateTime=self.now),
                    dict(RawAssignmentKey='path2',SourceUserId='u1',SkuId='s1',AssignedByGroupId='g1',AssignmentState='Error',AssignmentError='CountViolation',
                         DisabledPlanIds='plan1;plan2',LastUpdatedDateTime=self.now,SourceCollectedDateTime=self.now)]
        self.write_raw('M365_UserLicenseAssignments',self.paths)

    def write_raw(self,name,rows):
        t=self.raw_contract[name]
        p=self.root.joinpath(*t['area'].replace('\\','/').split('/'),t['name'])
        p.parent.mkdir(parents=True,exist_ok=True)
        with p.open('w',encoding='utf-8-sig',newline='') as f:
            w=csv.DictWriter(f,fieldnames=t['columns']);w.writeheader()
            for row in rows:w.writerow(dict({c:'' for c in t['columns']},**self.identity)|row)

    def enriched(self):
        identity,data,columns,hashes=report.prepare_data(self.root)
        rels=report.enrich_360(self.root,identity,data,columns,hashes)
        return data,columns,rels

    def test_multiple_paths_preserved_without_changing_summary(self):
        _,base,_,_=report.prepare_data(self.root)
        old=report.measures(base)
        data,cols,rels=self.enriched()
        self.assertEqual(len(data['LicenseAssignmentPath']),2)
        self.assertEqual(len(data['FactUserLicense']),1)
        self.assertEqual(report.measures(data)[:23],old)
        names = {measure['name'] for measure in report.measures(data)}
        self.assertTrue({'Managed device share', 'Source collection coverage',
                         'Assignment error rate', 'Mailbox link gap rate',
                         'Corporate devices', 'Corporate device country share'} <= names)
        self.assertEqual(len(rels),7)
        self.assertEqual(data['LicenseAssignmentPath'][1]['AssignmentError'],'CountViolation')
        self.assertEqual(data['LicenseAssignmentPath'][1]['GroupName'],'Fictional group')
        self.assertEqual(data['LicenseAssignmentPath'][1]['DisabledPlanIds'],'plan1;plan2')

    def test_date_ambiguity_offsets_and_sentinel_are_not_guessed(self):
        data,_,_=self.enriched()
        self.assertEqual(data['DimUser'][0]['CreationUtcDateTime'],'')
        self.assertEqual(data['DimUser'][0]['CreationRaw'],'09/01/2026 10:00:00')
        self.assertEqual(data['DimUser'][1]['CreationUtcDateTime'],'2026-09-01T10:00:00+00:00')
        self.assertEqual(data['DimDevice'][1]['SyncUtcDateTime'],'')
        self.assertEqual(data['DimDevice'][1]['SyncStatus'],'Sentinel / unknown')
        self.assertEqual(data['DeviceSource'][0]['ActivityUtcDateTime'],'')
        self.assertEqual(data['DeviceSource'][-1]['EnrollmentUtcDateTime'],'2026-09-01T10:00:00+00:00')

    def test_same_display_names_remain_distinct(self):
        data,_,_=self.enriched()
        for name,column in [('DimUser','UserSelection'),('DimDevice','DeviceSelection')]:
            self.assertEqual(len(set(r[column] for r in data[name])),2)
        self.assertEqual(data['DimDevice'][1]['AssociatedAccount'],'Not resolved')
        self.assertEqual(data['DimDevice'][1]['AssociationStatus'],'Unresolved')
        self.assertEqual(data['DimDevice'][0]['CountryLabel'],'FR')
        self.assertEqual(data['DimDevice'][0]['CountryStatus'],'Reported')
        self.assertEqual(data['DimDevice'][1]['CountryLabel'],'Unknown / unassigned')

    def test_missing_association_is_not_unresolved(self):
        rows=report.read_csv(self.path('CMDB_Devices'))[1];rows[1]['PrimaryUserId']=''
        self.write('CMDB_Devices',rows)
        data,_,_=self.enriched()
        self.assertEqual(data['DimDevice'][1]['AssociationStatus'],'Not provided')
        self.assertEqual(data['DimDevice'][1]['CountryStatus'],'No primary user')

    def test_entity_findings_attach_only_by_exact_type_and_key(self):
        self.write('CMDB_DataQuality',[dict(FindingId='q1',Severity='Warning',EntityType='Device',EntityId='fictional-prod|d1')])
        data,_,_=self.enriched()
        f=data['EntityFinding'][0]
        self.assertEqual(f['TenantDeviceKey'],'fictional-prod|d1')
        self.assertEqual(f['TenantUserKey'],'')
        self.assertEqual(f['TenantGroupKey'],'')

    def test_foreign_cmdb_row_rejected(self):
        rows=report.read_csv(self.path('CMDB_Users'))[1];rows[0]['TenantId']='foreign'
        self.write('CMDB_Users',rows)
        with self.assertRaisesRegex(ValueError,'360 source tenant mismatch'):self.enriched()

    def test_foreign_raw_row_rejected(self):
        self.paths[0]['TenantKey']='foreign'
        self.write_raw('M365_UserLicenseAssignments',self.paths)
        with self.assertRaisesRegex(ValueError,'tenant mismatch'):self.enriched()

    def test_duplicate_raw_path_rejected(self):
        self.write_raw('M365_UserLicenseAssignments',self.paths+[self.paths[0]])
        with self.assertRaisesRegex(ValueError,'duplicate 360 key'):self.enriched()

    def test_missing_raw_required_for_360_rejected_before_output(self):
        (self.root/'Raw/Intune/Intune_ManagedDevices.csv').unlink()
        with self.assertRaisesRegex(ValueError,'requires existing local source'):
            report.build(self.root,self.base/'new',True)
        self.assertFalse((self.base/'new').exists())

    def test_multiple_intune_candidates_retained_and_ranked(self):
        original=report.read_csv(self.root/'Raw/Intune/Intune_ManagedDevices.csv')[1][0]
        other=dict(original,ManagedDeviceId='intune0',LastSyncDateTime='2026-09-02T00:00:00Z')
        self.write_raw('Intune_ManagedDevices',[other,original])
        data,_,_=self.enriched()
        rows=[r for r in data['DeviceSource'] if r['SourceSystem']=='Intune']
        self.assertEqual(len(rows),2)
        self.assertEqual([r['SelectionStatus'] for r in rows],['Other candidate','Selected'])
        self.assertEqual(len(data['DimDevice']),2)

    def test_source_selection_drift_rejected(self):
        rows=report.read_csv(self.root/'Raw/Intune/Intune_ManagedDevices.csv')[1]
        rows[0]['UserId']='u2';self.write_raw('Intune_ManagedDevices',rows)
        with self.assertRaisesRegex(ValueError,'selected source differs'):self.enriched()

    def test_orphan_group_is_retained_without_a_false_relationship(self):
        self.paths[1]['AssignedByGroupId']='missing-group';self.write_raw('M365_UserLicenseAssignments',self.paths)
        data,_,_=self.enriched()
        self.assertEqual(data['LicenseAssignmentPath'][1]['TenantGroupKey'],'')
        self.assertEqual(data['LicenseAssignmentPath'][1]['GroupLinkStatus'],'Unresolved')
        self.assertEqual(data['LicenseAssignmentPath'][1]['AssignedByGroupId'],'missing-group')

    def test_inconsistent_summary_pair_set_rejected(self):
        self.paths[0]['SkuId']='s2';self.write_raw('M365_UserLicenseAssignments',self.paths)
        with self.assertRaisesRegex(ValueError,'reconcile with summary'):self.enriched()

    def test_nine_pages_bindings_and_source_immutability(self):
        hashes={p:report.sha(p) for p in self.root.rglob('*') if p.is_file()}
        output=self.base/'new'
        with contextlib.redirect_stdout(io.StringIO()):report.build(self.root,output,True)
        model=json.loads((output/'CMDB-REPORTS.SemanticModel/model.bim').read_text())['model']
        self.assertEqual(len(model['tables']),23)
        self.assertEqual(len(model['relationships']),13)
        self.assertTrue(all(r['crossFilteringBehavior']=='oneDirection' for r in model['relationships']))
        manifest=json.loads((output/'REPORT-MANIFEST.json').read_text())
        self.assertEqual(len(manifest['pages']),9)
        for name in ['device360','user360','group360']:
            page=json.loads((output/f'CMDB-REPORTS.Report/definition/pages/{name}/page.json').read_text())
            self.assertEqual(page['pageBinding']['type'],'Drillthrough')
            self.assertEqual(page['pageBinding']['parameters'][0]['boundFilter'],page['filterConfig']['filters'][0]['name'])
        self.assertEqual(hashes,{p:report.sha(p) for p in self.root.rglob('*') if p.is_file()})
        gates=[m for m in manifest['measures'] if m['name'].endswith('details')]
        self.assertEqual(len(gates),11)
        self.assertTrue(all('ALLSELECTED' in m['expression'] for m in gates))
        self.assertIsNone(next(m['expected'] for m in gates if m['name']=='Device details'))

    def test_report_captions_and_dropdown_body_space(self):
        data,_,_=self.enriched()
        pages=report.report_pages(report.measures(data),True)
        for page in pages:
            for v in page['visuals']:
                visual=v['visual']
                for role in visual.get('query',{}).get('queryState',{}).values():
                    for p in role['projections']:
                        self.assertTrue(p.get('displayName'), (page['name'],v['name']))
                if visual['visualType']=='slicer':
                    # 12pt title (~20px), top/bottom padding, and a full 32px
                    # native dropdown must fit before the next visual at 224.
                    padding=visual['visualContainerObjects']['padding'][0]['properties']
                    vertical=sum(float(padding[k]['expr']['Literal']['Value'].rstrip('D')) for k in ('top','bottom'))
                    self.assertGreaterEqual(v['position']['height']-vertical-20,32)
                    self.assertLessEqual(v['position']['y']+v['position']['height'],220)
        identity=next(p for p in pages if p['name']=='device360')['visuals'][5]
        headers=[p['displayName'] for p in identity['visual']['query']['queryState']['Values']['projections']]
        self.assertEqual(headers[:3],['Device','System','Version'])


if __name__=='__main__':unittest.main()
