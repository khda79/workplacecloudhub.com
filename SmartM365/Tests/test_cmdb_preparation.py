"""Synthetic, offline acceptance tests. Never read synchronized tenant data."""
import copy
import csv
import datetime as dt
import json
import sys
import tempfile
import unittest
import shutil
import uuid
import subprocess
from unittest import mock
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'SmartInventory/PreparedEvidence'))
import cmdb_prepare as pipeline

NOW = dt.datetime(2026, 1, 15, 12, tzinfo=dt.timezone.utc)
IDENTITY = {'TenantKey':'synthetic','OrganizationKey':'test','EnvironmentKey':'test','TenantId':'synthetic-tenant'}


class PreparationTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.gettempdir()) / ('SmartInventory-Cmdb-Tests-'+uuid.uuid4().hex)
        self.root.mkdir()
        self.source = self.root / 'DATA-LAST'; self.source.mkdir()
        self.output = self.root / 'DATA-POWERBI-CMDB'
        self.contract = pipeline.load_json(pipeline.CONTRACT)
        self.inputs = {s['name']:[] for s in self.contract['sources']}
        self.inputs.update({
            'users':[{'Object Id':'u1','User principal name':'user@synthetic.invalid','Display name':'Test user',
                      'AccountEnabled':'true','CollectedAtUtc':NOW.isoformat(),'Usage location':'FR','UserType':'Member'}],
            'entra_devices':[{'ObjectId':'eo1','DeviceId':'ed1','DisplayName':'Same name','OperatingSystem':'Windows'}],
            'managed':[{'ManagedDeviceId':'md1','AzureADDeviceId':'ed1','DeviceName':'Same name','OperatingSystem':'Windows',
                       'CollectedAtUtc':NOW.isoformat(),'UserId':'u1','LastSyncDateTime':NOW.isoformat(),
                       'ManagedDeviceOwnerType':'company','ComplianceState':'compliant','OsVersion':'10.0.26100'},
                      {'ManagedDeviceId':'md2','AzureADDeviceId':'ed2','DeviceName':'Same name','OperatingSystem':'Android',
                       'CollectedAtUtc':NOW.isoformat(),'UserId':'u1','LastSyncDateTime':NOW.isoformat(),'ManagedDeviceOwnerType':'personal'}],
            'apps':[{'AppId':'a1','AppName':'Editor','AppPublisher':'Vendor','Platform':'windows','AppVersion':'1','DeviceCount':'1',
                     'CollectionScope':'AllPlatforms','RelationCollectionScope':'All'},
                    {'AppId':'a2','AppName':' editor ','AppPublisher':'vendor','Platform':'WINDOWS','AppVersion':'2','DeviceCount':'1',
                     'CollectionScope':'AllPlatforms','RelationCollectionScope':'All'}],
            'app_relations':[{'AppId':'a1','DeviceId':'md1'},{'AppId':'a2','DeviceId':'md1'}],
            'skus':[{'Id':'s1','TenantSkuPartNumber':'SPE_E3','TenantPrepaidEnabled':'','TenantConsumedUnits':''}],
            'license_paths':[{'UserId':'u1','SkuId':'s1','AssignedByGroupId':'','AssignmentState':'Active','AssignmentError':''}],
            'plans':[{'SkuId':'s1','PlanId':'p1','PlanName':'Test plan','TenantProvisioningStatus':'Success'}],
            'user_plans':[{'UserId':'u1','SkuId':'s1','PlanId':'p1','StateCode':'D'}],
            'groups':[{'GroupId':'g1','DisplayName':'Group'}],
            'group_scope':[{'GroupId':'g1','MemberCount':'0','MemberCollectionStatus':'Collected'}],
            'mailboxes':[{'ExternalDirectoryObjectId':'u1','MailboxGuid':'ex1','PrimarySmtpAddress':'user@synthetic.invalid',
                          'RecipientTypeDetails':'UserMailbox','CollectedAtUtc':NOW.isoformat()}],
            'local_mailboxes':[{'ObjectGUID':'local1','PrimarySMTPaddress':'USER@synthetic.invalid','RecipientType':'UserMailbox',
                                'CollectedAtUtc':NOW.isoformat()}],
            'remote_mailboxes':[{'ObjectGuid':'remote1','PrimarySmtpAddress':'user@synthetic.invalid','RecipientTypeDetails':'RemoteUserMailbox',
                                 'CollectedAtUtc':NOW.isoformat()}],
            'ad_computers':[{'ObjectGUID':'ad-pc1','OperatingSystem':'Windows 10','Name':'AD unmanaged','Enabled':'true'}],
            'analytics':[{'ReportName':'Score','DeviceId':'md1','EndpointAnalyticsScore':'70'}],
            'readiness':[{'GraphId':'md1','UpgradeEligibility':'notCapable'}],
        })
        self.inputs['hardware']=[dict(ManagedDeviceId=r['ManagedDeviceId'],azureADDeviceId=r['AzureADDeviceId'],
            serialNumber=r['ManagedDeviceId'],manufacturer='Vendor',model='Test',totalStorageSpaceInBytes='107374182400',
            totalStorageSpaceInBytesStatus='Reported',freeStorageSpaceInBytes='1073741824',freeStorageSpaceInBytesStatus='Reported',
            physicalMemoryInBytes='8589934592',physicalMemoryInBytesStatus='Reported',CollectionStatus='Collected',CollectedAtUtc=NOW.isoformat())
            for r in self.inputs['managed']]
        self.write_inputs()

    def tearDown(self):
        if self.root.parent.resolve()==Path(tempfile.gettempdir()).resolve() and self.root.name.startswith('SmartInventory-Cmdb-Tests-'):
            shutil.rmtree(self.root)

    def write_inputs(self):
        files=[]
        for source in self.contract['sources']:
            data=[dict(row,TenantKey='synthetic') for row in self.inputs[source['name']]]
            columns=list(dict.fromkeys(source['columns']+[col for row in data for col in row]))
            path=self.source/source['file']
            with path.open('w',encoding='utf-8-sig',newline='') as stream:
                writer=csv.DictWriter(stream,fieldnames=columns);writer.writeheader();writer.writerows(data)
            files.append({'File':path.name,'Rows':len(data),'SHA256':pipeline.sha(path),
                          'Status':'Success','Errors':0,'IsPartialInventory':False,'Producer':'Synthetic producer',
                          'ScriptVersion':'test','RunId':'synthetic-run','StartedAtUtc':(NOW-dt.timedelta(minutes=5)).isoformat(),
                          'CompletedAtUtc':NOW.isoformat()})
        self.proof={'ContractVersion':'1.1',**IDENTITY,'Status':'Completed','IsPartialInventory':False,
                    'Errors':0,'RunId':'synthetic-run','Files':files}
        for producer in pipeline.load_json(pipeline.REGISTRY)['Producers']:
            for record in files:
                if record['File'] in producer['Files']:
                    record['Producer']=producer['Script'];record['Scope']=producer['Scope']
        self.write_proof()

    def write_proof(self):
        for producer in pipeline.load_json(pipeline.REGISTRY)['Producers']:
            files=[r for r in self.proof['Files'] if r['File'] in producer['Files']]
            proof=dict(self.proof, Owner='SmartInventory-CmdbSourceReceipt',
                       Producer=producer['Script'], ScriptVersion='test', Scope=producer['Scope'],
                       StartedAtUtc=files[0]['StartedAtUtc'] if files else (NOW-dt.timedelta(minutes=5)).isoformat(),
                       CompletedAtUtc=NOW.isoformat(), Files=files)
            (self.source/producer['Receipt']).write_text(json.dumps(proof),encoding='utf-8')

    def prepare(self, **kw):
        return pipeline.prepare(self.source,self.output,'synthetic',IDENTITY,now=NOW,**kw)

    def table(self,name):
        return list(pipeline.rows(self.output/(name+'.csv')))

    def write_native_rows(self, name, data):
        # Independent producer headers: never fabricate missing columns from the contract.
        definition=next(s for s in self.contract['sources'] if s['name']==name)
        path=self.source/definition['file']
        data=[dict(TenantKey='synthetic',**row) for row in data]
        columns=list(dict.fromkeys(column for row in data for column in row))
        with path.open('w',encoding='utf-8-sig',newline='') as stream:
            writer=csv.DictWriter(stream,fieldnames=columns);writer.writeheader();writer.writerows(data)
        record=next(r for r in self.proof['Files'] if r['File']==path.name)
        record.update(Rows=len(data),SHA256=pipeline.sha(path))
        self.write_proof()

    def alter_receipt(self, **updates):
        path=self.source/pipeline.load_json(pipeline.REGISTRY)['Producers'][0]['Receipt']
        proof=pipeline.load_json(path);proof.update(updates)
        path.write_text(json.dumps(proof),encoding='utf-8')

    def test_running_producer_blocks_last_output_replacement(self):
        self.unchanged_after(lambda:(self.alter_receipt(Status='Running'),self.prepare()),'Incomplete producer')

    def test_foreign_parent_identity_blocks_replacement(self):
        self.unchanged_after(lambda:(self.alter_receipt(TenantId='foreign'),self.prepare()),'tenant identity')

    def test_wrong_producer_parent_blocks_replacement(self):
        self.unchanged_after(lambda:(self.alter_receipt(Producer='another.ps1'),self.prepare()),'lineage')

    def test_wrong_full_scope_blocks_replacement(self):
        self.unchanged_after(lambda:(self.alter_receipt(Scope='Top'),self.prepare()),'full scope')

    def test_wrong_receipt_owner_blocks_replacement(self):
        self.unchanged_after(lambda:(self.alter_receipt(Owner='another-system'),self.prepare()),'owner')

    def test_foreign_source_organization_is_not_restamped(self):
        def reject():
            self.inputs['users'][0]['OrganizationKey']='foreign';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Source reporting identity')

    def test_parent_failure_blocks_replacement(self):
        self.unchanged_after(lambda:(self.alter_receipt(Status='Failed',Errors=1),self.prepare()),'Incomplete producer')

    def test_file_run_differing_from_parent_is_rejected(self):
        def reject():
            self.proof['Files'][0]['RunId']='different-run';self.write_proof();self.prepare()
        self.unchanged_after(reject,'lineage')

    def test_receipt_changed_after_swap_rolls_back(self):
        def fault(phase,source,stage):
            if phase=='after-swap':self.alter_receipt(Status='Running')
        self.unchanged_after(lambda:self.prepare(fault=fault),'Source proof changed')

    def test_actual_powershell_producer_receipts_are_accepted(self):
        helper=pipeline.REGISTRY.with_name('SmartM365-CmdbReceipt.ps1')
        quote=lambda value:"'"+str(value).replace("'","''")+"'"
        command=f"""$ErrorActionPreference='Stop'; . {quote(helper)}
function Test-SmartM365MaxItemsMode {{ return $false }}
function Get-SmartM365ScriptVersionFromFile {{ param($Path) return 'synthetic-test' }}
function WriteLog {{ param($Message,$Level) if($Level -eq 'WARNING'){{throw $Message}} }}
$global:SmartM365TenantKey='synthetic';$global:SmartM365OrganizationKey='test'
$global:SmartM365EnvironmentKey='test';$global:SmartM365TenantId='synthetic-tenant'
$directory={quote(self.source)}
$registry=Get-Content {quote(pipeline.REGISTRY)} -Raw | ConvertFrom-Json
foreach($producer in $registry.Producers){{
    $global:csvGeneratedPaths=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    Start-SmartM365CmdbSourceReceipt -ScriptPath $producer.Script -SourceRootPath $directory
    foreach($file in $producer.Files){{[void]$global:csvGeneratedPaths.Add((Join-Path $directory $file))}}
    if($producer.Files -contains 'Intune_DeviceHardware_All.csv'){{
        $hardwarePath=Join-Path $directory 'Intune_DeviceHardware_All.csv'
        $hardware=@(Import-Csv -LiteralPath $hardwarePath)
        foreach($row in $hardware){{$row.CollectedAtUtc=[datetimeoffset]::UtcNow.ToString('o')}}
        $hardware | Export-Csv -LiteralPath $hardwarePath -NoTypeInformation -Encoding utf8
    }}
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $producer.Scope
    Complete-SmartM365CmdbSourceReceipt -Status Success | Out-Null
}}"""
        result=subprocess.run(['pwsh','-NoProfile','-ExecutionPolicy','Bypass','-Command',command],capture_output=True,text=True,timeout=90)
        self.assertEqual(result.returncode,0,result.stdout+result.stderr)
        result=pipeline.prepare(self.source,self.output,'synthetic',IDENTITY,now=dt.datetime.now(dt.timezone.utc))
        self.assertEqual(result['GeneratedTables'],46)
        manifest=pipeline.load_json(self.output/pipeline.MANIFEST)
        self.assertEqual(len(manifest['SourceEvidence']['ProducerReceipts']),17)

    def unchanged_after(self, action, pattern):
        self.prepare()
        before={p.name:pipeline.sha(p) for p in self.output.iterdir()}
        with self.assertRaisesRegex((ValueError,FileNotFoundError),pattern):
            action()
        self.assertEqual(before,{p.name:pipeline.sha(p) for p in self.output.iterdir()})
        self.assertFalse(list(self.root.glob('.cmdb-stage-*')))
        self.assertFalse(list(self.root.glob('.cmdb-rollback-*')))

    def test_all_46_tables_and_one_current_manifest(self):
        result=self.prepare()
        self.assertEqual(result['GeneratedTables'],46)
        self.assertEqual(len(list(self.output.iterdir())),47)
        pipeline.validate_current(self.output,self.contract,'synthetic')
        self.assertFalse(list(self.output.rglob('*.json')))

    def test_validate_only_writes_nothing(self):
        before={p.name:pipeline.sha(p) for p in self.source.iterdir()}
        self.assertEqual(self.prepare(validate_only=True)['GeneratedTables'],0)
        self.assertFalse(self.output.exists())
        self.assertEqual(before,{p.name:pipeline.sha(p) for p in self.source.iterdir()})
        self.assertFalse((self.root/'.cmdb-preparation.lock').exists())

    def test_distinct_application_footprint_across_versions(self):
        self.prepare()
        top=self.table('TopApplication')
        self.assertEqual(len(top),1);self.assertEqual(top[0]['ReportedDeviceCount'],'1')
        self.assertEqual(top[0]['VersionCount'],'2')
        self.assertEqual(len(self.table('FactDeviceApplication')),2)

    def test_same_name_devices_and_mobile_are_retained(self):
        self.prepare()
        devices=self.table('DimDevice')
        self.assertEqual(len(devices),2)
        self.assertEqual({r['Ownership'] for r in devices},{'Corporate','Personal'})
        self.assertEqual(self.table('FactADIntuneCoverage')[0]['CoverageState'],'Native SID unavailable')
        self.assertEqual(self.table('FactHybridIdentityCoverage')[1]['OnPremisesOnlyCount'],'')

    def test_missing_capacity_is_blank_and_disabled_plan_is_false(self):
        self.prepare()
        self.assertEqual(self.table('DimLicenseSku')[0]['EnabledUnits'],'')
        self.assertEqual(self.table('FactUserServicePlan')[0]['IsEnabled'],'false')

    def test_remote_online_precedence_and_mailbox_types(self):
        self.prepare()
        hosting=self.table('FactMailboxHosting')
        self.assertEqual(len(hosting),1); self.assertEqual(hosting[0]['HostingLocation'],'Exchange Online')
        self.assertEqual(hosting[0]['MailboxTypeGroup'],'User mailboxes')

    def test_native_exo_header_and_guid_retain_addressless_mailboxes(self):
        self.inputs['local_mailboxes']=[];self.inputs['remote_mailboxes']=[];self.write_inputs()
        self.write_native_rows('mailboxes',[
            dict(ExternalDirectoryObjectId='',MailboxGuid='exo-'+str(i),PrimarySmtpAddress='',
                 RecipientTypeDetails='DiscoveryMailbox',CollectedAtUtc=NOW.isoformat()) for i in range(2)])
        self.prepare()
        rows=self.table('FactMailbox')
        self.assertEqual(len(rows),2)
        self.assertEqual(len({r['TenantMailboxKey'] for r in rows}),2)
        self.assertTrue(all(r['SourceSystem']=='mailboxes' for r in rows))
        definition=next(s for s in self.contract['sources'] if s['name']=='mailboxes')
        self.assertEqual(definition['key'],['TenantKey','MailboxGuid'])

    def test_native_local_header_and_recipient_type_are_preserved(self):
        self.inputs['mailboxes']=[];self.inputs['remote_mailboxes']=[];self.write_inputs()
        self.write_native_rows('local_mailboxes',[
            dict(ObjectGUID='local-1',PrimarySMTPaddress='user@synthetic.invalid',
                 RecipientType='SharedMailbox',CollectedAtUtc=NOW.isoformat()),
            dict(ObjectGUID='local-2',PrimarySMTPaddress='',RecipientType='UserMailbox',CollectedAtUtc=NOW.isoformat())])
        self.prepare()
        rows=self.table('FactMailboxHosting')
        self.assertEqual(len(rows),2)
        self.assertEqual({r['MailboxTypeGroup'] for r in rows},{'Shared mailboxes','User mailboxes'})
        self.assertTrue(all(r['HostingLocation']=='Exchange On-premises' for r in rows))
        definition=next(s for s in self.contract['sources'] if s['name']=='local_mailboxes')
        self.assertEqual(definition['key'],['TenantKey','ObjectGUID'])

    def test_renamed_exo_guid_header_cannot_replace_last_valid_output(self):
        def reject():
            row=dict(self.inputs['mailboxes'][0]);row['ExchangeGuid']=row.pop('MailboxGuid')
            self.write_native_rows('mailboxes',[row]);self.prepare()
        self.unchanged_after(reject,'CSV schema mismatch')

    def test_blank_native_local_guid_cannot_replace_last_valid_output(self):
        def reject():
            self.inputs['local_mailboxes'][0]['ObjectGUID']='';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Blank immutable key')

    def test_actual_activity_normalizer_is_consumed_without_collector_execution(self):
        script=Path(__file__).resolve().parents[1]/'SmartInventory/M365Inventory/Usage/SmartM365-M365UserActivity-Inventory.ps1'
        quote=lambda value:"'"+str(value).replace("'","''")+"'"
        command=f"""$ErrorActionPreference='Stop'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile({quote(script)},[ref]$tokens,[ref]$errors)
if($errors.Count){{throw 'Producer parser error'}}
foreach($name in @('ConvertTo-DateOrNull','Get-SourcePropertyValue','ConvertTo-ReportBool','Get-LatestActivity','ConvertFrom-M365UserActivityReport')){{
 $node=$ast.Find({{param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name}},$true)
 if($null -eq $node){{throw "Missing pure normalizer function: $name"}}
 . ([scriptblock]::Create($node.Extent.Text))
}}
$runId='synthetic-run';$Period='D30'
$raw=[pscustomobject]@{{'User Principal Name'='user@synthetic.invalid';'Report Refresh Date'='2026-01-14';
 'Exchange Last Activity Date'='2026-01-10';'Teams Last Activity Date'='2026-01-13';
 'Is Deleted'='True';'Assigned Products'='Synthetic product'}}
ConvertFrom-M365UserActivityReport -Rows @($raw) | ConvertTo-Json -Depth 4
"""
        completed=subprocess.run(['pwsh','-NoProfile','-Command',command],capture_output=True,text=True,check=True)
        row=json.loads(completed.stdout)
        self.assertNotIn('User Principal Name',row)
        self.assertNotIn('Report Refresh Date',row)
        self.write_native_rows('activity',[row]);self.prepare()
        activity=self.table('FactUserActivity')[0]
        self.assertEqual(activity['UserPrincipalName'],'user@synthetic.invalid')
        self.assertEqual(activity['ReportRefreshDate'],'2026-01-14')
        self.assertEqual(activity['LastActivityDate'],'2026-01-13')
        self.assertEqual(activity['LastActivityWorkload'],'Teams')
        self.assertEqual(activity['ExchangeLastActivityDate'],'2026-01-10')
        self.assertEqual(activity['HasAnyM365Activity'],'true')
        self.assertEqual(activity['IsDeleted'],'true')
        self.assertEqual(activity['AssignedProducts'],'Synthetic product')
        self.assertEqual(activity['MatchStatus'],'Resolved UPN')

    def test_activity_latest_date_includes_all_native_workloads_and_ties(self):
        fields={workload+'LastActivityDate':'' for workload in
                ('Exchange','OneDrive','SharePoint','Teams','SkypeForBusiness','Yammer')}
        fields.update(SkypeForBusinessLastActivityDate='2026-01-12',YammerLastActivityDate='2026-01-12')
        self.write_native_rows('activity',[dict(UserPrincipalName='user@synthetic.invalid',
            ReportRefreshDate='2026-01-14',IsDeleted='false',AssignedProducts='',**fields)])
        self.prepare()
        activity=self.table('FactUserActivity')[0]
        self.assertEqual(activity['LastActivityDate'],'2026-01-12')
        self.assertEqual(activity['LastActivityWorkload'],'SkypeForBusiness;Yammer')
        self.assertEqual(activity['HasAnyM365Activity'],'true')

    def test_missing_activity_workload_column_is_not_reported_as_no_activity(self):
        def reject():
            self.write_native_rows('activity',[dict(UserPrincipalName='user@synthetic.invalid',ReportRefreshDate='2026-01-14')])
            self.prepare()
        self.unchanged_after(reject,'CSV schema mismatch')

    def test_conflicting_analytics_scores_preserve_every_last_output_byte(self):
        def reject():
            self.inputs['analytics']=[dict(ReportName='EADeviceScoresV2',DeviceId='md1',EndpointAnalyticsScore=value)
                                      for value in ('71','88')]
            self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Duplicate immutable key: Intune_EndpointAnalytics_DevicePerformance.csv')

    def test_different_analytics_reports_for_one_device_are_not_duplicate_keys(self):
        self.inputs['analytics']=[dict(ReportName='EADeviceScoresV2',DeviceId='md1',EndpointAnalyticsScore='71'),
                                  dict(ReportName='EADevicePerformanceV2',DeviceId='md1',EndpointAnalyticsScore='')]
        self.write_inputs();self.prepare()
        rows=self.table('FactEndpointAnalyticsDevice')
        self.assertEqual(len(rows),2)
        scores={r['SourceSystem']:r['EndpointAnalyticsScore'] for r in rows}
        self.assertEqual(scores,{'EADeviceScoresV2':'71.0','EADevicePerformanceV2':''})

    def test_failed_producer_preserves_last(self):
        def reject():
            self.proof['Files'][0]['Status']='Failed';self.write_proof();self.prepare()
        self.unchanged_after(reject,'Incomplete producer')

    def test_old_acquisition_not_rescued_by_new_transport_time(self):
        def reject():
            producer=self.proof['Files'][0]['Producer']
            for record in self.proof['Files']:
                if record['Producer']==producer:record['StartedAtUtc']=(NOW-dt.timedelta(days=3)).isoformat()
            self.write_proof();self.prepare()
        self.unchanged_after(reject,'Stale acquisition')

    def test_detailed_hardware_overrides_list_default_and_keeps_source_date(self):
        self.inputs['managed'][0]['TotalStorageSpaceInBytes']='0'
        self.write_inputs();self.prepare()
        row=self.table('DeviceHardware')[0]
        self.assertEqual(row['Storage'],'100.00 GiB')
        self.assertEqual(float(row['PhysicalMemoryGiB']),8)
        self.assertEqual(row['HardwareCollectedDateTime'],NOW.isoformat())
        self.assertIn('Per-device GET',row['CollectionMode'])

    def test_hardware_missing_and_zero_are_not_fabricated_capacity(self):
        row=self.inputs['hardware'][0]
        row.update(totalStorageSpaceInBytes='',totalStorageSpaceInBytesStatus='Missing',physicalMemoryInBytes='0',physicalMemoryInBytesStatus='ZeroReported')
        self.write_inputs();self.prepare()
        row=self.table('DeviceHardware')[0]
        self.assertEqual(row['Storage'],'Not provided');self.assertEqual(row['PhysicalMemoryGiB'],'')
        self.assertEqual(row['MemoryStatus'],'ZeroReported')

    def test_failed_hardware_keeps_last_valid_output(self):
        def reject():
            self.inputs['hardware'][0]['CollectionStatus']='Failed';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Incomplete detailed hardware')

    def test_missing_hardware_identity_blocks_promotion(self):
        def reject():
            self.inputs['hardware'].pop();self.write_inputs();self.prepare()
        self.unchanged_after(reject,'identity scope differ')

    def test_hardware_foreign_correlation_blocks_promotion(self):
        def reject():
            self.inputs['hardware'][0]['azureADDeviceId']='another';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'correlation differs')

    def test_hardware_byte_status_mismatch_blocks_promotion(self):
        def reject():
            self.inputs['hardware'][0]['totalStorageSpaceInBytesStatus']='Missing';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'byte status')

    def test_hardware_fractional_capacity_is_rejected(self):
        def reject():
            self.inputs['hardware'][0]['physicalMemoryInBytes']='1.1';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'byte evidence')

    def test_hardware_free_space_cannot_exceed_total(self):
        def reject():
            self.inputs['hardware'][0]['freeStorageSpaceInBytes']='999999999999';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'free storage exceeds')

    def test_hardware_old_row_not_rescued_by_new_receipt(self):
        def reject():
            self.inputs['hardware'][0]['CollectedAtUtc']=(NOW-dt.timedelta(days=1)).isoformat();self.write_inputs();self.prepare()
        self.unchanged_after(reject,'observation outside')

    def test_country_findings_link_to_user_and_device_without_health_conflation(self):
        self.inputs['users'][0]['Usage location']=''
        self.inputs['managed'][0]['ComplianceState']='noncompliant'
        self.write_inputs();self.prepare()
        findings=self.table('EntityFinding')
        self.assertTrue(any(r['FindingType']=='UserCountryUnknown' and r['TenantUserKey'] for r in findings))
        self.assertTrue(any(r['FindingType']=='DeviceCountryUnknown' and r['TenantDeviceKey'] for r in findings))
        self.assertFalse(any('compliant' in r['FindingType'].casefold() for r in findings))

    def test_derived_country_is_information_not_warning(self):
        self.prepare()
        derived=[r for r in self.table('FactDataQuality') if r['FindingType']=='DeviceCountryDerived']
        self.assertEqual(len(derived),2)
        self.assertTrue(all(r['Severity']=='Information' for r in derived))

    def test_repeated_hardware_serial_retains_both_native_records(self):
        self.inputs['hardware'][1]['serialNumber']=self.inputs['hardware'][0]['serialNumber']
        self.write_inputs();self.prepare()
        self.assertEqual(len(self.table('DeviceHardware')),2)
        findings=[r for r in self.table('EntityFinding') if r['FindingType']=='RepeatedSerialValue']
        self.assertEqual(len(findings),2);self.assertTrue(all(r['TenantDeviceKey'] for r in findings))

    def test_license_errors_keep_each_native_path_and_link_users_and_groups(self):
        self.inputs['license_paths'][0]['AssignmentError']='CountViolation'
        self.inputs['license_paths'].append(dict(self.inputs['license_paths'][0],AssignedByGroupId='g1'))
        self.write_inputs();self.prepare()
        findings=[r for r in self.table('EntityFinding') if r['FindingType']=='ObservedLicenseAssignmentError']
        self.assertEqual(len(findings),2);self.assertTrue(all(r['TenantUserKey'] for r in findings))
        self.assertTrue(any(r['TenantGroupKey'] for r in findings))

    def seed_ad(self):
        sid='S-1-5-21-1-2-3-1001';dn='CN=Test,DC=synthetic,DC=invalid'
        self.inputs['users'][0]['OnPremisesSecurityIdentifier']=sid
        ticks=int((NOW-dt.datetime(1601,1,1,tzinfo=dt.timezone.utc)).total_seconds())*10000000
        self.inputs['ad_users']=[dict(ObjectGUID='ad-u1',ObjectSID=sid,UserPrincipalName='',Enabled='true',DistinguishedName=dn,
            DomainName='synthetic.invalid',manager='CN=Boss,DC=synthetic,DC=invalid',LastLogonTimestamp=str(ticks),WhenCreated='09/01/2020 10:00')]
        self.inputs['ad_objects']=[dict(ObjectGUID='ad-u1',ObjectSID=sid,DistinguishedName=dn,ObjectClass='user',PrimaryGroupID='513'),
                                  dict(ObjectGUID='boss',ObjectSID='S-1-5-21-1-2-3-1002',DistinguishedName='CN=Boss,DC=synthetic,DC=invalid',ObjectClass='user')]
        self.inputs['ad_domains']=[dict(DNSRoot='synthetic.invalid',DomainSID='S-1-5-21-1-2-3',Forest='synthetic.invalid')]
        self.inputs['ad_groups']=[dict(ObjectGUID='ad-g1',objectSid='S-1-5-21-1-2-3-513',MembersJson=json.dumps([dn]))]
        self.inputs['ad_members']=[dict(GroupObjectGUID='ad-g1',GroupSID='S-1-5-21-1-2-3-513',MemberDistinguishedName=dn,
            MemberObjectGUID='ad-u1',MemberSID=sid,MemberObjectClass='user',MembershipKind=kind,ResolutionStatus='Resolved') for kind in ('Direct','PrimaryGroup')]

    def test_ad_no_upn_manager_filetime_and_direct_primary_members_retained(self):
        self.seed_ad();self.write_inputs();self.prepare()
        user=self.table('ADUserSource')[0]
        self.assertEqual(user['UserPrincipalName'],'');self.assertTrue(user['TenantUserKey'])
        self.assertTrue(user['ManagerADObjectKey']);self.assertEqual(user['LastLogonUtcDateTime'],NOW.isoformat())
        self.assertEqual(user['CreationUtcDateTime'],'');self.assertEqual(user['CreationRaw'],'09/01/2020 10:00')
        self.assertEqual(len(self.table('ADDirectoryObject')),2);self.assertEqual(len(self.table('ADDomainSource')),1)
        members=self.table('ADMembership');self.assertEqual(len(members),2)
        self.assertTrue(all(r['TenantUserKey'] and r['TenantADGroupKey'] and r['TenantADObjectKey'] for r in members))

    def test_ad_duplicate_sid_does_not_fabricate_two_cloud_links(self):
        self.seed_ad();self.inputs['ad_users'].append(dict(self.inputs['ad_users'][0],ObjectGUID='duplicate'))
        self.write_inputs();self.prepare()
        self.assertTrue(all(not r['TenantUserKey'] and r['CloudMatchStatus']=='Ambiguous native SID' for r in self.table('ADUserSource')))

    def test_ad_groups_link_group_360_only_through_unique_native_sid(self):
        self.seed_ad();self.inputs['groups'][0]['OnPremisesSecurityIdentifier']='S-1-5-21-1-2-3-513'
        self.write_inputs();self.prepare()
        self.assertTrue(self.table('ADGroupSource')[0]['TenantGroupKey'])
        self.assertTrue(all(r['TenantGroupKey'] for r in self.table('ADMembership')))

    def test_ad_computer_duplicate_sid_is_ambiguous(self):
        self.inputs['ad_computers'][0]['SID']='sid'
        self.inputs['ad_computers'].append(dict(self.inputs['ad_computers'][0],ObjectGUID='ad-pc2'))
        self.inputs['entra_devices'][0]['OnPremisesSecurityIdentifier']='sid'
        self.write_inputs();self.prepare()
        self.assertTrue(all(r['CoverageState']=='Ambiguous Entra match' for r in self.table('FactADIntuneCoverage')))

    def test_ad_unresolved_external_member_is_retained(self):
        self.seed_ad();dn='CN=External,DC=outside,DC=invalid'
        self.inputs['ad_groups'][0]['MembersJson']=json.dumps([self.inputs['ad_members'][0]['MemberDistinguishedName'],dn])
        self.inputs['ad_members'].append(dict(GroupObjectGUID='ad-g1',GroupSID='S-1-5-21-1-2-3-513',MemberDistinguishedName=dn,
            MemberObjectGUID='',MemberSID='',MemberObjectClass='',MembershipKind='Direct',ResolutionStatus='UnresolvedOrExternal'))
        self.write_inputs();self.prepare()
        self.assertTrue(any(r['ResolutionStatus']=='UnresolvedOrExternal' and not r['TenantADObjectKey'] for r in self.table('ADMembership')))

    def test_ad_unresolved_primary_group_keeps_native_member_identity(self):
        self.seed_ad();self.inputs['ad_members'][1].update(GroupObjectGUID='',GroupSID='S-1-5-21-1-2-3-999',ResolutionStatus='UnresolvedPrimaryGroup')
        self.inputs['ad_objects'][0]['PrimaryGroupID']='999'
        self.write_inputs();self.prepare()
        member=[r for r in self.table('ADMembership') if r['MembershipKind']=='PrimaryGroup'][0]
        self.assertTrue(member['TenantADObjectKey']);self.assertEqual(member['TenantADGroupKey'],'')

    def test_ad_resolved_member_wrong_guid_blocks_promotion(self):
        self.seed_ad();self.write_inputs()
        def reject():
            self.inputs['ad_members'][0]['MemberObjectGUID']='missing';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'resolved identity')

    def test_ad_direct_membership_missing_from_evidence_blocks_promotion(self):
        self.seed_ad();self.write_inputs()
        def reject():
            self.inputs['ad_members'].pop(0);self.write_inputs();self.prepare()
        self.unchanged_after(reject,'differs from MembersJson')

    def test_entra_group_360_keeps_user_device_nested_and_other_native_members(self):
        self.inputs['group_scope'][0]['MemberCount']='4'
        self.inputs['group_members']=[dict(GroupId='g1',MemberId=mid,MemberType=kind,MembershipKind='Direct',CollectionStatus='Collected')
            for mid,kind in [('u1','#microsoft.graph.user'),('eo1','#microsoft.graph.device'),('g1','#microsoft.graph.group'),('sp1','#microsoft.graph.servicePrincipal')]]
        self.write_inputs();self.prepare()
        members=self.table('EntraGroupMembership');self.assertEqual(len(members),4)
        self.assertTrue(any(r['TenantUserKey'] for r in members));self.assertTrue(any(r['TenantDeviceKey'] for r in members))
        self.assertTrue(any(r['NestedTenantGroupKey'] for r in members));self.assertTrue(any(r['MemberId']=='sp1' and r['LinkStatus']=='Outside collected entity scope' for r in members))

    def test_unresolved_group_member_finding_links_group_360(self):
        self.inputs['group_scope'][0]['MemberCount']='1'
        self.inputs['group_members']=[dict(GroupId='g1',MemberId='missing',MemberType='#microsoft.graph.user',MembershipKind='Direct',CollectionStatus='Collected')]
        self.write_inputs();self.prepare()
        self.assertTrue(any(r['FindingType']=='UnresolvedGroupMember' and r['TenantGroupKey'] for r in self.table('EntityFinding')))

    def test_online_mailbox_precedence_does_not_leave_obsolete_local_link_warning(self):
        self.inputs['local_mailboxes'][0]['PrimarySMTPaddress']='other@synthetic.invalid'
        self.inputs['mailboxes'][0]['PrimarySmtpAddress']='other@synthetic.invalid'
        self.inputs['remote_mailboxes']=[]
        self.write_inputs();self.prepare()
        self.assertFalse(any(r['FindingType']=='UnlinkedMailbox' for r in self.table('FactDataQuality')))

    def test_technical_mailbox_without_external_id_is_information(self):
        self.inputs['mailboxes'].append(dict(ExternalDirectoryObjectId='',MailboxGuid='technical',PrimarySmtpAddress='technical@synthetic.invalid',RecipientTypeDetails='DiscoveryMailbox',CollectedAtUtc=NOW.isoformat()))
        self.write_inputs();self.prepare()
        row=[r for r in self.table('FactDataQuality') if r['FindingType']=='TechnicalMailboxWithoutUser'][0]
        self.assertEqual(row['Severity'],'Information')

    def test_missing_completion_proof_preserves_last(self):
        def reject():
            (self.source/pipeline.load_json(pipeline.REGISTRY)['Producers'][0]['Receipt']).unlink(); self.prepare()
        self.unchanged_after(reject,'')

    def test_incomplete_proof_preserves_last(self):
        def reject():
            self.proof['Files'].pop();self.write_proof();self.prepare()
        self.unchanged_after(reject,'Missing producer')

    def test_partial_scan_preserves_last(self):
        def reject():
            self.proof['Files'][0]['IsPartialInventory']=True;self.write_proof();self.prepare()
        self.unchanged_after(reject,'Incomplete producer')

    def test_hash_mismatch_preserves_last(self):
        def reject():
            with (self.source/self.contract['sources'][0]['file']).open('a') as stream:stream.write('unexpected\n')
            self.prepare()
        self.unchanged_after(reject,'hash mismatch')

    def test_wrong_row_count_preserves_last(self):
        def reject():
            self.proof['Files'][0]['Rows']+=1;self.write_proof();self.prepare()
        self.unchanged_after(reject,'row count mismatch')

    def test_duplicate_native_identity_preserves_last(self):
        def reject():
            self.inputs['managed'].append(copy.deepcopy(self.inputs['managed'][0]));self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Duplicate immutable')

    def test_foreign_tenant_preserves_last(self):
        self.prepare()
        path=self.source/self.contract['sources'][0]['file']
        path.write_text(path.read_text(encoding='utf-8-sig').replace('synthetic,','foreign,'),encoding='utf-8-sig')
        self.proof['Files'][0]['SHA256']=pipeline.sha(path);self.write_proof()
        with self.assertRaisesRegex(ValueError,'TenantKey'):self.prepare()

    def test_orphan_application_relation_preserves_last(self):
        def reject():
            self.inputs['app_relations'][0]['DeviceId']='unknown';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Orphan application')

    def test_partial_application_modes_preserve_last(self):
        def reject():
            self.inputs['apps'][0]['RelationCollectionScope']='Top';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'All-mode')

    def test_application_count_mismatch_preserves_last(self):
        def reject():
            self.inputs['apps'][0]['DeviceCount']='7';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'relation coverage')

    def test_group_completion_mismatch_preserves_last(self):
        def reject():
            self.inputs['group_scope'][0]['MemberCount']='1';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'membership count')

    def test_bad_score_preserves_last(self):
        def reject():
            self.inputs['analytics'][0]['EndpointAnalyticsScore']='101';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'outside 0-100')

    def test_unknown_readiness_state_preserves_last(self):
        def reject():
            self.inputs['readiness'][0]['UpgradeEligibility']='unsupported';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Unknown readiness')

    def test_unknown_plan_state_preserves_last(self):
        def reject():
            self.inputs['user_plans'][0]['StateCode']='unsupported';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Unknown compact')

    def test_mutation_during_swap_rolls_back(self):
        def fault(phase,source,stage):
            if phase=='after-swap':
                with (source/self.contract['sources'][0]['file']).open('a') as stream:stream.write('mutation')
        self.unchanged_after(lambda:self.prepare(fault=fault),'Source changed')

    def test_failed_post_swap_validation_rolls_back(self):
        def fault(phase,source,stage):
            if phase=='after-swap':(stage/'DimUser.csv').write_text('bad')
        self.unchanged_after(lambda:self.prepare(fault=fault),'Output hash')

    def test_successful_replacement_retains_no_history(self):
        self.prepare();self.prepare()
        self.assertFalse(list(self.root.glob('.cmdb-stage-*')))
        self.assertFalse(list(self.root.glob('.cmdb-rollback-*')))
        self.assertFalse((self.root/'DATA-POWERBI').exists())

    def test_concurrent_publisher_is_refused(self):
        self.prepare()
        with pipeline.PublicationLock(self.root/'.cmdb-preparation.lock'):
            with self.assertRaisesRegex(ValueError,'publication lock'):self.prepare()

    def test_unowned_directory_is_not_deleted(self):
        self.output.mkdir();(self.output/'keep.txt').write_text('not ours')
        with self.assertRaises(FileNotFoundError):self.prepare()
        self.assertEqual((self.output/'keep.txt').read_text(),'not ours')

    def test_wrong_output_root_is_refused(self):
        with self.assertRaisesRegex(ValueError,'sibling'):pipeline.paths(self.source,self.root/'DATA-POWERBI')

    def test_all_successful_empty_sources_have_valid_headers(self):
        self.inputs={name:[] for name in self.inputs};self.write_inputs();self.prepare()
        self.assertEqual(self.table('HardwareCoverage')[0]['RecordCount'],'0')
        self.assertEqual(pipeline.header(self.output/'DimUser.csv'),next(t for t in self.contract['tables'] if t['name']=='DimUser')['columns'])

    def test_autopilot_shared_serial_keeps_both_native_identities(self):
        self.inputs['autopilot']=[{'Autopilot ID':'ap1','Serial number':'same'},
                                  {'Autopilot ID':'ap2','Serial number':'same'}]
        self.write_inputs();self.prepare()
        self.assertEqual({r['AutopilotDeviceId'] for r in self.table('FactAutopilotDevice')},{'ap1','ap2'})

    def test_missing_required_producer_headers_preserve_last_output(self):
        for source,field in [('alerts','SourceReport'),('alerts','AlertName'),('teams','MemberCount'),
                             ('group_members','MemberType'),('group_members','CollectionStatus'),
                             ('ad_members','MemberObjectGUID'),('ad_members','GroupObjectGUID')]:
            with self.subTest(source=source,field=field):
                self.write_inputs()
                def reject():
                    definition=next(s for s in self.contract['sources'] if s['name']==source)
                    path=self.source/definition['file']
                    columns=[c for c in pipeline.header(path) if c!=field]
                    data=list(pipeline.rows(path))
                    with path.open('w',encoding='utf-8-sig',newline='') as stream:
                        writer=csv.DictWriter(stream,fieldnames=columns,extrasaction='ignore')
                        writer.writeheader();writer.writerows(data)
                    record=next(r for r in self.proof['Files'] if r['File']==path.name)
                    record['SHA256']=pipeline.sha(path);self.write_proof();self.prepare()
                self.unchanged_after(reject,'CSV schema mismatch')

    def test_duplicate_mailbox_smtp_in_one_source_preserves_last(self):
        def reject():
            duplicate=dict(self.inputs['mailboxes'][0], MailboxGuid='ex2')
            self.inputs['mailboxes'].append(duplicate);self.write_inputs();self.prepare()
        self.unchanged_after(reject,'share SMTP')

    def test_foreign_output_reporting_identity_is_rejected(self):
        self.prepare()
        path=self.output/'DimUser.csv'
        path.write_text(path.read_text(encoding='utf-8-sig').replace(',test,test,',',wrong,test,'),encoding='utf-8-sig')
        manifest=pipeline.load_json(self.output/pipeline.MANIFEST)
        manifest['OutputFiles'][path.name]['SHA256']=pipeline.sha(path)
        (self.output/pipeline.MANIFEST).write_text(json.dumps(manifest))
        with self.assertRaisesRegex(ValueError,'Reporting identity'):pipeline.validate_current(self.output,self.contract,'synthetic')

    def test_orphan_output_link_is_rejected_even_with_updated_hash(self):
        self.prepare()
        path=self.output/'FactUserLicense.csv'
        path.write_text(path.read_text(encoding='utf-8-sig').replace(self.table('DimUser')[0]['TenantUserKey'],'synthetic|missing'),encoding='utf-8-sig')
        manifest=pipeline.load_json(self.output/pipeline.MANIFEST)
        manifest['OutputFiles'][path.name]['SHA256']=pipeline.sha(path)
        (self.output/pipeline.MANIFEST).write_text(json.dumps(manifest))
        with self.assertRaisesRegex(ValueError,'Orphan output'):pipeline.validate_current(self.output,self.contract,'synthetic')

    def test_known_native_sid_chain_proves_intune_management(self):
        self.inputs['ad_computers'][0]['ObjectSID']='S-1-5-21-1'
        self.inputs['entra_devices'][0]['OnPremisesSecurityIdentifier']='S-1-5-21-1'
        self.write_inputs();self.prepare()
        self.assertEqual(self.table('FactADIntuneCoverage')[0]['CoverageState'],'Managed in Intune')
        self.assertEqual(self.table('FactADIntuneCoverage')[0]['IntuneManagedDeviceId'],'md1')

    def test_disabled_and_recent_user_activity_classification(self):
        self.inputs['users'][0]['LastSuccessfulSignInDateTime']=(NOW-dt.timedelta(days=15)).isoformat()
        self.write_inputs();self.prepare()
        self.assertEqual(self.table('DimUser')[0]['ActivityState'],'Active 30d')
        self.inputs['users'][0]['AccountEnabled']='false';self.write_inputs();self.prepare()
        self.assertEqual(self.table('DimUser')[0]['ActivityState'],'Disabled')

    def test_collaboration_membership_and_90_day_activity(self):
        self.inputs['teams']=[{'TeamId':'t1','TeamDisplayName':'Test Team','MemberCollectionStatus':'Collected',
                              'MemberCount':'1','OwnerCount':'1','GuestCount':'0','IsArchived':'false',
                              'LastActivityDateUtc':(NOW-dt.timedelta(days=100)).isoformat(),'CollectedAtUtc':NOW.isoformat()}]
        self.inputs['team_members']=[{'TeamId':'t1','UserId':'u1','Role':'Owner'}]
        self.inputs['sites']=[{'SiteId':'host,s1,w1','SiteIdentitySource':'Graph','SiteUrl':'https://synthetic.invalid/site',
                              'ReportRefreshDate':NOW.date().isoformat(),'CollectedAtUtc':NOW.isoformat(),
                              'LastActivityUtc':(NOW-dt.timedelta(days=10)).isoformat(),'StorageUsedMB':'2'}]
        self.write_inputs();self.prepare()
        self.assertEqual(self.table('DimTeam')[0]['ActivityState'],'Inactive (>90d)')
        self.assertEqual(self.table('DimTeam')[0]['MembershipCoverageStatus'],'Complete')
        self.assertEqual(self.table('DimSharePointSite')[0]['ActivityState'],'Active (90d)')
        self.assertEqual(float(self.table('DimSharePointSite')[0]['StorageUsedBytes']),2097152)

    def test_policy_families_native_payloads_and_parent_links(self):
        families=['SettingsCatalog','DeviceConfiguration','DeviceCompliance','WindowsFeatureUpdate','WindowsQualityUpdate']
        self.inputs['policies']=[{'PolicyFamily':f,'PolicyId':'p'+str(i),'NativeEvidenceJson':'{}',
                                 'AssignmentCollectionStatus':'Collected'} for i,f in enumerate(families)]
        self.inputs['policy_assignments']=[{'PolicyFamily':'SettingsCatalog','PolicyId':'p0','AssignmentId':'target',
                                           'NativeTargetJson':'{"@odata.type":"#microsoft.graph.allDevicesAssignmentTarget"}'}]
        self.write_inputs();self.prepare()
        self.assertEqual(len(self.table('DimIntuneConfigurationPolicy')),3)
        self.assertEqual(len(self.table('DimWindowsUpdatePolicy')),2)
        self.inputs['policy_assignments'][0]['PolicyId']='missing';self.write_inputs()
        with self.assertRaisesRegex(ValueError,'Policy assignment parent'):self.prepare()

    def test_naive_acquisition_time_is_rejected(self):
        self.proof['Files'][0]['CompletedAtUtc']='2026-01-15T12:00:00';self.write_proof()
        with self.assertRaisesRegex(ValueError,'explicit timezone'):self.prepare()

    def test_cleanup_error_does_not_undo_validated_publication(self):
        self.prepare()
        real_remove=pipeline.shutil.rmtree
        def fail_previous(path, *args, **kwargs):
            if Path(path).name.startswith('.cmdb-rollback-'):
                raise PermissionError('synthetic cleanup failure')
            return real_remove(path, *args, **kwargs)
        with mock.patch.object(pipeline.shutil,'rmtree',side_effect=fail_previous):
            self.assertEqual(self.prepare()['Status'],'PreparedWithCleanupWarning')
        pipeline.validate_current(self.output,self.contract,'synthetic')

    def test_unresolved_team_user_is_retained_not_silently_dropped(self):
        self.inputs['teams']=[{'TeamId':'t1','TeamDisplayName':'Team','MemberCollectionStatus':'Collected',
                              'MemberCount':'1','CollectedAtUtc':NOW.isoformat()}]
        self.inputs['team_members']=[{'TeamId':'t1','UserId':'external-user','Role':'Member'}]
        self.write_inputs();self.prepare()
        self.assertEqual(self.table('FactTeamMember')[0]['UserId'],'external-user')
        self.assertEqual(self.table('DimTeam')[0]['UnresolvedMemberCount'],'1')


if __name__ == '__main__':
    unittest.main(verbosity=2)
