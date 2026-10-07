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


class CsvReaderTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.gettempdir()) / ('SmartInventory-Cmdb-Reader-' + uuid.uuid4().hex)
        self.root.mkdir()
        self.path = self.root / 'synthetic.csv'

    def tearDown(self):
        if self.root.parent.resolve() == Path(tempfile.gettempdir()).resolve() and self.root.name.startswith('SmartInventory-Cmdb-Reader-'):
            shutil.rmtree(self.root)

    def test_large_multiline_quoted_membership_is_lossless(self):
        value = 'member,"quoted"\r\n' * 200000
        previous = csv.field_size_limit()
        with self.path.open('w', encoding='utf-8-sig', newline='') as stream:
            writer = csv.writer(stream)
            writer.writerow(['TenantKey', 'MembersJson'])
            writer.writerow(['synthetic', value])
        before = pipeline.sha(self.path)
        self.assertEqual(list(pipeline.rows(self.path)), [{'TenantKey':'synthetic', 'MembersJson':value}])
        self.assertEqual(pipeline.sha(self.path), before)
        self.assertEqual(csv.field_size_limit(), previous)

    def test_reader_restores_limit_after_malformed_row(self):
        previous = csv.field_size_limit()
        self.path.write_text('TenantKey,MembersJson\nsynthetic\n', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'Malformed logical CSV row'):
            list(pipeline.rows(self.path))
        self.assertEqual(csv.field_size_limit(), previous)

    def test_reader_restores_limit_when_closed_early(self):
        previous = csv.field_size_limit()
        self.path.write_text('TenantKey\nsynthetic\nsynthetic\n', encoding='utf-8')
        reader = pipeline.rows(self.path)
        next(reader)
        reader.close()
        self.assertEqual(csv.field_size_limit(), previous)

    def test_reader_retains_a_finite_field_limit(self):
        previous = csv.field_size_limit()
        try:
            csv.field_size_limit(512)
            self.path.write_text('TenantKey,MembersJson\nsynthetic,' + 'x' * 1025 + '\n', encoding='utf-8')
            with mock.patch.object(pipeline, 'CSV_FIELD_LIMIT', 1024):
                with self.assertRaises(csv.Error):
                    list(pipeline.rows(self.path))
            self.assertEqual(csv.field_size_limit(), 512)
        finally:
            csv.field_size_limit(previous)


class LicenseAssignmentParentTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix='SmartInventory-Cmdb-Parents-')
        self.source = Path(self.temporary.name)
        self.contract = pipeline.load_json(pipeline.CONTRACT)
        self.definitions = {item['name']: item for item in self.contract['sources']}
        self.write('users', ['Object Id'], [{'Object Id': 'u1'}])
        self.write('skus', ['Id'], [{'Id': 's1'}])
        self.write_paths([{'UserId': 'u1', 'SkuId': 's1'}])

    def tearDown(self):
        self.temporary.cleanup()

    def write(self, name, columns, data):
        with (self.source / self.definitions[name]['file']).open(
                'w', encoding='utf-8-sig', newline='') as stream:
            writer = csv.DictWriter(stream, fieldnames=columns)
            writer.writeheader()
            writer.writerows(data)

    def write_paths(self, data):
        self.write('license_paths', ['UserId', 'SkuId', 'AssignedByGroupId', 'AssignmentState'], data)

    def validate(self):
        return pipeline.validate_license_assignment_parents(self.source, self.contract)

    def test_normalized_native_ids_match_without_changing_files(self):
        self.write_paths([{'UserId': ' U1 ', 'SkuId': ' S1 ', 'AssignedByGroupId': '',
                           'AssignmentState': 'Active'},
                          {'UserId': 'u1', 'SkuId': 's1', 'AssignedByGroupId': 'unresolved-group',
                           'AssignmentState': 'Error'}])
        before = {p.name: pipeline.sha(p) for p in self.source.iterdir()}
        self.validate()
        self.assertEqual(before, {p.name: pipeline.sha(p) for p in self.source.iterdir()})

    def test_valid_empty_assignment_set_is_allowed(self):
        self.write_paths([])
        self.validate()

    def test_missing_users_count_rows_and_distinct_native_ids(self):
        self.write_paths([{'UserId': 'missing-user', 'SkuId': 's1', 'AssignedByGroupId': ''},
                          {'UserId': ' MISSING-USER ', 'SkuId': 's1', 'AssignedByGroupId': 'g1'}])
        with self.assertRaises(ValueError) as failure:
            self.validate()
        message = str(failure.exception)
        self.assertIn('MissingUserRows=2; MissingUserIds=1', message)
        self.assertIn('MissingSkuRows=0; MissingSkuIds=0', message)
        self.assertIn('M365_Users_Active.csv (UserId -> Object Id)', message)
        self.assertNotIn('missing-user', message)

    def test_missing_skus_are_distinguished_from_missing_users(self):
        self.write_paths([{'UserId': 'u1', 'SkuId': 'missing-sku'}])
        with self.assertRaises(ValueError) as failure:
            self.validate()
        self.assertIn('MissingUserRows=0; MissingUserIds=0', str(failure.exception))
        self.assertIn('MissingSkuRows=1; MissingSkuIds=1', str(failure.exception))
        self.assertIn('M365_Licenses_Tenant.csv (SkuId -> Id)', str(failure.exception))

    def test_complete_scan_reports_both_missing_parent_types(self):
        self.write_paths([{'UserId': 'missing-user', 'SkuId': 's1'},
                          {'UserId': 'u1', 'SkuId': 'missing-sku'},
                          {'UserId': 'other-user', 'SkuId': 'missing-sku'}])
        with self.assertRaises(ValueError) as failure:
            self.validate()
        self.assertIn('MissingUserRows=2; MissingUserIds=2', str(failure.exception))
        self.assertIn('MissingSkuRows=2; MissingSkuIds=1', str(failure.exception))

    def test_error_disabled_and_unknown_paths_are_not_filtered(self):
        for state in ['Active', 'ActiveWithError', 'Error', 'Disabled', '', 'Unknown']:
            with self.subTest(state=state):
                self.write_paths([{'UserId': 'missing-user', 'SkuId': 's1', 'AssignmentState': state}])
                with self.assertRaisesRegex(ValueError, 'MissingUserRows=1'):
                    self.validate()

    def test_blank_child_ids_cannot_match_blank_parent_ids(self):
        self.write('users', ['Object Id'], [{'Object Id': ''}])
        self.write('skus', ['Id'], [{'Id': ''}])
        self.write_paths([{'UserId': '', 'SkuId': ''}])
        with self.assertRaises(ValueError) as failure:
            self.validate()
        self.assertIn('MissingUserRows=1', str(failure.exception))
        self.assertIn('MissingSkuRows=1', str(failure.exception))


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
            'license_overview':[{'Id':'u1-s1','UserId':'u1','SkuId':'s1','Source':'Direct','GroupsAssigningSku':''}],
            'plans':[{'SkuId':'s1','PlanId':'p1','PlanName':'Test plan','TenantProvisioningStatus':'Success'}],
            'user_plans':[{'UserId':'u1','SkuId':'s1','PlanId':'p1','StateCode':'D'}],
            'groups':[{'GroupId':'g1','DisplayName':'Group'}],
            'group_scope':[{'GroupId':'g1','MemberCount':'0','MemberCollectionStatus':'Collected',
                            'DisplayName':'Group','MailEnabled':'false','SecurityEnabled':'true','GroupTypes':'',
                            'OnPremisesSecurityIdentifier':'','RunId':'synthetic-run',
                            'GroupCollectedAtUtc':(NOW-dt.timedelta(minutes=2)).isoformat(),'CollectedAtUtc':NOW.isoformat()}],
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
            if source['name'] == 'group_members':
                for row in data:
                    row.setdefault('RunId', 'synthetic-run')
                    row.setdefault('CollectedAtUtc', NOW.isoformat())
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

    def source_run(self, status='Collecting', unqualified=False, protocol=True):
        producer=pipeline.load_json(pipeline.REGISTRY)['Producers'][0]
        path=self.source/producer['Receipt']
        proof=pipeline.load_json(path)
        if protocol:
            proof['PublicationProtocol']=1
            path.write_text(json.dumps(proof),encoding='utf-8')
        run={'Owner':'SmartInventory-SourceRun','ContractVersion':'1.0',**IDENTITY,
             'Producer':producer['Script'],'RunId':'new-attempt','Status':status,
             'PublicationStarted':unqualified,'UnqualifiedPublication':unqualified}
        run_path=path.with_name(path.name.replace('.current.json.txt','.run.json.txt'))
        run_path.write_text(json.dumps(run),encoding='utf-8')
        return path,run_path,proof,run

    def test_collecting_and_prepublication_failure_preserve_previous_evidence(self):
        for status in ['Collecting','Failed']:
            with self.subTest(status=status):
                self.write_proof()
                path,run_path,proof,run=self.source_run(status)
                digest=pipeline.sha(path)
                records,receipts,_=pipeline.producer_records(self.source,IDENTITY)
                self.assertTrue(records)
                self.assertEqual(receipts[0]['RunId'],'synthetic-run')
                self.assertEqual(pipeline.sha(path),digest)
                self.assertEqual(receipts[0]['StartedAtUtc'],proof['StartedAtUtc'])

    def test_publishing_failed_replacement_and_retry_are_rejected(self):
        for status in ['Publishing','Failed','Collecting']:
            with self.subTest(status=status):
                self.source_run(status,unqualified=True)
                with self.assertRaisesRegex(ValueError,'publication is in progress or remains unqualified'):
                    pipeline.producer_records(self.source,IDENTITY)

    def test_new_attempt_fences_legacy_completed_proof_too(self):
        self.source_run('Publishing',unqualified=True,protocol=False)
        with self.assertRaisesRegex(ValueError,'publication is in progress'):
            pipeline.producer_records(self.source,IDENTITY)

    def test_run_state_cannot_replace_or_fabricate_completion_proof(self):
        path,_,_,_=self.source_run('Completed')
        path.unlink()
        with self.assertRaisesRegex(ValueError,'Missing producer completion proof'):
            pipeline.producer_records(self.source,IDENTITY)

    def test_missing_corrupt_foreign_and_untyped_run_states_are_rejected(self):
        for change in ['missing','corrupt','foreign','untyped','unknown']:
            with self.subTest(change=change):
                path,run_path,proof,run=self.source_run()
                if change=='missing': run_path.unlink()
                elif change=='corrupt': run_path.write_text('{',encoding='utf-8')
                else:
                    run.update({'TenantKey':'foreign'} if change=='foreign' else
                               {'UnqualifiedPublication':'false'} if change=='untyped' else {'Status':'Unknown'})
                    run_path.write_text(json.dumps(run),encoding='utf-8')
                with self.assertRaises((ValueError,json.JSONDecodeError)):
                    pipeline.producer_records(self.source,IDENTITY)

    def test_publication_started_after_validation_blocks_final_recheck(self):
        path,run_path,proof,run=self.source_run()
        records,receipts,registry_hash=pipeline.producer_records(self.source,IDENTITY)
        evidence={'ProducerReceipts':receipts,'RegistrySHA256':registry_hash,'Files':list(records.values())}
        pipeline.recheck_sources(self.source,evidence,self.contract)
        run.update(Status='Publishing',PublicationStarted=True,UnqualifiedPublication=True)
        run_path.write_text(json.dumps(run),encoding='utf-8')
        with self.assertRaisesRegex(ValueError,'publication is in progress'):
            pipeline.recheck_sources(self.source,evidence,self.contract)

    def test_collecting_does_not_reset_expired_acquisition_dates(self):
        self.source_run()
        records,_,_=pipeline.producer_records(self.source,IDENTITY)
        record=next(iter(records.values()))
        self.assertEqual(record['StartedAtUtc'],(NOW-dt.timedelta(minutes=5)).isoformat())
        expired=NOW+dt.timedelta(hours=self.contract['maxAgeHours']+1)
        with self.assertRaisesRegex(ValueError,'[Ff]resh|[Aa]ge|[Ee]xpir|[Ss]tale'):
            pipeline.validate_sources(self.source,self.contract,'synthetic',now=expired,identity=IDENTITY)

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

    def test_incomplete_producer_diagnostic_identifies_receipt_and_typed_flags(self):
        producer=pipeline.load_json(pipeline.REGISTRY)['Producers'][0]
        cases=[({'Status':'Running'}, "Status='Running'"),
               ({'Status':'Failed','IsPartialInventory':True,'Errors':1}, 'Errors=1 (int)'),
               ({'IsPartialInventory':True}, 'IsPartialInventory=True'),
               ({'IsPartialInventory':'false'}, "IsPartialInventory='false'"),
               ({'Errors':'0'}, "Errors='0' (str)"),
               ({'Errors':False}, 'Errors=False (bool)'),
               ({'Errors':None}, 'Errors=None (NoneType)'),
               ({'Errors':-1}, 'Errors=-1 (int)'),
               ({'Errors':1.0}, 'Errors=1.0 (float)')]
        for updates, expected in cases:
            with self.subTest(updates=updates):
                self.write_proof()
                self.alter_receipt(**updates)
                with self.assertRaises(ValueError) as caught:
                    pipeline.producer_records(self.source,IDENTITY)
                message=str(caught.exception)
                self.assertIn('Incomplete producer completion proof',message)
                self.assertIn('Producer='+repr(producer['Script']),message)
                self.assertIn('Receipt='+repr(producer['Receipt']),message)
                self.assertIn(expected,message)
                self.assertNotIn(IDENTITY['TenantId'],message)

    def test_missing_producer_diagnostic_identifies_receipt(self):
        producer=pipeline.load_json(pipeline.REGISTRY)['Producers'][5]
        (self.source/producer['Receipt']).unlink()
        with self.assertRaises(ValueError) as caught:
            pipeline.producer_records(self.source,IDENTITY)
        message=str(caught.exception)
        self.assertIn('Missing producer completion proof',message)
        self.assertIn('Producer='+repr(producer['Script']),message)
        self.assertIn('Receipt='+repr(producer['Receipt']),message)

    def test_receipt_diagnostic_alone_does_not_create_outputs_or_modify_sources(self):
        self.alter_receipt(Status='Failed',IsPartialInventory=True,Errors=1)
        before={p.name:pipeline.sha(p) for p in self.source.iterdir() if p.is_file()}
        with self.assertRaisesRegex(ValueError,'Incomplete producer completion proof'):
            pipeline.producer_records(self.source,IDENTITY)
        self.assertEqual(before,{p.name:pipeline.sha(p) for p in self.source.iterdir() if p.is_file()})
        self.assertFalse(self.output.exists())

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
    if($producer.Files -contains 'M365_EntraGroupMembershipScope.csv'){{
        $scopePath=Join-Path $directory 'M365_EntraGroupMembershipScope.csv'
        $scopeRows=@(Import-Csv -LiteralPath $scopePath)
        $catalogDate=[datetimeoffset]::UtcNow.ToString('o')
        foreach($row in $scopeRows){{$row.GroupCollectedAtUtc=$catalogDate;$row.CollectedAtUtc=$catalogDate}}
        $scopeRows | Export-Csv -LiteralPath $scopePath -NoTypeInformation -Encoding utf8
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

    def shared_receipts(self):
        for producer in pipeline.load_json(pipeline.REGISTRY)['Producers']:
            path = self.source / producer['Receipt']
            proof = pipeline.load_json(path)
            proof.update(Owner='SmartInventory-SourceReceipt', ContractVersion='1.2',
                         ScopeQualification='ConsumerScope', RequiredFiles=producer['Files'],
                         FullInventoryQualified=False, ConsumerScopeQualified=True)
            path.write_text(json.dumps(proof), encoding='utf-8')

    def test_shared_receipts_accept_additional_current_exports(self):
        self.shared_receipts()
        producer = pipeline.load_json(pipeline.REGISTRY)['Producers'][0]
        path = self.source / producer['Receipt']
        proof = pipeline.load_json(path)
        extra = dict(proof['Files'][0], File='AdditionalPublished.csv', Required=False)
        proof['Files'].append(extra)
        path.write_text(json.dumps(proof), encoding='utf-8')
        records, receipts, _ = pipeline.producer_records(self.source, IDENTITY)
        self.assertEqual(len(records), 34)
        self.assertEqual(len(receipts), 17)
        self.assertNotIn(extra['File'], records)
        self.assertEqual(self.prepare()['GeneratedTables'], 46)

    def test_license_overview_preserves_legitimate_repeated_display_rows(self):
        # Group display names are not immutable group IDs; never deduplicate this view.
        row = dict(Id='u1-s1', UserId='u1', SkuId='s1', Source='Group', GroupsAssigningSku='Same label')
        self.inputs['license_overview'] = [row.copy(), row.copy()]
        self.write_inputs()
        self.assertEqual(self.prepare()['GeneratedTables'], 46)
        health = next(r for r in self.table('SourceHealth') if r['SourceName'] == 'M365_Licenses_Users.csv')
        self.assertEqual(health['SourceRows'], '2')

    def test_license_overview_is_required_and_hash_bound(self):
        path = self.source / 'M365_Licenses_Users.csv'
        with path.open('a', encoding='utf-8') as stream:
            stream.write('changed\n')
        with self.assertRaisesRegex(ValueError, 'hash mismatch'):
            self.prepare()
        self.write_inputs()
        definition = next(p for p in pipeline.load_json(pipeline.REGISTRY)['Producers']
                          if p['Script'] == 'SmartM365-Licences-Inventory.ps1')
        receipt = self.source / definition['Receipt']
        proof = pipeline.load_json(receipt)
        proof['Files'] = [r for r in proof['Files'] if r['File'] != path.name]
        receipt.write_text(json.dumps(proof), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'Missing producer completion proof'):
            self.prepare()

    def test_generic_configured_scope_does_not_qualify_cmdb(self):
        self.shared_receipts()
        producer = pipeline.load_json(pipeline.REGISTRY)['Producers'][0]
        path = self.source / producer['Receipt']
        proof = pipeline.load_json(path)
        proof.update(ScopeQualification='ConfiguredOutputsOnly', IsPartialInventory=None)
        path.write_text(json.dumps(proof), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'consumer-scope'):
            pipeline.producer_records(self.source, IDENTITY)

    def test_shared_receipt_missing_required_export_is_rejected(self):
        self.shared_receipts()
        producer = pipeline.load_json(pipeline.REGISTRY)['Producers'][0]
        path = self.source / producer['Receipt']
        proof = pipeline.load_json(path)
        proof['Files'] = []
        path.write_text(json.dumps(proof), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'Missing producer completion proof'):
            pipeline.producer_records(self.source, IDENTITY)

    def test_shared_receipt_cannot_redefine_required_outputs(self):
        self.shared_receipts()
        producer = pipeline.load_json(pipeline.REGISTRY)['Producers'][0]
        path = self.source / producer['Receipt']
        proof = pipeline.load_json(path)
        proof['RequiredFiles'] = []
        path.write_text(json.dumps(proof), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'consumer-scope'):
            pipeline.producer_records(self.source, IDENTITY)

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

    def test_validate_only_rejects_orphan_license_user_without_output_or_source_writes(self):
        self.inputs['license_paths'][0]['UserId'] = 'missing-user'
        self.write_inputs()
        before = {p.name: pipeline.sha(p) for p in self.source.iterdir()}
        with mock.patch.object(pipeline, 'PublicationLock', side_effect=AssertionError('Lock must not start')):
            with self.assertRaisesRegex(ValueError, 'MissingUserRows=1; MissingUserIds=1'):
                self.prepare(validate_only=True)
        self.assertFalse(self.output.exists())
        self.assertFalse(list(self.root.glob('.cmdb-*')))
        self.assertEqual(before, {p.name: pipeline.sha(p) for p in self.source.iterdir()})

    def test_validate_only_rejects_orphan_license_sku_before_generation(self):
        self.inputs['license_paths'][0]['SkuId'] = 'missing-sku'
        self.write_inputs()
        with self.assertRaisesRegex(ValueError, 'MissingSkuRows=1; MissingSkuIds=1'):
            self.prepare(validate_only=True)
        self.assertFalse(self.output.exists())

    def test_orphan_license_parents_preserve_last_output_before_staging(self):
        self.prepare()
        before = {p.name: pipeline.sha(p) for p in self.output.iterdir()}
        self.inputs['license_paths'][0].update(UserId='missing-user', SkuId='missing-sku')
        self.write_inputs()
        sources = {p.name: pipeline.sha(p) for p in self.source.iterdir()}
        for validate_only in [True, False]:
            with self.subTest(validate_only=validate_only):
                with mock.patch.object(pipeline, 'PublicationLock', side_effect=AssertionError('Lock must not start')):
                    with self.assertRaises(ValueError) as failure:
                        self.prepare(validate_only=validate_only)
                self.assertIn('MissingUserRows=1; MissingUserIds=1', str(failure.exception))
                self.assertIn('MissingSkuRows=1; MissingSkuIds=1', str(failure.exception))
                self.assertEqual(before, {p.name: pipeline.sha(p) for p in self.output.iterdir()})
                self.assertEqual(sources, {p.name: pipeline.sha(p) for p in self.source.iterdir()})
                self.assertFalse(list(self.root.glob('.cmdb-stage-*')))
                self.assertFalse(list(self.root.glob('.cmdb-rollback-*')))

    def test_refreshed_native_user_unblocks_validation_and_preserves_error_path(self):
        self.inputs['license_paths'].append(dict(self.inputs['license_paths'][0],
                                                UserId='new-user', AssignmentState='Error'))
        self.write_inputs()
        with self.assertRaisesRegex(ValueError, 'MissingUserRows=1'):
            self.prepare(validate_only=True)
        new_user = dict(self.inputs['users'][0], **{'Object Id': 'new-user'})
        self.inputs['users'].append(new_user)
        self.write_inputs()
        self.assertEqual(self.prepare(validate_only=True)['Status'], 'ValidatedSources')
        self.assertEqual(self.prepare()['GeneratedTables'], 46)
        error_paths = [row for row in self.table('LicenseAssignmentPath')
                       if row['SourceUserId'] == 'new-user']
        self.assertEqual(len(error_paths), 1)
        self.assertEqual(error_paths[0]['AssignmentState'], 'Error')

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

    def test_native_large_ad_group_fields_pass_validation_and_generation(self):
        dns = [f'CN=member-{i},' + 'OU=Synthetic,' * 25 + 'DC=synthetic,DC=invalid' for i in range(10000)]
        members = json.dumps(dns)
        self.assertGreater(len(members),3300000)
        group_sid = 'S-1-5-21-1-2-3-2001'
        self.inputs['ad_groups']=[dict(ObjectGUID='ad-group-1',ObjectSID=group_sid,Name='Synthetic group',MembersJson=members)]
        self.inputs['ad_members']=[dict(GroupObjectGUID='ad-group-1',GroupSID=group_sid,MemberDistinguishedName=dn,
            MemberObjectGUID='',MemberSID='',MemberObjectClass='user',MembershipKind='Direct',
            ResolutionStatus='UnresolvedOrExternal') for dn in dns]
        self.write_inputs()
        path=self.source/'AD_Groups_AllDomains.csv'
        before=pipeline.sha(path)
        self.prepare()
        self.assertEqual(list(pipeline.rows(path))[0]['MembersJson'],members)
        self.assertEqual(pipeline.sha(path),before)
        self.assertEqual(self.table('ADGroupSource')[0]['ObjectGUID'],'ad-group-1')
        self.assertEqual(len(self.table('ADMembership')),len(dns))

    def test_large_ad_group_fields_do_not_weaken_duplicate_key_rejection(self):
        def reject():
            row=dict(ObjectGUID='ad-group-1',MembersJson=json.dumps(['x' * 200000]))
            self.inputs['ad_groups']=[row,dict(row)]
            self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Duplicate immutable key: AD_Groups_AllDomains.csv')

    def test_different_analytics_reports_for_one_device_are_not_duplicate_keys(self):
        self.inputs['analytics']=[dict(ReportName='EADeviceScoresV2',DeviceId='md1',EndpointAnalyticsScore='71'),
                                  dict(ReportName='EADevicePerformanceV2',DeviceId='md1',EndpointAnalyticsScore='')]
        self.write_inputs();self.prepare()
        rows=self.table('FactEndpointAnalyticsDevice')
        self.assertEqual(len(rows),2)
        scores={r['SourceSystem']:r['EndpointAnalyticsScore'] for r in rows}
        self.assertEqual(scores,{'EADeviceScoresV2':'71.0','EADevicePerformanceV2':''})

    def test_excluded_score_keeps_device_blank_score_and_explicit_qualification(self):
        self.prepare()
        devices_before = self.table('DimDevice')
        self.assertEqual(self.table('FactEndpointAnalyticsDevice')[0]['EndpointAnalyticsScore'], '70.0')
        # The collector removes every duplicate score row, not the inventory device.
        self.inputs['analytics'] = [dict(ReportName='EADevicePerformanceV2', DeviceId='md1',
                                        EndpointAnalyticsScore='', AppReliabilityScore='90')]
        self.write_inputs()
        producer = next(p for p in pipeline.load_json(pipeline.REGISTRY)['Producers']
                        if p['Scope'] == 'CMDB:analytics')
        receipt_path = self.source / producer['Receipt']
        receipt = pipeline.load_json(receipt_path)
        qualification = ('DuplicateScoreRowsExcluded: report=EADeviceScoresV2; rawRows=2; '
                         'publishedRows=0; excludedRows=2; excludedDevices=1; policy=ExcludeAllDuplicateKeys')
        receipt['Qualifications'] = [qualification]
        receipt_path.write_text(json.dumps(receipt), encoding='utf-8')
        self.prepare()
        self.assertEqual(self.table('DimDevice'), devices_before)
        rows = self.table('FactEndpointAnalyticsDevice')
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['DeviceId'], 'md1')
        self.assertEqual(rows[0]['EndpointAnalyticsScore'], '')  # not zero, not previous 70
        self.assertEqual(rows[0]['AppReliabilityScore'], '90.0')
        manifest = pipeline.load_json(self.output / pipeline.MANIFEST)
        published = next(p for p in manifest['SourceEvidence']['ProducerReceipts']
                         if p['Scope'] == 'CMDB:analytics')
        self.assertEqual(published['Qualifications'], [qualification])

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

    def test_ad_computers_without_dns_remain_identifiable_and_countable(self):
        self.seed_ad()
        self.inputs['ad_computers']=[]
        for index,dns in enumerate((None,'','   ','test.synthetic.invalid')):
            native=dict(ObjectGUID=f'ad-c{index}',SID=f'S-1-5-21-1-2-3-{2000+index}',
                DomainName='synthetic.invalid',DistinguishedName=f'CN=Computer{index},DC=synthetic,DC=invalid',
                Name=f'Computer{index}',DNSHostName=dns,Enabled='false',OperatingSystem='')
            self.inputs['ad_computers'].append(native)
            self.inputs['ad_objects'].append(dict(ObjectGUID=native['ObjectGUID'],ObjectSID=native['SID'],
                DistinguishedName=native['DistinguishedName'],ObjectClass='computer'))
        self.write_inputs();self.prepare()
        computers=self.table('ADComputerSource')
        self.assertEqual(len(computers),4)
        self.assertEqual(sum(not row['DNSHostName'] for row in computers),3)
        self.assertEqual(len({row['TenantADComputerKey'] for row in computers}),4)
        for row,native in zip(computers,self.inputs['ad_computers']):
            self.assertEqual(row['ObjectGUID'],native['ObjectGUID'])
            self.assertEqual(row['ObjectSID'],native['SID'])
            self.assertEqual(row['Name'],native['Name'])
            self.assertEqual(row['TenantDeviceKey'],'')

    def test_ad_duplicate_sid_does_not_fabricate_two_cloud_links(self):
        self.seed_ad();self.inputs['ad_users'].append(dict(self.inputs['ad_users'][0],ObjectGUID='duplicate'))
        self.write_inputs();self.prepare()
        self.assertTrue(all(not r['TenantUserKey'] and r['CloudMatchStatus']=='Ambiguous native SID' for r in self.table('ADUserSource')))

    def test_ad_groups_link_group_360_only_through_unique_native_sid(self):
        self.seed_ad();self.inputs['group_scope'][0]['OnPremisesSecurityIdentifier']='S-1-5-21-1-2-3-513'
        self.write_inputs();self.prepare()
        self.assertTrue(self.table('ADGroupSource')[0]['TenantGroupKey'])
        self.assertTrue(all(r['TenantGroupKey'] for r in self.table('ADMembership')))

    def test_ad_builtin_group_sid_repeated_across_domains_retains_both_memberships(self):
        self.seed_ad()
        dn=self.inputs['ad_objects'][0]['DistinguishedName']
        for guid,domain in [('builtin-a','domain-a'),('builtin-b','domain-b')]:
            self.inputs['ad_groups'].append(dict(ObjectGUID=guid,objectSid='S-1-5-32-544',
                DomainName=domain,MembersJson=json.dumps([dn])))
            self.inputs['ad_members'].append(dict(self.inputs['ad_members'][0],
                GroupObjectGUID=guid,GroupSID='S-1-5-32-544'))
        self.write_inputs();self.prepare()
        members=self.table('ADMembership')
        self.assertEqual(len(members),4)
        self.assertEqual(len({r['TenantADMembershipKey'] for r in members}),4)
        builtin=[r for r in members if r['GroupSID']=='S-1-5-32-544']
        self.assertEqual({r['GroupObjectGUID'] for r in builtin},{'builtin-a','builtin-b'})
        self.assertEqual(len({r['TenantADGroupKey'] for r in builtin}),2)
        self.assertTrue(all(r['TenantADObjectKey'] for r in builtin))

    def test_ad_true_duplicate_membership_preserves_last_output(self):
        self.seed_ad();self.write_inputs()
        def reject():
            self.inputs['ad_members'].append(dict(self.inputs['ad_members'][0]))
            self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Duplicate immutable key')

    def test_ad_direct_membership_cannot_omit_group_guid(self):
        self.seed_ad();self.write_inputs()
        def reject():
            self.inputs['ad_members'][0]['GroupObjectGUID']=''
            self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Blank immutable key')

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

    def test_group_properties_and_acquisition_use_workplace_scope_not_licensing(self):
        self.inputs['groups'][0].update(DisplayName='Independent licensing name',MailEnabled='true',
                                       SecurityEnabled='false',GroupTypes='Unified')
        self.inputs['group_scope'][0].update(DisplayName='Membership cohort name',GroupTypes='DynamicMembership')
        self.write_inputs();self.prepare()
        group=self.table('DimGroup')[0]
        self.assertEqual(group['DisplayName'],'Membership cohort name')
        self.assertEqual(group['MailEnabled'],'false');self.assertEqual(group['SecurityEnabled'],'true')
        self.assertEqual(group['GroupTypes'],'DynamicMembership')
        self.assertEqual(group['SourceCollectedDateTime'],(NOW-dt.timedelta(minutes=2)).isoformat())

    def test_independent_licensing_catalog_drift_does_not_fabricate_empty_groups(self):
        self.inputs['groups']=[{'GroupId':'later-group','DisplayName':'Later group'}]
        self.inputs['license_paths'][0]['AssignedByGroupId']='later-group'
        self.write_inputs();self.prepare()
        self.assertEqual([r['SourceGroupId'] for r in self.table('DimGroup')],['g1'])
        path=self.table('LicenseAssignmentPath')[0]
        self.assertEqual(path['AssignedByGroupId'],'later-group')
        self.assertEqual(path['TenantGroupKey'],'');self.assertEqual(path['GroupLinkStatus'],'Unresolved')
        self.assertTrue(any(r['FindingType']=='UnresolvedAssignmentGroup' for r in self.table('EntityFinding')))
        comparison=pipeline.load_json(self.output/pipeline.MANIFEST)['PreparationQualifications']['EntraGroupCatalogComparison']
        self.assertEqual([comparison[k] for k in ('CatalogGroups','LicensingCatalogGroups',
                                                'OnlyInMembershipCatalog','OnlyInLicensingCatalog')],[1,1,1,1])

    def test_group_catalog_mixed_row_runs_preserve_last(self):
        def reject():
            self.inputs['group_scope'].append(dict(self.inputs['group_scope'][0],GroupId='g2',RunId='other-run'))
            self.write_inputs();self.prepare()
        self.unchanged_after(reject,'mixed row run')

    def test_group_runtime_row_id_is_not_the_receipt_transaction_id(self):
        self.inputs['group_scope'][0]['RunId']='runtime-row-id'
        self.inputs['group_scope'][0]['MemberCount']='1'
        self.inputs['group_members']=[dict(GroupId='g1',MemberId='u1',MemberType='#microsoft.graph.user',
            MembershipKind='Direct',CollectionStatus='Collected',RunId='runtime-row-id')]
        self.write_inputs();self.prepare()
        self.assertEqual(len(self.table('EntraGroupMembership')),1)

    def test_group_catalog_timestamps_must_be_within_the_producer_and_before_membership(self):
        for field,value,pattern in [('GroupCollectedAtUtc',(NOW-dt.timedelta(hours=1)).isoformat(),'outside producer'),
                                    ('CollectedAtUtc',(NOW+dt.timedelta(minutes=1)).isoformat(),'outside producer'),
                                    ('GroupCollectedAtUtc','unqualified','explicit timezone'),
                                    ('GroupCollectedAtUtc',NOW.isoformat(),'acquired after')]:
            with self.subTest(field=field,value=value):
                original=copy.deepcopy(self.inputs['group_scope'][0])
                def reject():
                    self.inputs['group_scope'][0]['CollectedAtUtc']=(NOW-dt.timedelta(minutes=1)).isoformat()
                    self.inputs['group_scope'][0][field]=value;self.write_inputs();self.prepare()
                self.unchanged_after(reject,pattern)
                self.inputs['group_scope'][0]=original;self.write_inputs()

    def test_member_run_and_timestamp_must_match_the_group_scope(self):
        self.inputs['group_scope'][0]['MemberCount']='1'
        member=dict(GroupId='g1',MemberId='u1',MemberType='#microsoft.graph.user',
                    MembershipKind='Direct',CollectionStatus='Collected')
        self.inputs['group_members']=[member]
        for change in [dict(RunId='other-run'),dict(CollectedAtUtc=(NOW-dt.timedelta(minutes=1)).isoformat())]:
            with self.subTest(change=change):
                def reject():
                    self.inputs['group_members']=[dict(member,**change)];self.write_inputs();self.prepare()
                self.unchanged_after(reject,'membership run or acquisition')
                self.inputs['group_members']=[member];self.write_inputs()

    def test_orphan_membership_outside_scope_still_preserves_last(self):
        def reject():
            self.inputs['group_members']=[dict(GroupId='outside',MemberId='u1',MemberType='#microsoft.graph.user',
                                              MembershipKind='Direct',CollectionStatus='Collected')]
            self.write_inputs();self.prepare()
        self.unchanged_after(reject,'Unqualified Entra group membership')

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

    def test_orphan_application_device_is_retained_with_qualification(self):
        self.inputs['app_relations'][0]['DeviceId']='unknown';self.write_inputs()
        result=self.prepare()
        self.assertEqual(result['Status'],'Prepared')
        relation=next(r for r in self.table('FactDeviceApplication') if r['AppId']=='a1')
        self.assertEqual(relation['ManagedDeviceId'],'unknown')
        self.assertEqual(relation['DeviceLinkStatus'],'Unresolved')
        self.assertEqual(len(self.table('DimIntuneManagedDevice')),2)

    def test_partial_application_modes_preserve_last(self):
        def reject():
            self.inputs['apps'][0]['RelationCollectionScope']='Top';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'All-mode')

    def test_application_count_mismatch_is_qualified_not_rejected(self):
        self.inputs['apps'][0]['DeviceCount']='7';self.write_inputs();self.prepare()
        app=next(r for r in self.table('DimDetectedApplication') if r['AppId']=='a1')
        self.assertEqual(app['ReportedDeviceCount'],'7')
        self.assertEqual(app['ExactRelatedDeviceCount'],'1')
        self.assertEqual(app['RelationshipCoverageStatus'],'Relation count differs')

    def test_group_completion_mismatch_preserves_last(self):
        def reject():
            self.inputs['group_scope'][0]['MemberCount']='1';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'membership count')

    def test_bad_score_preserves_last(self):
        def reject():
            self.inputs['analytics'][0]['EndpointAnalyticsScore']='101';self.write_inputs();self.prepare()
        self.unchanged_after(reject,'outside 0-100')

    def test_unavailable_scores_keep_devices_reports_and_raw_bytes(self):
        self.prepare()
        devices_before=self.table('DimDevice')
        self.inputs['analytics']=[
            dict(ReportName='EADeviceScoresV2',DeviceId='md1',EndpointAnalyticsScore='0',
                 StartupScore='-1',AppReliabilityScore='-2',WorkFromAnywhereScore='100'),
            dict(ReportName='EADevicePerformanceV2',DeviceId='md1',EndpointAnalyticsScore='',
                 StartupScore='75.5',AppReliabilityScore='-1',WorkFromAnywhereScore='-2'),
            dict(ReportName='EADeviceScoresV2',DeviceId='md2',EndpointAnalyticsScore='-2',
                 StartupScore='-1',AppReliabilityScore='42',WorkFromAnywhereScore='')]
        self.write_inputs()
        raw_hashes={p.name:pipeline.sha(p) for p in self.source.iterdir() if p.is_file()}
        self.prepare()
        self.assertEqual(self.table('DimDevice'),devices_before)
        result=self.table('FactEndpointAnalyticsDevice')
        self.assertEqual(len(result),3)
        by_key={(r['DeviceId'],r['SourceSystem']):r for r in result}
        scores=by_key[('md1','EADeviceScoresV2')]
        self.assertEqual([scores[f] for f in ('EndpointAnalyticsScore','StartupPerformanceScore',
                         'AppReliabilityScore','WorkFromAnywhereScore')],['0.0','','','100.0'])
        self.assertEqual(by_key[('md1','EADevicePerformanceV2')]['StartupPerformanceScore'],'75.5')
        self.assertEqual(by_key[('md2','EADeviceScoresV2')]['EndpointAnalyticsScore'],'')
        self.assertEqual(raw_hashes,{p.name:pipeline.sha(p) for p in self.source.iterdir() if p.is_file()})
        qualification=pipeline.load_json(self.output/pipeline.MANIFEST)['PreparationQualifications']['EndpointAnalyticsUnavailableScores']
        self.assertEqual(qualification['DistinctDevices'],2)
        self.assertEqual(qualification['ScoreCells'],6)
        self.assertEqual(qualification['ByFieldAndSentinel'],[
            {'Field':'AppReliabilityScore','Sentinel':-2,'Count':1},
            {'Field':'AppReliabilityScore','Sentinel':-1,'Count':1},
            {'Field':'EndpointAnalyticsScore','Sentinel':-2,'Count':1},
            {'Field':'StartupPerformanceScore','Sentinel':-1,'Count':2},
            {'Field':'WorkFromAnywhereScore','Sentinel':-2,'Count':1}])
        self.assertNotIn('md1',json.dumps(qualification))
        self.assertTrue(all(r['SourceCollectedDateTime']==NOW.isoformat() for r in result))

    def test_invalid_scores_in_each_field_preserve_previous_publication(self):
        for field in ('EndpointAnalyticsScore','StartupScore','AppReliabilityScore','WorkFromAnywhereScore'):
            for value in ('-3','-0.1','100.1','NaN','Infinity','-Infinity','not-a-score'):
                with self.subTest(field=field,value=value):
                    self.inputs['analytics']=[dict(ReportName='EADeviceScoresV2',DeviceId='md1',EndpointAnalyticsScore='70')]
                    self.write_inputs()
                    def reject():
                        self.inputs['analytics'][0][field]=value
                        self.write_inputs()
                        self.prepare()
                    self.unchanged_after(reject,'outside 0-100|could not convert')

    def test_score_sentinels_do_not_relax_other_numeric_evidence(self):
        from cmdb_tables import number,endpoint_score
        for value in ('-1','-2'):
            self.assertEqual(endpoint_score(value),'')
            with self.assertRaisesRegex(ValueError,'non-negative'):
                number(value)

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
                             ('group_members','RunId'),('group_members','CollectedAtUtc'),
                             ('group_scope','DisplayName'),('group_scope','MailEnabled'),('group_scope','SecurityEnabled'),
                             ('group_scope','GroupTypes'),('group_scope','OnPremisesSecurityIdentifier'),
                             ('group_scope','RunId'),('group_scope','CollectedAtUtc'),('group_scope','GroupCollectedAtUtc'),
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
