"""Synthetic quarantine proof, whole-pipeline and derived-reader regression tests."""
import copy
import csv
import hashlib
import io
import json
import subprocess
import unittest
from contextlib import redirect_stdout
from pathlib import Path

import workplace_native_build as native
import workplace_native_sources as sources
import workplace_device_identity as device_identity
import test_native_evidence as fixture_module


class DeviceIdentityTests(unittest.TestCase):
    def setUp(self):
        self.case = fixture_module.NativeEvidenceTests(methodName='runTest')
        self.case.setUp()
        self.addCleanup(self.case.doCleanups)
        self.fixture = self.case.fixture
        self.source = self.fixture.source
        self.producer = next(p for p in sources.load_json(sources.REGISTRY)['Producers'] if p['Script'] == device_identity.PRODUCER)
        self.receipt_path = self.source / self.producer['Receipt']
        self.artifact = self.source / device_identity.FILE

    def install(self, resolution='Identity conflict', attempts=3):
        self.fixture.inputs['hardware'][0]['azureADDeviceId'] = 'contested-hardware-id'
        self.fixture.write_inputs()
        self.managed = list(sources.rows(self.source / 'Intune_ManagedDevices_All.csv'))
        self.hardware = list(sources.rows(self.source / 'Intune_DeviceHardware_All.csv'))
        def observation(number):
            m, h = (dict((k,v) for k,v in rows[0].items() if k not in device_identity.IDENTITY_FIELDS) for rows in (self.managed,self.hardware))
            return {'Attempt':number,'Managed':m,'Hardware':h,'NativeManaged':dict(Id=m['ManagedDeviceId'], Model='Complete synthetic native observation')}
        observations = [observation(n) for n in range(attempts+1)]
        if resolution == 'Converged':
            observations[-1]['Hardware']['azureADDeviceId'] = observations[-1]['Managed']['AzureADDeviceId']
            self.hardware[0]['azureADDeviceId'] = self.managed[0]['AzureADDeviceId']
            self.write_csv('Intune_DeviceHardware_All.csv', self.hardware)
        proof = sources.load_json(self.receipt_path)
        self.doc = dict(ContractVersion='0.5.0', TenantKey='synthetic',RunId=proof['RunId'],ManagedDeviceId='md1',Resolution=resolution,Observations=observations)
        self.write_artifact()
        proof.update(Owner='SmartInventory-SourceReceipt',ContractVersion='1.2',ScriptVersion='1.21',
                     RequiredFiles=self.producer['Files']+[device_identity.FILE], ScopeQualification='ConsumerScope',ConsumerScopeQualified=True)
        extra = dict(proof['Files'][0], File=device_identity.FILE)
        proof['Files'].append(extra)
        for record in proof['Files']:
            record['ScriptVersion']='1.21'
        self.receipt_path.write_text(json.dumps(proof), encoding='utf-8')
        self.resign()

    def write_csv(self, name, data):
        path = self.source / name
        with path.open('w',encoding='utf-8-sig',newline='') as stream:
            writer=csv.DictWriter(stream,fieldnames=list(data[0]));writer.writeheader();writer.writerows(data)

    def write_artifact(self):
        raw=json.dumps(self.doc,separators=(',',':'))
        self.write_csv(device_identity.FILE,[dict(**fixture_module.IDENTITY,RunId=self.doc['RunId'],ManagedDeviceId=self.doc['ManagedDeviceId'],
            ContractVersion='0.5.0',Resolution=self.doc['Resolution'],EvidenceJson=raw,EvidenceSHA256=hashlib.sha256(raw.encode()).hexdigest().upper())])

    def resign(self):
        proof=sources.load_json(self.receipt_path)
        for record in proof['Files']:
            record.update(SHA256=sources.sha(self.source/record['File']),Rows=len(list(sources.rows(self.source/record['File']))))
        self.receipt_path.write_text(json.dumps(proof),encoding='utf-8')

    def audit(self):
        with redirect_stdout(io.StringIO()):
            return self.case.audit()

    def build(self):
        with redirect_stdout(io.StringIO()):
            return native.build(self.source,self.case.output,self.audit())

    def table(self, name):
        return list(sources.rows(self.case.output/(name+'.csv')))

    def test_legacy_consistent_input_equivalence_is_preserved(self):
        result=self.build()
        self.assertEqual(result['Details']['DeviceIdentity']['ConflictCount'],0)
        self.assertEqual(result['Details']['DeviceIdentity']['ReReadCount'],0)
        self.assertEqual(len(self.table('DimIntuneManagedDevice')),2)

    def test_persistent_conflict_retains_native_identity_without_any_disputed_join(self):
        self.install()
        result=self.build()
        devices=self.table('DimDevice')
        conflict=next(r for r in devices if r['AssociationStatus']=='Identity conflict')
        self.assertEqual(conflict['SourceDeviceId'],'intune:md1')
        self.assertEqual(conflict['SourceSystems'],'Intune')
        self.assertEqual(conflict['CountryCode'],'')
        self.assertEqual(conflict['AssociatedAccount'],'Not resolved')
        self.assertEqual(len(devices),3)  # Independent Entra + technical Intune + unaffected Intune.
        self.assertEqual(len(self.table('DimIntuneManagedDevice')),2)
        self.assertEqual(len(self.table('DeviceHardware')),2)
        related=[r for r in self.table('FactUserDeviceRelationship') if r['TenantDeviceKey']==conflict['TenantDeviceKey']]
        self.assertEqual(related,[])
        self.assertEqual(next(r for r in self.table('FactDeviceCompliance') if r['TenantDeviceKey']==conflict['TenantDeviceKey'])['ComplianceState'],'Compliant')
        self.assertTrue(any(r['SelectionStatus']=='Identity conflict' for r in self.table('DeviceSource')))
        self.assertEqual(len([r for r in self.table('FactDataQuality') if r['FindingType']=='DeviceIdentityConflict']),1)
        for row in self.table('SourceHealth'):
            if row['SourceName'] in self.producer['Files']:
                self.assertEqual(row['Status'],'SuccessWithWarnings');self.assertIn('count=1',row['Coverage'])
        self.assertEqual(result['Details']['DeviceIdentity']['ManagedIdentities'],2)
        self.assertEqual(result['Details']['DeviceIdentity']['ConflictCount'],1)
        self.assertNotIn('md1',json.dumps(result['Details']['DeviceIdentity']))
        self.assertFalse(result['Publication'])

    def test_convergence_each_attempt_is_verified(self):
        for attempts in (1,2,3):
            with self.subTest(attempts=attempts):
                self.install('Converged',attempts)
                audit=self.audit()
                self.assertEqual(audit['Evidence']['DeviceIdentity']['ConvergedCount'],1)
                self.assertEqual(audit['Evidence']['DeviceIdentity']['ReReadCount'],attempts)

    def test_modern_consistent_empty_artifact_is_required_and_accepted(self):
        self.install('Converged',1)
        with self.artifact.open('w',encoding='utf-8',newline='') as stream:
            csv.writer(stream).writerow(sources.load_json(sources.CONTRACT)['deviceIdentity']['columns'])
        self.resign()
        self.assertEqual(self.audit()['Evidence']['DeviceIdentity']['ConflictCount'],0)
        self.artifact.unlink()
        with self.assertRaises((ValueError,FileNotFoundError)):
            self.audit()

    def test_unreported_divergence_stays_blocking(self):
        self.fixture.inputs['hardware'][0]['azureADDeviceId']='different'
        self.fixture.write_inputs()
        with self.assertRaisesRegex(ValueError,'correlation differs'):
            self.audit()

    def test_all_negative_proof_variants_stay_blocking(self):
        mutations={
            'tenant':lambda d:d.update(TenantKey='other-tenant'),
            'run':lambda d:d.update(RunId='other-run'),
            'id':lambda d:d.update(ManagedDeviceId='other-id'),
            'version':lambda d:d.update(ContractVersion='0.4.0'),
            'incomplete':lambda d:d['Observations'].pop(),
            'duplicate attempt':lambda d:d['Observations'][2].update(Attempt=1),
            'foreign managed':lambda d:d['Observations'][1]['Managed'].update(ManagedDeviceId='foreign'),
            'foreign hardware':lambda d:d['Observations'][1]['Hardware'].update(ManagedDeviceId='foreign'),
            'failed acquisition':lambda d:d['Observations'][1]['Hardware'].update(CollectionStatus='Failed'),
            'incomplete observation':lambda d:d['Observations'][1]['Hardware'].pop('model'),
            'invalid intermediate bytes':lambda d:d['Observations'][1]['Hardware'].update(physicalMemoryInBytes='-1'),
            'invalid intermediate storage':lambda d:d['Observations'][1]['Hardware'].update(freeStorageSpaceInBytes='999999999999999'),
            'outside interval':lambda d:d['Observations'][1]['Hardware'].update(CollectedAtUtc='2099-01-01T00:00:00Z'),
            'early convergence':lambda d:d['Observations'][1]['Hardware'].update(azureADDeviceId='ed1'),
            'invented convergence':lambda d:d.update(Resolution='Converged'),
            'altered published observation':lambda d:d['Observations'][0]['Managed'].update(DeviceName='fabricated'),
            'missing native observation':lambda d:d['Observations'][1].pop('NativeManaged'),
        }
        for name, mutate in mutations.items():
            with self.subTest(name=name):
                self.install();mutate(self.doc);self.write_artifact();self.resign()
                with self.assertRaises((ValueError,KeyError)):
                    self.audit()
                self.assertFalse(list(self.case.output.iterdir()))

    def test_sha_row_duplicates_and_downgrade_are_rejected(self):
        for variant in ('file hash','payload hash','duplicate','downgrade','wrong receipt run'):
            with self.subTest(variant=variant):
                self.install()
                if variant=='file hash':
                    with self.artifact.open('a') as stream:stream.write('\n')
                elif variant in ('payload hash','duplicate'):
                    data=list(sources.rows(self.artifact))
                    if variant=='payload hash':data[0]['EvidenceSHA256']='A'*64
                    else:data.append(copy.deepcopy(data[0]))
                    self.write_csv(device_identity.FILE,data);self.resign()
                else:
                    proof=sources.load_json(self.receipt_path)
                    if variant=='downgrade':proof['RequiredFiles']=self.producer['Files']
                    else:next(r for r in proof['Files'] if r['File']==device_identity.FILE)['RunId']='other-run'
                    self.receipt_path.write_text(json.dumps(proof),encoding='utf-8')
                with self.assertRaises(ValueError):self.audit()

    def test_artifact_is_rechecked_after_validation_before_output(self):
        self.install();audit=self.audit()
        with self.artifact.open('a') as stream:stream.write('\n')
        with self.assertRaisesRegex(ValueError,'Source changed'):
            native.build(self.source,self.case.output,audit)
        self.assertFalse(list(self.case.output.iterdir()))

    def test_older_reader_refuses_new_required_evidence(self):
        self.install()
        with self.assertRaisesRegex(ValueError,'consumer-scope declaration mismatch'):
            fixture_module.baseline.producer_records(self.source,fixture_module.IDENTITY)

    def test_fleet_scope_retains_21470_intune_identities_and_one_conflict(self):
        self.install()
        managed,hardware=copy.deepcopy(self.managed),copy.deepcopy(self.hardware)
        for number in range(2,21470):
            m,h=copy.deepcopy(managed[1]),copy.deepcopy(hardware[1])
            mid='synthetic-'+str(number)
            m.update(ManagedDeviceId=mid,AzureADDeviceId='cloud-'+mid)
            h.update(ManagedDeviceId=mid,azureADDeviceId='cloud-'+mid)
            managed.append(m);hardware.append(h)
        self.write_csv('Intune_ManagedDevices_All.csv',managed)
        self.write_csv('Intune_DeviceHardware_All.csv',hardware)
        self.resign()
        summary=self.audit()['Evidence']['DeviceIdentity']
        self.assertEqual(summary['ManagedIdentities'],21470)
        self.assertEqual(summary['ConflictCount'],1)
        self.assertEqual(summary['ManagedIdentities']-summary['ConflictCount'],21469)

    def windows(self):
        receipt=sources.load_json(self.receipt_path)
        schema=sources.load_json(Path(native.__file__).resolve().parents[1]/'config/source-schema.json')['files']['Intune_Devices_Inventory.csv']['columns']
        row=dict.fromkeys(schema,'')
        row.update({**fixture_module.IDENTITY,'Device ID':'md1','Device name':'Same name','Azure AD Device ID':'ed1',
            'Identity Status':'Identity conflict','Identity RunId':receipt['RunId'],'Identity Contract':'0.5.0',
            'Entra ObjectId':'','Entra DeviceId':'','Entra DisplayName':'','Entra ApproximateLastSignInDateTime':'',
            'OS':'Windows','OS version':'10.0.19045','Ownership':'company','Managed by':'MDM','Compliance':'compliant','Primary user UPN':'user@synthetic.invalid'})
        rows=[row]
        self.write_csv('Intune_Devices_Inventory.csv',rows)
        receipt['Files'].append(dict(receipt['Files'][0],File='Intune_Devices_Inventory.csv'))
        self.receipt_path.write_text(json.dumps(receipt),encoding='utf-8');self.resign()
        return rows

    def test_actual_powershell_windows_and_trust_readers_show_conflicts(self):
        self.install();self.windows()
        schema=sources.load_json(Path(native.__file__).resolve().parents[1]/'config/source-schema.json')['files']
        # Legacy facets required by the actual generator; synthetic empty evidence only.
        for name in ('M365_Entra_Devices.csv','AD_Computers_AllDomains.csv','Intune_Devices_UpgradeEligibility.csv','Intune_WindowsUpdate_Status.csv',
                     'Intune_EndpointAnalytics_DevicePerformance.csv','Intune_EndpointAnalytics_StartupDevices.csv','M365_Entra_Devices_HardwareIdConflicts.csv','M365_Entra_Devices_RemovalCandidates.csv'):
            with (self.source/name).open('w',encoding='utf-8-sig',newline='') as stream:
                csv.writer(stream).writerow(schema[name]['columns'])
        scripts=Path(native.__file__).parent
        outputs=self.fixture.root/'derived';outputs.mkdir()
        for script,args in (
            ('New-DeviceInventoryEvidence.ps1',['-OutputPath',str(outputs/'device.csv'),'-SignalsOutputPath',str(outputs/'signals.csv'),'-DirectorySummaryOutputPath',str(outputs/'directory.csv')]),
            ('New-DataTrustEvidence.ps1',['-OutputPath',str(outputs/'trust.csv')])):
            result=subprocess.run(['pwsh','-NoProfile','-ExecutionPolicy','Bypass','-File',str(scripts/script),'-DataRoot',str(self.fixture.root),*args],capture_output=True,text=True,timeout=60)
            self.assertEqual(result.returncode,0,result.stdout+result.stderr)
        row=list(sources.rows(outputs/'device.csv'))[0]
        self.assertIn('Identity conflict',row['Entra Match'])
        self.assertIn('Identity conflict',row['AD to Intune Match'])
        self.assertEqual(row['Country'],'Unknown')
        directory=list(sources.rows(outputs/'directory.csv'))[0]
        self.assertIn('conflicts=1',directory['Evidence Status'])
        self.assertEqual(directory['Enabled AD PCs Managed by Intune (%)'],'')
        trust=next(r for r in sources.rows(outputs/'trust.csv') if r['File Name']=='Intune_Devices_Inventory.csv')
        self.assertEqual(trust['Decision Readiness'],'Review required')
        self.assertIn('identity conflicts: 1',trust['Evidence Status'])
        self.assertIn(device_identity.FILE,trust['Recommended Action'])

    def test_real_core_export_roundtrips_powershell_observations_and_identity_columns(self):
        root=self.fixture.root/'producer';root.mkdir()
        repository=Path(native.__file__).resolve().parents[2]
        quote=lambda p:"'"+str(p).replace("'","''")+"'"
        command=r'''
$ErrorActionPreference='Stop'
$directory=Join-Path ROOT 'SmartM365/SmartInventory/M365Inventory/IntuneInventory/Devices'
$collector=Join-Path $directory 'SmartM365-Devices-Inventory.ps1'
$core=Import-Module (Join-Path ROOT 'SmartM365/Modules/SmartM365.Core/SmartM365.Core.psd1') -Force -PassThru
& $core {function script:WriteLog {param($Message,$Level)};function script:Invoke-SmartM365SharePointCsvUpload {param($LocalFilePath)};function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv {param($SourceFiles,$TimestampedPath)}}
function WriteLog {param($Message,$Level)}
function Start-Sleep {param($Seconds)}
$global:SmartM365TenantKey='synthetic';$global:SmartM365OrganizationKey='test';$global:SmartM365EnvironmentKey='test';$global:SmartM365TenantId='synthetic-tenant'
Set-SmartM365CoreContext -TenantKey synthetic -OrganizationKey test -EnvironmentKey test -TenantId synthetic-tenant -RetentionMaxCsv 0
$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$errors)
$getter=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-SafeProperty'},$true)
Set-Item Function:Get-SafeProperty ([scriptblock]::Create($getter.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
. (Join-Path $directory 'SmartM365-DeviceHardware.ps1')
. (Join-Path $directory 'SmartM365-DeviceIdentity.ps1')
function Get-MgDeviceManagementManagedDevice {param($ManagedDeviceId,$ErrorAction) @{Id=$ManagedDeviceId;AzureADDeviceId='managed-cloud';DeviceName='Synthetic';OperatingSystem='Windows'}}
function Invoke-MgGraphRequest {param($Method,$Uri,$ErrorAction) @{id='md1';azureADDeviceId='hardware-cloud'}}
Start-SmartM365CmdbSourceReceipt -ScriptPath $collector -SourceRootPath OUTPUT
Enable-SmartM365StagedSourcePublication
$run=Get-SmartM365CurrentSourcePublication
$device=Get-MgDeviceManagementManagedDevice md1
$stamp=[datetimeoffset]::UtcNow.ToString('o')
$hardware=@{md1=(ConvertTo-SmartM365HardwareRecord md1 (Invoke-MgGraphRequest))}
$resolved=Resolve-SmartM365DeviceIdentity -Devices @($device) -Hardware $hardware -CollectedAtUtc $stamp -TenantKey synthetic -RunId $run.RunId
$sets=@(
 @{Name='Intune_ManagedDevices_All';Data=$resolved.ManagedRows;Columns=(Get-SmartM365ManagedColumns)},
 @{Name='Intune_DeviceHardware_All';Data=$resolved.HardwareRows;Columns=@($resolved.HardwareRows[0].PSObject.Properties.Name)},
 @{Name='Intune_DeviceIdentityEvidence';Data=$resolved.EvidenceRows;Columns=@('RunId','ManagedDeviceId','ContractVersion','Resolution','EvidenceJson','EvidenceSHA256')},
 @{Name='Intune_Devices_Inventory';Data=@([pscustomobject]@{'Device ID'='md1';'Device name'='Synthetic';'Azure AD Device ID'='managed-cloud';'Identity Status'='Identity conflict';'Identity RunId'=$run.RunId;'Identity Contract'='0.5.0';'Entra ObjectId'='';'Entra DeviceId'='';'Entra DisplayName'='';'Entra ApproximateLastSignInDateTime'=''});Columns=@('Device ID','Device name','Azure AD Device ID','Identity Status','Identity RunId','Identity Contract','Entra ObjectId','Entra DeviceId','Entra DisplayName','Entra ApproximateLastSignInDateTime')}
)
foreach($set in $sets){Export-SmartM365Csv -Data $set.Data -Columns $set.Columns -TimestampedPath (Join-Path OUTPUT ($set.Name+'_history.csv')) -LatestPath (Join-Path OUTPUT ($set.Name+'.csv')) -NoWeeklyHistory | Out-Null}
Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope 'CMDB:managed,hardware'
Complete-SmartM365CmdbSourceReceipt -Status Success | Out-Null
'''.replace('ROOT',quote(repository)).replace('OUTPUT',quote(root))
        result=subprocess.run(['pwsh','-NoProfile','-ExecutionPolicy','Bypass','-Command',command],capture_output=True,text=True,timeout=60)
        self.assertEqual(result.returncode,0,result.stdout+result.stderr)
        summary=device_identity.validate_windows(root)
        self.assertEqual(summary['ConflictCount'],1)
        self.assertEqual(summary['ReReadCount'],3)
        artifact=list(sources.rows(root/device_identity.FILE))[0]
        self.assertEqual({k:artifact[k] for k in fixture_module.IDENTITY},fixture_module.IDENTITY)
        root51=self.fixture.root/'producer-ps51';root51.mkdir()
        result51=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-Command',command.replace(quote(root),quote(root51))],capture_output=True,text=True,timeout=60)
        self.assertEqual(result51.returncode,0,result51.stdout+result51.stderr)
        self.assertEqual(device_identity.validate_windows(root51)['ConflictCount'],1)

    def test_windows_reader_checks_status_run_hash_identity_and_no_entra_join(self):
        self.install();data=self.windows()
        self.assertEqual(device_identity.validate_windows(self.source)['ConflictIds'],['md1'])
        for field,value in [('Identity Status','Consistent'),('Identity RunId','foreign-run'),('Entra ObjectId','contested-link'),('Azure AD Device ID','other')]:
            with self.subTest(field=field):
                bad=copy.deepcopy(data);bad[0][field]=value
                self.write_csv('Intune_Devices_Inventory.csv',bad);self.resign()
                with self.assertRaises(ValueError):device_identity.validate_windows(self.source)

    def test_exchange_and_licensing_outputs_are_identical_with_quarantine(self):
        self.build()
        names=['FactMailbox','FactMailboxHosting','FactUserLicense','LicenseAssignmentPath','FactUserServicePlan','DimLicenseSku','DimUser','DimCountry']
        before={n:sources.sha(self.case.output/(n+'.csv')) for n in names}
        self.case.output=self.fixture.root/'quarantine-output';self.case.output.mkdir()
        self.install();self.build()
        self.assertEqual(before,{n:sources.sha(self.case.output/(n+'.csv')) for n in names})

    def test_native_stage_cli_returns_nonzero_and_no_promotion_for_bad_artifact(self):
        self.install();self.doc['RunId']='wrong';self.write_artifact();self.resign()
        import sys
        result=subprocess.run([sys.executable,'-B',str(Path(native.__file__)), '--source',str(self.source),'--tenant','synthetic','--audit',str(self.fixture.root/'audit.json')],capture_output=True,text=True)
        self.assertNotEqual(result.returncode,0)
        self.assertFalse((self.fixture.root/'audit.json').exists())
        self.assertFalse((self.source/'current.json.txt').exists())


if __name__=='__main__':
    unittest.main()
