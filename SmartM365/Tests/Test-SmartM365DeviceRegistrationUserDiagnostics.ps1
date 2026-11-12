<# .VERSION 1.4.2
Offline scenarios. Native device operations, events and networking are blocked or mocked.
#>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$app=Join-Path (Split-Path $PSScriptRoot -Parent) 'Devices\DeviceRegistrationTool'
. (Join-Path $app 'SmartM365-DeviceRegistration.UserDiagnostics.ps1')
$script:checks=0
function Assert-Drt($Condition,[string]$Message) { if (-not $Condition) {throw $Message}; $script:checks++ }
$now=[DateTimeOffset]::Parse('2026-11-10T12:00:00Z')
$named=Get-DrtNamedRegistryValues -Names @('','AutoEnrollMDM') -ValueReader {param($name);if($name){1}else{'unused registry default'}}
$roundtrip=$named|ConvertTo-Json|ConvertFrom-Json
Assert-Drt ($roundtrip.AutoEnrollMDM -eq 1 -and $roundtrip.PSObject.Properties.Count -ne 0) 'Unnamed registry defaults must not break PS5.1 diagnostic JSON'
Assert-Drt (-not $named.ContainsKey('')) 'Do not export empty JSON property names'
$base=@('AzureAdPrt : YES','AzureAdPrtUpdateTime : 2026-11-10 10:00:00.000 UTC','AzureAdPrtExpiryTime : 2026-11-24 10:00:00.000 UTC','WamDefaultSet : YES','CloudTgt : YES','OnPremTgt : YES','NgcSet : NO')
$history=@('AcquirePrtDiagnostics : PRESENT','Previous Prt Attempt : 2026-11-10 11:00:00.000 UTC','Attempt Status : 0xc00000d0','User Identity : other@contoso.example','HTTP status : 400','Server Error Code : invalid_request')
$s=ConvertFrom-DrtUserStatus -Lines ($base+$history) -CurrentUpn 'user@contoso.example' -CurrentAccount 'CONTOSO\user' -Now $now
Assert-Drt ($s.State -eq 'Present') 'Another account historical failure must not invalidate current PRT'
Assert-Drt ($s.Attempts.Count -eq 1 -and $s.Attempts[0].Failed -and -not $s.Attempts[0].RelevantFailure) 'Preserve the historical failure with attribution'
Assert-Drt ($s.Fields.CloudTgt -eq 'YES' -and $s.Fields.NgcSet -eq 'NO') 'Keep available SSO fields'
Assert-Drt (-not $s.RefreshAttempted) 'Parsing must never refresh a PRT'
$current=$history -replace 'other@contoso.example','user@contoso.example'
$s=ConvertFrom-DrtUserStatus -Lines ($base+$current) -CurrentUpn 'user@contoso.example' -Now $now
Assert-Drt ($s.State -eq 'AttentionRequired') 'Current recent account failure should be identified'
$old=$current -replace '2026-11-10 11:00:00','2026-11-01 11:00:00'
Assert-Drt ((ConvertFrom-DrtUserStatus -Lines ($base+$old) -CurrentUpn 'user@contoso.example' -Now $now).State -eq 'Present') 'Old failures are not current failures'
$unknown=$current -replace '2026-11-10 11:00:00.000 UTC','unavailable'
Assert-Drt ((ConvertFrom-DrtUserStatus -Lines ($base+$unknown) -CurrentUpn 'user@contoso.example' -Now $now).State -eq 'Present') 'Unknown attempt time is not a confirmed current failure'
Assert-Drt ((ConvertFrom-DrtUserStatus -Lines ($base -replace '2026-11-24 10:00:00','2026-11-10 13:00:00') -Now $now).State -eq 'Expiring') 'Detect upcoming expiration without renewing'
Assert-Drt ((ConvertFrom-DrtUserStatus -Lines ($base -replace '2026-11-24 10:00:00','2026-11-09 10:00:00') -Now $now).State -eq 'Expired') 'Detect expired PRT without renewing'
Assert-Drt ((ConvertFrom-DrtUserStatus -Lines @('AzureAdPrt : NO') -Now $now).State -eq 'Missing') 'Missing PRT state'
Assert-Drt ((ConvertFrom-DrtUserStatus -Lines @() -Now $now).State -eq 'Unknown') 'Missing snapshot is unknown'
Assert-Drt ($null -eq (Convert-DrtUtcDate '2026-11-10 10:00:00')) 'Do not invent a timezone'
Assert-Drt ((Convert-DrtUtcDate '2026-11-10T13:00:00+01:00').UtcDateTime.Hour -eq 12) 'Normalize explicit offsets'
$two=$base+$history+@('RefreshPrtDiagnostics : PRESENT','Previous Prt Attempt : 2026-11-10 11:30:00.000 UTC','Attempt Status : 0x0','User Identity : user@contoso.example','HTTP status : 200')
$s=ConvertFrom-DrtUserStatus -Lines $two -CurrentUpn 'user@contoso.example' -Now $now
Assert-Drt ($s.Attempts.Count -eq 2 -and $s.State -eq 'Present') 'Keep acquire/refresh records separate instead of picking an arbitrary last error'
foreach ($provider in @('Deploy Authority','Cloud Authority','Local Authority')) { Assert-Drt ((Get-DrtMdmClassification -Provider $provider -Omadm $true -Task $true) -eq 'WindowsInternal') ('Ignore Windows provider '+$provider) }
Assert-Drt ((Get-DrtMdmClassification -Provider 'WMI_Bridge_SCCM_Server' -Omadm $true -Task $true) -eq 'ManagementBridge') 'SCCM bridge is not foreign MDM'
Assert-Drt ((Get-DrtMdmClassification -Provider 'MS DM Server') -eq 'Intune') 'Recognize explicit Intune provider'
Assert-Drt ((Get-DrtMdmClassification -DiscoveryUrl 'https://enrollment.manage.microsoft.com/enrollmentserver/discovery.svc') -eq 'UnconfirmedIntuneCandidate') 'Discovery URL alone is not enrollment proof'
Assert-Drt ((Get-DrtMdmClassification -DiscoveryUrl 'https://enrollment.manage.microsoft.com/discovery' -EnrollmentType '6' -Upn 'fooUser@contoso.example') -eq 'IncompleteIntuneCandidate') 'Identify incomplete candidate without cleanup'
Assert-Drt ((Get-DrtMdmClassification -DiscoveryUrl 'https://enrollment.manage.microsoft.com/discovery' -EnrollmentType '6' -Upn 'fooUser@contoso.example' -Omadm $true) -eq 'UnconfirmedIntuneCandidate') 'Do not classify OMA-DM backed candidate as an empty stub'
Assert-Drt ((Get-DrtMdmClassification -Provider 'Vendor' -Omadm $true) -eq 'OtherMdm') 'Require corroboration for external provider'
Assert-Drt ((Get-DrtMdmClassification -Provider 'Vendor') -eq 'UnconfirmedOrInternal') 'Bare vendor name is not sufficient evidence'
Assert-Drt ((Get-DrtMdmClassification -DiscoveryUrl 'https://enrollment.manage.microsoft.com.attacker.example/') -ne 'Intune') 'Do not trust a substring URL'
$windcUrl='https://discovery.dm.microsoft.com/EnrollmentConfiguration?api-version=1.0'
foreach($type in '26','0x1a','0X1A'){
 Assert-Drt ((Get-DrtMdmClassification -Provider ' Microsoft Device Management ' -DiscoveryUrl $windcUrl -EnrollmentType $type -Omadm $true -Task $true) -eq 'MicrosoftLinkedEnrollment') 'Recognize corroborated Microsoft linked enrollment'
}
foreach($url in @('http://discovery.dm.microsoft.com/EnrollmentConfiguration','https://discovery.dm.microsoft.com.attacker.example/EnrollmentConfiguration','https://discovery.dm.microsoft.com@attacker.example/EnrollmentConfiguration','https://attacker.example/discovery.dm.microsoft.com/EnrollmentConfiguration','https://discovery.dm.microsoft.com/other','https://user@discovery.dm.microsoft.com/EnrollmentConfiguration','https://discovery.dm.microsoft.com:444/EnrollmentConfiguration','')){
 Assert-Drt ((Get-DrtMdmClassification -Provider 'Microsoft Device Management' -DiscoveryUrl $url -EnrollmentType '26' -Omadm $true) -ne 'MicrosoftLinkedEnrollment') 'Do not suppress an unrelated or malformed MDM entry'
}
Assert-Drt (-not (Test-DrtMicrosoftLinkedEnrollment)) 'Missing metadata is not recognition'
Assert-Drt ((Get-DrtMdmClassification -DiscoveryUrl $windcUrl -EnrollmentType '26') -ne 'MicrosoftLinkedEnrollment') 'Endpoint/type without provider is insufficient'
Assert-Drt ((Get-DrtMdmClassification -Provider 'Microsoft Device Management' -DiscoveryUrl $windcUrl -EnrollmentType '6') -ne 'MicrosoftLinkedEnrollment') 'Different enrollment type remains unconfirmed'
Assert-Drt ((Get-DrtMdmClassification -Provider 'Vendor' -DiscoveryUrl $windcUrl -EnrollmentType '26') -eq 'OtherMdm') 'Do not hide unrelated vendor evidence'
& {
 $linked='11111111-1111-4111-8111-111111111111';$primary='22222222-2222-4222-8222-222222222222';$script:linkedOnly=$true
 function Read-DrtRegistryEvidence {
  param($Path)
  $children=@();$values=@{}
  if($Path -eq 'SOFTWARE\Microsoft\Enrollments'){$children=if($script:linkedOnly){@($linked)}else{@($linked,$primary)}}
  elseif($Path -eq ('SOFTWARE\Microsoft\Enrollments\'+$linked)){$values=@{ProviderID='Microsoft Device Management';DiscoveryServiceFullURL=$windcUrl;EnrollmentType=26}}
  elseif($Path -eq ('SOFTWARE\Microsoft\Enrollments\'+$primary)){$values=@{ProviderID='MS DM Server';EnrollmentType=6}}
  return [pscustomobject]@{Path=$Path;State='Available';Values=$values;Children=$children;Detail='Fixture'}
 }
 function Read-DrtTaskEvidence {param($Path,[switch]$Folder);[pscustomobject]@{State='Available';Path=$Path;HasTasks=$true}}
 $one=Get-DrtEnrollmentEvidence
 Assert-Drt ($one.State -eq 'Unconfirmed' -and $one.Entries[0].Classification -eq 'MicrosoftLinkedEnrollment') 'Linked enrollment alone is not proof of primary Intune'
 $script:linkedOnly=$false;$both=Get-DrtEnrollmentEvidence
 Assert-Drt ($both.State -eq 'Detected' -and @($both.Entries|Where-Object Classification -eq 'OtherMdm').Count -eq 0) 'Primary Intune plus WinDC must not produce another-MDM warning'
 function Get-DrtUserDiagnostics {
  param($Result,$StatusLines,$EnrollmentEvidence)
  [pscustomobject]@{Context='User';RegistrationState='Healthy';Sso=[pscustomobject]@{State='Present';RefreshAttempted=$false;Attempts=@()};Enrollment=$both;Policies=[pscustomobject]@{MdmState='Enabled';AutomaticJoinTask=[pscustomobject]@{State='Available'}};Connectivity=@();Events=@();Coverage='Complete'}
 }
 $temp=Join-Path ([IO.Path]::GetTempPath()) ('Drt-WinDC-fixture-'+[guid]::NewGuid().ToString('N'));New-Item -ItemType Directory $temp|Out-Null
 try{
  $r=[pscustomobject]@{RunId='fixture';DsregSnapshotPath='';IntuneEnrolled=$true;MdmNonIntuneEnrollmentDetected=$true;MdmNonIntuneEnrollmentIds=$linked;MdmUnconfirmedIntuneEnrollmentIds='';NextAction=''}
  Add-DrtUserDiagnostics -Result $r -OutputDir $temp -EnrollmentEvidence $both
  Assert-Drt (-not $r.MdmNonIntuneEnrollmentDetected -and -not $r.MdmNonIntuneEnrollmentIds) 'Correct both false alert flag and misleading foreign-MDM ID list'
  Assert-Drt ($r.MdmMicrosoftLinkedEnrollmentIds -eq $linked -and $r.IntuneEnrolled) 'Preserve linked enrollment identity separately from primary Intune'
  $lines=Get-DrtSummaryLines $r.UserDiagnostics
  Assert-Drt (@($lines|Where-Object {$_ -match '^\[INFO\].*WinDC/MMP-C'}).Count -eq 1 -and @($lines|Where-Object {$_ -match '^\[WARNING\].*WinDC'}).Count -eq 0) 'Report presence informationally without claiming health'
  $roundtrip=Get-Content $r.UserDiagnosticsPath -Raw|ConvertFrom-Json
  Assert-Drt ($roundtrip.Enrollment.Entries[0].EnrollmentType -eq '26' -and $roundtrip.Enrollment.Entries[0].Classification -eq 'MicrosoftLinkedEnrollment') 'Keep raw type/classification in structured evidence'
 }finally{Remove-Item -LiteralPath $temp -Recurse -Force}
}
function Read-DrtRegistryEvidence {param($Path);[pscustomobject]@{Path=$Path;State=$script:registryState;Values=@{};Children=@();Detail='Offline fixture'}}
function Read-DrtTaskEvidence {param($Path,[switch]$Folder);[pscustomobject]@{Path=$Path;State='AccessDenied';HasTasks=$false;Detail='Offline ACL refusal'}}
$script:registryState='AccessDenied'
$e=Get-DrtEnrollmentEvidence
Assert-Drt ($e.State -eq 'Unknown' -and $e.Coverage -eq 'Partial') 'Refused root registry inspection must not become no enrollment'
Assert-Drt ((Get-DrtPolicyEvidence).MdmState -eq 'Unknown') 'Refused policy reads remain unknown'
Assert-Drt ((Get-DrtPolicyEvidence).AutomaticJoinTask.State -eq 'AccessDenied') 'Preserve task visibility refusal'
$script:registryState='Missing'
Assert-Drt ((Get-DrtEnrollmentEvidence).State -eq 'NotDetected') 'Observed missing root is distinct from ACL refusal'
& {
    $script:fixtureProvider='';$script:fixtureUrl='https://enrollment.manage.microsoft.com/discovery'
    function Read-DrtRegistryEvidence {
        param($Path)
        if ($Path -eq 'SOFTWARE\Microsoft\Enrollments') {return [pscustomobject]@{Path=$Path;State='Available';Children=@('00000000-0000-4000-8000-000000000001');Values=@{}}}
        if ($Path -like 'SOFTWARE\Microsoft\Enrollments\*') {return [pscustomobject]@{Path=$Path;State='Available';Children=@();Values=@{ProviderID=$script:fixtureProvider;DiscoveryServiceFullURL=$script:fixtureUrl;EnrollmentType='6';UPN='fooUser@contoso.example'}}}
        return [pscustomobject]@{Path=$Path;State='Missing';Children=@();Values=@{}}
    }
    function Read-DrtTaskEvidence {param($Path,[switch]$Folder);[pscustomobject]@{Path=$Path;State='NotVisible';HasTasks=$false;Detail='Invisible task is not proven absent'}}
    $partial=Get-DrtEnrollmentEvidence
    Assert-Drt ($partial.State -eq 'Unknown' -and $partial.Entries[0].Classification -eq 'UnconfirmedIntuneCandidate') 'Invisible tasks must not confirm an incomplete enrollment stub'
    $script:fixtureProvider='MS DM Server'
    $partial=Get-DrtEnrollmentEvidence
    Assert-Drt ($partial.State -eq 'Detected' -and $partial.Coverage -eq 'Partial') 'Keep positive Intune evidence even when complementary evidence is inaccessible'
    $script:fixtureProvider='WMI_Bridge_SCCM_Server';$script:fixtureUrl=''
    $bridge=Get-DrtEnrollmentEvidence
    Assert-Drt ($bridge.State -eq 'NotDetected' -and $bridge.Entries[0].Classification -eq 'ManagementBridge') 'SCCM bridge must not create foreign-MDM warnings'
}
$cfg=[pscustomobject]@{OutputRoot='custom';Nested=[pscustomobject]@{Enabled=$false}}
$template=[pscustomobject]@{OutputRoot='default';NewKey=2;Nested=[pscustomobject]@{Enabled=$true;Missing='added'}}
$added=New-Object Collections.Generic.List[string]
Sync-DrtConfigDefaults -Config $cfg -Template $template -Added $added
Assert-Drt ($cfg.OutputRoot -eq 'custom' -and -not $cfg.Nested.Enabled) 'Configuration synchronization preserves explicit local values'
Assert-Drt ($cfg.Nested.Missing -eq 'added' -and $added.Count -eq 2) 'Synchronize nested missing keys'
function Get-WinEvent { throw [UnauthorizedAccessException]::new('Offline event ACL refusal') }
Assert-Drt ((Get-DrtRecentEventEvidence -LogName Fixture -Since ([datetime]::Now)).State -eq 'Unavailable') 'Event access refusal is not an empty successful log'

# Exercise the actual User entry function, including its repair-switch suppression and finally hook.
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $app 'SmartM365-DeviceRegistration-Tool.ps1'),[ref]$tokens,[ref]$errors)
Assert-Drt ($errors.Count -eq 0) 'Main source parses'
foreach($f in $ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.FunctionDefinitionAst]}) { Invoke-Expression $f.Extent.Text }
$script:registryState='AccessDenied'
function Test-DrtEndpointConnectivity {param($HostName);[pscustomobject]@{Host=$HostName;Dns='Resolved';Tcp443='Reachable';Detail='Offline fixture; current-user context only'}}
function Get-DrtRecentEventEvidence {param($LogName,$Since);[pscustomobject]@{Log=$LogName;State='Unavailable';Events=@();Detail='Offline ACL refusal'}}
function Get-CimInstance {param($ClassName);[pscustomobject]@{PartOfDomain=$true;Domain='contoso.example'}}
function Get-CurrentUserAdminState {[pscustomobject]@{UserName='CONTOSO\user';UserUpn='user@contoso.example';IsDomainUser=$true;IsLocalAdmin=$false;IsElevated=$false}}
function Get-IntuneEnrollmentState {[pscustomobject]@{Enrolled=$false;Evidence='';Confidence='None'}}
function Get-MdmEnrollmentState {[pscustomobject]@{IntuneEnrollmentDetected=$false;NonIntuneMdmEnrollmentDetected=$false;UnconfirmedIntuneEnrollmentIds=''}}
function Get-DeviceRegistrationPolicyState { [pscustomobject]@{MdmPolicyStatus='MissingOrDisabled';MdmPolicyStatusMessage='MDM policy is missing.';PolicyIssues='MDM auto-enrollment policy is missing or disabled. | Automatic-Device-Join scheduled task is missing.'} }
function Write-SmartM365DeviceRegistrationLog {param($Message)}
function Invoke-DeviceRegistrationFileRetention {param($Paths,$KeepLatest)}
function Start-AutomaticDeviceJoinTask {throw 'User mode reached join trigger'}
function Start-IntuneAutoEnrollment {throw 'User mode reached enrollment trigger'}
function Remove-DeviceRegistrationMdmEnrollment {throw 'User mode reached MDM cleanup'}
function Invoke-DeviceRegistrationPrtRefresh {throw 'User mode reached PRT refresh'}
function Start-Process {throw 'User diagnostics reached process launcher'}
$scratch=Join-Path ([IO.Path]::GetTempPath()) ('DrtOffline-'+[guid]::NewGuid().ToString('N'))
if (-not [IO.Path]::GetFullPath($scratch).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)) {throw 'Fixture path is outside temp root'}
New-Item -ItemType Directory $scratch|Out-Null
try{
 $snapshot=Join-Path $scratch 'status.txt';$base+$history|Set-Content $snapshot -Encoding UTF8
 function Initialize-SmartM365DeviceRegistrationOutput {param($Root);[pscustomobject]@{Output=$scratch;RunLog=(Join-Path $scratch 'run.log');CsvPath=(Join-Path $scratch 'result.csv')}}
 function Get-DsregStatusSnapshot {param($DsregcmdPath,$OutputDir,$Phase);[pscustomobject]@{AzureAdJoined='YES';DeviceAuthStatus='SUCCESS';KeySignTest='PASSED';AzureAdPrt='YES';SnapshotPath=$snapshot;ExitCode=0}}
 $oldCsv=Join-Path $scratch 'schema-upgrade.csv'
 [pscustomobject]@{RunId='old';Status='HEALTHY'}|Export-Csv $oldCsv -NoTypeInformation
 Write-AtomicCsvAppend -Path $oldCsv -RunIdValue 'new' -RowObject ([pscustomobject]@{RunId='new';Status='HEALTHY';UserDiagnosticsPath='fictional-evidence.json';DiagnosticCoverage='Partial'})
 $rows=@(Import-Csv $oldCsv)
 Assert-Drt ($rows.Count -eq 2 -and $rows[0].RunId -eq 'old' -and $rows[0].Status -eq 'HEALTHY') 'CSV schema enrichment must preserve historical rows'
 Assert-Drt ($rows[1].UserDiagnosticsPath -eq 'fictional-evidence.json' -and $rows[1].DiagnosticCoverage -eq 'Partial') 'CSV schema enrichment must retain new diagnostic columns'
 function Write-AtomicCsvAppend {param($Path,$RowObject,$RunIdValue);Assert-Drt ($RowObject.DiagnosticCoverage -eq 'Partial') 'Finally diagnostics are included in CSV';Assert-Drt (-not $RowObject.PSObject.Properties['UserDiagnostics']) 'Detailed object is kept in sidecar, not stringified into CSV';$RowObject|Export-Csv $Path -NoTypeInformation}
 $Script:ToolName='Offline';$Script:Version='1.3';$Script:RunId='fixture';$Script:ComputerName='CTSO-TEST';$Script:RequireDomainConnectivity=$false;$Script:DeviceProfile='Auto';$script:SelectedLanguage='en'
 $result=Invoke-SmartM365DeviceRegistrationTool -Mode User -OutputRoot $scratch -NoTranscript -AllowDsregLeave -TriggerJoin -TriggerIntuneAutoEnrollment -AllowRemoveStaleIntuneEnrollment -AllowRemoveNonIntuneMdmEnrollment -RetryCount 1
 Assert-Drt ($result.ExitCode -eq 0 -and $result.Action -eq 'None') 'Healthy User execution remains successful without repair'
 Assert-Drt ($null -eq $result.IntuneEnrolled -and $result.LocalMdmState -eq 'Unknown') 'Partial evidence does not claim unenrolled'
 Assert-Drt (-not $result.LeaveAttempted -and -not $result.MdmCleanupAttempted -and -not $result.IntuneAutoEnrollAttempted) 'All repair switches are suppressed in User mode'
 Assert-Drt ($result.NextAction -eq 'CONTACT_SUPPORT_VERIFY_EVIDENCE') 'User guidance must not request an enrollment trigger'
 Assert-Drt (Test-Path $result.UserDiagnosticsPath) 'Structured evidence exported to isolated output'
 $decoded=$result|ConvertTo-Json -Depth 12|ConvertFrom-Json
 Assert-Drt ($decoded.UserDiagnostics.Context -eq 'User') 'Full enriched result must roundtrip under the endpoint PowerShell JSON reader'
 $text=Format-DeviceRegistrationResultText $result
 Assert-Drt ($text -match 'User diagnostics' -and $text -match 'TCP443=Reachable' -and $text -match 'User SSO: Present') 'GUI detail formatter includes new read-only sections'
 $findings=Get-DeviceRegistrationFindingLines $result
 Assert-Drt (-not @($findings|Where-Object {$_ -match 'Device is not enrolled in Intune|MDM policy is missing|scheduled task is missing'}).Count) 'Summary must not claim absence when evidence is refused'
 Assert-Drt (@($findings|Where-Object {$_ -match 'not verifiable|not attributed'}).Count -gt 0) 'Summary explains uncertainty and historical attribution'
}finally{Remove-Item -LiteralPath $scratch -Recurse -Force}
Write-Output ('PASS: {0} offline diagnostic assertions; no live device or network calls.' -f $script:checks)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC/lD1QB1tPJQUS
# I/ulm8A3o4tJwHLHXRCgeqUQtLm2cKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIOkZDR9X7FsUXH6qllE7bSR6BthFPBAYNDrHtAIlx7dZMA0GCSqG
# SIb3DQEBAQUABIIBgFaQ9sb+R/EmSzzs6N3AUjW1iPzR5Tk6tbzOPeubGmTur7F6
# 6z+dWHauf+/IVK7jAFNqMcFgAOXGHr/e1MNvnnHfS1/CJtVqZ5e0YfbjV9mrSxrh
# zJ+l0OAEgoSyCkml6028GhRAvJnB0iF1SohGPJCZZ0BgY7ftRA/+1kgOE8vb4OLr
# jRjcyEdsogvBot0RMJ2qEujQKa+zKfsfMwyzepIOrrzEHFpjJqeJn0YwJuLaiLTI
# 2GgVT7ErYKteTnkwFvJ+2WZ8/hEL8KASaDvt81C69N51JNqp/yFeE5XgAYibg5PX
# 0MQbuDQtvqLYA/IUVbkAbMooBa9OBSCAEPWfyfsp6XoszxVRK5WqtYZxjUKALCIB
# V7CaZnp/2ZaDaTG6e5Jc0WUTeaNVTIgXsI8M4GN0uOU8J+6lrxrBLESCfbO/50zg
# 0gbga/DySG5YUeJNnHtsa4lzAgJXuaSPL9llLBVZ2fEzld3B0q3FdrsTPGMfgTny
# NbIi+Am8GM+yQqrW0KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTIyMDAy
# MTdaMC8GCSqGSIb3DQEJBDEiBCC5PeNjOhB6jlFsei+6a2d3VPOrJwJyrM+U1jqR
# mI00NTANBgkqhkiG9w0BAQEFAASCAgAh45wexS8s2DlnmEQ6OHDeTNuWqB/plCnO
# 5qMb4MXfk5VB/fm237f+c5gsEpDECyiwIjchywKX5ChHADljMbvfk/qh+K3g6wKE
# w7I3/nYh1LMuT9qJ2YKxlZTzF5fAdAtRw28Dh2WMRrbMaANRV29dc8tb5xSQ2obC
# fpDLrsXaMuYkXbBOgS306yg9qqlq6sK4WOq5meqVbXK4OqxSNQ9fXGXtHJ9K9bay
# uqTpbLpy1pRkYBKxAjLj7gJv/vXA6xemGGpQ9/4RL7Z/qcsdLkn6MR2xj3/Q3aM+
# /tcUOPsIWOPzFQgYemr+dGjLu4iwldUfSp4s7p4CajLA/JGxKmL/xklWiD5ThoqN
# tvuKO90SrpiFkR8cahg/fTnCmxO6nTkRpGI7vYtnftZ1ZSCJ5pcTTxImm4WsOV8s
# u+P1B7SQ3B6VGvDBLK4hLqTyDo/iK+LXzR47/FDQ51EBYrZcyX2s8zMOczUwxubE
# qZ9v/yaqAecQn7N4a6hyoc/+jozGa0yX7hR1Ai/VMYzc2HHBWTusGjWvADxdS8dr
# spQQG9Tx7RAdk6MFVAXgWL3AQ68ZLaHo4TlrpeIUKn73B0hPXmPuuuYXCs4FpWZ7
# usBImGH71UUJqdD16RBhfNErs6EoT9Y1HLvYjoI9FJ/5F1J5+szdoF9MvtkZ1PeE
# BS27JBt6nA==
# SIG # End signature block
