#Requires -Version 7.0
# Synthetic-only tests. No collection, Power BI, client CSVs or publication.
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'DeviceSystemEvidence.psm1') -Force
$script:checks=0
function Assert-Test([bool]$Condition,[string]$Message) {$script:checks++; if(-not $Condition){throw $Message}}
function New-SyntheticSystem([string]$Id='device-one') {
    [pscustomobject]@{TenantKey='synthetic';DeviceId=$Id;DeviceName='not-a-join-key';AzureADDeviceId='cloud';UserPrincipalName='not-a-join-key';SecureBootStatus='Enabled';BIOSVersion='example';BIOSDate='2026-01-01';FirmwareType='Uefi';PlatformScriptLastUpdate='2026-01-02T00:00:00Z';LastRebootDate='2026-01-01 10:00:00';'Last Reboot Date'='2026-01-01 10:00:00';LastBootUpTime='2026-01-01 10:00:00';LastBootTime='2026-01-01 10:00:00';IsEncrypted='False';DeviceGuardVBSState='running';DeviceGuardVBSHardwareRequirement='meetHardwareRequirements';DeviceGuardCredentialGuardState='running';ManagementState='managed';ComplianceState='compliant';LastSyncDateTime='2026-01-02T00:00:00Z';ThirdPartyAVStatus='Detected';ThirdPartyAVProducts='Example AV';ExpectedThirdPartyAVName='Example AV';ExpectedThirdPartyAVStatus='Detected';WindowsUpdateRebootRequired='False';BitLockerOSProtectionStatus='On';BitLockerOSConversionStatus='FullyEncrypted'}
}
$inventory=[pscustomobject]@{TenantKey='synthetic';'Device ID'='device-one';'Primary user UPN'='user@example.invalid';Encrypted='True';Compliance='compliant'}
$user=[pscustomobject]@{TenantKey='synthetic';'User principal name'='user@example.invalid';CountryOrRegion='Germany';'Usage location'='FR'}
$prepared=[pscustomobject]@{'Device Source ID'='device-one'}
$asOf=[datetimeoffset]'2026-01-03T00:00:00Z'
function Invoke-Synthetic([object[]]$System=@((New-SyntheticSystem)),[object[]]$Inventory=@($inventory),[object[]]$Users=@($user),[object[]]$Prepared=@($prepared),[Nullable[double]]$Age=48,[string]$Zone) {
    @(New-DeviceSystemEvidence -SystemRows $System -InventoryRows $Inventory -UserRows $Users -PreparedInventoryRows $Prepared -MaxObservationAgeHours $Age -OffsetlessGraphTimeZoneId $Zone -AsOfUtc $asOf)
}
$base=Invoke-Synthetic
Assert-Test ($base.Count -eq 1 -and $base[0].Country -eq 'Germany') 'Primary-user CountryOrRegion is not preferred.'
foreach($value in @('', 'Unknown', 'Not a country')) {
    $user.CountryOrRegion=$value
    Assert-Test ((Invoke-Synthetic)[0].Country -eq 'Unknown') 'Missing or invalid country fell back to usage location.'
}
$user.CountryOrRegion='Germany'
Assert-Test ($base[0].'Device System Key' -eq 'synthetic|device-one') 'Tenant key is not part of identity.'
Assert-Test ($base[0].'Expected AV Name' -eq 'Example AV') 'Expected product name was replaced.'
Assert-Test ($base[0].Freshness -eq 'Within policy' -and [double]$base[0].'Observation Age Hours' -eq 24) 'Configured freshness is incorrect.'
Assert-Test ($base[0].'Encryption Comparison' -eq 'Conflict') 'Encryption disagreement was concealed.'
Assert-Test ($base[0].'Last Boot Time' -eq '2026-01-01 10:00:00') 'Boot aliases were not consolidated.'
$contract=Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\config\prepared-evidence-contract.json.txt') -Raw | ConvertFrom-Json
$table=@($contract.tables | Where-Object table -eq 'Device System Evidence')[0]
Assert-Test (($base[0].PSObject.Properties.Name -join '|') -eq ((Get-DeviceSystemEvidenceColumns) -join '|')) 'Output header differs from module contract.'
Assert-Test (($table.columns.name -join '|') -eq ((Get-DeviceSystemEvidenceColumns) -join '|')) 'Prepared output contract differs from module.'
foreach ($state in @('NotCollected','Unknown','Error','NotSupported','')) {
    $r=New-SyntheticSystem; $r.ExpectedThirdPartyAVStatus=$state; $r.SecureBootStatus=$state; $r.WindowsUpdateRebootRequired=$state; $r.BitLockerOSProtectionStatus=$state
    $e=(Invoke-Synthetic -System @($r))[0]; $expected=if($state){$state}else{'Missing'}
    Assert-Test ($e.'Expected AV State' -eq $expected -and $e.'Secure Boot State' -eq $expected -and $e.'Windows Update Reboot Required' -eq $expected -and $e.'BitLocker OS Protection' -eq $expected) 'Unavailable state became confirmed adverse.'
}
$r=New-SyntheticSystem; $r.ExpectedThirdPartyAVStatus='NotDetected'; $r.ThirdPartyAVStatus='NotDetected'; $r.ThirdPartyAVProducts=''
Assert-Test ((Invoke-Synthetic -System @($r))[0].'AV Qualification' -eq 'Evaluable registration observation') 'NotDetected registration is not evaluable.'
$r.ExpectedThirdPartyAVName=''
Assert-Test ((Invoke-Synthetic -System @($r))[0].'AV Qualification' -eq 'Expected product not configured') 'Missing expected product became evaluable.'
$r=New-SyntheticSystem; $r.FirmwareType='new-future-value'; $r.SecureBootStatus='Error'
$e=(Invoke-Synthetic -System @($r))[0]
Assert-Test ($e.'Firmware Type' -eq 'Unknown' -and $e.'Firmware Raw' -eq 'new-future-value' -and $e.'Secure Boot State' -eq 'Error') 'Future enum or error was silently coerced.'
$r=New-SyntheticSystem; $r.LastBootTime='2025-12-31 10:00:00'
Assert-Test ((Invoke-Synthetic -System @($r))[0].'Last Boot Quality' -eq 'Conflicting aliases') 'Conflicting boot aliases were hidden.'
foreach ($case in @(@('bad date','Invalid date'),@('2026-01-04T00:00:00Z','Future observation'),@('01/01/2026 00:00:00','Time zone unspecified'),@('2025-12-01T00:00:00Z','Stale'))) {
    $r=New-SyntheticSystem; $r.PlatformScriptLastUpdate=$case[0]
    Assert-Test ((Invoke-Synthetic -System @($r))[0].Freshness -eq $case[1]) 'Date/freshness qualifier changed.'
}
$r=New-SyntheticSystem; $r.PlatformScriptLastUpdate='02/01/2026 00:00:00'
Assert-Test ((Invoke-Synthetic -System @($r) -Zone 'UTC')[0].Freshness -eq 'Within policy') 'Explicit offsetless time-zone policy was ignored.'
Assert-Test ((Invoke-Synthetic -Age $null)[0].Freshness -eq 'Policy not configured') 'A fixed freshness threshold was silently invented.'
$unmatched=Invoke-Synthetic -System @((New-SyntheticSystem),(New-SyntheticSystem 'absent-id'))
Assert-Test ($unmatched.Count -eq 2 -and @($unmatched | Where-Object 'Inventory Match' -eq 'Unmatched').Count -eq 1 -and @($unmatched | Where-Object 'Prepared Inventory Match' -eq 'Unmatched').Count -eq 1) 'Unmatched device was dropped or matched by name.'
$ambiguous=(Invoke-Synthetic -Users @($user,$user))[0]
Assert-Test ($ambiguous.Country -eq 'Unknown' -and $ambiguous.'Data Quality Issues' -match 'Ambiguous primary-user') 'Ambiguous user match assigned a country.'
foreach ($action in @({Invoke-Synthetic -System @((New-SyntheticSystem),(New-SyntheticSystem))},{Invoke-Synthetic -Inventory @($inventory,$inventory)},{Invoke-Synthetic -Age 0},{ $r=New-SyntheticSystem; $r.TenantKey='other-tenant'; Invoke-Synthetic -System @($r) })) {
    $failed=$false
    try { & $action | Out-Null } catch {$failed=$true; Assert-Test ($_.Exception.Message -notmatch 'device-one|user@example') 'Failure leaked source identity.'}
    Assert-Test $failed 'Invalid identity or policy did not fail closed.'
}
Assert-Test (@(Invoke-Synthetic -System @()).Count -eq 0) 'Empty system population emitted invented devices.'
$devices=@('Enabled','Disabled','Error','Unknown','Not collected','NotSupported' | ForEach-Object {[pscustomobject]@{'Secure Boot State'=$_}})
$tokens=$null; $parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'New-SecurityEvidence.ps1'),[ref]$tokens,[ref]$parseErrors)
Assert-Test ($parseErrors.Count -eq 0) 'Security generator did not parse.'
foreach ($statement in $ast.EndBlock.Statements | Where-Object {$_.Extent.Text -match '^\$secureBoot(Covered|Healthy|Affected)\s*='}) {. ([scriptblock]::Create($statement.Extent.Text))}
Assert-Test ($secureBootCovered.Count -eq 2 -and $secureBootHealthy.Count -eq 1 -and $secureBootAffected.Count -eq 1) 'Secure Boot error became an adverse or evaluable entity.'
Assert-Test ($base[0].'Local Data Source' -eq 'LegacyPlatformScript' -and $base[0].'Last Console Logon Status' -eq 'NotCollected') 'Legacy provenance/console defaults changed.'
function New-RecurringSystem {
    $r=New-SyntheticSystem
    $r | Add-Member LocalDataSource 'IntuneRemediation'
    $r | Add-Member LocalDataFreshness 'Fresh'
    $r | Add-Member LocalObservationUtc '2026-01-02T00:00:00Z'
    $r | Add-Member LastConsoleLogonStatus 'Observed'
    $r | Add-Member LastConsoleLogonUser 'example\console-user'
    $r | Add-Member LastConsoleLogonSid 'S-1-5-21-1-2-3-1001'
    $r | Add-Member LastConsoleLogonUtc '2026-01-01T23:00:00Z'
    $r
}
$r=New-RecurringSystem; $r.PSObject.Properties.Remove('PlatformScriptLastUpdate')
$e=(Invoke-Synthetic -System @($r))[0]
Assert-Test ($e.'Observation UTC' -eq '2026-01-02T00:00:00.0000000Z' -and $e.'Local Data Freshness' -eq 'Fresh') 'Recurring timestamp without legacy alias failed.'
Assert-Test ($e.'Last Console Logon Status' -eq 'Observed' -and $e.'Last Console Logon User' -eq 'example\console-user' -and $e.'Last Console Logon UTC' -eq '2026-01-01T23:00:00.0000000Z') 'Qualified console evidence lost.'
Assert-Test ($e.'Primary User UPN' -eq 'user@example.invalid') 'Console identity overwrote primary-user ownership.'
Assert-Test ($e.'Last Console Logon SID' -eq 'S-1-5-21-1-2-3-1001') 'The prepared console SID was lost.'
$r=New-RecurringSystem; $r.LastConsoleLogonSid='not-a-sid'
Assert-Test ((Invoke-Synthetic -System @($r))[0].'Last Console Logon SID' -eq '') 'An invalid console SID was exposed.'
$r=New-RecurringSystem; $r.LastConsoleLogonStatus='Error'
Assert-Test ((Invoke-Synthetic -System @($r))[0].'Last Console Logon SID' -eq '') 'An error exposed an attributed console SID.'
foreach ($state in @('NotCollected','Stale','Error','Unknown','Pending','NotApplicable','Failed','ScriptError','InvalidOutput','InvalidTimestamp','FutureObservation','unexpected')) {
    $r=New-RecurringSystem; $r.LocalDataFreshness=$state
    $e=(Invoke-Synthetic -System @($r))[0]
    Assert-Test ($e.Freshness -ne 'Within policy' -and $e.'Last Console Logon User' -eq '' -and $e.'Last Console Logon Status' -ne 'Observed') 'Unqualified collection became current console evidence.'
    if ($state -ne 'unexpected') {
        Assert-Test ($e.'Local Data Freshness' -eq $state -and $e.Freshness -eq $state) 'The local collection state was lost in prepared evidence.'
    }
    Assert-Test ($e.'Last Console Logon SID' -eq '' -and $e.'Last Console Logon UTC' -eq '') 'An unavailable collection exposed console identity or time.'
}
$r=New-RecurringSystem; $r.LocalObservationUtc='2026-01-02T02:00:00+02:00'
Assert-Test ((Invoke-Synthetic -System @($r))[0].Freshness -eq 'Within policy') 'Equivalent observation aliases conflict.'
$r=New-RecurringSystem; $r.PlatformScriptLastUpdate='2026-01-02T01:00:00Z'
Assert-Test ((Invoke-Synthetic -System @($r))[0].'Observation UTC' -eq '') 'Contradictory timestamps qualified.'
foreach ($time in @('','invalid','2026-01-02T00:00:00','2026-01-04T00:00:00Z')) {
    $r=New-RecurringSystem; $r.LocalObservationUtc=$time; $r.PlatformScriptLastUpdate=''
    $e=(Invoke-Synthetic -System @($r) -Zone 'UTC')[0]
    Assert-Test ($e.Freshness -ne 'Within policy' -and $e.'Last Console Logon Status' -eq 'Unknown') 'Invalid recurring timestamp or future console observation qualified.'
}
foreach ($time in @('','invalid','2026-01-03T00:00:00Z','2026-01-01T23:00:00')) {
    $r=New-RecurringSystem; $r.LastConsoleLogonUtc=$time
    $e=(Invoke-Synthetic -System @($r))[0]
    Assert-Test ($e.'Last Console Logon Status' -eq 'Unknown' -and $e.'Last Console Logon User' -eq '') 'Unqualified console observation exposed as observed.'
}
$r=New-RecurringSystem; $r.LastConsoleLogonStatus='NotFoundInScan'; $r.LastConsoleLogonUser=''; $r.LastConsoleLogonUtc=''
Assert-Test ((Invoke-Synthetic -System @($r))[0].'Last Console Logon Status' -eq 'NotFoundInScan') 'NotFoundInScan became never-used or missing collection.'
$schemaScratch=Join-Path ([IO.Path]::GetTempPath()) ('DeviceSystemSchemaOffline-'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $schemaScratch)
$sourceSchema=Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\config\source-schema.json') -Raw | ConvertFrom-Json
$common=@($sourceSchema.files.'Intune_Devices_LocalSystem.csv'.columns)
$schemaCases=@(
    @{Extra=@('PlatformScriptLastUpdate');Accept=$true},
    @{Extra=@('LocalObservationUtc','LocalDataSource','LocalDataFreshness');Accept=$true},
    @{Extra=@('PlatformScriptLastUpdate','LocalObservationUtc','LocalDataSource','LocalDataFreshness');Accept=$true},
    @{Extra=@();Accept=$false},
    @{Extra=@('LocalObservationUtc');Accept=$false},
    @{Extra=@('PlatformScriptLastUpdate','LocalDataSource');Accept=$false},
    @{Extra=@('LocalObservationUtc','LocalDataFreshness');Accept=$false}
)
foreach($case in $schemaCases) {
    $schemaFile=Join-Path $schemaScratch ([guid]::NewGuid().ToString('N')+'.csv')
    $schemaRow=[ordered]@{}
    foreach($column in @($common)+@($case.Extra)) {$schemaRow[$column]=''}
    [pscustomobject]$schemaRow | Export-Csv -LiteralPath $schemaFile -NoTypeInformation
    $accepted=$true
    try {Assert-DeviceSystemCsvSchema -Path $schemaFile} catch {$accepted=$false}
    Assert-Test ($accepted -eq $case.Accept) 'Legacy/recurring header qualification changed.'
}
[pscustomobject]@{Passed=$true;Checks=$script:checks;Sources='Synthetic only';Collection=$false;Refresh=$false;Publication=$false} | ConvertTo-Json

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBIS9nPcPuqBOM5
# TDGtOOXKPbxOKpkM3BBPdeAYRve0AaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDxffGoJU6Up972895r1F3Nqg4/H42OduJs/+WWlcDc7MA0GCSqG
# SIb3DQEBAQUABIIBgFAA+GA2tqi9BG1U9Kt2u80pGTT7i1QowLMZLJXX550Kcwzz
# 2Wbqy6VyPk4qGtbOzcd5Ai1rEFVexGBM85gWG1s5g7H6rWx4vLvCcfS0/Iv4Az3J
# YIIUpOFy3XOfFGeIcNEG+DuJ/ihxaTRlDypb0V8R+OJopdcQRmpAHZqgazCiLAum
# OnPQFj5oOBALVIbvOA8VFgoIYBYIG69sp16B9zWHN8uca7XRhsnfNTap8xqxmelP
# cEtGY4ubdRQSweypnTehHJlUZVYHZeGEdvIJ909eda/LptMEJ8W6rq3uedG/gf9e
# kJqdV5LeFH1iEiwCVa/YIL7Q9xjZgaGAXOnSSZKUkfEWtwwYZKJlPT+DveMpHPkD
# MeUoSWK4x+NYjI3CsO2uzDyBSlctMOdFJ4ySOIAG44fOGGykPEsgKuMDaWJQ2twc
# NYL8aVaOG1PTYTvTVRlpfXcmtg9fZLmLDPpc3qLCz42fohmvkAqB0HlEOBkXB73Z
# p54tNF28gOe0CZWc0aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYyMzQx
# MDVaMC8GCSqGSIb3DQEJBDEiBCBa5cnIPdEBkcPHmAGjQaukFX4CsLuPCCIXq78N
# vgBzEzANBgkqhkiG9w0BAQEFAASCAgCitvGM/xbMrIxv9yCvb6U600G26wJdTPqj
# BYVPy77GhGD5oA3D+eKVOVruCDdvwRV0rM7tOtdxOhXKBEBW5vevJWfpOwo7MJ7G
# XfC9tErv4KKUuAjKZlV5/iEC9D0T3WSIRyWxggqdj3Poao3VkEEOF+q5HqVPuir8
# 9JQOUTcNLqOW8USIAdWO9Jr4kPJV2gVQjP0kVAY0MVykT7r1heMFS+VI+Fsw8z9b
# MLKfxTIVhzLVeUAC8IQKxs9HyrlGdr1cpucuLrDhCqRHcVABTN1qqAF0pibHXcRb
# uSJUXjS95eUOEBe9970zDff88s/YsPPosNVUTttoLkSP8wvZU1z1XpFzIbeAf9o1
# ql9/NNvibrzNP01UBt59X7RHQEWejvROpT8llWE4+47LrOHMjXKiRjtO7BBgBPn5
# 5IuVKjnOZovRfFRG5lzQrIHEYmIr5XTVZmhQjhNBkqFSUowHYETZLqwN+z2O1y9m
# 5HWZD4v7tC4SGaDRyYkKmq5d+iBYGs8mFag/e2hz2UbqjnLg1CdttBh+snD2Vcrq
# tW2J3G5qywHi8SIV5KFuHM/1JlAIKNEujD4demDOS/jWuswqKdKDtWs8XkgyJ3e9
# cF9w941qIru6IWtplqb74Y5wsV78AUBoGWtXiAq54SRxqoGSwffF5/9roSOXDx9J
# 8LhMijuI3w==
# SIG # End signature block
