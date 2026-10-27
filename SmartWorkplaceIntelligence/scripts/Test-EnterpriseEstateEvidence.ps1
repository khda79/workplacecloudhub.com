[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'EnterpriseEstateEvidence.psm1') -Force
$scope=@{TenantKey='fixture';OrganizationKey='fixture-org';EnvironmentKey='test';TenantId='11111111-1111-1111-1111-111111111111'}
function Row($Values){$h=[ordered]@{};foreach($k in $scope.Keys){$h[$k]=$scope[$k]};foreach($k in $Values.Keys){$h[$k]=$Values[$k]};[pscustomobject]$h}
function Assert($Value,[string]$Name){if(-not $Value){throw "FAILED: $Name"};Write-Host "PASS: $Name"}
function Fails([string]$Name,[scriptblock]$Body){$failed=$false;try{$null=& $Body}catch{$failed=$true};Assert $failed $Name}
$au='10000000-0000-0000-0000-000000000001';$su='10000000-0000-0000-0000-000000000002'
$pc='20000000-0000-0000-0000-000000000001';$server='20000000-0000-0000-0000-000000000002'
$cu='30000000-0000-0000-0000-000000000001';$guest='30000000-0000-0000-0000-000000000002'
$mi='40000000-0000-0000-0000-000000000001';$mobile='40000000-0000-0000-0000-000000000002'
$ad=@(Row @{ObjectGUID=$au;ObjectSID='S-1-5-21-100-200-300-1001';ImmutableId_AD='CaseSensitiveValue';UserPrincipalName='old@example.test';Enabled='false';Country='DE'})
$cloud=@((Row @{'Object Id'=$cu;OnPremisesSecurityIdentifier=$ad[0].ObjectSID;OnPremisesImmutableId='CaseSensitiveValue';'User principal name'='new@example.test';AccountEnabled='false';UserType='Member';CountryOrRegion='Germany';'Usage location'='de'}),(Row @{'Object Id'=$guest;OnPremisesSecurityIdentifier='';OnPremisesImmutableId='';'User principal name'='guest@example.test';AccountEnabled='true';UserType='Guest';'Usage location'='US'}))
$directory=@((Row @{ObjectGUID=$au;ObjectSID=$ad[0].ObjectSID;ObjectClass='user'}),(Row @{ObjectGUID=$su;ObjectSID='S-1-5-21-100-200-300-1002';ObjectClass='msDS-GroupManagedServiceAccount'}),(Row @{ObjectGUID=$pc;ObjectSID='S-1-5-21-100-200-300-1003';ObjectClass='computer'}),(Row @{ObjectGUID=$server;ObjectSID='S-1-5-21-100-200-300-1004';ObjectClass='computer'}))
$managed=@((Row @{ManagedDeviceId=$mi;AzureADDeviceId=$pc;OperatingSystem='Windows';UserId=$cu}),(Row @{ManagedDeviceId=$mobile;AzureADDeviceId='50000000-0000-0000-0000-000000000001';OperatingSystem='Android';UserId='' }))
$computers=@(Row @{ObjectGUID=$pc})
function Project { New-EnterpriseEstateProjection -ADUsers $ad -EntraUsers $cloud -DirectoryObjects $directory -ManagedDevices $managed -ADComputers $computers -ADObservationUtc '2026-01-01T00:00:00Z' -ADUserObservationUtc '2026-01-01T02:00:00Z' -EntraObservationUtc '2026-01-02T00:00:00Z' -IntuneObservationUtc '2026-01-01T01:00:00Z' }
$p=Project
Assert ($p.Accounts.Count -eq 3 -and $p.Devices.Count -eq 3) 'Deduplicated unions include service accounts, guests and AD-only computers'
Assert ($p.Audit.AccountStrongLinks -eq 1 -and $p.Audit.DeviceStrongLinks -eq 1) 'Only strong one-to-one identity links merge'
Assert ($p.Accounts[0].'AD State' -eq 'Disabled') 'Disabled accounts are retained'
Assert ($p.Accounts[0].'AD Observation UTC' -eq '2026-01-01T02:00:00Z' -and $p.Accounts[1].'AD Observation UTC' -eq '2026-01-01T00:00:00Z' -and $p.Devices[0].'AD Observation UTC' -eq '2026-01-01T00:00:00Z') 'User and directory publication times remain distinct'
Assert ($p.Accounts[0].Country -eq 'Germany' -and $p.Devices[0].Country -eq 'Germany') 'Country follows observed user evidence'
$cloud[0].CountryOrRegion='Germany'
$cloud[0].'Usage location'='FR'
$countryPreferred=Project
Assert ($countryPreferred.Accounts[0].Country -eq 'Germany' -and $countryPreferred.Devices[0].Country -eq 'Germany') 'CountryOrRegion takes priority over licensing usage location'
Assert ($countryPreferred.Accounts[0].'Country Basis' -eq 'Entra country' -and $countryPreferred.Devices[0].'Country Basis' -eq 'Primary-user proxy: Entra country') 'Country basis records directory-country attribution'
$cloud[0].CountryOrRegion='Unknown';$cloud[0].'Usage location'='de'
Assert ((Project).Devices[0].Country -eq 'Unknown' -and (Project).Accounts[0].Country -eq 'Unknown') 'Unknown Entra country never falls back to licensing or AD'
$cloud[0].CountryOrRegion=''
Assert ((Project).Devices[0].Country -eq 'Unknown') 'Blank Entra country never falls back to usage location'
$cloud[0].CountryOrRegion='Germany'
Assert ($p.Devices[1].Country -eq 'Unknown') 'AD-only country is not inferred'
Import-Module (Join-Path $PSScriptRoot 'AdCountryEvidence.psm1') -Force
$rules=New-AdCountryRules -Configuration @{TenantKey='fixture';AdDomainCountryMappings=@{'north.example.test'='DE'};AdComputerNameCountryPrefixes=@{DE='Germany';FR='France'}} -ExpectedTenantKey 'fixture'
$directory[3]|Add-Member DomainName 'north.example.test'
$directory[3]|Add-Member Name 'DE-SERVER'
$mapped=New-EnterpriseEstateProjection -ADUsers $ad -EntraUsers $cloud -DirectoryObjects $directory -ManagedDevices $managed -ADComputers $computers -ADObservationUtc '2026-01-01T00:00:00Z' -ADUserObservationUtc '2026-01-01T02:00:00Z' -EntraObservationUtc '2026-01-02T00:00:00Z' -IntuneObservationUtc '2026-01-01T01:00:00Z' -AdCountryRules $rules
Assert ($mapped.Devices[1].Country -eq 'Germany' -and $mapped.Devices[1].'Country Basis' -like 'Configured AD domain*') 'Configured AD-only country with explicit provenance'
Assert (($mapped.Accounts|ConvertTo-Json -Compress) -ceq ($p.Accounts|ConvertTo-Json -Compress)) 'AD computer rules do not change accounts'
Assert ($mapped.Devices.Count -eq $p.Devices.Count -and $mapped.Audit.DeviceStrongLinks -eq $p.Audit.DeviceStrongLinks) 'Country rules preserve device population and joins'
$directory[3].PSObject.Properties.Remove('DomainName');$directory[3].PSObject.Properties.Remove('Name')
Assert ($p.Devices[2].Platform -eq 'Android') 'All Intune platforms retained'
Assert ($p.Accounts[0].'UPN Comparison' -eq 'Different; ID match retained') 'UPN change does not override identity keys'
Assert ($p.Audit.NameBasedMerges -eq 0 -and $p.Audit.UPNBasedMerges -eq 0) 'No name-based merges'
$cloud[1].OnPremisesImmutableId='CaseSensitiveValue';Fails 'Duplicate secondary key fails closed' {Project};$cloud[1].OnPremisesImmutableId=''
$cloud[0].OnPremisesImmutableId='';$cloud[1].OnPremisesImmutableId='CaseSensitiveValue';Fails 'Conflicting SID and immutable ID fail closed' {Project};$cloud[0].OnPremisesImmutableId='CaseSensitiveValue';$cloud[1].OnPremisesImmutableId=''
$cloud[0].OnPremisesSecurityIdentifier='';$cloud[0].OnPremisesImmutableId='casesensitivevalue';$p=Project;Assert ($p.Audit.AccountStrongLinks -eq 0) 'Base64 immutable IDs are case-sensitive';$cloud[0].OnPremisesSecurityIdentifier=$ad[0].ObjectSID;$cloud[0].OnPremisesImmutableId='CaseSensitiveValue'
$cloud[1].TenantKey='other';Fails 'Cross-tenant evidence rejected' {Project};$cloud[1].TenantKey='fixture'
$managed[1].AzureADDeviceId=$pc;Fails 'Duplicate device bridge rejected' {Project};$managed[1].AzureADDeviceId='50000000-0000-0000-0000-000000000001'
$managed[1].ManagedDeviceId=$mi;Fails 'Duplicate managed-device ID rejected' {Project};$managed[1].ManagedDeviceId=$mobile
$directory[0].ObjectClass='contact';Fails 'Directory and user coverage mismatch rejected' {Project};$directory[0].ObjectClass='user'
$managed[1].AzureADDeviceId='00000000-0000-0000-0000-000000000000';Assert ((Project).Devices.Count -eq 3) 'Unregistered Intune devices retained without inventing a bridge'
$contract=Get-Content -LiteralPath (Join-Path (Split-Path $PSScriptRoot -Parent) 'config/prepared-evidence-contract.json.txt') -Raw|ConvertFrom-Json
Assert (@($contract.tables|Where-Object {$_.file -in @('AllAccountsEvidence.csv','AllDevicesEvidence.csv')}).Count -eq 2) 'Both source schemas declared'
# Exercise the complete file wrapper using only independently invented fixtures.
$fixtureRoot=Join-Path ([IO.Path]::GetTempPath()) ('EstateFixture-'+[guid]::NewGuid().ToString('N'))
$last=Join-Path $fixtureRoot 'DATA-LAST';$out=Join-Path ([IO.Path]::GetTempPath()) ('EstateStaging-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $last -Force|Out-Null
$managed[1].AzureADDeviceId='50000000-0000-0000-0000-000000000001'
foreach($c in $cloud){$c|Add-Member CollectedAtUtc '2026-01-01T00:00:00Z'}
foreach($m in $managed){$m|Add-Member CollectedAtUtc '2026-01-01T00:00:00Z'}
foreach($d in $directory){$d|Add-Member DomainName 'fixture.test'}
$ad|Export-Csv -LiteralPath (Join-Path $last 'AD_Users_AllDomains.csv') -NoTypeInformation
$cloud|Export-Csv -LiteralPath (Join-Path $last 'M365_Users_Active.csv') -NoTypeInformation
$directory|Export-Csv -LiteralPath (Join-Path $last 'AD_DirectoryObjects_AllDomains.csv') -NoTypeInformation
$managed|Export-Csv -LiteralPath (Join-Path $last 'Intune_ManagedDevices_All.csv') -NoTypeInformation
$computers|Export-Csv -LiteralPath (Join-Path $last 'AD_Computers_AllDomains.csv') -NoTypeInformation
$null=& (Join-Path $PSScriptRoot 'New-EnterpriseEstateEvidence.ps1') -DataRoot $fixtureRoot -OutputRoot $out
$stageAccounts=@(Import-Csv -LiteralPath (Join-Path $out 'AllAccountsEvidence.csv'))
Assert ($stageAccounts.Count -eq 3 -and -not $stageAccounts[0].PSObject.Properties['TenantKey']) 'Staging is compatible with publisher-owned TenantKey prefix'
Assert ($stageAccounts[0].PSObject.Properties['TenantId']) 'Staging retains source scope metadata'
$audit=Get-Content -LiteralPath (Join-Path $out 'EnterpriseEstateAudit.json.txt') -Raw|ConvertFrom-Json
Assert ($audit.Sources.Count -eq 5 -and -not $audit.Published -and -not $audit.Refreshed) 'Private audit records input hashes without activating data'
$rulesPath=Join-Path $fixtureRoot 'AdCountryRules.json.txt'
@{TenantKey='fixture';AdDomainCountryMappings=@{'fixture.test'='DE'};AdComputerNameCountryPrefixes=@{}}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $rulesPath
$mappedOut=Join-Path ([IO.Path]::GetTempPath()) ('EstateCountryStaging-'+[guid]::NewGuid().ToString('N'))
$null=& (Join-Path $PSScriptRoot 'New-EnterpriseEstateEvidence.ps1') -DataRoot $fixtureRoot -OutputRoot $mappedOut -AdCountryRulesPath $rulesPath
$mappedDevices=@(Import-Csv -LiteralPath (Join-Path $mappedOut 'AllDevicesEvidence.csv'))
Assert ($mappedDevices[1].Country -eq 'Germany' -and $mappedDevices[1].'Country Basis' -like 'Configured AD domain*') 'File wrapper consumes explicit private country-rule snapshot'
$mappedAudit=Get-Content -LiteralPath (Join-Path $mappedOut 'EnterpriseEstateAudit.json.txt') -Raw|ConvertFrom-Json
Assert ($mappedAudit.AdCountryRulesSHA256 -ceq (Get-FileHash -LiteralPath $rulesPath).Hash) 'Audit records exact country-rule snapshot hash'
Fails 'Existing staging is not overwritten' {& (Join-Path $PSScriptRoot 'New-EnterpriseEstateEvidence.ps1') -DataRoot $fixtureRoot -OutputRoot $out}
Fails 'Raw source tree is protected' {& (Join-Path $PSScriptRoot 'New-EnterpriseEstateEvidence.ps1') -DataRoot $fixtureRoot -OutputRoot (Join-Path $fixtureRoot 'new-staging')}
Fails 'PBIP output refused' {& (Join-Path $PSScriptRoot 'New-EnterpriseEstateEvidence.ps1') -DataRoot $fixtureRoot -OutputRoot (Join-Path $out 'pbip')}
# Intentionally retain these small data-free fixtures for inspection; no recursive cleanup.
Write-Host 'Enterprise estate offline tests passed. No collector, network, refresh or publication invoked.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAR7CDiaKQWUZuK
# 1SKKDAXt0+VEY09ka4OBB3P7Jz5iI6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKHKWKQCRn/JqUCu3eKDBlfKNGXW8EaqSn/SGx/9OlL3MA0GCSqG
# SIb3DQEBAQUABIIBgElpdmdmrNUSPjTvyu9oDydiSqYVVJqcrnmzSWwhBDEiUT0E
# jefT5EpotBgFOtNOBMuqzvoxXmkLMeO6YbfMF7RLyGKlz+1Ftp7y7mKzuM6Qr8OG
# oJ6dUayI+ksPP6ynMO7rw9VJSs+GEmKxa1lGaH1wPpKGX+1vMus7vrVj6D8R+h9w
# kfHifVjHMfDD3cFlqqrL3egS+2Li4KGHZZXesUjA2Ladn36PRcyUrjKaEEbRXA0W
# /gkcjHUQA0RzaubWmB+QtKZ1asNx0zBjNdu46N6ZW7mfoK8ZqUvT/3cYEivAMQ80
# Y/22BDo4zaDTEARbc1rgMauNCghHBXg6iVMbByiGSxITMhkE4n6BPq98hq7d2G1Q
# sozgYPY+mYb5Iw9eO89ccztVcI7nUVvXMsHNcVItqMsDSCG9l1K57gekUWNCq/1c
# DvmNrjH3bbK32iHmuzBwasVJd+fZ5/SpOTlg/EYpoKsiw5bjPMrMgqOM/yUnSxvf
# dZ0YRUIH3M0ZOmVCzaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYyMzQx
# MDVaMC8GCSqGSIb3DQEJBDEiBCADijvUEcnrBep0sYZXFnieXMJyIjkJyfH27Nhr
# JUzh9DANBgkqhkiG9w0BAQEFAASCAgBqHKbkUiOQOwuMAeEFFKYYxymZCHXoHeRv
# Otfnuzn2+5v7WAargG2SVypzB8k7SNysKPvsFP1bj4A7nE3U7IBPEzuRHpF0/+3K
# yUePWwnS3QvCay6bMu5+hfVeeFvN03xX+h+bhGxNlXNuiQKgaGJTvymkMx0lNX+F
# CemlCLMiprjoqb70HPvr488YaUgNtelsrbO2zfcMiF8aqYC81rHlqnLAx2XuJzte
# RwYrt+qfTmvY9ts0CLcIgDIc9lYMb6gJe5vsTOq8x1qAm6Ayje7Wm5GGlj82ZieN
# YlfYMijHx1Rsq4fZ5w3T9hMcT2IhPHw1cB9uvRqAjkFtn/ialGT9Mt3aTHbnxdIK
# Yv13fcxyjUKIfqHTFf11G20+bJ/QLfewEuG+KApTyKVjkS1drfZ91CLVb/1ENRoA
# +n7ckcEmswx/+HTYb0iiMbQ8u6XzazZj6Xcc9G+KvViiLCyccx5WjC0l0/BFRtvO
# +pYsem46iU7F8qAsywFGm6rrAj36/OGf4jXCuA6Pjlx9UJ1OuZ6Z5FRRncZdlSKf
# ICbz56SSH/OTE3V42pP79T6L5GERi56slIM16fO0tUF9YeI/7E4LYd4sYKfV2q0Y
# bcMPgwOzjEa+bOpuhME+o8rwAiOUzw3NYNVzZJEhyQt8gF4Ql+uQfZMSa4qwFKJ1
# DA7A6E3l4A==
# SIG # End signature block
