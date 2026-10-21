#Requires -Version 7.0
<#
.SYNOPSIS
  Offline checks for users deleted from the directory while the licensing run collects them.
.VERSION
1.0.3
.NOTES
  Synthetic identifiers only (GUIDs and example.invalid). Graph calls are replaced by mocks; no tenant access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Offline mocks keep the production command signatures.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Retry delays and warnings are silenced for the offline run.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Offline mocks change no state.')]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$null, [ref]$null)
$names = @('Invoke-GraphWithRetry', 'Test-LicensesGraphNotFound', 'Register-LicensesUserFailure', 'Assert-LicensesUserResolution')
$definitions = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name }, $true))
if ($definitions.Count -ne $names.Count) { throw 'Deleted-user functions are missing.' }
foreach ($definition in $definitions) { . ([scriptblock]::Create($definition.Extent.Text)) }
# The retry classifies Graph errors with the Core status helper, resolved through the module exports as in production.
Import-Module (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.98' -Force -WarningAction SilentlyContinue

$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$script:Log = [System.Collections.Generic.List[string]]::new()
function WriteLog { param([string]$Message, [string]$Level) $script:Log.Add("$Level $Message") }
function Start-Sleep { param([int]$Seconds) $null = $Seconds }
function Write-Warning { param([string]$Message) $null = $Message }
$script:LookupCalls = 0
$script:LookupBehavior = 'NotFound'
function Get-MgUser {
  param([string]$UserId, [string[]]$Property, $ErrorAction)
  $script:LookupCalls++
  switch ($script:LookupBehavior) {
    'NotFound' { throw [Exception]::new("Resource '$UserId' does not exist or one of its queried reference-property objects are not present. Status: 404 (NotFound) ErrorCode: Request_ResourceNotFound") }
    'Exists' { return [pscustomobject]@{ Id = $UserId } }
    'Throttled' { throw [Exception]::new('Status: 429 (TooManyRequests) ErrorCode: TooManyRequests') }
    'Forbidden' { throw [Exception]::new('Status: 403 (Forbidden) ErrorCode: Authorization_RequestDenied') }
  }
}
function Get-SyntheticErrorRecord([string]$Message, [Exception]$Inner) {
  try { if ($Inner) { throw [Exception]::new($Message, $Inner) } else { throw [Exception]::new($Message) } } catch { return $_ }
}
$notFound = Get-SyntheticErrorRecord "Resource 'synthetic' does not exist or one of its queried reference-property objects are not present. Status: 404 (NotFound) ErrorCode: Request_ResourceNotFound"
$forbidden = Get-SyntheticErrorRecord 'Status: 403 (Forbidden) ErrorCode: Authorization_RequestDenied'
function Get-SyntheticPathList { $paths = [System.Collections.Generic.List[object]]::new(); $paths.Add([pscustomobject]@{ UserId = 'kept' }); return ,$paths }
function Invoke-Case([string]$Behavior, $ErrorRecord, [bool]$DetailStage = $true) {
  $script:LookupBehavior = $Behavior; $script:LookupCalls = 0
  $paths = Get-SyntheticPathList
  $before = $paths.Count
  $paths.Add([pscustomobject]@{ UserId = '00000000-0000-0000-0000-000000000001' })
  $state = @{ UserResolutionFailures = 0; UsersDeletedDuringCollection = 0 }
  $result = Register-LicensesUserFailure -ErrorRecord $ErrorRecord -UserId '00000000-0000-0000-0000-000000000001' -LicenseDetailStage $DetailStage `
    -AssignmentPaths $paths -AssignmentPathCountBefore $before -State $state
  return [pscustomobject]@{ Result = $result; State = $state; Paths = $paths; Calls = $script:LookupCalls }
}

# Recognition of a Graph 404, also through an inner exception or a status property.
Check (Test-LicensesGraphNotFound -ErrorRecord $notFound) 'A Request_ResourceNotFound error was not recognized.'
Check (Test-LicensesGraphNotFound -ErrorRecord (Get-SyntheticErrorRecord 'Wrapper' ([Exception]::new('Status: 404 (NotFound)')))) 'An inner 404 was not recognized.'
$statusException = [Exception]::new('Graph error'); $statusException | Add-Member -NotePropertyName ResponseStatusCode -NotePropertyValue 404
Check (Test-LicensesGraphNotFound -ErrorRecord $statusException) 'A 404 status property was not recognized.'
Check (-not (Test-LicensesGraphNotFound -ErrorRecord $forbidden)) 'A 403 was taken for a 404.'

# 1. Confirmed deletion: excluded, counted once, its assignment paths removed, run allowed.
$case = Invoke-Case 'NotFound' $notFound
Check ($case.Result -eq 'Deleted' -and $case.State.UsersDeletedDuringCollection -eq 1 -and $case.State.UserResolutionFailures -eq 0) 'A confirmed deletion was not counted as deleted.'
Check ($case.Calls -eq 1) 'The deletion was not confirmed by a directory lookup.'
Check ($case.Paths.Count -eq 1 -and $case.Paths[0].UserId -eq 'kept') 'The deleted user assignment paths were not removed, or other paths were removed.'
$qualification = Assert-LicensesUserResolution -State $case.State -ListedUsers 47096
Check ($qualification -eq 'UsersDeletedDuringCollection=1') 'A confirmed deletion within the limit stopped the run.'
Check (@($script:Log | Where-Object { $_ -like 'WARNING*deleted from the directory during collection*' }).Count -ge 1) 'The deletion is not logged as a warning.'
Check (@($script:Log | Where-Object { $_ -match '@' }).Count -eq 0) 'A user principal name reached the deleted-user log.'

# 2. 404 not confirmed: the user still exists, the run fails.
$case = Invoke-Case 'Exists' $notFound
Check ($case.Result -eq 'Failure' -and $case.State.UserResolutionFailures -eq 1 -and $case.State.UsersDeletedDuringCollection -eq 0) 'An existing user with a 404 was excluded.'
$failed = $false; try { Assert-LicensesUserResolution -State $case.State -ListedUsers 47096 | Out-Null } catch { $failed = $_.Exception.Message -like '*could not be resolved*' }
Check $failed 'An unconfirmed 404 did not stop the run.'

# 3. The confirmation lookup itself fails (throttling, permission): not confirmed, the run fails.
foreach ($behavior in @('Throttled', 'Forbidden')) {
  $case = Invoke-Case $behavior $notFound
  Check ($case.Result -eq 'Failure' -and $case.State.UserResolutionFailures -eq 1) "A $behavior confirmation lookup excluded the user."
}

# 4. Any other error, or a 404 outside the license-detail call, fails without a deletion lookup.
$case = Invoke-Case 'NotFound' $forbidden
Check ($case.Result -eq 'Failure' -and $case.Calls -eq 0 -and $case.State.UserResolutionFailures -eq 1) 'A permission error was treated as a deletion.'
$case = Invoke-Case 'NotFound' $notFound $false
Check ($case.Result -eq 'Failure' -and $case.Calls -eq 0) 'A 404 outside the license-detail call was treated as a deletion.'
$case = Invoke-Case 'NotFound' (Get-SyntheticErrorRecord 'Status: 503 (ServiceUnavailable)')
Check ($case.Result -eq 'Failure' -and $case.Calls -eq 0) 'A service error was treated as a deletion.'

# 5. Mass deletion limit: 1 % of listed users, capped at 100 by default; both are configurable.
foreach ($limitCase in @(
    @{ Deleted = 100; Listed = 47096; Fails = $false }, @{ Deleted = 101; Listed = 47096; Fails = $true },
    @{ Deleted = 5; Listed = 500; Fails = $false }, @{ Deleted = 6; Listed = 500; Fails = $true },
    @{ Deleted = 11; Listed = 47096; MaxCount = 10; Fails = $true }, @{ Deleted = 30; Listed = 1000; MaxPercent = 5; Fails = $false })) {
  $state = @{ UserResolutionFailures = 0; UsersDeletedDuringCollection = $limitCase.Deleted }
  $arguments = @{ State = $state; ListedUsers = $limitCase.Listed }
  if ($limitCase.ContainsKey('MaxCount')) { $arguments.MaxDeletedCount = $limitCase.MaxCount }
  if ($limitCase.ContainsKey('MaxPercent')) { $arguments.MaxDeletedPercent = $limitCase.MaxPercent }
  $threw = $false; try { Assert-LicensesUserResolution @arguments | Out-Null } catch { $threw = $_.Exception.Message -like '*above the limit*' }
  Check ($threw -eq $limitCase.Fails) ("Deleted-user limit misapplied: {0} deleted of {1}." -f $limitCase.Deleted, $limitCase.Listed)
}
Check ($null -eq (Assert-LicensesUserResolution -State @{ UserResolutionFailures = 0; UsersDeletedDuringCollection = 0 } -ListedUsers 10)) 'A run without deletion returned a qualification.'

# 6. Collector wiring: stage flag around the license details, classification in the catch, limit before export,
#    qualification in the receipt, counter in the completion summary, one version.
$text = [IO.File]::ReadAllText($inventoryPath)
# The snapshot re-read (Licences 1.58) stays inside the deletion stage: a user deleted meanwhile is excluded.
Check ($text -match '(?s)\$licenseDetailStage = \$true\s+\$licenseDetails = Invoke-GraphWithRetry \{ Get-MgUserLicenseDetail[^\r\n]*.{0,600}?Resolve-LicensesUserSnapshot.{0,800}?\$licenseDetailStage = \$false') 'The license-detail call is not marked as the deletion stage.'
Check ($text -match 'Register-LicensesUserFailure -ErrorRecord \$_ ') 'The user loop does not classify its failures.'
Check ($text -match '\$usersDeletedQualification = Assert-LicensesUserResolution') 'The run does not apply the deleted-user rules.'
Check ($text -match "-Qualifications @\(\`$usersDeletedQualification") 'The deleted-user count does not reach the receipt.'
Check ($text -match 'Add-SmartM365SourceRunNote -Message \(''\{0\}: users deleted') 'The deleted-user count does not reach the run state.'
Check ($text -match '-UsersDeletedDuringCollection \(\[int\]\$userResolution\.UsersDeletedDuringCollection\)') 'The completion summary does not show the deleted-user count.'
$version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
Check ($version -and [version]$version -ge [version]'1.55' -and $text -match ('\$ScriptVersion = "' + [regex]::Escape($version) + '"')) 'Header and runtime versions differ or predate 1.55.'

"PASS: License deleted-user offline checks ($script:checks checks)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUJdvEwbRBNMc/
# J7mxaq/k+ffkkIlW7k6e7HvT/KeenaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDa6cN1/TgEu6R+CyQfojQR1YY5ZN4uJrJqphjJqSODoMA0GCSqG
# SIb3DQEBAQUABIIBgEr7W4sfDeOHrhK38870UMUQ7H+INtXjgSETNpt3t4DZeecC
# PvHX1TxfmyIrq4jS2czww3aLNlwuf9rPrUozb4OSJgQXMKC8QUPVIWBM5CCz2QRE
# ewsSpM96gDxgUtGjYzOBuB81nOv5bg4Zu5zi3jdcPmqstQOnBS3TXm+lcG0myJMp
# eTtOr7QGdFUk9oYcVGYDaB2cDZsmzeAt/pFWg1yJqYC32f5121vSUb/CRWRoFsx7
# DzOtTuA/sB1OQjydyeudC6TFn1pCoqQZz5j3gJgsyD9r8nmWOMAxRDYc0BgtjDiF
# uo1e5c1lczyBbOHXQI7xFdO7KRAEMcjdI4mx5f2fdO6R6s57WhETmugyE92UGJNF
# K3NStNjkzXLODvLwS4Tewk4ZUFxFQkLXFH9Cb1tr1iUk2OVx938MtKwCTeNbrn+Z
# fuG7BOYKHgrIxAkX4YmE8TBOKOP9vW5p2t3LeZqHwzLxG9mO3cqwlWk1QjTbLw4+
# yrp/TnevoPvFLVnmaKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExNDQ5
# MzBaMC8GCSqGSIb3DQEJBDEiBCCw4ZUuZ4bN0kxopXi9sp2dEWD46pxdiuVtcQDt
# VpZnaTANBgkqhkiG9w0BAQEFAASCAgBlTXLtDx35mAhWRVQ9WN4DjwBgdb/QUh3M
# O8ELgaykR5FNGpeBAQ3EkjRIxlRSve+DeWDkZbnwEwzc3uhQinYHifyRzAfrcnQ3
# 6Hril1Pi+Waonzy2Ll8kkskZSgbF9t/t/BnrTm7oj2PjMjQXjcSt0leMykZDgBFD
# KDMm9+RwTiw98kNyQyvh+yJdjxUcOihJW9uI6LGiO5b9mBeL3iQ8IDahYBO7tgrk
# 5BuSXeF2YDSd6StwBNxUkL8kMUjvxzPIEmk8QWCs8o334aHf0rSrh5nzu693/Bqp
# Qi+QQLZHhmUupPiL10X3R/ApY4iA+HR5UcT3KBfnmaANlBWHF6Mu6sqShe5D+V9E
# c673LsigLpOI3sXR8K6DxcNEGKykS8lCvDh3c2v1ouB6BFY0NBp7jnO5m3ZOh1SQ
# QK2ETiuylpn99ACPzo9CfOkniUHyqwFxGkeiE6TVPug5Wr1o0cek5Nq7gq0xixet
# uwY3UPPMIQ+dT433L4S3+Dm+aDzhbe8Ln0Y2CmBt4D/A65Rt6Jczy5U8gBt+fuUY
# ogdIpse3IvQ+lfKY3fzwYIBf+PCza1em+bXnjRq0bc0pt/Eu9px0SvuZov8pNEXw
# DebRc75mGjJdgYKR8XN8evOfcxS2p+OwOZqSqnTZF42C4Kl5X6XWy5Dy4bsz8zuP
# jYva5FpKRQ==
# SIG # End signature block
