<#
.SYNOPSIS
  Offline checks: the five usage reports read by the license report are qualified only by the UserActivity source receipt.
.VERSION
1.0.2
.NOTES
  Receipts are written by the real SmartM365.Core receipt commands in an isolated temporary folder.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$userActivityPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\Usage\SmartM365-M365UserActivity-Inventory.ps1')).Path
$coreModulePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1')).Path
Import-Module -Name $coreModulePath -MinimumVersion '1.0.90' -Force -ErrorAction Stop

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
$names = @(
  'Get-LicensesSuiteProduct', 'Get-LicensesOptionalProperty', 'ConvertTo-LicensesActivityDate', 'ConvertTo-LicensesMailboxSizeGb',
  'Get-LicensesQualifiedPartialAdCoverage', 'Get-LicensesCsvSource', 'Get-LicensesPublicationPath', 'Import-LicensesSourceCsv', 'Read-LicensesIndexedSource',
  'Get-LicensesMailboxGapSummary', 'Get-LicensesAdAccountActivitySummary', 'Get-LicensesFocusedUsageRows', 'Import-LicensesSourceCsvCached', 'Get-LicensesUserAdScopes', 'Get-LicensesUserScope', 'Get-LicensesScopeBreakdown', 'Get-LicensesOtherProduct'
)
$definitions = @($ast.FindAll({
  param($node)
  $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name
}, $true))
if ($definitions.Count -ne $names.Count) { throw 'License usage functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }
function WriteLog { param([string]$Message, [string]$Level) }

$script:Checks = 0
function Check {
  param([bool]$Condition, [string]$Label)
  if (-not $Condition) { throw "FAIL: $Label" }
  $script:Checks++
}

$usageFiles = @('M365_Users_Activity.csv','M365_Mailbox_Usage.csv','M365_Email_Activity.csv','M365_Apps_Usage_180D.csv','M365_OneDrive_Usage.csv')
$receiptName = 'SmartInventory_SmartM365-M365UserActivity-Inventory.current.json.txt'
$today = [datetime]::UtcNow.Date
$refresh = $today.AddDays(-3).ToString('yyyy-MM-dd')

$saved = @{}
foreach ($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')) {
  $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
  $saved[$name] = @{ Exists=($null -ne $variable); Value=$(if ($variable) { $variable.Value } else { $null }) }
}

function Set-TestTenant {
  param([string]$TenantKey)
  $global:SmartM365TenantKey = $TenantKey; $global:SmartM365OrganizationKey = 'contoso'
  $global:SmartM365EnvironmentKey = 'prod'; $global:SmartM365TenantId = 'synthetic-tenant'
}

function New-UsageFixture {
  param([string]$Name, [string]$TenantKey = 'prod', [string]$OneDrivePeriod = '180')
  $folder = Join-Path $script:Root $Name
  New-Item -ItemType Directory -Path $folder -Force | Out-Null
  $t = $TenantKey
  $upn = 'u1@example.invalid'
  @([pscustomobject]@{TenantKey='prod';UserId='u1';SkuPartNumber='SPE_E3'}) | Export-Csv -LiteralPath (Join-Path $folder 'M365_Licenses_Users.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;UserPrincipalName=$upn;ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$refresh;IsDeleted='False'}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_Users_Activity.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;'User Principal Name'=$upn;'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$refresh;'Is Deleted'='False'}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_Mailbox_Usage.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;'User Principal Name'=$upn;'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$refresh;'Is Deleted'='False';'Send Count'=1;'Read Count'=1}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_Email_Activity.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;'User Principal Name'=$upn;'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$refresh;Windows='No';Mac='No'}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_Apps_Usage_180D.csv') -NoTypeInformation -Encoding utf8NoBOM
  @([pscustomobject]@{TenantKey=$t;'Owner Principal Name'=$upn;'Is Deleted'='False';'Storage Used (Byte)'='1073741824';'Report Refresh Date'=$refresh;'Report Period'=$OneDrivePeriod}) |
    Export-Csv -LiteralPath (Join-Path $folder 'M365_OneDrive_Usage.csv') -NoTypeInformation -Encoding utf8NoBOM
  return $folder
}

function New-UserActivityReceipt {
  param([string]$Folder, [string[]]$Files, [string]$TenantKey = 'prod', [switch]$LeavePublishing)
  # Same producer, registry and scope as SmartM365-M365UserActivity-Inventory.ps1.
  Set-TestTenant $TenantKey
  $global:csvGeneratedPaths = $null
  Start-SmartM365CmdbSourceReceipt -ScriptPath $userActivityPath -SourceRootPath $Folder
  foreach ($file in $Files) { Register-SmartM365SourceCsv -Path (Join-Path $Folder $file) }
  if ($LeavePublishing) { return }
  Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope 'CMDB:activity'
  return (Complete-SmartM365CmdbSourceReceipt -Status 'Completed')
}

function Get-UsageSources {
  param([string]$Folder)
  # The license file is proven by a current-run publication, as during the licensing run.
  $licensePath = [IO.Path]::GetFullPath((Join-Path $Folder 'M365_Licenses_Users.csv'))
  $hashes = [hashtable]::new([StringComparer]::OrdinalIgnoreCase)
  $hashes[$licensePath] = (Get-FileHash -LiteralPath $licensePath -Algorithm SHA256).Hash
  $proof = [pscustomobject]@{ RunId='license-run'; Hashes=$hashes }
  $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $Folder -ExpectedTenantKey 'prod' -AsOfUtc $today -CurrentRunPublication $proof
  $byName = @{}
  foreach ($source in $usage.Sources) { if ($usageFiles -contains $source.Name) { $byName[$source.Name] = $source } }
  Check ($byName.Count -eq 5) 'The five usage reports are evaluated'
  return [pscustomobject]@{ Usage=$usage; Sources=$byName }
}

function Test-Unqualified {
  param($Source, [string]$Reason, [string]$Label)
  Check (-not $Source.Ready -and -not $Source.Forced -and $Source.Reason -ceq $Reason) "$Label ($($Source.Name)): expected '$Reason', got Ready=$($Source.Ready) '$($Source.Reason)'"
}

$script:Root = Join-Path $env:TEMP ('SmartM365-LicenseUsageReceipts-' + [guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Path $script:Root -Force | Out-Null

  # 1. Every report listed by a completed receipt with the same hash is qualified, never provisional.
  $folder = New-UsageFixture 'listed'
  Check ((New-UserActivityReceipt -Folder $folder -Files $usageFiles) -like '*.current.json.txt') 'Core completes the UserActivity receipt'
  $listed = Get-UsageSources $folder
  foreach ($name in $usageFiles) {
    $source = $listed.Sources[$name]
    Check ($source.Ready -and -not $source.Forced -and $source.Provenance -ceq 'collector completed' -and $source.Reason -eq '') "Listed report is qualified: $name"
    Check ($source.Date -eq $refresh) "Report date is its Report Refresh Date: $name"
  }
  Check ($listed.Usage.DowngradeReview.Unknown -eq 1) 'E3 review stays N/D without the other sources'

  # 2. A report present in DATA-LAST but absent from the receipt is not qualified.
  foreach ($name in @($usageFiles | Where-Object { $_ -ne 'M365_Users_Activity.csv' })) {
    $folder = New-UsageFixture ('absent-' + [IO.Path]::GetFileNameWithoutExtension($name))
    [void](New-UserActivityReceipt -Folder $folder -Files @($usageFiles | Where-Object { $_ -ne $name }))
    $result = Get-UsageSources $folder
    Test-Unqualified $result.Sources[$name] 'file missing from collector receipt' 'File absent from a Core receipt'
    foreach ($other in @($usageFiles | Where-Object { $_ -ne $name })) { Check $result.Sources[$other].Ready "Other listed report stays qualified: $other without $name" }
  }
  # M365_Users_Activity.csv is required by the producer registry: Core refuses a receipt without it,
  # so its absence is checked on a completed Core receipt whose record for that file is removed.
  $folder = New-UsageFixture 'absent-M365_Users_Activity'
  Check ($null -eq (New-UserActivityReceipt -Folder $folder -Files @($usageFiles | Where-Object { $_ -ne 'M365_Users_Activity.csv' })) -or
    -not (Test-Path -LiteralPath (Join-Path $folder $receiptName))) 'Core refuses a receipt without the required activity report'
  Test-Unqualified (Get-UsageSources $folder).Sources['M365_Users_Activity.csv'] 'collector receipt is missing' 'Refused Core receipt'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  $receiptPath = Join-Path $folder $receiptName
  $document = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
  $document.Files = @($document.Files | Where-Object File -ne 'M365_Users_Activity.csv')
  $document | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $receiptPath -Encoding utf8NoBOM
  Test-Unqualified (Get-UsageSources $folder).Sources['M365_Users_Activity.csv'] 'file missing from collector receipt' 'File absent from a Core receipt'

  # 3. A report whose bytes differ from the receipt is not qualified.
  $folder = New-UsageFixture 'hash'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  foreach ($name in $usageFiles) { Add-Content -LiteralPath (Join-Path $folder $name) -Value '' -NoNewline:$false }
  $hash = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $hash.Sources[$name] 'file hash differs from collector receipt' 'Changed file' }

  # 4. No receipt: fail-closed, the CSV alone is never used.
  $folder = New-UsageFixture 'missing'
  $missing = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $missing.Sources[$name] 'collector receipt is missing' 'Missing receipt' }

  # 5. A run that is publishing (or failed after publication) withholds the previous completed receipt.
  $folder = New-UsageFixture 'publishing'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  New-UserActivityReceipt -Folder $folder -Files @('M365_OneDrive_Usage.csv') -LeavePublishing
  try {
    $publishing = Get-UsageSources $folder
    foreach ($name in $usageFiles) { Test-Unqualified $publishing.Sources[$name] 'collector receipt is unreadable' 'Publishing run' }
  }
  finally { [void](Complete-SmartM365CmdbSourceReceipt -Status 'Failed') }
  $failed = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $failed.Sources[$name] 'collector receipt is unreadable' 'Run failed after publication' }

  # 6. A current receipt that is not Completed is not accepted.
  $folder = New-UsageFixture 'running'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  $receiptPath = Join-Path $folder $receiptName
  $document = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
  $document.Status = 'Running'
  $document | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $receiptPath -Encoding utf8NoBOM
  $running = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $running.Sources[$name] 'collector receipt is Running or partial' 'Running receipt' }

  # 7. Rows of another tenant, even under their own completed receipt, are not used.
  $folder = New-UsageFixture 'tenant' -TenantKey 'other'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles -TenantKey 'other')
  $tenant = Get-UsageSources $folder
  foreach ($name in $usageFiles) { Test-Unqualified $tenant.Sources[$name] 'another tenant is present' 'Foreign tenant' }

  # 8. OneDrive keeps the 180-day period of the other usage reports; another period is refused.
  $folder = New-UsageFixture 'period' -OneDrivePeriod '30'
  [void](New-UserActivityReceipt -Folder $folder -Files $usageFiles)
  $period = Get-UsageSources $folder
  Test-Unqualified $period.Sources['M365_OneDrive_Usage.csv'] 'unexpected report period' 'OneDrive period'
  Check $period.Sources['M365_Apps_Usage_180D.csv'].Ready 'Other reports are not affected by the OneDrive period'

  Write-Host ("PASS: license usage reports are qualified by the UserActivity receipt ({0} checks)." -f $script:Checks)
}
finally {
  foreach ($name in $saved.Keys) {
    if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
    else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
  }
  if (Test-Path -LiteralPath $script:Root) { Remove-Item -LiteralPath $script:Root -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCnFWYDWRwq1I3/
# KONd3KPwP1loo4sPH3dD8VT8DsXi7aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHNH/DieD+d3+l2e0uBwYl1XMjE/FyEjS1zB07NDeE/KMA0GCSqG
# SIb3DQEBAQUABIIBgAnxbNcJZ37fLNLwagnPvOEGiWTTcsUrvX9twZtbzbeEpekf
# 4UBizLynqEZ3BXvLXVNRtAmTtBCwGunROPhNUs2QqHupQn8/TVpvziPWex+1dtLt
# yhrxPZlSopbYfCvL+IV6koQcUcg6bSe3KIK7h3YXGIvUBR9Q2i6zOhtORqOwxl2N
# 9K+0ehEch+jOEh9JFQhhgzzkdw2EUK9AYudvGUhADRI4rsd71jzjqOXYD9xbZ51g
# EskMeUswMAItztIX2cr9+RYclSDP/pc/yYRfjpfomhBhQBN9Ou2JRpu40Jk/CtXu
# 0lEmXkYgVpy7c9tbU8ald/sqfEstJ0XmbNteVZ2UCdd+zz51GSY87k5jf9ZwXbqW
# EjM3/mQZu/VKSQLn39U6Fj6B5meEswrNEAcbaG5HPjqT7VzS+ffX6ADsXJt2GBdz
# T6j2XKiJgCpO3h10EwHgd4e0htXenrG4lzJtaY9YqjDL5Xlzl3AAH0My+0K+QE5N
# ZM+sF8pN1moryU/l8aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAxNDI2
# NDdaMC8GCSqGSIb3DQEJBDEiBCAsgpPQbWTzqBIDaYzDiUD2Ei/JauTbfUEnVHW6
# cdgcfTANBgkqhkiG9w0BAQEFAASCAgAH05SS0PPW++3WUhvGVaHT/mlPC4Sobu2W
# HcNszBm5Sj3HoCuy0J4Umlg21Z22/6LbQ41ixVFkfKTwo/URE2Jpr0f9V37o09fl
# VdaP9i+H0cHvF14misTJFgjACXwKw3aJ6On8sNXSeL+wD3+ICaVFAud0NfXuyA6f
# wInVSKPx0Heb4ZC1iqSjTjSp0cI9a2x2KjCjPLufrRge+E0cw8bsjMj1VNU8wnfP
# KWmAudQEQyc0PXqB7//khAj1k5TOTar6tvUI5OUImSsspNYkdJ+D0yDx/c8ZK8at
# X0TqVHhDNU3thxyBn+8IY5ZIjaLSpClMI8yPvyMyfGhdaXmGba9ZmDYAAWI5cSxC
# l1xxLE/+Nzrpb81W28N+nWRolMGbhFblX2RpxC8mgGCug9zqFy7KmeeeEsMPF0UA
# K9wT/hLZg8rESSQW2eOr7hNFL+4DY3GD31sTTLChgevJFfjXqvQKXtArrMkYwM9x
# Na3xsBG6XUI2N/TA93mutwlPMJBpPlTNKHmZSIqbDZuUDqfP7SIBrWaJGeXEc+vs
# LaM6NEaDhxThyx8rW2kq5eXERxidc8BkYDuCgBbTCNns5AwTR778Ohw+t83Ih9aS
# 2oQIbiRCwWCIAxRThlLD4DbrdWXucQXqbhTktnxmacNHtZ9mPlGLb/W+WK1BSsWd
# kdSSj03hrw==
# SIG # End signature block
