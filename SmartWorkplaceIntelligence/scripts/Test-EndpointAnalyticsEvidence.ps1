#Requires -Version 7.0
<#
.SYNOPSIS
Offline current and weekly Endpoint Analytics ambiguity refusal tests.
.VERSION
1.0.0
.NOTES
Uses synthetic CSVs only; no connected data, Power BI, collection or publication.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='New-TestRow constructs synthetic in-memory rows only.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns','',Justification='Get-TestDefinitions returns a set of AST function definitions for isolated tests.')]
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'EndpointAnalyticsEvidence.psm1') -Force -ErrorAction Stop
$script:Checks = 0
function Assert-Test {
    param([bool]$Condition,[string]$Message)
    $script:Checks++
    if (-not $Condition) { throw $Message }
}
function Get-TestFailure {
    param([scriptblock]$Action)
    try { & $Action | Out-Null } catch { return $_ }
    throw 'Expected synthetic conflict refusal was not raised.'
}
function Get-TestDefinitions {
    param([string]$Path,[string[]]$Names)
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw 'Production generator did not parse.' }
    foreach ($name in $Names) {
        $node=$ast.Find({param($item) $item -is [Management.Automation.Language.FunctionDefinitionAst] -and $item.Name -eq $name},$true)
        if (-not $node) { throw "Production function is missing: $name" }
        $node.Extent.Text
    }
}
function New-TestRow {
    param([string]$Device='synthetic-secret-device',[string]$Report='EADeviceScoresV2',[string]$Score='71',[string]$Date='2026-01-02T00:00:00Z')
    [pscustomobject]@{TenantKey='synthetic';ReportName=$Report;DeviceId=$Device;EndpointAnalyticsScore=$Score;ReportRefreshDate=$Date;StartupScore='80';AppReliabilityScore='90';StopErrorCount='0';CoreBootTime='20';CoreSignInTime='10'}
}
$first=New-TestRow
$second=New-TestRow -Score '88'
foreach ($rows in @(
    [pscustomobject]@{Rows=@($first,$second)},
    [pscustomobject]@{Rows=@($second,$first)},
    [pscustomobject]@{Rows=@($first,$first)},
    [pscustomobject]@{Rows=@($first,(New-TestRow -Score '88' -Date '2026-02-01T00:00:00Z'))},
    [pscustomobject]@{Rows=@($first,(New-TestRow -Device ' SYNTHETIC-SECRET-DEVICE '))},
    [pscustomobject]@{Rows=@(New-TestRow -Device '')},
    [pscustomobject]@{Rows=@(New-TestRow -Report '')},
    [pscustomobject]@{Rows=@([pscustomobject]@{DeviceId='synthetic';ReportName='EADeviceScoresV2'})},
    [pscustomobject]@{Rows=@($null)}
)) {
    $failure=Get-TestFailure { Assert-EndpointAnalyticsDeviceGrain -Rows $rows.Rows }
    Assert-Test ($failure.Exception.Message -match 'evidence rejected') 'Invalid or ambiguous identity was not rejected by the guard.'
    Assert-Test ($failure.Exception.Message -notmatch 'synthetic-secret-device|\b71\b|\b88\b') 'Guard exposed device or score values.'
}
$otherTenant=New-TestRow -Device 'synthetic-two'
$otherTenant.TenantKey='synthetic-other'
$failure=Get-TestFailure { Assert-EndpointAnalyticsDeviceGrain -Rows @($first,$otherTenant) }
Assert-Test ($failure.Exception.Message -match 'tenant populations=2') 'Mixed tenant populations were silently merged.'
Assert-EndpointAnalyticsDeviceGrain -Rows @()
$appOnly=New-TestRow -Report 'EADevicePerformanceV2' -Score ''
Assert-EndpointAnalyticsDeviceGrain -Rows @($first,$appOnly,(New-TestRow -Device 'synthetic-two'))
Assert-Test ($true) 'Empty evidence and distinct reports/devices are accepted.'

$currentPath=Join-Path $PSScriptRoot 'New-DeviceInventoryEvidence.ps1'
$trendPath=Join-Path $PSScriptRoot 'New-DeviceLifecycleExperienceTrendEvidence.ps1'
$definitions=@(Get-TestDefinitions $currentPath @('Get-NormalizedKey','Convert-ToDateTimeOrNull','Get-PreferredLatestIndex')) + @(Get-TestDefinitions $trendPath @('Get-LatestRowsByDevice'))
$module=New-Module -Name SyntheticEndpointSelectors -ScriptBlock {
    param($Definitions,$HelperPath)
    Import-Module $HelperPath -Force -ErrorAction Stop
    foreach ($definition in $Definitions) { . ([scriptblock]::Create($definition)) }
} -ArgumentList $definitions,(Join-Path $PSScriptRoot 'EndpointAnalyticsEvidence.psm1')
try {
    foreach ($rows in @([pscustomobject]@{Rows=@($first,$second)},[pscustomobject]@{Rows=@($second,$first)})) {
        $failure=Get-TestFailure { & $module { param($Rows) Get-PreferredLatestIndex -Rows $Rows -KeyProperty DeviceId -PreferredProperty EndpointAnalyticsScore -DateProperty ReportRefreshDate } $rows.Rows }
        Assert-Test ($failure.Exception.Message -match 'evidence rejected') 'Current selector chose an ambiguous score.'
        $failure=Get-TestFailure { & $module { param($Rows) Get-LatestRowsByDevice -Rows $Rows -PreferredProperty EndpointAnalyticsScore } $rows.Rows }
        Assert-Test ($failure.Exception.Message -match 'evidence rejected') 'Trend selector chose an ambiguous score.'
    }
    foreach ($rows in @([pscustomobject]@{Rows=@($first,$appOnly)},[pscustomobject]@{Rows=@($appOnly,$first)})) {
        $selected=& $module { param($Rows) Get-PreferredLatestIndex -Rows $Rows -KeyProperty DeviceId -PreferredProperty EndpointAnalyticsScore -DateProperty ReportRefreshDate } $rows.Rows
        Assert-Test ($selected['synthetic-secret-device'].EndpointAnalyticsScore -eq '71') 'Valid score-bearing report preference changed.'
        $selected=@(& $module { param($Rows) Get-LatestRowsByDevice -Rows $Rows -PreferredProperty EndpointAnalyticsScore } $rows.Rows)
        Assert-Test ($selected.Count -eq 1 -and $selected[0].EndpointAnalyticsScore -eq '71') 'Valid trend report preference changed.'
    }
} finally { Remove-Module $module -ErrorAction SilentlyContinue }

# Invoke complete generators against synthetic sources to verify existing output bytes survive.
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('EndpointEvidenceOffline-'+[guid]::NewGuid().ToString('N'))
try {
    [void](New-Item -Path $testRoot -ItemType Directory)
    $latest=Join-Path $testRoot 'DATA-LAST'
    [void](New-Item -Path $latest -ItemType Directory)
    foreach ($name in @('Intune_Devices_Inventory','M365_Users_Active','M365_Entra_Devices','AD_Computers_AllDomains','Intune_Devices_UpgradeEligibility','Intune_WindowsUpdate_Status','M365_Entra_Devices_HardwareIdConflicts','M365_Entra_Devices_RemovalCandidates')) {
        [IO.File]::WriteAllText((Join-Path $latest ($name+'.csv')), '"Synthetic empty header"'+[Environment]::NewLine)
    }
    @($first,$second) | Export-Csv -LiteralPath (Join-Path $latest 'Intune_EndpointAnalytics_DevicePerformance.csv') -NoTypeInformation -Encoding utf8NoBOM
    @(New-TestRow -Report 'EAStartupPerfDevicePerformanceV2') | Export-Csv -LiteralPath (Join-Path $latest 'Intune_EndpointAnalytics_StartupDevices.csv') -NoTypeInformation -Encoding utf8NoBOM
    $currentOutputs=@('current.csv','signals.csv','directory.csv') | ForEach-Object { Join-Path $testRoot $_ }
    foreach ($output in $currentOutputs) { [IO.File]::WriteAllText($output,'synthetic-existing-output') }
    $hashes=@($currentOutputs | ForEach-Object { (Get-FileHash -LiteralPath $_).Hash })
    $failure=Get-TestFailure { & $currentPath -DataRoot $testRoot -OutputPath $currentOutputs[0] -SignalsOutputPath $currentOutputs[1] -DirectorySummaryOutputPath $currentOutputs[2] }
    Assert-Test ($failure.Exception.Message -match 'evidence rejected') 'Full current generator failed for an unrelated reason.'
    for ($i=0;$i -lt $currentOutputs.Count;$i++) { Assert-Test ((Get-FileHash -LiteralPath $currentOutputs[$i]).Hash -eq $hashes[$i]) 'Current generator overwrote prior evidence after a conflict.' }
    @($first,$appOnly) | Export-Csv -LiteralPath (Join-Path $latest 'Intune_EndpointAnalytics_DevicePerformance.csv') -NoTypeInformation -Encoding utf8NoBOM
    $startupConflict=@(New-TestRow -Report 'EAStartupPerfDevicePerformanceV2'; New-TestRow -Report 'EAStartupPerfDevicePerformanceV2' -Score '88')
    $startupConflict[1].StopErrorCount='3'
    $startupConflict | Export-Csv -LiteralPath (Join-Path $latest 'Intune_EndpointAnalytics_StartupDevices.csv') -NoTypeInformation -Encoding utf8NoBOM
    $failure=Get-TestFailure { & $currentPath -DataRoot $testRoot -OutputPath $currentOutputs[0] -SignalsOutputPath $currentOutputs[1] -DirectorySummaryOutputPath $currentOutputs[2] }
    Assert-Test ($failure.Exception.Message -match 'evidence rejected') 'Full current generator selected conflicting startup evidence.'
    for ($i=0;$i -lt $currentOutputs.Count;$i++) { Assert-Test ((Get-FileHash -LiteralPath $currentOutputs[$i]).Hash -eq $hashes[$i]) 'Current generator changed an output after a startup conflict.' }

    $windowsRoot=Join-Path $testRoot 'DATA-ALL/Intune/Devices/Inventory/WeeklyHistory'
    $endpointRoot=Join-Path $testRoot 'DATA-ALL/Intune/EndpointAnalytics/WeeklyHistory'
    foreach ($week in @('2026-W01','2026-W02')) {
        $windowsWeek=Join-Path $windowsRoot $week
        $endpointWeek=Join-Path $endpointRoot $week
        [void](New-Item -Path $windowsWeek -ItemType Directory -Force)
        [void](New-Item -Path $endpointWeek -ItemType Directory -Force)
        @([pscustomobject]@{OS='Windows';'OS version'='10.0.22631'}) | Export-Csv -LiteralPath (Join-Path $windowsWeek 'Intune_Devices_Inventory.csv') -NoTypeInformation -Encoding utf8NoBOM
        @($first,$appOnly) | Export-Csv -LiteralPath (Join-Path $endpointWeek 'Intune_EndpointAnalytics_DevicePerformance.csv') -NoTypeInformation -Encoding utf8NoBOM
        @(New-TestRow -Report 'EAStartupPerfDevicePerformanceV2') | Export-Csv -LiteralPath (Join-Path $endpointWeek 'Intune_EndpointAnalytics_StartupDevices.csv') -NoTypeInformation -Encoding utf8NoBOM
    }
    $windowsOutput=Join-Path $testRoot 'windows.csv'
    $endpointOutput=Join-Path $testRoot 'endpoint.csv'
    $result=(& $trendPath -DataRoot $testRoot -WindowsOutputPath $windowsOutput -EndpointOutputPath $endpointOutput) | ConvertFrom-Json
    $trend=@(Import-Csv -LiteralPath $endpointOutput)
    Assert-Test ($result.EndpointSnapshots -eq 2 -and $trend.Count -eq 2 -and $trend[0].'Endpoint Analytics Covered Devices' -eq '1' -and $trend[0].'Endpoint Analytics Score' -eq '71.0') 'Healthy weekly generator changed its coverage or scores.'
    $windowsHash=(Get-FileHash -LiteralPath $windowsOutput).Hash
    $endpointHash=(Get-FileHash -LiteralPath $endpointOutput).Hash
    @($first,$second) | Export-Csv -LiteralPath (Join-Path $endpointWeek 'Intune_EndpointAnalytics_DevicePerformance.csv') -NoTypeInformation -Encoding utf8NoBOM
    $failure=Get-TestFailure { & $trendPath -DataRoot $testRoot -WindowsOutputPath $windowsOutput -EndpointOutputPath $endpointOutput }
    Assert-Test ($failure.Exception.Message -match 'evidence rejected') 'Full weekly generator failed for an unrelated reason.'
    Assert-Test ((Get-FileHash -LiteralPath $windowsOutput).Hash -eq $windowsHash -and (Get-FileHash -LiteralPath $endpointOutput).Hash -eq $endpointHash) 'Conflict replaced, skipped or shortened historical output.'
    Assert-Test (@(Import-Csv -LiteralPath (Join-Path $endpointWeek 'Intune_EndpointAnalytics_DevicePerformance.csv')).Count -eq 2) 'Generator rewrote the ambiguous input.'
    @($first,$appOnly) | Export-Csv -LiteralPath (Join-Path $endpointWeek 'Intune_EndpointAnalytics_DevicePerformance.csv') -NoTypeInformation -Encoding utf8NoBOM
    $startupConflict | Export-Csv -LiteralPath (Join-Path $endpointWeek 'Intune_EndpointAnalytics_StartupDevices.csv') -NoTypeInformation -Encoding utf8NoBOM
    $failure=Get-TestFailure { & $trendPath -DataRoot $testRoot -WindowsOutputPath $windowsOutput -EndpointOutputPath $endpointOutput }
    Assert-Test ($failure.Exception.Message -match 'evidence rejected') 'Full weekly generator selected conflicting startup evidence.'
    Assert-Test ((Get-FileHash -LiteralPath $windowsOutput).Hash -eq $windowsHash -and (Get-FileHash -LiteralPath $endpointOutput).Hash -eq $endpointHash) 'Weekly startup conflict replaced historical output.'

    & (Join-Path $PSScriptRoot 'Test-PreparedEvidencePipeline.ps1') -TestRoot $testRoot
    Assert-Test ($true) 'Existing prepared publication and stable historyKey regressions passed.'
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $tempRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if (-not $resolved.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'EndpointEvidenceOffline-*') { throw 'Unsafe synthetic cleanup path.' }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
[pscustomobject]@{Status='Passed';Checks=$script:Checks;Sources='Synthetic only';LiveCalls=0}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDO66VUsVsB/E/D
# UAeszEI5QYyXsP1Nk0QuPyltQ7/Bo6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMB7TddJbdU7jTgIuAbyTRyH9cT6XAIIA4g1Q+mAbh5BMA0GCSqG
# SIb3DQEBAQUABIIBgEfM7b7d+sZ+RcGr96SN759rQLy+Yw4qTgITKM5bR6IzUd3N
# /HgQ+8/K2siVLLThYatIQQGFmehMX2ke1kGcLONqUbwVR95sDt4vDVynznA0ga5C
# Nc/xHUfvITWIEmmNfd2a6reWGE5tTxRlnM37y+DR9cp+lVJ9+HjGIED/1unwrhrd
# rxSy0PQIunqT5UEImqFIMrFm1jma/QGwLxLw0qPN6qzOW5+hHzuLwpQ9N9hjndcd
# 8zluhCmHG4L46xS+aaKFrG3OdfGxejuK2F1+ItIJil013RcxFOQ044CsOmzMTOOv
# hqvKgWnOf1zFeuHI1mifdmuqf018XHr8+/i61TbHzupONkpSi1wykqFKNZhs4Bo7
# UEdr7PR6SKhHte6NURA5wc8wdGrxbNyvzHchju0TCSXYN14OCBoX09zmyHG5IwC8
# wWE5re1KPxU8MFFabG+dFQr7uLxDBV5J00NZQrcH0JVG6TIJa/98YopUuQ/UX2i8
# 7ETRlHLufm0kmzyACaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMyMjM3
# NDZaMC8GCSqGSIb3DQEJBDEiBCAHFHRLTPUsKxEpEttfhPfUybEP3FnEGPL2v+2O
# wjJ4VzANBgkqhkiG9w0BAQEFAASCAgB/i3QF4WQHuf8aDYiGngNo0e7I1FcgqNsD
# 08OpJam3l4dadQcBzSwSsxD/+g+k50xswoO5xtKd1y0xoF0S5mWYHyzbVMakzpej
# Sw4FWNq0xrDbPPBjgRCFUx0N+5nR+W27z/KS3ECirb//JUihTOO6djdRhgadjGri
# mixxMhoGgEqNpm4s4lK4JcupYPDrK2QhS12Cz+HaB1Ex5uvpiAHmPfgHp141gf4E
# nU7bSWg08ZDQiKFWKeDlvP5XL2t6Iq8/6Jbh4KxSnBnX5gMqxs4JCCsv9dS30ZHX
# u7CSRKFo762ncTQPHX4css0vHDjMJvbrrV1QLE7efWAXG9a1JIDKNz4q7IDAgL6h
# yLeP1OWgvJ28wqaQ2MJIWW+4Djs8kCmc45LipPWC0PuYHXrIss2UxhplZsaxD1uf
# 5EoykfFyR/N9KjPnItt4XgMTid5oyMywYN4FidHb2Se/nTCCON1+KsYbjF8k+/3Q
# KoSNz81avLbuG6Z+tcJnCnYWonzMKLlcu/3XkHwHO7o5R0X0hw74Owg6+KmYdfz6
# yd5Ni5ZWcn8m8opFf1h/bbGCXrshnn4ebIWacpKEzXxrcxlfN5JJXGsVyipowDLI
# J6eAn4CnlFoX4yd5S8adUkFcd5H6KO92jezSDOF/P2wkX4+5r6IIO9M+qb0g8R5l
# SoBlNr8jLw==
# SIG # End signature block
