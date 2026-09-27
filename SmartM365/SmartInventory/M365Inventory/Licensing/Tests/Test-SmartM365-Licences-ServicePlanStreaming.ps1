<#
.SYNOPSIS
    Offline regression test for disk-backed Microsoft 365 service-plan states.
.DESCRIPTION
    Verifies compact CSV writing and validation, detailed WeeklyHistory decoding,
    and safe retirement of duplicate weekly files. No Graph connection or
    production CSV is used.
.VERSION
1.0
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../../../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force

$inventoryScriptPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }

$requiredFunctions = @(
    'Get-ServicePlanStateCode',
    'ConvertFrom-ServicePlanStateCode',
    'ConvertTo-LicensesCsvField',
    'New-LicensesServicePlanStateWriter',
    'Write-LicensesServicePlanStateRow',
    'Close-LicensesServicePlanStateWriter',
    'Get-LicensesCsvDataRowCount',
    'Assert-LicensesServicePlanStateCsvFile',
    'Copy-LicensesCsvAtomically',
    'Publish-LicensesServicePlanStateCsvFile',
    'New-LicensesDetailedServicePlanHistorySource',
    'Get-LicensesIsoWeekName',
    'Publish-LicensesWeeklyHistory',
    'Remove-LegacyWeeklyServicePlanStateDuplicates'
)
$definitions = @($ast.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $requiredFunctions -contains $node.Name
}, $true) | Sort-Object { $_.Extent.StartOffset })
if ($definitions.Count -ne $requiredFunctions.Count) {
    throw "Expected $($requiredFunctions.Count) function definitions, found $($definitions.Count)."
}
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }

$script:TestLogs = [System.Collections.Generic.List[string]]::new()
function WriteLog {
    param([string]$Message, [string]$Level)
    $script:TestLogs.Add("$Level|$Message") | Out-Null
}
$script:MaxItemsMode = $false
function Test-SmartM365MaxItemsMode { return $script:MaxItemsMode }
function Add-SmartM365MaxItemsSuffixToBaseName {
    param([string]$BaseFileName)
    if ($script:MaxItemsMode) { return "${BaseFileName}_MAXITEMS-5" }
    return $BaseFileName
}
function Get-SmartM365MaxItemsSuffix { return '_MAXITEMS-5' }
function RemoveOldFiles {}
function Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) return $null }
function Assert-Equal {
    param($Actual, $Expected, [string]$Label)
    if ($Actual -ne $Expected) { throw "$Label expected '$Expected', got '$Actual'." }
}

$global:SmartM365OrganizationKey = 'org-test'
$global:SmartM365EnvironmentKey = 'test'
$global:SmartM365TenantId = 'tenant-id'
$global:EnableSharePointUpload = $false
$global:RetentionMaxCSV = 0
$global:csvGeneratedPaths = $null
$testRoot = Join-Path $env:TEMP ("SmartM365-LicenseStreamTest-" + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
    Assert-Equal (Get-LicensesIsoWeekName -Date ([datetime]::new(2026, 7, 20))) '2026-W30' 'ISO week name'
    $compactPath = Join-Path $testRoot 'states.building.csv'
    $writer = New-LicensesServicePlanStateWriter -Path $compactPath
    Write-LicensesServicePlanStateRow -Writer $writer -TenantKey 'tenant-test' -UserId 'user-1' -SkuId 'sku-1' -PlanId 'plan-1' -StateCode 'A' -RowNumber 1
    Write-LicensesServicePlanStateRow -Writer $writer -TenantKey 'tenant-test' -UserId 'user-1' -SkuId 'sku-1' -PlanId 'plan-2' -StateCode 'D' -RowNumber 2
    Write-LicensesServicePlanStateRow -Writer $writer -TenantKey 'tenant-test' -UserId 'user-2' -SkuId 'sku-2' -PlanId 'plan-3' -StateCode 'PP' -RowNumber 3
    Close-LicensesServicePlanStateWriter -Writer $writer

    Assert-LicensesServicePlanStateCsvFile -Path $compactPath -ExpectedRows 3
    Assert-Equal (Get-LicensesCsvDataRowCount -Path $compactPath) 3 'Compact rows'

    $detailedPath = Join-Path $testRoot 'M365_Licenses_UserServicePlanStates_Detailed.csv'
    $detailCount = New-LicensesDetailedServicePlanHistorySource -CompactCsvPath $compactPath -DetailedCsvPath $detailedPath
    Assert-Equal $detailCount 3 'Detailed rows'
    $detailRows = @(Import-Csv -LiteralPath $detailedPath)
    Assert-Equal $detailRows[0].IsEnabled 'True' 'Enabled state'
    Assert-Equal $detailRows[0].PlanStatus 'Success' 'Enabled status'
    Assert-Equal $detailRows[1].IsEnabled 'False' 'Disabled state'
    Assert-Equal $detailRows[1].PlanStatus 'Disabled' 'Disabled status'
    Assert-Equal $detailRows[2].PlanStatus 'PendingProvisioning' 'Pending status'

    $publishCurrent = Join-Path $testRoot 'publish-current'
    $publishLatest = Join-Path $testRoot 'publish-latest'
    New-Item -ItemType Directory -Path $publishCurrent, $publishLatest -Force | Out-Null
    $publishBuilding = Join-Path $publishCurrent 'states.building.csv'
    Copy-Item -LiteralPath $compactPath -Destination $publishBuilding
    $publication = Publish-LicensesServicePlanStateCsvFile -BuildingPath $publishBuilding -ExpectedRows 3 -CurrentOutputPath $publishCurrent -LatestOutputPath $publishLatest
    Assert-Equal (Test-Path -LiteralPath $publication.TimestampedPath) $true 'Timestamped publication'
    Assert-Equal (Test-Path -LiteralPath $publication.CurrentPath) $true 'Current publication'
    Assert-Equal (Test-Path -LiteralPath $publication.PublishedPath) $true 'Latest publication'
    Assert-Equal (Get-LicensesCsvDataRowCount -Path $publication.PublishedPath) 3 'Published rows'

    $script:MaxItemsMode = $true
    $maxCurrent = Join-Path $testRoot 'max-current'
    $maxLatest = Join-Path $testRoot 'max-latest'
    New-Item -ItemType Directory -Path $maxCurrent, $maxLatest -Force | Out-Null
    $maxBuilding = Join-Path $maxCurrent 'states.building.csv'
    Copy-Item -LiteralPath $compactPath -Destination $maxBuilding
    $maxPublication = Publish-LicensesServicePlanStateCsvFile -BuildingPath $maxBuilding -ExpectedRows 3 -CurrentOutputPath $maxCurrent -LatestOutputPath $maxLatest
    Assert-Equal ([System.IO.Path]::GetFileName($maxPublication.PublishedPath)) 'M365_Licenses_UserServicePlanStates_MAXITEMS-5.csv' 'MaxItems filename'
    Assert-Equal (Test-Path -LiteralPath (Join-Path $maxLatest 'M365_Licenses_UserServicePlanStates.csv')) $false 'Canonical MaxItems protection'
    $script:MaxItemsMode = $false
    $historyRoot = Join-Path $testRoot 'WeeklyHistory'
    $weekWithBoth = Join-Path $historyRoot '2026-W30'
    $weekLegacyOnly = Join-Path $historyRoot '2026-W29'
    New-Item -ItemType Directory -Path $weekWithBoth, $weekLegacyOnly -Force | Out-Null
    Copy-Item -LiteralPath $detailedPath -Destination (Join-Path $weekWithBoth 'M365_Licenses_UserServicePlanStates_Detailed.csv')
    Copy-Item -LiteralPath $detailedPath -Destination (Join-Path $weekWithBoth 'M365_Licenses_UserServicePlanStates.csv')
    Copy-Item -LiteralPath $detailedPath -Destination (Join-Path $weekLegacyOnly 'M365_Licenses_UserServicePlanStates.csv')

    $removed = Remove-LegacyWeeklyServicePlanStateDuplicates -HistoryRootPath $historyRoot -ExpectedCurrentWeekRows 3
    Assert-Equal $removed 1 'Removed duplicates'
    Assert-Equal (Test-Path -LiteralPath (Join-Path $weekWithBoth 'M365_Licenses_UserServicePlanStates.csv')) $false 'Duplicate removed'
    Assert-Equal (Test-Path -LiteralPath (Join-Path $weekWithBoth 'M365_Licenses_UserServicePlanStates_Detailed.csv')) $true 'Detailed preserved'
    Assert-Equal (Test-Path -LiteralPath (Join-Path $weekLegacyOnly 'M365_Licenses_UserServicePlanStates.csv')) $true 'Legacy-only week preserved'

    $manifest = Get-Content -LiteralPath (Join-Path $weekWithBoth 'manifest.json') -Raw | ConvertFrom-Json
    Assert-Equal @($manifest.Files | Where-Object { $_ -eq 'M365_Licenses_UserServicePlanStates.csv' }).Count 0 'Manifest duplicate removed'
    function Get-ScriptLocalConfigValue {
        param($Config, [string]$Name, $DefaultValue)
        return $DefaultValue
    }
    function Save-SmartM365WeeklyInventoryHistory {
        param(
            [string[]]$SourceFiles,
            [string]$HistoryRootPath,
            [int]$RetentionWeeks,
            [string]$HistoryLabel,
            [switch]$UploadChangedFilesOnly
        )
    }
    function Remove-LegacyWeeklyServicePlanStateDuplicates { throw 'simulated cleanup failure' }
    $ScriptLocalConfig = [pscustomobject]@{}
    $failureCurrent = Join-Path $testRoot 'failure-current'
    $failureLatest = Join-Path $testRoot 'failure-latest'
    New-Item -ItemType Directory -Path $failureCurrent, $failureLatest -Force | Out-Null
    $script:TestLogs.Clear()
    Publish-LicensesWeeklyHistory `
        -ServicePlanStateCsvPath $compactPath `
        -ExpectedServicePlanRows 3 `
        -CurrentOutputPath $failureCurrent `
        -LatestOutputPath $failureLatest
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like 'WARNING|Legacy WeeklyHistory service-plan duplicate cleanup failed*' }).Count 1 'Cleanup failure warning'
    Assert-Equal @(Get-ChildItem -LiteralPath $failureCurrent -Filter 'M365_Licenses_UserServicePlanStates_Detailed_*.csv' -File).Count 0 'Detailed cleanup after warning'

    Write-Host 'PASS: compact rows streamed; detailed states decoded; duplicate WeeklyHistory safely retired.' -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDEQcut5q1EalPu
# HThq8NE/wp8fSWyfci8Oqhs6XFoL8qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM2xzOp8UF7BJpzeKAyD6cVATi8RGsvTPhaTUAMkbWc9MA0GCSqG
# SIb3DQEBAQUABIIBgJRAItMPFDkNkfdunQ2J2+7I3IyP07pNe6xXXDKFvSA9thkd
# WFIiQ8hs9lo6bnF239NamZ6sWR62eOCw8beI2LAHs+IHB0H003L3LbTzgJbGvIFL
# tHecScK7k/KW+H5nzd+P1+W3Hlm4gYenmVMhCZLXkai0TQTgHHaaQe1l2Pqmq2u9
# RG4Xm+P0231m81mlQrwF/NF+6grAZ8uZRFJHso8qDQt+rScX/SxaGwi6rkD80u4D
# BsxpRN/nzPQco5Frlm7qZora/D+OmLMEF1sDhId3T5Rco/CKYThgSsoLQV+fh3GA
# uVdBm2h+ZT3W9r2kP1eXjYr8EeHF1VoV12AskWSIrxEgWGaMAifSrPRCCoYm9JQ8
# 5zegg4ULhE4KawNstLDe8yWxoxLGNts/ov3OeSu+1VlskugWSccw8vSyqpTKOkqt
# rVjCO6l6hsJaEtZ60yNDIK+rpehuvG21a1UNrrI/cdsV1C4qHHQx6wpzT2FwiQ9e
# d1EttOUnIc+7Tc3+XaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU3
# NDlaMC8GCSqGSIb3DQEJBDEiBCCsRL6LS+KlAS2O9lwy3NqsuDRU0XaJ2AFe16Hf
# U1P2kzANBgkqhkiG9w0BAQEFAASCAgCIS9yc1WSlJm18L69Z0UeRs/fJJ73CPM+b
# G8c+m+3J8il8cSz6hM6FKwv6n+1gt4QLF4a2/KbDQdacw2do1vr7FgkwAZN1GB3X
# JsokXMH+Ik5Zz4AI39WX90cH5l70JTe4yrKtt7GgwF2kXdt3JHQJoNX1i/n2GX2O
# q+tCNV8sco8Vf1nOoDjabmwOYFJKqtTlQOIawpnxnL4KK2QFEkjP//1ulFI01HTo
# uw0B8MP+mpOTJ24ZqvOhgCxSREK7kiTZ/CIWJqehqWYnd6Jzgy+uBExJhs7EhgWl
# tRWPl258kULBpcZxbKfTbJn24sMjuBO6A2bQ453SNRe0T5kvfey6zjQBzQOYeymn
# QL0MzAEMCm1uRDuQO3tcE9P/qNWC/1UaEJnDwE8G/g4XULCFcT4Ep1o8CssetU46
# NH2x/2kIPsiI2u7B2Gn4Yac3+gR7B86ThUbY9y7ue92EVIRkdnaBI9ysDEFBkuqy
# 5oggJCXBpAEgQXjr66TBo/wlfVAPRKpmfATN/GOESFtEzBpyoiysdnArI/NRRB7l
# LW5BsDXIktO36ykpKeK+umdADSg0aziy1GdqQxulklm6IKVfIn2muua88tYrsAyR
# NBJwYiJc8hGkw23LQc7oVYLiB5wxECijiDAvNDSlPBEZVv1iksYocLsFOLcgyMdm
# g/vyoAMnlg==
# SIG # End signature block
