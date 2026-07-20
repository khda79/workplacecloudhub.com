<#
.SYNOPSIS
    Offline regression test for Discovered Apps structured logging.
.DESCRIPTION
    Verifies 429 status inference, quiet retry error streams, explicit final CSV
    sample/physical row labels, and quiet retirement of an already absent
    SharePoint legacy file. No Graph connection or production CSV is used.
.VERSION
1.0
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$inventoryScriptPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Intune-DiscoveredApps-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }

$requiredFunctions = @(
    'Get-DiscoveredAppsGraphHttpErrorMessage',
    'Invoke-GraphPagedRequest',
    'Get-DiscoveredAppsCsvDataRowCount',
    'Complete-DiscoveredAppsStreamExport',
    'Remove-LegacyDiscoveredAppsDeviceDetailExport'
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

$coreModulePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psm1')).Path
$coreTokens = $null
$coreParseErrors = $null
$coreAst = [System.Management.Automation.Language.Parser]::ParseFile($coreModulePath, [ref]$coreTokens, [ref]$coreParseErrors)
if ($coreParseErrors.Count -gt 0) { throw ($coreParseErrors | Out-String) }
$coreDeleteDefinition = $coreAst.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Invoke-SmartM365GraphDeleteQuietly'
}, $true) | Select-Object -First 1
if (-not $coreDeleteDefinition) { throw 'Invoke-SmartM365GraphDeleteQuietly was not found in SmartM365.Core.' }
Invoke-Expression $coreDeleteDefinition.Extent.Text

$script:TestLogs = [System.Collections.Generic.List[string]]::new()
function WriteLog {
    param([string]$Message, [string]$Level)
    $script:TestLogs.Add("$Level|$Message") | Out-Null
}
function Assert-Equal {
    param($Actual, $Expected, [string]$Label)
    if ($Actual -ne $Expected) { throw "$Label expected '$Expected', got '$Actual'." }
}

$script:GraphAttempt = 0
$script:Stat_GraphCalls = 0
$script:Stat_ThrottleRetries = 0
$script:GraphRetryMaxSeconds = 1
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param(
        [string]$Method,
        [string]$Uri,
        [string]$OutputType,
        [string]$Body,
        [string]$ContentType,
        [switch]$SkipHttpErrorCheck,
        [string]$StatusCodeVariable
    )
    $script:GraphAttempt++
    if ($script:GraphAttempt -eq 1) {
        Set-Variable -Name $StatusCodeVariable -Value 429 -Scope 1
        return [pscustomobject]@{ error = [pscustomobject]@{ message = 'TooManyRequests synthetic test response' } }
    }
    Set-Variable -Name $StatusCodeVariable -Value 200 -Scope 1
    return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'device-1' }) }
}
function Get-ShortGraphErrorMessage { param($ErrorRecord) return [string]$ErrorRecord.Exception.Message }
function Get-GraphRetryDelaySeconds {
    param($ErrorRecord, [int]$Attempt, [int]$DefaultSeconds, [int]$MaximumSeconds)
    return 0
}
function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) }
function Test-SmartM365MaxItemsMode { return $false }
function Remove-SmartM365SharePointFile {
    [CmdletBinding()]
    param([string]$LocalFilePath)
    return $true
}
function Get-SmartM365GraphAccessToken { param([string]$Purpose) return 'synthetic-token' }
function Invoke-RestMethod {
    [CmdletBinding()]
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Headers,
        [switch]$SkipHttpErrorCheck,
        [string]$StatusCodeVariable,
        [string]$ResponseHeadersVariable
    )
    Set-Variable -Name $StatusCodeVariable -Value 404 -Scope 1
    Set-Variable -Name $ResponseHeadersVariable -Value @{} -Scope 1
    return [pscustomobject]@{ error = [pscustomobject]@{ message = 'itemNotFound' } }
}

$testRoot = Join-Path $env:TEMP ("SmartM365-DiscoveredApps-LoggingTest-" + [guid]::NewGuid().ToString('N'))
try {
    $pagedResult = @(Invoke-GraphPagedRequest -InitialUri 'https://graph.microsoft.com/v1.0/test' -MaxRetries 2 -DefaultRetrySeconds 0)
    Assert-Equal $pagedResult.Count 1 'Paged result count'
    Assert-Equal $script:Stat_ThrottleRetries 1 'Throttle retry count'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*Status=429; attempt 1/2*' }).Count 1 'Structured 429 status log'

    $quietDelete = Invoke-SmartM365GraphDeleteQuietly -Uri 'https://graph.microsoft.com/v1.0/test-delete' -MaxAttempts 1
    Assert-Equal $quietDelete.Success $true 'Quiet delete success'
    Assert-Equal $quietDelete.NotFound $true 'Quiet delete not found'
    Assert-Equal $quietDelete.StatusCode 404 'Quiet delete status'

    $outputPath = Join-Path $testRoot 'current'
    $latestPath = Join-Path $testRoot 'latest'
    New-Item -ItemType Directory -Path $outputPath, $latestPath -Force | Out-Null
    $partialPath = Join-Path $testRoot 'relations.partial.csv'
    $timestampedPath = Join-Path $outputPath 'Intune_DiscoveredApps_AppDeviceRelations_20260720_200000.csv'
    @(
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-1'; DeviceId = 'device-1' }
        [pscustomobject]@{ TenantKey = 'tenant-test'; AppId = 'app-2'; DeviceId = 'device-2' }
    ) | Export-Csv -LiteralPath $partialPath -NoTypeInformation -Encoding UTF8

    $global:csvGeneratedPaths = $null
    $publishedPath = Complete-DiscoveredAppsStreamExport `
        -PartialPath $partialPath `
        -TimestampedPath $timestampedPath `
        -OutputPath $outputPath `
        -GlobalPath $latestPath `
        -BaseFileName 'Intune_DiscoveredApps_AppDeviceRelations' `
        -ExpectedDataRows 2
    Assert-Equal $publishedPath $timestampedPath 'Published path'
    Assert-Equal (Test-Path -LiteralPath (Join-Path $latestPath 'Intune_DiscoveredApps_AppDeviceRelations.csv')) $true 'Latest relation CSV'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*SampleRows=1; CriticalFields=TenantKey, AppId, DeviceId*' }).Count 1 'SampleRows log'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*PhysicalRows=2*' }).Count 1 'PhysicalRows log'

    $global:EnableSharePointUpload = $true
    $legacyErrorOutput = @(Remove-LegacyDiscoveredAppsDeviceDetailExport -CurrentOutputPath $outputPath -LatestOutputPath $latestPath 2>&1)
    Assert-Equal @($legacyErrorOutput | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] }).Count 0 'Quiet legacy SharePoint 404'
    Assert-Equal @($script:TestLogs | Where-Object { $_ -like '*Legacy export Intune_DiscoveredApps_DeviceDetail.csv is disabled*' }).Count 1 'Legacy retirement log'

    Write-Host 'PASS: structured 429, explicit CSV row labels, and quiet legacy 404 verified.' -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOMOHTRu1IeCVg
# gDt2SDXnt36EtMZ5Mlve8eBBy7PlHaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIPECyyjqXy8f9oplkJ3ONFYL1ffwl+R5gvP0mGBTQZzXMA0GCSqG
# SIb3DQEBAQUABIIBgFyR9uM6f4Wk6yfHo07zYWC9+6I8MzCXgIUz+stTP4AWVWCF
# R3SI3db5hFNx4nJyQEe+ugGNmaII1Yv5z3SxyMQZ0hna7JfHml1gcGV+6xlr29dl
# Ok/TCe5HRUty2IOFLOrxKJ4w1tzC9YeDI/FU8FfXaVF5aVJmNlrPxR6I9nO+dllv
# ycFZDfE/GnZ1qdLZZuKNjQvuMkA5eKVq7uUbPpsryqXJ4n0tJn1Td+R8P6/JkAms
# eqcmbKzcG/LEZ5gwXyQK3fOU75aUb6RGlyB2WKaun9uU6EFYoFQB/z1CYrc4yz8S
# 6eFPquoR+l1533VZG2j845134zJdphrCF1RV/M4iBSai902npuj5FMefWal69dyG
# ZCHREGumPvTBrETBm/Zvz+pUmsp/ghwXAap8kTVvCSQjY3Fu13adCQZkXP5Bqsqo
# PBKOFCrXggae8xqCEkdFZOQwfjQpjL7F4vf2/XHiWaOWpWMWOSGUKxwJr+oJtLmU
# oSEqFImJ4WP+6J1LraGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjAxOTQ0
# NDhaMC8GCSqGSIb3DQEJBDEiBCDR2ah+H9w1cCqYqRo7uawYK9jFahJkPs1TOu2M
# 1X1NgDANBgkqhkiG9w0BAQEFAASCAgDARgF6bvrORE2gXxS42dPhe2eIfFR/eUdo
# 0uY/uhUaYHi6By+JYM0GcXXCaDKPtwIYrkBLJ4ruk/KpnrXrClCijfkvCao/zVBa
# BVlf4kmGsCRBg85U9ljg/Wc66AjyVh4GQ59TKSjpuyJ9+zUDAEGG1kR1sBYXC6Ur
# djmqD7O2M/KN5T9ktQStdcLyRM4Y+LAxPgjgwzVs3MUrwO/Ud7HKusVhzosyORn6
# cl8bnaN9xrUFGNdr4/BDXlaOfeKr9UbT1h9qGnGX1bO6YxtOPZM1lZ6kSHiOfW+m
# 5hIyTAWEYtvejdKJa1fV5nfu5RRcg2fHRkNBreyCu9RhmtQSOcah3PJ2ThzszKLC
# FEe2K4Ilf8j3YsmxwMjrh2oTJVDGRt2PknJ9m0Ri3pFIQrI9myXuWByS5MO0nb8H
# jEOtsM+vdXtf7to6ZF5jEdkgki8Z2xoj7lctEcjWl/HdYUQDIuDdvW0K1LhPn/Rs
# 8oD84+elxGdX73Z7dG3gTmgUI8SfIe3vzW73dO/I7utCe89iCEkySjluu09Iu/L2
# ybCayx+Dh39M3R7NvZMtx0KJkZA1BgLTJgjuMSuHFPAwkBVrniBkWgPMZ5EPrl1N
# I5FmaRUlqyyPy4oA4aO+CfXgQh80bMvWvMRefQ9ZVEeBKy2tM5zdm2eb5c6XxNgN
# WnR/nf3ozQ==
# SIG # End signature block
