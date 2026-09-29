[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$WindowsOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\WindowsLifecycleTrendEvidence.csv'),

    [string]$EndpointOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\EndpointExperienceTrendEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Convert-ToDoubleOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    $parsed = 0.0
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    return ([double]$Value).ToString('0.0########', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-WindowsGeneration {
    param([AllowNull()][object]$OperatingSystem, [AllowNull()][object]$OperatingSystemVersion)
    if ([string]$OperatingSystem -notmatch '(?i)windows') { return 'Not Windows' }
    $parts = ([string]$OperatingSystemVersion).Split('.')
    $build = 0
    if ($parts.Count -ge 3 -and [int]::TryParse($parts[2], [ref]$build)) {
        if ($build -ge 22000) { return 'Windows 11' }
        if ($build -ge 10240) { return 'Windows 10' }
    }
    return 'Windows legacy or unknown'
}

function Get-LatestRowsByDevice {
    param(
        [object[]]$Rows,
        [string]$PreferredProperty
    )

    $index = @{}
    foreach ($row in $Rows) {
        $key = ([string]$row.DeviceId).Trim().ToLowerInvariant()
        if (-not $key) { continue }
        if (-not $index.ContainsKey($key)) {
            $index[$key] = $row
            continue
        }

        $candidateHasValue = -not [string]::IsNullOrWhiteSpace([string]$row.$PreferredProperty)
        $currentHasValue = -not [string]::IsNullOrWhiteSpace([string]$index[$key].$PreferredProperty)
        $candidateDate = [datetime]::MinValue
        $currentDate = [datetime]::MinValue
        [void][datetime]::TryParse([string]$row.ReportRefreshDate, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$candidateDate)
        [void][datetime]::TryParse([string]$index[$key].ReportRefreshDate, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$currentDate)
        if (($candidateHasValue -and -not $currentHasValue) -or
            ($candidateHasValue -eq $currentHasValue -and $candidateDate -gt $currentDate)) {
            $index[$key] = $row
        }
    }
    return @($index.Values)
}

function Get-AverageOrNull {
    param([object[]]$Values)
    $numbers = @($Values | Where-Object { $null -ne $_ })
    if ($numbers.Count -eq 0) { return $null }
    return [math]::Round((($numbers | Measure-Object -Average).Average), 1)
}

$windowsHistoryRoot = Join-Path $DataRoot 'DATA-ALL\Intune\Devices\Inventory\WeeklyHistory'
$endpointHistoryRoot = Join-Path $DataRoot 'DATA-ALL\Intune\EndpointAnalytics\WeeklyHistory'
foreach ($requiredRoot in @($windowsHistoryRoot, $endpointHistoryRoot)) {
    if (-not (Test-Path -LiteralPath $requiredRoot -PathType Container)) {
        throw "Required weekly-history folder was not found: $requiredRoot"
    }
}

$windowsTrend = foreach ($weekFolder in (Get-ChildItem -LiteralPath $windowsHistoryRoot -Directory | Sort-Object Name)) {
    $inventoryPath = Join-Path $weekFolder.FullName 'Intune_Devices_Inventory.csv'
    if (-not (Test-Path -LiteralPath $inventoryPath -PathType Leaf)) { continue }

    $inventory = @(Import-Csv -LiteralPath $inventoryPath)
    $windowsRows = @($inventory | Where-Object { $_.OS -match '(?i)windows' })
    $windows10 = @($windowsRows | Where-Object { (Get-WindowsGeneration $_.OS $_.'OS version') -eq 'Windows 10' }).Count
    $windows11 = @($windowsRows | Where-Object { (Get-WindowsGeneration $_.OS $_.'OS version') -eq 'Windows 11' }).Count
    $knownModernWindows = $windows10 + $windows11
    $snapshotDate = (Get-Item -LiteralPath $inventoryPath).LastWriteTime.Date

    [pscustomobject][ordered]@{
        'Snapshot Date' = $snapshotDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        'Snapshot Week' = $weekFolder.Name
        'Managed Devices' = $inventory.Count
        'Windows Devices' = $windowsRows.Count
        'Windows 10 Devices' = $windows10
        'Windows 11 Devices' = $windows11
        'Windows 11 Adoption (%)' = Convert-ToInvariantDecimalText $(if ($knownModernWindows -gt 0) { $windows11 / $knownModernWindows } else { $null })
        'Evidence Status' = 'Observed'
    }
}

$endpointTrend = foreach ($weekFolder in (Get-ChildItem -LiteralPath $endpointHistoryRoot -Directory | Sort-Object Name)) {
    $performancePath = Join-Path $weekFolder.FullName 'Intune_EndpointAnalytics_DevicePerformance.csv'
    $startupPath = Join-Path $weekFolder.FullName 'Intune_EndpointAnalytics_StartupDevices.csv'
    if (-not (Test-Path -LiteralPath $performancePath -PathType Leaf) -or -not (Test-Path -LiteralPath $startupPath -PathType Leaf)) { continue }

    $performance = Get-LatestRowsByDevice -Rows @(Import-Csv -LiteralPath $performancePath) -PreferredProperty 'EndpointAnalyticsScore'
    $startup = Get-LatestRowsByDevice -Rows @(Import-Csv -LiteralPath $startupPath) -PreferredProperty 'StartupScore'
    $endpointScores = @($performance | ForEach-Object { Convert-ToDoubleOrNull $_.EndpointAnalyticsScore } | Where-Object { $null -ne $_ })
    $startupScores = @($startup | ForEach-Object { Convert-ToDoubleOrNull $_.StartupScore } | Where-Object { $null -ne $_ })
    $appScores = @($performance | ForEach-Object { Convert-ToDoubleOrNull $_.AppReliabilityScore } | Where-Object { $null -ne $_ -and $_ -ge 0 })
    $observedDates = @($performance + $startup | ForEach-Object {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse([string]$_.ReportRefreshDate, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { $parsed }
    })
    $snapshotDate = if ($observedDates.Count -gt 0) { ($observedDates | Sort-Object -Descending | Select-Object -First 1).Date } else { (Get-Item -LiteralPath $performancePath).LastWriteTime.Date }

    [pscustomobject][ordered]@{
        'Snapshot Date' = $snapshotDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        'Snapshot Week' = $weekFolder.Name
        'Endpoint Analytics Covered Devices' = $endpointScores.Count
        'Endpoint Analytics Score' = Convert-ToInvariantDecimalText (Get-AverageOrNull $endpointScores)
        'Endpoint Analytics Below 50 Devices' = @($endpointScores | Where-Object { $_ -lt 50 }).Count
        'Startup Covered Devices' = $startupScores.Count
        'Startup Score' = Convert-ToInvariantDecimalText (Get-AverageOrNull $startupScores)
        'Startup Below 50 Devices' = @($startupScores | Where-Object { $_ -lt 50 }).Count
        'App Reliability Covered Devices' = $appScores.Count
        'App Reliability Score' = Convert-ToInvariantDecimalText (Get-AverageOrNull $appScores)
        'Devices with Stop Errors' = @($startup | Where-Object { $null -ne (Convert-ToDoubleOrNull $_.StopErrorCount) -and (Convert-ToDoubleOrNull $_.StopErrorCount) -gt 0 }).Count
        'Devices with Boot Time over 60s' = @($startup | Where-Object { $null -ne (Convert-ToDoubleOrNull $_.CoreBootTime) -and (Convert-ToDoubleOrNull $_.CoreBootTime) -gt 60 }).Count
        'Devices with Sign-in Time over 30s' = @($startup | Where-Object { $null -ne (Convert-ToDoubleOrNull $_.CoreSignInTime) -and (Convert-ToDoubleOrNull $_.CoreSignInTime) -gt 30 }).Count
        'Evidence Status' = 'Observed'
    }
}

foreach ($outputPath in @($WindowsOutputPath, $EndpointOutputPath)) {
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
}

$windowsTrend | Export-Csv -LiteralPath $WindowsOutputPath -NoTypeInformation -Encoding utf8NoBOM
$endpointTrend | Export-Csv -LiteralPath $EndpointOutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject][ordered]@{
    WindowsOutputPath = $WindowsOutputPath
    WindowsSnapshots = $windowsTrend.Count
    EndpointOutputPath = $EndpointOutputPath
    EndpointSnapshots = $endpointTrend.Count
    LatestWindowsSnapshot = if ($windowsTrend.Count -gt 0) { $windowsTrend[-1].'Snapshot Date' } else { $null }
    LatestEndpointSnapshot = if ($endpointTrend.Count -gt 0) { $endpointTrend[-1].'Snapshot Date' } else { $null }
} | ConvertTo-Json -Depth 3

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDT/8eRjg60yTPI
# 6yMbPY+AP+Lj9jX4xexNTlyi0v7sVaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILp/TfIGhK1/8/TT19X70UBbqJOgK8GLous/Yx9e5mz2MA0GCSqG
# SIb3DQEBAQUABIIBgEchdOpf9ssU5YXG6OUGKdtsnFo6mIEttB83KxzGEKMRsgNr
# wntifzUEtBYK8OFNGsilfece5bRu3WbjEGSg8Jnes6IZPWYdkBGyw4Zv5FKmVf0E
# FFlDXcRxnBqIVnQ+s9SfiGEG9rgSALC+uWLh79IJitFoYVfMfQ4CoLhKplmsJe/B
# GQoPhgbc4kO3nS1AaqoqyYgH3/NTzIk+Em80TyPEIiuuzl1ubreLxmr+51kZkxkf
# RIiqioaiAqXmibnRP+vs8DGtwetyJkAwLat4tBdA71tK3HxuOPHSlci2Eu2+1zzW
# PRlyE1CfX5slpDLREdhHBHDE64Dm5VRQXGgiCs79mWO0OfhvreVcLeRJ9Q6jYWdS
# bwCdJEzAxQvi1UqBMrHt/8pTtGnYHiFwadjD8gljpJoWu1PhVzqRfBuVqNg6YaYk
# l9Sf9jtsQNneKLNOZjKzq9r6qPjUQYzSAVj0DrJm/PYC0N3Z4C8iXOO951Syck90
# 2mJgLyIFWH6fNXBefqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjkxOTEx
# MzBaMC8GCSqGSIb3DQEJBDEiBCBjEYA5iQZiavWQDDCO1ODjbcm8ZoTmRrJ+ObG5
# Lfn4pzANBgkqhkiG9w0BAQEFAASCAgCusnYUiBddtrblBNVP1H5rDWXm+d3rOcKx
# u9hUyIXn4hlkIJ7NTl+65sxf/c/lOa3DDRBiH3GocTzrGjEA8bdHy7l1N637D4Ph
# PlF1nJW95LFk4/XIjyBOcUaT4uLDXrpS5zH5nWc9lKFb2q1H79mVniYEycuZ3l0c
# 3TWT4KXbb1KOi5HFPwMJ5248zR73zyEbioKSjZBVilGGqoU+y0HAZ5oPuprsMQTC
# ZNF6pWaLnpiHLlYFM4NLtVwEX3Qvx/bpCYfr3prMA8YHuT5smO3Xk9dilqTufIvS
# IyB/jJ+fY+zWirjLOj5wIB0mKzdTdONe4flAdC7LbJkr40JG9JltG3UoMz17n3g0
# eaYmZL6Jynae9kZ8XV21Bl4WiwqbfymGzWD3f1V9zC+ZdSbgvIn4jRRm+MdKRSWA
# LT/UbMwRBBFknlncc+71fTDlMQMAxguNrgMzshV7N3gBNsZfPJDGjk5dwtOhvIUu
# lif2B+WVZ1PwM/V9qGa7me+mt4egBSaOpEiO2q6ol7PqyArEQd20N3+wm3ZKguHV
# GB9gE1nprU139rmdXbdlckBGAhNftVN5ZPFTXp9WeZX/9SRaotCdZ7sdsXa/xgfN
# +FxGO1Cafi/BkAxZwaFrdomWbG4qUzfLk/YnQ7UrNNwRoURJI4qCSuJlc+fOMElp
# AzYvI/QjVA==
# SIG # End signature block
