[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$InventoryOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ApplicationInventoryEvidence.csv'),

    [string]$TrendOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ApplicationTrendEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    return ([double]$Value).ToString('0.0########', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ApplicationProfile {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][datetime]$SnapshotDate,
        [Parameter(Mandatory = $true)][string]$SnapshotWeek,
        [switch]$IncludeDetail
    )

    $detail = [System.Collections.Generic.List[object]]::new()
    $fragmentedApplications = 0
    $highlyFragmentedApplications = 0
    $installationsOutsideDominant = [int64]0
    $longTailVersions = 0
    $unknownPublisherVersions = 0

    $groups = @($Rows | Group-Object AppName)
    foreach ($group in $groups) {
        $applicationName = if ([string]::IsNullOrWhiteSpace($group.Name)) { 'Unknown application' } else { $group.Name.Trim() }
        $versions = @($group.Group | ForEach-Object { [string]$_.AppVersion } | Sort-Object -Unique)
        $versionCount = $versions.Count
        $productTotal = [int64](($group.Group | Measure-Object DeviceCount -Sum).Sum)
        $dominant = $group.Group | Sort-Object -Property @{ Expression = { [int64]$_.DeviceCount }; Descending = $true }, @{ Expression = { [string]$_.AppVersion }; Descending = $true }, @{ Expression = { [string]$_.AppId }; Descending = $false } | Select-Object -First 1
        $dominantCount = [int64]$dominant.DeviceCount
        $dominantShare = if ($productTotal -gt 0) { $dominantCount / $productTotal } else { $null }
        $standardizationState = if ($versionCount -ge 10) { 'Highly fragmented' } elseif ($versionCount -gt 1) { 'Fragmented' } else { 'Single version' }
        $riskPriority = if ($versionCount -ge 20) { 'High' } elseif ($versionCount -ge 10) { 'Medium' } elseif ($versionCount -gt 1) { 'Low' } else { 'Standardized' }
        $recommendedAction = switch ($riskPriority) {
            'High' { 'Prioritize packaging and migration to a smaller approved version set.' }
            'Medium' { 'Review deployment rings and reduce avoidable version spread.' }
            'Low' { 'Confirm the preferred version and monitor remaining secondary versions.' }
            default { 'Maintain the approved version baseline.' }
        }

        if ($versionCount -gt 1) { $fragmentedApplications++ }
        if ($versionCount -ge 10) { $highlyFragmentedApplications++ }

        foreach ($row in $group.Group) {
            $deviceCount = [int64]$row.DeviceCount
            $isDominant = ([string]$row.AppId -eq [string]$dominant.AppId)
            $outsideDominant = if ($isDominant) { [int64]0 } else { $deviceCount }
            $publisher = if ([string]::IsNullOrWhiteSpace([string]$row.AppPublisher)) { 'Unknown publisher' } else { ([string]$row.AppPublisher).Trim() }
            $publisherState = if ($publisher -match '(?i)^unknown') { 'Unknown' } else { 'Known' }
            $footprintState = if ($isDominant) { 'Dominant version' } elseif ($deviceCount -le 5) { 'Long tail version' } else { 'Secondary version' }

            $installationsOutsideDominant += $outsideDominant
            if ($footprintState -eq 'Long tail version') { $longTailVersions++ }
            if ($publisherState -eq 'Unknown') { $unknownPublisherVersions++ }

            if ($IncludeDetail) {
                $detail.Add([pscustomobject][ordered]@{
                    'Application Source ID' = [string]$row.AppId
                    'Application Name' = $applicationName
                    'Application Version' = if ([string]::IsNullOrWhiteSpace([string]$row.AppVersion)) { 'Unknown version' } else { ([string]$row.AppVersion).Trim() }
                    'Publisher' = $publisher
                    'Publisher State' = $publisherState
                    'Platform' = if ([string]::IsNullOrWhiteSpace([string]$row.Platform)) { 'Unknown' } else { ([string]$row.Platform).Trim().ToLowerInvariant() }
                    'Device Count' = $deviceCount
                    'Product Total Installs' = $productTotal
                    'Product Version Count' = $versionCount
                    'Dominant Version Share (%)' = Convert-ToInvariantDecimalText $dominantShare
                    'Installs Outside Dominant Version' = $outsideDominant
                    'Standardization State' = $standardizationState
                    'Version Footprint State' = $footprintState
                    'Risk Priority' = $riskPriority
                    'Recommended Action' = $recommendedAction
                    'Snapshot Date' = $SnapshotDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
                    'Snapshot Week' = $SnapshotWeek
                    'Evidence Status' = 'Observed'
                })
            }
        }
    }

    $totalInstallations = [int64](($Rows | Measure-Object DeviceCount -Sum).Sum)
    $publishers = @($Rows | ForEach-Object { [string]$_.AppPublisher } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique).Count
    $standardizationRate = if ($groups.Count -gt 0) { $fragmentedApplications / $groups.Count } else { $null }
    $dominantCoverage = if ($totalInstallations -gt 0) { 1 - ($installationsOutsideDominant / $totalInstallations) } else { $null }

    return [pscustomobject]@{
        Detail = @($detail)
        Trend = [pscustomobject][ordered]@{
            'Snapshot Date' = $SnapshotDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            'Date Key' = [int]$SnapshotDate.ToString('yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture)
            'Snapshot Week' = $SnapshotWeek
            'Applications' = $groups.Count
            'Application Versions' = $Rows.Count
            'Publishers' = $publishers
            'Application Install Observations' = $totalInstallations
            'Fragmented Applications' = $fragmentedApplications
            'Highly Fragmented Applications' = $highlyFragmentedApplications
            'Installations Outside Dominant Version' = $installationsOutsideDominant
            'Long Tail Versions' = $longTailVersions
            'Unknown Publisher Versions' = $unknownPublisherVersions
            'Standardization Opportunity Rate (%)' = Convert-ToInvariantDecimalText $standardizationRate
            'Dominant Version Coverage (%)' = Convert-ToInvariantDecimalText $dominantCoverage
            'Evidence Status' = 'Observed'
        }
    }
}

$currentPath = Join-Path $DataRoot 'DATA-LAST\Intune_DiscoveredApps_Summary.csv'
$historyRoot = Join-Path $DataRoot 'DATA-ALL\Intune\Applications\DiscoveredApps\WeeklyHistory'
if (-not (Test-Path -LiteralPath $currentPath -PathType Leaf)) { throw "Required source was not found: $currentPath" }
if (-not (Test-Path -LiteralPath $historyRoot -PathType Container)) { throw "Required history folder was not found: $historyRoot" }

$currentItem = Get-Item -LiteralPath $currentPath
$currentRows = @(Import-Csv -LiteralPath $currentPath)
$latestHistoryFolder = Get-ChildItem -LiteralPath $historyRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1
$currentWeek = if ($latestHistoryFolder) { $latestHistoryFolder.Name } else { 'Current' }
$currentProfile = Get-ApplicationProfile -Rows $currentRows -SnapshotDate $currentItem.LastWriteTime.Date -SnapshotWeek $currentWeek -IncludeDetail

$trend = foreach ($weekFolder in (Get-ChildItem -LiteralPath $historyRoot -Directory | Sort-Object Name)) {
    $summaryPath = Join-Path $weekFolder.FullName 'Intune_DiscoveredApps_Summary.csv'
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { continue }
    $summaryItem = Get-Item -LiteralPath $summaryPath
    $weeklyProfile = Get-ApplicationProfile -Rows @(Import-Csv -LiteralPath $summaryPath) -SnapshotDate $summaryItem.LastWriteTime.Date -SnapshotWeek $weekFolder.Name
    $weeklyProfile.Trend
}

foreach ($outputPath in @($InventoryOutputPath, $TrendOutputPath)) {
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
}

$currentProfile.Detail | Export-Csv -LiteralPath $InventoryOutputPath -NoTypeInformation -Encoding utf8NoBOM
$trend | Export-Csv -LiteralPath $TrendOutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject][ordered]@{
    InventoryOutputPath = $InventoryOutputPath
    InventoryRows = $currentProfile.Detail.Count
    TrendOutputPath = $TrendOutputPath
    TrendSnapshots = @($trend).Count
    Applications = $currentProfile.Trend.Applications
    ApplicationVersions = $currentProfile.Trend.'Application Versions'
    ApplicationInstallObservations = $currentProfile.Trend.'Application Install Observations'
    FragmentedApplications = $currentProfile.Trend.'Fragmented Applications'
    InstallationsOutsideDominantVersion = $currentProfile.Trend.'Installations Outside Dominant Version'
    DominantVersionCoverage = $currentProfile.Trend.'Dominant Version Coverage (%)'
} | ConvertTo-Json -Depth 3

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCdgeX1r+HpKWeg
# Dng6fmNMU3f/BzZuGF1kPIY5g+PA26CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIF4DIBdQvNFtMPaEXOqxU6UOWll+BpZrlWX75Cngbl+jMA0GCSqG
# SIb3DQEBAQUABIIBgJ71bvOW2X3NOKTaXDOPI7L+8E8aSY2uzwTqmLMtkm0tCLOB
# WKEEWKO6OoetOxHFB6RLqEmnPsAnaNfSzYMkt0wGNldKQLh2Ll17hD+gbNLQNL3x
# M7hPXPqgK/2Kckl6Dw9ZHIYx+Y6w0uWKkVaHbX2CrUGol2kExN4JI21Twcc0hYfF
# bJylFOieN6EIBZxk7qduxOYjmU+GfM1jn2PZ9vA8xO18vzoItbYFvXl+x9uK5C3z
# CwA9qme3g9CSn72pq/xmLsmATeVwaHfNGIhwkSC+pDSKzREuTxQq9UihrC4pp7cz
# VypAKmOXMQ8t5S/diNNnFlA96Zb7sbxBGVxbvP4uOhkSKmD+zWNWb/QxyLQP0UNf
# PtZ54w+4o6SUNTgeXp5yWrGFS3+56q1nmXCniJKYq/EK3enLSwsAMKeLvwz9Egef
# wMoWhCU+rjOMDC9Z3h6jvLM7NghpLibykkkjfs+5nL4G/gtPYeSaw2sK1lJj34Tv
# 1k4Zv6wtHbAZY6/9Z6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjkxOTEx
# MjhaMC8GCSqGSIb3DQEJBDEiBCBIH/c+RP7rmwFQ0etnjMrlKk+qAv6IzRD6xMlJ
# tSh2QzANBgkqhkiG9w0BAQEFAASCAgCikDPwWz3oBhWIHP/kDzEaq94UQKfz99x+
# KOrc1F6HHHwSS/RhgjGd3PBWAxu+cDxaNbvnPdn57Cw+meHXDI6OhhgvFXmqpQWs
# nEGkka/MFnlKRja5QSCC6Nus1WYJHim6Rws9CnZ/gqKdQLpn1AvK/Lm2fS6BKRc0
# SY+lk+PIVp1UZ3Sq5SLwGtDin72bo1dIO2vJP9XyTAwHaWmJwV38NdTgKgXfti0H
# 2cArsEGcvUpkmIOFbEP1qLLlJsqUyKXdMw1WnA0M7ie5iHHX84PeFyv1jvNwPU0j
# v/pzOFSezZsVOVuNdWdFejV7Tw7pA+t0k+vwWzC9aDkOd5VAmFbnZDoHxKnCYfQ3
# g3yfMhvzy5C/wYPU1or6gpr3baJVebff5JdTM4qp2DPnGpUvJKbI8SLMzB0nbtuh
# k9JyUFA3crMoph1B/zwInD3+tNjmx7Zst4qqYcc0fEE1LPYq4ZPs4GP9PjVLefBm
# VDfIX1v2SOOVqXNHEqEORD5YZJo2fkRHAVXbUHbdFpFR1eZZp7i4Ja1Xu6v9DdSb
# cvave95a77WOyPC2dExsOLPVpSZ5pbxtRLzeAOBRY++xx//T15Y2edfBsKXd/3vx
# NmI/3Bk0oJQk0Wj4Z+S78oYmrYhBzyHSfzLkyKqDx4bvAqiqn9BM76IqBTpIfQUW
# 38jCp6g4jQ==
# SIG # End signature block
