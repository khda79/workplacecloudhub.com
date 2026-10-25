#Requires -Version 7.0
# Version 1.0.0
[CmdletBinding()]
param()
# Synthetic fixtures only. No tenant initialization, collectors, production files or Power BI.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:checks = 0
function Check([bool]$Pass, [string]$Message) { $script:checks++; if (-not $Pass) { throw $Message } }
Import-Module (Join-Path $PSScriptRoot 'DeviceReadinessGapEvidence.psm1') -Force
$reference = [datetimeoffset]'2026-10-25T08:00:00Z'

# Reboot time: offset-less values are never converted to UTC.
$fact = ConvertTo-RebootFact -Value '2026-10-15 23:30:00' -Reference $reference
Check ($fact.Basis -eq 'Offset-less source time' -and $fact.Utc -eq '' -and $fact.Days -eq 10 -and $fact.AgeBasis -match 'offset unknown') 'Offset-less time was converted or miscounted'
$fact = ConvertTo-RebootFact -Value '2026-10-26 01:00:00' -Reference $reference
Check ($fact.Days -eq 0 -and $fact.AgeBasis -ne 'Future') 'A local date one day ahead of UTC is not in the future'
$fact = ConvertTo-RebootFact -Value '2026-10-27 00:00:00' -Reference $reference
Check ($fact.AgeBasis -eq 'Future' -and $null -eq $fact.Days) 'Future offset-less value produced an age'
$fact = ConvertTo-RebootFact -Value '2026-10-24T23:30:00-02:00' -Reference $reference
Check ($fact.Basis -eq 'Explicit offset' -and $fact.Utc -eq '2026-10-25T01:30:00Z' -and $fact.Days -eq 0) 'Explicit offset was not converted to UTC'
$fact = ConvertTo-RebootFact -Value '2026-10-15T10:00:00Z' -Reference $reference
Check ($fact.Days -eq 10 -and $fact.AgeBasis -eq 'UTC calendar days') 'UTC value miscounted'
$fact = ConvertTo-RebootFact -Value '2026-10-25T09:00:00Z' -Reference $reference
Check ($fact.AgeBasis -eq 'Future' -and $null -eq $fact.Days) 'Future UTC value produced an age'
foreach ($value in 'Error', '31/02/2026 10:00:00', 'not a date') {
    $fact = ConvertTo-RebootFact -Value $value -Reference $reference
    Check ($fact.Basis -eq 'Invalid' -and $null -eq $fact.Days -and $fact.Utc -eq '') "Invalid value '$value' produced a date"
}
$fact = ConvertTo-RebootFact -Value '' -Reference $reference
Check ($fact.Basis -eq 'Missing' -and $null -eq $fact.Days) 'Missing value produced a date'

# Facts: AD priority, LocalSystem only when AD has no value, AD computers without Intune kept.
$intune = @(
    [pscustomobject]@{TenantKey = 't'; 'Device ID' = 'i-ad'; 'Device name' = 'PC-AD'; 'Azure AD Device ID' = 'E-AD'; PhysicalMemoryGB = '7,5'}
    [pscustomobject]@{TenantKey = 't'; 'Device ID' = 'i-local'; 'Device name' = 'PC-LOCAL'; 'Azure AD Device ID' = 'e-local'; PhysicalMemoryGB = 'Error'}
    [pscustomobject]@{TenantKey = 't'; 'Device ID' = 'i-invalid'; 'Device name' = 'PC-INVALID'; 'Azure AD Device ID' = 'e-invalid'; PhysicalMemoryGB = '16'})
$ad = @(
    [pscustomobject]@{TenantKey = 't'; ObjectGUID = 'G-AD'; ObjectGuidNormalized = '{E-AD}'; IntuneDeviceId = 'i-ad'; Name = 'PC-AD$'; LastRebootDateTime = '2026-10-01 08:00:00'}
    [pscustomobject]@{TenantKey = 't'; ObjectGUID = 'g-local'; ObjectGuidNormalized = 'e-local'; IntuneDeviceId = ''; Name = 'PC-LOCAL$'; LastRebootDateTime = ''}
    [pscustomobject]@{TenantKey = 't'; ObjectGUID = 'g-invalid'; ObjectGuidNormalized = 'e-invalid'; IntuneDeviceId = 'i-invalid'; Name = 'PC-INVALID$'; LastRebootDateTime = 'Error'}
    [pscustomobject]@{TenantKey = 't'; ObjectGUID = 'g-adonly'; ObjectGuidNormalized = 'e-adonly'; IntuneDeviceId = ''; Name = 'PC-ADONLY$'; LastRebootDateTime = '2026-09-01 08:00:00'})
$entra = @(
    [pscustomobject]@{DeviceId = 'e-ad'; IsPending = 'False'}
    [pscustomobject]@{DeviceId = 'e-adonly'; IsPending = 'True'})
$local = @(
    [pscustomobject]@{AzureADDeviceId = 'e-ad'; FirmwareType = 'Bios'; SecureBootStatus = 'NotSupported'; 'Last Reboot Date' = '2026-10-24 08:00:00'}
    [pscustomobject]@{AzureADDeviceId = 'e-local'; FirmwareType = 'Uefi'; SecureBootStatus = 'Enabled'; 'Last Reboot Date' = '2026-10-20 08:00:00'}
    [pscustomobject]@{AzureADDeviceId = 'e-invalid'; FirmwareType = 'Uefi'; SecureBootStatus = 'Error'; 'Last Reboot Date' = '2026-10-20 08:00:00'})
$rows = @(New-DeviceReadinessGapEvidence -IntuneRows $intune -EntraRows $entra -AdRows $ad -LocalSystemRows $local -Reference $reference)
Check ($rows.Count -eq 4) "Expected three Intune devices and one AD-only computer, got $($rows.Count)"
$byName = @{}; foreach ($row in $rows) { $byName[$row.'Device Name'] = $row }
$row = $byName['PC-AD']
Check ($row.'Inventory Scope' -eq 'Intune device' -and $row.'AD Object GUID' -eq 'e-ad' -and $row.'Azure AD Device ID' -eq 'e-ad') 'Intune/AD/Entra keys not normalized'
Check ($row.'Last Reboot Source' -eq 'AD computer' -and $row.'Days Since Last Reboot' -eq '24') 'AD reboot value lost its priority'
Check ($row.'Physical Memory (GB)' -eq '7.5' -and $row.'Firmware Type (source)' -eq 'Bios' -and $row.'Secure Boot Status (source)' -eq 'NotSupported' -and $row.'Entra Registration Pending' -eq 'False') 'Raw device facts changed'
$row = $byName['PC-LOCAL']
Check ($row.'Last Reboot Source' -eq 'Intune local system' -and $row.'Days Since Last Reboot' -eq '5' -and $row.'AD Object GUID' -eq 'e-local') 'LocalSystem fallback or AD match by Entra ID failed'
Check ($row.'Physical Memory (GB)' -eq '' -and $row.'Entra Registration Pending' -eq '') 'Unknown memory or pending state was not left blank'
$row = $byName['PC-INVALID']
Check ($row.'Last Reboot Source' -eq 'AD computer' -and $row.'Last Reboot Time Basis' -eq 'Invalid' -and $row.'Days Since Last Reboot' -eq '') 'An invalid AD value must not fall back nor produce an age'
$row = $byName['PC-ADONLY']
Check ($row.'Inventory Scope' -eq 'AD computer only' -and $row.'Device Source ID' -eq '' -and $row.'Entra Registration Pending' -eq 'True' -and $row.'Days Since Last Reboot' -eq '54' -and $row.TenantKey -eq 't') 'AD-only computer facts missing'
Check (@($rows | ForEach-Object { $_.PSObject.Properties.Name -join ',' } | Sort-Object -Unique).Count -eq 1 -and (($rows[0].PSObject.Properties.Name) -join ',') -ceq ((Get-DeviceReadinessGapColumns) -join ',')) 'Row columns differ from the declared columns'
$contract = (Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/prepared-evidence-contract.json.txt') -Raw | ConvertFrom-Json).tables | Where-Object file -EQ 'DeviceReadinessGapEvidence.csv'
Check (($contract.columns.name -join ',') -ceq ((Get-DeviceReadinessGapColumns) -join ',')) 'Declared columns differ from the contract'

# Complete generator: the existing outputs are byte-identical with and without the new output.
$root = Join-Path ([IO.Path]::GetTempPath()) ('DeviceReadinessGap-' + [guid]::NewGuid().ToString('N'))
try {
    $last = Join-Path $root 'DATA-LAST'; New-Item -ItemType Directory -Path $last -Force | Out-Null
    $schema = (Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/source-schema.json') -Raw | ConvertFrom-Json).files
    function Write-Source([string]$Name, [hashtable[]]$Rows) {
        # Declared source columns plus the extra exported columns a fixture row needs.
        $columns = [Collections.Generic.List[string]]::new(); foreach ($column in @($schema.$Name.columns) + @($Rows | ForEach-Object { $_.Keys })) { if ($column -and -not $columns.Contains($column)) { $columns.Add($column) } }
        $objects = foreach ($values in $Rows) { $row = [ordered]@{}; foreach ($column in $columns) { $row[$column] = $(if ($values.ContainsKey($column)) { $values[$column] } else { '' }) }; [pscustomobject]$row }
        if (@($objects).Count) { @($objects) | Export-Csv -LiteralPath (Join-Path $last $Name) -NoTypeInformation -Encoding utf8NoBOM }
        else { [IO.File]::WriteAllText((Join-Path $last $Name), (($columns | ForEach-Object { '"' + $_ + '"' }) -join ',') + [Environment]::NewLine) }
    }
    Write-Source 'Intune_Devices_Inventory.csv' @(@{TenantKey = 'synthetic-test'; 'Device ID' = 'i-1'; 'Device name' = 'PC-1'; 'Azure AD Device ID' = 'e-1'; OS = 'Windows'; 'OS version' = '10.0.19045.1'; PhysicalMemoryGB = '8'}, @{TenantKey = 'synthetic-test'; 'Device ID' = 'i-2'; 'Device name' = 'PC-2'; 'Azure AD Device ID' = 'e-2'; OS = 'Windows'; 'OS version' = '10.0.26100.1'; PhysicalMemoryGB = '4'})
    Write-Source 'M365_Users_Active.csv' @()
    Write-Source 'M365_Entra_Devices.csv' @(@{TenantKey = 'synthetic-test'; DeviceId = 'e-1'; IsPending = 'False'}, @{TenantKey = 'synthetic-test'; DeviceId = 'e-3'; IsPending = 'True'})
    Write-Source 'AD_Computers_AllDomains.csv' @(@{TenantKey = 'synthetic-test'; ObjectGUID = 'g-1'; ObjectGuidNormalized = 'e-1'; IntuneDeviceId = 'i-1'; Name = 'PC-1$'; Enabled = 'True'; OperatingSystem = 'Windows 10 Enterprise'; LastRebootDateTime = '2026-10-01 08:00:00'}, @{TenantKey = 'synthetic-test'; ObjectGUID = 'g-3'; ObjectGuidNormalized = 'e-3'; Name = 'PC-3$'; Enabled = 'True'; OperatingSystem = 'Windows 10 Enterprise'})
    foreach ($name in 'Intune_Devices_UpgradeEligibility.csv', 'Intune_WindowsUpdate_Status.csv', 'Intune_EndpointAnalytics_DevicePerformance.csv', 'Intune_EndpointAnalytics_StartupDevices.csv', 'M365_Entra_Devices_HardwareIdConflicts.csv', 'M365_Entra_Devices_RemovalCandidates.csv') { Write-Source $name @() }
    Write-Source 'Intune_Devices_LocalSystem.csv' @(@{TenantKey = 'synthetic-test'; AzureADDeviceId = 'e-2'; DeviceId = 'i-2'; FirmwareType = 'Uefi'; SecureBootStatus = 'Enabled'; 'Last Reboot Date' = '2026-10-20 08:00:00'; PlatformScriptLastUpdate = '2026-10-24T08:00:00Z'})
    $generator = Join-Path $PSScriptRoot 'New-DeviceInventoryEvidence.ps1'
    $outputs = @{}
    foreach ($mode in 'without', 'with') {
        $out = Join-Path $root $mode; New-Item -ItemType Directory -Path $out -Force | Out-Null
        $arguments = @{DataRoot = $root; OutputPath = (Join-Path $out 'DeviceInventoryEvidence.csv'); SignalsOutputPath = (Join-Path $out 'DeviceRiskEvidenceSummary.csv'); DirectorySummaryOutputPath = (Join-Path $out 'DeviceDirectorySummary.csv'); SystemEvidenceOutputPath = (Join-Path $out 'DeviceSystemEvidence.csv'); AsOfUtc = $reference}
        if ($mode -eq 'with') { $arguments.ReadinessGapOutputPath = Join-Path $out 'DeviceReadinessGapEvidence.csv' }
        & $generator @arguments | Out-Null
        $outputs[$mode] = $out
    }
    foreach ($name in 'DeviceInventoryEvidence.csv', 'DeviceRiskEvidenceSummary.csv', 'DeviceDirectorySummary.csv', 'DeviceSystemEvidence.csv') {
        Check ((Get-FileHash (Join-Path $outputs.without $name)).Hash -eq (Get-FileHash (Join-Path $outputs.with $name)).Hash) "Existing output changed: $name"
    }
    Check (-not (Test-Path -LiteralPath (Join-Path $outputs.without 'DeviceReadinessGapEvidence.csv'))) 'Gap facts written without being requested'
    $gap = @(Import-Csv -LiteralPath (Join-Path $outputs.with 'DeviceReadinessGapEvidence.csv'))
    Check ($gap.Count -eq 3 -and @($gap | Where-Object 'Inventory Scope' -EQ 'AD computer only').Count -eq 1) 'Generator gap facts scope changed'
    Check ((@($gap | Where-Object 'Device Name' -EQ 'PC-3')[0]).'Entra Registration Pending' -eq 'True') 'AD-only pending registration lost'
    Check ((@($gap | Where-Object 'Device Name' -EQ 'PC-2')[0]).'Last Reboot Source' -eq 'Intune local system') 'Generator LocalSystem fallback lost'
    Check ((Get-Content -LiteralPath (Join-Path $outputs.with 'DeviceReadinessGapEvidence.csv') -TotalCount 1) -ceq (((Get-DeviceReadinessGapColumns) | ForEach-Object { '"' + $_ + '"' }) -join ',')) 'Generator header differs from the declared columns'
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host "PASS: $script:checks device readiness gap checks (reboot time basis, AD priority, AD-only scope, byte-identical existing outputs). Synthetic fixtures only."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBQo3W+qIh/sSX
# jdRjCb19BULKOSYmDmYNwdYr17TQpqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEZJVVBwJJNAOWWUmwTpeARuxFkfJ5CUJ3vnw6miIsDFMA0GCSqG
# SIb3DQEBAQUABIIBgJwIK4tQQFK34PV533wJPeSbogg6hNWJ2HN2xcqyC0cv63i5
# O7FgpdF81BMvY1CrZpYuoJ4KQXq0ZixnNxqERbxIlAn7uWAu8XPmaffhWqoMN/ew
# 6Yeg5+YoYUyOj64ItJEtfyTdEc0lqhjESs3PaGblohrXoPt+EhqT1w0+r8ZP341X
# 8pGTj2xys1364/495jccB2JPLJbP94iXQ2L2KvSkDMWc40X4a3EOp0dHeO1I/tgv
# KKMk8rhGkltnt5HhAv4YgMYw+K5c94qNx/N324djzCChpFMPya5qFPcsePjYeApI
# ohOb6jQmbfvP9hg68zuMFe+jhLlFQRWszDQ2iFNYt+7xUTSIeEeHF3m2dR0/90ix
# QVBZMyPpbadDpXOGfCO4Brq3t2DZgpfMA5K/p1Ca7ghLDoCXbAgNfTuiRpOaAu+o
# w0XrooAyHxJY/Fa5R1Y+iv067j/bHybJXdczksKT/vZLvwcLSbdMT/rWvbVhb8ys
# NeyiutUieLfZT8/6m6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMjU3
# NTVaMC8GCSqGSIb3DQEJBDEiBCAIq1A1lRYi5SZeE7F6K0EP6sgk12a65m58SyFB
# sP9ZBzANBgkqhkiG9w0BAQEFAASCAgAOOV0RCz0oDnrB0aPeBoukXL1Dq5r1CX21
# oj+XftrTZmF7FfgkleTzgqAok8zpb6BE3WrlKe15WQgGCCUSRpfuhpRrA8SV0IK0
# pS1exKh6PdPcdfRZSHWe/OF5v7WGJJVfIErXahsClPRYv9TK0dHOWdj7ubHHrbSH
# 0q1Bk/JZLcOvbi9gHhz2lO0ivofHJerYr5D3c2TaZ79p1HC5GKdATp4GbZwPnj/M
# ndkp32Z8KlEG+y+mgA1s6nmXZij4LFjSoPZ2CYuVTV/jItj7V6IPpNGrLTMqqyEy
# b7CS1/CAx47KeKgq8uR0lT+PtblgA3GOVHdXlM1o4hhvxm2Zu+/jBc589oD64Vtn
# FzM9JT4wTgdovKCa5uWTR96cQMXrl4PsuEm2wTwWm2JeU4i6GrKaRRvYlL25YooC
# uo9YNVZPjXD2db+mLbU2hew5nY10pO2W+DdzIBeJeumMXG6CwmmP0d7Yt6HxWmA2
# sHKx/sxnPc47y3HF/rOjzqsR/8r7zDc5FUTJJtywQToO/aFqBxaqhMUojOhOTYVt
# zqL5zOADf6M4i+cyDfJgJpRFwynXGYoUeAnbpDItn1IM+q2azLPOKmR+h2wWOsLT
# gxUKzWP4dXusC0YA3uTZ9eHmTfNFitiVOFGfao4IXJoHvv9lWFN6pymMqoJ/uieF
# 364zvcrQ9w==
# SIG # End signature block
