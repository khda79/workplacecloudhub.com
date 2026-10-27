# Version 1.1.0
# Device readiness gap facts (docs/HYBRID-IDENTITY-DEVICE-READINESS-FINDINGS.md): raw facts only,
# no verdict. Thresholds and applicability are evaluated downstream.
Set-StrictMode -Version Latest

function Get-GapText([AllowNull()][object]$Row, [string[]]$Names) {
    if ($null -eq $Row) { return '' }
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return ([string]$property.Value).Trim() }
    }
    return ''
}

function Get-GapKey([AllowNull()][object]$Value) { ([string]$Value).Trim().Trim('{', '}').ToLowerInvariant() }

function New-GapIndex([object[]]$Rows, [string[]]$Names) {
    # First row wins, as in the legacy producer.
    $index = @{}
    foreach ($row in $Rows) { $key = Get-GapKey (Get-GapText $row $Names); if ($key -and -not $index.ContainsKey($key)) { $index[$key] = $row } }
    return $index
}

function ConvertTo-GapNumber([AllowNull()][object]$Value) {
    $text = ([string]$Value).Trim() -replace ',', '.'
    $number = 0.0
    if ($text -and [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and -not [double]::IsNaN($number) -and -not [double]::IsInfinity($number)) { return $number }
    return $null
}

function ConvertTo-RebootFact {
    # Offset-less source times are never converted to UTC: the age uses whole calendar days
    # between the source date and the reference UTC date, and says so.
    param([string]$Value, [Parameter(Mandatory)][datetimeoffset]$Reference)
    $result = [ordered]@{Utc = ''; Basis = ''; Days = $null; AgeBasis = '' }
    if (-not $Value) { $result.Basis = 'Missing'; return [pscustomobject]$result }
    $referenceDate = $Reference.UtcDateTime.Date
    if ($Value -match '(Z|[+-]\d\d:\d\d)$') {
        $dto = [datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$dto)) {
            $result.Basis = 'Explicit offset'
            $result.Utc = $dto.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
            if ($dto -gt $Reference) { $result.AgeBasis = 'Future' } else { $result.Days = ($referenceDate - $dto.UtcDateTime.Date).Days; $result.AgeBasis = 'UTC calendar days' }
            return [pscustomobject]$result
        }
    } else {
        $local = [datetime]::MinValue
        $formats = [string[]]@('yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm:ss.FFFFFFF', 'dd/MM/yyyy HH:mm:ss', 'yyyy-MM-dd')
        if ([datetime]::TryParseExact($Value, $formats, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$local)) {
            $result.Basis = 'Offset-less source time'
            # A local date can lead UTC by up to one calendar day; anything later is in the future.
            if ($local.Date -gt $referenceDate.AddDays(1)) { $result.AgeBasis = 'Future' }
            else { $result.Days = [math]::Max(0, ($referenceDate - $local.Date).Days); $result.AgeBasis = 'Source calendar days (offset unknown, +/-1 day)' }
            return [pscustomobject]$result
        }
    }
    $result.Basis = 'Invalid'
    return [pscustomobject]$result
}

function Get-DeviceReadinessGapColumns {
    @('TenantKey', 'Inventory Scope', 'Device Source ID', 'Device Name', 'Azure AD Device ID', 'AD Object GUID',
      'Entra Registration Pending', 'Physical Memory (GB)', 'Firmware Type (source)', 'Secure Boot Status (source)',
      'Last Reboot (source)', 'Last Reboot Source', 'Last Reboot Time Basis', 'Last Reboot UTC',
      'Days Since Last Reboot', 'Reboot Age Basis', 'Evidence Date')
}

function New-DeviceReadinessGapEvidence {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$IntuneRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$EntraRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AdRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LocalSystemRows,
        [Parameter(Mandatory)][datetimeoffset]$Reference
    )
    $entraById = New-GapIndex $EntraRows @('DeviceId')
    $adByIntuneId = New-GapIndex $AdRows @('IntuneDeviceId')
    $adByEntraId = New-GapIndex $AdRows @('ObjectGuidNormalized', 'EntraDeviceIdNormalized', 'AzureADDeviceId')
    $localByEntraId = New-GapIndex $LocalSystemRows @('AzureADDeviceId', 'DeviceId')
    $evidenceDate = $Reference.UtcDateTime.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
    $matchedAd = [Collections.Generic.HashSet[object]]::new()
    $rows = [Collections.Generic.List[object]]::new()
    foreach ($device in $IntuneRows) {
        $deviceId = Get-GapKey (Get-GapText $device @('Device ID'))
        $entraId = Get-GapKey (Get-GapText $device @('Azure AD Device ID', 'Entra DeviceId'))
        $conflict=(Get-GapText $device @('Identity Status')) -eq 'Identity conflict'
        if($conflict){$entraId=''}
        $entraRow = if ($entraId -and $entraById.ContainsKey($entraId)) { $entraById[$entraId] } else { $null }
        $adRow = if (-not $conflict -and $deviceId -and $adByIntuneId.ContainsKey($deviceId)) { $adByIntuneId[$deviceId] } elseif (-not $conflict -and $entraId -and $adByEntraId.ContainsKey($entraId)) { $adByEntraId[$entraId] } else { $null }
        $localRow = if ($entraId -and $localByEntraId.ContainsKey($entraId)) { $localByEntraId[$entraId] } else { $null }
        if ($null -ne $adRow) { [void]$matchedAd.Add($adRow) }
        $rows.Add((New-GapRow $(if($conflict){'Intune device - Identity conflict'}else{'Intune device'}) $device $entraId $entraRow $adRow $localRow $Reference $evidenceDate))
    }
    # AD computers without an Intune device keep their directory facts (hybrid join pending, reboot).
    foreach ($adRow in $AdRows) {
        if ($matchedAd.Contains($adRow)) { continue }
        $entraId = Get-GapKey (Get-GapText $adRow @('ObjectGuidNormalized', 'EntraDeviceIdNormalized', 'AzureADDeviceId'))
        $conflict=(Get-GapText $device @('Identity Status')) -eq 'Identity conflict'
        if($conflict){$entraId=''}
        $entraRow = if ($entraId -and $entraById.ContainsKey($entraId)) { $entraById[$entraId] } else { $null }
        $localRow = if ($entraId -and $localByEntraId.ContainsKey($entraId)) { $localByEntraId[$entraId] } else { $null }
        $rows.Add((New-GapRow 'AD computer only' $null $entraId $entraRow $adRow $localRow $Reference $evidenceDate))
    }
    return $rows.ToArray()
}

function New-GapRow {
    param([string]$Scope, [AllowNull()][object]$Device, [string]$EntraId, [AllowNull()][object]$EntraRow, [AllowNull()][object]$AdRow, [AllowNull()][object]$LocalRow, [datetimeoffset]$Reference, [string]$EvidenceDate)
    $entraRow = $EntraRow; $adRow = $AdRow; $localRow = $LocalRow; $entraId = $EntraId; $device = $Device
    # Legacy priority: the AD computer value, then the Intune local system value only when AD has none.
    $rebootRaw = Get-GapText $adRow @('LastRebootDateTime', 'LastRebootDate', 'LastBootUpTime')
    $rebootSource = if ($rebootRaw) { 'AD computer' } else { '' }
    if (-not $rebootRaw) { $rebootRaw = Get-GapText $localRow @('Last Reboot Date', 'LastRebootDate', 'LastBootUpTime'); if ($rebootRaw) { $rebootSource = 'Intune local system' } }
    $reboot = ConvertTo-RebootFact -Value $rebootRaw -Reference $Reference
    $pending = Get-GapText $entraRow @('IsPending')
    $memory = ConvertTo-GapNumber (Get-GapText $device @('PhysicalMemoryGB'))
    $tenant = Get-GapText $device @('TenantKey'); if (-not $tenant) { $tenant = Get-GapText $adRow @('TenantKey') }
    $name = Get-GapText $device @('Device name'); if (-not $name) { $name = (Get-GapText $adRow @('Name', 'SamAccountName')).TrimEnd('$') }
    [pscustomobject][ordered]@{
        'TenantKey' = $tenant
        'Inventory Scope' = $Scope
        'Device Source ID' = Get-GapText $device @('Device ID')
        'Device Name' = $name
        'Azure AD Device ID' = $entraId
        'AD Object GUID' = Get-GapKey (Get-GapText $adRow @('ObjectGuidNormalized', 'ObjectGUID'))
        'Entra Registration Pending' = if (-not $pending) { '' } elseif ($pending -match '^(?i:true|1|yes)$') { 'True' } elseif ($pending -match '^(?i:false|0|no)$') { 'False' } else { '' }
        'Physical Memory (GB)' = if ($null -ne $memory) { $memory.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) } else { '' }
        'Firmware Type (source)' = Get-GapText $localRow @('FirmwareType')
        'Secure Boot Status (source)' = Get-GapText $localRow @('SecureBootStatus')
        'Last Reboot (source)' = $rebootRaw
        'Last Reboot Source' = $rebootSource
        'Last Reboot Time Basis' = $reboot.Basis
        'Last Reboot UTC' = $reboot.Utc
        'Days Since Last Reboot' = if ($null -ne $reboot.Days) { [string]$reboot.Days } else { '' }
        'Reboot Age Basis' = $reboot.AgeBasis
        'Evidence Date' = $EvidenceDate
    }
}

Export-ModuleMember -Function New-DeviceReadinessGapEvidence, Get-DeviceReadinessGapColumns, ConvertTo-RebootFact

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJhbsAzMGWkGia
# bV8ogRyJFLE/In4unH/imIvipa+bB6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOsr5R/+/x8/JzbuZ5x7n1wNHz7fjGXmNHmRLr/RoL1uMA0GCSqG
# SIb3DQEBAQUABIIBgEDz6qg16G9sJOjZJJbno7Eq9qEtgtDu7Dpe2qLaAtvD2rY3
# XxW3y4aV7JPKbDgUYhWlZkyYRBSeTt/MiqinYByxdWcsqAaVouCYtge3sngoZqMl
# GTF7Rkdi/gHSG+MS1NK5hWQ72v1fuLbt5vqNik6m3IFRNzUjbEt3b+4Kc55ehbpq
# WOuPTZXcLEzu69xogH73jADBFUEPaaZxuEQEWamEi+4V3QeZL69W1yLJT7DH5DV6
# pffptVIRiyS+u/v5FH1iMaQmRqkUcwIij/Np3ntmmVS/HMWcBF+Ne1AsCnURAA7+
# Iu4beWHAh8tKJLHlX5mIvzX9NPjPPVUgBw2DFkrh09ZMNvXmsHAbbHw8c7t4mwpp
# Sp+w5Tk7Sss00rkef/BavngE/LA85fV0qCxyPh3VnZq3ZVHbDiUkz/eJqf3hBEuF
# CanEVum/w1vc9P91XdrQu7ApUILuni+1EdlfplWFYXaFFgNTyKW18VZVk11YKnD1
# FQHibuLEr/aDYAO/YKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMzIz
# NTNaMC8GCSqGSIb3DQEJBDEiBCCcj5Xoa5+RmGgy4zKQXIBEAqNH/YgAvWSuAOim
# WTBrdTANBgkqhkiG9w0BAQEFAASCAgANC2yHkqJezr+e7PtsmifRGEtDVRsQzJNc
# J0Q7CzylYvu/QoMyoREMAJ9Il4VWRZniGhEqZGwRofBy95kOVHjwKnzsM5ij8Qps
# UAWdYD8df1rNXD1SXyW9kocG8zzwb9NDuQE4xKCcJ6syiUoVSZB/K6mFd1o7UiuS
# mJRIk+ywoC1k8v2w961upMC2vRaUBcjWmKcoUkpLF7ax0ED/Px+FEmlY7Htaux0i
# i1WmHqwizrbrGzEms4DaOIl+ljRA9EzMYZ8CNohuAM/oLoeNS00/uwzOnIlKFkRE
# Uzm4uqeYejNKBj6tjRMUS+qDvvd9KDXBtA3P5xf+2KBZ/waBu6hXwZWUWlvCNV6v
# Ts5/cSuDRpAxWc4GDHKmc23DVOxLCgmy9orm1cGsct/ruvhSBSakEhuTUSslDKhj
# OPZdT+Vt79ZP9Wt+5ZdBzf+/iYYB4XnPYoDWx7wsaLKhB42yKQkWR15BsVG6KFk5
# Anml9+EsPcKcrXQAcodBsFsxyscNOIr+ysfwLWH74t1nsD2YqQdoocmrL9HGuPxs
# bfkK1DvDOL5LwJUKt65Be8aATHRS5zAzIPQ2lE6k/ctXKmKxFGxi1E0DhhrUznOQ
# vNjiCZtfLqCgcY+I350pPFVXNz5QqesMjHgCcDLeuuoVM3uTbd3VbRD8icW+1bdz
# taMqM3kkjg==
# SIG # End signature block
