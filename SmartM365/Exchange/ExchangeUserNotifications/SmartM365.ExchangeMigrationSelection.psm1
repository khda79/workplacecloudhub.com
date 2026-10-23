<#
.SYNOPSIS
Read-only migration selection normalization and recipient export.
.VERSION
1.0.0
#>
Set-StrictMode -Version 2.0
function Get-ExoSelectionValue {
    param($Object, [string]$Name, $Default = '')
    if ($null -eq $Object) { return $Default }
    if ($Object -is [Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
    }
    elseif ($Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $Default
}
function Test-ExoSelectionProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $false }
    if ($Object -is [Collections.IDictionary]) { return $Object.Contains($Name) }
    return $null -ne $Object.PSObject.Properties[$Name]
}
function Get-ExoSelectionSchedule {
    param($Object, [string]$Source, [string]$DisplayTimeZoneId)
    $result = [ordered]@{Known=$false;EffectiveDate='';ScheduledLocal='';ScheduledUtc='';DateSource=$Source;ScheduleReason='No scheduled completion date'}
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById($DisplayTimeZoneId)
    $utcValue = Get-ExoSelectionValue $Object 'CompleteAfterUTC' $null
    $hasUtcValue = -not [string]::IsNullOrWhiteSpace([string]$utcValue)
    $value = if ($hasUtcValue) { $utcValue } else { Get-ExoSelectionValue $Object 'CompleteAfter' $null }
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return [pscustomobject]$result }
    try {
        $utc = $null
        if ($value -is [datetimeoffset]) { $utc = $value.UtcDateTime }
        elseif ($value -is [datetime] -and $value.Kind -ne [DateTimeKind]::Unspecified) { $utc = $value.ToUniversalTime() }
        elseif ($hasUtcValue) {
            $utc = [datetime]::SpecifyKind([datetime]::Parse([string]$value, [Globalization.CultureInfo]::InvariantCulture), [DateTimeKind]::Utc)
        }
        elseif ([string]$value -match '(Z|[+-]\d{2}:\d{2})$') {
            $utc = [datetimeoffset]::Parse([string]$value, [Globalization.CultureInfo]::InvariantCulture).UtcDateTime
        }
        else {
            $sourceZoneId = [string](Get-ExoSelectionValue $Object 'TimeZone' '')
            if (-not $sourceZoneId) { throw 'Completion time has no verified time zone' }
            $sourceZone = [TimeZoneInfo]::FindSystemTimeZoneById($sourceZoneId)
            $wallTime = [datetime]::SpecifyKind([datetime]::Parse([string]$value, [Globalization.CultureInfo]::InvariantCulture), [DateTimeKind]::Unspecified)
            if ($sourceZone.IsAmbiguousTime($wallTime) -or $sourceZone.IsInvalidTime($wallTime)) { throw 'Ambiguous or invalid daylight-saving time' }
            $utc = [TimeZoneInfo]::ConvertTimeToUtc($wallTime, $sourceZone)
        }
        if ($utc.Year -ge 9998) { throw 'Completion is not scheduled' }
        $local = [TimeZoneInfo]::ConvertTimeFromUtc($utc, $zone)
        $result.Known = $true
        $result.EffectiveDate = $local.ToString('yyyy-MM-dd')
        $result.ScheduledLocal = $local.ToString('yyyy-MM-dd HH:mm') + ' [' + $DisplayTimeZoneId + ']'
        $result.ScheduledUtc = $utc.ToString('o')
        $result.ScheduleReason = 'Scheduled earliest completion; actual completion may be later'
    }
    catch { $result.ScheduleReason = $_.Exception.Message }
    return [pscustomobject]$result
}
function ConvertTo-ExoSelectionRecipient {
    param([object[]]$Records, [string]$DisplayTimeZoneId, [hashtable]$SentRegistry = @{})
    $rows = [Collections.Generic.List[object]]::new()
    foreach ($record in $Records) {
        $user = $record.User; $batch = $record.Batch; $stats = $record.Statistics
        $email = ([string](Get-ExoSelectionValue $user 'EmailAddress')).Trim().ToLowerInvariant()
        if (-not $email -or $email -notmatch '^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$') { throw 'Migration source contains a missing or invalid email address' }
        $hasUserSchedule = (Test-ExoSelectionProperty $stats 'CompleteAfterUTC') -or (Test-ExoSelectionProperty $stats 'CompleteAfter')
        $schedule = Get-ExoSelectionSchedule $stats 'User CompleteAfter' $DisplayTimeZoneId
        if (-not $schedule.Known -and $hasUserSchedule -and
            -not (Get-ExoSelectionValue $stats 'CompleteAfterUTC') -and -not (Get-ExoSelectionValue $stats 'CompleteAfter')) {
            $schedule = Get-ExoSelectionSchedule $batch 'Batch CompleteAfter' $DisplayTimeZoneId
        }
        if (-not $hasUserSchedule) {
            $schedule = Get-ExoSelectionSchedule $batch 'Batch proposal (individual override unverified)' $DisplayTimeZoneId
            $schedule.Known = $false
            $schedule.ScheduleReason = 'Individual schedule unavailable; explicitly confirm or replace the proposed date'
        }
        $status = [string](Get-ExoSelectionValue $user 'Status')
        $excluded = if ($status -match '^Completed') { 'Migration completed' } else { '' }
        if ($SentRegistry.ContainsKey($email)) { $excluded = 'Already notified (Force send required)' }
        $rows.Add([pscustomobject]@{
            Selected=($excluded -eq '' -and $schedule.Known -and $status -notmatch 'Failed|Stopped|Corrupt')
            EmailAddress=$email;BatchIds=[string](Get-ExoSelectionValue $batch 'Identity');MigrationStatus=$status
            EffectiveDate=$schedule.EffectiveDate;DateConfirmed=[bool]$schedule.Known;DateSource=$schedule.DateSource
            ScheduledLocal=$schedule.ScheduledLocal;ScheduledUtc=$schedule.ScheduledUtc;ScheduleReason=$schedule.ScheduleReason
            AlreadySent=$SentRegistry.ContainsKey($email);Completed=($status -match '^Completed');Conflict=$false;Exclusion=$excluded
        })
    }
    foreach ($group in ($rows | Group-Object EmailAddress)) {
        $row = $group.Group[0]
        $row.BatchIds = (@($group.Group.BatchIds | Sort-Object -Unique) -join ';')
        if ($group.Count -gt 1) {
            $row.AlreadySent = @($group.Group | Where-Object AlreadySent).Count -gt 0
            $row.Completed = @($group.Group | Where-Object Completed).Count -gt 0
            if (@($group.Group.EffectiveDate | Sort-Object -Unique).Count -gt 1 -or @($group.Group | Where-Object {-not $_.DateConfirmed}).Count) {
                $row.Selected=$false; $row.Conflict=$true; $row.DateConfirmed=$false
                $row.ScheduleReason='Conflicting or unverified dates across batches; an explicit date is required'
            }
            if ($row.Completed) { $row.Selected=$false; $row.Exclusion='Completed migration appears in selected batches' }
        }
        $row
    }
}
function Assert-ExoMigrationSelection {
    param([object[]]$Rows, [string]$TenantId, [switch]$ForceSend)
    $tenantGuid = [guid]::Empty
    if (-not [guid]::TryParse($TenantId, [ref]$tenantGuid) -or $tenantGuid -eq [guid]::Empty) { throw 'A verified tenant GUID is required' }
    if (-not $Rows.Count) { throw 'Select at least one recipient' }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $Rows) {
        $email = [string]$row.EmailAddress
        if ($email -notmatch '^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$' -or -not $seen.Add($email)) { throw 'Invalid or duplicated selected recipient' }
        $date = [datetime]::MinValue
        if (-not [datetime]::TryParseExact([string]$row.EffectiveDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) { throw "A valid explicit date is required for $email" }
        if (-not $row.DateConfirmed) { throw "Confirm the scheduled date for $email" }
        if ($row.Completed) { throw "Completed migration cannot be notified: $email" }
        if ($row.Conflict -and $row.DateSource -ne 'Operator override') { throw "Resolve conflicting dates for $email" }
        if ($row.AlreadySent -and -not $ForceSend) { throw "Recipient already notified: $email" }
        if ($date.Date -lt (Get-Date).Date -and -not $ForceSend) { throw "Date is in the past for $email" }
        if ([string]::IsNullOrWhiteSpace([string]$row.BatchIds)) { throw 'Missing source batch identity' }
    }
}
function Export-ExoMigrationSelection {
    param([object[]]$Rows, [string]$TenantId, [string]$TenantKey, [string]$Path, [string]$SourceReadAtUtc, [switch]$ForceSend)
    Assert-ExoMigrationSelection -Rows $Rows -TenantId $TenantId -ForceSend:$ForceSend
    $captured = [datetime]::UtcNow.ToString('o')
    $export = @(foreach ($row in $Rows) {
        [pscustomobject]@{EmailAddress=$row.EmailAddress;EffectiveDate=$row.EffectiveDate;SelectionTenantId=$TenantId;TenantKey=$TenantKey
            SelectionSource='ExchangeOnlineMigrationBatch';BatchIds=$row.BatchIds;MigrationStatus=$row.MigrationStatus
            DateSource=$row.DateSource;ScheduledUtc=$row.ScheduledUtc;ScheduledLocal=$row.ScheduledLocal
            ScheduleReason=$row.ScheduleReason;SelectedAtUtc=$captured;SourceReadAtUtc=$SourceReadAtUtc;DateConfirmed=$true}
    })
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    $export | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    return $export
}
function Assert-ExoMigrationSelectionCsv {
    param([object[]]$Rows, [string]$TenantId, [string]$TenantKey='')
    if (-not $Rows.Count) { throw 'Migration selection is empty' }
    $tenantGuid = [guid]::Empty
    if (-not [guid]::TryParse($TenantId, [ref]$tenantGuid) -or $tenantGuid -eq [guid]::Empty) { throw 'Configured tenant GUID is required for migration selection' }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $Rows) {
        if ([string](Get-ExoSelectionValue $row 'SelectionSource') -ne 'ExchangeOnlineMigrationBatch' -or
            [string](Get-ExoSelectionValue $row 'SelectionTenantId') -ine $TenantId -or
            [string](Get-ExoSelectionValue $row 'DateConfirmed') -ine 'True' -or
            -not [string](Get-ExoSelectionValue $row 'BatchIds')) { throw 'Invalid migration selection provenance or tenant' }
        if ($TenantKey -and [string](Get-ExoSelectionValue $row 'TenantKey') -cne $TenantKey) { throw 'Migration selection tenant profile changed' }
        foreach ($field in @('SourceReadAtUtc','SelectedAtUtc')) {
            $stamp=[datetimeoffset]::MinValue
            if (-not [datetimeoffset]::TryParse([string](Get-ExoSelectionValue $row $field),[ref]$stamp) -or
                ([datetimeoffset]::UtcNow-$stamp).TotalMinutes -gt 15 -or ($stamp-[datetimeoffset]::UtcNow).TotalMinutes -gt 1) { throw 'Migration selection expired or has invalid timestamps; reload recipients' }
        }
        $date = [datetime]::MinValue
        if (-not [datetime]::TryParseExact([string](Get-ExoSelectionValue $row 'EffectiveDate'), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) { throw 'Migration selection requires a date for every recipient' }
        $email = [string](Get-ExoSelectionValue $row 'EmailAddress')
        if ($email -notmatch '^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$' -or -not $seen.Add($email)) { throw 'Invalid or duplicated migration selection email' }
        if ([string](Get-ExoSelectionValue $row 'MigrationStatus') -match '^Completed') { throw 'Selection contains a completed migration' }
    }
}
Export-ModuleMember -Function Get-ExoSelectionValue, Test-ExoSelectionProperty, Get-ExoSelectionSchedule, ConvertTo-ExoSelectionRecipient, Assert-ExoMigrationSelection, Export-ExoMigrationSelection, Assert-ExoMigrationSelectionCsv

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCzqtnC9bs5lkJ6
# hUaE1sMSemPoQ/SpZDgBdI9gkCgU3qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILTe0fZRGkxJfoF1o0Z0DtDKbTDhZ9NbmmKVNMrUMhhiMA0GCSqG
# SIb3DQEBAQUABIIBgDgKukNN+kufNJBPqBFCdjYIDB+v48l+z29BHjNeA/LS1i00
# 9JX7wB/ohLzF/EfHS9PSDXNHbmxQej2lYVCyqHN+dwvq+8TtCgx2ziO89NhxufqB
# FSXn4fkvJc1YB5S7v0zTx0Y08cmKBgDj1HDJkEtPkUuXGmuhSdO2XyYDUMfPEptX
# 4UJdADmT1Gun9tgJcUxisvtTg2FQ1Agyy6eB5WZKffnCy+NEQBcUOVdKn5M3ZAJw
# j5x5hVlgQ6q6y+rTRQ54itXubuxpxqQjR5BKSyCDOdp+0c0zdX+PAy6aoePuy14h
# jXt2P5+qQnqTiPerdJfX4BE79GsW8IV777eJ8tgOuyBYUn/d8A/xnqwPBP6JPLJu
# lKpCioZmTsP4I7pr7L/TjK/QZyIPgJW/MNqIG/vpI2CuTS9yrlVHJCIS+b29d5zR
# 16tYj0zZWUSVADnAeCEJr0HkVC2lnva9tPkt/6Phvw8k2vUpfYi4dhSerItHYFZ8
# zhQKMn8w2FfbVSp+gqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzUz
# NTRaMC8GCSqGSIb3DQEJBDEiBCDOf9A8QaFUVg3bUl7tHA04luQsbD8MevzhRuB8
# wczfYjANBgkqhkiG9w0BAQEFAASCAgB5Z7vDv8xA65kgofZZzNGVR2JsYVj+V18u
# ztORS56+4tiNLy5c/iFJEXIaLBSnXds+ZRvyMoOQtJoOaWo/7XTVQfiCNetzT9Mk
# I/UUMZRR3ZXPZemRVpHNLwTHre8qWOKTkBgXqoLuy0sJ6tVpRk5zkLyjwJH9pFLf
# nTvrj9LwtFoMMlKAJZp4Z1H3FyB+wNe6iuv5J7g7bV6//Obq3MOy82nSU+hICpbu
# NLxQ6uRuaeHs52KBLl5aoffPR9gNXIPsEFqybzpgqeyFCI0JP/xqXmgK64ODJ2Cx
# xh6gW/lpr++cWnK4sK8DW5IPUZbL1qw/vemcMajFVZdpA3OZA3Hs6ux4XVqy8ZYO
# RkzXySYBcgqIPtWYO3UJprVhqg2BswFAo7n9A+WA7tlviDnbgOp/Eo3ueHnT7H27
# 0RiQ9OrF6TyyOKktzFi2OoEdYofwRekEdIE5ygCWHRMvGI1e/HYaKleSlso33w3D
# rEys7qfoXYo12L2+e9O3aStGHBI9mY1TtWuQrR8lZMpS7WF6LjdBZoRM2DO49opB
# +EBOJUl8goUypN2GYbeDI+t7o5dmRE6M3TFDhKpM7zZNCn8H02SFgTrlzHo/4GVl
# swEcBUbz/VT2+ceKoQ69PdvN55bNQJRgiHn4T444kKdA7UVelquzAcWUBNz9cKs4
# 7Rlht1JsJg==
# SIG # End signature block
