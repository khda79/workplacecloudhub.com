# Pure source-date qualification, version 1.0.2. No runtime, network or data writes.
function Get-IdentityProperty {
    param([AllowNull()][object]$Row, [string]$Name)
    if ($null -eq $Row) { return $null }
    if ($Row -is [System.Collections.IDictionary]) {
        if ($Row.Contains($Name)) { return [pscustomobject]@{Value=$Row[$Name]} }
        foreach ($key in $Row.Keys) {
            if ([string]::Equals([string]$key,$Name,[StringComparison]::OrdinalIgnoreCase)) {
                return [pscustomobject]@{Value=$Row[$key]}
            }
        }
        return $null
    }
    return $Row.PSObject.Properties[$Name]
}

function ConvertFrom-IdentityIsoEvidence {
    param([string]$Iso, [datetimeoffset]$Now)
    $utc=$null; $status='Invalid'
    if ($Iso -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$') {
        $date=[datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse($Iso,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$date)) {
            if ($date.Year -lt 1900) { $status='Sentinel' }
            elseif ($date -gt $Now) { $status='Future' }
            else { $status='Qualified'; $utc=$date.ToUniversalTime() }
        }
    } elseif ($Iso -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?$') {
        $date=[datetime]::MinValue
        if ([datetime]::TryParseExact($Iso,[string[]]@('yyyy-MM-ddTHH:mm:ss','yyyy-MM-ddTHH:mm:ss.FFFFFFF'),[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$date)) {
            $status=if($date.Year -lt 1900){'Sentinel'}else{'Timezone not established'}
        }
    }
    [pscustomobject]@{Utc=$utc;Status=$status}
}

function ConvertTo-IdentityNativeIso {
    param([AllowNull()][object]$Value)
    if ($Value -is [datetimeoffset]) { return $Value.ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [datetime]) { return $Value.ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    # A formatted/deserialized string is not evidence of its original Kind.
    return ''
}

function Read-IdentityDate {
    param([AllowNull()][object]$Row, [string]$Field, [string]$FileTimeField,
        [datetimeoffset]$Now = [datetimeoffset]::UtcNow)
    $p=Get-IdentityProperty $Row $Field; $raw=if($p){[string]$p.Value}else{''}
    $p=Get-IdentityProperty $Row ($Field+'Iso8601'); $iso=if($p){[string]$p.Value}else{''}
    $p=if($FileTimeField){Get-IdentityProperty $Row $FileTimeField}else{$null}
    $ticks=if($p){([string]$p.Value).Trim()}else{''}
    $status='Absent'; $proof='None'; $utc=$null
    # Supplied FILETIME is authoritative evidence, never a failed attempt followed
    # by a fallback to another column. Contradictory columns fail closed.
    if ($ticks -ne '') {
        $proof='AD FILETIME'; $status='Invalid'; $ft=$null
        if ($ticks -eq '0') { $status='Sentinel' }
        elseif ($ticks -match '^[0-9]+$') {
            try {
                $ft=[datetimeoffset]([datetime]::FromFileTimeUtc([long]$ticks))
                $status=if($ft.Year -lt 1900){'Sentinel'}elseif($ft -gt $Now){'Future'}else{'Qualified'}
            } catch { $status='Invalid' }
        }
        if ($iso -and $status -eq 'Qualified') {
            $other=ConvertFrom-IdentityIsoEvidence $iso $Now
            if ($other.Status -ne 'Qualified' -or $other.Utc -ne $ft) { $status='Conflicting source evidence' }
        }
        if ($status -eq 'Qualified') { $utc=$ft }
    } elseif ($iso) {
        $proof='Native producer round-trip'
        $other=ConvertFrom-IdentityIsoEvidence $iso $Now
        $utc=$other.Utc; $status=$other.Status
    } elseif ($raw) {
        $status=if($raw -match '^(?:\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4})(?:[T ]\d{2}:\d{2}:\d{2})'){'Unqualified legacy value'}else{'Invalid'}
    }
    [pscustomobject]@{Raw=$raw;IsoRaw=$iso;FileTimeRaw=$ticks;Utc=$utc;Status=$status;Proof=$proof}
}

function Merge-IdentityActivity {
    param([object[]]$Evidence, [int]$Days, [datetimeoffset]$Now = [datetimeoffset]::UtcNow)
    $excluded = @($Evidence | Where-Object { $_.Status -eq 'Not applicable' })
    $applicable = @($Evidence | Where-Object { $_.Status -ne 'Not applicable' })
    $qualified = @($applicable | Where-Object { $_.Status -eq 'Qualified' })
    $complete = $qualified.Count -eq $applicable.Count -and $applicable.Count -gt 0
    $latest = $null
    foreach ($item in $qualified) { if ($null -eq $latest -or $item.Utc -gt $latest) { $latest = $item.Utc } }
    $recent = $null; $old = $null
    if ($null -ne $latest -and $latest.UtcDateTime.Date -ge $Now.UtcDateTime.Date.AddDays(-$Days)) { $recent=$true; $old=$false }
    elseif ($complete) { $recent=$false; $old=$true }
    [pscustomobject]@{Utc=$latest; Recent=$recent; Old=$old; Coverage=$(if($complete){'Complete'}else{'Partial'});
        Status=$(if($recent -eq $true){'Recent observed activity'}elseif($old -eq $true){'Older observed activity'}else{'Unknown'});
        ApplicableCount=$applicable.Count; ExcludedCount=$excluded.Count;
        Exclusions=($excluded | ForEach-Object { $_.Proof }) -join ';';
        ConclusionScope='Applicable observed directory/mailbox/device signals only; not universal inactivity'}
}

function Get-IdentityMailboxActivitySet {
    param($AdRow, $LocalMailbox, $CloudMailbox, $CloudStats, $RemoteMailbox)
    $local = Read-IdentityDate $LocalMailbox LastLogonTime
    # A positively identified remote recipient has no local mailbox logon signal.
    # A missing mailbox row alone never establishes non-applicability.
    $remoteType = Get-IdentityProperty $RemoteMailbox 'RecipientTypeDetails'
    $adId = Get-IdentityProperty $AdRow 'ObjectGUID'
    $remoteId = Get-IdentityProperty $RemoteMailbox 'ObjectGUID'
    $sameIdentity = $adId -and $remoteId -and $adId.Value -and ([string]$adId.Value -eq [string]$remoteId.Value)
    if ($sameIdentity -and $remoteType -and [string]$remoteType.Value -match '^Remote(?:User|Shared|Room|Equipment)Mailbox$') {
        if ($null -eq $LocalMailbox) {
            $local = [pscustomobject]@{Utc=$null;Status='Not applicable';Proof='Native remote recipient type: local mailbox logon excluded';Raw=''}
        } else {
            $local = [pscustomobject]@{Utc=$null;Status='Ambiguous identity/scope';Proof='Local mailbox and remote recipient both matched';Raw=''}
        }
    }
    # Mailbox detail and statistics are alternative exports of the same cloud
    # logon signal, not independent required platforms. Absence of an alternate
    # export must not create a second expected source.
    $cloud = Read-IdentityDate $CloudMailbox LastLogonTime
    $stats = Read-IdentityDate $CloudStats LastLogonTime
    if ($cloud.Status -eq 'Absent') { $cloud = $stats }
    elseif ($stats.Status -ne 'Absent') {
        $merged = Merge-IdentityActivity @($cloud,$stats) 180
        if ($merged.Coverage -eq 'Complete') {
            $cloud = [pscustomobject]@{Utc=$merged.Utc;Status='Qualified';Proof='Cloud logon detail/statistics';Raw=''}
        } else {
            # Retain a recent positive qualified observation but keep incomplete
            # alternate evidence as a separate uncertainty in the same signal.
            return @((Read-IdentityDate $AdRow LastLogonDate LastLogonTimestamp),$local,$cloud,$stats)
        }
    }
    return @((Read-IdentityDate $AdRow LastLogonDate LastLogonTimestamp),$local,$cloud)
}

function New-IdentityRemoteScopeIndex {
    param([object[]]$Rows)
    $index=@{}
    foreach($row in $Rows) {
        $p=Get-IdentityProperty $row 'ObjectGUID'; if(-not $p -or -not $p.Value){continue}
        $key=([string]$p.Value).Trim().ToLowerInvariant()
        # Duplicate native identities cannot establish an exclusion, even if a
        # legacy SMTP/UPN map would keep the first row.
        if($index.ContainsKey($key)){$index[$key]=$null}else{$index[$key]=$row}
    }
    return $index
}

function Get-IdentityDeviceActivitySet {
    param($AdRow, $NativeEntra, $NativeManaged)
    # There is no positive exclusion proof in these exports: missing/ambiguous
    # native identity remains expected-but-unknown, never non-applicable.
    @((Read-IdentityDate $AdRow LastLogonDate LastLogonTimestamp),
      (Read-IdentityDate $NativeEntra ApproximateLastSignInDateTime),
      (Read-IdentityDate $NativeManaged LastSyncDateTime))
}

function Resolve-IdentityWindowsUpdateRows {
    param([object[]]$Rows)
    $rowsPresent=@($Rows | Where-Object { $null -ne $_ })
    if (-not $rowsPresent.Count) { return [pscustomobject]@{Row=$null;Status='Absent'} }
    if (@($rowsPresent | Where-Object { (Read-IdentityDate $_ ExportDateTime).Status -in @('Future','Sentinel','Invalid') }).Count) {
        return [pscustomobject]@{Row=$null;Status='Ambiguous: invalid, sentinel or future observation time'}
    }
    $policies=@($rowsPresent | ForEach-Object { $p=Get-IdentityProperty $_ 'PolicyId'; if($p){[string]$p.Value}else{''} } | Sort-Object -Unique)
    if ($policies.Count -gt 1) { return [pscustomobject]@{Row=$null;Status='Ambiguous: multiple policies'} }
    $fields=@('AggregateState_loc','AggregateState','CurrentDeviceUpdateStatus_loc','CurrentDeviceUpdateStatus','LatestAlertMessage_loc','LatestAlertMessage','BlockingReason','RiskBucket','UpgradeEligibility','UpgradeEligibilityLabel')
    $states=@($rowsPresent | ForEach-Object {
        $record=[ordered]@{}; foreach($field in $fields){$p=Get-IdentityProperty $_ $field;$record[$field]=if($p){[string]$p.Value}else{''}}
        ConvertTo-Json $record -Compress
    } | Sort-Object -Unique)
    if ($states.Count -eq 1) {
        # Return only the invariant business projection, not a row chosen by order.
        $row=$states[0] | ConvertFrom-Json
        return [pscustomobject]@{Row=$row;Status=$(if($rowsPresent.Count -eq 1){'Single observation; recency not established'}else{'Equivalent observations; recency not established'})}
    }
    if (-not $policies[0]) { return [pscustomobject]@{Row=$null;Status='Ambiguous: policy identity missing'} }
    $dated=@($rowsPresent | ForEach-Object { [pscustomobject]@{Row=$_;Date=(Read-IdentityDate $_ ExportDateTime)} })
    if (@($dated | Where-Object { $_.Date.Status -ne 'Qualified' }).Count) { return [pscustomobject]@{Row=$null;Status='Ambiguous: conflicting observations without complete date proof'} }
    $latest=($dated | Sort-Object { $_.Date.Utc } -Descending)[0].Date.Utc
    $tied=@($dated | Where-Object { $_.Date.Utc -eq $latest } | ForEach-Object { $_.Row })
    if ($tied.Count -eq $dated.Count) { return [pscustomobject]@{Row=$null;Status='Ambiguous: conflicting observations at equal qualified time'} }
    $resolved=Resolve-IdentityWindowsUpdateRows -Rows $tied
    if ($null -eq $resolved.Row) { return $resolved }
    [pscustomobject]@{Row=$resolved.Row;Status='Latest qualified observation within one policy'}
}

function Format-IdentityActivityCategory {
    param($Activity, [int]$Days)
    if ($Activity.Recent -eq $true) { return "Within last $Days days" }
    if ($Activity.Old -eq $true) { return "Older than $Days days" }
    return 'Unknown / unqualified activity'
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBLOhDp929WjtC
# 0DVfqe2fcb1cZymMrHJyErZJJ2OiZ6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINvEit+6nJyDmZu/rULl4Pw5Wx2xe61QVMhlZaVFXBLtMA0GCSqG
# SIb3DQEBAQUABIIBgGUs2AiEaQyw/M5I35rk7gE+93GEPFHbskZWKozyPH98hZIG
# x0DDqr+FYqmSFZ9xg5wzP9Q8nl9/Mi6uSEG3hyzQRMCEDy3EKRZkwKbolQY4KQGE
# OSrg1b08Ej+kU2XgUUqwS0T6c3T0SJ4q3K+9GLLeJqsrrbn2wxpK3S9hiSsEhWpo
# I77ODoFONqej5qrKmP4GN9879vo7WtbFx17i/t+9C5dXaHnMzl66Eox74yrewxeO
# k/wfSoTZni7cYdHLIT0wjVDbp2M5NS240m8sPcw6AetxzG+WiDrmpAbPM+7CZum8
# MHMWRoy2f3LbY42GvLAFHoCC6JpekX5rTjXPudKqCDONGmQNm2JCd3pmSPMNICYq
# ZUOW2/D8Hs75UKczN3Eyrp0Aott5Bx0EvjMpkjFmiHuHIzcwXib4/NgbwuCxD5F3
# ML1AzrBi3woOLjdMGY63xuBSM0WdUFdcZbnQ4LmRaW+jeJsxLZbrfFMNHaPkfAT7
# 7YuGoBKzPzPwtOVl86GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIyMTI2
# MjZaMC8GCSqGSIb3DQEJBDEiBCAJRPlKLtx4CZ9fzVJhngzl3Y8FJBFruRFk/nyO
# gcvOHDANBgkqhkiG9w0BAQEFAASCAgAoVODcWAe8MkXUMS8zFE9TjE7qtXjEu+tX
# JcL25KsTf3ZQZg6eec/IIa6pKBmvaPTtjVarafRJ9WIm8rmwyMQ1TpHWh+IjQmsJ
# oQGbFfQUcVv2fICcU0UXaYYKwOIvw8MYJuvraRvi7K8ek9g0yaJtu/l6qpuuSqCL
# nSfp+/Yuau3q59u0DevQBmg2irGoS0uJf4lyC45d5L3DSSLfdaKE7pNqqfrZ5QRq
# 3zX577Y5GdKlkSH76w7HeT9SvVtIO60ipAMfjQdz3uUQ4k8JVmZloeUvbENbxIjt
# ZCupHM6rBJVyRDEWWVf//3y+6xiFKpOG+xRNCfmc/+woXLGa2ZzkPegNnqcXswoB
# MSBNG2w7eszLT4X2mfOkGU32PBIjSVPTfvtIgY9gSzNMKTMmygEcBbLsop+c9gnR
# Ky9TWEaYMah5g3MGy3AfOiVjrIr3axWJbI9O1tI7zcv3ZH+jSbo1c81i8DFjXUh2
# ykP5LhJ8vTa5f3jnI15equZjuefF4ADLAUpi7gXAbB40/o+GSFABYttRdeWjRpD8
# SXDCihi8elc8HDLa4PbJdSROCkxY/bAHljMTZai3i3sRhJBg1c1/yO4pj70tIo3T
# NZPm5KQIFY4yZZHqJVGkBLBJq6uvY2EBQGVKPDoysC6DRsSS3ixn0cxpXansVQPu
# +dgCXMXKXw==
# SIG # End signature block
