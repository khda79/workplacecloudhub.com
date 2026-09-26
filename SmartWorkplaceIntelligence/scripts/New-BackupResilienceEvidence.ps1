[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [Parameter()]
    [string]$OutputRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private'),

    [Parameter()]
    [datetime]$AsOfDate = [datetime]::UtcNow.Date
)

$ErrorActionPreference = 'Stop'

function Get-NormalizedKey {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Test-TrueValue {
    param([AllowNull()][object]$Value)

    return ([string]$Value).Trim() -match '^(?i:true|1|yes)$'
}

function Get-DateValue {
    param([AllowNull()][object]$Value)

    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    return [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)
}

function Get-AgeDays {
    param([AllowNull()][Nullable[datetime]]$SnapshotDate)

    if ($null -eq $SnapshotDate) { return $null }
    return [math]::Max(0, [int][math]::Floor(($AsOfDate.Date - $SnapshotDate.Date).TotalDays))
}

function Get-FreshnessState {
    param([AllowNull()][Nullable[int]]$AgeDays)

    if ($null -eq $AgeDays) { return 'Unknown' }
    if ($AgeDays -le 30) { return 'Fresh' }
    return 'Stale'
}

$paths = @{
    Protected = Join-Path $DataRoot 'M365_Backup_ProtectedMailboxes.csv'
    Coverage  = Join-Path $DataRoot 'M365_BackupPolicyScope_MailboxCoverage.csv'
    Members   = Join-Path $DataRoot 'M365_BackupPolicyScope_GroupMembers.csv'
}

foreach ($path in $paths.Values) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required private evidence file was not found: $path"
    }
}

$protected = @(Import-Csv -LiteralPath $paths.Protected)
$coverage = @(Import-Csv -LiteralPath $paths.Coverage)
$members = @(Import-Csv -LiteralPath $paths.Members)

$protectedByKey = @{}
foreach ($row in $protected) {
    foreach ($value in @($row.MailboxId, $row.Mail, $row.UserPrincipalName)) {
        $key = Get-NormalizedKey $value
        if ($key -and -not $protectedByKey.ContainsKey($key)) {
            $protectedByKey[$key] = $row
        }
    }
}

$memberByKey = @{}
foreach ($row in $members) {
    foreach ($value in @($row.MemberId, $row.MemberMail, $row.MemberUserPrincipalName)) {
        $key = Get-NormalizedKey $value
        if ($key -and -not $memberByKey.ContainsKey($key)) {
            $memberByKey[$key] = $row
        }
    }
}

$coverageKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$matchedProtectionUnits = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$evidence = [System.Collections.Generic.List[object]]::new()

foreach ($row in $coverage) {
    $rowKeys = @(
        $row.MemberId,
        $row.PrimarySmtpAddress,
        $row.MailboxUserPrincipalName,
        $row.MemberUserPrincipalName,
        $row.MemberMail
    ) | ForEach-Object { Get-NormalizedKey $_ } | Where-Object { $_ } | Select-Object -Unique

    foreach ($key in $rowKeys) { [void]$coverageKeys.Add($key) }

    $protectedRow = $null
    foreach ($key in $rowKeys) {
        if ($protectedByKey.ContainsKey($key)) {
            $protectedRow = $protectedByKey[$key]
            break
        }
    }

    $memberRow = $null
    foreach ($key in $rowKeys) {
        if ($memberByKey.ContainsKey($key)) {
            $memberRow = $memberByKey[$key]
            break
        }
    }

    $mailboxFound = Test-TrueValue $row.MailboxFoundInInventory
    if ($protectedRow) { [void]$matchedProtectionUnits.Add([string]$protectedRow.ProtectionUnitId) }

    if (-not $mailboxFound) {
        $protectionState = 'Not applicable'
        $reconciliationState = 'Scope member without mailbox'
        $riskPriority = 'Medium'
        $riskSignal = 'Policy scope member has no mailbox inventory record'
        $recommendedAction = 'Reconcile the group member with mailbox inventory or remove the invalid policy-scope member.'
    }
    elseif ($protectedRow) {
        $protectionState = 'Protected'
        $reconciliationState = 'Expected and protected'
        if ($memberRow -and -not (Test-TrueValue $memberRow.AccountEnabled)) {
            $riskPriority = 'Medium'
            $riskSignal = 'Disabled account remains in expected backup scope'
            $recommendedAction = 'Confirm retention and recovery requirements, then remove obsolete disabled accounts from policy scope.'
        }
        else {
            $riskPriority = 'None'
            $riskSignal = 'No current protection gap'
            $recommendedAction = 'Maintain policy coverage and refresh evidence on schedule.'
        }
    }
    else {
        $protectionState = 'Not protected'
        $reconciliationState = 'Expected not protected'
        $riskPriority = 'High'
        $riskSignal = 'Expected mailbox is not present in protected-mailbox evidence'
        $recommendedAction = 'Add the mailbox to a protection policy and verify a successful protection unit.'
    }

    $protectionSnapshot = if ($protectedRow) { Get-DateValue $protectedRow.ExportDateTime } else { $null }
    $policySnapshot = Get-DateValue $row.RunDateUtc
    $evidenceAgeDays = @(
        Get-AgeDays $protectionSnapshot
        Get-AgeDays $policySnapshot
    ) | Where-Object { $null -ne $_ } | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

    $accountState = if (-not $memberRow) { 'Unknown' } elseif (Test-TrueValue $memberRow.AccountEnabled) { 'Enabled' } else { 'Disabled' }
    $entityKey = if ($row.MemberId) { $row.MemberId } elseif ($protectedRow.MailboxId) { $protectedRow.MailboxId } else { $row.PrimarySmtpAddress }
    $primaryAddress = @($row.PrimarySmtpAddress, $row.MailboxUserPrincipalName, $row.MemberMail, $protectedRow.Mail, $protectedRow.UserPrincipalName) | Where-Object { $_ } | Select-Object -First 1
    $displayName = @($row.MailboxDisplayName, $row.MemberDisplayName, $protectedRow.DisplayName, $primaryAddress) | Where-Object { $_ } | Select-Object -First 1
    $mailboxType = @($protectedRow.MailboxType, $row.RecipientTypeDetails, 'Unknown') | Where-Object { $_ } | Select-Object -First 1

    $evidence.Add([pscustomobject][ordered]@{
        'Backup Entity Key' = [string]$entityKey
        'Display Name' = [string]$displayName
        'Primary Address' = [string]$primaryAddress
        'Entity Type' = 'Mailbox'
        'Mailbox Type' = ([string]$mailboxType).Trim()
        'Account State' = $accountState
        'Expected Scope State' = 'In expected scope'
        'Protection State' = $protectionState
        'Reconciliation State' = $reconciliationState
        'Protection Status' = if ($protectedRow) { [string]$protectedRow.Status } else { 'Not observed' }
        'Policy Scope Group' = [string]$row.GroupDisplayName
        'Protection Policy' = if ($protectedRow) { [string]$protectedRow.ProtectionPolicyName } else { '' }
        'Error State' = if ($protectedRow -and ($protectedRow.ErrorCode -or $protectedRow.ErrorMessage)) { 'Error' } else { 'None' }
        'Error Code' = if ($protectedRow) { [string]$protectedRow.ErrorCode } else { '' }
        'Risk Priority' = $riskPriority
        'Risk Signal' = $riskSignal
        'Recommended Action' = $recommendedAction
        'Protection Snapshot Date' = if ($protectionSnapshot) { $protectionSnapshot.ToString('yyyy-MM-dd') } else { '' }
        'Policy Snapshot Date' = if ($policySnapshot) { $policySnapshot.ToString('yyyy-MM-dd') } else { '' }
        'Protection Evidence Age Days' = Get-AgeDays $protectionSnapshot
        'Policy Evidence Age Days' = Get-AgeDays $policySnapshot
        'Evidence Age Days' = $evidenceAgeDays
        'Freshness State' = Get-FreshnessState $evidenceAgeDays
        'Evidence Status' = 'Observed'
    })
}

foreach ($row in $protected) {
    if ($matchedProtectionUnits.Contains([string]$row.ProtectionUnitId)) { continue }

    $rowKeys = @($row.MailboxId, $row.Mail, $row.UserPrincipalName) | ForEach-Object { Get-NormalizedKey $_ } | Where-Object { $_ }
    if ($rowKeys | Where-Object { $coverageKeys.Contains($_) }) { continue }

    $protectionSnapshot = Get-DateValue $row.ExportDateTime
    $evidenceAgeDays = Get-AgeDays $protectionSnapshot
    $primaryAddress = @($row.Mail, $row.UserPrincipalName) | Where-Object { $_ } | Select-Object -First 1

    $evidence.Add([pscustomobject][ordered]@{
        'Backup Entity Key' = [string]$row.MailboxId
        'Display Name' = [string]$row.DisplayName
        'Primary Address' = [string]$primaryAddress
        'Entity Type' = 'Mailbox'
        'Mailbox Type' = if ($row.MailboxType) { [string]$row.MailboxType } else { 'Unknown' }
        'Account State' = 'Unknown'
        'Expected Scope State' = 'Outside expected scope'
        'Protection State' = 'Protected'
        'Reconciliation State' = 'Protected outside expected scope'
        'Protection Status' = [string]$row.Status
        'Policy Scope Group' = ''
        'Protection Policy' = [string]$row.ProtectionPolicyName
        'Error State' = if ($row.ErrorCode -or $row.ErrorMessage) { 'Error' } else { 'None' }
        'Error Code' = [string]$row.ErrorCode
        'Risk Priority' = 'Medium'
        'Risk Signal' = 'Protected mailbox is outside the expected policy scope'
        'Recommended Action' = 'Confirm policy intent and align the protected mailbox with the governed scope.'
        'Protection Snapshot Date' = if ($protectionSnapshot) { $protectionSnapshot.ToString('yyyy-MM-dd') } else { '' }
        'Policy Snapshot Date' = ''
        'Protection Evidence Age Days' = Get-AgeDays $protectionSnapshot
        'Policy Evidence Age Days' = $null
        'Evidence Age Days' = $evidenceAgeDays
        'Freshness State' = Get-FreshnessState $evidenceAgeDays
        'Evidence Status' = 'Observed'
    })
}

New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
$outputPath = Join-Path $OutputRoot 'BackupMailboxEvidence.csv'
$evidence | Sort-Object 'Reconciliation State', 'Primary Address' | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding utf8

$summary = [pscustomobject]@{
    OutputPath = $outputPath
    Rows = $evidence.Count
    ExpectedMailboxes = @($evidence | Where-Object { $_.'Expected Scope State' -eq 'In expected scope' -and $_.'Reconciliation State' -ne 'Scope member without mailbox' }).Count
    ExpectedAndProtected = @($evidence | Where-Object { $_.'Reconciliation State' -eq 'Expected and protected' }).Count
    ExpectedNotProtected = @($evidence | Where-Object { $_.'Reconciliation State' -eq 'Expected not protected' }).Count
    ScopeMembersWithoutMailbox = @($evidence | Where-Object { $_.'Reconciliation State' -eq 'Scope member without mailbox' }).Count
    ProtectedOutsideExpectedScope = @($evidence | Where-Object { $_.'Reconciliation State' -eq 'Protected outside expected scope' }).Count
    DisabledScopeMembers = @($evidence | Where-Object { $_.'Expected Scope State' -eq 'In expected scope' -and $_.'Account State' -eq 'Disabled' }).Count
    ProtectionErrors = @($evidence | Where-Object { $_.'Error State' -eq 'Error' }).Count
    OldestEvidenceDays = ($evidence | Measure-Object 'Evidence Age Days' -Maximum).Maximum
}

$summary | Format-List

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCANbXNjpy7gpt1G
# bCJcS5hpF7yECz13I/D/kdPz8G1B6aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGviYRoHCgaNTxiUWPu0Prect5SD8DJxmgldS985IwpBMA0GCSqG
# SIb3DQEBAQUABIIBgCKEBLoHGjB+FELbTgJtvyruejA9CqdQg9sphCi8phK9oQCb
# HqOB+DeILsLKirCDPC2nJ/4FNR1AVInXYFk+rh1jYn3xR1ddzA+Xp2vZwub+EZX6
# 8OY5deqNN0iXO4Y3/VkX8YobUmlfYO697jOSOyEOqqU8BGKEX86d70DzxmMgV40d
# zlmtPldKJLWRul2IogGbXTsnImRXlIIU8Dk9BhdjwGyUA+8oASxCQzf1jYwrMI/V
# WT+qMEv3UhPMS7XiwLwFUMzO91wCFnx6gsWGyZAtyva+gwv7SjQ9eJQvI5hj182Q
# EUJR5QtbyeAdT0FSkX6f+Sw6/OiWqIlUNnLtkbc8wSVtBM3bvh2td94dJyLG/Dar
# TITjxLPs1swgxTlj4/vKN2a+fTsVlgvP9EkJycn72C/PB1biV7CK7je8vOpHpeBz
# S39hChHQ/End++8QFJY9vp7aZ5a37Ts2v0fVHioWZ4keujCcE5U6VWCIrjbNygZr
# EnsIWy4QuPJrt1cxgaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjYyMzUw
# MDBaMC8GCSqGSIb3DQEJBDEiBCBaXHSridt2iNihgEUGHIWZVeUvJvf5qpksNmEX
# vJCzqjANBgkqhkiG9w0BAQEFAASCAgBAz3Osq+9PCSgfTLM0HCWzhUkXUZ3XxiNh
# 9wl8X6iMrC2l+LvWFnAZzZI2VlWoLqBZy5ePgvUH5QH/jbtViZ9p5jzvx38JCjUV
# WmB7Hxn3Rn89zLjhXE8MtIYDuy++QXDD5BNmQ4FPabvagndvIWUvREo9c3MpQJx8
# lXI1bCFxvY+avlva5AVAc9JJtSjxna6qbS9UnX0umQjC7N1JwFLGZLdl7FWJagF1
# AnAS/irhT66iTMnS1vQBW3ODtGI0lFKXb6g+MBeJNadzyfg2Y+fDMcRfmGdpevrz
# gCwzZczWg5OUKc2E4zVGTxqSbm0zD1Y9sAekRbqbCDfmzVSRDf+52NX+Bs9M+T+t
# R493dvx7rYCqEKjRXQqwT3Pv9RNsYF45NcE+ZiNtwjJKmaKo2wR5HVBbb0L06wj8
# 4tsYF/NbPiH6tt7U+Ydqtq3e+2ye83LX+Ce00Gn+ABQLgJLx6Fx5rk3CRbJp4Oae
# A6H3rP2ZUXgllOKt27NA2wzAleHmBitDPEQfjw7Rty+VhR/jon4MMufuatw2+VeG
# bhpcLQZI/kswGMiFuM64LAfpj4uC8fMuaPl8ENy1/RzVLEUaVltl7KxYQq2kaeeU
# 5BBQnowhFSZesSh/n6PTr8s5utufQHpOAuOofmpCwdcHYm62U+8IGlvwlocW8Jr2
# m57tby+A4g==
# SIG # End signature block
