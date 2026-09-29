[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$UserEvidencePath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\UserInventoryEvidence.csv'),

    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\MailboxEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-NormalizedKey {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Convert-ToDoubleOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $text = ([string]$Value).Trim().Replace([string][char]0x00A0, '').Replace(' ', '')
    $number = 0.0
    $cultures = if ($text.Contains(',') -and -not $text.Contains('.')) {
        @([Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [Globalization.CultureInfo]::InvariantCulture)
    } elseif ($text.Contains(',') -and $text.Contains('.') -and $text.LastIndexOf(',') -gt $text.LastIndexOf('.')) {
        @([Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [Globalization.CultureInfo]::InvariantCulture)
    } else {
        @([Globalization.CultureInfo]::InvariantCulture, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'))
    }
    foreach ($culture in $cultures) {
        if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, $culture, [ref]$number)) { return $number }
    }
    return $null
}

function Convert-ToDateOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) { return $date }
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [Globalization.DateTimeStyles]::AssumeLocal, [ref]$date)) { return $date }
    return $null
}

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([double]$Value).ToString('0.0', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ListCount {
    param([AllowNull()][object[]]$Values)
    $items = foreach ($value in $Values) {
        if ([string]::IsNullOrWhiteSpace([string]$value)) { continue }
        [string]$value -split '[;|,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
    return @($items | Sort-Object -Unique).Count
}

function Get-DelegationTypes {
    param([AllowNull()][object]$SendAs, [AllowNull()][object]$FullAccess, [AllowNull()][object]$SendOnBehalf)
    $types = @()
    if ((Get-ListCount @($SendAs)) -gt 0) { $types += 'Send As' }
    if ((Get-ListCount @($FullAccess)) -gt 0) { $types += 'Full Access' }
    if ((Get-ListCount @($SendOnBehalf)) -gt 0) { $types += 'Send on Behalf' }
    if ($types.Count -eq 0) { return 'None' }
    return ($types -join '; ')
}

function Get-ExchangeVersion {
    param([AllowNull()][object]$AdminDisplayVersion)
    $value = [string]$AdminDisplayVersion
    if ($value -match 'Version\s+15\.1') { return 'Exchange Server 2016' }
    if ($value -match 'Version\s+15\.2\s+\(Build\s+(\d+)') {
        if ([int]$Matches[1] -ge 2562) { return 'Exchange Server SE' }
        return 'Exchange Server 2019'
    }
    return 'Exchange Server - version not observed'
}

function Get-RiskAssessment {
    param(
        [AllowNull()][Nullable[double]]$MailboxSize,
        [string]$ArchiveState,
        [int]$DelegationCount,
        [AllowNull()][Nullable[datetime]]$LastActivity,
        [string]$ForwardingState,
        [datetime]$AsOfDate
    )
    if ($null -ne $MailboxSize -and $MailboxSize -ge 45 -and $ArchiveState -ne 'Enabled') {
        return @('High', 'Mailbox >=45 GB without an enabled archive', 'Enable or validate the archive and review retention before quota pressure causes disruption.')
    }
    if ($DelegationCount -ge 10) {
        return @('High', 'Mailbox has extensive delegation', 'Review Full Access, Send As, and Send on Behalf assignments with the mailbox owner.')
    }
    if ($null -ne $LastActivity -and ($AsOfDate - $LastActivity).TotalDays -gt 180) {
        return @('Medium', 'Mailbox has no recent observed activity', 'Confirm ownership and retention requirements before deprovisioning or migration action.')
    }
    if ($ForwardingState -eq 'Configured') {
        return @('Medium', 'Mailbox forwarding is configured', 'Validate forwarding destination, business need, and data-loss prevention controls.')
    }
    return @('None', 'No current mailbox attention signal', 'Maintain lifecycle, archive, and delegation governance and refresh evidence on schedule.')
}

$paths = @{
    Exo = Join-Path $DataRoot 'Exchange_EXO_Mailboxes_AllDomains.csv'
    Stats = Join-Path $DataRoot 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv'
    Archive = Join-Path $DataRoot 'Exchange_EXO_Mailboxes_AllDomains_Archive.csv'
    Permissions = Join-Path $DataRoot 'Exchange_EXO_Mailboxes_AllDomains_Permissions.csv'
    OnPrem = Join-Path $DataRoot 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
    Servers = Join-Path $DataRoot 'Exchange_OnPrem_Servers_Inventory.csv'
}
foreach ($path in @($paths.Values) + @($UserEvidencePath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Required private evidence file was not found: $path" }
}

$exo = @(Import-Csv -LiteralPath $paths.Exo)
$stats = @(Import-Csv -LiteralPath $paths.Stats)
$archives = @(Import-Csv -LiteralPath $paths.Archive)
$permissions = @(Import-Csv -LiteralPath $paths.Permissions)
$onPrem = @(Import-Csv -LiteralPath $paths.OnPrem)
$servers = @(Import-Csv -LiteralPath $paths.Servers -Delimiter ';')
$users = @(Import-Csv -LiteralPath $UserEvidencePath)
$asOfDate = (Get-Item -LiteralPath $paths.Exo).LastWriteTime.Date

function New-Index {
    param([object[]]$Rows, [string[]]$Properties)
    $index = @{}
    foreach ($row in $Rows) {
        foreach ($property in $Properties) {
            $key = Get-NormalizedKey $row.$property
            if ($key -and -not $index.ContainsKey($key)) { $index[$key] = $row }
        }
    }
    return $index
}

$statsIndex = New-Index $stats @('PrimarySmtpAddress','UserPrincipalName')
$archiveIndex = New-Index $archives @('PrimarySmtpAddress','UserPrincipalName')
$permissionIndex = New-Index $permissions @('PrimarySmtpAddress','UserPrincipalName')
$userIndex = New-Index $users @('User Principal Name')
$serverVersion = @{}
foreach ($server in $servers) {
    $key = Get-NormalizedKey $server.Name
    if ($key) { $serverVersion[$key] = Get-ExchangeVersion $server.AdminDisplayVersion }
}

$rows = [System.Collections.Generic.List[object]]::new()
$sequence = 0
foreach ($mailbox in $exo) {
    $sequence++
    $key = Get-NormalizedKey $mailbox.PrimarySmtpAddress
    if (-not $key) { $key = Get-NormalizedKey $mailbox.UserPrincipalName }
    $statsRow = if ($key -and $statsIndex.ContainsKey($key)) { $statsIndex[$key] } else { $null }
    $archiveRow = if ($key -and $archiveIndex.ContainsKey($key)) { $archiveIndex[$key] } else { $null }
    $permissionRow = if ($key -and $permissionIndex.ContainsKey($key)) { $permissionIndex[$key] } else { $null }
    $userRow = if ($key -and $userIndex.ContainsKey($key)) { $userIndex[$key] } else { $null }
    $size = Convert-ToDoubleOrNull $(if ($statsRow) { $statsRow.TotalItemSizeGB } else { $mailbox.TotalItemSizeGB })
    $archiveSize = Convert-ToDoubleOrNull $(if ($statsRow -and $statsRow.Archive_TotalItemSizeGB) { $statsRow.Archive_TotalItemSizeGB } elseif ($archiveRow) { $archiveRow.Archive_TotalItemSizeGB } else { $null })
    $archiveState = if (([string]$mailbox.ArchiveStatus -match '^(?i:active|enabled)$') -or ($null -ne $archiveSize -and $archiveSize -gt 0)) { 'Enabled' } else { 'Disabled' }
    $sendAs = if ($permissionRow) { $permissionRow.SendAs } else { $mailbox.SendAs }
    $fullAccess = if ($permissionRow) { $permissionRow.FullAccess } else { $mailbox.FullAccess }
    $sendOnBehalf = if ($permissionRow) { $permissionRow.GrantSendOnBehalfTo } else { $mailbox.GrantSendOnBehalfTo }
    $delegations = Get-ListCount @($sendAs, $fullAccess, $sendOnBehalf)
    $delegationTypes = Get-DelegationTypes $sendAs $fullAccess $sendOnBehalf
    $lastActivity = Convert-ToDateOrNull $(if ($statsRow -and $statsRow.LastUserActionTime) { $statsRow.LastUserActionTime } else { $mailbox.LastUserActionTime })
    $forwardingState = if ($mailbox.ForwardingSmtpAddress -or $mailbox.ForwardingAddress) { 'Configured' } else { 'None' }
    $risk = Get-RiskAssessment $size $archiveState $delegations $lastActivity $forwardingState $asOfDate
    $rows.Add([pscustomobject][ordered]@{
        'Mailbox Key' = ('Mailbox {0:D6}' -f $sequence)
        'Primary SMTP Address' = [string]$mailbox.PrimarySmtpAddress
        'Display Name' = [string]$mailbox.DisplayName
        'Country' = if ($userRow -and $userRow.Country) { [string]$userRow.Country } else { 'Unknown' }
        'Hosting Location' = 'Exchange Online'
        'Exchange Version' = 'Exchange Online'
        'Mailbox Type' = if ($mailbox.MailboxType) { [string]$mailbox.MailboxType } else { [string]$mailbox.RecipientTypeDetails }
        'Recipient Type' = [string]$mailbox.RecipientTypeDetails
        'Operational State' = if (([string]$mailbox.AccountEnabled).Trim() -match '^(?i:false|0|no)$') { 'Disabled' } else { 'Active' }
        'Mailbox Size GB' = Convert-ToInvariantDecimalText $size
        'Archive State' = $archiveState
        'Archive Size GB' = Convert-ToInvariantDecimalText $archiveSize
        'Total Storage GB' = Convert-ToInvariantDecimalText $(if ($null -eq $size -and $null -eq $archiveSize) { $null } else { [double]($size ?? 0) + [double]($archiveSize ?? 0) })
        'Delegation Count' = $delegations
        'Delegation Types' = $delegationTypes
        'Forwarding State' = $forwardingState
        'Last Activity Date' = if ($lastActivity) { $lastActivity.ToString('yyyy-MM-dd') } else { '' }
        'Hosting State' = 'Hosted online'
        'Risk Priority' = $risk[0]
        'Risk Signal' = $risk[1]
        'Recommended Action' = $risk[2]
        'Evidence Status' = 'Observed'
    })
}

foreach ($mailbox in $onPrem) {
    $sequence++
    $key = Get-NormalizedKey $mailbox.PrimarySMTPaddress
    if (-not $key) { $key = Get-NormalizedKey $mailbox.UserPrincipalName }
    $userRow = if ($key -and $userIndex.ContainsKey($key)) { $userIndex[$key] } else { $null }
    $serverKey = Get-NormalizedKey $mailbox.ServerName
    $exchangeVersion = if ($serverKey -and $serverVersion.ContainsKey($serverKey)) { $serverVersion[$serverKey] } else { 'Exchange Server - version not observed' }
    $sizeMb = Convert-ToDoubleOrNull $mailbox.'TotalItemSize-In-MB'
    $size = if ($null -ne $sizeMb) { $sizeMb / 1024 } else { $null }
    $archiveSizeMb = Convert-ToDoubleOrNull $mailbox.'ArchiveTotalItemSize-In-MB'
    $archiveSize = if ($null -ne $archiveSizeMb) { $archiveSizeMb / 1024 } else { $null }
    $archiveState = if (([string]$mailbox.ArchiveStatus -match '^(?i:active|enabled)$') -or ($null -ne $archiveSize -and $archiveSize -gt 0)) { 'Enabled' } else { 'Disabled' }
    $delegations = Get-ListCount @($mailbox.SendAs, $mailbox.FullAccess, $mailbox.GrantSendOnBehalfTo)
    $delegationTypes = Get-DelegationTypes $mailbox.SendAs $mailbox.FullAccess $mailbox.GrantSendOnBehalfTo
    $lastActivity = Convert-ToDateOrNull $mailbox.LastLogonTime
    $forwardingState = if ($mailbox.ForwardingSmtpAddress -or $mailbox.ForwardingAddress) { 'Configured' } else { 'None' }
    $risk = Get-RiskAssessment $size $archiveState $delegations $lastActivity $forwardingState $asOfDate
    $rows.Add([pscustomobject][ordered]@{
        'Mailbox Key' = ('Mailbox {0:D6}' -f $sequence)
        'Primary SMTP Address' = [string]$mailbox.PrimarySMTPaddress
        'Display Name' = [string]$mailbox.DisplayName
        'Country' = if ($userRow -and $userRow.Country) { [string]$userRow.Country } else { 'Unknown' }
        'Hosting Location' = 'On-premises'
        'Exchange Version' = $exchangeVersion
        'Mailbox Type' = if (([string]$mailbox.IsShared).Trim() -match '^(?i:true|1|yes)$') { 'SharedMailbox' } elseif (([string]$mailbox.IsResource).Trim() -match '^(?i:true|1|yes)$') { 'ResourceMailbox' } else { 'UserMailbox' }
        'Recipient Type' = [string]$mailbox.RecipientType
        'Operational State' = if (([string]$mailbox.AccountDisabled).Trim() -match '^(?i:true|1|yes)$') { 'Disabled' } else { 'Active' }
        'Mailbox Size GB' = Convert-ToInvariantDecimalText $size
        'Archive State' = $archiveState
        'Archive Size GB' = Convert-ToInvariantDecimalText $archiveSize
        'Total Storage GB' = Convert-ToInvariantDecimalText $(if ($null -eq $size -and $null -eq $archiveSize) { $null } else { [double]($size ?? 0) + [double]($archiveSize ?? 0) })
        'Delegation Count' = $delegations
        'Delegation Types' = $delegationTypes
        'Forwarding State' = $forwardingState
        'Last Activity Date' = if ($lastActivity) { $lastActivity.ToString('yyyy-MM-dd') } else { '' }
        'Hosting State' = 'Hosted on-premises'
        'Risk Priority' = $risk[0]
        'Risk Signal' = $risk[1]
        'Recommended Action' = $risk[2]
        'Evidence Status' = 'Observed'
    })
}

$outputDirectory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$rows | Sort-Object 'Hosting Location', 'Primary SMTP Address' | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject]@{
    OutputPath = $OutputPath
    Mailboxes = $rows.Count
    ExchangeOnline = @($rows | Where-Object 'Hosting Location' -eq 'Exchange Online').Count
    OnPremises = @($rows | Where-Object 'Hosting Location' -eq 'On-premises').Count
    WithArchive = @($rows | Where-Object 'Archive State' -eq 'Enabled').Count
    WithDelegations = @($rows | Where-Object 'Delegation Count' -gt 0).Count
    HighRisk = @($rows | Where-Object 'Risk Priority' -eq 'High').Count
} | Format-List

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCi60wOdX17wJ3k
# yUm8Q0pdIxvn5yJw2TBXzOm+7EONJaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDW4FX7b3XTU40fY5w5AnD3r6bIdAzP66OAl3Q3fwN/RMA0GCSqG
# SIb3DQEBAQUABIIBgFRtHMsBFqAeYR+OHgsqwaO0ucQrwwYqiWU628Ii75hck/Yb
# 05GeWqgn9L0kIJym8VV+9sEJC4P3niRIrN7YbXMhVM9ovBAgE5nC2sRjNVGBxUMW
# v4F0cmgL5YidZ4UdV9r5KQnnOvkJD0Kawbrvz31B6WxShZFfuG+OsqsnHOeuNkYj
# 0ojBKc3/SVdzVDQVopRg9YHREDrod0gdQG6C9eM9VG0c5QZoptL1ZjqulI6PkXLY
# HgwQy2B7uM7+NLhoETjWfn3gwfhrfAfFZKT69smx8v8y2ARJiVrfpGyuhwOf95dd
# /fFqtVYnL36NP1vA4FFI6zcZcopnQ5EyqhsRQV5QntZquJXLJ2MYpobVU9HC3F/1
# 2BudAvzTm1UV+D+0LczfV0MOn8m2h7xBSG0TM1HrkkydeG0vTdzsWrJbqRgt6qZr
# FjK6XdMewAjHPfLYPs1itxZJ3c70EcJeWUHQst/Gz/tcx0PezClptEGB7cPk/r1j
# yDOuwQl97K3tkC/K/qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjkxOTEx
# MzFaMC8GCSqGSIb3DQEJBDEiBCCAgoFTN9KQ90MloJk3wTOs+sr7DxEaQtaM3b5i
# kCSCJjANBgkqhkiG9w0BAQEFAASCAgCa+INhagE8mcimeY2sO1nqShMroBrChnBp
# qbT9+UWi9w9STyaD5rNja/lLpBuNgl8XdKGeU37TV2fOd6OS5FOhXT+pD29Q3JE3
# qKS4WRFcC//EzsOevzMcLQ0ZF0H+zz0LKnlCpmX5a8rzIIdq5JIMmMTuD1GxxC5g
# +YIbPdYth8RMc6CNZVuVsl9nsJxMVmGKuvwKX12CFrU7XOF3zfWR8LHdv7VmauuW
# pYl2QZDmEiYQjUjqt0IEIGHyE2mEi26l4l5YcPMbbWLnUsG/hOhMJMxfRpkYubXd
# PVmHQjSfFuUmGebbjsQ133OHh0uoNwe6BKAA/qKFB7oTZLrSf8Q0epVk0OY2Jm05
# RhHtezDu8BdAjBjPGYjeRqnnPUnPAsRp6v9DoSHFaQ/a8t97Qcopq5nMqZpkZoqJ
# W+oWiTrPyt+AFbFeHGz/HT4c9XQ2WdwpWlRje0ChAN+PLrn1IUK7BBs8GNYx/VJe
# +JwyoDnzanNyN+pVnvBKR3k8RSrcrX33Qy52fRp7iDwXPzh0ElQFm07Ijl8uTt2L
# L4+eCnQYkpPRqsOltUu8OwBiNDrG2OzZ515n6+zoTABtgmeXu/PpvO8xqNDz3kk0
# T2KzP4UeCobJuK0qbGpyBbC2CwrBOU55d48pDwZCoDNGUSkyCatkjBYhZgVYfaUZ
# i3rzLkFzYQ==
# SIG # End signature block
