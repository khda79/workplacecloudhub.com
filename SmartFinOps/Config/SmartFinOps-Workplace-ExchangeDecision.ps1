Set-StrictMode -Version 2.0

function Get-SmartFinOpsExchangeIdentityKey {
    [CmdletBinding()]
    param([AllowNull()]$Row)

    if ($null -eq $Row) { return '' }
    $identity = [string](Get-RowPropertyValue -Row $Row -Names @(
        'UserPrincipalName', 'User Principal Name', 'User principal name',
        'PrimarySmtpAddress', 'Primary SMTP Address'
    ))
    if ([string]::IsNullOrWhiteSpace($identity)) { return '' }
    return $identity.Trim().ToLowerInvariant()
}

function New-SmartFinOpsExchangeRowMap {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()][object[]]$Rows)

    $map = @{}
    foreach ($row in $Rows) {
        $key = Get-SmartFinOpsExchangeIdentityKey -Row $row
        if ($key) { $map[$key] = $row }
    }
    return $map
}

function ConvertTo-SmartFinOpsExchangeDecimalOrNull {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return ConvertTo-SmartFinOpsNumberOrNull $text
}

function Test-SmartFinOpsExchangeListValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $false }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    return $text -notin @('None', '[]', '{}', 'System.Object[]', 'N/A', '-')
}

function Get-SmartFinOpsExchangePermissionValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$PropertyName
    )

    $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $Rows) {
        if ($null -eq $row) { continue }
        $rawValue = Get-RowPropertyValue -Row $row -Names @($PropertyName)
        if (-not (Test-SmartFinOpsExchangeListValue -Value $rawValue)) { continue }
        foreach ($value in ([string]$rawValue -split ';')) {
            $trimmed = $value.Trim()
            if (Test-SmartFinOpsExchangeListValue -Value $trimmed) { [void]$values.Add($trimmed) }
        }
    }
    return (@($values | Sort-Object) -join ' | ')
}

function Get-SmartFinOpsExchangeDelegateValues {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()][AllowEmptyCollection()][object[]]$Rows)

    $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($propertyName in @('FullAccess', 'SendAs', 'GrantSendOnBehalfTo')) {
        $propertyValues = Get-SmartFinOpsExchangePermissionValue -Rows $Rows -PropertyName $propertyName
        foreach ($value in ($propertyValues -split ' \| ')) {
            if (Test-SmartFinOpsExchangeListValue -Value $value) { [void]$values.Add($value.Trim()) }
        }
    }
    return @($values | Sort-Object)
}

function New-SmartFinOpsExchangeOptimizationRow {
    [CmdletBinding()]
    param(
        [string]$FindingType, [string]$Identity, [string]$Severity, [string]$IssueType,
        [string]$Status, [string]$MailboxLocation, [AllowNull()]$Enabled, [string]$Details,
        [string]$RecipientTypeDetails = '', [string]$AccountStateStatus = '',
        [AllowNull()]$M365AccountEnabled = '', [AllowNull()]$ADEnabled = '', [AllowNull()]$EXOAccountEnabled = '',
        [AllowNull()]$HasDelegates = '', [AllowNull()]$DelegateCount = '',
        [string]$FullAccess = '', [string]$SendAs = '', [string]$GrantSendOnBehalfTo = '',
        [AllowNull()]$MailboxSizeGB = '', [string]$CapacityBand = '',
        [AllowNull()]$HasArchive = '', [AllowNull()]$ArchiveSizeGB = '',
        [AllowNull()]$LitigationHoldEnabled = '', [AllowNull()]$RetentionHoldEnabled = '',
        [AllowNull()]$IsLikelyServiceAccount = '', [string]$AssignedSkus = '',
        [string]$CurrentBaseLicense = '', [AllowNull()]$IndicativeMonthlyValueEUR = '',
        [string]$Currency = '', [string]$ExcludedReasons = '', [string]$Recommendation = '', [string]$Guardrail = ''
    )

    [pscustomobject]@{
        RunId = $script:RunId
        FindingType = $FindingType
        Identity = $Identity
        Severity = $Severity
        IssueType = $IssueType
        Status = $Status
        MailboxLocation = $MailboxLocation
        Enabled = $Enabled
        Details = $Details
        RecipientTypeDetails = $RecipientTypeDetails
        AccountStateStatus = $AccountStateStatus
        M365AccountEnabled = $M365AccountEnabled
        ADEnabled = $ADEnabled
        EXOAccountEnabled = $EXOAccountEnabled
        HasDelegates = $HasDelegates
        DelegateCount = $DelegateCount
        FullAccess = $FullAccess
        SendAs = $SendAs
        GrantSendOnBehalfTo = $GrantSendOnBehalfTo
        MailboxSizeGB = $MailboxSizeGB
        CapacityBand = $CapacityBand
        HasArchive = $HasArchive
        ArchiveSizeGB = $ArchiveSizeGB
        LitigationHoldEnabled = $LitigationHoldEnabled
        RetentionHoldEnabled = $RetentionHoldEnabled
        IsLikelyServiceAccount = $IsLikelyServiceAccount
        AssignedSkus = $AssignedSkus
        CurrentBaseLicense = $CurrentBaseLicense
        IndicativeMonthlyValueEUR = $IndicativeMonthlyValueEUR
        Currency = $Currency
        ExcludedReasons = $ExcludedReasons
        Recommendation = $Recommendation
        Guardrail = $Guardrail
    }
}

function New-SmartFinOpsSharedMailboxConversionRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MailboxRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MailboxStatsRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MailboxArchiveRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MailboxPermissionRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MailboxUsageRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$M365Users,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ADUsers,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$UserDecisionRows,
        [string]$Currency = 'EUR'
    )

    $statsByIdentity = New-SmartFinOpsExchangeRowMap -Rows $MailboxStatsRows
    $archiveByIdentity = New-SmartFinOpsExchangeRowMap -Rows $MailboxArchiveRows
    $permissionsByIdentity = New-SmartFinOpsExchangeRowMap -Rows $MailboxPermissionRows
    $usageByIdentity = New-SmartFinOpsExchangeRowMap -Rows $MailboxUsageRows
    $m365ByIdentity = New-SmartFinOpsExchangeRowMap -Rows $M365Users
    $adByIdentity = New-SmartFinOpsExchangeRowMap -Rows $ADUsers
    $decisionByIdentity = New-SmartFinOpsExchangeRowMap -Rows $UserDecisionRows
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($mailbox in $MailboxRows) {
        $recipientType = [string](Get-RowPropertyValue -Row $mailbox -Names @('RecipientTypeDetails'))
        if ($recipientType -ne 'UserMailbox') { continue }

        $key = Get-SmartFinOpsExchangeIdentityKey -Row $mailbox
        if (-not $key -or -not $decisionByIdentity.ContainsKey($key)) { continue }
        $decision = $decisionByIdentity[$key]
        $assignedSkus = [string](Get-RowPropertyValue -Row $decision -Names @('AssignedSkuPartNumbers'))
        if ([string]::IsNullOrWhiteSpace($assignedSkus)) { continue }

        $stats = if ($statsByIdentity.ContainsKey($key)) { $statsByIdentity[$key] } else { $null }
        $archive = if ($archiveByIdentity.ContainsKey($key)) { $archiveByIdentity[$key] } else { $null }
        $permissions = if ($permissionsByIdentity.ContainsKey($key)) { $permissionsByIdentity[$key] } else { $null }
        $usage = if ($usageByIdentity.ContainsKey($key)) { $usageByIdentity[$key] } else { $null }
        $m365User = if ($m365ByIdentity.ContainsKey($key)) { $m365ByIdentity[$key] } else { $null }
        $adUser = if ($adByIdentity.ContainsKey($key)) { $adByIdentity[$key] } else { $null }

        $permissionEvidence = @($mailbox, $permissions)
        $delegates = @(Get-SmartFinOpsExchangeDelegateValues -Rows $permissionEvidence)
        if ($delegates.Count -eq 0) { continue }
        $fullAccess = Get-SmartFinOpsExchangePermissionValue -Rows $permissionEvidence -PropertyName 'FullAccess'
        $sendAs = Get-SmartFinOpsExchangePermissionValue -Rows $permissionEvidence -PropertyName 'SendAs'
        $sendOnBehalf = Get-SmartFinOpsExchangePermissionValue -Rows $permissionEvidence -PropertyName 'GrantSendOnBehalfTo'

        $m365Enabled = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $m365User -Names @('AccountEnabled'))
        if ($null -eq $m365Enabled) {
            $blocked = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $m365User -Names @('Block credential'))
            if ($null -ne $blocked) { $m365Enabled = -not $blocked }
        }
        $adEnabled = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $adUser -Names @('Enabled'))
        $exoEnabled = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $mailbox -Names @('AccountEnabled'))
        $knownStates = @($m365Enabled, $adEnabled, $exoEnabled) | Where-Object { $null -ne $_ }
        $hasDisabledState = @($knownStates | Where-Object { $_ -eq $false }).Count -gt 0
        $hasEnabledState = @($knownStates | Where-Object { $_ -eq $true }).Count -gt 0
        $directoryDisabled = ($m365Enabled -eq $false) -or ($adEnabled -eq $false)
        $accountStateStatus = if ($hasDisabledState -and $hasEnabledState) { 'Conflict' } elseif ($directoryDisabled) { 'Disabled confirmed' } else { 'Not disabled in AD or Entra' }
        if ($accountStateStatus -eq 'Not disabled in AD or Entra') { continue }

        $mailboxSize = ConvertTo-SmartFinOpsExchangeDecimalOrNull (Get-RowPropertyValue -Row $mailbox -Names @('TotalItemSizeGB'))
        if ($null -eq $mailboxSize) { $mailboxSize = ConvertTo-SmartFinOpsExchangeDecimalOrNull (Get-RowPropertyValue -Row $stats -Names @('TotalItemSizeGB')) }
        if ($null -eq $mailboxSize) {
            $storageBytes = ConvertTo-SmartFinOpsExchangeDecimalOrNull (Get-RowPropertyValue -Row $usage -Names @('Storage Used (Byte)'))
            if ($null -ne $storageBytes) { $mailboxSize = [math]::Round([decimal]($storageBytes / 1GB), 2) }
        }
        $archiveSize = ConvertTo-SmartFinOpsExchangeDecimalOrNull (Get-RowPropertyValue -Row $archive -Names @('Archive_TotalItemSizeGB'))
        if ($null -eq $archiveSize) { $archiveSize = ConvertTo-SmartFinOpsExchangeDecimalOrNull (Get-RowPropertyValue -Row $stats -Names @('Archive_TotalItemSizeGB')) }
        $archiveStatus = ([string](Get-RowPropertyValue -Row $mailbox -Names @('ArchiveStatus'))).Trim()
        $hasArchiveFromUsage = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $usage -Names @('Has Archive'))
        $hasArchive = ($archiveStatus -match '^(Active|Enabled|True)$') -or ($null -ne $archiveSize -and $archiveSize -gt 0) -or ($hasArchiveFromUsage -eq $true)
        $litigationHold = (ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $mailbox -Names @('LitigationHoldEnabled'))) -eq $true
        $retentionHold = (ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $mailbox -Names @('RetentionHoldEnabled'))) -eq $true
        $isServiceAccount = ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $adUser -Names @('IsLikelyServiceAccount'))) -eq $true) -or ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $adUser -Names @('IsLikelyPrivilegedOrServiceAccount'))) -eq $true)

        $excludedReasons = New-Object System.Collections.Generic.List[string]
        if ($archiveStatus -notmatch '^(None|Disabled|False|Active|Enabled|True)$' -or
            $null -eq (ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $mailbox -Names @('LitigationHoldEnabled'))) -or
            $null -eq (ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $mailbox -Names @('RetentionHoldEnabled')))) { $excludedReasons.Add('Archive or hold state is unknown') | Out-Null }
        if ($mailboxSize -lt 0) { $excludedReasons.Add('Mailbox size is invalid') | Out-Null }
        if ($accountStateStatus -eq 'Conflict') { $excludedReasons.Add('Conflicting enabled/disabled account states') | Out-Null }
        if ($null -eq $mailboxSize) { $excludedReasons.Add('Mailbox size is unavailable') | Out-Null } elseif ($mailboxSize -ge 50) { $excludedReasons.Add('Mailbox size is at or above 50 GB') | Out-Null }
        if ($hasArchive) { $excludedReasons.Add('Archive mailbox evidence is present') | Out-Null }
        if ($litigationHold) { $excludedReasons.Add('Litigation hold is enabled') | Out-Null }
        if ($retentionHold) { $excludedReasons.Add('Retention hold is enabled') | Out-Null }
        if ($isServiceAccount) { $excludedReasons.Add('Service or privileged-account signal is present') | Out-Null }

        $capacityBand = if ($null -eq $mailboxSize) { 'Unknown' } elseif ($mailboxSize -lt 45) { 'Below 45 GB' } elseif ($mailboxSize -lt 50) { '45 to below 50 GB' } else { '50 GB or more' }
        $status = if ($excludedReasons.Count -gt 0) { 'Excluded' } elseif ($mailboxSize -lt 45) { 'Strong candidate' } else { 'Capacity review' }
        $severity = if ($status -eq 'Strong candidate') { 'Opportunity' } elseif ($status -eq 'Capacity review') { 'Warning' } else { 'Information' }
        $recommendation = if ($status -eq 'Strong candidate') {
            'Validate ownership and dependencies, convert the user mailbox to a shared mailbox, then remove or reuse the base license when contractually possible.'
        } elseif ($status -eq 'Capacity review') {
            'Review mailbox growth and reduce size below the operational safety threshold before considering conversion and license removal.'
        } else {
            'Do not use this mailbox as a license-removal candidate until every exclusion reason is resolved.'
        }
        $mailboxSizeOutput = if ($null -eq $mailboxSize) { '' } else { [math]::Round($mailboxSize, 2) }
        $archiveSizeOutput = if ($null -eq $archiveSize) { '' } else { [math]::Round($archiveSize, 2) }

        $rows.Add((New-SmartFinOpsExchangeOptimizationRow `
            -FindingType 'SharedMailboxConversion' `
            -Identity ([string](Get-RowPropertyValue -Row $mailbox -Names @('UserPrincipalName', 'PrimarySmtpAddress'))) `
            -Severity $severity -IssueType 'License realization path' -Status $status -MailboxLocation 'Exchange Online' `
            -Enabled $exoEnabled -Details "Licensed disabled UserMailbox with $($delegates.Count) configured delegate(s)." `
            -RecipientTypeDetails $recipientType -AccountStateStatus $accountStateStatus `
            -M365AccountEnabled $m365Enabled -ADEnabled $adEnabled -EXOAccountEnabled $exoEnabled `
            -HasDelegates $true -DelegateCount $delegates.Count -FullAccess $fullAccess -SendAs $sendAs -GrantSendOnBehalfTo $sendOnBehalf `
            -MailboxSizeGB $mailboxSizeOutput -CapacityBand $capacityBand -HasArchive $hasArchive -ArchiveSizeGB $archiveSizeOutput `
            -LitigationHoldEnabled $litigationHold -RetentionHoldEnabled $retentionHold -IsLikelyServiceAccount $isServiceAccount `
            -AssignedSkus $assignedSkus -CurrentBaseLicense ([string](Get-RowPropertyValue -Row $decision -Names @('CurrentBaseLicense'))) `
            -IndicativeMonthlyValueEUR (Get-RowPropertyValue -Row $decision -Names @('CurrentMonthlyPriceEUR')) -Currency $Currency `
            -ExcludedReasons ($excludedReasons -join ' | ') -Recommendation $recommendation `
            -Guardrail 'Review retention, legal requirements, forwarding, application access, ownership, delegate need, mailbox growth, and the commercial contract. Never convert or remove a license automatically.'
        )) | Out-Null
    }

    return @($rows | Sort-Object @{ Expression = { switch ($_.Status) { 'Strong candidate' { 1 } 'Capacity review' { 2 } default { 3 } } } }, MailboxSizeGB, Identity)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBhTqNAmRo99DpH
# Vi8JNBXG9gy7N6OckFbzzlg7ITzdc6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKPG7p8pfSzYZe7IRK3H3p6QVLSHWuhRlzeZjadNjAd3MA0GCSqG
# SIb3DQEBAQUABIIBgIzhL6ImglD/mV7WQj9ZWrhrM0EjnrPefIzRrn/uZPhfTpby
# dnLbHDj2kI0soUMhs1rhebmPwtx46S/qEllMt/D+xPeZ3WG6kTMC+zNMFSoOtH+k
# /NEQQ3RsVvYwMmXVIVg6R7piQMHTqcDR5H+hmEX7LrbVy0vLeRHyaUX73y/8MovD
# Ic66S8Y45JP4Y34EVlv1IAM8LcHU4b8BLKHRZ5yM6fkilLlXgBZiEJbmG4UHgMov
# RqRUBYHscT8nW9HrkbXWUvqX4bmPDXMNpcUmX9ryqn55GLDs2/INWBnxVfJbLeLK
# dtozveOdjYfO2KLy/Gc3nkc/+ZBL+FTljGoO57AH7W7JUO1XqOCFM/pvjeLduTJ2
# VVT/O+1LZA0EDF495KMdZ6+dbWGFd3qqXabQcBgCdFVHrEvMSx3Xlwl8hI4lJ1il
# ZaaFhyGa9hK4NScpVVnMLVyeaFIB1l1K4Gr1uJf5XvUXI8R8ajtjuFJr7pM70NK5
# zXDa75Wp50OSsffy2qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDkwODA4
# MjlaMC8GCSqGSIb3DQEJBDEiBCBFdzxIs59GEonU40pimRfb7pNW2HZXM/aoBYQ2
# HcfzwzANBgkqhkiG9w0BAQEFAASCAgA5/MEhbxUPVIksxguoygHEzlsKYXE0FOZ3
# dObng71+gl56SWOMUqiN1PXxXbHnGm/KhgxrJ3M5wMYvL/RhPsE6ClRo9E9+yNkC
# fOi5eVIXRDCZgdRrULy8hfJT23F9NC5t0WLSMhYeWYBH4ihT2en9Sfi7zzidt1WA
# VLzuUZUDpTD9stvt4sbnOVz4fLN2JktX7HFkT7pxz4WVF3HKk8TlbnvgE7DO8iGb
# C7UFrJ1tcBImfQO6rOGe6d89MGQ9kulnc0cULVl+HSgceFtmbgs/1nlDDvT3hDvo
# 5ef+EhyWgiEOoewQEQBiu/Zb6gO3P/7tcpGCy2N1R0jVtmg0AAqaSCh5pR6Ls3vD
# nI3qyJlScg7qYlSLAeyDizDn/ur2368/y3q3e/O0/Zduc4hBOpx6NNQJhnr+eyGa
# lnGsKr1h1AOGIO19dQTmQ+STZVFdXgpFDiw7O7G5L1BYW0SL0JKcb5jPRxsopXCf
# JFPk+SyrPmsJ1jbrUMXDytNnHhxAoMZNdUo9PbKmrW66QHogFxVa2krtyfmLIGvZ
# tiE06YCp2mRf/47zK3ghiTU3uPcyXHnHrtY4P+k1hVaIVCq2w9pBEoqwRC/QLeWd
# D0aTu5ccjlrx76MbnKnfGNKFwQTad6uEY4pO4ekGwmHiQjpZp8+9uf9bMRwy9yYD
# Oo5Zpd+ASQ==
# SIG # End signature block
