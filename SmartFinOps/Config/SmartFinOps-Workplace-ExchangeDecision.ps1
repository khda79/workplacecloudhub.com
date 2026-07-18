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
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

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
    $number = [decimal]0
    if ([decimal]::TryParse($text, [Globalization.NumberStyles]::Number, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    if ([decimal]::TryParse($text, [Globalization.NumberStyles]::Number, [Globalization.CultureInfo]::CurrentCulture, [ref]$number)) { return $number }
    return $null
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
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
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
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

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
