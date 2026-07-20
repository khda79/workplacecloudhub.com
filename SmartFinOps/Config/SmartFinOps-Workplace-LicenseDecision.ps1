Set-StrictMode -Version 2.0

function Get-SmartFinOpsUpnKey {
    [CmdletBinding()]
    param(
        [AllowNull()]$Row,
        [Parameter(Mandatory)][string[]]$Names
    )

    $value = [string](Get-RowPropertyValue -Row $Row -Names $Names)
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    return $value.Trim().ToLowerInvariant()
}

function Get-SmartFinOpsEvidenceRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Map,
        [Parameter(Mandatory)][string]$UserKey
    )

    if (-not $Map.ContainsKey($UserKey)) {
        $Map[$UserKey] = [pscustomobject]@{
            LatestDetailedActivity = $null
            RecentWorkloads = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            RecentM365Services = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            HasDesktopAppsActivation = $false
            HasRecentManagedDevice = $false
            HasMailboxStorageEvidence = $false
            MailboxStorageBytes = [decimal]0
            OneDriveStorageBytes = [decimal]0
            HasArchiveMailbox = $false
        }
    }
    return $Map[$UserKey]
}

function Add-SmartFinOpsEvidenceDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Evidence,
        [AllowNull()]$Value,
        [Parameter(Mandatory)][string]$Workload,
        [Parameter(Mandatory)][datetime]$RecentCutoff
    )

    $date = ConvertTo-DateTimeOrNull $Value
    if ($null -eq $date) { return }
    if ($null -eq $Evidence.LatestDetailedActivity -or $date -gt $Evidence.LatestDetailedActivity) {
        $Evidence.LatestDetailedActivity = $date
    }
    if ($date -ge $RecentCutoff) {
        [void]$Evidence.RecentWorkloads.Add($Workload)
        [void]$Evidence.RecentM365Services.Add($Workload)
    }
}

function New-SmartFinOpsUserEvidenceMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$UserActivityRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$MailboxUsageRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$OneDriveUsageRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AppsActivationRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TeamsActivityRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$EmailActivityRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SharePointActivityRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TeamsDeviceUsageRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$CopilotUsageRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TeamsPhoneUsageRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$IntuneDeviceRows,
        [Parameter(Mandatory)][datetime]$RecentCutoff
    )

    $map = @{}

    foreach ($row in $UserActivityRows) {
        if ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $row -Names @('IsDeleted', 'Is Deleted'))) -eq $true) { continue }
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('UserPrincipalName', 'User Principal Name')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        $workload = [string](Get-RowPropertyValue -Row $row -Names @('LastActivityWorkload'))
        if ([string]::IsNullOrWhiteSpace($workload)) { $workload = 'M365 aggregate' }
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('LastActivityDate')) -Workload $workload -RecentCutoff $RecentCutoff
    }

    foreach ($row in $MailboxUsageRows) {
        if ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $row -Names @('Is Deleted'))) -eq $true) { continue }
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('User Principal Name', 'UserPrincipalName')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('Last Activity Date')) -Workload 'Exchange mailbox' -RecentCutoff $RecentCutoff
        $storageValue = Get-RowPropertyValue -Row $row -Names @('Storage Used (Byte)')
        if (-not [string]::IsNullOrWhiteSpace([string]$storageValue)) {
            $evidence.HasMailboxStorageEvidence = $true
            $storage = ConvertTo-DecimalOrZero $storageValue
            if ($storage -gt $evidence.MailboxStorageBytes) { $evidence.MailboxStorageBytes = $storage }
        }
        if ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $row -Names @('Has Archive'))) -eq $true) { $evidence.HasArchiveMailbox = $true }
    }

    foreach ($row in $OneDriveUsageRows) {
        if ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $row -Names @('Is Deleted'))) -eq $true) { continue }
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('Owner Principal Name', 'User Principal Name', 'UserPrincipalName')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('Last Activity Date')) -Workload 'OneDrive' -RecentCutoff $RecentCutoff
        $storage = ConvertTo-DecimalOrZero (Get-RowPropertyValue -Row $row -Names @('Storage Used (Byte)'))
        if ($storage -gt $evidence.OneDriveStorageBytes) { $evidence.OneDriveStorageBytes = $storage }
    }

    foreach ($row in $AppsActivationRows) {
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('User Principal Name', 'UserPrincipalName')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        $productType = [string](Get-RowPropertyValue -Row $row -Names @('Product Type'))
        $lastActivation = Get-RowPropertyValue -Row $row -Names @('Last Activated Date')
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value $lastActivation -Workload 'Microsoft 365 Apps' -RecentCutoff $RecentCutoff
        $desktopCount = (ConvertTo-DecimalOrZero (Get-RowPropertyValue -Row $row -Names @('Windows'))) +
            (ConvertTo-DecimalOrZero (Get-RowPropertyValue -Row $row -Names @('Mac')))
        if ($productType -match 'MICROSOFT 365 APPS FOR ENTERPRISE' -and $desktopCount -gt 0) {
            $evidence.HasDesktopAppsActivation = $true
            [void]$evidence.RecentWorkloads.Add('Desktop applications')
        }
    }

    foreach ($row in $TeamsActivityRows) {
        if ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $row -Names @('Is Deleted'))) -eq $true) { continue }
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('User Principal Name', 'UserPrincipalName')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('Last Activity Date')) -Workload 'Teams' -RecentCutoff $RecentCutoff
    }

    foreach ($row in $EmailActivityRows) {
        if ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $row -Names @('Is Deleted'))) -eq $true) { continue }
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('User Principal Name', 'UserPrincipalName')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('Last Activity Date')) -Workload 'Email' -RecentCutoff $RecentCutoff
    }

    foreach ($row in $SharePointActivityRows) {
        if ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $row -Names @('Is Deleted', 'IsDeleted'))) -eq $true) { continue }
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('User Principal Name', 'UserPrincipalName')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('Last Activity Date', 'LastActivityDate')) -Workload 'SharePoint' -RecentCutoff $RecentCutoff
    }

    foreach ($row in $TeamsDeviceUsageRows) {
        if ((ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $row -Names @('Is Deleted', 'IsDeleted'))) -eq $true) { continue }
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('User Principal Name', 'UserPrincipalName')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('Last Activity Date', 'LastActivityDate')) -Workload 'Teams device usage' -RecentCutoff $RecentCutoff
    }

    foreach ($row in $CopilotUsageRows) {
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('UserPrincipalName', 'User Principal Name')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('LastActivityDate', 'Microsoft365CopilotLastActivityDate')) -Workload 'Microsoft 365 Copilot' -RecentCutoff $RecentCutoff
    }

    foreach ($row in $TeamsPhoneUsageRows) {
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('UserPrincipalName', 'User Principal Name')
        if (-not $key) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        Add-SmartFinOpsEvidenceDate -Evidence $evidence -Value (Get-RowPropertyValue -Row $row -Names @('LastCallDate', 'Last Call Date')) -Workload 'Teams Phone' -RecentCutoff $RecentCutoff
    }

    foreach ($row in $IntuneDeviceRows) {
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('Primary user UPN', 'Primary user email address', 'UserPrincipalName')
        if (-not $key) { continue }
        $lastCheckIn = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $row -Names @('Last check-in', 'LastSyncDateTime', 'Last check in'))
        if ($null -eq $lastCheckIn -or $lastCheckIn -lt $RecentCutoff) { continue }
        $evidence = Get-SmartFinOpsEvidenceRecord -Map $map -UserKey $key
        $evidence.HasRecentManagedDevice = $true
        [void]$evidence.RecentWorkloads.Add('Active Intune device')
    }

    return $map
}

function New-SmartFinOpsUserLicenseDecisionRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LicenseRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$M365Users,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ADUsers,
        [Parameter(Mandatory)][hashtable]$EvidenceMap,
        [Parameter(Mandatory)][datetime]$RecentCutoff,
        [Parameter(Mandatory)][AllowNull()]$PriceModel
    )

    $m365ByUpn = @{}
    foreach ($row in $M365Users) {
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('User principal name', 'UserPrincipalName')
        if ($key) { $m365ByUpn[$key] = $row }
    }
    $adByUpn = @{}
    foreach ($row in $ADUsers) {
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('UserPrincipalName', 'User principal name')
        if ($key) { $adByUpn[$key] = $row }
    }
    $licenseByUpn = @{}
    foreach ($row in $LicenseRows) {
        $key = Get-SmartFinOpsUpnKey -Row $row -Names @('User principal name', 'UserPrincipalName')
        if (-not $key) { continue }
        if (-not $licenseByUpn.ContainsKey($key)) { $licenseByUpn[$key] = [System.Collections.Generic.List[object]]::new() }
        $licenseByUpn[$key].Add($row)
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($key in ($licenseByUpn.Keys | Sort-Object)) {
        $licenses = @($licenseByUpn[$key])
        $skus = @($licenses | ForEach-Object { [string](Get-RowPropertyValue -Row $_ -Names @('SkuPartNumber', 'SKU part number')) } | Where-Object { $_ } | Sort-Object -Unique)
        $upn = [string](Get-RowPropertyValue -Row $licenses[0] -Names @('User principal name', 'UserPrincipalName'))
        $currentBaseSku = if ('SPE_E3' -in $skus) { 'SPE_E3' } elseif ('SPE_F1' -in $skus) { 'SPE_F1' } else { '' }
        $currentPlan = if ($currentBaseSku -eq 'SPE_E3') { 'M365 E3' } elseif ($currentBaseSku -eq 'SPE_F1') { 'M365 F3' } else { 'Other or no base suite' }
        $m365User = if ($m365ByUpn.ContainsKey($key)) { $m365ByUpn[$key] } else { $null }
        $adUser = if ($adByUpn.ContainsKey($key)) { $adByUpn[$key] } else { $null }
        $evidence = if ($EvidenceMap.ContainsKey($key)) { $EvidenceMap[$key] } else { Get-SmartFinOpsEvidenceRecord -Map $EvidenceMap -UserKey $key }

        $personaRaw = [string](Get-RowPropertyValue -Row $adUser -Names @('M365LicenseTargetPersona'))
        $targetPersona = if ($personaRaw -match '(?i)^M365\s*E3$|^E3$') {
            'M365 E3'
        }
        elseif ($personaRaw -match '(?i)^M365\s*F3$|^F3$') {
            'M365 F3'
        }
        elseif ($personaRaw -match '(?i)^None$|^No\s*license$') {
            'None'
        }
        else {
            'Undetermined'
        }
        $accountType = [string](Get-RowPropertyValue -Row $adUser -Names @('AccountType'))
        $m365Enabled = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $m365User -Names @('AccountEnabled'))
        if ($null -eq $m365Enabled) {
            $blocked = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $m365User -Names @('Block credential'))
            if ($null -ne $blocked) { $m365Enabled = -not $blocked }
        }
        $adEnabled = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $adUser -Names @('Enabled'))
        $lastSignIn = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $m365User -Names @('LastSignInDateTime', 'Last sign-in time'))
        $lastAdLogon = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $adUser -Names @('LastLogonDate', 'Last logon date'))
        $activityDates = @($evidence.LatestDetailedActivity, $lastSignIn, $lastAdLogon) | Where-Object { $null -ne $_ }
        $latestActivity = @($activityDates | Sort-Object -Descending | Select-Object -First 1)
        $latestActivityDate = if ($latestActivity.Count -gt 0) { [datetime]$latestActivity[0] } else { $null }
        $recentM365ServiceEvidence = @(@($evidence.RecentM365Services) | Sort-Object)
        $recentTechnicalEvidence = New-Object System.Collections.Generic.List[string]
        if ($lastSignIn -and $lastSignIn -ge $RecentCutoff) { $recentTechnicalEvidence.Add('Entra sign-in') | Out-Null }
        if ($lastAdLogon -and $lastAdLogon -ge $RecentCutoff) { $recentTechnicalEvidence.Add('AD logon') | Out-Null }
        if ($evidence.HasRecentManagedDevice) { $recentTechnicalEvidence.Add('Active Intune device') | Out-Null }
        $hasRecentM365ServiceActivity = $recentM365ServiceEvidence.Count -gt 0
        $hasRecentTechnicalPresence = $recentTechnicalEvidence.Count -gt 0
        $hasRecentObservedActivity = $hasRecentM365ServiceActivity -or $hasRecentTechnicalPresence
        $hasRecentActivity = ($latestActivityDate -and $latestActivityDate -ge $RecentCutoff) -or $evidence.HasDesktopAppsActivation -or $evidence.HasRecentManagedDevice

        $mailboxOverF3Limit = $evidence.MailboxStorageBytes -gt 2GB
        $oneDriveOverF3Limit = $evidence.OneDriveStorageBytes -gt 2GB
        $hasF3TechnicalBlocker = $mailboxOverF3Limit -or $oneDriveOverF3Limit -or $evidence.HasDesktopAppsActivation
        $f3TechnicalStatus = if ($hasF3TechnicalBlocker) { 'Blocked' } else { 'No observed blocker' }
        $isSharedMailbox = $accountType -match 'Shared Mailbox|Room Mailbox'
        $isSpecialAccount = $accountType -match 'Service Account|Generic Account|Admin Account|System Account'
        $isDisabled = ($m365Enabled -eq $false) -or ($adEnabled -eq $false)
        $isCurrentE3 = $currentBaseSku -eq 'SPE_E3'
        $isCurrentF3 = $currentBaseSku -eq 'SPE_F1'
        $isUnusedE3F3License = ($isCurrentE3 -or $isCurrentF3) -and (-not $hasRecentObservedActivity)
        $isPossiblyUnusedE3F3License = ($isCurrentE3 -or $isCurrentF3) -and
            $hasRecentTechnicalPresence -and
            (-not $hasRecentM365ServiceActivity) -and
            $evidence.HasMailboxStorageEvidence -and
            $evidence.MailboxStorageBytes -le 100MB
        $isE3WithoutObservedE3Capabilities = $isCurrentE3 -and
            $evidence.HasMailboxStorageEvidence -and
            $evidence.MailboxStorageBytes -lt 100GB -and
            (-not $evidence.HasDesktopAppsActivation)

        $recommended = 'Clarify license requirement'
        $confidence = 'Low'
        $basis = 'The current base license or M365LicenseTargetPersona does not support one of the approved decision paths.'
        $frontlineEligibilityStatus = 'Not applicable'

        if ($isSharedMailbox) {
            if ($evidence.MailboxStorageBytes -gt 50GB -or $evidence.HasArchiveMailbox) {
                $recommended = 'Keep appropriate license - shared mailbox'
                $confidence = 'High'
                $basis = 'Shared mailbox with more than 50 GB or an active archive; an appropriate license is still required.'
            }
            else {
                $recommended = 'Separate review - shared mailbox'
                $confidence = 'Medium'
                $basis = 'A shared mailbox must not be treated as a named user; validate size, archive, and usage before removing a license.'
            }
        }
        elseif ($isSpecialAccount) {
            $recommended = 'Separate review - special account'
            $confidence = 'Medium'
            $basis = "Account type '$accountType'; validate technical usage and required services before making any change."
        }
        elseif (($isCurrentE3 -or $isCurrentF3) -and $targetPersona -eq 'None') {
            if ($isDisabled) {
                $recommended = 'No license - candidate'
                $confidence = 'High'
                $basis = 'M365LicenseTargetPersona is None and the user account is disabled or blocked; confirm retention, departure, ownership, and regulatory requirements before removing the license.'
            }
            elseif (-not $hasRecentActivity) {
                $recommended = 'No license - review'
                $confidence = 'Medium'
                $basis = 'M365LicenseTargetPersona is None and no recent activity was observed across M365, detailed usage, AD/Entra sign-ins, or Intune devices.'
            }
            else {
                $recommended = 'Target persona conflict - active user review'
                $confidence = 'Medium'
                $basis = 'M365LicenseTargetPersona is None, but recent user activity or a managed-device signal was observed; do not remove the license until the persona and business need are reconciled.'
            }
        }
        elseif (($isCurrentE3 -or $isCurrentF3) -and $isDisabled) {
            $recommended = 'Account state conflict - target persona review'
            $confidence = 'Medium'
            $basis = "The account is disabled or blocked, but M365LicenseTargetPersona is '$targetPersona'; reconcile the persona, retention, and ownership before any license decision."
        }
        elseif ($isCurrentE3 -and $targetPersona -eq 'M365 F3') {
            $frontlineEligibilityStatus = 'Required - not proven by telemetry'
            if ($hasF3TechnicalBlocker) {
                $recommended = 'Keep M365 E3 - F3 technical blocker'
                $confidence = 'High'
                $basis = 'M365LicenseTargetPersona is F3, but a desktop application activation or Exchange/OneDrive storage above the 2 GB F3 limit was observed.'
            }
            elseif (-not $hasRecentActivity) {
                $recommended = 'Potential M365 F3 - activity and eligibility review'
                $confidence = 'Low'
                $basis = 'M365LicenseTargetPersona is F3 and no F3 technical blocker was observed, but recent activity is not available; validate both business need and Frontline eligibility.'
            }
            else {
                $recommended = 'Potential M365 F3 - Frontline eligibility required'
                $confidence = 'Medium'
                $basis = 'M365LicenseTargetPersona is F3, recent activity is observed, and no desktop or 2 GB storage blocker was found; documented Frontline eligibility is still required before changing the license.'
            }
        }
        elseif ($isCurrentF3 -and $targetPersona -eq 'M365 E3') {
            $recommended = 'M365 E3 capability review'
            $confidence = if ($hasF3TechnicalBlocker) { 'High' } else { 'Medium' }
            $basis = if ($hasF3TechnicalBlocker) {
                'M365LicenseTargetPersona is E3 and an observable F3 technical blocker exists; validate the required E3 capabilities and service impact.'
            }
            else {
                'M365LicenseTargetPersona is E3 while the current base license is F3; validate capability, quality, and user-experience requirements before any upgrade.'
            }
        }
        elseif ($isCurrentF3 -and $targetPersona -eq 'M365 F3') {
            if ($hasF3TechnicalBlocker) {
                $recommended = 'M365 F3 technical conflict - review'
                $confidence = 'High'
                $basis = 'The current license and target persona are F3, but a desktop application activation or Exchange/OneDrive storage above the 2 GB F3 limit was observed.'
            }
            elseif (-not $hasRecentActivity) {
                $recommended = 'Keep M365 F3 - activity review'
                $confidence = 'Low'
                $basis = 'The current license and target persona are F3, with no observed technical blocker, but recent activity was not found.'
            }
            else {
                $recommended = 'Keep current M365 F3 - no observed technical conflict'
                $confidence = 'Medium'
                $basis = 'The current license and M365LicenseTargetPersona are F3, recent activity is observed, and no desktop or 2 GB storage blocker was found.'
            }
        }
        elseif ($isCurrentE3 -and $targetPersona -eq 'M365 E3') {
            if (-not $hasRecentActivity) {
                $recommended = 'Keep M365 E3 - activity review'
                $confidence = 'Low'
                $basis = 'The current license and target persona are E3, but recent activity was not found; confirm the continuing business requirement.'
            }
            else {
                $recommended = 'Keep current M365 E3'
                $confidence = if ($hasF3TechnicalBlocker) { 'High' } else { 'Medium' }
                $basis = if ($hasF3TechnicalBlocker) {
                    'The current license and target persona are E3, and desktop or storage evidence supports an E3 capability requirement.'
                }
                else {
                    'The current license and M365LicenseTargetPersona are E3 and recent activity is observed.'
                }
            }
        }

        $currentPrice = if ($currentBaseSku) { Get-MonthlySkuPrice -PriceModel $PriceModel -SkuPartNumber $currentBaseSku } else { $null }
        $recommendedSku = if ($recommended -in @('Keep current M365 E3', 'Keep M365 E3 - activity review', 'Keep M365 E3 - F3 technical blocker', 'M365 E3 capability review')) {
            'SPE_E3'
        }
        elseif ($recommended -in @('Keep current M365 F3 - no observed technical conflict', 'Keep M365 F3 - activity review', 'Potential M365 F3 - Frontline eligibility required')) {
            'SPE_F1'
        }
        else {
            ''
        }
        $recommendedPrice = if ($recommendedSku) { Get-MonthlySkuPrice -PriceModel $PriceModel -SkuPartNumber $recommendedSku } elseif ($recommended -match '^No license') { [decimal]0 } else { $null }
        $monthlyDelta = if ($null -ne $currentPrice -and $null -ne $recommendedPrice) { [math]::Round(([decimal]$currentPrice - [decimal]$recommendedPrice), 2) } else { $null }
        $decisionClass = if ($recommended -eq 'No license - candidate') {
            'Recommended'
        }
        elseif ($recommended -eq 'Potential M365 F3 - Frontline eligibility required') {
            'Conditional'
        }
        elseif ($recommended -match '(?i)review' -or $recommended -match '^Separate review') {
            'Review'
        }
        else {
            'Keep'
        }

        $rows.Add([pscustomobject]@{
            RunId = $script:RunId
            UserPrincipalName = $upn
            AccountType = $accountType
            CurrentBaseLicense = $currentPlan
            CurrentBaseSku = $currentBaseSku
            AssignedSkuPartNumbers = ($skus -join ' | ')
            TargetPersonaRaw = $personaRaw
            TargetPersona = $targetPersona
            RecommendedLicense = $recommended
            DecisionClass = $decisionClass
            DecisionConfidence = $confidence
            DecisionBasis = $basis
            F3TechnicalStatus = $f3TechnicalStatus
            FrontlineEligibilityStatus = $frontlineEligibilityStatus
            LatestKnownActivity = if ($latestActivityDate) { $latestActivityDate.ToString('yyyy-MM-dd') } else { '' }
            RecentEvidence = (@($evidence.RecentWorkloads) | Sort-Object) -join ' | '
            RecentM365ServiceEvidence = $recentM365ServiceEvidence -join ' | '
            RecentTechnicalEvidence = $recentTechnicalEvidence.ToArray() -join ' | '
            HasRecentM365ServiceActivity = $hasRecentM365ServiceActivity
            HasRecentTechnicalPresence = $hasRecentTechnicalPresence
            HasDesktopAppsActivation = $evidence.HasDesktopAppsActivation
            HasRecentManagedDevice = $evidence.HasRecentManagedDevice
            HasMailboxStorageEvidence = $evidence.HasMailboxStorageEvidence
            MailboxStorageGB = [math]::Round([double]$evidence.MailboxStorageBytes / 1GB, 2)
            OneDriveStorageGB = [math]::Round([double]$evidence.OneDriveStorageBytes / 1GB, 2)
            IsUnusedE3F3License = $isUnusedE3F3License
            IsPossiblyUnusedE3F3License = $isPossiblyUnusedE3F3License
            IsE3WithoutObservedE3Capabilities = $isE3WithoutObservedE3Capabilities
            CurrentMonthlyPriceEUR = $currentPrice
            RecommendedMonthlyPriceEUR = $recommendedPrice
            IndicativeMonthlyDifferenceEUR = $monthlyDelta
        }) | Out-Null
    }

    return $rows.ToArray()
}




# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDEi4nJzkZXtK22
# kjmcg//G25xzrGxN4dTI3l6rucfakKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIPCX27EtNeyYDAipLhgN/6u7TWf4b/JiRZJIxRi4vVvrMA0GCSqG
# SIb3DQEBAQUABIIBgIIssZrqljGANb6RkNLo9jsQ7HTy8r0S/qplx1EKISocQtCx
# BnC+qP31tv5eWnKMHpfBYO84vuF+HD6CN3sEM9AlLPuiv/9vRDr0GmDMVjrfJUhE
# DP4FDis9AR++hfuYdhEivtUujPfXrtY45uCICdBDfr9uUoBE2xSUwdTevSGlUnUx
# ffCOcXV2QZi3eLOal0cGyRelXF1Q7J4Mkt/BTQ27/VkIoGBrbYoHNcDTcmU/WfNe
# rQzGv+W6tvtcHe3DzkbsILHp4NWG79235oyj4Nz2gjQNhzY/f2w1+bkP3OkVBIus
# ETy+PUjF7l8hur/ZKzB3Ajdr/V1zbmxm6fpf4/S9cu+pP8zUPpd4ptGAYLrkvjQx
# uR1UBz4BaS0DaA1ceeo+AI94A6imYgCl6CFwVp8dYJeJb1s1a718Ci5dDZcjZhs4
# /UeEpeK+lmWstcK3yBbBZoWNcbh5lc/59LZGdv9spzPDhpuQAHHi1ncVKQar7yYC
# +z1ckO9DjIvQoTfgGKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjAxMTI3
# MDZaMC8GCSqGSIb3DQEJBDEiBCDMEJ+Kd0GdOlpJebvUheEQ/iVJjq+w0prKbLdE
# W8Jp2DANBgkqhkiG9w0BAQEFAASCAgDJeobM5f9eoeqN8GirqsPeuQY6uPliddKs
# P2JexgkIynArQ3poZIVUoSKoWx2UCypWptzItsM7CGGshS+PunMXUiyWxCSJYQLw
# X46EfATTPvSKdOcfLOaZfmVgn4lVmp2zkI1sBRfS8sfo/BwkmJwInhQ/hUrxvPZu
# /SQo3ebKT48/u3HXianK3cYu11VUtPvp314PrDI0B4BPSx/l9WYDca0dNH/p5vYa
# L3taodQeArr3ndKuvtY1oGEqtQwweehYn4fTei4r5xbE/nDa4I2PjcMdQTbXKGSU
# qJ3yGsTsyizVG1C8HEKM9foj3vHKkgDIztLbylMHeY0BhtAsazmoJ0VwkRBUdH1o
# R1qBIrUuuyI7vjysaFyYHHAu94xg2CRVW6VXfA9P6cjoYTlysGipJWZGSRi/uQv9
# fJnhJRr8D6zYbrcC1TGrowQNp/qkigdVqodWuFQ7mQoYEvt6VYe48lGdAO5wgf+E
# tJv4mBhVrsoIQ8XSXpGZi26zW5JE4SsGt9QJuOoNnqsUEOnv6M6WJUp7OeIBjmMh
# +w0mVPn2wl5nNxKawItt9a0CSuh3pMIInUw2qnQG/I5NWEq8K607JuB2epPH7ruN
# ECmOF8gV9fIDkJGRB3f/JIO63CaQPKkc9YyhW6rLsc8HT8ld9rUjf0yWtl2psJ61
# vwUwDLmUXg==
# SIG # End signature block
