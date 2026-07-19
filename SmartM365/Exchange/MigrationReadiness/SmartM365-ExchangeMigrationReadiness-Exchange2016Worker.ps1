#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InputPath = '',
    [string]$OutputPath = '',
    [string]$ErrorPath = '',
    [string]$ProgressPath = '',
    [switch]$ValidateOnly,
    [switch]$SelfTest
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'
$workerVersion = '1.11.2'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
[Console]::InputEncoding = $utf8
$OutputEncoding = $utf8
$enabledChecks = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$addressQueryCache = @{}

function Test-WorkerCheckEnabled {
    param([Parameter(Mandatory)][string]$CheckId)
    return $enabledChecks.Count -eq 0 -or $enabledChecks.Contains($CheckId)
}

function Write-WorkerText {
    param([string]$Path, [string]$Text)
    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        [IO.File]::WriteAllText($Path, $Text, $utf8)
    }
}

function Write-WorkerProgress {
    param([int]$Current, [int]$Total, [string]$Message)
    Write-WorkerText -Path $ProgressPath -Text ("{0}|{1}|{2}" -f $Current, $Total, $Message)
}

function Test-WorkerNotFoundError {
    param([string]$Message)
    return -not [string]::IsNullOrWhiteSpace($Message) -and
        $Message -match "(?i)not found|couldn't be found|could not be found|does not exist|cannot be found"
}

function Invoke-WorkerCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{},
        [switch]$NotFoundIsEmpty
    )
    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{ Success = $false; Rows = @(); ErrorMessage = "$Name is unavailable." }
    }
    try {
        return [pscustomobject]@{
            Success = $true
            Rows = @(& $Name @Parameters -ErrorAction Stop)
            ErrorMessage = ''
        }
    }
    catch {
        if ($NotFoundIsEmpty -and (Test-WorkerNotFoundError -Message $_.Exception.Message)) {
            return [pscustomobject]@{ Success = $true; Rows = @(); ErrorMessage = '' }
        }
        return [pscustomobject]@{ Success = $false; Rows = @(); ErrorMessage = $_.Exception.Message }
    }
}

function ConvertTo-TextArray {
    param($Value)
    return @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function ConvertTo-MailboxEvidence {
    param($Mailbox)
    if ($null -eq $Mailbox) { return $null }
    return [pscustomobject][ordered]@{
        Identity = [string]$Mailbox.Identity
        DistinguishedName = [string]$Mailbox.DistinguishedName
        PrimarySmtpAddress = [string]$Mailbox.PrimarySmtpAddress
        WindowsEmailAddress = [string]$Mailbox.WindowsEmailAddress
        ExternalEmailAddress = [string]$Mailbox.ExternalEmailAddress
        EmailAddresses = @(ConvertTo-TextArray $Mailbox.EmailAddresses)
        RecipientTypeDetails = [string]$Mailbox.RecipientTypeDetails
        ExchangeGuid = [string]$Mailbox.ExchangeGuid
        ArchiveGuid = [string]$Mailbox.ArchiveGuid
        ArchiveStatus = [string]$Mailbox.ArchiveStatus
        ArchiveState = [string]$Mailbox.ArchiveState
        LegacyExchangeDN = [string]$Mailbox.LegacyExchangeDN
        Database = [string]$Mailbox.Database
        LitigationHoldEnabled = [bool]$Mailbox.LitigationHoldEnabled
        InPlaceHolds = @(ConvertTo-TextArray $Mailbox.InPlaceHolds)
        GrantSendOnBehalfTo = @(ConvertTo-TextArray $Mailbox.GrantSendOnBehalfTo)
        ForwardingSmtpAddress = [string]$Mailbox.ForwardingSmtpAddress
        ForwardingAddress = [string]$Mailbox.ForwardingAddress
        DeliverToMailboxAndForward = [bool]$Mailbox.DeliverToMailboxAndForward
        ModerationEnabled = [bool]$Mailbox.ModerationEnabled
        AcceptMessagesOnlyFrom = @(ConvertTo-TextArray $Mailbox.AcceptMessagesOnlyFrom)
        AcceptMessagesOnlyFromDLMembers = @(ConvertTo-TextArray $Mailbox.AcceptMessagesOnlyFromDLMembers)
        RejectMessagesFrom = @(ConvertTo-TextArray $Mailbox.RejectMessagesFrom)
        RejectMessagesFromDLMembers = @(ConvertTo-TextArray $Mailbox.RejectMessagesFromDLMembers)
        UseDatabaseQuotaDefaults = $Mailbox.UseDatabaseQuotaDefaults
        ProhibitSendReceiveQuota = [string]$Mailbox.ProhibitSendReceiveQuota
        LargeItemCount = if ($Mailbox.PSObject.Properties['LargeItemCount']) { [string]$Mailbox.LargeItemCount } else { '' }
        LargeItemCollectionStatus = if ($Mailbox.PSObject.Properties['LargeItemCollectionStatus']) { [string]$Mailbox.LargeItemCollectionStatus } else { 'NotCollected' }
    }
}

function ConvertTo-RecipientEvidence {
    param($Recipient)
    if ($null -eq $Recipient) { return $null }
    return [pscustomobject][ordered]@{
        Identity = [string]$Recipient.Identity
        DistinguishedName = [string]$Recipient.DistinguishedName
        PrimarySmtpAddress = [string]$Recipient.PrimarySmtpAddress
        WindowsEmailAddress = [string]$Recipient.WindowsEmailAddress
        ExternalEmailAddress = [string]$Recipient.ExternalEmailAddress
        EmailAddresses = @(ConvertTo-TextArray $Recipient.EmailAddresses)
        RecipientType = [string]$Recipient.RecipientType
        RecipientTypeDetails = [string]$Recipient.RecipientTypeDetails
        ExchangeGuid = [string]$Recipient.ExchangeGuid
        ArchiveGuid = [string]$Recipient.ArchiveGuid
        LegacyExchangeDN = [string]$Recipient.LegacyExchangeDN
    }
}

function ConvertTo-StatisticsEvidence {
    param($Statistics)
    if ($null -eq $Statistics) { return $null }
    return [pscustomobject][ordered]@{
        TotalItemSize = [string]$Statistics.TotalItemSize
        TotalDeletedItemSize = [string]$Statistics.TotalDeletedItemSize
        ItemCount = $Statistics.ItemCount
        DisconnectReason = [string]$Statistics.DisconnectReason
        LastLogonTime = $Statistics.LastLogonTime
    }
}

function Get-AddressConflictEvidence {
    param([string]$MailboxAddress, [string[]]$Addresses)
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($address in @($Addresses | ForEach-Object { ([string]$_ -replace '^(?i:smtp:)', '').Trim() } | Where-Object { $_ -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' } | Sort-Object -Unique)) {
        $safeAddress = $address.Replace("'", "''")
        $cacheKey = $address.ToLowerInvariant()
        if ($addressQueryCache.ContainsKey($cacheKey)) {
            $query = $addressQueryCache[$cacheKey]
        }
        else {
            $query = Invoke-WorkerCommand -Name 'Get-Recipient' -Parameters @{
                Filter = "EmailAddresses -eq 'smtp:$safeAddress' -or ExternalEmailAddress -eq 'smtp:$safeAddress'"
                ResultSize = 'Unlimited'
            }
            $addressQueryCache[$cacheKey] = $query
        }
        $owners = New-Object System.Collections.Generic.List[string]
        if ($query.Success) {
            foreach ($owner in @($query.Rows)) {
                $ownerAddress = [string]$owner.PrimarySmtpAddress
                $ownerIdentity = [string]$owner.Identity
                if ($ownerAddress -and $ownerAddress -ine $MailboxAddress) {
                    [void]$owners.Add(("{0} -> {1}" -f $address, $ownerIdentity))
                }
            }
        }
        [void]$items.Add([pscustomobject][ordered]@{
            Address = $address
            Available = [bool]$query.Success
            Conflicts = @($owners | Sort-Object -Unique)
            ErrorMessage = [string]$query.ErrorMessage
        })
    }
    $items.ToArray()
}

function Get-WorkerFailureEvidence {
    param([string]$EmailAddress, [string]$Message)
    $now = Get-Date
    return [pscustomobject][ordered]@{
        EmailAddress = $EmailAddress
        Available = $false
        Source = "Local Exchange 2016 Management Shell on $env:COMPUTERNAME"
        SourceTimestamp = $now
        Message = "Exchange 2016 mailbox collection failed: $Message"
        Mailboxes = @()
        RemoteMailboxes = @()
        MailUsers = @()
        Recipients = @()
        RecipientLookupAvailable = $false
        Statistics = @()
        StatisticsAvailable = $false
        ArchiveStatistics = @()
        ArchiveStatisticsAvailable = $false
        Permissions = @()
        PermissionsAvailable = $false
        HoldDataAvailable = $false
        FolderStatistics = @()
        FolderStatisticsAvailable = $false
        InboxRules = @()
        InboxRulesAvailable = $false
        DatabaseHealth = @()
        DatabaseHealthAvailable = $false
        DatabaseHealthSource = "Local Exchange 2016 Management Shell mailbox database on $env:COMPUTERNAME"
        DatabaseHealthSourceTimestamp = $now
        DeliveryRestrictionsAvailable = $false
        AddressConflicts = @()
    }
}
function Get-HybridEvidence {
    $collectedAt = Get-Date
    $empty = [pscustomobject]@{ Success = $false; Rows = @(); ErrorMessage = 'Check disabled for this run.' }
    $ews = if (Test-WorkerCheckEnabled -CheckId 'HYBRID-MRSPROXY') { Invoke-WorkerCommand -Name 'Get-WebServicesVirtualDirectory' } else { $empty }
    $org = if (Test-WorkerCheckEnabled -CheckId 'HYBRID-AUTODISCOVER-OAUTH') { Invoke-WorkerCommand -Name 'Get-OrganizationConfig' } else { $empty }
    $ioc = if (Test-WorkerCheckEnabled -CheckId 'HYBRID-AUTODISCOVER-OAUTH') { Invoke-WorkerCommand -Name 'Get-IntraOrganizationConnector' } else { $empty }
    $enabledEws = @($ews.Rows | Where-Object { [bool]$_.MRSProxyEnabled })
    $organization = @($org.Rows | Select-Object -First 1)
    $enabledIoc = @($ioc.Rows | Where-Object { [bool]$_.Enabled })
    $oauthEnabled = $organization.Count -eq 1 -and [bool]$organization[0].OAuth2ClientProfileEnabled
    return [pscustomobject][ordered]@{
        MrsProxyAvailable = [bool]$ews.Success
        MrsProxyEnabled = $enabledEws.Count -gt 0
        MrsProxyMessage = if (-not $ews.Success) { $ews.ErrorMessage } elseif ($enabledEws.Count -gt 0) { "MRSProxy enabled on $($enabledEws.Count) EWS virtual directorie(s)." } else { 'No EWS virtual directory has MRSProxy enabled.' }
        MrsProxySource = 'Local Exchange 2016 Management Shell EWS virtual directories'
        MrsProxySourceTimestamp = $collectedAt
        OAuthAvailable = [bool]($org.Success -and $ioc.Success)
        OAuthHealthy = [bool]($oauthEnabled -and $enabledIoc.Count -gt 0)
        OAuthMessage = if (-not $org.Success -or -not $ioc.Success) { (@($org.ErrorMessage, $ioc.ErrorMessage) | Where-Object { $_ } | Select-Object -Unique) -join ' | ' } else { "OAuth2ClientProfileEnabled=$oauthEnabled; EnabledIntraOrganizationConnector=$($enabledIoc.Count -gt 0)." }
        OAuthSource = 'Local Exchange 2016 Management Shell hybrid configuration'
        OAuthSourceTimestamp = $collectedAt
    }
}

if ($SelfTest) {
    $selfTestPath = Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-ExchangeMigrationReadiness-PS5SelfTest-{0}.clixml" -f [guid]::NewGuid().ToString('N'))
    try {
        $items = New-Object System.Collections.Generic.List[object]
        $errors = New-Object System.Collections.Generic.List[object]
        [void]$items.Add([pscustomobject]@{ Name = 'serialization'; Result = 'PASS' })
        [void]$errors.Add([pscustomobject]@{ EmailAddress = 'selftest@example.invalid'; Message = 'Expected self-test error record' })
        $payload = [pscustomobject][ordered]@{
            WorkerVersion = $workerVersion
            Items = $items.ToArray()
            Evidence = @((Get-WorkerFailureEvidence -EmailAddress 'selftest@example.invalid' -Message 'Expected self-test evidence'))
            Errors = $errors.ToArray()
        }
        $payload | Export-Clixml -LiteralPath $selfTestPath -Depth 12 -Force
        $roundTrip = Import-Clixml -LiteralPath $selfTestPath
        if (@($roundTrip.Items).Count -ne 1 -or [string]$roundTrip.Items[0].Result -ne 'PASS' -or @($roundTrip.Evidence).Count -ne 1 -or @($roundTrip.Errors).Count -ne 1) {
            throw 'Windows PowerShell 5.1 CLIXML nested collection round-trip returned an unexpected result.'
        }
        Write-Output "SELFTEST_OK|$workerVersion|$($PSVersionTable.PSVersion)"
        exit 0
    }
    catch {
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
    finally {
        if (Test-Path -LiteralPath $selfTestPath) { Remove-Item -LiteralPath $selfTestPath -Force -ErrorAction SilentlyContinue }
    }
}

try {
    $snapInName = 'Microsoft.Exchange.Management.PowerShell.SnapIn'
    if (-not (Get-PSSnapin -Name $snapInName -ErrorAction SilentlyContinue)) {
        Add-PSSnapin -Name $snapInName -ErrorAction Stop
    }
    foreach ($requiredCommand in @('Set-ADServerSettings', 'Get-Mailbox', 'Get-RemoteMailbox', 'Get-MailUser', 'Get-Recipient', 'Get-MailboxStatistics')) {
        if (-not (Get-Command -Name $requiredCommand -ErrorAction SilentlyContinue)) {
            throw "Required Exchange 2016 command '$requiredCommand' is unavailable after loading $snapInName."
        }
    }
    Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop | Out-Null

    if ($ValidateOnly) {
        Write-Output "VALIDATION_OK|$workerVersion|$env:COMPUTERNAME|$($PSVersionTable.PSVersion)"
        exit 0
    }
    if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
        throw "Mailbox input file not found: $InputPath"
    }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        throw 'OutputPath is required.'
    }

    $workerInput = Import-Clixml -LiteralPath $InputPath
    if ($workerInput -and $workerInput.PSObject.Properties['EmailAddresses']) {
        $emails = @($workerInput.EmailAddresses | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
        foreach ($checkId in @(ConvertTo-TextArray $workerInput.EnabledChecks)) { [void]$enabledChecks.Add($checkId) }
    }
    else {
        # Backward compatibility with the v1.11.1 plain string-array payload.
        $emails = @($workerInput | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
    }

    $startedAt = Get-Date
    $evidence = New-Object System.Collections.Generic.List[object]
    $collectionErrors = New-Object System.Collections.Generic.List[object]
    $databaseCache = @{}
    $index = 0
    foreach ($email in $emails) {
        $index++
        Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Recipient lookup - {0}" -f $email)
        try {
            $safeEmail = $email.Replace("'", "''")
            $identityFilter = "EmailAddresses -eq 'smtp:$safeEmail'"
            $recipientResult = Invoke-WorkerCommand -Name 'Get-Recipient' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' }
            $recipientTypes = @($recipientResult.Rows | ForEach-Object { [string]$_.RecipientTypeDetails } | Where-Object { $_ } | Sort-Object -Unique)
            $successfulEmpty = [pscustomobject]@{ Success = $true; Rows = @(); ErrorMessage = '' }
            $queryAllTypes = -not $recipientResult.Success
            $needMailbox = $queryAllTypes -or @($recipientTypes | Where-Object { $_ -match '^(User|Shared|Room|Equipment|Linked|Discovery)Mailbox$' }).Count -gt 0
            $needRemoteMailbox = $queryAllTypes -or @($recipientTypes | Where-Object { $_ -match '^Remote.*Mailbox$' }).Count -gt 0
            $needMailUser = $queryAllTypes -or $recipientTypes -contains 'MailUser'

            $mailboxResult = if ($needMailbox) { Invoke-WorkerCommand -Name 'Get-Mailbox' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' } } else { $successfulEmpty }
            $remoteMailboxResult = if ($needRemoteMailbox) { Invoke-WorkerCommand -Name 'Get-RemoteMailbox' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' } } else { $successfulEmpty }
            $mailUserResult = if ($needMailUser) { Invoke-WorkerCommand -Name 'Get-MailUser' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' } } else { $successfulEmpty }
            $mailboxes = @($mailboxResult.Rows | ForEach-Object { ConvertTo-MailboxEvidence $_ })
            $remoteMailboxes = @($remoteMailboxResult.Rows | ForEach-Object { ConvertTo-RecipientEvidence $_ })
            $mailUsers = @($mailUserResult.Rows | ForEach-Object { ConvertTo-RecipientEvidence $_ })
            $recipients = @($recipientResult.Rows | ForEach-Object { ConvertTo-RecipientEvidence $_ })

            $statistics = @()
            $statisticsAvailable = $false
            $archiveStatistics = @()
            $archiveStatisticsAvailable = $false
            $folderStatistics = @()
            $folderStatisticsAvailable = $false
            $inboxRules = @()
            $inboxRulesAvailable = $false
            $databaseHealth = @()
            $databaseHealthAvailable = $false
            $permissions = New-Object System.Collections.Generic.List[object]
            $mailboxPermissionSuccess = $false
            $sendAsPermissionSuccess = $false

            if ($mailboxes.Count -eq 1) {
                $identity = $mailboxes[0].Identity
                Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Mailbox statistics - {0}" -f $email)
                $statisticsResult = Invoke-WorkerCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $identity }
                $statistics = @($statisticsResult.Rows | ForEach-Object { ConvertTo-StatisticsEvidence $_ })
                $statisticsAvailable = [bool]$statisticsResult.Success

                if (Test-WorkerCheckEnabled -CheckId 'ARCHIVE-READINESS') {
                    $hasArchive = $mailboxes[0].ArchiveStatus -match '^(?i:Active|HostedPending|Local)$' -or ($mailboxes[0].ArchiveGuid -and $mailboxes[0].ArchiveGuid -notmatch '^0{8}-')
                    if ($hasArchive) {
                        $archiveResult = Invoke-WorkerCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $identity; Archive = $true }
                        $archiveStatistics = @($archiveResult.Rows | ForEach-Object { ConvertTo-StatisticsEvidence $_ })
                        $archiveStatisticsAvailable = [bool]$archiveResult.Success
                    }
                    else { $archiveStatisticsAvailable = $true }
                }

                if ((Test-WorkerCheckEnabled -CheckId 'MAILBOX-RECOVERABLE-ITEMS-QUOTA') -or (Test-WorkerCheckEnabled -CheckId 'MAILBOX-FOLDER-LIMITS')) {
                    Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Folder statistics - {0}" -f $email)
                    $folderResult = Invoke-WorkerCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $identity }
                    $folderStatistics = @($folderResult.Rows | ForEach-Object {
                        [pscustomobject][ordered]@{
                            Name = [string]$_.Name
                            FolderPath = [string]$_.FolderPath
                            FolderType = [string]$_.FolderType
                            ItemsInFolder = $_.ItemsInFolder
                            ItemsInFolderAndSubfolders = $_.ItemsInFolderAndSubfolders
                            FolderSize = [string]$_.FolderSize
                            FolderAndSubfolderSize = [string]$_.FolderAndSubfolderSize
                        }
                    })
                    $folderStatisticsAvailable = [bool]$folderResult.Success
                }

                if (Test-WorkerCheckEnabled -CheckId 'INBOX-FORWARDING-RULES') {
                    Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Inbox rules - {0}" -f $email)
                    $ruleResult = Invoke-WorkerCommand -Name 'Get-InboxRule' -Parameters @{ Mailbox = $identity; IncludeHidden = $true }
                    $inboxRules = @($ruleResult.Rows | ForEach-Object {
                        [pscustomobject][ordered]@{
                            Name = [string]$_.Name
                            Enabled = [bool]$_.Enabled
                            ForwardTo = @(ConvertTo-TextArray $_.ForwardTo)
                            RedirectTo = @(ConvertTo-TextArray $_.RedirectTo)
                            ForwardAsAttachmentTo = @(ConvertTo-TextArray $_.ForwardAsAttachmentTo)
                        }
                    })
                    $inboxRulesAvailable = [bool]$ruleResult.Success
                }

                if ((Test-WorkerCheckEnabled -CheckId 'EXCHANGE-DATABASE-HEALTH') -and $mailboxes[0].Database) {
                    $databaseKey = ([string]$mailboxes[0].Database).Trim().ToLowerInvariant()
                    if (-not $databaseCache.ContainsKey($databaseKey)) {
                        Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Database health - {0}" -f $email)
                        $databaseResult = Invoke-WorkerCommand -Name 'Get-MailboxDatabase' -Parameters @{ Identity = $mailboxes[0].Database; Status = $true }
                        $databaseRows = @($databaseResult.Rows | ForEach-Object {
                            [pscustomobject][ordered]@{
                                Identity = [string]$_.Identity
                                Name = [string]$_.Name
                                Mounted = [bool]$_.Mounted
                                Server = [string]$_.Server
                                ReplicationType = [string]$_.ReplicationType
                            }
                        })
                        $databaseCache[$databaseKey] = [pscustomobject]@{ Success = [bool]$databaseResult.Success; Rows = $databaseRows }
                    }
                    $databaseHealth = @($databaseCache[$databaseKey].Rows)
                    $databaseHealthAvailable = [bool]$databaseCache[$databaseKey].Success
                }

                if ((Test-WorkerCheckEnabled -CheckId 'PERMISSIONS-BASELINE') -or (Test-WorkerCheckEnabled -CheckId 'DELEGATE-MIGRATION-DEPENDENCY')) {
                    Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Delegated permissions - {0}" -f $email)
                    $mailboxPermissionResult = Invoke-WorkerCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $identity }
                    $mailboxPermissionSuccess = [bool]$mailboxPermissionResult.Success
                    foreach ($permission in @($mailboxPermissionResult.Rows)) {
                        $rights = @(ConvertTo-TextArray $permission.AccessRights)
                        $delegate = [string]$permission.User
                        if ($permission.IsInherited -or $rights -notcontains 'FullAccess' -or $delegate -match 'NT AUTHORITY|S-1-5-|SELF') { continue }
                        [void]$permissions.Add([pscustomobject]@{ PermissionType = 'FullAccess'; Delegate = $delegate; IsInherited = $false; Source = 'Exchange2016Worker' })
                    }

                    $sendAsPermissionResult = Invoke-WorkerCommand -Name 'Get-ADPermission' -Parameters @{ Identity = $mailboxes[0].DistinguishedName }
                    $sendAsPermissionSuccess = [bool]$sendAsPermissionResult.Success
                    foreach ($permission in @($sendAsPermissionResult.Rows)) {
                        if ($permission.IsInherited -or @(ConvertTo-TextArray $permission.ExtendedRights) -notcontains 'Send-As' -or $permission.Deny) { continue }
                        [void]$permissions.Add([pscustomobject]@{ PermissionType = 'SendAs'; Delegate = [string]$permission.User; IsInherited = $false; Source = 'Exchange2016Worker' })
                    }
                    foreach ($delegate in @($mailboxes[0].GrantSendOnBehalfTo)) {
                        [void]$permissions.Add([pscustomobject]@{ PermissionType = 'SendOnBehalf'; Delegate = [string]$delegate; IsInherited = $false; Source = 'Exchange2016Worker' })
                    }
                }
            }

            $addresses = @()
            if ($mailboxes.Count -eq 1) {
                $addresses += @($mailboxes[0].EmailAddresses)
                if ($mailboxes[0].ExternalEmailAddress) { $addresses += [string]$mailboxes[0].ExternalEmailAddress }
            }
            $lookupResults = @($recipientResult)
            if ($needMailbox) { $lookupResults += $mailboxResult }
            if ($needRemoteMailbox) { $lookupResults += $remoteMailboxResult }
            if ($needMailUser) { $lookupResults += $mailUserResult }
            $coreAvailable = @($lookupResults | Where-Object { -not $_.Success }).Count -eq 0
            $errors = @($lookupResults | Where-Object { -not $_.Success -and $_.ErrorMessage } | ForEach-Object { $_.ErrorMessage } | Sort-Object -Unique)
            $addressConflicts = @()
            if ((Test-WorkerCheckEnabled -CheckId 'PROXY-SMTP-GLOBAL-UNIQUE') -or (Test-WorkerCheckEnabled -CheckId 'TARGET-ADDRESS-GLOBAL-UNIQUE')) {
                Write-WorkerProgress -Current $index -Total $emails.Count -Message ("SMTP uniqueness - {0}" -f $email)
                $addressConflicts = @(Get-AddressConflictEvidence -MailboxAddress $email -Addresses $addresses)
            }
            [void]$evidence.Add([pscustomobject][ordered]@{
                EmailAddress = $email
                Available = $coreAvailable
                Source = "Local Exchange 2016 Management Shell on $env:COMPUTERNAME"
                SourceTimestamp = Get-Date
                Message = if ($coreAvailable) { 'Exchange 2016 evidence collected through direct local cmdlets.' } else { 'Exchange 2016 lookup failed: ' + ($errors -join ' | ') }
                Mailboxes = $mailboxes
                RemoteMailboxes = $remoteMailboxes
                MailUsers = $mailUsers
                Recipients = $recipients
                RecipientLookupAvailable = [bool]$recipientResult.Success
                Statistics = $statistics
                StatisticsAvailable = $statisticsAvailable
                ArchiveStatistics = $archiveStatistics
                ArchiveStatisticsAvailable = $archiveStatisticsAvailable
                Permissions = $permissions.ToArray()
                PermissionsAvailable = [bool]($mailboxPermissionSuccess -and $sendAsPermissionSuccess)
                HoldDataAvailable = [bool]($mailboxes.Count -eq 1 -and $mailboxResult.Success)
                FolderStatistics = $folderStatistics
                FolderStatisticsAvailable = $folderStatisticsAvailable
                InboxRules = $inboxRules
                InboxRulesAvailable = $inboxRulesAvailable
                DatabaseHealth = $databaseHealth
                DatabaseHealthAvailable = $databaseHealthAvailable
                DatabaseHealthSource = "Local Exchange 2016 Management Shell mailbox database on $env:COMPUTERNAME"
                DatabaseHealthSourceTimestamp = Get-Date
                DeliveryRestrictionsAvailable = [bool]($mailboxes.Count -eq 1 -and $mailboxResult.Success)
                AddressConflicts = $addressConflicts
            })
        }
        catch {
            $message = $_.Exception.Message
            [void]$collectionErrors.Add([pscustomobject][ordered]@{ EmailAddress = $email; Message = $message })
            [void]$evidence.Add((Get-WorkerFailureEvidence -EmailAddress $email -Message $message))
        }
    }
    $hybridEvidence = Get-HybridEvidence
    $result = [pscustomobject][ordered]@{
        WorkerVersion = $workerVersion
        ComputerName = $env:COMPUTERNAME
        PowerShellVersion = [string]$PSVersionTable.PSVersion
        SnapInName = $snapInName
        CollectedAt = Get-Date
        DurationSeconds = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 1)
        MailboxEvidenceCount = $evidence.Count
        MailboxObjectCount = [int](@($evidence.ToArray() | ForEach-Object { @($_.Mailboxes).Count } | Measure-Object -Sum).Sum)
        PermissionCount = @($evidence.ToArray() | ForEach-Object { @($_.Permissions) }).Count
        ErrorCount = $collectionErrors.Count
        Errors = $collectionErrors.ToArray()
        Evidence = $evidence.ToArray()
        Hybrid = $hybridEvidence
    }
    $result | Export-Clixml -LiteralPath $OutputPath -Depth 12 -Force
    Write-WorkerProgress -Current $emails.Count -Total $emails.Count -Message 'Complete'
}
catch {
    Write-WorkerText -Path $ErrorPath -Text ($_ | Out-String)
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqEiQfMJJTJFBa
# SPb0Z1FNLRBJ+/nCTqNqUuVBmSzcPKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHh5aHih+tLE7vizJRuRIx6XtODMSc4MBesyV123c3+xMA0GCSqG
# SIb3DQEBAQUABIIBgFWOhuCSta23D5C65Kiszb7Ea24Uy4LpIN/RITv8rnzKTZof
# DRTy+dOZiyGI8sxKt83v9ms8ARpAXGTrs/AkSF8Ftll0Y42th0S/XS+NFpuvdfT8
# fQ+IrQxzn6bSFQFZjuDook4rt3ZZ/ilyeIu8SQjtu5BG645Mw/mjsdOAvWjIrneD
# JGe2saSMEENtiMFs0xz23yYhh0bpwST3tOWSnb+2SnkxVokGiVNv2C/opY0u6yFq
# 1MgXtbZn2C7sXOJDYnkNV36h0UrlmGr3vHWM4h+onDXj9iSUx0hoGBQuOJDVWZ5T
# reyHGvF+uCT/u7KC4uStuqy8Rc0hMWnXQRlmO8sL1jitKBEPV/Mm8TFHqhJjIK7+
# Jb97/fQl96Y9ahmGYl1pg0Z3dOySg4Q3Dso4WGpALyETZoU5PQKd8YgPKiPPBqfV
# XsB1B0O+eo0Ce/hzVI7QmBantb7d65cq/yTtuMvgzGlJ3gg/F/arx8EY/qU0KIHs
# /eaRgE0FZq9WxNvLg6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxMTU2
# MzNaMC8GCSqGSIb3DQEJBDEiBCAwXcoT5/ZoJ6x/BZnv6IAJUd11PQ5y7TVZ6KSS
# coS51DANBgkqhkiG9w0BAQEFAASCAgAC2YEV2RJ+YoCKIb5lzzgmgZhQUs8Gcrf4
# RalsOwbtPVGVeHIcqGqveNR+vwyUvIC1F+00tAin8dy4QW32UuIsqDi4NBmo+cGr
# lDICQMqrwV9GWbmpRQ8EntgU3wIKJu3KyZwK5HCvYBN4cvJg96yJMKPpyxPdmj9D
# EyAhMNK/rnJc1X3YUosJWA+rSkHHtW219W8pMqozwc04KSSrIs3wG9wxzHYsCTfH
# IGcy30Ywy+0pqoXexBCzrgwtvfbuIdBQpTzlmwne4AWKTiG5nTi/nDkDOytlu4hb
# g6SptMZFY9/9ZXOR1l608Vai/GT8X3mvNR1xh6IrRe1U7zlNc1fbJqVaenGYxEZ1
# gA+wc3H6ZeS5CQQm/Zyq88oR3UG/+qtvUK2W4Ng6bFAxJKuUd74Q7/qOi6N+BYDB
# 0JLBL0hIKGJO1ecTOH7kEQxTRyX/0QwfQFan22EdkC4kvO40GaELKbmAVDpM6juK
# Uz27DMf0SmewLSO41hPQZ1xd0Gegr7XQMUIGHddrM2x5tBSeIOh47MzCmfBXA9Z5
# qipvQR7nVxmmk6/Xtt4BByt0aQtjIjP4eC1HrESueY6bydEYVmRPv18ykeAHSc3Z
# omAtZ7zHikFS8cTGBqATqylhSYtUfJwxL0MYrjdEFSAZU0/m0FK224G7BWzOhuE3
# 16J41XmqBA==
# SIG # End signature block
