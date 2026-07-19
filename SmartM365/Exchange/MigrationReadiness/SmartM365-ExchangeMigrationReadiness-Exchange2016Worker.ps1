#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InputPath = '',
    [string]$OutputPath = '',
    [string]$ErrorPath = '',
    [string]$ProgressPath = '',
    [switch]$ValidateOnly
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'
$workerVersion = '1.11.1'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
[Console]::InputEncoding = $utf8
$OutputEncoding = $utf8

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
        $query = Invoke-WorkerCommand -Name 'Get-Recipient' -Parameters @{
            Filter = "EmailAddresses -eq 'smtp:$safeAddress' -or ExternalEmailAddress -eq 'smtp:$safeAddress'"
            ResultSize = 'Unlimited'
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
    return @($items)
}

function Get-HybridEvidence {
    $collectedAt = Get-Date
    $ews = Invoke-WorkerCommand -Name 'Get-WebServicesVirtualDirectory'
    $org = Invoke-WorkerCommand -Name 'Get-OrganizationConfig'
    $ioc = Invoke-WorkerCommand -Name 'Get-IntraOrganizationConnector'
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

    $emails = @(Import-Clixml -LiteralPath $InputPath | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
    $evidence = New-Object System.Collections.Generic.List[object]
    $index = 0
    foreach ($email in $emails) {
        $index++
        Write-WorkerProgress -Current $index -Total $emails.Count -Message $email

        $safeEmail = $email.Replace("'", "''")
        $identityFilter = "EmailAddresses -eq 'smtp:$safeEmail'"
        $mailboxResult = Invoke-WorkerCommand -Name 'Get-Mailbox' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' }
        $remoteMailboxResult = Invoke-WorkerCommand -Name 'Get-RemoteMailbox' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' }
        $mailUserResult = Invoke-WorkerCommand -Name 'Get-MailUser' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' }
        $recipientResult = Invoke-WorkerCommand -Name 'Get-Recipient' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' }
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
            $statisticsResult = Invoke-WorkerCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $identity }
            $statistics = @($statisticsResult.Rows | ForEach-Object { ConvertTo-StatisticsEvidence $_ })
            $statisticsAvailable = [bool]$statisticsResult.Success

            $hasArchive = $mailboxes[0].ArchiveStatus -match '^(?i:Active|HostedPending|Local)$' -or ($mailboxes[0].ArchiveGuid -and $mailboxes[0].ArchiveGuid -notmatch '^0{8}-')
            if ($hasArchive) {
                $archiveResult = Invoke-WorkerCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $identity; Archive = $true }
                $archiveStatistics = @($archiveResult.Rows | ForEach-Object { ConvertTo-StatisticsEvidence $_ })
                $archiveStatisticsAvailable = [bool]$archiveResult.Success
            }

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

            if ($mailboxes[0].Database) {
                $databaseResult = Invoke-WorkerCommand -Name 'Get-MailboxDatabase' -Parameters @{ Identity = $mailboxes[0].Database; Status = $true }
                $databaseHealth = @($databaseResult.Rows | ForEach-Object {
                    [pscustomobject][ordered]@{
                        Identity = [string]$_.Identity
                        Name = [string]$_.Name
                        Mounted = [bool]$_.Mounted
                        Server = [string]$_.Server
                        ReplicationType = [string]$_.ReplicationType
                    }
                })
                $databaseHealthAvailable = [bool]$databaseResult.Success
            }

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

        $addresses = @()
        if ($mailboxes.Count -eq 1) {
            $addresses += @($mailboxes[0].EmailAddresses)
            if ($mailboxes[0].ExternalEmailAddress) { $addresses += [string]$mailboxes[0].ExternalEmailAddress }
        }
        $coreAvailable = [bool]($mailboxResult.Success -and $remoteMailboxResult.Success -and $mailUserResult.Success -and $recipientResult.Success)
        $lookupResults = @($mailboxResult, $remoteMailboxResult, $mailUserResult, $recipientResult)
        $errors = @($lookupResults | Where-Object { -not $_.Success -and $_.ErrorMessage } | ForEach-Object { $_.ErrorMessage } | Sort-Object -Unique)
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
            Permissions = @($permissions)
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
            AddressConflicts = @(Get-AddressConflictEvidence -MailboxAddress $email -Addresses $addresses)
        })
    }

    $result = [pscustomobject][ordered]@{
        WorkerVersion = $workerVersion
        ComputerName = $env:COMPUTERNAME
        PowerShellVersion = [string]$PSVersionTable.PSVersion
        SnapInName = $snapInName
        CollectedAt = Get-Date
        Evidence = @($evidence)
        Hybrid = Get-HybridEvidence
    }
    $result | Export-Clixml -LiteralPath $OutputPath -Depth 12 -Force
    Write-WorkerProgress -Current $emails.Count -Total $emails.Count -Message 'Complete'
}
catch {
    Write-WorkerText -Path $ErrorPath -Text ($_ | Out-String)
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
