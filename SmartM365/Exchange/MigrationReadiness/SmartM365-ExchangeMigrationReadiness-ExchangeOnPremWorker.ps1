#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InputPath = '',
    [string]$OutputPath = '',
    [string]$ErrorPath = '',
    [string]$ProgressPath = '',
    [string]$CancelPath = '',
    [string]$DiagnosticsDirectory = '',
    [switch]$RecipientBatchMode,
    [string]$RecipientBatchInputPath = '',
    [string]$RecipientBatchOutputPath = '',
    [string]$RecipientBatchErrorPath = '',
    [string]$RecipientBatchLogPath = '',
    [string]$RecipientBatchSnapInName = 'Microsoft.Exchange.Management.PowerShell.SnapIn',
    [switch]$ValidateOnly,
    [switch]$SelfTest
)

Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'
$workerVersion = '1.11.14'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
[Console]::InputEncoding = $utf8
$OutputEncoding = $utf8
$enabledChecks = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$workerStartedAt = Get-Date
$workerRunId = if ([string]::IsNullOrWhiteSpace($DiagnosticsDirectory)) { '-' } else { Split-Path $DiagnosticsDirectory -Leaf }
$script:WorkerLogPath = if ([string]::IsNullOrWhiteSpace($DiagnosticsDirectory)) { '' } else { Join-Path $DiagnosticsDirectory 'ExchangeOnPrem-Worker.log' }

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

function Write-WorkerChildLog {
    param([string]$Path, [string]$Level, [string]$Message)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    try {
        $directory = Split-Path -Parent $Path
        if ($directory -and -not (Test-Path -LiteralPath $directory -PathType Container)) {
            [void](New-Item -Path $directory -ItemType Directory -Force)
        }
        $line = "{0} [{1}] [PID={2}] [Run={3}] {4}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level.ToUpperInvariant(), $PID, $workerRunId, $Message
        [IO.File]::AppendAllText($Path, "$line`r`n", $utf8)
    }
    catch { $null = $_ }
}

function Write-WorkerLog {
    param([ValidateSet('DEBUG','INFO','WARN','ERROR','SUCCESS')][string]$Level,[string]$Message)
    Write-WorkerChildLog -Path $script:WorkerLogPath -Level $Level -Message ("[Component=ExchangeOnPremWorker] [ElapsedMs={0}] {1}" -f ([math]::Round(((Get-Date)-$workerStartedAt).TotalMilliseconds)),($Message -replace '[\r\n]+',' | '))
}

function Format-WorkerError {
    param([Parameter(Mandatory)]$ErrorRecord)
    $parts = New-Object System.Collections.Generic.List[string]
    [void]$parts.Add("ExceptionType=$($ErrorRecord.Exception.GetType().FullName)")
    [void]$parts.Add("Message=$($ErrorRecord.Exception.Message)")
    if($ErrorRecord.FullyQualifiedErrorId){[void]$parts.Add("FullyQualifiedErrorId=$($ErrorRecord.FullyQualifiedErrorId)")}
    if($ErrorRecord.CategoryInfo){[void]$parts.Add("CategoryInfo=$($ErrorRecord.CategoryInfo)")}
    if($ErrorRecord.InvocationInfo){[void]$parts.Add("Command=$($ErrorRecord.InvocationInfo.MyCommand.Name); Script=$($ErrorRecord.InvocationInfo.ScriptName); Line=$($ErrorRecord.InvocationInfo.ScriptLineNumber); SourceLine=$(if($ErrorRecord.InvocationInfo.Line){$ErrorRecord.InvocationInfo.Line.Trim()}else{''})")}
    if($ErrorRecord.ScriptStackTrace){[void]$parts.Add("ScriptStackTrace=$($ErrorRecord.ScriptStackTrace)")}
    $inner=$ErrorRecord.Exception.InnerException;$depth=0
    while($inner -and $depth -lt 5){$depth++;[void]$parts.Add("InnerException${depth}=$($inner.GetType().FullName): $($inner.Message)");$inner=$inner.InnerException}
    return ($parts -join ' | ')
}

function Get-WorkerParameterSummary {
    param([hashtable]$Parameters)
    return (@($Parameters.Keys | Sort-Object | ForEach-Object {
        $name=[string]$_;$value=$Parameters[$name]
        if($name -eq 'Filter'){"FilterLength=$(([string]$value).Length)"}
        elseif($value -is [array]){"$name.Count=$(@($value).Count)"}
        else{"$name=$value"}
    }) -join ';')
}

function ConvertTo-WorkerCommandLineArgument {
    param([Parameter(Mandatory)][string]$Value)
    if ($Value -notmatch '[\s"]') { return $Value }
    return '"' + $Value.Replace('"', '\"') + '"'
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
    $timer=[Diagnostics.Stopwatch]::StartNew()
    $parameterSummary=Get-WorkerParameterSummary -Parameters $Parameters
    Write-WorkerLog -Level DEBUG -Message "Command starting. Name=$Name; Parameters=$parameterSummary."
    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        $timer.Stop()
        $message="$Name is unavailable."
        Write-WorkerLog -Level ERROR -Message "Command unavailable. Name=$Name; DurationMs=$($timer.ElapsedMilliseconds); Message=$message."
        return [pscustomobject]@{ Success = $false; Rows = @(); ErrorMessage = $message }
    }
    try {
        $rows=@(& $Name @Parameters -ErrorAction Stop)
        $timer.Stop()
        Write-WorkerLog -Level SUCCESS -Message "Command completed. Name=$Name; DurationMs=$($timer.ElapsedMilliseconds); RowCount=$($rows.Count)."
        return [pscustomobject]@{ Success = $true; Rows = $rows; ErrorMessage = '' }
    }
    catch {
        $timer.Stop()
        if ($NotFoundIsEmpty -and (Test-WorkerNotFoundError -Message $_.Exception.Message)) {
            Write-WorkerLog -Level INFO -Message "Command returned expected not-found state. Name=$Name; DurationMs=$($timer.ElapsedMilliseconds); Message=$($_.Exception.Message)."
            return [pscustomobject]@{ Success = $true; Rows = @(); ErrorMessage = '' }
        }
        Write-WorkerLog -Level ERROR -Message "Command failed. Name=$Name; DurationMs=$($timer.ElapsedMilliseconds); $(Format-WorkerError $_)"
        return [pscustomobject]@{ Success = $false; Rows = @(); ErrorMessage = $_.Exception.Message }
    }
}

function Add-WorkerCommandFailure {
    param(
        [Parameter(Mandatory)]$CollectionErrors,
        [Parameter(Mandatory)]$MailboxErrors,
        [Parameter(Mandatory)][string]$EmailAddress,
        [Parameter(Mandatory)][string]$CheckId,
        [Parameter(Mandatory)][string]$CommandName,
        [Parameter(Mandatory)]$Result
    )
    if ($Result.Success) { return }
    $entry = [pscustomobject][ordered]@{
        EmailAddress = $EmailAddress
        CheckId = $CheckId
        Command = $CommandName
        Message = [string]$Result.ErrorMessage
        IsFatal = $false
    }
    [void]$CollectionErrors.Add($entry)
    [void]$MailboxErrors.Add($entry)
    Write-WorkerLog -Level WARN -Message "Partial mailbox evidence unavailable. EmailAddress=$EmailAddress; CheckId=$CheckId; Command=$CommandName; Message=$($entry.Message)."
}

function ConvertTo-TextArray {
    param($Value)
    return @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Test-WorkerMigrationRelevantDelegate {
    param([AllowNull()][string]$Identity)

    $value = ([string]$Identity).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { return $false }
    if ($value -match '^(?i:SELF|NT AUTHORITY\\SELF|S-1-5-10)$') { return $false }

    $shortName = if ($value -match '\\([^\\]+)$') { $Matches[1] } else { $value }
    if ($shortName -match '^(?i:Exchange Domain Servers|Exchange Servers|Exchange Services|Exchange Trusted Subsystem|Exchange Windows Permissions)$') {
        return $false
    }
    return $true
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

function Test-WorkerCancellation {
    return -not [string]::IsNullOrWhiteSpace($CancelPath) -and (Test-Path -LiteralPath $CancelPath -PathType Leaf)
}

function ConvertTo-NormalizedSmtpAddress {
    param($Value)
    $text = ([string]$Value).Trim()
    if ($text -match '^(?i:smtp:)(.+)$') { $text = $Matches[1].Trim() }
    if ($text -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return '' }
    return $text.ToLowerInvariant()
}

function Invoke-WorkerRecipientAddressBatch {
    param(
        [Parameter(Mandatory)][string[]]$Addresses,
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [Parameter(Mandatory)][string]$SnapInName,
        [int]$BatchNumber = 0,
        [string]$DiagnosticsDirectory = '',
        [ValidateRange(0,300)][int]$TestDelaySeconds = 0
    )

    $clauses = @(
        foreach ($address in $Addresses) {
            $safeAddress = $address.Replace("'", "''")
            "(EmailAddresses -eq 'smtp:$safeAddress' -or ExternalEmailAddress -eq 'smtp:$safeAddress')"
        }
    )
    $filter = $clauses -join ' -or '
    $batchRoot = Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-ExchangeMigrationReadiness\SmtpBatch-{0}" -f [guid]::NewGuid().ToString('N'))
    $batchInputPath = Join-Path $batchRoot 'input.clixml'
    $batchOutputPath = Join-Path $batchRoot 'output.clixml'
    $batchErrorPath = Join-Path $batchRoot 'error.txt'
    $logFileName = if ($BatchNumber -gt 0) { "ExchangeOnPrem-SmtpBatch-{0:D3}.log" -f $BatchNumber } else { "ExchangeOnPrem-SmtpBatch-{0}.log" -f [guid]::NewGuid().ToString('N') }
    $batchLogPath = if ([string]::IsNullOrWhiteSpace($DiagnosticsDirectory)) { Join-Path $batchRoot $logFileName } else { Join-Path $DiagnosticsDirectory $logFileName }
    $process = $null
    $started = Get-Date
    try {
        [void](New-Item -Path $batchRoot -ItemType Directory -Force)
        if (-not [string]::IsNullOrWhiteSpace($DiagnosticsDirectory)) {
            [void](New-Item -Path $DiagnosticsDirectory -ItemType Directory -Force)
        }
        [pscustomobject][ordered]@{ Addresses = @($Addresses); Filter = $filter; TestDelaySeconds = $TestDelaySeconds } | Export-Clixml -LiteralPath $batchInputPath -Depth 4 -Force
        Write-WorkerChildLog -Path $batchLogPath -Level INFO -Message ("Parent starting SMTP uniqueness child process; batch={0}; addresses={1}; timeout={2}s." -f $BatchNumber, $Addresses.Count, $TimeoutSeconds)
        Write-WorkerChildLog -Path $batchLogPath -Level INFO -Message ("Candidate addresses: {0}" -f ($Addresses -join ';'))

        $powershellPath = Join-Path $PSHOME 'powershell.exe'
        $arguments = @(
            '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath,
            '-RecipientBatchMode', '-RecipientBatchInputPath', $batchInputPath, '-RecipientBatchOutputPath', $batchOutputPath,
            '-RecipientBatchErrorPath', $batchErrorPath, '-RecipientBatchLogPath', $batchLogPath,
            '-RecipientBatchSnapInName', $SnapInName
        )
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $powershellPath
        $startInfo.Arguments = (@($arguments | ForEach-Object { ConvertTo-WorkerCommandLineArgument ([string]$_) }) -join ' ')
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $process = [Diagnostics.Process]::Start($startInfo)
        if (-not $process) { throw 'SMTP uniqueness child process could not be started.' }
        Write-WorkerChildLog -Path $batchLogPath -Level INFO -Message ("Parent observed child PID={0}." -f $process.Id)

        $deadline = (Get-Date).AddSeconds([math]::Max(5, $TimeoutSeconds))
        while (-not $process.WaitForExit(250)) {
            if (Test-WorkerCancellation) {
                Write-WorkerChildLog -Path $batchLogPath -Level WARN -Message 'Cancellation requested; terminating child process.'
                try { $process.Kill() } catch { Write-WorkerChildLog -Path $batchLogPath -Level ERROR -Message ("Child termination failed: {0}" -f $_.Exception.Message) }
                [void]$process.WaitForExit(5000)
                throw [OperationCanceledException]::new('Exchange on-premises worker cancellation requested.')
            }
            if ((Get-Date) -ge $deadline) {
                Write-WorkerChildLog -Path $batchLogPath -Level WARN -Message ("Timeout reached after {0} second(s); terminating child process." -f $TimeoutSeconds)
                try { $process.Kill() } catch { Write-WorkerChildLog -Path $batchLogPath -Level ERROR -Message ("Child termination failed: {0}" -f $_.Exception.Message) }
                [void]$process.WaitForExit(5000)
                return [pscustomobject]@{
                    Success = $false
                    TimedOut = $true
                    Rows = @()
                    DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
                    ErrorMessage = "Get-Recipient address batch exceeded $TimeoutSeconds second(s); child process terminated."
                    LogPath = $batchLogPath
                }
            }
        }

        if ($process.ExitCode -ne 0) {
            $errorMessage = if (Test-Path -LiteralPath $batchErrorPath -PathType Leaf) { [IO.File]::ReadAllText($batchErrorPath) } else { "Child process exit code $($process.ExitCode)." }
            Write-WorkerChildLog -Path $batchLogPath -Level ERROR -Message ("Child process failed with exit code {0}: {1}" -f $process.ExitCode, $errorMessage.Trim())
            return [pscustomobject]@{ Success = $false; TimedOut = $false; Rows = @(); DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 1); ErrorMessage = $errorMessage.Trim(); LogPath = $batchLogPath }
        }
        if (-not (Test-Path -LiteralPath $batchOutputPath -PathType Leaf)) {
            throw 'SMTP uniqueness child process completed without an output file.'
        }
        $rows = @(Import-Clixml -LiteralPath $batchOutputPath)
        Write-WorkerChildLog -Path $batchLogPath -Level SUCCESS -Message ("Parent imported {0} recipient row(s); duration={1}s." -f $rows.Count, ([math]::Round(((Get-Date) - $started).TotalSeconds, 1)))
        return [pscustomobject]@{ Success = $true; TimedOut = $false; Rows = $rows; DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 1); ErrorMessage = ''; LogPath = $batchLogPath }
    }
    catch [OperationCanceledException] { throw }
    catch {
        Write-WorkerChildLog -Path $batchLogPath -Level ERROR -Message (Format-WorkerError $_)
        return [pscustomobject]@{ Success = $false; TimedOut = $false; Rows = @(); DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 1); ErrorMessage = $_.Exception.Message; LogPath = $batchLogPath }
    }
    finally {
        if ($process) {
            if (-not $process.HasExited) { try { $process.Kill() } catch { $null = $_ } }
            $process.Dispose()
        }
        if (Test-Path -LiteralPath $batchRoot) { Remove-Item -LiteralPath $batchRoot -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
function Initialize-WorkerAddressConflictEvidence {
    param(
        [Parameter(Mandatory)]$Evidence,
        [Parameter(Mandatory)][hashtable]$CandidateAddressesByEmail,
        [Parameter(Mandatory)][int]$BatchSize,
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [Parameter(Mandatory)][string]$SnapInName,
        [string]$DiagnosticsDirectory = ''
    )

    $started = Get-Date
    $normalizedBatchSize = [math]::Max(1, [math]::Min(50, $BatchSize))
    $allAddresses = @($CandidateAddressesByEmail.Values | ForEach-Object { @($_) } | ForEach-Object { $_ } | Where-Object { $_ } | Sort-Object -Unique)
    $ownerIndex = @{}
    $queryState = @{}
    foreach ($address in $allAddresses) {
        $ownerIndex[$address] = New-Object System.Collections.Generic.List[object]
        $queryState[$address] = [pscustomobject]@{ Available = $false; ErrorMessage = 'Address was not queried.' }
    }

    $batchCount = if ($allAddresses.Count -eq 0) { 0 } else { [int][math]::Ceiling($allAddresses.Count / [double]$normalizedBatchSize) }
    $timeoutCount = 0
    $errorCount = 0
    $childLogPaths = New-Object System.Collections.Generic.List[string]
    for ($batchIndex = 0; $batchIndex -lt $batchCount; $batchIndex++) {
        if (Test-WorkerCancellation) { throw [OperationCanceledException]::new('Exchange on-premises worker cancellation requested.') }
        $offset = $batchIndex * $normalizedBatchSize
        $last = [math]::Min($allAddresses.Count - 1, $offset + $normalizedBatchSize - 1)
        $batchAddresses = @($allAddresses[$offset..$last])
        Write-WorkerProgress -Current ($batchIndex + 1) -Total $batchCount -Message ("SMTP uniqueness batch {0}/{1} - {2} address(es)" -f ($batchIndex + 1), $batchCount, $batchAddresses.Count)
        $query = Invoke-WorkerRecipientAddressBatch -Addresses $batchAddresses -TimeoutSeconds $TimeoutSeconds -SnapInName $SnapInName -BatchNumber ($batchIndex + 1) -DiagnosticsDirectory $DiagnosticsDirectory
        if (-not [string]::IsNullOrWhiteSpace([string]$query.LogPath)) { [void]$childLogPaths.Add([string]$query.LogPath) }
        if ($query.TimedOut) { $timeoutCount++ }
        if (-not $query.Success) { $errorCount++ }
        foreach ($address in $batchAddresses) {
            $queryState[$address] = [pscustomobject]@{ Available = [bool]$query.Success; ErrorMessage = [string]$query.ErrorMessage }
        }
        if (-not $query.Success) { continue }

        foreach ($owner in @($query.Rows)) {
            $ownerAddresses = @(
                @($owner.EmailAddresses) + @($owner.ExternalEmailAddress) + @($owner.PrimarySmtpAddress) |
                    ForEach-Object { ConvertTo-NormalizedSmtpAddress $_ } |
                    Where-Object { $_ } |
                    Sort-Object -Unique
            )
            foreach ($ownerAddress in $ownerAddresses) {
                if ($ownerIndex.ContainsKey($ownerAddress)) { [void]$ownerIndex[$ownerAddress].Add($owner) }
            }
        }
    }

    foreach ($entry in @($Evidence.ToArray())) {
        $email = ([string]$entry.EmailAddress).Trim().ToLowerInvariant()
        $expectedIdentities = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($object in @($entry.Mailboxes) + @($entry.RemoteMailboxes) + @($entry.MailUsers) + @($entry.Recipients)) {
            foreach ($identityValue in @($object.Identity, $object.DistinguishedName)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$identityValue)) { [void]$expectedIdentities.Add(([string]$identityValue).Trim()) }
            }
        }
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($address in @($CandidateAddressesByEmail[$email])) {
            $owners = New-Object System.Collections.Generic.List[string]
            foreach ($owner in @($ownerIndex[$address].ToArray())) {
                $ownerIdentity = [string]$owner.Identity
                $ownerDn = [string]$owner.DistinguishedName
                $ownerPrimary = ConvertTo-NormalizedSmtpAddress $owner.PrimarySmtpAddress
                $isExpectedOwner = ($ownerIdentity -and $expectedIdentities.Contains($ownerIdentity)) -or ($ownerDn -and $expectedIdentities.Contains($ownerDn)) -or ($ownerPrimary -eq $email)
                if (-not $isExpectedOwner) { [void]$owners.Add(("{0} -> {1}" -f $address, $ownerIdentity)) }
            }
            $state = $queryState[$address]
            [void]$items.Add([pscustomobject][ordered]@{
                Address = $address
                Available = [bool]$state.Available
                Conflicts = @($owners | Sort-Object -Unique)
                ErrorMessage = [string]$state.ErrorMessage
            })
        }
        $entry.AddressConflicts = $items.ToArray()
    }

    return [pscustomobject][ordered]@{
        CandidateAddressCount = $allAddresses.Count
        BatchCount = $batchCount
        TimeoutCount = $timeoutCount
        ErrorCount = $errorCount
        BatchSize = $normalizedBatchSize
        TimeoutSeconds = $TimeoutSeconds
        DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
        DiagnosticsDirectory = $DiagnosticsDirectory
        ChildLogPaths = $childLogPaths.ToArray()
    }
}
function Get-WorkerFailureEvidence {
    param([string]$EmailAddress, [string]$Message)
    $now = Get-Date
    return [pscustomobject][ordered]@{
        EmailAddress = $EmailAddress
        Available = $false
        Source = "Local Exchange on-premises Management Shell on $env:COMPUTERNAME"
        SourceTimestamp = $now
        Message = "Exchange on-premises mailbox collection failed: $Message"
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
        InboxRulesComplete = $false
        DatabaseHealth = @()
        DatabaseHealthAvailable = $false
        DatabaseHealthSource = "Local Exchange on-premises Management Shell mailbox database on $env:COMPUTERNAME"
        DatabaseHealthSourceTimestamp = $now
        DeliveryRestrictionsAvailable = $false
        AddressConflicts = @()
    }
}
function Get-HybridEvidence {
    $collectedAt = Get-Date
    $empty = [pscustomobject]@{ Success = $false; Rows = @(); ErrorMessage = 'Check disabled for this run.' }
    if (Test-WorkerCheckEnabled -CheckId 'HYBRID-MRSPROXY') {
        Write-WorkerProgress -Current 0 -Total 0 -Message 'Optional local diagnostic - reading all EWS virtual directories and configured MRS Proxy state'
        $ews = Invoke-WorkerCommand -Name 'Get-WebServicesVirtualDirectory'
    }
    else { $ews = $empty }
    if (Test-WorkerCheckEnabled -CheckId 'HYBRID-AUTODISCOVER-OAUTH') {
        Write-WorkerProgress -Current 0 -Total 0 -Message 'Hybrid readiness - reading organization OAuth configuration'
        $org = Invoke-WorkerCommand -Name 'Get-OrganizationConfig'
        Write-WorkerProgress -Current 0 -Total 0 -Message 'Hybrid readiness - reading intra-organization connectors'
        $ioc = Invoke-WorkerCommand -Name 'Get-IntraOrganizationConnector'
    }
    else { $org = $empty; $ioc = $empty }
    $enabledEws = @($ews.Rows | Where-Object { [bool]$_.MRSProxyEnabled })
    $organization = @($org.Rows | Select-Object -First 1)
    $enabledIoc = @($ioc.Rows | Where-Object { [bool]$_.Enabled })
    $oauthEnabled = $organization.Count -eq 1 -and [bool]$organization[0].OAuth2ClientProfileEnabled
    return [pscustomobject][ordered]@{
        MrsProxyAvailable = [bool]$ews.Success
        MrsProxyEnabled = $enabledEws.Count -gt 0
        MrsProxyMessage = if (-not $ews.Success) { $ews.ErrorMessage } elseif ($enabledEws.Count -gt 0) { "MRSProxy enabled on $($enabledEws.Count) EWS virtual directorie(s)." } else { 'No EWS virtual directory has MRSProxy enabled.' }
        MrsProxySource = 'Local Exchange on-premises Management Shell EWS virtual directories'
        MrsProxySourceTimestamp = $collectedAt
        OAuthAvailable = [bool]($org.Success -and $ioc.Success)
        OAuthHealthy = [bool]($oauthEnabled -and $enabledIoc.Count -gt 0)
        OAuthMessage = if (-not $org.Success -or -not $ioc.Success) { (@($org.ErrorMessage, $ioc.ErrorMessage) | Where-Object { $_ } | Select-Object -Unique) -join ' | ' } else { "OAuth2ClientProfileEnabled=$oauthEnabled; EnabledIntraOrganizationConnector=$($enabledIoc.Count -gt 0)." }
        OAuthSource = 'Local Exchange on-premises Management Shell hybrid configuration'
        OAuthSourceTimestamp = $collectedAt
    }
}


function Get-WorkerExchangeProductName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Major,
        [Parameter(Mandatory)][int]$Minor,
        [Parameter(Mandatory)][int]$Build
    )

    if ($Major -eq 15 -and $Minor -eq 1) { return 'Exchange Server 2016' }
    if (($Major -eq 15 -and $Minor -eq 2 -and $Build -ge 2562) -or ($Major -eq 15 -and $Minor -gt 2)) {
        return 'Exchange Server Subscription Edition'
    }
    if ($Major -eq 15 -and $Minor -eq 2) { return 'Exchange Server 2019' }
    throw "Unsupported Exchange server version '$Major.$Minor.$Build'. Supported products are Exchange Server 2016, 2019 and Subscription Edition."
}

function Get-WorkerExchangeServerInfo {
    [CmdletBinding()]
    param()

    $servers = @(Get-ExchangeServer -ErrorAction Stop)
    if ($servers.Count -eq 0) {
        throw 'Get-ExchangeServer returned no Exchange server.'
    }
    $localServers = @($servers | Where-Object {
        [string]$_.Name -ieq $env:COMPUTERNAME -or
        ([string]$_.Fqdn -split '\.')[0] -ieq $env:COMPUTERNAME
    })
    $server = if ($localServers.Count -gt 0) { $localServers[0] } else { $servers[0] }
    $versionText = [string]$server.AdminDisplayVersion
    if ($versionText -notmatch '(?i)Version\s+(?<Major>\d+)\.(?<Minor>\d+)\s+\(Build\s+(?<Build>\d+)\.(?<Revision>\d+)\)') {
        throw "Unable to identify the Exchange product from AdminDisplayVersion '$versionText'."
    }
    $major = [int]$Matches.Major
    $minor = [int]$Matches.Minor
    $build = [int]$Matches.Build
    $revision = [int]$Matches.Revision
    $product = Get-WorkerExchangeProductName -Major $major -Minor $minor -Build $build
    return [pscustomobject][ordered]@{
        ServerName = [string]$server.Name
        Fqdn = [string]$server.Fqdn
        Product = $product
        AdminDisplayVersion = $versionText
        Build = '{0}.{1}.{2}.{3}' -f $major,$minor,$build,$revision
        Edition = [string]$server.Edition
    }
}

if ($RecipientBatchMode) {
    try {
        Write-WorkerChildLog -Path $RecipientBatchLogPath -Level INFO -Message ("SMTP uniqueness child mode started; PowerShell={0}; computer={1}." -f $PSVersionTable.PSVersion, $env:COMPUTERNAME)
        if (-not (Test-Path -LiteralPath $RecipientBatchInputPath -PathType Leaf)) { throw "Recipient batch input file not found: $RecipientBatchInputPath" }
        if ([string]::IsNullOrWhiteSpace($RecipientBatchOutputPath)) { throw 'RecipientBatchOutputPath is required.' }
        $batchInput = Import-Clixml -LiteralPath $RecipientBatchInputPath
        $addresses = @(ConvertTo-TextArray $batchInput.Addresses)
        $recipientFilter = [string]$batchInput.Filter
        $testDelaySeconds = if ($batchInput.PSObject.Properties['TestDelaySeconds']) { [int]$batchInput.TestDelaySeconds } else { 0 }
        if ($testDelaySeconds -gt 0) {
            Write-WorkerChildLog -Path $RecipientBatchLogPath -Level INFO -Message ("Self-test delay requested: {0} second(s)." -f $testDelaySeconds)
            Start-Sleep -Seconds $testDelaySeconds
        }
        Write-WorkerChildLog -Path $RecipientBatchLogPath -Level INFO -Message ("Loading Exchange snap-in '{0}' for {1} address(es)." -f $RecipientBatchSnapInName, $addresses.Count)
        if (-not (Get-PSSnapin -Name $RecipientBatchSnapInName -ErrorAction SilentlyContinue)) {
            Add-PSSnapin -Name $RecipientBatchSnapInName -ErrorAction Stop
        }
        foreach ($requiredCommand in @('Set-ADServerSettings','Get-Recipient')) {
            if (-not (Get-Command -Name $requiredCommand -ErrorAction SilentlyContinue)) { throw "Required Exchange on-premises command '$requiredCommand' is unavailable." }
        }
        Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop | Out-Null
        Write-WorkerChildLog -Path $RecipientBatchLogPath -Level INFO -Message ("Executing forest-wide Get-Recipient filter; filterLength={0}." -f $recipientFilter.Length)
        $rows = @(Get-Recipient -Filter $recipientFilter -ResultSize Unlimited -ErrorAction Stop | Select-Object Identity,DistinguishedName,PrimarySmtpAddress,ExternalEmailAddress,EmailAddresses,RecipientType,RecipientTypeDetails)
        $rows | Export-Clixml -LiteralPath $RecipientBatchOutputPath -Depth 8 -Force
        Write-WorkerChildLog -Path $RecipientBatchLogPath -Level SUCCESS -Message ("Get-Recipient completed; rows={0}." -f $rows.Count)
        exit 0
    }
    catch {
        $message = $_ | Out-String
        Write-WorkerText -Path $RecipientBatchErrorPath -Text $message
        Write-WorkerChildLog -Path $RecipientBatchLogPath -Level ERROR -Message (Format-WorkerError $_)
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
}
if ($SelfTest) {
    Write-WorkerLog -Level INFO -Message "Self-test starting. Version=$workerVersion; PowerShell=$($PSVersionTable.PSVersion); DiagnosticsDirectory=$DiagnosticsDirectory."
    $selfTestPath = Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-ExchangeMigrationReadiness-PS5SelfTest-{0}.clixml" -f [guid]::NewGuid().ToString('N'))
    try {
        $items = New-Object System.Collections.Generic.List[object]
        $errors = New-Object System.Collections.Generic.List[object]
        [void]$items.Add([pscustomobject]@{ Name = 'serialization'; Result = 'PASS' })
        $commandErrors = New-Object System.Collections.Generic.List[object]
        $mailboxCommandErrors = New-Object System.Collections.Generic.List[object]
        Add-WorkerCommandFailure -CollectionErrors $commandErrors -MailboxErrors $mailboxCommandErrors -EmailAddress 'selftest@example.invalid' -CheckId 'INBOX-FORWARDING-RULES' -CommandName 'Get-InboxRule' -Result ([pscustomobject]@{ Success = $false; Rows = @(); ErrorMessage = 'Expected partial command error' })
        if ($commandErrors.Count -ne 1 -or $mailboxCommandErrors.Count -ne 1 -or $commandErrors[0].IsFatal) {
            throw 'Windows PowerShell 5.1 partial command error tracking self-test returned an unexpected result.'
        }
        [void]$items.Add([pscustomobject]@{ Name = 'partial-command-error-tracking'; Result = 'PASS' })
        if (
            (Test-WorkerMigrationRelevantDelegate -Identity 'NT AUTHORITY\SELF') -or
            (Test-WorkerMigrationRelevantDelegate -Identity 'DE\Exchange Trusted Subsystem') -or
            -not (Test-WorkerMigrationRelevantDelegate -Identity 'DE\50011SPDL01') -or
            -not (Test-WorkerMigrationRelevantDelegate -Identity 'S-1-5-21-111-222-333-444')
        ) {
            throw 'Migration-relevant delegate filtering self-test returned an unexpected result.'
        }
        [void]$items.Add([pscustomobject]@{ Name = 'migration-relevant-delegate-filtering'; Result = 'PASS' })
        $mappedProducts = @(
            Get-WorkerExchangeProductName -Major 15 -Minor 1 -Build 2507
            Get-WorkerExchangeProductName -Major 15 -Minor 2 -Build 1748
            Get-WorkerExchangeProductName -Major 15 -Minor 2 -Build 2562
        )
        if ($mappedProducts[0] -ne 'Exchange Server 2016' -or $mappedProducts[1] -ne 'Exchange Server 2019' -or $mappedProducts[2] -ne 'Exchange Server Subscription Edition') {
            throw 'Exchange on-premises build-to-product mapping self-test returned an unexpected result.'
        }
        [void]$items.Add([pscustomobject]@{ Name = 'exchange-product-version-mapping'; Result = 'PASS' })
        $batchMechanics = Invoke-WorkerRecipientAddressBatch -Addresses @('selftest@example.invalid') -TimeoutSeconds 5 -SnapInName 'SmartM365.Nonexistent.Exchange.SnapIn' -TestDelaySeconds 30 -DiagnosticsDirectory $DiagnosticsDirectory -BatchNumber 1
        if ($batchMechanics.Success -or -not $batchMechanics.TimedOut -or [string]::IsNullOrWhiteSpace([string]$batchMechanics.ErrorMessage) -or [double]$batchMechanics.DurationSeconds -gt 12) {
            throw 'Windows PowerShell 5.1 recipient batch timeout/termination self-test returned an unexpected result.'
        }
        [void]$items.Add([pscustomobject]@{ Name = 'recipient-batch-isolation'; Result = 'PASS' })
        [void]$errors.Add([pscustomobject]@{ EmailAddress = 'selftest@example.invalid'; Message = 'Expected self-test error record' })
        $payload = [pscustomobject][ordered]@{
            WorkerVersion = $workerVersion
            Items = $items.ToArray()
            Evidence = @((Get-WorkerFailureEvidence -EmailAddress 'selftest@example.invalid' -Message 'Expected self-test evidence'))
            Errors = $errors.ToArray()
        }
        $payload | Export-Clixml -LiteralPath $selfTestPath -Depth 12 -Force
        $roundTrip = Import-Clixml -LiteralPath $selfTestPath
        if (@($roundTrip.Items).Count -ne 5 -or [string]$roundTrip.Items[0].Result -ne 'PASS' -or @($roundTrip.Evidence).Count -ne 1 -or @($roundTrip.Errors).Count -ne 1) {
            throw 'Windows PowerShell 5.1 CLIXML nested collection round-trip returned an unexpected result.'
        }
        Write-WorkerLog -Level SUCCESS -Message 'Self-test completed successfully.'
        Write-Output "SELFTEST_OK|$workerVersion|$($PSVersionTable.PSVersion)"
        exit 0
    }
    catch {
        Write-WorkerLog -Level ERROR -Message "Self-test failed. $(Format-WorkerError $_)"
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
    finally {
        if (Test-Path -LiteralPath $selfTestPath) { Remove-Item -LiteralPath $selfTestPath -Force -ErrorAction SilentlyContinue }
    }
}

try {
    Write-WorkerLog -Level INFO -Message "Worker starting. Version=$workerVersion; PowerShell=$($PSVersionTable.PSVersion); Computer=$env:COMPUTERNAME; InputPath=$InputPath; OutputPath=$OutputPath; DiagnosticsDirectory=$DiagnosticsDirectory."
    $snapInTimer=[Diagnostics.Stopwatch]::StartNew()
    $snapInName = 'Microsoft.Exchange.Management.PowerShell.SnapIn'
    if (-not (Get-PSSnapin -Name $snapInName -ErrorAction SilentlyContinue)) {
        Add-PSSnapin -Name $snapInName -ErrorAction Stop
    }
    foreach ($requiredCommand in @('Set-ADServerSettings', 'Get-ExchangeServer', 'Get-Mailbox', 'Get-RemoteMailbox', 'Get-MailUser', 'Get-Recipient', 'Get-MailboxStatistics')) {
        if (-not (Get-Command -Name $requiredCommand -ErrorAction SilentlyContinue)) {
            throw "Required Exchange on-premises command '$requiredCommand' is unavailable after loading $snapInName."
        }
    }
    Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop | Out-Null
    $exchangeServerInfo = Get-WorkerExchangeServerInfo
    $snapInTimer.Stop()
    Write-WorkerLog -Level SUCCESS -Message "Exchange snap-in and required commands validated; Product=$($exchangeServerInfo.Product); Build=$($exchangeServerInfo.Build); Server=$($exchangeServerInfo.ServerName); Edition=$($exchangeServerInfo.Edition); ViewEntireForest enabled. DurationMs=$($snapInTimer.ElapsedMilliseconds); SnapIn=$snapInName."

    if ($ValidateOnly) {
        Write-Output "VALIDATION_OK|$workerVersion|$env:COMPUTERNAME|$($PSVersionTable.PSVersion)|$($exchangeServerInfo.Product)|$($exchangeServerInfo.Build)|$($exchangeServerInfo.ServerName)"
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
        $smtpBatchSize = if ($workerInput.PSObject.Properties['SmtpUniquenessBatchSize']) { [int]$workerInput.SmtpUniquenessBatchSize } else { 50 }
        $smtpBatchTimeoutSeconds = if ($workerInput.PSObject.Properties['SmtpUniquenessBatchTimeoutSeconds']) { [int]$workerInput.SmtpUniquenessBatchTimeoutSeconds } else { 60 }
    }
    else {
        # Backward compatibility with the v1.11.1 plain string-array payload.
        $emails = @($workerInput | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
        $smtpBatchSize = 25
        $smtpBatchTimeoutSeconds = 60
    }

    $smtpBatchSize = [math]::Max(1, [math]::Min(50, $smtpBatchSize))
    $smtpBatchTimeoutSeconds = [math]::Max(5, [math]::Min(300, $smtpBatchTimeoutSeconds))
    $startedAt = Get-Date
    Write-WorkerLog -Level INFO -Message "Runtime input loaded. MailboxCount=$($emails.Count); EnabledCheckCount=$($enabledChecks.Count); SmtpBatchSize=$smtpBatchSize; SmtpBatchTimeoutSeconds=$smtpBatchTimeoutSeconds."
    $evidence = New-Object System.Collections.Generic.List[object]
    $collectionErrors = New-Object System.Collections.Generic.List[object]
    $databaseCache = @{}
    $candidateAddressesByEmail = @{}
    $addressConflictMetrics = [pscustomobject]@{ CandidateAddressCount=0; BatchCount=0; TimeoutCount=0; ErrorCount=0; BatchSize=$smtpBatchSize; TimeoutSeconds=$smtpBatchTimeoutSeconds; DurationSeconds=0; DiagnosticsDirectory=$DiagnosticsDirectory; ChildLogPaths=@() }
    $index = 0
    foreach ($email in $emails) {
        if (Test-WorkerCancellation) { throw [OperationCanceledException]::new('Exchange on-premises worker cancellation requested.') }
        $index++
        $mailboxTimer=[Diagnostics.Stopwatch]::StartNew()
        $message = ''
        $mailboxErrors = New-Object System.Collections.Generic.List[object]
        Write-WorkerLog -Level INFO -Message "Mailbox collection starting. Index=$index/$($emails.Count); EmailAddress=$email."
        Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Recipient lookup - {0}" -f $email)
        try {
            $safeEmail = $email.Replace("'", "''")
            $identityFilter = "EmailAddresses -eq 'smtp:$safeEmail'"
            $recipientResult = Invoke-WorkerCommand -Name 'Get-Recipient' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' }
            Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'ONPREM-RECIPIENT-STATE' -CommandName 'Get-Recipient' -Result $recipientResult
            $recipientTypes = @($recipientResult.Rows | ForEach-Object { [string]$_.RecipientTypeDetails } | Where-Object { $_ } | Sort-Object -Unique)
            $successfulEmpty = [pscustomobject]@{ Success = $true; Rows = @(); ErrorMessage = '' }
            $queryAllTypes = -not $recipientResult.Success
            $needMailbox = $queryAllTypes -or @($recipientTypes | Where-Object { $_ -match '^(User|Shared|Room|Equipment|Linked|Discovery)Mailbox$' }).Count -gt 0
            $needRemoteMailbox = $queryAllTypes -or @($recipientTypes | Where-Object { $_ -match '^Remote.*Mailbox$' }).Count -gt 0
            $needMailUser = $queryAllTypes -or $recipientTypes -contains 'MailUser'

            $mailboxResult = if ($needMailbox) { Invoke-WorkerCommand -Name 'Get-Mailbox' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' } } else { $successfulEmpty }
            $remoteMailboxResult = if ($needRemoteMailbox) { Invoke-WorkerCommand -Name 'Get-RemoteMailbox' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' } } else { $successfulEmpty }
            $mailUserResult = if ($needMailUser) { Invoke-WorkerCommand -Name 'Get-MailUser' -Parameters @{ Filter = $identityFilter; ResultSize = 'Unlimited' } } else { $successfulEmpty }
            Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'ONPREM-RECIPIENT-STATE' -CommandName 'Get-Mailbox' -Result $mailboxResult
            Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'ONPREM-RECIPIENT-STATE' -CommandName 'Get-RemoteMailbox' -Result $remoteMailboxResult
            Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'ONPREM-RECIPIENT-STATE' -CommandName 'Get-MailUser' -Result $mailUserResult
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
            $inboxRulesComplete = $false
            $databaseHealth = @()
            $databaseHealthAvailable = $false
            $permissions = New-Object System.Collections.Generic.List[object]
            $mailboxPermissionSuccess = $false
            $sendAsPermissionSuccess = $false

            if ($mailboxes.Count -eq 1) {
                $identity = $mailboxes[0].Identity
                Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Mailbox statistics - {0}" -f $email)
                $statisticsResult = Invoke-WorkerCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $identity }
                Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'MAILBOX-TARGET-QUOTA' -CommandName 'Get-MailboxStatistics' -Result $statisticsResult
                $statistics = @($statisticsResult.Rows | ForEach-Object { ConvertTo-StatisticsEvidence $_ })
                $statisticsAvailable = [bool]$statisticsResult.Success

                if (Test-WorkerCheckEnabled -CheckId 'ARCHIVE-READINESS') {
                    $hasArchive = $mailboxes[0].ArchiveStatus -match '^(?i:Active|HostedPending|Local)$' -or ($mailboxes[0].ArchiveGuid -and $mailboxes[0].ArchiveGuid -notmatch '^0{8}-')
                    if ($hasArchive) {
                        $archiveResult = Invoke-WorkerCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $identity; Archive = $true }
                        Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'ARCHIVE-READINESS' -CommandName 'Get-MailboxStatistics -Archive' -Result $archiveResult
                        $archiveStatistics = @($archiveResult.Rows | ForEach-Object { ConvertTo-StatisticsEvidence $_ })
                        $archiveStatisticsAvailable = [bool]$archiveResult.Success
                    }
                    else { $archiveStatisticsAvailable = $true }
                }

                if ((Test-WorkerCheckEnabled -CheckId 'MAILBOX-RECOVERABLE-ITEMS-QUOTA') -or (Test-WorkerCheckEnabled -CheckId 'MAILBOX-FOLDER-LIMITS')) {
                    Write-WorkerProgress -Current $index -Total $emails.Count -Message ("Folder statistics - {0}" -f $email)
                    $folderResult = Invoke-WorkerCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $identity }
                    Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'MAILBOX-FOLDER-LIMITS' -CommandName 'Get-MailboxFolderStatistics' -Result $folderResult
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
                    $inboxRulesComplete = [bool]$ruleResult.Success
                    if (-not $ruleResult.Success) {
                        Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'INBOX-FORWARDING-RULES' -CommandName 'Get-InboxRule -IncludeHidden' -Result $ruleResult
                        Write-WorkerLog -Level WARN -Message "Get-InboxRule -IncludeHidden failed for $email; retrying visible rules without IncludeHidden. Error=$($ruleResult.ErrorMessage)"
                        $fallbackRuleResult = Invoke-WorkerCommand -Name 'Get-InboxRule' -Parameters @{ Mailbox = $identity }
                        if ($fallbackRuleResult.Success) {
                            $ruleResult = $fallbackRuleResult
                            Write-WorkerLog -Level WARN -Message "Visible inbox-rule fallback succeeded for $email; hidden-rule coverage remains incomplete. RowCount=$(@($ruleResult.Rows).Count)."
                        }
                        else {
                            Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'INBOX-FORWARDING-RULES' -CommandName 'Get-InboxRule fallback' -Result $fallbackRuleResult
                            $ruleResult = $fallbackRuleResult
                        }
                    }
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
                        Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'EXCHANGE-DATABASE-HEALTH' -CommandName 'Get-MailboxDatabase' -Result $databaseResult
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
                    Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'PERMISSIONS-BASELINE' -CommandName 'Get-MailboxPermission' -Result $mailboxPermissionResult
                    $mailboxPermissionSuccess = [bool]$mailboxPermissionResult.Success
                    foreach ($permission in @($mailboxPermissionResult.Rows)) {
                        $rights = @(ConvertTo-TextArray $permission.AccessRights)
                        $delegate = [string]$permission.User
                        if ($permission.IsInherited -or $rights -notcontains 'FullAccess' -or -not (Test-WorkerMigrationRelevantDelegate -Identity $delegate)) { continue }
                        [void]$permissions.Add([pscustomobject]@{ PermissionType = 'FullAccess'; Delegate = $delegate; IsInherited = $false; Source = 'ExchangeOnPremWorker' })
                    }

                    $sendAsPermissionResult = Invoke-WorkerCommand -Name 'Get-ADPermission' -Parameters @{ Identity = $mailboxes[0].DistinguishedName }
                    Add-WorkerCommandFailure -CollectionErrors $collectionErrors -MailboxErrors $mailboxErrors -EmailAddress $email -CheckId 'PERMISSIONS-BASELINE' -CommandName 'Get-ADPermission' -Result $sendAsPermissionResult
                    $sendAsPermissionSuccess = [bool]$sendAsPermissionResult.Success
                    foreach ($permission in @($sendAsPermissionResult.Rows)) {
                        $delegate = [string]$permission.User
                        if ($permission.IsInherited -or @(ConvertTo-TextArray $permission.ExtendedRights) -notcontains 'Send-As' -or $permission.Deny -or -not (Test-WorkerMigrationRelevantDelegate -Identity $delegate)) { continue }
                        [void]$permissions.Add([pscustomobject]@{ PermissionType = 'SendAs'; Delegate = $delegate; IsInherited = $false; Source = 'ExchangeOnPremWorker' })
                    }
                    foreach ($delegate in @($mailboxes[0].GrantSendOnBehalfTo)) {
                        if (-not (Test-WorkerMigrationRelevantDelegate -Identity ([string]$delegate))) { continue }
                        [void]$permissions.Add([pscustomobject]@{ PermissionType = 'SendOnBehalf'; Delegate = [string]$delegate; IsInherited = $false; Source = 'ExchangeOnPremWorker' })
                    }
                }
            }

            $candidateAddressValues = New-Object System.Collections.Generic.List[object]
            foreach ($mailObject in @($mailboxes) + @($remoteMailboxes) + @($mailUsers) + @($recipients)) {
                foreach ($value in @($mailObject.EmailAddresses) + @($mailObject.ExternalEmailAddress) + @($mailObject.PrimarySmtpAddress)) {
                    if ($value) { [void]$candidateAddressValues.Add($value) }
                }
            }
            $candidateAddressesByEmail[$email] = @(
                $candidateAddressValues.ToArray() |
                    ForEach-Object { ConvertTo-NormalizedSmtpAddress $_ } |
                    Where-Object { $_ } |
                    Sort-Object -Unique
            )
            $lookupResults = @($recipientResult)
            if ($needMailbox) { $lookupResults += $mailboxResult }
            if ($needRemoteMailbox) { $lookupResults += $remoteMailboxResult }
            if ($needMailUser) { $lookupResults += $mailUserResult }
            $coreAvailable = @($lookupResults | Where-Object { -not $_.Success }).Count -eq 0
            $errors = @($lookupResults | Where-Object { -not $_.Success -and $_.ErrorMessage } | ForEach-Object { $_.ErrorMessage } | Sort-Object -Unique)
            if ($mailboxErrors.Count -gt 0) {
                $message = 'Partial evidence unavailable: ' + (@($mailboxErrors | ForEach-Object { "$($_.Command): $($_.Message)" } | Select-Object -Unique) -join ' | ')
            }
            [void]$evidence.Add([pscustomobject][ordered]@{
                EmailAddress = $email
                Available = $coreAvailable
                Source = "Local Exchange on-premises Management Shell on $env:COMPUTERNAME"
                SourceTimestamp = Get-Date
                Message = if (-not $coreAvailable) { 'Exchange on-premises lookup failed: ' + ($errors -join ' | ') } elseif ($mailboxErrors.Count -gt 0) { $message } else { 'Exchange on-premises evidence collected through direct local cmdlets.' }
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
                InboxRulesComplete = $inboxRulesComplete
                DatabaseHealth = $databaseHealth
                DatabaseHealthAvailable = $databaseHealthAvailable
                DatabaseHealthSource = "Local Exchange on-premises Management Shell mailbox database on $env:COMPUTERNAME"
                DatabaseHealthSourceTimestamp = Get-Date
                DeliveryRestrictionsAvailable = [bool]($mailboxes.Count -eq 1 -and $mailboxResult.Success)
                AddressConflicts = @()
                PartialErrors = $mailboxErrors.ToArray()
            })
        }
        catch {
            $message = $_.Exception.Message
            Write-WorkerLog -Level ERROR -Message "Mailbox collection failed. Index=$index/$($emails.Count); EmailAddress=$email; $(Format-WorkerError $_)"
            [void]$collectionErrors.Add([pscustomobject][ordered]@{ EmailAddress = $email; CheckId = 'ONPREM-SOURCE'; Command = 'MailboxCollection'; Message = $message; IsFatal = $true })
            [void]$evidence.Add((Get-WorkerFailureEvidence -EmailAddress $email -Message $message))
        }
        finally {
            $mailboxTimer.Stop()
            Write-WorkerLog -Level $(if($message){'WARN'}else{'SUCCESS'}) -Message "Mailbox collection ended. Index=$index/$($emails.Count); EmailAddress=$email; DurationMs=$($mailboxTimer.ElapsedMilliseconds); EvidenceCount=$($evidence.Count); Error=$message."
        }
    }
    if ((Test-WorkerCheckEnabled -CheckId 'PROXY-SMTP-GLOBAL-UNIQUE') -or (Test-WorkerCheckEnabled -CheckId 'TARGET-ADDRESS-GLOBAL-UNIQUE')) {
        Write-WorkerLog -Level INFO -Message 'SMTP uniqueness collection starting.'
        $addressConflictMetrics = Initialize-WorkerAddressConflictEvidence -Evidence $evidence -CandidateAddressesByEmail $candidateAddressesByEmail -BatchSize $smtpBatchSize -TimeoutSeconds $smtpBatchTimeoutSeconds -SnapInName $snapInName -DiagnosticsDirectory $DiagnosticsDirectory
        Write-WorkerLog -Level $(if($addressConflictMetrics.ErrorCount){'WARN'}else{'SUCCESS'}) -Message "SMTP uniqueness collection ended. CandidateAddressCount=$($addressConflictMetrics.CandidateAddressCount); BatchCount=$($addressConflictMetrics.BatchCount); TimeoutCount=$($addressConflictMetrics.TimeoutCount); ErrorCount=$($addressConflictMetrics.ErrorCount); DurationSeconds=$($addressConflictMetrics.DurationSeconds)."
    }
    if (Test-WorkerCancellation) { throw [OperationCanceledException]::new('Exchange on-premises worker cancellation requested.') }
    Write-WorkerLog -Level INFO -Message 'Hybrid evidence collection starting.'
    $hybridEvidence = Get-HybridEvidence
    Write-WorkerLog -Level SUCCESS -Message "Hybrid evidence collection ended. MrsProxyAvailable=$($hybridEvidence.MrsProxyAvailable); MrsProxyEnabled=$($hybridEvidence.MrsProxyEnabled); OAuthAvailable=$($hybridEvidence.OAuthAvailable); OAuthHealthy=$($hybridEvidence.OAuthHealthy)."
    $result = [pscustomobject][ordered]@{
        WorkerVersion = $workerVersion
        ComputerName = $env:COMPUTERNAME
        PowerShellVersion = [string]$PSVersionTable.PSVersion
        SnapInName = $snapInName
        ExchangeProduct = $exchangeServerInfo.Product
        ExchangeBuild = $exchangeServerInfo.Build
        ExchangeServerName = $exchangeServerInfo.ServerName
        ExchangeServerFqdn = $exchangeServerInfo.Fqdn
        ExchangeEdition = $exchangeServerInfo.Edition
        CollectedAt = Get-Date
        DurationSeconds = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 1)
        MailboxEvidenceCount = $evidence.Count
        MailboxObjectCount = [int](@($evidence.ToArray() | ForEach-Object { @($_.Mailboxes).Count } | Measure-Object -Sum).Sum)
        PermissionCount = @($evidence.ToArray() | ForEach-Object { @($_.Permissions) }).Count
        ErrorCount = $collectionErrors.Count
        PartialErrorCount = @($collectionErrors | Where-Object { -not $_.IsFatal }).Count
        FatalErrorCount = @($collectionErrors | Where-Object { $_.IsFatal }).Count
        Errors = $collectionErrors.ToArray()
        SmtpUniquenessCandidateAddressCount = $addressConflictMetrics.CandidateAddressCount
        SmtpUniquenessBatchCount = $addressConflictMetrics.BatchCount
        SmtpUniquenessTimeoutCount = $addressConflictMetrics.TimeoutCount
        SmtpUniquenessErrorCount = $addressConflictMetrics.ErrorCount
        SmtpUniquenessBatchSize = $addressConflictMetrics.BatchSize
        SmtpUniquenessBatchTimeoutSeconds = $addressConflictMetrics.TimeoutSeconds
        SmtpUniquenessDurationSeconds = $addressConflictMetrics.DurationSeconds
        SmtpUniquenessDiagnosticsDirectory = $addressConflictMetrics.DiagnosticsDirectory
        SmtpUniquenessChildLogPaths = @($addressConflictMetrics.ChildLogPaths)
        Evidence = $evidence.ToArray()
        Hybrid = $hybridEvidence
    }
    $result | Export-Clixml -LiteralPath $OutputPath -Depth 12 -Force
    Write-WorkerLog -Level $(if($collectionErrors.Count){'WARN'}else{'SUCCESS'}) -Message "Worker evidence exported. MailboxEvidenceCount=$($evidence.Count); ErrorCount=$($collectionErrors.Count); PartialErrorCount=$($result.PartialErrorCount); FatalErrorCount=$($result.FatalErrorCount); TotalDurationSeconds=$($result.DurationSeconds); OutputPath=$OutputPath."
    Write-WorkerProgress -Current $emails.Count -Total $emails.Count -Message 'Complete'
}
catch {
    $diagnostic=Format-WorkerError $_
    Write-WorkerLog -Level ERROR -Message "Fatal worker error. $diagnostic"
    Write-WorkerText -Path $ErrorPath -Text "$(($_ | Out-String))`r`n$diagnostic"
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCz3AUDMhFjUliN
# fRajytQayRZSsuGsG3xcQffCTmCjMaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD1XHMqREtFOOe/x/l/Pi0a
# Lx02awfCfPle78pKQQakCjANBgkqhkiG9w0BAQEFAASCAYBsXbnJImvDdSzjDp1f
# 8gYDxTjnXm7UdprU09YkGHM3dgLICX+KFS4GIhcW8p4htKwxZb0dUHkj9pnRNrQs
# KsgRkBphHeEIQh5iTnqQuscekk6bXvsHQ2UnsA+Jx7UByskDnqZI9eVxP1JJ6ZD7
# KDZFjY9xiE0JKNbh11okl5jG1Jd7k54FgiBGBf0Kb9ckw6cA5cwaDO9bZ4WjztoT
# xo9sugcJL34ffIWoRRDUT4gakL3+gPFK5TEUfq4DYFPOJ7wVFdg2I9CCwV782nfn
# sdQpINYPRbGAlryURQUtmtjj3O3gr5czq44NcSSq+38bR6/gfhwT1dNt44bMj/iW
# Rx9c+TlwV7uacigWBSjD9LvrO2aLoKQ2ko+XVf8CKayRwWnpCLbTUzRu+9gJlhec
# hALKofoiqmdJKOupk2cEFNimYWiSmKw4EbDiQz3YOaYrYDMNVPHf1eTxDxV4QUbp
# xuLqZnnMLNSaijKSaw0nuMY9JvfuxEGnXMqghTzzvNYH9ro=
# SIG # End signature block
