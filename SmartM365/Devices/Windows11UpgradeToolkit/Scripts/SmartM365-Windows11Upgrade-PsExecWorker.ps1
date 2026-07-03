<#
.SYNOPSIS
    Internal worker used by SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1.

.DESCRIPTION
    Runs one remote computer operation in a background job. This file is operator-side only;
    the target device still receives only SmartM365-Invoke-Windows11UpgradeRepair.ps1.

.VERSION
0.1.4
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Computer,
    [Parameter(Mandatory = $true)][int]$CycleNumber,
    [Parameter(Mandatory = $true)][string]$RemoteScriptArgsJson,
    [string]$ResolvedPsExecPath,
    [Parameter(Mandatory = $true)][string]$LocalScriptPath,
    [Parameter(Mandatory = $true)][string]$RemoteBaseDir,
    [Parameter(Mandatory = $true)][string]$RemoteScriptPath,
    [Parameter(Mandatory = $true)][string]$RemoteSetupCacheRoot,
    [Parameter(Mandatory = $true)][string]$LogRoot,
    [Parameter(Mandatory = $true)][string]$CentralLogRoot,
    [string]$SetupSourcePath,
    [string]$SetupSourceMapPath,
    [ValidateSet('LocalCache','Share','Auto')]
    [string]$SetupExecutionMode = 'LocalCache',
    [string]$SetupMediaId = 'Win11',
    [string]$SetupLanguage = 'MatchSystem',
    [bool]$AllowSetupUpgrade = $false,
    [bool]$SkipSetupMediaPreCopy = $false,
    [bool]$SkipVirtualMachines = $false,
    [bool]$DryRun = $false,
    [bool]$NoCentralLogCollection = $false,
    [bool]$KeepCentralLogHistory = $false,
    [int]$PsExecTimeoutMinutes = 180,
    [string]$GlobalWorkerLeasePath,
    [string]$GlobalWorkerLeaseMutexName
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$RemoteScriptArgs = @()
if (-not [string]::IsNullOrWhiteSpace($RemoteScriptArgsJson)) {
    try {
        $decodedRemoteScriptArgs = $RemoteScriptArgsJson | ConvertFrom-Json -ErrorAction Stop
        if ($decodedRemoteScriptArgs -and $decodedRemoteScriptArgs.PSObject.Properties['Args']) {
            $RemoteScriptArgs = @($decodedRemoteScriptArgs.Args | ForEach-Object { [string]$_ })
        }
    }
    catch {
        throw ("Invalid remote script argument payload: {0}" -f $_.Exception.Message)
    }
}

function Invoke-WithWorkerLeaseMutex {
    param(
        [AllowNull()][string]$MutexName,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock
    )

    if ([string]::IsNullOrWhiteSpace($MutexName)) {
        & $ScriptBlock
        return
    }

    $mutex = New-Object System.Threading.Mutex($false, $MutexName)
    $acquired = $false
    try {
        try { $acquired = $mutex.WaitOne(30000) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if ($acquired) { & $ScriptBlock }
    }
    finally {
        if ($acquired) { try { $mutex.ReleaseMutex() } catch { } }
        $mutex.Dispose()
    }
}

function Update-WorkerLease {
    param(
        [AllowNull()][string]$LeasePath,
        [AllowNull()][string]$MutexName
    )

    if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }

    try {
        Invoke-WithWorkerLeaseMutex -MutexName $MutexName -ScriptBlock {
            if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
            $data = Get-Content -LiteralPath $LeasePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $processName = ''
            try { $processName = (Get-Process -Id $PID -ErrorAction Stop).ProcessName } catch { }
            $data | Add-Member -NotePropertyName WorkerProcessId -NotePropertyValue $PID -Force
            $data | Add-Member -NotePropertyName WorkerProcessName -NotePropertyValue $processName -Force
            $data | Add-Member -NotePropertyName WorkerStartedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
            $data | Add-Member -NotePropertyName LastUpdatedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
            $data | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $LeasePath -Encoding UTF8 -Force
        }
    }
    catch { }
}

Update-WorkerLease -LeasePath $GlobalWorkerLeasePath -MutexName $GlobalWorkerLeaseMutexName

function New-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Test-TcpPort {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Port,
        [int]$TimeoutMilliseconds = 3000
    )

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)) {
            return $false
        }
        $client.EndConnect($async)
        return $true
    }
    catch {
        return $false
    }
    finally {
        try { $client.Close() } catch { }
    }
}

function Test-RemoteAdminAccess {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    $state = [ordered]@{
        DnsResolved = $false
        DnsAddressList = ''
        PingReachable = $false
        SmbPort445Reachable = $false
        AdminShareReachable = $false
        AdminSharePath = "\\$ComputerName\ADMIN$"
        RootSharePath = "\\$ComputerName\C$"
        FailureType = ''
        Detail = ''
    }

    try {
        $dns = [System.Net.Dns]::GetHostEntry($ComputerName)
        $state.DnsResolved = $true
        $state.DnsAddressList = (($dns.AddressList | ForEach-Object { $_.IPAddressToString }) -join ';')
        Add-Content -LiteralPath $LogPath -Value ("[{0}] DNS resolved: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$state.DnsAddressList) -Encoding UTF8
    }
    catch {
        Add-Content -LiteralPath $LogPath -Value ("[{0}] WARN DNS resolution failed: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message) -Encoding UTF8
    }

    $state.PingReachable = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $state.PingReachable) {
        Add-Content -LiteralPath $LogPath -Value ("[{0}] WARN Ping failed. SMB checks will still run." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    }

    $state.SmbPort445Reachable = Test-TcpPort -ComputerName $ComputerName -Port 445
    if ($state.SmbPort445Reachable) {
        Add-Content -LiteralPath $LogPath -Value ("[{0}] SMB TCP 445 reachable." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    }
    else {
        Add-Content -LiteralPath $LogPath -Value ("[{0}] WARN SMB TCP 445 is not reachable." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    }

    $adminShareReachable = Test-Path -LiteralPath ([string]$state.AdminSharePath) -ErrorAction SilentlyContinue
    $rootShareReachable = Test-Path -LiteralPath ([string]$state.RootSharePath) -ErrorAction SilentlyContinue
    $state.AdminShareReachable = ($adminShareReachable -and $rootShareReachable)
    Add-Content -LiteralPath $LogPath -Value ("[{0}] Administrative shares: ADMIN$={1}; C$={2}; ADMINPath={3}; RootPath={4}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$adminShareReachable,$rootShareReachable,$state.AdminSharePath,$state.RootSharePath) -Encoding UTF8

    if (-not $state.AdminShareReachable) {
        if (-not $state.DnsResolved) {
            $state.FailureType = 'DNS_FAILED'
        }
        elseif (-not $state.SmbPort445Reachable) {
            $state.FailureType = 'SMB_PORT_445_UNREACHABLE'
        }
        elseif (-not $state.PingReachable) {
            $state.FailureType = 'PING_FAILED_ADMIN_SHARE_FAILED'
        }
        else {
            $state.FailureType = 'PING_OK_ADMIN_SHARE_FAILED'
        }
    }

    $state.Detail = ("DNS={0}; Addresses={1}; Ping={2}; Tcp445={3}; AdminShare={4}; FailureType={5}" -f $state.DnsResolved,$state.DnsAddressList,$state.PingReachable,$state.SmbPort445Reachable,$state.AdminShareReachable,$state.FailureType)
    return [pscustomobject]$state
}

function Set-ResultRemoteAccessState {
    param(
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][object]$State
    )

    $Result.DnsResolved = [string]$State.DnsResolved
    $Result.DnsAddressList = [string]$State.DnsAddressList
    $Result.PingReachable = [string]$State.PingReachable
    $Result.SmbPort445Reachable = [string]$State.SmbPort445Reachable
    $Result.AdminShareReachable = [string]$State.AdminShareReachable
    $Result.AdminShareFailureType = [string]$State.FailureType
}

function Convert-ToAdminSharePath {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$LocalPath
    )

    $full = [System.IO.Path]::GetFullPath($LocalPath)
    if ($full -notmatch '^[A-Za-z]:\\') {
        throw "Only local drive paths can be converted to admin share paths: $LocalPath"
    }
    $drive = $full.Substring(0,1)
    $rest = $full.Substring(3)
    return "\\$ComputerName\$drive`$\$rest"
}

function Copy-DirectoryContent {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )

    New-Directory -Path $DestinationPath
    Copy-Item -Path (Join-Path $SourcePath '*') -Destination $DestinationPath -Recurse -Force -ErrorAction Stop
}

function Copy-RemotePayload {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $ComputerName -LocalPath $RemoteBaseDir
    New-Directory -Path $remoteBaseShare
    $remoteScriptShare = Join-Path $remoteBaseShare (Split-Path -Leaf $RemoteScriptPath)
    Copy-Item -LiteralPath $LocalScriptPath -Destination $remoteScriptShare -Force -ErrorAction Stop

    $localHash = (Get-FileHash -LiteralPath $LocalScriptPath -Algorithm SHA256).Hash
    $remoteHash = (Get-FileHash -LiteralPath $remoteScriptShare -Algorithm SHA256).Hash
    if ($localHash -ne $remoteHash) {
        throw "Remote script copy hash mismatch."
    }

    Add-Content -LiteralPath $LogPath -Value ("[{0}] Remote script copied and verified." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
}

function Copy-SetupMediaToRemoteCache {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    if (-not $AllowSetupUpgrade) { return 'NotRequested' }
    if ($SkipSetupMediaPreCopy) { return 'ExistingMediaOnly' }
    if ($SetupExecutionMode -eq 'Share') { return 'ShareMode' }
    if ([string]::IsNullOrWhiteSpace($SetupSourcePath) -and [string]::IsNullOrWhiteSpace($SetupSourceMapPath)) { return 'TargetCacheNoSource' }

    Add-Content -LiteralPath $LogPath -Value ("[{0}] Setup media copy is target-side. SourcePath/Map will be validated from the target SYSTEM context: Source={1}; Map={2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$SetupSourcePath,$SetupSourceMapPath) -Encoding UTF8
    return 'TargetSideCache'
}

function Convert-ToPsExecRemoteArgument {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '""' }
    $text = [string]$Value
    if ($text.Length -eq 0) { return '""' }
    if ($text -match '^[A-Za-z0-9_:\\./@=-]+$') { return $text }
    return ('"{0}"' -f ($text -replace '"', '\"'))
}

$script:CentralLogBuckets = @('Success','ADMIN_SHARE_UNREACHABLE','Errors')

function Get-ResultValue {
    param(
        [Parameter(Mandatory = $true)][object]$Result,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($Result -is [System.Collections.IDictionary] -and $Result.Contains($Name)) {
        return [string]$Result[$Name]
    }
    if ($Result.PSObject.Properties[$Name]) {
        return [string]$Result.$Name
    }
    return ''
}

function Get-CentralLogBucket {
    param([Parameter(Mandatory = $true)][object]$Result)

    $launcherStatus = Get-ResultValue -Result $Result -Name 'LauncherStatus'
    if ($launcherStatus -eq 'ADMIN_SHARE_UNREACHABLE' -or $launcherStatus -eq 'DRYRUN_ADMIN_SHARE_UNREACHABLE') {
        return 'ADMIN_SHARE_UNREACHABLE'
    }

    $exitCodeText = Get-ResultValue -Result $Result -Name 'ExitCode'
    $exitCodeValue = 0
    if ([int]::TryParse($exitCodeText, [ref]$exitCodeValue) -and $exitCodeValue -eq 0) {
        return 'Success'
    }

    return 'Errors'
}

function New-CentralLogTarget {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Cycle,
        [Parameter(Mandatory = $true)][ValidateSet('Success','ADMIN_SHARE_UNREACHABLE','Errors')][string]$Bucket
    )

    if ($KeepCentralLogHistory) {
        return (Join-Path (Join-Path (Join-Path $CentralLogRoot $Bucket) $ComputerName) ("Cycle{0}_{1}" -f $Cycle,(Get-Date -Format 'yyyyMMdd-HHmmss')))
    }

    foreach ($otherBucket in $script:CentralLogBuckets) {
        $otherLatest = Join-Path (Join-Path (Join-Path $CentralLogRoot $otherBucket) $ComputerName) 'Latest'
        if (Test-Path -LiteralPath $otherLatest) {
            Remove-Item -LiteralPath $otherLatest -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    return (Join-Path (Join-Path (Join-Path $CentralLogRoot $Bucket) $ComputerName) 'Latest')
}

function Publish-LauncherEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Cycle,
        [Parameter(Mandatory = $true)][ValidateSet('Success','ADMIN_SHARE_UNREACHABLE','Errors')][string]$Bucket,
        [Parameter(Mandatory = $true)][string]$WorkerLogPath,
        [AllowNull()][string]$StdoutLogPath,
        [AllowNull()][string]$StderrLogPath
    )

    if ($NoCentralLogCollection) { return '' }

    $target = New-CentralLogTarget -ComputerName $ComputerName -Cycle $Cycle -Bucket $Bucket
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Directory -Path $target

    if (Test-Path -LiteralPath $WorkerLogPath -PathType Leaf) {
        Copy-Item -LiteralPath $WorkerLogPath -Destination (Join-Path $target (Split-Path -Leaf $WorkerLogPath)) -Force -ErrorAction SilentlyContinue
    }
    foreach ($extraLog in @($StdoutLogPath,$StderrLogPath)) {
        if (-not [string]::IsNullOrWhiteSpace($extraLog) -and (Test-Path -LiteralPath $extraLog -PathType Leaf)) {
            Copy-Item -LiteralPath $extraLog -Destination (Join-Path $target (Split-Path -Leaf $extraLog)) -Force -ErrorAction SilentlyContinue
        }
    }
    return $target
}

function Update-ResultFromLastRun {
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Result,
        [Parameter(Mandatory = $true)][object]$LastRun
    )

    $Result.RemoteStatus = [string]$LastRun.Status
    $Result.RemoteNextAction = [string]$LastRun.NextAction
    if ($Result.RemoteStatus) {
        $Result.LauncherStatus = $Result.RemoteStatus
        $Result.ExitCode = [string]$LastRun.ExitCode
    }
    if ($LastRun.PSObject.Properties['SetupCacheAction']) {
        $Result.SetupCacheAction = [string]$LastRun.SetupCacheAction
    }
    foreach ($propertyName in @('SetupDynamicUpdate','SelectedSetupSourcePath','SetupSourceSelectionDetail','DiskCleanupAction','DiskCleanupFreedGB','AdvancedDiskCleanupAction','AdvancedDiskCleanupFreedGB','DismCleanupAction','DismCleanupFreedGB','SetupCompletionRebootAction','SetupCompletionRebootDetail','SetupCompletionRebootUserCount','SetupCompletionRebootUsers','ControlledRebootAction','ControlledRebootDetail','ControlledRebootUserCount','ControlledRebootUsers')) {
        if ($LastRun.PSObject.Properties[$propertyName]) {
            $Result[$propertyName] = [string]$LastRun.$propertyName
        }
    }
}

function Collect-RemoteEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Cycle,
        [Parameter(Mandatory = $true)][ValidateSet('Success','ADMIN_SHARE_UNREACHABLE','Errors')][string]$Bucket
    )

    if ($NoCentralLogCollection) { return '' }

    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $ComputerName -LocalPath $RemoteBaseDir
    if (-not (Test-Path -LiteralPath $remoteBaseShare -PathType Container)) { return '' }

    $target = New-CentralLogTarget -ComputerName $ComputerName -Cycle $Cycle -Bucket $Bucket
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Directory -Path $target

    foreach ($child in @('Logs','Output','LastRun.json')) {
        $source = Join-Path $remoteBaseShare $child
        if (Test-Path -LiteralPath $source) {
            Copy-Item -LiteralPath $source -Destination $target -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    return $target
}
New-Directory -Path $LogRoot
$logPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}.log" -f $Computer,$CycleNumber,(Get-Date -Format 'yyyyMMdd-HHmmss'))
$stdoutPath = "$logPath.stdout.txt"
$stderrPath = "$logPath.stderr.txt"

$result = [ordered]@{
    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    ComputerName = $Computer
    CycleNumber = $CycleNumber
    LauncherStatus = 'UNKNOWN'
    RemoteStatus = ''
    RemoteNextAction = ''
    ExitCode = ''
    Detail = ''
    DnsResolved = ''
    DnsAddressList = ''
    PingReachable = ''
    SmbPort445Reachable = ''
    AdminShareReachable = ''
    AdminShareFailureType = ''
    RemotePayloadCopyAttempts = ''
    SetupCacheAction = ''
    SetupDynamicUpdate = ''
    SelectedSetupSourcePath = ''
    SetupSourceSelectionDetail = ''
    DiskCleanupAction = ''
    DiskCleanupFreedGB = ''
    AdvancedDiskCleanupAction = ''
    AdvancedDiskCleanupFreedGB = ''
    DismCleanupAction = ''
    DismCleanupFreedGB = ''
    SetupCompletionRebootAction = ''
    SetupCompletionRebootDetail = ''
    SetupCompletionRebootUserCount = ''
    SetupCompletionRebootUsers = ''
    ControlledRebootAction = ''
    ControlledRebootDetail = ''
    ControlledRebootUserCount = ''
    ControlledRebootUsers = ''
    RemoteLogsPath = ''
    PsExecLogPath = $logPath
}

try {
    Add-Content -LiteralPath $logPath -Value ("[{0}] Starting {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Computer) -Encoding UTF8

    $remoteAccess = Test-RemoteAdminAccess -ComputerName $Computer -LogPath $logPath
    Set-ResultRemoteAccessState -Result $result -State $remoteAccess

    if ($DryRun) {
        $result.LauncherStatus = if ($remoteAccess.AdminShareReachable) { 'DRYRUN_READY' } else { 'DRYRUN_ADMIN_SHARE_UNREACHABLE' }
        $result.Detail = $remoteAccess.Detail
        if (-not $remoteAccess.AdminShareReachable) {
            $result.RemoteLogsPath = Publish-LauncherEvidence -ComputerName $Computer -Cycle $CycleNumber -Bucket 'ADMIN_SHARE_UNREACHABLE' -WorkerLogPath $logPath -StdoutLogPath $stdoutPath -StderrLogPath $stderrPath
        }
        return [pscustomobject]$result
    }

    if ($SkipVirtualMachines) {
        Add-Content -LiteralPath $logPath -Value ("[{0}] Launcher-side VM pre-check skipped to avoid WinRM. Endpoint guard will check locally under PsExec/SYSTEM." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    }

    $payloadCopied = $false
    $payloadCopyError = ''
    $maxPayloadCopyAttempts = 3
    for ($payloadCopyAttempt = 1; $payloadCopyAttempt -le $maxPayloadCopyAttempts; $payloadCopyAttempt++) {
        $result.RemotePayloadCopyAttempts = [string]$payloadCopyAttempt
        if ($payloadCopyAttempt -gt 1) {
            $delaySeconds = [math]::Min(30, 10 * ($payloadCopyAttempt - 1))
            Add-Content -LiteralPath $logPath -Value ("[{0}] Retrying SMB/admin-share preflight and payload copy. Attempt={1}/{2}; DelaySeconds={3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$payloadCopyAttempt,$maxPayloadCopyAttempts,$delaySeconds) -Encoding UTF8
            Start-Sleep -Seconds $delaySeconds
            $remoteAccess = Test-RemoteAdminAccess -ComputerName $Computer -LogPath $logPath
            Set-ResultRemoteAccessState -Result $result -State $remoteAccess
        }

        if (-not $remoteAccess.AdminShareReachable) {
            $payloadCopyError = ("{0}: Required administrative shares are not reachable. {1}" -f $remoteAccess.FailureType,$remoteAccess.Detail)
            Add-Content -LiteralPath $logPath -Value ("[{0}] WARN Payload copy preflight failed. Attempt={1}/{2}; Detail={3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$payloadCopyAttempt,$maxPayloadCopyAttempts,$payloadCopyError) -Encoding UTF8
            continue
        }

        try {
            Copy-RemotePayload -ComputerName $Computer -LogPath $logPath
            $payloadCopied = $true
            break
        }
        catch {
            $payloadCopyError = ("Remote payload copy failed after successful SMB preflight. {0}; Error={1}" -f $remoteAccess.Detail,$_.Exception.Message)
            Add-Content -LiteralPath $logPath -Value ("[{0}] WARN Payload copy failed. Attempt={1}/{2}; Detail={3}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$payloadCopyAttempt,$maxPayloadCopyAttempts,$payloadCopyError) -Encoding UTF8
        }
    }

    if (-not $payloadCopied) {
        if (-not $remoteAccess.AdminShareReachable) {
            $result.LauncherStatus = 'ADMIN_SHARE_UNREACHABLE'
            $result.Detail = $payloadCopyError
            Add-Content -LiteralPath $logPath -Value ("[{0}] Skipping PsExec after retries: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$result.Detail) -Encoding UTF8
            $result.RemoteLogsPath = Publish-LauncherEvidence -ComputerName $Computer -Cycle $CycleNumber -Bucket 'ADMIN_SHARE_UNREACHABLE' -WorkerLogPath $logPath -StdoutLogPath $stdoutPath -StderrLogPath $stderrPath
            return [pscustomobject]$result
        }

        $result.LauncherStatus = 'REMOTE_PAYLOAD_COPY_FAILED'
        $result.Detail = $payloadCopyError
        Add-Content -LiteralPath $logPath -Value ("[{0}] ERROR Payload copy failed after {1} attempt(s): {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$maxPayloadCopyAttempts,$result.Detail) -Encoding UTF8
        $result.RemoteLogsPath = Publish-LauncherEvidence -ComputerName $Computer -Cycle $CycleNumber -Bucket 'Errors' -WorkerLogPath $logPath -StdoutLogPath $stdoutPath -StderrLogPath $stderrPath
        return [pscustomobject]$result
    }
    $result.SetupCacheAction = Copy-SetupMediaToRemoteCache -ComputerName $Computer -LogPath $logPath

    $remotePowerShellArgs = @($RemoteScriptArgs | ForEach-Object { Convert-ToPsExecRemoteArgument -Value $_ })
    $psexecArgs = @(
        "\\$Computer",
        '-accepteula',
        '-nobanner',
        '-s',
        'powershell.exe',
        '-NoProfile',
        '-ExecutionPolicy','Bypass',
        '-File',(Convert-ToPsExecRemoteArgument -Value $RemoteScriptPath)
    ) + $remotePowerShellArgs

    Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec command: {1} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$ResolvedPsExecPath,($psexecArgs -join ' ')) -Encoding UTF8
    $process = Start-Process -FilePath $ResolvedPsExecPath -ArgumentList $psexecArgs -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    $psExecTimedOut = $false
    $waitStarted = Get-Date
    $lastWaitLog = $waitStarted
    Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec started. Waiting for remote script completion; heartbeat every 60 seconds." -f $waitStarted.ToString('yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    if ($PsExecTimeoutMinutes -gt 0) {
        $timeoutAt = $waitStarted.AddMinutes($PsExecTimeoutMinutes)
        while (-not $process.WaitForExit(1000)) {
            $now = Get-Date
            if ($now -ge $timeoutAt) {
                $psExecTimedOut = $true
                Add-Content -LiteralPath $logPath -Value ("[{0}] ERROR PsExec timed out after {1} minute(s). Killing local PsExec process." -f $now.ToString('yyyy-MM-dd HH:mm:ss'),$PsExecTimeoutMinutes) -Encoding UTF8
                try { $process.Kill() } catch { }
                try { [void]$process.WaitForExit(5000) } catch { }
                $result.LauncherStatus = 'PSEXEC_TIMEOUT'
                $result.Detail = "Timed out after $PsExecTimeoutMinutes minute(s)."
                break
            }
            if (($now - $lastWaitLog).TotalSeconds -ge 60) {
                $elapsedMinutes = [math]::Round(($now - $waitStarted).TotalMinutes, 1)
                Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec still running for {1} minute(s). Remote script may be copying setup media or running Windows setup." -f $now.ToString('yyyy-MM-dd HH:mm:ss'),$elapsedMinutes) -Encoding UTF8
                $lastWaitLog = $now
            }
        }
    }
    else {
        while (-not $process.WaitForExit(1000)) {
            $now = Get-Date
            if (($now - $lastWaitLog).TotalSeconds -ge 60) {
                $elapsedMinutes = [math]::Round(($now - $waitStarted).TotalMinutes, 1)
                Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec still running for {1} minute(s). Remote script may be copying setup media or running Windows setup." -f $now.ToString('yyyy-MM-dd HH:mm:ss'),$elapsedMinutes) -Encoding UTF8
                $lastWaitLog = $now
            }
        }
    }

    try { $process.Refresh() } catch { }
    if ($result.LauncherStatus -eq 'UNKNOWN') {
        if (-not $psExecTimedOut -and $process.HasExited) {
            $result.ExitCode = [string]$process.ExitCode
            $result.LauncherStatus = if ($process.ExitCode -eq 0) { 'SUCCESS' } else { "PSEXEC_EXIT_$($process.ExitCode)" }
        }
    }

    $stdoutContent = @()
    if (Test-Path -LiteralPath $stdoutPath) {
        Add-Content -LiteralPath $logPath -Value '----- PsExec STDOUT -----' -Encoding UTF8
        $stdoutContent = @(Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue)
        $stdoutContent | Add-Content -LiteralPath $logPath -Encoding UTF8
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
    }
    $stderrContent = @()
    if (Test-Path -LiteralPath $stderrPath) {
        Add-Content -LiteralPath $logPath -Value '----- PsExec STDERR -----' -Encoding UTF8
        $stderrContent = @(Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue)
        $stderrContent | Add-Content -LiteralPath $logPath -Encoding UTF8
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
    if ([string]::IsNullOrWhiteSpace($result.ExitCode)) {
        $nativeExitLine = ($stderrContent | Where-Object { $_ -match 'with error code\s+-?\d+' } | Select-Object -Last 1)
        if ($nativeExitLine -and $nativeExitLine -match 'with error code\s+(?<Code>-?\d+)') {
            $result.ExitCode = $Matches.Code
            Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec exit code recovered from native STDERR: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$result.ExitCode) -Encoding UTF8
        }
    }

    Start-Sleep -Seconds 3
    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $Computer -LocalPath $RemoteBaseDir
    $remoteLastRunPath = Join-Path $remoteBaseShare 'LastRun.json'
    if (Test-Path -LiteralPath $remoteLastRunPath -PathType Leaf) {
        $lastRun = Get-Content -LiteralPath $remoteLastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        Update-ResultFromLastRun -Result $result -LastRun $lastRun
    }

    $centralLogBucket = Get-CentralLogBucket -Result $result
    $result.RemoteLogsPath = Collect-RemoteEvidence -ComputerName $Computer -Cycle $CycleNumber -Bucket $centralLogBucket
}
catch {
    $result.LauncherStatus = 'ERROR'
    $result.Detail = $_.Exception.Message
    Add-Content -LiteralPath $logPath -Value ("[{0}] ERROR {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message) -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($result.RemoteLogsPath)) {
        $result.RemoteLogsPath = Publish-LauncherEvidence -ComputerName $Computer -Cycle $CycleNumber -Bucket 'Errors' -WorkerLogPath $logPath -StdoutLogPath $stdoutPath -StderrLogPath $stderrPath
    }
}

[pscustomobject]$result
