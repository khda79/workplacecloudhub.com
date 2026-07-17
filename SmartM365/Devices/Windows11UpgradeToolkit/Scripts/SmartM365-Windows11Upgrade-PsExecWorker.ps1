<#
.SYNOPSIS
    Internal worker used by SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1.

.DESCRIPTION
    Runs one remote computer operation in a background job. This file is operator-side only;
    the target device still receives only SmartM365-Invoke-Windows11UpgradeRepair.ps1.

.VERSION
0.1.27
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
    [string]$CentralLogCollectionMode = 'Standard',
    [int]$PsExecTimeoutMinutes = 360,
    [string]$GlobalWorkerLeasePath,
    [string]$GlobalWorkerLeaseMutexName
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($CentralLogCollectionMode)) {
    $CentralLogCollectionMode = 'Standard'
}
elseif ($CentralLogCollectionMode -match '^\d+$') {
    $PsExecTimeoutMinutes = [int]$CentralLogCollectionMode
    $CentralLogCollectionMode = 'Standard'
}
elseif ($CentralLogCollectionMode -notin @('Standard','Full')) {
    throw "Invalid CentralLogCollectionMode '$CentralLogCollectionMode'. Expected Standard or Full."
}
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

function ConvertTo-SafePathSegment {
    param([AllowNull()][string]$Value)

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return 'UnknownComputer' }
    $safe = [regex]::Replace($text, '[^A-Za-z0-9._-]', '_')
    $safe = $safe.Trim(' ','.','_','-')
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'UnknownComputer' }
    return $safe
}

function Get-ComputerPathKey {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    $shortName = (($ComputerName -split '\.')[0]).Trim()
    if ([string]::IsNullOrWhiteSpace($shortName)) { $shortName = $ComputerName }
    $safeShortName = ConvertTo-SafePathSegment -Value $shortName

    $normalizedFullName = $ComputerName.Trim().ToLowerInvariant()
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalizedFullName)
        $hash = (($sha256.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join '').Substring(0,8)
    }
    finally {
        $sha256.Dispose()
    }

    return ('{0}-{1}' -f $safeShortName,$hash)
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

function Protect-RemoteToolkitPathAcl {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [switch]$Directory
    )

    $icacls = Join-Path $env:SystemRoot 'System32\icacls.exe'
    if (-not (Test-Path -LiteralPath $icacls -PathType Leaf)) {
        Add-Content -LiteralPath $LogPath -Value ("[{0}] WARN ACL hardening skipped because icacls.exe was not found. Path={1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Path) -Encoding UTF8
        return
    }

    $grants = if ($Directory) {
        @('*S-1-5-18:(OI)(CI)F','*S-1-5-32-544:(OI)(CI)F','*S-1-5-32-545:(OI)(CI)RX')
    }
    else {
        @('*S-1-5-18:F','*S-1-5-32-544:F','*S-1-5-32-545:RX')
    }

    $output = & $icacls $Path '/inheritance:r' '/grant:r' $grants 2>&1
    $exitCode = [int]$LASTEXITCODE
    if ($exitCode -ne 0) {
        throw ("ACL hardening failed. Path={0}; ExitCode={1}; Output={2}" -f $Path,$exitCode,(($output | Out-String).Trim()))
    }

    Add-Content -LiteralPath $LogPath -Value ("[{0}] ACL hardening applied. Path={1}; Directory={2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Path,[bool]$Directory) -Encoding UTF8
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
    Protect-RemoteToolkitPathAcl -Path $remoteBaseShare -LogPath $LogPath -Directory
    $remoteScriptShare = Join-Path $remoteBaseShare (Split-Path -Leaf $RemoteScriptPath)
    Copy-Item -LiteralPath $LocalScriptPath -Destination $remoteScriptShare -Force -ErrorAction Stop
    Protect-RemoteToolkitPathAcl -Path $remoteScriptShare -LogPath $LogPath

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

function Get-PsExecCommunicationFailureDetail {
    param([AllowEmptyCollection()][string[]]$Lines)

    $evidence = @($Lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' | '
    if ([string]::IsNullOrWhiteSpace($evidence)) { return '' }
    if ($evidence -match 'Error communicating with PsExec service on (?<Computer>[^:]+):\s*(?<Error>[^|\r\n]+)') {
        return ("PsExec communication lost after launch. Computer={0}; Error={1}" -f $Matches.Computer.Trim(),$Matches.Error.Trim())
    }
    if ($evidence -match 'Descripteur non valide|The handle is invalid') {
        return 'PsExec communication lost after launch. Error=Invalid handle from PsExec service.'
    }
    if ($evidence -match 'Erreur r.seau inattendue|unexpected network error') {
        return 'PsExec communication lost after launch. Error=Unexpected network error.'
    }
    if ($evidence -match 'Le chemin r.seau n.a pas .t. trouv.|network path was not found') {
        return 'PsExec communication lost after launch. Error=Network path was not found.'
    }
    return ''
}

$script:CentralLogBuckets = @('Success','ADMIN_SHARE_UNREACHABLE','InsufficientDisk','Compatibility','SetupSourceLanguageUnavailable','SetupMigrationProfileFailure','SetupMigrationProfileRepaired','SetupMigrationPluginFailure','SetupCopyLeaseTimeout','SetupMediaCopyTimeout','SetupMediaCopyFailure','SetupMediaManifestFailure','SetupProcessTimeout','SetupProcessInterrupted','PsExecTimeout','PsExecCommunicationLost','RemoteLogCollectionFailed','Errors')

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
    if ($launcherStatus -eq 'PSEXEC_COMMUNICATION_LOST') {
        return 'PsExecCommunicationLost'
    }
    if ($launcherStatus -eq 'REMOTE_LOG_COLLECTION_FAILED' -or $launcherStatus -eq 'REMOTE_RESULT_STALE' -or $launcherStatus -eq 'STALE_LASTRUN_IGNORED') {
        return 'RemoteLogCollectionFailed'
    }

    foreach ($statusName in @('RemoteStatus','LauncherStatus')) {
        $statusValue = Get-ResultValue -Result $Result -Name $statusName
        if ($statusValue -like 'INSUFFICIENT_DISK*') {
            return 'InsufficientDisk'
        }
        if ($statusValue -eq 'WINDOWS11_COMPAT_BLOCKER') {
            return 'Compatibility'
        }
        if ($statusValue -eq 'PSEXEC_TIMEOUT') {
            return 'PsExecTimeout'
        }
        if ($statusValue -eq 'PSEXEC_COMMUNICATION_LOST') {
            return 'PsExecCommunicationLost'
        }
        if ($statusValue -eq 'REMOTE_LOG_COLLECTION_FAILED' -or $statusValue -eq 'REMOTE_RESULT_STALE' -or $statusValue -eq 'STALE_LASTRUN_IGNORED') {
            return 'RemoteLogCollectionFailed'
        }
        if ($statusValue -eq 'SETUP_PROCESS_TIMEOUT') {
            return 'SetupProcessTimeout'
        }
        if ($statusValue -eq 'SETUP_MEDIA_COPY_TIMEOUT') {
            return 'SetupMediaCopyTimeout'
        }
        if ($statusValue -eq 'SETUP_MEDIA_COPY_FAILED') {
            return 'SetupMediaCopyFailure'
        }
        if ($statusValue -eq 'SETUP_MEDIA_MANIFEST_VALIDATION_FAILED') {
            return 'SetupMediaManifestFailure'
        }
        if ($statusValue -eq 'SETUP_PROCESS_MONITOR_INTERRUPTED') {
            return 'SetupProcessInterrupted'
        }
        if ($statusValue -eq 'SETUP_SOURCE_LANGUAGE_UNAVAILABLE') {
            return 'SetupSourceLanguageUnavailable'
        }
        if ($statusValue -eq 'SETUP_PROFILE_DUPLICATE_REPAIRED_REBOOT_REQUIRED') {
            return 'SetupMigrationProfileRepaired'
        }
        if ($statusValue -in @('SETUP_MIGRATION_PROFILE_FAILURE','SETUP_MIGRATION_PROFILE_REPAIR_FAILED')) {
            return 'SetupMigrationProfileFailure'
        }
        if ($statusValue -eq 'SETUP_MIGRATION_PLUGIN_FAILURE') {
            return 'SetupMigrationPluginFailure'
        }
        if ($statusValue -in @('SETUP_SUBNET_COPY_LEASE_TIMEOUT','SETUP_SOURCE_COPY_LEASE_TIMEOUT')) {
            return 'SetupCopyLeaseTimeout'
        }
    }

    foreach ($detailName in @('Detail','JobErrorMessage')) {
        $detailValue = Get-ResultValue -Result $Result -Name $detailName
        if ($detailValue -match '0xC1900200|0xC1900202|0xC1900208|Compatibility failure|Compatibility blocker') {
            return 'Compatibility'
        }
        if ($detailValue -match 'No setup source subfolder under .+ contains language .+ in sources\\lang\\.ini') {
            return 'SetupSourceLanguageUnavailable'
        }
        if ($detailValue -match '0x8007001F|Duplicate profile detected|Duplicate setup migration profile|SetupProfileDuplicate') {
            return 'SetupMigrationProfileFailure'
        }
        if ($detailValue -match '0x8007007F|Setup migration plugin failure|SetupMigrationPluginFailure|CscMig\\.dll|WSManMigrationPlugin\\.dll|RasMigPlugin\\.dll|LoadDllServer|LoadLibraryExW') {
            return 'SetupMigrationPluginFailure'
        }
        if ($detailValue -match 'setup\.exe timed out after') {
            return 'SetupProcessTimeout'
        }
        if ($detailValue -match 'Robocopy setup media copy timed out after') {
            return 'SetupMediaCopyTimeout'
        }
        if ($detailValue -match 'Robocopy setup media copy failed with exit code') {
            return 'SetupMediaCopyFailure'
        }
        if ($detailValue -match 'Setup media integrity manifest validation failed|Setup media integrity manifest contains|Setup media integrity check failed') {
            return 'SetupMediaManifestFailure'
        }
        if ($detailValue -match 'Setup monitoring was interrupted before setup\.exe exit was observed') {
            return 'SetupProcessInterrupted'
        }
        if ($detailValue -match 'Timed out waiting for setup (subnet|source) copy lease') {
            return 'SetupCopyLeaseTimeout'
        }
        if ($detailValue -match 'Error communicating with PsExec service|Descripteur non valide|Erreur r.seau inattendue|unexpected network error') {
            return 'PsExecCommunicationLost'
        }
        if ($detailValue -match 'Central log collection failed|Le chemin r.seau n.a pas .t. trouv.|network path was not found') {
            return 'RemoteLogCollectionFailed'
        }
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
        [Parameter(Mandatory = $true)][ValidateSet('Success','ADMIN_SHARE_UNREACHABLE','InsufficientDisk','Compatibility','SetupSourceLanguageUnavailable','SetupMigrationProfileFailure','SetupMigrationProfileRepaired','SetupMigrationPluginFailure','SetupCopyLeaseTimeout','SetupMediaCopyTimeout','SetupMediaCopyFailure','SetupMediaManifestFailure','SetupProcessTimeout','SetupProcessInterrupted','PsExecTimeout','PsExecCommunicationLost','RemoteLogCollectionFailed','Errors')][string]$Bucket
    )

    $computerPathKey = Get-ComputerPathKey -ComputerName $ComputerName

    if ($KeepCentralLogHistory) {
        return (Join-Path (Join-Path (Join-Path $CentralLogRoot $Bucket) $computerPathKey) ("Cycle{0}_{1}" -f $Cycle,(Get-Date -Format 'yyyyMMdd-HHmmss')))
    }

    foreach ($otherBucket in $script:CentralLogBuckets) {
        $otherLatest = Join-Path (Join-Path (Join-Path $CentralLogRoot $otherBucket) $computerPathKey) 'Latest'
        if (Test-Path -LiteralPath $otherLatest) {
            Remove-Item -LiteralPath $otherLatest -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    return (Join-Path (Join-Path (Join-Path $CentralLogRoot $Bucket) $computerPathKey) 'Latest')
}

function Publish-LauncherEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Cycle,
        [Parameter(Mandatory = $true)][ValidateSet('Success','ADMIN_SHARE_UNREACHABLE','InsufficientDisk','Compatibility','SetupSourceLanguageUnavailable','SetupMigrationProfileFailure','SetupMigrationProfileRepaired','SetupMigrationPluginFailure','SetupCopyLeaseTimeout','SetupMediaCopyTimeout','SetupMediaCopyFailure','SetupMediaManifestFailure','SetupProcessTimeout','SetupProcessInterrupted','PsExecTimeout','PsExecCommunicationLost','RemoteLogCollectionFailed','Errors')][string]$Bucket,
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
        Copy-Item -LiteralPath $WorkerLogPath -Destination (Join-Path $target 'Worker.log') -Force -ErrorAction SilentlyContinue
    }
    foreach ($extraLog in @($StdoutLogPath,$StderrLogPath)) {
        if (-not [string]::IsNullOrWhiteSpace($extraLog) -and (Test-Path -LiteralPath $extraLog -PathType Leaf)) {
            $extraName = Split-Path -Leaf $extraLog
            if ($extraName -like '*.stdout.txt') { $extraName = 'PsExec.stdout.txt' }
            elseif ($extraName -like '*.stderr.txt') { $extraName = 'PsExec.stderr.txt' }
            else { $extraName = ConvertTo-SafePathSegment -Value $extraName }
            Copy-Item -LiteralPath $extraLog -Destination (Join-Path $target $extraName) -Force -ErrorAction SilentlyContinue
        }
    }
    return $target
}

$script:CentralLogMaxStandardFileBytes = 5MB

function Add-CentralLogCollectionNote {
    param(
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][string]$Message
    )

    try {
        New-Directory -Path $TargetPath
        Add-Content -LiteralPath (Join-Path $TargetPath 'CentralLogCollection.skipped.txt') -Value $Message -Encoding UTF8
    }
    catch { }
}

function Copy-CentralEvidenceFile {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$TargetRoot,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [long]$MaxBytes = $script:CentralLogMaxStandardFileBytes
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { return }

    try {
        $item = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
        if ($CentralLogCollectionMode -ne 'Full' -and $item.Length -gt $MaxBytes) {
            Add-CentralLogCollectionNote -TargetPath $TargetRoot -Message ("Skipped large file: {0}; SizeMB={1:N2}; LimitMB={2:N2}; Source={3}" -f $RelativePath,($item.Length / 1MB),($MaxBytes / 1MB),$SourcePath)
            return
        }

        $destination = Join-Path $TargetRoot $RelativePath
        $destinationParent = Split-Path -Parent $destination
        if (-not [string]::IsNullOrWhiteSpace($destinationParent)) { New-Directory -Path $destinationParent }
        Copy-Item -LiteralPath $SourcePath -Destination $destination -Force -ErrorAction SilentlyContinue
    }
    catch {
        Add-CentralLogCollectionNote -TargetPath $TargetRoot -Message ("Failed to copy file: {0}; Error={1}" -f $SourcePath,$_.Exception.Message)
    }
}

function Copy-CentralEvidenceDirectorySmallFiles {
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string]$TargetRoot,
        [Parameter(Mandatory = $true)][string]$RelativeRoot,
        [long]$MaxBytes = $script:CentralLogMaxStandardFileBytes
    )

    if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) { return }

    $root = [System.IO.Path]::GetFullPath($SourceRoot).TrimEnd('\')
    foreach ($file in @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse -ErrorAction SilentlyContinue)) {
        $fileFullName = [System.IO.Path]::GetFullPath($file.FullName)
        $relativeChild = $fileFullName.Substring($root.Length).TrimStart('\')
        $relative = Join-Path $RelativeRoot $relativeChild
        Copy-CentralEvidenceFile -SourcePath $file.FullName -TargetRoot $TargetRoot -RelativePath $relative -MaxBytes $MaxBytes
    }
}

function Collect-StandardRemoteEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$RemoteBaseShare,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )

    Add-CentralLogCollectionNote -TargetPath $TargetPath -Message ("CentralLogCollectionMode=Standard; copied selected files only. Full logs remain on target under {0}." -f $RemoteBaseShare)

    Copy-CentralEvidenceFile -SourcePath (Join-Path $RemoteBaseShare 'LastRun.json') -TargetRoot $TargetPath -RelativePath 'LastRun.json' -MaxBytes 1MB

    $logsRoot = Join-Path $RemoteBaseShare 'Logs'
    if (Test-Path -LiteralPath $logsRoot -PathType Container) {
        foreach ($file in @(Get-ChildItem -LiteralPath $logsRoot -File -ErrorAction SilentlyContinue)) {
            Copy-CentralEvidenceFile -SourcePath $file.FullName -TargetRoot $TargetPath -RelativePath (Join-Path 'Logs' $file.Name) -MaxBytes 10MB
        }

        $setupLogRoot = Join-Path $logsRoot 'SetupUpgrade'
        if (Test-Path -LiteralPath $setupLogRoot -PathType Container) {
            foreach ($file in @(Get-ChildItem -LiteralPath $setupLogRoot -File -ErrorAction SilentlyContinue)) {
                $copy = ($file.Name -like 'Robocopy_*.log' -or $file.Name -like 'SetupDiag*.log' -or $file.Extension -in @('.json','.xml','.txt','.log'))
                if ($copy) {
                    Copy-CentralEvidenceFile -SourcePath $file.FullName -TargetRoot $TargetPath -RelativePath (Join-Path 'Logs\SetupUpgrade' $file.Name) -MaxBytes $script:CentralLogMaxStandardFileBytes
                }
            }
        }
    }

    Copy-CentralEvidenceDirectorySmallFiles -SourceRoot (Join-Path $RemoteBaseShare 'Output') -TargetRoot $TargetPath -RelativeRoot 'Output' -MaxBytes 2MB
}
function Get-RemoteEndpointLatestLogPath {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$RemoteBaseDir
    )

    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $ComputerName -LocalPath $RemoteBaseDir
    $remoteLogsShare = Join-Path $remoteBaseShare 'Logs'
    if (-not (Test-Path -LiteralPath $remoteLogsShare -PathType Container)) { return '' }

    $latest = Get-ChildItem -LiteralPath $remoteLogsShare -Filter 'SmartM365-Invoke-Windows11UpgradeRepair_*.log' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($latest) { return [string]$latest.FullName }
    return ''
}

function Get-RemoteEndpointSetupMonitorRecovery {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$RemoteBaseDir,
        [AllowNull()][string]$PsExecExitCode
    )

    $latestLog = Get-RemoteEndpointLatestLogPath -ComputerName $ComputerName -RemoteBaseDir $RemoteBaseDir
    if ([string]::IsNullOrWhiteSpace($latestLog)) { return $null }

    $tail = @(Get-Content -LiteralPath $latestLog -Tail 120 -ErrorAction SilentlyContinue)
    if ($tail.Count -eq 0) { return $null }

    $tailText = ($tail -join "`n")
    $hasFinal = ($tailText -match 'Final Status=')
    $setupStarted = ($tailText -match 'Starting setup upgrade:|setup\.exe started\.')
    $setupStillRunning = ($tailText -match 'setup\.exe still running\.')
    $setupExited = ($tailText -match 'setup\.exe exited\.')

    if (-not $hasFinal -and $setupStarted -and ($setupStillRunning -or -not $setupExited)) {
        return [pscustomobject]@{
            Status = 'SETUP_PROCESS_MONITOR_INTERRUPTED'
            NextAction = 'CHECK_SETUP_PROCESS_OR_OS_STATUS'
            ExitCode = '3'
            Detail = ("PsExec exited with code {0} before the endpoint wrote Final Status, but the remote endpoint log shows setup.exe was started and was still being monitored. RemoteLog={1}" -f $PsExecExitCode,$latestLog)
            RemoteLogPath = $latestLog
        }
    }

    return $null
}
function Test-LastRunFreshness {
    param(
        [Parameter(Mandatory = $true)]$LastRunItem,
        [Parameter(Mandatory = $true)]$LastRun,
        [Parameter(Mandatory = $true)][datetime]$WorkerStartedUtc
    )

    $minimumUtc = $WorkerStartedUtc.AddMinutes(-2)
    $lastWriteUtc = $LastRunItem.LastWriteTimeUtc
    $timeValues = New-Object System.Collections.Generic.List[datetime]
    foreach ($propertyName in @('EndTimeUtc','StartTimeUtc')) {
        if ($LastRun.PSObject.Properties[$propertyName] -and -not [string]::IsNullOrWhiteSpace([string]$LastRun.$propertyName)) {
            try {
                $parsed = [datetime]::Parse([string]$LastRun.$propertyName, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
                [void]$timeValues.Add($parsed.ToUniversalTime())
            }
            catch { }
        }
    }

    $latestContentUtc = $null
    if ($timeValues.Count -gt 0) {
        $latestContentUtc = @($timeValues | Sort-Object -Descending | Select-Object -First 1)[0]
    }

    $freshByWriteTime = ($lastWriteUtc -ge $minimumUtc)
    $freshByContentTime = ($null -ne $latestContentUtc -and $latestContentUtc -ge $minimumUtc)
    if ($freshByWriteTime -or $freshByContentTime) {
        return [pscustomobject]@{
            IsFresh = $true
            Detail = ("LastRun accepted. LastWriteTimeUtc={0:o}; LatestContentUtc={1}; WorkerStartedUtc={2:o}." -f $lastWriteUtc, $(if ($latestContentUtc) { $latestContentUtc.ToString('o') } else { '' }), $WorkerStartedUtc)
        }
    }

    return [pscustomobject]@{
        IsFresh = $false
        Detail = ("Ignored stale LastRun.json. LastWriteTimeUtc={0:o}; LatestContentUtc={1}; WorkerStartedUtc={2:o}; MinimumAcceptedUtc={3:o}." -f $lastWriteUtc, $(if ($latestContentUtc) { $latestContentUtc.ToString('o') } else { '' }), $WorkerStartedUtc, $minimumUtc)
    }
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
        if ($LastRun.PSObject.Properties['Detail']) {
            $Result.Detail = [string]$LastRun.Detail
        }
        else {
            $Result.Detail = ''
        }
    }
    if ($LastRun.PSObject.Properties['SetupCacheAction']) {
        $Result.SetupCacheAction = [string]$LastRun.SetupCacheAction
    }
    foreach ($propertyName in @('SetupDynamicUpdate','SelectedSetupSourcePath','SetupSourceSelectionDetail','DiskCleanupAction','DiskCleanupFreedGB','AdvancedDiskCleanupAction','AdvancedDiskCleanupFreedGB','DismCleanupAction','DismCleanupFreedGB','SetupCompletionRebootAction','SetupCompletionRebootDetail','SetupCompletionRebootUserCount','SetupCompletionRebootUsers','SetupProfileRepairAction','SetupProfileRepairDetail','SetupProfileRepairBlockingSid','SetupProfileRepairKeptSid','SetupProfileRepairProfilePath','SetupProfileRepairBackupPath','ControlledRebootAction','ControlledRebootDetail','ControlledRebootUserCount','ControlledRebootUsers','RetryAfterRebootAction','RetryAfterRebootDetail','RetryAfterRebootAttempt','RetryAfterRebootMaxAttempts','RetryAfterRebootTaskName')) {
        if ($LastRun.PSObject.Properties[$propertyName]) {
            $Result[$propertyName] = [string]$LastRun.$propertyName
        }
    }
}

function Collect-RemoteEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Cycle,
        [Parameter(Mandatory = $true)][ValidateSet('Success','ADMIN_SHARE_UNREACHABLE','InsufficientDisk','Compatibility','SetupSourceLanguageUnavailable','SetupMigrationProfileFailure','SetupMigrationProfileRepaired','SetupMigrationPluginFailure','SetupCopyLeaseTimeout','SetupMediaCopyTimeout','SetupMediaCopyFailure','SetupMediaManifestFailure','SetupProcessTimeout','SetupProcessInterrupted','PsExecTimeout','PsExecCommunicationLost','RemoteLogCollectionFailed','Errors')][string]$Bucket
    )

    if ($NoCentralLogCollection) { return '' }

    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $ComputerName -LocalPath $RemoteBaseDir
    if (-not (Test-Path -LiteralPath $remoteBaseShare -PathType Container)) { return '' }

    $target = New-CentralLogTarget -ComputerName $ComputerName -Cycle $Cycle -Bucket $Bucket
    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Directory -Path $target

    if ($CentralLogCollectionMode -eq 'Full') {
        foreach ($child in @('Logs','Output','LastRun.json')) {
            $source = Join-Path $remoteBaseShare $child
            if (Test-Path -LiteralPath $source) {
                Copy-Item -LiteralPath $source -Destination $target -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
    else {
        Collect-StandardRemoteEvidence -RemoteBaseShare $remoteBaseShare -TargetPath $target
    }
    return $target
}
New-Directory -Path $LogRoot
$workerStartedUtc = (Get-Date).ToUniversalTime()
$computerPathKey = Get-ComputerPathKey -ComputerName $Computer
$logPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}.log" -f $computerPathKey,$CycleNumber,(Get-Date -Format 'yyyyMMdd-HHmmss'))
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
    SetupProfileRepairAction = ''
    SetupProfileRepairDetail = ''
    SetupProfileRepairBlockingSid = ''
    SetupProfileRepairKeptSid = ''
    SetupProfileRepairProfilePath = ''
    SetupProfileRepairBackupPath = ''
    ControlledRebootAction = ''
    ControlledRebootDetail = ''
    ControlledRebootUserCount = ''
    ControlledRebootUsers = ''
    RetryAfterRebootAction = ''
    RetryAfterRebootDetail = ''
    RetryAfterRebootAttempt = ''
    RetryAfterRebootMaxAttempts = ''
    RetryAfterRebootTaskName = ''
    RemoteLogsPath = ''
    PsExecLogPath = $logPath
    JobErrorMessage = ''
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
            if ($null -eq $process.ExitCode) {
                $result.LauncherStatus = 'PSEXEC_EXIT_UNKNOWN'
                $result.Detail = 'PsExec process exited but no native exit code was available.'
            }
            else {
                $result.ExitCode = [string]$process.ExitCode
                $result.LauncherStatus = if ($process.ExitCode -eq 0) { 'SUCCESS' } else { "PSEXEC_EXIT_$($process.ExitCode)" }
            }
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
            $recoveredPsExecExitCode = 0
            if ([int]::TryParse([string]$result.ExitCode, [ref]$recoveredPsExecExitCode) -and $result.LauncherStatus -in @('UNKNOWN','PSEXEC_EXIT_UNKNOWN','PSEXEC_EXIT_')) {
                $result.LauncherStatus = if ($recoveredPsExecExitCode -eq 0) { 'SUCCESS' } else { "PSEXEC_EXIT_$recoveredPsExecExitCode" }
                if ($result.LauncherStatus -ne 'PSEXEC_EXIT_UNKNOWN' -and $result.Detail -eq 'PsExec process exited but no native exit code was available.') {
                    $result.Detail = ''
                }
            }
        }
    }

    $psExecCommunicationFailureDetail = Get-PsExecCommunicationFailureDetail -Lines $stderrContent
    if (-not [string]::IsNullOrWhiteSpace($psExecCommunicationFailureDetail) -and [string]::IsNullOrWhiteSpace($result.RemoteStatus)) {
        $result.LauncherStatus = 'PSEXEC_COMMUNICATION_LOST'
        $result.Detail = $psExecCommunicationFailureDetail
        Add-Content -LiteralPath $logPath -Value ("[{0}] WARN {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$psExecCommunicationFailureDetail) -Encoding UTF8
    }

    Start-Sleep -Seconds 3
    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $Computer -LocalPath $RemoteBaseDir
    $remoteLastRunPath = Join-Path $remoteBaseShare 'LastRun.json'
    if (Test-Path -LiteralPath $remoteLastRunPath -PathType Leaf) {
        $lastRunItem = Get-Item -LiteralPath $remoteLastRunPath -ErrorAction Stop
        $lastRun = Get-Content -LiteralPath $remoteLastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $lastRunFreshness = Test-LastRunFreshness -LastRunItem $lastRunItem -LastRun $lastRun -WorkerStartedUtc $workerStartedUtc
        if ($lastRunFreshness.IsFresh) {
            Update-ResultFromLastRun -Result $result -LastRun $lastRun
        }
        else {
            $previousStatus = $result.LauncherStatus
            $result.LauncherStatus = 'REMOTE_RESULT_STALE'
            $result.RemoteNextAction = 'CHECK_REMOTE_EXECUTION_AND_LASTRUN'
            $result.Detail = ("{0} PreviousLauncherStatus={1}" -f $lastRunFreshness.Detail,$previousStatus)
            Add-Content -LiteralPath $logPath -Value ("[{0}] WARN {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$result.Detail) -Encoding UTF8
        }
    }
    elseif ($result.LauncherStatus -like 'PSEXEC_EXIT_*' -or $result.LauncherStatus -eq 'PSEXEC_EXIT_UNKNOWN' -or $result.LauncherStatus -eq 'PSEXEC_COMMUNICATION_LOST') {
        $setupRecovery = Get-RemoteEndpointSetupMonitorRecovery -ComputerName $Computer -RemoteBaseDir $RemoteBaseDir -PsExecExitCode ([string]$result.ExitCode)
        if ($setupRecovery) {
            $result.LauncherStatus = [string]$setupRecovery.Status
            $result.RemoteNextAction = [string]$setupRecovery.NextAction
            $result.ExitCode = [string]$setupRecovery.ExitCode
            $result.Detail = [string]$setupRecovery.Detail
            Add-Content -LiteralPath $logPath -Value ("[{0}] WARN Reclassified PsExec exit as {1}: {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$result.LauncherStatus,$result.Detail) -Encoding UTF8
        }
    }

    $centralLogBucket = Get-CentralLogBucket -Result $result
    try {
        $result.RemoteLogsPath = Collect-RemoteEvidence -ComputerName $Computer -Cycle $CycleNumber -Bucket $centralLogBucket
    }
    catch {
        $centralLogError = "Central log collection failed: $($_.Exception.Message)"
        $result.JobErrorMessage = $centralLogError
        Add-Content -LiteralPath $logPath -Value ("[{0}] WARN {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$centralLogError) -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($result.RemoteStatus)) {
            if ($result.LauncherStatus -ne 'PSEXEC_COMMUNICATION_LOST') {
                $result.LauncherStatus = 'REMOTE_LOG_COLLECTION_FAILED'
                $result.Detail = $centralLogError
            }
            $centralLogBucket = Get-CentralLogBucket -Result $result
        }
    }
    if ([string]::IsNullOrWhiteSpace($result.RemoteLogsPath) -and [string]::IsNullOrWhiteSpace($result.RemoteStatus) -and $result.LauncherStatus -ne 'PSEXEC_COMMUNICATION_LOST') {
        if ($result.LauncherStatus -eq 'SUCCESS' -or $result.LauncherStatus -eq 'PSEXEC_EXIT_UNKNOWN' -or $result.LauncherStatus -like 'PSEXEC_EXIT_*') {
            $result.LauncherStatus = 'REMOTE_LOG_COLLECTION_FAILED'
            if ([string]::IsNullOrWhiteSpace($result.Detail) -or $result.Detail -eq 'PsExec process exited but no native exit code was available.') {
                $result.Detail = 'Remote script status could not be read because LastRun.json/log collection was unavailable after PsExec completed.'
            }
        }
    }
    if ([string]::IsNullOrWhiteSpace($result.RemoteLogsPath) -and $result.LauncherStatus -in @('PSEXEC_COMMUNICATION_LOST','REMOTE_LOG_COLLECTION_FAILED')) {
        $result.RemoteLogsPath = Publish-LauncherEvidence -ComputerName $Computer -Cycle $CycleNumber -Bucket (Get-CentralLogBucket -Result $result) -WorkerLogPath $logPath -StdoutLogPath $stdoutPath -StderrLogPath $stderrPath
    }
}
catch {
    if (-not [string]::IsNullOrWhiteSpace($result.RemoteStatus)) {
        $result.LauncherStatus = $result.RemoteStatus
        if ([string]::IsNullOrWhiteSpace($result.JobErrorMessage)) {
            $result.JobErrorMessage = $_.Exception.Message
        }
    }
    else {
        $result.LauncherStatus = 'ERROR'
        $result.Detail = $_.Exception.Message
    }
    Add-Content -LiteralPath $logPath -Value ("[{0}] ERROR {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message) -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($result.RemoteLogsPath)) {
        $result.RemoteLogsPath = Publish-LauncherEvidence -ComputerName $Computer -Cycle $CycleNumber -Bucket 'Errors' -WorkerLogPath $logPath -StdoutLogPath $stdoutPath -StderrLogPath $stderrPath
    }
}

[pscustomobject]$result

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDNEyJshQ7y8HQ/
# psXuyPba/XbveSTSzcVbLtgQUop7M6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAjNnFx2ZAWKBiHMNaZjoyRMgn7u+IDt3/xuNOUVzZHoMA0GCSqG
# SIb3DQEBAQUABIIBgDi3+BojRaxAHDCmNwFE0pJxf86avMWfEHBY1ZpZHwavJb48
# Qr7v8C/lenSXfXuhGG2x77KXH9f4lX/204N0DfIlWNKKhHev0pgB3LTwzeuZbUdN
# /jaS5tfXK4URXYttOXKxEI2Ui8jsVlmr30DP0tLopKxv0RNeE28TrNBc8QwymZ5i
# RpU+eb8dBJwCuFxDrUfgBpfYWRuQRkNi8zSLuSrKKn5zySzNxptwvqWo/jccyfDN
# EqF74ji9K3r5TPyoiVQuIiQrLD2eIidkNA0Ew3hwaE8qxJKiDGqC3COpdIjiRSRZ
# H4wA42xeTH2SvUJG9WbK/j1EfuS4fXWoSR5/DYVv7iy+Aqo4P/j+LV7b93THz0K1
# uMhwdDuf7lux5LmzJeqSLnL7V6PJh4NPRcVYgHFbSgpZMGnf/y9sdn4bnWahC/Lu
# ONRJjGykJKUv1kykJA2AOCSIBqBi8ovM1R+fIH2UdThtq+u1AMycM78s1GYBx9Qc
# XQs1mRJWN8OohnbBeaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTQxODA5
# NDhaMC8GCSqGSIb3DQEJBDEiBCBNwVbExhVFOiqDc9OSgabNtMAmU87dIr8qicV0
# hCuJmzANBgkqhkiG9w0BAQEFAASCAgB+1UjRpOXNkZMyHTH2Tk/dniT7nRfWKb4S
# CG0TslCS+mAe7nIqpJncjUQ28MM6xdf8OZybziDis2HPlGEOY4LCIWUhOx03y75R
# 8n8IOo7/eyAu4vDLgl+YMgwFllFAa4pzCzAEAfuOlaDHxueZk2/Bsl07kPHw21H0
# PPyh21XVWt0qiz6acfWOIz9HF1KInH/6OTpyyxxC/6tIczyN9AHgeceCwgtFW6F4
# dIPsPeL9U8JjPqga3XnKYDtlHt3jldx0ierHwDqH496RE1O8iSVA50chtrsQ7zg8
# otHUdqm4KdQKhK/t8r/mmSmLO9NK8iDe3rHR5oJ2WSjtaTY2tBoZ2v/seRJkro+P
# ZhiXZWIN0qZewX/EhtcUvC9LySaHsM1BsQ57tzlBm8V5iZOU7qVXxDVyCxAZtQ/r
# Pte5js7TG3K9tl5A1pnJvTN7uS8I1TS5aiAYMXjIUPFnEvqqSz9lrZsU8CF/Bf57
# CjbrY+7WPx3UfLLeBBOcKLGSR9A/cjkkJJQNR9FEc9v141jW27Ubf9ByoPI9YDbn
# dUCcgXGRwfyKdWyk42jcyNl3MwM2br60cR1aQnBqjoFeJWM5rt5kpTNv2i7FrOxh
# GKiRxz94RvIwfmMzHqoBBz8sdeF9IxU7upgOWGAWYHGh+7AM/bpep7rTNb5RDod7
# 5DeAwl7FeQ==
# SIG # End signature block
