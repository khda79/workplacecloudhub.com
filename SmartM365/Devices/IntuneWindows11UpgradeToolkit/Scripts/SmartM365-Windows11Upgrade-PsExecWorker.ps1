<#
.SYNOPSIS
    Internal worker used by SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1.

.DESCRIPTION
    Runs one remote computer operation in a background job. This file is operator-side only;
    the target device still receives only SmartM365-Invoke-Windows11UpgradeRepair.ps1.
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Computer,
    [Parameter(Mandatory = $true)][int]$CycleNumber,
    [Parameter(Mandatory = $true)][string[]]$RemoteScriptArgs,
    [string]$ResolvedPsExecPath,
    [Parameter(Mandatory = $true)][string]$LocalScriptPath,
    [Parameter(Mandatory = $true)][string]$RemoteBaseDir,
    [Parameter(Mandatory = $true)][string]$RemoteScriptPath,
    [Parameter(Mandatory = $true)][string]$RemoteSetupCacheRoot,
    [Parameter(Mandatory = $true)][string]$LogRoot,
    [Parameter(Mandatory = $true)][string]$CentralLogRoot,
    [string]$SetupSourcePath,
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

function Get-SetupMediaLanguages {
    param([Parameter(Mandatory = $true)][string]$MediaRoot)

    $langIni = Join-Path $MediaRoot 'sources\lang.ini'
    if (-not (Test-Path -LiteralPath $langIni -PathType Leaf)) { return @() }

    $languages = New-Object System.Collections.ArrayList
    $inAvailableSection = $false
    foreach ($rawLine in @(Get-Content -LiteralPath $langIni -ErrorAction Stop)) {
        $line = ([string]$rawLine).Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith(';')) { continue }
        if ($line -match '^\[(.+)\]$') {
            $inAvailableSection = ($Matches[1] -ieq 'Available UI Languages')
            continue
        }
        if ($inAvailableSection -and $line -match '^([^=]+)=') {
            $language = $Matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($language) -and -not $languages.Contains($language)) {
                [void]$languages.Add($language)
            }
        }
    }

    return @($languages.ToArray())
}

function Get-RemoteSystemLanguage {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
        if ($null -ne $os.OSLanguage) {
            return ([System.Globalization.CultureInfo]::GetCultureInfo([int]$os.OSLanguage)).Name
        }
    }
    catch { }

    return ''
}

function Resolve-ExpectedSetupLanguage {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    if ([string]::IsNullOrWhiteSpace($SetupLanguage)) { return '' }
    $expectedLanguage = $SetupLanguage.Trim()
    if ($expectedLanguage -in @('Any','None','Disabled')) { return '' }
    if ($expectedLanguage -in @('Auto','MatchSystem','System')) {
        $expectedLanguage = Get-RemoteSystemLanguage -ComputerName $ComputerName
        if ([string]::IsNullOrWhiteSpace($expectedLanguage)) {
            throw "Unable to detect remote system language for $ComputerName."
        }
        return $expectedLanguage
    }

    try {
        return ([System.Globalization.CultureInfo]::GetCultureInfo($expectedLanguage)).Name
    }
    catch {
        throw "Invalid SetupLanguage value '$SetupLanguage'. Use MatchSystem, Any, or a culture tag such as fr-FR or en-US."
    }
}

function Resolve-SetupSourceMediaPath {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [string]$ExpectedLanguage
    )

    if (Test-Path -LiteralPath (Join-Path $SourcePath 'setup.exe') -PathType Leaf) {
        return $SourcePath
    }

    if ([string]::IsNullOrWhiteSpace($ExpectedLanguage)) {
        throw "SetupSourcePath does not contain setup.exe and language matching is disabled, so no language subfolder can be selected: $SourcePath"
    }

    foreach ($child in @(Get-ChildItem -LiteralPath $SourcePath -Directory -ErrorAction SilentlyContinue)) {
        $languages = @(Get-SetupMediaLanguages -MediaRoot $child.FullName)
        if (@($languages | Where-Object { $_ -ieq $ExpectedLanguage } | Select-Object -First 1).Count -gt 0) {
            return $child.FullName
        }
    }

    throw ("No setup source subfolder under '{0}' contains language {1} in sources\lang.ini." -f $SourcePath,$ExpectedLanguage)
}

function Test-SetupSourceLanguage {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$ComputerName
    )

    $expectedLanguage = Resolve-ExpectedSetupLanguage -ComputerName $ComputerName
    if ([string]::IsNullOrWhiteSpace($expectedLanguage)) { return }

    $languages = @(Get-SetupMediaLanguages -MediaRoot $SourcePath)
    if ($languages.Count -eq 0) {
        throw "Setup media language could not be detected from sources\lang.ini. Expected=$expectedLanguage; Source=$SourcePath"
    }
    if (-not @($languages | Where-Object { $_ -ieq $expectedLanguage } | Select-Object -First 1)) {
        throw ("Setup media language mismatch. Expected={0}; Available={1}; Source={2}" -f $expectedLanguage,($languages -join ','),$SourcePath)
    }
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
    if ($SkipSetupMediaPreCopy) { return 'Skipped' }
    if ($SetupExecutionMode -eq 'Share') { return 'ShareMode' }
    if ([string]::IsNullOrWhiteSpace($SetupSourcePath)) { return 'NoSource' }
    if (-not (Test-Path -LiteralPath $SetupSourcePath -PathType Container)) {
        throw "SetupSourcePath not reachable from operator workstation: $SetupSourcePath"
    }

    $expectedLanguage = Resolve-ExpectedSetupLanguage -ComputerName $ComputerName
    $resolvedSetupSourcePath = Resolve-SetupSourceMediaPath -SourcePath $SetupSourcePath -ExpectedLanguage $expectedLanguage
    Add-Content -LiteralPath $LogPath -Value ("[{0}] Setup source resolved: {1}; ExpectedLanguage={2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$resolvedSetupSourcePath,$expectedLanguage) -Encoding UTF8

    $setupExe = Join-Path $resolvedSetupSourcePath 'setup.exe'
    $installWim = Join-Path $resolvedSetupSourcePath 'sources\install.wim'
    $installEsd = Join-Path $resolvedSetupSourcePath 'sources\install.esd'
    if (-not (Test-Path -LiteralPath $setupExe -PathType Leaf)) { throw "setup.exe not found in SetupSourcePath: $resolvedSetupSourcePath" }
    if (-not (Test-Path -LiteralPath $installWim -PathType Leaf) -and -not (Test-Path -LiteralPath $installEsd -PathType Leaf)) {
        throw "Setup media missing sources\install.wim or sources\install.esd: $resolvedSetupSourcePath"
    }
    Test-SetupSourceLanguage -SourcePath $resolvedSetupSourcePath -ComputerName $ComputerName

    $remoteCache = Join-Path $RemoteSetupCacheRoot $SetupMediaId
    $remoteCacheShare = Convert-ToAdminSharePath -ComputerName $ComputerName -LocalPath $remoteCache
    $remoteSetupShare = Join-Path $remoteCacheShare 'setup.exe'
    if (Test-Path -LiteralPath $remoteSetupShare -PathType Leaf) {
        Add-Content -LiteralPath $LogPath -Value ("[{0}] Setup cache already has setup.exe: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$remoteCache) -Encoding UTF8
        return 'AlreadyCached'
    }

    Add-Content -LiteralPath $LogPath -Value ("[{0}] Copying setup media to remote cache: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$remoteCache) -Encoding UTF8
    Copy-DirectoryContent -SourcePath $resolvedSetupSourcePath -DestinationPath $remoteCacheShare
    return 'Copied'
}

function Get-RemoteVirtualMachineSummary {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    $system = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $ComputerName -ErrorAction Stop
    $manufacturer = [string]$system.Manufacturer
    $model = [string]$system.Model
    $hypervisorPresent = $false
    try { $hypervisorPresent = [bool]$system.HypervisorPresent } catch { $hypervisorPresent = $false }
    $signature = ("{0} {1}" -f $manufacturer,$model)
    $patterns = @('Virtual Machine','VMware','VirtualBox','KVM','QEMU','Xen','HVM domU','Parallels','BHYVE','OpenStack','Google Compute Engine','Amazon EC2')
    $matchedPattern = ''
    foreach ($pattern in $patterns) {
        if ($signature -match [regex]::Escape($pattern)) {
            $matchedPattern = $pattern
            break
        }
    }

    $isVirtual = (-not [string]::IsNullOrWhiteSpace($matchedPattern)) -or $hypervisorPresent
    $evidence = if ($matchedPattern) {
        "Manufacturer=$manufacturer; Model=$model; Pattern=$matchedPattern"
    }
    elseif ($hypervisorPresent) {
        "Manufacturer=$manufacturer; Model=$model; HypervisorPresent=True"
    }
    else {
        "Manufacturer=$manufacturer; Model=$model"
    }

    [pscustomobject]@{
        IsVirtualMachine = $isVirtual
        Evidence = $evidence
    }
}

function Collect-RemoteEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Cycle
    )

    if ($NoCentralLogCollection) { return '' }

    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $ComputerName -LocalPath $RemoteBaseDir
    if (-not (Test-Path -LiteralPath $remoteBaseShare -PathType Container)) { return '' }

    $target = if ($KeepCentralLogHistory) {
        Join-Path (Join-Path $CentralLogRoot $ComputerName) ("Cycle{0}_{1}" -f $Cycle,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    else {
        Join-Path (Join-Path $CentralLogRoot $ComputerName) 'Latest'
    }

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
    SetupCacheAction = ''
    RemoteLogsPath = ''
    PsExecLogPath = $logPath
}

try {
    Add-Content -LiteralPath $logPath -Value ("[{0}] Starting {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Computer) -Encoding UTF8

    if ($DryRun) {
        $reachable = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue
        $adminShare = Test-Path -LiteralPath ("\\{0}\C$" -f $Computer)
        $result.LauncherStatus = if ($reachable -and $adminShare) { 'DRYRUN_READY' } else { 'DRYRUN_UNREACHABLE' }
        $result.Detail = "Ping=$reachable; AdminShare=$adminShare"
        return [pscustomobject]$result
    }

    if ($SkipVirtualMachines) {
        try {
            $vm = Get-RemoteVirtualMachineSummary -ComputerName $Computer
            if ($vm.IsVirtualMachine) {
                $result.LauncherStatus = 'SKIPPED_VIRTUAL_MACHINE'
                $result.RemoteStatus = 'SKIPPED_VIRTUAL_MACHINE'
                $result.RemoteNextAction = 'NO_ACTION_VIRTUAL_MACHINE'
                $result.ExitCode = '0'
                $result.Detail = $vm.Evidence
                $result.SetupCacheAction = 'SkippedVirtualMachine'
                Add-Content -LiteralPath $logPath -Value ("[{0}] Skipped virtual machine before payload/setup copy. {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$vm.Evidence) -Encoding UTF8
                return [pscustomobject]$result
            }
        }
        catch {
            Add-Content -LiteralPath $logPath -Value ("[{0}] WARNING Unable to pre-detect VM status, endpoint guard will still check: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message) -Encoding UTF8
        }
    }

    Copy-RemotePayload -ComputerName $Computer -LogPath $logPath
    $result.SetupCacheAction = Copy-SetupMediaToRemoteCache -ComputerName $Computer -LogPath $logPath

    $psexecArgs = @(
        "\\$Computer",
        '-accepteula',
        '-nobanner',
        '-s',
        'powershell.exe',
        '-NoProfile',
        '-ExecutionPolicy','Bypass',
        '-File',"`"$RemoteScriptPath`""
    ) + $RemoteScriptArgs

    Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec command: {1} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$ResolvedPsExecPath,($psexecArgs -join ' ')) -Encoding UTF8
    $process = Start-Process -FilePath $ResolvedPsExecPath -ArgumentList $psexecArgs -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath

    if ($PsExecTimeoutMinutes -gt 0) {
        if (-not $process.WaitForExit($PsExecTimeoutMinutes * 60 * 1000)) {
            try { $process.Kill() } catch { }
            $result.LauncherStatus = 'PSEXEC_TIMEOUT'
            $result.Detail = "Timed out after $PsExecTimeoutMinutes minute(s)."
        }
    }
    else {
        $process.WaitForExit()
    }

    if ($result.LauncherStatus -eq 'UNKNOWN') {
        $result.ExitCode = $process.ExitCode
        $result.LauncherStatus = if ($process.ExitCode -eq 0) { 'SUCCESS' } else { "PSEXEC_EXIT_$($process.ExitCode)" }
    }

    Start-Sleep -Seconds 3
    $result.RemoteLogsPath = Collect-RemoteEvidence -ComputerName $Computer -Cycle $CycleNumber
    if ($result.RemoteLogsPath) {
        $lastRunPath = Join-Path $result.RemoteLogsPath 'LastRun.json'
        if (Test-Path -LiteralPath $lastRunPath -PathType Leaf) {
            $lastRun = Get-Content -LiteralPath $lastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $result.RemoteStatus = [string]$lastRun.Status
            $result.RemoteNextAction = [string]$lastRun.NextAction
            if ($result.RemoteStatus) {
                $result.LauncherStatus = $result.RemoteStatus
                $result.ExitCode = [string]$lastRun.ExitCode
            }
        }
    }
}
catch {
    $result.LauncherStatus = 'ERROR'
    $result.Detail = $_.Exception.Message
    Add-Content -LiteralPath $logPath -Value ("[{0}] ERROR {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message) -Encoding UTF8
}

[pscustomobject]$result
