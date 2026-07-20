<#
.SYNOPSIS
Starts the Intune Hybrid Join repair LOT launcher GUI.

.VERSION
1.15
#>
param(
    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$GuiVersion = '1.15'

function Get-ToolkitRoot {
    $scriptPath = $PSCommandPath
    if (-not $scriptPath) {
        $scriptPath = $MyInvocation.MyCommand.Path
    }

    $scriptsRoot = Split-Path -Parent $scriptPath
    return Split-Path -Parent $scriptsRoot
}

function Get-ScriptHeaderVersion {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return 'missing'
    }

    try {
        $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $match = [regex]::Match($content, '(?m)^\.VERSION\s*\r?\n\s*([^\r\n]+)')
        if ($match.Success) {
            return $match.Groups[1].Value.Trim()
        }

        $scriptVersionMatch = [regex]::Match($content, '(?m)^\s*\$ScriptVersion\s*=\s*"([^"\r\n]+)"')
        if ($scriptVersionMatch.Success) {
            return $scriptVersionMatch.Groups[1].Value.Trim()
        }
    }
    catch {}

    return 'unknown'
}

function Write-GuiStartupLog {
    param([string]$Message)

    try {
        $line = '{0:yyyy-MM-dd HH:mm:ss} {1}' -f (Get-Date), $Message
        Add-Content -LiteralPath $script:GuiStartupLogPath -Value $line -Encoding UTF8
    }
    catch {}
}
function Import-SmartM365GuiSplash {
    $current = Get-ToolkitRoot
    while ($current) {
        $candidate = Join-Path $current 'SmartM365.GuiSplash.ps1'
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }

        $parent = Split-Path -Parent $current
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) {
            break
        }

        $current = $parent
    }

    return $false
}


function Read-ToolkitConfig {
    param([string]$Path)

    $config = @{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $config
    }

    foreach ($rawLine in @(Get-Content -LiteralPath $Path)) {
        $line = ([string]$rawLine).Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#') -or $line.StartsWith(';')) {
            continue
        }

        $match = [regex]::Match($line, '^(?<name>[A-Za-z_][A-Za-z0-9_]*)=(?<value>.*)$')
        if (-not $match.Success) {
            continue
        }

        $name = $match.Groups['name'].Value.Trim()
        if ($name -notmatch '^EHJIR_') {
            continue
        }

        $value = $match.Groups['value'].Value.Trim()
        if ($value.Length -ge 2 -and $value.StartsWith('"') -and $value.EndsWith('"')) {
            $value = $value.Substring(1, $value.Length - 2)
        }

        $config[$name] = $value
    }

    return $config
}

function Write-ToolkitConfig {
    param(
        [string]$Path,
        [hashtable]$Values
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# SmartM365 Intune Hybrid Join Toolkit GUI options.')
    $lines.Add('# This file is local and is updated by the LOT launcher GUI.')
    $lines.Add('')
    foreach ($name in @($Values.Keys | Sort-Object)) {
        if ($name -notmatch '^EHJIR_') { continue }
        $value = [string]$Values[$name]
        $lines.Add(('{0}={1}' -f $name,$value))
    }

    Set-Content -LiteralPath $Path -Encoding ASCII -Value @($lines)
}

function Get-ToolkitConfigValue {
    param(
        [string]$Name,
        [AllowNull()][string]$Default = $null
    )

    $processValue = [Environment]::GetEnvironmentVariable($Name, 'Process')
    if (-not [string]::IsNullOrWhiteSpace($processValue)) {
        return $processValue
    }

    if ($script:ToolkitConfig -and $script:ToolkitConfig.ContainsKey($Name)) {
        return [string]$script:ToolkitConfig[$Name]
    }

    if ($script:ToolkitDefaultEnvironment -and $script:ToolkitDefaultEnvironment.ContainsKey($Name)) {
        return [string]$script:ToolkitDefaultEnvironment[$Name]
    }

    return $Default
}

function Get-ConfiguredValue {
    param([string]$Name)
    return Get-ToolkitConfigValue -Name $Name -Default ([string]$script:ToolkitDefaultEnvironment[$Name])
}

function Get-TrimmedText {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

function Get-SafeLotName {
    param([string]$Name)

    $safe = [regex]::Replace((Get-TrimmedText -Value $Name), '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        throw 'Enter a LOT name.'
    }

    if ($safe -notmatch '^(?i)LOT-') {
        $safe = "LOT-$safe"
    }

    return $safe
}

function Get-LotsRoot {
    param([string]$RootPath)
    return Join-Path $RootPath 'Lots'
}

function Get-RunsRoot {
    param([string]$RootPath)
    return Join-Path $RootPath 'Runs'
}

function Get-LotRunsRoot {
    param(
        [string]$RootPath,
        [string]$LotName
    )

    return Join-Path (Get-RunsRoot -RootPath $RootPath) $LotName
}

function Get-ToolkitRootFromLotPath {
    param([string]$LotPath)

    $lotsRoot = Split-Path -Parent $LotPath
    return Split-Path -Parent $lotsRoot
}

function New-LotRunContext {
    param(
        [string]$RootPath,
        [string]$LotName
    )

    $runRoot = Get-LotRunsRoot -RootPath $RootPath -LotName $LotName
    $runPath = Join-Path $runRoot (Get-Date -Format 'yyyyMMdd-HHmmss')
    foreach ($folder in @('', 'Logs', 'PsExecLogs', 'Reports', 'CentralLogs', 'Archives', 'State')) {
        $path = if ([string]::IsNullOrWhiteSpace($folder)) { $runPath } else { Join-Path $runPath $folder }
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }

    [pscustomobject]@{
        RunPath        = $runPath
        LogRoot        = Join-Path $runPath 'PsExecLogs'
        ReportRoot     = Join-Path $runPath 'Reports'
        CentralLogRoot = Join-Path $runPath 'CentralLogs'
        ArchiveRoot    = Join-Path $runPath 'Archives'
    }
}

function Get-ComputerNamesFromFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return @()
    }

    Get-Content -LiteralPath $Path |
        ForEach-Object { Get-TrimmedText -Value $_ } |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        Select-Object -Unique
}

function Get-LotFolders {
    param([string]$Root)

    $lotsRoot = Get-LotsRoot -RootPath $Root
    if (-not (Test-Path -LiteralPath $lotsRoot -PathType Container)) {
        return @()
    }

    @(
        Get-ChildItem -LiteralPath $lotsRoot -Directory -Filter 'LOT-*' -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ine 'LOT-TEMPLATE' } |
            Sort-Object Name
    )
}

function Test-LotWrapperSet {
    param([string]$LotPath)

    $required = @(
        'Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd',
        'Run-IntuneHybridJoinRepairWithPsExec-Once.cmd',
        'Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd',
        'Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd'
    )

    foreach ($fileName in $required) {
        if (-not (Test-Path -LiteralPath (Join-Path $LotPath $fileName))) {
            return $false
        }
    }

    return $true
}

function Get-AdDomainText {
    param([string]$LotPath)

    $candidate = Join-Path $LotPath 'AdDomain.txt'
    if (Test-Path -LiteralPath $candidate) {
        $value = Get-TrimmedText -Value (Get-Content -LiteralPath $candidate -Raw)
        if ($value) {
            return $value
        }
    }

    return '(default)'
}

function Get-LotSummary {
    param([System.IO.DirectoryInfo]$Folder)

    $computersPath = Join-Path $Folder.FullName 'Computers.txt'
    $deviceCount = @(Get-ComputerNamesFromFile -Path $computersPath).Count
    $wrappersReady = Test-LotWrapperSet -LotPath $Folder.FullName

    [pscustomobject]@{
        Name          = $Folder.Name
        Path          = $Folder.FullName
        ComputersPath = $computersPath
        DeviceCount   = $deviceCount
        AdDomain      = Get-AdDomainText -LotPath $Folder.FullName
        WrappersReady = $wrappersReady
        Display       = ('{0} ({1} devices)' -f $Folder.Name, $deviceCount)
    }
}

function Invoke-LotWrapperRefresh {
    param(
        [string]$ToolkitRoot,
        [string]$LotPath
    )

    $null = $LotPath
    $script = Join-Path $ToolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1'
    if (-not (Test-Path -LiteralPath $script)) {
        throw "Wrapper refresh script not found: $script"
    }

    & (Resolve-GuiPowerShellPath) -NoProfile -ExecutionPolicy Bypass -File $script -RootPath $ToolkitRoot | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "LOT wrapper refresh failed with exit code $LASTEXITCODE."
    }
}

function New-ToolkitLotFolder {
    param(
        [string]$ToolkitRoot,
        [string]$Name
    )

    $safeName = Get-SafeLotName -Name $Name
    $lotsRoot = Get-LotsRoot -RootPath $ToolkitRoot
    New-Item -ItemType Directory -Path $lotsRoot -Force | Out-Null
    $lotPath = Join-Path $lotsRoot $safeName
    New-Item -ItemType Directory -Path $lotPath -Force | Out-Null

    $computersPath = Join-Path $lotPath 'Computers.txt'
    if (-not (Test-Path -LiteralPath $computersPath)) {
        New-Item -ItemType File -Path $computersPath -Force | Out-Null
    }
    $adDomainPath = Join-Path $lotPath 'AdDomain.txt'
    if (-not (Test-Path -LiteralPath $adDomainPath)) {
        New-Item -ItemType File -Path $adDomainPath -Force | Out-Null
    }

    Invoke-LotWrapperRefresh -ToolkitRoot $ToolkitRoot -LotPath $lotPath
    return $lotPath
}

function ConvertTo-CmdArgument {
    param([string]$Value)

    if ($null -eq $Value) {
        return '""'
    }

    return '"' + ($Value -replace '"', '\"') + '"'
}

function ConvertTo-CmdSetCommand {
    param(
        [string]$Name,
        [string]$Value
    )

    return 'set "{0}={1}"' -f $Name, ($Value -replace '"', '\"')
}


function ConvertTo-CmdWindowTitle {
    param([AllowNull()][string]$Value)

    $title = ([string]$Value -replace '[\r\n\t]+', ' ').Trim()
    $title = $title -replace '[&|<>^]', '-'
    if ($title.Length -gt 240) { $title = $title.Substring(0, 240) }
    return $title
}


function Get-ActiveToolkitLotRuns {
    param([Parameter(Mandatory = $true)][string]$ToolkitRoot)

    $runsRoot = Join-Path $ToolkitRoot 'Runs'
    if (-not (Test-Path -LiteralPath $runsRoot -PathType Container)) { return @() }

    $active = New-Object System.Collections.Generic.List[object]
    foreach ($stateFile in @(Get-ChildItem -LiteralPath $runsRoot -Filter 'ActiveLotRun_*.json' -File -Recurse -ErrorAction SilentlyContinue)) {
        try {
            $state = Get-Content -LiteralPath $stateFile.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ([string]$state.Toolkit -ne 'IntuneHybridJoinToolkit') { continue }
            $processId = [int]$state.ProcessId
            if ($processId -lt 1 -or -not (Get-Process -Id $processId -ErrorAction SilentlyContinue)) { continue }
            if ([string]::IsNullOrWhiteSpace([string]$state.SignalPath)) { continue }
            $active.Add([pscustomobject]@{ State = $state; StatePath = $stateFile.FullName; SignalPath = [string]$state.SignalPath })
        }
        catch { }
    }
    return @($active)
}

function Request-ToolkitLotStops {
    param(
        [Parameter(Mandatory = $true)][string]$ToolkitRoot,
        [switch]$DoNotForceExisting
    )

    $activeRuns = @(Get-ActiveToolkitLotRuns -ToolkitRoot $ToolkitRoot)
    $requested = 0
    $forced = 0
    $alreadyRequested = 0
    foreach ($activeRun in $activeRuns) {
        $signalExists = Test-Path -LiteralPath $activeRun.SignalPath -PathType Leaf
        if ($DoNotForceExisting -and $signalExists) {
            $alreadyRequested++
            continue
        }
        $force = $signalExists
        $signalParent = Split-Path -Parent $activeRun.SignalPath
        if (-not (Test-Path -LiteralPath $signalParent -PathType Container)) {
            New-Item -ItemType Directory -Path $signalParent -Force | Out-Null
        }
        [pscustomobject]@{
            Version = 1
            RequestedUtc = (Get-Date).ToUniversalTime().ToString('o')
            RequestedBy = [Environment]::UserName
            Source = 'GUI'
            Force = [bool]$force
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $activeRun.SignalPath -Encoding UTF8 -Force
        if ($force) { $forced++ } else { $requested++ }
    }

    return [pscustomobject]@{ Active = $activeRuns.Count; Requested = $requested; Forced = $forced; AlreadyRequested = $alreadyRequested }
}
function New-GuiLaunchCommandFile {
    param(
        [Parameter(Mandatory = $true)][string]$WorkingDirectory,
        [Parameter(Mandatory = $true)][string[]]$Commands,
        [string]$NamePrefix = 'EHJIR-GUI',
        [string]$WindowTitle,
        [switch]$PauseWhenDone
    )

    $launcherRoot = Join-Path $env:TEMP 'SmartM365\IntuneHybridJoinToolkit\GuiLaunchers'
    New-Item -ItemType Directory -Path $launcherRoot -Force | Out-Null

    $safePrefix = [regex]::Replace($NamePrefix, '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safePrefix)) { $safePrefix = 'EHJIR-GUI' }

    $launchPath = Join-Path $launcherRoot ('{0}_{1}.cmd' -f $safePrefix,(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    $launchLines = New-Object System.Collections.Generic.List[string]
    $launchLines.Add('@echo off')
    $launchLines.Add('setlocal')
    if (-not [string]::IsNullOrWhiteSpace($WindowTitle)) {
        $launchLines.Add(('title {0}' -f (ConvertTo-CmdWindowTitle -Value $WindowTitle)))
    }
    $launchLines.Add(('cd /d {0}' -f (ConvertTo-CmdArgument -Value $WorkingDirectory)))
    foreach ($command in @($Commands)) {
        if (-not [string]::IsNullOrWhiteSpace($command)) {
            $launchLines.Add($command)
        }
    }
    $launchLines.Add('set "EXITCODE=%ERRORLEVEL%"')
    if ($PauseWhenDone) {
        $launchLines.Add('echo.')
        $launchLines.Add('echo Finished with exit code %EXITCODE%.')
        $launchLines.Add('pause')
    }
    $launchLines.Add('exit /b %EXITCODE%')

    Set-Content -LiteralPath $launchPath -Value $launchLines -Encoding ASCII -Force
    return $launchPath
}

function Start-GuiLaunchCommandFile {
    param(
        [Parameter(Mandatory = $true)][string]$LaunchCommandPath,
        [Parameter(Mandatory = $true)][string]$WorkingDirectory
    )

    Start-Process -FilePath 'cmd.exe' -ArgumentList @('/c', (ConvertTo-CmdArgument -Value $LaunchCommandPath)) -WorkingDirectory $WorkingDirectory -Verb RunAs | Out-Null
}

function Start-LotHtmlReportOpenWatcher {
    param(
        [Parameter(Mandatory = $true)][string]$ReportRoot,
        [int]$TimeoutSeconds = 3600
    )

    $resolvedReportRoot = [System.IO.Path]::GetFullPath($ReportRoot)
    $escapedReportRoot = $resolvedReportRoot.Replace("'", "''")
    $watcherScript = @"
`$ErrorActionPreference = 'SilentlyContinue'
`$reportRoot = '$escapedReportRoot'
`$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
while ((Get-Date) -lt `$deadline) {
    `$report = Get-ChildItem -LiteralPath `$reportRoot -Filter 'PsExec_IntuneHybridJoinRepair_Summary_*.html' -File | Where-Object { `$_.Length -gt 0 } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (`$report) {
        Start-Process -FilePath `$report.FullName | Out-Null
        exit 0
    }
    Start-Sleep -Seconds 2
}
"@
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($watcherScript))
    Start-Process -FilePath (Resolve-GuiPowerShellPath) -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',$encodedCommand) -WindowStyle Hidden | Out-Null
}

function Resolve-GuiPsExecPath {
    param([string]$ToolkitRoot)

    $scriptPsExec = Join-Path $ToolkitRoot 'Scripts\PsExec.exe'
    if (Test-Path -LiteralPath $scriptPsExec) {
        return $scriptPsExec
    }

    $systemPsExec = Join-Path $env:WINDIR 'System32\PsExec.exe'
    if (Test-Path -LiteralPath $systemPsExec) {
        return $systemPsExec
    }

    $command = Get-Command -Name 'PsExec.exe' -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    return $null
}

function Resolve-GuiPowerShellPath {
    $powerShell7 = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    if (Test-Path -LiteralPath $powerShell7 -PathType Leaf) { return $powerShell7 }
    $command = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { return $command.Source }

    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $windowsPowerShell -PathType Leaf) { return $windowsPowerShell }
    $command = Get-Command -Name 'powershell.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { return $command.Source }
    throw 'Windows PowerShell or PowerShell 7 was not found.'
}

function Start-ToolkitLot {
    param(
        [pscustomobject]$Lot,
        [string]$Mode,
        [string[]]$ExtraArguments,
        [hashtable]$Environment
    )

    $lotToolkitRoot = Get-ToolkitRootFromLotPath -LotPath $Lot.Path
    Invoke-LotWrapperRefresh -ToolkitRoot $lotToolkitRoot -LotPath $Lot.Path
    $run = New-LotRunContext -RootPath $lotToolkitRoot -LotName $Lot.Name

    $wrapperMap = @{
        Loop                = 'Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd'
        Once                = 'Run-IntuneHybridJoinRepairWithPsExec-Once.cmd'
        LoopIgnoreRunGuard  = 'Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd'
        OnceIgnoreRunGuard  = 'Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd'
    }

    $wrapperPath = Join-Path $Lot.Path $wrapperMap[$Mode]
    if (-not (Test-Path -LiteralPath $wrapperPath)) {
        throw "Wrapper not found: $wrapperPath"
    }

    $effectiveExtraArguments = New-Object System.Collections.Generic.List[string]
    foreach ($argument in @($ExtraArguments)) {
        if ($null -ne $argument) {
            [void]$effectiveExtraArguments.Add([string]$argument)
        }
    }

    $effectiveEnvironment = @{}
    foreach ($key in $Environment.Keys) { $effectiveEnvironment[$key] = $Environment[$key] }
    $effectiveEnvironment['EHJIR_LOT_DIR'] = $Lot.Path
    $effectiveEnvironment['EHJIR_RUN_DIR'] = $run.RunPath

    $commands = New-Object System.Collections.Generic.List[string]
    foreach ($key in ($effectiveEnvironment.Keys | Sort-Object)) {
        $commands.Add((ConvertTo-CmdSetCommand -Name $key -Value ([string]$effectiveEnvironment[$key])))
    }

    $extra = ''
    if ($effectiveExtraArguments.Count -gt 0) {
        $extra = ' ' + (($effectiveExtraArguments | ForEach-Object { ConvertTo-CmdArgument -Value $_ }) -join ' ')
    }

    $commands.Add(('call {0}{1}' -f (ConvertTo-CmdArgument -Value $wrapperPath), $extra))
    $launchTitle = "{0} - started {1}" -f $Lot.Name,(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $launchCommandPath = New-GuiLaunchCommandFile -WorkingDirectory $run.RunPath -Commands @($commands) -NamePrefix ($Lot.Name + '-' + $Mode) -WindowTitle $launchTitle
    Start-GuiLaunchCommandFile -LaunchCommandPath $launchCommandPath -WorkingDirectory $run.RunPath
    Start-LotHtmlReportOpenWatcher -ReportRoot $run.ReportRoot
}

function New-SingleComputerRunContext {
    param(
        [string]$ToolkitRoot,
        [string]$ComputerName
    )

    $trimmed = $ComputerName.Trim().Trim([char]34)
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        throw 'Enter a computer name.'
    }

    $safeName = [regex]::Replace($trimmed, '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = 'Computer'
    }

    $root = Join-Path (Join-Path $ToolkitRoot 'Runs\SingleComputer') ("{0}_{1}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $safeName)
    foreach ($folder in @('', 'Logs', 'PsExecLogs', 'Reports', 'CentralLogs', 'Archives', 'State')) {
        $path = if ([string]::IsNullOrWhiteSpace($folder)) { $root } else { Join-Path $root $folder }
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }

    $computersPath = Join-Path $root 'Computers.txt'
    Set-Content -LiteralPath $computersPath -Value $trimmed -Encoding UTF8 -Force

    [pscustomobject]@{
        Root           = $root
        ComputersPath  = $computersPath
        ComputerName   = $trimmed
        Name           = $safeName
        LogRoot        = Join-Path $root 'PsExecLogs'
        ReportRoot     = Join-Path $root 'Reports'
        CentralLogRoot = Join-Path $root 'CentralLogs'
        ArchiveRoot    = Join-Path $root 'Archives'
    }
}

function Start-ToolkitSingleComputer {
    param(
        [string]$ToolkitRoot,
        [string]$ComputerName,
        [string]$Mode,
        [string[]]$ExtraArguments,
        [hashtable]$Environment
    )

    $context = New-SingleComputerRunContext -ToolkitRoot $ToolkitRoot -ComputerName $ComputerName
    $script = Join-Path $ToolkitRoot 'Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1'
    if (-not (Test-Path -LiteralPath $script)) {
        throw "Launcher script not found: $script"
    }

    $isDryRun = @($ExtraArguments) -contains '-DryRun'
    $psExecPath = if ($isDryRun) { $null } else { Resolve-GuiPsExecPath -ToolkitRoot $ToolkitRoot }
    if (-not $isDryRun -and -not $psExecPath) {
        throw 'PsExec.exe was not found in Scripts, System32, or PATH.'
    }

    $ignoreGuard = $Mode -like '*IgnoreRunGuard'

    $commands = New-Object System.Collections.Generic.List[string]
    foreach ($key in ($Environment.Keys | Sort-Object)) {
        $commands.Add((ConvertTo-CmdSetCommand -Name $key -Value ([string]$Environment[$key])))
    }

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $script,
        '-ComputerListPath', $context.ComputersPath,
        '-LogRoot', $context.LogRoot,
        '-ReportRoot', $context.ReportRoot,
        '-CentralLogRoot', $context.CentralLogRoot,
        '-ArchiveRoot', $context.ArchiveRoot,
        '-ThrottleLimit', '1',
        '-GlobalConcurrencyLimit', '0'
    )

    if ($psExecPath) {
        $arguments += @('-PsExecPath', $psExecPath)
    }
    $arguments += @('-TechnicianRunGuardHours',[string]$Environment.EHJIR_TECHNICIAN_RUN_GUARD_HOURS,'-CentralLogCollectionMode',[string]$Environment.EHJIR_CENTRAL_LOG_COLLECTION_MODE)
    if ([string]$Environment.EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY -eq '1') { $arguments += '-UseTechnicianRunGuardHistory' }
    if ([string]$Environment.EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY -eq '1') { $arguments += '-IgnoreTechnicianRunGuardHistory' }
    if ($Mode -like 'Once*') {
        $arguments += '-RunOnce'
    }
    if ($ignoreGuard) {
        $arguments += '-IgnoreRunGuard'
    }

    foreach ($argument in @($ExtraArguments)) {
        if ($null -ne $argument) {
            $arguments += [string]$argument
        }
    }

    $commands.Add(('{0} {1}' -f (ConvertTo-CmdArgument -Value (Resolve-GuiPowerShellPath)),(($arguments | ForEach-Object { ConvertTo-CmdArgument -Value $_ }) -join ' ')))
    $launchTitle = "Single {0} [{1}] - started {2}" -f $context.Name,$Mode,(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $launchCommandPath = New-GuiLaunchCommandFile -WorkingDirectory $context.Root -Commands @($commands) -NamePrefix ('Single-' + $context.Name) -WindowTitle $launchTitle -PauseWhenDone
    Start-GuiLaunchCommandFile -LaunchCommandPath $launchCommandPath -WorkingDirectory $context.Root
    return $context
}

function Wait-UiDelay {
    param([int]$Seconds)

    if ($Seconds -gt 0) {
        Start-Sleep -Seconds $Seconds
    }
}

function Open-TextFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType File -Path $Path -Force | Out-Null
    }

    Start-Process -FilePath 'notepad.exe' -ArgumentList (ConvertTo-CmdArgument -Value $Path) | Out-Null
}

function Open-FolderPath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    Start-Process -FilePath 'explorer.exe' -ArgumentList (ConvertTo-CmdArgument -Value $Path) | Out-Null
}

function Get-BooleanText {
    param([object]$CheckBox)
    if ([bool]$CheckBox.IsChecked) { return '1' }
    return '0'
}

function Get-IntText {
    param(
        [object]$TextBox,
        [int]$Default,
        [int]$Minimum = 0
    )

    $value = 0
    if (-not [int]::TryParse($TextBox.Text, [ref]$value)) {
        return [string]$Default
    }

    if ($value -lt $Minimum) {
        return [string]$Minimum
    }

    return [string]$value
}

$toolkitRoot = Get-ToolkitRoot
New-Item -ItemType Directory -Path (Get-LotsRoot -RootPath $toolkitRoot) -Force | Out-Null
New-Item -ItemType Directory -Path (Get-RunsRoot -RootPath $toolkitRoot) -Force | Out-Null
$script:GuiStartupLogPath = Join-Path $toolkitRoot 'GuiLauncherStartup.log'
$script:PsExecLauncherPath = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1'
$script:PsExecLauncherVersion = Get-ScriptHeaderVersion -Path $script:PsExecLauncherPath
$script:ToolkitConfigPath = Join-Path $toolkitRoot 'IntuneHybridJoinToolkit.config'
$script:ToolkitDefaultEnvironment = @{
    EHJIR_THROTTLE = '10'
    EHJIR_GLOBAL_CONCURRENCY_LIMIT = '15'
    EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = '0'
    EHJIR_DELAY_BETWEEN_CYCLES_MINUTES = '5'
    EHJIR_DISABLE_NIGHT_PAUSE = '0'
    EHJIR_NIGHT_PAUSE_START_HOUR = '20'
    EHJIR_NIGHT_PAUSE_END_HOUR = '7'
    EHJIR_INTUNE_RETRY_SLEEP_MINUTES = '3'
    EHJIR_INTUNE_RETRY_MAX_RETRIES = '5'
    EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS = '300'
    EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS = '3'
    EHJIR_STALE_CLEANUP_DELAY_SECONDS = '30'
    EHJIR_REBOOT_DELAY_SECONDS = '300'
    EHJIR_PSEXEC_TIMEOUT_MINUTES = '60'
    EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES = '15'
    EHJIR_ALLOW_DSREG_LEAVE = '1'
    EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT = '1'
    EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER = '1'
    EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE = '0'
    EHJIR_SKIP_VIRTUAL_MACHINES = '1'
    EHJIR_AD_DOMAIN = ''
    EHJIR_GUI_DRY_RUN = '0'
    EHJIR_GUI_AUDIT_ONLY = '0'
    EHJIR_GUI_IGNORE_RUN_GUARD_EVERY_CYCLE = '0'
    EHJIR_GUI_REMOVE_NON_INTUNE_MDM = '0'
    EHJIR_GUI_KEEP_CENTRAL_LOG_HISTORY = '0'
    EHJIR_GUI_NO_CENTRAL_LOG_COLLECTION = '0'
    EHJIR_GUI_SKIP_POST_CYCLE_INTUNE_INVENTORY = '0'
    EHJIR_GUI_MAX_CYCLES = '0'
    EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY = '1'
    EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY = '0'
    EHJIR_TECHNICIAN_RUN_GUARD_HOURS = '12'
    EHJIR_CENTRAL_LOG_COLLECTION_MODE = 'Standard'
}
$script:ToolkitConfig = Read-ToolkitConfig -Path $script:ToolkitConfigPath
Write-GuiStartupLog -Message ('GUI launcher startup; GuiVersion={0}; PsExecLauncherVersion={1}; Root={2}; Script={3}; PID={4}; User={5}' -f $GuiVersion,$script:PsExecLauncherVersion,$toolkitRoot,$PSCommandPath,$PID,[Environment]::UserName)
$launchAllLotStartDelaySeconds = 5

if ($ValidateOnly) {
    $lots = @(Get-LotFolders -Root $toolkitRoot | ForEach-Object { Get-LotSummary -Folder $_ })
    $ready = @($lots | Where-Object { $_.DeviceCount -gt 0 -and $_.WrappersReady }).Count
    Write-Output "Smart Intune Hybrid Join Toolkit LOT Launcher GUI validation completed. GuiVersion=$GuiVersion; PsExecLauncherVersion=$script:PsExecLauncherVersion; Lots=$($lots.Count); Ready=$ready"
    return
}

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$splash = $null
$splashModulePath = Import-SmartM365GuiSplash
if ($splashModulePath) {
    . $splashModulePath
    $splash = Start-SmartM365GuiSplash -Framework Wpf -ProductName 'Hybrid Join Repair LOT Launcher'
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SmartM365 Intune Hybrid Join Repair LOT Launcher"
        Width="1180" Height="790" MinWidth="1060" MinHeight="700"
        WindowStartupLocation="CenterScreen"
        Background="#F5F8FB" FontFamily="Segoe UI" FontSize="12"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
    <Window.Resources>
        <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
        <SolidColorBrush x:Key="TextBrush" Color="#1F2937"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#475569"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#D8E4F0"/>
        <Style TargetType="Button">
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#1F2937"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Margin" Value="0,4,0,10"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Padding" Value="8,5"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Margin" Value="0,4,0,10"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
        </Style>
        <Style x:Key="NumericTextBoxStyle" TargetType="TextBox">
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Margin" Value="0"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="TextAlignment" Value="Center"/>
        </Style>
        <Style x:Key="StepperButtonStyle" TargetType="Button">
            <Setter Property="Width" Value="32"/>
            <Setter Property="MinWidth" Value="32"/>
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Margin" Value="0"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#1F2937"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,4,18,4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="#475569"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#D8E4F0"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Margin" Value="0,0,6,0"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#F8FBFE"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Foreground" Value="#005A9E"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                                <Setter TargetName="TabBorder" Property="Background" Value="#E6F4FF"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="20" Margin="0,0,0,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="190"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <Border Background="#E7F3FF" BorderBrush="#B5DCFF" BorderThickness="1" CornerRadius="8" Padding="10,3" HorizontalAlignment="Left" Margin="0,0,0,14">
                        <TextBlock Text="SMARTM365 DEVICES" Foreground="#005A9E" FontWeight="SemiBold"/>
                    </Border>
                    <TextBlock Text="Hybrid Join Repair LOT Launcher" Foreground="{StaticResource TextBrush}" FontSize="28" FontWeight="SemiBold"/>
                    <TextBlock Text="Prepare and launch Intune Hybrid Join repair campaigns by LOT or single device." Foreground="{StaticResource MutedBrush}" Margin="0,6,0,0"/>
                </StackPanel>
                <Border x:Name="HeaderLogoLink" Grid.Column="1" BorderBrush="#D8E4F0" BorderThickness="1" CornerRadius="8" Background="#F8FBFF" Padding="12" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Cursor="Hand" ToolTip="Open WorkplaceCloudHub.com">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                        <Image x:Name="HeaderLogoImage" Width="128" Height="72" Stretch="Uniform" SnapsToDevicePixels="True" RenderOptions.BitmapScalingMode="HighQuality"/>
                        <TextBlock Text="workplacecloudhub.com" Foreground="#475569" FontSize="11" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <Border Grid.Row="1" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="16" Margin="0,0,0,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse Width="10" Height="10" Fill="#0078D4" Margin="2,0,12,0"/>
                    <StackPanel>
                        <TextBlock x:Name="StatusTitle" Text="Ready" FontWeight="SemiBold" Foreground="#0078D4"/>
                        <TextBlock x:Name="StatusText" Text="Select a LOT or create a new one." Foreground="{StaticResource MutedBrush}" Margin="0,3,0,0"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="RefreshButton" Content="Refresh"/>
                    <Button x:Name="StopAllButton" Content="Stop running LOTs" Background="#FFF4E5" Foreground="#9A3412" BorderBrush="#FDBA74" MinWidth="145"/>
                    <Button x:Name="LaunchAllButton" Content="Launch all ready LOTs" Background="#0078D4" Foreground="White" BorderBrush="#0078D4" MinWidth="160"/>
                </StackPanel>
            </Grid>
        </Border>

        <TabControl Grid.Row="2" x:Name="MainTabs" Background="Transparent" BorderThickness="0" Padding="0">
            <TabItem Header="Existing LOT">
                <Grid Margin="0,14,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="1.05*"/>
                        <ColumnDefinition Width="0.95*"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="0,0,8,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Campaign" Foreground="{StaticResource TextBrush}" FontSize="18" FontWeight="SemiBold"/>
                            <Grid Grid.Row="1" Margin="0,16,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="170"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <TextBlock Text="LOT"/>
                                    <ComboBox x:Name="LotCombo"/>
                                </StackPanel>
                                <Button Grid.Column="1" x:Name="LaunchLotButton" Content="Launch selected LOT" Background="#0078D4" Foreground="White" BorderBrush="#0078D4" Margin="12,22,0,10"/>
                            </Grid>
                            <UniformGrid Grid.Row="2" Columns="3" Margin="0,4,0,12">
                                <StackPanel>
                                    <TextBlock Text="Devices" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="LotDeviceCountText" Text="-" Foreground="{StaticResource MutedBrush}" Margin="0,4,0,0"/>
                                </StackPanel>
                                <StackPanel>
                                    <TextBlock Text="AD domain" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="LotAdDomainText" Text="-" Foreground="{StaticResource MutedBrush}" Margin="0,4,0,0"/>
                                </StackPanel>
                                <StackPanel>
                                    <TextBlock Text="Wrappers" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="LotWrappersText" Text="-" Foreground="{StaticResource MutedBrush}" Margin="0,4,0,0"/>
                                </StackPanel>
                            </UniformGrid>
                            <Grid Grid.Row="3">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="150"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0">
                                    <TextBlock Text="Launch mode"/>
                                    <ComboBox x:Name="LotModeCombo"/>
                                </StackPanel>
                                <StackPanel Grid.Column="1" Margin="12,0,0,0">
                                    <TextBlock Text="Worker limit"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="GlobalLimitDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="GlobalLimitText" Grid.Column="1" Text="15" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="GlobalLimitUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                </StackPanel>
                            </Grid>
                            <WrapPanel Grid.Row="4" VerticalAlignment="Bottom" HorizontalAlignment="Left">
                                <Button x:Name="OpenLotFolderButton" Content="Folder"/>
                                <Button x:Name="OpenLotComputersButton" Content="Computers.txt"/>
                                <Button x:Name="OpenLotAdDomainButton" Content="AD domain"/>
                                <Button x:Name="RefreshWrappersButton" Content="Refresh wrappers"/>
                            </WrapPanel>
                        </Grid>
                    </Border>
                    <Border Grid.Column="1" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="8,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Activity" Foreground="{StaticResource TextBrush}" FontSize="18" FontWeight="SemiBold"/>
                            <TextBox Grid.Row="1" x:Name="ActivityText" Margin="0,16,0,0" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <TabItem Header="Single PC">
                <Border CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="0,14,0,0">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,14,0">
                            <TextBlock Text="Single computer" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,16"/>
                            <TextBlock Text="Computer name"/>
                            <TextBox x:Name="SingleComputerText"/>
                            <TextBlock Text="Launch mode"/>
                            <ComboBox x:Name="SingleModeCombo"/>
                            <Button x:Name="LaunchSingleButton" Content="Launch single computer" Background="#0078D4" Foreground="White" BorderBrush="#0078D4" HorizontalAlignment="Left" MinWidth="180"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="14,0,0,0">
                            <TextBlock Text="Run folder" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,16"/>
                            <TextBox x:Name="SingleRunFolderText" IsReadOnly="True"/>
                            <Button x:Name="OpenSingleRunFolderButton" Content="Open folder" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </TabItem>

            <TabItem Header="New LOT">
                <Border CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="0,14,0,0">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,14,0">
                            <TextBlock Text="Create LOT" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,16"/>
                            <TextBlock Text="LOT name"/>
                            <TextBox x:Name="NewLotNameText"/>
                            <Button x:Name="CreateLotButton" Content="Create LOT" Background="#0078D4" Foreground="White" BorderBrush="#0078D4" HorizontalAlignment="Left" MinWidth="130"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="14,0,0,0">
                            <TextBlock Text="Computers.txt" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,16"/>
                            <TextBox x:Name="NewLotComputersPathText" IsReadOnly="True"/>
                            <Button x:Name="OpenNewLotComputersButton" Content="Open Computers.txt" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </TabItem>

            <TabItem Header="Options">
                <ScrollViewer Margin="0,14,0,0" VerticalScrollBarVisibility="Auto">
                    <Border CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Margin="0,0,18,0">
                                <TextBlock Text="Execution" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,12"/>
                                <WrapPanel>
                                    <CheckBox x:Name="DryRunCheck" Content="Dry run"/>
                                    <CheckBox x:Name="AuditOnlyCheck" Content="Audit only"/>
                                    <CheckBox x:Name="AllowDsregLeaveCheck" Content="Allow dsreg leave"/>
                                    <CheckBox x:Name="AllowStaleCleanupCheck" Content="Allow stale Intune cleanup"/>
                                    <CheckBox x:Name="AllowRebootNoUserCheck" Content="Allow reboot without user"/>
                                    <CheckBox x:Name="AllowRebootAfterLeaveCheck" Content="Allow reboot after leave"/>
                                    <CheckBox x:Name="IgnoreRunGuardCheck" Content="Ignore run guard every cycle"/>
                                    <CheckBox x:Name="RemoveNonIntuneMdmCheck" Content="Remove non-Intune MDM"/>
                                    <CheckBox x:Name="KeepCentralHistoryCheck" Content="Keep central log history"/>
                                    <CheckBox x:Name="NoCentralCollectionCheck" Content="No central log collection"/>
                                    <CheckBox x:Name="SkipPostCycleInventoryCheck" Content="Skip post-cycle Intune inventory"/>
                                    <CheckBox x:Name="SkipVirtualMachinesCheck" Content="Skip virtual machines"/>
                                    <CheckBox x:Name="NightPauseCheck" Content="Pause cycles at night"/>
                                </WrapPanel>
                                <TextBlock Text="AD domain override" Margin="0,18,0,0"/>
                                <TextBox x:Name="AdDomainOverrideText"/>
                            </StackPanel>
                            <Grid Grid.Column="1" Margin="18,0,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0" Margin="0,0,10,0">
                                    <TextBlock Text="Timing" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,12"/>
                                    <TextBlock Text="Throttle"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="ThrottleDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="ThrottleText" Grid.Column="1" Text="10" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="ThrottleUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Delay between cycles (minutes)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="CycleDelayDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="CycleDelayText" Grid.Column="1" Text="5" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="CycleDelayUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Night pause start hour"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="NightPauseStartDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="NightPauseStartText" Grid.Column="1" Text="20" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="NightPauseStartUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Night pause end hour"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="NightPauseEndDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="NightPauseEndText" Grid.Column="1" Text="7" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="NightPauseEndUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Max cycles (0 = unlimited)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="MaxCyclesDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="MaxCyclesText" Grid.Column="1" Text="0" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="MaxCyclesUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="PsExec timeout (minutes)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="PsExecTimeoutDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="PsExecTimeoutText" Grid.Column="1" Text="60" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="PsExecTimeoutUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                </StackPanel>
                                <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                    <TextBlock Text="Guards" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,12"/>
                                    <TextBlock Text="Global worker limit"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="GlobalLimitOptionDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="GlobalLimitOptionText" Grid.Column="1" Text="15" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="GlobalLimitOptionUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Global lease timeout (minutes)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="GlobalLeaseDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="GlobalLeaseText" Grid.Column="1" Text="0" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="GlobalLeaseUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Intune retry sleep (minutes)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="IntuneRetrySleepDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="IntuneRetrySleepText" Grid.Column="1" Text="3" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="IntuneRetrySleepUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Intune retry max"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="IntuneRetryMaxDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="IntuneRetryMaxText" Grid.Column="1" Text="5" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="IntuneRetryMaxUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Stale cleanup delay (seconds)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="StaleCleanupDelayDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="StaleCleanupDelayText" Grid.Column="1" Text="30" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="StaleCleanupDelayUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Reboot delay (seconds)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="RebootDelayDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="RebootDelayText" Grid.Column="1" Text="300" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="RebootDelayUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <Button x:Name="ResetDefaultsButton" Content="Reset defaults" HorizontalAlignment="Right" MinWidth="140" Margin="0,8,0,0"/>
                                </StackPanel>
                            </Grid>
                        </Grid>
                    </Border>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <Border Grid.Row="3" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="10" Margin="0,14,0,0">
            <TextBlock Text="SmartM365 - Intune Hybrid Join repair campaigns" Foreground="{StaticResource MutedBrush}"/>
        </Border>
    </Grid>
</Window>
'@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

$windowIconPath = Join-Path $toolkitRoot 'Scripts\WorkplaceCloudHub.ico'
if (-not (Test-Path -LiteralPath $windowIconPath -PathType Leaf)) {
    $windowIconPath = Join-Path $toolkitRoot 'WorkplaceCloudHub.ico'
}
if (Test-Path -LiteralPath $windowIconPath -PathType Leaf) {
    try {
        $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$windowIconPath)
    }
    catch {}
}

function Find-Control {
    param([string]$Name)
    return $window.FindName($Name)
}

$controls = @{}
@(
    'HeaderLogoLink','HeaderLogoImage','StatusTitle','StatusText','RefreshButton','StopAllButton','LaunchAllButton','LotCombo','LotDeviceCountText',
    'LotAdDomainText','LotWrappersText','LotModeCombo','GlobalLimitText','GlobalLimitDownButton','GlobalLimitUpButton','OpenLotFolderButton',
    'OpenLotComputersButton','OpenLotAdDomainButton','RefreshWrappersButton','LaunchLotButton',
    'ActivityText','SingleComputerText','SingleModeCombo','LaunchSingleButton','SingleRunFolderText',
    'OpenSingleRunFolderButton','NewLotNameText','CreateLotButton','NewLotComputersPathText',
    'OpenNewLotComputersButton','DryRunCheck','AuditOnlyCheck','AllowDsregLeaveCheck',
    'AllowStaleCleanupCheck','AllowRebootNoUserCheck','AllowRebootAfterLeaveCheck',
    'IgnoreRunGuardCheck','RemoveNonIntuneMdmCheck','KeepCentralHistoryCheck',
    'NoCentralCollectionCheck','SkipPostCycleInventoryCheck','SkipVirtualMachinesCheck',
    'AdDomainOverrideText','ThrottleText','ThrottleDownButton','ThrottleUpButton',
    'NightPauseCheck','CycleDelayText','CycleDelayDownButton','CycleDelayUpButton',
    'NightPauseStartText','NightPauseStartDownButton','NightPauseStartUpButton',
    'NightPauseEndText','NightPauseEndDownButton','NightPauseEndUpButton',
    'MaxCyclesText','MaxCyclesDownButton',
    'MaxCyclesUpButton','PsExecTimeoutText','PsExecTimeoutDownButton','PsExecTimeoutUpButton',
    'GlobalLimitOptionText','GlobalLimitOptionDownButton','GlobalLimitOptionUpButton','ResetDefaultsButton',
    'GlobalLeaseText','GlobalLeaseDownButton','GlobalLeaseUpButton',
    'IntuneRetrySleepText','IntuneRetrySleepDownButton','IntuneRetrySleepUpButton',
    'IntuneRetryMaxText','IntuneRetryMaxDownButton','IntuneRetryMaxUpButton',
    'StaleCleanupDelayText','StaleCleanupDelayDownButton','StaleCleanupDelayUpButton',
    'RebootDelayText','RebootDelayDownButton','RebootDelayUpButton'
) | ForEach-Object { $controls[$_] = Find-Control -Name $_ }

function Open-ExternalUrl {
    param([Parameter(Mandatory)][string]$Url)

    try {
        $psi = [System.Diagnostics.ProcessStartInfo]::new($Url)
        $psi.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    }
    catch {
        [System.Windows.MessageBox]::Show($window, "Unable to open:`r`n$Url`r`n`r`n$($_.Exception.Message)", 'SmartM365', 'OK', 'Warning') | Out-Null
    }
}

if ($controls.HeaderLogoLink) {
    $controls.HeaderLogoLink.Add_MouseLeftButtonUp({ Open-ExternalUrl -Url 'https://workplacecloudhub.com' })
}

$headerLogoPath = Join-Path $toolkitRoot 'WorkplaceCloudHub-lockup-WPF.png'
if (Test-Path -LiteralPath $headerLogoPath -PathType Leaf) {
    try {
        $bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bitmap.BeginInit()
        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.UriSource = [Uri]$headerLogoPath
        $bitmap.EndInit()
        $bitmap.Freeze()
        $controls.HeaderLogoImage.Source = $bitmap
    }
    catch {}
}

$script:Lots = @()
$script:SelectedLot = $null
$script:LastSingleRunFolder = $null
$script:LaunchedLotPaths = @{}

function Add-Status {
    param(
        [string]$Message,
        [string]$Title = 'Ready'
    )

    $controls.StatusTitle.Text = $Title
    $controls.StatusText.Text = $Message
    $controls.ActivityText.AppendText(('{0:HH:mm:ss}  {1}' -f (Get-Date), $Message) + [Environment]::NewLine)
    $controls.ActivityText.ScrollToEnd()
}

function Show-GuiError {
    param([string]$Message)
    [System.Windows.MessageBox]::Show($window, $Message, 'SmartM365', 'OK', 'Error') | Out-Null
}

function Initialize-Combo {
    param(
        [object]$Combo,
        [string[]]$Values,
        [string]$Selected
    )

    $Combo.Items.Clear()
    foreach ($value in $Values) {
        [void]$Combo.Items.Add($value)
    }
    $Combo.SelectedItem = $Selected
}

function Invoke-NumericStepperClick {
    param(
        [object]$Sender,
        [object]$EventArgs
    )

    $fieldName = 'numeric field'
    try {
        $settings = $Sender.Tag
        if (-not $settings) { return }

        $fieldName = [string]$settings.TextBoxName
        $textBox = $controls[$fieldName]
        if (-not $textBox) { return }

        $current = 0
        if (-not [int]::TryParse($textBox.Text, [ref]$current)) {
            $current = [int]$settings.Default
        }

        $next = $current + ([int]$settings.Direction * [int]$settings.Step)
        if ($next -lt [int]$settings.Minimum) {
            $next = [int]$settings.Minimum
        }

        if ($null -ne $settings.Maximum -and $next -gt [int]$settings.Maximum) {
            $next = [int]$settings.Maximum
        }

        $textBox.Text = [string]$next
        $textBox.CaretIndex = $textBox.Text.Length
        if ($EventArgs) { $EventArgs.Handled = $true }
    }
    catch {
        Show-GuiError ("Unable to update numeric value for {0}: {1}" -f $fieldName,$_.Exception.Message)
    }
}

function Register-NumericStepper {
    param(
        [string]$TextBoxName,
        [string]$DownButtonName,
        [string]$UpButtonName,
        [int]$Default,
        [int]$Minimum = 0,
        [Nullable[int]]$Maximum = $null,
        [int]$Step = 1
    )

    $textBox = $controls[$TextBoxName]
    $downButton = $controls[$DownButtonName]
    $upButton = $controls[$UpButtonName]
    if (-not $textBox -or -not $downButton -or -not $upButton) { return }

    $downButton.Tag = [pscustomobject]@{ TextBoxName = $TextBoxName; Default = $Default; Minimum = $Minimum; Maximum = $Maximum; Step = $Step; Direction = -1 }
    $upButton.Tag = [pscustomobject]@{ TextBoxName = $TextBoxName; Default = $Default; Minimum = $Minimum; Maximum = $Maximum; Step = $Step; Direction = 1 }
    $downButton.Add_Click({ param($sender, $eventArgs) Invoke-NumericStepperClick -Sender $sender -EventArgs $eventArgs })
    $upButton.Add_Click({ param($sender, $eventArgs) Invoke-NumericStepperClick -Sender $sender -EventArgs $eventArgs })
    $textBox.Add_PreviewTextInput({
        param($sender, $eventArgs)
        try {
            if ($eventArgs.Text -notmatch '^[0-9]+$') {
                $eventArgs.Handled = $true
            }
        }
        catch {
            $fieldName = 'numeric field'
            if ($sender -and $sender.Name) { $fieldName = [string]$sender.Name }
            Show-GuiError ("Unable to validate numeric input for {0}: {1}" -f $fieldName,$_.Exception.Message)
        }
    })
}

function Register-NumericSteppers {
    Register-NumericStepper -TextBoxName 'GlobalLimitText' -DownButtonName 'GlobalLimitDownButton' -UpButtonName 'GlobalLimitUpButton' -Default 15 -Minimum 1
    Register-NumericStepper -TextBoxName 'ThrottleText' -DownButtonName 'ThrottleDownButton' -UpButtonName 'ThrottleUpButton' -Default 10 -Minimum 1
    Register-NumericStepper -TextBoxName 'CycleDelayText' -DownButtonName 'CycleDelayDownButton' -UpButtonName 'CycleDelayUpButton' -Default 5 -Minimum 0
    Register-NumericStepper -TextBoxName 'NightPauseStartText' -DownButtonName 'NightPauseStartDownButton' -UpButtonName 'NightPauseStartUpButton' -Default 20 -Minimum 0 -Maximum 23
    Register-NumericStepper -TextBoxName 'NightPauseEndText' -DownButtonName 'NightPauseEndDownButton' -UpButtonName 'NightPauseEndUpButton' -Default 7 -Minimum 0 -Maximum 23
    Register-NumericStepper -TextBoxName 'MaxCyclesText' -DownButtonName 'MaxCyclesDownButton' -UpButtonName 'MaxCyclesUpButton' -Default 0 -Minimum 0
    Register-NumericStepper -TextBoxName 'PsExecTimeoutText' -DownButtonName 'PsExecTimeoutDownButton' -UpButtonName 'PsExecTimeoutUpButton' -Default 60 -Minimum 1 -Step 5
    Register-NumericStepper -TextBoxName 'GlobalLimitOptionText' -DownButtonName 'GlobalLimitOptionDownButton' -UpButtonName 'GlobalLimitOptionUpButton' -Default 15 -Minimum 1
    Register-NumericStepper -TextBoxName 'GlobalLeaseText' -DownButtonName 'GlobalLeaseDownButton' -UpButtonName 'GlobalLeaseUpButton' -Default 0 -Minimum 0
    Register-NumericStepper -TextBoxName 'IntuneRetrySleepText' -DownButtonName 'IntuneRetrySleepDownButton' -UpButtonName 'IntuneRetrySleepUpButton' -Default 3 -Minimum 0
    Register-NumericStepper -TextBoxName 'IntuneRetryMaxText' -DownButtonName 'IntuneRetryMaxDownButton' -UpButtonName 'IntuneRetryMaxUpButton' -Default 5 -Minimum 0
    Register-NumericStepper -TextBoxName 'StaleCleanupDelayText' -DownButtonName 'StaleCleanupDelayDownButton' -UpButtonName 'StaleCleanupDelayUpButton' -Default 30 -Minimum 0 -Step 5
    Register-NumericStepper -TextBoxName 'RebootDelayText' -DownButtonName 'RebootDelayDownButton' -UpButtonName 'RebootDelayUpButton' -Default 300 -Minimum 0 -Step 30
}

Register-NumericSteppers
Initialize-Combo -Combo $controls.LotModeCombo -Values @('Loop','Once','LoopIgnoreRunGuard','OnceIgnoreRunGuard') -Selected 'Loop'
Initialize-Combo -Combo $controls.SingleModeCombo -Values @('Once','OnceIgnoreRunGuard','Loop','LoopIgnoreRunGuard') -Selected 'Once'

function Initialize-Options {
    $controls.AllowDsregLeaveCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_ALLOW_DSREG_LEAVE') -eq '1')
    $controls.AllowStaleCleanupCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT') -eq '1')
    $controls.AllowRebootNoUserCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER') -eq '1')
    $controls.AllowRebootAfterLeaveCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE') -eq '1')
    $controls.SkipVirtualMachinesCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_SKIP_VIRTUAL_MACHINES') -eq '1')
    $controls.NightPauseCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_DISABLE_NIGHT_PAUSE') -ne '1')
    $controls.ThrottleText.Text = Get-ConfiguredValue 'EHJIR_THROTTLE'
    $controls.CycleDelayText.Text = Get-ConfiguredValue 'EHJIR_DELAY_BETWEEN_CYCLES_MINUTES'
    $controls.NightPauseStartText.Text = Get-ConfiguredValue 'EHJIR_NIGHT_PAUSE_START_HOUR'
    $controls.NightPauseEndText.Text = Get-ConfiguredValue 'EHJIR_NIGHT_PAUSE_END_HOUR'
    $controls.PsExecTimeoutText.Text = Get-ConfiguredValue 'EHJIR_PSEXEC_TIMEOUT_MINUTES'
    $controls.GlobalLimitText.Text = Get-ConfiguredValue 'EHJIR_GLOBAL_CONCURRENCY_LIMIT'
    $controls.GlobalLimitOptionText.Text = Get-ConfiguredValue 'EHJIR_GLOBAL_CONCURRENCY_LIMIT'
    $controls.GlobalLeaseText.Text = Get-ConfiguredValue 'EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES'
    $controls.IntuneRetrySleepText.Text = Get-ConfiguredValue 'EHJIR_INTUNE_RETRY_SLEEP_MINUTES'
    $controls.IntuneRetryMaxText.Text = Get-ConfiguredValue 'EHJIR_INTUNE_RETRY_MAX_RETRIES'
    $controls.StaleCleanupDelayText.Text = Get-ConfiguredValue 'EHJIR_STALE_CLEANUP_DELAY_SECONDS'
    $controls.RebootDelayText.Text = Get-ConfiguredValue 'EHJIR_REBOOT_DELAY_SECONDS'
    $controls.AdDomainOverrideText.Text = Get-ConfiguredValue 'EHJIR_AD_DOMAIN'
    $controls.DryRunCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_GUI_DRY_RUN') -eq '1')
    $controls.AuditOnlyCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_GUI_AUDIT_ONLY') -eq '1')
    $controls.IgnoreRunGuardCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_GUI_IGNORE_RUN_GUARD_EVERY_CYCLE') -eq '1')
    $controls.RemoveNonIntuneMdmCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_GUI_REMOVE_NON_INTUNE_MDM') -eq '1')
    $controls.KeepCentralHistoryCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_GUI_KEEP_CENTRAL_LOG_HISTORY') -eq '1')
    $controls.NoCentralCollectionCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_GUI_NO_CENTRAL_LOG_COLLECTION') -eq '1')
    $controls.SkipPostCycleInventoryCheck.IsChecked = ((Get-ConfiguredValue 'EHJIR_GUI_SKIP_POST_CYCLE_INTUNE_INVENTORY') -eq '1')
    $controls.MaxCyclesText.Text = Get-ConfiguredValue 'EHJIR_GUI_MAX_CYCLES'
}

function Get-ToolkitDefaultConfig {
    $defaults = @{}
    foreach ($key in @($script:ToolkitDefaultEnvironment.Keys)) {
        $defaults[$key] = [string]$script:ToolkitDefaultEnvironment[$key]
    }
    return $defaults
}

function Save-GuiOptions {
    param([switch]$Quiet)

    try {
        $environment = Get-LauncherOptionEnvironment
        Write-ToolkitConfig -Path $script:ToolkitConfigPath -Values $environment
        $script:ToolkitConfig = Read-ToolkitConfig -Path $script:ToolkitConfigPath
        if (-not $Quiet) {
            Add-Status -Title 'Options saved' -Message ("Saved GUI options to {0}." -f $script:ToolkitConfigPath)
        }
    }
    catch {
        if ($Quiet) { return }
        Show-GuiError ("Unable to save GUI options: {0}" -f $_.Exception.Message)
    }
}

function Reset-GuiOptionsToDefaults {
    try {
        Write-ToolkitConfig -Path $script:ToolkitConfigPath -Values (Get-ToolkitDefaultConfig)
        $script:ToolkitConfig = Read-ToolkitConfig -Path $script:ToolkitConfigPath
        Initialize-Options
        Add-Status -Title 'Defaults restored' -Message ("Reset GUI options to defaults in {0}." -f $script:ToolkitConfigPath)
    }
    catch {
        Show-GuiError ("Unable to reset GUI options: {0}" -f $_.Exception.Message)
    }
}

Initialize-Options

function Get-LauncherOptionArguments {
    $arguments = New-Object System.Collections.Generic.List[string]
    $maxCycles = Get-IntText -TextBox $controls.MaxCyclesText -Default 0 -Minimum 0
    if ([int]$maxCycles -gt 0) {
        $arguments.Add('-MaxCycles')
        $arguments.Add($maxCycles)
    }
    if ([bool]$controls.DryRunCheck.IsChecked) { $arguments.Add('-DryRun') }
    if ([bool]$controls.AuditOnlyCheck.IsChecked) { $arguments.Add('-AuditOnly') }
    if ([bool]$controls.IgnoreRunGuardCheck.IsChecked) { $arguments.Add('-IgnoreRunGuardEveryCycle') }
    if ([bool]$controls.RemoveNonIntuneMdmCheck.IsChecked) { $arguments.Add('-AllowRemoveNonIntuneMdmEnrollment') }
    if ([bool]$controls.KeepCentralHistoryCheck.IsChecked) { $arguments.Add('-KeepCentralLogHistory') }
    if ([bool]$controls.NoCentralCollectionCheck.IsChecked) { $arguments.Add('-NoCentralLogCollection') }
    if ([bool]$controls.SkipPostCycleInventoryCheck.IsChecked) { $arguments.Add('-SkipPostCycleIntuneInventory') }
    return @($arguments)
}

function Get-LauncherOptionEnvironment {
    $globalLimit = Get-IntText -TextBox $controls.GlobalLimitText -Default 15 -Minimum 1
    $controls.GlobalLimitOptionText.Text = $globalLimit

    $environment = @{
        EHJIR_THROTTLE                                 = Get-IntText -TextBox $controls.ThrottleText -Default 10 -Minimum 1
        EHJIR_GLOBAL_CONCURRENCY_LIMIT                 = $globalLimit
        EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = Get-IntText -TextBox $controls.GlobalLeaseText -Default 0 -Minimum 0
        EHJIR_DELAY_BETWEEN_CYCLES_MINUTES             = Get-IntText -TextBox $controls.CycleDelayText -Default 5 -Minimum 0
        EHJIR_DISABLE_NIGHT_PAUSE                      = if ([bool]$controls.NightPauseCheck.IsChecked) { '0' } else { '1' }
        EHJIR_NIGHT_PAUSE_START_HOUR                   = Get-IntText -TextBox $controls.NightPauseStartText -Default 20 -Minimum 0
        EHJIR_NIGHT_PAUSE_END_HOUR                     = Get-IntText -TextBox $controls.NightPauseEndText -Default 7 -Minimum 0
        EHJIR_INTUNE_RETRY_SLEEP_MINUTES               = Get-IntText -TextBox $controls.IntuneRetrySleepText -Default 3 -Minimum 0
        EHJIR_INTUNE_RETRY_MAX_RETRIES                 = Get-IntText -TextBox $controls.IntuneRetryMaxText -Default 5 -Minimum 0
        EHJIR_STALE_CLEANUP_DELAY_SECONDS              = Get-IntText -TextBox $controls.StaleCleanupDelayText -Default 30 -Minimum 0
        EHJIR_REBOOT_DELAY_SECONDS                     = Get-IntText -TextBox $controls.RebootDelayText -Default 300 -Minimum 0
        EHJIR_PSEXEC_TIMEOUT_MINUTES                   = Get-IntText -TextBox $controls.PsExecTimeoutText -Default 60 -Minimum 1
        EHJIR_ALLOW_DSREG_LEAVE                        = Get-BooleanText -CheckBox $controls.AllowDsregLeaveCheck
        EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT     = Get-BooleanText -CheckBox $controls.AllowStaleCleanupCheck
        EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER    = Get-BooleanText -CheckBox $controls.AllowRebootNoUserCheck
        EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE           = Get-BooleanText -CheckBox $controls.AllowRebootAfterLeaveCheck
        EHJIR_SKIP_VIRTUAL_MACHINES                    = Get-BooleanText -CheckBox $controls.SkipVirtualMachinesCheck
        EHJIR_GUI_DRY_RUN                              = Get-BooleanText -CheckBox $controls.DryRunCheck
        EHJIR_GUI_AUDIT_ONLY                           = Get-BooleanText -CheckBox $controls.AuditOnlyCheck
        EHJIR_GUI_IGNORE_RUN_GUARD_EVERY_CYCLE         = Get-BooleanText -CheckBox $controls.IgnoreRunGuardCheck
        EHJIR_GUI_REMOVE_NON_INTUNE_MDM                = Get-BooleanText -CheckBox $controls.RemoveNonIntuneMdmCheck
        EHJIR_GUI_KEEP_CENTRAL_LOG_HISTORY             = Get-BooleanText -CheckBox $controls.KeepCentralHistoryCheck
        EHJIR_GUI_NO_CENTRAL_LOG_COLLECTION            = Get-BooleanText -CheckBox $controls.NoCentralCollectionCheck
        EHJIR_GUI_SKIP_POST_CYCLE_INTUNE_INVENTORY     = Get-BooleanText -CheckBox $controls.SkipPostCycleInventoryCheck
        EHJIR_GUI_MAX_CYCLES                           = Get-IntText -TextBox $controls.MaxCyclesText -Default 0 -Minimum 0
        EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY         = Get-ConfiguredValue -Name 'EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY'
        EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY      = Get-ConfiguredValue -Name 'EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY'
        EHJIR_TECHNICIAN_RUN_GUARD_HOURS               = Get-ConfiguredValue -Name 'EHJIR_TECHNICIAN_RUN_GUARD_HOURS'
        EHJIR_CENTRAL_LOG_COLLECTION_MODE              = Get-ConfiguredValue -Name 'EHJIR_CENTRAL_LOG_COLLECTION_MODE'
    }

    if (-not [string]::IsNullOrWhiteSpace($controls.AdDomainOverrideText.Text)) {
        $environment.EHJIR_AD_DOMAIN = $controls.AdDomainOverrideText.Text.Trim()
    }

    return $environment
}

function Test-LotAlreadyLaunched {
    param([AllowNull()][object]$Lot)

    return ($null -ne $Lot -and -not [string]::IsNullOrWhiteSpace($Lot.Path) -and $script:LaunchedLotPaths.ContainsKey($Lot.Path))
}

function Register-LotLaunch {
    param([Parameter(Mandatory=$true)][object]$Lot)

    $script:LaunchedLotPaths[$Lot.Path] = [pscustomobject]@{
        Name      = $Lot.Name
        StartedAt = Get-Date
    }
}

function Get-LaunchableLotSummaries {
    @($script:Lots | Where-Object { $_.DeviceCount -gt 0 -and $_.WrappersReady -and -not (Test-LotAlreadyLaunched -Lot $_) })
}

function Set-LotUiAvailability {
    param([bool]$Enabled)

    foreach ($name in @('OpenLotFolderButton','OpenLotComputersButton','OpenLotAdDomainButton','RefreshWrappersButton','LaunchLotButton')) {
        $controls[$name].IsEnabled = $Enabled
    }
}

function Update-SelectedLotView {
    $script:SelectedLot = $controls.LotCombo.SelectedItem
    if (-not $script:SelectedLot) {
        $controls.LotDeviceCountText.Text = '-'
        $controls.LotAdDomainText.Text = '-'
        $controls.LotWrappersText.Text = '-'
        Set-LotUiAvailability -Enabled $false
        return
    }

    $controls.LotDeviceCountText.Text = [string]$script:SelectedLot.DeviceCount
    $controls.LotAdDomainText.Text = $script:SelectedLot.AdDomain
    $alreadyLaunched = Test-LotAlreadyLaunched -Lot $script:SelectedLot
    $controls.LotWrappersText.Text = if ($alreadyLaunched) { 'Launched' } elseif ($script:SelectedLot.WrappersReady) { 'Ready' } else { 'Missing' }
    Set-LotUiAvailability -Enabled $true
    $controls.LaunchLotButton.IsEnabled = (-not $alreadyLaunched -and $script:SelectedLot.DeviceCount -gt 0 -and $script:SelectedLot.WrappersReady)
    Add-Status -Title 'Selected' -Message ("{0}: {1} device(s)." -f $script:SelectedLot.Name, $script:SelectedLot.DeviceCount)
}

function Refresh-LotList {
    Invoke-LotWrapperRefresh -ToolkitRoot $toolkitRoot -LotPath $null
    $script:Lots = @(Get-LotFolders -Root $toolkitRoot | ForEach-Object { Get-LotSummary -Folder $_ })
    $previous = if ($script:SelectedLot) { $script:SelectedLot.Name } else { $null }

    $controls.LotCombo.Items.Clear()
    foreach ($lot in $script:Lots) {
        [void]$controls.LotCombo.Items.Add($lot)
    }

    $controls.LotCombo.DisplayMemberPath = 'Display'
    $selected = $script:Lots | Where-Object { $_.Name -eq $previous } | Select-Object -First 1
    if (-not $selected) {
        $selected = $script:Lots | Select-Object -First 1
    }

    $controls.LotCombo.SelectedItem = $selected
    Update-SelectedLotView
    Add-Status -Message ("{0} LOT(s) found. {1} ready for launch." -f $script:Lots.Count, @(Get-LaunchableLotSummaries).Count)
}

$controls.LotCombo.Add_SelectionChanged({ Update-SelectedLotView })
$controls.RefreshButton.Add_Click({ try { Refresh-LotList } catch { Show-GuiError $_.Exception.Message } })
$controls.StopAllButton.Add_Click({
    try {
        $stopResult = Request-ToolkitLotStops -ToolkitRoot $toolkitRoot
        if ($stopResult.Active -eq 0) { Add-Status -Title 'Stop' -Message 'No active LOT launcher was found.' }
        elseif ($stopResult.Forced -gt 0) { Add-Status -Title 'Forced stop' -Message ("Forced local stop requested for {0} active LOT(s). Verify remote targets before relaunch." -f $stopResult.Forced) }
        else { Add-Status -Title 'Controlled stop' -Message ("Stop requested for {0} active LOT(s). Click Stop again to force local workers if needed." -f $stopResult.Requested) }
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.OpenLotFolderButton.Add_Click({ if ($script:SelectedLot) { Open-FolderPath -Path $script:SelectedLot.Path } })
$controls.OpenLotComputersButton.Add_Click({ if ($script:SelectedLot) { Open-TextFile -Path $script:SelectedLot.ComputersPath } })
$controls.OpenLotAdDomainButton.Add_Click({ if ($script:SelectedLot) { Open-TextFile -Path (Join-Path $script:SelectedLot.Path 'AdDomain.txt') } })
$controls.RefreshWrappersButton.Add_Click({
    try {
        if (-not $script:SelectedLot) { return }
        Invoke-LotWrapperRefresh -ToolkitRoot $toolkitRoot -LotPath $script:SelectedLot.Path
        Add-Status -Message ("Wrappers refreshed for {0}." -f $script:SelectedLot.Name)
        Refresh-LotList
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.LaunchLotButton.Add_Click({
    try {
        if (-not $script:SelectedLot) { return }
        if ($script:SelectedLot.DeviceCount -le 0) { throw 'Selected LOT has no device in Computers.txt.' }
        if (-not $script:SelectedLot.WrappersReady) { throw 'Selected LOT wrappers are missing. Refresh wrappers first.' }
        if (Test-LotAlreadyLaunched -Lot $script:SelectedLot) {
            Add-Status -Title 'Skipped' -Message ("LOT {0} was already launched in this GUI session." -f $script:SelectedLot.Name)
            Update-SelectedLotView
            return
        }
        Save-GuiOptions -Quiet
        Start-ToolkitLot -Lot $script:SelectedLot -Mode ([string]$controls.LotModeCombo.SelectedItem) -ExtraArguments (Get-LauncherOptionArguments) -Environment (Get-LauncherOptionEnvironment)
        Register-LotLaunch -Lot $script:SelectedLot
        Update-SelectedLotView
        Add-Status -Title 'Launched' -Message ("Launched LOT {0}." -f $script:SelectedLot.Name)
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.LaunchAllButton.Add_Click({
    try {
        $lots = @(Get-LaunchableLotSummaries)
        if ($lots.Count -eq 0) { throw 'No ready LOT with devices was found.' }
        foreach ($lot in $lots) {
            Save-GuiOptions -Quiet
            Start-ToolkitLot -Lot $lot -Mode ([string]$controls.LotModeCombo.SelectedItem) -ExtraArguments (Get-LauncherOptionArguments) -Environment (Get-LauncherOptionEnvironment)
            Register-LotLaunch -Lot $lot
            Add-Status -Title 'Launch all' -Message ("Launched {0}. Next LOT starts in {1}s." -f $lot.Name, $launchAllLotStartDelaySeconds)
            Wait-UiDelay -Seconds $launchAllLotStartDelaySeconds
        }
        Refresh-LotList
        Add-Status -Title 'Launch all' -Message ("Launched {0} LOT(s)." -f $lots.Count)
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.LaunchSingleButton.Add_Click({
    try {
        $computer = $controls.SingleComputerText.Text.Trim()
        if (-not $computer) { throw 'Enter a computer name.' }
        Save-GuiOptions -Quiet
        $context = Start-ToolkitSingleComputer -ToolkitRoot $toolkitRoot -ComputerName $computer -Mode ([string]$controls.SingleModeCombo.SelectedItem) -ExtraArguments (Get-LauncherOptionArguments) -Environment (Get-LauncherOptionEnvironment)
        $script:LastSingleRunFolder = $context.Root
        $controls.SingleRunFolderText.Text = $context.Root
        Add-Status -Title 'Launched' -Message ("Launched single computer run for {0}." -f $computer)
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.OpenSingleRunFolderButton.Add_Click({
    if ($script:LastSingleRunFolder) {
        Open-FolderPath -Path $script:LastSingleRunFolder
    }
})
$controls.CreateLotButton.Add_Click({
    try {
        $lotPath = New-ToolkitLotFolder -ToolkitRoot $toolkitRoot -Name $controls.NewLotNameText.Text
        $computersPath = Join-Path $lotPath 'Computers.txt'
        $controls.NewLotComputersPathText.Text = $computersPath
        Add-Status -Title 'Created' -Message ("Created LOT {0}." -f (Split-Path -Leaf $lotPath))
        Refresh-LotList
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.OpenNewLotComputersButton.Add_Click({
    if ($controls.NewLotComputersPathText.Text) {
        Open-TextFile -Path $controls.NewLotComputersPathText.Text
    }
})
$script:SyncingGlobalLimitText = $false
function Sync-GlobalLimitText {
    param(
        [object]$SourceTextBox,
        [object]$TargetTextBox
    )

    if ($script:SyncingGlobalLimitText) { return }

    try {
        if (-not $SourceTextBox -or -not $TargetTextBox) { return }
        $value = [string]$SourceTextBox.Text
        if ([string]::IsNullOrWhiteSpace($value)) { return }
        if ([string]$TargetTextBox.Text -eq $value) { return }

        $script:SyncingGlobalLimitText = $true
        $TargetTextBox.Text = $value
    }
    catch {
        Show-GuiError ("Unable to synchronize worker limit fields: {0}" -f $_.Exception.Message)
    }
    finally {
        $script:SyncingGlobalLimitText = $false
    }
}

$controls.GlobalLimitOptionText.Add_TextChanged({ try { Sync-GlobalLimitText -SourceTextBox $controls.GlobalLimitOptionText -TargetTextBox $controls.GlobalLimitText } catch { Show-GuiError ("Unable to synchronize worker limit fields: {0}" -f $_.Exception.Message) } }.GetNewClosure())
$controls.GlobalLimitText.Add_TextChanged({ try { Sync-GlobalLimitText -SourceTextBox $controls.GlobalLimitText -TargetTextBox $controls.GlobalLimitOptionText } catch { Show-GuiError ("Unable to synchronize worker limit fields: {0}" -f $_.Exception.Message) } }.GetNewClosure())
$controls.ResetDefaultsButton.Add_Click({ Reset-GuiOptionsToDefaults })

$window.Add_Closing({
    param($sender, $eventArgs)

    Save-GuiOptions -Quiet
    $activeRuns = @(Get-ActiveToolkitLotRuns -ToolkitRoot $toolkitRoot)
    if ($activeRuns.Count -eq 0) { return }

    $choice = [System.Windows.MessageBox]::Show(
        ("{0} LOT launcher(s) are still running.`n`nYes: request a controlled stop, then close this GUI.`nNo: close this GUI and leave the LOTs running.`nCancel: keep this GUI open." -f $activeRuns.Count),
        'Running LOTs',
        [System.Windows.MessageBoxButton]::YesNoCancel,
        [System.Windows.MessageBoxImage]::Warning
    )

    if ($choice -eq [System.Windows.MessageBoxResult]::Cancel) {
        $eventArgs.Cancel = $true
        return
    }

    if ($choice -eq [System.Windows.MessageBoxResult]::Yes) {
        try {
            $stopResult = Request-ToolkitLotStops -ToolkitRoot $toolkitRoot -DoNotForceExisting
            Add-Status -Title 'Controlled stop' -Message ("Stop requested for {0} active LOT(s)." -f $stopResult.Requested)
        }
        catch { Show-GuiError $_.Exception.Message }
    }
})
$window.Add_Closed({
    if ($splash) {
        Close-SmartM365GuiSplash -Splash $splash
    }
})

if (Get-Command -Name Set-SmartM365WpfWindowVisible -ErrorAction SilentlyContinue) {
    $window.Add_SourceInitialized({
        Set-SmartM365WpfWindowVisible -Window $window
    })
}

try {
    Refresh-LotList
    Add-Status -Title 'Ready' -Message ('GUI {0}; PsExec launcher {1}; {2} LOT(s), {3} ready. Startup log: {4}' -f $GuiVersion,$script:PsExecLauncherVersion,$script:Lots.Count,@(Get-LaunchableLotSummaries).Count,$script:GuiStartupLogPath)
    if ($splash) {
        Close-SmartM365GuiSplash -Splash $splash
    }
} catch {
    if ($splash) {
        Close-SmartM365GuiSplash -Splash $splash
    }
    Show-GuiError $_.Exception.Message
}

[void]$window.ShowDialog()

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA1rRZVQ+9EGhsc
# HPGPK2lDq/ie+oxDyL5GP+jm3CAoNqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDG5Yr9zbtcPkXOm5pbCWL5
# D/lqdbctlZH+CDWl6/eiNjANBgkqhkiG9w0BAQEFAASCAYCJkgkKWv9zRYrzyvB9
# alQJaMHiKFLYL4fwVskAm3f5PtP4qKs99bZZYT1o8wv7boy32hCCGWYPQDUHEuad
# nZfd+jcpPT/xvQOU9KHxcwynhsEFYfw8OcbjlfabuCiOgtWbPv1fbVQkdBGMmhJr
# ZUX3fRy+CyKotsYyyolgJbSV4FzdwNMnQdla9b0a44gcuwewt9etRHXy61Fl8Cp3
# 9zBYcpajAAvBq+OOqa5mHQczmvjO5QygTSepNTjhB8eWj/5eIoR/AUQKwTUd+YrI
# Ead1YgjAmrTnLebnYvuLOuxwDT4ZNOKpB1higsXrPSxQb7C6b7wG7yfTAyBPuCA2
# 20gyYWQTJUcVfkXeA0gc1sXxd7STsMQWTNKVjiU28VTDgg3ZA7uejxoPV3PLEoMi
# aZ0IKRBK3sfd7TRZW/lN3TqP7pIbnDdkUY9J9m6lXY+2ziLcCt4HAximVTUXQ2X9
# Zky0Orm0G7mHAYVXIU9wkhglcjsvFlr5BHXOXUYZZFt7S3U=
# SIG # End signature block
