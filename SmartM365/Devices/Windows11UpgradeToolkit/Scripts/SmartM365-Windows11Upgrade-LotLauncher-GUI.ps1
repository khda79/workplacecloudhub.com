<#
.SYNOPSIS
Starts the Windows 11 Upgrade LOT launcher GUI.

.VERSION
0.1.47
#>
param(
    [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ToolkitRoot {
    $scriptPath = $PSCommandPath
    if (-not $scriptPath) {
        $scriptPath = $MyInvocation.MyCommand.Path
    }

    $scriptsRoot = Split-Path -Parent $scriptPath
    return Split-Path -Parent $scriptsRoot
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
        if ($name -notmatch '^W11UT_') {
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
    $lines.Add('# SmartM365 Windows 11 Upgrade Toolkit GUI options.')
    $lines.Add('# This file is local and is updated by the LOT launcher GUI.')
    $lines.Add('')
    foreach ($name in @($Values.Keys | Sort-Object)) {
        if ($name -notmatch '^W11UT_') { continue }
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

function Get-LotConfigPath {
    param([string]$LotPath)
    return Join-Path $LotPath 'Windows11UpgradeToolkit.config'
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
    foreach ($folder in @('', 'Logs', 'PsExecLogs', 'Reports', 'CentralLogs', 'State')) {
        $path = if ([string]::IsNullOrWhiteSpace($folder)) { $runPath } else { Join-Path $runPath $folder }
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }

    [pscustomobject]@{
        RunPath         = $runPath
        LogRoot         = Join-Path $runPath 'PsExecLogs'
        ReportRoot      = Join-Path $runPath 'Reports'
        CentralLogRoot  = Join-Path $runPath 'CentralLogs'
        LauncherLogRoot = Join-Path $runPath 'Logs'
    }
}

function Test-LotConfigCanOverrideEnvironmentValue {
    param(
        [string]$Name,
        [hashtable]$EnvironmentVariables = @{}
    )

    $processValue = [Environment]::GetEnvironmentVariable($Name, 'Process')
    if (-not [string]::IsNullOrWhiteSpace($processValue)) {
        return $false
    }

    if (-not $EnvironmentVariables.ContainsKey($Name)) {
        return $true
    }

    $currentValue = [string]$EnvironmentVariables[$Name]
    if ([string]::IsNullOrWhiteSpace($currentValue)) {
        return $true
    }

    return $false
}

function Get-EffectiveLotEnvironment {
    param(
        [string]$LotPath,
        [hashtable]$EnvironmentVariables = @{}
    )

    $effective = @{}
    foreach ($key in @($EnvironmentVariables.Keys)) {
        $effective[$key] = [string]$EnvironmentVariables[$key]
    }

    $lotConfig = Read-ToolkitConfig -Path (Get-LotConfigPath -LotPath $LotPath)
    foreach ($key in @($lotConfig.Keys)) {
        $lotValue = [string]$lotConfig[$key]
        if ([string]::IsNullOrWhiteSpace($lotValue)) {
            continue
        }

        if (Test-LotConfigCanOverrideEnvironmentValue -Name $key -EnvironmentVariables $effective) {
            $effective[$key] = $lotValue
        }
    }

    return $effective
}

function Get-SafeLotName {
    param([string]$Name)

    $safeName = [regex]::Replace($Name.Trim(), '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw 'Enter a LOT name.'
    }

    if ($safeName -notmatch '^(?i)LOT-') {
        $safeName = "LOT-$safeName"
    }

    return $safeName
}

function Get-ComputerNamesFromFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return @()
    }

    $seen = @{}
    $computers = New-Object System.Collections.Generic.List[string]
    foreach ($line in @(Get-Content -LiteralPath $Path)) {
        $name = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith('#')) {
            continue
        }

        $key = $name.ToUpperInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $computers.Add($name)
        }
    }

    return @($computers)
}

function Get-LotFolders {
    param([string]$RootPath)

    $lotsRoot = Get-LotsRoot -RootPath $RootPath
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
        'Run-Windows11UpgradeRepairWithPsExec-Loop.cmd',
        'Run-Windows11UpgradeRepairWithPsExec-Once.cmd',
        'Run-Windows11UpgradeRepairWithPsExec-Loop-IgnoreRunGuard.cmd',
        'Run-Windows11UpgradeRepairWithPsExec-Once-IgnoreRunGuard.cmd'
    )

    $missing = @(
        foreach ($fileName in $required) {
            if (-not (Test-Path -LiteralPath (Join-Path $LotPath $fileName) -PathType Leaf)) {
                $fileName
            }
        }
    )

    [pscustomobject]@{
        Ready   = ($missing.Count -eq 0)
        Missing = $missing
    }
}

function Get-LotSummary {
    param([string]$LotPath)

    $computersPath = Join-Path $LotPath 'Computers.txt'
    $wrappers = Test-LotWrapperSet -LotPath $LotPath
    $count = @(Get-ComputerNamesFromFile -Path $computersPath).Count
    $rootPath = Get-ToolkitRootFromLotPath -LotPath $LotPath
    $lotName = Split-Path -Leaf $LotPath

    [pscustomobject]@{
        Name            = $lotName
        Path            = $LotPath
        ComputersPath   = $computersPath
        ReportsPath     = Get-LotRunsRoot -RootPath $rootPath -LotName $lotName
        ComputerCount   = $count
        UpgradeScope    = 'Windows 10 readiness, policy repair, guarded Windows 11 upgrade'
        WrappersReady   = $wrappers.Ready
        MissingWrappers = $wrappers.Missing
        Display         = ('{0} ({1} devices)' -f $lotName, $count)
    }
}

function Invoke-LotWrapperRefresh {
    param([string]$RootPath)

    $script = Join-Path $RootPath 'Scripts\SmartM365-Windows11Upgrade-Update-LotCmdWrappers.ps1'
    if (-not (Test-Path -LiteralPath $script -PathType Leaf)) {
        throw "Wrapper refresh script not found: $script"
    }

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script -ToolkitRoot $RootPath
    if ($LASTEXITCODE -ne 0) {
        throw "LOT wrapper refresh failed with exit code $LASTEXITCODE."
    }
}

function New-ToolkitLotFolder {
    param(
        [string]$RootPath,
        [string]$Name
    )

    $safeName = Get-SafeLotName -Name $Name
    $lotsRoot = Get-LotsRoot -RootPath $RootPath
    New-Item -ItemType Directory -Path $lotsRoot -Force | Out-Null

    $lotPath = Join-Path $lotsRoot $safeName
    if (Test-Path -LiteralPath $lotPath) {
        throw "LOT folder already exists: $lotPath"
    }

    New-Item -ItemType Directory -Path $lotPath -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $lotPath 'Computers.txt') -Force | Out-Null

    $lotConfigPath = Get-LotConfigPath -LotPath $lotPath
    $templateConfigPath = Join-Path $RootPath 'Windows11UpgradeToolkit.config.template'
    if (Test-Path -LiteralPath $templateConfigPath -PathType Leaf) {
        Copy-Item -LiteralPath $templateConfigPath -Destination $lotConfigPath -Force
    }
    else {
        Set-Content -LiteralPath $lotConfigPath -Encoding ASCII -Value @(
            '# SmartM365 Windows 11 Upgrade Toolkit LOT defaults.'
            'W11UT_SETUP_SOURCE='
        )
    }

    Invoke-LotWrapperRefresh -RootPath $RootPath
    [pscustomobject]@{
        LotPath       = $lotPath
        ComputersPath = Join-Path $lotPath 'Computers.txt'
    }
}

function Assert-CmdLiteralValue {
    param(
        [AllowNull()][string]$Value,
        [string]$Context = 'CMD value'
    )

    if ($null -eq $Value) { return }
    if ($Value.IndexOf([char]0) -ge 0 -or $Value -match '[\r\n"%!]' ) {
        throw ("{0} contains a character that cannot be represented safely in a generated CMD launcher. Remove quotes, percent signs, exclamation marks, or line breaks." -f $Context)
    }
}

function ConvertTo-CmdArgument {
    param([string]$Value)

    if ($null -eq $Value) {
        return '""'
    }

    Assert-CmdLiteralValue -Value $Value -Context 'CMD argument'
    if ($Value -match '^[A-Za-z0-9_:\\./=-]+$') {
        return $Value
    }

    return ('"{0}"' -f $Value)
}

function ConvertTo-CmdSetCommand {
    param(
        [string]$Name,
        [AllowNull()][string]$Value
    )

    if ($Name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
        throw "Invalid environment variable name: $Name"
    }

    Assert-CmdLiteralValue -Value $Value -Context ("Environment variable {0}" -f $Name)
    return ('set "{0}={1}"' -f $Name, $Value)
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
            if ([string]$state.Toolkit -ne 'Windows11UpgradeToolkit') { continue }
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
        [string]$NamePrefix = 'W11UT-GUI',
        [string]$WindowTitle,
        [string]$EvidenceDirectory,
        [switch]$PauseWhenDone
    )

    $launcherRoot = Join-Path $env:TEMP 'SmartM365\Windows11UpgradeToolkit\GuiLaunchers'
    New-Item -ItemType Directory -Path $launcherRoot -Force | Out-Null

    $safePrefix = [regex]::Replace($NamePrefix, '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safePrefix)) { $safePrefix = 'W11UT-GUI' }

    $launchPath = Join-Path $launcherRoot ('{0}_{1}.cmd' -f $safePrefix,(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    $launchLines = New-Object System.Collections.Generic.List[string]
    $launchLines.Add('@echo off')
    $launchLines.Add('setlocal DisableDelayedExpansion')
    $launchLines.Add('echo GUI launcher : %~f0')
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
    if (-not [string]::IsNullOrWhiteSpace($EvidenceDirectory)) {
        New-Item -ItemType Directory -Path $EvidenceDirectory -Force | Out-Null
        $evidencePath = Join-Path $EvidenceDirectory 'GuiLaunchCommand.cmd'
        Copy-Item -LiteralPath $launchPath -Destination $evidencePath -Force
        $launcherHash = (Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash
        Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'GuiLaunchCommandEvidence.txt') -Encoding UTF8 -Force -Value @(
            'SmartM365 Windows 11 Upgrade Toolkit GUI launcher evidence'
            ('GeneratedUtc={0}' -f (Get-Date).ToUniversalTime().ToString('o'))
            ('TemporaryLauncher={0}' -f $launchPath)
            ('EvidenceLauncher={0}' -f $evidencePath)
            ('WorkingDirectory={0}' -f $WorkingDirectory)
            ('SHA256={0}' -f $launcherHash)
        )
    }
    return $launchPath
}

function Start-GuiLaunchCommandFile {
    param(
        [Parameter(Mandatory = $true)][string]$LaunchCommandPath,
        [Parameter(Mandatory = $true)][string]$WorkingDirectory
    )

    Assert-CmdLiteralValue -Value $LaunchCommandPath -Context 'GUI launcher path'
    $cmdArguments = '/d /s /c ""{0}""' -f $LaunchCommandPath
    Start-Process -FilePath 'cmd.exe' -ArgumentList $cmdArguments -WorkingDirectory $WorkingDirectory -Verb RunAs | Out-Null
}

function Start-LotHtmlReportOpenWatcher {
    param(
        [Parameter(Mandatory = $true)][string]$ReportRoot,
        [int]$TimeoutSeconds = 900
    )

    $resolvedReportRoot = [System.IO.Path]::GetFullPath($ReportRoot)
    $escapedReportRoot = $resolvedReportRoot.Replace("'", "''")
    $watcherScript = @"
`$ErrorActionPreference = 'SilentlyContinue'
`$reportRoot = '$escapedReportRoot'
`$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
while ((Get-Date) -lt `$deadline) {
    `$report = Get-ChildItem -LiteralPath `$reportRoot -Filter 'PsExec_Windows11Upgrade_Summary_*.html' -File | Where-Object { `$_.Length -gt 0 } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (`$report) {
        Start-Process -FilePath `$report.FullName | Out-Null
        exit 0
    }
    Start-Sleep -Seconds 2
}
"@
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($watcherScript))
    Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',$encodedCommand) -WindowStyle Hidden | Out-Null
}
function Get-SingleComputerRunRoot {
    param([string]$ToolkitRoot)
    return Join-Path (Get-RunsRoot -RootPath $ToolkitRoot) 'SingleComputer'
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

    $safeComputer = [regex]::Replace($trimmed, '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safeComputer)) {
        $safeComputer = 'Computer'
    }

    $runPath = Join-Path (Get-SingleComputerRunRoot -ToolkitRoot $ToolkitRoot) ("{0}_{1}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $safeComputer)
    New-Item -ItemType Directory -Path $runPath -Force | Out-Null

    $computersPath = Join-Path $runPath 'Computers.txt'
    Set-Content -LiteralPath $computersPath -Value $trimmed -Encoding UTF8 -Force

    foreach ($folder in @('Logs','PsExecLogs','Reports','CentralLogs','State')) {
        New-Item -ItemType Directory -Path (Join-Path $runPath $folder) -Force | Out-Null
    }

    [pscustomobject]@{
        ComputerName   = $trimmed
        SafeComputerName = $safeComputer
        RunPath        = $runPath
        ComputersPath  = $computersPath
        LogRoot        = Join-Path $runPath 'PsExecLogs'
        ReportRoot     = Join-Path $runPath 'Reports'
        CentralLogRoot = Join-Path $runPath 'CentralLogs'
        LauncherLogRoot = Join-Path $runPath 'Logs'
    }
}

function Resolve-GuiPsExecPath {
    param([string]$ToolkitRoot)

    $scriptPsExec = Join-Path $ToolkitRoot 'Scripts\PsExec.exe'
    if (Test-Path -LiteralPath $scriptPsExec -PathType Leaf) {
        return $scriptPsExec
    }

    $systemPsExec = Join-Path $env:WINDIR 'System32\PsExec.exe'
    if (Test-Path -LiteralPath $systemPsExec -PathType Leaf) {
        return $systemPsExec
    }

    $command = Get-Command -Name 'PsExec.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    throw 'PsExec.exe not found. Place it in Scripts, System32, or PATH before launching.'
}

function Copy-AutomaticInventorySnapshotToRun {
    param(
        [AllowNull()]$InventoryContext,
        [Parameter(Mandatory = $true)][string]$RunPath
    )

    if (-not $InventoryContext) { return @() }
    $copied = New-Object System.Collections.Generic.List[string]
    foreach ($source in @(
        [pscustomobject]@{ Property = 'AdInventoryCsv'; FileName = 'DevicesAD.csv' },
        [pscustomobject]@{ Property = 'IntuneInventoryCsv'; FileName = 'DevicesIntune.csv' }
    )) {
        if (-not $InventoryContext.PSObject.Properties[$source.Property]) { continue }
        $sourcePath = [string]$InventoryContext.PSObject.Properties[$source.Property].Value
        if ([string]::IsNullOrWhiteSpace($sourcePath)) { continue }
        if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
            throw ("Automatic inventory snapshot not found: {0}" -f $sourcePath)
        }

        $destinationPath = Join-Path $RunPath $source.FileName
        Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
        $copied.Add($destinationPath)
    }
    return @($copied)
}

function Start-ToolkitLot {
    param(
        [pscustomobject]$Lot,
        [string]$Mode,
        [string[]]$AdditionalArguments = @(),
        [hashtable]$EnvironmentVariables = @{},
        [AllowNull()]$InitialInventoryContext
    )

    $wrapperName = switch ($Mode) {
        'Loop' { 'Run-Windows11UpgradeRepairWithPsExec-Loop.cmd'; break }
        'Once' { 'Run-Windows11UpgradeRepairWithPsExec-Once.cmd'; break }
        'LoopIgnoreRunGuard' { 'Run-Windows11UpgradeRepairWithPsExec-Loop-IgnoreRunGuard.cmd'; break }
        'OnceIgnoreRunGuard' { 'Run-Windows11UpgradeRepairWithPsExec-Once-IgnoreRunGuard.cmd'; break }
        default { 'Run-Windows11UpgradeRepairWithPsExec-Loop.cmd' }
    }

    $wrapperPath = Join-Path $Lot.Path $wrapperName
    if (-not (Test-Path -LiteralPath $wrapperPath -PathType Leaf)) {
        throw "LOT wrapper not found: $wrapperPath"
    }

    [void](Resolve-GuiPsExecPath -ToolkitRoot $toolkitRoot)

    $effectiveEnvironment = Get-EffectiveLotEnvironment -LotPath $Lot.Path -EnvironmentVariables $EnvironmentVariables
    $run = New-LotRunContext -RootPath $toolkitRoot -LotName $Lot.Name
    [void](Copy-AutomaticInventorySnapshotToRun -InventoryContext $InitialInventoryContext -RunPath $run.RunPath)
    if ($InitialInventoryContext -and $InitialInventoryContext.PSObject.Properties['IntuneInventoryCsv'] -and -not [string]::IsNullOrWhiteSpace([string]$InitialInventoryContext.IntuneInventoryCsv)) {
        $effectiveEnvironment['W11UT_SKIP_INTUNE_INVENTORY_REFRESH'] = '1'
    }
    $effectiveEnvironment['W11UT_RUN_DIR'] = $run.RunPath
    $effectiveEnvironment['W11UT_LOT_DIR'] = $Lot.Path

    $commands = New-Object System.Collections.Generic.List[string]
    foreach ($name in @($effectiveEnvironment.Keys | Sort-Object)) {
        $value = [string]$effectiveEnvironment[$name]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $commands.Add((ConvertTo-CmdSetCommand -Name $name -Value $value))
        }
    }

    $commandParts = New-Object System.Collections.Generic.List[string]
    $commandParts.Add('call')
    $commandParts.Add((ConvertTo-CmdArgument -Value $wrapperPath))
    $commandParts.Add('-GlobalConcurrencyLimit')
    $commandParts.Add([string]$effectiveEnvironment.W11UT_GLOBAL_CONCURRENCY_LIMIT)
    $commandParts.Add('-GlobalConcurrencyLeaseTimeoutMinutes')
    $commandParts.Add([string]$effectiveEnvironment.W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES)
    foreach ($argument in @($AdditionalArguments)) {
        if (-not [string]::IsNullOrWhiteSpace($argument)) {
            $commandParts.Add((ConvertTo-CmdArgument -Value $argument))
        }
    }

    $commands.Add(($commandParts -join ' '))
    $launchTitle = "{0} - {1} - started {2}" -f $Lot.Name,$Mode,(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $launchCommandPath = New-GuiLaunchCommandFile -WorkingDirectory $run.RunPath -Commands @($commands) -NamePrefix ($Lot.Name + '-' + $Mode) -WindowTitle $launchTitle -EvidenceDirectory $run.LauncherLogRoot
    Start-GuiLaunchCommandFile -LaunchCommandPath $launchCommandPath -WorkingDirectory $run.RunPath
    Start-LotHtmlReportOpenWatcher -ReportRoot $run.ReportRoot
}

function Resolve-OrchestratorPowerShellPath {
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

function Start-ToolkitSingleComputer {
    param(
        [string]$ToolkitRoot,
        [string]$ComputerName,
        [string]$Mode,
        [string[]]$AdditionalArguments = @(),
        [hashtable]$EnvironmentVariables = @{}
    )

    $scriptPath = Join-Path $ToolkitRoot 'Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "Windows 11 upgrade PsExec launcher not found: $scriptPath"
    }

    $run = New-SingleComputerRunContext -ToolkitRoot $ToolkitRoot -ComputerName $ComputerName
    $isDryRun = @($AdditionalArguments) -contains '-DryRun'
    $psExecPath = if ($isDryRun) { $null } else { Resolve-GuiPsExecPath -ToolkitRoot $ToolkitRoot }
    $powerShellPath = Resolve-OrchestratorPowerShellPath

    $commandParts = New-Object System.Collections.Generic.List[string]
    foreach ($item in @(
        $powerShellPath,'-NoProfile','-ExecutionPolicy','Bypass','-File', $scriptPath,
        '-ComputerListPath', $run.ComputersPath,
        '-LogRoot', $run.LogRoot,
        '-ReportRoot', $run.ReportRoot,
        '-CentralLogRoot', $run.CentralLogRoot,
        '-LauncherLogRoot', $run.LauncherLogRoot,
        '-ThrottleLimit', '1',
        '-GlobalConcurrencyLimit', '0',
        '-GlobalConcurrencyLeaseTimeoutMinutes', [string]$EnvironmentVariables.W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES
    )) {
        $commandParts.Add((ConvertTo-CmdArgument -Value $item))
    }

    if ($psExecPath) {
        $commandParts.Add('-PsExecPath')
        $commandParts.Add((ConvertTo-CmdArgument -Value $psExecPath))
    }
    if ($Mode -in @('Once','OnceIgnoreRunGuard')) { $commandParts.Add('-RunOnce') }
    if ($Mode -in @('LoopIgnoreRunGuard','OnceIgnoreRunGuard')) { $commandParts.Add('-IgnoreRunGuard') }

    foreach ($pair in @(
        @{ Name = 'W11UT_AUDIT_ONLY'; Argument = '-AuditOnly' },
        @{ Name = 'W11UT_ALLOW_POLICY_REPAIR'; Argument = '-AllowPolicyRepair' },
        @{ Name = 'W11UT_ALLOW_WU_RESET'; Argument = '-AllowWUReset' },
        @{ Name = 'W11UT_ALLOW_FORCE_UPGRADE'; Argument = '-AllowForceUpgrade' },
        @{ Name = 'W11UT_ALLOW_SETUP_UPGRADE'; Argument = '-AllowSetupUpgrade' },
        @{ Name = 'W11UT_DIRECT_SETUP_UPGRADE'; Argument = '-DirectSetupUpgrade' },
        @{ Name = 'W11UT_ALLOW_REBOOT'; Argument = '-AllowReboot' },
        @{ Name = 'W11UT_SCHEDULE_RETRY_AFTER_REBOOT'; Argument = '-ScheduleRetryAfterReboot' },
        @{ Name = 'W11UT_ALLOW_SETUP_PROFILE_REPAIR'; Argument = '-AllowSetupProfileRepair' },
        @{ Name = 'W11UT_SETUP_REBOOT_WHEN_NO_USER'; Argument = '-AllowSetupCompletionRebootWhenNoUser' },
        @{ Name = 'W11UT_SKIP_VIRTUAL_MACHINES'; Argument = '-SkipVirtualMachines' },
        @{ Name = 'W11UT_ALLOW_DISK_CLEANUP'; Argument = '-AllowDiskCleanup' },
        @{ Name = 'W11UT_ALLOW_ADVANCED_DISK_CLEANUP'; Argument = '-AllowAdvancedDiskCleanup' },
        @{ Name = 'W11UT_SKIP_SETUP_MEDIA_PRECOPY'; Argument = '-SkipSetupMediaPreCopy' },
        @{ Name = 'W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY'; Argument = '-UseTechnicianRunGuardHistory' },
        @{ Name = 'W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY'; Argument = '-IgnoreTechnicianRunGuardHistory' },
        @{ Name = 'W11UT_SKIP_INTUNE_INVENTORY_REFRESH'; Argument = '-SkipIntuneInventoryRefresh' }
    )) {
        if ([string]$EnvironmentVariables[$pair.Name] -eq '1') {
            $commandParts.Add($pair.Argument)
        }
    }

    foreach ($pair in @(
        @{ Name = 'W11UT_SETUP_SOURCE'; Argument = '-SetupSourcePath' },
        @{ Name = 'W11UT_SETUP_SOURCE_MAP'; Argument = '-SetupSourceMapPath' },
        @{ Name = 'W11UT_SETUP_EXECUTION_MODE'; Argument = '-SetupExecutionMode' },
        @{ Name = 'W11UT_SETUP_MEDIA_ID'; Argument = '-SetupMediaId' },
        @{ Name = 'W11UT_SETUP_LANGUAGE'; Argument = '-SetupLanguage' },
        @{ Name = 'W11UT_SETUP_DYNAMIC_UPDATE'; Argument = '-SetupDynamicUpdate' },
        @{ Name = 'W11UT_SETUP_SOURCE_CANDIDATE_LIMIT'; Argument = '-SetupSourceCandidateLimit' },
        @{ Name = 'W11UT_SETUP_COPY_IPG_MS'; Argument = '-SetupMediaCopyIpGapMilliseconds' },
        @{ Name = 'W11UT_SETUP_COPY_JITTER_SECONDS'; Argument = '-SetupMediaCopyJitterSeconds' },
        @{ Name = 'W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT'; Argument = '-SetupSubnetConcurrencyLimit' },
        @{ Name = 'W11UT_SETUP_SUBNET_PREFIX_LENGTH'; Argument = '-SetupSubnetPrefixLength' },
        @{ Name = 'W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES'; Argument = '-SetupSubnetConcurrencyLeaseMinutes' },
        @{ Name = 'W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT'; Argument = '-SetupSubnetConcurrencyGateRoot' },
        @{ Name = 'W11UT_AD_DOMAIN'; Argument = '-AdDomain' },
        @{ Name = 'W11UT_INTUNE_TENANT_ID'; Argument = '-IntuneTenantId' },
        @{ Name = 'W11UT_INTUNE_INVENTORY_PAGE_SIZE'; Argument = '-IntuneInventoryPageSize' },
        @{ Name = 'W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS'; Argument = '-DelayBetweenComputersSeconds' },
        @{ Name = 'W11UT_DELAY_BETWEEN_CYCLES_MINUTES'; Argument = '-DelayBetweenCyclesMinutes' },
        @{ Name = 'W11UT_PSEXEC_TIMEOUT_MINUTES'; Argument = '-PsExecTimeoutMinutes' },
        @{ Name = 'W11UT_RUN_GUARD_HOURS'; Argument = '-RunGuardHours' },
        @{ Name = 'W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS'; Argument = '-RetryAfterRebootMaxAttempts' },
        @{ Name = 'W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS'; Argument = '-RetryAfterRebootDelaySeconds' },
        @{ Name = 'W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS'; Argument = '-ForceRequiredRebootWhenUptimeOverDays' }
    )) {
        $value = [string]$EnvironmentVariables[$pair.Name]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $commandParts.Add($pair.Argument)
            $commandParts.Add((ConvertTo-CmdArgument -Value $value))
        }
    }

    for ($index = 0; $index -lt @($AdditionalArguments).Count; $index++) {
        $argument = [string]$AdditionalArguments[$index]
        if ([string]::IsNullOrWhiteSpace($argument)) { continue }
        if ($argument -ieq '-ThrottleLimit') {
            $index++
            continue
        }

        $commandParts.Add((ConvertTo-CmdArgument -Value $argument))
    }

    $launchCommandPath = New-GuiLaunchCommandFile -WorkingDirectory $run.RunPath -Commands @(($commandParts -join ' ')) -NamePrefix ('Single-' + $run.SafeComputerName) -EvidenceDirectory $run.LauncherLogRoot -PauseWhenDone
    Start-GuiLaunchCommandFile -LaunchCommandPath $launchCommandPath -WorkingDirectory $run.RunPath
    return $run
}

function Wait-UiDelay {
    param([int]$Seconds)
    if ($Seconds -gt 0) {
        Start-Sleep -Seconds $Seconds
    }
}

function Open-TextFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        New-Item -ItemType File -Path $Path -Force | Out-Null
    }

    Start-Process -FilePath 'notepad.exe' -ArgumentList (ConvertTo-CmdArgument -Value $Path) | Out-Null
}

function Open-FolderPath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    Start-Process -FilePath 'explorer.exe' -ArgumentList (ConvertTo-CmdArgument -Value $Path) | Out-Null
}

function Test-IsUncPath {
    param([AllowNull()][string]$Path)
    return (-not [string]::IsNullOrWhiteSpace($Path) -and $Path.Trim().StartsWith('\\'))
}

function Split-SetupSourceText {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    @($Value -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
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
$script:ToolkitConfigPath = Join-Path $toolkitRoot 'Windows11UpgradeToolkit.config'
$script:ToolkitConfig = Read-ToolkitConfig -Path $script:ToolkitConfigPath
$script:SetupSourcePlaceholder = '\\server\share\Windows11'
$script:SetupSourcePlaceholderActive = $false
$script:ToolkitDefaultEnvironment = @{
    W11UT_AUDIT_ONLY = '0'
    W11UT_ALLOW_POLICY_REPAIR = '1'
    W11UT_ALLOW_WU_RESET = '1'
    W11UT_ALLOW_FORCE_UPGRADE = '1'
    W11UT_ALLOW_SETUP_UPGRADE = '1'
    W11UT_DIRECT_SETUP_UPGRADE = '0'
    W11UT_ALLOW_REBOOT = '1'
    W11UT_SCHEDULE_RETRY_AFTER_REBOOT = '1'
    W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS = '3'
    W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS = '300'
    W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS = '7'
    W11UT_SETUP_REBOOT_WHEN_NO_USER = '1'
    W11UT_ALLOW_SETUP_PROFILE_REPAIR = '1'
    W11UT_SKIP_VIRTUAL_MACHINES = '1'
    W11UT_ALLOW_DISK_CLEANUP = '1'
    W11UT_ALLOW_ADVANCED_DISK_CLEANUP = '0'
    W11UT_SKIP_SETUP_MEDIA_PRECOPY = '0'
    W11UT_SETUP_SOURCE = ''
    W11UT_SETUP_SOURCE_MAP = ''
    W11UT_SETUP_SOURCE_CANDIDATE_LIMIT = '5'
    W11UT_SETUP_COPY_IPG_MS = '20'
    W11UT_SETUP_COPY_JITTER_SECONDS = '300'
    W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT = '1'
    W11UT_SETUP_SUBNET_PREFIX_LENGTH = 'Auto'
    W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES = '90'
    W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT = ''
    W11UT_AD_DOMAIN = ''
    W11UT_INTUNE_TENANT_ID = ''
    W11UT_INTUNE_INVENTORY_PAGE_SIZE = '999'
    W11UT_SKIP_INTUNE_INVENTORY_REFRESH = '1'
    W11UT_SETUP_EXECUTION_MODE = 'LocalCache'
    W11UT_SETUP_MEDIA_ID = 'Win11'
    W11UT_SETUP_LANGUAGE = 'MatchSystem'
    W11UT_SETUP_DYNAMIC_UPDATE = 'Disable'
    W11UT_THROTTLE = '10'
    W11UT_GLOBAL_CONCURRENCY_LIMIT = '15'
    W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = '0'
    W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS = '0'
    W11UT_DELAY_BETWEEN_CYCLES_MINUTES = '10'
    W11UT_PSEXEC_TIMEOUT_MINUTES = '360'
    W11UT_CANCELLATION_DRAIN_TIMEOUT_MINUTES = '15'
    W11UT_GUI_DRY_RUN = '0'
    W11UT_GUI_KEEP_CENTRAL_LOG_HISTORY = '0'
    W11UT_GUI_NO_CENTRAL_LOG_COLLECTION = '0'
    W11UT_GUI_MAX_CYCLES = '0'
    W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY = '1'
    W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY = '0'
    W11UT_RUN_GUARD_HOURS = '3'
}
$launchAllLotStartDelaySeconds = 5

if ($ValidateOnly) {
    $count = @(Get-LotFolders -RootPath $toolkitRoot).Count
    Write-Output "SmartM365 Windows 11 Upgrade Toolkit LOT Launcher GUI validation completed. Lots=$count"
    return
}

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$splash = $null
$splashModulePath = Import-SmartM365GuiSplash
if ($splashModulePath) {
    . $splashModulePath
    $splash = Start-SmartM365GuiSplash -Framework Wpf -ProductName 'Windows 11 Upgrade LOT Launcher'
}

$updateCheckModulePath = Join-Path $toolkitRoot 'SmartM365.GuiUpdateCheck.ps1'
if (Test-Path -LiteralPath $updateCheckModulePath -PathType Leaf) {
    . $updateCheckModulePath
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SmartM365 Windows 11 Upgrade LOT Launcher"
        Width="1180" Height="810" MinWidth="1060" MinHeight="720"
        WindowStartupLocation="CenterScreen"
        Background="#F5F8FB" FontFamily="Segoe UI" FontSize="12"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
    <Window.Resources>
        <SolidColorBrush x:Key="TextBrush" Color="#1F2937"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#475569"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#D8E4F0"/>
        <Style TargetType="Button">
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Background" Value="White"/>
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
                    <TextBlock Text="Windows 11 Upgrade LOT Launcher" Foreground="{StaticResource TextBrush}" FontSize="28" FontWeight="SemiBold"/>
                    <TextBlock Text="Prepare and launch guarded Windows 11 upgrade campaigns by LOT or single device." Foreground="{StaticResource MutedBrush}" Margin="0,6,0,0"/>
                </StackPanel>
                <Border x:Name="HeaderLogoLink" Grid.Column="1" BorderBrush="#D8E4F0" BorderThickness="1" CornerRadius="8" Background="#F8FBFF" Padding="12" Cursor="Hand" ToolTip="Open WorkplaceCloudHub.com">
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
                                    <TextBlock Text="Scope" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="LotScopeText" Text="-" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,4,0,0"/>
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
                                <Button x:Name="OpenLotReportsButton" Content="Reports"/>
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

            <TabItem Header="Automatic LOT">
                <Grid Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="0.9*"/>
                        <ColumnDefinition Width="1.1*"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="12" Margin="0,0,8,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Text="Build from inventory" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,6"/>
                            <TextBlock Grid.Row="1" Text="Create a LOT from explicit Windows 10 records. Windows 11 evidence always excludes the device." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,0,0,6"/>
                            <StackPanel Grid.Row="2">
                                <TextBlock Text="Inventory source"/>
                                <ComboBox x:Name="AutomaticSourceCombo" Height="28"/>
                            </StackPanel>
                            <StackPanel Grid.Row="3">
                                <TextBlock Text="Intune refresh uses delegated Microsoft Graph sign-in." Foreground="{StaticResource MutedBrush}" Margin="0,2,0,2"/>
                                <CheckBox x:Name="AutomaticForceRefreshCheck" Content="Force inventory refresh this time" Margin="0,2,0,2" ToolTip="Ignore valid root caches for the next preview only."/>
                            </StackPanel>
                            <Grid Grid.Row="4">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="1.45*"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0" Margin="0,0,6,0">
                                    <TextBlock Text="LOT name"/>
                                    <TextBox x:Name="AutomaticLotNameText" Height="28"/>
                                </StackPanel>
                                <StackPanel Grid.Column="1" Margin="6,0,6,0">
                                    <TextBlock Text="Prefix(es)"/>
                                    <TextBox x:Name="AutomaticNamePrefixText" Height="28" ToolTip="Semicolon-separated prefixes, for example FR- or FR-;BE-"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Margin="6,0,0,0">
                                    <TextBlock Text="Contains"/>
                                    <TextBox x:Name="AutomaticNameContainsText" Height="28" ToolTip="Semicolon-separated literal values, for example -A- or -A-;-P-."/>
                                </StackPanel>
                            </Grid>
                            <Grid Grid.Row="5" Margin="0,4,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <CheckBox Grid.Column="0" x:Name="AutomaticExcludeIntuneCheck" Content="Exclude devices present in Intune" Margin="0,0,6,0" ToolTip="Requires AD + Intune inventory source."/>
                                <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="6,0,0,0">
                                    <CheckBox x:Name="AutomaticExcludeStaleAdCheck" Content="Exclude stale AD" ToolTip="Unknown or invalid LastLogonTimestampUtc values are also excluded."/>
                                    <TextBox x:Name="AutomaticLastLogonDaysText" Text="45" Width="38" Margin="6,0,4,0" IsEnabled="False" ToolTip="Maximum AD LastLogon age in days."/>
                                    <TextBlock Text="days" VerticalAlignment="Center"/>
                                </StackPanel>
                            </Grid>
                            <WrapPanel Grid.Row="6" Margin="0,4,0,0">
                                <Button x:Name="AutomaticPreviewButton" Content="Refresh and preview" MinWidth="155"/>
                                <Button x:Name="AutomaticCreateButton" Content="Create" Background="#0078D4" Foreground="White" BorderBrush="#0078D4" MinWidth="155"/>
                            </WrapPanel>
                        </Grid>
                    </Border>
                    <Border Grid.Column="1" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="8,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Selection preview" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}"/>
                            <TextBox Grid.Row="1" x:Name="AutomaticSummaryText" Margin="0,12,0,10" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Text="No automatic selection has been calculated yet."/>
                            <TextBox Grid.Row="2" x:Name="AutomaticEvidencePathText" IsReadOnly="True"/>
                            <Button Grid.Row="3" x:Name="AutomaticOpenEvidenceButton" Content="Open evidence folder" HorizontalAlignment="Left" IsEnabled="False"/>
                        </Grid>
                    </Border>
                </Grid>
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
                                    <CheckBox x:Name="AllowPolicyRepairCheck" Content="Allow policy repair"/>
                                    <CheckBox x:Name="AllowWUResetCheck" Content="Allow WU reset"/>
                                    <CheckBox x:Name="AllowForceUpgradeCheck" Content="Allow force upgrade"/>
                                    <CheckBox x:Name="AllowSetupUpgradeCheck" Content="Allow setup upgrade"/>
                                    <CheckBox x:Name="AllowRebootCheck" Content="Allow reboot"/>
                                    <CheckBox x:Name="ScheduleRetryAfterRebootCheck" Content="Retry at next startup"/>
                                    <CheckBox x:Name="SetupCompletionRebootCheck" Content="Reboot after setup if no user"/>
                                    <CheckBox x:Name="AllowSetupProfileRepairCheck" Content="Repair local duplicate setup profiles"/>
                                    <CheckBox x:Name="DirectSetupUpgradeCheck" Content="Direct setup upgrade"/>
                                    <CheckBox x:Name="SkipVirtualMachinesCheck" Content="Skip virtual machines"/>
                                    <CheckBox x:Name="SkipSetupPreCopyCheck" Content="Skip setup media pre-copy"/>
                                    <CheckBox x:Name="AllowDiskCleanupCheck" Content="Allow disk cleanup"/>
                                    <CheckBox x:Name="AllowAdvancedCleanupCheck" Content="Allow advanced cleanup"/>
                                    <CheckBox x:Name="KeepCentralHistoryCheck" Content="Keep central history"/>
                                    <CheckBox x:Name="NoCentralCollectionCheck" Content="No central collection"/>
                                    <CheckBox x:Name="UseTechRunGuardHistoryCheck" Content="Use technician run guard history"/>
                                    <CheckBox x:Name="IgnoreTechRunGuardHistoryCheck" Content="Ignore technician run guard history"/>
                                </WrapPanel>
                                <TextBlock Text="Force required reboot after uptime days" Margin="0,18,0,0"/>
                                <Grid Margin="0,4,0,10" Width="220" HorizontalAlignment="Left">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="32"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="32"/>
                                    </Grid.ColumnDefinitions>
                                    <Button x:Name="ForceRequiredRebootDaysDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                    <TextBox x:Name="ForceRequiredRebootDaysText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                    <Button x:Name="ForceRequiredRebootDaysUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                </Grid>
                                <TextBlock Text="Setup source" Margin="0,18,0,0"/>
                                <TextBox x:Name="SetupSourceText"/>
                                <TextBlock Text="Setup source map"/>
                                <TextBox x:Name="SetupSourceMapText"/>
                            </StackPanel>
                            <Grid Grid.Column="1" Margin="18,0,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0" Margin="0,0,10,0">
                                    <TextBlock Text="Setup" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,12"/>
                                    <TextBlock Text="Setup execution mode"/>
                                    <ComboBox x:Name="SetupModeCombo"/>
                                    <TextBlock Text="Setup media ID"/>
                                    <TextBox x:Name="SetupMediaIdText"/>
                                    <TextBlock Text="Setup language"/>
                                    <ComboBox x:Name="SetupLanguageCombo" IsEditable="True"/>
                                    <TextBlock Text="Dynamic update"/>
                                    <ComboBox x:Name="SetupDynamicUpdateCombo"/>
                                    <TextBlock Text="Setup source candidate limit"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="SetupCandidateLimitDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="SetupCandidateLimitText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="SetupCandidateLimitUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Copy IPG ms"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="SetupCopyIpgDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="SetupCopyIpgText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="SetupCopyIpgUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Copy jitter seconds"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="SetupCopyJitterDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="SetupCopyJitterText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="SetupCopyJitterUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Subnet copy limit"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="SetupSubnetLimitDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="SetupSubnetLimitText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="SetupSubnetLimitUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Subnet prefix length"/>
                                    <ComboBox x:Name="SetupSubnetPrefixCombo" IsEditable="True" Margin="0,4,0,10"/>
                                    <TextBlock Text="Subnet lease minutes"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="SetupSubnetLeaseDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="SetupSubnetLeaseText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="SetupSubnetLeaseUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Subnet gate root"/>
                                    <TextBox x:Name="SetupSubnetGateRootText"/>
                                </StackPanel>
                                <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                    <TextBlock Text="Timing" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,12"/>
                                    <TextBlock Text="Throttle"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="ThrottleDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="ThrottleText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="ThrottleUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Delay between computers (seconds)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="ComputerDelayDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="ComputerDelayText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="ComputerDelayUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Delay between cycles (minutes)"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="CycleDelayDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="CycleDelayText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="CycleDelayUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
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
                                        <TextBox x:Name="PsExecTimeoutText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="PsExecTimeoutUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                    </Grid>
                                    <TextBlock Text="Global worker limit"/>
                                    <Grid Margin="0,4,0,10">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="32"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="32"/>
                                        </Grid.ColumnDefinitions>
                                        <Button x:Name="GlobalLimitOptionDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                        <TextBox x:Name="GlobalLimitOptionText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
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
                                        <TextBox x:Name="GlobalLeaseText" Grid.Column="1" Style="{StaticResource NumericTextBoxStyle}"/>
                                        <Button x:Name="GlobalLeaseUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
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
            <TextBlock Text="SmartM365 - Windows 11 upgrade campaigns" Foreground="{StaticResource MutedBrush}"/>
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
    'LotScopeText','LotWrappersText','LotModeCombo','GlobalLimitText','GlobalLimitDownButton','GlobalLimitUpButton','OpenLotFolderButton',
    'OpenLotComputersButton','OpenLotReportsButton','RefreshWrappersButton','LaunchLotButton',
    'ActivityText','SingleComputerText','SingleModeCombo','LaunchSingleButton','SingleRunFolderText',
    'OpenSingleRunFolderButton','NewLotNameText','CreateLotButton','NewLotComputersPathText',
    'OpenNewLotComputersButton','DryRunCheck','AuditOnlyCheck','AllowPolicyRepairCheck',
    'AllowWUResetCheck','AllowForceUpgradeCheck','AllowSetupUpgradeCheck','AllowRebootCheck',
    'ScheduleRetryAfterRebootCheck','SetupCompletionRebootCheck','ForceRequiredRebootDaysText',
    'AutomaticSourceCombo','AutomaticLotNameText','AutomaticNamePrefixText','AutomaticNameContainsText','AutomaticForceRefreshCheck',
    'AutomaticExcludeIntuneCheck','AutomaticExcludeStaleAdCheck','AutomaticLastLogonDaysText',
    'AutomaticPreviewButton','AutomaticCreateButton','AutomaticSummaryText',
    'AutomaticEvidencePathText','AutomaticOpenEvidenceButton',
    'ForceRequiredRebootDaysDownButton','ForceRequiredRebootDaysUpButton','AllowSetupProfileRepairCheck',
    'DirectSetupUpgradeCheck','SkipVirtualMachinesCheck','SkipSetupPreCopyCheck',
    'AllowDiskCleanupCheck','AllowAdvancedCleanupCheck','KeepCentralHistoryCheck',
    'NoCentralCollectionCheck','UseTechRunGuardHistoryCheck','IgnoreTechRunGuardHistoryCheck','SetupSourceText','SetupSourceMapText','SetupModeCombo',
    'SetupMediaIdText','SetupLanguageCombo','SetupDynamicUpdateCombo','SetupCandidateLimitText',
    'SetupCandidateLimitDownButton','SetupCandidateLimitUpButton','SetupCopyIpgText',
    'SetupCopyIpgDownButton','SetupCopyIpgUpButton','SetupCopyJitterText',
    'SetupCopyJitterDownButton','SetupCopyJitterUpButton','SetupSubnetLimitText',
    'SetupSubnetLimitDownButton','SetupSubnetLimitUpButton','SetupSubnetPrefixCombo',
    'SetupSubnetLeaseText',
    'SetupSubnetLeaseDownButton','SetupSubnetLeaseUpButton','SetupSubnetGateRootText','ThrottleText','ThrottleDownButton',
    'ThrottleUpButton','ComputerDelayText','ComputerDelayDownButton','ComputerDelayUpButton',
    'CycleDelayText','CycleDelayDownButton','CycleDelayUpButton','MaxCyclesText','MaxCyclesDownButton',
    'MaxCyclesUpButton','PsExecTimeoutText','PsExecTimeoutDownButton','PsExecTimeoutUpButton',
    'GlobalLimitOptionText','GlobalLimitOptionDownButton','GlobalLimitOptionUpButton','ResetDefaultsButton',
    'GlobalLeaseText','GlobalLeaseDownButton','GlobalLeaseUpButton'
) | ForEach-Object { $controls[$_] = Find-Control -Name $_ }

function Open-ExternalUrl {
    param([Parameter(Mandatory)][string]$Url)

    try {
        $psi = [System.Diagnostics.ProcessStartInfo]::new($Url)
        $psi.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    }
    catch {
        [System.Windows.MessageBox]::Show($window, "Unable to open:
$Url

$($_.Exception.Message)", 'SmartM365', 'OK', 'Warning') | Out-Null
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
$script:UpdateCheckTimer = $null
$script:AutomaticInventoryProgressState = $null
$script:AutomaticLotNameTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:AutomaticGeneratedLotName = ''

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

function Get-GuiDialogOwner {
    if ($script:AutomaticInventoryProgressState -and $script:AutomaticInventoryProgressState.Window -and $script:AutomaticInventoryProgressState.Window.IsVisible) {
        return $script:AutomaticInventoryProgressState.Window
    }
    return $window
}

function Show-GuiError {
    param([string]$Message)
    [System.Windows.MessageBox]::Show((Get-GuiDialogOwner), $Message, 'SmartM365', 'OK', 'Error') | Out-Null
}

function Show-GuiWarningYesNo {
    param(
        [string]$Message,
        [string]$Title
    )

    $result = [System.Windows.MessageBox]::Show((Get-GuiDialogOwner), $Message, $Title, 'YesNo', 'Warning')
    return ($result -eq [System.Windows.MessageBoxResult]::Yes)
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
    $Combo.Text = $Selected
    $Combo.SelectedItem = $Selected
}

function Get-ConfiguredValue {
    param([string]$Name)
    return Get-ToolkitConfigValue -Name $Name -Default ([string]$script:ToolkitDefaultEnvironment[$Name])
}


function New-AutomaticInventoryProgressState {
    $progressXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="480" Height="230" WindowStyle="None" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        Background="#F5F8FB" FontFamily="Segoe UI" FontSize="12"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
    <Border BorderBrush="#B9DDF7" BorderThickness="1" CornerRadius="10" Background="White" Padding="22">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock x:Name="ProgressTitleText" Text="Preparing automatic LOT preview" FontSize="18" FontWeight="SemiBold" Foreground="#1F2937"/>
            <TextBlock Grid.Row="1" x:Name="ProgressStageText" Text="Checking inventory caches..." Margin="0,14,0,0" FontWeight="SemiBold" Foreground="#005A9E" TextWrapping="Wrap"/>
            <TextBlock Grid.Row="2" x:Name="ProgressDetailText" Text="Please wait." Margin="0,6,0,0" Foreground="#475569" TextWrapping="Wrap" MaxHeight="58"/>
            <ProgressBar Grid.Row="3" x:Name="InventoryProgressBar" Height="8" Margin="0,18,0,0" IsIndeterminate="True" Foreground="#0078D4"/>
            <TextBlock Grid.Row="4" x:Name="ProgressElapsedText" Text="Elapsed: 0 s" Margin="0,10,0,0" Foreground="#64748B" HorizontalAlignment="Right"/>
        </Grid>
    </Border>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$progressXaml)
    $progressWindow = [Windows.Markup.XamlReader]::Load($reader)
    $progressWindow.Owner = $window
    [pscustomobject]@{
        Window = $progressWindow
        StageText = $progressWindow.FindName('ProgressStageText')
        DetailText = $progressWindow.FindName('ProgressDetailText')
        ElapsedText = $progressWindow.FindName('ProgressElapsedText')
        StartedUtc = [datetime]::UtcNow
        Result = $null
        ErrorRecord = $null
    }
}

function Update-AutomaticInventoryProgress {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$Stage,
        [string]$Detail = ''
    )

    if (-not $State.Window -or -not $State.Window.IsVisible) { return }
    $State.StageText.Text = $Stage
    $State.DetailText.Text = if ([string]::IsNullOrWhiteSpace($Detail)) { 'Please wait.' } else { $Detail }
    $elapsedSeconds = [math]::Max(0, [math]::Floor(([datetime]::UtcNow - $State.StartedUtc).TotalSeconds))
    $State.ElapsedText.Text = "Elapsed: $elapsedSeconds s"
    $State.Window.UpdateLayout()
    $State.Window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Get-AutomaticSourceSelection {
    $value = [string]$controls.AutomaticSourceCombo.SelectedItem
    if ($value -eq 'AD') { return 'AD' }
    if ($value -eq 'Intune') { return 'Intune' }
    return 'Both'
}


function Test-AutomaticFilterSourceCompatibility {
    $source = Get-AutomaticSourceSelection
    if ([bool]$controls.AutomaticExcludeIntuneCheck.IsChecked -and $source -ne 'Both') {
        throw 'Exclude devices present in Intune requires the AD + Intune inventory source.'
    }
    if ([bool]$controls.AutomaticExcludeStaleAdCheck.IsChecked -and $source -eq 'Intune') {
        throw 'Exclude stale AD requires the AD or AD + Intune inventory source.'
    }
}

function Update-AutomaticFilterControlState {
    $controls.AutomaticLastLogonDaysText.IsEnabled = [bool]$controls.AutomaticExcludeStaleAdCheck.IsChecked
}

function Set-AutomaticFilterPreviewStale {
    param([string]$Message = 'Selection filter changed. Refresh the preview before creating the LOT.')
    $script:AutomaticPreviewSignature = ''
    $controls.AutomaticSummaryText.Text = $Message
}

function Get-AutomaticInventoryFileInfo {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][double]$FreshnessHours,
        [Parameter(Mandatory = $true)][string]$SourceName,
        [string]$ExpectedTenantId = '',
        [switch]$RequireDelegatedAuthentication
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ Source = $SourceName; Path = $Path; Exists = $false; Fresh = $false; AgeHours = [double]::PositiveInfinity; TenantId = ''; AuthenticationMode = ''; Scope = ''; ContentVerified = $false; Detail = 'Cache missing' }
    }

    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $ageHours = ((Get-Date) - $item.LastWriteTime).TotalHours
    $tenantId = ''
    $authenticationMode = ''
    $scope = ''
    $provenanceVerified = $true
    $contentVerified = $true
    $scopeDetail = ''
    $contentDetail = ''
    $firstRow = $null
    try {
        $firstRow = Import-Csv -LiteralPath $Path | Select-Object -First 1
        if (-not $firstRow) {
            $contentVerified = $false
            $contentDetail = '; cache content invalid (no data rows)'
        }
        else {
            $propertyNames = @($firstRow.PSObject.Properties.Name)
            $identityColumns = if ($SourceName -eq 'AD') { @('ComputerName','Name','DNSHostName') } else { @('DeviceName','ManagedDeviceName','ComputerName') }
            $hasIdentity = @($identityColumns | Where-Object { $propertyNames -contains $_ }).Count -gt 0
            $hasOperatingSystem = ($propertyNames -contains 'OperatingSystem')
            $hasOsVersion = if ($SourceName -eq 'Intune') { ($propertyNames -contains 'OSVersion' -or $propertyNames -contains 'OperatingSystemVersion') } else { $true }
            $contentVerified = ($hasIdentity -and $hasOperatingSystem -and $hasOsVersion)
            if (-not $contentVerified) {
                $contentDetail = "; cache content invalid (identity=$hasIdentity; operatingSystem=$hasOperatingSystem; osVersion=$hasOsVersion)"
            }
        }
    }
    catch {
        $contentVerified = $false
        $contentDetail = "; cache content unreadable ($($_.Exception.Message))"
    }

    if ($RequireDelegatedAuthentication) {
        $tenantId = if ($firstRow -and $firstRow.PSObject.Properties['InventoryTenantId']) { [string]$firstRow.InventoryTenantId } else { '' }
        $authenticationMode = if ($firstRow -and $firstRow.PSObject.Properties['InventoryAuthenticationMode']) { [string]$firstRow.InventoryAuthenticationMode } else { '' }
        $scope = if ($firstRow -and $firstRow.PSObject.Properties['InventoryScope']) { [string]$firstRow.InventoryScope } else { '' }
        $tenantMatches = (-not [string]::IsNullOrWhiteSpace($tenantId) -and ([string]::IsNullOrWhiteSpace($ExpectedTenantId) -or $tenantId -ieq $ExpectedTenantId))
        $authenticationMatches = ($authenticationMode -in @('DelegatedInteractive','DelegatedExistingSession'))
        $provenanceVerified = ($tenantMatches -and $authenticationMatches -and $scope -eq 'AllManagedDevices')
        $scopeDetail = if ($provenanceVerified) {
            "; tenant=$tenantId; auth=$authenticationMode; scope=$scope"
        }
        else {
            "; cache provenance mismatch (tenant='$tenantId'; auth='$authenticationMode'; scope='$scope'; expected tenant='$ExpectedTenantId'; expected auth='Delegated'; expected scope='AllManagedDevices')"
        }
    }

    [pscustomobject]@{
        Source = $SourceName
        Path = $Path
        Exists = $true
        Fresh = ($ageHours -le $FreshnessHours -and $contentVerified -and $provenanceVerified)
        AgeHours = $ageHours
        TenantId = $tenantId
        AuthenticationMode = $authenticationMode
        Scope = $scope
        ContentVerified = $contentVerified
        Detail = ('Cache age {0:N1}h; TTL {1:N0}h{2}{3}' -f $ageHours, $FreshnessHours, $contentDetail, $scopeDetail)
    }
}

function Get-GuiPowerShellProcessLogDetail {
    param(
        [string[]]$Paths,
        [int]$TailLines = 20
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($path in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $lines = @(Get-Content -LiteralPath $path -Tail $TailLines -ErrorAction SilentlyContinue)
        $content = ([string]($lines -join [Environment]::NewLine)).Trim()
        if (-not [string]::IsNullOrWhiteSpace($content)) {
            $parts.Add($content)
        }
    }
    return ($parts.ToArray() -join [Environment]::NewLine)
}

function Invoke-GuiPowerShellProcess {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$Activity,
        [switch]$Interactive,
        [scriptblock]$ProgressCallback
    )

    if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) { throw "Script not found: $ScriptPath" }
    Add-Status -Title 'Inventory refresh' -Message $Activity
    $previewWasEnabled = $controls.AutomaticPreviewButton.IsEnabled
    $createWasEnabled = $controls.AutomaticCreateButton.IsEnabled
    $controls.AutomaticPreviewButton.IsEnabled = $false
    $controls.AutomaticCreateButton.IsEnabled = $false
    $window.Cursor = [System.Windows.Input.Cursors]::Wait

    $stderrPath = "$LogPath.stderr.txt"
    $argumentParts = New-Object System.Collections.Generic.List[string]
    foreach ($value in @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath) + @($Arguments)) {
        $argumentParts.Add((ConvertTo-CmdArgument -Value ([string]$value)))
    }

    try {
        if ($ProgressCallback) { & $ProgressCallback $Activity 'Starting PowerShell inventory process...' }
        $startParameters = @{
            FilePath = 'powershell.exe'
            ArgumentList = ($argumentParts -join ' ')
            PassThru = $true
            RedirectStandardOutput = $LogPath
            RedirectStandardError = $stderrPath
        }
        if (-not $Interactive) { $startParameters.WindowStyle = 'Hidden' }
        $process = Start-Process @startParameters
        $nextProgressUpdateUtc = [datetime]::MinValue
        while (-not $process.HasExited) {
            $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
            if ($ProgressCallback -and [datetime]::UtcNow -ge $nextProgressUpdateUtc) {
                $liveDetail = Get-GuiPowerShellProcessLogDetail -Paths @($LogPath, $stderrPath) -TailLines 3
                & $ProgressCallback $Activity $liveDetail
                $nextProgressUpdateUtc = [datetime]::UtcNow.AddSeconds(1)
            }
            Start-Sleep -Milliseconds 200
            $process.Refresh()
        }
        $process.Refresh()
        if ($process.ExitCode -ne 0) {
            $detail = Get-GuiPowerShellProcessLogDetail -Paths @($stderrPath, $LogPath) -TailLines 20
            if ([string]::IsNullOrWhiteSpace($detail)) { $detail = 'The inventory process produced no error text.' }
            throw ("Inventory refresh failed with exit code {0}. Stdout={1}; Stderr={2}.{3}{4}" -f $process.ExitCode, $LogPath, $stderrPath, [Environment]::NewLine, $detail)
        }
        if ($ProgressCallback) {
            $detail = Get-GuiPowerShellProcessLogDetail -Paths @($LogPath) -TailLines 3
            & $ProgressCallback $Activity $detail
        }
    }
    finally {
        $window.Cursor = $null
        $controls.AutomaticPreviewButton.IsEnabled = $previewWasEnabled
        $controls.AutomaticCreateButton.IsEnabled = $createWasEnabled
    }
}

function Get-AutomaticInventorySnapshot {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('AD', 'Intune', 'Both')][string]$Source,
        [switch]$ForceRefresh,
        [scriptblock]$ProgressCallback
    )

    $automaticRoot = Join-Path (Get-RunsRoot -RootPath $toolkitRoot) 'AutomaticLotInventory'
    $sourceRunPath = Join-Path $automaticRoot ('Sources-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmssfff'))
    New-Item -ItemType Directory -Path $sourceRunPath -Force | Out-Null

    $requestedSources = if ($Source -eq 'Both') { @('AD', 'Intune') } else { @($Source) }
    $paths = @{ AD = ''; Intune = '' }
    $details = New-Object System.Collections.Generic.List[string]
    $failures = New-Object System.Collections.Generic.List[string]
    $tenantId = ''
    $authenticationMode = ''
    $configuredTenantId = Get-ConfiguredValue 'W11UT_INTUNE_TENANT_ID'

    if ($ProgressCallback) { & $ProgressCallback 'Checking root inventory caches...' "Requested source: $Source" }

    foreach ($requestedSource in $requestedSources) {
        $rootCsv = Join-Path $toolkitRoot $(if ($requestedSource -eq 'AD') { 'DevicesAD.csv' } else { 'DevicesIntune.csv' })
        $outputCsv = Join-Path $sourceRunPath $(if ($requestedSource -eq 'AD') { 'DevicesAD.csv' } else { 'DevicesIntune.csv' })
        try {
            $rootInfo = if ($requestedSource -eq 'AD') {
                Get-AutomaticInventoryFileInfo -Path $rootCsv -FreshnessHours 12 -SourceName 'AD'
            }
            else {
                Get-AutomaticInventoryFileInfo -Path $rootCsv -FreshnessHours 2 -SourceName 'Intune' -ExpectedTenantId $configuredTenantId -RequireDelegatedAuthentication
            }
        }
        catch {
            $rootInfo = [pscustomobject]@{
                Source = $requestedSource
                Path = $rootCsv
                Exists = (Test-Path -LiteralPath $rootCsv -PathType Leaf)
                Fresh = $false
                AgeHours = [double]::PositiveInfinity
                TenantId = ''
                AuthenticationMode = ''
                Scope = ''
                Detail = "Cache validation failed: $($_.Exception.Message)"
            }
        }

        if (-not $ForceRefresh -and $rootInfo.Fresh) {
            try {
                if ($ProgressCallback) { & $ProgressCallback "Reusing recent $requestedSource inventory..." ("{0}; copying root cache to {1}" -f $rootInfo.Detail, $outputCsv) }
                Copy-Item -LiteralPath $rootCsv -Destination $outputCsv -Force -ErrorAction Stop
                $paths[$requestedSource] = $outputCsv
                if ($requestedSource -eq 'Intune') {
                    $tenantId = [string]$rootInfo.TenantId
                    $authenticationMode = [string]$rootInfo.AuthenticationMode
                }
                $details.Add(("{0}: reused verified recent root cache; {1}; snapshot={2}" -f $requestedSource, $rootInfo.Detail, $outputCsv))
                continue
            }
            catch {
                $details.Add(("{0}: recent root cache copy failed; refresh required; detail={1}" -f $requestedSource, $_.Exception.Message))
            }
        }
        elseif ($ForceRefresh -and $rootInfo.Fresh) {
            $details.Add(("{0}: force refresh requested; verified root cache retained as fallback; {1}" -f $requestedSource, $rootInfo.Detail))
        }
        else {
            $details.Add(("{0}: root cache not reusable; {1}" -f $requestedSource, $rootInfo.Detail))
        }

        try {
            if ($requestedSource -eq 'AD') {
                $logPath = Join-Path $sourceRunPath 'DevicesAD.refresh.log'
                $exporter = Join-Path $toolkitRoot 'Scripts\SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1'
                $invokeParameters = @{
                    ScriptPath = $exporter
                    Arguments = @('-OutputPath', $outputCsv, '-ForceRefresh')
                    LogPath = $logPath
                    Activity = 'Reading a fresh complete AD forest inventory for the automatic LOT...'
                }
                if ($ProgressCallback) { $invokeParameters.ProgressCallback = $ProgressCallback }
                Invoke-GuiPowerShellProcess @invokeParameters
                if (-not (Test-Path -LiteralPath $outputCsv -PathType Leaf)) { throw "AD inventory CSV was not created: $outputCsv" }
                $paths.AD = $outputCsv
                $details.Add("AD: generated isolated automatic snapshot $outputCsv")
            }
            else {
                $logPath = Join-Path $sourceRunPath 'DevicesIntune.refresh.log'
                $exporter = Join-Path $toolkitRoot 'Scripts\SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1'
                $arguments = @('-OutputPath', $outputCsv, '-ForceRefresh')
                if (-not [string]::IsNullOrWhiteSpace($configuredTenantId)) { $arguments += @('-TenantId', $configuredTenantId) }
                $invokeParameters = @{
                    ScriptPath = $exporter
                    Arguments = $arguments
                    LogPath = $logPath
                    Activity = 'Waiting for delegated interactive Microsoft Graph sign-in to read a fresh complete Intune inventory...'
                    Interactive = $true
                }
                if ($ProgressCallback) { $invokeParameters.ProgressCallback = $ProgressCallback }
                Invoke-GuiPowerShellProcess @invokeParameters
                if (-not (Test-Path -LiteralPath $outputCsv -PathType Leaf)) { throw "Intune inventory CSV was not created: $outputCsv" }
                $intuneInfo = Get-AutomaticInventoryFileInfo -Path $outputCsv -FreshnessHours 2 -SourceName 'Intune' -ExpectedTenantId $configuredTenantId -RequireDelegatedAuthentication
                if (-not $intuneInfo.Fresh) { throw ("Generated Intune inventory provenance is invalid: {0}" -f $intuneInfo.Detail) }
                $paths.Intune = $outputCsv
                $tenantId = [string]$intuneInfo.TenantId
                $authenticationMode = [string]$intuneInfo.AuthenticationMode
                $details.Add("Intune: generated delegated interactive snapshot; tenant=$tenantId; path=$outputCsv")
            }
        }
        catch {
            $refreshError = $_.Exception.Message
            $fallbackAccepted = $false
            if ($rootInfo.Fresh) {
                $fallbackAccepted = Show-GuiWarningYesNo -Title ("{0} refresh failed" -f $requestedSource) -Message ((@(
                    ("The automatic {0} snapshot could not be generated:" -f $requestedSource)
                    $refreshError
                    ''
                    ("A verified recent root cache is available: {0}" -f $rootCsv)
                    ([string]$rootInfo.Detail)
                    ''
                    'Use this root cache explicitly as the fallback for this preview? A copy will still be preserved in the automatic evidence and run folders.'
                )) -join [Environment]::NewLine)
            }

            if ($fallbackAccepted) {
                Copy-Item -LiteralPath $rootCsv -Destination $outputCsv -Force -ErrorAction Stop
                $paths[$requestedSource] = $outputCsv
                if ($requestedSource -eq 'Intune') {
                    $tenantId = [string]$rootInfo.TenantId
                    $authenticationMode = [string]$rootInfo.AuthenticationMode
                }
                $details.Add(("{0}: root cache fallback explicitly accepted after refresh failure; {1}; snapshot={2}" -f $requestedSource, $rootInfo.Detail, $outputCsv))
                continue
            }
            $failures.Add(("{0}: {1}" -f $requestedSource, $refreshError))
        }
    }

    $availableCount = @($paths.Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count
    if ($failures.Count -gt 0) {
        if ($availableCount -eq 0) { throw ($failures -join [Environment]::NewLine) }
        $continuePartial = Show-GuiWarningYesNo -Title 'Partial inventory source' -Message ((@(
            'One selected inventory source failed:'
            ($failures -join [Environment]::NewLine)
            ''
            'Continue explicitly with the available source only? The preview and evidence will be marked partial.'
        )) -join [Environment]::NewLine)
        if (-not $continuePartial) { throw 'Automatic LOT selection cancelled because one inventory source failed.' }
    }

    if ($ProgressCallback) { & $ProgressCallback 'Inventory sources ready.' ($details.ToArray() -join [Environment]::NewLine) }
    [pscustomobject]@{
        RequestedSource = $Source
        AdInventoryCsv = [string]$paths.AD
        IntuneInventoryCsv = [string]$paths.Intune
        TenantId = $tenantId
        AuthenticationMode = $authenticationMode
        PartialSource = ($failures.Count -gt 0)
        ForceInventoryRefresh = [bool]$ForceRefresh
        SourceDetails = $details.ToArray()
        Failures = $failures.ToArray()
        SourceRunPath = $sourceRunPath
    }
}

function Invoke-AutomaticLotSelection {
    param(
        [Parameter(Mandatory = $true)]$InventoryContext,
        [switch]$Create
    )

    $engine = Join-Path $toolkitRoot 'Scripts\SmartM365-Windows11Upgrade-New-AutomaticLot.ps1'
    if (-not (Test-Path -LiteralPath $engine -PathType Leaf)) { throw "Automatic LOT engine not found: $engine" }
    $parameters = @{
        Source = [string]$InventoryContext.RequestedSource
        ToolkitRoot = $toolkitRoot
        LotName = [string]$controls.AutomaticLotNameText.Text
        EvidenceRoot = Join-Path (Get-RunsRoot -RootPath $toolkitRoot) 'AutomaticLotInventory'
    }
    $namePrefixText = [string]$controls.AutomaticNamePrefixText.Text
    if (-not [string]::IsNullOrWhiteSpace($namePrefixText)) { $parameters.ComputerNamePrefix = @($namePrefixText) }
    $nameContainsText = [string]$controls.AutomaticNameContainsText.Text
    if (-not [string]::IsNullOrWhiteSpace($nameContainsText)) { $parameters.ComputerNameContains = @($nameContainsText) }
    if ([bool]$controls.AutomaticExcludeIntuneCheck.IsChecked) { $parameters.ExcludeIntunePresent = $true }
    if ([bool]$controls.AutomaticExcludeStaleAdCheck.IsChecked) {
        $parameters.ExcludeStaleAd = $true
        $parameters.AdLastLogonMaxAgeDays = [int](Get-IntText -TextBox $controls.AutomaticLastLogonDaysText -Default 45 -Minimum 1)
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$InventoryContext.AdInventoryCsv)) { $parameters.AdInventoryCsv = [string]$InventoryContext.AdInventoryCsv }
    if (-not [string]::IsNullOrWhiteSpace([string]$InventoryContext.IntuneInventoryCsv)) { $parameters.IntuneInventoryCsv = [string]$InventoryContext.IntuneInventoryCsv }
    if ($InventoryContext.PartialSource) { $parameters.AllowPartialSource = $true }
    if ($Create) {
        $parameters.Create = $true
    }

    return & $engine @parameters
}

function Format-AutomaticLotSummary {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [Parameter(Mandatory = $true)]$InventoryContext
    )

    $summary = $Result.Summary
    @(
        "Requested source: $($summary.RequestedSource)"
        "Available source: $($summary.AvailableSources)$(if ($summary.PartialSource) { ' (PARTIAL)' } else { '' })"
        "Intune tenant: $(if ([string]::IsNullOrWhiteSpace([string]$InventoryContext.TenantId)) { 'Not used' } else { [string]$InventoryContext.TenantId })"
        "Intune auth: $(if ([string]::IsNullOrWhiteSpace([string]$InventoryContext.AuthenticationMode)) { 'Not used' } else { [string]$InventoryContext.AuthenticationMode })"
        @($InventoryContext.SourceDetails)
        "Computer prefix(es): $(if ([string]::IsNullOrWhiteSpace([string]$summary.ComputerNamePrefixes)) { 'All' } else { [string]$summary.ComputerNamePrefixes })"
        "Computer contains: $(if ([string]::IsNullOrWhiteSpace([string]$summary.ComputerNameContains)) { 'All' } else { [string]$summary.ComputerNameContains })"
        "Exclude devices present in Intune: $([bool]$summary.ExcludeIntunePresent)"
        "AD LastLogon filter: $(if ([bool]$summary.ExcludeStaleAd) { "Older than $($summary.ADLastLogonMaxAgeDays) day(s); unknown values excluded" } else { 'Disabled' })"
        "Unique inventory devices: $($summary.UniqueInventoryDevices)"
        "Matched by all filters: $($summary.NameFilterMatchedDevices)"
        "Filtered out: $($summary.NameFilterExcludedDevices)"
        "Filter exclusions: prefix=$($summary.PrefixFilterExcluded); contains=$($summary.ContainsFilterExcluded); Intune present=$($summary.IntunePresentFilterExcluded); AD LastLogon=$($summary.ADLastLogonFilterExcluded) (unknown=$($summary.ADLastLogonUnknownExcluded))"
        ''
        "AD rows: $($summary.ADRows); matching Windows 10 candidates: $($summary.ADWindows10Candidates)"
        "Intune rows: $($summary.IntuneRows); matching Windows 10 candidates: $($summary.IntuneWindows10Candidates)"
        "Selected unique devices: $($summary.SelectedDevices)"
        "Excluded devices: $($summary.ExcludedDevices)"
        "Windows 11 excluded: $($summary.Windows11Excluded)"
        "AD disabled excluded: $($summary.ADDisabledExcluded)"
        "Intune retired/wipe/delete excluded: $($summary.IntuneStateExcluded)"
        "Unknown OS excluded: $($summary.UnknownOSExcluded)"
        "AD name collisions excluded: $($summary.ADNameCollisions)"
        "Ignored duplicate Intune rows: $($summary.IntuneDuplicateRowsIgnored)"
        "Stale warnings: AD=$($summary.ADStaleWarnings); Intune=$($summary.IntuneStaleWarnings)"
        ''
        "LOT name: $($summary.LotName)"
        "Evidence: $($summary.EvidencePath)"
    ) -join [Environment]::NewLine
}

function Get-AutomaticPreviewSignature {
    return @(
        (Get-AutomaticSourceSelection)
        (Get-ConfiguredValue 'W11UT_INTUNE_TENANT_ID')
        [string]$controls.AutomaticLotNameText.Text.Trim()
        [string]$controls.AutomaticNamePrefixText.Text.Trim()
        [string]$controls.AutomaticNameContainsText.Text.Trim()
        [bool]$controls.AutomaticExcludeIntuneCheck.IsChecked
        [bool]$controls.AutomaticExcludeStaleAdCheck.IsChecked
        [string]$controls.AutomaticLastLogonDaysText.Text.Trim()
    ) -join '|'
}

function Get-AutomaticGeneratedLotName {
    param([AllowNull()][string]$PrefixText)

    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @(([string]$PrefixText) -split ';')) {
        $segment = [regex]::Replace($candidate.Trim().Trim([char]34).ToUpperInvariant(), '[^A-Z0-9_-]+', '-').Trim('-_')
        if (-not [string]::IsNullOrWhiteSpace($segment) -and -not $segments.Contains($segment)) {
            $segments.Add($segment)
        }
    }

    $prefixSegment = if ($segments.Count -gt 0) { "-$($segments -join '-')" } else { '' }
    return 'LOT-AUTO-W10{0}-{1}' -f $prefixSegment, $script:AutomaticLotNameTimestamp
}

function Update-AutomaticGeneratedLotName {
    $currentName = [string]$controls.AutomaticLotNameText.Text
    if (
        -not [string]::IsNullOrWhiteSpace($currentName) -and
        $currentName -cne [string]$script:AutomaticGeneratedLotName
    ) {
        return
    }

    $generatedName = Get-AutomaticGeneratedLotName -PrefixText ([string]$controls.AutomaticNamePrefixText.Text)
    $script:AutomaticGeneratedLotName = $generatedName
    $controls.AutomaticLotNameText.Text = $generatedName
}

function Update-AutomaticLotPreview {
    param(
        [switch]$ForceInventoryRefresh,
        [scriptblock]$ProgressCallback
    )

    $source = Get-AutomaticSourceSelection
    Test-AutomaticFilterSourceCompatibility
    $snapshotParameters = @{ Source = $source; ForceRefresh = [bool]$ForceInventoryRefresh }
    if ($ProgressCallback) { $snapshotParameters.ProgressCallback = $ProgressCallback }
    $context = Get-AutomaticInventorySnapshot @snapshotParameters
    if ($ProgressCallback) { & $ProgressCallback 'Loading inventories and building the selection preview...' "Source snapshots: $($context.SourceRunPath)" }
    $result = Invoke-AutomaticLotSelection -InventoryContext $context
    $controls.AutomaticLotNameText.Text = [string]$result.Summary.LotName
    $controls.AutomaticSummaryText.Text = Format-AutomaticLotSummary -Result $result -InventoryContext $context
    $controls.AutomaticEvidencePathText.Text = [string]$result.Summary.EvidencePath
    $controls.AutomaticOpenEvidenceButton.IsEnabled = -not [string]::IsNullOrWhiteSpace([string]$result.Summary.EvidencePath)
    $script:AutomaticPreviewContext = $context
    $script:AutomaticPreviewResult = $result
    $script:AutomaticPreviewSignature = Get-AutomaticPreviewSignature
    if ($ProgressCallback) { & $ProgressCallback 'Automatic LOT preview ready.' ("Selected {0} Windows 10 device(s); safety exclusions {1}; filtered out {2}." -f $result.Summary.SelectedDevices, $result.Summary.ExcludedDevices, $result.Summary.NameFilterExcludedDevices) }
    Add-Status -Title 'Automatic preview' -Message ("Selected {0} unique Windows 10 device(s); excluded {1}." -f $result.Summary.SelectedDevices, $result.Summary.ExcludedDevices)
    return $result
}

function New-AutomaticLotPreviewWork {
    param(
        [Parameter(Mandatory = $true)]$ProgressState,
        [Parameter(Mandatory = $true)][scriptblock]$Operation
    )

    return {
        try {
            $ProgressState.Result = & $Operation
        }
        catch {
            $ProgressState.ErrorRecord = $_
        }
        finally {
            if ($ProgressState.Window -and $ProgressState.Window.IsVisible) {
                $ProgressState.Window.Close()
            }
        }
    }.GetNewClosure()
}

function Invoke-AutomaticLotPreviewWithProgress {
    param([switch]$ForceInventoryRefresh)

    $state = New-AutomaticInventoryProgressState
    $script:AutomaticInventoryProgressState = $state
    $progressCallback = {
        param([string]$Stage, [string]$Detail)
        Update-AutomaticInventoryProgress -State $state -Stage $Stage -Detail $Detail
    }.GetNewClosure()

    $operation = {
        Update-AutomaticLotPreview -ForceInventoryRefresh:$ForceInventoryRefresh -ProgressCallback $progressCallback
    }.GetNewClosure()
    $work = New-AutomaticLotPreviewWork -ProgressState $state -Operation $operation
    $contentRendered = {
        [void]$state.Window.Dispatcher.BeginInvoke([action]$work, [System.Windows.Threading.DispatcherPriority]::Background)
    }.GetNewClosure()
    $state.Window.Add_ContentRendered($contentRendered)

    try {
        [void]$state.Window.ShowDialog()
    }
    finally {
        $script:AutomaticInventoryProgressState = $null
        if ($ForceInventoryRefresh) { $controls.AutomaticForceRefreshCheck.IsChecked = $false }
    }

    if ($state.ErrorRecord) { throw $state.ErrorRecord }
    return $state.Result
}

function Confirm-AutomaticLotCreate {
    param([Parameter(Mandatory = $true)]$Result)

    $summary = $Result.Summary
    $message = @(
        "Create $($summary.LotName)?"
        ''
        "Inventory source: $($summary.AvailableSources)$(if ($summary.PartialSource) { ' (PARTIAL)' } else { '' })"
        "Intune tenant: $(if ($summary.RequestedSource -eq 'AD') { 'Not used' } else { [string]$script:AutomaticPreviewContext.TenantId })"
        "Intune auth: $(if ($summary.RequestedSource -eq 'AD') { 'Not used' } else { [string]$script:AutomaticPreviewContext.AuthenticationMode })"
        "Computer prefix(es): $(if ([string]::IsNullOrWhiteSpace([string]$summary.ComputerNamePrefixes)) { 'All' } else { [string]$summary.ComputerNamePrefixes })"
        "Computer contains: $(if ([string]::IsNullOrWhiteSpace([string]$summary.ComputerNameContains)) { 'All' } else { [string]$summary.ComputerNameContains })"
        "Exclude devices present in Intune: $([bool]$summary.ExcludeIntunePresent)"
        "AD LastLogon filter: $(if ([bool]$summary.ExcludeStaleAd) { "Older than $($summary.ADLastLogonMaxAgeDays) day(s); unknown values excluded" } else { 'Disabled' })"
        "Matched by all filters: $($summary.NameFilterMatchedDevices)"
        "Filtered out: $($summary.NameFilterExcludedDevices)"
        "Selected Windows 10 devices: $($summary.SelectedDevices)"
        "Excluded devices: $($summary.ExcludedDevices)"
        "Stale warnings: AD=$($summary.ADStaleWarnings); Intune=$($summary.IntuneStaleWarnings)"
        ''
        'The LOT will be created but not launched. Open Existing LOT when you are ready to launch it.'
    ) -join [Environment]::NewLine
    return Show-GuiWarningYesNo -Title 'Create automatic LOT' -Message $message
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
    Register-NumericStepper -TextBoxName 'ForceRequiredRebootDaysText' -DownButtonName 'ForceRequiredRebootDaysDownButton' -UpButtonName 'ForceRequiredRebootDaysUpButton' -Default 7 -Minimum 0
    Register-NumericStepper -TextBoxName 'SetupCandidateLimitText' -DownButtonName 'SetupCandidateLimitDownButton' -UpButtonName 'SetupCandidateLimitUpButton' -Default 5 -Minimum 0
    Register-NumericStepper -TextBoxName 'SetupCopyIpgText' -DownButtonName 'SetupCopyIpgDownButton' -UpButtonName 'SetupCopyIpgUpButton' -Default 20 -Minimum 0 -Step 5
    Register-NumericStepper -TextBoxName 'SetupCopyJitterText' -DownButtonName 'SetupCopyJitterDownButton' -UpButtonName 'SetupCopyJitterUpButton' -Default 300 -Minimum 0 -Step 30
    Register-NumericStepper -TextBoxName 'SetupSubnetLimitText' -DownButtonName 'SetupSubnetLimitDownButton' -UpButtonName 'SetupSubnetLimitUpButton' -Default 1 -Minimum 0
    Register-NumericStepper -TextBoxName 'SetupSubnetLeaseText' -DownButtonName 'SetupSubnetLeaseDownButton' -UpButtonName 'SetupSubnetLeaseUpButton' -Default 90 -Minimum 1 -Step 5
    Register-NumericStepper -TextBoxName 'ThrottleText' -DownButtonName 'ThrottleDownButton' -UpButtonName 'ThrottleUpButton' -Default 10 -Minimum 1
    Register-NumericStepper -TextBoxName 'ComputerDelayText' -DownButtonName 'ComputerDelayDownButton' -UpButtonName 'ComputerDelayUpButton' -Default 0 -Minimum 0 -Step 5
    Register-NumericStepper -TextBoxName 'CycleDelayText' -DownButtonName 'CycleDelayDownButton' -UpButtonName 'CycleDelayUpButton' -Default 10 -Minimum 0
    Register-NumericStepper -TextBoxName 'MaxCyclesText' -DownButtonName 'MaxCyclesDownButton' -UpButtonName 'MaxCyclesUpButton' -Default 0 -Minimum 0
    Register-NumericStepper -TextBoxName 'PsExecTimeoutText' -DownButtonName 'PsExecTimeoutDownButton' -UpButtonName 'PsExecTimeoutUpButton' -Default 360 -Minimum 1 -Step 5
    Register-NumericStepper -TextBoxName 'GlobalLimitOptionText' -DownButtonName 'GlobalLimitOptionDownButton' -UpButtonName 'GlobalLimitOptionUpButton' -Default 15 -Minimum 1
    Register-NumericStepper -TextBoxName 'GlobalLeaseText' -DownButtonName 'GlobalLeaseDownButton' -UpButtonName 'GlobalLeaseUpButton' -Default 0 -Minimum 0
}

function Set-SetupSourceTextForeground {
    param([bool]$IsPlaceholder)

    if ($IsPlaceholder) {
        $controls.SetupSourceText.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(100, 116, 139))
        return
    }

    $controls.SetupSourceText.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(31, 41, 55))
}

function Set-SetupSourceTextValue {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        $script:SetupSourcePlaceholderActive = $true
        Set-SetupSourceTextForeground -IsPlaceholder $true
        $controls.SetupSourceText.Text = $script:SetupSourcePlaceholder
        return
    }

    $script:SetupSourcePlaceholderActive = $false
    Set-SetupSourceTextForeground -IsPlaceholder $false
    $controls.SetupSourceText.Text = $Value
}

function Clear-SetupSourcePlaceholder {
    if (-not $script:SetupSourcePlaceholderActive) { return }

    $script:SetupSourcePlaceholderActive = $false
    Set-SetupSourceTextForeground -IsPlaceholder $false
    $controls.SetupSourceText.Clear()
}

function Restore-SetupSourcePlaceholder {
    if (-not [string]::IsNullOrWhiteSpace($controls.SetupSourceText.Text)) { return }
    Set-SetupSourceTextValue -Value ''
}

function Get-SetupSourceText {
    if ($script:SetupSourcePlaceholderActive) { return '' }

    $value = $controls.SetupSourceText.Text.Trim()
    if ($value -eq $script:SetupSourcePlaceholder) { return '' }
    return $value
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
        $environment = Get-ToolkitOptionEnvironment
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

function Initialize-Options {
    Initialize-Combo -Combo $controls.LotModeCombo -Values @('Loop','Once','LoopIgnoreRunGuard','OnceIgnoreRunGuard') -Selected 'Loop'
    Initialize-Combo -Combo $controls.SingleModeCombo -Values @('Once','OnceIgnoreRunGuard','Loop','LoopIgnoreRunGuard') -Selected 'Once'
    Initialize-Combo -Combo $controls.AutomaticSourceCombo -Values @('AD + Intune','AD','Intune') -Selected 'AD + Intune'
    $controls.AutomaticExcludeIntuneCheck.IsChecked = $false
    $controls.AutomaticExcludeStaleAdCheck.IsChecked = $false
    if ([string]::IsNullOrWhiteSpace([string]$controls.AutomaticLastLogonDaysText.Text)) { $controls.AutomaticLastLogonDaysText.Text = '45' }
    Update-AutomaticFilterControlState
    Update-AutomaticGeneratedLotName
    Initialize-Combo -Combo $controls.SetupModeCombo -Values @('LocalCache','Share','Auto') -Selected (Get-ConfiguredValue 'W11UT_SETUP_EXECUTION_MODE')
    Initialize-Combo -Combo $controls.SetupLanguageCombo -Values @('MatchSystem','Any','fr-FR','en-GB','en-US','de-DE','es-ES','it-IT','nl-NL','pt-PT','pl-PL') -Selected (Get-ConfiguredValue 'W11UT_SETUP_LANGUAGE')
    Initialize-Combo -Combo $controls.SetupDynamicUpdateCombo -Values @('Disable','Enable','NoDrivers','NoLCU','NoDriversNoLCU') -Selected (Get-ConfiguredValue 'W11UT_SETUP_DYNAMIC_UPDATE')
    Initialize-Combo -Combo $controls.SetupSubnetPrefixCombo -Values @('Auto','20','21','22','23','24','25','26','27','28','29','30','31','32') -Selected (Get-ConfiguredValue 'W11UT_SETUP_SUBNET_PREFIX_LENGTH')

    $controls.AuditOnlyCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_AUDIT_ONLY') -eq '1')
    $controls.AllowPolicyRepairCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_ALLOW_POLICY_REPAIR') -eq '1')
    $controls.AllowWUResetCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_ALLOW_WU_RESET') -eq '1')
    $controls.AllowForceUpgradeCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_ALLOW_FORCE_UPGRADE') -eq '1')
    $controls.AllowSetupUpgradeCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_ALLOW_SETUP_UPGRADE') -eq '1')
    $controls.AllowRebootCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_ALLOW_REBOOT') -ne '0')
    $controls.ScheduleRetryAfterRebootCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_SCHEDULE_RETRY_AFTER_REBOOT') -ne '0')
    $controls.SetupCompletionRebootCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_SETUP_REBOOT_WHEN_NO_USER') -ne '0')
    $controls.ForceRequiredRebootDaysText.Text = Get-ConfiguredValue 'W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS'
    $controls.AllowSetupProfileRepairCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_ALLOW_SETUP_PROFILE_REPAIR') -ne '0')
    $controls.DirectSetupUpgradeCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_DIRECT_SETUP_UPGRADE') -eq '1')
    $controls.SkipVirtualMachinesCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_SKIP_VIRTUAL_MACHINES') -eq '1')
    $controls.SkipSetupPreCopyCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_SKIP_SETUP_MEDIA_PRECOPY') -eq '1')
    $controls.AllowDiskCleanupCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_ALLOW_DISK_CLEANUP') -eq '1')
    $controls.AllowAdvancedCleanupCheck.IsChecked = (
        (Get-ConfiguredValue 'W11UT_ALLOW_ADVANCED_DISK_CLEANUP') -eq '1' -or
        (Get-ConfiguredValue 'W11UT_ALLOW_DISM_COMPONENT_CLEANUP') -eq '1'
    )

    Set-SetupSourceTextValue -Value (Get-ConfiguredValue 'W11UT_SETUP_SOURCE')
    $controls.SetupSourceMapText.Text = Get-ConfiguredValue 'W11UT_SETUP_SOURCE_MAP'
    $controls.SetupMediaIdText.Text = Get-ConfiguredValue 'W11UT_SETUP_MEDIA_ID'
    $controls.SetupCandidateLimitText.Text = Get-ConfiguredValue 'W11UT_SETUP_SOURCE_CANDIDATE_LIMIT'
    $controls.SetupCopyIpgText.Text = Get-ConfiguredValue 'W11UT_SETUP_COPY_IPG_MS'
    $controls.SetupCopyJitterText.Text = Get-ConfiguredValue 'W11UT_SETUP_COPY_JITTER_SECONDS'
    $controls.SetupSubnetLimitText.Text = Get-ConfiguredValue 'W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT'
    $controls.SetupSubnetLeaseText.Text = Get-ConfiguredValue 'W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES'
    $controls.SetupSubnetGateRootText.Text = Get-ConfiguredValue 'W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT'
    $controls.ThrottleText.Text = Get-ConfiguredValue 'W11UT_THROTTLE'
    $controls.ComputerDelayText.Text = Get-ConfiguredValue 'W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS'
    $controls.CycleDelayText.Text = Get-ConfiguredValue 'W11UT_DELAY_BETWEEN_CYCLES_MINUTES'
    $controls.PsExecTimeoutText.Text = Get-ConfiguredValue 'W11UT_PSEXEC_TIMEOUT_MINUTES'
    $controls.GlobalLimitText.Text = Get-ConfiguredValue 'W11UT_GLOBAL_CONCURRENCY_LIMIT'
    $controls.GlobalLimitOptionText.Text = Get-ConfiguredValue 'W11UT_GLOBAL_CONCURRENCY_LIMIT'
    $controls.GlobalLeaseText.Text = Get-ConfiguredValue 'W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES'
    $controls.DryRunCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_GUI_DRY_RUN') -eq '1')
    $controls.KeepCentralHistoryCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_GUI_KEEP_CENTRAL_LOG_HISTORY') -eq '1')
    $controls.NoCentralCollectionCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_GUI_NO_CENTRAL_LOG_COLLECTION') -eq '1')
    $controls.UseTechRunGuardHistoryCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY') -ne '0')
    $controls.IgnoreTechRunGuardHistoryCheck.IsChecked = ((Get-ConfiguredValue 'W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY') -eq '1')
    $controls.MaxCyclesText.Text = Get-ConfiguredValue 'W11UT_GUI_MAX_CYCLES'
    Invoke-LauncherOptionStateUpdate
}

function Invoke-OptionCheckAvailabilityUpdate {
    param(
        [object]$CheckBox,
        [bool]$Enabled
    )

    $CheckBox.IsEnabled = $Enabled
    $CheckBox.Opacity = if ($Enabled) { 1.0 } else { 0.55 }
}

function Invoke-LauncherOptionStateUpdate {
    $auditOnly = [bool]$controls.AuditOnlyCheck.IsChecked
    $directSetup = [bool]$controls.DirectSetupUpgradeCheck.IsChecked

    $auditControlledChecks = @(
        $controls.AllowPolicyRepairCheck,
        $controls.AllowWUResetCheck,
        $controls.AllowForceUpgradeCheck,
        $controls.AllowSetupUpgradeCheck,
        $controls.DirectSetupUpgradeCheck,
        $controls.AllowRebootCheck,
        $controls.ScheduleRetryAfterRebootCheck,
        $controls.SetupCompletionRebootCheck,
        $controls.AllowSetupProfileRepairCheck,
        $controls.SkipVirtualMachinesCheck,
        $controls.AllowDiskCleanupCheck,
        $controls.AllowAdvancedCleanupCheck,
        $controls.SkipSetupPreCopyCheck,
        $controls.KeepCentralHistoryCheck,
        $controls.NoCentralCollectionCheck,
        $controls.DryRunCheck
    )

    $directSetupIgnoredChecks = @(
        $controls.AllowPolicyRepairCheck,
        $controls.AllowWUResetCheck,
        $controls.AllowForceUpgradeCheck,
        $controls.AllowSetupUpgradeCheck,
        $controls.AllowDiskCleanupCheck,
        $controls.AllowAdvancedCleanupCheck
    )

    foreach ($check in $auditControlledChecks) {
        if ($auditOnly) {
            $check.IsChecked = $false
        }
        Invoke-OptionCheckAvailabilityUpdate -CheckBox $check -Enabled (-not $auditOnly)
    }

    if ($auditOnly) {
        return
    }

    foreach ($check in $directSetupIgnoredChecks) {
        if ($directSetup) {
            $check.IsChecked = $false
        }
        Invoke-OptionCheckAvailabilityUpdate -CheckBox $check -Enabled (-not $directSetup)
    }
}

function Get-ToolkitOptionArguments {
    $arguments = New-Object System.Collections.Generic.List[string]
    $arguments.Add('-ThrottleLimit')
    $arguments.Add((Get-IntText -TextBox $controls.ThrottleText -Default 10 -Minimum 1))

    $maxCycles = Get-IntText -TextBox $controls.MaxCyclesText -Default 0 -Minimum 0
    if ([int]$maxCycles -gt 0) {
        $arguments.Add('-MaxCycles')
        $arguments.Add($maxCycles)
    }

    if ([bool]$controls.DryRunCheck.IsChecked) { $arguments.Add('-DryRun') }
    if ([bool]$controls.KeepCentralHistoryCheck.IsChecked) { $arguments.Add('-KeepCentralLogHistory') }
    if ([bool]$controls.NoCentralCollectionCheck.IsChecked) { $arguments.Add('-NoCentralLogCollection') }
    return @($arguments)
}

function Get-ToolkitOptionEnvironment {
    $globalLimit = Get-IntText -TextBox $controls.GlobalLimitText -Default 15 -Minimum 1
    $controls.GlobalLimitOptionText.Text = $globalLimit

    $environment = @{
        W11UT_GLOBAL_CONCURRENCY_LIMIT                 = $globalLimit
        W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = Get-IntText -TextBox $controls.GlobalLeaseText -Default 0 -Minimum 0
        W11UT_THROTTLE                                 = Get-IntText -TextBox $controls.ThrottleText -Default 10 -Minimum 1
        W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS          = Get-IntText -TextBox $controls.ComputerDelayText -Default 0 -Minimum 0
        W11UT_DELAY_BETWEEN_CYCLES_MINUTES             = Get-IntText -TextBox $controls.CycleDelayText -Default 10 -Minimum 0
        W11UT_PSEXEC_TIMEOUT_MINUTES                   = Get-IntText -TextBox $controls.PsExecTimeoutText -Default 360 -Minimum 1
        W11UT_SETUP_EXECUTION_MODE                     = [string]$controls.SetupModeCombo.Text
        W11UT_SETUP_LANGUAGE                           = [string]$controls.SetupLanguageCombo.Text
        W11UT_SETUP_DYNAMIC_UPDATE                     = [string]$controls.SetupDynamicUpdateCombo.Text
        W11UT_SETUP_SOURCE_CANDIDATE_LIMIT             = Get-IntText -TextBox $controls.SetupCandidateLimitText -Default 5 -Minimum 0
        W11UT_SETUP_COPY_IPG_MS                        = Get-IntText -TextBox $controls.SetupCopyIpgText -Default 20 -Minimum 0
        W11UT_SETUP_COPY_JITTER_SECONDS                = Get-IntText -TextBox $controls.SetupCopyJitterText -Default 300 -Minimum 0
        W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT           = Get-IntText -TextBox $controls.SetupSubnetLimitText -Default 1 -Minimum 0
        W11UT_SETUP_SUBNET_PREFIX_LENGTH               = [string]$controls.SetupSubnetPrefixCombo.Text
        W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES   = Get-IntText -TextBox $controls.SetupSubnetLeaseText -Default 90 -Minimum 1
        W11UT_AD_DOMAIN                                = Get-ConfiguredValue 'W11UT_AD_DOMAIN'
        W11UT_INTUNE_TENANT_ID                         = Get-ConfiguredValue 'W11UT_INTUNE_TENANT_ID'
        W11UT_INTUNE_INVENTORY_PAGE_SIZE               = Get-ConfiguredValue 'W11UT_INTUNE_INVENTORY_PAGE_SIZE'
        W11UT_SKIP_INTUNE_INVENTORY_REFRESH            = Get-ConfiguredValue 'W11UT_SKIP_INTUNE_INVENTORY_REFRESH'
        W11UT_AUDIT_ONLY                               = Get-BooleanText -CheckBox $controls.AuditOnlyCheck
        W11UT_ALLOW_POLICY_REPAIR                      = Get-BooleanText -CheckBox $controls.AllowPolicyRepairCheck
        W11UT_ALLOW_WU_RESET                           = Get-BooleanText -CheckBox $controls.AllowWUResetCheck
        W11UT_ALLOW_FORCE_UPGRADE                      = Get-BooleanText -CheckBox $controls.AllowForceUpgradeCheck
        W11UT_ALLOW_SETUP_UPGRADE                      = Get-BooleanText -CheckBox $controls.AllowSetupUpgradeCheck
        W11UT_DIRECT_SETUP_UPGRADE                     = Get-BooleanText -CheckBox $controls.DirectSetupUpgradeCheck
        W11UT_ALLOW_REBOOT                             = Get-BooleanText -CheckBox $controls.AllowRebootCheck
        W11UT_SCHEDULE_RETRY_AFTER_REBOOT       = Get-BooleanText -CheckBox $controls.ScheduleRetryAfterRebootCheck
        W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS   = Get-ConfiguredValue 'W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS'
        W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS  = Get-ConfiguredValue 'W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS'
        W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS = Get-IntText -TextBox $controls.ForceRequiredRebootDaysText -Default 7 -Minimum 0
        W11UT_SETUP_REBOOT_WHEN_NO_USER                = Get-BooleanText -CheckBox $controls.SetupCompletionRebootCheck
        W11UT_ALLOW_SETUP_PROFILE_REPAIR               = Get-BooleanText -CheckBox $controls.AllowSetupProfileRepairCheck
        W11UT_SKIP_VIRTUAL_MACHINES                    = Get-BooleanText -CheckBox $controls.SkipVirtualMachinesCheck
        W11UT_ALLOW_DISK_CLEANUP                       = Get-BooleanText -CheckBox $controls.AllowDiskCleanupCheck
        W11UT_ALLOW_ADVANCED_DISK_CLEANUP              = Get-BooleanText -CheckBox $controls.AllowAdvancedCleanupCheck
        W11UT_SKIP_SETUP_MEDIA_PRECOPY                 = Get-BooleanText -CheckBox $controls.SkipSetupPreCopyCheck
        W11UT_GUI_DRY_RUN                              = Get-BooleanText -CheckBox $controls.DryRunCheck
        W11UT_GUI_KEEP_CENTRAL_LOG_HISTORY             = Get-BooleanText -CheckBox $controls.KeepCentralHistoryCheck
        W11UT_GUI_NO_CENTRAL_LOG_COLLECTION            = Get-BooleanText -CheckBox $controls.NoCentralCollectionCheck
        W11UT_GUI_MAX_CYCLES                           = Get-IntText -TextBox $controls.MaxCyclesText -Default 0 -Minimum 0
        W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY         = Get-BooleanText -CheckBox $controls.UseTechRunGuardHistoryCheck
        W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY      = Get-BooleanText -CheckBox $controls.IgnoreTechRunGuardHistoryCheck
        W11UT_RUN_GUARD_HOURS                          = Get-ConfiguredValue 'W11UT_RUN_GUARD_HOURS'
    }

    foreach ($pair in @(
        @{ Key = 'W11UT_SETUP_MEDIA_ID'; Control = $controls.SetupMediaIdText },
        @{ Key = 'W11UT_SETUP_SOURCE_MAP'; Control = $controls.SetupSourceMapText },
        @{ Key = 'W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT'; Control = $controls.SetupSubnetGateRootText }
    )) {
        $value = $pair.Control.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $environment[$pair.Key] = $value
        }
    }

    $setupSource = Get-SetupSourceText
    if (-not [string]::IsNullOrWhiteSpace($setupSource)) {
        $environment['W11UT_SETUP_SOURCE'] = $setupSource
    }

    return $environment
}

function Test-SetupSourceBeforeLaunch {
    param(
        [hashtable]$EnvironmentVariables,
        [string]$ScopeName
    )

    $setupSource = [string]$EnvironmentVariables['W11UT_SETUP_SOURCE']
    $setupSourceMap = [string]$EnvironmentVariables['W11UT_SETUP_SOURCE_MAP']
    $mode = [string]$EnvironmentVariables['W11UT_SETUP_EXECUTION_MODE']
    $allowSetupUpgrade = ([string]$EnvironmentVariables['W11UT_ALLOW_SETUP_UPGRADE'] -eq '1' -or [string]$EnvironmentVariables['W11UT_DIRECT_SETUP_UPGRADE'] -eq '1')
    $skipSetupPreCopy = ([string]$EnvironmentVariables['W11UT_SKIP_SETUP_MEDIA_PRECOPY'] -eq '1')

    if (-not $allowSetupUpgrade) { return $true }
    if ($skipSetupPreCopy -and $mode -ne 'Share') { return $true }

    if ([string]::IsNullOrWhiteSpace($setupSource) -and [string]::IsNullOrWhiteSpace($setupSourceMap)) {
        Show-GuiError ("Setup source or setup source map is required for setup upgrade when target media copy is enabled.`n`nScope: {0}`n`nUse a UNC path reachable by target computers, or enable setup media pre-copy skip only when target cache is already valid." -f $ScopeName)
        $controls.SetupSourceText.Focus() | Out-Null
        return $false
    }

    $setupSources = @(
        Split-SetupSourceText -Value $setupSource
        if (-not [string]::IsNullOrWhiteSpace($setupSourceMap)) { $setupSourceMap }
    )
    $localSources = @($setupSources | Where-Object { -not (Test-IsUncPath -Path $_) })
    if ($setupSources.Count -gt 0 -and $localSources.Count -eq 0) {
        return $true
    }

    return Show-GuiWarningYesNo -Title 'Confirm local setup source' -Message ("One or more setup sources are not UNC paths:`n{0}`n`nIn LOT/PsExec mode, target computers run as SYSTEM and must read these paths themselves. Continue anyway?" -f ($localSources -join "`n"))
}

function Get-LaunchableLotSummaries {
    @($script:Lots | Where-Object { $_.ComputerCount -gt 0 -and $_.WrappersReady })
}

function Ensure-LotWrappersReady {
    param([Parameter(Mandatory = $true)][pscustomobject]$Lot)

    if ($Lot.WrappersReady) { return $Lot }

    Add-Status -Title 'Wrappers' -Message ("Wrappers missing for {0}. Refreshing wrappers before launch..." -f $Lot.Name)
    Invoke-LotWrapperRefresh -RootPath $toolkitRoot
    Refresh-LotList

    $refreshed = $script:Lots | Where-Object { $_.Name -eq $Lot.Name } | Select-Object -First 1
    if (-not $refreshed) { throw ("LOT disappeared after wrapper refresh: {0}" -f $Lot.Name) }
    if (-not $refreshed.WrappersReady) {
        $missing = @($refreshed.MissingWrappers) -join ', '
        throw ("Selected LOT wrappers are still missing after refresh: {0}. Missing={1}" -f $refreshed.Name,$missing)
    }

    return $refreshed
}

function Set-LotUiAvailability {
    param([bool]$Enabled)

    foreach ($name in @('OpenLotFolderButton','OpenLotComputersButton','OpenLotReportsButton','RefreshWrappersButton','LaunchLotButton')) {
        $controls[$name].IsEnabled = $Enabled
    }
}

function Update-SelectedLotView {
    $script:SelectedLot = $controls.LotCombo.SelectedItem
    if (-not $script:SelectedLot) {
        $controls.LotDeviceCountText.Text = '-'
        $controls.LotScopeText.Text = '-'
        $controls.LotWrappersText.Text = '-'
        Set-LotUiAvailability -Enabled $false
        return
    }

    $controls.LotDeviceCountText.Text = [string]$script:SelectedLot.ComputerCount
    $controls.LotScopeText.Text = $script:SelectedLot.UpgradeScope
    $controls.LotWrappersText.Text = if ($script:SelectedLot.WrappersReady) { 'Ready' } else { 'Missing' }
    Set-LotUiAvailability -Enabled $true
    Add-Status -Title 'Selected' -Message ("{0}: {1} device(s)." -f $script:SelectedLot.Name, $script:SelectedLot.ComputerCount)
}

function Refresh-LotList {
    $script:Lots = @(Get-LotFolders -RootPath $toolkitRoot | ForEach-Object { Get-LotSummary -LotPath $_.FullName })
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

function Confirm-UnlimitedCycleLaunch {
    $maxCycles = [int](Get-IntText -TextBox $controls.MaxCyclesText -Default 0 -Minimum 0)
    if ($maxCycles -gt 0) { return $true }

    $choice = [System.Windows.MessageBox]::Show(
        "Max cycles is 0, so this launcher will keep cycling until all devices are removed, the operator stops it, or an external cancellation is requested.`n`nFailure backoff is enabled, but an unlimited run can remain active for days. Continue?",
        'Unlimited LOT cycles',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning,
        [System.Windows.MessageBoxResult]::No
    )
    return ($choice -eq [System.Windows.MessageBoxResult]::Yes)
}

Register-NumericSteppers
Initialize-Options

$controls.SetupSourceText.Add_GotKeyboardFocus({ Clear-SetupSourcePlaceholder })
$controls.SetupSourceText.Add_LostKeyboardFocus({ Restore-SetupSourcePlaceholder })
$controls.AuditOnlyCheck.Add_Checked({ Invoke-LauncherOptionStateUpdate })
$controls.AuditOnlyCheck.Add_Unchecked({ Invoke-LauncherOptionStateUpdate })
$controls.DirectSetupUpgradeCheck.Add_Checked({ Invoke-LauncherOptionStateUpdate })
$controls.DirectSetupUpgradeCheck.Add_Unchecked({ Invoke-LauncherOptionStateUpdate })
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
$controls.OpenLotReportsButton.Add_Click({ if ($script:SelectedLot) { Open-FolderPath -Path $script:SelectedLot.ReportsPath } })
$controls.RefreshWrappersButton.Add_Click({
    try {
        Invoke-LotWrapperRefresh -RootPath $toolkitRoot
        Add-Status -Message 'Wrappers refreshed.'
        Refresh-LotList
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.LaunchLotButton.Add_Click({
    try {
        if (-not $script:SelectedLot) { return }
        if ($script:SelectedLot.ComputerCount -le 0) { throw 'Selected LOT has no device in Computers.txt.' }
        if (-not (Confirm-UnlimitedCycleLaunch)) { return }
        $script:SelectedLot = Ensure-LotWrappersReady -Lot $script:SelectedLot
        $environment = Get-ToolkitOptionEnvironment
        $effectiveEnvironment = Get-EffectiveLotEnvironment -LotPath $script:SelectedLot.Path -EnvironmentVariables $environment
        if (-not (Test-SetupSourceBeforeLaunch -EnvironmentVariables $effectiveEnvironment -ScopeName $script:SelectedLot.Name)) { return }
        Start-ToolkitLot -Lot $script:SelectedLot -Mode ([string]$controls.LotModeCombo.SelectedItem) -AdditionalArguments (Get-ToolkitOptionArguments) -EnvironmentVariables $environment
        Save-GuiOptions -Quiet
        Add-Status -Title 'Launched' -Message ("Launched LOT {0}." -f $script:SelectedLot.Name)
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.LaunchAllButton.Add_Click({
    try {
        if (-not (Confirm-UnlimitedCycleLaunch)) { return }
        $missingWrapperLots = @($script:Lots | Where-Object { $_.ComputerCount -gt 0 -and -not $_.WrappersReady })
        if ($missingWrapperLots.Count -gt 0) {
            Add-Status -Title 'Wrappers' -Message ("{0} LOT(s) have missing wrappers. Refreshing wrappers before launch all..." -f $missingWrapperLots.Count)
            Invoke-LotWrapperRefresh -RootPath $toolkitRoot
            Refresh-LotList
        }
        $lots = Get-LaunchableLotSummaries
        if ($lots.Count -eq 0) { throw 'No ready LOT with devices was found.' }
        foreach ($lot in $lots) {
            $environment = Get-ToolkitOptionEnvironment
            $effectiveEnvironment = Get-EffectiveLotEnvironment -LotPath $lot.Path -EnvironmentVariables $environment
            if (-not (Test-SetupSourceBeforeLaunch -EnvironmentVariables $effectiveEnvironment -ScopeName $lot.Name)) { return }
            Start-ToolkitLot -Lot $lot -Mode ([string]$controls.LotModeCombo.SelectedItem) -AdditionalArguments (Get-ToolkitOptionArguments) -EnvironmentVariables $environment
            Add-Status -Title 'Launch all' -Message ("Launched {0}. Next LOT starts in {1}s." -f $lot.Name, $launchAllLotStartDelaySeconds)
            Wait-UiDelay -Seconds $launchAllLotStartDelaySeconds
        }
        Save-GuiOptions -Quiet
        Add-Status -Title 'Launch all' -Message ("Launched {0} LOT(s)." -f $lots.Count)
    } catch {
        Show-GuiError $_.Exception.Message
    }
})
$controls.LaunchSingleButton.Add_Click({
    try {
        $computer = $controls.SingleComputerText.Text.Trim()
        if (-not $computer) { throw 'Enter a computer name.' }
        if (-not (Confirm-UnlimitedCycleLaunch)) { return }
        $environment = Get-ToolkitOptionEnvironment
        if (-not (Test-SetupSourceBeforeLaunch -EnvironmentVariables $environment -ScopeName $computer)) { return }
        $context = Start-ToolkitSingleComputer -ToolkitRoot $toolkitRoot -ComputerName $computer -Mode ([string]$controls.SingleModeCombo.SelectedItem) -AdditionalArguments (Get-ToolkitOptionArguments) -EnvironmentVariables $environment
        Save-GuiOptions -Quiet
        $script:LastSingleRunFolder = $context.RunPath
        $controls.SingleRunFolderText.Text = $context.RunPath
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
        $result = New-ToolkitLotFolder -RootPath $toolkitRoot -Name $controls.NewLotNameText.Text
        $controls.NewLotComputersPathText.Text = $result.ComputersPath
        Add-Status -Title 'Created' -Message ("Created LOT {0}." -f (Split-Path -Leaf $result.LotPath))
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

$script:AutomaticPreviewContext = $null
$script:AutomaticPreviewResult = $null
$script:AutomaticPreviewSignature = ''
$controls.AutomaticSourceCombo.Add_SelectionChanged({
    Set-AutomaticFilterPreviewStale -Message 'Inventory source changed. Refresh the preview before creating the LOT.'
})
$controls.AutomaticNamePrefixText.Add_TextChanged({
    Update-AutomaticGeneratedLotName
    Set-AutomaticFilterPreviewStale
})
$controls.AutomaticNameContainsText.Add_TextChanged({ Set-AutomaticFilterPreviewStale })
$controls.AutomaticExcludeIntuneCheck.Add_Checked({ Set-AutomaticFilterPreviewStale })
$controls.AutomaticExcludeIntuneCheck.Add_Unchecked({ Set-AutomaticFilterPreviewStale })
$controls.AutomaticExcludeStaleAdCheck.Add_Checked({
    Update-AutomaticFilterControlState
    Set-AutomaticFilterPreviewStale
})
$controls.AutomaticExcludeStaleAdCheck.Add_Unchecked({
    Update-AutomaticFilterControlState
    Set-AutomaticFilterPreviewStale
})
$controls.AutomaticLastLogonDaysText.Add_TextChanged({ Set-AutomaticFilterPreviewStale })
$controls.AutomaticPreviewButton.Add_Click({
    try {
        $forceRefresh = [bool]$controls.AutomaticForceRefreshCheck.IsChecked
        [void](Invoke-AutomaticLotPreviewWithProgress -ForceInventoryRefresh:$forceRefresh)
    }
    catch {
        Show-GuiError $_.Exception.Message
        Add-Status -Title 'Automatic preview failed' -Message $_.Exception.Message
    }
})
$controls.AutomaticOpenEvidenceButton.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace([string]$controls.AutomaticEvidencePathText.Text)) {
        Open-FolderPath -Path $controls.AutomaticEvidencePathText.Text
    }
})
$controls.AutomaticCreateButton.Add_Click({
    try {
        $currentSignature = Get-AutomaticPreviewSignature
        $forceRefresh = [bool]$controls.AutomaticForceRefreshCheck.IsChecked
        if ($forceRefresh -or $null -eq $script:AutomaticPreviewResult -or $script:AutomaticPreviewSignature -ne $currentSignature) {
            [void](Invoke-AutomaticLotPreviewWithProgress -ForceInventoryRefresh:$forceRefresh)
        }

        if ([int]$script:AutomaticPreviewResult.Summary.SelectedDevices -le 0) {
            throw 'No eligible Windows 10 device was selected. Review the exclusion evidence before creating a LOT.'
        }
        if (-not (Confirm-AutomaticLotCreate -Result $script:AutomaticPreviewResult)) { return }

        $created = Invoke-AutomaticLotSelection -InventoryContext $script:AutomaticPreviewContext -Create
        $lot = Get-LotSummary -LotPath ([string]$created.Summary.LotPath)

        $script:AutomaticPreviewResult = $created
        $controls.AutomaticSummaryText.Text = Format-AutomaticLotSummary -Result $created -InventoryContext $script:AutomaticPreviewContext
        $controls.AutomaticEvidencePathText.Text = [string]$created.Summary.EvidencePath
        $controls.AutomaticOpenEvidenceButton.IsEnabled = $true
        $script:SelectedLot = $lot
        Refresh-LotList
        Save-GuiOptions -Quiet
        Add-Status -Title 'Automatic LOT created' -Message ("Created {0} with {1} device(s). Open Existing LOT to launch it." -f $lot.Name, $lot.ComputerCount)
    }
    catch {
        Show-GuiError $_.Exception.Message
        Add-Status -Title 'Automatic LOT creation failed' -Message $_.Exception.Message
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

if (Get-Command -Name Start-SmartM365GuiUpdateCheck -ErrorAction SilentlyContinue) {
    $window.Add_ContentRendered({
        try {
            $manifestPath = Join-Path $toolkitRoot 'SmartM365.GuiUpdateCheck.psd1'
            $script:UpdateCheckTimer = Start-SmartM365GuiUpdateCheck -Owner $window -ManifestPath $manifestPath -AppRoot $toolkitRoot -OnStatus {
                param(
                    [string]$Message,
                    [string]$Title
                )
                Add-Status -Title $Title -Message $Message
            }
        }
        catch { [void]$_.Exception }
    })
}

try {
    Refresh-LotList
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBkWuHIUlibYOt0
# FjALWcsZcV7QWRbAQ1qY4j7P042ECKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOGK2rC2Ok62tBFypYiXKSf4E9gr3GDLDwP6ujHbgfUPMA0GCSqG
# SIb3DQEBAQUABIIBgH3nuW3K+XFfJrxODKmk8OxjGodQYQvR8ALSj1lO48JrMP34
# 5aYC4UMX1ahh8bBuoMayF+p+2skWxlg6o9z1mwo3diPXcQ51y+9I0dvZCu5aLpK8
# zwsIgLiBYRpQOJ3upQsaZy+3W3Tggjzn94J/T7/0YMNxxEMp4wvJrZCsaJo2zgFS
# rhg4AMcqZCVjb7Pov4692XIKgjqeJzhQBQQ/ch7o7JlSSytlNGKgsn1uEUXNjZ7I
# B6svjsmE7pz3CxNBd3czISX6wFxKWdxrhqVjR+IJ9pryFcfQolgDAKzBk0+krjV6
# WVHiSsH0RrUgkzZmF+Ekx0YQmKQPhStFA74ODWGCIDTW32xJYi/S+Y2+Rcs8+9wq
# Rseq5YytVLQDZDI/VeaOExeFi90NLO2GKb1bHaEeMftIs4Mb1qKn5LAjf3GqfoyX
# aXfaUTmpqHlZhO6D16Z4ziwBoUoHN+XpbAnvfCv8IQ0N7TxbwrtQb/9JPWq8CYAc
# LyyzJZx8heiP0tquraGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjYwOTA5
# NDBaMC8GCSqGSIb3DQEJBDEiBCCp4FOblN2nGgJR42xiJKeGzKxUV5VWvFG+ywTw
# 30gK7zANBgkqhkiG9w0BAQEFAASCAgCKMdUrAz+q+nKnBN7VnfqXSOsv4+oCGUw7
# LzoKat6NUJuf4QQYKtS1zQzfzxC42R1ThKLH0dv76vqhfHdqF6KUyXhdboPlR0Es
# eUCZ2PgTyVADJaaoVVmlhPxf5G0Lmj8j/pnXX9Q7gq+5Z64b0IC6xGCsUyx7LbMG
# ivP/o3bxrUxRPkwymRJbrMf3fPcLa+hT3/ImVzXx25JmMDuqfk+yubglgC73fYNi
# IisFn//lSm8Alh3HcRWAPX0VwSai7xfVpaPCgtz9I0dYnlosABa0sP2uusMS6eyX
# WfnlitIIIeJtzhP8uiTUqpW/n61jg5G4wv0n5yDreKEzUd3Fgqx1Z0Oe55+HHFTT
# 3qqXT5YlC3MyZHlThPesGTLk4kfH8FqE1KvQJZDpXrnxkAbED+7i6ZlewQu5g4P0
# Dv4lYLFon8jm5xeZvL0NWJMobq7+SThP/iIhv+2uYp/+E/Vwq1JSWO35Fi/cieAk
# Gdu/2VakcFH0YVxK8L5dIrzzFP5K5ADsywWXdwMOzqczgeC78lU556I5YI4Yhasn
# 2nGIIgwhrv++or5na0YaJr9jfPNT+AG99rd0/56MM20gafAOBGaTCfGorzwwOdmE
# 9TsrbnMkllcQz8y0JVFIuO9U8PWI0rEJd+NaAw0vNppwSwkVEH/Wm+hItxlSTyfP
# g2YyN1dBMg==
# SIG # End signature block
