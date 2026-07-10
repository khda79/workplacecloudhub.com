<#
.SYNOPSIS
Starts the Windows 11 Upgrade LOT launcher GUI.

.VERSION
0.1.28
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

function ConvertTo-CmdArgument {
    param([string]$Value)

    if ($null -eq $Value) {
        return '""'
    }

    if ($Value -match '^[A-Za-z0-9_:\\./=-]+$') {
        return $Value
    }

    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

function ConvertTo-CmdSetCommand {
    param(
        [string]$Name,
        [AllowNull()][string]$Value
    )

    if ($Name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
        throw "Invalid environment variable name: $Name"
    }

    return ('set "{0}={1}"' -f $Name, (($Value -replace '"', '\"')))
}


function ConvertTo-CmdWindowTitle {
    param([AllowNull()][string]$Value)

    $title = ([string]$Value -replace '[\r\n\t]+', ' ').Trim()
    $title = $title -replace '[&|<>^]', '-'
    if ($title.Length -gt 240) { $title = $title.Substring(0, 240) }
    return $title
}

function New-GuiLaunchCommandFile {
    param(
        [Parameter(Mandatory = $true)][string]$WorkingDirectory,
        [Parameter(Mandatory = $true)][string[]]$Commands,
        [string]$NamePrefix = 'W11UT-GUI',
        [string]$WindowTitle,
        [switch]$PauseWhenDone
    )

    $launcherRoot = Join-Path $env:TEMP 'SmartM365\Windows11UpgradeToolkit\GuiLaunchers'
    New-Item -ItemType Directory -Path $launcherRoot -Force | Out-Null

    $safePrefix = [regex]::Replace($NamePrefix, '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safePrefix)) { $safePrefix = 'W11UT-GUI' }

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

function Start-ToolkitLot {
    param(
        [pscustomobject]$Lot,
        [string]$Mode,
        [string[]]$AdditionalArguments = @(),
        [hashtable]$EnvironmentVariables = @{}
    )

    $wrapperName = switch ($Mode) {
        'Loop' { 'Run-Windows11UpgradeRepairWithPsExec-Loop.cmd'; break }
        'Once' { 'Run-Windows11UpgradeRepairWithPsExec-Once.cmd'; break }
        'LoopIgnoreRunGuard' { 'Run-Windows11UpgradeRepairWithPsExec-Loop-IgnoreRunGuard.cmd'; break }
        'OnceIgnoreRunGuard' { 'Run-Windows11UpgradeRepairWithPsExec-Once-IgnoreRunGuard.cmd'; break }
        default { 'Run-Windows11UpgradeRepairWithPsExec-Once.cmd' }
    }

    $wrapperPath = Join-Path $Lot.Path $wrapperName
    if (-not (Test-Path -LiteralPath $wrapperPath -PathType Leaf)) {
        throw "LOT wrapper not found: $wrapperPath"
    }

    [void](Resolve-GuiPsExecPath -ToolkitRoot $toolkitRoot)

    $effectiveEnvironment = Get-EffectiveLotEnvironment -LotPath $Lot.Path -EnvironmentVariables $EnvironmentVariables
    $run = New-LotRunContext -RootPath $toolkitRoot -LotName $Lot.Name
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
    $launchTitle = "{0} - started {1}" -f $Lot.Name,(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $launchCommandPath = New-GuiLaunchCommandFile -WorkingDirectory $run.RunPath -Commands @($commands) -NamePrefix ($Lot.Name + '-' + $Mode) -WindowTitle $launchTitle
    Start-GuiLaunchCommandFile -LaunchCommandPath $launchCommandPath -WorkingDirectory $run.RunPath
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

    $commandParts = New-Object System.Collections.Generic.List[string]
    foreach ($item in @(
        'powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File', $scriptPath,
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
        @{ Name = 'W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS'; Argument = '-RetryAfterRebootDelaySeconds' }
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

    $launchCommandPath = New-GuiLaunchCommandFile -WorkingDirectory $run.RunPath -Commands @(($commandParts -join ' ')) -NamePrefix ('Single-' + $run.SafeComputerName) -PauseWhenDone
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
    W11UT_SKIP_INTUNE_INVENTORY_REFRESH = '0'
    W11UT_SETUP_EXECUTION_MODE = 'LocalCache'
    W11UT_SETUP_MEDIA_ID = 'Win11'
    W11UT_SETUP_LANGUAGE = 'MatchSystem'
    W11UT_SETUP_DYNAMIC_UPDATE = 'Disable'
    W11UT_THROTTLE = '10'
    W11UT_GLOBAL_CONCURRENCY_LIMIT = '15'
    W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = '0'
    W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS = '0'
    W11UT_DELAY_BETWEEN_CYCLES_MINUTES = '5'
    W11UT_PSEXEC_TIMEOUT_MINUTES = '360'
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
    'HeaderLogoLink','HeaderLogoImage','StatusTitle','StatusText','RefreshButton','LaunchAllButton','LotCombo','LotDeviceCountText',
    'LotScopeText','LotWrappersText','LotModeCombo','GlobalLimitText','GlobalLimitDownButton','GlobalLimitUpButton','OpenLotFolderButton',
    'OpenLotComputersButton','OpenLotReportsButton','RefreshWrappersButton','LaunchLotButton',
    'ActivityText','SingleComputerText','SingleModeCombo','LaunchSingleButton','SingleRunFolderText',
    'OpenSingleRunFolderButton','NewLotNameText','CreateLotButton','NewLotComputersPathText',
    'OpenNewLotComputersButton','DryRunCheck','AuditOnlyCheck','AllowPolicyRepairCheck',
    'AllowWUResetCheck','AllowForceUpgradeCheck','AllowSetupUpgradeCheck','AllowRebootCheck',
    'ScheduleRetryAfterRebootCheck','SetupCompletionRebootCheck','AllowSetupProfileRepairCheck',
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

function Show-GuiWarningYesNo {
    param(
        [string]$Message,
        [string]$Title
    )

    $result = [System.Windows.MessageBox]::Show($window, $Message, $Title, 'YesNo', 'Warning')
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
    Register-NumericStepper -TextBoxName 'SetupCandidateLimitText' -DownButtonName 'SetupCandidateLimitDownButton' -UpButtonName 'SetupCandidateLimitUpButton' -Default 5 -Minimum 0
    Register-NumericStepper -TextBoxName 'SetupCopyIpgText' -DownButtonName 'SetupCopyIpgDownButton' -UpButtonName 'SetupCopyIpgUpButton' -Default 20 -Minimum 0 -Step 5
    Register-NumericStepper -TextBoxName 'SetupCopyJitterText' -DownButtonName 'SetupCopyJitterDownButton' -UpButtonName 'SetupCopyJitterUpButton' -Default 300 -Minimum 0 -Step 30
    Register-NumericStepper -TextBoxName 'SetupSubnetLimitText' -DownButtonName 'SetupSubnetLimitDownButton' -UpButtonName 'SetupSubnetLimitUpButton' -Default 1 -Minimum 0
    Register-NumericStepper -TextBoxName 'SetupSubnetLeaseText' -DownButtonName 'SetupSubnetLeaseDownButton' -UpButtonName 'SetupSubnetLeaseUpButton' -Default 90 -Minimum 1 -Step 5
    Register-NumericStepper -TextBoxName 'ThrottleText' -DownButtonName 'ThrottleDownButton' -UpButtonName 'ThrottleUpButton' -Default 10 -Minimum 1
    Register-NumericStepper -TextBoxName 'ComputerDelayText' -DownButtonName 'ComputerDelayDownButton' -UpButtonName 'ComputerDelayUpButton' -Default 0 -Minimum 0 -Step 5
    Register-NumericStepper -TextBoxName 'CycleDelayText' -DownButtonName 'CycleDelayDownButton' -UpButtonName 'CycleDelayUpButton' -Default 5 -Minimum 0
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
    Initialize-Combo -Combo $controls.LotModeCombo -Values @('Once','Loop','OnceIgnoreRunGuard','LoopIgnoreRunGuard') -Selected 'Once'
    Initialize-Combo -Combo $controls.SingleModeCombo -Values @('Once','OnceIgnoreRunGuard','Loop','LoopIgnoreRunGuard') -Selected 'Once'
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
        W11UT_DELAY_BETWEEN_CYCLES_MINUTES             = Get-IntText -TextBox $controls.CycleDelayText -Default 5 -Minimum 0
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

$window.Add_Closing({ Save-GuiOptions -Quiet })
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
