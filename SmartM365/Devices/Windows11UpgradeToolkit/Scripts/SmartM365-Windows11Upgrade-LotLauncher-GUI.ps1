<#
.SYNOPSIS
Creates and launches Smart Intune Windows 11 Upgrade Toolkit LOT folders from a GUI.

.DESCRIPTION
This operator GUI lists existing LOT-* folders for launch, shows their key
parameters, and offers a simplified new-LOT creation flow.

Operational LOT-* folders can contain real computer names and are ignored by Git.
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'

function Get-ToolkitRoot {
    $scriptDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    }
    else {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }

    return (Get-Item -LiteralPath (Split-Path -Parent $scriptDir) -ErrorAction Stop).FullName
}

function Read-ToolkitConfig {
    param([Parameter(Mandatory = $true)][string]$Path)

    $config = @{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $config
    }

    foreach ($rawLine in @(Get-Content -LiteralPath $Path -ErrorAction Stop)) {
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

function Get-ToolkitConfigValue {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()][string]$Default = $null
    )

    $value = [Environment]::GetEnvironmentVariable($Name, 'Process')
    if (-not [string]::IsNullOrWhiteSpace($value)) {
        return $value
    }

    if ($script:ToolkitConfig -and $script:ToolkitConfig.ContainsKey($Name)) {
        return [string]$script:ToolkitConfig[$Name]
    }

    return $Default
}

function Get-LotConfigPath {
    param([Parameter(Mandatory = $true)][string]$LotPath)

    return (Join-Path $LotPath 'Windows11UpgradeToolkit.config')
}

function Test-LotConfigCanOverrideEnvironmentValue {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
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

    if ($script:ToolkitConfig -and $script:ToolkitConfig.ContainsKey($Name) -and $currentValue -eq [string]$script:ToolkitConfig[$Name]) {
        return $true
    }

    if ($script:ToolkitDefaultEnvironment -and $script:ToolkitDefaultEnvironment.ContainsKey($Name) -and $currentValue -eq [string]$script:ToolkitDefaultEnvironment[$Name]) {
        return $true
    }

    return $false
}

function Get-EffectiveLotEnvironment {
    param(
        [Parameter(Mandatory = $true)][string]$LotPath,
        [hashtable]$EnvironmentVariables = @{}
    )

    $effective = @{}
    foreach ($key in @($EnvironmentVariables.Keys)) {
        $effective[$key] = [string]$EnvironmentVariables[$key]
    }

    $lotConfigPath = Get-LotConfigPath -LotPath $LotPath
    $lotConfig = Read-ToolkitConfig -Path $lotConfigPath
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
    param([Parameter(Mandatory = $true)][string]$LotName)

    $safeName = [regex]::Replace($LotName.Trim(), '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw 'Enter a LOT name.'
    }

    if ($safeName -notmatch '^(?i)LOT-') {
        $safeName = "LOT-$safeName"
    }

    return $safeName
}

function Get-ComputerNamesFromFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return @()
    }

    $seen = @{}
    $computers = New-Object System.Collections.ArrayList
    foreach ($line in @(Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $name = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith('#')) {
            continue
        }

        $key = $name.ToUpperInvariant()
        if ($seen.ContainsKey($key)) {
            continue
        }

        $seen[$key] = $true
        [void]$computers.Add($name)
    }

    return @($computers.ToArray())
}

function Get-LotFolders {
    param([Parameter(Mandatory = $true)][string]$RootPath)

    return @(
        Get-ChildItem -LiteralPath $RootPath -Directory -Filter 'LOT-*' -ErrorAction Stop |
            Where-Object { $_.Name -ine 'LOT-X' } |
            Sort-Object Name
    )
}

function Test-LotWrapperSet {
    param([Parameter(Mandatory = $true)][string]$LotPath)

    $wrapperNames = @(
        'Run-Windows11UpgradeRepairWithPsExec-Loop.cmd',
        'Run-Windows11UpgradeRepairWithPsExec-Once.cmd',
        'Run-Windows11UpgradeRepairWithPsExec-Loop-IgnoreRunGuard.cmd',
        'Run-Windows11UpgradeRepairWithPsExec-Once-IgnoreRunGuard.cmd'
    )

    $missing = @(
        foreach ($wrapperName in $wrapperNames) {
            $path = Join-Path $LotPath $wrapperName
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                $wrapperName
            }
        }
    )

    return [pscustomobject]@{
        Ready = ($missing.Count -eq 0)
        Missing = $missing
    }
}

function Get-LotSummary {
    param([Parameter(Mandatory = $true)][string]$LotPath)

    $computersPath = Join-Path $LotPath 'Computers.txt'
    $reportsPath = Join-Path $LotPath 'Reports'
    $computers = @(Get-ComputerNamesFromFile -Path $computersPath)
    $wrappers = Test-LotWrapperSet -LotPath $LotPath

    return [pscustomobject]@{
        Name = Split-Path -Leaf $LotPath
        Path = $LotPath
        ComputersPath = $computersPath
        ReportsPath = $reportsPath
        ComputerCount = $computers.Count
        UpgradeScope = 'Windows 10 readiness, policy repair, guarded Windows 11 upgrade'
        WrappersReady = $wrappers.Ready
        MissingWrappers = $wrappers.Missing
    }
}

function Invoke-LotWrapperRefresh {
    param([Parameter(Mandatory = $true)][string]$RootPath)

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
        [Parameter(Mandatory = $true)][string]$RootPath,
        [Parameter(Mandatory = $true)][string]$LotName
    )

    $rootItem = Get-Item -LiteralPath $RootPath -ErrorAction Stop
    $safeLotName = Get-SafeLotName -LotName $LotName
    $lotPath = Join-Path $rootItem.FullName $safeLotName
    if (Test-Path -LiteralPath $lotPath) {
        throw "LOT folder already exists: $lotPath"
    }

    New-Item -ItemType Directory -Path $lotPath -Force -ErrorAction Stop | Out-Null
    New-Item -ItemType File -Path (Join-Path $lotPath 'Computers.txt') -Force -ErrorAction Stop | Out-Null
    $lotConfigPath = Get-LotConfigPath -LotPath $lotPath
    if (-not (Test-Path -LiteralPath $lotConfigPath -PathType Leaf)) {
        $templateConfigPath = Join-Path $rootItem.FullName 'Windows11UpgradeToolkit.config.template'
        if (Test-Path -LiteralPath $templateConfigPath -PathType Leaf) {
            Copy-Item -LiteralPath $templateConfigPath -Destination $lotConfigPath -Force -ErrorAction Stop
        }
        else {
            Set-Content -LiteralPath $lotConfigPath -Value @(
                '# SmartM365 Windows 11 Upgrade Toolkit LOT defaults.'
                '# W11UT_SETUP_SOURCE should be a UNC path reachable from target computers.'
                'W11UT_SETUP_SOURCE='
            ) -Encoding ASCII -Force
        }
    }

    Invoke-LotWrapperRefresh -RootPath $rootItem.FullName

    return [pscustomobject]@{
        LotPath = $lotPath
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
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()][string]$Value
    )

    if ($Name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
        throw "Invalid environment variable name: $Name"
    }

    return ('set "{0}={1}"' -f $Name, (($Value -replace '"', '\"')))
}

function New-SingleComputerRunContext {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$ToolkitKey
    )

    $trimmed = $ComputerName.Trim().Trim([char]34)
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        throw 'Enter a computer name.'
    }

    $safeComputer = [regex]::Replace($trimmed, '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safeComputer)) { $safeComputer = 'Computer' }

    $runRoot = Get-SingleComputerRunRoot -ToolkitKey $ToolkitKey
    $runPath = Join-Path $runRoot ("{0}_{1}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'),$safeComputer)
    New-Item -ItemType Directory -Path $runPath -Force -ErrorAction Stop | Out-Null

    $computersPath = Join-Path $runPath 'Computers.txt'
    Set-Content -LiteralPath $computersPath -Value $trimmed -Encoding UTF8 -Force

    foreach ($folder in @('PsExecLogs','Reports','CentralLogs')) {
        New-Item -ItemType Directory -Path (Join-Path $runPath $folder) -Force -ErrorAction Stop | Out-Null
    }

    [pscustomobject]@{
        ComputerName = $trimmed
        RunPath = $runPath
        ComputersPath = $computersPath
        LogRoot = Join-Path $runPath 'PsExecLogs'
        ReportRoot = Join-Path $runPath 'Reports'
        CentralLogRoot = Join-Path $runPath 'CentralLogs'
    }
}

function Get-SingleComputerRunRoot {
    param([Parameter(Mandatory = $true)][string]$ToolkitKey)

    return (Join-Path $toolkitRoot 'SingleComputerRuns')
}

function Resolve-GuiPsExecPath {
    param([Parameter(Mandatory = $true)][string]$Scope)

    $psexecToolkitPath = Join-Path $toolkitRoot 'Scripts\PsExec.exe'
    $psexecSystem32Path = Join-Path $env:WINDIR 'System32\PsExec.exe'
    $psexecCommand = Get-Command -Name 'PsExec.exe' -CommandType Application -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $psexecToolkitPath -PathType Leaf) { return (Get-Item -LiteralPath $psexecToolkitPath).FullName }
    if (Test-Path -LiteralPath $psexecSystem32Path -PathType Leaf) { return (Get-Item -LiteralPath $psexecSystem32Path).FullName }
    if ($psexecCommand) { return $psexecCommand.Source }
    throw ("PsExec.exe not found. Place it in '{0}', in '{1}', or add PsExec.exe to PATH before launching the {2}." -f (Split-Path -Parent $psexecToolkitPath), (Split-Path -Parent $psexecSystem32Path), $Scope)
}

function Start-ToolkitLot {
    param(
        [Parameter(Mandatory = $true)][string]$LotPath,
        [Parameter(Mandatory = $true)][string]$Mode,
        [int]$GlobalConcurrencyLimit = 15,
        [int]$GlobalConcurrencyLeaseTimeoutMinutes = 0,
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

    $wrapperPath = Join-Path $LotPath $wrapperName
    if (-not (Test-Path -LiteralPath $wrapperPath -PathType Leaf)) {
        throw "LOT wrapper not found: $wrapperPath"
    }

    $psexecToolkitPath = Join-Path $toolkitRoot 'Scripts\PsExec.exe'
    $psexecSystem32Path = Join-Path $env:WINDIR 'System32\PsExec.exe'
    $psexecCommand = Get-Command -Name 'PsExec.exe' -CommandType Application -ErrorAction SilentlyContinue
    if (
        -not (Test-Path -LiteralPath $psexecToolkitPath -PathType Leaf) -and
        -not (Test-Path -LiteralPath $psexecSystem32Path -PathType Leaf) -and
        -not $psexecCommand
    ) {
        throw ("PsExec.exe not found. Place it in '{0}', in '{1}', or add PsExec.exe to PATH before launching the LOT." -f (Split-Path -Parent $psexecToolkitPath), (Split-Path -Parent $psexecSystem32Path))
    }

    if ($GlobalConcurrencyLimit -lt 1) { $GlobalConcurrencyLimit = 1 }
    $commandParts = @(
        (ConvertTo-CmdArgument -Value $wrapperPath),
        '-GlobalConcurrencyLimit',
        [string]$GlobalConcurrencyLimit,
        '-GlobalConcurrencyLeaseTimeoutMinutes',
        [string]$GlobalConcurrencyLeaseTimeoutMinutes
    )
    foreach ($argument in @($AdditionalArguments)) {
        if (-not [string]::IsNullOrWhiteSpace($argument)) {
            $commandParts += (ConvertTo-CmdArgument -Value $argument)
        }
    }

    $effectiveEnvironment = Get-EffectiveLotEnvironment -LotPath $LotPath -EnvironmentVariables $EnvironmentVariables
    $setCommands = @()
    foreach ($name in @($effectiveEnvironment.Keys | Sort-Object)) {
        $value = [string]$effectiveEnvironment[$name]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $setCommands += (ConvertTo-CmdSetCommand -Name $name -Value $value)
        }
    }

    $commandLine = (($setCommands + @($commandParts -join ' ')) -join ' & ')
    Start-Process -FilePath 'cmd.exe' -ArgumentList @('/k', $commandLine) -WorkingDirectory $LotPath -Verb RunAs
}

function Start-ToolkitSingleComputer {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$Mode,
        [int]$GlobalConcurrencyLimit = 15,
        [int]$GlobalConcurrencyLeaseTimeoutMinutes = 0,
        [string[]]$AdditionalArguments = @(),
        [hashtable]$EnvironmentVariables = @{}
    )

    $scriptPath = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "Windows 11 upgrade PsExec launcher not found: $scriptPath"
    }

    $isDryRun = @($AdditionalArguments) -contains '-DryRun'
    $psexecPath = ''
    if (-not $isDryRun) { $psexecPath = Resolve-GuiPsExecPath -Scope 'computer' }

    $run = New-SingleComputerRunContext -ComputerName $ComputerName -ToolkitKey 'Windows11UpgradeToolkit'
    if ($GlobalConcurrencyLimit -lt 1) { $GlobalConcurrencyLimit = 1 }

    $commandParts = @(
        'powershell.exe',
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        (ConvertTo-CmdArgument -Value $scriptPath),
        '-ComputerListPath',
        (ConvertTo-CmdArgument -Value $run.ComputersPath),
        '-LogRoot',
        (ConvertTo-CmdArgument -Value $run.LogRoot),
        '-ReportRoot',
        (ConvertTo-CmdArgument -Value $run.ReportRoot),
        '-CentralLogRoot',
        (ConvertTo-CmdArgument -Value $run.CentralLogRoot),
        '-GlobalConcurrencyLimit',
        [string]$GlobalConcurrencyLimit,
        '-GlobalConcurrencyLeaseTimeoutMinutes',
        [string]$GlobalConcurrencyLeaseTimeoutMinutes
    )
    if (-not [string]::IsNullOrWhiteSpace($psexecPath)) {
        $commandParts += '-PsExecPath'
        $commandParts += (ConvertTo-CmdArgument -Value $psexecPath)
    }

    if ($Mode -in @('Once','OnceIgnoreRunGuard')) { $commandParts += '-RunOnce' }
    if ($Mode -in @('LoopIgnoreRunGuard','OnceIgnoreRunGuard')) { $commandParts += '-IgnoreRunGuard' }

    foreach ($pair in @(
        @{ Name = 'W11UT_AUDIT_ONLY'; Argument = '-AuditOnly' },
        @{ Name = 'W11UT_ALLOW_POLICY_REPAIR'; Argument = '-AllowPolicyRepair' },
        @{ Name = 'W11UT_ALLOW_WU_RESET'; Argument = '-AllowWUReset' },
        @{ Name = 'W11UT_ALLOW_FORCE_UPGRADE'; Argument = '-AllowForceUpgrade' },
        @{ Name = 'W11UT_ALLOW_SETUP_UPGRADE'; Argument = '-AllowSetupUpgrade' },
        @{ Name = 'W11UT_DIRECT_SETUP_UPGRADE'; Argument = '-DirectSetupUpgrade' },
        @{ Name = 'W11UT_ALLOW_REBOOT'; Argument = '-AllowReboot' },
        @{ Name = 'W11UT_SKIP_VIRTUAL_MACHINES'; Argument = '-SkipVirtualMachines' },
        @{ Name = 'W11UT_ALLOW_DISK_CLEANUP'; Argument = '-AllowDiskCleanup' },
        @{ Name = 'W11UT_ALLOW_ADVANCED_DISK_CLEANUP'; Argument = '-AllowAdvancedDiskCleanup' },
        @{ Name = 'W11UT_SKIP_SETUP_MEDIA_PRECOPY'; Argument = '-SkipSetupMediaPreCopy' }
    )) {
        if ([string]$EnvironmentVariables[$pair.Name] -eq '1') { $commandParts += $pair.Argument }
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
        @{ Name = 'W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT'; Argument = '-SetupSourceConcurrencyLimit' },
        @{ Name = 'W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES'; Argument = '-SetupSourceConcurrencyLeaseMinutes' },
        @{ Name = 'W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT'; Argument = '-SetupSourceConcurrencyGateRoot' },
        @{ Name = 'W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS'; Argument = '-DelayBetweenComputersSeconds' },
        @{ Name = 'W11UT_DELAY_BETWEEN_CYCLES_MINUTES'; Argument = '-DelayBetweenCyclesMinutes' },
        @{ Name = 'W11UT_PSEXEC_TIMEOUT_MINUTES'; Argument = '-PsExecTimeoutMinutes' }
    )) {
        $value = [string]$EnvironmentVariables[$pair.Name]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $commandParts += $pair.Argument
            $commandParts += (ConvertTo-CmdArgument -Value $value)
        }
    }

    foreach ($argument in @($AdditionalArguments)) {
        if (-not [string]::IsNullOrWhiteSpace($argument)) {
            $commandParts += (ConvertTo-CmdArgument -Value $argument)
        }
    }

    $commandLine = $commandParts -join ' '
    Start-Process -FilePath 'cmd.exe' -ArgumentList @('/k', $commandLine) -WorkingDirectory $run.RunPath -Verb RunAs
    return $run
}

function Wait-UiDelay {
    param([int]$Seconds)

    if ($Seconds -le 0) {
        return
    }

    $end = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $end) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
}

function Open-TextFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        New-Item -ItemType File -Path $Path -Force -ErrorAction Stop | Out-Null
    }

    Start-Process -FilePath 'notepad.exe' -ArgumentList @("`"$Path`"")
}

function Open-FolderPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "Folder not found: $Path"
    }

    Start-Process -FilePath 'explorer.exe' -ArgumentList @("`"$Path`"")
}

function Open-OrCreateFolderPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }

    Start-Process -FilePath 'explorer.exe' -ArgumentList @("`"$Path`"")
}

$toolkitRoot = Get-ToolkitRoot
$script:ToolkitConfigPath = Join-Path $toolkitRoot 'Windows11UpgradeToolkit.config'
$script:ToolkitConfig = Read-ToolkitConfig -Path $script:ToolkitConfigPath
$script:ToolkitDefaultEnvironment = @{
    W11UT_AUDIT_ONLY = '0'
    W11UT_ALLOW_POLICY_REPAIR = '1'
    W11UT_ALLOW_WU_RESET = '1'
    W11UT_ALLOW_FORCE_UPGRADE = '1'
    W11UT_ALLOW_SETUP_UPGRADE = '1'
    W11UT_DIRECT_SETUP_UPGRADE = '0'
    W11UT_ALLOW_REBOOT = '1'
    W11UT_SKIP_VIRTUAL_MACHINES = '1'
    W11UT_ALLOW_DISK_CLEANUP = '1'
    W11UT_ALLOW_ADVANCED_DISK_CLEANUP = '0'
    W11UT_SKIP_SETUP_MEDIA_PRECOPY = '0'
    W11UT_SETUP_SOURCE_MAP = ''
    W11UT_SETUP_SOURCE_CANDIDATE_LIMIT = '5'
    W11UT_SETUP_COPY_IPG_MS = '0'
    W11UT_SETUP_COPY_JITTER_SECONDS = '0'
    W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT = '0'
    W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES = '240'
    W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT = ''
    W11UT_SETUP_EXECUTION_MODE = 'LocalCache'
    W11UT_SETUP_MEDIA_ID = 'Win11'
    W11UT_SETUP_LANGUAGE = 'MatchSystem'
    W11UT_SETUP_DYNAMIC_UPDATE = 'Disable'
    W11UT_THROTTLE = '10'
    W11UT_GLOBAL_CONCURRENCY_LIMIT = '15'
    W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = '0'
    W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS = '0'
    W11UT_DELAY_BETWEEN_CYCLES_MINUTES = '5'
    W11UT_PSEXEC_TIMEOUT_MINUTES = '180'
}
$launchAllLotStartDelaySeconds = 5

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$font = New-Object System.Drawing.Font('Segoe UI', 9)
$titleFont = New-Object System.Drawing.Font('Segoe UI Semibold', 19)
$sectionFont = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$smallFont = New-Object System.Drawing.Font('Segoe UI', 8)
$statusFont = New-Object System.Drawing.Font('Consolas', 9)

$colorBackground = [System.Drawing.ColorTranslator]::FromHtml('#F5F8FB')
$colorPanel = [System.Drawing.ColorTranslator]::FromHtml('#FFFFFF')
$colorPanelSoft = [System.Drawing.ColorTranslator]::FromHtml('#FAFCFE')
$colorHeaderPanel = [System.Drawing.ColorTranslator]::FromHtml('#E6F4FF')
$colorAccent = [System.Drawing.ColorTranslator]::FromHtml('#0078D4')
$colorAccentDark = [System.Drawing.ColorTranslator]::FromHtml('#005A9E')
$colorInk = [System.Drawing.ColorTranslator]::FromHtml('#1F2937')
$colorMuted = [System.Drawing.ColorTranslator]::FromHtml('#5F6B7A')
$colorBorder = [System.Drawing.ColorTranslator]::FromHtml('#DDE7F0')
$colorTextBoxBorder = [System.Drawing.ColorTranslator]::FromHtml('#B9C8D7')
$colorSuccess = [System.Drawing.ColorTranslator]::FromHtml('#027A48')
$colorDisabled = [System.Drawing.ColorTranslator]::FromHtml('#E4EAF1')
$colorDisabledText = [System.Drawing.ColorTranslator]::FromHtml('#7A8A99')

function Resolve-LogoIconPath {
    $candidatePaths = @(
        (Join-Path $PSScriptRoot 'SmartM365-logo.ico'),
        (Join-Path $toolkitRoot 'SmartM365-logo.ico'),
        (Join-Path $toolkitRoot 'Scripts\SmartM365-logo.ico')
    )

    foreach ($candidatePath in $candidatePaths) {
        if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
            return (Get-Item -LiteralPath $candidatePath -ErrorAction Stop).FullName
        }
    }

    return $null
}

function Resolve-LogoImagePath {
    $devicesRoot = Split-Path -Parent $toolkitRoot
    $candidatePaths = @(
        (Join-Path $toolkitRoot 'workplacecloudhub-v2.png'),
        (Join-Path $PSScriptRoot 'workplacecloudhub-v2.png'),
        (Join-Path $devicesRoot 'DeviceRegistrationTool\workplacecloudhub-v2.png')
    )

    foreach ($candidatePath in $candidatePaths) {
        if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
            return (Get-Item -LiteralPath $candidatePath -ErrorAction Stop).FullName
        }
    }

    return $null
}

function Set-FlatButtonStyle {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory = $true)][System.Drawing.Color]$BackColor,
        [Parameter(Mandatory = $true)][System.Drawing.Color]$ForeColor,
        [System.Drawing.Color]$BorderColor = $colorBorder
    )

    $Button.FlatStyle = 'Flat'
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.FlatAppearance.BorderColor = $BorderColor
    $Button.FlatAppearance.BorderSize = 1
    $Button.Height = 34
    $Button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function New-RoundedRectanglePath {
    param(
        [Parameter(Mandatory = $true)][System.Drawing.Rectangle]$Rectangle,
        [int]$Radius = 8
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $diameter = [Math]::Max(1, $Radius * 2)
    $path.AddArc($Rectangle.X, $Rectangle.Y, $diameter, $diameter, 180, 90)
    $path.AddArc($Rectangle.Right - $diameter, $Rectangle.Y, $diameter, $diameter, 270, 90)
    $path.AddArc($Rectangle.Right - $diameter, $Rectangle.Bottom - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($Rectangle.X, $Rectangle.Bottom - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function Add-SoftBorder {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Control]$Control,
        [System.Drawing.Color]$BorderColor = $colorBorder,
        [int]$Radius = 8
    )

    $Control.Add_Paint({
        param($sender, $eventArgs)

        $eventArgs.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $rect = New-Object System.Drawing.Rectangle(0, 0, ($sender.Width - 1), ($sender.Height - 1))
        $path = New-RoundedRectanglePath -Rectangle $rect -Radius $Radius
        $pen = New-Object System.Drawing.Pen($BorderColor, 1)
        try {
            $eventArgs.Graphics.DrawPath($pen, $path)
        }
        finally {
            $pen.Dispose()
            $path.Dispose()
        }
    }.GetNewClosure())
}

function Show-DeviceRegistrationTabPage {
    param([Parameter(Mandatory = $true)][System.Windows.Forms.Control]$Header)

    foreach ($tabHeader in $script:DeviceRegistrationTabHeaders) {
        $state = $tabHeader.Tag
        $isSelected = [object]::ReferenceEquals($tabHeader, $Header)
        $state.Selected = $isSelected
        $state.Page.Visible = $isSelected
        if ($isSelected) {
            $state.Page.BringToFront()
        }
        $tabHeader.Invalidate()
    }
}

function New-DeviceRegistrationTabHeader {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][System.Windows.Forms.Control]$Page
    )

    $tabTextFlags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor
        [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
        [System.Windows.Forms.TextFormatFlags]::EndEllipsis

    $header = New-Object System.Windows.Forms.Panel
    $header.Width = 122
    $header.Height = 34
    $header.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    $header.Cursor = [System.Windows.Forms.Cursors]::Hand
    $header.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $header.Tag = [pscustomobject]@{
        Text = $Text
        Page = $Page
        Selected = $false
        Hover = $false
    }

    $header.Add_Paint({
        param($sender, $eventArgs)

        $eventArgs.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $state = $sender.Tag
        $tabRect = New-Object System.Drawing.Rectangle(0, 0, ($sender.Width - 1), ($sender.Height - 1))
        $fillColor = if ($state.Selected) {
            $colorHeaderPanel
        }
        elseif ($state.Hover) {
            [System.Drawing.ColorTranslator]::FromHtml('#F8FBFE')
        }
        else {
            $colorPanel
        }
        $borderColor = if ($state.Selected -or $state.Hover) {
            [System.Drawing.ColorTranslator]::FromHtml('#B9DDF7')
        }
        else {
            $colorBorder
        }
        $textColor = if ($state.Selected) { $colorAccentDark } else { $colorMuted }

        $path = New-RoundedRectanglePath -Rectangle $tabRect -Radius 8
        $brush = New-Object System.Drawing.SolidBrush($fillColor)
        $pen = New-Object System.Drawing.Pen($borderColor, 1)
        try {
            $eventArgs.Graphics.FillPath($brush, $path)
            $eventArgs.Graphics.DrawPath($pen, $path)
            [System.Windows.Forms.TextRenderer]::DrawText(
                $eventArgs.Graphics,
                [string]$state.Text,
                $sender.Font,
                $tabRect,
                $textColor,
                $tabTextFlags
            )
        }
        finally {
            $pen.Dispose()
            $brush.Dispose()
            $path.Dispose()
        }
    }.GetNewClosure())

    $header.Add_MouseEnter({
        param($sender, $eventArgs)
        $sender.Tag.Hover = $true
        $sender.Invalidate()
    })
    $header.Add_MouseLeave({
        param($sender, $eventArgs)
        $sender.Tag.Hover = $false
        $sender.Invalidate()
    })
    $header.Add_Click({
        param($sender, $eventArgs)
        Show-DeviceRegistrationTabPage -Header $sender
    })

    return $header
}


function Add-AccentBar {
    param([Parameter(Mandatory = $true)][System.Windows.Forms.Control]$Control)

    $Control.Add_Paint({
        param($sender, $eventArgs)
    }.GetNewClosure())
}

function New-Label {
    param([string]$Text)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Dock = 'Fill'
    $label.TextAlign = 'MiddleLeft'
    $label.ForeColor = $colorMuted
    return $label
}

function New-ValueBox {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = 'Fill'
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    $box.ReadOnly = $true
    $box.BorderStyle = 'FixedSingle'
    $box.BackColor = $colorPanelSoft
    $box.ForeColor = $colorInk
    return $box
}

function New-EntryBox {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = 'Fill'
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    $box.BorderStyle = 'FixedSingle'
    $box.BackColor = $colorPanel
    $box.ForeColor = $colorInk
    return $box
}

function New-OptionsTextBox {
    param([Parameter(Mandatory = $true)][string]$Text)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = 'Fill'
    $box.Multiline = $true
    $box.ScrollBars = 'Vertical'
    $box.ReadOnly = $true
    $box.Font = $statusFont
    $box.BackColor = $colorPanelSoft
    $box.ForeColor = $colorInk
    $box.BorderStyle = 'FixedSingle'
    $box.Text = $Text.Trim()
    return $box
}

function New-SectionPanel {
    param([Parameter(Mandatory = $true)][string]$Title)

    $outer = New-Object System.Windows.Forms.Panel
    $outer.Dock = 'Fill'
    $outer.BackColor = $colorPanel
    $outer.Padding = New-Object System.Windows.Forms.Padding(14)
    Add-SoftBorder -Control $outer

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.ColumnCount = 1
    $layout.RowCount = 2
    $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.Dock = 'Fill'
    $titleLabel.Font = $sectionFont
    $titleLabel.ForeColor = $colorInk
    $titleLabel.TextAlign = 'MiddleLeft'

    $content = New-Object System.Windows.Forms.Panel
    $content.Dock = 'Fill'
    $content.BackColor = $colorPanel

    $layout.Controls.Add($titleLabel, 0, 0)
    $layout.Controls.Add($content, 0, 1)
    $outer.Controls.Add($layout)

    return [pscustomobject]@{
        Panel = $outer
        Content = $content
    }
}

function Set-ButtonEnabledStyle {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.Button]$Button,
        [bool]$Enabled,
        [System.Drawing.Color]$EnabledBackColor = $colorAccent,
        [System.Drawing.Color]$EnabledForeColor = ([System.Drawing.Color]::White)
    )

    if ($Enabled) {
        $Button.Enabled = $true
        Set-FlatButtonStyle -Button $Button -BackColor $EnabledBackColor -ForeColor $EnabledForeColor -BorderColor $EnabledBackColor
    }
    else {
        $Button.Enabled = $false
        Set-FlatButtonStyle -Button $Button -BackColor $colorDisabled -ForeColor $colorDisabledText -BorderColor $colorBorder
        $Button.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Smart Intune Windows 11 Upgrade Toolkit - LOT Launcher'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1240, 860)
$form.MinimumSize = New-Object System.Drawing.Size(1040, 760)
$form.Font = $font
$form.BackColor = $colorBackground

$logoIconPath = Resolve-LogoIconPath
$logoImagePath = Resolve-LogoImagePath
$script:LogoIcon = $null
$script:LogoImage = $null
if (-not [string]::IsNullOrWhiteSpace($logoIconPath)) {
    try {
        $script:LogoIcon = New-Object System.Drawing.Icon($logoIconPath, 48, 48)
        $form.Icon = $script:LogoIcon
    }
    catch {
        $script:LogoIcon = $null
    }
}

if (-not [string]::IsNullOrWhiteSpace($logoImagePath)) {
    try {
        $script:LogoImage = [System.Drawing.Image]::FromFile($logoImagePath)
    }
    catch {
        $script:LogoImage = $null
    }
}

$rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rootLayout.Dock = 'Fill'
$rootLayout.ColumnCount = 1
$rootLayout.RowCount = 4
$rootLayout.Padding = New-Object System.Windows.Forms.Padding(18)
$rootLayout.BackColor = $colorBackground
$rootLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 138))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 96))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54))) | Out-Null

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Fill'
$headerPanel.BackColor = $colorPanel
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(18)
Add-SoftBorder -Control $headerPanel

$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerLayout.Dock = 'Fill'
$headerLayout.ColumnCount = 2
$headerLayout.RowCount = 1
$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 142))) | Out-Null

$headerTextPanel = New-Object System.Windows.Forms.TableLayoutPanel
$headerTextPanel.Dock = 'Fill'
$headerTextPanel.ColumnCount = 1
$headerTextPanel.RowCount = 4
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null

$badgeLabel = New-Object System.Windows.Forms.Label
$badgeLabel.Text = 'SMARTM365'
$badgeLabel.AutoSize = $true
$badgeLabel.BackColor = $colorHeaderPanel
$badgeLabel.ForeColor = $colorAccentDark
$badgeLabel.Font = $smallFont
$badgeLabel.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'Windows 11 Upgrade LOT Launcher'
$titleLabel.Dock = 'Fill'
$titleLabel.Font = $titleFont
$titleLabel.ForeColor = $colorInk
$titleLabel.TextAlign = 'MiddleLeft'

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = 'Run an existing LOT or create a new empty LOT ready for Computers.txt.'
$subtitleLabel.Dock = 'Fill'
$subtitleLabel.ForeColor = $colorMuted
$subtitleLabel.TextAlign = 'MiddleLeft'

$psexecLabel = New-Object System.Windows.Forms.Label
$psexecLabel.Text = 'Diagnose Windows 10 blockers, repair update policy, and run guarded Windows 11 upgrades.'
$psexecLabel.Dock = 'Fill'
$psexecLabel.ForeColor = $colorMuted
$psexecLabel.Font = $smallFont
$psexecLabel.TextAlign = 'MiddleLeft'

$headerTextPanel.Controls.Add($badgeLabel, 0, 0)
$headerTextPanel.Controls.Add($titleLabel, 0, 1)
$headerTextPanel.Controls.Add($subtitleLabel, 0, 2)
$headerTextPanel.Controls.Add($psexecLabel, 0, 3)

$logoCard = New-Object System.Windows.Forms.Panel
$logoCard.Dock = 'Fill'
$logoCard.BackColor = $colorPanelSoft
$logoCard.Padding = New-Object System.Windows.Forms.Padding(10)
$logoCard.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
Add-SoftBorder -Control $logoCard

if ($script:LogoImage) {
    $logoPicture = New-Object System.Windows.Forms.PictureBox
    $logoPicture.Dock = 'Fill'
    $logoPicture.SizeMode = 'Zoom'
    $logoPicture.Image = $script:LogoImage
    $logoCard.Controls.Add($logoPicture)
}
else {
    $logoFallback = New-Object System.Windows.Forms.Label
    $logoFallback.Text = 'WorkplaceCloudHub'
    $logoFallback.Dock = 'Fill'
    $logoFallback.TextAlign = 'MiddleCenter'
    $logoFallback.ForeColor = $colorAccent
    $logoFallback.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
    $logoCard.Controls.Add($logoFallback)
}

$headerLayout.Controls.Add($headerTextPanel, 0, 0)
$headerLayout.Controls.Add($logoCard, 1, 0)
$headerPanel.Controls.Add($headerLayout)

$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Dock = 'Fill'
$actionPanel.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)
$actionPanel.BackColor = $colorPanel
$actionPanel.Padding = New-Object System.Windows.Forms.Padding(12)
Add-SoftBorder -Control $actionPanel

$actionLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionLayout.Dock = 'Fill'
$actionLayout.ColumnCount = 3
$actionLayout.RowCount = 1
$actionLayout.BackColor = $colorPanel
$actionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 22))) | Out-Null
$actionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$actionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 328))) | Out-Null

$actionDot = New-Object System.Windows.Forms.Panel
$actionDot.Width = 10
$actionDot.Height = 10
$actionDot.Margin = New-Object System.Windows.Forms.Padding(0, 24, 10, 0)
$actionDot.BackColor = $colorAccent
Add-SoftBorder -Control $actionDot -BorderColor $colorAccent -Radius 5

$actionTextPanel = New-Object System.Windows.Forms.TableLayoutPanel
$actionTextPanel.Dock = 'Fill'
$actionTextPanel.ColumnCount = 1
$actionTextPanel.RowCount = 3
$actionTextPanel.BackColor = $colorPanel
$actionTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$actionTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) | Out-Null
$actionTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$actionTitleLabel = New-Object System.Windows.Forms.Label
$actionTitleLabel.Dock = 'Fill'
$actionTitleLabel.Text = 'LOT: none selected'
$actionTitleLabel.ForeColor = $colorInk
$actionTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
$actionTitleLabel.TextAlign = 'MiddleLeft'

$actionSubtitleLabel = New-Object System.Windows.Forms.Label
$actionSubtitleLabel.Dock = 'Fill'
$actionSubtitleLabel.Text = 'Select or create a LOT.'
$actionSubtitleLabel.ForeColor = $colorAccent
$actionSubtitleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 8)
$actionSubtitleLabel.TextAlign = 'MiddleLeft'

$actionStatusLabel = New-Object System.Windows.Forms.Label
$actionStatusLabel.Dock = 'Fill'
$actionStatusLabel.Text = 'Ready'
$actionStatusLabel.ForeColor = $colorMuted
$actionStatusLabel.Font = $smallFont
$actionStatusLabel.TextAlign = 'MiddleLeft'

$actionTextPanel.Controls.Add($actionTitleLabel, 0, 0)
$actionTextPanel.Controls.Add($actionSubtitleLabel, 0, 1)
$actionTextPanel.Controls.Add($actionStatusLabel, 0, 2)

$actionButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionButtonsPanel.Dock = 'Fill'
$actionButtonsPanel.FlowDirection = 'RightToLeft'
$actionButtonsPanel.WrapContents = $false
$actionButtonsPanel.BackColor = $colorPanel
$actionButtonsPanel.Padding = New-Object System.Windows.Forms.Padding(0, 18, 0, 0)

$actionLaunchAllButton = New-Object System.Windows.Forms.Button
$actionLaunchAllButton.Text = 'Launch all'
$actionLaunchAllButton.Width = 130
Set-ButtonEnabledStyle -Button $actionLaunchAllButton -Enabled $false -EnabledBackColor $colorAccent

$actionRefreshButton = New-Object System.Windows.Forms.Button
$actionRefreshButton.Text = 'Refresh'
$actionRefreshButton.Width = 118
Set-FlatButtonStyle -Button $actionRefreshButton -BackColor $colorPanel -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$actionButtonsPanel.Controls.Add($actionLaunchAllButton)
$actionButtonsPanel.Controls.Add($actionRefreshButton)

$actionLayout.Controls.Add($actionDot, 0, 0)
$actionLayout.Controls.Add($actionTextPanel, 1, 0)
$actionLayout.Controls.Add($actionButtonsPanel, 2, 0)
$actionPanel.Controls.Add($actionLayout)

$tabShell = New-Object System.Windows.Forms.TableLayoutPanel
$tabShell.Dock = 'Fill'
$tabShell.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 12)
$tabShell.BackColor = $colorBackground
$tabShell.ColumnCount = 1
$tabShell.RowCount = 2
$tabShell.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$tabShell.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) | Out-Null
$tabShell.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$tabStrip = New-Object System.Windows.Forms.FlowLayoutPanel
$tabStrip.Dock = 'Fill'
$tabStrip.FlowDirection = 'LeftToRight'
$tabStrip.WrapContents = $false
$tabStrip.BackColor = $colorBackground
$tabStrip.Padding = New-Object System.Windows.Forms.Padding(0)
$tabStrip.Margin = New-Object System.Windows.Forms.Padding(0)

$tabContentPanel = New-Object System.Windows.Forms.Panel
$tabContentPanel.Dock = 'Fill'
$tabContentPanel.BackColor = $colorBackground
$tabContentPanel.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)

$tabShell.Controls.Add($tabStrip, 0, 0)
$tabShell.Controls.Add($tabContentPanel, 0, 1)

$script:DeviceRegistrationTabHeaders = New-Object System.Collections.ArrayList

$existingTab = New-Object System.Windows.Forms.Panel
$existingTab.Dock = 'Fill'
$existingTab.BackColor = $colorBackground

$singleTab = New-Object System.Windows.Forms.Panel
$singleTab.Dock = 'Fill'
$singleTab.BackColor = $colorBackground

$newTab = New-Object System.Windows.Forms.Panel
$newTab.Dock = 'Fill'
$newTab.BackColor = $colorBackground

$optionsTab = New-Object System.Windows.Forms.Panel
$optionsTab.Dock = 'Fill'
$optionsTab.BackColor = $colorBackground

$tabContentPanel.Controls.Add($existingTab)
$tabContentPanel.Controls.Add($singleTab)
$tabContentPanel.Controls.Add($newTab)
$tabContentPanel.Controls.Add($optionsTab)

$existingTabHeader = New-DeviceRegistrationTabHeader -Text 'Existing LOT' -Page $existingTab
$singleTabHeader = New-DeviceRegistrationTabHeader -Text 'Single PC' -Page $singleTab
$newTabHeader = New-DeviceRegistrationTabHeader -Text 'New LOT' -Page $newTab
$optionsTabHeader = New-DeviceRegistrationTabHeader -Text 'Options' -Page $optionsTab
$script:DeviceRegistrationTabHeaders.Add($existingTabHeader) | Out-Null
$script:DeviceRegistrationTabHeaders.Add($singleTabHeader) | Out-Null
$script:DeviceRegistrationTabHeaders.Add($newTabHeader) | Out-Null
$script:DeviceRegistrationTabHeaders.Add($optionsTabHeader) | Out-Null
$tabStrip.Controls.Add($existingTabHeader)
$tabStrip.Controls.Add($singleTabHeader)
$tabStrip.Controls.Add($newTabHeader)
$tabStrip.Controls.Add($optionsTabHeader)
Show-DeviceRegistrationTabPage -Header $existingTabHeader

$existingLayout = New-Object System.Windows.Forms.TableLayoutPanel
$existingLayout.Dock = 'Fill'
$existingLayout.ColumnCount = 1
$existingLayout.RowCount = 2
$existingLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$existingLayout.BackColor = $colorBackground
$existingLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$existingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 232))) | Out-Null
$existingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$existingSection = New-SectionPanel -Title 'Available lots'
$activitySection = New-SectionPanel -Title 'Activity'

$existingTable = New-Object System.Windows.Forms.TableLayoutPanel
$existingTable.Dock = 'Fill'
$existingTable.ColumnCount = 5
$existingTable.RowCount = 5
$existingTable.BackColor = $colorPanel
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 138))) | Out-Null
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 118))) | Out-Null
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 118))) | Out-Null
$existingTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 118))) | Out-Null
for ($i = 0; $i -lt 5; $i++) {
    $existingTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34))) | Out-Null
}

$lotCombo = New-Object System.Windows.Forms.ComboBox
$lotCombo.Dock = 'Fill'
$lotCombo.DropDownStyle = 'DropDownList'
$lotCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$lotCombo.FlatStyle = 'Flat'

$refreshLotsButton = New-Object System.Windows.Forms.Button
$refreshLotsButton.Text = 'Refresh'
$refreshLotsButton.Dock = 'Fill'
Set-FlatButtonStyle -Button $refreshLotsButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$openLotFolderButton = New-Object System.Windows.Forms.Button
$openLotFolderButton.Text = 'Folder'
$openLotFolderButton.Dock = 'Fill'
Set-FlatButtonStyle -Button $openLotFolderButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$deviceCountBox = New-ValueBox
$upgradeScopeBox = New-ValueBox

$existingModeCombo = New-Object System.Windows.Forms.ComboBox
$existingModeCombo.Dock = 'Fill'
$existingModeCombo.DropDownStyle = 'DropDownList'
$existingModeCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$existingModeCombo.FlatStyle = 'Flat'
$existingModeCombo.Items.AddRange([object[]]@('Once', 'Loop', 'OnceIgnoreRunGuard', 'LoopIgnoreRunGuard'))
$existingModeCombo.SelectedIndex = 0

$globalConcurrencyLimitBox = New-Object System.Windows.Forms.NumericUpDown
$globalConcurrencyLimitBox.Dock = 'Fill'
$globalConcurrencyLimitBox.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$globalConcurrencyLimitBox.Minimum = 1
$globalConcurrencyLimitBox.Maximum = 200
$globalConcurrencyLimitBox.Value = 15

$openComputersButton = New-Object System.Windows.Forms.Button
$openComputersButton.Text = 'Computers'
$openComputersButton.Dock = 'Fill'
Set-FlatButtonStyle -Button $openComputersButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$openReportsButton = New-Object System.Windows.Forms.Button
$openReportsButton.Text = 'Reports'
$openReportsButton.Dock = 'Fill'
Set-FlatButtonStyle -Button $openReportsButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$launchExistingButton = New-Object System.Windows.Forms.Button
$launchExistingButton.Text = 'Launch'
$launchExistingButton.Dock = 'Fill'
Set-ButtonEnabledStyle -Button $launchExistingButton -Enabled $false -EnabledBackColor $colorAccent

$launchAllLotsButton = New-Object System.Windows.Forms.Button
$launchAllLotsButton.Text = 'Launch all'
$launchAllLotsButton.Dock = 'Fill'
Set-ButtonEnabledStyle -Button $launchAllLotsButton -Enabled $false -EnabledBackColor $colorAccent

$refreshWrappersButton = New-Object System.Windows.Forms.Button
$refreshWrappersButton.Text = 'Wrappers'
$refreshWrappersButton.Dock = 'Fill'
Set-FlatButtonStyle -Button $refreshWrappersButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$existingTable.Controls.Add((New-Label 'LOT'), 0, 0)
$existingTable.Controls.Add($lotCombo, 1, 0)
$existingTable.SetColumnSpan($lotCombo, 2)
$existingTable.Controls.Add($refreshLotsButton, 3, 0)
$existingTable.Controls.Add($openLotFolderButton, 4, 0)

$existingTable.Controls.Add((New-Label 'Devices'), 0, 1)
$existingTable.Controls.Add($deviceCountBox, 1, 1)
$existingTable.SetColumnSpan($deviceCountBox, 2)
$existingTable.Controls.Add($openComputersButton, 3, 1)
$existingTable.Controls.Add($openReportsButton, 4, 1)

$existingTable.Controls.Add((New-Label 'Scope'), 0, 2)
$existingTable.Controls.Add($upgradeScopeBox, 1, 2)
$existingTable.SetColumnSpan($upgradeScopeBox, 4)

$existingTable.Controls.Add((New-Label 'Limit'), 0, 3)
$existingTable.Controls.Add($globalConcurrencyLimitBox, 1, 3)
$existingTable.SetColumnSpan($globalConcurrencyLimitBox, 4)

$existingTable.Controls.Add((New-Label 'Launch'), 0, 4)
$existingTable.Controls.Add($existingModeCombo, 1, 4)
$existingTable.Controls.Add($refreshWrappersButton, 2, 4)
$existingTable.Controls.Add($launchExistingButton, 3, 4)
$existingTable.SetColumnSpan($launchExistingButton, 2)

$existingSection.Content.Controls.Add($existingTable)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Dock = 'Fill'
$statusBox.Multiline = $true
$statusBox.ScrollBars = 'Vertical'
$statusBox.ReadOnly = $true
$statusBox.Font = $statusFont
$statusBox.BackColor = $colorPanelSoft
$statusBox.ForeColor = $colorInk
$statusBox.BorderStyle = 'FixedSingle'
$activitySection.Content.Controls.Add($statusBox)

$existingLayout.Controls.Add($existingSection.Panel, 0, 0)
$existingLayout.Controls.Add($activitySection.Panel, 0, 1)
$existingTab.Controls.Add($existingLayout)

$singleLayout = New-Object System.Windows.Forms.TableLayoutPanel
$singleLayout.Dock = 'Fill'
$singleLayout.ColumnCount = 1
$singleLayout.RowCount = 2
$singleLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$singleLayout.BackColor = $colorBackground
$singleLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$singleLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 174))) | Out-Null
$singleLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$singleSection = New-SectionPanel -Title 'Run one computer'
$singleTable = New-Object System.Windows.Forms.TableLayoutPanel
$singleTable.Dock = 'Top'
$singleTable.Height = 108
$singleTable.ColumnCount = 4
$singleTable.RowCount = 3
$singleTable.BackColor = $colorPanel
$singleTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 138))) | Out-Null
$singleTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$singleTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150))) | Out-Null
$singleTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150))) | Out-Null
for ($i = 0; $i -lt 3; $i++) {
    $singleTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) | Out-Null
}

$singleComputerBox = New-EntryBox

$singleModeCombo = New-Object System.Windows.Forms.ComboBox
$singleModeCombo.Dock = 'Fill'
$singleModeCombo.DropDownStyle = 'DropDownList'
$singleModeCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)
$singleModeCombo.FlatStyle = 'Flat'
$singleModeCombo.Items.AddRange([object[]]@('Once','OnceIgnoreRunGuard','Loop','LoopIgnoreRunGuard'))
$singleModeCombo.SelectedIndex = 0

$singleRunPathBox = New-ValueBox
$script:SingleComputerRunRoot = Get-SingleComputerRunRoot -ToolkitKey 'Windows11UpgradeToolkit'
$singleRunPathBox.Text = $script:SingleComputerRunRoot

$openSingleRunFolderButton = New-Object System.Windows.Forms.Button
$openSingleRunFolderButton.Text = 'Open run folder'
$openSingleRunFolderButton.Dock = 'Fill'
Set-FlatButtonStyle -Button $openSingleRunFolderButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$launchSingleComputerButton = New-Object System.Windows.Forms.Button
$launchSingleComputerButton.Text = 'Launch'
$launchSingleComputerButton.Dock = 'Fill'
Set-ButtonEnabledStyle -Button $launchSingleComputerButton -Enabled $true -EnabledBackColor $colorAccent

$singleTable.Controls.Add((New-Label 'Computer'), 0, 0)
$singleTable.Controls.Add($singleComputerBox, 1, 0)
$singleTable.Controls.Add((New-Label 'Mode'), 2, 0)
$singleTable.Controls.Add($singleModeCombo, 3, 0)

$singleTable.Controls.Add((New-Label 'Run folder'), 0, 1)
$singleTable.Controls.Add($singleRunPathBox, 1, 1)
$singleTable.SetColumnSpan($singleRunPathBox, 2)
$singleTable.Controls.Add($openSingleRunFolderButton, 3, 1)

$singleTable.Controls.Add((New-Label 'Launch'), 0, 2)
$singleTable.Controls.Add($launchSingleComputerButton, 1, 2)
$singleTable.SetColumnSpan($launchSingleComputerButton, 3)

$singleSection.Content.Controls.Add($singleTable)
$singleLayout.Controls.Add($singleSection.Panel, 0, 0)
$singleTab.Controls.Add($singleLayout)

$newLayout = New-Object System.Windows.Forms.TableLayoutPanel
$newLayout.Dock = 'Fill'
$newLayout.ColumnCount = 1
$newLayout.RowCount = 2
$newLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$newLayout.BackColor = $colorBackground
$newLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$newLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 174))) | Out-Null
$newLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$newSection = New-SectionPanel -Title 'Create an empty LOT'
$newTable = New-Object System.Windows.Forms.TableLayoutPanel
$newTable.Dock = 'Top'
$newTable.Height = 72
$newTable.ColumnCount = 3
$newTable.RowCount = 2
$newTable.BackColor = $colorPanel
$newTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 138))) | Out-Null
$newTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$newTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 150))) | Out-Null
for ($i = 0; $i -lt 2; $i++) {
    $newTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) | Out-Null
}

$newLotNameBox = New-EntryBox
$newComputersPathBox = New-ValueBox

$createLotButton = New-Object System.Windows.Forms.Button
$createLotButton.Text = 'Create'
$createLotButton.Dock = 'Fill'
Set-FlatButtonStyle -Button $createLotButton -BackColor $colorAccent -ForeColor ([System.Drawing.Color]::White) -BorderColor $colorAccent

$openNewComputersButton = New-Object System.Windows.Forms.Button
$openNewComputersButton.Text = 'Open Computers'
$openNewComputersButton.Dock = 'Fill'
Set-ButtonEnabledStyle -Button $openNewComputersButton -Enabled $false

$newTable.Controls.Add((New-Label 'LOT name'), 0, 0)
$newTable.Controls.Add($newLotNameBox, 1, 0)
$newTable.Controls.Add($createLotButton, 2, 0)

$newTable.Controls.Add((New-Label 'Computers.txt'), 0, 1)
$newTable.Controls.Add($newComputersPathBox, 1, 1)
$newTable.Controls.Add($openNewComputersButton, 2, 1)

$newSection.Content.Controls.Add($newTable)
$newLayout.Controls.Add($newSection.Panel, 0, 0)
$newTab.Controls.Add($newLayout)

$optionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$optionsLayout.Dock = 'Fill'
$optionsLayout.ColumnCount = 1
$optionsLayout.RowCount = 1
$optionsLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$optionsLayout.BackColor = $colorBackground
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$optionsSection = New-SectionPanel -Title 'Toolkit options'
$optionsScrollPanel = New-Object System.Windows.Forms.Panel
$optionsScrollPanel.Dock = 'Fill'
$optionsScrollPanel.AutoScroll = $true
$optionsScrollPanel.BackColor = $colorPanel

$optionsTable = New-Object System.Windows.Forms.TableLayoutPanel
$optionsTable.Dock = 'Top'
$optionsTable.AutoSize = $true
$optionsTable.ColumnCount = 4
$optionsTable.RowCount = 14
$optionsTable.BackColor = $colorPanel
$optionsTable.Padding = New-Object System.Windows.Forms.Padding(0, 2, 12, 2)
$optionsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 190))) | Out-Null
$optionsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 190))) | Out-Null
$optionsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 190))) | Out-Null
$optionsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
for ($i = 0; $i -lt 11; $i++) {
    $optionsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) | Out-Null
}

function New-OptionNumber {
    param([int]$Minimum, [int]$Maximum, [int]$Value)

    $number = New-Object System.Windows.Forms.NumericUpDown
    $number.Dock = 'Fill'
    $number.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)
    $number.Minimum = $Minimum
    $number.Maximum = $Maximum
    $number.Value = $Value
    return $number
}

function New-OptionCheck {
    param([string]$Text, [bool]$Checked = $false)

    $check = New-Object System.Windows.Forms.CheckBox
    $check.Text = $Text
    $check.Dock = 'Fill'
    $check.Checked = $Checked
    $check.ForeColor = $colorInk
    return $check
}

function Get-EnvSwitch {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [bool]$Default = $false
    )

    $value = Get-ToolkitConfigValue -Name $Name
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }

    if ($value -match '^(?i:1|true|yes|on)$') { return $true }
    if ($value -match '^(?i:0|false|no|off)$') { return $false }
    return $Default
}

$script:SetupSourcePlaceholder = '\\server\share\Windows11'

function Test-IsSetupSourcePlaceholder {
    param([AllowNull()][string]$Value)
    return ([string]$Value -eq $script:SetupSourcePlaceholder)
}

function Set-SetupSourcePlaceholder {
    if ([string]::IsNullOrWhiteSpace($optionSetupSourceBox.Text)) {
        $optionSetupSourceBox.Text = $script:SetupSourcePlaceholder
        $optionSetupSourceBox.ForeColor = $colorMuted
    }
}

function Get-SetupSourceText {
    $value = $optionSetupSourceBox.Text.Trim()
    if (Test-IsSetupSourcePlaceholder -Value $value) { return '' }
    return $value
}

function Test-IsUncPath {
    param([AllowNull()][string]$Path)
    return ([string]$Path -match '^\\\\[^\\]+\\[^\\]+')
}

function Test-IsLocalOrRelativePath {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (Test-IsUncPath -Path $Path) { return $false }
    return $true
}

function Split-SetupSourceText {
    param([AllowNull()][string]$Value)

    $sources = New-Object System.Collections.ArrayList
    foreach ($item in @([string]$Value -split ';')) {
        $source = ([string]$item).Trim().Trim([char]34)
        if (-not [string]::IsNullOrWhiteSpace($source)) {
            [void]$sources.Add($source)
        }
    }

    return @($sources.ToArray())
}

function Set-OptionNumberFromEnv {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.NumericUpDown]$Control,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $value = Get-ToolkitConfigValue -Name $Name
    $parsed = 0
    if ([int]::TryParse($value, [ref]$parsed)) {
        if ($parsed -ge [int]$Control.Minimum -and $parsed -le [int]$Control.Maximum) {
            $Control.Value = $parsed
        }
    }
}

$optionAuditOnlyCheck = New-OptionCheck -Text 'Audit only'
$optionAllowPolicyRepairCheck = New-OptionCheck -Text 'Allow policy repair'
$optionAllowWUResetCheck = New-OptionCheck -Text 'Allow WU reset'
$optionAllowForceUpgradeCheck = New-OptionCheck -Text 'Allow force upgrade'
$optionAllowSetupUpgradeCheck = New-OptionCheck -Text 'Allow setup upgrade'
$optionDirectSetupUpgradeCheck = New-OptionCheck -Text 'Direct setup upgrade'
$optionAllowRebootCheck = New-OptionCheck -Text 'Allow reboot'
$optionSkipVirtualMachinesCheck = New-OptionCheck -Text 'Skip virtual machines'
$optionAllowDiskCleanupCheck = New-OptionCheck -Text 'Allow disk cleanup'
$optionAllowAdvancedCleanupCheck = New-OptionCheck -Text 'Allow advanced cleanup'
$optionSkipSetupPreCopyCheck = New-OptionCheck -Text 'Use existing media only'
$optionKeepCentralHistoryCheck = New-OptionCheck -Text 'Keep central log history'
$optionNoCentralCollectionCheck = New-OptionCheck -Text 'No central log collection'
$optionDryRunCheck = New-OptionCheck -Text 'Dry run'

$optionSetupSourceBox = New-EntryBox
$optionSetupSourceBox.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)
$optionSetupSourceBox.Text = ''
$optionSetupSourceBox.Add_Enter({
    if (Test-IsSetupSourcePlaceholder -Value $optionSetupSourceBox.Text) {
        $optionSetupSourceBox.Text = ''
        $optionSetupSourceBox.ForeColor = $colorInk
    }
})
$optionSetupSourceBox.Add_Leave({
    Set-SetupSourcePlaceholder
})
Set-SetupSourcePlaceholder

$optionSetupSourceMapBox = New-EntryBox
$optionSetupSourceMapBox.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)

$optionSetupModeCombo = New-Object System.Windows.Forms.ComboBox
$optionSetupModeCombo.Dock = 'Fill'
$optionSetupModeCombo.DropDownStyle = 'DropDownList'
$optionSetupModeCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)
$optionSetupModeCombo.FlatStyle = 'Flat'
$optionSetupModeCombo.Items.AddRange([object[]]@('LocalCache', 'Share', 'Auto'))
$optionSetupModeCombo.SelectedIndex = 0

$optionSetupMediaIdBox = New-EntryBox
$optionSetupMediaIdBox.Text = 'Win11'
$optionSetupMediaIdBox.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)

$optionSetupLanguageCombo = New-Object System.Windows.Forms.ComboBox
$optionSetupLanguageCombo.Dock = 'Fill'
$optionSetupLanguageCombo.DropDownStyle = 'DropDown'
$optionSetupLanguageCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)
$optionSetupLanguageCombo.FlatStyle = 'Flat'
$optionSetupLanguageCombo.Items.AddRange([object[]]@('MatchSystem','Any','fr-FR','en-GB','en-US','de-DE','es-ES','it-IT','nl-NL','pt-PT','pl-PL'))
$optionSetupLanguageCombo.Text = 'MatchSystem'

$optionSetupDynamicUpdateCombo = New-Object System.Windows.Forms.ComboBox
$optionSetupDynamicUpdateCombo.Dock = 'Fill'
$optionSetupDynamicUpdateCombo.DropDownStyle = 'DropDownList'
$optionSetupDynamicUpdateCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)
$optionSetupDynamicUpdateCombo.FlatStyle = 'Flat'
$optionSetupDynamicUpdateCombo.Items.AddRange([object[]]@('Disable','Enable','NoDrivers','NoLCU','NoDriversNoLCU'))
$optionSetupDynamicUpdateCombo.SelectedIndex = 0

$optionThrottleBox = New-OptionNumber -Minimum 1 -Maximum 200 -Value 10
$optionDelayBetweenComputersBox = New-OptionNumber -Minimum 0 -Maximum 3600 -Value 0
$optionGlobalConcurrencyLimitBox = New-OptionNumber -Minimum 1 -Maximum 200 -Value 15
$optionGlobalConcurrencyLeaseTimeoutBox = New-OptionNumber -Minimum 0 -Maximum 1440 -Value 0
$optionDelayBetweenCyclesBox = New-OptionNumber -Minimum 0 -Maximum 1440 -Value 5
$optionMaxCyclesBox = New-OptionNumber -Minimum 0 -Maximum 1000 -Value 0
$optionPsExecTimeoutBox = New-OptionNumber -Minimum 0 -Maximum 1440 -Value 180
$optionSetupSourceCandidateLimitBox = New-OptionNumber -Minimum 0 -Maximum 100 -Value 5
$optionSetupCopyIpGapBox = New-OptionNumber -Minimum 0 -Maximum 10000 -Value 0
$optionSetupCopyJitterBox = New-OptionNumber -Minimum 0 -Maximum 86400 -Value 0

$optionAuditOnlyCheck.Checked = Get-EnvSwitch -Name 'W11UT_AUDIT_ONLY'
$optionAllowPolicyRepairCheck.Checked = Get-EnvSwitch -Name 'W11UT_ALLOW_POLICY_REPAIR' -Default $true
$optionAllowWUResetCheck.Checked = Get-EnvSwitch -Name 'W11UT_ALLOW_WU_RESET' -Default $true
$optionAllowForceUpgradeCheck.Checked = Get-EnvSwitch -Name 'W11UT_ALLOW_FORCE_UPGRADE' -Default $true
$optionAllowSetupUpgradeCheck.Checked = Get-EnvSwitch -Name 'W11UT_ALLOW_SETUP_UPGRADE' -Default $true
$optionDirectSetupUpgradeCheck.Checked = Get-EnvSwitch -Name 'W11UT_DIRECT_SETUP_UPGRADE'
$optionAllowRebootCheck.Checked = Get-EnvSwitch -Name 'W11UT_ALLOW_REBOOT' -Default $true
$optionSkipVirtualMachinesCheck.Checked = Get-EnvSwitch -Name 'W11UT_SKIP_VIRTUAL_MACHINES' -Default $true
$optionAllowDiskCleanupCheck.Checked = Get-EnvSwitch -Name 'W11UT_ALLOW_DISK_CLEANUP' -Default $true
$optionAllowAdvancedCleanupCheck.Checked = ((Get-EnvSwitch -Name 'W11UT_ALLOW_ADVANCED_DISK_CLEANUP') -or (Get-EnvSwitch -Name 'W11UT_ALLOW_DISM_COMPONENT_CLEANUP'))
$optionSkipSetupPreCopyCheck.Checked = Get-EnvSwitch -Name 'W11UT_SKIP_SETUP_MEDIA_PRECOPY'

$setupSourceDefault = Get-ToolkitConfigValue -Name 'W11UT_SETUP_SOURCE'
if (-not [string]::IsNullOrWhiteSpace($setupSourceDefault)) {
    $optionSetupSourceBox.Text = $setupSourceDefault
    $optionSetupSourceBox.ForeColor = $colorInk
}

$setupSourceMapDefault = Get-ToolkitConfigValue -Name 'W11UT_SETUP_SOURCE_MAP'
if (-not [string]::IsNullOrWhiteSpace($setupSourceMapDefault)) {
    $optionSetupSourceMapBox.Text = $setupSourceMapDefault
}

$setupModeDefault = Get-ToolkitConfigValue -Name 'W11UT_SETUP_EXECUTION_MODE'
if ($setupModeDefault -in @('LocalCache','Share','Auto')) {
    $optionSetupModeCombo.SelectedItem = $setupModeDefault
}

$setupMediaDefault = Get-ToolkitConfigValue -Name 'W11UT_SETUP_MEDIA_ID'
if (-not [string]::IsNullOrWhiteSpace($setupMediaDefault)) {
    $optionSetupMediaIdBox.Text = $setupMediaDefault
}

$setupLanguageDefault = Get-ToolkitConfigValue -Name 'W11UT_SETUP_LANGUAGE'
if (-not [string]::IsNullOrWhiteSpace($setupLanguageDefault)) {
    $optionSetupLanguageCombo.Text = $setupLanguageDefault
}

$setupDynamicUpdateDefault = Get-ToolkitConfigValue -Name 'W11UT_SETUP_DYNAMIC_UPDATE'
if ($setupDynamicUpdateDefault -in @('Enable','Disable','NoDrivers','NoLCU','NoDriversNoLCU')) {
    $optionSetupDynamicUpdateCombo.SelectedItem = $setupDynamicUpdateDefault
}

Set-OptionNumberFromEnv -Control $optionDelayBetweenCyclesBox -Name 'W11UT_DELAY_BETWEEN_CYCLES_MINUTES'
Set-OptionNumberFromEnv -Control $optionDelayBetweenComputersBox -Name 'W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS'
Set-OptionNumberFromEnv -Control $optionPsExecTimeoutBox -Name 'W11UT_PSEXEC_TIMEOUT_MINUTES'
Set-OptionNumberFromEnv -Control $optionThrottleBox -Name 'W11UT_THROTTLE'
Set-OptionNumberFromEnv -Control $optionGlobalConcurrencyLimitBox -Name 'W11UT_GLOBAL_CONCURRENCY_LIMIT'
Set-OptionNumberFromEnv -Control $optionGlobalConcurrencyLeaseTimeoutBox -Name 'W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES'
Set-OptionNumberFromEnv -Control $optionSetupSourceCandidateLimitBox -Name 'W11UT_SETUP_SOURCE_CANDIDATE_LIMIT'
Set-OptionNumberFromEnv -Control $optionSetupCopyIpGapBox -Name 'W11UT_SETUP_COPY_IPG_MS'
Set-OptionNumberFromEnv -Control $optionSetupCopyJitterBox -Name 'W11UT_SETUP_COPY_JITTER_SECONDS'

function Set-OptionCheckAvailability {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.CheckBox]$CheckBox,
        [Parameter(Mandatory = $true)][bool]$Enabled
    )

    $CheckBox.Enabled = $Enabled
    $CheckBox.ForeColor = if ($Enabled) { $colorInk } else { $colorMuted }
}

function Set-LauncherOptionState {
    $auditOnly = [bool]$optionAuditOnlyCheck.Checked
    $directSetup = [bool]$optionDirectSetupUpgradeCheck.Checked
    $auditControlledChecks = @(
        $optionAllowPolicyRepairCheck,
        $optionAllowWUResetCheck,
        $optionAllowForceUpgradeCheck,
        $optionAllowSetupUpgradeCheck,
        $optionDirectSetupUpgradeCheck,
        $optionAllowRebootCheck,
        $optionSkipVirtualMachinesCheck,
        $optionAllowDiskCleanupCheck,
        $optionAllowAdvancedCleanupCheck,
        $optionSkipSetupPreCopyCheck,
        $optionKeepCentralHistoryCheck,
        $optionNoCentralCollectionCheck,
        $optionDryRunCheck
    )
    $directSetupIgnoredChecks = @(
        $optionAllowPolicyRepairCheck,
        $optionAllowWUResetCheck,
        $optionAllowForceUpgradeCheck,
        $optionAllowSetupUpgradeCheck,
        $optionAllowRebootCheck,
        $optionAllowDiskCleanupCheck,
        $optionAllowAdvancedCleanupCheck
    )

    foreach ($check in $auditControlledChecks) {
        if ($auditOnly) {
            $check.Checked = $false
        }
        Set-OptionCheckAvailability -CheckBox $check -Enabled (-not $auditOnly)
    }

    if ($auditOnly) { return }

    foreach ($check in $directSetupIgnoredChecks) {
        if ($directSetup) {
            $check.Checked = $false
        }
        Set-OptionCheckAvailability -CheckBox $check -Enabled (-not $directSetup)
    }
}

$optionAuditOnlyCheck.Add_CheckedChanged({ Set-LauncherOptionState })
$optionDirectSetupUpgradeCheck.Add_CheckedChanged({ Set-LauncherOptionState })
Set-LauncherOptionState

$optionsTable.Controls.Add($optionAuditOnlyCheck, 0, 0)
$optionsTable.Controls.Add($optionAllowPolicyRepairCheck, 1, 0)
$optionsTable.Controls.Add($optionAllowWUResetCheck, 2, 0)
$optionsTable.Controls.Add($optionAllowForceUpgradeCheck, 3, 0)

$optionsTable.Controls.Add($optionAllowSetupUpgradeCheck, 0, 1)
$optionsTable.Controls.Add($optionAllowRebootCheck, 1, 1)
$optionsTable.Controls.Add($optionDirectSetupUpgradeCheck, 2, 1)
$optionsTable.Controls.Add($optionSkipVirtualMachinesCheck, 3, 1)

$optionsTable.Controls.Add((New-Label 'Setup source'), 0, 2)
$optionsTable.Controls.Add($optionSetupSourceBox, 1, 2)
$optionsTable.SetColumnSpan($optionSetupSourceBox, 3)

$optionsTable.Controls.Add((New-Label 'Setup source map'), 0, 3)
$optionsTable.Controls.Add($optionSetupSourceMapBox, 1, 3)
$optionsTable.SetColumnSpan($optionSetupSourceMapBox, 3)

$optionsTable.Controls.Add((New-Label 'Setup mode'), 0, 4)
$optionsTable.Controls.Add($optionSetupModeCombo, 1, 4)
$optionsTable.Controls.Add((New-Label 'Setup media id'), 2, 4)
$optionsTable.Controls.Add($optionSetupMediaIdBox, 3, 4)

$optionsTable.Controls.Add((New-Label 'Setup language'), 0, 5)
$optionsTable.Controls.Add($optionSetupLanguageCombo, 1, 5)
$optionsTable.Controls.Add((New-Label 'Dynamic update'), 2, 5)
$optionsTable.Controls.Add($optionSetupDynamicUpdateCombo, 3, 5)

$optionsTable.Controls.Add($optionSkipSetupPreCopyCheck, 0, 6)
$optionsTable.Controls.Add($optionAllowDiskCleanupCheck, 1, 6)
$optionsTable.Controls.Add($optionAllowAdvancedCleanupCheck, 2, 6)
$optionsTable.Controls.Add($optionKeepCentralHistoryCheck, 3, 6)

$optionsTable.Controls.Add($optionDryRunCheck, 0, 7)
$optionsTable.Controls.Add((New-Label 'Throttle per LOT'), 1, 7)
$optionsTable.Controls.Add($optionThrottleBox, 2, 7)

$optionsTable.Controls.Add((New-Label 'Computer delay sec'), 0, 8)
$optionsTable.Controls.Add($optionDelayBetweenComputersBox, 1, 8)
$optionsTable.Controls.Add((New-Label 'Delay between cycles'), 2, 8)
$optionsTable.Controls.Add($optionDelayBetweenCyclesBox, 3, 8)

$optionsTable.Controls.Add((New-Label 'Max cycles'), 0, 9)
$optionsTable.Controls.Add($optionMaxCyclesBox, 1, 9)
$optionsTable.Controls.Add((New-Label 'PsExec timeout min'), 2, 9)
$optionsTable.Controls.Add($optionPsExecTimeoutBox, 3, 9)

$optionsTable.Controls.Add((New-Label 'Global worker limit'), 0, 10)
$optionsTable.Controls.Add($optionGlobalConcurrencyLimitBox, 1, 10)
$optionsTable.Controls.Add((New-Label 'Global lease timeout min'), 2, 10)
$optionsTable.Controls.Add($optionGlobalConcurrencyLeaseTimeoutBox, 3, 10)

$optionsTable.Controls.Add((New-Label 'Copy IPG ms'), 0, 11)
$optionsTable.Controls.Add($optionSetupCopyIpGapBox, 1, 11)
$optionsTable.Controls.Add((New-Label 'Copy jitter sec'), 2, 11)
$optionsTable.Controls.Add($optionSetupCopyJitterBox, 3, 11)

$optionsTable.Controls.Add((New-Label 'Candidate limit'), 0, 12)
$optionsTable.Controls.Add($optionSetupSourceCandidateLimitBox, 1, 12)
$optionsTable.Controls.Add($optionNoCentralCollectionCheck, 2, 12)

$optionsNote = New-Object System.Windows.Forms.Label
$optionsNote.Text = 'For LOT/PsExec, Setup source must be a UNC path reachable by target computers. Use semicolons to list site shares; each target selects the nearest valid source.'
$optionsNote.Dock = 'Fill'
$optionsNote.ForeColor = $colorMuted
$optionsNote.TextAlign = 'MiddleLeft'
$optionsTable.Controls.Add($optionsNote, 0, 13)
$optionsTable.SetColumnSpan($optionsNote, 4)

$optionsScrollPanel.Controls.Add($optionsTable)
$optionsSection.Content.Controls.Add($optionsScrollPanel)
$optionsLayout.Controls.Add($optionsSection.Panel, 0, 0)
$optionsTab.Controls.Add($optionsLayout)

$globalConcurrencyLimitBox.Value = [int]$optionGlobalConcurrencyLimitBox.Value
$script:SyncingGlobalConcurrencyLimit = $false
$globalConcurrencyLimitBox.Add_ValueChanged({
    if ($script:SyncingGlobalConcurrencyLimit) { return }
    $script:SyncingGlobalConcurrencyLimit = $true
    try { $optionGlobalConcurrencyLimitBox.Value = [int]$globalConcurrencyLimitBox.Value }
    finally { $script:SyncingGlobalConcurrencyLimit = $false }
})
$optionGlobalConcurrencyLimitBox.Add_ValueChanged({
    if ($script:SyncingGlobalConcurrencyLimit) { return }
    $script:SyncingGlobalConcurrencyLimit = $true
    try { $globalConcurrencyLimitBox.Value = [int]$optionGlobalConcurrencyLimitBox.Value }
    finally { $script:SyncingGlobalConcurrencyLimit = $false }
})

$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = 'Fill'
$footerPanel.BackColor = $colorPanel
$footerPanel.Padding = New-Object System.Windows.Forms.Padding(12)
Add-SoftBorder -Control $footerPanel

$footerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$footerLayout.Dock = 'Fill'
$footerLayout.ColumnCount = 2
$footerLayout.RowCount = 1
$footerLayout.BackColor = $colorPanel
$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120))) | Out-Null

$footerStatus = New-Object System.Windows.Forms.Label
$footerStatus.Dock = 'Fill'
$footerStatus.Text = 'Ready'
$footerStatus.TextAlign = 'MiddleLeft'
$footerStatus.ForeColor = $colorMuted
$footerStatus.Font = $smallFont

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = 'Close'
$closeButton.Width = 94
Set-FlatButtonStyle -Button $closeButton -BackColor $colorPanel -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = 'Fill'
$buttonPanel.FlowDirection = 'RightToLeft'
$buttonPanel.BackColor = $colorPanel
$buttonPanel.Padding = New-Object System.Windows.Forms.Padding(0)
$buttonPanel.Controls.Add($closeButton)

$footerLayout.Controls.Add($footerStatus, 0, 0)
$footerLayout.Controls.Add($buttonPanel, 1, 0)
$footerPanel.Controls.Add($footerLayout)

$rootLayout.Controls.Add($headerPanel, 0, 0)
$rootLayout.Controls.Add($actionPanel, 0, 1)
$rootLayout.Controls.Add($tabShell, 0, 2)
$rootLayout.Controls.Add($footerPanel, 0, 3)
$form.Controls.Add($rootLayout)

$script:LotList = @()
$script:SelectedLotSummary = $null
$script:CreatedComputersPath = $null
$script:LastSingleRunPath = $null

function Add-Status {
    param([string]$Message)

    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    $footerStatus.Text = $Message
    $actionStatusLabel.Text = $Message
    if ([string]::IsNullOrWhiteSpace($statusBox.Text)) {
        $statusBox.Text = $line
    }
    else {
        $statusBox.AppendText([Environment]::NewLine + $line)
    }
}

function Get-ToolkitOptionArguments {
    $arguments = New-Object System.Collections.ArrayList

    [void]$arguments.Add('-ThrottleLimit')
    [void]$arguments.Add([string][int]$optionThrottleBox.Value)

    if ([int]$optionMaxCyclesBox.Value -gt 0) {
        [void]$arguments.Add('-MaxCycles')
        [void]$arguments.Add([string][int]$optionMaxCyclesBox.Value)
    }

    if ($optionDryRunCheck.Checked) { [void]$arguments.Add('-DryRun') }
    if ($optionKeepCentralHistoryCheck.Checked) { [void]$arguments.Add('-KeepCentralLogHistory') }
    if ($optionNoCentralCollectionCheck.Checked) { [void]$arguments.Add('-NoCentralLogCollection') }

    return @($arguments.ToArray())
}

function Get-ToolkitOptionEnvironment {
    $environment = @{
        W11UT_GLOBAL_CONCURRENCY_LIMIT = [string][int]$optionGlobalConcurrencyLimitBox.Value
        W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = [string][int]$optionGlobalConcurrencyLeaseTimeoutBox.Value
        W11UT_THROTTLE = [string][int]$optionThrottleBox.Value
        W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS = [string][int]$optionDelayBetweenComputersBox.Value
        W11UT_DELAY_BETWEEN_CYCLES_MINUTES = [string][int]$optionDelayBetweenCyclesBox.Value
        W11UT_PSEXEC_TIMEOUT_MINUTES = [string][int]$optionPsExecTimeoutBox.Value
        W11UT_SETUP_EXECUTION_MODE = [string]$optionSetupModeCombo.SelectedItem
        W11UT_SETUP_LANGUAGE = $optionSetupLanguageCombo.Text.Trim()
        W11UT_SETUP_DYNAMIC_UPDATE = [string]$optionSetupDynamicUpdateCombo.SelectedItem
        W11UT_SETUP_SOURCE_CANDIDATE_LIMIT = [string][int]$optionSetupSourceCandidateLimitBox.Value
        W11UT_SETUP_COPY_IPG_MS = [string][int]$optionSetupCopyIpGapBox.Value
        W11UT_SETUP_COPY_JITTER_SECONDS = [string][int]$optionSetupCopyJitterBox.Value
    }

    if (-not [string]::IsNullOrWhiteSpace($optionSetupMediaIdBox.Text)) {
        $environment["W11UT_SETUP_MEDIA_ID"] = $optionSetupMediaIdBox.Text.Trim()
    }

    $setupSource = Get-SetupSourceText
    if (-not [string]::IsNullOrWhiteSpace($setupSource)) {
        $environment["W11UT_SETUP_SOURCE"] = $setupSource
    }
    $setupSourceMap = $optionSetupSourceMapBox.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($setupSourceMap)) {
        $environment["W11UT_SETUP_SOURCE_MAP"] = $setupSourceMap
    }

    $environment["W11UT_AUDIT_ONLY"] = if ($optionAuditOnlyCheck.Checked) { "1" } else { "0" }
    $environment["W11UT_ALLOW_POLICY_REPAIR"] = if ($optionAllowPolicyRepairCheck.Checked) { "1" } else { "0" }
    $environment["W11UT_ALLOW_WU_RESET"] = if ($optionAllowWUResetCheck.Checked) { "1" } else { "0" }
$environment["W11UT_ALLOW_FORCE_UPGRADE"] = if ($optionAllowForceUpgradeCheck.Checked) { "1" } else { "0" }
$environment["W11UT_ALLOW_SETUP_UPGRADE"] = if ($optionAllowSetupUpgradeCheck.Checked) { "1" } else { "0" }
$environment["W11UT_DIRECT_SETUP_UPGRADE"] = if ($optionDirectSetupUpgradeCheck.Checked) { "1" } else { "0" }
$environment["W11UT_ALLOW_REBOOT"] = if ($optionAllowRebootCheck.Checked) { "1" } else { "0" }
    $environment["W11UT_SKIP_VIRTUAL_MACHINES"] = if ($optionSkipVirtualMachinesCheck.Checked) { "1" } else { "0" }
    $environment["W11UT_ALLOW_DISK_CLEANUP"] = if ($optionAllowDiskCleanupCheck.Checked) { "1" } else { "0" }
    $environment["W11UT_ALLOW_ADVANCED_DISK_CLEANUP"] = if ($optionAllowAdvancedCleanupCheck.Checked) { "1" } else { "0" }
    $environment["W11UT_SKIP_SETUP_MEDIA_PRECOPY"] = if ($optionSkipSetupPreCopyCheck.Checked) { "1" } else { "0" }

    return $environment
}

function Test-SetupSourceBeforeLaunch {
    param(
        [hashtable]$EnvironmentVariables = $null,
        [string]$ScopeName = 'current options'
    )

    if ($EnvironmentVariables) {
        $setupSource = [string]$EnvironmentVariables['W11UT_SETUP_SOURCE']
        $setupSourceMap = [string]$EnvironmentVariables['W11UT_SETUP_SOURCE_MAP']
        $mode = [string]$EnvironmentVariables['W11UT_SETUP_EXECUTION_MODE']
        $allowSetupUpgrade = ([string]$EnvironmentVariables['W11UT_ALLOW_SETUP_UPGRADE'] -eq '1' -or [string]$EnvironmentVariables['W11UT_DIRECT_SETUP_UPGRADE'] -eq '1')
        $skipSetupPreCopy = ([string]$EnvironmentVariables['W11UT_SKIP_SETUP_MEDIA_PRECOPY'] -eq '1')
    }
    else {
        $setupSource = Get-SetupSourceText
        $setupSourceMap = $optionSetupSourceMapBox.Text.Trim()
        $mode = [string]$optionSetupModeCombo.SelectedItem
        $allowSetupUpgrade = ($optionAllowSetupUpgradeCheck.Checked -or $optionDirectSetupUpgradeCheck.Checked)
        $skipSetupPreCopy = $optionSkipSetupPreCopyCheck.Checked
    }

    if (-not $allowSetupUpgrade) { return $true }
    if ($skipSetupPreCopy -and $mode -ne 'Share') { return $true }

    if ([string]::IsNullOrWhiteSpace($setupSource) -and [string]::IsNullOrWhiteSpace($setupSourceMap)) {
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            ("Setup source or setup source map is required for setup upgrade when target media copy is enabled.`r`n`r`nScope: {0}`r`n`r`nUse a UNC path reachable by the target computers, for example:`r`n{1}`r`n`r`nOr set W11UT_SETUP_SOURCE_MAP to a CSV map path, or enable Use existing media only if the target cache is already valid." -f $ScopeName,$script:SetupSourcePlaceholder),
            'Setup source required',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        if (-not $EnvironmentVariables) {
            $optionSetupSourceBox.Focus()
        }
        return $false
    }

    $setupSources = @(
        Split-SetupSourceText -Value $setupSource
        if (-not [string]::IsNullOrWhiteSpace($setupSourceMap)) { $setupSourceMap }
    )
    $localSources = @($setupSources | Where-Object { -not (Test-IsUncPath -Path $_) })
    if ($setupSources.Count -gt 0 -and $localSources.Count -eq 0) { return $true }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        $form,
        ("One or more setup sources are not UNC paths:`r`n{0}`r`n`r`nIn LOT/PsExec mode, the target computer runs as SYSTEM and must read these paths itself. A local or relative path usually exists only on the technician workstation.`r`n`r`nContinue anyway for a local/direct test or an explicitly shared identical path?" -f ($localSources -join "`r`n")),
        'Confirm local setup source',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    return ($answer -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Clear-LotDetails {
    $script:SelectedLotSummary = $null
    $deviceCountBox.Text = ''
    $upgradeScopeBox.Text = ''
    $actionTitleLabel.Text = 'LOT: none selected'
    $actionSubtitleLabel.Text = 'Select or create a LOT.'
    $actionStatusLabel.Text = 'Ready'
    Set-ButtonEnabledStyle -Button $launchExistingButton -Enabled $false -EnabledBackColor $colorAccent
}

function Get-LaunchableLotSummaries {
    $summaries = New-Object System.Collections.ArrayList
    foreach ($lot in @($script:LotList)) {
        try {
            $summary = Get-LotSummary -LotPath $lot.FullName
            if ($summary.ComputerCount -gt 0 -and $summary.WrappersReady) {
                [void]$summaries.Add($summary)
            }
        }
        catch {
            Add-Status ("Skipping {0}: {1}" -f $lot.Name, $_.Exception.Message)
        }
    }

    return @($summaries.ToArray())
}

function Update-ExistingLotControlState {
    param([bool]$HasLots)

    $lotCombo.Enabled = $HasLots
    $existingModeCombo.Enabled = $HasLots
    $globalConcurrencyLimitBox.Enabled = $HasLots
    $openLotFolderButton.Enabled = $HasLots
    $openComputersButton.Enabled = $HasLots
    $openReportsButton.Enabled = $HasLots
    $hasLaunchableLots = ($HasLots -and (Get-LaunchableLotSummaries).Count -gt 0)
    Set-ButtonEnabledStyle -Button $launchAllLotsButton -Enabled $hasLaunchableLots -EnabledBackColor $colorAccent
    Set-ButtonEnabledStyle -Button $actionLaunchAllButton -Enabled $hasLaunchableLots -EnabledBackColor $colorAccent
}

function Update-LotDetails {
    if ($lotCombo.SelectedIndex -lt 0 -or $lotCombo.SelectedIndex -ge $script:LotList.Count) {
        Clear-LotDetails
        return
    }

    $lot = $script:LotList[$lotCombo.SelectedIndex]
    $summary = Get-LotSummary -LotPath $lot.FullName
    $script:SelectedLotSummary = $summary

    $deviceCountBox.Text = [string]$summary.ComputerCount
    $upgradeScopeBox.Text = $summary.UpgradeScope
    $actionTitleLabel.Text = 'LOT: {0}' -f $summary.Name
    $actionSubtitleLabel.Text = '{0} device(s) - Windows 11 upgrade readiness' -f $summary.ComputerCount

    if (-not $summary.WrappersReady) {
        Add-Status ("Missing LOT wrappers: {0}" -f ($summary.MissingWrappers -join ', '))
    }

    $canLaunch = ($summary.ComputerCount -gt 0 -and $summary.WrappersReady)
    Set-ButtonEnabledStyle -Button $launchExistingButton -Enabled $canLaunch -EnabledBackColor $colorAccent
    Update-ExistingLotControlState -HasLots ($script:LotList.Count -gt 0)
}

function Refresh-LotList {
    param([string]$PreferredName)

    $selectedName = if (-not [string]::IsNullOrWhiteSpace($PreferredName)) {
        $PreferredName
    }
    elseif ($lotCombo.SelectedItem) {
        [string]$lotCombo.SelectedItem
    }
    else {
        ''
    }
    $script:LotList = @(Get-LotFolders -RootPath $toolkitRoot)

    $lotCombo.Items.Clear()
    foreach ($lot in $script:LotList) {
        $lotCombo.Items.Add($lot.Name) | Out-Null
    }

    if ($lotCombo.Items.Count -eq 0) {
        Clear-LotDetails
        Update-ExistingLotControlState -HasLots $false
        Add-Status 'No operational LOT-* folder found.'
        return
    }

    Update-ExistingLotControlState -HasLots $true

    $selectedIndex = 0
    if (-not [string]::IsNullOrWhiteSpace($selectedName)) {
        for ($i = 0; $i -lt $lotCombo.Items.Count; $i++) {
            if ([string]$lotCombo.Items[$i] -eq $selectedName) {
                $selectedIndex = $i
                break
            }
        }
    }

    $lotCombo.SelectedIndex = $selectedIndex
    Update-LotDetails
    Add-Status ("Loaded {0} LOT folder(s)." -f $lotCombo.Items.Count)
}

$refreshLotsButton.Add_Click({
    try {
        Refresh-LotList
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Refresh failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$actionRefreshButton.Add_Click({
    $refreshLotsButton.PerformClick()
})

$lotCombo.Add_SelectedIndexChanged({
    try {
        Update-LotDetails
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$actionLaunchAllButton.Add_Click({
    $launchAllLotsButton.PerformClick()
})

$openLotFolderButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw 'Select a LOT first.' }
        Open-FolderPath -Path $script:SelectedLotSummary.Path
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Open folder failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$openComputersButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw 'Select a LOT first.' }
        Open-TextFile -Path $script:SelectedLotSummary.ComputersPath
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$openReportsButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw 'Select a LOT first.' }
        Open-OrCreateFolderPath -Path $script:SelectedLotSummary.ReportsPath
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Open reports failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$refreshWrappersButton.Add_Click({
    try {
        Invoke-LotWrapperRefresh -RootPath $toolkitRoot
        Add-Status 'LOT wrappers refreshed.'
        Refresh-LotList
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Wrapper refresh failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchExistingButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw 'Select a LOT first.' }
        if ($script:SelectedLotSummary.ComputerCount -le 0) { throw 'Computers.txt is empty.' }

        $optionArguments = @(Get-ToolkitOptionArguments)
        $optionEnvironment = Get-ToolkitOptionEnvironment
        $effectiveEnvironment = Get-EffectiveLotEnvironment -LotPath $script:SelectedLotSummary.Path -EnvironmentVariables $optionEnvironment
        if (-not (Test-SetupSourceBeforeLaunch -EnvironmentVariables $effectiveEnvironment -ScopeName $script:SelectedLotSummary.Name)) { return }

        Add-Status ("Launching {0} in {1} mode. Global limit={2}. Args={3}; env={4}." -f $script:SelectedLotSummary.Name, $existingModeCombo.SelectedItem, [int]$globalConcurrencyLimitBox.Value, $optionArguments.Count, $effectiveEnvironment.Count)
        Start-ToolkitLot -LotPath $script:SelectedLotSummary.Path -Mode ([string]$existingModeCombo.SelectedItem) -GlobalConcurrencyLimit ([int]$globalConcurrencyLimitBox.Value) -GlobalConcurrencyLeaseTimeoutMinutes ([int]$optionGlobalConcurrencyLeaseTimeoutBox.Value) -AdditionalArguments $optionArguments -EnvironmentVariables $optionEnvironment
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'LOT launch failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchAllLotsButton.Add_Click({
    try {
        $launchableLots = @(Get-LaunchableLotSummaries)
        if ($launchableLots.Count -eq 0) {
            throw 'No launchable LOT found. Check Computers.txt and wrapper files.'
        }
        $mode = [string]$existingModeCombo.SelectedItem
        $limit = [int]$globalConcurrencyLimitBox.Value
        $optionArguments = @(Get-ToolkitOptionArguments)
        $optionEnvironment = Get-ToolkitOptionEnvironment

        foreach ($lotSummary in $launchableLots) {
            $effectiveEnvironment = Get-EffectiveLotEnvironment -LotPath $lotSummary.Path -EnvironmentVariables $optionEnvironment
            if (-not (Test-SetupSourceBeforeLaunch -EnvironmentVariables $effectiveEnvironment -ScopeName $lotSummary.Name)) { return }
        }

        Add-Status ("Launching {0}/{1} LOT folder(s) in {2} mode. Global limit={3}. Delay={4}s. Args={5}; env={6}." -f $launchableLots.Count, $script:LotList.Count, $mode, $limit, $launchAllLotStartDelaySeconds, $optionArguments.Count, $optionEnvironment.Count)

        for ($lotIndex = 0; $lotIndex -lt $launchableLots.Count; $lotIndex++) {
            $lotSummary = $launchableLots[$lotIndex]
            if ($lotIndex -gt 0 -and $launchAllLotStartDelaySeconds -gt 0) {
                Add-Status ("Waiting {0}s before launching {1}." -f $launchAllLotStartDelaySeconds, $lotSummary.Name)
                Wait-UiDelay -Seconds $launchAllLotStartDelaySeconds
            }

            Add-Status ("Launching {0}." -f $lotSummary.Name)
            Start-ToolkitLot -LotPath $lotSummary.Path -Mode $mode -GlobalConcurrencyLimit $limit -GlobalConcurrencyLeaseTimeoutMinutes ([int]$optionGlobalConcurrencyLeaseTimeoutBox.Value) -AdditionalArguments $optionArguments -EnvironmentVariables $optionEnvironment
        }

        $skippedCount = $script:LotList.Count - $launchableLots.Count
        if ($skippedCount -gt 0) {
            Add-Status ("Skipped {0} LOT folder(s) without devices or complete wrappers." -f $skippedCount)
        }
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Launch all failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchSingleComputerButton.Add_Click({
    try {
        if (-not (Test-SetupSourceBeforeLaunch)) { return }

        $computerName = $singleComputerBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($computerName)) { throw 'Enter a computer name.' }

        $optionArguments = @(Get-ToolkitOptionArguments)
        $optionEnvironment = Get-ToolkitOptionEnvironment
        Add-Status ("Launching single computer {0} in {1} mode. Global limit={2}. Args={3}; env={4}." -f $computerName, $singleModeCombo.SelectedItem, [int]$globalConcurrencyLimitBox.Value, $optionArguments.Count, $optionEnvironment.Count)
        $run = Start-ToolkitSingleComputer -ComputerName $computerName -Mode ([string]$singleModeCombo.SelectedItem) -GlobalConcurrencyLimit ([int]$globalConcurrencyLimitBox.Value) -GlobalConcurrencyLeaseTimeoutMinutes ([int]$optionGlobalConcurrencyLeaseTimeoutBox.Value) -AdditionalArguments $optionArguments -EnvironmentVariables $optionEnvironment
        $script:LastSingleRunPath = $run.RunPath
        $singleRunPathBox.Text = $run.RunPath
        Set-FlatButtonStyle -Button $openSingleRunFolderButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder
        Add-Status ("Single computer run folder: {0}" -f $run.RunPath)
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Single computer launch failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$openSingleRunFolderButton.Add_Click({
    try {
        $pathToOpen = $script:LastSingleRunPath
        if ([string]::IsNullOrWhiteSpace($pathToOpen)) {
            $pathToOpen = $script:SingleComputerRunRoot
        }
        if ([string]::IsNullOrWhiteSpace($pathToOpen)) {
            throw 'No single computer run folder configured.'
        }
        Open-OrCreateFolderPath -Path $pathToOpen
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Open run folder failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$createLotButton.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($newLotNameBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show($form, 'Enter a LOT name.', 'LOT name required', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $newLotNameBox.Focus()
            return
        }

        $result = New-ToolkitLotFolder -RootPath $toolkitRoot -LotName $newLotNameBox.Text.Trim()
        $newComputersPathBox.Text = $result.ComputersPath
        $script:CreatedComputersPath = $result.ComputersPath
        Set-ButtonEnabledStyle -Button $openNewComputersButton -Enabled $true

        Add-Status ("Created empty LOT: {0}" -f $result.LotPath)
        Refresh-LotList -PreferredName (Split-Path -Leaf $result.LotPath)
        Show-DeviceRegistrationTabPage -Header $newTabHeader

        $answer = [System.Windows.Forms.MessageBox]::Show(
            $form,
            ("LOT created:`r`n{0}`r`n`r`nOpen Computers.txt now?" -f $result.LotPath),
            'Fill Computers.txt',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            Open-TextFile -Path $result.ComputersPath
        }
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'LOT creation failed', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$openNewComputersButton.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($script:CreatedComputersPath)) {
            throw 'Create a LOT first.'
        }

        Open-TextFile -Path $script:CreatedComputersPath
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$closeButton.Add_Click({
    $form.Close()
})

$form.Add_Shown({
    try {
        Refresh-LotList
        $lotCombo.Focus() | Out-Null
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$form.Add_FormClosed({
    if ($script:LogoImage) {
        $script:LogoImage.Dispose()
        $script:LogoImage = $null
    }

    if ($script:LogoIcon) {
        $script:LogoIcon.Dispose()
        $script:LogoIcon = $null
    }
})

if ($ValidateOnly) {
    $folders = @(Get-LotFolders -RootPath $toolkitRoot)
    Write-Host ("Smart Intune Windows 11 Upgrade Toolkit LOT Launcher GUI validation completed. Lots={0}" -f $folders.Count)
    return
}

Add-Status 'Ready. Select an existing LOT or create a new empty LOT.'
[void]$form.ShowDialog()
