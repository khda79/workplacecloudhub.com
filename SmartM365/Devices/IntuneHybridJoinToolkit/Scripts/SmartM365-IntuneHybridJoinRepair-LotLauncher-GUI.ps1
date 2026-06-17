<#
.SYNOPSIS
Creates and launches Smart Intune Hybrid Join Toolkit LOT folders from a GUI.

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

$ErrorActionPreference = "Stop"

function Get-ToolkitRoot {
    $scriptDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    }
    else {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }

    return (Get-Item -LiteralPath (Split-Path -Parent $scriptDir) -ErrorAction Stop).FullName
}

function Get-SafeLotName {
    param([Parameter(Mandatory = $true)][string]$LotName)

    $safeName = [regex]::Replace($LotName.Trim(), "[^A-Za-z0-9._-]+", "-").Trim("-._")
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        throw "Enter a LOT name."
    }

    if ($safeName -notmatch "^(?i)LOT-") {
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
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith("#")) {
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
        Get-ChildItem -LiteralPath $RootPath -Directory -Filter "LOT-*" -ErrorAction Stop |
            Where-Object { $_.Name -ine "LOT-X" } |
            Sort-Object Name
    )
}

function Test-LotWrapperSet {
    param([Parameter(Mandatory = $true)][string]$LotPath)

    $wrapperNames = @(
        "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd",
        "Run-IntuneHybridJoinRepairWithPsExec-Once.cmd",
        "Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd",
        "Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd"
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

function Get-AdDomainText {
    param([Parameter(Mandatory = $true)][string]$LotPath)

    $adDomainPath = Join-Path $LotPath "AdDomain.txt"
    if (-not (Test-Path -LiteralPath $adDomainPath -PathType Leaf)) {
        return ""
    }

    foreach ($line in @(Get-Content -LiteralPath $adDomainPath -ErrorAction Stop)) {
        $value = ([string]$line).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }

    return ""
}

function Get-FileFreshnessText {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$FreshMinutes = 120
    )

    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $item) {
        return "Missing"
    }

    $age = (Get-Date) - $item.LastWriteTime
    $state = if ($age.TotalMinutes -le $FreshMinutes) { "Recent" } else { "Stale" }
    return ("{0}; {1:N1} min; {2}" -f $state,$age.TotalMinutes,$item.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
}

function Get-LotSummary {
    param(
        [Parameter(Mandatory = $true)][string]$RootPath,
        [Parameter(Mandatory = $true)][string]$LotPath
    )

    $computersPath = Join-Path $LotPath "Computers.txt"
    $adDomainPath = Join-Path $LotPath "AdDomain.txt"
    $lotAdCsvPath = Join-Path $LotPath "DevicesAD.csv"
    $rootAdCsvPath = Join-Path $RootPath "DevicesAD.csv"
    $rootIntuneCsvPath = Join-Path $RootPath "DevicesIntune.csv"
    $rootEntraCsvPath = Join-Path $RootPath "DevicesEntra.csv"
    $computers = @(Get-ComputerNamesFromFile -Path $computersPath)
    $adDomain = Get-AdDomainText -LotPath $LotPath
    $wrappers = Test-LotWrapperSet -LotPath $LotPath

    $adScope = if ([string]::IsNullOrWhiteSpace($adDomain)) {
        "Forest export from root DevicesAD.csv"
    }
    else {
        "Domain export: $adDomain"
    }

    $selectedAdCsv = if ([string]::IsNullOrWhiteSpace($adDomain)) { $rootAdCsvPath } else { $lotAdCsvPath }

    return [pscustomobject]@{
        Name = Split-Path -Leaf $LotPath
        Path = $LotPath
        ComputersPath = $computersPath
        ComputerCount = $computers.Count
        AdDomainPath = $adDomainPath
        AdDomain = $adDomain
        AdScope = $adScope
        SelectedAdCsv = $selectedAdCsv
        RootAdCsvStatus = Get-FileFreshnessText -Path $rootAdCsvPath -FreshMinutes 120
        SelectedAdCsvStatus = Get-FileFreshnessText -Path $selectedAdCsv -FreshMinutes 120
        IntuneCsvStatus = Get-FileFreshnessText -Path $rootIntuneCsvPath -FreshMinutes 120
        EntraCsvStatus = Get-FileFreshnessText -Path $rootEntraCsvPath -FreshMinutes 120
        WrappersReady = $wrappers.Ready
        MissingWrappers = $wrappers.Missing
    }
}

function Invoke-LotWrapperRefresh {
    param([Parameter(Mandatory = $true)][string]$RootPath)

    $updateScript = Join-Path $RootPath "Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1"
    if (-not (Test-Path -LiteralPath $updateScript -PathType Leaf)) {
        throw "LOT wrapper update script not found: $updateScript"
    }

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $updateScript -RootPath $RootPath
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
    New-Item -ItemType File -Path (Join-Path $lotPath "Computers.txt") -Force -ErrorAction Stop | Out-Null
    New-Item -ItemType File -Path (Join-Path $lotPath "AdDomain.txt") -Force -ErrorAction Stop | Out-Null

    Invoke-LotWrapperRefresh -RootPath $rootItem.FullName

    return [pscustomobject]@{
        LotPath = $lotPath
        ComputersPath = Join-Path $lotPath "Computers.txt"
        AdDomainPath = Join-Path $lotPath "AdDomain.txt"
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
        throw "Enter a computer name."
    }

    $safeComputer = [regex]::Replace($trimmed, "[^A-Za-z0-9._-]+", "-").Trim("-._")
    if ([string]::IsNullOrWhiteSpace($safeComputer)) { $safeComputer = "Computer" }

    $runRoot = Get-SingleComputerRunRoot -ToolkitKey $ToolkitKey
    $runPath = Join-Path $runRoot ("{0}_{1}" -f (Get-Date -Format "yyyyMMdd-HHmmss"),$safeComputer)
    New-Item -ItemType Directory -Path $runPath -Force -ErrorAction Stop | Out-Null

    $computersPath = Join-Path $runPath "Computers.txt"
    Set-Content -LiteralPath $computersPath -Value $trimmed -Encoding UTF8 -Force

    foreach ($folder in @("PsExecLogs","Reports","CentralLogs")) {
        New-Item -ItemType Directory -Path (Join-Path $runPath $folder) -Force -ErrorAction Stop | Out-Null
    }

    [pscustomobject]@{
        ComputerName = $trimmed
        RunPath = $runPath
        ComputersPath = $computersPath
        LogRoot = Join-Path $runPath "PsExecLogs"
        ReportRoot = Join-Path $runPath "Reports"
        CentralLogRoot = Join-Path $runPath "CentralLogs"
    }
}

function Get-SingleComputerRunRoot {
    param([Parameter(Mandatory = $true)][string]$ToolkitKey)

    return (Join-Path $toolkitRoot "SingleComputerRuns")
}

function Test-RecentFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$FreshMinutes = 120
    )

    $item = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $item) { return $false }
    return (((Get-Date) - $item.LastWriteTime).TotalMinutes -le $FreshMinutes)
}

function New-SingleComputerInventoryExportCommand {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [string]$Domain
    )

    $parts = @(
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (ConvertTo-CmdArgument -Value $ScriptPath),
        "-OutputPath",
        (ConvertTo-CmdArgument -Value $OutputPath),
        "-ComputerListPath",
        (ConvertTo-CmdArgument -Value $ComputerListPath),
        "-ForceRefresh"
    )
    if (-not [string]::IsNullOrWhiteSpace($Domain)) {
        $parts += "-Domain"
        $parts += (ConvertTo-CmdArgument -Value $Domain)
    }

    return ($parts -join " ")
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
        "Loop" { "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd"; break }
        "Once" { "Run-IntuneHybridJoinRepairWithPsExec-Once.cmd"; break }
        "LoopIgnoreRunGuard" { "Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd"; break }
        "OnceIgnoreRunGuard" { "Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd"; break }
        default { "Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd" }
    }

    $wrapperPath = Join-Path $LotPath $wrapperName
    if (-not (Test-Path -LiteralPath $wrapperPath -PathType Leaf)) {
        throw "LOT wrapper not found: $wrapperPath"
    }

    $psexecToolkitPath = Join-Path $toolkitRoot "Scripts\PsExec.exe"
    $psexecSystem32Path = Join-Path $env:WINDIR "System32\PsExec.exe"
    $psexecCommand = Get-Command -Name "PsExec.exe" -CommandType Application -ErrorAction SilentlyContinue
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
        "-GlobalConcurrencyLimit",
        [string]$GlobalConcurrencyLimit,
        "-GlobalConcurrencyLeaseTimeoutMinutes",
        [string]$GlobalConcurrencyLeaseTimeoutMinutes
    )
    foreach ($argument in @($AdditionalArguments)) {
        if (-not [string]::IsNullOrWhiteSpace($argument)) {
            $commandParts += (ConvertTo-CmdArgument -Value $argument)
        }
    }

    $setCommands = @()
    foreach ($name in @($EnvironmentVariables.Keys | Sort-Object)) {
        $value = [string]$EnvironmentVariables[$name]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $setCommands += (ConvertTo-CmdSetCommand -Name $name -Value $value)
        }
    }

    $commandLine = (($setCommands + @($commandParts -join " ")) -join " & ")
    Start-Process -FilePath "cmd.exe" -ArgumentList @("/k", $commandLine) -WorkingDirectory $LotPath -Verb RunAs
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

    $scriptPath = Join-Path $toolkitRoot "Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1"
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "Intune Hybrid Join PsExec launcher not found: $scriptPath"
    }

    $isDryRun = @($AdditionalArguments) -contains "-DryRun"
    if (-not $isDryRun) {
        $psexecToolkitPath = Join-Path $toolkitRoot "Scripts\PsExec.exe"
        $psexecSystem32Path = Join-Path $env:WINDIR "System32\PsExec.exe"
        $psexecCommand = Get-Command -Name "PsExec.exe" -CommandType Application -ErrorAction SilentlyContinue
        if (
            -not (Test-Path -LiteralPath $psexecToolkitPath -PathType Leaf) -and
            -not (Test-Path -LiteralPath $psexecSystem32Path -PathType Leaf) -and
            -not $psexecCommand
        ) {
            throw ("PsExec.exe not found. Place it in '{0}', in '{1}', or add PsExec.exe to PATH before launching the computer." -f (Split-Path -Parent $psexecToolkitPath), (Split-Path -Parent $psexecSystem32Path))
        }
    }

    $run = New-SingleComputerRunContext -ComputerName $ComputerName -ToolkitKey "IntuneHybridJoinToolkit"
    if ($GlobalConcurrencyLimit -lt 1) { $GlobalConcurrencyLimit = 1 }

    $commandParts = @(
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (ConvertTo-CmdArgument -Value $scriptPath),
        "-ComputerListPath",
        (ConvertTo-CmdArgument -Value $run.ComputersPath),
        "-LogRoot",
        (ConvertTo-CmdArgument -Value $run.LogRoot),
        "-ReportRoot",
        (ConvertTo-CmdArgument -Value $run.ReportRoot),
        "-CentralLogRoot",
        (ConvertTo-CmdArgument -Value $run.CentralLogRoot),
        "-GlobalConcurrencyLimit",
        [string]$GlobalConcurrencyLimit,
        "-GlobalConcurrencyLeaseTimeoutMinutes",
        [string]$GlobalConcurrencyLeaseTimeoutMinutes
    )

    $inventoryCommands = @()
    $intuneInventoryCsv = Join-Path $toolkitRoot "DevicesIntune.csv"
    $entraInventoryCsv = Join-Path $toolkitRoot "DevicesEntra.csv"
    $adRootInventoryCsv = Join-Path $toolkitRoot "DevicesAD.csv"
    $adDomain = [string]$EnvironmentVariables["EHJIR_AD_DOMAIN"]

    $rootInventoriesAreRecent = (
        (Test-RecentFile -Path $intuneInventoryCsv) -and
        (Test-RecentFile -Path $entraInventoryCsv) -and
        (Test-RecentFile -Path $adRootInventoryCsv)
    )

    if ($rootInventoriesAreRecent) {
        $commandParts += "-IntuneInventoryCsv"
        $commandParts += (ConvertTo-CmdArgument -Value $intuneInventoryCsv)
        $commandParts += "-EntraInventoryCsv"
        $commandParts += (ConvertTo-CmdArgument -Value $entraInventoryCsv)
        $commandParts += "-AdRootInventoryCsv"
        $commandParts += (ConvertTo-CmdArgument -Value $adRootInventoryCsv)
        $commandParts += "-AdInventoryCsv"
        $commandParts += (ConvertTo-CmdArgument -Value $adRootInventoryCsv)
    }
    else {
        $intuneInventoryCsv = Join-Path $run.RunPath "DevicesIntune.csv"
        $entraInventoryCsv = Join-Path $run.RunPath "DevicesEntra.csv"
        $adRootInventoryCsv = Join-Path $run.RunPath "DevicesAD.csv"

        $inventoryCommands += (New-SingleComputerInventoryExportCommand -ScriptPath (Join-Path $toolkitRoot "Scripts\SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1") -OutputPath $intuneInventoryCsv -ComputerListPath $run.ComputersPath)
        $inventoryCommands += (New-SingleComputerInventoryExportCommand -ScriptPath (Join-Path $toolkitRoot "Scripts\SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1") -OutputPath $entraInventoryCsv -ComputerListPath $run.ComputersPath)
        $inventoryCommands += (New-SingleComputerInventoryExportCommand -ScriptPath (Join-Path $toolkitRoot "Scripts\SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1") -OutputPath $adRootInventoryCsv -ComputerListPath $run.ComputersPath -Domain $adDomain)

        $commandParts += "-IntuneInventoryCsv"
        $commandParts += (ConvertTo-CmdArgument -Value $intuneInventoryCsv)
        $commandParts += "-EntraInventoryCsv"
        $commandParts += (ConvertTo-CmdArgument -Value $entraInventoryCsv)
        $commandParts += "-AdInventoryCsv"
        $commandParts += (ConvertTo-CmdArgument -Value $adRootInventoryCsv)
    }

    if (-not [string]::IsNullOrWhiteSpace($adDomain)) {
        $commandParts += "-AdDomain"
        $commandParts += (ConvertTo-CmdArgument -Value $adDomain)
    }

    if ($Mode -in @("Once","OnceIgnoreRunGuard")) { $commandParts += "-RunOnce" }
    if ($Mode -in @("LoopIgnoreRunGuard","OnceIgnoreRunGuard")) { $commandParts += "-IgnoreRunGuard" }

    if ([string]$EnvironmentVariables["EHJIR_ALLOW_DSREG_LEAVE"] -eq "1") {
        $commandParts += "-AllowDsregLeave"
    }
    else {
        $commandParts += "-AllowDsregLeave:`$false"
    }
    foreach ($pair in @(
        @{ Name = "EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT"; Argument = "-AllowRemoveStaleIntuneEnrollment" },
        @{ Name = "EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER"; Argument = "-AllowRebootWhenNoInteractiveUser" },
        @{ Name = "EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE"; Argument = "-AllowRebootAfterDsregLeave" },
        @{ Name = "EHJIR_SKIP_VIRTUAL_MACHINES"; Argument = "-SkipVirtualMachines" }
    )) {
        if ([string]$EnvironmentVariables[$pair.Name] -eq "1") { $commandParts += $pair.Argument }
    }

    foreach ($pair in @(
        @{ Name = "EHJIR_THROTTLE"; Argument = "-ThrottleLimit" },
        @{ Name = "EHJIR_DELAY_BETWEEN_CYCLES_MINUTES"; Argument = "-DelayBetweenCyclesMinutes" },
        @{ Name = "EHJIR_INTUNE_RETRY_SLEEP_MINUTES"; Argument = "-IntuneRetrySleepMinutes" },
        @{ Name = "EHJIR_INTUNE_RETRY_MAX_RETRIES"; Argument = "-IntuneRetryMaxRetries" },
        @{ Name = "EHJIR_STALE_CLEANUP_DELAY_SECONDS"; Argument = "-StaleCleanupDelaySeconds" },
        @{ Name = "EHJIR_REBOOT_DELAY_SECONDS"; Argument = "-RebootDelaySeconds" },
        @{ Name = "EHJIR_PSEXEC_TIMEOUT_MINUTES"; Argument = "-PsExecTimeoutMinutes" }
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

    $commandLine = if ($inventoryCommands.Count -gt 0) {
        (($inventoryCommands + @($commandParts -join " ")) -join " && ")
    }
    else {
        $commandParts -join " "
    }
    Start-Process -FilePath "cmd.exe" -ArgumentList @("/k", $commandLine) -WorkingDirectory $run.RunPath -Verb RunAs
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

    Start-Process -FilePath "notepad.exe" -ArgumentList @("`"$Path`"")
}

function Open-FolderPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "Folder not found: $Path"
    }

    Start-Process -FilePath "explorer.exe" -ArgumentList @("`"$Path`"")
}

$toolkitRoot = Get-ToolkitRoot
$launchAllLotStartDelaySeconds = 5

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$titleFont = New-Object System.Drawing.Font("Segoe UI Semibold", 19)
$sectionFont = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$smallFont = New-Object System.Drawing.Font("Segoe UI", 8)
$statusFont = New-Object System.Drawing.Font("Consolas", 9)

$colorBackground = [System.Drawing.ColorTranslator]::FromHtml("#F5F8FB")
$colorPanel = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF")
$colorPanelSoft = [System.Drawing.ColorTranslator]::FromHtml("#FAFCFE")
$colorHeaderPanel = [System.Drawing.ColorTranslator]::FromHtml("#E6F4FF")
$colorAccent = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")
$colorAccentDark = [System.Drawing.ColorTranslator]::FromHtml("#005A9E")
$colorInk = [System.Drawing.ColorTranslator]::FromHtml("#1F2937")
$colorMuted = [System.Drawing.ColorTranslator]::FromHtml("#5F6B7A")
$colorBorder = [System.Drawing.ColorTranslator]::FromHtml("#DDE7F0")
$colorTextBoxBorder = [System.Drawing.ColorTranslator]::FromHtml("#B9C8D7")
$colorSuccess = [System.Drawing.ColorTranslator]::FromHtml("#027A48")
$colorWarning = [System.Drawing.ColorTranslator]::FromHtml("#B54708")
$colorDisabled = [System.Drawing.ColorTranslator]::FromHtml("#E4EAF1")
$colorDisabledText = [System.Drawing.ColorTranslator]::FromHtml("#7A8A99")

function Resolve-LogoIconPath {
    $candidatePaths = @(
        (Join-Path $PSScriptRoot "SmartM365-logo.ico"),
        (Join-Path $toolkitRoot "SmartM365-logo.ico")
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
        (Join-Path $toolkitRoot "workplacecloudhub-v2.png"),
        (Join-Path $PSScriptRoot "workplacecloudhub-v2.png"),
        (Join-Path $devicesRoot "DeviceRegistrationTool\workplacecloudhub-v2.png")
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

    $Button.FlatStyle = "Flat"
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
    $header.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
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
            [System.Drawing.ColorTranslator]::FromHtml("#F8FBFE")
        }
        else {
            $colorPanel
        }
        $borderColor = if ($state.Selected -or $state.Hover) {
            [System.Drawing.ColorTranslator]::FromHtml("#B9DDF7")
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
    $label.Dock = "Fill"
    $label.TextAlign = "MiddleLeft"
    $label.ForeColor = $colorMuted
    return $label
}

function New-ValueBox {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = "Fill"
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    $box.ReadOnly = $true
    $box.BorderStyle = "FixedSingle"
    $box.BackColor = $colorPanelSoft
    $box.ForeColor = $colorInk
    return $box
}

function New-EntryBox {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = "Fill"
    $box.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    $box.BorderStyle = "FixedSingle"
    $box.BackColor = $colorPanel
    $box.ForeColor = $colorInk
    return $box
}

function New-OptionsTextBox {
    param([Parameter(Mandatory = $true)][string]$Text)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Dock = "Fill"
    $box.Multiline = $true
    $box.ScrollBars = "Vertical"
    $box.ReadOnly = $true
    $box.Font = $statusFont
    $box.BackColor = $colorPanelSoft
    $box.ForeColor = $colorInk
    $box.BorderStyle = "FixedSingle"
    $box.Text = $Text.Trim()
    return $box
}

function New-SectionPanel {
    param([Parameter(Mandatory = $true)][string]$Title)

    $outer = New-Object System.Windows.Forms.Panel
    $outer.Dock = "Fill"
    $outer.BackColor = $colorPanel
    $outer.Padding = New-Object System.Windows.Forms.Padding(14)
    Add-SoftBorder -Control $outer

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = "Fill"
    $layout.ColumnCount = 1
    $layout.RowCount = 2
    $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.Dock = "Fill"
    $titleLabel.Font = $sectionFont
    $titleLabel.ForeColor = $colorInk
    $titleLabel.TextAlign = "MiddleLeft"

    $content = New-Object System.Windows.Forms.Panel
    $content.Dock = "Fill"
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
$form.Text = "Smart Intune Hybrid Join Toolkit - LOT Launcher"
$form.StartPosition = "CenterScreen"
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
$rootLayout.Dock = "Fill"
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
$headerPanel.Dock = "Fill"
$headerPanel.BackColor = $colorPanel
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(18)
Add-SoftBorder -Control $headerPanel

$headerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$headerLayout.Dock = "Fill"
$headerLayout.ColumnCount = 2
$headerLayout.RowCount = 1
$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$headerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 142))) | Out-Null

$headerTextPanel = New-Object System.Windows.Forms.TableLayoutPanel
$headerTextPanel.Dock = "Fill"
$headerTextPanel.ColumnCount = 1
$headerTextPanel.RowCount = 4
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$headerTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null

$badgeLabel = New-Object System.Windows.Forms.Label
$badgeLabel.Text = "SMARTM365"
$badgeLabel.AutoSize = $true
$badgeLabel.BackColor = $colorHeaderPanel
$badgeLabel.ForeColor = $colorAccentDark
$badgeLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8)
$badgeLabel.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Intune Hybrid Join LOT Launcher"
$titleLabel.Dock = "Fill"
$titleLabel.ForeColor = $colorInk
$titleLabel.Font = $titleFont
$titleLabel.TextAlign = "MiddleLeft"

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Run an existing LOT or create an empty LOT ready for Computers.txt."
$subtitleLabel.Dock = "Fill"
$subtitleLabel.ForeColor = $colorMuted
$subtitleLabel.TextAlign = "MiddleLeft"

$psexecLabel = New-Object System.Windows.Forms.Label
$psexecLabel.Dock = "Fill"
$psexecLabel.ForeColor = $colorInk
$psexecLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
$psexecLabel.TextAlign = "MiddleLeft"

$headerTextPanel.Controls.Add($badgeLabel, 0, 0)
$headerTextPanel.Controls.Add($titleLabel, 0, 1)
$headerTextPanel.Controls.Add($subtitleLabel, 0, 2)
$headerTextPanel.Controls.Add($psexecLabel, 0, 3)

$logoCard = New-Object System.Windows.Forms.Panel
$logoCard.Dock = "Fill"
$logoCard.BackColor = $colorPanel
$logoCard.Padding = New-Object System.Windows.Forms.Padding(10)
$logoCard.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
Add-SoftBorder -Control $logoCard

if ($script:LogoImage) {
    $logoPicture = New-Object System.Windows.Forms.PictureBox
    $logoPicture.Dock = "Fill"
    $logoPicture.SizeMode = "Zoom"
    $logoPicture.Image = $script:LogoImage
    $logoPicture.BackColor = $colorPanel
    $logoCard.Controls.Add($logoPicture)
}
else {
    $logoFallback = New-Object System.Windows.Forms.Label
    $logoFallback.Dock = "Fill"
    $logoFallback.Text = "Workplace`r`nCloudHub"
    $logoFallback.ForeColor = $colorAccent
    $logoFallback.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
    $logoFallback.TextAlign = "MiddleCenter"
    $logoCard.Controls.Add($logoFallback)
}

$psexecToolkitPath = Join-Path $toolkitRoot "Scripts\PsExec.exe"
$psexecSystem32Path = Join-Path $env:WINDIR "System32\PsExec.exe"
if (Test-Path -LiteralPath $psexecToolkitPath -PathType Leaf) {
    $psexecLabel.Text = "PsExec ready in Scripts"
    $psexecLabel.ForeColor = $colorSuccess
}
elseif (Test-Path -LiteralPath $psexecSystem32Path -PathType Leaf) {
    $psexecLabel.Text = "PsExec ready in System32"
    $psexecLabel.ForeColor = $colorSuccess
}
else {
    $psexecLabel.Text = "PsExec not local"
    $psexecLabel.ForeColor = $colorWarning
}

$headerLayout.Controls.Add($headerTextPanel, 0, 0)
$headerLayout.Controls.Add($logoCard, 1, 0)
$headerPanel.Controls.Add($headerLayout)

$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Dock = "Fill"
$actionPanel.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)
$actionPanel.BackColor = $colorPanel
$actionPanel.Padding = New-Object System.Windows.Forms.Padding(12)
Add-SoftBorder -Control $actionPanel

$actionLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionLayout.Dock = "Fill"
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
$actionTextPanel.Dock = "Fill"
$actionTextPanel.ColumnCount = 1
$actionTextPanel.RowCount = 3
$actionTextPanel.BackColor = $colorPanel
$actionTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 24))) | Out-Null
$actionTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) | Out-Null
$actionTextPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$actionTitleLabel = New-Object System.Windows.Forms.Label
$actionTitleLabel.Dock = "Fill"
$actionTitleLabel.Text = "LOT: none selected"
$actionTitleLabel.ForeColor = $colorInk
$actionTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
$actionTitleLabel.TextAlign = "MiddleLeft"

$actionSubtitleLabel = New-Object System.Windows.Forms.Label
$actionSubtitleLabel.Dock = "Fill"
$actionSubtitleLabel.Text = "Select or create a LOT."
$actionSubtitleLabel.ForeColor = $colorAccent
$actionSubtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 8)
$actionSubtitleLabel.TextAlign = "MiddleLeft"

$actionStatusLabel = New-Object System.Windows.Forms.Label
$actionStatusLabel.Dock = "Fill"
$actionStatusLabel.Text = "Ready"
$actionStatusLabel.ForeColor = $colorMuted
$actionStatusLabel.Font = $smallFont
$actionStatusLabel.TextAlign = "MiddleLeft"

$actionTextPanel.Controls.Add($actionTitleLabel, 0, 0)
$actionTextPanel.Controls.Add($actionSubtitleLabel, 0, 1)
$actionTextPanel.Controls.Add($actionStatusLabel, 0, 2)

$actionButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionButtonsPanel.Dock = "Fill"
$actionButtonsPanel.FlowDirection = "RightToLeft"
$actionButtonsPanel.WrapContents = $false
$actionButtonsPanel.BackColor = $colorPanel
$actionButtonsPanel.Padding = New-Object System.Windows.Forms.Padding(0, 18, 0, 0)

$actionLaunchAllButton = New-Object System.Windows.Forms.Button
$actionLaunchAllButton.Text = "Launch all"
$actionLaunchAllButton.Width = 130
Set-ButtonEnabledStyle -Button $actionLaunchAllButton -Enabled $false -EnabledBackColor $colorAccent

$actionRefreshButton = New-Object System.Windows.Forms.Button
$actionRefreshButton.Text = "Refresh"
$actionRefreshButton.Width = 118
Set-FlatButtonStyle -Button $actionRefreshButton -BackColor $colorPanel -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$actionButtonsPanel.Controls.Add($actionLaunchAllButton)
$actionButtonsPanel.Controls.Add($actionRefreshButton)

$actionLayout.Controls.Add($actionDot, 0, 0)
$actionLayout.Controls.Add($actionTextPanel, 1, 0)
$actionLayout.Controls.Add($actionButtonsPanel, 2, 0)
$actionPanel.Controls.Add($actionLayout)

$tabShell = New-Object System.Windows.Forms.TableLayoutPanel
$tabShell.Dock = "Fill"
$tabShell.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 12)
$tabShell.BackColor = $colorBackground
$tabShell.ColumnCount = 1
$tabShell.RowCount = 2
$tabShell.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$tabShell.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) | Out-Null
$tabShell.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$tabStrip = New-Object System.Windows.Forms.FlowLayoutPanel
$tabStrip.Dock = "Fill"
$tabStrip.FlowDirection = "LeftToRight"
$tabStrip.WrapContents = $false
$tabStrip.BackColor = $colorBackground
$tabStrip.Padding = New-Object System.Windows.Forms.Padding(0)
$tabStrip.Margin = New-Object System.Windows.Forms.Padding(0)

$tabContentPanel = New-Object System.Windows.Forms.Panel
$tabContentPanel.Dock = "Fill"
$tabContentPanel.BackColor = $colorBackground
$tabContentPanel.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)

$tabShell.Controls.Add($tabStrip, 0, 0)
$tabShell.Controls.Add($tabContentPanel, 0, 1)

$script:DeviceRegistrationTabHeaders = New-Object System.Collections.ArrayList

$existingTab = New-Object System.Windows.Forms.Panel
$existingTab.Dock = "Fill"
$existingTab.BackColor = $colorBackground

$singleTab = New-Object System.Windows.Forms.Panel
$singleTab.Dock = "Fill"
$singleTab.BackColor = $colorBackground

$newTab = New-Object System.Windows.Forms.Panel
$newTab.Dock = "Fill"
$newTab.BackColor = $colorBackground

$optionsTab = New-Object System.Windows.Forms.Panel
$optionsTab.Dock = "Fill"
$optionsTab.BackColor = $colorBackground

$tabContentPanel.Controls.Add($existingTab)
$tabContentPanel.Controls.Add($singleTab)
$tabContentPanel.Controls.Add($newTab)
$tabContentPanel.Controls.Add($optionsTab)

$existingTabHeader = New-DeviceRegistrationTabHeader -Text "Existing LOT" -Page $existingTab
$singleTabHeader = New-DeviceRegistrationTabHeader -Text "Single PC" -Page $singleTab
$newTabHeader = New-DeviceRegistrationTabHeader -Text "New LOT" -Page $newTab
$optionsTabHeader = New-DeviceRegistrationTabHeader -Text "Options" -Page $optionsTab
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
$existingLayout.Dock = "Fill"
$existingLayout.ColumnCount = 1
$existingLayout.RowCount = 2
$existingLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$existingLayout.BackColor = $colorBackground
$existingLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$existingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 232))) | Out-Null
$existingLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$existingSection = New-SectionPanel -Title "Available lots"
$activitySection = New-SectionPanel -Title "Activity"

$existingTable = New-Object System.Windows.Forms.TableLayoutPanel
$existingTable.Dock = "Fill"
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
$lotCombo.Dock = "Fill"
$lotCombo.DropDownStyle = "DropDownList"
$lotCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$lotCombo.FlatStyle = "Flat"

$refreshLotsButton = New-Object System.Windows.Forms.Button
$refreshLotsButton.Text = "Refresh"
$refreshLotsButton.Dock = "Fill"
Set-FlatButtonStyle -Button $refreshLotsButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$openLotFolderButton = New-Object System.Windows.Forms.Button
$openLotFolderButton.Text = "Folder"
$openLotFolderButton.Dock = "Fill"
Set-FlatButtonStyle -Button $openLotFolderButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$deviceCountBox = New-ValueBox
$adScopeBox = New-ValueBox

$existingModeCombo = New-Object System.Windows.Forms.ComboBox
$existingModeCombo.Dock = "Fill"
$existingModeCombo.DropDownStyle = "DropDownList"
$existingModeCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$existingModeCombo.FlatStyle = "Flat"
$existingModeCombo.Items.AddRange([object[]]@("Loop", "Once", "LoopIgnoreRunGuard", "OnceIgnoreRunGuard"))
$existingModeCombo.SelectedIndex = 0

$globalConcurrencyLimitBox = New-Object System.Windows.Forms.NumericUpDown
$globalConcurrencyLimitBox.Dock = "Fill"
$globalConcurrencyLimitBox.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$globalConcurrencyLimitBox.Minimum = 1
$globalConcurrencyLimitBox.Maximum = 50
$globalConcurrencyLimitBox.Value = 15

$openComputersButton = New-Object System.Windows.Forms.Button
$openComputersButton.Text = "Computers"
$openComputersButton.Dock = "Fill"
Set-FlatButtonStyle -Button $openComputersButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$openAdDomainButton = New-Object System.Windows.Forms.Button
$openAdDomainButton.Text = "AD domain"
$openAdDomainButton.Dock = "Fill"
Set-FlatButtonStyle -Button $openAdDomainButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$launchExistingButton = New-Object System.Windows.Forms.Button
$launchExistingButton.Text = "Launch"
$launchExistingButton.Dock = "Fill"
Set-ButtonEnabledStyle -Button $launchExistingButton -Enabled $false -EnabledBackColor $colorAccent

$launchAllLotsButton = New-Object System.Windows.Forms.Button
$launchAllLotsButton.Text = "Launch all"
$launchAllLotsButton.Dock = "Fill"
Set-ButtonEnabledStyle -Button $launchAllLotsButton -Enabled $false -EnabledBackColor $colorAccent

$refreshWrappersButton = New-Object System.Windows.Forms.Button
$refreshWrappersButton.Text = "Wrappers"
$refreshWrappersButton.Dock = "Fill"
Set-FlatButtonStyle -Button $refreshWrappersButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$existingTable.Controls.Add((New-Label "LOT"), 0, 0)
$existingTable.Controls.Add($lotCombo, 1, 0)
$existingTable.SetColumnSpan($lotCombo, 2)
$existingTable.Controls.Add($refreshLotsButton, 3, 0)
$existingTable.Controls.Add($openLotFolderButton, 4, 0)

$existingTable.Controls.Add((New-Label "Devices"), 0, 1)
$existingTable.Controls.Add($deviceCountBox, 1, 1)
$existingTable.SetColumnSpan($deviceCountBox, 2)
$existingTable.Controls.Add($openComputersButton, 3, 1)
$existingTable.Controls.Add($openAdDomainButton, 4, 1)

$existingTable.Controls.Add((New-Label "AD scope"), 0, 2)
$existingTable.Controls.Add($adScopeBox, 1, 2)
$existingTable.SetColumnSpan($adScopeBox, 4)

$existingTable.Controls.Add((New-Label "Limit"), 0, 3)
$existingTable.Controls.Add($globalConcurrencyLimitBox, 1, 3)
$existingTable.SetColumnSpan($globalConcurrencyLimitBox, 4)

$existingTable.Controls.Add((New-Label "Launch"), 0, 4)
$existingTable.Controls.Add($existingModeCombo, 1, 4)
$existingTable.Controls.Add($refreshWrappersButton, 2, 4)
$existingTable.Controls.Add($launchExistingButton, 3, 4)
$existingTable.SetColumnSpan($launchExistingButton, 2)

$existingSection.Content.Controls.Add($existingTable)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Dock = "Fill"
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$statusBox.Font = $statusFont
$statusBox.BackColor = $colorPanelSoft
$statusBox.ForeColor = $colorInk
$statusBox.BorderStyle = "FixedSingle"
$activitySection.Content.Controls.Add($statusBox)

$existingLayout.Controls.Add($existingSection.Panel, 0, 0)
$existingLayout.Controls.Add($activitySection.Panel, 0, 1)
$existingTab.Controls.Add($existingLayout)

$singleLayout = New-Object System.Windows.Forms.TableLayoutPanel
$singleLayout.Dock = "Fill"
$singleLayout.ColumnCount = 1
$singleLayout.RowCount = 2
$singleLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$singleLayout.BackColor = $colorBackground
$singleLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$singleLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 174))) | Out-Null
$singleLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$singleSection = New-SectionPanel -Title "Run one computer"
$singleTable = New-Object System.Windows.Forms.TableLayoutPanel
$singleTable.Dock = "Top"
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
$singleModeCombo.Dock = "Fill"
$singleModeCombo.DropDownStyle = "DropDownList"
$singleModeCombo.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)
$singleModeCombo.FlatStyle = "Flat"
$singleModeCombo.Items.AddRange([object[]]@("Once","OnceIgnoreRunGuard","Loop","LoopIgnoreRunGuard"))
$singleModeCombo.SelectedIndex = 0

$singleRunPathBox = New-ValueBox
$script:SingleComputerRunRoot = Get-SingleComputerRunRoot -ToolkitKey "IntuneHybridJoinToolkit"
$singleRunPathBox.Text = $script:SingleComputerRunRoot

$openSingleRunFolderButton = New-Object System.Windows.Forms.Button
$openSingleRunFolderButton.Text = "Open run folder"
$openSingleRunFolderButton.Dock = "Fill"
Set-FlatButtonStyle -Button $openSingleRunFolderButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$launchSingleComputerButton = New-Object System.Windows.Forms.Button
$launchSingleComputerButton.Text = "Launch"
$launchSingleComputerButton.Dock = "Fill"
Set-ButtonEnabledStyle -Button $launchSingleComputerButton -Enabled $true -EnabledBackColor $colorAccent

$singleTable.Controls.Add((New-Label "Computer"), 0, 0)
$singleTable.Controls.Add($singleComputerBox, 1, 0)
$singleTable.Controls.Add((New-Label "Mode"), 2, 0)
$singleTable.Controls.Add($singleModeCombo, 3, 0)

$singleTable.Controls.Add((New-Label "Run folder"), 0, 1)
$singleTable.Controls.Add($singleRunPathBox, 1, 1)
$singleTable.SetColumnSpan($singleRunPathBox, 2)
$singleTable.Controls.Add($openSingleRunFolderButton, 3, 1)

$singleTable.Controls.Add((New-Label "Launch"), 0, 2)
$singleTable.Controls.Add($launchSingleComputerButton, 1, 2)
$singleTable.SetColumnSpan($launchSingleComputerButton, 3)

$singleSection.Content.Controls.Add($singleTable)
$singleLayout.Controls.Add($singleSection.Panel, 0, 0)
$singleTab.Controls.Add($singleLayout)

$newLayout = New-Object System.Windows.Forms.TableLayoutPanel
$newLayout.Dock = "Fill"
$newLayout.ColumnCount = 1
$newLayout.RowCount = 2
$newLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$newLayout.BackColor = $colorBackground
$newLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$newLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 174))) | Out-Null
$newLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$newSection = New-SectionPanel -Title "Create an empty LOT"
$newTable = New-Object System.Windows.Forms.TableLayoutPanel
$newTable.Dock = "Top"
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
$createLotButton.Text = "Create"
$createLotButton.Dock = "Fill"
Set-FlatButtonStyle -Button $createLotButton -BackColor $colorAccent -ForeColor ([System.Drawing.Color]::White) -BorderColor $colorAccent

$openNewComputersButton = New-Object System.Windows.Forms.Button
$openNewComputersButton.Text = "Open Computers"
$openNewComputersButton.Dock = "Fill"
Set-ButtonEnabledStyle -Button $openNewComputersButton -Enabled $false

$newTable.Controls.Add((New-Label "LOT name"), 0, 0)
$newTable.Controls.Add($newLotNameBox, 1, 0)
$newTable.Controls.Add($createLotButton, 2, 0)

$newTable.Controls.Add((New-Label "Computers.txt"), 0, 1)
$newTable.Controls.Add($newComputersPathBox, 1, 1)
$newTable.Controls.Add($openNewComputersButton, 2, 1)

$newSection.Content.Controls.Add($newTable)
$newLayout.Controls.Add($newSection.Panel, 0, 0)
$newTab.Controls.Add($newLayout)

$optionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$optionsLayout.Dock = "Fill"
$optionsLayout.ColumnCount = 1
$optionsLayout.RowCount = 1
$optionsLayout.Padding = New-Object System.Windows.Forms.Padding(10)
$optionsLayout.BackColor = $colorBackground
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$optionsSection = New-SectionPanel -Title "Toolkit options"
$optionsScrollPanel = New-Object System.Windows.Forms.Panel
$optionsScrollPanel.Dock = "Fill"
$optionsScrollPanel.AutoScroll = $true
$optionsScrollPanel.BackColor = $colorPanel

$optionsTable = New-Object System.Windows.Forms.TableLayoutPanel
$optionsTable.Dock = "Top"
$optionsTable.AutoSize = $true
$optionsTable.ColumnCount = 4
$optionsTable.RowCount = 10
$optionsTable.BackColor = $colorPanel
$optionsTable.Padding = New-Object System.Windows.Forms.Padding(0, 2, 12, 2)
$optionsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 190))) | Out-Null
$optionsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 170))) | Out-Null
$optionsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 190))) | Out-Null
$optionsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
for ($i = 0; $i -lt 10; $i++) {
    $optionsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36))) | Out-Null
}

function New-OptionNumber {
    param([int]$Minimum, [int]$Maximum, [int]$Value)

    $number = New-Object System.Windows.Forms.NumericUpDown
    $number.Dock = "Fill"
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
    $check.Dock = "Fill"
    $check.Checked = $Checked
    $check.ForeColor = $colorInk
    return $check
}

function Set-OptionNumberFromEnv {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Forms.NumericUpDown]$Control,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $value = [Environment]::GetEnvironmentVariable($Name, 'Process')
    $parsed = 0
    if ([int]::TryParse($value, [ref]$parsed)) {
        if ($parsed -ge [int]$Control.Minimum -and $parsed -le [int]$Control.Maximum) {
            $Control.Value = $parsed
        }
    }
}

function Get-EnvSwitch {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [bool]$Default = $false
    )

    $value = [Environment]::GetEnvironmentVariable($Name, 'Process')
    if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
    return @('1','true','yes','on') -contains $value.Trim().ToLowerInvariant()
}

$optionThrottleBox = New-OptionNumber -Minimum 1 -Maximum 200 -Value 10
$optionGlobalConcurrencyLimitBox = New-OptionNumber -Minimum 1 -Maximum 200 -Value 15
$optionGlobalConcurrencyLeaseTimeoutBox = New-OptionNumber -Minimum 0 -Maximum 1440 -Value 0
$optionDelayBetweenCyclesBox = New-OptionNumber -Minimum 0 -Maximum 1440 -Value 5
$optionMaxCyclesBox = New-OptionNumber -Minimum 0 -Maximum 1000 -Value 0
$optionPsExecTimeoutBox = New-OptionNumber -Minimum 0 -Maximum 1440 -Value 120
$optionIntuneRetrySleepBox = New-OptionNumber -Minimum 0 -Maximum 1440 -Value 5
$optionIntuneRetryMaxBox = New-OptionNumber -Minimum 0 -Maximum 100 -Value 5
$optionStaleCleanupDelayBox = New-OptionNumber -Minimum 0 -Maximum 86400 -Value 60
$optionRebootDelayBox = New-OptionNumber -Minimum 0 -Maximum 86400 -Value 180

$optionAdDomainBox = New-EntryBox
$optionAdDomainBox.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 3)

$adDomainDefault = [Environment]::GetEnvironmentVariable('EHJIR_AD_DOMAIN', 'Process')
if (-not [string]::IsNullOrWhiteSpace($adDomainDefault)) {
    $optionAdDomainBox.Text = $adDomainDefault
}

Set-OptionNumberFromEnv -Control $optionThrottleBox -Name 'EHJIR_THROTTLE'
Set-OptionNumberFromEnv -Control $optionGlobalConcurrencyLimitBox -Name 'EHJIR_GLOBAL_CONCURRENCY_LIMIT'
Set-OptionNumberFromEnv -Control $optionGlobalConcurrencyLeaseTimeoutBox -Name 'EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES'
Set-OptionNumberFromEnv -Control $optionDelayBetweenCyclesBox -Name 'EHJIR_DELAY_BETWEEN_CYCLES_MINUTES'
Set-OptionNumberFromEnv -Control $optionIntuneRetrySleepBox -Name 'EHJIR_INTUNE_RETRY_SLEEP_MINUTES'
Set-OptionNumberFromEnv -Control $optionIntuneRetryMaxBox -Name 'EHJIR_INTUNE_RETRY_MAX_RETRIES'
Set-OptionNumberFromEnv -Control $optionStaleCleanupDelayBox -Name 'EHJIR_STALE_CLEANUP_DELAY_SECONDS'
Set-OptionNumberFromEnv -Control $optionRebootDelayBox -Name 'EHJIR_REBOOT_DELAY_SECONDS'
Set-OptionNumberFromEnv -Control $optionPsExecTimeoutBox -Name 'EHJIR_PSEXEC_TIMEOUT_MINUTES'

$optionDryRunCheck = New-OptionCheck -Text "Dry run"
$optionAuditOnlyCheck = New-OptionCheck -Text "Audit only"
$optionAllowDsregLeaveCheck = New-OptionCheck -Text "Allow dsreg leave"
$optionAllowStaleIntuneCleanupCheck = New-OptionCheck -Text "Allow stale Intune cleanup"
$optionAllowRebootWhenNoUserCheck = New-OptionCheck -Text "Allow reboot no user"
$optionAllowRebootAfterLeaveCheck = New-OptionCheck -Text "Allow reboot after leave"
$optionIgnoreRunGuardEveryCycleCheck = New-OptionCheck -Text "Ignore run guard every cycle"
$optionRemoveNonIntuneMdmCheck = New-OptionCheck -Text "Remove non-Intune MDM"
$optionKeepCentralHistoryCheck = New-OptionCheck -Text "Keep central log history"
$optionNoCentralCollectionCheck = New-OptionCheck -Text "No central log collection"
$optionSkipPostCycleInventoryCheck = New-OptionCheck -Text "Skip post-cycle Intune inventory"
$optionSkipVirtualMachinesCheck = New-OptionCheck -Text "Skip virtual machines"
$optionAllowDsregLeaveCheck.Checked = Get-EnvSwitch -Name 'EHJIR_ALLOW_DSREG_LEAVE' -Default $true
$optionAllowStaleIntuneCleanupCheck.Checked = Get-EnvSwitch -Name 'EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT' -Default $true
$optionAllowRebootWhenNoUserCheck.Checked = Get-EnvSwitch -Name 'EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER' -Default $true
$optionAllowRebootAfterLeaveCheck.Checked = Get-EnvSwitch -Name 'EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE' -Default $true
$optionSkipVirtualMachinesCheck.Checked = Get-EnvSwitch -Name 'EHJIR_SKIP_VIRTUAL_MACHINES' -Default $true

$optionsTable.Controls.Add((New-Label "Throttle per LOT"), 0, 0)
$optionsTable.Controls.Add($optionThrottleBox, 1, 0)
$optionsTable.Controls.Add((New-Label "Delay between cycles"), 2, 0)
$optionsTable.Controls.Add($optionDelayBetweenCyclesBox, 3, 0)

$optionsTable.Controls.Add((New-Label "Max cycles"), 0, 1)
$optionsTable.Controls.Add($optionMaxCyclesBox, 1, 1)
$optionsTable.Controls.Add((New-Label "PsExec timeout min"), 2, 1)
$optionsTable.Controls.Add($optionPsExecTimeoutBox, 3, 1)

$optionsTable.Controls.Add((New-Label "Intune retry sleep min"), 0, 2)
$optionsTable.Controls.Add($optionIntuneRetrySleepBox, 1, 2)
$optionsTable.Controls.Add((New-Label "Intune retry max"), 2, 2)
$optionsTable.Controls.Add($optionIntuneRetryMaxBox, 3, 2)

$optionsTable.Controls.Add((New-Label "Stale cleanup seconds"), 0, 3)
$optionsTable.Controls.Add($optionStaleCleanupDelayBox, 1, 3)
$optionsTable.Controls.Add((New-Label "Reboot delay seconds"), 2, 3)
$optionsTable.Controls.Add($optionRebootDelayBox, 3, 3)

$optionsTable.Controls.Add((New-Label "AD domain override"), 0, 4)
$optionsTable.Controls.Add($optionAdDomainBox, 1, 4)
$optionsTable.SetColumnSpan($optionAdDomainBox, 3)

$optionsTable.Controls.Add($optionAllowDsregLeaveCheck, 0, 5)
$optionsTable.Controls.Add($optionAllowStaleIntuneCleanupCheck, 1, 5)
$optionsTable.Controls.Add($optionAllowRebootWhenNoUserCheck, 2, 5)
$optionsTable.Controls.Add($optionAllowRebootAfterLeaveCheck, 3, 5)

$optionsTable.Controls.Add($optionDryRunCheck, 0, 6)
$optionsTable.Controls.Add($optionAuditOnlyCheck, 1, 6)
$optionsTable.Controls.Add($optionIgnoreRunGuardEveryCycleCheck, 2, 6)
$optionsTable.Controls.Add($optionRemoveNonIntuneMdmCheck, 3, 6)

$optionsTable.Controls.Add($optionKeepCentralHistoryCheck, 0, 7)
$optionsTable.Controls.Add($optionNoCentralCollectionCheck, 1, 7)
$optionsTable.Controls.Add($optionSkipPostCycleInventoryCheck, 2, 7)
$optionsTable.Controls.Add($optionSkipVirtualMachinesCheck, 3, 7)

$optionsTable.Controls.Add((New-Label "Global worker limit"), 0, 8)
$optionsTable.Controls.Add($optionGlobalConcurrencyLimitBox, 1, 8)
$optionsTable.Controls.Add((New-Label "Global lease timeout min"), 2, 8)
$optionsTable.Controls.Add($optionGlobalConcurrencyLeaseTimeoutBox, 3, 8)

$optionsNote = New-Object System.Windows.Forms.Label
$optionsNote.Text = "Default LOT wrappers already enable guarded dsreg leave, stale Intune cleanup, and controlled reboot paths. These controls add or override launch arguments."
$optionsNote.Dock = "Fill"
$optionsNote.ForeColor = $colorMuted
$optionsNote.TextAlign = "MiddleLeft"
$optionsTable.Controls.Add($optionsNote, 0, 9)
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
$footerPanel.Dock = "Fill"
$footerPanel.BackColor = $colorPanel
$footerPanel.Padding = New-Object System.Windows.Forms.Padding(12)
Add-SoftBorder -Control $footerPanel

$footerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$footerLayout.Dock = "Fill"
$footerLayout.ColumnCount = 2
$footerLayout.RowCount = 1
$footerLayout.BackColor = $colorPanel
$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$footerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120))) | Out-Null

$footerStatus = New-Object System.Windows.Forms.Label
$footerStatus.Dock = "Fill"
$footerStatus.Text = "Ready"
$footerStatus.TextAlign = "MiddleLeft"
$footerStatus.ForeColor = $colorMuted
$footerStatus.Font = $smallFont

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Width = 94
Set-FlatButtonStyle -Button $closeButton -BackColor $colorPanel -ForeColor $colorInk -BorderColor $colorTextBoxBorder

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = "Fill"
$buttonPanel.FlowDirection = "RightToLeft"
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

    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
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

    if ([int]$optionMaxCyclesBox.Value -gt 0) {
        [void]$arguments.Add("-MaxCycles")
        [void]$arguments.Add([string][int]$optionMaxCyclesBox.Value)
    }

    if ($optionDryRunCheck.Checked) { [void]$arguments.Add("-DryRun") }
    if ($optionAuditOnlyCheck.Checked) { [void]$arguments.Add("-AuditOnly") }
    if ($optionIgnoreRunGuardEveryCycleCheck.Checked) { [void]$arguments.Add("-IgnoreRunGuardEveryCycle") }
    if ($optionRemoveNonIntuneMdmCheck.Checked) { [void]$arguments.Add("-AllowRemoveNonIntuneMdmEnrollment") }
    if ($optionKeepCentralHistoryCheck.Checked) { [void]$arguments.Add("-KeepCentralLogHistory") }
    if ($optionNoCentralCollectionCheck.Checked) { [void]$arguments.Add("-NoCentralLogCollection") }
    if ($optionSkipPostCycleInventoryCheck.Checked) { [void]$arguments.Add("-SkipPostCycleIntuneInventory") }

    return @($arguments.ToArray())
}

function Get-ToolkitOptionEnvironment {
    $environment = @{
        EHJIR_THROTTLE = [string][int]$optionThrottleBox.Value
        EHJIR_GLOBAL_CONCURRENCY_LIMIT = [string][int]$optionGlobalConcurrencyLimitBox.Value
        EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = [string][int]$optionGlobalConcurrencyLeaseTimeoutBox.Value
        EHJIR_DELAY_BETWEEN_CYCLES_MINUTES = [string][int]$optionDelayBetweenCyclesBox.Value
        EHJIR_INTUNE_RETRY_SLEEP_MINUTES = [string][int]$optionIntuneRetrySleepBox.Value
        EHJIR_INTUNE_RETRY_MAX_RETRIES = [string][int]$optionIntuneRetryMaxBox.Value
        EHJIR_STALE_CLEANUP_DELAY_SECONDS = [string][int]$optionStaleCleanupDelayBox.Value
        EHJIR_REBOOT_DELAY_SECONDS = [string][int]$optionRebootDelayBox.Value
        EHJIR_PSEXEC_TIMEOUT_MINUTES = [string][int]$optionPsExecTimeoutBox.Value
        EHJIR_ALLOW_DSREG_LEAVE = if ($optionAllowDsregLeaveCheck.Checked) { "1" } else { "0" }
        EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT = if ($optionAllowStaleIntuneCleanupCheck.Checked) { "1" } else { "0" }
        EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER = if ($optionAllowRebootWhenNoUserCheck.Checked) { "1" } else { "0" }
        EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE = if ($optionAllowRebootAfterLeaveCheck.Checked) { "1" } else { "0" }
        EHJIR_SKIP_VIRTUAL_MACHINES = if ($optionSkipVirtualMachinesCheck.Checked) { "1" } else { "0" }
    }

    if (-not [string]::IsNullOrWhiteSpace($optionAdDomainBox.Text)) {
        $environment["EHJIR_AD_DOMAIN"] = $optionAdDomainBox.Text.Trim()
    }

    return $environment
}

function Clear-LotDetails {
    $script:SelectedLotSummary = $null
    $deviceCountBox.Text = ""
    $adScopeBox.Text = ""
    $actionTitleLabel.Text = "LOT: none selected"
    $actionSubtitleLabel.Text = "Select or create a LOT."
    $actionStatusLabel.Text = "Ready"
    Set-ButtonEnabledStyle -Button $launchExistingButton -Enabled $false -EnabledBackColor $colorAccent
}

function Get-LaunchableLotSummaries {
    $summaries = New-Object System.Collections.ArrayList
    foreach ($lot in @($script:LotList)) {
        try {
            $summary = Get-LotSummary -RootPath $toolkitRoot -LotPath $lot.FullName
            if ($summary.ComputerCount -gt 0 -and $summary.WrappersReady) {
                [void]$summaries.Add($summary)
            }
        }
        catch {
            Add-Status ("Skipping {0}: {1}" -f $lot.Name,$_.Exception.Message)
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
    $openAdDomainButton.Enabled = $HasLots
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
    $summary = Get-LotSummary -RootPath $toolkitRoot -LotPath $lot.FullName
    $script:SelectedLotSummary = $summary

    $deviceCountBox.Text = [string]$summary.ComputerCount
    $adScopeBox.Text = $summary.AdScope
    $actionTitleLabel.Text = "LOT: {0}" -f $summary.Name
    $actionSubtitleLabel.Text = "{0} device(s) - {1}" -f $summary.ComputerCount, $summary.AdScope

    if (-not $summary.WrappersReady) {
        Add-Status ("Missing LOT wrappers: {0}" -f ($summary.MissingWrappers -join ", "))
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
        ""
    }
    $script:LotList = @(Get-LotFolders -RootPath $toolkitRoot)

    $lotCombo.Items.Clear()
    foreach ($lot in $script:LotList) {
        $lotCombo.Items.Add($lot.Name) | Out-Null
    }

    if ($lotCombo.Items.Count -eq 0) {
        Clear-LotDetails
        Update-ExistingLotControlState -HasLots $false
        Add-Status "No operational LOT-* folder found."
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
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Refresh failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
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
        if (-not $script:SelectedLotSummary) { throw "Select a LOT first." }
        Open-FolderPath -Path $script:SelectedLotSummary.Path
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Open folder failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$openComputersButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw "Select a LOT first." }
        Open-TextFile -Path $script:SelectedLotSummary.ComputersPath
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$openAdDomainButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw "Select a LOT first." }
        Open-TextFile -Path $script:SelectedLotSummary.AdDomainPath
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
    }
})

$refreshWrappersButton.Add_Click({
    try {
        Invoke-LotWrapperRefresh -RootPath $toolkitRoot
        Add-Status "LOT wrappers refreshed."
        Refresh-LotList
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Wrapper refresh failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchExistingButton.Add_Click({
    try {
        if (-not $script:SelectedLotSummary) { throw "Select a LOT first." }
        if ($script:SelectedLotSummary.ComputerCount -le 0) { throw "Computers.txt is empty." }

        $optionArguments = @(Get-ToolkitOptionArguments)
        $optionEnvironment = Get-ToolkitOptionEnvironment
        Add-Status ("Launching {0} in {1} mode. Global limit={2}. Args={3}; env={4}." -f $script:SelectedLotSummary.Name,$existingModeCombo.SelectedItem,[int]$globalConcurrencyLimitBox.Value,$optionArguments.Count,$optionEnvironment.Count)
        Start-ToolkitLot -LotPath $script:SelectedLotSummary.Path -Mode ([string]$existingModeCombo.SelectedItem) -GlobalConcurrencyLimit ([int]$globalConcurrencyLimitBox.Value) -GlobalConcurrencyLeaseTimeoutMinutes ([int]$optionGlobalConcurrencyLeaseTimeoutBox.Value) -AdditionalArguments $optionArguments -EnvironmentVariables $optionEnvironment
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "LOT launch failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchAllLotsButton.Add_Click({
    try {
        $launchableLots = @(Get-LaunchableLotSummaries)
        if ($launchableLots.Count -eq 0) {
            throw "No launchable LOT found. Check Computers.txt and wrapper files."
        }

        $mode = [string]$existingModeCombo.SelectedItem
        $limit = [int]$globalConcurrencyLimitBox.Value
        $optionArguments = @(Get-ToolkitOptionArguments)
        $optionEnvironment = Get-ToolkitOptionEnvironment
        Add-Status ("Launching {0}/{1} LOT folder(s) in {2} mode. Global limit={3}. Delay={4}s. Args={5}; env={6}." -f $launchableLots.Count,$script:LotList.Count,$mode,$limit,$launchAllLotStartDelaySeconds,$optionArguments.Count,$optionEnvironment.Count)

        for ($lotIndex = 0; $lotIndex -lt $launchableLots.Count; $lotIndex++) {
            $lotSummary = $launchableLots[$lotIndex]
            if ($lotIndex -gt 0 -and $launchAllLotStartDelaySeconds -gt 0) {
                Add-Status ("Waiting {0}s before launching {1}." -f $launchAllLotStartDelaySeconds,$lotSummary.Name)
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
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Launch all failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$launchSingleComputerButton.Add_Click({
    try {
        $computerName = $singleComputerBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($computerName)) { throw "Enter a computer name." }

        $optionArguments = @(Get-ToolkitOptionArguments)
        $optionEnvironment = Get-ToolkitOptionEnvironment
        Add-Status ("Launching single computer {0} in {1} mode. Global limit={2}. Args={3}; env={4}." -f $computerName,$singleModeCombo.SelectedItem,[int]$globalConcurrencyLimitBox.Value,$optionArguments.Count,$optionEnvironment.Count)
        $run = Start-ToolkitSingleComputer -ComputerName $computerName -Mode ([string]$singleModeCombo.SelectedItem) -GlobalConcurrencyLimit ([int]$globalConcurrencyLimitBox.Value) -GlobalConcurrencyLeaseTimeoutMinutes ([int]$optionGlobalConcurrencyLeaseTimeoutBox.Value) -AdditionalArguments $optionArguments -EnvironmentVariables $optionEnvironment
        $script:LastSingleRunPath = $run.RunPath
        $singleRunPathBox.Text = $run.RunPath
        Set-FlatButtonStyle -Button $openSingleRunFolderButton -BackColor $colorPanelSoft -ForeColor $colorInk -BorderColor $colorTextBoxBorder
        Add-Status ("Single computer run folder: {0}" -f $run.RunPath)
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Single computer launch failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$openSingleRunFolderButton.Add_Click({
    try {
        $pathToOpen = $script:LastSingleRunPath
        if ([string]::IsNullOrWhiteSpace($pathToOpen)) {
            $pathToOpen = $script:SingleComputerRunRoot
        }
        if ([string]::IsNullOrWhiteSpace($pathToOpen)) {
            throw "No single computer run folder configured."
        }
        if (-not (Test-Path -LiteralPath $pathToOpen -PathType Container)) {
            New-Item -ItemType Directory -Path $pathToOpen -Force -ErrorAction Stop | Out-Null
        }
        Open-FolderPath -Path $pathToOpen
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Open run folder failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$createLotButton.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($newLotNameBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show($form, "Enter a LOT name.", "LOT name required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
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
            "Fill Computers.txt",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            Open-TextFile -Path $result.ComputersPath
        }
    }
    catch {
        Add-Status ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "LOT creation failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$openNewComputersButton.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($script:CreatedComputersPath)) {
            throw "Create a LOT first."
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
    Write-Host ("Smart Intune Hybrid Join Toolkit LOT Launcher GUI validation completed. Lots={0}" -f $folders.Count)
    return
}

Add-Status "Ready. Select an existing LOT or create a new empty LOT."
[void]$form.ShowDialog()
