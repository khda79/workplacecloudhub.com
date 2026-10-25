<#
.SYNOPSIS
    Regression test: a Single PC launch from the GUI must apply the GUI options like a LOT launch.
.DESCRIPTION
    The GUI passes its options as EHJIR_* environment values. A LOT launch goes through
    Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd, which turns them into launcher parameters.
    A Single PC launch called the PsExec launcher directly with the environment only, which the
    launcher does not read: unchecked "Allow dsregcmd /leave" still ran with the launcher default
    (-AllowDsregLeave = $true) and "Skip virtual machines", reboot, cleanup, delay, night pause
    and timeout choices were ignored.
    This test extracts Start-ToolkitSingleComputer from the GUI with stubbed helpers and checks
    the generated launcher command line. Nothing is launched.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$toolkitRoot = Split-Path -Path $PSScriptRoot -Parent
$guiPath = Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1'

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$errors)
Write-TestResult -Condition (@($errors).Count -eq 0) -Message 'GUI parses without error'
foreach ($name in @('Start-ToolkitSingleComputer','Assert-CmdLiteralValue','ConvertTo-CmdArgument','ConvertTo-CmdSetCommand')) {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($definition) { . ([scriptblock]::Create($definition.Extent.Text)) }
}
$mapping = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-SingleComputerOptionArguments' }, $true)
if ($mapping) { . ([scriptblock]::Create($mapping.Extent.Text)) }

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('EHJIR-SinglePcTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'Scripts') -Force | Out-Null
Set-Content -LiteralPath (Join-Path $tempRoot 'Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1') -Value '# stub' -Encoding ASCII
$script:CapturedCommands = @()
function New-SingleComputerRunContext { param([string]$ToolkitRoot, [string]$ComputerName) [pscustomobject]@{ Root = $tempRoot; ComputersPath = (Join-Path $tempRoot 'Computers.txt'); ComputerName = $ComputerName; Name = $ComputerName; LogRoot = $tempRoot; ReportRoot = $tempRoot; CentralLogRoot = $tempRoot; ArchiveRoot = $tempRoot } }
function Resolve-GuiPsExecPath { param([string]$ToolkitRoot) return 'C:\Tools\PsExec64.exe' }
function Resolve-GuiPowerShellPath { return 'C:\Program Files\PowerShell\7\pwsh.exe' }
function New-GuiLaunchCommandFile { param([string]$WorkingDirectory, [string[]]$Commands, [string]$NamePrefix, [string]$WindowTitle, [switch]$PauseWhenDone) $script:CapturedCommands = @($Commands); return (Join-Path $tempRoot 'launch.cmd') }
function Start-GuiLaunchCommandFile { param([string]$LaunchCommandPath, [string]$WorkingDirectory) }

function Get-LaunchLine {
    param([hashtable]$Environment, [string[]]$ExtraArguments = @(), [string]$Mode = 'Once')
    $script:CapturedCommands = @()
    $fictionalComputer = 'PC-0142'
    [void](Start-ToolkitSingleComputer -ToolkitRoot $tempRoot -ComputerName $fictionalComputer -Mode $Mode -ExtraArguments $ExtraArguments -Environment $Environment)
    return [string](@($script:CapturedCommands | Where-Object { $_ -like '*SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1*' }) | Select-Object -Last 1)
}

$guiOff = @{
    EHJIR_THROTTLE = '10'; EHJIR_GLOBAL_CONCURRENCY_LIMIT = '15'; EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = '0'
    EHJIR_DELAY_BETWEEN_CYCLES_MINUTES = '7'; EHJIR_DISABLE_NIGHT_PAUSE = '1'; EHJIR_NIGHT_PAUSE_START_HOUR = '21'; EHJIR_NIGHT_PAUSE_END_HOUR = '6'
    EHJIR_INTUNE_RETRY_SLEEP_MINUTES = '4'; EHJIR_INTUNE_RETRY_MAX_RETRIES = '2'; EHJIR_STALE_CLEANUP_DELAY_SECONDS = '45'; EHJIR_REBOOT_DELAY_SECONDS = '240'; EHJIR_PSEXEC_TIMEOUT_MINUTES = '90'
    EHJIR_ALLOW_DSREG_LEAVE = '0'; EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT = '0'; EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER = '0'; EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE = '0'; EHJIR_SKIP_VIRTUAL_MACHINES = '1'
    EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY = '1'; EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY = '0'; EHJIR_TECHNICIAN_RUN_GUARD_HOURS = '12'; EHJIR_CENTRAL_LOG_COLLECTION_MODE = 'Standard'
}
try {
    $line = Get-LaunchLine -Environment $guiOff
    # The generated CMD quotes this token; PowerShell 7 -File binds "-AllowDsregLeave:$false" as false either way.
    Write-TestResult -Condition ($line -match '(^| )"?-AllowDsregLeave:\$false"?( |$)') -Message 'unchecked "Allow dsregcmd /leave" passes -AllowDsregLeave:$false' -Detail $line
    Write-TestResult -Condition ($line -notmatch '-AllowRemoveStaleIntuneEnrollment|-AllowRebootWhenNoInteractiveUser|-AllowRebootAfterDsregLeave') -Message 'unchecked cleanup and reboot authorizations are not passed'
    Write-TestResult -Condition ($line -match '(^| )-SkipVirtualMachines( |$)') -Message 'checked "Skip virtual machines" passes -SkipVirtualMachines'
    Write-TestResult -Condition ($line -match '(^| )-DisableNightPause( |$)' -and $line -match '-NightPauseStartHour 21' -and $line -match '-NightPauseEndHour 6') -Message 'night pause choice and hours are passed'
    Write-TestResult -Condition ($line -match '-DelayBetweenCyclesMinutes 7' -and $line -match '-IntuneRetrySleepMinutes 4' -and $line -match '-IntuneRetryMaxRetries 2' -and $line -match '-StaleCleanupDelaySeconds 45' -and $line -match '-RebootDelaySeconds 240' -and $line -match '-PsExecTimeoutMinutes 90') -Message 'delay, retry, cleanup, reboot delay and PsExec timeout values are passed'
    Write-TestResult -Condition ($line -match '-ThrottleLimit 1( |$)' -and $line -notmatch '-ThrottleLimit 10') -Message 'Single PC keeps one worker'

    $guiOn = $guiOff.Clone()
    $guiOn.EHJIR_ALLOW_DSREG_LEAVE = '1'; $guiOn.EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT = '1'; $guiOn.EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER = '1'; $guiOn.EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE = '1'; $guiOn.EHJIR_SKIP_VIRTUAL_MACHINES = '0'; $guiOn.EHJIR_DISABLE_NIGHT_PAUSE = '0'
    $line = Get-LaunchLine -Environment $guiOn
    Write-TestResult -Condition ($line -match '(^| )-AllowDsregLeave( |$)' -and $line -match '-AllowRemoveStaleIntuneEnrollment' -and $line -match '-AllowRebootWhenNoInteractiveUser' -and $line -match '-AllowRebootAfterDsregLeave') -Message 'checked authorizations are passed'
    Write-TestResult -Condition ($line -notmatch '-SkipVirtualMachines|-DisableNightPause') -Message 'unchecked "Skip virtual machines" and enabled night pause are not passed'

    $line = Get-LaunchLine -Environment $guiOff -ExtraArguments @('-DryRun','-AuditOnly') -Mode 'OnceIgnoreRunGuard'
    Write-TestResult -Condition ($line -match '-RunOnce' -and $line -match '-IgnoreRunGuard' -and $line -match '-DryRun' -and $line -match '-AuditOnly') -Message 'mode and extra GUI arguments are still passed'
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0
