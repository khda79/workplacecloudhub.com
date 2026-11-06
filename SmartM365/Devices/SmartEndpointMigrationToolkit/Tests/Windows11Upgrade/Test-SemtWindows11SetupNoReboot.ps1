<#
.SYNOPSIS
    Offline tests that Windows Setup never restarts the computer by itself.
.DESCRIPTION
    Extracts the real Invoke-SetupUpgrade and Resolve-SetupUpgradeSuccessOutcome functions of the
    Windows11Upgrade endpoint script. Runs Invoke-SetupUpgrade with a harmless short process
    (whoami.exe) in place of setup.exe and checks that the Windows Setup command line carries
    /noreboot, so that the restart that completes the upgrade stays decided by the toolkit (no
    interactive user connected, or uptime threshold). Then checks the outcome of a successful
    setup for each restart decision.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$w11Path = Join-Path $root 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($w11Path, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ('Windows 11 upgrade script does not parse: {0}' -f $parseErrors[0].Message) }
foreach ($name in 'Invoke-SetupUpgrade','Resolve-SetupUpgradeSuccessOutcome','Get-SetupExitCodeInfo','Convert-SignedExitCodeToHex','Format-SetupExitCodeInfo') {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Function not found in the Windows 11 upgrade script: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}

function Write-SmartLog { param($Message,$Level = 'INFO') [void]$script:logs.Add(('{0}|{1}' -f $Level,$Message)) }
function New-SmartDirectory { param($Path) }
function Format-SetupProcessSnapshot { param($Process) return 'snapshot' }

$script:logs = New-Object System.Collections.ArrayList
$script:SetupLogDir = Join-Path ([System.IO.Path]::GetTempPath()) 'SemtSetupNoReboot'
$SetupDynamicUpdate = 'Disable'
$SetupProcessHeartbeatSeconds = 300
$SetupProcessTimeoutMinutes = 5

# Windows Setup command line.
$fakeSetup = Join-Path $env:SystemRoot 'System32\whoami.exe'
$threw = $false
$exit = $null
try { $exit = Invoke-SetupUpgrade -SetupExePath $fakeSetup } catch { $threw = $true }
Assert-Test 'setup process runs and returns its exit code' ((-not $threw) -and $null -ne $exit)
$commandLine = @($script:logs | Where-Object { $_ -like '*Creating setup process*' }) | Select-Object -First 1
Assert-Test 'setup command line is logged' (-not [string]::IsNullOrWhiteSpace([string]$commandLine))
$argumentsText = if ($commandLine -match 'Arguments=(.*)$') { $Matches[1] } else { '' }
$arguments = @($argumentsText -split '\s+' | Where-Object { $_ })
Assert-Test 'setup command line carries /noreboot' ($arguments -contains '/noreboot')
Assert-Test 'setup command line keeps /auto upgrade /quiet /eula accept' (($argumentsText -match '^/auto upgrade /quiet ') -and ($argumentsText -match '/eula accept'))
Assert-Test 'setup command line keeps the dynamic update and copy logs settings' (($argumentsText -match '/dynamicupdate disable') -and ($argumentsText -match '/copylogs "'))

# The restart after a successful setup is decided by the toolkit.
$AllowSetupCompletionRebootWhenNoUser = $true
$script:rebootDecision = ''
function Invoke-SetupCompletionRebootWhenNoUser { return $script:rebootDecision }
$cases = @(
    @{ Decision = 'ScheduledNoUser'; Status = 'DIRECT_SETUP_UPGRADE_REBOOT_SCHEDULED_NO_USER' },
    @{ Decision = 'UserConnected'; Status = 'DIRECT_SETUP_UPGRADE_REBOOT_SKIPPED_USER_CONNECTED' },
    @{ Decision = 'UserDetectionFailed'; Status = 'DIRECT_SETUP_UPGRADE_REBOOT_SKIPPED_USER_DETECTION_FAILED' },
    @{ Decision = 'ForcedByUptime'; Status = 'DIRECT_SETUP_UPGRADE_SUCCESS_FORCE_REBOOT_UPTIME_EXCEEDED' }
)
foreach ($case in $cases) {
    $script:rebootDecision = $case.Decision
    $outcome = Resolve-SetupUpgradeSuccessOutcome -Prefix 'DIRECT_SETUP_UPGRADE' -SetupExitCode 0
    Assert-Test ("setup success with restart decision {0}: {1}" -f $case.Decision,$case.Status) ([string]$outcome.Status -eq $case.Status)
}

$AllowSetupCompletionRebootWhenNoUser = $false
$outcome = Resolve-SetupUpgradeSuccessOutcome -Prefix 'DIRECT_SETUP_UPGRADE' -SetupExitCode 0
Assert-Test 'restart not requested: setup started, monitored, no restart' ([string]$outcome.Status -eq 'DIRECT_SETUP_UPGRADE_STARTED' -and [string]$outcome.NextAction -eq 'MONITOR_SETUP_AND_REBOOT')

Write-Host ("Summary: Passed={0}; Failed={1}" -f $passed,$failed)
if ($failed -gt 0) { exit 1 }
exit 0
