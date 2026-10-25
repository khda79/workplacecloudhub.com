<#
.SYNOPSIS
    Regression test: DirectSetup must not re-enable reboots the operator disabled.
.DESCRIPTION
    Runs the option-default block of Scripts\Run-Windows11UpgradeRepairWithPsExec-Lot.cmd in a
    temporary CMD with chosen W11UT_* values and checks the resolved W11UT_ALLOW_REBOOT and
    W11UT_SETUP_REBOOT_WHEN_NO_USER. The GUI passes W11UT_ALLOW_REBOOT=0 when "Allow reboot" is
    unchecked; with W11UT_DIRECT_SETUP_UPGRADE=1 the wrapper used to force it back to 1, which
    passed -AllowReboot to the endpoint (pending-reboot handling and the uptime-forced reboot,
    which reboots even with a connected user). Nothing is launched: no PowerShell orchestrator,
    PsExec or endpoint runs.
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
$cmdPath = Join-Path $toolkitRoot 'Scripts\Run-Windows11UpgradeRepairWithPsExec-Lot.cmd'

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

$lines = [IO.File]::ReadAllLines($cmdPath)
$start = [Array]::FindIndex($lines, [Predicate[string]]{ param($line) $line -like 'rem Default LOT behavior*' })
$end = [Array]::FindIndex($lines, [Predicate[string]]{ param($line) $line -like 'set "W11UT_ALLOW_SETUP_SOURCE_VALIDATION=0"*' })
Write-TestResult -Condition ($start -ge 0 -and $end -gt $start) -Message 'option-default block found in the LOT wrapper'

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('W11UT-DirectRebootTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    function Invoke-DefaultBlock {
        param([hashtable]$Values)
        $probe = Join-Path $tempRoot ('probe-{0}.cmd' -f [guid]::NewGuid().ToString('N').Substring(0, 6))
        $body = New-Object System.Collections.Generic.List[string]
        $body.Add('@echo off')
        $body.Add('setlocal EnableExtensions EnableDelayedExpansion')
        foreach ($name in @('W11UT_DIRECT_SETUP_UPGRADE','W11UT_ALLOW_REBOOT','W11UT_SETUP_REBOOT_WHEN_NO_USER')) {
            if ($Values.ContainsKey($name)) { $body.Add(('set "{0}={1}"' -f $name,$Values[$name])) } else { $body.Add(('set "{0}="' -f $name)) }
        }
        for ($i = $start; $i -lt $end; $i++) { $body.Add($lines[$i]) }
        $body.Add('echo RESULT ALLOW_REBOOT=%W11UT_ALLOW_REBOOT%;REBOOT_WHEN_NO_USER=%W11UT_SETUP_REBOOT_WHEN_NO_USER%;DIRECT=%W11UT_DIRECT_SETUP_UPGRADE%')
        [IO.File]::WriteAllLines($probe, $body.ToArray(), [Text.Encoding]::ASCII)
        $output = @(& cmd.exe /d /c ('"{0}"' -f $probe) 2>&1 | ForEach-Object { [string]$_ })
        return (@($output | Where-Object { $_ -like 'RESULT *' }) -join '')
    }

    $result = Invoke-DefaultBlock -Values @{ W11UT_DIRECT_SETUP_UPGRADE = '1'; W11UT_ALLOW_REBOOT = '0'; W11UT_SETUP_REBOOT_WHEN_NO_USER = '0' }
    Write-TestResult -Condition ($result -match 'ALLOW_REBOOT=0;REBOOT_WHEN_NO_USER=0;DIRECT=1') -Message 'DirectSetup keeps reboots disabled when the operator disabled them' -Detail $result

    $result = Invoke-DefaultBlock -Values @{ W11UT_DIRECT_SETUP_UPGRADE = '1'; W11UT_ALLOW_REBOOT = '0' }
    Write-TestResult -Condition ($result -match 'ALLOW_REBOOT=0;REBOOT_WHEN_NO_USER=1;DIRECT=1') -Message 'DirectSetup keeps AllowReboot=0 and the setup-completion default' -Detail $result

    $result = Invoke-DefaultBlock -Values @{ W11UT_DIRECT_SETUP_UPGRADE = '1' }
    Write-TestResult -Condition ($result -match 'ALLOW_REBOOT=1;REBOOT_WHEN_NO_USER=1;DIRECT=1') -Message 'DirectSetup without explicit values keeps the documented defaults (reboots on)' -Detail $result

    $result = Invoke-DefaultBlock -Values @{ W11UT_DIRECT_SETUP_UPGRADE = '0'; W11UT_ALLOW_REBOOT = '0' }
    Write-TestResult -Condition ($result -match 'ALLOW_REBOOT=0;REBOOT_WHEN_NO_USER=1;DIRECT=0') -Message 'without DirectSetup an explicit AllowReboot=0 is kept' -Detail $result

    $text = [IO.File]::ReadAllText($cmdPath)
    Write-TestResult -Condition ($text -match '(?m)^if /I "%W11UT_ALLOW_REBOOT%"=="1" set "ACTION_ARGS=%ACTION_ARGS% -AllowReboot"') -Message '-AllowReboot is passed only when W11UT_ALLOW_REBOOT=1'
    Write-TestResult -Condition ($text -notmatch '(?m)^if /I "%W11UT_DIRECT_SETUP_UPGRADE%"=="1" set "W11UT_(ALLOW_REBOOT|SETUP_REBOOT_WHEN_NO_USER)=1"') -Message 'no line forces reboot options from DirectSetup'
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0
