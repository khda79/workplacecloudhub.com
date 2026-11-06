<#
.SYNOPSIS
    Offline tests of the robocopy exit code read by the Windows 11 setup media copy.
.DESCRIPTION
    Extracts the real Invoke-RobocopySetupMediaCopy function of the Windows11Upgrade endpoint
    script and runs it with a short real process (cmd.exe /c exit N) instead of robocopy.
    Checks that the real exit code is returned under Windows PowerShell 5.1, and that an
    exit code that cannot be read is treated as a failed copy (never as exit code 0).
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
foreach ($name in 'Invoke-RobocopySetupMediaCopy','Get-RobocopyLogErrorSummary','Format-SmartProcessArgument') {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Function not found in the Windows 11 upgrade script: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}

function Write-SmartLog { param($Message,$Level = 'INFO') [void]$script:logs.Add(('{0}|{1}' -f $Level,$Message)) }
function Update-SetupCacheLock { param($Lock) }

$SetupMediaCopyTimeoutMinutes = 5
$cmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
$logPath = Join-Path ([System.IO.Path]::GetTempPath()) ('SemtRobocopyExitCode-{0}.log' -f [guid]::NewGuid().ToString('N'))

try {
    foreach ($code in 0,1,8,16) {
        $script:logs = New-Object System.Collections.ArrayList
        $result = $null
        $threw = $false
        try { $result = Invoke-RobocopySetupMediaCopy -RobocopyPath $cmd -Arguments @('/c',"exit $code") -LogPath $logPath -CacheLock $null }
        catch { $threw = $true }
        Assert-Test "real exit code $code is returned" ((-not $threw) -and $result -is [int] -and $result -eq $code)
        Assert-Test "exit code $code is logged" (@($script:logs | Where-Object { $_ -like "*process exited*ExitCode=$code;*" }).Count -eq 1)
    }

    # A process object whose exit code cannot be read must fail the copy.
    Set-Content -LiteralPath $logPath -Value @('2026/11/06 11:49:08 ERREUR 53 (0x00000035) Copie du fichier \\server\share\sources\install.wim','ERREUR : NOMBRE DE TENTATIVES ATTEINT.') -Encoding UTF8
    function Start-Process {
        param($FilePath,$ArgumentList,[switch]$PassThru,[switch]$NoNewWindow,$ErrorAction)
        $fake = [pscustomobject]@{ Id = 4242; HasExited = $true; ExitCode = $null; Handle = [IntPtr]::Zero }
        $fake | Add-Member -MemberType ScriptMethod -Name Refresh -Value { }
        $fake | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { }
        return $fake
    }
    $script:logs = New-Object System.Collections.ArrayList
    $result = $null
    $message = ''
    try { $result = Invoke-RobocopySetupMediaCopy -RobocopyPath 'robocopy.exe' -Arguments @('a','b') -LogPath $logPath -CacheLock $null }
    catch { $message = $_.Exception.Message }
    Remove-Item Function:\Start-Process
    Assert-Test 'unreadable exit code is not returned as success' ($null -eq $result)
    Assert-Test 'unreadable exit code fails the copy' ($message -like '*exit code could not be read; the copy is treated as failed*')
    Assert-Test 'failure names the robocopy log' ($message -like "*Log=$logPath*")
    Assert-Test 'failure carries the robocopy error summary' ($message -like '*ErrorCodes=53*')
    Assert-Test 'failure is logged as ERROR' (@($script:logs | Where-Object { $_ -like 'ERROR|*could not be read*' }).Count -eq 1)
}
finally {
    Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $passed,$failed)
if ($failed -gt 0) { exit 1 }
exit 0
