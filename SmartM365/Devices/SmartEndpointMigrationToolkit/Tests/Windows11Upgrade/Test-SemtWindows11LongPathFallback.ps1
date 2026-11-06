<#
.SYNOPSIS
    Offline tests of the long literal path fallback used to hash the Windows 11 setup media.
.DESCRIPTION
    Some Windows PowerShell hosts refuse \\?\ paths ("Illegal characters in path"), so every file of a
    valid setup media failed the manifest check. Extracts the real functions of the Windows11Upgrade
    endpoint script and of the Intune package media integrity helper, and checks that:
    - a refused \\?\ open is retried once through the ordinary path and the hash is unchanged;
    - the fallback is logged once per run by the endpoint script;
    - a missing file, an access error or a path of 260 characters or more is never retried;
    - the startup host summary reports the PowerShell host and whether \\?\ paths are accepted.
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
$integrityPath = Join-Path $root 'Campaigns\Windows11Upgrade\IntuneWin32\Source\SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}

function Import-ScriptFunction([string]$Path, [string[]]$Names) {
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) { throw ('Script does not parse: {0}: {1}' -f $Path,$parseErrors[0].Message) }
    foreach ($name in $Names) {
        $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        if ($null -eq $definition) { throw "Function not found in ${Path}: $name" }
        . ([scriptblock]::Create($definition.Extent.Text))
        Set-Item -Path ("Function:\script:{0}" -f $name) -Value (Get-Item -LiteralPath ("Function:\{0}" -f $name)).ScriptBlock
    }
}

Import-ScriptFunction -Path $w11Path -Names @('ConvertTo-LongLiteralPath','Open-SmartFileReadStream','Get-SmartFileSha256','Get-SmartPowerShellHostSummary')
Import-ScriptFunction -Path $integrityPath -Names @('ConvertTo-SmartM365LongLiteralPath','Open-SmartM365FileReadStream','Get-SmartM365SetupMediaFileHash')

$script:logs = New-Object System.Collections.ArrayList
function Write-SmartLog { param($Message,$Level = 'INFO') [void]$script:logs.Add(('{0}|{1}' -f $Level,$Message)) }

# Host that refuses \\?\ paths exactly like the field failure: a real .NET ArgumentException raised by OpenRead.
$script:openCalls = New-Object System.Collections.ArrayList
$refusingOpen = {
    param($LiteralPath)
    [void]$script:openCalls.Add($LiteralPath)
    if ($LiteralPath.StartsWith('\\?\')) { return [System.IO.File]::OpenRead('refused|' + $LiteralPath) }
    return [System.IO.File]::OpenRead($LiteralPath)
}

function Read-StreamSha256($Stream) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return (($sha.ComputeHash($Stream) | ForEach-Object { $_.ToString('X2') }) -join '') }
    finally { $Stream.Dispose(); $sha.Dispose() }
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('SemtLongPathFallback-{0}' -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    $file = Join-Path $tempRoot '__chunk_data'
    [System.IO.File]::WriteAllBytes($file, [byte[]](1..200))
    $expected = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash

    # Normal host: unchanged behaviour.
    Assert-Test 'endpoint hash of a normal file is unchanged' ((Get-SmartFileSha256 -Path $file) -eq $expected)
    Assert-Test 'Intune helper hash of a normal file is unchanged' ((Get-SmartM365SetupMediaFileHash -Path $file) -eq $expected)
    Assert-Test 'normal host: no fallback logged' ($script:logs.Count -eq 0)

    # Endpoint: refused \\?\ path falls back to the ordinary path.
    $script:openCalls.Clear()
    $stream = Open-SmartFileReadStream -Path $file -OpenRead $refusingOpen
    Assert-Test 'endpoint: refused long literal path is reopened through the ordinary path' ((Read-StreamSha256 $stream) -eq $expected)
    Assert-Test 'endpoint: long literal path tried first, ordinary path second' ($script:openCalls.Count -eq 2 -and $script:openCalls[0] -like '\\?\*' -and $script:openCalls[1] -eq [System.IO.Path]::GetFullPath($file))
    Assert-Test 'endpoint: fallback logged as WARN' (@($script:logs | Where-Object { $_ -like 'WARN|Long literal path refused*' }).Count -eq 1)
    $stream = Open-SmartFileReadStream -Path $file -OpenRead $refusingOpen
    $stream.Dispose()
    Assert-Test 'endpoint: fallback logged once per run' (@($script:logs | Where-Object { $_ -like 'WARN|Long literal path refused*' }).Count -eq 1)

    # Intune helper: same fallback.
    $script:openCalls.Clear()
    $stream = Open-SmartM365FileReadStream -Path $file -OpenRead $refusingOpen
    Assert-Test 'Intune helper: refused long literal path is reopened through the ordinary path' ((Read-StreamSha256 $stream) -eq $expected)
    Assert-Test 'Intune helper: two opens (long literal, then ordinary)' ($script:openCalls.Count -eq 2)

    # Errors that are not a path refusal are never retried.
    $missing = Join-Path $tempRoot 'missing.bin'
    foreach ($case in @(
            @{ Name = 'endpoint'; Command = 'Open-SmartFileReadStream' },
            @{ Name = 'Intune helper'; Command = 'Open-SmartM365FileReadStream' })) {
        $script:openCalls.Clear()
        $threw = $false
        try { $null = & $case.Command -Path $missing -OpenRead $refusingOpen } catch { $threw = $true }
        Assert-Test ("{0}: refused path to a missing file still fails" -f $case.Name) $threw

        $script:openCalls.Clear()
        $threw = $false
        $message = ''
        $deniedOpen = { param($LiteralPath) [void]$script:openCalls.Add($LiteralPath); throw [System.UnauthorizedAccessException]::new('Access denied.') }
        try { $null = & $case.Command -Path $file -OpenRead $deniedOpen } catch { $threw = $true; $message = $_.Exception.Message }
        Assert-Test ("{0}: access denied is not retried" -f $case.Name) ($threw -and $script:openCalls.Count -eq 1 -and $message -like '*Access denied*')

        $longName = Join-Path $tempRoot (('d' * 120) + '\' + ('f' * 140) + '.bin')
        $script:openCalls.Clear()
        $threw = $false
        try { $null = & $case.Command -Path $longName -OpenRead $refusingOpen } catch { $threw = $true }
        Assert-Test ("{0}: path of 260 characters or more is not retried" -f $case.Name) ($threw -and $script:openCalls.Count -eq 1)
    }

    # Startup diagnostic.
    $summary = Get-SmartPowerShellHostSummary
    Assert-Test 'host summary names PowerShell, CLR, bitness and host' ($summary -match 'PSVersion=5\.1' -and $summary -match 'CLRVersion=4\.' -and $summary -match 'Is64BitProcess=' -and $summary -match 'Host=.+powershell')
    Assert-Test 'host summary reports long literal path support' ($summary -match 'LongLiteralPath=(Supported|Refused)$')
    $w11Text = [System.IO.File]::ReadAllText($w11Path)
    Assert-Test 'endpoint logs the PowerShell host at startup' ($w11Text -match 'Startup PowerShell host: \{0\}" -f \(Get-SmartPowerShellHostSummary\)')
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $passed,$failed)
if ($failed -gt 0) { exit 1 }
exit 0
