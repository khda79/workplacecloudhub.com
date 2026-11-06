<#
.SYNOPSIS
    Offline tests of the interactive user detection used before a Windows 11 controlled reboot.
.DESCRIPTION
    Extracts the real Get-InteractiveUserSessionSummary and Get-WtsSessionUserList functions of
    the Windows11Upgrade endpoint script. Replaces the WMI and Remote Desktop Services session
    sources with fakes and checks that:
    - a logon session without an account no longer fails the detection when the session list
      can be read and shows no user;
    - the detection still fails (no reboot) when such a session exists and the session list
      cannot be read;
    - a user seen by either source counts as connected.
    Also calls the real session list once on this computer.
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
$definitions = @{}
foreach ($name in 'Get-InteractiveUserSessionSummary','Get-WtsSessionUserList') {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($null -eq $definition) { throw "Function not found in the Windows 11 upgrade script: $name" }
    $definitions[$name] = $definition.Extent.Text
}

# Real Remote Desktop Services session list on this computer.
. ([scriptblock]::Create($definitions['Get-WtsSessionUserList']))
$realEntries = $null
$realError = ''
try { $realEntries = @(Get-WtsSessionUserList) } catch { $realError = $_.Exception.Message }
Assert-Test 'real session list is read without error' ([string]::IsNullOrWhiteSpace($realError))
Assert-Test 'real session list entries are "account|session|state"' (@($realEntries | Where-Object { [string]$_ -notmatch '^[^|]+\|\d+\|\d+$' }).Count -eq 0)
Assert-Test 'real session list shows the current user' (@($realEntries | Where-Object { ([string]$_ -split '\|')[0] -like "*\$env:USERNAME" }).Count -ge 1)

# Fake sources for the detection logic.
. ([scriptblock]::Create($definitions['Get-InteractiveUserSessionSummary']))
$script:fakeSessions = @()
$script:fakeAccounts = @{}
$script:cimFails = $false
$script:wtsEntries = @()
$script:wtsFails = $false

function Get-CimInstance {
    param($ClassName,$Filter,$ErrorAction)
    if ($script:cimFails) { throw 'WMI is not available.' }
    return @($script:fakeSessions)
}
function Get-CimAssociatedInstance {
    param($InputObject,$Association,$ErrorAction)
    $key = [string]$InputObject.LogonId
    if ($script:fakeAccounts.ContainsKey($key)) { return @($script:fakeAccounts[$key]) }
    return @()
}
function Get-WtsSessionUserList {
    if ($script:wtsFails) { throw 'The RPC server is unavailable.' }
    return @($script:wtsEntries)
}
function Set-Scenario {
    param([string[]]$SessionIds = @(), [hashtable]$Accounts = @{}, [string[]]$Wts = @(), [switch]$WtsFails, [switch]$CimFails)
    $script:fakeSessions = @($SessionIds | ForEach-Object { [pscustomobject]@{ LogonId = $_ } })
    $script:fakeAccounts = $Accounts
    $script:wtsEntries = $Wts
    $script:wtsFails = [bool]$WtsFails
    $script:cimFails = [bool]$CimFails
}
function New-FakeAccount([string]$Domain, [string]$Name) { return [pscustomobject]@{ Domain = $Domain; Name = $Name } }

# Field case: a logon session without an account, nobody in the session list.
Set-Scenario -SessionIds @('101') -Accounts @{} -Wts @()
$result = Get-InteractiveUserSessionSummary
Assert-Test 'orphan logon session, empty session list: detection succeeds' ($result.DetectionSucceeded -eq $true)
Assert-Test 'orphan logon session, empty session list: no user' ([int]$result.InteractiveUserCount -eq 0)
Assert-Test 'orphan logon session is reported in the detail' ([string]$result.Detail -match 'UnresolvedLogonSessions=1;')

# Orphan session and the session list cannot be read: fail closed.
Set-Scenario -SessionIds @('101') -Accounts @{} -WtsFails
$result = Get-InteractiveUserSessionSummary
Assert-Test 'orphan logon session, session list unreadable: detection fails' ($result.DetectionSucceeded -eq $false)
Assert-Test 'failure detail names both sources' ([string]$result.Detail -match 'could not be resolved' -and [string]$result.Detail -match 'RPC server is unavailable')

# Orphan session and a disconnected user in the session list: the user counts.
Set-Scenario -SessionIds @('101') -Accounts @{} -Wts @('CONTOSO\alice|2|4')
$result = Get-InteractiveUserSessionSummary
Assert-Test 'disconnected session-list user is counted' ($result.DetectionSucceeded -eq $true -and [int]$result.InteractiveUserCount -eq 1 -and [string]$result.InteractiveUsers -eq 'CONTOSO\alice')
Assert-Test 'session state is reported by name' ([string]$result.Detail -match 'CONTOSO\\alice \(session 2, Disconnected\)')

# No orphan session: the session list is only a cross-check, its failure keeps the WMI result.
Set-Scenario -SessionIds @('201') -Accounts @{ '201' = @(New-FakeAccount 'CONTOSO' 'bob') } -WtsFails
$result = Get-InteractiveUserSessionSummary
Assert-Test 'resolved WMI user, session list unreadable: detection succeeds with the user' ($result.DetectionSucceeded -eq $true -and [int]$result.InteractiveUserCount -eq 1 -and [string]$result.InteractiveUsers -eq 'CONTOSO\bob')
Assert-Test 'unreadable session list is reported in the detail' ([string]$result.Detail -match 'RdsSessions=unavailable: The RPC server is unavailable')

# Same user from both sources, different case: listed once.
Set-Scenario -SessionIds @('201') -Accounts @{ '201' = @(New-FakeAccount 'CONTOSO' 'bob') } -Wts @('contoso\BOB|1|0')
$result = Get-InteractiveUserSessionSummary
Assert-Test 'same user from both sources is counted once' ([int]$result.InteractiveUserCount -eq 1)

# A user seen only in the session list counts even without an orphan session.
Set-Scenario -SessionIds @() -Accounts @{} -Wts @('CONTOSO\carol|1|0')
$result = Get-InteractiveUserSessionSummary
Assert-Test 'user seen only in the session list blocks the reboot' ($result.DetectionSucceeded -eq $true -and [int]$result.InteractiveUserCount -eq 1)

# System and window-manager accounts in the session list are ignored.
Set-Scenario -SessionIds @('101') -Accounts @{} -Wts @('Window Manager\DWM-1|1|0','NT AUTHORITY\SYSTEM|0|4','CONTOSO\PC01$|0|4')
$result = Get-InteractiveUserSessionSummary
Assert-Test 'system and window-manager session-list accounts are ignored' ($result.DetectionSucceeded -eq $true -and [int]$result.InteractiveUserCount -eq 0)

# WMI failure still fails the detection.
Set-Scenario -CimFails -Wts @()
$result = Get-InteractiveUserSessionSummary
Assert-Test 'WMI failure: detection fails' ($result.DetectionSucceeded -eq $false)

# Nobody anywhere.
Set-Scenario -SessionIds @() -Accounts @{} -Wts @()
$result = Get-InteractiveUserSessionSummary
Assert-Test 'no session at all: detection succeeds with no user' ($result.DetectionSucceeded -eq $true -and [int]$result.InteractiveUserCount -eq 0)

Write-Host ("Summary: Passed={0}; Failed={1}" -f $passed,$failed)
if ($failed -gt 0) { exit 1 }
exit 0
