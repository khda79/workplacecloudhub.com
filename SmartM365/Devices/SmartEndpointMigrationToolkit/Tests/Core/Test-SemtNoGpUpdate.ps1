<#
.SYNOPSIS
    Offline tests: the endpoint scripts never run gpupdate.
.VERSION
    1.0.0
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$hj = Get-Content -LiteralPath (Join-Path $root 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1') -Raw
$w11 = Get-Content -LiteralPath (Join-Path $root 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1') -Raw
$prereq = Get-Content -LiteralPath (Join-Path $root 'Campaigns\Prerequisites\Scripts\SmartM365-Invoke-EndpointMigration-Prerequisites.ps1') -Raw
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}
function Import-Function([string]$Source, [string]$Name) {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ('Parse error in endpoint: ' + $errors[0].Message) }
    $f = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name },$true)
    if (-not $f) { throw "Missing function $Name" }
    return $f.Extent.Text
}

foreach ($endpoint in @(@{ Name = 'Hybrid Join'; Text = $hj }, @{ Name = 'Windows 11'; Text = $w11 }, @{ Name = 'Prerequisites'; Text = $prereq })) {
    Assert-Test ('{0} never starts gpupdate.exe' -f $endpoint.Name) ($endpoint.Text -notmatch "(?i)['""]gpupdate(\.exe)?['""]")
}
Assert-Test 'Hybrid Join has no initial or final gpupdate block' ($hj -notmatch 'Refreshing computer Group Policy' -and $hj -notmatch 'finalGpUpdateReasons' -and $hj -notmatch 'GpUpdateTimeoutSeconds')
Assert-Test 'Windows 11 has no gpupdate helper' ($w11 -notmatch 'function Invoke-SemtGpUpdateWithTimeout')

# Windows 11 policy repair removes the legacy values and returns without gpupdate.
. ([scriptblock]::Create((Import-Function $w11 'Repair-WindowsUpdatePolicyBlockers')))
function Write-SmartLog { param($Message,$Level) }
function Test-Path { param([string]$LiteralPath) $true }
$script:removed = New-Object System.Collections.Generic.List[string]
function Remove-ItemProperty { param([string]$LiteralPath,[string]$Name,$ErrorAction) $script:removed.Add($Name) }
$result = Repair-WindowsUpdatePolicyBlockers
Assert-Test 'Windows 11 policy repair removes the legacy Windows Update values' ($script:removed -contains 'WUServer' -and $script:removed -contains 'UseWUServer' -and $script:removed -contains 'TargetReleaseVersion')
Assert-Test 'Windows 11 policy repair reports no gpupdate timeout and no output file' (-not $result.TimedOut -and [string]::IsNullOrEmpty($result.OutputFile))

Write-Host ("Summary: Passed={0}; Failed={1}" -f $passed,$failed)
if ($failed -gt 0) { exit 1 }
exit 0
