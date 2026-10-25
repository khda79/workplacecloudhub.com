<#
.SYNOPSIS
    Offline tests of JSON date handling and of the supported PowerShell engines.
.DESCRIPTION
    Covers UTC ('Z') and offset dates, values already deserialized as DateTime/DateTimeOffset,
    process start time comparison, active and stale worker reservations, pending-operation
    markers, the JSON path without -DateKind, and the early refusal of unsupported engines.
    Run it on Windows PowerShell 5.1 and on PowerShell 7. Each line states whether the behaviour
    was EXECUTED on the running engine or SIMULATED (for example the path without -DateKind on an
    engine that has it). Nothing is claimed for an engine that is not installed.
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
$packageRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$engineName = '{0} {1}' -f $PSVersionTable.PSEdition,$PSVersionTable.PSVersion
$hasDateKind = (Get-Command -Name ConvertFrom-Json).Parameters.ContainsKey('DateKind')

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try { Write-TestResult -Condition ([bool](& $Check)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

Write-Host ("INFO: engine={0}; ConvertFrom-Json -DateKind available={1}" -f $engineName,$hasDateKind)
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-DateTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
$previousOverride = $env:SEMT_RUNTIME_ROOT_OVERRIDE
$env:SEMT_RUNTIME_ROOT_OVERRIDE = Join-Path $tempRoot 'Runtime'
try {
    Import-Module (Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
    $module = Get-Module SmartEndpointMigration.Core
    $utcText = '2026-10-24T21:00:28.1234567Z'
    $offsetText = '2026-10-24T23:00:28.1234567+02:00'
    $expected = [datetime]::new(2026, 10, 24, 21, 0, 28, [System.DateTimeKind]::Utc).AddTicks(1234567)
    $jsonPath = Join-Path $tempRoot 'dates.json'
    Set-Content -LiteralPath $jsonPath -Value ('{ "Utc": "' + $utcText + '", "Offset": "' + $offsetText + '", "List": [ "' + $utcText + '", { "Nested": "' + $offsetText + '" } ], "Text": "not a date" }') -Encoding UTF8

    foreach ($mode in @('Default','WithoutDateKind')) {
        $label = if ($mode -eq 'Default') { if ($hasDateKind) { 'EXECUTED with -DateKind String' } else { 'EXECUTED without -DateKind (engine has none)' } } else { if ($hasDateKind) { 'SIMULATED without -DateKind on an engine that has it (real ConvertFrom-Json date conversion)' } else { 'EXECUTED without -DateKind (engine has none)' } }
        & $module { param($Enabled) $script:SemtJsonDateKindSupported = $Enabled } ($mode -eq 'Default' -and $hasDateKind)
        $data = Read-SemtJsonFile -Path $jsonPath
        Test-Step ("{0}: UTC 'Z' value read as text" -f $label) { $data.Utc -is [string] }
        Test-Step ("{0}: UTC 'Z' value keeps its exact instant" -f $label) { (ConvertTo-SemtUtcDateTime -Value $data.Utc) -eq $expected }
        Test-Step ("{0}: offset value read as text with the same instant" -f $label) { $data.Offset -is [string] -and (ConvertTo-SemtUtcDateTime -Value $data.Offset) -eq $expected }
        Test-Step ("{0}: dates nested in arrays and objects normalized" -f $label) { @($data.List)[0] -is [string] -and (ConvertTo-SemtUtcDateTime -Value @($data.List)[0]) -eq $expected -and (ConvertTo-SemtUtcDateTime -Value @($data.List)[1].Nested) -eq $expected -and $data.Text -eq 'not a date' }
        # Process start time comparison through a JSON round trip (PID reuse guard).
        $start = & $module { Get-SemtProcessStartUtc -ProcessId $PID }
        $leasePath = Join-Path $tempRoot ('lease-' + $mode + '.json')
        Write-SemtJsonFileAtomic -Path $leasePath -InputObject ([pscustomobject]@{ LauncherProcessId = $PID; LauncherProcessStartUtc = $start; Shifted = ((ConvertTo-SemtUtcDateTime -Value $start).AddHours(2).ToString('o')) })
        $lease = Read-SemtJsonFile -Path $leasePath
        Test-Step ("{0}: live process recognised from its recorded start time" -f $label) { Test-SemtProcessAlive -ProcessId $PID -ExpectedStartUtc ([string]$lease.LauncherProcessStartUtc) }
        Test-Step ("{0}: a different start time is detected as PID reuse" -f $label) { -not (Test-SemtProcessAlive -ProcessId $PID -ExpectedStartUtc ([string]$lease.Shifted)) }
        # Active and stale reservations of the shared worker gate.
        $gate = Initialize-SemtWorkerGate -Limit 4 -LeaseTimeoutMinutes 120 -HeartbeatStaleMinutes 20 -Root (Join-Path $tempRoot ('Gate-' + $mode))
        New-Item -ItemType Directory -Path $gate.LeaseRoot -Force | Out-Null
        $now = (Get-Date).ToUniversalTime()
        Write-SemtJsonFileAtomic -Path (Join-Path $gate.LeaseRoot 'lease_active.json') -InputObject ([pscustomobject]@{ Campaign = 'HybridJoin'; Computer = 'PC-A'; Limit = 4; LauncherProcessId = $PID; LauncherProcessStartUtc = $start; CreatedUtc = $now.ToString('o'); HeartbeatUtc = $now.ToLocalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffzzz') })
        Write-SemtJsonFileAtomic -Path (Join-Path $gate.LeaseRoot 'lease_stale.json') -InputObject ([pscustomobject]@{ Campaign = 'HybridJoin'; Computer = 'PC-B'; Limit = 4; LauncherProcessId = $PID; LauncherProcessStartUtc = $start; CreatedUtc = $now.AddMinutes(-40).ToString('o'); HeartbeatUtc = $now.AddMinutes(-30).ToLocalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffzzz') })
        $snapshot = & $module { param($g) Get-SemtWorkerGateSnapshotUnlocked -Gate $g -RemoveStale } $gate
        Test-Step ("{0}: active reservation (offset heartbeat) kept" -f $label) { @($snapshot.Active | Where-Object { $_.Lease.Computer -eq 'PC-A' }).Count -eq 1 }
        Test-Step ("{0}: stale reservation (heartbeat 30 min old) recovered" -f $label) { @($snapshot.Removed | Where-Object { $_.Computer -eq 'PC-B' -and $_.Reason -eq 'HeartbeatStale' }).Count -eq 1 }
    }
    & $module { param($Enabled) $script:SemtJsonDateKindSupported = $Enabled } $hasDateKind

    Test-Step 'already deserialized DateTime (UTC and local) and DateTimeOffset convert to the same instant' {
        (ConvertTo-SemtUtcDateTime -Value $expected) -eq $expected -and (ConvertTo-SemtUtcDateTime -Value $expected.ToLocalTime()) -eq $expected -and (ConvertTo-SemtUtcDateTime -Value ([datetimeoffset]::new($expected).ToOffset([timespan]::FromHours(2)))) -eq $expected
    }
    Test-Step 'DateTime values inside a deserialized object are turned back into round-trip text' {
        $object = [pscustomobject]@{ A = $expected; B = @($expected.ToLocalTime(), [pscustomobject]@{ C = [datetimeoffset]::new($expected) }) }
        $converted = ConvertTo-SemtJsonDateText -Value $object
        $converted.A -eq $utcText -and @($converted.B)[0] -eq $utcText -and @($converted.B)[1].C -eq $utcText
    }

    # Pending-operation markers of the endpoint guard region.
    $endpoint = Join-Path $packageRoot 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
    $region = ([regex]::Match([IO.File]::ReadAllText($endpoint), '(?s)#region SEMT shared target guards v1.*?#endregion SEMT shared target guards v1')).Value
    . ([scriptblock]::Create($region))
    $targetRoot = Join-Path $tempRoot 'Target'
    $hjData = Join-Path $targetRoot 'HybridJoin'
    $pending = Join-Path $targetRoot 'State\TargetPendingOperation.json'
    New-Item -ItemType Directory -Path (Split-Path $pending -Parent) -Force | Out-Null
    Set-Content -LiteralPath $pending -Value ('{ "Campaign": "Windows11Upgrade", "Reason": "test", "ExpiresUtc": "' + (Get-Date).ToUniversalTime().AddHours(5).ToString('o') + '", "LastBootUtc": "", "ClearedByReboot": true }') -Encoding UTF8
    $guard = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R1 -CorrelationId C1
    Test-Step ("pending marker with a future UTC expiry blocks the other Campaign (EXECUTED on {0})" -f $engineName) { -not $guard.Acquired -and $guard.Status -eq 'TARGET_PENDING_OPERATION' }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $hjData
    Set-Content -LiteralPath $pending -Value ('{ "Campaign": "Windows11Upgrade", "Reason": "test", "ExpiresUtc": "' + (Get-Date).AddHours(-1).ToString('yyyy-MM-ddTHH:mm:ss.fffffffzzz') + '", "LastBootUtc": "", "ClearedByReboot": true }') -Encoding UTF8
    $guard = Enter-SemtTargetOperationGuard -Campaign HybridJoin -CampaignDataRoot $hjData -RunId R1 -CorrelationId C2
    Test-Step ("pending marker with a past offset expiry no longer blocks (EXECUTED on {0})" -f $engineName) { $guard.Acquired }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $hjData
    Test-Step 'endpoint reader without -DateKind turns DateTime values back into text (EXECUTED conversion of real ConvertFrom-Json output)' {
        $raw = ('{ "ExpiresUtc": "' + $utcText + '" }') | ConvertFrom-Json
        $converted = ConvertTo-SemtTargetJsonDateText -Value $raw
        $converted.ExpiresUtc -is [string] -and ([datetime]::Parse($converted.ExpiresUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)).ToUniversalTime() -eq $expected
    }

    # Supported engines.
    Test-Step 'supported engines: Windows PowerShell 5.1 and PowerShell 7.2+ accepted (version logic)' { (Test-SemtSupportedPowerShell -Edition Desktop -Version 5.1) -and (Test-SemtSupportedPowerShell -Edition Core -Version 7.2) -and (Test-SemtSupportedPowerShell -Edition Core -Version 7.6.6) }
    Test-Step 'unsupported engines refused: Windows PowerShell 5.0, PowerShell 6.2, 7.0 and 7.1 (SIMULATED versions)' { -not (Test-SemtSupportedPowerShell -Edition Desktop -Version 5.0) -and -not (Test-SemtSupportedPowerShell -Edition Core -Version 6.2) -and -not (Test-SemtSupportedPowerShell -Edition Core -Version 7.0) -and -not (Test-SemtSupportedPowerShell -Edition Core -Version 7.1.7) }
    Test-Step 'Assert-SemtSupportedPowerShell reports SEMT_UNSUPPORTED_POWERSHELL' { try { Assert-SemtSupportedPowerShell -Edition Core -Version 7.1; $false } catch { $_.Exception.Message -like 'SEMT_UNSUPPORTED_POWERSHELL:*' } }
    Test-Step 'running engine is accepted (EXECUTED)' { Test-SemtSupportedPowerShell }
    $psm1 = [IO.File]::ReadAllText((Join-Path $packageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psm1'))
    $guardText = ([regex]::Match($psm1, '(?s)\$semtEdition = .*?\r?\n\}')).Value
    Test-Step 'module refuses PowerShell 7.1 before defining any function (SIMULATED version table)' {
        $simulated = $guardText.Replace('$PSVersionTable', '$fakeVersionTable')
        $fakeVersionTable = @{ PSEdition = 'Core'; PSVersion = [version]'7.1.0' }
        try { . ([scriptblock]::Create($simulated)); $false } catch { $_.Exception.Message -like 'SEMT_UNSUPPORTED_POWERSHELL:*' }
    }
    Test-Step 'module guard runs before the function files are loaded' { $psm1.IndexOf('SEMT_UNSUPPORTED_POWERSHELL') -lt $psm1.IndexOf('. $functionPath') }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        $refusedData = Join-Path $tempRoot 'EndpointData'
        $ErrorActionPreference = 'Continue'
        $output = @(& (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File $endpoint -DataRoot $refusedData -CorrelationId 'date-test' 2>&1 | ForEach-Object { [string]$_ })
        $exit = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        Test-Step ("Windows 11 endpoint refused on {0} with an explicit status (EXECUTED)" -f $engineName) { $exit -eq 3 -and ($output -join "`n") -match 'FINAL_STATUS=UNSUPPORTED_POWERSHELL_HOST;.*CORRELATION_ID=date-test' }
        Test-Step 'refused endpoint wrote no state (data root not created)' { -not (Test-Path -LiteralPath $refusedData) }
        $hjEndpoint = Join-Path $packageRoot 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'
        $ErrorActionPreference = 'Continue'
        $output = @(& (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File $hjEndpoint -CorrelationId 'date-test' 2>&1 | ForEach-Object { [string]$_ })
        $exit = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        Test-Step ("Hybrid Join endpoint refused on {0} with an explicit status (EXECUTED)" -f $engineName) { $exit -eq 3 -and ($output -join "`n") -match 'FINAL_STATUS=UNSUPPORTED_POWERSHELL_HOST' }
        $guiCopy = Join-Path $tempRoot 'GuiPkg'
        foreach ($item in @('Assets','Campaigns','Config','Modules','Scripts')) { Copy-Item -LiteralPath (Join-Path $packageRoot $item) -Destination (Join-Path $guiCopy $item) -Recurse -Force }
        Get-ChildItem -LiteralPath (Join-Path $guiCopy 'Config') -Filter '*.local.json' -File | Remove-Item -Force
        $ErrorActionPreference = 'Continue'
        $output = @(& (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $guiCopy 'Scripts\SmartM365-EndpointMigration-GUI.ps1') -ValidateOnly 2>&1 | ForEach-Object { [string]$_ })
        $ErrorActionPreference = 'Stop'
        Test-Step ("GUI refused on {0} with SEMT_UNSUPPORTED_GUI_HOST and no state written (EXECUTED)" -f $engineName) { ($output -join "`n") -match 'SEMT_UNSUPPORTED_GUI_HOST' -and @(Get-ChildItem -LiteralPath (Join-Path $guiCopy 'Config') -Filter '*.local.json' -File).Count -eq 0 }
    }
    else {
        Write-Host 'INFO: endpoint and GUI refusal on PowerShell 7 is executed by the PowerShell 7 run of this test.'
    }
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    $env:SEMT_RUNTIME_ROOT_OVERRIDE = $previousOverride
    try { [IO.Directory]::Delete('\\?\' + $tempRoot, $true) } catch { Write-Host ("WARN: temp folder not fully removed: {0}" -f $tempRoot) }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0
