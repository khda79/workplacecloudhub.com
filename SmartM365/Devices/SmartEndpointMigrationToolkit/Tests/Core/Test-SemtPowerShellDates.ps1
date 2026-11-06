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
    # Inventory CSV dates: ISO is invariant; other text is the culture of the exporting PC (Export-Csv).
    $inventoryChecks = & {
        $saved = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
            # Each check in parentheses: the comma binds tighter than -eq.
            @(
                ((ConvertFrom-SemtInventoryDateText -Value '04/11/2026 11:19:11') -eq [datetime]::new(2026, 11, 4, 11, 19, 11, [DateTimeKind]::Utc)),
                ((ConvertFrom-SemtInventoryDateText -Value '27/10/2026 09:21:40') -eq [datetime]::new(2026, 10, 27, 9, 21, 40, [DateTimeKind]::Utc)),
                ((ConvertFrom-SemtInventoryDateText -Value '2026-04-11T11:19:11.0000000Z') -eq [datetime]::new(2026, 4, 11, 11, 19, 11, [DateTimeKind]::Utc)),
                ((ConvertFrom-SemtInventoryDateText -Value '2026-06-08 08:53:32') -eq [datetime]::new(2026, 6, 8, 8, 53, 32, [DateTimeKind]::Utc)),
                ($null -eq (ConvertFrom-SemtInventoryDateText -Value '')),
                ($null -eq (ConvertFrom-SemtInventoryDateText -Value 'not a date'))
            )
        }
        finally { [Threading.Thread]::CurrentThread.CurrentCulture = $saved }
    }
    Test-Step 'inventory date in fr-FR culture: 04/11/2026 is 4 November, not 11 April' { $inventoryChecks[0] }
    Test-Step 'inventory date in fr-FR culture: 27/10/2026 is read' { $inventoryChecks[1] }
    Test-Step 'inventory ISO dates are read invariantly whatever the culture' { $inventoryChecks[2] -and $inventoryChecks[3] }
    Test-Step 'empty or non-date inventory text gives no date' { $inventoryChecks[4] -and $inventoryChecks[5] }
    $jsonPath = Join-Path $tempRoot 'dates.json.txt'
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
        $leasePath = Join-Path $tempRoot ('lease-' + $mode + '.json.txt')
        Write-SemtJsonFileAtomic -Path $leasePath -InputObject ([pscustomobject]@{ LauncherProcessId = $PID; LauncherProcessStartUtc = $start; Shifted = ((ConvertTo-SemtUtcDateTime -Value $start).AddHours(2).ToString('o')) })
        $lease = Read-SemtJsonFile -Path $leasePath
        Test-Step ("{0}: live process recognised from its recorded start time" -f $label) { Test-SemtProcessAlive -ProcessId $PID -ExpectedStartUtc ([string]$lease.LauncherProcessStartUtc) }
        Test-Step ("{0}: a different start time is detected as PID reuse" -f $label) { -not (Test-SemtProcessAlive -ProcessId $PID -ExpectedStartUtc ([string]$lease.Shifted)) }
        # Active and stale reservations of the shared worker gate.
        $gate = Initialize-SemtWorkerGate -Limit 4 -LeaseTimeoutMinutes 120 -HeartbeatStaleMinutes 20 -Root (Join-Path $tempRoot ('Gate-' + $mode))
        New-Item -ItemType Directory -Path $gate.LeaseRoot -Force | Out-Null
        $now = (Get-Date).ToUniversalTime()
        Write-SemtJsonFileAtomic -Path (Join-Path $gate.LeaseRoot 'lease_active.json.txt') -InputObject ([pscustomobject]@{ Campaign = 'HybridJoin'; Computer = 'PC-A'; Limit = 4; LauncherProcessId = $PID; LauncherProcessStartUtc = $start; CreatedUtc = $now.ToString('o'); HeartbeatUtc = $now.ToLocalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffzzz') })
        Write-SemtJsonFileAtomic -Path (Join-Path $gate.LeaseRoot 'lease_stale.json.txt') -InputObject ([pscustomobject]@{ Campaign = 'HybridJoin'; Computer = 'PC-B'; Limit = 4; LauncherProcessId = $PID; LauncherProcessStartUtc = $start; CreatedUtc = $now.AddMinutes(-40).ToString('o'); HeartbeatUtc = $now.AddMinutes(-30).ToLocalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffzzz') })
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
    $pending = Join-Path $targetRoot 'State\TargetPendingOperation.json.txt'
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
        Get-ChildItem -LiteralPath (Join-Path $guiCopy 'Config') -Filter '*.local.json.txt' -File | Remove-Item -Force
        $ErrorActionPreference = 'Continue'
        $output = @(& (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $guiCopy 'Scripts\SmartM365-EndpointMigration-GUI.ps1') -ValidateOnly 2>&1 | ForEach-Object { [string]$_ })
        $ErrorActionPreference = 'Stop'
        Test-Step ("GUI refused on {0} with SEMT_UNSUPPORTED_GUI_HOST and no state written (EXECUTED)" -f $engineName) { ($output -join "`n") -match 'SEMT_UNSUPPORTED_GUI_HOST' -and @(Get-ChildItem -LiteralPath (Join-Path $guiCopy 'Config') -Filter '*.local.json.txt' -File).Count -eq 0 }
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB1/z8sJiksIIDP
# yn41wZF2MI73ZFkFWjGcCtbpvLRx8aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIDSQo2hkI6ENv5C/cUUzis1M/i12F8q18H8v2zhfau/hMA0GCSqG
# SIb3DQEBAQUABIIBgHnhfD0zzGi0PLaPeHOL2I0IVh/S9jzzdsxw7EMfCM1Ce9jf
# Kbtdysooz+g+w66V39IbiR6B3bkKbZgr4b8VPwFNxUd8Ftd3grs+Em01hInKOeAL
# n6zNbhYo1hgGCAFKou0h+mfwxg5i1In29B2hlmx4cyNEnPimEXOZEpG6iTin+TX7
# cj/YksDJzJqBPuF1a4bHrFIlfOOsVj7yOB0aD4VbJEmwtqbqjqZ13uHMr6L4ND/3
# 9l3edp7GVa7cTotXFGfYbkEFRtclSApnWxqHZM3u0WXncUnmmakVabcj4DSOS35/
# xL2vrwqVEpiUxRByiYFDo/29DnpHakPbodUnhPOFs8v+p+/CdAmudvmxcE6RlipI
# +zhBo5y/T7X2f2eZMxqruugtd0yZqqc5N767Yz16JAmc/yfP59/kJz2OMifkCpBI
# OjmToj1gUXFEWaQE1QUHG9pTvcWcv5a3AB80lBfNE0bLt76U2kimtg0FQ2psUIDk
# +Eu4Vb4EtxHCaWCpNaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDYxMzIz
# MDVaMC8GCSqGSIb3DQEJBDEiBCAFwES5CMVY4XbMIyH54T11gxZbBGjDA9Cpv4BQ
# 0UrXTTANBgkqhkiG9w0BAQEFAASCAgBafOSi6EOzrU43ObkWd9p++GydaVRwgwix
# 69bz9MkxQmWYQo7gObvC7KP5CJ374e/yH3bBQP9hhirIwDfegIHc4y2ar7qZDiZM
# Hy0FkcXOkXUe09I0FkAvyvN+7HtDRVf3ihs3UMNAo21iFvPkGUUgggi2cL1KAwPZ
# AFMaT27ELzam5eLP3l6V2xP5+i+DnVS6qj1UKuppt2FKOxdKi1LFqSGo5/Sid5dm
# 0DL/ASCbT28TZtvQ2vnrUUC7qrZaX44jDzEo4bHx50ZlWuRJhHBz+DCZng3/Hi/I
# onmMeWnxCJ3yYUhsPltyOazKjjSzCCWf3ASvyjAQ5WmGfKUZmm939xFM2ZDvDMRr
# H/jDlJHWihs3qiyvmOkS6FDLRR8NxzYmY53sLB/NOsOY4lPfq4CzaRCzL7/vhAT4
# 5kdcAULw6aJwj25e+ZdC0uSUgFPND8ZQS81i1jQTuXe9a+HX4Tw4GB/+HefMJWVT
# cGj5dXSbjmNCjB+WNHJxrpsihab1D8B9UfQb6V2dVq9W/BZZz+74DhH17djzbD1h
# BUXXYWFcpFXPVDs1Yy2xZXOh/W3t3IhreCNaWBs05btnoyFHBq+AoO4qznw0IZI4
# FsZ587lek1UQpBUcS3e1RETzn74HDmbpkgvuumsHOOdVn0vonuG63+2VaRg0+THC
# 559vyYT0Cg==
# SIG # End signature block
