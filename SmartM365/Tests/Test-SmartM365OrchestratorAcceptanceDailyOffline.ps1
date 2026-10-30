#Requires -Version 7.0
<#
.SYNOPSIS
Checks the daily server acceptance run by every orchestrator resident, offline.
.DESCRIPTION
Extracts the resident functions and runs them against a fake acceptance script in a temporary
folder: due time with the per-server offset, catch-up after startup when the published result is
older than 24 hours, one run at a time, exit code kept across a recycle (<log>.exitcode), time
limit, disabled mode, cluster keys and validation. Real hidden pwsh child processes are started;
no tenant, collector, mail or publication.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
$script:Checks = 0
function Assert-Daily {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$orchestratorText = [IO.File]::ReadAllText($orchestratorPath)
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$tokens, [ref]$errors)
foreach ($name in @('Get-PwshPath', 'Stop-ProcessTree', 'Get-OrchestratorAcceptanceDailyOffsetMinutes', 'Get-OrchestratorAcceptanceLastResultUtc', 'Test-OrchestratorAcceptanceDailyDue', 'Get-OrchestratorAcceptanceDailyRecord', 'Update-OrchestratorAcceptanceDaily')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    Assert-Daily ($null -ne $definition) "$name exists."
    . ([scriptblock]::Create($definition.Extent.Text))
}
$script:Log = [Collections.Generic.List[object]]::new()
function Write-OrchestratorLog { param([string]$Message, [string]$Level = 'INFO') $script:Log.Add([pscustomobject]@{ Message = $Message; Level = $Level }) }
$script:Saves = 0
function Save-OrchestratorState { $script:Saves++ }

# Due time: daily after the local time plus the offset, once; catch-up only at the first check after startup.
$day = [datetime]::new(2026, 10, 30)
$utc = { param([datetime]$Local) $Local.ToUniversalTime() }
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(16).AddMinutes(59) -DailyTime '17:00') -eq '') 'Not due before the daily time.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(17) -DailyTime '17:00') -eq 'Daily') 'Due at the daily time.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(17).AddMinutes(5) -DailyTime '17:00' -OffsetMinutes 10) -eq '') 'Offset delays the run.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(17).AddMinutes(10) -DailyTime '17:00' -OffsetMinutes 10) -eq 'Daily') 'Due at the daily time plus the offset.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(20) -DailyTime '17:00' -LastStartUtc (& $utc $day.AddHours(17).AddMinutes(31))) -eq '') 'Only one run per day.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(20) -DailyTime '17:00' -LastStartUtc (& $utc $day.AddDays(-1).AddHours(17).AddMinutes(31))) -eq 'Daily') 'A day missed while stopped runs later the same day.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -LastResultUtc (& $utc $day.AddHours(-28)) -StartupCheck) -eq 'CatchUp') 'Catch-up at startup when the result is older than 24 h.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -StartupCheck) -eq 'CatchUp') 'Catch-up at startup without any published result.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -LastResultUtc (& $utc $day.AddHours(-2)) -StartupCheck) -eq '') 'No catch-up with a recent result.'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -LastStartUtc (& $utc $day.AddHours(1)) -StartupCheck) -eq '') 'No catch-up loop after a recent start (recycle).'
Assert-Daily ((Test-OrchestratorAcceptanceDailyDue -Now $day.AddHours(2) -DailyTime '17:00' -LastResultUtc (& $utc $day.AddHours(-28))) -eq '') 'No catch-up after the first check.'

# Offset: 5 minutes per rank in the sorted list of expected servers; 0 when the server is not listed.
$local = ([string]$env:COMPUTERNAME).ToUpperInvariant()
$script:Settings = [pscustomobject]@{ ExpectedOrchestratorServers = @('ZZZZ-LAST', $local.ToLowerInvariant(), '0000-FIRST') }
$expectedIndex = [array]::IndexOf([string[]]@(@('ZZZZ-LAST', $local, '0000-FIRST') | Sort-Object -Unique), $local)
Assert-Daily ((Get-OrchestratorAcceptanceDailyOffsetMinutes) -eq $expectedIndex * 5) 'Offset follows the rank of this server.'
$script:Settings = [pscustomobject]@{ ExpectedOrchestratorServers = @('OTHER-A') }
Assert-Daily ((Get-OrchestratorAcceptanceDailyOffsetMinutes) -eq 0) 'Unlisted server has no offset.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-AcceptanceDaily-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
try {
    $null = New-Item -ItemType Directory -Path (Join-Path $root 'Shared') -Force
    $fakeScript = Join-Path $root 'Fake-Acceptance.ps1'
    function Get-OrchestratorAcceptanceScriptPath { return $fakeScript }
    $Tenant = 'offline'
    $script:Settings = [pscustomobject]@{
        AcceptanceDailyEnabled = $true; AcceptanceDailyTime = '00:00'; AcceptanceDailyTimeoutMinutes = 60
        ExpectedOrchestratorServers = @(); JobLogFolderPath = (Join-Path $root 'Logs'); SharedDataFolderPath = (Join-Path $root 'Shared')
    }

    # Published result: GeneratedAtUtc is read from the service account file of this server.
    Assert-Daily ((Get-OrchestratorAcceptanceLastResultUtc) -eq [datetime]::MinValue) 'No published result.'
    $serverFolder = Join-Path $root "Shared\$local"
    $null = New-Item -ItemType Directory -Path $serverFolder -Force
    [IO.File]::WriteAllText((Join-Path $serverFolder 'Orchestrator-Acceptance.json'), '{ "GeneratedAtUtc": "2026-10-29T03:00:00.0000000Z" }', [Text.UTF8Encoding]::new($false))
    Assert-Daily ((Get-OrchestratorAcceptanceLastResultUtc) -eq [datetime]::new(2026, 10, 29, 3, 0, 0, [DateTimeKind]::Utc)) 'Published GeneratedAtUtc is read.'

    # Disabled: nothing starts.
    $script:State = @{}
    $script:AcceptanceDailyStartupChecked = $false
    $script:Settings.AcceptanceDailyEnabled = $false
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily ($null -eq (Get-OrchestratorAcceptanceDailyRecord) -and $script:Log.Count -eq 0) 'Disabled: nothing starts.'
    $script:Settings.AcceptanceDailyEnabled = $true

    # Normal run: started hidden with -Tenant, one run at a time, exit code 1 reported as WARN.
    [IO.File]::WriteAllLines($fakeScript, [string[]]@('param([string]$Tenant)', '"fake acceptance tenant=$Tenant"', 'Start-Sleep -Seconds 2', 'exit 1'))
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    $record = Get-OrchestratorAcceptanceDailyRecord
    Assert-Daily ($null -ne $record -and [int]$record['Pid'] -gt 0 -and $script:Log[-1].Message -like 'Daily server acceptance started (daily): PID *') "Run started: $($script:Log[-1].Message)"
    $firstPid = [int]$record['Pid']
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily ([int](Get-OrchestratorAcceptanceDailyRecord)['Pid'] -eq $firstPid -and @($script:Log | Where-Object Message -like '*started*').Count -eq 1) 'No second run while the first one runs.'
    Wait-Process -Id $firstPid -Timeout 60 -ErrorAction SilentlyContinue
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    $finished = $script:Log[-1]
    Assert-Daily ($finished.Message -like 'Daily server acceptance finished: exit code 1 (at least one FAIL)*' -and $finished.Level -eq 'WARN') "Exit code 1 reported as WARN: $($finished.Message)"
    $logPath = [string]$record['LogPath']
    Assert-Daily ((Get-Content -LiteralPath $logPath -Raw) -match 'fake acceptance tenant=offline') 'The acceptance output and -Tenant are in the log.'
    Assert-Daily ((Split-Path $logPath -Leaf) -like "Acceptance-$($env:COMPUTERNAME)_*.log" -and (Test-Path -LiteralPath ($logPath + '.exitcode'))) 'Log and exit code file named for retention.'
    $record = Get-OrchestratorAcceptanceDailyRecord
    Assert-Daily (-not $record.Contains('Pid') -and $record['LastStartUtc']) 'Record cleared, last start kept.'
    $logCount = $script:Log.Count
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily ($script:Log.Count -eq $logCount) 'No second run the same day.'

    # Recycle: a record left by the previous instance is finished from the exit code file.
    $exitFile = Join-Path $root 'Logs\Acceptance-previous.log'
    [IO.File]::WriteAllText($exitFile + '.exitcode', '0')
    $script:State['AcceptanceDaily'] = @{ Pid = 999999; ProcessStartUtc = (Get-Date).ToUniversalTime().ToString('o'); LastStartUtc = (Get-Date).AddMinutes(-3).ToUniversalTime().ToString('o'); LogPath = $exitFile }
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily ($script:Log[-1].Message -like 'Daily server acceptance finished: exit code 0 (PASS or WARN)*' -and $script:Log[-1].Level -eq 'INFO') "Previous instance run reported: $($script:Log[-1].Message)"

    # Time limit: the process tree is stopped and reported.
    [IO.File]::WriteAllLines($fakeScript, [string[]]@('param([string]$Tenant)', 'Start-Sleep -Seconds 60'))
    $script:State = @{}
    $script:Settings.AcceptanceDailyTimeoutMinutes = 0
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    $slowPid = [int](Get-OrchestratorAcceptanceDailyRecord)['Pid']
    Start-Sleep -Milliseconds 500
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily (@($script:Log | Where-Object { $_.Message -like 'Daily server acceptance stopped after*' -and $_.Level -eq 'WARN' }).Count -eq 1) 'Time limit reported.'
    Start-Sleep -Seconds 1
    Assert-Daily ($null -eq (Get-Process -Id $slowPid -ErrorAction SilentlyContinue)) 'Process tree stopped at the limit.'
    $script:Settings.AcceptanceDailyTimeoutMinutes = 60

    # Missing script: warned once, not retried the same day.
    function Get-OrchestratorAcceptanceScriptPath { return (Join-Path $root 'Missing.ps1') }
    $script:State = @{}
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Update-OrchestratorAcceptanceDaily -Now (Get-Date)
    Assert-Daily (@($script:Log | Where-Object Message -like 'Daily server acceptance not started: script not found*').Count -eq 1) 'Missing script warned once.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

# Wiring: main loop outside the launch phase (Active or Inactive, no MaxConcurrency), retention, startup log.
$mainLoop = [regex]::Match($orchestratorText, '(?s)Test-OrchestratorRuntimeUpdate -Now \$now.{0,400}?Update-OrchestratorAcceptanceDaily -Now \$now.{0,500}?Update-OrchestratorDiskSpaceGuard -Now \$now.{0,400}?Invoke-LaunchPhase -Now \$now')
Assert-Daily $mainLoop.Success 'Called every tick, before and outside the launch phase.'
$update = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Update-OrchestratorAcceptanceDaily' }, $true).Extent.Text
Assert-Daily ($update -notmatch '\$script:Settings\.MaxConcurrency|Get-OrchestratorLocalServerState|Test-OrchestratorLocalServerActive|Send-Orchestrator|Send-SmartM365') 'No concurrency, server state or mail involved.'
Assert-Daily ($orchestratorText.Contains("Filter = 'Acceptance-*.log*'; Days = `$script:Settings.JobLogRetentionDays")) 'Acceptance logs follow the job log retention.'
Assert-Daily ($orchestratorText.Contains('Write-OrchestratorLog -Message ("Daily server acceptance: enabled={0}')) 'Settings logged at startup.'

# Cluster keys: template, central mapping and validation.
$cluster = Read-SmartM365OrchestratorJson -Path (Join-Path $base 'Orchestrator-Cluster.json.template')
Assert-Daily ($cluster.AcceptanceDailyEnabled -eq $true -and $cluster.AcceptanceDailyTime -eq '17:00' -and (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) 'Template keys are valid.'
foreach ($bad in @('4:30', '24:00', 'noon')) {
    $cluster.AcceptanceDailyTime = $bad
    Assert-Daily (-not (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) "AcceptanceDailyTime '$bad' rejected."
}
$cluster.AcceptanceDailyTime = '23:59'; $cluster.AcceptanceDailyEnabled = 'yes'
Assert-Daily (-not (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid) 'AcceptanceDailyEnabled must be a boolean.'
Assert-Daily ($orchestratorText.Contains("AcceptanceDailyEnabled = 'Boolean'") -and $orchestratorText.Contains("AcceptanceDailyTime = 'String'")) 'Central cluster keys are applied by the resident.'

"[{0}] PASS: daily server acceptance, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBeJjoDOwSbgQSi
# 8nywidl/8m08uW58fjmrKXOWGM50/6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJjBsle6L1kz+PSDHw0RC6X/Hy/W8IpFqFAGs9sgGZcSMA0GCSqG
# SIb3DQEBAQUABIIBgAfDH4ZwZh3vJbqKNsAdOxOa6XykeI0+OQsGAiRVt8/QsCkR
# tGt3oAMmyhItJISiFgmj+OZzantzilJnTh3YiA/bz3eA3sZFSD6dQP8nNeL7LH1O
# sPdYxLmyMQsT1RRzRc4QOlToWaW6zYPgfME5JljJkkepIK1Bz6ILY8aY2AzrmvwR
# 3f1+xQ0FtQHmTCEUT8h0Au5sxHmOP6lRRYo6fbeRe9gCBSQ+vUpm3CiP/7ZSPTjH
# fSeWdO288leGhwPIFccj2c3UrGAYUtQ6WXVX0yXPn60FuB9lCeBR6uXtCOptPsMq
# iW8UNAMc83Ub2bxXAYXOG7hE5pDyDLKDwiApFDmlZB9sQ7dsZYpajDxngN71JlT+
# /Y6+qtupbowipEoUYV6b/W11by6w4PggH1ivvMN3QD8GzxlfF9jdRaDJfmTlf8Da
# QOAQlb3e63dvJd8hi9/ngMe820dJcN59GUHLr0Z6YDOknwcckKXD7vAtxPnyuA5U
# +0m39F/kyHLVhG0IvaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAxMTEx
# MjBaMC8GCSqGSIb3DQEJBDEiBCBi4nneO5hmAyE5Qzp5BvREVMkz4Ln9iTTtbwxQ
# zwzcnTANBgkqhkiG9w0BAQEFAASCAgBx3tQzwcrbssYthqsAdPtyuW6mN9l+vfAs
# H07CjoynI7x1YKv08tgefJ9ZKv8ygzXjBpaCyc8UbpPNrPpQPyGmTjPwueOnoPDf
# HYCkMXkRsw8SDP7oxGJx7r9vtTCXjQNKEugewpomOee2t6lhCbR37z7Rgwl/etN7
# T9D4o1luxUlwMkBUTGv8BTWyvi3gxfLyN3uXBLDD2xRkkYmTc15Gs7FJQwYk4r57
# T8sS2k80VmjpU/Po+4d5f2rFd7i2ZdvIXYs0zyIDcaQB6xHVoNbPzPBayUq8whcc
# naq4643qEDr65BpLPkAOJw2zw7Owv2MgeMamsCAq0qakSXrW73Cwk/1PG0XR9RFh
# eclDlXY0FkaC4yg6XHAMOMkRyJTxmJ+gBNzzsSs8oCWKzdv6y25Dr/b5WWzIhGYc
# 5ybHIQpyNanizC8YZAquSc/TTTaTo2qPnxmoDa2LH4d1XZLeYsyHc8ccrdT5XloR
# lywSwqKOGZvv/Q+Xh7PMMTXUUjRe+Xr/m02HfKKpI0CstyUm5ianCmL3j9YiUTYN
# KlZ3stCbkhkzI4MlPwHBMJeIhwWuzkt5UCGitTGeBms5xLy9EenHTqboB6rt1VNM
# GABA4oypvNU90FiXEinu/xc1lGys57QYeC3YaztDRfwyK4I06js4oiz4QrAfWhNs
# /kY17hbQKg==
# SIG # End signature block
