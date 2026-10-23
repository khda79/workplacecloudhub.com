<#
.SYNOPSIS
    Tests device reboot guards, native failures, LOT argument transport and reminder task lifetime offline.
.VERSION
    1.0.0
#>
#requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:Checks = 0
$root = Split-Path -Parent $PSScriptRoot
$endpoint = Join-Path $root 'Devices\Windows11UpgradeToolkit\Scripts\SmartM365-Invoke-Windows11UpgradeRepair.ps1'
$lot = Join-Path $root 'Devices\IntuneHybridJoinToolkit\Scripts\SmartM365-IntuneHybridJoinRepair-New-AutomaticLot.ps1'
$taskScript = Join-Path $root 'Devices\DeviceRebootManager\Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DevicesSafety-' + [guid]::NewGuid().ToString('N'))

function Assert-Equal($Actual, $Expected, [string]$Message) {
    if ([string]$Actual -cne [string]$Expected) { throw "ASSERT FAILED: $Message. Expected=$Expected; Actual=$Actual" }
    $script:Checks++
}
function Import-TestFunction([string]$Path, [string]$Name) {
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$parseErrors)
    if ($parseErrors) { throw "Source failed to parse: $Path" }
    $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if (-not $node) { throw "Missing function: $Name" }
    Set-Item -Path "Function:\script:$Name" -Value $node.Body.GetScriptBlock()
}

# Only extracted functions run; endpoint entry points and all operational commands stay blocked.
function Get-CimInstance {
    [CmdletBinding()]
    param($ClassName, $Filter)
    if ($script:SessionMode -eq 'QueryFailure') { Write-Error 'Synthetic session query failure'; return }
    if ($script:SessionMode -eq 'NoSessions') { return }
    [pscustomobject]@{ LogonId = 'synthetic'; LogonType = 2 }
}
function Get-CimAssociatedInstance {
    [CmdletBinding()]
    param($InputObject, $Association)
    switch ($script:SessionMode) {
        'AssociationFailure' { Write-Error 'Synthetic association failure'; return }
        'EmptyAssociation' { return }
        'ServiceOnly' { [pscustomobject]@{ Domain = 'NT AUTHORITY'; Name = 'SYSTEM' }; return }
        default { [pscustomobject]@{ Domain = 'CONTOSO'; Name = 'FictionalUser' } }
    }
}
function shutdown.exe {
    param([Parameter(ValueFromRemainingArguments = $true)][object[]]$Arguments)
    $script:ShutdownCalls++
    if ($script:ShutdownThrows) { throw 'Synthetic process launch failure' }
    $global:LASTEXITCODE = $script:ShutdownExitCode
}
function Write-SmartLog { param($Message, $Level) }
function Send-UserRebootNotification { param($Context) }
function Publish-AutomaticLotProgress { param($Stage, $Detail) }
function Get-Command {
    param([string]$Name, $ErrorAction)
    if ($Name -eq 'pwsh.exe' -and $script:ForcePS51Child) { return }
    Microsoft.PowerShell.Core\Get-Command -Name $Name -ErrorAction Stop
}

function Invoke-TestRebootBranch([string]$Branch) {
    switch ($Branch) {
        'Pending' { Invoke-ControlledRebootWhenNoUser -Reason 'Synthetic pending reboot' }
        'Completion' { Invoke-SetupCompletionRebootWhenNoUser }
        'ProfileRepair' { Invoke-SetupProfileRepairRebootWhenNoUser -Reason 'Synthetic profile repair' }
    }
}

try {
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
    foreach ($name in @('Get-InteractiveUserSessionSummary', 'Test-RequiredRebootForceByUptime',
        'Invoke-RequiredRebootForcedByUptime', 'Invoke-ControlledRebootWhenNoUser',
        'Invoke-SetupCompletionRebootWhenNoUser', 'Invoke-SetupProfileRepairRebootWhenNoUser',
        'Resolve-PendingRebootOutcome', 'Resolve-SetupUpgradeSuccessOutcome')) {
        Import-TestFunction $endpoint $name
    }
    $AllowReboot = $true
    $AllowSetupCompletionRebootWhenNoUser = $true
    $AuditOnly = $false
    $ScheduleRetryAfterReboot = $false
    $ForceRequiredRebootWhenUptimeOverDays = 0
    $RebootDelaySeconds = 60
    $script:DeviceUptimeSummary = [pscustomobject]@{ DetectionSucceeded = $true; ThresholdExceeded = $true; UptimeDays = 10 }
    $script:ShutdownExitCode = 0
    $script:ShutdownThrows = $false
    foreach ($mode in @('QueryFailure', 'AssociationFailure', 'EmptyAssociation')) {
        $script:SessionMode = $mode
        Assert-Equal (Get-InteractiveUserSessionSummary).DetectionSucceeded $false "$mode remains unknown"
        foreach ($branch in @('Pending', 'Completion', 'ProfileRepair')) {
            $script:ShutdownCalls = 0
            Assert-Equal (Invoke-TestRebootBranch $branch) 'UserDetectionFailed' "$branch blocks $mode"
            Assert-Equal $script:ShutdownCalls 0 "$branch never reboots on $mode"
        }
    }
    $script:SessionMode = 'UserConnected'
    Assert-Equal (Get-InteractiveUserSessionSummary).InteractiveUserCount 1 'connected account is counted'
    foreach ($branch in @('Pending', 'Completion', 'ProfileRepair')) {
        $script:ShutdownCalls = 0
        Assert-Equal (Invoke-TestRebootBranch $branch) 'UserConnected' "$branch preserves connected user"
        Assert-Equal $script:ShutdownCalls 0 "$branch skips shutdown for connected user"
    }
    $script:SessionMode = 'ServiceOnly'
    Assert-Equal (Get-InteractiveUserSessionSummary).InteractiveUserCount 0 'system account is excluded'
    Assert-Equal (Get-InteractiveUserSessionSummary).DetectionSucceeded $true 'resolved system account is valid evidence'
    $script:SessionMode = 'NoSessions'
    foreach ($branch in @('Pending', 'Completion', 'ProfileRepair')) {
        foreach ($exitCode in @(0, 5)) {
            $script:ShutdownExitCode = $exitCode
            $script:ShutdownCalls = 0
            $expected = if ($exitCode -eq 0) { 'ScheduledNoUser' } else { 'ScheduleFailed' }
            Assert-Equal (Invoke-TestRebootBranch $branch) $expected "$branch respects shutdown exit $exitCode"
            Assert-Equal $script:ShutdownCalls 1 "$branch invokes only the shutdown mock"
        }
        $script:ShutdownThrows = $true
        Assert-Equal (Invoke-TestRebootBranch $branch) 'ScheduleFailed' "$branch reports launch exceptions"
        $script:ShutdownThrows = $false
    }
    $script:ShutdownExitCode = 5
    Assert-Equal (Resolve-PendingRebootOutcome -PendingRebootInfo ([pscustomobject]@{Source='Synthetic'}) -RebootReason 'Synthetic').ExitCode 1 'pending reboot outcome fails on native error'
    Assert-Equal (Resolve-SetupUpgradeSuccessOutcome -Prefix 'SETUP' -SetupExitCode 3010).ExitCode 1 'setup outcome fails on native error'
    $ForceRequiredRebootWhenUptimeOverDays = 7
    foreach ($exitCode in @(0, 5)) {
        $script:ShutdownExitCode = $exitCode
        $expected = if ($exitCode -eq 0) { 'ForcedByUptime' } else { 'ScheduleFailed' }
        Assert-Equal (Invoke-RequiredRebootForcedByUptime -Reason 'Synthetic uptime reboot' -Context PendingReboot) $expected "uptime reboot respects exit $exitCode"
    }

    Import-TestFunction $lot 'Invoke-AutomaticLotWrapperRefresh'
    $unicodeName = 'Folder With Spaces ' + [char]0x00e9 + [char]0x65e5
    $wrapperRoot = Join-Path $testRoot $unicodeName
    New-Item -ItemType Directory -Path $wrapperRoot -Force | Out-Null
    $wrapperScript = Join-Path $wrapperRoot 'Harmless Wrapper.ps1'
    Set-Content -LiteralPath $wrapperScript -Encoding UTF8 -Value @'
param([string]$RootPath)
[Console]::Out.WriteLine('ROOT=' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($RootPath)))
[Console]::Out.WriteLine(('O' * 70000))
[Console]::Error.WriteLine(('E' * 70000))
exit 7
'@
    foreach ($forcePS51 in @($false, $true)) {
        $script:ForcePS51Child = $forcePS51
        foreach ($suffix in @('', '\')) {
            $requestedRoot = $wrapperRoot + $suffix
            $result = Invoke-AutomaticLotWrapperRefresh -ScriptPath $wrapperScript -EffectiveToolkitRoot $requestedRoot
            Assert-Equal $result.ExitCode 7 "native failure retained with PS51 fallback=$forcePS51 and trailing slash='$suffix'"
            $expectedRoot = 'ROOT=' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($requestedRoot))
            Assert-Equal (@($result.Output | Where-Object { $_ -ceq $expectedRoot }).Count) 1 "Unicode path arrives unchanged with PS51 fallback=$forcePS51 and trailing slash='$suffix'"
            Assert-Equal (@($result.Output | Where-Object { $_.Length -eq 70000 }).Count) 2 'both redirected streams are drained without deadlock'
        }
    }

    # Execute the deployment entry point with every ScheduledTasks command replaced.
    function New-ScheduledTaskAction { param($Execute, $Argument) [pscustomobject]@{Execute=$Execute;Argument=$Argument} }
    function New-ScheduledTaskTrigger { param([switch]$AtLogOn,[switch]$Once,$At,$RepetitionInterval,$RepetitionDuration) 'SyntheticTrigger' }
    function New-ScheduledTaskPrincipal { param($GroupId,$RunLevel) 'SyntheticPrincipal' }
    function New-ScheduledTaskSettingsSet {
        param([switch]$AllowStartIfOnBatteries,[switch]$DontStopIfGoingOnBatteries,[switch]$StartWhenAvailable,$MultipleInstances,$ExecutionTimeLimit)
        $global:DevicesSafetyTaskCapture.TimeLimit = $ExecutionTimeLimit
        $global:DevicesSafetyTaskCapture.MultipleInstances = $MultipleInstances
        'SyntheticSettings'
    }
    function Register-ScheduledTask { param($TaskPath,$TaskName,$Action,$Trigger,$Principal,$Settings,$Description,[switch]$Force) $global:DevicesSafetyTaskCapture.Count++ }
    function Start-ScheduledTask { throw 'Unexpected scheduled task launch in offline test' }
    $global:DevicesSafetyTaskCapture = [pscustomobject]@{Count=0;TimeLimit=$null;MultipleInstances=''}
    Set-Content -LiteralPath (Join-Path $testRoot 'SmartM365-DeviceRebootManager-GUI.ps1') -Value '# Synthetic app fixture' -Encoding UTF8
    & $taskScript -InstallPath $testRoot | Out-Null
    Assert-Equal $global:DevicesSafetyTaskCapture.Count 1 'only mocked registration executes'
    Assert-Equal $global:DevicesSafetyTaskCapture.TimeLimit ([TimeSpan]::Zero) 'scheduler cannot cut off a valid postpone countdown'
    Assert-Equal $global:DevicesSafetyTaskCapture.MultipleInstances 'IgnoreNew' 'repeating triggers cannot open concurrent reminder instances'
    Write-Output ("PASS: {0} device safety checks; PowerShell {1}; no live CIM, reboot or scheduled task operation." -f $script:Checks,$PSVersionTable.PSVersion)
}
finally {
    Remove-Variable -Name DevicesSafetyTaskCapture -Scope Global -ErrorAction SilentlyContinue
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolvedRoot.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        [IO.Path]::GetFileName($resolvedRoot) -notlike 'SmartM365-DevicesSafety-*') { throw 'Unsafe test cleanup path' }
    if (Test-Path -LiteralPath $resolvedRoot) { Remove-Item -LiteralPath $resolvedRoot -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC/XB3NHxoScAPJ
# iAZmlN0supQD5kRonY9KWQe+Tv+QS6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIF5ZgpYUyWvsEyNxLQV2k1SUIpXa1mUcoskKQn4a1eltMA0GCSqG
# SIb3DQEBAQUABIIBgHBlTSuDFoUWocja10qMbfPuUTBLmYCZ6xNA0g0PLKNcgA7q
# UtCmqK6Hq0yUWnICtGQy2g/f3qV+Uq9nDK2VHS8dIAg7GzpkVnafNnmWDZXYjQqw
# JXD6hk5wQyaImohe2us777pbRc9Ysn++X/JCiW4bCIJ/F41d7OVp3AT6k0w82qQP
# 1LF5lwXOjNh02DEYXK5TwQ0GOWROk7GJADcy0Nig2xV6IykeGXA0aqPVBsx2E2SQ
# 5p0nU8fpbemdw0Wx/IhhvairpROXH/bEJqFU38PoE2zqUnr9qx3//bwTKnBaAn87
# Nrearp8Cvf407adG+JDs8FoROe42femiXDdk5Ptawbj1kvaCMvRHT0r7JwcqBhMQ
# jP+FTSTtjmTkTTnOZRBNRazVDEZh8gyMUeFX2EXbKvGXnRfei0Ir/MkADS2xkM3p
# H9PYU2GaTcZ6De+gtcs8fi/kPcwIZdLF5HpWEwmQ1ro7Su0wfcVwz0T+iBqiHK2n
# PtFtzDb5kAzNkLRW+aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzI3
# NTlaMC8GCSqGSIb3DQEJBDEiBCBIht0cfFiMjL2DHQFPaLw0JxTP3SXUwBYLjdpn
# nmxmcDANBgkqhkiG9w0BAQEFAASCAgA5+R2V6itnRAa5zeNrCdRxNDfokUddnx5Q
# bHp9tgaBZwZrs7sz5MQvQqfR9QQny5Ti46jqSXg4G0oKUqit9zfFdE4tEXzXxiPj
# gxvtq4QVbm423R+ksrKJ29g1qNn1UqosuxBGk8Kwnj/r7vPhJcaz3jjE7u/R0Ltw
# xCxk8PMirCr3IR59ydTh33Egn5hyhvFJ35Hngs5YWGMHQ+IXoHnhQdSvDLrerehm
# xQWmGyQcZQ64jC6cGxni8YxWajuxKf1vzlUDFNfwq5qXEo36XniRHNSZtE/dg4xQ
# 6YBJ3bMpd6kBvTDE1YDMziVV8lGGjiiKP1T19wOv1/qfkcDesRPaJrQEle+gXm6I
# HbzU12qVari8uhRHh1qB6QIJ5QeGH8nz+K1DrvTZmUmMmERAY+b9iLP4Ic8AXOxr
# pQPazxsvhLnUvhV/HiaB4Uqv4b3ZKmMVhy2gJrzBMS8YmuSeJuX2B4+tGhDwZYFO
# oZL/G7BGlb34YoU3cFUyWCD5cI2Vffz27SSW4SgxAO4TPEcNrKVWh9TpfiRFR5nD
# C8ZZWdTniE/3PXAgvqZDJ9q0Ulnk3c1l8tn0Lwa9ZE25KgTiUdzE00GXj7alXHyV
# j3t5GeBiazjmEuR+GXZnGTPSVhuANvK0RjU5PLNVaBjfl0ay9U82PRi9VysyeXU3
# 5ffXjSQmOg==
# SIG # End signature block
