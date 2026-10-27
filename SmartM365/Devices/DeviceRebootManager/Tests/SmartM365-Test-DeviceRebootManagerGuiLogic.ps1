#Requires -Version 5.1

<#
.SYNOPSIS
    Offline regression tests for the Smart Device Reboot Manager GUI logic.

.DESCRIPTION
    Loads the GUI functions from their AST without running the application, then
    checks the review fixes: boot-time failure handling, restart command and
    failure handling, persisted forced countdown, configuration booleans, template
    key merge, relaunch argument quoting and the locale-independent session check.
    No restart, scheduled task, CIM query or window is executed for real.

.VERSION
    1.0.0
#>

[CmdletBinding()]
param(
    [string]$GuiPath = ''
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($GuiPath)) {
    $GuiPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'SmartM365-DeviceRebootManager-GUI.ps1'
}

$script:Checks = 0
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($GuiPath, [ref]$tokens, [ref]$parseErrors)
Assert-True ($parseErrors.Count -eq 0) 'GUI script parses'

$testRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('SmartM365-DrmGuiLogic-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null

try {
    # Top-level function definitions only; the application body is never executed.
    foreach ($definition in $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
        . ([scriptblock]::Create($definition.Extent.Text))
    }

    $Demo = $false
    $EnableDebugLogging = $false
    $script:Ui = @{}
    $script:AppName = 'Smart Device Reboot Manager'
    $script:UserDataRoot = $testRoot
    $script:LogPath = Join-Path $testRoot 'test.log'
    $script:StatePath = Join-Path $testRoot 'state.json'
    $script:PreferencesPath = Join-Path $testRoot 'preferences.json'
    $script:CliBoundParameters = @{}
    $script:Effective = [ordered]@{ ReminderCooldownHours = 20; RequiredRestartAfterDays = 7 }
    $script:Strings = [pscustomobject]@{ RestartFailed = 'Synthetic restart failure' }
    $script:ResumedCountdownMinimumSeconds = 300
    $script:PreviewOnly = $false
    $script:PostponeCount = 0
    $script:DeadlineUtc = $null
    $script:PersistForcedCountdown = $false
    $script:DaysSinceReboot = 9
    $script:LastBoot = (Get-Date).AddDays(-9)

    # Configuration booleans: JSON strings must not turn "false" into $true.
    Assert-True ((ConvertTo-ConfigBoolean -Value 'false' -Name 'X') -eq $false) 'string false is false'
    Assert-True ((ConvertTo-ConfigBoolean -Value 'True' -Name 'X') -eq $true) 'string True is true'
    Assert-True ((ConvertTo-ConfigBoolean -Value 0 -Name 'X') -eq $false) 'zero is false'
    Assert-True ((ConvertTo-ConfigBoolean -Value $true -Name 'X') -eq $true) 'boolean true is kept'
    $invalidRejected = $false
    try { ConvertTo-ConfigBoolean -Value 'maybe' -Name 'X' | Out-Null } catch { $invalidRejected = $true }
    Assert-True $invalidRejected 'invalid boolean text is rejected'
    $config = [pscustomobject]@{ PreviewOnly = 'false'; NeverForceRestart = 'true' }
    Assert-True ((Get-ConfigValue -Config $config -Name 'PreviewOnly' -CurrentValue $true -Type bool) -eq $false) 'config PreviewOnly "false" stays false'
    Assert-True ((Get-ConfigValue -Config $config -Name 'NeverForceRestart' -CurrentValue $false -Type bool) -eq $true) 'config NeverForceRestart "true" is true'

    # Template keys are added without overwriting local values, including nested objects.
    $target = [pscustomobject]@{ A = 1; Nested = [pscustomobject]@{ X = 'local' } }
    $template = [pscustomobject]@{ A = 2; B = 3; Nested = [pscustomobject]@{ X = 'template'; Y = 4 } }
    $added = @(Add-MissingTemplateProperties -Target $target -Template $template)
    Assert-True (($added -join ',') -eq 'B,Nested.Y') 'missing template keys are reported'
    Assert-True ($target.A -eq 1 -and $target.Nested.X -eq 'local' -and $target.B -eq 3 -and $target.Nested.Y -eq 4) 'local values are preserved'
    $configPath = Join-Path $testRoot 'config.json'
    $templatePath = Join-Path $testRoot 'config.json.template'
    '{"A":1}' | Set-Content -LiteralPath $configPath -Encoding ASCII
    '{"A":2,"B":3}' | Set-Content -LiteralPath $templatePath -Encoding ASCII
    $merged = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    Add-MissingConfigKeys -Config $merged -ConfigPath $configPath -TemplatePath $templatePath
    $saved = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    Assert-True ($saved.A -eq 1 -and $saved.B -eq 3) 'merged configuration is saved when writable'

    # Uptime: a failed CIM query never produces a forced restart.
    function Get-CimInstance { throw 'Synthetic CIM failure' }
    Assert-True ($null -eq (Get-LastBootTime)) 'boot time failure returns null'
    Remove-Item -Path Function:\Get-CimInstance

    # Cooldown: a failed restart does not silence the next reminder; timestamps parse invariantly.
    $recent = (Get-Date).ToUniversalTime().AddHours(-1).ToString('o')
    Assert-True (Test-ReminderCooldown -State ([pscustomobject]@{ LastReminderUtc = $recent; LastReason = 'RemindLater' })) 'recent reminder is in cooldown'
    Assert-True (-not (Test-ReminderCooldown -State ([pscustomobject]@{ LastReminderUtc = $recent; LastReason = 'RestartFailed' }))) 'failed restart ignores cooldown'
    Assert-True (Test-ReminderCooldown -State ([pscustomobject]@{ LastReminderUtc = (Get-Date).ToUniversalTime().AddHours(-1); LastReason = 'Close' })) 'DateTime state values (PS7 JSON) are supported'

    # Forced countdown persistence per boot.
    $bootUtc = (Get-Date).ToUniversalTime().AddDays(-9)
    $future = (Get-Date).ToUniversalTime().AddMinutes(40)
    $state = [pscustomobject]@{ LastBootUtc = $bootUtc.ToString('o'); ForcedDeadlineUtc = $future.ToString('o'); PostponeCount = 2 }
    Assert-True (Restore-ForcedCountdownState -State $state -LastBoot $bootUtc.ToLocalTime().AddSeconds(30)) 'same boot resumes'
    Assert-True ([Math]::Abs(($script:DeadlineUtc - $future).TotalSeconds) -lt 1 -and $script:PostponeCount -eq 2) 'deadline and postpone count are restored'
    $state.ForcedDeadlineUtc = (Get-Date).ToUniversalTime().AddHours(-3).ToString('o')
    Assert-True (Restore-ForcedCountdownState -State $state -LastBoot $bootUtc.ToLocalTime()) 'expired deadline resumes'
    $graceSeconds = ($script:DeadlineUtc - (Get-Date).ToUniversalTime()).TotalSeconds
    Assert-True ($graceSeconds -gt 290 -and $graceSeconds -le 300) 'expired deadline gets the minimum grace only'
    Assert-True (-not (Restore-ForcedCountdownState -State $state -LastBoot $bootUtc.ToLocalTime().AddDays(1))) 'another boot starts a new countdown'
    Assert-True (-not (Restore-ForcedCountdownState -State ([pscustomobject]@{ LastBootUtc = $bootUtc.ToString('o') }) -LastBoot $bootUtc)) 'state without deadline starts a new countdown'
    $script:DeadlineUtc = (Get-Date).ToUniversalTime().AddSeconds(125)
    Update-SecondsRemainingFromDeadline
    Assert-True ($script:SecondsRemaining -ge 124 -and $script:SecondsRemaining -le 125) 'remaining seconds follow the wall clock'
    $script:PersistForcedCountdown = $true
    $script:PostponeCount = 1
    Save-ReminderState -Reason 'Postpone10Minutes' -DaysSinceReboot 9 -LastBoot $bootUtc.ToLocalTime()
    $persisted = Get-Content -LiteralPath $script:StatePath -Raw | ConvertFrom-Json
    Assert-True ($persisted.ForcedDeadlineUtc -and [int]$persisted.PostponeCount -eq 1) 'forced deadline and postpone count are persisted'
    $script:PersistForcedCountdown = $false

    # Restart: forced, waited, exit code checked; failure is reported and recorded.
    $script:StartProcessCalls = New-Object System.Collections.Generic.List[string]
    $script:SyntheticExitCode = 0
    $script:RestartFailedShown = 0
    function Start-Process { param($FilePath, $ArgumentList, $WindowStyle, [switch]$Wait, [switch]$PassThru)
        $script:StartProcessCalls.Add(('{0}|{1}|Wait={2}|PassThru={3}' -f $FilePath, $ArgumentList, [bool]$Wait, [bool]$PassThru))
        [pscustomobject]@{ ExitCode = $script:SyntheticExitCode }
    }
    function Restart-Computer { param([switch]$Force) if ($script:SyntheticExitCode -ne 0) { throw 'Synthetic Restart-Computer failure' } $script:StartProcessCalls.Add('Restart-Computer') }
    function Show-RestartFailedMessage { $script:RestartFailedShown++ }

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    Assert-True ((Invoke-SafeRestart -Reason 'Test' -Confirm:$false) -eq $true) 'successful restart returns true'
    if ($isAdmin) {
        Assert-True ($script:StartProcessCalls[0] -eq 'Restart-Computer') 'administrator restart uses Restart-Computer -Force'
    }
    else {
        Assert-True ($script:StartProcessCalls[0] -match 'shutdown\.exe\|/r /f /t 0\|Wait=True\|PassThru=True') 'standard user restart is forced and waited'
    }
    $script:SyntheticExitCode = 1190
    Assert-True ((Invoke-SafeRestart -Reason 'Test' -Confirm:$false) -eq $false) 'failed restart returns false'
    Assert-True ($script:RestartFailedShown -eq 1) 'failed restart is shown to the user'
    $failedState = Get-Content -LiteralPath $script:StatePath -Raw | ConvertFrom-Json
    Assert-True ($failedState.LastReason -eq 'RestartFailed') 'failed restart is recorded'
    Remove-Item -Path Function:\Start-Process, Function:\Restart-Computer

    # Relaunch quoting must survive powershell.exe -File parsing.
    $stubPath = Join-Path $testRoot 'Show-Args.ps1'
    @'
param([string]$A, [string]$B, [string]$C, [string]$D)
[Console]::Out.Write((@($A, $B, $C, $D) | ConvertTo-Json -Compress))
'@ | Set-Content -LiteralPath $stubPath -Encoding ASCII
    $values = @('C:\Path With Space\', 'Title "quoted" text', 'a\\"b\', 'plain')
    $commandLine = '-NoProfile -ExecutionPolicy Bypass -File {0} -A {1} -B {2} -C {3} -D {4}' -f (ConvertTo-ProcessArgument -Value $stubPath), (ConvertTo-ProcessArgument -Value $values[0]), (ConvertTo-ProcessArgument -Value $values[1]), (ConvertTo-ProcessArgument -Value $values[2]), (ConvertTo-ProcessArgument -Value $values[3])
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $startInfo.Arguments = $commandLine
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $process = [Diagnostics.Process]::Start($startInfo)
    $output = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()
    # PS 5.1 ConvertFrom-Json emits a JSON array as one object; unroll it.
    $bound = @(($output | ConvertFrom-Json) | ForEach-Object { $_ })
    for ($index = 0; $index -lt $values.Count; $index++) {
        Assert-True ($bound[$index] -ceq $values[$index]) ("relaunch argument {0} round-trips: [{1}] -> [{2}]" -f $index, $values[$index], $bound[$index])
    }

    # Session visibility uses WTS state, not localized 'query session' text.
    $script:Effective.TestRequiredRestart = $false
    $script:Effective.TestRecommendedRestart = $false
    Assert-True (Test-UserSessionVisible) 'current interactive session is visible'
    Assert-True ([bool]('SmartM365.DeviceRebootManager.WtsSession' -as [type])) 'WTS interop type is available'
    Assert-True (-not ((Get-Content -LiteralPath $script:LogPath -Raw) -match 'Session visibility check failed')) 'WTS session query succeeded'

    Write-Output ("PASS: {0} Device Reboot Manager GUI logic checks; PowerShell {1}; no restart, task, CIM or window operation." -f $script:Checks, $PSVersionTable.PSVersion)
}
finally {
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC780ea3rmoFOTJ
# SX5p4Kl8n6kKxTfpBTYS2RY6X3911qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKdosTcpk/890LpndR7YW8hjy/8YvBGIQaP+F/gCEzwdMA0GCSqG
# SIb3DQEBAQUABIIBgCGNMntLIJc4XZe3JNL1Y5A8atdddQC1VOnitdG/J3CoyTQH
# EMMca1KIYOdaO26Yx/0az8lTcXCoOkdDg0Etpimq0cnIOo52uM6pCKVHIs+3T4kT
# UAIeTkssQD3JEYlKcJfMayQBFOENuWBAKKMa2q1R7LEtwfmHJRRfg5MxaB0IImF7
# YmxnpirNuJpGDAWc53Ff309xR0jSOFZCi9Lu9eARB0XYM1j3CEZQNWvfBlygE53V
# g4GfGsFw7IOorKvIj1t9unWDvbfv5DBAqTn4QZsNZYndF3l9Edprv7HQMc6MvP+5
# pWE2UxzBtkhOeiRMHFkQoMBON0j9f9Oa+wOT2l5oKrL1NKV2Xywo9JMphA8uc6LQ
# DZnwI+a0QS31ufumRHHWuHulpncf8YgsvpAdbX4eDlizsK5Mfq4cqad2F8dEBjhC
# FQKuRlD7bye2DKcdzS8vrHSnQ8VfS5r7UiiTtsus7EF7g7mfXf37C0PEfNO33q4u
# dNU+aMMST8cLWXxC5KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMjQy
# MjVaMC8GCSqGSIb3DQEJBDEiBCCI9XN32AP3ABsDBUffmvpTCoOya1Np37hRQ3pM
# LzS4bTANBgkqhkiG9w0BAQEFAASCAgCnPp+veaaewk0z/nlEG3gyXm+KgZlP9TBf
# gtCoeUfQCaKf+ibr3C2ftEtYm09+3QSV0gyGGhmU/GBjsisYthgH8CzhOFMYyhEn
# wg8S2XYJ+QFVP36vU9bvzm6rcyoE49Fi1gTxBZk9zHSuX9KW8+V6h/n15JnqRIXe
# SYoegF7e5W61cikttc6wKTPz1IKU+xNyk9kLm8U7uMBWYPrE7RvFMMD36l7EkLVq
# JVT0mibFl3Zxr1k+TeidMhvXylsK0aQbHhIkK0ljm2CfbfUTgiXwxJM5au+2pYCp
# pG6x6Tpf6NfdtlXpnJOSjZzMV3wgSX6paFi0ihzEIb6ZQxQV2A23AFWGs8MFWKCk
# kgNb+OTPtRAnCxGz7KwPvs9CrYv0FJxx/KXXzl10WGGyM8HVAv/e0HknOfoUU0oY
# 0RCgd5WqIqlLJBlkWIBUrcqBLSGli+L9QxgWA1X4YaP8h3EoeLLbrqITebkqPQk/
# 1y2bYa4eaEDNRU3LXLs9aSM6qpsSRYQaXGPSN58Li7EMaZ2yulNwRmvUqYbDBsSE
# Hm9+eDgicDIJHNoyZl9I6jzcbSk58RLXMg8ficKvFenD9lqkNPwq/ulq+LauI10F
# x2vZUssoFiuwreM0MuNm+vgzo3RWGreHBIF3F3PxWPQoH5xiPpsxLfadijZU1Zys
# WVU1vy5eig==
# SIG # End signature block
