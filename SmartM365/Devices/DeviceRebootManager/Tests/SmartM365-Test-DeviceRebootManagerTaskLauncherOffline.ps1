#Requires -Version 5.1
<#
.SYNOPSIS
Tests task-only launcher repair with mocked scheduler operations.
.DESCRIPTION
Run elevated so the production administrator guard and protected backup ACL are
tested. No real task is registered, changed, started or stopped. Only allocated
temporary XML backups are written.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this offline test elevated; all scheduler operations are mocked.'
}
$script:Checks = 0
function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}
function Assert-Fails([scriptblock]$Code, [string]$Pattern) {
    $failed = $false
    try { & $Code | Out-Null } catch {
        $failed = $true
        Assert-True ($_.Exception.Message -like $Pattern) ('expected refusal: ' + $Pattern)
    }
    Assert-True $failed 'operation refused'
}
$repair = Join-Path (Split-Path $PSScriptRoot -Parent) 'Deploy\SmartM365-DeviceRebootManager-RepairTaskLauncher.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('WchSdrmTaskLauncher-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null

function Reset-Fixture {
    $global:SdrmTestBuild = '19045'
    $global:SdrmTestUpdates = 0
    $global:SdrmTestRestores = 0
    $global:SdrmTestDamageSection = $false
    $global:SdrmTestDamageAction = $false
    $global:SdrmTestOldExecute = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
    $unicode = [string][char]0x00e9
    $global:SdrmTestOldArguments = '-STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\WCH\SmartDeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1" -ConfigPath "C:\Configs\' + $unicode + ' & test\settings.json" -PreviewOnly -NeverForceRestart'
    $global:SdrmTestTask = [pscustomobject]@{
        Actions = @([pscustomobject]@{ Execute = $global:SdrmTestOldExecute; Arguments = $global:SdrmTestOldArguments; WorkingDirectory = 'C:\Program Files\WCH\SmartDeviceRebootManager' })
        Principal = [pscustomobject]@{ LogonType = 'Group' }
    }
    $global:SdrmTestDefinition = [xml]'<?xml version="1.0" encoding="UTF-16"?><Task xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task"><RegistrationInfo><Author>DEMO\Admin</Author><URI>\SmartM365\Device Reboot Manager</URI></RegistrationInfo><Triggers><LogonTrigger><Enabled>true</Enabled></LogonTrigger><TimeTrigger><Repetition><Interval>PT4H</Interval></Repetition><Enabled>true</Enabled></TimeTrigger></Triggers><Principals><Principal id="Author"><GroupId>S-1-5-32-545</GroupId><RunLevel>LeastPrivilege</RunLevel></Principal></Principals><Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><ExecutionTimeLimit>PT0S</ExecutionTimeLimit></Settings><Actions Context="Author"><Exec><Command/><Arguments/><WorkingDirectory/></Exec></Actions></Task>'
    Sync-FixtureXml
    $global:SdrmTestOriginalXml = $global:SdrmTestDefinition.OuterXml
}
function Sync-FixtureXml {
    $global:SdrmTestDefinition.Task.Actions.Exec.Command = [string]$global:SdrmTestTask.Actions[0].Execute
    $global:SdrmTestDefinition.Task.Actions.Exec.Arguments = [string]$global:SdrmTestTask.Actions[0].Arguments
    $global:SdrmTestDefinition.Task.Actions.Exec.WorkingDirectory = [string]$global:SdrmTestTask.Actions[0].WorkingDirectory
}
function Get-ScheduledTask { param($TaskPath, $TaskName, $ErrorAction) $global:SdrmTestTask }
function Get-CimInstance { param($ClassName, $ErrorAction) [pscustomobject]@{ BuildNumber = $global:SdrmTestBuild } }
function Export-ScheduledTask { param($TaskPath, $TaskName, $ErrorAction) $global:SdrmTestDefinition.OuterXml }
function New-ScheduledTaskAction {
    param($Execute, $Argument, $WorkingDirectory)
    [pscustomobject]@{ Execute = $Execute; Arguments = $Argument; WorkingDirectory = [string]$WorkingDirectory }
}
function Set-ScheduledTask {
    param($TaskPath, $TaskName, $Action, $ErrorAction)
    $global:SdrmTestUpdates++
    $global:SdrmTestTask.Actions = @($Action)
    Sync-FixtureXml
    if ($global:SdrmTestDamageSection) { $global:SdrmTestDefinition.Task.Settings.ExecutionTimeLimit = 'PT1M' }
    if ($global:SdrmTestDamageAction) { $global:SdrmTestTask.Actions[0].Arguments += ' unexpected' }
}
function Register-ScheduledTask {
    param($TaskPath, $TaskName, $Xml, [switch]$Force, $ErrorAction)
    $global:SdrmTestRestores++
    $global:SdrmTestDefinition = [xml]$Xml
    $global:SdrmTestTask.Actions = @([pscustomobject]@{
        Execute = [string]$global:SdrmTestDefinition.Task.Actions.Exec.Command
        Arguments = [string]$global:SdrmTestDefinition.Task.Actions.Exec.Arguments
        WorkingDirectory = [string]$global:SdrmTestDefinition.Task.Actions.Exec.WorkingDirectory
    })
}
function Start-ScheduledTask { throw 'A task must never be started by this repair.' }
function Stop-ScheduledTask { throw 'A task must never be stopped by this repair.' }

try {
    Reset-Fixture
    $previewRoot = Join-Path $testRoot 'preview'
    $preview = & $repair -BackupRoot $previewRoot
    Assert-True ($preview.Mode -eq 'Preview' -and -not $preview.Changed) 'default is a read-only preview'
    Assert-True ($global:SdrmTestUpdates -eq 0 -and $global:SdrmTestRestores -eq 0) 'preview does not touch tasks'
    Assert-True (-not (Test-Path $previewRoot)) 'preview does not create backups'
    Assert-True ($preview.NewArguments -ceq ('--headless "' + $global:SdrmTestOldExecute + '" ' + $global:SdrmTestOldArguments)) 'original quoted Unicode arguments preserved verbatim'
    $whatIf = & $repair -Apply -WhatIf -BackupRoot $previewRoot
    Assert-True (-not $whatIf.Changed -and $global:SdrmTestUpdates -eq 0 -and -not (Test-Path $previewRoot)) 'WhatIf never writes'

    $applied = & $repair -Apply -Confirm:$false -BackupRoot (Join-Path $testRoot 'apply')
    Assert-True ($applied.Changed -and $global:SdrmTestUpdates -eq 1 -and $global:SdrmTestRestores -eq 0) 'only the existing action is updated'
    Assert-True ($global:SdrmTestTask.Actions[0].Execute -ieq (Join-Path $env:SystemRoot 'System32\conhost.exe')) 'headless host selected'
    Assert-True ([IO.File]::ReadAllText($applied.BackupPath) -ceq $global:SdrmTestOriginalXml) 'original task XML preserved exactly'
    Assert-True ((Get-FileHash $applied.BackupPath).Hash -ceq $applied.BackupSHA256) 'backup checksum recorded'
    Assert-True ($applied.PreservedSections.Count -eq 4 -and -not $applied.TaskStarted) 'schedule, identity, settings and registration verified'
    $backupAcl = Get-Acl (Split-Path $applied.BackupPath -Parent)
    Assert-True ($backupAcl.AreAccessRulesProtected -and $backupAcl.Access.Count -eq 2) 'backup is restricted to administrators and SYSTEM'
    $again = & $repair -Apply -Confirm:$false -BackupRoot (Join-Path $testRoot 'again')
    Assert-True ($again.AlreadyHeadless -and -not $again.Changed -and $global:SdrmTestUpdates -eq 1) 'already repaired task is an idempotent no-op'
    Assert-True (-not (Test-Path (Join-Path $testRoot 'again'))) 'no extra backup for a no-op'

    Reset-Fixture; $global:SdrmTestBuild = '17763'
    Assert-Fails { & $repair -Apply -Confirm:$false -BackupRoot $testRoot } '*build 19041*'
    Assert-True ($global:SdrmTestUpdates -eq 0) 'unsupported OS is not silently changed'
    Reset-Fixture; $global:SdrmTestTask.Actions[0].Arguments = '-File "C:\OtherApp.ps1"'
    Assert-Fails { & $repair -Apply -Confirm:$false } '*does not directly identify*'
    Reset-Fixture; $global:SdrmTestTask.Actions[0].Execute = 'cmd.exe'
    Assert-Fails { & $repair -Apply -Confirm:$false } '*direct PowerShell*'
    Reset-Fixture; $global:SdrmTestTask.Actions += $global:SdrmTestTask.Actions[0]
    Assert-Fails { & $repair -Apply -Confirm:$false } '*exactly one*'
    Reset-Fixture; $global:SdrmTestTask.Principal.LogonType = 'Password'
    Assert-Fails { & $repair -Apply -Confirm:$false } '*interactive user/group*'
    Reset-Fixture
    $node = $global:SdrmTestDefinition.CreateElement('RegistrationTrigger', $global:SdrmTestDefinition.DocumentElement.NamespaceURI)
    [void]$global:SdrmTestDefinition.Task.Triggers.AppendChild($node)
    Assert-Fails { & $repair -Apply -Confirm:$false } '*registration trigger*'
    Assert-True ($global:SdrmTestUpdates -eq 0) 'registration trigger refused before a write'

    Reset-Fixture; $global:SdrmTestDamageSection = $true
    Assert-Fails { & $repair -Apply -Confirm:$false -BackupRoot (Join-Path $testRoot 'rollback-section') } '*original task was restored*'
    Assert-True ($global:SdrmTestRestores -eq 1 -and $global:SdrmTestDefinition.OuterXml -ceq $global:SdrmTestOriginalXml) 'unexpected settings change rolls back'
    Reset-Fixture; $global:SdrmTestDamageAction = $true
    Assert-Fails { & $repair -Apply -Confirm:$false -BackupRoot (Join-Path $testRoot 'rollback-action') } '*original task was restored*'
    Assert-True ($global:SdrmTestRestores -eq 1 -and $global:SdrmTestTask.Actions[0].Arguments -ceq $global:SdrmTestOldArguments) 'unexpected action change rolls back'
    "PASS: $script:Checks task-launcher repair checks; PowerShell $($PSVersionTable.PSVersion); scheduler fully mocked."
}
finally {
    $absolute = [IO.Path]::GetFullPath($testRoot)
    $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $absolute.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        [IO.Path]::GetFileName($absolute) -notlike 'WchSdrmTaskLauncher-*') { throw 'Unsafe test cleanup target.' }
    if (Test-Path -LiteralPath $absolute) { Remove-Item -LiteralPath $absolute -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqgR801mNchEE+
# l0BehlwvtVM2N9ETuCPFWZEnKbZLgaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDUBvSz/LQ1xwqy8LTtQPiv6Fc6sumDrV2j/PAj9NsCiMA0GCSqG
# SIb3DQEBAQUABIIBgF3pbKlX+cdyaXBQqjZWBtd3oHmJCIIB/llR4hXZgnD1Sn6b
# dmUDXSme/pf8NOinOPgITJ3HAlcXReZ5Mlbum4z8O2JWuEI4I6JTxshQGDdFFZZJ
# OfkGx/gLo8HwAk/pLe6x6flS+VNnEq6V0uQT6M3VqZpeWWoh4/uN4D3s6dzWsbnu
# 73E3PMKRhAjWvsqzW8KwAGU+OzbaebFEF2xI0060RLcPXotFdV35t1xXgPDjli6s
# ZKNZK5sTBFYM4ya7xfiLFWUJmJ6l/Sa0uC+4w5fKdqhwti9Zt5OYvkTmUB35hV+X
# +094oVVGdCE1ZyfHq5Ydm53eG38m8uO9i3wUOfEHrzflbsWJOXvg/OM6OkHbMPNB
# W3wLV/2snfYvXu6B+zkBUP+c3reL4BOaikq5b3wSTv0sPWtnyFJufPEl/zHeg6iG
# ANgi93vH9As5SZDxgbZjisc/efdwnJeiMGAj019yh4dNyKjIFGJgtJNniZUiKsh1
# jbTH7x+1lNNmrmJoLKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTgxNjIy
# NTlaMC8GCSqGSIb3DQEJBDEiBCAQGWBvASdoY2EFwHzoevKsyFsDetrmeSFTIs+2
# BcL3szANBgkqhkiG9w0BAQEFAASCAgCAJIBHQNDpOjjeqYhdQ2C9w7WVH6NH501S
# nDyf/klvmw61qrk0xD/ulh7K1gotxLNxsyiyzxNDBL6Azlkjp2Fnx05Ksmu47lZZ
# Z25S83c5+fJrJOTUpPyflRQy2ethB/v5IN0CImj0kWtMxufNlMd8kFx0x4MJoyCb
# n1VNvK0eNGMBeyEmoFRS0kKpH3f95O5xAUm5dFb2a3RydDeATLKrd5uhqXiebCxN
# 0UL0Bo3sxLKt10lvwi6XoTJTKYIsCeOQVPssGQxicqa5+abNDqebLn2vH5MNUc5m
# y332ty9j2Yb/+DyuFcBzK56QhdtRXb+hdsMbLQdUHLnScW1uanxivG8hdCfnVuUx
# /eWKBGQUAqMiCnYq40SIAWWgH7ZR7OahzgNSboFG70Q//ch4nqiZmJMwS0nwlo5r
# 7FTY3rYimsfc36uH4m6ziAivfFAvN7Mmamyh8CUIUxzIyMieTfdTcQkCD8cK07tK
# Shtein6BqrH21Y9si93V5N3gH7DnTVxwKrS0NZrmutwbvia12vnmT2bvJsCVquQF
# FSg+s20BRmeEyAm3zKxsMY2Pvq6z1qwLAvdyfdlzA8tY+rUnY1l0QP4ccUAy9bPO
# Ef/siHSFMgDxfcRUC1huvIjlu/qkngzyB4FZ1fcFKYvzXI2AjyGOUyV/q2isrgLO
# I/t6bW7Zhw==
# SIG # End signature block
