<#
.SYNOPSIS
    Offline regressions for Intune remediation detection and package pairing.
.VERSION
    1.0.0
.NOTES
    Runs original scripts in isolated child processes with mocked system commands.
    Does not change registry keys, services, scheduled tasks, Hello keys or tenants.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$packageRoot = Join-Path $repoRoot 'SmartM365/Intune/Remediation/Packages'
$childHost = Join-Path $PSHOME $(if ($PSVersionTable.PSEdition -eq 'Desktop') { 'powershell.exe' } else { 'pwsh.exe' })
$script:checks = 0

function Assert-Review {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}

function Invoke-DetectionScenario {
    param([string]$RelativePath, [string]$Setup, [int]$ExpectedExit, [string]$ExpectedOutput, [string]$Arguments = '')
    $sourcePath = (Join-Path $packageRoot $RelativePath).Replace("'", "''")
    # Block mutators even if a regression moves them before the intended safety gate.
    $guard = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
foreach ($name in @('New-Item','Add-Content','Remove-Item','Set-ItemProperty','Remove-ItemProperty',
    'Stop-Service','Start-Service','Restart-Service','Set-Service','Start-ScheduledTask',
    'Enable-ScheduledTask','Start-Process','takeown.exe','icacls.exe')) {
    Set-Item -Path ('Function:global:' + $name) -Value { throw 'UNEXPECTED_SIDE_EFFECT' }
}
function dsregcmd { throw 'UNEXPECTED_DEVICE_REGISTRATION' }
function nltest { throw 'UNEXPECTED_DOMAIN_LOOKUP' }
'@
    $command = $guard + "`n" + $Setup + "`n" +
        "& ([scriptblock]::Create([IO.File]::ReadAllText('$sourcePath'))) $Arguments"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $lines = & $childHost -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encoded 2>&1
    $code = $LASTEXITCODE
    $outputText = $lines -join "`n"
    Assert-Review ($code -eq $ExpectedExit) "Unexpected exit $code for $RelativePath : $outputText"
    Assert-Review ($outputText -match $ExpectedOutput) "Unexpected output for $RelativePath : $outputText"
    Assert-Review ($outputText -notmatch 'UNEXPECTED_') "An unmocked operational action was attempted: $outputText"
}

$policy = 'WindowsUpdate/Policy-Blockers/SmartM365-WindowsUpdate-Policy-Blockers-Detection.ps1'
$policySetup = @'
function Test-Path { param($Path) return ($Path -like '*WindowsUpdate*') }
function Get-ItemProperty {
    [CmdletBinding()]param($Path)
    if ($Path -like '*\AU') { [pscustomobject]@{ UseWUServer = 0; NoAutoUpdate = 0; AUOptions = 4 } }
    else { [pscustomobject]@{ DoNotConnectToWindowsUpdateInternetLocations = 0; DisableWindowsUpdateAccess = 0; SetDisableUXWUAccess = 0 } }
}
'@
Invoke-DetectionScenario $policy $policySetup 0 'no blocking policy detected'
Invoke-DetectionScenario $policy 'function Test-Path { param($Path) return $false }' 0 'no WSUS or Windows Update policy blocker'
foreach ($blockingName in @('UseWUServer', 'NoAutoUpdate', 'DoNotConnectToWindowsUpdateInternetLocations', 'WUServer')) {
    $value = if ($blockingName -eq 'WUServer') { "'https://wsus.contoso.example'" } else { '1' }
    $setup = 'function Test-Path { param($Path) return ($Path -like ''*WindowsUpdate*'') }; ' +
        'function Get-ItemProperty { [CmdletBinding()]param($Path) [pscustomobject]@{' + $blockingName + '=' + $value + '} }'
    Invoke-DetectionScenario $policy $setup 1 ('remediation required.*' + $blockingName)
}

$hello = 'WindowsHello/Reset-WindowsHello/SmartM365-Reset-WindowsHello-Detection.ps1'
$helloSetup = @'
function Test-Path { param($LiteralPath) return $true }
function Get-ChildItem { [CmdletBinding()]param($LiteralPath,$Path,[switch]$Recurse) [pscustomobject]@{Name='SyntheticHelloContainer'} }
function dsregcmd { 'DomainJoined : YES'; 'AzureAdJoined : YES' }
function Get-CimInstance { [CmdletBinding()]param($ClassName) [pscustomobject]@{PartOfDomain=$true;Domain='contoso.example'} }
'@
Invoke-DetectionScenario $hello ($helloSetup + "`nfunction nltest { param(`$DomainArgument) if (`$DomainArgument -ne '/dsgetdc:contoso.example') { throw 'UNEXPECTED_DOMAIN_ARGUMENT' }; 'DC unavailable'; `$global:LASTEXITCODE=1355 }") 0 'Status=Deferred Reason=ADUnavailableNoKeyReset'
# Native exit status determines connectivity, regardless of localized success text.
Invoke-DetectionScenario $hello ($helloSetup + "`nfunction nltest { 'Synthetic localized success'; `$global:LASTEXITCODE=0 }") 0 'Status=Healthy Reason=WindowsHelloConfiguredAndADReachable'
Invoke-DetectionScenario $hello ($helloSetup + "`nfunction nltest { throw 'Synthetic domain lookup failure' }") 0 'Status=Deferred'
Invoke-DetectionScenario $hello @'
function Test-Path { param($LiteralPath) return $false }
function Get-ChildItem { [CmdletBinding()]param($LiteralPath,$Path,[switch]$Recurse) }
function dsregcmd { 'DomainJoined : NO'; 'AzureAdJoined : YES' }
'@ 0 'Status=NotApplicable'

$helloReset = 'WindowsHello/Reset-WindowsHello/SmartM365-Reset-WindowsHello-Remediation.ps1'
Invoke-DetectionScenario $helloReset '' 0 'Status=Skipped Reason=ExplicitKeyResetApprovalRequired'
# Explicit manual opt-in remains available; every mutation is simulated here.
Invoke-DetectionScenario $helloReset @'
function New-Item { param($Path,$ItemType,[switch]$Force) }
function Add-Content { param($LiteralPath,$Value,$Encoding) }
function Get-Service { [CmdletBinding()]param($Name) return $null }
function Test-Path { param($LiteralPath) return $true }
function takeown.exe { }
function icacls.exe { }
function Remove-Item { [CmdletBinding()]param($LiteralPath,[switch]$Recurse,[switch]$Force) if ($LiteralPath -ne 'C:\Windows\ServiceProfiles\LocalService\AppData\Local\Microsoft\NGC') { throw 'UNEXPECTED_DELETE_TARGET' }; 'MOCK_HELLO_RESET' }
function dsregcmd { 'DomainJoined : YES'; 'AzureAdJoined : YES' }
'@ 0 'MOCK_HELLO_RESET' '-AllowWindowsHelloReset'

$ime = 'Intune-MDM/IntuneManagementExtension-Health/SmartM365-IntuneManagementExtension-Health-Detection.ps1'
$imeSetup = @'
function Get-Service { [CmdletBinding()]param($Name) [pscustomobject]@{Status='Running'} }
function Get-Process { [CmdletBinding()]param($Name) [pscustomobject]@{Id=42} }
function Test-Path { param($Path) return ($Path -eq 'C:\Program Files (x86)\Microsoft Intune Management Extension' -or $Path -like "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs*") }
function Get-Item { param($Path) [pscustomobject]@{LastWriteTime=(Get-Date)} }
function Select-String { param($Path,$Pattern,[switch]$Quiet) return $true }
'@
Invoke-DetectionScenario $ime $imeSetup 0 'Intune Management Extension is healthy'
Invoke-DetectionScenario $ime ($imeSetup + "`nfunction Get-Item { param(`$Path) [pscustomobject]@{LastWriteTime=(Get-Date).AddDays(-2)} }") 1 'IME appears inactive'
Invoke-DetectionScenario $ime ($imeSetup + "`nfunction Test-Path { param(`$Path) return (`$Path -notlike '*AgentExecutor.log') }") 1 'AgentExecutor.log is missing'

$guiPath = Join-Path $repoRoot 'SmartM365/Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1'
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$null, [ref]$parseErrors)
Assert-Review ($parseErrors.Count -eq 0) 'GUI source does not parse.'
$definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-MatchingLocalRemediationPath' }, $true)
. ([scriptblock]::Create($definition.Extent.Text))
$deviceRecovery = Join-Path $packageRoot 'Standalone/Device-Recovery'
$match = Get-MatchingLocalRemediationPath (Join-Path $deviceRecovery 'SmartM365-Time-Zone-Detection.ps1')
Assert-Review ([IO.Path]::GetFileName($match) -eq 'SmartM365-Time-Zone-Remediation.ps1') 'Time-Zone was not paired with its remediation.'
$match = Get-MatchingLocalRemediationPath (Join-Path $deviceRecovery 'SmartM365-Recovery-Check-Detection.ps1')
Assert-Review ([IO.Path]::GetFileName($match) -eq 'SmartM365-Recovery-Check-Remediation.ps1') 'Recovery-Check pairing regressed.'
Write-Output ('PASS: {0} checks; PowerShell {1}; no live endpoint or tenant operations.' -f $script:checks, $PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAnL8OEF0LmBU2C
# oeD3JwgjVzNI6ebs5EKjXBQ+SDnB9KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMsysfNUXN7/vM7g4vln1jir+tUCet7ZTu8k9iRHZipEMA0GCSqG
# SIb3DQEBAQUABIIBgLDWMWErT/WYeOsDUCi0fOjQQcD/5GlnwvVlFbiGCPN6X5d2
# d8MAdw6bocANeSn+RQUxj3KkGE5aA+J4SAR4p6xUI9B6RfTTrTybyYH+RidqO+lp
# Mkv4ZwR1JPlRskam3VWBeL8q7RSxujyHiZYKjt3pRLJh2bKV6hIQ5Gm+BRnjuXvz
# Q4HLjzz3bBZUr33aMPhsAm1pzAHfNtY/P2cJz1vTcmwK2g+Z11of0Ggsc5Jd8bLc
# DRxLkLcy6aIJA67NfEnW9sAx395gqt52Znk0yXOWFWps/Ql4jKxSCh5937l/5JIC
# u5eK2aUQBtJLqkpWguCsjkM5nzweDycH1UFJZsHpQ4fpxS2p/Il71N28/HAKDyAO
# ASSPOOkVKBwAjT26lMSjgixlwsgeD/gdo0zN3UKgfbEJrh10LBzOlIis3+YRK6VC
# Ry06e4qPpuDvnIq4ePHDxFi2T/o7ObS9i0IHPC9nJj4PHnaQZcIpaPecSzDhOpxr
# U/mW4BRTdYFUQxvksaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzU5
# NDFaMC8GCSqGSIb3DQEJBDEiBCBdCAacj7znQjJJNIaVHT1WKE4jTLeknpZGFfNJ
# W9h2zzANBgkqhkiG9w0BAQEFAASCAgCumDaTm1KhRY6xvfp5Hv5ekk4sRSr7ZDSa
# RUblzg7UnCzhq3VdfQ1/79BA5JJIq7dpz0zPxCbhlBA6zgjgaktq/h3ytCA4QCoc
# mhKI+dDlH7Z9EkdStceD+lgWCQzqAEDVliyaWRENqUKbhx0NOTxxfST8jZyHWuCY
# jSdvLd63g5mz/EpIPY5o7lU0GK/Jl6Ig+k+pdIpNa87EyeS0770svgJGitGTSRQ6
# c57Ch3YiXlBoX77jIixHcQrXFkvsqP/RfukHkfOCW7hJQ4WZwQ/uvGKQefAhs0g/
# YAJz7jRJ7xA0XO6YLqz6RdQR/6+FRDILPKFIq2Tyj3Pej4/hqUG980WvXxrKrT4O
# Jpowqz/DueIyOH4A8f9Cqf1DvpUK0QMjxnhs7nlyP7MgSx2HQe5ic8uYQ3d9EDcz
# PIlVRmXjCazAKwu27DnUlUlEcGgddMQZeBOOWpuzNtsbhLyQbSywZJkLIk6fF+3D
# 2B+o8rCo67pGZTJb3O212AOsauwqDBKIkd0Am6Q+2w0wlP9a9LystAZX1S9w41oj
# e4HM8s7AYd3qXnamojxe5iaLTbFLAXvTnU2s4jKdg97l8TYmAWtd2M8WtOjUIAJu
# qulZ/NamFTMxMl5C33v47gKV1F1HD9Zj0W/xUD5uLIh6A9w9b6r7/Zh7hdWv67hq
# et80fh/DZg==
# SIG # End signature block
