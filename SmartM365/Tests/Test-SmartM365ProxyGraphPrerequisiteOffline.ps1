#Requires -Version 5.1
<# Replays the prerequisite against synthetic modules and a mocked CurrentUser installer. No real installation or connection. #>
[CmdletBinding()]
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'
$checks = 0
function Check([bool]$Ok, [string]$Message) { if (-not $Ok) { throw $Message }; $script:checks++ }
function WriteLog { param($Message, $Level); $script:messages += [string]$Message }
function Install-Module {
    [CmdletBinding()]
    param($Name, $RequiredVersion, $Repository, $Scope, [switch]$Force, [switch]$AllowClobber)
    $script:installCalls++
    Check ($Name -eq 'Microsoft.Graph.Authentication' -and $RequiredVersion -eq '2.41.2') 'Bootstrap module/version must be constrained'
    Check ($Repository -eq 'PSGallery' -and $Scope -eq 'CurrentUser') 'Only CurrentUser installation from PSGallery is allowed'
    Check ($Force -and $AllowClobber -and $PSBoundParameters.ErrorAction -eq 'Stop') 'Installer must fail explicitly'
    if ($script:installMode -eq 'Denied') { throw 'Synthetic CurrentUser access denied' }
    if ($script:installMode -eq 'Unavailable') { throw 'Synthetic PSGallery unavailable' }
    if ($script:installMode -eq 'Invisible') { return }
    New-GraphFixture -Version $RequiredVersion
}
function Connect-MgGraph { throw 'Connection forbidden in offline test' }
$path = Join-Path $PSScriptRoot '../SmartInventory/ExchangeInventory/OnPremises/ProxyAddresses/SmartM365-Check-ProxyAddresses-Exchange.ps1'
if ($SourcePath) { $path = $SourcePath }
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
Check ($errors.Count -eq 0) 'Collector parser failed'
$definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Assert-ProxyAddressesGraphPrerequisite' }, $true)
Check ($null -ne $definition) 'Prerequisite function missing'
. ([scriptblock]::Create($definition.Extent.Text))
$call = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Assert-ProxyAddressesGraphPrerequisite' }, $true)
$snapin = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-PSSnapin' }, $true)
Check ($call.Extent.StartOffset -lt $snapin.Extent.StartOffset) 'Graph check must precede Exchange loading/collection'
Check ($call.Extent.Text -match '-EnableSharePointUpload \$global:EnableSharePointUpload -SendMailMode \$SendMailMode -SmtpServer \$SmtpServer') 'Call must use effective configuration'
Check ($call.Extent.StartOffset -gt $ast.Extent.Text.IndexOf('$ScriptLocalConfig = Get-ScriptLocalConfig')) 'Graph check must follow config resolution'
$savedPath = $env:PSModulePath
$root = Join-Path ([IO.Path]::GetTempPath()) ('proxy-graph-' + [guid]::NewGuid().ToString('N'))
$moduleBase = Join-Path $root 'Microsoft.Graph.Authentication'
$null = New-Item -ItemType Directory -Path $moduleBase -Force
$manifest = Join-Path $moduleBase 'Microsoft.Graph.Authentication.psd1'
$body = Join-Path $moduleBase 'Microsoft.Graph.Authentication.psm1'
function New-GraphFixture([string]$Version) {
    "@{RootModule='Microsoft.Graph.Authentication.psm1';ModuleVersion='$Version';GUID='ead3110d-d8f9-4a6b-976c-2d9ea09cf0fb';FunctionsToExport=@('Get-MgContext');CmdletsToExport=@();AliasesToExport=@();VariablesToExport=@()}" | Set-Content -LiteralPath $manifest -Encoding UTF8
    "function Get-MgContext { throw 'Graph context must not be queried by module prerequisite' }" | Set-Content -LiteralPath $body -Encoding UTF8
}
$script:installCalls = 0
$script:installMode = 'Unavailable'
try {
    Check (@(Get-Module Microsoft.Graph.Authentication).Count -eq 0) 'Run in a fresh process without a loaded Graph module'
    $env:PSModulePath = $root
    $cases = @(
        @{Name='SharePoint'; Upload=$true; Mode='SMTP'; Smtp='relay'; Needs=$true},
        @{Name='Graph mail'; Upload=$false; Mode='Graph'; Smtp='relay'; Needs=$true},
        @{Name='Both'; Upload=$false; Mode='Both'; Smtp='relay'; Needs=$true},
        @{Name='Implicit Graph'; Upload=$false; Mode=''; Smtp=''; Needs=$true},
        @{Name='Explicit SMTP'; Upload=$false; Mode='SMTP'; Smtp='relay'; Needs=$false},
        @{Name='Implicit SMTP'; Upload=$false; Mode=$null; Smtp='relay'; Needs=$false}
    )
    foreach ($case in $cases) {
        $script:installCalls = 0
        $messages = @(); $failure = $null
        try { Assert-ProxyAddressesGraphPrerequisite $case.Upload $case.Mode $case.Smtp } catch { $failure = $_ }
        Check (($null -ne $failure) -eq $case.Needs) ($case.Name + ': missing module gate')
        Check ($script:installCalls -eq [int]$case.Needs) ($case.Name + ': install only if required and absent')
        if ($case.Needs) {
            Check ($failure.Exception.Message -match 'Graph Authentication prerequisite failed.*PSModulePath=') ($case.Name + ': actionable diagnostic')
            Check ($failure.Exception.Message -match 'Synthetic PSGallery unavailable.*Scope=CurrentUser') ($case.Name + ': installation failure preserves cause and scope')
        }
    }
    New-GraphFixture -Version '1.2.3'
    foreach ($case in $cases | Where-Object Needs) {
        $script:installCalls = 0
        $messages = @()
        Assert-ProxyAddressesGraphPrerequisite $case.Upload $case.Mode $case.Smtp
        $loaded = Get-Module Microsoft.Graph.Authentication
        Check ($loaded.ModuleBase -eq $moduleBase) ($case.Name + ': only fixture loaded')
        Check (($messages -join '|') -match 'prerequisite OK; PowerShell=.*Version=1.2.3; Path=') ($case.Name + ': runtime/version/path logged')
        Check ($null -ne (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) ($case.Name + ': module remains available to later publication')
        Check ($script:installCalls -eq 0 -and $loaded.Version -eq [version]'1.2.3') ($case.Name + ': available module reused without upgrade')
        Remove-Module Microsoft.Graph.Authentication -Force
    }
    $script:installCalls = 0
    Import-Module -Name $manifest -Global -ErrorAction Stop
    # A loaded module remains usable even if its discovery path is no longer listed.
    $env:PSModulePath = Join-Path $root 'not-a-module-root'
    Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay'
    Check ($script:installCalls -eq 0 -and (Get-Module Microsoft.Graph.Authentication).Version -eq [version]'1.2.3') 'Loaded module reused without installation'
    Remove-Module Microsoft.Graph.Authentication -Force
    $env:PSModulePath = $root
    "throw 'Synthetic incompatible Graph module'" | Set-Content -LiteralPath $body -Encoding UTF8
    $failure = $null
    try { Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay' } catch { $failure = $_ }
    Check ($null -ne $failure -and $failure.Exception.Message -match 'Synthetic incompatible Graph module') 'Import failure must stop prerequisite'
    Check ($script:installCalls -eq 0) 'An incompatible available module must not be replaced automatically'
    Check ($env:PSModulePath -ceq $root) 'Prerequisite must not add PS7 paths'
    $messages = @()
    Assert-ProxyAddressesGraphPrerequisite $false 'SMTP' 'relay'
    Check (($messages -join '|') -match 'prerequisite skipped') 'SMTP-only skips even an incompatible module'
    Remove-Item -LiteralPath $manifest, $body -Force
    foreach ($mode in @('Denied','Invisible','Success')) {
        $script:installMode = $mode; $script:installCalls = 0
        $messages = @(); $failure = $null
        try { Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay' } catch { $failure = $_ }
        Check ($script:installCalls -eq 1) ($mode + ': one installation attempt')
        if ($mode -eq 'Success') {
            Check ($null -eq $failure -and (Get-Module Microsoft.Graph.Authentication).Version -eq [version]'2.41.2') 'Missing module installed then imported'
            Assert-ProxyAddressesGraphPrerequisite $true 'SMTP' 'relay'
            Check ($script:installCalls -eq 1) 'Second check does not reinstall'
            Check (($messages -join '|') -match 'absent; installing version 2.41.2.*Scope=CurrentUser') 'Installation scope and version logged'
        }
        else {
            $pattern = if ($mode -eq 'Denied') { 'Synthetic CurrentUser access denied' } else { 'still not discoverable' }
            Check ($null -ne $failure -and $failure.Exception.Message -match $pattern) ($mode + ': collection must remain blocked')
            Check ($null -eq (Get-Module Microsoft.Graph.Authentication)) ($mode + ': no false successful import')
        }
    }
}
finally {
    Get-Module Microsoft.Graph.Authentication | Where-Object { $_.ModuleBase -eq $moduleBase } | Remove-Module -Force
    $env:PSModulePath = $savedPath
    if (-not ([IO.Path]::GetFullPath($root)).StartsWith(([IO.Path]::GetFullPath([IO.Path]::GetTempPath())).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe fixture cleanup path' }
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
Write-Output ("PASS: {0} ProxyAddresses Graph prerequisite checks; {1}; synthetic only." -f $checks, $PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPneXj3ob/eLF2
# xVy+wplA7D+bQkjUBqor7Ri40vYTI6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIN++kqFa2jB+RwGWx3K1ewj+gIPqlMpr7dLAE4thFqWHMA0GCSqG
# SIb3DQEBAQUABIIBgJYGzmSaPkdyxlyyy3KDs21+nr/zPgqw6IutvqUifANSt/Hn
# gXTFmHcL8DF2q+uA5LiV62vo3dBympfJfLh5+b0ds4mZbDnKKR0ObPu975N8h+HG
# 1RjbPW0Dd7Qdu/KTjgUF+Mh0Yt7Xc082IyMIpwpWR4sy7F2MsfF4rx8ZaTPrv4kn
# MXwRoXN+Hk0iSb0hrfDMEfKTUpEljeig8PM5yEEHzAGu/SZ5I/bdWfLfDLiUKuoK
# jIokXxt/m9rXm1MtI8/KFvwOP3/3eyc39ukIBlJ/nYPHIuUyL9oHNpMXo8A/R2qQ
# IglD47JyURmFxoY0jyGkrL1msbYMYf+AmxNv2/yaf85LGkARhmF8r3EivV763vUg
# AgNPF1EOIlpTihTRVpMYLd61RpEHEEoOUNuCFUhYB7upudNLq5t2wijzilt1yFqQ
# bhhzTl1iHyfRoUiUdx3v8lAalpBQ+TjIjR3BAXg60m1BcGOdNNT67yJffjpqQlF0
# 9jpJGfM2If1gFcF8saGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMwMTE1
# NTVaMC8GCSqGSIb3DQEJBDEiBCAR4Vhl4oCmm47/tdWlQAeWMdq4cJdF4GsER8/s
# S+MHbTANBgkqhkiG9w0BAQEFAASCAgBQOuxE9yu+sW5wM9SNXZAUeXrg3nRwWYvB
# 6olZEYwS58/nfzpGi5c1wqbrJ+xG/JnnN5wxtUEm0+iKuNm8meOfJsJWSaOGXpq2
# NssfS91i7qsgTHHpZmVGibmwBhHs+9HnnRUgi7UVJ3KwRNuTw/kinL549ve0jgqW
# T9yhePO58WECLKJ9NQAx4ctZLroY/w7YM6Z053nPMz+7mW3jbTbTqYW5z9LimlC8
# qv49Amxz5g8qo/HVHHPAiZxcEdNeLEDKgVf2++4i2bToj9oL1VXRZSBa5gHlWpQs
# zkm+hYrZqmuxwEaIoObWQenyozoFsxqpGZGD43FK7to+TnHFr09dPVRWGaPn8mqd
# EgLCgYJPspYjf2BpkUCPjcz53LF1IGAj4Tqi0ZT33yij9bmprgdlzx+8jn2SNUHr
# cLbPBWciAmj6xtTk4BOYZDvcF78cgJuztwwia12DWtId6qXFIE4kAxP+Uj7tGi7d
# yKJ529GHWZucMrOhHw1iE1x1uck8/ZjEygW9dPdxLIy1E/Sx1wJPmR34Rc6UU9la
# Mj8/cmMvIyU72PIrB0Y5hRCm2SA0S+NyKCJHROJfqxXR7cv9RbHeBzuW+8iQrviw
# X18XP81mFoXNoD7tq/54VrHa5k9rzPAOxH2q4QanHy3gnPe+8iqktYu4NC4cMxnO
# e2G7HngrlQ==
# SIG # End signature block
