#Requires -Version 5.1
<#
.SYNOPSIS
Offline tests of SmartInventoryConfigRootPath: path mapping, templates and fail-closed reads.
.VERSION
1.1.0
.NOTES
Builds a synthetic SmartM365 tree under TEMP (copies of the JSON transport module, the tenant context and the
account-classification module) and runs the checks in Windows PowerShell 5.1 and PowerShell 7 when both are
available. No real configuration or tenant is touched.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

$inner = @'
param([string]$Root)
$ErrorActionPreference = 'Stop'
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Reject([scriptblock]$Action, [string]$Message) { $failed = $false; try { & $Action | Out-Null } catch { $failed = $true }; Check $failed $Message }
function Get-Long([string]$Path) { (Resolve-Path -LiteralPath $Path).ProviderPath }
$ws = Join-Path (Split-Path $Root -Parent) 'WS'
$code = Join-Path $Root 'SmartInventory'
$globalPath = Join-Path $Root 'Config\SmartM365.global.local.json.txt'
function Set-Global([string]$Value) {
    $document = [ordered]@{ WorkspaceRootPath = $ws; SmartInventoryConfigRootPath = $Value }
    [IO.File]::WriteAllText($globalPath, ($document | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
}
Set-Global ''
. (Join-Path $Root 'Config\SmartM365-TenantContext.ps1')
$foo = Join-Path $code 'Area\SmartM365-Foo.local.json'

Check ((ConvertTo-SmartM365InventoryConfigLocation $foo) -eq $foo) 'Empty setting changed the path.'
Check ((Resolve-SmartM365JsonConfigurationPath $foo) -eq "$foo.txt") 'Empty setting changed the resolution.'
Check ((Get-SmartM365JsonTemplateName "$foo.txt") -eq "$foo.template") 'Empty setting changed the template name.'

Set-Global '{{WorkspaceRootPath}}\SmartInventory\Config'
$target = Join-Path $ws 'SmartInventory\Config\Area\SmartM365-Foo.local.json'
Check ((ConvertTo-SmartM365InventoryConfigLocation $foo) -eq $target) 'Mapping to the configuration root is wrong.'
Check (Test-Path -LiteralPath (Split-Path $target -Parent) -PathType Container) 'Target folder was not created.'
Check ((ConvertTo-SmartM365InventoryConfigLocation $target) -eq $target) 'Mapping is not idempotent.'
Check ((ConvertFrom-SmartM365InventoryConfigLocation "$target.txt") -eq "$foo.txt") 'Inverse mapping is wrong.'
Check ((Get-SmartM365JsonTemplateName "$target.txt") -eq "$foo.template") 'Module template is not read from the code folder.'
Check ((Get-SmartM365JsonTemplatePath "$target.txt") -eq "$foo.template") 'Tenant context template is not read from the code folder.'
$config = Read-SmartM365JsonConfig -Path $foo -Required
Check ($config.Setting -eq 1) 'First run did not use the code folder template.'
Check ((Test-Path -LiteralPath "$target.txt") -and -not (Test-Path -LiteralPath "$foo.txt") -and -not (Test-Path -LiteralPath $foo)) 'First run created the file in the wrong folder.'
[IO.File]::WriteAllText("$foo.template", '{"Setting":1,"Added":2}')
$null = Read-SmartM365JsonConfig -Path $foo -Required
Check (((Get-Content -LiteralPath "$target.txt" -Raw | ConvertFrom-Json).Added) -eq 2) 'Template enrichment was not written to the configuration root.'
$bar = Join-Path $code 'Area\SmartM365-Bar.local.json'
[IO.File]::WriteAllText("$bar.txt", '{"A":1}')
Reject { Resolve-SmartM365JsonConfigurationPath $bar } 'A file left in the code folder was read from a second source.'
Remove-Item -LiteralPath "$bar.txt"
$tenantPath = Join-Path $Root 'Config\Tenants\x.local.json'
Check ((ConvertTo-SmartM365InventoryConfigLocation $tenantPath) -eq $tenantPath) 'A path outside SmartInventory was mapped.'
Set-Global ($code + '\Private'); Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'A root inside the code folder was accepted.'
Set-Global '{{DataAllRootPath}}\Config'; Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'An unsupported token was accepted.'
Set-Global 'Relative\Config'; Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'A relative root was accepted.'

if ($PSVersionTable.PSVersion.Major -ge 7) {
    # The account-classification module requires PowerShell 7.
    Set-Global '{{WorkspaceRootPath}}\SmartInventory\Config'
    Import-Module (Join-Path $code 'Common\SmartM365.AccountClassification.psd1') -Force
    $classification = Join-Path $ws 'SmartInventory\Config\Config\AccountClassification.local.json.txt'
    New-Item -ItemType Directory -Path (Split-Path $classification -Parent) -Force | Out-Null
    [IO.File]::WriteAllText($classification, '{}')
    Check ((Resolve-SmartM365AccountClassificationPath) -eq (Get-Long $classification)) 'Account classification was not read from the configuration root.'
}
'PASSED {0}' -f $script:passed
'@

function New-Fixture {
    $base = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-InvCfgRoot-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $root = Join-Path $base 'SmartM365'
    $copies = @(
        'Modules\SmartM365.Core\SmartM365.JsonTransport.psd1', 'Modules\SmartM365.Core\SmartM365.JsonTransport.psm1',
        'Config\SmartM365-JsonTransport.policy.psd1', 'Config\SmartM365-TenantContext.ps1', 'Config\SmartM365.global.local.json.template',
        'Config\Tenants\tenant.local.json.template', 'Config\Tenants\tenant.mail.local.json.template',
        'SmartInventory\Common\SmartM365.AccountClassification.psd1', 'SmartInventory\Common\SmartM365.AccountClassification.psm1',
        'SmartInventory\Config\AccountClassification.local.json.template')
    foreach ($relative in $copies) {
        $destination = Join-Path $root $relative
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $repoRoot $relative) -Destination $destination
    }
    [IO.File]::WriteAllText((Join-Path $root 'Config\Tenants\test.local.json.txt'), '{"ProfileKey":"test","OrganizationKey":"contoso","EnvironmentKey":"test","TenantKey":"contoso-test"}')
    $area = Join-Path $root 'SmartInventory\Area'
    New-Item -ItemType Directory -Path $area -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $area 'SmartM365-Foo.ps1'), '# owner script')
    [IO.File]::WriteAllText((Join-Path $area 'SmartM365-Foo.local.json.template'), '{"Setting":1}')
    $innerPath = Join-Path $base 'inner.ps1'
    [IO.File]::WriteAllText($innerPath, $inner, (New-Object Text.UTF8Encoding($true)))
    [pscustomobject]@{ Base = $base; Root = $root; Inner = $innerPath }
}

$engines = @()
$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (Test-Path -LiteralPath $windowsPowerShell) { $engines += $windowsPowerShell }
$pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
if ($pwsh) { $engines += $pwsh.Source }
if (-not $engines.Count) { throw 'No PowerShell engine found.' }
foreach ($engine in $engines) {
    $fixture = New-Fixture
    try {
        $output = & $engine -NoProfile -ExecutionPolicy Bypass -File $fixture.Inner -Root $fixture.Root 2>&1
        $text = ($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
        if ($LASTEXITCODE -ne 0 -or $text -notmatch 'PASSED \d+') { throw ("{0} failed:{1}{2}" -f $engine, [Environment]::NewLine, $text) }
        Write-Host ('{0}: {1}' -f (Split-Path $engine -Leaf), ([regex]::Match($text, 'PASSED \d+').Value)) -ForegroundColor Green
    }
    finally { Remove-Item -LiteralPath $fixture.Base -Recurse -Force -ErrorAction SilentlyContinue }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCu6E/mWTitPazG
# t4ao0WhThyd6FNngklX9u5vpanKVT6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIA2QnA8S9yJ9yijlmMROIAE0WcjL74hqFhDkX5MvbsFVMA0GCSqG
# SIb3DQEBAQUABIIBgKfeWTwyAPXWKCSXyEEfnquWs61R/opyBKu81wm7bEWvHVTa
# eyjUTaIH5XB61vjX093EjJ0IF0TcO9qnaEzUrGx8pswjTYW7cyVX1atL0seAGSJr
# nUwXFTUMOeHknyNHHM+4aAEbzWi3CTxM/pj8x6LbEkM2iJ+CmzfRkKu/aLKhiCo3
# yLI2DgwzRVpHEdaNtZemhbutaIcNkfLU6MYC1Bckfa2ggbeOmEv17fla3Hf7BpjD
# XY3T1YXQ3t9azZSnbXVDuKcM7EiBFVxOWeO3uyZBMqOXqfRVJ0kokZ/LTK6et32r
# Cd/Qptd/k1f1AB/Hq45Us/AXhovqx5Cofiejyml1yQo/bc2Vtu4kRK7Q42xcIdTF
# t/dcMgLixto1O0P8DKl4kr2LTBvkaDlStPNWBH0RSTphMUqgB+TJaa2GfpmkpaOV
# Bdc+0bgUVzg2H73SH37dkiplk1uk/snR5+2AepkEaI2HtQ3fWqy3zQB+AY21V0oM
# s+ZHVwzz5UP5rwl8LKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwNzM3
# MzZaMC8GCSqGSIb3DQEJBDEiBCDVL9mXRoTwhaA7bHL/d43iNymGSVPEpI78IfRe
# 0P7ivjANBgkqhkiG9w0BAQEFAASCAgCKcV7dgl7LYOzdMRZ8wk+QSoxmmRHi6xmx
# n9ctcbmM9A2frqFEesDGepcONo280jGJKahyJEaUiiy62BLu8wxVy7CUynSLvoqP
# TyiMWwR1dvvezfIlZ/XnTIIO/yZR3Zof7fNyLFzseOp223hTB1738QakJOykTch5
# 4iPidlM3wCmzaPZ7kJSTHOPnq3mK+g8xuRghL3tKKlis48mp5P3B0qNch2uL/VSC
# U4uJTzjHbTsIsqoMqNrL5WYjW0bIZgAOu0c7g6uX15E+gBz4e9QhTKbEBm8xZrEq
# O1zMut3bLW1Nv5wFuyUSE7Mv+Zys2ZxZ7k1IaYwKNyWymauhFLJTxD/k3yHk1BMv
# BYq8NwMsns6x1h1oFnQtASHbtB5moJBD0ElTqOR4iMEF+6rvLz/H1l4904zHxWHz
# 5ctU/MX4cI5QzjVXvmoarwzXDUR+/9FZJI6phGnZ9A/iHr2rVpoa5U//GyZZqnXa
# nLxD43LUlRDrA2M3LggsLJ0Q7YTbTkLx9DP+lIsihnTt9eJ5rLcPTNPw+z2B3MTJ
# 5Q70D9iHOdC2jT1UU4OTqPRFbi1B0xp+dpOPN9FQx66aIGV92UhGFyCQ6f0ydDxz
# Ig9NqFZ7VVLV9NnGGwPp4L+lqgfrLZEZ07PuyzRkufq7lq550qJ41RIuPgmO1+FL
# B7XPc2/2Lw==
# SIG # End signature block
