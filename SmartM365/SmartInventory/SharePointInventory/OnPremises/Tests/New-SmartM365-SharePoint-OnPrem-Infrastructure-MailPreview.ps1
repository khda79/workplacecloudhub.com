[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CsvRoot,
    [Parameter(Mandatory)][string]$OutputPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$folder = Split-Path $PSScriptRoot -Parent
$scriptPath = Join-Path $folder 'SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1'
$coreManifest = Join-Path $folder '..\..\..\Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
Import-Module $coreManifest -MinimumVersion '1.0.53' -ErrorAction Stop
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count) { throw 'Infrastructure collector does not parse.' }
foreach ($name in @('Get-ObservedProperty','Get-InfrastructureFarmEdition','New-InfrastructureMailHtml')) {
    $functionAst = @($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name}, $true))
    if ($functionAst.Count -ne 1) { throw "Missing mail function: $name" }
    . ([scriptblock]::Create($functionAst[0].Extent.Text))
}
$script:BuildEditionMapPath = Join-Path $folder 'SharePoint-OnPrem-BuildEditions.psd1'
function WriteLog { param($Message, $Level) Write-Verbose $Message }
$rows = @{}
foreach ($kind in @('Farms','Servers','ServiceApplications','WebApplications','WebApplicationZones','ContentDatabases')) {
    $path = Join-Path $CsvRoot "SharePoint_OnPrem_$kind.csv"
    $rows[$kind] = @(Import-Csv -LiteralPath $path -Delimiter ';' -ErrorAction Stop)
}
if ($rows.Farms.Count -ne 1) { throw 'Expected exactly one farm row.' }
$html = New-InfrastructureMailHtml -Status Preview -Rows $rows -TenantName ([string]$rows.Farms[0].TenantKey) -HostName 'Offline CSV preview'
$parent = Split-Path $OutputPath -Parent
if (-not (Test-Path -LiteralPath $parent -PathType Container)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
[IO.File]::WriteAllText($OutputPath, $html, [Text.UTF8Encoding]::new($false))
Write-Output "Preview written: $OutputPath"

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBkIDGTiSO+xWoS
# jEvJFz27yoehEEvupD8sMKayGwb+taCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCRDdsXaoIDuuCp15I63IYa
# cjVQiYVKcZr7H2hIsTd/zzANBgkqhkiG9w0BAQEFAASCAYCSAJ76k0h2a/C3kapG
# rFyPhVAssd5VgyBSVa13u7rjIgVuhYseF9XHMWmiuNJRCvlnZf+vohKRg2ZingTa
# Pw4Y3QFT2qa/PmBMWd6gt/sL2CXCr+tQte3lidwKbqGqNt/qqIpxW677xNdOoEhC
# ZwAV6aKm9hAWDj5AkHy2m9hR+5HRtdlMaiKn/5eQ/TPA2JIFGNtSSR1xk2BT3qf3
# Ykikwp/V4aARpXsdWyg2UXRcIfyjXxQub+jbf/tuC2TD1R94fdoxXj4dfPPp+32f
# I6NDlVzYiOaZ0zN29BLvG92WEzA1safjTFJfWWfAoScfXkvZ1eUYiCDHzSrrrcqU
# /a/3hZ1vXqlOs5e276Rd7XIC+gwRO7Uz6dqvQ3sZvqGQvpsOqGKonpU+NVsU72zu
# 6xbowX4FVPbwhMngzWQWtywGY9Gv9eMS8gfzyAN+Q+Z9yfS1BMm2s2dVmF4+hRSe
# iiM/4/mbOwzAkezZmTjOmTIou/avzzTvA1fE6NtubZEQPCM=
# SIG # End signature block
