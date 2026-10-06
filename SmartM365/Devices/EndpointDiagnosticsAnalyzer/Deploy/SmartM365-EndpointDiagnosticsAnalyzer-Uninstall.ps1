#Requires -Version 5.1
# Safe removal contract: 0.3.0-preview17
[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
param([ValidateSet('CurrentUser','AllUsers')][string]$InstallScope='CurrentUser',[string]$InstallPath='',[switch]$RemoveUserData)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(-not$InstallPath){$InstallPath=if($InstallScope -eq 'AllUsers'){Join-Path $env:ProgramFiles 'SmartM365\EndpointDiagnosticsAnalyzer'}else{Join-Path $env:LOCALAPPDATA 'Programs\SmartM365\EndpointDiagnosticsAnalyzer'}}
$target=[IO.Path]::GetFullPath($InstallPath).TrimEnd('\')
if(-not(Test-Path -LiteralPath $target -PathType Container)){return [pscustomobject]@{Result='NotInstalled';InstallPath=$target;UserDataRemoved=$false}}
if($target -eq [IO.Path]::GetPathRoot($target).TrimEnd('\') -or ((Get-Item -LiteralPath $target).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Unsafe uninstall target.'}
$metadataPath=Join-Path $target 'SmartM365-EndpointDiagnosticsAnalyzer.installation.json'
$metadata=Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
if($metadata.ProductName -ne 'Smart Endpoint Diagnostics Analyzer' -or $metadata.InstallScope -ne $InstallScope){throw 'Application ownership or installation scope could not be verified.'}
$owned=@('SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt','SmartM365-EndpointDiagnosticsAnalyzer.installation.json','SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1','HardwareReadiness.ps1','SmartM365.GuiSplash.ps1','Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd','WorkplaceCloudHub-lockup-WPF.png','WorkplaceCloudHub.ico')
foreach($entry in Get-ChildItem -LiteralPath $target -Force){
 if($entry.PSIsContainer -or $entry.Name -notin $owned -or ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint)){throw "Uninstall blocked: unowned entry must be preserved: $($entry.Name)"}
}
$programs=if($InstallScope -eq 'AllUsers'){Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\SmartM365'}else{Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\SmartM365'}
$shortcut=Join-Path $programs 'Smart Endpoint Diagnostics Analyzer.lnk'
$removeShortcut=$false
if($metadata.PSObject.Properties['ShortcutPath'] -and $metadata.ShortcutPath -eq $shortcut -and (Test-Path -LiteralPath $shortcut -PathType Leaf)){
 $shell=New-Object -ComObject WScript.Shell
 $link=$shell.CreateShortcut($shortcut)
 $removeShortcut=([IO.Path]::GetFullPath($link.TargetPath) -eq (Join-Path $target 'Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd'))
}
$userDataRemoved=$false
if(-not$PSCmdlet.ShouldProcess($target,'Remove verified application files')){return [pscustomobject]@{Result='NoChanges';InstallPath=$target;UserDataRemoved=$false}}
Remove-Item -LiteralPath $target -Recurse -Force
if($removeShortcut){Remove-Item -LiteralPath $shortcut -Force}
if($RemoveUserData){
 $dataPath=[IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'SmartM365\EndpointDiagnosticsAnalyzer'))
 $dataParent=[IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'SmartM365')).TrimEnd('\')+'\'
 if(-not$dataPath.StartsWith($dataParent,[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe user-data path.'}
 if(Test-Path -LiteralPath $dataPath){
  if((Get-Item -LiteralPath $dataPath).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'User-data reparse point cannot be removed.'}
  Remove-Item -LiteralPath $dataPath -Recurse -Force
 }
 $userDataRemoved=$true
}
[pscustomobject]@{Result='PASS';InstallPath=$target;ShortcutRemoved=$removeShortcut;UserDataRemoved=$userDataRemoved}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBKIUZCJ6BOZmKX
# QxGpkaBywmsRuMg1GxE04jqe55Ca7KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDW6fsgH9DSdDsXYmrjS+RR
# 6u6HqvwIOpZn7/bQ78pmtzANBgkqhkiG9w0BAQEFAASCAYBUDohF9XPf3MxSx9e/
# mLLa1nxAJEOY2fO4iRviz76Sh7tIxp7QZaUfP3tp36spaSGyBmOTXkuUkzvyh/Rk
# C0PggmTPFNvFw9vPGg0GTry+noOfTZeq6FFnGR4PksXnY0iAH9JWbqlR/NFIXpXm
# 1qK/bx1M41EXlePJQTTiKkLxvSmPdEb4pjMuLJQtKmDAZU7n5hUcBhCcWRZO+4mq
# aG57tlOtpIu8NKYsMTvA388grBgenHE0f5nFOJZtjdd9OnM9Zav3mJnh3fF19Jm7
# Pv9byjSXS42q9lLX33a0SjojGW8CGW/bvjZF+mXWrJycEibvGnaFG7MWbHg0OhLL
# xuRkbUeyK0lQNTTpezZ/nZp2dMQPb0jRbwliOPNgfpAx4JJx1TL+AvuDD2LiQL5z
# qu8J9Y6E8gtGCETtVl6F66fvVP3qNvRwRPC6qrq/LxqO7kLgAtYeyZwTr3pqS+mm
# +RPbewp7vyTYnl8R0WqH/1g0Zdgt17PYUGV+Hd10O1s4d0I=
# SIG # End signature block
