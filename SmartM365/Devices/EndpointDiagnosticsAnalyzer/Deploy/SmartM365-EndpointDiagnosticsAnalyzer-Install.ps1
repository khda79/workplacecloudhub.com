#Requires -Version 5.1
# Runtime installation contract: 0.3.0-preview17
[CmdletBinding()]
param(
 [ValidateSet('CurrentUser','AllUsers')][string]$InstallScope='CurrentUser',
 [string]$InstallPath='',
 [ValidateSet('PowerShellGallery','Intune','Local')][string]$PackageSource='Local',
 [switch]$SkipShortcut
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$product='Smart Endpoint Diagnostics Analyzer'
if($InstallScope -eq 'AllUsers'){
 $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
 if(-not([Security.Principal.WindowsPrincipal]::new($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'AllUsers installation requires elevation.'}
}
if(-not$InstallPath){$InstallPath=if($InstallScope -eq 'AllUsers'){Join-Path $env:ProgramFiles 'SmartM365\EndpointDiagnosticsAnalyzer'}else{Join-Path $env:LOCALAPPDATA 'Programs\SmartM365\EndpointDiagnosticsAnalyzer'}}
$target=[IO.Path]::GetFullPath($InstallPath).TrimEnd('\')
if($target -eq [IO.Path]::GetPathRoot($target).TrimEnd('\')){throw 'A filesystem root is not an installation target.'}
$sourceRoot=Split-Path -Parent $PSScriptRoot
$versionFile='SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt'
$metadataFile='SmartM365-EndpointDiagnosticsAnalyzer.installation.json'
$required=@('SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1','HardwareReadiness.ps1','SmartM365.GuiSplash.ps1','Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd','WorkplaceCloudHub-lockup-WPF.png','WorkplaceCloudHub.ico')
$version=Get-Content -LiteralPath (Join-Path $sourceRoot $versionFile) -Raw | ConvertFrom-Json
if($version.SchemaVersion -ne 1 -or $version.ProductName -ne $product -or -not$version.PackageVersion -or -not$version.PSObject.Properties['RuntimeHashes']){throw 'Invalid runtime integrity manifest; rebuild the package.'}
foreach($name in $required){
 $source=Join-Path $sourceRoot $name
 $expected=$version.RuntimeHashes.PSObject.Properties[$name]
 if(-not$expected -or -not(Test-Path -LiteralPath $source -PathType Leaf) -or (Get-Item -LiteralPath $source).Length -eq 0){throw "Required nonempty runtime file/hash missing: $name"}
 if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne [string]$expected.Value){throw "Runtime integrity failure: $name"}
 if([IO.Path]::GetExtension($name) -eq '.ps1'){
  $signature=Get-AuthenticodeSignature -LiteralPath $source
  if(-not$signature.SignerCertificate -or $signature.SignerCertificate.Thumbprint -ne 'D70ECB7B00377EBFB76B304C08DFC6620584E114' -or $signature.Status -notin @('Valid','NotTrusted')){throw "Runtime signature integrity failure: $name ($($signature.Status))"}
 }
}
if(Test-Path -LiteralPath $target){
 if((Get-Item -LiteralPath $target).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Reparse-point installation targets are not supported.'}
 if(@(Get-ChildItem -LiteralPath $target -Force).Count){
  $prior=Get-Content -LiteralPath (Join-Path $target $metadataFile) -Raw | ConvertFrom-Json
  if($prior.ProductName -ne $product){throw 'Nonempty destination does not belong to this application.'}
  if($prior.PackageSource -eq 'Intune' -and $PackageSource -ne 'Intune'){throw 'Intune owns this installation. Update it through Intune.'}
 }
}
$parent=Split-Path -Parent $target
New-Item -ItemType Directory -Path $parent -Force | Out-Null
$stage=$target+'.stage-'+[guid]::NewGuid().ToString('N')
$backup=$target+'.backup-'+[guid]::NewGuid().ToString('N')
New-Item -ItemType Directory -Path $stage | Out-Null
$shortcutPath='';$movedOld=$false;$activated=$false
try{
 foreach($name in @($versionFile)+$required){Copy-Item -LiteralPath (Join-Path $sourceRoot $name) -Destination (Join-Path $stage $name)}
 foreach($name in $required){if((Get-FileHash -LiteralPath (Join-Path $stage $name)).Hash -ne [string]$version.RuntimeHashes.PSObject.Properties[$name].Value){throw "Staged integrity failure: $name"}}
 if(-not$SkipShortcut){
  $programs=if($InstallScope -eq 'AllUsers'){Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\SmartM365'}else{Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\SmartM365'}
  $shortcutPath=Join-Path $programs 'Smart Endpoint Diagnostics Analyzer.lnk'
 }
 $metadata=[ordered]@{SchemaVersion=1;ProductName=$product;PackageVersion=[string]$version.PackageVersion;PackageSource=$PackageSource;InstallScope=$InstallScope;AutomaticUpdateEnabled=$false;InstalledAtUtc=[datetime]::UtcNow.ToString('o');ShortcutPath=$shortcutPath;VersionManifestHash=(Get-FileHash -LiteralPath (Join-Path $stage $versionFile)).Hash}
 $metadata | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $stage $metadataFile) -Encoding UTF8
 if(Test-Path -LiteralPath $target){Move-Item -LiteralPath $target -Destination $backup;$movedOld=$true}
 Move-Item -LiteralPath $stage -Destination $target;$activated=$true
 if(-not$SkipShortcut){
  New-Item -ItemType Directory -Path $programs -Force | Out-Null
  $shell=New-Object -ComObject WScript.Shell
  $shortcut=$shell.CreateShortcut($shortcutPath)
  $shortcut.TargetPath=Join-Path $target 'Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd'
  $shortcut.WorkingDirectory=$target;$shortcut.IconLocation=Join-Path $target 'WorkplaceCloudHub.ico'
  $shortcut.Description=$product;$shortcut.WindowStyle=7;$shortcut.Save()
 }
}catch{
 if($activated -and (Test-Path -LiteralPath $target)){Move-Item -LiteralPath $target -Destination ($stage+'.failed')}
 if($movedOld -and (Test-Path -LiteralPath $backup)){Move-Item -LiteralPath $backup -Destination $target}
 throw
}
[pscustomobject]@{Result='PASS';ProductName=$product;PackageVersion=[string]$version.PackageVersion;PackageSource=$PackageSource;InstallScope=$InstallScope;InstallPath=$target;ShortcutPath=$shortcutPath;AutomaticUpdateEnabled=$false;IntegrityVerified=$true;PreviousInstallationPath=if($movedOld){$backup}else{''}}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBPhHrPJAiy/AMi
# Vjql+CG9DD5982lmMotcUkT9QyZyFaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAuavELL3DtZIYZI9lxuVGb
# cGteGRu2DGyvMZhMLL2rHDANBgkqhkiG9w0BAQEFAASCAYAGo6g9sEHdn6U8UQbL
# DGxiZP/7mEue8LSvVFbHkI3Tm4tMJDNN0Flgb3xLO7mKiTbKWXDgw2xRhzCFCe0v
# QulRA46TNye78LrCIjTEB0esxlDiH1P/7H4PyI4whhNQXFGr7UdU3Rz+I4+a/w4R
# QT0teivr6rFXEz4dil0tRNHeH2IIukerp8oWNFgIYxG+zWIVQZIGvKlbY0KwY/xQ
# MMYRwEkJbjxWvGV4PSFFhZic6Nz7zBzGs7EFflt6zVlvS9lFtgazD4s0kPenGop9
# p6Re0g6nIk1kpJrceFZ1NLkOaBESzruYPjrFBezZJgCDHP3pnoyR4ZaiWeVvWBuL
# oO1y/42NtEhq3X/rOQEZ+kuS09nceY00bH4UvE80TrBvWlBwhirBRD5SFnyVgEsE
# DS1Fbm4LiXF5eF6OS3TFFQnKikLY/WE8SGagF0u3fytpv03ZpejBKB4PwUNfDadP
# BeDOZ/LS7ToSCGkPnUapJPqPW0cftJXle36yXh3j8zGClQE=
# SIG # End signature block
