#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
param(
    [string]$PackagePath='',
    [string]$OutputRoot=(Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\PowerShellGallery'),
    [string]$Repository='PSGallery',
    [string]$ApiKeyEnvironmentVariable='PSGALLERY_API_KEY',
    [switch]$Execute,
    [switch]$AllowPrereleasePublication,
    [switch]$ForceBuild
)
Set-StrictMode -Version Latest;$ErrorActionPreference='Stop'
if(-not$PackagePath){$build=& (Join-Path $PSScriptRoot 'SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1') -OutputRoot $OutputRoot -Force:$ForceBuild;$PackagePath=[string]$build.PackagePath}
$runtime=Join-Path $PackagePath 'Runtime'
$version=Get-Content -LiteralPath (Join-Path $runtime 'SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt') -Raw|ConvertFrom-Json
foreach($name in @('SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1','HardwareReadiness.ps1','SmartM365.GuiSplash.ps1','Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd','WorkplaceCloudHub-lockup-WPF.png','WorkplaceCloudHub.ico')){
    $hash=$version.RuntimeHashes.PSObject.Properties[$name]
    if(-not$hash -or (Get-FileHash -LiteralPath (Join-Path $runtime $name)).Hash -ne [string]$hash.Value){throw "Packaged runtime integrity mismatch: $name"}
}
foreach($file in Get-ChildItem -LiteralPath $PackagePath -File -Recurse | Where-Object Extension -in @('.ps1','.psm1','.psd1')){
    $signature=Get-AuthenticodeSignature -LiteralPath $file.FullName
    if($signature.Status -ne 'Valid' -or -not$signature.SignerCertificate -or $signature.SignerCertificate.Thumbprint -ne 'D70ECB7B00377EBFB76B304C08DFC6620584E114'){throw "Package signature verification failed: $($file.Name)"}
}
$manifestPath=Join-Path $PackagePath 'SmartM365.EndpointDiagnosticsAnalyzer.psd1';if(-not(Test-Path $manifestPath -PathType Leaf)){throw "Manifest not found: $manifestPath"};$manifest=Test-ModuleManifest $manifestPath;$pre=[string]$manifest.PrivateData.PSData['Prerelease'];$key=[Environment]::GetEnvironmentVariable($ApiKeyEnvironmentVariable);$preview=[pscustomobject]@{Mode=if($Execute){'Execute'}else{'Preview'};Repository=$Repository;ModuleName=$manifest.Name;Version=[string]$manifest.Version;Prerelease=$pre;PackagePath=(Resolve-Path $PackagePath).Path;ApiKeyEnvironmentVariable=$ApiKeyEnvironmentVariable;ApiKeyPresent=(-not[string]::IsNullOrWhiteSpace($key));PublicMetadataComplete=(-not[string]::IsNullOrWhiteSpace([string]$manifest.LicenseUri));PublicationAttempted=$false}
if(-not$Execute){return $preview};if($pre-and-not$AllowPrereleasePublication){throw "Prerelease approval is required: $pre"};if(-not$preview.PublicMetadataComplete){throw 'LicenseUri is required.'};if(-not$key){throw "API key environment variable is empty: $ApiKeyEnvironmentVariable"};if(-not(Get-Command Publish-PSResource -ErrorAction SilentlyContinue)){throw 'Publish-PSResource is unavailable.'};if($PSCmdlet.ShouldProcess("$($manifest.Name) $($manifest.Version) to $Repository",'Publish public PowerShell resource')){Publish-PSResource -Path $PackagePath -ApiKey $key -Repository $Repository -ErrorAction Stop;$preview.PublicationAttempted=$true};$preview

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCVjMW5guH1vDSV
# bomBkDMxdh/ZP8BA1/Jq2IuSSathoqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCtRd4XUk+CFsMXo7qqNGJN
# n1ucervgv1mhoeBU+encsDANBgkqhkiG9w0BAQEFAASCAYB9M5NtsE7nOxtdbeX8
# nexYRvBS6fRTSDdrgImnUNlyn7+LT3Svx6v096V0qHJdF1MKAJGOR4zwggNSmPF2
# FohyiaxFsF/q8c4BQRNmbpfTqwZ21d5VZrzso3XhFrye27pu0i51F8643gKQyLrs
# rwtv0TMlCM1w7VqINIjc7MEMFb/2JRSPmHufDYw+jB5EHtvC4gtqJTOXzU6L5UhL
# ZE9aB4E6O1MkMXccBF4/v+WwhgmGaDVpgco4oB1CoQuFt0fpU6vZomku8eU26BPF
# vH1qVex0XSBN5Vk32uvfDjmiFoGWX1Xd44oSMffhu9qj7psOGVKvYlcMeQP/GuBZ
# oeeNlB7b+hv0NjMuCIWdIhRGDpCK9vk3SC/HVWKpaK+L8aOkL5X8S+6+HHYhecRL
# NRL79WYi9aPmF6b6Pt1bXQhY1EpdCEJcNcXTcYf3kskT5KNBmc7p6Frf4XXpXNai
# eoDHscGN1iOE6L9e7YBrII4okf5JCxkcHUf5g41ZCBiFAAI=
# SIG # End signature block
