<#
.SYNOPSIS
    Builds one SmartM365 Windows 11 Upgrade Toolkit Intune Win32 package per setup language.
.DESCRIPTION
    Creates a staging source folder containing the endpoint script, Intune installer scripts, package manifest, and optionally one language-specific Windows setup media cache, then optionally runs IntuneWinAppUtil.exe.
.VERSION
    1.0.2
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidatePattern('^[a-z]{2}-[A-Z]{2}$')][string]$Language,
    [string]$MediaFolder,
    [string]$MediaId = 'Win11',
    [string]$SetupSourceRoot,
    [string]$PackageVersion,
    [string]$IntuneWinAppUtilPath,
    [string]$ContentPrepRoot = 'C:\tmp\SmartM365-W11UT-ContentPrep',
    [switch]$WithCacheOnly,
    [switch]$SkipIntuneWinBuild,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$toolkitRoot = Split-Path -Parent $scriptDir
$sourceTemplateRoot = Join-Path $scriptDir 'Source'
$endpointScript = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-Windows11UpgradeRepair.ps1'
$integrityHelperSource = Join-Path $sourceTemplateRoot 'SmartM365-SetupMediaIntegrity.ps1'
if (-not (Test-Path -LiteralPath $integrityHelperSource -PathType Leaf)) { throw "Setup media integrity helper not found: $integrityHelperSource" }
. $integrityHelperSource

function Get-EndpointVersion {
    param([string]$Path)
    $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $match = [regex]::Match($content, '(?m)^\$script:ScriptVersion\s*=\s*''(?<Version>[^'']+)''')
    if (-not $match.Success) { throw "Unable to read endpoint script version from $Path" }
    return $match.Groups['Version'].Value
}

function Get-DefaultMediaFolder {
    param([string]$Value)
    $parts = $Value -split '-'
    if ($parts.Count -ne 2) { throw "Invalid language: $Value" }
    return ('{0}-{1}' -f $parts[0].ToUpperInvariant(), $parts[1].ToLowerInvariant())
}

if ([string]::IsNullOrWhiteSpace($SetupSourceRoot)) { $SetupSourceRoot = Join-Path $toolkitRoot 'SetupSource' }
if ([string]::IsNullOrWhiteSpace($MediaFolder)) { $MediaFolder = Get-DefaultMediaFolder -Value $Language }
$endpointVersion = Get-EndpointVersion -Path $endpointScript
if ([string]::IsNullOrWhiteSpace($PackageVersion)) { $PackageVersion = $endpointVersion }

$packageSuffix = if ($WithCacheOnly) { '-WithCacheOnly' } else { '' }
$packageId = "SmartM365-Windows11UpgradeToolkit-$MediaId-$Language$packageSuffix"
$displayName = "Windows11UpgradeToolkit-$Language$packageSuffix"
$cacheFolder = "$MediaId-$Language"
$packageMode = if ($WithCacheOnly) { 'WithCacheOnly' } else { 'WithMedia' }
$mediaRoot = Join-Path $SetupSourceRoot $MediaFolder

if (-not $WithCacheOnly) {
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'setup.exe') -PathType Leaf)) { throw "setup.exe not found in media root: $mediaRoot" }
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'sources\install.wim') -PathType Leaf)) { throw "sources\install.wim not found in media root: $mediaRoot" }
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'SmartM365-SetupMediaManifest.sha256.csv') -PathType Leaf)) { throw "SmartM365-SetupMediaManifest.sha256.csv not found in media root: $mediaRoot" }
    $sourceIntegrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $mediaRoot
    Write-Host ("Validated setup source integrity. Files={0}; Bytes={1}; Root={2}" -f $sourceIntegrity.Files,$sourceIntegrity.Bytes,$sourceIntegrity.MediaRoot)
}

$buildRoot = Join-Path $scriptDir ("Build\$packageId")
$packageSource = Join-Path $buildRoot 'Source'
$outputRoot = Join-Path $scriptDir ("Output\$packageId")
$mediaDest = Join-Path $packageSource ("SetupMedia\$cacheFolder")

if ((Test-Path -LiteralPath $buildRoot) -and -not $Force) { throw "Build folder already exists: $buildRoot. Use -Force to recreate it." }
if (Test-Path -LiteralPath $buildRoot) { Remove-Item -LiteralPath $buildRoot -Recurse -Force }
New-Item -ItemType Directory -Path $packageSource,$outputRoot -Force | Out-Null
if (-not $WithCacheOnly) { New-Item -ItemType Directory -Path $mediaDest -Force | Out-Null }

Copy-Item -LiteralPath (Join-Path $sourceTemplateRoot 'Install.ps1') -Destination (Join-Path $packageSource 'Install.ps1') -Force
Copy-Item -LiteralPath (Join-Path $sourceTemplateRoot 'Run-IntuneUpgrade.ps1') -Destination (Join-Path $packageSource 'Run-IntuneUpgrade.ps1') -Force
Copy-Item -LiteralPath $integrityHelperSource -Destination (Join-Path $packageSource 'SmartM365-SetupMediaIntegrity.ps1') -Force
Copy-Item -LiteralPath $endpointScript -Destination (Join-Path $packageSource 'SmartM365-Invoke-Windows11UpgradeRepair.ps1') -Force

$manifest = [ordered]@{
    PackageId = $packageId
    PackageVersion = $PackageVersion
    DisplayName = $displayName
    MediaId = $MediaId
    Language = $Language
    MediaFolder = $MediaFolder
    SetupCacheFolder = $cacheFolder
    PackageMode = $packageMode
    RequiresExistingSetupCache = [bool]$WithCacheOnly
    EndpointVersion = $endpointVersion
    CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
}
$manifestPath = Join-Path $packageSource 'PackageManifest.json'
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding ASCII
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $outputRoot 'PackageManifest.json') -Force

$robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
if (-not $WithCacheOnly) {
    Write-Host "Copying setup media: $mediaRoot -> $mediaDest"
    & $robocopy $mediaRoot $mediaDest /MIR /R:2 /W:5 /NP /NFL /NDL | Out-Null
    $copyExit = [int]$LASTEXITCODE
    if ($copyExit -gt 7) { throw "Robocopy media copy failed with exit code $copyExit." }
    $stagedIntegrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $mediaDest
    Write-Host ("Validated staged setup media integrity. Files={0}; Bytes={1}; Root={2}" -f $stagedIntegrity.Files,$stagedIntegrity.Bytes,$stagedIntegrity.MediaRoot)
}
else {
    Write-Host "Building cache-only package. Setup media is not embedded; endpoint will require local cache: $cacheFolder"
}

$detectTemplate = Get-Content -LiteralPath (Join-Path $sourceTemplateRoot 'Detect-Template.ps1') -Raw
$detect = $detectTemplate.Replace('__PACKAGE_ID__', $packageId).Replace('__PACKAGE_VERSION__', $PackageVersion)
$detect | Set-Content -LiteralPath (Join-Path $outputRoot 'Detect.ps1') -Encoding ASCII

$commands = @"
Install command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1

Uninstall command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Uninstall

Detection script:
$outputRoot\Detect.ps1
"@
$commands | Set-Content -LiteralPath (Join-Path $outputRoot 'Intune-App-Commands.txt') -Encoding ASCII

$intuneWinPath = ''
if (-not $SkipIntuneWinBuild) {
    if ([string]::IsNullOrWhiteSpace($IntuneWinAppUtilPath)) { throw 'IntuneWinAppUtilPath is required unless -SkipIntuneWinBuild is used.' }
    if (-not (Test-Path -LiteralPath $IntuneWinAppUtilPath -PathType Leaf)) { throw "IntuneWinAppUtil.exe not found: $IntuneWinAppUtilPath" }

    $prepRoot = Join-Path $ContentPrepRoot $packageId
    $prepSource = Join-Path $prepRoot 'Source'
    $prepOutput = Join-Path $prepRoot 'Output'
    if (Test-Path -LiteralPath $prepRoot) { Remove-Item -LiteralPath $prepRoot -Recurse -Force }
    New-Item -ItemType Directory -Path $prepSource,$prepOutput -Force | Out-Null

    Write-Host "Preparing short IntuneWinAppUtil source path: $prepSource"
    & $robocopy $packageSource $prepSource /MIR /R:2 /W:5 /NP /NFL /NDL | Out-Null
    $prepCopyExit = [int]$LASTEXITCODE
    if ($prepCopyExit -gt 7) { throw "Robocopy content-prep copy failed with exit code $prepCopyExit." }
    if (-not $WithCacheOnly) {
        $prepMediaRoot = Join-Path $prepSource ("SetupMedia\{0}" -f $cacheFolder)
        $prepIntegrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $prepMediaRoot
        Write-Host ("Validated content-prep setup media integrity. Files={0}; Bytes={1}; Root={2}" -f $prepIntegrity.Files,$prepIntegrity.Bytes,$prepIntegrity.MediaRoot)
    }

    & $IntuneWinAppUtilPath -c $prepSource -s 'Install.ps1' -o $prepOutput -q
    if ($LASTEXITCODE -ne 0) { throw "IntuneWinAppUtil failed with exit code $LASTEXITCODE." }

    $generated = Join-Path $prepOutput 'Install.intunewin'
    if (-not (Test-Path -LiteralPath $generated -PathType Leaf)) { throw "Expected IntuneWinAppUtil output not found: $generated" }
    $intuneWinPath = Join-Path $outputRoot ("{0}.intunewin" -f $packageId)
    Copy-Item -LiteralPath $generated -Destination $intuneWinPath -Force
}

[pscustomobject]@{
    PackageId = $packageId
    DisplayName = $displayName
    PackageMode = $packageMode
    Source = $packageSource
    Output = $outputRoot
    IntuneWin = $intuneWinPath
    DetectionScript = Join-Path $outputRoot 'Detect.ps1'
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBCmzvbwSO3J1OL
# WyFlInosMdtHtpaUvnksDCHA+5yGR6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIOlKZCVNZViV/CYE/a62Mq0TNV4DoDY/lh5dSJzi2ZwiMA0GCSqG
# SIb3DQEBAQUABIIBgFOQ6Hj5jVlnwABD/R1KnsQdZPpuqiVosMeC6JvvYR/IC6xb
# Sayi58JlKI4FZ3PiiN7s5llMnt2tFqugK9+U6t0XofALb3cci2/+pax5G3oOXmaJ
# mGse6z30K2BQLkt42mv0IRckEboJgndezX9QY5TsNXxvHa98AC+c2G3YCuEDXvlP
# nhpC07O1twRvPSlXPaEzWMYc6hjMl8ejCfVycyvzKxkU4sj6cvfMGoK7jZG0WOxw
# qOCkVq4/m+iqbVuS5P2dZDTWKvzGGGcuGNsDiDPcIoMo+ZAfrH63BZsxA5KWnRhJ
# zyzGR7E+Je6iSX1zWq2/VwCoUMF14y8XqMSF/QWOilhTFKgmdKdeNaXPfeaafaKH
# zYDgE/LEuKjialdssMYwzZ6kE4OiYeQWM9zP3ztTecwJXCY0ZXq4e5a89G7CICqq
# +ZT6xPpSXMDJYPDFj3bsG8TdCutKCR3jKFzIpi/GvNB55gIdBh4NRb1XEuRWd6z8
# 7sSEiaYeOw/aOYUyEKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MzhaMC8GCSqGSIb3DQEJBDEiBCDYOkZm78bjqXvfRVfS3LLdR/VY2Z/iYsHDxDm3
# uDApJDANBgkqhkiG9w0BAQEFAASCAgCRmG0DsRsCetq0K0jAejzQKezh/NjbAq9t
# gXxeEiaY6MElfdroDkX8yK21ez4CCYpLW7N+bw8xEApKzm/7K8gOkoqs/2VsYs8u
# N22Tc6uCIAwstybGHGk4qwOMLcAuPXvYRiPUEK+4rJJWt1C7otJkIdEhG3IHxmVN
# ZCyDtBvzDIb74NNq5nQMsWvC81BKPcA17yUGr5LOrZQp/9+FcPYHYSVtRkWM7g9r
# 07/lOTGv47kFvcoYjumZPs/jaMItyIKUrvd55iNBKKVI+ziZhd0nuBl2rwKVOlb5
# SA1iqBu2NOPDz+GvOBTQA8YX2W7XqI9Nw7HGl48P5OVNjKHzwY6JJPs+VVX5AfhF
# lwUPyW/ClzaWzGwMyvKisWgKoeUdzuOvoQDt2DmuMaKiAX1PlounzwLWgMNVzmtu
# shzm46l0fH1tTzkyqG7a/f4Hx+RawjrtxtAlXG8ICuRVbT4TATHiOooX2QavON82
# vUGpTrqeXx3EflOe4fSgwIfNg08mdGWktWhKyvqZk1FhyDe+DyrSp44xHHE+PBTD
# OHnM8TkjWAXoLqal80xegHqVQnRii5GFCi0Ggn267KLD0Y1BkRbV2Dkpbo/rIAFB
# Q/CN45p27AvK/uSHlRzaY36p3x5rrmhXLeolSNTQdywZTzdTLtJK3spQhO12Ieow
# jD7H9e/rIw==
# SIG # End signature block
