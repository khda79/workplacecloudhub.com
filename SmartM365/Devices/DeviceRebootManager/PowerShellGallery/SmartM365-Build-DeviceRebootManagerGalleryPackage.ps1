#Requires -Version 5.1

<#
.SYNOPSIS
Builds the local PowerShell Gallery package for SmartM365 Device Reboot Manager.

.DESCRIPTION
Creates a clean module folder from an explicit allow-list, validates PowerShell
syntax and the pinned WorkplaceCloudHub Authenticode signer, and never copies
local runtime configuration or LOCAL_MEMORY.md.
.VERSION
1.0.1
#>

[CmdletBinding()]
param(
    [string]$SourceRoot = (Split-Path -Path $PSScriptRoot -Parent),
    [string]$OutputRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\PowerShellGallery'),
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [switch]$Force,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-NormalizedThumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Assert-SafeChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Child
    )

    $resolvedRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    $resolvedChild = [IO.Path]::GetFullPath($Child).TrimEnd('\', '/')
    $prefix = $resolvedRoot + [IO.Path]::DirectorySeparatorChar
    if ($resolvedChild -eq $resolvedRoot -or -not $resolvedChild.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe build target outside OutputRoot: $resolvedChild"
    }
}

function Assert-PowerShellSyntax {
    param([Parameter(Mandatory = $true)][string]$Path)

    $parseErrors = $null
    [void][Management.Automation.PSParser]::Tokenize(
        (Get-Content -LiteralPath $Path -Raw),
        [ref]$parseErrors
    )
    if ($parseErrors.Count -gt 0) {
        $messages = $parseErrors | ForEach-Object { $_.Message }
        throw ("PowerShell syntax validation failed for {0}: {1}" -f $Path,($messages -join '; '))
    }
}

function Assert-PinnedSignature {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    $actualThumbprint = if ($signature.SignerCertificate) {
        Get-NormalizedThumbprint $signature.SignerCertificate.Thumbprint
    }
    else {
        ''
    }

    if ($signature.Status -ne 'Valid' -or $actualThumbprint -ne (Get-NormalizedThumbprint $Thumbprint)) {
        throw ("Authenticode validation failed: file={0}; status={1}; signer={2}" -f $Path,$signature.Status,$actualThumbprint)
    }
}

$moduleTemplateRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Module'
$manifestSourcePath = Join-Path -Path $moduleTemplateRoot -ChildPath 'SmartM365.DeviceRebootManager.psd1'
if (-not (Test-Path -LiteralPath $manifestSourcePath -PathType Leaf)) {
    throw "Module manifest not found: $manifestSourcePath"
}

$manifest = Test-ModuleManifest -Path $manifestSourcePath
$moduleName = [string]$manifest.Name
$moduleVersion = [string]$manifest.Version
$prerelease = [string]$manifest.PrivateData.PSData['Prerelease']
$packageVersion = if ([string]::IsNullOrWhiteSpace($prerelease)) {
    $moduleVersion
}
else {
    '{0}-{1}' -f $moduleVersion,$prerelease
}

$versionManifestSourcePath = Join-Path -Path $SourceRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
if (-not (Test-Path -LiteralPath $versionManifestSourcePath -PathType Leaf)) {
    throw "Package version manifest not found: $versionManifestSourcePath"
}
$versionManifest = Get-Content -LiteralPath $versionManifestSourcePath -Raw | ConvertFrom-Json
if ([int]$versionManifest.SchemaVersion -ne 1 -or [string]$versionManifest.PackageVersion -ne $packageVersion) {
    throw ("Package version mismatch: manifest={0}; module={1}" -f $versionManifest.PackageVersion,$packageVersion)
}

$packagePath = Join-Path -Path (Join-Path -Path $OutputRoot -ChildPath $moduleName) -ChildPath $moduleVersion
Assert-SafeChildPath -Root $OutputRoot -Child $packagePath

if (Test-Path -LiteralPath $packagePath) {
    if (-not $Force) {
        throw "Build target already exists. Use -Force to replace it: $packagePath"
    }
    Remove-Item -LiteralPath $packagePath -Recurse -Force
}

New-Item -ItemType Directory -Path $packagePath -Force | Out-Null

$moduleFiles = @(
    'SmartM365.DeviceRebootManager.psd1'
    'SmartM365.DeviceRebootManager.psm1'
    'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
)

$runtimeFiles = @(
    'SmartM365-DeviceRebootManager.version.json.txt'
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-Demo.ps1'
    'Start-SmartM365-DeviceRebootManager-Demo.cmd'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json.template'
    'SmartM365.GuiSplash.ps1'
    'WorkplaceCloudHub.ico'
    'WorkplaceCloudHub-lockup-WPF.png'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
)

$deployFiles = @(
    'SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    'SmartM365-DeviceRebootManager-Detection.ps1'
    'SmartM365-DeviceRebootManager-Install.ps1'
    'SmartM365-DeviceRebootManager-Uninstall.ps1'
)

foreach ($relativePath in $moduleFiles) {
    $sourcePath = Join-Path -Path $moduleTemplateRoot -ChildPath $relativePath
    $destinationPath = Join-Path -Path $packagePath -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required module file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

foreach ($relativePath in $runtimeFiles) {
    $sourcePath = Join-Path -Path $SourceRoot -ChildPath $relativePath
    $destinationPath = Join-Path -Path (Join-Path $packagePath 'Runtime') -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required runtime file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

foreach ($relativePath in $deployFiles) {
    $sourcePath = Join-Path -Path (Join-Path $SourceRoot 'Deploy') -ChildPath $relativePath
    $destinationPath = Join-Path -Path (Join-Path $packagePath 'Runtime\Deploy') -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required deployment file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

$packagedFiles = @(Get-ChildItem -LiteralPath $packagePath -Recurse -File)
$powerShellFiles = @($packagedFiles | Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
foreach ($file in $powerShellFiles) {
    Assert-PowerShellSyntax -Path $file.FullName
    if (-not $SkipAuthenticodeValidation) {
        Assert-PinnedSignature -Path $file.FullName -Thumbprint $ExpectedSignerThumbprint
    }
}

$builtManifestPath = Join-Path -Path $packagePath -ChildPath 'SmartM365.DeviceRebootManager.psd1'
[void](Test-ModuleManifest -Path $builtManifestPath)

[pscustomobject]@{
    ModuleName               = $moduleName
    Version                  = $moduleVersion
    Prerelease               = $prerelease
    PackageVersion           = $packageVersion
    PackagePath              = $packagePath
    IntuneSourcePath         = Join-Path -Path $packagePath -ChildPath 'Runtime'
    FileCount                = $packagedFiles.Count
    PowerShellFileCount      = $powerShellFiles.Count
    SignaturesValidated      = (-not $SkipAuthenticodeValidation)
    ExpectedSignerThumbprint = Get-NormalizedThumbprint $ExpectedSignerThumbprint
    Ready                    = $true
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC3SwrkvxqNmOYK
# BwC/wT3n1b9CkmAxMA3OvEcDlD/qIaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDU0+MxiyrGpTAgDwvgUUa7UXjzmbPGZSYM0YJgs1l7UMA0GCSqG
# SIb3DQEBAQUABIIBgCmsaYrpde16qPLRxiHDm/nUQhEvrD72u1d8zALZkNkSER1i
# 8S01+lA4qHw0H2MUlQzZTlRsKJP2il2t8EVeosktBJ1d4LR1L0HLvd/okdsE6f/o
# 0G+HvunbG9gTsHd5NOjyRquaTwnlBapBmWUgBsdZ7icNoAKog9dDaXppyMS0AJ7f
# HkQQqeqqCvzPAufs6UwwcJb16ZkJsz/QgDcKC5E15Zc8Je7QJTEBRQxAXTXFZvcj
# Qp84INl3hXOVKE9z0FfYUPxlOutv1qO1a2t5RD90VnpMeDf2hyH0robD5XJIe/K8
# xRQlA5r0T7dazeJ/DgvH+GDjJo0THKqt2KV23thniG+dnd3pyQceQIdxzEW5g4W9
# 6hG65KOR/UDbfHXNeu5Gp364TYcjFlGZGEjiEu9QiAHOdUv98VjEgavIHIizDwzT
# cwIIngXvtxchYZvbQo1qnIJn68ZU1aho4QCNRnSmYomgmeH7uVUsbQS2+KOcOtlb
# N2U472QuzqvRKTz4S6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMzA5
# MzRaMC8GCSqGSIb3DQEJBDEiBCA0TTEx1bRXp/iPfVm8sWlzUWZUIsmouEyU4Nl1
# Vv9rmDANBgkqhkiG9w0BAQEFAASCAgCbVjbu6mV05WR6UkLZF47gjD2NNdRUVhfV
# bbvQad+YTyujKZTeMqBDtvRFdmv7IAx+bvv256kwDGRwVApr8EvLVVMlU2OVS6sR
# PVuWAPbabZPMqtn05NkzH/wxF6AO58a+uwGhUlkHaByb8JqP2N6m2yMTf/KraLV/
# C6oZdpcpNBVmNMILkbxkXKgSc4EiAQnYOk1l9xPB1npzAfkFsobTa72b3cujnxrn
# 02zDDLtR8EDtKTwSfIA7NVbDvYOs5g8ovkRC+BhChmaJnY9IwPJVieTjuAq6jq/D
# VtKR4hc3Lz9peAwR9UXpGqMYnU67/C2wpfkcTmAb96EGSb4Q/GfsuKJIfs3jSI1B
# lw/DEitPaqlfqtaTZe5JIAX0dp8f8KM/IMHeUnLUYAAmgBPhMyVuZUY4v+94avHj
# PFAsEs46+kJMC76bYIC4I63+JtJsBFL9N8bJ4KIribk8r+J7Mu7viI4ubP01A7Zz
# ynb3RKZaVZ+ea5VpOzJbWZvx9P4zcWRg2Rh8sD2/pYQVBVejIElre2xKvTO0a25v
# 0mYSOFTfEhd4Wz0DSRPFNOsuuvlxTMbaOlNHfqhwKwMavwBFeoXEYoxx0ZbbWDl+
# L2Ok6Qi883ewcmDYZbCwTaVMBQyh3fDTcOg6F69O6ZIk2VoOG7jaYhnac0ybOETr
# LOuyxppJVw==
# SIG # End signature block
