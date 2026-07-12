<#
.SYNOPSIS
    Installs the portable Python runtime used by SharePointMigration helpers.

.DESCRIPTION
    Downloads the official Windows embeddable Python package, verifies its
    SHA256 hash when provided, extracts it to Tools\Python, and writes a short
    runtime manifest. The install is local to this toolkit and does not modify
    PATH, registry keys, file associations, or machine-wide settings.

.EXAMPLE
    .\Scripts\SmartM365-SharePointMigration-InstallPortablePython.ps1

.EXAMPLE
    .\Scripts\SmartM365-SharePointMigration-InstallPortablePython.ps1 -Force

.EXAMPLE
    .\Scripts\SmartM365-SharePointMigration-InstallPortablePython.ps1 -PackagePath .\Tools\python-3.13.13-embed-amd64.zip -Force
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateNotNullOrEmpty()]
    [string]$PythonVersion = '3.13.13',

    [ValidateSet('amd64')]
    [string]$Architecture = 'amd64',

    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath = (Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'Tools\Python'),

    [string]$ExpectedSha256 = '8766A8775746235E23CF5AEE5027AB1060BB981D93110577ADCF3508AA0CBD55',

    [ValidateNotNullOrEmpty()]
    [string]$PackagePath,

    [switch]$Force,

    [switch]$KeepArchive
)

$ErrorActionPreference = 'Stop'

function Write-Step {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

$packageName = "python-$PythonVersion-embed-$Architecture.zip"
$downloadUrl = "https://www.python.org/ftp/python/$PythonVersion/$packageName"
$downloadDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'SmartM365-SharePointMigration'
$archivePath = Join-Path -Path $downloadDirectory -ChildPath $packageName
$stagingPath = Join-Path -Path (Split-Path -Parent $DestinationPath) -ChildPath ("Python.staging.{0}" -f ([guid]::NewGuid().ToString('N')))
$pythonExe = Join-Path -Path $DestinationPath -ChildPath 'python.exe'

if ((Test-Path -LiteralPath $pythonExe -PathType Leaf) -and -not $Force) {
    Write-Step "Portable Python already exists: $pythonExe"
    Write-Step "Use -Force to replace it."
    exit 0
}

New-Item -ItemType Directory -Path $downloadDirectory -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $DestinationPath) -Force | Out-Null

try {
    if (-not [string]::IsNullOrWhiteSpace($PackagePath)) {
        $resolvedPackagePath = (Resolve-Path -LiteralPath $PackagePath).ProviderPath
        $archivePath = $resolvedPackagePath
        Write-Step "Using local package: $archivePath"
    }
    else {
        Write-Step "Download: $downloadUrl"
        Invoke-WebRequest -Uri $downloadUrl -OutFile $archivePath -UseBasicParsing
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedSha256)) {
        $actualSha256 = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash
        if ($actualSha256 -ne $ExpectedSha256) {
            throw "SHA256 mismatch for $packageName. Expected $ExpectedSha256 but got $actualSha256."
        }

        Write-Step "SHA256 verified: $actualSha256"
    }

    if (Test-Path -LiteralPath $stagingPath) {
        Remove-Item -LiteralPath $stagingPath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $stagingPath -Force | Out-Null

    Write-Step "Extracting to staging folder."
    Expand-Archive -LiteralPath $archivePath -DestinationPath $stagingPath -Force

    if (-not (Test-Path -LiteralPath (Join-Path -Path $stagingPath -ChildPath 'python.exe') -PathType Leaf)) {
        throw "The downloaded package did not contain python.exe."
    }

    if (Test-Path -LiteralPath $DestinationPath) {
        if (-not $Force) {
            throw "Destination already exists: $DestinationPath. Use -Force to replace it."
        }

        if ($PSCmdlet.ShouldProcess($DestinationPath, 'Replace portable Python runtime')) {
            Remove-Item -LiteralPath $DestinationPath -Recurse -Force
        }
    }

    if ($PSCmdlet.ShouldProcess($DestinationPath, 'Install portable Python runtime')) {
        Move-Item -LiteralPath $stagingPath -Destination $DestinationPath

        $manifestPath = Join-Path -Path $DestinationPath -ChildPath 'PYTHON-PORTABLE.txt'
        $manifest = @(
            'Python portable runtime for SmartM365 SharePointMigration.'
            ''
            "Version: Python $PythonVersion embeddable package for Windows x64"
            "Source: https://www.python.org/downloads/release/python-$($PythonVersion.Replace('.', ''))/"
            "Package: $packageName"
            "SHA256: $ExpectedSha256"
            ''
            'This runtime is used by the comparison launchers before any system Python.'
            'It does not modify PATH, registry keys, file associations, or server-wide settings.'
        )
        Set-Content -LiteralPath $manifestPath -Value $manifest -Encoding UTF8

        $installedPython = Join-Path -Path $DestinationPath -ChildPath 'python.exe'
        $versionOutput = & $installedPython --version
        Write-Step "Installed: $installedPython"
        Write-Step "Runtime check: $versionOutput"
    }
}
finally {
    if (Test-Path -LiteralPath $stagingPath) {
        Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ([string]::IsNullOrWhiteSpace($PackagePath) -and -not $KeepArchive -and (Test-Path -LiteralPath $archivePath)) {
        Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCClh/c5JOZMGs6X
# zRRkkXQ0XWKgCx0IU1/kAn7jrRN6bqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCB4JC3z2+sh0Kgjw5T8vbPOpu8uv/hTlHYSf9PnVSknKDANBgkqhkiG9w0B
# AQEFAASCAYCvUNk/tTw9jcRIRgvSGsDofr2pOBiYPZkn2sjOhuWcRL5UYmLqrzOJ
# 8V4SBkfcTKIp+P5KjYv0vxeEKZDDvOe7Cq0bpqCwqD3nDOB6KT0V8Ay2dCPNMcYZ
# ncp83C8nk/0ajVxC7PxcYLuJdneaHig8SJTVfG2Uui9+F+Aje1YMpVgGmhowBo5G
# OgsmGdCvHqDQr9ckz8VaO6vjNkziW+WAatzn6ptq1qCpw1z1BSDJJWU8A5F65WDQ
# A00Ji+xYUJtjaheu/D6Bt1JVXAbYbV6hEAfpb5/kccEB9q+sv3E1V0OVzNhNcSLR
# o1EfWqjzk+MjI5bCJbu4Zm7rOBK1kNc3feJxwlyvIvqIxuLDOmF31ZbGHl984V/G
# ShJf36fAfYYYn59WGv5hAjb6R2vldXQi2sbpNV/XlTRZGE9lz3MhDaIhTJU+lzmK
# IYN4LPQ0rzqDQhxewQbjfD2LcSjukbnMlC7uXyfsPTEeujfwUywDUHs3c10pYB1S
# Ft4TFphM2/8=
# SIG # End signature block
