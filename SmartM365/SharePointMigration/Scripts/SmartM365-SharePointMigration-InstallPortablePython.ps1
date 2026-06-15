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
