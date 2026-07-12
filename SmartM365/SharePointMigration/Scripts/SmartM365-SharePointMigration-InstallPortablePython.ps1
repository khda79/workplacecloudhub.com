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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCClh/c5JOZMGs6X
# zRRkkXQ0XWKgCx0IU1/kAn7jrRN6bqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB4JC3z2+sh0Kgjw5T8
# vbPOpu8uv/hTlHYSf9PnVSknKDANBgkqhkiG9w0BAQEFAASCAYB7PBrI3zhIZJs8
# 4Hsf4wsWAZZeo/Famrazs5XKDSNtmBx9v47kJDpdFSUK6lHTUjo09KWNqB+u3wuD
# FB28BELa9brmnVjVM9P/0/jeKEHWR4UIf/M6oJTglqESyOJ9fDjpf3OHGzSq2Tl3
# AyMaxep2MD6wb4wBKRBQNQ2dyJz7mQ3q5Tbydld68k9UXzGPWlg7z3sB0My06KqP
# hdc5CXIA7jjMqs6pro6EaGrXBg+yzsU3AYslP3bRUNiQbad6P+e77rprsc7pXZXK
# ooM5k+uHq9iepWLVwqqc0K/aqVpjIOqbfQg7dwLZdb/E0lg/+7wvXnpdw0wSJQx6
# bCE8P/4TrdGGO9zbjN3aTKWjMJzz031ctEtVTYGe6h0VY+9erlKqMTknQh4rVK4d
# qbzkb6J4g8hS8iEG3BWdtk5kG2bLCDYjwOm45Z3BEDlRWCri5hHT4Ss2ydUB+GGd
# 20r9K7UHQQqic39zlRs7s/8QKxRS4WjCdKDhgFp5Cp9uPNQ4JG0=
# SIG # End signature block
