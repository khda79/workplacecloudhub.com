<#
.SYNOPSIS
    Creates SHA256 integrity manifests for Smart Endpoint Migration Toolkit Windows 11 Upgrade setup media folders.

.DESCRIPTION
    Generates SmartM365-SetupMediaManifest.sha256.csv for one Windows setup media root
    or for each direct media subfolder under a setup source root. The file name is kept because the
    Windows 11 Upgrade endpoint, the Intune package builder and the Intune installer validate it.
    Files are hashed through long literal paths (\\?\ and \\?\UNC\) so deep media trees are supported.
    Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/Scripts/New-SmartM365SetupMediaManifest.ps1 1.0.1 (see Docs\PROVENANCE.json.txt).

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'MediaRoot')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'MediaRoot')]
    [string]$MediaRoot,

    [Parameter(Mandatory = $true, ParameterSetName = 'SetupSourceRoot')]
    [string]$SetupSourceRoot,

    [string]$ManifestFileName = 'SmartM365-SetupMediaManifest.sha256.csv',

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'
$CacheManifestFileName = 'SmartM365-SetupMedia.json.txt'

function ConvertTo-LongLiteralPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath.StartsWith('\\?\')) { return $fullPath }
    if ($fullPath.StartsWith('\\')) { return '\\?\UNC\' + $fullPath.Substring(2) }
    return '\\?\' + $fullPath
}

function Get-SmartFileSha256 {
    param([Parameter(Mandatory = $true)][string]$Path)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead((ConvertTo-LongLiteralPath -Path $Path))
        $hashBytes = $sha.ComputeHash($stream)
        return (($hashBytes | ForEach-Object { $_.ToString('X2') }) -join '')
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
        if ($null -ne $sha) { $sha.Dispose() }
    }
}
function Resolve-ExistingDirectory {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
    $item = Get-Item -LiteralPath $resolved.ProviderPath -ErrorAction Stop
    if (-not $item.PSIsContainer) {
        throw "Path is not a directory: $Path"
    }
    return $item.FullName
}

function Test-WindowsSetupMediaRoot {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$Quiet
    )

    $setupExe = Join-Path $Path 'setup.exe'
    $sources = Join-Path $Path 'sources'
    $installWim = Join-Path $Path 'sources\install.wim'
    $installEsd = Join-Path $Path 'sources\install.esd'

    $ok = (Test-Path -LiteralPath $setupExe -PathType Leaf) -and
        (Test-Path -LiteralPath $sources -PathType Container) -and
        ((Test-Path -LiteralPath $installWim -PathType Leaf) -or (Test-Path -LiteralPath $installEsd -PathType Leaf))

    if (-not $ok -and -not $Quiet) {
        throw "Not a Windows setup media root: $Path. Expected setup.exe and sources\install.wim or sources\install.esd."
    }

    return $ok
}

function Get-RelativePathFromRoot {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    $pathFull = [System.IO.Path]::GetFullPath($Path)
    if (-not $pathFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside root. Root=$Root; Path=$Path"
    }
    return $pathFull.Substring($rootFull.Length).Replace('/', '\')
}

function Get-MediaRootsFromSetupSourceRoot {
    param([Parameter(Mandatory = $true)][string]$Root)

    $roots = New-Object System.Collections.ArrayList
    if (Test-WindowsSetupMediaRoot -Path $Root -Quiet) {
        [void]$roots.Add($Root)
    }

    foreach ($child in @(Get-ChildItem -LiteralPath $Root -Directory -Force -ErrorAction Stop)) {
        if (Test-WindowsSetupMediaRoot -Path $child.FullName -Quiet) {
            [void]$roots.Add($child.FullName)
        }
    }

    $unique = @($roots.ToArray() | Select-Object -Unique)
    if ($unique.Count -eq 0) {
        throw "No Windows setup media roots found under: $Root"
    }
    return $unique
}

function New-SetupMediaManifest {
    param([Parameter(Mandatory = $true)][string]$Root)

    Test-WindowsSetupMediaRoot -Path $Root | Out-Null

    $manifestPath = Join-Path $Root $ManifestFileName
    if ((Test-Path -LiteralPath $manifestPath -PathType Leaf) -and -not $Force) {
        throw "Manifest already exists. Use -Force to overwrite: $manifestPath"
    }

    $tempPath = Join-Path $Root (".{0}.{1}.tmp" -f $ManifestFileName,[guid]::NewGuid().ToString('N'))
    $excludeNames = @($ManifestFileName, $CacheManifestFileName, 'SmartM365-SetupMedia.json')
    $files = @(
        Get-ChildItem -LiteralPath $Root -Recurse -File -Force -ErrorAction Stop |
            Where-Object { $_.Name -notin $excludeNames -and $_.FullName -ne $tempPath } |
            Sort-Object FullName
    )

    if ($files.Count -eq 0) {
        throw "No files found under media root: $Root"
    }

    Write-Host ("Creating setup media manifest: Root={0}; Files={1}; Output={2}" -f $Root,$files.Count,$manifestPath)
    $index = 0
    $rows = foreach ($file in $files) {
        $index++
        if ($index -eq 1 -or $index -eq $files.Count -or ($index % 25) -eq 0) {
            Write-Progress -Activity 'Hashing Windows setup media' -Status $file.FullName -PercentComplete ([int](($index / [math]::Max(1, $files.Count)) * 100))
        }
        [pscustomobject]@{
            RelativePath = Get-RelativePathFromRoot -Root $Root -Path $file.FullName
            Length = [int64]$file.Length
            SHA256 = Get-SmartFileSha256 -Path $file.FullName
            LastWriteUtc = $file.LastWriteTimeUtc.ToString('o')
        }
    }
    Write-Progress -Activity 'Hashing Windows setup media' -Completed

    try {
        $rows | Export-Csv -LiteralPath $tempPath -NoTypeInformation -Encoding UTF8 -Force
        Move-Item -LiteralPath $tempPath -Destination $manifestPath -Force
    }
    finally {
        if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }

    $totalBytes = [int64](($files | Measure-Object -Property Length -Sum).Sum)
    [pscustomobject]@{
        MediaRoot = $Root
        ManifestPath = $manifestPath
        FileCount = $files.Count
        TotalBytes = $totalBytes
        TotalGB = [math]::Round($totalBytes / 1GB, 2)
    }
}

Write-Host "Smart Endpoint Migration Toolkit - Windows 11 Upgrade setup media manifest generator v$ScriptVersion"

if ($PSCmdlet.ParameterSetName -eq 'SetupSourceRoot') {
    $root = Resolve-ExistingDirectory -Path $SetupSourceRoot
    $mediaRoots = @(Get-MediaRootsFromSetupSourceRoot -Root $root)
}
else {
    $mediaRoots = @(Resolve-ExistingDirectory -Path $MediaRoot)
}

$results = foreach ($root in $mediaRoots) {
    New-SetupMediaManifest -Root $root
}

$results | Format-Table -AutoSize

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA9kZIV+puz5X3/
# NiJRPMIh8D2Wj9sdcxduybmvQHYDGqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCX6X7Zf3hC1EEjGvE8xF8e
# BNkz5C5xkuex7v9OKLrkQzANBgkqhkiG9w0BAQEFAASCAYB2hpw2rQ8D5I3l+DH8
# R8b6YuxTjLV3iF837Ji7Z9QIMbkUriORq9A/jZDuC8vzCq1Qc05va8VG5xwNYt1Z
# VJFcRIPjvZa3qouiCFSnVV5b5egN8gjugqWuqIohjs70BqnW8BFp6Ux+gr1O1YRt
# uHYZ2Yuz5tHYZj/vO0PuNcbmvlUyylofeiZ6OmkslJwdj2503tlyKVbXSKgdqQ7c
# U7NJfdWsAjViaK8B96+SvmvZ+hPMusAL1YYyVaKVEY+0MkzsyNTdY2PmuYx8iBxa
# krjnGcV0/NsQNbDSGstc9Ki3tB+C4UNDT9m+gc0Gmn9ep/wOfwOPEtAjooAiNZA/
# qWC0/bY/Z7GZUYWXyQlZAUjjf8k75pWG2N3o1PKmUfEnn5QSxC2lrRMiCQWvtcRv
# h0/s2MirjF6K37ezW+kV1DGKTKzoRjAve6K4xRlPbgyOcSIYFWAJjX2GCCLqW1vX
# XDa8XcKSxsUfZQMl6FZlFDPz2xUZoJyZa8WO4CxuG2zIQiU=
# SIG # End signature block
