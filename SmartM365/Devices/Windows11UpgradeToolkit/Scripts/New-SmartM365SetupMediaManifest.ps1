<#
.SYNOPSIS
    Creates SHA256 integrity manifests for Windows setup media folders.

.DESCRIPTION
    Generates SmartM365-SetupMediaManifest.sha256.csv for one Windows setup media root
    or for each direct media subfolder under a setup source root.

.VERSION
    1.0.1

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
$ScriptVersion = '1.0.1'
$CacheManifestFileName = 'SmartM365-SetupMedia.json'

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
    $excludeNames = @($ManifestFileName, $CacheManifestFileName)
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

Write-Host "SmartM365 setup media manifest generator v$ScriptVersion"

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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCjY8kqxsEu3QO5
# XDNHvORB3Sy7IDSI2MA2Je/gL+5RM6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCB4f+JalY4TqXrNtNSjcOOlKvn8k6P0LfyAg2QNRT5ySzANBgkqhkiG9w0B
# AQEFAASCAYCxguxDkBBfbmKBofUAu43MyrAt8x+YkQmABsT/K3PD3PamV7wIenYf
# 3QSSoUoScbeNPaP30rE50WrEmHqQeAWo5bCxC623RnkTM8ztJVJjNNihSVy/DOG5
# e1/i5fJtyIzvqML2hDhX4aVidQlZRJxXXUVoSPDOONVaAXmu31TNo9EfWHxiScis
# MXa6lfTbUBKs8LgxKeiHV5R4ZpGnGf0C1sszePdLq8vsSb7wMhvrt475n19ykY22
# e98F2OSxdgHtZTymKCY/72Ortmq75WGeh2a+6MDLgTVdfOYcUpbIseVZkcGZoSTo
# CgFXZU1zYuh9ibNlBGOHo9qeBnfW5lt5GahSkE13EVYztpIIy0mP7ngl4NgTp631
# WZvS/v8+nD5ZZUcn10TdKUpsFEFHSc4NM82Hx0c5EOgTUqGB1J3gHYSiuIpbfPyQ
# xBvVOUVhjhADfdHF4FVw9h9Ph0Cg310PGVHjqAaqrl2FxAVftgECenVmc6tcANtu
# zb/YNicm/y0=
# SIG # End signature block
