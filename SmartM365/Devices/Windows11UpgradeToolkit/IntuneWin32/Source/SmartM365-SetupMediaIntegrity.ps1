<#
.SYNOPSIS
    Validates a Windows setup media folder against its SmartM365 SHA256 manifest.
.VERSION
    1.0.0
#>

#requires -Version 5.1

function ConvertTo-SmartM365LongLiteralPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath.StartsWith('\\?\')) { return $fullPath }
    if ($fullPath.StartsWith('\\')) { return '\\?\UNC\' + $fullPath.Substring(2) }
    return '\\?\' + $fullPath
}

function Get-SmartM365SetupMediaFileHash {
    param([Parameter(Mandatory = $true)][string]$Path)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead((ConvertTo-SmartM365LongLiteralPath -Path $Path))
        $hashBytes = $sha.ComputeHash($stream)
        return (($hashBytes | ForEach-Object { $_.ToString('X2') }) -join '')
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
        if ($null -ne $sha) { $sha.Dispose() }
    }
}

function Test-SmartM365SetupMediaIntegrity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$MediaRoot,
        [string]$ManifestFileName = 'SmartM365-SetupMediaManifest.sha256.csv'
    )

    if (-not (Test-Path -LiteralPath $MediaRoot -PathType Container)) {
        throw "Setup media root not found: $MediaRoot"
    }

    $manifestPath = Join-Path $MediaRoot $ManifestFileName
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "Setup media integrity manifest not found: $manifestPath"
    }

    try {
        $rootFull = [System.IO.Path]::GetFullPath($MediaRoot).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
        $rows = @(Import-Csv -LiteralPath $manifestPath -ErrorAction Stop)
    }
    catch {
        throw ("Setup media integrity manifest cannot be read. Manifest={0}; Error={1}" -f $manifestPath,$_.Exception.Message)
    }

    if ($rows.Count -eq 0) { throw "Setup media integrity manifest is empty: $manifestPath" }
    foreach ($column in @('RelativePath','Length','SHA256')) {
        if (-not $rows[0].PSObject.Properties[$column]) {
            throw "Setup media integrity manifest is missing required column '$column': $manifestPath"
        }
    }

    $checkedFiles = 0
    $checkedBytes = 0L
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $relativePath = ([string]$row.RelativePath).Trim().TrimStart('\', '/')
        if ([string]::IsNullOrWhiteSpace($relativePath)) {
            throw ("Setup media integrity manifest contains an empty RelativePath. Manifest={0}; Row={1}" -f $manifestPath,$rowNumber)
        }
        if ([System.IO.Path]::IsPathRooted($relativePath) -or $relativePath -match '(^|[\\/])\.\.([\\/]|$)') {
            throw ("Setup media integrity manifest contains an unsafe RelativePath. Manifest={0}; Row={1}; RelativePath={2}" -f $manifestPath,$rowNumber,$relativePath)
        }
        if ($relativePath -match '[<>:"|?*\x00-\x1F]') {
            throw ("Setup media integrity manifest contains an invalid RelativePath character. Manifest={0}; Row={1}; RelativePath={2}" -f $manifestPath,$rowNumber,$relativePath)
        }

        try {
            $fileFull = [System.IO.Path]::GetFullPath((Join-Path $MediaRoot $relativePath))
        }
        catch {
            throw ("Setup media integrity manifest path cannot be resolved. Manifest={0}; Row={1}; RelativePath={2}; Error={3}" -f $manifestPath,$rowNumber,$relativePath,$_.Exception.Message)
        }
        if (-not $fileFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw ("Setup media integrity manifest path escapes media root. Manifest={0}; Row={1}; RelativePath={2}; File={3}" -f $manifestPath,$rowNumber,$relativePath,$fileFull)
        }
        if (-not (Test-Path -LiteralPath $fileFull -PathType Leaf)) {
            throw ("Setup media integrity check failed. File missing. Manifest={0}; Row={1}; RelativePath={2}; File={3}" -f $manifestPath,$rowNumber,$relativePath,$fileFull)
        }

        $expectedLength = 0L
        if (-not [int64]::TryParse([string]$row.Length, [ref]$expectedLength)) {
            throw ("Setup media integrity manifest contains an invalid Length. Manifest={0}; Row={1}; RelativePath={2}; Length={3}" -f $manifestPath,$rowNumber,$relativePath,$row.Length)
        }
        $item = Get-Item -LiteralPath $fileFull -ErrorAction Stop
        if ([int64]$item.Length -ne $expectedLength) {
            throw ("Setup media integrity check failed. Length mismatch. Manifest={0}; Row={1}; RelativePath={2}; File={3}; ExpectedLength={4}; ActualLength={5}" -f $manifestPath,$rowNumber,$relativePath,$fileFull,$expectedLength,$item.Length)
        }

        $expectedHash = ([string]$row.SHA256).Trim().ToUpperInvariant()
        if ($expectedHash -notmatch '^[A-F0-9]{64}$') {
            throw ("Setup media integrity manifest contains an invalid SHA256. Manifest={0}; Row={1}; RelativePath={2}; SHA256={3}" -f $manifestPath,$rowNumber,$relativePath,$row.SHA256)
        }
        try { $actualHash = (Get-SmartM365SetupMediaFileHash -Path $fileFull).ToUpperInvariant() }
        catch {
            throw ("Setup media integrity check failed. Cannot hash file. Manifest={0}; Row={1}; RelativePath={2}; File={3}; Error={4}" -f $manifestPath,$rowNumber,$relativePath,$fileFull,$_.Exception.Message)
        }
        if ($actualHash -ne $expectedHash) {
            throw ("Setup media integrity check failed. SHA256 mismatch. Manifest={0}; Row={1}; RelativePath={2}; File={3}; ExpectedSHA256={4}; ActualSHA256={5}" -f $manifestPath,$rowNumber,$relativePath,$fileFull,$expectedHash,$actualHash)
        }

        $checkedFiles++
        $checkedBytes += [int64]$item.Length
    }

    return [pscustomobject]@{
        MediaRoot = [System.IO.Path]::GetFullPath($MediaRoot)
        ManifestPath = $manifestPath
        Files = $checkedFiles
        Bytes = $checkedBytes
    }
}