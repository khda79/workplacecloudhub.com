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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCjY8kqxsEu3QO5
# XDNHvORB3Sy7IDSI2MA2Je/gL+5RM6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB4f+JalY4TqXrNtNSj
# cOOlKvn8k6P0LfyAg2QNRT5ySzANBgkqhkiG9w0BAQEFAASCAYB2CnsId4Qgzhst
# mxw66TBOJy38dtlDrv/26rZai47BPi3R7AicJrukfYPiJs967duDHOTD/y3wJPn/
# IPqq5NrT+FT7vpjo5nPx6s1Q3DRK7B+mr56QnPOhqHbt5Q8wd7GhUp7se30aij+X
# Mx1URKV9GcInLk3M9s4hxYi5ZeL+roQpLubxCeG1uQ39nlcT8WCiSYD5G9PNxpTB
# HA/UW8mxuzcpg95ljZqeu0vTS1J7nn1GxxyWCB7jldtDTA5ksAHawsovRpzzPwrO
# n4jdFwg7SE3Daw/fx+hGp7x0oTAEBNAunrgAvvgMKoA/lcYjEoEkM1YDet2TNZtQ
# WFwOwBeUPM39BAcXh1tmoCE9kMMTAgCpNU7mnzJOVEd2dhjL0J6SsKVAGi1/9G/9
# wCKvWZhLdySpDyaCbZ4jMLJUCa+tHCRcZgcz4FVyA34BUtSbeP4ekKiUCvN9gd2x
# i7NdSvjWtijCnavKcEIGDx1MlXR0L1yhp/KJIC6N0IN46J8eNQk=
# SIG # End signature block
