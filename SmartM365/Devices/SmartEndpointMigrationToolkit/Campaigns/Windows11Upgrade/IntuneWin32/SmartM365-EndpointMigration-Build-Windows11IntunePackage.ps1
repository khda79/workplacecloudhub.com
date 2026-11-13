<#
.SYNOPSIS
    Builds one Smart Endpoint Migration Toolkit Windows 11 Upgrade Intune Win32 package per setup language.
.DESCRIPTION
    Creates a staging source folder containing the endpoint script, Intune installer scripts, package manifest, and optionally one language-specific Windows setup media cache, then optionally runs IntuneWinAppUtil.exe.
    Package ids use the SmartM365-SEMT-Windows11Upgrade- prefix and app names use 'Smart Endpoint Migration Toolkit - Windows 11 Upgrade (<language>)' so they never collide with the previous Windows11UpgradeToolkit packages.
    Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/IntuneWin32/Build-SmartM365Windows11IntunePackage.ps1 1.0.2 (see Docs\PROVENANCE.json.txt).
.VERSION
    1.0.1
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
    [string]$ContentPrepRoot = 'C:\tmp\SmartM365-SEMT-W11-ContentPrep',
    [switch]$WithCacheOnly,
    [switch]$AllowSetupMediaDuplication,
    [switch]$SkipIntuneWinBuild,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Test-WchDemoAssetPath {
    param([string]$Path)
    $name=$Path.Replace('\','/')
    return $name -match '(?i)(^|/)DemoAssets(/|$)|(^|/)[^/]*demo[-_ ]?assets[^/]*\.(zip|nupkg|7z|rar|tar|tgz|gz|intunewin)$'
}

function Assert-WchNoDemoAssets {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string[]]$Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $budget=@{Entries=0;Bytes=[long]0}
    function Read-WchReleaseArchive {
        param([IO.Stream]$Stream,[string]$Name,[int]$Depth)
        if($Depth -gt 8){throw "Release archive nesting limit exceeded: $Name"}
        $zip=[IO.Compression.ZipArchive]::new($Stream,[IO.Compression.ZipArchiveMode]::Read,$true)
        try {
            foreach($entry in $zip.Entries){
                $budget.Entries++;$budget.Bytes+=$entry.Length
                if($budget.Entries -gt 100000 -or $budget.Bytes -gt 536870912){throw "Release archive inspection limit exceeded: $Name"}
                if(Test-WchDemoAssetPath $entry.FullName){throw "DemoAssets are local-only: $Name!$($entry.FullName)"}
                if($entry.Name){
                    $content=$entry.Open();$memory=$null
                    try {
                        $head=New-Object byte[] 4;$read=$content.Read($head,0,4)
                        $magic=$read -eq 4 -and $head[0] -eq 80 -and $head[1] -eq 75 -and (($head[2] -eq 3 -and $head[3] -eq 4) -or ($head[2] -eq 5 -and $head[3] -eq 6) -or ($head[2] -eq 7 -and $head[3] -eq 8))
                        if($magic -or $entry.FullName -match '(?i)\.(zip|nupkg)$'){
                            $memory=[IO.MemoryStream]::new();$memory.Write($head,0,$read);$content.CopyTo($memory);$memory.Position=0
                            Read-WchReleaseArchive $memory ($Name+'!'+$entry.FullName) ($Depth+1)
                        }
                    }finally{$content.Dispose();if($memory){$memory.Dispose()}}
                }
            }
        }finally{$zip.Dispose()}
    }
    foreach($inputPath in $Path){
        if(Test-WchDemoAssetPath $inputPath){throw "DemoAssets are local-only: $inputPath"}
        $item=Get-Item -LiteralPath $inputPath -Force -ErrorAction Stop
        $pending=[Collections.Generic.Stack[IO.FileSystemInfo]]::new();$pending.Push($item)
        while($pending.Count){
            $file=$pending.Pop()
            if(Test-WchDemoAssetPath $file.FullName){throw "DemoAssets are local-only: $($file.FullName)"}
            if($file.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "Release input must not be a reparse point: $($file.FullName)"}
            if($file.PSIsContainer){foreach($child in @(Get-ChildItem -LiteralPath $file.FullName -Force -ErrorAction Stop)){$pending.Push($child)}}else{
                $stream=[IO.File]::OpenRead($file.FullName)
                try {
                    $head=New-Object byte[] 4;$read=$stream.Read($head,0,4);$stream.Position=0
                    $magic=$read -eq 4 -and $head[0] -eq 80 -and $head[1] -eq 75 -and (($head[2] -eq 3 -and $head[3] -eq 4) -or ($head[2] -eq 5 -and $head[3] -eq 6) -or ($head[2] -eq 7 -and $head[3] -eq 8))
                    if($magic -or $file.Extension -match '(?i)^\.(zip|nupkg)$'){Read-WchReleaseArchive $stream $file.FullName 0}
                }finally{$stream.Dispose()}
            }
        }
    }
}


$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$toolkitRoot = Split-Path -Parent $scriptDir
$sourceTemplateRoot = Join-Path $scriptDir 'Source'
$endpointScriptFileName = 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$installerFileName = 'SmartM365-EndpointMigration-Install.ps1'
$runnerFileName = 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1'
$integrityHelperFileName = 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'
$detectTemplateFileName = 'SmartM365-EndpointMigration-Detect-Template.ps1'
$detectOutputFileName = 'SmartM365-EndpointMigration-Detect.ps1'
$endpointScript = Join-Path $toolkitRoot ('Scripts\{0}' -f $endpointScriptFileName)
$integrityHelperSource = Join-Path $sourceTemplateRoot $integrityHelperFileName
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
$packageId = "SmartM365-SEMT-Windows11Upgrade-$MediaId-$Language$packageSuffix"
$displayName = if ($WithCacheOnly) { "Smart Endpoint Migration Toolkit - Windows 11 Upgrade ($Language, WithCacheOnly)" } else { "Smart Endpoint Migration Toolkit - Windows 11 Upgrade ($Language)" }
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

Copy-Item -LiteralPath (Join-Path $sourceTemplateRoot $installerFileName) -Destination (Join-Path $packageSource $installerFileName) -Force
Copy-Item -LiteralPath (Join-Path $sourceTemplateRoot $runnerFileName) -Destination (Join-Path $packageSource $runnerFileName) -Force
Copy-Item -LiteralPath $integrityHelperSource -Destination (Join-Path $packageSource $integrityHelperFileName) -Force
Copy-Item -LiteralPath $endpointScript -Destination (Join-Path $packageSource $endpointScriptFileName) -Force

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
    AllowSetupMediaDuplication = [bool]$AllowSetupMediaDuplication
    EndpointVersion = $endpointVersion
    CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
}
$manifestPath = Join-Path $packageSource 'PackageManifest.json.txt'
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding ASCII
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $outputRoot 'PackageManifest.json.txt') -Force

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

$detectTemplate = Get-Content -LiteralPath (Join-Path $sourceTemplateRoot $detectTemplateFileName) -Raw
foreach ($placeholder in @('__PACKAGE_ID__','__PACKAGE_VERSION__')) {
    if (-not $detectTemplate.Contains($placeholder)) { throw "Detection template placeholder $placeholder not found: $detectTemplateFileName" }
}
$detect = $detectTemplate.Replace('__PACKAGE_ID__', $packageId).Replace('__PACKAGE_VERSION__', $PackageVersion)
if ($detect.Contains('__PACKAGE_')) { throw "Generated detection script still contains a placeholder: $detectOutputFileName" }
$detectOutputPath = Join-Path $outputRoot $detectOutputFileName
$detect | Set-Content -LiteralPath $detectOutputPath -Encoding ASCII

$commands = @"
Install command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\$installerFileName

Uninstall command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\$installerFileName -Uninstall

Detection script:
$detectOutputPath
"@
$commands | Set-Content -LiteralPath (Join-Path $outputRoot 'Intune-App-Commands.txt') -Encoding ASCII

Assert-WchNoDemoAssets -Path $packageSource
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

    Assert-WchNoDemoAssets -Path $prepSource
    & $IntuneWinAppUtilPath -c $prepSource -s $installerFileName -o $prepOutput -q
    if ($LASTEXITCODE -ne 0) { throw "IntuneWinAppUtil failed with exit code $LASTEXITCODE." }

    $generated = Join-Path $prepOutput ('{0}.intunewin' -f [System.IO.Path]::GetFileNameWithoutExtension($installerFileName))
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
    DetectionScript = $detectOutputPath
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCB0lu6ielcw2pn
# oC95LkqOhm3FmImA+hM6xkOFmcfMpqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDJOBmTRWlB5kUjWMhwakB78uvq9Rx1R9Vvzu8up8vl0MA0GCSqG
# SIb3DQEBAQUABIIBgHqGuMDGuJMlIibndHLhrwZnlkx3UEfCalRK9imSOM91KuQP
# xoDgu/s+DrYpRvT+KFNrXVF4SRvgYP8qpJ9JFByCr/+7D25BPAUcWZ/ea29JmiRv
# 8peMuMGou4G5oEX2wLtLMobagkroQt6zAPybjIqvGzHRXKHLr8k8r6fitbvluWCI
# 1Oh701Tk/aogN8egB6st3LGUVKAXAzTFCB6zptrYitde8bR2fKizSPxGQXTJ0lyu
# f25AygOCBL4UB78XyvMquDPhH4VyJk9Li7B90OfnoeWIhzQKXEQqCIMStgNLrjB8
# F5NutGOMA2zgEvF2BmYS+ylODHPewRXdKvQUX0s71qbQGHZocJkjQyHyHT0fJhy7
# QDNogjS6OAL6TjOVF9V8wM2jVsG+s/76tZ07uvvZ8sYa3gD/6tbclH7RL1aV0L9G
# DJgmr8BQIhpquBlihzCyKqwLRSaVWHrCEp+0IGlzvzxkaMBVZIO9k3xVocXS/M4S
# f5an0jJtrFfAyySoyKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMDE1
# NTJaMC8GCSqGSIb3DQEJBDEiBCAdR0lIDkakfVx+JmyZTJ6OZZ5euQdWUPs1r3YO
# lU2MPzANBgkqhkiG9w0BAQEFAASCAgAGgiHoQnSXO/IJTFwTqXkRPjHJySPjv/dX
# iMY4wUqXDsgEt3dKfNCGhE1wnm+DZXyc2K5YmcalnLcVaS0bMQQrxqNO/+SnSY14
# vGIH9o4IhSkOR/yQAmk50UUu90rG+ARcu7oUstU5vF8IMW+5ksYW2EFeACXHr+go
# Z/FYOZR3NyzWsuAJOefKmV6BPUhY7u1zi5tMACxms150nhPY6Uj68Rjk5jJyTK+A
# 1L9SbxIWolg4u44h62BVZnofkvQm1Ou1wB4SlG2pDzMGx2KgBat2/DfS4naK63gp
# tggQQM9MBX5C80R3koOZaDr0zZAsHsjMaNgRI3bOLBjGHRMvEzbr31rocGjXG+ga
# VbO7vSUnVOzI3OsSRpYBuH88WlAE3NeCyAjlfNUewhC1mTnn1pX9VAapr5k11c28
# xKxEiKiSaVTI9Pb4bF9+TuDCHborPJ3ZSMCIt0Y9ycNkMWHDZpt6DuHzQ9VT1AqG
# h+H8ynfXe5ZFNIWUIV24T80nB0Eup5NyvG+MV3OVg0ma+4X13xc9f9pGGTmchSpx
# dGgCLNVZwketDYYQtlfGcsA8L85RKEffXzIhmySI45GsoGag5D8JnPvRMa9rpeJ3
# 3JiwI/MIXEMg2XGP6HHyVfuQoPwmvE1xbkaIc2MNLjGGClNIMP1TiiT+gaNyFqxv
# POgUoISg1w==
# SIG # End signature block
