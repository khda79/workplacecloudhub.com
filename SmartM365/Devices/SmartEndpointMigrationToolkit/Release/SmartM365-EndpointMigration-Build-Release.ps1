#Requires -Version 5.1

<#
.SYNOPSIS
    Validates and optionally builds the standalone Smart Endpoint Migration Toolkit release.

.DESCRIPTION
    Preview is the default and writes nothing. -Build stages the package files, writes the
    distribution manifest (SmartEndpointMigrationToolkit.components.json.txt: every file with
    its SHA256, size and declared version) inside the staged package, then creates a
    deterministic ZIP (sorted entries, fixed timestamp), a SHA256 checksum file and the
    release notes in the output folder.

    The file set is the package folder minus operational and private content: Runs, the
    content of Waves, Inventory, DemoAssets, Release\Output and Release\Work, *.local.json,
    LOCAL_MEMORY.md, logs, inventories, setup media and IntuneWin32 build output. The build
    fails when such a file would be shipped, when a PowerShell file does not parse, when a
    script header version differs from the version manifest or when a component is not signed
    by the pinned signer (unless -AllowUnsigned is given for an offline qualification build,
    which is then marked RequireSignature=false and UNSIGNED in the release notes).

    Sign the sources before building a production release: the manifest records the hashes of
    the signed files. This script never publishes, tags, pushes or deploys anything.

.EXAMPLE
    .\SmartM365-EndpointMigration-Build-Release.ps1

.EXAMPLE
    .\SmartM365-EndpointMigration-Build-Release.ps1 -Build -AllowUnsigned

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$OutputRoot,
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [switch]$Build,
    [switch]$AllowUnsigned,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Host {
    param([Parameter(Position = 0, ValueFromRemainingArguments = $true)]$Object, [System.ConsoleColor]$ForegroundColor)
    $text = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),(@($Object) -join ' ')
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { Microsoft.PowerShell.Utility\Write-Host $text -ForegroundColor $ForegroundColor } else { Microsoft.PowerShell.Utility\Write-Host $text }
}

if ([string]::IsNullOrWhiteSpace($SourceRoot)) { $SourceRoot = Split-Path -Path $PSScriptRoot -Parent }
$SourceRoot = (Resolve-Path -LiteralPath $SourceRoot).Path.TrimEnd('\')
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Join-Path $SourceRoot 'Release\Output' }

$versionManifest = Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartEndpointMigrationToolkit.version.json.txt') -Raw | ConvertFrom-Json
$packageVersion = [string]$versionManifest.PackageVersion
if ($packageVersion -notmatch '^\d+\.\d+\.\d+$') { throw "Invalid PackageVersion: $packageVersion" }
$artifactBaseName = 'SmartEndpointMigrationToolkit-{0}' -f $packageVersion
$releaseTag = [string]$versionManifest.ReleaseTag
if ($releaseTag -ne ('endpoint-migration-toolkit-v{0}' -f $packageVersion)) { throw "ReleaseTag does not match PackageVersion: $releaseTag" }

$excludedPatterns = @(
    '^Runs\\(?!\.gitkeep$)',
    '^Waves\\[^\\]+\\[^\\]+\\',
    '^Inventory\\',
    '^DemoAssets\\',
    '^Release\\(Output|Work)\\',
    '\.local\.json$',
    '(^|\\)LOCAL_MEMORY\.md$',
    '\.log$',
    '(^|\\)Devices(AD|Intune|Entra)[^\\]*\.csv$',
    '(^|\\)Computers(\.[A-Za-z]+)?\.txt$',
    '^Campaigns\\Windows11Upgrade\\SetupSource\\(?!README\.md$|\.gitignore$)',
    '^Campaigns\\Windows11Upgrade\\SetupSourceGates\\(?!README\.md$|\.gitignore$)',
    '^Campaigns\\Windows11Upgrade\\IntuneWin32\\(Build|Output|SetupMedia)\\',
    '(^|\\)\.vs\\',
    '\.(tmp|bak|zip|intunewin|iso|wim|esd)$',
    '^SmartEndpointMigrationToolkit\.components\.json\.txt$'
)
$releaseFiles = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
foreach ($file in @(Get-ChildItem -LiteralPath $SourceRoot -Recurse -File -Force)) {
    $relative = $file.FullName.Substring($SourceRoot.Length).TrimStart('\')
    $excluded = $false
    foreach ($pattern in $excludedPatterns) { if ($relative -match $pattern) { $excluded = $true; break } }
    if ($excluded) { $skipped.Add($relative) } else { $releaseFiles.Add($relative) }
}
$releaseFiles = @($releaseFiles | Sort-Object)
Write-Host ("Package {0}: {1} file(s) selected, {2} operational/private file(s) excluded." -f $packageVersion,$releaseFiles.Count,$skipped.Count)

$problems = New-Object System.Collections.Generic.List[string]
$entries = New-Object System.Collections.Generic.List[object]
$unsigned = 0
foreach ($relative in $releaseFiles) {
    $full = Join-Path $SourceRoot $relative
    $version = ''
    $signature = ''
    if ($relative -match '\.(ps1|psm1|psd1)$') {
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($full, [ref]$tokens, [ref]$errors)
        if (@($errors).Count -gt 0) { $problems.Add(('Parse error: {0}: {1}' -f $relative,$errors[0].Message)) }
        $text = [IO.File]::ReadAllText($full)
        $headerMatch = [regex]::Match($text, '(?m)^\s*\.VERSION\s*\r?\n\s*(?<Version>\d+\.\d+\.\d+)')
        if ($headerMatch.Success) { $version = $headerMatch.Groups['Version'].Value }
        elseif ($relative -match '\.psd1$') {
            $moduleVersion = [regex]::Match($text, "ModuleVersion\s*=\s*'(?<Version>[^']+)'")
            if ($moduleVersion.Success) { $version = $moduleVersion.Groups['Version'].Value }
        }
        $sig = Get-AuthenticodeSignature -LiteralPath $full
        $signature = [string]$sig.Status
        if ($sig.Status -eq 'Valid') {
            if (([string]$sig.SignerCertificate.Thumbprint).ToUpperInvariant() -ne $ExpectedSignerThumbprint.ToUpperInvariant()) { $problems.Add(('Unexpected signer: {0}: {1}' -f $relative,$sig.SignerCertificate.Thumbprint)) }
        }
        elseif ($sig.Status -eq 'NotSigned') { $unsigned++ }
        else { $problems.Add(('Invalid signature: {0}: {1}' -f $relative,$sig.Status)) }
    }
    $item = Get-Item -LiteralPath $full
    $entries.Add([pscustomobject]@{ Path = $relative.Replace('\','/'); Sha256 = (Get-FileHash -LiteralPath $full -Algorithm SHA256).Hash; Size = $item.Length; Version = $version; Signature = $signature })
}

$declared = [ordered]@{
    'Scripts/SmartM365-Invoke-EndpointMigrationWithPsExec.ps1' = $versionManifest.Components.Engine
    'Scripts/SmartM365-EndpointMigration-PsExecWorker.ps1' = $versionManifest.Components.Worker
    'Modules/SmartEndpointMigration.Core/SmartEndpointMigration.Core.psd1' = $versionManifest.Components.CoreModule
    'Scripts/SmartM365-EndpointMigration-GUI.ps1' = $versionManifest.Components.GUI
    'Campaigns/HybridJoin/Scripts/SmartM365-Invoke-EndpointMigration-HybridJoin.ps1' = $versionManifest.Components.HybridJoinEndpoint
    'Campaigns/Windows11Upgrade/Scripts/SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1' = $versionManifest.Components.Windows11UpgradeEndpoint
}
foreach ($path in $declared.Keys) {
    $entry = @($entries | Where-Object { $_.Path -eq $path })
    if ($entry.Count -ne 1) { $problems.Add(('Declared component missing from the release: {0}' -f $path)); continue }
    if ($entry[0].Version -ne [string]$declared[$path]) { $problems.Add(('Version mismatch: {0}: header={1}; version manifest={2}' -f $path,$entry[0].Version,$declared[$path])) }
}
# Shared SmartM365.Gui* libraries are verbatim copies; their provenance is in Docs/PROVENANCE.json.txt.
foreach ($entry in @($entries | Where-Object { $_.Path -match '\.(ps1|psm1)$' -and $_.Path -notmatch '(^|/)SmartM365\.Gui[A-Za-z.]*\.ps1$' -and [string]::IsNullOrWhiteSpace($_.Version) })) { $problems.Add(('Missing .VERSION header: {0}' -f $entry.Path)) }
foreach ($required in @('README.md','Start-SmartEndpointMigrationToolkit-GUI.cmd','Start-SmartEndpointMigrationToolkit-Demo.cmd','Scripts/Run-EndpointMigrationWave.cmd','Config/SmartEndpointMigrationToolkit.options.json.txt','Docs/PROVENANCE.json.txt')) {
    if (@($entries | Where-Object { $_.Path -eq $required }).Count -eq 0) { $problems.Add(('Required release file missing: {0}' -f $required)) }
}
if ($unsigned -gt 0 -and -not $AllowUnsigned) { $problems.Add(('{0} PowerShell file(s) are not signed. Sign the sources first, or use -AllowUnsigned for an offline qualification build.' -f $unsigned)) }

if ($problems.Count -gt 0) {
    foreach ($problem in $problems) { Write-Host ('ERROR: ' + $problem) -ForegroundColor Red }
    throw ('Release validation failed with {0} problem(s).' -f $problems.Count)
}
Write-Host ("Validation passed: {0} file(s); unsigned PowerShell files: {1}." -f $entries.Count,$unsigned)
if (-not $Build) {
    Write-Host 'Preview only. Use -Build to create the release artifacts. Nothing was written.'
    return [pscustomobject]@{ Version = $packageVersion; Files = $entries.Count; Unsigned = $unsigned; Built = $false }
}

$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$workRoot = Join-Path $OutputRoot ('Work-{0}' -f $packageVersion)
$stageRoot = Join-Path $workRoot 'SmartEndpointMigrationToolkit'
$zipPath = Join-Path $OutputRoot ($artifactBaseName + '.zip')
$shaPath = Join-Path $OutputRoot ($artifactBaseName + '.sha256')
$notesPath = Join-Path $OutputRoot ('RELEASE-NOTES-{0}.md' -f $packageVersion)
foreach ($existing in @($zipPath,$shaPath,$notesPath)) { if ((Test-Path -LiteralPath $existing) -and -not $Force) { throw "Release artifact already exists (use -Force to replace): $existing" } }
if (Test-Path -LiteralPath $workRoot) { Remove-Item -LiteralPath $workRoot -Recurse -Force }
New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null

foreach ($entry in $entries) {
    $target = Join-Path $stageRoot ($entry.Path.Replace('/','\'))
    New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $SourceRoot ($entry.Path.Replace('/','\'))) -Destination $target -Force
}
$buildDate = [datetime]::ParseExact([string]$versionManifest.ReleaseDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
$componentManifest = [pscustomobject]@{
    SchemaVersion = 1
    Product = 'SmartEndpointMigrationToolkit'
    PackageVersion = $packageVersion
    ReleaseTag = $releaseTag
    WorkerContractVersion = [int]$versionManifest.WorkerContractVersion
    RequireSignature = (-not [bool]$AllowUnsigned -or $unsigned -eq 0)
    Signer = $ExpectedSignerThumbprint
    BuiltOn = $buildDate.ToString('yyyy-MM-dd')
    Files = @($entries | ForEach-Object { [pscustomobject]@{ Path = $_.Path; Sha256 = $_.Sha256; Size = $_.Size; Version = $_.Version } })
}
$manifestPath = Join-Path $stageRoot 'SmartEndpointMigrationToolkit.components.json.txt'
[IO.File]::WriteAllText($manifestPath, ($componentManifest | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
$fixedTime = [DateTimeOffset]::new($buildDate.Year, $buildDate.Month, $buildDate.Day, 0, 0, 0, [TimeSpan]::Zero)
$stream = [IO.File]::Open($zipPath, [IO.FileMode]::CreateNew)
try {
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $all = @(Get-ChildItem -LiteralPath $stageRoot -Recurse -File | ForEach-Object { $_.FullName.Substring($stageRoot.Length).TrimStart('\').Replace('\','/') } | Sort-Object { $_ } -CaseSensitive)
        foreach ($relative in $all) {
            $zipEntry = $archive.CreateEntry(('SmartEndpointMigrationToolkit/{0}' -f $relative), [IO.Compression.CompressionLevel]::Optimal)
            $zipEntry.LastWriteTime = $fixedTime
            $source = [IO.File]::OpenRead((Join-Path $stageRoot $relative.Replace('/','\')))
            try { $target = $zipEntry.Open(); try { $source.CopyTo($target) } finally { $target.Dispose() } } finally { $source.Dispose() }
        }
    }
    finally { $archive.Dispose() }
}
finally { $stream.Dispose() }

$zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash
$manifestHash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
[IO.File]::WriteAllLines($shaPath, @(('{0} *{1}' -f $zipHash,(Split-Path $zipPath -Leaf)),('{0} *SmartEndpointMigrationToolkit/SmartEndpointMigrationToolkit.components.json.txt' -f $manifestHash)), [Text.Encoding]::ASCII)
$signedText = if ($componentManifest.RequireSignature) { 'Signed release: every PowerShell component is signed by the pinned WorkplaceCloudHub signer and RequireSignature=true.' } else { ('UNSIGNED QUALIFICATION BUILD: {0} PowerShell file(s) are not signed and RequireSignature=false. Do not use this package for production endpoints; rebuild from signed sources.' -f $unsigned) }
$notes = @(
    ('# Smart Endpoint Migration Toolkit {0}' -f $packageVersion),
    '',
    ('Release tag (not created by this script): `{0}`' -f $releaseTag),
    '',
    $signedText,
    '',
    '## Artifacts',
    '',
    ('- `{0}` (SHA256 `{1}`)' -f (Split-Path $zipPath -Leaf),$zipHash),
    ('- `{0}`' -f (Split-Path $shaPath -Leaf)),
    ('- Distribution manifest inside the package: `SmartEndpointMigrationToolkit.components.json.txt` (SHA256 `{0}`, {1} files)' -f $manifestHash,$entries.Count),
    '',
    '## Components',
    ''
) + @($declared.Keys | ForEach-Object { '- `{0}` {1}' -f $_,$declared[$_] }) + @(
    '',
    '## Scope',
    '',
    '- Campaigns: Hybrid Join and Intune enrollment, Windows 11 upgrade.',
    '- Waves, Single PC Runs, Automatic Wave, Launch All Waves (selected Campaign only), controlled stop.',
    '- Coexists with IntuneHybridJoinToolkit and Windows11UpgradeToolkit; no configuration, Wave, inventory, history or report is migrated. Existing Windows 11 setup media is reused or moved as documented in Docs/LEGACY-COEXISTENCE.md.',
    '- Windows 11 setup media: one read-only resolution for PsExec and Intune (new cache, legacy media owned by the previous Intune package used in place, idle legacy media adopted with verified rollback, no silent duplication). Legacy media that is present but not write-restricted is Untrusted (explicit status, never used, permissions never changed, second copy only with the duplication authorization); unreadable permissions are Indeterminate and block.',
    '- Scripts run as SYSTEM on targets (endpoint started by PsExec, Intune runner, retry-after-reboot tasks) are stored under C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign> with owner Administrators and a protected ACL; ProgramData keeps data only. An untrusted program folder stops the operation (SEMT_PROGRAM_FOLDER_UNTRUSTED).',
    '- ProgramData: only the SEMT components of the data folder chain get owner Administrators and a protected DACL; shared parents (WCH, WCH\SmartM365) are verified, never changed; a read-only pass refuses reparse points, hard links and existing content without verified provenance before any change; failure is blocking (SEMT_DATA_FOLDER_UNTRUSTED, DATA_FOLDER_UNTRUSTED). The Intune runner verifies both folders before it runs anything. Setup media without its manifest, or without a required package reference, is refused. Intune WithCacheOnly for devices with reusable media, WithMedia otherwise.',
    '- PowerShell: GUI and endpoint scripts on Windows PowerShell 5.1; engine, worker and tools on Windows PowerShell 5.1 or PowerShell 7.2+. Unsupported engines are refused before any state is written.',
    '',
    'See README.md, Docs/TEST-REPORT.md and Docs/ACCEPTANCE-PROCEDURE.md.'
)
[IO.File]::WriteAllLines($notesPath, $notes, [Text.UTF8Encoding]::new($false))
Remove-Item -LiteralPath $workRoot -Recurse -Force
Write-Host ("Built {0} ({1}); checksums {2}; notes {3}." -f $zipPath,$zipHash,$shaPath,$notesPath)
return [pscustomobject]@{ Version = $packageVersion; Files = $entries.Count; Unsigned = $unsigned; Built = $true; ZipPath = $zipPath; ZipSha256 = $zipHash; ShaPath = $shaPath; NotesPath = $notesPath; RequireSignature = $componentManifest.RequireSignature }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCOkXv5S3eUcSm6
# 3DJ2yT4CWW1+6U7N5KOSAM0lb2Nxj6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFXY1eOqDz2wzPcqniod8DKVXSTaSvamASRtz6n2767HMA0GCSqG
# SIb3DQEBAQUABIIBgFZf9PtJfhH3gGvteFR1JEGiBvWuZXpdrf+7fOWL31GcptcI
# zHV0yYM9y+URrPYhTjmXeOeS+Vv0yH3MVqMVlFbgoje8AOf/Tnxx8fC6cFzQwboG
# Tg0QtC8CIZN1efuz52Byqse2dWrkVovy96vbON56QxLfHcewdMjrToADB19/DE7v
# Lp4aIvWxFursm7Skbm/Nj7kWVBQ++9wkOce5ntlpGBr/CDpeGpd5Lo54Vl88eVQ+
# Zb66bNaDgfiLlm944i88DT/W4Uwzas2k3fB0BdI+YhamMDttCrICoplTCKcsB2C+
# 8ytznoEc5J9ZLh37XwLtLC2qSa2Ci/N3+pAYE73Sl6USEyeUm3Fgp2y0hBQsh2bz
# RKX8VA/MMull/779Ngv/W/WEcExmnLU4dHSNmCay0ekG7MlOZuD7dPJ7S+AAc78x
# aeuxzMv2xrBOZH+Jmg6EmrwdukrZoFNB9Pm3GwM4I/gCtVwyyHG1ZzsvmrJZlzTW
# 1PpIlyGHaIxAzM3XnKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTRaMC8GCSqGSIb3DQEJBDEiBCBnjYT3jDVzzPYM0uc+AckjOfawLH7s0yEeDKr6
# cFPV6jANBgkqhkiG9w0BAQEFAASCAgBUxPWM4Ykr6I4kFai/rglEsOmVnhQG5lN/
# Xy9KmxE58BwQmVCWp9sH2d5DA4g9bCKMcX7KXUOO55WRb3i1VteKzUNWGw2/ASh3
# vZtkkR93JD+qWJCGMTp2LzW4XzvoikqQXs5Y+dsKDHHrVz9Z3il2FMhrFOZwd8z+
# Tgv6Ot7q5zRnyGNzte/s7URf72LLs/TyPIegQPsU8fpNkGUAxKjE3EN6kKliZxl9
# iFvuUu4jEJA5rz9MmJmS6IdzbXNmUH4Y5pavAWNTPtB1gg3G9o6Gtm36EDge1m5r
# IeqzVKnMTvxEJqTQGHVdrIhclKYYLge6pjgTUPlpVIiph+LRr7Z/0HysoeZRBiUR
# PHQbhRP/qk37BVrcSJd0vBxP6qxYbnfPA8D1F6HEvFp1eTBvF6Gv8xZodbcX8bqW
# T2qXrHEkPeaAWokYgi6nlOc+sdXQqGHVF/sof03t7QSL/dpPhN2aJT1PlNWeAPur
# NbVHkK0lStW1Djq2yRFzz/aXEuZUclJT4jUm+Zn1p66vl4J/wu9qIcfZ/SCfsYzr
# YvFaQR54tl5tl2A5mbe1WxhlQarN1YgP55kdQrnR3QPOJ0flFG/5LMQj5Z7bUi4T
# 3tNqg5eF7/Kvm8Jwgrrnk+OjJltghtSfil9j9aj0O+thOZp9PIZOIyvguON7+8lb
# EhuqXkt7lA==
# SIG # End signature block
