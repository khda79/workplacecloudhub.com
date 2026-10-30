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
foreach ($required in @('README.md','Start-SmartEndpointMigrationToolkit-GUI.cmd','Start-SmartEndpointMigrationToolkit-Demo.cmd','Export-UserLicensesCsv.cmd','Scripts/Run-EndpointMigrationWave.cmd','Config/SmartEndpointMigrationToolkit.options.json.txt','Docs/PROVENANCE.json.txt')) {
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBPMRERDivkuEgL
# 89WLDs01SUCKbI9j1FSuviI9hpSftKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA3U5IdhCc7W3O6AFEW2DJa
# JShMxHYfYetGcen5xoFGXzANBgkqhkiG9w0BAQEFAASCAYAKxXMzaOQVkG9R2mUG
# qwaCdJAFLBqSTjG6PY/pCoZRgGN1iV9Ycfz48pVz9q0NEG95HGvP3HIcSIIR6vuO
# 0BRPE3qOwvSaSTwUY1rrm8CLf36mGm4d49HCQga6j3XKz6MUEGEaUsVA9R67aPrB
# uQNx+zs4PLFmVMQxxW7F5pntImYZQ/g3H6rANWdgaLBB6yfCG5qqse578V9i8MxQ
# yZ8SQmbPJIA6owLSXrDMEF6aSZwtSVjmEfa4zl8AIf2VvYt8PYS/Il3H0rY5bQ4a
# yz8lkyJdtgZhnC6hKQYYE2yORu38nrCNlfNEEQSwWkB+EzqxPz54lRgvB69p6tp4
# TRdH0p62XtgXveNKWcCbVJgJc2zLKBxbSWSO/zAtDMXrRJyXWuxLNLLhdPYNo9EX
# zfg9zS1eYQ9GseFiEUMT7WNO8PL7O+I3BrfvkyZOisv04o2GtM+6rBpWMFaWfd2V
# kqwQnXk2fmbpXg6oO+B8oZdRPrsu6i6zzuKQdc3Ar/lyCGs=
# SIG # End signature block
