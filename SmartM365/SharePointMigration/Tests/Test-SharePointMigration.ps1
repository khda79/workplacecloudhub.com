<#
.SYNOPSIS
    Runs offline launcher and dashboard regression checks using synthetic scans.
.VERSION
    1.0.1
#>
[CmdletBinding()]
param([string]$PythonPath)
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SharePointMigration-tests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null

function Import-TestFunctions {
    param([string]$Path, [string[]]$Names)
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
    foreach ($name in $Names) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $node) { throw "Function missing: $name" }
        Set-Item -Path "function:script:$name" -Value $node.Body.GetScriptBlock()
    }
}
function Assert-Test { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:Checks++ }
$script:Checks = 0
try {
    Import-TestFunctions (Join-Path $project 'Scripts/Launchers/Generic/SmartM365-SharePointMigration-Launcher.ps1') @(
        'Get-PathMappingRows', 'Get-ComparisonConfigValue', 'Get-CsvScanEvidence', 'Assert-CsvScanAgeDifference', 'Invoke-PermissionComparison'
    )
    Import-TestFunctions (Join-Path $project 'SmartM365-SharePointMigration-GUI.ps1') @('Get-SelectedMigrationFolderName')

    $mapping = Join-Path $testRoot 'migration.mapping.txt'
    foreach ($delimiter in @(' ', "`t", ';', ',')) {
        [IO.File]::WriteAllText($mapping, "# comment`n/source%20web${delimiter}/target%20web`n")
        $rows = @(Get-PathMappingRows $mapping)
        Assert-Test ($rows.Count -eq 1 -and $rows[0].Target -eq '/target%20web') 'Valid mapping changed.'
    }
    foreach ($content in @('/source /target ignored', '/source', '# empty', '/source web /target web')) {
        [IO.File]::WriteAllText($mapping, $content)
        $blocked = $false
        try { Get-PathMappingRows $mapping | Out-Null } catch { $blocked = $true }
        Assert-Test $blocked "Invalid mapping accepted: $content"
    }

    # Test only the GUI's folder selection logic; invoking a WPF action here can display a modal dialog.
    $script:CurrentMigration = [pscustomobject]@{ Name = 'Folder Example'; Config = @{ Name = 'Different report label' } }
    Assert-Test ((Get-SelectedMigrationFolderName) -eq 'Folder Example') 'GUI must use the discovered directory, not the report label.'

    # Run the actual permission comparison orchestration with all external work mocked.
    function Resolve-MigrationPath { param($Path) Join-Path $testRoot $Path }
    function New-MigrationRunTimestamp { Get-Date -Format 'yyyyMMdd-HHmmss' }
    function New-MigrationLogPath { param($Action, $Timestamp) Join-Path $testRoot 'test.log' }
    function Start-Transcript { param($Path, [switch]$Force, [switch]$WhatIf) }
    function Stop-LauncherTranscript { param($Path) }
    function Write-Info { param($Message, $Color) }
    function Open-DirectoryInExplorer { param($Path) }
    function Resolve-ComparisonPathMappingsFile { $null }
    function Get-MigrationEndpointType { param($Side) 'SPO' }
    function Get-MigrationEndpointPermissionLibraryOnly { param($Side) $false }
    function Update-EntraUsersCacheForComparison { $script:CacheCalls++; return $SourceCsv }
    function Invoke-PythonScript { param($Script, $Arguments) $script:PythonCalls++ }
    function Read-Host { param($Prompt) return $script:Answer }
    $Config = @{
        Name = 'Synthetic'; Source = @{ PermissionRootPath = '/source' }; Target = @{ PermissionRootPath = '/target' }
        Output = @{ SourcePermissionScans = 'source'; TargetPermissionScans = 'target'; PermissionComparisons = 'comparisons' }
        Comparison = @{ PermissionMaxScanAgeDifferenceHours = 24; ShareGateReplacementCharacter = '_' }
    }
    $ProjectRoot = $project
    $SourceCsv = Join-Path $testRoot 'source.csv'; $TargetCsv = Join-Path $testRoot 'target.csv'
    [IO.File]::WriteAllText($SourceCsv, 'synthetic'); [IO.File]::WriteAllText($TargetCsv, 'synthetic')
    foreach ($case in @(
        @{ Hours = 48; Max = 24; Force = $false; Interactive = $false; Answer = ''; Block = $true },
        @{ Hours = 2; Max = 1; Force = $false; Interactive = $false; Answer = ''; Block = $true },
        @{ Hours = 24; Max = 24; Force = $false; Interactive = $false; Answer = ''; Block = $false },
        @{ Hours = 48; Max = 24; Force = $true; Interactive = $false; Answer = ''; Block = $false },
        @{ Hours = 48; Max = 24; Force = $false; Interactive = $true; Answer = 'NO'; Block = $true },
        @{ Hours = 48; Max = 24; Force = $false; Interactive = $true; Answer = 'YES'; Block = $false }
    )) {
        $now = [datetime]::UtcNow
        (Get-Item $SourceCsv).LastWriteTimeUtc = $now
        (Get-Item $TargetCsv).LastWriteTimeUtc = $now.AddHours(-$case.Hours)
        $Config.Comparison.PermissionMaxScanAgeDifferenceHours = $case.Max
        $Force = $case.Force; $script:LauncherNonInteractive = -not $case.Interactive; $script:Answer = $case.Answer
        $script:CacheCalls = 0; $script:PythonCalls = 0; $blocked = $false
        try { Invoke-PermissionComparison } catch {
            if ($_ -notmatch 'scan age difference') { throw }
            $blocked = $true
        }
        Assert-Test ($blocked -eq $case.Block) 'Permission scan age decision incorrect.'
        $expectedCalls = if ($case.Block) { 0 } else { 1 }
        Assert-Test ($script:CacheCalls -eq $expectedCalls -and $script:PythonCalls -eq $expectedCalls) 'Age validation must precede Entra/cache and comparison work.'
    }

    # A copy changes LastWriteTime, but the legacy filename retains the original scan date.
    $legacy = Join-Path $testRoot 'SP2019-FileInventory-Synthetic-20260927-120000.csv'
    [IO.File]::WriteAllText($legacy, "Name`nfile`n")
    (Get-Item $legacy).LastWriteTimeUtc = [datetime]::UtcNow
    $legacyEvidence = Get-CsvScanEvidence -CsvPath $legacy -WarningAction SilentlyContinue
    Assert-Test ($legacyEvidence.CompletedAtUtc -lt [datetime]::UtcNow.AddDays(-1)) 'Copied legacy scan must use its filename date.'
    $legacyPeer = Join-Path $testRoot 'SPO-FileInventory-Synthetic-20260927-130000.csv'
    [IO.File]::WriteAllText($legacyPeer, "Name`nfile`n")
    (Get-Item $legacyPeer).LastWriteTimeUtc = [datetime]::UtcNow
    $Force = $false; $script:LauncherNonInteractive = $true
    $absoluteAgeBlocked = $false
    try {
        Assert-CsvScanAgeDifference -SourceCsvPath $legacy -TargetCsvPath $legacyPeer -MaxAgeDifferenceHours 12 -MaxAgeHours 24 -WarningAction SilentlyContinue
    } catch { $absoluteAgeBlocked = $_ -match 'absolute ages' }
    Assert-Test $absoluteAgeBlocked 'Two copied old scans with a small gap must be blocked.'

    $receiptCsv = Join-Path $testRoot 'SP2019-FileInventory-Synthetic-20261002-120000.csv'
    [IO.File]::WriteAllText($receiptCsv, "Name`nfile`n")
    $python = if ($PythonPath) { $PythonPath } else { Join-Path $project 'Tools/Python/python.exe' }
    & $python (Join-Path $project 'Scripts/Compare/scan_evidence.py') --csv $receiptCsv --side Source --kind File --scope 'https://example.com/site' | Out-Null
    Assert-Test ($LASTEXITCODE -eq 0) 'Manifest generation failed.'
    $receiptEvidence = Get-CsvScanEvidence -CsvPath $receiptCsv
    Assert-Test ($receiptEvidence.Rows -eq 1 -and $receiptEvidence.Provenance -like '*.manifest.json.txt') 'Manifest data was not loaded.'
    $manifestPath = "$receiptCsv.manifest.json.txt"
    $manifestText = [IO.File]::ReadAllText($manifestPath)
    $manifestText = [regex]::Replace($manifestText, '"CompletedAtUtc":\s*"[^"]+"', '"CompletedAtUtc": "2026-10-03T14:26:52.084752+00:00"')
    [IO.File]::WriteAllText($manifestPath, $manifestText)
    $originalCulture = [Globalization.CultureInfo]::CurrentCulture
    try {
        [Globalization.CultureInfo]::CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
        $datedEvidence = Get-CsvScanEvidence -CsvPath $receiptCsv
    } finally {
        [Globalization.CultureInfo]::CurrentCulture = $originalCulture
    }
    $expectedUtc = ([datetimeoffset]::Parse('2026-10-03T14:26:52.084752+00:00', [Globalization.CultureInfo]::InvariantCulture)).UtcDateTime
    Assert-Test ($datedEvidence.CompletedAtUtc -eq $expectedUtc) 'ISO manifest timestamp changed month/day under fr-FR culture.'
    [IO.File]::AppendAllText($receiptCsv, "tampered`n")
    $tamperBlocked = $false
    try { Get-CsvScanEvidence -CsvPath $receiptCsv | Out-Null } catch { $tamperBlocked = $true }
    Assert-Test $tamperBlocked 'Changed inventory was accepted despite manifest hash mismatch.'
    $emptyCsv = Join-Path $testRoot 'SPO-FileInventory-Synthetic-20261002-120000.csv'
    [IO.File]::WriteAllText($emptyCsv, "Name`n")
    & $python (Join-Path $project 'Scripts/Compare/scan_evidence.py') --csv $emptyCsv --side Target --kind File --scope 'https://example.com/site' | Out-Null
    Assert-Test ($LASTEXITCODE -eq 0) 'Empty inventory manifest generation failed.'
    $emptyBlocked = $false
    try { Assert-CsvScanAgeDifference -SourceCsvPath $emptyCsv -TargetCsvPath $emptyCsv -MaxAgeDifferenceHours 12 -MaxAgeHours 24 } catch {
        $emptyBlocked = $_ -match 'zero data rows'
    }
    Assert-Test $emptyBlocked 'Header-only inventory must not be compared.'
    $freshSource = Join-Path $testRoot 'SP2019-FileInventory-Synthetic-fresh.csv'
    $freshTarget = Join-Path $testRoot 'SPO-FileInventory-Synthetic-fresh.csv'
    [IO.File]::WriteAllText($freshSource, "Name`nfile`n")
    [IO.File]::WriteAllText($freshTarget, "Name`n")
    & $python (Join-Path $project 'Scripts/Compare/scan_evidence.py') --csv $freshSource --side Source --kind File --scope 'https://example.com/source' | Out-Null
    Assert-Test ($LASTEXITCODE -eq 0) 'Fresh source inventory manifest generation failed.'
    & $python (Join-Path $project 'Scripts/Compare/scan_evidence.py') --csv $freshTarget --side Target --kind File --scope 'https://example.com/target' | Out-Null
    Assert-Test ($LASTEXITCODE -eq 0) 'Fresh target inventory manifest generation failed.'
    $Force = $false; $script:LauncherNonInteractive = $true
    Assert-CsvScanAgeDifference -SourceCsvPath $freshSource -TargetCsvPath $freshTarget -MaxAgeDifferenceHours 12 -MaxAgeHours 24 -Label 'File inventory' -AllowEmptyTarget -WarningAction SilentlyContinue
    $permissionBlocked = $false
    try { Assert-CsvScanAgeDifference -SourceCsvPath $freshSource -TargetCsvPath $freshTarget -MaxAgeDifferenceHours 12 -MaxAgeHours 24 -Label 'Permission inventory' } catch {
        $permissionBlocked = $_ -match 'target inventory contains zero data rows'
    }
    Assert-Test $permissionBlocked 'Permission comparison must still reject an empty target inventory.'
    $sourceBlocked = $false
    try { Assert-CsvScanAgeDifference -SourceCsvPath $freshTarget -TargetCsvPath $freshTarget -MaxAgeDifferenceHours 12 -MaxAgeHours 24 -Label 'File inventory' -AllowEmptyTarget -WarningAction SilentlyContinue } catch {
        $sourceBlocked = $_ -match 'source inventory contains zero data rows'
    }
    Assert-Test $sourceBlocked 'File comparison must still reject an empty source inventory.'
    Write-Output "PASS: $script:Checks offline PowerShell assertions; no tenant calls or launched processes."
}
finally {
    $resolvedRoot = [IO.Path]::GetFullPath($testRoot)
    $allowedRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolvedRoot.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolvedRoot -Leaf) -like 'SharePointMigration-tests-*') {
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBP2dY6wI3auDFA
# oj5YtQsFYBcB8CCuUa/0IJPUAwhLs6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD+np+kCZPjVPUz+iqi7o1S
# 3j/0bbiA34lZYwXniG7VEzANBgkqhkiG9w0BAQEFAASCAYACTBAfV5MwSoo6xyCa
# 6IsoHDLWtGoI1ZQt0meNng6y/VUqs17RITH7rwy/gn+ZPybA6h1s4+6Si4J75AbA
# 6UPG/l7RfRwi3Dq+MXN67iPPrAXgvJrv51Ed8rQAZCv3Qjaf/5gRmn6B0vGWpU5g
# uj8DMFDX3xIGZLjylvgcOABuz+GeBT6spnSzTw731mGi2RKFbeBUw8/k+K/mocma
# z1MeafiVEdZ4WApakxU/xAb1dmIARw4NJP6YTInipErQpqxvcBu7DyTB35aOl5Im
# XCQqofyK6fkp5KBs7BQroKR/1WplG1tBI+JTAvSHJVb5P2E/ix/JgWM/FqZw5JXp
# Fgtua1+iOwgGsqjMZGGCAU4JE4l8hj45tlGVER4eASnulJHbefFZrUXfzAFAe0LO
# 1td0gGZSJ5ymXuCvIpr8HAfWUgpWNaNsSxljfzx408uRi1ce9vchfPA8K7FpoJ4r
# EGDCct0tip0LIbnOXDipqfmXQqdDFS9qLgZC7THmB5NuBzM=
# SIG # End signature block
