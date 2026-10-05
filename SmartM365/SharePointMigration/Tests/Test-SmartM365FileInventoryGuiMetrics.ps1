<#
.SYNOPSIS
    Verify source and target scan metrics and the GUI's metadata-only display.
.VERSION
    1.0.1
#>
#Requires -Version 7.4

$ErrorActionPreference = 'Stop'
function Import-FunctionFromScript {
    param([string]$Path, [string]$Name)
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ($errors | ForEach-Object Message) }
    $definition = $ast.Find({
        param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name
    },$true)
    if (-not $definition) { throw "Missing function $Name in $Path" }
    $body = $definition.Body.Extent.Text
    Set-Item -Path ("function:global:{0}" -f $Name) -Value ([scriptblock]::Create($body.Substring(1, $body.Length - 2)))
}
function Write-Info { param([string]$Color, [string]$Message) }

$project = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
$gui = Join-Path $project 'SmartM365-SharePointMigration-GUI.ps1'
Import-FunctionFromScript -Path $gui -Name 'Format-FileInventoryVolume'
Import-FunctionFromScript -Path $gui -Name 'Get-FileInventoryMetricCsvHash'
Import-FunctionFromScript -Path $gui -Name 'Get-FileInventoryMetricPresentation'
$root = Join-Path $PSScriptRoot ('.file-inventory-gui-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {
    throw 'Unsafe test path.'
}
try {
    [void](New-Item -ItemType Directory -Path $root)
    foreach ($side in @('Source','Target')) {
        $scriptPath = Join-Path $project ("Scripts\Inventory\SmartM365-SharePoint{0}-FileInventory.ps1" -f $side)
        Import-FunctionFromScript -Path $scriptPath -Name 'Add-FileInventoryMetric'
        Import-FunctionFromScript -Path $scriptPath -Name 'Write-FileInventoryMetrics'
        $script:MetricFileSizes = @{}
        $script:MetricFolders = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $script:MetricRows = [long]0
        $script:MetricKnownSizeBytes = [long]0
        $script:MetricMissingPathRows = [long]0
        $script:MetricInvalidLibraryRows = [long]0
        $script:MetricDuplicateRows = [long]0
        $script:MetricConflictingSizeRows = [long]0
        $path = Join-Path $root ("{0}-FileInventory-Fixture-20261004-150000.csv" -f $(if ($side -eq 'Source') { 'SP2019' } else { 'SPO' }))
        'ServerRelativeUrl;LibraryUrl;SizeBytes' | Set-Content -LiteralPath $path -Encoding utf8
        Add-FileInventoryMetric -Row ([pscustomobject]@{ ServerRelativeUrl='/sites/a/Docs/one.txt'; LibraryUrl='/sites/a/Docs'; SizeBytes=1024 })
        Add-FileInventoryMetric -Row ([pscustomobject]@{ ServerRelativeUrl='/sites/a/Docs/Folder/two.txt'; LibraryUrl='/sites/a/Docs'; SizeBytes=2048 })
        Add-FileInventoryMetric -Row ([pscustomobject]@{ ServerRelativeUrl='/sites/a/Docs/folder/TWO.txt'; LibraryUrl='/sites/a/Docs'; SizeBytes=2048 })
        if ($side -eq 'Target') {
            Add-FileInventoryMetric -Row ([pscustomobject]@{ ServerRelativeUrl='/sites/a/Docs/Folder/three.txt'; LibraryUrl='/sites/a/Docs'; SizeBytes=$null })
        }
        Write-FileInventoryMetrics -CsvPath $path
        $metadata = Get-Content -LiteralPath "$path.metrics.json.txt" -Raw | ConvertFrom-Json
        if ($metadata.Files -ne $(if ($side -eq 'Source') { 2 } else { 3 }) -or
            $metadata.FoldersWithFiles -ne 1 -or $metadata.KnownSizeBytes -ne 3072 -or
            $metadata.DuplicateRows -ne 1 -or $metadata.CsvSha256 -ne (Get-FileHash -LiteralPath $path).Hash) { throw "Wrong scan metrics for $side" }
        $scan = [pscustomobject]@{ File=(Get-Item -LiteralPath $path); Date=[datetime]'2026-10-04 15:00:00'; Provenance='fixture' }
        $display = Get-FileInventoryMetricPresentation -Scan $scan
        if ($display.Files -ne [string]$metadata.Files -or $display.Folders -ne '1') { throw "GUI did not read $side metrics." }
        if ($side -eq 'Source' -and $display.Volume -notmatch '^3[.,]0 KiB$') { throw 'Source volume was not displayed.' }
        if ($side -eq 'Target' -and $display.Volume -ne '—') { throw 'Incomplete target volume was displayed as complete.' }
        # A new hash-bound sidecar survives synchronization changing the timestamp.
        (Get-Item -LiteralPath $path).LastWriteTimeUtc = [datetime]::UtcNow.AddDays(-1)
        $display = Get-FileInventoryMetricPresentation -Scan $scan
        if ($display.Files -ne [string]$metadata.Files -or $display.Evidence -notmatch 'SHA256 verified') { throw 'Hash-bound metrics were hidden after synchronization.' }
        # Earlier sidecars can be recovered only when whole-second truncation and a valid receipt agree.
        $legacyMetadata = $metadata | Select-Object * -ExcludeProperty CsvSha256
        $legacyMetadata | ConvertTo-Json | Set-Content -LiteralPath "$path.metrics.json.txt" -Encoding utf8
        $rounded = [long]$metadata.CsvLastWriteTimeUtcTicks - ([long]$metadata.CsvLastWriteTimeUtcTicks % [TimeSpan]::TicksPerSecond)
        (Get-Item -LiteralPath $path).LastWriteTimeUtc = [datetime]::new($rounded,[DateTimeKind]::Utc)
        @{SchemaVersion=1;InventoryFile=(Get-Item $path).Name;Rows=$metadata.Rows;Sha256=$metadata.CsvSha256} |
            ConvertTo-Json | Set-Content -LiteralPath "$path.manifest.json.txt"
        $display = Get-FileInventoryMetricPresentation -Scan $scan
        if ($display.Files -ne [string]$metadata.Files -or $display.Evidence -notmatch 'receipt SHA256 verified') { throw 'Synchronized legacy metrics were not restored from verified receipt.' }
        Remove-Item -LiteralPath "$path.manifest.json.txt"
        if ((Get-FileInventoryMetricPresentation -Scan $scan).Files -ne '—') { throw 'Rounded legacy metrics without a receipt were accepted.' }
        $metadata | ConvertTo-Json | Set-Content -LiteralPath "$path.metrics.json.txt" -Encoding utf8
        if ($side -eq 'Source') {
            $before = [IO.File]::ReadAllText($path)
            [IO.File]::WriteAllText($path,$before.Replace('SizeBytes','FakeBytes'),[Text.UTF8Encoding]::new($false))
            if ((Get-FileInventoryMetricPresentation -Scan $scan).Files -ne '—') { throw 'Same-length changed content was accepted.' }
            Add-Content -LiteralPath $path -Value 'changed'
            if ((Get-FileInventoryMetricPresentation -Scan $scan).Files -ne '—') { throw 'Stale sidecar was accepted.' }
        }
    }
    $legacy = Join-Path $root 'SP2019-FileInventory-Legacy.csv'
    'ServerRelativeUrl;LibraryUrl;SizeBytes' | Set-Content -LiteralPath $legacy -Encoding utf8
    $legacyScan = [pscustomobject]@{ File=(Get-Item -LiteralPath $legacy); Date=[datetime]'2026-10-03 15:00:00'; Provenance='legacy' }
    $legacyDisplay = Get-FileInventoryMetricPresentation -Scan $legacyScan
    if ($legacyDisplay.Files -ne '—' -or $legacyDisplay.Table -notmatch 'Rerun scan') {
        throw 'A legacy CSV without scan metrics was presented as measured.'
    }
    'SharePointMigration scan metrics and GUI metadata test passed.'
}
finally {
    if (Test-Path -LiteralPath $root -PathType Container) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDgdIXHvOuzRxRV
# Yg0e+NznouPnF6yGGHoYTKoP+9is56CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAT3V19sgXXJwyJvZhR8p1S
# tQ93Wgjb13HmYs/HebEq3jANBgkqhkiG9w0BAQEFAASCAYA7S/4eVLQMpqUbbi7R
# 6MleKOdWp7mnwEk75jd0NUMYeyS9p0NPnNUB3zknZNtueS355lMvKu9HCleec2XQ
# WBZjSiAFXRHW7tIrO+lzEzeYnb7ElT2E93E7c1TOu5oMbgys9nLbCoPFZWH3tblD
# U3zvd7NICdZ+2ImOjxSBoGE088eUVlPpcL+8CfgWkv2d2HzTkygLhuNy8Np9eSAj
# b1yFGKV8dqWonRjHm/I88mz4Uza6Rmc2udLZwqgp0kVern0Rr4JFFoct7f+EWrtF
# 3U63lqs+z/NXYbvOZeGlgpro5ITZOBt5vZNQolBuUbMmQl9gaZDSqZLP4c669En/
# f7CF66CB8SaIPaZSzIq4QF78TsB2/d0ygckGa0FTNGVysZ4z/bFhlvQZyvg1P4uS
# mY2dE/0lHav1xB71bmU0roYSNLGqrlFZUd6jh1lX4ljc56ybnBZUGSuuxgC8L5mq
# 0jmMc3ultO8yvAD8kCGAMLp3O/fLIcC1+C/YQdMyM02gUNk=
# SIG # End signature block
