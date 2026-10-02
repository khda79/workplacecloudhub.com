<#
.SYNOPSIS
    Runs offline launcher and dashboard regression checks using synthetic scans.
.VERSION
    1.0.0
#>
[CmdletBinding()]
param()
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
        'Get-PathMappingRows', 'Get-ComparisonConfigValue', 'Assert-CsvScanAgeDifference', 'Invoke-PermissionComparison'
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC2/w0D9fx7TScY
# +g7jFnyJ9hD5Sk6wPIp+vVe5hcvb2aCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD+ppLU6HkIiq2PVKX7KGk3
# F410xv8gUb3DB3oZ7Eto1DANBgkqhkiG9w0BAQEFAASCAYAkK8mHXcUQKIP9NMkY
# /JPNnf8of5WIcxb9EyEVuBSImeHNdeAUdylseH8ywTHj3FCGqWnvQ6iqmJg2UlLv
# IH5Z16r2PZQXa6Qw4JIM4U416+V8Gc9hovnd3HFcl/VcZKC4aJHCZjXwHZG0mJ3c
# AcAxL+8TliV+Xox22SN3aUFXyngISixDiHX94ZbX1FSBl/5thAXcx+/epVJq4Up9
# XNtYo85ZJ5Nvv2fPQET4dX7iWyBnC8kHtix0g4aiTyUep5q126FCRBSvE26/firw
# n0UBg0u0exEc81mjmqppS+EUEov+ELp8DfnaSlkIc9K+OwNDz8nEo4Q6g7NbE5nc
# nTGmoDMyMxp+Vjj6jE4XehiWssXONLU/J5hPfIxH6dhlJASYU92e5zv/uubWWC+Q
# Zj/Seks4oFLCVlrTryEZwCpl+Vo01ErlSxlUVg9elYHiJvUqpkutX049obZGC/R2
# qEOSCQRfIgZTTbGlt/xsdAmARdCiCsHxiiboLKuaDQYP5Ek=
# SIG # End signature block
