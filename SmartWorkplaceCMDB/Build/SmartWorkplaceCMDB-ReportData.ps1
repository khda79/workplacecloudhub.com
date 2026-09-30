<#
.SYNOPSIS
Builds the current Power BI report-only tables after a complete CMDB build.

.DESCRIPTION
Keeps the 28 collector contract tables unchanged. The 19 report tables and
their validated CI hardware evidence are replaced as one current directory;
no dated report-data history is retained.
#>
[CmdletBinding()]
param(
    [Alias('ProfileKey')][string]$Tenant = 'default',
    [string]$OrganizationKey, [string]$EnvironmentKey, [string]$TenantKey,
    [string]$TenantId, [string]$DataRootPath, [string]$DataAllRootPath,
    [string]$LatestOutputRootPath, [string]$LogRootPath,
    [string]$GlobalConfigPath, [string]$TenantConfigPath,
    [switch]$NoConfigWrite, [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$projectRoot = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1') -Force
$bound = @{}
foreach ($key in $PSBoundParameters.Keys) { $bound[$key] = $PSBoundParameters[$key] }
$context = Resolve-SmartWorkplaceCMDBContext -BoundParameters $bound `
    -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath `
    -NoConfigWrite:($ValidateOnly -or $NoConfigWrite)
$paths = $context.Paths
$reportConfig = $context.Configuration['ReportData']
if ($null -eq $reportConfig -or -not [bool]$reportConfig['Enabled']) {
    throw 'ReportData.Enabled must be true for the configured Full report-data step.'
}
$inventoryRoot = [string]$reportConfig['SmartInventoryLatestOutputRootPath']
if ([string]::IsNullOrWhiteSpace($inventoryRoot)) {
    throw 'ReportData.SmartInventoryLatestOutputRootPath is required.'
}
$inventoryRoot = [IO.Path]::GetFullPath($inventoryRoot)
$localMailboxes = Join-Path $inventoryRoot 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
$remoteMailboxes = Join-Path $inventoryRoot 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv'
$source = [IO.Path]::GetFullPath($paths.LatestOutputRootPath)
$powerBI = Join-Path $source 'PowerBI'
$destination = Join-Path $powerBI 'Report'
$raw = Join-Path $source 'Raw'
$hardware = Join-Path $raw 'Intune\Intune_DeviceHardware.csv'
foreach ($path in @($source, $powerBI, $raw, $hardware, $localMailboxes, $remoteMailboxes)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required report-data input is missing: '$path'." }
}
$python = Get-Command python -CommandType Application -ErrorAction Stop |
    Select-Object -First 1
$versionText = & $python.Source -c 'import sys; print(".".join(map(str, sys.version_info[:3])))'
if ($LASTEXITCODE -ne 0 -or [version]$versionText -lt [version]'3.10') {
    throw 'Python 3.10 or later is required for report-data preparation.'
}
$generator = Join-Path $projectRoot 'PowerBI\prepare_current_report_data.py'
if ($ValidateOnly) {
    [pscustomobject]@{Status='Validated'; PythonVersion=$versionText;
        SmartInventoryRoot=$inventoryRoot; ReportRoot=$destination}
    return
}

$runId = [guid]::NewGuid().ToString('N')
$ciWorkRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartWorkplaceCMDB-CI-' + $runId)
$ciOutput = Join-Path $ciWorkRoot 'CI'
$stage = Join-Path $powerBI ('Report.stage.' + $runId)
$previous = Join-Path $powerBI ('Report.previous.' + $runId)
foreach ($target in @($destination, $stage, $previous)) {
    if ([IO.Path]::GetFullPath((Split-Path $target -Parent)) -ne
        [IO.Path]::GetFullPath($powerBI)) {
        throw "Unsafe report-data target: '$target'."
    }
}
$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
if (-not [IO.Path]::GetFullPath($ciWorkRoot).StartsWith(
        $tempRoot + [IO.Path]::DirectorySeparatorChar,
        [StringComparison]::OrdinalIgnoreCase)) {
    throw "Unsafe CI staging target: '$ciWorkRoot'."
}
$promoted = $false
$movedPrevious = $false
$promotionValidated = $false
try {
    New-Item -ItemType Directory -Path $ciWorkRoot | Out-Null
    Import-Module (Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.CI\SmartWorkplaceCMDB.CI.psm1') -Force
    Export-SmartWorkplaceCMDBCIRegistry -InputRootPath (Join-Path $source 'CMDB') `
        -RawRootPath $raw -HardwareInputPath $hardware -IncludeContext `
        -OrganizationKey $paths.OrganizationKey -EnvironmentKey $paths.EnvironmentKey `
        -TenantKey $paths.TenantKey -TenantId $paths.TenantId `
        -OutputDirectory $ciOutput | Out-Null
    $arguments = @($generator, '--data-root', $source, '--output', $stage,
        '--ci-hardware', (Join-Path $ciOutput 'CMDB_CIDeviceHardware.csv'),
        '--exchange-onprem-local', $localMailboxes,
        '--exchange-onprem-remote', $remoteMailboxes)
    & $python.Source @arguments
    if ($LASTEXITCODE -ne 0) { throw "Report-data generator failed with exit code $LASTEXITCODE." }
    & $python.Source $generator --data-root $source --output $stage --validate-only
    if ($LASTEXITCODE -ne 0) { throw 'Staged report-data validation failed.' }
    if (Test-Path -LiteralPath $destination) {
        $oldManifestPath = Join-Path $destination 'report-data.manifest.json.txt'
        if (-not (Test-Path -LiteralPath $oldManifestPath -PathType Leaf)) {
            throw "Existing Report folder has no CMDB report-data manifest: '$destination'."
        }
        $oldManifest = Get-Content -LiteralPath $oldManifestPath -Raw | ConvertFrom-Json
        if ([string]$oldManifest.identity.TenantKey -cne [string]$paths.TenantKey -or
            @($oldManifest.outputHashes.PSObject.Properties).Count -ne 19) {
            throw 'Existing Report folder is not the expected 19-table tenant snapshot.'
        }
        [IO.Directory]::Move($destination, $previous)
        $movedPrevious = $true
    }
    [IO.Directory]::Move($stage, $destination)
    $promoted = $true
    & $python.Source $generator --data-root $source --output $destination --validate-only
    if ($LASTEXITCODE -ne 0) { throw 'Promoted report-data validation failed.' }
    $promotionValidated = $true
    if ($movedPrevious) {
        try {
            [IO.Directory]::Delete($previous, $true)
            $movedPrevious = $false
        }
        catch {
            Write-Warning ("Validated Report is current, but the old Report directory could not be removed: {0}. {1}" -f
                $previous, $_.Exception.Message)
        }
    }
    [pscustomobject]@{Status='Completed'; ReportRoot=$destination; Tables=19;
        PythonVersion=$versionText}
}
catch {
    if ($promoted -and -not $promotionValidated -and $movedPrevious -and
        (Test-Path -LiteralPath $destination -PathType Container)) {
        $failed = Join-Path $powerBI ('Report.failed.' + $runId)
        [IO.Directory]::Move($destination, $failed)
        [IO.Directory]::Move($previous, $destination)
    }
    throw
}
finally {
    if (Test-Path -LiteralPath $stage -PathType Container) {
        try { [IO.Directory]::Delete($stage, $true) }
        catch { Write-Warning ("Temporary report stage could not be removed: {0}" -f $stage) }
    }
    if (Test-Path -LiteralPath $ciWorkRoot -PathType Container) {
        try { [IO.Directory]::Delete($ciWorkRoot, $true) }
        catch { Write-Warning ("Temporary CI stage could not be removed: {0}" -f $ciWorkRoot) }
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBLNz/yItKRT0zH
# Jz965n6EsAML878HDnb98YEBhzudhqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDCuKDCKfs20lZMTc/woLl/
# Ih38TjsqculbRw0q/agbpjANBgkqhkiG9w0BAQEFAASCAYAYsIP3SiiwuyVc642h
# OHtYGwejndDuiZxov1MSzofqLurBzGf1BZd3CMx9t94vkrKO5wanLgOh2fcpUtOC
# PKYidUtgfL0ENvYD8rGPEIgAWNM1gNHkKYFBL/Urnk7UGRzfn32iMp9sILIM//7w
# Hgn/Df/nHyFo30lc8MDz3iLa2n4o++ZHZtdm8nqkvuCbJWErjjYxVAXNEhivGB/b
# wU+vzHWe5Jxs1pOVaxtc9IoUcVWIHdPDxslicgaVYhkd79G2UpshzKN17MlU24wG
# 408HU3FPS8d7fdCzhl002Aqvl4IcWj6afhqHeOhGA3JAb4znW+dodCdf1Z4JH3YR
# 6S86nrSDgnEIJyPxyLoikrhb3flnqUoSAIYUjgFdV2pMrS1cwNjv76HWJLYEoqW1
# +1VDZCURf0i5gzw/iv1PwrvmIzjkB+plDsWcIXCd/5S7yuFh8ERIzH6HMzg+wJJv
# WBPXoThda5ZhWp4fBzsMgBI4swXYfA1i06hYq4axN6mUW5s=
# SIG # End signature block
