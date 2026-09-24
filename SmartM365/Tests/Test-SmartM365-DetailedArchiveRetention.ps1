<#
.SYNOPSIS
Offline tests for the seven-day detailed archive retention policy.
.VERSION
1.1.0
#>
#requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$referenceTime = [datetime]'2026-09-23T12:00:00'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Import-FunctionFromScript {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Name
    )

    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    Assert-True ($parseErrors.Count -eq 0) "PowerShell parse errors in $Path"
    $definition = @($ast.FindAll({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name
    }, $true))[0]
    Assert-True ($null -ne $definition) "Function $Name not found in $Path"
    $bodyText = $definition.Body.Extent.Text
    return [scriptblock]::Create($bodyText.Substring(1, $bodyText.Length - 2))
}

function Log { param([string]$Message) }
function Warn { param([string]$Message) }
function Write-Log { param([string]$Message, [string]$Level, [string]$Stage) }
function Test-FileLocked { param([string]$Path) return $false }

$cases = @(
    [pscustomobject]@{
        Name = 'Exchange HybridIdentity Issues'
        Path = Join-Path $smartM365Root 'SmartInventory\ExchangeInventory\Migration\SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1'
        Prefix = 'Exchange_HybridIdentity_Issues'
        Format = 'yyyyMMdd-HHmmss'
        Version = '1.17'
        ArchiveWriteMarker = 'CopyCsv $main $archive'
        RetentionCallMarker = "Remove-DetailedArchiveFilesOlderThan -Folder (Join-Path `$OutputFolder 'Archive')"
    },
    [pscustomobject]@{
        Name = 'Windows11 Readiness Issues'
        Path = Join-Path $smartM365Root 'SmartInventory\M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1'
        Prefix = 'Intune_Windows11_Readiness_Issues'
        Format = 'yyyyMMdd-HHmmss'
        Version = '1.23'
        ArchiveWriteMarker = 'CopyCsv $main $archive'
        RetentionCallMarker = "Remove-DetailedArchiveFilesOlderThan -Folder (Join-Path `$OutputFolder 'Archive')"
    },
    [pscustomobject]@{
        Name = 'Windows Update Status'
        Path = Join-Path $smartM365Root 'SmartInventory\M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-WinUpdate_Status_From_Intune.ps1'
        Prefix = 'Intune_WindowsUpdate_Status'
        Format = 'yyyyMMdd_HHmmss'
        Version = '1.41'
        ArchiveWriteMarker = 'Copy-FileAtomic -SourcePath $SourceCsv -DestinationFinal $archiveFinal'
        RetentionCallMarker = 'Remove-DetailedArchiveFilesOlderThan -Folder $ArchiveFolder'
    }
)

foreach ($case in $cases) {
    $source = Get-Content -LiteralPath $case.Path -Raw
    Assert-True ($source -match ('(?m)^\s*\.VERSION\s*\r?\n' + [regex]::Escape($case.Version) + '\s*$')) "$($case.Name) metadata version"
    Assert-True ($source -match ('\$ScriptVersion\s*=\s*["'']' + [regex]::Escape($case.Version) + '["'']')) "$($case.Name) runtime version"
    Assert-True ($source -match '\[int\]\$DetailedArchiveRetentionDays\s*=\s*7') "$($case.Name) default retention is seven days"
    $localConfigTemplatePath = $case.Path -replace '\.ps1$', '.local.json.template'
    $weeklyHistoryEvidence = $source
    if (Test-Path -LiteralPath $localConfigTemplatePath) { $weeklyHistoryEvidence += Get-Content -LiteralPath $localConfigTemplatePath -Raw }
    Assert-True ($weeklyHistoryEvidence.Contains('WeeklyHistory')) "$($case.Name) weekly history remains configured"

    $writeIndex = $source.IndexOf($case.ArchiveWriteMarker, [StringComparison]::Ordinal)
    $retentionIndex = $source.IndexOf($case.RetentionCallMarker, [StringComparison]::Ordinal)
    Assert-True ($writeIndex -ge 0) "$($case.Name) archive write marker"
    Assert-True ($retentionIndex -gt $writeIndex) "$($case.Name) retention runs only after archive publication"

    $retentionFunction = Import-FunctionFromScript -Path $case.Path -Name 'Remove-DetailedArchiveFilesOlderThan'
    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ArchiveRetention-' + [guid]::NewGuid().ToString('N'))
    $archiveFolder = Join-Path $tempRoot 'Archive'
    $weeklyFolder = Join-Path $tempRoot 'WeeklyHistory\2026-W39'
    New-Item -ItemType Directory -Path $archiveFolder,$weeklyFolder -Force | Out-Null

    try {
        $oldStamp = $referenceTime.AddDays(-7).AddSeconds(-1).ToString($case.Format, [Globalization.CultureInfo]::InvariantCulture)
        $cutoffStamp = $referenceTime.AddDays(-7).ToString($case.Format, [Globalization.CultureInfo]::InvariantCulture)
        $freshStamp = $referenceTime.AddHours(-1).ToString($case.Format, [Globalization.CultureInfo]::InvariantCulture)
        $oldPath = Join-Path $archiveFolder ("{0}_{1}.csv" -f $case.Prefix,$oldStamp)
        $cutoffPath = Join-Path $archiveFolder ("{0}_{1}.csv" -f $case.Prefix,$cutoffStamp)
        $freshPath = Join-Path $archiveFolder ("{0}_{1}.csv" -f $case.Prefix,$freshStamp)
        $unrelatedPath = Join-Path $archiveFolder ("Unrelated_{0}.csv" -f $oldStamp)
        $malformedPath = Join-Path $archiveFolder ("{0}_20260901.csv" -f $case.Prefix)
        $weeklyPath = Join-Path $weeklyFolder ("{0}_{1}.csv" -f $case.Prefix,$oldStamp)
        foreach ($path in @($oldPath,$cutoffPath,$freshPath,$unrelatedPath,$malformedPath,$weeklyPath)) {
            Set-Content -LiteralPath $path -Value 'test' -Encoding utf8
        }

        & $retentionFunction -Folder $archiveFolder -BaseNameWithoutExt $case.Prefix -RetentionDays 7 -ReferenceTime $referenceTime

        Assert-True (-not (Test-Path -LiteralPath $oldPath)) "$($case.Name) removes a matching snapshot older than seven days"
        Assert-True (Test-Path -LiteralPath $cutoffPath) "$($case.Name) retains a snapshot exactly at the cutoff"
        Assert-True (Test-Path -LiteralPath $freshPath) "$($case.Name) retains a fresh snapshot"
        Assert-True (Test-Path -LiteralPath $unrelatedPath) "$($case.Name) retains another collector's snapshot"
        Assert-True (Test-Path -LiteralPath $malformedPath) "$($case.Name) retains a malformed timestamp instead of guessing"
        Assert-True (Test-Path -LiteralPath $weeklyPath) "$($case.Name) does not recurse into WeeklyHistory"
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }
}

$statusSource = Get-Content -LiteralPath $cases[2].Path -Raw
Assert-True (-not $statusSource.Contains('$global:RetentionMaxCSV')) 'Windows Update Status no longer applies count-based CSV archive retention'
Assert-True ($statusSource.Contains('Prune-Files -Folder $LogsPath')) 'Windows Update Status keeps the independent log-count retention'

$sharedModuleCases = @(
    [pscustomobject]@{
        Name = 'SmartM365.Core'
        Path = Join-Path $smartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psm1'
        ManifestPath = Join-Path $smartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        ExpectedVersion = '1.0.57'
        ExplicitExport = $false
    },
    [pscustomobject]@{
        Name = 'SmartM365 Windows PowerShell 5 compatibility module'
        Path = Join-Path $smartM365Root 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psm1'
        ManifestPath = Join-Path $smartM365Root 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
        ExpectedVersion = '1.0.41'
        ExplicitExport = $true
    }
)

foreach ($moduleCase in $sharedModuleCases) {
    $manifest = Import-PowerShellDataFile -LiteralPath $moduleCase.ManifestPath
    Assert-True ([string]$manifest.ModuleVersion -eq $moduleCase.ExpectedVersion) "$($moduleCase.Name) version"
    if ($moduleCase.ExplicitExport) {
        Assert-True ($manifest.FunctionsToExport -contains 'Remove-SmartM365TimestampedFilesOlderThan') "$($moduleCase.Name) exports the timestamp retention helper"
        Assert-True ($manifest.FunctionsToExport -contains 'Remove-SmartM365TimestampedDirectoriesOlderThan') "$($moduleCase.Name) exports the timestamp directory retention helper"
    }

    $retentionFunction = Import-FunctionFromScript -Path $moduleCase.Path -Name 'Remove-SmartM365TimestampedFilesOlderThan'
    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SharedArchiveRetention-' + [guid]::NewGuid().ToString('N'))
    $archiveFolder = Join-Path $tempRoot 'Archive'
    $weeklyFolder = Join-Path $archiveFolder 'WeeklyHistory\2026-W39'
    New-Item -ItemType Directory -Path $archiveFolder,$weeklyFolder -Force | Out-Null

    try {
        $global:csvGeneratedPaths = @()
        $paths = [ordered]@{
            OldUnderscoreSeconds = Join-Path $archiveFolder 'Collector_20260916_115959.csv'
            OldHyphenSeconds = Join-Path $archiveFolder 'Collector_20260916-115959.csv'
            OldUnderscoreMinutes = Join-Path $archiveFolder 'Collector_20260916_1159.csv'
            OldHyphenMinutes = Join-Path $archiveFolder 'Collector_20260916-1159.csv'
            ExactCutoff = Join-Path $archiveFolder 'Collector_20260916_120000.csv'
            Fresh = Join-Path $archiveFolder 'Collector_20260923-110000.csv'
            Canonical = Join-Path $archiveFolder 'Collector.csv'
            InvalidDate = Join-Path $archiveFolder 'Collector_20261340_250000.csv'
            History = Join-Path $archiveFolder 'Collector_History_20260901_000000.csv'
            Cache = Join-Path $archiveFolder 'Collector_Cache_20260901_000000.csv'
            Resume = Join-Path $archiveFolder 'Collector_Resume_20260901_000000.csv'
            Excluded = Join-Path $archiveFolder 'Collector_20260901_000000.csv'
            Weekly = Join-Path $weeklyFolder 'Collector_20260901_000000.csv'
        }
        foreach ($path in $paths.Values) { Set-Content -LiteralPath $path -Value 'test' -Encoding utf8 }

        & $retentionFunction -FolderPath $archiveFolder -FilePattern '*.csv' -RetentionDays 7 -ExcludeFiles @($paths.Excluded) -ReferenceTime $referenceTime

        foreach ($key in @('OldUnderscoreSeconds','OldHyphenSeconds','OldUnderscoreMinutes','OldHyphenMinutes')) {
            Assert-True (-not (Test-Path -LiteralPath $paths[$key])) "$($moduleCase.Name) removes $key"
        }
        foreach ($key in @('ExactCutoff','Fresh','Canonical','InvalidDate','History','Cache','Resume','Excluded','Weekly')) {
            Assert-True (Test-Path -LiteralPath $paths[$key]) "$($moduleCase.Name) preserves $key"
        }

        $guardedOldPath = Join-Path $archiveFolder 'Collector_Guarded_20260901_000000.csv'
        Set-Content -LiteralPath $guardedOldPath -Value 'test' -Encoding utf8
        & $retentionFunction -FolderPath $archiveFolder -FilePattern '*.csv' -RetentionDays 7 -RequireCurrentRunPublication -ReferenceTime $referenceTime
        Assert-True (Test-Path -LiteralPath $guardedOldPath) "$($moduleCase.Name) skips failed-run cleanup without publication evidence"

        $currentPublicationPath = Join-Path $archiveFolder 'Collector_Current.csv'
        Set-Content -LiteralPath $currentPublicationPath -Value 'test' -Encoding utf8
        $global:csvGeneratedPaths = @($currentPublicationPath)
        & $retentionFunction -FolderPath $archiveFolder -FilePattern '*.csv' -RetentionDays 7 -RequireCurrentRunPublication -ReferenceTime $referenceTime
        Assert-True (-not (Test-Path -LiteralPath $guardedOldPath)) "$($moduleCase.Name) cleans old files after current-run publication evidence"

        $directoryRetentionFunction = Import-FunctionFromScript -Path $moduleCase.Path -Name 'Remove-SmartM365TimestampedDirectoriesOlderThan'
        $directoryRoot = Join-Path $tempRoot 'RunDirectories'
        $runDirectories = [ordered]@{
            Old = Join-Path $directoryRoot '20260916-115959'
            OldWithCollisionSuffix = Join-Path $directoryRoot '20260916-115959-2'
            ExactCutoff = Join-Path $directoryRoot '20260916-120000'
            Fresh = Join-Path $directoryRoot '20260923_110000'
            Invalid = Join-Path $directoryRoot '20261340-250000'
            WeeklyHistory = Join-Path $directoryRoot 'WeeklyHistory'
            Excluded = Join-Path $directoryRoot '20260901-000000'
        }
        foreach ($directoryPath in $runDirectories.Values) {
            New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $directoryPath 'evidence.txt') -Value 'test' -Encoding utf8
        }

        & $directoryRetentionFunction -RootPath $directoryRoot -RetentionDays 7 -ExcludeDirectories @($runDirectories.Excluded) -ReferenceTime $referenceTime

        foreach ($key in @('Old','OldWithCollisionSuffix')) {
            Assert-True (-not (Test-Path -LiteralPath $runDirectories[$key])) "$($moduleCase.Name) removes timestamped directory $key"
        }
        foreach ($key in @('ExactCutoff','Fresh','Invalid','WeeklyHistory','Excluded')) {
            Assert-True (Test-Path -LiteralPath $runDirectories[$key]) "$($moduleCase.Name) preserves timestamped directory $key"
        }
    }
    finally {
        $global:csvGeneratedPaths = @()
        if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }
}

$inventoryRoot = Join-Path $smartM365Root 'SmartInventory'
$additionalRetentionCollectors = @(
    'ActiveDirectoryInventory\SmartM365-ActiveDirectory-HealthCheck.ps1',
    'ExchangeInventory\BackupProtection\SmartM365-M365-BackupPolicyScope-Inventory.ps1',
    'ExchangeInventory\OnPremises\ServersAndStorage\SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1',
    'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-Compliance-Inventory.ps1',
    'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-UpgradeEligibility.ps1',
    'M365Inventory\IntuneInventory\EndpointAnalytics\SmartM365-EndpointAnalytics-Inventory.ps1',
    'M365Inventory\IntuneInventory\SmartM365-Intune-ExportRemediationScripts.ps1',
    'M365Inventory\IntuneInventory\WindowsUpdate\AutopatchAlerts\SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1',
    'M365Inventory\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1'
)
foreach ($relativePath in $additionalRetentionCollectors) {
    $source = Get-Content -LiteralPath (Join-Path $inventoryRoot $relativePath) -Raw
    Assert-True ($source -match 'Remove-(?:Core)?SmartM365Timestamped(?:Files|Directories)OlderThan') "$relativePath applies seven-day timestamp retention"
}

$legacyDataRetentionCalls = foreach ($scriptPath in Get-ChildItem -LiteralPath $inventoryRoot -Filter '*.ps1' -Recurse) {
    $lineNumber = 0
    foreach ($line in Get-Content -LiteralPath $scriptPath.FullName) {
        $lineNumber++
        if ($line -match '(?:RemoveOldFiles|Remove-OldFiles).*?(?:\.csv|\.xlsx|RetentionMaxCSV)' -and
            $line -notmatch 'Remove-SmartM365TimestampedFilesOlderThan') {
            '{0}:{1}: {2}' -f $scriptPath.FullName,$lineNumber,$line.Trim()
        }
    }
}
Assert-True (@($legacyDataRetentionCalls).Count -eq 0) ("No collector keeps count-based CSV/XLSX retention. Found: {0}" -f (@($legacyDataRetentionCalls) -join ' | '))

Write-Host "Detailed archive retention tests passed for $($cases.Count) dedicated collectors and $($sharedModuleCases.Count) shared modules." -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBG1MUz4fNjYAco
# /sAKo6mQ3c8EBmf5kKUjec0zp6gD56CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDsez1PvQyajh39HQvMPakQA4rdszMUDKNlkBsQdXFEZMA0GCSqG
# SIb3DQEBAQUABIIBgC2JBqOrXhuAa0kdnAalHt70emmTDVm0UBv6UzeaJhFkN1NU
# pHbNITVFm8eW4RJK7IbSP7EK+7bVHusJpXkRLRb5gw75Aa7zqv5bShn1xJwJAAwi
# 0ktyN6oUSW5pFL7RkFjieCcjKRKlmYJ5K7Hu9kVUB/XNE+fVSDxK45ZzbpQ9WKcJ
# e0VJ1pY07rHdoQoeUriJEBV8eYSnTGMThF3mpkh2jUiwd7L9fDlCl5DVdCKu0TVb
# r3Dq6r+DX5ZKemNvqEsTwNYcZgnqSxkogYDozPqHpG9M6iEU28rQ7ZYjjZkDDFEW
# 1zfPnkNVqwFfEMA3Uq15NhAPfBTC1XojWU+NuIEl0p66wqDlmZgO3M9w7aep7cWu
# fRsAdxnenTwuf+StJVUKQGfcPINM6+p/w0kIcpb0Out0jU1zAdKtbm8bQo5BbKOH
# ur9tHEZW7ISXq5s4IF1FqKAhARzWzJvjsYsldvgzHqVZ1tkZ/nbzE7zJps28CiQ0
# J+QfGnFnaUv00hx1XKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjQwODQz
# MjVaMC8GCSqGSIb3DQEJBDEiBCCiN16+2nN7RObMSNUhXvCH5pTibJjRCpazXeT+
# lTKRbjANBgkqhkiG9w0BAQEFAASCAgB8vav63ifoMv+Ouw3oyVVuciEEEVX3c1RA
# TTe5k6QQqILzkYrxQferF0wNvqZ4r2TnuOfQTDDBhy/EUV3lkRD9SujV7JFz3hHr
# /tB9auFznFqIEZzDUcQ4V8Q5EszK1K/sUg0lA18GYHbVo9yXfYNKL47dZvuBLuI4
# tVdFu4FaYH1+EEeACPY+jhnoItlBjBvAjw8JQRAxt5UmLBsaEDlZ49pqoiuF7/No
# aneLP1ecip5DYQ6YwrVgyYWsO62CdtRBJYWAe/roTw0gD/ktC/HVdwHZReeH8L+V
# FCky9KFC9wsszG6JJLYR59xwJCyIuEEyi8MMjY/ceoWevYILH9VjmhDBLJuzZMk+
# QEQMiJOyV3mCmPxbF6K360UTKIbawxwVo+TNBQK6xSojYUb+XrzsG2YHX1ReROHQ
# Y0JIX2PanXGFYLKt9AStwXH2g++eRH16qepjYJl8gIwj+JAu/QlAB6UdQxP1GsQx
# nQ3TAbrvOVe2OfPmwXFZyGuWlQJvkTdncEc9ryd7iN53QdWMMORSz4Ni7pjXWIgu
# Pl5kJIRaEwbN++dbMgxaFwN7eyLCdz0NGPND+i6mQgD/NpiJ+amHOos4TdfkzyYC
# sPpRgvYIG3UGsVz4VZn0dvD918ihKxPxglnjcXU6t6pIhCk6ZvWjDeSJ3//YwW5J
# F+zrl6DHog==
# SIG # End signature block
