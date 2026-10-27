<#
.SYNOPSIS
Offline tests for the seven-day detailed archive retention policy.
.VERSION
1.1.4
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
function Test-SmartM365InterruptedRun { return $false }

$cases = @(
    [pscustomobject]@{
        Name = 'Windows Update Status'
        Path = Join-Path $smartM365Root 'SmartInventory\M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-WinUpdate_Status_From_Intune.ps1'
        Prefix = 'Intune_WindowsUpdate_Status'
        Format = 'yyyyMMdd_HHmmss'
        MinimumVersion = '1.48'
        ArchiveWriteMarker = 'Copy-FileAtomic -SourcePath $SourceCsv -DestinationFinal $archiveFinal'
        RetentionCallMarker = 'Remove-DetailedArchiveFilesOlderThan -Folder $ArchiveFolder'
    }
)

foreach ($case in $cases) {
    $source = Get-Content -LiteralPath $case.Path -Raw
    # Version floor: later releases keep the retention contract; header and runtime versions must agree.
    $headerVersion = [regex]::Match($source, '(?m)^\s*\.VERSION\s*\r?\n\s*([0-9]+(?:\.[0-9]+)+)\s*$').Groups[1].Value
    Assert-True ($headerVersion -and [version]$headerVersion -ge [version]$case.MinimumVersion) "$($case.Name) metadata version"
    Assert-True ($source -match ('\$ScriptVersion\s*=\s*["'']' + [regex]::Escape($headerVersion) + '["'']')) "$($case.Name) runtime version"
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

$statusSource = Get-Content -LiteralPath $cases[0].Path -Raw
Assert-True (-not $statusSource.Contains('$global:RetentionMaxCSV')) 'Windows Update Status no longer applies count-based CSV archive retention'
Assert-True ($statusSource.Contains('Prune-Files -Folder $LogsPath')) 'Windows Update Status keeps the independent log-count retention'

$sharedModuleCases = @(
    [pscustomobject]@{
        Name = 'SmartM365.Core'
        Path = Join-Path $smartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psm1'
        ManifestPath = Join-Path $smartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        MinimumVersion = '1.0.88'
        ExplicitExport = $false
    },
    [pscustomobject]@{
        Name = 'SmartM365 Windows PowerShell 5 compatibility module'
        Path = Join-Path $smartM365Root 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psm1'
        ManifestPath = Join-Path $smartM365Root 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'
        MinimumVersion = '1.0.58'
        ExplicitExport = $true
    }
)

foreach ($moduleCase in $sharedModuleCases) {
    $manifest = Import-PowerShellDataFile -LiteralPath $moduleCase.ManifestPath
    Assert-True ([version]$manifest.ModuleVersion -ge [version]$moduleCase.MinimumVersion) "$($moduleCase.Name) version"
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
