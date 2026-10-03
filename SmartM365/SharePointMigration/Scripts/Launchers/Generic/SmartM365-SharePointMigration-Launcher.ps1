<#
.SYNOPSIS
    Generic launcher for a migration folder.

.DESCRIPTION
    Reads Migrations\<MigrationName>\migration.config.psd1 and runs the
    requested inventory, comparison, or permission action.

.VERSION
    1.0.23
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$MigrationName,

    [Parameter(Mandatory = $true)]
    [ValidateSet(
        'ScanSourceFiles',
        'ScanTargetFiles',
        'CompareFiles',
        'ScanSourcePermissions',
        'ScanTargetPermissions',
        'ComparePermissions',
        'CompareSourceHistory',
        'CompareScanHistory',
        'ComparePermissionScanHistory'
    )]
    [string]$Action,

    [string]$SourceCsv,

    [string]$TargetCsv,

    [string]$OldCsv,

    [string]$NewCsv,

    [ValidateSet('Source', 'Target')]
    [string]$HistorySide = 'Source',

    [switch]$ForceAuthentication = $true,

    [switch]$DeviceLogin,

    [switch]$UseCertificate,

    [switch]$Force,

    [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'
$script:LauncherPreviousConsoleMarker = [string]$env:SPMIG_CONSOLE_LIFECYCLE_ACTIVE
$script:LauncherOwnsLifecycle = [string]::IsNullOrWhiteSpace($script:LauncherPreviousConsoleMarker)
if ($script:LauncherOwnsLifecycle) {
    Microsoft.PowerShell.Utility\Write-Host ('{0} Script  : {1} v{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $MyInvocation.MyCommand.Name, '1.0.23') -ForegroundColor Cyan
}
$script:LauncherNonInteractive = [bool]$NonInteractive
. (Join-Path -Path $PSScriptRoot -ChildPath '..\SmartM365-SharePointMigration-LauncherCommon.ps1')

$ProjectRoot = [System.IO.Path]::GetFullPath((Get-LauncherProjectRoot -LauncherScriptRoot $PSScriptRoot))
$Migration = Import-LauncherMigrationConfig -ProjectRoot $ProjectRoot -MigrationName $MigrationName
$MigrationRoot = $Migration.Root
$Config = $Migration.Data

if ($Config.Name -eq 'NewMigration') {
    throw "The migration template cannot be executed directly. Copy Migrations_Template to Migrations\<MigrationName>, update migration.config.psd1, then run the launcher from the copied migration."
}

function Get-ConsoleTimestamp {
    Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
}

function Write-Info {
    param(
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    Microsoft.PowerShell.Utility\Write-Host ("{0} {1}" -f (Get-ConsoleTimestamp), $Message) -ForegroundColor $Color
}

function Write-LauncherIntro {
    if (-not $script:LauncherOwnsLifecycle) { return }
    $bannerLines = @(
        '================================================================================',
        ' SmartM365 by WorkplaceCloudHub',
        ' Website : https://workplacecloudhub.com',
        ' GitHub  : https://github.com/khda79/workplacecloudhub.com',
        '================================================================================'
    )
    $colors = @('DarkCyan', 'Cyan', 'Yellow', 'Yellow', 'DarkCyan')
    for ($index = 0; $index -lt $bannerLines.Count; $index++) {
        Microsoft.PowerShell.Utility\Write-Host $bannerLines[$index] -ForegroundColor $colors[$index]
    }
    Write-Info ('Migration : {0}' -f $MigrationName) Cyan
    Write-Info ('Action    : {0}' -f $Action) Cyan
    Write-Info ('Started   : {0}' -f $script:LauncherRunStartedAt.ToString('yyyy-MM-dd HH:mm:ss')) Cyan
}

function Write-LauncherSummary {
    param([AllowNull()]$Failure)

    if (-not $script:LauncherOwnsLifecycle) { return }

    $endedAt = Get-Date
    $duration = $endedAt - $script:LauncherRunStartedAt
    $durationText = '{0:00}:{1:00}:{2:00}' -f ([int][Math]::Floor($duration.TotalHours)), $duration.Minutes, $duration.Seconds
    $failed = $null -ne $Failure
    $status = if ($failed) { 'FAILED' } else { 'SUCCESS' }
    $color = if ($failed) { [ConsoleColor]::Red } else { [ConsoleColor]::Green }
    $lines = @(
        '================================================================================',
        ('SmartM365 by WorkplaceCloudHub - {0}' -f $(if ($failed) { 'Execution failed' } else { 'Execution completed' })),
        ('Migration : {0}' -f $MigrationName),
        ('Action    : {0}' -f $Action),
        ('Status    : {0}' -f $status),
        ('Duration  : {0}' -f $durationText)
    )
    if (-not [string]::IsNullOrWhiteSpace($script:LauncherRunLogPath)) {
        $lines += ('Log       : {0}' -f $script:LauncherRunLogPath)
    }
    if ($failed) {
        $lines += ('Reason    : {0}' -f $Failure.Exception.Message)
    }
    elseif ($Action -like 'Compare*') {
        $lines += 'Result    : Comparison generated; review the report for migration differences.'
    }
    $lines += '================================================================================'

    $timestampedLines = foreach ($line in $lines) {
        '{0} {1}' -f (Get-ConsoleTimestamp), $line
    }
    for ($index = 0; $index -lt $timestampedLines.Count; $index++) {
        $lineColor = if ($lines[$index] -like 'Status*') { $color } elseif ($index -eq 0 -or $index -eq ($lines.Count - 1)) { [ConsoleColor]::DarkCyan } else { [ConsoleColor]::Gray }
        Microsoft.PowerShell.Utility\Write-Host $timestampedLines[$index] -ForegroundColor $lineColor
    }
    if (-not [string]::IsNullOrWhiteSpace($script:LauncherRunLogPath) -and
        (Test-Path -LiteralPath $script:LauncherRunLogPath -PathType Leaf)) {
        try {
            Add-Content -LiteralPath $script:LauncherRunLogPath -Value $timestampedLines -Encoding UTF8
        }
        catch {
            Write-Info ('Could not append launcher summary to run log: {0}' -f $_.Exception.Message) Yellow
        }
    }
}

function Get-PythonCommand {
    $portablePython = Join-Path -Path $ProjectRoot -ChildPath 'Tools\Python\python.exe'
    if (Test-Path -LiteralPath $portablePython -PathType Leaf) {
        & $portablePython --version *> $null
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{ Executable = $portablePython; Arguments = @(); Source = 'Portable' }
        }
    }

    $python = Get-Command -Name python -ErrorAction SilentlyContinue
    if ($python) {
        & $python.Source --version *> $null
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{ Executable = $python.Source; Arguments = @(); Source = 'PATH' }
        }
    }

    $pyLauncher = Get-Command -Name py -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        & $pyLauncher.Source -3 --version *> $null
        if ($LASTEXITCODE -eq 0) {
            return [pscustomobject]@{ Executable = $pyLauncher.Source; Arguments = @('-3'); Source = 'PyLauncher' }
        }
    }

    throw "Python 3 is required. Expected portable Python at: $portablePython"
}

function Invoke-PythonScript {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Script,

        [string[]]$Arguments = @()
    )

    if (-not (Test-Path -LiteralPath $Script -PathType Leaf)) {
        throw "Python script not found: $Script"
    }

    $pythonCommand = Get-PythonCommand
    Write-Info ("Running Python script: {0} ({1})" -f $Script, $pythonCommand.Source) Cyan

    & $pythonCommand.Executable @($pythonCommand.Arguments) $Script @Arguments 2>&1 | ForEach-Object {
        Microsoft.PowerShell.Utility\Write-Host ([string]$_)
    }

    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw ("Python script failed with exit code {0}: {1}" -f $exitCode, $Script)
    }
}
function Resolve-MigrationPath {
    param([string]$Path)
    Resolve-LauncherMigrationPath -MigrationRoot $MigrationRoot -Path $Path
}

function Get-MigrationLogDirectory {
    $logsPath = 'logs'
    if ($Config.ContainsKey('Output') -and $Config.Output.ContainsKey('Logs') -and -not [string]::IsNullOrWhiteSpace([string]$Config.Output.Logs)) {
        $logsPath = [string]$Config.Output.Logs
    }

    $logDirectory = Resolve-MigrationPath $logsPath
    if (-not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    }

    $logDirectory
}

function New-MigrationRunTimestamp {
    $runId = if ($env:SPMIG_GUI_RUN_ID -match '^[0-9a-fA-F]{32}$') {
        $env:SPMIG_GUI_RUN_ID.ToLowerInvariant()
    } else { [guid]::NewGuid().ToString('N') }
    '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $runId
}

function New-MigrationLogPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action,

        [string]$Timestamp = (New-MigrationRunTimestamp)
    )

    $path = Join-Path -Path (Get-MigrationLogDirectory) -ChildPath ("{0}-{1}-{2}.log" -f $Config.Name, $Action, $Timestamp)
    if ($Action -eq $script:Action) {
        $script:LauncherRunLogPath = $path
    }
    $path
}

function Add-TimestampToLogFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $timestampPattern = '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} '
    $content = Get-Content -LiteralPath $Path
    $updated = foreach ($line in $content) {
        if ($line -match $timestampPattern -or [string]::IsNullOrWhiteSpace($line)) {
            $line
        }
        else {
            "{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line
        }
    }

    Set-Content -LiteralPath $Path -Value $updated -Encoding UTF8 -WhatIf:$false
}

function Stop-LauncherTranscript {
    param([string]$Path)

    try {
        Stop-Transcript | Out-Null
    }
    catch {
        return
    }

    Add-TimestampToLogFile -Path $Path
}

function Open-DirectoryInExplorer {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        $resolvedPath = [string](Resolve-Path -LiteralPath $Path).ProviderPath
        $resolvedPath = ($resolvedPath -replace ([string][char]0), '').Trim()
        $explorerPath = $resolvedPath
        $root = [System.IO.Path]::GetPathRoot($resolvedPath)

        if ($root -match '^[A-Za-z]:\\$') {
            $driveName = $root.Substring(0, 1)
            $drive = Get-PSDrive -Name $driveName -ErrorAction SilentlyContinue
            $uncRoot = $null
            if ($drive -and $drive.DisplayRoot -and $drive.DisplayRoot.StartsWith('\\')) {
                $uncRoot = $drive.DisplayRoot
            }
            else {
                try {
                    $logicalDisk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}:'" -f $driveName) -ErrorAction Stop
                    if ($logicalDisk.ProviderName -and $logicalDisk.ProviderName.StartsWith('\\')) {
                        $uncRoot = $logicalDisk.ProviderName
                    }
                }
                catch {
                    $logicalDisk = Get-WmiObject -Class Win32_LogicalDisk -Filter ("DeviceID='{0}:'" -f $driveName) -ErrorAction SilentlyContinue
                    if ($logicalDisk -and $logicalDisk.ProviderName -and $logicalDisk.ProviderName.StartsWith('\\')) {
                        $uncRoot = $logicalDisk.ProviderName
                    }
                }
            }

            if ($uncRoot) {
                $uncRoot = (([string]$uncRoot) -replace ([string][char]0), '').Trim()
                $relativePath = $resolvedPath.Substring($root.Length).TrimStart('\')
                $candidatePaths = New-Object 'System.Collections.Generic.List[string]'
                $candidatePaths.Add((Join-Path -Path $uncRoot -ChildPath $relativePath))

                $uncLeaf = Split-Path -Path $uncRoot.TrimEnd('\') -Leaf
                if (-not [string]::IsNullOrWhiteSpace($uncLeaf) -and $relativePath.StartsWith(("{0}\" -f $uncLeaf), [System.StringComparison]::OrdinalIgnoreCase)) {
                    $candidatePaths.Add((Join-Path -Path $uncRoot -ChildPath $relativePath.Substring($uncLeaf.Length + 1)))
                }

                foreach ($candidatePath in $candidatePaths) {
                    $cleanCandidatePath = (([string]$candidatePath) -replace ([string][char]0), '').Trim()
                    if (Test-Path -LiteralPath $cleanCandidatePath -PathType Container) {
                        $explorerPath = $cleanCandidatePath
                        break
                    }
                }
            }
        }

        $explorerPath = (([string]$explorerPath) -replace ([string][char]0), '').Trim()
        if (-not (Test-Path -LiteralPath $explorerPath -PathType Container)) {
            Write-Warning ("Comparison directory cannot be opened because it does not exist: {0}" -f $explorerPath)
            return
        }

        Write-Info ("Opening comparison directory: {0}" -f $explorerPath) Cyan

        try {
            Invoke-Item -LiteralPath $explorerPath -ErrorAction Stop
        }
        catch {
            Start-Process -FilePath explorer.exe -ArgumentList ('"{0}"' -f $explorerPath) | Out-Null
        }
    }
    catch {
        Write-Warning ("Could not open comparison directory '{0}': {1}" -f $Path, $_.Exception.Message)
    }
}

function Get-SPOAuthConfig {
    $configPath = Join-Path -Path $ProjectRoot -ChildPath 'Config\SPOAuth.local.psd1'
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        return @{}
    }

    Import-PowerShellDataFile -LiteralPath $configPath
}

function Get-LatestCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$Filter,

        [switch]$Recurse
    )

    $parameters = @{
        Path = $Directory
        File = $true
        Filter = $Filter
        ErrorAction = 'SilentlyContinue'
    }
    if ($Recurse) {
        $parameters.Recurse = $true
    }

    $latest = Get-ChildItem @parameters |
        Where-Object { $_.Name -notlike '*-Errors.csv' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $latest) {
        throw "No CSV found in $Directory with filter $Filter."
    }

    $latest.FullName
}

function Get-ComparisonConfigValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        $DefaultValue
    )

    if ($Config.ContainsKey('Comparison') -and $Config.Comparison.ContainsKey($Name)) {
        $value = $Config.Comparison[$Name]
        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
            return $value
        }
    }

    return $DefaultValue
}

function Get-ComparisonConfigBool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [bool]$DefaultValue = $false
    )

    $value = Get-ComparisonConfigValue -Name $Name -DefaultValue $DefaultValue
    if ($value -is [bool]) {
        return $value
    }

    try {
        return [System.Convert]::ToBoolean([string]$value)
    }
    catch {
        throw "Comparison.$Name must be a boolean value. Current value: $value"
    }
}

function Assert-CsvScanAgeDifference {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceCsvPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetCsvPath,

        [Parameter(Mandatory = $true)]
        [double]$MaxAgeDifferenceHours,

        [double]$MaxAgeHours = 24,

        [string]$Label = 'Inventory'
    )

    $sourceEvidence = Get-CsvScanEvidence -CsvPath $SourceCsvPath
    $targetEvidence = Get-CsvScanEvidence -CsvPath $TargetCsvPath
    if (($null -ne $sourceEvidence.Rows -and $sourceEvidence.Rows -eq 0) -or
        ($null -ne $targetEvidence.Rows -and $targetEvidence.Rows -eq 0)) {
        throw "$Label cannot be compared: a completed inventory contains zero data rows. Source: $SourceCsvPath; Target: $TargetCsvPath"
    }
    $ageDifference = ($sourceEvidence.CompletedAtUtc - $targetEvidence.CompletedAtUtc).Duration()
    $sourceAge = ([datetime]::UtcNow - $sourceEvidence.CompletedAtUtc).TotalHours
    $targetAge = ([datetime]::UtcNow - $targetEvidence.CompletedAtUtc).TotalHours
    if ($ageDifference.TotalHours -le $MaxAgeDifferenceHours -and
        $sourceAge -le $MaxAgeHours -and $targetAge -le $MaxAgeHours -and
        $sourceAge -ge -0.25 -and $targetAge -ge -0.25) { return }

    $message = ("{0} scan age difference is {1:n2}h (maximum {2:n2}h); absolute ages: source {3:n2}h, target {4:n2}h (maximum {5:n2}h). Source evidence: {6}; target evidence: {7}." -f `
        $Label, $ageDifference.TotalHours, $MaxAgeDifferenceHours,
        $sourceAge, $targetAge, $MaxAgeHours,
        $sourceEvidence.Provenance, $targetEvidence.Provenance)

    if ($Force) {
        Write-Warning ("{0} Continuing because -Force was specified." -f $message)
        return
    }

    if ($script:LauncherNonInteractive) {
        throw ("{0} Rerun source/target scans closer together, or use -Force only after reviewing the risk." -f $message)
    }

    Write-Warning $message
    $answer = Read-Host 'Continue anyway? Type YES to continue'
    if ($answer -ne 'YES') {
        throw 'Comparison cancelled because scan age difference was not accepted.'
    }
}
function Get-CsvScanEvidence {
    param([Parameter(Mandatory = $true)][string]$CsvPath)
    $item = Get-Item -LiteralPath $CsvPath -ErrorAction Stop
    $manifestPath = "$CsvPath.manifest.json.txt"
    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($manifest.SchemaVersion -ne 1 -or $manifest.InventoryFile -ne $item.Name -or
            [string]::IsNullOrWhiteSpace([string]$manifest.Sha256) -or
            [string]::IsNullOrWhiteSpace([string]$manifest.CompletedAtUtc)) {
            throw "Invalid scan manifest: $manifestPath"
        }
        $actualHash = (Get-FileHash -LiteralPath $CsvPath -Algorithm SHA256).Hash
        if ($actualHash -ne [string]$manifest.Sha256) {
            throw "Scan CSV hash differs from its manifest: $CsvPath"
        }
        $completedAt = $manifest.CompletedAtUtc
        $completedAtUtc = if ($completedAt -is [datetime]) {
            $completedAt.ToUniversalTime()
        } elseif ($completedAt -is [datetimeoffset]) {
            $completedAt.UtcDateTime
        } else {
            ([datetimeoffset]::Parse([string]$completedAt, [Globalization.CultureInfo]::InvariantCulture)).UtcDateTime
        }
        return [pscustomobject]@{
            CompletedAtUtc = $completedAtUtc
            Rows = [int]$manifest.Rows
            Provenance = $manifestPath
        }
    }
    # Legacy scans have no receipt. Their timestamp in the name survives file copies.
    if ($item.Name -match '-(?<date>\d{8})-(?<time>\d{6})(?:-[^.]+)?\.csv$') {
        $localTime = [datetime]::ParseExact(
            ($Matches.date + $Matches.time), 'yyyyMMddHHmmss',
            [Globalization.CultureInfo]::InvariantCulture)
        Write-Warning "Legacy scan has no manifest; using filename timestamp: $CsvPath"
        return [pscustomobject]@{ CompletedAtUtc = $localTime.ToUniversalTime(); Rows = $null; Provenance = 'filename (legacy)' }
    }
    Write-Warning "Scan has no manifest or standard filename; using mutable file timestamp: $CsvPath"
    return [pscustomobject]@{ CompletedAtUtc = $item.LastWriteTimeUtc; Rows = $null; Provenance = 'LastWriteTimeUtc (unverified)' }
}

function Write-CsvScanManifest {
    param([string]$CsvPath, [string]$Side, [string]$Kind, [string]$Scope)
    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        throw "Inventory did not publish its final CSV: $CsvPath"
    }
    $python = Get-PythonCommand
    $helper = Join-Path $ProjectRoot 'Scripts\Compare\scan_evidence.py'
    $pythonArguments = @($python.Arguments) + @($helper, '--csv', $CsvPath, '--side', $Side, '--kind', $Kind, '--scope', $Scope)
    & $python.Executable @pythonArguments | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Could not create scan manifest for $CsvPath" }
    Write-Info "Scan manifest: $CsvPath.manifest.json.txt" DarkCyan
}
function Resolve-ComparisonPathMappingsFile {
    $mappingPath = [string](Get-ComparisonConfigValue -Name 'PathMappingsFile' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($mappingPath)) {
        return $null
    }

    $resolvedPath = Resolve-MigrationPath $mappingPath
    if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        throw "Path mapping file not found: $resolvedPath"
    }

    return $resolvedPath
}

function Get-PathMappingRows {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $rows = @()
    $lineNumber = 0
    foreach ($rawLine in [System.IO.File]::ReadAllLines($Path)) {
        $lineNumber++
        $line = $rawLine.Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) {
            continue
        }

        $parts = @($line -split '[\t;, ]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($parts.Count -ne 2) {
            throw "Invalid path mapping at $Path`:$lineNumber. Expected exactly two fields: <source-url-or-path> <target-url-or-path>. Encode spaces as %20; put comments on separate lines."
        }

        $rows += [pscustomobject]@{
            Source = $parts[0]
            Target = $parts[1]
        }
    }

    if ($rows.Count -eq 0) {
        throw "Path mapping file does not contain any active mapping rows: $Path"
    }

    $rows
}

function New-UrlsFileFromPathMappings {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    $mappingPath = Resolve-ComparisonPathMappingsFile
    if (-not $mappingPath) {
        throw "$Side URLs file is not configured. Set $Side.UrlsFile or Comparison.PathMappingsFile."
    }

    $mappingRows = @(Get-PathMappingRows -Path $mappingPath)
    $propertyName = if ($Side -eq 'Source') { 'Source' } else { 'Target' }
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $urls = [System.Collections.Generic.List[string]]::new()

    foreach ($row in $mappingRows) {
        $url = [string]$row.$propertyName
        if (-not [string]::IsNullOrWhiteSpace($url) -and $seen.Add($url)) {
            $urls.Add($url)
        }
    }

    if ($urls.Count -eq 0) {
        throw "No $Side URLs found in path mapping file: $mappingPath"
    }

    $generatedRoot = 'operations\generated'
    if ($Config.ContainsKey('Output') -and $Config.Output.ContainsKey('GeneratedOperations') -and -not [string]::IsNullOrWhiteSpace([string]$Config.Output.GeneratedOperations)) {
        $generatedRoot = [string]$Config.Output.GeneratedOperations
    }

    $generatedDirectory = Resolve-MigrationPath $generatedRoot
    New-Item -ItemType Directory -Path $generatedDirectory -Force | Out-Null
    $sideName = $Side.ToLowerInvariant()
    $outputPath = Join-Path -Path $generatedDirectory -ChildPath ("{0}-{1}-urls-from-mapping-{2}.txt" -f $Config.Name, $sideName, [guid]::NewGuid().ToString('N'))
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllLines($outputPath, $urls.ToArray(), $utf8NoBom)

    Write-Info ("Derived {0} URLs file from mapping: {1} ({2} URLs)" -f $Side.ToLowerInvariant(), $outputPath, $urls.Count) DarkCyan
    return $outputPath
}

function Resolve-MigrationUrlsFile {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    $configSection = if ($Side -eq 'Source') { $Config.Source } else { $Config.Target }
    $configuredPath = ''
    if ($configSection.ContainsKey('UrlsFile')) {
        $configuredPath = [string]$configSection.UrlsFile
    }

    if (-not [string]::IsNullOrWhiteSpace($configuredPath)) {
        $resolvedPath = Resolve-MigrationPath $configuredPath
        if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
            throw "$Side URLs file not found: $resolvedPath"
        }

        return $resolvedPath
    }

    New-UrlsFileFromPathMappings -Side $Side
}
function Get-MigrationEndpointConfig {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    if ($Side -eq 'Source') {
        return $Config.Source
    }

    return $Config.Target
}

function Get-MigrationEndpointType {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    $endpointConfig = Get-MigrationEndpointConfig -Side $Side
    $defaultType = if ($Side -eq 'Source') { 'SP2019' } else { 'SPO' }
    $rawType = $defaultType
    if ($endpointConfig.ContainsKey('Type') -and -not [string]::IsNullOrWhiteSpace([string]$endpointConfig.Type)) {
        $rawType = [string]$endpointConfig.Type
    }

    $normalizedType = $rawType.Trim().ToUpperInvariant()
    switch -Regex ($normalizedType) {
        '^(SP2016|SHAREPOINT2016|2016)$' { return 'SP2016' }
        '^(SP2019|SHAREPOINT2019|2019)$' { return 'SP2019' }
        '^(SPO|SHAREPOINTONLINE|ONLINE)$' { return 'SPO' }
    }

    throw "$Side.Type '$rawType' is not supported. Allowed values: SP2016, SP2019, SPO."
}

function Test-MigrationEndpointIsSPO {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    (Get-MigrationEndpointType -Side $Side) -eq 'SPO'
}

function Get-MigrationEndpointModifiedTimeZone {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    $configKey = if ($Side -eq 'Source') { 'SourceModifiedTimeZone' } else { 'TargetModifiedTimeZone' }
    $rawMode = [string](Get-ComparisonConfigValue -Name $configKey -DefaultValue 'Auto')
    if ([string]::IsNullOrWhiteSpace($rawMode)) {
        $rawMode = 'Auto'
    }

    switch ($rawMode.Trim().ToUpperInvariant()) {
        'AUTO' {
            if (Test-MigrationEndpointIsSPO -Side $Side) { return 'UTC' }
            return 'Local'
        }
        'NONE' { return 'Raw' }
        'RAW' { return 'Raw' }
        'LOCAL' { return 'Local' }
        'UTC' { return 'UTC' }
    }

    throw "$configKey '$rawMode' is not supported. Allowed values: Auto, Raw, Local, UTC."
}

function Get-MigrationEndpointPermissionLibraryOnly {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    if ($Side -eq 'Source') {
        return [bool]$Config.Permissions.SourceDocumentLibrariesOnly
    }

    return [bool]$Config.Permissions.TargetDocumentLibrariesOnly
}

function Add-SPOInventoryAuthenticationParameters {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters
    )

    $authConfig = Get-SPOAuthConfig
    foreach ($key in @('ClientId', 'Tenant', 'TenantId')) {
        if ($authConfig.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$authConfig[$key])) {
            $Parameters[$key] = $authConfig[$key]
        }
    }

    if ($UseCertificate) {
        if ($authConfig.ContainsKey('Thumbprint') -and -not [string]::IsNullOrWhiteSpace([string]$authConfig.Thumbprint)) {
            $Parameters.Thumbprint = $authConfig.Thumbprint
        }

        if (-not $Parameters.ContainsKey('ClientId') -or -not $Parameters.ContainsKey('Thumbprint')) {
            throw "Certificate authentication requires ClientId and Thumbprint in Config\SPOAuth.local.psd1."
        }
    }
    elseif ($DeviceLogin) {
        $Parameters.DeviceLogin = $true
    }
    else {
        $Parameters.Interactive = $true
    }

    if ($ForceAuthentication) {
        $Parameters.ForceAuthentication = $true
    }
}

function Add-OnPremInventoryScopeParameters {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Parameters,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    $endpointConfig = Get-MigrationEndpointConfig -Side $Side
    if ($endpointConfig.ContainsKey('WebApplicationUrl') -and -not [string]::IsNullOrWhiteSpace([string]$endpointConfig.WebApplicationUrl)) {
        $Parameters.WebApplicationUrl = [string]$endpointConfig.WebApplicationUrl
        $Parameters.UseSiteUrlFilter = $true
        $Parameters.SiteUrlsFile = Resolve-MigrationUrlsFile -Side $Side
        return
    }

    if ($endpointConfig.ContainsKey('SiteUrl') -and -not [string]::IsNullOrWhiteSpace([string]$endpointConfig.SiteUrl)) {
        $Parameters.SiteUrl = [string]$endpointConfig.SiteUrl
        return
    }

    throw "$Side.Type $(Get-MigrationEndpointType -Side $Side) requires $Side.WebApplicationUrl for mapped multi-site scans, or $Side.SiteUrl for a single site scan."
}

function New-MigrationInventoryOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side,

        [Parameter(Mandatory = $true)]
        [ValidateSet('File', 'Permission')]
        [string]$InventoryKind,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string]$Timestamp
    )

    $endpointType = Get-MigrationEndpointType -Side $Side
    Join-Path -Path $OutputDirectory -ChildPath ("{0}-{1}Inventory-{2}-{3}.csv" -f $endpointType, $InventoryKind, $Config.Name, $Timestamp)
}

function Invoke-FileInventoryScan {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    $timestamp = New-MigrationRunTimestamp
    $endpointType = Get-MigrationEndpointType -Side $Side
    $isSPO = Test-MigrationEndpointIsSPO -Side $Side
    $scriptName = if ($isSPO) { 'SmartM365-SharePointTarget-FileInventory.ps1' } else { 'SmartM365-SharePointSource-FileInventory.ps1' }
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath ("Scripts\Inventory\{0}" -f $scriptName)
    $outputDirectory = if ($Side -eq 'Source') { Resolve-MigrationPath $Config.Output.SourceFileScans } else { Resolve-MigrationPath $Config.Output.TargetFileScans }
    $actionName = if ($Side -eq 'Source') { 'ScanSourceFiles' } else { 'ScanTargetFiles' }
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

    $parameters = @{
        OutputPath = (New-MigrationInventoryOutputPath -Side $Side -InventoryKind 'File' -OutputDirectory $outputDirectory -Timestamp $timestamp)
        LogPath = (New-MigrationLogPath -Action $actionName -Timestamp $timestamp)
    }

    if ($isSPO) {
        $parameters.WebUrlsFile = Resolve-MigrationUrlsFile -Side $Side
        Add-SPOInventoryAuthenticationParameters -Parameters $parameters
    }
    else {
        Add-OnPremInventoryScopeParameters -Parameters $parameters -Side $Side
    }

    Write-Info ("Starting {0} file scan for migration '{1}' ({2})" -f $Side.ToLowerInvariant(), $Config.Name, $endpointType) Cyan
    $authMode = if ($parameters.ContainsKey('Thumbprint')) { 'Certificate' } elseif ($parameters.ContainsKey('DeviceLogin') -and $parameters.DeviceLogin) { 'DeviceLogin' } elseif ($parameters.ContainsKey('Interactive') -and $parameters.Interactive) { 'Interactive' } else { 'Default' }
    $inputKey = @('WebUrlsFile', 'WebApplicationUrl', 'SiteUrl', 'WebUrl', 'TenantAdminUrl') | Where-Object { $parameters.ContainsKey($_) } | Select-Object -First 1
    $inputValue = if ($inputKey) { $parameters[$inputKey] } else { '<none>' }
    $forceAuthenticationEnabled = $parameters.ContainsKey('ForceAuthentication') -and [bool]$parameters.ForceAuthentication
    $useSiteUrlFilterEnabled = $parameters.ContainsKey('UseSiteUrlFilter') -and [bool]$parameters.UseSiteUrlFilter
    $siteUrlsFileLabel = if ($parameters.ContainsKey('SiteUrlsFile')) { $parameters.SiteUrlsFile } else { '<none>' }
    Write-Info ("File scan options: AuthMode={0}; ForceAuthentication={1}; ParameterSet={2}; Input={3}; UseSiteUrlFilter={4}; SiteUrlsFile={5}" -f $authMode, $forceAuthenticationEnabled, $inputKey, $inputValue, $useSiteUrlFilterEnabled, $siteUrlsFileLabel) DarkCyan
    & $scriptPath @parameters
    Write-CsvScanManifest -CsvPath $parameters.OutputPath -Side $Side -Kind 'File' -Scope ([string]$inputValue)
}

function Invoke-PermissionInventoryScan {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    $timestamp = New-MigrationRunTimestamp
    $endpointType = Get-MigrationEndpointType -Side $Side
    $isSPO = Test-MigrationEndpointIsSPO -Side $Side
    $scriptName = if ($isSPO) { 'SmartM365-SharePointTarget-PermissionInventory.ps1' } else { 'SmartM365-SharePointSource-PermissionInventory.ps1' }
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath ("Scripts\Inventory\{0}" -f $scriptName)
    $outputDirectory = if ($Side -eq 'Source') { Resolve-MigrationPath $Config.Output.SourcePermissionScans } else { Resolve-MigrationPath $Config.Output.TargetPermissionScans }
    $actionName = if ($Side -eq 'Source') { 'ScanSourcePermissions' } else { 'ScanTargetPermissions' }
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

    $parameters = @{
        OutputPath = (New-MigrationInventoryOutputPath -Side $Side -InventoryKind 'Permission' -OutputDirectory $outputDirectory -Timestamp $timestamp)
        LogPath = (New-MigrationLogPath -Action $actionName -Timestamp $timestamp)
        DocumentLibrariesOnly = (Get-MigrationEndpointPermissionLibraryOnly -Side $Side)
        IncludeItemPermissions = [bool]$Config.Permissions.IncludeItemPermissions
        ItemProgressInterval = [int]$Config.Permissions.ItemProgressInterval
    }

    if ($isSPO) {
        $parameters.WebUrlsFile = Resolve-MigrationUrlsFile -Side $Side
        Add-SPOInventoryAuthenticationParameters -Parameters $parameters
    }
    else {
        Add-OnPremInventoryScopeParameters -Parameters $parameters -Side $Side
    }

    Write-Info ("Starting {0} permission scan for migration '{1}' ({2})" -f $Side.ToLowerInvariant(), $Config.Name, $endpointType) Cyan
    $permissionScope = if ($parameters.DocumentLibrariesOnly) { 'DocumentLibrariesOnly' } else { 'AllListsAndLibraries' }
    $authMode = if ($parameters.ContainsKey('Thumbprint')) { 'Certificate' } elseif ($parameters.ContainsKey('DeviceLogin') -and $parameters.DeviceLogin) { 'DeviceLogin' } elseif ($parameters.ContainsKey('Interactive') -and $parameters.Interactive) { 'Interactive' } else { 'Default' }
    Write-Info ("Permission scan options: Scope={0}; IncludeItemPermissions={1}; ItemProgressInterval={2}; AuthMode={3}; ForceAuthentication={4}" -f $permissionScope, [bool]$parameters.IncludeItemPermissions, [int]$parameters.ItemProgressInterval, $authMode, [bool]$parameters.ForceAuthentication) DarkCyan
    & $scriptPath @parameters
    $scopeValue = if ($parameters.ContainsKey('WebUrlsFile')) { $parameters.WebUrlsFile } elseif ($parameters.ContainsKey('SiteUrl')) { $parameters.SiteUrl } else { $parameters.WebApplicationUrl }
    Write-CsvScanManifest -CsvPath $parameters.OutputPath -Side $Side -Kind 'Permission' -Scope ([string]$scopeValue)
}

function Invoke-SourceFileScan {
    Invoke-FileInventoryScan -Side 'Source'
}

function Invoke-TargetFileScan {
    Invoke-FileInventoryScan -Side 'Target'
}

function Invoke-SourcePermissionScan {
    Invoke-PermissionInventoryScan -Side 'Source'
}

function Invoke-TargetPermissionScan {
    Invoke-PermissionInventoryScan -Side 'Target'
}
function Invoke-FileComparison {
    $timestamp = New-MigrationRunTimestamp
    $logPath = New-MigrationLogPath -Action 'CompareFiles' -Timestamp $timestamp
    $sourceDirectory = Resolve-MigrationPath $Config.Output.SourceFileScans
    $targetDirectory = Resolve-MigrationPath $Config.Output.TargetFileScans
    $outputDirectory = Join-Path -Path (Resolve-MigrationPath $Config.Output.FileComparisons) -ChildPath ("{0}-files-{1}" -f $Config.Name, $timestamp)
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    Start-Transcript -Path $logPath -Force -WhatIf:$false | Out-Null

    try {
        Write-Info ("Run log: {0}" -f $logPath) Cyan
        if ([string]::IsNullOrWhiteSpace($SourceCsv)) {
            $SourceCsv = Get-LatestCsv -Directory $sourceDirectory -Filter ("{0}-FileInventory-{1}-*.csv" -f (Get-MigrationEndpointType -Side 'Source'), $Config.Name) -Recurse
        }
        if ([string]::IsNullOrWhiteSpace($TargetCsv)) {
            $TargetCsv = Get-LatestCsv -Directory $targetDirectory -Filter ("{0}-FileInventory-{1}-*.csv" -f (Get-MigrationEndpointType -Side 'Target'), $Config.Name)
        }

        $maxScanAgeDifferenceHours = [double](Get-ComparisonConfigValue -Name 'MaxScanAgeDifferenceHours' -DefaultValue 12)
        $maxScanAgeHours = [double](Get-ComparisonConfigValue -Name 'MaxScanAgeHours' -DefaultValue 24)
        $modifiedDateToleranceMinutes = [double](Get-ComparisonConfigValue -Name 'ModifiedDateToleranceMinutes' -DefaultValue 0)
        $sourceModifiedTimeZone = Get-MigrationEndpointModifiedTimeZone -Side 'Source'
        $targetModifiedTimeZone = Get-MigrationEndpointModifiedTimeZone -Side 'Target'
        Assert-CsvScanAgeDifference `
            -SourceCsvPath $SourceCsv `
            -TargetCsvPath $TargetCsv `
            -MaxAgeDifferenceHours $maxScanAgeDifferenceHours `
            -MaxAgeHours $maxScanAgeHours `
            -Label 'File inventory'

        $compareScript = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Compare\compare_sp_source_target_file_inventories.py'
        $pathMappingsFile = Resolve-ComparisonPathMappingsFile
        $compareArguments = @(
            '--source-csv', $SourceCsv,
            '--target-csv', $TargetCsv,
            '--output-directory', $outputDirectory,
            '--source-web-urls-file', (Resolve-MigrationUrlsFile -Side 'Source'),
            '--target-web-urls-file', (Resolve-MigrationUrlsFile -Side 'Target'),
            '--comparison-name', ("{0}-{1}-vs-{2}" -f $Config.Name, (Get-MigrationEndpointType -Side 'Source'), (Get-MigrationEndpointType -Side 'Target')),
            '--size-tolerance-bytes', ([string]$Config.Comparison.SizeToleranceBytes),
            '--max-scan-age-difference-hours', ([string]$maxScanAgeDifferenceHours),
            '--max-scan-age-hours', ([string]$maxScanAgeHours),
            '--modified-date-tolerance-minutes', ([string]$modifiedDateToleranceMinutes),
            '--source-modified-time-zone', $sourceModifiedTimeZone,
            '--target-modified-time-zone', $targetModifiedTimeZone,
            '--sharegate-replacement-character', ([string]$Config.Comparison.ShareGateReplacementCharacter)
        )
        foreach ($targetPrefix in @($Config.Target.PrefixToRemove)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$targetPrefix)) {
                $compareArguments += @('--target-prefix', [string]$targetPrefix)
            }
        }
        if ($pathMappingsFile) {
            $compareArguments += @('--path-mapping-file', $pathMappingsFile)
        }
        Invoke-PythonScript -Script $compareScript -Arguments $compareArguments

        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Export\export_comparison_to_excel.py') -Arguments @(
            '--comparison-directory', $outputDirectory,
            '--output-xlsx', (Join-Path $outputDirectory ("{0}-{1}-vs-{2}-Comparison-{3}.xlsx" -f $Config.Name, (Get-MigrationEndpointType -Side 'Source'), (Get-MigrationEndpointType -Side 'Target'), $timestamp))
        )
        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Export\export_duplicate_keys_to_excel.py') -Arguments @(
            '--comparison-directory', $outputDirectory,
            '--output-xlsx', (Join-Path $outputDirectory ("DuplicateKeys-{0}.xlsx" -f $timestamp))
        )
        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Export\export_libraries_to_delete_excel.py') -Arguments @(
            '--comparison-directory', $outputDirectory,
            '--output-xlsx', (Join-Path $outputDirectory ("Libraries-To-Delete-{0}.xlsx" -f $timestamp))
        )
        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Export\export_files_to_delete_excel.py') -Arguments @(
            '--comparison-directory', $outputDirectory,
            '--output-xlsx', (Join-Path $outputDirectory ("Files-To-Delete-{0}.xlsx" -f $timestamp))
        )

        $deleteScriptArguments = @('--comparison-directory', $outputDirectory)
        if (Get-ComparisonConfigBool -Name 'AllowDuplicateKeysForDeleteScript' -DefaultValue $false) {
            $deleteScriptArguments += '--allow-duplicate-keys'
        }
        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Generate\generate_sp_target_delete_libraries_script.py') -Arguments ($deleteScriptArguments + @(
            '--output-ps1', (Join-Path $outputDirectory ("SmartM365-SharePointTarget-ExtraLibrariesRemove-{0}.ps1" -f $timestamp))
        ))
        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Generate\generate_sp_target_delete_files_script.py') -Arguments ($deleteScriptArguments + @(
            '--output-ps1', (Join-Path $outputDirectory ("SmartM365-SharePointTarget-ExtraFilesRemove-{0}.ps1" -f $timestamp))
        ))
        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Generate\generate_sp_target_delete_folders_script.py') -Arguments @(
            '--comparison-directory', $outputDirectory,
            '--output-ps1', (Join-Path $outputDirectory ("SmartM365-SharePointTarget-ExtraFoldersRemove-{0}.ps1" -f $timestamp))
        )

        Write-Info ("File comparison completed: {0}" -f $outputDirectory) Green
        Open-DirectoryInExplorer -Path $outputDirectory
    }
    finally {
        Stop-LauncherTranscript -Path $logPath
    }
}

function Get-GeneratedOperationsDirectory {
    $generatedRoot = 'operations\generated'
    if ($Config.ContainsKey('Output') -and $Config.Output.ContainsKey('GeneratedOperations') -and -not [string]::IsNullOrWhiteSpace([string]$Config.Output.GeneratedOperations)) {
        $generatedRoot = [string]$Config.Output.GeneratedOperations
    }

    $generatedDirectory = Resolve-MigrationPath $generatedRoot
    New-Item -ItemType Directory -Path $generatedDirectory -Force | Out-Null
    return $generatedDirectory
}

function Resolve-EntraUsersCachePath {
    $configuredPath = [string](Get-ComparisonConfigValue -Name 'EntraUsersCachePath' -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($configuredPath)) {
        return (Resolve-MigrationPath $configuredPath)
    }

    return (Join-Path -Path (Get-GeneratedOperationsDirectory) -ChildPath ("{0}-entra-users-cache.csv" -f $Config.Name))
}

function Get-AuthConfigValue {
    param(
        [hashtable]$AuthConfig,
        [string[]]$Names
    )

    foreach ($name in $Names) {
        if ($AuthConfig.ContainsKey($name) -and -not [string]::IsNullOrWhiteSpace([string]$AuthConfig[$name])) {
            return [string]$AuthConfig[$name]
        }
    }

    return ''
}

function Update-EntraUsersCacheForComparison {
    if (-not (Get-ComparisonConfigBool -Name 'EntraUsersCacheEnabled' -DefaultValue $true)) {
        throw "Entra users cache is mandatory for permission comparisons. Set Comparison.EntraUsersCacheEnabled = `$true."
    }

    if ((Get-MigrationEndpointType -Side 'Source') -ne 'SPO' -and (Get-MigrationEndpointType -Side 'Target') -ne 'SPO') {
        throw "Entra users cache is mandatory for permission comparisons. At least one endpoint must be SPO to resolve destination Entra users."
    }

    $cachePath = Resolve-EntraUsersCachePath
    $maxAgeHours = [double](Get-ComparisonConfigValue -Name 'EntraUsersCacheMaxAgeHours' -DefaultValue 6)
    if ($maxAgeHours -le 0) { throw 'Comparison.EntraUsersCacheMaxAgeHours must be greater than zero.' }
    $maxAgeHours = [math]::Min($maxAgeHours, 6)
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Inventory\SmartM365-EntraUsers-CacheExport.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "Entra users cache export script not found: $scriptPath"
    }

    $authConfig = Get-SPOAuthConfig
    $tenantId = Get-AuthConfigValue -AuthConfig $authConfig -Names @('TenantId', 'Tenant')
    $clientId = Get-AuthConfigValue -AuthConfig $authConfig -Names @('ClientId', 'AppId')
    $thumbprint = Get-AuthConfigValue -AuthConfig $authConfig -Names @('Thumbprint', 'Thumb')

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $scriptPath,
        '-OutputPath', $cachePath,
        '-MaxCacheAgeHours', ([string]$maxAgeHours)
    )

    if ($ForceAuthentication) {
        $arguments += '-Connect'
    }
    if ($UseCertificate) {
        if (-not ($tenantId -and $clientId -and $thumbprint)) {
            throw "Certificate authentication for Entra users cache requires TenantId, ClientId and Thumbprint in Config\SPOAuth.local.psd1."
        }

        $arguments += @('-TenantId', $tenantId, '-AppId', $clientId, '-CertificateThumbprint', $thumbprint)
    }
    else {
        $arguments += '-InteractiveAuth'
        if ($DeviceLogin) {
            $arguments += '-DeviceLogin'
        }
    }

    Write-Info ("Ensuring Entra users cache: {0}" -f $cachePath) Cyan
    & pwsh @arguments 2>&1 | ForEach-Object { Microsoft.PowerShell.Utility\Write-Host ([string]$_) }
    if ($LASTEXITCODE -ne 0) {
        throw ("Entra users cache export failed with exit code {0}: {1}" -f $LASTEXITCODE, $scriptPath)
    }

    return $cachePath
}

function Invoke-PermissionComparison {
    $timestamp = New-MigrationRunTimestamp
    $logPath = New-MigrationLogPath -Action 'ComparePermissions' -Timestamp $timestamp
    $sourceDirectory = Resolve-MigrationPath $Config.Output.SourcePermissionScans
    $targetDirectory = Resolve-MigrationPath $Config.Output.TargetPermissionScans
    $outputDirectory = Join-Path -Path (Resolve-MigrationPath $Config.Output.PermissionComparisons) -ChildPath ("{0}-permissions-{1}" -f $Config.Name, $timestamp)
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    Start-Transcript -Path $logPath -Force -WhatIf:$false | Out-Null

    try {
        Write-Info ("Run log: {0}" -f $logPath) Cyan
        if ([string]::IsNullOrWhiteSpace($SourceCsv)) {
            $SourceCsv = Get-LatestCsv -Directory $sourceDirectory -Filter ("{0}-PermissionInventory-{1}-*.csv" -f (Get-MigrationEndpointType -Side 'Source'), $Config.Name) -Recurse
        }
        if ([string]::IsNullOrWhiteSpace($TargetCsv)) {
            $TargetCsv = Get-LatestCsv -Directory $targetDirectory -Filter ("{0}-PermissionInventory-{1}-*.csv" -f (Get-MigrationEndpointType -Side 'Target'), $Config.Name)
        }

        $maxScanAgeDifferenceHours = [double](Get-ComparisonConfigValue -Name 'PermissionMaxScanAgeDifferenceHours' -DefaultValue 24)
        $maxScanAgeHours = [double](Get-ComparisonConfigValue -Name 'PermissionMaxScanAgeHours' -DefaultValue 48)
        Assert-CsvScanAgeDifference `
            -SourceCsvPath $SourceCsv `
            -TargetCsvPath $TargetCsv `
            -MaxAgeDifferenceHours $maxScanAgeDifferenceHours `
            -MaxAgeHours $maxScanAgeHours `
            -Label 'Permission inventory'

        $pathMappingsFile = Resolve-ComparisonPathMappingsFile
        $entraUsersCachePath = Update-EntraUsersCacheForComparison
        if ([string]::IsNullOrWhiteSpace($entraUsersCachePath) -or -not (Test-Path -LiteralPath $entraUsersCachePath -PathType Leaf)) {
            throw "Entra users cache is mandatory for permission comparisons, but no cache file was produced."
        }

        $permissionCompareArguments = @(
            '--source-csv', $SourceCsv,
            '--target-csv', $TargetCsv,
            '--output-directory', $outputDirectory,
            '--source-root-path', $Config.Source.PermissionRootPath,
            '--target-root-path', $Config.Target.PermissionRootPath,
            '--sharegate-replacement-character', ([string]$Config.Comparison.ShareGateReplacementCharacter),
            '--comparison-name', ("{0}-{1}-vs-{2}-Permissions" -f $Config.Name, (Get-MigrationEndpointType -Side 'Source'), (Get-MigrationEndpointType -Side 'Target')),
            '--max-scan-age-difference-hours', ([string]$maxScanAgeDifferenceHours),
            '--max-scan-age-hours', ([string]$maxScanAgeHours)
        )
        if (Get-MigrationEndpointPermissionLibraryOnly -Side 'Source') {
            $permissionCompareArguments += '--source-scan-document-libraries-only'
        }
        if (Get-MigrationEndpointPermissionLibraryOnly -Side 'Target') {
            $permissionCompareArguments += '--target-scan-document-libraries-only'
        }
        if ($pathMappingsFile) {
            $permissionCompareArguments += @('--path-mapping-file', $pathMappingsFile)
        }
        $permissionCompareArguments += @('--entra-users-csv', $entraUsersCachePath)
        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Compare\compare_sp_source_target_permissions.py') -Arguments $permissionCompareArguments

        Write-Info ("Permission comparison completed: {0}" -f $outputDirectory) Green
        Open-DirectoryInExplorer -Path $outputDirectory
    }
    finally {
        Stop-LauncherTranscript -Path $logPath
    }
}

function Get-FileHistoryComparisonRoot {
    if ($Config.ContainsKey('Output') -and $Config.Output.ContainsKey('FileHistoryComparisons') -and -not [string]::IsNullOrWhiteSpace([string]$Config.Output.FileHistoryComparisons)) {
        return (Resolve-MigrationPath $Config.Output.FileHistoryComparisons)
    }

    if ($Config.ContainsKey('Output') -and $Config.Output.ContainsKey('SourceHistoryComparisons') -and -not [string]::IsNullOrWhiteSpace([string]$Config.Output.SourceHistoryComparisons)) {
        return (Resolve-MigrationPath $Config.Output.SourceHistoryComparisons)
    }

    return (Resolve-MigrationPath 'comparisons\scan-history')
}

function Invoke-FileHistoryComparison {
    param(
        [ValidateSet('Source', 'Target')]
        [string]$Side = 'Source',

        [string]$LogAction = 'CompareScanHistory'
    )

    $timestamp = New-MigrationRunTimestamp
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Compare\SmartM365-SharePointSource-FileInventoryHistoryCompare.ps1'
    $endpointType = Get-MigrationEndpointType -Side $Side
    $scanDirectory = if ($Side -eq 'Source') { Resolve-MigrationPath $Config.Output.SourceFileScans } else { Resolve-MigrationPath $Config.Output.TargetFileScans }
    $outputDirectory = Join-Path -Path (Get-FileHistoryComparisonRoot) -ChildPath ("{0}-{1}-Changes-{2}" -f $Config.Name, $endpointType, $timestamp)

    $parameters = @{
        ScanDirectory = $scanDirectory
        InventoryNameFilter = ("*{0}*" -f $Config.Name)
        WebUrlsFile = (Resolve-MigrationUrlsFile -Side $Side)
        OutputDirectory = $outputDirectory
        LogPath = (New-MigrationLogPath -Action $LogAction -Timestamp $timestamp)
        ComparisonName = ("{0}-{1}-history" -f $Config.Name, $endpointType)
        EndpointType = $endpointType
        Side = $Side
        SizeToleranceBytes = [long]$Config.Comparison.SizeToleranceBytes
        ModifiedDateToleranceMinutes = [double](Get-ComparisonConfigValue -Name 'ModifiedDateToleranceMinutes' -DefaultValue 0)
    }

    if (-not [string]::IsNullOrWhiteSpace($OldCsv)) {
        $parameters.OldCsv = $OldCsv
    }
    if (-not [string]::IsNullOrWhiteSpace($NewCsv)) {
        $parameters.NewCsv = $NewCsv
    }
    if ($Force) {
        $parameters.Force = $true
    }

    Write-Info ("Starting {0} scan history comparison for migration '{1}'" -f $Side.ToLowerInvariant(), $Config.Name) Cyan
    & $scriptPath @parameters
    Open-DirectoryInExplorer -Path $outputDirectory
}

function Invoke-PermissionHistoryComparison {
    param([ValidateSet('Source', 'Target')][string]$Side = 'Source')

    $scanDirectory = if ($Side -eq 'Source') {
        Resolve-MigrationPath $Config.Output.SourcePermissionScans
    } else { Resolve-MigrationPath $Config.Output.TargetPermissionScans }
    $endpointType = Get-MigrationEndpointType -Side $Side
    $filter = '{0}-PermissionInventory-{1}-*.csv' -f $endpointType, $Config.Name
    $scans = @(Get-ChildItem -LiteralPath $scanDirectory -Filter $filter -File -Recurse -ErrorAction Stop |
        Where-Object { $_.Name -notlike '*-Errors.csv' } | Sort-Object {
            if ($_.Name -match '-(?<date>\d{8})-(?<time>\d{6})(?:-[^.]+)?\.csv$') {
                [datetime]::ParseExact(($Matches.date + $Matches.time), 'yyyyMMddHHmmss',
                    [Globalization.CultureInfo]::InvariantCulture)
            } else { $_.LastWriteTime }
        } -Descending)
    if ($scans.Count -lt 2 -and (-not $OldCsv -or -not $NewCsv)) {
        throw "At least two completed $Side permission scans are required."
    }
    $older = if ($OldCsv) { $OldCsv } else { $scans[1].FullName }
    $newer = if ($NewCsv) { $NewCsv } else { $scans[0].FullName }
    $scanRoot = [IO.Path]::GetFullPath($scanDirectory).TrimEnd('\') + '\'
    foreach ($path in @($older, $newer)) {
        $full = [IO.Path]::GetFullPath($path)
        if (-not $full.StartsWith($scanRoot, [StringComparison]::OrdinalIgnoreCase) -or
            -not (Test-Path -LiteralPath $full -PathType Leaf) -or
            [IO.Path]::GetFileName($full) -notlike $filter -or
            [IO.Path]::GetFileName($full) -like '*-Errors.csv') {
            throw "Selected permission scan is outside the $Side inventory or has an unexpected name: $path"
        }
    }
    if ([string]::Equals([IO.Path]::GetFullPath($older), [IO.Path]::GetFullPath($newer),
            [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Previous and current permission scans must be different files.'
    }

    $historyRoot = if ($Config.Output.ContainsKey('PermissionHistoryComparisons') -and
        -not [string]::IsNullOrWhiteSpace([string]$Config.Output.PermissionHistoryComparisons)) {
        Resolve-MigrationPath $Config.Output.PermissionHistoryComparisons
    } else { Resolve-MigrationPath 'comparisons\permission-scan-history' }
    $timestamp = New-MigrationRunTimestamp
    $outputDirectory = Join-Path $historyRoot ('{0}-{1}-PermissionChanges-{2}' -f $Config.Name, $endpointType, $timestamp)
    $logPath = New-MigrationLogPath -Action 'ComparePermissionScanHistory' -Timestamp $timestamp
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    Start-Transcript -Path $logPath -Force -WhatIf:$false | Out-Null
    try {
        Write-Info ("Run log: {0}" -f $logPath) Cyan
        Write-Info ("Comparing $Side permission scans: $older -> $newer") Cyan
        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Compare\compare_sp_permission_inventory_history.py') `
            -Arguments @('--old-csv', $older, '--new-csv', $newer,
                '--output-directory', $outputDirectory,
                '--comparison-name', ('{0}-{1}-permission-history' -f $Config.Name, $endpointType))
        Write-Info ("Permission scan history comparison completed: {0}" -f $outputDirectory) Green
        if (-not $script:LauncherNonInteractive) { Open-DirectoryInExplorer -Path $outputDirectory }
    }
    finally { Stop-LauncherTranscript -Path $logPath }
}
# SharePoint Server snap-ins are unavailable in PowerShell 7. Re-enter the same
# launcher under Windows PowerShell, retaining explicit switches and literal paths.
if ($PSVersionTable.PSVersion.Major -gt 5 -and $Action -match '^Scan(Source|Target)(Files|Permissions)$') {
    $scanSide = $Matches[1]
    if (-not (Test-MigrationEndpointIsSPO -Side $scanSide)) {
        $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        if (-not (Test-Path -LiteralPath $windowsPowerShell -PathType Leaf)) {
            throw 'SharePoint Server scans require Windows PowerShell 5.1 on a SharePoint server.'
        }
        $argumentExpressions = foreach ($entry in $PSBoundParameters.GetEnumerator()) {
            if ($entry.Value -is [System.Management.Automation.SwitchParameter] -or $entry.Value -is [bool]) {
                '-{0}:${1}' -f $entry.Key, ([bool]$entry.Value).ToString().ToLowerInvariant()
            }
            else {
                "-{0} '{1}'" -f $entry.Key, ([string]$entry.Value).Replace("'", "''")
            }
        }
        $commandText = "& '{0}' {1}" -f $PSCommandPath.Replace("'", "''"), ($argumentExpressions -join ' ')
        $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($commandText))
        Write-Info 'Starting SharePoint Server inventory in Windows PowerShell 5.1.' Cyan
        & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand
        exit $LASTEXITCODE
    }
}

$script:LauncherRunStartedAt = Get-Date
$script:LauncherRunLogPath = ''
$script:LauncherFailure = $null
$env:SPMIG_CONSOLE_LIFECYCLE_ACTIVE = '1'
Write-LauncherIntro
try {
    switch ($Action) {
        'ScanSourceFiles' { Invoke-SourceFileScan }
        'ScanTargetFiles' { Invoke-TargetFileScan }
        'CompareFiles' { Invoke-FileComparison }
        'ScanSourcePermissions' { Invoke-SourcePermissionScan }
        'ScanTargetPermissions' { Invoke-TargetPermissionScan }
        'ComparePermissions' { Invoke-PermissionComparison }
        'CompareSourceHistory' { Invoke-FileHistoryComparison -Side 'Source' -LogAction 'CompareSourceHistory' }
        'CompareScanHistory' { Invoke-FileHistoryComparison -Side $HistorySide -LogAction 'CompareScanHistory' }
        'ComparePermissionScanHistory' { Invoke-PermissionHistoryComparison -Side $HistorySide }
    }
}
catch {
    $script:LauncherFailure = $_
    throw
}
finally {
    Write-LauncherSummary -Failure $script:LauncherFailure
    if ($script:LauncherPreviousConsoleMarker) { $env:SPMIG_CONSOLE_LIFECYCLE_ACTIVE = $script:LauncherPreviousConsoleMarker }
    else { Remove-Item Env:SPMIG_CONSOLE_LIFECYCLE_ACTIVE -ErrorAction SilentlyContinue }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDyXZgL3wOJ9Ugm
# repSw7k6nGpTG0u8IbOEzsb2Jqdt+qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAA6LZFZLMMQg7HVelP5D+u
# FIRbU+gkvDNI77caqFEpyDANBgkqhkiG9w0BAQEFAASCAYAxbNeQa+d3Y/JqkktQ
# y9z2J6H89iQhmLWlcEZv2oW4F3Hsral1ghF3YUBh2v/47Po02OHWMkukCemu3e46
# IYFuSyvd9fXFjm7vHDwu3IvZ9+mJc6OBPoJhuiQpS8Hr48kBdVeSQSY0D7tw02uN
# PKy3S/Y29Dgwrdz5fHwzKiYGGJPAiuAQ/Ba7CHH6cSwq0WSLCYFasedl+FmgMwMS
# KZdCkJE33jxI+MOz059Y6FrZKI2cYK483DmZqCaaWFlECmzwxZWk1DPM6NDwjLlM
# XXHtqTRTuf6NiOYd9meh7i6FPo7tnivS/6FhveqQ+nOgJ1i6viFNlMqFwTT+C4bP
# soFy6GtojUfmn7L/2Uzmh9Mth8GczoXWNpgn55GX3mmoBGQFRbeltkm8JhhkgIv0
# n0yA88nROc+/3rK0w4NR/8EQGXcJ9xJLATEl63ZMr0oHNWdNckoqtryGobfu+zgb
# PLnEjJgegJMYnCtES//I8nOmgrA+qSD/z7Di2caDlBAgABI=
# SIG # End signature block
