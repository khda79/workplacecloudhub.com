<#
.SYNOPSIS
    Run source SharePoint Server inventories for configured migrations.

.DESCRIPTION
    Run on a SharePoint farm server under an account with SharePoint Shell
    access. Copies the signed inventory scripts and their lifecycle helper to
    a unique local folder, then runs one Windows PowerShell 5.1 scan at a time.
    Source URLs come from Source.UrlsFile or migration.mapping.txt; an absent
    filter is an error so a batch cannot accidentally scan an entire farm.

.VERSION
    1.0.5

.EXAMPLE
    powershell.exe -NoProfile -File .\SmartM365-SharePointMigration-SourceScanBatch.ps1 -PlanOnly

.EXAMPLE
    powershell.exe -NoProfile -File .\SmartM365-SharePointMigration-SourceScanBatch.ps1 -InventoryMode Both
#>

[CmdletBinding()]
param(
    [string[]]$MigrationNames = @(),
    [ValidateSet('FilesOnly', 'PermissionsOnly', 'Both')]
    [string]$InventoryMode = 'Both',
    [switch]$PlanOnly,
    [ValidatePattern('^\d{8}-\d{6}-[a-f0-9]{8}$')]
    [string]$BatchId = '',
    [string]$ProjectRoot = '',
    [string]$LocalRunRoot = (Join-Path $env:TEMP 'SmartM365-SharePointMigration-SourceBatch')
)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -eq 5) {
    $windowsModules = Join-Path $PSHOME 'Modules'
    $env:PSModulePath = $windowsModules + [System.IO.Path]::PathSeparator + $env:PSModulePath
}
$script:BatchLogPath = $null
$script:Results = New-Object 'System.Collections.Generic.List[object]'

function Write-Host {
    param([Parameter(Position = 0, ValueFromRemainingArguments = $true)][object[]]$Object)
    foreach ($part in $Object) {
        foreach ($line in ([string]$part -split "`r?`n")) {
            Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line)
        }
    }
}

function Write-BatchLine {
    param([string]$Message)
    foreach ($line in ($Message -split "`r?`n")) {
        $entry = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line
        Microsoft.PowerShell.Utility\Write-Host $entry
        if ($script:BatchLogPath) {
            Add-Content -LiteralPath $script:BatchLogPath -Value $entry -Encoding UTF8
        }
    }
}

function Resolve-MigrationFile {
    param([string]$MigrationRoot, [string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    if ([System.IO.Path]::IsPathRooted($Value)) { return $Value }
    return (Join-Path $MigrationRoot $Value)
}

function Quote-PowerShellLiteral {
    param([string]$Value)
    return ("'{0}'" -f $Value.Replace("'", "''"))
}

function Get-SourceUrls {
    param([string]$MigrationRoot, [hashtable]$Config)

    $urlsPath = Resolve-MigrationFile -MigrationRoot $MigrationRoot -Value ([string]$Config.Source.UrlsFile)
    if ($urlsPath) {
        if (-not (Test-Path -LiteralPath $urlsPath -PathType Leaf)) {
            throw "Configured source URLs file not found: $urlsPath"
        }
        $urls = @(Get-Content -LiteralPath $urlsPath | ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') })
    }
    else {
        $mappingValue = if ($Config.Comparison) { [string]$Config.Comparison.PathMappingsFile } else { '' }
        $mappingPath = Resolve-MigrationFile -MigrationRoot $MigrationRoot -Value $mappingValue
        if (-not $mappingPath -or -not (Test-Path -LiteralPath $mappingPath -PathType Leaf)) {
            throw "Source URL filter missing for $($Config.Name). Configure Source.UrlsFile or Comparison.PathMappingsFile."
        }
        $urls = @()
        foreach ($rawLine in (Get-Content -LiteralPath $mappingPath)) {
            $line = $rawLine.Trim()
            if (-not $line -or $line.StartsWith('#')) { continue }
            $parts = @($line -split '[\t;, ]+' | Where-Object { $_ })
            if ($parts.Count -ne 2) {
                throw "Invalid source mapping in $mappingPath`: expected two URL fields."
            }
            $urls += $parts[0]
        }
    }

    $urls = @($urls | Select-Object -Unique)
    if ($urls.Count -eq 0) { throw "No source URLs found for $($Config.Name)." }
    foreach ($url in $urls) {
        $uri = $null
        if (-not [uri]::TryCreate($url, [uriKind]::Absolute, [ref]$uri) -or
            $uri.Scheme -notin @('http', 'https')) {
            throw "Invalid source URL for $($Config.Name): $url"
        }
    }
    return $urls
}

function Get-SourceMigrations {
    param([string]$Root, [string[]]$Names)
    $migrationsRoot = Get-SmartM365MigrationsRoot -ProjectRoot $Root
    if (-not (Test-Path -LiteralPath $migrationsRoot -PathType Container)) {
        throw "Migrations directory not found: $migrationsRoot"
    }
    $available = @{}
    foreach ($directory in (Get-ChildItem -LiteralPath $migrationsRoot -Directory)) {
        if ($Names.Count -gt 0 -and $Names -notcontains $directory.Name) { continue }
        $configPath = Join-Path $directory.FullName 'migration.config.psd1'
        if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { continue }
        $config = Import-PowerShellDataFile -LiteralPath $configPath
        if ($config.Name -ne $directory.Name) { throw "Migration name differs from folder name in $configPath" }
        if ($directory.Name -eq 'NewMigration') { continue }
        if (-not $config.Source -or $config.Source.Type -notin @('SP2016', 'SP2019')) {
            if ($Names -contains $directory.Name) { throw "Migration $($directory.Name) is not a SharePoint Server source." }
            continue
        }
        if ([string]::IsNullOrWhiteSpace([string]$config.Source.WebApplicationUrl)) {
            throw "Source.WebApplicationUrl is missing in $configPath"
        }
        foreach ($key in @('SourceFileScans', 'SourcePermissionScans', 'Logs')) {
            if (-not $config.Output -or [string]::IsNullOrWhiteSpace([string]$config.Output[$key])) {
                throw "Output.$key is missing in $configPath"
            }
        }
        $available[$directory.Name] = [pscustomobject]@{
            Name = $directory.Name
            Root = $directory.FullName
            Config = $config
            Urls = @(Get-SourceUrls -MigrationRoot $directory.FullName -Config $config)
        }
    }
    if ($Names.Count -gt 0) {
        $selected = New-Object 'System.Collections.Generic.List[object]'
        foreach ($name in $Names) {
            if (-not $available.ContainsKey($name)) { throw "Unknown or unsupported migration: $name" }
            if (@($selected | Where-Object { $_.Name -eq $name }).Count -eq 0) { $selected.Add($available[$name]) }
        }
        return @($selected.ToArray())
    }
    return @($available.Values | Sort-Object Name)
}

function Copy-SourceScript {
    param([string]$RelativePath, [string]$Destination)
    $source = Join-Path $ProjectRoot $RelativePath
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Required script not found: $source" }
    Copy-Item -LiteralPath $source -Destination $Destination -Force -ErrorAction Stop
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne
        (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash) {
        throw "Local script copy differs from source: $Destination"
    }
    $signature = Get-AuthenticodeSignature -LiteralPath $Destination
    if ($signature.Status -ne 'Valid') { throw "Local script signature is not valid: $Destination ($($signature.Status))" }
    Write-BatchLine ("Copied signed script: {0}" -f $Destination)
}

function Copy-VerifiedDependency {
    param([string]$RelativePath, [string]$Destination)
    $source = Join-Path $ProjectRoot $RelativePath
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Required dependency not found: $source" }
    Copy-Item -LiteralPath $source -Destination $Destination -Force -ErrorAction Stop
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne
        (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash) {
        throw "Local dependency copy differs from source: $Destination"
    }
    Write-BatchLine ("Copied dependency: {0}" -f $Destination)
}

function Copy-PortablePython {
    param([string]$SourceDirectory, [string]$DestinationDirectory)
    $sourceRoot = (Get-Item -LiteralPath $SourceDirectory -ErrorAction Stop).FullName.TrimEnd('\')
    $files = @(Get-ChildItem -LiteralPath $sourceRoot -File -Recurse -ErrorAction Stop)
    if ($files.Count -eq 0) { throw "Portable Python directory is empty: $sourceRoot" }
    foreach ($file in $files) {
        $relativePath = $file.FullName.Substring($sourceRoot.Length).TrimStart('\')
        $destination = Join-Path $DestinationDirectory $relativePath
        New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force -ErrorAction Stop | Out-Null
        Copy-Item -LiteralPath $file.FullName -Destination $destination -Force -ErrorAction Stop
        if ((Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash -ne
            (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash) {
            throw "Local Python runtime copy differs from source: $destination"
        }
    }
    Write-BatchLine ("Copied and verified {0} portable Python files locally: {1}" -f $files.Count, $DestinationDirectory)
}

function Get-PythonCommand {
    param([string]$LocalRoot)
    $portableSource = Join-Path $ProjectRoot 'Tools\Python'
    $portableSourceExe = Join-Path $portableSource 'python.exe'
    if (Test-Path -LiteralPath $portableSourceExe -PathType Leaf) {
        $portableLocal = Join-Path $LocalRoot 'Python'
        Copy-PortablePython -SourceDirectory $portableSource -DestinationDirectory $portableLocal
        $portableExe = Join-Path $portableLocal 'python.exe'
        $previousPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            & $portableExe --version *> $null
            $pythonExit = $LASTEXITCODE
        }
        finally { $ErrorActionPreference = $previousPreference }
        if ($pythonExit -ne 0) { throw "Local portable Python could not start: $portableExe" }
        return [pscustomobject]@{ Executable = $portableExe; Arguments = @() }
    }
    foreach ($name in @('python', 'py')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if (-not $command) { continue }
        if ([string]$command.Source -like '\\*') { continue }
        $prefix = if ($name -eq 'py') { @('-3') } else { @() }
        $previousPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            & $command.Source @prefix --version *> $null
            $pythonExit = $LASTEXITCODE
        }
        finally { $ErrorActionPreference = $previousPreference }
        if ($pythonExit -eq 0) { return [pscustomobject]@{ Executable = $command.Source; Arguments = $prefix } }
    }
    throw "Python 3 is required to create scan manifests. Expected portable Python at: $portableSourceExe"
}

function Add-BatchResult {
    param([object]$Job, [string]$Status, [int]$ExitCode, [string]$OutputPath, [string]$LogPath)
    $script:Results.Add([pscustomobject]@{
        Migration = $Job.Name
        Action = $Job.Action
        Status = $Status
        ExitCode = $ExitCode
        Started = $Job.Started.ToString('o')
        Finished = (Get-Date).ToString('o')
        OutputCsv = $OutputPath
        RunLog = $LogPath
    })
    Write-BatchLine ("{0}: {1} {2} (exit {3}); CSV: {4}; log: {5}" -f $Status, $Job.Name, $Job.Action, $ExitCode, $OutputPath, $LogPath)
}

function Update-BatchMigrationComplexity {
    # Informative score, computed right after the migration's last planned phase; never changes the batch result.
    param([object]$Migration, [object]$Python, [string]$LocalCompare)
    try {
        $scanned = @($script:Results | Where-Object { $_.Migration -eq $Migration.Name -and $_.Status -eq 'SUCCESS' }).Count -gt 0
        if (-not $scanned -or -not $script:ComplexityHelpersLoaded) { return }
        if ($null -eq $script:ComplexityDependenciesReady) {
            $script:ComplexityDependenciesReady = $false
            foreach ($name in @('migration_complexity.py', 'compare_sp_source_target_file_inventories.py',
                    'compare_sp_source_target_permissions.py', 'report_html.py', 'comparison_artifacts.py',
                    'comparison_history_chart.py')) {
                Copy-VerifiedDependency -RelativePath ("Scripts\Compare\{0}" -f $name) -Destination (Join-Path $LocalCompare $name)
            }
            $script:ComplexityDependenciesReady = $true
        }
        if (-not $script:ComplexityDependenciesReady) { return }
        $result = Invoke-SmartM365MigrationComplexity -Migration $Migration -ProjectRoot $ProjectRoot `
            -SourceType ([string]$Migration.Config.Source.Type) -PythonExecutable $Python.Executable `
            -PythonArguments @($Python.Arguments) -ScriptPath (Join-Path $LocalCompare 'migration_complexity.py')
        foreach ($line in @($result.Output)) { Write-BatchLine ([string]$line) }
        if ($result.Status -eq 'Succeeded') {
            Write-BatchLine ("Migration complexity updated for {0}: {1}" -f $Migration.Name, $result.Message)
        }
        else {
            Write-BatchLine ("WARNING Migration complexity not updated for {0} ({1}); the GUI will calculate it: {2}" -f $Migration.Name, $result.Status, $result.Message)
        }
    }
    catch {
        Write-BatchLine ("WARNING Migration complexity skipped for {0}; the GUI will calculate it: {1}" -f $Migration.Name, $_.Exception.Message)
    }
    finally {
        if ($script:ComplexityPendingMarkers -and $script:ComplexityPendingMarkers.ContainsKey($Migration.Name)) {
            Clear-SmartM365ComplexityPending -Path $script:ComplexityPendingMarkers[$Migration.Name]
            $script:ComplexityPendingMarkers.Remove($Migration.Name)
        }
    }
}

$requestedNames = New-Object 'System.Collections.Generic.List[string]'
foreach ($value in $MigrationNames) {
    foreach ($name in ($value -split ',')) {
        if (-not [string]::IsNullOrWhiteSpace($name)) { $requestedNames.Add($name.Trim()) }
    }
}
if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Join-Path $PSScriptRoot '..\..\..'
}
$ProjectRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
# The entry script runs from a local copy, so the storage helper is read from the toolkit.
$storageHelper = Join-Path $ProjectRoot 'Scripts\Launchers\SmartM365-SharePointMigration-Storage.ps1'
if (-not (Test-Path -LiteralPath $storageHelper -PathType Leaf)) {
    $storageHelper = Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-Storage.ps1'
}
. $storageHelper
$script:ComplexityHelpersLoaded = $false
$script:ComplexityDependenciesReady = $null
$script:ComplexityPendingMarkers = @{}
try {
    . (Join-Path $ProjectRoot 'Scripts\Launchers\SmartM365-SharePointMigration-LauncherCommon.ps1')
    $script:ComplexityHelpersLoaded = $true
}
catch { Write-BatchLine ("WARNING Migration complexity helpers unavailable; the GUI will calculate it: {0}" -f $_.Exception.Message) }
$LocalRunRoot = [System.IO.Path]::GetFullPath($LocalRunRoot)
if ($LocalRunRoot.StartsWith('\\')) { throw 'LocalRunRoot must be on a local disk of the farm server.' }
if (-not $PlanOnly) {
    if (-not $BatchId) { $BatchId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)) }
    $batchRoot = Get-SmartM365MigrationLogsRoot -ProjectRoot $ProjectRoot -ChildPath "source-scan-batches\$batchId"
    Write-BatchLine ("Creating source batch log: {0}" -f (Join-Path $batchRoot 'batch.log'))
    New-Item -ItemType Directory -Path $batchRoot -ErrorAction Stop | Out-Null
    $script:BatchLogPath = Join-Path $batchRoot 'batch.log'
    Write-BatchLine ("Source batch preflight started. Logs: {0}" -f $batchRoot)
}
try {
Write-BatchLine ("Loading source migration configurations from: {0}" -f (Get-SmartM365MigrationsRoot -ProjectRoot $ProjectRoot))
$migrations = @(Get-SourceMigrations -Root $ProjectRoot -Names $requestedNames.ToArray())
if ($migrations.Count -eq 0) { throw 'No configured SharePoint Server source migrations found.' }
$actions = @()
if ($InventoryMode -in @('FilesOnly', 'Both')) { $actions += 'ScanSourceFiles' }
if ($InventoryMode -in @('PermissionsOnly', 'Both')) { $actions += 'ScanSourcePermissions' }

Write-BatchLine ("Source scan plan: {0} migrations; {1} scans; maximum parallel=1." -f $migrations.Count, ($migrations.Count * $actions.Count))
foreach ($action in $actions) {
    foreach ($migration in $migrations) {
        Write-BatchLine ("PLAN {0} {1} {2} ({3} source URLs)" -f $action, $migration.Name, $migration.Config.Source.WebApplicationUrl, $migration.Urls.Count)
    }
}
if ($PlanOnly) {
    Write-BatchLine 'Plan only: no scan started and no output created.'
    return
}

Write-BatchLine 'Checking Windows PowerShell 5.1 and the SharePoint snap-in.'
$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $windowsPowerShell -PathType Leaf)) {
    throw 'Windows PowerShell 5.1 is required on the SharePoint farm server.'
}
& $windowsPowerShell -NoProfile -Command "if (Get-PSSnapin -Registered -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) { exit 0 } else { exit 2 }"
if ($LASTEXITCODE -ne 0) {
    throw 'Microsoft.SharePoint.PowerShell snap-in is not registered. Run this batch on a SharePoint farm server.'
}
Write-BatchLine 'Checking the scan manifest script on the shared project.'
$manifestSource = Join-Path $ProjectRoot 'Scripts\Compare\scan_evidence.py'
if (-not (Test-Path -LiteralPath $manifestSource -PathType Leaf)) { throw "Scan manifest script not found: $manifestSource" }
}
catch {
    Write-BatchLine ("Source batch preflight failed: {0}" -f $_.Exception.Message)
    throw
}

$localRoot = Join-Path $LocalRunRoot $batchId
$localInventory = Join-Path $localRoot 'Inventory'
$localLaunchers = Join-Path $localRoot 'Launchers'
$localCompare = Join-Path $localRoot 'Compare'
New-Item -ItemType Directory -Path $localInventory, $localLaunchers, $localCompare -Force | Out-Null
Write-BatchLine ("Batch started. Logs: {0}; local execution copy: {1}" -f $batchRoot, $localRoot)

$fileScript = Join-Path $localInventory 'SmartM365-SharePointSource-FileInventory.ps1'
$permissionScript = Join-Path $localInventory 'SmartM365-SharePointSource-PermissionInventory.ps1'
$script:BatchAborted = $false

try {
    Copy-SourceScript -RelativePath 'Scripts\Inventory\SmartM365-SharePointSource-FileInventory.ps1' -Destination $fileScript
    Copy-SourceScript -RelativePath 'Scripts\Inventory\SmartM365-SharePointSource-PermissionInventory.ps1' -Destination $permissionScript
    Copy-SourceScript -RelativePath 'Scripts\Inventory\SmartM365-SharePointSource-Assessment.ps1' -Destination (Join-Path $localInventory 'SmartM365-SharePointSource-Assessment.ps1')
    $localAssessmentConfig = Join-Path $localRoot 'Config'
    [void](New-Item -ItemType Directory -Path $localAssessmentConfig -Force)
    Copy-VerifiedDependency -RelativePath 'Config\source-assessment.rules.json.template' -Destination (Join-Path $localAssessmentConfig 'source-assessment.rules.json.template')
    if (Test-Path -LiteralPath (Join-Path $ProjectRoot 'Config\source-assessment.rules.json.txt')) {
        Copy-VerifiedDependency -RelativePath 'Config\source-assessment.rules.json.txt' -Destination (Join-Path $localAssessmentConfig 'source-assessment.rules.json.txt')
    }
    Copy-SourceScript -RelativePath 'Scripts\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1' -Destination (Join-Path $localLaunchers 'SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
    $manifestScript = Join-Path $localCompare 'scan_evidence.py'
    Copy-VerifiedDependency -RelativePath 'Scripts\Compare\scan_evidence.py' -Destination $manifestScript
    Copy-VerifiedDependency -RelativePath 'Scripts\console_lifecycle.py' -Destination (Join-Path $localRoot 'console_lifecycle.py')
    $python = Get-PythonCommand -LocalRoot $localRoot
    if ($script:ComplexityHelpersLoaded) {
        # Tell the GUI that this batch will calculate complexity after each migration's last phase.
        foreach ($migration in $migrations) {
            $script:ComplexityPendingMarkers[$migration.Name] = Set-SmartM365ComplexityPending -MigrationRoot $migration.Root -Reason "Source scan batch $batchId"
        }
    }
    foreach ($action in $actions) {
        Write-BatchLine ("Starting phase: {0}" -f $action)
        foreach ($migration in $migrations) {
            $job = [pscustomobject]@{ Name = $migration.Name; Action = $action; Started = Get-Date }
            $outputPath = ''
            $logPath = ''
            try {
            $urlsFile = Join-Path $batchRoot ("{0}-source-urls.txt" -f $migration.Name)
            [System.IO.File]::WriteAllLines($urlsFile, [string[]]$migration.Urls, [System.Text.UTF8Encoding]::new($true))
            $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
            $kind = if ($action -eq 'ScanSourceFiles') { 'File' } else { 'Permission' }
            $outputKey = if ($kind -eq 'File') { 'SourceFileScans' } else { 'SourcePermissionScans' }
            $outputDir = Resolve-MigrationFile -MigrationRoot $migration.Root -Value ([string]$migration.Config.Output[$outputKey])
            $logsDir = Resolve-MigrationFile -MigrationRoot $migration.Root -Value ([string]$migration.Config.Output.Logs)
            New-Item -ItemType Directory -Path $outputDir, $logsDir -Force | Out-Null
            $outputPath = Join-Path $outputDir ("{0}-{1}Inventory-{2}-{3}.csv" -f $migration.Config.Source.Type, $kind, $migration.Name, $timestamp)
            $logPath = Join-Path $logsDir ("{0}-{1}-{2}.log" -f $migration.Name, $action, $timestamp)
            $scriptPath = if ($kind -eq 'File') { $fileScript } else { $permissionScript }
            $commandText = '& {0} -WebApplicationUrl {1} -OutputPath {2} -LogPath {3} -UseSiteUrlFilter -SiteUrlsFile {4}' -f @(
                (Quote-PowerShellLiteral $scriptPath),
                (Quote-PowerShellLiteral ([string]$migration.Config.Source.WebApplicationUrl)),
                (Quote-PowerShellLiteral $outputPath),
                (Quote-PowerShellLiteral $logPath),
                (Quote-PowerShellLiteral $urlsFile))
            if ($kind -eq 'Permission') {
                $commandText += ' -ItemProgressInterval {0} -IncludeItemPermissions:${1}' -f @(
                    [int]$migration.Config.Permissions.ItemProgressInterval,
                    ([bool]$migration.Config.Permissions.IncludeItemPermissions).ToString().ToLowerInvariant())
                if ([bool]$migration.Config.Permissions.SourceDocumentLibrariesOnly) { $commandText += ' -DocumentLibrariesOnly' }
            }

            Write-BatchLine ("START {0} {1}; {2} source URLs" -f $migration.Name, $action, $migration.Urls.Count)
            $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($commandText))
            & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand
            $code = $LASTEXITCODE
            if ($null -eq $code) { $code = 1 }
            $errorPath = Join-Path $outputDir ("{0}-Errors.csv" -f [System.IO.Path]::GetFileNameWithoutExtension($outputPath))
            if ($code -eq 0 -and (-not (Test-Path -LiteralPath $outputPath -PathType Leaf) -or
                (Test-Path -LiteralPath $errorPath -PathType Leaf))) {
                Write-BatchLine ("Scan reported success without a complete CSV, or it recorded errors: {0}" -f $outputPath)
                $code = 1
            }
            if ($code -eq 0 -and $kind -eq 'File' -and
                -not (Test-Path -LiteralPath "$outputPath.metrics.json.txt" -PathType Leaf)) {
                Write-BatchLine ("File scan metrics are missing: {0}.metrics.json.txt" -f $outputPath)
                $code = 1
            }
            if ($code -eq 0) {
                $previousPreference = $ErrorActionPreference
                try {
                    $ErrorActionPreference = 'Continue'
                    $manifestOutput = & $python.Executable @($python.Arguments) $manifestScript --csv $outputPath --side Source --kind $kind --scope $urlsFile 2>&1
                    $manifestExit = $LASTEXITCODE
                }
                finally { $ErrorActionPreference = $previousPreference }
                foreach ($line in @($manifestOutput)) { Write-BatchLine ([string]$line) }
                if ($manifestExit -ne 0 -or -not (Test-Path -LiteralPath "$outputPath.manifest.json.txt" -PathType Leaf)) {
                    Write-BatchLine ("Scan manifest was not published: {0}.manifest.json.txt" -f $outputPath)
                    $code = 1
                }
            }
            Add-BatchResult -Job $job -Status $(if ($code -eq 0) { 'SUCCESS' } else { 'FAILED' }) -ExitCode $code -OutputPath $outputPath -LogPath $logPath
            }
            catch {
                Write-BatchLine ("FAILED {0} {1}: {2}" -f $migration.Name, $action, $_.Exception.Message)
                Add-BatchResult -Job $job -Status 'FAILED' -ExitCode 1 -OutputPath $outputPath -LogPath $logPath
            }
            if ($action -eq $actions[-1]) {
                Update-BatchMigrationComplexity -Migration $migration -Python $python -LocalCompare $localCompare
            }
        }
    }
}
catch {
    $script:BatchAborted = $true
    Write-BatchLine ("Batch aborted: {0}" -f $_.Exception.Message)
    throw
}
finally {
    foreach ($marker in @($script:ComplexityPendingMarkers.Values)) { Clear-SmartM365ComplexityPending -Path $marker }
    $summaryPath = Join-Path $batchRoot 'summary.csv'
    if ($script:Results.Count -gt 0) {
        $script:Results.ToArray() | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8
    }
    else {
        Set-Content -LiteralPath $summaryPath -Value 'Migration,Action,Status,ExitCode,Started,Finished,OutputCsv,RunLog' -Encoding UTF8
    }
    $failed = @($script:Results | Where-Object { $_.Status -ne 'SUCCESS' }).Count
    if ($script:BatchAborted) { $failed++ }
    Write-BatchLine ("Batch finished: {0} successful, {1} failed. Summary: {2}" -f ($script:Results.Count - $failed), $failed, $summaryPath)
}
if (@($script:Results | Where-Object { $_.Status -ne 'SUCCESS' }).Count -gt 0) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA+f9F65EmWOPiV
# C45G05O2WDkdHy5Qq3OyYmFw/SdCWaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIARwNU/6Pbfk5Whh2rvt5wyekvtRqQAOLdNBEnm7Ml4oMA0GCSqG
# SIb3DQEBAQUABIIBgDwvNthozGWMYspYAsezR+ciUEfHEkIVzjeD6tQKLoaxyRjm
# 9M3jxqLdx6DpKACMnPyY+xVdu3f8AmetjoSorttPp8iGC58rj5hYF+TXKgMHjQps
# qj06Znnyf4aN3QWiOnu3yLMYUXikHrLEhjEzox/1WSYE/KpMG8t77H9r/EKqcArw
# VNWQ1S5BnPb0jbIN8+DBOSMpTj+ytJDSjEnZq9nGEysf51ovs94r1VouKijOCs57
# OUfFz7uPhZcnpnt1aYGHhDh2AGxfkCz9qf4zk51aFF+68DZtlTPuEGrXBX4AZZe3
# hSD3IhlB9j/JRu00oHDPd21CfJvrIJT6ckwDZ7sEXMBlPdmXtxMnR+GbrP9kEs98
# fajtsMn+p5b9yTTRHaaMxkgdI8O4oRd5QD/t1WTh5vcZ25RlMTyGNCpInIE1kaD5
# LXygfzsq/hNJIWIRhQyfne5Cg7bDJvJY31ZUx5KeMrevJSSSbDz3F6lOPasa+CgP
# Q7EaNzk9KYWeyuexcaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzVaMC8GCSqGSIb3DQEJBDEiBCB4ppy2+fkrRFmlocz7DSkVOu2XCegO5oRWWwT0
# VMMNdjANBgkqhkiG9w0BAQEFAASCAgAPp3V9Vu/J8M1aKZntvKjE3FjkRp1OcSAS
# 1FjT+V9lcLtht9Eu9q5IQGRtQpDLL3Kp4wdTaJw5aa9gXpyET07mHUOMj7etzad5
# WqDzKy6jGlGYgvPygxHbF4POjjhKNJhBMWgITC5iq6r23L59NOCLZeZ98LmtawKT
# jXCgyrXuUn/PBVheI25WmnPlXhISEqVG7CqfGwBH0QBh1BVM/ASl4WFTCtctfZo1
# SFMJIE0F2kL1MqVK5xuUlEj7G0SCeAowLSqoIkhiNzRBZSApUA1p4iNQ5SOZRrIG
# whTM6u55eBqgTrSKigu0E28dSLDZKWw6lDJP96eEKhYuQHZCuLR2kXagJFcBKYUx
# UvUTU9PtDflmHJ2P2AoQtMd6UQrAO3c7aBKnVwRe5R8JPHvJhZvUXnCIKe+p0knK
# pUbdJvSxWFQqIayQoqi7wT0Six8d9pXa4fnMD8a3COUG+k3bwBmycHCajzMibUSs
# ECWeV63uHMjbIXXqcUvNNbdRTbX4YAN+TCdOGQP2la1EMi6OVddSHKS4vmdDPb45
# JMGGTOm3Gw9YAS9C6uEqHyAUiF7kBGxxXqwyz7nGem4kHKekKqERFi4ylpZ5ZEDo
# zrOxua9nyCM5Dl+1B//5nTtLEpPGxffgAXCRLZkzwqJh3pNhFhVfNh8w+2WuLe4M
# NnhmA/koVw==
# SIG # End signature block
