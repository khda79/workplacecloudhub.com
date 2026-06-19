<#
.SYNOPSIS
    Generic launcher for a migration folder.

.DESCRIPTION
    Reads Migrations\<MigrationName>\migration.config.psd1 and runs the
    requested inventory, comparison, or permission action.
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
        'CompareSourceHistory'
    )]
    [string]$Action,

    [string]$SourceCsv,

    [string]$TargetCsv,

    [switch]$ForceAuthentication = $true,

    [switch]$DeviceLogin,

    [switch]$UseCertificate,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
. (Join-Path -Path $PSScriptRoot -ChildPath '..\SmartM365-SharePointMigration-LauncherCommon.ps1')

$ProjectRoot = Get-LauncherProjectRoot -LauncherScriptRoot $PSScriptRoot
$Migration = Import-LauncherMigrationConfig -ProjectRoot $ProjectRoot -MigrationName $MigrationName
$MigrationRoot = $Migration.Root
$Config = $Migration.Data

if ($MigrationName -eq '_Template' -or $Config.Name -eq 'NewMigration') {
    throw "The migration template cannot be executed directly. Copy Migrations\_Template to Migrations\<MigrationName>, update migration.config.psd1, then run the launcher from the copied migration."
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

function New-MigrationLogPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action,

        [string]$Timestamp = (Get-Date -Format 'yyyyMMdd-HHmmss')
    )

    Join-Path -Path (Get-MigrationLogDirectory) -ChildPath ("{0}-{1}-{2}.log" -f $Config.Name, $Action, $Timestamp)
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
        $resolvedPath = (Resolve-Path -LiteralPath $Path).ProviderPath
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
                $explorerPath = Join-Path -Path $uncRoot -ChildPath $resolvedPath.Substring($root.Length)
            }
        }

        Write-Info ("Opening comparison directory: {0}" -f $explorerPath) Cyan
        Start-Process -FilePath explorer.exe -ArgumentList ('"{0}"' -f $explorerPath) | Out-Null
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

function Assert-CsvScanAgeDifference {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceCsvPath,

        [Parameter(Mandatory = $true)]
        [string]$TargetCsvPath,

        [Parameter(Mandatory = $true)]
        [double]$MaxAgeDifferenceHours,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if ($MaxAgeDifferenceHours -lt 0) {
        Write-Info ("Scan age check disabled for {0}." -f $Label) DarkYellow
        return
    }

    $sourceItem = Get-Item -LiteralPath $SourceCsvPath
    $targetItem = Get-Item -LiteralPath $TargetCsvPath
    $ageDifference = ($targetItem.LastWriteTime - $sourceItem.LastWriteTime).Duration()
    $message = "{0} scan age difference is {1:n2}h. Maximum allowed: {2:n2}h. Source: {3}; Target: {4}" -f $Label, $ageDifference.TotalHours, $MaxAgeDifferenceHours, $sourceItem.LastWriteTime, $targetItem.LastWriteTime

    if ($ageDifference.TotalHours -gt $MaxAgeDifferenceHours) {
        if ($Force) {
            Write-Info ("WARNING: {0} Continuing because -Force was used." -f $message) Yellow
            return
        }

        throw ("{0} Rerun source/target scans closer together, or use -Force only after reviewing the risk." -f $message)
    }

    Write-Info $message DarkCyan
}

function Invoke-PythonScript {
    param(
        [string]$Script,
        [string[]]$Arguments
    )

    $pythonCommand = Get-PythonCommand
    Write-Info ("Using Python: {0} ({1})" -f $pythonCommand.Executable, $pythonCommand.Source) Cyan
    & $pythonCommand.Executable @($pythonCommand.Arguments) $Script @Arguments 2>&1 | ForEach-Object {
        Write-Info ([string]$_)
    }

    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "Python script failed with exit code $exitCode`: $Script"
    }
}

function Invoke-SourceFileScan {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Inventory\SmartM365-SharePointSource-FileInventory.ps1'
    $outputDirectory = Resolve-MigrationPath $Config.Output.SourceFileScans
    $urlsFile = Resolve-MigrationPath $Config.Source.UrlsFile
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

    $parameters = @{
        WebApplicationUrl = $Config.Source.WebApplicationUrl
        UseSiteUrlFilter = $true
        SiteUrlsFile = $urlsFile
        OutputPath = (Join-Path -Path $outputDirectory -ChildPath ("SP2019-FileInventory-{0}-{1}.csv" -f $Config.Name, $timestamp))
        LogPath = (New-MigrationLogPath -Action 'ScanSourceFiles' -Timestamp $timestamp)
    }

    Write-Info ("Starting source file scan for migration '{0}'" -f $Config.Name) Cyan
    & $scriptPath @parameters
}

function Invoke-TargetFileScan {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Inventory\SmartM365-SharePointTarget-FileInventory.ps1'
    $outputDirectory = Resolve-MigrationPath $Config.Output.TargetFileScans
    $urlsFile = Resolve-MigrationPath $Config.Target.UrlsFile
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

    $parameters = @{
        WebUrlsFile = $urlsFile
        OutputPath = (Join-Path -Path $outputDirectory -ChildPath ("SPO-FileInventory-{0}-{1}.csv" -f $Config.Name, $timestamp))
        LogPath = (New-MigrationLogPath -Action 'ScanTargetFiles' -Timestamp $timestamp)
    }

    $authConfig = Get-SPOAuthConfig
    foreach ($key in @('ClientId', 'Tenant', 'TenantId', 'Thumbprint')) {
        if ($authConfig.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$authConfig[$key])) {
            $parameters[$key] = $authConfig[$key]
        }
    }

    if ($UseCertificate) {
        if (-not $parameters.ContainsKey('ClientId') -or -not $parameters.ContainsKey('Thumbprint')) {
            throw "Certificate authentication requires ClientId and Thumbprint in Config\SPOAuth.local.psd1."
        }
    }
    elseif ($DeviceLogin) {
        $parameters.DeviceLogin = $true
    }
    else {
        $parameters.Interactive = $true
    }

    if ($ForceAuthentication) {
        $parameters.ForceAuthentication = $true
    }

    Write-Info ("Starting target file scan for migration '{0}'" -f $Config.Name) Cyan
    & $scriptPath @parameters
}

function Invoke-SourcePermissionScan {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Inventory\SmartM365-SharePointSource-PermissionInventory.ps1'
    $outputDirectory = Resolve-MigrationPath $Config.Output.SourcePermissionScans
    $urlsFile = Resolve-MigrationPath $Config.Source.UrlsFile
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

    $parameters = @{
        WebApplicationUrl = $Config.Source.WebApplicationUrl
        UseSiteUrlFilter = $true
        SiteUrlsFile = $urlsFile
        OutputPath = (Join-Path -Path $outputDirectory -ChildPath ("SP2019-PermissionInventory-{0}-{1}.csv" -f $Config.Name, $timestamp))
        LogPath = (New-MigrationLogPath -Action 'ScanSourcePermissions' -Timestamp $timestamp)
        DocumentLibrariesOnly = [bool]$Config.Permissions.SourceDocumentLibrariesOnly
        IncludeItemPermissions = [bool]$Config.Permissions.IncludeItemPermissions
        ItemProgressInterval = [int]$Config.Permissions.ItemProgressInterval
    }

    Write-Info ("Starting source permission scan for migration '{0}'" -f $Config.Name) Cyan
    & $scriptPath @parameters
}

function Invoke-TargetPermissionScan {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Inventory\SmartM365-SharePointTarget-PermissionInventory.ps1'
    $outputDirectory = Resolve-MigrationPath $Config.Output.TargetPermissionScans
    $urlsFile = Resolve-MigrationPath $Config.Target.UrlsFile
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

    $parameters = @{
        WebUrlsFile = $urlsFile
        OutputPath = (Join-Path -Path $outputDirectory -ChildPath ("SPO-PermissionInventory-{0}-{1}.csv" -f $Config.Name, $timestamp))
        LogPath = (New-MigrationLogPath -Action 'ScanTargetPermissions' -Timestamp $timestamp)
        DocumentLibrariesOnly = [bool]$Config.Permissions.TargetDocumentLibrariesOnly
        IncludeItemPermissions = [bool]$Config.Permissions.IncludeItemPermissions
        ItemProgressInterval = [int]$Config.Permissions.ItemProgressInterval
    }

    $authConfig = Get-SPOAuthConfig
    foreach ($key in @('ClientId', 'Tenant', 'TenantId', 'Thumbprint')) {
        if ($authConfig.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$authConfig[$key])) {
            $parameters[$key] = $authConfig[$key]
        }
    }

    if ($UseCertificate) {
        if (-not $parameters.ContainsKey('ClientId') -or -not $parameters.ContainsKey('Thumbprint')) {
            throw "Certificate authentication requires ClientId and Thumbprint in Config\SPOAuth.local.psd1."
        }
    }
    elseif ($DeviceLogin) {
        $parameters.DeviceLogin = $true
    }
    else {
        $parameters.Interactive = $true
    }

    if ($ForceAuthentication) {
        $parameters.ForceAuthentication = $true
    }

    Write-Info ("Starting target permission scan for migration '{0}'" -f $Config.Name) Cyan
    & $scriptPath @parameters
}

function Invoke-FileComparison {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $logPath = New-MigrationLogPath -Action 'CompareFiles' -Timestamp $timestamp
    $sourceDirectory = Resolve-MigrationPath $Config.Output.SourceFileScans
    $targetDirectory = Resolve-MigrationPath $Config.Output.TargetFileScans
    $outputDirectory = Join-Path -Path (Resolve-MigrationPath $Config.Output.FileComparisons) -ChildPath ("{0}-files-{1}" -f $Config.Name, $timestamp)
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    Start-Transcript -Path $logPath -Force -WhatIf:$false | Out-Null

    try {
        Write-Info ("Run log: {0}" -f $logPath) Cyan
        if ([string]::IsNullOrWhiteSpace($SourceCsv)) {
            $SourceCsv = Get-LatestCsv -Directory $sourceDirectory -Filter ("SP2019-FileInventory-{0}-*.csv" -f $Config.Name) -Recurse
        }
        if ([string]::IsNullOrWhiteSpace($TargetCsv)) {
            $TargetCsv = Get-LatestCsv -Directory $targetDirectory -Filter ("SPO-FileInventory-{0}-*.csv" -f $Config.Name)
        }

        $maxScanAgeDifferenceHours = [double](Get-ComparisonConfigValue -Name 'MaxScanAgeDifferenceHours' -DefaultValue 12)
        $modifiedDateToleranceMinutes = [double](Get-ComparisonConfigValue -Name 'ModifiedDateToleranceMinutes' -DefaultValue 0)
        $sourceModifiedTimeZone = [string](Get-ComparisonConfigValue -Name 'SourceModifiedTimeZone' -DefaultValue 'Local')
        $targetModifiedTimeZone = [string](Get-ComparisonConfigValue -Name 'TargetModifiedTimeZone' -DefaultValue 'UTC')
        Assert-CsvScanAgeDifference `
            -SourceCsvPath $SourceCsv `
            -TargetCsvPath $TargetCsv `
            -MaxAgeDifferenceHours $maxScanAgeDifferenceHours `
            -Label 'File inventory'

        $compareScript = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Compare\compare_sp_source_target_file_inventories.py'
        Invoke-PythonScript -Script $compareScript -Arguments @(
            '--source-csv', $SourceCsv,
            '--target-csv', $TargetCsv,
            '--output-directory', $outputDirectory,
            '--source-web-urls-file', (Resolve-MigrationPath $Config.Source.UrlsFile),
            '--target-web-urls-file', (Resolve-MigrationPath $Config.Target.UrlsFile),
            '--target-prefix', $Config.Target.PrefixToRemove,
            '--comparison-name', ("{0}-SP2019-vs-SPO" -f $Config.Name),
            '--size-tolerance-bytes', ([string]$Config.Comparison.SizeToleranceBytes),
            '--modified-date-tolerance-minutes', ([string]$modifiedDateToleranceMinutes),
            '--source-modified-time-zone', $sourceModifiedTimeZone,
            '--target-modified-time-zone', $targetModifiedTimeZone,
            '--sharegate-replacement-character', ([string]$Config.Comparison.ShareGateReplacementCharacter)
        )

        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Export\export_comparison_to_excel.py') -Arguments @(
            '--comparison-directory', $outputDirectory,
            '--output-xlsx', (Join-Path $outputDirectory ("{0}-SP2019-vs-SPO-Comparison-{1}.xlsx" -f $Config.Name, $timestamp))
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

        Write-Info ("File comparison completed: {0}" -f $outputDirectory) Green
        Open-DirectoryInExplorer -Path $outputDirectory
    }
    finally {
        Stop-LauncherTranscript -Path $logPath
    }
}

function Invoke-PermissionComparison {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $logPath = New-MigrationLogPath -Action 'ComparePermissions' -Timestamp $timestamp
    $sourceDirectory = Resolve-MigrationPath $Config.Output.SourcePermissionScans
    $targetDirectory = Resolve-MigrationPath $Config.Output.TargetPermissionScans
    $outputDirectory = Join-Path -Path (Resolve-MigrationPath $Config.Output.PermissionComparisons) -ChildPath ("{0}-permissions-{1}" -f $Config.Name, $timestamp)
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    Start-Transcript -Path $logPath -Force -WhatIf:$false | Out-Null

    try {
        Write-Info ("Run log: {0}" -f $logPath) Cyan
        if ([string]::IsNullOrWhiteSpace($SourceCsv)) {
            $SourceCsv = Get-LatestCsv -Directory $sourceDirectory -Filter ("SP2019-PermissionInventory-{0}-*.csv" -f $Config.Name) -Recurse
        }
        if ([string]::IsNullOrWhiteSpace($TargetCsv)) {
            $TargetCsv = Get-LatestCsv -Directory $targetDirectory -Filter ("SPO-PermissionInventory-{0}-*.csv" -f $Config.Name)
        }

        Invoke-PythonScript -Script (Join-Path $ProjectRoot 'Scripts\Compare\compare_sp_source_target_permissions.py') -Arguments @(
            '--source-csv', $SourceCsv,
            '--target-csv', $TargetCsv,
            '--output-directory', $outputDirectory,
            '--source-root-path', $Config.Source.PermissionRootPath,
            '--target-root-path', $Config.Target.PermissionRootPath,
            '--comparison-name', ("{0}-SP2019-vs-SPO-Permissions" -f $Config.Name)
        )

        Write-Info ("Permission comparison completed: {0}" -f $outputDirectory) Green
        Open-DirectoryInExplorer -Path $outputDirectory
    }
    finally {
        Stop-LauncherTranscript -Path $logPath
    }
}

function Invoke-SourceHistoryComparison {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $scriptPath = Join-Path -Path $ProjectRoot -ChildPath 'Scripts\Compare\SmartM365-SharePointSource-FileInventoryHistoryCompare.ps1'
    $outputDirectory = Join-Path -Path (Resolve-MigrationPath $Config.Output.SourceHistoryComparisons) -ChildPath ("SP2019-Changes-{0}" -f $timestamp)

    $parameters = @{
        ScanDirectory = Resolve-MigrationPath $Config.Output.SourceFileScans
        InventoryNameFilter = ("*{0}*" -f $Config.Name)
        WebUrlsFile = Resolve-MigrationPath $Config.Source.UrlsFile
        OutputDirectory = $outputDirectory
        LogPath = (New-MigrationLogPath -Action 'CompareSourceHistory' -Timestamp $timestamp)
        ComparisonName = ("{0}-SP2019-vs-SP2019" -f $Config.Name)
        SizeToleranceBytes = [long]$Config.Comparison.SizeToleranceBytes
        ModifiedDateToleranceMinutes = [double](Get-ComparisonConfigValue -Name 'ModifiedDateToleranceMinutes' -DefaultValue 0)
    }

    if ($Force) {
        $parameters.Force = $true
    }

    Write-Info ("Starting source history comparison for migration '{0}'" -f $Config.Name) Cyan
    & $scriptPath @parameters
    Open-DirectoryInExplorer -Path $outputDirectory
}

switch ($Action) {
    'ScanSourceFiles' { Invoke-SourceFileScan }
    'ScanTargetFiles' { Invoke-TargetFileScan }
    'CompareFiles' { Invoke-FileComparison }
    'ScanSourcePermissions' { Invoke-SourcePermissionScan }
    'ScanTargetPermissions' { Invoke-TargetPermissionScan }
    'ComparePermissions' { Invoke-PermissionComparison }
    'CompareSourceHistory' { Invoke-SourceHistoryComparison }
}
