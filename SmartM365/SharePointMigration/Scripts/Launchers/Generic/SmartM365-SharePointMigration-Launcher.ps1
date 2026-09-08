<#
.SYNOPSIS
    Generic launcher for a migration folder.

.DESCRIPTION
    Reads Migrations\<MigrationName>\migration.config.psd1 and runs the
    requested inventory, comparison, or permission action.

.VERSION
    1.0.18
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
        'CompareScanHistory'
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
$script:LauncherNonInteractive = [bool]$NonInteractive
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

        [string]$Label = 'Inventory'
    )

    $sourceItem = Get-Item -LiteralPath $SourceCsvPath -ErrorAction Stop
    $targetItem = Get-Item -LiteralPath $TargetCsvPath -ErrorAction Stop
    $ageDifference = ($sourceItem.LastWriteTime - $targetItem.LastWriteTime).Duration()

    if ($ageDifference.TotalHours -le $MaxAgeDifferenceHours) {
        return
    }

    $message = ("{0} scan age difference is {1:n2}h. Maximum allowed: {2:n2}h. Source: {3}; Target: {4}" -f `
        $Label,
        $ageDifference.TotalHours,
        $MaxAgeDifferenceHours,
        $sourceItem.LastWriteTime,
        $targetItem.LastWriteTime)

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
    $outputPath = Join-Path -Path $generatedDirectory -ChildPath ("{0}-{1}-urls-from-mapping.txt" -f $Config.Name, $sideName)
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

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
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
}

function Invoke-PermissionInventoryScan {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Source', 'Target')]
        [string]$Side
    )

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
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
            $SourceCsv = Get-LatestCsv -Directory $sourceDirectory -Filter ("{0}-FileInventory-{1}-*.csv" -f (Get-MigrationEndpointType -Side 'Source'), $Config.Name) -Recurse
        }
        if ([string]::IsNullOrWhiteSpace($TargetCsv)) {
            $TargetCsv = Get-LatestCsv -Directory $targetDirectory -Filter ("{0}-FileInventory-{1}-*.csv" -f (Get-MigrationEndpointType -Side 'Target'), $Config.Name)
        }

        $maxScanAgeDifferenceHours = [double](Get-ComparisonConfigValue -Name 'MaxScanAgeDifferenceHours' -DefaultValue 12)
        $modifiedDateToleranceMinutes = [double](Get-ComparisonConfigValue -Name 'ModifiedDateToleranceMinutes' -DefaultValue 0)
        $sourceModifiedTimeZone = Get-MigrationEndpointModifiedTimeZone -Side 'Source'
        $targetModifiedTimeZone = Get-MigrationEndpointModifiedTimeZone -Side 'Target'
        Assert-CsvScanAgeDifference `
            -SourceCsvPath $SourceCsv `
            -TargetCsvPath $TargetCsv `
            -MaxAgeDifferenceHours $maxScanAgeDifferenceHours `
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
    $maxAgeHours = [double](Get-ComparisonConfigValue -Name 'EntraUsersCacheMaxAgeHours' -DefaultValue 24)
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
            $SourceCsv = Get-LatestCsv -Directory $sourceDirectory -Filter ("{0}-PermissionInventory-{1}-*.csv" -f (Get-MigrationEndpointType -Side 'Source'), $Config.Name) -Recurse
        }
        if ([string]::IsNullOrWhiteSpace($TargetCsv)) {
            $TargetCsv = Get-LatestCsv -Directory $targetDirectory -Filter ("{0}-PermissionInventory-{1}-*.csv" -f (Get-MigrationEndpointType -Side 'Target'), $Config.Name)
        }

        $maxScanAgeDifferenceHours = [double](Get-ComparisonConfigValue -Name 'PermissionMaxScanAgeDifferenceHours' -DefaultValue 24)
        Assert-CsvScanAgeDifference `
            -SourceCsvPath $SourceCsv `
            -TargetCsvPath $TargetCsv `
            -MaxAgeDifferenceHours $maxScanAgeDifferenceHours `
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
            '--comparison-name', ("{0}-{1}-vs-{2}-Permissions" -f $Config.Name, (Get-MigrationEndpointType -Side 'Source'), (Get-MigrationEndpointType -Side 'Target'))
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

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
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

switch ($Action) {
    'ScanSourceFiles' { Invoke-SourceFileScan }
    'ScanTargetFiles' { Invoke-TargetFileScan }
    'CompareFiles' { Invoke-FileComparison }
    'ScanSourcePermissions' { Invoke-SourcePermissionScan }
    'ScanTargetPermissions' { Invoke-TargetPermissionScan }
    'ComparePermissions' { Invoke-PermissionComparison }
    'CompareSourceHistory' { Invoke-FileHistoryComparison -Side 'Source' -LogAction 'CompareSourceHistory' }
    'CompareScanHistory' { Invoke-FileHistoryComparison -Side $HistorySide -LogAction 'CompareScanHistory' }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBtFI66055UB17c
# 3NFZu9u+D0+i/WSrdI7vVA8TsvsbyaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII6eTMHn3Hv+2h/LEuN0QV4jMHGFWvxG1dc52AC15k0UMA0GCSqG
# SIb3DQEBAQUABIIBgFoC1RwAgydMA0zGa7mFG5FST0rsLaRlpPq6D8Tox3FtNawe
# GeQvlJxi2QmY2QTq1hPGGcmJFJv4aryPwPr+HULiqVPEc5DozcnEDvXfpsn6CPgz
# N4tqpB7o2MglUvx144ryZyTbNnkCye4Fb9DmzjZSXdM3rUHqMt9PFMlVkA9Y30uu
# DRjTX+fUQ4sBIvAapBjSabgB1I/Z1rfIdxjfSne+01PbIGrsoi8kwIiiZSMgD84j
# H6JJCu27xvMF3Aur5hb21QnVpLDLym4MK1/Rl3mAyk35moFByuR85O8ta54CtDnH
# Yj/hknFKodZE3/ryBj0mtnefNnxIRAHlWuJAIsK9kVegj/PM8JYfslqHxhtQvaLo
# TpJXqbXwCqiJd8zFvEAeIf5QML2qH07gpRWfUDvnz79GerSphqIo/z4mZ7WsA+wD
# WJy5YF2F2Ngl/E+7ZDI5H/0vzqYBNl+nkB/9Zye4wejfhnUEuhGyYwT1f6oyGSr4
# /EanUrz3bRdFirYndKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDgxNjA2
# MTVaMC8GCSqGSIb3DQEJBDEiBCDWVxebJ9Q3reEMaEhvNsuALhqxNJw/BK3x4oGS
# wRn21DANBgkqhkiG9w0BAQEFAASCAgBSqFhy9cFQMc7qKsGZ3z82oVyK6oVwraXV
# N5yL/jKzcbOS1lkTlxa6rbKGF5xfy+e7UoV1TpS2rPjvNBeoCz4HJmTBnFY2eqa1
# /mRcZbYGfhMLBLQWI2Si2NCF8U+X0zzM93f2tmWjWZOeR1HhLisUHJ+Ti3CcLBha
# 7YsHfGb8K/9IVAIi0nMDGncJ8X5T3D2cXAMeMtumjAzaHKwbn+diOfT+sb0rNnOA
# oHjFk2cNb9dpuE7f80+P9fAlyKHfqhwZRd+iIfkf0QmA1mxqqfbQk/tM/ik7ntyq
# sahDY+HuuFkokLKDJCGuNEoqVswxj0h519jPzVAGr8DXlhj+G14cvMIb+x3A3Cf0
# f8jXpO0v9qqBUCR3/fhxCxPSEFTmH56LuMN43INWwrO/isT9J8yyC5Yh+BCZdnbZ
# xt3WvIaXykV9R6XZCtOA5hPMHhYMy3tnMx4W6pYCYxx3n97IvJxvEtd0c6IH2LKZ
# HTo0DvD+I/35pdg2mhYi9Nx/aq8MfO+mtwFk5cx6PRYTerZ2VmN4NJSwVKiF6Dx3
# bffwDFB59vcfObJ7XhszFIfD13CAMlRotmTfK8MeqngGYDk4NRuaYMV5dvZxwvcH
# cW26Pgsgomw85iCmB62vWNPEBm2aI3HEWtQDORYJq0813zokz1KWdqpkZ2hQ9//w
# 111M7T2/RA==
# SIG # End signature block
