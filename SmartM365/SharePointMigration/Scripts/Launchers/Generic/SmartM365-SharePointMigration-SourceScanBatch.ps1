<#
.SYNOPSIS
    Run source SharePoint Server inventories for configured migrations.

.DESCRIPTION
    Run on a SharePoint farm server under an account with SharePoint Shell
    access. Copies the signed inventory scripts and their lifecycle helper to
    a unique local folder, then runs one Windows PowerShell 5.1 scan at a time.
    Source URLs come from Source.UrlsFile or migration.mapping.txt; an absent
    filter is an error so a batch cannot accidentally scan an entire farm.

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
    $migrationsRoot = Join-Path $Root 'Migrations'
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
$LocalRunRoot = [System.IO.Path]::GetFullPath($LocalRunRoot)
if ($LocalRunRoot.StartsWith('\\')) { throw 'LocalRunRoot must be on a local disk of the farm server.' }
if (-not $PlanOnly) {
    $batchId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
    $batchRoot = Join-Path $ProjectRoot "Migrations\logs\source-scan-batches\$batchId"
    Write-BatchLine ("Creating source batch log: {0}" -f (Join-Path $batchRoot 'batch.log'))
    New-Item -ItemType Directory -Path $batchRoot -Force -ErrorAction Stop | Out-Null
    $script:BatchLogPath = Join-Path $batchRoot 'batch.log'
    Write-BatchLine ("Source batch preflight started. Logs: {0}" -f $batchRoot)
}
try {
Write-BatchLine ("Loading source migration configurations from: {0}" -f (Join-Path $ProjectRoot 'Migrations'))
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
    Copy-SourceScript -RelativePath 'Scripts\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1' -Destination (Join-Path $localLaunchers 'SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
    $manifestScript = Join-Path $localCompare 'scan_evidence.py'
    Copy-VerifiedDependency -RelativePath 'Scripts\Compare\scan_evidence.py' -Destination $manifestScript
    Copy-VerifiedDependency -RelativePath 'Scripts\console_lifecycle.py' -Destination (Join-Path $localRoot 'console_lifecycle.py')
    $python = Get-PythonCommand -LocalRoot $localRoot
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
        }
    }
}
catch {
    $script:BatchAborted = $true
    Write-BatchLine ("Batch aborted: {0}" -f $_.Exception.Message)
    throw
}
finally {
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC1OQKIM8uF2zdh
# eB9Dcx6LLwJEREje2BESEqXrkj6f6aCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDqMxRJGwAG8kxd6ITceluc
# B3z3OPbh8K34MaWO1QROBjANBgkqhkiG9w0BAQEFAASCAYCWYDBLZGRFldq4a+d7
# +wH+NIHZ0cIRTY4t+x/rpE0kuGn8SNJaWaOkhaGOgstebEXIKI95H9b/2W8H4zqJ
# hE75MiwpasGmFT1rqUQd12vdv1O+41fxko0ykjLyh8QVOxOso2MegRYcpdNYMSNg
# xGAApHv5OIiJ88kh18KGtVP5G3PPWrtcUg+vG1cjK85X4C1s0D0zcd5vjfZewTH4
# /sc0qZ0pVUJiSKQFkLHSP+ZWyvX2Y40uEH+vrvnq5zeV2/uXXRR7M4irTnZ73YLe
# Z2fThqpfa0TtsIW7DWIkecoWCLSts2Y9of5S871miibpDVnrnzF9XVaodn4DcGSL
# qo8/jgRr8SuBAI8BYk+x1OEDWge40BxTikqr2CdEyGUWXwbxEC0Q67SE/bYqBVeR
# pPHjR2gTLmignGQssat5zZd3z0yHbQEH0s6OvUKUHXk44YJqSYePXIl1nQUtmZ23
# Dx5K0yyqlw2fC3u4DBNNTPSLEXU24TUqItilNxSBJ/Dj3BA=
# SIG # End signature block
