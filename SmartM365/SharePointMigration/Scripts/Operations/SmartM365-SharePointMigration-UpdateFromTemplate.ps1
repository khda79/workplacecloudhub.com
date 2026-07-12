<#
.SYNOPSIS
    Updates local migration folders from Migrations\_Template.

.DESCRIPTION
    Synchronizes safe template-owned content into migration folders:
    - copies/updates launcher CMD files from _Template\launchers
    - creates missing folder structure from _Template
    - creates the missing mapping placeholder file only if absent
    - audits migration.config.psd1 without overwriting it

    Migration configuration files are intentionally not overwritten because they
    contain environment-specific source and target values.
.VERSION
    1.0.0
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]]$MigrationName,

    [string]$MigrationsRoot,

    [switch]$RemoveUnknownLaunchers,

    [switch]$NoConfigReport,

    [ValidateNotNullOrEmpty()]
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'

function Resolve-DefaultMigrationsRoot {
    $currentDirectory = if ($PSScriptRoot) {
        Get-Item -LiteralPath $PSScriptRoot
    }
    else {
        Get-Item -LiteralPath (Get-Location).Path
    }

    while ($currentDirectory) {
        $candidate = Join-Path -Path $currentDirectory.FullName -ChildPath 'Migrations'
        $templateCandidate = Join-Path -Path $candidate -ChildPath '_Template'
        if (Test-Path -LiteralPath $templateCandidate -PathType Container) {
            return $candidate
        }

        $currentDirectory = $currentDirectory.Parent
    }

    throw 'Could not resolve the Migrations directory. Use -MigrationsRoot.'
}

if ([string]::IsNullOrWhiteSpace($MigrationsRoot)) {
    $MigrationsRoot = Resolve-DefaultMigrationsRoot
}
else {
    $MigrationsRoot = (Resolve-Path -LiteralPath $MigrationsRoot).ProviderPath
}

$TemplateRoot = Join-Path -Path $MigrationsRoot -ChildPath '_Template'
$TemplateConfigPath = Join-Path -Path $TemplateRoot -ChildPath 'migration.config.psd1'

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

function Stop-TimestampedTranscript {
    param([string]$Path)

    try {
        Stop-Transcript | Out-Null
    }
    catch {
        return
    }

    $script:TranscriptStarted = $false
    Add-TimestampToLogFile -Path $Path
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path -Path (Join-Path -Path $MigrationsRoot -ChildPath 'logs') -ChildPath ("SmartM365-SharePointMigration-UpdateFromTemplate-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
}

$logDirectory = Split-Path -Path $LogPath -Parent
if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

$script:TranscriptStarted = $false
trap {
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
    throw $_
}

Start-Transcript -Path $LogPath -Force -WhatIf:$false | Out-Null
$script:TranscriptStarted = $true
Write-Info ("Run log: {0}" -f $LogPath) Cyan

function Get-HashtableLeafPath {
    param(
        [Parameter(Mandatory = $true)]
        $Value,

        [string]$Prefix = ''
    )

    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in ($Value.Keys | Sort-Object)) {
            $childPrefix = if ([string]::IsNullOrWhiteSpace($Prefix)) { [string]$key } else { "{0}.{1}" -f $Prefix, $key }
            foreach ($path in Get-HashtableLeafPath -Value $Value[$key] -Prefix $childPrefix) {
                $path
            }
        }
        return
    }

    $Prefix
}

function Get-RelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $baseUri = [System.Uri]((Resolve-Path -LiteralPath $BasePath).ProviderPath.TrimEnd('\') + '\')
    $pathUri = [System.Uri]((Resolve-Path -LiteralPath $Path).ProviderPath)
    [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($pathUri).ToString()).Replace('/', '\')
}

function Copy-TemplateFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [switch]$OnlyIfMissing
    )

    $destinationExists = Test-Path -LiteralPath $DestinationPath -PathType Leaf
    if ($OnlyIfMissing -and $destinationExists) {
        return 'SkippedExisting'
    }

    $destinationDirectory = Split-Path -Path $DestinationPath -Parent
    if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($destinationDirectory, 'Create directory')) {
            New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
        }
    }

    $action = if ($destinationExists) { 'Update file' } else { 'Create file' }
    if ($PSCmdlet.ShouldProcess($DestinationPath, $action)) {
        Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
    }

    if ($destinationExists) { 'Updated' } else { 'Created' }
}

function Sync-TemplateDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MigrationRoot
    )

    $created = 0
    $templateKeepFiles = Get-ChildItem -LiteralPath $TemplateRoot -File -Filter '.gitkeep' -Recurse |
        Where-Object { $_.FullName -notlike (Join-Path -Path $TemplateRoot -ChildPath 'launchers*') }

    $templateDirectorySet = New-Object 'System.Collections.Generic.SortedSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($keepFile in $templateKeepFiles) {
        $currentDirectory = $keepFile.Directory
        while ($currentDirectory -and ($currentDirectory.FullName -ne $TemplateRoot)) {
            [void]$templateDirectorySet.Add($currentDirectory.FullName)
            $currentDirectory = $currentDirectory.Parent
        }
    }

    foreach ($directoryPath in $templateDirectorySet) {
        $relativePath = Get-RelativePath -BasePath $TemplateRoot -Path $directoryPath
        $destinationDirectory = Join-Path -Path $MigrationRoot -ChildPath $relativePath
        if (-not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
            if ($PSCmdlet.ShouldProcess($destinationDirectory, 'Create template directory')) {
                New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
            }
            $created++
        }
    }

    $created
}

function Sync-TemplateLaunchers {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MigrationRoot
    )

    $templateLauncherRoot = Join-Path -Path $TemplateRoot -ChildPath 'launchers'
    $destinationLauncherRoot = Join-Path -Path $MigrationRoot -ChildPath 'launchers'
    $updated = 0
    $created = 0
    $removed = 0

    if (-not (Test-Path -LiteralPath $destinationLauncherRoot -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($destinationLauncherRoot, 'Create launchers directory')) {
            New-Item -ItemType Directory -Path $destinationLauncherRoot -Force | Out-Null
        }
    }

    $templateLauncherNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($templateFile in Get-ChildItem -LiteralPath $templateLauncherRoot -File -Filter '*.cmd') {
        [void]$templateLauncherNames.Add($templateFile.Name)
        $destinationPath = Join-Path -Path $destinationLauncherRoot -ChildPath $templateFile.Name
        $result = Copy-TemplateFile -SourcePath $templateFile.FullName -DestinationPath $destinationPath
        if ($result -eq 'Created') { $created++ }
        elseif ($result -eq 'Updated') { $updated++ }
    }

    if ($RemoveUnknownLaunchers) {
        foreach ($existingLauncher in Get-ChildItem -LiteralPath $destinationLauncherRoot -File -Filter '*.cmd' -ErrorAction SilentlyContinue) {
            if (-not $templateLauncherNames.Contains($existingLauncher.Name)) {
                if ($PSCmdlet.ShouldProcess($existingLauncher.FullName, 'Remove launcher not present in template')) {
                    Remove-Item -LiteralPath $existingLauncher.FullName -Force
                }
                $removed++
            }
        }
    }

    [pscustomobject]@{
        Created = $created
        Updated = $updated
        Removed = $removed
    }
}

function Ensure-TemplateConfigFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MigrationRoot
    )

    $created = 0
    foreach ($fileName in @('migration.mapping.txt')) {
        $sourcePath = Join-Path -Path $TemplateRoot -ChildPath $fileName
        $destinationPath = Join-Path -Path $MigrationRoot -ChildPath $fileName
        $result = Copy-TemplateFile -SourcePath $sourcePath -DestinationPath $destinationPath -OnlyIfMissing
        if ($result -eq 'Created') {
            $created++
        }
    }

    $created
}

function Test-MigrationConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MigrationRoot,

        [Parameter(Mandatory = $true)]
        [string]$MigrationName,

        [Parameter(Mandatory = $true)]
        [hashtable]$TemplateConfig
    )

    $configPath = Join-Path -Path $MigrationRoot -ChildPath 'migration.config.psd1'
    $reportPath = Join-Path -Path $MigrationRoot -ChildPath 'migration.config.template-diff.txt'

    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        if (-not $NoConfigReport) {
            $content = @(
                "Migration config is missing.",
                "Create it from _Template\migration.config.psd1 and update environment-specific values."
            )
            if ($PSCmdlet.ShouldProcess($reportPath, 'Write config diff report')) {
                Set-Content -LiteralPath $reportPath -Value $content -Encoding UTF8
            }
        }
        return [pscustomobject]@{
            Status = 'MissingConfig'
            MissingKeys = @()
            ExtraKeys = @()
            ReportPath = $reportPath
        }
    }

    $migrationConfig = Import-PowerShellDataFile -LiteralPath $configPath
    $templateKeys = @(Get-HashtableLeafPath -Value $TemplateConfig)
    $migrationKeys = @(Get-HashtableLeafPath -Value $migrationConfig)

    $migrationKeySet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $migrationKeys) { [void]$migrationKeySet.Add($key) }

    $templateKeySet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $templateKeys) { [void]$templateKeySet.Add($key) }

    $missingKeys = @($templateKeys | Where-Object { -not $migrationKeySet.Contains($_) })
    $extraKeys = @($migrationKeys | Where-Object { -not $templateKeySet.Contains($_) })

    if (-not $NoConfigReport) {
        $content = New-Object 'System.Collections.Generic.List[string]'
        $content.Add(("Migration: {0}" -f $MigrationName))
        $content.Add(("Checked at: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
        $content.Add('')
        $content.Add('Config files are not overwritten automatically.')
        $content.Add('Add missing keys manually if needed, keeping migration-specific values.')
        $content.Add('')
        $content.Add('Missing keys from template:')
        if ($missingKeys.Count) {
            foreach ($key in $missingKeys) { $content.Add(("  - {0}" -f $key)) }
        }
        else {
            $content.Add('  <none>')
        }
        $content.Add('')
        $content.Add('Extra local keys not in template:')
        if ($extraKeys.Count) {
            foreach ($key in $extraKeys) { $content.Add(("  - {0}" -f $key)) }
        }
        else {
            $content.Add('  <none>')
        }

        if ($PSCmdlet.ShouldProcess($reportPath, 'Write config diff report')) {
            Set-Content -LiteralPath $reportPath -Value $content -Encoding UTF8
        }
    }

    [pscustomobject]@{
        Status = if ($missingKeys.Count) { 'MissingKeys' } else { 'OK' }
        MissingKeys = $missingKeys
        ExtraKeys = $extraKeys
        ReportPath = $reportPath
    }
}

if (-not (Test-Path -LiteralPath $TemplateRoot -PathType Container)) {
    throw "Template directory not found: $TemplateRoot"
}

if (-not (Test-Path -LiteralPath $TemplateConfigPath -PathType Leaf)) {
    throw "Template config not found: $TemplateConfigPath"
}

$templateConfig = Import-PowerShellDataFile -LiteralPath $TemplateConfigPath

$migrationDirectories = Get-ChildItem -LiteralPath $MigrationsRoot -Directory |
    Where-Object { $_.Name -ne '_Template' }

if ($MigrationName) {
    $wanted = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $MigrationName) { [void]$wanted.Add($name) }
    $migrationDirectories = @($migrationDirectories | Where-Object { $wanted.Contains($_.Name) })
}

if (-not $migrationDirectories) {
    Write-Info 'No migration folders found to update.' Yellow
    if ($script:TranscriptStarted) {
        Stop-TimestampedTranscript -Path $LogPath
    }
    return
}

Write-Info ("Template: {0}" -f $TemplateRoot) Cyan
Write-Info ("Migrations to update: {0}" -f (($migrationDirectories | Select-Object -ExpandProperty Name) -join ', ')) Cyan

foreach ($migrationDirectory in $migrationDirectories) {
    Write-Info ("Updating migration: {0}" -f $migrationDirectory.Name) Cyan

    $directoryCount = Sync-TemplateDirectories -MigrationRoot $migrationDirectory.FullName
    $launcherResult = Sync-TemplateLaunchers -MigrationRoot $migrationDirectory.FullName
    $configFileCount = Ensure-TemplateConfigFiles -MigrationRoot $migrationDirectory.FullName
    $configResult = Test-MigrationConfig `
        -MigrationRoot $migrationDirectory.FullName `
        -MigrationName $migrationDirectory.Name `
        -TemplateConfig $templateConfig

    Write-Info ("  Directories created: {0}" -f $directoryCount)
    Write-Info ("  Launchers created/updated/removed: {0}/{1}/{2}" -f $launcherResult.Created, $launcherResult.Updated, $launcherResult.Removed)
    Write-Info ("  Missing mapping files created: {0}" -f $configFileCount)
    Write-Info ("  Config status: {0}; missing keys: {1}; report: {2}" -f $configResult.Status, $configResult.MissingKeys.Count, $configResult.ReportPath)
}

if ($script:TranscriptStarted) {
    Stop-TimestampedTranscript -Path $LogPath
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOefkPAh55Wc2E
# n0SOzeqM/pRJxSf3wuBuQDJ6OQxw3aCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBCccQxWg6pAu34MzEX
# otWNbyEDCMNPLVh5KMonNm8GbDANBgkqhkiG9w0BAQEFAASCAYAfcj69+3icZ2qN
# UcrOqtAnJz5b0SF2S94GIEwr3BvHkIZmZ1rmKUwY2mBOk67jELYrXNcYGynLp2Ab
# k2NIpm/hB71/rBoSuDmgykYPIoiXbjPehI5USzOnOma8rYwNfn/q2m8JD3+CqDUU
# ydxT8/0eEiP6s53dacYEI3t8r9aN/9LzOs84OPg9a+bnl6cHsG4nzsgpzceJmxG5
# 2c7lSAJRwYwAdsNpXACdIPDO3+k4VxG0QgLBG5Vh/2QQoBTEBiubH+drLIfPBISv
# t0wL0GYcN8TM2OxmBV0/vUMrN0mjGvRhIKo7ZOlHs0apDjeSpvdsx+kZfwFPDIkP
# ABjTJpEPz7xgljTPzUzMHW14f6gf7UmKiXl+kOPVHp9hhWVhTvZJIaaQ/8zXbLOq
# vDoylLIZlFvWcJmvAiw3heKKWah0hxe/E/7yaHrD0naS2XNsCdbIA6BpTmf8V1Il
# DAOfpjbpBhREJcluJZxNjezhmVcUcWrZ+MxRa+IhEDIxcMl/I6w=
# SIG # End signature block
