<#
.SYNOPSIS
    Shared helpers for SharePoint migration launchers, source scan batches and the GUI.
.VERSION
    1.0.5
#>

. (Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-SharePointMigration-Storage.ps1')

function Get-LauncherProjectRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LauncherScriptRoot
    )

    return (Resolve-Path -LiteralPath (Join-Path -Path $LauncherScriptRoot -ChildPath '..\..\..')).ProviderPath
}

function Import-LauncherMigrationConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [string]$MigrationName,

        [string]$ConfigFileName = 'migration.config.psd1'
    )

    $migrationRoot = Join-Path -Path (Get-SmartM365MigrationsRoot -ProjectRoot $ProjectRoot) -ChildPath $MigrationName
    $configPath = Join-Path -Path $migrationRoot -ChildPath $ConfigFileName
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        throw "Migration config not found: $configPath"
    }

    [pscustomobject]@{
        Root = $migrationRoot
        Path = $configPath
        Data = Import-PowerShellDataFile -LiteralPath $configPath
    }
}

function Resolve-LauncherMigrationPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MigrationRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path -Path $MigrationRoot -ChildPath $Path
}

function Use-DefaultParameterValue {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BoundParameters,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        $Value
    )

    if ($BoundParameters.ContainsKey($Name)) {
        return $null
    }

    return $Value
}

# Inventory ordering uses collection completion, never the copy/synchronization time.
function Get-SmartM365InventoryScans {
    param([string]$Directory, [string]$Filter, [switch]$Recurse)
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return @() }
    @(Get-ChildItem -LiteralPath $Directory -Filter $Filter -File -Recurse:$Recurse -ErrorAction Stop |
        Where-Object {
            $_.Name -notlike '*-Errors.csv' -and
            $_.Name -notlike 'AssessmentDetails-*' -and $_.Name -notlike '*.assessment.*.csv' -and
            -not (Test-Path -LiteralPath (Join-Path $_.DirectoryName ($_.BaseName + '-Errors.csv')) -PathType Leaf)
        } | ForEach-Object {
            $file = $_; $date = $file.LastWriteTime; $rows = $null
            $provenance = 'file timestamp (unverified)'
            if ($file.Name -match '-(?<date>\d{8})-(?<time>\d{6})(?:-[^.]+)?\.csv$') {
                try {
                    $date = [datetime]::ParseExact(($Matches.date + $Matches.time), 'yyyyMMddHHmmss', [Globalization.CultureInfo]::InvariantCulture)
                    $provenance = 'CSV filename (legacy scan)'
                } catch { return }
            }
            $receiptPath = "$($file.FullName).manifest.json.txt"
            if (Test-Path -LiteralPath $receiptPath -PathType Leaf) {
                try {
                    $receipt = Get-Content -LiteralPath $receiptPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    $count = [long]0
                    if ($receipt.SchemaVersion -ne 1 -or $receipt.InventoryFile -ne $file.Name -or
                        [string]$receipt.Sha256 -notmatch '^[a-fA-F0-9]{64}$' -or
                        -not [long]::TryParse([string]$receipt.Rows, [ref]$count) -or $count -lt 0) { return }
                    $date = [datetimeoffset]::Parse([string]$receipt.CompletedAtUtc, [Globalization.CultureInfo]::InvariantCulture).LocalDateTime
                    $rows = $count; $provenance = 'scan receipt (hash not recalculated)'
                } catch { return }
            }
            [pscustomobject]@{ File=$file; Date=$date; Rows=$rows; Provenance=$provenance }
        } | Sort-Object @{Expression='Date'; Descending=$true}, @{Expression={ $_.File.Name }; Descending=$true})
}

# Migration complexity is scored from source scans only. Source scans compute it after
# publishing a scan; the GUI recalculates only when the stored result does not match.
function Get-SmartM365ComplexityEntraCache {
    param([Parameter(Mandatory = $true)]$Migration, [Parameter(Mandatory = $true)][string]$ProjectRoot)
    $comparison = $Migration.Config['Comparison']
    $configured = if ($comparison) { [string]$comparison['EntraUsersCachePath'] } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($configured)) {
        $path = Resolve-LauncherMigrationPath -MigrationRoot $Migration.Root -Path $configured
        if (Test-Path -LiteralPath $path -PathType Leaf) { return Get-Item -LiteralPath $path }
        return $null
    }
    $directory = Join-Path $ProjectRoot '_Local\EntraUsersCache'
    # The GUI asks once per migration on each refresh: reuse the folder listing for a few seconds.
    $memo = Get-Variable -Name SmartM365ComplexityEntraMemo -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($memo -and $memo.Directory -eq $directory -and ((Get-Date) - $memo.At).TotalSeconds -lt 15) { return $memo.Item }
    $item = $null
    if (Test-Path -LiteralPath $directory -PathType Container) {
        $item = Get-ChildItem -LiteralPath $directory -Filter '*.csv' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    }
    $script:SmartM365ComplexityEntraMemo = [pscustomobject]@{ Directory = $directory; At = Get-Date; Item = $item }
    return $item
}

function Get-SmartM365ComplexityFileKey {
    # Second precision: synchronized copies can truncate sub-second timestamps.
    param([System.IO.FileInfo]$File)
    if (-not $File) { return '' }
    $ticks = $File.LastWriteTimeUtc.Ticks
    return '{0}:{1}:{2}' -f $File.Name, $File.Length, ($ticks - ($ticks % [TimeSpan]::TicksPerSecond))
}

function Get-SmartM365ComplexityConfigHash {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    $item = Get-Item -LiteralPath $Path
    $key = '{0}|{1}|{2}' -f $item.FullName, $item.Length, $item.LastWriteTimeUtc.Ticks
    $memo = Get-Variable -Name SmartM365ComplexityConfigHashMemo -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($memo -and $memo.Key -eq $key) { return $memo.Hash }
    $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    $script:SmartM365ComplexityConfigHashMemo = [pscustomobject]@{ Key = $key; Hash = $hash }
    return $hash
}

# A source scan announces a pending complexity calculation so the GUI does not start the
# same calculation meanwhile. Markers older than MaxAgeHours are ignored (interrupted scans).
function Set-SmartM365ComplexityPending {
    param([Parameter(Mandatory = $true)][string]$MigrationRoot, [string]$Reason = 'Source scan')
    $path = Join-Path $MigrationRoot ('migration-complexity.pending.{0}.txt' -f [guid]::NewGuid().ToString('N'))
    try {
        Set-Content -LiteralPath $path -Value ('{0} {1} on {2}' -f (Get-Date).ToUniversalTime().ToString('o'), $Reason, $env:COMPUTERNAME) -Encoding UTF8 -ErrorAction Stop
        return $path
    }
    catch { return '' }
}

function Clear-SmartM365ComplexityPending {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
}

function Get-SmartM365ComplexityPending {
    param([Parameter(Mandatory = $true)][string]$MigrationRoot, [double]$MaxAgeHours = 24)
    if (-not (Test-Path -LiteralPath $MigrationRoot -PathType Container)) { return $null }
    return Get-ChildItem -LiteralPath $MigrationRoot -Filter 'migration-complexity.pending.*.txt' -File -ErrorAction SilentlyContinue |
        Where-Object { ((Get-Date).ToUniversalTime() - $_.LastWriteTimeUtc).TotalHours -lt $MaxAgeHours } |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
}

function Get-SmartM365PythonCommand {
    <# Portable Tools\Python first, then python, then the py -3 launcher; each must start. #>
    param([Parameter(Mandatory = $true)][string]$ProjectRoot, [switch]$LocalOnly)
    $candidates = @([pscustomobject]@{ Executable = (Join-Path $ProjectRoot 'Tools\Python\python.exe'); Arguments = @(); Source = 'Portable' })
    foreach ($name in @('python', 'py')) {
        $command = Get-Command -Name $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $command -or ($LocalOnly -and [string]$command.Source -like '\\*')) { continue }
        $arguments = if ($name -eq 'py') { @('-3') } else { @() }
        $candidates += [pscustomobject]@{ Executable = $command.Source; Arguments = $arguments; Source = $(if ($name -eq 'py') { 'PyLauncher' } else { 'PATH' }) }
    }
    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate.Executable -PathType Leaf)) { continue }
        $previousPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            & $candidate.Executable @($candidate.Arguments) --version *> $null
            $exitCode = $LASTEXITCODE
        }
        finally { $ErrorActionPreference = $previousPreference }
        if ($exitCode -eq 0) { return $candidate }
    }
    throw ('Python 3 is required. Expected portable Python at: {0}' -f (Join-Path $ProjectRoot 'Tools\Python\python.exe'))
}

function Get-SmartM365ComplexityScanInputs {
    param([Parameter(Mandatory = $true)]$Migration, [Parameter(Mandatory = $true)][string]$SourceType)
    $cfg = $Migration.Config
    $result = @{ FileScan = $null; PermissionScan = $null }
    if (-not $cfg['Output']) { return [pscustomobject]$result }
    foreach ($spec in @(
            @{ Key = 'FileScan'; Output = 'SourceFileScans'; Kind = 'File' },
            @{ Key = 'PermissionScan'; Output = 'SourcePermissionScans'; Kind = 'Permission' })) {
        $relative = [string]$cfg['Output'][$spec.Output]
        if ([string]::IsNullOrWhiteSpace($relative)) { continue }
        $directory = Resolve-LauncherMigrationPath -MigrationRoot $Migration.Root -Path $relative
        $filter = '{0}-{1}Inventory-{2}-*.csv' -f $SourceType, $spec.Kind, [string]$cfg['Name']
        $scan = @(Get-SmartM365InventoryScans -Directory $directory -Filter $filter -Recurse) | Select-Object -First 1
        if ($scan) { $result[$spec.Key] = $scan.File }
    }
    return [pscustomobject]$result
}

function Get-SmartM365ComplexityRequest {
    <# Inputs, arguments and signature of the source-only complexity score for one migration. #>
    param(
        [Parameter(Mandatory = $true)]$Migration,
        [Parameter(Mandatory = $true)][string]$ProjectRoot,
        [System.IO.FileInfo]$FileScan,
        [System.IO.FileInfo]$PermissionScan,
        [string]$ScriptPath = ''
    )
    if (-not $FileScan) { return $null }
    $cfg = $Migration.Config
    $entraCache = if ($PermissionScan) { Get-SmartM365ComplexityEntraCache -Migration $Migration -ProjectRoot $ProjectRoot } else { $null }
    $mapping = $null
    $mappingName = if ($cfg['Comparison']) { [string]$cfg['Comparison']['PathMappingsFile'] } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($mappingName)) {
        $mappingPath = Resolve-LauncherMigrationPath -MigrationRoot $Migration.Root -Path $mappingName
        if (Test-Path -LiteralPath $mappingPath -PathType Leaf) { $mapping = Get-Item -LiteralPath $mappingPath }
    }
    $targetSiteUrl = if ($cfg['Target']) { [string]$cfg['Target']['SiteUrl'] } else { '' }
    $configDirectory = Join-Path $ProjectRoot 'Config'
    $configPath = Join-Path $configDirectory 'migration-complexity.json.txt'
    $configHash = Get-SmartM365ComplexityConfigHash -Path $configPath
    $assessment = $null
    $assessmentPath = Join-Path $Migration.Root 'source-complexity-assessment.json.txt'
    if (Test-Path -LiteralPath $assessmentPath -PathType Leaf) { $assessment = Get-Item -LiteralPath $assessmentPath }
    $assessmentHash = if ($assessment) { Get-SmartM365ComplexityConfigHash -Path $assessmentPath } else { '' }
    $automaticHashes = [ordered]@{File='';Permission='';FileDetails='';PermissionDetails=''}
    foreach ($spec in @(@{Key='File';Scan=$FileScan},@{Key='Permission';Scan=$PermissionScan})) {
        if ($spec.Scan) {
            $automaticPath = $spec.Scan.FullName + '.assessment.json.txt'
            if (Test-Path -LiteralPath $automaticPath -PathType Leaf) {
                $automaticHashes[$spec.Key] = Get-SmartM365ComplexityConfigHash -Path $automaticPath
                try {
                    $auto = Get-Content -LiteralPath $automaticPath -Raw -Encoding UTF8 | ConvertFrom-Json
                    $detailName = [string]$auto.DetailsFile
                    if ($detailName -and [IO.Path]::GetFileName($detailName) -ceq $detailName) {
                        $detailPath = Join-Path $spec.Scan.DirectoryName $detailName
                        if (Test-Path -LiteralPath $detailPath -PathType Leaf) { $automaticHashes[$spec.Key+'Details'] = Get-SmartM365ComplexityConfigHash -Path $detailPath }
                    }
                } catch { } # Invalid sidecars are diagnosed by the calculator and remain N/A.
            }
        }
    }
    $assessmentTarget = if ($cfg['Target'] -and $cfg['Target']['Type'] -and [string]$cfg['Target']['Type'] -ne 'SPO') { 'Other' } else { 'SPO' }
    if ([string]::IsNullOrWhiteSpace($ScriptPath)) { $ScriptPath = Join-Path $ProjectRoot 'Scripts\Compare\migration_complexity.py' }
    $output = Join-Path $Migration.Root 'migration-complexity.json.txt'
    $arguments = @(
        $ScriptPath,
        '--migration-name', $Migration.Name,
        '--file-inventory', $FileScan.FullName,
        '--config-dir', $configDirectory,
        '--output', $output
    )
    if ($PermissionScan) { $arguments += @('--permission-inventory', $PermissionScan.FullName) }
    if ($entraCache) { $arguments += @('--entra-users-csv', $entraCache.FullName) }
    if ($mapping) { $arguments += @('--path-mapping-file', $mapping.FullName) }
    if ($targetSiteUrl) { $arguments += @('--target-site-url', $targetSiteUrl) }
    if ($assessment) { $arguments += @('--source-assessment', $assessment.FullName) }
    $arguments += @('--assessment-target-platform', $assessmentTarget)
    # Scans have unique timestamped names; the Entra cache and the mapping keep their name when rewritten.
    $inputKeys = foreach ($item in @($FileScan, $PermissionScan, $entraCache, $mapping)) { Get-SmartM365ComplexityFileKey -File $item }
    $signature = (@($Migration.Name, $configHash, $targetSiteUrl, $assessmentHash, $automaticHashes.File, $automaticHashes.Permission, $automaticHashes.FileDetails, $automaticHashes.PermissionDetails, $assessmentTarget) + @($inputKeys)) -join '|'
    [pscustomobject]@{
        Migration = $Migration.Name; Output = $output; Arguments = $arguments; Signature = $signature
        FileScan = $FileScan; PermissionScan = $PermissionScan; EntraCache = $entraCache; Mapping = $mapping
        TargetSiteUrl = $targetSiteUrl; ConfigHash = $configHash
        Assessment = $assessment; AssessmentHash = $assessmentHash
        AutomaticAssessmentHashes = [pscustomobject]$automaticHashes
    }
}

function Invoke-SmartM365MigrationComplexity {
    <# Scores the latest source scans after a scan. Never throws: a failure must not fail the scan. #>
    param(
        [Parameter(Mandatory = $true)]$Migration,
        [Parameter(Mandatory = $true)][string]$ProjectRoot,
        [Parameter(Mandatory = $true)][string]$SourceType,
        [Parameter(Mandatory = $true)][string]$PythonExecutable,
        [string[]]$PythonArguments = @(),
        [string]$ScriptPath = ''
    )
    try {
        $inputs = Get-SmartM365ComplexityScanInputs -Migration $Migration -SourceType $SourceType
        $request = Get-SmartM365ComplexityRequest -Migration $Migration -ProjectRoot $ProjectRoot `
            -FileScan $inputs.FileScan -PermissionScan $inputs.PermissionScan -ScriptPath $ScriptPath
        if (-not $request) {
            return [pscustomobject]@{ Status = 'Skipped'; Message = 'No completed source file scan is available.'; Output = @() }
        }
        $previousPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'Continue'
            $output = @(& $PythonExecutable @PythonArguments @($request.Arguments) 2>&1 | ForEach-Object { [string]$_ })
            $exitCode = $LASTEXITCODE
        }
        finally { $ErrorActionPreference = $previousPreference }
        if ($exitCode -ne 0) {
            return [pscustomobject]@{ Status = 'Failed'; Message = "Python exit code $exitCode."; Output = $output }
        }
        return [pscustomobject]@{ Status = 'Succeeded'; Message = $request.Output; Output = $output }
    }
    catch {
        return [pscustomobject]@{ Status = 'Failed'; Message = $_.Exception.Message; Output = @() }
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCtlJ1gawh9sKcz
# 8CAOdScZiy2oRWJPpQuy9pK/bkSTTqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBiBJT4y8vqzD1bNtIw1BEz
# EZg5hlw65XWurtj1+q/w3zANBgkqhkiG9w0BAQEFAASCAYBrfAnJDwY7usMA5jy4
# MrMnKsvzPfHPXdX9ySSS8Qp4kkCXjoJ83yyLLsLzsVJYUglUp6O8YvIA/qUmTQJY
# aJcuhR6k54okvrwoWzpKVSWUPJRJu0KQ9xvJA+7ObOzyVyZpNSz504EE6g8+4jl5
# 5RyjZGDsNZj/PJmVBNWNhWB3j29l2/YXIgxF9GMUb9AL0OXSktIzClkEcOdhswyo
# HyD/yOyJ6E59FTJcPLk37WJGNHEIBix5nYuOjq1obyTGFnfpPmBBoNQR26vuvhNe
# Zq9hDfSWUclgHtLkYmhhl7N/8bmbbWu9MlFpxgIv1H9eldEjAVXg21wb/9cCVPfj
# 4jefiAXYPKvOqiSKTmvpba+UYgPh+Sc9ZBzIQ6u9QBCcQ6BilyMuZQ2wQqp5x3LS
# fk1jJhDRC7KhcvWo4fwCquUjoGn2vEGndsKpSNR5VLgnf3uJ6RSvP4r0mk3nA/Jd
# zbzRCeuHFZz+gDKEfrIwXneXHbqKd0cCVWhZBndYUschYQ8=
# SIG # End signature block
