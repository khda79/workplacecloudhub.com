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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCvEWoryyauJccD
# 4nqBMgNHv1DGOna20Rlta3opZJynrqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID6t5L31jx2TMkPzulto8Ughf05QvUA/pFAvFnveqy5VMA0GCSqG
# SIb3DQEBAQUABIIBgIv0VQ1YWpPqdW0xejOTMtPDcG1qQRFJwFZLP/N1xoy+bj80
# mWmlBcJnf9UiewxLGJoamy81Opve5GXDJy8SD56iSSuE/xVzrTo8PvIfXOCBlUR4
# H70r5OFf/1hXHjpfj325RvvJCpwrLTfOvZ9BRJTxOo5UBT6FBz0yV/UMpNGpQnXC
# tWxRK4uTMKsASMlTypn839k6jMf2sJuCf8ihuQjreAf0DlUJPZSzHIosoEI7204b
# NihmbmB5ry2bP31FBpi++0kD6A053Te4GLaANnYI9NhSjZYPEIplAfNL18KShbhc
# u2qIVGEdcJAwkmsn4Uxbk+RCO0QbcA0apauAmG8bEPvcA675HkOiR3oKqSpObRS5
# B2s31efdUOhYTVPaW1M88K9xYCSM6mFSeOJKlurvk2YCSU1GxYWZw0IJsFgFq3tK
# cO1+lWC+7+VInckKRt46XWM2ImtiYDPCGkbog9S87skQFi61h3GjESyupRj8G2sn
# H+clgQW1r9A2d3g6IKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzVaMC8GCSqGSIb3DQEJBDEiBCCyNC5gy/pHfR/qF4O5vRT6JU1q90e6fQipPHqk
# NF/i7jANBgkqhkiG9w0BAQEFAASCAgAgaAw4QzPItEdFD2mTctao6mSB6nXmPvmd
# UjPtzDsPGtbyEjDdL07PTKqD9SUJ8Is3lEsA/go5vvoiiCOYkUhkJII+j7+Nuyd/
# Hb6MIhdSJha9lFGcOlQeXvIMf7sEUAjbmy6317UGIEQJ43VQLFQN3kYaWhX6U1x8
# sV9n7vA2SaSCceUdAHIBMRpuQ4XCbT+fbGJcri6q7obklzsK905K8Kekc1+kpZTw
# fO8S/0QmG9qBxdHu+tQzJyotgRVhtsdUavkV4WBaHyB6Bq3F4YyS5/KNFjVOB16e
# ZymqXLi126hjGGPqY7HhrnHP8TsNdmo/TB26fLGAZp+CZm7ko1W2sBCs2W+3nHkS
# 2p0K8YkE5m8ePH6iFLSRljGRe834l+s32fWGxgmV0t05Gayycp1q11sCjouH8m6S
# DzEn2bmMpU3UyaMj3m5jIdEhkpHKkEC326CaeIvUOvAUSIDoBkqpaFTGWyCkzWnG
# vVypWWnwaHWWocy3m7JeDhSIH0RR4lhOKdiUfCNH5DeCLI3q3CmvzGmjW3z3NIBU
# JSaPw8WjwBIhj6WoLH8vKkcfmDKpiLX7Ke/RdWeQkepEyJR5xXc7TUnUqo9NuTet
# BnSkzJSx8GQ5kpAxlkf4WvpwHo2X95G3j5JjILsX9jm4Lkl1TXN1/81cTCESOd2v
# m3XZUGIyyw==
# SIG # End signature block
