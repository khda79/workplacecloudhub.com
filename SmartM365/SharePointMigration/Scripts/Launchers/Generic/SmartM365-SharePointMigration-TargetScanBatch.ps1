<#
.SYNOPSIS
    Run destination inventories for configured SharePoint migrations.

.DESCRIPTION
    Runs target file scans first, then target permission scans. Both
    authentication modes run up to two scans at once. Interactive scans
    use separate visible consoles for sign-in. The limit applies to this batch.

.VERSION
    1.0.3

.EXAMPLE
    pwsh -File .\SmartM365-SharePointMigration-TargetScanBatch.ps1 -PlanOnly

.EXAMPLE
    pwsh -File .\SmartM365-SharePointMigration-TargetScanBatch.ps1 -InventoryMode Both

.EXAMPLE
    pwsh -File .\SmartM365-SharePointMigration-TargetScanBatch.ps1 -AuthMode Certificate -MaxParallel 2
#>

[CmdletBinding()]
param(
    [string[]]$MigrationNames = @(),
    [ValidateSet('FilesOnly', 'PermissionsOnly', 'Both')]
    [string]$InventoryMode = 'Both',
    [ValidateSet('Interactive', 'Certificate')]
    [string]$AuthMode = 'Interactive',
    [ValidateRange(0, 2)]
    [int]$MaxParallel = 0,
    [ValidateRange(0, 300)]
    [int]$LaunchDelaySeconds = 15,
    [switch]$PlanOnly,
    [ValidatePattern('^\d{8}-\d{6}-[a-f0-9]{8}$')]
    [string]$BatchId = '',
    [string]$ProjectRoot = (Join-Path $PSScriptRoot '..\..\..'),
    [string]$LauncherPath = (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-Launcher.ps1')
)

$ErrorActionPreference = 'Stop'
$script:BatchLogPath = $null

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

function Get-BatchMigrations {
    param([string]$Root, [string[]]$Names)

    $migrationsRoot = Join-Path $Root 'Migrations'
    if (-not (Test-Path -LiteralPath $migrationsRoot -PathType Container)) {
        throw "Migrations directory not found: $migrationsRoot"
    }

    $available = @{}
    foreach ($directory in (Get-ChildItem -LiteralPath $migrationsRoot -Directory)) {
        $configPath = Join-Path $directory.FullName 'migration.config.psd1'
        if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { continue }
        $config = Import-PowerShellDataFile -LiteralPath $configPath
        if ($config.Name -ne $directory.Name) {
            throw "Migration name differs from folder name in $configPath"
        }
        if ($directory.Name -eq 'NewMigration') { continue }
        if (-not $config.Target -or $config.Target.Type -ne 'SPO' -or
            [string]::IsNullOrWhiteSpace([string]$config.Target.SiteUrl)) {
            if ($Names -contains $directory.Name) {
                throw "Migration $($directory.Name) needs Target.Type=SPO and Target.SiteUrl."
            }
            continue
        }
        $available[$directory.Name] = [pscustomobject]@{
            Name = $directory.Name
            TargetUrl = [string]$config.Target.SiteUrl
        }
    }

    if ($Names.Count -gt 0) {
        $selected = [System.Collections.Generic.List[object]]::new()
        foreach ($name in $Names) {
            if (-not $available.ContainsKey($name)) { throw "Unknown or unsupported migration: $name" }
            if (@($selected | Where-Object Name -EQ $name).Count -eq 0) {
                $selected.Add($available[$name])
            }
        }
        return @($selected.ToArray())
    }
    return @($available.Values | Sort-Object Name)
}

function Assert-CertificateAuth {
    param([string]$Root)
    $configPath = Join-Path $Root 'Config\SPOAuth.local.psd1'
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        throw "Certificate auth requires $configPath"
    }
    $auth = Import-PowerShellDataFile -LiteralPath $configPath
    foreach ($key in @('ClientId', 'TenantId', 'Thumbprint')) {
        if ([string]::IsNullOrWhiteSpace([string]$auth[$key])) {
            throw "Certificate auth requires $key in $configPath"
        }
    }
    $thumbprint = ([string]$auth.Thumbprint).Replace(' ', '')
    $certificate = @(Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $thumbprint -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
        Select-Object -First 1)
    if ($certificate.Count -eq 0) {
        throw 'Certificate auth requires the configured, unexpired certificate with its private key on this machine.'
    }
}

function Add-BatchResult {
    param([object]$Job, [string]$Status, [int]$ExitCode, [string]$OutputPath = '')
    $runLog = $outputCsv = ''
    try {
        if (Test-Path -LiteralPath $Job.ResultPath -PathType Leaf) {
            $receipt = Get-Content -LiteralPath $Job.ResultPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($receipt.RunId -ne $Job.RunId -or $receipt.Migration -ne $Job.Migration -or $receipt.Action -ne $Job.Action) {
                throw 'The launcher receipt does not match this batch action.'
            }
            $runLog = [string]$receipt.LogPath
            $outputCsv = [string]$receipt.OutputCsv
        }
        else { Write-BatchLine ("Run receipt unavailable for {0} {1}; no output path was inferred." -f $Job.Migration, $Job.Action) }
    }
    catch { Write-BatchLine ("Could not read run receipt for {0} {1}: {2}" -f $Job.Migration, $Job.Action, $_.Exception.Message) }
    $script:Results.Add([pscustomobject]@{
        Migration = $Job.Migration
        Action = $Job.Action
        Status = $Status
        ExitCode = $ExitCode
        Started = $Job.Started.ToString('o')
        Finished = (Get-Date).ToString('o')
        OutputLog = $(if ($runLog) { $runLog } else { $OutputPath })
        ConsoleLog = $OutputPath
        RunLog = $runLog
        OutputCsv = $outputCsv
    })
    Write-BatchLine ("{0}: {1} {2} (exit {3})" -f $Status, $Job.Migration, $Job.Action, $ExitCode)
}

$script:Results = [System.Collections.Generic.List[object]]::new()
$script:ActiveProcesses = [System.Collections.Generic.List[object]]::new()
$requestedNames = [System.Collections.Generic.List[string]]::new()
foreach ($value in $MigrationNames) {
    foreach ($name in ($value -split ',')) {
        if (-not [string]::IsNullOrWhiteSpace($name)) { $requestedNames.Add($name.Trim()) }
    }
}
$ProjectRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
if (-not $PlanOnly) {
    if (-not $BatchId) { $BatchId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)) }
    $batchRoot = Join-Path $ProjectRoot "Migrations\logs\target-scan-batches\$batchId"
    Write-BatchLine ("Creating destination batch log: {0}" -f (Join-Path $batchRoot 'batch.log'))
    New-Item -ItemType Directory -Path $batchRoot -ErrorAction Stop | Out-Null
    $script:BatchLogPath = Join-Path $batchRoot 'batch.log'
    Write-BatchLine ("Destination batch preflight started. Logs: {0}" -f $batchRoot)
}
try {
Write-BatchLine 'Checking PowerShell, launcher, and authentication settings.'
$pwshCommand = Get-Command pwsh -ErrorAction Stop
$LauncherPath = [System.IO.Path]::GetFullPath($LauncherPath)
if (-not (Test-Path -LiteralPath $LauncherPath -PathType Leaf)) {
    throw "Launcher not found: $LauncherPath"
}
$limit = if ($MaxParallel -gt 0) { $MaxParallel } else { 2 }

Write-BatchLine ("Loading destination migration configurations from: {0}" -f (Join-Path $ProjectRoot 'Migrations'))
$migrations = @(Get-BatchMigrations -Root $ProjectRoot -Names $requestedNames.ToArray())
if ($migrations.Count -eq 0) { throw 'No configured SPO migrations found.' }
$actions = @()
if ($InventoryMode -in @('FilesOnly', 'Both')) { $actions += 'ScanTargetFiles' }
if ($InventoryMode -in @('PermissionsOnly', 'Both')) { $actions += 'ScanTargetPermissions' }

Write-BatchLine ("Destination scan plan: {0} migrations; {1} actions; authentication={2}; maximum parallel={3}." -f $migrations.Count, $actions.Count, $AuthMode, $limit)
foreach ($action in $actions) {
    foreach ($migration in $migrations) {
        Write-BatchLine ("PLAN {0} {1} {2}" -f $action, $migration.Name, $migration.TargetUrl)
    }
}
if ($PlanOnly) {
    Write-BatchLine 'Plan only: no scan started and no output created.'
    return
}

Write-BatchLine 'Checking destination authentication and PnP.PowerShell.'
if ($AuthMode -eq 'Certificate') { Assert-CertificateAuth -Root $ProjectRoot }
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'PnP.PowerShell is required in PowerShell 7 before destination scans can run.'
}
}
catch {
    Write-BatchLine ("Destination batch preflight failed: {0}" -f $_.Exception.Message)
    throw
}

Write-BatchLine ("Batch started. Logs: {0}" -f $batchRoot)
$interrupted = $true
try {
    foreach ($action in $actions) {
        Write-BatchLine ("Starting phase: {0}" -f $action)
        $pending = [System.Collections.Generic.Queue[object]]::new()
        foreach ($migration in $migrations) { $pending.Enqueue($migration) }
        $running = [System.Collections.Generic.List[object]]::new()
        $lastLaunch = [datetime]::MinValue
        while ($pending.Count -gt 0 -or $running.Count -gt 0) {
            while ($pending.Count -gt 0 -and $running.Count -lt $limit -and
                   ((Get-Date) - $lastLaunch).TotalSeconds -ge $LaunchDelaySeconds) {
                $migration = $pending.Dequeue()
                $runId = [guid]::NewGuid().ToString('N')
                $job = [pscustomobject]@{ Migration = $migration.Name; Action = $action; Started = Get-Date; RunId=$runId; ResultPath=(Join-Path $batchRoot "$runId.result.json.txt") }
                $prefix = '{0}-{1}' -f $migration.Name, $action
                $stdout = Join-Path $batchRoot "$prefix.stdout.log"
                $stderr = Join-Path $batchRoot "$prefix.stderr.log"
                $arguments = @('-NoLogo', '-NoProfile')
                if ($AuthMode -eq 'Certificate') { $arguments += '-NonInteractive' }
                $arguments += @('-File', ('"{0}"' -f $LauncherPath),
                    '-MigrationName', ('"{0}"' -f $migration.Name), '-Action', $action,
                    '-RunResultPath', ('"{0}"' -f $job.ResultPath), '-RunResultId', $job.RunId)
                if ($AuthMode -eq 'Certificate') { $arguments += '-UseCertificate' }
                $outputLog = if ($AuthMode -eq 'Certificate') { $stdout } else { '' }
                try {
                    $startParameters = @{
                        FilePath = $pwshCommand.Source
                        ArgumentList = $arguments
                        WorkingDirectory = $env:TEMP
                        PassThru = $true
                    }
                    if ($AuthMode -eq 'Certificate') {
                        $startParameters.WindowStyle = 'Hidden'
                        $startParameters.RedirectStandardOutput = $stdout
                        $startParameters.RedirectStandardError = $stderr
                    }
                    else {
                        $startParameters.WindowStyle = 'Normal'
                    }
                    $process = Start-Process @startParameters
                    $running.Add([pscustomobject]@{ Job = $job; Process = $process; Output = $outputLog })
                    $script:ActiveProcesses.Add([pscustomobject]@{ Job = $job; Process = $process; Output = $outputLog })
                    $lastLaunch = Get-Date
                    Write-BatchLine ("START {0} {1}; PID {2}; active {3}/{4}" -f $job.Migration, $action, $process.Id, $running.Count, $limit)
                    if ($AuthMode -eq 'Interactive') {
                        Write-BatchLine ("Sign-in for {0}: complete the prompt in its scan window if requested." -f $job.Migration)
                    }
                }
                catch {
                    Write-BatchLine ("FAILED to start {0} {1}: {2}" -f $job.Migration, $action, $_.Exception.Message)
                    Add-BatchResult -Job $job -Status 'FAILED' -ExitCode 1 -OutputPath $outputLog
                }
            }

            foreach ($item in @($running.ToArray())) {
                $item.Process.Refresh()
                if (-not $item.Process.HasExited) { continue }
                $code = $item.Process.ExitCode
                Add-BatchResult -Job $item.Job -Status $(if ($code -eq 0) { 'SUCCESS' } else { 'FAILED' }) -ExitCode $code -OutputPath $item.Output
                [void]$running.Remove($item)
                foreach ($active in @($script:ActiveProcesses.ToArray())) {
                    if ($active.Process.Id -eq $item.Process.Id) { [void]$script:ActiveProcesses.Remove($active) }
                }
                $item.Process.Dispose()
            }
            if ($pending.Count -gt 0 -or $running.Count -gt 0) {
                Start-Sleep -Seconds 5
            }
        }
    }
    $interrupted = $false
}
finally {
    if ($interrupted) {
        foreach ($active in @($script:ActiveProcesses.ToArray())) {
            try {
                if (-not $active.Process.HasExited) {
                    Stop-Process -Id $active.Process.Id -Force -ErrorAction Stop
                }
            }
            catch {
                Write-BatchLine ("Could not stop PID {0}: {1}" -f $active.Process.Id, $_.Exception.Message)
            }
            Add-BatchResult -Job $active.Job -Status 'INTERRUPTED' -ExitCode 1 -OutputPath $active.Output
            $active.Process.Dispose()
        }
    }
    $summaryPath = Join-Path $batchRoot 'summary.csv'
    $script:Results.ToArray() | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8
    $failed = @($script:Results | Where-Object Status -NE 'SUCCESS').Count
    Write-BatchLine ("Batch finished: {0} successful, {1} failed. Summary: {2}" -f ($script:Results.Count - $failed), $failed, $summaryPath)
}
if ($interrupted -or @($script:Results | Where-Object Status -NE 'SUCCESS').Count -gt 0) { exit 1 }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJ5B9jj3OdcEHb
# 4HwjvvFVEhJW2gSyxPQdN1RVDfvvBaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDrRcCWq5ENeG5Gy9Nn8L/C
# Q87lt+/Jr/n9TpORO5KdcjANBgkqhkiG9w0BAQEFAASCAYAc5sc+ZzS7R3pnBmCV
# 41PIj17cj/tIP4bBsINnNDG/ICbY8KqSF1hN8+0ljCDtcrp5dnWeBe65VrfXtEkz
# DN19icsJFt2OT5g4l08p9BnPhvakZbQ3j2iggH4DzH1/HhCrVpmzGPRhnWGtEvV8
# z2/Og4ZsM/WUxwoU8uKWzu+LarDYCSTEEwdJbPDcA62UgWh1BZjI+b6Bw3VSNAEH
# 1QEgqRHL7lTHsoomBkQ+rVURg5z8Nx8csicDeq9xtiqRFekGmEYcmveRRBgvP1X1
# cZ16uojqM1lXBvs4m/znFp9LKxaZSc2xOHPGdOAr2zpgWvS77o0wMBKlkMqGTJ8S
# ITGhdPuxfkpvJdWhvUbVWv7V/N5IKaMqvbZMWoVwP2tpuy7W0FoCw2v+JF2MPefS
# ChSLqXLt5zBLDPBxhje1dFx/iU+28RqzKQnblvjJXALadxpoOOaHvXXIcdpn3OAk
# yfN9jHZQzogKseTQgjy+g+ex4eQrmgfMWxmmffHUwsVY6NY=
# SIG # End signature block
