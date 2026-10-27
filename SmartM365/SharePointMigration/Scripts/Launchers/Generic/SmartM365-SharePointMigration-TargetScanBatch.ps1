<#
.SYNOPSIS
    Run destination inventories for configured SharePoint migrations.

.DESCRIPTION
    Runs target file scans first, then target permission scans. Both
    authentication modes run up to two scans at once. Interactive scans
    use separate visible consoles for sign-in. The limit applies to this batch.

.VERSION
    1.0.4

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

    $migrationsRoot = Get-SmartM365MigrationsRoot -ProjectRoot $Root
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
. (Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-Storage.ps1')
if (-not $PlanOnly) {
    if (-not $BatchId) { $BatchId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)) }
    $batchRoot = Get-SmartM365MigrationLogsRoot -ProjectRoot $ProjectRoot -ChildPath "target-scan-batches\$batchId"
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

Write-BatchLine ("Loading destination migration configurations from: {0}" -f (Get-SmartM365MigrationsRoot -ProjectRoot $ProjectRoot))
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDqmwN4UmpaFuHj
# hykEnxteIR7cIIE9rnV5ZoX9AvERZ6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJ2lmPx5Ei7vQ89e3Hpl/2tm0O+MwVQI74Ps8AwOosc3MA0GCSqG
# SIb3DQEBAQUABIIBgHLaC6JS0eXNSn5FCZ/YkiTr1HgunYhaB9CEWEx/4iXrqrqI
# jd3l0ypaguNvj+GVRwyPfCZOsMigPKBDDsyxo955HmEIeSM+s8NQsM7P31BwfPXE
# Ks04NB2Q7hKK+tzCESKl1ykZVgOLtdc1qXN8FGk3qX7sFIrN1IAD7V+WDiQhlgD6
# UVa4ZaM2Ls1PYlVXlAMrrQJGPeIsNxCw2duQi2ZtlnCW2OGXwoN0nOO7xFYABIEN
# /2BJ6yAE/xzYWuftenIXwDhDThxqqkIipE+iQ0I/PmYKLsOyKUOg65kjA3hAczmY
# KkBWvRvlg5NLZFq/2fydileoEtuPLlszDdqwTK5zaUYADhIHbkT5WylNqAEQFN6X
# R8dOQMxAM9vXzEv2jei+hDeRn/dUlSGUir171w/No8XCK1ZM0CWL7m+fyJQOEe0Q
# 8zWYvVG92YokyLVLR9E/vbTC47OuUz85oXb6GpfA8wX9cNdXwrCikf4TdHoY2EzP
# H9OaDNve2Zautg3H86GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxNzQy
# MzVaMC8GCSqGSIb3DQEJBDEiBCC/KP08W2wDXUl6r9CbdanRqqanfhHk3XTk/BFG
# HMAg+TANBgkqhkiG9w0BAQEFAASCAgCNBGdpCDgY4B7dbQdwuOEvLxktsS74fY2H
# W8oDTt5tT8wVJQLwSF4Oh/mNLwmrQLrKn07VfpY+nP0P+969Bt4qzpbxRKC9LiiE
# BQtL1LRoqzsIVcxy4tkPUX7IFOTCAl/HdNAC92iaMeB31kIILmXm596c2Gpv1H2W
# cTvbRUoz9/t4ZJPfUEVASvZlbl3j/C+JUipc/DfSlLoeFkFS9pwSDQ+gEH/ZIfiV
# H/obHlPub9NjyAwCSWUDZq6nfuxpYNrXDFLlf+PgwByhvp6H3CHho2u0+cIoxWt6
# w33H8uPNkk8w9es8f366MEYSsWix5XU8ktwHzvLYYRB1MK0IslqVfjOLev6NM7Dp
# qL9+DXPlmPnWYINurc7QTIjaOTOzJdBd0tkDKECWJDzZnMpRTc1oNJt3RIeSmjOD
# pgb+4iHCT77u7N1d/5ufnUfcKeFklApvNeo54bh36sCWKhZdcUgrSQjHtmpIzyIn
# fkwrMBzviGmyReKv9QDi9p/C3o25njqGw4w/o4QT893jWWRDCbwAgJU86rS6VHnL
# 3otuwa5kG10+53rulZ0RdWVEc2y3rA4zCNvXfb4XM5j+iUVqUabtDyMFB330g7xS
# OqgIjwZ57jIgtLol2SZs+1wHQQxKzCZA4UWS+4gL4syfZbbrav1RAWacstTXhHg6
# zN59VBE0Nw==
# SIG # End signature block
