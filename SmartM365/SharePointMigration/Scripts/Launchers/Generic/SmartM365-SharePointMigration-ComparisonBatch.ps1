<#
.SYNOPSIS
    Compare source and destination inventories for all configured migrations.
.DESCRIPTION
    Runs file comparisons first, then permission comparisons, sequentially.
    Uses the existing migration launcher to select the latest inventories,
    validate scan evidence, refresh the Entra cache, and generate reports.
    Continues after individual failures and returns a nonzero batch exit code.
.VERSION
    1.0.1
.EXAMPLE
    .\SmartM365-SharePointMigration-ComparisonBatch.ps1 -PlanOnly
.EXAMPLE
    .\SmartM365-SharePointMigration-ComparisonBatch.ps1 -ComparisonMode Both
#>
#requires -Version 7.4
[CmdletBinding()]
param(
    [string[]]$MigrationNames = @(),
    [ValidateSet('FilesOnly', 'PermissionsOnly', 'Both')]
    [string]$ComparisonMode = 'Both',
    [ValidateSet('Interactive', 'Certificate')]
    [string]$AuthMode = 'Interactive',
    [switch]$PlanOnly,
    [ValidatePattern('^\d{8}-\d{6}-[a-f0-9]{8}$')]
    [string]$BatchId = '',
    [switch]$Force,
    [string]$ProjectRoot = (Join-Path $PSScriptRoot '..\..\..'),
    [string]$LauncherPath = ''
)

$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$script:BatchLogPath = $null
$script:Results = [Collections.Generic.List[object]]::new()

function Write-Host {
    param([Parameter(Position = 0, ValueFromRemainingArguments = $true)][object[]]$Object)
    foreach ($part in $Object) {
        foreach ($line in ([string]$part -split "`r?`n")) {
            Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line)
        }
    }
}

function Write-BatchLine {
    param([string]$Message, [string]$ConsoleLog = '')
    foreach ($line in ($Message -split "`r?`n")) {
        $entry = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line
        Microsoft.PowerShell.Utility\Write-Host $entry
        if ($script:BatchLogPath) { Add-Content -LiteralPath $script:BatchLogPath -Value $entry -Encoding utf8 }
        if ($ConsoleLog) { Add-Content -LiteralPath $ConsoleLog -Value $entry -Encoding utf8 }
    }
}

function Get-ComparisonMigrations {
    param([string]$Root, [string[]]$Names)
    $directory = Join-Path $Root 'Migrations'
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { throw "Migrations directory not found: $directory" }
    $available = @{}
    foreach ($item in (Get-ChildItem -LiteralPath $directory -Directory)) {
        if ($item.Name -eq 'NewMigration') { continue }
        $path = Join-Path $item.FullName 'migration.config.psd1'
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $configError = ''
        try {
            $config = Import-PowerShellDataFile -LiteralPath $path
            if ($config.Name -ne $item.Name) { throw "Migration name differs from folder name: $path" }
        }
        catch { $configError = $_.Exception.Message }
        $available[$item.Name] = [pscustomobject]@{ Name=$item.Name; ConfigError=$configError }
    }
    if ($Names.Count) {
        $seen = @{}
        foreach ($name in $Names) {
            if (-not $available.ContainsKey($name)) { throw "Unknown migration: $name" }
            if ($seen.ContainsKey($name)) { continue }
            $seen[$name] = $true
            $available[$name]
        }
    }
    else { $available.Values | Sort-Object Name }
}

$ProjectRoot = [IO.Path]::GetFullPath($ProjectRoot)
if ([string]::IsNullOrWhiteSpace($LauncherPath)) {
    $LauncherPath = Join-Path $ProjectRoot 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1'
}
$requested = @($MigrationNames | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$actions = @()
if ($ComparisonMode -in @('FilesOnly','Both')) { $actions += 'CompareFiles' }
if ($ComparisonMode -in @('PermissionsOnly','Both')) { $actions += 'ComparePermissions' }
$summaryPath = $null
$interrupted = $true
try {
    if (-not $PlanOnly) {
        $id = if ($BatchId) { $BatchId } else { '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N').Substring(0,8) }
        $batchRoot = Join-Path $ProjectRoot "Migrations\logs\comparison-batches\$id"
        Write-BatchLine ("Creating comparison batch logs: {0}" -f $batchRoot)
        [void](New-Item -ItemType Directory -Path $batchRoot)
        $script:BatchLogPath = Join-Path $batchRoot 'batch.log'
        $summaryPath = Join-Path $batchRoot 'summary.csv'
    }
    Write-BatchLine 'Comparison batch preflight started.'
    $pwsh = (Get-Command pwsh -ErrorAction Stop).Source
    $LauncherPath = [IO.Path]::GetFullPath($LauncherPath)
    if (-not (Test-Path -LiteralPath $LauncherPath -PathType Leaf)) { throw "Migration launcher not found: $LauncherPath" }
    $migrations = @(Get-ComparisonMigrations -Root $ProjectRoot -Names $requested)
    if ($migrations.Count -eq 0) { throw 'No configured migrations found.' }
    Write-BatchLine ("Comparison plan: {0} migrations; {1} comparisons; authentication={2}; sequential execution." -f
        $migrations.Count, ($migrations.Count * $actions.Count), $AuthMode)
    foreach ($action in $actions) {
        foreach ($migration in $migrations) {
            Write-BatchLine ("PLAN {0} {1}" -f $migration.Name, $action)
            if ($migration.ConfigError) { Write-BatchLine ("INVALID {0}: {1}" -f $migration.Name, $migration.ConfigError) }
        }
    }
    if ($PlanOnly) {
        Write-BatchLine 'Plan only: no comparisons, authentication, or output created.'
        if (@($migrations | Where-Object ConfigError).Count) { exit 1 }
        return
    }
    if ($Force) { Write-BatchLine 'WARNING: -Force allows comparisons outside the configured scan age limits.' }
    Write-BatchLine 'Using the latest source/target inventories selected by each migration launcher. Missing or invalid evidence will fail that comparison.'
    if ($AuthMode -eq 'Interactive' -and $ComparisonMode -ne 'FilesOnly') {
        Write-BatchLine 'Permission comparisons may request Entra sign-in if the users cache needs refreshing.'
    }
    foreach ($action in $actions) {
        Write-BatchLine ("Starting phase: {0}" -f $action)
        foreach ($migration in $migrations) {
            $started = Get-Date
            $consoleLog = Join-Path $batchRoot ("{0}-{1}.console.log" -f $migration.Name,$action)
            $status = 'FAILED'; $code = 1; $message = ''
            Write-BatchLine -Message ("START {0} {1}" -f $migration.Name,$action) -ConsoleLog $consoleLog
            try {
                if ($migration.ConfigError) { throw $migration.ConfigError }
                $childArguments = @('-NoLogo','-NoProfile')
                if ($AuthMode -eq 'Certificate') { $childArguments += '-NonInteractive' }
                $childArguments += @('-File',$LauncherPath,'-MigrationName',$migration.Name,'-Action',$action,'-NonInteractive')
                if ($AuthMode -eq 'Certificate') { $childArguments += '-UseCertificate' }
                if ($Force) { $childArguments += '-Force' }
                $global:LASTEXITCODE = 0
                & $pwsh @childArguments 2>&1 | ForEach-Object {
                    Write-BatchLine -Message ("{0} {1}: {2}" -f $migration.Name,$action,[string]$_) -ConsoleLog $consoleLog
                }
                $code = $LASTEXITCODE
                if ($null -eq $code) { $code = 1 }
                if ($code -eq 0) { $status = 'SUCCESS' }
                else { $message = "Comparison launcher returned exit $code. See console and migration logs." }
            }
            catch {
                $message = $_.Exception.Message
                Write-BatchLine -Message ("ERROR {0} {1}: {2}" -f $migration.Name,$action,$message) -ConsoleLog $consoleLog
            }
            finally {
                $script:Results.Add([pscustomobject]@{
                    Migration=$migration.Name; Action=$action; Status=$status; ExitCode=$code
                    Started=$started.ToString('o'); Finished=(Get-Date).ToString('o'); ConsoleLog=$consoleLog; Error=$message
                })
                $script:Results.ToArray() | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding utf8
                Write-BatchLine ("{0}: {1} {2} (exit {3})" -f $status,$migration.Name,$action,$code)
            }
        }
    }
    $interrupted = $false
}
catch {
    Write-BatchLine ("Comparison batch failed: {0}" -f $_.Exception.Message)
    throw
}
finally {
    if ($summaryPath) {
        $script:Results.ToArray() | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding utf8
        $failed = @($script:Results | Where-Object Status -NE SUCCESS).Count
        Write-BatchLine ("Batch finished: {0} successful; {1} failed; interrupted={2}; summary: {3}" -f
            ($script:Results.Count-$failed),$failed,$interrupted,$summaryPath)
    }
}
if ($interrupted -or @($script:Results | Where-Object Status -NE SUCCESS).Count) { exit 1 }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDf2R3kc4ocgyER
# jTB/LY2y9mFoYWHiqsE/vfx9v1YZjaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBBNwAQ7qRdYoMulr0/lIlc
# xSYYU55M2DuR0CWECPlAfzANBgkqhkiG9w0BAQEFAASCAYBLI+KHYcaAnDchiOgI
# tNOhheJZa2It1iq6CiWLE/+tvGBTOTb2JY7Mg1xl3edzG1eKYRYGb3ZOqh46kPRi
# jZyvWpRShYvR+1SMSrvr8favOdao/7R+8AuMRqPM8bEzTdPNpjxaRQEfdVzJQ7zb
# UFvHvHDrulrTbGRerIwbcxIsMtNUlsCgrykT7isWeHR27q/JQaVxjE5Pa3QlrVVU
# if4tElnrSCntV+Te/OpB58tFNFOCB6UliyGc7OjKR2Jp2LGaIhkcDEgrY7VRKGlO
# AepDq7bu3QSDm5stFA6H61t8JLT/5XWPtQHwKp0hNN6Z17fblCF0YWo7m4S87rJW
# cbjurNiWeCTIJRIuJIx6BFhbSMdbd3hViW9HhRvYxJC7YDiy+JsCv0QltdU2/N/U
# UWnVUo3v6esE6WRxmeY/+JxA932UUydDbrxBk5K9FlVCX6hz/AavnipH3pFkc4I9
# ix6feFwbFY0iOsn5ll6v1q087fDHIM+0SAxRq2x0PZ35LHA=
# SIG # End signature block
