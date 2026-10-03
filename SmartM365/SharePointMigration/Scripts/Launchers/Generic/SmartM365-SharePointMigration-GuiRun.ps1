<#
.SYNOPSIS
    Runs a SharePointMigration entry script.
.VERSION
    1.0.2
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ActivityPath,
    [Parameter(Mandatory = $true)][string]$MigrationName,
    [string]$Action,
    [string]$OperationPath,
    [ValidateSet('Interactive', 'DeviceLogin', 'Certificate')][string]$AuthMode = 'Interactive',
    [string]$SourceCsv,
    [string]$TargetCsv,
    [string]$OldCsv,
    [string]$NewCsv,
    [ValidateSet('Source', 'Target')][string]$HistorySide = 'Source'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-GuiActivity.ps1')
$projectRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$migrationRoot = [System.IO.Path]::GetFullPath((Join-Path (Join-Path $projectRoot 'Migrations') $MigrationName))
$activityRoot = Get-SmartM365GuiActivityDirectory -ProjectRoot $projectRoot

try {
    if (-not [System.IO.Path]::GetFullPath($ActivityPath).StartsWith(
            ($activityRoot.TrimEnd('\') + '\'), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Activity path is outside the shared activity directory.'
    }
    $runId = (Read-SmartM365GuiActivity -Path $ActivityPath).Id
    if (-not (Test-Path -LiteralPath $migrationRoot -PathType Container)) {
        throw "Migration folder not found: $MigrationName"
    }
    Write-SmartM365GuiActivityEvent -Path $ActivityPath -Status 'Running' -Detail 'Child process starting.'
    if ($OperationPath) {
        $operationFull = [System.IO.Path]::GetFullPath($OperationPath)
        if (-not $operationFull.StartsWith(($migrationRoot.TrimEnd('\') + '\'),
                [System.StringComparison]::OrdinalIgnoreCase) -or
            -not (Test-Path -LiteralPath $operationFull -PathType Leaf)) {
            throw 'Operation launcher must be inside the selected migration.'
        }
        $executable = 'cmd.exe'
        $arguments = @('/d', '/c', ('"{0}"' -f $operationFull))
    }
    else {
        $executable = if ($PSVersionTable.PSVersion.Major -ge 7) { (Get-Process -Id $PID).Path } else { 'powershell.exe' }
        $launcher = Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-Launcher.ps1'
        $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $launcher),
            '-MigrationName', ('"{0}"' -f $MigrationName), '-Action', $Action)
        if ($AuthMode -eq 'DeviceLogin') { $arguments += '-DeviceLogin' }
        if ($AuthMode -eq 'Certificate') { $arguments += '-UseCertificate' }
        foreach ($entry in @(@('SourceCsv', $SourceCsv), @('TargetCsv', $TargetCsv),
                @('OldCsv', $OldCsv), @('NewCsv', $NewCsv))) {
            if ($entry[1]) { $arguments += @('-' + $entry[0], ('"{0}"' -f $entry[1])) }
        }
        if ($Action -in @('CompareScanHistory', 'ComparePermissionScanHistory')) {
            $arguments += @('-HistorySide', $HistorySide)
        }
    }
    $env:SPMIG_GUI_RUN_ID = $runId
    $process = Start-Process -FilePath $executable -ArgumentList $arguments -WorkingDirectory $projectRoot `
        -NoNewWindow -Wait -PassThru -ErrorAction Stop
    Remove-Item Env:SPMIG_GUI_RUN_ID -ErrorAction SilentlyContinue
    $code = [int]$process.ExitCode
    $script:ConsoleLifecycleStatus = if ($code -eq 0) { 'SUCCESS' } elseif ($code -eq 2) { 'CANCELLED' } else { 'FAILED' }
    $logPath = ''
    $logsDirectory = if ($OperationPath) {
        Join-Path $projectRoot 'Scripts\Operations\logs'
    } else { Join-Path $migrationRoot 'logs' }
    if (Test-Path -LiteralPath $logsDirectory -PathType Container) {
        $match = Get-ChildItem -LiteralPath $logsDirectory -Filter "*-$runId.log" -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($match) { $logPath = $match.FullName }
    }
    $status = if ($code -eq 0) { 'Succeeded' } elseif ($code -eq 2) { 'Cancelled' } else { 'Failed' }
    $detail = "Child process exited with code $code."
    if ($OperationPath -and $code -eq 0 -and $logPath) {
        $modeLine = Select-String -LiteralPath $logPath -Pattern 'WhatIfMode:\s*(True|False)' `
            -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($modeLine) {
            if ($modeLine.Matches[0].Groups[1].Value -eq 'True') {
                $status = 'Previewed'
                $detail = 'Site operation preview completed; no change applied.'
            }
            else {
                $status = 'Applied'
                $detail = 'Site operation applied; inspect the run log for details.'
            }
        }
    }
    Write-SmartM365GuiActivityEvent -Path $ActivityPath -Status $status -ExitCode $code `
        -LogPath $logPath -Detail $detail
    Write-Host ("{0} {1}: {2} (exit code {3}). Activity: {4}" -f `
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $MigrationName,
        $(if ($OperationPath) { 'Operation' } else { $Action }), $code, $ActivityPath)
}
catch {
    Remove-Item Env:SPMIG_GUI_RUN_ID -ErrorAction SilentlyContinue
    Write-SmartM365GuiActivityEvent -Path $ActivityPath -Status 'Failed' -ExitCode 1 `
        -Detail $_.Exception.Message
    Write-Error $_
}
}
catch {
    $script:ConsoleLifecycleFailure = $_
    throw
}
finally {
    Complete-SmartM365MigrationConsoleLifecycle -Context $script:ConsoleLifecycleContext -Failure $script:ConsoleLifecycleFailure -Status $script:ConsoleLifecycleStatus
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAsju3nYlDWoFe6
# XNjP9kwUuK59/XrOsm5sCucwyUV1nKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBDo4oDPqz9yZb+01ellQ1Z
# ZwVcXzjxU8gGVIgwu5xHezANBgkqhkiG9w0BAQEFAASCAYBL0veMmzgr9olUI5p1
# M4MPuHey/VjI+x08UqhL2JAP9ncvsuptnNi5NtjK13HW7Xs0RLf2MKnA0WXyaEjG
# KcbFZrAvmQQ/4/HDkUpNdSRnMhVDdIBvfn6s8b9oYyOM1uipHQcOiEC1AkuCv2bu
# TiVUBIL14pQ+XYiWX7rUSnEBMGc5daE4xm+OY9EhzbYHHcxOz5uMCf73/+EoGCNz
# MDB+8WfTT0Oz/wFhO2R8RxKVv1JoTJMi9YULP/EvSBhi65CccJB7LJtsAR4bZDrn
# lsRKdAxiOT1EajpLgnwZyVs97rjY695sZbuyQ8no2OHAtrqsnpL5l958g5zreGx9
# 63KhzqLMvd6cOIgLETMyopzBJSlGI0qth0v0AmLmZ+5vn8TqElYQXScIphLwkBpG
# kDgwcgtDjtxoHPAjWF0LTl5iJysTTfR6wcEDzuoYxadgHzpo+h9dlHSkVF/6BrhV
# XoZc2cw5CSG72PqW8cgqw+IUm+Uh9EbpgFBtCdZWmn5haQA=
# SIG # End signature block
