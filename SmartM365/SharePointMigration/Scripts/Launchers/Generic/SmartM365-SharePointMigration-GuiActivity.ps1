function Get-SmartM365GuiActivityDirectory {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)
    $directory = Join-Path $ProjectRoot 'Migrations\logs\gui-activity'
    [void][System.IO.Directory]::CreateDirectory($directory)
    return $directory
}

function Write-SmartM365GuiActivityEvent {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Status,
        [string]$Detail = '',
        [int]$ExitCode = -1,
        [string]$LogPath = '',
        [string]$Migration = ''
    )
    $record = [ordered]@{
        Status = $Status
        Detail = $Detail
        ExitCode = $ExitCode
        LogPath = $LogPath
        Migration = $Migration
    }
    $line = '{0} {1}{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        ([char]0x7C), ($record | ConvertTo-Json -Compress -Depth 4)
    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($line + "`r`n")
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Append,
        [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
    try { $stream.Write($bytes, 0, $bytes.Length) }
    finally { $stream.Dispose() }
}

function New-SmartM365GuiActivity {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectRoot,
        [Parameter(Mandatory = $true)][string]$Migration,
        [Parameter(Mandatory = $true)][string]$Action
    )
    $directory = Get-SmartM365GuiActivityDirectory -ProjectRoot $ProjectRoot
    $id = [guid]::NewGuid().ToString('N')
    $path = Join-Path $directory ('{0}-{1}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $id)
    $actor = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $record = [ordered]@{
        Id = $id
        Utc = [DateTime]::UtcNow.ToString('o')
        Actor = $actor
        Machine = [Environment]::MachineName
        Migration = $Migration
        Action = $Action
        Status = 'Started'
    }
    $line = '{0} |{1}{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        ($record | ConvertTo-Json -Compress -Depth 4), "`r`n"
    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($line)
    $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
    try { $stream.Write($bytes, 0, $bytes.Length) }
    finally { $stream.Dispose() }
    return $path
}

function Read-SmartM365GuiActivity {
    param([Parameter(Mandatory = $true)][string]$Path)
    $lines = [System.IO.File]::ReadAllLines($Path)
    if ($lines.Count -eq 0) { return $null }
    $first = ($lines[0] -split ' \|', 2)[1] | ConvertFrom-Json
    if (-not $first) { return $null }
    $last = $first
    foreach ($line in ($lines | Select-Object -Skip 1)) {
        try { $last = ($line -split ' \|', 2)[1] | ConvertFrom-Json }
        catch { continue }
    }
    $displayTime = ([datetime]$first.Utc).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $status = [string]$last.Status
    if ($status -in @('Started', 'Running', 'Validating') -and
        [datetime]::UtcNow.Subtract(([datetime]$first.Utc).ToUniversalTime()).TotalHours -ge 24) {
        $status = 'Unconfirmed (>24h)'
    }
    [pscustomobject]@{
        Id = $first.Id
        Utc = $first.Utc
        Actor = $first.Actor
        Machine = $first.Machine
        Migration = if ($last.Migration) { $last.Migration } else { $first.Migration }
        Action = $first.Action
        Status = $status
        Detail = $last.Detail
        ExitCode = $last.ExitCode
        LogPath = $last.LogPath
        FullName = $Path
        Display = ('{0}  {1}  {2}  {3}  {4}' -f $displayTime,
            $first.Actor, $(if ($last.Migration) { $last.Migration } else { $first.Migration }), $first.Action, $status)
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqpi/cLSiOmIr1
# /bpBc6cFyrbGqwRSaspa2vsaB1nljKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAClcYeNeY+gpXKrxYY+5pp
# jlzo3NWQnQM/HK47Bz+xFzANBgkqhkiG9w0BAQEFAASCAYCA3pzLPYcjXNRhCa57
# KF5Iul/pgcs4YR9+ta6EdHHE8SP1uueKcLTlutPd2cWNJzh0r4XIe/Y53GDQkZMm
# U8x+99YqMUNJ+PbSpemXsA01roZGTpbA2BPncG9OWOEcwB1f8O2uSTv1xHFZOO90
# 3dU+ZrL9TX4GJ6gAKh0zokYDoajpm/a3Lzd+EldHnN9lMHjUVbaLRiCHZGHH7LCh
# ycIWB1ZSrtWYt6zqWSmVR/RfgvLJs5fuHHSop2LuwXOnttwMm5fLLA4Ts4wBtbm4
# v1Debv7m3b4Y+nOLOiKS/8kFoY0ISd8CvCOWZ7BpTP1p+UmzAiyCLkPYm0LQKs0h
# QezinG82plN3ZS4G7gbKCYgFD94oNA6oJFyMqtzMH98bkyZBoCr4wG55IbL6rAl+
# CAkw+BC7Ma6fq1WpakRQEXkhyuxVWRplAPJzkspZ6A2Wh4mMOM9AVmjpOH7cOTIu
# WjgFl7zfeErrPzqFyOCrtMQc17b43Nw++oXgIjmHIRd1bzY=
# SIG # End signature block
