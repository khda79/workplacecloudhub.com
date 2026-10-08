Set-StrictMode -Version 2.0

function ConvertTo-SmartM365OrchestratorJobStopUtc {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetimeoffset]) { return $Value.ToUniversalTime() }
    if ($Value -is [datetime]) { return [datetimeoffset]::new($Value.ToUniversalTime(), [timespan]::Zero) }
    $parsed = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
        return $parsed.ToUniversalTime()
    }
    return $null
}

function Get-SmartM365OrchestratorJobStopFolder {
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$Server)
    if ($Server -notmatch '^[A-Za-z0-9._-]+$' -or $Server -in @('.', '..')) { throw 'Invalid orchestrator server name.' }
    Join-Path (Join-Path $SharedDataFolderPath $Server) 'JobStopRequests'
}

function Request-SmartM365OrchestratorJobStop {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$JobName,
        [Parameter(Mandatory)][int]$ProcessId,
        [Parameter(Mandatory)][string]$StartTime,
        [Parameter(Mandatory)][string]$Reason,
        [Parameter(Mandatory)][string]$Tenant
    )
    $Reason = $Reason.Trim()
    if ($Reason.Length -lt 3 -or $Reason.Length -gt 1000) { throw 'Enter a stop reason (3 to 1000 characters).' }
    if ([string]::IsNullOrWhiteSpace($JobName) -or $ProcessId -le 0) { throw 'Select a running job with a valid process ID.' }
    $start = ConvertTo-SmartM365OrchestratorJobStopUtc $StartTime
    if ($null -eq $start) { throw 'The selected job has no valid start time.' }
    $folder = Get-SmartM365OrchestratorJobStopFolder $SharedDataFolderPath $Server
    $heartbeatPath = Join-Path (Join-Path $SharedDataFolderPath $Server) 'Orchestrator-Heartbeat.json'
    $heartbeat = Read-SmartM365OrchestratorJson -Path $heartbeatPath
    $stamp = ConvertTo-SmartM365OrchestratorJobStopUtc $heartbeat.Timestamp
    if ($null -eq $stamp -or
        ([datetimeoffset]::UtcNow - $stamp).TotalMinutes -gt 5 -or
        ($stamp - [datetimeoffset]::UtcNow).TotalMinutes -gt 5 -or
        [string]$heartbeat.Lifecycle -ne 'Running' -or
        -not $heartbeat.PSObject.Properties['JobStopProtocol'] -or [int]$heartbeat.JobStopProtocol -lt 1 -or
        [string]$heartbeat.Tenant -ne $Tenant) { throw 'The owning orchestrator is not ready to accept job stop requests.' }
    $matching = @($heartbeat.RunningJobs | Where-Object {
        $currentStart = ConvertTo-SmartM365OrchestratorJobStopUtc $_.StartTime
        [string]$_.Name -eq $JobName -and [int]$_.Pid -eq $ProcessId -and
        $null -ne $currentStart -and $currentStart.Ticks -eq $start.Ticks
    })
    if ($matching.Count -ne 1) { throw 'The selected run changed. Refresh Operations and select the current run.' }
    $requestId = [guid]::NewGuid().ToString('N')
    $request = [pscustomobject][ordered]@{
        SchemaVersion = 1; RequestId = $requestId; Tenant = $Tenant; Server = $Server
        JobName = $JobName; ProcessId = $ProcessId; StartTime = $StartTime
        RequestedAtUtc = [datetimeoffset]::UtcNow.ToString('o')
        RequestedBy = [Environment]::UserName; RequestedFrom = [Environment]::MachineName
        Reason = $Reason; Status = 'Requested'; UpdatedAtUtc = ''; Detail = ''
    }
    $path = Join-Path $folder ("{0}_{1}.json.txt" -f [datetimeoffset]::UtcNow.ToString('yyyyMMddTHHmmssfffZ'), $requestId)
    Write-SmartM365OrchestratorJsonAtomically -Path $path -Document $request -CreateNew
    [pscustomobject]@{ RequestId = $requestId; Path = $path; Server = $Server; JobName = $JobName }
}

function Get-SmartM365OrchestratorJobStopRequests {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$Server)
    $folder = Get-SmartM365OrchestratorJobStopFolder $SharedDataFolderPath $Server
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { return @() }
    @(
        foreach ($file in @(Get-ChildItem -LiteralPath $folder -Filter '*.json.txt' -File -ErrorAction Stop)) {
            try { $document = Read-SmartM365OrchestratorJson -Path $file.FullName }
            catch { Write-Warning ("Malformed job stop request skipped: {0}: {1}" -f $file.FullName, $_.Exception.Message); continue }
            if ([string]$document.Status -in @('Requested', 'Stopping')) {
                [pscustomobject]@{ Path = $file.FullName; Document = $document }
            }
        }
    )
}

function Set-SmartM365OrchestratorJobStopRequestStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('Stopping', 'Stopped', 'AlreadyExited', 'Rejected')][string]$Status,
        [string]$Detail = ''
    )
    $request = Read-SmartM365OrchestratorJson -Path $Path
    $request.Status = $Status
    $request.UpdatedAtUtc = [datetimeoffset]::UtcNow.ToString('o')
    $request.Detail = $Detail
    Write-SmartM365OrchestratorJsonAtomically -Path $Path -Document $request
}

Export-ModuleMember -Function ConvertTo-SmartM365OrchestratorJobStopUtc, Get-SmartM365OrchestratorJobStopFolder, Request-SmartM365OrchestratorJobStop, Get-SmartM365OrchestratorJobStopRequests, Set-SmartM365OrchestratorJobStopRequestStatus

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA4xYQw4N042ghE
# 6673s5UCFg9ARc6ovCsHrrMf5NL9TKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBu3ruzJDcAMuwaYJOGgjKR
# HA6aWL8nkMmklhZ9mk7roDANBgkqhkiG9w0BAQEFAASCAYBdI4ov+YKS2QJ5leP+
# bgFh84bqFhk0oPXk7Ae4dLIvgT6GnxG/pknIM1ojy6ibodPIPM9Oy8I0aJDip5h5
# 8MWnIJuNKU0y8UR4Y/jQBGZTJtnflzMORGGthWzMtWQ25i84W6D87zTa2BS0J4SW
# g8CfjJYm9FjmXE9CD11+SLIODgeY4h/croxQtMhOSL7hsk4sCcgnWvdK51yCrUj2
# SL31Uk+sFVBMU1sUPjovuQjpB8ropFcpzzRiC8q3K+ioNYS/gkHhwCbkUGK3+0n2
# lJRUuGokd462J3pD+CyYloWY1Ypx2ux96vnZLwp04mWSo+Z+pN6zIttEhxTcLrmI
# Nausue8OPkXjzXwW828P5dOuj7rK5+tfhU2nZZ61iwmIsiJgiMofaw/WEdc/sFYI
# 0hUhsUs/R17eIVcsExTBSKF7fYbrmMNKry1QeePMEujbJqPN5R8caIp2d9BOv6/R
# X8eNUNsn/jyiUbCo/piNhNzGSP0fAgnxy+Gwmj/gJfjJ9Qs=
# SIG # End signature block
