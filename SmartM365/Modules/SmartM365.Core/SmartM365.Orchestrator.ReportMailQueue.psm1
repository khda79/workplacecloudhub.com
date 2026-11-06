<#
.SYNOPSIS
    Offline queue primitives for deferred report mail.
.VERSION
    1.0.1
.DESCRIPTION
    Does not connect to mail services or select a production path. The caller must
    verify the queue directory and its access controls before each operation.
#>
Set-StrictMode -Version Latest

function Assert-ReportMailQueueDirectory {
    param([string]$Path, [scriptblock]$AssertTrustedDirectory)
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if (-not $item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw 'The report mail queue must be an existing, non-linked directory.'
    }
    if ($null -eq $AssertTrustedDirectory) { throw 'A trusted-directory verifier is required.' }
    $result = @(& $AssertTrustedDirectory $item.FullName)
    if ($result.Count -ne 1 -or $result[0] -isnot [bool] -or -not $result[0]) {
        throw 'The report mail queue directory was not verified as trusted.'
    }
    return $item.FullName
}

function Get-ReportMailQueueId {
    param([Parameter(Mandatory)][string]$TenantKey, [Parameter(Mandatory)][string]$DeduplicationKey)
    $bytes = [Text.Encoding]::UTF8.GetBytes(($TenantKey.Length.ToString([Globalization.CultureInfo]::InvariantCulture) + ':' + $TenantKey + $DeduplicationKey))
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Get-ReportMailPayloadHash {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Open-ReportMailQueueGuard {
    param([Parameter(Mandatory)][string]$Path)
    if (Test-Path -LiteralPath $Path) {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw 'Linked or invalid queue lock is not accepted.'
        }
    }
    return [IO.File]::Open($Path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
}

function Write-ReportMailQueueRecord {
    param([string]$Path, $Record, [switch]$Create)
    $tmp = Join-Path (Split-Path -Parent $Path) ('.' + [guid]::NewGuid().ToString('N') + '.pending')
    try {
        $json = ConvertTo-Json -InputObject $Record -Depth 32 -Compress
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($json)
        $stream = [IO.File]::Open($tmp, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
        if ($Create) { [IO.File]::Move($tmp, $Path) }
        else {
            $backup = Join-Path (Split-Path -Parent $Path) ('.' + [guid]::NewGuid().ToString('N') + '.backup')
            [IO.File]::Replace($tmp, $Path, $backup)
            Remove-Item -LiteralPath $backup -Force -ErrorAction Stop
        }
    } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } }
}

function Read-ReportMailQueueRecord {
    param([Parameter(Mandatory)][string]$Path)
    # A batch reader must not block another worker's atomic state replacement.
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {
        $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true)
        try { return ($reader.ReadToEnd() | ConvertFrom-Json -ErrorAction Stop) }
        finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

function Add-SmartM365ReportMailQueueItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueuePath,
        [Parameter(Mandatory)][scriptblock]$AssertTrustedDirectory,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TenantKey,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DeduplicationKey,
        [Parameter(Mandatory)][datetimeoffset]$NotBeforeUtc,
        [Parameter(Mandatory)][hashtable]$Message
    )
    $root = Assert-ReportMailQueueDirectory $QueuePath $AssertTrustedDirectory
    $id = Get-ReportMailQueueId $TenantKey $DeduplicationKey
    $payloadBytes = [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $Message -Depth 32 -Compress))
    $payloadHash = Get-ReportMailPayloadHash $payloadBytes
    $lockPath = Join-Path $root ($id + '.lock')
    $guard = Open-ReportMailQueueGuard $lockPath
    try {
        $null = Assert-ReportMailQueueDirectory $root $AssertTrustedDirectory
        $path = Join-Path $root ($id + '.json')
        if (Test-Path -LiteralPath $path) {
            $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked queue record is not accepted.' }
            $existing = Read-ReportMailQueueRecord -Path $path
            if ($existing.Id -ne $id -or $existing.TenantKey -cne $TenantKey) { throw 'Queue record identity mismatch.' }
            if ($existing.PayloadHash -cne $payloadHash) { throw 'Queue deduplication key is already used for a different message.' }
            return [pscustomobject]@{ Id = $id; Status = [string]$existing.Status; Existing = $true }
        }
        $record = [ordered]@{
            SchemaVersion = 1; Id = $id; TenantKey = $TenantKey; Status = 'Pending'
            CreatedAtUtc = [datetimeoffset]::UtcNow.ToString('o'); NotBeforeUnixMs = $NotBeforeUtc.ToUnixTimeMilliseconds()
            Attempts = 0; PayloadHash = $payloadHash; PayloadBase64 = [Convert]::ToBase64String($payloadBytes)
        }
        Write-ReportMailQueueRecord -Path $path -Record $record -Create
        return [pscustomobject]@{ Id = $id; Status = 'Pending'; Existing = $false }
    } finally { $guard.Dispose() }
}

function Invoke-SmartM365ReportMailQueueItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueuePath,
        [Parameter(Mandatory)][scriptblock]$AssertTrustedDirectory,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{64}$')][string]$Id,
        [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][datetimeoffset]$Now,
        [Parameter(Mandatory)][scriptblock]$SendAction,
        [ValidateRange(1, 1440)][int]$RetryMinutes = 5
    )
    $root = Assert-ReportMailQueueDirectory $QueuePath $AssertTrustedDirectory
    $guard = Open-ReportMailQueueGuard (Join-Path $root ($Id + '.lock'))
    try {
        $null = Assert-ReportMailQueueDirectory $root $AssertTrustedDirectory
        $path = Join-Path $root ($Id + '.json')
        $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked queue record is not accepted.' }
        $record = Read-ReportMailQueueRecord -Path $path
        if ($record.SchemaVersion -ne 1 -or $record.Id -cne $Id -or $record.TenantKey -cne $TenantKey -or $record.Status -notin @('Pending','Sending','Sent')) { throw 'Invalid queue record.' }
        if ($record.Status -eq 'Sent') { return 'Sent' }
        # The exclusive guard is held: Sending now means a previous outcome is unknown.
        # Never automatically replay it, even after a process restart.
        if ($record.Status -eq 'Sending') { return 'Uncertain' }
        if ($record.NotBeforeUnixMs -isnot [long] -and $record.NotBeforeUnixMs -isnot [int]) { throw 'Invalid queue delivery timestamp.' }
        if ($Now.ToUnixTimeMilliseconds() -lt [long]$record.NotBeforeUnixMs) { return 'Pending' }
        $payloadBytes = [Convert]::FromBase64String([string]$record.PayloadBase64)
        if ((Get-ReportMailPayloadHash $payloadBytes) -cne $record.PayloadHash) { throw 'Queue payload integrity check failed.' }
        $message = [Text.Encoding]::UTF8.GetString($payloadBytes) | ConvertFrom-Json
        $record.Attempts = [int]$record.Attempts + 1
        $record.Status = 'Sending'
        $record | Add-Member -NotePropertyName AttemptStartedAtUtc -NotePropertyValue ([datetimeoffset]::UtcNow.ToString('o')) -Force
        $record | Add-Member -NotePropertyName AttemptHost -NotePropertyValue ([Environment]::MachineName) -Force
        # This checkpoint must succeed before the transport can accept a message.
        Write-ReportMailQueueRecord -Path $path -Record $record
        try {
            $transportResult = @(& $SendAction $message)
            if ($transportResult.Count -ne 1 -or $transportResult[0] -isnot [bool] -or -not $transportResult[0]) {
                throw 'The transport did not explicitly confirm acceptance.'
            }
        } catch {
            $record.Status = 'Pending'
            $record.NotBeforeUnixMs = $Now.AddMinutes($RetryMinutes).ToUnixTimeMilliseconds()
            Write-ReportMailQueueRecord -Path $path -Record $record
            return 'Retry'
        }
        $record.Status = 'Sent'
        $record.PayloadBase64 = ''
        $record | Add-Member -NotePropertyName SentAtUtc -NotePropertyValue ([datetimeoffset]::UtcNow.ToString('o')) -Force
        try { Write-ReportMailQueueRecord -Path $path -Record $record }
        catch {
            Write-Warning ("Report mail {0}: transport accepted, but final state persistence failed. Automatic replay is blocked; reconcile delivery before recovery. {1}" -f $Id, $_.Exception.Message)
            return 'Uncertain'
        }
        return 'Sent'
    } finally { $guard.Dispose() }
}

function Invoke-SmartM365ReportMailQueueBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueuePath,
        [Parameter(Mandatory)][scriptblock]$AssertTrustedDirectory,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TenantKey,
        [Parameter(Mandatory)][datetimeoffset]$Now,
        [Parameter(Mandatory)][scriptblock]$SendAction,
        [ValidateRange(1, 1000)][int]$MaximumItems = 50
    )
    $root = Assert-ReportMailQueueDirectory $QueuePath $AssertTrustedDirectory
    $counts = [ordered]@{ Sent=0; Pending=0; Retry=0; Busy=0; Invalid=0; Uncertain=0 }
    $items = @(Get-ChildItem -LiteralPath $root -Filter '*.json' -File -ErrorAction Stop | Where-Object { $_.BaseName -cmatch '^[0-9a-f]{64}$' } | Sort-Object Name)
    $attempted = 0
    foreach ($item in $items) {
        if ($attempted -ge $MaximumItems) { break }
        try {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked queue record is not accepted.' }
            $metadata = Read-ReportMailQueueRecord -Path $item.FullName
            if ($metadata.SchemaVersion -ne 1 -or $metadata.Id -cne $item.BaseName -or $metadata.Status -notin @('Pending','Sending','Sent')) { throw 'Invalid queue record.' }
            if ($metadata.TenantKey -cne $TenantKey -or $metadata.Status -eq 'Sent') { continue }
            if ($metadata.NotBeforeUnixMs -isnot [long] -and $metadata.NotBeforeUnixMs -isnot [int]) { throw 'Invalid queue delivery timestamp.' }
            if ($Now.ToUnixTimeMilliseconds() -lt [long]$metadata.NotBeforeUnixMs) { $counts.Pending++; continue }
            $attempted++
            $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $root -AssertTrustedDirectory $AssertTrustedDirectory -Id $item.BaseName -TenantKey $TenantKey -Now $Now -SendAction $SendAction
            if ($status -eq 'Uncertain') { $attempted-- }
            $counts[$status]++
        } catch [IO.IOException] { $counts.Busy++ }
        catch { $counts.Invalid++ }
    }
    return [pscustomobject]$counts
}

Export-ModuleMember -Function Add-SmartM365ReportMailQueueItem, Invoke-SmartM365ReportMailQueueItem, Invoke-SmartM365ReportMailQueueBatch

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCnoMbpkZQynJC/
# xbfbdA+vY9dNZ4svub46yYuOFe2XwqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD8PMFihrcYFngiXZBBeTW0
# DxbFgoFGLi3uPXeN1SP+TzANBgkqhkiG9w0BAQEFAASCAYBxV6rsnz4/I7NNbKJQ
# ZnIHQjXLMhOGJH79uSf+Z7+bmA+Mw/IeWrjfcbhNxNJ49Fli+HHmFiA7VTsfr+5W
# CpK0YSA5mVGuDMYKAOTMgkJH//BC94+JpNZFjVJrPcclzsw9GA88CBFdrq39UzC5
# L66iwwqQwBz9UGBvoge6KsWrRFRX+x5drf/THQa6959GiQqEhYZ8byzZjejaU+ii
# PvI4NHKO+kWegRTUqC9sOdkEOuRr50Cf8qGLcY/6oNIM4atHRnRMit5JrfGhqFkg
# 0dwerZ5lVbJSIqXWgvnucQA5JvNDdZlvtkc6QEKSAhE3TRWSGoWXvE9jJObs0COP
# gXyyTRdwNu95kzEc1a56b0xbHTZVHXOP6BVhW9Nfuwz7Mc3s6n1ihNnrXMUgVdf+
# jBEYcloMFtNSAnJMjYIT7fHVdsxHYsWcHCg3HRmsnfFJYJg5Q2iAh94jJsmeOq55
# RhWebmEGPt6XINO/WcIcLPGuMtH/LBM7P97qFb9Lwmf1qz8=
# SIG # End signature block
