<#
.SYNOPSIS
    Offline queue primitives for deferred report mail.
.VERSION
    1.2.0
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
        [Parameter(Mandatory)][hashtable]$Message,
        [ValidatePattern('^(?:[0-9a-f]{64})?$')][string]$SupersedeKey = ''
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
        if ($SupersedeKey) { $record.SupersedeKey = $SupersedeKey }
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
        [scriptblock]$IsExpired,
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
        if ($IsExpired -and (& $IsExpired $message)) {
            # Keep the terminal status readable by existing queue metadata tools,
            # and record expiry separately from a successful transport submission.
            $record.Status = 'Sent'
            $record.PayloadBase64 = ''
            $record | Add-Member -NotePropertyName Expired -NotePropertyValue $true -Force
            $record | Add-Member -NotePropertyName ExpiredAtUtc -NotePropertyValue $Now.ToUniversalTime().ToString('o') -Force
            Write-ReportMailQueueRecord -Path $path -Record $record
            return 'Expired'
        }
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

function Set-SmartM365ReportMailQueueItemSuperseded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$QueuePath,
        [Parameter(Mandatory)][scriptblock]$AssertTrustedDirectory,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{64}$')][string]$Id,
        [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{64}$')][string]$SupersededBy
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
        if ($record.Status -eq 'Sending') { return 'Uncertain' }
        # Recorded as Sent for readers of earlier versions; the payload is no longer needed.
        $record.Status = 'Sent'
        $record.PayloadBase64 = ''
        $record | Add-Member -NotePropertyName Superseded -NotePropertyValue $true -Force
        $record | Add-Member -NotePropertyName SupersededBy -NotePropertyValue $SupersededBy -Force
        $record | Add-Member -NotePropertyName SentAtUtc -NotePropertyValue ([datetimeoffset]::UtcNow.ToString('o')) -Force
        Write-ReportMailQueueRecord -Path $path -Record $record
        return 'Superseded'
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
        [scriptblock]$IsExpired,
        [ValidateRange(1, 1000)][int]$MaximumItems = 50
    )
    $root = Assert-ReportMailQueueDirectory $QueuePath $AssertTrustedDirectory
    $counts = [ordered]@{ Sent=0; Pending=0; Retry=0; Busy=0; Invalid=0; Uncertain=0; Superseded=0; Expired=0 }
    $items = @(Get-ChildItem -LiteralPath $root -Filter '*.json' -File -ErrorAction Stop | Where-Object { $_.BaseName -cmatch '^[0-9a-f]{64}$' } | Sort-Object Name)
    # Latest pending record per supersede key; older pending copies are not delivered.
    $latest = @{}
    foreach ($item in $items) {
        try {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
            $metadata = Read-ReportMailQueueRecord -Path $item.FullName
            $keyProperty = $metadata.PSObject.Properties['SupersedeKey']
            if (-not $keyProperty -or -not $keyProperty.Value -or $metadata.TenantKey -cne $TenantKey -or $metadata.Status -ne 'Pending' -or $metadata.Id -cne $item.BaseName) { continue }
            # PowerShell 7 may already convert the ISO timestamp to DateTime; keep its full precision.
            $created = $metadata.CreatedAtUtc
            $createdUtc = if ($created -is [datetime]) { [datetimeoffset]$created.ToUniversalTime() } elseif ($created -is [datetimeoffset]) { $created.ToUniversalTime() } else { [datetimeoffset]::Parse([string]$created, [Globalization.CultureInfo]::InvariantCulture) }
            $candidate = [pscustomobject]@{ Id = [string]$metadata.Id; CreatedAtUtc = $createdUtc }
            $current = $latest[[string]$keyProperty.Value]
            if ($null -eq $current -or $candidate.CreatedAtUtc -gt $current.CreatedAtUtc -or ($candidate.CreatedAtUtc -eq $current.CreatedAtUtc -and [string]::CompareOrdinal($candidate.Id, $current.Id) -gt 0)) {
                $latest[[string]$keyProperty.Value] = $candidate
            }
        } catch { continue }
    }
    $attempted = 0
    foreach ($item in $items) {
        if ($attempted -ge $MaximumItems) { break }
        try {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Linked queue record is not accepted.' }
            $metadata = Read-ReportMailQueueRecord -Path $item.FullName
            if ($metadata.SchemaVersion -ne 1 -or $metadata.Id -cne $item.BaseName -or $metadata.Status -notin @('Pending','Sending','Sent')) { throw 'Invalid queue record.' }
            if ($metadata.TenantKey -cne $TenantKey -or $metadata.Status -eq 'Sent') { continue }
            $keyProperty = $metadata.PSObject.Properties['SupersedeKey']
            if ($metadata.Status -eq 'Pending' -and $keyProperty -and $keyProperty.Value -and $latest.ContainsKey([string]$keyProperty.Value) -and $latest[[string]$keyProperty.Value].Id -cne $item.BaseName) {
                $counts[(Set-SmartM365ReportMailQueueItemSuperseded -QueuePath $root -AssertTrustedDirectory $AssertTrustedDirectory -Id $item.BaseName -TenantKey $TenantKey -SupersededBy $latest[[string]$keyProperty.Value].Id)]++
                continue
            }
            if ($metadata.NotBeforeUnixMs -isnot [long] -and $metadata.NotBeforeUnixMs -isnot [int]) { throw 'Invalid queue delivery timestamp.' }
            if ($Now.ToUnixTimeMilliseconds() -lt [long]$metadata.NotBeforeUnixMs) { $counts.Pending++; continue }
            $attempted++
            $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $root -AssertTrustedDirectory $AssertTrustedDirectory -Id $item.BaseName -TenantKey $TenantKey -Now $Now -SendAction $SendAction -IsExpired $IsExpired
            if ($status -eq 'Uncertain') { $attempted-- }
            $counts[$status]++
        } catch [IO.IOException] { $counts.Busy++ }
        catch { $counts.Invalid++ }
    }
    return [pscustomobject]$counts
}

Export-ModuleMember -Function Add-SmartM365ReportMailQueueItem, Invoke-SmartM365ReportMailQueueItem, Invoke-SmartM365ReportMailQueueBatch, Set-SmartM365ReportMailQueueItemSuperseded

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBu9dZTY7aPZ0y+
# 6dxIViYW7ZM9LKJ9qRqed3ADKPQU1KCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgAceGKDnIQVA9
# YDrGQplhKRmbRED97AqGcL6m39xqWYcwDQYJKoZIhvcNAQEBBQAEggGAqWrMlZAQ
# ZJetC21tFq+fj6LH4koCrQsRNtvRiTickFrEs2P7bAyd7ZJwJIryJaeO0xeCtpX0
# TPssLk4LcHxWWSrZKGe8MBOyroMiuiq04vA9TnFDJMjKxqvZKvO8sgGXKbOxlKs4
# 2H3PUHrMsoEU9/hNuHLxMrwXros9RIQEZewchyZ5FoAyOD8Epz/JJweruCZ/DhMo
# smgvF73k5LKXjLCHJPobMc3hJIKy+Q2ag1CTiQb4HY2AJm1m7KMWrDmY0QOrrZug
# 1u/FD1Uvsv00wJ5GpAbhutshHYPvwh/rcFWoB9s4gHYwlA7anBxys+jZt9lbPQ5Q
# NtodWpPTOEkcqeQBK8E9fs/Nl9fBPMB6jCV5pJLCtMQuZkfqXzIOolSa/dsIIiRU
# 5sw/fe6jYfQ9keBPZMw/X7tlYPi0IpLDTSPQ0/+vsilCEJIradKFEWWZI7nLwF2b
# efHuYQD5rcJ8Sjz5Z4RCCElKhTIliTXtixj6Yv74K1RnUVJAVuVhWDw4oYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDgwNjA5WjA/BgkqhkiG9w0BCQQxMgQwIHoDbxqIZNREkHaHA5XwRMR1
# Rd2NVK6a+cBsDXPJ2N/tv+Sl1Th42g8vjozHUAglMA0GCSqGSIb3DQEBAQUABIIC
# AGboKR1qtw/250PTKEEjbfIP7YW3TS3lQoqpq913RQgrtAOMi+bjVyFKoEJgr5gr
# Bbmw0/uyuqG5kc8q4Ztckehpu0pS7Ejdv2H1BmCztyzO2AGHuzkDl1HQIXPUEtmj
# ovMJ1UxymgxNlw7/cYMVoueRXwW7RUygJvYX1ZW/T7dJem1hi29StQCzGc7g3K8Z
# rqf823pkn5rErcVXuNWyXRE9FAvbKJCegvyNqj0uwHwLKdjQmJj47zbqfP4n6WQ2
# UluVE0yz3PszFxOtHgJDHWiCCcMQ0wZfiHtdmOqe2xaQjhl5ffUG57vvphRK+w3H
# miDMcY4ZvMF4RZi5mXT+waLwfz3ZN1nHCWdAtgcOmiKJGKb8246mTol7oyvurunr
# OIx1BpbOTO6IDXVcc5/uZY0QgHTxyJsvpRmpY1uuUCSEtPp1uz+7te9G647yc0iS
# AnUYi5ylxIOYv9Oib/DgY7JRL4J2VtNeXdKUqAGwO3OjL8HdtEZ3+oRTdt67GHXR
# 7R9kPj81FCPIjQuqltZd5tASXfw4X4VG9xHU0Cs/26dwNaSxC1frbmzFUKcsQESI
# RHSVgZp2fQOsYwBAteVKQopn4U0mAjCwpBL3lgwMtDvMa+eK7nI4+TxDz1C7qoDb
# e8zwsSY4ToYQMFy5MlJJ1tC+cVDlGLF+liLroMOUHn1i
# SIG # End signature block
