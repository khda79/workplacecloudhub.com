<#
.SYNOPSIS
    Offline queue primitives for deferred report mail.
.VERSION
    1.1.0
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
        [ValidateRange(1, 1000)][int]$MaximumItems = 50
    )
    $root = Assert-ReportMailQueueDirectory $QueuePath $AssertTrustedDirectory
    $counts = [ordered]@{ Sent=0; Pending=0; Retry=0; Busy=0; Invalid=0; Uncertain=0; Superseded=0 }
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
            $status = Invoke-SmartM365ReportMailQueueItem -QueuePath $root -AssertTrustedDirectory $AssertTrustedDirectory -Id $item.BaseName -TenantKey $TenantKey -Now $Now -SendAction $SendAction
            if ($status -eq 'Uncertain') { $attempted-- }
            $counts[$status]++
        } catch [IO.IOException] { $counts.Busy++ }
        catch { $counts.Invalid++ }
    }
    return [pscustomobject]$counts
}

Export-ModuleMember -Function Add-SmartM365ReportMailQueueItem, Invoke-SmartM365ReportMailQueueItem, Invoke-SmartM365ReportMailQueueBatch, Set-SmartM365ReportMailQueueItemSuperseded

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDS0Ddp9z+JOxQn
# oM/1CdCMzBoMgKewazJQ/lg5jiODFqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILwSmdz3ydu+Jx/LOL7wJmOU+32a7JjdhyXruKYAaOx7MA0GCSqG
# SIb3DQEBAQUABIIBgDw79ADq6YUbkVxGr1kBQmjAh5UcvAqtedW8TIECj6dW9EdF
# xlGDHTnGo1kWlGmTFpdFMcjTZIX3Mpigg2hZ/XN/lmkjQh+PzVNT235PXs6697G6
# kanSMVck3PC575lOwJPYGy+v1lbNTwgllqX16c3lgHywZWR4jf4y7JvZC/s6+EW+
# JaSrnXDwiMawOzwm1gdyt454f/6l6GYUGCoVP8k7LhmIDD8Ipj2ANQdoDUn9onsY
# BzRYc+9YU115xliwpgtPzD0bkAwSd9Ue4FqUbzcs+EH58SZge7RXoBwcSIOKNNkK
# dO+JqeS9FQ55Jsaut71/kHqH4rmTO7TdsvkMBUKX/LZKJgnIeB8Ei7lJ2HIi+Swy
# q++s2+jLZ3+3bfBXzEWTXVoJlGna1Whhyh+/T/mGXjJRaFiWMdl+O34rLdvfFmOE
# yWNMJKMjJuiiPYJ27JssfTSX8fpGYLOPIO9Uf8sUoGOqLIK61fDrYY9MTv3JgzYW
# 1TiaK7mjwC9p2W7jhaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTgxMzAx
# NDhaMC8GCSqGSIb3DQEJBDEiBCDXu+4P2TEuOlDtlHsBrPv6EPRHzF2cETweuKjO
# qkG7MTANBgkqhkiG9w0BAQEFAASCAgCPIeE7BDyCLps8j4hQnPVWygNfX1Y4gIFm
# PBsN1KKY48G3cwF7HHHt3RlFKoXHv9aFGzG156gjEjw2Sc18NI05vNR6z+CupL0T
# MDkYTP4G8yp7DR0qK1vrX6yHK+joRdM31OIjV1kP92m24pBlcKt3Ubi/lbhQlypu
# A9yaOhP1HbCiDHyKWztK822ULX7czVC74eJyJAtBG/e/f6eeQP1cpbqHv9O0fHFV
# AhJiLqck5dkdy9ADVGSROrd9G4tGCjt7eGD78jHguLGI0iI+GeewevO/0eePvQvD
# a1+8E2bAUOh89NLpw+JN0VNGZCbywyg9Si7dRStzKXXaU/uB2tLSmj6OciBXrlYJ
# 5DLk1Xy0TYaqAVqBUSKTOKL3am8h+BteTcUE9ThzPLb7vKg+Q6JaVZdBuSCcH0Dr
# vlwtOb6e0ZPwjkKtstuQiuLxuuxKphCrHWnI0KUeWQJ9koPBm/j3Hs5E1qVsxeEr
# JNVIh1YaofgehC/oUnmX9MCmc/td8E3WTxN6spp6Cn4PvypvKyO1E0iE0QsmGS/j
# KXmM+O0RQV79WkX2tPxD9B75ZtZLeUSxKDOSpk00okt29GxXfWgTBbpolFaTI5xv
# 5XwaPNBgej6Y6eoZdy/grgTGcUi14OOFMQz4CNbTDzEFkslQdizjZaxuYrvE9rHo
# t00/49ZDVg==
# SIG # End signature block
