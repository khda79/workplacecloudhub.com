#Requires -Version 7.0

<#
.SYNOPSIS
    SharePoint Upload Module.
    Provides Microsoft Graph REST-based SharePoint Online file upload
    for use in any PowerShell script that holds a Graph Bearer token.

.DESCRIPTION
    Exposes two public functions:

        Resolve-SmartM365SpDriveId
            Resolves the SharePoint document library drive ID from site coordinates.
            Cache the result and pass it to Invoke-SmartM365SpFileUpload to avoid
            re-resolving on every call when uploading multiple files.

        Invoke-SmartM365SpFileUpload
            Uploads a single local file to a SharePoint Online document library.
            - Skips upload if the remote file is already up to date (timestamp comparison).
            - Files <= ChunkSize : direct PUT upload.
            - Files >  ChunkSize : resumable chunked upload with per-chunk retry.
            - Preserves source file timestamps via PATCH after upload (non-blocking).
            - Accepts an optional Logger scriptblock to bridge with the caller's
              Write-Log function.
            - Returns a structured [pscustomobject] result; never throws.

    AUTHENTICATION
        Uses Invoke-RestMethod with a pre-acquired Graph Bearer token passed as a
        hashtable: @{ Authorization = "Bearer <token>" }
        The caller is responsible for token acquisition and refresh.
        Required Graph permission: Sites.Selected with site-level write grant

    LOGGER INTERFACE
        The optional -Logger scriptblock receives three positional arguments:
            param($Message, $Level, $Stage)
        where Level is "INFO", "WARN", or "ERROR" and Stage is "SP".
        Compatible with the Write-Log function signature.

        Example:
            $spLogger = { param($msg, $lvl, $stg) Write-Log $msg $lvl $stg }

    USAGE EXAMPLE
        Import-Module ".\Modules\SmartM365.SharePoint\SmartM365.SharePoint.psd1" -ErrorAction Stop

        $spDriveId = Resolve-SmartM365SpDriveId `
            -Headers             $headers `
            -SiteHostname        "contoso.sharepoint.com" `
            -SitePath            "/sites/workplace-data" `
            -LibraryDisplayName  "Documents" `
            -Logger              { param($m,$l,$s) Write-Log $m $l $s }

        $spResult = Invoke-SmartM365SpFileUpload `
            -Headers           $headers `
            -LocalFilePath     $CsvFinal `
            -TargetFolderPath  "SmartM365-DATA/CSV" `
            -DriveId           $spDriveId `
            -Logger            { param($m,$l,$s) Write-Log $m $l $s }

        if ($spResult.Status -eq "Error") {
            Write-Log "SharePoint upload failed: $($spResult.Error)" "WARN" "SP"
        }

VERSION
    1.0.0
#>

Set-StrictMode -Version Latest

# ==========================================================
# Internal: Write-SpLog
# ==========================================================
function Write-SpLog {
    param(
        [string]$Message,
        [string]$Level  = "INFO",
        [string]$Stage  = "SP",
        [scriptblock]$Logger = $null
    )
    if ($Logger) {
        try { & $Logger $Message $Level $Stage } catch { }
    }
    else {
        switch ($Level) {
            "WARN"  { Write-Warning  "[$Stage] $Message" }
            "ERROR" { Write-Warning  "[$Stage][ERROR] $Message" }
            default { Write-Verbose  "[$Stage] $Message" }
        }
    }
}

# ==========================================================
# Public: Resolve-SmartM365SpDriveId
# ==========================================================
function Resolve-SmartM365SpDriveId {
    <#
    .SYNOPSIS
        Resolves the drive ID for a SharePoint Online document library.
    .PARAMETER Headers
        Hashtable containing the Graph Bearer token: @{ Authorization = "Bearer <token>" }
    .PARAMETER SiteHostname
        SharePoint hostname, e.g. "contoso.sharepoint.com"
    .PARAMETER SitePath
        Site path, e.g. "/sites/workplace-data"
    .PARAMETER LibraryDisplayName
        Display name of the document library, e.g. "Documents"
    .PARAMETER Logger
        Optional scriptblock bridging to the caller's logging function.
        Signature: param($Message, $Level, $Stage)
    .OUTPUTS
        [string] Drive ID
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$SiteHostname,
        [Parameter(Mandatory)][string]$SitePath,
        [Parameter(Mandatory)][string]$LibraryDisplayName,
        [scriptblock]$Logger = $null
    )

    Write-SpLog "Resolving site: $SiteHostname$SitePath" "INFO" "SP" -Logger $Logger

    $siteUri    = "https://graph.microsoft.com/v1.0/sites/$($SiteHostname):$($SitePath)"
    $site       = Invoke-RestMethod -Method GET -Uri $siteUri -Headers $Headers
    $siteId     = $site.id
    Write-SpLog "Site ID resolved: $siteId" "INFO" "SP" -Logger $Logger

    $drivesUri  = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
    $drivesResp = Invoke-RestMethod -Method GET -Uri $drivesUri -Headers $Headers
    $driveList  = @($drivesResp.value)

    # Match by display name, with accent-normalized fallback
    $normalize  = { param($s) $s.Normalize([System.Text.NormalizationForm]::FormD) -replace '\p{M}', '' }
    $drive      = $driveList | Where-Object { $_.name -ieq $LibraryDisplayName }
    if (-not $drive) {
        $targetNorm = & $normalize $LibraryDisplayName
        $drive      = $driveList | Where-Object { (& $normalize $_.name) -ieq $targetNorm }
    }
    if (-not $drive) {
        $available = ($driveList | ForEach-Object { $_.name }) -join " | "
        throw "SharePoint drive '$LibraryDisplayName' not found on site '$SiteHostname$SitePath'. Available drives: $available"
    }

    Write-SpLog "Drive ID resolved: $($drive.id) (library: '$($drive.name)')" "INFO" "SP" -Logger $Logger
    return $drive.id
}

# ==========================================================
# Internal: Get-SpRemoteLastModified
# ==========================================================
function Get-SpRemoteLastModified {
    param(
        [hashtable]$Headers,
        [string]$DriveId,
        [string]$DestRelativePath
    )
    try {
        $uri  = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($DestRelativePath)?`$select=lastModifiedDateTime,id"
        $item = Invoke-RestMethod -Method GET -Uri $uri -Headers $Headers -ErrorAction Stop
        $raw  = $item.lastModifiedDateTime
        if ($raw) {
            # Invoke-RestMethod may auto-deserialize ISO 8601 strings to [datetime].
            # Calling .ToString() without a format on a [datetime] uses the current culture
            # (fr-FR: dd/MM/yyyy), which swaps day/month when re-parsed with InvariantCulture.
            # Fix: if already a [datetime], use it directly; otherwise parse the raw string as-is.
            $dt = if ($raw -is [datetime]) {
                $raw
            }
            else {
                [datetime]::Parse(
                    [string]$raw,
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::RoundtripKind
                )
            }
            if ($dt.Kind -eq [System.DateTimeKind]::Utc)         { return $dt.ToLocalTime() }
            if ($dt.Kind -eq [System.DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($dt, [System.DateTimeKind]::Utc).ToLocalTime() }
            return $dt
        }
    }
    catch { }
    return $null
}

# ==========================================================
# Internal: Invoke-SpSmallFileUpload
# ==========================================================
function Invoke-SpSmallFileUpload {
    param(
        [hashtable]$Headers,
        [string]$DriveId,
        [string]$DestRelativePath,
        [string]$LocalFilePath
    )
    $uri           = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($DestRelativePath):/content"
    $bytes         = [System.IO.File]::ReadAllBytes($LocalFilePath)
    $putHeaders    = $Headers.Clone()
    $putHeaders["Content-Type"] = "application/octet-stream"
    return Invoke-RestMethod -Method PUT -Uri $uri -Headers $putHeaders -Body $bytes
}

# ==========================================================
# Internal: Invoke-SpLargeFileUpload
# ==========================================================
function Invoke-SpLargeFileUpload {
    param(
        [hashtable]$Headers,
        [string]$DriveId,
        [string]$DestRelativePath,
        [string]$LocalFilePath,
        [long]$FileSize,
        [long]$ChunkSize,
        [scriptblock]$Logger = $null
    )

    $sessionUri     = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($DestRelativePath):/createUploadSession"
    $sessionBody    = @{ item = @{ "@microsoft.graph.conflictBehavior" = "replace" } } | ConvertTo-Json -Depth 5
    $sessionHeaders = $Headers.Clone()
    $sessionHeaders["Content-Type"] = "application/json"
    $maxRetries     = 3
    $retryDelay     = 10
    $session        = $null

    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            $session = Invoke-RestMethod -Method POST -Uri $sessionUri -Headers $sessionHeaders -Body $sessionBody
            break
        }
        catch {
            $errMsg = $_.Exception.Message
            if ($errMsg -match "nameAlreadyExists|Conflict" -and $attempt -lt $maxRetries) {
                Write-SpLog "Upload session conflict (attempt $attempt/$maxRetries) — retrying in ${retryDelay}s..." "WARN" "SP" -Logger $Logger
                Start-Sleep -Seconds $retryDelay
            }
            else { throw }
        }
    }

    $uploadUrl  = $session.uploadUrl
    $stream     = [System.IO.File]::OpenRead($LocalFilePath)
    $actualSize = $stream.Length
    if ($actualSize -ne $FileSize) {
        Write-SpLog "File size changed since scan: declared=$FileSize actual=$actualSize — using actual." "WARN" "SP" -Logger $Logger
        $FileSize = $actualSize
    }
    $buffer = [byte[]]::new($ChunkSize)
    $offset = [long]0
    $result = $null

    Write-SpLog "Large file upload started: $([math]::Round($FileSize / 1MB, 2)) MB — resumable session." "INFO" "SP" -Logger $Logger

    try {
        while ($offset -lt $FileSize) {
            $read  = $stream.Read($buffer, 0, $ChunkSize)
            if ($read -eq 0) { break }
            $chunk = $buffer[0..($read - 1)]
            $end   = $offset + $read - 1

            $chunkHeaders = @{
                "Content-Range" = "bytes $offset-$end/$FileSize"
                "Content-Type"  = "application/octet-stream"
            }
            $chunkMaxRetries = 3
            $chunkRetryDelay = 5

            for ($chunkAttempt = 1; $chunkAttempt -le $chunkMaxRetries; $chunkAttempt++) {
                try {
                    $result = Invoke-RestMethod -Method PUT -Uri $uploadUrl -Headers $chunkHeaders -Body ([byte[]]$chunk)
                    break
                }
                catch {
                    $chunkErr    = $_.Exception.Message
                    $isLastChunk = ($offset + $read) -ge $FileSize
                    if ($isLastChunk -and $chunkErr -match "400|Bad Request") {
                        Write-SpLog "Last chunk returned 400 — treating as success (SharePoint accepted bytes)." "WARN" "SP" -Logger $Logger
                        break
                    }
                    if ($chunkAttempt -lt $chunkMaxRetries) {
                        Write-SpLog "Chunk error at offset $offset (attempt $chunkAttempt/$chunkMaxRetries): $chunkErr — retrying in ${chunkRetryDelay}s..." "WARN" "SP" -Logger $Logger
                        Start-Sleep -Seconds $chunkRetryDelay
                    }
                    else { throw }
                }
            }
            $offset += $read
            $pct = [math]::Round($offset / $FileSize * 100)
            Write-SpLog "Progress: $pct% ($offset / $FileSize bytes)" "INFO" "SP" -Logger $Logger
        }
    }
    finally {
        $stream.Dispose()
    }

    return $result
}

# ==========================================================
# Internal: Set-SpFileTimestamp
# ==========================================================
function Set-SpFileTimestamp {
    param(
        [hashtable]$Headers,
        [string]$DriveId,
        [string]$DestRelativePath,
        [datetime]$LastModified,
        [datetime]$Created
    )
    $uri          = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($DestRelativePath)"
    $patchHeaders = $Headers.Clone()
    $patchHeaders["Content-Type"] = "application/json"
    $bodyJson = @{
        fileSystemInfo = @{
            lastModifiedDateTime = $LastModified.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            createdDateTime      = $Created.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
    } | ConvertTo-Json -Compress
    Invoke-RestMethod -Method PATCH -Uri $uri -Headers $patchHeaders -Body $bodyJson | Out-Null
}

# ==========================================================
# Public: Invoke-SmartM365SpFileUpload
# ==========================================================
function Invoke-SmartM365SpFileUpload {
    <#
    .SYNOPSIS
        Uploads a single local file to a SharePoint Online document library.
    .DESCRIPTION
        - Requires a pre-acquired Graph Bearer token in -Headers.
        - Skips upload if the remote file is already up to date (timestamp comparison).
        - Files <= ChunkSize : direct PUT upload.
        - Files >  ChunkSize : resumable chunked upload with per-chunk retry.
        - Preserves source timestamps via PATCH after upload (non-blocking on failure).
        - Accepts an optional -Logger scriptblock bridging to the caller's Write-Log.
        - Never throws: errors are captured in the returned result object.
          Check $result.Status -eq "Error" and $result.Error in the caller.
    .PARAMETER Headers
        Hashtable containing the Graph Bearer token: @{ Authorization = "Bearer <token>" }
    .PARAMETER LocalFilePath
        Full path to the local file to upload.
    .PARAMETER TargetFolderPath
        Relative path of the target folder inside the drive root, e.g. "SmartM365-DATA/CSV"
    .PARAMETER DriveId
        Pre-resolved drive ID (recommended for performance). If omitted, SiteHostname,
        SitePath and LibraryDisplayName must be provided for on-the-fly resolution.
    .PARAMETER SiteHostname
        Required only when DriveId is omitted. e.g. "contoso.sharepoint.com"
    .PARAMETER SitePath
        Required only when DriveId is omitted. e.g. "/sites/workplace-data"
    .PARAMETER LibraryDisplayName
        Required only when DriveId is omitted. e.g. "Documents"
    .PARAMETER ChunkSize
        Byte threshold above which the resumable chunked upload is used. Default: 10 MB.
    .PARAMETER Logger
        Optional scriptblock bridging to the caller's logging function.
        Signature: param($Message, $Level, $Stage)
        Example: { param($m, $l, $s) Write-Log $m $l $s }
    .OUTPUTS
        [pscustomobject]@{
            Status      # "Success" | "Skipped" | "Error"
            FileName    # Uploaded file name
            SizeKB      # File size in KB
            DurationMs  # Elapsed time in milliseconds
            DriveId     # Resolved or provided drive ID
            DestPath    # Relative destination path on SharePoint
            Error       # Error message (empty on success)
        }
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$LocalFilePath,
        [Parameter(Mandatory)][string]$TargetFolderPath,
        [string]$DriveId             = "",
        [string]$SiteHostname        = "",
        [string]$SitePath            = "",
        [string]$LibraryDisplayName  = "",
        [long]$ChunkSize             = 10MB,
        [bool]$ForceUpload           = $true,
        [scriptblock]$Logger         = $null
    )

    $result = [pscustomobject]@{
        Status     = "Pending"
        FileName   = ""
        SizeKB     = 0
        DurationMs = 0
        DriveId    = $DriveId
        DestPath   = ""
        Error      = ""
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # Validate local file
        if (-not (Test-Path -LiteralPath $LocalFilePath -PathType Leaf)) {
            throw "Local file not found: $LocalFilePath"
        }
        $fileInfo        = Get-Item -LiteralPath $LocalFilePath
        $result.FileName = $fileInfo.Name
        $result.SizeKB   = [math]::Round($fileInfo.Length / 1KB, 1)

        # Resolve drive ID if not provided
        if ([string]::IsNullOrWhiteSpace($DriveId)) {
            if ([string]::IsNullOrWhiteSpace($SiteHostname) -or
                [string]::IsNullOrWhiteSpace($SitePath)     -or
                [string]::IsNullOrWhiteSpace($LibraryDisplayName)) {
                throw "DriveId not provided. SiteHostname, SitePath and LibraryDisplayName are all required to resolve it automatically."
            }
            $DriveId        = Resolve-SmartM365SpDriveId -Headers $Headers -SiteHostname $SiteHostname -SitePath $SitePath -LibraryDisplayName $LibraryDisplayName -Logger $Logger
            $result.DriveId = $DriveId
        }

        $destPath        = "$TargetFolderPath/$($fileInfo.Name)"
        $result.DestPath = $destPath

        # Skip check (only when ForceUpload = $false)
        if (-not $ForceUpload) {
            $remoteTs = Get-SpRemoteLastModified -Headers $Headers -DriveId $DriveId -DestRelativePath $destPath
            if ($remoteTs -and $fileInfo.LastWriteTime -le $remoteTs) {
                $result.Status = "Skipped"
                Write-SpLog "SKIPPED (not newer): $($fileInfo.Name) | Local=$($fileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')) Remote=$($remoteTs.ToString('yyyy-MM-dd HH:mm:ss'))" "INFO" "SP" -Logger $Logger
                $sw.Stop()
                $result.DurationMs = $sw.ElapsedMilliseconds
                return $result
            }
        }

        Write-SpLog "Uploading: $($fileInfo.Name) ($($result.SizeKB) KB) → $destPath" "INFO" "SP" -Logger $Logger

        if ($fileInfo.Length -gt $ChunkSize) {
            Invoke-SpLargeFileUpload -Headers $Headers -DriveId $DriveId -DestRelativePath $destPath `
                -LocalFilePath $LocalFilePath -FileSize $fileInfo.Length -ChunkSize $ChunkSize -Logger $Logger | Out-Null
        }
        else {
            Invoke-SpSmallFileUpload -Headers $Headers -DriveId $DriveId -DestRelativePath $destPath `
                -LocalFilePath $LocalFilePath | Out-Null
        }

        # Preserve source timestamps (non-blocking on failure)
        try {
            Set-SpFileTimestamp -Headers $Headers -DriveId $DriveId -DestRelativePath $destPath `
                -LastModified $fileInfo.LastWriteTime -Created $fileInfo.CreationTime
        }
        catch {
            Write-SpLog "Timestamp update failed (non-blocking): $($_.Exception.Message)" "WARN" "SP" -Logger $Logger
        }

        $sw.Stop()
        $result.Status     = "Success"
        $result.DurationMs = $sw.ElapsedMilliseconds
        Write-SpLog "SUCCESS: $($fileInfo.Name) uploaded in $($sw.ElapsedMilliseconds) ms" "INFO" "SP" -Logger $Logger
    }
    catch {
        $sw.Stop()
        $result.Status     = "Error"
        $result.DurationMs = $sw.ElapsedMilliseconds
        $result.Error      = $_.Exception.Message
        Write-SpLog "FAILED: $($_.Exception.Message)" "ERROR" "SP" -Logger $Logger
    }

    return $result
}

# ==========================================================
# Module exports
# ==========================================================
Export-ModuleMember -Function @(
    'Resolve-SmartM365SpDriveId',
    'Invoke-SmartM365SpFileUpload'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCpJPBh+t0ULq0N
# Z2mOkgsB1+Gup8UohLikgihVlm7np6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIJ6fXXNc7Kof0tuj4W7j10kCONgqGa/LKDpe5cLA7LrXMA0GCSqG
# SIb3DQEBAQUABIIBgJIx8cHKlCi6aFwDy4aHNV5Dr+nTQLIPgjAKBaFW5wvsSYvn
# 2oB2Ekg7iXl0tpuUietUVSBcfGo4ti3EXzhembMiOrt0aD7E3FOm3mRRyAd5qMnf
# SrZmXBJ9FFGM0qtYTeQ2an+X6RmqNz5Yri4cLi7tMf1iszPxdC6/Qrva55nda5uZ
# ZlmpasAn8OJeaz8g6iMWZzG/xxMKEkSfbKDHOYIywtE1xROHLRKoNRe19/kyh9oi
# hSDfLtDFwTbjExU59KvUIjK+NjU40mX3uhlKHRNODqQ4ymYuxtUHkojm7W0nFhXw
# X0pWgsXSsOxdnaBE4jIqy55rRBb6Bt7ysJDFIID0HZ+bPJYdyKBddtFujZuz38XI
# xphNcqfloj+71JmtGOduxtMdR5EEbSxzbzcgvbA9OszS3ncY4LYRSR5w7BOiilG1
# 4P42Im095sXJaKzkrtPPeR+uj/lmQpoKLwjfvSlRpfkwgWPQ4xFqTmyIi31IOmxV
# pHAo5SKkJUxo4LqwdKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MDlaMC8GCSqGSIb3DQEJBDEiBCDMwqA3V7Nal5hGb0hqh2vns1dWg8wmoawiYBnC
# hoqg0jANBgkqhkiG9w0BAQEFAASCAgAjwhhxI4NITKxqFtL/F8Cjl40TzsnYxRNw
# fJUYAQwkUFHr+3/Xc4sLp6WUMRIG0dJa32+9zpD3nqS6JwwIZNM1r1bPOrkXhbPP
# BsLD2FJkNDYHSjVBbZxbYoWcbWuxmg09voDuWXZqfDYpOW2yXlejDw/yEsnn2iGW
# rtxg1ihkIzfvngKY8j+nrQteDi4vVsnfWEaddCGaA91nE7u3zE/SPCshzWC32qt6
# r/95aLpSD4B2HDRjs6PlCW/yeezdFa0DvbtmRIaMLSc4xKld7DKFO948vsxYC9ke
# p3KqCuBfjuS1vHd7O5iEo9G5tptogns+tBbI3OOscu6xPrQRpITUySWeuYcBZM9Q
# iUgqdR6GMfN/Gz7p6lYbRWOrvJ+KG6KsbkAFx+mkzNwIplUhKI1DCXTYDqpE8vh5
# x8Y2t7aDytSXZCcqcYqbfW/U0+EmmTv7jCaaGmjzl6tPrQXZNMN3UjvrwgYHWv1a
# z+NyJxlqM22F0uBPIGI1hcJyDNvovFccx1hdOAvfAgZqCixoBq2QbD9wWuyJBaiy
# jvF1vYgP0DoqIVTfYJ5P4XKASNDnu9ohQdFW/2OEDOqCDIBX1HwSfmWfL31bTShS
# tlRZjhgJpZrdZu4QX8twGzNZmWd/JvGkcSeh6MhViJSm/u9K3E+5PRoP9d70HfD5
# Axos/gz68A==
# SIG # End signature block
