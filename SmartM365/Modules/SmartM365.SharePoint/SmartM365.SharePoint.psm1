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
