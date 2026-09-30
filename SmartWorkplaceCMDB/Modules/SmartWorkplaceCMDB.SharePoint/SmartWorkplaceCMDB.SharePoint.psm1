# SmartWorkplaceCMDB.SharePoint
# Version: 0.1.3

$script:SmartWorkplaceCMDBSharePointVersion = '0.1.3'
$script:DriveIdCache = @{}
$script:FolderCache = @{}

function ConvertTo-SmartWorkplaceCMDBGraphDrivePath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    return ((($Path -replace '\\', '/').Trim('/') -split '/') |
        ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
}

function Get-SmartWorkplaceCMDBSharePointRelativePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [Parameter(Mandatory)][string]$DataAllRootPath,
        [Parameter(Mandatory)][string]$LatestOutputRootPath,
        [Parameter(Mandatory)][string]$LogRootPath
    )

    $file = [IO.Path]::GetFullPath($LocalFilePath)
    $roots = [ordered]@{
        'DATA-ALL'  = [IO.Path]::GetFullPath($DataAllRootPath)
        'DATA-LAST' = [IO.Path]::GetFullPath($LatestOutputRootPath)
        'LOG-ALL'   = [IO.Path]::GetFullPath($LogRootPath)
    }

    foreach ($entry in $roots.GetEnumerator()) {
        $root = $entry.Value.TrimEnd('\', '/')
        $prefix = $root + [IO.Path]::DirectorySeparatorChar
        if ($file.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
            $relative = $file.Substring($prefix.Length)
            return ('{0}/{1}' -f $entry.Key, ($relative -replace '\\', '/'))
        }
    }

    throw "Local file is outside the configured SmartWorkplaceCMDB data roots: '$LocalFilePath'."
}

function Get-SmartWorkplaceCMDBSharePointContentType {
    [CmdletBinding()]
    param([Parameter(Mandatory)][IO.FileInfo]$FileInfo)

    if ($FileInfo.Extension -ieq '.csv') {
        return 'text/csv'
    }
    if ($FileInfo.Extension -ieq '.html') {
        return 'text/html; charset=utf-8'
    }
    if ($FileInfo.Extension -ieq '.log' -or
        $FileInfo.Name.EndsWith(
            '.transcript.txt',
            [StringComparison]::OrdinalIgnoreCase
        )) {
        return 'text/plain; charset=utf-8'
    }
    if ($FileInfo.Name.EndsWith(
            '.status.json.txt',
            [StringComparison]::OrdinalIgnoreCase
        ) -or $FileInfo.Name.EndsWith(
            '.manifest.json.txt',
            [StringComparison]::OrdinalIgnoreCase
        )) {
        return 'application/json; charset=utf-8'
    }
    throw 'Only CSV, HTML, LOG, .transcript.txt, .status.json.txt, and .manifest.json.txt files can be published by this publisher.'
}

function Test-SmartWorkplaceCMDBSharePointConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$CertificateThumbprint,
        [Parameter(Mandatory)][string]$SiteHostname,
        [Parameter(Mandatory)][string]$SitePath,
        [Parameter(Mandatory)][string]$LibraryDisplayName,
        [Parameter(Mandatory)][string]$TargetFolderPath
    )

    $tenantGuid = [guid]::Empty
    if (-not [guid]::TryParse($TenantId, [ref]$tenantGuid)) {
        throw 'MicrosoftGraph.TenantId must contain an Entra tenant GUID.'
    }
    $clientGuid = [guid]::Empty
    if (-not [guid]::TryParse($ClientId, [ref]$clientGuid)) {
        throw 'MicrosoftGraph.ClientId must contain an application registration GUID.'
    }
    if ($CertificateThumbprint -notmatch '^[a-fA-F0-9]{40,64}$') {
        throw 'MicrosoftGraph.CertificateThumbprint must contain 40 to 64 hexadecimal characters.'
    }
    if ($SiteHostname -notmatch '^[a-zA-Z0-9.-]+\.sharepoint\.com$') {
        throw 'SharePoint.SiteHostname must contain a SharePoint Online hostname.'
    }
    if (-not $SitePath.StartsWith('/')) {
        throw 'SharePoint.SitePath must start with a forward slash.'
    }
    $normalizedTarget = ($TargetFolderPath -replace '\\', '/').Trim('/')
    if ([string]::IsNullOrWhiteSpace($normalizedTarget) -or
        $normalizedTarget -split '/' -contains '..') {
        throw 'SharePoint.TargetFolderPath must contain a safe relative folder path.'
    }
    if ([string]::IsNullOrWhiteSpace($LibraryDisplayName)) {
        throw 'SharePoint.LibraryDisplayName is required.'
    }

    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if ($null -eq $module) {
        throw 'Microsoft.Graph.Authentication is required for SharePoint publication.'
    }

    $certificate = @('Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') |
        ForEach-Object {
            Get-ChildItem -LiteralPath $_ -ErrorAction SilentlyContinue
        } |
        Where-Object Thumbprint -eq $CertificateThumbprint |
        Select-Object -First 1
    if ($null -eq $certificate) {
        throw 'The configured Microsoft Graph certificate was not found.'
    }
    if (-not $certificate.HasPrivateKey) {
        throw 'The configured Microsoft Graph certificate has no accessible private key.'
    }
    if ($certificate.NotAfter -le (Get-Date)) {
        throw 'The configured Microsoft Graph certificate is expired.'
    }

    return [pscustomobject]@{
        TenantId = $tenantGuid.ToString()
        ClientId = $clientGuid.ToString()
        CertificateThumbprint = $certificate.Thumbprint
        SiteHostname = $SiteHostname.Trim()
        SitePath = '/' + $SitePath.Trim('/')
        LibraryDisplayName = $LibraryDisplayName.Trim()
        TargetFolderPath = $normalizedTarget
        AuthenticationModuleVersion = $module.Version.ToString()
    }
}

function Get-SmartWorkplaceCMDBGraphStatusCode {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    try {
        if ($ErrorRecord.Exception.Response) {
            return [int]$ErrorRecord.Exception.Response.StatusCode
        }
    }
    catch {
        Write-Verbose 'Unable to read the Microsoft Graph HTTP status code.'
    }
    return 0
}

function Invoke-SmartWorkplaceCMDBGraphRequestWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST', 'PUT')]
        [string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [AllowNull()]$Body,
        [string]$ContentType = 'application/json',
        [string]$InputFilePath,
        [int]$MaximumAttempts = 4,
        [string]$Operation = 'Microsoft Graph request'
    )

    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        try {
            $parameters = @{
                Method = $Method
                Uri = $Uri
                ErrorAction = 'Stop'
            }
            if ($null -ne $Body) {
                $parameters['Body'] = $Body
            }
            if (-not [string]::IsNullOrWhiteSpace($ContentType)) {
                $parameters['ContentType'] = $ContentType
            }
            if (-not [string]::IsNullOrWhiteSpace($InputFilePath)) {
                $parameters['InputFilePath'] = $InputFilePath
            }
            return Invoke-MgGraphRequest @parameters
        }
        catch {
            $statusCode = Get-SmartWorkplaceCMDBGraphStatusCode -ErrorRecord $_
            $transient = $statusCode -in @(409, 429, 500, 502, 503, 504) -or
                $_.Exception.Message -match '(?i)throttl|timeout|temporarily unavailable'
            if (-not $transient -or $attempt -ge $MaximumAttempts) {
                throw "$Operation failed. Status=$statusCode; $($_.Exception.Message)"
            }
            $delay = [math]::Min(60, 5 * $attempt)
            Write-Warning "$Operation transient failure. Status=$statusCode; retrying in ${delay}s."
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-SmartWorkplaceCMDBSharePointDriveId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Configuration
    )

    $cacheKey = '{0}|{1}|{2}' -f
        $Configuration.SiteHostname,
        $Configuration.SitePath,
        $Configuration.LibraryDisplayName
    if ($script:DriveIdCache.ContainsKey($cacheKey)) {
        return [string]$script:DriveIdCache[$cacheKey]
    }

    $site = Invoke-SmartWorkplaceCMDBGraphRequestWithRetry `
        -Method GET `
        -Uri ('https://graph.microsoft.com/v1.0/sites/{0}:{1}' -f
            $Configuration.SiteHostname,
            $Configuration.SitePath) `
        -Operation 'Resolve SharePoint site'
    $drives = Invoke-SmartWorkplaceCMDBGraphRequestWithRetry `
        -Method GET `
        -Uri ('https://graph.microsoft.com/v1.0/sites/{0}/drives' -f $site.id) `
        -Operation 'Resolve SharePoint document libraries'

    $normalize = {
        param($Text)
        if ($null -eq $Text) { return '' }
        return (([string]$Text).Normalize(
                [Text.NormalizationForm]::FormD
            ) -replace '\p{M}', '')
    }
    $expected = & $normalize $Configuration.LibraryDisplayName
    $drive = @($drives.value | Where-Object {
            (& $normalize $_.name) -ieq $expected
        } | Select-Object -First 1)[0]
    if ($null -eq $drive) {
        $available = @($drives.value | ForEach-Object name) -join ', '
        throw "SharePoint library '$($Configuration.LibraryDisplayName)' was not found. Available: $available"
    }

    $script:DriveIdCache[$cacheKey] = [string]$drive.id
    return [string]$drive.id
}

function Confirm-SmartWorkplaceCMDBSharePointFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DriveId,
        [Parameter(Mandatory)][string]$FolderPath
    )

    $normalized = ($FolderPath -replace '\\', '/').Trim('/')
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return
    }

    $parent = ''
    foreach ($segment in $normalized -split '/') {
        $current = if ([string]::IsNullOrWhiteSpace($parent)) {
            $segment
        }
        else {
            "$parent/$segment"
        }
        $cacheKey = "$DriveId|$current"
        if ($script:FolderCache.ContainsKey($cacheKey)) {
            $parent = $current
            continue
        }

        $encodedCurrent = ConvertTo-SmartWorkplaceCMDBGraphDrivePath -Path $current
        $exists = $true
        try {
            Invoke-SmartWorkplaceCMDBGraphRequestWithRetry `
                -Method GET `
                -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encodedCurrent" `
                -Operation "Resolve SharePoint folder '$current'" | Out-Null
        }
        catch {
            if ($_.Exception.Message -notmatch 'Status=404') {
                throw
            }
            $exists = $false
        }

        if (-not $exists) {
            $body = @{
                name = $segment
                folder = @{}
                '@microsoft.graph.conflictBehavior' = 'fail'
            } | ConvertTo-Json -Depth 5
            $uri = if ([string]::IsNullOrWhiteSpace($parent)) {
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
            }
            else {
                $encodedParent = ConvertTo-SmartWorkplaceCMDBGraphDrivePath -Path $parent
                "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${encodedParent}:/children"
            }
            Invoke-SmartWorkplaceCMDBGraphRequestWithRetry `
                -Method POST `
                -Uri $uri `
                -Body $body `
                -Operation "Create SharePoint folder '$current'" | Out-Null
        }

        $script:FolderCache[$cacheKey] = $true
        $parent = $current
    }
}

function Invoke-SmartWorkplaceCMDBSharePointLargeFileUpload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [Parameter(Mandatory)][string]$DriveId,
        [Parameter(Mandatory)][string]$EncodedTargetPath,
        [int]$ChunkSizeBytes = 10485760
    )

    $sessionBody = @{
        item = @{
            '@microsoft.graph.conflictBehavior' = 'replace'
        }
    } | ConvertTo-Json -Depth 5
    $session = Invoke-SmartWorkplaceCMDBGraphRequestWithRetry `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${EncodedTargetPath}:/createUploadSession" `
        -Body $sessionBody `
        -Operation 'Create SharePoint upload session'
    if ([string]::IsNullOrWhiteSpace([string]$session.uploadUrl)) {
        throw 'SharePoint upload session did not return an upload URL.'
    }

    $chunkMultiple = 327680
    if (($ChunkSizeBytes % $chunkMultiple) -ne 0) {
        $ChunkSizeBytes = [int]([math]::Floor(
                $ChunkSizeBytes / $chunkMultiple
            ) * $chunkMultiple)
    }
    if ($ChunkSizeBytes -lt $chunkMultiple) {
        $ChunkSizeBytes = $chunkMultiple
    }

    $stream = [IO.File]::OpenRead($LocalFilePath)
    $lastResponse = $null
    try {
        $length = [int64]$stream.Length
        $buffer = New-Object byte[] $ChunkSizeBytes
        $offset = [int64]0
        while ($offset -lt $length) {
            $remaining = $length - $offset
            $requested = [int][math]::Min($ChunkSizeBytes, $remaining)
            $read = $stream.Read($buffer, 0, $requested)
            if ($read -le 0) {
                break
            }
            $chunk = New-Object byte[] $read
            [Array]::Copy($buffer, 0, $chunk, 0, $read)
            $end = $offset + $read - 1

            $uploaded = $false
            for ($attempt = 1; -not $uploaded -and $attempt -le 4; $attempt++) {
                try {
                    $response = Invoke-WebRequest `
                        -Method PUT `
                        -Uri ([string]$session.uploadUrl) `
                        -Headers @{
                            'Content-Range' = "bytes $offset-$end/$length"
                        } `
                        -ContentType 'application/octet-stream' `
                        -Body $chunk `
                        -SkipHttpErrorCheck `
                        -ErrorAction Stop
                    if ([int]$response.StatusCode -notin @(200, 201, 202)) {
                        throw "Chunk upload returned HTTP $($response.StatusCode)."
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$response.Content)) {
                        try {
                            $lastResponse = $response.Content |
                                ConvertFrom-Json -ErrorAction Stop
                        }
                        catch {
                            $lastResponse = $response.Content
                        }
                    }
                    $uploaded = $true
                }
                catch {
                    if ($attempt -ge 4) {
                        throw "SharePoint chunk upload failed for bytes $offset-$end/$length. $($_.Exception.Message)"
                    }
                    Start-Sleep -Seconds ([math]::Min(60, 5 * $attempt))
                }
            }
            $offset += $read
        }
    }
    finally {
        $stream.Dispose()
    }
    return $lastResponse
}
function Publish-SmartWorkplaceCMDBSharePointFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$LocalFilePath,
        [Parameter(Mandatory)][string]$DataAllRootPath,
        [Parameter(Mandatory)][string]$LatestOutputRootPath,
        [Parameter(Mandatory)][string]$LogRootPath,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$CertificateThumbprint,
        [Parameter(Mandatory)][string]$SiteHostname,
        [Parameter(Mandatory)][string]$SitePath,
        [Parameter(Mandatory)][string]$LibraryDisplayName,
        [Parameter(Mandatory)][string]$TargetFolderPath
    )

    $configuration = Test-SmartWorkplaceCMDBSharePointConfiguration `
        -TenantId $TenantId `
        -ClientId $ClientId `
        -CertificateThumbprint $CertificateThumbprint `
        -SiteHostname $SiteHostname `
        -SitePath $SitePath `
        -LibraryDisplayName $LibraryDisplayName `
        -TargetFolderPath $TargetFolderPath

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $records = New-Object System.Collections.Generic.List[object]
    $connected = $false
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-MgGraph `
            -TenantId $configuration.TenantId `
            -ClientId $configuration.ClientId `
            -CertificateThumbprint $configuration.CertificateThumbprint `
            -ContextScope Process `
            -NoWelcome `
            -ErrorAction Stop | Out-Null
        $connected = $true
        $driveId = Get-SmartWorkplaceCMDBSharePointDriveId -Configuration $configuration

        foreach ($path in @($LocalFilePath | Sort-Object -Unique)) {
            $fileInfo = $null
            try {
                $fileInfo = Get-Item -LiteralPath $path -ErrorAction Stop
                $contentType = Get-SmartWorkplaceCMDBSharePointContentType `
                    -FileInfo $fileInfo
                $relativePath = Get-SmartWorkplaceCMDBSharePointRelativePath `
                    -LocalFilePath $fileInfo.FullName `
                    -DataAllRootPath $DataAllRootPath `
                    -LatestOutputRootPath $LatestOutputRootPath `
                    -LogRootPath $LogRootPath
                $sharePointPath = '{0}/{1}' -f
                    $configuration.TargetFolderPath.TrimEnd('/'),
                    $relativePath.TrimStart('/')
                $folderPath = Split-Path ($sharePointPath -replace '/', '\') -Parent
                Confirm-SmartWorkplaceCMDBSharePointFolder `
                    -DriveId $driveId `
                    -FolderPath ($folderPath -replace '\\', '/')
                $encodedPath = ConvertTo-SmartWorkplaceCMDBGraphDrivePath -Path $sharePointPath
                $uploaded = if ($fileInfo.Length -gt 250MB) {
                    Invoke-SmartWorkplaceCMDBSharePointLargeFileUpload `
                        -LocalFilePath $fileInfo.FullName `
                        -DriveId $driveId `
                        -EncodedTargetPath $encodedPath
                }
                else {
                    Invoke-SmartWorkplaceCMDBGraphRequestWithRetry `
                        -Method PUT `
                        -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/root:/${encodedPath}:/content" `
                        -InputFilePath $fileInfo.FullName `
                        -ContentType $contentType `
                        -Operation "Upload '$relativePath'"
                }

                $records.Add([pscustomobject][ordered]@{
                    LocalFilePath = $fileInfo.FullName
                    RelativePath = $relativePath
                    SharePointPath = $sharePointPath
                    Status = 'Uploaded'
                    WebUrl = [string]$uploaded.webUrl
                    Size = $fileInfo.Length
                    Error = ''
                })
            }
            catch {
                $records.Add([pscustomobject][ordered]@{
                    LocalFilePath = if ($fileInfo) { $fileInfo.FullName } else { $path }
                    RelativePath = ''
                    SharePointPath = ''
                    Status = 'Failed'
                    WebUrl = ''
                    Size = if ($fileInfo) { $fileInfo.Length } else { 0 }
                    Error = $_.Exception.Message
                })
            }
        }
    }
    finally {
        if ($connected) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }

    return @($records.ToArray())
}

Export-ModuleMember -Function @(
    'Get-SmartWorkplaceCMDBSharePointRelativePath',
    'Test-SmartWorkplaceCMDBSharePointConfiguration',
    'Publish-SmartWorkplaceCMDBSharePointFile'
)

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAr5HTb1pigTwxs
# Gdr1tNVonXoAHNpZazQuRNvcnvjJHqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC1Zdatagu503kr2qVqC24m
# nOtp2ts2hCxoPnkE0QglYzANBgkqhkiG9w0BAQEFAASCAYBOmVqiY/1RH8nOTh4S
# K6ScgRFuCl4IQM2Tyf1xB3Ql8NhHtcjfp/d01v4cDgCUPjHjHWsqD3k6m7m7mz5C
# pLanp4eNqHvRO0HBITypvKGg9GrBFqTHCqCg1LSFnTcxewaHGjwwy3s4gdfzyRoE
# wy+06sBkP6tSlMz2qELIAgyvD9fqCfruExAfWsDMMJMM4ElM+c3yhmNm9057FtmT
# m1Ox83bOKIe6PCPAzdxkgGZRWsyzMd/0K2J4IyDSRSbG7iPUkOJvTfxC9mlAyw6/
# 3iEVAUb7d3vksqVMf3owpQg7ADlYf22aHiq/d3WFRQq+vxMrnBhcahkK/jbBV/AA
# 8Ub76G4RZUYp+msVozw26ElBdDjjLTVYSrKKSo4HKsL6KS89VU9vOTcbgXT1oNEn
# yNUsY3YmWnbJdO/rR2uK+a7WIgnBO/fT/D6JSSpGQOEdeA9YGs27MgG2EdLAPqWp
# 5QeUrqqaNrDTxC3PTDi2RsFtr7amVeHH50fG7H8vZ2Mj/yQ=
# SIG # End signature block
