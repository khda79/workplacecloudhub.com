# SmartWorkplaceCMDB.SharePoint
# Version: 0.1.0

$script:SmartWorkplaceCMDBSharePointVersion = '0.1.0'
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
                if ($fileInfo.Extension -ine '.csv') {
                    throw 'Only CSV files can be published by this publisher.'
                }
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
                        -ContentType 'text/csv' `
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDMpSHVtO8H+rvH
# U5Odoi88c0z759SMgfB7yMQM48NxjqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDrt52AiGtjBOZ7UkthNDDsWE6knm4vERORlHdvikyd2MA0GCSqG
# SIb3DQEBAQUABIIBgKmQApESQjh91w/fG0GYk6Wd8urcwV9O7tXb9lE6rdBdlsSl
# 9LbC3OTpg8BzfBGv3Gwk+4m+d4jqF3aeZG8fJJm1Ku3KCfNRp516CsHc1TB/6lMC
# ihOFiDI9kQc14lX3HcZ3QkHLmeE9jfVJ+WC23XuI283kWaazxbFoFyX/WIfdC81R
# WjLi5vBHUp7ZhKnIW0Hha5W41FmzS3itDyzW0ZjstTdPDL0pHLvTIQkeBll4Y1J+
# IiKNat/SJBcrL10Y9MpMofpWfTOvtfncVOqSY56rB7+9VlZHHclvFTxIcmlzKnw5
# zA29eCarxPYf3G08CshPsL03B8XIF2Y8mTtpYr2MGGedCmrqhn3jhFk9DOZkorIk
# KtzdrJkdF5hF5zsM91MwkOLTwOvmT2REesGaG+6oZgtpPZh7yCIGGjNmGM1wnbmM
# CnK/wXNUk14SdVPbU31ZiPDwWaMvVoQmqM7sJVHeVYpIwxyIOgH8htxV430hrs0R
# zwMQgV90+F1KF54Qi6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMjAy
# MDFaMC8GCSqGSIb3DQEJBDEiBCAIqSNzBCMynxnB/RG2IcwIo/ImJz/YpTMV+xv3
# lsJPLTANBgkqhkiG9w0BAQEFAASCAgBo5hVrrrG8hoYshzn93JahVW+js+shVh83
# UGA3u6sBRR+ncKpq9sMrm9OdXxRbHVes6Q3jsnfTI3Gokoz6on+MbKHenivre89t
# Sb3OuOaKacBmnFoRoMC65+0rb92xK5tLkRq2jcAZiR6UsKdZZXaZOcS2vf7yTu6P
# JpiH7yA+L3TLT4g3rd67VZaFd0UtuTiDk9XWNonGX4AMD8wY+Nb2s9EwfcdupZIJ
# Ba9IOVW1nxrnFcBWdwJ7ZAH8dFBBPajvHgHSjB0mLGC820TB7vSpL9pK5A8Nz2rL
# xLBK3q2jzKId3EWZddFSfvKdZTFYma0uUxyv25OH6qrYbEHhwt0umc/5iV+8xVFL
# cRqlUduFGGD8aJgyT0NZKUflzv3BTU2LZLmd1/XxcRA/Sk0R8tMTbVGfLcCjbixb
# zlz1euFIsnZ7oWUforQcYBZLnhT8swhBMuhq6Gc3ws+Fsb0Zumz+fmsKK17mlxV5
# d3F9kY2vUyEIoIFADBO1p7wpyfd0svZHdLY4/GTVBtzHjBda8k87QLcU7vR6n5zf
# FVx5CHoI32gG+achCnBpfja+S3zqL/6Gg5T+/By/PW77XNiaABIhYBjKFRnbe4Gy
# XSLeJovFjzHSQ5oLzLv65z9ev+fo5PVBcyukHlaKfMzEay0V6R5biRG3TII7B5Cr
# R8E9vANRzg==
# SIG # End signature block
