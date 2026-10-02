<#
.SYNOPSIS
    Verifies SharePoint URLs and creates one local migration from the template.
.DESCRIPTION
    The migration directory is created only after every configured web and the
    SharePoint Online admin endpoint have passed authenticated read checks.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RequestPath,

    [Parameter(Mandatory = $true)]
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$migrationsRoot = Join-Path $projectRoot 'Migrations'

function Assert-WebUrl {
    param([string]$Value, [string]$Label, [string]$Type)
    $uri = $null
    if (-not [uri]::TryCreate($Value, [System.UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -notin @('http', 'https') -or
        ($Type -eq 'SPO' -and $uri.Scheme -ne 'https') -or
        $uri.UserInfo -or $uri.Query -or $uri.Fragment -or
        $Value -match '\s') {
        throw "$Label must be an absolute SharePoint web URL without credentials, spaces, query or fragment. SPO requires HTTPS."
    }
    return $uri.AbsoluteUri.TrimEnd('/')
}

function Get-CanonicalWebUrl {
    param([string]$Value)
    $uri = [uri]$Value
    return $uri.AbsoluteUri.TrimEnd('/')
}

function Get-CommonPermissionPath {
    param([string[]]$Urls)
    $common = @(([uri]$Urls[0]).AbsolutePath.Trim('/') -split '/' | Where-Object { $_ })
    foreach ($url in $Urls | Select-Object -Skip 1) {
        $segments = @(([uri]$url).AbsolutePath.Trim('/') -split '/' | Where-Object { $_ })
        $limit = [Math]::Min($common.Count, $segments.Count)
        $count = 0
        while ($count -lt $limit -and $common[$count] -ieq $segments[$count]) { $count++ }
        $common = @($common | Select-Object -First $count)
    }
    if ($common.Count -eq 0) { return '/' }
    return '/' + ($common -join '/')
}

function Assert-ExactWeb {
    param([string]$Expected, [string]$Actual, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Actual) -or
        -not [string]::Equals((Get-CanonicalWebUrl $Expected), (Get-CanonicalWebUrl $Actual),
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label returned a different web URL: $Actual (requested $Expected)."
    }
}

function Get-ValidationFailure {
    param([string]$Label, [string]$Url, [System.Exception]$Exception)
    $message = [string]$Exception.Message
    $status = $null
    if ($Exception.Response -and $Exception.Response.StatusCode) {
        $status = [int]$Exception.Response.StatusCode
    }
    if ($status -eq 404 -or $message -match '(?i)\b(NotFound|Not Found|404)\b') {
        return "$Label does not exist or its URL is incorrect: $Url. $message"
    }
    if ($status -in @(401, 403) -or $message -match '(?i)\b(Unauthorized|Forbidden|Access Denied|401|403)\b') {
        return "$Label could not be verified because access was denied: $Url. $message"
    }
    return "$Label could not be verified: $Url. $message"
}

function Test-OnPremWeb {
    param([string]$Url, [string]$Label)
    $apiUrl = '{0}/_api/web?$select=Url' -f $Url.TrimEnd('/')
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -UseDefaultCredentials `
            -Headers @{ Accept = 'application/json;odata=verbose' } -TimeoutSec 30 -ErrorAction Stop
        $actual = if ($response.Url) { [string]$response.Url }
                  elseif ($response.d -and $response.d.Url) { [string]$response.d.Url }
                  else { '' }
        Assert-ExactWeb -Expected $Url -Actual $actual -Label $Label
    }
    catch {
        throw (Get-ValidationFailure -Label $Label -Url $Url -Exception $_.Exception)
    }
}

function Get-SPOAuthParameters {
    param([string]$Mode)
    $authPath = Join-Path $projectRoot 'Config\SPOAuth.local.psd1'
    $config = if (Test-Path -LiteralPath $authPath -PathType Leaf) {
        Import-PowerShellDataFile -LiteralPath $authPath
    } else { @{} }
    $parameters = @{ ReturnConnection = $true; ErrorAction = 'Stop' }
    if ($Mode -eq 'Certificate') {
        foreach ($key in @('ClientId', 'Tenant', 'Thumbprint')) {
            if ([string]::IsNullOrWhiteSpace([string]$config[$key])) {
                throw "Certificate authentication requires $key in Config\SPOAuth.local.psd1."
            }
        }
        $parameters.ClientId = [string]$config.ClientId
        $parameters.Tenant = if ($config.TenantId) { [string]$config.TenantId } else { [string]$config.Tenant }
        $parameters.Thumbprint = [string]$config.Thumbprint
    }
    elseif ($Mode -eq 'DeviceLogin') {
        if ([string]::IsNullOrWhiteSpace([string]$config.Tenant)) {
            throw 'Device login requires Tenant in Config\SPOAuth.local.psd1.'
        }
        $parameters.DeviceLogin = $true
        $parameters.Tenant = [string]$config.Tenant
        if ($config.ClientId) { $parameters.ClientId = [string]$config.ClientId }
    }
    else {
        $parameters.Interactive = $true
        if ($config.ClientId) { $parameters.ClientId = [string]$config.ClientId }
    }
    return $parameters
}

function Test-SPOWeb {
    param([string]$Url, [string]$Label, [hashtable]$AuthParameters)
    try {
        $connection = Connect-PnPOnline -Url $Url @AuthParameters
        $web = Get-PnPWeb -Connection $connection -Includes Url -ErrorAction Stop
        Assert-ExactWeb -Expected $Url -Actual ([string]$web.Url) -Label $Label
    }
    catch {
        throw (Get-ValidationFailure -Label $Label -Url $Url -Exception $_.Exception)
    }
}

function Test-SPOAdmin {
    param([string]$Url, [hashtable]$AuthParameters)
    try {
        $connection = Connect-PnPOnline -Url $Url @AuthParameters
        $tenant = Get-PnPTenant -Connection $connection -ErrorAction Stop
        if ($null -eq $tenant) { throw 'The tenant admin endpoint returned no tenant data.' }
    }
    catch {
        throw (Get-ValidationFailure -Label 'Target admin URL' -Url $Url -Exception $_.Exception)
    }
}

function ConvertTo-Psd1Value {
    param($Value, [int]$Depth = 0)
    $indent = '    ' * $Depth
    if ($null -eq $Value) { return '$null' }
    if ($Value -is [bool]) { return $(if ($Value) { '$true' } else { '$false' }) }
    if ($Value -is [string]) { return "'" + $Value.Replace("'", "''") + "'" }
    if ($Value -is [System.Collections.IDictionary]) {
        $lines = [System.Collections.Generic.List[string]]::new()
        foreach ($key in $Value.Keys) {
            if ([string]$key -notmatch '^[A-Za-z][A-Za-z0-9]*$') { throw "Unsupported config key: $key" }
            $lines.Add(('{0}    {1} = {2}' -f $indent, $key, (ConvertTo-Psd1Value -Value $Value[$key] -Depth ($Depth + 1))))
        }
        return "@{`r`n" + ($lines -join "`r`n") + "`r`n$indent}"
    }
    if ($Value -is [System.Array]) {
        return '@(' + (($Value | ForEach-Object { ConvertTo-Psd1Value -Value $_ -Depth ($Depth + 1) }) -join ', ') + ')'
    }
    if ($Value -is [System.ValueType]) {
        return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0}', $Value)
    }
    throw "Unsupported configuration value type: $($Value.GetType().FullName)"
}

function Assert-MigrationRequest {
    param($Request)
    $name = [string]$Request.Name
    if ($name -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9._-]{0,62}[A-Za-z0-9])?$' -or
        $name -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$' -or
        $name -in @('_Template', '_Local', 'logs')) {
        throw 'Migration name must be a new Windows-safe folder name (letters, digits, dot, dash or underscore).'
    }
    if ($Request.SourceType -notin @('SP2016', 'SP2019', 'SPO') -or
        $Request.TargetType -notin @('SP2016', 'SP2019', 'SPO')) {
        throw 'Source and target types must be SP2016, SP2019 or SPO.'
    }
    if ($Request.AuthMode -notin @('Interactive', 'DeviceLogin', 'Certificate')) {
        throw 'Invalid SharePoint Online authentication mode.'
    }
    if ($null -eq $Request.Mappings -or @($Request.Mappings).Count -eq 0) {
        throw 'At least one source-to-target mapping is required.'
    }
    $folder = Join-Path $migrationsRoot $name
    if (Test-Path -LiteralPath $folder) { throw "Migration folder already exists: $folder" }
    return $folder
}

function Invoke-NewMigration {
    param($Request)
    $destination = Assert-MigrationRequest -Request $Request
    $sourceType = [string]$Request.SourceType
    $targetType = [string]$Request.TargetType
    $mappings = [System.Collections.Generic.List[object]]::new()
    foreach ($row in @($Request.Mappings)) {
        $source = Assert-WebUrl -Value ([string]$row.SourceUrl) -Label 'Source URL' -Type $sourceType
        $target = Assert-WebUrl -Value ([string]$row.TargetUrl) -Label 'Target URL' -Type $targetType
        $mappings.Add([pscustomobject]@{ SourceUrl = $source; TargetUrl = $target })
    }
    $sourceApp = if ($sourceType -ne 'SPO') {
        $raw = if ($Request.SourceWebApplicationUrl) { [string]$Request.SourceWebApplicationUrl }
               else { ([uri]$mappings[0].SourceUrl).GetLeftPart([System.UriPartial]::Authority) }
        Assert-WebUrl -Value $raw -Label 'Source web application URL' -Type $sourceType
    } else { '' }
    $targetApp = if ($targetType -ne 'SPO') {
        $raw = if ($Request.TargetWebApplicationUrl) { [string]$Request.TargetWebApplicationUrl }
               else { ([uri]$mappings[0].TargetUrl).GetLeftPart([System.UriPartial]::Authority) }
        Assert-WebUrl -Value $raw -Label 'Target web application URL' -Type $targetType
    } else { '' }
    $adminUrl = if ($targetType -eq 'SPO') {
        $url = Assert-WebUrl -Value ([string]$Request.TargetAdminUrl) -Label 'Target admin URL' -Type 'SPO'
        if (([uri]$url).AbsolutePath -ne '/') { throw 'Target admin URL must be the admin host root.' }
        $url
    } else { '' }
    if (($sourceApp -and ([uri]$sourceApp).AbsolutePath -ne '/') -or
        ($targetApp -and ([uri]$targetApp).AbsolutePath -ne '/')) {
        throw 'SharePoint Server web application URLs must be host roots without a site path.'
    }
    foreach ($row in $mappings) {
        if ($sourceApp -and
            ([uri]$row.SourceUrl).GetLeftPart([System.UriPartial]::Authority) -ine
            ([uri]$sourceApp).GetLeftPart([System.UriPartial]::Authority)) {
            throw "Source mapping is outside the selected web application: $($row.SourceUrl)"
        }
        if ($targetApp -and
            ([uri]$row.TargetUrl).GetLeftPart([System.UriPartial]::Authority) -ine
            ([uri]$targetApp).GetLeftPart([System.UriPartial]::Authority)) {
            throw "Target mapping is outside the selected web application: $($row.TargetUrl)"
        }
    }
    $sourceRoot = if ($Request.SourcePermissionRootPath) { [string]$Request.SourcePermissionRootPath }
                  else { Get-CommonPermissionPath -Urls @($mappings | ForEach-Object SourceUrl) }
    $targetRoot = if ($Request.TargetPermissionRootPath) { [string]$Request.TargetPermissionRootPath }
                  else { Get-CommonPermissionPath -Urls @($mappings | ForEach-Object TargetUrl) }
    foreach ($root in @($sourceRoot, $targetRoot)) {
        if ($root -notmatch '^/' -or [regex]::IsMatch($root, "[`r`n']")) {
            throw 'Permission root paths must start with / and contain no quotes or line breaks.'
        }
    }

    $auth = $null
    if ($sourceType -eq 'SPO' -or $targetType -eq 'SPO') {
        Import-Module PnP.PowerShell -ErrorAction Stop
        $auth = Get-SPOAuthParameters -Mode ([string]$Request.AuthMode)
    }
    $checked = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $mappings) {
        if ($checked.Add("$sourceType|$($row.SourceUrl)")) {
            if ($sourceType -eq 'SPO') { Test-SPOWeb -Url $row.SourceUrl -Label 'Source web' -AuthParameters $auth }
            else { Test-OnPremWeb -Url $row.SourceUrl -Label 'Source web' }
        }
        if ($checked.Add("$targetType|$($row.TargetUrl)")) {
            if ($targetType -eq 'SPO') { Test-SPOWeb -Url $row.TargetUrl -Label 'Target web' -AuthParameters $auth }
            else { Test-OnPremWeb -Url $row.TargetUrl -Label 'Target web' }
        }
    }
    if ($adminUrl) { Test-SPOAdmin -Url $adminUrl -AuthParameters $auth }

    if (Test-Path -LiteralPath $destination) { throw "Migration folder appeared during validation: $destination" }
    $template = Join-Path $migrationsRoot '_Template'
    if (-not (Test-Path -LiteralPath $template -PathType Container)) { throw "Template not found: $template" }
    $staging = Join-Path $migrationsRoot ('.new-{0}-{1}' -f $Request.Name, [guid]::NewGuid().ToString('N'))
    try {
        Copy-Item -LiteralPath $template -Destination $staging -Recurse -ErrorAction Stop
        $config = Import-PowerShellDataFile -LiteralPath (Join-Path $template 'migration.config.psd1')
        $config.Name = [string]$Request.Name
        $config.Source.Type = $sourceType
        $config.Source.WebApplicationUrl = $sourceApp
        $config.Source.SiteUrl = if ($sourceType -eq 'SPO') { $mappings[0].SourceUrl } else { '' }
        $config.Source.PermissionRootPath = $sourceRoot
        $config.Target.Type = $targetType
        $config.Target.WebApplicationUrl = $targetApp
        $config.Target.SiteUrl = if ($targetType -eq 'SPO') { $mappings[0].TargetUrl } else { '' }
        $config.Target.TenantAdminUrl = $adminUrl
        $config.Target.PermissionRootPath = $targetRoot
        $config.Target.PrefixToRemove = ''
        $utf8 = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText((Join-Path $staging 'migration.config.psd1'),
            ((ConvertTo-Psd1Value -Value $config) + "`r`n"), $utf8)
        $mappingText = (($mappings | ForEach-Object { '{0} {1}' -f $_.SourceUrl, $_.TargetUrl }) -join "`r`n") + "`r`n"
        [System.IO.File]::WriteAllText((Join-Path $staging 'migration.mapping.txt'), $mappingText, $utf8)
        $saved = Import-PowerShellDataFile -LiteralPath (Join-Path $staging 'migration.config.psd1')
        if ($saved.Name -ne $Request.Name -or $saved.Source.Type -ne $sourceType -or $saved.Target.Type -ne $targetType) {
            throw 'Generated migration configuration did not pass read-back validation.'
        }
        [System.IO.Directory]::Move($staging, $destination)
    }
    finally {
        if (Test-Path -LiteralPath $staging -PathType Container) {
            $rootFull = [System.IO.Path]::GetFullPath($migrationsRoot).TrimEnd('\') + '\'
            $stageFull = [System.IO.Path]::GetFullPath($staging)
            if (-not $stageFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Refusing to clean an unexpected staging path: $stageFull"
            }
            Remove-Item -LiteralPath $stageFull -Recurse -Force
        }
    }
    return [pscustomobject]@{ Success = $true; MigrationName = [string]$Request.Name; Folder = $destination; ValidatedUrls = ($checked.Count + [int][bool]$adminUrl) }
}

if ($MyInvocation.InvocationName -ne '.') {
    try {
        $request = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $result = Invoke-NewMigration -Request $request
        $exitCode = 0
    }
    catch {
        $result = [pscustomobject]@{ Success = $false; Message = $_.Exception.Message }
        $exitCode = 1
    }
    try {
        $json = $result | ConvertTo-Json -Depth 8
        [System.IO.File]::WriteAllText($ResultPath, $json, [System.Text.UTF8Encoding]::new($false))
    }
    catch {
        Write-Error "Could not write migration result: $($_.Exception.Message)"
        exit 1
    }
    exit $exitCode
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD6fi3AgqjIi6w+
# x0oQVZj15kSj7YasA1eKTvpCHDRlGaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB56SN+6eEsDlElDXNGHbLD
# SVTjOQDmll4sEf5bFmaHejANBgkqhkiG9w0BAQEFAASCAYBe2aXjGcKoyxpktVcR
# j3ZCEz2XbjT+aaucrB5pzZFKVpwH380Rwjo7IH70hlE/tsAGscbbbk6zcb94J+rz
# Nq4DQGLoO5GlvMP7jUpHQNUE8iUwxboInu8AEVQZnccRryNsq/KeQV+qgRNgx5+f
# PGdb+bN8xZ8ttheuRdCnFvAWLK7J5hNHVcJ6DPaGXt2Ig9lb9oE+G190dhJG2PQg
# itEHId2JEWOz0yfszbY3avL0P+FTh9iCtDgrmZWLx3I29xrh7BgTovhmZOgtGcCi
# tMfO0jg+1gGz8wjvIa4MzqCwHvep9e7lJJEZvfQGOl20UeiLzCH8zipTG43vr4fd
# cGZ0WnQnYrYOzLdpSLUZetcxWsoMK3W8EwsZ4Hzgpz0jRtpBmnhMrVkjycj59pLZ
# r840K8Yi1BOqAYh9ALLRBCz0dEBcg8fwPedAigDWlmVVQFe7CrsUBCbgcAFxNG4Q
# dmW1gaXNnrho+Q2FO3vbieLFbkbaMmdZJX2QSSyFnMigjcE=
# SIG # End signature block
