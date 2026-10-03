<#
.SYNOPSIS
    Verifies SharePoint URLs and creates one local migration from the template.
.DESCRIPTION
    The migration directory is created only after every configured web and the
    SharePoint Online admin endpoint have passed authenticated read checks.
.VERSION
    1.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RequestPath,

    [Parameter(Mandatory = $true)]
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = if ($MyInvocation.InvocationName -eq '.') { [pscustomobject]@{ OwnsConsole=$false } } else { Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath }
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {
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
    $template = Join-Path $projectRoot 'Migrations_Template'
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
        $script:ConsoleLifecycleStatus = 'FAILED'
        exit 1
    }
    $script:ConsoleLifecycleStatus = $(if (($exitCode) -eq 0) { 'SUCCESS' } else { 'FAILED' })
    exit $exitCode
}
}
catch {
    $script:ConsoleLifecycleFailure = $_
    throw
}
finally {
    Complete-SmartM365MigrationConsoleLifecycle -Context $script:ConsoleLifecycleContext -Failure $script:ConsoleLifecycleFailure -Status $script:ConsoleLifecycleStatus
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqDpzSTuzQP2v5
# HikFCBmTuIXRZWN/dMI0RC9wLVlZg6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEDmICzqkwpnxRh30eAlqNP+6R0X1Emzy24+Oa4Maih8MA0GCSqG
# SIb3DQEBAQUABIIBgHTVYrr+0ZMErCiOBOjedGs2naI0DyMBW4Rl2w5SQSOJP1nC
# xFPkANkbSXUQEnZeXkCDs5Af2yszTBDJ8ocyPgK+6SfEzegGuu5lW9+dlUQmk+4B
# sW0CvxBYLbgpkAaMRamKyyGc0g98WAvOcGDBCH1JIQPyn50DFU5NnrJyo1CgVzom
# oKSNd8m7VWUriSsbhU2nP6FA8PS+xf4KYBZYsOv2ob/81ureAHkyzukQldg8eLoW
# ypT6akhd70wtsQ+FWc4A2xBPBt7o6spY8ZxPShrNOEeyAA3B/2LMUEvWPuhd4WCx
# Y4P2ZqoLQSSy5xGR0IccUnZMvly6/oHQZbYXKUWBPtCgwVUoFVKTRJ9BeVR6sCIF
# S6J47wFdchtNCDBD6xIS82zDAqejOoEPBMV9gZ/tZ51orCMt/D08guPj4Rx9Xhjf
# y9dksk6lDs33j9H6YF0DH5e98dZ0oW0TVU42xllR35IXUltWnsWF75TFWe+w2BgR
# jKZaQmS0eeKdMzOj8qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNDA4
# MzNaMC8GCSqGSIb3DQEJBDEiBCBNQBM+ZrY61b9v7U/ZvzC+OAmq4AZFsB9Aq7lP
# PeRcujANBgkqhkiG9w0BAQEFAASCAgAp/5YIx8LeAPCCikfi/f6MC4pdv4+ilCJs
# Bk/eiAjWbJHBkggCFEZdN69VuhN8epycpmETT40bn1ZgIHVmvSuC2h7B2s+kRAC/
# V+4Y5WLFBpW8aDNOzqgrjTO2oggp3DYAGioCRd+kk9tAzdDlvu0HtqZ4eW2FDtps
# GN92fnuYOHvFmOauuPkIYmpkpd4wtWgxX1WhFNNLSUEdm5Kk61Kbb2w0dyWsH0Lr
# w34KyzPbZUzHXxl6G9EdBoiDmIzglCeOuMuA2GALhJKjMsSePgOMEv63zH3+yjSw
# hvmioRy6k+bK+tPyx/7seqk7LuUChyQBZ2w7BTD1LwCMDcLKuH+hgjwnv1EGQtiS
# 8ZfQb8z5HGNNjEDBAMrDjhfBglIgEC6vs9q8a/yKk0JIbaov/9A8LHWY3XR8VHjb
# AVj5EucT2HMydQndVKrcCyHiWmkTZelFtGlRJQvbTRrV5WQHiJpEottO87cqZGpU
# T7UtkoztuGy2CQDQpC46Vn8W+bcPsCkZTa+sHsRByfZbU54xRABYfYgPjtmn3Due
# LGLRdYvKZnFyWoqFVVfTJ4Qw/7XIEyMRmAenLOPuYD4ZSzUOwNuxQq7FPFRv3ONM
# Dh/CY+zQJfp1z9ay7qpauOYw2YXFpAMv5YMutENXvSOqFqKj4hYalYz6axLjMxSk
# scY9OfiXaQ==
# SIG # End signature block
