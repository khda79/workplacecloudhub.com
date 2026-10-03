<#
.SYNOPSIS
    Exports a lightweight Microsoft Entra users cache for SharePoint migration comparisons.

.DESCRIPTION
    Reuses a fresh cache when it is younger than MaxCacheAgeHours. Otherwise connects to
    Microsoft Graph and exports identity fields used to normalize SharePoint permission
    principals during source/target comparisons.

.VERSION
    1.0.3
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [double]$MaxCacheAgeHours = 6,

    [switch]$Connect,

    [switch]$InteractiveAuth,

    [switch]$DeviceLogin,

    [switch]$ForceRefresh,

    [string]$TenantId,

    [string]$AppId,

    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {

function Write-CacheInfo {
    param([string]$Message, [ConsoleColor]$Color = [ConsoleColor]::Gray)
    Write-Host ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) -ForegroundColor $Color
}

function Test-GraphConnection {
    try {
        Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id&$top=1' -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Ensure-GraphUsersModule {
    $required = @('Microsoft.Graph.Users', 'Microsoft.Graph.Authentication')
    $missing = @($required | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
    if ($missing.Count -gt 0) {
        $account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        Write-CacheInfo ("Missing Microsoft Graph modules for {0} in PowerShell {1}: {2}" -f $account, $PSVersionTable.PSVersion, ($missing -join ', ')) Yellow
        try { $answer = Read-Host 'Install only the required modules from PSGallery for the current user? [Y/N]' }
        catch { throw "Module installation could not be confirmed interactively: $($_.Exception.Message)" }
        if ([string]$answer -notmatch '^(?i:y|yes|o|oui)$') {
            throw 'Required Microsoft Graph module installation was declined; the Entra users cache cannot be refreshed.'
        }
        if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) {
            throw 'Install-Module is unavailable in this PowerShell session; install PowerShellGet for the current user.'
        }
        try {
            if ('Microsoft.Graph.Users' -in $missing) {
                Write-CacheInfo 'Installing Microsoft.Graph.Users for CurrentUser (Authentication is a dependency)...' Cyan
                Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
            }
            if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
                Write-CacheInfo 'Installing Microsoft.Graph.Authentication for CurrentUser...' Cyan
                Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
            }
        }
        catch { throw "Required Microsoft Graph module installation failed for ${account}: $($_.Exception.Message)" }
    }
    foreach ($name in $required) {
        try { Import-Module -Name $name -ErrorAction Stop }
        catch { throw "Required module $name could not be loaded: $($_.Exception.Message)" }
    }
    foreach ($command in @('Get-MgUser', 'Get-MgContext', 'Connect-MgGraph', 'Invoke-MgGraphRequest')) {
        if (-not (Get-Command -Name $command -ErrorAction SilentlyContinue)) {
            throw "Required Microsoft Graph command is unavailable after module import: $command"
        }
    }
    $versions = @($required | ForEach-Object { $module = Get-Module -Name $_; '{0}={1}' -f $_, $module.Version })
    Write-CacheInfo ("Microsoft Graph modules ready: {0}" -f ($versions -join '; ')) DarkCyan
}

function Enter-EntraCacheLock {
    param([string]$Path)
    $deadline = (Get-Date).AddMinutes(10)
    $announced = $false
    while ($true) {
        try {
            return [System.IO.File]::Open($Path, [System.IO.FileMode]::OpenOrCreate,
                [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] {
            if ((Get-Date) -ge $deadline) { throw "Timed out waiting for Entra cache lock: $Path" }
            if (-not $announced) {
                Write-CacheInfo "Another user is refreshing the Entra cache; waiting for $Path" Yellow
                $announced = $true
            }
            Start-Sleep -Seconds 2
        }
    }
}

function Connect-EntraUsersGraph {
    $context = $null
    try { $context = Get-MgContext -ErrorAction SilentlyContinue } catch { }

    if (-not $Connect -and $context -and (Test-GraphConnection)) {
        Write-CacheInfo 'Existing Microsoft Graph session detected. Reusing current connection.' DarkCyan
        return
    }

    if ($Connect -and (Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }

    $connectParams = @{ NoWelcome = $true; ErrorAction = 'Stop' }
    if (-not $InteractiveAuth -and -not [string]::IsNullOrWhiteSpace($TenantId) -and -not [string]::IsNullOrWhiteSpace($AppId) -and -not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        $connectParams.TenantId = $TenantId
        $connectParams.ClientId = $AppId
        $connectParams.CertificateThumbprint = $CertificateThumbprint
        Write-CacheInfo 'Connecting to Microsoft Graph with app-only certificate authentication.' DarkCyan
    }
    else {
        $connectParams.Scopes = @('User.Read.All', 'Directory.Read.All')
        if ($DeviceLogin) {
            $connectParams.UseDeviceCode = $true
        }
        Write-CacheInfo 'Connecting to Microsoft Graph with delegated authentication.' DarkCyan
    }

    Connect-MgGraph @connectParams | Out-Null
    if (-not (Test-GraphConnection)) {
        throw 'Microsoft Graph connection validation failed.'
    }
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

if ($MaxCacheAgeHours -le 0) { throw 'MaxCacheAgeHours must be greater than zero.' }
$MaxCacheAgeHours = [math]::Min($MaxCacheAgeHours, 6)
$cacheLock = Enter-EntraCacheLock -Path ("{0}.lock.txt" -f $OutputPath)
$tempPath = $null
try {
if (-not $ForceRefresh -and (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    $cacheItem = Get-Item -LiteralPath $OutputPath -ErrorAction Stop
    $cacheAgeHours = ((Get-Date) - $cacheItem.LastWriteTime).TotalHours
    $cacheHeader = if ($cacheItem.Length -gt 0) { Get-Content -LiteralPath $OutputPath -TotalCount 1 -ErrorAction Stop } else { '' }
    if ($cacheAgeHours -ge 0 -and $cacheAgeHours -lt $MaxCacheAgeHours -and $cacheHeader -match '^"?Id"?,') {
        Write-CacheInfo ("Using existing Entra users cache: {0} (age {1:n2}h, max {2:n2}h)" -f $OutputPath, $cacheAgeHours, $MaxCacheAgeHours) Green
        return
    }

    Write-CacheInfo ("Entra users cache is stale or invalid: {0} (age {1:n2}h, max {2:n2}h)" -f $OutputPath, $cacheAgeHours, $MaxCacheAgeHours) Yellow
}
elseif ($ForceRefresh) {
    Write-CacheInfo "Entra users cache refresh was requested for $OutputPath; Microsoft Graph export is required." Yellow
}
else {
    Write-CacheInfo "No reusable Entra users cache at $OutputPath; Microsoft Graph export is required." Yellow
}

Ensure-GraphUsersModule
Connect-EntraUsersGraph

$selectProperties = @(
    'id',
    'displayName',
    'userPrincipalName',
    'mail',
    'proxyAddresses',
    'otherMails',
    'onPremisesUserPrincipalName',
    'onPremisesSamAccountName',
    'onPremisesSecurityIdentifier',
    'accountEnabled',
    'userType'
)

Write-CacheInfo 'Exporting Microsoft Entra users cache...' Cyan
$users = Get-MgUser -All -Property $selectProperties -ErrorAction Stop
if (@($users).Count -eq 0) { throw 'Microsoft Graph returned no users; the Entra cache was not published.' }
$rows = foreach ($user in $users) {
    [pscustomobject]@{
        Id                          = $user.Id
        DisplayName                 = $user.DisplayName
        UserPrincipalName           = $user.UserPrincipalName
        Mail                        = $user.Mail
        ProxyAddresses              = (($user.ProxyAddresses | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join ';')
        OtherMails                  = (($user.OtherMails | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join ';')
        OnPremisesUserPrincipalName = $user.OnPremisesUserPrincipalName
        OnPremisesSamAccountName    = $user.OnPremisesSamAccountName
        OnPremisesSecurityIdentifier = $user.OnPremisesSecurityIdentifier
        AccountEnabled              = $user.AccountEnabled
        UserType                    = $user.UserType
        ExportedAt                  = (Get-Date).ToString('s')
    }
}

$tempPath = "{0}.{1}.tmp" -f $OutputPath, [guid]::NewGuid().ToString('N')
$rows | Export-Csv -LiteralPath $tempPath -NoTypeInformation -Encoding UTF8
Move-Item -LiteralPath $tempPath -Destination $OutputPath -Force
Write-CacheInfo ("Entra users cache exported: {0} ({1} users)" -f $OutputPath, @($rows).Count) Green
$tempPath = $null
}
finally {
    if ($tempPath -and (Test-Path -LiteralPath $tempPath)) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
    $cacheLock.Dispose()
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD+l6r/uGkjN3GA
# Yn56NIxBKbEqc/WnewgNDLRmYZHL8aCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAQmPDAU7ADKgX1D82EyuBa
# BuQ8GDkx2esxkBBJWgnKSDANBgkqhkiG9w0BAQEFAASCAYCGP8TDcSPbNXe44QAH
# D/s4QAxFZk6ycdScJ4IdZhNcdEpm4ytuxZh1Lezn91ZU6ooZ0i06mVcPv7/2MG9L
# Os56INNWnA1R26qjyD3BWz2ZETD4SoV1Mz/UYugKuUbF69R8WFD6JF5B4/G/jCBm
# nxPr7tBiqyiwE4T1hjcFX90yLnnVp8woqQr65WtlxJien6PljI7Uv8qi+y9GaHyS
# BM+1oEydAhpaWzvudPdufitKBdzFd37xLUC78rCauQP7drZu3am5ZTepkk9fCbLO
# +SHoMp6600CY40Cf0sT+CpVykHrTfxU/PLGRIb1oX2s46Ip3VNZfNUXRgmwhzhcJ
# awGhOWekzwnV60qiWN/x0oGg1zBEPb9KDugKW3BiwiQEW3KJZW3sBJNVuUEazfs5
# me9z6x00kyw37Ouf7TNa2S+tBGj6q+lDmQCXSFS5H5Efq5WGm+evu3u8nuV0rXG3
# G1ipa7U2lyMPCDiyBiVpu37lQHyvXR4UYNLapWWwiYG4UFs=
# SIG # End signature block
