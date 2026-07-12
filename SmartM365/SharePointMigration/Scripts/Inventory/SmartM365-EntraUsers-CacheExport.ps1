<#
.SYNOPSIS
    Exports a lightweight Microsoft Entra users cache for SharePoint migration comparisons.

.DESCRIPTION
    Reuses a fresh cache when it is younger than MaxCacheAgeHours. Otherwise connects to
    Microsoft Graph and exports identity fields used to normalize SharePoint permission
    principals during source/target comparisons.

.VERSION
    1.0.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [double]$MaxCacheAgeHours = 24,

    [switch]$Connect,

    [switch]$InteractiveAuth,

    [switch]$DeviceLogin,

    [switch]$ForceRefresh,

    [string]$TenantId,

    [string]$AppId,

    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'

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
    if (-not (Get-Command -Name Get-MgUser -ErrorAction SilentlyContinue)) {
        Import-Module Microsoft.Graph.Users -ErrorAction Stop
    }
    if (-not (Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue)) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
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

if (-not $ForceRefresh -and (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
    $cacheItem = Get-Item -LiteralPath $OutputPath -ErrorAction Stop
    $cacheAgeHours = ((Get-Date) - $cacheItem.LastWriteTime).TotalHours
    if ($cacheAgeHours -lt $MaxCacheAgeHours) {
        Write-CacheInfo ("Using existing Entra users cache: {0} (age {1:n2}h, max {2:n2}h)" -f $OutputPath, $cacheAgeHours, $MaxCacheAgeHours) Green
        return
    }

    Write-CacheInfo ("Entra users cache is stale: {0} (age {1:n2}h, max {2:n2}h)" -f $OutputPath, $cacheAgeHours, $MaxCacheAgeHours) Yellow
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

$tempPath = "{0}.tmp" -f $OutputPath
$rows | Export-Csv -LiteralPath $tempPath -NoTypeInformation -Encoding UTF8
Move-Item -LiteralPath $tempPath -Destination $OutputPath -Force
Write-CacheInfo ("Entra users cache exported: {0} ({1} users)" -f $OutputPath, @($rows).Count) Green

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBwba7jYbChwkor
# i3kHqLUIuhjCdSf5VDH7Vfltlmd56KCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCBHxvUn+gu0WLA5vxmC2PChDvgiWsUDRcpcwA8fuD7cwjANBgkqhkiG9w0B
# AQEFAASCAYAO9NUPPFG9foCvQR/E69k6kd6trug/+wzhz8WmTqBbpiGVAentMSRi
# /s6BrxOStOddIn/hp4dlciCRjprnUR4UZ2zOxjMr5ZYjaFe1Xr8Aan7hb2cz1qFv
# B5e8oOvvbE/BS7vu6rtBbRpqYfTB/tU5dinnrHV4dmFdonAdf5GyhEL8KjNvLLfz
# hbBiOX401TNg9RvnPtEB95yLVixSsHoTe6xAwSHWC8tNQL9BndMtboGdyJ6ozl8t
# i/W1Jpk/7sh+NK2ks9RNk/wJxS4Z8jzb4rm5yxINNNv2Y7XNPS2GMqRmjmmOFaxT
# vnqTg3K+btQFirZ/acG66UAKA1X14m6hN80YaTAQjF/viA6CwK5wcKwle1mAuQ0w
# TK5ahfu6efRTH6FsjjYfctoFkQB/nGE8byI5Li/Vl0WYXdCySUmI39YtV+kLehbL
# pWiqOQMxsKLpV5bVziNnIgz3MUH7bIOqev68/S8/ajtBhflbXK9IYJrc6yJcO9js
# VPXmxQf6qmY=
# SIG # End signature block
