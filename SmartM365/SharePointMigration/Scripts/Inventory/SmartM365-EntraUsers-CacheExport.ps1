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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBwba7jYbChwkor
# i3kHqLUIuhjCdSf5VDH7Vfltlmd56KCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBHxvUn+gu0WLA5vxmC
# 2PChDvgiWsUDRcpcwA8fuD7cwjANBgkqhkiG9w0BAQEFAASCAYCmxbdWAM1CibWT
# 1ts++fS5svOOXf8HOlw2T/qEwj5ZEcmCHUkuaAiFDpOpvPC0JQslsAVJKHdPEwc/
# /pIFNPQLT7NJMRToVltMcBbWs3cfcG0qEyXydFuPhIA2VwK7Xg46SbJ2dtapp/7k
# qT3KrXulM0Cef98CoYzyoUh1TlqajbCzDZJmcK09foLsHc+q3ct4eFiIDZ5F0cc2
# GQ3A94ic5ZgRU2HJDQ5jBy9ADikeHe4Z9OeIqH1kmIAWNQpQG5XdrbavySY8Tg8e
# uGIcwJjvS8bjnfPguYTt37ZJE55EtzUaL5w+VikLSBHLV9fEuZgT7U18RFTO7E9e
# F/5CB3Byutrm46ZI1VnAfmwe3SLs62g3Q2UQO3E5MVi+T0q/dN6y9hO/gqTmMJ4f
# Mq+4lDjKSAGvWlYPFUWdGbys3ZfCOOsjuWeWV4/wtYtDZI3yLAFmD5Vi6pIiqfAd
# 5neh5eZ7xC/WqhdV58SV01YD5Lg7c4x1KOTNNRdgpcclDVisMT8=
# SIG # End signature block
