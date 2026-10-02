<#
.SYNOPSIS
    Disables SharePoint page comments on a SharePoint Online site collection.

.DESCRIPTION
    Uses Microsoft.Online.SharePoint.PowerShell and Set-SPOSite.
    This script must be run with Windows PowerShell 5.1 through powershell.exe.

    Set-SPOSite -CommentsOnSitePagesDisabled applies the setting at site collection level.
    SharePoint Online also applies it to the subsites in that site collection.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantAdminUrl,

    [ValidateNotNullOrEmpty()]
    [string]$OutputName = "SharePointSite",

    [switch]$WhatIfMode
)

$ErrorActionPreference = "Stop"

function ConvertTo-SafeFileName {
    param([string]$Value)

    $safe = [regex]::Replace($Value, '[^\w\.-]+', '-')
    $safe = $safe.Trim('-')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return "SharePointSite"
    }

    return $safe
}

$ScriptName = "Disable-{0}-SharePointPageComments" -f (ConvertTo-SafeFileName -Value $OutputName)
$ScriptVersion = "3.0.0"
$runGuid = if ($env:SPMIG_GUI_RUN_ID -match '^[0-9a-fA-F]{32}$') {
    $env:SPMIG_GUI_RUN_ID.ToLowerInvariant()
} else { [guid]::NewGuid().ToString('N') }
$RunId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $runGuid

$ScriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
    Split-Path -Parent $PSCommandPath
}
elseif (-not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    (Get-Location).Path
}

$LogFolder = Join-Path -Path $ScriptRoot -ChildPath "logs"
$OutputFolder = Join-Path -Path $ScriptRoot -ChildPath "Output"
$LogFile = Join-Path -Path $LogFolder -ChildPath "$ScriptName-v$ScriptVersion-$RunId.log"
$CsvFile = Join-Path -Path $OutputFolder -ChildPath "$ScriptName-v$ScriptVersion-$RunId.csv"

function Write-RunLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    $color = switch ($Level) {
        "SUCCESS" { "Green" }
        "WARN" { "Yellow" }
        "ERROR" { "Red" }
        default { "Gray" }
    }

    Write-Host $line -ForegroundColor $color
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Import-SPOAdminModule {
    $module = Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell" |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $module) {
        throw "Required module 'Microsoft.Online.SharePoint.PowerShell' is not installed. Install it in Windows PowerShell 5.1."
    }

    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
    Write-RunLog "SPO module version: $($module.Version)"
    Write-RunLog "SPO module path: $($module.Path)"
}

function New-ResultRow {
    return [ordered]@{
        RunId                               = $RunId
        ScriptVersion                       = $ScriptVersion
        Timestamp                           = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        WindowsUser                         = "$env:USERDOMAIN\$env:USERNAME"
        SiteUrl                             = $SiteUrl
        TenantAdminUrl                      = $TenantAdminUrl
        PreviousCommentsOnSitePagesDisabled = $null
        NewCommentsOnSitePagesDisabled      = $null
        Status                              = "Pending"
        Message                             = $null
    }
}

try {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null

    Write-RunLog "Starting $ScriptName v$ScriptVersion. RunId: $RunId"
    Write-RunLog "ComputerName: $env:COMPUTERNAME"
    Write-RunLog "Windows user: $env:USERDOMAIN\$env:USERNAME"
    Write-RunLog "PowerShell edition: $($PSVersionTable.PSEdition)"
    Write-RunLog "PowerShell version: $($PSVersionTable.PSVersion)"
    Write-RunLog "Site URL: $SiteUrl"
    Write-RunLog "Tenant admin URL: $TenantAdminUrl"
    Write-RunLog "WhatIfMode: $WhatIfMode"
    Write-RunLog "Log file: $LogFile"
    Write-RunLog "CSV file: $CsvFile"

    if ($PSVersionTable.PSEdition -ne "Desktop") {
        throw "This script must be run with Windows PowerShell 5.1. Use powershell.exe, not pwsh.exe."
    }

    Import-SPOAdminModule

    $result = New-ResultRow

    Write-RunLog "Connecting to SharePoint admin center: $TenantAdminUrl"
    Connect-SPOService -Url $TenantAdminUrl

    $siteBefore = Get-SPOSite -Identity $SiteUrl
    $result.PreviousCommentsOnSitePagesDisabled = $siteBefore.CommentsOnSitePagesDisabled
    Write-RunLog "Current CommentsOnSitePagesDisabled: $($siteBefore.CommentsOnSitePagesDisabled)"

    if ($siteBefore.CommentsOnSitePagesDisabled -eq $true) {
        $result.NewCommentsOnSitePagesDisabled = $true
        $result.Status = "AlreadyDisabled"
        $result.Message = "Comments were already disabled on the site collection."
        Write-RunLog $result.Message "SUCCESS"
    }
    elseif ($WhatIfMode) {
        $result.NewCommentsOnSitePagesDisabled = $siteBefore.CommentsOnSitePagesDisabled
        $result.Status = "WhatIf"
        $result.Message = "No change applied. Comments would be disabled with Set-SPOSite."
        Write-RunLog $result.Message "WARN"
    }
    else {
        Write-RunLog "Disabling comments with Set-SPOSite."
        Set-SPOSite -Identity $SiteUrl -CommentsOnSitePagesDisabled $true

        $siteAfter = Get-SPOSite -Identity $SiteUrl
        $result.NewCommentsOnSitePagesDisabled = $siteAfter.CommentsOnSitePagesDisabled
        $result.Status = "Success"
        $result.Message = "Comments disabled on the site collection."
        Write-RunLog "New CommentsOnSitePagesDisabled: $($siteAfter.CommentsOnSitePagesDisabled)" "SUCCESS"
    }

    [pscustomobject]$result | Export-Csv -LiteralPath $CsvFile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-RunLog "CSV report exported to: $CsvFile" "SUCCESS"
    Write-RunLog "Completed $ScriptName v$ScriptVersion." "SUCCESS"
}
catch {
    if ($null -eq $result) {
        $result = New-ResultRow
    }

    $result.Status = "Error"
    $result.Message = $_.Exception.Message
    [pscustomobject]$result | Export-Csv -LiteralPath $CsvFile -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    Write-RunLog $_.Exception.Message "ERROR"
    throw
}
finally {
    try {
        Disconnect-SPOService -ErrorAction SilentlyContinue
    }
    catch {
        $null = $_
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB80s4hyvjNg3G1
# d3gv4EzMTCHKcneJO2NTdNh8ch2/B6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB12ThZ+KVsxyGqfTU91hTV
# xTxczcmOFuCN8/A0U2/suzANBgkqhkiG9w0BAQEFAASCAYAwJYbuo4TgQWv8SCB9
# gSfqIOWL8H5fh9Ab7q0A5+BFqRNdiTfPWvebXS8fk+shDCL7DUL3Mv8I8Gh6QejL
# Piqwp5dOaWAqapoNtJYnBSO3vBR9CiN5WGhi2GKrZMPbJMG4gPD6/BubysDZiM9+
# rGY8XEEFMpHB+s61zY9dRJ0h2ujb1rV/vr5o59fgvlipS5cywTyK7wnAwfFqen+U
# N3H6PLT2q7RCp5GKv3dW1rdS2BM6KPIo+KlTwimEGSTHBsstGaHRvP0Q1PciciMe
# 8LabaDI2dt4guIoicYBBRpuf9gX+InFrcre5pJnUnouD3ocNFzEWARIpjeR88W7Z
# NcaKuSNYIjSI8KQ7uYREJMD34I4NfytNUhNxhyXHeQHx+ERjz1LgTO2gtny5IaVW
# OJzxH/kLENbRrLMmjVSZ6JZ6osuFbH1P2glKrNKkzgcEU1dNoRRhemgnaW8OysCh
# ULtSYeMh0MmXmebh2Oe13rBHPpYuN9Rp4fnX0IEptK2Bujw=
# SIG # End signature block
