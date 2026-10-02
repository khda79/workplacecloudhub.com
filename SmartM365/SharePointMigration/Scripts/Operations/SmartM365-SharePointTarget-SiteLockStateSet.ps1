<#
.SYNOPSIS
    Sets the lock state of a SharePoint Online site collection.

.DESCRIPTION
    Uses Microsoft.Online.SharePoint.PowerShell and Set-SPOSite.
    This script must be run with Windows PowerShell 5.1 through powershell.exe.

    The default lock state is ReadOnly. Use -WhatIfMode to generate a report
    without applying changes.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantAdminUrl,

    [ValidateSet("ReadOnly", "Unlock", "NoAccess")]
    [string]$LockState = "ReadOnly",

    [switch]$WhatIfMode
)

$ErrorActionPreference = "Stop"

$ScriptName = "SmartM365-SharePointTarget-SiteLockStateSet"
$ScriptVersion = "1.0.0"
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
        RunId             = $RunId
        ScriptVersion     = $ScriptVersion
        Timestamp         = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        WindowsUser       = "$env:USERDOMAIN\$env:USERNAME"
        ComputerName      = $env:COMPUTERNAME
        SiteUrl           = $SiteUrl
        TenantAdminUrl    = $TenantAdminUrl
        RequestedLockState = $LockState
        PreviousLockState = $null
        NewLockState      = $null
        Status            = "Pending"
        Message           = $null
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
    Write-RunLog "Requested lock state: $LockState"
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
    $result.PreviousLockState = $siteBefore.LockState
    Write-RunLog "Current LockState: $($siteBefore.LockState)"

    if ($siteBefore.LockState -eq $LockState) {
        $result.NewLockState = $siteBefore.LockState
        $result.Status = "AlreadySet"
        $result.Message = "Site collection lock state is already '$LockState'."
        Write-RunLog $result.Message "SUCCESS"
    }
    elseif ($WhatIfMode) {
        $result.NewLockState = $siteBefore.LockState
        $result.Status = "WhatIf"
        $result.Message = "No change applied. Site collection lock state would be set to '$LockState'."
        Write-RunLog $result.Message "WARN"
    }
    else {
        Write-RunLog "Setting site collection LockState to '$LockState'."
        Set-SPOSite -Identity $SiteUrl -LockState $LockState

        $siteAfter = Get-SPOSite -Identity $SiteUrl
        $result.NewLockState = $siteAfter.LockState
        $result.Status = "Success"
        $result.Message = "Site collection lock state set to '$($siteAfter.LockState)'."
        Write-RunLog "New LockState: $($siteAfter.LockState)" "SUCCESS"
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCACMq2UcjpyRBp
# V+7wDxrlCbqyOYWTz5xAEgmTeLqjE6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDj1I+JZ0jCM1UHSxQI1hsA
# Vw+28ByXWquvkOi7Sg9u+DANBgkqhkiG9w0BAQEFAASCAYCHlHgB45NL4oYhOMKO
# kwYQm/w8MplgNEqafr2A0gV2uS8Qu/02rm9QzKPpM9EINapFl1rsJMMyQJS3cMv9
# uXV7Jzzb7oiSNtjxfpMAddnPNMUfvQC494jUNBCw+7nZGxAnCgdb0USFlRDrcZCW
# E15TG/n2QQFzmKIh1XQdlzOWkEC78bzzwb7LZ6KEzIR/kuZUyfo8lcFRaMVH6qF6
# s2r9TLL4G6njmtAwycZHMOrQz63JrwP6TAMNPdmxAXcLL/y/qTaApvzB4OTzdz3R
# vKJd3x7e3wFiTywNtzaTu8OpboIJcjBz0QV7w7orSJ8OStvcvEmDLVC0uclBCjp5
# fpMZITrVheptvJgfuBnr9VcQNNKA6FIpEMzHNcsSD671wFea/XYk5MuRoj70kYHz
# +FquPzIkSkVKvhS4rj91cM5vu3vcbEEYeL7abMLxOTuC9xeAYdoMC+XENxV9/YY/
# tyKW0qhpgpBN6FduFt9mz27gRoPcrLHQpUI9XRK5kdinpVU=
# SIG # End signature block
