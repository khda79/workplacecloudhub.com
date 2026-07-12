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
$RunId = Get-Date -Format "yyyyMMdd-HHmmss"

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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC5Brtz3K4CPkOc
# Q442C2YS/r0RERawmZZaDKsNcUzuZ6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC0bViQiFCW01P6dwuP
# NVN4SmY1sT9SIZWBq0aHfQ7E+jANBgkqhkiG9w0BAQEFAASCAYAHQBSyg92KeOW7
# Me+HcuVUQVdSmCmH6JwnbAByG3/nxSHWKhas9sLOOM55vLK3UQsWbRHrwjhK0UR3
# kh2KrQIi0Ln1Xar/WBhr5GcaJj0/JHwsWe8HJCjmMol6TV9w5GlEBVao7vSc3AKB
# ZqaNedS1HShcSF7A1GNKr2U1PVtg7NkUazMZxhui026M/TjnJ/yLb2Y6dxhH0OVI
# /PZ6PqUFgOpCWlchK+bFutPVnZ3LSSD5K5moaE/9HC+UsZfjOqkJ9pLmTmyTJcCQ
# wfUhLHl5n/7VPTmTREbDgxuR6YfeB7yeb2s0EYSmx6o80vPtKsx9sDKNpx8WZNc/
# ushhryYwGQE57R2vOYj5p2q5q679hwQp9VA+rfvl/ElnlZzIr2Y0tknnHM9Ios4U
# UKtM1qyESgMNA8M/dhE70HtbDxjqxGk+qjhKbsd4XeNj3+VufSqwzkTlChWOPkRU
# vmhnqHwA1Stg8+FIdkjRqqr9G3vJ53KvS4RshbaUX/zgIVf3wjU=
# SIG # End signature block
