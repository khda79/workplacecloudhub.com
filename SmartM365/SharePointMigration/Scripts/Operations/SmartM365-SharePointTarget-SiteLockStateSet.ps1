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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBISgnu9wM09J4Z
# 6QJwOxDfLrCGJvITfvhAt87f1Nbw36CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDILqMrl2t6gSobvfV7
# cI/Rm5tGrNyVPfV5cVdDOpqHIjANBgkqhkiG9w0BAQEFAASCAYBplEFj8jkFTEBQ
# KKERUOVOpynmUyp+aq+GXi4L1NYcy3RbxYDbRf+J/4Qci85X3G+oqrOIPW1pM6nw
# tiQKhs94oPvg6Vl0YI6Lt1ptMwJjNqzNy/KUib/8/JFH1/LYx2zkEvr18IQS7QKu
# 904VdTHVBOcMdxF6WoE7aSg8WGIzElhLOh3DB4wC2c3cQKdUATYdadvFOsnOceOh
# TViOaM36mJ2GmMZTe6lfMT27U293uFwEj8BXiT/7djsiiNSO1zIioC9SfMVAnZg1
# 4TpgtJvp8TgyWWfFC4F0JPXXfsL4Ku4jGj3XJVqmZmzMoMuNwQSBlOoB18ckXv3m
# 2bjw4taxge+bC7NrAoF8tD4x/TeaXqmOaIXwO3m0HETMCVIAAIiS5f6OWz5K7ybf
# fCfRrIKZaXGg7qH0fiMkoukAo64qb8yxSfoacTHNvs+JCX9rHnhwhSLMpsoB0b6A
# 9vN5DPReQ0BYOX0paDWtdATZDnMI+qbro7OoPf4/NqWiKojbxmE=
# SIG # End signature block
