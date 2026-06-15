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
