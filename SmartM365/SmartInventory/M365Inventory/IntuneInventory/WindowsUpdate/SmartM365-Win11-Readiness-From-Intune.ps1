<#
.SYNOPSIS
SmartM365 - Windows 11 Readiness export from Intune Endpoint Analytics (Work from anywhere) via Microsoft Graph (beta).

.DESCRIPTION
- Queries Graph beta: deviceManagement/userExperienceAnalyticsWorkFromAnywhereMetrics('allDevices')/metricDevices
- Exports all device readiness records with NormalizedDeviceName, UpgradeEligibilityLabel, AzureAdJoinType
- De-duplication by NormalizedDeviceName (audit CSV for duplicates)
- Atomic CSV writes (DATA-ALL, DATA-LAST, Archive)
- Optional SharePoint upload via SmartM365.Core module (same Graph token)
- Token auto-refresh: Graph token refreshed automatically if age > 55 min
- Emails:
    - Fatal error HTML email (on failure)
    - Optional success/summary HTML email with KPI tables

.REQUIREMENTS
- PowerShell 7+
- MSAL.PS module
- App-only certificate authentication for Graph
- SmartM365.Core module (Modules\SmartM365.Core\SmartM365.Core.psd1) [if SharePoint upload enabled]
- Graph app permissions: DeviceManagementManagedDevices.Read.All (validated before CSV export)

PARAMETERS
  -DryRun             : Run all API calls and processing; skip file writes and emails
  -MaxDevices         : Limit number of devices processed (0 = all; useful for testing)
  -SummaryEmailMode   : Always | OnChange | Never (default: Always)
  -EnableSummaryEmail : Toggle success email (default: $true)
  -EnableErrorEmail   : Toggle error email (default: $true)

OUTPUT
Primary:
{{DataAllRootPath}}\Intune\WindowsUpdate\Win11Readiness\Intune_Devices_Win11Readiness.csv
Archive (timestamped, keep 10):
{{DataAllRootPath}}\Intune\WindowsUpdate\Win11Readiness\Archive\Intune_Devices_Win11Readiness_yyyyMMdd_HHmmss.csv
DATA-LAST (fixed name only):
{{LatestCsvFolderPath}}\Intune_Devices_Win11Readiness.csv
Logs (timestamped, keep 10):
{{DataAllRootPath}}\Intune\WindowsUpdate\Win11Readiness\Logs\SmartM365-Win11-Readiness-From-Intune-yyyyMMdd_HHmmss.log

.VERSION
1.4
#>

param(
    [string]$Tenant = 'test',
    [switch]$DryRun,
    [int]$MaxDevices = 0,
    [ValidateSet("Always","OnChange","Never")][string]$SummaryEmailMode = "Always",
    [bool]$EnableSummaryEmail = $true,
    [bool]$EnableErrorEmail = $true,
    [int]$MaxItems = 0
)
if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
    foreach ($smartM365LimitName in @('TopUsers','TopMailboxes','MaxDevices','MaxSites','MaxTeams','MaxApps','MaxPolicies','Limit','MaxPages')) {
        $smartM365LimitVariable = Get-Variable -Name $smartM365LimitName -Scope Script -ErrorAction SilentlyContinue
        if ($smartM365LimitVariable -and -not $PSBoundParameters.ContainsKey($smartM365LimitName) -and $null -ne $smartM365LimitVariable.Value) {
            Set-Variable -Name $smartM365LimitName -Value ([int]$MaxItems) -Scope Script
        }
    }
}

# ==========================================================
# SmartM365 - Version
# ==========================================================
$ScriptVersion = "1.4"

# ==========================================================
# SmartM365 - App-only authentication parameters
# ==========================================================
$AppId     = "00000000-0000-0000-0000-000000000000"
$TenantId  = "00000000-0000-0000-0000-000000000000"
$Thumb     = "0000000000000000000000000000000000000000"
$OrgDomain = "contoso.onmicrosoft.com"

# ==========================================================
# SmartM365 - Output paths (UNC)
# ==========================================================
$BasePath = Join-Path $PSScriptRoot "Output\Win11Readiness"

$CsvName  = "Intune_Devices_Win11Readiness.csv"
$CsvFinal = Join-Path $BasePath $CsvName
$CsvTemp  = Join-Path $BasePath "$CsvName.tmp"

$ArchivePath = Join-Path $BasePath "Archive"
$LogsPath    = Join-Path $BasePath "Logs"

$DataLastPath = Join-Path $PSScriptRoot "Output\DATA-LAST"
$CsvLastFinal = Join-Path $DataLastPath $CsvName
$CsvLastTemp  = Join-Path $DataLastPath "$CsvName.tmp"

# ==========================================================
# SmartM365 - SharePoint upload
# ==========================================================
$EnableSharePointUpload = $true
$SP_SiteHostname        = "contoso.sharepoint.com"
$SP_SitePath            = "/sites/SMART-M365"
$SP_LibraryDisplayName  = "Documents"
$SP_TargetFolderPath    = "SMART-M365/DATA"
[long]$SP_ChunkSize     = 10MB

# ==========================================================
# SmartM365 - Email configuration
# ==========================================================
$SmtpPort          = 25
$From              = "noreply@example.com"
$UseIntegratedAuth = $true

$ToSummary = ""
$BccAll    = ""

$SmtpHostName = ""
$RelayIp      = $null
$UseSsl       = $false

$SummaryStatePath = Join-Path $BasePath "Intune_Devices_Win11Readiness.lastcount.txt"

# ==========================================================
# SmartM365 - Console rendering options
# ==========================================================
$ConsolePretty      = $true
$ConsoleShowRunId   = $false
$ConsoleShowGetUrls = $false

# ==========================================================
# SmartM365 - Run metadata
# ==========================================================
$ScriptName = "SmartM365-Win11-Readiness-From-Intune"
$RunId      = [guid]::NewGuid().ToString()
$RunStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath    = Join-Path $LogsPath "$ScriptName-$RunStamp.log"

# ==========================================================
# SmartM365 - Tenant/config integration
# ==========================================================
function Find-SmartM365Root {
    $d = $PSScriptRoot
    while ($d) {
        $tenantContext = Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1'
        $coreModule = Join-Path -Path $d -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if ((Test-Path -LiteralPath $tenantContext) -and (Test-Path -LiteralPath $coreModule)) { return $d }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365 root not found from script path.'
}
function Get-SmartM365ScriptLocalConfig {
    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ('{0}.local.json' -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
        if (-not (Test-Path -LiteralPath $templatePath)) { throw "Local configuration file not found and template is missing: $configPath" }
        Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
        Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
    }
    return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}
function Resolve-SmartM365ScriptValue {
    param([AllowNull()]$Value)
    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $matches) {
            $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $tokenValue = Resolve-SmartM365ScriptValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}
function Get-SmartM365ScriptConfigValue {
    param([Parameter(Mandatory = $true)]$Config,[Parameter(Mandatory = $true)][string]$Name,$DefaultValue)
    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -isnot [string]) { return Resolve-SmartM365ScriptValue -Value $property.Value }
        $text = $property.Value.Trim()
        if ($text -and $text -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return Resolve-SmartM365ScriptValue -Value $property.Value }
    }
    $globalProperty = $script:SmartM365EffectiveConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) { return $DefaultValue }
        return Resolve-SmartM365ScriptValue -Value $globalProperty.Value
    }
    return $DefaultValue
}
$SmartM365Root = Find-SmartM365Root
. (Join-Path -Path $SmartM365Root -ChildPath 'Config\SmartM365-TenantContext.ps1')
$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
Import-Module (Join-Path -Path $SmartM365Root -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1') -Force -ErrorAction Stop
Initialize-SmartM365DefaultCsvValidationRules
$ScriptLocalConfig = Get-SmartM365ScriptLocalConfig
$AppId = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue ''
$TenantId = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue ''
$Thumb = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'Thumbprint' -DefaultValue (Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '')
$OrgDomain = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
$BasePath = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue (Join-Path (Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'DataAllRootPath' -DefaultValue $PSScriptRoot) 'Intune\WindowsUpdate\Win11Readiness')
$DataLastPath = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue $BasePath
$LogAllRootPath = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue (Join-Path $BasePath 'Logs')
$CsvName = 'Intune_Devices_Win11Readiness.csv'
$CsvFinal = Join-Path $BasePath $CsvName
$CsvTemp = Join-Path $BasePath "$CsvName.tmp"
$ArchivePath = Join-Path $BasePath 'Archive'
$LogsPath = Join-Path $LogAllRootPath $ScriptName
$CsvLastFinal = Join-Path $DataLastPath $CsvName
$CsvLastTemp = Join-Path $DataLastPath "$CsvName.tmp"
$SummaryStatePath = Join-Path $BasePath 'Intune_Devices_Win11Readiness.lastcount.txt'
$RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogPath = Join-Path $LogsPath "$ScriptName-$RunStamp.log"
$SmartM365SharePointUploadEnabled = [bool](Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:EnableSharePointUpload = $SmartM365SharePointUploadEnabled
$global:SharePointSiteHostname = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$global:AppId = $AppId
$global:TenantId = $TenantId
$global:Thumbprint = $Thumb
$EnableSharePointUpload = $false
$From = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue ''
$ToSummary = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue (Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'ErrorMailTo' -DefaultValue '')
$SmtpHostName = Get-SmartM365ScriptConfigValue -Config $ScriptLocalConfig -Name 'SmtpServer' -DefaultValue ''
if ([string]::IsNullOrWhiteSpace($SmtpHostName)) { $EnableSummaryEmail = $false; $EnableErrorEmail = $false }

# ==========================================================
# SmartM365 - Core helpers
# ==========================================================
function Assert-PS7 {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw "PowerShell 7+ is required. Current version: $($PSVersionTable.PSVersion)"
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO",
        [Parameter()][string]$Stage = ""
    )

    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $fileLine = "[$ts][$Level][$ScriptName][$RunId] $Message"
    Add-Content -Path $LogPath -Value $fileLine -Encoding UTF8

    if (-not $ConsolePretty) { Write-Host $fileLine; return }

    $prefix    = if ([string]::IsNullOrWhiteSpace($Stage)) { "" } else { "[$Stage] " }
    $rid       = if ($ConsoleShowRunId) { " ($RunId)" } else { "" }
    $consoleTs = (Get-Date).ToString("HH:mm:ss")
    $line      = "{0} {1}{2}{3}" -f $consoleTs, $prefix, $Message, $rid

    switch ($Level) {
        "INFO"  { Write-Host $line -ForegroundColor Gray }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        "ERROR" { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Cleanup-TempFiles {
    foreach ($p in @($CsvTemp, $CsvLastTemp)) {
        if ($p -and (Test-Path $p)) {
            try { Remove-Item -Path $p -Force -ErrorAction Stop } catch { }
        }
    }
}

function Prune-Files {
    param(
        [Parameter(Mandatory=$true)][string]$Folder,
        [Parameter(Mandatory=$true)][string]$Filter,
        [Parameter(Mandatory=$true)][int]$Keep
    )

    if (-not (Test-Path $Folder)) { return }
    $files = Get-ChildItem -Path $Folder -File -Filter $Filter -ErrorAction SilentlyContinue |
             Sort-Object -Property LastWriteTime -Descending
    if (-not $files) { return }

    $toRemove = $files | Select-Object -Skip $Keep
    foreach ($f in $toRemove) {
        try {
            Remove-Item -Path $f.FullName -Force -ErrorAction Stop
            Write-Log "Pruned old file: $($f.FullName)" "INFO" "PRUNE"
        }
        catch {
            Write-Log "Failed to prune file: $($f.FullName). Error: $($_.Exception.Message)" "WARN" "PRUNE"
        }
    }
}

# ==========================================================
# SmartM365 - Network / SMTP IPv4 resolution
# ==========================================================
function Resolve-IPv4Address {
    param([Parameter(Mandatory=$true)][string]$HostName)

    $addresses = [System.Net.Dns]::GetHostAddresses($HostName) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }
    $first = $addresses | Select-Object -First 1
    if (-not $first) { throw "No IPv4 address found for $HostName" }
    return $first.IPAddressToString
}

# ==========================================================
# SmartM365 - Mail helpers
# ==========================================================
function New-HtmlMailMessage {
    param(
        [Parameter(Mandatory=$true)][string]$From,
        [Parameter(Mandatory=$true)][string]$To,
        [Parameter(Mandatory=$true)][string]$Subject,
        [Parameter(Mandatory=$true)][string]$HtmlBody,
        [Parameter()][string]$BccAll = $null
    )
    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = New-Object System.Net.Mail.MailAddress($From)
    foreach ($addr in ($To -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $mail.To.Add((New-Object System.Net.Mail.MailAddress($addr.Trim())))
    }
    if ($BccAll) {
        foreach ($addr in ($BccAll -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            $mail.Bcc.Add((New-Object System.Net.Mail.MailAddress($addr.Trim())))
        }
    }
    $mail.Subject    = $Subject
    $mail.Body       = $HtmlBody
    $mail.IsBodyHtml = $true
    $utf8 = [System.Text.Encoding]::UTF8
    $mail.SubjectEncoding = $utf8
    $mail.BodyEncoding    = $utf8
    $mail.HeadersEncoding = $utf8
    return $mail
}

function Send-HtmlMail {
    param(
        [Parameter(Mandatory=$true)][System.Net.Mail.MailMessage]$MailMessage,
        [Parameter(Mandatory=$true)][string]$SmtpEndpoint,
        [Parameter(Mandatory=$true)][int]$SmtpPort,
        [Parameter(Mandatory=$true)][bool]$UseIntegratedAuth,
        [Parameter()][bool]$UseSsl = $false
    )

    $client = New-Object System.Net.Mail.SmtpClient($SmtpEndpoint, $SmtpPort)
    $client.EnableSsl = $UseSsl
    if ($UseIntegratedAuth) { $client.UseDefaultCredentials = $true } else { $client.UseDefaultCredentials = $false }
    try { $client.Send($MailMessage) }
    finally { $MailMessage.Dispose(); $client.Dispose() }
}

function Send-FatalErrorEmail {
    param([Parameter(Mandatory=$true)][string]$Subject,[Parameter(Mandatory=$true)][string]$HtmlBody)

    if (-not $EnableErrorEmail) { return }
    if ($DryRun) { Write-Log "DryRun: skipping fatal error email." "INFO" "DRYRUN"; return }
    if ([string]::IsNullOrWhiteSpace($ToSummary)) { return }

    try {
        $mail = New-HtmlMailMessage -From $From -To $ToSummary -Subject $Subject -HtmlBody $HtmlBody -BccAll $BccAll
        Send-HtmlMail -MailMessage $mail -SmtpEndpoint $script:SmtpEndpoint -SmtpPort $SmtpPort -UseIntegratedAuth $UseIntegratedAuth -UseSsl $UseSsl
        Write-Log "Fatal error email sent." "INFO" "MAIL"
    }
    catch {
        Write-Log "Failed to send fatal error email: $($_.Exception.Message)" "WARN" "MAIL"
    }
}

function Send-SummaryEmail {
    param([Parameter(Mandatory=$true)][string]$Subject,[Parameter(Mandatory=$true)][string]$HtmlBody)

    if (-not $EnableSummaryEmail) { return }
    if ($DryRun) { Write-Log "DryRun: skipping summary email." "INFO" "DRYRUN"; return }
    if ([string]::IsNullOrWhiteSpace($ToSummary)) { return }

    try {
        $mail = New-HtmlMailMessage -From $From -To $ToSummary -Subject $Subject -HtmlBody $HtmlBody -BccAll $BccAll
        Send-HtmlMail -MailMessage $mail -SmtpEndpoint $script:SmtpEndpoint -SmtpPort $SmtpPort -UseIntegratedAuth $UseIntegratedAuth -UseSsl $UseSsl
        Write-Log "Summary email sent." "INFO" "MAIL"
    }
    catch {
        Write-Log "Failed to send summary email: $($_.Exception.Message)" "ERROR" "MAIL"; throw
    }
}

function Should-SendSummaryEmail {
    param([Parameter(Mandatory=$true)][int]$CurrentCount,[Parameter(Mandatory=$true)][string]$Mode,[Parameter(Mandatory=$true)][string]$StatePath)

    if (-not $EnableSummaryEmail) { return $false }
    switch ($Mode) {
        "Never"    { return $false }
        "Always"   { return $true }
        "OnChange" {
            $previous = $null
            if (Test-Path $StatePath) { $previous = (Get-Content -Path $StatePath -ErrorAction SilentlyContinue | Select-Object -First 1) }
            if ([string]::IsNullOrWhiteSpace($previous)) { return $true }
            $prevInt = 0
            if (-not [int]::TryParse($previous, [ref]$prevInt)) { return $true }
            return ($prevInt -ne $CurrentCount)
        }
        default { return $false }
    }
}

function Save-SummaryState {
    param([Parameter(Mandatory=$true)][int]$Count,[Parameter(Mandatory=$true)][string]$StatePath)
    if ($DryRun) { return }
    try { Set-Content -Path $StatePath -Value $Count -Encoding UTF8 } catch { }
}

# ==========================================================
# SmartM365 - HTML helpers
# ==========================================================
function Html-Encode {
    param([Parameter(Mandatory=$false)][object]$Value)
    if ($null -eq $Value) { return "" }
    $s = [string]$Value
    $s = $s.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace('"',"&quot;").Replace("'","&#39;")
    return $s
}

function Convert-ObjectsToHtmlTable {
    param(
        [Parameter(Mandatory=$true)][AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory=$true)][string[]]$Columns,
        [Parameter(Mandatory=$true)][string]$Title
    )

    if ($null -eq $Rows) { $Rows = @() }
    if (-not $Rows -or $Rows.Count -eq 0) {
        return "<h3 style='margin:16px 0 8px 0;'>$(Html-Encode $Title)</h3><div style='font-family:Segoe UI,Arial;font-size:12px;color:#666;'>No data.</div>"
    }

    $th  = ($Columns | ForEach-Object { "<th style='padding:6px;border:1px solid #ddd;background:#f5f5f5;text-align:left;'>$(Html-Encode $_)</th>" }) -join ""
    $trs = ($Rows | ForEach-Object {
        $row = $_
        $tds = ($Columns | ForEach-Object { "<td style='padding:6px;border:1px solid #ddd;'>$(Html-Encode $row.$_)</td>" }) -join ""
        "<tr>$tds</tr>"
    }) -join "`n"

    return @"
<h3 style="margin:16px 0 8px 0;">$(Html-Encode $Title)</h3>
<table style="border-collapse:collapse;font-family:Segoe UI,Arial;font-size:12px;">
<tr>$th</tr>
$trs
</table>
"@
}

# ==========================================================
# SmartM365 - Auth helpers
# ==========================================================
function Get-AppCertificate {
    param([Parameter(Mandatory=$true)][string]$Thumbprint)

    $thumb = $Thumbprint.Replace(" ", "").ToUpper()
    $cert = Get-Item -Path "Cert:\CurrentUser\My\$thumb" -ErrorAction SilentlyContinue
    if (-not $cert) { $cert = Get-Item -Path "Cert:\LocalMachine\My\$thumb" -ErrorAction SilentlyContinue }
    if (-not $cert) { throw "Certificate not found for thumbprint: $thumb" }
    if (-not $cert.HasPrivateKey) { throw "Certificate has no private key: $thumb" }
    return $cert
}

function Get-GraphTokenWithCert {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    if (-not (Get-Module -ListAvailable -Name "MSAL.PS")) {
        throw "MSAL.PS module is not installed. Install it with: Install-Module MSAL.PS -Scope AllUsers"
    }
    Import-Module MSAL.PS -ErrorAction Stop
    $token = Get-MsalToken -TenantId $TenantId -ClientId $ClientId -ClientCertificate $Certificate -Scopes "https://graph.microsoft.com/.default"
    if (-not $token -or [string]::IsNullOrWhiteSpace($token.AccessToken)) { throw "Failed to acquire access token using MSAL.PS." }
    return $token.AccessToken
}

# ==========================================================
# SmartM365 - Graph token auto-refresh
# ==========================================================
function Refresh-GraphTokenIfNeeded {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory=$true)][ref]$Headers,
        [Parameter()][int]$RefreshThresholdMinutes = 55
    )

    if ($script:TokenAcquiredAt -eq [datetime]::MinValue) { return }

    $elapsed = (Get-Date) - $script:TokenAcquiredAt
    if ($elapsed.TotalMinutes -lt $RefreshThresholdMinutes) { return }

    Write-Log "Graph token age $([int]$elapsed.TotalMinutes) min >= threshold $RefreshThresholdMinutes min. Refreshing..." "INFO" "AUTH"
    try {
        $newToken = Get-GraphTokenWithCert -TenantId $TenantId -ClientId $ClientId -Certificate $Certificate
        $Headers.Value = @{ Authorization = "Bearer $newToken" }
        $script:TokenAcquiredAt = Get-Date
        Write-Log "Graph token refreshed successfully at $($script:TokenAcquiredAt.ToString('HH:mm:ss'))." "INFO" "AUTH"
    }
    catch {
        Write-Log "Failed to refresh Graph token: $($_.Exception.Message). Continuing with existing token." "WARN" "AUTH"
    }
}

# ==========================================================
# SmartM365 - Graph REST helpers
# ==========================================================
function Try-GetGraphErrorMessage {
    param([Parameter(Mandatory=$true)]$ErrorRecord)

    $bodyText = $null
    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.GetResponseStream()) {
            $sr = New-Object System.IO.StreamReader($ErrorRecord.Exception.Response.GetResponseStream())
            $bodyText = $sr.ReadToEnd()
        }
    } catch { }
    if ([string]::IsNullOrWhiteSpace($bodyText)) { return $null }
    try {
        $obj = $bodyText | ConvertFrom-Json -ErrorAction Stop
        if ($obj.error -and $obj.error.message) { return [string]$obj.error.message }
    } catch { }
    return $bodyText
}

function Build-GraphUrl {
    param(
        [Parameter(Mandatory=$true)][string]$BaseUrl,
        [Parameter(Mandatory=$true)][string]$Select
    )

    $ub = [System.UriBuilder]::new($BaseUrl)
    $encodedSelect = [System.Uri]::EscapeDataString($Select)
    $ub.Query = "`$select=$encodedSelect"
    return $ub.Uri.AbsoluteUri
}

function Invoke-GraphGetPagedRest {
    param(
        [Parameter(Mandatory=$true)][string]$InitialUrl,
        [Parameter(Mandatory=$true)][hashtable]$Headers
    )

    $null = [System.Uri]::new($InitialUrl)

    $items = New-Object System.Collections.Generic.List[object]
    $url = $InitialUrl
    $script:GraphPage = 0

    while ($url) {
        if ($ConsoleShowGetUrls) {
            Write-Log "GET $url" "INFO" "GRAPH"
        }
        else {
            $script:GraphPage++
            Write-Log ("Fetching page {0}" -f $script:GraphPage) "INFO" "GRAPH"
            $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $auditLine = "[$ts][INFO][$ScriptName][$RunId] GET $url"
            Add-Content -Path $LogPath -Value $auditLine -Encoding UTF8
        }

        try {
            $resp = Invoke-RestMethod -Method GET -Uri $url -Headers $Headers
        }
        catch {
            $body = $null
            try {
                if ($_.Exception.Response -and $_.Exception.Response.GetResponseStream()) {
                    $sr = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $body = $sr.ReadToEnd()
                }
            }
            catch { }
            if ($body) { Write-Log "Graph error body: $body" "ERROR" "GRAPH" }
            throw
        }

        if ($resp.value) {
            foreach ($i in $resp.value) { [void]$items.Add($i) }
        }

        $url = $resp.'@odata.nextLink'
    }

    return $items
}

# ==========================================================
# SmartM365 - File helpers (atomic copy + archive)
# ==========================================================
function Copy-FileAtomic {
    param([Parameter(Mandatory=$true)][string]$SourcePath,[Parameter(Mandatory=$true)][string]$DestinationFinal,[Parameter(Mandatory=$true)][string]$DestinationTemp)

    if (-not (Test-Path $SourcePath)) { throw "Source file not found: $SourcePath" }
    Ensure-Directory -Path (Split-Path -Parent $DestinationFinal)
    if (Test-Path $DestinationTemp) { Remove-Item -Path $DestinationTemp -Force -ErrorAction SilentlyContinue }
    Copy-Item -Path $SourcePath -Destination $DestinationTemp -Force
    Move-Item -Path $DestinationTemp -Destination $DestinationFinal -Force
}

function Write-ArchiveCopyAndPrune {
    param(
        [Parameter(Mandatory=$true)][string]$SourceCsv,
        [Parameter(Mandatory=$true)][string]$ArchiveFolder,
        [Parameter(Mandatory=$true)][string]$BaseNameWithoutExt,
        [Parameter(Mandatory=$true)][string]$RunStamp,
        [Parameter(Mandatory=$true)][int]$Keep
    )

    Ensure-Directory -Path $ArchiveFolder
    $archiveFinal = Join-Path $ArchiveFolder ("{0}_{1}.csv"     -f $BaseNameWithoutExt, $RunStamp)
    $archiveTemp  = Join-Path $ArchiveFolder ("{0}_{1}.csv.tmp" -f $BaseNameWithoutExt, $RunStamp)
    Copy-FileAtomic -SourcePath $SourceCsv -DestinationFinal $archiveFinal -DestinationTemp $archiveTemp
    Write-Log "Archive CSV created: $archiveFinal" "INFO" "DATA"
    Prune-Files -Folder $ArchiveFolder -Filter ("{0}_*.csv" -f $BaseNameWithoutExt) -Keep $Keep
}

# ==========================================================
# SmartM365 - Business helpers (normalization + mapping)
# ==========================================================
function Get-NormalizedDeviceName {
    param([AllowNull()][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    return $Name.Trim().ToUpper()
}

function Get-UpgradeEligibilityLabel {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return "UNKNOWN" }
    $v = ("$Value").Trim()

    switch -Regex ($v) {
        '^(0|upgraded)$'   { return "UPGRADED" }
        '^(3|capable)$'    { return "CAPABLE" }
        '^(2|notcapable)$' { return "NOT CAPABLE" }
        '^(unknown)$'      { return "UNKNOWN" }
        default            { return "UNKNOWN" }
    }
}

# ==========================================================
# SmartM365 - Summary helpers
# ==========================================================
function Get-ReadinessSummary {
    param([Parameter(Mandatory=$true)][object[]]$Rows)

    $total = $Rows.Count

    $byLabel = $Rows |
        Group-Object -Property UpgradeEligibilityLabel |
        Sort-Object -Property Count -Descending |
        ForEach-Object { [pscustomobject]@{ Label = $_.Name; Count = $_.Count } }

    $byCode = $Rows |
        Group-Object -Property UpgradeEligibility |
        Sort-Object -Property Count -Descending |
        ForEach-Object { [pscustomobject]@{ Code = $_.Name; Count = $_.Count } }

    return [pscustomobject]@{
        Total   = $total
        ByLabel = $byLabel
        ByCode  = $byCode
    }
}

# ==========================================================
# SmartM365 - Main
# ==========================================================
$ErrorActionPreference  = "Stop"
$script:SmtpEndpoint    = $null
$script:TokenAcquiredAt = [datetime]::MinValue
$scriptStart = Get-Date

try {
    Assert-PS7

    Ensure-Directory -Path $BasePath
    Ensure-Directory -Path $ArchivePath
    Ensure-Directory -Path $LogsPath
    Ensure-Directory -Path $DataLastPath

    Prune-Files -Folder $LogsPath -Filter "$ScriptName-*.log" -Keep 10

    Invoke-SmartM365Preflight -ScriptName $ScriptName -RequiredModules @('MSAL.PS') -OutputPaths @($BasePath,$ArchivePath,$LogsPath,$DataLastPath) | Out-Null

    # ----------------------------------------------------------
    # SmartM365.Core is imported during tenant/config initialization.
    # Legacy SharePoint upload remains disabled; SmartM365 upload runs after CSV export.
    # ----------------------------------------------------------
    if ($SmartM365SharePointUploadEnabled) {
        Write-Log "SmartM365 SharePoint upload is enabled." "INFO" "SP"
    }

    if ($ConsolePretty) {
        Write-Host ""
        Write-Host ("=" * 78) -ForegroundColor DarkGray
        Write-Host ("{0}  v{1}  {2}" -f $ScriptName, $ScriptVersion, (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -ForegroundColor White
        Write-Host ("Host: {0}   RunId: {1}" -f $env:COMPUTERNAME, $RunId) -ForegroundColor DarkGray
        if ($DryRun) { Write-Host "*** DRY-RUN MODE: no files written, no emails sent ***" -ForegroundColor Cyan }
        Write-Host ("=" * 78) -ForegroundColor DarkGray
    }

    Write-Log "Starting export. Version=$ScriptVersion DryRun=$DryRun MaxDevices=$MaxDevices" "INFO"

        if ($EnableSummaryEmail -or $EnableErrorEmail) {
        $script:SmtpEndpoint = if ($RelayIp) { $RelayIp } else { Resolve-IPv4Address -HostName $SmtpHostName }
        Write-Log "SMTP endpoint: $script:SmtpEndpoint (IntegratedAuth=$UseIntegratedAuth HostName=$SmtpHostName)" "INFO" "SMTP"
    }
    else {
        Write-Log "SMTP email disabled because SmtpServer is not configured." "INFO" "SMTP"
    }

    $cert = Get-AppCertificate -Thumbprint $Thumb
    Write-Log "Certificate loaded: Subject=$($cert.Subject)" "INFO" "AUTH"

    $token = Get-GraphTokenWithCert -TenantId $TenantId -ClientId $AppId -Certificate $cert
    $script:TokenAcquiredAt = Get-Date
    Write-Log "Access token acquired at $($script:TokenAcquiredAt.ToString('HH:mm:ss'))." "INFO" "AUTH"

    $headers = @{ Authorization = "Bearer $token" }

    Invoke-SmartM365Preflight `
        -ScriptName $ScriptName `
        -RequiredGraphApplicationPermissions @('DeviceManagementManagedDevices.Read.All') `
        -GraphAccessToken $token `
        -GraphProbeUris @("https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsWorkFromAnywhereMetrics('allDevices')/metricDevices?`$top=1") | Out-Null

    # ----------------------------------------------------------
    # Graph query
    # ----------------------------------------------------------
    $baseUrl = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsWorkFromAnywhereMetrics('allDevices')/metricDevices"
    $select  = "id,deviceName,upgradeEligibility,azureAdJoinType"
    $url     = Build-GraphUrl -BaseUrl $baseUrl -Select $select
    Write-Log "Graph query prepared (select=$select)" "INFO" "GRAPH"

    Refresh-GraphTokenIfNeeded -TenantId $TenantId -ClientId $AppId -Certificate $cert -Headers ([ref]$headers)

    $data = Invoke-GraphGetPagedRest -InitialUrl $url -Headers $headers
    Write-Log "Devices retrieved: $($data.Count)" "INFO" "DATA"

    # ----------------------------------------------------------
    # MaxDevices limit
    # ----------------------------------------------------------
    if ($MaxDevices -gt 0 -and $data.Count -gt $MaxDevices) {
        Write-Log "MaxDevices=${MaxDevices}: limiting to first $MaxDevices devices out of $($data.Count)." "WARN" "DATA"
        $data = $data | Select-Object -First $MaxDevices
    }

    # ----------------------------------------------------------
    # Transform
    # ----------------------------------------------------------
    $exportTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    $out = foreach ($d in $data) {
        [pscustomobject]@{
            DeviceName              = $d.deviceName
            NormalizedDeviceName    = Get-NormalizedDeviceName -Name $d.deviceName
            UpgradeEligibility      = $d.upgradeEligibility
            UpgradeEligibilityLabel = Get-UpgradeEligibilityLabel -Value $d.upgradeEligibility
            AzureAdJoinType         = $d.azureAdJoinType
            GraphId                 = $d.id
            OrgDomain               = $OrgDomain
            ExportDateTime          = $exportTime
            RunId                   = $RunId
        }
    }

    # ----------------------------------------------------------
    # De-duplication by NormalizedDeviceName + audit CSV
    # ----------------------------------------------------------
    $beforeCount = $out.Count

    $dupGroupsByName = $out |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.NormalizedDeviceName) } |
        Group-Object -Property NormalizedDeviceName |
        Where-Object { $_.Count -gt 1 }

    if ($dupGroupsByName) {
        if (-not $DryRun) {
            $dupAuditPath = Join-Path $BasePath ("Intune_Devices_Win11Readiness_DUPLICATES_ByName_{0}.csv" -f $RunStamp)
            $dupGroupsByName | ForEach-Object { $_.Group } | Export-Csv -Path $dupAuditPath -NoTypeInformation -Encoding UTF8
            Write-Log "Duplicate audit CSV created (by NormalizedDeviceName): $dupAuditPath (Groups=$($dupGroupsByName.Count))." "WARN" "DATA"
        }
        else {
            Write-Log "DryRun: $($dupGroupsByName.Count) duplicate group(s) detected (by NormalizedDeviceName); audit CSV skipped." "WARN" "DRYRUN"
        }
    }
    else {
        Write-Log "No duplicates detected (by NormalizedDeviceName)." "INFO" "DATA"
    }

    $out = $out | Sort-Object NormalizedDeviceName, DeviceName | Sort-Object NormalizedDeviceName -Unique

    $afterCount = $out.Count
    $dupCount   = $beforeCount - $afterCount

    if ($dupCount -gt 0) {
        Write-Log "Duplicates removed by NormalizedDeviceName: $dupCount (Before=$beforeCount, After=$afterCount)." "WARN" "DATA"
    }
    else {
        Write-Log "No duplicates removed (Count=$afterCount)." "INFO" "DATA"
    }

    $count = $afterCount

    # ----------------------------------------------------------
    # CSV export - DATA-ALL (atomic), DATA-LAST (atomic), Archive
    # ----------------------------------------------------------
    $spUploadStatus = "Disabled"
    $spUploadError  = ""
    $spDestPath     = ""

    if ($DryRun) {
        Write-Log "DryRun: skipping write of $CsvFinal." "INFO" "DRYRUN"
    }
    else {
        if (Test-Path $CsvTemp) { Remove-Item -Path $CsvTemp -Force -ErrorAction SilentlyContinue }
        $out | Sort-Object DeviceName | Export-Csv -Path $CsvTemp -NoTypeInformation -Encoding UTF8
        Move-Item -Path $CsvTemp -Destination $CsvFinal -Force
        Write-Log "CSV exported: $CsvFinal" "INFO" "DATA"

        Copy-FileAtomic -SourcePath $CsvFinal -DestinationFinal $CsvLastFinal -DestinationTemp $CsvLastTemp
        Write-Log "DATA-LAST updated: $CsvLastFinal" "INFO" "DATA"

        $archiveBase = [System.IO.Path]::GetFileNameWithoutExtension($CsvName)
        Write-ArchiveCopyAndPrune -SourceCsv $CsvFinal -ArchiveFolder $ArchivePath -BaseNameWithoutExt $archiveBase -RunStamp $RunStamp -Keep 10

        if ($SmartM365SharePointUploadEnabled) {
            foreach ($uploadFile in @($CsvFinal, $CsvLastFinal)) {
                try {
                    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $uploadFile | Out-Null
                    Write-Log "SharePoint upload completed through SmartM365.Core: $uploadFile" "INFO" "SP"
                }
                catch {
                    Write-Log "SharePoint upload failed through SmartM365.Core: $($_.Exception.Message)" "WARN" "SP"
                }
            }
        }

        # ----------------------------------------------------------
        # SharePoint upload (non-blocking - uses same Graph token)
        # ----------------------------------------------------------
        if ($EnableSharePointUpload) {
            try {
                Refresh-GraphTokenIfNeeded -TenantId $TenantId -ClientId $AppId -Certificate $cert -Headers ([ref]$headers)
                Write-Log "SharePoint upload starting: $SP_SiteHostname$SP_SitePath / $SP_LibraryDisplayName / $SP_TargetFolderPath" "INFO" "SP"

                $spLogger  = { param($m, $l, $s) Write-Log $m $l $s }
                $spDriveId = Resolve-EmeritSpDriveId -Headers $headers -SiteHostname $SP_SiteHostname -SitePath $SP_SitePath -LibraryDisplayName $SP_LibraryDisplayName -Logger $spLogger
                $spResult  = Invoke-EmeritSpFileUpload -Headers $headers -LocalFilePath $CsvFinal -TargetFolderPath $SP_TargetFolderPath -DriveId $spDriveId -ChunkSize $SP_ChunkSize -Logger $spLogger

                $spUploadStatus = $spResult.Status
                $spDestPath     = $spResult.DestPath
                Write-Log "SharePoint upload: $spUploadStatus ($($spResult.DurationMs) ms) -> $spDestPath" "INFO" "SP"

                if ($spResult.Status -eq "Error") {
                    $spUploadError = $spResult.Error
                    Write-Log "SharePoint upload error (non-blocking): $spUploadError" "WARN" "SP"
                }
            }
            catch {
                $spUploadStatus = "Error"
                $spUploadError  = $_.Exception.Message
                Write-Log "SharePoint upload failed (non-blocking): $spUploadError" "WARN" "SP"
            }
        }
    }

    if ($DryRun -and $EnableSharePointUpload) {
        $spUploadStatus = "Skipped (DryRun)"
        Write-Log "DryRun: skipping SharePoint upload." "INFO" "DRYRUN"
    }

    # ----------------------------------------------------------
    # KPI calculations
    # ----------------------------------------------------------
    $summary = Get-ReadinessSummary -Rows $out
    Write-Log "KPIs: Total=$count Duplicates=$dupCount Before=$beforeCount" "INFO" "KPI"

    # ----------------------------------------------------------
    # Success email
    # ----------------------------------------------------------
    $should = Should-SendSummaryEmail -CurrentCount $count -Mode $SummaryEmailMode -StatePath $SummaryStatePath
    if ($should) {
        $duration = New-TimeSpan -Start $scriptStart -End (Get-Date)

        $byLabelTable = Convert-ObjectsToHtmlTable -Rows $summary.ByLabel -Columns @("Label","Count") -Title "Breakdown by Eligibility Label"
        $byCodeTable  = Convert-ObjectsToHtmlTable -Rows $summary.ByCode  -Columns @("Code","Count")  -Title "Breakdown by Raw Code"

        $spStatusColor = if ($spUploadStatus -eq "Error") { "color:#b00020;" } elseif ($spUploadStatus -eq "Success") { "color:#007700;" } else { "" }

        $subject = "[SmartM365] Win11 Readiness export - SUCCESS - Devices=$count Duplicates=$dupCount"

        $html = @"
<html><body style="font-family:Segoe UI,Arial;">
<h2 style="margin:0 0 10px 0;">Windows 11 Readiness export - SUCCESS</h2>
<ul>
<li><b>Script</b>: $(Html-Encode $ScriptName) v$(Html-Encode $ScriptVersion)</li>
<li><b>RunId</b>: $(Html-Encode $RunId)</li>
<li><b>Host</b>: $(Html-Encode $env:COMPUTERNAME)</li>
<li><b>Total devices</b>: $count</li>
<li><b>Duplicates removed</b>: $dupCount (before=$beforeCount)</li>
<li><b>Execution time</b>: $([int]$duration.TotalSeconds) seconds</li>
<li><b>ExportDateTime</b>: $(Html-Encode $exportTime)</li>
$(if ($DryRun) { "<li><b style='color:#b00020;'>DRY-RUN MODE: no files written</b></li>" })
</ul>

<h3 style="margin:16px 0 8px 0;">Output CSV</h3>
<ul>
<li><b>DATA-ALL</b>: $(Html-Encode $CsvFinal)</li>
<li><b>DATA-LAST</b>: $(Html-Encode $CsvLastFinal)</li>
</ul>

<h3 style="margin:16px 0 8px 0;">SharePoint Upload</h3>
<ul>
<li><b>Status</b>: <span style="$spStatusColor">$(Html-Encode $spUploadStatus)</span></li>
<li><b>Target</b>: $(Html-Encode "$SP_SiteHostname$SP_SitePath / $SP_LibraryDisplayName / $SP_TargetFolderPath")</li>
$(if ($spDestPath)    { "<li><b>Destination</b>: $(Html-Encode $spDestPath)</li>" })
$(if ($spUploadError) { "<li><b style='color:#b00020;'>Error</b>: $(Html-Encode $spUploadError)</li>" })
</ul>

$byLabelTable
$byCodeTable

</body></html>
"@

        Send-SummaryEmail -Subject $subject -HtmlBody $html
        Save-SummaryState -Count $count -StatePath $SummaryStatePath
    }
    else {
        Write-Log "Summary email skipped (Mode=$SummaryEmailMode)." "INFO" "MAIL"
    }

    Write-Log "Completed successfully. Version=$ScriptVersion DryRun=$DryRun Devices=$count Duplicates=$dupCount" "INFO"
}
catch {
    $err = $_
    Write-Log "Fatal error: $($err.Exception.Message)" "ERROR"

    $duration = New-TimeSpan -Start $scriptStart -End (Get-Date)
    $subject  = "[SmartM365] Win11 Readiness export - FAILED - $ScriptName"
    $html = @"
<html><body style="font-family:Segoe UI,Arial;">
<h2 style="margin:0 0 10px 0;color:#b00020;">Windows 11 Readiness export - FAILED</h2>
<ul>
<li><b>Script</b>: $(Html-Encode $ScriptName) v$(Html-Encode $ScriptVersion)</li>
<li><b>RunId</b>: $(Html-Encode $RunId)</li>
<li><b>Host</b>: $(Html-Encode $env:COMPUTERNAME)</li>
<li><b>Execution time</b>: $([int]$duration.TotalSeconds) seconds</li>
</ul>
<h3 style="margin:16px 0 8px 0;">Error</h3>
<pre style="white-space:pre-wrap;background:#f6f6f6;border:1px solid #ddd;padding:10px;">$(Html-Encode ($err | Out-String))</pre>
</body></html>
"@

    Send-FatalErrorEmail -Subject $subject -HtmlBody $html
    throw
}
finally {
    Cleanup-TempFiles
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCROpkDTZOUeqwk
# F/GoHY7i0Q4yn5Lm1NX7yFbBER/lwqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDzuhJFvWt6SSAWv4l7
# J+xL+iAwCSrxb5itjSl5DJiaajANBgkqhkiG9w0BAQEFAASCAYAqen5OzuGYk5Z0
# nSkT8xDdCwCmFvmyLqc+2S2WSJL9dS3jtd8NlyN02j/hVnbX04aM2ujmbw0PNL7N
# 2VOaT5S96ZkG4AlY5QEaQChuay94JqSiG2PhAcDZHm5gVv17VR5jXyDt03Uy+id3
# XDmHMN9GHYE9Etkgqh0pl5uE4I0kygh/qZF9z4ueYT4Yw4CIudMZh+RpTEvok37P
# QBb9Zju/86ywlDmeQ47LaKMM/y4xZ54pw6ZhAJaL7rGtAeucXtMeVgZHqpxKQKtS
# kAGtZFU/hAGV1mGy3aeaeQVujvKM7ME7uuT8xIZR+YpGmKMSPBzubBz79fbpNkH0
# qBYWvtCcIkY9IkCGpf3nL/Q0XPit4Ro8T3/GH2DN4LpQKDu2brFa8B9kYwDpilaF
# E4A+1RQL+mw+uVGBpkfQl17dvLeVAWD4RAlv4Jxtxz2zIu+HlXL/Nd/JCGJPSplP
# j+PdtaPcqwFKcFlurh23v4ZLiF1wTgYkHEHHymZNaBjrGc3VoAA=
# SIG # End signature block
