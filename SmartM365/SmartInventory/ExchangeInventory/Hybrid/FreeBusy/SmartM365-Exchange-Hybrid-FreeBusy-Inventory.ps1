<#
.SYNOPSIS
    Weekly Exchange hybrid free/busy configuration check with Microsoft CSS-Exchange Free/Busy Checker.

.DESCRIPTION
    Runs the vendored Free/Busy Checker (ThirdParty\CSS-Exchange\FreeBusyChecker, see PROVENANCE.md) on the
    local Exchange server and publishes its results:
    - Exchange_Hybrid_FreeBusy_Summary.csv: one row per side (Exchange on-premises, Exchange Online) and
      authentication (DAuth, OAuth) with the run status, the side status and the result counts;
    - Exchange_Hybrid_FreeBusy_Findings.csv: one row per result line (status OK, Warning, Error, Info or
      NotAuthorized), with values that can reveal names, addresses, URLs, certificates or identifiers masked.

    Before every run the collector copies Free/Busy Checker and its runner into a local work folder that
    only administrators can change (default %ProgramFiles%\SmartM365\FreeBusy\Work), checks every parent,
    checks the copy (expected SHA256 and Microsoft signature) and the runner, and checks both again just
    before the start. The runner runs in its own Windows PowerShell process with a time limit. Exchange
    Online uses the tenant app and its certificate only (prefix EO): without that session the on-premises
    part is still published and Exchange Online is NotConnected. Free/Busy Checker always runs with
    -SkipVersionCheck alone, so it never updates itself. Its HTML report and transcript stay in the work
    folder (last run only) and are read as text only; they are never published.

.VERSION
1.0.1
.REQUIREMENTS
    Windows PowerShell 5.1, elevated, on an Exchange server with the Exchange Management Tools.
    Modules: SmartM365 WindowsPowerShell5 compatibility module; ActiveDirectory (Windows feature RSAT-AD-PowerShell);
    ExchangeOnlineManagement 3.x (installed for the current user by the runner when missing).
    Exchange Online: the tenant app with its certificate; Free/Busy Checker cmdlets not exposed to the app are reported NotAuthorized.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Script Name : SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1
    Version     : 1.0.1
    Author      : https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [string]$Tenant = "test",

    [string]$OutputRoot,

    [int]$MaxItems = 0
)
# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
if ($MaxItems -gt 0) {
    throw "-MaxItems is not supported by SmartM365-Exchange-Hybrid-FreeBusy-Inventory because Free/Busy Checker always checks the whole hybrid configuration."
}
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath

$ScriptName = "SmartM365-Exchange-Hybrid-FreeBusy-Inventory"
$ScriptVersion = "1.0.1"
$RunId = (Get-Date).ToString("yyyyMMdd-HHmmss")
# Vendored Microsoft CSS-Exchange Free/Busy Checker (see ThirdParty\CSS-Exchange\FreeBusyChecker\PROVENANCE.md).
$FreeBusyCheckerRelativePath = 'ThirdParty\CSS-Exchange\FreeBusyChecker\FreeBusyChecker.ps1'
$FreeBusyCheckerVersion = '26.06.06.1238'
$FreeBusyCheckerExpectedSha256 = '5E9BCF887AB3A06F29208EB0C3CAD87D0B38C30019F18BB8DE27791038EEEC34'
$FreeBusyRunnerFileName = 'SmartM365-Exchange-Hybrid-FreeBusy-Runner.ps1'

$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

function Merge-SmartM365ScriptLocalConfig {
    param(
        [Parameter(Mandatory)]
        [string]$ConfigPath
    )

    $localConfig = Read-SmartM365JsonConfig -Path $ConfigPath -Required
    if ($null -eq $localConfig -or $localConfig.Count -eq 0) { return }

    $effectiveConfig = ConvertTo-SmartM365Hashtable -InputObject $script:SmartM365EffectiveConfig
    foreach ($key in $localConfig.Keys) {
        $value = $localConfig[$key]
        if ($value -is [string] -and $value -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { continue }
        $effectiveConfig[$key] = $value
    }

    $script:SmartM365EffectiveConfig = [pscustomobject]$effectiveConfig
}

Merge-SmartM365ScriptLocalConfig -ConfigPath (Join-Path -Path $PSScriptRoot -ChildPath "$ScriptName.local.json")

function Resolve-SmartM365ConfigTokenValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    $resolved = [string]$Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $tokenMatches) {
            $tokenName = $match.Groups['Name'].Value
            $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$tokenName]
            if ($null -eq $property -or $null -eq $property.Value) { continue }

            $tokenValue = Resolve-SmartM365ConfigTokenValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Resolve-SmartM365ConfigTokenValue -Value '{{DataAllRootPath}}\Exchange\Hybrid\FreeBusy'
}
else {
    $OutputRoot = Resolve-SmartM365ConfigTokenValue -Value $OutputRoot
}
$OutputFolder = Join-Path $OutputRoot $RunId
$LogRoot = Resolve-SmartM365ConfigTokenValue -Value '{{LogAllRootPath}}'
if ([string]::IsNullOrWhiteSpace($LogRoot) -or $LogRoot -eq '{{LogAllRootPath}}') {
    $LogRoot = Join-Path -Path (Split-Path -Path $OutputRoot -Parent) -ChildPath 'LOG-ALL'
}
$LogFolder = Join-Path -Path $LogRoot -ChildPath $ScriptName
$LogFile = Join-Path $LogFolder "$ScriptName-$RunId.log"
$TranscriptFile = Join-Path $LogFolder "$ScriptName-$RunId`_Transcript.log"
$script:TranscriptStarted = $false
$script:FreeBusyGeneratedCsvPaths = New-Object 'System.Collections.Generic.List[string]'
$script:FreeBusySharePointUploads = New-Object System.Collections.ArrayList
$script:FreeBusyWarningCount = 0
$script:FreeBusyErrorCount = 0

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    if ($Level -eq 'WARN') { $script:FreeBusyWarningCount++ }
    if ($Level -eq 'ERROR') { $script:FreeBusyErrorCount++ }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { "$timestamp [$Level] $_" })
    $line | ForEach-Object { Write-Host $_ }
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Stop-FreeBusyTranscriptSafely {
    if (-not $script:TranscriptStarted) { return }
    try { Stop-SmartM365RunTranscript | Out-Null } catch { Write-Verbose "Transcript already stopped: $($_.Exception.Message)" }
    $script:TranscriptStarted = $false
}

function Complete-FreeBusyRun {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][datetime]$Started,
        [AllowNull()][string]$ErrorMessage
    )

    $ended = Get-Date
    Write-Log 'Execution summary:'
    Write-Log ("  Status: {0}" -f $Status)
    Write-Log ("  Duration: {0}" -f ($ended - $Started))
    Write-Log ("  Started: {0}" -f $Started.ToString('yyyy-MM-dd HH:mm:ss zzz'))
    Write-Log ("  Ended: {0}" -f $ended.ToString('yyyy-MM-dd HH:mm:ss zzz'))
    Write-Log ("  ScriptName: {0}" -f $ScriptName)
    Write-Log ("  ScriptVersion: {0}" -f $ScriptVersion)
    Write-Log ("  TenantProfile: {0}" -f $Tenant)
    Write-Log ("  TenantKey: {0}" -f $global:SmartM365TenantKey)
    Write-Log ("  OutputPath: {0}" -f $OutputFolder)
    Write-Log ("  LogTextFile: {0}" -f $LogFile)
    Write-Log ("  TranscriptFile: {0}" -f $TranscriptFile)
    Write-Log ("  GeneratedCsvFiles: {0}" -f $script:FreeBusyGeneratedCsvPaths.Count)
    Write-Log ("  Warnings: {0}" -f $script:FreeBusyWarningCount)
    Write-Log ("  Errors: {0}" -f $script:FreeBusyErrorCount)
    if (-not [string]::IsNullOrWhiteSpace($ErrorMessage)) { Write-Log ("  ErrorMessage: {0}" -f $ErrorMessage) }

    Stop-FreeBusyTranscriptSafely
    if (Get-Command Complete-CoreSmartM365SourceReceipt -ErrorAction SilentlyContinue) {
        $sourceReceipt = Complete-CoreSmartM365SourceReceipt -Status $Status -ErrorCount $script:FreeBusyErrorCount
        if ($sourceReceipt) { Invoke-FreeBusySharePointUpload -LocalFilePath $sourceReceipt }
    }
    Invoke-FreeBusySharePointUpload -LocalFilePath $LogFile
    if (Test-Path -LiteralPath $TranscriptFile -PathType Leaf) {
        Invoke-FreeBusySharePointUpload -LocalFilePath $TranscriptFile
    }
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Get-SmartM365EffectiveConfigValue {
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [AllowNull()]
        [object]$DefaultValue = $null
    )

    if ($null -ne $script:SmartM365EffectiveConfig) {
        $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$Name]
        if ($null -ne $property -and $null -ne $property.Value) {
            if ($property.Value -is [string] -and [string]::IsNullOrWhiteSpace($property.Value)) {
                return $DefaultValue
            }

            return Resolve-SmartM365ConfigTokenValue -Value $property.Value
        }
    }

    return $DefaultValue
}

function Import-SmartM365CoreModule {
    if ((Get-Command Publish-CoreSmartM365Csv -ErrorAction SilentlyContinue) -and (Get-Command CoreSendEmailHtmlReport -ErrorAction SilentlyContinue)) {
        return
    }

    $d = $PSScriptRoot
    while ($d) {
        $moduleCandidates = if ($PSVersionTable.PSVersion.Major -lt 6) {
            @(
                (Join-Path -Path $d -ChildPath 'Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psd1'),
                (Join-Path -Path $d -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1')
            )
        }
        else {
            @(
                (Join-Path -Path $d -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1')
            )
        }

        foreach ($modulePath in $moduleCandidates) {
            if (Test-Path -LiteralPath $modulePath) {
                $minimumVersion = if ($modulePath -like '*WindowsPowerShell5*') { '1.0.57' } else { '1.0.87' }
                Import-Module -Name $modulePath -MinimumVersion $minimumVersion -Prefix Core -ErrorAction Stop
                return
            }
        }

        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }

    throw 'SmartM365.Core module not found.'
}

function Invoke-FreeBusySharePointUpload {
    param(
        [Parameter(Mandatory)]
        [string]$LocalFilePath
    )

    $enabled = [bool](Get-SmartM365EffectiveConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false)
    if (-not $enabled) { return }

    $thumbprint = Get-SmartM365EffectiveConfigValue -Name 'Thumbprint' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($thumbprint)) {
        $thumbprint = Get-SmartM365EffectiveConfigValue -Name 'Thumb' -DefaultValue ''
    }

    try {
        Import-SmartM365CoreModule
        $record = Invoke-CoreSmartM365SharePointCsvUpload `
            -LocalFilePath $LocalFilePath `
            -Enabled $true `
            -SiteHostname (Get-SmartM365EffectiveConfigValue -Name 'SharePointSiteHostname' -DefaultValue '') `
            -SitePath (Get-SmartM365EffectiveConfigValue -Name 'SharePointSitePath' -DefaultValue '') `
            -LibraryDisplayName (Get-SmartM365EffectiveConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents') `
            -TargetFolderPath (Get-SmartM365EffectiveConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue '') `
            -AppId (Get-SmartM365EffectiveConfigValue -Name 'AppId' -DefaultValue '') `
            -TenantId (Get-SmartM365EffectiveConfigValue -Name 'TenantId' -DefaultValue '') `
            -Thumbprint $thumbprint
        if ($record) {
            [void]$script:FreeBusySharePointUploads.Add($record)
        }
        Write-Log "SharePoint upload requested: $LocalFilePath"
        return $record
    }
    catch {
        Write-Log "SharePoint upload failed (non-blocking): $($_.Exception.Message)" "WARN"
    }
}

function Export-FreeBusyCsv {
    param(
        [AllowNull()][object[]]$Data,
        [Parameter(Mandatory)]
        [string]$Path,
        [string[]]$Columns = @()
    )

    Import-SmartM365CoreModule
    $latestPath = $null
    $latestCsvFolder = [string](Get-SmartM365EffectiveConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($latestCsvFolder)) {
        Ensure-Directory -Path $latestCsvFolder
        $latestPath = Join-Path -Path $latestCsvFolder -ChildPath (Split-Path -Path $Path -Leaf)
    }
    Publish-CoreSmartM365Csv -Data @($Data) -TimestampedPath $Path -LatestPath $latestPath -Columns $Columns -Delimiter ';' -NoSharePointUpload | Out-Null
    $script:FreeBusyGeneratedCsvPaths.Add($Path) | Out-Null
    Invoke-FreeBusySharePointUpload -LocalFilePath $Path

    if (-not [string]::IsNullOrWhiteSpace($latestPath)) {
        $script:FreeBusyGeneratedCsvPaths.Add($latestPath) | Out-Null
        Write-Log ("CSV latest copy written to: {0}" -f $latestPath)
        Invoke-FreeBusySharePointUpload -LocalFilePath $latestPath
    }
}

function Save-FreeBusyWeeklyHistory {
    $weeklyHistoryEnabled = [bool](Get-SmartM365EffectiveConfigValue -Name 'EnableWeeklyHistory' -DefaultValue $true)
    if (-not $weeklyHistoryEnabled) { return }

    $weeklyHistoryPath = [string](Get-SmartM365EffectiveConfigValue -Name 'WeeklyHistoryFolderPath' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($weeklyHistoryPath)) { return }

    $weeklyHistoryPath = Resolve-SmartM365ConfigTokenValue -Value $weeklyHistoryPath
    $sourceCsvPaths = @($script:FreeBusyGeneratedCsvPaths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -Unique)
    if ($sourceCsvPaths.Count -eq 0) { return }

    try {
        Import-SmartM365CoreModule
        if (Get-Command Add-CoreSmartM365WeeklyHistory -ErrorAction SilentlyContinue) {
            Add-CoreSmartM365WeeklyHistory -SourceCsvPaths $sourceCsvPaths -HistoryRootPath $weeklyHistoryPath -RetentionWeeks ([int](Get-SmartM365EffectiveConfigValue -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)) -HistoryLabel 'Exchange hybrid free/busy' | Out-Null
        }
    }
    catch {
        Write-Log "Weekly history upload failed (non-blocking): $($_.Exception.Message)" 'WARN'
    }
}

function Format-HtmlValue {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return "N/A"
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Send-FreeBusyHtmlReport {
    param(
        [Parameter(Mandatory)][string]$HtmlReportPath,
        [Parameter(Mandatory)][string]$Subject,
        [string[]]$Attachments = @()
    )

    if (-not (Test-Path -LiteralPath $HtmlReportPath -PathType Leaf)) {
        throw "HTML report not found: $HtmlReportPath"
    }

    Import-SmartM365CoreModule

    $from = [string](Get-SmartM365EffectiveConfigValue -Name 'From' -DefaultValue '')
    $to = [string](Get-SmartM365EffectiveConfigValue -Name 'To' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($to)) {
        $to = [string](Get-SmartM365EffectiveConfigValue -Name 'ErrorMailTo' -DefaultValue '')
    }

    $smtpServer = [string](Get-SmartM365EffectiveConfigValue -Name 'SmtpServer' -DefaultValue '')
    $sendMailMode = [string](Get-SmartM365EffectiveConfigValue -Name 'SendMailMode' -DefaultValue '')
    $cc = [string](Get-SmartM365EffectiveConfigValue -Name 'Cc' -DefaultValue '')
    $existingAttachments = @($Attachments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Leaf) })

    $reportHtml = Get-Content -LiteralPath $HtmlReportPath -Raw -ErrorAction Stop
    if ($reportHtml -notmatch 'SmartM365EmailTemplate:v1') {
        $reportHtml = "<!-- SmartM365EmailTemplate:v1 -->`r`n$reportHtml"
    }

    $mailParams = @{
        From        = $from
        To          = $to
        Subject     = $Subject
        BodyHtml    = $reportHtml
        Attachments = $existingAttachments
    }
    if (-not [string]::IsNullOrWhiteSpace($smtpServer)) { $mailParams['SmtpServer'] = $smtpServer }
    if (-not [string]::IsNullOrWhiteSpace($sendMailMode)) { $mailParams['SendMailMode'] = $sendMailMode }
    if (-not [string]::IsNullOrWhiteSpace($cc)) { $mailParams['Cc'] = $cc }

    CoreSendEmailHtmlReport @mailParams
}

function Get-FreeBusyCsvColumn {
    param([Parameter(Mandatory)][ValidateSet('Findings', 'Summary')][string]$Kind)

    if ($Kind -eq 'Findings') {
        return @('ServerName', 'Side', 'Auth', 'Category', 'RowNumber', 'Item', 'Status', 'Value', 'ValueMasked', 'FreeBusyCheckerVersion', 'GenerationTime')
    }
    return @('ServerName', 'Side', 'Auth', 'RunStatus', 'SideStatus', 'RunMessage', 'NotAuthorizedCmdlets', 'MissingRoles', 'DurationSeconds', 'FreeBusyCheckerVersion', 'GenerationTime', 'OkCount', 'WarningCount', 'ErrorCount', 'InfoCount', 'NotAuthorizedCount')
}

function Get-FreeBusyCmdletScope {
    # Where each Exchange Online cmdlet's results appear in the report, and the roles that contain it.
    return @{
        'Get-EOMailbox'                         = [pscustomobject]@{ Auth = 'DAuth'; Category = 'Get-Mailbox'; Roles = 'View-Only Recipients or Mail Recipients' }
        'Get-EOOrganizationRelationship'        = [pscustomobject]@{ Auth = 'DAuth'; Category = 'Get-OrganizationRelationship'; Roles = 'View-Only Configuration' }
        'Get-EOFederatedOrganizationIdentifier' = [pscustomobject]@{ Auth = 'DAuth'; Category = 'Get-FederatedOrganizationIdentifier'; Roles = 'View-Only Configuration' }
        'Get-EOSharingPolicy'                   = [pscustomobject]@{ Auth = 'DAuth'; Category = 'Get-SharingPolicy'; Roles = 'View-Only Configuration' }
        'Test-EOOrganizationRelationship'       = [pscustomobject]@{ Auth = 'DAuth'; Category = 'Test-OrganizationRelationship'; Roles = 'View-Only Configuration or O365SupportViewConfig' }
        'Get-EOIntraOrganizationConnector'      = [pscustomobject]@{ Auth = 'OAuth'; Category = 'Get-IntraOrganizationConnector'; Roles = 'View-Only Configuration' }
        'Get-EOIntraOrganizationConfiguration'  = [pscustomobject]@{ Auth = 'OAuth'; Category = 'Get-IntraOrganizationConfiguration'; Roles = 'View-Only Configuration' }
        'Get-EOOnPremisesOrganization'          = [pscustomobject]@{ Auth = 'OAuth'; Category = 'Get-IntraOrganizationConfiguration'; Roles = 'View-Only Configuration' }
        'Get-EOAuthServer'                      = [pscustomobject]@{ Auth = 'OAuth'; Category = 'Get-AuthServer'; Roles = 'View-Only Configuration' }
        'Test-EOOAuthConnectivity'              = [pscustomobject]@{ Auth = 'OAuth'; Category = 'Test-OAuthConnectivity'; Roles = 'Mail Recipients or Organization Transport Settings' }
    }
}

function Get-FreeBusySetting {
    $timeoutMinutes = [int](Get-SmartM365EffectiveConfigValue -Name 'FreeBusyCheckerTimeoutMinutes' -DefaultValue 30)
    $workRoot = [string](Get-SmartM365EffectiveConfigValue -Name 'FreeBusyWorkFolderPath' -DefaultValue '')
    if ($timeoutMinutes -lt 1 -or $timeoutMinutes -gt 110) { throw "FreeBusyCheckerTimeoutMinutes must be between 1 and 110 (the orchestrator job allows 120)." }
    if ([string]::IsNullOrWhiteSpace($workRoot)) {
        # Under Program Files, where standard users cannot create folders: the verified copy runs elevated from there.
        $workRoot = Join-Path -Path $env:ProgramFiles -ChildPath 'SmartM365\FreeBusy\Work'
    }
    # The report and transcript hold the hybrid configuration: they stay on a local disk only.
    if ($workRoot.StartsWith('\\') -or -not [System.IO.Path]::IsPathRooted($workRoot)) {
        throw "FreeBusyWorkFolderPath must be a local absolute path, not a share: $workRoot"
    }
    $thumbprint = [string](Get-SmartM365EffectiveConfigValue -Name 'Thumbprint' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($thumbprint)) { $thumbprint = [string](Get-SmartM365EffectiveConfigValue -Name 'Thumb' -DefaultValue '') }

    return [pscustomobject]@{
        TimeoutMinutes        = $timeoutMinutes
        WorkFolderPath        = $workRoot
        AppId                 = [string](Get-SmartM365EffectiveConfigValue -Name 'AppId' -DefaultValue '')
        CertificateThumbprint = $thumbprint
        Organization          = [string](Get-SmartM365EffectiveConfigValue -Name 'OrgDomain' -DefaultValue '')
    }
}

function Test-FreeBusyCheckerIntegrity {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ExpectedSha256
    )

    $result = [pscustomobject]@{ IsValid = $false; Reason = $null; Sha256 = $null; SignatureStatus = $null; Signer = $null }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.Reason = "Free/Busy Checker script not found: $Path"
        return $result
    }

    $result.Sha256 = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($result.Sha256 -ne $ExpectedSha256.ToUpperInvariant()) {
        $result.Reason = "SHA256 mismatch: expected $($ExpectedSha256.ToUpperInvariant()), found $($result.Sha256)."
        return $result
    }

    $signature = Get-AuthenticodeSignature -FilePath $Path
    $result.SignatureStatus = [string]$signature.Status
    if ($signature.SignerCertificate) { $result.Signer = [string]$signature.SignerCertificate.Subject }
    if ([string]$signature.Status -ne 'Valid' -or -not $signature.SignerCertificate) {
        $result.Reason = "Authenticode signature is not valid (status $($signature.Status))."
        return $result
    }
    $subject = [string]$signature.SignerCertificate.Subject
    $issuer = [string]$signature.SignerCertificate.Issuer
    if ($subject -notmatch '(^|,\s*)CN=Microsoft Corporation(,|$)' -or $subject -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)' -or $issuer -notmatch '(^|,\s*)O=Microsoft Corporation(,|$)') {
        $result.Reason = "Authenticode signer is not Microsoft Corporation: $subject"
        return $result
    }

    $result.IsValid = $true
    return $result
}

function Test-FreeBusyRunnerIntegrity {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$ExpectedSha256,
        [AllowNull()][string]$CollectorSignerThumbprint
    )

    # The runner copy must be the repository runner; when the collector is signed, the runner must be signed by the same certificate.
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return "Free/Busy runner not found: $Path" }
    $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if ($hash -ne $ExpectedSha256) { return "Free/Busy runner SHA256 differs from the repository runner ($hash)." }
    if (-not [string]::IsNullOrWhiteSpace($CollectorSignerThumbprint)) {
        $signature = Get-AuthenticodeSignature -FilePath $Path
        if ([string]$signature.Status -ne 'Valid' -or -not $signature.SignerCertificate -or [string]$signature.SignerCertificate.Thumbprint -ne $CollectorSignerThumbprint) {
            return "Free/Busy runner is not signed by the collector's certificate (status $($signature.Status))."
        }
    }
    return $null
}

function Test-FreeBusyElevation {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Same work-folder checks as the Exchange infrastructure collector (copied verbatim; a test keeps them identical).
function Get-HealthCheckerUnsafeAclEntry {
    param([Parameter(Mandatory = $true)][System.Security.AccessControl.FileSystemSecurity]$Acl)

    # SYSTEM, Administrators, TrustedInstaller; CREATOR OWNER only grants rights to whoever creates a child.
    $trustedSids = @('S-1-5-18', 'S-1-5-32-544', 'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464')
    $writeRights = [System.Security.AccessControl.FileSystemRights]'WriteData, AppendData, WriteExtendedAttributes, DeleteSubdirectoriesAndFiles, WriteAttributes, Delete, ChangePermissions, TakeOwnership'
    $unsafe = New-Object System.Collections.Generic.List[string]
    $owner = [string]$Acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value
    if ($owner -notin $trustedSids) { $unsafe.Add("owner $owner") }
    foreach ($rule in @($Acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))) {
        $sid = [string]$rule.IdentityReference.Value
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow -or $sid -in $trustedSids -or $sid -eq 'S-1-3-0') { continue }
        # Generic rights (GENERIC_ALL 0x10000000, GENERIC_WRITE 0x40000000) are not mapped by .NET: test them too.
        $mask = [int64][int]$rule.FileSystemRights -band 4294967295
        if (($mask -band [int64]$writeRights) -ne 0 -or ($mask -band 0x50000000) -ne 0) { $unsafe.Add(('{0} {1}' -f $sid, $rule.FileSystemRights)) }
    }
    return $unsafe.ToArray()
}

function Get-HealthCheckerWorkFolderAncestor {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string[]]$TrustedBase = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:SystemRoot)
    )

    # Parents below the first system folder, nearest first; system folders and the drive root are trusted.
    $bases = @($TrustedBase | Where-Object { $_ } | ForEach-Object { [System.IO.Path]::GetFullPath($_).TrimEnd('\') })
    $ancestors = New-Object System.Collections.Generic.List[string]
    $current = Split-Path -Path ([System.IO.Path]::GetFullPath($Path).TrimEnd('\')) -Parent
    while ($current) {
        $normalized = $current.TrimEnd('\')
        if ($normalized.Length -le 2 -or $normalized -in $bases) { break }
        $ancestors.Add($normalized)
        $current = Split-Path -Path $normalized -Parent
    }
    return $ancestors.ToArray()
}

function Set-HealthCheckerWorkFolderAclRule {
    param([Parameter(Mandatory)][System.Security.AccessControl.DirectorySecurity]$Acl)

    # Administrators own the work folder; only SYSTEM and Administrators have access, inheritance removed.
    $Acl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')))
    $Acl.SetAccessRuleProtection($true, $false)
    foreach ($existingRule in @($Acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))) {
        [void]$Acl.RemoveAccessRuleSpecific($existingRule)
    }
    foreach ($sid in @('S-1-5-18', 'S-1-5-32-544')) {
        $identity = New-Object System.Security.Principal.SecurityIdentifier($sid)
        $Acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($identity, 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow')))
    }
    return $Acl
}

function Initialize-HealthCheckerWorkFolder {
    param([Parameter(Mandatory)][string]$WorkRoot)

    # The verified Health Checker copy runs elevated from here: only administrators may change this folder or its parents.
    $ancestors = @(Get-HealthCheckerWorkFolderAncestor -Path $WorkRoot)
    [array]::Reverse($ancestors)
    foreach ($ancestor in $ancestors) {
        if (-not (Test-Path -LiteralPath $ancestor -PathType Container)) {
            New-Item -ItemType Directory -Path $ancestor -Force | Out-Null
            $ancestorAcl = Get-Acl -LiteralPath $ancestor
            $ancestorAcl.SetOwner((New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')))
            Set-Acl -LiteralPath $ancestor -AclObject $ancestorAcl
        }
        $unsafe = @(Get-HealthCheckerUnsafeAclEntry -Acl (Get-Acl -LiteralPath $ancestor))
        if ($unsafe.Count -gt 0) {
            throw ("A parent of the Health Checker work folder is not restricted to administrators ({0}): {1}" -f ($unsafe -join '; '), $ancestor)
        }
    }
    if (-not (Test-Path -LiteralPath $WorkRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $WorkRoot -Force | Out-Null
    }
    Set-Acl -LiteralPath $WorkRoot -AclObject (Set-HealthCheckerWorkFolderAclRule -Acl (Get-Acl -LiteralPath $WorkRoot))
}

function Invoke-FreeBusyRunnerProcess {
    param(
        [Parameter(Mandatory)][string]$RunnerPath,
        [Parameter(Mandatory)][string]$FreeBusyCheckerPath,
        [Parameter(Mandatory)][string]$StatusPath,
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [Parameter(Mandatory)][object]$Settings,
        [string]$PowerShellPath = (Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe')
    )

    $workFolder = Split-Path -Path $RunnerPath -Parent
    $consolePath = Join-Path -Path $workFolder -ChildPath 'FreeBusyRunner-Console.txt'
    $errorPath = Join-Path -Path $workFolder -ChildPath 'FreeBusyRunner-Errors.txt'
    $arguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $RunnerPath), '-FreeBusyCheckerPath', ('"{0}"' -f $FreeBusyCheckerPath), '-ExpectedSha256', $FreeBusyCheckerExpectedSha256, '-StatusPath', ('"{0}"' -f $StatusPath))
    $result = [pscustomobject]@{ Status = $null; Message = $null; ExitCode = $null; DurationSeconds = $null }
    $started = Get-Date
    # The app identifiers reach the runner through its environment, not its command line.
    $environment = @{ SMARTM365_FREEBUSY_APP_ID = $Settings.AppId; SMARTM365_FREEBUSY_CERT_THUMBPRINT = $Settings.CertificateThumbprint; SMARTM365_FREEBUSY_ORGANIZATION = $Settings.Organization }
    try {
        foreach ($name in $environment.Keys) { [Environment]::SetEnvironmentVariable($name, [string]$environment[$name], 'Process') }
        $process = Start-Process -FilePath $PowerShellPath -ArgumentList $arguments -PassThru -WindowStyle Hidden -WorkingDirectory $workFolder -RedirectStandardOutput $consolePath -RedirectStandardError $errorPath -ErrorAction Stop
        $null = $process.Handle
    }
    catch {
        $result.Status = 'LaunchFailed'
        $result.Message = "Free/Busy runner could not be started: $($_.Exception.Message)"
        $result.DurationSeconds = 0
        return $result
    }
    finally {
        foreach ($name in $environment.Keys) { Remove-Item -LiteralPath "Env:\$name" -ErrorAction SilentlyContinue }
    }

    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        & taskkill.exe /PID $process.Id /T /F 2>&1 | Out-Null
        try { [void]$process.WaitForExit(30000) } catch { Write-Verbose "Stopped Free/Busy runner already gone: $($_.Exception.Message)" }
        $result.Status = 'Timeout'
        $result.Message = "Free/Busy Checker exceeded the $([math]::Round($TimeoutSeconds / 60, 1)) minute limit and was stopped."
        $result.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
        return $result
    }
    $result.ExitCode = $process.ExitCode
    $result.DurationSeconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 0)
    $result.Status = 'Exited'
    return $result
}

function Read-FreeBusyRunnerStatus {
    param([Parameter(Mandatory)][string]$Path)

    # Written by the verified runner in the protected work folder; plain JSON, nothing is deserialized into objects.
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $value = [System.IO.File]::ReadAllText($Path) | ConvertFrom-Json
    $notAuthorized = @()
    if ($value.PSObject.Properties['NotAuthorizedCmdlets'] -and $null -ne $value.NotAuthorizedCmdlets) {
        $notAuthorized = @($value.NotAuthorizedCmdlets | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^(Get|Test)-EO[A-Za-z]+$' })
    }
    $text = { param($Name) if ($value.PSObject.Properties[$Name] -and $null -ne $value.$Name) { [string]$value.$Name } else { $null } }
    return [pscustomobject]@{
        Stage                 = & $text 'Stage'
        RunStatus             = & $text 'RunStatus'
        Message               = & $text 'Message'
        ExchangeOnline        = & $text 'ExchangeOnline'
        ExchangeOnlineMessage = & $text 'ExchangeOnlineMessage'
        NotAuthorizedCmdlets  = $notAuthorized
    }
}

function ConvertTo-FreeBusySafeValue {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '' }
    $text = [regex]::Replace($Value, '<[^>]+>', ' ')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    $text = [regex]::Replace($text, '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', '[email]')
    $text = [regex]::Replace($text, '(?i)\b(?:https?|net\.tcp)://[^\s;,]+', '[url]')
    $text = [regex]::Replace($text, '\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b', '[guid]')
    $text = [regex]::Replace($text, '\b[0-9a-fA-F]{40}\b', '[thumbprint]')
    $text = [regex]::Replace($text, '\b(?:\d{1,3}\.){3}\d{1,3}\b', '[ip]')
    $text = [regex]::Replace($text, '(?i)\b(?:CN|OU|O|DC|L|S|C)=[^,;]+', '[subject]')
    # Host and domain names (labels with at least one letter); versions such as 15.1.2507 are kept.
    $text = [regex]::Replace($text, '\b(?=[A-Za-z0-9-]*[A-Za-z])[A-Za-z0-9-]+(?:\.(?=[A-Za-z0-9-]*[A-Za-z])[A-Za-z0-9-]+)+\b', '[domain]')
    $text = [regex]::Replace($text, '\s+', ' ').Trim()
    if ($text.Length -gt 1000) { $text = $text.Substring(0, 997) + '...' }
    return $text
}

function Get-FreeBusyMaskedItem {
    # Items whose values are identifiers, names, addresses or endpoints of the tenant or the organization.
    return @('Identity', 'Name', 'Domain Names', 'DomainNames', 'Domains', 'UserName', 'Forest Name', 'ForestName', 'ProxyUrl', 'TokenIssuerUris', 'TokenIssuerUri', 'Application Uri', 'ApplicationUri', 'TarGetApplicationUri', 'TargetApplicationUri', 'TarGetAutoDiscoverEpr', 'TargetAutodiscoverEpr', 'TarGetOwaURL', 'TargetOwaURL', 'TarGetSharingEpr', 'TargetSharingEpr', 'Thumbprint', 'Subject', 'Issuer', 'AccountNamespace', 'Delegation Federation Trust', 'DelegationFederationTrust', 'OrganizationGuid', 'ApplicationIdentifier', 'Realm', 'Token Issuer Metadata EPR', 'Token Issuer EPR', 'TokenIssuerMetadataEpr', 'TokenIssuerEpr', 'InternalUrl', 'ExternalUrl', 'DiscoveryEndpoint', 'OnPremiseDiscoveryEndpoint', 'TargetAddressDomains', 'TarGetAddressDomains', 'TargetServiceEpr', 'MetadataUrl', 'AuthMetadataUrl', 'Service Instance', 'HybridDomains', 'InboundConnector', 'OutboundConnector', 'OrganizationName', 'LinkedMailbox', 'Mailbox', 'CertificateThumbprint', 'CurrentCertificateThumbprint', 'NextCertificateThumbprint', 'PreviousCertificateThumbprint', 'ServiceName')
}

function ConvertFrom-FreeBusyCheckerHtml {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Html,
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$FreeBusyCheckerVersion,
        [Parameter(Mandatory)][string]$GenerationTime
    )

    # Text parsing only: the report is never loaded as markup or deserialized.
    $sectionPattern = '(?is)<th[^>]*>\s*(?:<b>)?\s*Exchange (On Premise|Online) (DAuth|OAuth) Configuration\s*(?:</b>)?\s*</th>'
    $summaryPattern = '(?is)<th[^>]*>\s*(?:<b>)?\s*Summary\s*-\s*(.+?)\s*(?:</b>)?\s*</th>'
    $statusMap = @{ green = 'OK'; red = 'Error'; yellow = 'Warning'; orange = 'Warning' }
    $masked = Get-FreeBusyMaskedItem
    $findings = New-Object 'System.Collections.Generic.List[object]'
    $rowNumber = 0
    $sections = @([regex]::Matches($Html, $sectionPattern))
    for ($s = 0; $s -lt $sections.Count; $s++) {
        $side = if ($sections[$s].Groups[1].Value -match 'Online') { 'ExchangeOnline' } else { 'ExchangeOnPremises' }
        $auth = $sections[$s].Groups[2].Value
        $sectionStart = $sections[$s].Index + $sections[$s].Length
        $sectionEnd = if ($s + 1 -lt $sections.Count) { $sections[$s + 1].Index } else { $Html.Length }
        $sectionText = $Html.Substring($sectionStart, $sectionEnd - $sectionStart)
        $summaries = @([regex]::Matches($sectionText, $summaryPattern))
        for ($c = 0; $c -lt $summaries.Count; $c++) {
            $category = ConvertTo-FreeBusySafeValue -Value ($summaries[$c].Groups[1].Value -replace '^(?i)On[ -]Premises?\s+', '')
            $blockStart = $summaries[$c].Index + $summaries[$c].Length
            $blockEnd = if ($c + 1 -lt $summaries.Count) { $summaries[$c + 1].Index } else { $sectionText.Length }
            $block = $sectionText.Substring($blockStart, $blockEnd - $blockStart)
            # The first cell repeats the command line (with mailbox and domain names): results are in the following cells.
            $firstCellEnd = $block.IndexOf('</td>', [StringComparison]::OrdinalIgnoreCase)
            if ($firstCellEnd -ge 0) { $block = $block.Substring($firstCellEnd + 5) }
            foreach ($div in @([regex]::Matches($block, '(?is)<div([^>]*)>(.*?)(?=</div>|<div)'))) {
                $attributes = $div.Groups[1].Value
                # Page headings between tables are not results.
                if ($attributes -match "(?i)class\s*=\s*['""]?Black") { continue }
                $inner = $div.Groups[2].Value
                $label = ''
                $valueText = $inner
                $labelMatch = [regex]::Match($inner, '(?is)^\s*<b>\s*(.*?)\s*</b>(.*)$')
                if ($labelMatch.Success -and $labelMatch.Groups[1].Value -match ':\s*$') {
                    $label = ConvertTo-FreeBusySafeValue -Value ($labelMatch.Groups[1].Value -replace ':\s*$', '')
                    $valueText = $labelMatch.Groups[2].Value
                }
                $value = ConvertTo-FreeBusySafeValue -Value $valueText
                if ([string]::IsNullOrWhiteSpace($label) -and [string]::IsNullOrWhiteSpace($value)) { continue }
                if ($value -match '^=+$') { continue }
                $color = ''
                $classMatch = [regex]::Match($attributes, "(?i)class\s*=\s*['""]?(green|red|yellow|orange)")
                $styleMatch = [regex]::Match($inner + ' ' + $attributes, "(?i)color\s*:\s*['""]?\s*(green|red|yellow|orange)\b")
                if ($classMatch.Success) { $color = $classMatch.Groups[1].Value.ToLowerInvariant() }
                elseif ($styleMatch.Success) { $color = $styleMatch.Groups[1].Value.ToLowerInvariant() }
                $status = if ($statusMap.ContainsKey($color)) { $statusMap[$color] } else { 'Info' }
                $isMasked = $label -in $masked
                $rowNumber++
                $findings.Add([pscustomobject]@{
                    ServerName             = $ServerName
                    Side                   = $side
                    Auth                   = $auth
                    Category               = $category
                    RowNumber              = $rowNumber
                    Item                   = $label
                    Status                 = $status
                    Value                  = $(if ($isMasked) { '[masked]' } else { $value })
                    ValueMasked            = $isMasked
                    FreeBusyCheckerVersion = $FreeBusyCheckerVersion
                    GenerationTime         = $GenerationTime
                }) | Out-Null
            }
        }
    }
    return $findings.ToArray()
}

function Add-FreeBusyNotAuthorizedFinding {
    param(
        [AllowEmptyCollection()][object[]]$Findings = @(),
        [AllowEmptyCollection()][string[]]$NotAuthorizedCmdlets = @(),
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$FreeBusyCheckerVersion,
        [Parameter(Mandatory)][string]$GenerationTime
    )

    # A cmdlet the app cannot run is reported as NotAuthorized, never as a configuration error.
    $scope = Get-FreeBusyCmdletScope
    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($finding in $Findings) { $rows.Add($finding) | Out-Null }
    $nextRow = if ($rows.Count -gt 0) { [int](@($rows | ForEach-Object { [int]$_.RowNumber } | Measure-Object -Maximum).Maximum) } else { 0 }
    foreach ($cmdlet in $NotAuthorizedCmdlets) {
        if (-not $scope.ContainsKey($cmdlet)) { continue }
        $target = $scope[$cmdlet]
        $message = "Not available to the tenant app (missing Exchange role: $($target.Roles)); $cmdlet was not run."
        $matched = @($rows | Where-Object { $_.Side -eq 'ExchangeOnline' -and $_.Auth -eq $target.Auth -and $_.Category -eq $target.Category })
        foreach ($row in $matched) {
            $row.Status = 'NotAuthorized'
            $row.Value = $message
            $row.ValueMasked = $false
        }
        if ($matched.Count -eq 0) {
            $nextRow++
            $rows.Add([pscustomobject]@{ ServerName = $ServerName; Side = 'ExchangeOnline'; Auth = $target.Auth; Category = $target.Category; RowNumber = $nextRow; Item = $cmdlet; Status = 'NotAuthorized'; Value = $message; ValueMasked = $false; FreeBusyCheckerVersion = $FreeBusyCheckerVersion; GenerationTime = $GenerationTime }) | Out-Null
        }
    }
    return $rows.ToArray()
}

function New-FreeBusySummaryRow {
    param(
        [AllowEmptyCollection()][object[]]$Findings = @(),
        [Parameter(Mandatory)][string]$RunStatus,
        [AllowNull()][string]$RunMessage,
        [AllowNull()][object]$RunnerStatus,
        [AllowNull()][object]$DurationSeconds,
        [Parameter(Mandatory)][string]$ServerName,
        [Parameter(Mandatory)][string]$FreeBusyCheckerVersion,
        [Parameter(Mandatory)][string]$GenerationTime
    )

    $scope = Get-FreeBusyCmdletScope
    $blocking = $RunStatus -notin @('Completed', 'EXONotConnected', 'EXONotAuthorized')
    $notAuthorized = if ($RunnerStatus) { @($RunnerStatus.NotAuthorizedCmdlets) } else { @() }
    $exchangeOnline = if ($RunnerStatus) { [string]$RunnerStatus.ExchangeOnline } else { 'NotConnected' }
    $rows = foreach ($side in @('ExchangeOnPremises', 'ExchangeOnline')) {
        foreach ($auth in @('DAuth', 'OAuth')) {
            $sideRows = @($Findings | Where-Object { $_.Side -eq $side -and $_.Auth -eq $auth })
            $sideCmdlets = @($notAuthorized | Where-Object { $side -eq 'ExchangeOnline' -and $scope.ContainsKey($_) -and $scope[$_].Auth -eq $auth })
            $sideStatus = if ($blocking) { 'NotRun' }
                elseif ($side -eq 'ExchangeOnline' -and $exchangeOnline -eq 'NotConnected') { 'NotConnected' }
                elseif ($sideCmdlets.Count -gt 0) { 'NotAuthorizedPartial' }
                elseif ($sideRows.Count -gt 0) { 'Completed' }
                else { 'NotRun' }
            [pscustomobject]@{
                ServerName             = $ServerName
                Side                   = $side
                Auth                   = $auth
                RunStatus              = $RunStatus
                SideStatus             = $sideStatus
                RunMessage             = $RunMessage
                NotAuthorizedCmdlets   = ($sideCmdlets -join '; ')
                MissingRoles           = (@($sideCmdlets | ForEach-Object { $scope[$_].Roles } | Select-Object -Unique) -join '; ')
                DurationSeconds        = $DurationSeconds
                FreeBusyCheckerVersion = $FreeBusyCheckerVersion
                GenerationTime         = $GenerationTime
                OkCount                = $(if ($sideStatus -eq 'NotRun' -or $sideStatus -eq 'NotConnected') { $null } else { @($sideRows | Where-Object Status -eq 'OK').Count })
                WarningCount           = $(if ($sideStatus -eq 'NotRun' -or $sideStatus -eq 'NotConnected') { $null } else { @($sideRows | Where-Object Status -eq 'Warning').Count })
                ErrorCount             = $(if ($sideStatus -eq 'NotRun' -or $sideStatus -eq 'NotConnected') { $null } else { @($sideRows | Where-Object Status -eq 'Error').Count })
                InfoCount              = $(if ($sideStatus -eq 'NotRun' -or $sideStatus -eq 'NotConnected') { $null } else { @($sideRows | Where-Object Status -eq 'Info').Count })
                NotAuthorizedCount     = $(if ($sideStatus -eq 'NotRun' -or $sideStatus -eq 'NotConnected') { $null } else { @($sideRows | Where-Object Status -eq 'NotAuthorized').Count })
            }
        }
    }
    return @($rows)
}

function Resolve-FreeBusyRunStatus {
    param(
        [AllowNull()][object]$Process,
        [AllowNull()][object]$RunnerStatus,
        [bool]$HtmlFound,
        [bool]$ParseFailed
    )

    if ($Process -and $Process.Status -in @('LaunchFailed', 'Timeout')) { return $Process.Status }
    if ($RunnerStatus -and -not [string]::IsNullOrWhiteSpace($RunnerStatus.RunStatus)) { return $RunnerStatus.RunStatus }
    if ($ParseFailed) { return 'ParseFailed' }
    if (-not $HtmlFound) { return 'NoOutput' }
    if ($RunnerStatus -and $RunnerStatus.ExchangeOnline -eq 'NotConnected') { return 'EXONotConnected' }
    if ($RunnerStatus -and @($RunnerStatus.NotAuthorizedCmdlets).Count -gt 0) { return 'EXONotAuthorized' }
    return 'Completed'
}

function ConvertTo-FreeBusyReportHtml {
    param(
        [AllowEmptyCollection()][object[]]$Summary = @(),
        [AllowEmptyCollection()][object[]]$Findings = @(),
        [Parameter(Mandatory)][string]$RunStatus,
        [AllowNull()][string]$RunMessage
    )

    $statusColor = if ($RunStatus -eq 'Completed') { '#166534' } elseif ($RunStatus -in @('EXONotConnected', 'EXONotAuthorized')) { '#9a3412' } else { '#b91c1c' }
    $summaryRows = foreach ($row in $Summary) {
        $sideColor = switch ([string]$row.SideStatus) { 'Completed' { '#166534'; break } 'NotRun' { '#64748b'; break } default { '#9a3412' } }
        "<tr><td>$(Format-HtmlValue $row.Side)</td><td>$(Format-HtmlValue $row.Auth)</td><td style='font-weight:700;color:$sideColor;'>$(Format-HtmlValue $row.SideStatus)</td><td style='text-align:right;'>$(Format-HtmlValue $row.ErrorCount)</td><td style='text-align:right;'>$(Format-HtmlValue $row.WarningCount)</td><td style='text-align:right;'>$(Format-HtmlValue $row.NotAuthorizedCount)</td><td>$(Format-HtmlValue $row.MissingRoles)</td></tr>"
    }
    $issues = @($Findings | Where-Object { $_.Status -in @('Error', 'NotAuthorized', 'Warning') } | Sort-Object @{ Expression = { switch ($_.Status) { 'Error' { 0 } 'NotAuthorized' { 1 } default { 2 } } } }, Side, Auth, RowNumber)
    $shown = @($issues | Select-Object -First 100)
    $issueRows = foreach ($row in $shown) {
        $color = switch ([string]$row.Status) { 'Error' { '#b91c1c'; break } 'NotAuthorized' { '#9a3412'; break } default { '#a16207' } }
        "<tr><td style='font-weight:700;color:$color;'>$(Format-HtmlValue $row.Status)</td><td>$(Format-HtmlValue $row.Side)</td><td>$(Format-HtmlValue $row.Auth)</td><td>$(Format-HtmlValue $row.Category)</td><td>$(Format-HtmlValue $row.Item)</td><td>$(Format-HtmlValue $row.Value)</td></tr>"
    }
    $issuesBlock = if ($shown.Count -gt 0) {
        $note = if ($issues.Count -gt $shown.Count) { "<p style='font-size:12px;color:#64748b;'>Showing the first $($shown.Count) of $($issues.Count) results; all are in Exchange_Hybrid_FreeBusy_Findings.csv.</p>" } else { '' }
        "<h2 style='font-size:17px;'>Errors, not authorized checks and warnings</h2><table cellpadding='6' cellspacing='0' style='width:100%;border-collapse:collapse;font-size:12px;border:1px solid #e5e7eb;'><tr style='background:#f8fafc;'><th align='left'>Status</th><th align='left'>Side</th><th align='left'>Auth</th><th align='left'>Check</th><th align='left'>Item</th><th align='left'>Value</th></tr>$($issueRows -join "`r`n")</table>$note"
    } else { "<p>No error, warning or not authorized check was reported.</p>" }
    $message = if ([string]::IsNullOrWhiteSpace($RunMessage)) { '' } else { "<p>$(Format-HtmlValue $RunMessage)</p>" }

    return @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Exchange Hybrid Free/Busy</title></head>
<body style="margin:0;padding:24px;background:#f6f8fb;font-family:Segoe UI,Arial,sans-serif;color:#0f172a;">
<div style="max-width:1100px;margin:0 auto;background:#ffffff;border:1px solid #dbe3ef;border-radius:10px;padding:24px;">
<div style="font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#0f766e;font-weight:700;">Smart365 Exchange Hybrid</div>
<h1 style="font-size:24px;margin:6px 0 12px 0;">Exchange hybrid free/busy configuration</h1>
<p>Microsoft CSS-Exchange Free/Busy Checker $(Format-HtmlValue $FreeBusyCheckerVersion) - run status <strong style="color:$statusColor;">$(Format-HtmlValue $RunStatus)</strong>. RunId $(Format-HtmlValue $RunId).</p>
$message
<p style="font-size:12px;color:#64748b;">Sides that were not run or not connected show N/A: their state is not measured. Values that can reveal names, addresses, URLs, certificates or identifiers are masked. Not authorized checks are Exchange Online cmdlets the tenant app cannot run; they are not configuration errors.</p>
<table cellpadding="6" cellspacing="0" style="width:100%;border-collapse:collapse;font-size:12px;border:1px solid #e5e7eb;">
<tr style="background:#f8fafc;"><th align="left">Side</th><th align="left">Auth</th><th align="left">Status</th><th align="right">Errors</th><th align="right">Warnings</th><th align="right">Not authorized</th><th align="left">Missing roles</th></tr>
$($summaryRows -join "`r`n")
</table>
$issuesBlock
</div></body></html>
"@
}

function Clear-FreeBusyWorkFolder {
    param([Parameter(Mandatory)][string]$WorkRoot)

    # Only the last run is kept: the previous copy, report and transcript are removed first.
    foreach ($item in @(Get-ChildItem -LiteralPath $WorkRoot -Force -ErrorAction SilentlyContinue)) {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
    }
}

$StartTime = Get-Date
$script:CompletionStatus = 'Auto'
try {
    Ensure-Directory -Path $OutputFolder
    Ensure-Directory -Path $LogFolder
    Start-SmartM365RunTranscript -Path $TranscriptFile -Append | Out-Null
    $script:TranscriptStarted = $true
    Write-Log "Starting $ScriptName v$ScriptVersion. RunId: $RunId"
    Write-Log "Output folder: $OutputFolder"

    Import-SmartM365CoreModule
    Start-CoreSmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath ([string](Get-SmartM365EffectiveConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')) -ScopeParameters @{ FreeBusyCheckerVersion = $FreeBusyCheckerVersion }
    Invoke-CoreSmartM365Preflight -ScriptName $ScriptName -OutputPaths @($OutputFolder, $LogFolder) | Out-Null

    $serverName = [string]$env:COMPUTERNAME
    $generationTime = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz')
    $runnerStatus = $null
    $process = $null
    $findings = @()
    $runStatus = $null
    $runMessage = $null
    $htmlFound = $false
    $parseFailed = $false

    $settings = $null
    try { $settings = Get-FreeBusySetting }
    catch { $runStatus = 'ConfigurationInvalid'; $runMessage = "Free/Busy Checker was not run: invalid configuration. $($_.Exception.Message)" }

    if (-not $runStatus -and -not (Test-FreeBusyElevation)) {
        $runStatus = 'NotElevated'
        $runMessage = 'Free/Busy Checker was not run: it requires an elevated (administrator) session on the Exchange server.'
    }
    if (-not $runStatus) {
        try {
            Initialize-HealthCheckerWorkFolder -WorkRoot $settings.WorkFolderPath
            Clear-FreeBusyWorkFolder -WorkRoot $settings.WorkFolderPath
        }
        catch {
            $runStatus = 'UnsafeWorkFolder'
            $runMessage = "Free/Busy Checker was not run: the work folder is not restricted to administrators. $($_.Exception.Message)"
        }
    }
    if (-not $runStatus) {
        $sourceChecker = Join-Path -Path $PSScriptRoot -ChildPath $FreeBusyCheckerRelativePath
        $sourceRunner = Join-Path -Path $PSScriptRoot -ChildPath $FreeBusyRunnerFileName
        $checkerCopy = Join-Path -Path $settings.WorkFolderPath -ChildPath 'FreeBusyChecker.ps1'
        $runnerCopy = Join-Path -Path $settings.WorkFolderPath -ChildPath $FreeBusyRunnerFileName
        $statusPath = Join-Path -Path $settings.WorkFolderPath -ChildPath 'FreeBusyRunner-Status.json'
        $collectorSignature = Get-AuthenticodeSignature -FilePath $PSCommandPath
        $collectorSigner = if ([string]$collectorSignature.Status -eq 'Valid' -and $collectorSignature.SignerCertificate) { [string]$collectorSignature.SignerCertificate.Thumbprint } else { $null }
        $runnerSha256 = if (Test-Path -LiteralPath $sourceRunner -PathType Leaf) { (Get-FileHash -LiteralPath $sourceRunner -Algorithm SHA256).Hash } else { '' }
        $integrity = Test-FreeBusyCheckerIntegrity -Path $sourceChecker -ExpectedSha256 $FreeBusyCheckerExpectedSha256
        if ($integrity.IsValid) {
            [System.IO.File]::WriteAllBytes($checkerCopy, [System.IO.File]::ReadAllBytes($sourceChecker))
            if ($runnerSha256) { [System.IO.File]::WriteAllBytes($runnerCopy, [System.IO.File]::ReadAllBytes($sourceRunner)) }
            # Checked again on the copies immediately before the start, so the files that run are the ones just verified.
            $integrity = Test-FreeBusyCheckerIntegrity -Path $checkerCopy -ExpectedSha256 $FreeBusyCheckerExpectedSha256
        }
        $runnerProblem = if ($runnerSha256) { Test-FreeBusyRunnerIntegrity -Path $runnerCopy -ExpectedSha256 $runnerSha256 -CollectorSignerThumbprint $collectorSigner } else { "Free/Busy runner not found: $sourceRunner" }
        if (-not $integrity.IsValid) {
            $runStatus = 'IntegrityFailed'
            $runMessage = "Free/Busy Checker was not run: $($integrity.Reason)"
        }
        elseif ($runnerProblem) {
            $runStatus = 'IntegrityFailed'
            $runMessage = "Free/Busy Checker was not run: $runnerProblem"
        }
        else {
            Write-Log ("Free/Busy Checker {0} and its runner verified. Time limit: {1} min." -f $FreeBusyCheckerVersion, $settings.TimeoutMinutes)
            $process = Invoke-FreeBusyRunnerProcess -RunnerPath $runnerCopy -FreeBusyCheckerPath $checkerCopy -StatusPath $statusPath -TimeoutSeconds ($settings.TimeoutMinutes * 60) -Settings $settings
            $runnerStatus = Read-FreeBusyRunnerStatus -Path $statusPath
            $htmlFile = @(Get-ChildItem -LiteralPath $settings.WorkFolderPath -Filter 'FBCheckerOutput_*.html' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
            if ($htmlFile.Count -gt 0) {
                $htmlFound = $true
                try {
                    $findings = @(ConvertFrom-FreeBusyCheckerHtml -Html ([System.IO.File]::ReadAllText($htmlFile[0].FullName)) -ServerName $serverName -FreeBusyCheckerVersion $FreeBusyCheckerVersion -GenerationTime $generationTime)
                    if ($runnerStatus -and @($runnerStatus.NotAuthorizedCmdlets).Count -gt 0) {
                        $findings = @(Add-FreeBusyNotAuthorizedFinding -Findings $findings -NotAuthorizedCmdlets @($runnerStatus.NotAuthorizedCmdlets) -ServerName $serverName -FreeBusyCheckerVersion $FreeBusyCheckerVersion -GenerationTime $generationTime)
                    }
                }
                catch {
                    $parseFailed = $true
                    $findings = @()
                    $runMessage = "The Free/Busy Checker report could not be read: $($_.Exception.Message)"
                }
            }
            $runStatus = Resolve-FreeBusyRunStatus -Process $process -RunnerStatus $runnerStatus -HtmlFound $htmlFound -ParseFailed $parseFailed
            if (-not $runMessage) {
                $runMessage = if ($process.Status -in @('LaunchFailed', 'Timeout')) { $process.Message }
                    elseif ($runnerStatus -and $runnerStatus.Message) { $runnerStatus.Message }
                    elseif ($runnerStatus -and $runnerStatus.ExchangeOnlineMessage) { $runnerStatus.ExchangeOnlineMessage }
                    elseif (-not $htmlFound) { "Free/Busy Checker produced no report (runner exit code $($process.ExitCode), stage $(if ($runnerStatus) { $runnerStatus.Stage } else { 'unknown' }))." }
                    else { $null }
            }
        }
    }

    if ($runStatus -ne 'Completed') {
        Write-Log ("Free/Busy Checker run status {0}. {1} Raw output, if any, stays in the local work folder." -f $runStatus, $runMessage) 'WARN'
    }
    $summary = @(New-FreeBusySummaryRow -Findings $findings -RunStatus $runStatus -RunMessage $runMessage -RunnerStatus $runnerStatus -DurationSeconds $(if ($process) { $process.DurationSeconds } else { $null }) -ServerName $serverName -FreeBusyCheckerVersion $FreeBusyCheckerVersion -GenerationTime $generationTime)
    Write-Log ("Free/Busy Checker: {0} finding(s); {1} error(s); {2} not authorized check(s)." -f $findings.Count, @($findings | Where-Object Status -eq 'Error').Count, @($findings | Where-Object Status -eq 'NotAuthorized').Count)

    $summaryPath = Join-Path $OutputFolder 'Exchange_Hybrid_FreeBusy_Summary.csv'
    Export-FreeBusyCsv -Data @($summary) -Path $summaryPath -Columns (Get-FreeBusyCsvColumn -Kind Summary)
    $findingsPath = Join-Path $OutputFolder 'Exchange_Hybrid_FreeBusy_Findings.csv'
    Export-FreeBusyCsv -Data @($findings) -Path $findingsPath -Columns (Get-FreeBusyCsvColumn -Kind Findings)
    Write-Log "Free/Busy summary and findings exported to: $summaryPath; $findingsPath"

    $reportPath = Join-Path $OutputFolder 'Exchange_Hybrid_FreeBusy_Report.html'
    Set-Content -LiteralPath $reportPath -Value (ConvertTo-FreeBusyReportHtml -Summary $summary -Findings $findings -RunStatus $runStatus -RunMessage $runMessage) -Encoding UTF8
    Invoke-FreeBusySharePointUpload -LocalFilePath $reportPath
    Save-FreeBusyWeeklyHistory
    Remove-CoreSmartM365TimestampedDirectoriesOlderThan -RootPath $OutputRoot -RetentionDays 7 -ExcludeDirectories @($OutputFolder) -LogFile $LogFile

    Send-FreeBusyHtmlReport -HtmlReportPath $reportPath -Subject "SmartM365 Exchange Hybrid Free/Busy check - $RunId" -Attachments @($summaryPath)
    Write-Log "Free/Busy report email sent."

    $completionStatus = if ($script:FreeBusyWarningCount -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
    $script:CompletionStatus = $completionStatus
    Complete-FreeBusyRun -Status $completionStatus -Started $StartTime -ErrorMessage $null
    if ($completionStatus -eq 'CompletedWithWarnings') { $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(3); exit $smartM365RunGuard.ExitCode }

    $smartM365RunGuard.ReachedTerminal=$true
}
catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_

    $script:CompletionStatus = 'Failed'
    $message = $_.Exception.Message
    try { Write-Log -Level "ERROR" -Message $message } catch { Write-Verbose "Log unavailable: $($_.Exception.Message)" }
    try { Complete-FreeBusyRun -Status 'Failed' -Started $StartTime -ErrorMessage $message } catch { Write-Verbose "Completion unavailable: $($_.Exception.Message)" }
    throw
}
finally {
    $smartM365RunGuard.InCleanup=$true

    # 'Auto' here means the run never reached its end (Ctrl+C or a stopped pipeline): the shared banner shows
    # CANCELLED and the lifecycle finalizer exits with code 130. No receipt or success status was written.
    $bannerStatus = $script:CompletionStatus
    if ($bannerStatus -eq 'Auto') {
        try { Write-Log -Level 'WARN' -Message 'Run interrupted before completion (cancelled): no receipt or success status was written.' } catch { $script:FreeBusyWarningCount++ }
    }
    Write-SmartM365CompletionBanner -Status $bannerStatus -ScriptName $ScriptName -StartedAt $StartTime -WarningCount $script:FreeBusyWarningCount -ErrorCount $script:FreeBusyErrorCount
}
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDIoncNQ/4TzFx3
# /aI8ASxyZ2WDMyTH2fz2vdjAk9WfZ6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDuh02rmm1ftrKLENvgIdq65OhVGNKYSg3XafZhWGey5MA0GCSqG
# SIb3DQEBAQUABIIBgBiT6V/7/0CN+1hIB8xwGWvCmxTCrThW4mv4hbjEiOdeIR4a
# wWIkM5zmPN0aQSa4mpoQ0gEVr74yg38fT+d7CndyQtE1Y6QhpwqL52GuFPbVNrU1
# ypXDm1tJtbn6UiehJqxDS0P1pInFkjQlqgXRMg5/FZMOktekwEt6gxR4U4ta8Q8f
# V8xLIV6Ch1HS2wcKvSh9S9w7htPZh4kOpxjASFr0eoaxFdZgr9sxC4zGhx/BmlZs
# ns05DEBuztUamYEddPsDUwmjozhUr6SHPBJKoh4VUwajLylalVVE6BLn5Q21B5Z2
# VP4P2mGeMp/dfz3hokQoBSFOAv091oWn6e6+uywCuXfuuOu9ViVSAKzjKmfss9a8
# AKx9az49Qxv17Jd94nYU41+M9zTz1grbX/aJoiqzdCUqhCtJxEEKXdBfDj7oNe+W
# HIFlTvzD5olrpkK48VxLN8KL6pTkb3gV8NswFWUs00+YP9HjLVbaFXIcZsyfwJ/3
# F7hY4eWfDwDaW4H0JKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEyMDMy
# NDRaMC8GCSqGSIb3DQEJBDEiBCChn2PpWAGzsQ77bh7p57x1e/xFqhMnGA3OhzSs
# FIuOrzANBgkqhkiG9w0BAQEFAASCAgBlGLI51FrlyRHYAzWIQqo8euoFdzVRYgS3
# 13Jp5dyy3T0aMbBMrtMJmnA2uMXnicWNX6t8orDrVOdbrdmSE9YI0YBuokUbZMg+
# pmmbYqWgKW2CHmeE9f4AexaHeSyfSUpo7OFOjZUjvLU6Izo/vzwwfUZtrD4YgdLi
# gOExIMreqxxzedKsZMnGA57bCvSbikHP0E4O+Fu9F4karFbRzdVaJaufpB24VxuQ
# OQDYNRt8aAJ2Pq3Hd02tW4r6L3YG7VL6dErNTxgH5cEAyIJspx6Gj+4kQevl6FNV
# jhNjzEhGKlZ+MsXHJteFHO2tUNv+RixnrUvvJIuSHuKdCrioYc+wpBxi7w89cFX5
# rzFrP7XSNQoLaLMmvvnNx2LLZ9NKu1vSsE8LDpVF3nIPEfhtj2auWP3vnnpL5GqD
# 2E7GEAXwFo55RcwFB4IbYenXTtHhTUjbDI7Omb6ipJqk0t9NBN4Grv1njytZTftD
# OVZXO3INYsppkKtXTjj8mE4j5vubXNecBvYROoVYhr3omUcP7ckMMlzFjOk2FfXy
# iPrLRqsZ0QUaruPimDuMupjiYcYND4so9JqXZs49FMqpGrMuQI3rwhFTJ8fWwgbF
# nPgNS2vHX/CKZighVspUJQRaoko6o2A+WLQGOyB43sbBPno8Jj9WZ+6aVbpuZQN9
# M6NLquH4Og==
# SIG # End signature block
