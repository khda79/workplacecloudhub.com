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
1.0.12
.REQUIREMENTS
    Windows PowerShell 5.1, elevated, on an Exchange server with the Exchange Management Tools.
    Modules: SmartM365 WindowsPowerShell5 compatibility module; ActiveDirectory (Windows feature RSAT-AD-PowerShell);
    ExchangeOnlineManagement 3.x (installed for the current user by the runner when missing).
    Exchange Online: the tenant app with its certificate; Free/Busy Checker cmdlets not exposed to the app are reported NotAuthorized.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Script Name : SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1
    Version     : 1.0.12
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
. (Join-Path $smartM365LifecycleRoot 'SmartInventory/Common/SmartM365.ReportRenderers.ps1')
. (Join-Path $smartM365LifecycleRoot 'SmartInventory/Common/SmartM365.ReportInputs.ps1')
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
$ScriptVersion = "1.0.12"
$RunId = (Get-Date).ToString("yyyyMMdd-HHmmss")
# Vendored Microsoft CSS-Exchange Free/Busy Checker (see ThirdParty\CSS-Exchange\FreeBusyChecker\PROVENANCE.md).
$FreeBusyCheckerRelativePath = 'ThirdParty\CSS-Exchange\FreeBusyChecker\FreeBusyChecker.ps1'
$FreeBusyCheckerVersion = '26.06.06.1238'
$FreeBusyCheckerExpectedSha256 = '5E9BCF887AB3A06F29208EB0C3CAD87D0B38C30019F18BB8DE27791038EEEC34'
$FreeBusyRunnerFileName = 'SmartM365-Exchange-Hybrid-FreeBusy-Runner.ps1'

$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

function Merge-SmartM365ScriptLocalConfig {
    param([Parameter(Mandatory)][string]$ConfigPath)
    $localConfig = Read-SmartM365ScriptConfiguration -ConfigurationPath $ConfigPath
    $script:SmartM365EffectiveConfig = Merge-SmartM365Configuration -EffectiveConfig $script:SmartM365EffectiveConfig -LocalConfig $localConfig -Mode Overlay
}

Merge-SmartM365ScriptLocalConfig -ConfigPath (Join-Path -Path $PSScriptRoot -ChildPath "$ScriptName.local.json")

function Resolve-SmartM365ConfigTokenValue {
    param([AllowNull()]$Value)
    Resolve-SmartM365ConfigurationTokens -Value $Value -Config $script:SmartM365EffectiveConfig
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
    param([Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)
    Get-SmartM365ConfigurationValue -Name $Name -EffectiveConfig $script:SmartM365EffectiveConfig -DefaultValue $DefaultValue -EffectiveFallback Blank
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
                $minimumVersion = if ($modulePath -like '*WindowsPowerShell5*') { '1.0.76' } else { '1.0.130' }
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

    $policies=Get-Variable SmartM365CollectorUploadPolicies -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    $enabled = if ($null -ne $policies) { [bool]$global:EnableSharePointUpload } else { [bool](Get-SmartM365EffectiveConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false) }
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
        # Callers do not read the record (it is kept in the uploads list); returning it wrote it to the job log.
        return
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

    $reportHtml = Add-SmartM365FreeBusyTemplateMarker (Get-Content -LiteralPath $HtmlReportPath -Raw -ErrorAction Stop)

    $mailParams = @{
        From        = $from
        To          = $to
        Subject     = $Subject
        BodyHtml    = $reportHtml
        Attachments = $existingAttachments
        ReportCadenceKey = 'Exchange.FreeBusy'
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

# Free/Busy Checker checks OAuth only when the Intra Organization Connector is enabled and no -Auth is
# given (-Auth would select a parameter set that lets it update itself): its DAuth checks do not apply.
$FreeBusyDAuthNotApplicableReason = 'Intra Organization Connector enabled: free/busy uses OAuth, DAuth not checked'
function Test-FreeBusyOAuthOnly {
    param([AllowEmptyString()][string]$Html)

    return [regex]::IsMatch($Html, '(?i)Checking OAuth only as Free Busy Lookup is done using OAuth when the Intra Organization Connector is Enabled')
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
        [Parameter(Mandatory)][string]$GenerationTime,
        [switch]$OAuthOnly
    )

    # A cmdlet the app cannot run is reported as NotAuthorized, never as a configuration error.
    $scope = Get-FreeBusyCmdletScope
    $rows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($finding in $Findings) { $rows.Add($finding) | Out-Null }
    $nextRow = if ($rows.Count -gt 0) { [int](@($rows | ForEach-Object { [int]$_.RowNumber } | Measure-Object -Maximum).Maximum) } else { 0 }
    foreach ($cmdlet in $NotAuthorizedCmdlets) {
        if (-not $scope.ContainsKey($cmdlet)) { continue }
        $target = $scope[$cmdlet]
        if ($OAuthOnly -and $target.Auth -eq 'DAuth') { continue }
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
        [Parameter(Mandatory)][string]$GenerationTime,
        [switch]$OAuthOnly
    )

    $scope = Get-FreeBusyCmdletScope
    $blocking = $RunStatus -notin @('Completed', 'EXONotConnected', 'EXONotAuthorized')
    $notAuthorized = @(if ($RunnerStatus) { $RunnerStatus.NotAuthorizedCmdlets })
    $exchangeOnline = if ($RunnerStatus) { [string]$RunnerStatus.ExchangeOnline } else { 'NotConnected' }
    $rows = foreach ($side in @('ExchangeOnPremises', 'ExchangeOnline')) {
        foreach ($auth in @('DAuth', 'OAuth')) {
            $sideRows = @($Findings | Where-Object { $_.Side -eq $side -and $_.Auth -eq $auth })
            # Missing cmdlets and roles are reported only when Free/Busy Checker actually ran.
            $notApplicable = (-not $blocking) -and $OAuthOnly -and $auth -eq 'DAuth'
            # @() outside the if: an if returning one cmdlet yields a scalar, and .Count on it fails under strict mode in 5.1.
            $sideCmdlets = @(if (-not ($blocking -or $notApplicable)) { $notAuthorized | Where-Object { $side -eq 'ExchangeOnline' -and $scope.ContainsKey($_) -and $scope[$_].Auth -eq $auth } })
            $sideStatus = if ($blocking) { 'NotRun' }
                elseif ($side -eq 'ExchangeOnline' -and $exchangeOnline -eq 'NotConnected') { 'NotConnected' }
                elseif ($notApplicable) { 'NotApplicable' }
                elseif ($sideCmdlets.Count -gt 0) { 'NotAuthorizedPartial' }
                elseif ($sideRows.Count -gt 0) { 'Completed' }
                else { 'NotRun' }
            [pscustomobject]@{
                ServerName             = $ServerName
                Side                   = $side
                Auth                   = $auth
                RunStatus              = $RunStatus
                SideStatus             = $sideStatus
                RunMessage             = $(if ($sideStatus -eq 'NotApplicable') { $FreeBusyDAuthNotApplicableReason } else { $RunMessage })
                NotAuthorizedCmdlets   = ($sideCmdlets -join '; ')
                MissingRoles           = (@($sideCmdlets | ForEach-Object { $scope[$_].Roles } | Select-Object -Unique) -join '; ')
                DurationSeconds        = $DurationSeconds
                FreeBusyCheckerVersion = $FreeBusyCheckerVersion
                GenerationTime         = $GenerationTime
                OkCount                = $(if ($sideStatus -in @('NotRun', 'NotConnected', 'NotApplicable')) { $null } else { @($sideRows | Where-Object Status -eq 'OK').Count })
                WarningCount           = $(if ($sideStatus -in @('NotRun', 'NotConnected', 'NotApplicable')) { $null } else { @($sideRows | Where-Object Status -eq 'Warning').Count })
                ErrorCount             = $(if ($sideStatus -in @('NotRun', 'NotConnected', 'NotApplicable')) { $null } else { @($sideRows | Where-Object Status -eq 'Error').Count })
                InfoCount              = $(if ($sideStatus -in @('NotRun', 'NotConnected', 'NotApplicable')) { $null } else { @($sideRows | Where-Object Status -eq 'Info').Count })
                NotAuthorizedCount     = $(if ($sideStatus -in @('NotRun', 'NotConnected', 'NotApplicable')) { $null } else { @($sideRows | Where-Object Status -eq 'NotAuthorized').Count })
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
        [bool]$ParseFailed,
        [switch]$OAuthOnly
    )

    if ($Process -and $Process.Status -in @('LaunchFailed', 'Timeout')) { return $Process.Status }
    if ($RunnerStatus -and -not [string]::IsNullOrWhiteSpace($RunnerStatus.RunStatus)) { return $RunnerStatus.RunStatus }
    if ($ParseFailed) { return 'ParseFailed' }
    if (-not $HtmlFound) { return 'NoOutput' }
    if ($RunnerStatus -and $RunnerStatus.ExchangeOnline -eq 'NotConnected') { return 'EXONotConnected' }
    $scope = Get-FreeBusyCmdletScope
    $applicable = @(@($(if ($RunnerStatus) { $RunnerStatus.NotAuthorizedCmdlets } else { @() })) | Where-Object { $_ -and -not ($OAuthOnly -and $scope.ContainsKey($_) -and $scope[$_].Auth -eq 'DAuth') })
    if ($applicable.Count -gt 0) { return 'EXONotAuthorized' }
    return 'Completed'
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
    $global:EnableSharePointUpload=[bool](Get-SmartM365EffectiveConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false)
    . (Join-Path $smartM365LifecycleRoot 'Modules\SmartM365.Core\SmartM365-SharePointUploadPolicy.ps1')
    Initialize-SmartM365CollectorUploadContext -LocalConfig $script:SmartM365EffectiveConfig -DataAllRoot $OutputRoot -DataLastRoot ([string](Get-SmartM365EffectiveConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')) -LogAllRoot $LogFolder -ExternalActionsAllowed $true
    # The WindowsPowerShell5 Core uploads (weekly history included) read the SharePoint target from globals.
    # Set them from the effective configuration so DATA-ALL publication does not depend on a caller default.
    foreach ($smartM365UploadSetting in @('SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath','AppId','TenantId')) {
        $smartM365UploadValue = [string](Get-SmartM365EffectiveConfigValue -Name $smartM365UploadSetting -DefaultValue '')
        if (-not [string]::IsNullOrWhiteSpace($smartM365UploadValue)) { Set-Variable -Name $smartM365UploadSetting -Scope Global -Value $smartM365UploadValue }
    }
    $smartM365UploadThumbprint = [string](Get-SmartM365EffectiveConfigValue -Name 'Thumbprint' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($smartM365UploadThumbprint)) { $smartM365UploadThumbprint = [string](Get-SmartM365EffectiveConfigValue -Name 'Thumb' -DefaultValue '') }
    if (-not [string]::IsNullOrWhiteSpace($smartM365UploadThumbprint)) { $global:Thumbprint = $smartM365UploadThumbprint }
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
$oauthOnly = $false
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
                    $htmlText = [System.IO.File]::ReadAllText($htmlFile[0].FullName)
                    $oauthOnly = Test-FreeBusyOAuthOnly -Html $htmlText
                    if ($oauthOnly) { Write-Log "Free/Busy Checker checked OAuth only: $FreeBusyDAuthNotApplicableReason." }
                    $findings = @(ConvertFrom-FreeBusyCheckerHtml -Html $htmlText -ServerName $serverName -FreeBusyCheckerVersion $FreeBusyCheckerVersion -GenerationTime $generationTime)
                    if ($runnerStatus -and @($runnerStatus.NotAuthorizedCmdlets).Count -gt 0) {
                        $findings = @(Add-FreeBusyNotAuthorizedFinding -Findings $findings -NotAuthorizedCmdlets @($runnerStatus.NotAuthorizedCmdlets) -ServerName $serverName -FreeBusyCheckerVersion $FreeBusyCheckerVersion -GenerationTime $generationTime -OAuthOnly:$oauthOnly)
                    }
                }
                catch {
                    $parseFailed = $true
                    $findings = @()
                    $runMessage = "The Free/Busy Checker report could not be read: $($_.Exception.Message)"
                }
            }
            $runStatus = Resolve-FreeBusyRunStatus -Process $process -RunnerStatus $runnerStatus -HtmlFound $htmlFound -ParseFailed $parseFailed -OAuthOnly:$oauthOnly
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
    $summary = @(New-FreeBusySummaryRow -Findings $findings -RunStatus $runStatus -RunMessage $runMessage -RunnerStatus $runnerStatus -DurationSeconds $(if ($process) { $process.DurationSeconds } else { $null }) -ServerName $serverName -FreeBusyCheckerVersion $FreeBusyCheckerVersion -GenerationTime $generationTime -OAuthOnly:$oauthOnly)
    Write-Log ("Free/Busy Checker: {0} finding(s); {1} error(s); {2} not authorized check(s)." -f $findings.Count, @($findings | Where-Object Status -eq 'Error').Count, @($findings | Where-Object Status -eq 'NotAuthorized').Count)

    $summaryPath = Join-Path $OutputFolder 'Exchange_Hybrid_FreeBusy_Summary.csv'
    Export-FreeBusyCsv -Data @($summary) -Path $summaryPath -Columns (Get-FreeBusyCsvColumn -Kind Summary)
    $findingsPath = Join-Path $OutputFolder 'Exchange_Hybrid_FreeBusy_Findings.csv'
    Export-FreeBusyCsv -Data @($findings) -Path $findingsPath -Columns (Get-FreeBusyCsvColumn -Kind Findings)
    Write-Log "Free/Busy summary and findings exported to: $summaryPath; $findingsPath"

    $reportLatest = [string](Get-SmartM365EffectiveConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
    if ($reportLatest) {
        Save-SmartM365ReportInputs -Producer 'SmartM365-Exchange-Hybrid-FreeBusy-Inventory' -LatestFolder $reportLatest -TenantKey $global:SmartM365TenantKey -Files @('Exchange_Hybrid_FreeBusy_Summary.csv','Exchange_Hybrid_FreeBusy_Findings.csv') -InputData @{
            RunId=$RunId; RunStatus=$runStatus; RunMessage=$runMessage; CheckerVersion=$FreeBusyCheckerVersion; SummaryFile=(Get-SmartM365ReportFileReference -Path $summaryPath)
        }
    }
    $reportPath = Join-Path $OutputFolder 'Exchange_Hybrid_FreeBusy_Report.html'
    Set-Content -LiteralPath $reportPath -Value (ConvertTo-FreeBusyReportHtml -Summary $summary -Findings $findings -RunStatus $runStatus -RunMessage $runMessage) -Encoding UTF8
    Invoke-FreeBusySharePointUpload -LocalFilePath $reportPath
    Save-FreeBusyWeeklyHistory
    Remove-CoreSmartM365TimestampedDirectoriesOlderThan -RootPath $OutputRoot -RetentionDays 7 -ExcludeDirectories @($OutputFolder) -LogFile $LogFile

    Send-FreeBusyHtmlReport -HtmlReportPath $reportPath -Subject "SmartM365 Exchange Hybrid Free/Busy check - $RunId" -Attachments @($summaryPath)
    if ((Get-Variable SmartM365LastMailDelivery -Scope Global -ValueOnly -ErrorAction SilentlyContinue) -notin @('Skipped','Busy')) { Write-Log "Free/Busy report email accepted for delivery." }

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
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBRl6PLj2Wx1Hsp
# emh7UgTMci53h1Stv4rSlkFxjEWOCaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgUEfY3rd2LGlj
# KiNrN+v90dxtxx/tHRlwGRBejd0HRaEwDQYJKoZIhvcNAQEBBQAEggGARHntFYJ6
# OmvVw4A5G/UeLwvMEzL+dlcotSwKoAPx4NbAhEs8hojR5PQsQp3s3TOR8ax37LDD
# SRD0m6VHuvbYX5/KD3QQDzGworInM5K4SrIFB97PmzZ0iOeFaOVLegoTDX0lcWjK
# 78+Ok/cuRCChdPcRNohoNICu1M2mJgQ2TQ/sWfPm3CwozcNwVJkaDcPHMZQwr4RU
# OEqR/EcECfgx76hYkLG7uxarhbUtQ/IybmhuDwJt0BlVIbaYDvh+4k832+cIp8aE
# 8+5kjbKAXblxGJLpZ0/ExX4AlLj8e/YADiV8aTH3VxeRyzFMS5ZF8aS41aSOyNlw
# VHzXdYt3ZtGcAcnvj22mAnVwN54eY760feubKCcH3PneCBHFSjUZ53VyxYPC1WCz
# fo+Z1TVTcWyQpZ+Pbw9u1C6/XmX9nAEZw7WwLyLtmr6xqIAmmAOnT9+h9V8uUj26
# ERRHr8RriXrwTTtqvB2agZmtsZh/9O4Mg0pWzc/aX5Tmppeazor/qFr2oYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNDUxWjA/BgkqhkiG9w0BCQQxMgQwQKnfyjx7j+/h5whfhcyC353z
# zPFkbca+x7s/MWGgAKuFAPxW9GPh5VYOb5LhUcMaMA0GCSqGSIb3DQEBAQUABIIC
# AKUDBUKuUQje3iA8GZtVvJxipvnk2Z4hpacjYaSM+flZmOiBTifZOmSh+077FvRx
# /JDu7d1FEDzdjFeS3McTQMur1DEkxy7avub54PQgV9aGpP68PQ4jn56vjgwVur9r
# rC9CPfOjdJZAf+Kjx8tKnZ7AfMQQfy3q8eEZjeuV0KsRcx97UUtJk8V+mYRchVy2
# zHlcKTZlpPhFwBcIAm/g7/S5f8bMZAGo5Xriz+99KDM90PX1K0qkB2DF1javSMMC
# tmDj0y/4kldnVqwkLAoKGVniybbdSBSC4KLwoWSOSs7KnGehGJ8vblow0y/8+42W
# X5+GDOncpVjqNxE2tRiz1vGgV5m5fw4VYUc0HdE9advmPpLZc1yNLe5xXVt0lhPX
# akUg+0m4m1s7p/NyD1IeUhBq7nEHMlGE68gJHN6Hx8lW3++tJDEODOQnzP4U1YXK
# O4bThusX4U43Or/8EIyOPpQrZFjrafhiPsxAakSqzNlBQVy1g2GbTQ8ZT9A+juJu
# 84xMMYJeSuo9X7wm/EPBadEDeT9sEGxstRbMw9AnKhApKB/rrecUUfGn2xJgQXxc
# tFL9XrlUjDTvcjhBrKSjgvpxj24zuGuHqi2/KivZAljLyqP+WVekw7t2/AEVb3XJ
# 1wDfMHosU32b2e59XIIU3fVWqiazVP3X11VLHmAtwAVv
# SIG # End signature block
