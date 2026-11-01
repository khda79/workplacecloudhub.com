<#
.SYNOPSIS
    Audits and optionally remediates proxy addresses for local and remote Exchange user/shared mailboxes in specified Organizational Units (OUs), or all OUs.

.DESCRIPTION
    This script checks whether each local or remote UserMailbox and SharedMailbox has its expected proxy address.
    Local mailbox addresses are derived from the Exchange Alias. Remote mailboxes use their existing RemoteRoutingAddress first and fall back to Alias plus ExpectedSuffix only when no usable routing address exists.
    Before remediation, the expected address is checked against every mail-enabled Exchange recipient in the forest.
    If the expected address is missing and the -AddMissingAddress switch is specified, the script can add the address.
    The script can target multiple OUs via -OrganizationalUnit (string array) or all OUs with -AllOrganizationalUnit.
    It generates detailed, summary, and remediation CSV reports, groups them into a single Excel workbook, and maintains logs.
    Designed to run on an Exchange 2016 server with the Management Tools installed.

.PARAMETER OrganizationalUnit
    Distinguished names of the target OU(s). Accepts multiple OUs.
    No default scope is assumed. Configure at least one OU or use -AllOrganizationalUnit.

.PARAMETER AllOrganizationalUnit
    If specified, the script ignores OU filtering and fetches all User/Shared mailboxes in the forest.

.PARAMETER ExpectedSuffix
    The expected domain suffix used with Alias for local mailboxes and as the remote mailbox fallback (default: "tenant.mail.onmicrosoft.com").

.PARAMETER AddMissingAddress
    If specified, missing proxy addresses will be added. Recipients managed by an email address policy are reported but never modified.

.PARAMETER OutputPath
    Path to the output directory for reports and logs.

.PARAMETER SmtpServer
    SMTP relay to send the HTML summary email.

.PARAMETER SmtpPort
    Port for SMTP relay (default: 25).

.PARAMETER SendMailMode
    Mail transport mode: Graph, SMTP, or Both. Both uses Graph first and falls back to SMTP.

.PARAMETER From
    From address for notification email.

.PARAMETER To
    To recipients for notification email. Multiple recipients supported with ';' or ','.

.PARAMETER Cc
    CC recipients for notification email. Multiple recipients supported with ';' or ','.

.PARAMETER Subject
    Subject for notification email.

.PARAMETER MaxMailAttachmentMB
    Maximum workbook attachment size in MB. Larger workbooks are referenced in the email body but not attached.

.REQUIREMENTS
    Windows PowerShell 5.1 on an Exchange 2016/on-premises management host.
    Modules/snap-ins: SmartM365 WindowsPowerShell5 compatibility module; Exchange Management snap-in; ActiveDirectory module; ImportExcel module.
    Minimum permissions: Exchange on-premises recipient read access and AD read access for proxyAddresses, primary SMTP, UPN and related recipient attributes.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    - Requires Exchange 2016 Management Tools.
    - Generates detailed, summary, and added addresses CSV reports plus an Excel workbook with Check, ExistingProxyAddresses, DuplicateAliases, PlannedChanges, Summary, and Added worksheets.
    - Maintains logs and cleans up old files automatically.

.VERSION
1.38
.AUTHOR
    https://github.com/khda79/workplacecloudhub.com
    Minimum permissions: Windows PowerShell 5.1, Exchange 2016 Management snap-in, ActiveDirectory module, Exchange recipient read access, and AD read access.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$Tenant = 'test',
[Parameter(Mandatory=$false)]
    [string[]]$OrganizationalUnit = @(),

    [Parameter(Mandatory=$false)]
    [switch]$AllOrganizationalUnit,

    [Parameter(Mandatory=$false)]
    [string]$ExpectedSuffix = "tenant.mail.onmicrosoft.com",

    [Parameter(Mandatory=$false)]
    [switch]$AddMissingAddress,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [string]$SmtpServer = "",

    [Parameter(Mandatory=$false)]
    [int]$SmtpPort = 25,

    [Parameter(Mandatory=$false)]
    [ValidateSet("Graph","SMTP","Both")]
    [string]$SendMailMode = "",

    [Parameter(Mandatory=$false)]
    [string]$From = "",

    [Parameter(Mandatory=$false)]
    [string]$To = "",

    [Parameter(Mandatory=$false)]
    [string]$Cc = "",

    [Parameter(Mandatory=$false)]
    [string]$Subject = "Check ProxyAddresses Report",

    [Parameter(Mandatory=$false)]
    [ValidateRange(0, 100)]
    [int]$MaxMailAttachmentMB = 5,

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
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null


function Get-ScriptLocalConfig {
    param()
    Read-SmartM365ScriptConfiguration -ScriptPath $PSCommandPath
}

function Resolve-SmartM365ConfigValue {
    param([AllowNull()]$Value)
    Resolve-SmartM365ConfigurationTokens -Value $Value -Config $script:SmartM365GlobalConfig
}
function Get-ScriptLocalConfigValue {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)
    Get-SmartM365ConfigurationValue -Name $Name -LocalConfig $Config -EffectiveConfig $script:SmartM365GlobalConfig -DefaultValue $DefaultValue -EffectiveFallback Blank
}

function Ensure-ProxyAddressesImportExcelAllUsers {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $requiredRoot = Join-Path $env:ProgramFiles 'WindowsPowerShell/Modules/ImportExcel'
    $ready = Initialize-SmartM365ImportExcel -Scope AllUsers -RequiredModuleRoot $requiredRoot -InstallIfMissing
    return $ready.Path
}

function New-ProxyAddressesWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$CsvFiles,
        [Parameter(Mandatory = $true)][string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($script:ImportExcelManifestPath)) {
        throw 'The validated AllUsers ImportExcel manifest path is unavailable.'
    }
    Initialize-SmartM365ImportExcel -Scope AllUsers -ManifestPath $script:ImportExcelManifestPath | Out-Null

    $parentPath = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parentPath)) {
        New-Item -Path $parentPath -ItemType Directory -Force | Out-Null
    }
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }

    foreach ($csv in $CsvFiles) {
        $rows = @()
        if ($csv.PSObject.Properties.Name -contains 'Rows') {
            $rows = @($csv.Rows)
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$csv.Path) -and (Test-Path -LiteralPath $csv.Path)) {
            $rows = @(Import-Csv -LiteralPath $csv.Path)
        }

        $isEmpty = $rows.Count -eq 0
        if ($isEmpty) {
            $placeholder = [ordered]@{}
            foreach ($column in @($csv.EmptyColumns)) {
                $placeholder[[string]$column] = ''
            }
            $rows = @([pscustomobject]$placeholder)
        }

        $exportParameters = @{
            Path          = $Path
            WorksheetName = [string]$csv.WorksheetName
            AutoSize              = $true
            FreezeTopRow          = $true
            BoldTopRow            = $true
            AutoFilter            = $true
            NoHyperLinkConversion = '*'
        }

        if (-not $isEmpty) {
            $exportParameters.TableName = [string]$csv.TableName
        }

        $rows | Export-Excel @exportParameters

        if ($isEmpty) {
            $package = Open-ExcelPackage -Path $Path
            try {
                $package.Workbook.Worksheets[[string]$csv.WorksheetName].DeleteRow(2)
            }
            finally {
                Close-ExcelPackage -ExcelPackage $package
            }
        }
    }

    return (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName
}

function Get-RecipientObjectKey {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)]$Recipient)

    $guid = ([string]$Recipient.Guid).Trim()
    if ($guid -and $guid -ne [guid]::Empty.ToString()) {
        return ('guid:{0}' -f $guid.ToLowerInvariant())
    }

    $distinguishedName = ([string]$Recipient.DistinguishedName).Trim()
    if ($distinguishedName) {
        return ('dn:{0}' -f $distinguishedName.ToLowerInvariant())
    }

    return ('identity:{0}' -f ([string]$Recipient.Identity).Trim().ToLowerInvariant())
}

function ConvertTo-SmtpProxyAddress {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    $address = ([string]$Value).Trim() -replace '(?i)^smtp:', ''
    if ($address -notmatch '^[^@\s]+@[^@\s]+$') { return '' }
    return ('smtp:{0}' -f $address.ToLowerInvariant())
}

function ConvertTo-ProxyAddressParts {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([AllowNull()]$Value)

    $raw = ([string]$Value).Trim()
    $prefix = ''
    $addressValue = $raw
    $separatorIndex = $raw.IndexOf(':')
    if ($separatorIndex -gt 0) {
        $prefix = $raw.Substring(0, $separatorIndex)
        $addressValue = $raw.Substring($separatorIndex + 1)
    }

    $normalizedSmtpAddress = ''
    if ([string]::IsNullOrWhiteSpace($prefix) -or $prefix -ieq 'smtp') {
        $normalizedSmtpAddress = ConvertTo-SmtpProxyAddress -Value $raw
    }

    [pscustomobject]@{
        ProxyAddressType     = $prefix
        ProxyAddressValue    = $addressValue
        NormalizedSmtpAddress = $normalizedSmtpAddress
        IsPrimarySmtp        = ($prefix -ceq 'SMTP')
    }
}

function New-UniqueSuggestedSmtpAddress {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowEmptyString()][string]$Alias,
        [AllowEmptyString()][string]$ExpectedSuffix,
        [Parameter(Mandatory = $true)][System.Collections.Generic.HashSet[string]]$ReservedAddresses
    )

    if ([string]::IsNullOrWhiteSpace($Alias) -or [string]::IsNullOrWhiteSpace($ExpectedSuffix)) { return '' }

    $suffix = $ExpectedSuffix.Trim().TrimStart('@').TrimEnd('.').ToLowerInvariant()
    $localPart = $Alias.Trim().ToLowerInvariant() -replace '[^a-z0-9._+-]', '-'
    $localPart = $localPart.Trim('.-'.ToCharArray())
    if ([string]::IsNullOrWhiteSpace($localPart)) { return '' }

    for ($candidateIndex = 2; $candidateIndex -le 999; $candidateIndex++) {
        $candidate = ConvertTo-SmtpProxyAddress -Value ('{0}-{1}@{2}' -f $localPart, $candidateIndex, $suffix)
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if (-not $ReservedAddresses.Contains($candidate)) {
            [void]$ReservedAddresses.Add($candidate)
            return $candidate
        }
    }

    return ''
}

function Assert-ProxyAddressesGraphPrerequisite {
    [CmdletBinding()]
    param(
        [bool]$EnableSharePointUpload,
        [AllowEmptyString()][string]$SendMailMode,
        [AllowEmptyString()][string]$SmtpServer
    )

    $effectiveMailMode = ([string]$SendMailMode).Trim()
    if (-not $effectiveMailMode) {
        $effectiveMailMode = if ([string]::IsNullOrWhiteSpace($SmtpServer)) { 'Graph' } else { 'SMTP' }
    }
    if (-not $EnableSharePointUpload -and $effectiveMailMode -notin @('Graph','Both')) {
        WriteLog -Message 'Graph Authentication prerequisite skipped: SharePoint upload disabled and mail uses SMTP.'
        return
    }

    $moduleName = 'Microsoft.Graph.Authentication'
    $runtime = '{0} {1}' -f $PSVersionTable.PSEdition, $PSVersionTable.PSVersion
    WriteLog -Message ("Checking {0} prerequisite; PowerShell={1}; SharePointUpload={2}; MailMode={3}." -f $moduleName, $runtime, $EnableSharePointUpload, $effectiveMailMode)
    try {
        $loaded = Get-Module -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        if ($null -eq $loaded) {
            $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction Stop | Sort-Object Version -Descending | Select-Object -First 1
            if ($null -eq $available) {
                # Pin the missing-module bootstrap to the PS5.1-compatible release; never upgrade an available module.
                $installVersion = '2.41.2'
                WriteLog -Message ("{0} is absent; installing version {1} from PSGallery with Scope=CurrentUser for {2}." -f $moduleName, $installVersion, [Environment]::UserName)
                Install-Module -Name $moduleName -RequiredVersion $installVersion -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction Stop | Sort-Object Version -Descending | Select-Object -First 1
                if ($null -eq $available) { throw 'Installation completed but Graph Authentication is still not discoverable in this process PSModulePath.' }
            }
            Import-Module -Name $available.Path -Global -ErrorAction Stop
            $loaded = Get-Module -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        }
        if ($null -eq $loaded) { throw 'Import completed without a loaded Graph Authentication module.' }
        WriteLog -Message ("Graph Authentication prerequisite OK; PowerShell={0}; Version={1}; Path={2}." -f $runtime, $loaded.Version, $loaded.Path)
    }
    catch {
        throw ("Graph Authentication prerequisite failed under PowerShell {0}: {1} An available module is reused; installation is attempted only when absent, with Scope=CurrentUser. No Graph connection was attempted. PSModulePath={2}" -f $runtime, $_.Exception.Message, $env:PSModulePath)
    }
}

function Test-SmtpAddressSuffix {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowEmptyString()][string]$ProxyAddress,
        [AllowEmptyString()][string]$Suffix
    )

    if ([string]::IsNullOrWhiteSpace($ProxyAddress) -or [string]::IsNullOrWhiteSpace($Suffix)) {
        return $false
    }

    $address = $ProxyAddress -replace '(?i)^smtp:', ''
    $atIndex = $address.LastIndexOf('@')
    if ($atIndex -lt 1 -or $atIndex -eq ($address.Length - 1)) { return $false }
    return $address.Substring($atIndex + 1).TrimEnd('.') -ieq $Suffix.Trim().TrimStart('@').TrimEnd('.')
}

$ScriptLocalConfig = Get-ScriptLocalConfig

$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$ErrorActionPreference = 'Stop'
    if (-not $PSBoundParameters.ContainsKey('OrganizationalUnit')) { $OrganizationalUnit = @((Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrganizationalUnit' -DefaultValue $OrganizationalUnit)) }
    foreach ($configName in @('ExpectedSuffix','SendMailMode','SmtpServer','From','To')) {
        if (-not $PSBoundParameters.ContainsKey($configName)) {
            Set-Variable -Name $configName -Value (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name $configName -DefaultValue (Get-Variable -Name $configName -ValueOnly)) -Scope Local
        }
    }
    if (-not $PSBoundParameters.ContainsKey('Subject')) {
        $subjectProperty = $ScriptLocalConfig.PSObject.Properties['Subject']
        if ($null -ne $subjectProperty -and $null -ne $subjectProperty.Value) {
            if ($subjectProperty.Value -is [string]) {
                $localSubject = $subjectProperty.Value.Trim()
                if ($localSubject -and $localSubject -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                    $Subject = Resolve-SmartM365ConfigValue -Value $subjectProperty.Value
                }
            }
            else {
                $Subject = Resolve-SmartM365ConfigValue -Value $subjectProperty.Value
            }
        }
    }

    #region Module Import and Initialization
$ScriptVersion = "1.38"
    $TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
    $OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ProxyAddressesCsvLogFolderPath' -DefaultValue $OutputPath
    $LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
    function Join-ModulePath {
        param([Parameter(Mandatory)][string]$FileName)
        $repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..')).Path
        return (Join-Path (Join-Path (Join-Path (Join-Path $repoRoot 'Modules') 'SmartM365.Core') 'Compatibility\WindowsPowerShell5') $FileName)
    }

    try {
        Write-Host "Loading module SmartM365-WindowsPowerShell5.psd1..."
        Import-Module -Name (Join-ModulePath 'SmartM365-WindowsPowerShell5.psd1') -MinimumVersion '1.0.70' -ErrorAction Stop
        $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
        . (Join-ModulePath '..\..\SmartM365-SharePointUploadPolicy.ps1')
        Initialize-SmartM365CollectorUploadContext -LocalConfig $ScriptLocalConfig -GlobalConfig $global:SmartM365GlobalConfig -DataAllRoot $InitializeOutputPath -DataLastRoot $LatestCsvFolderPath -LogAllRoot $global:LogPath -ExternalActionsAllowed $true
        Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $LatestCsvFolderPath -ScopeParameters @{OrganizationalUnit=$OrganizationalUnit;AllOrganizationalUnit=$AllOrganizationalUnit;AddMissingAddress=[bool]$AddMissingAddress}
        Start-SmartM365RunTranscript -Path $global:logTranscriptFile -Append
        WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
        $OutputPath = $InitializeOutputPath
        WriteLog -Message "Starting $TaskName..."
    } catch {
        Write-Host "Initialization failed: $($_.Exception.Message)`n$($_.ScriptStackTrace)" -ForegroundColor Red
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(0); exit $smartM365RunGuard.ExitCode
    }
    #endregion

    $OrganizationalUnit = @(
        $OrganizationalUnit |
            ForEach-Object { if ($null -ne $_) { ([string]$_).Trim() } } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
    if (-not $AllOrganizationalUnit -and $OrganizationalUnit.Count -eq 0) {
        $scopeError = 'No OrganizationalUnit is configured. Set OrganizationalUnit in the local JSON or specify -AllOrganizationalUnit.'
        WriteLog -Message $scopeError -Level 'ERROR'
        throw $scopeError
    }

    $timestamp      = Get-Date -Format 'yyyyMMdd-HHmm'
    $outDetail      = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Check_{0}.csv" -f $timestamp)
    $outSummary     = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Summary_{0}.csv" -f $timestamp)
    $outAdded       = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Added_{0}.csv" -f $timestamp)
    $outWorkbook    = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_{0}.xlsx" -f $timestamp)
    $latestDetail   = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Check.csv' } else { $null }
    $latestSummary  = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Summary.csv' } else { $null }
    $latestAdded    = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Added.csv' } else { $null }
    $latestWorkbook = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses.xlsx' } else { $null }
    if (Test-SmartM365MaxItemsMode) {
        $outDetail = Add-SmartM365MaxItemsSuffixToCsvPath -Path $outDetail
        $outSummary = Add-SmartM365MaxItemsSuffixToCsvPath -Path $outSummary
        $outAdded = Add-SmartM365MaxItemsSuffixToCsvPath -Path $outAdded
        $outWorkbook = Add-SmartM365MaxItemsSuffixToCsvPath -Path $outWorkbook
        if ($latestDetail) { $latestDetail = Add-SmartM365MaxItemsSuffixToCsvPath -Path $latestDetail }
        if ($latestSummary) { $latestSummary = Add-SmartM365MaxItemsSuffixToCsvPath -Path $latestSummary }
        if ($latestAdded) { $latestAdded = Add-SmartM365MaxItemsSuffixToCsvPath -Path $latestAdded }
        if ($latestWorkbook) { $latestWorkbook = Add-SmartM365MaxItemsSuffixToCsvPath -Path $latestWorkbook }
    }

    Write-Host "=== Start $(Get-Date) ==="
    if ($AllOrganizationalUnit) {
        Write-Host "Scope       : ALL (no OU filtering, entire forest)"
        WriteLog -Message "Target OU(s): ALL (entire forest)"
    } else {
        $OuListString = ($OrganizationalUnit -join "; ")
        Write-Host "Scope       : $OuListString"
        WriteLog -Message "Target OU(s): $OuListString"
    }
    Write-Host "Expected suffix: $ExpectedSuffix"
    Write-Host "AddMissing   : $AddMissingAddress"
    Write-Host "Detail CSV  : $outDetail"
    Write-Host "Summary CSV : $outSummary"
    Write-Host "Added CSV   : $outAdded"
    Write-Host "Excel       : $outWorkbook"

    Assert-ProxyAddressesGraphPrerequisite -EnableSharePointUpload $global:EnableSharePointUpload -SendMailMode $SendMailMode -SmtpServer $SmtpServer

    # -------------------------------
    # Exchange PSSnapin availability
    # -------------------------------
    $snapinName = "Microsoft.Exchange.Management.PowerShell.SnapIn"

    if (-not (Get-PSSnapin $snapinName -Registered -ErrorAction SilentlyContinue)) {
        Write-Error "The Exchange Management PSSnapin '$snapinName' is not registered on this server.`nThis script must be run on an Exchange 2016 server where the Management Tools are installed."
        Stop-SmartM365RunTranscript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
    }

    # Verify the PowerShell security dependency and Excel before Exchange loads type data.
    $script:ImportExcelManifestPath = Ensure-ProxyAddressesImportExcelAllUsers

    if (-not (Get-PSSnapin $snapinName -ErrorAction SilentlyContinue)) {
        Write-Host "The Exchange PSSnapin '$snapinName' is not loaded in the current session. Attempting to load it..."
        try {
            Add-PSSnapin $snapinName -ErrorAction Stop
            Write-Host "The Exchange PSSnapin was loaded successfully."
        }
        catch {
            Write-Error "Failed to load the Exchange PSSnapin '$snapinName'. Error: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            Write-Error "Ensure you are running this script on an Exchange 2016 server and have the necessary permissions.`nAlternatively, try running this script from the Exchange Management Shell or use a script that connects via PowerShell Remoting to localhost."
            Stop-SmartM365RunTranscript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
        }
    } else {
        Write-Host "The Exchange PSSnapin '$snapinName' is already loaded."
    }

    $requiredExchangeCommands = @('Get-Recipient', 'Get-RemoteMailbox')
    if ($AddMissingAddress) { $requiredExchangeCommands += @('Set-Mailbox', 'Set-RemoteMailbox') }
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredModules @('ImportExcel') -RequiredCommands $requiredExchangeCommands -RequireExchangeOnPrem -RequireActiveDirectoryRead -InstallModuleScope AllUsers | Out-Null
    if ([string]::IsNullOrWhiteSpace($script:ImportExcelManifestPath)) {
        $script:ImportExcelManifestPath = Ensure-ProxyAddressesImportExcelAllUsers
    }
    if ([string]::IsNullOrWhiteSpace($script:ImportExcelManifestPath)) {
        throw 'ImportExcel is not installed in the AllUsers Windows PowerShell module path after preflight.'
    }
    Get-Module -Name ImportExcel | Remove-Module -Force -ErrorAction SilentlyContinue
    Initialize-SmartM365ImportExcel -Scope AllUsers -ManifestPath $script:ImportExcelManifestPath | Out-Null
    $importExcelModule = Get-Module -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
    WriteLog -Message ("ImportExcel module loaded: version={0}; path={1}" -f $importExcelModule.Version, $importExcelModule.Path)

    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-Error "The Get-Mailbox cmdlet is still not available after attempting to load the snap-in.`nThis could indicate an issue with the Exchange Management Tools installation."
        Stop-SmartM365RunTranscript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
    }

    if (-not (Get-Command Get-Recipient -ErrorAction SilentlyContinue)) {
        Write-Error "Get-Recipient is unavailable. Ensure the Exchange Management Tools are installed, or run from the Exchange Management Shell."
        Stop-SmartM365RunTranscript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
    }

    # --- Enable ViewEntireForest
    $script:PrevViewEntireForest = $null
    $script:ChangedViewEntireForest = $false
    try {
        $currentSettings = Get-ADServerSettings
        $script:PrevViewEntireForest = $currentSettings.ViewEntireForest
        if (-not $currentSettings.ViewEntireForest) {
            Set-ADServerSettings -ViewEntireForest $true
            $script:ChangedViewEntireForest = $true
            Write-Host "ViewEntireForest enabled for this session."
        } else {
            Write-Host "ViewEntireForest already enabled."
        }
    } catch {
        Write-Warning "Unable to read/set ViewEntireForest : $($_.Exception.Message)`n$($_.ScriptStackTrace)"
    }

    # Collections & counters
    $results                = New-Object System.Collections.Generic.List[object]
    $addedOperations        = New-Object System.Collections.Generic.List[object]
    $plannedChangeRows     = New-Object System.Collections.Generic.List[object]
    $existingProxyAddressRows = New-Object System.Collections.Generic.List[object]
    $duplicateAliasRows = New-Object System.Collections.Generic.List[object]

    $okCount                = 0
    $missingCount           = 0
    $noPrimaryCount         = 0
    $noAliasCount           = 0
    $noExpectedAddressCount = 0
    $addedCount             = 0
    $addFailedCount         = 0
    $policyEnabledCount     = 0
    $policySkippedCount     = 0
    $localMailboxCount      = 0
    $remoteMailboxCount     = 0
    $remoteRoutingAddressCount = 0
    $remoteAliasFallbackCount = 0
    $remoteAliasFallbackMissingBlockedCount = 0
    $remoteRoutingSuffixMismatchCount = 0
    $remoteMailboxLookupMissCount = 0
    $duplicateExpectedMissingCount = 0
    $addressAlreadyAssignedMissingCount = 0

    # Pre/post remediation counters
    $preMissing             = 0
    $postMissing            = 0

    $recipientResultSize = if ($MaxItems -gt 0) { $MaxItems } else { 'Unlimited' }
    $localRecipientTypes = @('UserMailbox', 'SharedMailbox')
    $remoteRecipientTypes = @('RemoteUserMailbox', 'RemoteSharedMailbox')
    $recipientTypes = @($localRecipientTypes + $remoteRecipientTypes)
    $allMailEnabledRecipients = @()

    # Retrieve the forest-wide recipient set once when possible. It is also used to
    # detect an expected SMTP address already assigned to any mail-enabled object.
    if ($AllOrganizationalUnit) {
        try {
            Write-Host "Fetching all mail-enabled recipients from the entire forest..."
            $allMailEnabledRecipients = @(Get-Recipient -ResultSize Unlimited -ErrorAction Stop)
            $recipients = @($allMailEnabledRecipients | Where-Object { ([string]$_.RecipientTypeDetails) -in $recipientTypes })
        }
        catch {
            Write-Error "Failed to Get-Recipient for ALL OU: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            Stop-SmartM365RunTranscript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            throw
        }
    }
    else {
        $scopedRecipients = @()
        foreach ($ou in $OrganizationalUnit) {
            try {
                Write-Host "Fetching recipients from OU: $ou"
                $recs = Get-Recipient -OrganizationalUnit $ou `
                                      -ResultSize $recipientResultSize `
                                      -RecipientTypeDetails $recipientTypes `
                                      -ErrorAction Stop
                if ($recs) { $scopedRecipients += $recs }
            }
            catch {
                Write-Warning "Get-Recipient failed for OU '$ou' : $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            }
        }
        $recipients = @($scopedRecipients | Sort-Object -Property Guid -Unique)

        try {
            Write-Host "Fetching all mail-enabled recipients for forest-wide SMTP uniqueness checks..."
            $allMailEnabledRecipients = @(Get-Recipient -ResultSize Unlimited -ErrorAction Stop)
        }
        catch {
            Write-Error "Failed to retrieve forest-wide recipients for SMTP uniqueness checks: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            throw
        }
    }

    if ($MaxItems -gt 0) {
        $preMaxItemsRecipientCount = @($recipients).Count
        $recipients = @($recipients | Sort-Object PrimarySmtpAddress,Name | Select-Object -First $MaxItems)
        WriteLog -Message ("MaxItems enabled: restricted recipients from {0} to {1}; forest-wide conflict checks remain unlimited." -f $preMaxItemsRecipientCount, @($recipients).Count) -Level 'WARNING'
    }

    $total   = @($recipients).Count
    $counter = 0
    Write-Host "User/Shared mailboxes retrieved: $total"
    Write-Host "Mail-enabled recipients indexed for SMTP conflicts: $(@($allMailEnabledRecipients).Count)"

    # Build one forest-wide SMTP owner index. The nested owner map deduplicates
    # differently-cased copies of the same proxy on the same recipient.
    $smtpAddressOwners = @{}
    foreach ($recipientForAddressIndex in $allMailEnabledRecipients) {
        $ownerKey = Get-RecipientObjectKey -Recipient $recipientForAddressIndex
        $ownerDescription = '{0}:{1}' -f ([string]$recipientForAddressIndex.RecipientTypeDetails), ([string]$recipientForAddressIndex.Identity)
        foreach ($proxyAddress in @($recipientForAddressIndex.EmailAddresses)) {
            $normalizedProxyAddress = ConvertTo-SmtpProxyAddress -Value $proxyAddress
            if (-not $normalizedProxyAddress) { continue }
            if (-not $smtpAddressOwners.ContainsKey($normalizedProxyAddress)) {
                $smtpAddressOwners[$normalizedProxyAddress] = @{}
            }
            $smtpAddressOwners[$normalizedProxyAddress][$ownerKey] = $ownerDescription
        }
    }

    # Retrieve all remote mailbox routing attributes in one Exchange call, then
    # resolve target objects locally by GUID, DN, or Alias.
    $remoteMailboxByGuid = @{}
    $remoteMailboxByDistinguishedName = @{}
    $remoteMailboxByAlias = @{}
    if (@($recipients | Where-Object { ([string]$_.RecipientTypeDetails) -in $remoteRecipientTypes }).Count -gt 0) {
        Write-Host "Fetching RemoteMailbox routing attributes in one batch..."
        $remoteMailboxes = @(Get-RemoteMailbox -ResultSize Unlimited -ErrorAction Stop)
        foreach ($remoteMailbox in $remoteMailboxes) {
            $remoteGuid = ([string]$remoteMailbox.Guid).Trim().ToLowerInvariant()
            if ($remoteGuid -and $remoteGuid -ne [guid]::Empty.ToString()) { $remoteMailboxByGuid[$remoteGuid] = $remoteMailbox }
            $remoteDn = ([string]$remoteMailbox.DistinguishedName).Trim().ToLowerInvariant()
            if ($remoteDn) { $remoteMailboxByDistinguishedName[$remoteDn] = $remoteMailbox }
            $remoteAlias = ([string]$remoteMailbox.Alias).Trim().ToLowerInvariant()
            if ($remoteAlias) { $remoteMailboxByAlias[$remoteAlias] = $remoteMailbox }
        }
        Write-Host "RemoteMailbox routing objects retrieved: $($remoteMailboxes.Count)"
    }

    # Resolve the expected address once per target. RemoteRoutingAddress is the
    # authoritative value for remote mailboxes; Alias@ExpectedSuffix is fallback.
    $recipientPlans = New-Object System.Collections.Generic.List[object]
    foreach ($rec in $recipients) {
        $isRemoteMailbox = ([string]$rec.RecipientTypeDetails) -in $remoteRecipientTypes
        $alias = ([string]$rec.Alias).Trim()
        $samAccountName = ([string]$rec.SamAccountName).Trim()
        $recipientKey = Get-RecipientObjectKey -Recipient $rec
        $remoteMailbox = $null
        $remoteRoutingAddressRaw = ''
        $remoteRoutingAddress = ''
        $expectedAddress = ''
        $expectedAddressSource = ''
        $routingWarning = ''
        $remoteRoutingAddressSuffixMatchesExpected = $null

        if ([string]::IsNullOrWhiteSpace($alias)) { $noAliasCount++ }

        if ($isRemoteMailbox) {
            $remoteGuid = ([string]$rec.Guid).Trim().ToLowerInvariant()
            $remoteDn = ([string]$rec.DistinguishedName).Trim().ToLowerInvariant()
            $remoteAliasKey = $alias.ToLowerInvariant()
            if ($remoteGuid -and $remoteMailboxByGuid.ContainsKey($remoteGuid)) {
                $remoteMailbox = $remoteMailboxByGuid[$remoteGuid]
            }
            elseif ($remoteDn -and $remoteMailboxByDistinguishedName.ContainsKey($remoteDn)) {
                $remoteMailbox = $remoteMailboxByDistinguishedName[$remoteDn]
            }
            elseif ($remoteAliasKey -and $remoteMailboxByAlias.ContainsKey($remoteAliasKey)) {
                $remoteMailbox = $remoteMailboxByAlias[$remoteAliasKey]
            }

            if ($null -ne $remoteMailbox) {
                $remoteRoutingAddressRaw = ([string]$remoteMailbox.RemoteRoutingAddress).Trim()
                $remoteRoutingAddress = ConvertTo-SmtpProxyAddress -Value $remoteMailbox.RemoteRoutingAddress
            }
            else {
                $remoteMailboxLookupMissCount++
                $routingWarning = 'RemoteMailbox lookup failed; Alias fallback used when available.'
            }

            if ($remoteRoutingAddress) {
                $expectedAddress = $remoteRoutingAddress
                $expectedAddressSource = 'RemoteRoutingAddress'
                $remoteRoutingAddressCount++
                $remoteRoutingAddressSuffixMatchesExpected = Test-SmtpAddressSuffix -ProxyAddress $remoteRoutingAddress -Suffix $ExpectedSuffix
                if (-not $remoteRoutingAddressSuffixMatchesExpected) {
                    $remoteRoutingSuffixMismatchCount++
                    $routingWarning = 'RemoteRoutingAddress suffix differs from ExpectedSuffix.'
                }
            }
            elseif ($alias) {
                $expectedAddress = ConvertTo-SmtpProxyAddress -Value ("{0}@{1}" -f $alias, $ExpectedSuffix)
                $expectedAddressSource = 'AliasFallback'
                $remoteAliasFallbackCount++
                if ($remoteRoutingAddressRaw) {
                    $routingWarning = 'RemoteRoutingAddress is invalid; Alias fallback used.'
                }
                elseif (-not $routingWarning) {
                    $routingWarning = 'RemoteRoutingAddress is empty; Alias fallback used.'
                }
            }
        }
        elseif ($alias) {
            $expectedAddress = ConvertTo-SmtpProxyAddress -Value ("{0}@{1}" -f $alias, $ExpectedSuffix)
            $expectedAddressSource = 'Alias'
        }

        if (-not $expectedAddress) {
            $expectedAddressSource = 'Unavailable'
            $noExpectedAddressCount++
        }

        $recipientPlans.Add([pscustomobject]@{
            Recipient                                = $rec
            RecipientKey                             = $recipientKey
            Alias                                    = $alias
            SamAccountName                           = $samAccountName
            IsRemoteMailbox                          = $isRemoteMailbox
            MailboxLocation                          = if ($isRemoteMailbox) { 'Remote' } else { 'OnPremises' }
            RemoteRoutingAddress                     = $remoteRoutingAddress
            RemoteRoutingAddressRaw                  = $remoteRoutingAddressRaw
            RemoteRoutingAddressSuffixMatchesExpected = $remoteRoutingAddressSuffixMatchesExpected
            ExpectedAddress                          = $expectedAddress
            ExpectedAddressSource                    = $expectedAddressSource
            RoutingWarning                           = $routingWarning
        })
    }

    $expectedAddressCounts = @{}
    $expectedAddressPlanOwners = @{}
    foreach ($plan in $recipientPlans) {
        if (-not $plan.ExpectedAddress) { continue }
        if (-not $expectedAddressPlanOwners.ContainsKey($plan.ExpectedAddress)) {
            $expectedAddressPlanOwners[$plan.ExpectedAddress] = @{}
        }
        $expectedAddressPlanOwners[$plan.ExpectedAddress][$plan.RecipientKey] = '{0}:{1}' -f ([string]$plan.Recipient.RecipientTypeDetails), ([string]$plan.Recipient.Identity)
        if ($expectedAddressCounts.ContainsKey($plan.ExpectedAddress)) {
            $expectedAddressCounts[$plan.ExpectedAddress]++
        }
        else {
            $expectedAddressCounts[$plan.ExpectedAddress] = 1
        }
    }

    $duplicateExpectedAddresses = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $duplicateExpectedRecipientCount = 0
    foreach ($expectedAddressCountEntry in $expectedAddressCounts.GetEnumerator()) {
        if ([int]$expectedAddressCountEntry.Value -le 1) { continue }
        [void]$duplicateExpectedAddresses.Add([string]$expectedAddressCountEntry.Key)
        $duplicateExpectedRecipientCount += [int]$expectedAddressCountEntry.Value
    }
    $duplicateExpectedAddressGroupCount = $duplicateExpectedAddresses.Count
    if ($duplicateExpectedAddressGroupCount -gt 0) {
        WriteLog -Message ("Duplicate expected proxy addresses detected: groups={0}; recipients={1}. Conflicting addresses are blocked from remediation." -f $duplicateExpectedAddressGroupCount, $duplicateExpectedRecipientCount) -Level 'WARNING'
    }

    $addressAlreadyAssignedConflictAddresses = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $suggestedAddressReservations = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($existingSmtpAddress in @($smtpAddressOwners.Keys)) { [void]$suggestedAddressReservations.Add([string]$existingSmtpAddress) }
    foreach ($plannedExpectedAddress in @($expectedAddressCounts.Keys)) { [void]$suggestedAddressReservations.Add([string]$plannedExpectedAddress) }
    foreach ($plan in $recipientPlans) {
        $counter++
        if ($counter -eq 1 -or ($counter % 250) -eq 0 -or $counter -eq $total) {
            Write-Progress -Activity "Checking EmailAddresses..." -Status "$counter / $total" -PercentComplete (($counter / [Math]::Max($total, 1)) * 100)
        }

        $rec = $plan.Recipient
        $primary = $null
        try { $primary = $rec.PrimarySmtpAddress } catch { $primary = $null }

        $email = ''
        if ($primary -and $primary.ToString() -match '.+@.+') { $email = $primary.ToString() }
        $status = ''
        $policyWarning = $false
        $expectedAddress = [string]$plan.ExpectedAddress
        $isDuplicateExpectedAddress = $expectedAddress -and $duplicateExpectedAddresses.Contains($expectedAddress)
        $conflictOwners = @()
        $duplicateExpectedPeers = @()
        if ($expectedAddress -and $smtpAddressOwners.ContainsKey($expectedAddress)) {
            foreach ($ownerEntry in $smtpAddressOwners[$expectedAddress].GetEnumerator()) {
                if ([string]$ownerEntry.Key -ne [string]$plan.RecipientKey) {
                    $conflictOwners += [string]$ownerEntry.Value
                }
            }
        }
        if ($isDuplicateExpectedAddress -and $expectedAddressPlanOwners.ContainsKey($expectedAddress)) {
            foreach ($peerEntry in $expectedAddressPlanOwners[$expectedAddress].GetEnumerator()) {
                if ([string]$peerEntry.Key -ne [string]$plan.RecipientKey) {
                    $duplicateExpectedPeers += [string]$peerEntry.Value
                }
            }
        }
        $isAddressAlreadyAssigned = $conflictOwners.Count -gt 0
        $isExpectedAddressConflict = $isDuplicateExpectedAddress -or $isAddressAlreadyAssigned
        $conflictReason = @()
        if ($isDuplicateExpectedAddress) { $conflictReason += 'DuplicateExpectedAddress' }
        if ($isAddressAlreadyAssigned) {
            $conflictReason += 'AddressAlreadyAssigned'
            [void]$addressAlreadyAssignedConflictAddresses.Add($expectedAddress)
        }

        $suggestedUniqueAddress = ''
        $suggestedUniqueAddressReason = ''
        if ($isExpectedAddressConflict) {
            if ($plan.ExpectedAddressSource -eq 'Alias' -and -not [string]::IsNullOrWhiteSpace($plan.Alias)) {
                $suggestedUniqueAddress = New-UniqueSuggestedSmtpAddress -Alias $plan.Alias -ExpectedSuffix $ExpectedSuffix -ReservedAddresses $suggestedAddressReservations
                if (-not [string]::IsNullOrWhiteSpace($suggestedUniqueAddress)) {
                    if ($isDuplicateExpectedAddress -and $isAddressAlreadyAssigned) {
                        $suggestedUniqueAddressReason = 'AliasDuplicateAndAddressAssignedCandidate'
                    }
                    elseif ($isDuplicateExpectedAddress) {
                        $suggestedUniqueAddressReason = 'AliasDuplicateCandidate'
                    }
                    elseif ($isAddressAlreadyAssigned) {
                        $suggestedUniqueAddressReason = 'AddressAlreadyAssignedCandidate'
                    }
                }
            }
            elseif ($plan.ExpectedAddressSource -eq 'RemoteRoutingAddress') {
                $suggestedUniqueAddressReason = 'RemoteRoutingAddressConflict-NoSuggestion'
            }
            elseif ($plan.ExpectedAddressSource -eq 'AliasFallback') {
                $suggestedUniqueAddressReason = 'RemoteRoutingAddressUnavailable-NoSuggestion'
            }
        }

        if ($plan.IsRemoteMailbox) { $remoteMailboxCount++ } else { $localMailboxCount++ }

        $proxyAddressIndex = 0
        foreach ($recipientProxyAddress in @($rec.EmailAddresses)) {
            $proxyAddressRaw = ([string]$recipientProxyAddress).Trim()
            if ([string]::IsNullOrWhiteSpace($proxyAddressRaw)) { continue }

            $proxyAddressIndex++
            $proxyAddressParts = ConvertTo-ProxyAddressParts -Value $proxyAddressRaw
            $existingProxyAddressRows.Add([PSCustomObject]@{
                Identity               = $rec.Identity
                Alias                  = $plan.Alias
                SamAccountName         = $plan.SamAccountName
                DisplayName            = $rec.DisplayName
                RecipientType          = $rec.RecipientTypeDetails
                MailboxLocation        = $plan.MailboxLocation
                PrimaryAddress         = $email
                RemoteRoutingAddress   = $plan.RemoteRoutingAddress
                AddressIndex           = $proxyAddressIndex
                ProxyAddress           = $proxyAddressRaw
                ProxyAddressType       = $proxyAddressParts.ProxyAddressType
                ProxyAddressValue      = $proxyAddressParts.ProxyAddressValue
                NormalizedSmtpAddress  = $proxyAddressParts.NormalizedSmtpAddress
                IsPrimarySmtp          = $proxyAddressParts.IsPrimarySmtp
                IsExpectedAddress      = (-not [string]::IsNullOrWhiteSpace($expectedAddress) -and $proxyAddressParts.NormalizedSmtpAddress -ieq $expectedAddress)
                EmailAddressPolicyEnabled = $rec.EmailAddressPolicyEnabled
            })
        }

        # Track Email Address Policy status without flooding the console.
        if ($rec.EmailAddressPolicyEnabled -eq $true) {
            $policyWarning = $true
            $policyEnabledCount++
        }

        if ($primary -and $primary.ToString() -match '.+@.+') {
            $email = $primary.ToString()

            if (-not $expectedAddress) {
                $status = 'No expected address'
            }
            else {
                $exists = $false
                foreach ($recipientProxyAddress in @($rec.EmailAddresses)) {
                    $normalizedRecipientProxyAddress = ConvertTo-SmtpProxyAddress -Value $recipientProxyAddress
                    if ($normalizedRecipientProxyAddress -ieq $expectedAddress) {
                        $exists = $true
                        break
                    }
                }

                if ($exists) {
                    $status = 'OK'
                    $okCount++
                }
                else {
                    $preMissing++
                    $missingCount++

                    $isSafeWriteCandidate = (-not $isExpectedAddressConflict) -and (-not ($plan.IsRemoteMailbox -and $plan.ExpectedAddressSource -eq 'AliasFallback')) -and ($rec.EmailAddressPolicyEnabled -ne $true)
                    if ($isSafeWriteCandidate) {
                        $plannedCmdlet = if ($plan.IsRemoteMailbox) { 'Set-RemoteMailbox' } else { 'Set-Mailbox' }
                        $plannedExecutionMode = if ($AddMissingAddress) { 'WriteRun' } else { 'ReadOnlyPreview' }
                        $plannedCurrentStatus = if ($AddMissingAddress) { 'Missing -> WillAdd' } else { 'Missing' }
                        $identityForCommand = ([string]$rec.Identity).Replace("'", "''")
                        $addressForCommand = ([string]$expectedAddress).Replace("'", "''")
                        $plannedChangeRows.Add([PSCustomObject]@{
                            Action                    = 'AddProxyAddress'
                            ExecutionMode             = $plannedExecutionMode
                            WouldRunInWriteMode       = $true
                            Identity                  = $rec.Identity
                            Alias                     = $plan.Alias
                            SamAccountName            = $plan.SamAccountName
                            DisplayName               = $rec.DisplayName
                            RecipientType             = $rec.RecipientTypeDetails
                            MailboxLocation           = $plan.MailboxLocation
                            PrimarySmtp               = $email
                            RemoteRoutingAddress      = $plan.RemoteRoutingAddress
                            ExpectedAddressSource     = $plan.ExpectedAddressSource
                            PlannedProxy              = $expectedAddress
                            CurrentStatus             = $plannedCurrentStatus
                            Reason                    = 'SafeMissingExpectedProxyAddress'
                            EmailAddressPolicyEnabled = $rec.EmailAddressPolicyEnabled
                            Cmdlet                    = $plannedCmdlet
                            CommandPreview            = ("{0} -Identity '{1}' -EmailAddresses @{{ add = '{2}' }}" -f $plannedCmdlet, $identityForCommand, $addressForCommand)
                        })
                    }

                    if ($isExpectedAddressConflict) {
                        if ($isDuplicateExpectedAddress) { $duplicateExpectedMissingCount++ }
                        if ($isAddressAlreadyAssigned) { $addressAlreadyAssignedMissingCount++ }
                        if ($isDuplicateExpectedAddress -and $isAddressAlreadyAssigned) {
                            $status = 'Missing -> AddressConflict'
                        }
                        elseif ($isDuplicateExpectedAddress) {
                            $status = 'Missing -> DuplicateExpectedAddress'
                        }
                        else {
                            $status = 'Missing -> AddressAlreadyAssigned'
                        }
                    }
                    elseif ($plan.IsRemoteMailbox -and $plan.ExpectedAddressSource -eq 'AliasFallback') {
                        $remoteAliasFallbackMissingBlockedCount++
                        $status = 'Missing -> RemoteRoutingAddressUnavailable'
                    }
                    elseif ($AddMissingAddress -and $rec.EmailAddressPolicyEnabled -eq $true) {
                        $status = 'Missing -> SkippedPolicyEnabled'
                        $policySkippedCount++
                    }
                    elseif ($AddMissingAddress) {
                        $status = 'Missing -> WillAdd'
                        try {
                            $params = @{
                                Identity       = $rec.Identity
                                EmailAddresses = @{ add = $expectedAddress }
                                ErrorAction    = 'Stop'
                            }

                            if ($PSCmdlet.ShouldProcess($rec.Identity, "Add proxy address $expectedAddress")) {
                                if ($plan.IsRemoteMailbox) {
                                    Set-RemoteMailbox @params
                                }
                                else {
                                    Set-Mailbox @params
                                }
                                Write-Host "apply $expectedAddress for $email"
                                $addedCount++
                                $status = 'Added'

                                $addedOperations.Add([PSCustomObject]@{
                                    Identity             = $rec.Identity
                                    Alias                = $plan.Alias
                                    SamAccountName       = $plan.SamAccountName
                                    DisplayName          = $rec.DisplayName
                                    RecipientType        = $rec.RecipientTypeDetails
                                    MailboxLocation      = $plan.MailboxLocation
                                    RemoteRoutingAddress = $plan.RemoteRoutingAddress
                                    ExpectedAddressSource = $plan.ExpectedAddressSource
                                    AddedProxy           = $expectedAddress
                                    PrimarySmtp          = $email
                                    When                 = (Get-Date)
                                })
                            }
                            else {
                                $status = 'Missing -> Skipped by WhatIf'
                            }
                        }
                        catch {
                            $addFailedCount++
                            $status = 'Missing -> AddFailed'
                            Write-Warning "[$($rec.Identity)] Failed to add $expectedAddress : $($_.Exception.Message)`n$($_.ScriptStackTrace)"
                        }
                    }
                    else {
                        $status = 'Missing'
                    }

                    if ($status -like 'Missing*') {
                        $postMissing++
                    }
                }
            }
        }
        else {
            $status = 'No primary address'
            $noPrimaryCount++
        }

        if ($status -notlike 'Missing*') {
            $suggestedUniqueAddress = ''
            $suggestedUniqueAddressReason = ''
        }

        $results.Add([PSCustomObject]@{
            Identity                                  = $rec.Identity
            Alias                                     = $plan.Alias
            SamAccountName                            = $plan.SamAccountName
            DisplayName                               = $rec.DisplayName
            RecipientType                             = $rec.RecipientTypeDetails
            MailboxLocation                           = $plan.MailboxLocation
            PrimaryAddress                            = $email
            RemoteRoutingAddress                      = $plan.RemoteRoutingAddress
            RemoteRoutingAddressSuffixMatchesExpected = $plan.RemoteRoutingAddressSuffixMatchesExpected
            ExpectedAddress                           = $expectedAddress
            ExpectedAddressSource                     = $plan.ExpectedAddressSource
            ExpectedAddressConflict                   = $isExpectedAddressConflict
            ExpectedAddressConflictReason             = ($conflictReason -join ';')
            ExpectedAddressConflictOwners             = ($conflictOwners -join '; ')
            ExpectedAddressDuplicatePeers             = ($duplicateExpectedPeers -join '; ')
            SuggestedUniqueAddress                    = $suggestedUniqueAddress
            SuggestedUniqueAddressReason              = $suggestedUniqueAddressReason
            RoutingWarning                            = $plan.RoutingWarning
            Status                                    = $status
            EmailAddressPolicyEnabled                 = $rec.EmailAddressPolicyEnabled
            PolicyWarning                             = $policyWarning
        })
    }
    $duplicateAliasGroups = @($results | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Alias) } | Group-Object Alias | Where-Object { $_.Count -gt 1 } | Sort-Object @{Expression='Count';Descending=$true}, Name)
    foreach ($duplicateAliasGroup in $duplicateAliasGroups) {
        foreach ($duplicateAliasRecipient in @($duplicateAliasGroup.Group | Sort-Object DisplayName, Identity)) {
            $duplicateAliasRows.Add([PSCustomObject]@{
                Alias                         = $duplicateAliasGroup.Name
                DuplicateAliasCount           = $duplicateAliasGroup.Count
                Identity                      = $duplicateAliasRecipient.Identity
                SamAccountName                = $duplicateAliasRecipient.SamAccountName
                DisplayName                   = $duplicateAliasRecipient.DisplayName
                RecipientType                 = $duplicateAliasRecipient.RecipientType
                MailboxLocation               = $duplicateAliasRecipient.MailboxLocation
                PrimaryAddress                = $duplicateAliasRecipient.PrimaryAddress
                RemoteRoutingAddress          = $duplicateAliasRecipient.RemoteRoutingAddress
                ExpectedAddress               = $duplicateAliasRecipient.ExpectedAddress
                ExpectedAddressSource         = $duplicateAliasRecipient.ExpectedAddressSource
                ExpectedAddressDuplicatePeers = $duplicateAliasRecipient.ExpectedAddressDuplicatePeers
                SuggestedUniqueAddress        = $duplicateAliasRecipient.SuggestedUniqueAddress
                SuggestedUniqueAddressReason  = $duplicateAliasRecipient.SuggestedUniqueAddressReason
                Status                        = $duplicateAliasRecipient.Status
                EmailAddressPolicyEnabled     = $duplicateAliasRecipient.EmailAddressPolicyEnabled
            })
        }
    }

    Write-Progress -Activity "Checking EmailAddresses..." -Completed
    if ($policyEnabledCount -gt 0) {
        Write-Host "Email address policy enabled recipients: $policyEnabledCount. Details are available in the detail CSV."
    }
    if ($duplicateExpectedAddressGroupCount -gt 0) {
        Write-Host "Duplicate expected proxy addresses: $duplicateExpectedAddressGroupCount group(s), $duplicateExpectedRecipientCount recipient(s), $duplicateExpectedMissingCount missing address(es) blocked." -ForegroundColor Yellow
    }
    if ($addressAlreadyAssignedConflictAddresses.Count -gt 0) {
        Write-Host "Expected addresses already assigned to another Exchange recipient: $($addressAlreadyAssignedConflictAddresses.Count) address(es), $addressAlreadyAssignedMissingCount missing address(es) blocked." -ForegroundColor Yellow
    }
    if ($remoteAliasFallbackMissingBlockedCount -gt 0) {
        Write-Host "Remote mailboxes missing RemoteRoutingAddress: $remoteAliasFallbackMissingBlockedCount address(es) blocked from remediation." -ForegroundColor Yellow
    }
    if ($remoteRoutingSuffixMismatchCount -gt 0) {
        Write-Host "RemoteRoutingAddress suffix mismatches: $remoteRoutingSuffixMismatchCount. Review RoutingWarning in the detail CSV." -ForegroundColor Yellow
    }

    $publishResults = @()
    $addedPublish = $null
    $detailPublish = Publish-SmartM365Csv -Data @($results | Sort-Object Status, Identity) -TimestampedPath $outDetail -LatestPath $latestDetail
    if ($detailPublish) { $publishResults += $detailPublish }
    if ($addedOperations.Count -gt 0) {
        $addedPublish = Publish-SmartM365Csv -Data @($addedOperations) -TimestampedPath $outAdded -LatestPath $latestAdded
        if ($addedPublish) { $publishResults += $addedPublish }
    }

    $summary = @(
        [PSCustomObject]@{ Summary = "Total recipients processed";                         Count = $results.Count },
        [PSCustomObject]@{ Summary = "Existing proxy addresses listed";                  Count = $existingProxyAddressRows.Count },
        [PSCustomObject]@{ Summary = "Duplicate alias groups";                            Count = $duplicateAliasGroups.Count },
        [PSCustomObject]@{ Summary = "Recipients sharing duplicate alias";                Count = $duplicateAliasRows.Count },
        [PSCustomObject]@{ Summary = "On-premises mailboxes processed";                    Count = $localMailboxCount },
        [PSCustomObject]@{ Summary = "Remote mailboxes processed";                         Count = $remoteMailboxCount },
        [PSCustomObject]@{ Summary = "With expected address present";                      Count = $okCount },
        [PSCustomObject]@{ Summary = "With expected address missing";                      Count = $missingCount },
        [PSCustomObject]@{ Summary = "Planned address additions if Write is enabled";       Count = $plannedChangeRows.Count },
        [PSCustomObject]@{ Summary = "Pre-remediation missing";                            Count = $preMissing },
        [PSCustomObject]@{ Summary = "Post-remediation still missing";                     Count = $postMissing },
        [PSCustomObject]@{ Summary = "With no primary address";                            Count = $noPrimaryCount },
        [PSCustomObject]@{ Summary = "With no Alias";                                      Count = $noAliasCount },
        [PSCustomObject]@{ Summary = "With no resolvable expected address";                Count = $noExpectedAddressCount },
        [PSCustomObject]@{ Summary = "RemoteRoutingAddress used";                          Count = $remoteRoutingAddressCount },
        [PSCustomObject]@{ Summary = "Remote mailboxes using Alias fallback";              Count = $remoteAliasFallbackCount },
        [PSCustomObject]@{ Summary = "RemoteRoutingAddress suffix mismatches";             Count = $remoteRoutingSuffixMismatchCount },
        [PSCustomObject]@{ Summary = "Remote Alias fallback missing addresses blocked";     Count = $remoteAliasFallbackMissingBlockedCount },
        [PSCustomObject]@{ Summary = "RemoteMailbox lookup misses";                        Count = $remoteMailboxLookupMissCount },
        [PSCustomObject]@{ Summary = "With email address policy enabled";                  Count = $policyEnabledCount },
        [PSCustomObject]@{ Summary = "Additions skipped by email policy";                  Count = $policySkippedCount },
        [PSCustomObject]@{ Summary = "Duplicate expected address groups";                  Count = $duplicateExpectedAddressGroupCount },
        [PSCustomObject]@{ Summary = "Recipients sharing expected address";                Count = $duplicateExpectedRecipientCount },
        [PSCustomObject]@{ Summary = "Missing addresses blocked by duplicate";             Count = $duplicateExpectedMissingCount },
        [PSCustomObject]@{ Summary = "Expected addresses assigned to another recipient";   Count = $addressAlreadyAssignedConflictAddresses.Count },
        [PSCustomObject]@{ Summary = "Missing addresses blocked because already assigned"; Count = $addressAlreadyAssignedMissingCount },
        [PSCustomObject]@{ Summary = "Addresses successfully added";                       Count = $addedCount },
        [PSCustomObject]@{ Summary = "Address additions failed";                           Count = $addFailedCount }
    )
    $summaryPublish = Publish-SmartM365Csv -Data @($summary) -TimestampedPath $outSummary -LatestPath $latestSummary
    if ($summaryPublish) { $publishResults += $summaryPublish }

    $workbookCsvFiles = @(
        [pscustomobject]@{
            Path          = $detailPublish.TimestampedPath
            WorksheetName = 'Check'
            TableName     = 'ProxyAddressesCheck'
            EmptyColumns  = @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Identity','Alias','SamAccountName','DisplayName','RecipientType','MailboxLocation','PrimaryAddress','RemoteRoutingAddress','RemoteRoutingAddressSuffixMatchesExpected','ExpectedAddress','ExpectedAddressSource','ExpectedAddressConflict','ExpectedAddressConflictReason','ExpectedAddressConflictOwners','ExpectedAddressDuplicatePeers','SuggestedUniqueAddress','SuggestedUniqueAddressReason','RoutingWarning','Status','EmailAddressPolicyEnabled','PolicyWarning')
        }
        [pscustomobject]@{
            Rows          = @($existingProxyAddressRows | Sort-Object Identity, AddressIndex)
            WorksheetName = 'ExistingProxyAddresses'
            TableName     = 'ProxyAddressesExisting'
            EmptyColumns  = @('Identity','Alias','SamAccountName','DisplayName','RecipientType','MailboxLocation','PrimaryAddress','RemoteRoutingAddress','AddressIndex','ProxyAddress','ProxyAddressType','ProxyAddressValue','NormalizedSmtpAddress','IsPrimarySmtp','IsExpectedAddress','EmailAddressPolicyEnabled')
        }
        [pscustomobject]@{
            Rows          = @($duplicateAliasRows | Sort-Object @{Expression='DuplicateAliasCount';Descending=$true}, Alias, DisplayName)
            WorksheetName = 'DuplicateAliases'
            TableName     = 'ProxyAddressesDuplicateAliases'
            EmptyColumns  = @('Alias','DuplicateAliasCount','Identity','SamAccountName','DisplayName','RecipientType','MailboxLocation','PrimaryAddress','RemoteRoutingAddress','ExpectedAddress','ExpectedAddressSource','ExpectedAddressDuplicatePeers','SuggestedUniqueAddress','SuggestedUniqueAddressReason','Status','EmailAddressPolicyEnabled')
        }
        [pscustomobject]@{
            Rows          = @($plannedChangeRows | Sort-Object MailboxLocation, DisplayName, Identity)
            WorksheetName = 'PlannedChanges'
            TableName     = 'ProxyAddressesPlannedChanges'
            EmptyColumns  = @('Action','ExecutionMode','WouldRunInWriteMode','Identity','Alias','SamAccountName','DisplayName','RecipientType','MailboxLocation','PrimarySmtp','RemoteRoutingAddress','ExpectedAddressSource','PlannedProxy','CurrentStatus','Reason','EmailAddressPolicyEnabled','Cmdlet','CommandPreview')
        }
        [pscustomobject]@{
            Path          = $summaryPublish.TimestampedPath
            WorksheetName = 'Summary'
            TableName     = 'ProxyAddressesSummary'
            EmptyColumns  = @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Summary','Count')
        }
        [pscustomobject]@{
            Path          = if ($addedPublish) { $addedPublish.TimestampedPath } else { $null }
            WorksheetName = 'Added'
            TableName     = 'ProxyAddressesAdded'
            EmptyColumns  = @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Identity','Alias','SamAccountName','DisplayName','RecipientType','MailboxLocation','RemoteRoutingAddress','ExpectedAddressSource','AddedProxy','PrimarySmtp','When')
        }
    )
    New-ProxyAddressesWorkbook -CsvFiles $workbookCsvFiles -Path $outWorkbook | Out-Null
    WriteLog -Message "Excel workbook exported to: $outWorkbook"

    if ($latestWorkbook) {
        $latestWorkbookFolder = Split-Path -Path $latestWorkbook -Parent
        if (-not (Test-Path -LiteralPath $latestWorkbookFolder)) {
            New-Item -Path $latestWorkbookFolder -ItemType Directory -Force | Out-Null
        }
        Copy-Item -LiteralPath $outWorkbook -Destination $latestWorkbook -Force
        WriteLog -Message "Excel latest copy written to: $latestWorkbook"
    }

    $workbookSharePointUploads = @()
    foreach ($workbookUploadPath in @($outWorkbook, $latestWorkbook) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique) {
        $workbookUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $workbookUploadPath
        if ($workbookUpload) { $workbookSharePointUploads += $workbookUpload }
    }
    # The workbook bypasses Publish-SmartM365Csv: apply the same seven-day SharePoint retention once this run's copy is uploaded.
    $outWorkbookFullPath = [IO.Path]::GetFullPath($outWorkbook)
    if (@($workbookSharePointUploads | Where-Object { [string]::Equals([IO.Path]::GetFullPath([string]$_.LocalFilePath), $outWorkbookFullPath, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0) {
        Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath $outWorkbook -RetentionDays 7 | Out-Null
    }
    if ($workbookSharePointUploads.Count -gt 0) {
        $publishResults += [pscustomobject]@{ SharePointUploads = @($workbookSharePointUploads) }
    }

    $mailLinkRows = @($publishResults | ForEach-Object { $_.SharePointUploads } |
        Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.WebUrl) } |
        ForEach-Object {
            [pscustomobject]@{
                TenantKey = [string]$global:SmartM365TenantKey
                FileName = if ($_.FileName) { [string]$_.FileName } else { Split-Path -Path ([string]$_.LocalFilePath) -Leaf }
                SharePointPath = [string]$_.SharePointPath
                WebUrl = [string]$_.WebUrl
            }
        } | Sort-Object FileName, WebUrl -Unique)
    if ($mailLinkRows.Count -gt 0) {
        $mailLinksTimestamped = Join-Path $OutputPath ("Exchange_OnPrem_ProxyAddresses_MailLinks_{0}.csv" -f $timestamp)
        $mailLinksLatest = if ($LatestCsvFolderPath) { Join-Path $LatestCsvFolderPath 'Exchange_OnPrem_ProxyAddresses_MailLinks.csv' } else { $null }
        if (Test-SmartM365MaxItemsMode) {
            $mailLinksTimestamped = Add-SmartM365MaxItemsSuffixToCsvPath -Path $mailLinksTimestamped
            if ($mailLinksLatest) { $mailLinksLatest = Add-SmartM365MaxItemsSuffixToCsvPath -Path $mailLinksLatest }
        }
        $null = Publish-SmartM365Csv -Data $mailLinkRows -TimestampedPath $mailLinksTimestamped -LatestPath $mailLinksLatest
        WriteLog -Message ("ProxyAddresses SharePoint link CSV published: {0}" -f $mailLinksLatest)
    }

    Write-Host "`n===== Summary ====="
    foreach ($item in $summary) {
        switch ($item.Summary) {
            "With expected address missing"       { Write-Host "$($item.Summary): $($item.Count)" -ForegroundColor Yellow }
            "With expected address present"       { Write-Host "$($item.Summary): $($item.Count)" -ForegroundColor Green }
            default                               { Write-Host "$($item.Summary): $($item.Count)" }
        }
    }

    Write-Host "Detail: $outDetail"
    if (Test-Path $outAdded)        { Write-Host "Added: $outAdded" }
    Write-Host "Summary: $outSummary"
    Write-Host "Excel: $outWorkbook"

WriteLog -Message 'ProxyAddresses report email is disabled; CSV and workbook outputs are available for the mailbox summary.'

    try {
        if ($script:ChangedViewEntireForest -and $script:PrevViewEntireForest -ne $null) {
            Set-ADServerSettings -ViewEntireForest $script:PrevViewEntireForest
            $script:ChangedViewEntireForest=$false
            Write-Host "ViewEntireForest restored to initial state ($script:PrevViewEntireForest)."
        }
    } catch {
        Write-Warning "Unable to restore ViewEntireForest : $($_.Exception.Message)`n$($_.ScriptStackTrace)"
    }

    #region Cleanup
	# Clean up old CSV, Excel, and log files.
	# Automatically excludes all generated CSVs via global:csvGeneratedPaths + current transcript and log files via global variables.
	Remove-SmartM365TimestampedFilesOlderThan -FolderPath $OutputPath -FilePattern '*.csv' -RetentionDays 7 -LogFile $global:logTextFile
	Remove-SmartM365TimestampedFilesOlderThan -FolderPath $OutputPath -FilePattern 'Exchange_OnPrem_ProxyAddresses_*.xlsx' -RetentionDays 7 -LogFile $global:logTextFile
	RemoveOldFiles -Path $logPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
    WriteLog -Message "$TaskName completed."
    try {
        $smartM365TranscriptPath = $null
        $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue
        if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
            $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
        }
        else {
            $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue
            if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value }
        }
        Stop-SmartM365RunTranscript | Out-Null
        if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath }
    }
    catch {}
    Complete-SmartM365ExecutionContext -Status Success
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard -Cleanup {
        $changed=Get-Variable ChangedViewEntireForest -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        $previous=Get-Variable PrevViewEntireForest -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        if($changed -and $null -ne $previous){
            Invoke-SmartM365RunCleanup $smartM365RunGuard 'AD session settings' { Set-ADServerSettings -ViewEntireForest $previous; $script:ChangedViewEntireForest=$false }
        }
    }
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}

    #endregion

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDdDxSEU4YPx7bf
# Vxs6BX9TvbuzV8hTntwcxoZ81x+Q3KCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg3P8FW1exh1An
# tIoe1mx9IaZeK6tUkDyf2jI/cktlHVEwDQYJKoZIhvcNAQEBBQAEggGAKzmX7H11
# Iv3NF65crvc6vp12/QE9WkWpYwpSKx6RHIkicC+HFAhGjWSm2xkzhTWlUZbPPoVg
# KSBIZljDMctJfGJ40fAloSJjOOv7ABnQDKBv45rOa7nG9r9HXBYdAcAQ254Tiehz
# +Pr+N26RizhmgXIJRKRnRFjZeim3qunJU/SWOHNafUnYS6zQeeyd3Qurn3nRgTQF
# jchg6Md1js58AyuAs1LMDFy8DSpCsjs/MB+sUjaSHfu7rOr4RzbdJABe955pr/BC
# roTWbMW1PJ4/06UQ9M7ZYSIP+9cgI1L8cFTWmx9EkdRvDdPh4ANhSAuCOtalD8x4
# DF4M6uFNH1ukutTYwsqa8oBrB2ejZaI+6qXNcrOsyXwmaVEEONit5YFT4C8mO5KS
# ZJQdVgnjhUexM+HRyv0Ohvi9mec9Mf8I2vymrFTKqZLiUe2PEdYChrjt09EUDgHe
# dbI95KAeUFfvjy+LrVoVZEt5MWD4GCipRbAuj1svRyhSeSoFl1vEou5AoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTAxMTY1OTEwWjA/BgkqhkiG9w0BCQQxMgQweSEkXjgl7j5J5ClqQIeCEpCk
# c7tmZpGM5UfI2mROITUc/ib8QhKLIlNTdXGhVnUkMA0GCSqGSIb3DQEBAQUABIIC
# AEa1rPXCjMul6p8zNIgGqIJXC/uRBxUYf/gACEwvOmVE2z64/+FVRN/Go80iDXqt
# vyp/DVtS0X2OLDkYs9SgiOaOuYuK8JiQ6TKRDfmX6Q0JEDRfcCvTYwnrUG0KdlRw
# JXOIpeo9G7KVMnSQAasLSxOefgnaxFk00bDRuyAszwPKPw6mA7wU135MJxMZ9/c3
# Uh7Nf1RfGV/vFjm4QSlFL8gAbjcHKp8ajSZLjw6lP/Ezok/iHwJx/rsIjOY9EwSc
# QEtQr7UN1kAguBvLDbw7zoaBD+s9qUT07UYYo+a+n9zz0H9UaDylMV1WL+S9T4rU
# LPeiZD4EHmqGVoqPjNtOg3gWN8SP0/Gq1sVyCQvISq2niOl5WxezLAQyPXDfopJn
# t/kpGXjUwElv51Cdoe/ykROcN3QFSqu1+QQBXewh4KR/M5en2scSy6jgDrx/+iHR
# 3X8E22v4lVBFNIPJbLKntQXV8xwAcjV+yTvmslQ2bWWP4lhPxPiRCkRaDdodPIUe
# IjO5gSCjNqNVZWPCdKIc3Zcc3CaH08rFX3M8t+F9cYVYT/Yp0lMahkRS7Mz8f1TU
# qMi8FKXwAOlfUDUXI6chs+KNHsj+w9swlLyjPzOwglrcRGrG69vzov9UsKxv0TZQ
# s8Og6hLhkHmxZ47e7+RkFnYd0+gNfpM3ktlSf9i8VbKC
# SIG # End signature block
