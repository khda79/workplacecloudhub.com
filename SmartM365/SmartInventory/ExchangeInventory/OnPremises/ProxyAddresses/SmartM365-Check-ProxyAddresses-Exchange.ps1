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
    - Generates detailed, summary, and added addresses CSV reports plus an Excel workbook with Check, ExistingProxyAddresses, DuplicateAliases, Summary, and Added worksheets.
    - Maintains logs and cleans up old files automatically.

.VERSION
1.21

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
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
            Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null
        }
        else {
            if (-not (Test-Path -LiteralPath $templatePath)) {
                $message = @(
                    "Local configuration file not found: $configPath",
                    "Template to copy is also missing: $templatePath",
                    'Create the .local.json file from a safe template, then run the script again.'
                ) -join [Environment]::NewLine
                throw $message
            }

            Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
            Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
            Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
        }
    }

    try {
        return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
}

function Resolve-SmartM365ConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    if ($Value -notmatch '\{\{[^}]+\}\}') {
        return $Value
    }

    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($ScriptRoot) { $ScriptRoot } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }

            $tokenValue = Resolve-SmartM365ConfigValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}
function Get-ScriptLocalConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                return Resolve-SmartM365ConfigValue -Value $property.Value
            }
        }
        else {
            return Resolve-SmartM365ConfigValue -Value $property.Value
        }
    }


    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $PSCommandPath -Parent }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $globalProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) {
            return $DefaultValue
        }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

function New-ProxyAddressesWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object[]]$CsvFiles,
        [Parameter(Mandatory = $true)][string]$Path
    )

    Import-Module -Name ImportExcel -ErrorAction Stop

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
            AutoSize      = $true
            FreezeTopRow  = $true
            BoldTopRow    = $true
            AutoFilter    = $true
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
    if (-not $PSBoundParameters.ContainsKey('OrganizationalUnit')) { $OrganizationalUnit = @(Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrganizationalUnit' -DefaultValue $OrganizationalUnit) }
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
    $ScriptVersion = "1.21"
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
        Import-Module -Name (Join-ModulePath 'SmartM365-WindowsPowerShell5.psd1') -MinimumVersion '1.0.18' -ErrorAction Stop
        $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
        Start-Transcript -Path $global:logTranscriptFile -Append
        WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
        $OutputPath = $InitializeOutputPath
        WriteLog -Message "Starting $TaskName..."
    } catch {
        Write-Host "Initialization failed: $($_.Exception.Message)`n$($_.ScriptStackTrace)" -ForegroundColor Red
        exit
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

    # -------------------------------
    # Exchange PSSnapin availability
    # -------------------------------
    $snapinName = "Microsoft.Exchange.Management.PowerShell.SnapIn"

    if (-not (Get-PSSnapin $snapinName -Registered -ErrorAction SilentlyContinue)) {
        Write-Error "The Exchange Management PSSnapin '$snapinName' is not registered on this server.`nThis script must be run on an Exchange 2016 server where the Management Tools are installed."
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        exit 1
    }

    if (-not (Get-PSSnapin $snapinName -ErrorAction SilentlyContinue)) {
        Write-Host "The Exchange PSSnapin '$snapinName' is not loaded in the current session. Attempting to load it..."
        try {
            Add-PSSnapin $snapinName -ErrorAction Stop
            Write-Host "The Exchange PSSnapin was loaded successfully."
        }
        catch {
            Write-Error "Failed to load the Exchange PSSnapin '$snapinName'. Error: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            Write-Error "Ensure you are running this script on an Exchange 2016 server and have the necessary permissions.`nAlternatively, try running this script from the Exchange Management Shell or use a script that connects via PowerShell Remoting to localhost."
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            exit 1
        }
    } else {
        Write-Host "The Exchange PSSnapin '$snapinName' is already loaded."
    }

    $requiredExchangeCommands = @('Get-Recipient', 'Get-RemoteMailbox')
    if ($AddMissingAddress) { $requiredExchangeCommands += @('Set-Mailbox', 'Set-RemoteMailbox') }
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredModules @('ImportExcel') -RequiredCommands $requiredExchangeCommands -RequireExchangeOnPrem -RequireActiveDirectoryRead | Out-Null
    Import-Module -Name ImportExcel -ErrorAction Stop
    $importExcelModule = Get-Module -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
    WriteLog -Message ("ImportExcel module loaded: version={0}; path={1}" -f $importExcelModule.Version, $importExcelModule.Path)

    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-Error "The Get-Mailbox cmdlet is still not available after attempting to load the snap-in.`nThis could indicate an issue with the Exchange Management Tools installation."
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        exit 1
    }

    if (-not (Get-Command Get-Recipient -ErrorAction SilentlyContinue)) {
        Write-Error "Get-Recipient is unavailable. Ensure the Exchange Management Tools are installed, or run from the Exchange Management Shell."
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        exit 1
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
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
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
    if ($workbookSharePointUploads.Count -gt 0) {
        $publishResults += [pscustomobject]@{ SharePointUploads = @($workbookSharePointUploads) }
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

# ===========================
# === Email notification ====
# ===========================
try {
    function Encode([string]$s) { return (ConvertTo-SmartM365EmailHtmlText $s) }

    $MailTo = @($To) -split '[;,]\s*' | Where-Object { $_ -and $_.Trim() -ne '' }
    $MailCc = @($Cc) -split '[;,]\s*' | Where-Object { $_ -and $_.Trim() -ne '' }

    $MailFrom    = $From
    $MailSubject = $Subject

    # Attach only the consolidated workbook when it is small enough for mail transport.
    # CSV files and the workbook always remain available through the body paths/SharePoint links.
    $attachments = @()
    $workbookAttachmentNote = $null
    if (Test-Path -LiteralPath $outWorkbook -PathType Leaf) {
        $workbookItem = Get-Item -LiteralPath $outWorkbook -ErrorAction SilentlyContinue
        if ($workbookItem) {
            $workbookSizeMB = [Math]::Round(($workbookItem.Length / 1MB), 2)
            if ($MaxMailAttachmentMB -le 0) {
                $workbookAttachmentNote = "Excel workbook attachment disabled by MaxMailAttachmentMB=0. Workbook size: $workbookSizeMB MB."
                WriteLog -Message $workbookAttachmentNote -Level 'WARNING'
            }
            elseif ($workbookItem.Length -le ($MaxMailAttachmentMB * 1MB)) {
                $attachments += $outWorkbook
            }
            else {
                $workbookAttachmentNote = "Excel workbook not attached because size $workbookSizeMB MB exceeds MaxMailAttachmentMB=$MaxMailAttachmentMB MB. Use the body path or SharePoint link instead."
                WriteLog -Message $workbookAttachmentNote -Level 'WARNING'
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($MailFrom) -or -not $MailTo) {
        Write-Host "Email skipped: incomplete email parameters (From/To)."
    }
    else {
        $totalCount = [int](($summary | Where-Object { $_.Summary -eq 'Total recipients processed' } | Select-Object -First 1).Count)
        $presentCount = [int](($summary | Where-Object { $_.Summary -eq 'With expected address present' } | Select-Object -First 1).Count)
        $missingCount = [int](($summary | Where-Object { $_.Summary -eq 'With expected address missing' } | Select-Object -First 1).Count)
        $localMailboxCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'On-premises mailboxes processed' } | Select-Object -First 1).Count)
        $remoteMailboxCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Remote mailboxes processed' } | Select-Object -First 1).Count)
        $noAliasCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'With no Alias' } | Select-Object -First 1).Count)
        $noExpectedAddressCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'With no resolvable expected address' } | Select-Object -First 1).Count)
        $remoteRoutingAddressCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'RemoteRoutingAddress used' } | Select-Object -First 1).Count)
        $remoteAliasFallbackCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Remote mailboxes using Alias fallback' } | Select-Object -First 1).Count)
        $remoteRoutingSuffixMismatchCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'RemoteRoutingAddress suffix mismatches' } | Select-Object -First 1).Count)
        $remoteAliasFallbackMissingBlockedCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Remote Alias fallback missing addresses blocked' } | Select-Object -First 1).Count)
        $remoteMailboxLookupMissCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'RemoteMailbox lookup misses' } | Select-Object -First 1).Count)
        $addedCount = [int](($summary | Where-Object { $_.Summary -eq 'Addresses successfully added' } | Select-Object -First 1).Count)
        $policyEnabledCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'With email address policy enabled' } | Select-Object -First 1).Count)
        $policySkippedCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Additions skipped by email policy' } | Select-Object -First 1).Count)
        $duplicateAliasGroupCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Duplicate alias groups' } | Select-Object -First 1).Count)
        $duplicateAliasRecipientCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Recipients sharing duplicate alias' } | Select-Object -First 1).Count)
        $duplicateExpectedAddressGroupCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Duplicate expected address groups' } | Select-Object -First 1).Count)
        $duplicateExpectedRecipientCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Recipients sharing expected address' } | Select-Object -First 1).Count)
        $duplicateExpectedMissingCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Missing addresses blocked by duplicate' } | Select-Object -First 1).Count)
        $addressAlreadyAssignedCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Expected addresses assigned to another recipient' } | Select-Object -First 1).Count)
        $addressAlreadyAssignedMissingCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'Missing addresses blocked because already assigned' } | Select-Object -First 1).Count)
        $effectiveSendMailMode = if ([string]::IsNullOrWhiteSpace($SendMailMode)) { if ([string]::IsNullOrWhiteSpace($SmtpServer)) { 'Graph' } else { 'SMTP' } } else { $SendMailMode.Trim() }

        $summaryRowsForEmail = @(
            [pscustomobject]@{ Label = 'Total recipients processed'; Value = $totalCount }
            [pscustomobject]@{ Label = 'On-premises mailboxes'; Value = $localMailboxCountForMail }
            [pscustomobject]@{ Label = 'Remote mailboxes'; Value = $remoteMailboxCountForMail }
            [pscustomobject]@{ Label = 'With expected address present'; Value = $presentCount }
            [pscustomobject]@{ Label = 'With expected address missing'; Value = $missingCount }
            [pscustomobject]@{ Label = 'With no Alias'; Value = $noAliasCountForMail }
            [pscustomobject]@{ Label = 'With no resolvable expected address'; Value = $noExpectedAddressCountForMail }
            [pscustomobject]@{ Label = 'RemoteRoutingAddress used'; Value = $remoteRoutingAddressCountForMail }
            [pscustomobject]@{ Label = 'Remote mailboxes using Alias fallback'; Value = $remoteAliasFallbackCountForMail }
            [pscustomobject]@{ Label = 'RemoteRoutingAddress suffix mismatches'; Value = $remoteRoutingSuffixMismatchCountForMail }
            [pscustomobject]@{ Label = 'Remote Alias fallback missing addresses blocked'; Value = $remoteAliasFallbackMissingBlockedCountForMail }
            [pscustomobject]@{ Label = 'RemoteMailbox lookup misses'; Value = $remoteMailboxLookupMissCountForMail }
            [pscustomobject]@{ Label = 'Email address policy enabled'; Value = $policyEnabledCountForMail }
            [pscustomobject]@{ Label = 'Additions skipped by email policy'; Value = $policySkippedCountForMail }
            [pscustomobject]@{ Label = 'Duplicate alias groups'; Value = $duplicateAliasGroupCountForMail }
            [pscustomobject]@{ Label = 'Recipients sharing duplicate alias'; Value = $duplicateAliasRecipientCountForMail }
            [pscustomobject]@{ Label = 'Duplicate expected address groups'; Value = $duplicateExpectedAddressGroupCountForMail }
            [pscustomobject]@{ Label = 'Recipients sharing expected address'; Value = $duplicateExpectedRecipientCountForMail }
            [pscustomobject]@{ Label = 'Missing addresses blocked by duplicate'; Value = $duplicateExpectedMissingCountForMail }
            [pscustomobject]@{ Label = 'Expected addresses assigned elsewhere'; Value = $addressAlreadyAssignedCountForMail }
            [pscustomobject]@{ Label = 'Missing addresses blocked because assigned'; Value = $addressAlreadyAssignedMissingCountForMail }
            [pscustomobject]@{ Label = 'Addresses added'; Value = $addedCount }
        )

        $pathRows = @(
            [pscustomobject]@{ Label = 'Detail CSV'; Path = $outDetail }
            [pscustomobject]@{ Label = 'Summary CSV'; Path = $outSummary }
            [pscustomobject]@{ Label = 'Excel workbook'; Path = $outWorkbook }
        )
        if (Test-Path $outAdded) { $pathRows += [pscustomobject]@{ Label = 'Added CSV'; Path = $outAdded } }

        $scopeHtml = if ($AllOrganizationalUnit) { 'ALL (entire forest)' } else { ($OrganizationalUnit | ForEach-Object { Encode $_ }) -join '<br/>' }
        $modeLabel = if ($AddMissingAddress) { 'Write mode' } else { 'Read-only mode' }
        $scopeSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr><td style="width:180px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Mode</td><td style="border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;">$(Encode $modeLabel)</td></tr>
  <tr><td style="width:180px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Expected suffix</td><td style="border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;">$(Encode $ExpectedSuffix)</td></tr>
  <tr><td style="width:180px;background:#f8fafc;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Scope</td><td style="padding:10px 12px;font-size:13px;color:#334155;word-break:break-all;">$scopeHtml</td></tr>
</table>
"@

        $sections = @([pscustomobject]@{ Title = 'Scope'; Html = $scopeSectionHtml })
        if (-not [string]::IsNullOrWhiteSpace($workbookAttachmentNote)) {
            $attachmentSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #fde68a;background:#fffbeb;">
  <tr><td style="padding:12px 14px;font-size:13px;color:#92400e;word-break:break-word;">$(Encode $workbookAttachmentNote)</td></tr>
</table>
"@
            $sections += [pscustomobject]@{ Title = 'Email attachment notice'; Html = $attachmentSectionHtml }
        }

        $duplicateAliasPreviewRows = @($duplicateAliasRows | Sort-Object @{Expression='DuplicateAliasCount';Descending=$true}, Alias, DisplayName | Select-Object -First 50)
        if ($duplicateAliasPreviewRows.Count -gt 0) {
            $duplicateAliasRowsHtml = foreach ($row in $duplicateAliasPreviewRows) {
                "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(Encode $row.Alias)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(Encode ([string]$row.DuplicateAliasCount))</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(Encode $row.DisplayName)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.PrimaryAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.ExpectedAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.SuggestedUniqueAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;font-weight:700;color:#92400e;`">$(Encode $row.Status)</td></tr>"
            }
            $duplicateAliasSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Alias</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Count</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Display name</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Primary SMTP</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Expected proxy</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Suggested proxy</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Status</th>
  </tr>
  $($duplicateAliasRowsHtml -join "`n")
</table>
"@
            $sections += [pscustomobject]@{ Title = 'Top 50 duplicate aliases'; Html = $duplicateAliasSectionHtml }
        }

        $missingPreviewRows = @($results | Where-Object { $_.Status -like 'Missing*' } | Sort-Object Status, DisplayName | Select-Object -First 50)
        if ($missingPreviewRows.Count -gt 0) {
            $missingRowsHtml = foreach ($row in $missingPreviewRows) {
                "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(Encode $row.DisplayName)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(Encode $row.Alias)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.PrimaryAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.ExpectedAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;font-weight:700;color:#92400e;`">$(Encode $row.Status)</td></tr>"
            }
            $missingSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Display name</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Alias</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Primary SMTP</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Expected proxy</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Status</th>
  </tr>
  $($missingRowsHtml -join "`n")
</table>
"@
            $sections += [pscustomobject]@{ Title = 'Top 50 missing proxy addresses'; Html = $missingSectionHtml }
        }

        $sharePointRecords = @($publishResults | ForEach-Object { $_.SharePointUploads } | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_.WebUrl) })
        if ($sharePointRecords.Count -gt 0) {
            $sharePointRowsHtml = foreach ($record in $sharePointRecords) {
                $label = if ($record.FileName) { [string]$record.FileName } else { [string]$record.SharePointPath }
                $pathText = if ($record.SharePointPath) { [string]$record.SharePointPath } else { [string]$record.WebUrl }
                "<tr><td style=`"width:220px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;`">$(Encode $label)</td><td style=`"border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;word-break:break-all;`"><a href=`"$(Encode $record.WebUrl)`" style=`"color:#2563eb;text-decoration:underline;`">$(Encode $pathText)</a></td></tr>"
            }
            $sharePointSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  $($sharePointRowsHtml -join "`n")
</table>
"@
            $sections += [pscustomobject]@{ Title = 'SharePoint links'; Html = $sharePointSectionHtml }
        }

        $severity = if ($AddMissingAddress -and $addedCount -gt 0) { 'Success' } elseif ($missingCount -gt 0) { 'Warning' } else { 'Success' }
        $actionTitle = if ($missingCount -gt 0) { 'Review required' } else { '' }
        $actionHtml = if ($missingCount -gt 0) { 'Review missing proxy addresses before remediation. Recipients managed by an email address policy, duplicate expected addresses, addresses already assigned to another mail-enabled recipient, and remote mailboxes without a usable RemoteRoutingAddress are always skipped in write mode.' } else { '' }
        $message = if ($missingCount -gt 0) { 'Exchange on-premises proxy address audit found missing expected proxy addresses.' } else { 'Exchange on-premises proxy address audit completed without missing expected proxy addresses.' }

        $body = New-SmartM365EmailBody `
            -Title $MailSubject `
            -Category 'SmartM365 Exchange OnPrem' `
            -Severity $severity `
            -Tenant $Tenant `
            -HostName $env:COMPUTERNAME `
            -Message $message `
            -ActionTitle $actionTitle `
            -ActionHtml $actionHtml `
            -SummaryRows $summaryRowsForEmail `
            -PathRows $pathRows `
            -Sections $sections `
            -Footer 'This automated message was generated by SmartM365. Use the exported CSV files, Excel workbook, and SharePoint links as the source of truth.'

        SendEmailHtmlReport -SendMailMode $effectiveSendMailMode -SmtpServer $SmtpServer -SmtpPort $SmtpPort -From $MailFrom -To ($MailTo -join ';') -Cc ($MailCc -join ';') -Subject $MailSubject -BodyHtml $body -Attachments $attachments -AllowAttachments
        Write-Host "Email sent to '$($MailTo -join ';')' via $effectiveSendMailMode."
    }
} catch {
    Write-Warning "Email send failed: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
}
# === End Email notification ===

    try {
        if ($script:ChangedViewEntireForest -and $script:PrevViewEntireForest -ne $null) {
            Set-ADServerSettings -ViewEntireForest $script:PrevViewEntireForest
            Write-Host "ViewEntireForest restored to initial state ($script:PrevViewEntireForest)."
        }
    } catch {
        Write-Warning "Unable to restore ViewEntireForest : $($_.Exception.Message)`n$($_.ScriptStackTrace)"
    }

    #region Cleanup
	# Clean up old CSV, Excel, and log files.
	# Automatically excludes all generated CSVs via global:csvGeneratedPaths + current transcript and log files via global variables.
	RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
	RemoveOldFiles -Path $OutputPath -Filter "Exchange_OnPrem_ProxyAddresses_*.xlsx" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
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
        Stop-Transcript | Out-Null
        if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath }
    }
    catch {}
    Complete-SmartM365ExecutionContext -Status Success
    #endregion
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCdRFOhAgH9TR6b
# aFTm1dnSg1X5mVa+i4CFC6SWtY5886CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIF2lnpYeMAmzNUSqi8fcMU9dM2iYVewDKH2suy52T2w+MA0GCSqG
# SIb3DQEBAQUABIIBgDISYGGdNtMSRU5ffM5GuI5HQh6i2AtqSQw1JcX0e1kcQbJ4
# XpXKwqarqZcE40JTKsd5yhePOCkrd1/2edJLeud/j3ZUOClBqc2TtsYYclQXIJj5
# SQVJ/ZqISDPM7jdlq3tP7NIXEqSu6lvz+RmqrmNaqD883gSSnTgENROcg1Z8MAOH
# E4iQmjiKtQyKWLmR+7t1+EAMWtVNb37UrGUkl4+DPuX2COBTS6lnwzaqf1XM/1qd
# C9KSwyn1PwQXm+N3wYiULDzGxhIAvOw41ROPJeUN9ibaTdFjS3xHgTkhrdyhkmvu
# PjFfQxwXlzvNXPiolrMWD5KIwR9M9BEcxztwgCqvtFiUxQBEqrc4PRk0KhEh8UXw
# JMsZVbsr/8GVTGYW+x2Dy4nL6/Wyjc4SzYWkA/YUMXEdXBTEbcHDFQAtkEAmHJoZ
# b3fCcj884gHa0PU8R6h4BxoARUCaJ9Ld9QykdBjcmfPPFiz3Cvkac8l3CT2JsvK0
# pWOBVAl+XZEGvasmMKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTcxNjUw
# MzJaMC8GCSqGSIb3DQEJBDEiBCBntZswj0/Pwb/30ggQmfxcZIgDJ6avQs59gSTV
# OqZ50TANBgkqhkiG9w0BAQEFAASCAgCI3yyKs3aNKxufrZCyA8UT1pTXhANdwKxJ
# 242uI+5/IoykCQ6CWHv1fsrFMI5CJsmuzM/JZnKzq6TJG378BlTkk/h4wsFyX195
# EVhzQ6T7dEzUjSeZSVRi5Ssf1an5BOFkBuDXRwiABHLGZ609zKx2cQkN495xou7P
# aiHIVy4/q0URPgxUWq0ZNNS/Wzc8lpyJRFAa61pSpsvkqe5RJ6sbuCjTkESvzHCw
# WUqJ9+PjukYusPBRcTNR6Xv8cmDyLVKpmCCxO4cU58yzyIni7MtZPGDrXgT8niqW
# AEJiO7kif1sC0SkHZAhp8UA4ajs4RClLMcJasxh6kNC5acRwibWgWCd6l0c6aHj4
# F82qi7NA2NFxV79eL20OAxqyDiiFPed4Rov1daWxVuFVA5awXyMHedbKV4CLhLX4
# XL7ovvNOMzi6ezenNO0lqLRMP7qCaSze8Jag7qz8P4IBswRFeUS9A2H2qHhKUOrf
# 3Pv6bhO9U5AF9jYIjsman9+t2UwFIa4P94XqtiiQE/Dj+CHpl8CAz2LVdTkFxkpt
# CaLqwWqStQ2hoTeGk2RhsFjdcA7snfjammJVw/3OiYnwXA77GSWbwWbUTW0E5HSZ
# zCRCUTh7LslU+hbduPD7W0CyhXY8bXeXvlawly/tsrblvpjbxf0XcljdgrHMtfIf
# IiiCvruzTQ==
# SIG # End signature block
