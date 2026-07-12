<#
.SYNOPSIS
    Audits and optionally remediates proxy addresses for on-premises Exchange **User & Shared** mailboxes in specified Organizational Units (OUs), or all OUs.

.DESCRIPTION
    This script checks whether each **UserMailbox** and **SharedMailbox** has a proxy address with the expected domain suffix (e.g., tenant.mail.onmicrosoft.com).
    If the expected address is missing and the -AddMissingAddress switch is specified, the script can add the address (either restricted by an allowlist CSV, or for all mailboxes when -SkipAllowListCsv is used).
    The script can target multiple OUs via -OrganizationalUnit (string array) or all OUs with -AllOrganizationalUnit.
    It generates detailed, summary, and remediation CSV reports and maintains logs.
    Designed to run on an Exchange 2016 server with the Management Tools installed.

.PARAMETER OrganizationalUnit
    Distinguished names of the target OU(s). Accepts multiple OUs.
    Default: @("DC=example,DC=com")

.PARAMETER AllOrganizationalUnit
    If specified, the script ignores OU filtering and fetches all User/Shared mailboxes in the forest.

.PARAMETER ExpectedSuffix
    The expected domain suffix for proxy addresses (default: "tenant.mail.onmicrosoft.com").

.PARAMETER AddMissingAddress
    If specified, missing proxy addresses will be added depending on allowlist logic.

.PARAMETER SkipAllowListCsv
    If specified, the allowlist CSV is ignored. When used with -AddMissingAddress, the script will add the expected address for ALL mailboxes that are missing it.
    This switch also works without -AddMissingAddress (auditing only) to ignore allowlist evaluation in the reports.

.PARAMETER AllowListCsv
    Path to the CSV file containing the allowlist of mailboxes eligible for address addition (used unless -SkipAllowListCsv is present).

.PARAMETER CsvEmailColumn
    The column name in the allowlist CSV containing the primary SMTP addresses (default: "PrimarySmtpAddress").

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


.REQUIREMENTS
    Windows PowerShell 5.1 on an Exchange 2016/on-premises management host.
    Modules/snap-ins: SmartM365 WindowsPowerShell5 compatibility module; Exchange Management snap-in; ActiveDirectory module.
    Minimum permissions: Exchange on-premises recipient read access and AD read access for proxyAddresses, primary SMTP, UPN and related recipient attributes.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    - Requires Exchange 2016 Management Tools.
    - Generates detailed, summary, and added addresses CSV reports.
    - Maintains logs and cleans up old files automatically.

.VERSION
1.12

.AUTHOR
    https://github.com/khda79/workplacecloudhub.com
    Minimum permissions: Windows PowerShell 5.1, Exchange 2016 Management snap-in, ActiveDirectory module, Exchange recipient read access, and AD read access.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$Tenant = 'test',
[Parameter(Mandatory=$false)]
    [string[]]$OrganizationalUnit = @("DC=example,DC=com"),

    [Parameter(Mandatory=$false)]
    [switch]$AllOrganizationalUnit = $true,

    [Parameter(Mandatory=$false)]
    [string]$ExpectedSuffix = "tenant.mail.onmicrosoft.com",

    [Parameter(Mandatory=$false)]
    [switch]$AddMissingAddress = $false,

    # Skip allowlist usage; active with or without -AddMissingAddress
    [Parameter(Mandatory=$false)]
    [switch]$SkipAllowListCsv = $True,

    [Parameter(Mandatory=$false)]
    [string]$AllowListCsv,

    [Parameter(Mandatory=$false)]
    [string]$CsvEmailColumn = "PrimarySmtpAddress",

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
    $ScriptVersion = "1.12"
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

    if (-not $AllowListCsv) {
        $ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
        $AllowListCsv = Join-Path $ScriptDirectory "Check-ProxyAddresses-Exchange-AllowListCsv.csv"
    }

    $timestamp       = Get-Date -Format 'yyyyMMdd-HHmm'
$outDetail       = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Check_{0}.csv" -f $timestamp)
$outSummary      = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Summary_{0}.csv" -f $timestamp)
$outAdded        = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Added_{0}.csv" -f $timestamp)
$outAllowMissing = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_MissingInAllowList_{0}.csv" -f $timestamp)
$latestDetail       = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Check.csv' } else { $null }
$latestSummary      = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Summary.csv' } else { $null }
$latestAdded        = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Added.csv' } else { $null }
$latestAllowMissing = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_MissingInAllowList.csv' } else { $null }

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
    Write-Host "SkipAllowLst : $SkipAllowListCsv"
    if ($PSBoundParameters.ContainsKey('AllowListCsv')) {
        Write-Host "AllowListCsv : $AllowListCsv"
        Write-Host "CsvEmailCol  : $CsvEmailColumn"
    }
    Write-Host "Detail CSV  : $outDetail"
    Write-Host "Summary CSV : $outSummary"
    Write-Host "Added CSV   : $outAdded"
    Write-Host "Allowlist missing CSV: $outAllowMissing"

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

    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequireExchangeOnPrem -RequireActiveDirectoryRead | Out-Null

    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-Error "The Get-Mailbox cmdlet is still not available after attempting to load the snap-in.`nThis could indicate an issue with the Exchange Management Tools installation."
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        exit 1
    }

    if (-not (Get-Command Get-RemoteMailbox -ErrorAction SilentlyContinue) -and -not (Get-Command Get-Recipient -ErrorAction SilentlyContinue)) {
        Write-Error "Neither Get-RemoteMailbox nor Get-Recipient are available. Ensure the Exchange Management Tools are installed, or run from the Exchange Management Shell."
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

    # --- Allowlist CSV ---
    if ($AddMissingAddress -and -not $SkipAllowListCsv) {
        if ([string]::IsNullOrWhiteSpace($AllowListCsv)) {
            Write-Error "The parameter -AllowListCsv is required when -AddMissingAddress is specified (unless -SkipAllowListCsv is used)."
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            exit 1
        }
    }

    $script:AllowSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $script:AllowListLoaded = $false
    $script:AllowListRowCount = 0

    if (-not $SkipAllowListCsv -and -not [string]::IsNullOrWhiteSpace($AllowListCsv)) {
        if (-not (Test-Path -LiteralPath $AllowListCsv)) {
            Write-Error "CSV file not found: $AllowListCsv"
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            exit 1
        }
        try {
            $rows = Import-Csv -LiteralPath $AllowListCsv
            if (-not $rows) {
                Write-Warning "The CSV is empty: $AllowListCsv - allowlist will not be applied."
            } else {
                foreach ($row in $rows) {
                    if ($null -ne $row.$CsvEmailColumn -and ($row.$CsvEmailColumn).ToString().Trim() -ne "") {
                        [void]$script:AllowSet.Add(($row.$CsvEmailColumn).ToString().Trim())
                        $script:AllowListRowCount++
                    }
                }
                $script:AllowListLoaded = $true
                Write-Host "Allowlist loaded: $script:AllowListRowCount entries (column '$CsvEmailColumn')."
            }
        } catch {
            Write-Error "Failed to load CSV: $AllowListCsv - $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            exit 1
        }
    } elseif ($SkipAllowListCsv) {
        Write-Host "SkipAllowListCsv: allowlist CSV will be ignored; additions (if any) will not be restricted by CSV."
    }

    # Confirmation only when using allowlist CSV (not when skipping)
    if ($AddMissingAddress -and -not $SkipAllowListCsv) {
        Write-Host "You have specified -AddMissingAddress and provided -AllowListCsv: $AllowListCsv"
        $confirmation = Read-Host "Do you want to continue using this AllowListCsv file? (Y/N)"
        if ($confirmation -match '^[Yy]$') {
            Write-Host "Continuing script execution with AllowListCsv: $AllowListCsv"
        } else {
            Write-Error "Script stopped by user: confirmation to use AllowListCsv was denied."
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            exit 1
        }
    }

    # Collections & counters
    $results                = New-Object System.Collections.Generic.List[object]
    $addedOperations        = New-Object System.Collections.Generic.List[object]

    $okCount                = 0
    $missingCount           = 0
    $missingInAllowCount    = 0
    $missingNotInAllowCount = 0
    $noPrimaryCount         = 0
    $addedCount             = 0
    $addFailedCount         = 0
    $policyEnabledCount    = 0

    # Pre/post remediation counters
    $preMissing             = 0
    $postMissing            = 0

    $recipientResultSize = if ($MaxItems -gt 0) { $MaxItems } else { 'Unlimited' }

    # Recipient retrieval: either all OUs (no filter) or each OU provided
    if ($AllOrganizationalUnit) {
        try {
            Write-Host "Fetching recipients from ALL Organizational Units..."
            $recipients = Get-Recipient -ResultSize $recipientResultSize `
                                        -RecipientTypeDetails UserMailbox, SharedMailbox `
                                        -ErrorAction Stop
        } catch {
            Write-Error "Failed to Get-Recipient for ALL OU: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            throw
        }
    } else {
        $allRecipients = @()
        foreach ($ou in $OrganizationalUnit) {
            try {
                Write-Host "Fetching recipients from OU: $ou"
                $recs = Get-Recipient -OrganizationalUnit $ou `
                                      -ResultSize $recipientResultSize `
                                      -RecipientTypeDetails UserMailbox, SharedMailbox `
                                      -ErrorAction Stop
                if ($recs) { $allRecipients += $recs }
            } catch {
                Write-Warning "Get-Recipient failed for OU '$ou' : $($_.Exception.Message)`n$($_.ScriptStackTrace)"
            }
        }
        # Deduplicate by Guid
        $recipients = $allRecipients | Sort-Object -Property Guid -Unique
    }
    if ($MaxItems -gt 0) {
        $preMaxItemsRecipientCount = @($recipients).Count
        $recipients = @($recipients | Sort-Object PrimarySmtpAddress,Name | Select-Object -First $MaxItems)
        WriteLog -Message ("MaxItems enabled: restricted recipients from {0} to {1}." -f $preMaxItemsRecipientCount, @($recipients).Count) -Level 'WARNING'
    }

    $total   = $recipients.Count
    $counter = 0
    Write-Host "User/Shared mailboxes retrieved: $total"

    foreach ($rec in $recipients) {
        $counter++
        Write-Progress -Activity "Checking EmailAddresses..." -Status "$counter / $total" -PercentComplete (($counter / $total) * 100)

        $primary = $null
        try { $primary = $rec.PrimarySmtpAddress } catch { $primary = $null }

        $email           = ""
        $expectedAddress = ""
        $status          = ""
        $allowListed     = $false
        $policyWarning   = $false

        # Track Email Address Policy status without flooding the console.
        if ($rec.EmailAddressPolicyEnabled -eq $true) {
            $policyWarning = $true
            $policyEnabledCount++
        }

        if ($primary -and $primary.ToString() -match '.+@.+') {
            $email = $primary.ToString()
            $localPart = $email.Split('@')[0]
            $expectedAddress = "smtp:{0}@{1}" -f $localPart, $ExpectedSuffix

            $addresses = @($rec.EmailAddresses) | ForEach-Object { $_.ToString() }
            $exists    = $addresses | Where-Object { $_ -ieq $expectedAddress }

            if ($script:AllowListLoaded -and -not [string]::IsNullOrWhiteSpace($email)) {
                $allowListed = $script:AllowSet.Contains($email)
            }

            if ($exists) {
                $status = "OK"
                $okCount++
            } else {
                # missing before remediation
                $preMissing++
                $missingCount++

                # Decide if we may add
                $mayAdd = $AddMissingAddress -and ( $SkipAllowListCsv -or ($script:AllowListLoaded -and $allowListed) )

                if ($mayAdd) {
                    if (-not $SkipAllowListCsv -and $allowListed) { $missingInAllowCount++ }
                    if ($SkipAllowListCsv) { $status = "Missing -> WillAdd (SkipAllowList)" } else { $status = "Missing -> WillAdd" }

                    try {
                        $params = @{
                            Identity       = $rec.Identity
                            EmailAddresses = @{ add = $expectedAddress }
                            ErrorAction    = 'Stop'
                        }

                        if ($PSCmdlet.ShouldProcess($rec.Identity, "Add proxy address $expectedAddress")) {
                            Set-Mailbox @params
                            Write-Host "apply $expectedAddress for $email"
                            $addedCount++
                            $status = "Added"

                            $addedOperations.Add([PSCustomObject]@{
                                Identity      = $rec.Identity
                                DisplayName   = $rec.DisplayName
                                RecipientType = $rec.RecipientTypeDetails
                                AddedProxy    = $expectedAddress
                                PrimarySmtp   = $email
                                When          = (Get-Date)
                            })
                        } else {
                            $status = "Missing -> Skipped by WhatIf"
                        }
                    } catch {
                        $addFailedCount++
                        $status = "Missing -> AddFailed"
                        Write-Warning "[$($rec.Identity)] Failed to add $expectedAddress : $($_.Exception.Message)`n$($_.ScriptStackTrace)"
                    }
                } else {
                    # No add attempt
                    if ($script:AllowListLoaded) {
                        if ($allowListed) {
                            $missingInAllowCount++
                            $status = "Missing (InAllowList)"
                        } else {
                            $missingNotInAllowCount++
                            $status = "Missing (NotInAllowList)"
                        }
                    } else {
                        # No allowlist in use, count as not-in-allowlist for reporting symmetry
                        $missingNotInAllowCount++
                        $status = "Missing (NotInAllowList)"
                    }
                }

                # post-remediation still missing?
                if ($status -like 'Missing*' -or $status -eq 'Missing -> AddFailed' -or $status -eq 'Missing -> Skipped by WhatIf') {
                    $postMissing++
                }
            }
        }
        else {
            $status = "No primary address"
            $noPrimaryCount++
        }

        $results.Add([PSCustomObject]@{
            Identity                  = $rec.Identity
            DisplayName               = $rec.DisplayName
            RecipientType             = $rec.RecipientTypeDetails
            PrimaryAddress            = $email
            ExpectedAddress           = $expectedAddress
            Status                    = $status
            AllowListMatch            = $allowListed
            EmailAddressPolicyEnabled = $rec.EmailAddressPolicyEnabled
            PolicyWarning             = $policyWarning
        })
    }

    if ($policyEnabledCount -gt 0) {
        Write-Host "Email address policy enabled recipients: $policyEnabledCount. Details are available in the detail CSV."
    }

    $publishResults = @()
    $detailPublish = Publish-SmartM365Csv -Data @($results | Sort-Object Status, Identity) -TimestampedPath $outDetail -LatestPath $latestDetail
    if ($detailPublish) { $publishResults += $detailPublish }
    if ($addedOperations.Count -gt 0) {
        $addedPublish = Publish-SmartM365Csv -Data @($addedOperations) -TimestampedPath $outAdded -LatestPath $latestAdded
        if ($addedPublish) { $publishResults += $addedPublish }
    }

    # Dedicated CSV for "missing but in allowlist" (exclude Added)
    $allowMissing = $results | Where-Object {
        $_.AllowListMatch -and $_.Status -like 'Missing*' -and $_.Status -notlike '*NotInAllowList*' -and $_.Status -ne 'Added'
    }
    if ($allowMissing.Count -gt 0) {
        $allowMissingPublish = Publish-SmartM365Csv -Data @($allowMissing | Select-Object Identity, DisplayName, PrimaryAddress, ExpectedAddress, Status, EmailAddressPolicyEnabled, PolicyWarning) -TimestampedPath $outAllowMissing -LatestPath $latestAllowMissing
        if ($allowMissingPublish) { $publishResults += $allowMissingPublish }
    }

    $summary = @(
        [PSCustomObject]@{ Summary = "Total recipients processed";            Count = $results.Count },
        [PSCustomObject]@{ Summary = "With expected address present";         Count = $okCount },
        [PSCustomObject]@{ Summary = "With expected address missing";         Count = $missingCount },
        [PSCustomObject]@{ Summary = "Pre-remediation missing";               Count = $preMissing },
        [PSCustomObject]@{ Summary = "Post-remediation still missing";        Count = $postMissing },
        [PSCustomObject]@{ Summary = "With missing but in allowlist";         Count = $missingInAllowCount },
        [PSCustomObject]@{ Summary = "With missing but NOT in allowlist";     Count = $missingNotInAllowCount },
        [PSCustomObject]@{ Summary = "With no primary address";               Count = $noPrimaryCount },
        [PSCustomObject]@{ Summary = "With email address policy enabled";     Count = $policyEnabledCount },
        [PSCustomObject]@{ Summary = "Addresses successfully added";          Count = $addedCount },
        [PSCustomObject]@{ Summary = "Address additions failed";              Count = $addFailedCount },
        [PSCustomObject]@{ Summary = "Allowlist entries loaded";              Count = $script:AllowListRowCount }
    )
    $summaryPublish = Publish-SmartM365Csv -Data @($summary) -TimestampedPath $outSummary -LatestPath $latestSummary
    if ($summaryPublish) { $publishResults += $summaryPublish }

    Write-Host "`n===== Summary ====="
    foreach ($item in $summary) {
        switch ($item.Summary) {
            "With expected address missing"       { Write-Host "$($item.Summary): $($item.Count)" -ForegroundColor Yellow }
            "With expected address present"       { Write-Host "$($item.Summary): $($item.Count)" -ForegroundColor Green }
            "With missing but in allowlist"       { Write-Host "$($item.Summary): $($item.Count)" -ForegroundColor Cyan }
            "With missing but NOT in allowlist"   { Write-Host "$($item.Summary): $($item.Count)" -ForegroundColor Magenta }
            default                               { Write-Host "$($item.Summary): $($item.Count)" }
        }
    }

    Write-Host "Detail: $outDetail"
    if (Test-Path $outAdded)        { Write-Host "Added: $outAdded" }
    if (Test-Path $outAllowMissing) { Write-Host "Missing (in allowlist) : $outAllowMissing" }
    Write-Host "Summary: $outSummary"

# ===========================
# === Email notification ====
# ===========================
try {
    function Encode([string]$s) { return (ConvertTo-SmartM365EmailHtmlText $s) }

    $MailTo = @($To) -split '[;,]\s*' | Where-Object { $_ -and $_.Trim() -ne '' }
    $MailCc = @($Cc) -split '[;,]\s*' | Where-Object { $_ -and $_.Trim() -ne '' }

    $MailFrom    = $From
    $MailSubject = $Subject

    $attachments = @()
    if (Test-Path $outDetail)        { $attachments += $outDetail }
    if (Test-Path $outSummary)       { $attachments += $outSummary }
    if (Test-Path $outAdded)         { $attachments += $outAdded }
    if (Test-Path $outAllowMissing)  { $attachments += $outAllowMissing }

    if ([string]::IsNullOrWhiteSpace($MailFrom) -or -not $MailTo) {
        Write-Host "Email skipped: incomplete email parameters (From/To)."
    }
    else {
        $totalCount = [int](($summary | Where-Object { $_.Summary -eq 'Total recipients processed' } | Select-Object -First 1).Count)
        $presentCount = [int](($summary | Where-Object { $_.Summary -eq 'With expected address present' } | Select-Object -First 1).Count)
        $missingCount = [int](($summary | Where-Object { $_.Summary -eq 'With expected address missing' } | Select-Object -First 1).Count)
        $allowMissingCount = [int](($summary | Where-Object { $_.Summary -eq 'With missing but in allowlist' } | Select-Object -First 1).Count)
        $notAllowMissingCount = [int](($summary | Where-Object { $_.Summary -eq 'With missing but NOT in allowlist' } | Select-Object -First 1).Count)
        $addedCount = [int](($summary | Where-Object { $_.Summary -eq 'Addresses successfully added' } | Select-Object -First 1).Count)
        $policyEnabledCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'With email address policy enabled' } | Select-Object -First 1).Count)
        $effectiveSendMailMode = if ([string]::IsNullOrWhiteSpace($SendMailMode)) { if ([string]::IsNullOrWhiteSpace($SmtpServer)) { 'Graph' } else { 'SMTP' } } else { $SendMailMode.Trim() }

        $summaryRowsForEmail = @(
            [pscustomobject]@{ Label = 'Total recipients processed'; Value = $totalCount }
            [pscustomobject]@{ Label = 'With expected address present'; Value = $presentCount }
            [pscustomobject]@{ Label = 'With expected address missing'; Value = $missingCount }
            [pscustomobject]@{ Label = 'Missing but in allowlist'; Value = $allowMissingCount }
            [pscustomobject]@{ Label = 'Missing but not in allowlist'; Value = $notAllowMissingCount }
            [pscustomobject]@{ Label = 'Email address policy enabled'; Value = $policyEnabledCountForMail }
            [pscustomobject]@{ Label = 'Addresses added'; Value = $addedCount }
        )

        $pathRows = @(
            [pscustomobject]@{ Label = 'Detail CSV'; Path = $outDetail }
            [pscustomobject]@{ Label = 'Summary CSV'; Path = $outSummary }
        )
        if (Test-Path $outAdded) { $pathRows += [pscustomobject]@{ Label = 'Added CSV'; Path = $outAdded } }
        if (Test-Path $outAllowMissing) { $pathRows += [pscustomobject]@{ Label = 'Allowlist missing CSV'; Path = $outAllowMissing } }

        $scopeHtml = if ($AllOrganizationalUnit) { 'ALL (entire forest)' } else { ($OrganizationalUnit | ForEach-Object { Encode $_ }) -join '<br/>' }
        $modeLabel = if ($AddMissingAddress) { 'Write mode' } else { 'Read-only mode' }
        $allowListMode = if ($SkipAllowListCsv) { 'Allowlist skipped' } else { 'Allowlist enforced' }
        $scopeSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr><td style="width:180px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Mode</td><td style="border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;">$(Encode $modeLabel)</td></tr>
  <tr><td style="width:180px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Allowlist</td><td style="border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;">$(Encode $allowListMode)</td></tr>
  <tr><td style="width:180px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Expected suffix</td><td style="border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;color:#334155;">$(Encode $ExpectedSuffix)</td></tr>
  <tr><td style="width:180px;background:#f8fafc;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Scope</td><td style="padding:10px 12px;font-size:13px;color:#334155;word-break:break-all;">$scopeHtml</td></tr>
</table>
"@

        $sections = @([pscustomobject]@{ Title = 'Scope'; Html = $scopeSectionHtml })

        $missingPreviewRows = @($results | Where-Object { $_.Status -like 'Missing*' } | Sort-Object Status, DisplayName | Select-Object -First 50)
        if ($missingPreviewRows.Count -gt 0) {
            $missingRowsHtml = foreach ($row in $missingPreviewRows) {
                "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(Encode $row.DisplayName)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.PrimaryAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.ExpectedAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;font-weight:700;color:#92400e;`">$(Encode $row.Status)</td></tr>"
            }
            $missingSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Display name</th>
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
        $actionHtml = if ($missingCount -gt 0) { 'Review missing proxy addresses before remediation. Use write mode only after validating the scope and allowlist decision.' } else { '' }
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
            -Footer 'This automated message was generated by SmartM365. Use the exported CSV files and SharePoint links as the source of truth.'

        SendEmailHtmlReport -SendMailMode $effectiveSendMailMode -SmtpServer $SmtpServer -SmtpPort $SmtpPort -From $MailFrom -To ($MailTo -join ';') -Cc ($MailCc -join ';') -Subject $MailSubject -BodyHtml $body -Attachments $attachments
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
	# Clean up old CSV files + old log files
	# Automatically excludes all generated CSVs via global:csvGeneratedPaths + current transcript and log files via global variables
	RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDoePA6UFCEP1A/
# nwLnNnjJMN/76gukgUipIuvqAz+9d6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCDpqR5i7s+7LaqOy2/3Hl6/exa6bdYOe431NqgdlFSbNDANBgkqhkiG9w0B
# AQEFAASCAYAaVqBCWGeGNz64R8j06F4GXujT1eOC52t9o1PYsyU7xbrPndm3utcr
# M1ci0SddgpuJJB3xUAmlpW9SwtSbW5bhjSvfYJeB8dP6VF58itk5Qv+5PCzTtGbh
# RvP2/sZf5APSOYpnqGX7KOhlpMA6NLLvMULljbyH93geehWyugdhIa8dLHItpojk
# UBthWq6dTYllzu+fsH4E8GvMcCKYFKM52tnP74IAjphHky4RQjNvETGjf8dam10o
# xAIarvBecsnG6erF0DAoCs6G+z2+u+dENbzroI/lN4qbfrQKh+5gV03OBGeVyt3c
# ZE8BkSoRXeL3QeuiIhv4dtnOfzyyODZdcEvyDa5FoL2UYHDRwNRWwzGIav5quR04
# drzIIw7kejUK5ewx4RekoY10KVgQSNd3BW2kgf2Rg6ht8dTMMR/nmwoP1GMIR5XK
# aCmqmPTDa9l7jFHUGR0mTfO869jV1VRHJwWDUwm3nyc+6w2vTsN9gHmz0AkMx8Cg
# ONzMA5QXFaM=
# SIG # End signature block
