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

.NOTES
    - Requires Exchange 2016 Management Tools.
    - Generates detailed, summary, and added addresses CSV reports.
    - Maintains logs and cleans up old files automatically.

.VERSION
1.8

.AUTHOR
    https://github.com/khda79/workplacecloudhub.com
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
    [string]$Subject = "Check ProxyAddresses Report"
)
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
            Write-Host 'Edit the local JSON now if needed. Press Enter to continue with the current file values.' -ForegroundColor Yellow
            Read-Host 'Press Enter to continue'
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
    $ScriptVersion = "1.8"
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
        Import-Module (Join-ModulePath 'SmartM365-WindowsPowerShell5.psd1') -ErrorAction Stop
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

    # Recipient retrieval: either all OUs (no filter) or each OU provided
    if ($AllOrganizationalUnit) {
        try {
            Write-Host "Fetching recipients from ALL Organizational Units..."
            $recipients = Get-Recipient -ResultSize Unlimited `
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
                                      -ResultSize Unlimited `
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

    Publish-SmartM365Csv -Data @($results | Sort-Object Status, Identity) -TimestampedPath $outDetail -LatestPath $latestDetail | Out-Null
    if ($addedOperations.Count -gt 0) {
        Publish-SmartM365Csv -Data @($addedOperations) -TimestampedPath $outAdded -LatestPath $latestAdded | Out-Null
    }

    # Dedicated CSV for "missing but in allowlist" (exclude Added)
    $allowMissing = $results | Where-Object {
        $_.AllowListMatch -and $_.Status -like 'Missing*' -and $_.Status -notlike '*NotInAllowList*' -and $_.Status -ne 'Added'
    }
    if ($allowMissing.Count -gt 0) {
        Publish-SmartM365Csv -Data @($allowMissing | Select-Object Identity, DisplayName, PrimaryAddress, ExpectedAddress, Status, EmailAddressPolicyEnabled, PolicyWarning) -TimestampedPath $outAllowMissing -LatestPath $latestAllowMissing | Out-Null
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
    Publish-SmartM365Csv -Data @($summary) -TimestampedPath $outSummary -LatestPath $latestSummary | Out-Null

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
    # HTML encoding helper
    Add-Type -AssemblyName System.Web | Out-Null
    function Encode([string]$s) { return [System.Web.HttpUtility]::HtmlEncode($s) }

    # Recipients (support ; or ,)
    $MailTo = @($To) -split '[;,]\s*' | Where-Object { $_ -and $_.Trim() -ne '' }
    $MailCc = @($Cc) -split '[;,]\s*' | Where-Object { $_ -and $_.Trim() -ne '' }

    $MailFrom    = $From
    $MailSubject = $Subject

    # Attach CSV reports
    $attachments = @()
    if (Test-Path $outDetail)        { $attachments += $outDetail }
    if (Test-Path $outSummary)       { $attachments += $outSummary }
    if (Test-Path $outAdded)         { $attachments += $outAdded }
    if (Test-Path $outAllowMissing)  { $attachments += $outAllowMissing }

    if ([string]::IsNullOrWhiteSpace($MailFrom) -or -not $MailTo) {
        Write-Host "Email skipped: incomplete email parameters (From/To)."
    }
    else {
        $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $totalCount = [int](($summary | Where-Object { $_.Summary -eq 'Total recipients processed' } | Select-Object -First 1).Count)
        $presentCount = [int](($summary | Where-Object { $_.Summary -eq 'With expected address present' } | Select-Object -First 1).Count)
        $missingCount = [int](($summary | Where-Object { $_.Summary -eq 'With expected address missing' } | Select-Object -First 1).Count)
        $allowMissingCount = [int](($summary | Where-Object { $_.Summary -eq 'With missing but in allowlist' } | Select-Object -First 1).Count)
        $notAllowMissingCount = [int](($summary | Where-Object { $_.Summary -eq 'With missing but NOT in allowlist' } | Select-Object -First 1).Count)
        $addedCount = [int](($summary | Where-Object { $_.Summary -eq 'Addresses successfully added' } | Select-Object -First 1).Count)
        $modeLabel = if ($AddMissingAddress) { 'Write mode' } else { 'Read-only mode' }
        $modeColor = if ($AddMissingAddress) { '#b45309' } else { '#047857' }
        $allowListMode = if ($SkipAllowListCsv) { 'Allowlist skipped' } else { 'Allowlist enforced' }
        $effectiveSendMailMode = if ([string]::IsNullOrWhiteSpace($SendMailMode)) { if ([string]::IsNullOrWhiteSpace($SmtpServer)) { 'Graph' } else { 'SMTP' } } else { $SendMailMode.Trim() }
        $transportLabel = switch ($effectiveSendMailMode) { 'Graph' { 'Microsoft Graph' } 'SMTP' { "SMTP $SmtpServer`:$SmtpPort" } 'Both' { "Microsoft Graph with SMTP fallback $SmtpServer`:$SmtpPort" } default { $effectiveSendMailMode } }

        # Build summary table rows
        $rowsSummary = foreach ($row in $summary) {
            "<tr><td style='padding:10px 12px;border-bottom:1px solid #e5e7eb;color:#334155;'>$(Encode $row.Summary)</td><td style='padding:10px 12px;border-bottom:1px solid #e5e7eb;text-align:right;font-weight:700;color:#0f172a;'>$([string]$row.Count)</td></tr>"
        }

        # Build HTML section listing "missing but in allowlist"
        if (-not $allowMissing) {
            $allowMissing = $results | Where-Object {
                $_.AllowListMatch -and $_.Status -like 'Missing*' -and $_.Status -notlike '*NotInAllowList*' -and $_.Status -ne 'Added'
            }
        }

        $rowsAllow = foreach ($row in ($allowMissing | Sort-Object DisplayName)) {
            "<tr><td style='padding:9px 10px;border-bottom:1px solid #e5e7eb;color:#0f172a;'>$(Encode $row.DisplayName)</td><td style='padding:9px 10px;border-bottom:1px solid #e5e7eb;color:#475569;'>$(Encode $row.PrimaryAddress)</td><td style='padding:9px 10px;border-bottom:1px solid #e5e7eb;color:#475569;'>$(Encode $row.ExpectedAddress)</td><td style='padding:9px 10px;border-bottom:1px solid #e5e7eb;color:#b45309;font-weight:600;'>$(Encode $row.Status)</td></tr>"
        }

        $allowSection = if (@($allowMissing).Count -gt 0) {
@"
<div style="margin-top:22px;">
  <h3 style="margin:0 0 10px 0;font-size:16px;color:#0f172a;">Allowlisted missing addresses</h3>
  <table role="presentation" style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
    <thead>
      <tr style="background:#f8fafc;">
        <th style="padding:10px;text-align:left;color:#475569;font-size:12px;text-transform:uppercase;">Display name</th>
        <th style="padding:10px;text-align:left;color:#475569;font-size:12px;text-transform:uppercase;">Primary SMTP</th>
        <th style="padding:10px;text-align:left;color:#475569;font-size:12px;text-transform:uppercase;">Expected proxy</th>
        <th style="padding:10px;text-align:left;color:#475569;font-size:12px;text-transform:uppercase;">Status</th>
      </tr>
    </thead>
    <tbody>
      $($rowsAllow -join "`n")
    </tbody>
  </table>
</div>
"@
        } else {
            '<div style="margin-top:22px;padding:14px 16px;border:1px solid #bbf7d0;background:#f0fdf4;border-radius:8px;color:#166534;font-weight:600;">No allowlisted mailbox is missing the expected address.</div>'
        }

        # OU section text
        $ouHtml = if ($AllOrganizationalUnit) { "ALL (entire forest)" } else { ($OrganizationalUnit | ForEach-Object { Encode $_ }) -join "<br/>" }
        $attachmentNames = @($attachments | ForEach-Object { "<code style='background:#f1f5f9;border:1px solid #e2e8f0;border-radius:4px;padding:2px 5px;color:#334155;'>$(Encode (Split-Path -Leaf $_))</code>" }) -join ' '

        $body = @"
<html>
  <body style="margin:0;padding:0;background:#f6f8fb;font-family:Segoe UI,Arial,sans-serif;color:#0f172a;">
    <div style="max-width:980px;margin:0 auto;padding:24px;">
      <div style="background:#ffffff;border:1px solid #dbe3ef;border-radius:10px;overflow:hidden;">
        <div style="background:linear-gradient(135deg,#0f766e,#2563eb);padding:22px 26px;color:#ffffff;">
          <div style="font-size:12px;text-transform:uppercase;letter-spacing:.08em;opacity:.9;">Smart365 Exchange OnPrem</div>
          <h1 style="margin:6px 0 0 0;font-size:24px;line-height:1.25;">$(Encode $MailSubject)</h1>
          <div style="margin-top:10px;font-size:13px;opacity:.95;">Generated $now - Transport: $(Encode $transportLabel)</div>
        </div>

        <div style="padding:22px 26px;">
          <div style="margin-bottom:18px;">
            <span style="display:inline-block;background:$modeColor;color:#ffffff;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:700;">$modeLabel</span>
            <span style="display:inline-block;background:#eef2ff;color:#3730a3;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:700;margin-left:6px;">$allowListMode</span>
          </div>

          <table role="presentation" style="border-collapse:collapse;width:100%;margin-bottom:20px;">
            <tr>
              <td style="width:20%;padding:0 8px 8px 0;"><div style="border:1px solid #dbe3ef;border-radius:8px;padding:14px;background:#f8fafc;"><div style="font-size:11px;text-transform:uppercase;color:#64748b;font-weight:700;">Total checked</div><div style="font-size:26px;font-weight:800;color:#0f172a;margin-top:4px;">$totalCount</div></div></td>
              <td style="width:20%;padding:0 8px 8px 0;"><div style="border:1px solid #bbf7d0;border-radius:8px;padding:14px;background:#f0fdf4;"><div style="font-size:11px;text-transform:uppercase;color:#166534;font-weight:700;">Present</div><div style="font-size:26px;font-weight:800;color:#166534;margin-top:4px;">$presentCount</div></div></td>
              <td style="width:20%;padding:0 8px 8px 0;"><div style="border:1px solid #fed7aa;border-radius:8px;padding:14px;background:#fff7ed;"><div style="font-size:11px;text-transform:uppercase;color:#9a3412;font-weight:700;">Missing</div><div style="font-size:26px;font-weight:800;color:#9a3412;margin-top:4px;">$missingCount</div></div></td>
              <td style="width:20%;padding:0 8px 8px 0;"><div style="border:1px solid #bfdbfe;border-radius:8px;padding:14px;background:#eff6ff;"><div style="font-size:11px;text-transform:uppercase;color:#1d4ed8;font-weight:700;">Allowlisted</div><div style="font-size:26px;font-weight:800;color:#1d4ed8;margin-top:4px;">$allowMissingCount</div></div></td>
              <td style="width:20%;padding:0 0 8px 0;"><div style="border:1px solid #e9d5ff;border-radius:8px;padding:14px;background:#faf5ff;"><div style="font-size:11px;text-transform:uppercase;color:#7e22ce;font-weight:700;">Added</div><div style="font-size:26px;font-weight:800;color:#7e22ce;margin-top:4px;">$addedCount</div></div></td>
            </tr>
          </table>

          <div style="border:1px solid #e5e7eb;border-radius:8px;background:#ffffff;padding:14px 16px;margin-bottom:20px;">
            <div style="font-size:13px;color:#334155;line-height:1.55;">
              <strong>Scope:</strong> $ouHtml<br/>
              <strong>Expected suffix:</strong> $(Encode $ExpectedSuffix)<br/>
              <strong>Not in allowlist:</strong> $notAllowMissingCount
            </div>
          </div>

          <h3 style="margin:0 0 10px 0;font-size:16px;color:#0f172a;">Summary</h3>
          <table role="presentation" style="border-collapse:collapse;width:100%;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
            <thead>
              <tr style="background:#f8fafc;">
                <th style="padding:10px 12px;text-align:left;color:#475569;font-size:12px;text-transform:uppercase;">Metric</th>
                <th style="padding:10px 12px;text-align:right;color:#475569;font-size:12px;text-transform:uppercase;">Count</th>
              </tr>
            </thead>
            <tbody>
              $($rowsSummary -join "`n")
            </tbody>
          </table>

          $allowSection

          <div style="margin-top:22px;padding-top:14px;border-top:1px solid #e5e7eb;color:#64748b;font-size:12px;line-height:1.6;">
            Detailed CSV reports are attached.<br/>
            $attachmentNames
          </div>
        </div>
      </div>
    </div>
  </body>
</html>
"@

        SendEmailHtmlReport -SendMailMode $effectiveSendMailMode -SmtpServer $SmtpServer -SmtpPort $SmtpPort -From $MailFrom -To ($MailTo -join ';') -Cc ($MailCc -join ';') -Subject $MailSubject -BodyHtml $body -Attachments $attachments
        Write-Host "Email sent to '$($MailTo -join ';')' via $transportLabel."
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
    Complete-SmartM365ExecutionContext -Status Success
    Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
    #endregion
