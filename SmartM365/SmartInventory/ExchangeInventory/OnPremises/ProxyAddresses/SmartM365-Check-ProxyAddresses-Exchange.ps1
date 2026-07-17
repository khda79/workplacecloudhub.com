<#
.SYNOPSIS
    Audits and optionally remediates proxy addresses for on-premises Exchange **User & Shared** mailboxes in specified Organizational Units (OUs), or all OUs.

.DESCRIPTION
    This script checks whether each **UserMailbox** and **SharedMailbox** has a proxy address with the expected domain suffix (e.g., tenant.mail.onmicrosoft.com).
    The expected proxy address local part is derived from the recipient SamAccountName.
    If the expected address is missing and the -AddMissingAddress switch is specified, the script can add the address.
    The script can target multiple OUs via -OrganizationalUnit (string array) or all OUs with -AllOrganizationalUnit.
    It generates detailed, summary, and remediation CSV reports and maintains logs.
    Designed to run on an Exchange 2016 server with the Management Tools installed.

.PARAMETER OrganizationalUnit
    Distinguished names of the target OU(s). Accepts multiple OUs.
    No default scope is assumed. Configure at least one OU or use -AllOrganizationalUnit.

.PARAMETER AllOrganizationalUnit
    If specified, the script ignores OU filtering and fetches all User/Shared mailboxes in the forest.

.PARAMETER ExpectedSuffix
    The expected domain suffix for proxy addresses (default: "tenant.mail.onmicrosoft.com").

.PARAMETER AddMissingAddress
    If specified, missing proxy addresses will be added.

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
1.13

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
    $ScriptVersion = "1.13"
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

    $timestamp       = Get-Date -Format 'yyyyMMdd-HHmm'
$outDetail       = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Check_{0}.csv" -f $timestamp)
$outSummary      = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Summary_{0}.csv" -f $timestamp)
$outAdded        = Join-Path -Path $OutputPath -ChildPath ("Exchange_OnPrem_ProxyAddresses_Added_{0}.csv" -f $timestamp)
$latestDetail       = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Check.csv' } else { $null }
$latestSummary      = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Summary.csv' } else { $null }
$latestAdded        = if ($LatestCsvFolderPath) { Join-Path -Path $LatestCsvFolderPath -ChildPath 'Exchange_OnPrem_ProxyAddresses_Added.csv' } else { $null }

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

    # Collections & counters
    $results                = New-Object System.Collections.Generic.List[object]
    $addedOperations        = New-Object System.Collections.Generic.List[object]

    $okCount                = 0
    $missingCount           = 0
    $noPrimaryCount         = 0
    $noSamAccountNameCount  = 0
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
        $samAccountName  = ([string]$rec.SamAccountName).Trim()
        $expectedAddress = ""
        $status          = ""
        $policyWarning   = $false

        # Track Email Address Policy status without flooding the console.
        if ($rec.EmailAddressPolicyEnabled -eq $true) {
            $policyWarning = $true
            $policyEnabledCount++
        }

        if ($primary -and $primary.ToString() -match '.+@.+') {
            $email = $primary.ToString()

            if ([string]::IsNullOrWhiteSpace($samAccountName)) {
                $status = "No SamAccountName"
                $noSamAccountNameCount++
            }
            else {
                $expectedAddress = "smtp:{0}@{1}" -f $samAccountName, $ExpectedSuffix
                $addresses = @($rec.EmailAddresses) | ForEach-Object { $_.ToString() }
                $exists = $addresses | Where-Object { $_ -ieq $expectedAddress }

                if ($exists) {
                    $status = "OK"
                    $okCount++
                }
                else {
                    $preMissing++
                    $missingCount++

                    if ($AddMissingAddress) {
                        $status = "Missing -> WillAdd"
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
                                    Identity       = $rec.Identity
                                    SamAccountName = $samAccountName
                                    DisplayName    = $rec.DisplayName
                                    RecipientType  = $rec.RecipientTypeDetails
                                    AddedProxy     = $expectedAddress
                                    PrimarySmtp    = $email
                                    When           = (Get-Date)
                                })
                            }
                            else {
                                $status = "Missing -> Skipped by WhatIf"
                            }
                        }
                        catch {
                            $addFailedCount++
                            $status = "Missing -> AddFailed"
                            Write-Warning "[$($rec.Identity)] Failed to add $expectedAddress : $($_.Exception.Message)`n$($_.ScriptStackTrace)"
                        }
                    }
                    else {
                        $status = "Missing"
                    }

                    if ($status -like 'Missing*') {
                        $postMissing++
                    }
                }
            }
        }
        else {
            $status = "No primary address"
            $noPrimaryCount++
        }

        $results.Add([PSCustomObject]@{
            Identity                  = $rec.Identity
            SamAccountName            = $samAccountName
            DisplayName               = $rec.DisplayName
            RecipientType             = $rec.RecipientTypeDetails
            PrimaryAddress            = $email
            ExpectedAddress           = $expectedAddress
            Status                    = $status
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

    $summary = @(
        [PSCustomObject]@{ Summary = "Total recipients processed";            Count = $results.Count },
        [PSCustomObject]@{ Summary = "With expected address present";         Count = $okCount },
        [PSCustomObject]@{ Summary = "With expected address missing";         Count = $missingCount },
        [PSCustomObject]@{ Summary = "Pre-remediation missing";               Count = $preMissing },
        [PSCustomObject]@{ Summary = "Post-remediation still missing";        Count = $postMissing },
        [PSCustomObject]@{ Summary = "With no primary address";               Count = $noPrimaryCount },
        [PSCustomObject]@{ Summary = "With no SamAccountName";                Count = $noSamAccountNameCount },
        [PSCustomObject]@{ Summary = "With email address policy enabled";     Count = $policyEnabledCount },
        [PSCustomObject]@{ Summary = "Addresses successfully added";          Count = $addedCount },
        [PSCustomObject]@{ Summary = "Address additions failed";              Count = $addFailedCount }
    )
    $summaryPublish = Publish-SmartM365Csv -Data @($summary) -TimestampedPath $outSummary -LatestPath $latestSummary
    if ($summaryPublish) { $publishResults += $summaryPublish }

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

    if ([string]::IsNullOrWhiteSpace($MailFrom) -or -not $MailTo) {
        Write-Host "Email skipped: incomplete email parameters (From/To)."
    }
    else {
        $totalCount = [int](($summary | Where-Object { $_.Summary -eq 'Total recipients processed' } | Select-Object -First 1).Count)
        $presentCount = [int](($summary | Where-Object { $_.Summary -eq 'With expected address present' } | Select-Object -First 1).Count)
        $missingCount = [int](($summary | Where-Object { $_.Summary -eq 'With expected address missing' } | Select-Object -First 1).Count)
        $noSamAccountNameCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'With no SamAccountName' } | Select-Object -First 1).Count)
        $addedCount = [int](($summary | Where-Object { $_.Summary -eq 'Addresses successfully added' } | Select-Object -First 1).Count)
        $policyEnabledCountForMail = [int](($summary | Where-Object { $_.Summary -eq 'With email address policy enabled' } | Select-Object -First 1).Count)
        $effectiveSendMailMode = if ([string]::IsNullOrWhiteSpace($SendMailMode)) { if ([string]::IsNullOrWhiteSpace($SmtpServer)) { 'Graph' } else { 'SMTP' } } else { $SendMailMode.Trim() }

        $summaryRowsForEmail = @(
            [pscustomobject]@{ Label = 'Total recipients processed'; Value = $totalCount }
            [pscustomobject]@{ Label = 'With expected address present'; Value = $presentCount }
            [pscustomobject]@{ Label = 'With expected address missing'; Value = $missingCount }
            [pscustomobject]@{ Label = 'With no SamAccountName'; Value = $noSamAccountNameCountForMail }
            [pscustomobject]@{ Label = 'Email address policy enabled'; Value = $policyEnabledCountForMail }
            [pscustomobject]@{ Label = 'Addresses added'; Value = $addedCount }
        )

        $pathRows = @(
            [pscustomobject]@{ Label = 'Detail CSV'; Path = $outDetail }
            [pscustomobject]@{ Label = 'Summary CSV'; Path = $outSummary }
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

        $missingPreviewRows = @($results | Where-Object { $_.Status -like 'Missing*' } | Sort-Object Status, DisplayName | Select-Object -First 50)
        if ($missingPreviewRows.Count -gt 0) {
            $missingRowsHtml = foreach ($row in $missingPreviewRows) {
                "<tr><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(Encode $row.DisplayName)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;`">$(Encode $row.SamAccountName)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.PrimaryAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;color:#334155;word-break:break-all;`">$(Encode $row.ExpectedAddress)</td><td style=`"border-bottom:1px solid #eef2f7;padding:9px 10px;font-size:12px;font-weight:700;color:#92400e;`">$(Encode $row.Status)</td></tr>"
            }
            $missingSectionHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
  <tr>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">Display name</th>
    <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px;font-size:12px;color:#475569;text-transform:uppercase;">SamAccountName</th>
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
        $actionHtml = if ($missingCount -gt 0) { 'Review missing proxy addresses before remediation. Use write mode only after validating the scope.' } else { '' }
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDrP5ddFQZEixSX
# lic0sb23fybrCxaEPW7GRv2R+kZhVaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAwOItZOXg8DR4hhy9RCoVZ2KfHLNWaYAF56kFCak0EgMA0GCSqG
# SIb3DQEBAQUABIIBgAB36+d4B9pzkTUo22hmwagU1Z7frhdgMxPSRajotr/6bxZh
# gzB/ygg7uvEkMrjMUdcYF/dEmadbp9aoUuox4vs0Jgu903JujyuOaPyg5luZdJaS
# EGkCzUpyDohZy+FCOjbPuBa0RFkTL+J0ieuOe/3bQYuhuONCnqWWhM7E2UDPJigu
# G0kzaTlUL4p6AdUkXUip3ti8s0UHMdKqLZp7IfbQsOmaXcNdktuv+Qs2NwRPD+TH
# LV74oNHmqe2EivA9HtteDPjm8QAmwvysid1+mLpoON2ctP4U5RC0Z1tpplvXFeWO
# NFgVAQhIRfAYLdDIidLpd6xcXsuXSVF1XoQbRrmA5k/Skg8VAnXphABJAhrWbyQS
# NVpqEKJ9PKEokmpNTWYGsVmeAl7XLgy8+PDMYjsBXnehjDlQvL6e4QD02DsFoh3/
# cZaTJS10bVUltd8Vhf5k2XBt/Ug/v//vJSUnvbONzgKFCBE2Ejdk0I0FAZaY4mxa
# fmxq1tX4U/W/inCDGaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTcxMDI1
# MTZaMC8GCSqGSIb3DQEJBDEiBCDTEWGdk+qXg2Qzk0rkvdxLEgxP4YeeJEvayntT
# 6ESZ3TANBgkqhkiG9w0BAQEFAASCAgCtfwuzenj0So5DzFewfjKb9biVaGezelAA
# tlXVfV4/4LAqOMC/9PQStLQpVLVe+tnuydqlqZfypUOS2rPdUiBLfClW5yDcIQSX
# +SM2L0eVB1LJhqGDHYsd2EEaUw2NIVK9rsN/xZ1MMbBlaoy9T2UutsLYtAONEjDN
# /O9JKgS26HSoc6Mw4Reef6E/4/qZurDzSwnJY2VcuMO3dWQJ5BEeF9U+BqVbGhKo
# v6T2BIJv+ZR5HqFjT7ebPYEF/ywBrtmEr0o5PvExfNvU0dK1yFyYTnugGhAtFRtA
# D5OBR54YbkjpgE46eV6oRnmgKVH8fKdgL7tRZqQ9lv38VmrogsY54AExUGVCj3Jb
# ANSwt7bYtDTD3bnDRgsGGFfonrEIatan3hC+dW1LFYY483sJucZI++v3qWDdSZQR
# kLp/7OPz9r2Fa8bL4pmUVaNse3L3/uFj8MyrXkDpf8F9OswtqRkSXqP9KosBB3OW
# Zs8eGC9YFTexQEI/lvPrbrrijVWYvb/mBusIAb9ZdPA51SryECdubRgqr60IL4M4
# ZaBBmxIkM/2oc3ZTkdRxQ7xA4wW+YVM0nXpNNKMP4D9aTK3t2i72a/l9XWSQI4nb
# eS4W4rWCbcdRN0jE0H/mgdEgy/Te4kGOy6EPu06iD68sPpqfRpfkluC1J7Dz/m7s
# gNbRAdj26w==
# SIG # End signature block
