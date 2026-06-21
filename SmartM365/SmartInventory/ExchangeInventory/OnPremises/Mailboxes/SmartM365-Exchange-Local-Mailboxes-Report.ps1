<#
.SYNOPSIS
Builds an on-premises Exchange mailbox report from local inventory outputs.

.DESCRIPTION
Reads local Exchange mailbox inventory exports and generates consolidated report files for review and follow-up.
Runtime-specific paths are loaded from the script local configuration file.

.PARAMETER TargetDomains
Optional list of domains to include in the report.

.PARAMETER OutputPath
Optional output folder override.

.PARAMETER FileFreshnessHours
Maximum age, in hours, accepted for source inventory files.

.NOTES
Author    : https://github.com/khda79/workplacecloudhub.com
Version   : 1.0
Requires  : Windows PowerShell 5.1 and Exchange Management Tools
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
[Parameter(Mandatory=$false)]
    [string[]]$TargetDomains,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 168)]
    [int]$FileFreshnessHours = 4
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
        return [pscustomobject]@{}
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
#region Module Import and Initialization
$ScriptVersion = "1.0"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxReportCsvLogFolderPath' -DefaultValue $OutputPath
function Join-ModulePath {
    param([Parameter(Mandatory)][string]$FileName)
    $repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..')).Path
    return (Join-Path (Join-Path (Join-Path (Join-Path $repoRoot 'Modules') 'SmartM365.Core') 'Compatibility\WindowsPowerShell5') $FileName)
}

$LimitResultSize = $null
if ($LimitResultSize) {
    $TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion (LimitResultSize : $LimitResultSize)..."
}

function EnsureExchangePSSnapinLoaded {
    [CmdletBinding()]
    param (
        [string]$SnapinName = "Microsoft.Exchange.Management.PowerShell.SnapIn"
    )

    # Check if the PSSnapin is registered on the server
    if (-not (Get-PSSnapin $SnapinName -Registered -ErrorAction SilentlyContinue)) {
        Write-Error "The Exchange Management PSSnapin '$SnapinName' is not registered on this server."
        Write-Error "This script must be run on an Exchange 2016 server where the Management Tools are installed."
        return $false
    }

    # Check if the PSSnapin is already loaded in the current session
    if (-not (Get-PSSnapin $SnapinName -ErrorAction SilentlyContinue)) {
        Write-Verbose "The Exchange PSSnapin '$SnapinName' is not loaded in the current session. Attempting to load it..."
        try {
            Add-PSSnapin $SnapinName -ErrorAction Stop
            Write-Verbose "The Exchange PSSnapin was loaded successfully."
        }
        catch {
            Write-Error "Failed to load the Exchange PSSnapin '$SnapinName'. Error: $($_.Exception.Message)"
            Write-Error "Ensure you are running this script on an Exchange 2016 server and have the necessary permissions."
            Write-Error "Alternatively, try running this script from the Exchange Management Shell or use a script that connects via PowerShell Remoting to localhost."
            return $false
        }
    } else {
        Write-Verbose "The Exchange PSSnapin '$SnapinName' is already loaded."
    }

    # Verify that the Get-Mailbox command is now available
    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-Error "The Get-Mailbox cmdlet is still not available after attempting to load the snap-in."
        Write-Error "This could indicate an issue with the Exchange Management Tools installation."
        return $false
    }

    return $true
}

try {
    Write-Host "Loading module SmartM365-WindowsPowerShell5.psd1..."
    Import-Module (Join-ModulePath 'SmartM365-WindowsPowerShell5.psd1') -ErrorAction Stop
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append
    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
    WriteLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
    WriteLog -Message "Parameter FileFreshnessHours: $FileFreshnessHours"
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    # If the module fails to load, we cannot send emails (functions live there); exiting is acceptable here.
    exit
}
#endregion

#region Daily Success Lock
function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        return [pscustomobject]@{}
    }

    try {
        return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
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
$LockRoot = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LockRoot' -DefaultValue "C:\ProgramData\SmartM365\Locks"
$LockScriptName = "Exchange-Local-Mailboxes"
$LockToday = Get-Date -Format "yyyy-MM-dd"
$DailySuccessLockFile = Join-Path $LockRoot "$LockScriptName-SUCCESS-$LockToday.lock"

if (!(Test-Path -LiteralPath $LockRoot)) {
    New-Item -Path $LockRoot -ItemType Directory -Force | Out-Null
}

if (Test-Path -LiteralPath $DailySuccessLockFile) {
    Write-Output "[$LockScriptName] Already succeeded on $LockToday. Skipping execution."
    try { WriteLog -Message "Daily success lock exists for $LockToday ($DailySuccessLockFile). Execution skipped." "WARN" } catch {}
    exit 0
}

$ScriptSucceeded = $false
#endregion
# Predeclare paths & timers & data for email
$StartTime = Get-Date
$csvOutputPath    = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_DailyStats.csv"
$summaryCsvPath   = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_DailyStats_Summary.csv"
$summaryWithTotal = $null

# Prepare secondary copy destination from local configuration (LatestCsvFolderPath)
$csvOutputPath2 = $null
try {
    $pathForSecondCopy = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
    if ($pathForSecondCopy) {
        $csvOutputPath2 = Join-Path -Path $pathForSecondCopy -ChildPath "Exchange_OnPrem_Mailboxes_DailyStats.csv"
        WriteLog -Message "Second CSV copy target resolved to $csvOutputPath2"
    } else {
        WriteLog -Message "LatestCsvFolderPath not found or empty. Secondary CSV copy disabled." "WARN"
    }
} catch {
    WriteLog -Message "Failed to resolve LatestCsvFolderPath for secondary CSV copy: $($_.Exception.Message)" "WARN"
}

# Resolve input CSV paths from local configuration (do NOT contact Exchange yet)
try {
    $InputRemoteMailboxes = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RemoteMailboxCsvLogFolderPath' -DefaultValue ''
$InputRemoteMailboxesFile = Join-Path $InputRemoteMailboxes "Exchange_OnPrem_RemoteMailboxes_AllDomains.csv"
} catch {
    Write-Error "Error retrieving local configuration path InputRemoteMailboxes $_"
    WriteLog -Message "ERROR retrieving local configuration path: $_" "ERROR"
    # Do not exit here; we will fallback to live Exchange if needed
}

try {
    $InputLocalMailboxes = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCsvLogFolderPath' -DefaultValue ''
    # FIX: Local file must be the 'Local' CSV, not the 'Remote' CSV
$InputLocalMailboxesFile = Join-Path $InputLocalMailboxes "Exchange_OnPrem_Mailboxes_AllDomains.csv"
} catch {
    Write-Error "Error retrieving local configuration path InputLocalMailboxesFile $_"
    WriteLog -Message "ERROR retrieving local configuration path: $_" "ERROR"
    # Do not exit here; we will fallback to live Exchange if needed
}

# --- Decide whether to use cached CSVs or live Exchange ---
function Test-FileIsFresh {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$MaxAgeHours = 4
    )
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $age = (Get-Date) - (Get-Item -LiteralPath $Path).LastWriteTime
        return ($age.TotalHours -lt $MaxAgeHours)
    } catch {
        Write-Warning "Failed to read LastWriteTime for '$Path'. Error: $($_.Exception.Message)"
        return $false
    }
}

$localFresh  = if ($InputLocalMailboxesFile)  { Test-FileIsFresh -Path $InputLocalMailboxesFile  -MaxAgeHours $FileFreshnessHours } else { $false }
$remoteFresh = if ($InputRemoteMailboxesFile) { Test-FileIsFresh -Path $InputRemoteMailboxesFile -MaxAgeHours $FileFreshnessHours } else { $false }
$useCache    = $localFresh -and $remoteFresh

WriteLog -Message "Local CSV fresh:  $localFresh ($InputLocalMailboxesFile)"
WriteLog -Message "Remote CSV fresh: $remoteFresh ($InputRemoteMailboxesFile)"
Write-Output "CSV cache decision: useCache = $useCache (threshold = $FileFreshnessHours hours)"

try {
    Write-Output "------------------------------------------------------------"
    Write-Output "Starting Exchange Mailbox Report Script at $(Get-Date)."

    # --- FILTERS (common) ---
    $allowedTypes = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox', 'RemoteUserMailbox', 'RemoteSharedMailbox')

    # --- Acquire mailbox data (CSV cache preferred) ---
    Write-Output "Step 1/3: Acquiring mailboxes..."
    $fullMailboxData = [System.Collections.Generic.List[PSObject]]::new()
    $statsHashTable  = @{}  # used only in live mode

    if ($useCache) {
        Write-Output "Using cached CSV files (both < $FileFreshnessHours hours)."

        # If your CSVs are semicolon-delimited, add: -Delimiter ';'
        $onPremRaw  = Import-Csv -Path $InputLocalMailboxesFile
        $remoteRaw  = Import-Csv -Path $InputRemoteMailboxesFile

        # Filter local (RecipientType or RecipientTypeDetails)
        $onPremMailboxes = $onPremRaw | Where-Object {
            ($_.RecipientType -and ($_.RecipientType -in @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox'))) -or
            ($_.RecipientTypeDetails -and ($_.RecipientTypeDetails -in @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox')))
        }

        # Filter remote (RecipientTypeDetails)
        $remoteMailboxes = $remoteRaw | Where-Object {
            $_.RecipientTypeDetails -in @('RemoteUserMailbox','RemoteSharedMailbox')
        }

        # Assemble local from CSV
        foreach ($row in $onPremMailboxes) {
            $rt = if ($row.PSObject.Properties.Name -contains 'RecipientTypeDetails' -and $row.RecipientTypeDetails) { 
                $row.RecipientTypeDetails 
            } else { 
                $row.RecipientType 
            }

            $dn = $row.DistinguishedName
            $adDomain = $null
            if ($dn) {
                $parts = $dn -split ',DC='
                if ($parts.Count -gt 1) { $adDomain = $parts[1..($parts.Count - 1)] -join '.' }
            }

            # Normalize AccountDisabled
            $accDisabled = $false
            if ($row.AccountDisabled) { $accDisabled = ($row.AccountDisabled.ToString() -match '^(?i:true|1|yes)$') }

            # Parse TotalItemSize-In-MB as double
            $totalMb = 0.0
            if ($row.'TotalItemSize-In-MB') { [void][double]::TryParse($row.'TotalItemSize-In-MB'.ToString().Replace(',','.'), [ref]$totalMb) }

            $fullMailboxData.Add([PSCustomObject]@{
                DistinguishedName    = $dn
                AccountDisabled      = $accDisabled
                RecipientTypeDetails = $rt
                ADDomain             = $adDomain
                TotalItemSizeToMB    = [math]::Round($totalMb, 2)
            })
        }

        # Assemble remote from CSV
        foreach ($row in $remoteMailboxes) {
            $dn = $row.DistinguishedName
            $adDomain = $null
            if ($dn) {
                $parts = $dn -split ',DC='
                if ($parts.Count -gt 1) { $adDomain = $parts[1..($parts.Count - 1)] -join '.' }
            }

            $accDisabled = $false
            if ($row.AccountDisabled) { $accDisabled = ($row.AccountDisabled.ToString() -match '^(?i:true|1|yes)$') }

            $fullMailboxData.Add([PSCustomObject]@{
                DistinguishedName    = $dn
                AccountDisabled      = $accDisabled
                RecipientTypeDetails = $row.RecipientTypeDetails
                ADDomain             = $adDomain
                TotalItemSizeToMB    = 0
            })
        }

        Write-Output "From cache: processed $($fullMailboxData.Count) mailboxes."

    } else {
        Write-Output "CSV cache not available or stale (older than $FileFreshnessHours hours). Falling back to live Exchange."

        if (-not (EnsureExchangePSSnapinLoaded)) {
            Write-Error "Exchange environment not ready. Stopping."
            $errorMessage = "Exchange environment not ready. Stopping."
            $body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
            SendEmailHtmlReport -BodyHtml $body
            throw $errorMessage
        }
        Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequireExchangeOnPrem | Out-Null

        # View entire forest in AD
        try {
            Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
            WriteLog -Message "Set-ADServerSettings -ViewEntireForest $true applied successfully."
        } catch {
            $errorMessage = "CRITICAL ERROR: Failed to apply 'Set-ADServerSettings -ViewEntireForest $true'. Ensure RSAT AD tools are installed and you have necessary permissions. Message: $($_.Exception.Message)"
            WriteLog -message  $errorMessage
            $body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
            SendEmailHtmlReport -BodyHtml $body
            throw $errorMessage
        }

        $getMailboxParams = @{ ResultSize = 'Unlimited' }

        Write-Output "Filtering on mailbox types: $($allowedTypes -join ', ')"
        Write-Output "Gathering on-premises mailboxes..."
        $onPremMailboxes = Get-Mailbox @getMailboxParams | Where-Object { $allowedTypes -contains $_.RecipientTypeDetails.ToString() }

        Write-Output "Gathering remote mailboxes (Hybrid)..."
        $remoteMailboxes = Get-RemoteMailbox @getMailboxParams | Where-Object { $allowedTypes -contains $_.RecipientTypeDetails.ToString() }

        Write-Output "Found $($onPremMailboxes.Count) on-premises mailboxes and $($remoteMailboxes.Count) remote mailboxes."
        Write-Output "Processing statistics in batches for performance..."

        # Collect mailbox statistics in batches
        if ($onPremMailboxes.Count -gt 0) {
            $batchSize = 100
            for ($i = 0; $i -lt $onPremMailboxes.Count; $i += $batchSize) {
                $end = [Math]::Min($i + $batchSize - 1, $onPremMailboxes.Count - 1)
                $batch = $onPremMailboxes[$i..$end]
                try {
                    $batchStats = $batch | Get-MailboxStatistics -ErrorAction Stop
                    foreach ($stat in $batchStats) {
                        if ($stat.DisplayName) { $statsHashTable[$stat.DisplayName] = $stat }
                    }
                } catch {
                    Write-Warning "An error occurred processing a batch of mailboxes. Some statistics might be missing. Error: $($_.Exception.Message)"
                }
            }
        }

        # Assemble local mailboxes
        foreach ($mailbox in $onPremMailboxes) {
            $stats = $statsHashTable[$mailbox.DisplayName]
            $TotalItemSizeToMB = if ($null -ne $stats -and $null -ne $stats.TotalItemSize) { $stats.TotalItemSize.Value.ToMB() } else { 0 }
            $fullMailboxData.Add([PSCustomObject]@{
                DistinguishedName    = $mailbox.DistinguishedName
                AccountDisabled      = $mailbox.AccountDisabled 
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
                ADDomain             = ($mailbox.DistinguishedName -split ',DC=')[1..($mailbox.DistinguishedName.Split(',DC=').Count - 1)] -join '.'
                TotalItemSizeToMB    = $TotalItemSizeToMB
            })
        }

        # Assemble remote mailboxes
        foreach ($mailbox in $remoteMailboxes) {
            $fullMailboxData.Add([PSCustomObject]@{
                DistinguishedName    = $mailbox.DistinguishedName
                AccountDisabled      = $mailbox.AccountDisabled 
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
                ADDomain             = ($mailbox.DistinguishedName -split ',DC=')[1..($mailbox.DistinguishedName.Split(',DC=').Count - 1)] -join '.'
                TotalItemSizeToMB    = 0
            })
        }

        Write-Output "Successfully processed $($fullMailboxData.Count) mailboxes (live)."
    }

    if ($fullMailboxData.Count -eq 0) {
        Write-Warning "No mailboxes were found or processed. The report will not be generated."
        return
    }

    # Optional: filter by -TargetDomains (ADDomain = derived from DN)
    if ($TargetDomains -and $TargetDomains.Count -gt 0) {
        Write-Output "Applying TargetDomains filter: $($TargetDomains -join ', ')"
        $fullMailboxData = $fullMailboxData | Where-Object { $_.ADDomain -in $TargetDomains }
        Write-Output "After domain filter: $($fullMailboxData.Count) mailboxes."
        if ($fullMailboxData.Count -eq 0) {
            Write-Warning "No mailboxes matched the requested TargetDomains."
        }
    }

    # Step 2: Define the report columns.
    Write-Output "Step 2/3: Defining report columns..."
    $allMailboxTypes = $allowedTypes | Sort-Object

    # Step 3: Group and generate the report...
    Write-Output "Step 3/3: Grouping data and generating the report..."
    $groupedByDomain = $fullMailboxData | Group-Object -Property ADDomain

    $report = $groupedByDomain | ForEach-Object {
        $domain = $_.Name
        $mailboxesInDomain = $_.Group
        $measureResult = $mailboxesInDomain | Where-Object { $_.TotalItemSizeToMB -ne $null } | Measure-Object -Property TotalItemSizeToMB -Sum
        $totalSizeMB = $measureResult.Sum
        $totalSizeGB = [math]::Round($totalSizeMB / 1024, 2)
        $mailboxCount = $mailboxesInDomain.Count
        $remoteMailboxTypes = @('RemoteUserMailbox', 'RemoteSharedMailbox')
        $localMailboxCount = ($mailboxesInDomain | Where-Object { $_.RecipientTypeDetails -notin $remoteMailboxTypes }).Count
        $remoteMailboxCount = ($mailboxesInDomain | Where-Object { $_.RecipientTypeDetails -in $remoteMailboxTypes }).Count
        $disabledCount = ($mailboxesInDomain | Where-Object { $_.AccountDisabled -eq $true }).Count
        $enabledCount = ($mailboxesInDomain | Where-Object { $_.AccountDisabled -eq $false }).Count
        
        $outputRow = [ordered]@{
            "Date" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            "DomainName" = $domain
            "TotalMailboxCount" = $mailboxCount
            "TotalLocalMailboxCount" = $localMailboxCount
            "TotalRemoteMailboxCount" = $remoteMailboxCount
            "EnabledAccounts" = $enabledCount
            "DisabledAccounts" = $disabledCount
            "TotalLocalMailboxSizeGB" = $totalSizeGB
        }
        
        $countsInThisDomain = $mailboxesInDomain | Group-Object -Property RecipientTypeDetails -NoElement
        foreach ($type in $allMailboxTypes) {
            $countForType = ($countsInThisDomain | Where-Object { $_.Name -eq $type }).Count
            $outputRow[$type] = if ($countForType) { $countForType } else { 0 }
        }
        
        [PSCustomObject]$outputRow
    }

    Write-Output "Report object generated with $($report.Count) rows."

    # --- Logic for appending to the CSV file ---
    try {
        if (Test-Path -Path $csvOutputPath) {
            Write-Output "CSV file exists. Appending new data..."
            $report | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Select-Object -Skip 1 | Add-Content -Path $csvOutputPath -Encoding UTF8
        } else {
            Write-Output "CSV file does not exist. Creating new file..."
            $report | Export-Csv -Path $csvOutputPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        }
        Write-Output "Successfully finished writing to CSV file."
    } catch {
        throw "Failed to write the report to '$csvOutputPath'. Please check file permissions. Error: $($_.Exception.Message)"
    }

    # --- NEW: Copy the CSV to csvOutputPath2 (LatestCsvFolderPath) ---
    try {
        if ($csvOutputPath2) {
            $destDir = Split-Path -Path $csvOutputPath2 -Parent
            if (-not (Test-Path -LiteralPath $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                WriteLog -Message "Created destination directory for secondary CSV copy: $destDir"
            }
            Copy-Item -Path $csvOutputPath -Destination $csvOutputPath2 -Force
            WriteLog -Message "Secondary CSV copy done: $csvOutputPath -> $csvOutputPath2"
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvOutputPath2
            $global:CsvCopyOk = $true
        } else {
            $global:CsvCopyOk = $false
            WriteLog -Message "Secondary CSV copy skipped: csvOutputPath2 not defined." "WARN"
        }
    } catch {
        $global:CsvCopyOk = $false
        Write-Warning "Failed to copy CSV to '$csvOutputPath2'. Error: $($_.Exception.Message)"
        WriteLog -Message "Failed to copy CSV to '$csvOutputPath2'. Error: $($_.Exception.Message)" "WARN"
    }

    # ---------------------------
    # Step 4: End-of-run SUMMARY
    # ---------------------------
    try {
        Write-Output ""
        Write-Output "End-of-run Summary (by domain):"

        $summaryColumns = @(
            'DomainName',
            'TotalMailboxCount',
            'TotalLocalMailboxCount',
            'TotalRemoteMailboxCount',
            'EnabledAccounts',
            'DisabledAccounts',
            'TotalLocalMailboxSizeGB'
        )

        $summary = $report | Select-Object $summaryColumns

        # Build TOTAL row
        $totalRow = [ordered]@{ 'DomainName' = 'TOTAL' }
        foreach ($col in $summaryColumns) {
            if ($col -ne 'DomainName') {
                $sum = ($summary | Measure-Object -Property $col -Sum).Sum
                if ($col -eq 'TotalLocalMailboxSizeGB') { $sum = [math]::Round([double]$sum, 2) }
                $totalRow[$col] = $sum
            }
        }
        $summaryWithTotal = @($summary + ([PSCustomObject]$totalRow))

        # Display to console
        $summaryWithTotal | Sort-Object {
            if ($_.DomainName -eq 'TOTAL') { [int]::MaxValue } else { $_.DomainName }
        } | Format-Table -AutoSize | Out-Host

        # Save to a summary CSV (overwrite with latest run)
        $summaryWithTotal | Export-Csv -Path $summaryCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $summaryCsvPath
        WriteLog -Message "Summary saved to $summaryCsvPath"
    }
    catch {
        Write-Warning "Failed to produce summary table. Error: $($_.Exception.Message)"
    }
    
    #region Cleanup
	# Clean up old CSV files + old log files
	# Automatically excludes all generated CSVs via global:csvGeneratedPaths + current transcript and log files via global variables
	RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
	RemoveOldFiles -Path $logPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
    WriteLog -Message "$TaskName completed."
    $ScriptSucceeded = $true
    # Do not Stop-Transcript here; handled in finally
    #endregion
}
catch {
    Write-Error "A critical error occurred and the script has been stopped."
    Write-Error "ERROR: $($_.Exception.Message)"
    Write-Error "DETAILS: $($_.ToString())"
    # THROW instead of EXIT to ensure the final email still gets sent in 'finally'
    throw
}
finally {
    Write-Output "Script execution finished at $(Get-Date)."
    Write-Output "------------------------------------------------------------"

    # Mark daily success only if the main run succeeded
    if ($ScriptSucceeded) {
        try {
            New-Item -Path $DailySuccessLockFile -ItemType File -Force | Out-Null
            try { WriteLog -Message "Daily success lock created: $DailySuccessLockFile" } catch {}
        } catch {
            Write-Warning "Failed to create daily success lock: $($_.Exception.Message)"
            try { WriteLog -Message "Failed to create daily success lock ($DailySuccessLockFile): $($_.Exception.Message)" "WARN" } catch {}
        }
    }

    # Ensure transcript is closed before email
    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}

    # Compose and send the end-of-run email using SendFileListEmailReport
    try {
        # Compute elapsed time
        $elapsed = (Get-Date) - $StartTime
        $elapsedFormatted = "{0:00}:{1:00}:{2:00}" -f [int]$elapsed.Hours, [int]$elapsed.Minutes, [int]$elapsed.Seconds

        # Derive totals if $report exists
        $totalsLine = ""
        if ($null -ne $report -and $report.Count -gt 0) {
            $sumTotal      = ($report | Measure-Object -Property TotalMailboxCount -Sum).Sum
            $sumLocal      = ($report | Measure-Object -Property TotalLocalMailboxCount -Sum).Sum
            $sumRemote     = ($report | Measure-Object -Property TotalRemoteMailboxCount -Sum).Sum
            $sumEnabled    = ($report | Measure-Object -Property EnabledAccounts -Sum).Sum
            $sumDisabled   = ($report | Measure-Object -Property DisabledAccounts -Sum).Sum
            $sumSizeGB     = [math]::Round( ($report | Measure-Object -Property TotalLocalMailboxSizeGB -Sum).Sum, 2)
            $domainCount   = $report.Count
            $totalsLine    = "Domains: $domainCount | Total: $sumTotal | Local: $sumLocal | Remote: $sumRemote | Enabled: $sumEnabled | Disabled: $sumDisabled | LocalSize(GB): $sumSizeGB"
        } else {
            $totalsLine = "No aggregated totals available."
        }

        $runMode    = if ($useCache) { "Completed (CSV cache mode)" } else { "Completed (Live Exchange mode)" }
        $emailTitle = "$TaskName - $runMode"

        # Build HTML summary table for the email body
        $summaryHtml = ""
        if ($null -ne $summaryWithTotal -and $summaryWithTotal.Count -gt 0) {
            $summaryHtml += @"
<h3 style='font-family:Segoe UI,Arial,sans-serif; margin:12px 0 6px;'>Mailbox Summary by Domain</h3>
<table style='border-collapse:collapse; font-family:Segoe UI,Arial,sans-serif; font-size:13px;'>
  <thead>
    <tr style='background:#f3f3f3;'>
      <th style='border:1px solid #ddd; padding:6px 8px; text-align:left;'>DomainName</th>
      <th style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>TotalMailboxCount</th>
      <th style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>TotalLocalMailboxCount</th>
      <th style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>TotalRemoteMailboxCount</th>
      <th style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>EnabledAccounts</th>
      <th style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>DisabledAccounts</th>
      <th style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>TotalLocalMailboxSizeGB</th>
    </tr>
  </thead>
  <tbody>
"@
            foreach ($row in $summaryWithTotal) {
                $isTotal = ($row.DomainName -eq 'TOTAL')
                $rowStyle = if ($isTotal) { "font-weight:bold; background:#fff8e1;" } else { "" }

                $summaryHtml += @"
    <tr style='$rowStyle'>
      <td style='border:1px solid #ddd; padding:6px 8px; text-align:left;'>$($row.DomainName)</td>
      <td style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>$($row.TotalMailboxCount)</td>
      <td style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>$($row.TotalLocalMailboxCount)</td>
      <td style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>$($row.TotalRemoteMailboxCount)</td>
      <td style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>$($row.EnabledAccounts)</td>
      <td style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>$($row.DisabledAccounts)</td>
      <td style='border:1px solid #ddd; padding:6px 8px; text-align:right;'>$($row.'TotalLocalMailboxSizeGB')</td>
    </tr>
"@
            }

            $summaryHtml += @"
  </tbody>
</table>
"@
        } else {
            $summaryHtml = "<p style='font-family:Segoe UI,Arial,sans-serif; font-size:13px;'>No summary data available.</p>"
        }

        # Compose the Message (header + summary HTML)
        $emailMessage = @"
Host: $env:COMPUTERNAME<br/>
Mode: $runMode<br/>
Elapsed: $elapsedFormatted<br/>
$totalsLine<br/><br/>
$summaryHtml
<br/>
Outputs:<br/>
- $csvOutputPath<br/>
- $summaryCsvPath<br/>
- $csvOutputPath2
"@

        # Include the CSVs (main, summary, and the second copy if present)
        $filesForReport = @()
        if (Test-Path $csvOutputPath)   { $filesForReport += $csvOutputPath }
        if (Test-Path $summaryCsvPath)  { $filesForReport += $summaryCsvPath }
        if ($csvOutputPath2 -and (Test-Path $csvOutputPath2)) { $filesForReport += $csvOutputPath2 }

        SendFileListEmailReport -Files $filesForReport -Title $emailTitle -Message $emailMessage
    } catch {
        Write-Warning "Failed to send end-of-run email. Error: $($_.Exception.Message)"
    }
}
