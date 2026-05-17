<#
.SYNOPSIS
    Inventories all local Exchange mailboxes in an on-premises environment.

.DESCRIPTION
    This script scans and exports detailed information about Exchange mailboxes, including:
      - Full mailbox properties (name, alias, UPN, OU, etc.)
      - Permissions (FullAccess, SendAs, excluding NT AUTHORITY\SELF)
      - Mailbox statistics (size, item count, last logon, etc.)
      - Archive status and retention policies
      - Mobile device associations

    Supports both full-forest and targeted OU scans. Results are exported to CSV files.
    Includes robust logging, error handling, and backup of previous exports.
    Designed for Exchange 2016 servers with management tools and AD modules installed.
    Parameters allow customization of output paths, permission inclusion, and overwrite behavior.

.NOTES
    Version: 1.0
    Author: https://github.com/khda79/M365
    Requirements: Exchange 2016 Management Tools, Active Directory module
#>

[CmdletBinding()]
param (
    [string]$Tenant = 'test',
[Parameter(Mandatory = $false)]
    [string]$OutputPath,
	[Parameter(Mandatory = $false)]
    [string]$OutputPathOnlyADPermission,
    [Parameter(Mandatory = $false)]
    [string[]]$IncludedOrganizationalUnit = @(), 
    [Parameter(Mandatory = $false)]
    [bool]$DetectAllDomains = $true, 
    [Parameter(Mandatory = $false)]
    [bool]$IncludeADPermission = $false,
    [Parameter(Mandatory = $false)]
    [bool]$OnlyADPermission = $false,
    [Parameter(Mandatory = $false)]
    [bool]$ForceOverwriteCSV = $true
)
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $p = Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'
        if (Test-Path -LiteralPath $p) { return $p }
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'SmartM365.global.local.json'
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'SmartM365.global.local.json'
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
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCsvLogFolderPath' -DefaultValue $OutputPath
$LimitResultSize = $null
if ($LimitResultSize) {
    $TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion (LimitResultSize : $LimitResultSize)..."
}
$scriptdatamailbox = $false
$scriptdatamegewithperm = $true
$scriptdatamegewithbatch = $true


# Atomic CSV export helper
function Export-CsvAtomic {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateSet("UTF8","UTF8BOM","Unicode","ASCII","Default")]
        [string]$Encoding = "UTF8"
    )

    $parent = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    $leaf = Split-Path -Path $Path -Leaf
    $tmp = Join-Path -Path $parent -ChildPath ("{0}.tmp.{1}" -f $leaf, ([guid]::NewGuid().ToString("N")))

    try {
        $InputObject | Export-Csv -Path $tmp -NoTypeInformation -Encoding $Encoding -ErrorAction Stop
        Move-Item -LiteralPath $tmp -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        if (Test-Path -LiteralPath $tmp) {
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        }
    }
}

[string]$inputFolderCSVfiles
[string]$excludeMailboxesFile
[string]$outputConsolidatedCsvPath
[string]$outputConsolidatedCsvPath2
[string]$outputFileNamePrefix = "MigrationBatch"

[bool]$ExcludeAllDisabledAccounts = $false
[bool]$ExcludeDisabledAccountsExceptForSharedMailboxes = $false
[bool]$ExcludeDisabledMailboxes = $false
[bool]$IncludeSpecificLastLogonCriteria = $false
[bool]$SimpleBatchCsv = $true
[bool]$EnabledBatchFileCreation = $false
[string[]]$ExcludeSamAccountNamePatterns
[string[]]$excludeFullAccessSamAccounts
[string[]]$excludeSendAsSamAccounts
[string[]]$excludeSendOnBehalfToSamAccounts

[double]$maxBatchSizeMB = 500000.0
[int]$MaxMailboxesPerBatchFile = 500

[string]$SendFileListEmailReportFileName

function Join-ModulePath {
param([Parameter(Mandatory)][string]$FileName)
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
return (Join-Path (Join-Path (Join-Path (Join-Path $repoRoot 'Modules') 'SmartM365.Core') 'Compatibility\WindowsPowerShell5') $FileName)
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
	WriteLog -Message $MyInvocation.MyCommand.Name
	WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
	$OutputPath = $InitializeOutputPath
	WriteLog -Message "Starting $TaskName..."
	WriteLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
	exit 1
}
#endregion

if (-not (EnsureExchangePSSnapinLoaded)) {
    Write-Error "Exchange environment not ready. Exiting script."
	$errorMessage = "Exchange environment not ready. Exiting script."
	$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
	SendEmailHtmlReport -BodyHtml $body
    exit 1
}
$StartTime = Get-Date

try {
	$OutputPathOnlyADPermission = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxOnlyAdPermissionCsvLogFolderPath' -DefaultValue ''
} catch {
	Write-Error "Error retrieving local configuration path LastOutput $_"
	WriteLog -Message "ERROR retrieving local configuration path: $_" "ERROR"
	return
}

if (-not (Test-Path $OutputPathOnlyADPermission)) {
	Write-Host "The share '$OutputPathOnlyADPermission' is not available. Stopping the script." -ForegroundColor Red
	$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : The share '$OutputPathOnlyADPermission' is not available. Stopping the script."
	SendEmailHtmlReport -BodyHtml $body
	exit
} else {
	Write-Host "The network share '$OutputPathOnlyADPermission' is available. Continuing the script..." -ForegroundColor Green
Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath, $OutputPathOnlyADPermission) -RequireExchangeOnPrem -RequireActiveDirectoryRead | Out-Null

}

# Refined parameter logic for permissions
if ($OnlyADPermission) {
    $IncludeADPermission = $true
    if ([string]::IsNullOrWhiteSpace($OutputPathOnlyADPermission)) {
        throw "OutputPathOnlyADPermission is required when OnlyADPermission is set"
    }
    $OutputPath = $OutputPathOnlyADPermission

    # In OnlyADPermission mode, do not merge with permissions inventory nor batch naming.
    $scriptdatamegewithperm = $false
    $scriptdatamegewithbatch = $false
}
if ($IncludedOrganizationalUnit.Count -ne 0) {
    $DetectAllDomains = $false
}
$TimeStampForLogOnly = Get-Date -Format "yyyyMMdd_HHmmss" # Timestamp for log files to ensure uniqueness per run
				
# Initialize global caches
if (-not $Global:DomainInfoCache) {
    $Global:DomainInfoCache = @{} # For NetBIOSName to DNSRoot (FQDN) mapping
}
if (-not $Global:ADObjectCache) {
    $Global:ADObjectCache = @{} # For SID to ADObject mapping (or $null if not found)
}
if (-not $Global:GroupMemberCache) {
    $Global:GroupMemberCache = @{} # For Group DN to member SamAccountNames list mapping (or error/empty marker)
}
if (-not $Global:DomainFQDNToNetBIOSCache) {
    $Global:DomainFQDNToNetBIOSCache = @{} # For Domain FQDN to NetBIOS Name mapping
}
$Global:ScriptOverallMailboxData = @() # To accumulate all mailbox data for a final combined CSV
$Script:MailboxesProcessingLogFile = $null # Initialize script-scoped variable for detailed log file path

# Variable to track successful script completion
$InventoryCompletedSuccessfully = $false

		
try { # Main try block for script execution and interruption handling
    WriteLog -Message "Starting script '$PSCommandPath' - Version $ScriptVersion"
    Write-Host "Starting script '$($MyInvocation.MyCommand.Name)' - Version $ScriptVersion ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
    Write-Host "Output Path for this run: $OutputPath"
    WriteLog -Message "Output Path for this run: $OutputPath"
    Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)"
    WriteLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
    WriteLog -Message "Effective permission flags: IncludeADPermission = $IncludeADPermission, OnlyADPermission = $OnlyADPermission, ForceOverwriteCSV = $ForceOverwriteCSV"
    Write-Host "Effective permission flags: IncludeADPermission = $IncludeADPermission, OnlyADPermission = $OnlyADPermission, ForceOverwriteCSV = $ForceOverwriteCSV"
    Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))

    # Set AD server settings to view the entire forest
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

    #region Function Definitions

    # This function contains the detailed mailbox processing logic
    function MailboxesProcessing2 {
        param (
            [Parameter(Mandatory=$true)]
            [string[]]$IncludedLDAPPaths
        )
        $FunctionStartTime = Get-Date

        # --- START NEW LOGIC FOR LOG FILE IDENTIFIER ---
        $logPathIdentifier = "AllMailboxes"
        if ($IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0) {
            if ($IncludedLDAPPaths.Count -eq 1) {
                $firstPath = $IncludedLDAPPaths[0]
                $pathParts = $firstPath -split ','
                $identifierPartAttempt = $pathParts[0]
                foreach ($part in $pathParts) {
                    if ($part.StartsWith("OU=", [System.StringComparison]::OrdinalIgnoreCase) -or $part.StartsWith("CN=", [System.StringComparison]::OrdinalIgnoreCase)) {
                        $identifierPartAttempt = $part
                        break
                    }
                }
                $identifierPartClean = $identifierPartAttempt -replace '^(OU=|CN=|DC=)', ''
                $identifierPartClean = $identifierPartClean -replace '[^a-zA-Z0-9_-]', ''
                if ($identifierPartClean.Length -gt 25) {
                    $identifierPartClean = $identifierPartClean.Substring(0, 25)
                }
                if (-not [string]::IsNullOrWhiteSpace($identifierPartClean)) {
                    $logPathIdentifier = $identifierPartClean
                } else {
                    $logPathIdentifier = "SinglePath_UnknownFormat"
                }
            } else {
                $logPathIdentifier = "MultiplePaths_($($IncludedLDAPPaths.Count))"
            }
        }
        # --- END NEW LOGIC FOR LOG FILE IDENTIFIER ---

        # Use script-scoped variable for the detailed log file path
        $Script:MailboxesProcessingLogFile = Join-Path -Path $logPath -ChildPath "MailboxesProcessingDetails_Local_${logPathIdentifier}_$TimeStampForLogOnly.log"

        # Logging function specific to MailboxesProcessing2
        function Write-LogMailboxesProcessing {
            param (
                [Parameter(Mandatory = $true)]
                [string]$Message
            )
            $ProcessingTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $ProcessingLogEntry = "$ProcessingTimestamp - $Message"
            try {
                Add-Content -Path $Script:MailboxesProcessingLogFile -Value $ProcessingLogEntry -ErrorAction Stop
            } catch {
                $mainLogErrorMessage = "ERROR (MailboxesProcessing2): Error writing to MailboxesProcessing log file ($($Script:MailboxesProcessingLogFile)): $($_.Exception.Message). Original Message: $Message"
                WriteLog -message $mainLogErrorMessage
                Write-Host -ForegroundColor Red $mainLogErrorMessage
            }
        }
       
        # --- START CSV File Overwrite Check (Combined CSV) ---
        # This check is now primarily handled in the main script logic before calling Process-SpecificDomain or MailboxesProcessing
        # However, for the case where !DetectAllDomains and a single OU is provided, the per-domain/path CSV check is relevant here.
        if (-not $DetectAllDomains -and $IncludedLDAPPaths -and $IncludedLDAPPaths.Count -eq 1) {
            $singlePath = $IncludedLDAPPaths[0]
            $singlePathDomainName = "UnknownPathDomain" # Default
            if ($singlePath -match "DC=([^,]+)") {
                $tempDcParts = @()
                $singlePath -split ',' | Where-Object {$_ -like "DC=*"} | ForEach-Object {$tempDcParts += $_ -replace "DC="}
                if ($tempDcParts.Count -gt 0) {
                    $singlePathDomainName = $tempDcParts -join "."
                }
            } else {
                $singlePathDomainName = $singlePath -replace '^(OU=|CN=)','' -replace '[^a-zA-Z0-9.-]','_'
                if ($singlePathDomainName.Length -gt 30) {$singlePathDomainName = $singlePathDomainName.Substring(0,30)}
            }
           
            $perDomainCsvFileName = "Exchange_OnPrem_Mailboxes_$($singlePathDomainName)$(if($OnlyADPermission){'_OnlyADPermission'}else{''}).csv"
			$perDomainCsvFileFullPath = Join-Path -Path $OutputPath -ChildPath $perDomainCsvFileName

			# Define the base backup directory (e.g., C:\Users\A_khadaw8899\Documents\ExchangeMailboxesInventory\Backup)
			$baseExchangeMailboxesInventoryPath = (Get-Item $OutputPath).Parent.FullName # Adjusted to get ExchangeMailboxesInventory base
			$backupBaseDir = Join-Path -Path $baseExchangeMailboxesInventoryPath -ChildPath "Backup"

			if (Test-Path $perDomainCsvFileFullPath) {
				if (-not $ForceOverwriteCSV) {
					# Create a timestamped directory inside the backup base directory
					$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
					$currentBackupDir = Join-Path -Path $backupBaseDir -ChildPath $timestamp

					# Create the full backup path including the timestamped directory
					if (-not (Test-Path $currentBackupDir)) {
						New-Item -ItemType Directory -Path $currentBackupDir -Force | Out-Null
						Write-LogMailboxesProcessing "INFO: Created timestamped backup directory: '$currentBackupDir'."
					}

					# The backup file name should be the same as the source file name
					$backupFileName = Split-Path -Path $perDomainCsvFileFullPath -Leaf
					$backupFilePath = Join-Path -Path $currentBackupDir -ChildPath $backupFileName

					$message = "Per-domain/path CSV file '$perDomainCsvFileFullPath' for '$singlePath' already exists and -ForceOverwriteCSV is `$false. Backing up the existing file to '$backupFilePath'."
					Write-LogMailboxesProcessing "WARNING: $message"
					Write-Host -ForegroundColor Yellow $message

					# Copy the file to the backup directory
					try {
						Copy-Item -Path $perDomainCsvFileFullPath -Destination $backupFilePath -Force -ErrorAction Stop
						Write-LogMailboxesProcessing "INFO: Successfully backed up '$perDomainCsvFileFullPath' to '$backupFilePath'."

						# After backing up, remove the original file
						Remove-Item -Path $perDomainCsvFileFullPath -ErrorAction Stop
						Write-LogMailboxesProcessing "INFO: Successfully removed original file '$perDomainCsvFileFullPath'."
					} catch {
						$errorMessage = "CRITICAL ERROR: Failed to process the per-domain CSV file '$perDomainCsvFileFullPath'. Error: $($_.Exception.Message)"
						Write-LogMailboxesProcessing $errorMessage
						Write-Host -ForegroundColor Red $errorMessage
						return # Stop processing for this domain if copy or removal fails
					}
				} else {
					$message = "Per-domain/path CSV file '$perDomainCsvFileFullPath' for '$singlePath' exists and -ForceOverwriteCSV is `$true. Existing file will be deleted."
					Write-LogMailboxesProcessing "INFO: $message"
					Write-Host -ForegroundColor Cyan $message
					try {
						Remove-Item -Path $perDomainCsvFileFullPath -Force -ErrorAction Stop
						Write-LogMailboxesProcessing "INFO: File '$perDomainCsvFileFullPath' deleted successfully."
					} catch {
						$errorMessage = "ERROR: Could not delete existing per-domain/path CSV file '$perDomainCsvFileFullPath'. Message: $($_.Exception.Message)"
						Write-LogMailboxesProcessing $errorMessage
						Write-Error $errorMessage
					}
				}
			}
        }
        # --- END Per-Path/Domain CSV Overwrite Check ---


        $output = @()
        $Domains = @()
        $MailboxesByDomain = @{}
        $ShowAll = $false
        $ExpandGroups = $true
        $i = 0
        $totalMailBox = 0
        $MailboxScanStartTime = $null
       
        Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
        Write-LogMailboxesProcessing "MailboxesProcessing2: Starting data collection and export. Log Identifier: $logPathIdentifier"
        if ($IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0) {
            Write-Host "Filtering by provided LDAP paths ($($IncludedLDAPPaths.Count) paths) ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-LogMailboxesProcessing "Filtering by LDAP paths: $($IncludedLDAPPaths -join '; ')"
        }
        Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))

        $AllMailbox = @()
        if ($IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0) {
            Write-Host -ForegroundColor:Cyan "Retrieving mailboxes from specified paths. Please wait, this may take some time… $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            $totalOUs = $IncludedLDAPPaths.Count
            $ouCounter = 0
            foreach ($ou in $IncludedLDAPPaths) {
                $ouCounter++
                Write-Host -ForegroundColor:Yellow "Checking existence and retrieving mailboxes from '$ou' (Path $ouCounter of $totalOUs)..."
                Write-LogMailboxesProcessing "Attempting to retrieve mailboxes from path: '$ou' (Path $ouCounter of $totalOUs)"
                try {
                    $adObjectExists = $false
                    if ($ou.StartsWith("OU=", [System.StringComparison]::OrdinalIgnoreCase)) {
                        if (Get-ADOrganizationalUnit -Identity $ou -ErrorAction SilentlyContinue) {
                            $adObjectExists = $true
                        }
                    } elseif ($ou.StartsWith("DC=", [System.StringComparison]::OrdinalIgnoreCase)) {
                        if (Get-ADDomain -Identity ($ou -replace 'DC=','' -replace ',','.') -ErrorAction SilentlyContinue) {
                             $adObjectExists = $true
                        } else {
                             Write-LogMailboxesProcessing "Note: Could not verify domain DN '$ou' with Get-ADDomain. Get-Mailbox will be attempted."
                             $adObjectExists = $true # Assume it might be a valid scope for Get-Mailbox anyway
                        }
                    } else {
                         Write-LogMailboxesProcessing "Path '$ou' does not start with OU= or DC=. Assuming it's a specific DN Get-Mailbox might handle or it's an error. Attempting Get-Mailbox."
                         $adObjectExists = $true # Assume Get-Mailbox might handle it
                    }

                    if ($adObjectExists) {
                        Write-Host -ForegroundColor:Yellow "Retrieving mailboxes from '$ou'..."
                        Write-LogMailboxesProcessing "Retrieving mailboxes from '$ou'."
						if ($LimitResultSize) {
							$mailboxesInPath = Get-Mailbox -OrganizationalUnit $ou -ResultSize $LimitResultSize -ErrorAction Stop
						} else {
							$mailboxesInPath = Get-Mailbox -OrganizationalUnit $ou -ResultSize Unlimited -ErrorAction Stop
						}
                        $AllMailbox += $mailboxesInPath
                        Write-LogMailboxesProcessing "Found $($mailboxesInPath.Count) mailboxes in '$ou'."
                    } else {
                        $warningMessage = "The path '$ou' was not positively identified as an existing OU or resolvable Domain DN. It will be skipped."
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage"
                    }
                }
                catch {
                    $errorMessage = "Error while retrieving mailboxes from path '$ou': $($_.Exception.Message)."
                    Write-Error $errorMessage
                    Write-LogMailboxesProcessing "ERROR: $errorMessage"
                }
            }
        }
        else
        {
            Write-Host -ForegroundColor:Cyan "Retrieving ALL mailboxes (no specific paths provided). Please wait, this may take some time… $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-LogMailboxesProcessing "Retrieving ALL mailboxes (no specific paths provided)..."
            try {                
				if ($LimitResultSize) {
					$AllMailbox = Get-Mailbox -ResultSize $LimitResultSize -ErrorAction Stop
				} else {
					$AllMailbox = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
				}
				
            } catch {
                $errorMessage = "Error while retrieving all mailboxes: $($_.Exception.Message)."
                Write-Error $errorMessage
                Write-LogMailboxesProcessing "ERROR: $errorMessage"
            }
        }

        $totalMailBox = $AllMailbox.Count
        Write-Host -ForegroundColor Cyan "Total number of mailboxes to process for current scope: $totalMailBox"
        Write-LogMailboxesProcessing "Total number of mailboxes to process for current scope: $totalMailBox"

        if ($totalMailBox -eq 0) {
            $warningMessage = "No mailboxes were found to process for the current scope. $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-Warning $warningMessage
            Write-LogMailboxesProcessing $warningMessage
        }
        else
        {
            Write-Host -ForegroundColor:Cyan "Processing mailboxes. Please wait, this will likely take some time… $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            if ($OnlyADPermission) {
                Write-Host -ForegroundColor:Cyan "OnlyADPermission mode is active!"
                Write-LogMailboxesProcessing "OnlyADPermission mode is active. Standard mailbox details will be skipped."
            }
            Write-LogMailboxesProcessing "Starting detailed processing of $totalMailBox mailboxes for current scope."
            $MailboxScanStartTime = Get-Date
           
            $LocalAllOutput = @() # Collection for data from this specific MailboxesProcessing2 call

            Foreach ($Mbx in $AllMailbox) {
                $i++
                $Domain = ""
                try {
                    $DNParts = $Mbx.DistinguishedName -split ","
                    $DomainParts = $DNParts | Where-Object { $_ -like "DC=*" } | ForEach-Object { $_ -replace "DC=", "" }
                    $Domain = $DomainParts -join "."
                } catch {
                    $warningMessage = "Could not determine domain for mailbox $($Mbx.Name) from DN: $($Mbx.DistinguishedName)"
                    Write-Warning $warningMessage
                    Write-LogMailboxesProcessing "WARNING: $warningMessage"
                    $Domain = "Unknown"
                }

                Write-LogMailboxesProcessing "Processing mailbox ($i of $totalMailBox): $($Mbx.Name) - ($Domain)"
                $MbxStartTime = Get-Date
               
                if ($Domains -notcontains $Domain) {
                    $Domains += $Domain
                    $MailboxesByDomain[$Domain] = @()
                }       
               
                $userObj = New-Object PSObject
                $userObj | Add-Member NoteProperty -Name "DomainName" -Value $Domain
                $userObj | Add-Member NoteProperty -Name "BatchName" -Value ""
                $userObj | Add-Member NoteProperty -Name "Name" -Value $Mbx.Name
                $userObj | Add-Member NoteProperty -Name "DisplayName" -Value $Mbx.DisplayName
                $userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
                $userObj | Add-Member NoteProperty -Name "UserPrincipalName" -Value $Mbx.UserPrincipalName   
                $userObj | Add-Member NoteProperty -Name "SamAccountName" -Value $Mbx.SamAccountName
               
                if ($OnlyADPermission -eq $false)
                {
                    $userObj | Add-Member NoteProperty -Name "IsMailboxEnabled" -Value $Mbx.IsMailboxEnabled
                    $userObj | Add-Member NoteProperty -Name "IsValid" -Value $Mbx.IsValid
                    $userObj | Add-Member NoteProperty -Name "IsShared" -Value $Mbx.IsShared
                    $userObj | Add-Member NoteProperty -Name "IsLinked" -Value $Mbx.IsLinked
                    $userObj | Add-Member NoteProperty -Name "IsResource" -Value $Mbx.IsResource
                    $userObj | Add-Member NoteProperty -Name "AccountDisabled" -Value $Mbx.AccountDisabled
                    $userObj | Add-Member NoteProperty -Name "DistinguishedName" -Value $Mbx.DistinguishedName   
                    $userObj | Add-Member NoteProperty -Name "RecipientType" -Value $(if($null -ne $Mbx.RecipientTypeDetails){$Mbx.RecipientTypeDetails.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "RecipientOU" -Value $Mbx.OrganizationalUnit
                    $userObj | Add-Member NoteProperty -Name "PrimarySMTPaddress" -Value $(if($null -ne $Mbx.PrimarySmtpAddress){$Mbx.PrimarySmtpAddress.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "EmailAddresses" -Value (($Mbx.EmailAddresses | Where-Object {$_.PrefixString -eq 'smtp'} | ForEach-Object {$_.SmtpAddress}) -join ";")
                    $userObj | Add-Member NoteProperty -Name "Database" -Value $(if($null -ne $Mbx.Database){$Mbx.Database.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ServerName" -Value $Mbx.ServerName
                    $userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults   
                    $userObj | Add-Member NoteProperty -Name "ArchiveName" -Value ($Mbx.ArchiveName -join ";")
                    $userObj | Add-Member NoteProperty -Name "ArchiveStatus" -Value $(if($null -ne $Mbx.ArchiveStatus){$Mbx.ArchiveStatus.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ArchiveState" -Value $(if($null -ne $Mbx.ArchiveState){$Mbx.ArchiveState.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ArchiveQuota" -Value $(if($null -ne $Mbx.ArchiveQuota){$Mbx.ArchiveQuota.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ForwardingAddress" -Value $(if($null -ne $Mbx.ForwardingAddress){$Mbx.ForwardingAddress.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ForwardingSmtpAddress" -Value $(if($null -ne $Mbx.ForwardingSmtpAddress){$Mbx.ForwardingSmtpAddress.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "DeliverToMailboxAndForward" -Value $Mbx.DeliverToMailboxAndForward 
                    $userObj | Add-Member NoteProperty -Name "Department" -Value $Mbx.Department
                    $userObj | Add-Member NoteProperty -Name "Office" -Value $Mbx.Office
                    $userObj | Add-Member NoteProperty -Name "Manager" -Value $(if($null -ne $Mbx.Manager){$Mbx.Manager.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "WhenMailboxCreated" -Value $Mbx.WhenMailboxCreated
                    $userObj | Add-Member NoteProperty -Name "WhenChanged" -Value $Mbx.WhenChanged
                    $userObj | Add-Member NoteProperty -Name "WhenCreated" -Value $Mbx.WhenCreated   
                    $userObj | Add-Member NoteProperty -Name "MailboxLocations" -Value ($Mbx.MailboxLocations -join ";")
                    $userObj | Add-Member NoteProperty -Name "Identity" -Value $(if($null -ne $Mbx.Identity){$Mbx.Identity.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "ObjectCategory" -Value $(if($null -ne $Mbx.ObjectCategory){$Mbx.ObjectCategory.ToString()}else{""})
                    $userObj | Add-Member NoteProperty -Name "GrantSendOnBehalfTo" -Value (($Mbx.GrantSendOnBehalfTo | ForEach-Object {if($null -ne $_){$_.ToString()}else{""}}) -join ";")
                   
                    $ProhibitSendReceiveQuota = "N/A"
                    try {
                        if (($Mbx.UseDatabaseQuotaDefaults -eq $true)) {
                            $db = Get-MailboxDatabase $Mbx.Database -ErrorAction Stop
                            if ($db.ProhibitSendReceiveQuota.IsUnlimited) { $ProhibitSendReceiveQuota = "Unlimited" }
                            else { $ProhibitSendReceiveQuota = $db.ProhibitSendReceiveQuota.Value.ToMB() }
                        } elseif ($Mbx.ProhibitSendReceiveQuota.IsUnlimited) {
                            $ProhibitSendReceiveQuota = "Unlimited"
                        } else {
                            $ProhibitSendReceiveQuota = $Mbx.ProhibitSendReceiveQuota.Value.ToMB()
                        }
                    } catch {
                        $warningMessage = "Could not determine ProhibitSendReceiveQuota for $($Mbx.Name): $($_.Exception.Message)"
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage"
                        $ProhibitSendReceiveQuota = "Error"
                    }
                    $userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $ProhibitSendReceiveQuota
                   
                    # Initialize Mailbox Statistics properties
                    $userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value "N/A"
                    $userObj | Add-Member NoteProperty -Name "TotalItemSize-In-MB" -Value "N/A"
                    $userObj | Add-Member NoteProperty -Name "ItemCount" -Value "N/A"
                    $userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value "N/A"
                    $userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize-In-MB" -Value "N/A"

                    $MbxStatsLogStartTime = Get-Date
                    try {           
                        $Stats = Get-MailboxStatistics -Identity $Mbx.DistinguishedName -WarningAction SilentlyContinue -ErrorAction Stop
                        if ($Stats) {
                            $userObj."LastLogonTime" = $Stats.LastLogonTime
							$userObj."TotalItemSize-In-MB" = $(if($null -ne $Stats.TotalItemSize -and $Stats.TotalItemSize.IsUnlimited -eq $false) {$Stats.TotalItemSize.Value.ToMB()} elseif($null -ne $Stats.TotalItemSize -and $Stats.TotalItemSize.IsUnlimited -eq $true) {$Stats.TotalItemSize.Value.ToMB()} else {"N/A"})
                            $userObj."ItemCount" = $Stats.ItemCount
                            $userObj."DeletedItemCount" = $Stats.DeletedItemCount
                            $userObj."TotalDeletedItemSize-In-MB" = $(if($null -ne $Stats.TotalDeletedItemSize -and $Stats.TotalDeletedItemSize.IsUnlimited -eq $false) {$Stats.TotalDeletedItemSize.Value.ToMB()} elseif($null -ne $Stats.TotalDeletedItemSize -and $Stats.TotalDeletedItemSize.IsUnlimited -eq $true) {$Stats.TotalDeletedItemSize.Value.ToMB()} else {"N/A"})           
                        }
                    }
                    catch {
                        $warningMessage = "An error occurred while retrieving mailbox statistics for $($Mbx.DistinguishedName)."
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage Error: $($_.Exception.Message)"
                        $userObj."LastLogonTime" = "Error"
                        $userObj."TotalItemSize-In-MB" = "Error"
                        $userObj."ItemCount" = "Error"
                        $userObj."DeletedItemCount" = "Error"
                        $userObj."TotalDeletedItemSize-In-MB" = "Error"
                    }
                   
                    $ArchiveTotalItemSizeMB = "N/A"
                    $ArchiveTotalItemCount = "N/A"
                    if ($Mbx.ArchiveGuid -ne [System.Guid]::Empty) {
                        try {           
                            $MbxArchiveStats = Get-MailboxStatistics -Identity $Mbx.DistinguishedName -Archive -WarningAction SilentlyContinue -ErrorAction Stop
                            if ($MbxArchiveStats) {
                                $ArchiveTotalItemSizeMB = $(if($null -ne $MbxArchiveStats.TotalItemSize -and $MbxArchiveStats.TotalItemSize.IsUnlimited -eq $false) {$MbxArchiveStats.TotalItemSize.Value.ToMB()} elseif ($null -ne $MbxArchiveStats.TotalItemSize -and $MbxArchiveStats.TotalItemSize.IsUnlimited -eq $true) {"Unlimited"} else {"N/A"})
                                $ArchiveTotalItemCount = $MbxArchiveStats.ItemCount
                            }
                        }
                        catch {
                            $warningMessage = "An error occurred while retrieving archive mailbox statistics for $($Mbx.DistinguishedName)."
                            Write-Warning $warningMessage
                            Write-LogMailboxesProcessing "WARNING: $warningMessage Error: $($_.Exception.Message)"
                            $ArchiveTotalItemSizeMB = "Error"
                            $ArchiveTotalItemCount = "Error"
                        }           
                    }
                    $userObj | Add-Member NoteProperty -Name "ArchiveTotalItemSize-In-MB" -Value $ArchiveTotalItemSizeMB
                    $userObj | Add-Member NoteProperty -Name "ArchiveItemCount" -Value $ArchiveTotalItemCount   
                   
                    $LargeItemCount = "N/A" # Logic for this is not implemented
                    $LargeItemThresholdMBValue = 35 # Example threshold
                    $userObj | Add-Member NoteProperty -Name "LargeItemCount-Over-$($LargeItemThresholdMBValue)MB" -Value $LargeItemCount
                   
                    $MbxStatsLogEndTime = Get-Date
                    Write-LogMailboxesProcessing " -------- Processing time for MailboxStatistics for $($Mbx.Name): $($MbxStatsLogEndTime - $MbxStatsLogStartTime)"         
                   
                    $UserExceptions = @()
                    $AdditionalUserFilters = ""
                    if (-not $ShowAll) {
                        $UserExceptions = @(
                            'S-1-*', # Well-known SIDs often cause issues or are irrelevant
                            "*\Organization Management", "*\Domain Admins", "*\Enterprise Admins",
                            "*\Exchange Services", "*\Exchange Trusted Subsystem", "*\Exchange Servers",
                            "*\Exchange View-Only Administrators", "*\Exchange Admins",
                            "*\Managed Availability Servers", "*\Public Folder Administrators",
                            "*\Exchange Domain Servers", "*\Exchange Organization Administrators",
                            "NT AUTHORITY\*"
                        )
                    }

                    $ExceptionsRegex = '^(?!)$' # Regex that matches nothing
                    if ($UserExceptions.Count -gt 0) {
                        $ExceptionMatches = @($UserExceptions | ForEach-Object {[System.Text.RegularExpressions.Regex]::Escape($_)})
                        $ExceptionsRegex = '^(' + ($ExceptionMatches -join '|') + ')$'
                        $ExceptionsRegex = $ExceptionsRegex -replace '\\\*','.*' # Convert wildcard * to regex .*
                    }
                   
                    $MbxPermsLogStartTime = Get-Date
                    $FullAccessUsersList = @()
                    $FullAccessUserCount = 0
                    try {
                        $fullPermsRaw = Get-MailboxPermission -Identity $Mbx.DistinguishedName -ErrorAction SilentlyContinue | Where-Object { $_.AccessRights -contains 'FullAccess' -and (-not $_.IsInherited) -and (-not $_.Deny) }
                        if ($fullPermsRaw) {
                            foreach ($permEntry in $fullPermsRaw) {
                                $UserIDString = if ($null -ne $permEntry.User) {$permEntry.User.ToString()} else {"<UnknownUserSID: $($permEntry.User.SecurityIdentifier.Value)>"}
                                $UserSID = $permEntry.User.SecurityIdentifier.Value
                                $adObject = $null
                                $resolvedBy = "N/A"
                                $formattedADObjectIdentity = $UserIDString # Default if not resolved

                                if ($UserIDString -notmatch $ExceptionsRegex) {
                                    if ($ExpandGroups) {
                                        if ($Global:ADObjectCache.ContainsKey($UserSID)) {
                                            $adObject = $Global:ADObjectCache[$UserSID]
                                            if ($adObject) {
                                                $resolvedBy = "Cache (Object Found)"
                                                Write-LogMailboxesProcessing "        Found SID $UserSID ($UserIDString) in ADObjectCache."
                                            } else {
                                                $resolvedBy = "Cache (Not Found Marker)"
                                                Write-LogMailboxesProcessing "        Found SID $UserSID ($UserIDString) in ADObjectCache as 'Not Found'. Skipping AD lookups."
                                            }
                                        } else { # Not in cache, try to resolve
                                            # Attempt 1: If UserIDString is DOMAIN\User format (and not a SID string)
                                            if ($UserIDString -match '.+\\.+' -and -not ($UserIDString -like "S-1-*")) {
                                                $NetBIOSDomainName, $sAMAccountNameToFind = $UserIDString.Split('\')
                                                Write-LogMailboxesProcessing "        Attempting to resolve '$UserIDString' via NetBIOS domain '$NetBIOSDomainName' and sAMAccountName '$sAMAccountNameToFind'."
                                                $TargetDomainDNSRoot = $null
                                                if ($Global:DomainInfoCache.ContainsKey($NetBIOSDomainName)) {
                                                    $TargetDomainDNSRoot = $Global:DomainInfoCache[$NetBIOSDomainName]
                                                    Write-LogMailboxesProcessing "          Found '$NetBIOSDomainName' in DomainInfoCache. DNSRoot: $TargetDomainDNSRoot"
                                                } else {
                                                    try {
                                                        $TargetDomainObj = Get-ADDomain -Identity $NetBIOSDomainName -ErrorAction Stop
                                                        if ($TargetDomainObj) {
                                                            $TargetDomainDNSRoot = $TargetDomainObj.DNSRoot
                                                            $Global:DomainInfoCache[$NetBIOSDomainName] = $TargetDomainDNSRoot
                                                            Write-LogMailboxesProcessing "          Retrieved and cached DNSRoot for '$NetBIOSDomainName': $TargetDomainDNSRoot"
                                                        }
                                                    } catch { Write-LogMailboxesProcessing "          Error retrieving domain info for NetBIOS '$NetBIOSDomainName': $($_.Exception.Message)" }
                                                }
                                               
                                                if ($TargetDomainDNSRoot) {
                                                    try {
                                                        $adObject = Get-ADObject -Filter "SamAccountName -eq '$sAMAccountNameToFind'" -Server $TargetDomainDNSRoot -Properties ObjectClass, SamAccountName, DistinguishedName -ErrorAction Stop
                                                        if ($adObject) { $resolvedBy = "TargetedDomain ($TargetDomainDNSRoot)" }
                                                    } catch { Write-LogMailboxesProcessing "          Error resolving '$UserIDString' via targeted domain search on '$TargetDomainDNSRoot': $($_.Exception.Message)" }
                                                }
                                            }

                                            # Attempt 2: Resolve by SID via GC (if not resolved yet)
                                            if (-not $adObject) {
                                                $gcServer = $null
                                                try {
                                                    $currentForest = Get-ADForest -ErrorAction SilentlyContinue
                                                    if ($currentForest) { $gcServer = $currentForest.GlobalCatalogs | Get-Random -ErrorAction SilentlyContinue }
                                                } catch { Write-LogMailboxesProcessing "          Warning: Could not get Global Catalog server list for SID $UserSID. Error: $($_.Exception.Message)" }

                                                if ($gcServer) {
                                                    Write-LogMailboxesProcessing "        Attempting to resolve SID $UserSID for '$UserIDString' via GC: $gcServer"
                                                    try {
                                                        $adObject = Get-ADObject -Identity $UserSID -Server $gcServer -Properties ObjectClass, SamAccountName, DistinguishedName -ErrorAction Stop
                                                        if ($adObject) { $resolvedBy = "GC ($gcServer)" }
                                                    } catch { Write-LogMailboxesProcessing "          GC lookup for SID $UserSID failed or object not found. Error: $($_.Exception.Message)" }
                                                } else { Write-LogMailboxesProcessing "          Skipping GC lookup for SID $UserSID as no GC server was found/available."}
                                            }
                                           
                                            # Attempt 3: Resolve by SID via default domain (if still not resolved)
                                            if (-not $adObject) {
                                                Write-LogMailboxesProcessing "        Attempting default domain search for SID $UserSID for '$UserIDString'."
                                                try {
                                                    $adObject = Get-ADObject -Identity $UserSID -Properties ObjectClass, SamAccountName, DistinguishedName -ErrorAction Stop
                                                    if ($adObject) { $resolvedBy = "DefaultDomain" }
                                                } catch { Write-LogMailboxesProcessing "          Default domain search for SID $UserSID failed or object not found. Error: $($_.Exception.Message)" }
                                            }
                                            $Global:ADObjectCache[$UserSID] = $adObject # Cache the result (object or $null)
                                        }

                                        if ($adObject) { # If object was resolved (from cache or new lookup)
                                            # Get NetBIOS domain name for the resolved object for consistent formatting
                                            $objectDomainDistinguishedName = ($adObject.DistinguishedName -replace "^.+?(?=DC=)","") # Extract domain part of DN
                                            $objectDomainFQDN = ($objectDomainDistinguishedName -replace 'DC=','' -replace ',','.')
                                            $objectNetBIOSDomain = $null

                                            if ($Global:DomainFQDNToNetBIOSCache.ContainsKey($objectDomainFQDN)) {
                                                $objectNetBIOSDomain = $Global:DomainFQDNToNetBIOSCache[$objectDomainFQDN]
                                            } else {
                                                try {
                                                    $adDomainForObject = Get-ADDomain -Identity $objectDomainFQDN -ErrorAction Stop # Query domain by FQDN
                                                    if ($adDomainForObject) {
                                                        $objectNetBIOSDomain = $adDomainForObject.NetBIOSName
                                                        $Global:DomainFQDNToNetBIOSCache[$objectDomainFQDN] = $objectNetBIOSDomain
                                                    } else { Write-LogMailboxesProcessing "          Warning: Get-ADDomain returned null for FQDN '$objectDomainFQDN' (object $($adObject.SamAccountName))" }
                                                } catch { Write-LogMailboxesProcessing "          Warning: Could not get NetBIOS for domain FQDN '$objectDomainFQDN' of object '$($adObject.SamAccountName)': $($_.Exception.Message)" }
                                            }
                                            $formattedADObjectIdentity = if ($objectNetBIOSDomain) { "$objectNetBIOSDomain\$($adObject.SamAccountName)" } else { $adObject.SamAccountName }

                                            if ($resolvedBy -ne "Cache (Not Found Marker)") { # Avoid logging for already known 'not found'
                                                 Write-LogMailboxesProcessing "        Successfully resolved SID $UserSID ($UserIDString) to '$formattedADObjectIdentity' (DN: '$($adObject.DistinguishedName)') via $resolvedBy."
                                            }

                                            if ($adObject.ObjectClass -contains 'group') {
                                                Write-LogMailboxesProcessing "        Expanding group $formattedADObjectIdentity (DN: $($adObject.DistinguishedName)) for FullAccess on $($Mbx.Name)"
                                                $groupMembersSam = @()
                                                $groupExpansionStatusString = "$formattedADObjectIdentity (Group - "

                                                if ($Global:GroupMemberCache.ContainsKey($adObject.DistinguishedName)) {
                                                    $cachedEntry = $Global:GroupMemberCache[$adObject.DistinguishedName]
                                                    if ($cachedEntry -is [array]) { # Successfully expanded before
                                                        $groupMembersSam = $cachedEntry
                                                        Write-LogMailboxesProcessing "          Found group '$($adObject.SamAccountName)' in GroupMemberCache. Members count: $($groupMembersSam.Count)"
                                                    } elseif ($cachedEntry -is [string] -and ($cachedEntry -like "*Error Expanding Members*" -or $cachedEntry -like "*No Members Listed*")) {
                                                        $FullAccessUsersList += $cachedEntry # Add error/status from cache
                                                    }
                                                } else { # Not in cache, try to expand
                                                    $groupDomainForExpansionFQDN = $objectDomainFQDN # Use the group's own domain FQDN
                                                    try {
                                                        $groupMembersADPrincipal = $null
                                                        if ($groupDomainForExpansionFQDN) {
                                                            Write-LogMailboxesProcessing "          Targeting domain '$groupDomainForExpansionFQDN' for Get-ADGroupMember for group '$($adObject.DistinguishedName)'"
                                                            $groupMembersADPrincipal = Get-ADGroupMember -Identity $adObject.DistinguishedName -Server $groupDomainForExpansionFQDN -Recursive -ErrorAction Stop
                                                        } else {
                                                            Write-LogMailboxesProcessing "          Could not determine specific domain FQDN for group '$($adObject.DistinguishedName)'. Attempting Get-ADGroupMember using DN without specifying server."
                                                            $groupMembersADPrincipal = Get-ADGroupMember -Identity $adObject.DistinguishedName -Recursive -ErrorAction Stop
                                                        }

                                                        if ($groupMembersADPrincipal) {
                                                            $groupMembersSam = $groupMembersADPrincipal | ForEach-Object {$_.SamAccountName}
                                                            $Global:GroupMemberCache[$adObject.DistinguishedName] = $groupMembersSam # Cache successful expansion
                                                        } else { # Group has no members
                                                            $statusMsg = "$groupExpansionStatusString" + "No Members Listed)"
                                                            $FullAccessUsersList += $statusMsg
                                                            $Global:GroupMemberCache[$adObject.DistinguishedName] = $statusMsg # Cache no members status
                                                        }
                                                    } catch {
                                                        $statusMsg = "$groupExpansionStatusString" + "Error Expanding Members: $($_.Exception.Message.Split([Environment]::NewLine)[0]))"
                                                        Write-LogMailboxesProcessing "          $statusMsg"
                                                        $FullAccessUsersList += $statusMsg
                                                        $Global:GroupMemberCache[$adObject.DistinguishedName] = $statusMsg # Cache error status
                                                    }
                                                }
                                               
                                                if ($groupMembersSam.Count -gt 0) {
                                                    foreach($sam in $groupMembersSam){
                                                        $memberNetBIOSDomain = $objectNetBIOSDomain # Assume members are in the same domain as the group for formatting
                                                        $memberFormattedIdentity = if ($memberNetBIOSDomain) { "$memberNetBIOSDomain\$sam" } else { $sam }
                                                        if ($memberFormattedIdentity -notmatch $ExceptionsRegex) { # Check exception for each member
                                                            $FullAccessUsersList += $memberFormattedIdentity
                                                        }
                                                    }
                                                }
                                            } else { # Not a group, just a user
                                                $FullAccessUsersList += $formattedADObjectIdentity
                                            }
                                        } else { # adObject is $null (could not be resolved)
                                            if ($resolvedBy -ne "Cache (Not Found Marker)") { # Avoid logging for already known 'not found'
                                                Write-LogMailboxesProcessing "        SID $UserSID ($UserIDString) could not be resolved to an AD object after all attempts."
                                            }
                                            $FullAccessUsersList += "$UserIDString (SID Not Found in AD)"
                                        }
                                    } else { # ExpandGroups is $false
                                        $FullAccessUsersList += $UserIDString
                                    }
                                } # End if notmatch $ExceptionsRegex
                            }
                            $FullAccessUsersList = $FullAccessUsersList | Select-Object -Unique
                            $FullAccessUserCount = $FullAccessUsersList.Count
                        }
                        $userObj | Add-Member NoteProperty -Name "FullAccessCount" -Value $FullAccessUserCount   
                        $userObj | Add-Member NoteProperty -Name "FullAccess" -Value ($FullAccessUsersList -join ";")
                    } catch {
                        $userObj | Add-Member NoteProperty -Name "FullAccessCount" -Value "Error"
                        $userObj | Add-Member NoteProperty -Name "FullAccess" -Value "Error retrieving permissions"
                        $warningMessage = "An error occurred while retrieving MailboxPermission for mailbox $($Mbx.DistinguishedName): $($_.Exception.Message)"
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage"
                        Write-LogMailboxesProcessing "        Outer catch for MailboxPermission processing for '$UserIDString' on $($Mbx.Name): $($_.Exception.Message)"
                    }
                    $MbxPermsLogEndTime = Get-Date
                    Write-LogMailboxesProcessing " -------- Processing time for MailboxPermission for $($Mbx.Name): $($MbxPermsLogEndTime - $MbxPermsLogStartTime)"   
                } # End if ($OnlyADPermission -eq $false)
				else
				{
					$userObj | Add-Member NoteProperty -Name "PrimarySMTPaddress" -Value $(if($null -ne $Mbx.PrimarySmtpAddress){$Mbx.PrimarySmtpAddress.ToString()}else{""})
				}
				
                If ($IncludeADPermission)
                {
					try {
						$MbxADPermsLogStartTime = Get-Date   # log start time

						# Make sure these are reset for each mailbox
						$SendAsUsersList  = @()
						$SendAsUserCount  = 0

						$sendAsPermissions = Get-ADPermission -Identity $Mbx.DistinguishedName | Where-Object {
							($_.ExtendedRights -like "*Send*") -and
							($_.IsInherited -eq $false) -and
							($_.User -notlike "NT AUTHORITY\SELF") -and
							($_.User -notmatch '^S-1-5-21-')   # Ignore unresolved SIDs
						}

						if ($sendAsPermissions) {
							Write-Host "`nMailbox: $($Mbx.DisplayName)"
							foreach ($perm in $sendAsPermissions) {
								$SendAsUsersList += $perm.User
								Write-Host "  SendAs accordé à : $($perm.User)"
							}
							$SendAsUsersList = $SendAsUsersList | Select-Object -Unique
							$SendAsUserCount = $SendAsUsersList.Count
						}

						$userObj | Add-Member NoteProperty -Name "SendAsCount" -Value $SendAsUserCount -Force
						$userObj | Add-Member NoteProperty -Name "SendAs"      -Value ($SendAsUsersList -join ";") -Force
					}
					catch {
						$userObj | Add-Member NoteProperty -Name "SendAsCount" -Value "Error" -Force
						$userObj | Add-Member NoteProperty -Name "SendAs"      -Value "Error retrieving SendAs" -Force
						$warningMessage = "An error occurred while retrieving ADPermission (SendAs) for mailbox $($Mbx.DistinguishedName): $($_.Exception.Message)"
						Write-Warning $warningMessage
						Write-LogMailboxesProcessing "WARNING: $warningMessage"
					}
                    $MbxADPermsLogEndTime = Get-Date
                    Write-LogMailboxesProcessing " -------- Processing time for ADPermission (SendAs) for $($Mbx.Name): $($MbxADPermsLogEndTime - $MbxADPermsLogStartTime)"               
                }
                else
                {
                    $userObj | Add-Member NoteProperty -Name "SendAsCount" -Value "NotChecked"
                    $userObj | Add-Member NoteProperty -Name "SendAs" -Value "NotChecked"
                }       

                if ($OnlyADPermission -eq $false)
                {                   
                    $MbxMobileLogStartTime = Get-Date     
                    $MobileDevicePresent = $false
                    $ClientTypes = @()
                    $FriendlyNames = @()
                    $DeviceTypes = @()
                    $LastSyncTimes = @()
                    $MobileDeviceError = ""

                    try {
                        $MobileDeviceIdentity = $Mbx.UserPrincipalName
                        if (-not $MobileDeviceIdentity -and $Mbx.SamAccountName) {
                            $MobileDeviceIdentity = $Mbx.SamAccountName
                        } elseif (-not $MobileDeviceIdentity) {
                            $MobileDeviceIdentity = $Mbx.DistinguishedName # Fallback, though UPN/SAM is preferred for Get-MobileDevice
                        }

                        if ($MobileDeviceIdentity) {
                            $MobileDevices = Get-MobileDevice -Mailbox $MobileDeviceIdentity -ErrorAction SilentlyContinue
                            if ($MobileDevices) {
                                 $MobileDevicePresent = $true
                                 $ClientTypes = $MobileDevices.ClientType | ForEach-Object {if($null -ne $_){$_.ToString()}else{""}}
                                 $FriendlyNames = $MobileDevices.FriendlyName
                                 $DeviceTypes = $MobileDevices.DeviceType
                                 $LastSyncTimes = $MobileDevices.LastSyncTime | ForEach-Object {if($null -ne $_){$_}else{""}} # Keep as datetime objects for now, will be stringified by join
                            }
                        } else {
                            $warningMessage = "Could not determine a suitable identity for Get-MobileDevice for $($Mbx.Name)."
                            Write-Warning $warningMessage
                            Write-LogMailboxesProcessing "WARNING: $warningMessage"
                            $MobileDeviceError = "No Identity"
                        }
                    } catch {
                        $warningMessage = "An error occurred while retrieving MobileDevice for mailbox $($Mbx.DistinguishedName): $($_.Exception.Message)"
                        Write-Warning $warningMessage
                        Write-LogMailboxesProcessing "WARNING: $warningMessage"
                        $MobileDeviceError = "Error"
                    }
                    $userObj | Add-Member NoteProperty -Name "MobileDeviceAssociated" -Value $(if($MobileDeviceError){$MobileDeviceError} else {$MobileDevicePresent})
                    $userObj | Add-Member NoteProperty -Name "MobileClientTypes" -Value ($ClientTypes -join ";")
                    $userObj | Add-Member NoteProperty -Name "MobileFriendlyNames" -Value ($FriendlyNames -join ";")
                    $userObj | Add-Member NoteProperty -Name "MobileDeviceTypes" -Value ($DeviceTypes -join ";")
                    $userObj | Add-Member NoteProperty -Name "MobileLastSyncTimes" -Value ($LastSyncTimes -join ";")
					
					$userObj | Add-Member NoteProperty -Name "RemoteRoutingAddress" -Value $Mbx.RemoteRoutingAddress
                   
                    $MbxMobileLogEndTime = Get-Date
                    Write-LogMailboxesProcessing " -------- Processing time for MobileDevice for $($Mbx.Name): $($MbxMobileLogEndTime - $MbxMobileLogStartTime)"    
                }
               
                # ObjectGUID from AD (via Exchange Mailbox object property - no extra AD query needed)
                $userObj | Add-Member NoteProperty -Name "ObjectGUID" -Value $(if ($null -ne $Mbx.Guid) { $Mbx.Guid.ToString() } else { "" })

                # Add to the local collection for this function's scope
                $output += $userObj
               
                # Add to domain-specific collection (for per-domain CSVs if !DetectAllDomains and multiple paths)
                $MailboxesByDomain[$Domain] += $userObj
               
                $MbxEndTime = Get-Date
                $MbxTotalTimeTaken = $MbxEndTime - $MbxStartTime
                Write-LogMailboxesProcessing " ------ Total processing time for mailbox $($Mbx.Name): $($MbxTotalTimeTaken.ToString())"

                if ($totalMailBox -gt 0) {
                    $TimeElapsed = (Get-Date) - $MailboxScanStartTime
                    $AvgTimePerMailbox = if ($i -gt 0) { $TimeElapsed.TotalSeconds / $i } else { 0 }
                    $RemainingMailboxes = $totalMailBox - $i
                    $EstimatedTimeRemainingSeconds = if ($AvgTimePerMailbox -gt 0) { $RemainingMailboxes * $AvgTimePerMailbox } else { 0 }
                    $TimeSpan = [TimeSpan]::FromSeconds($EstimatedTimeRemainingSeconds)
                    $EstimatedTimeRemaining = "{0:d2}:{1:d2}:{2:d2}" -f $TimeSpan.Hours, $TimeSpan.Minutes, $TimeSpan.Seconds
                    Write-Progress -Activity "Processing mailbox $($Mbx.Name) - ($($Domain))" -Status "Scanned: $i of $totalMailBox ($([int](($i / $totalMailBox) * 100))%). Estimated time remaining: $EstimatedTimeRemaining" -PercentComplete (($i / $totalMailBox) * 100)
                }
            }
            Write-Progress -Activity "Mailbox Processing" -Status "Completed." -Completed
            Write-LogMailboxesProcessing "END - Detailed processing of mailboxes."
            Write-Host "`nEND - Detailed processing of mailboxes ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"    
            Write-LogMailboxesProcessing "DEBUG: Entering per-path/domain CSV export logic."
            Write-LogMailboxesProcessing "DEBUG: IncludedLDAPPaths = $($IncludedLDAPPaths)"
            Write-LogMailboxesProcessing "DEBUG: IncludedLDAPPaths.Count = $($IncludedLDAPPaths.Count)"
            Write-LogMailboxesProcessing "DEBUG: DetectAllDomains = $($DetectAllDomains)"

            if ((-not $DetectAllDomains -and $IncludedLDAPPaths -and $IncludedLDAPPaths.Count -gt 0)) {
                Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
                Write-LogMailboxesProcessing "Starting export of per-path/per-OU data to CSVs (as -DetectAllDomains is false)."
                Write-Host "Exporting per-path/per-OU data to CSVs ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
               
                if ($IncludedLDAPPaths.Count -eq 1) { # Single OU/Path
                    $singlePathForCsv = $IncludedLDAPPaths[0]
                    $csvIdentifierForSinglePath = $logPathIdentifier # Use the already generated logPathIdentifier
                    $OutputDataForPath = $output # All output from this run corresponds to this single path

                    if ($OutputDataForPath -and $OutputDataForPath.Count -gt 0) {
                        $FileNameSuffix = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
        $PerPathCsvFile = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_$($csvIdentifierForSinglePath)${FileNameSuffix}"
                        try {
                            Export-CsvAtomic -InputObject $OutputDataForPath -Path $PerPathCsvFile -Encoding UTF8
                            Write-Host -ForegroundColor:Green "Mailbox data for path '$($singlePathForCsv)' (identified as '$csvIdentifierForSinglePath') exported to: $PerPathCsvFile"
                            Write-LogMailboxesProcessing "Mailbox data for path '$($singlePathForCsv)' exported to: $PerPathCsvFile"
                        } catch {
                            $errorMessage = "Failed to export CSV for path '$($singlePathForCsv)': $($_.Exception.Message)"
                            Write-Error $errorMessage
                            Write-LogMailboxesProcessing "ERROR: $errorMessage"
                        }
                    } else {
                        Write-LogMailboxesProcessing "No data to export for path '$singlePathForCsv'."
                    }
                } elseif ($IncludedLDAPPaths.Count -gt 1) { # Multiple OUs/Paths
                    foreach ($DomainToExportInOUContext in $Domains) { 
                        $OutputDataForDomainInOU = $MailboxesByDomain[$DomainToExportInOUContext]
                        if ($OutputDataForDomainInOU -and $OutputDataForDomainInOU.Count -gt 0) {
                            $FileNameSuffix = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
                $csvBaseNameForOUContext = "Exchange_OnPrem_Mailboxes_TargetedScope_Domain_$($DomainToExportInOUContext)"
                            $PerDomainCsvFileInOUContext = Join-Path -Path $OutputPath -ChildPath "${csvBaseNameForOUContext}${FileNameSuffix}"
                            
                            if ($ForceOverwriteCSV -and (Test-Path $PerDomainCsvFileInOUContext)) {
                                Write-LogMailboxesProcessing "INFO: Per-scope/domain CSV file '$PerDomainCsvFileInOUContext' exists and ForceOverwriteCSV is enabled. Deleting file."
                                try { Remove-Item -Path $PerDomainCsvFileInOUContext -Force -ErrorAction Stop } 
                                catch { Write-LogMailboxesProcessing "ERROR: Could not delete existing per-scope/domain CSV file '$PerDomainCsvFileInOUContext'. Message: $($_.Exception.Message)" }
                            } elseif ((Test-Path $PerDomainCsvFileInOUContext) -and (-not $ForceOverwriteCSV)) {
                                Write-LogMailboxesProcessing "WARNING: CSV file '$PerDomainCsvFileInOUContext' exists and -ForceOverwriteCSV is false. Skipping export for this specific domain context within targeted OUs."
                                Write-Host -ForegroundColor Yellow "CSV file '$PerDomainCsvFileInOUContext' exists and -ForceOverwriteCSV is false. Skipping export for this domain context."
                                continue # Skip to next domain context
                            }

                            try {
                                Export-CsvAtomic -InputObject $OutputDataForDomainInOU -Path $PerDomainCsvFileInOUContext -Encoding UTF8
                                Write-Host -ForegroundColor:Green "Mailbox data for domain scope '$($DomainToExportInOUContext)' (from targeted OUs) exported to: $PerDomainCsvFileInOUContext"
                                Write-LogMailboxesProcessing "Mailbox data for domain scope '$($DomainToExportInOUContext)' (from targeted OUs) exported to: $PerDomainCsvFileInOUContext"
                            } catch {
                                $errorMessage = "Failed to export CSV for domain scope '$($DomainToExportInOUContext)' (from targeted OUs): $($_.Exception.Message)"
                                Write-Error $errorMessage
                                Write-LogMailboxesProcessing "ERROR: $errorMessage"
                            }
                        } else {
                            Write-LogMailboxesProcessing "No data to export for domain scope '$DomainToExportInOUContext' (from targeted OUs)."
                        }
                    }
                }
                Write-Host "END - Exporting per-path/per-OU data to CSVs ... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
                Write-LogMailboxesProcessing "END - Exporting per-path/per-OU data to CSVs."
            }
        } # End else (if $totalMailbox -gt 0)

        Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
        $FunctionEndTime = Get-Date
        $FunctionTotalTimeTaken = $null

        if ($MailboxScanStartTime) { # If scan actually started
            $FunctionTotalTimeTaken = $FunctionEndTime - $MailboxScanStartTime
            if ($totalMailBox -gt 0) {
                $TotalSeconds = $FunctionTotalTimeTaken.TotalSeconds
                $AverageSecondsPerMailbox = $TotalSeconds / $totalMailBox
                $AverageTimePerMailbox = [TimeSpan]::FromSeconds($AverageSecondsPerMailbox)
                Write-Host "MailboxesProcessing2: Total processing time for $totalMailBox mailboxes: $($FunctionTotalTimeTaken.ToString())"
                Write-Host "MailboxesProcessing2: Average processing time per mailbox: $($AverageTimePerMailbox.ToString())"
                Write-LogMailboxesProcessing "MailboxesProcessing2: Total processing time for $totalMailBox mailboxes: $($FunctionTotalTimeTaken.ToString())"
                Write-LogMailboxesProcessing "MailboxesProcessing2: Average processing time per mailbox: $($AverageTimePerMailbox.ToString())"
            } else { # Scan started but no mailboxes found
                Write-LogMailboxesProcessing "MailboxesProcessing2: No mailboxes processed, cannot calculate average time."
                Write-Host "MailboxesProcessing2: Total execution time (scan started, no mailboxes processed): $($FunctionTotalTimeTaken.ToString())"
                Write-LogMailboxesProcessing "MailboxesProcessing2: Total execution time (scan started, no mailboxes processed): $($FunctionTotalTimeTaken.ToString())"
            }
        } else { # Mailbox scan phase was skipped (e.g., no mailboxes initially found)
             $ApproximateStartTime = $FunctionStartTime # Use function start time as best guess
             if (Test-Path $Script:MailboxesProcessingLogFile) { # Use script-scoped variable
                 try {
                     $ApproximateStartTime = (Get-Item $Script:MailboxesProcessingLogFile -ErrorAction Stop).CreationTime
                 } catch {
                     Write-LogMailboxesProcessing "Warning: Could not get CreationTime of MailboxesProcessingLogFile. Using function start time for duration calculation."
                 }
             }
             $FunctionTotalTimeTaken = $FunctionEndTime - $ApproximateStartTime
             Write-Host "MailboxesProcessing2: Total execution time (no mailbox scan phase or scan start time unavailable): $($FunctionTotalTimeTaken.ToString())"
             Write-LogMailboxesProcessing "MailboxesProcessing2: Total execution time (no mailbox scan phase or scan start time unavailable): $($FunctionTotalTimeTaken.ToString())"
        }
        return $output # Return the collected data for this scope
    } # End of MailboxesProcessing2 function

    # Wrapper function that calls MailboxesProcessing2
    function MailboxesProcessing {
        param (
            [Parameter(Mandatory=$true)]
            [string[]]$IncludedLDAPPaths
        )
        process {
            Write-Host "`n  [MailboxesProcessing] Function called." -ForegroundColor Magenta
            WriteLog -Message "  [MailboxesProcessing] Wrapper function called. Will call MailboxesProcessing2."
            WriteLog -Message "  [MailboxesProcessing] Effective permission flags for this call: IncludeADPermission = $IncludeADPermission, OnlyADPermission = $OnlyADPermission"

            # MailboxesProcessing2 will return its collected data
            $processedData = MailboxesProcessing2 -IncludedLDAPPaths $IncludedLDAPPaths
           
            WriteLog -Message "  [MailboxesProcessing] Finished call to MailboxesProcessing2."
            return $processedData # Pass the data up
        }
    }

    # Function to process a specific domain when DetectAllDomains is $true
    function Process-SpecificDomain {
        param (
            [Parameter(Mandatory=$true)]
            [System.DirectoryServices.ActiveDirectory.Domain]$CurrentDomain
        )
        process {
            $domainName = $CurrentDomain.Name
            $distinguishedName = ($domainName.Split('.') | ForEach-Object { "DC=$_" }) -join ','
           
            WriteLog -Message "Starting global processing for domain: '$domainName' (DN: '$distinguishedName')"

            $isForestRootDomain = $false
            if ($CurrentDomain.Forest -and $CurrentDomain.Forest.RootDomain -and ($CurrentDomain.Name -eq $CurrentDomain.Forest.RootDomain.Name)) {
                $isForestRootDomain = $true
                WriteLog -Message "  -> Domain '$domainName' IS the AD forest root."
            } else {
                if ($CurrentDomain.Forest -and $CurrentDomain.Forest.RootDomain) {
                     WriteLog -Message "  -> Domain '$domainName' IS NOT the AD forest root (The forest root is '$($CurrentDomain.Forest.RootDomain.Name)')."
                } else {
                     WriteLog -Message "  -> Domain '$domainName' IS NOT the AD forest root (Could not confirm the current forest root via this specific domain object)."
                }
            }
           
            Write-Host -ForegroundColor:Cyan "Processing domain '$domainName' $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Write-Host "  LDAP DN: $distinguishedName"
           
            # --- START Per-Domain CSV Overwrite/Skip/Load Check for DetectAllDomains mode ---
    $perDomainCsvBaseName = "Exchange_OnPrem_Mailboxes_$($domainName)"
            $perDomainCsvSuffix = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
            $perDomainCsvFullPath = Join-Path -Path $OutputPath -ChildPath "$($perDomainCsvBaseName)${perDomainCsvSuffix}"

            if (Test-Path $perDomainCsvFullPath) {
                if (-not $ForceOverwriteCSV) {
                    # MODIFICATION START: Load data from existing CSV instead of just skipping
                    $loadMessage = "Le fichier CSV par domaine '$perDomainCsvFullPath' pour le domaine '$domainName' existe déjà et -ForceOverwriteCSV est `$false. Tentative de chargement des données depuis ce fichier."
                    WriteLog -Message "INFO: $loadMessage"
                    Write-Host -ForegroundColor Cyan $loadMessage

                    try {
                        $existingData = Import-Csv -Path $perDomainCsvFullPath -Encoding UTF8 -ErrorAction Stop
                        if ($existingData -and $existingData.Count -gt 0) {
                            $Global:ScriptOverallMailboxData += $existingData
                            $successMessage = "Chargement réussi de $($existingData.Count) enregistrements depuis '$perDomainCsvFullPath' et ajoutés aux données globales. Le traitement en direct pour le domaine '$domainName' sera ignoré."
                            WriteLog -Message "INFO: $successMessage"
                            Write-Host -ForegroundColor Green $successMessage
                        } elseif ($existingData) { # File exists but is empty
                             $emptyFileMessage = "Le fichier CSV '$perDomainCsvFullPath' existe mais est vide. Aucun enregistrement n'a été chargé. Le traitement en direct pour le domaine '$domainName' sera ignoré."
                             WriteLog -Message "INFO: $emptyFileMessage" # Changed from WARNING to INFO as it's an expected scenario
                             Write-Host -ForegroundColor Cyan $emptyFileMessage
                        } else { # Should not happen if Import-Csv did not error but returned $null for $existingData
                            $nullDataMessage = "L'importation depuis '$perDomainCsvFullPath' n'a retourné aucune donnée (possiblement un fichier vide ou mal formé non détecté comme erreur). Le traitement en direct pour le domaine '$domainName' sera ignoré."
                            WriteLog -Message "WARNING: $nullDataMessage"
                            Write-Host -ForegroundColor Yellow $nullDataMessage
                        }
                    } catch {
                        $importErrorMessage = "ERREUR: Échec de l'importation des données depuis le CSV existant '$perDomainCsvFullPath' pour le domaine '$domainName'. Message: $($_.Exception.Message). Le traitement en direct pour ce domaine sera ignoré."
                        WriteLog -message $importErrorMessage
                        Write-Error $importErrorMessage # Keep as Write-Error for visibility
                    }
                    return # Skip live processing for this domain as data is loaded or attempt was made
                    # MODIFICATION END
                } else { # ForceOverwriteCSV is $true
                    $message = "Per-domain CSV file '$perDomainCsvFullPath' for domain '$domainName' exists and -ForceOverwriteCSV is `$true. Existing file will be deleted before processing."
                    WriteLog -Message "INFO: $message"
                    Write-Host -ForegroundColor Cyan $message
                    try {
                        Remove-Item -Path $perDomainCsvFullPath -Force -ErrorAction Stop
                        WriteLog -Message "INFO: File '$perDomainCsvFullPath' for domain '$domainName' deleted successfully."
                    } catch {
                        $errorMessage = "ERROR: Could not delete existing per-domain CSV file '$perDomainCsvFullPath' for domain '$domainName'. Message: $($_.Exception.Message)"
                        WriteLog -Message "ERROR: $errorMessage"
                        Write-Error $errorMessage
                        # Decide if script should stop or continue if deletion fails. Original script implies continuation.
                    }
                }
            }
            # --- END Per-Domain CSV Overwrite/Skip/Load Check ---

            [string[]]$pathsForMailboxProcessing = @()
            $domainDataFromProcessing = $null # Initialize to null

            if ($isForestRootDomain) {
                Write-Host "  Domain Status: AD Forest Root." -ForegroundColor Yellow
                WriteLog -Message "  Domain '$domainName' is the forest root. Attempting to retrieve first-level OUs."
               
                try {
                    $firstLevelOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $distinguishedName -SearchScope OneLevel -Server $domainName -ErrorAction Stop | Select-Object -ExpandProperty DistinguishedName
                   
                    if ($firstLevelOUs -and $firstLevelOUs.Count -gt 0) {
                        $pathsForMailboxProcessing = $firstLevelOUs
                        Write-Host "    -> $($firstLevelOUs.Count) first-level OUs found under '$distinguishedName'. They will be passed to MailboxesProcessing."
                        $firstLevelOUs | ForEach-Object { WriteLog -Message "      -> First-level OU to include for MailboxesProcessing: '$_'" }
                    } else {
                        WriteLog -Message "  WARNING: No first-level OUs found under the root domain '$domainName'. Processing the domain root ('$distinguishedName') itself."
                        Write-Host "  WARNING: No first-level OUs found under the root domain '$domainName'. Processing the domain root itself." -ForegroundColor Yellow
                        $pathsForMailboxProcessing = @($distinguishedName)
                    }
                } catch {
                    # If Get-ADOrganizationalUnit fails, fall back to processing the domain root itself.
                    $errorMessage = "ERROR retrieving first-level OUs for the root domain '$domainName'. Message: $($_.Exception.Message). Will attempt to process the domain root ('$distinguishedName') itself."
                    WriteLog -message $errorMessage
                    Write-Warning $errorMessage
                    $pathsForMailboxProcessing = @($distinguishedName)
                }
                $domainDataFromProcessing = MailboxesProcessing -IncludedLDAPPaths $pathsForMailboxProcessing
                if ($null -ne $domainDataFromProcessing) { $Global:ScriptOverallMailboxData += $domainDataFromProcessing }


            } else { # Not the forest root domain
                Write-Host "  Domain Status: Not the AD Forest Root."
                WriteLog -Message "  Domain '$domainName' is not the forest root. The domain's own DN ('$distinguishedName') will be used for MailboxesProcessing."
                $pathsForMailboxProcessing = @($distinguishedName)
                $domainDataFromProcessing = MailboxesProcessing -IncludedLDAPPaths $pathsForMailboxProcessing
                if ($null -ne $domainDataFromProcessing) { $Global:ScriptOverallMailboxData += $domainDataFromProcessing }
            }
           
            # Export data for THIS specific domain if it was processed live (not loaded from existing CSV)
            # This is the per-domain CSV that might be loaded in future runs.
            if ($domainDataFromProcessing -and $domainDataFromProcessing.Count -gt 0) {
                try {
                    Export-CsvAtomic -InputObject $domainDataFromProcessing -Path $perDomainCsvFullPath -Encoding UTF8
                    WriteLog -Message "INFO: Live processed data for domain '$domainName' exported to '$perDomainCsvFullPath'."
                    Write-Host -ForegroundColor Green "Live processed data for domain '$domainName' exported to '$perDomainCsvFullPath'."
                } catch {
                    $exportError = "ERROR: Failed to export live processed data for domain '$domainName' to '$perDomainCsvFullPath'. Message: $($_.Exception.Message)"
                    WriteLog -message $exportError
                    Write-Error $exportError
                }
            } elseif ($domainDataFromProcessing) { # Processed live, but no mailboxes found
                 WriteLog -Message "INFO: Live processing for domain '$domainName' yielded no mailboxes. Per-domain CSV '$perDomainCsvFullPath' will not be created or will be empty if pre-deleted."
            }
            # If data was loaded from existing CSV, $domainDataFromProcessing would be $null here.

            Write-Host "  -> Other processing logic for domain '$domainName' (outside MailboxesProcessing) completed."
            WriteLog -Message "Finished global processing for domain: '$domainName'. IsForestRoot: $isForestRootDomain"
        }
    }
    #endregion Function Definitions

    #region Main Script Block
    $InventoryCompletedSuccessfully = $false # Initialize completion flag
    try { # Main try block for script execution
        if ($DetectAllDomains)
        {
            Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
            WriteLog -Message "DetectAllDomains mode enabled. Checking for Active Directory module..."
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                $errorMessage = "CRITICAL ERROR: The Active Directory module is not installed. Please install it before running this script. The script will stop."
                WriteLog -message $errorMessage				
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				SendEmailHtmlReport -BodyHtml $body
                throw $errorMessage
            }
            WriteLog -Message "Importing Active Directory module..."
            Write-Host "Importing Active Directory module... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
            Import-Module ActiveDirectory -ErrorAction Stop
           
            WriteLog -Message "Retrieving forest information..."
            $forest = $null
            try {
                $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            } catch {
                $errorMessage = "CRITICAL ERROR: Failed to contact Active Directory forest. Check connectivity and permissions. Message: $($_.Exception.Message). The script will stop."
                WriteLog -message $errorMessage
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				SendEmailHtmlReport -BodyHtml $body
                throw $errorMessage
            }

            if (-not $forest) {
                $errorMessage = "CRITICAL ERROR: Failed to retrieve current Active Directory forest information (forest object is null). The script will stop."
                WriteLog -message $errorMessage
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				SendEmailHtmlReport -BodyHtml $body
                throw $errorMessage
            }
            WriteLog -Message "Current forest: $($forest.Name)"

            if (-not $forest.RootDomain) {
                $errorMessage = "CRITICAL ERROR: Failed to determine the root domain of the Active Directory forest '$($forest.Name)'. The script will stop."
                WriteLog -message $errorMessage
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				SendEmailHtmlReport -BodyHtml $body
                throw $errorMessage
            }
            WriteLog -Message "Detected forest root domain: $($forest.RootDomain.Name)"
			
			# --- START Check for existing COMBINED CSV file if DetectAllDomains is $true ---
			$combinedCsvFileSuffixGlobal = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
    $globalCombinedCsvFile = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_AllDomains${combinedCsvFileSuffixGlobal}"

			$baseExchangeMailboxesInventoryPath = (Get-Item $OutputPath).Parent.Parent.FullName # Go up two levels to ExchangeMailboxesInventory
			$backupBaseDir = Join-Path -Path $baseExchangeMailboxesInventoryPath -ChildPath "Backup"

			if (Test-Path $globalCombinedCsvFile) {
				if (-not $ForceOverwriteCSV) {
					# Create a timestamped directory inside the backup base directory
					$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
					$currentBackupDir = Join-Path -Path $backupBaseDir -ChildPath $timestamp

					# Create the full backup path including the timestamped directory
					if (-not (Test-Path $currentBackupDir)) {
						New-Item -ItemType Directory -Path $currentBackupDir -Force | Out-Null
						WriteLog -Message "INFO: Created timestamped backup directory: '$currentBackupDir'"
					}

					# The backup file name should be the same as the source file name
					$backupFileName = Split-Path -Path $globalCombinedCsvFile -Leaf
					$backupFilePath = Join-Path -Path $currentBackupDir -ChildPath $backupFileName

					WriteLog -Message "INFO: Combined CSV file '$globalCombinedCsvFile' already exists and -ForceOverwriteCSV is `$false. Backing up the existing file to '$backupFilePath'."
					Write-Host -ForegroundColor Yellow "WARNING: Combined CSV file '$globalCombinedCsvFile' already exists. Backing it up to '$backupFilePath' before continuing."

					# Copy the file to the backup directory
					try {
						# Copy the file
						Copy-Item -Path $globalCombinedCsvFile -Destination $backupFilePath -Force -ErrorAction Stop
						WriteLog -Message "INFO: Successfully backed up '$globalCombinedCsvFile' to '$backupFilePath'."

						# After successful backup, remove the original file
						Remove-Item -Path $globalCombinedCsvFile -ErrorAction Stop
						WriteLog -Message "INFO: Successfully removed original combined CSV file '$globalCombinedCsvFile'."

					} catch {
						# Handle errors during copy or removal
						$errorMessage = "CRITICAL ERROR: Failed to process the combined CSV file '$globalCombinedCsvFile' (copy or removal failed). Error: $($_.Exception.Message)"
						WriteLog -message $errorMessage
						Write-Host -ForegroundColor Red $errorMessage
						$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
						SendEmailHtmlReport -BodyHtml $body
						exit 1 # Stop script if processing fails
					}
				} else {
					WriteLog -Message "INFO: Combined CSV file '$globalCombinedCsvFile' exists and -ForceOverwriteCSV is `$true. It will be deleted before new combined export at the end of processing."
					# Deletion will happen just before the final export of the combined file.
				}
			}
            $domainsToProcess = $forest.Domains
            if (-not $domainsToProcess -or $domainsToProcess.Count -eq 0) {
                $errorMessage = "CRITICAL ERROR: No domains were found in the forest '$($forest.Name)'. Check Active Directory configuration. The script will stop."
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				SendEmailHtmlReport -BodyHtml $body
                WriteLog -message $errorMessage
                throw $errorMessage
            }

            WriteLog -Message "Number of domains to process: $($domainsToProcess.Count)"

            if ($domainsToProcess.Count -gt 0) {
                Write-Host "`nList of domains to process:" -ForegroundColor Green
                $domainsToProcess | Select-Object Name, @{Name="ExistingCSV"; Expression={
                    $domainCsvBaseName = "Exchange_OnPrem_Mailboxes_$($_.Name)"
                    $domainCsvSuffix = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }
                    $domainCsvFullPath = Join-Path -Path $OutputPath -ChildPath "$($domainCsvBaseName)${domainCsvSuffix}"
                    if (Test-Path $domainCsvFullPath) {"Yes"} else {"No"}
                }} | Format-Table -AutoSize
                Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
            }

            WriteLog -Message "Starting processing of each domain via the Process-SpecificDomain function..."
            foreach ($domain in $domainsToProcess) {
                Process-SpecificDomain -CurrentDomain $domain
            }
            WriteLog -Message "Finished processing all domains."
            # Export the globally accumulated data to the AllDomains CSV
            if ($Global:ScriptOverallMailboxData.Count -gt 0) {
                WriteLog -Message "Exporting combined data for all processed domains to '$globalCombinedCsvFile'..."
                if ((Test-Path $globalCombinedCsvFile) -and $ForceOverwriteCSV) { # File might have been created by a previous version or if script was interrupted
                    try {
                        Remove-Item -Path $globalCombinedCsvFile -Force -ErrorAction Stop
                        WriteLog -Message "INFO: Successfully deleted existing combined CSV file '$globalCombinedCsvFile' due to ForceOverwriteCSV."
                    } catch {
                        WriteLog -Message "ERROR: Failed to delete existing combined CSV file '$globalCombinedCsvFile'. Export may fail or append. Error: $($_.Exception.Message)"
                    }
                }
                try {
                    Export-CsvAtomic -InputObject $Global:ScriptOverallMailboxData -Path $globalCombinedCsvFile -Encoding UTF8
                    WriteLog -Message "Successfully exported combined data to '$globalCombinedCsvFile'."
                    Write-Host -ForegroundColor Green "All processed mailbox data exported to: $globalCombinedCsvFile"
					$InputCsvForDuplicateScan = $globalCombinedCsvFile
					$scriptdatamailbox = $true						
					$SendFileListEmailReportFileName = $globalCombinedCsvFile
                } catch {
                    $errorMessage = "Failed to export combined data to '$globalCombinedCsvFile': $($_.Exception.Message)"
                    WriteLog -Message "ERROR: $errorMessage"					
                    Write-Error $errorMessage
					$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Failed to export combined data to '$globalCombinedCsvFile': $($_.Exception.Message)"
					SendEmailHtmlReport -BodyHtml $body
                }
            } else {
                WriteLog -Message "No data accumulated in \$Global:ScriptOverallMailboxData. Combined 'AllDomains' CSV will not be created or will be empty if it was pre-deleted."
                Write-Host -ForegroundColor Yellow "No mailbox data was collected or loaded. The combined file '$globalCombinedCsvFile' will not be created or will be empty."
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : No mailbox data was collected or loaded. The combined file '$globalCombinedCsvFile' will not be created or will be empty."
				SendEmailHtmlReport -BodyHtml $body
            }

        }
        else # DetectAllDomains is $false
        {
            WriteLog -Message "DetectAllDomains mode is $false."
            Write-Host "DetectAllDomains mode is $false." -ForegroundColor Green
           
            $combinedCsvFileForNonDetectAll = $null
            $createCombinedCsvForNonDetectAll = $true # Assume we create a combined CSV unless it's a single OU
            $combinedCsvFileSuffixLocal = if ($OnlyADPermission) { "_OnlyADPermission.csv" } else { ".csv" }

            if ($IncludedOrganizationalUnit -and $IncludedOrganizationalUnit.Count -eq 1) {
                $createCombinedCsvForNonDetectAll = $false # For single OU, MailboxesProcessing2 handles its own CSV. No separate "combined" for one.
                WriteLog -Message "INFO: Combined CSV will not be created for non-DetectAllDomains mode as only one specific OU/path is processed. Its specific CSV will be generated by MailboxesProcessing2 if data is found."
            } elseif ($IncludedOrganizationalUnit -and $IncludedOrganizationalUnit.Count -gt 1) {
                $pathIdentifiers = foreach ($path in $IncludedOrganizationalUnit) {
                    $firstPart = ($path -split ',')[0] -replace '^(OU=|CN=|DC=)','' -replace '[^a-zA-Z0-9_-]',''
                    if ($firstPart.Length -gt 15) { $firstPart = $firstPart.Substring(0,15) }
                    $firstPart
                }
                $identifierString = ($pathIdentifiers | Select-Object -Unique) -join "_"
                if ($identifierString.Length -gt 50) { $identifierString = $identifierString.Substring(0,50)}
                $combinedCsvFileForNonDetectAll = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_AllTargetedOUs_${identifierString}${combinedCsvFileSuffixLocal}"
            } else { # No OUs specified, processing current scope
                $combinedCsvFileForNonDetectAll = Join-Path -Path $OutputPath -ChildPath "Exchange_OnPrem_Mailboxes_AllCurrentScope${combinedCsvFileSuffixLocal}"
            }

            if ($createCombinedCsvForNonDetectAll -and $combinedCsvFileForNonDetectAll) {
                if (Test-Path $combinedCsvFileForNonDetectAll) {
                    if (-not $ForceOverwriteCSV) {
                        $errorMessage = "CRITICAL ERROR: Combined CSV file '$combinedCsvFileForNonDetectAll' for specified OUs/scope already exists and -ForceOverwriteCSV is `$false. Script will stop."
                        WriteLog -message $errorMessage
                        Write-Host -ForegroundColor Red $errorMessage
						$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
						SendEmailHtmlReport -BodyHtml $body
                        exit 1
                    } else {
                        WriteLog -Message "INFO: Combined CSV file '$combinedCsvFileForNonDetectAll' exists and -ForceOverwriteCSV is `$true. It will be deleted."
                        try { Remove-Item -Path $combinedCsvFileForNonDetectAll -Force -ErrorAction Stop }
                        catch { WriteLog -Message "ERROR: Could not delete '$combinedCsvFileForNonDetectAll'. Message: $($_.Exception.Message)" }
                    }
                }
            }

            if ($IncludedOrganizationalUnit -and $IncludedOrganizationalUnit.Count -gt 0) {
                WriteLog -Message "Processing specified Organizational Units. $($IncludedOrganizationalUnit.Count) OU(s) provided via -IncludedOrganizationalUnit parameter."
                Write-Host "Processing specified Organizational Units. $($IncludedOrganizationalUnit.Count) OU(s) provided." -ForegroundColor Green
                $IncludedOrganizationalUnit | ForEach-Object { WriteLog -Message "  - Will process OU: $_" }
            } else {
                WriteLog -Message "No Organizational Units specified via -IncludedOrganizationalUnit parameter. MailboxesProcessing2 will attempt to retrieve all mailboxes in the current Exchange scope."
                Write-Host "Warning: No Organizational Units specified via -IncludedOrganizationalUnit parameter. MailboxesProcessing2 will attempt to retrieve all mailboxes in the current Exchange scope." -ForegroundColor Yellow
            }
           
            $ouData = MailboxesProcessing -IncludedLDAPPaths $IncludedOrganizationalUnit # This will call MailboxesProcessing2
            if ($null -ne $ouData) {
                 $Global:ScriptOverallMailboxData += $ouData
            }
            # Export the combined data if applicable for !DetectAllDomains (multiple OUs or current scope)
            if ($createCombinedCsvForNonDetectAll -and $Global:ScriptOverallMailboxData.Count -gt 0 -and $combinedCsvFileForNonDetectAll) {
                WriteLog -Message "Exporting combined data for specified OUs/scope to '$combinedCsvFileForNonDetectAll'..."
                # Overwrite check already performed, or ForceOverwriteCSV is true (file deleted)
                try {
                    Export-CsvAtomic -InputObject $Global:ScriptOverallMailboxData -Path $combinedCsvFileForNonDetectAll -Encoding UTF8
                    WriteLog -Message "Successfully exported combined data to '$combinedCsvFileForNonDetectAll'."
                    Write-Host -ForegroundColor Green "All processed mailbox data for specified scope exported to: $combinedCsvFileForNonDetectAll"
					$InputCsvForDuplicateScan = $combinedCsvFileForNonDetectAll					
					$scriptdatamailbox = $true
					$SendFileListEmailReportFileName = $combinedCsvFileForNonDetectAll	
                } catch {
                    $errorMessage = "Failed to export combined data to '$combinedCsvFileForNonDetectAll': $($_.Exception.Message)"
                    WriteLog -Message "ERROR: $errorMessage"
                    Write-Error $errorMessage
                }
            } elseif ($createCombinedCsvForNonDetectAll -and (-not $Global:ScriptOverallMailboxData -or $Global:ScriptOverallMailboxData.Count -eq 0)) {
                WriteLog -Message "No data accumulated in \$Global:ScriptOverallMailboxData for specified OUs/scope. Combined CSV '$combinedCsvFileForNonDetectAll' will not be created or will be empty if pre-deleted."
                Write-Host -ForegroundColor Yellow "No mailbox data was collected for the specified OUs/scope. The combined file '$combinedCsvFileForNonDetectAll' will not be created or will be empty."
            }
        }
        WriteLog -Message "END of script main logic execution."
        Write-Host "END of script main logic execution... $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
        $InventoryCompletedSuccessfully = $true # Set flag for successful completion
    #endregion Main Script Block
    }
    finally { # This is the finally for the main try block that starts after variable initialization
        $EndTime = Get-Date
        if ($InventoryCompletedSuccessfully) {
            WriteLog -Message "Inventory finished. Total duration: $($EndTime - $StartTime)."
            Write-Host "Inventory finished. Total duration: $($EndTime - $StartTime)."
        } else { # Script was interrupted or had an unhandled terminating error
            $interruptionMessage = "Inventory interrupted or terminated due to an error. Total duration: $($EndTime - $StartTime)."
            Write-Host -ForegroundColor Yellow $interruptionMessage
        }
    }
}
finally {

if ($scriptdatamailbox -eq $true) {
	Write-Host "-----------------------------------------------------------------------------------------"
	Write-Host "Remove Duplicate in Export Exchange 2016 mailboxes inventory ..."
	WriteLog -Message "Remove Duplicate in Export Exchange 2016 mailboxes inventory ..."
	
    # Define paths
	$TempOutputCsv = [System.IO.Path]::ChangeExtension($InputCsvForDuplicateScan, $null) + "_WithoutDuplicateSMTP.csv"
    $LogFileDuplicate = $global:LogTextFile -replace '\.log$', '_duplicate.log'

    # Load CSV
    WriteLog -Message "Load CSV : $InputCsvForDuplicateScan"
    $rows = Import-Csv -Path $InputCsvForDuplicateScan

    # Create hashtable to track unique addresses (case-insensitive)
    WriteLog -Message "Create hashtable to track unique addresses"
    $seen = @{}
    $uniqueRows = @()
    $duplicates = @()

    foreach ($row in $rows) {
        $address = $row.PrimarySMTPaddress.ToLower()
        if ($seen.ContainsKey($address)) {
            $duplicates += $row
        } else {
            $seen[$address] = $true
            $uniqueRows += $row
        }
    }

    # Check if duplicates exist
    if ($duplicates.Count -eq 0) {
        WriteLog -Message "No duplicates found. CSV file remains unchanged."

        # Clean up temporary file if it exists
        if (Test-Path $TempOutputCsv) {
            try {
                Remove-Item -Path $TempOutputCsv -Force
                WriteLog -Message "Temporary file removed: $TempOutputCsv"
            } catch {
                WriteLog -Message "Failed to remove temporary file: $_" "ERROR"
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : No Duplicates Primary SMTP found - Failed to remove temporary file: $_"
				SendEmailHtmlReport -BodyHtml $body
				exit 1
            }
        }
    } else {
        WriteLog -Message "Duplicates Primary SMTP detected. Processing cleanup..."		
        # Export cleaned CSV to temporary file
        try {
            Export-CsvAtomic -InputObject $uniqueRows -Path $TempOutputCsv -Encoding UTF8
        } catch {
            WriteLog -Message "Failed to export cleaned CSV: $_" "ERROR"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Duplicates Primary SMTP detected - Failed to export cleaned CSV: $_"
			SendEmailHtmlReport -BodyHtml $body
            exit 1
        }

        # Log duplicates with timestamp
        $logEntries = @()
        foreach ($dup in $duplicates) {
            $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Duplicate found - PrimarySMTPaddress: $($dup.PrimarySMTPaddress), DistinguishedName: $($dup.DistinguishedName)"
            WriteLog -Message $entry
            $logEntries += $entry
        }
        try {
            $logEntries | Out-File -FilePath $LogFileDuplicate -Encoding UTF8
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Duplicates Primary SMTP detected."
			SendEmailHtmlReport -BodyHtml $body -Attachments $LogFileDuplicate
        } catch {			
            WriteLog -Message "Failed to write log file: $_" "ERROR"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Duplicates Primary SMTP detected - Failed to write log file: $_"
			SendEmailHtmlReport -BodyHtml $body
            exit 1
        }

        # Rename original file by appending '-versionoriginal'
        $originalName = Split-Path -Leaf $InputCsvForDuplicateScan
        $originalFolder = Split-Path -Parent $InputCsvForDuplicateScan
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($originalName)
        $extension = [System.IO.Path]::GetExtension($originalName)
        $renamedPath = Join-Path $originalFolder "$baseName-versionoriginal$extension"

        try {
            if (Test-Path $renamedPath) {
                if ($ForceOverwriteCSV) {
                    WriteLog -Message "File already exists: $renamedPath. Removing due to ForceOverwriteCSV = $true"
                    Remove-Item -Path $renamedPath -Force
                } else {
                    WriteLog -Message "File already exists: $renamedPath. Rename aborted." "ERROR"
					$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Duplicates Primary SMTP detected - File already exists: $renamedPath. Rename aborted."
					SendEmailHtmlReport -BodyHtml $body
                    exit 1
                }
            }
            Rename-Item -Path $InputCsvForDuplicateScan -NewName $renamedPath
            WriteLog -Message "Original file renamed to: $renamedPath"
        } catch {
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Failed to rename original file: $_"
			SendEmailHtmlReport -BodyHtml $body
            WriteLog -Message "Failed to rename original file: $_" "ERROR"
            exit 1
        }

        # Move cleaned file back to original name
        try {
            if (Test-Path $InputCsvForDuplicateScan) {
                if ($ForceOverwriteCSV) {
                    WriteLog -Message "File already exists: $InputCsvForDuplicateScan. Removing due to ForceOverwriteCSV = $true"
                    Remove-Item -Path $InputCsvForDuplicateScan -Force
                } else {
                    WriteLog -Message "File already exists: $InputCsvForDuplicateScan. Move aborted." "ERROR"
					$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : File already exists: $InputCsvForDuplicateScan. Move aborted."
					SendEmailHtmlReport -BodyHtml $body
                    exit 1
                }
            }
            Move-Item -Path $TempOutputCsv -Destination $InputCsvForDuplicateScan
            WriteLog -Message "Cleaned file moved to: $InputCsvForDuplicateScan"
        } catch {
            WriteLog -Message "Failed to move cleaned file: $_" "ERROR"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : Failed to move cleaned file: $_"
			SendEmailHtmlReport -BodyHtml $body
            exit 1
        }

        # Summary log
        WriteLog -Message "Processing complete."
        WriteLog -Message "Total rows processed: $($rows.Count)"
        WriteLog -Message "Unique rows retained: $($uniqueRows.Count)"
        WriteLog -Message "Duplicates found: $($duplicates.Count)"
        WriteLog -Message "Cleaned file saved to: $InputCsvForDuplicateScan"
        WriteLog -Message "Duplicates logged in: $LogFile"
        WriteLog -Message "Original file renamed to: $renamedPath"
    }
}

if ($scriptdatamailbox -eq $true -and $scriptdatamegewithperm -eq $true -and (-not $OnlyADPermission)) {
	Write-Host "-----------------------------------------------------------------------------------------"
	Write-Host "Combine Export Exchange 2016 mailboxes inventory with last permission inventory ..."
	try {
		$SourceDirectory = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxOnlyAdPermissionCsvLogFolderPath' -DefaultValue ''
	} catch {
		$errorMessage = "Error retrieving local configuration path LastOutput $_"
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}

	if (-not (Test-Path $SourceDirectory)) {
		$errorMessage = "The share '$SourceDirectory' is not available. Stopping the script."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}

	try {
		$DestinationDirectory = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCsvLogFolderPath' -DefaultValue ''
	} catch {
		$errorMessage = "Error retrieving local configuration path DestinationDirectory $_"
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}

	if (-not (Test-Path $DestinationDirectory)) {		
		$errorMessage = "The share '$DestinationDirectory' is not available. Stopping the script."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}
	
	try {
		$ScriptCsvLogFolderPathectory = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCombinedPermissionCsvLogFolderPath' -DefaultValue ''
	} catch {		
		$errorMessage = "Error retrieving local configuration path ScriptCsvLogFolderPathectory $_"
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}

	if (-not (Test-Path $ScriptCsvLogFolderPathectory)) {
		$errorMessage = "The share '$ScriptCsvLogFolderPathectory' is not available. Stopping the script."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}

	# --- Step 2: Process files based on mode (Consolidation vs. Standard) ---

	$allDomainsFile = Get-ChildItem -Path $DestinationDirectory -Filter "*_AllDomains.csv" -File | Select-Object -First 1
	WriteLog -Message "$allDomainsFile found in the Destination directory."
	
	$sourceFiles = Get-ChildItem -Path $SourceDirectory -Filter "*_OnlyADPermission.csv" -File

	if ($sourceFiles.Count -eq 0) {
		WriteLog -Message "No '*_OnlyADPermission.csv' files found in the source directory. Nothing to process." "WARNING"
		$errorMessage = "Error, No '*_OnlyADPermission.csv' files found in the source directory. Nothing to process."
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}

	if ($allDomainsFile) {
		# --- Consolidation Mode: Merge all sources into one destination ---
		WriteLog -Message "Found '$($allDomainsFile.Name)'. Activating consolidation mode."
		$fullPaths = $sourceFiles | Select-Object -ExpandProperty FullName
		WriteLog -Message "All $($sourceFiles.Count) source files will be merged to update this single destination file:`n$($fullPaths -join "`n")"
		try {
			Write-Verbose "Creating a consolidated lookup table from all source files..."
			$consolidatedLookupTable = @{}
			foreach ($sourceFile in $sourceFiles) {
				Write-Verbose "  -> Adding source: $($sourceFile.Name)"
				Import-Csv -Path $sourceFile.FullName | ForEach-Object {
					if (-not [string]::IsNullOrEmpty($_.UserPrincipalName)) {
						$consolidatedLookupTable[$_.UserPrincipalName] = $_
					}
				}
			}
			Write-Verbose "Consolidated lookup table created with $($consolidatedLookupTable.Count) unique entries."

			Write-Verbose "Importing data from '$($allDomainsFile.Name)'..."
			$destinationData = Import-Csv -Path $allDomainsFile.FullName
			
			$updatedCount = 0
			$notFoundCount = 0

			$destinationData | ForEach-Object {
				$currentUserPrincipalName = $_.UserPrincipalName
				if ($consolidatedLookupTable.ContainsKey($currentUserPrincipalName)) {
					$_.SendAsCount = $consolidatedLookupTable[$currentUserPrincipalName].SendAsCount
					$_.SendAs = $consolidatedLookupTable[$currentUserPrincipalName].SendAs
					$updatedCount++
				} else {
					$notFoundCount++
				}
			}
			
			$outputFilePath = Join-Path -Path $ScriptCsvLogFolderPathectory -ChildPath $allDomainsFile.Name
			$SendFileListEmailReportFileName = $outputFilePath
			WriteLog -Message "Exporting updated data to '$outputFilePath'..."
			if ($PSCmdlet.ShouldProcess($outputFilePath, "Export Updated Consolidated CSV")) {
				Export-CsvAtomic -InputObject $destinationData -Path $outputFilePath -Encoding UTF8
			}
			
			WriteLog -Message "--------------------------------------------------"
			WriteLog -Message "Consolidation complete."
			WriteLog -Message "File updated: $outputFilePath"
			WriteLog -Message "Users updated: $updatedCount. Not found: $notFoundCount."

		} catch {
			Write-Error "An unexpected error occurred while processing the consolidated file '$($allDomainsFile.Name)'. Details: $_"
			$errorMessage = "ERROR retrieving local configuration path: $_"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
			SendEmailHtmlReport -BodyHtml $body
			throw
		}

	} else {
		# --- Standard Mode: One-to-one file matching ---
		WriteLog -Message "No '*_AllDomains.csv' file found. Proceeding with standard one-to-one matching."
		
		$summary = @{ Processed = 0; Skipped = 0; Errors = 0; SkippedFiles = [System.Collections.Generic.List[string]]::new() }

		foreach ($sourceFile in $sourceFiles) {
			try {
				WriteLog -Message "Processing source file: $($sourceFile.Name)"
				$baseName = $sourceFile.BaseName.Replace('_OnlyADPermission', '')
				$destinationFile = Get-ChildItem -Path $DestinationDirectory -Filter "$baseName*.csv" -File | Select-Object -First 1

				if (-not $destinationFile) {
					WriteLog -Message "No matching destination file found for '$baseName*' for '$($sourceFile.Name)'. File skipped." "WARNING"
					$summary.Skipped++; $summary.SkippedFiles.Add($sourceFile.Name)
					continue
				}

				WriteLog -Message "Match found: '$($sourceFile.Name)' -> '$($destinationFile.Name)'"

				$lookupTable = @{}
				Import-Csv -Path $sourceFile.FullName | ForEach-Object {
					if (-not [string]::IsNullOrEmpty($_.UserPrincipalName)) {
						$lookupTable[$_.UserPrincipalName] = $_
					}
				}

				$destinationData = Import-Csv -Path $destinationFile.FullName
				$updatedCount = 0; $notFoundCount = 0

				$destinationData | ForEach-Object {
					if ($lookupTable.ContainsKey($_.UserPrincipalName)) {
						$_.SendAsCount = $lookupTable[$_.UserPrincipalName].SendAsCount
						$_.SendAs = $lookupTable[$_.UserPrincipalName].SendAs
						$updatedCount++
					} else {
						$notFoundCount++
					}
				}
				
				$outputFilePath = Join-Path -Path $ScriptCsvLogFolderPathectory -ChildPath $destinationFile.Name
				$SendFileListEmailReportFileName = $outputFilePath
				if ($PSCmdlet.ShouldProcess($outputFilePath, "Export Updated CSV")) {
					Export-CsvAtomic -InputObject $destinationData -Path $outputFilePath -Encoding UTF8
				}
				
				$summary.Processed++
				WriteLog -Message "Update for '$($destinationFile.Name)' finished. Users updated: $updatedCount. Not found: $notFoundCount."

			} catch {
				Write-Error "An unexpected error occurred while processing file '$($sourceFile.Name)'. Details: $_"
				$summary.Errors++
				$errorMessage = "ERROR retrieving local configuration path: $_"
				$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
				SendEmailHtmlReport -BodyHtml $body
				throw
			}
		}

		# --- Final Summary for Standard Mode ---
		WriteLog -Message "--------------------------------------------------"
		WriteLog -Message "All files processed."
		WriteLog -Message "Files processed successfully: $($summary.Processed)"
		WriteLog -Message "Files skipped (no destination): $($summary.Skipped)"
		WriteLog -Message "Errors: $($summary.Errors)"

		if ($summary.SkippedFiles.Count -gt 0) {
			WriteLog -Message "List of skipped files:" "WARNING"
			$summary.SkippedFiles | ForEach-Object { Write-Warning "- $_" }
		}
	}	
}

if ($scriptdatamailbox -eq $true -and $scriptdatamegewithbatch -eq $true -and (-not $OnlyADPermission)) {
	Write-Host "-----------------------------------------------------------------------------------------"
	Write-Host "Combine Export Exchange 2016 mailboxes inventory with batch name ..."
	$StartTime = Get-Date
	$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss" # Timestamp for unique filenames
		try {
		$OutputPathBatch = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ExchangeMigrationBatchFolderPath' -DefaultValue ''
	} catch {		
		$errorMessage = "Error retrieving local configuration path OutputPathBatch $_"
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}

	if (-not (Test-Path $OutputPathBatch)) {
		WriteLog -Message "The share '$OutputPathBatch' is not available. Stopping the script." "ERROR"
		$errorMessage = "The share '$OutputPathBatch' is not available. Stopping the script."
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}
	
	$outputFolderForBatchCsv = Join-Path -Path $OutputPathBatch -ChildPath "Output-Batches"      # For individual batch CSV files
	$outputFolderForConsolidatedCsv = Join-Path -Path $OutputPathBatch -ChildPath "Output-ConsolidatedCsv" # For consolidated summary CSVs
	$outputFolderForBatchCsvBackup = Join-Path -Path $OutputPathBatch -ChildPath "Output-Batches-Backup" # For consolidated summary CSVs
	# Create output folders if they don't exist

	if (-not (Test-Path $outputFolderForBatchCsv)) {
		try {
			New-Item -ItemType Directory -Path $outputFolderForBatchCsv -Force -ErrorAction Stop | Out-Null
		} catch {
			$errorMessage = "Unable to create output folder for batches: '$outputFolderForBatchCsv'. Error: $($_.Exception.Message)"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
			SendEmailHtmlReport -BodyHtml $body
			throw
		}
	}
	if (-not (Test-Path $outputFolderForConsolidatedCsv)) {
		try {
			New-Item -ItemType Directory -Path $outputFolderForConsolidatedCsv -Force -ErrorAction Stop | Out-Null
		} catch {
			$errorMessage = "Unable to create output folder for consolidated CSVs: '$outputFolderForConsolidatedCsv'. Error: $($_.Exception.Message)"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
			SendEmailHtmlReport -BodyHtml $body
			throw
		}
	}
	if (-not (Test-Path $outputFolderForBatchCsvBackup)) {
		try {
			New-Item -ItemType Directory -Path $outputFolderForBatchCsvBackup -Force -ErrorAction Stop | Out-Null
		} catch {
			$errorMessage = "Unable to create output folder for consolidated CSVs: '$outputFolderForBatchCsvBackup'. Error: $($_.Exception.Message)"
			$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
			SendEmailHtmlReport -BodyHtml $body
			throw
		}
	}

	try {
		$inputFolderCSVfiles = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalMailboxCombinedPermissionCsvLogFolderPath' -DefaultValue ''
	} catch {
		$errorMessage = "Error retrieving local configuration path inputFolderCSVfiles $_"
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
		throw
	}

	if (-not (Test-Path $inputFolderCSVfiles)) {
		$errorMessage = "The share '$inputFolderCSVfiles' is not available. Stopping the script."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
	} else {
		Write-Host "The network share '$inputFolderCSVfiles' is available. Continuing the script..." -ForegroundColor Green
	}

	$timestampCsvBackup = Get-Date -Format "yyyyMMdd_HHmmss"
	$archiveFolder = Join-Path -Path $outputFolderForBatchCsvBackup -ChildPath "archive_$timestampCsvBackup"

	if ((Get-ChildItem -Path $outputFolderForBatchCsv -File).Count -gt 0) {
		Write-Host -ForegroundColor Yellow "Warning: The processing folder '$outputFolderForBatchCsv' is not empty. Moving existing files to an archive subfolder."
		WriteLog -Message "Warning: The processing folder '$outputFolderForBatchCsv' is not empty. Moving existing files to archive subfolder '$archiveFolder'."

		try {
			# Ensure the backup folder exists
			if (-not (Test-Path -Path $outputFolderForBatchCsvBackup -PathType Container)) {
				Write-Host -ForegroundColor Cyan "The backup folder '$outputFolderForBatchCsvBackup' does not exist. Creating it..."
				WriteLog -Message "The backup folder '$outputFolderForBatchCsvBackup' does not exist. Creating it..."
				New-Item -Path $outputFolderForBatchCsvBackup -ItemType Directory -ErrorAction Stop
			}

			# Create the timestamped archive subfolder
			New-Item -Path $archiveFolder -ItemType Directory -ErrorAction Stop
			WriteLog -Message "Archive folder created: '$archiveFolder'."

			# Move files from the processing folder to the archive folder
			Get-ChildItem -Path $outputFolderForBatchCsv -File | Move-Item -Destination $archiveFolder -Force -ErrorAction Stop
			WriteLog -Message "Successfully moved existing files to '$archiveFolder'."
			Write-Host -ForegroundColor Green "Existing files successfully moved to '$archiveFolder'. Proceeding with script execution."
		} catch {
			Write-Host -ForegroundColor Red "Error: Failed to move existing files to the archive folder. Script aborted."
			WriteLog -Message "Error: Failed to move existing files to the archive folder. Error: $($_.Exception.Message). Script aborted."
			exit 1
		}
	} else {
		WriteLog -Message "The processing folder '$outputFolderForBatchCsv' is empty. Proceeding with script execution."
		Write-Host -ForegroundColor Green "The processing folder '$outputFolderForBatchCsv' is empty. Proceeding with script execution."
	}

	if (-not $PSBoundParameters.ContainsKey('outputConsolidatedCsvPath')) {
		# Clarified filename to reflect its content: All source data, BatchName column updated by this script's run
		$outputConsolidatedCsvPath = Join-Path -Path $outputFolderForConsolidatedCsv -ChildPath "${outputFileNamePrefix}_AllSourceData_BatchNameUpdated_$TimeStamp.csv"
		WriteLog -Message "Parameter 'outputConsolidatedCsvPath' (all source mailboxes, BatchName updated by script) not specified, defaulting to: '$outputConsolidatedCsvPath'"
	}

	if (-not $PSBoundParameters.ContainsKey('outputConsolidatedCsvPath2')) {
		# Clarified filename to reflect its content: All source data, BatchName column updated by this script's run
        $outputConsolidatedCsvPath2 = Join-Path -Path $outputFolderForConsolidatedCsv -ChildPath "Exchange_OnPrem_Mailboxes_AllDomains.csv"
		WriteLog -Message "Parameter 'outputConsolidatedCsvPath2' (all source mailboxes, BatchName updated by script) not specified, defaulting to: '$outputConsolidatedCsvPath2'"
	}

	WriteLog -Message "--- Script Parameters Used ---"
	WriteLog -Message "InputFolderCSVfiles: '$inputFolderCSVfiles'"
	WriteLog -Message "ExcludeMailboxesFile: '$excludeMailboxesFile'"
	WriteLog -Message "OutputFolder: '$outputFolder'"
	WriteLog -Message "OutputConsolidatedCsvPath (All Source Data, BatchName updated): '$outputConsolidatedCsvPath'"
	WriteLog -Message "OutputConsolidatedCsvPath2 (All Source Data, BatchName updated): '$outputConsolidatedCsvPath2'"
	WriteLog -Message "OutputFileNamePrefix: '$outputFileNamePrefix'"
	WriteLog -Message "IncludeSpecificLastLogonCriteria: $IncludeSpecificLastLogonCriteria (If True, overrides other disabled account filters)"
	if (-not $IncludeSpecificLastLogonCriteria) {
		WriteLog -Message "ExcludeAllDisabledAccounts: $ExcludeAllDisabledAccounts (Default=False, Excludes ALL disabled if True)"
		WriteLog -Message "ExcludeDisabledAccountsExceptForSharedMailboxes: $ExcludeDisabledAccountsExceptForSharedMailboxes (Default=True, Excludes disabled non-shared if True and ExcludeAllDisabledAccounts=False)"
		WriteLog -Message "ExcludeDisabledMailboxes: $ExcludeDisabledMailboxes (Default=True, Excludes if True)"
	} else {
		WriteLog -Message "ExcludeAllDisabledAccounts, ExcludeDisabledAccountsExceptForSharedMailboxes, ExcludeDisabledMailboxes are IGNORED because IncludeSpecificLastLogonCriteria is True."
	}
	if ($PSBoundParameters.ContainsKey('ExcludeSamAccountNamePatterns') -and $ExcludeSamAccountNamePatterns) { WriteLog -Message "ExcludeSamAccountNamePatterns: $($ExcludeSamAccountNamePatterns -join ', ')" } else { WriteLog -Message "ExcludeSamAccountNamePatterns: Not specified." }
	WriteLog -Message "SimpleBatchCsv (Minimal columns in individual batch CSVs): $SimpleBatchCsv"
	WriteLog -Message "MaxBatchSizeMB (Max total size per batch in MB): $maxBatchSizeMB"
	WriteLog -Message "MaxMailboxesPerBatchFile (Max items per batch file): $MaxMailboxesPerBatchFile"
	if ($excludeFullAccessSamAccounts) { WriteLog -Message "ExcludeFullAccessSamAccounts: $($excludeFullAccessSamAccounts -join ', ')" } else { WriteLog -Message "ExcludeFullAccessSamAccounts: Not specified." }
	if ($excludeSendAsSamAccounts) { WriteLog -Message "ExcludeSendAsSamAccounts: $($excludeSendAsSamAccounts -join ', ')" } else { WriteLog -Message "ExcludeSendAsSamAccounts: Not specified." }
	if ($excludeSendOnBehalfToSamAccounts) { WriteLog -Message "ExcludeSendOnBehalfToSamAccounts: $($excludeSendOnBehalfToSamAccounts -join ', ')" } else { WriteLog -Message "ExcludeSendOnBehalfToSamAccounts: Not specified." }
	WriteLog -Message "--- End of Script Parameters ---"

	try {
		Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1))
	} catch {
		Write-Warning "Could not determine console width. Skipping console separator line."
	}

	# RECIPIENT TYPE CONSTANTS
	$sharedMailboxRecipientTypeName = "SharedMailbox"
	$discoveryMailboxRecipientTypeName = "DiscoveryMailbox"
	$roomMailboxRecipientTypeName = "RoomMailbox"
	# Add other types if needed for filtering

	# CSV Column name constants from input files
	$ouColumnNameInCsv = "RecipientOU"
	$primarySmtpAddressColumnNameInCsv = "PrimarySmtpAddress"
	$recipientTypeColumnNameInCsv = "RecipientType" # Used for SharedMailbox exception
	$accountDisabledColumnNameInCsv = "AccountDisabled" # Used for disable filtering
	$isMailboxEnabledColumnNameInCsv = "IsMailboxEnabled"
	$lastLogonTimeColumnNameInCsv = "LastLogonTime" # NEW - For new filtering criteria
	$fullAccessColumnNameInCsv = "FullAccess"
	$sendAsColumnNameInCsv = "SendAs"
	$sendOnBehalfToColumnNameInCsv = "GrantSendOnBehalfTo"
	$displayNameColumnNameInCsv = "DisplayName"
	$aliasColumnNameInCsv = "Alias"
	$samAccountNameColumnNameInCsv = "SamAccountName" # Ensure it's here
	$identityColumnNameInCsv = "Identity"
	$domainNameColumnNameInCsv = "DomainName"
	$totalItemSizeColumnNameInCsv = "TotalItemSize-In-MB" # Crucial for the size limit
	# This is the name of the batch name column expected in source CSVs AND the column this script will update/use.
	$batchNameColumn = "BatchName"

	# Column name added by this script to track source file (if not already present)
	$originalSourceFileColumnAddedByScript = "OriginalSourceFile"

	# Date threshold for new criteria
	$lastLogonDateThreshold = Get-Date "2024-05-01"

	Write-Host -ForegroundColor White "Checking input folder for CSV files... $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
	WriteLog -Message "Checking input folder: '$inputFolderCSVfiles' for CSV files."

	if (-not (Test-Path $inputFolderCSVfiles -PathType Container)) {
		$errorMessage = "The input folder '$inputFolderCSVfiles' does not exist or is not a directory. Script cannot continue."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
	}

	$allCsvFilesInSourceFolder = Get-ChildItem -Path $inputFolderCSVfiles -Filter "*.csv" -File
	$prioritizedFileSuffixToProcess = "AllDomains.csv"
	$foundPrioritizedFiles = $allCsvFilesInSourceFolder | Where-Object { $_.Name -like "*$prioritizedFileSuffixToProcess" }

	$csvFilesToProcessForImport = $null

	if ($foundPrioritizedFiles.Count -gt 0) {
		$csvFilesToProcessForImport = @($foundPrioritizedFiles)
		$infoMessageText = ("Found $($foundPrioritizedFiles.Count) CSV file(s) ending with '$prioritizedFileSuffixToProcess': " +
							"$($foundPrioritizedFiles.Name -join ', '). " +
							"Only these files will be processed.")
		Write-Host -ForegroundColor Green $infoMessageText
		WriteLog -Message $infoMessageText
	} else {
		$csvFilesToProcessForImport = $allCsvFilesInSourceFolder
		if ($csvFilesToProcessForImport.Count -gt 0) {
			$infoMessageText = ("No CSV files ending with '$prioritizedFileSuffixToProcess' were found. " +
								"Processing all $($csvFilesToProcessForImport.Count) CSV files found in '$inputFolderCSVfiles': " +
								"$($csvFilesToProcessForImport.Name -join ', ').")
			Write-Host $infoMessageText
			WriteLog -Message $infoMessageText
		}
	}

	if ($null -eq $csvFilesToProcessForImport -or $csvFilesToProcessForImport.Count -eq 0) {
		$errorMessage = "No CSV files found in the input folder '$inputFolderCSVfiles' matching the criteria for processing. Ensure the folder contains valid CSV files. Script cannot continue."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
	}

	if (-not $PSBoundParameters.ContainsKey('excludeMailboxesFile')) {
		$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
		$excludeMailboxesFile = Join-Path -Path $scriptDirectory -ChildPath "Exchange-MigrationBatch-Mailboxes-Excluded.csv"
		WriteLog -Message "Parameter -ExcludeMailboxesFile not specified, defaulting to: '$excludeMailboxesFile'"
	}

	$excludedMailboxPrimarySmtpAddresses = @()

	if (-not (Test-Path $excludeMailboxesFile)) {
		Write-Warning "The exclusion file '$excludeMailboxesFile' does not exist. Mailbox exclusion based on this file will be skipped."
		WriteLog -Message "The exclusion file '$excludeMailboxesFile' does not exist. Mailbox exclusion based on this file will be skipped."
	} else {
		try {
			# Ensure PrimarySmtpAddress exists before expanding
			$csvData = Import-Csv $excludeMailboxesFile -ErrorAction Stop
			if ($csvData -and $csvData[0].PSObject.Properties.Name -contains $primarySmtpAddressColumnNameInCsv) {
				$excludedMailboxPrimarySmtpAddresses = $csvData | Select-Object -ExpandProperty $primarySmtpAddressColumnNameInCsv
				Write-Host "Importing $($excludedMailboxPrimarySmtpAddresses.Count) users to exclude from file '$excludeMailboxesFile'."
				WriteLog -Message "Importing $($excludedMailboxPrimarySmtpAddresses.Count) users to exclude from file '$excludeMailboxesFile'."
			} else {
				 Write-Warning "Exclusion file '$excludeMailboxesFile' found but does not contain the required column '$primarySmtpAddressColumnNameInCsv' or is empty. Skipping exclusion based on this file."
				 WriteLog -Message "Exclusion file '$excludeMailboxesFile' found but does not contain the required column '$primarySmtpAddressColumnNameInCsv' or is empty. Skipping exclusion based on this file."
			}
		} catch {
			$errorMessage = "Error importing the exclusion file '$excludeMailboxesFile': $($_.Exception.Message). Mailbox exclusion based on this file will be skipped."
			Write-Error $errorMessage
			WriteLog -Message $errorMessage "ERROR"
		}
	}

	WriteLog -Message "Individual batch CSV files will be saved to: '$outputFolderForBatchCsv'"
	WriteLog -Message "Consolidated summary CSV files will be saved to: '$outputFolderForConsolidatedCsv'"

	$allMailboxesOriginalData = [System.Collections.Generic.List[PSObject]]::new() # Use Generic List for better performance
	$isFirstFileProcessed = $false
	$inputFileImportStatistics = [System.Collections.Generic.List[PSObject]]::new() # Use Generic List

	$totalFilesToProcessInQueue = $csvFilesToProcessForImport.Count
	$filesProcessedCounter = 0

	Write-Host "Starting import and validation of input CSV files..."
	foreach ($csvFileItem in $csvFilesToProcessForImport) {
		$filesProcessedCounter++
		$statusMessageForFileImport = "File {0} of {1}: {2}" -f $filesProcessedCounter, $totalFilesToProcessInQueue, $csvFileItem.Name
		Write-Progress -Activity "Processing input CSV files" -Status $statusMessageForFileImport -PercentComplete (($filesProcessedCounter / $totalFilesToProcessInQueue) * 100)

		WriteLog -Message "Processing input CSV file: $($csvFileItem.FullName)"
		$mailboxesInCurrentFileBeforeFilter = 0
		$disabledAccountsInCurrentFile = 0
		$disabledMailboxesInCurrentFile = 0
		$currentFileProcessingStatusMessage = "Error during processing"
		$fileStatObject = [PSCustomObject]@{
			FileName = $csvFileItem.Name
			MailboxesBeforeFiltering = 0
			DisabledAccountsInFile = 0
			DisabledMailboxesInFile = 0
			MailboxesAfterFiltering = 0
			Status = $currentFileProcessingStatusMessage
		}

		try {
			$mailboxesFromCurrentCsv = Import-Csv -Path $csvFileItem.FullName -ErrorAction Stop

			if ($null -eq $mailboxesFromCurrentCsv -or $mailboxesFromCurrentCsv.Count -eq 0) {
				Write-Warning "CSV file '$($csvFileItem.FullName)' is empty or could not be imported properly. Skipping this file."
				WriteLog -Message "CSV file '$($csvFileItem.FullName)' is empty or could not be imported properly. Skipping this file."
				$fileStatObject.Status = "Skipped (File Empty or Import Error)"
				$inputFileImportStatistics.Add($fileStatObject)
				continue
			}

			$mailboxesInCurrentFileBeforeFilter = $mailboxesFromCurrentCsv.Count
			$firstObjectProperties = $mailboxesFromCurrentCsv[0].PSObject.Properties.Name

			# Check for column existence before attempting to count
			$hasAccountDisabledCol = $firstObjectProperties -contains $accountDisabledColumnNameInCsv
			$hasMbxEnabledCol = $firstObjectProperties -contains $isMailboxEnabledColumnNameInCsv
			$hasRecipientTypeCol = $firstObjectProperties -contains $recipientTypeColumnNameInCsv # Needed for stats
			$hasLastLogonTimeCol = $firstObjectProperties -contains $lastLogonTimeColumnNameInCsv # Needed for new criteria stats (optional here, checked strongly later)


			if ($hasAccountDisabledCol) {
				$disabledAccountsInCurrentFile = ($mailboxesFromCurrentCsv | Where-Object { $_.$accountDisabledColumnNameInCsv -eq 'True' }).Count
			} else { WriteLog -Message "Warning: Column '$accountDisabledColumnNameInCsv' not found in '$($csvFileItem.FullName)'. Cannot count disabled accounts in this file." }

			if ($hasMbxEnabledCol) {
				$disabledMailboxesInCurrentFile = ($mailboxesFromCurrentCsv | Where-Object { $_.$isMailboxEnabledColumnNameInCsv -eq 'False' }).Count
			} else { WriteLog -Message "Warning: Column '$isMailboxEnabledColumnNameInCsv' not found in '$($csvFileItem.FullName)'. Cannot count disabled mailboxes in this file." }

			# Add required columns if they don't exist and add to the main list
			$mailboxesFromCurrentCsv | ForEach-Object {
				$currentObject = $_
				# Ensure OriginalSourceFile column exists and is set
				if (-not ($currentObject.PSObject.Properties.Name -contains $originalSourceFileColumnAddedByScript)) {
					$currentObject | Add-Member -MemberType NoteProperty -Name $originalSourceFileColumnAddedByScript -Value $csvFileItem.Name
				} else {
					$currentObject.$originalSourceFileColumnAddedByScript = $csvFileItem.Name
				}
				# Ensure BatchName column exists and is initialized (important for later update)
				if (-not ($currentObject.PSObject.Properties.Name -contains $batchNameColumn)) {
					 $currentObject | Add-Member -MemberType NoteProperty -Name $batchNameColumn -Value ""
				}
				# Add the processed object to the list
				$allMailboxesOriginalData.Add($currentObject)
			}

			# Column validation happens only on the first successfully processed file
			if (-not $isFirstFileProcessed) {
				 # Check against the first object loaded into the main list from this file
				 $firstLoadedObjectProperties = $allMailboxesOriginalData[0].PSObject.Properties.Name
				 # Define required columns (ensure all necessary ones are listed)
				 $requiredInputColumnsList = @(
					 $ouColumnNameInCsv, $primarySmtpAddressColumnNameInCsv,
					 $recipientTypeColumnNameInCsv, # Needed for SharedMailbox exception
					 $accountDisabledColumnNameInCsv, # Needed for disable filtering
					 $isMailboxEnabledColumnNameInCsv, $fullAccessColumnNameInCsv,
					 $sendAsColumnNameInCsv, $sendOnBehalfToColumnNameInCsv, $samAccountNameColumnNameInCsv, # Ensure SamAccountName is required
					 $identityColumnNameInCsv, $domainNameColumnNameInCsv, $totalItemSizeColumnNameInCsv # Ensure size column is required
				 )
				 # Conditionally add LastLogonTime to required list if new feature is enabled
				 if ($IncludeSpecificLastLogonCriteria) {
					 if (-not ($firstLoadedObjectProperties -contains $lastLogonTimeColumnNameInCsv)) {
						$errorMessage = ("The parameter -IncludeSpecificLastLogonCriteria is True, but the required column '$lastLogonTimeColumnNameInCsv' is missing in the first processed input CSV file '$($csvFileItem.FullName)'. " +
										 "Aborting script.")
						Write-Progress -Activity "Processing input CSV files" -Completed
						WriteLog -Message $errorMessage "ERROR"
						$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
						SendEmailHtmlReport -BodyHtml $body
					 }
					 # Add to list for general check if it was already there, or just note it's needed
					 if ($requiredInputColumnsList -notcontains $lastLogonTimeColumnNameInCsv) {
						 # $requiredInputColumnsList += $lastLogonTimeColumnNameInCsv # Not strictly needed here as we checked above
					 }
				 }


				$missingInputColumnsList = @()
				foreach ($columnNameToVerify in $requiredInputColumnsList) {
					if (-not ($firstLoadedObjectProperties -contains $columnNameToVerify)) {
						$missingInputColumnsList += $columnNameToVerify
					}
				}

				if ($missingInputColumnsList.Count -gt 0) {
					# Improve error message readability
					$errorMessage = ("The first processed input CSV file '$($csvFileItem.FullName)' is missing one or more required columns for script operation. " +
									 "Missing columns: $($missingInputColumnsList -join ', '). Please ensure all required columns are present. Aborting script.")
					Write-Progress -Activity "Processing input CSV files" -Completed
					WriteLog -Message $errorMessage "ERROR"
					$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
					SendEmailHtmlReport -BodyHtml $body
				}
				$isFirstFileProcessed = $true
				WriteLog -Message "Column structure successfully validated against first processed file '$($csvFileItem.FullName)' for essential columns (including '$recipientTypeColumnNameInCsv', '$accountDisabledColumnNameInCsv', '$isMailboxEnabledColumnNameInCsv', '$totalItemSizeColumnNameInCsv', '$samAccountNameColumnNameInCsv')."
				if ($IncludeSpecificLastLogonCriteria) {
					WriteLog -Message "Additionally, column '$lastLogonTimeColumnNameInCsv' was confirmed as present for -IncludeSpecificLastLogonCriteria."
				}
			}

			WriteLog -Message ("Successfully imported $mailboxesInCurrentFileBeforeFilter mailboxes from '$($csvFileItem.FullName)'. " +
						 "Raw Disabled Accounts in file: $disabledAccountsInCurrentFile, Raw Disabled Mailboxes in file: $disabledMailboxesInCurrentFile. " +
						 "Total mailboxes loaded so far: $($allMailboxesOriginalData.Count).")
			$currentFileProcessingStatusMessage = "Processed"

			$fileStatObject.MailboxesBeforeFiltering = $mailboxesInCurrentFileBeforeFilter
			$fileStatObject.DisabledAccountsInFile = $disabledAccountsInCurrentFile
			$fileStatObject.DisabledMailboxesInFile = $disabledMailboxesInCurrentFile
			$fileStatObject.Status = $currentFileProcessingStatusMessage
			$inputFileImportStatistics.Add($fileStatObject)

		} catch {
			$errorMessage = "Error processing CSV file '$($csvFileItem.FullName)': $($_.Exception.Message). This file will be skipped."
			Write-Error $errorMessage
			WriteLog -Message $errorMessage "ERROR"
			$fileStatObject.Status = "$($currentFileProcessingStatusMessage): $($_.Exception.Message)"
			$inputFileImportStatistics.Add($fileStatObject) # Add stats even if error occurred after initial count
		}
	}

	Write-Progress -Activity "Processing input CSV files" -Completed

	if ($allMailboxesOriginalData.Count -eq 0) {
		$errorMessage = "No mailboxes were loaded from any of the input CSV files. Aborting script."
		WriteLog -Message $errorMessage "ERROR"
		$body = NewSimpleEmailBody -Title $TaskName -Message "$TaskName : $errorMessage"
		SendEmailHtmlReport -BodyHtml $body
	}

	$countBeforeGlobalFiltering = $allMailboxesOriginalData.Count
	WriteLog -Message "Starting global filtering for batch processing. Total mailboxes before any filtering: $countBeforeGlobalFiltering."

	# Use List<T>.FindAll() or iterate for filtering
	$filteredMailboxesForBatchesList = [System.Collections.Generic.List[PSObject]]::new()
	$excludedByRecipientTypeCount = 0
	$excludedByExclusionFileCount = 0
	$excludedByAllDisabledCount = 0 # For original logic
	$excludedByDisabledNonSharedCount = 0 # For original logic
	$excludedByMbxDisabledCount = 0 # For original logic
	$excludedBySamAccountNamePatternCount = 0 
	$excludedByNewLastLogonCriteriaCount = 0 # NEW - Counter for new criteria
	$includedByNewLastLogonCriteria_AcctDisabledTrue = 0 # NEW
	$includedByNewLastLogonCriteria_AcctDisabledFalseLastLogonOK = 0 # NEW
	$excludedByMissingColumnsForNewCriteriaCount = 0 # NEW

	$allMailboxesOriginalData | ForEach-Object {
		$mailbox = $_
		$includeMailbox = $true # Assume inclusion unless a filter excludes it
		$exclusionReason = "" # Track why it was excluded for logging

		# Check required properties exist before filtering (general ones)
		$hasRecipientType = $mailbox.PSObject.Properties.Name -contains $recipientTypeColumnNameInCsv
		$hasPrimarySmtp = $mailbox.PSObject.Properties.Name -contains $primarySmtpAddressColumnNameInCsv
		$hasSamAccountName = $mailbox.PSObject.Properties.Name -contains $samAccountNameColumnNameInCsv
		
		# --- Primary Inclusion/Exclusion based on New LastLogonTime Criteria (if active) ---
		if ($IncludeSpecificLastLogonCriteria) {
			$currentMailboxPrimarySmtp = if ($hasPrimarySmtp) { $mailbox.$primarySmtpAddressColumnNameInCsv } else { "UnknownPrimarySmtp" }
			$hasAccountDisabledForNewLogic = $mailbox.PSObject.Properties.Name -contains $accountDisabledColumnNameInCsv
			$hasLastLogonTimeForNewLogic = $mailbox.PSObject.Properties.Name -contains $lastLogonTimeColumnNameInCsv

			if (-not $hasAccountDisabledForNewLogic -or -not $hasLastLogonTimeForNewLogic) {
				WriteLog -Message "Warning: Mailbox '$currentMailboxPrimarySmtp' is missing '$accountDisabledColumnNameInCsv' or '$lastLogonTimeColumnNameInCsv'. Cannot apply new logon criteria. Excluding from batch consideration."
				$includeMailbox = $false
				$exclusionReason = "Missing columns for LastLogon criteria ($accountDisabledColumnNameInCsv or $lastLogonTimeColumnNameInCsv)"
				$excludedByMissingColumnsForNewCriteriaCount++
			} else {
				$isAccountDisabledFlagForNewLogic = ($mailbox.$accountDisabledColumnNameInCsv -eq 'True')

				if ($isAccountDisabledFlagForNewLogic) {
					# Include if AccountDisabled is True
					Write-Verbose "Including '$currentMailboxPrimarySmtp' because AccountDisabled is True (New LastLogon Criteria Active)."
					$includedByNewLastLogonCriteria_AcctDisabledTrue++
					# $includeMailbox remains true
				} else {
					# AccountDisabled is False, check LastLogonTime
					$lastLogonTimeStringValue = $mailbox.$lastLogonTimeColumnNameInCsv
					$lastLogonTimeDateValue = $null
					
					if ([string]::IsNullOrWhiteSpace($lastLogonTimeStringValue)) {
						WriteLog -Message "Warning: LastLogonTime is empty or whitespace for '$currentMailboxPrimarySmtp' (AccountDisabled=False). Excluding from new criteria."
						$includeMailbox = $false
						$exclusionReason = "Empty LastLogonTime (AccountDisabled False, New LastLogon Criteria Active)"
						$excludedByNewLastLogonCriteriaCount++
					} else {
						try {
							$lastLogonTimeDateValue = Get-Date $lastLogonTimeStringValue
							if ($lastLogonTimeDateValue -gt $lastLogonDateThreshold) {
								# Include if LastLogonTime > threshold
								Write-Verbose "Including '$currentMailboxPrimarySmtp' because AccountDisabled is False and LastLogonTime ($lastLogonTimeDateValue) > $lastLogonDateThreshold (New LastLogon Criteria Active)."
								$includedByNewLastLogonCriteria_AcctDisabledFalseLastLogonOK++
								# $includeMailbox remains true
							} else {
								$includeMailbox = $false
								$exclusionReason = "AccountDisabled False, LastLogonTime ($lastLogonTimeDateValue) not after $lastLogonDateThreshold (New LastLogon Criteria Active)"
								$excludedByNewLastLogonCriteriaCount++
								Write-Verbose "Excluding '$currentMailboxPrimarySmtp' due to: $exclusionReason"
							}
						} catch {
							WriteLog -Message "Warning: Could not parse LastLogonTime '$lastLogonTimeStringValue' for mailbox '$currentMailboxPrimarySmtp' (AccountDisabled=False). Excluding from new criteria. Error: $($_.Exception.Message)"
							$includeMailbox = $false
							$exclusionReason = "Unparseable LastLogonTime (AccountDisabled False, New LastLogon Criteria Active)"
							$excludedByNewLastLogonCriteriaCount++
						}
					}
				}
			}
		} # End of $IncludeSpecificLastLogonCriteria block

		# --- Standard Filters (Applied if mailbox is still a candidate) ---
		# Filter 1: Recipient Type (must have column)
		if ($includeMailbox -and $hasRecipientType) {
			if ($mailbox.$recipientTypeColumnNameInCsv -eq $discoveryMailboxRecipientTypeName -or $mailbox.$recipientTypeColumnNameInCsv -eq $roomMailboxRecipientTypeName) {
				$includeMailbox = $false
				$exclusionReason = "RecipientType ($($mailbox.$recipientTypeColumnNameInCsv))"
				$excludedByRecipientTypeCount++
				Write-Verbose "Excluding $($mailbox.$primarySmtpAddressColumnNameInCsv) due to RecipientType: $($mailbox.$recipientTypeColumnNameInCsv)"
			}
		}

		# Filter 2: Exclusion File (must have column, only if still included)
		if ($includeMailbox -and $hasPrimarySmtp -and $excludedMailboxPrimarySmtpAddresses.Count -gt 0) {
			if ($excludedMailboxPrimarySmtpAddresses -contains $mailbox.$primarySmtpAddressColumnNameInCsv) {
				$includeMailbox = $false
				$exclusionReason = "Exclusion File ($excludeMailboxesFile)"
				$excludedByExclusionFileCount++
				Write-Verbose "Excluding $($mailbox.$primarySmtpAddressColumnNameInCsv) due to exclusion file: $excludeMailboxesFile"
			}
		}
		
		# Filter 3 (was 5): SamAccountName Patterns
		if ($includeMailbox -and $hasSamAccountName -and $PSBoundParameters.ContainsKey('ExcludeSamAccountNamePatterns') -and $ExcludeSamAccountNamePatterns.Count -gt 0) {
			$currentSamAccountName = $mailbox.$samAccountNameColumnNameInCsv
			if (-not [string]::IsNullOrEmpty($currentSamAccountName)) {
				foreach ($pattern in $ExcludeSamAccountNamePatterns) {
					if ($currentSamAccountName -ilike $pattern) {
						$includeMailbox = $false
						$exclusionReason = "SamAccountName ('$currentSamAccountName') matches pattern '$pattern' (case-insensitive, using -ilike)"
						$excludedBySamAccountNamePatternCount++
						Write-Verbose "Excluding $($mailbox.$primarySmtpAddressColumnNameInCsv) due to SamAccountName '$currentSamAccountName' matching pattern '$pattern'."
						break 
					}
				}
			}
		}

		# --- Original Disabled Account/Mailbox Filters (ONLY if New LastLogonTime Criteria is NOT active) ---
		if ($includeMailbox -and -not $IncludeSpecificLastLogonCriteria) {
			$hasAccountDisabledForOrigLogic = $mailbox.PSObject.Properties.Name -contains $accountDisabledColumnNameInCsv
			$hasMbxEnabledForOrigLogic = $mailbox.PSObject.Properties.Name -contains $isMailboxEnabledColumnNameInCsv
			# Note: $hasRecipientType already checked

			# Original Filter for Disabled Accounts 
			if ($includeMailbox -and $hasAccountDisabledForOrigLogic -and $hasRecipientType) {
				$isAccountDisabledFlag = ($mailbox.$accountDisabledColumnNameInCsv -eq "True")
				$isSharedMailboxFlag = ($mailbox.$recipientTypeColumnNameInCsv -eq $sharedMailboxRecipientTypeName)

				if ($isAccountDisabledFlag) {
					if ($ExcludeAllDisabledAccounts) {
						$includeMailbox = $false
						$exclusionReason = "Account Disabled (ExcludeAllDisabledAccounts=True)"
						$excludedByAllDisabledCount++
						Write-Verbose "Excluding $($mailbox.$primarySmtpAddressColumnNameInCsv) because account is disabled and ExcludeAllDisabledAccounts is True."
					} elseif ($ExcludeDisabledAccountsExceptForSharedMailboxes) {
						if (-not $isSharedMailboxFlag) {
							$includeMailbox = $false
							$exclusionReason = "Account Disabled (Exclude...ExceptShared=True, Not Shared)"
							$excludedByDisabledNonSharedCount++
							Write-Verbose "Excluding $($mailbox.$primarySmtpAddressColumnNameInCsv) because account is disabled, ExcludeDisabledAccountsExceptForSharedMailboxes is True, and it is NOT a Shared Mailbox."
						} else {
							Write-Verbose "Including disabled Shared Mailbox $($mailbox.$primarySmtpAddressColumnNameInCsv) because ExcludeDisabledAccountsExceptForSharedMailboxes is True."
							# $includeMailbox remains true
						}
					}
					# Else (both parameters False): Disabled account is included by default by this specific filter.
				}
			}

			# Original Filter for Disabled Mailboxes 
			 if ($includeMailbox -and $ExcludeDisabledMailboxes -and $hasMbxEnabledForOrigLogic) {
				if ($mailbox.$isMailboxEnabledColumnNameInCsv -eq "False") {
					$includeMailbox = $false
					$exclusionReason = "Mailbox Disabled (IsMailboxEnabled=False, ExcludeDisabledMailboxes=True)"
					$excludedByMbxDisabledCount++
					Write-Verbose "Excluding $($mailbox.$primarySmtpAddressColumnNameInCsv) because mailbox is disabled and ExcludeDisabledMailboxes is True."
				}
			}
		} # End of original disabled filters block

		# If mailbox passed all applicable filters, add it to the list
		if ($includeMailbox) {
			$filteredMailboxesForBatchesList.Add($mailbox)
		} else {
			# Optional: Log excluded mailboxes and reason (can be verbose)
			# WriteLog -Message "Excluded Mailbox: $($mailbox.$primarySmtpAddressColumnNameInCsv) - Reason: $exclusionReason"
		}
	} # End ForEach-Object over $allMailboxesOriginalData

	$filteredMailboxesForBatches = @($filteredMailboxesForBatchesList)
	$countAfterGlobalFiltering = $filteredMailboxesForBatches.Count

	WriteLog -Message "Filtering complete. Mailboxes remaining for batching: $countAfterGlobalFiltering."
	WriteLog -Message "Filter Breakdown:"
	WriteLog -Message "  - Excluded by Recipient Type ('$discoveryMailboxRecipientTypeName', '$roomMailboxRecipientTypeName'): $excludedByRecipientTypeCount"
	if ($excludedMailboxPrimarySmtpAddresses.Count -gt 0) {
		WriteLog -Message "  - Excluded by Exclusion File '$excludeMailboxesFile': $excludedByExclusionFileCount"
	}
	if ($PSBoundParameters.ContainsKey('ExcludeSamAccountNamePatterns') -and $ExcludeSamAccountNamePatterns.Count -gt 0) {
		WriteLog -Message "  - Excluded by SamAccountName pattern (Patterns: '$($ExcludeSamAccountNamePatterns -join ", ")'): $excludedBySamAccountNamePatternCount"
	}

	if ($IncludeSpecificLastLogonCriteria) {
		WriteLog -Message "  --- Using New LastLogonTime Criteria ---"
		WriteLog -Message "    - Included because AccountDisabled=True: $includedByNewLastLogonCriteria_AcctDisabledTrue"
		WriteLog -Message "    - Included because AccountDisabled=False AND LastLogonTime > $($lastLogonDateThreshold.ToString('yyyy-MM-dd')): $includedByNewLastLogonCriteria_AcctDisabledFalseLastLogonOK"
		WriteLog -Message "    - Excluded by New LastLogonTime Criteria (e.g., AccountDisabled=False and LastLogonTime not meeting threshold, or unparseable date): $excludedByNewLastLogonCriteriaCount"
		WriteLog -Message "    - Excluded due to missing '$accountDisabledColumnNameInCsv' or '$lastLogonTimeColumnNameInCsv' columns for New Criteria: $excludedByMissingColumnsForNewCriteriaCount"
	} else {
		WriteLog -Message "  --- Using Original Disabled Account/Mailbox Filters ---"
		if ($ExcludeAllDisabledAccounts) {
			WriteLog -Message "  - Excluded because Account Disabled (ExcludeAllDisabledAccounts=True): $excludedByAllDisabledCount"
		} elseif ($ExcludeDisabledAccountsExceptForSharedMailboxes) {
			WriteLog -Message "  - Excluded because Account Disabled and Not Shared Mailbox (ExcludeDisabledAccountsExceptForSharedMailboxes=True): $excludedByDisabledNonSharedCount"
		} else {
			WriteLog -Message "  - No accounts excluded based on disabled status (both Exclude*DisabledAccount* flags are False, or overridden by IncludeSpecificLastLogonCriteria)."
		}
		if ($ExcludeDisabledMailboxes) {
			WriteLog -Message "  - Excluded because Mailbox Disabled (IsMailboxEnabled=False, ExcludeDisabledMailboxes=True): $excludedByMbxDisabledCount"
		} else {
			WriteLog -Message "  - No mailboxes excluded based on IsMailboxEnabled=False (ExcludeDisabledMailboxes=False, or overridden by IncludeSpecificLastLogonCriteria)."
		}
	}


	# Sanity check:
	$totalExcludedDirectlyCount = $excludedByRecipientTypeCount + $excludedByExclusionFileCount + $excludedBySamAccountNamePatternCount
	if ($IncludeSpecificLastLogonCriteria) {
		$totalExcludedDirectlyCount += $excludedByNewLastLogonCriteriaCount + $excludedByMissingColumnsForNewCriteriaCount
	} else {
		$totalExcludedDirectlyCount += $excludedByAllDisabledCount + $excludedByDisabledNonSharedCount + $excludedByMbxDisabledCount
	}
	# This calculation is complex because filters are sequential. The sum of individual exclusion counts might be > (Total - Included) if a mailbox matches multiple criteria.
	# A more accurate way to get total excluded is $countBeforeGlobalFiltering - $countAfterGlobalFiltering
	$totalMailboxesExcludedOverall = $countBeforeGlobalFiltering - $countAfterGlobalFiltering
	WriteLog -Message "Total mailboxes excluded by all active filters (overall): $totalMailboxesExcludedOverall"


	# Update statistics based on the final filtered list
	foreach ($statEntryObject in $inputFileImportStatistics) {
		if ($statEntryObject.Status -eq "Processed") {
			# Count mailboxes from this specific file that are present in the final filtered list
			$countAfterFilteringForFile = ($filteredMailboxesForBatches | Where-Object { $_.$originalSourceFileColumnAddedByScript -eq $statEntryObject.FileName }).Count
			$statEntryObject.MailboxesAfterFiltering = $countAfterFilteringForFile
		}
	}

	# --- Display Summary Report ---
	Write-Host
	Write-Host -ForegroundColor Magenta "Input CSV Files & Global Filters Summary Report:"
	WriteLog -Message "Input CSV Files & Global Filters Summary Report:"
	Write-Host -ForegroundColor Cyan "Global Filters Applied (for mailboxes included in batch generation):"
	WriteLog -Message "Global Filters Applied (for mailboxes included in batch generation):"
	Write-Host "  - Recipient types '$discoveryMailboxRecipientTypeName' and '$roomMailboxRecipientTypeName': Always Excluded"
	WriteLog -Message "  - Recipient types '$discoveryMailboxRecipientTypeName' and '$roomMailboxRecipientTypeName': Always Excluded"

	# Exclusion File Summary
	if ($excludedMailboxPrimarySmtpAddresses.Count -gt 0) {
		Write-Host "  - SMTP addresses from exclusion file '$excludeMailboxesFile': Excluded ($excludedByExclusionFileCount matched out of $($excludedMailboxPrimarySmtpAddresses.Count) listed)"
		WriteLog -Message "  - SMTP addresses from exclusion file '$excludeMailboxesFile': Excluded ($excludedByExclusionFileCount matched out of $($excludedMailboxPrimarySmtpAddresses.Count) listed)"
	} else {
		if (Test-Path $excludeMailboxesFile) {
			Write-Host "  - Exclusion file '$excludeMailboxesFile': File was found but was empty, missing column, or no mailboxes matched for exclusion."
			WriteLog -Message "  - Exclusion file '$excludeMailboxesFile': File was found but was empty, missing column, or no mailboxes matched for exclusion."
		} else {
			Write-Host "  - Exclusion file '$excludeMailboxesFile': File not found or not specified, so file-based exclusion was not applied."
			WriteLog -Message "  - Exclusion file '$excludeMailboxesFile': File not found or not specified, so file-based exclusion was not applied."
		}
	}

	# SamAccountName Pattern Exclusion Summary
	if ($PSBoundParameters.ContainsKey('ExcludeSamAccountNamePatterns') -and $ExcludeSamAccountNamePatterns.Count -gt 0) {
		Write-Host "  - SamAccountNames matching patterns ('$($ExcludeSamAccountNamePatterns -join "', '")'): Excluded ($excludedBySamAccountNamePatternCount matched)"
		WriteLog -Message "  - SamAccountNames matching patterns ('$($ExcludeSamAccountNamePatterns -join "', '")'): Excluded ($excludedBySamAccountNamePatternCount matched)"
	} else {
		Write-Host "  - SamAccountName pattern exclusion: Not specified or no patterns provided."
		WriteLog -Message "  - SamAccountName pattern exclusion: Not specified or no patterns provided."
	}

	# New LastLogonTime Criteria Summary
	if ($IncludeSpecificLastLogonCriteria) {
		Write-Host -ForegroundColor Cyan "  --- Specific LastLogonTime Criteria Active ---"
		WriteLog -Message "  --- Specific LastLogonTime Criteria Active ---"
		Write-Host "    - Mailboxes with '$accountDisabledColumnNameInCsv = True': Included"
		WriteLog -Message "    - Mailboxes with '$accountDisabledColumnNameInCsv = True': Included (Actual included: $includedByNewLastLogonCriteria_AcctDisabledTrue)"
		Write-Host "    - Mailboxes with '$accountDisabledColumnNameInCsv = False' AND '$lastLogonTimeColumnNameInCsv > $($lastLogonDateThreshold.ToString('yyyy-MM-dd'))': Included"
		WriteLog -Message "    - Mailboxes with '$accountDisabledColumnNameInCsv = False' AND '$lastLogonTimeColumnNameInCsv > $($lastLogonDateThreshold.ToString('yyyy-MM-dd'))': Included (Actual included: $includedByNewLastLogonCriteria_AcctDisabledFalseLastLogonOK)"
		Write-Host "    - Other mailboxes (AccountDisabled=False, LastLogonTime not meeting criteria, or missing/unparseable data): Excluded ($($excludedByNewLastLogonCriteriaCount + $excludedByMissingColumnsForNewCriteriaCount) total)"
		WriteLog -Message "    - Other mailboxes (AccountDisabled=False, LastLogonTime not meeting criteria, or missing/unparseable data): Excluded (Details: $excludedByNewLastLogonCriteriaCount by date, $excludedByMissingColumnsForNewCriteriaCount by missing columns)"
		Write-Host "    (Original disabled account/mailbox filters ExcludeAllDisabledAccounts, ExcludeDisabledAccountsExceptForSharedMailboxes, ExcludeDisabledMailboxes were IGNORED)"
		WriteLog -Message "    (Original disabled account/mailbox filters ExcludeAllDisabledAccounts, ExcludeDisabledAccountsExceptForSharedMailboxes, ExcludeDisabledMailboxes were IGNORED)"
	} else {
		# Original Disabled Account Summary
		Write-Host -ForegroundColor Cyan "  --- Original Disabled Account/Mailbox Filters Active ---"
		WriteLog -Message "  --- Original Disabled Account/Mailbox Filters Active ---"
		if ($ExcludeAllDisabledAccounts) {
			Write-Host "  - Disabled accounts ($accountDisabledColumnNameInCsv = True): All Excluded (Parameter ExcludeAllDisabledAccounts = `$true)"
			WriteLog -Message "  - Disabled accounts ($accountDisabledColumnNameInCsv = True): All Excluded (Parameter ExcludeAllDisabledAccounts = `$true)"
		} elseif ($ExcludeDisabledAccountsExceptForSharedMailboxes) {
			Write-Host "  - Disabled accounts ($accountDisabledColumnNameInCsv = True): Excluded, EXCEPT for SharedMailboxes (Parameter ExcludeDisabledAccountsExceptForSharedMailboxes = `$true)"
			WriteLog -Message "  - Disabled accounts ($accountDisabledColumnNameInCsv = True): Excluded, EXCEPT for SharedMailboxes (Parameter ExcludeDisabledAccountsExceptForSharedMailboxes = `$true)"
		} else {
			Write-Host "  - Disabled accounts ($accountDisabledColumnNameInCsv = True): Included (Both Exclude*DisabledAccount* parameters are `$false)"
			WriteLog -Message "  - Disabled accounts ($accountDisabledColumnNameInCsv = True): Included (Both Exclude*DisabledAccount* parameters are `$false)"
		}

		# Original Disabled Mailbox Summary
		if ($ExcludeDisabledMailboxes) {
			Write-Host "  - Disabled mailboxes ($isMailboxEnabledColumnNameInCsv = False): Excluded (Parameter ExcludeDisabledMailboxes = `$true)"
			WriteLog -Message "  - Disabled mailboxes ($isMailboxEnabledColumnNameInCsv = False): Excluded (Parameter ExcludeDisabledMailboxes = `$true)"
		} else {
			Write-Host "  - Disabled mailboxes ($isMailboxEnabledColumnNameInCsv = False): Included (Parameter ExcludeDisabledMailboxes = `$false)"
			WriteLog -Message "  - Disabled mailboxes ($isMailboxEnabledColumnNameInCsv = False): Included (Parameter ExcludeDisabledMailboxes = `$false)"
		}
	}
	Write-Host

	Write-Host -ForegroundColor Yellow ("-" * ($host.UI.RawUI.WindowSize.Width -1))
	Write-Host -ForegroundColor Yellow "Input CSV Files Statistics (MailboxesAfterFiltering shows items eligible for batching from that source file):"
	WriteLog -Message "Input CSV Files Statistics (MailboxesAfterFiltering shows items eligible for batching from that source file):"

	$inputCsvSummaryTableFormat = @(
		@{Label="File Name"; Expression={$_.FileName}; Align="Left"; Width=35},
		@{Label="Total In CSV"; Expression={$_.MailboxesBeforeFiltering}; Align="Right"},
		@{Label="Disabled Acc."; Expression={$_.DisabledAccountsInFile}; Align="Right"}, # Raw count from file
		@{Label="Mbx Disabled"; Expression={$_.DisabledMailboxesInFile}; Align="Right"}, # Raw count from file
		@{Label="Post-Filter"; Expression={$_.MailboxesAfterFiltering}; Align="Right"}, # Count AFTER all filters
		@{Label="Status"; Expression={$_.Status}; Align="Left"; Width=30}
	)

	$summaryDataForDisplayTable = [System.Collections.Generic.List[PSObject]]::new() # Use Generic List

	if ($inputFileImportStatistics.Count -gt 0) {
		$inputFileImportStatistics.ForEach({ $summaryDataForDisplayTable.Add($_) })

		# Ensure properties exist and are numeric before summing
		$totalMailboxesBeforeFilteringInStats = ($inputFileImportStatistics | Where-Object {$_.MailboxesBeforeFiltering -is [int] -or $_.MailboxesBeforeFiltering -is [double]} | Measure-Object -Property MailboxesBeforeFiltering -Sum).Sum
		$totalDisabledAccountsInStats = ($inputFileImportStatistics | Where-Object {$_.DisabledAccountsInFile -is [int] -or $_.DisabledAccountsInFile -is [double]} | Measure-Object -Property DisabledAccountsInFile -Sum).Sum
		$totalDisabledMailboxesInStats = ($inputFileImportStatistics | Where-Object {$_.DisabledMailboxesInFile -is [int] -or $_.DisabledMailboxesInFile -is [double]} | Measure-Object -Property DisabledMailboxesInFile -Sum).Sum
		$totalMailboxesAfterFilteringInStats = ($inputFileImportStatistics | Where-Object {$_.MailboxesAfterFiltering -is [int] -or $_.MailboxesAfterFiltering -is [double]} | Measure-Object -Property MailboxesAfterFiltering -Sum).Sum

		$totalRowForFileStatsSummary = [PSCustomObject]@{
			FileName                 = "TOTALS"
			MailboxesBeforeFiltering = $totalMailboxesBeforeFilteringInStats
			DisabledAccountsInFile   = $totalDisabledAccountsInStats
			DisabledMailboxesInFile  = $totalDisabledMailboxesInStats
			MailboxesAfterFiltering  = $totalMailboxesAfterFilteringInStats
			Status                   = "---"
		}
		$summaryDataForDisplayTable.Add($totalRowForFileStatsSummary)
	} else {
		Write-Host "No CSV file statistics to display as no files were processed or statistics collected."
		WriteLog -Message "No CSV file statistics to display as no files were processed or statistics collected."
	}

	if ($summaryDataForDisplayTable.Count -gt 0) {
		$summaryTableOutputString = $summaryDataForDisplayTable | Format-Table $inputCsvSummaryTableFormat -AutoSize | Out-String
		Write-Host $summaryTableOutputString
		Write-Host -ForegroundColor Yellow ("-" * ($host.UI.RawUI.WindowSize.Width -1))

		$summaryTableOutputString.Split([Environment]::NewLine) | ForEach-Object {
			if (-not [string]::IsNullOrWhiteSpace($_)) {
				WriteLog -Message  $_
			}
		}
	}

	$unfilteredOutputCsvFileName = Join-Path -Path $outputFolderForConsolidatedCsv -ChildPath "${outputFileNamePrefix}_AllInputMailboxes_UnfilteredOriginal_$TimeStamp.csv"
	Write-Host -ForegroundColor Magenta "Generating consolidated CSV of all originally loaded mailboxes (unfiltered, original columns)..."
	WriteLog -Message "Generating consolidated CSV of all originally loaded mailboxes (unfiltered, original columns) to '$unfilteredOutputCsvFileName'..."
	if ($allMailboxesOriginalData.Count -gt 0) {
		try {
			# Export directly from the List
			Export-CsvAtomic -InputObject $allMailboxesOriginalData -Path $unfilteredOutputCsvFileName -Encoding UTF8
			Write-Host -ForegroundColor Green "Successfully exported $($allMailboxesOriginalData.Count) mailboxes to unfiltered consolidated CSV (original data): $unfilteredOutputCsvFileName"
			WriteLog -Message "Successfully exported $($allMailboxesOriginalData.Count) mailboxes to unfiltered consolidated CSV (original data): $unfilteredOutputCsvFileName"
		} catch {
			$errorMessage = "Error exporting unfiltered consolidated CSV (original data) to '$unfilteredOutputCsvFileName': $($_.Exception.Message)"
			Write-Error $errorMessage
			WriteLog -Message $errorMessage "ERROR"
		}
	} else {
		Write-Warning "No mailboxes were loaded from input files, so the unfiltered consolidated CSV (original data) was not generated."
		WriteLog -Message "No mailboxes were loaded from input files, so the unfiltered consolidated CSV (original data) was not generated."
	}

	Write-Host -ForegroundColor Magenta "Preparing to generate migration batch CSV files..."
	WriteLog -Message "Preparing to generate migration batch CSV files (Max items per batch: $MaxMailboxesPerBatchFile, Max size per batch: $maxBatchSizeMB MB)..."

	$mailboxToAssignedBatchNameMap = @{} # Map to store PrimarySmtpAddress -> BatchName assigned by THIS script run
	$globalBatchesGeneratedSoFarCount = 0 # Global counter for progress

	if ($filteredMailboxesForBatches.Count -eq 0) {
		Write-Warning "No mailboxes are eligible for batching after all filters were applied. No migration batch CSV files will be generated."
		WriteLog -Message "No mailboxes are eligible for batching after all filters were applied. No migration batch CSV files will be generated."
	} else {
		WriteLog -Message "Total mailboxes eligible for batching (after all filters): $($filteredMailboxesForBatches.Count)"

		# Group mailboxes by domain
		$mailboxesGroupedByDomain = $filteredMailboxesForBatches | Group-Object -Property $domainNameColumnNameInCsv
		$domainBatchGenerationSummaryData = @{} # Store summary data [DomainName] -> [PSCustomObject]@{MailboxCount=X; BatchCount=Y}

		# Estimate total batches for progress (less accurate now due to size limit, but provides a starting point)
		$estimatedTotalBatchesToGenerateCount = 0
		foreach ($domainGroupToCount in $mailboxesGroupedByDomain) {
			$mailboxesInThisDomainForCount = $domainGroupToCount.Group.Count
			if ($mailboxesInThisDomainForCount -gt 0) {
				 # Initial estimate based on count, size limit might increase this
				 $batchesNeededForThisDomain = [math]::Ceiling($mailboxesInThisDomainForCount / $MaxMailboxesPerBatchFile)
				 $estimatedTotalBatchesToGenerateCount += $batchesNeededForThisDomain
			}
		}
		WriteLog -Message "Estimated total migration batches based on count limit: $estimatedTotalBatchesToGenerateCount (Actual number may be higher due to size limit of $maxBatchSizeMB MB per batch)."
		if ($estimatedTotalBatchesToGenerateCount -gt 0) {
			Write-Host "Preparing to generate migration batch file(s)... (Estimate: $estimatedTotalBatchesToGenerateCount)"
		}

		# Process each domain
		foreach ($domainItemGroup in $mailboxesGroupedByDomain) {
			$currentProcessingDomainName = $domainItemGroup.Name
			# Get mailboxes for the current domain as a modifiable list
			$mailboxesInCurrentDomainList = [System.Collections.Generic.List[PSObject]]::new($domainItemGroup.Group)
			$totalMailboxesInCurrentDomainForBatching = $mailboxesInCurrentDomainList.Count
			$mailboxesProcessedInDomain = 0 # Track mailboxes actually placed in batches for this domain

			if ($totalMailboxesInCurrentDomainForBatching -eq 0) {
				WriteLog -Message "Skipping domain '$currentProcessingDomainName' as it has no mailboxes remaining after filtering." # Message clarification
				$domainBatchGenerationSummaryData[$currentProcessingDomainName] = [PSCustomObject]@{ MailboxCount = 0; BatchCount = 0 }
				continue
			}
			WriteLog -Message "Processing domain '$currentProcessingDomainName' with $totalMailboxesInCurrentDomainForBatching eligible mailboxes for batching."

			$batchCounterForThisDomain = 1 # This will be used for the actual batch number in filename
			$batchesGeneratedForThisDomain = 0 # Track batches actually created for this domain
			$currentBatchMailboxes = [System.Collections.Generic.List[PSObject]]::new()
			$currentBatchTotalSizeMB = 0.0
			

			# Use a while loop to consume mailboxes from the list
			while ($mailboxesInCurrentDomainList.Count -gt 0) {
				$mailboxToAdd = $mailboxesInCurrentDomainList[0] # Peek at the next mailbox
				$mailboxSizeMB = 0.0
				$sizeParseSuccess = $false
				$sizeValueString = $mailboxToAdd.$totalItemSizeColumnNameInCsv

				# Try parsing the mailbox size robustly
				if (-not [string]::IsNullOrWhiteSpace($sizeValueString)) {
					# Remove potential suffixes and trim whitespace before parsing
					$sizeValueString = $sizeValueString -replace 'MB$','' -replace 'GB$','' -replace ',','.' # Handle potential suffixes and commas
					try {
						$sizeValueString = $sizeValueString.Trim()
						# Attempt to parse as double using InvariantCulture first
						if ([double]::TryParse($sizeValueString, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$mailboxSizeMB)) {
							$sizeParseSuccess = $true
						} else {
							# Fallback try with current culture if invariant fails
							if ([double]::TryParse($sizeValueString, [ref]$mailboxSizeMB)) {
								$sizeParseSuccess = $true
							} else {
								WriteLog -Message "Warning: Could not parse '$($mailboxToAdd.$totalItemSizeColumnNameInCsv)' as a number for mailbox '$($mailboxToAdd.$primarySmtpAddressColumnNameInCsv)'. Assuming size 0 MB for batch calculation."
								$mailboxSizeMB = 0.0 # Explicitly set to 0 if parse fails
							}
						}
					} catch {
						WriteLog -Message "Warning: Error parsing size '$($mailboxToAdd.$totalItemSizeColumnNameInCsv)' for mailbox '$($mailboxToAdd.$primarySmtpAddressColumnNameInCsv)'. Assuming size 0 MB. Error: $($_.Exception.Message)"
						$mailboxSizeMB = 0.0 # Explicitly set to 0 on error
					}
				} else {
					 WriteLog -Message "Warning: Missing or empty size value ('$totalItemSizeColumnNameInCsv') for mailbox '$($mailboxToAdd.$primarySmtpAddressColumnNameInCsv)'. Assuming size 0 MB for batch calculation."
					 $mailboxSizeMB = 0.0 # Explicitly set to 0 if missing/empty
				}

				# Decision logic: Can this mailbox be added?
				$canAddMailboxToCurrentBatch = $false
				$mustFinalizeCurrentBatchFirst = $false

				if ($currentBatchMailboxes.Count -eq 0) {
					# If batch is empty, always try to add the mailbox.
					# Handle edge case: Single mailbox exceeds the limit
					if ($mailboxSizeMB -gt $maxBatchSizeMB) {
						WriteLog -Message "Warning: Mailbox '$($mailboxToAdd.$primarySmtpAddressColumnNameInCsv)' (Size: $([math]::Round($mailboxSizeMB, 2)) MB) exceeds the maximum batch size ($maxBatchSizeMB MB) and will be placed in its own batch."
						$canAddMailboxToCurrentBatch = $true
						# This will be added, and then the batch will be finalized immediately after.
					} else {
						# It fits in an empty batch (or is 0 size)
						$canAddMailboxToCurrentBatch = $true
					}
				} else {
					# Batch is not empty, check limits
					if (($currentBatchMailboxes.Count -lt $MaxMailboxesPerBatchFile) -and (($currentBatchTotalSizeMB + $mailboxSizeMB) -le $maxBatchSizeMB)) {
						# Fits based on count and size
						$canAddMailboxToCurrentBatch = $true
					} else {
						# Doesn't fit, finalize the current batch first, then this mailbox will start a new batch.
						$mustFinalizeCurrentBatchFirst = $true
					}
				}

				# Action: Finalize Batch (if needed because it's full or a single large item is next)
				if ($mustFinalizeCurrentBatchFirst -and $currentBatchMailboxes.Count -gt 0) {
					$batchesGeneratedForThisDomain++ 
					$globalBatchesGeneratedSoFarCount++
					$batchNameGeneratedByScript = "$outputFileNamePrefix-$($currentProcessingDomainName.Replace('.', '-'))-$TimeStamp-Batch$batchesGeneratedForThisDomain"
					$outputBatchFilePath = Join-Path -Path $outputFolderForBatchCsv -ChildPath "$batchNameGeneratedByScript.csv"

					if ($estimatedTotalBatchesToGenerateCount -gt 0) {
						$progressCompletionPercent = ($globalBatchesGeneratedSoFarCount / $estimatedTotalBatchesToGenerateCount) * 100
						if ($progressCompletionPercent -gt 100) { $progressCompletionPercent = 100 } 
						$statusMessageForProgress = "Batch {0} (Est. Total: {1}): {2}" -f $globalBatchesGeneratedSoFarCount, $estimatedTotalBatchesToGenerateCount, (Split-Path $outputBatchFilePath -Leaf)
						Write-Progress -Activity "Generating Migration Batches" -Status $statusMessageForProgress -PercentComplete $progressCompletionPercent -CurrentOperation "Domain: $currentProcessingDomainName (Batch $batchesGeneratedForThisDomain)"
					}

					$currentBatchMailboxes | ForEach-Object {
						$mbxInBatch = $_
						$mbxInBatch.$batchNameColumn = $batchNameGeneratedByScript 
						if (-not $mailboxToAssignedBatchNameMap.ContainsKey($mbxInBatch.$primarySmtpAddressColumnNameInCsv)) {
							$mailboxToAssignedBatchNameMap[$mbxInBatch.$primarySmtpAddressColumnNameInCsv] = $batchNameGeneratedByScript
						} else { WriteLog -Message "Warning: Mailbox '$($mbxInBatch.$primarySmtpAddressColumnNameInCsv)' re-assigned to batch '$batchNameGeneratedByScript'." }
					}

					$columnsToSelectForBatchCsv = if ($SimpleBatchCsv) { @($primarySmtpAddressColumnNameInCsv) }
												   else { @($primarySmtpAddressColumnNameInCsv, $displayNameColumnNameInCsv, $aliasColumnNameInCsv, $samAccountNameColumnNameInCsv, $ouColumnNameInCsv, $identityColumnNameInCsv, $recipientTypeColumnNameInCsv, $fullAccessColumnNameInCsv, $sendAsColumnNameInCsv, $sendOnBehalfToColumnNameInCsv, $totalItemSizeColumnNameInCsv, $batchNameColumn, $originalSourceFileColumnAddedByScript) }
					try {
						if ($EnabledBatchFileCreation) { Export-CsvAtomic -InputObject ($currentBatchMailboxes | Select-Object $columnsToSelectForBatchCsv) -Path $outputBatchFilePath -Encoding UTF8 }
						WriteLog -Message ("Generated batch file: $outputBatchFilePath with $($currentBatchMailboxes.Count) mailboxes (Total Size: $([math]::Round($currentBatchTotalSizeMB, 2)) MB). Batch Name assigned: $batchNameGeneratedByScript.")
					} catch { Write-Error "Error generating batch file '$outputBatchFilePath': $($_.Exception.Message)"; WriteLog -Message "Error generating batch file '$outputBatchFilePath': $($_.Exception.Message)" }

					$currentBatchMailboxes.Clear()
					$currentBatchTotalSizeMB = 0.0
					# The mailbox $mailboxToAdd that triggered this finalization has NOT been added yet. It will be added to the new, now empty, batch.
				}

				# Action: Add Mailbox to current (possibly new) batch
				if ($canAddMailboxToCurrentBatch) {
					$currentBatchMailboxes.Add($mailboxToAdd)
					$currentBatchTotalSizeMB += $mailboxSizeMB 
					$mailboxesProcessedInDomain++ 
					$mailboxesInCurrentDomainList.RemoveAt(0) # Consume the mailbox from the domain list

					# If this mailbox is a single large item that filled an empty batch, or if it's the last item for the domain and fills the batch limits
					$isSingleLargeItemBatch = ($currentBatchMailboxes.Count -eq 1 -and $mailboxSizeMB -gt $maxBatchSizeMB)
					$isBatchFullByCount = ($currentBatchMailboxes.Count -eq $MaxMailboxesPerBatchFile)
					$isBatchFullBySize = ($currentBatchTotalSizeMB -ge $maxBatchSizeMB) # Using -ge in case of exact match

					if ($isSingleLargeItemBatch -or $isBatchFullByCount -or $isBatchFullBySize) {
						# Finalize this batch immediately
						$batchesGeneratedForThisDomain++
						$globalBatchesGeneratedSoFarCount++
						$batchNameGeneratedByScript = "$outputFileNamePrefix-$($currentProcessingDomainName.Replace('.', '-'))-$TimeStamp-Batch$batchesGeneratedForThisDomain"
						$outputBatchFilePath = Join-Path -Path $outputFolderForBatchCsv -ChildPath "$batchNameGeneratedByScript.csv"

						if ($estimatedTotalBatchesToGenerateCount -gt 0) { $progressCompletionPercent = ($globalBatchesGeneratedSoFarCount / $estimatedTotalBatchesToGenerateCount) * 100; if ($progressCompletionPercent -gt 100) { $progressCompletionPercent = 100 }; $statusMessageForProgress = "Batch {0} (Est. Total: {1}): {2}" -f $globalBatchesGeneratedSoFarCount, $estimatedTotalBatchesToGenerateCount, (Split-Path $outputBatchFilePath -Leaf); Write-Progress -Activity "Generating Migration Batches" -Status $statusMessageForProgress -PercentComplete $progressCompletionPercent -CurrentOperation "Domain: $currentProcessingDomainName (Batch $batchesGeneratedForThisDomain)"}
						
						$currentBatchMailboxes | ForEach-Object { $mbxInBatch = $_; $mbxInBatch.$batchNameColumn = $batchNameGeneratedByScript; if (-not $mailboxToAssignedBatchNameMap.ContainsKey($mbxInBatch.$primarySmtpAddressColumnNameInCsv)) { $mailboxToAssignedBatchNameMap[$mbxInBatch.$primarySmtpAddressColumnNameInCsv] = $batchNameGeneratedByScript } else { WriteLog -Message "Warning: Mailbox '$($mbxInBatch.$primarySmtpAddressColumnNameInCsv)' re-assigned to batch '$batchNameGeneratedByScript'." } }
						
						$columnsToSelectForBatchCsv = if ($SimpleBatchCsv) { @($primarySmtpAddressColumnNameInCsv) } else { @($primarySmtpAddressColumnNameInCsv, $displayNameColumnNameInCsv, $aliasColumnNameInCsv, $samAccountNameColumnNameInCsv, $ouColumnNameInCsv, $identityColumnNameInCsv, $recipientTypeColumnNameInCsv, $fullAccessColumnNameInCsv, $sendAsColumnNameInCsv, $sendOnBehalfToColumnNameInCsv, $totalItemSizeColumnNameInCsv, $batchNameColumn, $originalSourceFileColumnAddedByScript) }
						
						try {
							if ($EnabledBatchFileCreation) { Export-CsvAtomic -InputObject ($currentBatchMailboxes | Select-Object $columnsToSelectForBatchCsv) -Path $outputBatchFilePath -Encoding UTF8; WriteLog -Message ("Generated batch file (due to limit/large item): $outputBatchFilePath with $($currentBatchMailboxes.Count) mailbox(es) (Total Size: $([math]::Round($currentBatchTotalSizeMB, 2)) MB). Batch Name assigned: $batchNameGeneratedByScript.") }
							}
						catch {
							Write-Error "Error generating batch file (due to limit/large item) '$outputBatchFilePath': $($_.Exception.Message)"; WriteLog -Message "Error generating batch file (due to limit/large item) '$outputBatchFilePath': $($_.Exception.Message)"
							}
						$currentBatchMailboxes.Clear()
						$currentBatchTotalSizeMB = 0.0
					}
				} elseif (-not $mustFinalizeCurrentBatchFirst) {
					# This case should ideally not be hit if logic is correct.
					# It means mailbox couldn't be added, but batch wasn't marked for finalization.
					# This could happen if $canAddMailboxToCurrentBatch was false from the start (e.g. failed a check not related to size/count in an empty batch - not current logic)
					WriteLog -Message "Internal Logic Warning: Mailbox '$($mailboxToAdd.$primarySmtpAddressColumnNameInCsv)' was not added and batch finalization was not triggered. Check logic. Skipping mailbox to prevent loop."
					$mailboxesInCurrentDomainList.RemoveAt(0) # Consume (skip) the mailbox
				}
				# If $mustFinalizeCurrentBatchFirst was true, $mailboxToAdd was not added in this iteration,
				# but the previous batch was finalized. $mailboxToAdd will be re-evaluated for the new empty batch in the next iteration.
			} # End while loop processing mailboxes in domain

			# Finalize the very last batch for the domain if it contains mailboxes
			 if ($currentBatchMailboxes.Count -gt 0) {
				$batchesGeneratedForThisDomain++ 
				$globalBatchesGeneratedSoFarCount++
				$batchNameGeneratedByScript = "$outputFileNamePrefix-$($currentProcessingDomainName.Replace('.', '-'))-$TimeStamp-Batch$batchesGeneratedForThisDomain"
				$outputBatchFilePath = Join-Path -Path $outputFolderForBatchCsv -ChildPath "$batchNameGeneratedByScript.csv"

				if ($estimatedTotalBatchesToGenerateCount -gt 0) { $progressCompletionPercent = ($globalBatchesGeneratedSoFarCount / $estimatedTotalBatchesToGenerateCount) * 100; if ($progressCompletionPercent -gt 100) { $progressCompletionPercent = 100 }; $statusMessageForProgress = "Batch {0} (Est. Total: {1}): {2}" -f $globalBatchesGeneratedSoFarCount, $estimatedTotalBatchesToGenerateCount, (Split-Path $outputBatchFilePath -Leaf); Write-Progress -Activity "Generating Migration Batches" -Status $statusMessageForProgress -PercentComplete $progressCompletionPercent -CurrentOperation "Domain: $currentProcessingDomainName (Batch $batchesGeneratedForThisDomain)"}

				$currentBatchMailboxes | ForEach-Object { $mbxInBatch = $_; $mbxInBatch.$batchNameColumn = $batchNameGeneratedByScript; if (-not $mailboxToAssignedBatchNameMap.ContainsKey($mbxInBatch.$primarySmtpAddressColumnNameInCsv)) { $mailboxToAssignedBatchNameMap[$mbxInBatch.$primarySmtpAddressColumnNameInCsv] = $batchNameGeneratedByScript } else { WriteLog -Message "Warning: Mailbox '$($mbxInBatch.$primarySmtpAddressColumnNameInCsv)' re-assigned to batch '$batchNameGeneratedByScript'." } }

				$columnsToSelectForBatchCsv = if ($SimpleBatchCsv) { @($primarySmtpAddressColumnNameInCsv) } else { @($primarySmtpAddressColumnNameInCsv, $displayNameColumnNameInCsv, $aliasColumnNameInCsv, $samAccountNameColumnNameInCsv, $ouColumnNameInCsv, $identityColumnNameInCsv, $recipientTypeColumnNameInCsv, $fullAccessColumnNameInCsv, $sendAsColumnNameInCsv, $sendOnBehalfToColumnNameInCsv, $totalItemSizeColumnNameInCsv, $batchNameColumn, $originalSourceFileColumnAddedByScript) }
				
				try {
					if ($EnabledBatchFileCreation) { Export-CsvAtomic -InputObject ($currentBatchMailboxes | Select-Object $columnsToSelectForBatchCsv) -Path $outputBatchFilePath -Encoding UTF8; WriteLog -Message  ("Generated FINAL batch file for domain '$currentProcessingDomainName': $outputBatchFilePath with $($currentBatchMailboxes.Count) mailboxes (Total Size: $([math]::Round($currentBatchTotalSizeMB, 2)) MB). Batch Name assigned: $batchNameGeneratedByScript.") }
					}
				catch { Write-Error "Error generating FINAL batch file '$outputBatchFilePath': $($_.Exception.Message)"; WriteLog -Message "Error generating FINAL batch file '$outputBatchFilePath': $($_.Exception.Message)" }
			}

			# Record summary for the domain using the actual counts
			$domainBatchGenerationSummaryData[$currentProcessingDomainName] = [PSCustomObject]@{
				MailboxCount = $mailboxesProcessedInDomain # Mailboxes successfully placed in batches
				BatchCount   = $batchesGeneratedForThisDomain # Batches actually created
			}
			WriteLog -Message "Domain '$currentProcessingDomainName': Processed $mailboxesProcessedInDomain mailboxes into $batchesGeneratedForThisDomain batch(es)."

		} # End foreach domain group

		# Complete the progress bar
		Write-Progress -Activity "Generating Migration Batches" -Completed

		# --- Update and Export Consolidated CSV with Batch Names ---
		WriteLog -Message "Preparing data for consolidated CSV (all source mailboxes, BatchName column updated by script) to '$outputConsolidatedCsvPath'..."
		if ($allMailboxesOriginalData.Count -gt 0) {
			# Iterate through the original list and update BatchName using the map
			$allMailboxesOriginalData | ForEach-Object {
				$currentMailboxObject = $_
				# Ensure BatchName column exists (should have been added during import)
				if (-not ($currentMailboxObject.PSObject.Properties.Name -contains $batchNameColumn)) {
					$currentMailboxObject | Add-Member -MemberType NoteProperty -Name $batchNameColumn -Value ""
				}

				# Update BatchName if it was assigned during this run
				if ($mailboxToAssignedBatchNameMap.ContainsKey($currentMailboxObject.$primarySmtpAddressColumnNameInCsv)) {
					$currentMailboxObject.$batchNameColumn = $mailboxToAssignedBatchNameMap[$currentMailboxObject.$primarySmtpAddressColumnNameInCsv]
				}
				# Else: leave the BatchName as it was (empty or potentially a value from the source CSV if not overwritten)
			}

			try {            
				Export-CsvAtomic -InputObject $allMailboxesOriginalData -Path $outputConsolidatedCsvPath -Encoding UTF8
								
				Write-Host -ForegroundColor Green "Successfully exported $($allMailboxesOriginalData.Count) total mailboxes to consolidated CSV (all source columns, BatchName updated by script): $outputConsolidatedCsvPath"
				WriteLog -Message "Successfully exported $($allMailboxesOriginalData.Count) total mailboxes to consolidated CSV (all source columns, BatchName updated by script): $outputConsolidatedCsvPath"

				# Copy the generated CSV to the second path
				Copy-Item -Path $outputConsolidatedCsvPath -Destination $outputConsolidatedCsvPath2 -Force -ErrorAction Stop
				$outputFolderForConsolidatedCsv3 = (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '')
            $outputConsolidatedCsvPath3 = Join-Path -Path $outputFolderForConsolidatedCsv3 -ChildPath "Exchange_OnPrem_Mailboxes_AllDomains.csv"
				Copy-Item -Path $outputConsolidatedCsvPath -Destination $outputConsolidatedCsvPath3 -Force -ErrorAction Stop
				Invoke-SmartM365SharePointCsvUpload -LocalFilePath $outputConsolidatedCsvPath3
				$SendFileListEmailReportFileName = $outputConsolidatedCsvPath3

				Write-Host -ForegroundColor Green "Successfully copied consolidated CSV to: $outputConsolidatedCsvPath2"
				WriteLog -Message "Successfully copied consolidated CSV to: $outputConsolidatedCsvPath2"

			} catch {
				$errorMessage = "Error exporting consolidated CSV (all source mailboxes, BatchName updated by script) to '$outputConsolidatedCsvPath': $($_.Exception.Message)"
				Write-Error $errorMessage
				WriteLog -Message $errorMessage "ERROR"
			}
		} else {
			Write-Warning "No mailboxes were originally loaded, so the consolidated CSV (all source mailboxes, BatchName updated by script) was not generated."
			WriteLog -Message "No mailboxes were originally loaded, so the consolidated CSV (all source mailboxes, BatchName updated by script) was not generated."
		}
	} # End else (mailboxes were eligible for batching)

	Write-Host
	Write-Host -ForegroundColor Magenta "Batch Generation Summary by Domain:"
	WriteLog -Message "Batch Generation Summary by Domain:"

	if ($domainBatchGenerationSummaryData -and $domainBatchGenerationSummaryData.Count -gt 0) {
		$summaryObjectsForDomainTable = [System.Collections.Generic.List[PSObject]]::new() # Use Generic List

		# Sort domains by name for consistent output
		$domainBatchGenerationSummaryData.GetEnumerator() | Sort-Object Name | ForEach-Object {
			$summaryObjectsForDomainTable.Add([PSCustomObject]@{
				Domain              = $_.Name
				'Batched Mailboxes' = $_.Value.MailboxCount # Renamed for clarity
				'Batches Generated' = $_.Value.BatchCount
			})
		}

		if ($summaryObjectsForDomainTable.Count -gt 0) {
			# Calculate totals from the generated summary objects
			$totalBatchedMailboxesInSummary = ($summaryObjectsForDomainTable | Measure-Object -Property 'Batched Mailboxes' -Sum).Sum
			$totalBatchesGeneratedInSummary = ($summaryObjectsForDomainTable | Measure-Object -Property 'Batches Generated' -Sum).Sum

			$totalSummaryRowObject = [PSCustomObject]@{
				Domain              = "TOTAL"
				'Batched Mailboxes' = $totalBatchedMailboxesInSummary
				'Batches Generated' = $totalBatchesGeneratedInSummary
			}
			$summaryObjectsForDomainTable.Add($totalSummaryRowObject) # Add total row at the end
		}

		$summaryTableOutputString = $summaryObjectsForDomainTable | Format-Table -AutoSize | Out-String
		Write-Host $summaryTableOutputString

		# Log the summary table
		$summaryTableOutputString.Split([Environment]::NewLine) | ForEach-Object {
			if (-not [string]::IsNullOrWhiteSpace($_)) {
				WriteLog -Message  $_
			}
		}
	} else {
		$noSummaryMessageText = "No batches were generated, so no domain summary table is available."
		Write-Host $noSummaryMessageText
		WriteLog -Message  $noSummaryMessageText
	}
	try { Write-Host ('-' * ($host.UI.RawUI.WindowSize.Width - 1)) } catch {}

	$EndTime = Get-Date
	Write-Host -ForegroundColor White "Script finished. Total execution time: $($EndTime - $StartTime)."
	WriteLog -Message "Script finished. Total execution time: $($EndTime - $StartTime)."
	
}
if ($InventoryCompletedSuccessfully -eq $false) { 
	$interruptionMessageRedundant = "Script (outer finally) interrupted or terminated due to an error. Total duration: $($EndTimeFinalRedundant - $StartTime)."
	Write-Host -ForegroundColor Yellow $interruptionMessageRedundant
}
Else
{
	#SendFileListEmailReport -Files @($SendFileListEmailReportFileName) -Title "$TaskName" -Message "$TaskName : All processed mailbox data exported to $SendFileListEmailReportFileName."
					
	    #region Run Summary
    try {
        $EndTimeSummary = Get-Date
        $durationSummary = $EndTimeSummary - $StartTime

        $includedOuText = ""
        if ($IncludedOrganizationalUnit -and $IncludedOrganizationalUnit.Count -gt 0) {
            $includedOuText = ($IncludedOrganizationalUnit -join "; ")
        }

        $mailboxRecordCount = $null
        if ($Global:ScriptOverallMailboxData) {
            $mailboxRecordCount = $Global:ScriptOverallMailboxData.Count
        }

        Write-Host
        Write-Host -ForegroundColor Cyan "Run Summary:"
        WriteLog -Message "Run Summary:"

        Write-Host "  ScriptVersion     : $ScriptVersion"
        Write-Host "  StartTime         : $($StartTime.ToString('o'))"
        Write-Host "  EndTime           : $($EndTimeSummary.ToString('o'))"
        Write-Host "  Duration          : $durationSummary"
        Write-Host "  Host              : $env:COMPUTERNAME"
        Write-Host "  User              : $env:USERNAME"
        Write-Host "  OnlyADPermission  : $OnlyADPermission"
        Write-Host "  IncludeADPermission: $IncludeADPermission"
        Write-Host "  DetectAllDomains  : $DetectAllDomains"
        Write-Host "  IncludedOUs       : $includedOuText"
        Write-Host "  OutputPath        : $OutputPath"
        if ($SendFileListEmailReportFileName) { Write-Host "  PrimaryOutputFile : $SendFileListEmailReportFileName" }
        if ($mailboxRecordCount -ne $null) { Write-Host "  MailboxRecords    : $mailboxRecordCount" }

        WriteLog -Message ("  ScriptVersion      : {0}" -f $ScriptVersion)
        WriteLog -Message ("  StartTime          : {0}" -f $StartTime.ToString('o'))
        WriteLog -Message ("  EndTime            : {0}" -f $EndTimeSummary.ToString('o'))
        WriteLog -Message ("  Duration           : {0}" -f $durationSummary)
        WriteLog -Message ("  Host               : {0}" -f $env:COMPUTERNAME)
        WriteLog -Message ("  User               : {0}" -f $env:USERNAME)
        WriteLog -Message ("  OnlyADPermission   : {0}" -f $OnlyADPermission)
        WriteLog -Message ("  IncludeADPermission: {0}" -f $IncludeADPermission)
        WriteLog -Message ("  DetectAllDomains   : {0}" -f $DetectAllDomains)
        WriteLog -Message ("  IncludedOUs        : {0}" -f $includedOuText)
        WriteLog -Message ("  OutputPath         : {0}" -f $OutputPath)
        if ($SendFileListEmailReportFileName) { WriteLog -Message ("  PrimaryOutputFile  : {0}" -f $SendFileListEmailReportFileName) }
        if ($mailboxRecordCount -ne $null) { WriteLog -Message ("  MailboxRecords     : {0}" -f $mailboxRecordCount) }
    } catch {
        WriteLog -Message "Failed to build Run Summary: $_" "ERROR"
    }
    #endregion Run Summary
	#region Cleanup
	# Clean up old CSV files + old log files
	# Automatically excludes all generated CSVs via global:csvGeneratedPaths + current transcript and log files via global variables
	RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
	RemoveOldFiles -Path $logPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
	WriteLog -Message "$TaskName completed."
	Stop-Transcript
	#endregion
}
}
#End of script



