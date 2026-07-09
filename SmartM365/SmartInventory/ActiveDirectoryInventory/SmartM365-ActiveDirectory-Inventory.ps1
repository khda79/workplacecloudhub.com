<#
.SYNOPSIS
    Active Directory full inventory (OUs, Computers, Users, Groups, Contacts) across one or multiple domains.

.DESCRIPTION
    This script performs a comprehensive inventory of Active Directory across:
    - Organizational Units (OUs)
    - Computers
    - Users
    - Groups
    - Contacts

    It:
    - Discovers all domains in the forest or uses a subset passed via -TargetDomains
    - Exports detailed CSVs per domain in a "Not-CSV-Combined" folder
    - Combines all per-domain CSVs into global "AllDomains" CSVs
    - Analyzes duplicate UserPrincipalNames and SMTP proxy addresses across all domains
    - Uses the shared framework (SmartM365.Core / InitializeScriptEnvironment)
    - Logs to text + transcript
    - Copies combined CSVs to a local configuration path (LatestCsvFolderPath)
    - Cleans old CSV/log files
    - Sends an email notification in case of a global error (SendEmailHtmlReport)

.VERSION
1.8

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',

    [Parameter(Mandatory = $false)]
    [string[]]$TargetDomains,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [int]$DomainParallelThrottleLimit = 0,

    [Parameter(Mandatory = $false)]
    [switch]$DomainWorker,

    [Parameter(Mandatory = $false)]
    [string]$DomainWorkerTempFolder,

    [Parameter(Mandatory = $false)]
    [switch]$ReportOnly,

    [Parameter(Mandatory = $false)]
    [switch]$DuplicateAnalysisOnly,

    [Parameter(Mandatory = $false)]
    [switch]$ForceSendDuplicateNotification,

    [Parameter(Mandatory = $false)]
    [switch]$SkipDailyReport
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

function Get-SmartM365AdInventorySendMailMode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config
    )

    $configuredMode = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'SendMailMode' -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($configuredMode)) {
        $configuredMode = $configuredMode.Trim()
    }

    if ([string]::IsNullOrWhiteSpace($configuredMode) -or $configuredMode -in @('__USE_GLOBAL__', 'USE_GLOBAL')) {
        $smtpServer = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'SmtpServer' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($smtpServer)) { return 'Graph' }
        return 'SMTP'
    }

    switch ($configuredMode.ToLowerInvariant()) {
        'graph' { return 'Graph' }
        'smtp'  { return 'SMTP' }
        'both'  { return 'Both' }
        default { throw ("Invalid SendMailMode '{0}'. Use Graph, SMTP, or Both." -f $configuredMode) }
    }
}

function Send-SmartM365AdInventoryEmailHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$BodyHtml,
        [string]$To = ''
    )

    $effectiveSendMailMode = Get-SmartM365AdInventorySendMailMode -Config $ScriptLocalConfig
    $mailParams = @{
        Subject      = $Subject
        BodyHtml     = $BodyHtml
        SendMailMode = $effectiveSendMailMode
    }
    if (-not [string]::IsNullOrWhiteSpace($To)) {
        $mailParams['To'] = $To
    }

    SendEmailHtmlReport @mailParams
}



$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$DomainFriendlyNames = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DomainFriendlyNames' -DefaultValue ([pscustomobject]@{})
$IntuneEnrollmentGroupPattern = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'IntuneEnrollmentGroupPattern' -DefaultValue '*GG_INTUNE_ENROLLMENT*'
$Windows11UpgradeGroupPattern = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Windows11UpgradeGroupPattern' -DefaultValue '*GG_INTUNE_UPGRADEW11*'
$EnableOuInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableOuInventory' -DefaultValue $true)
$EnableComputerInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableComputerInventory' -DefaultValue $true)
$EnableUserInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableUserInventory' -DefaultValue $true)
$EnableGroupInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableGroupInventory' -DefaultValue $true)
$EnableContactInventory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableContactInventory' -DefaultValue $true)
$EnableDuplicateAnalysis = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDuplicateAnalysis' -DefaultValue $true)
$EnableDuplicateNotification = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDuplicateNotification' -DefaultValue $true)
$DuplicateNotificationLastSentFilePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DuplicateNotificationLastSentFilePath' -DefaultValue ''
$DeleteTemporaryPerDomainCsv = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeleteTemporaryPerDomainCsv' -DefaultValue $true)
$EnableWeeklyHistory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableWeeklyHistory' -DefaultValue $true)
$WeeklyHistoryFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryFolderPath' -DefaultValue ''
$WeeklyHistoryRetentionWeeks = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)
$EnableDailyReport = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDailyReport' -DefaultValue $true)
$DailyReportAllowedOS = @(Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyReportAllowedOS' -DefaultValue @('Windows 7*','Windows 8*','Windows 10*','Windows 11*'))
$DailyReportInactiveDays = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyReportInactiveDays' -DefaultValue 90)
$EnableDailyReportLock = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDailyReportLock' -DefaultValue $true)
$DailyReportLockRoot = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyReportLockRoot' -DefaultValue 'C:\ProgramData\SmartM365\Locks'
$DailyReportLockName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyReportLockName' -DefaultValue 'SmartM365-ActiveDirectory-Inventory-DailyReport'
$ConfiguredDomainParallelThrottleLimit = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DomainParallelThrottleLimit' -DefaultValue 1)
$EffectiveDomainParallelThrottleLimit = if ($DomainParallelThrottleLimit -gt 0) { $DomainParallelThrottleLimit } else { $ConfiguredDomainParallelThrottleLimit }
if ($EffectiveDomainParallelThrottleLimit -lt 1) { $EffectiveDomainParallelThrottleLimit = 1 }
if ($DomainWorker) { $EffectiveDomainParallelThrottleLimit = 1 }

# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host ("Current PowerShell version: {0}" -f $PSVersionTable.PSVersion) -ForegroundColor Yellow
    exit 1
}

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module $modulePath -ErrorAction Stop
} catch {
    Write-Host ("Failed to import SmartM365.Core module from '{0}' : {1}" -f $modulePath, $_) -ForegroundColor Red
    exit 1
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.8"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ActiveDirectoryInventoryCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:LogTranscriptFile -Append

    WriteLog -Message ("Script environment initialized at {0}" -f $InitializeOutputPath)
    $OutputPath = $InitializeOutputPath
    if ([string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath)) {
        $WeeklyHistoryFolderPath = Join-Path -Path $OutputPath -ChildPath 'WeeklyHistory'
    }
    WriteLog -Message ("Starting {0}" -f $TaskName)
}
catch {
    Write-Host ("Initialization failed: {0}" -f $_) -ForegroundColor Red
    exit 1
}

# ==========================================================
# MAIN TRY / CATCH / FINALLY
# ==========================================================
try {
    if (-not $ReportOnly -and -not $DuplicateAnalysisOnly) {
        Import-Module ActiveDirectory -ErrorAction Stop
        WriteLog -Message "ActiveDirectory module imported successfully."
        Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredModules @('ActiveDirectory') -RequireActiveDirectoryRead | Out-Null
    }
    else {
        Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) | Out-Null
        if ($DuplicateAnalysisOnly) {
            WriteLog -Message "DuplicateAnalysisOnly mode enabled. Live Active Directory inventory collection will be skipped."
        }
        else {
            WriteLog -Message "ReportOnly mode enabled. Live Active Directory inventory collection will be skipped."
        }
    }

    # ----------------------------------------------------------
    # RETRY CONFIGURATION
    # ----------------------------------------------------------
    $MaxRetries        = 3
    $RetryDelaySeconds = 30

    # ----------------------------------------------------------
    # TRANSIENT AD ERROR DETECTION
    # ----------------------------------------------------------
    function Test-IsTransientADError {
        param(
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.ErrorRecord]$ErrorRecord
        )

        $ex = $ErrorRecord.Exception
        while ($null -ne $ex) {
            if ($ex -is [Microsoft.ActiveDirectory.Management.ADServerDownException]) { return $true }
            if ($ex.Message -match 'Unable to contact the server')                    { return $true }
            if ($ex.Message -match 'The server is not operational')                   { return $true }
            if ($ex.Message -match 'invalid enumeration context')                     { return $true }
            $ex = $ex.InnerException
        }
        return $false
    }

    $utcDate     = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $baseFolder  = Join-Path $OutputPath "Not-CSV-Combined"
    $tempFolder = $null
    if ($DuplicateAnalysisOnly) {
        WriteLog -Message "Temporary per-domain export folder skipped in DuplicateAnalysisOnly mode."
    }
    elseif ($DomainWorker -and -not [string]::IsNullOrWhiteSpace($DomainWorkerTempFolder)) {
        $tempFolder = $DomainWorkerTempFolder
    }
    else {
        $tempFolder = Join-Path $baseFolder $utcDate
    }

    if (-not [string]::IsNullOrWhiteSpace($tempFolder)) {
        $null = New-Item -ItemType Directory -Path $tempFolder -Force
        WriteLog -Message ("Temporary per-domain export folder ready: {0}" -f $tempFolder)
    }
    function Get-DomainNameShort {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainName
        )

        if ([string]::IsNullOrWhiteSpace($DomainName)) {
            return $null
        }

        $domainNameLower = $DomainName.Trim().ToLowerInvariant()
        $configuredName = $DomainFriendlyNames.PSObject.Properties[$domainNameLower]
        if ($null -ne $configuredName -and -not [string]::IsNullOrWhiteSpace([string]$configuredName.Value)) {
            return [string]$configuredName.Value
        }

        $dotPos = $domainNameLower.IndexOf('.')
        if ($dotPos -gt 0) {
            return $domainNameLower.Substring(0, $dotPos).ToUpperInvariant()
        }

        return $domainNameLower.ToUpperInvariant()
    }

    function Get-NormalizedDomainAndSam {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainNameShort,

            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$SamAccountName
        )

        $domainPart = if ([string]::IsNullOrWhiteSpace($DomainNameShort)) { '' } else { $DomainNameShort.Trim() }
        $samPart    = if ([string]::IsNullOrWhiteSpace($SamAccountName)) { '' } else { $SamAccountName.Trim() }

        if ($samPart.EndsWith('$')) {
            $samPart = $samPart.Substring(0, $samPart.Length - 1)
        }

        $value = "{0}\{1}" -f $domainPart, $samPart
        $value = $value.Trim().ToLowerInvariant()
        $value = $value -replace '\u00A0', ''
        $value = $value -replace ' ', ''
        $value = $value -replace '\t', ''

        return $value
    }

    function Convert-GuidToImmutableId {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$ObjectGuid
        )
        try {
            if ([string]::IsNullOrWhiteSpace($ObjectGuid)) {
                return $null
            }
            $guid  = [System.Guid]::Parse($ObjectGuid)
            $bytes = $guid.ToByteArray()
            return [System.Convert]::ToBase64String($bytes)
        }
        catch {
            return $null
        }
    }

    function Remove-OldFiles {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,

            [Parameter(Mandatory = $true)]
            [string]$Filter,

            [Parameter(Mandatory = $true)]
            [int]$OlderThanDays
        )

        if (-not (Test-Path -Path $Path)) {
            return
        }

        $limit = (Get-Date).AddDays(-$OlderThanDays)
        Get-ChildItem -Path $Path -Filter $Filter -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $limit } |
            ForEach-Object {
                try {
                    Remove-Item -Path $_.FullName -Force -ErrorAction Stop
                    WriteLog -Message ("Deleted old file: {0}" -f $_.FullName)
                }
                catch {
                    WriteLog -Message ("Failed to delete old file '{0}': {1}" -f $_.FullName, $_)
                }
            }
    }

    function Combine-CsvFiles {
        param(
            [Parameter(Mandatory = $true)]
            [string]$SourceFolder,

            [Parameter(Mandatory = $true)]
            [string]$Filter,

            [Parameter(Mandatory = $true)]
            [string]$DestinationFile
        )

        $files = Get-ChildItem -Path $SourceFolder -Filter $Filter -File | Sort-Object Name
        if (-not $files) {
            if (Test-Path -LiteralPath $DestinationFile) {
                Remove-Item -LiteralPath $DestinationFile -Force -ErrorAction Stop
                WriteLog -Message ("Removed stale combined CSV because no source files were found for filter '{0}': {1}" -f $Filter, $DestinationFile)
            }
            WriteLog -Message ("No CSV files found for filter '{0}' in '{1}'" -f $Filter, $SourceFolder)
            return
        }

        $isFirstFile = $true
        [int64]$combinedRowCount = 0

        foreach ($file in $files) {
            $rows = @(Import-Csv -Path $file.FullName -ErrorAction Stop)
            if ($rows.Count -eq 0) { continue }

            if ($isFirstFile) {
                $rows | Export-Csv -Path $DestinationFile -NoTypeInformation -Encoding UTF8
                $isFirstFile = $false
            }
            else {
                $rows | Export-Csv -Path $DestinationFile -NoTypeInformation -Encoding UTF8 -Append
            }

            $combinedRowCount += $rows.Count
        }

        if ($isFirstFile) {
            if (Test-Path -LiteralPath $DestinationFile) {
                Remove-Item -LiteralPath $DestinationFile -Force -ErrorAction Stop
                WriteLog -Message ("Removed stale combined CSV because no data rows were found for filter '{0}': {1}" -f $Filter, $DestinationFile)
            }
            WriteLog -Message ("No data rows found while combining filter '{0}' in '{1}'" -f $Filter, $SourceFolder)
            return
        }

        WriteLog -Message ("Combined {0} file(s), {1} row(s), into '{2}'" -f $files.Count, $combinedRowCount, $DestinationFile)
    }

    function Save-WeeklyInventoryHistory {
        param(
            [Parameter(Mandatory = $true)]
            [string[]]$SourceFiles,

            [Parameter(Mandatory = $true)]
            [string]$HistoryRootPath,

            [Parameter(Mandatory = $false)]
            [int]$RetentionWeeks = 52
        )

        if ([string]::IsNullOrWhiteSpace($HistoryRootPath)) {
            WriteLog -Message "Weekly AD inventory history skipped: WeeklyHistoryFolderPath is empty."
            return
        }

        $existingSourceFiles = @($SourceFiles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_) })
        if ($existingSourceFiles.Count -eq 0) {
            WriteLog -Message "Weekly AD inventory history skipped: no source CSV file found."
            return
        }

        $now = Get-Date
        $isoYear = [System.Globalization.ISOWeek]::GetYear($now)
        $isoWeek = [System.Globalization.ISOWeek]::GetWeekOfYear($now)
        $weekName = "{0}-W{1:00}" -f $isoYear, $isoWeek
        $weekFolder = Join-Path -Path $HistoryRootPath -ChildPath $weekName

        if (Test-Path -LiteralPath $weekFolder) {
            $existingCsv = @(Get-ChildItem -LiteralPath $weekFolder -Filter '*.csv' -File -ErrorAction SilentlyContinue)
            if ($existingCsv.Count -gt 0) {
                WriteLog -Message ("Weekly AD inventory history already exists for {0}. Snapshot skipped: {1}" -f $weekName, $weekFolder)
                return
            }
        }

        New-Item -Path $weekFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        $copiedFiles = New-Object System.Collections.Generic.List[string]

        foreach ($sourceFile in $existingSourceFiles) {
            $destinationFile = Join-Path -Path $weekFolder -ChildPath ([System.IO.Path]::GetFileName($sourceFile))
            Copy-Item -LiteralPath $sourceFile -Destination $destinationFile -Force -ErrorAction Stop
            [void]$copiedFiles.Add($destinationFile)
        }

        $manifest = [PSCustomObject][ordered]@{
            CreatedAt       = (Get-Date).ToString('o')
            Week            = $weekName
            SourceOutputPath = $OutputPath
            Files           = @($copiedFiles | ForEach-Object { [System.IO.Path]::GetFileName($_) })
        }
        $manifestPath = Join-Path -Path $weekFolder -ChildPath 'manifest.json'
        $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

        WriteLog -Message ("Weekly AD inventory history saved for {0}: {1} file(s) in {2}" -f $weekName, $copiedFiles.Count, $weekFolder)

        if ($RetentionWeeks -gt 0) {
            $oldWeekFolders = @(Get-ChildItem -LiteralPath $HistoryRootPath -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^\d{4}-W\d{2}$' } |
                Sort-Object Name -Descending |
                Select-Object -Skip $RetentionWeeks)

            foreach ($oldWeekFolder in $oldWeekFolders) {
                try {
                    Remove-Item -LiteralPath $oldWeekFolder.FullName -Recurse -Force -ErrorAction Stop
                    WriteLog -Message ("Deleted old weekly AD inventory history folder: {0}" -f $oldWeekFolder.FullName)
                }
                catch {
                    WriteLog -Message ("Failed to delete old weekly AD inventory history folder '{0}': {1}" -f $oldWeekFolder.FullName, $_) -Level "WARNING"
                }
            }
        }
    }

    function Remove-TemporaryInventoryFolder {
        param(
            [Parameter(Mandatory = $true)]
            [string]$TempFolder,

            [Parameter(Mandatory = $true)]
            [string]$BaseFolder
        )

        if (-not $DeleteTemporaryPerDomainCsv) {
            WriteLog -Message ("Temporary per-domain CSV folder kept because DeleteTemporaryPerDomainCsv is disabled: {0}" -f $TempFolder)
            return
        }

        if ([string]::IsNullOrWhiteSpace($TempFolder) -or -not (Test-Path -LiteralPath $TempFolder)) {
            return
        }

        try {
            $resolvedTemp = (Resolve-Path -LiteralPath $TempFolder -ErrorAction Stop).Path.TrimEnd('\')
            $resolvedBase = (Resolve-Path -LiteralPath $BaseFolder -ErrorAction Stop).Path.TrimEnd('\')
            $basePrefix = $resolvedBase + '\'

            if ($resolvedTemp -eq $resolvedBase -or -not $resolvedTemp.StartsWith($basePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                WriteLog -Message ("Temporary per-domain CSV cleanup skipped because the path is outside the expected base folder. Temp: {0}. Base: {1}" -f $resolvedTemp, $resolvedBase) -Level "WARNING"
                return
            }

            Remove-Item -LiteralPath $resolvedTemp -Recurse -Force -ErrorAction Stop
            WriteLog -Message ("Deleted temporary per-domain CSV folder: {0}" -f $resolvedTemp)
        }
        catch {
            WriteLog -Message ("Failed to delete temporary per-domain CSV folder '{0}': {1}" -f $TempFolder, $_) -Level "WARNING"
        }
    }

    function ConvertTo-DailyReportBoolean {
        [CmdletBinding()]
        param([AllowNull()]$Value)

        if ($Value -is [bool]) { return $Value }
        if ($null -eq $Value) { return $false }
        return ([string]$Value).Trim() -eq 'True'
    }

    function ConvertTo-DailyReportDate {
        [CmdletBinding()]
        param([AllowNull()]$Value)

        if ($Value -is [datetime]) { return $Value }
        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

        $parsedDate = [datetime]::MinValue
        if ([datetime]::TryParse([string]$Value, [ref]$parsedDate)) { return $parsedDate }
        return $null
    }

    function Get-DailyReportSimpleOS {
        [CmdletBinding()]
        param([AllowNull()][string]$OperatingSystem)

        switch -Wildcard ($OperatingSystem) {
            '*Windows 11*' { return 'Windows 11' }
            '*Windows 10*' { return 'Windows 10' }
            '*Windows 8*'  { return 'Windows 8' }
            '*Windows 7*'  { return 'Windows 7' }
            default        { return 'Unknown' }
        }
    }

    function Test-DailyReportWildcardMatch {
        [CmdletBinding()]
        param(
            [AllowNull()][string]$Value,
            [string[]]$Patterns
        )

        foreach ($pattern in @($Patterns)) {
            if ($Value -like $pattern) { return $true }
        }
        return $false
    }

    function Get-InventoryDailyReportSource {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$SourceFolder,
            [Parameter(Mandatory = $true)][string[]]$AllowedOS,
            [string[]]$TargetDomains
        )

        if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
            WriteLog -Message 'Daily report source skipped: source folder is empty.' -Level 'WARNING'
            return $null
        }

        $computersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Computers_AllDomains.csv'
        $usersCsv = Join-Path -Path $SourceFolder -ChildPath 'AD_Users_AllDomains.csv'

        foreach ($sourceCsv in @($computersCsv, $usersCsv)) {
            if (-not (Test-Path -LiteralPath $sourceCsv)) {
                WriteLog -Message ("Daily report source unavailable. Missing file: {0}" -f $sourceCsv) -Level 'WARNING'
                return $null
            }
        }

        $targetDomainSet = @{}
        foreach ($domain in @($TargetDomains | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            $targetDomainSet[$domain.ToLowerInvariant()] = $true
        }
        $hasTargetDomainFilter = $targetDomainSet.Count -gt 0

        $computers = @(Import-Csv -LiteralPath $computersCsv -Encoding UTF8 | Where-Object {
            $domainName = [string]$_.DomainName
            ((-not $hasTargetDomainFilter) -or $targetDomainSet.ContainsKey($domainName.ToLowerInvariant())) -and
            (Test-DailyReportWildcardMatch -Value ([string]$_.OperatingSystem) -Patterns $AllowedOS)
        } | ForEach-Object {
            [PSCustomObject]@{
                Enabled                = ConvertTo-DailyReportBoolean $_.Enabled
                SimpleOS               = Get-DailyReportSimpleOS -OperatingSystem ([string]$_.OperatingSystem)
                ADDomain               = [string]$_.DomainName
                OperatingSystem        = [string]$_.OperatingSystem
                OperatingSystemVersion = [string]$_.OperatingSystemVersion
                LastLogonDate          = ConvertTo-DailyReportDate $_.LastLogonDate
            }
        })

        $users = @(Import-Csv -LiteralPath $usersCsv -Encoding UTF8 | Where-Object {
            $domainName = [string]$_.DomainName
            ((-not $hasTargetDomainFilter) -or $targetDomainSet.ContainsKey($domainName.ToLowerInvariant())) -and
            (-not [string]::IsNullOrWhiteSpace([string]$_.UserPrincipalName))
        } | ForEach-Object {
            [PSCustomObject]@{
                ADDomain          = [string]$_.DomainName
                Enabled           = ConvertTo-DailyReportBoolean $_.Enabled
                LastLogonDate     = ConvertTo-DailyReportDate $_.LastLogonDate
                UserPrincipalName = [string]$_.UserPrincipalName
            }
        })

        if ($computers.Count -eq 0 -and $users.Count -eq 0) {
            WriteLog -Message 'Daily report source produced no reportable rows.' -Level 'WARNING'
            return $null
        }

        $domains = @(@($computers | Select-Object -ExpandProperty ADDomain) + @($users | Select-Object -ExpandProperty ADDomain) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique)

        WriteLog -Message ("Daily report source loaded. Computers: {0}; Users: {1}; Domains: {2}" -f $computers.Count, $users.Count, ($domains -join ', '))
        return [PSCustomObject]@{
            Domains   = $domains
            Computers = $computers
            Users     = $users
        }
    }

    function Write-DailyReportCsv {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]$Rows,
            [Parameter(Mandatory = $true)][string]$OutputFilePath,
            [string]$LatestFolderPath
        )

        $rowsArray = @($Rows)
        if ($rowsArray.Count -eq 0) {
            WriteLog -Message ("Daily report CSV skipped because there are no rows: {0}" -f $OutputFilePath) -Level 'WARNING'
            return
        }

        if (Test-Path -LiteralPath $OutputFilePath) {
            $rowsArray | ConvertTo-Csv -NoTypeInformation -Delimiter ';' | Select-Object -Skip 1 | Add-Content -Path $OutputFilePath -Encoding UTF8
            WriteLog -Message ("Daily report rows appended to CSV: {0}" -f $OutputFilePath)
        }
        else {
            $rowsArray | Export-Csv -Path $OutputFilePath -NoTypeInformation -Delimiter ';' -Encoding UTF8
            WriteLog -Message ("Daily report CSV created: {0}" -f $OutputFilePath)
        }
        if (-not $global:csvGeneratedPaths -or -not ($global:csvGeneratedPaths -is [System.Collections.Generic.HashSet[string]])) {
            $existingCsvGeneratedPaths = @($global:csvGeneratedPaths)
            $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($existingCsvGeneratedPath in $existingCsvGeneratedPaths) {
                if (-not [string]::IsNullOrWhiteSpace([string]$existingCsvGeneratedPath)) {
                    [void]$global:csvGeneratedPaths.Add([string]$existingCsvGeneratedPath)
                }
            }
        }
        [void]$global:csvGeneratedPaths.Add($OutputFilePath)
        if (-not [string]::IsNullOrWhiteSpace($LatestFolderPath)) {
            if (-not (Test-Path -LiteralPath $LatestFolderPath)) {
                New-Item -Path $LatestFolderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                WriteLog -Message ("Created missing LatestCsvFolderPath directory: {0}" -f $LatestFolderPath)
            }

            $latestFilePath = Join-Path -Path $LatestFolderPath -ChildPath ([System.IO.Path]::GetFileName($OutputFilePath))
            Copy-Item -LiteralPath $OutputFilePath -Destination $latestFilePath -Force -ErrorAction Stop
            WriteLog -Message ("Daily report CSV copied to LatestCsvFolderPath: {0}" -f $latestFilePath)
        }

        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $OutputFilePath
    }

    function Invoke-ActiveDirectoryDailyReport {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$SourceFolder,
            [Parameter(Mandatory = $true)][string]$ReportOutputPath,
            [string]$LatestFolderPath,
            [string[]]$AllowedOS = @('Windows 7*','Windows 8*','Windows 10*','Windows 11*'),
            [string[]]$TargetDomains,
            [int]$InactiveDays = 90,
            [bool]$UseDailyLock = $true,
            [string]$LockRoot,
            [string]$LockName
        )

        if ($SkipDailyReport) {
            WriteLog -Message 'Active Directory daily report skipped because -SkipDailyReport was specified.'
            return $false
        }

        if (-not $EnableDailyReport) {
            WriteLog -Message 'Active Directory daily report skipped because EnableDailyReport is disabled.'
            return $false
        }

        $today = Get-Date -Format 'yyyy-MM-dd'
        $dailyLockFile = $null
        if ($UseDailyLock) {
            if ([string]::IsNullOrWhiteSpace($LockRoot)) { $LockRoot = 'C:\ProgramData\SmartM365\Locks' }
            if ([string]::IsNullOrWhiteSpace($LockName)) { $LockName = 'SmartM365-ActiveDirectory-Inventory-DailyReport' }
            if (-not (Test-Path -LiteralPath $LockRoot)) { New-Item -Path $LockRoot -ItemType Directory -Force | Out-Null }
            $dailyLockFile = Join-Path -Path $LockRoot -ChildPath ("{0}-SUCCESS-{1}.lock" -f $LockName, $today)
            if (Test-Path -LiteralPath $dailyLockFile) {
                WriteLog -Message ("[{0}] Daily report already succeeded on {1}. Skipping." -f $LockName, $today)
                return $false
            }
        }

        $reportSource = Get-InventoryDailyReportSource -SourceFolder $SourceFolder -AllowedOS $AllowedOS -TargetDomains $TargetDomains
        if ($null -eq $reportSource) { return $false }

        $nowText = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $osReportColumns = @('Windows 7', 'Windows 8', 'Windows 10', 'Windows 11')

        $computerRows = @($reportSource.Computers | Group-Object -Property ADDomain | ForEach-Object {
            $domainName = $_.Name
            $computersInGroup = $_.Group
            $enabledComputersInGroup = @($computersInGroup | Where-Object { $_.Enabled -eq $true })
            $disabledComputersInGroup = @($computersInGroup | Where-Object { $_.Enabled -eq $false })
            $countsByEnabledOS = @($enabledComputersInGroup | Group-Object -Property SimpleOS -NoElement)
            $countsByDisabledOS = @($disabledComputersInGroup | Group-Object -Property SimpleOS -NoElement)

            $outputRow = [ordered]@{
                Date             = $nowText
                DomainName       = $domainName
                TotalComputers   = $computersInGroup.Count
                EnabledAccounts  = $enabledComputersInGroup.Count
                DisabledAccounts = $disabledComputersInGroup.Count
            }

            foreach ($os in $osReportColumns) {
                $enabledOsGroup = @($countsByEnabledOS | Where-Object { $_.Name -eq $os } | Select-Object -First 1)
                $disabledOsGroup = @($countsByDisabledOS | Where-Object { $_.Name -eq $os } | Select-Object -First 1)
                $outputRow["$os Enabled"] = if ($enabledOsGroup.Count -gt 0) { $enabledOsGroup[0].Count } else { 0 }
                $outputRow["$os Disabled"] = if ($disabledOsGroup.Count -gt 0) { $disabledOsGroup[0].Count } else { 0 }
            }

            [PSCustomObject]$outputRow
        })

        $inactiveThreshold = (Get-Date).AddDays(-1 * $InactiveDays)
        $userRows = @($reportSource.Users | Where-Object { -not [string]::IsNullOrEmpty($_.ADDomain) } | Group-Object -Property ADDomain | ForEach-Object {
            $domainName = $_.Name
            $usersInGroup = $_.Group
            $enabledUsers = @($usersInGroup | Where-Object { $_.Enabled -eq $true })
            $disabledUsers = @($usersInGroup | Where-Object { $_.Enabled -eq $false })
            $inactiveUsers = @($usersInGroup | Where-Object { $null -eq $_.LastLogonDate -or $_.LastLogonDate -lt $inactiveThreshold })

            [PSCustomObject]([ordered]@{
                Date = $nowText
                DomainName = $domainName
                TotalUsers = $usersInGroup.Count
                EnabledUsers = $enabledUsers.Count
                DisabledUsers = $disabledUsers.Count
                Inactive90DaysUsers = $inactiveUsers.Count
            })
        })

        $computerReportPath = Join-Path -Path $ReportOutputPath -ChildPath 'AD_Computers_DailyStats.csv'
        $userReportPath = Join-Path -Path $ReportOutputPath -ChildPath 'AD_Users_DailyStats.csv'
        Write-DailyReportCsv -Rows $computerRows -OutputFilePath $computerReportPath -LatestFolderPath $LatestFolderPath
        Write-DailyReportCsv -Rows $userRows -OutputFilePath $userReportPath -LatestFolderPath $LatestFolderPath

        if ($UseDailyLock -and -not [string]::IsNullOrWhiteSpace($dailyLockFile)) {
            New-Item -Path $dailyLockFile -ItemType File -Force | Out-Null
            WriteLog -Message ("[{0}] Daily report success lock created for {1}." -f $LockName, $today)
        }

        WriteLog -Message 'Active Directory daily reports generated successfully.'
        return $true
    }
    function Get-ADStringValue {
        param([object]$Value)
        if ($null -eq $Value) { return $null }
        if ($Value -is [string]) { return $Value }
        if ($Value.Count -eq 0) { return $null }
        return [string]($Value[0])
    }


    function Get-CountryNameFromCode {
        param(
            [Parameter(Mandatory = $true)]
            [string]$CountryCode
        )

        $CountryLookup = @{
            "AE" = "United Arab Emirates"
            "AT" = "Austria"
            "BE" = "Belgium"
            "BR" = "Brazil"
            "CH" = "Switzerland"
            "CL" = "Chile"
            "CN" = "China"
            "CZ" = "Czech Republic"
            "DE" = "Germany"
            "ES" = "Spain"
            "FR" = "France"
            "GB" = "United Kingdom"
            "HR" = "Croatia"
            "IE" = "Ireland"
            "IT" = "Italy"
            "LU" = "Luxembourg"
            "LV" = "Latvia"
            "MX" = "Mexico"
            "NL" = "Netherlands"
            "PL" = "Poland"
            "PT" = "Portugal"
            "SI" = "Slovenia"
            "UY" = "Uruguay"
        }

        if ($CountryLookup.ContainsKey($CountryCode.ToUpper())) {
            return $CountryLookup[$CountryCode.ToUpper()]
        }
        else {
            return "Country code not found"
        }
    }

    function Test-UserAccountControlFlag {
        param(
            [Parameter(Mandatory = $true)]
            [int]$UserAccountControlValue,

            [Parameter(Mandatory = $true)]
            [int]$FlagToCheck
        )

        return (($UserAccountControlValue -band $FlagToCheck) -eq $FlagToCheck)
    }


    function Test-GroupMembershipByName {
        param(
            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string[]]$GroupNames,

            [Parameter(Mandatory = $true)]
            [string]$GroupNameToFind
        )

        if ($null -eq $GroupNames -or $GroupNames.Count -eq 0) {
            return $false
        }

        return ($GroupNames -contains $GroupNameToFind)
    }


    function Get-ComputerGroupNames {
        param(
            [Parameter(Mandatory = $true)]
            [object]$Computer,

            [Parameter(Mandatory = $true)]
            [string]$Server,

            [Parameter(Mandatory = $false)]
            [AllowNull()]
            [string]$DomainSid,

            [Parameter(Mandatory = $true)]
            [bool]$ResolveNestedGroups,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupNameByDNCache,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupParentsByDNCache,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupNameBySIDCache
        )

        if ($ResolveNestedGroups) {
            try {
                return @(Get-ADPrincipalGroupMembership -Identity $Computer -Server $Server | Select-Object -ExpandProperty Name)
            }
            catch {
                return @()
            }
        }

        $startDns = New-Object System.Collections.Generic.List[string]
        if ($Computer.MemberOf) {
            foreach ($memberDn in $Computer.MemberOf) {
                if (-not [string]::IsNullOrWhiteSpace([string]$memberDn)) {
                    [void]$startDns.Add([string]$memberDn)
                }
            }
        }

        if ($Computer.primaryGroupID -and $DomainSid) {
            try {
                $pgSid = ('{0}-{1}' -f $DomainSid, $Computer.primaryGroupID)
                $pgDn  = $null

                if ($GroupNameBySIDCache.ContainsKey($pgSid)) {
                    $pgDn = $GroupNameBySIDCache[$pgSid]
                }
                else {
                    $pgObj = Get-ADGroup -Identity $pgSid -Server $Server -Properties DistinguishedName
                    if ($pgObj -and $pgObj.DistinguishedName) {
                        $pgDn = [string]$pgObj.DistinguishedName
                        $GroupNameBySIDCache[$pgSid] = $pgDn
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($pgDn)) {
                    [void]$startDns.Add($pgDn)
                }
            }
            catch { }
        }

        if ($startDns.Count -eq 0) {
            return @()
        }

        $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $queue   = New-Object System.Collections.Queue
        $names   = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($dn in $startDns) {
            if ($visited.Add($dn)) {
                $queue.Enqueue($dn)
            }
        }

        while ($queue.Count -gt 0) {
            $gdn      = [string]$queue.Dequeue()
            $gName    = $null
            $gParents = $null

            if ($GroupNameByDNCache.ContainsKey($gdn)) {
                $gName = $GroupNameByDNCache[$gdn]
                if ($GroupParentsByDNCache.ContainsKey($gdn)) {
                    $gParents = $GroupParentsByDNCache[$gdn]
                }
            }
            else {
                try {
                    $gObj = Get-ADGroup -Identity $gdn -Server $Server -Properties Name, MemberOf
                    if ($gObj) {
                        $gName    = $gObj.Name
                        $gParents = @($gObj.MemberOf)
                        $GroupNameByDNCache[$gdn]    = $gName
                        $GroupParentsByDNCache[$gdn] = $gParents
                    }
                }
                catch {
                    $gParents = @()
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($gName)) {
                [void]$names.Add($gName)
            }

            if ($gParents) {
                foreach ($parentDn in $gParents) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$parentDn)) {
                        $parentDnString = [string]$parentDn
                        if ($visited.Add($parentDnString)) {
                            $queue.Enqueue($parentDnString)
                        }
                    }
                }
            }
        }

        return @($names)
    }


    # ----------------------------------------------------------
    # PATH VERIFICATION
    # ----------------------------------------------------------
    $destinationRootPath = $null
    try {
        $destinationRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''

        if ([string]::IsNullOrWhiteSpace($destinationRootPath)) {
            WriteLog -Message "WARNING: LatestCsvFolderPath not found in local configuration or returned empty. Combined CSV copy will be skipped."
            $destinationRootPath = $null
        }
        else {
            $destinationRootPath = $destinationRootPath.Trim()
            WriteLog -Message ("LatestCsvFolderPath resolved to: {0}" -f $destinationRootPath)
        }
    }
    catch {
        WriteLog -Message ("WARNING: Failed to resolve LatestCsvFolderPath: {0}. Combined CSV copy will be skipped." -f $_)
        $destinationRootPath = $null
    }

    if ($ReportOnly) {
        $reportSourceFolder = $destinationRootPath
        if ([string]::IsNullOrWhiteSpace($reportSourceFolder)) {
            $reportSourceFolder = $OutputPath
        }

        Invoke-ActiveDirectoryDailyReport -SourceFolder $reportSourceFolder `
            -ReportOutputPath $OutputPath `
            -LatestFolderPath $destinationRootPath `
            -AllowedOS $DailyReportAllowedOS `
            -TargetDomains $TargetDomains `
            -InactiveDays $DailyReportInactiveDays `
            -UseDailyLock $EnableDailyReportLock `
            -LockRoot $DailyReportLockRoot `
            -LockName $DailyReportLockName | Out-Null

        Remove-OldFiles -Path $OutputPath -Filter "*.csv" -OlderThanDays 30
        Remove-OldFiles -Path $OutputPath -Filter "*.log" -OlderThanDays 30
        WriteLog -Message ("{0} completed successfully in ReportOnly mode." -f $TaskName)
        return
    }

    $combinedUsersCsv     = Join-Path $OutputPath "AD_Users_AllDomains.csv"
    $combinedComputersCsv = Join-Path $OutputPath "AD_Computers_AllDomains.csv"
    $combinedGroupsCsv    = Join-Path $OutputPath "AD_Groups_AllDomains.csv"
    $combinedOusCsv       = Join-Path $OutputPath "AD_OUs_AllDomains.csv"
    $combinedContactsCsv  = Join-Path $OutputPath "AD_Contacts_AllDomains.csv"
    $duplicateUpnCsv      = Join-Path $OutputPath "AD_Users_DuplicateUPN.csv"
    $duplicateSmtpCsv     = Join-Path $OutputPath "AD_Users_DuplicateSMTP.csv"

    if (-not $DuplicateAnalysisOnly) {

    if ($TargetDomains -and $TargetDomains.Count -gt 0) {
        $DomainsToProcess = $TargetDomains
        WriteLog -Message ("Using explicitly provided target domains: {0}" -f ($DomainsToProcess -join ', '))
    }
    else {
        try {
            $forest = Get-ADForest -ErrorAction Stop
            $DomainsToProcess = $forest.Domains
            WriteLog -Message ("Discovered forest domains: {0}" -f ($DomainsToProcess -join ', '))
        }
        catch {
            throw "Unable to retrieve forest domains. $_"
        }
    }

    $skipDomainLoop = $false
    $DomainsToProcess = @($DomainsToProcess)

    if (-not $DomainWorker -and $EffectiveDomainParallelThrottleLimit -gt 1 -and $DomainsToProcess.Count -gt 1) {
        WriteLog -Message ("Starting parallel domain inventory with throttle limit {0}. Domains: {1}" -f $EffectiveDomainParallelThrottleLimit, ($DomainsToProcess -join ', '))

        $pwshPath = (Get-Process -Id $PID).Path
        $scriptPath = $PSCommandPath
        $workerJobs = @()
        $jobFailures = New-Object System.Collections.Generic.List[string]

        $receiveDomainJob = {
            param([System.Management.Automation.Job]$Job)

            $jobOutput = Receive-Job -Job $Job -ErrorAction SilentlyContinue 2>&1
            if ($jobOutput) {
                foreach ($line in $jobOutput) {
                    WriteLog -Message ("[{0}] {1}" -f $Job.Name, $line)
                }
            }

            if ($Job.State -ne 'Completed') {
                $reason = if ($Job.ChildJobs.Count -gt 0 -and $Job.ChildJobs[0].JobStateInfo.Reason) { $Job.ChildJobs[0].JobStateInfo.Reason.Message } else { $Job.State }
                [void]$jobFailures.Add(("{0}: {1}" -f $Job.Name, $reason))
            }

            Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
        }

        foreach ($domainName in $DomainsToProcess) {
            while (@($workerJobs | Where-Object { $_.State -eq 'Running' }).Count -ge $EffectiveDomainParallelThrottleLimit) {
                $completedJob = Wait-Job -Job $workerJobs -Any -Timeout 5
                if ($completedJob) {
                    & $receiveDomainJob $completedJob
                    $workerJobs = @($workerJobs | Where-Object { $_.Id -ne $completedJob.Id })
                }
            }

            $safeJobName = $domainName -replace '[^a-zA-Z0-9.-]', '_'
            $workerJobs += Start-Job -Name ("ADInventory-{0}" -f $safeJobName) -ScriptBlock {
                param($PwshPath, $ScriptPath, $TenantName, $DomainName, $OutputPathValue, $TempFolderPath)

                & $PwshPath -NoProfile -ExecutionPolicy Bypass -File $ScriptPath -Tenant $TenantName -TargetDomains $DomainName -OutputPath $OutputPathValue -DomainWorker -DomainWorkerTempFolder $TempFolderPath
                if ($LASTEXITCODE -ne 0) {
                    throw ("Domain worker failed for {0} with exit code {1}." -f $DomainName, $LASTEXITCODE)
                }
            } -ArgumentList $pwshPath, $scriptPath, $Tenant, $domainName, $OutputPath, $tempFolder
        }

        while ($workerJobs.Count -gt 0) {
            $completedJob = Wait-Job -Job $workerJobs -Any
            & $receiveDomainJob $completedJob
            $workerJobs = @($workerJobs | Where-Object { $_.Id -ne $completedJob.Id })
        }

        if ($jobFailures.Count -gt 0) {
            throw ("Parallel domain inventory failed: {0}" -f ($jobFailures -join ' | '))
        }

        WriteLog -Message "Parallel domain inventory workers completed successfully."
        $skipDomainLoop = $true
    }

    if (-not $skipDomainLoop) {
    foreach ($currentDomainName in $DomainsToProcess) {
        WriteLog -Message ("Starting inventory for domain '{0}'" -f $currentDomainName)

        $domainAttempt = 0
        $domainSuccess = $false

        while (-not $domainSuccess -and $domainAttempt -lt $MaxRetries) {
            $domainAttempt++
            if ($domainAttempt -gt 1) {
                WriteLog -Message ("Retrying inventory for domain '{0}' (attempt {1}/{2}) after {3}s delay..." -f $currentDomainName, $domainAttempt, $MaxRetries, $RetryDelaySeconds)
                Start-Sleep -Seconds $RetryDelaySeconds
            }

        try {

        $safeDomainFileName = $currentDomainName -replace '[^a-zA-Z0-9\.-]', '_'

        # ------------------------------------------------------
        # OU INVENTORY
        # ------------------------------------------------------
        if ($EnableOuInventory) {
        try {
            $CurrentObjectType = "OUs"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_OUs_{0}.csv" -f $safeDomainFileName)
            [int64]$ouCount = 0

            Get-ADOrganizationalUnit -Filter * -Server $currentDomainName -Properties Name, DistinguishedName, description, managedBy |
                ForEach-Object { [void]($ouCount++); $_ } |
                Select-Object `
                    @{Name = 'DomainName';   Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';   Expression = { $CurrentObjectType }},
                    Name,
                    DistinguishedName,
                    @{Name = 'Description'; Expression = { $_.Description -replace "`r", " -R " -replace "`n", " -N " }},
                    managedBy |
                Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported OUs for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $ouCount)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("OU inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("OU inventory skipped for domain '{0}' because EnableOuInventory is disabled." -f $currentDomainName)
        }

        # ------------------------------------------------------
        # COMPUTER INVENTORY
        # ------------------------------------------------------
        if ($EnableComputerInventory) {
        try {
            $CurrentObjectType = "Computers"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Computers_{0}.csv" -f $safeDomainFileName)

            $EnablePCFilter = $true
            if ($EnablePCFilter) {
                $computerFilter = {
                    (OperatingSystem -like "*Windows*" -and OperatingSystem -notlike "*Server*") -or (OperatingSystem -notlike "*")
                }
            }
            else {
                $computerFilter = { $true }
            }

            $GroupNameByDNCache     = @{}
            $GroupParentsByDNCache = @{}
            $GroupNameBySIDCache   = @{}

            try {
                $domainObj = Get-ADDomain -Server $currentDomainName
                $domainSid = $domainObj.DomainSID.Value
            }
            catch {
                $domainSid = $null
            }

            $ResolveNestedComputerGroups = $false
            [int64]$computerCount = 0

            Get-ADComputer -Filter $computerFilter -Server $currentDomainName -Properties SamAccountName, Name, DistinguishedName, Enabled, DNSHostName, OperatingSystem, operatingSystemHotfix, operatingSystemServicePack, operatingSystemVersion, LastLogonDate, LastLogonTimestamp, Description, IPv4Address, WhenCreated, WhenChanged, pwdLastSet, CanonicalName, MemberOf, primaryGroupID, ObjectGUID, ObjectSID, SIDHistory, extensionAttribute1, extensionAttribute2, extensionAttribute3, extensionAttribute4, extensionAttribute5, extensionAttribute6, extensionAttribute7, extensionAttribute8, extensionAttribute9, extensionAttribute10, extensionAttribute11, extensionAttribute12, extensionAttribute13, extensionAttribute14, extensionAttribute15 |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.DNSHostName) } |
                ForEach-Object {
                    [void]($computerCount++)
                    $computer = $_
                    $computerGroupNames = Get-ComputerGroupNames -Computer $computer -Server $currentDomainName -DomainSid $domainSid -ResolveNestedGroups:$ResolveNestedComputerGroups -GroupNameByDNCache $GroupNameByDNCache -GroupParentsByDNCache $GroupParentsByDNCache -GroupNameBySIDCache $GroupNameBySIDCache
                    $hasGroupAddIntune  = [bool]($computerGroupNames | Where-Object { $_ -like $IntuneEnrollmentGroupPattern })
                    $hasGroupUpgradeW11 = [bool]($computerGroupNames | Where-Object { $_ -like $Windows11UpgradeGroupPattern })

                    [PSCustomObject][ordered]@{
                        DomainName              = $currentDomainName
                        ObjectType              = $CurrentObjectType
                        SamAccountName          = $computer.SamAccountName
                        Name                    = $computer.Name
                        DistinguishedName       = $computer.DistinguishedName
                        Enabled                 = $computer.Enabled
                        DNSHostName             = $computer.DNSHostName
                        OperatingSystem         = $computer.OperatingSystem
                        operatingSystemHotfix   = $computer.operatingSystemHotfix
                        operatingSystemServicePack = $computer.operatingSystemServicePack
                        operatingSystemVersion  = $computer.operatingSystemVersion
                        LastLogonDate           = if ($computer.LastLogonTimestamp -ne $null -and $computer.LastLogonTimestamp -ne 0) { [datetime]::FromFileTime($computer.LastLogonTimestamp) } else { '' }
                        Description             = $computer.Description -replace "`r", " -R " -replace "`n", " -N "
                        IPv4Address             = $computer.IPv4Address
                        WhenCreated             = $computer.WhenCreated
                        WhenChanged             = $computer.WhenChanged
                        pwdLastSetDate          = if ($computer.pwdLastSet -ne $null -and $computer.pwdLastSet -ne 0) { [datetime]::FromFileTime($computer.pwdLastSet) } else { '' }
                        CanonicalName           = $computer.CanonicalName -replace "`r", " -R " -replace "`n", " -N "
                        MemberOfDNs             = if ($computer.MemberOf) { ($computer.MemberOf) -join ';' } else { '' }
                        PrimaryGroupName        = if ($computer.primaryGroupID -and $domainSid) {
                            $pgSid = ('{0}-{1}' -f $domainSid, $computer.primaryGroupID)
                            if ($GroupNameBySIDCache.ContainsKey($pgSid)) {
                                $pgDn = $GroupNameBySIDCache[$pgSid]
                                if (-not [string]::IsNullOrWhiteSpace($pgDn)) {
                                    if ($GroupNameByDNCache.ContainsKey($pgDn)) {
                                        $GroupNameByDNCache[$pgDn]
                                    }
                                    else {
                                        try {
                                            $g = Get-ADGroup -Identity $pgSid -Server $currentDomainName
                                            $GroupNameByDNCache[$pgDn] = $g.Name
                                            $g.Name
                                        }
                                        catch { '' }
                                    }
                                }
                                else {
                                    ''
                                }
                            }
                            else {
                                try {
                                    $g = Get-ADGroup -Identity $pgSid -Server $currentDomainName -Properties DistinguishedName
                                    if ($g -and $g.DistinguishedName) {
                                        $GroupNameBySIDCache[$pgSid] = [string]$g.DistinguishedName
                                        $GroupNameByDNCache[[string]$g.DistinguishedName] = $g.Name
                                    }
                                    $g.Name
                                }
                                catch { '' }
                            }
                        } else { '' }
                        ObjectGUID              = $computer.ObjectGUID
                        SID                     = $computer.ObjectSID.Value
                        SIDHistory              = if ($computer.SIDHistory) { ($computer.SIDHistory | ForEach-Object { $_.Value }) -join ';' } else { '' }
                        extensionAttribute1     = Get-ADStringValue $computer.extensionAttribute1
                        extensionAttribute2     = Get-ADStringValue $computer.extensionAttribute2
                        extensionAttribute3     = Get-ADStringValue $computer.extensionAttribute3
                        extensionAttribute4     = Get-ADStringValue $computer.extensionAttribute4
                        extensionAttribute5     = Get-ADStringValue $computer.extensionAttribute5
                        extensionAttribute6     = Get-ADStringValue $computer.extensionAttribute6
                        extensionAttribute7     = Get-ADStringValue $computer.extensionAttribute7
                        extensionAttribute8     = Get-ADStringValue $computer.extensionAttribute8
                        extensionAttribute9     = Get-ADStringValue $computer.extensionAttribute9
                        extensionAttribute10    = Get-ADStringValue $computer.extensionAttribute10
                        extensionAttribute11    = Get-ADStringValue $computer.extensionAttribute11
                        extensionAttribute12    = Get-ADStringValue $computer.extensionAttribute12
                        extensionAttribute13    = Get-ADStringValue $computer.extensionAttribute13
                        extensionAttribute14    = Get-ADStringValue $computer.extensionAttribute14
                        extensionAttribute15    = Get-ADStringValue $computer.extensionAttribute15
                        DomainNameShort         = Get-DomainNameShort -DomainName $currentDomainName
                        DomainAndSam            = Get-NormalizedDomainAndSam -DomainNameShort (Get-DomainNameShort -DomainName $currentDomainName) -SamAccountName $computer.SamAccountName
                        ImmutableId_AD          = Convert-GuidToImmutableId -ObjectGuid ([string]$computer.ObjectGUID)
                        Has_Group_AddIntune     = $hasGroupAddIntune
                        Has_Group_UpgradeW11    = $hasGroupUpgradeW11
                    }
                } |
                Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Computers for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $computerCount)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Computer inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("Computer inventory skipped for domain '{0}' because EnableComputerInventory is disabled." -f $currentDomainName)
        }

        # ------------------------------------------------------
        # USER INVENTORY
        # ------------------------------------------------------
        if ($EnableUserInventory) {
        try {
            $CurrentObjectType = "Users"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Users_{0}.csv" -f $safeDomainFileName)
            [int64]$userCount = 0

            [int]$DomainExcludedUsersNoUpn = 0
            try {
                $DomainExcludedUsersNoUpn = (Get-ADUser -LDAPFilter "(&(objectCategory=person)(objectClass=user)(!(userPrincipalName=*)))" -Server $currentDomainName -ResultSetSize $null).Count
            }
            catch {
                WriteLog -Message ("WARNING: Failed to count users without UPN for domain '{0}': {1}" -f $currentDomainName, $_)
                $DomainExcludedUsersNoUpn = 0
            }
            WriteLog -Message ("Users excluded because of missing UPN for domain '{0}': {1}" -f $currentDomainName, $DomainExcludedUsersNoUpn)

            Get-ADUser -LDAPFilter "(&(objectCategory=person)(objectClass=user)(userPrincipalName=*))" -Server $currentDomainName -ResultSetSize $null -Properties SamAccountName, sAMAccountType, Name, DistinguishedName, UserPrincipalName, Enabled, manager, LastLogonTimestamp, DisplayName, GivenName, Surname, Description, Department, Title, Company, Office, TelephoneNumber, MobilePhone, EmailAddress, StreetAddress, City, PostalCode, Country, WhenCreated, WhenChanged, AccountExpirationDate, pwdLastSet, badPwdCount, badPasswordTime, LogonCount, userAccountControl, msDS-ManagedPassword, ProxyAddresses, MemberOf, CanonicalName, ObjectGUID, targetAddress, ObjectSID, SIDHistory, extensionAttribute1, extensionAttribute2, extensionAttribute3, extensionAttribute4, extensionAttribute5, extensionAttribute6, extensionAttribute7, extensionAttribute8, extensionAttribute9, extensionAttribute10, extensionAttribute11, extensionAttribute12, extensionAttribute13, extensionAttribute14, extensionAttribute15 |
                Select-Object `
                    @{Name = 'DomainName';           Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';           Expression = { $CurrentObjectType }},
                    SamAccountName,
                    sAMAccountType,
                    @{Name = 'Name';                Expression = { $_.Name        -replace "`r", " -R " -replace "`n", " -N " }},
                    DistinguishedName,
                    UserPrincipalName,
                    Enabled,
                    manager,
                    LastLogonTimestamp,
                    @{Name = 'LastLogonDate';       Expression = {
                        if ($_.LastLogonTimestamp -eq 0 -or $_.LastLogonTimestamp -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.LastLogonTimestamp)
                        }
                    }},
                    @{Name = 'DisplayName';          Expression = { $_.DisplayName  -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'GivenName';            Expression = { $_.GivenName    -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Surname';             Expression = { $_.Surname     -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Description';         Expression = { $_.Description -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Department';          Expression = { $_.Department  -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Title';               Expression = { $_.Title       -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Company';             Expression = { $_.Company     -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'Office';              Expression = { $_.Office      -replace "`r", " -R " -replace "`n", " -N " }},
                    TelephoneNumber,
                    MobilePhone,
                    EmailAddress,
                    @{Name = 'StreetAddress';       Expression = { $_.StreetAddress -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'City';                Expression = { $_.City          -replace "`r", " -R " -replace "`n", " -N " }},
                    PostalCode,
                    Country,
                    @{Name = 'CountryName';         Expression = {
                        if (-not [string]::IsNullOrWhiteSpace($_.Country)) {
                            Get-CountryNameFromCode -CountryCode $_.Country
                        }
                        else {
                            "Unknown"
                        }
                    }},
                    WhenCreated,
                    WhenChanged,
                    AccountExpirationDate,
                    @{Name = 'PasswordLastSetDate'; Expression = {
                        if ($_.pwdLastSet -eq 0 -or $_.pwdLastSet -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.pwdLastSet)
                        }
                    }},
                    badPwdCount,
                    @{Name = 'BadPasswordDate';     Expression = {
                        if ($_.badPasswordTime -eq 0 -or $_.badPasswordTime -eq $null) {
                            ""
                        }
                        else {
                            [datetime]::FromFileTime($_.badPasswordTime)
                        }
                    }},
                    LogonCount,
                    userAccountControl,
                    @{Name = 'IsNormalAccount';         Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 64 }},
                    @{Name = 'IsScriptAccount';         Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 1 }},
                    @{Name = 'IsPasswordNeverExpires';  Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 65536 }},
                    @{Name = 'IsTrustedForDelegation';  Expression = { Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 524288 }},
                    @{Name = 'MemberOfGroups';          Expression = { $_.MemberOf      -join ";" }},
                    @{Name = 'ProxyAddresses';          Expression = { $_.ProxyAddresses -join ";" }},
                    @{Name = 'CanonicalName';           Expression = { $_.CanonicalName -replace "`r", " -R " -replace "`n", " -N " }},
                    @{Name = 'ObjectGUID';              Expression = { $_.ObjectGUID }},
                    @{Name = 'TargetAddress';           Expression = { $_.targetAddress }},
                    @{Name = 'ObjectSID';               Expression = { $_.ObjectSID.Value }},
                    @{Name = 'ObjectSIDHistory';        Expression = {
                        if ($_.SIDHistory) {
                            ($_.SIDHistory | ForEach-Object { $_.Value }) -join ";"
                        }
                        else {
                            ""
                        }
                    }},
                    @{Name = 'extensionAttribute1';          Expression = { Get-ADStringValue $_.extensionAttribute1 }},
                    @{Name = 'extensionAttribute2';          Expression = { Get-ADStringValue $_.extensionAttribute2 }},
                    @{Name = 'extensionAttribute3';          Expression = { Get-ADStringValue $_.extensionAttribute3 }},
                    @{Name = 'extensionAttribute4';          Expression = { Get-ADStringValue $_.extensionAttribute4 }},
                    @{Name = 'extensionAttribute5';          Expression = { Get-ADStringValue $_.extensionAttribute5 }},
                    @{Name = 'extensionAttribute6';          Expression = { Get-ADStringValue $_.extensionAttribute6 }},
                    @{Name = 'extensionAttribute7';          Expression = { Get-ADStringValue $_.extensionAttribute7 }},
                    @{Name = 'extensionAttribute8';          Expression = { Get-ADStringValue $_.extensionAttribute8 }},
                    @{Name = 'extensionAttribute9';          Expression = { Get-ADStringValue $_.extensionAttribute9 }},
                    @{Name = 'extensionAttribute10';         Expression = { Get-ADStringValue $_.extensionAttribute10 }},
                    @{Name = 'extensionAttribute11';         Expression = { Get-ADStringValue $_.extensionAttribute11 }},
                    @{Name = 'extensionAttribute12';         Expression = { Get-ADStringValue $_.extensionAttribute12 }},
                    @{Name = 'extensionAttribute13';         Expression = { Get-ADStringValue $_.extensionAttribute13 }},
                    @{Name = 'extensionAttribute14';         Expression = { Get-ADStringValue $_.extensionAttribute14 }},
                    @{Name = 'extensionAttribute15';         Expression = { Get-ADStringValue $_.extensionAttribute15 }},
                    @{Name = 'MustChangePasswordAtNextLogon'; Expression = { ($_.pwdLastSet -eq 0) -and (-not (Test-UserAccountControlFlag -UserAccountControlValue $_.userAccountControl -FlagToCheck 65536)) }},
                    @{Name = 'DomainNameShort';         Expression = { Get-DomainNameShort -DomainName $currentDomainName }},
                    @{Name = 'DomainAndSam';            Expression = { Get-NormalizedDomainAndSam -DomainNameShort (Get-DomainNameShort -DomainName $currentDomainName) -SamAccountName $_.SamAccountName }},
                    @{Name = 'ImmutableId_AD';          Expression = { Convert-GuidToImmutableId -ObjectGuid ([string]$_.ObjectGUID) }} |
                ForEach-Object { [void]($userCount++); $_ } |
                Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Users for domain '{0}' to '{1}'. Count: {2}. Excluded without UPN: {3}" -f $currentDomainName, $outputCsvFilePath, $userCount, $DomainExcludedUsersNoUpn)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("User inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("User inventory skipped for domain '{0}' because EnableUserInventory is disabled." -f $currentDomainName)
        }

        # ------------------------------------------------------
        # GROUP INVENTORY
        # ------------------------------------------------------
        if ($EnableGroupInventory) {
        try {
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Groups_{0}.csv" -f $safeDomainFileName)
            $GroupData = @(Get-ADGroup -Filter * -Server $currentDomainName -Properties CanonicalName, CN, Created, createTimeStamp, Deleted, Description, DisplayName, DistinguishedName, GroupCategory, GroupScope, GroupType, HomePage, LastKnownParent, mail, ManagedBy, MemberOf, Members, Modified, modifyTimeStamp, Name, ObjectCategory, ObjectClass, ObjectGUID, objectSid, ProtectedFromAccidentalDeletion, SamAccountName, SIDHistory, whenChanged, whenCreated |
                Select-Object `
                    @{Name = 'DomainName'; Expression = { [string]$currentDomainName }},
                    @{Name = 'CanonicalName'; Expression = { Get-ADStringValue $_.CanonicalName }},
                    @{Name = 'CN'; Expression = { Get-ADStringValue $_.CN }},
                    @{Name = 'Created'; Expression = { $_.Created }},
                    @{Name = 'createTimeStamp'; Expression = { $_.createTimeStamp }},
                    @{Name = 'Deleted'; Expression = { $_.Deleted }},
                    @{Name = 'Description'; Expression = { Get-ADStringValue $_.Description }},
                    @{Name = 'DisplayName'; Expression = { Get-ADStringValue $_.DisplayName }},
                    @{Name = 'DistinguishedName'; Expression = { Get-ADStringValue $_.DistinguishedName }},
                    @{Name = 'GroupCategory'; Expression = { Get-ADStringValue $_.GroupCategory }},
                    @{Name = 'GroupScope'; Expression = { Get-ADStringValue $_.GroupScope }},
                    @{Name = 'GroupType'; Expression = { $_.GroupType }},
                    @{Name = 'HomePage'; Expression = { Get-ADStringValue $_.HomePage }},
                    @{Name = 'LastKnownParent'; Expression = { Get-ADStringValue $_.LastKnownParent }},
                    @{Name = 'mail'; Expression = { Get-ADStringValue $_.mail }},
                    @{Name = 'ManagedBy'; Expression = { Get-ADStringValue $_.ManagedBy }},
                    @{Name = 'MemberOf'; Expression = { if ($_.MemberOf) { ($_.MemberOf -join ';') } else { '' } }},
                    @{Name = 'Members'; Expression = { if ($_.Members) { ($_.Members -join ';') } else { '' } }},
                    @{Name = 'Modified'; Expression = { $_.Modified }},
                    @{Name = 'modifyTimeStamp'; Expression = { $_.modifyTimeStamp }},
                    @{Name = 'Name'; Expression = { Get-ADStringValue $_.Name }},
                    @{Name = 'ObjectCategory'; Expression = { Get-ADStringValue $_.ObjectCategory }},
                    @{Name = 'ObjectClass'; Expression = { Get-ADStringValue $_.ObjectClass }},
                    @{Name = 'ObjectGUID'; Expression = { if ($_.ObjectGUID) { $_.ObjectGUID.Guid } else { $null } }},
                    @{Name = 'objectSid'; Expression = { $_.objectSid.Value }},
                    @{Name = 'ProtectedFromAccidentalDeletion'; Expression = { $_.ProtectedFromAccidentalDeletion }},
                    @{Name = 'SamAccountName'; Expression = { Get-ADStringValue $_.SamAccountName }},
                    @{Name = 'SIDHistory'; Expression = { if ($_.SIDHistory) { ($_.SIDHistory.Value) -join ';' } else { '' } }},
                    @{Name = 'whenChanged'; Expression = { $_.whenChanged }},
                    @{Name = 'whenCreated'; Expression = { $_.whenCreated }}
            )
            $GroupData | Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8
            WriteLog -Message ("Exported Groups for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $GroupData.Count)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Group inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("Group inventory skipped for domain '{0}' because EnableGroupInventory is disabled." -f $currentDomainName)
        }

        # ------------------------------------------------------
        # CONTACT INVENTORY
        # ------------------------------------------------------
        if ($EnableContactInventory) {
        try {
            $CurrentObjectType = "Contacts"
            $outputCsvFilePath = Join-Path $tempFolder ("AD_Contacts_{0}.csv" -f $safeDomainFileName)
            [int64]$contactCount = 0

            Get-ADObject -Filter { ObjectClass -eq "contact" } -Server $currentDomainName -Properties DisplayName, ProxyAddresses, Mail |
                ForEach-Object { [void]($contactCount++); $_ } |
                Select-Object `
                    @{Name = 'DomainName';      Expression = { $currentDomainName }},
                    @{Name = 'ObjectType';      Expression = { $CurrentObjectType }},
                    DisplayName,
                    @{Name = 'ProxyAddresses'; Expression = { $_.ProxyAddresses -join ";" }},
                    Mail |
                Export-Csv $outputCsvFilePath -NoTypeInformation -Encoding UTF8

            WriteLog -Message ("Exported Contacts for domain '{0}' to '{1}'. Count: {2}" -f $currentDomainName, $outputCsvFilePath, $contactCount)
        }
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) { throw }
            WriteLog -Message ("Contact inventory failed for domain '{0}': {1}" -f $currentDomainName, $_)
        }
        }
        else {
            WriteLog -Message ("Contact inventory skipped for domain '{0}' because EnableContactInventory is disabled." -f $currentDomainName)
        }

        $domainSuccess = $true

        } # end outer domain try
        catch {
            if (Test-IsTransientADError -ErrorRecord $_) {
                if ($domainAttempt -lt $MaxRetries) {
                    WriteLog -Message ("WARNING: Transient AD connectivity error for domain '{0}' (attempt {1}/{2}): {3}" -f $currentDomainName, $domainAttempt, $MaxRetries, $_.Exception.Message)
                    # Loop continues: next iteration will wait and retry
                }
                else {
                    WriteLog -Message ("ERROR: Transient AD connectivity error for domain '{0}' persisted after {1} attempt(s). Skipping domain. Error: {2}" -f $currentDomainName, $domainAttempt, $_.Exception.Message)
                }
            }
            else {
                WriteLog -Message ("FATAL: Unhandled error for domain '{0}' (attempt {1}/{2}): {3}. Continuing with next domain." -f $currentDomainName, $domainAttempt, $MaxRetries, $_)
                break
            }
        }

        } # end while retry loop

        if (-not $domainSuccess) {
            WriteLog -Message ("WARNING: Domain '{0}' could not be fully inventoried after {1} attempt(s). Skipping." -f $currentDomainName, $domainAttempt)
        }
    }
    }

    if ($DomainWorker) {
        WriteLog -Message "Domain worker completed; combined CSV generation is handled by the parent process."
        return
    }

    # ------------------------------------------------------
    # COMBINE PER-DOMAIN CSV FILES
    # ------------------------------------------------------
    $combinedUsersCsv     = Join-Path $OutputPath "AD_Users_AllDomains.csv"
    $combinedComputersCsv = Join-Path $OutputPath "AD_Computers_AllDomains.csv"
    $combinedGroupsCsv    = Join-Path $OutputPath "AD_Groups_AllDomains.csv"
    $combinedOusCsv       = Join-Path $OutputPath "AD_OUs_AllDomains.csv"
    $combinedContactsCsv  = Join-Path $OutputPath "AD_Contacts_AllDomains.csv"

    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Users_*.csv"     -DestinationFile $combinedUsersCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Computers_*.csv" -DestinationFile $combinedComputersCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Groups_*.csv"    -DestinationFile $combinedGroupsCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_OUs_*.csv"       -DestinationFile $combinedOusCsv
    Combine-CsvFiles -SourceFolder $tempFolder -Filter "AD_Contacts_*.csv"  -DestinationFile $combinedContactsCsv

    # ------------------------------------------------------
    # Copy combined CSVs to the latest CSV folder
    # ------------------------------------------------------
    try {
        if ($destinationRootPath) {
            if (-not (Test-Path -LiteralPath $destinationRootPath)) {
                New-Item -Path $destinationRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                WriteLog -Message ("Created missing LatestCsvFolderPath directory: {0}" -f $destinationRootPath)
            }
            foreach ($combinedCsv in @($combinedUsersCsv, $combinedComputersCsv, $combinedGroupsCsv, $combinedOusCsv, $combinedContactsCsv)) {
                if (Test-Path -Path $combinedCsv) {
                    $destinationFile = Join-Path $destinationRootPath ([System.IO.Path]::GetFileName($combinedCsv))
                    Copy-Item -LiteralPath $combinedCsv -Destination $destinationFile -Force -ErrorAction Stop
                    WriteLog -Message ("Copied combined CSV '{0}' to '{1}'" -f $combinedCsv, $destinationFile)
                    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $destinationFile
                }
                else {
                    WriteLog -Message ("Combined CSV not found, skipping copy: {0}" -f $combinedCsv)
                }
            }
        }
        else {
            WriteLog -Message "LatestCsvFolderPath unavailable. Combined CSV copy skipped."
        }
    }
    catch {
        WriteLog -Message ("Copy to LatestCsvFolderPath failed: {0}" -f $_)
    }

    }
    else {
        WriteLog -Message "DuplicateAnalysisOnly mode: using existing combined users CSV for duplicate analysis."
        if ($destinationRootPath) {
            $latestUsersCsv = Join-Path $destinationRootPath "AD_Users_AllDomains.csv"
            if (Test-Path -LiteralPath $latestUsersCsv) {
                $combinedUsersCsv = $latestUsersCsv
                WriteLog -Message ("DuplicateAnalysisOnly source CSV: {0}" -f $combinedUsersCsv)
            }
            else {
                WriteLog -Message ("Latest AD users CSV not found, falling back to OutputPath: {0}" -f $latestUsersCsv) -Level "WARNING"
            }

            foreach ($latestName in @('AD_Computers_AllDomains.csv', 'AD_Groups_AllDomains.csv', 'AD_OUs_AllDomains.csv', 'AD_Contacts_AllDomains.csv')) {
                $candidatePath = Join-Path $destinationRootPath $latestName
                if (-not (Test-Path -LiteralPath $candidatePath)) { continue }
                switch ($latestName) {
                    'AD_Computers_AllDomains.csv' { $combinedComputersCsv = $candidatePath }
                    'AD_Groups_AllDomains.csv'    { $combinedGroupsCsv = $candidatePath }
                    'AD_OUs_AllDomains.csv'       { $combinedOusCsv = $candidatePath }
                    'AD_Contacts_AllDomains.csv'  { $combinedContactsCsv = $candidatePath }
                }
            }
        }
    }

    # ------------------------------------------------------
    # DUPLICATE USER IDENTITY ANALYSIS
    # ------------------------------------------------------
    if ($EnableDuplicateAnalysis) {
        try {
            WriteLog -Message "Starting duplicate UPN and SMTP proxy address analysis..."

            if (-not (Test-Path -Path $combinedUsersCsv)) {
                WriteLog -Message ("WARNING: Combined users CSV not found, skipping duplicate analysis: {0}" -f $combinedUsersCsv)
            }
            else {
                $allUsers = @(Import-Csv -Path $combinedUsersCsv -Encoding UTF8)
                WriteLog -Message ("Loaded {0} users from combined CSV for duplicate analysis." -f $allUsers.Count)

                $upnMap = @{}
                $smtpMap = @{}

                foreach ($u in $allUsers) {
                    $upnKey = ([string]$u.UserPrincipalName).Trim().ToLowerInvariant()
                    if (-not [string]::IsNullOrWhiteSpace($upnKey)) {
                        if (-not $upnMap.ContainsKey($upnKey)) {
                            $upnMap[$upnKey] = [System.Collections.Generic.List[object]]::new()
                        }
                        [void]$upnMap[$upnKey].Add($u)
                    }

                    if ([string]::IsNullOrWhiteSpace($u.ProxyAddresses)) { continue }
                    foreach ($entry in ([string]$u.ProxyAddresses -split ';')) {
                        $entry = $entry.Trim()
                        if ($entry -notmatch '^smtp:(.+)$') { continue }

                        $smtpAddress = $Matches[1].Trim()
                        if ([string]::IsNullOrWhiteSpace($smtpAddress)) { continue }

                        $smtpKey = $smtpAddress.ToLowerInvariant()
                        if (-not $smtpMap.ContainsKey($smtpKey)) {
                            $smtpMap[$smtpKey] = [System.Collections.Generic.List[object]]::new()
                        }

                        [void]$smtpMap[$smtpKey].Add([PSCustomObject]@{
                            SmtpAddress       = $smtpAddress
                            IsUppercaseSMTP   = $entry.StartsWith('SMTP:', [System.StringComparison]::Ordinal)
                            UserPrincipalName = $u.UserPrincipalName
                            DomainName        = $u.DomainName
                            DomainNameShort   = $u.DomainNameShort
                            SamAccountName    = $u.SamAccountName
                            DisplayName       = $u.DisplayName
                            Enabled           = $u.Enabled
                            DistinguishedName = $u.DistinguishedName
                        })
                    }
                }

                $duplicateUpnRows = @(
                    foreach ($upnKey in @($upnMap.Keys | Sort-Object)) {
                        $users = $upnMap[$upnKey]
                        if ($users.Count -le 1) { continue }

                        foreach ($u in @($users | Sort-Object DomainName, SamAccountName)) {
                            [PSCustomObject][ordered]@{
                                UserPrincipalName   = $u.UserPrincipalName
                                UPN_OccurrenceCount = $users.Count
                                DomainName          = $u.DomainName
                                DomainNameShort     = $u.DomainNameShort
                                SamAccountName      = $u.SamAccountName
                                DisplayName         = $u.DisplayName
                                Enabled             = $u.Enabled
                                DistinguishedName   = $u.DistinguishedName
                            }
                        }
                    }
                )

                $duplicateUpnRows | Export-Csv -Path $duplicateUpnCsv -NoTypeInformation -Encoding UTF8
                $upnDuplicateCount = @($upnMap.Keys | Where-Object { $upnMap[$_].Count -gt 1 }).Count
                WriteLog -Message ("Duplicate UPN analysis complete. Distinct duplicate UPNs: {0}. Affected accounts: {1}. Output: {2}" -f $upnDuplicateCount, $duplicateUpnRows.Count, $duplicateUpnCsv)

                $duplicateSmtpRows = @(
                    foreach ($smtpKey in @($smtpMap.Keys | Sort-Object)) {
                        $entries = $smtpMap[$smtpKey]
                        if ($entries.Count -le 1) { continue }

                        foreach ($e in @($entries | Sort-Object DomainName, UserPrincipalName, SmtpAddress)) {
                            [PSCustomObject][ordered]@{
                                SmtpAddress          = $e.SmtpAddress
                                SMTP_OccurrenceCount = $entries.Count
                                IsUppercaseSMTP      = $e.IsUppercaseSMTP
                                UserPrincipalName    = $e.UserPrincipalName
                                DomainName           = $e.DomainName
                                DomainNameShort      = $e.DomainNameShort
                                SamAccountName       = $e.SamAccountName
                                DisplayName          = $e.DisplayName
                                Enabled              = $e.Enabled
                                DistinguishedName    = $e.DistinguishedName
                            }
                        }
                    }
                )

                $duplicateSmtpRows | Export-Csv -Path $duplicateSmtpCsv -NoTypeInformation -Encoding UTF8
                $smtpDuplicateCount = @($smtpMap.Keys | Where-Object { $smtpMap[$_].Count -gt 1 }).Count
                WriteLog -Message ("Duplicate SMTP analysis complete. Distinct duplicate addresses: {0}. Affected entries: {1}. Output: {2}" -f $smtpDuplicateCount, $duplicateSmtpRows.Count, $duplicateSmtpCsv)

                foreach ($duplicateCsv in @($duplicateUpnCsv, $duplicateSmtpCsv)) {
                    if ($destinationRootPath -and (Test-Path -Path $duplicateCsv)) {
                        $destinationFile = Join-Path $destinationRootPath ([System.IO.Path]::GetFileName($duplicateCsv))
                        Copy-Item -LiteralPath $duplicateCsv -Destination $destinationFile -Force -ErrorAction Stop
                        WriteLog -Message ("Copied '{0}' to '{1}'" -f $duplicateCsv, $destinationFile)
                        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $destinationFile
                    }
                }

                $hasDuplicateIdentities = (($upnDuplicateCount -gt 0) -or ($smtpDuplicateCount -gt 0))
                if ($EnableDuplicateNotification -and $hasDuplicateIdentities) {
                    if ([string]::IsNullOrWhiteSpace($DuplicateNotificationLastSentFilePath)) {
                        $DuplicateNotificationLastSentFilePath = Join-Path $OutputPath 'AD_DuplicateNotification_LastSent.txt'
                    }

                    $todayStamp = (Get-Date).ToString('yyyy-MM-dd')
                    $lastSentStamp = ''
                    if (Test-Path -LiteralPath $DuplicateNotificationLastSentFilePath) {
                        $lastSentStamp = (Get-Content -LiteralPath $DuplicateNotificationLastSentFilePath -Raw -ErrorAction SilentlyContinue).Trim()
                    }

                    if ($lastSentStamp -eq $todayStamp -and -not $ForceSendDuplicateNotification) {
                        WriteLog -Message ("Duplicate identity notification already sent today ({0}). Use -ForceSendDuplicateNotification to resend." -f $todayStamp)
                    }
                    else {
                        $notificationFolder = Split-Path -Path $DuplicateNotificationLastSentFilePath -Parent
                        if (-not [string]::IsNullOrWhiteSpace($notificationFolder) -and -not (Test-Path -LiteralPath $notificationFolder)) {
                            New-Item -Path $notificationFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }

                        $emailSubject = "SmartM365 Active Directory duplicate identities detected"
                        $notificationGeneratedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
                        $htmlSourceUsersCsv = [System.Net.WebUtility]::HtmlEncode([string]$combinedUsersCsv)
                        $htmlDuplicateUpnCsv = [System.Net.WebUtility]::HtmlEncode([string]$duplicateUpnCsv)
                        $htmlDuplicateSmtpCsv = [System.Net.WebUtility]::HtmlEncode([string]$duplicateSmtpCsv)
                        $htmlTenant = [System.Net.WebUtility]::HtmlEncode([string]$Tenant)
                        $htmlHost = [System.Net.WebUtility]::HtmlEncode([string]$env:COMPUTERNAME)
                        $htmlGeneratedAt = [System.Net.WebUtility]::HtmlEncode([string]$notificationGeneratedAt)
                        $emailBody = @"
<!doctype html>
<html>
<body style="margin:0;padding:0;background:#f4f6f8;font-family:Segoe UI,Arial,sans-serif;color:#111827;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;padding:24px 0;">
    <tr>
      <td align="center">
        <table role="presentation" width="760" cellpadding="0" cellspacing="0" style="width:760px;max-width:760px;background:#ffffff;border:1px solid #d9e2ec;border-radius:6px;overflow:hidden;">
          <tr>
            <td style="background:#0f172a;color:#ffffff;padding:20px 24px;">
              <div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#93c5fd;font-weight:700;">SmartM365 Active Directory</div>
              <div style="font-size:24px;line-height:30px;font-weight:700;margin-top:6px;">Duplicate identities detected</div>
              <div style="font-size:13px;color:#cbd5e1;margin-top:8px;">Tenant: $htmlTenant &nbsp;|&nbsp; Host: $htmlHost &nbsp;|&nbsp; Generated: $htmlGeneratedAt</div>
            </td>
          </tr>
          <tr>
            <td style="padding:22px 24px 8px 24px;">
              <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">
                <tr>
                  <td style="background:#fff7ed;border:1px solid #fed7aa;border-left:5px solid #f97316;padding:12px 14px;border-radius:4px;">
                    <div style="font-size:14px;font-weight:700;color:#9a3412;">Action required</div>
                    <div style="font-size:13px;line-height:19px;color:#7c2d12;margin-top:4px;">Review the duplicate UPN and SMTP CSV files before identity cleanup, migration, or synchronization decisions.</div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:14px 24px 4px 24px;">
              <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
                <tr>
                  <th align="left" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Check</th>
                  <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Distinct duplicates</th>
                  <th align="right" style="background:#f8fafc;border-bottom:1px solid #d9e2ec;padding:10px 12px;font-size:12px;color:#475569;text-transform:uppercase;">Affected entries</th>
                </tr>
                <tr>
                  <td style="border-bottom:1px solid #eef2f7;padding:12px;font-size:14px;font-weight:600;">UserPrincipalName</td>
                  <td align="right" style="border-bottom:1px solid #eef2f7;padding:12px;font-size:18px;font-weight:700;color:#b45309;">$upnDuplicateCount</td>
                  <td align="right" style="border-bottom:1px solid #eef2f7;padding:12px;font-size:18px;font-weight:700;color:#b45309;">$($duplicateUpnRows.Count)</td>
                </tr>
                <tr>
                  <td style="padding:12px;font-size:14px;font-weight:600;">SMTP proxy address</td>
                  <td align="right" style="padding:12px;font-size:18px;font-weight:700;color:#b45309;">$smtpDuplicateCount</td>
                  <td align="right" style="padding:12px;font-size:18px;font-weight:700;color:#b45309;">$($duplicateSmtpRows.Count)</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 24px 0 24px;">
              <div style="font-size:15px;font-weight:700;color:#111827;margin-bottom:8px;">Source dataset</div>
              <div style="font-family:Consolas,'Courier New',monospace;font-size:12px;line-height:18px;background:#f8fafc;border:1px solid #d9e2ec;border-radius:4px;padding:10px 12px;color:#334155;word-break:break-all;">$htmlSourceUsersCsv</div>
            </td>
          </tr>
          <tr>
            <td style="padding:18px 24px 22px 24px;">
              <div style="font-size:15px;font-weight:700;color:#111827;margin-bottom:8px;">Generated output files</div>
              <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid #d9e2ec;">
                <tr>
                  <td style="width:150px;background:#f8fafc;border-bottom:1px solid #eef2f7;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Duplicate UPN</td>
                  <td style="border-bottom:1px solid #eef2f7;padding:10px 12px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;">$htmlDuplicateUpnCsv</td>
                </tr>
                <tr>
                  <td style="width:150px;background:#f8fafc;padding:10px 12px;font-size:13px;font-weight:700;color:#334155;">Duplicate SMTP</td>
                  <td style="padding:10px 12px;font-family:Consolas,'Courier New',monospace;font-size:12px;color:#334155;word-break:break-all;">$htmlDuplicateSmtpCsv</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="background:#f8fafc;border-top:1px solid #d9e2ec;padding:12px 24px;color:#64748b;font-size:12px;line-height:18px;">
              This automated message was generated by SmartM365. Use the attached/exported CSV paths above as the source of truth for remediation.
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"@
                        $duplicateNotificationTo = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue '')
                        if ([string]::IsNullOrWhiteSpace($duplicateNotificationTo)) {
                            throw 'Duplicate identity notification requires To in configuration. ErrorMailTo is reserved for error notifications.'
                        }
                        Send-SmartM365AdInventoryEmailHtmlReport -Subject $emailSubject -BodyHtml $emailBody -To $duplicateNotificationTo
                        Set-Content -LiteralPath $DuplicateNotificationLastSentFilePath -Value $todayStamp -Encoding UTF8
                        WriteLog -Message ("Duplicate identity notification sent. Last-sent marker updated: {0}" -f $DuplicateNotificationLastSentFilePath)
                    }
                }
                elseif (-not $EnableDuplicateNotification) {
                    WriteLog -Message "Duplicate identity notification skipped because EnableDuplicateNotification is disabled."
                }
            }
        }
        catch {
            $duplicateAnalysisError = $_
            WriteLog -Message ("Duplicate user identity analysis failed: {0}" -f $duplicateAnalysisError)

            try {
                $emailSubject = "[ERROR] SmartM365 Active Directory duplicate identity analysis"
                $emailBody = @"
<html>
<body>
    <h3>Active Directory Duplicate Identity Analysis - Error</h3>
    <p><b>Script:</b> $($MyInvocation.MyCommand.Name)</p>
    <p><b>Version:</b> $ScriptVersion</p>
    <p><b>Tenant:</b> $Tenant</p>
    <p><b>Host:</b> $env:COMPUTERNAME</p>
    <p><b>Date:</b> $(Get-Date)</p>
    <p><b>Error:</b></p>
    <pre>$duplicateAnalysisError</pre>
</body>
</html>
"@
                Send-SmartM365AdInventoryEmailHtmlReport -Subject $emailSubject -BodyHtml $emailBody
                WriteLog -Message "Duplicate identity analysis error email notification sent."
            }
            catch {
                WriteLog -Message ("Failed to send duplicate identity analysis error email notification: {0}" -f $_)
            }
        }
    }
    else {
        WriteLog -Message "Duplicate UPN and SMTP proxy address analysis skipped because EnableDuplicateAnalysis is disabled."
    }

    # ------------------------------------------------------
    # DAILY ACTIVE DIRECTORY REPORTS
    # ------------------------------------------------------
    if (-not $DuplicateAnalysisOnly -and -not $SkipDailyReport) {
        Invoke-ActiveDirectoryDailyReport -SourceFolder $OutputPath `
            -ReportOutputPath $OutputPath `
            -LatestFolderPath $destinationRootPath `
            -AllowedOS $DailyReportAllowedOS `
            -TargetDomains $TargetDomains `
            -InactiveDays $DailyReportInactiveDays `
            -UseDailyLock $EnableDailyReportLock `
            -LockRoot $DailyReportLockRoot `
            -LockName $DailyReportLockName | Out-Null
    }
    elseif ($DuplicateAnalysisOnly) {
        WriteLog -Message "Daily Active Directory report skipped in DuplicateAnalysisOnly mode."
    }
    else {
        WriteLog -Message "Daily Active Directory report skipped because -SkipDailyReport was specified."
    }
    # ------------------------------------------------------
    # WEEKLY INVENTORY HISTORY
    # ------------------------------------------------------
    if ($EnableWeeklyHistory) {
        try {
            Save-WeeklyInventoryHistory -SourceFiles @(
                $combinedUsersCsv,
                $combinedComputersCsv,
                $combinedGroupsCsv,
                $combinedOusCsv,
                $combinedContactsCsv,
                $duplicateUpnCsv,
                $duplicateSmtpCsv
            ) -HistoryRootPath $WeeklyHistoryFolderPath -RetentionWeeks $WeeklyHistoryRetentionWeeks
        }
        catch {
            WriteLog -Message ("Weekly AD inventory history failed: {0}" -f $_) -Level "WARNING"
        }
    }
    else {
        WriteLog -Message "Weekly AD inventory history skipped because EnableWeeklyHistory is disabled."
    }

    # ------------------------------------------------------
    # CLEANUP TEMPORARY PER-DOMAIN CSV FILES
    # ------------------------------------------------------
    if (-not $DuplicateAnalysisOnly) {
        Remove-TemporaryInventoryFolder -TempFolder $tempFolder -BaseFolder $baseFolder
    }
    else {
        WriteLog -Message "Temporary per-domain cleanup skipped in DuplicateAnalysisOnly mode."
    }

    # ------------------------------------------------------
    # CLEANUP OLD FILES
    # ------------------------------------------------------
    Remove-OldFiles -Path $OutputPath -Filter "*.csv" -OlderThanDays 30
    Remove-OldFiles -Path $OutputPath -Filter "*.log" -OlderThanDays 30

    WriteLog -Message ("{0} completed successfully." -f $TaskName)
}
catch {
    $globalError = $_
    WriteLog -Message ("Fatal error in script: {0}" -f $globalError)

    try {
        $emailSubject = "[ERROR] $TaskName"
        $emailBody = @"
<html>
<body>
    <h3>Active Directory Inventory - Global Error</h3>
    <p><b>Script:</b> $($MyInvocation.MyCommand.Name)</p>
    <p><b>Version:</b> $ScriptVersion</p>
    <p><b>Host:</b> $env:COMPUTERNAME</p>
    <p><b>Date:</b> $(Get-Date)</p>
    <p><b>Error:</b></p>
    <pre>$globalError</pre>
</body>
</html>
"@
        Send-SmartM365AdInventoryEmailHtmlReport -Subject $emailSubject -BodyHtml $emailBody
        WriteLog -Message "Global error email notification sent."
    }
    catch {
        WriteLog -Message ("Failed to send global error email notification: {0}" -f $_)
    }
}
finally {
    try {
        WriteLog -Message ("Stopping transcript for script '{0}'" -f $MyInvocation.MyCommand.Name)
        Stop-Transcript | Out-Null
        try {
            $smartM365TranscriptPath = $null
            $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue
            if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
                $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
            }
            else {
                $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue
                if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
                    $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
                }
            }
            if ($smartM365TranscriptPath) {
                Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath
            }
        }
        catch {
        }
    }
    catch {
        Write-Host ("Failed to stop transcript: {0}" -f $_) -ForegroundColor Yellow
    }

    try {
        Complete-SmartM365ExecutionContext -Status Auto -ErrorRecord $globalError
    }
    catch {
        Write-Host ("Failed to write execution summary: {0}" -f $_) -ForegroundColor Yellow
    }
}
