<#
.SYNOPSIS
Exports Microsoft 365 Copilot user usage from Microsoft Graph reports.

.DESCRIPTION
Downloads the official Microsoft 365 Copilot usage user detail report, normalizes
the v2 and v1 schemas into a stable SmartFinOps-oriented CSV, and publishes the
result to tenant-isolated DATA-ALL, DATA-LAST, WeeklyHistory, and SharePoint when
enabled.

Version v2 is preferred. A compatibility fallback to v1 is used only when v2 is
not available for the requested period or Microsoft Graph rejects the v2 report
version. Version v1 does not expose prompt counters, active usage days, Microsoft
365 app, Edge, agent, or Copilot Chat work/web activity dates; those stable CSV
columns remain empty after a v1 fallback.

This collector never requests or exports prompt text.

.VERSION
1.0.0

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permission: Reports.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
Official API:
https://learn.microsoft.com/microsoft-365/copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [ValidateSet('D7', 'D28', 'D30', 'D90', 'D180')]
    [string]$Period = 'D180',
    [ValidateSet('v1', 'v2')]
    [string]$ReportVersion = 'v2',
    [string]$OutputPath,
    [string]$LatestCsvFolderPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [switch]$ValidateOnly,
    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxItems = 0
)

if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
}

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'
$TaskName = "SmartM365-CopilotUsage-Inventory v$ScriptVersion"
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$rawReportPath = $null

if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "This script requires PowerShell 7 or later. Current version: $($PSVersionTable.PSVersion)"
}

$tenantContextPath = & {
    $directory = $PSScriptRoot
    while ($directory) {
        $candidates = @(
            (Join-Path -Path $directory -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $directory -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($candidate in $candidates) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $directory -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $directory) { break }
        $directory = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}

. $tenantContextPath
$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

$localConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-CopilotUsage-Inventory.local.json'
$localConfigTemplatePath = "$localConfigPath.template"
Initialize-SmartM365LocalJsonFromTemplate `
    -Path $localConfigPath `
    -TemplatePath $localConfigTemplatePath `
    -ConfigDescription 'SmartM365 Copilot usage local configuration' | Out-Null
$script:SmartM365LocalConfig = Read-SmartM365JsonConfig -Path $localConfigPath -Required

function Resolve-SmartM365TokenValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $resolved = $Value
    for ($index = 0; $index -lt 10; $index++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $tokenValue = Resolve-SmartM365TokenValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-SmartM365ConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()]$DefaultValue = $null
    )

    $localValue = $null
    if ($script:SmartM365LocalConfig -is [System.Collections.IDictionary] -and $script:SmartM365LocalConfig.Contains($Name)) {
        $localValue = $script:SmartM365LocalConfig[$Name]
    }
    elseif ($null -ne $script:SmartM365LocalConfig) {
        $localProperty = $script:SmartM365LocalConfig.PSObject.Properties[$Name]
        if ($null -ne $localProperty) { $localValue = $localProperty.Value }
    }

    if ($null -ne $localValue) {
        $localText = [string]$localValue
        if (-not [string]::IsNullOrWhiteSpace($localText) -and $localText -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
            return Resolve-SmartM365TokenValue -Value $localValue
        }
    }

    $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    return Resolve-SmartM365TokenValue -Value $property.Value
}

function Get-SourcePropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$Row,
        [Parameter(Mandatory)][string[]]$Names
    )

    if ($null -eq $Row) { return $null }
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }
    return $null
}

function ConvertTo-CopilotDateText {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $parsedDate = [datetime]::MinValue
    if ([datetime]::TryParse(
            $text,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AssumeUniversal,
            [ref]$parsedDate
        )) {
        return $parsedDate.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $text
}

function ConvertTo-CopilotCounter {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    $parsedValue = [long]0
    if ([long]::TryParse(
            ([string]$Value).Trim(),
            [System.Globalization.NumberStyles]::Integer,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$parsedValue
        )) {
        return $parsedValue
    }
    return ([string]$Value).Trim()
}

function Get-CopilotUsageColumns {
    [CmdletBinding()]
    param()

    return @(
        'RunId',
        'ReportVersion',
        'ReportPeriod',
        'ReportRefreshDate',
        'UserPrincipalName',
        'DisplayName',
        'LastActivityDate',
        'Microsoft365CopilotLastActivityDate',
        'CopilotChatLastActivityDate',
        'CopilotChatWorkLastActivityDate',
        'CopilotChatWebLastActivityDate',
        'TeamsCopilotLastActivityDate',
        'WordCopilotLastActivityDate',
        'ExcelCopilotLastActivityDate',
        'PowerPointCopilotLastActivityDate',
        'OutlookCopilotLastActivityDate',
        'OneNoteCopilotLastActivityDate',
        'LoopCopilotLastActivityDate',
        'EdgeLastActivityDate',
        'CopilotAgentLastActivityDate',
        'ActiveUsageDays',
        'PromptsSubmittedAllApps',
        'CopilotChatWorkPromptsSubmitted',
        'CopilotChatWebPromptsSubmitted'
    )
}

function ConvertFrom-CopilotUsageReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][ValidateSet('v1', 'v2')][string]$EffectiveReportVersion
    )

    foreach ($row in $Rows) {
        [pscustomobject][ordered]@{
            RunId = $runId
            ReportVersion = $EffectiveReportVersion
            ReportPeriod = $Period
            ReportRefreshDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @('Report Refresh Date', 'reportRefreshDate'))
            UserPrincipalName = [string](Get-SourcePropertyValue -Row $row -Names @('User Principal Name', 'UserPrincipalName', 'userPrincipalName'))
            DisplayName = [string](Get-SourcePropertyValue -Row $row -Names @('Display Name', 'DisplayName', 'displayName'))
            LastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @('Last Activity Date', 'lastActivityDate'))
            Microsoft365CopilotLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @(
                    'Microsoft 365 Copilot Last Activity Date',
                    'Microsoft 365 App Last Activity Date',
                    'Microsoft365 Copilot Last Activity Date',
                    'microsoft365CopilotLastActivityDate',
                    'microsoft365AppLastActivityDate'
                ))
            CopilotChatLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @(
                    'Copilot Chat Last Activity Date',
                    'copilotChatLastActivityDate'
                ))
            CopilotChatWorkLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @(
                    'Copilot Chat (work) Last Activity Date',
                    'Copilot Chat Work Last Activity Date',
                    'copilotChatWorkLastActivityDate'
                ))
            CopilotChatWebLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @(
                    'Copilot Chat (web) Last Activity Date',
                    'Copilot Chat Web Last Activity Date',
                    'copilotChatWebLastActivityDate'
                ))
            TeamsCopilotLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @(
                    'Microsoft Teams Copilot Last Activity Date',
                    'Teams Copilot Last Activity Date',
                    'microsoftTeamsCopilotLastActivityDate',
                    'teamsCopilotLastActivityDate'
                ))
            WordCopilotLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @('Word Copilot Last Activity Date', 'wordCopilotLastActivityDate'))
            ExcelCopilotLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @('Excel Copilot Last Activity Date', 'excelCopilotLastActivityDate'))
            PowerPointCopilotLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @('PowerPoint Copilot Last Activity Date', 'powerPointCopilotLastActivityDate'))
            OutlookCopilotLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @('Outlook Copilot Last Activity Date', 'outlookCopilotLastActivityDate'))
            OneNoteCopilotLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @('OneNote Copilot Last Activity Date', 'oneNoteCopilotLastActivityDate'))
            LoopCopilotLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @('Loop Copilot Last Activity Date', 'loopCopilotLastActivityDate'))
            EdgeLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @(
                    'Edge Last Activity Date',
                    'Microsoft Edge Last Activity Date',
                    'edgeLastActivityDate',
                    'microsoftEdgeLastActivityDate'
                ))
            CopilotAgentLastActivityDate = ConvertTo-CopilotDateText (Get-SourcePropertyValue -Row $row -Names @(
                    'Copilot Agent Last Activity Date',
                    'Copilot Agents Last Activity Date',
                    'copilotAgentLastActivityDate'
                ))
            ActiveUsageDays = ConvertTo-CopilotCounter (Get-SourcePropertyValue -Row $row -Names @(
                    'Active Usage Days',
                    'Active Days',
                    'activeUsageDays',
                    'activeDays'
                ))
            PromptsSubmittedAllApps = ConvertTo-CopilotCounter (Get-SourcePropertyValue -Row $row -Names @(
                    'Prompts Submitted for All Apps',
                    'Prompts submitted (any app)',
                    'Prompts Submitted (Any App)',
                    'All Apps Prompts Submitted',
                    'Prompts Submitted',
                    'promptsSubmittedAllApps',
                    'allAppsPromptsSubmitted'
                ))
            CopilotChatWorkPromptsSubmitted = ConvertTo-CopilotCounter (Get-SourcePropertyValue -Row $row -Names @(
                    'Prompts Submitted for Copilot Chat (work)',
                    'Copilot Chat (work) Prompts Submitted',
                    'Copilot Chat Work Prompts Submitted',
                    'copilotChatWorkPromptsSubmitted'
                ))
            CopilotChatWebPromptsSubmitted = ConvertTo-CopilotCounter (Get-SourcePropertyValue -Row $row -Names @(
                    'Prompts Submitted for Copilot Chat (web)',
                    'Copilot Chat (web) Prompts Submitted',
                    'Copilot Chat Web Prompts Submitted',
                    'copilotChatWebPromptsSubmitted'
                ))
        }
    }
}

function Get-CopilotUsageReportUri {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('v1', 'v2')][string]$Version,
        [Parameter(Mandatory)][string]$RequestedPeriod
    )

    return "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$RequestedPeriod',version='$Version')?`$format=text/csv"
}

function Get-GraphHttpStatusCode {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    foreach ($candidate in @(
            $ErrorRecord.Exception.ResponseStatusCode,
            $ErrorRecord.Exception.StatusCode,
            $(if ($ErrorRecord.Exception.Response) { $ErrorRecord.Exception.Response.StatusCode })
        )) {
        if ($null -eq $candidate) { continue }
        try { return [int]$candidate } catch {}
    }
    return $null
}

function Test-CopilotV2FallbackEligible {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    $statusCode = Get-GraphHttpStatusCode -ErrorRecord $ErrorRecord
    if ($statusCode -in @(400, 404, 405, 406, 422, 501)) { return $true }
    $message = [string]$ErrorRecord.Exception.Message
    return ($message -match '(?i)(unsupported|not supported|invalid).{0,80}(report )?version|version.{0,80}(unsupported|not supported|invalid)')
}

function Resolve-CopilotInitialReportVersion {
    [CmdletBinding()]
    param()

    if ($ReportVersion -eq 'v1' -and $Period -eq 'D28') {
        throw "Report version v1 does not support period D28. Use -ReportVersion v2 or choose D7, D30, D90, or D180."
    }
    if ($ReportVersion -eq 'v2' -and $Period -eq 'D30') {
        WriteLog 'Microsoft Graph report v2 does not support D30. Falling back to v1, which supports D30.' 'WARNING'
        return 'v1'
    }
    return $ReportVersion
}

function Invoke-CopilotUsageReportDownload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][ValidateSet('v1', 'v2')][string]$InitialVersion
    )

    $effectiveVersion = $InitialVersion
    $uri = Get-CopilotUsageReportUri -Version $effectiveVersion -RequestedPeriod $Period
    try {
        WriteLog ("Downloading Microsoft 365 Copilot usage report. Version={0}; Period={1}" -f $effectiveVersion, $Period) 'INFO'
        Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $DestinationPath -ProgressAction SilentlyContinue -ErrorAction Stop | Out-Null
    }
    catch {
        if ($effectiveVersion -ne 'v2' -or $Period -eq 'D28' -or -not (Test-CopilotV2FallbackEligible -ErrorRecord $_)) {
            throw
        }

        $statusCode = Get-GraphHttpStatusCode -ErrorRecord $_
        WriteLog ("Microsoft Graph rejected report v2. Falling back to v1. Status={0}; Error={1}" -f $statusCode, $_.Exception.Message) 'WARNING'
        if (Test-Path -LiteralPath $DestinationPath) {
            Remove-Item -LiteralPath $DestinationPath -Force -ErrorAction SilentlyContinue
        }
        $effectiveVersion = 'v1'
        $uri = Get-CopilotUsageReportUri -Version $effectiveVersion -RequestedPeriod $Period
        Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $DestinationPath -ProgressAction SilentlyContinue -ErrorAction Stop | Out-Null
    }

    return $effectiveVersion
}

function Test-CopilotEmptyCsvSchema {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Columns)

    $validationPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("M365_Copilot_UserUsage_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        Write-SmartM365CsvAtomically -Data @() -Path $validationPath -Columns $Columns -NoTenantKey
        $header = Get-Content -LiteralPath $validationPath -TotalCount 1
        $expectedHeader = (($Columns | ForEach-Object { '"{0}"' -f $_ }) -join ',')
        if ($header -ne $expectedHeader) {
            throw "Empty CSV schema validation failed. Expected '$expectedHeader'; got '$header'."
        }
    }
    finally {
        if (Test-Path -LiteralPath $validationPath) {
            Remove-Item -LiteralPath $validationPath -Force -ErrorAction SilentlyContinue
        }
    }
}

$dataAllRoot = Resolve-SmartM365TokenValue -Value (Get-SmartM365ConfigValue -Name 'DataAllRootPath' -DefaultValue '')
$logAllRoot = Resolve-SmartM365TokenValue -Value (Get-SmartM365ConfigValue -Name 'LogAllRootPath' -DefaultValue '')
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $LatestCsvFolderPath = Resolve-SmartM365TokenValue -Value (Get-SmartM365ConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $configuredOutputPath = Resolve-SmartM365TokenValue -Value (Get-SmartM365ConfigValue -Name 'OutputPath' -DefaultValue '')
    $OutputPath = if ([string]::IsNullOrWhiteSpace($configuredOutputPath)) {
        Join-Path -Path $dataAllRoot -ChildPath 'M365\Usage\Copilot'
    }
    else {
        $configuredOutputPath
    }
}

$runOutputRoot = Join-Path -Path $OutputPath -ChildPath $runId
$logFileBaseName = 'SmartM365-CopilotUsage-Inventory'
$logFolder = Join-Path -Path $logAllRoot -ChildPath $logFileBaseName
$logPath = Join-Path -Path $logFolder -ChildPath ("{0}_{1}.log" -f $logFileBaseName, $runId)
$rawReportPath = Join-Path -Path $runOutputRoot -ChildPath ("M365_Copilot_UserUsage_Source_{0}.csv" -f $runId)

$modulePath = Join-Path -Path ([string](Get-SmartM365ConfigValue -Name 'SmartM365RootPath' -DefaultValue (Split-Path -Path $PSScriptRoot -Parent))) -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $modulePath -MinimumVersion '1.0.41' -Force -ErrorAction Stop

$global:RetentionMaxCSV = [int](Get-SmartM365ConfigValue -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-SmartM365ConfigValue -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-SmartM365ConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-SmartM365ConfigValue -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-SmartM365ConfigValue -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-SmartM365ConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-SmartM365ConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue ''

$global:SmartM365ExecutionStartTime = Get-Date
$global:SmartM365ExecutionSummaryWritten = $false
$global:SmartM365ScriptName = $TaskName
Set-SmartM365CoreContext `
    -RunId $runId `
    -RunOutputRoot $runOutputRoot `
    -LatestOutputRoot $LatestCsvFolderPath `
    -LogPath $logPath `
    -RetentionMaxCsv $global:RetentionMaxCSV `
    -RetentionMaxLogs $global:RetentionMaxLogs
$global:LogTextFile = $logPath
$global:logTextFile = $logPath
$global:LogPath = $logFolder
$global:logTranscriptFile = Join-Path -Path $logFolder -ChildPath ("{0}_{1}_Transcript.log" -f $logFileBaseName, $runId)

$copilotColumns = @(Get-CopilotUsageColumns)
$global:SmartM365CsvValidationRules = if ($global:SmartM365CsvValidationRules -is [hashtable]) {
    $global:SmartM365CsvValidationRules
}
else {
    @{}
}
Add-SmartM365CsvValidationRule `
    -Rules $global:SmartM365CsvValidationRules `
    -BaseFileName 'M365_Copilot_UserUsage' `
    -CriticalFields @('TenantKey', 'RunId', 'ReportVersion', 'ReportPeriod', 'ReportRefreshDate', 'UserPrincipalName') `
    -RequiredColumns @($copilotColumns) `
    -AllowEmptyDataset

function Stop-SmartM365CopilotUsageTranscript {
    [CmdletBinding()]
    param()

    try {
        Stop-Transcript | Out-Null
        if ($global:logTranscriptFile -and (Get-Command Update-SmartM365TimestampedTranscript -ErrorAction SilentlyContinue)) {
            Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile
        }
    }
    catch {}
}

try {
    foreach ($folder in @($runOutputRoot, $LatestCsvFolderPath, $logFolder)) {
        if ([string]::IsNullOrWhiteSpace([string]$folder)) {
            throw 'A required SmartM365 output or log path resolved to an empty value.'
        }
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        }
    }

    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    WriteLog ("Starting {0}. Tenant={1}; Period={2}; RequestedReportVersion={3}; RunId={4}" -f $TaskName, $Tenant, $Period, $ReportVersion, $runId) 'INFO'

    $initialReportVersion = Resolve-CopilotInitialReportVersion

    if ($ValidateOnly) {
        Invoke-SmartM365Preflight `
            -ScriptName $TaskName `
            -RequiredModules @('Microsoft.Graph.Authentication') `
            -RequiredCommands @('Get-MgContext', 'Invoke-MgGraphRequest') `
            -OutputPaths @($runOutputRoot, $LatestCsvFolderPath) | Out-Null

        $null = Get-CopilotUsageReportUri -Version $initialReportVersion -RequestedPeriod $Period
        Test-CopilotEmptyCsvSchema -Columns (@('TenantKey') + $copilotColumns)
        $validationSummary = "Tenant=$Tenant; Period=$Period; RequestedVersion=$ReportVersion; InitialVersion=$initialReportVersion; EmptySchemaColumns=$($copilotColumns.Count + 1); GraphPermission=Reports.Read.All; No cloud request performed"
        WriteLog ("Validation completed. {0}" -f $validationSummary) 'SUCCESS'
        Stop-SmartM365CopilotUsageTranscript
        Complete-SmartM365ExecutionContext -Status Auto
        return
    }

    if ($Connect -or $null -eq (Get-MgContext -ErrorAction SilentlyContinue)) {
        $connectParams = @{
            Graph = $true
            ExchangeOnline = $false
            GraphScopes = @('Reports.Read.All')
        }
        if (-not $InteractiveAuth) {
            $connectParams.AppId = [string](Get-SmartM365ConfigValue -Name 'AppId' -DefaultValue '')
            $connectParams.Thumbprint = [string](Get-SmartM365ConfigValue -Name 'Thumbprint' -DefaultValue '')
            $connectParams.TenantId = [string](Get-SmartM365ConfigValue -Name 'TenantId' -DefaultValue '')
        }

        $connectResult = Connect-SmartM365CloudSession @connectParams
        if (-not $connectResult.GraphConnected) {
            throw 'Microsoft Graph connection failed. Check app-only certificate settings or use -InteractiveAuth.'
        }
    }

    Invoke-SmartM365Preflight `
        -ScriptName $TaskName `
        -RequiredModules @('Microsoft.Graph.Authentication') `
        -RequiredCommands @('Get-MgContext', 'Invoke-MgGraphRequest') `
        -RequiredGraphApplicationPermissions @('Reports.Read.All') `
        -OutputPaths @($runOutputRoot, $LatestCsvFolderPath) | Out-Null

    $effectiveReportVersion = Invoke-CopilotUsageReportDownload -DestinationPath $rawReportPath -InitialVersion $initialReportVersion
    $sourceRows = if (Test-Path -LiteralPath $rawReportPath) {
        @(Import-Csv -LiteralPath $rawReportPath)
    }
    else {
        @()
    }
    $sourceRows = @(Limit-SmartM365RowsForMaxItems -Data $sourceRows)
    $normalizedRows = @(ConvertFrom-CopilotUsageReport -Rows $sourceRows -EffectiveReportVersion $effectiveReportVersion)

    $exportResult = Export-SmartM365Csv `
        -BaseFileName 'M365_Copilot_UserUsage' `
        -OutputPath $runOutputRoot `
        -GlobalPath $LatestCsvFolderPath `
        -Data $normalizedRows `
        -Columns $copilotColumns

    $summary = "Tenant=$Tenant; Period=$Period; ReportVersion=$effectiveReportVersion; Users=$($normalizedRows.Count); LatestCsv=$($exportResult.LatestPath)"
    WriteLog ("Microsoft 365 Copilot usage inventory completed. {0}" -f $summary) 'SUCCESS'
    Send-SmartM365TeamsNotification `
        -Level SUCCESS `
        -Channel Infos `
        -Title 'SmartM365 Microsoft 365 Copilot usage inventory completed' `
        -Message $summary `
        -ResultSummary $summary `
        -Facts @{
            Tenant = $Tenant
            Period = $Period
            ReportVersion = $effectiveReportVersion
            Users = $normalizedRows.Count
            LatestCsv = $exportResult.LatestPath
            RunId = $runId
        } | Out-Null

    Stop-SmartM365CopilotUsageTranscript
    Complete-SmartM365ExecutionContext -Status Auto
}
catch {
    $message = $_.Exception.Message
    WriteLog ("Microsoft 365 Copilot usage inventory failed: {0}" -f $message) 'ERROR'
    try {
        Send-SmartM365TeamsNotification `
            -Level ERROR `
            -Channel Alerts `
            -Title 'SmartM365 Microsoft 365 Copilot usage inventory failed' `
            -Message $message `
            -Facts @{
                Tenant = $Tenant
                Period = $Period
                RequestedReportVersion = $ReportVersion
                RunId = $runId
                LogPath = $logPath
                OutputPath = $runOutputRoot
            } | Out-Null
    }
    catch {
        WriteLog ("Teams alert notification failed: {0}" -f $_.Exception.Message) 'WARNING'
    }
    Stop-SmartM365CopilotUsageTranscript
    try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage 'CopilotUsageInventory' } catch {}
    throw
}
finally {
    if ($rawReportPath -and (Test-Path -LiteralPath $rawReportPath)) {
        Remove-Item -LiteralPath $rawReportPath -Force -ErrorAction SilentlyContinue
    }
    try { RemoveOldFiles -Path $OutputPath -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile } catch {}
    try { RemoveOldFiles -Path $logFolder -Filter '*.log' -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile } catch {}
    try { Stop-SmartM365CopilotUsageTranscript } catch {}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBUOHhL9zLHz/mf
# gnU4uFqFLjKvl3j3LhuzBVTVpac0M6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIES7GCmWNN6oFAkEwmAR40sKYCVJfdDF1Zp7FXWbpc6MA0GCSqG
# SIb3DQEBAQUABIIBgI+2kwj408mIJW/U57+5UTYd/140rzvQBZmVLxRMFI2p/hBZ
# 4j4HHSrk9i/t/J3FPYCsf+3q37jI3G79Aq5myRrC6k4l+yh16KflyJ7E/0kIul9f
# z9m636PelIVwRycu4lAxzsTuH04E1aXR+7vbrG0gxkBVB15mne7xjJSxhbnmXkvV
# Hj+/op86meQrRE6MokJiPY9Hf79KqppmPT4uXpoxdRqvY87ZB/4CCHgVONtvPpJp
# UM85XrKov7ztXcZ+Roh7X7HWd4WoLXpAYIFzoxrYQZa6DmKYyMj0o5TX1ACOOlv0
# 62ixBRzeK2D9wcAvX8u9ZNCAFB5NgrG0lXi7w+9lpJOORXBfoiBUnYVwBtu8oQcN
# c3KAQBfDU57Tn+5+YaSmUmL7usPZDSR6OLMjUgvtN1zxXSaS0m/QBeBR1BU5mVVg
# eehj0HBYYr7GFd63Go40BZACECrk5gJR5PugqeRU5CiPBbVjjB8dOPI9N3YjOkiS
# OrIuN/bBekvO7HjnJqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTgyMDEx
# NTZaMC8GCSqGSIb3DQEJBDEiBCBknDD+BXtryXccqNt+EnaBtpMG2YmHmC6/aPAL
# zOiUGTANBgkqhkiG9w0BAQEFAASCAgB7Wb+487Y8LDQe7lxToY++s1zIxRErq5Gg
# jmDduVf8rinYrcTeFfgEWhlZsLp4flDes0Lsan/U0UpvcUjnmhaZ/r6frx6E//TF
# TfwGse4vZKI60a5jj0lUo3TYW9ISGJbH4nL8K2yyfZ2aE+OVIs59v8AdmlZeswSO
# zPhmQJwEjaYRxY6WYv4ss68727hiAaE7kKz7GTOBBghBbIAordM8kFOBMwrIHrmH
# /5CWi/N++OdRqDRnUSJ3C/pexAybwGE3Yk/f0vbfWXkXz0wkLMT/ojRCk01s0eLr
# IREeG4c+WrHN+2Yg8rEmd2uy2iNBpSUsiK0CFuDZN6RLovgX/axTOk/ttxoyKZuk
# dq1BKhw3FalF26+UPEizhb+xyCtC4r1alxWgq3EFKL8KnklZNCvccj9VZAdDCnFa
# u1hLyylBdYIXyBPhkt58ljh1FUOD9Lxkp8HM6DHfeYZrwE33RhihAnNuoFOXWEst
# iyyF4kab0YjnT7GDD4i8zcGSoDaoso3KbiRzbo0gXp78TMHQNis6mgTiemICK6g0
# y42vae29OsL9VWyT1XKxdZTJ7amjM2hYf9omRXEjhSWDlpPeCennJUS9/UuxIvyI
# jk/uZVehnvyaG4HgsSkE+KyClWNAWIi4aqKaQBykkQ+I38nogW5UREsdd701oZXp
# m2ikoMQ0Ug==
# SIG # End signature block
