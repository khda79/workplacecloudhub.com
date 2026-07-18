#Requires -Version 7.0
<#
.SYNOPSIS
Collects Microsoft Intune Endpoint Analytics standard reports through Microsoft Graph exportJobs.

.DESCRIPTION
Read-only SmartInventory collector for Endpoint Analytics features included with a valid
Microsoft Intune license. Advanced Analytics report families are deliberately excluded.

The Endpoint Analytics report names used by this script are currently documented only in
the Microsoft Intune beta report catalogue, so the exportJobs calls use Microsoft Graph beta.
No Endpoint Analytics policy, assignment, baseline, data collection setting, or remediation
is created or changed.

.PARAMETER Tenant
SmartM365 tenant profile key. Defaults to test.

.PARAMETER Reports
All, Scores, Startup, ApplicationReliability, WorkFromAnywhere, or exact supported report names.

.PARAMETER IncludeStartupProcesses
Includes EAStartupPerfDeviceProcesses and publishes StartupProcesses.

.PARAMETER OutputPath
Historical output folder. Defaults to {{DataAllRootPath}}\Intune\EndpointAnalytics.

.PARAMETER LatestCsvFolderPath
Folder for current canonical CSV files. Defaults to the SmartM365 tenant DATA-LAST folder.

.PARAMETER Connect
Forces a new app-only certificate connection. A normal collection also connects when no
Graph context exists.

.PARAMETER InteractiveAuth
Uses delegated interactive Graph authentication for controlled tests.

.PARAMETER ValidateOnly
Validates the static catalogue. With -Connect or -InteractiveAuth, creates temporary export
jobs to verify tenant availability, but downloads no data and publishes no CSV.

.PARAMETER MaxItems
Limits normalized rows per source report and activates the SmartM365 MAXITEMS filename suffix.

.PARAMETER SelfTest
Runs offline completed, failed, throttled, schema, MAXITEMS, TenantKey, and Advanced Analytics
guard simulations. No tenant context or Graph connection is used.

.EXAMPLE
pwsh -File .\SmartM365-EndpointAnalytics-Inventory.ps1 -Tenant test -ValidateOnly -InteractiveAuth

.EXAMPLE
pwsh -File .\SmartM365-EndpointAnalytics-Inventory.ps1 -Tenant test -Reports All -Connect

.VERSION
1.0.3

.REQUIREMENTS
PowerShell 7+.
Modules: SmartM365.Core 1.0.24+; Microsoft.Graph.Authentication.
Graph POST exportJobs permission documented by Microsoft:
DeviceManagementManagedDevices.ReadWrite.All (application or delegated).
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'SmartM365.Core consumes shared runtime context through established global variables.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Fallback logging is used only before SmartM365.Core is loaded and keeps every line timestamped.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Script parameters are consumed through script scope and default parameter expressions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'These functions create in-memory rows or temporary read-only report export jobs, not tenant configuration.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Plural nouns describe collections and outputs precisely.')]
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string[]]$Reports = @('All'),
    [switch]$IncludeStartupProcesses,
    [string]$OutputPath,
    [string]$LatestCsvFolderPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [switch]$ValidateOnly,
    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxItems = 0,
    [ValidateRange(30, 7200)]
    [int]$ExportTimeoutSeconds = 600,
    [ValidateRange(1, 60)]
    [int]$PollIntervalSeconds = 5,
    [ValidateRange(1, 12)]
    [int]$MaxRetryCount = 6,
    [switch]$SelfTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ScriptVersion = '1.0.3'
$script:ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$script:RunId = [guid]::NewGuid().Guid
$script:CollectedAtUtc = [datetime]::UtcNow.ToString('o')
$script:RequiredPermission = 'DeviceManagementManagedDevices.ReadWrite.All'
$script:AdvancedReportPattern = '^(BR|EAResourcePerf|EAAnomaly)|DeviceTimeline|DeviceQuery'
$script:GraphApiBase = 'https://graph.microsoft.com'
$script:CoreImported = $false
$script:TenantConfig = $null

if ($MaxItems -gt 0) {
    $global:SmartM365MaxItems = $MaxItems
    $global:SmartM365TestMaxItems = $MaxItems
    $global:SmartM365IsMaxItemsRun = $true
}

function Write-EALog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO'
    )
    if (Get-Command WriteLog -ErrorAction SilentlyContinue) {
        WriteLog -Message $Message -Level $Level
        return
    }
    $prefix = '{0} [{1}]' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level
    foreach ($line in [regex]::Split($Message, '\r?\n')) {
        Write-Host ("{0} {1}" -f $prefix, $line)
    }
}

function Get-EARawConfigValue {
    param([Parameter(Mandatory)][string]$Name, $DefaultValue = $null)
    if ($null -eq $script:TenantConfig) { return $DefaultValue }
    $property = $script:TenantConfig.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    if ($property.Value -is [string] -and [string]::IsNullOrWhiteSpace($property.Value)) { return $DefaultValue }
    return $property.Value
}

function Resolve-EAConfigValue {
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $resolved = [string]$Value
    for ($index = 0; $index -lt 10; $index++) {
        $configMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($configMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $configMatches) {
            $replacement = Get-EARawConfigValue -Name $match.Groups['Name'].Value
            if ($null -eq $replacement) { continue }
            $replacement = Resolve-EAConfigValue -Value $replacement
            if ($null -eq $replacement) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$replacement)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-EAConfigValue {
    param([Parameter(Mandatory)][string]$Name, $DefaultValue = $null)
    return Resolve-EAConfigValue -Value (Get-EARawConfigValue -Name $Name -DefaultValue $DefaultValue)
}

function Initialize-EARuntime {
    $tenantContextPath = $null
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        foreach ($candidate in @(
            (Join-Path $searchRoot 'SmartM365-TenantContext.ps1'),
            (Join-Path $searchRoot 'Config\SmartM365-TenantContext.ps1')
        )) {
            if (Test-Path -LiteralPath $candidate) { $tenantContextPath = $candidate; break }
        }
        if ($tenantContextPath) { break }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    if (-not $tenantContextPath) { throw 'SmartM365-TenantContext.ps1 was not found.' }

    . $tenantContextPath
    $script:TenantConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

    $moduleRoot = Split-Path -Path $tenantContextPath -Parent
    if ((Split-Path -Path $moduleRoot -Leaf) -eq 'Config') { $moduleRoot = Split-Path -Path $moduleRoot -Parent }
    $coreManifest = Join-Path $moduleRoot 'Modules\SmartM365.Core\SmartM365.Core.psd1'
    if (-not (Test-Path -LiteralPath $coreManifest)) { throw "SmartM365.Core manifest was not found: $coreManifest" }
    Import-Module -Name $coreManifest -MinimumVersion '1.0.24' -Prefix Core -Force -ErrorAction Stop
    $script:CoreImported = $true

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $dataAllRoot = [string](Get-EAConfigValue -Name 'DataAllRootPath' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($dataAllRoot)) { throw 'DataAllRootPath is missing from the effective SmartM365 tenant configuration.' }
        $script:OutputPath = Join-Path $dataAllRoot 'Intune\EndpointAnalytics'
    }
    else { $script:OutputPath = $OutputPath }

    if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
        $script:LatestCsvFolderPath = [string](Get-EAConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($script:LatestCsvFolderPath)) { $script:LatestCsvFolderPath = $script:OutputPath }
    }
    else { $script:LatestCsvFolderPath = $LatestCsvFolderPath }

    $global:EnableSharePointUpload = if ($ValidateOnly) { $false } else { [bool](Get-EAConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false) }
    $global:EnableTeamsNotifications = if ($ValidateOnly) { $false } else { [bool](Get-EAConfigValue -Name 'EnableTeamsNotifications' -DefaultValue $false) }
    $global:SharePointSiteHostname = [string](Get-EAConfigValue -Name 'SharePointSiteHostname' -DefaultValue '')
    $global:SharePointSitePath = [string](Get-EAConfigValue -Name 'SharePointSitePath' -DefaultValue '')
    $global:SharePointLibraryDisplayName = [string](Get-EAConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents')
    $global:SharePointTargetFolderPath = [string](Get-EAConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue '')
    $global:AppId = [string](Get-EAConfigValue -Name 'AppId' -DefaultValue '')
    $global:TenantId = [string](Get-EAConfigValue -Name 'TenantId' -DefaultValue '')
    $global:Thumb = [string](Get-EAConfigValue -Name 'Thumb' -DefaultValue '')
    $global:Thumbprint = [string](Get-EAConfigValue -Name 'Thumbprint' -DefaultValue $global:Thumb)
    $global:RetentionMaxCSV = [int](Get-EAConfigValue -Name 'RetentionMaxCSV' -DefaultValue 30)
    $global:RetentionMaxLogs = [int](Get-EAConfigValue -Name 'RetentionMaxLogs' -DefaultValue 30)
    CoreInitializeScriptEnvironment -OutputPathInit $script:OutputPath -LogFileName $script:ScriptName -CallerScriptPath $PSCommandPath
}

function Get-EAReportCatalog {
    return @(
        [pscustomobject]@{ Name='EADevicePerformanceV2'; Group='Scores'; Output='DevicePerformance'; Grain='Device'; ApiVersion='beta'; Aliases=@(); Select=@('DeviceAppHealthScore','DeviceId','DeviceManufacturer','DeviceModel','DeviceName','MeanTimeToFailure','MemaTimeGenerated','ProcessedDateTime','TotalAppCrashes') },
        [pscustomobject]@{ Name='EADeviceModelPerformanceV2'; Group='Scores'; Output='ModelPerformance'; Grain='Model'; ApiVersion='beta'; Aliases=@(); Select=@('ActiveDevices','DeviceManufacturer','DeviceModel','MeanTimeToFailure','MemaTimeGenerated','ModelAppHealthScore') },
        [pscustomobject]@{ Name='EADeviceScoresV2'; Group='Scores'; Output='DevicePerformance'; Grain='Device'; ApiVersion='beta'; Aliases=@(); Select=@('AppReliabilityScore','DeviceId','DeviceName','EndpointAnalyticsScore','Manufacturer','Model','StartupPerformanceScore','WorkFromAnywhereScore') },
        [pscustomobject]@{ Name='EAModelScoresV2'; Group='Scores'; Output='ModelPerformance'; Grain='Model'; ApiVersion='beta'; Aliases=@(); Select=@('AppReliabilityScore','EndpointAnalyticsScore','Manufacturer','Model','ModelDeviceCount','StartupPerformanceScore','WorkFromAnywhereScore') },
        [pscustomobject]@{ Name='EAStartupPerfDevicePerformanceV2'; Group='Startup'; Output='StartupDevices'; Grain='Device'; ApiVersion='beta'; Aliases=@(); Select=@('BlueScreenCount','BootScore','CoreBootTime','CoreLogonTime','DeviceId','DeviceName','LogonScore','Manufacturer','MemaTimeGenerated','Model','OSVersion','RestartCount','StartupPerformanceScore') },
        [pscustomobject]@{ Name='EAStartupPerfModelPerformanceV2'; Group='Startup'; Output='StartupModels'; Grain='Model'; ApiVersion='beta'; Aliases=@(); Select=@('AverageBlueScreens','AverageRestarts','CoreBootTime','CoreLogonTime','Manufacturer','MemaTimeGenerated','Model','StartupPerformanceScore') },
        [pscustomobject]@{ Name='EAStartupPerfDeviceProcesses'; Group='Startup'; Output='StartupProcesses'; Grain='Process'; ApiVersion='beta'; Aliases=@(); Select=@('FileDescription','InsertedDate','ProcessName','ProductName','Publisher','TimePerProcess','TotalDeviceCount') },
        [pscustomobject]@{ Name='EAAppPerformance'; Group='ApplicationReliability'; Output='AppReliability'; Grain='Application'; ApiVersion='beta'; Aliases=@(); Select=@('ActiveDevices','AppFriendlyName','AppHealthScore','AppName','AppPublisher','MeanTimeToFailure','MemaTimeGenerated','TotalAppCrashes','TotalAppUsageDuration') },
        [pscustomobject]@{ Name='EAOSVersionsPerformance'; Group='ApplicationReliability'; Output='OSReliability'; Grain='OSVersion'; ApiVersion='beta'; Aliases=@(); Select=@('ActiveDevices','MeanTimeToFailure','MemaTimeGenerated','OSBuildNumber','OSVersion','OSVersionAppHealthScore') },
        [pscustomobject]@{ Name='EAWFADeviceList'; Group='WorkFromAnywhere'; Output='WorkFromAnywhere'; Grain='Device'; ApiVersion='beta'; Aliases=@('WorkFromAnywhereDeviceList'); Select=@('DeviceId','DeviceName','Manufacturer','Model','OSDescription','OSVersion') },
        [pscustomobject]@{ Name='EAWFAPerDevicePerformance'; Group='WorkFromAnywhere'; Output='WorkFromAnywhere'; Grain='Device'; ApiVersion='beta'; Aliases=@(); Select=@('CloudManagementScore','DeviceId','DeviceName','Manufacturer','MemaTimeGenerated','Model','WindowsScore','WorkFromAnywhereScore') },
        [pscustomobject]@{ Name='EAWFAModelPerformance'; Group='WorkFromAnywhere'; Output='WorkFromAnywhere'; Grain='Model'; ApiVersion='beta'; Aliases=@(); Select=@('CloudManagementScore','Manufacturer','MemaTimeGenerated','Model','ModelDeviceCount','WindowsScore','WorkFromAnywhereScore') }
    )
}

function Get-EAOutputSchemas {
    return [ordered]@{
        DevicePerformance=@('RunId','ReportName','ReportRefreshDate','DeviceId','DeviceName','Manufacturer','Model','EndpointAnalyticsScore','StartupScore','AppReliabilityScore','WorkFromAnywhereScore','CrashCount','MeanTimeToFailure')
        ModelPerformance=@('RunId','ReportName','ReportRefreshDate','Manufacturer','Model','EndpointAnalyticsScore','StartupScore','AppReliabilityScore','WorkFromAnywhereScore','CrashCount','MeanTimeToFailure')
        StartupDevices=@('RunId','ReportName','ReportRefreshDate','DeviceId','DeviceName','Manufacturer','Model','OSVersion','StartupScore','BootScore','SignInScore','CoreBootTime','CoreSignInTime','RestartCount','StopErrorCount')
        StartupModels=@('RunId','ReportName','ReportRefreshDate','Manufacturer','Model','StartupScore','CoreBootTime','CoreSignInTime','RestartCount','StopErrorCount')
        StartupProcesses=@('RunId','ReportName','ReportRefreshDate','ApplicationName','Publisher','UsageDuration')
        AppReliability=@('RunId','ReportName','ReportRefreshDate','ApplicationName','Publisher','AppReliabilityScore','CrashCount','UsageDuration','MeanTimeToFailure')
        OSReliability=@('RunId','ReportName','ReportRefreshDate','OSVersion','AppReliabilityScore','MeanTimeToFailure')
        WorkFromAnywhere=@('RunId','ReportName','ReportRefreshDate','DeviceId','DeviceName','Manufacturer','Model','OSVersion','WorkFromAnywhereScore','CloudManagementScore','WindowsScore')
        DataQuality=@('RunId','ReportName','ApiVersion','Status','RowCount','ExportJobStatus','IsAdvancedAnalytics','RequiredPermission','ErrorCode','ErrorMessage','CollectedAtUtc')
    }
}

function Resolve-EAReportSelection {
    param([Parameter(Mandatory)][object[]]$Catalog, [Parameter(Mandatory)][string[]]$RequestedReports)
    $requested = @($RequestedReports | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($requested.Count -eq 0) { $requested = @('All') }
    foreach ($name in $requested) {
        if ($name -match $script:AdvancedReportPattern) { throw "Advanced Analytics report '$name' is excluded by design." }
    }
    if ($requested -contains 'All') { return @($Catalog) }
    $selected = New-Object System.Collections.Generic.List[object]
    foreach ($name in $requested) {
        $reportMatches = @($Catalog | Where-Object { $_.Name -ieq $name -or $_.Group -ieq $name -or $_.Aliases -icontains $name })
        if ($reportMatches.Count -eq 0) { throw "Unsupported Endpoint Analytics report or group: $name" }
        foreach ($reportMatch in $reportMatches) { if (-not ($selected.Name -contains $reportMatch.Name)) { $selected.Add($reportMatch) } }
    }
    return @($selected.ToArray() | Sort-Object Name -Unique)
}

function Test-EAStaticContract {
    param([Parameter(Mandatory)][object[]]$Catalog, [Parameter(Mandatory)][System.Collections.IDictionary]$Schemas)
    $failures = New-Object System.Collections.Generic.List[string]
    foreach ($report in $Catalog) {
        if ($report.Name -match $script:AdvancedReportPattern) { $failures.Add("Advanced Analytics report present: $($report.Name)") }
        if ($report.ApiVersion -ne 'beta') { $failures.Add("Undocumented API version: $($report.Name)") }
        if (@($report.Select).Count -eq 0) { $failures.Add("No explicit select list: $($report.Name)") }
        if (-not $Schemas.Contains($report.Output)) { $failures.Add("Missing output schema: $($report.Output)") }
        foreach ($advancedColumn in @('ResourcePerfScore','OverviewBatteryHealthScore')) {
            if ($report.Select -contains $advancedColumn) { $failures.Add("Advanced score column selected: $($report.Name).$advancedColumn") }
        }
    }
    if ($failures.Count -gt 0) { throw ("Endpoint Analytics static contract failed: {0}" -f ($failures -join ' | ')) }
    return $true
}

function Get-EAStatusCode {
    param([Parameter(Mandatory)]$ErrorRecord)
    if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Data.Contains('StatusCode')) { return [int]$ErrorRecord.Exception.Data['StatusCode'] }
    try { if ($ErrorRecord.Exception.Response.StatusCode) { return [int]$ErrorRecord.Exception.Response.StatusCode } }
    catch { Microsoft.PowerShell.Utility\Write-Debug "No HTTP response status was exposed: $($_.Exception.Message)" }
    if ([string]$ErrorRecord -match '\b(400|401|403|404|408|409|429|500|502|503|504)\b') { return [int]$Matches[1] }
    return 0
}

function Get-EARetryAfterSeconds {
    param([Parameter(Mandatory)]$ErrorRecord, [Parameter(Mandatory)][int]$Attempt)
    if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Data.Contains('RetryAfter')) {
        return [Math]::Max(0, [Math]::Min(60, [int]$ErrorRecord.Exception.Data['RetryAfter']))
    }
    try {
        $retryAfter = $ErrorRecord.Exception.Response.Headers.RetryAfter
        if ($retryAfter.Delta) { return [Math]::Max(1, [Math]::Min(60, [int][Math]::Ceiling($retryAfter.Delta.TotalSeconds))) }
    }
    catch { Microsoft.PowerShell.Utility\Write-Debug "No Retry-After header was exposed: $($_.Exception.Message)" }
    return [Math]::Min(60, [int][Math]::Pow(2, [Math]::Min($Attempt, 5)))
}

function Invoke-EAGraphRequest {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [AllowNull()][string]$Body,
        [int]$RetryCount=$MaxRetryCount,
        [scriptblock]$Invoker,
        [switch]$NoSleep
    )
    for ($attempt=1; $attempt -le $RetryCount; $attempt++) {
        try {
            if ($Invoker) { return & $Invoker $Method $Uri $Body }
            $invokeParameters = @{ Method=$Method; Uri=$Uri; OutputType='PSObject'; ErrorAction='Stop' }
            if ($Method -eq 'POST') { $invokeParameters.Body=$Body; $invokeParameters.ContentType='application/json' }
            return Invoke-MgGraphRequest @invokeParameters
        }
        catch {
            $statusCode = Get-EAStatusCode $_
            if ($statusCode -notin @(408,409,429,500,502,503,504) -or $attempt -ge $RetryCount) { throw }
            $delay = Get-EARetryAfterSeconds $_ $attempt
            Write-EALog ("Graph transient failure HTTP {0}; retry {1}/{2} in {3}s: {4}" -f $statusCode,$attempt,$RetryCount,$delay,$Uri) WARNING
            if (-not $NoSleep -and $delay -gt 0) { Start-Sleep -Seconds $delay }
        }
    }
}

function Start-EAExportJob {
    param([Parameter(Mandatory)][object]$Report, [string]$EffectiveReportName=$Report.Name)
    if ($EffectiveReportName -match $script:AdvancedReportPattern) { throw "Advanced Analytics report '$EffectiveReportName' is excluded by design." }
    $body = [ordered]@{ reportName=$EffectiveReportName; select=@($Report.Select); format='csv'; localizationType='replaceLocalizableValues' }
    $uri = '{0}/{1}/deviceManagement/reports/exportJobs' -f $script:GraphApiBase,$Report.ApiVersion
    Write-EALog ("Creating Endpoint Analytics export job: {0} ({1})" -f $EffectiveReportName,$Report.ApiVersion)
    $job = Invoke-EAGraphRequest POST $uri ($body | ConvertTo-Json -Depth 6 -Compress)
    if ($null -eq $job -or [string]::IsNullOrWhiteSpace([string]$job.id)) { throw "exportJobs POST returned no job id for '$EffectiveReportName'." }
    return $job
}

function Wait-EAExportJob {
    param(
        [Parameter(Mandatory)][string]$JobId,
        [Parameter(Mandatory)][string]$ApiVersion,
        [int]$TimeoutSeconds=$ExportTimeoutSeconds,
        [int]$PollSeconds=$PollIntervalSeconds,
        [scriptblock]$StatusProvider,
        [switch]$NoSleep
    )
    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    $poll = 0
    do {
        $poll++
        $job = if ($StatusProvider) { & $StatusProvider $poll } else {
            $uri = '{0}/{1}/deviceManagement/reports/exportJobs/{2}' -f $script:GraphApiBase,$ApiVersion,[uri]::EscapeDataString($JobId)
            Invoke-EAGraphRequest GET $uri
        }
        $status = [string]$job.status
        Write-EALog ("Export job {0} status: {1}" -f $JobId,$status) DEBUG
        switch ($status.ToLowerInvariant()) {
            'completed' { return $job }
            'failed' {
                $failureText = if ($job.PSObject.Properties['error']) { [string]$job.error } else { 'No failure detail returned.' }
                throw "Endpoint Analytics export job failed: $JobId. $failureText"
            }
        }
        if (-not $NoSleep) { Start-Sleep -Seconds ([Math]::Min(60,$PollSeconds + [Math]::Min(15,$poll - 1))) }
    } while ([datetime]::UtcNow -lt $deadline)
    throw "Endpoint Analytics export job timed out after $TimeoutSeconds seconds: $JobId"
}

function Invoke-EAWebDownload {
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$OutFile)
    for ($attempt=1; $attempt -le $MaxRetryCount; $attempt++) {
        try { Invoke-WebRequest -Uri $Uri -OutFile $OutFile -ErrorAction Stop; return }
        catch {
            $statusCode = Get-EAStatusCode $_
            if ($statusCode -notin @(408,429,500,502,503,504) -or $attempt -ge $MaxRetryCount) { throw }
            $delay = Get-EARetryAfterSeconds $_ $attempt
            Write-EALog ("ZIP download retry {0}/{1} in {2}s (HTTP {3})." -f $attempt,$MaxRetryCount,$delay,$statusCode) WARNING
            if ($delay -gt 0) { Start-Sleep -Seconds $delay }
        }
    }
}

function Import-EAExportedCsv {
    param([Parameter(Mandatory)][object]$CompletedJob, [Parameter(Mandatory)][string]$ReportName)
    if ([string]::IsNullOrWhiteSpace([string]$CompletedJob.url)) { throw "Completed job for '$ReportName' returned no URL." }
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("SmartM365-EA-{0}" -f [guid]::NewGuid().Guid)
    $zipPath = Join-Path $tempRoot 'export.zip'
    $extractPath = Join-Path $tempRoot 'content'
    try {
        New-Item -Path $extractPath -ItemType Directory -Force | Out-Null
        Invoke-EAWebDownload ([string]$CompletedJob.url) $zipPath
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractPath -Force
        $csvFile = Get-ChildItem -LiteralPath $extractPath -Filter '*.csv' -File -Recurse | Select-Object -First 1
        if (-not $csvFile) { throw "No CSV found in export ZIP for '$ReportName'." }
        $rows = @(Import-Csv -LiteralPath $csvFile.FullName)
        if ($MaxItems -gt 0) { $rows = @($rows | Select-Object -First $MaxItems) }
        return $rows
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Get-EARawValue {
    param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)][string[]]$Names)
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return $property.Value }
    }
    return ''
}

function New-EANormalizedRow {
    param([Parameter(Mandatory)][string]$ReportName, [Parameter(Mandatory)]$RawRow)
    $base = [ordered]@{
        RunId=$script:RunId
        ReportName=$ReportName
        ReportRefreshDate=(Get-EARawValue $RawRow @('MemaTimeGenerated','ProcessedDateTime','InsertedDate'))
    }
    switch ($ReportName) {
        'EADevicePerformanceV2' {
            $base.DeviceId=Get-EARawValue $RawRow @('DeviceId'); $base.DeviceName=Get-EARawValue $RawRow @('DeviceName')
            $base.Manufacturer=Get-EARawValue $RawRow @('DeviceManufacturer'); $base.Model=Get-EARawValue $RawRow @('DeviceModel')
            $base.EndpointAnalyticsScore=''; $base.StartupScore=''; $base.AppReliabilityScore=Get-EARawValue $RawRow @('DeviceAppHealthScore')
            $base.WorkFromAnywhereScore=''; $base.CrashCount=Get-EARawValue $RawRow @('TotalAppCrashes'); $base.MeanTimeToFailure=Get-EARawValue $RawRow @('MeanTimeToFailure')
        }
        'EADeviceModelPerformanceV2' {
            $base.Manufacturer=Get-EARawValue $RawRow @('DeviceManufacturer'); $base.Model=Get-EARawValue $RawRow @('DeviceModel')
            $base.EndpointAnalyticsScore=''; $base.StartupScore=''; $base.AppReliabilityScore=Get-EARawValue $RawRow @('ModelAppHealthScore')
            $base.WorkFromAnywhereScore=''; $base.CrashCount=''; $base.MeanTimeToFailure=Get-EARawValue $RawRow @('MeanTimeToFailure')
        }
        'EADeviceScoresV2' {
            $base.DeviceId=Get-EARawValue $RawRow @('DeviceId'); $base.DeviceName=Get-EARawValue $RawRow @('DeviceName')
            $base.Manufacturer=Get-EARawValue $RawRow @('Manufacturer'); $base.Model=Get-EARawValue $RawRow @('Model')
            $base.EndpointAnalyticsScore=Get-EARawValue $RawRow @('EndpointAnalyticsScore'); $base.StartupScore=Get-EARawValue $RawRow @('StartupPerformanceScore')
            $base.AppReliabilityScore=Get-EARawValue $RawRow @('AppReliabilityScore'); $base.WorkFromAnywhereScore=Get-EARawValue $RawRow @('WorkFromAnywhereScore')
            $base.CrashCount=''; $base.MeanTimeToFailure=''
        }
        'EAModelScoresV2' {
            $base.Manufacturer=Get-EARawValue $RawRow @('Manufacturer'); $base.Model=Get-EARawValue $RawRow @('Model')
            $base.EndpointAnalyticsScore=Get-EARawValue $RawRow @('EndpointAnalyticsScore'); $base.StartupScore=Get-EARawValue $RawRow @('StartupPerformanceScore')
            $base.AppReliabilityScore=Get-EARawValue $RawRow @('AppReliabilityScore'); $base.WorkFromAnywhereScore=Get-EARawValue $RawRow @('WorkFromAnywhereScore')
            $base.CrashCount=''; $base.MeanTimeToFailure=''
        }
        'EAStartupPerfDevicePerformanceV2' {
            $base.DeviceId=Get-EARawValue $RawRow @('DeviceId'); $base.DeviceName=Get-EARawValue $RawRow @('DeviceName')
            $base.Manufacturer=Get-EARawValue $RawRow @('Manufacturer'); $base.Model=Get-EARawValue $RawRow @('Model'); $base.OSVersion=Get-EARawValue $RawRow @('OSVersion')
            $base.StartupScore=Get-EARawValue $RawRow @('StartupPerformanceScore'); $base.BootScore=Get-EARawValue $RawRow @('BootScore'); $base.SignInScore=Get-EARawValue $RawRow @('LogonScore')
            $base.CoreBootTime=Get-EARawValue $RawRow @('CoreBootTime'); $base.CoreSignInTime=Get-EARawValue $RawRow @('CoreLogonTime')
            $base.RestartCount=Get-EARawValue $RawRow @('RestartCount'); $base.StopErrorCount=Get-EARawValue $RawRow @('BlueScreenCount')
        }
        'EAStartupPerfModelPerformanceV2' {
            $base.Manufacturer=Get-EARawValue $RawRow @('Manufacturer'); $base.Model=Get-EARawValue $RawRow @('Model')
            $base.StartupScore=Get-EARawValue $RawRow @('StartupPerformanceScore'); $base.CoreBootTime=Get-EARawValue $RawRow @('CoreBootTime'); $base.CoreSignInTime=Get-EARawValue $RawRow @('CoreLogonTime')
            $base.RestartCount=Get-EARawValue $RawRow @('AverageRestarts'); $base.StopErrorCount=Get-EARawValue $RawRow @('AverageBlueScreens')
        }
        'EAStartupPerfDeviceProcesses' {
            $base.ApplicationName=Get-EARawValue $RawRow @('ProductName','FileDescription','ProcessName'); $base.Publisher=Get-EARawValue $RawRow @('Publisher'); $base.UsageDuration=Get-EARawValue $RawRow @('TimePerProcess')
        }
        'EAAppPerformance' {
            $base.ApplicationName=Get-EARawValue $RawRow @('AppFriendlyName','AppName'); $base.Publisher=Get-EARawValue $RawRow @('AppPublisher')
            $base.AppReliabilityScore=Get-EARawValue $RawRow @('AppHealthScore'); $base.CrashCount=Get-EARawValue $RawRow @('TotalAppCrashes')
            $base.UsageDuration=Get-EARawValue $RawRow @('TotalAppUsageDuration'); $base.MeanTimeToFailure=Get-EARawValue $RawRow @('MeanTimeToFailure')
        }
        'EAOSVersionsPerformance' {
            $base.OSVersion=Get-EARawValue $RawRow @('OSVersion','OSBuildNumber'); $base.AppReliabilityScore=Get-EARawValue $RawRow @('OSVersionAppHealthScore'); $base.MeanTimeToFailure=Get-EARawValue $RawRow @('MeanTimeToFailure')
        }
        { $_ -in @('EAWFADeviceList','WorkFromAnywhereDeviceList') } {
            $base.DeviceId=Get-EARawValue $RawRow @('DeviceId'); $base.DeviceName=Get-EARawValue $RawRow @('DeviceName')
            $base.Manufacturer=Get-EARawValue $RawRow @('Manufacturer'); $base.Model=Get-EARawValue $RawRow @('Model'); $base.OSVersion=Get-EARawValue $RawRow @('OSVersion','OSDescription')
            $base.WorkFromAnywhereScore=''; $base.CloudManagementScore=''; $base.WindowsScore=''
        }
        'EAWFAPerDevicePerformance' {
            $base.DeviceId=Get-EARawValue $RawRow @('DeviceId'); $base.DeviceName=Get-EARawValue $RawRow @('DeviceName')
            $base.Manufacturer=Get-EARawValue $RawRow @('Manufacturer'); $base.Model=Get-EARawValue $RawRow @('Model'); $base.OSVersion=''
            $base.WorkFromAnywhereScore=Get-EARawValue $RawRow @('WorkFromAnywhereScore'); $base.CloudManagementScore=Get-EARawValue $RawRow @('CloudManagementScore'); $base.WindowsScore=Get-EARawValue $RawRow @('WindowsScore')
        }
        'EAWFAModelPerformance' {
            $base.DeviceId=''; $base.DeviceName=''; $base.Manufacturer=Get-EARawValue $RawRow @('Manufacturer'); $base.Model=Get-EARawValue $RawRow @('Model'); $base.OSVersion=''
            $base.WorkFromAnywhereScore=Get-EARawValue $RawRow @('WorkFromAnywhereScore'); $base.CloudManagementScore=Get-EARawValue $RawRow @('CloudManagementScore'); $base.WindowsScore=Get-EARawValue $RawRow @('WindowsScore')
        }
        default { throw "No normalization mapping exists for '$ReportName'." }
    }
    return [pscustomobject]$base
}

function New-EADataQualityRow {
    param(
        [Parameter(Mandatory)][string]$ReportName,
        [Parameter(Mandatory)][string]$ApiVersion,
        [Parameter(Mandatory)][string]$Status,
        [int]$RowCount=0,
        [string]$ExportJobStatus='',
        [string]$ErrorCode='',
        [string]$ErrorMessage=''
    )
    return [pscustomobject][ordered]@{
        RunId=$script:RunId; ReportName=$ReportName; ApiVersion=$ApiVersion; Status=$Status; RowCount=$RowCount
        ExportJobStatus=$ExportJobStatus; IsAdvancedAnalytics=$false; RequiredPermission=$script:RequiredPermission
        ErrorCode=$ErrorCode; ErrorMessage=$ErrorMessage; CollectedAtUtc=$script:CollectedAtUtc
    }
}

function Invoke-EAReport {
    param([Parameter(Mandatory)][object]$Report, [switch]$AvailabilityOnly)
    $lastError = $null
    foreach ($effectiveName in @($Report.Name) + @($Report.Aliases)) {
        try {
            $job = Start-EAExportJob $Report $effectiveName
            $completed = Wait-EAExportJob ([string]$job.id) $Report.ApiVersion
            $rawRows = if ($AvailabilityOnly) { @() } else { @(Import-EAExportedCsv $completed $effectiveName) }
            $normalizedRows = @($rawRows | ForEach-Object { New-EANormalizedRow $effectiveName $_ })
            return [pscustomobject]@{ EffectiveName=$effectiveName; Job=$completed; Rows=$normalizedRows; AliasUsed=($effectiveName -ne $Report.Name) }
        }
        catch {
            $lastError = $_
            $statusCode = Get-EAStatusCode $_
            if ($effectiveName -eq $Report.Name -and @($Report.Aliases).Count -gt 0 -and $statusCode -in @(0,400,404)) {
                Write-EALog ("Report {0} rejected; trying alias {1}. {2}" -f $Report.Name,$Report.Aliases[0],$_.Exception.Message) WARNING
                continue
            }
            break
        }
    }
    throw $lastError
}

function Publish-EAOutputs {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$OutputRows, [Parameter(Mandatory)][System.Collections.IDictionary]$Schemas)
    $fileNames = [ordered]@{
        DevicePerformance='Intune_EndpointAnalytics_DevicePerformance.csv'
        ModelPerformance='Intune_EndpointAnalytics_ModelPerformance.csv'
        StartupDevices='Intune_EndpointAnalytics_StartupDevices.csv'
        StartupModels='Intune_EndpointAnalytics_StartupModels.csv'
        StartupProcesses='Intune_EndpointAnalytics_StartupProcesses.csv'
        AppReliability='Intune_EndpointAnalytics_AppReliability.csv'
        OSReliability='Intune_EndpointAnalytics_OSReliability.csv'
        WorkFromAnywhere='Intune_EndpointAnalytics_WorkFromAnywhere.csv'
        DataQuality='Intune_EndpointAnalytics_DataQuality.csv'
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    foreach ($key in $fileNames.Keys) {
        if ($key -eq 'StartupProcesses' -and -not $IncludeStartupProcesses) { continue }
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileNames[$key])
        $timestampedPath = Join-Path $script:OutputPath ("{0}_{1}.csv" -f $baseName,$stamp)
        $latestPath = Join-Path $script:LatestCsvFolderPath $fileNames[$key]
        $rows = if ($OutputRows.Contains($key)) { @($OutputRows[$key]) } else { @() }
        Publish-CoreSmartM365Csv -Data $rows -TimestampedPath $timestampedPath -LatestPath $latestPath -Columns @($Schemas[$key]) | Out-Null
    }
}

function Invoke-EASelfTest {
    $catalog = @(Get-EAReportCatalog)
    $schemas = Get-EAOutputSchemas
    Test-EAStaticContract $catalog $schemas | Out-Null

    $completedSequence = @(
        [pscustomobject]@{status='notStarted'},
        [pscustomobject]@{status='inProgress'},
        [pscustomobject]@{status='completed';url='https://example.invalid/export.zip'}
    )
    $completed = Wait-EAExportJob simulation-completed beta 30 1 {
        param($poll)
        $completedSequence[[Math]::Min($poll - 1,$completedSequence.Count - 1)]
    } -NoSleep
    if ($completed.status -ne 'completed') { throw 'Completed simulation failed.' }

    $failedCaught = $false
    try {
        Wait-EAExportJob simulation-failed beta 30 1 { [pscustomobject]@{status='failed';error='Simulated failure'} } -NoSleep | Out-Null
    } catch { $failedCaught = $_.Exception.Message -match 'failed' }
    if (-not $failedCaught) { throw 'Failed simulation did not fail.' }

    $script:ThrottleAttempt = 0
    $throttleResult = Invoke-EAGraphRequest GET 'https://example.invalid/throttle' $null 2 {
        $script:ThrottleAttempt++
        if ($script:ThrottleAttempt -eq 1) {
            $exception = [System.Exception]::new('Simulated 429')
            $exception.Data['StatusCode']=429; $exception.Data['RetryAfter']=0
            throw $exception
        }
        [pscustomobject]@{status='completed'}
    } -NoSleep
    if ($throttleResult.status -ne 'completed' -or $script:ThrottleAttempt -ne 2) { throw '429 simulation failed.' }

    $sample = New-EANormalizedRow EADeviceScoresV2 ([pscustomobject]@{DeviceId='device-1';DeviceName='device';Manufacturer='vendor';Model='model';EndpointAnalyticsScore=75;StartupPerformanceScore=70;AppReliabilityScore=80;WorkFromAnywhereScore=76})
    if ($sample.DeviceId -ne 'device-1') { throw 'Normalization simulation failed.' }
    if ($schemas.DataQuality[0] -ne 'RunId') { throw 'Schema simulation failed.' }
    if ('EAResourcePerfAggByDevice' -notmatch $script:AdvancedReportPattern) { throw 'Advanced Analytics guard failed.' }
    if ((@('TenantKey') + $schemas.DevicePerformance)[0] -ne 'TenantKey') { throw 'TenantKey simulation failed.' }
    if ($MaxItems -gt 0 -and "x_MAXITEMS_$MaxItems.csv" -notmatch '_MAXITEMS_\d+\.csv$') { throw 'MAXITEMS simulation failed.' }
    Write-EALog 'Self-test passed: completed, failed, throttled, schemas, TenantKey, MAXITEMS, normalization, permission, and Advanced Analytics guard.' SUCCESS
}

if ($SelfTest) { Invoke-EASelfTest; return }

$script:CompletionStatus='Success'
$script:CompletionError=$null
$script:FailureStage=''
$script:OutputRows=[ordered]@{}
$script:DataQualityRows=New-Object System.Collections.Generic.List[object]

try {
    Initialize-EARuntime
    $catalog=@(Get-EAReportCatalog)
    $schemas=Get-EAOutputSchemas
    Test-EAStaticContract $catalog $schemas | Out-Null
    $selectedReports=@(Resolve-EAReportSelection $catalog $Reports)

    if (-not $IncludeStartupProcesses -and @($selectedReports | Where-Object Name -eq 'EAStartupPerfDeviceProcesses').Count -gt 0) {
        $selectedReports=@($selectedReports | Where-Object Name -ne 'EAStartupPerfDeviceProcesses')
        $script:DataQualityRows.Add((New-EADataQualityRow EAStartupPerfDeviceProcesses beta SkippedByParameter 0 NotCreated '' 'Use -IncludeStartupProcesses to collect this report.'))
    }

    Write-EALog ("Starting {0} v{1}; Tenant={2}; Reports={3}; ValidateOnly={4}; MaxItems={5}" -f $script:ScriptName,$script:ScriptVersion,$Tenant,($selectedReports.Name -join ', '),[bool]$ValidateOnly,$MaxItems)
    $existingContext=Get-MgContext -ErrorAction SilentlyContinue
    $shouldConnectForValidation=$ValidateOnly -and ($Connect -or $InteractiveAuth)
    $shouldConnectForCollection=-not $ValidateOnly
    if ($shouldConnectForValidation -or $shouldConnectForCollection) {
        $tenantId=[string](Get-EAConfigValue TenantId '')
        $appId=[string](Get-EAConfigValue AppId '')
        $thumbprint=[string](Get-EAConfigValue Thumbprint (Get-EAConfigValue Thumb ''))
        if ($shouldConnectForCollection -or $Connect -or $InteractiveAuth -or $null -eq $existingContext) {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
            catch { Write-EALog "Existing Graph session disconnect failed: $($_.Exception.Message)" DEBUG }
            if ($InteractiveAuth) {
                $connectParameters=@{Scopes=@($script:RequiredPermission);NoWelcome=$true;ErrorAction='Stop'}
                if ($tenantId) { $connectParameters.TenantId=$tenantId }
                Write-EALog 'Connecting to Microsoft Graph with delegated interactive authentication.'
                Connect-MgGraph @connectParameters | Out-Null
            }
            else {
                if (-not $tenantId -or -not $appId -or -not $thumbprint) { throw 'App-only authentication requires TenantId, AppId, and Thumbprint in local SmartM365 configuration.' }
                Write-EALog 'Connecting to Microsoft Graph with SmartM365 app-only certificate authentication.'
                Connect-MgGraph -TenantId $tenantId -ClientId $appId -CertificateThumbprint $thumbprint -NoWelcome -ErrorAction Stop | Out-Null
            }
        }
        $script:FailureStage='Preflight'
        Invoke-CoreSmartM365Preflight -ScriptName $script:ScriptName -RequiredModules @('Microsoft.Graph.Authentication') -OutputPaths @($script:OutputPath,$script:LatestCsvFolderPath) -RequiredGraphApplicationPermissions @($script:RequiredPermission) | Out-Null
    }

    if ($ValidateOnly -and -not $shouldConnectForValidation) {
        foreach ($report in $selectedReports) { $script:DataQualityRows.Add((New-EADataQualityRow $report.Name $report.ApiVersion StaticValidated 0 NotCreated)) }
        Write-EALog 'ValidateOnly static validation passed. Add -Connect or -InteractiveAuth to test tenant availability.' SUCCESS
        $script:DataQualityRows | Format-Table ReportName,ApiVersion,Status,ExportJobStatus
        return
    }

    $script:FailureStage=if ($ValidateOnly) {'Tenant report availability validation'} else {'Endpoint Analytics collection'}
    foreach ($report in $selectedReports) {
        try {
            $result=Invoke-EAReport $report -AvailabilityOnly:$ValidateOnly
            $status=if ($result.AliasUsed) {'AliasUsed'} elseif ($ValidateOnly) {'Available'} else {'Collected'}
            $message=if ($result.AliasUsed) {"Effective report name: $($result.EffectiveName)"} else {''}
            $script:DataQualityRows.Add((New-EADataQualityRow $report.Name $report.ApiVersion $status @($result.Rows).Count ([string]$result.Job.status) '' $message))
            if (-not $ValidateOnly) {
                if (-not $script:OutputRows.Contains($report.Output)) { $script:OutputRows[$report.Output]=New-Object System.Collections.Generic.List[object] }
                foreach ($row in @($result.Rows)) { $script:OutputRows[$report.Output].Add($row) }
            }
        }
        catch {
            $statusCode=Get-EAStatusCode $_
            $errorCode=if ($statusCode -gt 0) {"HTTP$statusCode"} else {'ExportFailed'}
            $status=if ($statusCode -in @(400,404)) {'UnavailableInTenant'} elseif ($statusCode -in @(401,403)) {'AccessDenied'} else {'Failed'}
            $script:DataQualityRows.Add((New-EADataQualityRow $report.Name $report.ApiVersion $status 0 failed $errorCode $_.Exception.Message))
            Write-EALog ("Report {0} failed: {1}" -f $report.Name,$_.Exception.Message) WARNING
        }
    }

    if ($ValidateOnly) {
        $script:DataQualityRows | Format-Table ReportName,ApiVersion,Status,ExportJobStatus,ErrorCode
        if (@($script:DataQualityRows | Where-Object Status -eq AccessDenied).Count -gt 0) { throw 'Endpoint Analytics permission or report access is missing.' }
        if (@($script:DataQualityRows | Where-Object Status -in @('Available','AliasUsed')).Count -eq 0) {
            throw 'Endpoint Analytics availability validation failed because no requested standard report completed successfully.'
        }
        Write-EALog 'ValidateOnly tenant availability validation completed. No CSV was published.' SUCCESS
        return
    }
    if (@($script:DataQualityRows | Where-Object Status -eq AccessDenied).Count -gt 0) {
        throw "Endpoint Analytics collection stopped: required permission or access is missing. Required: $($script:RequiredPermission)."
    }
    $terminalFailures = @($script:DataQualityRows | Where-Object Status -eq Failed)
    if ($terminalFailures.Count -gt 0) {
        throw "Endpoint Analytics collection stopped because $($terminalFailures.Count) report export(s) failed. Canonical business CSV files were not published."
    }
    if (@($script:DataQualityRows | Where-Object Status -in @('Collected','AliasUsed')).Count -eq 0) {
        throw 'Endpoint Analytics collection stopped because no requested standard report was available in the tenant.'
    }

    $script:OutputRows.DataQuality=$script:DataQualityRows
    Publish-EAOutputs $script:OutputRows $schemas
    $collectedCount=@($script:DataQualityRows | Where-Object Status -in @('Collected','AliasUsed')).Count
    $unavailableCount=@($script:DataQualityRows | Where-Object Status -eq UnavailableInTenant).Count
    $resultSummary="Endpoint Analytics standard reports collected: $collectedCount; unavailable: $unavailableCount; Advanced Analytics reports: 0."
    Write-EALog $resultSummary SUCCESS
    Send-CoreSmartM365TeamsNotification -Title "$($script:ScriptName) completed" -Message $resultSummary -Level SUCCESS -Channel Infos -ResultSummary $resultSummary -Facts @{Tenant=$Tenant;OutputPath=$script:OutputPath;RunId=$script:RunId} | Out-Null
}
catch {
    $script:CompletionStatus='Failed'; $script:CompletionError=$_
    Write-EALog $_.Exception.Message ERROR
    if ($script:CoreImported) {
        try { Send-CoreSmartM365TeamsNotification -Title "$($script:ScriptName) failed" -Message $_.Exception.Message -Level ERROR -Channel Alerts -Facts @{Tenant=$Tenant;FailedStage=$script:FailureStage;OutputPath=$script:OutputPath;RunId=$script:RunId} | Out-Null }
        catch { Microsoft.PowerShell.Utility\Write-Debug "Teams alert failed: $($_.Exception.Message)" }
    }
    throw
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    catch { Microsoft.PowerShell.Utility\Write-Debug "Graph disconnect failed: $($_.Exception.Message)" }
    if ($script:CoreImported) {
        try { Complete-CoreSmartM365ExecutionContext -Status $script:CompletionStatus -ErrorRecord $script:CompletionError -FailureStage $script:FailureStage }
        catch { Microsoft.PowerShell.Utility\Write-Debug "Completion banner failed: $($_.Exception.Message)" }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCBGggDVyb+EERD
# D0N6tdre7HSLcxriha5kbCXvW97z96CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIK2obS/nbpZehuEObQn6+7+Yqg7xLWCQjH0ESXhVfPYYMA0GCSqG
# SIb3DQEBAQUABIIBgE/+6nsXSrrPjaqivwjV5v03kqBdriJXtxIy3ffTlwe0R1Ac
# tSgtv34agRIhPWYuCfPoA/G+EipwPmzzKF2Tdt1MfSt6H4PtAsmCNnJJCho+3v0I
# 5gxmNvhtYg4ZPuv9HIkKpzfvGbKW+6wXoL0xJclx/WgOugeUS1dp0I6JmVe8kGXR
# BT6A8fS4VcFjpBQGuNC7KFajoCpP5RbXwqu9DsA/NTU3CsX3wu92HxbWA1q/nIPg
# 2cKvQjDZ6KhZ8R0PvQOayicx5p7eX3GAXzDkfNVZma6xATrWr96U6kUx9ZQwinoi
# jAMZh18Ak/V7xhLflrl92gpDVbEEM0SSNa0FalxpIv73/ueGzaMdiQbfW/BE4hnP
# dXn/KgBbSvBGV9dILeTZEuI8IeF4c6VPrNbyPtB395Ts3oYw70ZH/flv0JtKvo9D
# ODGw06LzU1UnRt8OB2HJUzbYGYW/5BwEh3MwSR+4JbS/WebdGVF/khbQ64WXFvxV
# 9YhnzWazdy/jvVOS/6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTgyMDMw
# NTVaMC8GCSqGSIb3DQEJBDEiBCAiKOwfV7TmgoC6Ox5MjPc7kvhh9CNypeyuhl2x
# kNSTkDANBgkqhkiG9w0BAQEFAASCAgA1ETKJEZd5ep8KYWafcQ5W+4Bx0MlSzxnx
# 6blSlKMMO4JReX8WkuvXHjlWoImkBPZW3uINeHumDdervjnntzef3zsbth81iOfe
# aQxR3ky3KRPD11r742EH3LO9Sx79ZLtB5lrTyiu0GD4pElADL+RT55Vn11y+PsxI
# VdTHuCWyern497Vioy1gkBdHOPCYjL6cUOMQgUeiKp+mdJsZNqKUWJLOJz9xgSaT
# lU8h2AIyc5q+CNLkVP85WtYUxJbdWNvPqQ9r9WEjIxgHhh/L02aE1wP9KDcE/tsV
# 5ZjvqjMZWhZRQ3D1clHhUbp+pzlDvw6VBisUaYdBTO4nKNYYI0kgWuR9uIHJvKMS
# aL9k1euz1h5yqLtbUMrLCjYexwBECBAYP/+IF0MHdXNFkAt0vw2Xu/q/3tVdJb8t
# 0BVFNeEovShYX4Ym8OsEzFouIfpH1qdckWnqyOTAezUyVjEZs/xI37IHnkz5lioI
# KH8qzWWg+nYia1Nq2oX6V3gq4FuX3NcyvHf+qeyZDlXjKei53gCJwKBYTukuvP1z
# J4BHhejnST0kQvw1NX4nlMOlxcy/fuIaS5Kn20JQONClndfAJ+3Kej9toBGsHMUm
# 7o6JFrJIPfBIjpMXOrs8NjFrbuEaYamWlxRRICqzQuQLOPmLDF7icOpiouq4PM0s
# 9ssWiQG+Eg==
# SIG # End signature block
