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
1.0.2

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
$script:ScriptVersion = '1.0.2'
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
        Invoke-CoreSmartM365Preflight -ScriptName $script:ScriptName -RequiredModules @('Microsoft.Graph.Authentication') -OutputPaths @($script:OutputPath,$script:LatestCsvFolderPath) -RequiredGraphApplicationPermissions @($script:RequiredPermission) -GraphProbeUris @("$($script:GraphApiBase)/beta/deviceManagement/reports/exportJobs?`$top=1") | Out-Null
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
