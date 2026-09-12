<#
.SYNOPSIS
Collects Windows update alert/status and Endpoint Analytics device scores.

.DESCRIPTION
Uses Microsoft Intune report export jobs without changing devices, policies,
assignments, baselines, or remediations. Creating the temporary export-job
resource currently requires DeviceManagementManagedDevices.ReadWrite.All.

.VERSION
1.0.2
#>
[CmdletBinding(DefaultParameterSetName = 'Graph')]
param(
    [Alias('ProfileKey')][string]$Tenant = 'default',
    [string]$OrganizationKey,
    [string]$EnvironmentKey,
    [string]$TenantKey,
    [string]$TenantId,
    [string]$DataRootPath,
    [string]$DataAllRootPath,
    [string]$LatestOutputRootPath,
    [string]$LogRootPath,
    [string]$GlobalConfigPath,
    [string]$TenantConfigPath,
    [Parameter(ParameterSetName = 'Fixture', Mandatory)][string]$InputJsonPath,
    [ValidateRange(0, 2147483647)][int]$MaxItems = 0,
    [ValidateRange(30, 1800)][int]$ExportTimeoutSeconds = 300,
    [ValidateRange(1, 60)][int]$PollSeconds = 5,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.0.2'
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Get-ConfigurationSection {
    param([System.Collections.IDictionary]$Configuration, [string]$Name)
    if ($Configuration.Contains($Name) -and $Configuration[$Name] -is [System.Collections.IDictionary]) {
        return $Configuration[$Name]
    }
    return [ordered]@{}
}

function Get-ConfigurationText {
    param([System.Collections.IDictionary]$Configuration, [string]$Name)
    if ($Configuration.Contains($Name)) { return ([string]$Configuration[$Name]).Trim() }
    return ''
}

function Get-GraphValue {
    param([AllowNull()]$InputObject, [string]$Name)
    return Get-SmartWorkplaceCMDBGraphObjectValue -InputObject $InputObject -Name $Name
}

function Get-CleanText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value -replace "`r`n|`n|`r", ' ').Trim()
}

function Get-PreferredText {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)][string]$Field)
    $values = @($Rows | ForEach-Object { [string]$_.$Field } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($values.Count -eq 0) { return '' }
    $ranked = @($values | Group-Object { $_.ToLowerInvariant() } | Sort-Object @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Ascending = $true })
    return [string]$ranked[0].Group[0]
}

function Get-MaxScoreText {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)][string]$Field)
    $scores = @($Rows | ForEach-Object { [string]$_.$Field } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { [double]::Parse($_, [Globalization.CultureInfo]::InvariantCulture) })
    if ($scores.Count -eq 0) { return '' }
    return (@($scores | Measure-Object -Maximum)[0].Maximum).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-DateText {
    param([AllowNull()]$Value, [string]$Field, [string]$Key)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    $date = [datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$date)) {
        throw "$Field '$Value' is invalid for '$Key'."
    }
    return $date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffffffZ', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ScoreText {
    param([AllowNull()]$Value, [string]$Field, [string]$Key)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    $score = 0.0
    if (-not [double]::TryParse([string]$Value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$score)) {
        throw "$Field '$Value' is outside 0..100 for '$Key'."
    }
    if ($score -eq -1) { return '' }
    if ($score -lt 0 -or $score -gt 100) {
        throw "$Field '$Value' is outside 0..100 for '$Key'."
    }
    return $score.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)
}

function Invoke-IntuneReportExport {
    param(
        [Parameter(Mandatory)][string]$ReportName,
        [Parameter(Mandatory)][string[]]$Select,
        [string]$Filter,
        [ValidateSet('v1.0', 'beta')][string]$ApiVersion = 'beta'
    )

    $body = [ordered]@{
        reportName = $ReportName
        select = @($Select)
        format = 'csv'
        localizationType = 'replaceLocalizableValues'
    }
    if (-not [string]::IsNullOrWhiteSpace($Filter)) { $body['filter'] = $Filter }
    $bodyJson = $body | ConvertTo-Json -Depth 6 -Compress
    $uri = "https://graph.microsoft.com/$ApiVersion/deviceManagement/reports/exportJobs"
    $post = {
        param($RequestUri)
        Invoke-MgGraphRequest -Method POST -Uri $RequestUri -Body $bodyJson -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
    }
    $job = Invoke-SmartWorkplaceCMDBGraphRequestWithRetry -Uri $uri -RequestScript $post
    $jobId = Get-CleanText (Get-GraphValue $job 'id')
    if ([string]::IsNullOrWhiteSpace($jobId)) { throw "Intune report '$ReportName' returned no export job ID." }

    $deadline = [datetime]::UtcNow.AddSeconds($ExportTimeoutSeconds)
    do {
        $statusUri = "$uri/$([uri]::EscapeDataString($jobId))"
        $completed = Invoke-SmartWorkplaceCMDBGraphRequestWithRetry -Uri $statusUri
        $status = (Get-CleanText (Get-GraphValue $completed 'status')).ToLowerInvariant()
        if ($status -eq 'completed') { break }
        if ($status -eq 'failed') { throw "Intune report '$ReportName' export job failed: $jobId." }
        Start-Sleep -Seconds $PollSeconds
    } while ([datetime]::UtcNow -lt $deadline)
    if ($status -ne 'completed') { throw "Intune report '$ReportName' export job timed out after $ExportTimeoutSeconds seconds." }

    $downloadUrl = Get-CleanText (Get-GraphValue $completed 'url')
    if ([string]::IsNullOrWhiteSpace($downloadUrl)) { throw "Intune report '$ReportName' completed without a download URL." }
    $temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ("SmartWorkplaceCMDB-IntuneReport-{0}" -f [guid]::NewGuid().ToString('N'))
    $zipPath = Join-Path $temporaryRoot 'report.zip'
    $extractPath = Join-Path $temporaryRoot 'content'
    try {
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -ErrorAction Stop
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractPath -Force
        $csv = Get-ChildItem -LiteralPath $extractPath -File -Filter '*.csv' | Select-Object -First 1
        if ($null -eq $csv) { throw "Intune report '$ReportName' archive contains no CSV." }
        return @(Import-Csv -LiteralPath $csv.FullName)
    }
    finally {
        if (Test-Path -LiteralPath $temporaryRoot) {
            Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptRoot)
$coreModule = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Core\SmartWorkplaceCMDB.Core.psd1'
$graphModule = Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.Graph\SmartWorkplaceCMDB.Graph.psd1'
$rawContractPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
Import-Module $coreModule -Force
Import-Module $graphModule -Force

$boundParameters = @{}
foreach ($key in $PSBoundParameters.Keys) { $boundParameters[$key] = $PSBoundParameters[$key] }
$context = Resolve-SmartWorkplaceCMDBContext -BoundParameters $boundParameters -GlobalConfigPath $GlobalConfigPath -TenantConfigPath $TenantConfigPath -NoConfigWrite:($ValidateOnly -or $NoConfigWrite -or $PSCmdlet.ParameterSetName -eq 'Fixture')
$paths = Resolve-SmartWorkplaceCMDBCollectionPaths -Paths $context.Paths -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems -ExplicitDataRoot:([bool]$DataRootPath) -NoWrite:$ValidateOnly
$contract = Get-SmartWorkplaceCMDBTableContract -Path $rawContractPath
$tableNames = @('Intune_WindowsUpdateAlerts.csv', 'Intune_EndpointAnalyticsDeviceScores.csv')
$tables = @{}
$latestPaths = @{}
foreach ($name in $tableNames) {
    $matches = @($contract.tables | Where-Object name -eq $name)
    if ($matches.Count -ne 1) { throw "Raw contract definition missing or duplicated: $name" }
    $tables[$name] = $matches[0]
    $latestPaths[$name] = [IO.Path]::GetFullPath((Join-Path $paths.LatestOutputRootPath (Join-Path ([string]$matches[0].area) $name)))
}

$mode = if ($ValidateOnly) { 'Validate' } elseif ($PSCmdlet.ParameterSetName -eq 'Fixture') { 'Fixture' } else { 'Collect' }
$runtime = Start-SmartWorkplaceCMDBExecutionContext -Context $context -ScriptPath $PSCommandPath -ScriptVersion $ScriptVersion -Mode $mode -NoWrite:$ValidateOnly
$executionError = $null
$connected = $false
try {
    $configuration = Get-ConfigurationSection $context.Configuration 'MicrosoftGraph'
    $clientId = Get-ConfigurationText $configuration 'ClientId'
    $thumbprint = Get-ConfigurationText $configuration 'CertificateThumbprint'
    $fixture = $null
    $readiness = $null
    if ($PSCmdlet.ParameterSetName -eq 'Fixture') {
        $InputJsonPath = [IO.Path]::GetFullPath($InputJsonPath)
        $fixture = Get-Content -LiteralPath $InputJsonPath -Raw | ConvertFrom-Json
    }
    else {
        $readiness = Test-SmartWorkplaceCMDBGraphAppOnlyReadiness -TenantId $paths.TenantId -ClientId $clientId -CertificateThumbprint $thumbprint
    }

    if ($ValidateOnly) {
        [pscustomobject]@{
            Status = 'Valid'
            ScriptVersion = $ScriptVersion
            SourceMode = if ($null -ne $fixture) { 'OfflineJson' } else { 'MicrosoftGraphExportJobs' }
            RawContractVersion = [string]$contract.contractVersion
            RequiredGraphPermissions = 'DeviceManagementConfiguration.Read.All;DeviceManagementManagedDevices.ReadWrite.All'
            OutputCount = $tableNames.Count
        } | Format-List
        return
    }

    $alertSourceRows = @()
    $analyticsSourceRows = @()
    if ($null -ne $fixture) {
        $alertSourceRows = @($fixture.windowsUpdateAlerts)
        $analyticsSourceRows = @($fixture.endpointAnalyticsDeviceScores)
    }
    else {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Connect-MgGraph -TenantId $readiness.TenantId -ClientId $readiness.ClientId -CertificateThumbprint $readiness.CertificateThumbprint -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
        $connected = $true
        $graphContext = Get-MgContext -ErrorAction Stop
        if ($null -eq $graphContext -or [string]$graphContext.TenantId -ne $readiness.TenantId) { throw 'Microsoft Graph connected to an unexpected tenant.' }

        $policyPath = Join-Path $paths.LatestOutputRootPath 'Raw\Intune\Intune_WindowsUpdatePolicies.csv'
        if (-not (Test-Path -LiteralPath $policyPath -PathType Leaf)) { throw "Windows update policy snapshot is required before report collection: '$policyPath'." }
        $policies = @(Import-SmartWorkplaceCMDBSourceCsv -LiteralPath $policyPath -Paths $paths)
        $reportSelect = @('DeviceId', 'DeviceName', 'PolicyId', 'EventDateTimeUTC', 'LastWUScanTimeUTC', 'AggregateState', 'CurrentDeviceUpdateStatus', 'LatestAlertMessage')
        foreach ($policy in $policies) {
            $policyId = Get-CleanText $policy.PolicyId
            if ([string]::IsNullOrWhiteSpace($policyId)) { continue }
            $reportName = if ([string]$policy.PolicyType -eq 'Feature') { 'FeatureUpdateDeviceState' } elseif ([string]$policy.PolicyType -eq 'Quality') { 'QualityUpdateDeviceStatusByPolicy' } else { '' }
            if ([string]::IsNullOrWhiteSpace($reportName)) { continue }
            $filter = "PolicyId eq '$($policyId.Replace("'", "''"))'"
            $reportRows = @(Invoke-IntuneReportExport -ReportName $reportName -Select $reportSelect -Filter $filter -ApiVersion v1.0)
            foreach ($row in $reportRows) {
                $row | Add-Member -NotePropertyName sourceReport -NotePropertyValue $reportName -Force
                $alertSourceRows += $row
            }
        }
        $analyticsSourceRows = @(Invoke-IntuneReportExport -ReportName 'EADeviceScoresV2' -Select @('AppReliabilityScore', 'DeviceId', 'DeviceName', 'EndpointAnalyticsScore', 'Manufacturer', 'Model', 'StartupPerformanceScore', 'WorkFromAnywhereScore') -ApiVersion beta)
    }

    if ($MaxItems -gt 0) {
        $alertSourceRows = @($alertSourceRows | Select-Object -First $MaxItems)
        $analyticsSourceRows = @($analyticsSourceRows | Select-Object -First $MaxItems)
    }
    $collected = [datetime]::UtcNow.ToString('o')
    $alertRows = @($alertSourceRows | ForEach-Object {
        $deviceId = Get-CleanText (Get-GraphValue $_ 'deviceId')
        $policyId = Get-CleanText (Get-GraphValue $_ 'policyId')
        $reportName = Get-CleanText (Get-GraphValue $_ 'sourceReport')
        if ([string]::IsNullOrWhiteSpace($deviceId) -or [string]::IsNullOrWhiteSpace($policyId) -or [string]::IsNullOrWhiteSpace($reportName)) { throw 'Windows update report row is missing sourceReport, deviceId, or policyId.' }
        [pscustomobject][ordered]@{
            SourceSystem = 'MicrosoftIntuneReports'
            SourceReport = $reportName
            DeviceId = $deviceId
            DeviceName = Get-CleanText (Get-GraphValue $_ 'deviceName')
            PolicyId = $policyId
            EventDateTimeUTC = Get-DateText (Get-GraphValue $_ 'eventDateTimeUTC') 'eventDateTimeUTC' $deviceId
            LastWUScanTimeUTC = Get-DateText (Get-GraphValue $_ 'lastWUScanTimeUTC') 'lastWUScanTimeUTC' $deviceId
            AggregateState = Get-CleanText (Get-GraphValue $_ 'aggregateState')
            CurrentDeviceUpdateStatus = Get-CleanText (Get-GraphValue $_ 'currentDeviceUpdateStatus')
            LatestAlertMessage = Get-CleanText (Get-GraphValue $_ 'latestAlertMessage')
            SourceCollectedDateTime = $collected
        }
    })
    $analyticsRows = @($analyticsSourceRows | ForEach-Object {
        $deviceId = Get-CleanText (Get-GraphValue $_ 'deviceId')
        if ([string]::IsNullOrWhiteSpace($deviceId)) { throw 'Endpoint Analytics report row is missing deviceId.' }
        [pscustomobject][ordered]@{
            SourceSystem = 'MicrosoftIntuneReports'
            DeviceId = $deviceId
            DeviceName = Get-CleanText (Get-GraphValue $_ 'deviceName')
            Manufacturer = Get-CleanText (Get-GraphValue $_ 'manufacturer')
            Model = Get-CleanText (Get-GraphValue $_ 'model')
            EndpointAnalyticsScore = Get-ScoreText (Get-GraphValue $_ 'endpointAnalyticsScore') 'endpointAnalyticsScore' $deviceId
            StartupPerformanceScore = Get-ScoreText (Get-GraphValue $_ 'startupPerformanceScore') 'startupPerformanceScore' $deviceId
            AppReliabilityScore = Get-ScoreText (Get-GraphValue $_ 'appReliabilityScore') 'appReliabilityScore' $deviceId
            WorkFromAnywhereScore = Get-ScoreText (Get-GraphValue $_ 'workFromAnywhereScore') 'workFromAnywhereScore' $deviceId
            SourceCollectedDateTime = $collected
        }
    })
    $analyticsGroups = @($analyticsRows | Group-Object DeviceId)
    $duplicateAnalyticsGroups = @($analyticsGroups | Where-Object Count -gt 1)
    if ($duplicateAnalyticsGroups.Count -gt 0) {
        $conflictingAnalyticsGroupCount = 0
        $collapsedAnalyticsRows = New-Object System.Collections.Generic.List[object]
        foreach ($group in @($analyticsGroups | Sort-Object Name)) {
            $hasConflict = $false
            foreach ($field in @('DeviceName', 'Manufacturer', 'Model', 'EndpointAnalyticsScore', 'StartupPerformanceScore', 'AppReliabilityScore', 'WorkFromAnywhereScore')) {
                $distinct = @($group.Group | ForEach-Object { ([string]$_.$field).ToLowerInvariant() } | Sort-Object -Unique)
                if ($distinct.Count -gt 1) { $hasConflict = $true }
            }
            if ($hasConflict) { $conflictingAnalyticsGroupCount++ }
            $collapsedAnalyticsRows.Add([pscustomobject][ordered]@{
                    SourceSystem = 'MicrosoftIntuneReports'
                    DeviceId = [string]$group.Name
                    DeviceName = Get-PreferredText -Rows @($group.Group) -Field 'DeviceName'
                    Manufacturer = Get-PreferredText -Rows @($group.Group) -Field 'Manufacturer'
                    Model = Get-PreferredText -Rows @($group.Group) -Field 'Model'
                    EndpointAnalyticsScore = Get-MaxScoreText -Rows @($group.Group) -Field 'EndpointAnalyticsScore'
                    StartupPerformanceScore = Get-MaxScoreText -Rows @($group.Group) -Field 'StartupPerformanceScore'
                    AppReliabilityScore = Get-MaxScoreText -Rows @($group.Group) -Field 'AppReliabilityScore'
                    WorkFromAnywhereScore = Get-MaxScoreText -Rows @($group.Group) -Field 'WorkFromAnywhereScore'
                    SourceCollectedDateTime = $collected
                })
        }
        Write-Warning ("Intune returned {0} duplicate Endpoint Analytics device key(s), including {1} with conflicting attributes. Canonical text values and maximum available scores were retained." -f $duplicateAnalyticsGroups.Count, $conflictingAnalyticsGroupCount)
        $analyticsRows = @($collapsedAnalyticsRows.ToArray())
    }
    $rowsByTable = @{
        'Intune_WindowsUpdateAlerts.csv' = $alertRows
        'Intune_EndpointAnalyticsDeviceScores.csv' = $analyticsRows
    }
    foreach ($name in $tableNames) {
        $key = if ($name -eq 'Intune_WindowsUpdateAlerts.csv') { { "$($_.SourceReport)|$($_.PolicyId)|$($_.DeviceId)|$($_.EventDateTimeUTC)" } } else { 'DeviceId' }
        $duplicates = @($rowsByTable[$name] | Group-Object $key | Where-Object Count -gt 1)
        if ($duplicates.Count) { throw "Duplicate Intune analytics keys returned for ${name}: $($duplicates.Name -join ', ')" }
    }

    $published = @()
    foreach ($name in $tableNames) {
        $run = $null
        try {
            $run = Start-SmartWorkplaceCMDBSourceCollection -Paths $paths -RawPath @($latestPaths[$name]) -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems
            $stamp = [datetime]::UtcNow
            $baseName = [IO.Path]::GetFileNameWithoutExtension($name)
            $history = Join-Path $paths.DataAllRootPath ('Intune\Analytics\{0}\{1}\{2}_{3}.csv' -f $stamp.ToString('yyyy'), $stamp.ToString('MM'), $baseName, $stamp.ToString('yyyyMMdd-HHmmssfff'))
            Publish-SmartWorkplaceCMDBSourceCsv -Run $run -InputObject @($rowsByTable[$name]) -Columns @($tables[$name].columns | ForEach-Object { [string]$_ }) -HistoryPath $history -LatestPath $latestPaths[$name] -ContractPath $rawContractPath -ContractTableName $name | Out-Null
            $published += $latestPaths[$name]
        }
        catch {
            if ($null -ne $run) { Complete-SmartWorkplaceCMDBSourceCollection -Run $run -Failed }
            throw
        }
    }
    Write-Information ("SmartWorkplaceCMDB Intune analytics collection completed. UpdateRows={0}; EndpointAnalyticsDevices={1}." -f $alertRows.Count, $analyticsRows.Count) -InformationAction Continue
    [pscustomobject]@{
        Status = 'Completed'
        ScriptVersion = $ScriptVersion
        WindowsUpdateRowCount = $alertRows.Count
        EndpointAnalyticsDeviceCount = $analyticsRows.Count
        PublishedPath = $published
    }
}
catch {
    $executionError = $_
    throw
}
finally {
    if ($connected -and (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue)) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    Complete-SmartWorkplaceCMDBExecutionContext -RuntimeContext $runtime -ErrorRecord $executionError
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCXAXcw8gFXeJGM
# fzXVyFcBnZndKiq8H4WpqEqovClBRaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEILtVBUqee69chkRfI2I+JQ8EFAESnYGwY+iQARRAWizUMA0GCSqG
# SIb3DQEBAQUABIIBgK+nbQqsZ211aOKnk5qdxX2v6KV7qdOh538gE1MjmO8Rh5qL
# +DreT9DOn4lY+2fqMYw4xOcK5L3Du1O3JZlwVDJOStyyUGRAPBetQ3dXfEjRPGej
# 1VV4twcozPt3fyWWG7BeAX0M6EmKCHCo7rXuDHViruKyeGkhw9P4PkwtCs3aHDcb
# htGCN+u4qRyOuyb0BuqpR5jTXgV06HDI2EV1uLDeH2bhb8M2TTeg7ROw1VWCvMA2
# wCz5sT/EyQH+9J3iOwG3LKnmH3nfxdbbOvBDC2IEK5Ly/hqtBeNjbxW+LaGBzdKo
# 7HtDjIVTTivd0wfMojJsLD5rQXS+NHynHASPV1I8yJT47IajzB5tPfRxSFcGkl0N
# m7SRQhNQCnDKAgcOHnhPFYsVe+qiAjWnWQDBSIRLIHq/FEmQ+jxLc383fIC01shf
# D+h/D1nalh/xNGL5m68JIGQRNIMBMKERcIJHhHwz096yaXCAUxEyNi6mVgWCYIiT
# XQ8Ulc6pGUyU4SukeqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIyMTMw
# MDVaMC8GCSqGSIb3DQEJBDEiBCA4Rc2SfaOFY0ohd67RdbbTrOFJk45yfyt7Z4DD
# eIZnBDANBgkqhkiG9w0BAQEFAASCAgCeyBOaKccCUsmViOzt2GPKTUZPtVc/mEq0
# Go0/OBTFKJqQrlZyuPEfNtojUciksfqo0h+x6ONL4m+VR8tW2Fo/dMIJZRFhnSgW
# hhFlGr7OVG6FUgq/NgoLe7vx/Lsvy7Lj5ijiy6hM7EvcWfRjR0R1iHOKRVg/txg/
# cHQizMr/MsRf5GR8lsppo8H/q/jCiJ9l3CqQQwlZLbzoMdh1VpiWga7m/67gaos8
# wZCWZ/C/aOLWm9u7xUqqn2r9wsa83RpcX8b6c72keQf0E5EYJBuIxarrM+uTSQbU
# 7moIXbhHVwSVnAjVWtNdsk3OyIIBkDC+leHmxEQdsff3KdMOlctjWUe/j2Z7urxS
# OUA7x+4DbpkZlDeBqZkR5DHUM0oQJiP02sDAijTz+qA1sNiBOqNxoxnSOTOEovgO
# g/bdU3Vwnnc3+HK6pbEDHw5jrdMPb/APWaTrvVeaEiA4xNuc0JWQtc7QldjZpksM
# u5zHR9k76nA63A32ep/xS8gYxMyVuPZJiVwAwE9mBlCGPsF+KWlVmWqYeUEOA5Mg
# JN+U7PvlklcQj0QqX4vX+7KbjbbsVrDUsfOKAXp2yIXHA5OdWcRqaMCbg/gA8zW3
# 0eJoXXdp70XFl7mHuAkWLHcVu8ZW5SpDm7ZnauGF+AjempffRJxIuYNZLJ0HoDe4
# dOw6vQ2wdw==
# SIG # End signature block
