<#
.SYNOPSIS
Collects Windows update status, Endpoint Analytics scores, and Windows upgrade eligibility.

.DESCRIPTION
Uses Microsoft Intune report export jobs and the read-only Work From Anywhere
device API without changing devices, policies, assignments, baselines, or
remediations. Creating the temporary export-job resource currently requires
DeviceManagementManagedDevices.ReadWrite.All.

.VERSION
1.0.5
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
    [ValidateRange(30, 1800)][int]$ExportTimeoutSeconds = 900,
    [ValidateRange(1, 60)][int]$PollSeconds = 5,
    [switch]$NoConfigWrite,
    [switch]$ValidateOnly
)

$ScriptVersion = '1.0.5'
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

function Get-UpdateAlertBaseKey {
    param([Parameter(Mandatory)]$Row)
    return @(
        (Get-CleanText $Row.SourceReport).ToLowerInvariant(),
        (Get-CleanText $Row.PolicyId).ToLowerInvariant(),
        (Get-CleanText $Row.DeviceId).ToLowerInvariant(),
        (Get-CleanText $Row.EventDateTimeUTC).ToLowerInvariant()
    ) -join '|'
}

function Get-UpdateAlertEvidenceFingerprint {
    param([Parameter(Mandatory)]$Row)
    $payload = @(
        (Get-CleanText $Row.LastWUScanTimeUTC).ToLowerInvariant(),
        (Get-CleanText $Row.AggregateState).ToLowerInvariant(),
        (Get-CleanText $Row.CurrentDeviceUpdateStatus).ToLowerInvariant(),
        (Get-CleanText $Row.LatestAlertMessage).ToLowerInvariant()
    ) -join [char]31
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
        return ([BitConverter]::ToString($algorithm.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $algorithm.Dispose()
    }
}

function Get-UpdateAlertEvidenceKey {
    param([Parameter(Mandatory)]$Row)
    return '{0}|{1}' -f (Get-UpdateAlertBaseKey $Row), (Get-UpdateAlertEvidenceFingerprint $Row)
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

function Get-UpgradeEligibilityText {
    param([AllowNull()]$Value, [string]$Key)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    switch (([string]$Value).Trim().ToLowerInvariant()) {
        { $_ -in @('0', 'upgraded') } { return 'upgraded' }
        { $_ -in @('1', 'unknown', 'undetermined', 'notapplicable') } { return 'unknown' }
        { $_ -in @('2', 'noteligible', 'notcapable', 'notready') } { return 'notCapable' }
        { $_ -in @('3', 'eligible', 'capable', 'ready') } { return 'capable' }
        { $_ -in @('4', 'unknownfuturevalue') } { return 'unknownFutureValue' }
        default { throw "upgradeEligibility '$Value' is not recognized for '$Key'." }
    }
}

function Invoke-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    $nextLink = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextLink)) {
        $pageUri = $nextLink
        $page = Invoke-IntuneExportRead -ReportName 'WorkFromAnywhere' -Operation 'metric page read' -RequestScript {
            Invoke-MgGraphRequest -Method GET -Uri $pageUri -ErrorAction Stop
        }
        foreach ($item in @(Get-GraphValue $page 'value')) {
            if ($null -ne $item) { $items.Add($item) }
        }
        $nextLink = Get-CleanText (Get-GraphValue $page '@odata.nextLink')
    }
    return @($items.ToArray())
}

function Get-UpgradeEligibilityGraphRows {
    $select = 'id,deviceId,deviceName,upgradeEligibility'
    $primaryError = ''
    $primaryStatus = 0
    try {
        $metricsUri = "https://graph.microsoft.com/v1.0/deviceManagement/userExperienceAnalyticsWorkFromAnywhereMetrics?`$select=id"
        $metrics = @(Invoke-GraphCollection -Uri $metricsUri)
        $metricIds = @($metrics | ForEach-Object { Get-CleanText (Get-GraphValue $_ 'id') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
        if ($metricIds.Count -eq 0) { throw 'Microsoft Graph returned no Work From Anywhere metric identifier.' }
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($metricId in $metricIds) {
            $escapedMetricId = [uri]::EscapeDataString($metricId)
            $deviceUri = "https://graph.microsoft.com/v1.0/deviceManagement/userExperienceAnalyticsWorkFromAnywhereMetrics/$escapedMetricId/metricDevices?`$select=$select"
            foreach ($row in @(Invoke-GraphCollection -Uri $deviceUri)) {
                $row | Add-Member -NotePropertyName sourceMetricId -NotePropertyValue $metricId -Force
                $rows.Add($row)
            }
        }
        return [pscustomobject]@{ Status = 'CollectedV1'; Rows = @($rows.ToArray()); Error = '' }
    }
    catch {
        $primaryError = $_.Exception.Message
        $primaryStatus = (Get-IntuneExportErrorDetails -ErrorRecord $_).StatusCode
        Write-Warning ("The documented Work From Anywhere metric route was unavailable; trying the SmartInventory allDevices beta route. Error: {0}" -f $primaryError)
    }

    try {
        $fallbackUri = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsWorkFromAnywhereMetrics('allDevices')/metricDevices?`$select=$select"
        $rows = @(Invoke-GraphCollection -Uri $fallbackUri)
        foreach ($row in $rows) {
            $row | Add-Member -NotePropertyName sourceMetricId -NotePropertyValue 'allDevices' -Force
        }
        return [pscustomobject]@{ Status = 'CollectedBetaAllDevices'; Rows = $rows; Error = '' }
    }
    catch {
        $fallbackError = $_.Exception.Message
        $fallbackStatus = (Get-IntuneExportErrorDetails -ErrorRecord $_).StatusCode
        $message = "Work From Anywhere device readiness is unavailable. V1=$primaryError; BetaAllDevices=$fallbackError"
        if ($primaryStatus -in @(408, 429, 500, 502, 503, 504) -or
            $fallbackStatus -in @(408, 429, 500, 502, 503, 504)) {
            throw "$message. A transient Graph failure must not be published as an empty snapshot."
        }
        Write-Warning $message
        return [pscustomobject]@{ Status = 'Unavailable'; Rows = @(); Error = $message }
    }
}

function Get-IntuneExportErrorDetails {
    param([Parameter(Mandatory)]$ErrorRecord)

    $exception = $ErrorRecord.Exception
    $response = if ($null -ne $exception.PSObject.Properties['Response']) { $exception.Response } else { $null }
    $statusCode = 0
    if ($null -ne $response -and $null -ne $response.PSObject.Properties['StatusCode']) {
        try { $statusCode = [int]$response.StatusCode } catch { }
    }
    $message = [string]$exception.Message
    if ($statusCode -eq 0) {
        $match = [regex]::Match($message, '(?<!\d)(400|401|403|404|408|409|429|500|502|503|504)(?!\d)')
        if ($match.Success) { $statusCode = [int]$match.Groups[1].Value }
        elseif ($message -match '(?i)\bService\s*Unavailable\b') { $statusCode = 503 }
        elseif ($message -match '(?i)\bTooManyRequests\b|\bToo\s+Many\s+Requests\b') { $statusCode = 429 }
    }

    $retryAfter = ''
    $requestId = ''
    $clientRequestId = ''
    if ($null -ne $exception.Data) {
        if ($exception.Data.Contains('Retry-After')) { $retryAfter = [string]$exception.Data['Retry-After'] }
        if ($exception.Data.Contains('request-id')) { $requestId = [string]$exception.Data['request-id'] }
        if ($exception.Data.Contains('client-request-id')) { $clientRequestId = [string]$exception.Data['client-request-id'] }
    }
    if ($null -ne $response -and $null -ne $response.PSObject.Properties['Headers']) {
        $headers = $response.Headers
        if ([string]::IsNullOrWhiteSpace($retryAfter)) {
            try { $retryAfter = [string]$headers['Retry-After'] } catch { }
        }
        if ([string]::IsNullOrWhiteSpace($requestId)) {
            try { $requestId = [string]$headers['request-id'] } catch { }
        }
        if ([string]::IsNullOrWhiteSpace($clientRequestId)) {
            try { $clientRequestId = [string]$headers['client-request-id'] } catch { }
        }
    }
    if ([string]::IsNullOrWhiteSpace($requestId)) {
        $match = [regex]::Match($message, '(?i)(?<!client-)\brequest-id\s*:\s*([0-9a-f-]{36})')
        if ($match.Success) { $requestId = $match.Groups[1].Value }
    }
    if ([string]::IsNullOrWhiteSpace($clientRequestId)) {
        $match = [regex]::Match($message, '(?i)\bclient-request-id\s*:\s*([0-9a-f-]{36})')
        if ($match.Success) { $clientRequestId = $match.Groups[1].Value }
    }
    return [pscustomobject]@{
        StatusCode = $statusCode
        RetryAfter = $retryAfter
        RequestId = $requestId
        ClientRequestId = $clientRequestId
    }
}

function Invoke-IntuneExportRead {
    param(
        [Parameter(Mandatory)][string]$ReportName,
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][scriptblock]$RequestScript,
        [ValidateRange(1, 1800)][int]$RetryWindowSeconds = 600,
        [scriptblock]$SleepScript = { param($Seconds) Start-Sleep -Seconds $Seconds },
        [scriptblock]$ClockScript = { [datetime]::UtcNow }
    )

    $deadline = (& $ClockScript).AddSeconds($RetryWindowSeconds)
    $attempt = 0
    while ($true) {
        $attempt++
        try { return & $RequestScript }
        catch {
            $details = Get-IntuneExportErrorDetails -ErrorRecord $_
            $suffix = if ($details.RequestId) { "; RequestId=$($details.RequestId)" } else { '' }
            if ($details.ClientRequestId) { $suffix += "; ClientRequestId=$($details.ClientRequestId)" }
            $context = "Intune report '$ReportName' $Operation failed. Status=$($details.StatusCode); Attempt=$attempt$suffix."
            if ($details.StatusCode -notin @(408, 429, 500, 502, 503, 504)) { throw $context }

            $delay = [Math]::Min(180, [int](15 * [Math]::Pow(2, [Math]::Min($attempt - 1, 4))))
            if (-not [string]::IsNullOrWhiteSpace($details.RetryAfter)) {
                $seconds = 0
                if ([int]::TryParse($details.RetryAfter.Trim(), [ref]$seconds) -and $seconds -gt 0) {
                    $delay = $seconds
                }
                else {
                    $retryDate = [datetimeoffset]::MinValue
                    if ([datetimeoffset]::TryParse($details.RetryAfter.Trim(), [ref]$retryDate)) {
                        $delay = [Math]::Max(1, [int][Math]::Ceiling(($retryDate.UtcDateTime - (& $ClockScript)).TotalSeconds))
                    }
                }
            }
            if ((& $ClockScript).AddSeconds($delay) -gt $deadline) {
                throw "$context Retry window exhausted; no incomplete report was published."
            }
            Write-Warning "$context Retrying the same read in $delay second(s)."
            & $SleepScript $delay
        }
    }
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
    Write-Information "Intune report '$ReportName': creating export job ($ApiVersion)." -InformationAction Continue
    try { $job = & $post $uri }
    catch {
        $details = Get-IntuneExportErrorDetails -ErrorRecord $_
        $suffix = if ($details.RequestId) { "; RequestId=$($details.RequestId)" } else { '' }
        if ($details.ClientRequestId) { $suffix += "; ClientRequestId=$($details.ClientRequestId)" }
        throw "Intune report '$ReportName' export job creation failed. Status=$($details.StatusCode)$suffix. The POST outcome is unknown; no collector-level POST retry was attempted."
    }
    $jobId = Get-CleanText (Get-GraphValue $job 'id')
    if ([string]::IsNullOrWhiteSpace($jobId)) { throw "Intune report '$ReportName' returned no export job ID." }

    $deadline = [datetime]::UtcNow.AddSeconds($ExportTimeoutSeconds)
    do {
        $statusUri = "$uri/$([uri]::EscapeDataString($jobId))"
        $completed = Invoke-IntuneExportRead -ReportName $ReportName -Operation "job status (JobId=$jobId)" -RequestScript {
            Invoke-MgGraphRequest -Method GET -Uri $statusUri -ErrorAction Stop
        }
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
        Invoke-IntuneExportRead -ReportName $ReportName -Operation "archive download (JobId=$jobId)" -RequestScript {
            if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
            Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -ErrorAction Stop | Out-Null
        } | Out-Null
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

function Publish-IntuneAnalyticsBatch {
    param(
        [Parameter(Mandatory)]$Paths,
        [Parameter(Mandatory)][string[]]$TableNames,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Tables,
        [Parameter(Mandatory)][System.Collections.IDictionary]$LatestPaths,
        [Parameter(Mandatory)][System.Collections.IDictionary]$RowsByTable,
        [Parameter(Mandatory)][string]$ContractPath,
        [switch]$Fixture,
        [int]$MaxItems,
        [scriptblock]$BeforePromotionScript = { param($Destination, $Index) }
    )

    $batchId = [guid]::NewGuid().ToString('N')
    $stagingRoot = Join-Path $Paths.DataRootPath ('.staging\IntuneAnalytics\' + $batchId)
    $stagedLatestRoot = Join-Path $stagingRoot 'DATA-LAST'
    $backupRoot = Join-Path $stagingRoot 'backups'
    $stamp = [datetime]::UtcNow
    $promotions = New-Object System.Collections.Generic.List[object]
    $completedPromotions = New-Object System.Collections.Generic.List[object]
    $run = $null
    $completed = $false
    try {
        foreach ($name in $TableNames) {
            $table = $Tables[$name]
            $stagedPath = Join-Path $stagedLatestRoot (Join-Path ([string]$table.area) $name)
            Export-SmartWorkplaceCMDBCsv -InputObject @($RowsByTable[$name]) `
                -Columns @($table.columns | ForEach-Object { [string]$_ }) `
                -Path $stagedPath -TenantKey $Paths.TenantKey `
                -OrganizationKey $Paths.OrganizationKey -EnvironmentKey $Paths.EnvironmentKey `
                -TenantId $Paths.TenantId
            $baseName = [IO.Path]::GetFileNameWithoutExtension($name)
            $history = Join-Path $Paths.DataAllRootPath ('Intune\Analytics\{0}\{1}\{2}_{3}.csv' -f
                $stamp.ToString('yyyy'), $stamp.ToString('MM'), $baseName, $stamp.ToString('yyyyMMdd-HHmmssfff'))
            foreach ($destination in @($history, $LatestPaths[$name])) {
                $promotions.Add([pscustomobject]@{
                    Source = $stagedPath
                    Destination = [IO.Path]::GetFullPath($destination)
                    Backup = Join-Path $backupRoot ([guid]::NewGuid().ToString('N') + '.csv')
                })
            }
        }
        $results = @(Test-SmartWorkplaceCMDBCsvContract -LatestOutputRootPath $stagedLatestRoot -ContractPath $ContractPath)
        foreach ($name in $TableNames) {
            $result = @($results | Where-Object Name -eq $name)
            if ($result.Count -ne 1 -or $result[0].Status -ne 'Valid') {
                throw "Staged Intune analytics CSV '$name' does not satisfy its contract. No output was promoted."
            }
        }

        $run = Start-SmartWorkplaceCMDBSourceCollection -Paths $Paths `
            -RawPath @($TableNames | ForEach-Object { $LatestPaths[$_] }) -Fixture:$Fixture -MaxItems $MaxItems
        $promotionIndex = 0
        foreach ($promotion in $promotions) {
            $promotionIndex++
            & $BeforePromotionScript $promotion.Destination $promotionIndex
            $folder = Split-Path $promotion.Destination -Parent
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
            $hadPrevious = Test-Path -LiteralPath $promotion.Destination -PathType Leaf
            if ($hadPrevious) {
                New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
                Copy-Item -LiteralPath $promotion.Destination -Destination $promotion.Backup -Force
            }
            $candidate = $promotion.Destination + '.candidate.' + $batchId
            try {
                Copy-Item -LiteralPath $promotion.Source -Destination $candidate -Force
                Move-Item -LiteralPath $candidate -Destination $promotion.Destination -Force
            }
            finally {
                if (Test-Path -LiteralPath $candidate) { Remove-Item -LiteralPath $candidate -Force }
            }
            $completedPromotions.Add([pscustomobject]@{
                Destination = $promotion.Destination
                Backup = $promotion.Backup
                HadPrevious = $hadPrevious
            })
        }
        $publishedResults = @(Test-SmartWorkplaceCMDBCsvContract -LatestOutputRootPath $Paths.LatestOutputRootPath -ContractPath $ContractPath)
        foreach ($name in $TableNames) {
            $result = @($publishedResults | Where-Object Name -eq $name)
            if ($result.Count -ne 1 -or $result[0].Status -ne 'Valid') {
                throw "Published Intune analytics CSV '$name' does not satisfy its contract."
            }
        }
        Complete-SmartWorkplaceCMDBSourceCollection -Run $run
        $completed = $true
        return @($TableNames | ForEach-Object { $LatestPaths[$_] })
    }
    catch {
        $originalError = $_
        for ($index = $completedPromotions.Count - 1; $index -ge 0; $index--) {
            $promotion = $completedPromotions[$index]
            if ($promotion.HadPrevious) {
                Copy-Item -LiteralPath $promotion.Backup -Destination $promotion.Destination -Force
            }
            elseif (Test-Path -LiteralPath $promotion.Destination) {
                Remove-Item -LiteralPath $promotion.Destination -Force
            }
        }
        if ($null -ne $run -and -not $completed) {
            Complete-SmartWorkplaceCMDBSourceCollection -Run $run -Failed
        }
        throw $originalError
    }
    finally {
        if (Test-Path -LiteralPath $stagingRoot) {
            Remove-Item -LiteralPath $stagingRoot -Recurse -Force
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
$tableNames = @(
    'Intune_WindowsUpdateAlerts.csv',
    'Intune_EndpointAnalyticsDeviceScores.csv',
    'Intune_EndpointAnalyticsUpgradeEligibility.csv'
)
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
            SourceMode = if ($null -ne $fixture) { 'OfflineJson' } else { 'MicrosoftGraphExportJobsAndEndpointAnalytics' }
            RawContractVersion = [string]$contract.contractVersion
            RequiredGraphPermissions = 'DeviceManagementConfiguration.Read.All;DeviceManagementManagedDevices.ReadWrite.All'
            OutputCount = $tableNames.Count
        } | Format-List
        return
    }

    $alertSourceRows = @()
    $analyticsSourceRows = @()
    $upgradeEligibilitySourceRows = @()
    $upgradeEligibilityCollectionStatus = 'NotStarted'
    if ($null -ne $fixture) {
        $alertSourceRows = @($fixture.windowsUpdateAlerts)
        $analyticsSourceRows = @($fixture.endpointAnalyticsDeviceScores)
        $upgradeEligibilitySourceRows = @($fixture.endpointAnalyticsUpgradeEligibility)
        $upgradeEligibilityCollectionStatus = 'Fixture'
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
        $upgradeEligibilityResult = Get-UpgradeEligibilityGraphRows
        $upgradeEligibilitySourceRows = @($upgradeEligibilityResult.Rows)
        $upgradeEligibilityCollectionStatus = [string]$upgradeEligibilityResult.Status
    }

    if ($MaxItems -gt 0) {
        $alertSourceRows = @($alertSourceRows | Select-Object -First $MaxItems)
        $analyticsSourceRows = @($analyticsSourceRows | Select-Object -First $MaxItems)
        $upgradeEligibilitySourceRows = @($upgradeEligibilitySourceRows | Select-Object -First $MaxItems)
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
    $alertGroups = @($alertRows | Group-Object { Get-UpdateAlertEvidenceKey $_ })
    $duplicateAlertRowCount = 0
    $collapsedAlertRows = New-Object System.Collections.Generic.List[object]
    foreach ($group in @($alertGroups | Sort-Object Name)) {
        if ($group.Count -gt 1) { $duplicateAlertRowCount += $group.Count - 1 }
        $selected = @($group.Group | Sort-Object DeviceName, SourceReport, PolicyId, DeviceId, EventDateTimeUTC, LastWUScanTimeUTC, AggregateState, CurrentDeviceUpdateStatus, LatestAlertMessage)[0]
        $collapsedAlertRows.Add($selected)
    }
    $alertRows = @($collapsedAlertRows.ToArray())
    $distinctEvidenceCollisionCount = @($alertRows |
        Group-Object { Get-UpdateAlertBaseKey $_ } |
        Where-Object Count -gt 1).Count
    if ($duplicateAlertRowCount -gt 0 -or $distinctEvidenceCollisionCount -gt 0) {
        Write-Warning ((
            'Intune returned {0} repeated Windows update evidence row(s) and {1} event key(s) with distinct evidence. ' +
            'Repeated evidence was collapsed; distinct source evidence was retained with deterministic fingerprints.') -f
            $duplicateAlertRowCount,
            $distinctEvidenceCollisionCount)
    }
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
    $upgradeEligibilityRows = @($upgradeEligibilitySourceRows | ForEach-Object {
        $deviceId = Get-CleanText (Get-GraphValue $_ 'deviceId')
        $metricDeviceId = Get-CleanText (Get-GraphValue $_ 'id')
        $metricId = Get-CleanText (Get-GraphValue $_ 'sourceMetricId')
        $deviceIdSource = 'ReportedDeviceId'
        if ([string]::IsNullOrWhiteSpace($deviceId) -and $upgradeEligibilityCollectionStatus -eq 'CollectedBetaAllDevices') {
            $deviceId = $metricDeviceId
            $deviceIdSource = 'BetaMetricDeviceId'
        }
        if ([string]::IsNullOrWhiteSpace($deviceId)) { throw 'Endpoint Analytics upgrade eligibility row has no usable Intune device ID.' }
        $eligibility = Get-UpgradeEligibilityText (Get-GraphValue $_ 'upgradeEligibility') $deviceId
        if ([string]::IsNullOrWhiteSpace($eligibility)) { throw "Endpoint Analytics upgrade eligibility is missing for '$deviceId'." }
        [pscustomobject][ordered]@{
            SourceSystem = 'MicrosoftGraphEndpointAnalytics'
            MetricId = $metricId
            MetricDeviceId = $metricDeviceId
            DeviceId = $deviceId
            DeviceIdSource = $deviceIdSource
            DeviceName = Get-CleanText (Get-GraphValue $_ 'deviceName')
            UpgradeEligibility = $eligibility
            SourceCollectedDateTime = $collected
        }
    })
    if ($upgradeEligibilityRows.Count -gt 0) {
        $collapsedUpgradeEligibilityRows = New-Object System.Collections.Generic.List[object]
        foreach ($group in @($upgradeEligibilityRows | Group-Object DeviceId | Sort-Object Name)) {
            $eligibilityValues = @($group.Group | ForEach-Object { [string]$_.UpgradeEligibility } | Sort-Object -Unique)
            if ($eligibilityValues.Count -gt 1) {
                throw "Conflicting Endpoint Analytics upgrade eligibility values returned for device '$($group.Name)': $($eligibilityValues -join ', ')."
            }
            $selected = @($group.Group | Sort-Object MetricId, MetricDeviceId)[0]
            $collapsedUpgradeEligibilityRows.Add($selected)
        }
        $upgradeEligibilityRows = @($collapsedUpgradeEligibilityRows.ToArray())
    }
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
        'Intune_EndpointAnalyticsUpgradeEligibility.csv' = $upgradeEligibilityRows
    }
    foreach ($name in $tableNames) {
        $key = if ($name -eq 'Intune_WindowsUpdateAlerts.csv') { { Get-UpdateAlertEvidenceKey $_ } } else { 'DeviceId' }
        $duplicates = @($rowsByTable[$name] | Group-Object $key | Where-Object Count -gt 1)
        if ($duplicates.Count) { throw "Duplicate Intune analytics keys returned for ${name}: $($duplicates.Name -join ', ')" }
    }

    $published = @(Publish-IntuneAnalyticsBatch -Paths $paths -TableNames $tableNames `
        -Tables $tables -LatestPaths $latestPaths -RowsByTable $rowsByTable `
        -ContractPath $rawContractPath -Fixture:($PSCmdlet.ParameterSetName -eq 'Fixture') -MaxItems $MaxItems)
    Write-Information ("SmartWorkplaceCMDB Intune analytics collection completed. UpdateRows={0}; EndpointAnalyticsDevices={1}; UpgradeEligibilityDevices={2}; UpgradeEligibilityStatus={3}." -f $alertRows.Count, $analyticsRows.Count, $upgradeEligibilityRows.Count, $upgradeEligibilityCollectionStatus) -InformationAction Continue
    [pscustomobject]@{
        Status = 'Completed'
        ScriptVersion = $ScriptVersion
        WindowsUpdateRowCount = $alertRows.Count
        EndpointAnalyticsDeviceCount = $analyticsRows.Count
        UpgradeEligibilityDeviceCount = $upgradeEligibilityRows.Count
        UpgradeEligibilityCollectionStatus = $upgradeEligibilityCollectionStatus
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBWNtktz2xckR56
# t2CpePEWZ77NLyTC6/Q20JZmMeJ4+6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIE8Qn2iuJc8kzgkmuqZG8VRbrqbh7KpdPHJsWgvUIkjRMA0GCSqG
# SIb3DQEBAQUABIIBgA2yL61DSWXrseTc/QS5Xxap33NWilgvhFDbwhpIUTAz3Ixw
# BdDdeBuT9LDT5kPpKbRaxKvQgTqxuG0hcDpkefe2lPe+MF7mYwGCF6EmIJM0nssx
# WRVCXSDfVRBUO5UL8S1gDCQGKy+Xvr1hfehYAMfm+J1HNR1NWb52Vrl+O3gWvKVa
# n3FLwIR885Ivzb7jUU7nLSnsADFmWJXES6VbBVB0V8AuS5gFmkXKtsjyJiGnlyjs
# BsLml1kHgGXRL9+R4FgQtt7TkaI5XRBQj93SAXIcHYx4ugvFoCxmDIqB6649q8Vm
# YmoXPu6KJzuM5/WKQUS+AvQ6SB0B8HFwwWtZ8DV2RQ3Wmtq2QNdNculCI7q6+VH7
# TDpGlnWliQ+sYT60m3KbjQtNQIeWX6a3HN0l3VFvs4EbBQUctpi+logvou8Vw2z+
# 3zRyq9ZPEHikhUKvjiOEeVHI2xtD5Owi1UH7WEVOqDcD0oYIYNS/+BLy9vWVwXQv
# ZeGD2lY+a9p36PNquKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjIxODU0
# MzFaMC8GCSqGSIb3DQEJBDEiBCBGsu1/3kY3/s0hYCfcipsaOiY39FN9L/yI2Ynz
# /2xwFjANBgkqhkiG9w0BAQEFAASCAgBWPyBetwk/WqYv6QXjmK20J31oyxDbtaOF
# 0LM5lqZzvm5/lXo8KkAHlaIxLKcEZWRBqqa/sFlV2Fe17XxrHhEBKAWDFBzsKdEb
# c6yViszDC5xVqZifeB05/wUvreR/6XS7YGdrFX1OwayDng0ywhVTd9Ju9llD7JhI
# dgWMPqvyk0rDtdNnoh+pbGi+T1zt2CFEKm8GR8i0LI+RuD/v8uFzHWJFqpWKFi5W
# yO4q0KGzPMflRHTdw6JZ6XHh6NTFv0VBW/9Ojfq83A4K7cv/IfEoZNA/6W8UYzp0
# cQX1LFrY6DKwur1ChTyV01r4sv4rJZhuw1VnXHu6SVHqM3SOflXtAwh5i64o8l5U
# MvE4R7UrnQsWYWXEkMd8aM4fayyfzrxcVOYtm2yGIroO4CGVkDux+ZZDG7yZ8SRt
# OfsMS9jHpqJKVOaL5u5I+mvENkr3++EdFZuHXBtMlh32/t4kWerAECQjjDZ6xA1O
# BePz8OfUtUTnJITYxVpYuPmExsznoKE7JqX4JFJ6bf/la3lthrkTPtYg43fRzlnV
# RDuyZmwhToulKt0JMsLkstfCovQ+bmaHLfvoxMz8qBJAqxEpwtf1joKajXODBRRs
# zxEvtdvevJ1dmA3nPMVewixqnuHdcfdqMhREGyXpIWgQZC5vKXdkYu8MVTxVY9zK
# AcPRheejxg==
# SIG # End signature block
