Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-SmartM365EvidenceProperty {
    param([AllowNull()]$InputObject,[Parameter(Mandatory)][string[]]$Names,[AllowNull()]$DefaultValue='')
    if ($null -eq $InputObject) { return $DefaultValue }
    foreach ($name in $Names) {
        if ($InputObject -is [System.Collections.IDictionary] -and $InputObject.Contains($name)) { return $InputObject[$name] }
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }
    return $DefaultValue
}

function ConvertTo-SmartM365EvidenceText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return (@($Value) | ForEach-Object { [string]$_ }) -join ';'
    }
    return [string]$Value
}

function ConvertTo-SmartM365EvidenceJson {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ($Value | ConvertTo-Json -Depth 20 -Compress)
}

function Resolve-SmartM365EvidenceToken {
    param([AllowNull()]$Value,[Parameter(Mandatory)]$EffectiveConfig)
    if ($Value -isnot [string]) { return $Value }
    $resolved = [string]$Value
    for ($iteration=0; $iteration -lt 10; $iteration++) {
        $tokenMatches = [regex]::Matches($resolved,'\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $property = $EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $resolved = $resolved.Replace($match.Value,[string]$property.Value)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-SmartM365EvidenceConfig {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)]$EffectiveConfig
    )
    $base = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $folder = Split-Path -Path $ScriptPath -Parent
    $localPath = Join-Path $folder "$base.local.json"; $localPath = Resolve-SmartM365JsonConfigurationPath -Path $localPath
    $sourcePath = $localPath
    if (-not (Test-Path -LiteralPath $localPath -PathType Leaf -ErrorAction Stop)) {
        $names = Get-SmartM365JsonNames -Path $localPath
        $templatePaths = @(($names.Preferred + '.template'), (Get-SmartM365JsonTemplateName -Path $localPath))
        $sourcePath = $null
        foreach ($candidate in $templatePaths) {
            if (Test-Path -LiteralPath $candidate -ErrorAction Stop) {
                if (-not (Test-Path -LiteralPath $candidate -PathType Leaf -ErrorAction Stop)) {
                    throw "Collector configuration template is not a file: $candidate"
                }
                $sourcePath = $candidate
                break
            }
        }
        if ($null -eq $sourcePath) { throw "Collector configuration templates not found: $($templatePaths -join '; ')" }
    }
    # A malformed preferred file must fail, never silently select older configuration.
    $local = Get-Content -LiteralPath $sourcePath -Raw | ConvertFrom-Json -NoEnumerate -ErrorAction Stop
    if ($local -isnot [System.Management.Automation.PSCustomObject]) { throw "Collector configuration must be a JSON object: $sourcePath" }
    $result = [ordered]@{}
    foreach ($property in $EffectiveConfig.PSObject.Properties) { $result[$property.Name] = $property.Value }
    foreach ($property in $local.PSObject.Properties) {
        if ($property.Value -is [string] -and $property.Value.Trim() -in @('__USE_GLOBAL__','USE_GLOBAL')) { continue }
        $result[$property.Name] = Resolve-SmartM365EvidenceToken -Value $property.Value -EffectiveConfig ([pscustomobject]$result)
    }
    foreach ($key in @($result.Keys)) { $result[$key] = Resolve-SmartM365EvidenceToken -Value $result[$key] -EffectiveConfig ([pscustomobject]$result) }
    return [pscustomobject]$result
}

function Initialize-SmartM365EvidenceRuntime {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)]$EffectiveConfig,
        [Parameter(Mandatory)][string]$DefaultOutputRelativePath,
        [string]$OutputPath,
        [string]$LatestCsvFolderPath,
        [switch]$ValidateOnly,
        [switch]$EnableConfiguredExternalActions
    )
    $config = Get-SmartM365EvidenceConfig -ScriptPath $ScriptPath -EffectiveConfig $EffectiveConfig
    $root = [string](Get-SmartM365EvidenceProperty $EffectiveConfig @('SmartM365RootPath'))
    $coreManifest = Join-Path $root 'Modules\SmartM365.Core\SmartM365.Core.psd1'
    # Keep receipt exports visible to the caller and preserve the loaded Core context.
    Import-Module -Name $coreManifest -MinimumVersion '1.0.99' -Global -ErrorAction Stop
    $dataAllRoot = [string](Get-SmartM365EvidenceProperty $config @('DataAllRootPath'))
    $resolvedOutput = if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath } elseif (-not [string]::IsNullOrWhiteSpace([string](Get-SmartM365EvidenceProperty $config @('ScriptCsvLogFolderPath')))) { [string](Get-SmartM365EvidenceProperty $config @('ScriptCsvLogFolderPath')) } else { Join-Path $dataAllRoot $DefaultOutputRelativePath }
    $resolvedLatest = if (-not [string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $LatestCsvFolderPath } else { [string](Get-SmartM365EvidenceProperty $config @('LatestCsvFolderPath')) }
    if ([string]::IsNullOrWhiteSpace($resolvedOutput) -or [string]::IsNullOrWhiteSpace($resolvedLatest)) { throw 'Collector output paths are not configured.' }

    $externalActionsEnabled = $EnableConfiguredExternalActions -and -not $ValidateOnly
    $global:EnableSharePointUpload = $externalActionsEnabled -and [bool](Get-SmartM365EvidenceProperty $config @('EnableSharePointUpload') $false)
    $global:EnableTeamsNotifications = $externalActionsEnabled -and [bool](Get-SmartM365EvidenceProperty $config @('EnableTeamsNotifications') $false)
    foreach ($name in @('SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath','AppId','TenantId','Thumb','Thumbprint','SmtpServer','From','To','Cc','ErrorMailTo')) {
        Set-Variable -Name $name -Scope Global -Value ([string](Get-SmartM365EvidenceProperty $config @($name)))
    }
    $global:RetentionMaxCSV = [int](Get-SmartM365EvidenceProperty $config @('RetentionMaxCSV') 30)
    $global:RetentionMaxLogs = [int](Get-SmartM365EvidenceProperty $config @('RetentionMaxLogs') 30)
    InitializeScriptEnvironment -OutputPathInit $resolvedOutput -LogFileName ([System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)) -CallerScriptPath $ScriptPath | Out-Null
    return [pscustomobject]@{ Config=$config; OutputPath=$resolvedOutput; LatestCsvFolderPath=$resolvedLatest; RunId=[guid]::NewGuid().Guid; CollectedAtUtc=[datetime]::UtcNow.ToString('o') }
}

# The Graph SDK retries 429/503/504 by itself (MaxRetry 3) before Invoke-MgGraphRequest throws: stacked under
# this module's retries that made up to 4 requests per attempt and hid Retry-After ("Too many retries
# performed"). -DisableSdkRetry turns the SDK retry off for the whole process so that this module is the only
# retry level. It is set here, right after Connect-MgGraph and before any request, and never changed later:
# with SDK 2.30 an absolute-URI Invoke-MgGraphRequest empties the session's AzureADEndpoint, and the client
# that Set-MgRequestContext forces to be rebuilt then fails app-only certificate auth with "Invalid URI"
# before sending. Connect-MgGraph restores the environment, so this point is safe.
function Connect-SmartM365EvidenceGraph {
    param([Parameter(Mandatory)]$Runtime,[Parameter(Mandatory)][string[]]$Scopes,[switch]$InteractiveAuth,[switch]$DisableSdkRetry)
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    try { if (Get-MgContext -ErrorAction SilentlyContinue) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false } } catch {}
    if ($InteractiveAuth) { Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop | Out-Null }
    else {
        $appId = [string](Get-SmartM365EvidenceProperty $Runtime.Config @('AppId'))
        $tenantId = [string](Get-SmartM365EvidenceProperty $Runtime.Config @('TenantId'))
        $thumb = [string](Get-SmartM365EvidenceProperty $Runtime.Config @('Thumbprint','Thumb'))
        if ([string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($tenantId) -or [string]::IsNullOrWhiteSpace($thumb)) { throw 'AppId, TenantId, and Thumbprint are required for app-only Graph authentication.' }
        Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $thumb -NoWelcome -ErrorAction Stop | Out-Null
    }
    Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $true
    if ($DisableSdkRetry) { Set-MgRequestContext -MaxRetry 0 | Out-Null }
}

function Get-SmartM365EvidenceStatusCode {
    param([Parameter(Mandatory)]$ErrorRecord)
    # Response, Graph SDK status or "HTTP 503"/"Status: 500" text; a number elsewhere (a URL) is not a status.
    $status = Get-SmartM365GraphErrorStatusCode -ErrorRecord $ErrorRecord
    if ($status) { return [int]$status }
    $message = [string]$ErrorRecord.Exception.Message
    if ($message -match '(?i)\bTooManyRequests\b') { return 429 }
    if ($message -match '(?i)\bServiceUnavailable\b') { return 503 }
    return 0
}

# Clock seam: offline tests replace it to drive retry budgets and export deadlines without waiting.
function Get-SmartM365EvidenceUtcNow { [datetime]::UtcNow }

# Read only: the SDK retry is never changed here (see Connect-SmartM365EvidenceGraph -DisableSdkRetry).
function Assert-SmartM365EvidenceSdkRetryOff {
    param([string]$Operation='')
    if (-not (Get-Command Get-MgRequestContext -ErrorAction SilentlyContinue)) { return }
    $maxRetry = [int](Get-MgRequestContext).MaxRetry
    if ($maxRetry -ne 0) { Write-SmartM365EvidenceLog "$Operation runs with the Graph SDK retry on (MaxRetry $maxRetry): connect with -DisableSdkRetry so that the retry budget is the only retry level." 'WARNING' }
}

function Invoke-SmartM365EvidenceGraphRequest {
    # MaxRetryCount bounds the attempts; a RetryBudgetSeconds above 0 bounds the elapsed time instead.
    param([ValidateSet('GET','POST')][string]$Method='GET',[Parameter(Mandatory)][string]$Uri,[AllowNull()][string]$Body,[int]$MaxRetryCount=6,[ValidateRange(0,86400)][int]$RetryBudgetSeconds=0,[string]$Operation='',[scriptblock]$Invoker,[switch]$NoSleep)
    if ([string]::IsNullOrWhiteSpace($Operation)) {
        $path = try { ([uri]$Uri).AbsolutePath } catch { '<invalid uri>' }
        $Operation = "Graph $Method $path"
    }
    $started = Get-SmartM365EvidenceUtcNow
    for ($attempt=1; ; $attempt++) {
        try {
            if ($Invoker) { return & $Invoker $Method $Uri $Body }
            $parameters = @{Method=$Method;Uri=$Uri;OutputType='PSObject';ErrorAction='Stop'}
            if ($Method -eq 'POST') { $parameters.Body=$Body; $parameters.ContentType='application/json' }
            return Invoke-MgGraphRequest @parameters
        }
        catch {
            $status = Get-SmartM365EvidenceStatusCode $_
            # 401/403 never retried; 408/409/429/5xx, timeouts and connection resets retried.
            $retryable = $status -notin @(401,403) -and ($status -in @(408,409,429,500,502,503,504) -or ($status -eq 0 -and (Test-SmartM365GraphTransientError -ErrorRecord $_ -StatusCode $null)))
            if (-not $retryable) { throw }
            $statusText = if ($status) { "HTTP $status" } else { 'a transient error' }
            $elapsed = [int][Math]::Floor(((Get-SmartM365EvidenceUtcNow) - $started).TotalSeconds)
            $wait = Get-SmartM365RetryDelay -Attempt $attempt -ErrorRecord $_ -BaseSeconds 5 -MaximumSeconds 300
            $exhausted = if ($RetryBudgetSeconds -gt 0) { $elapsed + $wait.Seconds -gt $RetryBudgetSeconds } else { $attempt -ge $MaxRetryCount }
            if ($exhausted) {
                $failure = [System.IO.IOException]::new(("{0} failed after {1} attempt(s) over {2}s ({3}). Last error: {4}" -f $Operation,$attempt,$elapsed,$statusText,$_.Exception.Message), $_.Exception)
                if ($status) { $failure.Data['StatusCode'] = $status }
                $failure.Data['Attempts'] = $attempt
                throw $failure
            }
            $limit = if ($RetryBudgetSeconds -gt 0) { "budget ${RetryBudgetSeconds}s" } else { "attempt limit $MaxRetryCount" }
            # INFO as before: a retry that recovers must not turn a successful run into CompletedWithWarnings.
            if (-not $Invoker) { Write-SmartM365EvidenceLog ("{0} returned {1} on attempt {2}; retrying in {3}s ({4}; elapsed {5}s, {6})." -f $Operation,$statusText,$attempt,$wait.Seconds,$wait.Source,$elapsed,$limit) }
            if (-not $NoSleep) { Start-Sleep -Seconds $wait.Seconds }
        }
    }
}

function Get-SmartM365EvidenceGraphCollection {
    param([Parameter(Mandatory)][string]$Uri,[ValidateRange(1,100000)][int]$MaxPages=2000,[int]$MaxItems=0,[scriptblock]$Invoker)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    $page = 0
    while (-not [string]::IsNullOrWhiteSpace($next) -and $page -lt $MaxPages) {
        $page++
        $response = Invoke-SmartM365EvidenceGraphRequest -Uri $next -Invoker $Invoker -NoSleep:$($null-ne$Invoker)
        $values = @(Get-SmartM365EvidenceProperty $response @('value') @())
        foreach ($value in $values) {
            $items.Add($value)
            if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { return @($items) }
        }
        $next = [string](Get-SmartM365EvidenceProperty $response @('@odata.nextLink') '')
    }
    if (-not [string]::IsNullOrWhiteSpace($next)) { throw "Graph pagination exceeded MaxPages=$MaxPages." }
    return @($items)
}

function Invoke-SmartM365EvidenceExportReport {
    param(
        [Parameter(Mandatory)][string]$ReportName,
        [string]$Filter='',
        [string[]]$Select=@(),
        [ValidateSet('v1.0','beta')][string]$ApiVersion='v1.0',
        [int]$TimeoutSeconds=600,
        [int]$PollSeconds=5,
        [int]$MaxItems=0,
        [ValidateRange(0,86400)][int]$CreateRetryBudgetSeconds=0,
        [ValidateRange(0,86400)][int]$ProgressLogSeconds=90,
        # A job still notStarted/inProgress after this many seconds is abandoned and created again, once; 0 = never.
        [ValidateRange(0,86400)][int]$StallSeconds=0
    )
    $base = 'https://graph.microsoft.com'
    $uri = "$base/$ApiVersion/deviceManagement/reports/exportJobs"
    $requestBody = [ordered]@{reportName=$ReportName;format='csv';localizationType='replaceLocalizableValues'}
    if (-not [string]::IsNullOrWhiteSpace($Filter)) { $requestBody.filter=$Filter }
    if (@($Select).Count -gt 0) { $requestBody.select=@($Select) }
    $body = $requestBody | ConvertTo-Json -Depth 8 -Compress
    # A CreateRetryBudgetSeconds above 0 gives the job creation that much time and status polls up to 300
    # seconds each; it expects the SDK retry off from the connection (warned otherwise, never changed here).
    $budgeted = $CreateRetryBudgetSeconds -gt 0
    $createRetry = if ($budgeted) { @{RetryBudgetSeconds=$CreateRetryBudgetSeconds} } else { @{} }
    $pollRetry = if ($budgeted) { @{RetryBudgetSeconds=[Math]::Min(300,$CreateRetryBudgetSeconds)} } else { @{} }
    if ($budgeted) { Assert-SmartM365EvidenceSdkRetryOff -Operation "Intune export job creation ($ReportName)" }
    $deadline = $null
    $recreated = $false
    :job while ($true) {
        $job = Invoke-SmartM365EvidenceGraphRequest -Method POST -Uri $uri -Body $body -Operation "Intune export job creation ($ReportName)" @createRetry
        $jobId = [string](Get-SmartM365EvidenceProperty $job @('id'))
        if ([string]::IsNullOrWhiteSpace($jobId)) { throw "No export job id returned for $ReportName." }
        $jobStarted = Get-SmartM365EvidenceUtcNow
        if ($null -eq $deadline) { $reportStarted = $jobStarted; $deadline = $jobStarted.AddSeconds($TimeoutSeconds) }
        Write-SmartM365EvidenceLog "Intune export job $jobId created for $ReportName."
        $nextProgress = $jobStarted.AddSeconds($ProgressLogSeconds)
        while ($true) {
            $job = Invoke-SmartM365EvidenceGraphRequest -Uri "$uri/$([uri]::EscapeDataString($jobId))" -Operation "Intune export job status ($ReportName)" @pollRetry
            $status = [string](Get-SmartM365EvidenceProperty $job @('status'))
            if ($status -ieq 'completed') { break job }
            if ($status -ieq 'failed') { throw "Intune export job $jobId failed for $ReportName." }
            $now = Get-SmartM365EvidenceUtcNow
            $jobElapsed = [int][Math]::Floor(($now - $jobStarted).TotalSeconds)
            $totalElapsed = [int][Math]::Floor(($now - $reportStarted).TotalSeconds)
            if ($ProgressLogSeconds -gt 0 -and $now -ge $nextProgress) {
                Write-SmartM365EvidenceLog "Intune export job $jobId for $ReportName is '$status' after ${jobElapsed}s (report elapsed ${totalElapsed}s of ${TimeoutSeconds}s)."
                $nextProgress = $now.AddSeconds($ProgressLogSeconds)
            }
            if ($now -ge $deadline) {
                $note = if ($recreated) { ' The job had already been created again once.' } else { '' }
                throw "Intune export job $jobId for $ReportName timed out after ${totalElapsed}s (limit ${TimeoutSeconds}s); last status '$status'.$note"
            }
            if ($StallSeconds -gt 0 -and -not $recreated -and $jobElapsed -ge $StallSeconds -and ($status -in @('notStarted','inProgress') -or [string]::IsNullOrWhiteSpace($status))) {
                Write-SmartM365EvidenceLog "Intune export job $jobId for $ReportName is still '$status' after ${jobElapsed}s; abandoning it and creating the job again (once)." 'WARNING'
                $recreated = $true
                continue job
            }
            Start-Sleep -Seconds $PollSeconds
        }
    }
    $downloadUrl = [string](Get-SmartM365EvidenceProperty $job @('url'))
    # The download URL can lag the completed status by a few seconds.
    for ($urlPoll = 1; [string]::IsNullOrWhiteSpace($downloadUrl) -and $urlPoll -le 3; $urlPoll++) {
        Write-SmartM365EvidenceLog "Completed export job for $ReportName has no URL yet; polling again ($urlPoll/3)." 'WARNING'
        Start-Sleep -Seconds $PollSeconds
        $job = Invoke-SmartM365EvidenceGraphRequest -Uri "$uri/$([uri]::EscapeDataString($jobId))" -Operation "Intune export job status ($ReportName)" @pollRetry
        $downloadUrl = [string](Get-SmartM365EvidenceProperty $job @('url'))
    }
    if ([string]::IsNullOrWhiteSpace($downloadUrl)) { throw "Completed export job returned no URL for $ReportName." }
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) "SmartM365-Evidence-$([guid]::NewGuid().Guid)"
    try {
        $extract = Join-Path $tempRoot 'content'; $zip = Join-Path $tempRoot 'export.zip'
        New-Item -Path $extract -ItemType Directory -Force | Out-Null
        Invoke-SmartM365ExportFileDownload -Url $downloadUrl -OutFile $zip -Operation "Intune export download ($ReportName)" | Out-Null
        Expand-Archive -LiteralPath $zip -DestinationPath $extract -Force
        $csv = Get-ChildItem -LiteralPath $extract -Filter '*.csv' -File -Recurse | Select-Object -First 1
        if (-not $csv) { throw "No CSV found in the Intune export for $ReportName." }
        $rows = @(Import-Csv -LiteralPath $csv.FullName)
        if ($MaxItems -gt 0) { $rows = @($rows | Select-Object -First $MaxItems) }
        return $rows
    }
    finally {
        if (Test-Path -LiteralPath $tempRoot) {
            try { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction Stop }
            catch { Write-Warning "Temporary report cleanup failed: $($_.Exception.Message)" }
        }
    }
}

function Export-SmartM365EvidenceDataset {
    param([Parameter(Mandatory)]$Runtime,[Parameter(Mandatory)][string]$BaseFileName,[AllowEmptyCollection()][object[]]$Rows,[Parameter(Mandatory)][string[]]$Columns,[switch]$NoWeeklyHistory,[switch]$SingleSerialization,[switch]$NoTenantKey,[switch]$LocalSerialization)
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $timestamped = Join-Path $Runtime.OutputPath "${BaseFileName}_${stamp}.csv"
    $latest = Join-Path $Runtime.LatestCsvFolderPath "${BaseFileName}.csv"
    $options = @{}
    if ($SingleSerialization) { $options.SingleSerialization = $true }
    # Rows that already carry the tenant identity columns skip the generic enrichment.
    if ($NoTenantKey) { $options.NoTenantKey = $true }
    if ($LocalSerialization) { $options.LocalSerialization = $true }
    return Export-SmartM365Csv -Data $Rows -TimestampedPath $timestamped -LatestPath $latest -Columns $Columns -NoWeeklyHistory:$NoWeeklyHistory @options
}

function Write-SmartM365EvidenceLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR','DEBUG')][string]$Level='INFO'
    )
    WriteLog -Message $Message -Level $Level
}

function Complete-SmartM365EvidenceRuntime {
    param(
        [ValidateSet('Auto','Success','CompletedWithWarnings','Failed','Cancelled')][string]$Status='Auto',
        [AllowNull()]$ErrorRecord,
        [string]$FailureStage='',
        [switch]$CloseTranscriptBeforeUpload
    )
    $options = @{}
    if ($CloseTranscriptBeforeUpload) { $options.CloseTranscriptBeforeUpload = $true }
    Complete-SmartM365ExecutionContext -Status $Status -ErrorRecord $ErrorRecord -FailureStage $FailureStage @options
}

Export-ModuleMember -Function Get-SmartM365EvidenceProperty,ConvertTo-SmartM365EvidenceText,ConvertTo-SmartM365EvidenceJson,Get-SmartM365EvidenceConfig,Initialize-SmartM365EvidenceRuntime,Connect-SmartM365EvidenceGraph,Get-SmartM365EvidenceStatusCode,Invoke-SmartM365EvidenceGraphRequest,Get-SmartM365EvidenceGraphCollection,Invoke-SmartM365EvidenceExportReport,Export-SmartM365EvidenceDataset,Write-SmartM365EvidenceLog,Complete-SmartM365EvidenceRuntime

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB1Ja02SV7Ddrdr
# LvUV9Bq+bbat2czWOcX1K5dLE6GWJaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIVEdCX/lipb72beK7BdruDsWDn6VCqGMhD4c1XQUs0xMA0GCSqG
# SIb3DQEBAQUABIIBgB44WS5DSgGoxl8aq1hmgnPzMzhmvFzJJwLEKCUVKUWmU6xv
# BJJQWPx2MKgRqayZaUr7PLzP2RFJqs6hMgIKB/BR1OD7Izu6AvdTIr+Nh757i9N8
# qa+8lc+d2AgyOkqPUmVF79+CWXoEEBz1Cj4es9623P6omZv8u0YfbU3zVJOXg3mt
# PqjMGf8Nz+ldZP2oSmUFB1pasB8rki44TKKOXVEBzUVqA24fVA4CoCd4ZATOj5E9
# jpLpfBvosfVq6yLQ+iMwMJG4fIrpZAb7afwF00wEqrX8YoPTcNvjAd8WXeN0T6LY
# 2HDlzYhDBIkJANUxxw5FBRa1+fBIEFpKVbrSacGuEXF/5j2r1QS0JVXh/YFDTW2s
# 2ZEOvBfH9U6CY51ESRketQTGmDNju7ztV+MCg0nqXp9qhzxctQHihXOTmvJs9nX0
# GiJM24vYygBLI/vijC10VHyvQO8FsNO2rIGsj0AROeDOLoRqxO/6CNJz95KvTOwt
# s3TJkNuPinnr5DlS0KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIxNzI5
# NThaMC8GCSqGSIb3DQEJBDEiBCA4sZQ+p++cUEhXmhf954ckHkhI46Y0XSrtiIIu
# 5yI7RDANBgkqhkiG9w0BAQEFAASCAgBK5aVODeoV6Gu4/CRIlBE9PDSCNCOIK+h5
# CUq+aiOQ53LT/PRB3L8Vkh2K24TcoGnlF+bbufvwpV7QSp5yaeVbE5kaiRU8mwrQ
# ySj2afZfNE+vVPdmxaWCKgEAihzWs/xJJ2zlYboVyXGWv5nJfwONWEvCSGSiefmC
# M0pLPZnQC48vswD5s9pNSDtpyt/uC4KfYzkt/CU0ZonzjjvbDomcO3rxUm/dtZM3
# 89ukXGWjEUqneMLtq0vFIK6igEi34AMhHSn3QrbUTOEjCab0GmfRuGBzr/92PwGl
# q2Rl2ya8uuMYIXJiKRmtFn8MkE33OQmqjb31HEJl3reYan4Qw7XGw7FQIU30NDHb
# aH81dB92c2Ec3H2L8kA+6OhURkvYG86tgN5wuSCoELUfEN+C86oGV/OVTHTogbAP
# /m0K72x/PTd3Wa7nsqnW6Pldk7m6R/4YrK4RkcZHC9BZDasoOceEBnGX2l9QAGWD
# BOgJgteGSuLhgE7xL8vHQQOPjF6lsbMiD7e6hwMvqJq7LZ8H1g/mL/QpoXUa0lO9
# jGS0vvDiIrTUB0p/x0aCiDkRcDKrAAavt5OPPZ0nKu1IOA3BlU0WaJUh7i+JORW4
# rrqvRFQdvjH1IUzt5oAKuIc1Iyuv6mS+STclKVaSj8XjUuwHQZWrC26T/WF4yrtr
# vVq/Gyxt3w==
# SIG # End signature block
