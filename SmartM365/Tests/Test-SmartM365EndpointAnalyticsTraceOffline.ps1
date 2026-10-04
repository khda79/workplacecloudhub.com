#Requires -Version 7.0
<#
.SYNOPSIS
Offline Endpoint Analytics logging, transcript and rejected-row evidence tests.
.VERSION
1.0.2
.NOTES
Loads production AST functions and the main try/finally only. Configuration, Core
actions, Graph calls, downloads, delays and CSV publication are simulated.
Real transcripts and synthetic diagnostics exist only in a disposable TEMP folder.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Production trace helpers consume the established Core global context; tests restore it afterwards.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='Synthetic console output verifies real transcript capture.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets','',Justification='Start-Sleep is mocked inside an isolated module; no test waits.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Mock actions only mutate synthetic in-memory lists or disposable TEMP fixtures.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Mock signatures match production helpers; unused parameters deliberately trigger no external actions.')]
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$path = Join-Path (Split-Path $PSScriptRoot -Parent) 'SmartInventory/M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Endpoint Analytics parse failed.' }
$definitions = @($ast.EndBlock.Statements |
    Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] } |
    ForEach-Object { $_.Extent.Text })
$main = @($ast.EndBlock.Statements |
    Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] })[-1].Extent.Text
$corePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Modules/SmartM365.Core/SmartM365.Core.psm1'
$coreAst = [Management.Automation.Language.Parser]::ParseFile($corePath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Core parse failed.' }
$coreDefinitions = foreach ($name in @('WriteLog','Format-SmartM365LogLine','RemoveOldFiles','Test-FileLocked')) {
    $node = $coreAst.Find({ param($item) $item -is [Management.Automation.Language.FunctionDefinitionAst] -and $item.Name -eq $name },$true)
    if (-not $node) { throw "Missing Core function: $name" }
    $node.Extent.Text -replace '^function WriteLog', 'function CoreWriteLog' -replace '^function RemoveOldFiles', 'function CoreRemoveOldFiles'
}
$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$fixtureRoot = Join-Path $tempRoot ('SmartM365-EATrace-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $fixtureRoot
$savedGlobals = @{}
foreach ($name in @('LogTextFile','logTranscriptFile','EnableSharePointUpload','RetentionMaxLogs','csvGeneratedPaths','SmartM365WarningCount','SmartM365ErrorCount')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $savedGlobals[$name] = @{Exists=($null -ne $variable);Value=$(if ($variable) { $variable.Value } else { $null })}
}
$module = $null
try {
    $module = New-Module -Name SyntheticEndpointTrace -ScriptBlock {
        param($Definitions,$Main,$FixtureRoot,$CoreDefinitions)
        Set-StrictMode -Version Latest
        foreach ($definition in $Definitions) { . ([scriptblock]::Create($definition)) }
        $script:Main = [scriptblock]::Create($Main)
        $script:Root = $FixtureRoot
        $script:Checks = 0
        function Assert-Test {
            param([bool]$Condition,[string]$Message)
            $script:Checks++
            if (-not $Condition) { throw $Message }
        }
        function Get-TestFailure {
            param([scriptblock]$Action)
            try { & $Action | Out-Null } catch { return $_ }
            throw 'Expected synthetic failure was not raised.'
        }
        function Reset-TestState {
            param([string]$Name)
            $script:RunId = $Name
            $script:OutputPath = Join-Path $script:Root $Name
            $null = New-Item -ItemType Directory -Path $script:OutputPath -Force
            $script:LatestCsvFolderPath = Join-Path $script:OutputPath 'canonical'
            $script:TenantConfig = [pscustomobject]@{TenantKey='synthetic-tenant';TenantId='synthetic-tenant';AppId='synthetic-app';Thumbprint='synthetic-cert'}
            $global:LogTextFile = Join-Path $script:OutputPath 'synthetic.log'
            $global:logTranscriptFile = Join-Path $script:OutputPath 'synthetic.transcript.txt'
            $global:EnableSharePointUpload = $false
            $global:RetentionMaxLogs = 5
            $script:GrainDiagnosticPaths = [Collections.Generic.List[string]]::new()
            $script:Logs = [Collections.Generic.List[string]]::new()
            $script:Uploads = [Collections.Generic.List[object]]::new()
            $script:Retention = [Collections.Generic.List[object]]::new()
            $script:CoreImported = $true
            $script:TranscriptStarted = $false
            $script:ScopeComplete = $null
            $script:CompletionStatus = 'Success'
            $script:CompletionError = $null
            $script:FailureStage = ''
            $script:OutputRows = [ordered]@{}
            $script:DataQualityRows = [Collections.Generic.List[object]]::new()
            $script:ScoreExclusions = [Collections.Generic.List[object]]::new()
            $script:Mails = [Collections.Generic.List[object]]::new()
            $script:ScopeQualifications = @()
            $script:ScriptName = 'SyntheticEndpointAnalytics'
            $script:ScriptVersion = 'synthetic-version'
            $script:CollectedAtUtc = [datetime]::UtcNow.ToString('o')
            $script:RequiredPermission = 'synthetic-permission'
            $script:AdvancedReportPattern = '^(BR|EAResourcePerf|EAAnomaly)|DeviceTimeline|DeviceQuery'
            $script:Reports = @('EADeviceScoresV2')
            $script:Tenant = 'synthetic'
            $script:IncludeStartupProcesses = $true
            $script:ValidateOnly = $false
            $script:Connect = $false
            $script:InteractiveAuth = $false
            $script:MaxItems = 0
            $script:ReportConsistencyAttempts = 3
            $script:ReportConsistencyRetryDelaySeconds = 15
            $script:Imports = 0
            $script:Jobs = 0
            $script:Delays = [Collections.Generic.List[int]]::new()
            $script:CanonicalWrites = 0
            $script:PublicationFails = $false
            $script:MailFails = $false
            $script:MockConnections = 0
            $script:PreflightFails = $false
            $script:CompletionFails = $false
            $script:TranscriptUploadDeferred = $false
            $script:UploadFails = $false
            $script:Exports = @([pscustomobject]@{Rows=@(
                [pscustomobject]@{DeviceId='private-device';EndpointAnalyticsScore=71;ExtraEvidence='first'},
                [pscustomobject]@{DeviceId='private-device';EndpointAnalyticsScore=88;ExtraEvidence='second'}
            )})
        }
        function CoreWriteLog {
            param($Message,$Level)
            $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
            $script:Logs.Add($line)
            Add-Content -LiteralPath $global:LogTextFile -Value $line
            Write-Host $line
        }
        function WriteLog { throw 'The old unprefixed logger must not be called.' }
        function Initialize-EARuntime { $script:CoreImported = $true }
        function Get-MgContext { [CmdletBinding()]param() return $null }
        function Disconnect-MgGraph { [CmdletBinding()]param() }
        function Connect-MgGraph {
            [CmdletBinding()]param($TenantId,$ClientId,$CertificateThumbprint,[switch]$NoWelcome)
            $script:MockConnections++
        }
        function Invoke-CoreSmartM365Preflight {
            param($ScriptName,$RequiredModules,$OutputPaths,$RequiredGraphApplicationPermissions)
            if ($script:PreflightFails) { throw 'Synthetic preflight failure.' }
        }
        function Start-EAExportJob {
            param($Report,$EffectiveReportName)
            if ($EffectiveReportName -notin @('EADeviceScoresV2','EAWFADeviceList')) { throw 'Unexpected alias/report.' }
            $script:Jobs++
            [pscustomobject]@{id=('synthetic-job-' + $script:Jobs)}
        }
        function Wait-EAExportJob {
            param($JobId,$ApiVersion)
            [pscustomobject]@{id=$JobId;status='completed';url='https://example.invalid/signed-download?secret=not-retained'}
        }
        function Import-EAExportedCsv {
            param($CompletedJob,$ReportName)
            $script:Imports++
            return @($script:Exports[[Math]::Min($script:Imports-1,$script:Exports.Count-1)].Rows)
        }
        function Start-Sleep { param($Seconds) $script:Delays.Add($Seconds) }
        function Publish-CoreSmartM365Csv {
            param($Data,$TimestampedPath,$LatestPath,$Columns)
            if ($script:PublicationFails) { throw 'Synthetic publication failure.' }
            $script:CanonicalWrites++
        }
        function Remove-CoreSmartM365TimestampedFilesOlderThan { param($FolderPath,$FilePattern,$RetentionDays,$LogFile) }
        function Send-CoreSmartM365TeamsNotification { param($Title,$Message,$Level,$Channel,$ResultSummary,$Facts) }
        function Set-CoreSmartM365CmdbSourceScope { param($CompleteScope,$Scope,$Qualifications) $script:ScopeComplete = $CompleteScope; $script:ScopeQualifications = @($Qualifications) }
        function Get-EAConfigValue { param($Name,$Default) if ($Name -eq 'ErrorMailTo') { 'reviewer@example.invalid' } elseif ($script:TenantConfig.PSObject.Properties[$Name]) { $script:TenantConfig.$Name } else { $Default } }
        function New-CoreSmartM365EmailBody {
            param($Title,$Category,$Severity,$Message,$SummaryData,$Sections,$PathRows,$Tenant)
            Assert-Test ($Severity -eq 'Warning' -and $SummaryData.Status -eq 'CompletedWithWarnings') 'Exclusion mail lost warning semantics.'
            $Message + ($Sections.Html -join '')
        }
        function CoreSendEmailHtmlReport {
            param($To,$Cc,$Subject,$BodyHtml,$MailPurpose)
            if ($script:MailFails) { throw 'Synthetic mail failure.' }
            $script:Mails.Add([pscustomobject]@{To=$To;Cc=$Cc;Subject=$Subject;Body=$BodyHtml;Purpose=$MailPurpose;Writes=$script:CanonicalWrites})
        }
        function Complete-CoreSmartM365ExecutionContext {
            param($Status,$ErrorRecord,$FailureStage,[switch]$DeferTranscriptUpload)
            $script:TranscriptUploadDeferred = [bool]$DeferTranscriptUpload
            Write-EALog ('SYNTHETIC FINAL BANNER: ' + $Status)
            if ($script:CompletionFails) { throw 'Synthetic completion failure.' }
        }
        function Invoke-CoreSmartM365SharePointCsvUpload {
            param($LocalFilePath,[switch]$EnsureParentFolders)
            $script:Uploads.Add([pscustomobject]@{Path=$LocalFilePath;Closed=(-not $script:TranscriptStarted);EnsureFolders=[bool]$EnsureParentFolders;Content=(Get-Content -LiteralPath $LocalFilePath -Raw)})
            if ($script:UploadFails) { throw 'Synthetic upload failure.' }
        }
        function CoreRemoveOldFiles {
            param($Path,$Filter,$KeepCount,$ExcludeFiles)
            $script:Retention.Add([pscustomobject]@{Path=$Path;Filter=$Filter;KeepCount=$KeepCount;ExcludeFiles=@($ExcludeFiles)})
        }

        function Invoke-TraceTest {
            Reset-TestState 'row-classification'
            $raw = @(
                [pscustomobject]@{DeviceId='conflict';EndpointAnalyticsScore=71;Extra='A';Nullable=$null},
                [pscustomobject]@{DeviceId=' CONFLICT ';EndpointAnalyticsScore=88;Extra='a'},
                [pscustomobject]@{DeviceId='identical';EndpointAnalyticsScore=80},
                [pscustomobject]@{DeviceId='identical';EndpointAnalyticsScore=80},
                [pscustomobject]@{DeviceId='unique';EndpointAnalyticsScore=90},
                [pscustomobject]@{DeviceId=' ';EndpointAnalyticsScore=0},
                $null
            )
            $before = ConvertTo-Json -InputObject $raw -Depth 20
            $d = Get-EARejectedRowDiagnostic $raw EADeviceScoresV2 synthetic-job 1
            Assert-Test ($d.TotalRawRows -eq 7 -and $d.DuplicateExcessRows -eq 2 -and $d.InvalidIdentityRows -eq 2) 'Rejected-row counts changed.'
            Assert-Test (-not $d.CanonicalPublicationAllowed -and $d.DuplicateGroups.Count -eq 2) 'Diagnostic incorrectly qualified canonical publication.'
            $conflict = @($d.DuplicateGroups | Where-Object DeviceId -eq conflict)[0]
            $identical = @($d.DuplicateGroups | Where-Object DeviceId -eq identical)[0]
            Assert-Test ($conflict.Classification -eq 'DifferentRawRows' -and $identical.Classification -eq 'IdenticalRawRows') 'Identical/conflicting raw rows were confused.'
            Assert-Test ($conflict.DifferingColumns -contains 'Extra' -and $conflict.DifferingColumns -contains 'Nullable' -and $conflict.DifferingColumns -contains 'EndpointAnalyticsScore') 'Raw differences, case or missing-versus-null evidence were lost.'
            Assert-Test (($conflict.RowNumbers -join ',') -eq '1,2' -and ($d.InvalidRows.RowNumber -join ',') -eq '6,7') 'Original row locations were lost.'
            Assert-Test ((ConvertTo-Json -InputObject $raw -Depth 20) -ceq $before) 'Diagnostic generation mutated source rows.'
            Assert-Test ($conflict.RawRows[0].Extra -ceq 'A' -and $conflict.RawRows[1].Extra -ceq 'a') 'Unnormalized raw evidence was changed.'
            Assert-Test ((ConvertTo-Json $d -Depth 30) -notmatch 'unique') 'Unique valid rows unnecessarily entered the diagnostic.'
            $empty = Get-EARejectedRowDiagnostic @() EADeviceScoresV2 synthetic-job 1
            Assert-Test ($empty.DuplicateGroups.Count -eq 0 -and $empty.InvalidRows.Count -eq 0) 'Empty diagnostic handling failed.'
            Save-EARejectedRowDiagnostic $raw EADeviceScoresV2 synthetic-job 1
            Save-EARejectedRowDiagnostic $raw EADeviceScoresV2 synthetic-job 2
            Assert-Test ($script:GrainDiagnosticPaths.Count -eq 2 -and $script:GrainDiagnosticPaths[0] -cne $script:GrainDiagnosticPaths[1]) 'Attempts overwrote one another.'
            $text = Get-Content -LiteralPath $script:GrainDiagnosticPaths[0] -Raw
            $saved = $text | ConvertFrom-Json
            Assert-Test ($saved.RunId -eq 'row-classification' -and $saved.TenantKey -eq 'synthetic-tenant' -and $saved.ExportJobId -eq 'synthetic-job' -and $saved.Attempt -eq 1) 'Diagnostic provenance is incomplete.'
            Assert-Test ($text -notmatch 'signed-download|not-retained' -and (Split-Path $script:GrainDiagnosticPaths[0] -Leaf) -like '*.json.txt') 'Diagnostic transport leaked a URL or used the wrong extension.'
            Assert-Test ((Split-Path $script:GrainDiagnosticPaths[0] -Parent) -eq (Join-Path $script:OutputPath Diagnostics)) 'Diagnostic escaped its private script log folder.'
            Assert-Test (($script:Logs -join ' ') -notmatch 'conflict|identical|\b71\b|\b88\b') 'Count-only trace logs leaked raw device evidence.'
            Complete-EATraceArtifacts
            Assert-Test ($script:Uploads.Count -eq 0) 'Disabled upload triggered an external action.'
            Assert-Test ($script:Retention.Count -eq 1 -and $script:Retention[0].KeepCount -eq 3 -and $script:Retention[0].ExcludeFiles.Count -eq 2 -and $script:Retention[0].Filter -eq 'Intune_EndpointAnalytics_RejectedRows_*.json.txt') 'Retention lost the current-run exclusion or scoped filter.'
            $global:RetentionMaxLogs = 0
            Complete-EATraceArtifacts
            Assert-Test ($script:Retention.Count -eq 1) 'Disabled retention was not respected.'
            $failure = Get-TestFailure { Save-EARejectedRowDiagnostic $raw '../unsafe' synthetic-job 1 }
            Assert-Test ($failure.Exception.Message -match 'Unsafe') 'Unsafe diagnostic file identity was accepted.'

            Reset-TestState 'persistent-rejection'
            $script:Reports = @('EAWFADeviceList')
            $global:EnableSharePointUpload = $true
            $failure = Get-TestFailure { & $script:Main }
            Assert-Test ($failure.Exception.Message -match 'Canonical business CSV files were not published') 'Collector failure boundary changed.'
            Assert-Test ($failure.Exception.Message -match 'Failed reports: EAWFADeviceList \[ExportFailed\]:' -and $failure.Exception.Message -match 'duplicate device rows=1' -and $failure.Exception.Message -notmatch 'private-device|ExtraEvidence|\b71\b|\b88\b') 'Final failure omitted the report/root cause or exposed raw row values.'
            Assert-Test ($script:TranscriptUploadDeferred) 'Core attempted to upload the active transcript before the final banner.'
            Assert-Test ($script:Imports -eq 3 -and $script:Jobs -eq 3 -and ($script:Delays -join ',') -eq '15,30') 'Fresh complete exports or bounded retries changed.'
            Assert-Test ($script:CanonicalWrites -eq 0 -and $script:ScopeComplete -eq $false -and $script:CompletionStatus -eq 'Failed') 'Failed collection was admitted or published.'
            Assert-Test ($script:GrainDiagnosticPaths.Count -eq 3) 'An exhausted export attempt lacks private evidence.'
            Assert-Test (-not $script:TranscriptStarted -and (Test-Path -LiteralPath $global:logTranscriptFile)) 'Failure did not close its transcript.'
            $transcript = Get-Content -LiteralPath $global:logTranscriptFile -Raw
            # The end marker is localized by PowerShell; use its timestamp/footer shape.
            Assert-Test ($transcript -match 'Starting SyntheticEndpointAnalytics vsynthetic-version' -and $transcript -match 'SYNTHETIC FINAL BANNER: Failed' -and $transcript -match '\d{14}\r?\n\*+\s*$') 'Transcript omitted the startup/version, final banner or closing marker.'
            $log = Get-Content -LiteralPath $global:LogTextFile -Raw
            Assert-Test ($log -match 'duplicate device rows=1' -and $log -notmatch 'private-device|ExtraEvidence|\b71\b|\b88\b') 'Prefixed logging lost the failure or leaked raw values.'
            Assert-Test ($script:Uploads.Count -eq 4 -and @($script:Uploads | Where-Object { -not $_.Closed }).Count -eq 0) 'Trace upload happened before transcript closure or omitted an artifact.'
            Assert-Test (@($script:Uploads | Where-Object { -not $_.EnsureFolders }).Count -eq 0) 'New diagnostic folders would not be created on SharePoint.'
            Assert-Test ($script:Uploads[-1].Content -match 'SYNTHETIC FINAL BANNER: Failed' -and $script:Uploads[-1].Content -match '\d{14}\r?\n\*+\s*$') 'Uploaded transcript was incomplete.'
            Assert-Test ($script:MockConnections -eq 1) 'Unexpected mock connection count.'

            Reset-TestState 'recover-clean-export'
            $script:Exports += [pscustomobject]@{Rows=@([pscustomobject]@{DeviceId='clean';EndpointAnalyticsScore=90})}
            & $script:Main | Out-Null
            Assert-Test ($script:Imports -eq 2 -and $script:CanonicalWrites -eq 9 -and $script:CompletionStatus -eq 'Success') 'Clean fresh retry was not published normally.'
            Assert-Test ($script:GrainDiagnosticPaths.Count -eq 1 -and -not $script:TranscriptStarted) 'Clean retry lost rejected evidence or transcript closure.'
            Assert-Test ($script:OutputRows.DevicePerformance.Count -eq 1 -and $script:OutputRows.DevicePerformance[0].DeviceId -eq 'clean') 'Rejected rows contaminated the clean output.'
            Assert-Test ($script:Uploads.Count -eq 0) 'Unconfigured external action ran after success.'
            Assert-Test ($script:Mails.Count -eq 0) 'Clean retry sent an exclusion alert.'

            Reset-TestState 'exclude-ambiguous-scores'
            $script:Exports[0].Rows += [pscustomobject]@{DeviceId='clean';EndpointAnalyticsScore=90}
            & $script:Main | Out-Null
            Assert-Test ($script:Imports -eq 3 -and $script:CanonicalWrites -eq 9 -and $script:CompletionStatus -eq 'CompletedWithWarnings') 'Qualified exclusion did not publish fresh valid outputs with warning status.'
            Assert-Test ($script:OutputRows.DevicePerformance.Count -eq 1 -and $script:OutputRows.DevicePerformance[0].DeviceId -eq 'clean' -and $script:OutputRows.DevicePerformance[0].EndpointAnalyticsScore -eq 90) 'Exclusion chose a conflicting row or changed valid scores.'
            $quality = @($script:DataQualityRows)[0]
            Assert-Test ($quality.Status -eq 'CollectedWithExclusions' -and $quality.RawRowCount -eq 3 -and $quality.RowCount -eq 1 -and $quality.ExcludedRowCount -eq 2 -and $quality.ExcludedDeviceCount -eq 1) 'Exclusion accounting was not published in DataQuality.'
            Assert-Test ($script:Mails.Count -eq 1 -and $script:Mails[0].Writes -eq 9 -and $script:Mails[0].Purpose -eq 'Error' -and $script:Mails[0].Cc -eq '' -and $script:Mails[0].To -eq 'reviewer@example.invalid') 'Exclusion alert was not sent after publication via the Core error-mail route.'
            Assert-Test ($script:Mails[0].Body -match 'private-device' -and $script:Mails[0].Body -match '71' -and $script:Mails[0].Body -match '88' -and $script:Mails[0].Body -notmatch 'ExtraEvidence') 'Exclusion alert omitted the device/conflicting scores or leaked unnecessary source fields.'
            Assert-Test ($script:GrainDiagnosticPaths.Count -eq 3 -and -not $script:TranscriptStarted) 'Qualified exclusion lost raw evidence or transcript closure.'
            Assert-Test (($script:ScopeQualifications -join '') -match 'excludedRows=2; excludedDevices=1' -and ($script:ScopeQualifications -join '') -notmatch 'private-device|\b71\b|\b88\b') 'Producer qualification omitted exclusions or exposed private row values.'
            $performanceReport = @(Get-EAReportCatalog | Where-Object Name -eq EADevicePerformanceV2)[0]
            $cleanQuality = New-EADataQualityRow -ReportName $performanceReport.Name -ApiVersion beta -Status Collected -RowCount 1
            $scope = Get-EACollectionScope -QualityRows @($quality,$cleanQuality) -ItemLimit 0
            Assert-Test ($scope.CompleteScope -and $scope.Qualifications.Count -eq 1) 'Complete acquired scope with explicit exclusions was rejected.'
            Assert-Test (-not (Get-EACollectionScope -QualityRows @($quality,$cleanQuality) -ItemLimit 1).CompleteScope) 'Limited acquisition was admitted as complete.'
            Assert-Test (-not (Get-EACollectionScope -QualityRows @($quality) -ItemLimit 0).CompleteScope) 'Missing required performance report was admitted as complete.'
            $schemas = Get-EAOutputSchemas
            Assert-Test ($schemas.DataQuality -contains 'RawRowCount' -and $schemas.DataQuality -contains 'ExcludedRowCount' -and $schemas.DataQuality -contains 'ExcludedDeviceCount') 'CSV schema dropped exclusion counts.'

            Reset-TestState 'all-scores-excluded'
            $script:Exports[0].Rows[0] | Add-Member -NotePropertyName DeviceName -NotePropertyValue '<script>alert("test")</script>'
            & $script:Main | Out-Null
            Assert-Test ($script:OutputRows.DevicePerformance.Count -eq 0 -and $script:CanonicalWrites -eq 9 -and $script:Mails.Count -eq 1) 'All-score exclusion failed stable empty-output publication.'
            Assert-Test ($script:Mails[0].Body -match '&lt;script&gt;' -and $script:Mails[0].Body -notmatch '<script>') 'Raw device text was not HTML escaped in the alert.'

            Reset-TestState 'publication-fails-before-alert'
            $script:PublicationFails=$true
            $failure=Get-TestFailure { & $script:Main }
            Assert-Test ($failure.Exception.Message -eq 'Synthetic publication failure.' -and $script:Mails.Count -eq 0 -and $script:CompletionStatus -eq 'Failed') 'Alert claimed successful publication before valid outputs existed.'

            Reset-TestState 'mail-fails-after-publication'
            $script:MailFails=$true
            $failure=Get-TestFailure { & $script:Main }
            Assert-Test ($failure.Exception.Message -eq 'Synthetic mail failure.' -and $script:CanonicalWrites -eq 9 -and $script:CompletionStatus -eq 'Failed' -and -not $script:TranscriptStarted) 'Mail failure was hidden, removed valid CSVs or left the transcript open.'

            Reset-TestState 'diagnostic-write-failure'
            # Save failure is injected without changing the production consistency code.
            function Save-EARejectedRowDiagnostic { param($RawRows,$ReportName,$ExportJobId,$Attempt) throw 'Synthetic diagnostic failure.' }
            $global:LogTextFile = Join-Path $script:OutputPath 'synthetic.log'
            $report = @(Get-EAReportCatalog | Where-Object Name -eq EADeviceScoresV2)[0]
            $failure = Get-TestFailure { Invoke-EAReport $report }
            Assert-Test ([bool]$failure.Exception.Data['EndpointAnalyticsGrain'] -and $failure.Exception.Message -notmatch 'Synthetic diagnostic failure') 'Diagnostic failure masked the blocking grain failure.'
            Assert-Test ($script:Jobs -eq 3 -and $script:CanonicalWrites -eq 0 -and @($script:Logs | Where-Object { $_ -match 'original grain failure remains blocking' }).Count -eq 3) 'Diagnostic failure bypassed the retry or publication guard.'
            # Restore the actual production saver for subsequent lifecycle tests.
            $saver = @($Definitions | Where-Object { $_ -match '^function Save-EARejectedRowDiagnostic' })[0]
            . ([scriptblock]::Create($saver))

            Reset-TestState 'preflight-failure'
            $script:PreflightFails = $true
            $script:CompletionFails = $true
            $failure = Get-TestFailure { & $script:Main }
            Assert-Test ($failure.Exception.Message -eq 'Synthetic preflight failure.' -and $script:Jobs -eq 0 -and $script:CanonicalWrites -eq 0) 'Completion masked preflight failure or started collection.'
            Assert-Test (-not $script:TranscriptStarted -and (Get-Content -LiteralPath $global:logTranscriptFile -Raw) -match 'SYNTHETIC FINAL BANNER: Failed') 'Completion failure left the transcript open.'

            Reset-TestState 'static-validation'
            $script:ValidateOnly = $true
            & $script:Main | Out-Null
            Assert-Test ($script:Jobs -eq 0 -and $script:MockConnections -eq 0 -and $script:CanonicalWrites -eq 0 -and $script:Uploads.Count -eq 0 -and -not $script:TranscriptStarted) 'Static validation ran an external action or left its transcript open.'

            Reset-TestState 'upload-failure'
            Start-EATranscript
            Write-EALog 'Synthetic upload-failure transcript.'
            $global:EnableSharePointUpload = $true
            $script:UploadFails = $true
            Complete-EATraceArtifacts
            Assert-Test (-not $script:TranscriptStarted -and (Test-Path -LiteralPath $global:logTranscriptFile) -and $script:Uploads.Count -eq 1) 'Upload failure lost local transcript evidence.'

            Reset-TestState 'core-trace-integration'
            foreach ($definition in $CoreDefinitions) { . ([scriptblock]::Create($definition)) }
            function Invoke-SmartM365TeamsNotificationFromLog { param($Message,$Level) }
            $global:csvGeneratedPaths = @()
            $global:SmartM365WarningCount = 0
            $global:SmartM365ErrorCount = 0
            Write-EALog "Synthetic first line`nSynthetic second line" WARNING
            $lines = @(Get-Content -LiteralPath $global:LogTextFile)
            Assert-Test ($lines.Count -eq 2 -and @($lines | Where-Object { $_ -notmatch '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[WARNING\] ' }).Count -eq 0 -and $global:SmartM365WarningCount -eq 1) 'Actual prefixed Core logger failed timestamps, physical lines or warning counters.'
            $diagnosticFolder = Join-Path $script:OutputPath Diagnostics
            $null = New-Item -ItemType Directory -Path $diagnosticFolder
            $unrelated = Join-Path $diagnosticFolder 'unrelated.json.txt'
            Set-Content -LiteralPath $unrelated -Value '{}'
            foreach ($index in 1..7) {
                $old = Join-Path $diagnosticFolder ("Intune_EndpointAnalytics_RejectedRows_synthetic_old_$index.json.txt")
                Set-Content -LiteralPath $old -Value '{}'
                (Get-Item -LiteralPath $old).LastWriteTime = (Get-Date).AddDays(-$index)
            }
            Save-EARejectedRowDiagnostic $raw EADeviceScoresV2 synthetic-job 1
            Save-EARejectedRowDiagnostic $raw EADeviceScoresV2 synthetic-job 2
            Complete-EATraceArtifacts
            $remaining = @(Get-ChildItem -LiteralPath $diagnosticFolder -Filter 'Intune_EndpointAnalytics_RejectedRows_*.json.txt')
            Assert-Test ($remaining.Count -eq 5 -and (Test-Path -LiteralPath $unrelated) -and @($script:GrainDiagnosticPaths | Where-Object { -not (Test-Path -LiteralPath $_) }).Count -eq 0) 'Actual shared retention did not preserve current evidence, bound old diagnostics or protect unrelated files.'
            Assert-Test (@($remaining | Where-Object Name -like '*old_*' | Sort-Object Name | Select-Object -ExpandProperty Name) -join ',' -eq 'Intune_EndpointAnalytics_RejectedRows_synthetic_old_1.json.txt,Intune_EndpointAnalytics_RejectedRows_synthetic_old_2.json.txt,Intune_EndpointAnalytics_RejectedRows_synthetic_old_3.json.txt') 'Actual shared retention kept stale rather than most recent diagnostics.'
            [pscustomobject]@{Status='Passed';Checks=$script:Checks;LiveCalls=0;RealDelays=0;ProductionWrites=0}
        }
    } -ArgumentList $definitions,$main,$fixtureRoot,$coreDefinitions
    & $module { Invoke-TraceTest }
}
finally {
    if ($module) {
        & $module { if ($script:TranscriptStarted) { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } }
        Remove-Module $module -Force -ErrorAction SilentlyContinue
    }
    foreach ($name in $savedGlobals.Keys) {
        if ($savedGlobals[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $savedGlobals[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    $resolved = [IO.Path]::GetFullPath($fixtureRoot)
    if (-not $resolved.StartsWith($tempRoot,[StringComparison]::OrdinalIgnoreCase) -or
        (Split-Path $resolved -Leaf) -notmatch '^SmartM365-EATrace-[a-f0-9]{32}$') {
        throw 'Unsafe synthetic fixture cleanup target.'
    }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDHwYk2p/UbqThu
# Acjtyf953T3H6Ys9PnYZqRKNEELrFKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID9SfsjNYnwPaiR8FmY59TOOE5MDx3+rjFtnlVyr71XlMA0GCSqG
# SIb3DQEBAQUABIIBgGL+HYAEIDdphlFL+DYtu0OpJgEewkblykj/zsbZ28RHzjfw
# iMwoR0MaaZnlz3EDW5aUqXlWVw6OrUwU1UxckSRCZYCgrgtE+BEoaVRCqnRoWO4E
# FQ8pucpGR3861FP3IhjaWAeSO9zMDtm3aVDnnkI2VCxfVsRQOCuGxK9Jxof7JjBv
# Q6sxpMeGPiZlpCVQAqptjJXa6v96IQ2KAaksMT6aWW7a68SCShHfZZu9b/MbHNGf
# gDHKGuPxrea2k9EUWFiITU4TMatxhfFh2ENraNCOKi56qIsy+jrT+krloj6AjA90
# CIeatMB2gWHDRFjRi5Racw1C+3n+RaBmsuOeIMeAydlzhsOmwDGJ/VcYFe6LoL9h
# 0Pc16yWGB6PMcDSrjTojcxsQgMZjsD4eCfGv2cdYVWz1r2c5Q7wJ41nu5eKfUWQj
# XKx5X0WSuyvCnQH5eR2SXDDrRSQhbms/5bgNOVBQtJ2Lm0yyGiajOu838wvVsOjC
# vTHRQ3ZU5/M9QpBwKqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDQyMTAx
# NTZaMC8GCSqGSIb3DQEJBDEiBCCiWCUn6HahPGkTvSeYlHak3xTYcwxlrL5PvkFh
# TU/3mjANBgkqhkiG9w0BAQEFAASCAgCdyQRWtE45NlGpzVYejY8YEhO7LBvWmW0Q
# s/rKtDJAMGnE3ACZv40uaFL7xpBbGpiUVjzPxvCZa6Q1/9NiSGYpCM/aIZJRdI/x
# 2u3yZWNllgbmLMqwMMBwGeidTjdVr/x0pMtx/P+/ipvEOjVYe/gcEqIiXcCGUyyK
# NTJibgdI/EcEKxH9fmBNUCmKqLt6UPYbgDrHIoFXN7vfmMQJnXJCCj1xv7EysEXw
# LQfyFo1mvf+uFZv2/eDMnfG7EvVYQ4C6opWqqgurFtdrBOuIsI8MrpGNc5QwXm+o
# S7rOB+31kC4VEW+IxOjkUUf7RRZEOaQlorHMFDpYp8vDXysNaec5uzkan/gC8WfF
# /v+J35gFyyQRN+8Q+7eWPydCKcGnUmad/WE4SjnJsw2XP2wwytslT/QdeMgyZYEZ
# aoZtYIu00hAhbaYEVZSTh0Oo/zorz1vpdcrgjxaBC9myfkQbHjuWohSJSXHn8ETt
# M/9ZWoSM6Oa7PLLduX0/MYIxThr5k7baHEEpHXxf+X5K/EWbcTUpZxAifxX+7KPH
# 2iq6fQZbt1JeKCYhXWLPR99ozfZhLx/Seateu2mbVgac46rHXsylu3lJDo6GThbE
# jHZGS/RqS2AzrGdcEOoL4r0+XE6mK9qewD7aitDuXdnUSfhfFZDRyCAT9DmdbjvW
# tHofh2+KIw==
# SIG # End signature block
