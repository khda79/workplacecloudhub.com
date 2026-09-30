<#
.SYNOPSIS
Runs offline tests for the orchestrator SharePoint operational-folder mirror.
.VERSION
1.0.3
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) { throw $Message }
}

$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$orchestratorPath = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator\SmartM365-Inventory-Orchestrator.ps1'
$coreModulePath = Join-Path -Path $smartM365Root -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) { throw 'Orchestrator source failed parsing.' }

$functionNames = @(
    'Get-OrchestratorSharePointMirrorRelativePath',
    'Test-OrchestratorSharePointMirrorFile',
    'Get-OrchestratorSharePointMirrorSnapshot',
    'Read-OrchestratorSharePointMirrorState',
    'Save-OrchestratorSharePointMirrorState',
    'Enter-OrchestratorSharePointMirrorLock',
    'Exit-OrchestratorSharePointMirrorLock',
    'Invoke-OrchestratorEnsureSharePointFolder',
    'Invoke-OrchestratorSharePointMirror',
    'Invoke-OrchestratorPeriodicSharePointUpload',
    'Test-OrchestratorPlannedRecycle'
)
$definitions = foreach ($name in $functionNames) {
    $node = $ast.Find({ param($candidate) $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq $name }, $true)
    if (-not $node) { throw "Missing orchestrator function: $name" }
    $node.Extent.Text
}
# The snapshot refreshes the shared heartbeat during long scans; not under test here.
$definitions = @($definitions) + 'function Update-OrchestratorHeartbeatDuringLongOperation { }'
$mirrorModule =New-Module -ScriptBlock ([scriptblock]::Create($definitions -join [Environment]::NewLine))

$temporaryRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('SmartM365-OrchestratorSharePointMirror-' + [guid]::NewGuid().ToString('N'))
try {
    $sharedRoot = Join-Path -Path $temporaryRoot -ChildPath 'Tenant\DATA-ALL\Orchestrator'
    $paths = @(
        'Config\Versions\v1',
        'Audit',
        'Election\Claims\SyntheticJob',
        'Election\Concurrency',
        'PipelineRuns\batch-1\Jobs'
    )
    foreach ($relativePath in $paths) { New-Item -ItemType Directory -Path (Join-Path $sharedRoot $relativePath) -Force | Out-Null }

    Set-Content -LiteralPath (Join-Path $sharedRoot 'Config\Orchestrator-Jobs.json') -Value '{"Jobs":[]}' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $sharedRoot 'Config\Versions\v1\Orchestrator-Jobs.json') -Value '{"Jobs":[]}' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $sharedRoot 'Audit\Orchestrator_ConfigChanges.csv') -Value 'VersionId' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $sharedRoot 'Election\Orchestrator-ElectionPlan.json') -Value '{"PlanId":"p1"}' -Encoding utf8
    $leasePath = Join-Path $sharedRoot 'Election\Concurrency\SharedRuntime.json'
    Set-Content -LiteralPath $leasePath -Value '{"LeaseId":"l1"}' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $sharedRoot 'PipelineRuns\batch-1\request.json') -Value '{"BatchId":"batch-1"}' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $sharedRoot 'PipelineRuns\batch-1\Jobs\SyntheticJob.json') -Value '{"Status":"Pending"}' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $sharedRoot 'Election\Concurrency\transient.lock') -Value 'lock' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $sharedRoot 'PipelineRuns\batch-1\request.json.tmp') -Value 'partial' -Encoding utf8
    Set-Content -LiteralPath (Join-Path $sharedRoot 'Audit\notes.txt') -Value 'not managed' -Encoding utf8

    $snapshot = & $mirrorModule { param($Root) Get-OrchestratorSharePointMirrorSnapshot -SharedDataFolderPath $Root } $sharedRoot
    $relativeFiles = @($snapshot.Files | ForEach-Object { [string]$_.RelativePath })
    Assert-True -Condition ($relativeFiles.Count -eq 7) -Message "Unexpected mirrored file count: $($relativeFiles.Count)."
    Assert-True -Condition ('DATA-ALL/Orchestrator/Config/Orchestrator-Jobs.json' -in $relativeFiles) -Message 'Current configuration was not selected.'
    Assert-True -Condition ('DATA-ALL/Orchestrator/PipelineRuns/batch-1/Jobs/SyntheticJob.json' -in $relativeFiles) -Message 'Pipeline status was not selected.'
    Assert-True -Condition (-not ($relativeFiles -match '(?i)\.lock$|\.tmp$|notes\.txt$')) -Message 'Transient or unsupported files were selected.'
    Assert-True -Condition (@($snapshot.Folders | Where-Object RelativePath -eq 'DATA-ALL/Orchestrator/Config').Count -eq 1) -Message 'Config folder was not selected.'
    Assert-True -Condition (@($snapshot.Folders | Where-Object RelativePath -eq 'DATA-ALL/Orchestrator/PipelineRuns').Count -eq 1) -Message 'PipelineRuns folder was not selected.'

    $initializeProbe = & $mirrorModule {
        $script:Settings = [pscustomobject]@{
            SharePointSiteHostname = 'tenant.sharepoint.test'
            SharePointSitePath = '/sites/SMART-M365'
            SharePointLibraryDisplayName = 'Documents'
            SharePointTargetFolderPath = 'SMART-M365'
        }
        $script:SharePointEnsuredFolderState = @{}
        $script:InitializedFolders = [Collections.Generic.List[string]]::new()
        function script:Initialize-SmartM365SharePointFolder {
            [CmdletBinding()]
            param(
                [string]$SharePointRelativeFolderPath,
                [bool]$Enabled,
                [string]$SiteHostname,
                [string]$SitePath,
                [string]$LibraryDisplayName,
                [string]$TargetFolderPath
            )
            $script:InitializedFolders.Add($SharePointRelativeFolderPath)
            return $true
        }
        $result = Invoke-OrchestratorEnsureSharePointFolder -RelativePath 'DATA-ALL/Orchestrator/Config'
        [pscustomobject]@{ Result = $result; Calls = @($script:InitializedFolders.ToArray()) }
    }
    Assert-True -Condition ($initializeProbe.Result -and $initializeProbe.Calls.Count -eq 1 -and $initializeProbe.Calls[0] -ceq 'DATA-ALL/Orchestrator/Config') -Message 'The orchestrator did not call Initialize-SmartM365SharePointFolder exactly once.'

    & $mirrorModule {
        param($Root)
        $script:Settings = [pscustomobject]@{
            SharedDataFolderPath = $Root
            SharePointMirrorStatePath = (Join-Path $Root 'Orchestrator-SharePointMirrorState.json')
            SharePointMirrorLockPath = (Join-Path $Root 'Orchestrator-SharePointMirror.lock')
            OrchestratorSharePointUploadIntervalMinutes = 60
        }
        $script:SharePointEnsuredFolderState = @{}
        $script:Uploads = [Collections.Generic.List[string]]::new()
        $script:Deletes = [Collections.Generic.List[string]]::new()
        $script:EnsuredFolders = [Collections.Generic.List[string]]::new()
        $script:Logs = [Collections.Generic.List[string]]::new()
        $script:ThrowSnapshot = $false
        $script:FailUploadPath = ''
        $script:OriginalSnapshot = (Get-Command Get-OrchestratorSharePointMirrorSnapshot -CommandType Function).ScriptBlock

        function script:Get-OrchestratorSharePointMirrorSnapshot {
            param([string]$SharedDataFolderPath)
            if ($script:ThrowSnapshot) { throw 'Synthetic incomplete scan.' }
            & $script:OriginalSnapshot -SharedDataFolderPath $SharedDataFolderPath
        }
        function script:Test-OrchestratorSharePointUploadConfigured { return $true }
        function script:Invoke-OrchestratorEnsureSharePointFolder {
            param([string]$RelativePath)
            $script:EnsuredFolders.Add($RelativePath)
            return $true
        }
        function script:Invoke-OrchestratorSharePointUpload {
            param([string]$LocalFilePath, [string]$Reason, [switch]$Force)
            $script:Uploads.Add($LocalFilePath)
            if ($LocalFilePath -eq $script:FailUploadPath) { return $false }
            return $true
        }
        function script:Invoke-OrchestratorSharePointDelete {
            param([string]$LocalFilePath, [string]$Reason)
            $script:Deletes.Add($LocalFilePath)
            return $true
        }
        function script:Write-OrchestratorLog { param([string]$Message, [string]$Level) $script:Logs.Add($Message) }
        function script:Write-FileAtomically {
            param([string]$Path, [string]$Content)
            [IO.File]::WriteAllText($Path, $Content, [Text.UTF8Encoding]::new($false))
        }
    } $sharedRoot

    & $mirrorModule { Invoke-OrchestratorSharePointMirror }
    $firstUploadCount = & $mirrorModule { $script:Uploads.Count }
    Assert-True -Condition ($firstUploadCount -eq 7) -Message "Initial mirror uploaded $firstUploadCount files instead of 7."
    Assert-True -Condition ((& $mirrorModule { $script:EnsuredFolders.Count }) -ge 4) -Message 'Operational folders were not ensured.'

    & $mirrorModule { Invoke-OrchestratorSharePointMirror }
    Assert-True -Condition ((& $mirrorModule { $script:Uploads.Count }) -eq $firstUploadCount) -Message 'Unchanged files were uploaded again.'

    Set-Content -LiteralPath (Join-Path $sharedRoot 'Config\Orchestrator-Jobs.json') -Value '{"Jobs":[{"Name":"Changed"}]}' -Encoding utf8
    & $mirrorModule { Invoke-OrchestratorSharePointMirror }
    Assert-True -Condition ((& $mirrorModule { $script:Uploads.Count }) -eq ($firstUploadCount + 1)) -Message 'Changed configuration was not uploaded exactly once.'

    Copy-Item -LiteralPath $leasePath -Destination ($leasePath+'.txt')
    $pairedSnapshot=& $mirrorModule {param($Root)Get-OrchestratorSharePointMirrorSnapshot -SharedDataFolderPath $Root} $sharedRoot
    Assert-True -Condition (@($pairedSnapshot.Files | Where-Object RelativePath -like '*/Concurrency/SharedRuntime*').Count -eq 1) -Message 'Identical JSON pair mirrored twice.'
    & $mirrorModule { Invoke-OrchestratorSharePointMirror }
    Assert-True -Condition ((& $mirrorModule {$script:Deletes.Count}) -eq 0) -Message 'Renamed lease triggered remote deletion.'
    Remove-Item -LiteralPath $leasePath -Force
    $leasePath += '.txt'
    Remove-Item -LiteralPath $leasePath -Force
    & $mirrorModule { Invoke-OrchestratorSharePointMirror }
    Assert-True -Condition ((& $mirrorModule { $script:Deletes.Count }) -eq 1) -Message 'Expired concurrency lease was not removed from SharePoint.'

    Set-Content -LiteralPath $leasePath -Value '{"LeaseId":"l2"}' -Encoding utf8
    & $mirrorModule { Invoke-OrchestratorSharePointMirror }
    Remove-Item -LiteralPath $leasePath -Force
    $statePath = Join-Path $sharedRoot 'Orchestrator-SharePointMirrorState.json'
    $stateHashBeforeFailure = (Get-FileHash -LiteralPath $statePath -Algorithm SHA256).Hash
    $deleteCountBeforeFailure = & $mirrorModule { $script:Deletes.Count }
    & $mirrorModule { $script:ThrowSnapshot = $true; Invoke-OrchestratorSharePointMirror; $script:ThrowSnapshot = $false }
    Assert-True -Condition ((& $mirrorModule { $script:Deletes.Count }) -eq $deleteCountBeforeFailure) -Message 'Incomplete scan triggered a remote deletion.'
    Assert-True -Condition ((Get-FileHash -LiteralPath $statePath -Algorithm SHA256).Hash -eq $stateHashBeforeFailure) -Message 'Incomplete scan changed the valid mirror state.'

    # Execute the actual finalization call, with only remote operations mocked.
    $finalCall = $ast.FindAll({ param($node)
        $node -is [Management.Automation.Language.CommandAst] -and
        $node.GetCommandName() -eq 'Invoke-OrchestratorPeriodicSharePointUpload'
    }, $true) | Where-Object { $_.Extent.Text -match '\(Get-Date\)' } | Select-Object -Last 1
    Assert-True -Condition ($null -ne $finalCall) -Message 'Final synchronization call missing.'
    $finalScript = [scriptblock]::Create($finalCall.Extent.Text)
    & $mirrorModule {
        $script:Settings | Add-Member NoteProperty OrchestratorRunsCsvPath ''
        $script:Settings | Add-Member NoteProperty StatePath ''
        $script:Settings | Add-Member NoteProperty HeartbeatPath ''
        $script:LastSharePointUploadAttempt = Get-Date
        # A final synchronization outside a planned recycle mirrors the operational folders.
        $script:ExitCode = 0
        $script:OrchestratorStopReason = 'Stopped'
        function script:Get-OrchestratorLogPath { param($Date) return '' }
        function script:Get-JobRunsCsvPath { return '' }
        $script:Uploads.Clear()
    }
    $changedPath = Join-Path $sharedRoot 'Config\Orchestrator-Jobs.json'
    $newPath = Join-Path $sharedRoot 'Audit\new.json.txt'
    $retryPath = Join-Path $sharedRoot 'Audit\retry.json.txt'
    Set-Content -LiteralPath $retryPath -Value '{"Retry":true}' -Encoding utf8
    & $mirrorModule { param($Path) $script:FailUploadPath = $Path; Invoke-OrchestratorSharePointMirror; $script:Uploads.Clear() } $retryPath
    Set-Content -LiteralPath $changedPath -Value '{"Jobs":[{"Name":"FinalChanged"}]}' -Encoding utf8
    Set-Content -LiteralPath $newPath -Value '{"New":true}' -Encoding utf8
    & $mirrorModule { Invoke-OrchestratorPeriodicSharePointUpload -Now (Get-Date) }
    Assert-True -Condition ((& $mirrorModule { $script:Uploads.Count }) -eq 0) -Message 'Normal interval gate was bypassed.'
    & $mirrorModule { param($Call) $script:FailUploadPath = ''; & $Call } $finalScript
    $finalUploads = @(& $mirrorModule { $script:Uploads.ToArray() })
    Assert-True -Condition ($finalUploads.Count -eq 3) -Message 'Final synchronization resent unchanged files or missed pending files.'
    foreach ($expected in @($changedPath, $newPath, $retryPath)) {
        Assert-True -Condition ($expected -in $finalUploads) -Message "Final synchronization omitted $expected."
    }
    & $mirrorModule { param($Call) $script:Uploads.Clear(); $script:Settings.OrchestratorSharePointUploadIntervalMinutes = 0; & $Call } $finalScript
    Assert-True -Condition ((& $mirrorModule { $script:Uploads.Count }) -eq 0) -Message 'Repeated finalization resent unchanged files.'
    Set-Content -LiteralPath $newPath -Value '{"New":"changed with interval disabled"}' -Encoding utf8
    & $mirrorModule { param($Call) & $Call } $finalScript
    Assert-True -Condition ((& $mirrorModule { $script:Uploads.Count }) -eq 1) -Message 'Finalization did not run with periodic uploads disabled.'
    # A planned recycle (runtime update or lifetime) defers the mirror to the next instance.
    Set-Content -LiteralPath $newPath -Value '{"New":"changed before planned recycle"}' -Encoding utf8
    & $mirrorModule { param($Call) $script:Uploads.Clear(); $script:Logs.Clear(); $script:OrchestratorStopReason = 'RuntimeUpdate'; & $Call } $finalScript
    $recycleLogs = @(& $mirrorModule { $script:Logs.ToArray() })
    Assert-True -Condition ((& $mirrorModule { $script:Uploads.Count }) -eq 0 -and @($recycleLogs | Where-Object { $_ -like '*deferred to the next orchestrator instance*' }).Count -eq 1) -Message 'A planned recycle did not defer the operational mirror.'
    & $mirrorModule { param($Call) $script:OrchestratorStopReason = 'Stopped'; & $Call } $finalScript
    Assert-True -Condition ((& $mirrorModule { $script:Uploads.Count }) -eq 1) -Message 'The change deferred by a planned recycle was not mirrored by the next finalization.'

    $moduleWarnings = @()
    Import-Module -Name $coreModulePath -MinimumVersion '1.0.56' -Force -ErrorAction Stop -WarningVariable moduleWarnings
    Assert-True -Condition ($moduleWarnings.Count -eq 0) -Message ("SmartM365.Core import emitted warning(s): {0}" -f ($moduleWarnings -join ' | '))
    $approvedVerbs = @(Get-Verb | Select-Object -ExpandProperty Verb)
    $unapprovedCommands = @(Get-Command -Module SmartM365.Core | Where-Object { $_.Name -match '-' -and $_.Name.Split('-')[0] -notin $approvedVerbs })
    $unapprovedCommandNames = @($unapprovedCommands | ForEach-Object { $_.Name } | Sort-Object)
    Assert-True -Condition ($unapprovedCommands.Count -eq 0) -Message ("SmartM365.Core exports command(s) with unapproved verbs: {0}" -f ($unapprovedCommandNames -join ', '))
    Assert-True -Condition ($null -ne (Get-Command Initialize-SmartM365SharePointFolder -Module SmartM365.Core -ErrorAction SilentlyContinue)) -Message 'Initialize-SmartM365SharePointFolder is not exported.'
    Assert-True -Condition ($null -eq (Get-Command Ensure-SmartM365SharePointFolder -Module SmartM365.Core -ErrorAction SilentlyContinue)) -Message 'The obsolete Ensure-SmartM365SharePointFolder command is still exported.'
    $coreModule = Get-Module SmartM365.Core | Select-Object -First 1
    & $coreModule {
        $script:SmartM365SharePointFolderPathCache = @{}
        $script:FolderRequests = [Collections.Generic.List[string]]::new()
        function script:Invoke-SmartM365GraphRestWithRetry {
            param($Method, $Uri, $Body, $ContentType, $Operation)
            $script:FolderRequests.Add("$Method $Uri")
            return [pscustomobject]@{ id = 'created' }
        }
        $folderPath = 'SMART-M365/DATA/DATA-ALL/Orchestrator/Config/Versions'
        if (-not (Ensure-SmartM365SharePointDriveFolderPath -DriveId 'drive-1' -FolderPath $folderPath)) { throw 'Folder creation returned false.' }
        $firstRequestCount = $script:FolderRequests.Count
        if ($firstRequestCount -ne 6) { throw "Expected 6 segment creation requests, got $firstRequestCount." }
        if (-not (Ensure-SmartM365SharePointDriveFolderPath -DriveId 'drive-1' -FolderPath $folderPath)) { throw 'Cached folder creation returned false.' }
        if ($script:FolderRequests.Count -ne $firstRequestCount) { throw 'Folder cache did not prevent duplicate Graph requests.' }

        $script:SmartM365SharePointFolderPathCache = @{}
        function script:Invoke-SmartM365GraphRestWithRetry {
            param($Method, $Uri, $Body, $ContentType, $Operation)
            throw 'Ensure SharePoint folder failed. Method=POST; Status=409; Body={"error":{"code":"nameAlreadyExists"}}'
        }
        if (-not (Ensure-SmartM365SharePointDriveFolderPath -DriveId 'drive-2' -FolderPath 'SMART-M365/DATA')) {
            throw 'An existing SharePoint folder was not treated as an idempotent success.'
        }
    }

    # A transiently missing lease must not turn this scan into a remote deletion.
    Set-Content -LiteralPath $leasePath -Value '{"LeaseId":"deferred"}' -Encoding utf8
    & $mirrorModule { Invoke-OrchestratorSharePointMirror }
    Remove-Item -LiteralPath $leasePath -Force
    $beforeDeferredDeletes = & $mirrorModule { $script:Deletes.Count }
    & $mirrorModule {
        function script:Get-OrchestratorSharePointMirrorSnapshot {
            param([string]$SharedDataFolderPath)
            $result = & $script:OriginalSnapshot -SharedDataFolderPath $SharedDataFolderPath
            $result.DeferredPaths = @('DATA-ALL/Orchestrator/Election/Concurrency/SharedRuntime.json.txt','DATA-ALL/Orchestrator/Election/Concurrency/SharedRuntime.json')
            $result
        }
        Invoke-OrchestratorSharePointMirror
    }
    Assert-True -Condition ((& $mirrorModule {$script:Deletes.Count}) -eq $beforeDeferredDeletes) -Message 'Deferred scan deleted remote lease.'
    $deferredState = & $mirrorModule { Read-OrchestratorSharePointMirrorState $script:Settings.SharePointMirrorStatePath }
    Assert-True -Condition (@($deferredState.Files | Where-Object RelativePath -like '*/Concurrency/SharedRuntime.json.txt').Count -eq 1) -Message 'Deferred scan lost previous mirror entry.'
    & $mirrorModule {
        Set-Item Function:script:Get-OrchestratorSharePointMirrorSnapshot -Value $script:OriginalSnapshot
        Invoke-OrchestratorSharePointMirror
    }
    Assert-True -Condition ((& $mirrorModule {$script:Deletes.Count}) -eq ($beforeDeferredDeletes + 1)) -Message 'Next complete scan did not reconcile expired lease.'
    "ORCHESTRATOR_SHAREPOINT_MIRROR_TEST_OK Files=$($relativeFiles.Count); InitialUploads=$firstUploadCount; ExpiredLeaseDeletes=1"
}
finally {
    if ($mirrorModule) { Remove-Module $mirrorModule -Force -ErrorAction SilentlyContinue }
    Remove-Module SmartM365.Core -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $temporaryRoot) { Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
