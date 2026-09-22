<#
.SYNOPSIS
Runs offline pipeline and production-manifest contract tests for the SmartM365 orchestrator.
.VERSION
1.0.1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$orchestratorRoot = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator'
$smartInventoryRoot = Split-Path -Path $orchestratorRoot -Parent
$pipelineModuleFile = Join-Path -Path $orchestratorRoot -ChildPath 'SmartM365.Orchestrator.Pipeline.psm1'
$orchestratorPath = Join-Path -Path $orchestratorRoot -ChildPath 'SmartM365-Inventory-Orchestrator.ps1'
$jobsTemplatePath = Join-Path -Path $orchestratorRoot -ChildPath 'Orchestrator-Jobs.json.template'
$pipelineModule = Import-Module -Name $pipelineModuleFile -Force -PassThru -ErrorAction Stop

function Assert-True {
    param([Parameter(Mandatory)][bool]$Condition, [Parameter(Mandatory)][string]$Message)
    if (-not $Condition) { throw $Message }
}

function New-SyntheticJob {
    param([string]$Name, [string]$Group, [bool]$Enabled = $true, [string]$Mode = 'Elected', [string[]]$DependsOn = @())
    [pscustomobject][ordered]@{
        Name = $Name
        Group = $Group
        Enabled = $Enabled
        AssignmentMode = $Mode
        DependsOn = @($DependsOn)
        RequiredCapabilities = @('SharedRuntime')
    }
}

$temporaryRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('SmartM365-OrchestratorPipeline-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $temporaryRoot -Force | Out-Null
    $document = [pscustomobject]@{
        Jobs = @(
            (New-SyntheticJob -Name 'Parent' -Group 'M365'),
            (New-SyntheticJob -Name 'Child' -Group 'Intune' -DependsOn @('Parent')),
            (New-SyntheticJob -Name 'Disabled' -Group 'Intune' -Enabled $false),
            (New-SyntheticJob -Name 'Manual' -Group 'Intune' -Mode 'Manual')
        )
    }

    $selection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -Pipeline Intune
    Assert-True ($selection.SelectedCount -eq 2) 'Dependency closure did not include the enabled parent.'
    Assert-True ($selection.AddedDependencyCount -eq 1) 'Dependency closure count is incorrect.'
    Assert-True (@($selection.SelectedJobs.Name) -contains 'Parent') 'The dependency parent is absent.'
    Assert-True (@($selection.ExcludedJobs | Where-Object Reason -EQ 'Disabled').Count -eq 1) 'Disabled exclusion was not reported.'
    Assert-True (@($selection.ExcludedJobs | Where-Object Reason -EQ 'Manual').Count -eq 1) 'Manual exclusion was not reported.'

    $selectionAgain = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -Pipeline Intune
    Assert-True ((@($selection.SelectedJobs.Name) -join '|') -eq (@($selectionAgain.SelectedJobs.Name) -join '|')) 'Selection is not idempotent.'

    $request = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $temporaryRoot -Tenant test -Pipeline Intune -Selection $selection -ManifestHash 'ABC123' -RequestedBy 'offline-test' -RequestedFrom 'offline-test'
    Assert-True ($request.OverallStatus -eq 'Running' -and $request.PendingCount -eq 2) 'New request status is incorrect.'
    Assert-True (Test-Path -LiteralPath $request.RequestPath -PathType Leaf) 'Atomic request document is missing.'

    $retrySource = Join-Path -Path $temporaryRoot -ChildPath 'retry-source.tmp'
    $retryDestination = Join-Path -Path $temporaryRoot -ChildPath 'retry-destination.json'
    [IO.File]::WriteAllText($retrySource, '{"retry":true}', [Text.UTF8Encoding]::new($false))
    $retryAttempts = & $pipelineModule {
        param($SourcePath, $DestinationPath)
        $state = [pscustomobject]@{ Attempts = 0 }
        $operation = {
            param($Source, $Destination)
            $state.Attempts++
            if ($state.Attempts -lt 3) { throw [IO.IOException]::new('Synthetic transient move failure.') }
            [IO.File]::Move($Source, $Destination, $true)
        }.GetNewClosure()
        Move-SmartM365OrchestratorPipelineFileWithRetry `
            -SourcePath $SourcePath `
            -DestinationPath $DestinationPath `
            -MaximumAttempts 5 `
            -RetryDelayMilliseconds 1 `
            -MoveOperation $operation
        $state.Attempts
    } $retrySource $retryDestination
    Assert-True ($retryAttempts -eq 3 -and (Test-Path -LiteralPath $retryDestination -PathType Leaf)) 'Transient pipeline status replacement was not retried successfully.'

    $boundedRetryAttempts = & $pipelineModule {
        param($SourcePath, $DestinationPath)
        $state = [pscustomobject]@{ Attempts = 0 }
        $operation = {
            param($Source, $Destination)
            $state.Attempts++
            throw [UnauthorizedAccessException]::new('Synthetic persistent access denial.')
        }.GetNewClosure()
        try {
            Move-SmartM365OrchestratorPipelineFileWithRetry `
                -SourcePath $SourcePath `
                -DestinationPath $DestinationPath `
                -MaximumAttempts 2 `
                -RetryDelayMilliseconds 1 `
                -MoveOperation $operation
        }
        catch [UnauthorizedAccessException] { return $state.Attempts }
        throw 'Persistent access denial was unexpectedly swallowed.'
    } $retryDestination (Join-Path -Path $temporaryRoot -ChildPath 'never-created.json')
    Assert-True ($boundedRetryAttempts -eq 2) 'Persistent pipeline replacement failure did not stop at the configured retry bound.'

    $concurrentRejected = $false
    try {
        New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $temporaryRoot -Tenant test -Pipeline Intune -Selection $selection -ManifestHash 'ABC123' | Out-Null
    }
    catch { $concurrentRejected = $_.Exception.Message -like '*active pipeline request already exists*' }
    Assert-True $concurrentRejected 'A concurrent active pipeline request was accepted.'

    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId -JobName Parent -Status Success -Attempt 0 -OwnerServer SERVER-A | Out-Null
    $retryNotBefore = [datetime]::UtcNow.AddMinutes(5).ToString('o')
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId -JobName Child -Status RetryScheduled -Attempt 1 -OwnerServer SERVER-B -NotBeforeUtc $retryNotBefore | Out-Null
    $retrying = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId
    $retryingChild = @($retrying.Jobs | Where-Object JobName -EQ Child)[0]
    Assert-True (-not $retrying.IsTerminal -and $retryingChild.Attempt -eq 1 -and $retryingChild.NotBeforeUtc -eq $retryNotBefore) 'Retry timing/attempt metadata was not preserved.'
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId -JobName Child -Status CompletedWithWarnings -Attempt 0 -OwnerServer SERVER-B | Out-Null
    $completed = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $temporaryRoot -BatchId $request.BatchId
    Assert-True ($completed.IsTerminal -and $completed.OverallStatus -eq 'CompletedWithWarnings') 'Warning completion was not aggregated correctly.'

    $secondRequest = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $temporaryRoot -Tenant test -Pipeline Intune -Selection $selection -ManifestHash 'ABC123'
    Assert-True ($secondRequest.BatchId -ne $request.BatchId) 'A completed request prevented the next unique batch.'
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $secondRequest.BatchId -JobName Parent -Status Failed -Attempt 1 -OwnerServer SERVER-A -Detail 'Synthetic failure' | Out-Null
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $temporaryRoot -BatchId $secondRequest.BatchId -JobName Child -Status BlockedDependencyFailed -OwnerServer SERVER-B | Out-Null
    $failed = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $temporaryRoot -BatchId $secondRequest.BatchId
    Assert-True ($failed.IsTerminal -and $failed.OverallStatus -eq 'Failed' -and $failed.FailedCount -eq 2) 'Failure aggregation is incorrect.'

    $source = Get-Content -LiteralPath $orchestratorPath -Raw
    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$tokens, [ref]$parseErrors)
    Assert-True (@($parseErrors).Count -eq 0) 'Orchestrator source failed parser validation.'
    $scanDefinitions = foreach ($functionName in @('Update-OrchestratorPipelineRequests', 'Get-OrchestratorPipelineDependencyStatus')) {
        $node = $ast.Find({ param($candidate) $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq $functionName }, $true)
        if ($null -eq $node) { throw "Missing orchestrator function: $functionName" }
        $node.Extent.Text
    }

    $scanRoot = Join-Path -Path $temporaryRoot -ChildPath 'Scan'
    New-Item -ItemType Directory -Path $scanRoot -Force | Out-Null
    $scanManifestPath = Join-Path -Path $scanRoot -ChildPath 'Orchestrator-Jobs.json'
    $document | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $scanManifestPath -Encoding utf8
    $scanSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -Pipeline Full
    $scanRequest = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $scanRoot -Tenant test -Pipeline Full -Selection $scanSelection -ManifestHash ((Get-FileHash -LiteralPath $scanManifestPath -Algorithm SHA256).Hash)
    Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $scanRoot -BatchId $scanRequest.BatchId -JobName Child -Status RetryScheduled -Attempt 1 -OwnerServer SERVER-B -NotBeforeUtc ([datetime]::UtcNow.AddMinutes(10).ToString('o')) | Out-Null
    $scanModule = New-Module -ScriptBlock ([scriptblock]::Create($scanDefinitions -join [Environment]::NewLine))
    $scanResult = & $scanModule {
        param($Root, $ManifestPath, $JobsDocument, $BatchId)
        $script:Settings = @{ SharedDataFolderPath = $Root; JobsManifestPath = $ManifestPath }
        $script:Manifest = @{ JobsByName = @{ Parent = $JobsDocument.Jobs[0]; Child = $JobsDocument.Jobs[1] } }
        $script:PipelinePending = @{}
        $script:Tenant = 'test'
        function script:Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message) }
        function script:Set-OrchestratorPipelineJobStatus { param($BatchId, $JobName, $Status, $Detail) throw "Unexpected rejection: $JobName/$Status/$Detail" }
        Update-OrchestratorPipelineRequests
        [pscustomobject]@{
            PendingNames = @($script:PipelinePending.Keys)
            ParentStatus = Get-OrchestratorPipelineDependencyStatus -BatchId $BatchId -JobName Parent
        }
    } $scanRoot $scanManifestPath $document $scanRequest.BatchId
    Assert-True ($scanResult.PendingNames.Count -eq 1 -and $scanResult.PendingNames[0] -eq 'Parent') 'A retry was made runnable before NotBeforeUtc.'
    Assert-True ($scanResult.ParentStatus -eq 'Pending') 'Shared dependency status was not read from the batch.'

    $mailNode = $ast.Find({ param($candidate) $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq 'Send-OrchestratorRuntimeUpdateEmail' }, $true)
    Assert-True ($null -ne $mailNode) 'Runtime update email function is missing.'
    $mailModule = New-Module -ScriptBlock ([scriptblock]::Create($mailNode.Extent.Text))
    $mailResult = & $mailModule {
        $script:RuntimeUpdateBaseline = [pscustomobject]@{ ScriptVersion = [version]'1.5.13'; CoreVersion = [version]'1.0.49' }
        $script:RunningJobs = @{ JobA = @{}; JobB = @{} }
        $script:LastRuntimeUpdateMailIdentity = ''
        $script:Tenant = 'test'
        $script:ScriptName = 'SmartM365-Inventory-Orchestrator'
        $script:ScriptVersion = '1.5.14'
        $script:MailCalls = New-Object 'System.Collections.Generic.List[object]'
        function script:ConvertTo-HtmlText { param($Text) [Net.WebUtility]::HtmlEncode([string]$Text) }
        function script:Send-OrchestratorMail {
            param($Subject, $HtmlBody)
            $script:MailCalls.Add([pscustomobject]@{ Subject = $Subject; Body = $HtmlBody })
            return $true
        }
        $candidate = [pscustomobject]@{ ScriptVersion = [version]'1.5.14'; CoreVersion = [version]'1.0.50'; Identity = 'offline-runtime-1.5.14' }
        $first = Send-OrchestratorRuntimeUpdateEmail -Candidate $candidate -DetectedAt ([datetime]'2026-09-21T13:00:00Z')
        $second = Send-OrchestratorRuntimeUpdateEmail -Candidate $candidate -DetectedAt ([datetime]'2026-09-21T13:01:00Z')
        [pscustomobject]@{ First = $first; Second = $second; Calls = $script:MailCalls.ToArray() }
    }
    Assert-True ($mailResult.First -and $mailResult.Second -and $mailResult.Calls.Count -eq 1) 'Runtime update mail anti-duplication failed.'
    Assert-True ($mailResult.Calls[0].Subject -like '*1.5.13 -> 1.5.14*') 'Runtime update mail subject does not contain the version transition.'
    Assert-True ($mailResult.Calls[0].Body -like '*Stable, parser-valid and Authenticode-approved*' -and $mailResult.Calls[0].Body -like '*Detached jobs still supervised*') 'Runtime update mail body is incomplete.'

    $productionDocument = Get-Content -LiteralPath $jobsTemplatePath -Raw | ConvertFrom-Json -Depth 100
    $missingExternalActionOptIns = foreach ($job in @($productionDocument.Jobs)) {
        if ([string]::IsNullOrWhiteSpace([string]$job.ScriptPath)) { continue }
        $jobScriptPath = Join-Path -Path $smartInventoryRoot -ChildPath ([string]$job.ScriptPath)
        if (-not (Test-Path -LiteralPath $jobScriptPath -PathType Leaf)) { continue }
        $jobScriptSource = Get-Content -LiteralPath $jobScriptPath -Raw
        $supportsExternalActionOptIn = $jobScriptSource -match '\[switch\]\s*\$EnableConfiguredExternalActions'
        $enablesExternalActions = [string]$job.Arguments -match '(^|\s)-EnableConfiguredExternalActions(?=\s|$)'
        if ($supportsExternalActionOptIn -and -not $enablesExternalActions) { [string]$job.Name }
    }
    Assert-True (@($missingExternalActionOptIns).Count -eq 0) ("Jobs declaring EnableConfiguredExternalActions must opt in explicitly: {0}" -f (@($missingExternalActionOptIns) -join ', '))
    $productionSelection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $productionDocument -Pipeline Full
    Assert-True ($productionSelection.SelectedCount -eq 42) 'Published Full selection must contain the 42 enabled non-manual jobs.'
    Assert-True ($productionSelection.ExcludedCount -eq 6) 'Published Full selection must exclude the 6 disabled/manual jobs.'

    $facadeRoot = Join-Path -Path $temporaryRoot -ChildPath 'Facade'
    $facadeElectionRoot = Join-Path -Path $facadeRoot -ChildPath 'Election'
    New-Item -ItemType Directory -Path $facadeElectionRoot -Force | Out-Null
    $facadeManifestPath = Join-Path -Path $facadeRoot -ChildPath 'Orchestrator-Jobs.json'
    Copy-Item -LiteralPath $jobsTemplatePath -Destination $facadeManifestPath
    $facadePlan = [pscustomobject]@{
        PlanId = 'offline-plan'
        Assignments = @($productionSelection.SelectedJobs | ForEach-Object { [pscustomobject]@{ JobName = $_.Name; OwnerServer = $env:COMPUTERNAME } })
    }
    $facadePlan | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $facadeElectionRoot 'Orchestrator-ElectionPlan.json') -Encoding utf8
    $facadePath = Join-Path -Path $orchestratorRoot -ChildPath 'SmartM365-Inventory-Pipeline.ps1'
    $facadeOutput = & (Join-Path $PSHOME 'pwsh.exe') -NoProfile -ExecutionPolicy Bypass -File $facadePath -Tenant test -Pipeline Full -Collect -NoWait -JobsManifestPath $facadeManifestPath -SharedDataFolderPath $facadeRoot 2>&1
    Assert-True ($LASTEXITCODE -eq 0) ("The offline Collect/NoWait façade failed: {0}" -f (@($facadeOutput) -join [Environment]::NewLine))
    Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $facadeRoot 'PipelineRuns') -Directory).Count -eq 1) 'The façade did not publish exactly one batch.'

    foreach ($requiredHook in @(
        'Update-OrchestratorPipelineRequests',
        'Get-OrchestratorPipelineDependencyStatus',
        'PipelineBatchId',
        'Set-OrchestratorPipelineJobStatus',
        'Send-OrchestratorRuntimeUpdateEmail -Candidate $candidate',
        "if (`$reason -in @('due', 'forced'))"
    )) {
        Assert-True ($source.Contains($requiredHook)) "Orchestrator pipeline hook is missing: $requiredHook"
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) { Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ("[{0}] SmartM365 Orchestrator pipeline tests passed." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA3+r+uJnKnpJh8
# AIfxgxLQiMRAcT/0+i6ZkXIxi0U1/qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII/78ImjwCl79QhQEhqJPd40s5UIyG28svs9dwYGRMAbMA0GCSqG
# SIb3DQEBAQUABIIBgCXPkSc9gteRvg+gdYuD1NlV2SlXfQ/RxIXY21MIu6A7KVPI
# pflUCWxbr0eIWDWvAcROzaWnN/LmfbAlsCeBiJGsHyjujyOaVePbBheS7d6B/HeX
# HD/wOYBiZiMRVkVtjn1lnjWkT23vOkK5+aPoUPQfTkAspw8TjmiECS/Y1b2GeAxa
# drTUVYBMXKjREKJSDgi3BsMxO2Tq+5lgAp9LnGjXyzPjnWuzr6Zc9IFvcdTydc4Y
# PK5gx7NeyEgeQg1riYdxW+TGE4waSa8Z5/isSCMMmf75E80rYgmQcGu6DoUcMNki
# RgGwvhzHkIpPJD9ETVt6uR5poAro8ORuja2HjxkaJXcdfrMYp5f7EqJMPKIE++2C
# WerxqfvzzZFxIQk52zqg8bw23kuw/wQIhsYYW85lE0CCntbISMqIy6ou4G73+7PA
# vwmmn2qJVGAK/CaubU4SnDlKGg2Jugw7OKibmbRB2uPS/d1dbnyVyd0ZoZ1RKzer
# 32MhX2yqG1k9mikVx6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjIxNjA3
# MzJaMC8GCSqGSIb3DQEJBDEiBCCU6QigtMMNC1V2j31Ysz0Q2uq6kpYZEzfIUOvh
# V0mhbDANBgkqhkiG9w0BAQEFAASCAgAmrI1CWKtrwdlNFTqE0zbHiXX8deMkFJiF
# b0V1nvClk21U07m91yJkpGoq/NeFaj4paIJS/ujso1KZCrhQBklVRSHIOx3xvWbK
# 1hI3IcXGczlgB/S8Io8//Plv6uD+P36efE/Fu4DaGIhiq/zP72ClN/kmGA5WYkzH
# uoxDH23X8ZPAiIKeuD8ze/+YdieRi7XrQm/8VVYCyhSfDMB1H37ew1df8I1ZUwWr
# 1cLZRM8iRCh7rkwqOfqlbfJWVauxDm181OeRhgFjSobds8ZJzphW5uCD6ZyoGwR4
# 8KjNH6UtLSpMBtH/4DitnCStczK17lTQBDCRRiUDAHWZ4LfD2qMb/qhecKRWuNUa
# vAIl6fqJq/0+OAhcfBVXzmW3EF5IOlUSOGcMKU87u5QSV85BlXFjEq71NkRRwKX8
# 0tfe4VYGD/CvkAoEuJG993LW0DoOF4wVXaeS88KCyA181nGLsCD+qjQmCXJtDg8g
# MdONfY18qHL8v3blz5ShuciV++fb8e3hWvt2UnO223NdnG3FgtQZweIMoGRkhX+7
# NfnkIWQWCf5Y/isa597Y94f+TBevsifAke8WgN1jBZEnUrIYsB8qlt4TjkGScsfj
# mHnp/3DAGaFuZUVXNV7BBsfwsZ2WgxOgxmig7wLwL+qDFlLpACsLL+DL3q4yGXOe
# t9cCZMm2yA==
# SIG # End signature block
