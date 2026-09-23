<#
.SYNOPSIS
Runs offline tests for the orchestrator SharePoint operational-folder mirror.
.VERSION
1.0.1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
    'Invoke-OrchestratorSharePointMirror'
)
$definitions = foreach ($name in $functionNames) {
    $node = $ast.Find({ param($candidate) $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq $name }, $true)
    if (-not $node) { throw "Missing orchestrator function: $name" }
    $node.Extent.Text
}
$mirrorModule = New-Module -ScriptBlock ([scriptblock]::Create($definitions -join [Environment]::NewLine))

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

    "ORCHESTRATOR_SHAREPOINT_MIRROR_TEST_OK Files=$($relativeFiles.Count); InitialUploads=$firstUploadCount; ExpiredLeaseDeletes=1"
}
finally {
    if ($mirrorModule) { Remove-Module $mirrorModule -Force -ErrorAction SilentlyContinue }
    Remove-Module SmartM365.Core -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $temporaryRoot) { Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAtBMqOMG5jXPNh
# MedhZqzAbcVsmMuklchswwjE2mgVpaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEoY5R3z/ttiFnLLRE41mxtJdwcj/FZmx1chzTPmT94eMA0GCSqG
# SIb3DQEBAQUABIIBgKEysfGh8uoXXLzC8kXbNiVhISQjXZwzSzwSO0xNC83AR7l/
# DIHIMEc6mO/TqR/zhG6axkVgw8y+foR2EEmZ1Lw7Gdivc6WQVtBkkKIPn3QSvyKr
# u75lhCTgFbRVxlAMZJ6H5RQgyqD+qgivJnZSxxbJg5IYc0zhjMEbnUJ2rltYqkFN
# Aa1WBhxKceHoH2YQfQKJgW29YAQrrc0x+iQ6d7k8gvC9WgmE6eSV6HCiZI3W2udG
# QGruVB8C7jvyP8Emi2Im1se5wOJQkz7y7/WjpTQZDXWLu/1uaRdSVUZ/1v/VSqSr
# fRzXhQ3n3WI7hzqb6CbTdxbHFlDAd8gWG45zpkZrYRofScS76PTAV+kEbWe3frbg
# sx77PS3NWmAsy4alzXwJkzCsaE1hWXM//TX7RG8MqojriQdVen53fp4JaGw6//8h
# sOh1j/I6YkorXd6XMSoRbVC3PKVOmqb02Sj2SPTTyCkrEkK7HB15PkRx13wmIfVD
# d9+bkBPxTDeKNY2Su6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjMxNTEz
# NDBaMC8GCSqGSIb3DQEJBDEiBCCUsBZs3WZUeL5clvM+OsseZK5OHY63kdR8WDUp
# AN2a7TANBgkqhkiG9w0BAQEFAASCAgCWN+5YRniDWlg+uY0N24Pwtgr3SyI6SYdR
# YxqRHX7PlVfk3GvJkhkSFIwiZHJr79JL0ArsfmUeDjZ66FO8Jr7INcX3FLxJfDh8
# we3BECOGI2f2zCy5wGVI9WwfR+PP5PsNqXgyHGFgRbqVVOxoXTvliZumAdnnboP0
# o6fSAyInO9A7m7ph6hoWctr4iiyXQcHduubL2JkGycTQ3q6AHPYyvDfYvD6HqJTm
# YpLUEmdaSJOPrx3mafY/hRC+e4iQnJochsT+kAPEdC0ypIv8BxPp26qEiPUmmK/L
# Dc1R3vwjqpNg0doOcbC+HOzwnnU2MIqk5h4XOvPWosbudG0KuKL7wmQbA4aLKuKi
# xxtaHwl3iwA4KGDyaKCHuI8NehVdYU00/7qbcEVNy2KIUEciwxxnE/ZfEge8TMuc
# Wl0ypMslbh9lKCReZ0dHHj7zC/ONEFwJpa77oTUzXxewtUVp7dSXDNOE8sqVNCSd
# OVUG6+q1mvcQISOSappc7l9Cwr72JCc0UirI0Iw6SN1majZ+kZdmF/u/bhOziLCa
# M4qPottWpqFNwN1k3vte2NzEo+xS+NczyM7Z39qfksKBalyXyUMH1Wt9gi65D1oj
# +03sSnw+YJuHYJv1kl5HB2zww5PvNhMfCS6UAh9SrcW1js/lDBbHI7V0aiafWbms
# b8lgF5HLlw==
# SIG # End signature block
