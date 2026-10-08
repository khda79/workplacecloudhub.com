#Requires -Version 7.0
<#
.SYNOPSIS
Offline activation merge and preparation-to-publication boundary tests.
.VERSION
1.0.2
.NOTES
Extracts helpers only; no tenant initialization, collections, authentication,
SharePoint transport or shared configuration publication. Temporary fixtures only.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$smartRoot=Split-Path $PSScriptRoot -Parent
$preparedRoot=Join-Path $smartRoot 'SmartInventory/PreparedEvidence'
$testCount=0
function Assert-True {param([bool]$Value,[string]$Label) if(-not $Value){throw $Label};$script:testCount++}
function Assert-Rejected {param([scriptblock]$Body,[string]$Label) $failed=$false;try{& $Body | Out-Null}catch{$failed=$true};Assert-True $failed $Label}
function Import-Helper {
    param([string]$Path,[string]$Name)
    $parseErrors=$null;$tokens=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    Assert-True ($parseErrors.Count -eq 0) "Parse failed: $Name"
    $functions=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name},$true))
    Assert-True ($functions.Count -eq 1) "Helper missing or ambiguous: $Name"
    . ([scriptblock]::Create($functions[0].Extent.Text.Replace("function $Name", "function script:$Name")))
}
Import-Helper (Join-Path $preparedRoot 'SmartM365-CmdbEvidence-Prepare.ps1') 'Invoke-SmartM365CmdbPreparedPublication'
Import-Helper (Join-Path $preparedRoot 'SmartM365-CmdbEvidence-Orchestrator.ps1') 'New-SmartM365CmdbOrchestratorJobsDocument'
$template=Get-Content (Join-Path $smartRoot 'SmartInventory/Orchestrator/Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
$optInNames=@('M365-WorkplaceScope-Inventory','CmdbEvidence-Prepare')
foreach($name in $optInNames){
    $entries=@($template.Jobs | Where-Object Name -eq $name)
    Assert-True ($entries.Count -eq 1 -and -not $entries[0].Enabled -and $entries[0].RequiresExplicitActivation) 'Template must not silently activate CMDB jobs.'
}
foreach($name in @('M365-Teams-Inventory','M365-SPO-Inventory','Intune-WindowsAutopilot-Inventory')){
    $entry=@($template.Jobs | Where-Object Name -eq $name)[0]
    Assert-True ($entry.Schedule.Type -eq 'Daily') 'A core CMDB producer exceeds its 48-hour cadence.'
}
$original=$template | ConvertTo-Json -Depth 50 | ConvertFrom-Json
$original.Jobs=@($original.Jobs | Where-Object Name -notin $optInNames)
$candidate=Get-Content (Join-Path $preparedRoot 'cmdb-orchestrator-integration.json.txt') -Raw | ConvertFrom-Json
$before=$original | ConvertTo-Json -Depth 50 -Compress
$arguments=@{JobsDocument=$original;ProposedJobs=$candidate.ProposedJobs;ExecutionServers=@('WORKER-TEST');PrepareTime='07:30';WorkplaceScopeTime='00:05'}
$templateBefore=$template | ConvertTo-Json -Depth 50 -Compress
$activated=New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $template -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'
Assert-True ($activated.Jobs.Count -eq $template.Jobs.Count) 'Activation duplicated opt-in template entries.'
Assert-True (($template | ConvertTo-Json -Depth 50 -Compress) -ceq $templateBefore) 'Activation changed input template.'
foreach($name in $optInNames){
    $entry=@($activated.Jobs | Where-Object Name -eq $name)[0]
    Assert-True ($entry.Enabled -and $entry.AssignmentMode -eq 'Pinned' -and -not $entry.PSObject.Properties['RequiresExplicitActivation']) 'Template entry was not explicitly qualified for the selected worker.'
}
Assert-True (@($activated.Jobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')[0].Arguments -ceq '-EnableConfiguredExternalActions') 'Activation removed configured native uploads.'
$unmarked=$template | ConvertTo-Json -Depth 50 | ConvertFrom-Json
@($unmarked.Jobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')[0].PSObject.Properties.Remove('RequiresExplicitActivation')
Assert-Rejected {New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $unmarked -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'} 'A deliberately disabled native job was activated without review.'
$merged=New-SmartM365CmdbOrchestratorJobsDocument @arguments
Assert-True (($original | ConvertTo-Json -Depth 50 -Compress) -ceq $before) 'Input document changed.'
Assert-True ($merged.Jobs.Count -eq $original.Jobs.Count+2) 'Unexpected job count.'
foreach($job in $original.Jobs){
    $match=@($merged.Jobs | Where-Object Name -eq $job.Name)
    Assert-True ($match.Count -eq 1 -and ($match[0] | ConvertTo-Json -Depth 50 -Compress) -ceq ($job | ConvertTo-Json -Depth 50 -Compress)) "Existing job altered: $($job.Name)"
}
$prepare=@($merged.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare')[0]
$scope=@($merged.Jobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')[0]
foreach($job in @($prepare,$scope)){
    Assert-True ($job.Enabled -and $job.AssignmentMode -eq 'Pinned' -and @($job.AllowedServers).Count -eq 1 -and $job.AllowedServers[0] -eq 'WORKER-TEST') 'Unverified worker eligibility.'
    Assert-True (-not $job.PSObject.Properties['ConditionalGraphAppRoles']) 'Review-only property escaped.'
}
Assert-True ($prepare.Arguments -ceq '-Publish' -and $prepare.RequiredGraphAppRoles -contains 'Sites.Selected' -and $prepare.RequiredCapabilities -contains 'Graph') 'Verified publication requirements missing.'
Assert-True ($prepare.DependsOn.Count -eq 17 -and $prepare.DependencyMode -eq 'FreshSuccess' -and $prepare.DependencyMaxAgeHours -eq 240) 'Dependency gates changed.'
Assert-True ($prepare.DependsOn -contains 'EXO-Mailboxes-Inventory-Fast' -and $prepare.DependsOn -notcontains 'EXO-Mailboxes-Inventory') 'Preparation waits for weekly stats instead of daily mailbox acquisition.'
Assert-True ($prepare.Schedule.Times[0] -eq '07:30' -and $scope.Schedule.Times[0] -eq '00:05') 'Daily schedule changed.'
$again=New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $merged -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'
Assert-True (($merged | ConvertTo-Json -Depth 50 -Compress) -ceq ($again | ConvertTo-Json -Depth 50 -Compress)) 'Merge is not idempotent.'
$scope.Arguments='-EnableConfiguredExternalActions'
$retained=New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $merged -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'
Assert-True (@($retained.Jobs | Where-Object Name -eq $scope.Name)[0].Arguments -ceq $scope.Arguments) 'Existing native scope configuration overwritten.'
$prepare.Arguments='-ValidateOnly'
Assert-Rejected {New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $merged -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'} 'Conflicting CMDB job accepted.'
Assert-Rejected {New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $original -ProposedJobs $candidate.ProposedJobs -ExecutionServers ' ' -PrepareTime '07:30' -WorkplaceScopeTime '00:05'} 'Empty worker accepted.'
Assert-Rejected {New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $original -ProposedJobs $candidate.ProposedJobs -ExecutionServers @('ONE','TWO') -PrepareTime '07:30' -WorkplaceScopeTime '00:05'} 'Multiple pinned workers accepted.'
$duplicate=$original | ConvertTo-Json -Depth 50 | ConvertFrom-Json
$duplicate.Jobs=@($duplicate.Jobs)+@($scope,$scope)
Assert-Rejected {New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $duplicate -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'} 'Duplicate scope accepted.'
$disabled=$original | ConvertTo-Json -Depth 50 | ConvertFrom-Json
@($disabled.Jobs | Where-Object Name -eq 'Intune-DiscoveredApps-Inventory')[0].Enabled=$false
Assert-Rejected {New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $disabled -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'} 'Disabled dependency silently skipped.'
$conflict=$original | ConvertTo-Json -Depth 50 | ConvertFrom-Json
$wrongScope=$scope | ConvertTo-Json -Depth 25 | ConvertFrom-Json
$wrongScope.ScriptPath='unreviewed.ps1'
$conflict.Jobs=@($conflict.Jobs)+@($wrongScope)
Assert-Rejected {New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $conflict -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'} 'Conflicting native scope overwritten.'
Import-Module (Join-Path $smartRoot 'SmartInventory/Orchestrator/SmartM365.Orchestrator.Management.psm1') -Force
$valid=New-SmartM365CmdbOrchestratorJobsDocument @arguments
$validation=Test-SmartM365OrchestratorJobsDocument -Document $valid
Assert-True $validation.Valid ('Actual jobs validator failed: '+($validation.Errors -join '; '))
$cluster=[pscustomobject]@{ExpectedOrchestratorServers=@('WORKER-TEST')}
Assert-True (Test-SmartM365OrchestratorClusterDocument -Document $cluster).Valid 'Cluster rejected.'
# Template's unrelated pinned workers need not belong to this synthetic cluster.
$own=[pscustomobject]@{Jobs=@($valid.Jobs | Where-Object Name -in @('CmdbEvidence-Prepare','M365-WorkplaceScope-Inventory'))}
Assert-True (Test-SmartM365OrchestratorConfigurationConsistency -JobsDocument $own -ClusterDocument $cluster).Valid 'New job consistency rejected.'
Import-Module (Join-Path $smartRoot 'SmartInventory/Orchestrator/SmartM365.Orchestrator.Pipeline.psm1') -Force
$selection=Get-SmartM365OrchestratorPipelineSelection -JobsDocument $valid -JobName 'CmdbEvidence-Prepare' -IncludeDependencies
Assert-True ($selection.SelectedJobs.Count -eq 18) 'Native producer omitted from dependency selection.'
$contract=Get-Content (Join-Path $preparedRoot 'cmdb-prepared-contract.json.txt') -Raw | ConvertFrom-Json
Assert-True ($contract.maxAgeHours -eq 48 -and $contract.freshnessGroups[0].maxAgeHours -eq 240 -and $contract.freshnessGroups[0].warningAgeHours -eq 168) 'Freshness policy changed.'
$activation=Get-Content (Join-Path $preparedRoot 'SmartM365-CmdbEvidence-Orchestrator.ps1') -Raw
Assert-True ($activation -match 'if\(\$Apply\)' -and $activation -match '-ExpectedJobsHash \$snapshot.JobsHash -ExpectedClusterHash \$snapshot.ClusterHash') 'Optimistic explicit activation guard missing.'
Assert-True ($activation -notmatch 'Submit-SmartM365|Request-SmartM365|Invoke-RestMethod|Send-.*Mail') 'Activation entry point expanded into execution or external notifications.'
Assert-True ($activation -match 'SmartM365TeamsNotificationInProgress=\$true' -and $activation -match 'SmartM365TeamsNotificationInProgress=\$previous' -and $activation -match 'EnableSharePointUpload=\$false') 'Local validation external-action guard missing.'
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('cmdb-automation-test-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture | Out-Null
try {
    $manifest=Join-Path $fixture 'current.json.txt'
    [IO.File]::WriteAllText($manifest,'{"Status":"SyntheticOnly"}')
    $script:calls=0
    $invoker={param($tenant,$root,$hash)
        Assert-True ($tenant -eq 'test' -and $root -eq $fixture -and $hash -ceq (Get-FileHash $manifest -Algorithm SHA256).Hash) 'Publication is not bound to newly prepared bytes.'
        $script:calls++;return 0
    }
    foreach($status in @('Prepared','PreparedWithCleanupWarning')){
        Invoke-SmartM365CmdbPreparedPublication -PreparationResult ([pscustomobject]@{Status=$status}) -PreparedRoot $fixture -TenantProfile test -PublisherInvoker $invoker
    }
    Assert-True ($script:calls -eq 2) 'Success did not reach publisher exactly once.'
    foreach($status in @('Failed','ValidatedSources','Unknown')){
        Assert-Rejected {Invoke-SmartM365CmdbPreparedPublication -PreparationResult ([pscustomobject]@{Status=$status}) -PreparedRoot $fixture -TenantProfile test -PublisherInvoker $invoker} "Invalid status accepted: $status"
    }
    Assert-True ($script:calls -eq 2) 'Failed generation reached transport.'
    foreach($returnValue in @(1,'0',$null)){
        $badInvoker={$returnValue}.GetNewClosure()
        Assert-Rejected {Invoke-SmartM365CmdbPreparedPublication -PreparationResult ([pscustomobject]@{Status='Prepared'}) -PreparedRoot $fixture -TenantProfile test -PublisherInvoker $badInvoker} 'Invalid publisher result reported success.'
    }
    Assert-Rejected {Invoke-SmartM365CmdbPreparedPublication -PreparationResult ([pscustomobject]@{Status='Prepared'}) -PreparedRoot (Join-Path $fixture 'absent') -TenantProfile test -PublisherInvoker $invoker} 'Missing manifest accepted.'
    $source=Get-Content (Join-Path $preparedRoot 'SmartM365-CmdbEvidence-Prepare.ps1') -Raw
    Assert-True ($source -match 'if\(\$Publish -and -not \$ValidateOnly\)') 'ValidateOnly publication exclusion missing.'
} finally {
    $resolved=[IO.Path]::GetFullPath($fixture)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if(-not $resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notmatch '^cmdb-automation-test-[a-f0-9]{32}$'){throw 'Unsafe temporary cleanup refused.'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
Write-Output "PASS: $testCount CMDB automation offline checks. No collections or live publication."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDlZoVB7X1yXoHo
# FUpC1YB6pLN3eYdhAOeFzqwFDjptyqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICxD5emOuLLjLt0fh66voUx+wHVTcj0xQaMOHp8JMvOZMA0GCSqG
# SIb3DQEBAQUABIIBgGjvwwEp2Q+wwvn9T/T70tBbclVeaHkOnmSQSvooRTCdS26X
# fgNlKBZpUED3EyulX0sNAMKUpc0kICgaslD7jAxZ/3pqscfmthAUPWcsHCjf76tM
# kZJdCGulj6pHseMdkDqV7lIZ5LhfBk3dE2WH1ZgGoXidJIFCSW0yiAqOR8HZiMED
# Hqwc6AI5w0EKBMo/W/jN8mJAWZN63hlgxZ2K47AInf8Qf/V37j/Hx/ICvvDFX5Vk
# v5KBspIAA9u1p0hMq5bV6zGX09tvoS7N/iyIQ+4rGGZshItlTv2w+IkMH3Gg6l7J
# r7D7uSbBDVucswfWTbm9mlpoWQL0ZFbwWrcHC5BeCOX5+Xkmd7gv+VVuslFoLjaH
# 9w+RucZAaZqvbFV1CDephogRxUBJs/EVCen75UWsPUAkqlhHHsbzcZKGkqykPrEr
# CCtFo09CArikhJY8qKQL0BPWuTet0IL76SBcQlrkDPZJ6RwgWQua2WID6RB6iqfy
# 5UPwgBljYcQHxyzUYKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxNDQ1
# NDlaMC8GCSqGSIb3DQEJBDEiBCB43lqlb91is5+GMvlt7B/6uihiqaHNn9Mr3rVH
# AEAfSjANBgkqhkiG9w0BAQEFAASCAgBEVFeY93tqM9LyHWr9eQEH7ZQT30Mfilu5
# V9UD/I4r7Ku8ly1J2++OIruQdYMIPrvRyp4d6P+SmsczYvfjmMPopr6UBfl/WbBI
# M78TL87N5KpvTlkrihYTAgKH/YcxZENoTRCCzmLhdNh776piz1CoW1XVg+3WxkQM
# To8/2cfcYsCzCsWWMXRSNKygO1bePj/PEzLPrX1uPQKH6BCoRrbXAage3mueZgEL
# crehRmDt8OxMXi3YaEmDVPBO+IazKZ2ng7B9yjVTejBhnQS4mcjP3tTQgNX2f6YQ
# 7ZmGYDFDnwk9YUU7nQXcFkp75FTF5/H+mKPgqc89Kgw/bZc/fwa23poyZsuPVPB2
# WDLrChqUhdDdwaJh4vhpVzxqa66OnM2sXPDwPOl6D2kaLnxuuGHidIP77NMlfBF6
# +xoVTRNe7s6veUMPD8jntiplkyeEpdztuYvSBUX7OskMEuyunIJ+plY5/CgOdHA5
# j7UhWdY4EeiA1cEI3EEtfICTamXVLDEWZdrikc/Fyf3x+DMV1TVbyDtlTtaztSAO
# /Bwz+w1i50+SSOXTZYvOyORS2KqgTJ+brDXNUHH23xjq6aY5PFznKoT/bqVGsP9F
# XTnrf6FUiIjdhkXTlnpYM4zpXIfWm40Gwvb5mTHLViGWsQHNlOWYn3TXBRHVTFtU
# s7kJAKnQWA==
# SIG # End signature block
