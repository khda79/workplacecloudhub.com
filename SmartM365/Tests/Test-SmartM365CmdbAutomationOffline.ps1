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
    Assert-True ($entries.Count -eq 1) 'Template integration job missing.'
    Assert-True ($entries[0].Enabled -and -not $entries[0].PSObject.Properties['RequiresExplicitActivation']) "$name must be enabled in the template."
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
$activationTemplate=$template | ConvertTo-Json -Depth 50 | ConvertFrom-Json
$activationScope=@($activationTemplate.Jobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')[0]
$activationScope.Enabled=$false
$activationScope | Add-Member -NotePropertyName RequiresExplicitActivation -NotePropertyValue $true
$activationPrepare=@($activationTemplate.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare')[0]
$activationPrepare.Enabled=$false
$activationPrepare | Add-Member -NotePropertyName RequiresExplicitActivation -NotePropertyValue $true
$activated=New-SmartM365CmdbOrchestratorJobsDocument -JobsDocument $activationTemplate -ProposedJobs $candidate.ProposedJobs -ExecutionServers WORKER-TEST -PrepareTime '07:30' -WorkplaceScopeTime '00:05'
Assert-True ($activated.Jobs.Count -eq $template.Jobs.Count) 'Activation duplicated opt-in template entries.'
Assert-True (($template | ConvertTo-Json -Depth 50 -Compress) -ceq $templateBefore) 'Activation changed input template.'
foreach($name in $optInNames){
    $entry=@($activated.Jobs | Where-Object Name -eq $name)[0]
    Assert-True ($entry.Enabled -and $entry.AssignmentMode -eq 'Pinned' -and -not $entry.PSObject.Properties['RequiresExplicitActivation']) 'Template entry was not explicitly qualified for the selected worker.'
}
Assert-True (@($activated.Jobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')[0].Arguments -ceq '-EnableConfiguredExternalActions') 'Activation removed configured native uploads.'
$unmarked=$activationTemplate | ConvertTo-Json -Depth 50 | ConvertFrom-Json
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDWe2rS+bhnxpNC
# 9Kuak1G9VFj58GKol81S0yMXnSTHxqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDgjy+MY2dnRIXy5d6NDU5g
# a8kiuuFEQTY1iONjqbtZuDANBgkqhkiG9w0BAQEFAASCAYCmlfPjUNJBkBa2f0GG
# +MXiaZsHSwHRcSLiDqWWdFiGcdy3YnWQmUA5PRv1D3JlfZXZyASIBioiraztkc5R
# oZ+30AgadbfVyqbwOF3EusWAGH0QypUAAP2TmX81PnRto9pDFS7suh+/VObRsH4d
# +CNl2mj3QxtBhfMavuOLKiugOcuHhtzVBL5gipaAUiQYF6Fg980YEsLSVKJhZ2F8
# 8gl+yjxePcNHDZpTdmTTkjw2o7fOlwAME7vJtX1IyvNh5cp+pv86liGaMvdTF0Ig
# VAxvguQrVp41HoPcjC5pZuVfR56Zx5vXPGssNkNrdweCu6If4nnYE1zhV5vwjrPM
# 01FiRYT6jO0fgmFLx6LqEwRtPh9hOof4HPy1PCE2DO1UrmdYqKXLUPi4BNX6OF6z
# iROWVJGJmNQZTsiGoLRn76tVSxvLH7ffuec9W8GxTLfSgW1y5QO2Z6+BXayamVFq
# RvmIWiWEJtY/rmwBB2Er8vI8ygS/77JpdhmLkkQLhi5covA=
# SIG # End signature block
