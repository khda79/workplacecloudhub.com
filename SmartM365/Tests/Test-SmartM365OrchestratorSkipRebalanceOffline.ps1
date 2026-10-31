<#
.SYNOPSIS
Exercise real ownership tracking, schedule selection and launch admission in memory.
The concurrency gate captures admitted occurrences and prevents every child launch.
.VERSION
1.0.1
#>
#requires -Version 7.0
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$source=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Orchestrator parser failed.'}
function Get-FunctionText([string]$Name){
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true)
    if($null -eq $node){throw "Missing function: $Name"};$node.Extent.Text
}
$names=@('Get-ElectedJobOwner','Test-OrchestratorLocalServerActive','Test-JobAllowedOnServer','Update-OrchestratorElectedOwnership',
    'Get-JobOccurrencesInWindow','Get-DueOccurrence','ConvertTo-StateTime','ConvertFrom-StateTime',
    'Set-OccurrenceSkipped','Invoke-LaunchPhase')
$definitions=@($names | ForEach-Object {Get-FunctionText $_})
$module=New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")+@'

# Match the resident scheduler, which does not enable StrictMode.
Set-StrictMode -Off
function Assert-Case($Condition,[string]$Message){if(-not $Condition){throw $Message}}
function Reset-Case {
    $script:Due=[datetime]::new(2026,10,27,0,25,0,[DateTimeKind]::Local)
    $script:Acquired=$script:Due.AddHours(10).AddMinutes(31)
    $script:Job=[pscustomobject]@{Name='SyntheticSkip';AssignmentMode='Elected';Enabled=$true;DependsOn=@();ConcurrencyKey='synthetic';Schedule=[pscustomobject]@{Type='Daily';Times=@('00:25');DaysOfWeek=@();MissedRunPolicy='Skip'}}
    $script:Manifest=[pscustomobject]@{OrderedJobs=@($script:Job);JobsByName=@{SyntheticSkip=$script:Job}}
    $script:ElectedOwnershipSince=@{}
    $script:ElectionPlan=[pscustomobject]@{PlanId='plan-a';Assignments=@([pscustomobject]@{JobName=$script:Job.Name;OwnerServer=$env:COMPUTERNAME})}
    $script:State=@{LastScheduledOccurrence=$script:Due.AddDays(-2).ToString('o');LastStatus='Success';PendingRetry=$null}
    $script:MaintenanceHealthy=$true;$script:MaintenanceControl=[pscustomobject]@{Enabled=$false}
    $script:StatePersistenceHealthy=$true;$script:FailSave=$false
    $script:Settings=[pscustomobject]@{MaxConcurrency=1;SharedDataFolderPath='unused'}
    $script:RunningJobs=@{};$script:ForcedPending=@();$script:PipelinePending=@{}
    $script:MaintenanceDue=$null;$script:MaintenanceCoverage=$null;$script:CoverageError=$false;$script:Saves=0;$script:Cleared=0
    $script:State.MaintenanceCatchUpRevision=8
    $script:Admitted=[Collections.Generic.List[object]]::new();$script:Rows=[Collections.Generic.List[object]]::new()
}
function Get-JobState {param($JobName) $script:State}
function Save-OrchestratorState {$script:Saves++;if($script:FailSave){$script:StatePersistenceHealthy=$false}}
function Add-JobRunCsvRow {param($JobName,$ScheduledTime,$StartTime,$EndTime,$DurationSec,$ExitCode,$Status,$RetryCount,$LogPath) $script:Rows.Add([pscustomobject]@{Name=$JobName;Occurrence=$ScheduledTime;Status=$Status})}
function Write-OrchestratorLog {param($Message,$Level)}
function Write-OrchestratorRuntimeUpdateWarning {param($Key,$Message,$Now)}
function Clear-DependencyWaitLog {param($JobName) $script:Cleared++}
function Test-JobSelected {param($JobName) $true}
function Get-JobEffectiveAllowedServers {param($Job) ,@($env:COMPUTERNAME)}
function Get-MaintenanceCatchUpOccurrence {param($Job,$State,$Control) $script:MaintenanceDue}
function Find-OrchestratorPeerMaintenanceCoverage {param($JobName,$Revision,$Occurrence) if($script:CoverageError){throw 'Synthetic peer read failure'}; $script:MaintenanceCoverage}
function Complete-MaintenanceCatchUpOccurrence {param($State,$Occurrence) $State.MaintenanceCatchUpLastOccurrence=$Occurrence.ToString('o')}
function Get-OrchestratorOccurrenceCoveredByLaterClaim {param($Job,$ExpectedOccurrence,$Now) $null}
function Test-SmartM365OrchestratorPipelineCancellation {param($SharedDataFolderPath,$BatchId) $false}
function Test-OrchestratorLocalConcurrencyBlocked {
    param($Job,$Occurrence,$LaunchedThisTick,[switch]$IsForced,$Now)
    $script:Admitted.Add([pscustomobject]@{Name=$Job.Name;Occurrence=$Occurrence;Forced=[bool]$IsForced})
    return $true
}
function Start-InventoryJob {throw 'A synthetic test must never launch a child.'}
'@))
$results=[Collections.Generic.List[object]]::new()
function Case([string]$Name,[scriptblock]$Body){
    & $module {Reset-Case}
    try {& $module $Body;$results.Add([pscustomobject]@{Name=$Name;Passed=$true})}
    catch {throw "${Name}: $($_.Exception.Message) $($_.ScriptStackTrace)"}
}
try {
    Case 'Late new owner skips and records the old occurrence once' {
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 0 -and $script:Rows.Count -eq 1) 'Missed occurrence reached launch gate.'
        Assert-Case ($script:Rows[0].Status -eq 'Skipped' -and $script:Rows[0].Occurrence -eq $script:Due) 'Wrong skipped occurrence/status.'
        Assert-Case ($script:Saves -eq 1 -and $script:Cleared -eq 1) 'Skip was not persisted or wait cleared.'
        Invoke-LaunchPhase $script:Acquired.AddMinutes(1)
        Assert-Case ($script:Rows.Count -eq 1 -and $script:Admitted.Count -eq 0) 'Missed occurrence repeated on next tick.'
    }
    Case 'Plan ID change with same owner preserves queued due work' {
        Update-OrchestratorElectedOwnership $script:Due.AddMinutes(-1)
        $script:ElectionPlan.PlanId='plan-b'
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Rows.Count -eq 0) 'Same-owner delayed occurrence was skipped.'
    }
    Case 'Next daily occurrence launches after old occurrence was skipped' {
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Update-OrchestratorElectedOwnership $script:Due.AddDays(1).AddSeconds(30)
        Invoke-LaunchPhase $script:Due.AddDays(1).AddSeconds(30)
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Admitted[0].Occurrence -eq $script:Due.AddDays(1)) 'Next day was suppressed.'
    }
    Case 'Ownership acquired exactly at an occurrence does not skip it' {
        Update-OrchestratorElectedOwnership $script:Due
        Invoke-LaunchPhase $script:Due.AddSeconds(30)
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Rows.Count -eq 0) 'Boundary occurrence was skipped.'
    }
    Case 'Loss and reacquisition reset the boundary' {
        Update-OrchestratorElectedOwnership $script:Due.AddMinutes(-1)
        $script:ElectionPlan.Assignments[0].OwnerServer='synthetic-other-server'
        Update-OrchestratorElectedOwnership $script:Due.AddMinutes(1)
        Invoke-LaunchPhase $script:Due.AddMinutes(1)
        Assert-Case ($script:Admitted.Count -eq 0 -and $script:ElectedOwnershipSince.Count -eq 0) 'Lost owner remained eligible.'
        $script:ElectionPlan.Assignments[0].OwnerServer=$env:COMPUTERNAME
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Rows.Count -eq 1 -and $script:Admitted.Count -eq 0) 'Reacquisition replayed an old occurrence.'
    }
    Case 'Temporary plan loss blocks launches without resetting ownership' {
        Update-OrchestratorElectedOwnership $script:Due.AddMinutes(-1)
        $plan=$script:ElectionPlan;$script:ElectionPlan=$null
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 0) 'Missing plan admitted a job.'
        $script:ElectionPlan=$plan
        Update-OrchestratorElectedOwnership $script:Acquired.AddMinutes(1)
        Invoke-LaunchPhase $script:Acquired.AddMinutes(1)
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Rows.Count -eq 0) 'Transient plan loss reset the boundary.'
    }
    Case 'RunOnce retains late catch-up' {
        $script:Job.Schedule.MissedRunPolicy='RunOnce'
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Rows.Count -eq 0) 'RunOnce was skipped.'
    }
    Case 'Pending automatic retry retains its occurrence' {
        $script:State.PendingRetry=@{NotBefore=$script:Acquired.AddMinutes(-1).ToString('o');ScheduledOccurrence=$script:Due.ToString('o');Attempt=1}
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Admitted[0].Occurrence -eq $script:Due -and $script:Rows.Count -eq 0) 'Retry was skipped.'
    }
    Case 'Force remains explicit' {
        $script:ForcedPending=@($script:Job.Name)
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Admitted[0].Forced -and $script:Rows.Count -eq 0) 'Force was skipped.'
    }
    Case 'Manual Pipeline retains its old requested occurrence' {
        $script:PipelinePending[$script:Job.Name]=@{Occurrence=$script:Due;Attempt=0;BatchId='synthetic'}
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Rows.Count -eq 0) 'Pipeline was skipped.'
    }
    Case 'Explicit maintenance catch-up remains eligible' {
        $script:MaintenanceDue=$script:Due
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:Rows.Count -eq 0) 'Maintenance catch-up was skipped.'
    }
    Case 'Peer already handled maintenance occurrence' {
        $script:MaintenanceDue=$script:Due
        $script:MaintenanceCoverage=[pscustomobject]@{Server='peer';Revision=12}
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 0 -and $script:Saves -eq 1 -and $script:State.MaintenanceCatchUpLastOccurrence -eq $script:Due.ToString('o')) 'Peer coverage did not persist the cursor.'
    }
    Case 'Unreadable peer state pauses catch-up' {
        $script:MaintenanceDue=$script:Due;$script:CoverageError=$true
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 0 -and $script:Saves -eq 0) 'Unreadable peer state admitted or consumed catch-up.'
    }
    Case 'Running job is not relaunched after reassignment' {
        $script:RunningJobs[$script:Job.Name]=@{ProcessId=123}
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 0 -and $script:RunningJobs.ContainsKey($script:Job.Name)) 'Running job was stopped or admitted again.'
    }
    Case 'Disabled job remains disabled' {
        $script:Job.Enabled=$false
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 0 -and $script:Rows.Count -eq 0) 'Disabled job was scheduled.'
    }
    Case 'Pinned job keeps its current scheduling behavior' {
        $script:Job.AssignmentMode='Pinned'
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case ($script:Admitted.Count -eq 1 -and $script:ElectedOwnershipSince.Count -eq 0) 'Pinned job was affected.'
    }
    Case 'Weekly Skip ignores old occurrence and admits the next week' {
        $script:Job.Schedule.Type='Weekly';$script:Job.Schedule.DaysOfWeek=@($script:Due.DayOfWeek.ToString())
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Invoke-LaunchPhase $script:Due.AddDays(7).AddSeconds(30)
        Assert-Case ($script:Rows.Count -eq 1 -and $script:Admitted.Count -eq 1 -and $script:Admitted[0].Occurrence -eq $script:Due.AddDays(7)) 'Weekly boundary failed.'
    }
    Case 'Multiple daily times preserve the next slot' {
        $script:Job.Schedule.Times=@('00:25','12:00')
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Invoke-LaunchPhase $script:Due.Date.AddHours(12).AddSeconds(30)
        Assert-Case ($script:Rows.Count -eq 1 -and $script:Admitted.Count -eq 1 -and $script:Admitted[0].Occurrence.Hour -eq 12) 'Next daily slot was skipped.'
    }
    Case 'Failed persistence ends the launch phase' {
        $script:FailSave=$true
        Update-OrchestratorElectedOwnership $script:Acquired
        Invoke-LaunchPhase $script:Acquired
        Assert-Case (-not $script:StatePersistenceHealthy -and $script:Admitted.Count -eq 0) 'Failed persistence admitted a launch.'
    }
    $text=Get-Content -LiteralPath $source -Raw
    if($text -notmatch 'Update-OrchestratorElectedOwnership -Now \(Get-Date\)\s+Invoke-MissedRunCatchUp'){throw 'Startup ownership seed is missing.'}
    if($text -notmatch 'Update-OrchestratorElectionPlan\s+Update-OrchestratorElectedOwnership -Now \(Get-Date\)\s+Update-OrchestratorJobStopRequests'){throw 'Tick ownership tracking is missing.'}
    $results | Format-Table -AutoSize
    Write-Output ("PASS: {0} synthetic scheduling scenarios; startup/tick wiring checked; no live execution." -f $results.Count)
} finally {Remove-Module $module -Force}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBYmjokz78qBAbT
# //XkkcQXxmiAHyHfdbDPNJSXjglCzaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDQVDZM8iniqm4RLxo4Mw8oeNSmaMSpjvrYdinP7tDAqMA0GCSqG
# SIb3DQEBAQUABIIBgAPqZjYJfpYyWeMOS2S0gD0n4cHqs9QYB2KlSko5A56AGn/C
# oc7isLnFqRN1hbzA3wiYlUkVE8Py2OE++o8FxsSTmjM7pTWrL0qap2Xk0IOOmFtQ
# O8k54VP67XkY9voKaxOl2NmdlUHCgYmMnDh+/IghVpYh5AUk0lJ4O1n/dCfTajeZ
# tEIyAkFA3hc0B3s+2DuOeQ8STCTMXrh+OFNMQAzQEah+ZVflyIJ8pntTRHB2J5Td
# XFVpcTPBzyzeCNd7SVyHzjvqxmSCK2baFrKVcptxl0l1hTNwpbluQtMkQjK5owMJ
# SRXQ5mLIqXvyeJnpcefglnvh/d+IZbsumVBQe4kl9zq42XchCUrGUPkGM2Y4Fie4
# EoBEXAtfPcxfEWhY2Poxb1Qdq5WZ+VpAghdz7VI9Novl7FGHbqwx02pLozIwwQ0i
# 7ng6fmBLIj5CnaYIwyGU43bs5S3ll6jZPHk/lHipar2pl2pe4lqc3iJ39ofK5YpC
# QUPb5w3rkvoxVYHZaKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExNTQ3
# MTZaMC8GCSqGSIb3DQEJBDEiBCB2Ls3YBJY37C7PHjYfMzjVc9IZptFQcHEtzUJ4
# RxMCpjANBgkqhkiG9w0BAQEFAASCAgCXJU2CUvE/bqc89s4iiBNUetsSy0qbiijm
# FAUy6C9/+dEqj7+XYji2Qtksxjis4ppaslSV7IKlpV8a0yQbLtBv0EEgFcKsLWEt
# 1POpIdq1Y+kMR56vkQpE2Fc7OwNdeH6tSiv7WtcK8vscmtv71JZWvhmEfz4qiF8c
# K7w0HKIAqtdjzbXU45UCl8u3MAXf1iq6+oRRSJqp84KCgU0zV/DTdDQVv8KpWaSR
# q5+QlKbDZWYjS0bGXHGrniuLv6MFLrzD27MNw782Fd1bMsuk02aoPmryCiuRWrwb
# b7PkmpMUo2F/EjFEGBM0Od6nU9K0x1Z2+L/qfYK7cupeGLNMBOQSNZRa5c8U4IFT
# blmW2tvUQc+3UsuApuM0fhrYpowcbwIbC5kTqeC0tCOfQ0LSKpWQ6OHJKH7AH8n+
# 7pWzwVfH3zP6WHKH6mO34BVqJlVJodLaSJdD/rzMzV39IPil5sqrqO0Au6lW4hZd
# eXgwHRgxjRoWhToSvJf9Evc0ojU2nF/1WmpJ+0KrP9LP7Uqrf1B9geZftr1dF6+K
# hUgcO/7jfPb6qrIeHOizY3ZJd7ZCeOH4paTpRBNjKoJDSIgc/vgHnzAyxLkeXd1w
# tzPC3dYusabEB6iwdFLiF9nIgDuH1618K4CPt2jTjkhT7rwQeoYwT7afT7epob6U
# Jahw76FmjA==
# SIG # End signature block
