<#
.SYNOPSIS
Offline tests for the FreshSuccess dependency activity gate: an active dependency expected to end
before the downstream occurrence blocks it; a longer staged run does not.
.VERSION
1.0.0
.NOTES
Synthetic claims and run states under TEMP only. No collector, scheduler, tenant or network access.
#>
#requires -Version 7.0
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic tenant key read by the receipt fallback; the test process ends afterwards.')]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
$transport = Get-Module SmartM365.JsonTransport
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$null, [ref]$null)
$defs = foreach ($name in @('Get-JobOccurrencesInWindow', 'Get-LatestPastOccurrence', 'Get-OrchestratorDependencyActivity', 'Get-OrchestratorDependencyMaxAge', 'Get-OrchestratorDependencyFreshStatus', 'Get-OrchestratorClaimOccurrenceUtc', 'ConvertTo-OrchestratorUtcTime')) {
    $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true).Extent.Text
}
$gate = New-Module -ScriptBlock ([scriptblock]::Create($defs -join "`n"))
$root = Join-Path $env:TEMP ('dependency-activity-' + [guid]::NewGuid().ToString('N'))
$shared = Join-Path $root 'DATA-ALL/Orchestrator'
$claims = Join-Path $shared 'Election/Claims'
$latest = Join-Path $root 'DATA-LAST'
New-Item -ItemType Directory $claims, $latest | Out-Null
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$managedApps = [pscustomobject]@{ Name = 'Intune-ManagedApps-Inventory'; ScriptPath = 'M365Inventory\WorkplaceScope\SmartM365-IntuneManagedApps-Inventory.ps1'; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 240; EstimatedDurationMinutes = 60; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('05:00'); DaysOfWeek = @() } }
$apps = [pscustomobject]@{ Name = 'Intune-DiscoveredApps-Inventory'; ScriptPath = 'M365Inventory\IntuneInventory\Applications\SmartM365-Intune-DiscoveredApps-Inventory.ps1'; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 8640; EstimatedDurationMinutes = 1500; Schedule = [pscustomobject]@{ Type = 'Weekly'; Times = @('20:00'); DaysOfWeek = @('Saturday') } }
$fast = [pscustomobject]@{ Name = 'EXO-Mailboxes-Inventory-Fast'; ScriptPath = 'ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1'; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 1380; EstimatedDurationMinutes = 90; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('00:15'); DaysOfWeek = @() } }
$late = [pscustomobject]@{ Name = 'Late-Producer'; ScriptPath = 'X\SmartM365-Late-Producer.ps1'; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 60; EstimatedDurationMinutes = 5; Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('07:40'); DaysOfWeek = @() } }
$prepare = [pscustomobject]@{ Name = 'WorkplaceEvidence-Prepare'; ScriptPath = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'; Enabled = $true; AssignmentMode = 'Elected'; DependencyMode = 'FreshSuccess'; DependencyMaxAgeHours = 48; DependsOn = @('Intune-ManagedApps-Inventory'); Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('07:45'); DaysOfWeek = @() } }
$jobs = @($managedApps, $apps, $fast, $late, $prepare)

# Claims are written directly so that their safe window follows the simulated clock, not the real one.
function Write-SyntheticClaim($Job, [datetime]$Occurrence, [string]$Status, $SafeUntil = $null) {
    $folder = Join-Path $claims $Job.Name
    New-Item -ItemType Directory $folder -Force | Out-Null
    $until = if ($null -eq $SafeUntil) { $Occurrence.AddDays(10) } else { [datetime]$SafeUntil }
    $document = [ordered]@{ SchemaVersion = 1; ClaimId = [guid]::NewGuid().ToString('N'); JobName = $Job.Name; OccurrenceUtc = $Occurrence.ToUniversalTime().ToString('o'); OwnerServer = 'REMOTE'; PlanId = 'synthetic'; OrchestratorPid = 1; Status = $Status; Attempt = 0; CreatedAtUtc = $Occurrence.ToUniversalTime().ToString('o'); UpdatedAtUtc = $Occurrence.ToUniversalTime().AddMinutes(1).ToString('o'); SafeUntilUtc = $until.ToUniversalTime().ToString('o') }
    $path = Join-Path $folder ($Occurrence.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ') + '.json.txt')
    [IO.File]::WriteAllText($path, ($document | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
}
function Write-SyntheticRunState($Job, [string]$Status, $Staged) {
    $scriptName = [IO.Path]::GetFileName($Job.ScriptPath)
    $document = [ordered]@{ Owner = 'SmartInventory-SourceRun'; ContractVersion = '1.0'; Producer = $scriptName; Status = $Status; RunId = [guid]::NewGuid().ToString('N'); StartedAtUtc = [datetime]::UtcNow.ToString('o'); PublicationStarted = $false; UnqualifiedPublication = $false }
    if ($null -ne $Staged) { $document.StagedPublication = $Staged }
    $path = Join-Path $latest ('SmartInventory_{0}.run.json.txt' -f [IO.Path]::GetFileNameWithoutExtension($scriptName))
    [IO.File]::WriteAllText($path, ($document | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
}
function Activity($Job, [datetime]$Downstream, [datetime]$Now) {
    & $gate { param($j, $d, $n) Get-OrchestratorDependencyActivity -DependencyJob $j -DownstreamOccurrence $d -Now $n } $Job $Downstream $Now
}

try {
    & $transport { function script:Get-SmartM365JsonTransportPolicy { @{ Mode = 'JsonText'; QualifiedUncRoots = @(); QualifiedLocalDataRoots = @() } } }
    & $gate { param($c, $s, $j) $script:Settings = [pscustomobject]@{ ElectionClaimsPath = $c; SharedDataFolderPath = $s; ElectionClaimGraceMinutes = 15 }; $script:Manifest = [pscustomobject]@{ OrderedJobs = $j }; $script:DependencyFreshCache = @{} } $claims $shared $jobs
    $global:SmartM365TenantKey = 'synthetic'

    # Weekday: ManagedApps retry at 05:58 blocks the 07:45 preparation, staged or not.
    $monday = [datetime]::new(2026, 10, 19, 0, 0, 0, [DateTimeKind]::Local)
    $d0745 = $monday.AddHours(7).AddMinutes(45)
    Write-SyntheticClaim $managedApps $monday.AddDays(-1).AddHours(5) 'Success'
    foreach ($status in @('Claimed', 'Running', 'RetryScheduled')) {
        Write-SyntheticClaim $managedApps $monday.AddHours(5) $status
        foreach ($staged in @($false, $true)) {
            Write-SyntheticRunState $managedApps 'Collecting' $staged
            $result = Activity $managedApps $d0745 $monday.AddHours(5).AddMinutes(58)
            Check ($result.State -eq 'Waiting' -and $result.Detail -match 'expected to finish before') "Morning dependency $status (staged=$staged) did not block the 07:45 preparation."
        }
    }
    # Final failure today with a fresh success yesterday: activity clear, freshness accepts yesterday.
    Write-SyntheticClaim $managedApps $monday.AddHours(5) 'Failed'
    Check ((Activity $managedApps $d0745 $monday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A final failure still blocked preparation.'
    $fresh = & $gate { param($j, $n) $script:DependencyFreshCache = @{}; Get-OrchestratorDependencyFreshStatus -DependencyJob $j -MaxAgeHours (Get-OrchestratorDependencyMaxAge -DependencyJob $j -OverrideHours 48 -Now $n) -Now $n } $managedApps $monday.AddHours(7).AddMinutes(45)
    Check ($fresh -eq 'Ready') 'Yesterday success was not accepted after a final failure today.'
    Write-SyntheticClaim $managedApps $monday.AddHours(5) 'Success'
    Check ((Activity $managedApps $d0745 $monday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A success today still blocked preparation.'
    Write-SyntheticClaim $managedApps $monday.AddHours(5) 'CompletedWithWarnings'
    Check ((Activity $managedApps $d0745 $monday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A success with warnings today still blocked preparation.'

    # Catch-up of the previous day occurrence while today's run is active: staged only is safe.
    $tuesday = $monday.AddDays(1)
    Write-SyntheticClaim $managedApps $tuesday.AddHours(5) 'Running'
    Write-SyntheticRunState $managedApps 'Collecting' $false
    Check ((Activity $managedApps $d0745 $tuesday.AddHours(5).AddMinutes(56)).State -eq 'Waiting') 'A direct-publication run did not block a catch-up preparation.'
    Write-SyntheticRunState $managedApps 'Collecting' $null
    Check ((Activity $managedApps $d0745 $tuesday.AddHours(5).AddMinutes(56)).State -eq 'Waiting') 'A run state without the staging flag was trusted.'
    Write-SyntheticRunState $managedApps 'Collecting' $true
    Check ((Activity $managedApps $d0745 $tuesday.AddHours(5).AddMinutes(56)).State -eq 'Clear') 'A staged run blocked a catch-up preparation.'
    Write-SyntheticRunState $managedApps 'Publishing' $true
    $result = Activity $managedApps $d0745 $tuesday.AddHours(5).AddMinutes(56)
    Check ($result.State -eq 'Waiting' -and $result.Detail -match 'promoting') 'A staged run promoting its CSVs did not block.'
    Write-SyntheticClaim $managedApps $tuesday.AddHours(5) 'Success'

    # Saturday 07:45: DiscoveredApps is due at 20:00 only; the previous weekly run succeeded.
    $saturday = [datetime]::new(2026, 10, 24, 0, 0, 0, [DateTimeKind]::Local)
    Write-SyntheticClaim $apps $saturday.AddDays(-7).AddHours(20) 'Success'
    $result = Activity $apps $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(45)
    Check ($result.State -eq 'Clear') 'The Saturday 20:00 weekly run blocked the Saturday morning preparation.'

    # Sunday 07:45: the 25-hour weekly run is in progress.
    $sunday = $saturday.AddDays(1)
    Write-SyntheticClaim $apps $saturday.AddHours(20) 'Running'
    Write-SyntheticRunState $apps 'Collecting' $true
    $result = Activity $apps $sunday.AddHours(7).AddMinutes(45) $sunday.AddHours(7).AddMinutes(45)
    Check ($result.State -eq 'Clear' -and $result.Detail -match 'staged') 'A staged long weekly run blocked the Sunday preparation.'
    Write-SyntheticRunState $apps 'Collecting' $false
    Check ((Activity $apps $sunday.AddHours(7).AddMinutes(45) $sunday.AddHours(7).AddMinutes(45)).State -eq 'Waiting') 'A direct-publication long run did not block.'
    Write-SyntheticRunState $apps 'Collecting' $true

    # An expired non-terminal claim (lost owner) does not block; stale archives are ignored.
    Write-SyntheticClaim $apps $saturday.AddHours(20) 'Running' $sunday.AddHours(6)
    Write-SyntheticRunState $apps 'Collecting' $false
    Check ((Activity $apps $sunday.AddHours(7).AddMinutes(45) $sunday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'An expired claim blocked preparation.'
    $folder = Join-Path $claims $apps.Name
    $archive = Join-Path $folder ($sunday.ToUniversalTime().AddHours(1).ToString('yyyyMMddTHHmmssfffZ') + '.json.stale.abc.json.txt')
    [IO.File]::WriteAllText($archive, '{"JobName":"Intune-DiscoveredApps-Inventory","Status":"Running"}')
    Check ((Activity $apps $sunday.AddHours(7).AddMinutes(45) $sunday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A stale claim archive was read as an execution.'

    # Saturday EXO Fast queued behind the full run since 00:15: not started beyond the grace.
    Check ((Activity $fast $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(45)).State -eq 'Clear') 'A queued dependency held preparation all Saturday.'
    # A dependency due shortly before the downstream occurrence and not yet claimed blocks during the grace.
    $result = Activity $late $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(46)
    Check ($result.State -eq 'Waiting' -and $result.Detail -match 'not started') 'A due, unclaimed dependency did not block during the claim grace.'
    Check ((Activity $late $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(56)).State -eq 'Clear') 'An unclaimed dependency blocked beyond the claim grace.'

    # A script shared by several jobs cannot prove staging through its single run state.
    $full = [pscustomobject]@{ Name = 'EXO-Mailboxes-Inventory'; ScriptPath = $fast.ScriptPath; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 2160; EstimatedDurationMinutes = 1200; Schedule = [pscustomobject]@{ Type = 'Weekly'; Times = @('20:00'); DaysOfWeek = @('Friday') } }
    & $gate { param($j) $script:Manifest = [pscustomobject]@{ OrderedJobs = $j } } (@($jobs) + $full)
    Write-SyntheticClaim $full $saturday.AddDays(-1).AddHours(20) 'Running'
    Write-SyntheticRunState $full 'Collecting' $true
    Check ((Activity $full $saturday.AddHours(7).AddMinutes(45) $saturday.AddHours(7).AddMinutes(45)).State -eq 'Waiting') 'A shared-script run state was trusted.'

    # Run state reader.
    Write-SyntheticRunState $apps 'Collecting' $true
    $state = Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $shared -Job $apps
    Check ($state.Known -and $state.StagedPublication -and $state.Status -eq 'Collecting') 'Staged run state not read.'
    Check (-not (Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $shared -Job $prepare).Known) 'Absent run state reported as known.'
    Check (-not (Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath (Join-Path $root 'Other') -Job $apps).Known) 'Run state read outside the DATA-ALL/Orchestrator layout.'

    # GUI readiness explains the same wait.
    Write-SyntheticClaim $managedApps $sunday.AddHours(5) 'RetryScheduled'
    Write-SyntheticRunState $managedApps 'Failed' $false
    $readiness = @(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $shared -JobsDocument ([pscustomobject]@{ Jobs = $jobs }) -JobName $prepare.Name -Now $sunday.AddHours(7))
    Check ($readiness.Count -eq 1 -and $readiness[0].State -eq 'Waiting' -and $readiness[0].Detail -match 'RetryScheduled') 'GUI readiness ignored a dependency awaiting retry.'
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Output "PASS: $script:passed dependency activity checks. Synthetic claims only."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC2YqdujI3qRsTe
# IKmJdtC+l2s5l6AxbuXnmGtn0G3AC6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM3LWo3+AnSl0DeYdjz6bT8Xd1tn+Kpz571JXNY+altZMA0GCSqG
# SIb3DQEBAQUABIIBgIXZzzQf8P09cKXZcceuXhpHTjo7frldXUh79wLddonB+00i
# ue9lbj3gWjB4LKJ33R+nfPS5es9YkXNPAUMTRQFwefwYnSy5xXR3kvqgFFqJXKY1
# Iflth9QI/wjTbLyJWbWghR3BaR4nSBG5RkyyUIqZMFGF4UDB7GqyS8wxM9KkFIVR
# sH+uXWujOLYzsC25a+AoP37Jw2z2EZ3VPvwF8thRkmuw5Nsq313rYO+1BsKDwqtM
# 8pl1WQbyMBjNx5L+zHvLTEEZe3UCm9oS+2BckGfdqb8mllS2Hfr7auFOwrnxpSYu
# o4YVue9da9h8CG7OsaplAcsp3S8EVnxPiAxXaIhlObXkIEPjtZAXp8MlHSXtlT/D
# FEwEiJQsU1RPBKAJdYc3KyDIqIOV9cLpijFM+4qpVfCF8/VtF6Vp3dROxgXWH3ch
# Dsbfcl1gBIljdDOLdKMfBWTxvpcBzu711AaHmBtBMC10EobLPYXB1lcz1YI7DcJj
# cWDUCtoaQURTz0h6SqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAwNjIx
# NDlaMC8GCSqGSIb3DQEJBDEiBCCZ3z8KgxaRNvlp/ERFXvztBmRBm66/rtPhr0/v
# yUYYsTANBgkqhkiG9w0BAQEFAASCAgCKRT+3DhzloB5YJZM7nCoNvGpVLC6xaedS
# /hL7FbvorkxMY2kf5dB/6S7+uUEJaPRdW6FdZFGaJlz98c0S2Eyjmku33ca8BpRr
# wyIP6Se2No+TR2nT+WMs6Gg9JwwoL0YZzMtiXQqh2rP9TR5B1VEJ6DX6QwEa4AOk
# OyTgJjEfwDHpusuU8Y859l6DAXsbvvvNOm9eRCzYfaRjtKe4odVktAyEcr2bbdzp
# 3UwKPs6kCFpp8CUGDeJK4pyjr2wDUIHOQYszmhbATUveaa2vFM7aEyj6YVGX67cM
# mQRj94s0m2Nh6/eBAOoeyc4SKu7uzwbMmewrq648A0nY3PS81Ai2FsCMGjBGIZdg
# Vu1PVNp+H9y6hMCJudKO9bhE9FCCN2p0VMZeKb6NyaH+oI+R0Qp3w67Y1MdmiDr4
# 6U1Rq2JZ2Ct5b0SLbSirS2MPeIwII2SIvDcwV+feM8kbGqhlmo9TNXHvKLOcS09q
# w2j1He97qWNodLbnVP7B4iDs+IHNMrHmQvHDI/frKkzIKX+dkMojdUFNopZawrta
# hJMJ7Sp3kQ02YY4NbxNhOl3N01SNBpCJmWCCkm0LtUr9BjGbwAGTMhdAqSclRGwy
# Bd3tS7SsnkVZ03k3Fip9wMEb6IP1ys8uRrtsEePbZ7woBu5sR4potkyEJV7pOW1I
# HZdBb8lYBg==
# SIG # End signature block
