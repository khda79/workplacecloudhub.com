<#
.SYNOPSIS
Offline tests for the elected RunOnce catch-up guard: a due occurrence already covered by a newer or later-started shared claim is not replayed.
.VERSION
1.1
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$base=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
$transport=Get-Module SmartM365.JsonTransport
$policy=& $transport {(Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock}
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'),[ref]$null,[ref]$null)
function Get-FunctionText([string]$Name){$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true).Extent.Text}
$defs=foreach($name in @('Get-OrchestratorOccurrenceCoveredByLaterClaim','Get-OrchestratorClaimOccurrenceUtc','ConvertTo-OrchestratorUtcTime')){Get-FunctionText $name}
$gate=New-Module -ScriptBlock ([scriptblock]::Create($defs -join "`n"))
$root=Join-Path $env:TEMP ('elected-catchup-'+[guid]::NewGuid().ToString('N'))
$claims=Join-Path $root 'Claims'
New-Item -ItemType Directory $claims|Out-Null
$script:passed=0
function Check($Value,$Message){if(-not $Value){throw $Message};$script:passed++}
try {
    & $transport {function script:Get-SmartM365JsonTransportPolicy {@{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()}}}
    & $gate {param($c) $script:Settings=[pscustomobject]@{ElectionClaimsPath=$c};$script:DependencyFreshCache=@{}} $claims

    # Incident shape: the schedule moved from 08:30 to 07:45; server A ran the 08:30 occurrence,
    # server B became owner and evaluates the previous day's 07:45 occurrence at 05:54.
    $job=[pscustomobject]@{Name='CatchUpJob';AssignmentMode='Elected';Schedule=[pscustomobject]@{Type='Daily';Times=@('07:45');DaysOfWeek=@()}}
    $now=[datetime]::new(2026,10,20,5,54,0,[DateTimeKind]::Local)
    $due=[datetime]::new(2026,10,19,7,45,0,[DateTimeKind]::Local)
    $previousSchedule=[datetime]::new(2026,10,19,8,30,0,[DateTimeKind]::Local)
    $folder=Join-Path $claims $job.Name
    function Coverage { & $gate {param($j,$o,$n) $script:DependencyFreshCache=@{}; Get-OrchestratorOccurrenceCoveredByLaterClaim -Job $j -ExpectedOccurrence $o -Now $n} $job $due $now }
    function Initialize-ClaimFolder { if(Test-Path $folder){Get-ChildItem $folder -File|Remove-Item -Force} }
    function Write-Claim([datetime]$Occurrence,[string]$Status,[datetime]$Created,[datetime]$Updated,[string]$JobName=$job.Name) {
        New-Item -ItemType Directory $folder -Force|Out-Null
        $document=[ordered]@{SchemaVersion=1;ClaimId=[guid]::NewGuid().ToString('N');JobName=$JobName;OccurrenceUtc=$Occurrence.ToUniversalTime().ToString('o');OwnerServer='SERVER-A';PlanId='synthetic';OrchestratorPid=1;Status=$Status;Attempt=0;CreatedAtUtc=$Created.ToUniversalTime().ToString('o');UpdatedAtUtc=$Updated.ToUniversalTime().ToString('o')}
        $path=Join-Path $folder ($Occurrence.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')
        [IO.File]::WriteAllText($path,($document|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
        $path
    }

    Check ($null -eq (Coverage)) 'Missing claims folder covered the due occurrence.'
    New-Item -ItemType Directory $folder -Force|Out-Null
    Check ($null -eq (Coverage)) 'Empty claims folder covered the due occurrence.'

    foreach($status in @('Success','CompletedWithWarnings')) {
        Initialize-ClaimFolder
        Write-Claim $previousSchedule $status $previousSchedule $previousSchedule.AddMinutes(40)|Out-Null
        $coverage=Coverage
        Check ($null -ne $coverage -and $coverage.Occurrence -eq $previousSchedule -and $coverage.Status -eq $status -and $coverage.OwnerServer -eq 'SERVER-A') "Newer $status claim did not cover the missed occurrence."
    }
    foreach($status in @('Failed','TimedOut','Interrupted','Cancelled','Claimed','Running','RetryScheduled')) {
        Initialize-ClaimFolder
        Write-Claim $previousSchedule $status $previousSchedule $previousSchedule.AddMinutes(40)|Out-Null
        Check ($null -eq (Coverage)) "Newer $status claim covered the missed occurrence."
    }

    # An older occurrence counts only when its run started at or after the due occurrence.
    $older=$due.AddDays(-1)
    Initialize-ClaimFolder
    Write-Claim $older 'Success' $due.AddMinutes(15) $due.AddMinutes(55)|Out-Null
    $coverage=Coverage
    Check ($null -ne $coverage -and $coverage.Occurrence -eq $older) 'Late catch-up started after the due occurrence did not cover it.'
    Initialize-ClaimFolder
    Write-Claim $older 'Success' $older $older.AddMinutes(40)|Out-Null
    Check ($null -eq (Coverage)) 'An older run started before the due occurrence covered it.'
    Initialize-ClaimFolder
    Write-Claim $older 'Success' $due.AddMinutes(-30) $due.AddMinutes(30)|Out-Null
    Check ($null -eq (Coverage)) 'A run in progress at the due occurrence covered it.'

    # The due occurrence itself is left to the existing peer claim check.
    Initialize-ClaimFolder
    Write-Claim $due 'Success' $due $due.AddMinutes(40)|Out-Null
    Check ($null -eq (Coverage)) 'The claim of the due occurrence itself was reported as a later coverage.'

    # Claims that cannot be trusted never cover.
    Initialize-ClaimFolder
    Write-Claim $now.AddHours(2) 'Success' $now.AddHours(2) $now.AddHours(3)|Out-Null
    Check ($null -eq (Coverage)) 'A claim for a future occurrence covered the due occurrence.'
    Initialize-ClaimFolder
    Write-Claim $previousSchedule 'Success' $previousSchedule $now.AddHours(1)|Out-Null
    Check ($null -eq (Coverage)) 'A claim completed in the future covered the due occurrence.'
    Initialize-ClaimFolder
    Write-Claim $previousSchedule 'Success' $previousSchedule $previousSchedule.AddMinutes(-5)|Out-Null
    Check ($null -eq (Coverage)) 'A claim completed before it started covered the due occurrence.'
    Initialize-ClaimFolder
    Write-Claim $previousSchedule 'Success' $previousSchedule $previousSchedule.AddMinutes(40) 'OtherJob'|Out-Null
    Check ($null -eq (Coverage)) 'A claim of another job covered the due occurrence.'
    Initialize-ClaimFolder
    $corrupt=Write-Claim $previousSchedule 'Success' $previousSchedule $previousSchedule.AddMinutes(40)
    [IO.File]::WriteAllText($corrupt,'{')
    Check ($null -eq (Coverage)) 'An unreadable claim covered the due occurrence.'
    # A trusted claim still covers when a newer claim is unreadable.
    Write-Claim $due.AddMinutes(30) 'Success' $due.AddMinutes(30) $due.AddMinutes(70)|Out-Null
    $coverage=Coverage
    Check ($null -ne $coverage -and $coverage.Occurrence -eq $due.AddMinutes(30)) 'An unreadable newer claim hid a valid covering claim.'

    # Without coverage the next occurrence is unaffected: normal 07:45 run after the incident day.
    Initialize-ClaimFolder
    Write-Claim $previousSchedule 'Success' $previousSchedule $previousSchedule.AddMinutes(40)|Out-Null
    $today=[datetime]::new(2026,10,20,7,45,0,[DateTimeKind]::Local)
    $todayCoverage=& $gate {param($j,$o,$n) $script:DependencyFreshCache=@{}; Get-OrchestratorOccurrenceCoveredByLaterClaim -Job $j -ExpectedOccurrence $o -Now $n} $job $today $today.AddSeconds(20)
    Check ($null -eq $todayCoverage) 'The next scheduled occurrence was covered by the previous day.'

    # The launch phase applies the guard to due elected occurrences before dependency gates.
    $launch=Get-FunctionText 'Invoke-LaunchPhase'
    $guardIndex=$launch.IndexOf('Get-OrchestratorOccurrenceCoveredByLaterClaim')
    Check ($guardIndex -gt 0 -and $guardIndex -lt $launch.IndexOf('# Dependency gate')) 'Launch phase does not apply the covered-occurrence guard before dependency gates.'
    Check ($launch -match "\`$reason -eq 'due' -and \`$job\.AssignmentMode -eq 'Elected'") 'Covered-occurrence guard is not limited to due elected occurrences.'
    Check ($launch.Substring($guardIndex,1200) -match 'Set-OccurrenceSkipped') 'Covered occurrence is not recorded as Skipped.'

    # A momentarily unreachable share must mean "not covered" with a warning, never a scheduler crash.
    $launchAst=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-LaunchPhase'},$true)
    $guardIf=$launchAst.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -match "\`$reason -eq 'due' -and \`$job\.AssignmentMode -eq 'Elected'" -and $n.Extent.Text.Contains('Get-OrchestratorOccurrenceCoveredByLaterClaim')},$true)
    Check ($null -ne $guardIf) 'Covered-occurrence guard block not found in the launch phase.'
    $harness=New-Module -ScriptBlock ([scriptblock]::Create(($defs -join "`n") + @"

function Get-ChildItem { throw 'Synthetic share outage.' }
function Write-OrchestratorRuntimeUpdateWarning { param(`$Key,`$Message,`$Now) `$script:warnings.Add(`$Message) }
function Set-OccurrenceSkipped { param(`$JobName,`$Occurrence,`$Reason) `$script:skipped=`$true }
function Clear-DependencyWaitLog { param(`$JobName) }
function Invoke-Guard { param(`$job,`$occurrence,`$Now)
    `$name=`$job.Name; `$reason='due'; `$script:reachedLaunch=`$false
    foreach(`$tick in 1) {
$($guardIf.Extent.Text)
        `$script:reachedLaunch=`$true
    }
}
"@))
    try {
        $outage=& $harness {param($c,$j,$o,$n) $script:Settings=[pscustomobject]@{ElectionClaimsPath=$c};$script:DependencyFreshCache=@{};$script:warnings=[Collections.Generic.List[string]]::new();$script:skipped=$false
            Invoke-Guard -job $j -occurrence $o -Now $n
            [pscustomobject]@{Reached=$script:reachedLaunch;Skipped=$script:skipped;Warnings=@($script:warnings)}} $claims $job $due $now
        Check ($outage.Reached -and -not $outage.Skipped) 'A share outage skipped the due occurrence or stopped the launch phase.'
        Check ($outage.Warnings.Count -eq 1 -and $outage.Warnings[0] -match 'Synthetic share outage') 'A share outage was not reported as a warning.'
    } finally { Remove-Module $harness -Force }
    [pscustomobject]@{Passed=$script:passed;Fixture=$root;Scope='Synthetic claims only; no production jobs or network access'}
} finally {
    & $transport {param($p) Set-Item Function:script:Get-SmartM365JsonTransportPolicy $p} $policy
    Remove-Module $gate -Force
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA9Fn4wl7vnOmy9
# Tf4rYJ4zXTkbAyu4ER8qq22obXgaG6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPk+fPGaMCDUcRVxfSwTfvKJZhiSbCC0rAp4Xgx0qU6lMA0GCSqG
# SIb3DQEBAQUABIIBgHeZETJn9d2GnIBzV0DPk5VcBaF41eG/vRnxEhO3w9Aj6Tm7
# 8Gl1c6LN8RKhsrnXyLvwJm5j8dgHI+H/Us+s5WFjkXk9vU6gnHrUDvWqFvqrXTXR
# PxFawARhwMZ82taHqYMEyDb0aRojF7JdY10uBMr3b7p2O54X7F5fO8TOR2TAq33b
# snesjk06JxOTQiAcWQdEQzRyPpkNSQVBIOE55cElWcZywJCERfmXDgf3QABgO+jP
# lk09R5BfJg6MozPtmn9iWqy5VWNDPMIHL4Pdkh3FOtg6UV/Cg/n+qp9RgZzBau0d
# p9DTFsG0S67S5CY2z88GdDzxGY4Px860pYOK4J3s/nYJos3UtR5kETetsIJ+QoI4
# dGKw/vEU7QziWD51qV7GMS9pXau+kPWfKyQGTq4yEwDKYdM8XzDlul18Idqv2ojF
# G0eq6+t0sf8E4PHQ5fRpMj0WiGHMN7K3ls4wKCrjbpgebzKvKqn0IPw8i18yl0UR
# W7+lJFBOFQpvtyy7oqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMDEz
# MTFaMC8GCSqGSIb3DQEJBDEiBCAn8WKsTDbhl8RYbc9YAoYFhcYVdJsw0y1X9uN2
# otRlyjANBgkqhkiG9w0BAQEFAASCAgCLtxUsD3z47LMBcTusCoCbDZC4efkHjlIQ
# eMEU5tJaKDJfSHTgvum0oDebJRqOLY16lwjxpaXUXVIxs8xZoiXw2CD0pu9Vy4Mf
# WrsIb15Ui7DUVlfIAdrPwDTumPV3kmcc90XN1DEP/8jSVYWYApVOXtxr/ztLhDoE
# 8+uR3GgJC6YmXZemXcU8kJKThXIx3BZUuK8G9HPnUYl6AjMILmiTevd1LfhPZMRI
# JRNOuRX9+O64x2uJghspE8boQFF5jEhFYtB79PmeqR4ZdmDQSOe5jHaczjqMrsWY
# L0LSYdreYN1gSuT1MyYs5E3dpxKok+aeYsmjAMsmtgHuRAesuPiRO+mGkDr+qqWX
# xBRMdjga1WhES29h4L323yAGrhv0lmY/w9CHBgoJ4bGAyv4glb4+v10V9m989y9w
# UkxmSswP1k576J0ML4Vqn0vJvR0nR50zgYlBtzZPxtyu2tQg4NCMcs6fQpQ+PfuS
# 5K5PN2qKYn6ONV4+yURCTO22tVwXLCflp77sV2bMMgcRLbo9HccQcPUGCY1Oaj6F
# cZESuSMZnCbYuWp3lyzvORwq8O0yvYBCMvEx1Kd1ESjau6a1UEtRncpvAwUOBmNT
# /7Lbra+f+QuRFzzR+FJTd9qx2cw1JiCB4290b4aRELTVUhM+ya2e8Ou65G9jOiHJ
# QxNa/gNyDQ==
# SIG # End signature block
