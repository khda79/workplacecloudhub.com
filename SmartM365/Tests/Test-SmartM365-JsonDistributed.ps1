#requires -Version 7.0
<#
.SYNOPSIS
Offline JSON transport tests for distributed orchestrator state.
.VERSION
1.0.2
#>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$modulePath=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1'
Import-Module $modulePath -Force
$distributed=Get-Module SmartM365.Orchestrator.Distributed
$transport=Get-Module SmartM365.JsonTransport
$policy=& $transport { (Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock }
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonDistributed-' + [guid]::NewGuid().ToString('N'))
$claims=Join-Path $root 'Claims';$leases=Join-Path $root 'Concurrency'
$null=New-Item -ItemType Directory $claims,$leases -Force
$script:passed=0
function Check([bool]$Value,[string]$Message){if(!$Value){throw $Message};$script:passed++}
function Reject([scriptblock]$Action,[string]$Message){$failed=$false;try{& $Action|Out-Null}catch{$failed=$true};Check $failed $Message}
try {
    & $transport { function script:Get-SmartM365JsonTransportPolicy { @{Mode='Readers';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} } }
    $occurrence=[datetime]::UtcNow
    $claim=Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName Fixture -Occurrence $occurrence -OwnerServer SYNTHETIC-A -PlanId synthetic
    $lease=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey fixture -JobName Fixture -Occurrence $occurrence -OwnerServer SYNTHETIC-A
    $claimHash=(Get-FileHash $claim.ClaimPath).Hash;$leaseHash=(Get-FileHash $lease.LeasePath).Hash
    & $transport { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} } }
    & $distributed { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} } }
    Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases
    $readClaim=Get-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName Fixture -Occurrence $occurrence
    $readLease=Get-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey fixture
    Check ($readClaim.Claim.ClaimId -eq $claim.Claim.ClaimId -and (Get-FileHash $readClaim.ClaimPath).Hash -eq $claimHash) 'Claim migration changed identifier or content.'
    Check ($readLease.Lease.LeaseId -eq $lease.Lease.LeaseId -and (Get-FileHash $readLease.LeasePath).Hash -eq $leaseHash) 'Lease migration changed identifier or content.'
    Check (!(Test-Path $claim.ClaimPath) -and !(Test-Path $lease.LeasePath)) 'Distributed migration retained legacy outputs.'
    & $distributed {
        $script:OriginalResolver = (Get-Command Resolve-DistributedJsonPath).ScriptBlock
        $script:ResolvedPaths = [Collections.Generic.List[string]]::new()
        function script:Resolve-DistributedJsonPath {
            param([string]$Path)
            $script:ResolvedPaths.Add($Path)
            & $script:OriginalResolver -Path $Path
        }
    }
    $progress = [Collections.Generic.List[string]]::new()
    Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases -OnProgress {param($message) $progress.Add($message)}
    $resolved = @(& $distributed {$script:ResolvedPaths.ToArray()})
    Check ($resolved.Count -eq 1 -and $resolved[0] -eq $lease.LeasePath) 'Completed claim was reprocessed or live lease skipped.'
    Check ($progress.Count -ge 2 -and $progress[-1] -match 'scan complete') 'Startup progress missing.'
    $claimJournal=(Get-SmartM365JsonNames $claim.ClaimPath).Journal
    $journalBytes=[IO.File]::ReadAllBytes($claimJournal)
    $claimBytes=[IO.File]::ReadAllBytes($readClaim.ClaimPath)
    [IO.File]::AppendAllText($claimJournal,(@{Owner='Orchestrator distributed state';Phase='Prepared';SHA256=$claimHash}|ConvertTo-Json -Compress)+[Environment]::NewLine)
    & $distributed {$script:ResolvedPaths.Clear()}
    Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases
    Check (($claim.ClaimPath -in @(& $distributed {$script:ResolvedPaths.ToArray()})) -and (Get-Content $claimJournal -Tail 1 | ConvertFrom-Json).Phase -eq 'Completed') 'Interrupted rename was not resumed.'
    Check ((Get-FileHash $readClaim.ClaimPath).Hash -eq $claimHash) 'Recovery changed historical bytes.'
    [IO.File]::AppendAllText($claimJournal,'{"Phase":')
    Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases -WarningAction SilentlyContinue
    Check ((Get-FileHash $readClaim.ClaimPath).Hash -eq $claimHash) 'Truncated receipt recovery changed claim.'
    [IO.File]::WriteAllBytes($claimJournal,$journalBytes)
    Copy-Item $readClaim.ClaimPath $claim.ClaimPath
    Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases
    Check (!(Test-Path $claim.ClaimPath)) 'Identical duplicate was skipped.'
    [IO.File]::WriteAllBytes($claim.ClaimPath,$claimBytes)
    [IO.File]::WriteAllText($readClaim.ClaimPath,'{"different":true}')
    Reject {Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases} 'Divergent pair was skipped.'
    [IO.File]::WriteAllBytes($readClaim.ClaimPath,$claimBytes)
    Remove-Item -LiteralPath $claim.ClaimPath
    [IO.File]::WriteAllBytes($claimJournal,$journalBytes)
    [IO.File]::WriteAllText($readClaim.ClaimPath,'{')
    Reject {Get-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName Fixture -Occurrence $occurrence} 'Invalid completed claim accepted by consumer.'
    [IO.File]::WriteAllBytes($readClaim.ClaimPath,$claimBytes)
    [IO.File]::WriteAllBytes($claimJournal,$journalBytes)
    # A natively produced JSON-text claim has no migration receipt to repair.
    Remove-Item -LiteralPath $claimJournal
    & $distributed {$script:ResolvedPaths.Clear()}
    Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases
    Check ($claim.ClaimPath -notin @(& $distributed {$script:ResolvedPaths.ToArray()})) 'Native JSON-text claim unnecessarily migrated.'
    [IO.File]::WriteAllBytes($claimJournal,$journalBytes)
    & $distributed {Set-Item Function:script:Resolve-DistributedJsonPath $script:OriginalResolver}
    # Startup race seen in production (1.1.18): a peer releases a live lease between the listing and its read.
    $vanishing=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey vanishing -JobName Vanishing -Occurrence $occurrence -OwnerServer SYNTHETIC-B
    $vanishingPath=(Get-SmartM365JsonNames $vanishing.LeasePath).Legacy
    & $distributed {
        param($Target)
        $script:VanishTarget=$Target; $script:VanishDelete=$true
        function script:Resolve-DistributedJsonPath {
            param([string]$Path)
            if ($Path -ne $script:VanishTarget) { return & $script:OriginalResolver -Path $Path }
            if ($script:VanishDelete) { Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($Path)) -File | Where-Object { $_.Name -like ([IO.Path]::GetFileName($Path) + '*') } | Remove-Item -Force }
            throw [IO.FileNotFoundException]::new('Could not find file ' + $Path)
        }
    } $vanishingPath
    $progress.Clear()
    Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases -OnProgress {param($message) $progress.Add($message)}
    Check (@($progress | Where-Object { $_ -match 'disappeared after the listing' }).Count -eq 1 -and $progress[-1] -match 'vanished=1') 'A lease released during the startup scan stopped the orchestrator or was not reported.'
    & $distributed {$script:VanishTarget=''}
    $kept=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey vanishing -JobName Vanishing -Occurrence $occurrence -OwnerServer SYNTHETIC-B
    & $distributed {$script:VanishTarget=$args[0]; $script:VanishDelete=$false} (Get-SmartM365JsonNames $kept.LeasePath).Legacy
    Reject {Convert-SmartM365OrchestratorDistributedHistory -ClaimsRootPath $claims -LeasesRootPath $leases} 'A read failure on a lease that still exists was ignored.'
    & $distributed {Set-Item Function:script:Resolve-DistributedJsonPath $script:OriginalResolver}
    Check (Exit-SmartM365OrchestratorConcurrencyLease -LeasePath $kept.LeasePath -LeaseId $kept.Lease.LeaseId -OwnerServer SYNTHETIC-B) 'Race fixture lease could not be released.'
    $competitor=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey fixture -JobName Other -Occurrence $occurrence -OwnerServer SYNTHETIC-B
    Check (!$competitor.Acquired) 'Migration allowed a competing lease.'
    $null=Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $claim.ClaimPath -OwnerServer SYNTHETIC-A -Status Success
    Check ((Get-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName Fixture -Occurrence $occurrence).Claim.Status -eq 'Success') 'Persisted legacy path reference did not resolve after migration.'
    Check (Exit-SmartM365OrchestratorConcurrencyLease -LeasePath $lease.LeasePath -LeaseId $lease.Lease.LeaseId -OwnerServer SYNTHETIC-A) 'Lease could not be released using persisted legacy path.'
    Check (!(Test-Path $readLease.LeasePath)) 'Released preferred lease remains.'
    foreach ($cycle in 1..2) {
        $again=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey fixture -JobName Fixture -Occurrence $occurrence.AddMinutes($cycle) -OwnerServer SYNTHETIC-A
        Check $again.Acquired 'Released key could not be acquired again.'
        if ($cycle -eq 1) {
            # Journal left by 1.1.6 after refusing a release of a reused key.
            $journal=(Get-SmartM365JsonNames $again.LeasePath).Journal
            [IO.File]::AppendAllText($journal,(@{Owner='Orchestrator distributed state';Phase='Failed';SHA256='';Detail='Migration journal belongs to another owner.'}|ConvertTo-Json -Compress)+[Environment]::NewLine)
        }
        Check (Exit-SmartM365OrchestratorConcurrencyLease -LeasePath $again.LeasePath -LeaseId $again.Lease.LeaseId -OwnerServer SYNTHETIC-A) 'Repeated release or existing failed-journal recovery failed.'
    }
    $foreign=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey fixture -JobName Fixture -Occurrence $occurrence.AddMinutes(3) -OwnerServer SYNTHETIC-A
    $journal=(Get-SmartM365JsonNames $foreign.LeasePath).Journal
    [IO.File]::AppendAllText($journal,(@{Owner='Unrelated owner';Phase='Failed';SHA256='';Detail='Synthetic'}|ConvertTo-Json -Compress)+[Environment]::NewLine)
    $foreignHash=(Get-FileHash $journal).Hash
    foreach ($retry in 1..2) { Reject {Get-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey fixture} 'Foreign owner accepted during recovery.' }
    Check ((Get-FileHash $journal).Hash -eq $foreignHash -and (Test-Path $foreign.LeasePath)) 'Foreign journal or lease changed on rejection.'
    Copy-Item $readClaim.ClaimPath $claim.ClaimPath
    [IO.File]::WriteAllText($readClaim.ClaimPath,'{')
    Reject {Get-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName Fixture -Occurrence $occurrence} 'Invalid preferred claim fell back.'
    $contended=Join-Path $root 'Contended'
    $jobs=@()
    try {
        foreach($owner in @('SYNTHETIC-A','SYNTHETIC-B')) {
            $jobs += Start-Job -ScriptBlock {
                param($moduleFile,$claimRoot,$when,$server)
                $ErrorActionPreference='Stop'
                Import-Module $moduleFile -Force
                & (Get-Module SmartM365.JsonTransport) { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()} } }
                Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claimRoot -JobName Concurrent -Occurrence $when -OwnerServer $server -PlanId synthetic
            } -ArgumentList $modulePath,$contended,$occurrence,$owner
        }
        $null=$jobs|Wait-Job -Timeout 20
        if(@($jobs|Where-Object State -ne Completed).Count){throw 'Synthetic claim contention timed out or failed.'}
        $results=@($jobs|Receive-Job -ErrorAction Stop)
        Check (@($results|Where-Object Acquired).Count -eq 1) 'Two contenders acquired the same occurrence.'
        Check (@($results|Where-Object {$_.ClaimPath -notlike '*.json.txt'}).Count -eq 0) 'Contended claim used legacy extension.'
    } finally { $jobs|Stop-Job -ErrorAction SilentlyContinue; $jobs|Remove-Job -Force -ErrorAction SilentlyContinue }
    [pscustomobject]@{Passed=$script:passed;FixtureRoot=$root;Evidence='Synthetic distributed files and local worker processes; no inventory'}
} finally {
    & $distributed { Remove-Item Function:script:Get-SmartM365JsonTransportPolicy -ErrorAction SilentlyContinue }
    & $transport {param($original) Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $original} $policy
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUHQRVe/Rfql+E
# 9ZooucYS8qc2duZyEsOhMC6i3dacxaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEQm+0piChf7C4GUv2oUvlHaZinNguvi/6tJv61Bl5U4MA0GCSqG
# SIb3DQEBAQUABIIBgCdEAbYohh8kqARzChGtYw87U6yk+38Ee+QAiCjox85MulYe
# 6cuRxxVuketrbxNx7n3qLZKogtbM/bH8AAijzYxc3U7Ah5a7ZZYToKCHcBU4nX5L
# 2q1qrSr9Ict/qd7+i3GC/U9GTKHEalQ3aoE/cY4SB4++4rukyi/3oJemuJrkyWXM
# wigL8yxAH1bC5TBPsHf1G9MfFTJn7xd0axGmeraCYOxx5Un1HrdgYvJosEFx/kUC
# S6Xt6fODxZMkQW5m8GwdsUPbfx1Mc+qfbSkMb8MlJF5b2mRYBL+l3fa7wmM5QDDV
# Cj7JxCRb1Vd6QDdNZlzcHFIaKInJLdN+qLjqfbKOjlrGr56R0nOc5PMLm1fb1iUF
# 7fxmArcCqpHJZxs0Hj9yf3D17snF3QgdsQ8403X0NKX0dbVT4vWiw70BD41s/mSY
# w+/oX5RNOBlLbK05x8d69I2x+uFo/q+R+JxK3W7Ct3UhsOnJQbDZY5HlliLlC6X0
# 30LW8M/xcfmHSjFJr6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxNTI1
# NTdaMC8GCSqGSIb3DQEJBDEiBCAALgxDsr0ll9Uro7Els6WrwH+VPO9FsMBWwts2
# c9gXYTANBgkqhkiG9w0BAQEFAASCAgCPrMbWel7PCaGSrCA2bCzdiGbvsjhNMVVB
# z+cAaBPbu+rIVeDZtni21UpSNUlYjNvX5wZO2ekCkvKXwM9EZzJzjPH0paj1oHBj
# w63GZdXsuPUrceTXvfwPBdQ8FPP+GZrlQiLyxUjNtKKjmapkjukwom8DNriH8NoK
# yxzQ2bTH1co0o+7zeNN265+sPKRsV1fGkvRIsz5SSOruToA1zD3aWtfEVOdAcwpZ
# Zw9wf4MoEjDa2VwsqaROnNciSu6liTU5FF6ACOG/OYcf8+4A6A2iZciJhOAlidjs
# cFW8iAcLQqa2opsCBhGs13d4ltFgQj/JWOtKHBywfHmDQ3+khAvt6fvdpm5efkOf
# vPIfEDp8BNUTmbpbCD9AQxtEWjirH1LQdX9sUv5b2obAXh38jNbGeFTRL7TBjtnD
# BZXjGMRUbdNd9uyqg5QdJ8zOfHerxVu8eu2yxwzfc3JBFeuzl+SeeII4OSnT4Q2j
# Yh8DA4mKL+1b7gPkkp0RRvVo2kthBOhMEjgLpFKZnLj90LA6mwVrITM7XPPxAiXk
# sRFCvGUCRfPhYy1w/MnKiTxyQ+q/uKhLVm92OSRPBeD9cPs7ZPXiXNasaRAe8nS/
# csZI3quB1s76CBW+Bxc5mijTAA/ZInWkFusXRFYYrx4qT1nk8qXFzFN+ttH58ALe
# ZWrkv7HWbA==
# SIG # End signature block
