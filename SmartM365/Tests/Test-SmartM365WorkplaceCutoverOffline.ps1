#Requires -Version 7.2
# Synthetic configuration tests only. No operational preparation or collection.
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$orchestrator=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $orchestrator 'SmartM365.WorkplaceCutover.psm1') -Force
$checks=0
function Check([bool]$Value,[string]$Message){if(-not $Value){throw $Message};$script:checks++}
function Reject([scriptblock]$Action,[string]$Pattern){
    $message='';try{& $Action | Out-Null}catch{$message=$_.Exception.Message}
    Check ($message -match $Pattern) "Expected '$Pattern'; got '$message'"
}
function Json($Value){ConvertTo-Json -InputObject $Value -Depth 100 -Compress}
function Clone($Value){Json $Value | ConvertFrom-Json -Depth 100}
$template=Get-Content (Join-Path $orchestrator 'Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
Check (-not ($template.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare').Enabled) 'Legacy template job enabled'
$old=Clone $template
$w=$old.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare'
$m=$old.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare'
$m.Enabled=$true
$native=@('M365-WorkplaceScope-Inventory','M365-VerifiedDomains-Inventory','Intune-WindowsAutopilot-Inventory','Intune-AutopatchAlerts-Inventory')
$w.DependsOn=@($w.DependsOn | Where-Object {$_ -notin $native})
$w.Schedule.Times=@('08:30')
$w.Arguments='-Tenant test'
$w | Add-Member NoteProperty PrivateExtra 'preserved'
$before=Json $old
$plan=New-SmartM365WorkplaceCutoverPlan $old
Check ((Json $old) -ceq $before) 'Input mutated'
Check ($plan.Changes.Count -eq 5) 'Expected four dependencies and one disable'
$afterW=$plan.Jobs.Jobs | Where-Object Name -eq $w.Name
foreach($property in $w.PSObject.Properties){
    if($property.Name -eq 'DependsOn'){continue}
    Check ((Json $afterW.($property.Name)) -ceq (Json $property.Value)) "Workplace field changed: $($property.Name)"
}
foreach($job in $old.Jobs){
    if($job.Name -in @($w.Name,$m.Name)){continue}
    Check ((Json ($plan.Jobs.Jobs | Where-Object Name -eq $job.Name)) -ceq (Json $job)) "Unrelated job changed: $($job.Name)"
}
$expectedLegacy=Clone $m;$expectedLegacy.Enabled=$false
Check ((Json ($plan.Jobs.Jobs | Where-Object Name -eq $m.Name)) -ceq (Json $expectedLegacy)) 'Legacy fields changed beyond Enabled'
Check ((New-SmartM365WorkplaceCutoverPlan $plan.Jobs).Changes.Count -eq 0) 'Not idempotent'
$bad=Clone $old;($bad.Jobs | Where-Object Name -eq $native[0]).Enabled=$false
Reject {New-SmartM365WorkplaceCutoverPlan $bad} 'enabled and schedulable'
$bad=Clone $old;$bad.Jobs+=Clone $m
Reject {New-SmartM365WorkplaceCutoverPlan $bad} 'duplicate'
$bad=Clone $old;($bad.Jobs | Where-Object Name -eq $w.Name).DependsOn+= $m.Name
Reject {New-SmartM365WorkplaceCutoverPlan $bad} 'still depends'
$bad=Clone $old;($bad.Jobs | Where-Object Name -eq $w.Name).Enabled=$false
Reject {New-SmartM365WorkplaceCutoverPlan $bad} 'already be enabled'
$bad=Clone $old;($bad.Jobs | Where-Object Name -eq $w.Name).DependencyMode='LatestOccurrence'
Reject {New-SmartM365WorkplaceCutoverPlan $bad} 'FreshSuccess'
$bad=Clone $old;$bad.Jobs=@($bad.Jobs | Where-Object Name -ne $m.Name)
Check ((New-SmartM365WorkplaceCutoverPlan $bad).Changes.Count -eq 4) 'Absent legacy job unsupported'

$root=Join-Path ([IO.Path]::GetTempPath()) ('WorkplaceCutoverTest-'+[guid]::NewGuid().ToString('N'))
try{
    $shared=Join-Path $root 'Orchestrator'
    $null=New-Item -ItemType Directory -Path (Join-Path $shared 'Config'),(Join-Path $shared 'Audit'),(Join-Path $shared 'WORKER-TEST') -Force
    $cluster=Get-Content (Join-Path $orchestrator 'Orchestrator-Cluster.json.template') -Raw | ConvertFrom-Json
    $cluster.ExpectedOrchestratorServers=@('WORKER-TEST')
    $cluster.ElectionWeightsByServer=[pscustomobject]@{'WORKER-TEST'=1.0}
    $cluster.ServerJobPolicies=[pscustomobject]@{}
    $jobsPath=Join-Path $shared 'Config/Orchestrator-Jobs.json.txt'
    $clusterPath=Join-Path $shared 'Config/Orchestrator-Cluster.json.txt'
    $heartbeatPath=Join-Path $shared 'WORKER-TEST/Orchestrator-Heartbeat.json.txt'
    [IO.File]::WriteAllText($jobsPath,(Json $old))
    [IO.File]::WriteAllText($clusterPath,(Json $cluster))
    # A legacy version must NOT be migrated by preview.
    $version=Join-Path $shared 'Config/Versions/20261001T000000000Z_deadbeef'
    $null=New-Item -ItemType Directory -Path $version -Force
    [IO.File]::WriteAllText((Join-Path $version 'Orchestrator-Jobs.before.json'),(Json $old))
    $filesBefore=@(Get-ChildItem $shared -File -Recurse | Get-FileHash | Select-Object Path,Hash)
    $preview=Get-SmartM365WorkplaceCutoverPlan $shared
    Check ((Json @(Get-ChildItem $shared -File -Recurse | Get-FileHash | Select-Object Path,Hash)) -ceq (Json $filesBefore)) 'Preview wrote files'
    Check ($preview.Changes.Count -eq 5) 'Disk preview differs'
    $publish=@{SharedDataFolderPath=$shared;ExpectedJobsHash=$preview.JobsHash;ExpectedClusterHash=$preview.ClusterHash;
        BackupDirectory=(Join-Path $root 'Backup');NoManualPreparationConfirmed=$true;Confirm=$false}
    $stale=@{}+$publish;$stale.ExpectedJobsHash='0'*64
    Reject {Publish-SmartM365WorkplaceCutover @stale} 'changed since preview'
    Reject {Publish-SmartM365WorkplaceCutover @publish} 'Could not find|not find|introuvable'
    $hb=[pscustomobject]@{Timestamp=[datetimeoffset]::UtcNow.ToString('o');Lifecycle='Running';StatePersistenceHealthy=$true;RunningJobs=@();PendingJobs=@()}
    [IO.File]::WriteAllText($heartbeatPath,(Json $hb))
    $fileCount=@(Get-ChildItem $root -Recurse -File).Count
    Publish-SmartM365WorkplaceCutover @publish -WhatIf
    Check (@(Get-ChildItem $root -Recurse -File).Count -eq $fileCount) 'WhatIf wrote files'
    Check (-not (Test-Path $publish.BackupDirectory)) 'WhatIf created backup'
    $hb.RunningJobs=@([pscustomobject]@{Name='CmdbEvidence-Prepare'})
    [IO.File]::WriteAllText($heartbeatPath,(Json $hb))
    Reject {Publish-SmartM365WorkplaceCutover @publish} 'running or pending'
    $hb.RunningJobs=@();$hb.PendingJobs=@([pscustomobject]@{Name='WorkplaceEvidence-Prepare'})
    [IO.File]::WriteAllText($heartbeatPath,(Json $hb))
    Reject {Publish-SmartM365WorkplaceCutover @publish} 'running or pending'
    $hb.PendingJobs=@();$hb.Timestamp=[datetimeoffset]::UtcNow.AddDays(-1).ToString('o')
    [IO.File]::WriteAllText($heartbeatPath,(Json $hb))
    Reject {Publish-SmartM365WorkplaceCutover @publish} 'Fresh running heartbeat'
    $hb.Timestamp=[datetimeoffset]::UtcNow.ToString('o')
    [IO.File]::WriteAllText($heartbeatPath,(Json $hb))
    $invalid=@{}+$publish;$invalid.BackupDirectory=Join-Path $shared 'backup'
    Reject {Publish-SmartM365WorkplaceCutover @invalid} 'outside the repository'
    $invalid=@{}+$publish;$invalid.NoManualPreparationConfirmed=$false
    Reject {Publish-SmartM365WorkplaceCutover @invalid} 'Confirm that no manually'
    # Real configuration publication API against this synthetic directory only.
    $result=Publish-SmartM365WorkplaceCutover @publish
    Check ($result.Status -eq 'Published') 'Synthetic publication failed'
    Check ((Get-FileHash (Join-Path $result.BackupPath 'Orchestrator-Jobs.json.txt')).Hash -eq $preview.JobsHash) 'Backup jobs mismatch'
    Check ((Get-FileHash (Join-Path $result.BackupPath 'Orchestrator-Cluster.json.txt')).Hash -eq $preview.ClusterHash) 'Backup cluster mismatch'
    $readBack=Get-SmartM365WorkplaceCutoverPlan $shared
    Check ($readBack.Changes.Count -eq 0) 'Read-back is not converged'
    Check ((Json $readBack.Cluster) -ceq (Json $cluster)) 'Cluster values changed'
    $publish.ExpectedJobsHash=$readBack.JobsHash;$publish.ExpectedClusterHash=$readBack.ClusterHash
    Check ((Publish-SmartM365WorkplaceCutover @publish).Status -eq 'AlreadyApplied') 'Repeated publication not a no-op'
    Check (-not (Test-Path (Join-Path $shared 'Requests'))) 'Job request created'
    Write-Output "Workplace cutover offline checks passed: $checks. Synthetic configuration only."
}finally{
    $resolved=[IO.Path]::GetFullPath($root)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/')
    if($resolved.StartsWith($temp+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -and
       [IO.Path]::GetFileName($resolved) -match '^WorkplaceCutoverTest-[a-f0-9]{32}$' -and (Test-Path -LiteralPath $resolved)){
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBEQjbCxcmUTLja
# JTer92SdsCLJ2DGfSmirNfV7ws1ZJ6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM6uLRUzdFJHBr8DT3u59FBq69jaUW0bW2lF0PC0B27SMA0GCSqG
# SIb3DQEBAQUABIIBgAclEmMGhWR3Dv4beAky/hE6W44tPJubaEK3VcJ34ytl72iI
# eS/GsNtEstLq5kkbqoOa3jBTCeJB6I/xswLFEaIACiWBhep9Pr71YkrkiVelPrU4
# 0yTD5XSyf+uTnfhGUdXjLpuKsLXOO+EOqQS5xy9cnVlRb7X8wuW+6WQDjB6See+N
# pXBRDqsSNIP+YEYZN5Gn0jh+b95uq4nsFYRV4sQDxSESu9Emy2DsVYaSovgnVATa
# KdaFH4g7yuIoJ3IkJPmBdJkTgJxNAq6OEHeud4TiKosRlIRTvs1nQEJ1FTVFiaqZ
# yd0tXHjyUIv3Nze0O8LqES48BSP6nNoj7J8Jw5CZ+mSO9rZBptGzjeG5fayRBFbR
# lEzkbDUd2QTSi0Tsu8QW3VzK/im02faIRmu7VB7VMcS+RkfDQAUO7iR0zoLJjnoF
# HRkPwdjqzWA7n0QXtTxHuXzNGjGZJ+HyLfbAk5KrdRl7djxEz0lDjALZcuNTq+qG
# pyUdbG3OM0mT7sUke6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxMTMx
# MThaMC8GCSqGSIb3DQEJBDEiBCDzKNA8iLqdLX7K7bapvr0d/T4XEyfzTZAOTtEU
# VufIDDANBgkqhkiG9w0BAQEFAASCAgAQS8r2Z8JK+PSXSzcD0WJ+UxRfkrnPFy/E
# /15/3NWr745kDnPSjvIY2SJJxc2athYPFJBLqytAFwC9l2PMAc7SI8Gr8YsU3O2C
# 6xsCAU4NFdgRf4hfSRYBz4TGgmgwWrSK9q4H31rdPnJr1I3y6u30mMfX+Tt2ZXon
# ZBs5dqbuXdvu5JXr3fESWR97eHaoFgaKrAgNuRuAnEEcT+f6DW3pnrl16c2CsarS
# 80GmA2WiXF6iSrUs0p94BrLuqU3TPxJ4eJZaDpOGkKfANYuIwomPhxykuv9BBMNQ
# if8wcqNcZAWcjwoj7w5EP7ZKarJ/sLKuCn1e2Ycjk91shGXtumB8ZoOo9YXQfPM8
# IQlgxERQnBVgJtbgvxbS7/rFKHBgX4jh37WU86SKD+JpxijnDf3Lk4c4OjWTOjmG
# hPyWloohgAWAXete+D/tg3DSJRq1Wxk2uPtWahxB7v2NDKRZnG18Zg8Kuc1OZMEa
# Zb5Ht/3e2bPaAHLbHCHbxK3xQBpCr/3sJjldp2McRM2EJU8vMeViUdtN1kguFHmS
# 1YRhM3HAVZ0Z4mRhxNUNfWIUQIG83i+I/ee2EFKb2x2qYW4rhd3CzbDkFwmL7llN
# TUyOUIexfqdqU1dCkyLvw3orcNdupLptlRJAyoET8pUtjPq6ifLWwyp49MJo6FDt
# suesWuFUaQ==
# SIG # End signature block
