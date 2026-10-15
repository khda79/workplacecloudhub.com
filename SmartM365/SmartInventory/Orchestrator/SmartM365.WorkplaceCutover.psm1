#Requires -Version 7.2
# Machine-readable configuration helper, version 1.0.0.
# No tenant initialization, collection, preparation, transport or Desktop action.
Set-StrictMode -Version Latest

function New-SmartM365WorkplaceCutoverPlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$JobsDocument)
    $copy = $JobsDocument | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $workplace = @($copy.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')
    $legacy = @($copy.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare')
    if ($workplace.Count -ne 1 -or $legacy.Count -gt 1) { throw 'Missing or duplicate preparation job.' }
    $workplace = $workplace[0]
    if ($workplace.Enabled -ne $true -or $workplace.AssignmentMode -eq 'Manual' -or
        $workplace.ScriptPath.Replace('\','/') -cne 'PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1') {
        throw 'Workplace preparation must already be enabled with the expected entry point.'
    }
    if ($workplace.DependencyMode -ne 'FreshSuccess') { throw 'Workplace must use FreshSuccess dependencies.' }
    foreach ($job in $copy.Jobs) {
        if ('CmdbEvidence-Prepare' -in @($job.DependsOn)) {
            throw "A job still depends on the legacy preparation: $($job.Name)."
        }
    }
    $required = @('M365-WorkplaceScope-Inventory','M365-VerifiedDomains-Inventory',
                  'Intune-WindowsAutopilot-Inventory','Intune-AutopatchAlerts-Inventory')
    if ($legacy.Count -eq 1) {
        if ($legacy[0].ScriptPath.Replace('\','/') -cne 'PreparedEvidence/SmartM365-CmdbEvidence-Prepare.ps1') {
            throw 'Unexpected legacy preparation entry point.'
        }
        $required += @($legacy[0].DependsOn)
    }
    $changes = [Collections.Generic.List[object]]::new()
    foreach ($name in @($required | Select-Object -Unique)) {
        $producer = @($copy.Jobs | Where-Object Name -eq $name)
        if ($producer.Count -ne 1 -or $producer[0].Enabled -ne $true -or $producer[0].AssignmentMode -eq 'Manual') {
            throw "Required native dependency must be enabled and schedulable: $name."
        }
        if ($name -notin @($workplace.DependsOn)) {
            $workplace.DependsOn = @($workplace.DependsOn) + $name
            $changes.Add([pscustomobject]@{Job=$workplace.Name; Field='DependsOn'; Before='Absent'; After=$name})
        }
    }
    if ($legacy.Count -eq 1 -and $legacy[0].Enabled -ne $false) {
        $legacy[0].Enabled = $false
        $changes.Add([pscustomobject]@{Job=$legacy[0].Name; Field='Enabled'; Before=$true; After=$false})
    }
    [pscustomobject]@{Jobs=$copy; Changes=@($changes.ToArray()); RetirementPending=$true}
}

function Read-WorkplaceCutoverFile {
    param([Parameter(Mandatory)][string]$Path)
    # Deliberately bypass migration-capable configuration readers in preview.
    if (Test-Path -LiteralPath ($Path+'.txt') -PathType Leaf) { $Path += '.txt' }
    $bytes = [IO.File]::ReadAllBytes($Path)
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    $text = [Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xfeff)
    [pscustomobject]@{Path=$Path; Bytes=$bytes; Hash=$hash; Document=($text | ConvertFrom-Json -Depth 100 -ErrorAction Stop)}
}

function Get-SmartM365WorkplaceCutoverPlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath)
    $jobs = Read-WorkplaceCutoverFile (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Jobs.json')
    $cluster = Read-WorkplaceCutoverFile (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Cluster.json')
    $plan = New-SmartM365WorkplaceCutoverPlan -JobsDocument $jobs.Document
    [pscustomobject]@{Changes=$plan.Changes; Jobs=$plan.Jobs; Cluster=$cluster.Document;
        JobsHash=$jobs.Hash; ClusterHash=$cluster.Hash; JobsFile=$jobs; ClusterFile=$cluster}
}

function Assert-WorkplaceCutoverIdle {
    param([string]$SharedDataFolderPath, $Cluster)
    $servers = @($Cluster.ExpectedOrchestratorServers)
    if (-not $servers.Count) { throw 'No expected orchestrator servers.' }
    foreach ($server in $servers) {
        if ($server -notmatch '^[a-zA-Z0-9][a-zA-Z0-9.-]*$') { throw 'Invalid orchestrator server name.' }
        $hb = (Read-WorkplaceCutoverFile (Join-Path $SharedDataFolderPath "$server/Orchestrator-Heartbeat.json")).Document
        $timestamp = if ($hb.Timestamp -is [datetime]) { [datetimeoffset]$hb.Timestamp } else {
            [datetimeoffset]::Parse([string]$hb.Timestamp,[Globalization.CultureInfo]::InvariantCulture)
        }
        $age = ([datetimeoffset]::UtcNow - $timestamp).TotalMinutes
        if ($age -lt -1 -or $age -gt [double]$Cluster.PeerHeartbeatStaleMinutes -or
            $hb.Lifecycle -ne 'Running' -or $hb.StatePersistenceHealthy -ne $true) {
            throw "Fresh running heartbeat with healthy state persistence required: $server."
        }
        foreach ($job in @($hb.RunningJobs) + @($hb.PendingJobs)) {
            if ($job.Name -in @('WorkplaceEvidence-Prepare','CmdbEvidence-Prepare')) {
                throw "Preparation is running or pending on $server. No configuration published."
            }
        }
    }
}

function Publish-SmartM365WorkplaceCutover {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$ExpectedJobsHash,
        [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$ExpectedClusterHash,
        [Parameter(Mandatory)][string]$BackupDirectory,
        [Parameter(Mandatory)][switch]$NoManualPreparationConfirmed
    )
    $ErrorActionPreference = 'Stop'
    if (-not $NoManualPreparationConfirmed) { throw 'Confirm that no manually launched preparation or transfer is running.' }
    $plan = Get-SmartM365WorkplaceCutoverPlan $SharedDataFolderPath
    if ($plan.JobsHash -ne $ExpectedJobsHash -or $plan.ClusterHash -ne $ExpectedClusterHash) {
        throw 'Configuration changed since preview. Review a new plan.'
    }
    if (-not $plan.Changes.Count) { return [pscustomobject]@{Status='AlreadyApplied'; Changes=@()} }
    Assert-WorkplaceCutoverIdle $SharedDataFolderPath $plan.Cluster
    if (-not [IO.Path]::IsPathFullyQualified($BackupDirectory)) { throw 'BackupDirectory must be absolute and outside the repository and orchestrator root.' }
    $backupRoot = [IO.Path]::GetFullPath($BackupDirectory).TrimEnd('\','/')
    $repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../../..')).TrimEnd('\','/')
    $sharedRoot = [IO.Path]::GetFullPath($SharedDataFolderPath).TrimEnd('\','/')
    foreach ($excluded in @($repoRoot,$sharedRoot)) {
        if ($backupRoot -eq $excluded -or $backupRoot.StartsWith($excluded+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {
            throw 'BackupDirectory must be outside the repository and orchestrator root.'
        }
    }
    $ancestor = $backupRoot
    while ($ancestor) {
        if (Test-Path -LiteralPath $ancestor) {
            $item = Get-Item -LiteralPath $ancestor -Force
            if (-not $item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                throw 'BackupDirectory must not traverse a file, junction or symbolic link.'
            }
        }
        $parent = Split-Path -Path $ancestor -Parent
        if ($parent -eq $ancestor) { break }
        $ancestor = $parent
    }
    if (-not $PSCmdlet.ShouldProcess($SharedDataFolderPath,'Disable legacy CMDB job and merge native dependencies; no job request')) { return }
    Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Management.psm1') -ErrorAction Stop
    $validation = Test-SmartM365OrchestratorJobsDocument $plan.Jobs
    $consistency = Test-SmartM365OrchestratorConfigurationConsistency $plan.Jobs $plan.Cluster
    if (-not $validation.Valid -or -not $consistency.Valid) {
        throw "Invalid cutover: $(@($validation.Errors)+@($consistency.Errors) -join '; ')"
    }
    $backup = Join-Path $backupRoot ('WorkplaceCutover-'+[guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $backup -ErrorAction Stop
    foreach ($file in @($plan.JobsFile,$plan.ClusterFile)) {
        $target = Join-Path $backup ([IO.Path]::GetFileName($file.Path))
        [IO.File]::WriteAllBytes($target,$file.Bytes)
        if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $file.Hash) { throw 'Backup verification failed. No configuration published.' }
    }
    # Heartbeats are not a distributed lock; operator confirmation also covers direct launchers.
    Assert-WorkplaceCutoverIdle $SharedDataFolderPath $plan.Cluster
    $result = Publish-SmartM365OrchestratorConfiguration -SharedDataFolderPath $SharedDataFolderPath `
        -JobsDocument $plan.Jobs -ClusterDocument $plan.Cluster `
        -ExpectedJobsHash $ExpectedJobsHash -ExpectedClusterHash $ExpectedClusterHash `
        -ChangeSummary 'Workplace cutover: disable legacy preparation and merge native dependencies; preserve schedules and collectors'
    $after = Get-SmartM365WorkplaceCutoverPlan $SharedDataFolderPath
    if ($after.Changes.Count -ne 0 -or $after.JobsHash -ne $result.JobsHash -or $after.ClusterHash -ne $result.ClusterHash) {
        throw "Post-publication verification failed; inspect configuration and backup $backup. Do not rerun preparation."
    }
    [pscustomobject]@{Status='Published'; VersionId=$result.VersionId; BackupPath=$backup;
        JobsHash=$result.JobsHash; ClusterHash=$result.ClusterHash; Changes=$plan.Changes}
}

Export-ModuleMember -Function New-SmartM365WorkplaceCutoverPlan, Get-SmartM365WorkplaceCutoverPlan, Publish-SmartM365WorkplaceCutover

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA0V5efNVpgfcna
# Zphx7FBAymZwrFp00feqEhxmNYLDBKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIEia2b6KdfKK9OnBDtetYW/tlCDj/gJm9GchPT6yiEqMA0GCSqG
# SIb3DQEBAQUABIIBgGE4li0oBHRzhLPdM8VlSatMYglhLkCJZJ37mo4Kb90Cv5FU
# jATGAaJg4i3jdaQtITPU7TtQ/ochi69YqDLbJSoIQCnY7YHgbhovk2+DlnGwl6x7
# d5jGNfh/TIHILfKmpkQEc3aWRiK71m1ojSQWcEfLg2nteaBWUTDD1fq3oYDN2Sr9
# dCgYDzqWRNWqUI2FTNv3oeHB/saKoRwkhgB6Pjd/2cdi0tPn348LACruz2Si0yIx
# gSwapP9h3peB9Awv8zHrZcvfvIL+UXDdP/Ypq4unUNvolSMJkJXGXQXUsM1Ctwmf
# SZlG3QqpxWb75i8fIBf7yJs0qzO5DUcf98FBXxghv0DVyIerauejarcBhi7nkce5
# VAQ0ZTEa5j4MHWcvrJP/YSwxF9ZapAubb8D5IolJ6HwjRuJiEnLrpzgr69fXFKpc
# ZTliZ7ZNk4JAHICMNhciBfgSKMlQEEaVdIYiDQMi0FqSeWiYccHyeu1JHmnl9oAi
# QaxXgSmqF4cct07qcKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxMTMx
# MThaMC8GCSqGSIb3DQEJBDEiBCAOmNHbUjNrEZxOW5FA4IWtUzzzh1Q4wJFYnr0t
# wnVW8TANBgkqhkiG9w0BAQEFAASCAgCAHlW4QyGqQZd1ZDqETOYfyqnJfh+ab0LK
# mnPegRwKff4HR8kfRnppXcdtDSGfOOQEFXyRIeRIk6mgMXkiB92LncSLIKPCNUi8
# rBk6qOZzIkiNBhWdyfuBmSywiccBXeI/u/RcuNwFq6bzv8BVFEfwg01vd6rZQY+9
# 1Eg7z7SgSORYUW6xIC7LTOcpFuZZLFbg5GDnl8XKDYNxxRPHVdNilBp9QccqeEx/
# ZVDMOwbUuey8ZOTX/yMVyCLYv5LYA/8McVAPZXFTCcd9cir3grE2gjjr9xGU//RB
# KGJDLxlT6sfmMt+ZGT3e87hOHKCcaf6VXfA9kWlqGV9IkpeGdupcf9gquG97cHQD
# I6SzV6DIBU+z36arGzdy6N/cTaUSvk3Z9PeRCLaBQnH/CDH/VTdgMogsc+NRR/n/
# DUAYelBAwg+ut3k7OCFU7N04HNT8/1QKCfEPWJMLnsYWCfDXLeZ1i62FgD64Z/2Z
# +FB1wI1ckUGWSB+lqf2EVxUqUUwfLVPurQ4o8jBTMzqlKPnWBKym29vOi4QhnM9i
# vvT9MsPy7aUkUiHaH22BYgm4jAr3zKwyEDHelR58ghY+21V2mfjkgXqOsI4Uys3B
# 22BXw68jPTEA5fRvUbsUbV2OS2FRhbkqfmelaf3ggPVlLwhxAiPk/5fLq6shtSkO
# lWmqGwrHNA==
# SIG # End signature block
