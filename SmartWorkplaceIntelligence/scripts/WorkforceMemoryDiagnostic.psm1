# Manual opt-in diagnostic only. No collector, publisher, cleanup or model access.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedMetadata.psm1') -Force

function Get-WorkforceSystemMemory {
    try {
        $memory=Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Memory -OperationTimeoutSec 2 -ErrorAction Stop
        [pscustomobject]@{AvailableBytes=[long]$memory.AvailableBytes;CommittedBytes=[long]$memory.CommittedBytes;CommitLimitBytes=[long]$memory.CommitLimit;Error=$null}
    } catch { [pscustomobject]@{AvailableBytes=$null;CommittedBytes=$null;CommitLimitBytes=$null;Error='System memory counters unavailable'} }
}

function Invoke-MonitoredWorkforceProcess {
    param([string]$WorkerPath,[string[]]$WorkerArguments,[string]$RunRoot)
    $stagePath=Join-Path $RunRoot 'stages.ndjson'
    $samplesPath=Join-Path $RunRoot 'memory.csv'
    $stdout=Join-Path $RunRoot 'worker.stdout.log';$stderr=Join-Path $RunRoot 'worker.stderr.log'
    $shell=(Get-Process -Id $PID).Path
    # Start-Process joins ArgumentList on Windows; quote every path explicitly.
    $arguments=@('-NoProfile','-File',$WorkerPath)+$WorkerArguments
    foreach($argument in $arguments){if($argument -match '["\r\n]'){throw 'Invalid diagnostic argument.'}}
    $quoted=@($arguments | ForEach-Object {'"'+$_+'"'}) -join ' '
    $started=[datetime]::UtcNow
    $os=Get-CimInstance Win32_OperatingSystem -OperationTimeoutSec 2 -ErrorAction SilentlyContinue
    $hostInfo=@{StartedUtc=$started.ToString('O');ParentProcessId=$PID;PowerShellVersion=$PSVersionTable.PSVersion.ToString();Is64BitProcess=[Environment]::Is64BitProcess;TotalPhysicalBytes=if($os){[long]$os.TotalVisibleMemorySize*1024}else{$null};SampleIntervalSeconds=5;WorkerSHA256=(Get-FileHash -LiteralPath $WorkerPath).Hash;MonitorSHA256=(Get-FileHash -LiteralPath $PSCommandPath).Hash;Publication=$false;InputMode='Existing raw exports, read only; no atomic snapshot claimed'}
    $environmentPath=Resolve-SmartM365OwnedJsonPath -Path (Join-Path $RunRoot 'environment.json') -Owner 'WorkplaceEvidence-Prepare/diagnostic' -Validate {param($document) if($document.Publication -ne $false){throw 'Invalid diagnostic audit.'}}
    Write-SmartM365JsonBytesAtomically -Path $environmentPath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($hostInfo | ConvertTo-Json))) -Validate {param($document) if($document.Publication -ne $false){throw 'Invalid diagnostic audit.'}} | Out-Null
    $worker=Start-Process -FilePath $shell -ArgumentList $quoted -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru
    $lastStage='Worker startup';$announcedStage='';$samples=0;$sampleErrors=0;$peakPrivate=0L;$minimumAvailable=$null;$maximumCommitted=$null;$peakWorkingSet=0L
    try {
        do {
            try {
                if(Test-Path -LiteralPath $stagePath){
                    try {$last=Get-Content -LiteralPath $stagePath -Tail 1 | ConvertFrom-Json; if($last.Stage){$lastStage=[string]$last.Stage}} catch { }
                }
                $system=Get-WorkforceSystemMemory
                if($system.Error){$sampleErrors++}
                $worker.Refresh()
                $private=$null;$working=$null
                if(-not $worker.HasExited){$private=$worker.PrivateMemorySize64;$working=$worker.WorkingSet64;$peakWorkingSet=[math]::Max($peakWorkingSet,$worker.PeakWorkingSet64);$peakPrivate=[math]::Max($peakPrivate,$private)}
                if($null -ne $system.AvailableBytes -and ($null -eq $minimumAvailable -or $system.AvailableBytes -lt $minimumAvailable)){$minimumAvailable=$system.AvailableBytes}
                if($null -ne $system.CommittedBytes -and ($null -eq $maximumCommitted -or $system.CommittedBytes -gt $maximumCommitted)){$maximumCommitted=$system.CommittedBytes}
                [pscustomobject]@{Utc=[datetime]::UtcNow.ToString('O');WorkerProcessId=$worker.Id;Stage=$lastStage;Exited=$worker.HasExited;PrivateBytes=$private;WorkingSetBytes=$working;AvailablePhysicalBytes=$system.AvailableBytes;SystemCommittedBytes=$system.CommittedBytes;SystemCommitLimitBytes=$system.CommitLimitBytes;CounterError=$system.Error} |
                    Export-Csv -LiteralPath $samplesPath -NoTypeInformation -Append -Encoding utf8
                $samples++
                if($lastStage -ne $announcedStage){Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Workforce diagnostic stage: {1}' -f (Get-Date),$lastStage);$announcedStage=$lastStage}
            } catch {$sampleErrors++;Write-Warning "Diagnostic sample unavailable: $($_.Exception.GetType().Name)"}
            if($worker.HasExited){break}
            $null=$worker.WaitForExit(5000)
        } while($true)
        $worker.WaitForExit()
        $result=[pscustomobject]@{RunRoot=$RunRoot;WorkerProcessId=$worker.Id;ExitCode=$worker.ExitCode;LastStage=$lastStage;DurationSeconds=[math]::Round(([datetime]::UtcNow-$started).TotalSeconds,1);Samples=$samples;SampleErrors=$sampleErrors;SampledPeakPrivateBytes=$peakPrivate;ObservedPeakWorkingSetBytes=$peakWorkingSet;MinimumAvailablePhysicalBytes=$minimumAvailable;MaximumSystemCommittedBytes=$maximumCommitted;Publication=$false}
        $resultPath=Resolve-SmartM365OwnedJsonPath -Path (Join-Path $RunRoot 'result.json') -Owner 'WorkplaceEvidence-Prepare/diagnostic' -Validate {param($document) if($document.RunRoot -ne $RunRoot){throw 'Diagnostic run mismatch.'}}
        Write-SmartM365JsonBytesAtomically -Path $resultPath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($result | ConvertTo-Json))) -Validate {param($document) if($document.RunRoot -ne $RunRoot){throw 'Diagnostic run mismatch.'}} | Out-Null
        $result
    } finally {
        # Never automatically cancel or kill a workload being measured.
        if(-not $worker.HasExited){Write-Warning "Diagnostic worker $($worker.Id) still running; waiting without cancelling it.";$worker.WaitForExit()}
        $worker.Dispose()
    }
}

function Invoke-WorkforceMemoryDiagnostic {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataRoot,[Parameter(Mandatory)][string]$WorkRoot,
        [Parameter(Mandatory)][string]$AccountClassificationConfigPath,[string]$MappingRoot)
    if($PSVersionTable.PSVersion.Major -lt 7 -or -not [Environment]::Is64BitProcess){throw 'PowerShell 7 x64 is required.'}
    $raw=(Resolve-Path -LiteralPath $DataRoot).ProviderPath.TrimEnd('\')
    $work=[IO.Path]::GetFullPath($WorkRoot).TrimEnd('\')
    if($work -eq $raw -or $work.StartsWith($raw+'\',[StringComparison]::OrdinalIgnoreCase) -or $raw.StartsWith($work+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Diagnostic work and raw data roots must be disjoint.'}
    if(-not $MappingRoot){$MappingRoot=$raw}
    foreach($path in $AccountClassificationConfigPath,(Join-Path $MappingRoot 'SmartWorkplaceIntelligence-PersonaClassification.xlsx'),(Join-Path $MappingRoot 'SmartWorkplaceIntelligence-SiteClassification.xlsx')){if(-not(Test-Path -LiteralPath $path -PathType Leaf)){throw "Diagnostic input missing: $path"}}
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $workLock=[IO.File]::Open((Join-Path $work '.preparation.lock'),'OpenOrCreate','ReadWrite','None')
    $sourceLock=$null
    try {
        $sourceLock=[IO.File]::Open((Join-Path $raw '.prepared-source.lock'),'OpenOrCreate','ReadWrite','None')
        Convert-PreparedAuditNames -Root $work -Family workforce-diagnostics
        $run=Join-Path $work ('workforce-diagnostics/'+[datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8))
        $outputs=Join-Path $run 'outputs';New-Item -ItemType Directory -Path $outputs -Force | Out-Null
        Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Workforce-only diagnostic. Private logs/output: {1}. No DATA-POWERBI publication; no collectors.' -f (Get-Date),$run)
        $worker=Join-Path $PSScriptRoot 'New-WorkforceIdentityEvidence.ps1'
        $arguments=@('-DataRoot',$raw,'-DiagnosticStagePath',(Join-Path $run 'stages.ndjson'),'-AccountClassificationConfigPath',$AccountClassificationConfigPath,
            '-PersonaClassificationPath',(Join-Path $MappingRoot 'SmartWorkplaceIntelligence-PersonaClassification.xlsx'),'-SiteClassificationPath',(Join-Path $MappingRoot 'SmartWorkplaceIntelligence-SiteClassification.xlsx'),
            '-UserOutputPath',(Join-Path $outputs 'UserInventoryEvidence.csv'),'-HistoryOutputPath',(Join-Path $outputs 'UserActivityHistoryEvidence.csv'),
            '-TrendOutputPath',(Join-Path $outputs 'WorkforceTrendEvidence.csv'),'-IdentityOutputPath',(Join-Path $outputs 'IdentityReconciliationEvidence.csv'),'-SignalsOutputPath',(Join-Path $outputs 'WorkforceOperationalSignals.csv'))
        Invoke-MonitoredWorkforceProcess -WorkerPath $worker -WorkerArguments $arguments -RunRoot $run
    } finally {if($sourceLock){$sourceLock.Dispose()};$workLock.Dispose()}
}
Export-ModuleMember -Function Invoke-WorkforceMemoryDiagnostic

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDDM0+RXL13MXrQ
# L/8CeOIlzVhizGgWMG+j63GezyN2IaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMlpPdCgrk58BloMYM7RZRqmovP29IqSibKXfBZIXmsuMA0GCSqG
# SIb3DQEBAQUABIIBgE3amVYtGbw3mZInG6HGIePGqujj1eJOU2YqgoGfOVNAP0gm
# ivdh6dyP4SXSGpb+ElRRn0WPicNmj2ijAdlHGZJ4XBNj8kZZ6U9u+ML4UOCpja1C
# Blv7vjEZ8ynVSkxq4Xltj8J8kqMYwFYAfJSe31myf82kCEhewmYq7pbFlS0mxnjU
# kYe69xrC4v2HI54y0YVXt6127T+uVLZKS5EyEFpJ7xIJEdvl9zYfDoqicsMwjyo5
# WN01u/Eunux8RX1X5/yvxoa2DIt9I7Zl1AoEO5NJ6DKj//+7r1DRVFHayk4dAVHo
# odMfPYWVXHgqObWnxvzVffPGRiCmhJhtOGQZXDrk3j1qJUMaGptKF41IHBIbtWqK
# Rj7QHAq5AlTrzIqxcDUEw2ULWP8SNfxcGJILslWoi5OQV+5/ovr8aaJzbwJ7HDLd
# BIExnQ7j9zf+93BZoJ+p0GNhTsSMk6h6wNsRDKiEZKq7VMT18B5KJlx2TAE0utWB
# HCARp//jOyYMOY80ZKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU4
# MDBaMC8GCSqGSIb3DQEJBDEiBCCAQ+b/fFtKv/Cf3Y/xS7y5piWcAvo4kAqrKYJK
# gNrmxTANBgkqhkiG9w0BAQEFAASCAgAJ6LRJw+2hDEo9502ns5aIC6KgtgtmPLPK
# YtSDsWKl16Eo7PgsaaHiY4WKJ5PFDUr6LC2tiCUTrg1fZjOHXtaotkTAXrV3SqrQ
# 6Bz1+2MuI3l+GNDyyO8JEibtXys39ac4AhdBuJvWiYCHKBOvNwaQ4smvv+EmWeVp
# fMBdMEVYAkkF3GiQb91NNygz+b69Rjo8+HhHU1U449fROsCDvnDLBVCqdMo7Qe85
# Uvn+S0Dj1hpfS2TyY2BGw1VUU7yj4D/dTVLUrFgnAlcvfOf+NBYmJuLd2eEwlQyA
# Aipz90Cd/rwynD0ZX8KfwshCr7nERCchwFjkfRgX792FUqNvPXC93kQ87mKZmDFR
# GeCD/rXsrjhk86ZB4nMm5Ve3/8x1xmUHbsFz5thUOTqL7LlzJqDK43hOyoNLexFz
# 9bR0zc2snpguEXzlOtTAKEENNZFwIT8y5aAuqjBOpQXTMvsuxqVhaay4pSiFWRJA
# LrupWj8XHqJA5nYVU1yDZgm1AjS84gVRhu0vCH/EIPUgT2veI0irQQI6vVcgj0z3
# rDoOVfLvgl+5Chsmas1S5ySYbfjAU/QmTv7MphaM+2PzMOIFbWjWe8SC4u1i2aS1
# b+u4Up246TjuLWIrZ92lFMn+AfUYkZ+jjGUJdV7pKIdvZUtViUiCr4iuye3V3VrU
# mf5+8yYQUQ==
# SIG # End signature block
