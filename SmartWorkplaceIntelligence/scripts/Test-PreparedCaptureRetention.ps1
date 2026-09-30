[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Synthetic local fixtures. Never invoke tenant collectors or touch production roots.
$ErrorActionPreference='Stop'
$fixture=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$raw=Join-Path $fixture 'raw'; $work=Join-Path $fixture 'work'; $output=Join-Path $raw 'DATA-POWERBI'
New-Item -ItemType Directory -Path (Join-Path $raw 'DATA-LAST'),(Join-Path $fixture 'config') -Force | Out-Null
$csv=Join-Path $raw 'DATA-LAST/Example.csv'
$valid="`"TenantKey`",`"Value`"`r`n`"synthetic-test`",`"one`"`r`n"
[IO.File]::WriteAllText($csv,$valid)
@{currentFiles=@('Example.csv');mappingFiles=@();dailyFiles=@();history=@()} | ConvertTo-Json | Set-Content (Join-Path $fixture 'config/prepared-source-contract.json')
$module=Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
& $module {param($p) $script:ProductRoot=$p} $fixture
$script:checks=0
function Check([bool]$Pass,[string]$Message){if(-not $Pass){throw $Message};$script:checks++}
function Reject([scriptblock]$Action,[string]$Pattern){$message='';try{& $Action | Out-Null}catch{$message=$_.Exception.Message};Check ($message -match $Pattern) "Expected '$Pattern', got '$message'"}
$argsForRun=@{DataRoot=$raw;WorkRoot=$work;OutputRoot=$output;TenantKey='synthetic-test';AccountClassificationConfigPath=(Join-Path $fixture 'unused.psd1');ValidateOnly=$true}
try {
    $entry=@(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath (Join-Path $fixture 'config/prepared-source-contract.json'))[0]
    # A source newer than its initial plan is accepted; actual captured hash is recorded.
    [IO.File]::WriteAllText($csv,$valid.Replace('one','newer'))
    $snapshot=Join-Path $fixture 'direct-copy'
    $captured=& $module {param($e,$s) Copy-PreparedStableSource -Entry $e -SnapshotRoot $s -RetryDelayMs 0} $entry $snapshot
    Check ($captured.SHA256 -ne $entry.SHA256 -and $captured.SHA256 -eq (Get-FileHash $csv).Hash) 'Captured provenance used obsolete planned bytes'
    Check ($captured.CaptureAttempts -eq 1) 'Stable file unexpectedly retried'
    # Deterministic mutation between the initial metadata read and stream opening.
    & $module {
        param($p)
        $script:mutationPath=$p; $script:mutationReads=0
        function script:Get-Item {
            param([string]$LiteralPath,[switch]$Force)
            $item=Microsoft.PowerShell.Management\Get-Item -LiteralPath $LiteralPath -Force:$Force
            if($LiteralPath -eq $script:mutationPath){
                $script:mutationReads++
                if($script:mutationReads -eq 1){$null=$item.Length;[IO.File]::AppendAllText($LiteralPath,"`r`n")}
            }
            $item
        }
    } $csv
    try {$captured=& $module {param($e,$s) Copy-PreparedStableSource -Entry $e -SnapshotRoot $s -RetryDelayMs 0} $entry $snapshot}
    finally {& $module {Remove-Item Function:script:Get-Item}}
    Check ($captured.CaptureAttempts -eq 2) 'Changing file did not retry independently'
    Check ($captured.SHA256 -eq (Get-FileHash $csv).Hash) 'Retry did not capture latest stable bytes'
    # A non-readable source gets exactly three attempts, with its relative path.
    $locked=[IO.File]::Open($csv,'Open','ReadWrite','None')
    try {Reject {& $module {param($e,$s) Copy-PreparedStableSource -Entry $e -SnapshotRoot $s -RetryDelayMs 0} $entry $snapshot} 'after 3 attempts: DATA-LAST[/\\]Example.csv'}
    finally {$locked.Dispose()}
    [IO.File]::WriteAllText($csv,$valid)
    # Raw updates after capture must not affect validation of the frozen copy.
    & $module {
        param($p)
        $script:rawToChange=$p
        $script:originalValidator=${function:Test-PreparedSourceTenants}
        function script:Test-PreparedSourceTenants {
            param([object[]]$Plan,[string]$TenantKey,[string]$SnapshotRoot)
            if(-not $SnapshotRoot){throw 'Validation unexpectedly reads live sources'}
            [IO.File]::WriteAllText($script:rawToChange,'changed by synthetic collector')
            & $script:originalValidator -Plan $Plan -TenantKey $TenantKey -SnapshotRoot $SnapshotRoot
        }
    } $csv
    try {$result=Invoke-PreparedEvidencePipeline @argsForRun}
    finally {& $module {Set-Item Function:script:Test-PreparedSourceTenants $script:originalValidator}}
    Check ($result.Identity.Rows -eq 1 -and -not $result.Publication) 'Post-capture raw update invalidated good snapshot'
    Check (-not(Test-Path (Join-Path $result.DiagnosticPath 'source'))) 'Successful preflight retained source payload'
    Check ((Test-Path (Join-Path $result.DiagnosticPath 'capture.json.txt')) -or (Test-Path (Join-Path $result.DiagnosticPath 'capture.json'))) 'Capture audit missing'
    Check (([IO.File]::ReadAllText($csv)) -ceq 'changed by synthetic collector') 'Cleanup touched live source'
    Check (-not(Test-Path $output)) 'Preflight published output'
    $bad=$argsForRun.Clone();$bad.WorkRoot=$fixture
    Reject {Invoke-PreparedEvidencePipeline @bad} 'must not contain DataRoot'
    # Failed source validation also cleans payload and retains diagnostics.
    [IO.File]::WriteAllText($csv,$valid.Replace('synthetic-test','wrong-test'))
    Reject {Invoke-PreparedEvidencePipeline @argsForRun} 'incompatible TenantKey'
    Check (@(Get-ChildItem $work -Recurse -Directory | Where-Object Name -In 'source','prepared').Count -eq 0) 'Failed run leaked source/staging payload'
    Check (@(Get-ChildItem $work -Recurse -File | Where-Object Name -in @('failure.json','failure.json.txt')).Count -eq 1) 'Failed run diagnostic missing'
    # Recover only explicitly marked interrupted runs, not unowned legacy payloads.
    $crashId=[guid]::NewGuid().ToString('N');$crash=Join-Path $work $crashId
    New-Item -ItemType Directory -Path (Join-Path $crash 'source'),(Join-Path $crash 'prepared'),(Join-Path $work 'unowned/source') -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $crash 'source/partial.csv'),'synthetic')
    [IO.File]::WriteAllText((Join-Path $work 'unowned/source/keep.csv'),'keep')
    @{Owner='PreparedEvidencePipeline/v1';RunId=$crashId;TenantKey='synthetic-test';DataRoot=$raw} | ConvertTo-Json | Set-Content (Join-Path $crash 'run.json')
    [IO.File]::WriteAllText($csv,$valid)
    $null=Invoke-PreparedEvidencePipeline @argsForRun
    Check (-not(Test-Path (Join-Path $crash 'source'))) 'Interrupted owned run was not recovered'
    Check (Test-Path (Join-Path $work 'unowned/source/keep.csv')) 'Unowned payload removed'
    Reject {& $module {param($w) Remove-PreparedOwnedPath $w '../raw'} $work} 'strictly below'
    Reject {& $module {param($w) Remove-PreparedOwnedPath $w '.'} $work} 'strictly below'
    $outside=Join-Path $fixture 'outside';New-Item -ItemType Directory -Path $outside -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $outside 'keep.txt'),'keep')
    New-Item -ItemType Junction -Path (Join-Path $work 'junction') -Target $outside | Out-Null
    Reject {& $module {param($w) Remove-PreparedOwnedPath $w 'junction'} $work} 'linked path'
    Check (Test-Path (Join-Path $outside 'keep.txt')) 'Linked target was changed'
    $foreignId=[guid]::NewGuid().ToString('N');$foreign=Join-Path $work $foreignId
    New-Item -ItemType Directory -Path (Join-Path $foreign 'source') -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $foreign 'source/keep.csv'),'keep')
    @{Owner='PreparedEvidencePipeline/v1';RunId=$foreignId;TenantKey='another-test';DataRoot=$raw} | ConvertTo-Json | Set-Content (Join-Path $foreign 'run.json')
    Reject {& $module {param($w,$id,$r) Clear-PreparedRunPayload $w $id $r 'synthetic-test'} $work $foreignId $raw} 'ownership mismatch'
    Check (Test-Path (Join-Path $foreign 'source/keep.csv')) 'Foreign-tenant payload removed'
    # Four publications retain only current + previous; audit and history stay.
    $staging=Join-Path $fixture 'staging';New-Item -ItemType Directory -Path $staging | Out-Null
    $contract=Join-Path $fixture 'output-contract.json'
    @{tables=@(@{table='Test Trend';file='Trend.csv';uniqueKey=@('Date');columns=@(@{name='Date';type='dateTime'},@{name='Value';type='int64'})})} | ConvertTo-Json -Depth 8 | Set-Content $contract
    $pub=@{StagingRoot=$staging;OutputRoot=$output;TenantKey='synthetic-test';Provenance=@{Mode='SyntheticTest'};ContractPath=$contract}
    $rawHash=(Get-FileHash $csv).Hash
    $published=@(foreach($i in 1..4){@([pscustomobject]@{Date='2026-01-01';Value=$i}) | Export-Csv (Join-Path $staging 'Trend.csv') -NoTypeInformation;Publish-PreparedEvidenceBatch @pub})
    $pointer=Get-Content (Join-Path $output 'current.json.txt') -Raw | ConvertFrom-Json
    Check ($pointer.BatchId -eq $published[3].BatchId -and $pointer.PreviousBatchId -eq $published[2].BatchId) 'Wrong protected batch pair'
    Check (@(Get-ChildItem (Join-Path $output 'batches') -Directory).Count -eq 2) 'Retention did not keep exactly two batches'
    Check (@(Get-ChildItem (Join-Path $output 'retired') -Directory).Count -eq 2) 'Retired audit missing'
    Check (@(Get-ChildItem (Join-Path $output 'retired') -Recurse -Filter '*.csv').Count -eq 0) 'Retired CSV payload retained'
    Check ((Get-FileHash $csv).Hash -eq $rawHash) 'Retention changed raw data'
    $pointerHash=(Get-FileHash (Join-Path $output 'current.json.txt')).Hash
    @([pscustomobject]@{Date='2026-01-01';Value='bad'}) | Export-Csv (Join-Path $staging 'Trend.csv') -NoTypeInformation
    Reject {Publish-PreparedEvidenceBatch @pub} 'Batch validation failed'
    Check ((Get-FileHash (Join-Path $output 'current.json.txt')).Hash -eq $pointerHash) 'Failure advanced current pointer'
    Check (@(Get-ChildItem (Join-Path $output 'batches') -Directory).Count -eq 2) 'Failure removed a good batch or leaked CSVs'
    Check (@(Get-ChildItem (Join-Path $output 'failed') -Recurse -Filter failure.json.txt).Count -eq 1) 'Publication failure diagnostic missing'
    $unknown=Join-Path $output 'batches/unowned'
    New-Item -ItemType Directory -Path $unknown | Out-Null
    [IO.File]::WriteAllText((Join-Path $unknown 'keep.csv'),'keep')
    & $module {param($o,$p) Remove-PreparedObsoleteBatches $o $p} $output $pointer
    Check (Test-Path (Join-Path $unknown 'keep.csv')) 'Unreferenced batch swept'
    # Only the exact downloaded mapping folder may be cleaned.
    $mapping=Join-Path $work ('mappings/'+[guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $mapping -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $mapping 'synthetic.xlsx'),'synthetic')
    Remove-PreparedMappingWorkbooks $work $mapping
    Check (-not(Test-Path $mapping)) 'Downloaded mapping payload retained'
    Reject {Remove-PreparedMappingWorkbooks $work $raw} 'exact run-owned'
    Write-Host "PASS: $script:checks stable capture, bounded retry, post-capture update, cleanup, retention, failure and path safety checks. Synthetic data only."
} finally {Remove-Module $module}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDFmATqDlLS2efk
# lGJyTRqCfAznqhGSyDB+GAmJMPe7D6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOSBtfOr0Er80lFLTPGstm/qiLCf2dSZXUNcMqKwnmddMA0GCSqG
# SIb3DQEBAQUABIIBgCsnllQcl+Zmnn+t/31DkeijVpAyhxRflvdfLpCQOriU8kwZ
# zrTIoXZuBsm85uLmLy4U8+pxLdhImlYF8O7Wk6jR2YvUT89MvunWzyp0y73psbr2
# Df1mIDtOiYcwWsW/11qm+2Q+2aZ/7viR+EQ5FbW4rZVxogNcOQqhD1ClhBVspgAv
# A+rHbnpSNgLKzHgfxpDqU/uw2dp49a6E5lbrQJGYkOvjtrBzyUQLVU1BRmsMdQNF
# 8XdHVnvFYXNRRxAiU4kxlFwQ/VsrthOkL/NwuGSNADors6YcvDZXl6fQeYfkzmEm
# RKfOw0Ka8SUbHcbyvqL167Tt6344twU1/jwgGYebntYM2ZNYjmgoGX+/D6YgQHqJ
# laQFR/zsvc+ggCOvB76zO9VBczHmTHtWJfsM/naSoFnXFAi2V47Jk1b9z+UEkzil
# qXcoVIZE5M/GqWGdDWaLvrdzJZmvv+YIl05ksnOksl+k6A+vIE2d0gHjycvTiDj4
# K2yhKclHx+0/3gqdoKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU4
# MDBaMC8GCSqGSIb3DQEJBDEiBCBoeOblxBKliNRDnbRLSMjQ1UQJCR4elrPxwWSM
# vv2+qzANBgkqhkiG9w0BAQEFAASCAgC0JWCoU9S2B8iZnKhl1BramTYdzSp3U64L
# ff0WF9Czjfhb/He9NMhbzdZMoyehXz0PKskypzoogTV/wi7CoWPgL5l+/SD2W7TN
# v7DB58YuVsj5cmoDGoTVwluzE98HhYUdSNTKHAiGkikOJLF+oYebunatkSESwebw
# JAOQQ1Bhm6OcRDuZZ4JsXkLDCxaTqJ1EFJ72FvTddTRPqTRGQ8CfL55NQfKKBMoM
# caZvVqjHRbuECh8ZEZzq8QsXVvENTvi4hydqH1k67yGzyc3emDNEafHpWlMxnnxc
# Fyu20vja+vKmY9RNVmnoGrFQqN0UIVJDGWz8tLNpTF1DoQABpdOqg/3tqAIRrWqP
# 9MMnwzUVmIePkdolDVaILJno3vKSpZicL2h5A8SZexfof05p+tWpO/hENzqGUItS
# OY1l76a/whn33Or4oWsRTUJ60ZGwnUuSXEzqtwk1f7sw96DGvk7zxxOu7ywAB/Kx
# HO83kWjvTnqe0qYLhjcniUYOhu3Kk/vpqASXJvWhj6YAH1wZL9gAMRTsSrEEKg23
# 1nRrKQ0bK233H7sgCv288t1JtyPsX1EbSvn9oESYJPwhuAHDPLaKqBhp3DAfFAbJ
# pBrJmW0GhOV1QwOXRU8M6dmPgzwgBlavC+ZmqA+glK182wyLZgWB9aAvb36UEWaC
# 0GM6zMAfxw==
# SIG # End signature block
