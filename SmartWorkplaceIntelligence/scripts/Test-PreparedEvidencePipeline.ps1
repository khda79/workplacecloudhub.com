[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$staging=Join-Path $root 'staging'
$output=Join-Path $root 'DATA-POWERBI'
New-Item -ItemType Directory -Path $staging -Force | Out-Null
$contract=Join-Path $root 'contract.json'
@{tables=@(@{table='Test Trend';file='Trend.csv';historyKey=@('Date');uniqueKey=@('Date');columns=@(@{name='Date';type='dateTime'},@{name='Value';type='int64'},@{name='Note';type='string'})})} | ConvertTo-Json -Depth 8 | Set-Content $contract
function Fixture($Rows) { $Rows | Export-Csv (Join-Path $staging 'Trend.csv') -NoTypeInformation }
function ExpectFailure([string]$Name,[scriptblock]$Action,[string]$Like='*') {
    $before=(Get-FileHash (Join-Path $output 'current.json.txt')).Hash
    $failed=$false
    try { & $Action | Out-Null } catch { $failed=$_.Exception.Message -like $Like }
    if (-not $failed) { throw "Expected rejection: $Name" }
    if ((Get-FileHash (Join-Path $output 'current.json.txt')).Hash -ne $before) { throw "Last good pointer changed: $Name" }
    Write-Host "PASS: $Name rejected; last good batch preserved."
}
$argsForPublish=@{StagingRoot=$staging;OutputRoot=$output;TenantKey='synthetic-test';Provenance=@{Mode='SyntheticTest'};ContractPath=$contract}
Fixture @([pscustomobject]@{Date='2026-01-01';Value='2';Note="comma, quote `" and`nnewline"})
$first=Publish-PreparedEvidenceBatch @argsForPublish
$row=Import-Csv (Join-Path $first.BatchPath 'Trend.csv')
if ($row.TenantKey -ne 'synthetic-test' -or $row.Note -notmatch "`n") { throw 'CSV roundtrip failed.' }
Write-Host 'PASS: logical CSV records and tenant prefix preserved.'
Fixture @([pscustomobject]@{Date='2026-01-01';Value='invalid';Note='x'})
ExpectFailure 'Invalid numeric value' { Publish-PreparedEvidenceBatch @argsForPublish }
Fixture @([pscustomobject]@{Date='2026-01-02';Value='2';Note='x'})
ExpectFailure 'Historical date regression' { Publish-PreparedEvidenceBatch @argsForPublish }
Fixture @([pscustomobject]@{Date='2026-01-01';Value='2';Note='x'},[pscustomobject]@{Date='2026-01-01';Value='3';Note='x'})
ExpectFailure 'Duplicate keys' { Publish-PreparedEvidenceBatch @argsForPublish }
Fixture @([pscustomobject]@{Date='2026-01-01';Value='2';Note='x'})
$argsForPublish.TenantKey='another-test'
ExpectFailure 'Cross-tenant output' { Publish-PreparedEvidenceBatch @argsForPublish }
$argsForPublish.TenantKey='synthetic-test'
$lock=[IO.File]::Open((Join-Path $output '.publication.lock'),'OpenOrCreate','ReadWrite','None')
try { ExpectFailure 'Concurrent publisher' { Publish-PreparedEvidenceBatch @argsForPublish } } finally {$lock.Dispose()}
[IO.File]::WriteAllText((Join-Path $staging 'Trend.csv'),'"Date","Value","Note"'+[Environment]::NewLine)
ExpectFailure 'Unexpected empty table' { Publish-PreparedEvidenceBatch @argsForPublish }
Fixture @([pscustomobject]@{Date='2026-01-01';Value='2';Note='x'},[pscustomobject]@{Date='2026-01-02';Value='3';Note='x'})
$next=Publish-PreparedEvidenceBatch @argsForPublish
$pointer=Get-Content (Join-Path $output 'current.json.txt') -Raw | ConvertFrom-Json
if ($pointer.PreviousBatchId -ne $first.BatchId -or -not(Test-Path $first.BatchPath)) {throw 'Previous version lost.'}
Write-Host 'PASS: complete replacement retains previous batch.'
ExpectFailure 'Changed base batch during staging' { Publish-PreparedEvidenceBatch @argsForPublish -ExpectedPreviousBatchId $first.BatchId } '*Published batch changed*'
$guarded = Publish-PreparedEvidenceBatch @argsForPublish -ExpectedPreviousBatchId $next.BatchId -SkipRetention
if (-not (Test-Path -LiteralPath $first.BatchPath) -or -not (Test-Path -LiteralPath $next.BatchPath)) { throw 'SkipRetention removed a prior batch.' }
$next = $guarded
Write-Host 'PASS: guarded publication retains all prior batches when requested.'
function Contract([int]$KeyVersion) {
    @{tables=@(@{table='Test Trend';file='Trend.csv';historyKey=@('Date');historyKeyVersion=$KeyVersion;uniqueKey=@('Date');columns=@(@{name='Date';type='dateTime'},@{name='Value';type='int64'},@{name='Note';type='string'})})} | ConvertTo-Json -Depth 8 | Set-Content $contract
}
Contract 2
Fixture @([pscustomobject]@{Date='2026-01-05';Value='4';Note='x'})
$regrained=Publish-PreparedEvidenceBatch @argsForPublish -WarningAction SilentlyContinue
$manifest=Get-Content (Join-Path $regrained.BatchPath 'batch.json.txt') -Raw | ConvertFrom-Json
$reset=@($manifest.HistoryKeyResets)
if ($reset.Count -ne 1 -or $reset[0].Table -ne 'Test Trend' -or $reset[0].PreviousVersion -ne 1 -or $reset[0].Version -ne 2 -or $reset[0].PreviousBatchId -ne $next.BatchId -or $manifest.HistoryKeyVersions.'Test Trend' -ne 2) { throw 'Key-version reset was not recorded.' }
Write-Host 'PASS: contract key-version change resets the comparison once and is recorded.'
Fixture @([pscustomobject]@{Date='2026-01-12';Value='5';Note='x'})
ExpectFailure 'Historical date regression after key-version change' { Publish-PreparedEvidenceBatch @argsForPublish }
Contract 1
Fixture @([pscustomobject]@{Date='2026-01-05';Value='4';Note='x'})
ExpectFailure 'Key-version rollback' { Publish-PreparedEvidenceBatch @argsForPublish }
$module=Get-Module PreparedEvidencePipeline
$moveRoot=Join-Path $root 'move'
New-Item -ItemType Directory -Path $moveRoot -Force | Out-Null
$target=Join-Path $moveRoot 'current.json.txt'
$marker=Join-Path $moveRoot 'locked.marker'
Set-Content -LiteralPath $target -Value 'old' -NoNewline
Set-Content -LiteralPath (Join-Path $moveRoot 'next.tmp') -Value 'new' -NoNewline
$holder=Start-ThreadJob -ScriptBlock {
    $stream=[IO.File]::Open($using:target,'Open','Read','None')
    try { Set-Content -LiteralPath $using:marker -Value 'locked'; Start-Sleep -Milliseconds 2500 } finally { $stream.Dispose() }
}
while (-not (Test-Path -LiteralPath $marker)) { Start-Sleep -Milliseconds 50 }
$retries=@(& $module { param($s,$d) Move-PreparedFile -Source $s -Destination $d -Overwrite -RetryDelaySeconds @(1,1,1,1,1) } (Join-Path $moveRoot 'next.tmp') $target 3>&1)
$holder | Wait-Job | Remove-Job
if ((Get-Content -LiteralPath $target -Raw) -ne 'new' -or $retries.Count -lt 1) { throw 'Transient lock was not retried.' }
Write-Host "PASS: transiently locked target replaced after $($retries.Count) traced retry(ies)."
Set-Content -LiteralPath (Join-Path $moveRoot 'next.tmp') -Value 'newer' -NoNewline
$lock=[IO.File]::Open($target,'Open','Read','None')
try {
    $failed=$false
    try { & $module { param($s,$d) Move-PreparedFile -Source $s -Destination $d -Overwrite -RetryDelaySeconds @(0,0) -WarningAction SilentlyContinue } (Join-Path $moveRoot 'next.tmp') $target }
    catch { $failed=$_.Exception.Message -like "*after 3 attempt(s): $target*" }
} finally { $lock.Dispose() }
if (-not $failed -or (Get-Content -LiteralPath $target -Raw) -ne 'new' -or -not (Test-Path -LiteralPath (Join-Path $moveRoot 'next.tmp'))) { throw 'Persistent lock was not rejected with the target path.' }
Write-Host 'PASS: persistently locked target rejected with its path; source and target preserved.'
# Weekly history: the declared key is the week; provisional snapshot dates and per-user values may change.
$weekRoot=Join-Path $root 'weekly'
$staging=Join-Path $weekRoot 'staging'; $output=Join-Path $weekRoot 'DATA-POWERBI'
New-Item -ItemType Directory -Path $staging -Force | Out-Null
$contract=Join-Path $weekRoot 'contract.json'
function WeekContract([switch]$WithoutKey) {
    $table=@{table='Week History';file='WeekHistory.csv';uniqueKey=@('Week Label','User');columns=@(@{name='Snapshot Date';type='dateTime'},@{name='Week Label';type='string'},@{name='User';type='string'},@{name='Last Activity Date';type='dateTime'})}
    if (-not $WithoutKey) { $table.historyKey=@('Week Label') }
    @{tables=@($table)} | ConvertTo-Json -Depth 8 | Set-Content $contract
}
function WeekFixture($Rows) { $Rows | Export-Csv (Join-Path $staging 'WeekHistory.csv') -NoTypeInformation }
$argsForPublish=@{StagingRoot=$staging;OutputRoot=$output;TenantKey='synthetic-test';Provenance=@{Mode='SyntheticTest'};ContractPath=$contract}
WeekContract
WeekFixture @([pscustomobject]@{'Snapshot Date'='2026-09-24';'Week Label'='2026-W39';User='u1';'Last Activity Date'='2026-09-20'},[pscustomobject]@{'Snapshot Date'='2026-09-28';'Week Label'='2026-W40';User='u1';'Last Activity Date'='2026-09-27'})
$null=Publish-PreparedEvidenceBatch @argsForPublish
WeekFixture @([pscustomobject]@{'Snapshot Date'='2026-09-24';'Week Label'='2026-W39';User='u1';'Last Activity Date'='2026-09-20'},[pscustomobject]@{'Snapshot Date'='2026-10-01';'Week Label'='2026-W40';User='u1';'Last Activity Date'='2026-09-30'})
$null=Publish-PreparedEvidenceBatch @argsForPublish
Write-Host 'PASS: current-week snapshot date and per-user activity changes are accepted.'
WeekFixture @([pscustomobject]@{'Snapshot Date'='2026-10-01';'Week Label'='2026-W40';User='u1';'Last Activity Date'='2026-09-30'})
ExpectFailure 'Lost historical week' { Publish-PreparedEvidenceBatch @argsForPublish } '*missing Week Label=2026-W39*'
WeekContract -WithoutKey
WeekFixture @([pscustomobject]@{'Snapshot Date'='2026-09-24';'Week Label'='2026-W39';User='u1';'Last Activity Date'='2026-09-20'},[pscustomobject]@{'Snapshot Date'='2026-10-01';'Week Label'='2026-W40';User='u1';'Last Activity Date'='2026-09-30'})
ExpectFailure 'History table without a declared key' { Publish-PreparedEvidenceBatch @argsForPublish } '*No historical comparison key*'
Write-Host 'All synthetic publication tests passed. No tenant API or Power BI access.'

# An observed Windows 10 population can legitimately become empty. This explicit
# per-table contract does not weaken the existing unexpected-empty protection.
$zeroRoot=Join-Path $root 'qualified-zero'
$staging=Join-Path $zeroRoot 'staging';$output=Join-Path $zeroRoot 'DATA-POWERBI'
New-Item -ItemType Directory -Path $staging -Force | Out-Null
$contract=Join-Path $zeroRoot 'contract.json.txt'
$zeroContract=@{tables=@(@{table='Windows Migration Evidence';file='WindowsMigrationEvidence.csv';allowEmpty=$true;uniqueKey=@('Device Source ID');columns=@(@{name='Device Source ID';type='string'})})}
$zeroContract | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $contract
[IO.File]::WriteAllText((Join-Path $staging 'WindowsMigrationEvidence.csv'),'"Device Source ID"'+[Environment]::NewLine)
$argsForPublish=@{StagingRoot=$staging;OutputRoot=$output;TenantKey='synthetic-test';Provenance=@{Mode='SyntheticTest'};ContractPath=$contract}
$zero=Publish-PreparedEvidenceBatch @argsForPublish
if((Get-Content -LiteralPath (Join-Path $zero.BatchPath 'batch.json.txt') -Raw | ConvertFrom-Json).Files[0].Rows -ne 0){throw 'Qualified zero publication failed.'}
Write-Host 'PASS: explicit migration zero allowed with a header, validation and tenant receipt.'
$zeroContract.tables[0].allowEmpty='true'
$zeroContract | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $contract
ExpectFailure 'Non-boolean empty-table allowance' { Publish-PreparedEvidenceBatch @argsForPublish } '*Unexpected empty output*'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCIxuYqBB8EXHS+
# cpytPEX4uLjyx+GIYlTtBiHCFOyD66CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHTj19idu7Fdg4xWfHlc7DKY1rBxqZiJ74q0VKyi0Cv0MA0GCSqG
# SIb3DQEBAQUABIIBgIxdmppEgLLovyE8I33RD2EPkdfMu/FeIV5Vhbvtq/OJE+Bj
# n1rI2nVOCWWIMneaul6Jc76f9LZ8j3/k1cOhQjJFX1+3X/ihlgEt7KNtoMxRWdwc
# UTAe+rPXyH9sxev6rmcWup/prBi3LQUBh3mF+LJWuIgvgmRiFHznQ7UxAqILn3ON
# scXyOLY075FFEH0KhZslnhsvm2WUDVXi69zGqQnfWXCJL8YGWCe66FkaiF1V02vD
# X5voUUXwx7LTvF2lcG5hNyDru5m2PU3ZYBLvUAS4xP1GlQMhkGrMzWBexTyJ9vQ0
# T7Wwbt9kI1YSuO6rH/IRcNrSFpixPIo/rLuPZxk3kJhx5fwBkVCVXF/mvt99wq0d
# C1b+YpoCoxABn0JfiAT4rK1NlGcEnCrUiFBhRz1TFyM1aFpu3qIVz0pMsIRw1CBE
# l63RjNMsJlzsBULgULlzgz+1/CFBLKF8HA1LuOKXWOdhvPvMnZGDZQ2MJkKrxSIS
# D8fi2RIHZmuddREcQaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDkxNDI1
# MTZaMC8GCSqGSIb3DQEJBDEiBCAYfw3Tm8G8lC1RAV+vCRQEyNQm8cQdgyKUncSR
# RZDp8TANBgkqhkiG9w0BAQEFAASCAgBYD7TDIexP3Dky/CelFcGe+2bUEGa3YldS
# FteUg94DwbTm6wg+TpR0pWLY5SJuGlV0p1dDtzhxoPAHR2Ni2ecXmgb7+WMoY3+Z
# SnYQuDZ/qTyeMKXvuCcFzTUoEkVVyqmctHtAjR+TLuQIaZdg44T3QU+wY1aYuFxa
# l9liraZcHL80BY9e8XBXgQDctn+BLpYAgJgYCBvH9PcOOkpvY3SbZn2P3xqSTqSZ
# AUSc9pfShk9+t1KP9HQNUN3cFNuhtf3Lx1xx3qVZE9kIxsyZkDTuBMt6AgHOWppf
# XW6FsdfQsCwX2P2zDvAaMcC51YtB+TGK2wFZyXTrA5viym3HZLhr83bj/7jfodBW
# WiPabRjQJo5OJayDuFgq+Rexg9tbucusdVadZ1esaTYuJW0fTZYZabZGCdRCG4X2
# Abpukx69NsT7HxKGiY7IFmy3x2ZIRqXTJjh2G8L8WN5sPrLDKbL3r7dartd6nH41
# mZe7l92NnwWSuXyluOMke33zJYju7oOnYtrgJ1TW+3rtw3rkkJv2y692ftbni99h
# dFfdACpzpI68/0H1uHkoAhGhDGsezXiY7GT0W4RWmCCHBy5goLhfTEb0uAMaPK4I
# jilZL4DqihqD5qvocegMX7JfXBz2aUUS08LaSqKrWVLfjx+49cvLRsoVoGOblpRZ
# kuNtF8stmw==
# SIG # End signature block
