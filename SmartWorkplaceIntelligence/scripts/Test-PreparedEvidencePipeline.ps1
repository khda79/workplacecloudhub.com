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
$firstManifest=Get-Content (Join-Path $first.BatchPath 'batch.json.txt') -Raw | ConvertFrom-Json
if ($firstManifest.SharedMailboxLicenseEvidence.State -ne 'N/D' -or (Test-Path (Join-Path $first.BatchPath 'shared-mailbox-evidence-draft.json.txt'))) { throw 'Unqualified shared mailbox evidence or draft cleanup was mishandled.' }
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
# Rolling window: a week at or before the cutoff may leave; a later week still may not.
function IsoWeek([int]$WeeksAgo) { $d=[datetime]::UtcNow.Date.AddDays(-7*$WeeksAgo); '{0:D4}-W{1:D2}' -f [Globalization.ISOWeek]::GetYear($d),[Globalization.ISOWeek]::GetWeekOfYear($d) }
$expired=IsoWeek 12; $inside=IsoWeek 11; $current=IsoWeek 0
$output=Join-Path $weekRoot 'window\DATA-POWERBI'; $argsForPublish.OutputRoot=$output
WeekFixture @([pscustomobject]@{'Snapshot Date'='2026-09-24';'Week Label'=$expired;User='u1';'Last Activity Date'='2026-09-20'},[pscustomobject]@{'Snapshot Date'='2026-09-28';'Week Label'=$inside;User='u1';'Last Activity Date'='2026-09-27'},[pscustomobject]@{'Snapshot Date'='2026-10-01';'Week Label'=$current;User='u1';'Last Activity Date'='2026-09-30'})
$null=Publish-PreparedEvidenceBatch @argsForPublish -HistoryWindowWeeks 12 -InformationAction SilentlyContinue
WeekFixture @([pscustomobject]@{'Snapshot Date'='2026-10-01';'Week Label'=$current;User='u1';'Last Activity Date'='2026-09-30'})
ExpectFailure 'Lost week inside the window' { Publish-PreparedEvidenceBatch @argsForPublish -HistoryWindowWeeks 12 } "*missing Week Label=$inside*"
ExpectFailure 'Lost expired week without a window' { Publish-PreparedEvidenceBatch @argsForPublish } "*missing Week Label=$expired*"
WeekFixture @([pscustomobject]@{'Snapshot Date'='2026-09-28';'Week Label'=$inside;User='u1';'Last Activity Date'='2026-09-27'},[pscustomobject]@{'Snapshot Date'='2026-10-01';'Week Label'=$current;User='u1';'Last Activity Date'='2026-09-30'})
$windowed=Publish-PreparedEvidenceBatch @argsForPublish -HistoryWindowWeeks 12 -InformationAction SilentlyContinue
$expiration=@((Get-Content (Join-Path $windowed.BatchPath 'batch.json.txt') -Raw | ConvertFrom-Json).HistoryWindowExpirations)
if ($expiration.Count -ne 1 -or $expiration[0].Table -ne 'Week History' -or $expiration[0].WindowWeeks -ne 12 -or $expiration[0].CutoffWeek -ne $expired -or @($expiration[0].Weeks) -join ',' -ne $expired) { throw 'Window expiration was not recorded.' }
Write-Host 'PASS: a week past the rolling window leaves the history and is recorded; a week inside it is still protected.'
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
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB4aMrqbwZx+VtP
# Yk/HLpiux236ZamFuQBp01khDVMQ6KCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgDmPE6Yr6bevo
# EBe9Ae9HsISOKCk3VQZkSWg+/idywJkwDQYJKoZIhvcNAQEBBQAEggGAlN6g+tkh
# /UGQKtCwQxuZeUDeQRo1n+qNpUgDpJ+aQzINSYJmsDIdPiqDiY86e07UFmwa4y1q
# qpUylDGvAjTNZv49NABqHvUBo1p/+iqyNFlILhQtbtu3MpEsHdNsaem7jGq6N7f/
# nlgARXOejPHyQydEBFLfLfC7+bJm4yTKaDEO2/DK8HtJUh++6NNZEfw+3bnP1xYl
# EWqL/h8Gq5QGP5o9pRsllBGrz41ExJ8LFYJwvsWoSh1fPIAnaXNPSyDtNk57ZCGi
# iE1a3mvu8xFVYOmrsayNFrZeTXCkPsGQx/WIPom8DVMCzOZWie7Y7FOjUZB4BIAz
# EJjBu+milnP9/7xCwYEUT7qLGTLbSAXmg1eGyTmWLpUqm9SvR02o57XtTh7cORYw
# ote8O9TF0lKjWFIeoRvOfJuVxKZQ58t+pMpDtGYcvWCX07ZNWtnjrHqsiyxaFbT4
# Y+74ziMq1fSKofVLkoVngLZlnLTzibEBBM6kNcMgaT+u4QxAtrC8U4zIoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTA1MTU0NzMzWjA/BgkqhkiG9w0BCQQxMgQwOpovsAXogO5EwwhP00AbWfch
# 72gODIe5XNG0rssXbXrjvi6U6nAQdaUx2WoHUPhHMA0GCSqGSIb3DQEBAQUABIIC
# AAnsP5G7JXdL5n48CvIC4O8J5V7by3XrnqJHjkY6arf0eivawOg3QJ5ogvVhERZI
# oD5L8Ypqh9qvoJTcpsNNZSgk6Z3mbrgKnkkNTZUvV8wYdS/imoBvnVuO4wMJkf01
# bTqaPX9DeSImHB7tb3tSd7jmoAQG1i+ia8HtgTZyUogTbnK/vlvAo8YiW9HEzpXu
# ARGE9g5R1xbJup/xwL60HE0v8jY/fS/7NEB6keK83xB3wF2daG2c5FE5fHSF7Zi7
# AoOzI8WE5e1MDheCi6QzwdsEJrVBbU30KpkCcgEWlIv14sxV9Vm020TFEvQTbN7t
# TRlIWCbWDM0uHJD0in+uMhUKxsl4R/dWxcRuU4VgU+SnQ/A0iVMmvaWqFyOTLGvh
# A9lCGcKUb5wVIxzkpvB/sfawHSr3IL49a2Qpucv69o3odsmNXcEOKXAOwQ7XD9l9
# KETl53NjLC0iFap6F0gJJ0x5+k82C8eYOywrhVVT21QNz1Uasl6CFUW3Fb5c5kFS
# Yx3l+N3h8emFP/6ZzP1/KqmvuY7n+/AjlShDNjMQ2oR+cBs0owcrmSdStSd7V8Nc
# 9n9Ah+r4oxnY3+opX6H2AruZjzm4HwGNqST1lVuyjURMecw2oCZknOKBUAFE3TqR
# 19m/qq3XLIpB1iwmdMSHddeGVQb7ByWXysIqfFOEq3/U
# SIG # End signature block
