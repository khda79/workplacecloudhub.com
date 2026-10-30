#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of the parallel live statistics prefetch of the EXO mailboxes producer: parallel values
equal the sequential ones, partial errors, simulated throttling, worker connection failure, deadline,
one persistent session per worker, and unchanged Fast and permissions-only paths.
.VERSION
1.0.1
.NOTES
Synthetic mailboxes and mocked Exchange cmdlets only. No tenant, Exchange, API, mail or synchronized data access.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$producerPath = Join-Path $root 'SmartInventory/ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1'
$evidencePath = Join-Path $root 'SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

$text = [IO.File]::ReadAllText($producerPath)
$ast = [Management.Automation.Language.Parser]::ParseFile($producerPath, [ref]$null, [ref]$null)
$evidenceAst = [Management.Automation.Language.Parser]::ParseFile($evidencePath, [ref]$null, [ref]$null)
function Get-FunctionText($Tree, [string]$Name) { $Tree.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true).Extent.Text }

# ---------- Static wiring ----------
Check ($text -match '\.VERSION\r?\n1\.30\r?\n' -and $text -match '\$ScriptVersion = "1\.30"') 'EXO mailboxes version is not 1.30.'
Check ($text -match "Get-ScriptLocalConfigValue -Config \`$ScriptLocalConfig -Name 'StatsParallelThrottle' -DefaultValue 4") 'StatsParallelThrottle does not default to 4 from the local configuration.'
$template = Get-Content -LiteralPath (Join-Path (Split-Path $producerPath) 'SmartM365-EXO-Mailboxes-Inventory.local.json.template') -Raw | ConvertFrom-Json
Check ($template.StatsParallelThrottle -eq 4) 'The local configuration template does not expose StatsParallelThrottle = 4.'
Check ($text -match 'if \(\$UseLiveStats -and \$StatsParallelThrottleEffective -gt 1 -and \(-not \$InteractiveAuth\)') 'The prefetch is not limited to live stats with a throttle above 1.'
Check ($text -match '\$UseLiveStats = \(\$IncludeStats -and \(-not \$ExcludeStats\) -and \(-not \$PermissionsOnly\)\)') 'Live stats can now run with Fast or permissions-only.'
Check ($text -match '\$prefetchedStats\.ContainsKey\(\$counter - 1\)' -and $text -match '\$statsIndex = -1\r?\n\s*foreach \(\$statsMailbox in \$mailboxes\) \{\r?\n\s*\$statsIndex\+\+') 'Prefetched values are not keyed by the sequential loop position.'
Check ($text -match 'if \(\$null -eq \$liveStats\) \{ \$liveStats = Get-MailboxLiveStatsRecord -Mailbox \$mb') 'A mailbox missing from the prefetch is no longer collected sequentially.'
Check ($text -match "catch \[System\.Management\.Automation\.PipelineStoppedException\] \{ throw \}") 'Cancellation during the prefetch is swallowed.'
Check ($text -match '(?s)Live stats prefetch completed.*?try \{ Invoke-EXOProactiveReconnect \}') 'The main session is not refreshed after the worker sessions.'
$prefetchText = Get-FunctionText $ast 'Invoke-MailboxStatsParallelPrefetch'
Check ($prefetchText -notmatch 'Disconnect-ExchangeOnline') 'A stats worker disconnects and would clear sibling EXO connections.'
Check (@($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-ExoStatsExact' }, $true)).Count -eq 2) 'Live stats are collected outside Get-MailboxLiveStatsRecord.'
Check ($text -match 'workers=\{0\}, dispatched=\{1\}, prefetched=\{2\}.*seconds=\{8\}, throughput=\{9\} mailboxes/min') 'Worker count, duration and throughput are not logged.'

# ---------- Runtime with mocked Exchange cmdlets ----------
$state = [System.Collections.Concurrent.ConcurrentDictionary[string,int]]::new()
[AppDomain]::CurrentDomain.SetData('SmartM365ParallelStatsTest', $state)
$mocks = @'
function Get-SyntheticNumber([string]$Identity) { [int]([regex]::Match($Identity, '\d+').Value) }
function New-SyntheticSize([double]$Bytes) {
    $value = [pscustomobject]@{ Bytes = $Bytes }
    $value | Add-Member -MemberType ScriptMethod -Name ToBytes -Value { [double]$this.Bytes }
    [pscustomobject]@{ Value = $value }
}
function Test-SyntheticThrottled([string]$Identity) {
    $state = [AppDomain]::CurrentDomain.GetData('SmartM365ParallelStatsTest')
    $calls = $state.AddOrUpdate("calls|$Identity", 1, { param($k, $v) $v + 1 })
    if ($Identity -like '*always-empty*') { return $true }
    $limit = 0
    if ($state.TryGetValue("throttle|$Identity", [ref]$limit)) { return ($calls -le $limit) }
    return $false
}
function Get-EXOMailboxStatistics {
    [CmdletBinding()]
    param([string]$Identity, [switch]$Archive, [string[]]$Properties)
    if ($Identity -like 'broken*') { throw "Synthetic terminating failure for $Identity" }
    if (Test-SyntheticThrottled $(if ($Archive) { "archive|$Identity" } else { $Identity })) { return $null }
    $n = Get-SyntheticNumber $Identity
    $result = [pscustomobject]@{ TotalItemSize = (New-SyntheticSize (($n + 1) * 123456789.0 * $(if ($Archive) { 3 } else { 1 }))); ItemCount = $n * 10 + $(if ($Archive) { 1 } else { 0 }) }
    if ($n % 5 -ne 0) { $result | Add-Member -NotePropertyName LastLogonTime -NotePropertyValue ([datetime]::new(2026, 1, 1, 8, 0, 0, [DateTimeKind]::Utc).AddMinutes($n)) }
    if ($Properties -contains 'LastUserActionTime') { $result | Add-Member -NotePropertyName LastUserActionTime -NotePropertyValue ([datetime]::new(2026, 2, 1, 8, 0, 0, [DateTimeKind]::Utc).AddMinutes($n)) }
    $result
}
function Get-MailboxStatistics {
    [CmdletBinding()]
    param([string]$Identity, [switch]$Archive)
    if (Test-SyntheticThrottled ("legacy|" + $Identity)) { return $null }
    $n = Get-SyntheticNumber $Identity
    [pscustomobject]@{ LastLogonTime = [datetime]::new(2025, 6, 1, 0, 0, 0, [DateTimeKind]::Utc).AddHours($n); ItemCount = $n; TotalItemSize = '1.5 GB (1,610,612,736 bytes)' }
}
'@
$definitions = (@(
        (Get-FunctionText $ast 'Convert-QuotaToGB'),
        (Get-FunctionText $evidenceAst 'ConvertTo-IdentityNativeIso'),
        (Get-FunctionText $ast 'Get-ExoStatsExact'),
        (Get-FunctionText $ast 'Get-MailboxLiveStatsRecord')
    ) -join "`n") + "`n" + $mocks
. ([scriptblock]::Create($definitions))
. ([scriptblock]::Create($prefetchText))
$connect = @'
param($Label)
$state = [AppDomain]::CurrentDomain.GetData('SmartM365ParallelStatsTest')
[void]$state.AddOrUpdate('connect', 1, { param($k, $v) $v + 1 })
if ($Label -eq 'fail') { throw 'Synthetic EXO connection refusal' }
'@
function New-Items([string[]]$Identities) {
    $index = -1
    foreach ($identity in $Identities) {
        $index++
        $status = if ($index % 3 -eq 0) { 'Active' } else { 'None' }
        [pscustomobject]@{ Index = $index; Mailbox = [pscustomobject]@{ UserPrincipalName = $identity; ArchiveStatus = $status } }
    }
}
function Get-ValueJson($Values) { $Values | Select-Object ArchiveSizeGB, ArchiveCount, ArchiveLastLogon, ArchiveLastLogonIso, TotalSizeGB, ItemCount, LastLogon, LastLogonIso, LastAction, LastActionIso | ConvertTo-Json -Compress }
function Invoke-Prefetch($Items, [int]$Throttle = 4, [string]$Label = 'ok', [datetime]$Deadline = [datetime]::MaxValue, [int]$EmptyRetries = 2, [switch]$Luat) {
    Invoke-MailboxStatsParallelPrefetch -Mailboxes $Items -ThrottleLimit $Throttle -WorkerDefinitions $definitions -ConnectScript $connect -ConnectArguments @($Label) `
        -ConnectRetries 1 -IncludeLastUserActionTime:$Luat -EmptyResultRetries $EmptyRetries -RetryDelaySeconds 0 -DeadlineUtc $Deadline -ProgressInterval 10 `
        -OnProgress { param($done, $all, $elapsed) $script:progress.Add($done) } -OnWorkerFailure { param($id, $message) $script:workerFailures.Add("$id|$message") }
}
function Reset-State { $state.Clear(); $script:progress = [Collections.Generic.List[int]]::new(); $script:workerFailures = [Collections.Generic.List[string]]::new() }

# Parallel values equal the sequential values, keyed by loop position, with one session per worker.
$identities = @(1..41 | ForEach-Object { 'user{0:000}@contoso.test' -f $_ })
foreach ($luat in @($false, $true)) {
    Reset-State
    $items = @(New-Items $identities)
    $sequential = @{}
    foreach ($item in $items) { $sequential[$item.Index] = Get-ValueJson (Get-MailboxLiveStatsRecord -Mailbox $item.Mailbox -IncludeLastUserActionTime:$luat) }
    Reset-State
    $summary = Invoke-Prefetch $items -Luat:$luat
    Check ($summary.WorkerCount -eq 4 -and $summary.Dispatched -eq 41 -and $summary.Prefetched -eq 41 -and $summary.Results.Count -eq 41) "Parallel prefetch did not return every mailbox (LUAT=$luat)."
    foreach ($item in $items) {
        Check ((Get-ValueJson $summary.Results[$item.Index]) -ceq $sequential[$item.Index]) "Parallel values differ from sequential values for position $($item.Index) (LUAT=$luat)."
    }
    $connects = 0; [void]$state.TryGetValue('connect', [ref]$connects)
    Check ($connects -eq 4) 'Workers did not open exactly one persistent session each.'
    Check (($summary.Retried + $summary.Failed + $summary.EmptyAfterRetries + $summary.ConnectFailedWorkers + $summary.DeadlineSkipped) -eq 0) 'A clean run reported retries or failures.'
    Check ((@($script:progress) -join ',') -eq '10,20,30,40') 'Prefetch progress was not reported at the interval.'
}
$sample = $summary.Results[0]
Check ($sample.TotalSizeGB -eq 0.23 -and $sample.ArchiveCount -eq 11 -and $sample.LastLogonIso -like '2026-01-01T08:01:00*' -and $sample.LastActionIso -like '2026-02-01T08:01:00*') 'Synthetic values were not converted as expected.'
Check ($summary.Results[4].LastLogonIso -like '2025-06-01T05:00:00*') 'The legacy LastLogonTime fallback was not used.'

# A single worker request (throttle 2 over one mailbox) still uses one worker.
Reset-State
$one = Invoke-Prefetch @(New-Items @('user007@contoso.test')) -Throttle 2
Check ($one.WorkerCount -eq 1 -and $one.Results.Count -eq 1) 'One mailbox did not use a single worker.'

# Partial errors and simulated throttling: retried values are kept, failures are left to the sequential loop.
Reset-State
[void]$state.TryAdd('throttle|user002@contoso.test', 2)
[void]$state.TryAdd('throttle|legacy|user002@contoso.test', 2)
$mixed = @(New-Items @('user001@contoso.test', 'user002@contoso.test', 'broken003@contoso.test', 'always-empty004@contoso.test', 'user005@contoso.test'))
$summary = Invoke-Prefetch $mixed -Throttle 2
Check ($summary.Results.ContainsKey(1) -and $summary.Retried -ge 1) 'A throttled mailbox was not retried into a complete value.'
Check (-not $summary.Results.ContainsKey(2) -and $summary.Failed -eq 1) 'A failing mailbox was not left to the sequential loop.'
Check (-not $summary.Results.ContainsKey(3) -and $summary.EmptyAfterRetries -eq 1) 'An always-empty mailbox was not left to the sequential loop after its retries.'
Check ($summary.Results.ContainsKey(0) -and $summary.Results.ContainsKey(4) -and $summary.Prefetched -eq 3) 'Healthy mailboxes were lost next to failures.'
$calls = 0; [void]$state.TryGetValue('calls|always-empty004@contoso.test', [ref]$calls)
Check ($calls -eq 3) 'The empty-result retry count is not EmptyResultRetries + 1 primary calls.'

# Worker connection failure: no values, one failure callback per worker, every mailbox stays sequential.
Reset-State
$summary = Invoke-Prefetch @(New-Items $identities) -Label 'fail'
Check ($summary.ConnectFailedWorkers -eq 4 -and $summary.Results.Count -eq 0 -and $script:workerFailures.Count -eq 4) 'Failed worker sessions were not reported or produced values.'
Check (@($script:workerFailures | Where-Object { $_ -match 'Synthetic EXO connection refusal' }).Count -eq 4) 'The worker connection error was not reported.'

# Deadline (MaxRunMinutes) reached: workers stop before calling Exchange.
Reset-State
$summary = Invoke-Prefetch @(New-Items $identities) -Deadline ([datetime]::UtcNow.AddMinutes(-1))
$anyCall = @($state.Keys | Where-Object { $_ -like 'calls|*' }).Count
Check ($summary.DeadlineSkipped -eq 41 -and $summary.Results.Count -eq 0 -and $anyCall -eq 0) 'Workers kept calling Exchange after the deadline.'

# Empty input returns an empty summary without workers.
$empty = Invoke-MailboxStatsParallelPrefetch -Mailboxes @() -ThrottleLimit 4 -WorkerDefinitions $definitions -ConnectScript $connect
Check ($empty.WorkerCount -eq 0 -and $empty.Results.Count -eq 0) 'An empty prefetch started workers.'

[AppDomain]::CurrentDomain.SetData('SmartM365ParallelStatsTest', $null)
Write-Host ("PASS: {0} parallel live stats offline checks. Synthetic mailboxes only; no tenant, Exchange or publication." -f $script:checks)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBu+mFmGBf+Z2J2
# 7hQR/cVkVg+wKV4hUEK1J/SL503aSKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBGpCmrpSi1yry6TqbKhbFVGfxjAPPWvO/0zvG0BfuTgMA0GCSqG
# SIb3DQEBAQUABIIBgHijLZawNJYAGGHWAiMsegiBW6gQ6sXrRMbdKx3IwQDZTEnq
# HOxRwr9Pb0NP4Mx+j4+eXVbKIJAasRp0a7qkrlU+17XRCk/SHlLeB/K6kYGIP8LE
# qD8VgqXUF7rjI7GAtO+YUXZUdloJs881Eqv/MyTfyeLA0F3rxuVOBWkKzAf7bNTP
# CsvE3UKsIZ+3LdbATaeJoKl9gEhKjvaSaHfaYrW/A/YZbkv46Al20YQsmV9u54SO
# z5FlAktNnFRZ+Kma0NLgBFyvuqCC4pPAV6U24ANDnWlOZPUVtwV5YoMd8bkIG8nr
# sI52WX90WmEOOrvdbwljfkRf1p2N3B/tQgpbPZSCjzf1CRWMxWJOYIKXeBplrCB0
# tFxmciJqXLQnPxcG7r4pRl+n5SJszFKCVFn4Jo+MyZ3oE65sxjvZPwhvpMg5m6CK
# nzLx8XH1CVcbKt5jP6Qo6w+l3AAuMq/ktp67zBrQXsBp0qb1JUDqJUOEPm2ggzV7
# NZhNGvvaPrlczTSs26GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwNzU3
# MzBaMC8GCSqGSIb3DQEJBDEiBCC3OtEYxtwCwIghi0b5hj4nLEyiDX3RN7jg3Zx0
# AsLQ0TANBgkqhkiG9w0BAQEFAASCAgBM4AZKgAjSpi7fs2t1PzjdV9wXhmh/hEGR
# uK0srjven7bTecqEmgQ9wWUjnAahu3o12imNfJQKP9CcgrXweYJAZOYqZuWbgUXI
# +yzNZC1A8+xdLirMToC1ffpyNw/Jk8GrTnHomA8jz8jScFz6wu9+F5vWuy8PO4MI
# JLQGNzy9EB07vdA5kYHmun/NVJEkzV+JmrB2Cc3vSW6PCyfArVG5AfKk5qgPV9kc
# jurGUdHhgRWXzxntWLwP6YpEQQD0bpMS+0SiElIaq2bm5wi82jjG2oti2Tbf55fY
# VgjNlNrkU48xhpY3RZZ7d3OR3zD5B88bpb1rTzSkN/KX3mbZqxaDzb5JLecqkYcY
# ivQi1YaLL4wRF5UjMZJdeCMXm97UIXjKY3BZVPG2FfjTJnh2ZA2tsTKsBC5w8mhZ
# Vwf8e+retF+04ylRV8u+hy6IK/ZMlNcugPohGU+Yw4QL9Arg2IgGVUaAXn/fjH+9
# zowkwSFWyYW9dVPB3nSdk5tmfcdnDopBNdrVLst5/Sm+nxZ199NLG4FwU7z7PRmx
# jFS5ivQFCLvveeJQlkWHq9IEXkSW4SSHsUZEprUTulZsQr0cShl4sXLNQS7SrWvZ
# aX/v4R6pfIGpjK5ZqxXau2jR/5EzSxn0UcGEgA2gBvj3Yd2MbOCGy0yxNg3DVtwT
# VwF2VFlk1Q==
# SIG # End signature block
