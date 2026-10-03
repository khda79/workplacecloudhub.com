<#
.SYNOPSIS
    Offline contract tests for the read-only farm diagnostic.
.VERSION
    1.0.2
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-FarmDiagnostic.ps1'
. $scriptPath -Project 'Synthetic' -Around ([datetime]'2026-10-03T00:00:00')
$root = Join-Path $PSScriptRoot ('.farm-diagnostic-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test path.' }

function Assert-Equal {
    param($Actual,$Expected,[string]$Label)
    if ([string]$Actual -cne [string]$Expected) { throw "$Label`: expected '$Expected', got '$Actual'." }
}

try {
    [void](New-Item -ItemType Directory -Path $root -Force)
    $zone = [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time')
    $center = [datetime]::SpecifyKind([datetime]'2026-10-03T00:00:00',[DateTimeKind]::Unspecified)
    $windows = @(Get-FarmDiagWindows -Mode Around -Center $center -Minutes 30 -Zone $zone)
    Assert-Equal $windows.Count 1 'Around window count'
    Assert-Equal ($windows[0].PeakUtc.ToString('o')) '2026-10-02T22:00:00.0000000Z' 'Local to UTC conversion'
    $rangeWindows = @(Get-FarmDiagWindows -Mode Range -Start $center.AddMinutes(-5) -End $center.AddMinutes(5) -Zone $zone)
    Assert-Equal $rangeWindows.Count 1 'Range window count'
    $peakPath = Join-Path $root 'AccessFailures-5min.csv'
    @'
WindowUtc,Lines,Items,Source,Destination,Undetermined
2026-10-02 22:00 UTC,9,4,9,0,0
2026-10-02 22:05 UTC,4,2,4,0,0
'@ | Out-File -LiteralPath $peakPath -Encoding utf8
    $fromPeaks = @(Get-FarmDiagWindows -Mode Peaks -Minutes 10 -PeaksPath $peakPath -Zone $zone)
    Assert-Equal $fromPeaks.Count 2 'CSV peak window count'
    Assert-Equal $fromPeaks[0].ShareGateLines 9 'CSV line count'
    $fromPeaksWithEmptyDates = @(Get-FarmDiagWindows -Mode Peaks -Start $null -End $null -Center $null -Minutes 10 -PeaksPath $peakPath -Zone $zone)
    Assert-Equal $fromPeaksWithEmptyDates.Count 2 'CSV-only window selection with unbound dates'
    . $scriptPath -Project 'Synthetic' -ShareGatePeaksCsv $peakPath -WindowMinutes 30 -DryRun
    Assert-Equal $script:FarmDiagMode 'Peaks' 'CSV-only parameter set without Around prompt'
    $eligibleServers = @(Get-FarmDiagEligibleServers @(
        [pscustomobject]@{ Address='WEB01'; Role='WebFrontEnd'; Type='Application' },
        [pscustomobject]@{ Address='APP01'; Role='ApplicationWithSearch'; Type='Application' },
        [pscustomobject]@{ Address='SQL01'; Role='Invalid'; Type='Application' },
        [pscustomobject]@{ Address='SMTP01'; Role='External'; Type='Application' },
        [pscustomobject]@{ Address='DB01'; Role='Custom'; Type='Database' }
    ))
    Assert-Equal $eligibleServers.Count 2 'Only SharePoint servers are eligible for collection'
    Assert-Equal ((@($eligibleServers | ForEach-Object { [string]$_.Address })) -join ',') 'WEB01,APP01' 'External farm entries are excluded'
    Assert-Equal (ConvertTo-FarmDiagUncPath 'WFE02' 'D:\IIS\Logs') '\\WFE02\d$\IIS\Logs' 'Remote drive mapping'
    $rerun = Get-FarmDiagRerunCommand 'WFE02' $zone.Id $true
    if ($rerun -notmatch '-Project "Synthetic"' -or $rerun -notmatch '-Servers "WFE02"' -or $rerun -notmatch '-DryRun$') { throw 'Local rerun command is incomplete.' }
    try { ConvertTo-FarmDiagUtc ([datetime]::SpecifyKind([datetime]'2026-10-25T02:30:00',[DateTimeKind]::Unspecified)) $zone | Out-Null; throw 'Ambiguous local time was accepted.' }
    catch { if ($_.Exception.Message -notmatch 'ambiguous') { throw } }

    [xml]$config = @'
<configuration><system.applicationHost><applicationPools><add name="ContentPool"><cpu limit="15000" action="Throttle" resetInterval="00:05:00"/><recycling><periodicRestart time="29:00:00" privateMemory="1024"><schedule><add value="00:00:00"/></schedule></periodicRestart></recycling><processModel idleTimeout="00:20:00"/></add></applicationPools><sites><siteDefaults><logFile directory="D:\IIS\Logs"/></siteDefaults><site name="Content" id="7"><bindings><binding protocol="https" bindingInformation="*:443:source.example"/></bindings><application path="/" applicationPool="ContentPool"/><logFile directory="D:\IIS\Logs"/></site></sites></system.applicationHost></configuration>
'@
    $inventory = Get-FarmDiagIisInventory $config 'WFE02'
    Assert-Equal $inventory.Pools.Count 1 'IIS pool count'
    Assert-Equal $inventory.Pools[0].CpuLimitPercent 15 'CPU percent conversion'
    Assert-Equal $inventory.Pools[0].CpuAction 'Throttle' 'CPU action'
    Assert-Equal $inventory.Pools[0].RecycleSchedule '00:00:00' 'Recycle schedule'
    Assert-Equal $inventory.Sites[0].Pool 'ContentPool' 'Site pool mapping'

    $logRoot = Join-Path $root 'IIS'
    $logFolder = Join-Path $logRoot 'W3SVC7'
    [void](New-Item -ItemType Directory -Path $logFolder -Force)
    $inventory.Sites[0].LogDirectory = $logRoot
    $logPath = Join-Path $logFolder 'u_ex.log'
    $writer = New-Object IO.StreamWriter($logPath,$false,[Text.UTF8Encoding]::new($false))
    try {
        $writer.WriteLine('#Fields: date time s-ip cs-method cs-uri-stem cs-username c-ip sc-status sc-substatus sc-win32-status cs(User-Agent)')
        for($i=0;$i -lt 5000;$i++){ $writer.WriteLine('2026-10-01 10:00:00 10.0.0.1 GET /_vti_bin/client.svc - 10.0.0.2 401 1 2148074252 Browser') }
        $writer.WriteLine('2026-10-02 22:00:10 10.0.0.1 GET /_vti_bin/client.svc user 10.0.0.2 401 1 2148074254 Browser')
        $writer.WriteLine('2026-10-02 22:00:20 10.0.0.1 GET /_vti_bin/client.svc user 10.0.0.2 200 0 0 Browser')
        $writer.WriteLine('2026-10-02 22:00:30 10.0.0.1 GET /_vti_bin/client.svc robot 10.0.0.3 401 1 0 MS+Search+6.0+Robot')
    }
    finally { $writer.Dispose() }
    $result = Get-FarmDiagIisRequests 'LOCAL' $inventory.Sites $windows $true $zone
    Assert-Equal $result.Rows.Count 2 'Relevant IIS groups'
    Assert-Equal $result.ExcludedRobot 1 'Robot exclusion'
    Assert-Equal $result.LinesRead 5004 'Streaming line count'
    Assert-Equal @($result.Rows | Where-Object Win32Meaning -EQ 'SEC_E_NO_CREDENTIALS').Count 1 'SSPI interpretation'
    if ($result.DurationSeconds -lt 0) { throw 'Negative IIS read duration.' }

    function Get-WinEvent {
        param([string]$ComputerName,[hashtable]$FilterHashtable,[string]$ErrorAction)
        if ($ComputerName -ne 'WFE02') { throw 'The event query did not target the remote server through RPC.' }
        $id = if ($FilterHashtable.ContainsKey('Id')) { 5210 } else { 5199 }
        $record = if ($id -eq 5210) { 1 } else { 2 }
        $event = [pscustomobject]@{ Id=$id; RecordId=$record; ProviderName='Microsoft-Windows-WAS'; Message='synthetic event'; TimeCreated=[datetime]'2026-10-03T00:00:00' }
        $event | Add-Member -MemberType ScriptMethod -Name ToXml -Value { '<Event><System><TimeCreated SystemTime="2026-10-02T22:00:00.0000000Z"/></System></Event>' }
        return $event
    }
    $rpcEvents = @(Get-FarmDiagEvents 'WFE02' 'System' @(5210) $windows $zone 7)
    Assert-Equal $rpcEvents.Count 2 'Targeted and other WAS events through RPC'
    Assert-Equal @($rpcEvents | Where-Object EventId -EQ 5199).Count 1 'Other WAS event in window'

    function Get-SPLogEvent {
        param([string]$Directory,[datetime]$StartTime,[datetime]$EndTime,[string]$ErrorAction)
        if ($Directory -ne $root) { throw 'ULS did not use the configured directory.' }
        [pscustomobject]@{ Timestamp=[datetime]'2026-10-03T00:00:01'; Level='High'; Category='Authentication'; Process='w3wp'; Message='Application Authentication Pipeline Failure' }
        [pscustomobject]@{ Timestamp=[datetime]'2026-10-03T00:00:02'; Level='Critical'; Category='Config Cache'; Process='w3wp'; Message='synthetic' }
    }
    $ulsResult = Get-FarmDiagUls 'LOCAL' $root $windows $zone 1 $true
    Assert-Equal $ulsResult.Rows.Count 1 'ULS event cap'
    Assert-Equal $ulsResult.Truncated $true 'ULS truncation marker'

    $events = @(
        [pscustomobject]@{ Server='WFE02'; EventId=5210; Utc='2026-09-30T22:00:00Z'; Local='2026-10-01 00:00:00'; InPeakWindow=$false },
        [pscustomobject]@{ Server='WFE02'; EventId=5210; Utc='2026-10-01T22:00:00Z'; Local='2026-10-02 00:00:00'; InPeakWindow=$false },
        [pscustomobject]@{ Server='WFE02'; EventId=5210; Utc='2026-10-02T22:00:00Z'; Local='2026-10-03 00:00:00'; InPeakWindow=$true }
    )
    $recurrences = @(Get-FarmDiagNightlyRecurrences $events 7 $windows $zone)
    Assert-Equal $recurrences.Count 1 'Night recurrence count'
    Assert-Equal $recurrences[0].Nights 3 'Distinct nights'
    $correlation = @(Get-FarmDiagCorrelation $windows $events @() $result.Rows @() $zone)
    Assert-Equal $correlation[0].WasEvents 1 'WAS in peak'
    Assert-Equal $correlation[0].Iis401Count 1 'IIS 401 in peak'
    Assert-Equal $correlation[0].StartLocal '2026-10-02 23:30:00' 'Local correlation start'
    $timeline = @(Get-FarmDiagTimeline $events @() $result.Rows @())
    Assert-Equal $timeline.Count 3 'Window timeline rows'
    $risks = @(Get-FarmDiagPoolRisks $inventory.Pools $windows $zone)
    Assert-Equal @($risks | Where-Object Risk -EQ 'CpuThrottle').Count 1 'CPU throttle risk'
    Assert-Equal @($risks | Where-Object Risk -EQ 'ScheduledRecycleNearWindow').Count 1 'Recycling window risk'
    $csvPath = Write-FarmDiagCsv $root 'Farm-Test.csv' @('Server','EventId') $events
    Assert-Equal @(Import-Csv -LiteralPath $csvPath).Count 3 'Atomic CSV result'
    $summary = [pscustomobject]@{ Project='Synthetic'; Status='Complete'; GeneratedAtUtc='2026-10-03T00:00:00Z'; PeakCount=1; Coverage=@([pscustomobject]@{Server='WFE02';Source='IIS logs';Status='Complete';Detail='fixture';DurationSeconds=1.2}) }
    $htmlPath = Join-Path $root 'Farm-Report.html'
    Write-FarmDiagHtml $htmlPath $summary $inventory.Pools $events $result.Rows $correlation $recurrences $timeline $risks
    if (-not ((Get-Content -LiteralPath $htmlPath -Raw) -match 'ScheduledRecycleNearWindow')) { throw 'HTML report omitted the window risk.' }
    Assert-Equal @(Get-ChildItem -LiteralPath $root -File -Filter '*.tmp-*').Count 0 'No temporary output'
    Write-Output 'Farm diagnostic offline contract tests passed.'
}
finally {
    if ((Test-Path -LiteralPath $root) -and $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCvOV073VX4X7Kr
# apOzv/nfdUlU307jHmDbzGIxW7MedaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIN4/TBjiPQR+hWZZl+SqgkeBzAkQ5GsBm/uz3gv4liHvMA0GCSqG
# SIb3DQEBAQUABIIBgHshhgW3fvk0sAkWEmwTat2aLw4ZvNVnlEAH6hYcaqXjW7p1
# /m5mZ5/b5hguQ1XcMnfDY5+iK4TD5qjhJbCOWifZPBzmQ9ceO8LPFNHEmoEXIwgn
# uCUSKZUK9uKlEwFvFo7z8ClIyidwYIMsl3kvncjRHGd+SZJDFQY+ac7flEbj+DC1
# 5XXLR9QYRmipZMyqmGG+XDpG8Co7bcktFLYiM7sjJg4wcF2O5l6ah0OlI22oaBux
# P0A+VW4P/WQocaboBEceQToKl0xduMWTpjgaqCEq/tvk6mlJ1xIFJlTesM9pnDRN
# W/k/xv3Oo3m2b3bWPwwNSKIFYl8I5PYz8Optt7ZFoL/Jg80j2Sv3Oz5sWvVWubDV
# s/c2g1m2TRo7C/lLG4DwjPAlBHISLeobgq6y6ikozgEBxcoRXFRjq8VhRvfCLRd9
# lD+MbyMLECEPbL8Xw+JIomikHi/sIryRk/pYqDlb3yD1qlSDc5P/06xJd1MEylTN
# eoq4lkXKKmtQk2uZKaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxMzMx
# NTVaMC8GCSqGSIb3DQEJBDEiBCBKmOSYQ/Qv0fbIHyPT2uNvcI4FYBYsJFrCMMDl
# KfP37zANBgkqhkiG9w0BAQEFAASCAgAUquKu1xCbc/fI7RigDJoJqT/JrNeM3y7j
# j+LqNsd2OvQJosdsWs7EEIx06REvsUTBZYP9Ha7fzIn0HRQb4QbhrReZbWwsDQyJ
# 9qp2dlZnLbdsUTDusaeniQY98UY6AvYde4uRx9UXQMCgRmP32GYQoUHU11H8guaD
# Oc4hApvodG5+qVxMacOVuglPnCYQHavKBRTOH88kwKdT83ZAN1RsLX4GnCBsiAFa
# d4TY9zqNq3NzuNI6+uYe7q+gcsOqqteYLc+EpQxAgCWA2lzeiXvJTJl34JfsdU9S
# b2qG9OzMepu6ruP5LVDAw8STo+EkueT9/xvzFBkqOTveo5Rf3BNNMHe7iALD+hhZ
# 6aCVJj7MAIL/VNurDuRo/3joo6WPp5cY2xKtdzeVsXMwxWYXDYaRNLVfu9RkITt5
# UuE2r6WnNAbOuDWLwQjuvi9SM9MBXsR+VbJtQurJDfx/5mOlLC0/d1fNPmsHd6jM
# uYsSh4KRvLC4fe4M90UujDCN+Sg6oJd3c/u3fR2lHBSuEsYvviutI5OiObHHOnhx
# NQsetT17ioG3ZXFxM37c/wY0IQGbYOxrGgObl2bp5vJ4uQn8afa/HQBaqraC7FP3
# cg7qI4AhUcrEomh4nCeTTMX/R4HIt8rIMLJ/Bb91RVC/APXtMbc0jaof8OQdPo3B
# asUR4ZExWQ==
# SIG # End signature block
