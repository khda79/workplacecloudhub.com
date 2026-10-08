<#
.SYNOPSIS
    Offline contract tests for the read-only farm diagnostic.
.VERSION
    1.0.4
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
    $logStart = $script:FarmDiagLog.Count
    Write-FarmDiagLog "First diagnostic action`nSecond diagnostic action"
    $logged = @($script:FarmDiagLog | Select-Object -Skip $logStart)
    Assert-Equal $logged.Count 2 'Multiline console actions have separate timestamps'
    if (@($logged | Where-Object { $_ -notmatch '^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] ' }).Count) { throw 'Console action is missing a timestamp.' }
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
    Assert-Equal $inventory.Sites[0].LogPeriod 'Daily' 'Default IIS log period'

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
    $header = '#Fields: date time s-ip cs-method cs-uri-stem cs-username c-ip sc-status sc-substatus sc-win32-status cs(User-Agent)'
    @($header,'2026-09-24 10:00:00 10.0.0.1 GET /_vti_bin/client.svc old 10.0.0.4 401 1 0 Browser') | Set-Content -LiteralPath (Join-Path $logFolder 'u_ex260924.log')
    @($header,'2026-10-02 22:00:40 10.0.0.1 GET /_vti_bin/client.svc current 10.0.0.5 401 1 0 Browser') | Set-Content -LiteralPath (Join-Path $logFolder 'u_ex261002.log')
    @($header,'2026-10-02 22:00:50 10.0.0.1 GET /_vti_bin/client.svc rollover 10.0.0.6 401 1 0 Browser') | Set-Content -LiteralPath (Join-Path $logFolder 'u_ex261001.log')
    $progressMessages = New-Object 'System.Collections.Generic.List[string]'
    $result = Get-FarmDiagIisRequests 'LOCAL' $inventory.Sites $windows $true $zone { param($message) $progressMessages.Add($message) | Out-Null }
    Assert-Equal $result.Rows.Count 4 'Relevant IIS groups'
    Assert-Equal $result.ExcludedRobot 1 'Robot exclusion'
    Assert-Equal $result.FilesRead 3 'Relevant dated files and unknown names are opened'
    Assert-Equal $result.FilesSkipped 1 'Old daily IIS file is skipped'
    Assert-Equal $result.LinesRead 5008 'Streaming line count'
    if (-not (@($progressMessages | Where-Object { $_ -match 'selected 3 of 4 log files' }).Count)) { throw 'IIS selection progress was not reported.' }
    if (-not (@($progressMessages | Where-Object { $_ -match 'listing .*W3SVC7' }).Count)) { throw 'IIS folder action was not reported.' }
    if (-not (@($progressMessages | Where-Object { $_ -match 'closed u_ex261002.log' }).Count)) { throw 'IIS file completion was not reported.' }
    Assert-Equal @($result.Rows | Where-Object Win32Meaning -EQ 'SEC_E_NO_CREDENTIALS').Count 1 'SSPI interpretation'
    if ($result.DurationSeconds -lt 0) { throw 'Negative IIS read duration.' }
    $inventory.Sites[0].LogPeriod = 'Monthly'
    $monthlyResult = Get-FarmDiagIisRequests 'LOCAL' $inventory.Sites $windows $true $zone
    Assert-Equal $monthlyResult.FilesSkipped 0 'Non-daily IIS rotation does not skip files'
    Assert-Equal $monthlyResult.LinesRead 5010 'Non-daily IIS rotation scans all files'
    $inventory.Sites[0].LogPeriod = 'Daily'
    $missingFolder = Join-Path $logRoot 'W3SVC8'
    [void](New-Item -ItemType Directory -Path $missingFolder -Force)
    @($header,'2026-09-24 10:00:00 10.0.0.1 GET /_vti_bin/client.svc old 10.0.0.4 401 1 0 Browser') | Set-Content -LiteralPath (Join-Path $missingFolder 'u_ex260924.log')
    $missingSite = [pscustomobject]@{ Site='Old'; SiteId='8'; LogDirectory=$logRoot; LogFormat='W3C'; LogPeriod='Daily'; Pool='ContentPool' }
    $missingResult = Get-FarmDiagIisRequests 'LOCAL' @($missingSite) $windows $true $zone
    Assert-Equal $missingResult.FilesRead 0 'No out-of-window IIS file is opened'
    Assert-Equal $missingResult.MissingPaths.Count 1 'No candidate IIS file is reported missing'
    $hourlyFolder = Join-Path $logRoot 'W3SVC9'
    [void](New-Item -ItemType Directory -Path $hourlyFolder -Force)
    @($header,'2026-10-02 22:00:00 10.0.0.1 GET /_vti_bin/client.svc hourly 10.0.0.7 401 1 0 Browser') | Set-Content -LiteralPath (Join-Path $hourlyFolder 'u_ex26100222.log')
    @($header,'2026-09-24 10:00:00 10.0.0.1 GET /_vti_bin/client.svc old 10.0.0.4 401 1 0 Browser') | Set-Content -LiteralPath (Join-Path $hourlyFolder 'u_ex26092410.log')
    $hourlySite = [pscustomobject]@{ Site='Hourly'; SiteId='9'; LogDirectory=$logRoot; LogFormat='W3C'; LogPeriod='Hourly'; Pool='ContentPool' }
    $hourlyResult = Get-FarmDiagIisRequests 'LOCAL' @($hourlySite) $windows $true $zone
    Assert-Equal $hourlyResult.FilesRead 1 'Hourly IIS file in the requested window is opened'
    Assert-Equal $hourlyResult.FilesSkipped 1 'Old hourly IIS file is skipped'

    function Get-WinEvent {
        param([string]$ComputerName,[hashtable]$FilterHashtable,[string]$ErrorAction)
        if ($ComputerName -ne 'WFE02') { throw 'The event query did not target the remote server through RPC.' }
        $id = if ($FilterHashtable.ContainsKey('Id')) { 5210 } else { 5199 }
        $record = if ($id -eq 5210) { 1 } else { 2 }
        $event = [pscustomobject]@{ Id=$id; RecordId=$record; ProviderName='Microsoft-Windows-WAS'; Message='synthetic event'; TimeCreated=[datetime]'2026-10-03T00:00:00' }
        $event | Add-Member -MemberType ScriptMethod -Name ToXml -Value { '<Event><System><TimeCreated SystemTime="2026-10-02T22:00:00.0000000Z"/></System></Event>' }
        return $event
    }
    $eventProgress = New-Object 'System.Collections.Generic.List[string]'
    $rpcEvents = @(Get-FarmDiagEvents 'WFE02' 'System' @(5210) $windows $zone 7 -Progress { param($message) $eventProgress.Add($message) | Out-Null })
    Assert-Equal $rpcEvents.Count 2 'Targeted and other WAS events through RPC'
    Assert-Equal @($rpcEvents | Where-Object EventId -EQ 5199).Count 1 'Other WAS event in window'
    if (-not (@($eventProgress | Where-Object { $_ -match 'RPC query IDs=5210' }).Count)) { throw 'RPC event query was not reported.' }
    if (-not (@($eventProgress | Where-Object { $_ -match 'peak window 1/1 returned 1 WAS records' }).Count)) { throw 'WAS peak window result was not reported.' }
    $eventProgress.Clear()
    $securityEvents = @(Get-FarmDiagEvents 'WFE02' 'Security' @(4625) $windows $zone -Progress { param($message) $eventProgress.Add($message) | Out-Null })
    Assert-Equal $securityEvents.Count 1 'Security RPC query with progress callback'
    if (-not (@($eventProgress | Where-Object { $_ -match 'retained 1 records' }).Count)) { throw 'Security event result was not reported.' }

    function Get-SPLogEvent {
        param([string]$Directory,[datetime]$StartTime,[datetime]$EndTime,[string]$ErrorAction)
        if ($Directory -ne $root) { throw 'ULS did not use the configured directory.' }
        [pscustomobject]@{ Timestamp=[datetime]'2026-10-03T00:00:01'; Level='High'; Category='Authentication'; Process='w3wp'; Message='Application Authentication Pipeline Failure' }
        [pscustomobject]@{ Timestamp=[datetime]'2026-10-03T00:00:02'; Level='Critical'; Category='Config Cache'; Process='w3wp'; Message='synthetic' }
    }
    $ulsProgress = New-Object 'System.Collections.Generic.List[string]'
    $ulsResult = Get-FarmDiagUls 'LOCAL' $root $windows $zone 1 $true { param($message) $ulsProgress.Add($message) | Out-Null }
    Assert-Equal $ulsResult.Rows.Count 1 'ULS event cap'
    Assert-Equal $ulsResult.Truncated $true 'ULS truncation marker'
    if (-not (@($ulsProgress | Where-Object { $_ -match 'checking .*farm-diagnostic-test-' }).Count)) { throw 'ULS directory check was not reported.' }
    if (-not (@($ulsProgress | Where-Object { $_ -match 'window 1/1: querying ' }).Count)) { throw 'ULS window query was not reported.' }
    if (-not (@($ulsProgress | Where-Object { $_ -match 'window 1/1 completed:' }).Count)) { throw 'ULS window completion was not reported.' }

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
    Assert-Equal $correlation[0].Iis401Count 3 'IIS 401 in peak'
    Assert-Equal $correlation[0].StartLocal '2026-10-02 23:30:00' 'Local correlation start'
    $timeline = @(Get-FarmDiagTimeline $events @() $result.Rows @())
    Assert-Equal $timeline.Count 5 'Window timeline rows'
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDZTJyZxBkytpFl
# +92EA3Lef8Cx1wWVZrh40qAK8Bf/5aCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDvtvTEftqinKzVeapY6tXd
# 34VU2i9qxn8nVjQJN0kwZTANBgkqhkiG9w0BAQEFAASCAYBCaOB9liijEWHDDfP0
# xMSt0V9W1/ay8Hf+L0gOgfieHYa6C4dzd/nPu71R+zovvE/N09Mq2geksmAEfyd4
# yZcGQ7Te21P2g2OWC4ST+xcTjbLPG3oA33QVM1RC41WA5ANGi6Gpy5pZ8HZdZ0Ef
# t9S8i5Bxjja3BmcuPWt0ne14tMXcPZNzOQOeuGwlEA6idk5HbAj5peoLPuNTPPI/
# ZBxUJPoqwKuamKRxhXPQAXASh5ul5HNSum38lbj7oFVaP4snTH03dTNVEVC3ek//
# SxW633piFbbOru/bte8P6dt2ACOLV6jImUQwxtwsmHT0SaiD6/4fpb4XDHrY6Jri
# xnjKvCWQF+xZQPv+CQEy4knCNOPhs3rmklcdM1OJEdl6sBWyacoOGEytcQVJGKbL
# GR3DfSHHgcvx23Zm+ciXrwQddS6phDcSlyuRvBIJZGzlT8Mf994/RB59RubyI4tl
# SMEbq8sGV5KkD1z3RT3znNDhE8x0i6AK++uwm4pEVykXJqc=
# SIG # End signature block
