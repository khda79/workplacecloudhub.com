<#
.SYNOPSIS
    Offline tests for the business-hours mail delivery window: schedule, routing,
    deferred non-administrator copy and supersede-by-latest queue delivery.
.DESCRIPTION
    Synthetic configuration, recipients and queue directory. Transport, ACL checks and
    the call stack are mocked; no mail service, share or production file is contacted.
.VERSION
    1.0.0
#>
$ErrorActionPreference = 'Stop'
$core = Join-Path $PSScriptRoot '../Modules/SmartM365.Core'
Import-Module (Join-Path $core 'SmartM365.Orchestrator.ReportMailSchedule.psm1') -Global -Force
Import-Module (Join-Path $core 'SmartM365.Orchestrator.ReportMailQueue.psm1') -Global -Force
. (Join-Path $core 'SmartM365-MailAdminRouting.ps1')
. (Join-Path $core 'SmartM365-MailMaintenance.ps1')
. (Join-Path $core 'SmartM365-DeferredMail.ps1')

$script:checks = 0
function Assert-Window { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw "FAIL: $Message" }; $script:checks++ }

# ---- 1. Window calculation (Romance Standard Time) --------------------------------
$zone = 'Romance Standard Time'
$week = @('Monday','Tuesday','Wednesday','Thursday','Friday')
$cases = @(
    @{ Now='2026-11-18T09:00:00Z'; Open=$true;  Next='2026-11-18T09:00:00Z'; Label='Wednesday 10:00 open' }
    @{ Now='2026-11-18T06:59:00Z'; Open=$false; Next='2026-11-18T07:00:00Z'; Label='Wednesday 07:59 closed until 08:00' }
    @{ Now='2026-11-18T07:00:00Z'; Open=$true;  Next='2026-11-18T07:00:00Z'; Label='08:00 is inside the window' }
    @{ Now='2026-11-18T18:00:00Z'; Open=$false; Next='2026-11-19T07:00:00Z'; Label='19:00 is outside the window' }
    @{ Now='2026-11-20T19:30:00Z'; Open=$false; Next='2026-11-23T07:00:00Z'; Label='Friday evening waits for Monday' }
    @{ Now='2026-11-21T10:00:00Z'; Open=$false; Next='2026-11-23T07:00:00Z'; Label='Saturday waits for Monday' }
    @{ Now='2026-03-29T10:00:00Z'; Open=$false; Next='2026-03-30T06:00:00Z'; Label='Monday after spring DST opens at 06:00 UTC' }
    @{ Now='2026-10-25T10:00:00Z'; Open=$false; Next='2026-10-26T07:00:00Z'; Label='Monday after autumn DST opens at 07:00 UTC' }
)
foreach ($case in $cases) {
    $state = Get-SmartM365ReportMailWindowState -Now ([datetimeoffset]::Parse($case.Now)) -Days $week -NotBefore '08:00' -NotAfter '19:00' -TimeZoneId $zone
    Assert-Window ($state.IsOpen -eq $case.Open) "$($case.Label): IsOpen"
    Assert-Window ($state.NextOpenUtc -eq [datetimeoffset]::Parse($case.Next)) "$($case.Label): next opening $($state.NextOpenUtc.ToString('o'))"
}
$all = Get-SmartM365ReportMailWindowState -Now ([datetimeoffset]::Parse('2026-11-21T22:30:00Z')) -Days @('saturday') -NotBefore '08:00' -NotAfter '24:00' -TimeZoneId $zone
Assert-Window ($all.IsOpen) 'NotAfter 24:00 keeps the window open until midnight'
foreach ($bad in @(@{Days=@('Mon');After='19:00'}, @{Days=@('Monday','monday');After='19:00'}, @{Days=@('1');After='19:00'}, @{Days=$week;After='08:00'}, @{Days=$week;After='07:00'})) {
    $failed = $false
    try { $null = Get-SmartM365ReportMailWindowState -Now ([datetimeoffset]::UtcNow) -Days $bad.Days -NotBefore '08:00' -NotAfter $bad.After -TimeZoneId $zone } catch { $failed = $true }
    Assert-Window $failed "Invalid window $($bad.Days -join ',') / $($bad.After) was accepted"
}

# ---- 2. Recipient routing -------------------------------------------------------------
$script:clock = [datetimeoffset]::Parse('2026-11-18T09:00:00Z')
$script:caller = 'C:\synthetic\SmartInventory\Collector.ps1'
$script:adminOnly = $false
$script:schedule = [pscustomobject]@{ Enabled=$true; NotBefore='08:00'; NotAfter='19:00'; Days=$week; TimeZone=$zone; QueuePath=''; TrustedSids=@('S-1-5-18') }
$script:cfg = [pscustomobject]@{ ProfileKey='demo'; Smart365Administrators='admin@example.invalid'; To='team@example.invalid'; ErrorMailTo='admin@example.invalid'; ReportMailSchedule=$script:schedule }
$script:logs = [Collections.Generic.List[string]]::new()
function Get-SmartM365MailUtcNow { $script:clock }
function Get-SmartM365EffectiveModuleGlobalConfig { $script:cfg }
function Resolve-SmartM365ConfigValue { param($Value) $Value }
function Get-SmartM365MaintenanceMailRoot { param($Config) 'C:\synthetic\Orchestrator' }
function Get-SmartM365AdminMailControl { param($SharedDataFolderPath) [pscustomobject]@{ AdministratorsOnly=$script:adminOnly; Revision=1 } }
function Get-SmartM365MaintenanceMailState { param($SharedDataFolderPath) [pscustomobject]@{ Enabled=$false; Revision=0 } }
function Get-SmartM365CallerLocalConfig { [pscustomobject]@{} }
function Get-ModuleLocalConfig { [pscustomobject]@{} }
function Get-ModuleLocalConfigValue { param($Config,$Name,$DefaultValue) if ($Name -eq 'ErrorMailTo') { 'admin@example.invalid' } else { $DefaultValue } }
function Get-PSCallStack { [pscustomobject]@{ ScriptName=$script:caller } }
function WriteLog { param($Message,$Level) $script:logs.Add("$Level $Message") }
function Split-Recipients { param([string]$Value) @($Value -split ';' | Where-Object { $_ } | Sort-Object) -join ';' }

$route = Resolve-SmartM365UnifiedMailRoute -To 'team@example.invalid' -Cc 'copy@example.invalid' -MailPurpose Report
Assert-Window ((Split-Recipients ($route.To + ';' + $route.Cc)) -eq 'admin@example.invalid;copy@example.invalid;team@example.invalid') 'Inside the window normal recipients and administrators receive the mail'
$script:clock = [datetimeoffset]::Parse('2026-11-21T10:00:00Z')
foreach ($purpose in @('Report','Error','Maintenance')) {
    $route = Resolve-SmartM365UnifiedMailRoute -To 'team@example.invalid' -Cc 'copy@example.invalid' -Bcc 'hidden@example.invalid' -MailPurpose $purpose
    Assert-Window ($route.To -eq 'admin@example.invalid' -and -not $route.Cc -and -not $route.Bcc) "Outside the window $purpose mail goes to administrators only"
}
$script:caller = 'C:\synthetic\Exchange\ExchangeUserNotifications\Campaign.ps1'
$route = Resolve-SmartM365UnifiedMailRoute -To 'user@example.invalid' -MailPurpose Report
Assert-Window ((Split-Recipients $route.To) -eq 'admin@example.invalid;user@example.invalid') 'User communication campaigns ignore the window'
$script:caller = 'C:\synthetic\SmartInventory\Collector.ps1'
$script:schedule.NotAfter = '07:00'
$route = Resolve-SmartM365UnifiedMailRoute -To 'team@example.invalid' -MailPurpose Report
Assert-Window ($route.To -eq 'admin@example.invalid') 'An invalid window restricts mail to administrators'
Assert-Window (@($script:logs | Where-Object { $_ -like 'WARNING Mail delivery window is invalid*' }).Count -ge 1) 'An invalid window is logged'
$script:schedule.NotAfter = '19:00'
$legacy = [pscustomobject]@{ Enabled=$true; NotBefore='08:00'; TimeZone=$zone; QueuePath=''; TrustedSids=@('S-1-5-18') }
$script:cfg.ReportMailSchedule = $legacy
$route = Resolve-SmartM365UnifiedMailRoute -To 'team@example.invalid' -MailPurpose Report
Assert-Window ((Split-Recipients $route.To) -eq 'admin@example.invalid;team@example.invalid') 'Without Days and NotAfter the legacy cutoff does not restrict recipients'
$script:cfg.ReportMailSchedule = $script:schedule

$script:clock = [datetimeoffset]::Parse('2026-11-23T08:00:00Z')
$global:SmartM365MailExcludeAdministrators = $true
try {
    $route = Resolve-SmartM365UnifiedMailRoute -To 'admin@example.invalid' -Cc 'team@example.invalid;ADMIN@example.invalid' -MailPurpose Report
    Assert-Window ($route.To -eq 'team@example.invalid' -and -not $route.Cc) 'Queued copy excludes administrators and promotes Cc'
    $route = Resolve-SmartM365UnifiedMailRoute -To 'admin@example.invalid' -MailPurpose Report
    Assert-Window ([string]::IsNullOrEmpty($route.To)) 'Queued copy for administrators only has no recipient left'
    $script:adminOnly = $true
    $route = Resolve-SmartM365UnifiedMailRoute -To 'team@example.invalid' -MailPurpose Report
    Assert-Window ([string]::IsNullOrEmpty($route.To)) 'Administrators-only routing drops the queued copy'
    $script:adminOnly = $false
} finally { Remove-Variable -Name SmartM365MailExcludeAdministrators -Scope Global -ErrorAction SilentlyContinue }

# ---- 3. Deferred copy, supersede and dispatch -----------------------------------------
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-MailWindow-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $root
$script:schedule.QueuePath = $root
$script:delivered = [Collections.Generic.List[object]]::new()
function Test-SmartM365DeferredMailDirectory { param($Path,$TrustedSids) $Path -eq $root }
function Initialize-SmartM365DeferredMailRuntime { }
function Set-Acl { param($LiteralPath,$AclObject) }
function SendEmailHtmlReport {
    param($From,$To,$Cc,$Subject,$BodyHtml,$Attachments,$AllowAttachments,$MailPurpose,$DeliveryClass,$SendMailMode,$SmtpServer,$SmtpPort)
    $final = Resolve-SmartM365UnifiedMailRoute -To $To -Cc $Cc -MailPurpose Report
    $script:delivered.Add([pscustomobject]@{ Subject=$Subject; To=$final.To; Cc=$final.Cc })
}
try {
    $message = { param($subject,$to) @{ From='sender@example.invalid'; To=$to; Cc='admin@example.invalid'; Subject=$subject; BodyHtml='<p>synthetic</p>'; SendMailMode='Graph'; SmtpServer=''; SmtpPort=25 } }
    $script:clock = [datetimeoffset]::Parse('2026-11-18T09:00:00Z')
    Assert-Window (-not (Add-SmartM365DeferredReport -Message (& $message 'Daily 2026-11-18' 'team@example.invalid') -Attachments @() -MailPurpose Report -DeliveryClass Report)) 'Inside the window the report is sent directly'
    Assert-Window (@(Get-ChildItem $root -Filter '*.json').Count -eq 0) 'Inside the window nothing is queued'
    foreach ($day in @('2026-11-21T02:00:00Z','2026-11-22T02:00:00Z','2026-11-23T02:00:00Z')) {
        $script:clock = [datetimeoffset]::Parse($day)
        $global:SmartM365LastMailDelivery = ''
        $continue = Add-SmartM365DeferredReport -Message (& $message ('Daily ' + $day.Substring(0,10)) 'team@example.invalid') -Attachments @() -MailPurpose Report -DeliveryClass Report
        Assert-Window (-not $continue) "Outside the window the administrator copy is still sent immediately ($day)"
        Assert-Window ($global:SmartM365LastMailDelivery -eq 'Queued') "Outside the window the submission is recorded as Queued ($day)"
        Start-Sleep -Milliseconds 20
    }
    $script:clock = [datetimeoffset]::Parse('2026-11-23T02:00:00Z')
    $null = Add-SmartM365DeferredReport -Message (& $message 'Weekly licences' 'domain@example.invalid') -Attachments @() -MailPurpose Report -DeliveryClass Report
    $null = Add-SmartM365DeferredReport -Message (& $message 'Daily error' 'team@example.invalid') -Attachments @() -MailPurpose Error -DeliveryClass Report
    $null = Add-SmartM365DeferredReport -Message (& $message 'Operational' 'team@example.invalid') -Attachments @() -MailPurpose Report -DeliveryClass Operational
    $records = @(Get-ChildItem $root -Filter '*.json' | ForEach-Object { Get-Content -Raw $_.FullName | ConvertFrom-Json })
    Assert-Window ($records.Count -eq 4) 'Only reports are queued (errors and operational notices are not)'
    Assert-Window (@($records | Where-Object { $_.NotBeforeUnixMs -eq [datetimeoffset]::Parse('2026-11-23T07:00:00Z').ToUnixTimeMilliseconds() }).Count -eq 4) 'Queued copies wait for Monday 08:00'
    Assert-Window (@($records.SupersedeKey | Sort-Object -Unique).Count -eq 2) 'Daily reports share one supersede key; the weekly report has its own'

    $script:clock = [datetimeoffset]::Parse('2026-11-23T06:59:00Z')
    $null = Invoke-SmartM365DeferredMail
    Assert-Window ($script:delivered.Count -eq 0) 'The queue is not dispatched before the window opens'
    $script:clock = [datetimeoffset]::Parse('2026-11-23T07:00:00Z')
    $result = Invoke-SmartM365DeferredMail
    Assert-Window ($result.Sent -eq 2 -and $result.Superseded -eq 2) "Monday sends the latest daily report and the weekly report ($($result | ConvertTo-Json -Compress))"
    $daily = @($script:delivered | Where-Object { $_.Subject -like 'Daily*' })
    Assert-Window ($daily.Count -eq 1 -and $daily[0].Subject -eq 'Daily 2026-11-23' -and $daily[0].To -eq 'team@example.invalid') 'The latest daily report is the one delivered'
    Assert-Window (@($script:delivered | Where-Object { "$($_.To);$($_.Cc)" -match 'admin@' }).Count -eq 0) 'Delivered queued copies exclude administrators'
    Assert-Window (-not (Get-Variable SmartM365MailExcludeAdministrators -Scope Global -ErrorAction SilentlyContinue)) 'The exclusion flag is cleared after dispatch'
    $superseded = @(Get-ChildItem $root -Filter '*.json' | ForEach-Object { Get-Content -Raw $_.FullName | ConvertFrom-Json } | Where-Object { $_.PSObject.Properties['Superseded'] })
    Assert-Window ($superseded.Count -eq 2 -and @($superseded | Where-Object { $_.Status -ne 'Sent' -or $_.PayloadBase64 }).Count -eq 0) 'Superseded copies are closed as Sent without payload'
    $again = Invoke-SmartM365DeferredMail
    Assert-Window ($again.Sent -eq 0 -and $script:delivered.Count -eq 2) 'A second pass sends nothing more'

    # The queued copy of a report addressed only to administrators is closed without sending.
    $script:clock = [datetimeoffset]::Parse('2026-11-24T02:00:00Z')
    $null = Add-SmartM365DeferredReport -Message (& $message 'Admin only' 'admin@example.invalid') -Attachments @() -MailPurpose Report -DeliveryClass Report
    $script:clock = [datetimeoffset]::Parse('2026-11-24T07:00:00Z')
    $result = Invoke-SmartM365DeferredMail
    Assert-Window ($result.Sent -eq 1 -and $script:delivered.Count -eq 2) 'A queued copy with only administrators is closed without a second delivery'
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host ("PASS: {0} mail delivery window checks; PS {1}; no mail sent." -f $script:checks, $PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCAQMs5gXvr6UZF
# Y0LPnRER7OsorXrBbaTA3SHRsXRUzqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAQ0kLU2txEmaaPjjqsBbprXYN5xqmCaNQRyZygGeznBMA0GCSqG
# SIb3DQEBAQUABIIBgB7SZrYBXB7r1fJWI/dSKCw1HF1HBLSaZha4EB5+C7Jj0pK/
# M4liRn0aLzHslIcWPaDbxqkL9515wZ2y/5B22nZ8uaZ8er8hKaLZ5VB60qN/cRn7
# idrg5EAS5A6IB6fXqPEzKp+OX5AWU0NK6GpwqTLCIJWyZFAp7WT4ZVD/fafsDSlK
# 5yJxAFDJZ389I1Ap3M1oei1qDg//z1OJD3pMwxalLJY8/SwH3LcR2xKOXyIJLN+h
# 5X09faqE5tGxGkXM347qc6jQ36E0GdqkaiKcm6CCFmSEvE8e410ZyC/7tNbls+T6
# +G9JFezAlcqKrBwRitx/i/33r9KoODd/5uwziUDHIvevucn+JnEEJ5G9LCcBIlA8
# 1l8cUDXJ4ORbuqhXtBhC58I0nCHohtQ9X4pMAvQ1kyswVUSKcApsfAKFc6aio2vo
# hp0QRDAZxULKsP7/MjB8aPCPluf+WZkvzIfIcmet5nbaF2eI3kqwzhaSCodo0o6q
# +lbQN0vWef0ocRipg6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTgxMzAx
# NDlaMC8GCSqGSIb3DQEJBDEiBCCILTAfsVEqOWjjdDEOf7x4QlWZbGzRBcvDhNzX
# +DoekTANBgkqhkiG9w0BAQEFAASCAgCsLjujWRZ/aHEvU7a4dEUk1jw1R5QmxYn0
# YgAFoEgy7v8J99veULNEwWVOtthFquZfwmQXRHqAuoc1bsLXbkMjizQrwG1llKQg
# Uu1deq2N1BGgozv+ubh5gSxsS9DANurLKpU/JooPujLs9DVIHgoMITITYkil3XPV
# WQXstANTRs23BWvn0B9oo/Nueka8uwbp7trtH61yQ2jXhPrahs0KWoD0cyBX9MNb
# KyGFMUObdumYuhB8RtL3J3L6S8aCuCr6AbcSdbCCgzfCG2qcjPc7t/rlpEHD+O5b
# WqlLh8rG3mqG0RA11/MCaiH2I61X8N83qw6+y71qHxu1/MCN6JnkbJHXz5NpgQj9
# r2VVkFIeFKoRsUTvFmRDTGneOUDafq9kflDw6Lrgz2EPlw41JLXLF2ulqaC6UwMR
# XxQSgAEUteaw1S+VBtEJzxX8OUi/wNdvS3p1GriV18DYs8X6gT/P2F416hs3zcSs
# II1xJP5Yq+vAnZGvDbIBXqrnh7Pll/3NApduz3IcgVB1HTmuDvuHhJ0v4aBWWOEY
# 0+IqRAkPz+ywqe8LFkkT0XhpWVJzB/EOrS0+ntaddBqARoa0DvypqRCfKzCA+MKz
# iMNX6fCL4Z1x2O+4TVbjV+YQE06yvp0NABMxxZ17EYoBhEibGSbTZKkkk83DR3GM
# d08HGEL+qA==
# SIG # End signature block
