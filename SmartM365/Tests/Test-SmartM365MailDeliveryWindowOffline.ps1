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
