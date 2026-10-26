#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic data and simulated commands for the existing Orchestrator WPF window.
.VERSION
1.3.0
.NOTES
No tenant context, management modules, shared folders or resident workers.
#>

# Pure Configuration view helpers (no module import, no file access) shared with the GUI module.
. (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.ConfigView.Functions.ps1')

function Get-OrchestratorDemoNextRun {
    param($Job,[datetime]$Clock)
    if(-not $Job.Enabled){return 'Disabled'}
    if($Job.Assignment -eq 'Manual'){return 'Manual'}
    for($offset=0;$offset -le 7;$offset++){
        $day=$Clock.Date.AddDays($offset)
        if($Job.Frequency -eq 'Weekly' -and [string]$day.DayOfWeek -notin $Job.ScheduleDays){continue}
        foreach($text in @($Job.Times -split ',' | ForEach-Object {$_.Trim()} | Sort-Object)){
            $time=[timespan]::Zero
            if(-not [timespan]::TryParseExact($text,'hh\:mm',[cultureinfo]::InvariantCulture,[ref]$time)){throw 'The demo catalog contains an invalid schedule time.'}
            $candidate=$day.Add($time)
            if($candidate -ge $Clock){return $candidate.ToString('yyyy-MM-dd HH:mm')}
        }
    }
    return 'No occurrence'
}

function New-OrchestratorDemoData {
    param([ValidateSet('Overview','Issues','Maintenance')][string]$Scenario)
    $clock = [datetime]'2026-10-18T09:30:00'
    $servers = @(
        foreach ($i in 1..3) {
            [pscustomobject]@{
                Server=('CTSO-INV-{0:00}' -f $i); Online=$true; HeartbeatAgeMinutes=0.2
                Weight=1; Policy=''; Capabilities='AD, Graph, EXO, SharePoint'; AssignedJobs=4
                State='Running'; LastHeartbeat='2026-10-18 09:30'; Running=0; Pending=0
                RecycleIn='02:30:00'; Pid=(4100+$i); StatePersistence='Healthy (sample)'
                Version='1.5.47'; Maintenance='Active'; MaintenanceAcknowledgement='Ready'
            }
        }
    )
    # Read only the neutral shipped template, never Orchestrator-Jobs runtime JSON.
    $catalogPath=Join-Path $PSScriptRoot 'Orchestrator-Jobs.json.template'
    if(-not (Test-Path -LiteralPath $catalogPath -PathType Leaf)){throw 'The demo requires the shipped Orchestrator-Jobs.json.template. Copy it with the GUI.'}
    $catalog=Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
    if(-not $catalog.Jobs -or @($catalog.Jobs.Name | Sort-Object -Unique).Count -ne $catalog.Jobs.Count){throw 'The demo job catalog is empty or contains duplicate names.'}
    $jobs = @(
        for ($i=0; $i -lt $catalog.Jobs.Count; $i++) {
            $source=$catalog.Jobs[$i]
            $scheduleDays=@(Get-OrchestratorGuiPropertyValue $source.Schedule 'DaysOfWeek' @())
            $owner = $servers[$i % 3].Server
            $row=[pscustomobject]@{
                Name=[string]$source.Name; Arguments=[string]$source.Arguments
                Enabled=[bool]$source.Enabled; Frequency=[string]$source.Schedule.Type; Times=(@($source.Schedule.Times) -join ', ')
                Days=$(if($source.Schedule.Type -eq 'Weekly'){$scheduleDays -join ', '}else{'Every day'})
                ScheduleDays=$scheduleDays
                MissedRunPolicy=[string]$source.Schedule.MissedRunPolicy
                Assignment=[string]$source.AssignmentMode; Server=$owner; NextRun=''
                Group=[string]$source.Group; Health=$(if($source.Enabled){'Healthy'}else{'Disabled'})
                LastStatus='Success'; LastSuccess='2026-10-18 08:20'
                SuccessAgeHours=1.2; ExpectedMaxAgeHours=$(if($source.Schedule.Type -eq 'Weekly'){168}else{24})
                AverageDurationMinutes=[double]$source.EstimatedDurationMinutes
                TimeoutMinutes=[int]$source.TimeoutMinutes; MaxRetries=[int]$source.MaxRetries; RetryDelaySeconds=[int]$source.RetryDelaySeconds
                DependsOn=(@($source.DependsOn) -join ', ')
                DependencyMode=[string](Get-OrchestratorGuiPropertyValue $source 'DependencyMode' 'LatestOccurrence')
                DependencyMaxAgeHours=[int](Get-OrchestratorGuiPropertyValue $source 'DependencyMaxAgeHours' 0)
            }
            $row.NextRun=Get-OrchestratorDemoNextRun -Job $row -Clock $clock
            $row
        }
    )
    $history = @(
        for ($day=0; $day -lt 7; $day++) {
            for ($i=0; $i -lt $jobs.Count; $i++) {
                $start = $clock.Date.AddDays(-$day).AddHours(1).AddMinutes([math]::Floor(360*$i/[double]$jobs.Count))
                $duration=3+($i%18)
                $status = if (($day+$i) % 19 -eq 0) {'Failed'} elseif (($day+$i) % 11 -eq 0) {'CompletedWithWarnings'} else {'Success'}
                [pscustomobject]@{
                    EventTime=$start.AddMinutes($duration).ToString('yyyy-MM-dd HH:mm:ss')
                    StartTime=$start.ToString('yyyy-MM-dd HH:mm:ss'); Server=$jobs[$i].Server
                    JobName=$jobs[$i].Name; Status=$status; DurationSec=($duration*60)
                    ExitCode=$(if ($status -eq 'Failed') {1} elseif ($status -eq 'CompletedWithWarnings') {3} else {0})
                    RetryCount=0; LogPath=('Demo/Logs/{0}.log' -f $jobs[$i].Name)
                }
            }
        }
    )
    $deviceJob=@($jobs | Where-Object Name -EQ 'Intune-Devices-Inventory')[0]
    $teamsJob=@($jobs | Where-Object Name -EQ 'M365-Teams-Inventory')[0]
    $spoJob=@($jobs | Where-Object Name -EQ 'M365-SPO-Inventory')[0]
    $activeUsersJob=@($jobs | Where-Object Name -EQ 'M365-ActiveUsers-Inventory')[0]
    $running = @(
        [pscustomobject]@{Server=$deviceJob.Server;ServerState='Running';Job=$deviceJob.Name;Pid=5201;Started='2026-10-18 09:22';DurationMinutes=8}
        [pscustomobject]@{Server=$teamsJob.Server;ServerState='Running';Job=$teamsJob.Name;Pid=5202;Started='2026-10-18 09:26';DurationMinutes=4}
    )
    $servers[0].Running=1; $servers[1].Running=1; $servers[2].Pending=1
    $pending = @([pscustomobject]@{Server=$spoJob.Server;ServerState='Running';Job=$spoJob.Name;Reason='Dependency';WaitingMinutes=4;Details='Waiting for the current Teams inventory (sample).'})
    $incidents = @()
    $logEvents = [Collections.Generic.List[object]]::new()
    foreach ($server in $servers) {
        $logEvents.Add([pscustomobject]@{Time='2026-10-18 06:00:04';Server=$server.Server;Level='INFO';Category='Lifecycle';Count=1;First='2026-10-18 06:00:04';Message=('SmartM365-Inventory-Orchestrator v1.5.47 starting (PID {0}, tenant demo, sample).' -f $server.Pid);Source='Demo/Logs/Orchestrator.log'})
        $logEvents.Add([pscustomobject]@{Time='2026-10-18 05:59:58';Server=$server.Server;Level='INFO';Category='Lifecycle';Count=1;First='2026-10-18 05:59:58';Message='SmartM365-Inventory-Orchestrator exiting normally (exit code 0).';Source='Demo/Logs/Orchestrator.log'})
        $logEvents.Add([pscustomobject]@{Time='2026-10-18 07:15:00';Server=$server.Server;Level='INFO';Category='Maintenance';Count=1;First='2026-10-18 07:15:00';Message='Maintenance control applied: revision=3; enabled=False; catchUpAll=True; reason=Sample planning resumed.';Source='Demo/Logs/Orchestrator.log'})
    }
    $logEvents.Add([pscustomobject]@{Time='2026-10-18 08:41:12';Server='CTSO-INV-02';Level='WARN';Category='Warning';Count=2;First='2026-10-18 08:12:40';Message='Capability TeamsPowerShell unavailable on the sample server.';Source='Demo/Logs/Orchestrator.log'})
    if ($Scenario -eq 'Issues') {
        $logEvents.Add([pscustomobject]@{Time='2026-10-18 09:24:31';Server='CTSO-INV-03';Level='WARN';Category='Heartbeat';Count=14;First='2026-10-18 09:11:02';Message='Failed to write heartbeat: sample shared folder write timeout.';Source='Demo/Logs/Orchestrator.log'})
        $logEvents.Add([pscustomobject]@{Time='2026-10-18 09:24:30';Server='CTSO-INV-03';Level='ERROR';Category='StatePersistence';Count=3;First='2026-10-18 09:12:30';Message='State persistence unavailable after 60 second(s) of UNC write retries; the scheduler remains alive but new job launches are paused (sample).';Source='Demo/Logs/Orchestrator.log'})
        $logEvents.Add([pscustomobject]@{Time='2026-10-18 09:24:29';Server='CTSO-INV-03';Level='ERROR';Category='ConfigReload';Count=14;First='2026-10-18 09:11:01';Message='Central cluster configuration reload rejected; the last valid cluster settings stay in effect (sample).';Source='Demo/Logs/Orchestrator.log'})
        $logEvents.Add([pscustomobject]@{Time='2026-10-18 09:27:45';Server='CTSO-INV-03';Level='WARN';Category='StopRequest';Count=1;First='2026-10-18 09:27:45';Message='Orphan stop request retired: target PID 4103 no longer runs (sample).';Source='Demo/Logs/Orchestrator.log'})
    }
    if ($Scenario -eq 'Issues') {
        $complianceJob=@($jobs | Where-Object Name -EQ 'Intune-Devices-Compliance-Inventory')[0]
        $complianceJob.Health='Stale'; $complianceJob.LastStatus='Failed'; $complianceJob.SuccessAgeHours=27
        $complianceJob.LastSuccess='2026-10-17 06:30'
        $incidents = @([pscustomobject]@{ObservedBy='CTSO-INV-02';Issue='Intune compliance source is stale. A retry is recommended (sample).'})
    }
    $requests = @([pscustomobject]@{
        BatchId='DEMO-20261018-090000'; Selection='Selected'; Created='2026-10-18 09:00'; Status='Running'
        Jobs=3; Pending=1; Failed=0; Cancelled=0; RequestedBy='demo.operator@contoso.example'
        Rows=@(
            [pscustomobject]@{JobName=$activeUsersJob.Name;Status='Success';OwnerServer=$activeUsersJob.Server;UpdatedAtUtc='2026-10-18 07:05';Detail='Sample inventory completed.'}
            [pscustomobject]@{JobName=$teamsJob.Name;Status='Running';OwnerServer=$teamsJob.Server;UpdatedAtUtc='2026-10-18 07:26';Detail='Sample channel inventory in progress.'}
            [pscustomobject]@{JobName=$spoJob.Name;Status='Pending';OwnerServer=$spoJob.Server;UpdatedAtUtc='2026-10-18 07:26';Detail='Waiting for Teams (sample).'}
        )
    })
    [pscustomobject]@{
        Clock=$clock; Scenario=$Scenario; Jobs=$jobs; Servers=$servers; History=$history
        Running=$running; Pending=$pending; Incidents=$incidents; Requests=$requests; LogEvents=@($logEvents | Sort-Object Time -Descending)
        Maintenance=($Scenario -eq 'Maintenance'); Sequence=1
        Versions=@([pscustomobject]@{VersionId='DEMO-v003';Created='2026-10-18 08:45'},[pscustomobject]@{VersionId='DEMO-v002';Created='2026-10-17 16:20'})
    }
}

function Write-OrchestratorDemoActivity {
    param([string]$Message)
    $c = $script:OrchestratorDemo.Controls
    $c.StatusText.Text = 'DEMO: ' + $Message
    $c.ActivityBox.AppendText("[2026-10-18 09:30:00] [DEMO] $Message`r`n")
    $c.ActivityBox.ScrollToEnd()
}

function Update-OrchestratorDemoHistory {
    $s = $script:OrchestratorDemo
    $c = $s.Controls
    $rows = @($s.Data.History | Where-Object {
        $date = [datetime]::ParseExact($_.EventTime,'yyyy-MM-dd HH:mm:ss',[cultureinfo]::InvariantCulture)
        (-not $c.HistoryFromPicker.SelectedDate -or $date.Date -ge $c.HistoryFromPicker.SelectedDate) -and
        (-not $c.HistoryToPicker.SelectedDate -or $date.Date -le $c.HistoryToPicker.SelectedDate) -and
        ($c.HistoryServerCombo.SelectedItem -eq 'All' -or $_.Server -eq $c.HistoryServerCombo.SelectedItem) -and
        ($c.HistoryStatusCombo.SelectedItem -eq 'All' -or $_.Status -eq $c.HistoryStatusCombo.SelectedItem) -and
        ([string]::IsNullOrWhiteSpace($c.HistoryJobCombo.Text) -or $c.HistoryJobCombo.Text -eq 'All' -or $_.JobName -eq $c.HistoryJobCombo.Text)
    } | Sort-Object EventTime -Descending)
    $s.FilteredHistory = $rows
    $c.HistoryGrid.ItemsSource = $rows
}

function Update-OrchestratorDemoEvents {
    $s = $script:OrchestratorDemo
    $c = $s.Controls
    $since = $s.Data.Clock.AddHours(-1 * [int]$c.EventsPeriodCombo.SelectedItem.Hours).ToString('yyyy-MM-dd HH:mm:ss')
    $level = [string]$c.EventsLevelCombo.SelectedItem.Value
    $text = $c.EventsTextBox.Text.Trim()
    $rows = @($s.Data.LogEvents | Where-Object {
        ($c.EventsServerCombo.SelectedItem -eq 'All' -or $_.Server -eq $c.EventsServerCombo.SelectedItem) -and
        ($level -eq 'All' -or $_.Level -eq $level -or ($level -eq 'ErrorsAndWarnings' -and $_.Level -in @('ERROR','WARN'))) -and
        ($c.EventsCategoryCombo.SelectedItem -eq 'All' -or $_.Category -eq $c.EventsCategoryCombo.SelectedItem) -and
        [string]::CompareOrdinal($_.Time, $since) -ge 0 -and (-not $text -or $_.Message -like "*$text*")
    })
    $s.FilteredEvents = $rows
    $c.EventsGrid.ItemsSource = $rows
    $c.EventsSummaryText.Text = 'DEMO - {0} sample event row(s). Resident logs are never read in demo mode.' -f $rows.Count
}

function New-OrchestratorDemoDashboard {
    # Same shape as Get-SmartM365OrchestratorDashboard, built from the fictional model only.
    param([Parameter(Mandatory)]$Data)
    $clock = $Data.Clock
    $actions = [Collections.Generic.List[object]]::new()
    $failed = @($Data.History | Where-Object { $_.Status -eq 'Failed' -and [datetime]$_.EventTime -ge $clock.AddHours(-24) })
    foreach ($row in $failed | Select-Object -First 3) { $actions.Add([pscustomobject]@{ Severity = 'Critical'; Area = 'Job Failing'; Item = $row.JobName; Detail = 'Sample failure at {0}.' -f $row.EventTime; Target = 'Planning' }) }
    foreach ($job in @($Data.Jobs | Where-Object Health -eq 'Stale')) { $actions.Add([pscustomobject]@{ Severity = 'Warning'; Area = 'Job Stale'; Item = $job.Name; Detail = ('Last sample success {0} ({1} h).' -f $job.LastSuccess, $job.SuccessAgeHours); Target = 'Planning' }) }
    foreach ($incident in @($Data.Incidents)) { $actions.Add([pscustomobject]@{ Severity = 'Warning'; Area = 'Peer monitoring'; Item = $incident.ObservedBy; Detail = $incident.Issue; Target = 'Operations' }) }
    if ($Data.Maintenance) { $actions.Add([pscustomobject]@{ Severity = 'Warning'; Area = 'Maintenance'; Item = 'Scheduling'; Detail = 'Sample maintenance is ON: scheduled launches are suspended.'; Target = 'Operations' }) }
    foreach ($logEvent in @($Data.LogEvents | Where-Object { $_.Level -eq 'ERROR' })) { $actions.Add([pscustomobject]@{ Severity = 'Warning'; Area = 'Log ' + $logEvent.Category; Item = $logEvent.Server; Detail = $logEvent.Message; Target = 'Events' }) }
    $sorted = @($actions | Sort-Object @{ Expression = { @{ Critical = 0; Warning = 1; Info = 2 }[$_.Severity] } }, Area)
    $enabled = @($Data.Jobs | Where-Object Enabled)
    $unhealthy = @($Data.Jobs | Where-Object { $_.Health -in @('Stale', 'Failing') })
    $errors = @($Data.LogEvents | Where-Object Level -eq 'ERROR').Count
    $tiles = @(
        [pscustomobject]@{ Label = 'Jobs'; Value = [string]$Data.Jobs.Count; Detail = '{0} disabled (sample)' -f ($Data.Jobs.Count - $enabled.Count); Level = 'Normal'; Target = 'Planning' }
        [pscustomobject]@{ Label = 'Enabled'; Value = [string]$enabled.Count; Detail = 'scheduled or manual'; Level = 'Normal'; Target = 'Planning' }
        [pscustomobject]@{ Label = 'Servers online'; Value = '{0}/{0}' -f $Data.Servers.Count; Detail = 'all sample heartbeats fresh'; Level = 'Normal'; Target = 'Operations' }
        [pscustomobject]@{ Label = 'Running'; Value = [string]$Data.Running.Count; Detail = 'jobs in progress'; Level = 'Normal'; Target = 'Operations' }
        [pscustomobject]@{ Label = 'Pending'; Value = [string]$Data.Pending.Count; Detail = $(if ($Data.Pending.Count) { 'mostly Dependency' } else { 'nothing waiting' }); Level = $(if ($Data.Pending.Count) { 'Info' } else { 'Normal' }); Target = 'Operations' }
        [pscustomobject]@{ Label = 'Job health'; Value = [string]$unhealthy.Count; Detail = '{0} stale or failing' -f $unhealthy.Count; Level = $(if ($unhealthy.Count) { 'Warning' } else { 'Normal' }); Target = 'Planning' }
        [pscustomobject]@{ Label = 'Failures 24 h'; Value = [string]$failed.Count; Detail = $(if ($failed.Count) { (@($failed.JobName | Sort-Object -Unique) | Select-Object -First 2) -join ', ' } else { 'no failed run' }); Level = $(if ($failed.Count) { 'Critical' } else { 'Normal' }); Target = 'Failures24h' }
        [pscustomobject]@{ Label = 'Log errors 24 h'; Value = [string]$errors; Detail = '{0} warning(s)' -f @($Data.LogEvents | Where-Object Level -eq 'WARN').Count; Level = $(if ($errors) { 'Warning' } else { 'Normal' }); Target = 'Events' }
        [pscustomobject]@{ Label = 'Certificate'; Value = '212 d'; Detail = 'days before expiry (sample)'; Level = 'Normal'; Target = 'Operations' }
    )
    $servers = @(foreach ($server in $Data.Servers) {
        [pscustomobject]@{ Server = $server.Server; State = 'Online'; Version = $server.Version; Running = $server.Running; Pending = $server.Pending; RecycleIn = '14h30'
            Owned = $server.AssignedJobs; Certificate = '212 d'; StatePersistence = 'OK'; Capabilities = $server.Capabilities }
    })
    $items = [Collections.Generic.List[object]]::new()
    foreach ($job in @($Data.Jobs | Where-Object { $_.Enabled -and $_.Assignment -ne 'Manual' })) {
        foreach ($day in @($clock.Date, $clock.Date.AddDays(1))) {
            if ($job.Frequency -eq 'Weekly' -and [string]$day.DayOfWeek -notin @($job.ScheduleDays)) { continue }
            foreach ($text in @([string]$job.Times -split ',\s*' | Where-Object { $_ })) {
                $time = [timespan]::Zero
                if (-not [timespan]::TryParseExact($text, 'hh\:mm', [cultureinfo]::InvariantCulture, [ref]$time)) { continue }
                $start = $day.Add($time)
                if ($start -lt $clock -or $start -gt $clock.AddHours(24)) { continue }
                $items.Add([pscustomobject]@{ Lane = $job.Server; Job = $job.Name; Start = $start; End = $start.AddMinutes([math]::Max(5, [double]$job.AverageDurationMinutes)); Kind = $(if ($Data.Maintenance) { 'Suspended' } else { 'Planned' }) })
            }
        }
    }
    foreach ($running in @($Data.Running)) { $start = [datetime]$running.Started; $items.Add([pscustomobject]@{ Lane = $running.Server; Job = $running.Job; Start = $start; End = $clock.AddMinutes(25); Kind = 'Running' }) }
    $critical = @($sorted | Where-Object Severity -eq 'Critical').Count; $warning = @($sorted | Where-Object Severity -eq 'Warning').Count
    $summary = 'DEMO - {0}/{0} sample server(s) online | {1} running | {2} pending | {3} failure(s) in 24 h' -f $Data.Servers.Count, $Data.Running.Count, $Data.Pending.Count, $failed.Count
    [pscustomobject]@{
        Banner = $(if ($critical) { [pscustomobject]@{ Level = 'Critical'; Text = '{0} critical and {1} warning item(s) to review. {2}' -f $critical, $warning, $summary } } elseif ($warning) { [pscustomobject]@{ Level = 'Warning'; Text = '{0} warning item(s) to review. {1}' -f $warning, $summary } } else { [pscustomobject]@{ Level = 'Normal'; Text = 'Everything looks healthy. ' + $summary } })
        Tiles = $tiles; Actions = $sorted; Servers = $servers
        Timeline = [pscustomobject]@{ From = $clock; To = $clock.AddHours(24); Lanes = @($Data.Servers.Server); Items = @($items | Sort-Object Start) }
        GeneratedAt = $clock
    }
}

function New-OrchestratorDemoConfiguration {
    # Fictional Contoso files built from the shipped templates only; no runtime JSON is read.
    $configFolder = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'Config'
    $readTemplate = { param($name) Get-Content -LiteralPath (Join-Path $configFolder $name) -Raw -Encoding UTF8 | ConvertFrom-Json }
    $globalTemplate = & $readTemplate 'SmartM365.global.local.json.template'
    $tenantTemplate = & $readTemplate 'Tenants\tenant.local.json.template'
    $mailTemplate = & $readTemplate 'Tenants\tenant.mail.local.json.template'
    $global = $globalTemplate | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $global.DefaultTenant = 'contoso'; $global.MailClientName = 'Contoso (DEMO)'
    $tenant = $tenantTemplate | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $tenant.ProfileKey = 'contoso'; $tenant.OrganizationKey = 'contoso'; $tenant.EnvironmentKey = 'demo'; $tenant.TenantKey = 'contoso-demo'; $tenant.Environment = 'Demo'
    $tenant.AppId = '11111111-2222-3333-4444-555555555555'; $tenant.TenantId = '66666666-7777-8888-9999-000000000000'
    $tenant.Thumb = 'DEMO000000000000000000000000000000000000'; $tenant.Thumbprint = 'DEMO000000000000000000000000000000000000'
    $tenant.OrgDomain = 'contoso.onmicrosoft.com'; $tenant.ErrorMailTo = 'it-alerts@contoso.example'; $tenant.To = 'it-reports@contoso.example'
    $tenant.From = 'smartm365@contoso.example'; $tenant.MailSendAccessPolicyGroup = 'smartm365-mailsend@contoso.example'
    $tenant.TeamsAlertsWebhookUrl = 'https://contoso.example/demo-webhook/alerts'; $tenant.EnableSharePointUpload = $true
    $tenant.PSObject.Properties.Remove('RemoteRoutingDomain')
    $tenant | Add-Member -NotePropertyName 'LegacyReportFolder' -NotePropertyValue 'D:\Reports (sample key not in the template)'
    $mail = $mailTemplate | ConvertTo-Json -Depth 20 | ConvertFrom-Json
    $mail.Scripts.'SmartM365-Licences-Inventory'.From = 'licensing@contoso.example'
    $mail.Scripts.'SmartM365-Licences-Inventory'.Cc = 'finance@contoso.example'
    $descriptions = Get-Content -LiteralPath (Join-Path $configFolder 'SmartM365.ConfigDescriptions.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $files = @(
        [pscustomobject]@{ Kind = 'Global'; Title = 'SmartM365.global.local.json'; Path = 'Config\SmartM365.global.local.json.txt (DEMO)'; Status = 'OK'; Error = ''; Document = $global; Template = $globalTemplate; DescriptionKey = 'SmartM365.global.local.json'; SourceName = 'global' }
        [pscustomobject]@{ Kind = 'Tenant'; Title = 'Tenants\contoso.local.json'; Path = 'Config\Tenants\contoso.local.json.txt (DEMO)'; Status = 'OK'; Error = ''; Document = $tenant; Template = $tenantTemplate; DescriptionKey = 'Tenants/tenant.local.json'; SourceName = 'tenant' }
        [pscustomobject]@{ Kind = 'Mail'; Title = 'Tenants\contoso.mail.local.json'; Path = 'Config\Tenants\contoso.mail.local.json.txt (DEMO)'; Status = 'OK'; Error = ''; Document = $mail; Template = $mailTemplate; DescriptionKey = 'Tenants/tenant.mail.local.json'; SourceName = 'mail' }
    )
    $view = New-SmartM365OrchestratorConfigurationView -ProfileKey 'contoso' -Files $files -Descriptions $descriptions
    $view | Add-Member -NotePropertyName ConfigFolder -NotePropertyValue $configFolder
    $view | Add-Member -NotePropertyName ReadAt -NotePropertyValue $script:OrchestratorDemo.Data.Clock
    return $view
}

function Update-OrchestratorDemoConfiguration {
    $s = $script:OrchestratorDemo; $c = $s.Controls
    if ($null -eq $s.Configuration) {
        try { $s.Configuration = New-OrchestratorDemoConfiguration }
        catch { $c.ConfigSummaryText.Text = 'DEMO configuration unavailable: copy Config templates and SmartM365.ConfigDescriptions.json with the GUI. ' + $_.Exception.Message; return }
    }
    $c.ConfigShowSensitiveButton.Content = if ($s.ConfigurationShow) { 'Hide' } else { 'Show' }
    foreach ($pair in @(@('Global', 'ConfigGlobal'), @('Tenant', 'ConfigTenant'), @('Mail', 'ConfigMail'))) {
        $file = @($s.Configuration.Files | Where-Object Kind -eq $pair[0])[0]
        $rows = @(Select-SmartM365OrchestratorConfigRows -Rows @($file.Rows) -Filter $c.ConfigFilterBox.Text.Trim() -ShowSensitive:$s.ConfigurationShow)
        $c[$pair[1] + 'Grid'].ItemsSource = $rows
        $c[$pair[1] + 'Tab'].Header = '{0} ({1})' -f $pair[0], [IO.Path]::GetFileName(($file.Path -replace ' \(DEMO\)$', ''))
        $c[$pair[1] + 'StatusText'].Text = 'DEMO: {0}. {1} key(s), {2} shown; {3} not in template, {4} missing from file; {5} sensitive value(s) {6}.' -f $file.Path, @($file.Rows).Count, $rows.Count, $file.NotInTemplate, $file.MissingFromFile, @($file.Rows | Where-Object Sensitive).Count, $(if ($s.ConfigurationShow) { 'shown' } else { 'hidden' })
    }
    $c.ConfigSummaryText.Text = 'DEMO - fictional Contoso configuration built from the shipped templates. Read-only: nothing is written; sensitive values are hidden unless Show is pressed.'
}

function Update-OrchestratorDemoViews {
    $s=$script:OrchestratorDemo; $d=$s.Data; $c=$s.Controls
    foreach($server in $d.Servers){
        $server.AssignedJobs=@($d.Jobs | Where-Object Server -EQ $server.Server).Count
        $server.Running=@($d.Running | Where-Object Server -EQ $server.Server).Count
        $server.Pending=@($d.Pending | Where-Object Server -EQ $server.Server).Count
        $server.Maintenance=if($d.Maintenance){'Paused'}else{'Active'}
        $server.MaintenanceAcknowledgement=if($d.Maintenance){'Applied (sample)'}else{'Ready'}
    }
    $c.ServersGrid.ItemsSource=@($d.Servers)
    $c.OperationsServersGrid.ItemsSource=@($d.Servers)
    $c.PlanningGrid.ItemsSource=@($d.Jobs)
    $c.OperationsRunningGrid.ItemsSource=@($d.Running)
    $c.OperationsPendingGrid.ItemsSource=@($d.Pending)
    $c.OperationsIncidentsGrid.ItemsSource=@($d.Incidents)
    $c.RequestsGrid.ItemsSource=@($d.Requests)
    $c.VersionsGrid.ItemsSource=@($d.Versions)
    $s.Dashboard=New-OrchestratorDemoDashboard -Data $d
    Set-OrchestratorDashboardView -Controls $c -Dashboard $s.Dashboard
    $c.MaintenanceBannerText.Text=if($d.Maintenance){'DEMO - Scheduling maintenance enabled'}else{'DEMO - Scheduling active'}
    $c.MaintenanceDetailText.Text=if($d.Maintenance){'Contoso demo maintenance window. Running sample jobs remain visible.'}else{'All three sample residents are ready. Actions affect this demonstration only.'}
    $c.EnableMaintenanceButton.IsEnabled=-not $d.Maintenance
    $c.DisableMaintenanceButton.IsEnabled=$d.Maintenance
    $c.StopRunningJobButton.IsEnabled=$d.Running.Count -gt 0
    $c.CancelRequestButton.IsEnabled=@($d.Requests | Where-Object Pending -GT 0).Count -gt 0
    $c.CancelAllRequestsButton.IsEnabled=$c.CancelRequestButton.IsEnabled
    if($d.Jobs.Count){$c.PlanningGrid.SelectedIndex=0}
    if($d.Servers.Count){$c.ServersGrid.SelectedIndex=0}
    if($d.Requests.Count){$c.RequestsGrid.SelectedIndex=0}
    if($d.Running.Count){$c.OperationsRunningGrid.SelectedIndex=0}
    if($d.Versions.Count){$c.VersionsGrid.SelectedIndex=0}
    Update-OrchestratorDemoHistory
    Update-OrchestratorDemoEvents
    Update-OrchestratorDemoConfiguration
}

function Export-OrchestratorDemoHistory {
    param([ValidateSet('Html','Csv')][string]$Format='Html')
    $s=$script:OrchestratorDemo
    # The only disk output in the demo. The root is generated per launch, never read from runtime config.
    [void][IO.Directory]::CreateDirectory($s.OutputRoot)
    $columns=@('EventTime','StartTime','Server','JobName','Status','DurationSec','ExitCode','RetryCount','LogPath')
    $path=Join-Path $s.OutputRoot ('SmartM365-Orchestrator-History-DEMO.'+$Format.ToLowerInvariant())
    if($Format -eq 'Html') {
        $body=$s.FilteredHistory | Select-Object $columns | ConvertTo-Html -Title 'SmartM365 Orchestrator history - DEMO' -PreContent '<h1>SmartM365 Orchestrator history</h1><p><strong>DEMO - Sample data</strong> | Contoso | 2026-10-18 09:30 | Synthetic execution history</p>'
        [IO.File]::WriteAllText($path,($body -join [Environment]::NewLine),[Text.UTF8Encoding]::new($false))
    } else {
        $s.FilteredHistory | Select-Object $columns | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8
    }
    $displayPath='%TEMP%\SmartM365-Demo\InventoryOrchestrator\'+(Split-Path $s.OutputRoot -Leaf)+'\'+(Split-Path $path -Leaf)
    Write-OrchestratorDemoActivity "Sample $Format history exported to $displayPath"
    return $path
}

function Invoke-OrchestratorDemoAction {
    param([string]$Action)
    $s=$script:OrchestratorDemo; $d=$s.Data; $c=$s.Controls
    switch($Action) {
        'RefreshButton' { Update-OrchestratorDemoViews; Write-OrchestratorDemoActivity 'Sample views refreshed. Clock remains frozen.'; return }
        'HistoryRefreshButton' { Update-OrchestratorDemoHistory; Write-OrchestratorDemoActivity "$($s.FilteredHistory.Count) sample history rows match the filters."; return }
        'Failures24hButton' {
            $c.HistoryFromPicker.SelectedDate=$d.Clock.Date; $c.HistoryToPicker.SelectedDate=$d.Clock.Date
            $c.HistoryStatusCombo.SelectedItem='Failed'; $c.MainTabs.SelectedIndex=3
            Update-OrchestratorDemoHistory; return
        }
        'ExportHtmlButton' { [void](Export-OrchestratorDemoHistory -Format Html); return }
        'ExportCsvButton' { [void](Export-OrchestratorDemoHistory -Format Csv); return }
        'EnableMaintenanceButton' { $d.Maintenance=$true }
        'DisableMaintenanceButton' { $d.Maintenance=$false }
        'ValidateButton' {
            Write-OrchestratorDemoActivity 'Sample draft is valid. Validation was performed on fictional data only.'
            return
        }
        'ApplyJobButton' {
            $job=$c.PlanningGrid.SelectedItem; if(-not $job){return}
            $times=@($c.TimesBox.Text -split ',' | ForEach-Object {$_.Trim()})
            foreach($text in $times){
                $time=[timespan]::Zero
                if(-not [timespan]::TryParseExact($text,'hh\:mm',[cultureinfo]::InvariantCulture,[ref]$time) -or $time.TotalHours -ge 24){Write-OrchestratorDemoActivity 'Enter sample times as HH:mm, separated by commas.';return}
            }
            $numeric=@{}
            foreach($pair in @(@('TimeoutBox','TimeoutMinutes'),@('RetriesBox','MaxRetries'),@('RetryDelayBox','RetryDelaySeconds'),@('DurationBox','AverageDurationMinutes'),@('DependencyMaxAgeBox','DependencyMaxAgeHours'))){
                $value=0
                if(-not [int]::TryParse($c[$pair[0]].Text,[ref]$value) -or $value -lt 0){Write-OrchestratorDemoActivity 'Enter non-negative sample numeric settings.';return}
                $numeric[$pair[1]]=$value
            }
            $job.Enabled=[bool]$c.JobEnabledCheck.IsChecked
            $job.Times=$c.TimesBox.Text; $job.DependsOn=$c.DependsOnBox.Text
            $job.Frequency=[string]$c.ScheduleTypeCombo.SelectedItem.Content
            $job.Assignment=[string]$c.AssignmentCombo.SelectedItem.Content
            $job.DependencyMode=[string]$c.DependencyModeCombo.SelectedItem.Content
            $job.MissedRunPolicy=[string]$c.MissedPolicyCombo.SelectedItem.Content
            $dayNames=@('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday')
            $job.ScheduleDays=@($dayNames | Where-Object {$c[$_+'Check'].IsChecked})
            $job.Days=if($job.Frequency -eq 'Weekly'){$job.ScheduleDays -join ', '}else{'Every day'}
            foreach($key in $numeric.Keys){$job.$key=$numeric[$key]}
            if($job.Assignment -eq 'Pinned' -and $c.PinnedServerCombo.SelectedItem){$job.Server=[string]$c.PinnedServerCombo.SelectedItem}
            $job.NextRun=Get-OrchestratorDemoNextRun -Job $job -Clock $d.Clock
        }
        'ApplyServerButton' {
            $server=$c.ServersGrid.SelectedItem; if(-not $server){return}
            $weight=0.0
            if(-not [double]::TryParse($c.ServerWeightBox.Text,[ref]$weight) -or $weight -le 0){Write-OrchestratorDemoActivity 'Enter a positive sample server weight.';return}
            $server.Weight=$weight
        }
        'AddServerButton' {
            $name=$c.NewServerBox.Text.Trim()
            if($name -notmatch '^CTSO-[A-Z0-9-]{1,30}$' -or $name -in $d.Servers.Server){Write-OrchestratorDemoActivity 'Use a unique fictional CTSO- server name.';return}
            $copy=$d.Servers[0].PSObject.Copy(); $copy.Server=$name; $copy.AssignedJobs=0; $copy.Running=0
            $d.Servers=@($d.Servers)+$copy
        }
        'RemoveServerButton' {
            $row=$c.ServersGrid.SelectedItem
            if(-not $row -or $d.Servers.Count -le 1){return}
            $remaining=@($d.Servers | Where-Object Server -NE $row.Server)
            foreach($job in $d.Jobs){if($job.Server -eq $row.Server){$job.Server=$remaining[0].Server}}
            foreach($item in @($d.Running)+@($d.Pending)){if($item.Server -eq $row.Server){$item.Server=$remaining[0].Server}}
            $d.Servers=$remaining
        }
        'RequestRunButton' {
            $job=$c.PlanningGrid.SelectedItem; if(-not $job){return}
            $d.Sequence++
            $d.Requests=@([pscustomobject]@{BatchId=('DEMO-REQUEST-{0:000}' -f $d.Sequence);Selection='Selected';Created='2026-10-18 09:30';Status='Pending';Jobs=1;Pending=1;Failed=0;Cancelled=0;RequestedBy='demo.operator@contoso.example';Rows=@([pscustomobject]@{JobName=$job.Name;Status='Pending';OwnerServer=$job.Server;UpdatedAtUtc='2026-10-18 07:30';Detail='Simulated request only.'})})+@($d.Requests)
        }
        {$_ -in @('CancelRequestButton','CancelAllRequestsButton')} {
            $targets=if($Action -eq 'CancelAllRequestsButton'){@($d.Requests)}else{@($c.RequestsGrid.SelectedItem)}
            foreach($request in $targets){
                if($null -eq $request){continue}
                foreach($row in $request.Rows){if($row.Status -eq 'Pending'){$row.Status='Cancelled';$row.Detail='Cancelled in demo memory.';$request.Cancelled++}}
                $request.Pending=0
                $request.Status=if(@($request.Rows | Where-Object Status -EQ 'Running').Count){'Running'}else{'Cancelled'}
            }
        }
        'StopRunningJobButton' {
            $row=$c.OperationsRunningGrid.SelectedItem
            if($row){
                $d.Running=@($d.Running | Where-Object Job -NE $row.Job)
                foreach($request in $d.Requests){foreach($item in $request.Rows){if($item.JobName -eq $row.Job -and $item.Status -eq 'Running'){$item.Status='Cancelled';$item.Detail='Stopped in demo memory.';$request.Cancelled++}}}
            }
        }
        'RebalanceButton' { for($i=0;$i -lt $d.Jobs.Count;$i++){$d.Jobs[$i].Server=$d.Servers[($i+1)%$d.Servers.Count].Server} }
        'RollbackButton' { $s.Data=New-OrchestratorDemoData -Scenario $d.Scenario }
        'PublishButton' {
            $d.Sequence++
            $d.Versions=@([pscustomobject]@{VersionId=('DEMO-v{0:000}' -f (3+$d.Sequence));Created='2026-10-18 09:30'})+@($d.Versions)
        }
    }
    Update-OrchestratorDemoViews
    Write-OrchestratorDemoActivity "$Action simulated in memory. No operational request or configuration was written."
}

function Select-OrchestratorDemoTarget {
    param([string]$Target)
    $c=$script:OrchestratorDemo.Controls
    $header=switch($Target){'Failures24h'{'History'}default{$Target}}
    if($Target -eq 'Failures24h'){$c.HistoryStatusCombo.SelectedItem='Failed';Update-OrchestratorDemoHistory}
    if($Target -eq 'Events'){$c.EventsLevelCombo.SelectedIndex=1}
    foreach($tab in $c.MainTabs.Items){if($tab.Header -eq $header){$c.MainTabs.SelectedItem=$tab;break}}
}

function Invoke-OrchestratorDemo {
    param([string]$Xaml,[string]$Scenario,[string]$View,[int]$Width,[int]$Height,[switch]$ValidateOnly)
    if([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA){throw 'The demo requires STA. Use the Demo CMD launcher.'}
    $window=ConvertFrom-OrchestratorGuiXaml -Text $Xaml
    $window.Title='SmartM365 Orchestrator | DEMO - Sample data'
    $window.WindowState='Normal'; $window.Width=$Width; $window.Height=$Height
    $controls=@{}
    foreach($match in [regex]::Matches($Xaml,'x:Name="(?<Name>[A-Za-z][A-Za-z0-9_]*)"')){
        $name=$match.Groups['Name'].Value
        $control=$window.FindName($name)
        if($control){$controls[$name]=$control}
    }
    $script:OrchestratorDemo=[pscustomobject]@{
        Mode='Demo';Window=$window;Controls=$controls;Data=(New-OrchestratorDemoData -Scenario $Scenario)
        FilteredHistory=@();FilteredEvents=@();Dashboard=$null;Configuration=$null;ConfigurationShow=$false;OutputRoot=(Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Demo/InventoryOrchestrator/'+[guid]::NewGuid().ToString('N')))
    }
    $s=$script:OrchestratorDemo; $c=$s.Controls
    $body=$window.Content; $window.Content=$null
    $dock=[System.Windows.Controls.DockPanel]::new()
    $bar=[System.Windows.Controls.StackPanel]::new();$bar.Orientation='Horizontal';$bar.Margin='18,8,18,0'
    [System.Windows.Controls.DockPanel]::SetDock($bar,'Top')
    $label=[System.Windows.Controls.TextBlock]::new();$label.Text='DEMO - Sample data | Scene';$label.VerticalAlignment='Center';$label.Margin='0,0,12,0'
    [void]$bar.Children.Add($label)
    $combo=[System.Windows.Controls.ComboBox]::new();$combo.Width=170;$combo.ItemsSource=@('Overview','Issues','Maintenance');$combo.SelectedItem=$Scenario
    $c.DemoScenarioCombo=$combo
    [void]$bar.Children.Add($combo);[void]$dock.Children.Add($bar);[void]$dock.Children.Add($body);$window.Content=$dock
    $combo.Add_SelectionChanged({
        $s=$script:OrchestratorDemo
        $s.Data=New-OrchestratorDemoData -Scenario ([string]$s.Controls.DemoScenarioCombo.SelectedItem)
        Update-OrchestratorDemoViews
        Write-OrchestratorDemoActivity ('Sample scene loaded: '+$s.Data.Scenario)
    })
    $c.SharedPathText.Text='Contoso | DEMO - Sample data | All actions are simulated'
    $c.ConnectionText.Text='Offline demo'; $c.LastRefreshText.Text='2026-10-18 09:30 (frozen)'
    $c.FooterText.Text='DEMO - Sample data. Exports use a separate temporary demo folder.'
    $c.VersionText.Text='v1.6.0 | DEMO'
    $c.AutoRefreshCheck.IsChecked=$false; $c.AutoRefreshCheck.IsEnabled=$false
    # Keep both running sample rows visible at the default capture size.
    $operationsLayout=$c.OperationsServersGrid.Parent
    $operationsLayout.RowDefinitions[1].Height=[System.Windows.GridLength]::new(110)
    $operationsLayout.RowDefinitions[3].Height=[System.Windows.GridLength]::new(150)
    $c.HistoryFromPicker.SelectedDate=$s.Data.Clock.AddDays(-6).Date
    $c.HistoryToPicker.SelectedDate=$s.Data.Clock.Date
    $c.HistoryServerCombo.ItemsSource=@('All')+@($s.Data.Servers.Server);$c.HistoryServerCombo.SelectedIndex=0
    $c.HistoryJobCombo.ItemsSource=@('All')+@($s.Data.Jobs.Name);$c.HistoryJobCombo.SelectedIndex=0
    $c.HistoryStatusCombo.ItemsSource=@('All','Success','CompletedWithWarnings','Failed','Cancelled');$c.HistoryStatusCombo.SelectedIndex=0
    $c.PinnedServerCombo.ItemsSource=@($s.Data.Servers.Server)
    $c.EventsServerCombo.ItemsSource=@('All')+@($s.Data.Servers.Server);$c.EventsServerCombo.SelectedIndex=0
    $c.EventsLevelCombo.DisplayMemberPath='Label'
    $c.EventsLevelCombo.ItemsSource=@(foreach($pair in @(@('All','All'),@('Errors and warnings','ErrorsAndWarnings'),@('ERROR','ERROR'),@('WARN','WARN'),@('INFO','INFO'))){[pscustomobject]@{Label=$pair[0];Value=$pair[1]}});$c.EventsLevelCombo.SelectedIndex=0
    $c.EventsCategoryCombo.ItemsSource=@('All','Lifecycle','StopRequest','Heartbeat','StatePersistence','ConfigReload','Maintenance','Error','Warning');$c.EventsCategoryCombo.SelectedIndex=0
    $c.EventsPeriodCombo.DisplayMemberPath='Label'
    $c.EventsPeriodCombo.ItemsSource=@(foreach($pair in @(@('1 h',1),@('6 h',6),@('24 h',24),@('3 days',72),@('7 days',168))){[pscustomobject]@{Label=$pair[0];Hours=$pair[1]}});$c.EventsPeriodCombo.SelectedIndex=2
    $c.OperationsMailsGrid.ItemsSource=@([pscustomobject]@{Time='2026-10-18 08:20';Server='CTSO-INV-01';Subject='DEMO - Daily inventory completed (sample)'})
    foreach($name in @('MaintenanceReasonBox','JobStopReasonBox','CancellationReasonBox')){$c[$name].Text='Contoso demonstration'}
    $c.PlanningGrid.Add_SelectionChanged({
        $c=$script:OrchestratorDemo.Controls;$job=$c.PlanningGrid.SelectedItem;if(-not $job){return}
        $c.SelectedJobText.Text=$job.Name;$c.JobArgumentsBox.Text=$job.Arguments;$c.JobEnabledCheck.IsChecked=$job.Enabled
        $c.TimesBox.Text=$job.Times;$c.DependsOnBox.Text=$job.DependsOn
        foreach($pair in @(@('ScheduleTypeCombo',$job.Frequency),@('AssignmentCombo',$job.Assignment),@('DependencyModeCombo',$job.DependencyMode),@('MissedPolicyCombo',$job.MissedRunPolicy))){
            foreach($item in $c[$pair[0]].Items){if($item.Content -eq $pair[1]){$c[$pair[0]].SelectedItem=$item;break}}
        }
        foreach($pair in @(@('TimeoutBox','TimeoutMinutes'),@('RetriesBox','MaxRetries'),@('RetryDelayBox','RetryDelaySeconds'),@('DurationBox','AverageDurationMinutes'),@('DependencyMaxAgeBox','DependencyMaxAgeHours'))){$c[$pair[0]].Text=[string]$job.($pair[1])}
        foreach($dayName in @('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday')){$c[$dayName+'Check'].IsChecked=$dayName -in $job.ScheduleDays}
        $dependents=@($script:OrchestratorDemo.Data.Jobs | Where-Object {$job.Name -in @($_.DependsOn -split ',\s*')} | ForEach-Object Name)
        $c.DependentsText.Text=if($dependents.Count){'Dependents: '+($dependents -join ', ')}else{'No dependent jobs in the shipped catalog.'}
        $c.ReadinessGrid.ItemsSource=@(foreach($dependency in @($job.DependsOn -split ',\s*' | Where-Object {$_})){
            [pscustomobject]@{Dependency=$dependency;State='Ready';AgeHours=0.5;MaxAgeHours=$job.DependencyMaxAgeHours;Detail='Sample successful occurrence.'}
        })
    })
    $c.ServersGrid.Add_SelectionChanged({$c=$script:OrchestratorDemo.Controls;$row=$c.ServersGrid.SelectedItem;if($row){$c.SelectedServerText.Text=$row.Server;$c.ServerWeightBox.Text=[string]$row.Weight}})
    $c.RequestsGrid.Add_SelectionChanged({$c=$script:OrchestratorDemo.Controls;$row=$c.RequestsGrid.SelectedItem;$c.RequestJobsGrid.ItemsSource=if($row){@($row.Rows)}else{@()}})
    $c.AssignmentCombo.Add_SelectionChanged({$c=$script:OrchestratorDemo.Controls;$c.PinnedServerCombo.IsEnabled=($c.AssignmentCombo.SelectedItem.Content -eq 'Pinned')})
    $c.ScheduleTypeCombo.Add_SelectionChanged({$c=$script:OrchestratorDemo.Controls;$c.DaysPanel.IsEnabled=($c.ScheduleTypeCombo.SelectedItem.Content -eq 'Weekly')})
    foreach($name in @('RefreshButton','ValidateButton','RebalanceButton','PublishButton','ApplyJobButton','ApplyServerButton','HistoryRefreshButton','ExportHtmlButton','ExportCsvButton','Failures24hButton','EnableMaintenanceButton','DisableMaintenanceButton','StopRunningJobButton','RequestRunButton','CancelRequestButton','CancelAllRequestsButton','AddServerButton','RemoveServerButton','RollbackButton')){
        $c[$name].Tag=$name
        $c[$name].Add_Click({param($demoControl)try{Invoke-OrchestratorDemoAction -Action ([string]$demoControl.Tag)}catch{Write-OrchestratorDemoActivity $_.Exception.Message}})
    }
    $c.HistoryGrid.Add_MouseDoubleClick({$c=$script:OrchestratorDemo.Controls;if($c.HistoryGrid.SelectedItem){$c.MainTabs.SelectedIndex=6;Write-OrchestratorDemoActivity ('Sample log: '+$c.HistoryGrid.SelectedItem.LogPath)}})
    $c.OperationsMailsGrid.Add_MouseDoubleClick({Write-OrchestratorDemoActivity 'Sample mail: Contoso inventory completed successfully. No email was sent.'})
    foreach($name in @('EventsServerCombo','EventsLevelCombo','EventsCategoryCombo','EventsPeriodCombo')){$c[$name].Add_SelectionChanged({Update-OrchestratorDemoEvents})}
    $c.EventsTextBox.Add_TextChanged({Update-OrchestratorDemoEvents})
    $c.ConfigFilterBox.Add_TextChanged({Update-OrchestratorDemoConfiguration})
    $c.ConfigShowSensitiveButton.Add_Click({$script:OrchestratorDemo.ConfigurationShow=-not $script:OrchestratorDemo.ConfigurationShow;Update-OrchestratorDemoConfiguration})
    $c.ConfigRefreshButton.Add_Click({Update-OrchestratorDemoConfiguration;Write-OrchestratorDemoActivity 'Sample configuration refreshed. No configuration file was read.'})
    $c.ConfigOpenFolderButton.Add_Click({Write-OrchestratorDemoActivity 'Open folder is disabled in demo mode.'})
    $c.DashboardTimelineCanvas.Add_SizeChanged({param($sender) Update-OrchestratorDashboardTimeline -Canvas $sender})
    $c.DashboardTilesList.AddHandler([System.Windows.UIElement]::MouseLeftButtonUpEvent,[System.Windows.Input.MouseButtonEventHandler]{param($sender,$eventArgs)$tile=$eventArgs.OriginalSource.DataContext;if($tile -and $tile.PSObject.Properties['Target']){Select-OrchestratorDemoTarget ([string]$tile.Target)}})
    $c.DashboardActionsGrid.Add_MouseDoubleClick({$row=$script:OrchestratorDemo.Controls.DashboardActionsGrid.SelectedItem;if($row){Select-OrchestratorDemoTarget ([string]$row.Target)}})
    $c.EventsRefreshButton.Add_Click({Update-OrchestratorDemoEvents;Write-OrchestratorDemoActivity 'Sample events refreshed. No log file was read.'})
    $c.EventsGrid.Add_SelectionChanged({$c=$script:OrchestratorDemo.Controls;$row=$c.EventsGrid.SelectedItem;$c.EventsDetailBox.Text=if($row){'[{0}][{1}][{2}] {3}{4}{5}' -f $row.Time,$row.Server,$row.Level,$row.Category,[Environment]::NewLine,$row.Message}else{''}})
    $logo=Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png'
    if(Test-Path -LiteralPath $logo){$c.HeaderLogo.Source=[System.Windows.Media.Imaging.BitmapFrame]::Create([uri]$logo)}
    $icon=Join-Path $PSScriptRoot 'WorkplaceCloudHub.ico'
    if(Test-Path -LiteralPath $icon){$window.Icon=[System.Windows.Media.Imaging.BitmapFrame]::Create([uri]$icon)}
    $c.HeaderLogo.Cursor='Hand';$c.HeaderLogo.ToolTip='Open WorkplaceCloudHub.com'
    $c.HeaderLogo.Add_MouseLeftButtonUp({Start-Process 'https://workplacecloudhub.com'})
    Update-OrchestratorDemoViews
    Write-OrchestratorDemoActivity 'Contoso sample data loaded. No connection, resident worker or collector was started.'
    foreach($tab in $c.MainTabs.Items){if($tab.Header -eq $View){$c.MainTabs.SelectedItem=$tab;break}}
    if($ValidateOnly){return $s}
    $splash=$null
    try {
        . (Join-Path $PSScriptRoot 'SmartM365.GuiSplash.ps1')
        $splash=Start-SmartM365GuiSplash -ProductName 'SmartM365 Orchestrator' -Subtitle 'DEMO - Sample data' -MinimumDurationMs 6000
        $script:OrchestratorDemoSplash=$splash
        $window.Add_ContentRendered({Hide-SmartM365GuiSplash -Splash $script:OrchestratorDemoSplash})
        [void]$window.ShowDialog()
    } finally {if($splash){Close-SmartM365GuiSplash -Splash $splash}}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCJzLEk46eon7u8
# 4nw58bx9ZB1r9TaTGoza+3O7v41iRKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOimbzz+Qx0vtIUCCbVFvzKBFX8ND6aK/0Y1H9pK/4HeMA0GCSqG
# SIb3DQEBAQUABIIBgCYw0Ib/XSAyv0OcxBrXmAkXSpMxLlQaaZk2rX6gqCG7NThW
# sFFyiGjvqN9okx2Pm7pFGXLew9e3f7gK+T/QRIicXf7uZTtwVlExoEtZsKmQRi1+
# MPaZuha03hKpJBL7FNYUl7XGnfidDP7q4uv0diTkmv3T4aJ4Y8UYy5j5CqmEmkOG
# +wUNWvaro/KlnXdJ4Wf9Sreov+gKvvKsMsBpDOAOA6483sDD7/WyudbSctUEp/Bu
# puzUvaXmayRA6k++mPUC6osN7lJLjWk30M9fRG+UJYBykIc1/rivf0pfkP1ckEvx
# cyiVPEwcgN5Rjyd8vcrp40XyLZeAxAJBNAw8q6BAxh1m9OB9vDO/9gQAuedER/6q
# Jz6dx93PYBDjd9p7KlisyR+B3iWy2M5ypdPbIX8W2OPfjUDzytJEzNC6H2qBqyKi
# LcctQ3To73hvKj+9e7/Q52OkImDcFzODGXsiMCwKA2DtLOPrWzAWm5sa2KCSJSsT
# v0ea9VSB0TFQ96neZaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxOTI2
# MjlaMC8GCSqGSIb3DQEJBDEiBCBIoQ+u/kZ3zAdIA0wEV9du398gLm12DIVVXZOl
# JZCOATANBgkqhkiG9w0BAQEFAASCAgCItLrNSQ6B09/1RUElwUBEJXjRrSPt/9Nv
# u+fRpaaVY42xYUevSNjHZyZkzkyb0hB6+cCnHfSmQRSJT6i6o4e2pCuPp5COpegf
# RjJ5nEVkBpMbukSrVyzVFZkLstUfHoTbFzl+bF/qNM0zY7MMuL+B6qa9Cl5v9j54
# H8G8iwdSMBbj+fkOHTNIfC+DShDPk8Pi8zvzJsUek9i32gZV3SVFLi5kD6LA0dZL
# egKzHP9/ThH49OrepfpgdqqgnD6K6OB2S13EvOxgtwq7OANLA+SFJnKtGRegqnWZ
# TGFDPGoeDg7RJqixZWZB1UaVMYAKj1vTaoey652p65N0E3DIGBSbFcYqUFpD32Ag
# FBbJl6eIk7PClpRGdDPdlnzm5Bgm/rZKNY7BsP0+9F49pV4h4q1/nSiqwng41wCc
# ukDEuHOOXgusN9T97XwziytVkgUYMBZlvgklJEJm6MuBImXvAks6yqFImo9Gyam9
# LOZ4fDd/n0ajlhjJcA8yk+FpJqBMrbYDxYIhw60MhNsls+kQwTzYCDBWAFaLdvB1
# OtXdx3v2vitKky2ZJ9Hg+/h2t+EdCVADbN3PyIDqvW4mvqQU5DjHQ6HI5lMSdP99
# kxJm9Dx6u2j4rCAZnwGYiY8RahUDsxt/iNM0fVFbqufhbvbKLQWQfuUFgY7mCKYg
# fRO5t+DK3Q==
# SIG # End signature block
