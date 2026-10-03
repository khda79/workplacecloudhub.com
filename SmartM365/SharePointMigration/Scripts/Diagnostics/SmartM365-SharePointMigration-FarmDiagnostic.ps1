<#
.SYNOPSIS
    Collect read-only SharePoint farm diagnostics and correlate them with ShareGate peaks.
.VERSION
    1.0.1
#>
#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName='Peaks')]
param(
    [Parameter(Mandatory)][ValidatePattern('^(?!\.\.?$)[A-Za-z0-9][A-Za-z0-9 ._-]*$')][string]$Project,
    [Parameter(Mandatory,ParameterSetName='Range')][datetime]$StartTime,
    [Parameter(Mandatory,ParameterSetName='Range')][datetime]$EndTime,
    [Parameter(Mandatory,ParameterSetName='Around')][datetime]$Around,
    [Parameter(ParameterSetName='Around')][Parameter(ParameterSetName='Peaks')][ValidateRange(1,1440)][int]$WindowMinutes = 30,
    [Parameter(ParameterSetName='Range')][Parameter(ParameterSetName='Around')][Parameter(Mandatory,ParameterSetName='Peaks')][string]$ShareGatePeaksCsv,
    [string]$TimeZone = '',
    [string[]]$Servers,
    [ValidateRange(0,60)][int]$NightlyScanDays = 7,
    [ValidateRange(1,100000)][int]$MaxUlsEvents = 10000,
    [string]$ToolkitRoot = '',
    [string]$OutputPath = '',
    [switch]$DryRun,
    [Parameter(Mandatory,ParameterSetName='Copy')][string]$CopyResultsFrom = ''
)

$ErrorActionPreference = 'Stop'
$script:FarmDiagVersion = '1.0.1'
$script:FarmDiagScriptPath = $PSCommandPath
$script:FarmDiagMode = $PSCmdlet.ParameterSetName
$script:FarmDiagToolkitRoot = if ($ToolkitRoot) { $ToolkitRoot } else { [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')) }
$script:FarmDiagDefaultRoot = Join-Path $script:FarmDiagToolkitRoot 'Migrations'
$script:FarmDiagCoverage = New-Object 'System.Collections.Generic.List[object]'
$script:FarmDiagLog = New-Object 'System.Collections.Generic.List[string]'

function Write-FarmDiagLog {
    param([string]$Message)
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    [void]$script:FarmDiagLog.Add($line)
    Microsoft.PowerShell.Utility\Write-Host $line
}

function Add-FarmDiagCoverage {
    param([string]$Server,[string]$Source,[string]$Status,[string]$Detail,[double]$DurationSeconds = 0)
    $script:FarmDiagCoverage.Add([pscustomobject]@{ Server=$Server; Source=$Source; Status=$Status; Detail=$Detail; DurationSeconds=[math]::Round($DurationSeconds,2) })
}

function ConvertTo-FarmDiagUtc {
    param([datetime]$Value,[TimeZoneInfo]$Zone)
    if ($Value.Kind -eq [DateTimeKind]::Utc) { return $Value }
    if ($Value.Kind -eq [DateTimeKind]::Local) { return $Value.ToUniversalTime() }
    $wall = [DateTime]::SpecifyKind($Value,[DateTimeKind]::Unspecified)
    if ($Zone.IsInvalidTime($wall) -or $Zone.IsAmbiguousTime($wall)) { throw "The farm-local time '$wall' is invalid or ambiguous in $($Zone.Id). Supply an unambiguous UTC value." }
    return [TimeZoneInfo]::ConvertTimeToUtc($wall,$Zone)
}

function ConvertTo-FarmDiagLocal {
    param([datetime]$Utc,[TimeZoneInfo]$Zone)
    return [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::SpecifyKind($Utc,[DateTimeKind]::Utc),$Zone)
}

function Get-FarmDiagWindows {
    param([string]$Mode,[datetime]$Start,[datetime]$End,[datetime]$Center,[int]$Minutes,[string]$PeaksPath,[TimeZoneInfo]$Zone)
    $windows = New-Object 'System.Collections.Generic.List[object]'
    if ($Mode -eq 'Range') {
        $from = ConvertTo-FarmDiagUtc $Start $Zone; $to = ConvertTo-FarmDiagUtc $End $Zone
        if ($to -le $from) { throw 'EndTime must be after StartTime.' }
        $windows.Add([pscustomobject]@{ StartUtc=$from; EndUtc=$to; PeakUtc=$null; ShareGateLines=0 })
    }
    elseif ($Mode -eq 'Around') {
        $utc = ConvertTo-FarmDiagUtc $Center $Zone
        $windows.Add([pscustomobject]@{ StartUtc=$utc.AddMinutes(-$Minutes); EndUtc=$utc.AddMinutes($Minutes); PeakUtc=$utc; ShareGateLines=0 })
    }
    if ($PeaksPath) {
        if (-not (Test-Path -LiteralPath $PeaksPath -PathType Leaf)) { throw "ShareGate peaks CSV is missing: $PeaksPath" }
        $peakCount = 0
        foreach ($row in @(Import-Csv -LiteralPath $PeaksPath)) {
            if (-not $row.WindowUtc -or $row.WindowUtc -eq '(unparsed timestamp)') { continue }
            $stamp = [datetime]::MinValue
            if (-not [datetime]::TryParseExact(([string]$row.WindowUtc -replace '\s+UTC$',''), 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$stamp)) { throw "Invalid WindowUtc in peaks CSV: $($row.WindowUtc)" }
            $utc = [DateTime]::SpecifyKind($stamp,[DateTimeKind]::Utc)
            $lines = 0; [void][int]::TryParse([string]$row.Lines,[ref]$lines)
            $windows.Add([pscustomobject]@{ StartUtc=$utc.AddMinutes(-$Minutes); EndUtc=$utc.AddMinutes(5+$Minutes); PeakUtc=$utc; ShareGateLines=$lines })
            $peakCount++
        }
        if (-not $peakCount) { throw 'The ShareGate peaks CSV has no valid five-minute UTC windows.' }
    }
    return @($windows | Sort-Object StartUtc,EndUtc -Unique)
}

function Get-FarmDiagServerTimeZone {
    param([string]$Server,[bool]$IsLocal)
    if ($IsLocal) { return [TimeZoneInfo]::Local.Id }
    $registry = Get-WmiObject -Namespace root\default -Class StdRegProv -ComputerName $Server -ErrorAction Stop
    $result = $registry.GetStringValue(2147483650,'SYSTEM\CurrentControlSet\Control\TimeZoneInformation','TimeZoneKeyName')
    if ($result.ReturnValue -ne 0 -or -not $result.sValue) { throw "Remote time zone registry lookup failed with code $($result.ReturnValue)." }
    return [string]$result.sValue
}

function Get-FarmDiagAuditStatus {
    param([string]$Server,[bool]$IsLocal)
    if ($IsLocal) { $lines = @(& auditpol.exe /get /category:* /r 2>$null) }
    else {
        try {
            Test-WSMan -ComputerName $Server -ErrorAction Stop | Out-Null
            $lines = @(Invoke-Command -ComputerName $Server -ScriptBlock { auditpol.exe /get /category:* /r } -ErrorAction Stop)
        }
        catch { return 'Unknown - remote audit policy unavailable without WinRM' }
    }
    if (-not $lines.Count) { return 'Unknown - auditpol returned no policy' }
    try { $policy = @($lines | ConvertFrom-Csv -ErrorAction Stop) }
    catch { return 'Unknown - auditpol output could not be parsed' }
    $found = 0; $failureEnabled = 0
    foreach ($row in $policy) {
        $values = @($row.PSObject.Properties | ForEach-Object { [string]$_.Value })
        if ((($values -join '|') -match '(?i)0CCE92(15|3F|42)-') -and $values.Count) {
            $found++
            $setting = 0
            if ([int]::TryParse($values[-1],[ref]$setting) -and ($setting -band 2)) { $failureEnabled++ }
        }
    }
    if (-not $found) { return 'Unknown - relevant audit subcategories were not found' }
    if ($failureEnabled) { return "Active - failure audit enabled in $failureEnabled relevant subcategory(s)" }
    return 'Inactive - failure audit disabled for relevant subcategories'
}

function Test-FarmDiagInWindow {
    param([datetime]$Utc,[object[]]$Windows)
    foreach ($window in $Windows) { if ($Utc -ge $window.StartUtc -and $Utc -lt $window.EndUtc) { return $true } }
    return $false
}

function ConvertTo-FarmDiagUncPath {
    param([string]$Server,[string]$Path)
    $expanded = $Path.Replace('%SystemDrive%','C:').Replace('%systemdrive%','C:').Replace('%SystemRoot%','C:\Windows').Replace('%windir%','C:\Windows')
    if ($expanded -notmatch '^([A-Za-z]):\\(.*)$') { throw "Cannot map remote path to an administrative share: $Path" }
    return ('\\{0}\{1}$\{2}' -f $Server,$Matches[1].ToLowerInvariant(),$Matches[2])
}

function Get-FarmDiagUlsDirectory {
    param([string]$Server,[string]$LogLocation,[bool]$IsLocal)
    if ($IsLocal) {
        $localPath = $LogLocation.Replace('%CommonProgramFiles%',$env:CommonProgramFiles).Replace('%SystemDrive%',$env:SystemDrive)
        return $localPath
    }
    $remotePath = $LogLocation
    if ($remotePath -match '(?i)%CommonProgramFiles%') {
        $registry = Get-WmiObject -Namespace root\default -Class StdRegProv -ComputerName $Server -ErrorAction Stop
        $value = $registry.GetStringValue(2147483650,'SOFTWARE\Microsoft\Windows\CurrentVersion','CommonFilesDir')
        if ($value.ReturnValue -ne 0 -or -not $value.sValue) { throw "Cannot resolve CommonProgramFiles on $Server from the remote registry." }
        $remotePath = [regex]::Replace($remotePath,'(?i)%CommonProgramFiles%',[string]$value.sValue)
    }
    return ConvertTo-FarmDiagUncPath $Server $remotePath
}

function Get-FarmDiagIisConfig {
    param([string]$Server,[bool]$IsLocal)
    if ($IsLocal) {
        Import-Module WebAdministration -ErrorAction Stop
        Get-Website -ErrorAction Stop | Out-Null
        $path = Join-Path $env:windir 'System32\inetsrv\config\applicationHost.config'
    }
    else { $path = '\\{0}\c$\Windows\System32\inetsrv\config\applicationHost.config' -f $Server }
    [xml]$xml = [IO.File]::ReadAllText($path)
    return $xml
}

function Get-FarmDiagIisInventory {
    param([xml]$Config,[string]$Server)
    $sites = New-Object 'System.Collections.Generic.List[object]'
    $pools = New-Object 'System.Collections.Generic.List[object]'
    $root = $Config.configuration.'system.applicationHost'
    foreach ($pool in @($root.applicationPools.add)) {
        $cpu = $pool.cpu; $recycle = $pool.recycling.periodicRestart; $process = $pool.processModel
        $limit = 0; if ($cpu.limit) { $limit = [int]$cpu.limit }
        $schedule = @($recycle.schedule.add | ForEach-Object { $_.value }) -join ';'
        $pools.Add([pscustomobject]@{ Server=$Server; Pool=$pool.name; CpuLimitPercent=($limit/1000); CpuAction=[string]$cpu.action; CpuResetInterval=[string]$cpu.resetInterval; RecycleTime=[string]$recycle.time; RecycleSchedule=$schedule; MemoryKb=[string]$recycle.memory; PrivateMemoryKb=[string]$recycle.privateMemory; IdleTimeout=[string]$process.idleTimeout })
    }
    foreach ($site in @($root.sites.site)) {
        $bindings = @($site.bindings.binding | ForEach-Object { '{0}:{1}' -f $_.protocol,$_.bindingInformation }) -join ';'
        $sitePool = @($site.application | Where-Object { $_.path -eq '/' } | Select-Object -First 1 | ForEach-Object applicationPool)
        $directory = [string]$site.logFile.directory
        if (-not $directory) { $directory = [string]$root.sites.siteDefaults.logFile.directory }
        $format = [string]$site.logFile.logFormat
        if (-not $format) { $format = [string]$root.sites.siteDefaults.logFile.logFormat }
        if (-not $format) { $format = 'W3C' }
        $sites.Add([pscustomobject]@{ Server=$Server; Site=$site.name; SiteId=[string]$site.id; Bindings=$bindings; Pool=($sitePool -join ''); LogDirectory=$directory; LogFormat=$format })
    }
    return [pscustomobject]@{ Sites=$sites.ToArray(); Pools=$pools.ToArray() }
}

function Get-FarmDiagIisRequests {
    param([string]$Server,[object[]]$Sites,[object[]]$Windows,[bool]$IsLocal,[TimeZoneInfo]$Zone)
    $rows = @{}; $excluded = 0; $filesRead = 0; $linesRead = 0
    $missing = New-Object 'System.Collections.Generic.List[string]'
    $watch = [Diagnostics.Stopwatch]::StartNew()
    foreach ($site in $Sites) {
        if (-not $site.LogDirectory -or -not $site.SiteId) { continue }
        if ($site.LogFormat -ne 'W3C') { $missing.Add("$($site.Site): IIS log format '$($site.LogFormat)' is not W3C."); continue }
        $root = if ($IsLocal) { $site.LogDirectory.Replace('%SystemDrive%','C:').Replace('%SystemRoot%','C:\Windows') } else { ConvertTo-FarmDiagUncPath $Server $site.LogDirectory }
        $folder = Join-Path $root ('W3SVC' + $site.SiteId)
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { $missing.Add($folder); continue }
        try { $files = @(Get-ChildItem -LiteralPath $folder -File -Filter '*.log' -ErrorAction Stop) }
        catch { $missing.Add($folder + ': ' + $_.Exception.Message); continue }
        if (-not $files.Count) { $missing.Add($folder + ': no W3C log files'); continue }
        foreach ($file in $files) {
            $filesRead++
            try { $reader = New-Object IO.StreamReader($file.FullName) }
            catch { $missing.Add($file.FullName + ': ' + $_.Exception.Message); continue }
            try {
                $fields = @(); $index = @{}; $sawFields = $false
                while (-not $reader.EndOfStream) {
                    $line = $reader.ReadLine(); $linesRead++
                    if ($line.StartsWith('#Fields:')) {
                        $fields = @($line.Substring(8).Trim() -split '\s+')
                        $sawFields = $true
                        $index = @{}; for($i=0;$i -lt $fields.Count;$i++){ $index[$fields[$i]]=$i }
                        continue
                    }
                    if ($line.StartsWith('#') -or -not $index.ContainsKey('date') -or -not $index.ContainsKey('time') -or -not $index.ContainsKey('cs-uri-stem')) { continue }
                    $parts = $line -split '\s+'
                    if ($parts.Count -lt $fields.Count) { continue }
                    $dateText = $parts[$index['date']] + ' ' + $parts[$index['time']]
                    $stamp = [datetime]::MinValue
                    if (-not [datetime]::TryParseExact($dateText,'yyyy-MM-dd HH:mm:ss',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$stamp)) { continue }
                    $utc = [DateTime]::SpecifyKind($stamp,[DateTimeKind]::Utc)
                    if (-not (Test-FarmDiagInWindow $utc $Windows)) { continue }
                    $uri = $parts[$index['cs-uri-stem']]
                    if ($uri -notmatch '(?i)(client\.svc|/_vti_bin/)') { continue }
                    $agent = if ($index.ContainsKey('cs(User-Agent)')) { $parts[$index['cs(User-Agent)']] } else { '' }
                    if ($agent -match '(?i)MS[+ ]Search[+ ]6\.0[+ ]Robot') { $excluded++; continue }
                    $account = if ($index.ContainsKey('cs-username')) { $parts[$index['cs-username']] } else { '' }
                    $ip = if ($index.ContainsKey('c-ip')) { $parts[$index['c-ip']] } else { '' }
                    $status = if ($index.ContainsKey('sc-status')) { $parts[$index['sc-status']] } else { '' }
                    $sub = if ($index.ContainsKey('sc-substatus')) { $parts[$index['sc-substatus']] } else { '' }
                    $win32 = if ($index.ContainsKey('sc-win32-status')) { $parts[$index['sc-win32-status']] } else { '' }
                    $bin = $utc.AddMinutes(-($utc.Minute % 5)).AddSeconds(-$utc.Second)
                    $key = @($site.SiteId,$bin.ToString('o'),$account,$ip,$status,$sub,$win32) -join [char]31
                    if (-not $rows.ContainsKey($key)) {
                        $meaning = switch ($win32) { '2148074254' { 'SEC_E_NO_CREDENTIALS' } '2148074252' { 'SEC_E_LOGON_DENIED' } default { '' } }
                        $rows[$key] = [pscustomobject]@{ Server=$Server; Site=$site.Site; SiteId=$site.SiteId; Pool=$site.Pool; WindowUtc=$bin.ToString('o'); WindowLocal=(ConvertTo-FarmDiagLocal $bin $Zone).ToString('yyyy-MM-dd HH:mm:ss'); Account=$account; ClientIp=$ip; Status=$status; SubStatus=$sub; Win32Status=$win32; Win32Meaning=$meaning; Count=0 }
                    }
                    $rows[$key].Count++
                }
                if (-not $sawFields) { $missing.Add($file.FullName + ': W3C #Fields header missing') }
            }
            finally { $reader.Dispose() }
        }
    }
    $watch.Stop()
    return [pscustomobject]@{ Rows=@($rows.Values); ExcludedRobot=$excluded; FilesRead=$filesRead; LinesRead=$linesRead; DurationSeconds=$watch.Elapsed.TotalSeconds; MissingPaths=$missing.ToArray() }
}

function Get-FarmDiagEvents {
    param([string]$Server,[string]$LogName,[int[]]$Ids,[object[]]$Windows,[TimeZoneInfo]$Zone,[int]$NightDays = 0)
    $from = @($Windows | ForEach-Object StartUtc | Sort-Object | Select-Object -First 1)[0]
    $to = @($Windows | ForEach-Object EndUtc | Sort-Object -Descending | Select-Object -First 1)[0]
    if ($NightDays -gt 0) { $from = $from.AddDays(-$NightDays) }
    $filter = @{ LogName=$LogName; Id=$Ids; StartTime=$from.ToLocalTime(); EndTime=$to.ToLocalTime() }
    try { $events = @(Get-WinEvent -ComputerName $Server -FilterHashtable $filter -ErrorAction Stop) }
    catch {
        if ($_.FullyQualifiedErrorId -match 'NoMatchingEventsFound' -or $_.Exception.Message -match 'No events were found') { $events = @() }
        else { throw }
    }
    if ($LogName -eq 'System') {
        foreach($window in $Windows) {
            $allWasFilter = @{ LogName='System'; StartTime=$window.StartUtc.ToLocalTime(); EndTime=$window.EndUtc.ToLocalTime() }
            try { $events += @(Get-WinEvent -ComputerName $Server -FilterHashtable $allWasFilter -ErrorAction Stop | Where-Object { [string]$_.ProviderName -match '(?i)WAS' }) }
            catch { if ($_.FullyQualifiedErrorId -notmatch 'NoMatchingEventsFound' -and $_.Exception.Message -notmatch 'No events were found') { throw } }
        }
    }
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $seen = @{}
    foreach ($event in $events) {
        if ($LogName -eq 'System' -and [string]$event.ProviderName -notmatch '(?i)WAS') { continue }
        $key = [string]$event.RecordId
        if ($key -and $seen.ContainsKey($key)) { continue }
        if ($key) { $seen[$key] = $true }
        $utcText = [regex]::Match($event.ToXml(),'<TimeCreated\s+SystemTime="([^"]+)"').Groups[1].Value
        $utc = if ($utcText) { [datetime]::Parse($utcText,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AdjustToUniversal) } else { $event.TimeCreated.ToUniversalTime() }
        $utc = [DateTime]::SpecifyKind($utc,[DateTimeKind]::Utc)
        if ($NightDays -eq 0 -and -not (Test-FarmDiagInWindow $utc $Windows)) { continue }
        $rows.Add([pscustomobject]@{ Server=$Server; Log=$LogName; RecordId=$event.RecordId; EventId=$event.Id; Provider=$event.ProviderName; Utc=$utc.ToString('o'); Local=(ConvertTo-FarmDiagLocal $utc $Zone).ToString('yyyy-MM-dd HH:mm:ss'); InPeakWindow=(Test-FarmDiagInWindow $utc $Windows); Message=([string]$event.Message -replace '\r?\n',' ') })
    }
    return $rows.ToArray()
}

function Get-FarmDiagNightlyRecurrences {
    param([object[]]$Events,[int]$Days,[object[]]$Windows,[TimeZoneInfo]$Zone)
    $groups = @{}
    foreach ($event in $Events) {
        $clock = [datetime]::ParseExact($event.Local,'yyyy-MM-dd HH:mm:ss',[Globalization.CultureInfo]::InvariantCulture)
        $inClockWindow = $false
        foreach($window in $Windows) {
            $localStart = ConvertTo-FarmDiagLocal $window.StartUtc $Zone
            $localEnd = ConvertTo-FarmDiagLocal $window.EndUtc $Zone
            $startClock = $localStart.TimeOfDay; $endClock = $localEnd.TimeOfDay
            if (($localEnd-$localStart).TotalHours -ge 24 -or ($startClock -le $endClock -and $clock.TimeOfDay -ge $startClock -and $clock.TimeOfDay -le $endClock) -or ($startClock -gt $endClock -and ($clock.TimeOfDay -ge $startClock -or $clock.TimeOfDay -le $endClock))) { $inClockWindow = $true; break }
        }
        if (-not $inClockWindow) { continue }
        $bucket = '{0:D2}:{1:D2}' -f $clock.Hour,([int]([math]::Floor($clock.Minute/5)*5))
        $key = @($event.Server,$event.EventId,$bucket) -join [char]31
        if (-not $groups.ContainsKey($key)) { $groups[$key] = [pscustomobject]@{ Server=$event.Server; EventId=$event.EventId; LocalFiveMinuteBucket=$bucket; Nights=@{}; Events=0 } }
        $groups[$key].Nights[$clock.ToString('yyyy-MM-dd')] = $true
        $groups[$key].Events++
    }
    return @($groups.Values | Where-Object { $_.Nights.Count -ge 2 } | ForEach-Object {
        [pscustomobject]@{ Server=$_.Server; EventId=$_.EventId; LocalFiveMinuteBucket=$_.LocalFiveMinuteBucket; Nights=$_.Nights.Count; Events=$_.Events; ScanDays=$Days }
    } | Sort-Object Server,EventId,LocalFiveMinuteBucket)
}

function Get-FarmDiagCorrelation {
    param([object[]]$Windows,[object[]]$Was,[object[]]$Security,[object[]]$Iis,[object[]]$Uls,[TimeZoneInfo]$Zone = [TimeZoneInfo]::Local)
    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach ($window in $Windows) {
        $start = $window.StartUtc; $end = $window.EndUtc
        $wasRows = @($Was | Where-Object { $stamp=[DateTimeOffset]::Parse([string]$_.Utc).UtcDateTime; $stamp -ge $start -and $stamp -lt $end })
        $securityRows = @($Security | Where-Object { $stamp=[DateTimeOffset]::Parse([string]$_.Utc).UtcDateTime; $stamp -ge $start -and $stamp -lt $end })
        $iisRows = @($Iis | Where-Object { $stamp=[DateTimeOffset]::Parse([string]$_.WindowUtc).UtcDateTime; $stamp -ge $start -and $stamp -lt $end })
        $ulsRows = @($Uls | Where-Object { $stamp=[DateTimeOffset]::Parse([string]$_.Utc).UtcDateTime; $stamp -ge $start -and $stamp -lt $end })
        $iis401 = @($iisRows | Where-Object Status -EQ '401' | Measure-Object -Property Count -Sum | ForEach-Object Sum)
        $result.Add([pscustomobject]@{ StartUtc=$start.ToString('o'); StartLocal=(ConvertTo-FarmDiagLocal $start $Zone).ToString('yyyy-MM-dd HH:mm:ss'); EndUtc=$end.ToString('o'); EndLocal=(ConvertTo-FarmDiagLocal $end $Zone).ToString('yyyy-MM-dd HH:mm:ss'); TimeZone=$Zone.Id; PeakUtc=$(if($window.PeakUtc){$window.PeakUtc.ToString('o')}else{''}); ShareGateLines=$window.ShareGateLines; WasEvents=$wasRows.Count; SecurityEvents=$securityRows.Count; Iis401Count=$(if($iis401.Count -and $null -ne $iis401[0]){$iis401[0]}else{0}); UlsEvents=$ulsRows.Count; Servers=(@($wasRows+$securityRows+$iisRows+$ulsRows | ForEach-Object Server | Sort-Object -Unique) -join ';') })
    }
    return $result.ToArray()
}

function Get-FarmDiagTimeline {
    param([object[]]$Was,[object[]]$Security,[object[]]$Iis,[object[]]$Uls)
    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach($row in $Was) {
        if (-not $row.InPeakWindow) { continue }
        $result.Add([pscustomobject]@{ Utc=$row.Utc; Local=$row.Local; Server=$row.Server; Source='WAS'; Event=[string]$row.EventId; Detail=$row.Message; Count=1 })
    }
    foreach($row in $Security) {
        $result.Add([pscustomobject]@{ Utc=$row.Utc; Local=$row.Local; Server=$row.Server; Source='Security'; Event=[string]$row.EventId; Detail=$row.Message; Count=1 })
    }
    foreach($row in $Iis) {
        $result.Add([pscustomobject]@{ Utc=$row.WindowUtc; Local=$row.WindowLocal; Server=$row.Server; Source='IIS'; Event=($row.Status+'.'+$row.SubStatus); Detail=($row.Site+' | '+$row.Pool+' | '+$row.Win32Meaning+' | '+$row.Account+' | '+$row.ClientIp); Count=$row.Count })
    }
    foreach($row in $Uls) {
        $result.Add([pscustomobject]@{ Utc=$row.Utc; Local=$row.Local; Server=$row.Server; Source='ULS'; Event=$row.Level; Detail=($row.Category+' | '+$row.Message); Count=1 })
    }
    return @($result.ToArray() | Sort-Object Utc,Server,Source)
}

function Get-FarmDiagPoolRisks {
    param([object[]]$Pools,[object[]]$Windows,[TimeZoneInfo]$Zone)
    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach($pool in $Pools) {
        if ([double]$pool.CpuLimitPercent -gt 0 -and $pool.CpuAction -eq 'Throttle') {
            $result.Add([pscustomobject]@{ Server=$pool.Server; Pool=$pool.Pool; Risk='CpuThrottle'; WindowStartUtc=''; Detail=("CPU limit $($pool.CpuLimitPercent)% with Throttle; review ThrottleUnderLoad suitability.") })
        }
        foreach($schedule in @($pool.RecycleSchedule -split ';' | Where-Object { $_ })) {
            $span = [TimeSpan]::Zero
            if (-not [TimeSpan]::TryParse($schedule,[ref]$span)) { continue }
            foreach($window in $Windows) {
                $start = ConvertTo-FarmDiagLocal $window.StartUtc $Zone
                $end = ConvertTo-FarmDiagLocal $window.EndUtc $Zone
                foreach($offset in @(-1,0,1)) {
                    $recycle = $start.Date.AddDays($offset).Add($span)
                    if ($recycle -ge $start.AddMinutes(-15) -and $recycle -le $end.AddMinutes(15)) {
                        $result.Add([pscustomobject]@{ Server=$pool.Server; Pool=$pool.Pool; Risk='ScheduledRecycleNearWindow'; WindowStartUtc=$window.StartUtc.ToString('o'); Detail=("Scheduled recycling at local $schedule overlaps or is within 15 minutes of the window.") })
                        break
                    }
                }
            }
        }
    }
    return $result.ToArray()
}

function Get-FarmDiagUls {
    param([string]$Server,[string]$LogLocation,[object[]]$Windows,[TimeZoneInfo]$Zone,[int]$MaxEvents,[bool]$IsLocal)
    if (-not $LogLocation) { throw 'Farm diagnostic configuration did not provide LogLocation.' }
    $directory = Get-FarmDiagUlsDirectory $Server $LogLocation $IsLocal
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { throw "ULS directory is inaccessible: $directory" }
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $seen = @{}
    $truncated = $false
    $ambiguous = 0
    foreach ($window in $Windows) {
        $localStart = ConvertTo-FarmDiagLocal $window.StartUtc $Zone
        $localEnd = ConvertTo-FarmDiagLocal $window.EndUtc $Zone
        $candidates = @(Get-SPLogEvent -Directory $directory -StartTime $localStart -EndTime $localEnd -ErrorAction Stop | Where-Object {
            [string]$_.Level -match '(?i)High|Critical|Unexpected' -or ([string]$_.Message + ' ' + [string]$_.Category) -match '(?i)authenticat|w3wp|config.?cache|process.*start'
        } | Select-Object -First ($MaxEvents + 1))
        foreach ($entry in $candidates) {
            $level = [string]$entry.Level; $message = [string]$entry.Message; $category = [string]$entry.Category
            $localStamp = [datetime]$entry.Timestamp
            try { $utc = ConvertTo-FarmDiagUtc ([DateTime]::SpecifyKind($localStamp,[DateTimeKind]::Unspecified)) $Zone }
            catch { $ambiguous++; continue }
            $key = @($utc.ToString('o'),$level,$category,$message,[string]$entry.Process) -join [char]31
            if ($seen.ContainsKey($key)) { continue }
            if ($rows.Count -ge $MaxEvents) { $truncated = $true; break }
            $seen[$key] = $true
            $rows.Add([pscustomobject]@{ Server=$Server; Utc=$utc.ToString('o'); Local=$localStamp.ToString('yyyy-MM-dd HH:mm:ss'); Level=$level; Category=$category; Process=[string]$entry.Process; Message=($message -replace '\r?\n',' ') })
        }
        if ($truncated) { break }
    }
    return [pscustomobject]@{ Rows=$rows.ToArray(); Truncated=$truncated; AmbiguousTimestamps=$ambiguous; Directory=$directory }
}

function Write-FarmDiagCsv {
    param([string]$Directory,[string]$Name,[string[]]$Columns,[object[]]$Rows)
    $path = Join-Path $Directory $Name
    $temp = $path + '.tmp-' + [guid]::NewGuid().ToString('N')
    if ($Rows.Count) { $Rows | Select-Object -Property $Columns | Export-Csv -LiteralPath $temp -NoTypeInformation -Encoding UTF8 }
    else { [IO.File]::WriteAllText($temp,('"' + ($Columns -join '","') + '"' + [Environment]::NewLine),[Text.UTF8Encoding]::new($true)) }
    Move-Item -LiteralPath $temp -Destination $path -Force
    return $path
}

function Write-FarmDiagText {
    param([string]$Path,[string]$Value)
    $temp = $Path + '.tmp-' + [guid]::NewGuid().ToString('N')
    [IO.File]::WriteAllText($temp,$Value,[Text.UTF8Encoding]::new($true))
    Move-Item -LiteralPath $temp -Destination $Path -Force
}

function Get-FarmDiagOutput {
    param([string]$ProjectName,[string]$Override,[bool]$Preview)
    $root = if ($Override) { $Override } else { Join-Path (Join-Path $script:FarmDiagDefaultRoot $ProjectName) 'ShareGate\Diagnostics' }
    $name = 'Farm-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
    try {
        if (-not (Test-Path -LiteralPath $root)) { [void](New-Item -ItemType Directory -Path $root -Force) }
        $probe = Join-Path $root ('.farm-write-probe-' + [guid]::NewGuid().ToString('N'))
        [IO.File]::WriteAllText($probe,'probe'); Remove-Item -LiteralPath $probe -Force
        return [pscustomobject]@{ Path=(Join-Path $root $name); Fallback=$false; DesiredRoot=$root }
    }
    catch {
        $fallback = Join-Path 'C:\Temp\SPFarmDiag' (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
        Write-FarmDiagLog "Output share is not writable: $root. Local fallback: $fallback. Error: $($_.Exception.Message)"
        if (-not $Preview) { [void](New-Item -ItemType Directory -Path $fallback -Force) }
        return [pscustomobject]@{ Path=$fallback; Fallback=$true; DesiredRoot=$root }
    }
}

function Get-FarmDiagRerunCommand {
    param([string]$Server,[string]$ZoneId,[bool]$Preview)
    $scriptPath = Join-Path $script:FarmDiagToolkitRoot 'Scripts\Diagnostics\SmartM365-SharePointMigration-FarmDiagnostic.ps1'
    $command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -Project "{1}" -Servers "{2}" -TimeZone "{3}" -NightlyScanDays {4} -MaxUlsEvents {5} -ToolkitRoot "{6}"' -f $scriptPath,$Project,$Server,$ZoneId,$NightlyScanDays,$MaxUlsEvents,$script:FarmDiagToolkitRoot
    if ($script:FarmDiagMode -eq 'Range') { $command += ' -StartTime "{0:o}" -EndTime "{1:o}"' -f $StartTime,$EndTime }
    elseif ($script:FarmDiagMode -eq 'Around') { $command += ' -Around "{0:o}" -WindowMinutes {1}' -f $Around,$WindowMinutes }
    else { $command += ' -WindowMinutes {0}' -f $WindowMinutes }
    if ($ShareGatePeaksCsv) { $command += ' -ShareGatePeaksCsv "{0}"' -f $ShareGatePeaksCsv }
    if ($Preview) { $command += ' -DryRun' }
    return $command
}

function Write-FarmDiagHtml {
    param([string]$Path,[object]$Summary,[object[]]$Pools,[object[]]$Events,[object[]]$IisRows,[object[]]$Correlation,[object[]]$Recurrences,[object[]]$Timeline,[object[]]$Risks)
    $risk = @($Pools | Where-Object { $_.CpuLimitPercent -gt 0 -and $_.CpuAction -eq 'Throttle' })
    $html = New-Object 'System.Collections.Generic.List[string]'
    $html.Add('<!doctype html><html><head><meta charset="utf-8"><title>SharePoint farm diagnostics</title><style>body{font:14px Segoe UI,Arial;margin:30px;color:#243143}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccd5df;padding:6px;text-align:left}th{background:#e9f1fa}h1,h2{color:#164c75}.warn{color:#9b3e20}</style></head><body>')
    $html.Add('<h1>SharePoint farm diagnostics</h1>')
    $html.Add('<p>Project: '+[Net.WebUtility]::HtmlEncode($Summary.Project)+' | Status: '+[Net.WebUtility]::HtmlEncode($Summary.Status)+' | Generated UTC: '+[Net.WebUtility]::HtmlEncode($Summary.GeneratedAtUtc)+'</p>')
    $html.Add('<h2>Coverage</h2><table><tr><th>Server</th><th>Source</th><th>Status</th><th>Detail</th><th>Read duration (s)</th></tr>')
    foreach ($row in $Summary.Coverage) { $html.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f [Net.WebUtility]::HtmlEncode([string]$row.Server),[Net.WebUtility]::HtmlEncode([string]$row.Source),[Net.WebUtility]::HtmlEncode([string]$row.Status),[Net.WebUtility]::HtmlEncode([string]$row.Detail),$row.DurationSeconds)) }
    $html.Add('</table><h2>CPU throttle configuration</h2>')
    if ($risk.Count) { foreach ($pool in $risk) { $html.Add('<p class="warn">'+[Net.WebUtility]::HtmlEncode(('{0}: {1} limits CPU to {2}% with Throttle.' -f $pool.Server,$pool.Pool,$pool.CpuLimitPercent))+'</p>') } }
    else { $html.Add('<p>No Throttle pool was observed in the collected configuration.</p>') }
    $html.Add('<h2>Correlation</h2><p>ShareGate peaks: '+$Summary.PeakCount+'; WAS events: '+$Events.Count+'; IIS groups: '+$IisRows.Count+'. Temporal proximity is evidence for review, not proof of causation.</p>')
    $html.Add('<table><tr><th>Start UTC</th><th>Start local</th><th>End UTC</th><th>End local</th><th>ShareGate lines</th><th>WAS</th><th>Security</th><th>IIS 401</th><th>ULS</th><th>Servers</th></tr>')
    foreach($row in $Correlation){$html.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td></tr>' -f $row.StartUtc,$row.StartLocal,$row.EndUtc,$row.EndLocal,$row.ShareGateLines,$row.WasEvents,$row.SecurityEvents,$row.Iis401Count,$row.UlsEvents,[Net.WebUtility]::HtmlEncode($row.Servers)))}
    $html.Add('</table><h2>Recurring WAS events</h2><table><tr><th>Server</th><th>Event ID</th><th>Local five-minute bucket</th><th>Nights</th><th>Events</th></tr>')
    foreach($row in $Recurrences){$html.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f [Net.WebUtility]::HtmlEncode($row.Server),$row.EventId,$row.LocalFiveMinuteBucket,$row.Nights,$row.Events))}
    $html.Add('</table>')
    $html.Add('<h2>Window timeline</h2><p>The table shows the first 100 rows. Farm-Timeline.csv contains the full timeline.</p><table><tr><th>UTC</th><th>Local</th><th>Server</th><th>Source</th><th>Event</th><th>Count</th><th>Detail</th></tr>')
    foreach($row in @($Timeline | Select-Object -First 100)){$html.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>' -f $row.Utc,$row.Local,[Net.WebUtility]::HtmlEncode($row.Server),$row.Source,[Net.WebUtility]::HtmlEncode($row.Event),$row.Count,[Net.WebUtility]::HtmlEncode($row.Detail)))}
    $html.Add('</table><h2>Configuration and window risks</h2><table><tr><th>Server</th><th>Pool</th><th>Risk</th><th>Detail</th></tr>')
    foreach($row in $Risks){$html.Add(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>' -f [Net.WebUtility]::HtmlEncode($row.Server),[Net.WebUtility]::HtmlEncode($row.Pool),$row.Risk,[Net.WebUtility]::HtmlEncode($row.Detail)))}
    $html.Add('</table>')
    $html.Add('<h2>Recommendations for operator review</h2><ul><li>Review migration windows overlapping recurring WAS events or planned recycling.</li><li>Review whether ThrottleUnderLoad is appropriate for pools with sustained Throttle events. No setting is applied by this diagnostic.</li><li>Check source authentication and worker-process restarts before attributing IIS 401 responses to a migration failure.</li></ul>')
    $html.Add('</body></html>')
    Write-FarmDiagText $Path ($html -join [Environment]::NewLine)
}

function Invoke-FarmDiagnostic {
    Write-FarmDiagLog "Farm diagnostic v$script:FarmDiagVersion started."
    if ($CopyResultsFrom) {
        $target = Join-Path (Join-Path $script:FarmDiagDefaultRoot $Project) 'ShareGate\Diagnostics'
        if (-not (Test-Path -LiteralPath $CopyResultsFrom -PathType Container)) { throw "Results folder is missing: $CopyResultsFrom" }
        [void](New-Item -ItemType Directory -Path $target -Force)
        Copy-Item -LiteralPath $CopyResultsFrom -Destination $target -Recurse -ErrorAction Stop
        Write-FarmDiagLog "Copied output to $target"
        return
    }
    if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -ne 1) { throw 'Windows PowerShell 5.1 is required; PowerShell 7 is not supported.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run Windows PowerShell as Administrator on a farm server.' }
    try { $Host.Runspace.ThreadOptions = 'ReuseThread' } catch { Write-FarmDiagLog 'ReuseThread could not be enabled; continuing.' }
    if (-not (Get-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
        try { Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop }
        catch { try { Import-Module SharePointServer -ErrorAction Stop } catch { throw 'Cannot load Microsoft.SharePoint.PowerShell snap-in or SharePointServer module.' } }
    }
    try { $null = Get-SPFarm -ErrorAction Stop } catch { throw "Cannot access the SharePoint farm. The account must be a SharePoint Shell Admin. Original error: $($_.Exception.Message)" }
    $farmServers = @(Get-SPServer -ErrorAction Stop)
    $instances = @(Get-SPServiceInstance -ErrorAction Stop)
    $diagnosticConfig = Get-SPDiagnosticConfig -ErrorAction Stop
    $zone = if ($TimeZone) { [TimeZoneInfo]::FindSystemTimeZoneById($TimeZone) } else { [TimeZoneInfo]::Local }
    $windows = @(Get-FarmDiagWindows -Mode $script:FarmDiagMode -Start $StartTime -End $EndTime -Center $Around -Minutes $WindowMinutes -PeaksPath $ShareGatePeaksCsv -Zone $zone)
    $selected = @($farmServers | Where-Object { [string]$_.Type -notmatch '(?i)database|sql' })
    if ($Servers) {
        $unknown = @($Servers | Where-Object { $_ -notin @($selected | ForEach-Object Address) })
        if ($unknown.Count) { throw "Servers are not members of the discovered SharePoint farm: $($unknown -join ', ')" }
        $selected = @($selected | Where-Object { $_.Address -in $Servers })
    }
    if (-not $selected.Count) { throw 'No SharePoint application server was discovered.' }
    $output = Get-FarmDiagOutput -ProjectName $Project -Override $OutputPath -Preview ([bool]$DryRun)
    Write-FarmDiagLog "Output: $($output.Path)"
    foreach ($window in $windows) { Write-FarmDiagLog ("Window UTC: {0:o} to {1:o}; ShareGate lines: {2}" -f $window.StartUtc,$window.EndUtc,$window.ShareGateLines) }
    foreach ($server in $selected) { Write-FarmDiagLog "Discovered server: $($server.Address); role=$($server.Role)" }
    if ($DryRun) {
        foreach($server in $selected) {
            $name = [string]$server.Address
            $local = $name -ieq $env:COMPUTERNAME -or $name.Split('.')[0] -ieq $env:COMPUTERNAME
            $configPath = if($local){Join-Path $env:windir 'System32\inetsrv\config\applicationHost.config'}else{'\\{0}\c$\Windows\System32\inetsrv\config\applicationHost.config' -f $name}
            try { Write-FarmDiagLog "Preflight $name IIS config readable: $(Test-Path -LiteralPath $configPath -PathType Leaf)" }
            catch { Write-FarmDiagLog "Preflight $name IIS config: unavailable ($($_.Exception.Message))" }
            try { $null = Get-WinEvent -ComputerName $name -ListLog System -ErrorAction Stop; Write-FarmDiagLog "Preflight $name event RPC: available" }
            catch { Write-FarmDiagLog "Preflight $name event RPC: unavailable ($($_.Exception.Message))" }
            try {
                $ulsPath = Get-FarmDiagUlsDirectory $name ([string]$diagnosticConfig.LogLocation) $local
                Write-FarmDiagLog "Preflight $name ULS directory readable: $(Test-Path -LiteralPath $ulsPath -PathType Container)"
            }
            catch { Write-FarmDiagLog "Preflight $name ULS directory: unavailable ($($_.Exception.Message))" }
        }
        Write-FarmDiagLog 'DryRun completed. No farm logs or IIS log files were collected.'
        return
    }
    [void](New-Item -ItemType Directory -Path $output.Path -Force)
    $topology = New-Object 'System.Collections.Generic.List[object]'
    $sites = New-Object 'System.Collections.Generic.List[object]'
    $pools = New-Object 'System.Collections.Generic.List[object]'
    $was = New-Object 'System.Collections.Generic.List[object]'
    $security = New-Object 'System.Collections.Generic.List[object]'
    $iis = New-Object 'System.Collections.Generic.List[object]'
    $uls = New-Object 'System.Collections.Generic.List[object]'
    foreach ($server in $selected) {
        $name = [string]$server.Address; $local = $name -ieq $env:COMPUTERNAME -or $name.Split('.')[0] -ieq $env:COMPUTERNAME
        $webService = @($instances | Where-Object { [string]$_.Server.Address -ieq $name -and [string]$_.TypeName -match '(?i)Microsoft SharePoint Foundation Web Application' -and [string]$_.Status -eq 'Online' }).Count -gt 0
        $serverZone = ''
        try { $serverZone = Get-FarmDiagServerTimeZone $name $local; Add-FarmDiagCoverage $name 'Time zone' 'Complete' $serverZone }
        catch { Add-FarmDiagCoverage $name 'Time zone' 'Missing' $_.Exception.Message }
        $topology.Add([pscustomobject]@{ Server=$name; Role=[string]$server.Role; Type=[string]$server.Type; EffectiveWfe=$webService; FarmTimeZone=$zone.Id; ServerTimeZone=$serverZone; IsLocal=$local })
        if ($serverZone -and $serverZone -ne $zone.Id) { Write-FarmDiagLog "Time zone mismatch on $name`: server=$serverZone; farm=$($zone.Id)." }
        try {
                $config = Get-FarmDiagIisConfig $name $local
                $inventory = Get-FarmDiagIisInventory $config $name
                foreach($row in $inventory.Sites){ $sites.Add($row) }
                foreach($row in $inventory.Pools){ $pools.Add($row) }
                Add-FarmDiagCoverage $name 'IIS configuration' 'Complete' ('Sites='+$inventory.Sites.Count)
                try {
                    $result = Get-FarmDiagIisRequests $name $inventory.Sites $windows $local $zone
                    foreach($row in $result.Rows){ $iis.Add($row) }
                    $iisStatus = if($result.MissingPaths.Count -and -not $result.FilesRead){'Missing'}elseif($result.MissingPaths.Count){'Partial'}else{'Complete'}
                    Add-FarmDiagCoverage $name 'IIS logs' $iisStatus ("Files=$($result.FilesRead); Lines=$($result.LinesRead); RobotExcluded=$($result.ExcludedRobot); Missing=$($result.MissingPaths -join ';')") $result.DurationSeconds
                }
                catch { Add-FarmDiagCoverage $name 'IIS logs' 'Missing' $_.Exception.Message }
            }
        catch { Add-FarmDiagCoverage $name 'IIS configuration' 'Missing' $_.Exception.Message; Add-FarmDiagCoverage $name 'IIS logs' 'Missing' 'IIS configuration was unavailable.' }
        try {
            $rows = @(Get-FarmDiagEvents $name 'System' @(5210,5074,5075,5011,5009,5002,5186,5117) $windows $zone $NightlyScanDays)
            foreach($row in $rows){ $was.Add($row) }
            Add-FarmDiagCoverage $name 'WAS events' 'Complete' ("Events=$($rows.Count)")
        }
        catch { Add-FarmDiagCoverage $name 'WAS events' 'Missing' $_.Exception.Message }
        try {
            $audit = Get-FarmDiagAuditStatus $name $local
            $rows = @(Get-FarmDiagEvents $name 'Security' @(4625,4771,4776) $windows $zone)
            foreach($row in $rows){ $security.Add($row) }
            $auditState = if($audit -like 'Inactive*'){'NotApplicable'}elseif($audit -like 'Unknown*' -and -not $rows.Count){'Partial'}else{'Complete'}
            Add-FarmDiagCoverage $name 'Security events' $auditState ("Events=$($rows.Count); Audit=$audit")
        }
        catch { Add-FarmDiagCoverage $name 'Security events' 'Missing' $_.Exception.Message }
        try {
            $result = Get-FarmDiagUls $name ([string]$diagnosticConfig.LogLocation) $windows $zone $MaxUlsEvents $local
            foreach($row in $result.Rows){ $uls.Add($row) }
            Add-FarmDiagCoverage $name 'ULS' $(if($result.Truncated -or $result.AmbiguousTimestamps){'Partial'}else{'Complete'}) ("Events=$($result.Rows.Count); AmbiguousTimestamps=$($result.AmbiguousTimestamps); Directory=$($result.Directory)")
        }
        catch { Add-FarmDiagCoverage $name 'ULS' 'Missing' $_.Exception.Message }
        if (@($script:FarmDiagCoverage | Where-Object { $_.Server -eq $name -and $_.Status -eq 'Missing' }).Count -and -not $local) {
            Add-FarmDiagCoverage $name 'Local rerun' 'Action' ((Get-FarmDiagRerunCommand $name $zone.Id $true) + ' | ' + (Get-FarmDiagRerunCommand $name $zone.Id $false))
            Write-FarmDiagLog "Local DryRun on $name`: $(Get-FarmDiagRerunCommand $name $zone.Id $true)"
            Write-FarmDiagLog "Local run on $name`: $(Get-FarmDiagRerunCommand $name $zone.Id $false)"
        }
    }
    $prefix = 'Farm-'
    Write-FarmDiagCsv $output.Path ($prefix+'Topology.csv') @('Server','Role','Type','EffectiveWfe','FarmTimeZone','ServerTimeZone','IsLocal') $topology.ToArray() | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'IisSites.csv') @('Server','Site','SiteId','Bindings','Pool','LogDirectory','LogFormat') $sites.ToArray() | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'IisPools.csv') @('Server','Pool','CpuLimitPercent','CpuAction','CpuResetInterval','RecycleTime','RecycleSchedule','MemoryKb','PrivateMemoryKb','IdleTimeout') $pools.ToArray() | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'WasEvents.csv') @('Server','Log','RecordId','EventId','Provider','Utc','Local','InPeakWindow','Message') $was.ToArray() | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'SecurityEvents.csv') @('Server','Log','RecordId','EventId','Provider','Utc','Local','InPeakWindow','Message') $security.ToArray() | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'IisRequests.csv') @('Server','Site','SiteId','Pool','WindowUtc','WindowLocal','Account','ClientIp','Status','SubStatus','Win32Status','Win32Meaning','Count') $iis.ToArray() | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'UlsEvents.csv') @('Server','Utc','Local','Level','Category','Process','Message') $uls.ToArray() | Out-Null
    $correlation = @(Get-FarmDiagCorrelation $windows $was.ToArray() $security.ToArray() $iis.ToArray() $uls.ToArray() $zone)
    $recurrences = if($NightlyScanDays){@(Get-FarmDiagNightlyRecurrences $was.ToArray() $NightlyScanDays $windows $zone)}else{@()}
    $timeline = @(Get-FarmDiagTimeline $was.ToArray() $security.ToArray() $iis.ToArray() $uls.ToArray())
    $risks = @(Get-FarmDiagPoolRisks $pools.ToArray() $windows $zone)
    Write-FarmDiagCsv $output.Path ($prefix+'Correlation.csv') @('StartUtc','StartLocal','EndUtc','EndLocal','TimeZone','PeakUtc','ShareGateLines','WasEvents','SecurityEvents','Iis401Count','UlsEvents','Servers') $correlation | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'NightlyRecurrences.csv') @('Server','EventId','LocalFiveMinuteBucket','Nights','Events','ScanDays') $recurrences | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'Timeline.csv') @('Utc','Local','Server','Source','Event','Detail','Count') $timeline | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'Risks.csv') @('Server','Pool','Risk','WindowStartUtc','Detail') $risks | Out-Null
    Write-FarmDiagCsv $output.Path ($prefix+'Coverage.csv') @('Server','Source','Status','Detail','DurationSeconds') $script:FarmDiagCoverage.ToArray() | Out-Null
    $status = if(@($script:FarmDiagCoverage | Where-Object { $_.Status -in @('Missing','Partial') }).Count){'Partial'}else{'Complete'}
    $summary = [pscustomobject]@{ SchemaVersion=1; Project=$Project; Status=$status; GeneratedAtUtc=[datetime]::UtcNow.ToString('o'); TimeZone=$zone.Id; Windows=$windows; PeakCount=$windows.Count; Coverage=$script:FarmDiagCoverage.ToArray(); Correlation=$correlation; Recurrences=$recurrences; OutputDirectory=$output.Path; ReportPath=(Join-Path $output.Path 'Farm-Report.html') }
    Write-FarmDiagHtml $summary.ReportPath $summary $pools.ToArray() $was.ToArray() $iis.ToArray() $correlation $recurrences $timeline $risks
    if ($output.Fallback) {
        Write-FarmDiagLog ('Copy command: powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -Project "{1}" -CopyResultsFrom "{2}"' -f $script:FarmDiagScriptPath,$Project,$output.Path)
    }
    Write-FarmDiagLog "Final output: $($output.Path)"
    Write-FarmDiagText (Join-Path $output.Path 'Farm.log') (($script:FarmDiagLog -join [Environment]::NewLine)+[Environment]::NewLine)
    Write-FarmDiagText (Join-Path $output.Path 'Farm-Summary.json.txt') (($summary | ConvertTo-Json -Depth 8)+[Environment]::NewLine)
}

if ($MyInvocation.InvocationName -ne '.') {
    try { Invoke-FarmDiagnostic }
    catch { Write-FarmDiagLog "ERROR: $($_.Exception.Message)"; exit 1 }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCbljvA+YVHhvw3
# dSY6aog/t7f59CZc8GK2CagXCHjvj6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFutpF+VxmehYzcxRO0kuWd7PQD+WgRSILazYysGPV6VMA0GCSqG
# SIb3DQEBAQUABIIBgDBPWuwAHe/0/EoVf2FrS7WWcFN1lDNG8avBcjVVhYXqu7Xt
# Vrt1RobvKKU+NjYQqYq1NIQiJsgQQERplgy6Mo0F3r24upiwk742C/6EIGQb/YR2
# nJhZEU9odsRH5gqt+6cTiGRXFjPEwVJCIrsxXZbhQbeWcwnQg0Gx26Sr3iTjVmiC
# GkGfDlG9LGfZLvpxkP59GpLSNCw9SEok+wK9XT803lNfThIo8vno12rjne7zeAlo
# A0VOCjbDBRQYlNSQfz7vANLL2RmEBuc546EnEbWs0eJhu7LfaxfVH7JLM6CRF2Ot
# 6BmL5dLYyLzYxqhRF4I8SRCFAsKqzdv0hTnqA6/IoZmUarKj4oLgInvCPft2G1au
# x8G5GhqzmmNN43qHHMftvyy3TM8lMcr1u26Yf2jPxyAxAV/yGFD0Tz+monH+9zhw
# J2jire0+Kfd62ld493ZXuKskUd1uI4PfGAbLs1I7gQ7IRpVZITpjb8dBV29BzoBw
# kU2F43NAg6H+XyO6KaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxMzAw
# NTZaMC8GCSqGSIb3DQEJBDEiBCCVDCj1H+u02xFR3M23OBPxKr7ALeiV+0xLRaSm
# Fr1DYTANBgkqhkiG9w0BAQEFAASCAgCZqA64aXO2pCVcdTVtSCK288HdyPqBoN9S
# oyiz9sAgWDISJaFWFHFS7E5S9QGE5oBY0fX342aNkO7QN5eN9Q+gJySOC8L4XH6q
# QPmGw62lownKCjCWWYjeHMU2hGISQoI0tiIkicSfSrYxb2Erx2C9oCs0fVKun4EC
# cEeskBhFkexxrganBsicQ0OYCygkbasdV57i1E2WH/inq6Gneza7YRfDSEvviNRN
# 20ke7UAVcGkXznWHOdC1h1j2Drq33yg0lf7xFLACiIbyOY1QeyDPFaT7Me7rDsf4
# 5NW04qR3y5Kv4MRcBnHQdv+8M165Hy0OuoSflU08d5/QmjO1sBNkfPg45EYqVkLX
# pZ+oBvm9+zGXDEBAvkvQzi6yIC65K5bCWflYQZ/jSqnS3ZPYp9RMYH16e2wVx6TN
# GSjeNmsEwqR8YijJPOc017cSrg1MrrU8a6h3KwwuWW5/nEk2P/5j1jFXH0s+32AP
# YU2xlXL4jY9rLDrGR2O5r2XyEnNJMlg/BaGTTkWCuYj1E+TnOLrZ/weo8lKWlexG
# gVPg/VxPDZkKqQliDfmmBm9L8RL4RCwtMIblv4N+/bQ0sYBXcTgD5gbSc7lDksWP
# pctRo2f7/iCZy/+kvTGx8nSzKzhnpt29aakXhtnXLwmXKLUEBbpariV8ZrEoLdtZ
# wokafQ03kw==
# SIG # End signature block
