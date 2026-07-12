<#
.SYNOPSIS
Resident scheduler that runs SmartInventory scripts unattended at their configured times.

.DESCRIPTION
SmartM365-Inventory-Orchestrator is a bounded-lifetime resident process started by a single
Windows Task Scheduler task (at server startup plus a daily trigger). It loops with a
one-minute tick, launches each job exactly at its scheduled occurrences, then exits cleanly
after a configurable maximum lifetime (default 24 hours) so Task Scheduler restarts a fresh
instance (memory recycling).

Jobs are declared in Orchestrator-Jobs.json (hot reloaded on change). Each job runs in its
own detached pwsh child process with stdout/stderr redirected to a per-run log file, so a
running job always survives an orchestrator exit. On startup the new orchestrator instance
re-adopts children recorded as Running in Orchestrator-State.json (PID plus StartTime match)
and resumes their supervision (timeout, exit code, retries). Jobs longer than 24 hours can
therefore span several orchestrator lifecycles without being interrupted.

Features:
- Daily and weekly schedules, one or several times per day, per-job MissedRunPolicy
  (RunOnce catch-up or Skip).
- Server allowlist: the AllowedServers configuration value lists the servers allowed to
  run jobs by default (empty = all servers), and each job may carry its own
  AllowedServers list in the manifest. Jobs not allowed on the local computer name are
  never launched here (including through -Force).
- Per-job PowerShellEdition: jobs run in pwsh (PowerShell7, default) or powershell.exe
  (WindowsPowerShell) child processes; Exchange on-premises scripts require
  WindowsPowerShell.
- Per-server folders: the orchestrator data and log folders are automatically suffixed
  with the local computer name, so several servers can share the same UNC
  DataAllRootPath/LogAllRootPath without state, lock or log collisions.
- Global MaxConcurrency with queueing (a due occurrence is never lost while waiting).
- DependsOn chains executed in topological order; dependents blocked when a parent with
  ContinueOnError=false finally fails, or when the dependency wait timeout is exceeded.
- Per-job overlap guard based on the state file (effective across recycles) and a global
  lock file so two orchestrator instances never run together for the same tenant.
- Timeout with process-tree kill, retry policy (MaxRetries / RetryDelaySeconds).
- Tenant-wide orchestrator lifecycle CSV, daily job-runs CSV (atomic writes), orchestrator
  log with daily rotation, one log per job execution, retention cleanup and heartbeat.
- HTML error email on final job failure (JobMailMode Always/OnError/Never), optional
  daily HTML summary email, fatal error email if the orchestrator itself crashes.
  Mail uses the shared SmartM365.Core mail helper, so the orchestrator supports the same
  Graph/SMTP/Both transport selection, branding and HTML copy behavior as inventory scripts.
- Optional Authenticode validation before job launch, configurable and usable
  in Audit mode before switching to Enforce.

.PARAMETER Tenant
Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER Connect
Passed through to every launched inventory script (forces a fresh sign-in in the scripts).

.PARAMETER DryRun
Prints the planned occurrences for the next 24 hours per job (plus pending catch-up runs),
then exits without launching anything.

.PARAMETER Once
Runs a single tick (adoption, supervision, due launches), saves state, then exits.
Launched children keep running detached. Intended for tests.

.PARAMETER Force
Job names to launch immediately on the first tick even if not due. Bypasses the schedule
and the dependency gate, but still honors the per-job overlap guard and MaxConcurrency.

.PARAMETER Only
Restricts launching to the listed job names. Supervision of already running jobs is not
affected.

.PARAMETER Skip
Excludes the listed job names from launching.

.PARAMETER MaxConcurrency
Overrides the configured global concurrency limit (default 2).

.PARAMETER MaxLifetimeHours
Overrides the configured maximum lifetime in hours (default 24).

.PARAMETER JobsManifestPath
Overrides the jobs manifest path. Default: Orchestrator-Jobs.json next to this script.

.PARAMETER StatePath
Overrides the state file path. Default: Orchestrator-State.json in the orchestrator data
folder ({{DataAllRootPath}}\Orchestrator).

.PARAMETER Stop
Writes a stop request for the currently running orchestrator instance of this tenant, then
waits for it to exit cleanly. No inventory job process is killed; running jobs remain
detached and are re-adopted by the next orchestrator instance.

.PARAMETER StopTimeoutSeconds
Maximum time to wait for a -Stop request to be consumed. Defaults to 180 seconds.

.VERSION
1.3.11

.REQUIREMENTS
    PowerShell 7+.
    Config/SmartM365-TenantContext.ps1 (SmartM365 tenant context helper).
    SmartM365.Core. Microsoft Graph modules are required only when the orchestrator mail
    transport is Graph or Both; each job script manages its own inventory connections
    inside its own child process.

.NOTES
    Version : 1.3.11
    Author: https://github.com/khda79/workplacecloudhub.com
    Exit codes: 0 = normal end (recycle, DryRun, Once), 1 = unexpected fatal error,
    2 = configuration or manifest error at startup, 3 = another live instance holds the lock.
#>

param(
    [string]$Tenant = 'test',
    [switch]$Connect,
    [switch]$DryRun,
    [switch]$Once,
    [string[]]$Force = @(),
    [string[]]$Only = @(),
    [string[]]$Skip = @(),
    [switch]$Stop,
    [int]$StopTimeoutSeconds = 180,
    [int]$MaxConcurrency = 0,
    [int]$MaxLifetimeHours = 0,
    [string]$JobsManifestPath = '',
    [string]$StatePath = ''
)

$ErrorActionPreference = 'Stop'

$ScriptVersion = '1.3.11'
$ScriptName = 'SmartM365-Inventory-Orchestrator'

# Normalize list parameters: when launched through pwsh -File, a value such as
# "-Force JobA,JobB" arrives as a single comma-separated string.
$Force = @($Force | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$Only = @($Only | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$Skip = @($Skip | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# ==========================================================
# Console timestamps (orchestrator standard): shadow Write-Host
# ==========================================================
function Write-Host {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [System.Object]$Object,
        [switch]$NoNewline,
        [System.Object]$Separator,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )

    $optional = @{}
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $optional['ForegroundColor'] = $ForegroundColor }
    if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $optional['BackgroundColor'] = $BackgroundColor }
    if ($PSBoundParameters.ContainsKey('Separator')) { $optional['Separator'] = $Separator }
    if ($NoNewline) { $optional['NoNewline'] = $true }

    $text = ''
    if ($null -ne $Object) { $text = [string]$Object }

    if ([string]::IsNullOrEmpty($text)) {
        Microsoft.PowerShell.Utility\Write-Host $text @optional
        return
    }

    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    foreach ($line in ($text -split "`r`n|`n")) {
        if ([string]::IsNullOrEmpty($line)) {
            Microsoft.PowerShell.Utility\Write-Host '' @optional
        }
        else {
            Microsoft.PowerShell.Utility\Write-Host ("[{0}] {1}" -f $stamp, $line) @optional
        }
    }
}

# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 2
}

# ==========================================================
# SmartM365 - Tenant/config integration
# ==========================================================
function Find-SmartM365TenantContextPath {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}

function Find-SmartM365CoreModulePath {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }

        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365.Core module manifest not found.'
}
function Get-SmartM365ScriptLocalConfig {
    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ('{0}.local.json' -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
        if (-not (Test-Path -LiteralPath $templatePath)) {
            throw "Local configuration file not found and template is missing: $configPath"
        }
        Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
        Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
        Write-Host "Review the new .local.json values; the orchestrator continues with template defaults." -ForegroundColor Yellow
    }
    $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    if (Get-Command Sync-SmartM365JsonConfigWithTemplate -ErrorAction SilentlyContinue) {
        return (Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $configPath)
    }
    return $config
}

function Resolve-SmartM365ScriptValue {
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $tokenValue = Resolve-SmartM365ScriptValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-SmartM365ScriptConfigValue {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -isnot [string]) { return Resolve-SmartM365ScriptValue -Value $property.Value }
        $text = $property.Value.Trim()
        if ($text -and $text -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return Resolve-SmartM365ScriptValue -Value $property.Value }
    }
    $globalProperty = $script:SmartM365EffectiveConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) { return $DefaultValue }
        return Resolve-SmartM365ScriptValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

function Get-SmartM365ScriptConfigInt {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        [int]$DefaultValue
    )

    $value = Get-SmartM365ScriptConfigValue -Config $Config -Name $Name -DefaultValue $DefaultValue
    if ($null -eq $value) { return $DefaultValue }
    try { return [int]$value } catch { return $DefaultValue }
}

function Get-SmartM365ScriptConfigBool {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        [bool]$DefaultValue
    )

    $value = Get-SmartM365ScriptConfigValue -Config $Config -Name $Name -DefaultValue $DefaultValue
    if ($null -eq $value) { return $DefaultValue }
    if ($value -is [bool]) { return $value }
    $text = ([string]$value).Trim()
    if ($text -in @('true', 'True', '1')) { return $true }
    if ($text -in @('false', 'False', '0')) { return $false }
    return $DefaultValue
}

# ==========================================================
# Logging (orchestrator log with daily rotation)
# ==========================================================
function Get-OrchestratorLogPath {
    param([datetime]$Date = (Get-Date))
    return Join-Path -Path $script:Settings.OrchestratorLogFolderPath -ChildPath ("{0}_{1}_{2}.log" -f $ScriptName, $env:COMPUTERNAME, $Date.ToString('yyyyMMdd'))
}

function Write-OrchestratorLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    foreach ($line in ($Message -split "`r`n|`n")) {
        $fileLine = "[{0}][{1}] {2}" -f $stamp, $Level, $line
        if ($script:LogReady) {
            try { Add-Content -Path (Get-OrchestratorLogPath) -Value $fileLine -Encoding UTF8 } catch { }
        }
        switch ($Level) {
            'INFO' { Write-Host $line -ForegroundColor Gray }
            'WARN' { Write-Host $line -ForegroundColor Yellow }
            'ERROR' { Write-Host $line -ForegroundColor Red }
        }
    }
}

function Test-OrchestratorSharePointUploadConfigured {
    if (-not $script:Settings.SharePointUploadEnabled) { return $false }

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($name in @('SharePointSiteHostname', 'SharePointSitePath', 'SharePointLibraryDisplayName', 'SharePointTargetFolderPath')) {
        if ([string]::IsNullOrWhiteSpace([string]$script:Settings.$name)) { $missing.Add($name) }
    }
    foreach ($name in @('AppId', 'TenantId', 'Thumbprint')) {
        if ([string]::IsNullOrWhiteSpace([string](Get-Variable -Name $name -Scope Global -ValueOnly -ErrorAction SilentlyContinue))) { $missing.Add($name) }
    }

    if ($missing.Count -gt 0) {
        $key = ($missing | Sort-Object) -join ','
        if ($script:LastSharePointConfigWarningKey -ne $key) {
            Write-OrchestratorLog -Message ("SharePoint upload enabled but not configured for orchestrator uploads. Missing: {0}" -f ($missing -join ', ')) -Level ERROR
            $script:LastSharePointConfigWarningKey = $key
        }
        return $false
    }
    return $true
}

function Invoke-OrchestratorSharePointUpload {
    param(
        [Parameter(Mandatory = $true)][string]$LocalFilePath,
        [string]$Reason = 'orchestrator artifact',
        [switch]$Force
    )

    if (-not (Test-OrchestratorSharePointUploadConfigured)) { return $false }
    if ([string]::IsNullOrWhiteSpace($LocalFilePath) -or -not (Test-Path -LiteralPath $LocalFilePath)) { return $false }

    $fileInfo = $null
    try { $fileInfo = Get-Item -LiteralPath $LocalFilePath -ErrorAction Stop }
    catch {
        Write-OrchestratorLog -Message ("SharePoint upload skipped for {0}: cannot read '{1}'. {2}" -f $Reason, $LocalFilePath, $_.Exception.Message) -Level ERROR
        return $false
    }

    $key = $fileInfo.FullName.ToLowerInvariant()
    $signature = "{0}|{1}" -f $fileInfo.Length, $fileInfo.LastWriteTimeUtc.Ticks
    if (-not $Force -and $script:SharePointUploadedFileState.ContainsKey($key) -and $script:SharePointUploadedFileState[$key] -eq $signature) {
        return $true
    }

    try {
        $record = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $fileInfo.FullName -Enabled $true -SiteHostname $script:Settings.SharePointSiteHostname -SitePath $script:Settings.SharePointSitePath -LibraryDisplayName $script:Settings.SharePointLibraryDisplayName -TargetFolderPath $script:Settings.SharePointTargetFolderPath -ErrorAction Stop
        if ($record) {
            $script:SharePointUploadedFileState[$key] = $signature
            $target = if (-not [string]::IsNullOrWhiteSpace([string]$record.WebUrl)) { [string]$record.WebUrl } else { [string]$record.SharePointPath }
            Write-OrchestratorLog -Message ("SharePoint upload OK ({0}): {1}" -f $Reason, $target)
            return $true
        }
        Write-OrchestratorLog -Message ("SharePoint upload failed for {0}: no uploaded item returned for '{1}'." -f $Reason, $fileInfo.FullName) -Level ERROR
        return $false
    }
    catch {
        Write-OrchestratorLog -Message ("SharePoint upload failed for {0}: {1}" -f $Reason, $_.Exception.Message) -Level ERROR
        return $false
    }
}

function Invoke-OrchestratorMailHtmlUploads {
    param([int]$PreviousCount = 0)

    if (-not $global:SmartM365MailHtmlFiles) { return }
    $files = @($global:SmartM365MailHtmlFiles)
    if ($files.Count -le $PreviousCount) { return }
    foreach ($file in @($files | Select-Object -Skip $PreviousCount)) {
        Invoke-OrchestratorSharePointUpload -LocalFilePath ([string]$file) -Reason 'mail HTML copy' -Force | Out-Null
    }
}

function Invoke-OrchestratorPeriodicSharePointUpload {
    param(
        [Parameter(Mandatory = $true)][datetime]$Now,
        [switch]$Force
    )

    if (-not $Force) {
        $interval = [int]$script:Settings.OrchestratorSharePointUploadIntervalMinutes
        if ($interval -le 0) { return }
        if ($script:LastSharePointUploadAttempt -ne [datetime]::MinValue -and ($Now - $script:LastSharePointUploadAttempt).TotalMinutes -lt $interval) { return }
    }
    $script:LastSharePointUploadAttempt = $Now

    foreach ($path in @(
        (Get-OrchestratorLogPath -Date $Now),
        $script:Settings.OrchestratorRunsCsvPath,
        (Get-JobRunsCsvPath),
        $script:Settings.StatePath,
        $script:Settings.HeartbeatPath
    )) {
        if (-not [string]::IsNullOrWhiteSpace([string]$path)) {
            Invoke-OrchestratorSharePointUpload -LocalFilePath ([string]$path) -Reason 'periodic orchestrator upload' -Force:$Force | Out-Null
        }
    }

    if ($global:SmartM365MailHtmlFiles) {
        foreach ($file in @($global:SmartM365MailHtmlFiles)) {
            Invoke-OrchestratorSharePointUpload -LocalFilePath ([string]$file) -Reason 'mail HTML copy' -Force:$Force | Out-Null
        }
    }
}

function Write-DependencyWaitLog {
    param(
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][string[]]$BlockingDependencies,
        [Parameter(Mandatory = $true)][datetime]$Now
    )

    $dependencies = @($BlockingDependencies | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if ($dependencies.Count -eq 0) { $dependencies = @('unknown') }
    $dependencyKey = $dependencies -join '|'
    $state = $null
    if ($script:DependencyWaitLogState.ContainsKey($JobName)) { $state = $script:DependencyWaitLogState[$JobName] }

    $shouldLog = $false
    $prefix = 'waiting for dependencies'
    if ($null -eq $state -or $state.DependencyKey -ne $dependencyKey) {
        $shouldLog = $true
        $state = [pscustomobject]@{
            DependencyKey = $dependencyKey
            FirstSeen = $Now
            LastLogged = [datetime]::MinValue
        }
    }
    else {
        $interval = [int]$script:Settings.DependencyWaitLogIntervalMinutes
        if ($interval -gt 0 -and ($Now - [datetime]$state.LastLogged).TotalMinutes -ge $interval) {
            $shouldLog = $true
            $prefix = ("still waiting for dependencies after {0} min" -f [int]($Now - [datetime]$state.FirstSeen).TotalMinutes)
        }
    }

    if ($shouldLog) {
        Write-OrchestratorLog -Message ("Job {0}: {1}: {2}." -f $JobName, $prefix, ($dependencies -join ', '))
        $state.LastLogged = $Now
        $script:DependencyWaitLogState[$JobName] = $state
    }
}

function Clear-DependencyWaitLog {
    param([Parameter(Mandatory = $true)][string]$JobName)

    if ($script:DependencyWaitLogState.ContainsKey($JobName)) {
        $state = $script:DependencyWaitLogState[$JobName]
        $duration = [int]((Get-Date) - [datetime]$state.FirstSeen).TotalMinutes
        Write-OrchestratorLog -Message ("Job {0}: dependencies cleared after {1} min; launch can continue." -f $JobName, $duration)
        $script:DependencyWaitLogState.Remove($JobName)
    }
}
# ==========================================================
# Atomic file helpers
# ==========================================================
function Write-FileAtomically {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Content
    )

    $folder = Split-Path -Path $Path -Parent
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $temp = "{0}.{1}.tmp" -f $Path, $PID
    [System.IO.File]::WriteAllText($temp, $Content)
    Move-Item -LiteralPath $temp -Destination $Path -Force
}

function ConvertTo-StateTime {
    param([Parameter(Mandatory = $true)][datetime]$Value)
    return $Value.ToString('o')
}

function ConvertFrom-StateTime {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    return [datetime]::Parse($Text, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
}

# ==========================================================
# Mail helpers (System.Net.Mail + SmtpClient, explicit UTF-8)
# ==========================================================
function Resolve-IPv4Address {
    param([Parameter(Mandatory = $true)][string]$HostName)

    $addresses = [System.Net.Dns]::GetHostAddresses($HostName) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }
    $first = $addresses | Select-Object -First 1
    if (-not $first) { throw "No IPv4 address found for $HostName" }
    return $first.IPAddressToString
}

function ConvertTo-HtmlText {
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function New-HtmlMailMessage {
    param(
        [Parameter(Mandatory = $true)][string]$From,
        [Parameter(Mandatory = $true)][string]$To,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$HtmlBody,
        [string]$Cc = '',
        [string]$Bcc = ''
    )

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = New-Object System.Net.Mail.MailAddress($From)
    foreach ($addr in ($To -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $mail.To.Add((New-Object System.Net.Mail.MailAddress($addr.Trim())))
    }
    foreach ($addr in ($Cc -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $mail.CC.Add((New-Object System.Net.Mail.MailAddress($addr.Trim())))
    }
    foreach ($addr in ($Bcc -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $mail.Bcc.Add((New-Object System.Net.Mail.MailAddress($addr.Trim())))
    }
    $mail.Subject = $Subject
    $mail.Body = $HtmlBody
    $mail.IsBodyHtml = $true
    $utf8 = [System.Text.Encoding]::UTF8
    $mail.SubjectEncoding = $utf8
    $mail.BodyEncoding = $utf8
    $mail.HeadersEncoding = $utf8
    return $mail
}

function Send-HtmlMail {
    param(
        [Parameter(Mandatory = $true)][System.Net.Mail.MailMessage]$MailMessage,
        [Parameter(Mandatory = $true)][string]$SmtpEndpoint,
        [Parameter(Mandatory = $true)][int]$SmtpPort,
        [Parameter(Mandatory = $true)][bool]$UseIntegratedAuth,
        [bool]$UseSsl = $false
    )

    $client = New-Object System.Net.Mail.SmtpClient($SmtpEndpoint, $SmtpPort)
    $client.EnableSsl = $UseSsl
    if ($UseIntegratedAuth) { $client.UseDefaultCredentials = $true } else { $client.UseDefaultCredentials = $false }
    try { $client.Send($MailMessage) }
    finally { $MailMessage.Dispose(); $client.Dispose() }
}

function Send-OrchestratorMail {
    param(
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$HtmlBody,
        [switch]$IsError
    )

    if (-not $script:Settings.MailEnabled) {
        Write-OrchestratorLog -Message ("Mail disabled ({0}); email not sent: {1}" -f $script:Settings.MailConfigIssue, $Subject) -Level WARN
        return $false
    }
    $to = $script:Settings.MailTo
    if ($IsError -and -not [string]::IsNullOrWhiteSpace($script:Settings.ErrorMailTo)) { $to = $script:Settings.ErrorMailTo }
    if ([string]::IsNullOrWhiteSpace($to)) {
        Write-OrchestratorLog -Message ("No recipient resolved; email not sent: {0}" -f $Subject) -Level WARN
        return $false
    }
    try {
        if (-not [string]::IsNullOrWhiteSpace($script:Settings.MailBcc)) {
            Write-OrchestratorLog -Message "Bcc is configured but is not supported by the shared SmartM365 mail helper; Bcc is ignored for orchestrator mail." -Level WARN
        }

        $mailHtmlCountBefore = if ($global:SmartM365MailHtmlFiles) { @($global:SmartM365MailHtmlFiles).Count } else { 0 }

        Send-SmartM365Mail -SmtpServer $script:Settings.SmtpServer -SmtpPort $script:Settings.SmtpPort -SendMailMode $script:Settings.SendMailMode -From $script:Settings.MailFrom -To $to -Cc $script:Settings.MailCc -Subject $Subject -BodyHtml $HtmlBody -BodyAsHtml -HighPriority:$IsError -ErrorAction Stop
        Write-OrchestratorLog -Message ("Email sent via {0}: {1}" -f $script:Settings.SendMailMode, $Subject)

        Invoke-OrchestratorMailHtmlUploads -PreviousCount $mailHtmlCountBefore
        return $true
    }
    catch {
        Write-OrchestratorLog -Message ("Failed to send email via {0} '{1}': {2}" -f $script:Settings.SendMailMode, $Subject, $_.Exception.Message) -Level WARN
        return $false
    }
}

function Get-JobLogTailHtml {
    param([AllowNull()][AllowEmptyString()][string]$LogPath)

    if ([string]::IsNullOrWhiteSpace($LogPath) -or -not (Test-Path -LiteralPath $LogPath)) { return '' }
    try {
        $tail = Get-Content -LiteralPath $LogPath -Tail 40 -ErrorAction Stop
        if (-not $tail) { return '' }
        $safe = ConvertTo-HtmlText -Text (($tail | ForEach-Object { [string]$_ }) -join [Environment]::NewLine)
        return "<p><b>Last job log lines:</b></p><pre style='background-color:#F5F5F5;border:1px solid #DDDDDD;padding:8px;font-family:Consolas,monospace;font-size:12px;'>$safe</pre>"
    }
    catch { return '' }
}

function Send-JobResultEmail {
    param(
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][hashtable]$Details,
        [AllowNull()][AllowEmptyString()][string]$LogPath
    )

    $isFailure = $Status -ne 'Success'
    $color = '#107C10'
    if ($isFailure) { $color = '#D13438' }
    $rows = ''
    foreach ($key in $Details.Keys) {
        $rows += "<tr><td style='padding:3px 10px;border:1px solid #DDDDDD;'><b>{0}</b></td><td style='padding:3px 10px;border:1px solid #DDDDDD;'>{1}</td></tr>" -f (ConvertTo-HtmlText -Text $key), (ConvertTo-HtmlText -Text ([string]$Details[$key]))
    }
    $body = @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#1F2937;'>
<h2 style='color:$color;'>Job $(ConvertTo-HtmlText -Text $JobName): $(ConvertTo-HtmlText -Text $Status)</h2>
<table style='border-collapse:collapse;'>$rows</table>
$(Get-JobLogTailHtml -LogPath $LogPath)
<p style='color:#5F6B7A;'>Sent by $ScriptName v$ScriptVersion on $(ConvertTo-HtmlText -Text $env:COMPUTERNAME) (tenant $(ConvertTo-HtmlText -Text $Tenant)).</p>
</body></html>
"@
    $statusWord = 'succeeded'
    if ($isFailure) { $statusWord = "failed ($Status)" }
    $subject = "[SmartM365 Orchestrator][$Tenant] Job $JobName $statusWord"
    Send-OrchestratorMail -Subject $subject -HtmlBody $body -IsError:$isFailure | Out-Null
}

function Send-OrchestratorFatalEmail {
    param([Parameter(Mandatory = $true)][string]$ErrorText)

    $body = @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#1F2937;'>
<h2 style='color:#D13438;'>Orchestrator fatal error</h2>
<p>$ScriptName v$ScriptVersion crashed on $(ConvertTo-HtmlText -Text $env:COMPUTERNAME) (tenant $(ConvertTo-HtmlText -Text $Tenant)).</p>
<pre style='background-color:#F5F5F5;border:1px solid #DDDDDD;padding:8px;font-family:Consolas,monospace;font-size:12px;'>$(ConvertTo-HtmlText -Text $ErrorText)</pre>
<p>Running jobs (detached pwsh children) are not affected and will be re-adopted at the next start.</p>
</body></html>
"@
    Send-OrchestratorMail -Subject "[SmartM365 Orchestrator][$Tenant] Orchestrator fatal error" -HtmlBody $body -IsError | Out-Null
}

# ==========================================================
# Jobs manifest (load, validate, hot reload)
# ==========================================================
function ConvertTo-NormalizedJob {
    param(
        [Parameter(Mandatory = $true)]$RawJob,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[string]]$Errors
    )

    $name = ''
    if ($RawJob.PSObject.Properties['Name'] -and $RawJob.Name) { $name = [string]$RawJob.Name }
    if ([string]::IsNullOrWhiteSpace($name)) { $Errors.Add('A job has an empty Name.'); return $null }
    if ($name -notmatch '^[A-Za-z0-9._-]+$') { $Errors.Add("Job '$name': Name may only contain letters, digits, dot, underscore and dash."); return $null }

    $scriptPath = ''
    if ($RawJob.PSObject.Properties['ScriptPath'] -and $RawJob.ScriptPath) { $scriptPath = [string]$RawJob.ScriptPath }
    if ([string]::IsNullOrWhiteSpace($scriptPath)) { $Errors.Add("Job '$name': ScriptPath is required.") }
    $launcherPath = ''
    if ($RawJob.PSObject.Properties['LauncherPath'] -and $RawJob.LauncherPath) { $launcherPath = [string]$RawJob.LauncherPath }

    $arguments = ''
    if ($RawJob.PSObject.Properties['Arguments'] -and $RawJob.Arguments) { $arguments = [string]$RawJob.Arguments }
    $enabled = $false
    if ($RawJob.PSObject.Properties['Enabled']) { $enabled = [bool]$RawJob.Enabled }
    $group = ''
    if ($RawJob.PSObject.Properties['Group'] -and $RawJob.Group) { $group = [string]$RawJob.Group }
    $dependsOn = @()
    if ($RawJob.PSObject.Properties['DependsOn'] -and $null -ne $RawJob.DependsOn) { $dependsOn = @($RawJob.DependsOn | ForEach-Object { [string]$_ }) }
    $timeoutMinutes = 240
    if ($RawJob.PSObject.Properties['TimeoutMinutes'] -and $null -ne $RawJob.TimeoutMinutes) { $timeoutMinutes = [int]$RawJob.TimeoutMinutes }
    $maxRetries = 0
    if ($RawJob.PSObject.Properties['MaxRetries'] -and $null -ne $RawJob.MaxRetries) { $maxRetries = [int]$RawJob.MaxRetries }
    $retryDelaySeconds = 300
    if ($RawJob.PSObject.Properties['RetryDelaySeconds'] -and $null -ne $RawJob.RetryDelaySeconds) { $retryDelaySeconds = [int]$RawJob.RetryDelaySeconds }
    $continueOnError = $true
    if ($RawJob.PSObject.Properties['ContinueOnError']) { $continueOnError = [bool]$RawJob.ContinueOnError }
    $allowedServers = @()
    if ($RawJob.PSObject.Properties['AllowedServers'] -and $null -ne $RawJob.AllowedServers) {
        $allowedServers = @($RawJob.AllowedServers | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    }
    $requiredLogPatterns = @('Execution context:|Environment initialized successfully|Starting .* v')
    if ($RawJob.PSObject.Properties['RequiredLogPatterns']) {
        $requiredLogPatterns = @($RawJob.RequiredLogPatterns | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    }
    $minimumSuccessDurationSeconds = 0
    if ($RawJob.PSObject.Properties['MinimumSuccessDurationSeconds'] -and $null -ne $RawJob.MinimumSuccessDurationSeconds) {
        $minimumSuccessDurationSeconds = [int]$RawJob.MinimumSuccessDurationSeconds
    }
    $dependencyWaitTimeoutMinutes = 0
    if ($RawJob.PSObject.Properties['DependencyWaitTimeoutMinutes'] -and $null -ne $RawJob.DependencyWaitTimeoutMinutes) {
        $dependencyWaitTimeoutMinutes = [int]$RawJob.DependencyWaitTimeoutMinutes
    }
    $powerShellEdition = 'PowerShell7'
    if ($RawJob.PSObject.Properties['PowerShellEdition'] -and $RawJob.PowerShellEdition) { $powerShellEdition = [string]$RawJob.PowerShellEdition }
    if ($powerShellEdition -notin @('PowerShell7', 'WindowsPowerShell')) {
        $Errors.Add("Job '$name': PowerShellEdition must be 'PowerShell7' or 'WindowsPowerShell'.")
    }

    $scheduleType = ''
    $times = @()
    $daysOfWeek = @()
    $missedRunPolicy = 'RunOnce'
    if (-not $RawJob.PSObject.Properties['Schedule'] -or $null -eq $RawJob.Schedule) {
        $Errors.Add("Job '$name': Schedule is required.")
    }
    else {
        $schedule = $RawJob.Schedule
        if ($schedule.PSObject.Properties['Type'] -and $schedule.Type) { $scheduleType = [string]$schedule.Type }
        if ($scheduleType -notin @('Daily', 'Weekly')) { $Errors.Add("Job '$name': Schedule.Type must be 'Daily' or 'Weekly'.") }
        if ($schedule.PSObject.Properties['Times'] -and $null -ne $schedule.Times) { $times = @($schedule.Times | ForEach-Object { [string]$_ }) }
        if ($times.Count -eq 0) { $Errors.Add("Job '$name': Schedule.Times must contain at least one 'HH:mm' value.") }
        foreach ($t in $times) {
            if ($t -notmatch '^\d{2}:\d{2}$') { $Errors.Add("Job '$name': invalid Schedule time '$t' (expected HH:mm)."); continue }
            try { [TimeSpan]::ParseExact($t, 'hh\:mm', [System.Globalization.CultureInfo]::InvariantCulture) | Out-Null }
            catch { $Errors.Add("Job '$name': invalid Schedule time '$t' (expected HH:mm, 00:00 to 23:59).") }
        }
        if ($scheduleType -eq 'Weekly') {
            if ($schedule.PSObject.Properties['DaysOfWeek'] -and $null -ne $schedule.DaysOfWeek) { $daysOfWeek = @($schedule.DaysOfWeek | ForEach-Object { [string]$_ }) }
            if ($daysOfWeek.Count -eq 0) { $Errors.Add("Job '$name': Weekly schedule requires DaysOfWeek.") }
            $validDays = @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')
            foreach ($d in $daysOfWeek) {
                if ($d -notin $validDays) { $Errors.Add("Job '$name': invalid DaysOfWeek value '$d'.") }
            }
        }
        if ($schedule.PSObject.Properties['MissedRunPolicy'] -and $schedule.MissedRunPolicy) { $missedRunPolicy = [string]$schedule.MissedRunPolicy }
        if ($missedRunPolicy -notin @('RunOnce', 'Skip')) { $Errors.Add("Job '$name': Schedule.MissedRunPolicy must be 'RunOnce' or 'Skip'.") }
    }

    return [pscustomobject]@{
        Name = $name
        ScriptPath = $scriptPath
        LauncherPath = $launcherPath
        Arguments = $arguments
        Enabled = $enabled
        Group = $group
        DependsOn = $dependsOn
        TimeoutMinutes = $timeoutMinutes
        MaxRetries = $maxRetries
        RetryDelaySeconds = $retryDelaySeconds
        ContinueOnError = $continueOnError
        AllowedServers = $allowedServers
        RequiredLogPatterns = $requiredLogPatterns
        MinimumSuccessDurationSeconds = $minimumSuccessDurationSeconds
        DependencyWaitTimeoutMinutes = $dependencyWaitTimeoutMinutes
        PowerShellEdition = $powerShellEdition
        Schedule = [pscustomobject]@{
            Type = $scheduleType
            Times = $times
            DaysOfWeek = $daysOfWeek
            MissedRunPolicy = $missedRunPolicy
        }
    }
}

function Get-TopologicalJobOrder {
    param([Parameter(Mandatory = $true)]$Jobs)

    $inDegree = @{}
    $children = @{}
    foreach ($j in $Jobs) {
        $inDegree[$j.Name] = 0
        $children[$j.Name] = New-Object System.Collections.Generic.List[string]
    }
    foreach ($j in $Jobs) {
        foreach ($dep in $j.DependsOn) {
            $children[$dep].Add($j.Name)
            $inDegree[$j.Name] = $inDegree[$j.Name] + 1
        }
    }
    $queue = New-Object System.Collections.Generic.Queue[string]
    foreach ($j in $Jobs) {
        if ($inDegree[$j.Name] -eq 0) { $queue.Enqueue($j.Name) }
    }
    $orderedNames = New-Object System.Collections.Generic.List[string]
    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $orderedNames.Add($current)
        foreach ($child in $children[$current]) {
            $inDegree[$child] = $inDegree[$child] - 1
            if ($inDegree[$child] -eq 0) { $queue.Enqueue($child) }
        }
    }
    if ($orderedNames.Count -ne $Jobs.Count) {
        throw 'Dependency cycle detected in DependsOn definitions.'
    }
    $byName = @{}
    foreach ($j in $Jobs) { $byName[$j.Name] = $j }
    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($n in $orderedNames) { $ordered.Add($byName[$n]) }
    return $ordered
}

function Read-JobsManifest {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "Jobs manifest not found: $Path" }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    if (-not $raw.PSObject.Properties['Jobs'] -or $null -eq $raw.Jobs) { throw "Jobs manifest has no 'Jobs' array: $Path" }

    $errors = New-Object System.Collections.Generic.List[string]
    $jobs = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($rawJob in @($raw.Jobs)) {
        $job = ConvertTo-NormalizedJob -RawJob $rawJob -Errors $errors
        if ($null -eq $job) { continue }
        if ($seen.ContainsKey($job.Name)) { $errors.Add("Duplicate job name: $($job.Name)"); continue }
        $seen[$job.Name] = $true
        $jobs.Add($job)
    }
    foreach ($job in $jobs) {
        foreach ($dep in $job.DependsOn) {
            if (-not $seen.ContainsKey($dep)) { $errors.Add("Job '$($job.Name)': DependsOn references unknown job '$dep'.") }
        }
    }
    $ordered = $null
    if ($errors.Count -eq 0) {
        try { $ordered = @(Get-TopologicalJobOrder -Jobs $jobs) }
        catch { $errors.Add($_.Exception.Message) }
    }
    if ($errors.Count -gt 0) {
        throw ("Invalid jobs manifest '{0}':{1}{2}" -f $Path, [Environment]::NewLine, ($errors -join [Environment]::NewLine))
    }

    $byName = @{}
    foreach ($j in $ordered) { $byName[$j.Name] = $j }
    return [pscustomobject]@{
        Path = $Path
        FileWriteTimeUtc = (Get-Item -LiteralPath $Path).LastWriteTimeUtc
        OrderedJobs = $ordered
        JobsByName = $byName
    }
}

function Update-JobsManifestIfChanged {
    if (-not (Test-Path -LiteralPath $script:Settings.JobsManifestPath)) {
        Write-OrchestratorLog -Message ("Jobs manifest disappeared: {0}; keeping the last valid manifest in memory." -f $script:Settings.JobsManifestPath) -Level WARN
        return
    }
    $writeTime = (Get-Item -LiteralPath $script:Settings.JobsManifestPath).LastWriteTimeUtc
    if ($writeTime -eq $script:Manifest.FileWriteTimeUtc) { return }
    try {
        $newManifest = Read-JobsManifest -Path $script:Settings.JobsManifestPath
        $script:Manifest = $newManifest
        Write-OrchestratorLog -Message ("Jobs manifest reloaded ({0} jobs)." -f $newManifest.OrderedJobs.Count)
        Write-ServerAllowlistSummary
        Initialize-NewJobStates
    }
    catch {
        $message = $_.Exception.Message
        Write-OrchestratorLog -Message ("Jobs manifest reload rejected; the last valid manifest stays in effect. {0}" -f $message) -Level ERROR
        $body = @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#1F2937;'>
<h2 style='color:#D13438;'>Invalid jobs manifest rejected</h2>
<p>The orchestrator keeps running with the last valid manifest.</p>
<pre style='background-color:#F5F5F5;border:1px solid #DDDDDD;padding:8px;font-family:Consolas,monospace;font-size:12px;'>$(ConvertTo-HtmlText -Text $message)</pre>
</body></html>
"@
        Send-OrchestratorMail -Subject "[SmartM365 Orchestrator][$Tenant] Invalid jobs manifest rejected" -HtmlBody $body -IsError | Out-Null
        # Remember the rejected write time so the error is not repeated every tick.
        $script:Manifest.FileWriteTimeUtc = $writeTime
    }
}

# ==========================================================
# State file
# ==========================================================
function Read-OrchestratorState {
    if (Test-Path -LiteralPath $script:Settings.StatePath) {
        try {
            $state = Get-Content -LiteralPath $script:Settings.StatePath -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
            if ($state -is [hashtable] -and $state.ContainsKey('Jobs')) { return $state }
            Write-OrchestratorLog -Message "State file has an unexpected shape; starting with a fresh state." -Level WARN
        }
        catch {
            Write-OrchestratorLog -Message ("Failed to read state file '{0}': {1}; starting with a fresh state." -f $script:Settings.StatePath, $_.Exception.Message) -Level WARN
        }
    }
    return @{
        SchemaVersion = 1
        LastDailySummaryDate = ''
        Jobs = @{}
    }
}

function Save-OrchestratorState {
    $script:State['UpdatedUtc'] = (Get-Date).ToUniversalTime().ToString('o')
    Write-FileAtomically -Path $script:Settings.StatePath -Content ($script:State | ConvertTo-Json -Depth 10)
}

function Get-JobState {
    param([Parameter(Mandatory = $true)][string]$JobName)

    if (-not $script:State.Jobs.ContainsKey($JobName)) {
        $script:State.Jobs[$JobName] = @{
            LastScheduledOccurrence = $null
            LastRunStart = $null
            LastRunEnd = $null
            LastStatus = ''
            LastExitCode = $null
            RetryCount = 0
            Running = $null
            PendingRetry = $null
        }
    }
    return $script:State.Jobs[$JobName]
}

# ==========================================================
# Schedule computation
# ==========================================================
function Get-JobOccurrencesInWindow {
    param(
        [Parameter(Mandatory = $true)]$Job,
        [Parameter(Mandatory = $true)][datetime]$WindowStart,
        [Parameter(Mandatory = $true)][datetime]$WindowEnd
    )

    $result = New-Object System.Collections.Generic.List[datetime]
    if ($WindowStart -gt $WindowEnd) { return $result }
    $schedule = $Job.Schedule
    $day = $WindowStart.Date
    while ($day -le $WindowEnd.Date) {
        $dayApplies = $false
        if ($schedule.Type -eq 'Daily') { $dayApplies = $true }
        elseif ($schedule.Type -eq 'Weekly') { $dayApplies = ($schedule.DaysOfWeek -contains $day.DayOfWeek.ToString()) }
        if ($dayApplies) {
            foreach ($t in $schedule.Times) {
                $ts = [TimeSpan]::ParseExact($t, 'hh\:mm', [System.Globalization.CultureInfo]::InvariantCulture)
                $occurrence = $day.Add($ts)
                if ($occurrence -ge $WindowStart -and $occurrence -le $WindowEnd) { $result.Add($occurrence) }
            }
        }
        $day = $day.AddDays(1)
    }
    $result.Sort()
    return $result
}

function Get-DueOccurrence {
    # Returns the most recent occurrence <= Now that is newer than the last handled one.
    param(
        [Parameter(Mandatory = $true)]$Job,
        [AllowNull()]$LastOccurrence,
        [Parameter(Mandatory = $true)][datetime]$Now
    )

    $windowStart = $Now.Date.AddDays(-8)
    if ($null -ne $LastOccurrence) { $windowStart = ([datetime]$LastOccurrence).AddSeconds(1) }
    $occurrences = Get-JobOccurrencesInWindow -Job $Job -WindowStart $windowStart -WindowEnd $Now
    if ($occurrences.Count -eq 0) { return $null }
    return $occurrences[$occurrences.Count - 1]
}

function Get-LatestPastOccurrence {
    param(
        [Parameter(Mandatory = $true)]$Job,
        [Parameter(Mandatory = $true)][datetime]$Now
    )

    $occurrences = Get-JobOccurrencesInWindow -Job $Job -WindowStart $Now.Date.AddDays(-8) -WindowEnd $Now
    if ($occurrences.Count -eq 0) { return $null }
    return $occurrences[$occurrences.Count - 1]
}

# ==========================================================
# Job-runs CSV (daily file, atomic writes)
# ==========================================================
function Get-JobRunsCsvPath {
    return Join-Path -Path $script:Settings.JobRunsFolderPath -ChildPath ("Orchestrator_JobRuns_{0}.csv" -f (Get-Date).ToString('yyyyMMdd'))
}

function ConvertTo-CsvField {
    param([AllowNull()][AllowEmptyString()][string]$Value)
    if ($null -eq $Value) { $Value = '' }
    return '"{0}"' -f $Value.Replace('"', '""')
}

function Add-JobRunCsvRow {
    param(
        [Parameter(Mandatory = $true)][string]$JobName,
        [AllowNull()]$ScheduledTime,
        [AllowNull()]$StartTime,
        [AllowNull()]$EndTime,
        [AllowNull()]$DurationSec,
        [AllowNull()]$ExitCode,
        [Parameter(Mandatory = $true)][string]$Status,
        [int]$RetryCount = 0,
        [string]$LogPath = ''
    )

    $format = 'yyyy-MM-dd HH:mm:ss'
    $scheduledText = ''
    if ($null -ne $ScheduledTime) { $scheduledText = ([datetime]$ScheduledTime).ToString($format) }
    $startText = ''
    if ($null -ne $StartTime) { $startText = ([datetime]$StartTime).ToString($format) }
    $endText = ''
    if ($null -ne $EndTime) { $endText = ([datetime]$EndTime).ToString($format) }
    $durationText = ''
    if ($null -ne $DurationSec) { $durationText = [string][int]$DurationSec }
    $exitText = ''
    if ($null -ne $ExitCode) { $exitText = [string]$ExitCode }

    $fields = @($JobName, $scheduledText, $startText, $endText, $durationText, $exitText, $Status, [string]$RetryCount, $LogPath)
    $line = (($fields | ForEach-Object { ConvertTo-CsvField -Value $_ }) -join ',')
    $csvPath = Get-JobRunsCsvPath
    $header = '"JobName","ScheduledTime","StartTime","EndTime","DurationSec","ExitCode","Status","RetryCount","LogPath"'
    $content = ''
    if (Test-Path -LiteralPath $csvPath) {
        $content = [System.IO.File]::ReadAllText($csvPath)
        if ($content.Length -gt 0 -and -not $content.EndsWith("`n")) { $content += [Environment]::NewLine }
    }
    else {
        $content = $header + [Environment]::NewLine
    }
    $content += $line + [Environment]::NewLine
    Write-FileAtomically -Path $csvPath -Content $content
}

# ==========================================================
# Orchestrator lifecycle runs CSV (shared per tenant)
# ==========================================================
function Get-OrchestratorRunMode {
    if ($DryRun) { return 'DryRun' }
    if ($Once) { return 'Once' }
    return 'Resident'
}

function Get-OrchestratorRunUserName {
    try {
        return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    catch {
        if (-not [string]::IsNullOrWhiteSpace($env:USERDOMAIN)) {
            return '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME
        }
        return [string]$env:USERNAME
    }
}

function Protect-OrchestratorRunText {
    param([AllowNull()][AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $cleanText = (($Text -replace '[\r\n]+', ' ') -replace '\s{2,}', ' ').Trim()
    if ($cleanText.Length -gt 1000) { $cleanText = $cleanText.Substring(0, 1000) }
    return $cleanText
}

function Test-OrchestratorRunsCsvLockException {
    param([AllowNull()][Exception]$Exception)

    $current = $Exception
    while ($null -ne $current) {
        if ($current -is [System.IO.IOException] -or $current -is [UnauthorizedAccessException]) {
            return $true
        }
        $current = $current.InnerException
    }

    return $false
}

function Invoke-WithOrchestratorRunsCsvLock {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [AllowEmptyCollection()][object[]]$ArgumentList = @(),
        [int]$TimeoutSeconds = 0,
        [int]$RetryDelayMilliseconds = 250
    )

    if ($TimeoutSeconds -le 0) {
        $TimeoutSeconds = 120
        if ($script:Settings -and $script:Settings.PSObject.Properties['OrchestratorRunsCsvLockTimeoutSeconds']) {
            $TimeoutSeconds = [int]$script:Settings.OrchestratorRunsCsvLockTimeoutSeconds
        }
    }

    $lockStream = $null
    $deadline = (Get-Date).AddSeconds([Math]::Max(1, $TimeoutSeconds))
    $lastException = $null
    $attempt = 0

    while ((Get-Date) -lt $deadline) {
        try {
            $lockStream = [System.IO.File]::Open(
                $script:Settings.OrchestratorRunsLockPath,
                [System.IO.FileMode]::OpenOrCreate,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None
            )
            break
        }
        catch {
            $lastException = $_.Exception
            if (-not (Test-OrchestratorRunsCsvLockException -Exception $lastException)) { throw }
            $attempt++
            if (($attempt % 40) -eq 0) {
                Write-OrchestratorLog -Message ("Waiting for orchestrator-runs CSV lock: {0}" -f $script:Settings.OrchestratorRunsLockPath) -Level WARN
            }
            Start-Sleep -Milliseconds ([Math]::Max(50, $RetryDelayMilliseconds))
        }
    }

    if ($null -eq $lockStream) {
        $detail = if ($lastException) { $lastException.Message } else { 'timeout elapsed' }
        throw ("Could not acquire the orchestrator-runs CSV lock after {0}s: {1}. Last error: {2}" -f $TimeoutSeconds, $script:Settings.OrchestratorRunsLockPath, $detail)
    }

    try {
        & $Action @ArgumentList
    }
    finally {
        $lockStream.Dispose()
    }
}
function Read-OrchestratorRunsCsv {
    $csvPath = $script:Settings.OrchestratorRunsCsvPath
    if (-not (Test-Path -LiteralPath $csvPath)) { return @() }
    if ((Get-Item -LiteralPath $csvPath).Length -eq 0) { return @() }
    return @(Import-Csv -LiteralPath $csvPath)
}

function Write-OrchestratorRunsCsv {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows)

    $columns = @(
        'RunId',
        'Tenant',
        'StartDateTime',
        'StartDateTimeUtc',
        'EndDateTime',
        'EndDateTimeUtc',
        'DurationSec',
        'Server',
        'User',
        'ProcessId',
        'OrchestratorVersion',
        'PowerShellVersion',
        'Mode',
        'Connect',
        'Status',
        'ExitCode',
        'StopReason',
        'ErrorMessage'
    )
    $csvLines = @($Rows | Select-Object -Property $columns | ConvertTo-Csv -NoTypeInformation)
    $content = ($csvLines -join [Environment]::NewLine) + [Environment]::NewLine
    Write-FileAtomically -Path $script:Settings.OrchestratorRunsCsvPath -Content $content
}

function Test-OrchestratorRunProcessActive {
    param(
        [Parameter(Mandatory = $true)][int]$ProcessId,
        [Parameter(Mandatory = $true)][string]$StartDateTime
    )

    try {
        $process = Get-Process -Id $ProcessId -ErrorAction Stop
        if ($process.ProcessName -notin @('pwsh', 'powershell')) { return $false }
        $expectedStart = ConvertFrom-StateTime -Text $StartDateTime
        if ($null -eq $expectedStart) { return $false }
        return [math]::Abs(($process.StartTime - $expectedStart).TotalSeconds) -le 5
    }
    catch {
        return $false
    }
}

function Add-OrchestratorRunTracking {
    $runId = [guid]::NewGuid().ToString()
    $startTime = $script:StartTime
    $detectedAt = (Get-Date).ToString('o')
    $server = [string]$env:COMPUTERNAME

    $action = {
        param($runId, $startTime, $detectedAt, $server)
        $rows = @(Read-OrchestratorRunsCsv)
        foreach ($row in $rows) {
            if ($row.Status -ne 'Running' -or $row.Tenant -ne $Tenant -or $row.Server -ne $server) { continue }
            $isActive = $false
            [int]$recordedPid = 0
            if ([int]::TryParse([string]$row.ProcessId, [ref]$recordedPid)) {
                $isActive = Test-OrchestratorRunProcessActive -ProcessId $recordedPid -StartDateTime ([string]$row.StartDateTime)
            }
            if ($isActive) { continue }

            $row.Status = 'Interrupted'
            $row.StopReason = 'UnexpectedStop'
            $row.ErrorMessage = "Detected at $detectedAt by a later orchestrator start; the recorded process is no longer running."
        }

        $rows += [pscustomobject][ordered]@{
            RunId = $runId
            Tenant = $Tenant
            StartDateTime = $startTime.ToString('o')
            StartDateTimeUtc = $startTime.ToUniversalTime().ToString('o')
            EndDateTime = ''
            EndDateTimeUtc = ''
            DurationSec = ''
            Server = $server
            User = Get-OrchestratorRunUserName
            ProcessId = [string]$PID
            OrchestratorVersion = $ScriptVersion
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
            Mode = Get-OrchestratorRunMode
            Connect = [string]$Connect.IsPresent
            Status = 'Running'
            ExitCode = ''
            StopReason = ''
            ErrorMessage = ''
        }
        Write-OrchestratorRunsCsv -Rows $rows
    }
    try {
        Invoke-WithOrchestratorRunsCsvLock -Action $action -ArgumentList @($runId, $startTime, $detectedAt, $server)
        $script:OrchestratorRunId = $runId
        $script:OrchestratorRunRegistered = $true
    }
    catch {
        Write-OrchestratorLog -Message ("Lifecycle run tracking skipped; orchestrator continues. {0}" -f $_.Exception.Message) -Level WARN
        $script:OrchestratorRunRegistered = $false
    }
}

function Complete-OrchestratorRunTracking {
    param(
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][int]$ExitCode,
        [Parameter(Mandatory = $true)][string]$StopReason,
        [AllowNull()][AllowEmptyString()][string]$ErrorMessage = ''
    )

    if (-not $script:OrchestratorRunRegistered) { return }
    $endTime = Get-Date

    $action = {
        param($endTime, $Status, $ExitCode, $StopReason, $ErrorMessage, $RunId)
        $rows = @(Read-OrchestratorRunsCsv)
        $row = $rows | Where-Object { $_.RunId -eq $RunId } | Select-Object -First 1
        if ($null -eq $row) { throw "Orchestrator run '$RunId' was not found in the lifecycle CSV." }

        $startTime = ConvertFrom-StateTime -Text ([string]$row.StartDateTime)
        $row.EndDateTime = $endTime.ToString('o')
        $row.EndDateTimeUtc = $endTime.ToUniversalTime().ToString('o')
        if ($null -ne $startTime) {
            $row.DurationSec = [string][int][math]::Max(0, ($endTime - $startTime).TotalSeconds)
        }
        $row.Status = $Status
        $row.ExitCode = [string]$ExitCode
        $row.StopReason = $StopReason
        $row.ErrorMessage = Protect-OrchestratorRunText -Text $ErrorMessage
        Write-OrchestratorRunsCsv -Rows $rows
    }
    try {
        Invoke-WithOrchestratorRunsCsvLock -Action $action -ArgumentList @($endTime, $Status, $ExitCode, $StopReason, $ErrorMessage, $script:OrchestratorRunId)
    }
    catch {
        Write-OrchestratorLog -Message ("Lifecycle run completion update skipped. {0}" -f $_.Exception.Message) -Level WARN
    }
}

# ==========================================================
# Global lock and heartbeat
# ==========================================================
function Enter-OrchestratorLock {
    $lockPath = $script:Settings.LockPath
    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        try {
            $stream = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
            try {
                $payload = @{ Pid = $PID; StartTimeUtc = (Get-Date).ToUniversalTime().ToString('o'); ComputerName = $env:COMPUTERNAME } | ConvertTo-Json
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
                $stream.Write($bytes, 0, $bytes.Length)
            }
            finally { $stream.Dispose() }
            return $true
        }
        catch [System.IO.IOException] {
            $ownerPid = 0
            try {
                $existing = Get-Content -LiteralPath $lockPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                if ($existing.PSObject.Properties['Pid']) { $ownerPid = [int]$existing.Pid }
            }
            catch { }
            $ownerProcess = $null
            if ($ownerPid -gt 0) { $ownerProcess = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue }
            if ($ownerProcess -and $ownerProcess.ProcessName -eq 'pwsh') {
                Write-OrchestratorLog -Message ("Another orchestrator instance is already running (PID {0}); exiting." -f $ownerPid) -Level ERROR
                return $false
            }
            Write-OrchestratorLog -Message ("Stale orchestrator lock found (PID {0} is gone); recovering the lock." -f $ownerPid) -Level WARN
            try { Remove-Item -LiteralPath $lockPath -Force -ErrorAction Stop }
            catch {
                Write-OrchestratorLog -Message ("Failed to remove stale lock '{0}': {1}" -f $lockPath, $_.Exception.Message) -Level ERROR
                return $false
            }
        }
    }
    Write-OrchestratorLog -Message "Could not acquire the orchestrator lock after stale-lock recovery." -Level ERROR
    return $false
}

function Exit-OrchestratorLock {
    if (-not $script:LockOwned) { return }
    try { Remove-Item -LiteralPath $script:Settings.LockPath -Force -ErrorAction Stop } catch { }
    $script:LockOwned = $false
}

function Get-OrchestratorLockOwner {
    if (-not (Test-Path -LiteralPath $script:Settings.LockPath)) { return $null }

    $ownerPid = 0
    $ownerStartTimeUtc = ''
    try {
        $existing = Get-Content -LiteralPath $script:Settings.LockPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($existing.PSObject.Properties['Pid']) { $ownerPid = [int]$existing.Pid }
        if ($existing.PSObject.Properties['StartTimeUtc']) { $ownerStartTimeUtc = [string]$existing.StartTimeUtc }
    }
    catch {
        return $null
    }

    if ($ownerPid -le 0) { return $null }

    $process = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
    if (-not $process -or $process.ProcessName -ne 'pwsh') { return $null }

    $commandLine = ''
    try {
        $cim = Get-CimInstance -ClassName Win32_Process -Filter ("ProcessId={0}" -f $ownerPid) -ErrorAction Stop
        $commandLine = [string]$cim.CommandLine
    }
    catch { }

    if ($commandLine -and $commandLine -notmatch 'SmartM365-Inventory-Orchestrator\\.ps1') { return $null }

    return [pscustomobject]@{
        Pid = $ownerPid
        StartTimeUtc = $ownerStartTimeUtc
        Process = $process
        CommandLine = $commandLine
    }
}

function Request-OrchestratorStop {
    param([int]$TimeoutSeconds = 180)

    $owner = Get-OrchestratorLockOwner
    if (-not $owner) {
        if (Test-Path -LiteralPath $script:Settings.StopRequestPath) {
            try { Remove-Item -LiteralPath $script:Settings.StopRequestPath -Force -ErrorAction Stop } catch { }
        }
        Write-OrchestratorLog -Message ("No live orchestrator instance found for tenant {0} on {1}; no stop request was left pending." -f $Tenant, $env:COMPUTERNAME) -Level WARN
        Write-Host ("No live orchestrator instance found for tenant {0} on {1}." -f $Tenant, $env:COMPUTERNAME) -ForegroundColor Yellow
        return 0
    }

    $payload = [pscustomobject]@{
        RequestedAtLocal = (Get-Date).ToString('o')
        RequestedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        RequestedBy = Get-OrchestratorRunUserName
        RequestedFrom = $env:COMPUTERNAME
        RequestProcessId = $PID
        TargetPid = $owner.Pid
        TargetStartTimeUtc = $owner.StartTimeUtc
        Tenant = $Tenant
        Reason = 'ManualStopRequest'
    }
    $payload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:Settings.StopRequestPath -Encoding UTF8
    Write-OrchestratorLog -Message ("Stop request written for orchestrator PID {0}: {1}" -f $owner.Pid, $script:Settings.StopRequestPath)
    Write-Host ("Stop request written for orchestrator PID {0}. Waiting up to {1} second(s)..." -f $owner.Pid, $TimeoutSeconds) -ForegroundColor Cyan

    $deadline = (Get-Date).AddSeconds([math]::Max(1, $TimeoutSeconds))
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        $stillRunning = Get-Process -Id $owner.Pid -ErrorAction SilentlyContinue
        if (-not $stillRunning -or -not (Test-Path -LiteralPath $script:Settings.LockPath)) {
            Write-OrchestratorLog -Message ("Orchestrator PID {0} stopped after manual stop request." -f $owner.Pid)
            Write-Host ("Orchestrator PID {0} stopped." -f $owner.Pid) -ForegroundColor Green
            return 0
        }
    }

    Write-OrchestratorLog -Message ("Timed out waiting for orchestrator PID {0} to stop after {1} second(s). Stop request remains at {2}." -f $owner.Pid, $TimeoutSeconds, $script:Settings.StopRequestPath) -Level ERROR
    Write-Host ("Timed out waiting for orchestrator PID {0}. Stop request remains at {1}." -f $owner.Pid, $script:Settings.StopRequestPath) -ForegroundColor Red
    return 1
}

function Test-OrchestratorStopRequested {
    if (-not (Test-Path -LiteralPath $script:Settings.StopRequestPath)) { return $false }

    $requestText = ''
    try { $requestText = Get-Content -LiteralPath $script:Settings.StopRequestPath -Raw -ErrorAction Stop }
    catch { }

    try { Remove-Item -LiteralPath $script:Settings.StopRequestPath -Force -ErrorAction Stop } catch { }

    $requestedBy = 'unknown'
    $requestedFrom = 'unknown'
    $requestedAt = 'unknown'
    try {
        $request = $requestText | ConvertFrom-Json -ErrorAction Stop
        if ($request.PSObject.Properties['RequestedBy']) { $requestedBy = [string]$request.RequestedBy }
        if ($request.PSObject.Properties['RequestedFrom']) { $requestedFrom = [string]$request.RequestedFrom }
        if ($request.PSObject.Properties['RequestedAtLocal']) { $requestedAt = [string]$request.RequestedAtLocal }
    }
    catch { }

    Write-OrchestratorLog -Message ("Manual stop request consumed. RequestedBy={0}; RequestedFrom={1}; RequestedAt={2}. No running inventory child job will be killed." -f $requestedBy, $requestedFrom, $requestedAt)
    return $true
}

function Write-OrchestratorHeartbeat {
    $now = Get-Date
    $running = @()
    foreach ($name in $script:RunningJobs.Keys) {
        $info = $script:RunningJobs[$name]
        $running += @{
            Name = $name
            Pid = $info.Process.Id
            StartTime = ConvertTo-StateTime -Value $info.StartTime
            ScheduledOccurrence = ConvertTo-StateTime -Value $info.Occurrence
        }
    }
    $heartbeat = @{
        Timestamp = $now.ToString('o')
        Pid = $PID
        Tenant = $Tenant
        ScriptVersion = $ScriptVersion
        LifetimeDeadline = ConvertTo-StateTime -Value $script:LifetimeDeadline
        RunningJobs = $running
    }
    try { Write-FileAtomically -Path $script:Settings.HeartbeatPath -Content ($heartbeat | ConvertTo-Json -Depth 5) }
    catch { Write-OrchestratorLog -Message ("Failed to write heartbeat: {0}" -f $_.Exception.Message) -Level WARN }

    $interval = [int]$script:Settings.OrchestratorHeartbeatLogIntervalMinutes
    if ($interval -le 0) { return }
    if ($script:LastHeartbeatLogTime -ne [datetime]::MinValue -and ($now - $script:LastHeartbeatLogTime).TotalMinutes -lt $interval) { return }

    $script:LastHeartbeatLogTime = $now
    $enabledJobs = 0
    if ($script:Manifest -and $script:Manifest.OrderedJobs) {
        $enabledJobs = @($script:Manifest.OrderedJobs | Where-Object { $_.Enabled }).Count
    }
    $runningNames = @($running | ForEach-Object { [string]$_.Name } | Sort-Object)
    $runningText = if ($runningNames.Count -gt 0) { $runningNames -join ', ' } else { 'none' }
    $remaining = [TimeSpan]::Zero
    if ($script:LifetimeDeadline -gt $now) { $remaining = $script:LifetimeDeadline - $now }
    Write-OrchestratorLog -Message ("Heartbeat: alive; runningJobs={0}/{1}; enabledJobs={2}; nextTickSeconds={3}; lifetimeRemaining={4}; running={5}." -f $runningNames.Count, $script:Settings.MaxConcurrency, $enabledJobs, $script:Settings.TickSeconds, $remaining.ToString('hh\:mm\:ss'), $runningText)
}

# ==========================================================
# Authenticode validation (optional, audit/enforce)
# ==========================================================
function ConvertTo-OrchestratorStringList {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @($Value -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    return @($Value | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
}

function Resolve-OrchestratorAuthenticodeCertificatePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path -Path $PSScriptRoot -ChildPath $Path))
}

function Install-OrchestratorAuthenticodeTrustedCertificates {
    if (-not $script:Settings.AuthenticodeValidationEnabled) { return }
    if (-not $script:Settings.AuthenticodeInstallTrustedCertificates) { return }

    $certificatePaths = @(ConvertTo-OrchestratorStringList -Value $script:Settings.AuthenticodeTrustedCertificatePaths)
    if ($certificatePaths.Count -eq 0) {
        Write-OrchestratorLog -Message 'Authenticode trust install enabled but no AuthenticodeTrustedCertificatePaths are configured.' -Level WARN
        return
    }

    $targetStores = @()
    if ($script:Settings.AuthenticodeInstallTrustedRoot) { $targetStores += 'Root' }
    if ($script:Settings.AuthenticodeInstallTrustedPublisher) { $targetStores += 'TrustedPublisher' }
    if ($targetStores.Count -eq 0) {
        Write-OrchestratorLog -Message 'Authenticode trust install enabled but both target stores are disabled.' -Level WARN
        return
    }

    foreach ($configuredPath in $certificatePaths) {
        $certPath = Resolve-OrchestratorAuthenticodeCertificatePath -Path $configuredPath
        if (-not (Test-Path -LiteralPath $certPath)) {
            $message = "Authenticode trusted certificate file not found: $certPath"
            $level = if ($script:Settings.AuthenticodeValidationMode -eq 'Enforce') { 'ERROR' } else { 'WARN' }
            Write-OrchestratorLog -Message $message -Level $level
            if ($script:Settings.AuthenticodeValidationMode -eq 'Enforce') { throw $message }
            continue
        }

        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certPath)
        $thumbprint = ([string]$cert.Thumbprint).Replace(' ', '').ToUpperInvariant()
        $allowedThumbprints = @($script:Settings.AuthenticodeAllowedThumbprints)
        if ($allowedThumbprints.Count -gt 0 -and $allowedThumbprints -notcontains $thumbprint) {
            $message = "Authenticode trusted certificate skipped because thumbprint $thumbprint is not in AuthenticodeAllowedThumbprints: $certPath"
            $level = if ($script:Settings.AuthenticodeValidationMode -eq 'Enforce') { 'ERROR' } else { 'WARN' }
            Write-OrchestratorLog -Message $message -Level $level
            if ($script:Settings.AuthenticodeValidationMode -eq 'Enforce') { throw $message }
            continue
        }

        foreach ($storeName in $targetStores) {
            $store = [System.Security.Cryptography.X509Certificates.X509Store]::new($storeName, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
            try {
                $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
                $existing = @($store.Certificates | Where-Object { ([string]$_.Thumbprint).Replace(' ', '').ToUpperInvariant() -eq $thumbprint })
                if ($existing.Count -gt 0) {
                    Write-OrchestratorLog -Message ("Authenticode trusted certificate already present: store=CurrentUser\{0}; thumbprint={1}; subject={2}" -f $storeName, $thumbprint, $cert.Subject)
                }
                else {
                    $store.Add($cert)
                    Write-OrchestratorLog -Message ("Authenticode trusted certificate installed: store=CurrentUser\{0}; thumbprint={1}; subject={2}; source={3}" -f $storeName, $thumbprint, $cert.Subject, $certPath)
                }
            }
            catch {
                $message = "Authenticode trusted certificate install failed: store=CurrentUser\$storeName; thumbprint=$thumbprint; error=$($_.Exception.Message)"
                $level = if ($script:Settings.AuthenticodeValidationMode -eq 'Enforce') { 'ERROR' } else { 'WARN' }
                Write-OrchestratorLog -Message $message -Level $level
                if ($script:Settings.AuthenticodeValidationMode -eq 'Enforce') { throw $message }
            }
            finally {
                if ($store) { $store.Close() }
            }
        }
    }
}

function Get-OrchestratorAuthenticodeFileList {
    param(
        [Parameter(Mandatory = $true)]$Job,
        [Parameter(Mandatory = $true)][string]$ScriptFullPath
    )

    $items = New-Object System.Collections.Generic.List[object]
    $items.Add([pscustomobject]@{ Role = 'JobScript'; Path = $ScriptFullPath })

    if ($script:Settings.AuthenticodeCheckCoreModule) {
        $coreManifest = [string]$script:Settings.CoreModulePath
        if (-not [string]::IsNullOrWhiteSpace($coreManifest)) {
            $items.Add([pscustomobject]@{ Role = 'SmartM365.Core manifest'; Path = $coreManifest })
            $coreScript = Join-Path -Path (Split-Path -Path $coreManifest -Parent) -ChildPath 'SmartM365.Core.psm1'
            $items.Add([pscustomobject]@{ Role = 'SmartM365.Core module'; Path = $coreScript })
        }
    }

    if ($script:Settings.AuthenticodeCheckWindowsPowerShellModule -and $Job.PowerShellEdition -eq 'WindowsPowerShell') {
        $coreRoot = Split-Path -Path ([string]$script:Settings.CoreModulePath) -Parent
        $ps5Root = Join-Path -Path $coreRoot -ChildPath 'Compatibility\WindowsPowerShell5'
        foreach ($name in @('SmartM365-WindowsPowerShell5.psd1', 'SmartM365-WindowsPowerShell5.psm1')) {
            $candidate = Join-Path -Path $ps5Root -ChildPath $name
            if (Test-Path -LiteralPath $candidate) {
                $items.Add([pscustomobject]@{ Role = 'SmartM365 WindowsPowerShell5 compatibility'; Path = $candidate })
            }
        }
    }

    return @($items)
}

function Test-OrchestratorAuthenticodeFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Role
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{
            Role = $Role
            Path = $Path
            Status = 'Missing'
            Thumbprint = ''
            Subject = ''
            Signer = ''
            IsValid = $false
            Detail = 'File not found.'
        }
    }

    try {
        $signature = Get-AuthenticodeSignature -FilePath $Path -ErrorAction Stop
        $thumbprint = ''
        $subject = ''
        if ($signature.SignerCertificate) {
            $thumbprint = [string]$signature.SignerCertificate.Thumbprint
            $subject = [string]$signature.SignerCertificate.Subject
        }

        $allowedThumbprints = @($script:Settings.AuthenticodeAllowedThumbprints)
        $thumbprintAllowed = $true
        if ($allowedThumbprints.Count -gt 0) {
            $thumbprintAllowed = $false
            if (-not [string]::IsNullOrWhiteSpace($thumbprint)) {
                $thumbprintAllowed = ($allowedThumbprints -contains $thumbprint.ToUpperInvariant())
            }
        }

        $isValid = ([string]$signature.Status -eq 'Valid') -and $thumbprintAllowed
        $detail = [string]$signature.StatusMessage
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = [string]$signature.Status }
        if ([string]$signature.Status -eq 'Valid' -and -not $thumbprintAllowed) {
            $detail = 'Signature is valid, but the signer thumbprint is not in AuthenticodeAllowedThumbprints.'
        }

        return [pscustomobject]@{
            Role = $Role
            Path = $Path
            Status = [string]$signature.Status
            Thumbprint = $thumbprint
            Subject = $subject
            Signer = $subject
            IsValid = $isValid
            Detail = $detail
        }
    }
    catch {
        return [pscustomobject]@{
            Role = $Role
            Path = $Path
            Status = 'Error'
            Thumbprint = ''
            Subject = ''
            Signer = ''
            IsValid = $false
            Detail = $_.Exception.Message
        }
    }
}

function New-OrchestratorAuthenticodeHtmlReport {
    param(
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)]$Results
    )

    $rows = @()
    foreach ($result in $Results) {
        $color = if ($result.IsValid) { '#107C10' } else { '#D13438' }
        $rows += "<tr><td style='padding:4px 8px;border:1px solid #DDDDDD;'>$(ConvertTo-HtmlText -Text $result.Role)</td><td style='padding:4px 8px;border:1px solid #DDDDDD;color:$color;'>$(ConvertTo-HtmlText -Text $result.Status)</td><td style='padding:4px 8px;border:1px solid #DDDDDD;'>$(ConvertTo-HtmlText -Text $result.Thumbprint)</td><td style='padding:4px 8px;border:1px solid #DDDDDD;'>$(ConvertTo-HtmlText -Text $result.Path)</td><td style='padding:4px 8px;border:1px solid #DDDDDD;'>$(ConvertTo-HtmlText -Text $result.Detail)</td></tr>"
    }

    return @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#1F2937;'>
<h2 style='color:#D13438;'>Authenticode validation rejected job $(ConvertTo-HtmlText -Text $JobName)</h2>
<table style='border-collapse:collapse;'>
<tr><th style='padding:4px 8px;border:1px solid #DDDDDD;'>Role</th><th style='padding:4px 8px;border:1px solid #DDDDDD;'>Status</th><th style='padding:4px 8px;border:1px solid #DDDDDD;'>Thumbprint</th><th style='padding:4px 8px;border:1px solid #DDDDDD;'>Path</th><th style='padding:4px 8px;border:1px solid #DDDDDD;'>Detail</th></tr>
$($rows -join "`n")
</table>
<p>Mode: $(ConvertTo-HtmlText -Text $script:Settings.AuthenticodeValidationMode). Tenant: $(ConvertTo-HtmlText -Text $Tenant). Server: $(ConvertTo-HtmlText -Text $env:COMPUTERNAME).</p>
</body></html>
"@
}

function Invoke-OrchestratorAuthenticodeValidation {
    param(
        [Parameter(Mandatory = $true)]$Job,
        [Parameter(Mandatory = $true)][string]$ScriptFullPath
    )

    if (-not $script:Settings.AuthenticodeValidationEnabled) {
        return [pscustomobject]@{ Allowed = $true; Results = @(); Summary = 'Authenticode validation disabled.' }
    }

    $results = @()
    foreach ($item in (Get-OrchestratorAuthenticodeFileList -Job $Job -ScriptFullPath $ScriptFullPath)) {
        $result = Test-OrchestratorAuthenticodeFile -Path $item.Path -Role $item.Role
        $results += $result
        if ($result.IsValid) {
            Write-OrchestratorLog -Message ("Authenticode {0}: valid; thumbprint={1}; subject={2}; path={3}" -f $result.Role, $result.Thumbprint, $result.Subject, $result.Path)
        }
        else {
            $level = if ($script:Settings.AuthenticodeValidationMode -eq 'Enforce') { 'ERROR' } else { 'WARN' }
            Write-OrchestratorLog -Message ("Authenticode {0}: {1}; thumbprint={2}; path={3}; detail={4}" -f $result.Role, $result.Status, $result.Thumbprint, $result.Path, $result.Detail) -Level $level
        }
    }

    $failed = @($results | Where-Object { -not $_.IsValid })
    if ($failed.Count -eq 0) {
        return [pscustomobject]@{ Allowed = $true; Results = $results; Summary = 'Authenticode validation passed.' }
    }

    $summary = "Authenticode validation found $($failed.Count) invalid or untrusted file(s). Mode=$($script:Settings.AuthenticodeValidationMode)."
    if ($script:Settings.AuthenticodeValidationMode -eq 'Enforce') {
        $body = New-OrchestratorAuthenticodeHtmlReport -JobName $Job.Name -Results $results
        Send-OrchestratorMail -Subject ("[CRITICAL][SmartM365 Orchestrator][$Tenant] Authenticode rejected job $($Job.Name)") -HtmlBody $body -IsError | Out-Null
        return [pscustomobject]@{ Allowed = $false; Results = $results; Summary = $summary }
    }

    return [pscustomobject]@{ Allowed = $true; Results = $results; Summary = $summary }
}

# Child process management
# ==========================================================
function Get-PwshPath {
    $candidate = Join-Path -Path $PSHOME -ChildPath 'pwsh.exe'
    if (Test-Path -LiteralPath $candidate) { return $candidate }
    $command = Get-Command -Name pwsh.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw 'pwsh.exe not found.'
}

function Get-JobEngine {
    # Returns the child-process engine for a job: pwsh (default) or Windows
    # PowerShell 5.1 for jobs that declare PowerShellEdition = WindowsPowerShell
    # (Exchange on-premises scripts).
    param([AllowNull()]$Job)

    if ($null -ne $Job -and $Job.PowerShellEdition -eq 'WindowsPowerShell') {
        $path = Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'
        if (-not (Test-Path -LiteralPath $path)) { throw 'powershell.exe (Windows PowerShell 5.1) not found.' }
        return [pscustomobject]@{ Path = $path; ProcessName = 'powershell' }
    }
    return [pscustomobject]@{ Path = (Get-PwshPath); ProcessName = 'pwsh' }
}

function Resolve-OrchestratorJobPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = [Environment]::ExpandEnvironmentVariables($Path)
    $resolved = $resolved.Replace('{{Tenant}}', $Tenant).Replace('{{TenantKey}}', $Tenant)
    if ([System.IO.Path]::IsPathRooted($resolved)) { return $resolved }
    return (Join-Path -Path $script:SmartInventoryRoot -ChildPath $resolved)
}
function Stop-ProcessTree {
    param([Parameter(Mandatory = $true)][int]$TargetPid)

    try {
        $output = & taskkill.exe /PID $TargetPid /T /F 2>&1
        Write-OrchestratorLog -Message ("taskkill /T /F on PID {0}: {1}" -f $TargetPid, (($output | ForEach-Object { [string]$_ }) -join ' | '))
    }
    catch {
        Write-OrchestratorLog -Message ("Failed to kill process tree of PID {0}: {1}" -f $TargetPid, $_.Exception.Message) -Level WARN
    }
}

function Test-ProcessMatchesRecord {
    param(
        [Parameter(Mandatory = $true)][int]$RecordedPid,
        [Parameter(Mandatory = $true)][datetime]$ExpectedStartTime,
        [string]$ExpectedProcessName = 'pwsh'
    )

    $process = Get-Process -Id $RecordedPid -ErrorAction SilentlyContinue
    if (-not $process) { return $null }
    if ($process.ProcessName -ne $ExpectedProcessName) { return $null }
    $actualStart = $null
    try { $actualStart = $process.StartTime } catch { return $null }
    if ([math]::Abs(($actualStart - $ExpectedStartTime).TotalSeconds) -gt 5) { return $null }
    # Pin the process handle now so the exit code stays readable after the process ends.
    try { $null = $process.Handle } catch { }
    return $process
}

function Start-InventoryJob {
    param(
        [Parameter(Mandatory = $true)]$Job,
        [Parameter(Mandatory = $true)][datetime]$Occurrence,
        [int]$Attempt = 0
    )

    $state = Get-JobState -JobName $Job.Name
    $scriptFullPath = Resolve-OrchestratorJobPath -Path $Job.ScriptPath
    $launcherFullPath = ''
    if ($Job.PSObject.Properties['LauncherPath'] -and -not [string]::IsNullOrWhiteSpace([string]$Job.LauncherPath)) {
        $launcherFullPath = Resolve-OrchestratorJobPath -Path ([string]$Job.LauncherPath)
    }
    $useLauncher = -not [string]::IsNullOrWhiteSpace($launcherFullPath)
    $startTime = Get-Date
    $logFolder = $script:Settings.JobLogFolderPath
    if (-not (Test-Path -LiteralPath $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
    $logPath = Join-Path -Path $logFolder -ChildPath ("Job-{0}_{1}_{2}.log" -f $Job.Name, $env:COMPUTERNAME, $startTime.ToString('yyyyMMdd_HHmmss'))

    if (-not (Test-Path -LiteralPath $scriptFullPath)) {
        Write-OrchestratorLog -Message ("Job {0}: script not found: {1}" -f $Job.Name, $scriptFullPath) -Level ERROR
        $runInfo = @{ StartTime = $startTime; Occurrence = $Occurrence; LogPath = ''; Attempt = $Attempt; TimeoutMinutes = $Job.TimeoutMinutes }
        Complete-JobRun -JobName $Job.Name -RunInfo $runInfo -StatusHint 'LaunchFailed' -ExitCode $null -EndTime $startTime -ErrorText ("Script not found: {0}" -f $scriptFullPath)
        return
    }

    if ($useLauncher -and -not (Test-Path -LiteralPath $launcherFullPath)) {
        Write-OrchestratorLog -Message ("Job {0}: launcher not found: {1}" -f $Job.Name, $launcherFullPath) -Level ERROR
        $runInfo = @{ StartTime = $startTime; Occurrence = $Occurrence; LogPath = ''; Attempt = $Attempt; TimeoutMinutes = $Job.TimeoutMinutes }
        Complete-JobRun -JobName $Job.Name -RunInfo $runInfo -StatusHint 'LaunchFailed' -ExitCode $null -EndTime $startTime -ErrorText ("Launcher not found: {0}" -f $launcherFullPath)
        return
    }
    $authenticodeResult = Invoke-OrchestratorAuthenticodeValidation -Job $Job -ScriptFullPath $scriptFullPath
    if (-not $authenticodeResult.Allowed) {
        Write-OrchestratorLog -Message ("Job {0}: launch rejected by Authenticode validation. {1}" -f $Job.Name, $authenticodeResult.Summary) -Level ERROR
        $rejectLine = "[{0}] {1} rejected job {2}: AuthenticodeRejected. {3}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $ScriptName, $Job.Name, $authenticodeResult.Summary
        try { [System.IO.File]::WriteAllText($logPath, $rejectLine + [Environment]::NewLine) } catch { }
        $runInfo = @{ StartTime = $startTime; Occurrence = $Occurrence; LogPath = $logPath; Attempt = $Attempt; TimeoutMinutes = $Job.TimeoutMinutes }
        Complete-JobRun -JobName $Job.Name -RunInfo $runInfo -StatusHint 'LaunchFailed' -ExitCode $null -EndTime (Get-Date) -ErrorText ("AuthenticodeRejected: {0}" -f $authenticodeResult.Summary)
        return
    }

    $escapedScript = $scriptFullPath.Replace("'", "''")
    $escapedLog = $logPath.Replace("'", "''")
    $escapedTenant = $Tenant.Replace("'", "''")
    $argumentPart = ''
    if (-not [string]::IsNullOrWhiteSpace($Job.Arguments)) { $argumentPart = ' ' + $Job.Arguments.Trim() }
    $connectPart = ''
    if ($Connect) { $connectPart = ' -Connect' }
    $engine = Get-JobEngine -Job $Job
    # Out-File:Encoding pins the *>> redirection to UTF-8: Windows PowerShell 5.1
    # would otherwise append UTF-16 output to the UTF-8 job log.
    if ($useLauncher) {
        $cmdLine = 'call "' + $launcherFullPath.Replace('"', '""') + '"' + $argumentPart
        $escapedCmdLine = $cmdLine.Replace("'", "''")
        $command = "`$ErrorActionPreference = 'Continue'; " +
            "`$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'; " +
            "& `$env:ComSpec /d /c '" + $escapedCmdLine + "' *>> '" + $escapedLog + "'; " +
            "if (-not `$?) { exit 1 } elseif (`$null -ne `$LASTEXITCODE) { exit `$LASTEXITCODE } else { exit 0 }"
    }
    else {
        $command = "`$ErrorActionPreference = 'Continue'; " +
            "`$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'; " +
            "& '" + $escapedScript + "'" + $argumentPart + " -Tenant '" + $escapedTenant + "'" + $connectPart + " *>> '" + $escapedLog + "'; " +
            "if (-not `$?) { exit 1 } elseif (`$null -ne `$LASTEXITCODE) { exit `$LASTEXITCODE } else { exit 0 }"
    }
    $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($command))

    $headerLines = @(
        ("[{0}] {1} launching job {2} (attempt {3}, scheduled {4})" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss'), $ScriptName, $Job.Name, $Attempt, $Occurrence.ToString('yyyy-MM-dd HH:mm:ss')),
        ("[{0}] Script path verified: {1}" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss'), $scriptFullPath),
        ("[{0}] Launch mode: {1}" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss'), $(if ($useLauncher) { 'LauncherPath' } else { 'ScriptPath' })),
        ("[{0}] Launcher path: {1}" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss'), $(if ($useLauncher) { $launcherFullPath } else { 'n/a' })),
        ("[{0}] PowerShell engine: {1} ({2})" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss'), $engine.Path, $Job.PowerShellEdition),
        ("[{0}] Child command: {1}" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss'), $command),
        ("[{0}] Start-Process arguments: -NoProfile -ExecutionPolicy Bypass -EncodedCommand <encoded; decoded child command logged above>" -f $startTime.ToString('yyyy-MM-dd HH:mm:ss'))
    )
    try { [System.IO.File]::WriteAllLines($logPath, $headerLines, [System.Text.UTF8Encoding]::new($false)) } catch { }

    try {
        $process = Start-Process -FilePath $engine.Path -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encodedCommand) -WindowStyle Hidden -PassThru
    }
    catch {
        Write-OrchestratorLog -Message ("Job {0}: failed to start child process: {1}" -f $Job.Name, $_.Exception.Message) -Level ERROR
        $runInfo = @{ StartTime = $startTime; Occurrence = $Occurrence; LogPath = $logPath; Attempt = $Attempt; TimeoutMinutes = $Job.TimeoutMinutes }
        Complete-JobRun -JobName $Job.Name -RunInfo $runInfo -StatusHint 'LaunchFailed' -ExitCode $null -EndTime (Get-Date) -ErrorText $_.Exception.Message
        return
    }

    $processStart = $startTime
    try { $processStart = $process.StartTime } catch { }
    try { $null = $process.Handle } catch { }

    $script:RunningJobs[$Job.Name] = @{
        Process = $process
        StartTime = $processStart
        Occurrence = $Occurrence
        LogPath = $logPath
        Attempt = $Attempt
        TimeoutMinutes = $Job.TimeoutMinutes
        ProcessName = $engine.ProcessName
    }
    $state.Running = @{
        Pid = $process.Id
        StartTime = ConvertTo-StateTime -Value $processStart
        ScheduledOccurrence = ConvertTo-StateTime -Value $Occurrence
        LogPath = $logPath
        Attempt = $Attempt
        TimeoutMinutes = $Job.TimeoutMinutes
        ProcessName = $engine.ProcessName
    }
    $state.LastRunStart = ConvertTo-StateTime -Value $processStart
    $state.LastStatus = 'Running'
    $state.RetryCount = $Attempt
    $state.PendingRetry = $null
    Save-OrchestratorState

    $launchTarget = if ($useLauncher) { $launcherFullPath } else { $scriptFullPath }
    Write-OrchestratorLog -Message ("Job {0}: started PID {1} ({2}, attempt {3}, scheduled {4}, timeout {5} min, log {6}, target {7}; command {8})." -f $Job.Name, $process.Id, $engine.ProcessName, $Attempt, $Occurrence.ToString('yyyy-MM-dd HH:mm'), $Job.TimeoutMinutes, $logPath, $launchTarget, $command)
}

function Test-JobRunSuccessEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$JobName,
        [AllowNull()][string]$LogPath,
        [int]$DurationSec,
        [AllowNull()]$ManifestJob
    )

    $issues = New-Object System.Collections.Generic.List[string]
    if ($null -eq $ManifestJob) { return [pscustomobject]@{ IsValid = $true; Message = '' } }

    $minimumDuration = 0
    if ($ManifestJob.PSObject.Properties['MinimumSuccessDurationSeconds'] -and $null -ne $ManifestJob.MinimumSuccessDurationSeconds) {
        $minimumDuration = [int]$ManifestJob.MinimumSuccessDurationSeconds
    }
    if ($minimumDuration -gt 0 -and $DurationSec -lt $minimumDuration) {
        $issues.Add(("duration {0}s is lower than required minimum {1}s" -f $DurationSec, $minimumDuration))
    }

    $patterns = @()
    if ($ManifestJob.PSObject.Properties['RequiredLogPatterns'] -and $null -ne $ManifestJob.RequiredLogPatterns) {
        $patterns = @($ManifestJob.RequiredLogPatterns | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    }
    if ($patterns.Count -gt 0) {
        if ([string]::IsNullOrWhiteSpace($LogPath) -or -not (Test-Path -LiteralPath $LogPath)) {
            $issues.Add('job log is missing; required success evidence cannot be verified')
        }
        else {
            $logText = ''
            try { $logText = [System.IO.File]::ReadAllText($LogPath) }
            catch { $issues.Add(("job log cannot be read: {0}" -f $_.Exception.Message)) }
            if (-not [string]::IsNullOrEmpty($logText)) {
                foreach ($pattern in $patterns) {
                    if ($logText -notmatch $pattern) {
                        $issues.Add(("job log does not contain required success evidence pattern: {0}" -f $pattern))
                    }
                }
            }
            elseif ($issues.Count -eq 0) {
                $issues.Add('job log is empty; required success evidence cannot be verified')
            }
        }
    }

    if ($issues.Count -gt 0) {
        return [pscustomobject]@{ IsValid = $false; Message = ($issues -join '; ') }
    }
    return [pscustomobject]@{ IsValid = $true; Message = '' }
}

function Complete-JobRun {
    param(
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][hashtable]$RunInfo,
        [Parameter(Mandatory = $true)][ValidateSet('Exited', 'TimedOut', 'Interrupted', 'LaunchFailed')][string]$StatusHint,
        [AllowNull()]$ExitCode,
        [Parameter(Mandatory = $true)][datetime]$EndTime,
        [string]$ErrorText = ''
    )

    $state = Get-JobState -JobName $JobName
    $manifestJob = $null
    if ($script:Manifest -and $script:Manifest.JobsByName.ContainsKey($JobName)) { $manifestJob = $script:Manifest.JobsByName[$JobName] }

    $status = 'Failed'
    switch ($StatusHint) {
        'TimedOut' { $status = 'TimedOut' }
        'Interrupted' { $status = 'Interrupted' }
        'LaunchFailed' { $status = 'Failed' }
        'Exited' {
            if ($null -eq $ExitCode) {
                $status = 'Failed'
                Write-OrchestratorLog -Message ("Job {0}: exit code could not be read; the run is treated as Failed." -f $JobName) -Level ERROR
            }
            elseif ([int]$ExitCode -eq 0) { $status = 'Success' }
            else { $status = 'Failed' }
        }
    }

    $durationSec = [int][math]::Max(0, ($EndTime - [datetime]$RunInfo.StartTime).TotalSeconds)
    $attempt = [int]$RunInfo.Attempt

    if ($status -eq 'Success') {
        $evidence = Test-JobRunSuccessEvidence -JobName $JobName -LogPath ([string]$RunInfo.LogPath) -DurationSec $durationSec -ManifestJob $manifestJob
        if (-not $evidence.IsValid) {
            $status = 'Failed'
            $message = "Job success evidence validation failed: $($evidence.Message)"
            if ([string]::IsNullOrWhiteSpace($ErrorText)) { $ErrorText = $message } else { $ErrorText = $ErrorText + '; ' + $message }
            Write-OrchestratorLog -Message ("Job {0}: {1}" -f $JobName, $message) -Level ERROR
        }
    }

    $retryScheduled = $false
    if ($status -ne 'Success' -and $null -ne $manifestJob -and $attempt -lt $manifestJob.MaxRetries) {
        $notBefore = $EndTime.AddSeconds($manifestJob.RetryDelaySeconds)
        $state.PendingRetry = @{
            NotBefore = ConvertTo-StateTime -Value $notBefore
            Attempt = $attempt + 1
            ScheduledOccurrence = ConvertTo-StateTime -Value ([datetime]$RunInfo.Occurrence)
        }
        $retryScheduled = $true
    }

    $state.LastRunEnd = ConvertTo-StateTime -Value $EndTime
    $state.LastExitCode = $ExitCode
    $state.LastStatus = $status
    $state.Running = $null
    if ($script:RunningJobs.ContainsKey($JobName)) { $script:RunningJobs.Remove($JobName) }
    Save-OrchestratorState
    $csvStatus = $status
    if ($retryScheduled) { $csvStatus = 'Retried' }
    Add-JobRunCsvRow -JobName $JobName -ScheduledTime $RunInfo.Occurrence -StartTime $RunInfo.StartTime -EndTime $EndTime -DurationSec $durationSec -ExitCode $ExitCode -Status $csvStatus -RetryCount $attempt -LogPath ([string]$RunInfo.LogPath)


    Invoke-OrchestratorSharePointUpload -LocalFilePath ([string]$RunInfo.LogPath) -Reason ("job {0} log" -f $JobName) -Force | Out-Null
    Invoke-OrchestratorSharePointUpload -LocalFilePath (Get-JobRunsCsvPath) -Reason 'job-runs CSV' -Force | Out-Null
    $exitText = 'n/a'
    if ($null -ne $ExitCode) { $exitText = [string]$ExitCode }
    $level = 'INFO'
    if ($status -ne 'Success') { $level = 'ERROR' }
    Write-OrchestratorLog -Message ("Job {0}: finished with status {1} (exit code {2}, duration {3}s, attempt {4})." -f $JobName, $status, $exitText, $durationSec, $attempt) -Level $level
    if ($retryScheduled) {
        Write-OrchestratorLog -Message ("Job {0}: retry {1}/{2} scheduled after {3}s." -f $JobName, ($attempt + 1), $manifestJob.MaxRetries, $manifestJob.RetryDelaySeconds) -Level WARN
    }

    $jobMailMode = $script:Settings.JobMailMode
    $shouldEmail = $false
    if ($jobMailMode -eq 'Always' -and -not $retryScheduled) { $shouldEmail = $true }
    elseif ($jobMailMode -eq 'OnError' -and $status -ne 'Success' -and -not $retryScheduled) { $shouldEmail = $true }
    if ($shouldEmail) {
        $details = [ordered]@{
            'Job' = $JobName
            'Status' = $status
            'Scheduled time' = ([datetime]$RunInfo.Occurrence).ToString('yyyy-MM-dd HH:mm:ss')
            'Start time' = ([datetime]$RunInfo.StartTime).ToString('yyyy-MM-dd HH:mm:ss')
            'End time' = $EndTime.ToString('yyyy-MM-dd HH:mm:ss')
            'Duration (s)' = $durationSec
            'Exit code' = $exitText
            'Attempt' = $attempt
            'Log path' = [string]$RunInfo.LogPath
        }
        if (-not [string]::IsNullOrWhiteSpace($ErrorText)) { $details['Error'] = $ErrorText }
        $detailsHashtable = @{}
        foreach ($key in $details.Keys) { $detailsHashtable[$key] = $details[$key] }
        Send-JobResultEmail -JobName $JobName -Status $status -Details $detailsHashtable -LogPath ([string]$RunInfo.LogPath)
    }
}

function Update-RunningJobs {
    param([Parameter(Mandatory = $true)][datetime]$Now)

    foreach ($name in @($script:RunningJobs.Keys)) {
        $info = $script:RunningJobs[$name]
        $process = $info.Process
        $exited = $false
        try {
            $process.Refresh()
            $exited = $process.HasExited
        }
        catch { $exited = $true }

        if ($exited) {
            $exitCode = $null
            try { $exitCode = $process.ExitCode } catch { }
            $endTime = $Now
            try { $endTime = $process.ExitTime } catch { }
            Complete-JobRun -JobName $name -RunInfo $info -StatusHint 'Exited' -ExitCode $exitCode -EndTime $endTime
            continue
        }

        $timeoutMinutes = [int]$info.TimeoutMinutes
        if ($script:Manifest.JobsByName.ContainsKey($name)) { $timeoutMinutes = [int]$script:Manifest.JobsByName[$name].TimeoutMinutes }
        if ($timeoutMinutes -gt 0 -and ($Now - [datetime]$info.StartTime).TotalMinutes -gt $timeoutMinutes) {
            Write-OrchestratorLog -Message ("Job {0}: timeout after {1} minutes (limit {2}); killing the process tree of PID {3}." -f $name, [int]($Now - [datetime]$info.StartTime).TotalMinutes, $timeoutMinutes, $process.Id) -Level ERROR
            Stop-ProcessTree -TargetPid $process.Id
            try { $process.WaitForExit(10000) | Out-Null } catch { }
            Complete-JobRun -JobName $name -RunInfo $info -StatusHint 'TimedOut' -ExitCode $null -EndTime (Get-Date)
        }
    }
}

function Restore-RunningJobs {
    foreach ($jobName in @($script:State.Jobs.Keys)) {
        $state = $script:State.Jobs[$jobName]
        if ($null -eq $state.Running) { continue }
        $record = $state.Running
        $expectedStart = ConvertFrom-StateTime -Text ([string]$record.StartTime)
        $recordedPid = [int]$record.Pid
        $expectedProcessName = 'pwsh'
        if ($record.ContainsKey('ProcessName') -and $record.ProcessName) { $expectedProcessName = [string]$record.ProcessName }
        $process = $null
        if ($null -ne $expectedStart) { $process = Test-ProcessMatchesRecord -RecordedPid $recordedPid -ExpectedStartTime $expectedStart -ExpectedProcessName $expectedProcessName }

        $runInfo = @{
            StartTime = $expectedStart
            Occurrence = ConvertFrom-StateTime -Text ([string]$record.ScheduledOccurrence)
            LogPath = [string]$record.LogPath
            Attempt = [int]$record.Attempt
            TimeoutMinutes = [int]$record.TimeoutMinutes
        }
        if ($null -eq $runInfo.StartTime) { $runInfo.StartTime = Get-Date }
        if ($null -eq $runInfo.Occurrence) { $runInfo.Occurrence = $runInfo.StartTime }

        if ($null -ne $process) {
            $runInfo['Process'] = $process
            $script:RunningJobs[$jobName] = $runInfo
            Write-OrchestratorLog -Message ("Job {0}: re-adopted running PID {1} (started {2}); supervision resumed with the original timeout window." -f $jobName, $recordedPid, $expectedStart.ToString('yyyy-MM-dd HH:mm:ss'))
        }
        else {
            Write-OrchestratorLog -Message ("Job {0}: recorded PID {1} is gone or does not match; the run is marked Interrupted." -f $jobName, $recordedPid) -Level WARN
            Complete-JobRun -JobName $jobName -RunInfo $runInfo -StatusHint 'Interrupted' -ExitCode $null -EndTime (Get-Date) -ErrorText 'The orchestrator restart found no matching pwsh process for this run.'
        }
    }
}

# ==========================================================
# Launch phase (due occurrences, retries, forced, dependencies)
# ==========================================================
function Test-JobSelected {
    param([Parameter(Mandatory = $true)][string]$JobName)

    if ($Skip.Count -gt 0 -and $Skip -contains $JobName) { return $false }
    if ($Only.Count -gt 0 -and $Only -notcontains $JobName) { return $false }
    return $true
}

function Get-JobEffectiveAllowedServers {
    param([Parameter(Mandatory = $true)]$Job)

    $list = @($Job.AllowedServers)
    if ($list.Count -eq 0) { $list = @($script:Settings.DefaultAllowedServers) }
    return $list
}

function Test-JobAllowedOnServer {
    # Empty effective list = every server is allowed. Comparison uses the local
    # computer name (case-insensitive).
    param([Parameter(Mandatory = $true)]$Job)

    $list = Get-JobEffectiveAllowedServers -Job $Job
    if ($list.Count -eq 0) { return $true }
    return ($list -contains $env:COMPUTERNAME)
}

function Write-ServerAllowlistSummary {
    $blocked = @($script:Manifest.OrderedJobs | Where-Object { -not (Test-JobAllowedOnServer -Job $_) } | ForEach-Object { $_.Name })
    if ($blocked.Count -gt 0) {
        Write-OrchestratorLog -Message ("Jobs not allowed on this server ({0}): {1}" -f $env:COMPUTERNAME, ($blocked -join ', '))
    }
}

function Set-OccurrenceSkipped {
    param(
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][datetime]$Occurrence,
        [Parameter(Mandatory = $true)][string]$Reason
    )

    $state = Get-JobState -JobName $JobName
    $state.LastScheduledOccurrence = ConvertTo-StateTime -Value $Occurrence
    $state.LastStatus = 'Skipped'
    Save-OrchestratorState

    Add-JobRunCsvRow -JobName $JobName -ScheduledTime $Occurrence -StartTime $null -EndTime $null -DurationSec $null -ExitCode $null -Status 'Skipped' -RetryCount 0 -LogPath ''
    Write-OrchestratorLog -Message ("Job {0}: occurrence {1} skipped ({2})." -f $JobName, $Occurrence.ToString('yyyy-MM-dd HH:mm'), $Reason) -Level WARN
}

function Set-OccurrenceBlocked {
    param(
        [Parameter(Mandatory = $true)][string]$JobName,
        [Parameter(Mandatory = $true)][datetime]$Occurrence,
        [Parameter(Mandatory = $true)][ValidateSet('BlockedDependencyFailed', 'BlockedDependencyTimeout')][string]$Status,
        [Parameter(Mandatory = $true)][string]$Reason
    )

    $state = Get-JobState -JobName $JobName
    $state.LastScheduledOccurrence = ConvertTo-StateTime -Value $Occurrence
    $state.LastStatus = $Status
    $state.PendingRetry = $null
    Save-OrchestratorState

    Add-JobRunCsvRow -JobName $JobName -ScheduledTime $Occurrence -StartTime $null -EndTime (Get-Date) -DurationSec 0 -ExitCode $null -Status $Status -RetryCount 0 -LogPath ''
    Write-OrchestratorLog -Message ("Job {0}: occurrence {1} blocked with status {2} ({3})." -f $JobName, $Occurrence.ToString('yyyy-MM-dd HH:mm'), $Status, $Reason) -Level ERROR
}
function Initialize-NewJobStates {
    # New jobs (no state yet) start fast-forwarded: an occurrence from before the job
    # existed is not a missed occurrence.
    $now = Get-Date
    foreach ($job in $script:Manifest.OrderedJobs) {
        if ($script:State.Jobs.ContainsKey($job.Name)) { continue }
        $state = Get-JobState -JobName $job.Name
        $latest = Get-LatestPastOccurrence -Job $job -Now $now
        if ($null -ne $latest) { $state.LastScheduledOccurrence = ConvertTo-StateTime -Value $latest }
        Write-OrchestratorLog -Message ("Job {0}: initialized state (fast-forwarded to {1})." -f $job.Name, $(if ($null -ne $latest) { $latest.ToString('yyyy-MM-dd HH:mm') } else { 'no past occurrence' }))
    }
    Save-OrchestratorState

}

function Invoke-MissedRunCatchUp {
    # Startup only: apply MissedRunPolicy=Skip by fast-forwarding past occurrences.
    # RunOnce jobs keep their due occurrence and are caught up by the launch phase.
    $now = Get-Date
    foreach ($job in $script:Manifest.OrderedJobs) {
        if (-not $job.Enabled) { continue }
        if (-not (Test-JobAllowedOnServer -Job $job)) { continue }
        if ($job.Schedule.MissedRunPolicy -ne 'Skip') { continue }
        $state = Get-JobState -JobName $job.Name
        if ($script:RunningJobs.ContainsKey($job.Name)) { continue }
        $lastOccurrence = ConvertFrom-StateTime -Text ([string]$state.LastScheduledOccurrence)
        $due = Get-DueOccurrence -Job $job -LastOccurrence $lastOccurrence -Now $now
        if ($null -ne $due) {
            Set-OccurrenceSkipped -JobName $job.Name -Occurrence $due -Reason 'missed occurrence, MissedRunPolicy=Skip'
        }
    }
}

function Invoke-LaunchPhase {
    param([Parameter(Mandatory = $true)][datetime]$Now)

    $limit = $script:Settings.MaxConcurrency
    $launchedThisTick = New-Object System.Collections.Generic.List[string]

    foreach ($job in $script:Manifest.OrderedJobs) {
        $name = $job.Name
        if (-not (Test-JobSelected -JobName $name)) { continue }
        $isForced = ($script:ForcedPending -contains $name)
        if (-not (Test-JobAllowedOnServer -Job $job)) {
            if ($isForced) {
                Write-OrchestratorLog -Message ("Job {0}: -Force refused; this server is not in the job allowlist ({1})." -f $name, ((Get-JobEffectiveAllowedServers -Job $job) -join ', ')) -Level WARN
                $script:ForcedPending = @($script:ForcedPending | Where-Object { $_ -ne $name })
            }
            continue
        }
        $state = Get-JobState -JobName $name
        $lastOccurrence = ConvertFrom-StateTime -Text ([string]$state.LastScheduledOccurrence)

        # Per-job overlap guard: a job still running at its next occurrence is not
        # relaunched; the new occurrence is marked Skipped.
        if ($script:RunningJobs.ContainsKey($name)) {
            if ($job.Enabled) {
                $due = Get-DueOccurrence -Job $job -LastOccurrence $lastOccurrence -Now $Now
                if ($null -ne $due) { Set-OccurrenceSkipped -JobName $name -Occurrence $due -Reason 'previous run still in progress (overlap guard)' }
            }
            if ($isForced) {
                Write-OrchestratorLog -Message ("Job {0}: -Force ignored because the job is already running." -f $name) -Level WARN
                $script:ForcedPending = @($script:ForcedPending | Where-Object { $_ -ne $name })
            }
            continue
        }

        $occurrence = $null
        $attempt = 0
        $reason = ''
        if ($isForced) {
            $occurrence = $Now
            $reason = 'forced'
        }
        elseif ($null -ne $state.PendingRetry) {
            $notBefore = ConvertFrom-StateTime -Text ([string]$state.PendingRetry.NotBefore)
            if ($null -ne $notBefore -and $notBefore -le $Now) {
                $occurrence = ConvertFrom-StateTime -Text ([string]$state.PendingRetry.ScheduledOccurrence)
                if ($null -eq $occurrence) { $occurrence = $Now }
                $attempt = [int]$state.PendingRetry.Attempt
                $reason = 'retry'
            }
            else { continue }
        }
        elseif ($job.Enabled) {
            $occurrence = Get-DueOccurrence -Job $job -LastOccurrence $lastOccurrence -Now $Now
            $reason = 'due'
        }
        if ($null -eq $occurrence) { continue }

        # Dependency gate (not applied to forced runs): wait while a dependency is
        # running, launched earlier in this tick, pending a retry, or itself due.
        if (-not $isForced) {
            $deferred = $false
            $blockedByParent = $false
            $blockedDependency = ''
            $blockingDependencies = New-Object System.Collections.Generic.List[string]
            foreach ($dep in $job.DependsOn) {
                $depState = Get-JobState -JobName $dep
                $depJob = $null
                if ($script:Manifest.JobsByName.ContainsKey($dep)) { $depJob = $script:Manifest.JobsByName[$dep] }

                if ($script:RunningJobs.ContainsKey($dep) -or $launchedThisTick -contains $dep) {
                    $deferred = $true
                    $blockingDependencies.Add($dep)
                    continue
                }
                if ($null -ne $depState.PendingRetry) {
                    $deferred = $true
                    $blockingDependencies.Add($dep)
                    continue
                }

                # A dependency that is not allowed on this server never runs here and
                # must not block its dependents.
                if ($null -ne $depJob -and $depJob.Enabled -and (Test-JobAllowedOnServer -Job $depJob)) {
                    $depLast = ConvertFrom-StateTime -Text ([string]$depState.LastScheduledOccurrence)
                    $depDue = Get-DueOccurrence -Job $depJob -LastOccurrence $depLast -Now $Now
                    if ($null -ne $depDue) {
                        $deferred = $true
                        $blockingDependencies.Add($dep)
                        continue
                    }
                    if ($depState.LastStatus -in @('Failed', 'TimedOut', 'Interrupted') -and -not $depJob.ContinueOnError) {
                        $depOccurrence = ConvertFrom-StateTime -Text ([string]$depState.LastScheduledOccurrence)
                        if ($null -ne $depOccurrence -and $depOccurrence -eq $occurrence) { $blockedByParent = $true; $blockedDependency = $dep; break }
                    }
                }
            }
            if ($blockedByParent) {
                Set-OccurrenceBlocked -JobName $name -Occurrence $occurrence -Status 'BlockedDependencyFailed' -Reason ("dependency '{0}' finally failed with ContinueOnError=false" -f $blockedDependency)
                Clear-DependencyWaitLog -JobName $name
                continue
            }
            if ($deferred) {
                $dependencyWaitTimeout = [int]$job.DependencyWaitTimeoutMinutes
                if ($dependencyWaitTimeout -le 0) { $dependencyWaitTimeout = [int]$script:Settings.DependencyWaitTimeoutMinutes }
                if ($dependencyWaitTimeout -gt 0 -and $script:DependencyWaitLogState.ContainsKey($name)) {
                    $waitState = $script:DependencyWaitLogState[$name]
                    $waitMinutes = [int]($Now - [datetime]$waitState.FirstSeen).TotalMinutes
                    if ($waitMinutes -ge $dependencyWaitTimeout) {
                        Set-OccurrenceBlocked -JobName $name -Occurrence $occurrence -Status 'BlockedDependencyTimeout' -Reason ("dependencies still blocking after {0} min: {1}" -f $waitMinutes, (@($blockingDependencies) -join ', '))
                        Clear-DependencyWaitLog -JobName $name
                        continue
                    }
                }
                Write-DependencyWaitLog -JobName $name -BlockingDependencies @($blockingDependencies) -Now $Now
                continue
            }
            Clear-DependencyWaitLog -JobName $name
        }

        if ($script:RunningJobs.Count -ge $limit) {
            Write-OrchestratorLog -Message ("Job {0}: {1} occurrence {2} is queued; concurrency limit reached ({3}/{4})." -f $name, $reason, $occurrence.ToString('yyyy-MM-dd HH:mm'), $script:RunningJobs.Count, $limit)
            continue
        }

        if ($reason -in @('due', 'forced')) {
            $state.LastScheduledOccurrence = ConvertTo-StateTime -Value $occurrence
        }
        if ($isForced) {
            $script:ForcedPending = @($script:ForcedPending | Where-Object { $_ -ne $name })
            Write-OrchestratorLog -Message ("Job {0}: launching now (forced)." -f $name)
        }
        Start-InventoryJob -Job $job -Occurrence $occurrence -Attempt $attempt
        $launchedThisTick.Add($name)
    }
}

# ==========================================================
# Daily summary email
# ==========================================================
function Send-DailySummaryIfDue {
    param([Parameter(Mandatory = $true)][datetime]$Now)

    if (-not $script:Settings.SendDailySummaryEmail) { return }
    $today = $Now.ToString('yyyy-MM-dd')
    if ($script:State.LastDailySummaryDate -eq $today) { return }
    $summaryTime = $null
    try { $summaryTime = [TimeSpan]::ParseExact($script:Settings.DailySummaryTime, 'hh\:mm', [System.Globalization.CultureInfo]::InvariantCulture) }
    catch {
        Write-OrchestratorLog -Message ("Invalid DailySummaryTime '{0}'; expected HH:mm." -f $script:Settings.DailySummaryTime) -Level WARN
        return
    }
    if ($Now.TimeOfDay -lt $summaryTime) { return }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($day in @($Now.AddDays(-1), $Now)) {
        $csvPath = Join-Path -Path $script:Settings.JobRunsFolderPath -ChildPath ("Orchestrator_JobRuns_{0}.csv" -f $day.ToString('yyyyMMdd'))
        if (-not (Test-Path -LiteralPath $csvPath)) { continue }
        try {
            foreach ($row in (Import-Csv -LiteralPath $csvPath)) {
                $reference = $row.StartTime
                if ([string]::IsNullOrWhiteSpace($reference)) { $reference = $row.ScheduledTime }
                $parsed = $null
                if (-not [string]::IsNullOrWhiteSpace($reference)) {
                    try { $parsed = [datetime]::ParseExact($reference, 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) } catch { }
                }
                if ($null -ne $parsed -and $parsed -ge $Now.AddHours(-24)) { $rows.Add($row) }
            }
        }
        catch {
            Write-OrchestratorLog -Message ("Failed to read job-runs CSV '{0}': {1}" -f $csvPath, $_.Exception.Message) -Level WARN
        }
    }

    $failedCount = @($rows | Where-Object { $_.Status -in @('Failed', 'TimedOut', 'Interrupted') }).Count
    $tableRows = ''
    foreach ($row in $rows) {
        $color = '#1F2937'
        switch ($row.Status) {
            'Success' { $color = '#107C10' }
            'Failed' { $color = '#D13438' }
            'TimedOut' { $color = '#D13438' }
            'Interrupted' { $color = '#D13438' }
            'Skipped' { $color = '#FF8C00' }
            'Retried' { $color = '#FF8C00' }
        }
        $tableRows += "<tr>" +
            "<td style='padding:3px 10px;border:1px solid #DDDDDD;'>$(ConvertTo-HtmlText -Text $row.JobName)</td>" +
            "<td style='padding:3px 10px;border:1px solid #DDDDDD;'>$(ConvertTo-HtmlText -Text $row.ScheduledTime)</td>" +
            "<td style='padding:3px 10px;border:1px solid #DDDDDD;'>$(ConvertTo-HtmlText -Text $row.StartTime)</td>" +
            "<td style='padding:3px 10px;border:1px solid #DDDDDD;'>$(ConvertTo-HtmlText -Text $row.EndTime)</td>" +
            "<td style='padding:3px 10px;border:1px solid #DDDDDD;text-align:right;'>$(ConvertTo-HtmlText -Text $row.DurationSec)</td>" +
            "<td style='padding:3px 10px;border:1px solid #DDDDDD;color:$color;font-weight:bold;'>$(ConvertTo-HtmlText -Text $row.Status)</td>" +
            "<td style='padding:3px 10px;border:1px solid #DDDDDD;text-align:right;'>$(ConvertTo-HtmlText -Text $row.RetryCount)</td>" +
            "</tr>"
    }
    if ([string]::IsNullOrEmpty($tableRows)) {
        $tableRows = "<tr><td colspan='7' style='padding:3px 10px;border:1px solid #DDDDDD;'>No job execution in the last 24 hours.</td></tr>"
    }
    $body = @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#1F2937;'>
<h2>SmartInventory orchestrator daily summary</h2>
<p>Tenant $(ConvertTo-HtmlText -Text $Tenant) on $(ConvertTo-HtmlText -Text $env:COMPUTERNAME): $($rows.Count) execution(s) in the last 24 hours, $failedCount failure(s).</p>
<table style='border-collapse:collapse;'>
<tr style='background-color:#E6F4FF;'>
<th style='padding:3px 10px;border:1px solid #DDDDDD;text-align:left;'>Job</th>
<th style='padding:3px 10px;border:1px solid #DDDDDD;text-align:left;'>Scheduled</th>
<th style='padding:3px 10px;border:1px solid #DDDDDD;text-align:left;'>Start</th>
<th style='padding:3px 10px;border:1px solid #DDDDDD;text-align:left;'>End</th>
<th style='padding:3px 10px;border:1px solid #DDDDDD;text-align:left;'>Duration (s)</th>
<th style='padding:3px 10px;border:1px solid #DDDDDD;text-align:left;'>Status</th>
<th style='padding:3px 10px;border:1px solid #DDDDDD;text-align:left;'>Retries</th>
</tr>
$tableRows
</table>
<p style='color:#5F6B7A;'>Sent by $ScriptName v$ScriptVersion.</p>
</body></html>
"@
    $subject = "[SmartM365 Orchestrator][$Tenant] Daily summary: $($rows.Count) run(s), $failedCount failure(s)"
    if (Send-OrchestratorMail -Subject $subject -HtmlBody $body) {
        $script:State.LastDailySummaryDate = $today
        Save-OrchestratorState

    }
}

# ==========================================================
# Retention cleanup
# ==========================================================
function Invoke-RetentionCleanup {
    $targets = @(
        @{ Folder = $script:Settings.OrchestratorLogFolderPath; Filter = '*.log'; Days = $script:Settings.OrchestratorLogRetentionDays; Recurse = $false },
        @{ Folder = $script:Settings.JobLogFolderPath; Filter = '*.log'; Days = $script:Settings.JobLogRetentionDays; Recurse = $true },
        @{ Folder = $script:Settings.JobRunsFolderPath; Filter = 'Orchestrator_JobRuns_*.csv'; Days = $script:Settings.JobRunsCsvRetentionDays; Recurse = $false }
    )
    foreach ($target in $targets) {
        if ($target.Days -le 0) { continue }
        if (-not (Test-Path -LiteralPath $target.Folder)) { continue }
        $cutoff = (Get-Date).AddDays(-1 * $target.Days)
        $items = @(Get-ChildItem -Path $target.Folder -File -Filter $target.Filter -Recurse:($target.Recurse) -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff })
        foreach ($item in $items) {
            try {
                Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop
                Write-OrchestratorLog -Message ("Retention: removed {0}" -f $item.FullName)
            }
            catch {
                Write-OrchestratorLog -Message ("Retention: failed to remove {0}: {1}" -f $item.FullName, $_.Exception.Message) -Level WARN
            }
        }
    }
}

# ==========================================================
# Dry run plan
# ==========================================================
function Show-DryRunPlan {
    $now = Get-Date
    Write-OrchestratorLog -Message ("Dry run: planned occurrences from {0} to {1}." -f $now.ToString('yyyy-MM-dd HH:mm'), $now.AddHours(24).ToString('yyyy-MM-dd HH:mm'))
    foreach ($job in $script:Manifest.OrderedJobs) {
        $schedule = $job.Schedule
        $scheduleText = ''
        if ($schedule.Type -eq 'Daily') { $scheduleText = 'Daily at ' + ($schedule.Times -join ', ') }
        else { $scheduleText = 'Weekly on ' + ($schedule.DaysOfWeek -join ', ') + ' at ' + ($schedule.Times -join ', ') }
        $flags = @()
        if (-not $job.Enabled) { $flags += 'disabled' }
        if (-not (Test-JobSelected -JobName $job.Name)) { $flags += 'filtered out by -Only/-Skip' }
        if ($job.DependsOn.Count -gt 0) { $flags += ('depends on ' + ($job.DependsOn -join ', ')) }
        if ($job.PowerShellEdition -eq 'WindowsPowerShell') { $flags += 'WindowsPowerShell' }
        if ($job.PSObject.Properties['LauncherPath'] -and -not [string]::IsNullOrWhiteSpace([string]$job.LauncherPath)) { $flags += ('launcher ' + [string]$job.LauncherPath) }
        if (-not (Test-JobAllowedOnServer -Job $job)) { $flags += ('not allowed on this server; allowed: ' + ((Get-JobEffectiveAllowedServers -Job $job) -join ', ')) }
        $flagText = ''
        if ($flags.Count -gt 0) { $flagText = ' [' + ($flags -join '; ') + ']' }
        Write-OrchestratorLog -Message ("Job {0} (group {1}): {2}; policy {3}; timeout {4} min; retries {5}{6}" -f $job.Name, $job.Group, $scheduleText, $schedule.MissedRunPolicy, $job.TimeoutMinutes, $job.MaxRetries, $flagText)

        if (-not $job.Enabled -or -not (Test-JobSelected -JobName $job.Name) -or -not (Test-JobAllowedOnServer -Job $job)) { continue }
        $state = $null
        if ($script:State.Jobs.ContainsKey($job.Name)) { $state = $script:State.Jobs[$job.Name] }
        if ($null -ne $state) {
            $lastOccurrence = ConvertFrom-StateTime -Text ([string]$state.LastScheduledOccurrence)
            $due = Get-DueOccurrence -Job $job -LastOccurrence $lastOccurrence -Now $now
            if ($null -ne $due) {
                if ($schedule.MissedRunPolicy -eq 'RunOnce') {
                    Write-OrchestratorLog -Message ("  Catch-up: missed occurrence {0} would run once at startup (MissedRunPolicy=RunOnce)." -f $due.ToString('yyyy-MM-dd HH:mm'))
                }
                else {
                    Write-OrchestratorLog -Message ("  Catch-up: missed occurrence {0} would be skipped (MissedRunPolicy=Skip)." -f $due.ToString('yyyy-MM-dd HH:mm'))
                }
            }
            if ($null -ne $state.Running) {
                Write-OrchestratorLog -Message ("  State: job recorded as Running (PID {0}); a real start would try to re-adopt it." -f $state.Running.Pid)
            }
        }
        $upcoming = Get-JobOccurrencesInWindow -Job $job -WindowStart $now -WindowEnd $now.AddHours(24)
        if ($upcoming.Count -eq 0) {
            Write-OrchestratorLog -Message "  Next 24h: no occurrence."
        }
        else {
            Write-OrchestratorLog -Message ("  Next 24h: {0}" -f (($upcoming | ForEach-Object { $_.ToString('yyyy-MM-dd HH:mm') }) -join ', '))
        }
    }
    Write-OrchestratorLog -Message "Dry run complete; nothing was launched."
}

# ==========================================================
# Initialization
# ==========================================================
$script:LogReady = $false
$script:LockOwned = $false
$script:RunningJobs = @{}
$script:ForcedPending = @()
$script:SmtpEndpoint = ''
$script:Manifest = $null
$script:State = $null
$script:DependencyWaitLogState = @{}
$script:SharePointUploadedFileState = @{}
$script:LastSharePointUploadAttempt = [datetime]::MinValue
$script:LastHeartbeatLogTime = [datetime]::MinValue
$script:LastSharePointConfigWarningKey = ''
try {
    $tenantContextPath = Find-SmartM365TenantContextPath
    $coreModulePath = Find-SmartM365CoreModulePath
    Import-Module $coreModulePath -Force -ErrorAction Stop
    . $tenantContextPath
    $script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
    $localConfig = Get-SmartM365ScriptLocalConfig

    $script:SmartInventoryRoot = Split-Path -Path $PSScriptRoot -Parent

    $dataFolder = Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'OrchestratorDataFolderPath' -DefaultValue (Join-Path -Path $PSScriptRoot -ChildPath 'Output')
    $logFolder = Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'OrchestratorLogFolderPath' -DefaultValue (Join-Path -Path $dataFolder -ChildPath 'Logs')

    $sharedDataFolder = $dataFolder
    if ((Split-Path -Path $sharedDataFolder -Leaf) -eq $env:COMPUTERNAME) {
        $sharedDataFolder = Split-Path -Path $sharedDataFolder -Parent
    }

    # Per-server isolation: DataAllRootPath/LogAllRootPath may be a UNC share used by
    # several orchestrator servers. State, lock, heartbeat, job-runs CSVs and logs are
    # therefore suffixed with the local computer name so instances never collide.
    if ((Split-Path -Path $dataFolder -Leaf) -ne $env:COMPUTERNAME) { $dataFolder = Join-Path -Path $dataFolder -ChildPath $env:COMPUTERNAME }
    if ((Split-Path -Path $logFolder -Leaf) -ne $env:COMPUTERNAME) { $logFolder = Join-Path -Path $logFolder -ChildPath $env:COMPUTERNAME }

    $effectiveMaxConcurrency = Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'MaxConcurrency' -DefaultValue 2
    if ($MaxConcurrency -gt 0) { $effectiveMaxConcurrency = $MaxConcurrency }
    $effectiveMaxLifetimeHours = Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'MaxLifetimeHours' -DefaultValue 24
    if ($MaxLifetimeHours -gt 0) { $effectiveMaxLifetimeHours = $MaxLifetimeHours }

    $effectiveManifestPath = Join-Path -Path $PSScriptRoot -ChildPath 'Orchestrator-Jobs.json'
    if (-not [string]::IsNullOrWhiteSpace($JobsManifestPath)) { $effectiveManifestPath = $JobsManifestPath }
    $effectiveStatePath = Join-Path -Path $dataFolder -ChildPath 'Orchestrator-State.json'
    if (-not [string]::IsNullOrWhiteSpace($StatePath)) { $effectiveStatePath = $StatePath }

    $mailFrom = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'From' -DefaultValue '')
    $mailTo = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'To' -DefaultValue '')
    $errorMailTo = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'ErrorMailTo' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($mailTo)) { $mailTo = $errorMailTo }
    if ([string]::IsNullOrWhiteSpace($errorMailTo)) { $errorMailTo = $mailTo }
    $smtpServer = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'SmtpServer' -DefaultValue '')
    $sendMailMode = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'SendMailMode' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($sendMailMode)) { $sendMailMode = if ([string]::IsNullOrWhiteSpace($smtpServer)) { 'Graph' } else { 'SMTP' } }
    $sendMailMode = $sendMailMode.Trim()
    if ($sendMailMode -notin @('Graph', 'SMTP', 'Both')) { $sendMailMode = 'Graph' }

    $global:AppId = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'AppId' -DefaultValue '')
    $global:TenantId = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'TenantId' -DefaultValue '')
    $global:Thumb = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'Thumb' -DefaultValue '')
    $global:Thumbprint = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'Thumbprint' -DefaultValue $global:Thumb)

    $mailConfigIssues = @()
    if ([string]::IsNullOrWhiteSpace($mailFrom)) { $mailConfigIssues += 'From missing' }
    if ([string]::IsNullOrWhiteSpace($mailTo)) { $mailConfigIssues += 'To/ErrorMailTo missing' }
    $canGraphMail = -not [string]::IsNullOrWhiteSpace($global:AppId) -and -not [string]::IsNullOrWhiteSpace($global:TenantId) -and -not [string]::IsNullOrWhiteSpace($global:Thumbprint)
    $canSmtpMail = -not [string]::IsNullOrWhiteSpace($smtpServer)
    if ($sendMailMode -eq 'Graph' -and -not $canGraphMail) { $mailConfigIssues += 'Graph app auth missing (AppId/TenantId/Thumbprint)' }
    if ($sendMailMode -eq 'SMTP' -and -not $canSmtpMail) { $mailConfigIssues += 'SmtpServer missing for SMTP mode' }
    if ($sendMailMode -eq 'Both' -and -not ($canGraphMail -or $canSmtpMail)) { $mailConfigIssues += 'Graph app auth and SMTP fallback missing' }
    $mailConfigIssueText = if ($mailConfigIssues.Count -gt 0) { $mailConfigIssues -join '; ' } else { '' }
    $authenticodeValidationEnabled = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'AuthenticodeValidationEnabled' -DefaultValue $false
    $authenticodeValidationMode = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'AuthenticodeValidationMode' -DefaultValue 'Audit')
    if ($authenticodeValidationMode -notin @('Audit', 'Enforce')) { $authenticodeValidationMode = 'Audit' }
    $authenticodeAllowedThumbprints = @(ConvertTo-OrchestratorStringList -Value (Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'AuthenticodeAllowedThumbprints' -DefaultValue @()) | ForEach-Object { $_.Replace(' ', '').ToUpperInvariant() } | Where-Object { $_ })
    $authenticodeCheckCoreModule = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'AuthenticodeCheckCoreModule' -DefaultValue $true
    $authenticodeCheckWindowsPowerShellModule = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'AuthenticodeCheckWindowsPowerShellModule' -DefaultValue $true
    $authenticodeInstallTrustedCertificates = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'AuthenticodeInstallTrustedCertificates' -DefaultValue $false
    $authenticodeTrustedCertificatePaths = @(ConvertTo-OrchestratorStringList -Value (Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'AuthenticodeTrustedCertificatePaths' -DefaultValue @()))
    $authenticodeInstallTrustedRoot = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'AuthenticodeInstallTrustedRoot' -DefaultValue $true
    $authenticodeInstallTrustedPublisher = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'AuthenticodeInstallTrustedPublisher' -DefaultValue $true

    # AllowedServers: default server allowlist for every job (empty = all servers).
    # Accepts a JSON array or a comma/semicolon separated string.
    $allowedServersRaw = Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'AllowedServers' -DefaultValue @()
    $defaultAllowedServers = @()
    if ($null -ne $allowedServersRaw) {
        if ($allowedServersRaw -is [string]) { $defaultAllowedServers = @($allowedServersRaw -split '[;,]') }
        else { $defaultAllowedServers = @($allowedServersRaw) }
        $defaultAllowedServers = @($defaultAllowedServers | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    }

    # JobMailMode is intentionally a dedicated key: the ecosystem-wide SendMailMode key
    # carries the mail transport (Graph/SMTP/Both), not a notification policy.
    $jobMailMode = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'JobMailMode' -DefaultValue 'OnError')
    if ($jobMailMode -notin @('Always', 'OnError', 'Never')) {
        Write-Host ("Invalid JobMailMode '{0}'; falling back to OnError." -f $jobMailMode) -ForegroundColor Yellow
        $jobMailMode = 'OnError'
    }

    $script:Settings = [pscustomobject]@{
        OrchestratorDataFolderPath = $dataFolder
        OrchestratorLogFolderPath = $logFolder
        JobLogFolderPath = (Join-Path -Path $logFolder -ChildPath 'Jobs')
        JobRunsFolderPath = (Join-Path -Path $dataFolder -ChildPath 'JobRuns')
        OrchestratorRunsCsvPath = (Join-Path -Path $sharedDataFolder -ChildPath 'Orchestrator_Runs.csv')
        OrchestratorRunsLockPath = (Join-Path -Path $sharedDataFolder -ChildPath 'Orchestrator_Runs.lock')
        StatePath = $effectiveStatePath
        LockPath = (Join-Path -Path $dataFolder -ChildPath 'Orchestrator.lock')
        HeartbeatPath = (Join-Path -Path $dataFolder -ChildPath 'Orchestrator-Heartbeat.json')
        StopRequestPath = (Join-Path -Path $dataFolder -ChildPath 'Orchestrator-StopRequested.json')
        JobsManifestPath = $effectiveManifestPath
        MaxConcurrency = [math]::Max(1, $effectiveMaxConcurrency)
        MaxLifetimeHours = [math]::Max(1, $effectiveMaxLifetimeHours)
        TickSeconds = [math]::Max(15, (Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'TickSeconds' -DefaultValue 60))
        DependencyWaitLogIntervalMinutes = [math]::Max(0, (Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'DependencyWaitLogIntervalMinutes' -DefaultValue 30))
        DependencyWaitTimeoutMinutes = [math]::Max(0, (Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'DependencyWaitTimeoutMinutes' -DefaultValue 1440))
        OrchestratorRunsCsvLockTimeoutSeconds = [math]::Max(1, (Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'OrchestratorRunsCsvLockTimeoutSeconds' -DefaultValue 300))
        OrchestratorHeartbeatLogIntervalMinutes = [math]::Max(0, (Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'OrchestratorHeartbeatLogIntervalMinutes' -DefaultValue 30))
        OrchestratorSharePointUploadIntervalMinutes = [math]::Max(0, (Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'OrchestratorSharePointUploadIntervalMinutes' -DefaultValue 60))
        SharePointUploadEnabled = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'EnableSharePointUpload' -DefaultValue ([bool]$global:EnableSharePointUpload)
        SharePointSiteHostname = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'SharePointSiteHostname' -DefaultValue $global:SharePointSiteHostname)
        SharePointSitePath = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'SharePointSitePath' -DefaultValue $global:SharePointSitePath)
        SharePointLibraryDisplayName = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'SharePointLibraryDisplayName' -DefaultValue $global:SharePointLibraryDisplayName)
        SharePointTargetFolderPath = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'SharePointTargetFolderPath' -DefaultValue $global:SharePointTargetFolderPath)
        CoreModulePath = $coreModulePath
        AuthenticodeValidationEnabled = $authenticodeValidationEnabled
        AuthenticodeValidationMode = $authenticodeValidationMode
        AuthenticodeAllowedThumbprints = $authenticodeAllowedThumbprints
        AuthenticodeCheckCoreModule = $authenticodeCheckCoreModule
        AuthenticodeCheckWindowsPowerShellModule = $authenticodeCheckWindowsPowerShellModule
        AuthenticodeInstallTrustedCertificates = $authenticodeInstallTrustedCertificates
        AuthenticodeTrustedCertificatePaths = $authenticodeTrustedCertificatePaths
        AuthenticodeInstallTrustedRoot = $authenticodeInstallTrustedRoot
        AuthenticodeInstallTrustedPublisher = $authenticodeInstallTrustedPublisher
        MailFrom = $mailFrom
        MailTo = $mailTo
        ErrorMailTo = $errorMailTo
        MailCc = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'Cc' -DefaultValue '')
        MailBcc = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'Bcc' -DefaultValue '')
        SmtpServer = $smtpServer
        SmtpPort = Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'SmtpPort' -DefaultValue 25
        SendMailMode = $sendMailMode
        GraphMailConfigured = $canGraphMail
        SmtpMailConfigured = $canSmtpMail
        UseIntegratedAuth = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'UseIntegratedAuth' -DefaultValue $true
        UseSsl = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'UseSsl' -DefaultValue $false
        RelayIp = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'RelayIp' -DefaultValue '')
        JobMailMode = $jobMailMode
        DefaultAllowedServers = $defaultAllowedServers
        SendDailySummaryEmail = Get-SmartM365ScriptConfigBool -Config $localConfig -Name 'SendDailySummaryEmail' -DefaultValue $false
        DailySummaryTime = [string](Get-SmartM365ScriptConfigValue -Config $localConfig -Name 'DailySummaryTime' -DefaultValue '07:00')
        OrchestratorLogRetentionDays = Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'OrchestratorLogRetentionDays' -DefaultValue 30
        JobLogRetentionDays = Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'JobLogRetentionDays' -DefaultValue 30
        JobRunsCsvRetentionDays = Get-SmartM365ScriptConfigInt -Config $localConfig -Name 'JobRunsCsvRetentionDays' -DefaultValue 90
        MailEnabled = ($mailConfigIssues.Count -eq 0)
        MailConfigIssue = $mailConfigIssueText
    }

    foreach ($folder in @($sharedDataFolder, $script:Settings.OrchestratorDataFolderPath, $script:Settings.OrchestratorLogFolderPath, $script:Settings.JobLogFolderPath, $script:Settings.JobRunsFolderPath)) {
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    }
    $script:LogReady = $true
    $global:LogTextFile = Get-OrchestratorLogPath
    $global:SmartM365WarningCount = 0
    $global:SmartM365ErrorCount = 0
}
catch {
    Write-Host ("Configuration error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 2
}

# ==========================================================
# Main
# ==========================================================
$script:ExitCode = 0
$script:StartTime = Get-Date
$script:LifetimeDeadline = $script:StartTime.AddHours($script:Settings.MaxLifetimeHours)
$script:OrchestratorRunId = ''
$script:OrchestratorRunRegistered = $false
$script:OrchestratorStopReason = 'NormalShutdown'
$script:OrchestratorRunErrorText = ''

try {
    if ($Stop) {
        $script:ExitCode = Request-OrchestratorStop -TimeoutSeconds $StopTimeoutSeconds
        $script:OrchestratorStopReason = 'StopRequestSubmitted'
        exit $script:ExitCode
    }

    Add-OrchestratorRunTracking
    Write-OrchestratorLog -Message ("{0} v{1} starting (PID {2}, tenant {3}, MaxConcurrency {4}, MaxLifetimeHours {5})." -f $ScriptName, $ScriptVersion, $PID, $Tenant, $script:Settings.MaxConcurrency, $script:Settings.MaxLifetimeHours)
    Write-OrchestratorLog -Message ("Runtime context: server={0}; user={1}; pid={2}; PowerShell={3}; edition={4}; process64bit={5}." -f $env:COMPUTERNAME, (Get-OrchestratorRunUserName), $PID, $PSVersionTable.PSVersion, $PSVersionTable.PSEdition, [Environment]::Is64BitProcess)
    Write-OrchestratorLog -Message ("Paths: scriptRoot={0}; currentDirectory={1}; data={2}; log={3}; jobLogs={4}; state={5}; heartbeat={6}; stopRequest={7}; runCsv={8}." -f $PSScriptRoot, (Get-Location).Path, $script:Settings.OrchestratorDataFolderPath, $script:Settings.OrchestratorLogFolderPath, $script:Settings.JobLogFolderPath, $script:Settings.StatePath, $script:Settings.HeartbeatPath, $script:Settings.StopRequestPath, $script:Settings.OrchestratorRunsCsvPath)
    Write-OrchestratorLog -Message ("Mail context: enabled={0}; mode={1}; graphConfigured={2}; smtpConfigured={3}; fromConfigured={4}; recipientConfigured={5}; jobMailMode={6}; dailySummary={7}." -f $script:Settings.MailEnabled, $script:Settings.SendMailMode, $script:Settings.GraphMailConfigured, $script:Settings.SmtpMailConfigured, (-not [string]::IsNullOrWhiteSpace($script:Settings.MailFrom)), (-not [string]::IsNullOrWhiteSpace($script:Settings.MailTo)), $script:Settings.JobMailMode, $script:Settings.SendDailySummaryEmail)

    Write-OrchestratorLog -Message ("Orchestrator upload context: sharePointEnabled={0}; target={1}; uploadIntervalMinutes={2}; dependencyWaitLogIntervalMinutes={3}; dependencyWaitTimeoutMinutes={4}; heartbeatLogIntervalMinutes={5}; runsCsvLockTimeoutSeconds={6}." -f $script:Settings.SharePointUploadEnabled, $script:Settings.SharePointTargetFolderPath, $script:Settings.OrchestratorSharePointUploadIntervalMinutes, $script:Settings.DependencyWaitLogIntervalMinutes, $script:Settings.DependencyWaitTimeoutMinutes, $script:Settings.OrchestratorHeartbeatLogIntervalMinutes, $script:Settings.OrchestratorRunsCsvLockTimeoutSeconds)
    Write-OrchestratorLog -Message ("Authenticode context: enabled={0}; mode={1}; allowedThumbprints={2}; checkCoreModule={3}; checkWindowsPowerShellModule={4}; installTrustedCertificates={5}; trustedCertificatePaths={6}; installRoot={7}; installTrustedPublisher={8}." -f $script:Settings.AuthenticodeValidationEnabled, $script:Settings.AuthenticodeValidationMode, @($script:Settings.AuthenticodeAllowedThumbprints).Count, $script:Settings.AuthenticodeCheckCoreModule, $script:Settings.AuthenticodeCheckWindowsPowerShellModule, $script:Settings.AuthenticodeInstallTrustedCertificates, @($script:Settings.AuthenticodeTrustedCertificatePaths).Count, $script:Settings.AuthenticodeInstallTrustedRoot, $script:Settings.AuthenticodeInstallTrustedPublisher)
    Install-OrchestratorAuthenticodeTrustedCertificates
    if (-not $script:Settings.MailEnabled) {
        Write-OrchestratorLog -Message ("Email notifications are disabled ({0})." -f $script:Settings.MailConfigIssue) -Level WARN
    }

    if (-not (Test-Path -LiteralPath $script:Settings.JobsManifestPath)) {
        $manifestTemplatePath = '{0}.template' -f $script:Settings.JobsManifestPath
        if (Test-Path -LiteralPath $manifestTemplatePath) {
            Copy-Item -LiteralPath $manifestTemplatePath -Destination $script:Settings.JobsManifestPath -ErrorAction Stop
            Write-OrchestratorLog -Message ("Created jobs manifest from template: {0}. Adjust Enabled/Schedule/AllowedServers locally; this runtime manifest stays out of Git." -f $script:Settings.JobsManifestPath) -Level WARN
        }
    }

    try {
        $script:Manifest = Read-JobsManifest -Path $script:Settings.JobsManifestPath
        Write-OrchestratorLog -Message ("Jobs manifest loaded: {0} ({1} jobs, {2} enabled)." -f $script:Settings.JobsManifestPath, @($script:Manifest.OrderedJobs).Count, @($script:Manifest.OrderedJobs | Where-Object { $_.Enabled }).Count)
        Write-ServerAllowlistSummary
    }
    catch {
        Write-OrchestratorLog -Message $_.Exception.Message -Level ERROR
        Send-OrchestratorFatalEmail -ErrorText $_.Exception.Message
        $script:ExitCode = 2
        $script:OrchestratorStopReason = 'ManifestError'
        throw
    }

    $script:State = Read-OrchestratorState

    foreach ($requested in @($Force) + @($Only) + @($Skip)) {
        if (-not [string]::IsNullOrWhiteSpace($requested) -and -not $script:Manifest.JobsByName.ContainsKey($requested)) {
            Write-OrchestratorLog -Message ("Job name '{0}' from -Force/-Only/-Skip is not in the manifest." -f $requested) -Level WARN
        }
    }

    if ($DryRun) {
        Show-DryRunPlan
        $script:OrchestratorStopReason = 'DryRun'
        exit 0
    }

    if (-not (Enter-OrchestratorLock)) {
        $script:ExitCode = 3
        $script:OrchestratorStopReason = 'LockConflict'
        exit 3
    }
    $script:LockOwned = $true

    Restore-RunningJobs
    Initialize-NewJobStates
    Invoke-MissedRunCatchUp
    Invoke-RetentionCleanup
    $script:ForcedPending = @($Force | Where-Object { $script:Manifest.JobsByName.ContainsKey($_) })

    $lastCleanupDate = (Get-Date).Date
    while ($true) {
        $now = Get-Date
        if (Test-OrchestratorStopRequested) {
            $script:OrchestratorStopReason = 'ManualStopRequest'
            break
        }
        Update-JobsManifestIfChanged
        Update-RunningJobs -Now $now
        Invoke-LaunchPhase -Now $now
        Send-DailySummaryIfDue -Now $now
        Write-OrchestratorHeartbeat
        Save-OrchestratorState

        Invoke-OrchestratorPeriodicSharePointUpload -Now $now
        if ($now.Date -ne $lastCleanupDate) {
            $previousLogDate = $lastCleanupDate
            Invoke-RetentionCleanup
            Invoke-OrchestratorSharePointUpload -LocalFilePath (Get-OrchestratorLogPath -Date $previousLogDate) -Reason 'daily orchestrator log rollover' -Force | Out-Null
            $lastCleanupDate = $now.Date
        }

        if ($Once) {
            Write-OrchestratorLog -Message "Single tick complete (-Once); exiting. Launched jobs keep running detached."
            $script:OrchestratorStopReason = 'Once'
            break
        }
        if ((Get-Date) -ge $script:LifetimeDeadline) {
            Write-OrchestratorLog -Message ("Maximum lifetime of {0} hour(s) reached; recycling. No new job will be launched by this instance." -f $script:Settings.MaxLifetimeHours)
            $script:OrchestratorStopReason = 'MaxLifetime'
            break
        }
        Start-Sleep -Seconds $script:Settings.TickSeconds
    }

    if ($script:RunningJobs.Count -gt 0) {
        foreach ($name in $script:RunningJobs.Keys) {
            $info = $script:RunningJobs[$name]
            Write-OrchestratorLog -Message ("Job {0}: left running detached (PID {1}, started {2}); the next orchestrator instance will re-adopt it." -f $name, $info.Process.Id, ([datetime]$info.StartTime).ToString('yyyy-MM-dd HH:mm:ss'))
        }
    }
    Write-OrchestratorLog -Message ("{0} exiting normally (exit code 0)." -f $ScriptName)
}
catch {
    if ($script:ExitCode -eq 0) { $script:ExitCode = 1 }
    $errorText = "{0}`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace
    $script:OrchestratorRunErrorText = $errorText
    if ($script:OrchestratorStopReason -eq 'NormalShutdown') {
        $script:OrchestratorStopReason = 'FatalError'
    }
    Write-OrchestratorLog -Message ("Fatal orchestrator error: {0}" -f $errorText) -Level ERROR
    if ($script:ExitCode -eq 1) { Send-OrchestratorFatalEmail -ErrorText $errorText }
}
finally {
    if ($null -ne $script:State -and -not $DryRun) {
        try { Save-OrchestratorState } catch { }
    }
    Exit-OrchestratorLock
    $runStatus = 'Failed'
    if ($script:ExitCode -eq 0) { $runStatus = 'Success' }
    elseif ($script:ExitCode -eq 3) { $runStatus = 'Rejected' }
    try {
        Complete-OrchestratorRunTracking -Status $runStatus -ExitCode $script:ExitCode -StopReason $script:OrchestratorStopReason -ErrorMessage $script:OrchestratorRunErrorText
    }
    catch {
        Write-OrchestratorLog -Message ("Failed to finalize orchestrator lifecycle tracking: {0}" -f $_.Exception.Message) -Level WARN
    }
    if (-not $DryRun -and -not $Stop) {
        try {
            Invoke-OrchestratorPeriodicSharePointUpload -Now (Get-Date) -Force
        }
        catch {
            Write-OrchestratorLog -Message ("Final orchestrator SharePoint upload failed: {0}" -f $_.Exception.Message) -Level ERROR
        }
    }
}

exit $script:ExitCode

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAL0nNM2bGnWLhl
# mMacZB+NiQNi+DHwMPZhPnl64qIltKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCCpLkMrFTwIgeNqaxri27m8gsRiXxJHIffxoWY6nofNHDANBgkqhkiG9w0B
# AQEFAASCAYBZTvoA+jHBob4OOI2crfUx4i3X6RX+o9/O376yRTOc+KscL2ZPO5zs
# nQEf4vyDgVa3mGO06xpeTDQSvV9JeiLZXfTMGaLFRjopaguVioA/Xgwcbqz7wA/V
# FNjYPBfe5AwD0SxNLoac0J/xVhsR2tY1uWmTQBJXcyq2LcC170Cj+lazyqZE1cxQ
# p5Ufzg3zz3KbFryvFQMBqoYUfjcxowcK0TA4wRXmAEZc9F8xcM4AOq2N6bbhtLxB
# h+0p3JkRjBRyyaAZTw/2aITIAMJ2cowz16yvTb2rVu801pifWtRaetml1OD9c53e
# lrCW+WVnyNYalA/6jzhqNmLf1elIwXf/fE2PRMRIIpfTcsEpjnmEFrjQafaZWn8U
# 1tYGefSnXrQ0632/ZqXy46FwQePfM5jvBhvxnvDTZNdZ53FVPxNDkcJTIYQdDLLE
# ZShkUlKMM4qmxQzUl3hD7ZS0TNMuRO8hyhNutCYJPOKy5tKbBBxlJ2knfsCW8Cn+
# Mlp2jtnNFTw=
# SIG # End signature block
