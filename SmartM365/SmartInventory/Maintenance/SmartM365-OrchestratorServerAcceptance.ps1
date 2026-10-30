<#
.SYNOPSIS
Acceptance test of a server before it becomes an active SmartM365 orchestrator server.

.DESCRIPTION
For each capability the orchestrator can assign, the script runs a real but limited read
(about 100 objects), checks the returned data, then discards it: rows stay in memory and no CSV
is written. A full run publishes its statuses (no row value) for the orchestrator console in
<orchestrator shared data folder>\<SERVER>: Orchestrator-Acceptance.json under the account of the
orchestrator scheduled task, Orchestrator-Acceptance.OtherAccount.json otherwise. When the tenant enables
EnableSharePointUpload, that result file and the run log are uploaded to SharePoint like the other
scripts (-NoSharePointUpload skips it). No other write reaches the tenant: no Teams notification, and
no mail unless -SendTestMail is given.

Capabilities, in order:
- RunAccount: running identity; -ExpectedAccount fails the run under another account.
- Modules: PowerShell 7 and Windows PowerShell 5.1, the modules the jobs import (Graph submodules
  included), same version for every installed Microsoft.Graph module, snap-ins.
- CodeSigning: Authenticode status of this script, the Core module and the orchestrator files.
- Network: HTTPS reachability of the Microsoft 365 endpoints, proxy in use, latency.
- Clock: offset against the Date header of the Entra sign-in endpoint (certificate sign-in).
- Certificate: same readiness test as the orchestrator (store, expiry, private key signature).
- SharedRuntime: DATA-ALL, DATA-LAST and LOG-ALL read/write, SMB latency of a witness folder under
  WorkspaceRootPath (created then removed), read latency of the code share, DATA and local free space.
- AD: forest domains and up to SampleSize users per domain (PowerShell 7, like AD-Inventory),
  plus an ActiveDirectory module check in Windows PowerShell 5.1.
- Graph: up to SampleSize users, Entra devices and Intune managed devices.
- GraphPermissions: application roles of the token against the roles declared by the enabled jobs.
- EXO: up to SampleSize mailboxes.
- ExchangeOnPrem: up to SampleSize mailboxes through the Exchange snap-in (Windows PowerShell 5.1).
- SharePointOnPrem: farm and up to SmallSampleSize site collections (Windows PowerShell 5.1).
- TeamsPowerShell: same calls as the Teams job: connection (same TenantId) and up to
  SmallSampleSize phone number assignments. A failure is a non-blocking WARN (the election keeps
  the Teams job on other servers), FAIL only when requested with -Capability.
- Adaxes: services, domains and LDAP filter of the Adaxes collector configuration; up to
  SampleSize users per domain, one child process per domain (-AdaxesDomainTimeoutSeconds); the
  domains after a timed-out one are not read.
- Mail: mail configuration (Graph Mail.Send or SMTP banner); a test mail only with -SendTestMail.

A capability whose component is not installed is SKIPPED, or FAIL when it was requested with
-Capability. Extractions run in a child PowerShell process with a time limit; after a time-out
the last step announced by the child is reported. The jobs of the
catalog that this server could run are listed at the end (same rule as the election).

.VERSION
1.2.0

.EXAMPLE
./SmartM365-OrchestratorServerAcceptance.ps1 -Tenant prod
./SmartM365-OrchestratorServerAcceptance.ps1 -Tenant prod -Capability AD, ExchangeOnPrem
./SmartM365-OrchestratorServerAcceptance.ps1 -Tenant prod -ExpectedAccount 'DOMAIN\service-account'
./SmartM365-OrchestratorServerAcceptance.ps1 -Tenant prod -Capability Mail -SendTestMail

.NOTES
Exit code 1 when at least one capability is FAIL, 0 otherwise. The log goes to LOG-ALL like the
other scripts (SmartM365-OrchestratorServerAcceptance folder); it holds counts, never row values.
#>
#requires -Version 7.2
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [ValidateSet('All', 'RunAccount', 'Modules', 'CodeSigning', 'Network', 'Clock', 'Certificate', 'SharedRuntime', 'AD', 'Graph', 'GraphPermissions', 'EXO', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell', 'Adaxes', 'Mail')]
    [string[]]$Capability = @('All'),
    [ValidateRange(1, 1000)][int]$SampleSize = 100,
    [ValidateRange(1, 100)][int]$SmallSampleSize = 10,
    [ValidateRange(30, 3600)][int]$CapabilityTimeoutSeconds = 600,
    [ValidateRange(30, 3600)][int]$AdaxesDomainTimeoutSeconds = 120,
    [ValidateRange(5, 500)][int]$SmbIterations = 20,
    [ValidateRange(1, 10000)][double]$SmbWarningMilliseconds = 15,
    [ValidateRange(1, 10000)][double]$SmbFailureMilliseconds = 50,
    [string]$ExpectedAccount = '',
    [switch]$SendTestMail,
    [switch]$NoSharePointUpload,
    [string]$TestMailTo = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# Figures in the log and the tables use a dot whatever the regional settings of the server.
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::InvariantCulture

$script:CapabilityOrder = @('RunAccount', 'Modules', 'CodeSigning', 'Network', 'Clock', 'Certificate', 'SharedRuntime', 'AD', 'Graph', 'GraphPermissions', 'EXO', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell', 'Adaxes', 'Mail')
# Capability names the orchestrator election uses.
$script:OrchestratorCapabilities = @('SharedRuntime', 'AD', 'Graph', 'EXO', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell', 'Adaxes')
# A failure of these capabilities does not block the activation (WARN): the election simply keeps
# their jobs (M365-TeamsPhonePstnUsage-Inventory) on other servers. FAIL only when requested with -Capability.
$script:NonBlockingCapabilities = @('TeamsPowerShell')
# Modules the SmartInventory jobs import; a minimum only where a collector or the orchestrator probe sets one.
# Every Microsoft.Graph module must also be at the Microsoft.Graph.Authentication version (checked below).
$script:ExpectedModules = @(
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Users'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Groups'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Identity.DirectoryManagement'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Reports'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'Microsoft.Graph.DeviceManagement'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Beta.Groups'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Beta.DeviceManagement'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'ImportExcel'; Minimum = ''; Required = $true }
    [pscustomobject]@{ Name = 'MicrosoftTeams'; Minimum = '4.7.1'; Required = $false }
    [pscustomobject]@{ Name = 'ActiveDirectory'; Minimum = ''; Required = $false }
    [pscustomobject]@{ Name = 'Adaxes'; Minimum = ''; Required = $false }
)
$script:GuidPattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
$script:SmtpPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
$script:ResultMarker = 'SMARTM365_ACCEPTANCE_JSON:'

#region Pure helpers (covered by the offline test)

function Get-SmartM365AcceptanceWorstStatus {
    param([AllowEmptyCollection()][string[]]$Status = @())
    foreach ($candidate in @('FAIL', 'WARN', 'PASS', 'SKIPPED')) { if ($Status -contains $candidate) { return $candidate } }
    return 'PASS'
}

function Get-SmartM365AcceptancePercentile {
    param([AllowEmptyCollection()][double[]]$Values = @(), [ValidateRange(0, 100)][double]$Percent = 50)
    if ($Values.Count -eq 0) { return $null }
    $sorted = @($Values | Sort-Object)
    $index = [int][math]::Ceiling($Percent / 100 * $sorted.Count) - 1
    return [math]::Round($sorted[[math]::Min([math]::Max($index, 0), $sorted.Count - 1)], 1)
}

function Get-SmartM365AcceptanceLatencyStatus {
    param([AllowNull()]$MedianMilliseconds, [double]$WarningMilliseconds, [double]$FailureMilliseconds)
    if ($null -eq $MedianMilliseconds) { return 'FAIL' }
    if ([double]$MedianMilliseconds -ge $FailureMilliseconds) { return 'FAIL' }
    if ([double]$MedianMilliseconds -ge $WarningMilliseconds) { return 'WARN' }
    return 'PASS'
}

function Get-SmartM365AcceptanceClockStatus {
    param([double]$OffsetSeconds, [double]$WarningSeconds = 60, [double]$FailureSeconds = 300)
    $absolute = [math]::Abs($OffsetSeconds)
    if ($absolute -ge $FailureSeconds) { return 'FAIL' }
    if ($absolute -ge $WarningSeconds) { return 'WARN' }
    return 'PASS'
}

function Test-SmartM365AcceptanceDataset {
    <#
    Checks rows returned by a limited read. Column rules: Name; Required (every row filled);
    MinFilledPercent; Pattern (on filled values); Unique; Date (parsable, between 1990 and tomorrow).
    #>
    param(
        [AllowEmptyCollection()][object[]]$Rows = @(),
        [Parameter(Mandatory = $true)][int]$SampleSize,
        [Parameter(Mandatory = $true)][object[]]$Columns,
        [switch]$AllowEmpty,
        [datetime]$NowUtc = [datetime]::UtcNow
    )

    $issues = [System.Collections.Generic.List[string]]::new()
    $statuses = [System.Collections.Generic.List[string]]::new()
    $count = $Rows.Count
    if ($count -eq 0) {
        $statuses.Add($(if ($AllowEmpty) { 'WARN' } else { 'FAIL' }))
        $issues.Add('no row returned')
    }
    elseif ($count -gt $SampleSize) {
        $statuses.Add('FAIL')
        $issues.Add(('{0} rows returned for a limit of {1}' -f $count, $SampleSize))
    }

    $minimumDate = [datetime]::new(1990, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
    $maximumDate = $NowUtc.AddDays(1)
    foreach ($column in $Columns) {
        if ($count -eq 0) { break }
        $name = [string]$column.Name
        $values = @(foreach ($row in $Rows) {
                $property = if ($null -ne $row) { $row.PSObject.Properties[$name] } else { $null }
                if ($null -eq $property -or $null -eq $property.Value) { '' } else { ([string]$property.Value).Trim() }
            })
        $filled = @($values | Where-Object { $_ -ne '' })
        $emptyCount = $count - $filled.Count
        $required = $column.PSObject.Properties['Required'] -and [bool]$column.Required
        if ($required -and $emptyCount -gt 0) {
            $statuses.Add('FAIL'); $issues.Add(('{0} empty in {1} row(s)' -f $name, $emptyCount))
        }
        elseif ($column.PSObject.Properties['MinFilledPercent']) {
            $percent = [math]::Round($filled.Count * 100 / $count, 1)
            if ($percent -lt [double]$column.MinFilledPercent) {
                $statuses.Add('WARN'); $issues.Add(('{0} filled in {1} % of rows (expected {2} %)' -f $name, $percent, $column.MinFilledPercent))
            }
        }
        if ($column.PSObject.Properties['Pattern'] -and [string]$column.Pattern) {
            $invalid = @($filled | Where-Object { $_ -notmatch [string]$column.Pattern })
            if ($invalid.Count -gt 0) {
                $statuses.Add($(if ($required) { 'FAIL' } else { 'WARN' }))
                $issues.Add(('{0} has {1} value(s) in an unexpected format' -f $name, $invalid.Count))
            }
        }
        if ($column.PSObject.Properties['Date'] -and [bool]$column.Date) {
            $outOfRange = 0
            foreach ($value in $filled) {
                $parsed = [datetime]::MinValue
                $ok = [datetime]::TryParse($value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)
                if (-not $ok -or $parsed -lt $minimumDate -or $parsed -gt $maximumDate) { $outOfRange++ }
            }
            if ($outOfRange -gt 0) { $statuses.Add('WARN'); $issues.Add(('{0} has {1} unreadable or implausible date(s)' -f $name, $outOfRange)) }
        }
        if ($column.PSObject.Properties['Unique'] -and [bool]$column.Unique) {
            $duplicates = @($filled | Group-Object -NoElement | Where-Object Count -gt 1)
            if ($duplicates.Count -gt 0) {
                $statuses.Add('FAIL'); $issues.Add(('{0} has {1} duplicated value(s)' -f $name, $duplicates.Count))
            }
        }
    }

    return [pscustomobject]@{
        Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses)
        Count = $count
        Issues = @($issues)
    }
}

function Get-SmartM365AcceptanceDatasetSpec {
    param([Parameter(Mandatory = $true)][string]$Name)
    $guid = $script:GuidPattern
    $smtp = $script:SmtpPattern
    switch ($Name) {
        'ADUsers' {
            return [pscustomobject]@{ AllowEmpty = $false; Columns = @(
                    [pscustomobject]@{ Name = 'ObjectGUID'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'SamAccountName'; Required = $true; Unique = $true }
                    [pscustomobject]@{ Name = 'UserPrincipalName'; MinFilledPercent = 50; Pattern = '^[^@\s]+@[^@\s]+$' }
                    [pscustomobject]@{ Name = 'WhenCreated'; MinFilledPercent = 90; Date = $true }) }
        }
        'GraphUsers' {
            return [pscustomobject]@{ AllowEmpty = $false; Columns = @(
                    [pscustomobject]@{ Name = 'Id'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'UserPrincipalName'; Required = $true; Pattern = '^[^@\s]+@[^@\s]+$'; Unique = $true }
                    [pscustomobject]@{ Name = 'CreatedDateTime'; Date = $true }) }
        }
        'EntraDevices' {
            return [pscustomobject]@{ AllowEmpty = $false; Columns = @(
                    [pscustomobject]@{ Name = 'Id'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'DeviceId'; MinFilledPercent = 90; Pattern = $guid }
                    [pscustomobject]@{ Name = 'DisplayName'; Required = $true }) }
        }
        'IntuneDevices' {
            return [pscustomobject]@{ AllowEmpty = $false; Columns = @(
                    [pscustomobject]@{ Name = 'Id'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'DeviceName'; Required = $true }
                    [pscustomobject]@{ Name = 'AzureADDeviceId'; MinFilledPercent = 50; Pattern = $guid }
                    [pscustomobject]@{ Name = 'LastSyncDateTime'; Date = $true }) }
        }
        'EXOMailboxes' {
            return [pscustomobject]@{ AllowEmpty = $false; Columns = @(
                    [pscustomobject]@{ Name = 'Guid'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'PrimarySmtpAddress'; Required = $true; Pattern = $smtp; Unique = $true }
                    [pscustomobject]@{ Name = 'ExternalDirectoryObjectId'; MinFilledPercent = 90; Pattern = $guid }
                    [pscustomobject]@{ Name = 'RecipientTypeDetails'; Required = $true }) }
        }
        'OnPremMailboxes' {
            return [pscustomobject]@{ AllowEmpty = $false; Columns = @(
                    [pscustomobject]@{ Name = 'Guid'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'ExchangeGuid'; Required = $true; Pattern = $guid }
                    [pscustomobject]@{ Name = 'PrimarySmtpAddress'; Required = $true; Pattern = $smtp; Unique = $true }
                    [pscustomobject]@{ Name = 'RecipientTypeDetails'; Required = $true }
                    [pscustomobject]@{ Name = 'Database' }) }
        }
        'SharePointSites' {
            return [pscustomobject]@{ AllowEmpty = $true; Columns = @(
                    [pscustomobject]@{ Name = 'Id'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'Url'; Required = $true; Pattern = '^https?://'; Unique = $true }) }
        }
        'TeamsPhoneNumbers' {
            # A tenant without Teams Phone numbers is possible: an empty answer only warns.
            return [pscustomobject]@{ AllowEmpty = $true; Columns = @(
                    [pscustomobject]@{ Name = 'TelephoneNumber'; Required = $true; Pattern = '^\+?[0-9]'; Unique = $true }
                    [pscustomobject]@{ Name = 'NumberType'; Required = $true }
                    [pscustomobject]@{ Name = 'ActivationState'; MinFilledPercent = 90 }) }
        }
        'AdaxesUsers' {
            return [pscustomobject]@{ AllowEmpty = $false; Columns = @(
                    [pscustomobject]@{ Name = 'ObjectGUID'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'SamAccountName'; Required = $true }
                    [pscustomobject]@{ Name = 'UserPrincipalName'; MinFilledPercent = 50; Pattern = '^[^@\s]+@[^@\s]+$' }) }
        }
        default { throw "Unknown dataset: $Name" }
    }
}

function ConvertFrom-SmartM365AcceptanceChildOutput {
    param([AllowEmptyString()][string]$StandardOutput, [AllowEmptyString()][string]$StandardError = '', [int]$ExitCode = 0)
    $match = [regex]::Match([string]$StandardOutput, '(?m)^' + [regex]::Escape($script:ResultMarker) + '(?<Json>.+?)\r?$')
    if (-not $match.Success) {
        $detail = (([string]$StandardError + ' ' + [string]$StandardOutput).Trim() -replace '[\r\n]+', ' ')
        if ($detail.Length -gt 600) { $detail = $detail.Substring(0, 600) }
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "Child process exited with code $ExitCode without a result." }
        return [pscustomobject]@{ Ready = $false; ComponentMissing = $false; Detail = $detail; Roles = @(); Datasets = @(); Info = [pscustomobject]@{} }
    }
    $result = $match.Groups['Json'].Value | ConvertFrom-Json -ErrorAction Stop
    foreach ($name in @('Ready', 'ComponentMissing', 'Detail', 'Roles', 'Datasets', 'Info')) {
        if (-not $result.PSObject.Properties[$name]) {
            $default = switch ($name) { 'Ready' { $false } 'ComponentMissing' { $false } 'Detail' { '' } 'Info' { [pscustomobject]@{} } default { @() } }
            $result | Add-Member -NotePropertyName $name -NotePropertyValue $default
        }
    }
    return $result
}

function Get-SmartM365AcceptanceLastStage {
    <# Children announce each step on stdout; after a time-out the last one tells where it hung. #>
    param([AllowEmptyString()][string]$StandardOutput)
    $stages = [regex]::Matches([string]$StandardOutput, '(?m)^SMARTM365_ACCEPTANCE_STAGE:(?<Stage>.+?)\r?$')
    if ($stages.Count -eq 0) { return '' }
    return $stages[$stages.Count - 1].Groups['Stage'].Value.Trim()
}

function Get-SmartM365AcceptanceChildAssessment {
    <# Turns a child result into a capability status. Datasets with the same Name and several scopes
       (domains) are partial when only some scopes fail: WARN; all scopes failing is FAIL. #>
    param([Parameter(Mandatory = $true)]$Child, [Parameter(Mandatory = $true)][int]$SampleSize, [switch]$Requested)

    if ([bool]$Child.ComponentMissing) {
        return [pscustomobject]@{ Status = $(if ($Requested) { 'FAIL' } else { 'SKIPPED' }); Detail = [string]$Child.Detail; Extractions = @() }
    }
    if (-not [bool]$Child.Ready) { return [pscustomobject]@{ Status = 'FAIL'; Detail = [string]$Child.Detail; Extractions = @() } }

    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    # One line per data set and scope for the final extraction table (counts and checks only).
    $extractions = [System.Collections.Generic.List[object]]::new()
    if ([string]$Child.Detail) { $parts.Add([string]$Child.Detail) }
    foreach ($group in @(@($Child.Datasets) | Group-Object Name)) {
        $failedScopes = 0
        foreach ($dataset in $group.Group) {
            $label = if ([string]$dataset.Scope) { '{0} {1}' -f $dataset.Name, $dataset.Scope } else { [string]$dataset.Name }
            if ([string]$dataset.Error) {
                $failedScopes++
                $parts.Add(('{0}: read failed: {1}' -f $label, $dataset.Error))
                $extractions.Add([pscustomobject]@{ Dataset = [string]$dataset.Name; Scope = [string]$dataset.Scope; Rows = '-'; Seconds = $dataset.DurationSeconds; Status = 'FAIL'; Checks = 'read failed: ' + [string]$dataset.Error })
                continue
            }
            $spec = Get-SmartM365AcceptanceDatasetSpec -Name ([string]$dataset.Name)
            $check = Test-SmartM365AcceptanceDataset -Rows @($dataset.Rows) -SampleSize $SampleSize -Columns $spec.Columns -AllowEmpty:([bool]$spec.AllowEmpty)
            $statuses.Add($check.Status)
            $text = '{0}: {1} row(s) in {2} s' -f $label, $check.Count, $dataset.DurationSeconds
            if ($check.Issues.Count -gt 0) { $text += ' (' + ($check.Issues -join '; ') + ')' }
            $parts.Add($text)
            $extractions.Add([pscustomobject]@{ Dataset = [string]$dataset.Name; Scope = [string]$dataset.Scope; Rows = $check.Count; Seconds = $dataset.DurationSeconds; Status = $check.Status; Checks = $(if ($check.Issues.Count) { $check.Issues -join '; ' } else { 'keys filled, formats valid, no duplicate' }) })
        }
        if ($failedScopes -gt 0) {
            $groupStatus = if ($failedScopes -eq $group.Count) { 'FAIL' } else { 'WARN' }
            $statuses.Add($groupStatus)
            # A domain failing while others read is a warning for the capability, not a failure.
            foreach ($row in @($extractions | Where-Object { $_.Dataset -eq [string]$group.Name -and $_.Rows -eq '-' })) { $row.Status = $groupStatus }
        }
    }
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join ' | '); Extractions = $extractions.ToArray() }
}

function Get-SmartM365AcceptanceModuleAssessment {
    <# Installed: objects with Name and Version (every installed version). #>
    param([AllowEmptyCollection()][object[]]$Installed = @(), [Parameter(Mandatory = $true)][object[]]$Expected)

    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($module in $Expected) {
        $versions = @($Installed | Where-Object { [string]$_.Name -eq [string]$module.Name } | ForEach-Object { [version][string]$_.Version } | Sort-Object -Descending)
        if ($versions.Count -eq 0) {
            $statuses.Add($(if ([bool]$module.Required) { 'FAIL' } else { 'WARN' }))
            $parts.Add(('{0} missing' -f $module.Name))
            continue
        }
        if ([string]$module.Minimum -and $versions[0] -lt [version][string]$module.Minimum) {
            $statuses.Add($(if ([bool]$module.Required) { 'FAIL' } else { 'WARN' }))
            $parts.Add(('{0} {1} below {2}' -f $module.Name, $versions[0], $module.Minimum))
            continue
        }
        $parts.Add(('{0} {1}' -f $module.Name, $versions[0]))
    }
    # Graph submodules of another version than Microsoft.Graph.Authentication fail to load together.
    $authentication = @($Installed | Where-Object { [string]$_.Name -eq 'Microsoft.Graph.Authentication' } | ForEach-Object { [version][string]$_.Version } | Sort-Object -Descending | Select-Object -First 1)
    if ($authentication.Count -eq 1) {
        $mismatched = @($Installed | Where-Object { [string]$_.Name -like 'Microsoft.Graph.*' -and [string]$_.Name -ne 'Microsoft.Graph.Authentication' } |
                Group-Object Name | ForEach-Object {
                    $latest = @($_.Group | ForEach-Object { [version][string]$_.Version } | Sort-Object -Descending)[0]
                    if ($latest -ne $authentication[0]) { '{0} {1}' -f $_.Name, $latest }
                })
        if ($mismatched.Count -gt 0) {
            $statuses.Add('WARN')
            $parts.Add(('Graph modules not at {0}: {1}' -f $authentication[0], ($mismatched -join ', ')))
        }
    }
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join '; ') }
}

function ConvertTo-SmartM365AcceptanceStringList {
    param([AllowNull()]$Value)
    return @(@($Value) | ForEach-Object { if ($null -ne $_) { ([string]$_).Trim() } } | Where-Object { $_ })
}

function Get-SmartM365AcceptanceRequiredGraphRole {
    <# Roles declared by the enabled jobs, each with the jobs that need it. #>
    param([AllowEmptyCollection()][object[]]$Jobs = @(), [AllowEmptyCollection()][string[]]$AdditionalRoles = @())
    $byRole = [ordered]@{}
    foreach ($job in @($Jobs | Where-Object { -not $_.PSObject.Properties['Enabled'] -or [bool]$_.Enabled })) {
        $roles = if ($job.PSObject.Properties['RequiredGraphAppRoles']) { @(ConvertTo-SmartM365AcceptanceStringList $job.RequiredGraphAppRoles) } else { @() }
        foreach ($role in $roles) {
            if (-not $byRole.Contains($role)) { $byRole[$role] = [System.Collections.Generic.List[string]]::new() }
            $byRole[$role].Add([string]$job.Name)
        }
    }
    foreach ($role in @($AdditionalRoles | Where-Object { $_ })) {
        if (-not $byRole.Contains($role)) { $byRole[$role] = [System.Collections.Generic.List[string]]::new() }
        $byRole[$role].Add('(configuration)')
    }
    return @($byRole.Keys | Sort-Object | ForEach-Object { [pscustomobject]@{ Role = $_; Jobs = @($byRole[$_]) } })
}

function Get-SmartM365AcceptanceJobEligibility {
    <# Same rule as Test-SmartM365OrchestratorCapabilityMatch: every capability ready and every Graph role granted. #>
    param([AllowEmptyCollection()][object[]]$Jobs = @(), [AllowEmptyCollection()][string[]]$ReadyCapabilities = @(), [AllowEmptyCollection()][string[]]$GraphRoles = @())
    foreach ($job in @($Jobs | Where-Object { -not $_.PSObject.Properties['Enabled'] -or [bool]$_.Enabled })) {
        $capabilities = if ($job.PSObject.Properties['RequiredCapabilities']) { @(ConvertTo-SmartM365AcceptanceStringList $job.RequiredCapabilities) } else { @() }
        $roles = if ($job.PSObject.Properties['RequiredGraphAppRoles']) { @(ConvertTo-SmartM365AcceptanceStringList $job.RequiredGraphAppRoles) } else { @() }
        $missingCapabilities = @($capabilities | Where-Object { $_ -notin $ReadyCapabilities } | Sort-Object -Unique)
        $missingRoles = @($roles | Where-Object { $_ -notin $GraphRoles } | Sort-Object -Unique)
        [pscustomobject]@{
            Name = [string]$job.Name
            Eligible = ($missingCapabilities.Count -eq 0 -and $missingRoles.Count -eq 0)
            MissingCapabilities = $missingCapabilities
            MissingGraphAppRoles = $missingRoles
        }
    }
}

function Resolve-SmartM365AcceptanceTokens {
    <# Resolves {{Name}} tokens from the effective configuration and extra values. #>
    param([AllowEmptyString()][string]$Value, [Parameter(Mandatory = $true)][hashtable]$Tokens)
    $resolved = [string]$Value
    for ($pass = 0; $pass -lt 10 -and $resolved -match '\{\{'; $pass++) {
        $changed = $false
        foreach ($match in [regex]::Matches($resolved, '\{\{([A-Za-z0-9_.-]+)\}\}')) {
            $name = $match.Groups[1].Value
            if ($Tokens.ContainsKey($name) -and $null -ne $Tokens[$name] -and [string]$Tokens[$name] -ne '') {
                $resolved = $resolved.Replace($match.Value, [string]$Tokens[$name]); $changed = $true
            }
        }
        if (-not $changed) { break }
    }
    if ($resolved -match '\{\{') { throw "Configuration value cannot be resolved: $Value" }
    return $resolved
}

function Get-SmartM365AcceptanceShortResult {
    <# One short line per capability for the final table, read from the detail of the check. #>
    param([Parameter(Mandatory = $true)][string]$Capability, [string]$Status, [AllowEmptyString()][string]$Detail)
    $text = [string]$Detail
    $first = (($text -split ' \| ')[0] -split '; ')[0]
    if ($Status -eq 'SKIPPED') { return 'component not installed on this server' }
    $short = switch ($Capability) {
        'RunAccount' {
            $account = [regex]::Match($text, 'running as (\S+?);').Groups[1].Value
            if ($text -match 'account of the orchestrator scheduled task\)') { '{0} (orchestrator service account)' -f $account }
            elseif ($text -match 'scheduled task runs as (\S+?):') { '{0}, not the service account {1}' -f $account, $Matches[1] }
            else { $account }
        }
        'Modules' {
            $problems = @($text -split '; ' | Where-Object { $_ -match 'missing|below|not at|OneDrive profile copy|Windows PowerShell 5\.1:' })
            if ($problems.Count) { $problems -join '; ' } else { 'required modules present' }
        }
        'CodeSigning' {
            $unsigned = @($text -split '; ' | Where-Object { $_ -match '\.ps[dm]?1: ' -and $_ -notmatch ': Valid' } | ForEach-Object { ($_ -split ' by ')[0] })
            if ($unsigned.Count) { $unsigned -join '; ' } else { 'checked files signed (Valid)' }
        }
        'Network' {
            $reachable = [regex]::Matches($text, ': HTTP \d+').Count
            $down = @([regex]::Matches($text, '(?:^|; )(\S+?)(?: \(Teams\))?: unreachable') | ForEach-Object { $_.Groups[1].Value })
            '{0} of {1} endpoint(s) reachable{2}' -f $reachable, ($reachable + $down.Count), $(if ($down.Count) { '; unreachable: ' + ($down -join ', ') } else { '' })
        }
        'Clock' { $offset = [regex]::Match($text, 'local clock (\S+) s'); if ($offset.Success) { 'offset {0} s' -f $offset.Groups[1].Value } else { $first } }
        'Certificate' {
            $expiry = [regex]::Match($text, 'expires (\S+) \S+ UTC \((-?\d+) day')
            if ($Status -ne 'FAIL' -and $expiry.Success) { 'private key usable, expires {0} ({1} days)' -f $expiry.Groups[1].Value, $expiry.Groups[2].Value } else { $first }
        }
        'SharedRuntime' {
            $parts = @(@($text -split '; ' | Where-Object { $_ -match 'not accessible|unreachable|could not|No answer|witness: ' }))
            $smb = [regex]::Match($text, 'witness[^:]*: median (\S+) ms'); if ($smb.Success) { $parts += 'DATA SMB median {0} ms' -f $smb.Groups[1].Value }
            $code = [regex]::Match($text, 'read/metadata/list: median (\S+) ms'); if ($code.Success) { $parts += 'code {0} ms' -f $code.Groups[1].Value }
            $free = [regex]::Match($text, 'DATA free (\S+) %'); if ($free.Success) { $parts += 'DATA free {0} %' -f $free.Groups[1].Value }
            $parts -join ', '
        }
        { $_ -in @('AD', 'Graph', 'EXO', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell', 'Adaxes') } {
            $read = [regex]::Matches($text, ': (\d+) row\(s\)')
            $failedScopes = [regex]::Matches($text, ': read failed').Count
            if ($read.Count -eq 0 -and $failedScopes -eq 0) { $first }
            else {
                $rows = 0; foreach ($match in $read) { $rows += [int]$match.Groups[1].Value }
                $notes = @()
                if ($failedScopes) { $notes += '{0} read failure(s)' -f $failedScopes }
                if ($text -match '^non-blocking failure') { $notes += 'non-blocking' }
                if ($text -match 'unexpected format') { $notes += 'format warnings' }
                if ($text -match 'duplicated') { $notes += 'duplicates' }
                if ($text -match 'users skipped because') { $notes += 'unreadable users skipped' }
                if ($text -match 'check the Adaxes permissions') { $notes += 'check Adaxes permissions' }
                if ($text -match 'Windows PowerShell 5\.1 ActiveDirectory read failed') { $notes += 'AD read failed in Windows PowerShell 5.1' }
                '{0} of {1} data set(s) read, {2} row(s){3}' -f $read.Count, ($read.Count + $failedScopes), $rows, $(if ($notes.Count) { ' (' + ($notes -join ', ') + ')' } else { '' })
            }
        }
        'Mail' { ($text -split '; ') -join ', ' }
        default { $first }
    }
    $short = [string]$short
    if ($short.Length -gt 90) { $short = $short.Substring(0, 87) + '...' }
    return $short
}

function Format-SmartM365AcceptanceTable {
    <# Text table for the final summary. Returns { Status; Text } lines; Status (when the row has one) colors the line. #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string[]]$Columns,
        [AllowEmptyCollection()][object[]]$Rows = @(),
        [int]$MaxColumnWidth = 70
    )
    $cells = @(foreach ($row in $Rows) {
            , @(foreach ($column in $Columns) {
                    $property = $row.PSObject.Properties[$column]
                    $value = if ($null -eq $property -or $null -eq $property.Value) { '' } elseif ($property.Value -is [double] -or $property.Value -is [decimal]) { ([double]$property.Value).ToString('0.0', [Globalization.CultureInfo]::InvariantCulture) } elseif ($property.Value -is [IFormattable]) { $property.Value.ToString($null, [Globalization.CultureInfo]::InvariantCulture) } else { [string]$property.Value }
                    $value = $value -replace '[\r\n]+', ' '
                    if ($value.Length -gt $MaxColumnWidth) { $value.Substring(0, $MaxColumnWidth - 3) + '...' } else { $value }
                })
        })
    $widths = @(for ($index = 0; $index -lt $Columns.Count; $index++) {
            $width = $Columns[$index].Length
            foreach ($rowCells in $cells) { if ($rowCells[$index].Length -gt $width) { $width = $rowCells[$index].Length } }
            $width
        })
    $join = { param($values) ((@(for ($index = 0; $index -lt $values.Count; $index++) { ([string]$values[$index]).PadRight($widths[$index]) }) -join '  ').TrimEnd()) }
    $lines = [System.Collections.Generic.List[object]]::new()
    $lines.Add([pscustomobject]@{ Status = ''; Text = '' })
    $lines.Add([pscustomobject]@{ Status = ''; Text = $Title })
    $lines.Add([pscustomobject]@{ Status = ''; Text = (& $join $Columns) })
    $lines.Add([pscustomobject]@{ Status = ''; Text = (& $join @($widths | ForEach-Object { '-' * $_ })) })
    if ($cells.Count -eq 0) { $lines.Add([pscustomobject]@{ Status = ''; Text = '(none)' }) }
    for ($rowIndex = 0; $rowIndex -lt $cells.Count; $rowIndex++) {
        $statusProperty = $Rows[$rowIndex].PSObject.Properties['Status']
        $lines.Add([pscustomobject]@{ Status = $(if ($statusProperty) { [string]$statusProperty.Value } else { '' }); Text = (& $join $cells[$rowIndex]) })
    }
    return $lines.ToArray()
}

function New-SmartM365AcceptanceResultDocument {
    <# Result published for the orchestrator console (display only): one entry per tested capability.
       Status is what the election would see: a non-blocking failure is published as FAIL. #>
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Results,
        [Parameter(Mandatory = $true)][string]$ServerName,
        [string]$AcceptanceVersion = '',
        [string]$RunAccount = '',
        [bool]$ServiceAccount = $false,
        [AllowNull()]$Python = $null,
        [datetime]$GeneratedAtUtc = [datetime]::UtcNow
    )
    $entries = [System.Collections.Generic.List[object]]::new()
    foreach ($result in $Results) {
        $status = if ($result.PSObject.Properties['PublishedStatus'] -and [string]$result.PublishedStatus) { [string]$result.PublishedStatus } else { [string]$result.Status }
        $detail = [string]$result.Detail
        if ($detail.Length -gt 1000) { $detail = $detail.Substring(0, 997) + '...' }
        $entries.Add([ordered]@{ Name = [string]$result.Capability; Status = $status; Detail = $detail })
    }
    if ($null -ne $Python) { $entries.Add([ordered]@{ Name = 'Python'; Status = [string]$Python.Status; Detail = [string]$Python.Detail }) }
    return [ordered]@{
        SchemaVersion = 1
        ServerName = $ServerName.ToUpperInvariant()
        GeneratedAtUtc = $GeneratedAtUtc.ToUniversalTime().ToString('o')
        AcceptanceVersion = $AcceptanceVersion
        RunAccount = $RunAccount
        ServiceAccount = $ServiceAccount
        Results = $entries.ToArray()
    }
}

function Write-SmartM365AcceptanceResultFile {
    <# Writes <server folder>\Orchestrator-Acceptance.json (orchestrator service account) or
       Orchestrator-Acceptance.OtherAccount.json (any other account) atomically, UTF-8 without BOM. #>
    param([Parameter(Mandatory = $true)][string]$ServerFolder, [Parameter(Mandatory = $true)]$Document)
    $name = if ([bool]$Document.ServiceAccount) { 'Orchestrator-Acceptance.json' } else { 'Orchestrator-Acceptance.OtherAccount.json' }
    $null = [IO.Directory]::CreateDirectory($ServerFolder)
    $target = Join-Path $ServerFolder $name
    $temporary = Join-Path $ServerFolder ('{0}.{1}.tmp' -f $name, [guid]::NewGuid().ToString('N'))
    $backup = Join-Path $ServerFolder ('{0}.{1}.bak' -f $name, [guid]::NewGuid().ToString('N'))
    try {
        [IO.File]::WriteAllText($temporary, (ConvertTo-Json -InputObject $Document -Depth 6), [Text.UTF8Encoding]::new($false))
        if ([IO.File]::Exists($target)) { [IO.File]::Replace($temporary, $target, $backup) }
        else { [IO.File]::Move($temporary, $target) }
    }
    finally { if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }; if ([IO.File]::Exists($backup)) { [IO.File]::Delete($backup) } }
    return $target
}

function Publish-SmartM365AcceptanceSharePointResult {
    param([Parameter(Mandatory = $true)][string]$ResultPath, [Parameter(Mandatory = $true)][string]$WorkFolder)
    if ((Get-SmartM365JsonTransportPolicy).Mode -ne 'JsonText') {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $ResultPath -EnsureParentFolders | Out-Null
        return
    }
    $relative = Get-SmartM365SharePointRelativeFilePath -LocalFilePath $ResultPath
    if ($relative -notmatch '^DATA-ALL/') { throw 'The acceptance result is not under DATA-ALL; its SharePoint destination cannot be preserved.' }
    $staged = Join-Path $WorkFolder (($relative -replace '/', [IO.Path]::DirectorySeparatorChar) + '.txt')
    [IO.Directory]::CreateDirectory((Split-Path -Parent $staged)) | Out-Null
    Copy-Item -LiteralPath $ResultPath -Destination $staged -ErrorAction Stop
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $staged -EnsureParentFolders | Out-Null
}

function Get-SmartM365AcceptanceVerdict {
    param([AllowEmptyCollection()][object[]]$Results = @())
    $failed = @($Results | Where-Object Status -eq 'FAIL' | ForEach-Object Capability)
    $warned = @($Results | Where-Object Status -eq 'WARN' | ForEach-Object Capability)
    if ($failed.Count) { return [pscustomobject]@{ Status = 'FAIL'; Text = ('NOT READY: fix {0} before activating this server, or keep the jobs that need them away from it.' -f ($failed -join ', ')) } }
    if ($warned.Count) { return [pscustomobject]@{ Status = 'WARN'; Text = ('READY WITH WARNINGS: review {0}.' -f ($warned -join ', ')) } }
    return [pscustomobject]@{ Status = 'PASS'; Text = 'READY: every tested capability passed.' }
}

#endregion

#region Runtime helpers

function Add-AcceptanceMeasurement {
    <# Collects one measured value for the final summary (no row value of the tenant). #>
    param([string]$Section, [string]$Item, [AllowEmptyString()][string]$Value, [string]$Status = 'PASS')
    $script:Measurements.Add([pscustomobject]@{ Section = $Section; Item = $Item; Value = $Value; Status = $Status })
}

function Get-AcceptanceConfigValue {
    param([Parameter(Mandatory = $true)][string]$Name, $Default = '')
    $value = $null
    if ($script:TenantConfig -is [System.Collections.IDictionary]) { if ($script:TenantConfig.Contains($Name)) { $value = $script:TenantConfig[$Name] } }
    elseif ($null -ne $script:TenantConfig -and $script:TenantConfig.PSObject.Properties[$Name]) { $value = $script:TenantConfig.$Name }
    if ($null -eq $value -or ($value -is [string] -and ([string]::IsNullOrWhiteSpace($value) -or $value.Trim() -in @('__USE_GLOBAL__', 'USE_GLOBAL')))) { return $Default }
    return $value
}

function Write-AcceptanceLog {
    param([AllowEmptyString()][string]$Message, [string]$Level = 'INFO')
    if (Get-Command -Name WriteLog -ErrorAction SilentlyContinue) { WriteLog -Message $Message -Level $Level }
    else { Write-Host ('{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message) }
}

function New-AcceptanceResult {
    # Ready: what the election would see (PASS or WARN before a non-blocking failure is downgraded).
    param([string]$Name, [string]$Status, [string]$Detail, [double]$DurationSeconds = 0, [bool]$Ready = $false)
    [pscustomobject]@{ Capability = $Name; Status = $Status; DurationSeconds = [math]::Round($DurationSeconds, 1); Detail = $Detail; Ready = $Ready }
}

function Invoke-AcceptanceChildProcess {
    <# Runs a script in a separate PowerShell process with a time limit; the rows come back on stdout. #>
    param(
        [Parameter(Mandatory = $true)][string]$EnginePath,
        [Parameter(Mandatory = $true)][string]$ScriptText,
        [hashtable]$Environment = @{},
        [int]$TimeoutSeconds = 600
    )
    if (-not (Test-Path -LiteralPath $EnginePath -PathType Leaf)) {
        return [pscustomobject]@{ Ready = $false; ComponentMissing = $false; Detail = "PowerShell engine not found: $EnginePath"; Roles = @(); Datasets = @(); Info = [pscustomobject]@{} }
    }
    $fullScript = $script:ChildPreamble + "`n" + $ScriptText
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($fullScript))
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $EnginePath
    $startInfo.Arguments = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encoded"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    $startInfo.StandardErrorEncoding = [Text.UTF8Encoding]::new($false)
    foreach ($key in $Environment.Keys) { $startInfo.EnvironmentVariables[[string]$key] = [string]$Environment[$key] }
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) { throw "Could not start $EnginePath." }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit([math]::Max(1, $TimeoutSeconds) * 1000)) {
            try {
                if ($PSVersionTable.PSVersion.Major -ge 6) { $process.Kill($true) }
                else { $process.Kill() }
                $null = $process.WaitForExit(5000)
            }
            catch { Write-Verbose ("Child process kill failed: {0}" -f $_.Exception.Message) }
            # The output written before the stop names the step that hung.
            $partial = if ($stdoutTask.Wait(5000)) { $stdoutTask.Result } else { '' }
            $stage = Get-SmartM365AcceptanceLastStage -StandardOutput $partial
            $detail = "No result within $TimeoutSeconds second(s); the child process was stopped" + $(if ($stage) { " during: $stage." } else { ' before its first step.' })
            return [pscustomobject]@{ Ready = $false; ComponentMissing = $false; TimedOut = $true; Detail = $detail; Roles = @(); Datasets = @(); Info = [pscustomobject]@{} }
        }
        return ConvertFrom-SmartM365AcceptanceChildOutput -StandardOutput $stdoutTask.GetAwaiter().GetResult() -StandardError $stderrTask.GetAwaiter().GetResult() -ExitCode $process.ExitCode
    }
    catch {
        return [pscustomobject]@{ Ready = $false; ComponentMissing = $false; Detail = $_.Exception.Message; Roles = @(); Datasets = @(); Info = [pscustomobject]@{} }
    }
    finally { $process.Dispose() }
}

function Measure-AcceptanceSmbLatency {
    <# Create/write, read, metadata and delete calls in a witness folder; values in milliseconds. #>
    param([Parameter(Mandatory = $true)][string]$FolderPath, [int]$Iterations)
    $samples = [System.Collections.Generic.List[double]]::new()
    $payload = [byte[]]::new(4096)
    [Random]::new(42).NextBytes($payload)
    $watch = [System.Diagnostics.Stopwatch]::new()
    for ($index = 0; $index -lt $Iterations; $index++) {
        $file = [IO.Path]::Combine($FolderPath, ('witness.{0}.tmp' -f $index))
        $watch.Restart(); [IO.File]::WriteAllBytes($file, $payload); $samples.Add($watch.Elapsed.TotalMilliseconds)
        $watch.Restart(); $read = [IO.File]::ReadAllBytes($file); $samples.Add($watch.Elapsed.TotalMilliseconds)
        if ($read.Length -ne $payload.Length -or -not [Linq.Enumerable]::SequenceEqual($read, $payload)) { throw "Witness file content differs after reading back: $file" }
        $watch.Restart(); $null = [IO.File]::Exists($file); $samples.Add($watch.Elapsed.TotalMilliseconds)
        $watch.Restart(); [IO.File]::Delete($file); $samples.Add($watch.Elapsed.TotalMilliseconds)
    }
    return @($samples)
}

function Measure-AcceptanceReadLatency {
    <# Read-only calls on the code share: file read, metadata and folder listing. #>
    param([Parameter(Mandatory = $true)][string]$FilePath, [Parameter(Mandatory = $true)][string]$FolderPath, [int]$Iterations)
    $samples = [System.Collections.Generic.List[double]]::new()
    $watch = [System.Diagnostics.Stopwatch]::new()
    for ($index = 0; $index -lt $Iterations; $index++) {
        $watch.Restart(); $null = [IO.File]::ReadAllBytes($FilePath); $samples.Add($watch.Elapsed.TotalMilliseconds)
        $watch.Restart(); $null = [IO.File]::GetLastWriteTimeUtc($FilePath); $samples.Add($watch.Elapsed.TotalMilliseconds)
        $watch.Restart(); $enumerator = [IO.Directory]::EnumerateFileSystemEntries($FolderPath).GetEnumerator()
        try { $null = $enumerator.MoveNext() } finally { $enumerator.Dispose() }
        $samples.Add($watch.Elapsed.TotalMilliseconds)
    }
    return @($samples)
}

function Format-AcceptanceLatency {
    param([AllowEmptyCollection()][double[]]$Samples)
    return ('median {0} ms, p95 {1} ms over {2} call(s)' -f (Get-SmartM365AcceptancePercentile -Values $Samples -Percent 50), (Get-SmartM365AcceptancePercentile -Values $Samples -Percent 95), $Samples.Count)
}

function Invoke-AcceptanceWebProbe {
    param([Parameter(Mandatory = $true)][string]$Uri, [int]$TimeoutSeconds = 20)
    $target = [uri]$Uri
    $proxy = [System.Net.WebRequest]::DefaultWebProxy
    $proxyText = 'direct'
    # PowerShell 7 returns no proxy URI (null) for a direct connection.
    try {
        if ($null -ne $proxy -and -not $proxy.IsBypassed($target)) {
            $proxyUri = $proxy.GetProxy($target)
            if ($null -ne $proxyUri -and $proxyUri.AbsoluteUri -ne $target.AbsoluteUri) { $proxyText = 'proxy ' + $proxyUri.Authority }
        }
    }
    catch { $proxyText = 'proxy unknown' }
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $response = Invoke-WebRequest -Uri $Uri -Method Get -TimeoutSec $TimeoutSeconds -SkipHttpErrorCheck -UseBasicParsing -ErrorAction Stop
        $watch.Stop()
        $dateHeader = ''
        if ($response.Headers.ContainsKey('Date')) { $dateHeader = [string](@($response.Headers['Date'])[0]) }
        return [pscustomobject]@{ Uri = $Uri; Reachable = $true; StatusCode = [int]$response.StatusCode; Milliseconds = [math]::Round($watch.Elapsed.TotalMilliseconds); Proxy = $proxyText; Date = $dateHeader; Error = ''; ReceivedAtUtc = [datetime]::UtcNow }
    }
    catch {
        $watch.Stop()
        $messages = @(); $exception = $_.Exception
        while ($null -ne $exception) { if ($messages -notcontains $exception.Message) { $messages += $exception.Message }; $exception = $exception.InnerException }
        return [pscustomobject]@{ Uri = $Uri; Reachable = $false; StatusCode = 0; Milliseconds = [math]::Round($watch.Elapsed.TotalMilliseconds); Proxy = $proxyText; Date = ''; Error = ($messages -join ' -> '); ReceivedAtUtc = [datetime]::UtcNow }
    }
}

#endregion

#region Child scripts (no row is written anywhere: they come back on stdout only)

$script:ChildPreamble = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }
$sampleSize = [int]$env:SMART_SAMPLE_SIZE
$smallSampleSize = [int]$env:SMART_SMALL_SAMPLE_SIZE
function ConvertTo-AcceptanceText($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('o') }
    return [string]$Value
}
function Write-AcceptanceResult($Result) {
    'SMARTM365_ACCEPTANCE_JSON:' + (ConvertTo-Json -InputObject $Result -Compress -Depth 8)
}
function Get-AcceptanceErrorText($ErrorRecord) {
    # The outer message often only says "see inner exception": keep every distinct message of the chain.
    $messages = @(); $exception = $ErrorRecord.Exception
    while ($null -ne $exception) { if ($messages -notcontains $exception.Message) { $messages += $exception.Message }; $exception = $exception.InnerException }
    return ($messages -join ' -> ')
}
function Write-AcceptanceStage([string]$Text) {
    # Straight to stdout, outside the pipeline: the parent reads the last step after a time-out.
    [Console]::Out.WriteLine('SMARTM365_ACCEPTANCE_STAGE:' + $Text + ' (' + (Get-Date -Format 'HH:mm:ss') + ')')
    [Console]::Out.Flush()
}
function Invoke-AcceptanceDataset([string]$Name, [string]$Scope, [scriptblock]$Read) {
    Write-AcceptanceStage ('reading {0} {1}' -f $Name, $Scope)
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try { $rows = @(& $Read); $errorText = '' }
    catch { $rows = @(); $errorText = Get-AcceptanceErrorText $_ }
    @{ Name = $Name; Scope = $Scope; Error = $errorText; DurationSeconds = [math]::Round($watch.Elapsed.TotalSeconds, 1); Rows = @($rows) }
}
function Invoke-AcceptanceRetry([scriptblock]$Action) {
    for ($attempt = 1; ; $attempt++) {
        try { return & $Action }
        catch {
            $message = [string]$_.Exception.Message
            $transient = $message -notmatch '(?i)\b(?:401|403)\b|Unauthorized|Forbidden|AADSTS' -and $message -match '(?i)\b(?:408|429|500|502|503|504)\b|ServiceUnavailable|TooManyRequests|throttl|timed? ?out|temporarily unavailable|connection (?:was )?(?:reset|closed|aborted|forcibly closed)'
            if (-not $transient -or $attempt -ge 3) { throw }
            Start-Sleep -Seconds (5 * $attempt)
        }
    }
}
'@

$script:AdChildScript = @'
try {
    # -SkipEditionCheck: on Windows Server 2016 the module has no Core flag and loads through Windows PowerShell compatibility.
    if (-not (Get-Module -ListAvailable -SkipEditionCheck -Name ActiveDirectory)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'ActiveDirectory PowerShell module is not installed on this server.' }; return }
    Import-Module ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue
    Write-AcceptanceStage 'Get-ADForest'
    $forest = Get-ADForest -ErrorAction Stop
    $domains = @($forest.Domains | ForEach-Object { [string]$_ } | Sort-Object)
    $datasets = @(foreach ($domain in $domains) {
            Invoke-AcceptanceDataset 'ADUsers' $domain {
                Get-ADUser -Server $domain -Filter * -ResultSetSize $sampleSize -Properties UserPrincipalName, WhenCreated -ErrorAction Stop | ForEach-Object {
                    [pscustomobject]@{ ObjectGUID = ConvertTo-AcceptanceText $_.ObjectGUID; SamAccountName = ConvertTo-AcceptanceText $_.SamAccountName; UserPrincipalName = ConvertTo-AcceptanceText $_.UserPrincipalName; WhenCreated = ConvertTo-AcceptanceText $_.WhenCreated }
                }
            }
        })
    Write-AcceptanceResult @{ Ready = $true; Detail = ('forest {0}: {1} domain(s): {2}' -f $forest.Name, $domains.Count, ($domains -join ', ')); Datasets = $datasets }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = (Get-AcceptanceErrorText $_) } }
'@

$script:AdWindowsPowerShellScript = @'
try {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'ActiveDirectory PowerShell module is not installed for Windows PowerShell 5.1.' }; return }
    Import-Module ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue
    $domain = Get-ADDomain -ErrorAction Stop
    Write-AcceptanceResult @{ Ready = $true; Detail = ('Windows PowerShell {0}: ActiveDirectory read OK ({1})' -f $PSVersionTable.PSVersion, $domain.DNSRoot) }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = ('Windows PowerShell 5.1 ActiveDirectory read failed: {0}' -f $_.Exception.Message) } }
'@

$script:GraphChildScript = @'
try {
    if (-not (Get-Module -ListAvailable -SkipEditionCheck -Name Microsoft.Graph.Authentication)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'Microsoft.Graph.Authentication PowerShell module is not installed on this server.' }; return }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    Write-AcceptanceStage 'Graph sign-in'
    Invoke-AcceptanceRetry { Connect-MgGraph -TenantId $env:SMART_TENANT_ID -ClientId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -NoWelcome -ErrorAction Stop }
    $context = Get-MgContext
    $roles = @($context.Scopes | Sort-Object -Unique)
    function Get-AcceptanceGraphPage([string]$Uri) {
        $response = Invoke-AcceptanceRetry { Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop }
        @($response.value)
    }
    $datasets = @(
        Invoke-AcceptanceDataset 'GraphUsers' '' {
            Get-AcceptanceGraphPage ('https://graph.microsoft.com/v1.0/users?$top={0}&$select=id,userPrincipalName,createdDateTime' -f $sampleSize) | ForEach-Object {
                [pscustomobject]@{ Id = ConvertTo-AcceptanceText $_.id; UserPrincipalName = ConvertTo-AcceptanceText $_.userPrincipalName; CreatedDateTime = ConvertTo-AcceptanceText $_.createdDateTime }
            }
        }
        Invoke-AcceptanceDataset 'EntraDevices' '' {
            Get-AcceptanceGraphPage ('https://graph.microsoft.com/v1.0/devices?$top={0}&$select=id,deviceId,displayName' -f $sampleSize) | ForEach-Object {
                [pscustomobject]@{ Id = ConvertTo-AcceptanceText $_.id; DeviceId = ConvertTo-AcceptanceText $_.deviceId; DisplayName = ConvertTo-AcceptanceText $_.displayName }
            }
        }
        Invoke-AcceptanceDataset 'IntuneDevices' '' {
            Get-AcceptanceGraphPage ('https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top={0}&$select=id,deviceName,azureADDeviceId,lastSyncDateTime' -f $sampleSize) | ForEach-Object {
                [pscustomobject]@{ Id = ConvertTo-AcceptanceText $_.id; DeviceName = ConvertTo-AcceptanceText $_.deviceName; AzureADDeviceId = ConvertTo-AcceptanceText $_.azureADDeviceId; LastSyncDateTime = ConvertTo-AcceptanceText $_.lastSyncDateTime }
            }
        }
    )
    Write-AcceptanceResult @{ Ready = $true; Detail = ('Graph app-only sign-in with {0} application role(s)' -f $roles.Count); Roles = $roles; Datasets = $datasets }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = (Get-AcceptanceErrorText $_) } }
finally { try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { } }
'@

$script:ExoChildScript = @'
try {
    if (-not (Get-Module -ListAvailable -SkipEditionCheck -Name ExchangeOnlineManagement)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'ExchangeOnlineManagement PowerShell module is not installed on this server.' }; return }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
    Write-AcceptanceStage 'Exchange Online connection'
    Invoke-AcceptanceRetry { Connect-ExchangeOnline -AppId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -Organization $env:SMART_ORGANIZATION -ShowBanner:$false -ErrorAction Stop }
    $datasets = @(
        Invoke-AcceptanceDataset 'EXOMailboxes' '' {
            Get-EXOMailbox -ResultSize $sampleSize -ErrorAction Stop -WarningAction SilentlyContinue | ForEach-Object {
                [pscustomobject]@{ Guid = ConvertTo-AcceptanceText $_.Guid; PrimarySmtpAddress = ConvertTo-AcceptanceText $_.PrimarySmtpAddress; ExternalDirectoryObjectId = ConvertTo-AcceptanceText $_.ExternalDirectoryObjectId; RecipientTypeDetails = ConvertTo-AcceptanceText $_.RecipientTypeDetails }
            }
        }
    )
    Write-AcceptanceResult @{ Ready = $true; Detail = ('Exchange Online app-only connection to {0}' -f $env:SMART_ORGANIZATION); Datasets = $datasets }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = (Get-AcceptanceErrorText $_) } }
finally { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { } }
'@

$script:ExchangeOnPremChildScript = @'
try {
    if (-not (Get-PSSnapin -Registered -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'Exchange Management Shell snap-in Microsoft.Exchange.Management.PowerShell.SnapIn is not registered on this server.' }; return }
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
    # Session setting only: lets Get-Mailbox read every domain of the forest.
    Set-ADServerSettings -ViewEntireForest $true
    $servers = @(Get-ExchangeServer -ErrorAction Stop)
    $datasets = @(
        Invoke-AcceptanceDataset 'OnPremMailboxes' '' {
            Get-Mailbox -ResultSize $sampleSize -ErrorAction Stop -WarningAction SilentlyContinue | ForEach-Object {
                [pscustomobject]@{ Guid = ConvertTo-AcceptanceText $_.Guid; ExchangeGuid = ConvertTo-AcceptanceText $_.ExchangeGuid; PrimarySmtpAddress = ConvertTo-AcceptanceText $_.PrimarySmtpAddress; RecipientTypeDetails = ConvertTo-AcceptanceText $_.RecipientTypeDetails; Database = ConvertTo-AcceptanceText $_.Database }
            }
        }
    )
    Write-AcceptanceResult @{ Ready = $true; Detail = ('Windows PowerShell {0}, {1} Exchange server(s) visible' -f $PSVersionTable.PSVersion, $servers.Count); Datasets = $datasets }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = (Get-AcceptanceErrorText $_) } }
'@

$script:SharePointOnPremChildScript = @'
try {
    if (-not (Get-PSSnapin -Registered -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'SharePoint Server PowerShell snap-in Microsoft.SharePoint.PowerShell is not registered on this server.' }; return }
    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
    $farm = Get-SPFarm -ErrorAction Stop
    $webApplications = @(Get-SPWebApplication -ErrorAction Stop)
    $datasets = @(
        Invoke-AcceptanceDataset 'SharePointSites' '' {
            foreach ($site in @(Get-SPSite -Limit ([string]$smallSampleSize) -ErrorAction Stop)) {
                try { [pscustomobject]@{ Id = ConvertTo-AcceptanceText $site.ID; Url = ConvertTo-AcceptanceText $site.Url } }
                finally { $site.Dispose() }
            }
        }
    )
    Write-AcceptanceResult @{ Ready = $true; Detail = ('farm build {0}, {1} web application(s)' -f $farm.BuildVersion, $webApplications.Count); Datasets = $datasets }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = (Get-AcceptanceErrorText $_) } }
'@

$script:TeamsChildScript = @'
try {
    if (-not (Get-Module -ListAvailable -SkipEditionCheck -Name MicrosoftTeams)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'MicrosoftTeams PowerShell module is not installed on this server.' }; return }
    Import-Module MicrosoftTeams -MinimumVersion 4.7.1 -ErrorAction Stop
    Write-AcceptanceStage 'Teams connection'
    # Same calls as the Teams job (M365-TeamsPhonePstnUsage-Inventory): Get-CsTenant is not available app-only in recent modules.
    # MicrosoftTeams 7.x loads its Cs cmdlets during Connect-MicrosoftTeams: connect at script level, then look for them.
    $connection = Connect-MicrosoftTeams -TenantId $env:SMART_TENANT_ID -ApplicationId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -ErrorAction Stop
    if (-not (Get-Command -Name Get-CsPhoneNumberAssignment -ErrorAction SilentlyContinue)) {
        $csCount = @(Get-Command -Name 'Get-Cs*' -ErrorAction SilentlyContinue).Count
        throw ('Get-CsPhoneNumberAssignment is not available after Connect-MicrosoftTeams (MicrosoftTeams {0} from {1}; {2} Get-Cs* command(s) loaded).' -f (Get-Module -Name MicrosoftTeams).Version, (Get-Module -Name MicrosoftTeams).ModuleBase, $csCount)
    }
    $connectedTenant = [string](@($connection | Where-Object { $_ } | ForEach-Object { if ($_.PSObject.Properties['TenantId']) { $_.TenantId } }) | Select-Object -First 1)
    if ($connectedTenant -and $connectedTenant -ne [string]$env:SMART_TENANT_ID) { throw ('Connect-MicrosoftTeams returned TenantId {0}, not the configured tenant.' -f $connectedTenant) }
    $datasets = @(
        Invoke-AcceptanceDataset 'TeamsPhoneNumbers' '' {
            Get-CsPhoneNumberAssignment -Top $smallSampleSize -ErrorAction Stop | Select-Object -First $smallSampleSize | ForEach-Object {
                [pscustomobject]@{ TelephoneNumber = ConvertTo-AcceptanceText $_.TelephoneNumber; NumberType = ConvertTo-AcceptanceText $_.NumberType; ActivationState = ConvertTo-AcceptanceText $_.ActivationState }
            }
        }
    )
    Write-AcceptanceResult @{ Ready = $true; Detail = ('Teams app-only connection{0}; MicrosoftTeams {1} from {2}' -f $(if ($connectedTenant) { ', tenant identifier verified' } else { '' }), (Get-Module -Name MicrosoftTeams).Version, (Get-Module -Name MicrosoftTeams).ModuleBase); Datasets = $datasets }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = (Get-AcceptanceErrorText $_) } }
finally {
    try { if (Get-Command -Name Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) { Disconnect-MicrosoftTeams -ErrorAction Stop | Out-Null } } catch { }
}
'@

$script:AdaxesChildScript = @'
try {
    $moduleQuery = @{ ListAvailable = $true; Name = 'Adaxes' }
    if ($PSVersionTable.PSVersion.Major -ge 6) { $moduleQuery.SkipEditionCheck = $true }
    if (-not (Get-Module @moduleQuery)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'Adaxes PowerShell module is not installed on this server.' }; return }
    $services = @($env:SMART_ADAXES_SERVICES | ConvertFrom-Json | Where-Object { $_ })
    if ($services.Count -eq 0) { throw 'AdaxesService is not configured in the Adaxes collector configuration.' }
    $ldapFilter = [string]$env:SMART_ADAXES_LDAP_FILTER
    if ([string]::IsNullOrWhiteSpace($ldapFilter)) { throw 'LdapFilter is not configured in the Adaxes collector configuration.' }
    Write-AcceptanceStage 'importing the Adaxes module'
    Import-Module Adaxes -ErrorAction Stop
    $domains = @($env:SMART_ADAXES_DOMAINS | ConvertFrom-Json | Where-Object { $_ } | Sort-Object -Unique)
    $domainSource = 'configuration'
    if ($domains.Count -eq 0) {
        $domains = @([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Domains | ForEach-Object { [string]$_.Name } | Sort-Object -Unique)
        $domainSource = 'current forest'
    }
    $hasResultSetSize = (Get-Command -Name Get-AdmUser).Parameters.ContainsKey('ResultSetSize')
    # Like the collector: a user whose objectClass the account cannot read ends the whole read; it is
    # excluded by DN and the domain is read again (25 at most). Only the count is reported.
    $excluded = [System.Collections.Generic.List[string]]::new()
    function ConvertTo-AcceptanceLdapValue([string]$Value) {
        $builder = [System.Text.StringBuilder]::new()
        foreach ($char in $Value.ToCharArray()) {
            switch ($char) { '\' { [void]$builder.Append('\5c') } '*' { [void]$builder.Append('\2a') } '(' { [void]$builder.Append('\28') } ')' { [void]$builder.Append('\29') } ([char]0) { [void]$builder.Append('\00') } default { [void]$builder.Append($char) } }
        }
        return $builder.ToString()
    }
    $datasets = @(foreach ($domain in $domains) {
            Invoke-AcceptanceDataset 'AdaxesUsers' $domain {
                $errors = @()
                foreach ($service in $services) {
                    while ($true) {
                        $filter = if ($excluded.Count -eq 0) { $ldapFilter } else { '(&{0}{1})' -f $ldapFilter, (($excluded | ForEach-Object { '(!(distinguishedName={0}))' -f (ConvertTo-AcceptanceLdapValue $_) }) -join '') }
                        try {
                            Write-AcceptanceStage ('Get-AdmUser through service {0} on {1} ({2} excluded object(s))' -f $service, $domain, $excluded.Count)
                            $parameters = @{ AdaxesService = $service; Server = $domain; LDAPFilter = $filter; Properties = @('userPrincipalName', 'sAMAccountName', 'objectGUID'); ErrorAction = 'Stop' }
                            if ($hasResultSetSize) { $parameters['ResultSetSize'] = $sampleSize }
                            return @(Get-AdmUser @parameters | Select-Object -First $sampleSize | ForEach-Object {
                                    [pscustomobject]@{ ObjectGUID = ConvertTo-AcceptanceText $_.ObjectGUID; SamAccountName = ConvertTo-AcceptanceText $_.SamAccountName; UserPrincipalName = ConvertTo-AcceptanceText $_.UserPrincipalName }
                                })
                        }
                        catch {
                            $match = [regex]::Match([string]$_.Exception.Message, "attribute 'objectClass' of the directory object '(?<Dn>.+)' is unavailable")
                            if ($match.Success -and -not $excluded.Contains($match.Groups['Dn'].Value) -and $excluded.Count -lt 25) { $excluded.Add($match.Groups['Dn'].Value); continue }
                            $errors += ('{0}: {1}' -f $service, (Get-AcceptanceErrorText $_))
                            break
                        }
                    }
                }
                throw ($errors -join ' | ')
            }
        })
    Write-AcceptanceResult @{ Ready = $true; Detail = ('Adaxes module {0}; service(s) {1}' -f (Get-Module -Name Adaxes).Version, ($services -join ', ')); Datasets = $datasets; Info = @{ DomainSource = $domainSource; ExcludedObjects = $excluded.Count } }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = (Get-AcceptanceErrorText $_) } }
'@

$script:WindowsPowerShellInventoryScript = @'
$result = @{ Ready = $true; Detail = ''; Info = @{
        Version = [string]$PSVersionTable.PSVersion
        ActiveDirectory = [bool](Get-Module -ListAvailable -Name ActiveDirectory)
        ExchangeSnapIn = [bool](Get-PSSnapin -Registered -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)
        SharePointSnapIn = [bool](Get-PSSnapin -Registered -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)
    } }
Write-AcceptanceResult $result
'@

#endregion

#region Capability checks

function Get-AcceptanceTaskAccountMatch {
    <# Accounts of the orchestrator scheduled task of this computer and whether the running account is one of them. #>
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $taskAccounts = @(Get-ScheduledTask -TaskName 'SmartM365 Inventory Orchestrator*' -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.Principal.UserId } | Where-Object { $_ } | Sort-Object -Unique)
    $userName = $identity.Name.Split('\')[-1]
    $match = @($taskAccounts | Where-Object { [string]::Equals($_, $identity.Name, [StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($_.Split('\')[-1], $userName, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
    return [pscustomobject]@{ TaskAccounts = $taskAccounts; IsTaskAccount = $match }
}

function Test-AcceptanceRunAccount {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $elevated = ([Security.Principal.WindowsPrincipal]$identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $detail = 'running as {0}; elevated={1}' -f $identity.Name, $elevated
    if ($ExpectedAccount) {
        if (-not [string]::Equals($identity.Name, $ExpectedAccount.Trim(), [StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{ Status = 'FAIL'; Detail = "$detail; expected $ExpectedAccount" }
        }
        return [pscustomobject]@{ Status = 'PASS'; Detail = "$detail (expected account)" }
    }
    # Without -ExpectedAccount, the account of the orchestrator scheduled task of this computer is the reference.
    $task = Get-AcceptanceTaskAccountMatch
    if ($task.TaskAccounts.Count -eq 0) { return [pscustomobject]@{ Status = 'PASS'; Detail = "$detail; no orchestrator scheduled task on this computer to compare with" } }
    if ($task.IsTaskAccount) { return [pscustomobject]@{ Status = 'PASS'; Detail = "$detail (account of the orchestrator scheduled task)" } }
    return [pscustomobject]@{ Status = 'PASS'; Detail = ('{0}; the orchestrator scheduled task runs as {1}: run the test again as that account' -f $detail, ($task.TaskAccounts -join ', ')) }
}

function Test-AcceptanceModules {
    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add(('PowerShell {0}' -f $PSVersionTable.PSVersion))
    if ($PSVersionTable.PSVersion -lt [version]'7.2') { $statuses.Add('FAIL') }
    $installed = @(Get-Module -ListAvailable -SkipEditionCheck -Name @($script:ExpectedModules.Name + 'Microsoft.Graph.*') -ErrorAction SilentlyContinue | Select-Object Name, Version)
    $assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $installed -Expected $script:ExpectedModules
    $statuses.Add($assessment.Status); $parts.Add($assessment.Detail)
    Add-AcceptanceMeasurement -Section 'Modules' -Item 'PowerShell 7' -Value ([string]$PSVersionTable.PSVersion) -Status $(if ($PSVersionTable.PSVersion -lt [version]'7.2') { 'FAIL' } else { 'PASS' })
    foreach ($module in $script:ExpectedModules) {
        $latest = @($installed | Where-Object { [string]$_.Name -eq [string]$module.Name } | ForEach-Object { [version][string]$_.Version } | Sort-Object -Descending | Select-Object -First 1)
        $missingStatus = if ([bool]$module.Required) { 'FAIL' } else { 'WARN' }
        $moduleStatus = if ($latest.Count -eq 0) { $missingStatus } elseif ([string]$module.Minimum -and $latest[0] -lt [version][string]$module.Minimum) { $missingStatus } else { 'PASS' }
        $value = if ($latest.Count) { [string]$latest[0] } else { 'missing' }
        # Import-Module takes the first folder of PSModulePath that holds the module: a copy in a synced
        # OneDrive profile can shadow the machine copy and load without its files (on-demand placeholders).
        $firstCopy = @(Get-Module -ListAvailable -SkipEditionCheck -Name ([string]$module.Name) -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($firstCopy.Count -and [string]$firstCopy[0].ModuleBase -match '(?i)\\OneDrive') {
            $value += ' (loaded from a OneDrive profile copy: {0})' -f $firstCopy[0].ModuleBase
            if ($moduleStatus -eq 'PASS') { $moduleStatus = 'WARN'; $statuses.Add('WARN'); $parts.Add(('{0} is loaded from a OneDrive profile copy' -f $module.Name)) }
        }
        if ([string]$module.Minimum) { $value += ' (minimum {0})' -f $module.Minimum }
        if (-not [bool]$module.Required) { $value += ', optional' }
        Add-AcceptanceMeasurement -Section 'Modules' -Item ([string]$module.Name) -Value $value -Status $moduleStatus
    }
    $python = Test-AcceptancePython
    # Orchestrator capability Python (orchestrator 1.5.80): ready when Python 3.10+ answers.
    $script:PythonReady = ($python.Status -eq 'PASS')
    $script:PythonResult = $python
    $statuses.Add($python.Status); $parts.Add($python.Detail)
    Add-AcceptanceMeasurement -Section 'Modules' -Item 'Python 3.10+ (WorkplaceEvidence-Prepare)' -Value $python.Detail -Status $python.Status
    $windows = Invoke-AcceptanceChildProcess -EnginePath $script:WindowsPowerShellPath -ScriptText $script:WindowsPowerShellInventoryScript -Environment $script:ChildEnvironment -TimeoutSeconds 120
    if (-not [bool]$windows.Ready) {
        $statuses.Add('FAIL'); $parts.Add('Windows PowerShell 5.1: ' + $windows.Detail)
        Add-AcceptanceMeasurement -Section 'Modules' -Item 'Windows PowerShell 5.1' -Value ([string]$windows.Detail) -Status 'FAIL'
    }
    else {
        $info = $windows.Info
        if ([version][string]$info.Version -lt [version]'5.1') { $statuses.Add('FAIL') }
        $parts.Add(('Windows PowerShell {0} (ActiveDirectory={1}, Exchange snap-in={2}, SharePoint snap-in={3})' -f $info.Version, $info.ActiveDirectory, $info.ExchangeSnapIn, $info.SharePointSnapIn))
        $script:WindowsPowerShellInfo = $info
        Add-AcceptanceMeasurement -Section 'Modules' -Item 'Windows PowerShell' -Value ([string]$info.Version)
        Add-AcceptanceMeasurement -Section 'Modules' -Item 'ActiveDirectory (5.1)' -Value $(if ([bool]$info.ActiveDirectory) { 'present' } else { 'missing' }) -Status $(if ([bool]$info.ActiveDirectory) { 'PASS' } else { 'WARN' })
        Add-AcceptanceMeasurement -Section 'Modules' -Item 'Exchange snap-in (5.1)' -Value $(if ([bool]$info.ExchangeSnapIn) { 'registered' } else { 'not registered' }) -Status $(if ([bool]$info.ExchangeSnapIn) { 'PASS' } else { 'SKIPPED' })
        Add-AcceptanceMeasurement -Section 'Modules' -Item 'SharePoint snap-in (5.1)' -Value $(if ([bool]$info.SharePointSnapIn) { 'registered' } else { 'not registered' }) -Status $(if ([bool]$info.SharePointSnapIn) { 'PASS' } else { 'SKIPPED' })
    }
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join '; ') }
}

function Test-AcceptanceCodeSigning {
    $files = @(
        $PSCommandPath
        (Join-Path $script:SmartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psm1')
        (Join-Path $script:SmartInventoryRoot 'Orchestrator\SmartM365.Orchestrator.Distributed.psm1')
        (Join-Path $script:SmartInventoryRoot 'Orchestrator\SmartM365-Inventory-Orchestrator.ps1')
    )
    $policy = [string](Get-ExecutionPolicy)
    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add("execution policy $policy")
    foreach ($file in $files) {
        $signature = Get-AuthenticodeSignature -LiteralPath $file
        $status = [string]$signature.Status
        $fileStatus = if ($status -eq 'Valid') { 'PASS' } elseif ($policy -eq 'AllSigned') { 'FAIL' } else { 'WARN' }
        if ($fileStatus -ne 'PASS') { $statuses.Add($fileStatus) }
        $signer = if ($signature.SignerCertificate) { ' by ' + $signature.SignerCertificate.Subject } else { '' }
        $parts.Add(('{0}: {1}{2}' -f [IO.Path]::GetFileName($file), $status, $signer))
        Add-AcceptanceMeasurement -Section 'Code signing' -Item ([IO.Path]::GetFileName($file)) -Value $status -Status $fileStatus
    }
    Add-AcceptanceMeasurement -Section 'Code signing' -Item 'Execution policy' -Value $policy
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join '; ') }
}

function Test-AcceptanceNetwork {
    $endpoints = [System.Collections.Generic.List[string]]::new()
    foreach ($uri in @('https://login.microsoftonline.com/', 'https://graph.microsoft.com/v1.0/', 'https://outlook.office365.com/')) { $endpoints.Add($uri) }
    $sharePointHost = [string](Get-AcceptanceConfigValue -Name 'SharePointSiteHostname')
    if ($sharePointHost -and $sharePointHost -notmatch '(?i)^contoso\.') { $endpoints.Add("https://$sharePointHost/") }
    # Teams PowerShell endpoints: only the Teams job needs them, an unreachable one is a warning.
    $optional = @('https://teams.microsoft.com/', 'https://api.interfaces.records.teams.microsoft.com/')
    foreach ($uri in $optional) { $endpoints.Add($uri) }
    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($uri in $endpoints) {
        $probe = Invoke-AcceptanceWebProbe -Uri $uri
        $item = '{0}{1}' -f ([uri]$uri).Host, $(if ($uri -in $optional) { ' (Teams)' } else { '' })
        if (-not $probe.Reachable) {
            $endpointStatus = if ($uri -in $optional) { 'WARN' } else { 'FAIL' }
            $statuses.Add($endpointStatus)
            $parts.Add(('{0}: unreachable through {1}: {2}' -f $item, $probe.Proxy, $probe.Error))
            Add-AcceptanceMeasurement -Section 'Network' -Item $item -Value ('unreachable ({0}): {1}' -f $probe.Proxy, $probe.Error) -Status $endpointStatus
            continue
        }
        $endpointStatus = if ($probe.Milliseconds -ge 2000) { 'WARN' } else { 'PASS' }
        if ($endpointStatus -ne 'PASS') { $statuses.Add($endpointStatus) }
        $parts.Add(('{0}: HTTP {1} in {2} ms ({3})' -f ([uri]$uri).Host, $probe.StatusCode, $probe.Milliseconds, $probe.Proxy))
        Add-AcceptanceMeasurement -Section 'Network' -Item $item -Value ('HTTP {0} in {1} ms ({2})' -f $probe.StatusCode, $probe.Milliseconds, $probe.Proxy) -Status $endpointStatus
    }
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join '; ') }
}

function Test-AcceptanceClock {
    $sentAtUtc = [datetime]::UtcNow
    $probe = Invoke-AcceptanceWebProbe -Uri 'https://login.microsoftonline.com/'
    $parsed = [DateTimeOffset]::MinValue
    if (-not $probe.Reachable -or -not [DateTimeOffset]::TryParse($probe.Date, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) {
        return [pscustomobject]@{ Status = 'WARN'; Detail = 'no Date header from login.microsoftonline.com; time offset not measured' }
    }
    # The server stamps the response around the middle of the round trip; the header has a one-second precision.
    $localMiddle = $sentAtUtc.AddTicks([long](($probe.ReceivedAtUtc - $sentAtUtc).Ticks / 2))
    $offset = [math]::Round(($localMiddle - $parsed.UtcDateTime).TotalSeconds, 1)
    $status = Get-SmartM365AcceptanceClockStatus -OffsetSeconds $offset
    $source = ''
    try { $source = (& w32tm /query /source 2>$null | Select-Object -First 1) } catch { $source = '' }
    Add-AcceptanceMeasurement -Section 'Clock and certificate' -Item 'Clock offset from Entra' -Value ('{0} s (source {1})' -f $offset, $(if ($source) { ([string]$source).Trim() } else { 'unknown' })) -Status $status
    return [pscustomobject]@{ Status = $status; Detail = ('local clock {0} s from Entra time (warning 60 s, failure 300 s); time source {1}' -f $offset, $(if ($source) { ([string]$source).Trim() } else { 'unknown' })) }
}

function Test-AcceptanceCertificate {
    if (-not $script:Thumbprint) { return [pscustomobject]@{ Status = 'FAIL'; Detail = 'Thumbprint is not configured for this tenant.' } }
    $readiness = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $script:Thumbprint -Cache @{} -CacheMinutes 0 -ExpiryWarningDays 30
    $status = if (-not $readiness.Ready) { 'FAIL' } elseif ($readiness.ExpiresSoon) { 'WARN' } else { 'PASS' }
    Add-AcceptanceMeasurement -Section 'Clock and certificate' -Item 'Application certificate' -Value $(if ($readiness.Ready) { '{0}, private key usable, expires {1} ({2} days)' -f $readiness.StorePath, ([datetime]$readiness.NotAfterUtc).ToUniversalTime().ToString('yyyy-MM-dd'), $readiness.DaysRemaining } else { [string]$readiness.Reason }) -Status $status
    $hint = ''
    if ([string]$readiness.Reason -in @('NotFound', 'PrivateKeyInaccessible') -and -not $ExpectedAccount) {
        $hint = '; expected when the certificate is in the personal store of the service account: run the test as that account'
    }
    return [pscustomobject]@{ Status = $status; Detail = ('{0} (tested as {1}){2}' -f $readiness.Detail, [Environment]::UserName, $hint) }
}

function Test-AcceptanceSharedRuntime {
    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()

    $checks = @(foreach ($item in $script:DataRoots.GetEnumerator()) { [pscustomobject]@{ Name = $item.Key; Path = $item.Value; Access = 'ReadWrite'; PathType = 'Container' } })
    foreach ($probe in @(Test-SmartM365OrchestratorPathAccess -Checks $checks -TimeoutSeconds 30)) {
        if (-not $probe.Ready) { $statuses.Add('FAIL') }
        $parts.Add(('{0}: {1}' -f $probe.Name, $(if ($probe.Ready) { 'read/write OK' } else { $probe.Detail })))
        Add-AcceptanceMeasurement -Section 'Shares and disks' -Item $probe.Name -Value $(if ($probe.Ready) { 'read/write OK' } else { $probe.Detail }) -Status $(if ($probe.Ready) { 'PASS' } else { 'FAIL' })
    }

    # Witness folder under WorkspaceRootPath: created, measured, removed.
    $witnessFolder = Join-Path $script:WorkspaceRootPath ('.smartm365-acceptance.{0}.{1}' -f $env:COMPUTERNAME, [guid]::NewGuid().ToString('N'))
    try {
        $rootProbe = @(Test-SmartM365OrchestratorPathAccess -Checks @([pscustomobject]@{ Name = 'WorkspaceRootPath'; Path = $script:WorkspaceRootPath; Access = 'Read'; PathType = 'Container' }) -TimeoutSeconds 30)[0]
        if (-not $rootProbe.Ready) { throw ('WorkspaceRootPath is not reachable: {0}' -f $rootProbe.Detail) }
        $null = [IO.Directory]::CreateDirectory($witnessFolder)
        $samples = Measure-AcceptanceSmbLatency -FolderPath $witnessFolder -Iterations $SmbIterations
        $median = Get-SmartM365AcceptancePercentile -Values $samples -Percent 50
        $statuses.Add((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds $median -WarningMilliseconds $SmbWarningMilliseconds -FailureMilliseconds $SmbFailureMilliseconds))
        $parts.Add(('DATA witness write/read/metadata/delete: {0}' -f (Format-AcceptanceLatency $samples)))
        Add-AcceptanceMeasurement -Section 'Shares and disks' -Item 'DATA SMB latency' -Value (Format-AcceptanceLatency $samples) -Status (Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds $median -WarningMilliseconds $SmbWarningMilliseconds -FailureMilliseconds $SmbFailureMilliseconds)
    }
    catch {
        $statuses.Add('FAIL'); $parts.Add(('DATA witness: {0}' -f $_.Exception.Message))
        Add-AcceptanceMeasurement -Section 'Shares and disks' -Item 'DATA SMB latency' -Value $_.Exception.Message -Status 'FAIL'
    }
    finally {
        if ([IO.Directory]::Exists($witnessFolder)) {
            try { [IO.Directory]::Delete($witnessFolder, $true) }
            catch { $statuses.Add('WARN'); $parts.Add(('witness folder could not be removed: {0}: {1}' -f $witnessFolder, $_.Exception.Message)) }
        }
    }

    try {
        $codeFile = Join-Path $script:SmartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        $samples = Measure-AcceptanceReadLatency -FilePath $codeFile -FolderPath $script:SmartInventoryRoot -Iterations $SmbIterations
        $median = Get-SmartM365AcceptancePercentile -Values $samples -Percent 50
        $statuses.Add((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds $median -WarningMilliseconds $SmbWarningMilliseconds -FailureMilliseconds $SmbFailureMilliseconds))
        $location = if ($script:SmartM365Root.StartsWith('\\')) { 'code share' } else { 'local code folder' }
        $parts.Add(('{0} read/metadata/list: {1}' -f $location, (Format-AcceptanceLatency $samples)))
        Add-AcceptanceMeasurement -Section 'Shares and disks' -Item ('{0} latency' -f $location) -Value (Format-AcceptanceLatency $samples) -Status (Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds $median -WarningMilliseconds $SmbWarningMilliseconds -FailureMilliseconds $SmbFailureMilliseconds)
    }
    catch {
        $statuses.Add('FAIL'); $parts.Add(('code share read: {0}' -f $_.Exception.Message))
        Add-AcceptanceMeasurement -Section 'Shares and disks' -Item 'code share latency' -Value $_.Exception.Message -Status 'FAIL'
    }

    $space = Get-SmartM365OrchestratorVolumeFreeSpace -Path $script:WorkspaceRootPath -TimeoutSeconds 30
    if (-not $space.Ready) {
        $statuses.Add('WARN'); $parts.Add(('DATA free space not measured: {0}' -f $space.Error))
        Add-AcceptanceMeasurement -Section 'Shares and disks' -Item 'DATA free space' -Value ('not measured: {0}' -f $space.Error) -Status 'WARN'
    }
    else {
        $spaceStatus = if ([double]$space.FreePercent -lt 10) { 'FAIL' } elseif ([double]$space.FreePercent -lt 15) { 'WARN' } else { 'PASS' }
        $statuses.Add($spaceStatus)
        $parts.Add(('DATA free {0} % ({1} GiB of {2} GiB)' -f $space.FreePercent, $space.FreeGB, $space.TotalGB))
        Add-AcceptanceMeasurement -Section 'Shares and disks' -Item 'DATA free space' -Value ('{0} % ({1} GiB of {2} GiB)' -f $space.FreePercent, $space.FreeGB, $space.TotalGB) -Status $spaceStatus
    }
    foreach ($drive in @(@([IO.Path]::GetPathRoot($env:SystemRoot), [IO.Path]::GetPathRoot([IO.Path]::GetTempPath())) | Sort-Object -Unique)) {
        $info = [IO.DriveInfo]::new($drive)
        $freeGB = [math]::Round($info.AvailableFreeSpace / 1GB, 1)
        $driveStatus = if ($freeGB -lt 2) { 'FAIL' } elseif ($freeGB -lt 10) { 'WARN' } else { 'PASS' }
        $statuses.Add($driveStatus)
        $parts.Add(('local {0} free {1} GiB' -f $drive.TrimEnd('\'), $freeGB))
        Add-AcceptanceMeasurement -Section 'Shares and disks' -Item ('Local {0} free space' -f $drive.TrimEnd('\')) -Value ('{0} GiB' -f $freeGB) -Status $driveStatus
    }
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join '; ') }
}

function Test-AcceptanceAd {
    $child = Invoke-AcceptanceChildProcess -EnginePath $script:PowerShellPath -ScriptText $script:AdChildScript -Environment $script:ChildEnvironment -TimeoutSeconds $CapabilityTimeoutSeconds
    $assessment = Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize $SampleSize -Requested:($script:Requested -contains 'AD')
    if ($assessment.Status -eq 'SKIPPED') { return $assessment }
    # The Exchange on-premises collectors read AD from Windows PowerShell 5.1.
    $windows = Invoke-AcceptanceChildProcess -EnginePath $script:WindowsPowerShellPath -ScriptText $script:AdWindowsPowerShellScript -Environment $script:ChildEnvironment -TimeoutSeconds 180
    $windowsStatus = if ([bool]$windows.Ready) { 'PASS' } else { 'WARN' }
    Add-AcceptanceMeasurement -Section 'Modules' -Item 'ActiveDirectory read (5.1)' -Value ([string]$windows.Detail) -Status $windowsStatus
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($assessment.Status, $windowsStatus); Detail = ('{0} | {1}' -f $assessment.Detail, $windows.Detail); Extractions = $assessment.Extractions }
}

function Get-AcceptanceGraphChild {
    if ($null -eq $script:GraphChild) {
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $script:GraphChild = Invoke-AcceptanceChildProcess -EnginePath $script:PowerShellPath -ScriptText $script:GraphChildScript -Environment $script:ChildEnvironment -TimeoutSeconds $CapabilityTimeoutSeconds
        $script:GraphChildSeconds = $watch.Elapsed.TotalSeconds
    }
    return $script:GraphChild
}

function Test-AcceptanceGraph {
    if (-not $script:CloudConfigured) { return [pscustomobject]@{ Status = 'FAIL'; Detail = 'AppId, TenantId or Thumbprint is not configured.' } }
    return Get-SmartM365AcceptanceChildAssessment -Child (Get-AcceptanceGraphChild) -SampleSize $SampleSize -Requested:($script:Requested -contains 'Graph')
}

function Get-AcceptanceAdditionalGraphRole {
    $roles = @()
    if ($script:MailMode -in @('Graph', 'Both')) { $roles += 'Mail.Send' }
    $upload = Get-AcceptanceConfigValue -Name 'EnableSharePointUpload' -Default $false
    if ($upload -eq $true -or [string]$upload -eq 'true') { $roles += 'Sites.Selected' }
    return $roles
}

function Test-AcceptanceGraphPermissions {
    if (-not $script:CloudConfigured) { return [pscustomobject]@{ Status = 'FAIL'; Detail = 'AppId, TenantId or Thumbprint is not configured.' } }
    $graph = Get-AcceptanceGraphChild
    if (-not [bool]$graph.Ready) { return [pscustomobject]@{ Status = 'FAIL'; Detail = 'Graph sign-in failed: ' + $graph.Detail } }
    $granted = @($graph.Roles | ForEach-Object { [string]$_ })
    $required = @(Get-SmartM365AcceptanceRequiredGraphRole -Jobs $script:CatalogJobs -AdditionalRoles @(Get-AcceptanceAdditionalGraphRole))
    $missing = @($required | Where-Object { $_.Role -notin $granted })
    if ($missing.Count -eq 0) { return [pscustomobject]@{ Status = 'PASS'; Detail = ('{0} required role(s) granted among {1}' -f $required.Count, $granted.Count) } }
    $text = ($missing | ForEach-Object { '{0} (jobs: {1})' -f $_.Role, ($_.Jobs -join ', ') }) -join '; '
    # Roles belong to the application, not to the server: every server sees the same gap.
    return [pscustomobject]@{ Status = 'WARN'; Detail = ('{0} of {1} required role(s) missing for the application: {2}' -f $missing.Count, $required.Count, $text) }
}

function Test-AcceptanceExo {
    if (-not $script:CloudConfigured -or -not $script:Organization) { return [pscustomobject]@{ Status = 'FAIL'; Detail = 'AppId, Thumbprint or Exchange Online organization is not configured.' } }
    $child = Invoke-AcceptanceChildProcess -EnginePath $script:PowerShellPath -ScriptText $script:ExoChildScript -Environment $script:ChildEnvironment -TimeoutSeconds $CapabilityTimeoutSeconds
    return Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize $SampleSize -Requested:($script:Requested -contains 'EXO')
}

function Test-AcceptanceExchangeOnPrem {
    $child = Invoke-AcceptanceChildProcess -EnginePath $script:WindowsPowerShellPath -ScriptText $script:ExchangeOnPremChildScript -Environment $script:ChildEnvironment -TimeoutSeconds $CapabilityTimeoutSeconds
    return Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize $SampleSize -Requested:($script:Requested -contains 'ExchangeOnPrem')
}

function Test-AcceptanceSharePointOnPrem {
    $child = Invoke-AcceptanceChildProcess -EnginePath $script:WindowsPowerShellPath -ScriptText $script:SharePointOnPremChildScript -Environment $script:ChildEnvironment -TimeoutSeconds $CapabilityTimeoutSeconds
    return Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize $SmallSampleSize -Requested:($script:Requested -contains 'SharePointOnPrem')
}

function Test-AcceptanceTeams {
    $child = Invoke-AcceptanceChildProcess -EnginePath $script:PowerShellPath -ScriptText $script:TeamsChildScript -Environment $script:ChildEnvironment -TimeoutSeconds $CapabilityTimeoutSeconds
    if (-not [bool]$child.ComponentMissing -and -not $script:CloudConfigured) { return [pscustomobject]@{ Status = 'FAIL'; Detail = 'AppId, TenantId or Thumbprint is not configured.' } }
    return Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize $SmallSampleSize -Requested:($script:Requested -contains 'TeamsPowerShell')
}

function Get-AcceptanceCollectorConfig {
    # Read-only: a script configuration is read where its script reads it, never migrated here.
    param([Parameter(Mandatory = $true)][string]$RelativePath)
    $path = ConvertTo-SmartM365InventoryConfigLocation -Path (Join-Path $script:SmartInventoryRoot $RelativePath)
    $readPath = Get-SmartM365JsonReadPath -Path $path -Optional
    if (-not $readPath) { return $null }
    return (Read-SmartM365JsonDocument -Path $path).Document
}

function Get-AcceptanceOrchestratorServerFolder {
    <# <orchestrator shared data folder>\<SERVER>: where the orchestrator publishes its heartbeat and capabilities. #>
    $orchestratorFolder = Join-Path $script:SmartInventoryRoot 'Orchestrator'
    $config = Get-AcceptanceCollectorConfig -RelativePath 'Orchestrator\SmartM365-Inventory-Orchestrator.local.json'
    if ($null -eq $config) { $config = Get-Content -LiteralPath (Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator.local.json.template') -Raw | ConvertFrom-Json }
    $dataFolder = '{{DataAllRootPath}}\Orchestrator'
    if ($config.PSObject.Properties['OrchestratorDataFolderPath'] -and -not [string]::IsNullOrWhiteSpace([string]$config.OrchestratorDataFolderPath) -and [string]$config.OrchestratorDataFolderPath -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { $dataFolder = [string]$config.OrchestratorDataFolderPath }
    $tokens = @{} + $script:Tokens
    foreach ($item in $script:DataRoots.GetEnumerator()) { $tokens[$item.Key] = $item.Value }
    $dataFolder = Resolve-SmartM365AcceptanceTokens -Value $dataFolder -Tokens $tokens
    if ((Split-Path -Path $dataFolder -Leaf) -eq $env:COMPUTERNAME) { $dataFolder = Split-Path -Path $dataFolder -Parent }
    return (Join-Path $dataFolder $env:COMPUTERNAME.ToUpperInvariant())
}

function Get-AcceptanceAdaxesConfig { Get-AcceptanceCollectorConfig -RelativePath 'ActiveDirectoryInventory\SmartM365-ActiveDirectory-Adaxes-Inventory.local.json' }

function Test-AcceptancePython {
    <# Python 3.10+ for WorkplaceEvidence-Prepare: no orchestrator capability covers it, so the job could be elected here. #>
    $command = 'python'
    try {
        $prepareConfig = Get-AcceptanceCollectorConfig -RelativePath 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.local.json'
        if ($null -ne $prepareConfig -and $prepareConfig.PSObject.Properties['PreparedPythonCommand'] -and [string]$prepareConfig.PreparedPythonCommand) { $command = [string]$prepareConfig.PreparedPythonCommand }
    }
    catch { Write-Verbose ("Prepare configuration not read: {0}" -f $_.Exception.Message) }
    $python = Get-Command -Name $command -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    $impact = 'WorkplaceEvidence-Prepare would fail on this server'
    if ($null -eq $python) { return [pscustomobject]@{ Status = 'WARN'; Detail = ("'{0}' not found: {1}" -f $command, $impact) } }
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new($python.Source, '-c "import sys; print(\"%d.%d.%d\" % sys.version_info[:3])"')
    $startInfo.UseShellExecute = $false; $startInfo.RedirectStandardOutput = $true; $startInfo.RedirectStandardError = $true; $startInfo.CreateNoWindow = $true
    $process = [System.Diagnostics.Process]::Start($startInfo)
    try {
        $outputTask = $process.StandardOutput.ReadToEndAsync()
        if (-not $process.WaitForExit(20000)) { try { $process.Kill($true) } catch { Write-Verbose 'Python probe kill failed.' }; return [pscustomobject]@{ Status = 'WARN'; Detail = ('{0} did not answer within 20 s: {1}' -f $python.Source, $impact) } }
        $version = ([string]$outputTask.GetAwaiter().GetResult()).Trim()
    }
    finally { $process.Dispose() }
    # The Microsoft Store alias answers without a version.
    if ($version -notmatch '^\d+\.\d+\.\d+$') { return [pscustomobject]@{ Status = 'WARN'; Detail = ('{0} gives no Python version (Store alias?): {1}' -f $python.Source, $impact) } }
    if ([version]$version -lt [version]'3.10') { return [pscustomobject]@{ Status = 'WARN'; Detail = ('Python {0} at {1} is below 3.10: {2}' -f $version, $python.Source, $impact) } }
    return [pscustomobject]@{ Status = 'PASS'; Detail = ('Python {0} at {1}' -f $version, $python.Source) }
}

function Test-AcceptanceAdaxes {
    $config = $null
    $configError = ''
    try { $config = Get-AcceptanceAdaxesConfig } catch { $configError = $_.Exception.Message }
    $environment = $script:ChildEnvironment.Clone()
    $services = @(); $domains = @(); $filter = ''
    if ($null -ne $config) {
        if ($config.PSObject.Properties['AdaxesService']) { $services = @(ConvertTo-SmartM365AcceptanceStringList $config.AdaxesService) }
        if ($config.PSObject.Properties['Domains']) { $domains = @(ConvertTo-SmartM365AcceptanceStringList $config.Domains) }
        if ($config.PSObject.Properties['LdapFilter']) { $filter = [string]$config.LdapFilter }
    }
    $environment['SMART_ADAXES_SERVICES'] = ConvertTo-Json -InputObject @($services) -Compress
    $environment['SMART_ADAXES_LDAP_FILTER'] = $filter
    $domainSource = 'configuration'
    if ($domains.Count -eq 0) {
        $domainSource = 'current forest'
        try { $domains = @([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Domains | ForEach-Object { [string]$_.Name } | Sort-Object -Unique) }
        catch { return [pscustomobject]@{ Status = 'FAIL'; Detail = 'Domains is empty in the Adaxes configuration and the current forest cannot be read: ' + $_.Exception.Message } }
    }

    # One child per domain with its own time limit: a hung domain or service does not hide the others.
    $datasets = [System.Collections.Generic.List[object]]::new()
    $childDetail = ''
    $timedOutDomain = ''
    $excludedNotes = [System.Collections.Generic.List[string]]::new()
    $child = $null
    foreach ($domain in $domains) {
        if ($timedOutDomain) {
            $datasets.Add([pscustomobject]@{ Name = 'AdaxesUsers'; Scope = $domain; Error = "not read: $timedOutDomain timed out first"; DurationSeconds = 0; Rows = @() })
            continue
        }
        $environment['SMART_ADAXES_DOMAINS'] = ConvertTo-Json -InputObject @($domain) -Compress
        $child = Invoke-AcceptanceChildProcess -EnginePath $script:PowerShellPath -ScriptText $script:AdaxesChildScript -Environment $environment -TimeoutSeconds ([math]::Min($AdaxesDomainTimeoutSeconds, $CapabilityTimeoutSeconds))
        if ([bool]$child.ComponentMissing) { break }
        if (-not [bool]$child.Ready) {
            # A configuration or module error is the same for every domain: report it once.
            if (-not ($child.PSObject.Properties['TimedOut'] -and [bool]$child.TimedOut)) {
                if ($datasets.Count -eq 0) { break }
                $datasets.Add([pscustomobject]@{ Name = 'AdaxesUsers'; Scope = $domain; Error = [string]$child.Detail; DurationSeconds = 0; Rows = @() })
                continue
            }
            $timedOutDomain = $domain
            $datasets.Add([pscustomobject]@{ Name = 'AdaxesUsers'; Scope = $domain; Error = [string]$child.Detail; DurationSeconds = [math]::Min($AdaxesDomainTimeoutSeconds, $CapabilityTimeoutSeconds); Rows = @() })
            continue
        }
        if (-not $childDetail) { $childDetail = [string]$child.Detail }
        foreach ($dataset in @($child.Datasets)) { $datasets.Add($dataset) }
        if ($child.Info.PSObject.Properties['ExcludedObjects'] -and [int]$child.Info.ExcludedObjects -gt 0) { $excludedNotes.Add(('{0}: {1}' -f $domain, $child.Info.ExcludedObjects)) }
    }
    if ($null -ne $child -and ([bool]$child.ComponentMissing -or (-not [bool]$child.Ready -and $datasets.Count -eq 0))) {
        $assessment = Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize $SampleSize -Requested:($script:Requested -contains 'Adaxes')
    }
    else {
        $combined = [pscustomobject]@{ Ready = $true; ComponentMissing = $false; Detail = ('{0}; {1} domain(s) from the {2}' -f $childDetail, $domains.Count, $domainSource).TrimStart('; '); Roles = @(); Datasets = @($datasets); Info = [pscustomobject]@{} }
        $assessment = Get-SmartM365AcceptanceChildAssessment -Child $combined -SampleSize $SampleSize -Requested:($script:Requested -contains 'Adaxes')
    }
    if ($configError) { $assessment.Detail = 'Adaxes collector configuration unreadable: ' + $configError + ' | ' + $assessment.Detail }
    elseif ($null -eq $config -and $assessment.Status -ne 'SKIPPED') { $assessment.Detail = 'Adaxes collector configuration not found | ' + $assessment.Detail }
    if ($excludedNotes.Count -gt 0) {
        # The collector skips them the same way and ends with warnings.
        $assessment.Status = Get-SmartM365AcceptanceWorstStatus -Status @($assessment.Status, 'WARN')
        $assessment.Detail = ('users skipped because the running account cannot read their objectClass, as the collector does ({0}) | {1}' -f ($excludedNotes -join ', '), $assessment.Detail)
    }
    # Adaxes answers but hides the users the account may not read: almost nothing everywhere means missing rights.
    $readDatasets = @($datasets | Where-Object { -not [string]$_.Error })
    if ($readDatasets.Count -gt 0 -and @($readDatasets | Where-Object { @($_.Rows).Count -ge 2 }).Count -eq 0) {
        $assessment.Detail = ('{0} has no or almost no user in every domain read: check the Adaxes permissions of this account (security role assignment) | {1}' -f [Security.Principal.WindowsIdentity]::GetCurrent().Name, $assessment.Detail)
    }
    return $assessment
}

function Test-AcceptanceMail {
    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    $from = [string](Get-AcceptanceConfigValue -Name 'From')
    $to = if ($TestMailTo) { $TestMailTo } else { [string](Get-AcceptanceConfigValue -Name 'ErrorMailTo' -Default (Get-AcceptanceConfigValue -Name 'To')) }
    $smtpServer = [string](Get-AcceptanceConfigValue -Name 'SmtpServer')
    $parts.Add(('mode {0}' -f $script:MailMode))
    if (-not $from) { $statuses.Add('FAIL'); $parts.Add('From is not configured') }
    if (-not $to) { $statuses.Add('FAIL'); $parts.Add('ErrorMailTo/To is not configured') }

    if ($script:MailMode -in @('Graph', 'Both')) {
        if (-not $script:CloudConfigured) { $statuses.Add('FAIL'); $parts.Add('Graph app sign-in is not configured') }
        else {
            $graph = Get-AcceptanceGraphChild
            if (-not [bool]$graph.Ready) { $statuses.Add('FAIL'); $parts.Add('Graph sign-in failed') }
            elseif (@($graph.Roles) -notcontains 'Mail.Send') { $statuses.Add('FAIL'); $parts.Add('Mail.Send is not granted to the application') }
            else { $parts.Add('Mail.Send granted') }
        }
    }
    if ($script:MailMode -in @('SMTP', 'Both')) {
        if (-not $smtpServer) { $statuses.Add($(if ($script:MailMode -eq 'SMTP') { 'FAIL' } else { 'WARN' })); $parts.Add('SmtpServer is not configured') }
        else {
            $client = [System.Net.Sockets.TcpClient]::new()
            try {
                if (-not $client.ConnectAsync($smtpServer, 25).Wait(10000)) { throw 'no connection within 10 s' }
                $stream = $client.GetStream(); $stream.ReadTimeout = 10000
                $banner = [IO.StreamReader]::new($stream).ReadLine()
                if ([string]$banner -notmatch '^220') { throw "unexpected banner: $banner" }
                $parts.Add(('SMTP {0}:25 answered 220' -f $smtpServer))
            }
            catch { $statuses.Add('FAIL'); $parts.Add(('SMTP {0}:25: {1}' -f $smtpServer, $_.Exception.Message)) }
            finally { $client.Dispose() }
        }
    }

    if (-not $SendTestMail) { $parts.Add('no mail sent (use -SendTestMail)') }
    elseif ((Get-SmartM365AcceptanceWorstStatus -Status @($statuses)) -eq 'FAIL') { $parts.Add('test mail not sent: configuration incomplete') }
    else {
        try {
            $global:AppId = $script:AppId; $global:TenantId = $script:TenantId; $global:Thumbprint = $script:Thumbprint
            $body = '<p>SmartM365 orchestrator server acceptance test mail from {0} ({1}).</p>' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Send-SmartM365Mail -SendMailMode $script:MailMode -SmtpServer $smtpServer -From $from -To $to -Subject ('SmartM365 orchestrator server acceptance - {0}' -f $env:COMPUTERNAME) -BodyHtml $body -BodyAsHtml -MailPurpose Maintenance | Out-Null
            $parts.Add(('test mail sent to {0}' -f $to))
        }
        catch { $statuses.Add('FAIL'); $parts.Add(('test mail failed: {0}' -f $_.Exception.Message)) }
    }
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join '; ') }
}

#endregion

function Move-SmartM365AcceptanceLegacyLogs {
    param([string]$LogRoot, [string]$ServerFolder, [string]$ServerName, [string]$ScriptName)
    $pattern = '^' + [regex]::Escape($ScriptName) + '-\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.log$'
    foreach ($file in @(Get-ChildItem -LiteralPath $LogRoot -File -ErrorAction Stop | Where-Object { $_.Name -match $pattern })) {
        try {
            # Leave recent or unidentifiable files in place; they may belong to an active older run.
            if ($file.LastWriteTimeUtc -gt [datetime]::UtcNow.AddHours(-2)) { continue }
            $matches = @(Select-String -LiteralPath $file.FullName -Pattern 'MachineName:\s*(\S+)' -ErrorAction Stop)
            if ($matches.Count -ne 1 -or $matches[0].Matches[0].Groups[1].Value -ine $ServerName) { continue }
            $target = Join-Path $ServerFolder $file.Name
            if (-not (Test-Path -LiteralPath $target)) { Move-Item -LiteralPath $file.FullName -Destination $target -ErrorAction Stop }
        }
        catch { Write-Warning ('Legacy acceptance log was left in place: {0}: {1}' -f $file.Name, $_.Exception.Message) }
    }
}

function Remove-SmartM365AcceptanceOldLogs {
    param([string]$ServerFolder, [string]$ScriptName)
    $pattern = '^' + [regex]::Escape($ScriptName) + '-\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.log$'
    $files = @(Get-ChildItem -LiteralPath $ServerFolder -File -ErrorAction Stop | Where-Object { $_.Name -match $pattern } | Sort-Object LastWriteTimeUtc, Name -Descending)
    foreach ($file in @($files | Select-Object -Skip 2)) { Remove-Item -LiteralPath $file.FullName -ErrorAction Stop }
}

#region Main

$script:SmartInventoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd('\')
$script:SmartM365Root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\')
. (Join-Path $script:SmartM365Root 'Config\SmartM365-TenantContext.ps1')
$script:TenantConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
Import-Module (Join-Path $script:SmartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.117' -ErrorAction Stop
Import-Module (Join-Path $script:SmartInventoryRoot 'Orchestrator\SmartM365.Orchestrator.Distributed.psm1') -ErrorAction Stop

$scriptName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$workFolder = Join-Path ([IO.Path]::GetTempPath()) ('{0}_{1}_{2}' -f $scriptName, (Get-Date -Format 'yyyyMMdd_HHmmss'), [guid]::NewGuid().ToString('N').Substring(0, 8))
$resolvedLogRoot = & (Get-Module -Name SmartM365.Core) {
    Get-ModuleLocalConfigValue -Config (Get-ModuleLocalConfig) -Name 'LogAllRootPath' -DefaultValue ''
}
if ([string]::IsNullOrWhiteSpace($resolvedLogRoot)) { throw 'LogAllRootPath is required for acceptance logs.' }
$logRoot = Join-Path $resolvedLogRoot $scriptName
$serverLogFolder = Join-Path $logRoot ([Environment]::MachineName)
[IO.Directory]::CreateDirectory($serverLogFolder) | Out-Null
# Refuse overlapping runs on this server so retention never removes an active run.
$logLock = [IO.File]::Open((Join-Path $serverLogFolder '.acceptance.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
try {
    $previousRetention = $global:RetentionMaxLogs
    $global:RetentionMaxLogs = 0
    try { $null = InitializeScriptEnvironment -OutputPathInit $workFolder -LogFileName $scriptName -CallerScriptPath $PSCommandPath -LogSubfolder ([Environment]::MachineName) }
    finally { $global:RetentionMaxLogs = $previousRetention }
    if ([IO.Path]::GetFullPath($global:LogPath).TrimEnd('\') -ine [IO.Path]::GetFullPath($serverLogFolder).TrimEnd('\')) { throw 'Core and tenant log roots differ; log retention was not run.' }
    Move-SmartM365AcceptanceLegacyLogs -LogRoot $logRoot -ServerFolder $serverLogFolder -ServerName ([Environment]::MachineName) -ScriptName $scriptName
}
catch { $logLock.Dispose(); throw }
# SharePoint upload of the result file and the log follows the tenant setting, like the other scripts;
# no Teams notification is sent from log lines.
$global:AppId = [string](Get-AcceptanceConfigValue -Name 'AppId')
$global:TenantId = [string](Get-AcceptanceConfigValue -Name 'TenantId')
$global:Thumbprint = [string](Get-AcceptanceConfigValue -Name 'Thumbprint' -Default (Get-AcceptanceConfigValue -Name 'Thumb'))
$uploadSetting = Get-AcceptanceConfigValue -Name 'EnableSharePointUpload' -Default $false
$global:EnableSharePointUpload = (-not $NoSharePointUpload) -and ($uploadSetting -eq $true -or [string]$uploadSetting -in @('true', 'True', '1'))
$global:SharePointSiteHostname = [string](Get-AcceptanceConfigValue -Name 'SharePointSiteHostname')
$global:SharePointSitePath = [string](Get-AcceptanceConfigValue -Name 'SharePointSitePath')
$global:SharePointLibraryDisplayName = [string](Get-AcceptanceConfigValue -Name 'SharePointLibraryDisplayName' -Default 'Documents')
$global:SharePointTargetFolderPath = [string](Get-AcceptanceConfigValue -Name 'SharePointTargetFolderPath')
& (Get-Module -Name SmartM365.Core) { $script:SmartM365TeamsNotificationInProgress = $true }

$results = [System.Collections.Generic.List[object]]::new()
$script:PythonReady = $false
$script:PythonResult = $null
$extractionRows = [System.Collections.Generic.List[object]]::new()
$script:Measurements = [System.Collections.Generic.List[object]]::new()
$startedAt = Get-Date
$exitCode = 0
try {
    $script:AppId = [string](Get-AcceptanceConfigValue -Name 'AppId')
    $script:TenantId = [string](Get-AcceptanceConfigValue -Name 'TenantId')
    $script:Thumbprint = ([string](Get-AcceptanceConfigValue -Name 'Thumbprint' -Default (Get-AcceptanceConfigValue -Name 'Thumb')) -replace '\s', '').ToUpperInvariant()
    $script:Organization = [string](Get-AcceptanceConfigValue -Name 'ExchangeOnlineOrganization' -Default (Get-AcceptanceConfigValue -Name 'OrgDomain'))
    $script:CloudConfigured = [bool]($script:AppId -and $script:TenantId -and $script:Thumbprint)
    $smtpServerSetting = [string](Get-AcceptanceConfigValue -Name 'SmtpServer')
    $script:MailMode = [string](Get-AcceptanceConfigValue -Name 'SendMailMode' -Default $(if ($smtpServerSetting) { 'SMTP' } else { 'Graph' }))
    if ($script:MailMode -notin @('Graph', 'SMTP', 'Both')) { $script:MailMode = 'Graph' }

    $tokens = @{ SmartM365RootPath = $script:SmartM365Root }
    $script:Tokens = $tokens
    $configProperties = if ($script:TenantConfig -is [System.Collections.IDictionary]) { $script:TenantConfig.Keys } else { $script:TenantConfig.PSObject.Properties.Name }
    foreach ($name in $configProperties) { $value = Get-AcceptanceConfigValue -Name $name -Default $null; if ($value -is [string]) { $tokens[$name] = $value } }
    $script:WorkspaceRootPath = Resolve-SmartM365AcceptanceTokens -Value ([string](Get-AcceptanceConfigValue -Name 'WorkspaceRootPath')) -Tokens $tokens
    $script:DataRoots = [ordered]@{}
    foreach ($name in @('DataAllRootPath', 'LatestCsvFolderPath', 'LogAllRootPath')) {
        $script:DataRoots[$name] = Resolve-SmartM365AcceptanceTokens -Value ([string](Get-AcceptanceConfigValue -Name $name)) -Tokens $tokens
    }

    $script:PowerShellPath = (Get-Process -Id $PID).Path
    $script:WindowsPowerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $script:ChildEnvironment = @{
        SMART_TENANT_ID = $script:TenantId; SMART_APP_ID = $script:AppId; SMART_CERT_THUMBPRINT = $script:Thumbprint
        SMART_ORGANIZATION = $script:Organization; SMART_SAMPLE_SIZE = [string]$SampleSize; SMART_SMALL_SAMPLE_SIZE = [string]$SmallSampleSize
    }
    $script:GraphChild = $null
    $script:GraphChildSeconds = 0
    $catalog = Get-Content -LiteralPath (Join-Path $script:SmartInventoryRoot 'Orchestrator\Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
    $script:CatalogJobs = @($catalog.Jobs)

    $script:Requested = @($Capability | Where-Object { $_ -ne 'All' })
    $selected = if ($Capability -contains 'All') { $script:CapabilityOrder } else { @($script:CapabilityOrder | Where-Object { $_ -in $script:Requested }) }
    Write-AcceptanceLog ('Orchestrator server acceptance on {0}: tenant {1}, capabilities {2}, sample {3} (small {4}). Rows stay in memory; only the result file and the log may go to SharePoint.' -f $env:COMPUTERNAME, $Tenant, ($selected -join ', '), $SampleSize, $SmallSampleSize)

    $checks = @{
        RunAccount = { Test-AcceptanceRunAccount }; Modules = { Test-AcceptanceModules }; CodeSigning = { Test-AcceptanceCodeSigning }
        Network = { Test-AcceptanceNetwork }; Clock = { Test-AcceptanceClock }; Certificate = { Test-AcceptanceCertificate }
        SharedRuntime = { Test-AcceptanceSharedRuntime }; AD = { Test-AcceptanceAd }; Graph = { Test-AcceptanceGraph }
        GraphPermissions = { Test-AcceptanceGraphPermissions }; EXO = { Test-AcceptanceExo }; ExchangeOnPrem = { Test-AcceptanceExchangeOnPrem }
        SharePointOnPrem = { Test-AcceptanceSharePointOnPrem }; TeamsPowerShell = { Test-AcceptanceTeams }; Adaxes = { Test-AcceptanceAdaxes }
        Mail = { Test-AcceptanceMail }
    }
    foreach ($name in $selected) {
        Write-AcceptanceLog ('{0}: started.' -f $name)
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $graphBefore = $script:GraphChildSeconds
        try { $outcome = & $checks[$name] }
        catch { $outcome = [pscustomobject]@{ Status = 'FAIL'; Detail = $_.Exception.Message } }
        $seconds = $watch.Elapsed.TotalSeconds
        $status = [string]$outcome.Status
        $detail = [string]$outcome.Detail
        $isReady = $status -in @('PASS', 'WARN')
        $nonBlocking = $status -eq 'FAIL' -and $name -in $script:NonBlockingCapabilities -and $script:Requested -notcontains $name
        if ($nonBlocking) { $status = 'WARN'; $detail = 'non-blocking failure (the orchestrator keeps the jobs needing it on other servers) | ' + $detail }
        $result = New-AcceptanceResult -Name $name -Status $status -Detail $detail -DurationSeconds $seconds -Ready $isReady
        $result | Add-Member -NotePropertyName PublishedStatus -NotePropertyValue ([string]$outcome.Status)
        $results.Add($result)
        if ($outcome.PSObject.Properties['Extractions']) {
            foreach ($row in @($outcome.Extractions)) {
                $rowStatus = if ($nonBlocking -and $row.Status -eq 'FAIL') { 'WARN' } else { $row.Status }
                $extractionRows.Add([pscustomobject]@{ Status = $rowStatus; Capability = $name; Dataset = $row.Dataset; Scope = $row.Scope; Rows = $row.Rows; Seconds = $row.Seconds; Checks = $row.Checks })
            }
        }
        $level = switch ($result.Status) { 'PASS' { 'SUCCESS' } 'FAIL' { 'ERROR' } 'WARN' { 'WARNING' } default { 'INFO' } }
        $graphNote = if ($script:GraphChildSeconds -ne $graphBefore) { ' (includes the shared Graph sign-in and reads)' } else { '' }
        Write-AcceptanceLog ('{0}: {1} in {2} s{3}: {4}' -f $name, $result.Status, $result.DurationSeconds, $graphNote, $result.Detail) $level
    }

    $eligibleCount = -1
    $notEligibleRows = @()
    if ($Capability -contains 'All') {
        $ready = @($results | Where-Object { $_.Capability -in $script:OrchestratorCapabilities -and [bool]$_.Ready } | ForEach-Object Capability)
        if ($script:PythonReady) { $ready += 'Python' }
        $roles = if ($null -ne $script:GraphChild -and [bool]$script:GraphChild.Ready) { @($script:GraphChild.Roles | ForEach-Object { [string]$_ }) } else { @() }
        $eligibility = @(Get-SmartM365AcceptanceJobEligibility -Jobs $script:CatalogJobs -ReadyCapabilities $ready -GraphRoles $roles)
        $eligible = @($eligibility | Where-Object Eligible)
        Write-AcceptanceLog ('Catalog jobs this server could run: {0} of {1} (ready capabilities: {2}).' -f $eligible.Count, $eligibility.Count, $(if ($ready.Count) { $ready -join ', ' } else { 'none' }))
        foreach ($job in @($eligibility | Where-Object { -not $_.Eligible })) {
            $reasons = @()
            if ($job.MissingCapabilities.Count) { $reasons += 'capabilities ' + ($job.MissingCapabilities -join ', ') }
            if ($job.MissingGraphAppRoles.Count) { $reasons += 'Graph roles ' + ($job.MissingGraphAppRoles -join ', ') }
            Write-AcceptanceLog ('  not eligible: {0}: missing {1}' -f $job.Name, ($reasons -join '; '))
        }
        $eligibleCount = $eligible.Count
        $notEligibleRows = @($eligibility | Where-Object { -not $_.Eligible } | ForEach-Object {
                [pscustomobject]@{ Job = $_.Name; 'Missing capabilities' = ($_.MissingCapabilities -join ', '); 'Missing Graph roles' = ($_.MissingGraphAppRoles -join ', ') }
            })
    }

    # Final summary in several tables; the full details are in the lines above.
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $runAccount = @($results | Where-Object Capability -eq 'RunAccount' | Select-Object -First 1)
    $serverRows = @(
        [pscustomobject]@{ Item = 'Server'; Value = $env:COMPUTERNAME }
        [pscustomobject]@{ Item = 'Operating system'; Value = [string](Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption }
        [pscustomobject]@{ Item = 'Account'; Value = $(if ($runAccount.Count) { Get-SmartM365AcceptanceShortResult -Capability 'RunAccount' -Status $runAccount[0].Status -Detail $runAccount[0].Detail } else { $identity }) }
        [pscustomobject]@{ Item = 'Tenant profile'; Value = $Tenant }
        [pscustomobject]@{ Item = 'PowerShell'; Value = ('{0} (Windows PowerShell {1})' -f $PSVersionTable.PSVersion, $(if ($script:WindowsPowerShellInfo) { $script:WindowsPowerShellInfo.Version } else { 'not checked' })) }
        [pscustomobject]@{ Item = 'Code'; Value = $script:SmartM365Root }
        [pscustomobject]@{ Item = 'Data root'; Value = $script:WorkspaceRootPath }
        [pscustomobject]@{ Item = 'Acceptance script'; Value = ('{0}, sample {1} (small {2})' -f $global:SmartM365ScriptVersion, $SampleSize, $SmallSampleSize) }
        [pscustomobject]@{ Item = 'Started / duration'; Value = ('{0} / {1:hh\:mm\:ss}' -f $startedAt.ToString('yyyy-MM-dd HH:mm:ss'), ((Get-Date) - $startedAt)) }
    )
    $capabilityRows = @($results | ForEach-Object {
            [pscustomobject]@{ Status = $_.Status; Capability = $_.Capability; Seconds = [double]$_.DurationSeconds; Result = (Get-SmartM365AcceptanceShortResult -Capability $_.Capability -Status $_.Status -Detail $_.Detail) }
        })
    $tables = [System.Collections.Generic.List[object]]::new()
    $rule = [pscustomobject]@{ Status = ''; Text = ('=' * 110) }
    $tables.Add($rule)
    $tables.Add([pscustomobject]@{ Status = ''; Text = ('ACCEPTANCE SUMMARY - {0}' -f $env:COMPUTERNAME) })
    $tables.Add($rule)
    foreach ($line in @(Format-SmartM365AcceptanceTable -Title '1. Server and run' -Columns @('Item', 'Value') -Rows $serverRows -MaxColumnWidth 90)) { $tables.Add($line) }
    foreach ($line in @(Format-SmartM365AcceptanceTable -Title '2. Capabilities' -Columns @('Status', 'Capability', 'Seconds', 'Result') -Rows $capabilityRows -MaxColumnWidth 90)) { $tables.Add($line) }
    foreach ($line in @(Format-SmartM365AcceptanceTable -Title '3. Real extractions (rows kept in memory only, then dropped)' -Columns @('Status', 'Capability', 'Dataset', 'Scope', 'Rows', 'Seconds', 'Checks') -Rows @($extractionRows) -MaxColumnWidth 60)) { $tables.Add($line) }
    foreach ($section in @($script:Measurements | ForEach-Object Section | Select-Object -Unique)) {
        $sectionRows = @($script:Measurements | Where-Object Section -eq $section | ForEach-Object { [pscustomobject]@{ Status = $_.Status; Item = $_.Item; Value = $_.Value } })
        foreach ($line in @(Format-SmartM365AcceptanceTable -Title ('4. Measurements - {0}' -f $section) -Columns @('Status', 'Item', 'Value') -Rows $sectionRows -MaxColumnWidth 80)) { $tables.Add($line) }
    }
    if ($eligibleCount -ge 0) {
        $jobTitle = '5. Catalog jobs: {0} of {1} could run on this server; the others need' -f $eligibleCount, ($eligibleCount + $notEligibleRows.Count)
        foreach ($line in @(Format-SmartM365AcceptanceTable -Title $jobTitle -Columns @('Job', 'Missing capabilities', 'Missing Graph roles') -Rows $notEligibleRows -MaxColumnWidth 60)) { $tables.Add($line) }
    }
    $tables.Add([pscustomobject]@{ Status = ''; Text = '' })
    $tables.Add($rule)
    $verdict = Get-SmartM365AcceptanceVerdict -Results @($results)
    $tables.Add([pscustomobject]@{ Status = ''; Text = ('Totals: {0} PASS, {1} WARN, {2} FAIL, {3} SKIPPED' -f @($results | Where-Object Status -eq 'PASS').Count, @($results | Where-Object Status -eq 'WARN').Count, @($results | Where-Object Status -eq 'FAIL').Count, @($results | Where-Object Status -eq 'SKIPPED').Count) })
    $tables.Add([pscustomobject]@{ Status = $verdict.Status; Text = ('Verdict: {0}' -f $verdict.Text) })
    $tables.Add($rule)
    foreach ($line in $tables) {
        $level = switch ($line.Status) { 'PASS' { 'SUCCESS' } 'FAIL' { 'ERROR' } 'WARN' { 'WARNING' } default { 'INFO' } }
        Write-AcceptanceLog $line.Text $level
    }
    $overall = Get-SmartM365AcceptanceWorstStatus -Status @($results | Where-Object Status -ne 'SKIPPED' | ForEach-Object Status)
    if ($overall -eq 'FAIL') { $exitCode = 1 }

    # Published for the orchestrator console only after a full run: a partial run would hide capabilities.
    if ($Capability -contains 'All') {
        try {
            $document = New-SmartM365AcceptanceResultDocument -Results @($results) -ServerName $env:COMPUTERNAME -AcceptanceVersion ([string]$global:SmartM365ScriptVersion) -RunAccount ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -ServiceAccount ([bool](Get-AcceptanceTaskAccountMatch).IsTaskAccount) -Python $script:PythonResult
            $published = Write-SmartM365AcceptanceResultFile -ServerFolder (Get-AcceptanceOrchestratorServerFolder) -Document $document
            Write-AcceptanceLog ('Acceptance result published for the orchestrator console: {0}{1}' -f $published, $(if ($document.ServiceAccount) { '' } else { ' (not the orchestrator service account: shown apart, it does not validate the server)' }))
            if ($global:EnableSharePointUpload) { Publish-SmartM365AcceptanceSharePointResult -ResultPath $published -WorkFolder $workFolder }
            else { Write-AcceptanceLog ('SharePoint upload skipped: {0}.' -f $(if ($NoSharePointUpload) { '-NoSharePointUpload' } else { 'EnableSharePointUpload is off for this tenant' })) }
        }
        catch { Write-AcceptanceLog ('Acceptance result not published: {0}' -f $_.Exception.Message) 'WARNING' }
    }
    else { Write-AcceptanceLog 'Acceptance result not published: only a full run (-Capability All) is published for the orchestrator console.' }
    Write-AcceptanceLog ('Acceptance result: {0} ({1} PASS, {2} WARN, {3} FAIL, {4} SKIPPED).' -f $overall,
        @($results | Where-Object Status -eq 'PASS').Count, @($results | Where-Object Status -eq 'WARN').Count,
        @($results | Where-Object Status -eq 'FAIL').Count, @($results | Where-Object Status -eq 'SKIPPED').Count) $(if ($overall -eq 'FAIL') { 'ERROR' } elseif ($overall -eq 'WARN') { 'WARNING' } else { 'SUCCESS' })
}
catch {
    $exitCode = 1
    Write-AcceptanceLog ('Acceptance stopped: {0}' -f $_.Exception.Message) 'ERROR'
}
finally {
    $script:GraphChild = $null
    if ([IO.Directory]::Exists($workFolder)) {
        try { [IO.Directory]::Delete($workFolder, $true) }
        catch { Write-AcceptanceLog ('Work folder could not be removed: {0}: {1}' -f $workFolder, $_.Exception.Message) 'WARNING' }
    }
    $completion = if ($exitCode -ne 0) { 'Failed' } elseif (@($results | Where-Object Status -eq 'WARN').Count -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
    try { Complete-SmartM365ExecutionContext -Status $completion -EnsureLogUploadParentFolders }
    finally {
        try { Remove-SmartM365AcceptanceOldLogs -ServerFolder $serverLogFolder -ScriptName $scriptName }
        catch { Write-Warning ('Acceptance log retention failed: {0}' -f $_.Exception.Message) }
        $logLock.Dispose()
    }
}
exit $exitCode

#endregion

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDyD5tpto1AsDfT
# HjtFsxRp2qRi5+hKR/W+29T5O/pAjqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAlZvoNc2/KZpLagVHkBLHS
# 08j47ek5PRucTsGhhKAZAzANBgkqhkiG9w0BAQEFAASCAYAyWri0A0nkV0e24iZd
# 7MI+udel83+GWAlXSPlXQkg4qSPC4uHxfH+XXRF1zvUnkHUUfLpMT8hVORQlQuxa
# lC/IBbT56epqtqzAhOXSytRyzrQcnMd70G6uTPiG2RoldgrSG1Wa/z3fuPRtzNFT
# +Ubwsf9hAmGgd798G6pvw5K/UeWhFIEgH+5g337bSlaSoyPaTZ6svo9RfFluY2lg
# BkvuXQiyeBQGKTPSGrVZ3tc0xyk88eKlnzPobmN58SQH4A/Uz4pLH/ifBmTQmXK5
# jM7nExZHmzadh7oCQ+O1Zj6Iiy/fwM1QZP3KSHel0TT2OwmvVNtT8VoQkAePnRdT
# Y1fPzG6XfcjyD/5ZIl9x1q7J4i/jQYa1MNGQ/1SY0bnYbUJyIIlg+OJHMGnhvxi1
# EzXgiatXAu4Io2/Vx3BjeEhsVvHownr/jCksy2nykmQ0ednTJxXoFyc2U6B6tuH9
# GCN3x4ShQGghfUydwTIAVoy5PsiE6bDLjoogA5ahw7vtrlQ=
# SIG # End signature block
