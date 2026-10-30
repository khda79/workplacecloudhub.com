<#
.SYNOPSIS
Acceptance test of a server before it becomes an active SmartM365 orchestrator server.

.DESCRIPTION
For each capability the orchestrator can assign, the script runs a real but limited read
(about 100 objects), checks the returned data, then discards it: rows stay in memory and no CSV
is written. Nothing is written to the tenant: no SharePoint log upload, no Teams notification,
and no mail unless -SendTestMail is given.

Capabilities, in order:
- RunAccount: running identity; -ExpectedAccount fails the run under another account.
- Modules: PowerShell 7 and Windows PowerShell 5.1, module versions against the minimums used by
  the collectors, same version for every installed Microsoft.Graph module, snap-ins.
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
- TeamsPowerShell: tenant (same TenantId) and up to SmallSampleSize users.
- Adaxes: services, domains and LDAP filter of the Adaxes collector configuration; up to
  SampleSize users per domain.
- Mail: mail configuration (Graph Mail.Send or SMTP banner); a test mail only with -SendTestMail.

A capability whose component is not installed is SKIPPED, or FAIL when it was requested with
-Capability. Extractions run in a child PowerShell process with a time limit. The jobs of the
catalog that this server could run are listed at the end (same rule as the election).

.VERSION
1.0.0

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
    [ValidateRange(5, 500)][int]$SmbIterations = 20,
    [ValidateRange(1, 10000)][double]$SmbWarningMilliseconds = 15,
    [ValidateRange(1, 10000)][double]$SmbFailureMilliseconds = 50,
    [string]$ExpectedAccount = '',
    [switch]$SendTestMail,
    [string]$TestMailTo = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:CapabilityOrder = @('RunAccount', 'Modules', 'CodeSigning', 'Network', 'Clock', 'Certificate', 'SharedRuntime', 'AD', 'Graph', 'GraphPermissions', 'EXO', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell', 'Adaxes', 'Mail')
# Capability names the orchestrator election uses.
$script:OrchestratorCapabilities = @('SharedRuntime', 'AD', 'Graph', 'EXO', 'ExchangeOnPrem', 'SharePointOnPrem', 'TeamsPowerShell', 'Adaxes')
# Minimums taken from the collectors (Import-Module -MinimumVersion); an empty minimum checks presence only.
$script:ExpectedModules = @(
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Minimum = '2.37.0'; Required = $true }
    [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Minimum = '3.7.2'; Required = $true }
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
                    [pscustomobject]@{ Name = 'Database'; MinFilledPercent = 90 }) }
        }
        'SharePointSites' {
            return [pscustomobject]@{ AllowEmpty = $true; Columns = @(
                    [pscustomobject]@{ Name = 'Id'; Required = $true; Pattern = $guid; Unique = $true }
                    [pscustomobject]@{ Name = 'Url'; Required = $true; Pattern = '^https?://'; Unique = $true }) }
        }
        'TeamsUsers' {
            return [pscustomobject]@{ AllowEmpty = $false; Columns = @(
                    [pscustomobject]@{ Name = 'Identity'; Required = $true; Unique = $true }
                    [pscustomobject]@{ Name = 'UserPrincipalName'; Required = $true; Pattern = '^[^@\s]+@[^@\s]+$'; Unique = $true }) }
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

function Get-SmartM365AcceptanceChildAssessment {
    <# Turns a child result into a capability status. Datasets with the same Name and several scopes
       (domains) are partial when only some scopes fail: WARN; all scopes failing is FAIL. #>
    param([Parameter(Mandatory = $true)]$Child, [Parameter(Mandatory = $true)][int]$SampleSize, [switch]$Requested)

    if ([bool]$Child.ComponentMissing) {
        return [pscustomobject]@{ Status = $(if ($Requested) { 'FAIL' } else { 'SKIPPED' }); Detail = [string]$Child.Detail }
    }
    if (-not [bool]$Child.Ready) { return [pscustomobject]@{ Status = 'FAIL'; Detail = [string]$Child.Detail } }

    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    if ([string]$Child.Detail) { $parts.Add([string]$Child.Detail) }
    foreach ($group in @(@($Child.Datasets) | Group-Object Name)) {
        $failedScopes = 0
        foreach ($dataset in $group.Group) {
            $label = if ([string]$dataset.Scope) { '{0} {1}' -f $dataset.Name, $dataset.Scope } else { [string]$dataset.Name }
            if ([string]$dataset.Error) {
                $failedScopes++
                $parts.Add(('{0}: read failed: {1}' -f $label, $dataset.Error))
                continue
            }
            $spec = Get-SmartM365AcceptanceDatasetSpec -Name ([string]$dataset.Name)
            $check = Test-SmartM365AcceptanceDataset -Rows @($dataset.Rows) -SampleSize $SampleSize -Columns $spec.Columns -AllowEmpty:([bool]$spec.AllowEmpty)
            $statuses.Add($check.Status)
            $text = '{0}: {1} row(s) in {2} s' -f $label, $check.Count, $dataset.DurationSeconds
            if ($check.Issues.Count -gt 0) { $text += ' (' + ($check.Issues -join '; ') + ')' }
            $parts.Add($text)
        }
        if ($failedScopes -gt 0) { $statuses.Add($(if ($failedScopes -eq $group.Count) { 'FAIL' } else { 'WARN' })) }
    }
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join ' | ') }
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
        $roles = if ($job.PSObject.Properties['RequiredGraphAppRoles']) { ConvertTo-SmartM365AcceptanceStringList $job.RequiredGraphAppRoles } else { @() }
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
        $capabilities = if ($job.PSObject.Properties['RequiredCapabilities']) { ConvertTo-SmartM365AcceptanceStringList $job.RequiredCapabilities } else { @() }
        $roles = if ($job.PSObject.Properties['RequiredGraphAppRoles']) { ConvertTo-SmartM365AcceptanceStringList $job.RequiredGraphAppRoles } else { @() }
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

#endregion

#region Runtime helpers

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
    param([string]$Name, [string]$Status, [string]$Detail, [double]$DurationSeconds = 0)
    [pscustomobject]@{ Capability = $Name; Status = $Status; DurationSeconds = [math]::Round($DurationSeconds, 1); Detail = $Detail }
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
    foreach ($key in $Environment.Keys) { $startInfo.Environment[[string]$key] = [string]$Environment[$key] }
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) { throw "Could not start $EnginePath." }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit([math]::Max(1, $TimeoutSeconds) * 1000)) {
            try { $process.Kill($true) } catch { Write-Verbose ("Child process kill failed: {0}" -f $_.Exception.Message) }
            return [pscustomobject]@{ Ready = $false; ComponentMissing = $false; Detail = "No result within $TimeoutSeconds second(s); the child process was stopped."; Roles = @(); Datasets = @(); Info = [pscustomobject]@{} }
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
    if ($null -ne $proxy -and -not $proxy.IsBypassed($target)) { $proxyText = 'proxy ' + $proxy.GetProxy($target).Authority }
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
        return [pscustomobject]@{ Uri = $Uri; Reachable = $false; StatusCode = 0; Milliseconds = [math]::Round($watch.Elapsed.TotalMilliseconds); Proxy = $proxyText; Date = ''; Error = $_.Exception.Message; ReceivedAtUtc = [datetime]::UtcNow }
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
function Invoke-AcceptanceDataset([string]$Name, [string]$Scope, [scriptblock]$Read) {
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try { $rows = @(& $Read); $errorText = '' }
    catch { $rows = @(); $errorText = $_.Exception.Message }
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
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'ActiveDirectory PowerShell module is not installed on this server.' }; return }
    Import-Module ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue
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
catch { Write-AcceptanceResult @{ Ready = $false; Detail = $_.Exception.Message } }
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
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'Microsoft.Graph.Authentication PowerShell module is not installed on this server.' }; return }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
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
catch { Write-AcceptanceResult @{ Ready = $false; Detail = $_.Exception.Message } }
finally { try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { } }
'@

$script:ExoChildScript = @'
try {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'ExchangeOnlineManagement PowerShell module is not installed on this server.' }; return }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
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
catch { Write-AcceptanceResult @{ Ready = $false; Detail = $_.Exception.Message } }
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
catch { Write-AcceptanceResult @{ Ready = $false; Detail = $_.Exception.Message } }
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
catch { Write-AcceptanceResult @{ Ready = $false; Detail = $_.Exception.Message } }
'@

$script:TeamsChildScript = @'
try {
    if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'MicrosoftTeams PowerShell module is not installed on this server.' }; return }
    Import-Module MicrosoftTeams -MinimumVersion 4.7.1 -ErrorAction Stop
    Invoke-AcceptanceRetry { Connect-MicrosoftTeams -TenantId $env:SMART_TENANT_ID -ApplicationId $env:SMART_APP_ID -CertificateThumbprint $env:SMART_CERT_THUMBPRINT -ErrorAction Stop | Out-Null }
    $tenant = Get-CsTenant -ErrorAction Stop
    if ([string]$tenant.TenantId -ne [string]$env:SMART_TENANT_ID) { throw ('Get-CsTenant returned TenantId {0}, not the configured tenant.' -f $tenant.TenantId) }
    $datasets = @(
        Invoke-AcceptanceDataset 'TeamsUsers' '' {
            Get-CsOnlineUser -ResultSize $smallSampleSize -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{ Identity = ConvertTo-AcceptanceText $_.Identity; UserPrincipalName = ConvertTo-AcceptanceText $_.UserPrincipalName }
            }
        }
    )
    Write-AcceptanceResult @{ Ready = $true; Detail = 'Teams app-only connection, tenant identifier verified'; Datasets = $datasets }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = $_.Exception.Message } }
finally {
    try { if (Get-Command -Name Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue) { Disconnect-MicrosoftTeams -ErrorAction Stop | Out-Null } } catch { }
}
'@

$script:AdaxesChildScript = @'
try {
    if (-not (Get-Module -ListAvailable -Name Adaxes)) { Write-AcceptanceResult @{ Ready = $false; ComponentMissing = $true; Detail = 'Adaxes PowerShell module is not installed on this server.' }; return }
    $services = @($env:SMART_ADAXES_SERVICES | ConvertFrom-Json | Where-Object { $_ })
    if ($services.Count -eq 0) { throw 'AdaxesService is not configured in the Adaxes collector configuration.' }
    $ldapFilter = [string]$env:SMART_ADAXES_LDAP_FILTER
    if ([string]::IsNullOrWhiteSpace($ldapFilter)) { throw 'LdapFilter is not configured in the Adaxes collector configuration.' }
    Import-Module Adaxes -ErrorAction Stop
    $domains = @($env:SMART_ADAXES_DOMAINS | ConvertFrom-Json | Where-Object { $_ } | Sort-Object -Unique)
    $domainSource = 'configuration'
    if ($domains.Count -eq 0) {
        $domains = @([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Domains | ForEach-Object { [string]$_.Name } | Sort-Object -Unique)
        $domainSource = 'current forest'
    }
    $hasResultSetSize = (Get-Command -Name Get-AdmUser).Parameters.ContainsKey('ResultSetSize')
    $datasets = @(foreach ($domain in $domains) {
            Invoke-AcceptanceDataset 'AdaxesUsers' $domain {
                $errors = @()
                foreach ($service in $services) {
                    try {
                        $parameters = @{ AdaxesService = $service; Server = $domain; LDAPFilter = $ldapFilter; Properties = @('userPrincipalName', 'sAMAccountName', 'objectGUID'); ErrorAction = 'Stop' }
                        if ($hasResultSetSize) { $parameters['ResultSetSize'] = $sampleSize }
                        return @(Get-AdmUser @parameters | Select-Object -First $sampleSize | ForEach-Object {
                                [pscustomobject]@{ ObjectGUID = ConvertTo-AcceptanceText $_.ObjectGUID; SamAccountName = ConvertTo-AcceptanceText $_.SamAccountName; UserPrincipalName = ConvertTo-AcceptanceText $_.UserPrincipalName }
                            })
                    }
                    catch { $errors += ('{0}: {1}' -f $service, $_.Exception.Message) }
                }
                throw ($errors -join ' | ')
            }
        })
    Write-AcceptanceResult @{ Ready = $true; Detail = ('Adaxes module {0}; service(s) {1}; {2} domain(s) from the {3}' -f (Get-Module -Name Adaxes).Version, ($services -join ', '), $domains.Count, $domainSource); Datasets = $datasets }
}
catch { Write-AcceptanceResult @{ Ready = $false; Detail = $_.Exception.Message } }
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
    return [pscustomobject]@{ Status = 'PASS'; Detail = "$detail; run it again under the orchestrator service account with -ExpectedAccount" }
}

function Test-AcceptanceModules {
    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add(('PowerShell {0}' -f $PSVersionTable.PSVersion))
    if ($PSVersionTable.PSVersion -lt [version]'7.2') { $statuses.Add('FAIL') }
    $installed = @(Get-Module -ListAvailable -Name @($script:ExpectedModules.Name + 'Microsoft.Graph.*') -ErrorAction SilentlyContinue | Select-Object Name, Version)
    $assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $installed -Expected $script:ExpectedModules
    $statuses.Add($assessment.Status); $parts.Add($assessment.Detail)
    $windows = Invoke-AcceptanceChildProcess -EnginePath $script:WindowsPowerShellPath -ScriptText $script:WindowsPowerShellInventoryScript -Environment $script:ChildEnvironment -TimeoutSeconds 120
    if (-not [bool]$windows.Ready) { $statuses.Add('FAIL'); $parts.Add('Windows PowerShell 5.1: ' + $windows.Detail) }
    else {
        $info = $windows.Info
        if ([version][string]$info.Version -lt [version]'5.1') { $statuses.Add('FAIL') }
        $parts.Add(('Windows PowerShell {0} (ActiveDirectory={1}, Exchange snap-in={2}, SharePoint snap-in={3})' -f $info.Version, $info.ActiveDirectory, $info.ExchangeSnapIn, $info.SharePointSnapIn))
        $script:WindowsPowerShellInfo = $info
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
        if ($status -ne 'Valid') { $statuses.Add($(if ($policy -eq 'AllSigned') { 'FAIL' } else { 'WARN' })) }
        $signer = if ($signature.SignerCertificate) { ' by ' + $signature.SignerCertificate.Subject } else { '' }
        $parts.Add(('{0}: {1}{2}' -f [IO.Path]::GetFileName($file), $status, $signer))
    }
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($statuses); Detail = ($parts -join '; ') }
}

function Test-AcceptanceNetwork {
    $endpoints = [System.Collections.Generic.List[string]]::new()
    foreach ($uri in @('https://login.microsoftonline.com/', 'https://graph.microsoft.com/v1.0/', 'https://outlook.office365.com/')) { $endpoints.Add($uri) }
    $sharePointHost = [string](Get-AcceptanceConfigValue -Name 'SharePointSiteHostname')
    if ($sharePointHost -and $sharePointHost -notmatch '(?i)^contoso\.') { $endpoints.Add("https://$sharePointHost/") }
    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($uri in $endpoints) {
        $probe = Invoke-AcceptanceWebProbe -Uri $uri
        if (-not $probe.Reachable) { $statuses.Add('FAIL'); $parts.Add(('{0}: unreachable through {1}: {2}' -f ([uri]$uri).Host, $probe.Proxy, $probe.Error)); continue }
        if ($probe.Milliseconds -ge 2000) { $statuses.Add('WARN') }
        $parts.Add(('{0}: HTTP {1} in {2} ms ({3})' -f ([uri]$uri).Host, $probe.StatusCode, $probe.Milliseconds, $probe.Proxy))
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
    return [pscustomobject]@{ Status = $status; Detail = ('local clock {0} s from Entra time (warning 60 s, failure 300 s); time source {1}' -f $offset, $(if ($source) { ([string]$source).Trim() } else { 'unknown' })) }
}

function Test-AcceptanceCertificate {
    if (-not $script:Thumbprint) { return [pscustomobject]@{ Status = 'FAIL'; Detail = 'Thumbprint is not configured for this tenant.' } }
    $readiness = Get-SmartM365OrchestratorCachedCertificateReadiness -Thumbprint $script:Thumbprint -Cache @{} -CacheMinutes 0 -ExpiryWarningDays 30
    $status = if (-not $readiness.Ready) { 'FAIL' } elseif ($readiness.ExpiresSoon) { 'WARN' } else { 'PASS' }
    return [pscustomobject]@{ Status = $status; Detail = ('{0} (tested as {1})' -f $readiness.Detail, [Environment]::UserName) }
}

function Test-AcceptanceSharedRuntime {
    $statuses = [System.Collections.Generic.List[string]]::new()
    $parts = [System.Collections.Generic.List[string]]::new()

    $checks = @(foreach ($item in $script:DataRoots.GetEnumerator()) { [pscustomobject]@{ Name = $item.Key; Path = $item.Value; Access = 'ReadWrite'; PathType = 'Container' } })
    foreach ($probe in @(Test-SmartM365OrchestratorPathAccess -Checks $checks -TimeoutSeconds 30)) {
        if (-not $probe.Ready) { $statuses.Add('FAIL') }
        $parts.Add(('{0}: {1}' -f $probe.Name, $(if ($probe.Ready) { 'read/write OK' } else { $probe.Detail })))
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
    }
    catch { $statuses.Add('FAIL'); $parts.Add(('DATA witness: {0}' -f $_.Exception.Message)) }
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
    }
    catch { $statuses.Add('FAIL'); $parts.Add(('code share read: {0}' -f $_.Exception.Message)) }

    $space = Get-SmartM365OrchestratorVolumeFreeSpace -Path $script:WorkspaceRootPath -TimeoutSeconds 30
    if (-not $space.Ready) { $statuses.Add('WARN'); $parts.Add(('DATA free space not measured: {0}' -f $space.Error)) }
    else {
        $statuses.Add($(if ([double]$space.FreePercent -lt 10) { 'FAIL' } elseif ([double]$space.FreePercent -lt 15) { 'WARN' } else { 'PASS' }))
        $parts.Add(('DATA free {0} % ({1} GiB of {2} GiB)' -f $space.FreePercent, $space.FreeGB, $space.TotalGB))
    }
    foreach ($drive in @(@([IO.Path]::GetPathRoot($env:SystemRoot), [IO.Path]::GetPathRoot([IO.Path]::GetTempPath())) | Sort-Object -Unique)) {
        $info = [IO.DriveInfo]::new($drive)
        $freeGB = [math]::Round($info.AvailableFreeSpace / 1GB, 1)
        $statuses.Add($(if ($freeGB -lt 2) { 'FAIL' } elseif ($freeGB -lt 10) { 'WARN' } else { 'PASS' }))
        $parts.Add(('local {0} free {1} GiB' -f $drive.TrimEnd('\'), $freeGB))
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
    return [pscustomobject]@{ Status = Get-SmartM365AcceptanceWorstStatus -Status @($assessment.Status, $windowsStatus); Detail = ('{0} | {1}' -f $assessment.Detail, $windows.Detail) }
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

function Get-AcceptanceAdaxesConfig {
    # Read-only: the collector configuration is read where the collector reads it, never migrated here.
    $path = ConvertTo-SmartM365InventoryConfigLocation -Path (Join-Path $script:SmartInventoryRoot 'ActiveDirectoryInventory\SmartM365-ActiveDirectory-Adaxes-Inventory.local.json')
    $readPath = Get-SmartM365JsonReadPath -Path $path -Optional
    if (-not $readPath) { return $null }
    return (Read-SmartM365JsonDocument -Path $path).Document
}

function Test-AcceptanceAdaxes {
    $config = $null
    $configError = ''
    try { $config = Get-AcceptanceAdaxesConfig } catch { $configError = $_.Exception.Message }
    $environment = $script:ChildEnvironment.Clone()
    $services = @(); $domains = @(); $filter = ''
    if ($null -ne $config) {
        if ($config.PSObject.Properties['AdaxesService']) { $services = ConvertTo-SmartM365AcceptanceStringList $config.AdaxesService }
        if ($config.PSObject.Properties['Domains']) { $domains = ConvertTo-SmartM365AcceptanceStringList $config.Domains }
        if ($config.PSObject.Properties['LdapFilter']) { $filter = [string]$config.LdapFilter }
    }
    $environment['SMART_ADAXES_SERVICES'] = ConvertTo-Json -InputObject @($services) -Compress
    $environment['SMART_ADAXES_DOMAINS'] = ConvertTo-Json -InputObject @($domains) -Compress
    $environment['SMART_ADAXES_LDAP_FILTER'] = $filter
    $child = Invoke-AcceptanceChildProcess -EnginePath $script:PowerShellPath -ScriptText $script:AdaxesChildScript -Environment $environment -TimeoutSeconds $CapabilityTimeoutSeconds
    $assessment = Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize $SampleSize -Requested:($script:Requested -contains 'Adaxes')
    if ($configError) { $assessment.Detail = 'Adaxes collector configuration unreadable: ' + $configError + ' | ' + $assessment.Detail }
    elseif ($null -eq $config -and $assessment.Status -ne 'SKIPPED') { $assessment.Detail = 'Adaxes collector configuration not found | ' + $assessment.Detail }
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

#region Main

$script:SmartInventoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd('\')
$script:SmartM365Root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\')
. (Join-Path $script:SmartM365Root 'Config\SmartM365-TenantContext.ps1')
$script:TenantConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
Import-Module (Join-Path $script:SmartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1') -ErrorAction Stop
Import-Module (Join-Path $script:SmartInventoryRoot 'Orchestrator\SmartM365.Orchestrator.Distributed.psm1') -ErrorAction Stop

$scriptName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$workFolder = Join-Path ([IO.Path]::GetTempPath()) ('{0}_{1}_{2}' -f $scriptName, (Get-Date -Format 'yyyyMMdd_HHmmss'), [guid]::NewGuid().ToString('N').Substring(0, 8))
$null = InitializeScriptEnvironment -OutputPathInit $workFolder -LogFileName $scriptName -CallerScriptPath $PSCommandPath
# Acceptance only: the log stays on the server share, no Teams notification is sent from log lines.
$global:EnableSharePointUpload = $false
& (Get-Module -Name SmartM365.Core) { $script:SmartM365TeamsNotificationInProgress = $true }

$results = [System.Collections.Generic.List[object]]::new()
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
    Write-AcceptanceLog ('Orchestrator server acceptance on {0}: tenant {1}, capabilities {2}, sample {3} (small {4}). Rows stay in memory; nothing is written to the tenant.' -f $env:COMPUTERNAME, $Tenant, ($selected -join ', '), $SampleSize, $SmallSampleSize)

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
        $result = New-AcceptanceResult -Name $name -Status ([string]$outcome.Status) -Detail ([string]$outcome.Detail) -DurationSeconds $seconds
        $results.Add($result)
        $level = switch ($result.Status) { 'PASS' { 'SUCCESS' } 'FAIL' { 'ERROR' } 'WARN' { 'WARNING' } default { 'INFO' } }
        $graphNote = if ($script:GraphChildSeconds -ne $graphBefore) { ' (includes the shared Graph sign-in and reads)' } else { '' }
        Write-AcceptanceLog ('{0}: {1} in {2} s{3}: {4}' -f $name, $result.Status, $result.DurationSeconds, $graphNote, $result.Detail) $level
    }

    if ($Capability -contains 'All') {
        $ready = @($results | Where-Object { $_.Capability -in $script:OrchestratorCapabilities -and $_.Status -in @('PASS', 'WARN') } | ForEach-Object Capability)
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
    }

    Write-AcceptanceLog 'Summary:'
    Write-AcceptanceLog ('  {0,-8} {1,-17} {2,8}  {3}' -f 'Status', 'Capability', 'Seconds', 'Detail')
    foreach ($result in $results) {
        $detail = if ($result.Detail.Length -gt 160) { $result.Detail.Substring(0, 157) + '...' } else { $result.Detail }
        Write-AcceptanceLog ('  {0,-8} {1,-17} {2,8}  {3}' -f $result.Status, $result.Capability, $result.DurationSeconds, $detail)
    }
    $overall = Get-SmartM365AcceptanceWorstStatus -Status @($results | Where-Object Status -ne 'SKIPPED' | ForEach-Object Status)
    if ($overall -eq 'FAIL') { $exitCode = 1 }
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
    $global:EnableSharePointUpload = $false
    $completion = if ($exitCode -ne 0) { 'Failed' } elseif (@($results | Where-Object Status -eq 'WARN').Count -gt 0) { 'CompletedWithWarnings' } else { 'Success' }
    Complete-SmartM365ExecutionContext -Status $completion
}
exit $exitCode

#endregion
