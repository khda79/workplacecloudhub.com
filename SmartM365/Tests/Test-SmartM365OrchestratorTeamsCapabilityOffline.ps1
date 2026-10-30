#Requires -Version 7.0
<#
.SYNOPSIS
Checks the TeamsPowerShell read-only capability probe, offline.
.DESCRIPTION
Verifies that the read-only TeamsPowerShell probe uses the app-only call of
M365-TeamsPhonePstnUsage-Inventory (Get-CsPhoneNumberAssignment -Top 1) instead of Get-CsTenant,
which the job does not use, and runs the probe body against stubbed MicrosoftTeams commands:
ready, module missing (ComponentMissing), command missing and query failure. The stub exposes
Get-CsPhoneNumberAssignment only after Connect-MicrosoftTeams, the stricter case.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Distributed.psm1') -Force
$script:Checks = 0
function Assert-Teams {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$distributedText = [IO.File]::ReadAllText((Join-Path $folder 'SmartM365.Orchestrator.Distributed.psm1'))
$readOnly = [regex]::Match($distributedText, "(?s)\`$teamsScript = @'(?<Body>.*?)'@\s*\r?\n\s*(?<Call>\`$results\.Add\(\(Invoke-IsolatedCapabilityProbe -Name 'TeamsPowerShell'[^\r\n]*)")
Assert-Teams $readOnly.Success 'Read-only Teams probe exists.'
Assert-Teams ($readOnly.Groups['Call'].Value -match '-EnginePath \$pwshPath' -and $readOnly.Groups['Call'].Value -match '-TimeoutSeconds \$ProbeTimeoutSeconds') 'Probe runs in bounded PowerShell 7.'
$body = $readOnly.Groups['Body'].Value
Assert-Teams ($body -notmatch '(?m)^[^#\r\n]*\bGet-CsTenant\b') 'Probe no longer calls Get-CsTenant.'
Assert-Teams ($body -match 'Get-CsPhoneNumberAssignment -Top 1 -ErrorAction Stop') 'Probe reads one phone number assignment.'
Assert-Teams ($body -match 'Connect-MicrosoftTeams -TenantId \$env:SMART_TENANT_ID -ApplicationId \$env:SMART_APP_ID -CertificateThumbprint \$env:SMART_CERT_THUMBPRINT') 'Probe connects app-only with the certificate.'
Assert-Teams ($body -notmatch '(?i)\b(Set|New|Remove|Grant|Update|Enable|Disable|Register|Unregister)-Cs') 'Probe never changes Teams.'
Assert-Teams ($body.IndexOf('Connect-MicrosoftTeams -TenantId') -ge 0 -and $body.IndexOf('Connect-MicrosoftTeams -TenantId') -lt $body.IndexOf('Get-Command -Name Get-CsPhoneNumberAssignment')) 'Probe connects before looking for Get-CsPhoneNumberAssignment.'
$tokens = $null; $errors = $null
[void][Management.Automation.Language.Parser]::ParseInput($body, [ref]$tokens, [ref]$errors)
Assert-Teams ($errors.Count -eq 0) 'Probe script parses.'

# Probe body against stubs: functions shadow the MicrosoftTeams commands in a child scope.
function Invoke-TeamsProbeBody {
    param([bool]$ModuleInstalled = $true, [bool]$CommandAvailable = $true, [string]$QueryError = '')
    $state = @{ Top = $null; Connected = $false; Disconnected = $false }
    $output = & {
        function Get-Module { param([switch]$ListAvailable, [string]$Name) if ($ModuleInstalled) { [pscustomobject]@{ Name = 'MicrosoftTeams'; Version = [version]'7.8.0' } } }
        function Import-Module { param([string]$Name, [string]$MinimumVersion) }
        function Get-Command { param([string]$Name) if ($Name -eq 'Get-CsPhoneNumberAssignment' -and -not ($CommandAvailable -and $state.Connected)) { return $null } [pscustomobject]@{ Name = $Name } }
        function Connect-MicrosoftTeams { param($TenantId, $ApplicationId, $CertificateThumbprint) $state.Connected = $true; 'connected' }
        function Disconnect-MicrosoftTeams { $state.Disconnected = $true }
        function Get-CsPhoneNumberAssignment { param([int]$Top) $state.Top = $Top; if ($QueryError) { throw $QueryError }; [pscustomobject]@{ TelephoneNumber = '+10000000000' } }
        . ([scriptblock]::Create($body))
    }
    $line = @($output | Where-Object { [string]$_ -like 'SMARTM365_CAPABILITY_JSON:*' })
    if ($line.Count -ne 1) { throw "FAIL: probe output: $($output -join ' | ')" }
    $result = ([string]$line[0]).Substring('SMARTM365_CAPABILITY_JSON:'.Length) | ConvertFrom-Json
    return [pscustomobject]@{ Ready = [bool]$result.Ready; Detail = [string]$result.Detail; Top = $state.Top; Connected = $state.Connected; Disconnected = $state.Disconnected }
}

$ready = Invoke-TeamsProbeBody
Assert-Teams ($ready.Ready -and $ready.Top -eq 1 -and $ready.Connected -and $ready.Disconnected -and $ready.Detail -like '*Get-CsPhoneNumberAssignment*') "Ready probe: $($ready.Detail)"
$missing = Invoke-TeamsProbeBody -ModuleInstalled $false
Assert-Teams (-not $missing.Ready -and $missing.Detail -eq 'MicrosoftTeams PowerShell module is not installed on this server.' -and -not $missing.Connected) 'Missing module keeps its message.'
Assert-Teams (Test-SmartM365OrchestratorCapabilityComponentMissing -Result ([pscustomobject]@{ Ready = $false; Detail = $missing.Detail })) 'Missing module is a missing component.'
$noCommand = Invoke-TeamsProbeBody -CommandAvailable $false
Assert-Teams (-not $noCommand.Ready -and $noCommand.Detail -like '*Get-CsPhoneNumberAssignment is not available*after Connect-MicrosoftTeams*' -and $noCommand.Connected -and $noCommand.Disconnected) "Module without the command: $($noCommand.Detail)"
Assert-Teams (-not (Test-SmartM365OrchestratorCapabilityComponentMissing -Result ([pscustomobject]@{ Ready = $false; Detail = $noCommand.Detail }))) 'Module without the command stays a warning.'
$denied = Invoke-TeamsProbeBody -QueryError 'Forbidden: the service principal has no Teams role.'
Assert-Teams (-not $denied.Ready -and $denied.Detail -like 'Forbidden*' -and $denied.Disconnected) "Query failure: $($denied.Detail)"

"[{0}] PASS: TeamsPowerShell read-only probe, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks
