#Requires -Version 7.0
<#
.SYNOPSIS
Collects Microsoft Teams Phone PSTN and Direct Routing usage through Microsoft Graph.

.DESCRIPTION
Uses the Microsoft Graph v1.0 getPstnCalls and getDirectRoutingCalls functions with
app-only certificate authentication. Detailed phone numbers are masked. The primary
SmartFinOps output aggregates usage by user without calculating license cost or
performing any financial allocation.

.VERSION
1.2

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Microsoft Graph application permission: CallRecords.Read.All.
    Delegated authentication is not supported by the Graph call-record APIs.
    Optional phone assignments: MicrosoftTeams 4.7.1+; Organization.Read.All;
    a supported Teams Entra role such as Teams Telephony Administrator.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'SmartM365.Core uses global execution context variables for logs, CSV publication, notifications, and retention.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Final command-line status is intentional.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'Cleanup and optional metadata probes intentionally ignore non-actionable failures.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'New-TeamsPhoneDateWindow creates in-memory date window objects only.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Collection-oriented helper names are intentionally explicit.')]
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [ValidateRange(1, 730)]
    [int]$LookbackDays = 90,
    [Nullable[datetimeoffset]]$FromDate,
    [Nullable[datetimeoffset]]$ToDate,
    [switch]$IncludePstn,
    [switch]$IncludeDirectRouting,
    [switch]$IncludePhoneAssignments,
    [switch]$ValidateOnly,
    [ValidateRange(0, 2147483647)]
    [int]$MaxItems = 0,
    [string]$OutputPath,
    [string]$LatestCsvFolderPath
)

if (-not $PSBoundParameters.ContainsKey('IncludePstn')) { $IncludePstn = $true }
if (-not $PSBoundParameters.ContainsKey('IncludeDirectRouting')) { $IncludeDirectRouting = $true }

if ($MaxItems -gt 0) {
    $global:SmartM365MaxItems = $MaxItems
    $global:SmartM365TestMaxItems = $MaxItems
    $global:SmartM365IsMaxItemsRun = $true
}

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$ScriptVersion = '1.2'
$ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$TaskName = "$ScriptBaseName v$ScriptVersion"
$RunId = [guid]::NewGuid().Guid
$RunStarted = Get-Date
$CurrentOperation = 'Initialize'
$script:TeamsPowerShellConnected = $false

$tenantContextPath = & {
    $directory = $PSScriptRoot
    while ($directory) {
        foreach ($candidate in @(
            (Join-Path -Path $directory -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $directory -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $directory -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $directory) { break }
        $directory = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}

. $tenantContextPath
$TenantContext = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$tenantContextDirectory = Split-Path -Path $tenantContextPath -Parent
$SmartM365Root = if ((Split-Path -Path $tenantContextDirectory -Leaf) -ieq 'Config') {
    Split-Path -Path $tenantContextDirectory -Parent
}
else {
    $tenantContextDirectory
}

$coreModulePath = Join-Path -Path $SmartM365Root -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $coreModulePath -MinimumVersion '1.0.41' -Force -ErrorAction Stop

$LocalConfigPath = Join-Path -Path $PSScriptRoot -ChildPath "$ScriptBaseName.local.json"
$LocalTemplatePath = "$LocalConfigPath.template"
if (-not (Test-Path -LiteralPath $LocalConfigPath)) {
    Initialize-SmartM365LocalJsonFromTemplate `
        -Path $LocalConfigPath `
        -TemplatePath $LocalTemplatePath `
        -ConfigDescription 'Teams Phone PSTN usage local configuration' | Out-Null
}
$ScriptConfig = Get-Content -LiteralPath $LocalConfigPath -Raw | ConvertFrom-Json

function Resolve-TeamsPhoneConfigToken {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($Value -isnot [string]) { return $Value }
    $resolved = $Value
    for ($iteration = 0; $iteration -lt 10; $iteration++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $property = $TenantContext.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$property.Value)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-TeamsPhoneConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$DefaultValue
    )

    $localProperty = $ScriptConfig.PSObject.Properties[$Name]
    if ($null -ne $localProperty -and $null -ne $localProperty.Value) {
        $localValue = $localProperty.Value
        if ($localValue -isnot [string] -or (
            -not [string]::IsNullOrWhiteSpace($localValue) -and
            $localValue.Trim() -notin @('__USE_GLOBAL__', 'USE_GLOBAL')
        )) {
            return Resolve-TeamsPhoneConfigToken -Value $localValue
        }
    }

    $tenantProperty = $TenantContext.PSObject.Properties[$Name]
    if ($null -ne $tenantProperty -and $null -ne $tenantProperty.Value) {
        return Resolve-TeamsPhoneConfigToken -Value $tenantProperty.Value
    }
    return Resolve-TeamsPhoneConfigToken -Value $DefaultValue
}

function ConvertTo-TeamsPhoneUtcText {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    try {
        return ([datetimeoffset]$Value).ToUniversalTime().ToString(
            'yyyy-MM-ddTHH:mm:ss.fffZ',
            [Globalization.CultureInfo]::InvariantCulture
        )
    }
    catch {
        return ''
    }
}

function ConvertTo-TeamsPhoneInvariantNumber {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Value,
        [string]$Format = '0.00'
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    try {
        $number = [Convert]::ToDouble($Value, [Globalization.CultureInfo]::InvariantCulture)
        return $number.ToString($Format, [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        return [string]$Value
    }
}

function ConvertTo-TeamsPhoneDouble {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return 0.0 }
    try { return [Convert]::ToDouble($Value, [Globalization.CultureInfo]::InvariantCulture) }
    catch {
        $parsed = 0.0
        if ([double]::TryParse(
            [string]$Value,
            [Globalization.NumberStyles]::Any,
            [Globalization.CultureInfo]::CurrentCulture,
            [ref]$parsed
        )) {
            return $parsed
        }
        return 0.0
    }
}

function Get-TeamsPhonePropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Object,
        [Parameter(Mandatory)][string[]]$Names
    )

    if ($null -eq $Object) { return $null }
    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }
    return $null
}

function Join-TeamsPhoneDistinctValues {
    [CmdletBinding()]
    param([AllowNull()][object[]]$Values)

    return (@(
        $Values |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    ) -join ';')
}

function Protect-TeamsPhoneNumber {
    [CmdletBinding()]
    param([AllowNull()][object]$Number)

    $text = ([string]$Number).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $digitPositions = New-Object System.Collections.Generic.List[int]
    for ($index = 0; $index -lt $text.Length; $index++) {
        if ([char]::IsDigit($text[$index])) { [void]$digitPositions.Add($index) }
    }
    if ($digitPositions.Count -eq 0) { return '[REDACTED]' }

    $visibleDigits = [Math]::Min(4, $digitPositions.Count)
    $maskUntil = $digitPositions.Count - $visibleDigits
    $characters = $text.ToCharArray()
    for ($digitIndex = 0; $digitIndex -lt $maskUntil; $digitIndex++) {
        $characters[$digitPositions[$digitIndex]] = '*'
    }
    return -join $characters
}

function Get-TeamsPhoneDirection {
    [CmdletBinding()]
    param([AllowNull()][object]$CallType)

    $normalized = ([string]$CallType).Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) { return 'Unknown' }
    if ($normalized -match '(?i)(^|_)(in|inbound)(_|$)' -or $normalized -match '(?i)(Byot|DirectRouting|Dr)In') {
        return 'Inbound'
    }
    if ($normalized -match '(?i)(^|_)(out|outbound)(_|$)' -or $normalized -match '(?i)(Byot|DirectRouting|Dr)Out') {
        return 'Outbound'
    }
    return 'Unknown'
}

function Resolve-TeamsPhoneDateRange {
    [CmdletBinding()]
    param(
        [Nullable[datetimeoffset]]$RequestedFrom,
        [Nullable[datetimeoffset]]$RequestedTo,
        [Parameter(Mandatory)][int]$RequestedLookbackDays
    )

    $nowUtc = [datetimeoffset]::UtcNow
    $toUtc = if ($null -ne $RequestedTo) { $RequestedTo.ToUniversalTime() } else { $nowUtc }
    $fromUtc = if ($null -ne $RequestedFrom) {
        $RequestedFrom.ToUniversalTime()
    }
    else {
        $toUtc.AddDays(-$RequestedLookbackDays)
    }

    if ($null -ne $RequestedFrom -and $null -eq $RequestedTo) {
        $toUtc = $fromUtc.AddDays($RequestedLookbackDays)
        if ($toUtc -gt $nowUtc) { $toUtc = $nowUtc }
    }
    if ($toUtc -gt $nowUtc.AddMinutes(1)) {
        throw "ToDate must not be in the future. Requested UTC value: $(ConvertTo-TeamsPhoneUtcText $toUtc)"
    }
    if ($fromUtc -ge $toUtc) {
        throw 'FromDate must be earlier than ToDate after conversion to UTC.'
    }

    return [pscustomobject]@{
        FromUtc = $fromUtc
        ToUtc = $toUtc
        TotalDays = [math]::Round(($toUtc - $fromUtc).TotalDays, 3)
    }
}

function New-TeamsPhoneDateWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][datetimeoffset]$FromUtc,
        [Parameter(Mandatory)][datetimeoffset]$ToUtc,
        [ValidateRange(1, 90)][int]$MaximumDays = 90
    )

    $cursor = $FromUtc.ToUniversalTime()
    $end = $ToUtc.ToUniversalTime()
    while ($cursor -lt $end) {
        $windowTo = $cursor.AddDays($MaximumDays)
        if ($windowTo -gt $end) { $windowTo = $end }
        [pscustomobject]@{
            FromUtc = $cursor
            ToUtc = $windowTo
            FromText = ConvertTo-TeamsPhoneUtcText $cursor
            ToText = ConvertTo-TeamsPhoneUtcText $windowTo
        }
        $cursor = $windowTo
    }
}

function Get-TeamsPhoneRetryAfterSeconds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [int]$DefaultSeconds = 10,
        [int]$MaximumSeconds = 300
    )

    $value = $null
    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.Headers) {
            $values = $ErrorRecord.Exception.Response.Headers.GetValues('Retry-After')
            $value = @($values | Select-Object -First 1)[0]
        }
    }
    catch {}
    if ($null -eq $value) {
        try { $value = $ErrorRecord.Exception.Data['Retry-After'] } catch {}
    }

    $seconds = 0
    if ($null -ne $value -and [int]::TryParse([string]$value, [ref]$seconds) -and $seconds -gt 0) {
        return [Math]::Min($seconds, $MaximumSeconds)
    }
    $retryDate = [datetimeoffset]::MinValue
    if ($null -ne $value -and [datetimeoffset]::TryParse([string]$value, [ref]$retryDate)) {
        $seconds = [int][Math]::Ceiling(($retryDate.ToUniversalTime() - [datetimeoffset]::UtcNow).TotalSeconds)
        if ($seconds -gt 0) { return [Math]::Min($seconds, $MaximumSeconds) }
    }
    return [Math]::Min([Math]::Max(1, $DefaultSeconds), $MaximumSeconds)
}

function Invoke-TeamsPhoneGraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Operation,
        [ValidateRange(1, 10)][int]$MaximumAttempts = 6
    )

    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            $statusCode = $null
            try {
                if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }
            }
            catch {}
            $isTransient = $statusCode -in @(408, 409, 429, 500, 502, 503, 504) -or
                $_.Exception.Message -match '(?i)throttl|TooManyRequests|temporarily unavailable|timeout|timed out'
            if (-not $isTransient -or $attempt -ge $MaximumAttempts) { throw }

            $fallbackDelay = [Math]::Min(300, [Math]::Pow(2, $attempt) * 5)
            $delay = Get-TeamsPhoneRetryAfterSeconds `
                -ErrorRecord $_ `
                -DefaultSeconds ([int]$fallbackDelay) `
                -MaximumSeconds 300
            WriteLog -Message (
                '{0} throttled or transiently unavailable. Status={1}; attempt {2}/{3}; retry in {4}s.' -f
                $Operation,
                $(if ($null -ne $statusCode) { $statusCode } else { 'unknown' }),
                $attempt,
                $MaximumAttempts,
                $delay
            ) -Level 'WARNING'
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-SmartM365GraphCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InitialUri,
        [Parameter(Mandatory)][string]$Operation,
        [ValidateRange(0, 2147483647)][int]$MaximumItems = 0,
        [scriptblock]$RequestInvoker
    )

    $items = New-Object System.Collections.Generic.List[object]
    $nextLink = $InitialUri
    $pageCount = 0
    while (-not [string]::IsNullOrWhiteSpace($nextLink)) {
        $pageCount++
        $response = if ($null -ne $RequestInvoker) {
            & $RequestInvoker $nextLink
        }
        else {
            Invoke-TeamsPhoneGraphRequest -Uri $nextLink -Operation "$Operation page $pageCount"
        }
        foreach ($item in @($response.value)) {
            [void]$items.Add($item)
            if ($MaximumItems -gt 0 -and $items.Count -ge $MaximumItems) { break }
        }
        if ($MaximumItems -gt 0 -and $items.Count -ge $MaximumItems) { break }
        $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
        $nextLink = if ($null -ne $nextLinkProperty) { [string]$nextLinkProperty.Value } else { '' }
    }

    return [pscustomobject]@{
        Items = $items.ToArray()
        PageCount = $pageCount
    }
}

function Test-TeamsPhoneGraphAuthorizationFailure {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    $statusCode = $null
    try {
        if ($ErrorRecord.Exception.Response) { $statusCode = [int]$ErrorRecord.Exception.Response.StatusCode }
    }
    catch {}
    return $statusCode -in @(401, 403) -or
        $ErrorRecord.Exception.Message -match '(?i)Authorization_RequestDenied|Forbidden|Unauthorized|CallRecords.Read.All'
}

function Get-TeamsPhoneCallRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Pstn', 'DirectRouting')][string]$Source,
        [Parameter(Mandatory)][object[]]$Windows,
        [ValidateRange(0, 2147483647)][int]$MaximumItems = 0
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $pageCount = 0
    foreach ($window in $Windows) {
        if ($MaximumItems -gt 0 -and $rows.Count -ge $MaximumItems) { break }
        $functionName = if ($Source -eq 'Pstn') { 'getPstnCalls' } else { 'getDirectRoutingCalls' }
        $uri = 'https://graph.microsoft.com/v1.0/communications/callRecords/{0}(fromDateTime={1},toDateTime={2})' -f `
            $functionName,
            $window.FromText,
            $window.ToText
        $remaining = if ($MaximumItems -gt 0) { $MaximumItems - $rows.Count } else { 0 }
        try {
            $result = Get-SmartM365GraphCollection `
                -InitialUri $uri `
                -Operation "$Source calls $($window.FromText) to $($window.ToText)" `
                -MaximumItems $remaining
        }
        catch {
            if (Test-TeamsPhoneGraphAuthorizationFailure -ErrorRecord $_) {
                throw (
                    'Microsoft Graph application permission CallRecords.Read.All is missing, not admin-consented, or the connection is not app-only. ' +
                    'The getPstnCalls and getDirectRoutingCalls APIs do not support delegated authentication. ' +
                    "Original error: $($_.Exception.Message)"
                )
            }
            throw
        }
        $pageCount += $result.PageCount
        foreach ($row in @($result.Items)) {
            $rowId = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('id'))
            if ([string]::IsNullOrWhiteSpace($rowId)) {
                $rowId = '{0}|{1}|{2}|{3}' -f
                    (Get-TeamsPhonePropertyValue -Object $row -Names @('callId', 'correlationId')),
                    (Get-TeamsPhonePropertyValue -Object $row -Names @('userId')),
                    (Get-TeamsPhonePropertyValue -Object $row -Names @('startDateTime')),
                    (Get-TeamsPhonePropertyValue -Object $row -Names @('duration'))
            }
            if ($seen.Add($rowId)) { [void]$rows.Add($row) }
            if ($MaximumItems -gt 0 -and $rows.Count -ge $MaximumItems) { break }
        }
    }
    WriteLog -Message ("{0} collection completed. Rows={1}; Pages={2}; Windows={3}" -f $Source, $rows.Count, $pageCount, $Windows.Count) -Level 'INFO'
    return $rows.ToArray()
}

function ConvertTo-TeamsPhonePstnDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$RangeFromUtc,
        [Parameter(Mandatory)][string]$RangeToUtc
    )

    foreach ($row in $Rows) {
        $duration = ConvertTo-TeamsPhoneDouble (Get-TeamsPhonePropertyValue -Object $row -Names @('duration'))
        [pscustomobject][ordered]@{
            RunId = $RunId
            RangeFromUtc = $RangeFromUtc
            RangeToUtc = $RangeToUtc
            Id = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('id'))
            CallId = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('callId'))
            UserId = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('userId'))
            UserPrincipalName = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('userPrincipalName'))
            UserDisplayName = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('userDisplayName'))
            StartDateTimeUtc = ConvertTo-TeamsPhoneUtcText (Get-TeamsPhonePropertyValue -Object $row -Names @('startDateTime'))
            EndDateTimeUtc = ConvertTo-TeamsPhoneUtcText (Get-TeamsPhonePropertyValue -Object $row -Names @('endDateTime'))
            DurationSeconds = ConvertTo-TeamsPhoneInvariantNumber $duration '0'
            Minutes = ConvertTo-TeamsPhoneInvariantNumber ($duration / 60.0)
            Direction = Get-TeamsPhoneDirection (Get-TeamsPhonePropertyValue -Object $row -Names @('callType'))
            CallType = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('callType'))
            Charge = ConvertTo-TeamsPhoneInvariantNumber (Get-TeamsPhonePropertyValue -Object $row -Names @('charge'))
            ConnectionCharge = ConvertTo-TeamsPhoneInvariantNumber (Get-TeamsPhonePropertyValue -Object $row -Names @('connectionCharge'))
            Currency = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('currency'))
            DestinationContext = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('destinationContext'))
            DestinationName = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('destinationName'))
            LicenseCapability = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('licenseCapability'))
            InventoryType = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('inventoryType'))
            Operator = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('operator'))
            CallDurationSource = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('callDurationSource'))
            UsageCountryCode = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('usageCountryCode'))
            TenantCountryCode = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('tenantCountryCode'))
            CallerNumberMasked = Protect-TeamsPhoneNumber (Get-TeamsPhonePropertyValue -Object $row -Names @('callerNumber'))
            CalleeNumberMasked = Protect-TeamsPhoneNumber (Get-TeamsPhonePropertyValue -Object $row -Names @('calleeNumber'))
            ConferenceId = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('conferenceId'))
        }
    }
}

function ConvertTo-TeamsPhoneDirectRoutingDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$RangeFromUtc,
        [Parameter(Mandatory)][string]$RangeToUtc
    )

    foreach ($row in $Rows) {
        $duration = ConvertTo-TeamsPhoneDouble (Get-TeamsPhonePropertyValue -Object $row -Names @('duration'))
        [pscustomobject][ordered]@{
            RunId = $RunId
            RangeFromUtc = $RangeFromUtc
            RangeToUtc = $RangeToUtc
            Id = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('id'))
            CorrelationId = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('correlationId'))
            UserId = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('userId'))
            UserPrincipalName = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('userPrincipalName'))
            UserDisplayName = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('userDisplayName'))
            InviteDateTimeUtc = ConvertTo-TeamsPhoneUtcText (Get-TeamsPhonePropertyValue -Object $row -Names @('inviteDateTime'))
            StartDateTimeUtc = ConvertTo-TeamsPhoneUtcText (Get-TeamsPhonePropertyValue -Object $row -Names @('startDateTime'))
            FailureDateTimeUtc = ConvertTo-TeamsPhoneUtcText (Get-TeamsPhonePropertyValue -Object $row -Names @('failureDateTime'))
            EndDateTimeUtc = ConvertTo-TeamsPhoneUtcText (Get-TeamsPhonePropertyValue -Object $row -Names @('endDateTime'))
            DurationSeconds = ConvertTo-TeamsPhoneInvariantNumber $duration '0'
            Minutes = ConvertTo-TeamsPhoneInvariantNumber ($duration / 60.0)
            Direction = Get-TeamsPhoneDirection (Get-TeamsPhonePropertyValue -Object $row -Names @('callType'))
            CallType = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('callType'))
            SuccessfulCall = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('successfulCall'))
            CallerNumberMasked = Protect-TeamsPhoneNumber (Get-TeamsPhonePropertyValue -Object $row -Names @('callerNumber'))
            CalleeNumberMasked = Protect-TeamsPhoneNumber (Get-TeamsPhonePropertyValue -Object $row -Names @('calleeNumber'))
            MediaPathLocation = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('mediaPathLocation'))
            SignalingLocation = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('signalingLocation'))
            FinalSipCode = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('finalSipCode'))
            CallEndSubReason = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('callEndSubReason'))
            FinalSipCodePhrase = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('finalSipCodePhrase'))
            TrunkFullyQualifiedDomainName = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('trunkFullyQualifiedDomainName'))
            MediaBypassEnabled = [string](Get-TeamsPhonePropertyValue -Object $row -Names @('mediaBypassEnabled'))
        }
    }
}

function ConvertTo-TeamsPhoneUserUsage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$PstnRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DirectRoutingRows
    )

    $events = New-Object System.Collections.Generic.List[object]
    foreach ($row in $PstnRows) {
        $upn = ([string]$row.UserPrincipalName).Trim()
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        [void]$events.Add([pscustomobject]@{
            UserPrincipalName = $upn
            StartDateTime = [string]$row.StartDateTimeUtc
            DurationSeconds = ConvertTo-TeamsPhoneDouble $row.DurationSeconds
            Direction = [string]$row.Direction
            DestinationContext = [string]$row.DestinationContext
            Charge = ConvertTo-TeamsPhoneDouble $row.Charge
            Currency = [string]$row.Currency
            LicenseCapability = [string]$row.LicenseCapability
            Operator = [string]$row.Operator
            Source = 'Pstn'
        })
    }
    foreach ($row in $DirectRoutingRows) {
        $upn = ([string]$row.UserPrincipalName).Trim()
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }
        [void]$events.Add([pscustomobject]@{
            UserPrincipalName = $upn
            StartDateTime = [string]$row.StartDateTimeUtc
            DurationSeconds = ConvertTo-TeamsPhoneDouble $row.DurationSeconds
            Direction = [string]$row.Direction
            DestinationContext = ''
            Charge = 0.0
            Currency = ''
            LicenseCapability = ''
            Operator = ''
            Source = 'DirectRouting'
        })
    }

    foreach ($group in @($events.ToArray() | Group-Object { $_.UserPrincipalName.ToLowerInvariant() })) {
        $groupRows = @($group.Group)
        $pstnGroupRows = @($groupRows | Where-Object Source -eq 'Pstn')
        $directGroupRows = @($groupRows | Where-Object Source -eq 'DirectRouting')
        $datedRows = @($groupRows | ForEach-Object {
            $parsed = [datetimeoffset]::MinValue
            if ([datetimeoffset]::TryParse(
                $_.StartDateTime,
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::AssumeUniversal,
                [ref]$parsed
            )) {
                [pscustomobject]@{ Row = $_; Date = $parsed.ToUniversalTime() }
            }
        })
        $lastDate = @($datedRows | Sort-Object Date -Descending | Select-Object -First 1)
        $activeDays = @($datedRows | ForEach-Object { $_.Date.ToString('yyyy-MM-dd') } | Sort-Object -Unique).Count
        $totalSeconds = ($groupRows | Measure-Object -Property DurationSeconds -Sum).Sum
        $domesticSeconds = ($pstnGroupRows | Where-Object { $_.DestinationContext -match '(?i)^domestic$' } | Measure-Object -Property DurationSeconds -Sum).Sum
        $internationalSeconds = ($pstnGroupRows | Where-Object { $_.DestinationContext -match '(?i)^international$' } | Measure-Object -Property DurationSeconds -Sum).Sum
        $charge = ($pstnGroupRows | Measure-Object -Property Charge -Sum).Sum

        [pscustomobject][ordered]@{
            UserPrincipalName = [string]$groupRows[0].UserPrincipalName
            LastCallDate = if ($lastDate.Count -gt 0) { $lastDate[0].Date.ToString('yyyy-MM-dd') } else { '' }
            ActiveDays = $activeDays
            InboundCallCount = @($groupRows | Where-Object Direction -eq 'Inbound').Count
            OutboundCallCount = @($groupRows | Where-Object Direction -eq 'Outbound').Count
            TotalCallCount = $groupRows.Count
            TotalMinutes = ConvertTo-TeamsPhoneInvariantNumber ($totalSeconds / 60.0)
            DomesticMinutes = if ($pstnGroupRows.Count -gt 0) { ConvertTo-TeamsPhoneInvariantNumber ($domesticSeconds / 60.0) } else { '' }
            InternationalMinutes = if ($pstnGroupRows.Count -gt 0) { ConvertTo-TeamsPhoneInvariantNumber ($internationalSeconds / 60.0) } else { '' }
            Charge = if ($pstnGroupRows.Count -gt 0) { ConvertTo-TeamsPhoneInvariantNumber $charge } else { '' }
            Currency = Join-TeamsPhoneDistinctValues $pstnGroupRows.Currency
            LicenseCapability = Join-TeamsPhoneDistinctValues $pstnGroupRows.LicenseCapability
            Operator = Join-TeamsPhoneDistinctValues $pstnGroupRows.Operator
            HasPstnUsage = ($pstnGroupRows.Count -gt 0)
            HasDirectRoutingUsage = ($directGroupRows.Count -gt 0)
        }
    }
}

function Connect-TeamsPhoneAssignmentSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ApplicationId,
        [Parameter(Mandatory)][string]$DirectoryTenantId,
        [Parameter(Mandatory)][string]$CertificateThumbprint
    )

    $module = Get-Module -ListAvailable -Name MicrosoftTeams |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if ($null -eq $module) {
        throw (
            'Phone assignment inventory requires the officially supported MicrosoftTeams PowerShell module. ' +
            'Install MicrosoftTeams 4.7.1 or later, then rerun with -IncludePhoneAssignments.'
        )
    }
    if ($module.Version -lt [version]'4.7.1') {
        throw "Phone assignment app-only authentication requires MicrosoftTeams 4.7.1 or later. Installed version: $($module.Version)."
    }
    Import-Module -Name $module.Path -Force -ErrorAction Stop
    if (-not (Get-Command -Name Get-CsPhoneNumberAssignment -ErrorAction SilentlyContinue)) {
        throw 'Get-CsPhoneNumberAssignment is not available in the installed MicrosoftTeams module.'
    }

    try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null } catch {}
    try {
        Connect-MicrosoftTeams `
            -ApplicationId $ApplicationId `
            -TenantId $DirectoryTenantId `
            -CertificateThumbprint $CertificateThumbprint `
            -ErrorAction Stop | Out-Null
        $script:TeamsPowerShellConnected = $true
    }
    catch {
        throw (
            'MicrosoftTeams app-only connection for phone assignments failed. This optional inventory requires ' +
            'Microsoft Graph application permission Organization.Read.All and a supported Teams Entra role on the service principal ' +
            '(least voice-focused option: Teams Telephony Administrator). ' +
            "Original error: $($_.Exception.Message)"
        )
    }
}

function Get-TeamsPhoneAssignments {
    [CmdletBinding()]
    param([ValidateRange(0, 2147483647)][int]$MaximumItems = 0)

    $rows = New-Object System.Collections.Generic.List[object]
    $skip = 0
    $pageSize = 1000
    while ($true) {
        $requestedTop = if ($MaximumItems -gt 0) {
            [Math]::Min($pageSize, $MaximumItems - $rows.Count)
        }
        else {
            $pageSize
        }
        if ($requestedTop -le 0) { break }
        try {
            $page = @(Get-CsPhoneNumberAssignment -Skip $skip -Top $requestedTop -ErrorAction Stop)
        }
        catch {
            throw (
                'Get-CsPhoneNumberAssignment failed. Verify MicrosoftTeams app-only RBAC, Organization.Read.All, ' +
                "and the service principal Teams role. Original error: $($_.Exception.Message)"
            )
        }
        foreach ($number in $page) {
            [void]$rows.Add([pscustomobject][ordered]@{
                RunId = $RunId
                TelephoneNumberMasked = Protect-TeamsPhoneNumber (Get-TeamsPhonePropertyValue -Object $number -Names @('TelephoneNumber'))
                OperatorId = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('OperatorId'))
                NumberType = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('NumberType'))
                ActivationState = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('ActivationState'))
                AssignedPstnTargetId = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('AssignedPstnTargetId', 'TargetId'))
                TargetType = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('TargetType'))
                AssignmentCategory = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('AssignmentCategory'))
                Capability = Join-TeamsPhoneDistinctValues (Get-TeamsPhonePropertyValue -Object $number -Names @('Capability', 'AcquiredCapabilities'))
                PstnAssignmentStatus = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('PstnAssignmentStatus', 'AssignmentStatus'))
                PstnPartnerName = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('PstnPartnerName', 'PartnerName'))
                NumberSource = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('NumberSource'))
                IsoCountryCode = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('IsoCountryCode'))
                City = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('City', 'PlaceName'))
                LocationId = [string](Get-TeamsPhonePropertyValue -Object $number -Names @('LocationId'))
            })
        }
        if ($page.Count -lt $requestedTop) { break }
        $skip += $page.Count
        if ($MaximumItems -gt 0 -and $rows.Count -ge $MaximumItems) { break }
    }
    WriteLog -Message ("Phone assignment collection completed. Rows={0}" -f $rows.Count) -Level 'INFO'
    return $rows.ToArray()
}

$PstnColumns = @(
    'RunId', 'RangeFromUtc', 'RangeToUtc', 'Id', 'CallId', 'UserId', 'UserPrincipalName', 'UserDisplayName',
    'StartDateTimeUtc', 'EndDateTimeUtc', 'DurationSeconds', 'Minutes', 'Direction', 'CallType', 'Charge',
    'ConnectionCharge', 'Currency', 'DestinationContext', 'DestinationName', 'LicenseCapability', 'InventoryType',
    'Operator', 'CallDurationSource', 'UsageCountryCode', 'TenantCountryCode', 'CallerNumberMasked',
    'CalleeNumberMasked', 'ConferenceId'
)
$DirectRoutingColumns = @(
    'RunId', 'RangeFromUtc', 'RangeToUtc', 'Id', 'CorrelationId', 'UserId', 'UserPrincipalName',
    'UserDisplayName', 'InviteDateTimeUtc', 'StartDateTimeUtc', 'FailureDateTimeUtc', 'EndDateTimeUtc',
    'DurationSeconds', 'Minutes', 'Direction', 'CallType', 'SuccessfulCall', 'CallerNumberMasked',
    'CalleeNumberMasked', 'MediaPathLocation', 'SignalingLocation', 'FinalSipCode', 'CallEndSubReason',
    'FinalSipCodePhrase', 'TrunkFullyQualifiedDomainName', 'MediaBypassEnabled'
)
$UserUsageColumns = @(
    'UserPrincipalName', 'LastCallDate', 'ActiveDays', 'InboundCallCount', 'OutboundCallCount', 'TotalCallCount',
    'TotalMinutes', 'DomesticMinutes', 'InternationalMinutes', 'Charge', 'Currency', 'LicenseCapability',
    'Operator', 'HasPstnUsage', 'HasDirectRoutingUsage'
)
$PhoneAssignmentColumns = @(
    'RunId', 'TelephoneNumberMasked', 'OperatorId', 'NumberType', 'ActivationState', 'AssignedPstnTargetId',
    'TargetType', 'AssignmentCategory', 'Capability', 'PstnAssignmentStatus', 'PstnPartnerName', 'NumberSource',
    'IsoCountryCode', 'City', 'LocationId'
)

$dateRange = Resolve-TeamsPhoneDateRange `
    -RequestedFrom $FromDate `
    -RequestedTo $ToDate `
    -RequestedLookbackDays $LookbackDays
$windows = @(New-TeamsPhoneDateWindow -FromUtc $dateRange.FromUtc -ToUtc $dateRange.ToUtc -MaximumDays 90)
$rangeFromText = ConvertTo-TeamsPhoneUtcText $dateRange.FromUtc
$rangeToText = ConvertTo-TeamsPhoneUtcText $dateRange.ToUtc

$dataAllRoot = [string](Get-TeamsPhoneConfigValue -Name 'DataAllRootPath' -DefaultValue '')
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $LatestCsvFolderPath = [string](Get-TeamsPhoneConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $defaultOutputPath = if (-not [string]::IsNullOrWhiteSpace($dataAllRoot)) {
        Join-Path -Path $dataAllRoot -ChildPath 'M365\Teams\PhonePstnUsage'
    }
    else {
        ''
    }
    $OutputPath = [string](Get-TeamsPhoneConfigValue `
        -Name 'TeamsPhonePstnUsageCsvLogFolderPath' `
        -DefaultValue $defaultOutputPath)
}
if ([string]::IsNullOrWhiteSpace($OutputPath) -or [string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    throw 'OutputPath and LatestCsvFolderPath must resolve to non-empty tenant-scoped paths.'
}

$global:RetentionMaxCSV = [int](Get-TeamsPhoneConfigValue -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:EnableTeamsNotifications = [bool](Get-TeamsPhoneConfigValue -Name 'EnableTeamsNotifications' -DefaultValue $false)
$global:RetentionMaxLogs = [int](Get-TeamsPhoneConfigValue -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-TeamsPhoneConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = [string](Get-TeamsPhoneConfigValue -Name 'SharePointSiteHostname' -DefaultValue '')
$global:SharePointSitePath = [string](Get-TeamsPhoneConfigValue -Name 'SharePointSitePath' -DefaultValue '')
$global:SharePointLibraryDisplayName = [string](Get-TeamsPhoneConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents')
$global:SharePointTargetFolderPath = [string](Get-TeamsPhoneConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue '')
$global:EnableWeeklyHistory = [bool](Get-TeamsPhoneConfigValue -Name 'EnableWeeklyHistory' -DefaultValue $true)
$global:WeeklyHistoryFolderPath = [string](Get-TeamsPhoneConfigValue `
    -Name 'WeeklyHistoryFolderPath' `
    -DefaultValue (Join-Path -Path $OutputPath -ChildPath 'WeeklyHistory'))
$global:WeeklyHistoryRetentionWeeks = [int](Get-TeamsPhoneConfigValue -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)
if ($ValidateOnly) {
    $global:EnableSharePointUpload = $false
    $global:EnableTeamsNotifications = $false
}

$global:SmartM365ExecutionStartTime = $RunStarted
$global:SmartM365ExecutionSummaryWritten = $false
$global:SmartM365ScriptName = $ScriptBaseName

$AppId = [string](Get-TeamsPhoneConfigValue -Name 'AppId' -DefaultValue '')
$TenantId = [string](Get-TeamsPhoneConfigValue -Name 'TenantId' -DefaultValue '')
$OrgDomain = [string](Get-TeamsPhoneConfigValue -Name 'OrgDomain' -DefaultValue '')
$Thumbprint = [string](Get-TeamsPhoneConfigValue -Name 'Thumbprint' -DefaultValue '')
if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
    $Thumbprint = [string](Get-TeamsPhoneConfigValue -Name 'Thumb' -DefaultValue '')
}

try {
    $CurrentOperation = 'Initialize script environment'
    $OutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $ScriptBaseName
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    WriteLog -Message (
        'Starting {0}. Tenant={1}; FromUtc={2}; ToUtc={3}; Windows={4}; IncludePstn={5}; IncludeDirectRouting={6}; IncludePhoneAssignments={7}; MaxItems={8}' -f
        $TaskName,
        $Tenant,
        $rangeFromText,
        $rangeToText,
        $windows.Count,
        [bool]$IncludePstn,
        [bool]$IncludeDirectRouting,
        [bool]$IncludePhoneAssignments,
        $MaxItems
    ) -Level 'INFO'

    Add-SmartM365CsvValidationRule `
        -Rules $global:SmartM365CsvValidationRules `
        -BaseFileName 'M365_Teams_PSTNCalls' `
        -CriticalFields @('Id') `
        -RequiredColumns $PstnColumns `
        -AllowEmptyDataset
    Add-SmartM365CsvValidationRule `
        -Rules $global:SmartM365CsvValidationRules `
        -BaseFileName 'M365_Teams_DirectRoutingCalls' `
        -CriticalFields @('Id') `
        -RequiredColumns $DirectRoutingColumns `
        -AllowEmptyDataset
    Add-SmartM365CsvValidationRule `
        -Rules $global:SmartM365CsvValidationRules `
        -BaseFileName 'M365_Teams_PhoneUserUsage' `
        -CriticalFields @('UserPrincipalName') `
        -RequiredColumns $UserUsageColumns `
        -AllowEmptyDataset
    Add-SmartM365CsvValidationRule `
        -Rules $global:SmartM365CsvValidationRules `
        -BaseFileName 'M365_Teams_PhoneAssignments' `
        -CriticalFields @('TelephoneNumberMasked', 'NumberType') `
        -RequiredColumns $PhoneAssignmentColumns `
        -AllowEmptyDataset

    if ($ValidateOnly) {
        $CurrentOperation = 'Validate configuration and schemas'
        Assert-SmartM365CsvDataCompleteness -Data @() -Columns $PstnColumns -BaseFileName 'M365_Teams_PSTNCalls'
        Assert-SmartM365CsvDataCompleteness -Data @() -Columns $DirectRoutingColumns -BaseFileName 'M365_Teams_DirectRoutingCalls'
        Assert-SmartM365CsvDataCompleteness -Data @() -Columns $UserUsageColumns -BaseFileName 'M365_Teams_PhoneUserUsage'
        Assert-SmartM365CsvDataCompleteness -Data @() -Columns $PhoneAssignmentColumns -BaseFileName 'M365_Teams_PhoneAssignments'
        $validationSummary = (
            'ValidateOnly passed. Tenant={0}; UTC range={1} to {2}; windows={3}; schemas=4; no cloud connection and no CSV publication.' -f
            $Tenant,
            $rangeFromText,
            $rangeToText,
            $windows.Count
        )
        WriteLog -Message $validationSummary -Level 'SUCCESS'
        try { Stop-Transcript | Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } catch {}
        Complete-SmartM365ExecutionContext -Status Success
        return
    }

    if (-not $IncludePstn -and -not $IncludeDirectRouting -and -not $IncludePhoneAssignments) {
        throw 'At least one collection switch must be enabled.'
    }

    $CurrentOperation = 'Connect Microsoft Graph app-only'
    Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -VerboseDisconnect:$true
    $connection = Connect-SmartM365CloudSession `
        -ExchangeOnline:$false `
        -Graph:$true `
        -AppId $AppId `
        -Thumbprint $Thumbprint `
        -TenantId $TenantId `
        -Organization $OrgDomain `
        -GraphScopes @('CallRecords.Read.All')
    if (-not $connection.GraphConnected) {
        throw 'Microsoft Graph app-only connection failed. Check AppId, TenantId, certificate thumbprint, and local tenant configuration.'
    }
    $graphContext = Get-MgContext -ErrorAction Stop
    if ([string]$graphContext.AuthType -notmatch '(?i)AppOnly') {
        throw (
            "Microsoft Graph authentication type '$($graphContext.AuthType)' is not app-only. " +
            'The PSTN and Direct Routing call-record APIs do not support delegated authentication.'
        )
    }

    $CurrentOperation = 'Run preflight'
    Invoke-SmartM365Preflight `
        -ScriptName $TaskName `
        -RequiredModules @('Microsoft.Graph.Authentication') `
        -RequiredCommands @('Invoke-MgGraphRequest') `
        -OutputPaths @($OutputPath, $LatestCsvFolderPath) | Out-Null
    WriteLog -Message 'Required Microsoft Graph application permission: CallRecords.Read.All. Live authorization is verified by the selected call-record endpoint before CSV publication.' -Level 'INFO'

    $rawPstnRows = @()
    $rawDirectRoutingRows = @()
    if ($IncludePstn) {
        $CurrentOperation = 'Collect PSTN and Operator Connect calls'
        $rawPstnRows = @(Get-TeamsPhoneCallRows -Source Pstn -Windows $windows -MaximumItems $MaxItems)
    }
    if ($IncludeDirectRouting) {
        $CurrentOperation = 'Collect Direct Routing calls'
        $rawDirectRoutingRows = @(Get-TeamsPhoneCallRows -Source DirectRouting -Windows $windows -MaximumItems $MaxItems)
    }

    $CurrentOperation = 'Normalize call rows'
    $pstnRows = @(ConvertTo-TeamsPhonePstnDetail -Rows $rawPstnRows -RangeFromUtc $rangeFromText -RangeToUtc $rangeToText)
    $directRoutingRows = @(ConvertTo-TeamsPhoneDirectRoutingDetail -Rows $rawDirectRoutingRows -RangeFromUtc $rangeFromText -RangeToUtc $rangeToText)
    $userUsageRows = @(ConvertTo-TeamsPhoneUserUsage -PstnRows $pstnRows -DirectRoutingRows $directRoutingRows)

    $assignmentRows = @()
    if ($IncludePhoneAssignments) {
        $CurrentOperation = 'Connect MicrosoftTeams app-only for phone assignments'
        Connect-TeamsPhoneAssignmentSession `
            -ApplicationId $AppId `
            -DirectoryTenantId $TenantId `
            -CertificateThumbprint $Thumbprint
        $CurrentOperation = 'Collect phone assignments'
        $assignmentRows = @(Get-TeamsPhoneAssignments -MaximumItems $MaxItems)
    }

    $CurrentOperation = 'Publish CSV files'
    $null = Export-SmartM365Csv `
        -BaseFileName 'M365_Teams_PSTNCalls' `
        -OutputPath $OutputPath `
        -GlobalPath $LatestCsvFolderPath `
        -Data $pstnRows `
        -Columns $PstnColumns
    $null = Export-SmartM365Csv `
        -BaseFileName 'M365_Teams_DirectRoutingCalls' `
        -OutputPath $OutputPath `
        -GlobalPath $LatestCsvFolderPath `
        -Data $directRoutingRows `
        -Columns $DirectRoutingColumns
    $usageResult = Export-SmartM365Csv `
        -BaseFileName 'M365_Teams_PhoneUserUsage' `
        -OutputPath $OutputPath `
        -GlobalPath $LatestCsvFolderPath `
        -Data $userUsageRows `
        -Columns $UserUsageColumns
    if ($IncludePhoneAssignments) {
        $null = Export-SmartM365Csv `
            -BaseFileName 'M365_Teams_PhoneAssignments' `
            -OutputPath $OutputPath `
            -GlobalPath $LatestCsvFolderPath `
            -Data $assignmentRows `
            -Columns $PhoneAssignmentColumns
    }

    $summary = (
        'PSTN calls={0}; Direct Routing calls={1}; Users={2}; Phone assignments={3}; UTC range={4} to {5}' -f
        $pstnRows.Count,
        $directRoutingRows.Count,
        $userUsageRows.Count,
        $assignmentRows.Count,
        $rangeFromText,
        $rangeToText
    )
    WriteLog -Message ("Teams Phone PSTN usage inventory completed. {0}" -f $summary) -Level 'SUCCESS'
    Send-SmartM365TeamsNotification `
        -Level SUCCESS `
        -Channel Infos `
        -Title 'SmartM365 Teams Phone PSTN usage inventory completed' `
        -Message $summary `
        -ResultSummary $summary `
        -Facts @{
            Tenant = $Tenant
            RangeFromUtc = $rangeFromText
            RangeToUtc = $rangeToText
            PstnRows = $pstnRows.Count
            DirectRoutingRows = $directRoutingRows.Count
            UserUsageRows = $userUsageRows.Count
            PhoneAssignmentRows = $assignmentRows.Count
            PrimaryCsv = [string]$usageResult.LatestPath
            RunId = $RunId
        } | Out-Null

    try { Stop-Transcript | Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    Write-Host ("Teams Phone PSTN usage inventory completed. {0}" -f $summary)
}
catch {
    $errorRecord = $_
    $message = $errorRecord.Exception.Message
    try {
        WriteLog -Message ("Teams Phone PSTN usage inventory failed during {0}: {1}" -f $CurrentOperation, $message) -Level 'ERROR'
    }
    catch {}
    try {
        Send-SmartM365TeamsNotification `
            -Level ERROR `
            -Channel Alerts `
            -Title 'SmartM365 Teams Phone PSTN usage inventory failed' `
            -Message $message `
            -Facts @{
                Tenant = $Tenant
                FailedOperation = $CurrentOperation
                RangeFromUtc = $rangeFromText
                RangeToUtc = $rangeToText
                Computer = $env:COMPUTERNAME
                LogPath = [string]$global:LogTextFile
                TranscriptPath = [string]$global:logTranscriptFile
                OutputPath = $OutputPath
                RunId = $RunId
            } | Out-Null
    }
    catch {}
    try { Stop-Transcript | Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } catch {}
    try {
        Complete-SmartM365ExecutionContext `
            -Status Failed `
            -ErrorRecord $errorRecord `
            -FailureStage $CurrentOperation
    }
    catch {}
    throw
}
finally {
    if ($script:TeamsPowerShellConnected) {
        try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null } catch {}
        $script:TeamsPowerShellConnected = $false
    }
    try { Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true } catch {}
    try { Stop-Transcript | Out-Null } catch {}
}
