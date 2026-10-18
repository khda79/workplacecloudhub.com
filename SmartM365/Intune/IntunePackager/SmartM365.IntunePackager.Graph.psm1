#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Graph publication engine for Smart IntunePackager (win32LobApp, Graph beta).

.DESCRIPTION
    Depends only on Microsoft.Graph.Authentication (Connect-MgGraph, Invoke-MgGraphRequest).
    Every Graph call is a REST request. The publication flow is split into one function per
    step so the engine can later be shared by other SmartM365 publishers:

    CreateApp -> ContentVersion -> ContentFile -> SasWait -> Upload -> Commit -> CommitWait -> Finalize

    Graph throttling (429/503/504) honors Retry-After. Azure Storage uploads use Put Block /
    Put Block List with retries and SAS renewal (before expiry and after a 403).

    Assignments are intentionally out of scope: administrators assign the app in the Intune
    admin center. The Assignments region below is reserved for a future shared engine.

.VERSION
    1.0.11

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Net.Http
Import-Module (Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.IntunePackager.Core.psm1') -Scope Local

$script:IpGraphScopes = @('DeviceManagementApps.ReadWrite.All')
# Set when Connect-IpGraph changed the persisted DisableLoginByWAM option; restored by Disconnect-IpGraph.
$script:IpRestoreLoginByWam = $false
$script:IpGraphSettings = @{
    BaseUri            = 'https://graph.microsoft.com/beta'
    MaxRetries         = 5
    PollSeconds        = 5
    PollTimeoutMinutes = 30
}
$script:IpRetryableGraphStatus = @(429, 503, 504)
$script:IpRetryableAzureStatus = @(408, 429, 500, 502, 503, 504)
$script:IpAzureStorageVersion = '2020-10-02'
$script:IpHttpStatusNames = @{
    BadRequest = 400; Unauthorized = 401; Forbidden = 403; NotFound = 404; Conflict = 409
    TooManyRequests = 429; InternalServerError = 500; BadGateway = 502; ServiceUnavailable = 503; GatewayTimeout = 504
}

#region Settings and connection

function Get-IpGraphScopes {
    [CmdletBinding()]
    param()
    return @($script:IpGraphScopes)
}

function Set-IpGraphSettings {
    [CmdletBinding()]
    param(
        [string]$BaseUri,
        [ValidateRange(0, 20)][int]$MaxRetries = -1,
        [ValidateRange(1, 300)][int]$PollSeconds = 0,
        [ValidateRange(1, 720)][int]$PollTimeoutMinutes = 0
    )
    if ($BaseUri) { $script:IpGraphSettings.BaseUri = $BaseUri.TrimEnd('/') }
    if ($PSBoundParameters.ContainsKey('MaxRetries')) { $script:IpGraphSettings.MaxRetries = $MaxRetries }
    if ($PollSeconds -gt 0) { $script:IpGraphSettings.PollSeconds = $PollSeconds }
    if ($PollTimeoutMinutes -gt 0) { $script:IpGraphSettings.PollTimeoutMinutes = $PollTimeoutMinutes }
}

function Get-IpGraphSettings {
    [CmdletBinding()]
    param()
    return [pscustomobject]$script:IpGraphSettings.Clone()
}

function Assert-IpGraphModule {
    [CmdletBinding()]
    param()
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
        throw 'Required module Microsoft.Graph.Authentication is not installed. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

function Get-IpGraphContext {
    [CmdletBinding()]
    param()
    try { return (Get-MgContext -ErrorAction SilentlyContinue) } catch { return $null }
}

function Test-IpGraphBrokerLoaded {
    <#
    .SYNOPSIS
        Returns $true when the native WAM broker runtime (msalruntime*.dll) is loaded in this process.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    try { return [bool](@((Get-Process -Id $PID).Modules | Where-Object { $_.ModuleName -like 'msalruntime*.dll' }).Count) }
    catch { return $false }
}

function Get-IpGraphRuntimeInfo {
    [CmdletBinding()]
    param()
    $module = Get-Module -Name 'Microsoft.Graph.Authentication' | Sort-Object Version -Descending | Select-Object -First 1
    return [pscustomobject]@{
        PowerShell         = $PSVersionTable.PSVersion.ToString()
        DotNet             = [Environment]::Version.ToString()
        GraphAuthentication = if ($module) { $module.Version.ToString() } else { '' }
        WamBrokerLoaded    = Test-IpGraphBrokerLoaded
    }
}

function Get-IpProfileValue {
    param([psobject]$Profile, [string]$Name, [string]$Default = '')
    $property = $Profile.PSObject.Properties[$Name]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) { return $Default }
    return [string]$property.Value
}

function Resolve-IpSignInMode {
    <#
    .SYNOPSIS
        Returns the effective sign-in mode: Certificate, DeviceCode, Browser or Wam.

    .DESCRIPTION
        Auto selects Browser (system browser, no WAM broker) when the tenant profile has its own
        ClientId, otherwise Wam. With the default Microsoft Graph PowerShell client the SDK always
        uses WAM, whatever Set-MgGraphOption -DisableLoginByWAM says.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][psobject]$TenantProfile, [switch]$UseDeviceCode)

    if ((Get-IpProfileValue $TenantProfile 'AuthMode') -eq 'Certificate') { return 'Certificate' }
    if ($UseDeviceCode) { return 'DeviceCode' }
    $mode = Get-IpProfileValue $TenantProfile 'DelegatedSignIn' 'Auto'
    if ($mode -eq 'Auto') { $mode = if (Get-IpProfileValue $TenantProfile 'ClientId') { 'Browser' } else { 'Wam' } }
    return $mode
}

function Disconnect-IpGraph {
    [CmdletBinding()]
    param()
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    catch { Write-Verbose 'No Microsoft Graph session required disconnection.' }
    if ($script:IpRestoreLoginByWam) {
        # Restore the persisted SDK option changed for a Browser sign-in.
        try { Set-MgGraphOption -DisableLoginByWAM $false }
        catch { Write-IpLog ("Unable to restore the Microsoft Graph DisableLoginByWAM option: {0}" -f $_.Exception.Message) -Level WARN }
        $script:IpRestoreLoginByWam = $false
    }
}

function Connect-IpGraph {
    <#
    .SYNOPSIS
        Disconnects any existing Graph session, then connects with the tenant profile
        (delegated interactive or application certificate).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$TenantProfile,
        [switch]$UseDeviceCode,
        # Delegated scopes; defaults to the publication scope. Read-only tools pass DeviceManagementApps.Read.All.
        [string[]]$Scopes = @()
    )

    if ($Scopes.Count -eq 0) { $Scopes = Get-IpGraphScopes }
    Assert-IpGraphModule
    Disconnect-IpGraph

    $tenantId = Get-IpProfileValue $TenantProfile 'TenantId'
    $clientId = Get-IpProfileValue $TenantProfile 'ClientId'
    $signIn = Resolve-IpSignInMode -TenantProfile $TenantProfile -UseDeviceCode:$UseDeviceCode
    $parameters = @{ NoWelcome = $true; ContextScope = 'Process'; ErrorAction = 'Stop' }
    if ($tenantId) { $parameters.TenantId = $tenantId }
    else { Write-IpLog ("Tenant profile {0} has no TenantId: the signed-in account selects the tenant." -f $TenantProfile.Name) -Level WARN }

    switch ($signIn) {
        'Certificate' {
            $parameters.ClientId = $clientId
            $parameters.CertificateThumbprint = Get-IpProfileValue $TenantProfile 'CertificateThumbprint'
        }
        default {
            $parameters.Scopes = $Scopes
            if ($clientId) { $parameters.ClientId = $clientId }
            if ($signIn -eq 'DeviceCode') { $parameters.UseDeviceCode = $true }
        }
    }
    if ($signIn -eq 'Browser') {
        $module = Get-Module -Name 'Microsoft.Graph.Authentication' | Sort-Object Version -Descending | Select-Object -First 1
        # DisableLoginByWAM with a custom ClientId is broken before 2.35.1 (msgraph-sdk-powershell issue 3518).
        if ($module -and $module.Version -lt [version]'2.35.1') { throw "DelegatedSignIn=Browser requires Microsoft.Graph.Authentication 2.35.1 or later (loaded: $($module.Version))." }
        if (-not [bool](Get-MgGraphOption).DisableLoginByWAM) {
            Set-MgGraphOption -DisableLoginByWAM $true
            $script:IpRestoreLoginByWam = $true
        }
    }
    if ($signIn -eq 'Wam') {
        Write-IpLog 'Sign-in uses the Web Account Manager broker (default Microsoft Graph PowerShell client). The native broker can crash PowerShell at process exit; use DelegatedSignIn=Browser with your own ClientId, DeviceCode or Certificate to avoid it.' -Level WARN
        # The SDK parents the WAM window to the console window, which can be hidden or behind other windows.
        Write-IpLog "If you don't see the sign-in window, check the taskbar or press Alt+Tab."
    }
    Write-IpLog ("Connecting to Microsoft Graph. Tenant profile={0}; SignIn={1}; Scopes={2}" -f $TenantProfile.Name, $signIn, ($Scopes -join ', '))

    try { Connect-MgGraph @parameters | Out-Null }
    catch {
        if ($script:IpRestoreLoginByWam) { Disconnect-IpGraph }
        # Report the authentication error itself, not the ActionPreferenceStopException wrapper.
        $exception = $_.Exception
        if ($exception -is [Management.Automation.ActionPreferenceStopException] -and $exception.ErrorRecord -and $exception.ErrorRecord.Exception) { $exception = $exception.ErrorRecord.Exception }
        throw [InvalidOperationException]::new(("[Connect] Sign-in to tenant profile '{0}' failed or was canceled: {1}" -f $TenantProfile.Name, $exception.Message), $_.Exception)
    }
    $context = Get-IpGraphContext
    if ($null -eq $context) { throw 'Microsoft Graph connection failed: no context was returned.' }
    $tenantKind = Get-IpProfileValue $TenantProfile 'TenantIdKind'
    if (-not $tenantKind -and $tenantId) { $tenantKind = if (Test-IpGuid $tenantId) { 'Guid' } else { 'Domain' } }
    if ($tenantKind -eq 'Guid' -and [string]$context.TenantId -ne $tenantId) {
        Disconnect-IpGraph
        throw ("Connected tenant {0} does not match the tenant profile {1}." -f $context.TenantId, $TenantProfile.Name)
    }
    if ($tenantKind -eq 'Domain') { Write-IpLog ("Tenant profile {0} uses the domain {1}; connected TenantId={2}." -f $TenantProfile.Name, $tenantId, $context.TenantId) }
    $runtime = Get-IpGraphRuntimeInfo
    Write-IpLog ("Graph runtime: PowerShell={0}; .NET={1}; Microsoft.Graph.Authentication={2}; WamBrokerLoaded={3}" -f $runtime.PowerShell, $runtime.DotNet, $runtime.GraphAuthentication, $runtime.WamBrokerLoaded)

    $grantedScopes = @($context.Scopes)
    $missing = @($Scopes | Where-Object { $_ -notin $grantedScopes -and ($_ -replace '\.Read\.All$', '.ReadWrite.All') -notin $grantedScopes })
    if ($missing.Count -gt 0) {
        Write-IpLog ("Granted permissions do not include: {0}. Graph requests may fail with 403." -f ($missing -join ', ')) -Level WARN
    }

    $identity = if ($context.Account) { [string]$context.Account } else { 'application ' + [string]$context.ClientId }
    Write-IpLog ("Connected to Microsoft Graph. TenantId={0}; Identity={1}; AuthType={2}" -f $context.TenantId, $identity, $context.AuthType)
    return [pscustomobject]@{
        TenantId = [string]$context.TenantId
        Identity = $identity
        AuthType = [string]$context.AuthType
        Scopes   = $grantedScopes
    }
}

#endregion

#region Graph requests

function Start-IpSleep {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][double]$Seconds)
    if ($Seconds -gt 0) { Start-Sleep -Milliseconds ([int]($Seconds * 1000)) }
}

function Invoke-IpGraphRawRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [AllowNull()][string]$Body
    )

    if ($null -ne $Body) {
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
    }
    return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType PSObject -ErrorAction Stop
}

function Get-IpHttpErrorInfo {
    <#
    .SYNOPSIS
        Extracts HTTP status, Retry-After, message and request id from a Graph SDK error.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    $exception = if ($ErrorRecord -is [Management.Automation.ErrorRecord]) { $ErrorRecord.Exception } else { $ErrorRecord }
    $statusCode = 0
    $retryAfter = $null
    $requestId = ''
    $message = [string]$exception.Message

    $current = $exception
    while ($null -ne $current -and $statusCode -eq 0) {
        $responseProperty = $current.PSObject.Properties['Response']
        if ($null -ne $responseProperty -and $null -ne $responseProperty.Value) {
            $response = $responseProperty.Value
            if ($response.PSObject.Properties['StatusCode']) { $statusCode = [int]$response.StatusCode }
            if ($response.PSObject.Properties['Headers'] -and $null -ne $response.Headers) {
                $header = $response.Headers.PSObject.Properties['RetryAfter']
                if ($null -ne $header -and $null -ne $header.Value) {
                    if ($null -ne $header.Value.Delta) { $retryAfter = [int][Math]::Ceiling($header.Value.Delta.TotalSeconds) }
                    elseif ($null -ne $header.Value.Date) { $retryAfter = [int][Math]::Max(0, [Math]::Ceiling(($header.Value.Date - [datetimeoffset]::UtcNow).TotalSeconds)) }
                }
            }
        }
        $current = $current.InnerException
    }

    if ($statusCode -eq 0) {
        $statusLine = [regex]::Match($message, '(?m)^HTTP/\d(\.\d)?\s+(?<Code>[1-5]\d\d)\b')
        $numeric = [regex]::Match($message, '\b(4\d\d|5\d\d)\b')
        $named = [regex]::Match($message, '\b(' + (($script:IpHttpStatusNames.Keys | Sort-Object) -join '|') + ')\b')
        if ($statusLine.Success) { $statusCode = [int]$statusLine.Groups['Code'].Value }
        elseif ($named.Success) { $statusCode = $script:IpHttpStatusNames[$named.Value] }
        elseif ($numeric.Success) { $statusCode = [int]$numeric.Value }
    }

    $details = ''
    if ($ErrorRecord -is [Management.Automation.ErrorRecord] -and $null -ne $ErrorRecord.ErrorDetails) { $details = [string]$ErrorRecord.ErrorDetails.Message }
    # Invoke-MgGraphRequest puts the raw HTTP response (request line, headers, JSON body) in the
    # exception message: keep only the Graph error body for the user-facing message.
    $jsonStart = $message.IndexOf('{"error"')
    if (-not $details -and $jsonStart -ge 0) { $details = $message.Substring($jsonStart).Trim() }
    $headerRequestId = [regex]::Match($message, '(?im)^request-id:\s*(?<Id>[0-9a-f-]{36})\s*$')
    if ($headerRequestId.Success) { $requestId = $headerRequestId.Groups['Id'].Value }
    if ($details) {
        try {
            $parsed = $details | ConvertFrom-Json -ErrorAction Stop
            if ($parsed.PSObject.Properties['error']) {
                if ($parsed.error.PSObject.Properties['message'] -and $parsed.error.message) { $message = [string]$parsed.error.message }
                if ($parsed.error.PSObject.Properties['innerError'] -and $parsed.error.innerError.PSObject.Properties['request-id']) {
                    $requestId = [string]$parsed.error.innerError.'request-id'
                }
            }
        }
        catch { $message = $details }
    }
    elseif ($message -match '\r?\n') {
        $message = ([regex]::Split($message, '\r?\n') | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()
    }

    return [pscustomobject]@{
        StatusCode        = $statusCode
        RetryAfterSeconds = $retryAfter
        Message           = $message
        RequestId         = $requestId
    }
}

function Get-IpRetryDelaySeconds {
    [CmdletBinding()]
    param([int]$Attempt, [AllowNull()]$RetryAfterSeconds)
    if ($null -ne $RetryAfterSeconds -and [int]$RetryAfterSeconds -gt 0) { return [Math]::Min(300, [int]$RetryAfterSeconds) }
    return [Math]::Min(60, [Math]::Pow(2, $Attempt))
}

function Resolve-IpGraphUri {
    param([string]$Uri)
    if ($Uri -match '^(https://|<)') { return $Uri }
    return $script:IpGraphSettings.BaseUri + '/' + $Uri.TrimStart('/')
}

function Invoke-IpGraphRequest {
    <#
    .SYNOPSIS
        Sends a Graph REST request with throttling retries and a step-qualified error message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        $Body,
        [string]$Step = 'Graph',
        [int]$MaxRetries = -1
    )

    if ($MaxRetries -lt 0) { $MaxRetries = [int]$script:IpGraphSettings.MaxRetries }
    $fullUri = Resolve-IpGraphUri -Uri $Uri
    $json = if ($PSBoundParameters.ContainsKey('Body')) { ConvertTo-Json -InputObject $Body -Depth 30 -Compress } else { $null }

    for ($attempt = 1; ; $attempt++) {
        try {
            return (Invoke-IpGraphRawRequest -Method $Method -Uri $fullUri -Body $json)
        }
        catch {
            $info = Get-IpHttpErrorInfo -ErrorRecord $_
            $path = ([Uri]$fullUri).AbsolutePath
            if ($info.StatusCode -in $script:IpRetryableGraphStatus -and $attempt -le $MaxRetries) {
                $delay = Get-IpRetryDelaySeconds -Attempt $attempt -RetryAfterSeconds $info.RetryAfterSeconds
                Write-IpLog ("[{0}] {1} {2} returned HTTP {3}. Retry {4}/{5} in {6} s." -f $Step, $Method, $path, $info.StatusCode, $attempt, $MaxRetries, $delay) -Level WARN
                Start-IpSleep -Seconds $delay
                continue
            }
            $statusText = if ($info.StatusCode) { "HTTP $($info.StatusCode)" } else { 'no HTTP status' }
            $requestText = if ($info.RequestId) { " RequestId=$($info.RequestId)." } else { '' }
            throw [InvalidOperationException]::new(("[{0}] {1} {2} failed ({3}): {4}{5}" -f $Step, $Method, $path, $statusText, $info.Message, $requestText), $_.Exception)
        }
    }
}

function Get-IpGraphCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Uri, [string]$Step = 'Graph')

    $items = [Collections.Generic.List[object]]::new()
    $visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $next = $Uri
    while (-not [string]::IsNullOrWhiteSpace($next)) {
        # @odata.nextLink is absolute and already carries $filter, $select and the skip token.
        if (-not $visited.Add($next)) { throw "[$Step] Graph returned the same @odata.nextLink twice; paging stopped." }
        $page = Invoke-IpGraphRequest -Method GET -Uri $next -Step $Step
        $values = $page.PSObject.Properties['value']
        if ($null -ne $values) { foreach ($item in @($values.Value)) { if ($null -ne $item) { $items.Add($item) } } }
        $next = if ($page.PSObject.Properties['@odata.nextLink']) { [string]$page.'@odata.nextLink' } else { '' }
    }
    return $items.ToArray()
}

#endregion

#region App lookup

function Test-IpWin32AppObject {
    param($App)
    return ($null -ne $App -and $App.PSObject.Properties['@odata.type'] -and [string]$App.'@odata.type' -eq '#microsoft.graph.win32LobApp')
}

function Get-IpWin32App {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$AppId, [string]$Step = 'GetApp')

    if (-not (Test-IpGuid $AppId)) { throw "[$Step] Invalid app id: $AppId" }
    $app = Invoke-IpGraphRequest -Method GET -Uri "deviceAppManagement/mobileApps/$AppId" -Step $Step
    if (-not (Test-IpWin32AppObject $app)) { throw ("[{0}] Intune app {1} is not a Win32 app (type {2})." -f $Step, $AppId, $app.'@odata.type') }
    return $app
}

function Get-IpObjectText {
    param($InputObject, [string]$Name)
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return '' }
    return [string]$property.Value
}

function Find-IpWin32App {
    <#
    .SYNOPSIS
        Lists Win32 apps with base-type properties only, optionally filtered by display name
        (exact or contains, case-insensitive, evaluated locally).

    .DESCRIPTION
        $select is limited to mobileApp properties: on the mobileApps collection Graph validates
        $select against the base type, so win32LobApp properties such as displayVersion are
        rejected (HTTP 400). Use Get-IpWin32AppDetails for one app's Win32 properties. The
        type filter keeps large Win32 objects (detection rules, icons) out of the listing.
    #>
    [CmdletBinding()]
    param([string]$DisplayName, [switch]$Contains)

    $filter = [Uri]::EscapeDataString("isof('microsoft.graph.win32LobApp')")
    $select = [Uri]::EscapeDataString('id,displayName,publisher,lastModifiedDateTime')
    $all = @(Get-IpGraphCollection -Uri ('deviceAppManagement/mobileApps?$filter={0}&$select={1}' -f $filter, $select) -Step 'FindApp')
    $apps = $all
    if ($DisplayName) {
        if ($Contains) { $apps = @($all | Where-Object { (Get-IpObjectText $_ 'displayName') -like ('*' + [WildcardPattern]::Escape($DisplayName) + '*') }) }
        else { $apps = @($all | Where-Object { (Get-IpObjectText $_ 'displayName') -eq $DisplayName }) }
    }
    Write-IpLog ("[FindApp] Graph returned {0} Win32 app(s); {1} match the name filter '{2}'." -f $all.Count, $apps.Count, $DisplayName)
    return @(foreach ($app in $apps) {
            [pscustomobject]@{
                id                   = Get-IpObjectText $app 'id'
                displayName          = Get-IpObjectText $app 'displayName'
                publisher            = Get-IpObjectText $app 'publisher'
                # Raw value (string or DateTime): the GUI converts it to local time.
                lastModifiedDateTime = if ($app.PSObject.Properties['lastModifiedDateTime']) { $app.lastModifiedDateTime } else { $null }
            }
        })
}

function ConvertTo-IpWin32AppSummary {
    <#
    .SYNOPSIS
        Returns the properties shown for a selected Win32 app.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$App)
    return [pscustomobject]@{
        id                      = Get-IpObjectText $App 'id'
        displayName             = Get-IpObjectText $App 'displayName'
        displayVersion          = Get-IpObjectText $App 'displayVersion'
        publisher               = Get-IpObjectText $App 'publisher'
        committedContentVersion = Get-IpObjectText $App 'committedContentVersion'
        setupFilePath           = Get-IpObjectText $App 'setupFilePath'
        publishingState         = Get-IpObjectText $App 'publishingState'
        isAssigned              = Get-IpObjectText $App 'isAssigned'
        lastModifiedDateTime    = if ($App.PSObject.Properties['lastModifiedDateTime']) { $App.lastModifiedDateTime } else { $null }
    }
}

function Get-IpWin32AppDetails {
    <#
    .SYNOPSIS
        Reads one Win32 app by id (GET) and returns the properties shown for the selected app.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$AppId)
    return (ConvertTo-IpWin32AppSummary -App (Get-IpWin32App -AppId $AppId -Step 'AppDetails'))
}

#endregion

#region Publication steps

function New-IpWin32App {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([Parameter(Mandatory = $true)][Collections.IDictionary]$Body)

    if (-not $PSCmdlet.ShouldProcess([string]$Body.displayName, 'Create Intune Win32 app')) { return $null }
    $app = Invoke-IpGraphRequest -Method POST -Uri 'deviceAppManagement/mobileApps' -Body $Body -Step 'CreateApp'
    if ([string]::IsNullOrWhiteSpace([string]$app.id)) { throw '[CreateApp] Graph did not return an app id.' }
    Write-IpLog "[CreateApp] Created Intune Win32 app: $($app.id)"
    return $app
}

function New-IpContentVersion {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([Parameter(Mandatory = $true)][string]$AppId)

    if (-not $PSCmdlet.ShouldProcess($AppId, 'Create content version')) { return $null }
    $version = Invoke-IpGraphRequest -Method POST -Uri "deviceAppManagement/mobileApps/$AppId/microsoft.graph.win32LobApp/contentVersions" -Body @{} -Step 'ContentVersion'
    if ([string]::IsNullOrWhiteSpace([string]$version.id)) { throw '[ContentVersion] Graph did not return a content version id.' }
    Write-IpLog "[ContentVersion] Created content version: $($version.id)"
    return [string]$version.id
}

function New-IpContentFile {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][string]$AppId,
        [Parameter(Mandatory = $true)][string]$ContentVersionId,
        [Parameter(Mandatory = $true)][Collections.IDictionary]$Body
    )

    if (-not $PSCmdlet.ShouldProcess($AppId, 'Create content file')) { return $null }
    $filesUri = "deviceAppManagement/mobileApps/$AppId/microsoft.graph.win32LobApp/contentVersions/$ContentVersionId/files"
    $file = Invoke-IpGraphRequest -Method POST -Uri $filesUri -Body $Body -Step 'ContentFile'
    if ([string]::IsNullOrWhiteSpace([string]$file.id)) { throw '[ContentFile] Graph did not return a content file id.' }
    Write-IpLog "[ContentFile] Created content file: $($file.id)"
    return [pscustomobject]@{ FileId = [string]$file.id; FileUri = ('{0}/{1}' -f $filesUri, $file.id) }
}

function Get-IpGraphObjectDiagnosticText {
    param($Value)
    if ($null -eq $Value) { return '<null>' }
    $parts = foreach ($property in @($Value.PSObject.Properties | Sort-Object Name)) {
        if ($property.Name -match 'azureStorageUri|Uri$|Sas|Secret|Token|Key|manifest') { continue }
        $item = $property.Value
        if ($null -ne $item -and $item -isnot [string] -and $item -isnot [ValueType]) { $item = ConvertTo-Json -InputObject $item -Depth 5 -Compress }
        '{0}={1}' -f $property.Name, $item
    }
    return (@($parts) -join '; ')
}

function Wait-IpContentFileState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$FileUri,
        [Parameter(Mandatory = $true)][string[]]$SuccessStates,
        [Parameter(Mandatory = $true)][string[]]$FailureStates,
        [Parameter(Mandatory = $true)][string]$Step,
        [switch]$RequireAzureStorageUri,
        [scriptblock]$CancelCheck
    )

    $deadline = (Get-Date).AddMinutes([int]$script:IpGraphSettings.PollTimeoutMinutes)
    $last = $null
    $lastState = ''
    do {
        if ($CancelCheck -and (& $CancelCheck)) { throw [OperationCanceledException]::new("[$Step] Publication canceled.") }
        $file = Invoke-IpGraphRequest -Method GET -Uri $FileUri -Step $Step
        $last = $file
        $state = [string]$file.uploadState
        if ($state -in $FailureStates) { throw ("[{0}] Content file entered failure state {1}. Detail: {2}" -f $Step, $state, (Get-IpGraphObjectDiagnosticText $file)) }
        if ($state -in $SuccessStates -and (-not $RequireAzureStorageUri -or -not [string]::IsNullOrWhiteSpace([string]$file.azureStorageUri))) { return $file }
        if ($state -ne $lastState) { Write-IpLog ("[{0}] Waiting for content file state. Current={1}" -f $Step, $state); $lastState = $state }
        Start-IpSleep -Seconds ([int]$script:IpGraphSettings.PollSeconds)
    } while ((Get-Date) -lt $deadline)

    throw ("[{0}] Timed out after {1} minutes waiting for {2}. Last detail: {3}" -f $Step, $script:IpGraphSettings.PollTimeoutMinutes, ($SuccessStates -join ','), (Get-IpGraphObjectDiagnosticText $last))
}

function Request-IpSasRenewal {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$FileUri, [scriptblock]$CancelCheck)

    Write-IpLog '[Upload] Requesting a renewed Azure Storage SAS URI from Intune.'
    Invoke-IpGraphRequest -Method POST -Uri "$FileUri/renewUpload" -Body @{} -Step 'SasRenewal' | Out-Null
    $file = Wait-IpContentFileState -FileUri $FileUri -SuccessStates @('azureStorageUriRenewalSuccess') -FailureStates @('azureStorageUriRenewalFailed', 'azureStorageUriRequestFailed') -Step 'SasRenewal' -RequireAzureStorageUri -CancelCheck $CancelCheck
    Write-IpLog ("[SasRenewal] Renewed SAS URI received (state {0}). Resuming the upload." -f $file.uploadState)
    return [string]$file.azureStorageUri
}

#endregion

#region Azure Storage upload

function Get-IpSasExpiryUtc {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Uri)
    try {
        foreach ($part in ([Uri]$Uri).Query.TrimStart('?').Split('&')) {
            $pair = $part.Split('=', 2)
            if ($pair.Count -eq 2 -and $pair[0] -eq 'se') {
                return ([datetimeoffset]::Parse([Net.WebUtility]::UrlDecode($pair[1]), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal)).UtcDateTime
            }
        }
    }
    catch { Write-Verbose 'SAS expiry could not be parsed.' }
    return [datetime]::MinValue
}

function Test-IpSasRenewalNeeded {
    [CmdletBinding()]
    param([datetime]$ExpiryUtc, [datetime]$NowUtc = [datetime]::UtcNow, [int]$RenewBeforeMinutes = 5)
    if ($ExpiryUtc -eq [datetime]::MinValue) { return $false }
    return ($ExpiryUtc -le $NowUtc.AddMinutes($RenewBeforeMinutes))
}

function Join-IpSasQuery {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Uri, [Parameter(Mandatory = $true)][string]$Query)
    if ($Uri.Contains('?')) { return "$Uri&$Query" }
    return "$Uri`?$Query"
}

function New-IpBlockId {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][int]$Index)
    return [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(('block-{0:D8}' -f $Index)))
}

function New-IpBlockListXml {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string[]]$BlockIds)
    $builder = [Text.StringBuilder]::new('<?xml version="1.0" encoding="utf-8"?><BlockList>')
    foreach ($id in $BlockIds) { [void]$builder.AppendFormat('<Latest>{0}</Latest>', [Security.SecurityElement]::Escape($id)) }
    [void]$builder.Append('</BlockList>')
    return $builder.ToString()
}

function Invoke-IpAzureBlobRequest {
    <#
    .SYNOPSIS
        Sends one PUT to Azure Storage and returns the status code and body.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$Client,
        [Parameter(Mandatory = $true)][string]$Uri,
        [byte[]]$Bytes,
        [string]$Xml
    )

    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Put, [Uri]$Uri)
    $response = $null
    try {
        $request.Headers.Add('x-ms-version', $script:IpAzureStorageVersion)
        if ($PSBoundParameters.ContainsKey('Xml')) { $request.Content = [Net.Http.StringContent]::new($Xml, [Text.Encoding]::UTF8, 'application/xml') }
        else { $request.Content = [Net.Http.ByteArrayContent]::new($Bytes) }
        $response = $Client.SendAsync($request).GetAwaiter().GetResult()
        $body = if ($response.IsSuccessStatusCode) { '' } else { $response.Content.ReadAsStringAsync().GetAwaiter().GetResult() }
        return [pscustomobject]@{ StatusCode = [int]$response.StatusCode; Body = $body }
    }
    finally {
        if ($response) { $response.Dispose() }
        $request.Dispose()
    }
}

function Send-IpAzureBlobWithRetry {
    <#
    .SYNOPSIS
        Sends one block or block list with retries. Returns 'Ok' or 'AuthFailed' (caller renews the SAS).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][Net.Http.HttpClient]$Client,
        [Parameter(Mandatory = $true)][string]$Uri,
        [byte[]]$Bytes,
        [string]$Xml,
        [Parameter(Mandatory = $true)][string]$Operation,
        [ValidateRange(1, 20)][int]$MaxRetries = 5
    )

    $requestArgs = @{ Client = $Client; Uri = $Uri }
    if ($PSBoundParameters.ContainsKey('Xml')) { $requestArgs.Xml = $Xml } else { $requestArgs.Bytes = $Bytes }

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $failure = ''
        try {
            $response = Invoke-IpAzureBlobRequest @requestArgs
            if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) { return 'Ok' }
            if ($response.StatusCode -eq 403) { return 'AuthFailed' }
            if ($response.StatusCode -notin $script:IpRetryableAzureStatus) {
                throw [InvalidOperationException]::new(("[Upload] {0} failed with HTTP {1}: {2}" -f $Operation, $response.StatusCode, $response.Body))
            }
            $failure = "HTTP $($response.StatusCode)"
        }
        catch [InvalidOperationException] { throw }
        catch { $failure = $_.Exception.Message }

        if ($attempt -ge $MaxRetries) { throw ("[Upload] {0} failed after {1} attempts: {2}" -f $Operation, $MaxRetries, $failure) }
        $delay = [Math]::Min(60, 5 * $attempt)
        Write-IpLog ("[Upload] {0} attempt {1}/{2} failed ({3}). Retrying in {4} s." -f $Operation, $attempt, $MaxRetries, $failure, $delay) -Level WARN
        Start-IpSleep -Seconds $delay
    }
}

function Send-IpIntuneWinContent {
    <#
    .SYNOPSIS
        Streams the encrypted payload from the .intunewin archive to Azure Storage in blocks.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$IntuneWinPath,
        [Parameter(Mandatory = $true)][string]$ContentEntryName,
        [Parameter(Mandatory = $true)][string]$AzureStorageUri,
        [Parameter(Mandatory = $true)][string]$FileUri,
        [ValidateRange(1, 100)][int]$BlockSizeMB = 16,
        [ValidateRange(1, 20)][int]$MaxRetries = 5,
        [ValidateRange(1, 60)][int]$RenewBeforeExpiryMinutes = 5,
        [ValidateRange(1, 240)][int]$HttpTimeoutMinutes = 30,
        # Validation option: requests a SAS renewal before this block number, counted from 1 (0 = disabled).
        [ValidateRange(0, 100000)][int]$ForceSasRenewalAtBlock = 0,
        [scriptblock]$ProgressCallback,
        [scriptblock]$CancelCheck
    )

    $state = @{ Uri = $AzureStorageUri; Expiry = (Get-IpSasExpiryUtc -Uri $AzureStorageUri) }
    $renew = {
        $state.Uri = Request-IpSasRenewal -FileUri $FileUri -CancelCheck $CancelCheck
        $state.Expiry = Get-IpSasExpiryUtc -Uri $state.Uri
        if ($state.Expiry -ne [datetime]::MinValue) { Write-IpLog ("[Upload] Renewed SAS expires at {0:u}." -f $state.Expiry) }
    }
    $send = {
        param([string]$Query, [byte[]]$Bytes, [string]$Xml, [string]$Operation)
        if (Test-IpSasRenewalNeeded -ExpiryUtc $state.Expiry -RenewBeforeMinutes $RenewBeforeExpiryMinutes) { & $renew }
        for ($authAttempt = 1; $authAttempt -le 2; $authAttempt++) {
            $sendArgs = @{ Client = $client; Uri = (Join-IpSasQuery -Uri $state.Uri -Query $Query); Operation = $Operation; MaxRetries = $MaxRetries }
            if ($Xml) { $sendArgs.Xml = $Xml } else { $sendArgs.Bytes = $Bytes }
            if ((Send-IpAzureBlobWithRetry @sendArgs) -eq 'Ok') { return }
            if ($authAttempt -eq 2) { throw "[Upload] $Operation was rejected with HTTP 403 after SAS renewal." }
            Write-IpLog "[Upload] $Operation returned HTTP 403 (expired or invalid SAS). Renewing and retrying." -Level WARN
            & $renew
        }
    }

    $blockSize = $BlockSizeMB * 1MB
    $blockIds = [Collections.Generic.List[string]]::new()
    $client = [Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromMinutes($HttpTimeoutMinutes)
    $archive = [IO.Compression.ZipFile]::OpenRead($IntuneWinPath)
    try {
        $entry = $archive.Entries | Where-Object { $_.FullName -eq $ContentEntryName } | Select-Object -First 1
        if (-not $entry) { throw "[Upload] Encrypted payload entry not found: $ContentEntryName" }
        $total = [int64]$entry.Length
        $blockCount = [int][Math]::Max(1, [Math]::Ceiling($total / [double]$blockSize))
        $stream = $entry.Open()
        try {
            $buffer = [byte[]]::new($blockSize)
            $sent = [int64]0
            $index = 0
            while ($true) {
                # Fill the whole buffer: deflate streams may return partial reads.
                $read = 0
                while ($read -lt $blockSize) {
                    $count = $stream.Read($buffer, $read, $blockSize - $read)
                    if ($count -le 0) { break }
                    $read += $count
                }
                if ($read -le 0) { break }
                if ($CancelCheck -and (& $CancelCheck)) { throw [OperationCanceledException]::new('[Upload] Publication canceled.') }

                $payload = $buffer
                if ($read -lt $blockSize) {
                    $payload = [byte[]]::new($read)
                    [Array]::Copy($buffer, $payload, $read)
                }
                $blockId = New-IpBlockId -Index $index
                # Block numbers in messages start at 1; block ids keep the zero-based index.
                $number = $index + 1
                if ($ForceSasRenewalAtBlock -gt 0 -and $number -eq $ForceSasRenewalAtBlock) {
                    $after = if ($number -gt 1) { 'after block {0}/{1}, ' -f ($number - 1), $blockCount } else { '' }
                    Write-IpLog ("[Upload] Validation option: forcing a SAS renewal {0}before block {1}/{2}." -f $after, $number, $blockCount) -Level WARN
                    & $renew
                }
                & $send ('comp=block&blockid=' + [Uri]::EscapeDataString($blockId)) $payload '' ("block {0}/{1}" -f $number, $blockCount)
                $blockIds.Add($blockId)
                $sent += $read
                $index++
                if ($ProgressCallback) { & $ProgressCallback $sent $total }
                if ($read -lt $blockSize) { break }
            }
        }
        finally { $stream.Dispose() }

        if ($blockIds.Count -eq 0) { throw '[Upload] The encrypted payload is empty.' }
        & $send 'comp=blocklist' $null (New-IpBlockListXml -BlockIds $blockIds.ToArray()) 'block list commit'
        Write-IpLog ("[Upload] Uploaded {0} block(s), {1}." -f $blockIds.Count, (Format-IpByteSize $sent))
        return $blockIds.Count
    }
    finally {
        $archive.Dispose()
        $client.Dispose()
    }
}

#endregion

#region Commit and finalize

function Submit-IpContentFileCommit {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][string]$FileUri,
        [Parameter(Mandatory = $true)][Collections.IDictionary]$CommitBody,
        [ValidateRange(0, 5)][int]$CommitRetries = 1,
        [scriptblock]$CancelCheck
    )

    if (-not $PSCmdlet.ShouldProcess($FileUri, 'Commit content file')) { return }
    for ($attempt = 0; ; $attempt++) {
        Invoke-IpGraphRequest -Method POST -Uri "$FileUri/commit" -Body $CommitBody -Step 'Commit' | Out-Null
        Write-IpLog '[Commit] Commit requested.'
        try {
            Wait-IpContentFileState -FileUri $FileUri -SuccessStates @('commitFileSuccess') -FailureStates @('commitFileFailed') -Step 'CommitWait' -CancelCheck $CancelCheck | Out-Null
            Write-IpLog '[CommitWait] Content file committed.'
            return
        }
        catch {
            if ($attempt -ge $CommitRetries -or $_.Exception.Message -notmatch 'commitFileFailed') { throw }
            Write-IpLog ("[CommitWait] Commit failed; retrying commit ({0}/{1})." -f ($attempt + 1), $CommitRetries) -Level WARN
        }
    }
}

function Complete-IpAppContentVersion {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][string]$AppId,
        [Parameter(Mandatory = $true)][Collections.IDictionary]$PatchBody
    )

    if (-not $PSCmdlet.ShouldProcess($AppId, "Set committedContentVersion $($PatchBody.committedContentVersion)")) { return }
    Invoke-IpGraphRequest -Method PATCH -Uri "deviceAppManagement/mobileApps/$AppId" -Body $PatchBody -Step 'Finalize' | Out-Null
    Write-IpLog "[Finalize] App metadata applied and content version $($PatchBody.committedContentVersion) committed."
}

#endregion

#region Assignments (reserved)
# Not implemented in v1: Smart IntunePackager never assigns apps. Administrators create
# assignments in the Intune admin center. Future shared assignment functions belong here.
#endregion

#region Orchestration

function Get-IpPublishPlan {
    <#
    .SYNOPSIS
        Returns the ordered Graph requests a publication would send (payloads redacted for display).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$Definition,
        [Parameter(Mandatory = $true)][Collections.IDictionary]$AppBody,
        [Parameter(Mandatory = $true)][psobject]$IntuneWinMetadata,
        [int]$BlockSizeMB = 16
    )

    $appId = if ($Definition.Publish.Mode -eq 'UpdateExisting') { [string]$Definition.Publish.ExistingAppId } else { '<new-app-id>' }
    $appUri = "deviceAppManagement/mobileApps/$appId"
    $fileUri = "$appUri/microsoft.graph.win32LobApp/contentVersions/<content-version-id>/files/<file-id>"
    $blocks = [int][Math]::Max(1, [Math]::Ceiling([double]$IntuneWinMetadata.SizeEncrypted / ($BlockSizeMB * 1MB)))
    $steps = [Collections.Generic.List[object]]::new()
    $add = { param($step, $method, $uri, $body, $note) $steps.Add([pscustomobject][ordered]@{ Step = $step; Method = $method; Uri = (Resolve-IpGraphUri $uri); Note = $note; Body = (Protect-IpPayloadForDisplay $body) }) }

    if ($Definition.Publish.Mode -eq 'UpdateExisting') { & $add 'GetApp' 'GET' $appUri $null 'Verify the existing Win32 app.' }
    else { & $add 'CreateApp' 'POST' 'deviceAppManagement/mobileApps' $AppBody 'Create the win32LobApp.' }
    & $add 'ContentVersion' 'POST' "$appUri/microsoft.graph.win32LobApp/contentVersions" @{} 'Create a content version.'
    & $add 'ContentFile' 'POST' "$appUri/microsoft.graph.win32LobApp/contentVersions/<content-version-id>/files" (New-IpContentFileBody $IntuneWinMetadata) 'Declare the encrypted content file.'
    & $add 'SasWait' 'GET' $fileUri $null 'Wait for azureStorageUriRequestSuccess.'
    & $add 'Upload' 'PUT' '<azure-storage-sas-uri>' $null ("{0} block(s) of up to {1} MB, then Put Block List; SAS renewed when needed." -f $blocks, $BlockSizeMB)
    & $add 'Commit' 'POST' "$fileUri/commit" (New-IpCommitBody $IntuneWinMetadata) 'Commit with fileEncryptionInfo from Detection.xml.'
    & $add 'CommitWait' 'GET' $fileUri $null 'Wait for commitFileSuccess.'
    if ($Definition.Publish.Mode -eq 'UpdateExisting') { & $add 'Finalize' 'PATCH' $appUri (New-IpFinalizePatchBody -AppBody $AppBody -ContentVersionId '<content-version-id>') 'Differential PATCH computed from the current app at publish time: only changed properties are sent (body shown = all editable properties); a changed detection goes through rules with the requirement rules kept.' }
    else { & $add 'Finalize' 'PATCH' $appUri (New-IpFinalizePatchBody -AppBody $AppBody -ContentVersionId '<content-version-id>') 'Apply metadata and set committedContentVersion.' }
    & $add 'Verify' 'GET' $appUri $null 'Read the app again and check the content version and the preserved properties.'
    return $steps.ToArray()
}

function Publish-IpWin32App {
    <#
    .SYNOPSIS
        Publishes a .intunewin package as an Intune Win32 app (create, or new content version
        of an existing app). With -DryRun or -WhatIf, returns the planned requests and
        payloads without calling Graph. The app is never assigned.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][psobject]$Definition,
        [Parameter(Mandatory = $true)][string]$IntuneWinPath,
        [string]$BaseDirectory = '',
        [psobject]$UploadSettings,
        [switch]$DryRun,
        [string]$DryRunOutputFolder = '',
        [scriptblock]$ProgressCallback,
        [scriptblock]$CancelCheck,
        # UpdateExisting: last modification time read before the publication (import or confirmation).
        [string]$ExpectedLastModified = '',
        # UpdateExisting: publish even if the app changed in Intune since it was read.
        [switch]$Force
    )

    $blockSizeMB = 16; $maxRetries = 5; $renewBefore = 5; $httpTimeout = 30; $forceRenewal = 0
    if ($null -ne $UploadSettings) {
        if ($UploadSettings.PSObject.Properties['BlockSizeMB'] -and [int]$UploadSettings.BlockSizeMB -gt 0) { $blockSizeMB = [int]$UploadSettings.BlockSizeMB }
        if ($UploadSettings.PSObject.Properties['MaxRetries'] -and [int]$UploadSettings.MaxRetries -gt 0) { $maxRetries = [int]$UploadSettings.MaxRetries }
        if ($UploadSettings.PSObject.Properties['RenewSasBeforeExpiryMinutes'] -and [int]$UploadSettings.RenewSasBeforeExpiryMinutes -gt 0) { $renewBefore = [int]$UploadSettings.RenewSasBeforeExpiryMinutes }
        if ($UploadSettings.PSObject.Properties['HttpTimeoutMinutes'] -and [int]$UploadSettings.HttpTimeoutMinutes -gt 0) { $httpTimeout = [int]$UploadSettings.HttpTimeoutMinutes }
        if ($UploadSettings.PSObject.Properties['ForceSasRenewalAtBlock'] -and [int]$UploadSettings.ForceSasRenewalAtBlock -gt 0) { $forceRenewal = [int]$UploadSettings.ForceSasRenewalAtBlock }
    }
    if ($blockSizeMB -gt 100) { throw 'Configuration error: Upload.BlockSizeMB must be between 1 and 100.' }

    $metadata = Read-IpIntuneWinMetadata -Path $IntuneWinPath
    $appBody = New-IpWin32LobAppBody -Definition $Definition -IntuneWinMetadata $metadata -BaseDirectory $BaseDirectory
    $displayName = [string]$appBody.displayName
    $mode = [string]$Definition.Publish.Mode

    # UpdateExisting: read the app and check concurrency before the confirmation, so the operator is
    # not asked to confirm a publication that is then blocked. A block is thrown without logging here:
    # the caller logs it once (exception Data['IpConcurrency'] = $true).
    $existing = $null
    if ($mode -eq 'UpdateExisting' -and -not $DryRun -and -not $WhatIfPreference) {
        if ($null -eq (Get-IpGraphContext)) { throw 'Not connected to Microsoft Graph. Connect before publishing.' }
        try { $existing = Get-IpWin32App -AppId $Definition.Publish.ExistingAppId }
        catch {
            $message = $_.Exception.Message
            if ($message -notmatch '^\[') { $message = "[GetApp] $message" }
            Write-IpLog ("Publication failed at step GetApp: {0}" -f $message) -Level ERROR
            throw [InvalidOperationException]::new(("Publication failed at step GetApp: {0}" -f $message), $_.Exception)
        }
        $concurrency = Test-IpImportConcurrency -Definition $Definition -Current $existing -ExpectedLastModified $ExpectedLastModified
        if ($concurrency -and -not $Force) {
            $blocked = [InvalidOperationException]::new($concurrency)
            $blocked.Data['IpConcurrency'] = $true
            throw $blocked
        }
    }

    $execute = (-not $DryRun) -and $PSCmdlet.ShouldProcess($displayName, "Publish Intune Win32 app ($mode)")
    if (-not $execute) {
        $plan = @(Get-IpPublishPlan -Definition $Definition -AppBody $appBody -IntuneWinMetadata $metadata -BlockSizeMB $blockSizeMB)
        Write-IpLog ("Dry run: {0} Graph request(s) planned for '{1}' ({2}). Microsoft Graph was not called." -f $plan.Count, $displayName, $mode)
        foreach ($step in $plan) { Write-IpLog ("Dry run step {0}: {1} {2}. {3}" -f $step.Step, $step.Method, $step.Uri, $step.Note) }
        if ($DryRunOutputFolder) {
            [IO.Directory]::CreateDirectory($DryRunOutputFolder) | Out-Null
            for ($i = 0; $i -lt $plan.Count; $i++) {
                Save-IpJsonFile -Path (Join-Path $DryRunOutputFolder ('{0:D2}-{1}.json' -f ($i + 1), $plan[$i].Step)) -InputObject $plan[$i] -Confirm:$false -WhatIf:$false
            }
            Write-IpLog "Dry run payloads written to: $DryRunOutputFolder"
        }
        return [pscustomobject]@{ Mode = 'DryRun'; DisplayName = $displayName; PublishMode = $mode; AppId = $null; PortalUrl = $null; Assigned = $false; Plan = $plan }
    }

    if ($null -eq (Get-IpGraphContext)) { throw 'Not connected to Microsoft Graph. Connect before publishing.' }

    $step = 'Prepare'
    $appId = ''
    $createdApp = $false
    try {
        if ($mode -eq 'UpdateExisting') {
            $step = 'GetApp'
            $appId = [string]$existing.id
            if ([string]::IsNullOrWhiteSpace([string]$existing.committedContentVersion)) {
                Write-IpLog "[GetApp] Existing app $appId has no committed content version yet; a new content version will be committed." -Level WARN
            }
            Write-IpLog ("[GetApp] Updating existing app: {0} ({1})" -f $existing.displayName, $appId)
            if ($concurrency) { Write-IpLog ($concurrency + ' Publication forced by the operator.') -Level WARN }
            # Computed before the upload so that an unsupported case stops before any content is sent.
            $update = New-IpUpdatePatchBody -AppBody $appBody -Current $existing -ContentVersionId 'pending' -Definition $Definition
            $changeText = if ($update.Changes.Count) { (@($update.Changes | ForEach-Object { $_.Property }) -join ', ') } else { 'none' }
            Write-IpLog ("[GetApp] Differential update: {0} change(s): {1}. Requirement rules kept: {2}." -f $update.Changes.Count, $changeText, $update.PreservedRequirementRules)
            foreach ($warning in $update.Warnings) { Write-IpLog ('[GetApp] ' + $warning) -Level WARN }
        }
        else {
            $step = 'CreateApp'
            $appId = [string](New-IpWin32App -Body $appBody -Confirm:$false).id
            $createdApp = $true
        }

        $step = 'ContentVersion'
        $contentVersionId = New-IpContentVersion -AppId $appId -Confirm:$false
        $step = 'ContentFile'
        $file = New-IpContentFile -AppId $appId -ContentVersionId $contentVersionId -Body (New-IpContentFileBody $metadata) -Confirm:$false
        $step = 'SasWait'
        $fileState = Wait-IpContentFileState -FileUri $file.FileUri -SuccessStates @('azureStorageUriRequestSuccess') -FailureStates @('azureStorageUriRequestFailed') -Step 'SasWait' -RequireAzureStorageUri -CancelCheck $CancelCheck
        $step = 'Upload'
        Write-IpLog ("[Upload] Uploading encrypted payload. Size={0}; BlockSize={1}; MaxRetries={2}" -f (Format-IpByteSize $metadata.SizeEncrypted), (Format-IpByteSize ($blockSizeMB * 1MB)), $maxRetries)
        [void](Send-IpIntuneWinContent -IntuneWinPath $metadata.IntuneWinPath -ContentEntryName $metadata.ContentEntryName -AzureStorageUri ([string]$fileState.azureStorageUri) -FileUri $file.FileUri -BlockSizeMB $blockSizeMB -MaxRetries $maxRetries -RenewBeforeExpiryMinutes $renewBefore -HttpTimeoutMinutes $httpTimeout -ForceSasRenewalAtBlock $forceRenewal -ProgressCallback $ProgressCallback -CancelCheck $CancelCheck)
        $step = 'Commit'
        Submit-IpContentFileCommit -FileUri $file.FileUri -CommitBody (New-IpCommitBody $metadata) -CancelCheck $CancelCheck -Confirm:$false
        $step = 'Finalize'
        if ($mode -eq 'UpdateExisting') {
            $update.Body.committedContentVersion = $contentVersionId
            Complete-IpAppContentVersion -AppId $appId -PatchBody $update.Body -Confirm:$false
        }
        else {
            Complete-IpAppContentVersion -AppId $appId -PatchBody (New-IpFinalizePatchBody -AppBody $appBody -ContentVersionId $contentVersionId) -Confirm:$false
        }
    }
    catch {
        $message = $_.Exception.Message
        if ($message -notmatch '^\[') { $message = "[$step] $message" }
        if ($createdApp -and $appId) {
            Write-IpLog ("The app {0} was created but its content is not committed. Delete it in the Intune admin center, or publish again with Publish.Mode=UpdateExisting and ExistingAppId={0}." -f $appId) -Level WARN
        }
        Write-IpLog ("Publication failed at step {0}: {1}" -f $step, $message) -Level ERROR
        if ($_.Exception -is [OperationCanceledException]) { throw }
        throw [InvalidOperationException]::new(("Publication failed at step {0}: {1}" -f $step, $message), $_.Exception)
    }

    # Post-publication check: content version, preserved properties and requirement rules.
    $verificationIssues = @()
    $verifiedState = $null
    try {
        $after = Get-IpWin32App -AppId $appId -Step 'Verify'
        $verifiedState = [pscustomobject]@{
            AppId                = [string]$after.id
            DisplayName          = [string]$after.displayName
            ContentVersion       = [string]$after.committedContentVersion
            LastModifiedDateTime = ConvertTo-IpIsoUtc (Get-IpObjectProperty $after 'lastModifiedDateTime')
        }
        if ($mode -eq 'UpdateExisting') { $verificationIssues = @(Test-IpUpdateResult -Before $existing -After $after -Update $update -ContentVersionId $contentVersionId) }
        elseif ([string]$after.committedContentVersion -ne $contentVersionId) { $verificationIssues = @(("committedContentVersion is {0}, expected {1}." -f $after.committedContentVersion, $contentVersionId)) }
        if ($verificationIssues.Count) { foreach ($issue in $verificationIssues) { Write-IpLog ('[Verify] ' + $issue) -Level WARN } }
        else { Write-IpLog ('[Verify] Content version {0} committed; preserved properties and requirement rules unchanged.' -f $contentVersionId) }
    }
    catch {
        $verificationIssues = @('The app could not be read again for verification: ' + $_.Exception.Message)
        Write-IpLog ('[Verify] ' + $verificationIssues[0]) -Level WARN
    }

    $portalUrl = Get-IpIntunePortalAppUrl -AppId $appId
    $context = Get-IpGraphContext
    Write-IpLog ("Publication completed. AppId={0}; ContentVersion={1}; DisplayName={2}" -f $appId, $contentVersionId, $displayName)
    Write-IpLog ("Intune admin center: {0}" -f $portalUrl)
    Write-IpLog 'The app is not assigned. Create assignments in the Intune admin center.' -Level WARN
    return [pscustomobject]@{
        Mode               = 'Publish'
        DisplayName        = $displayName
        PublishMode        = $mode
        AppId              = $appId
        ContentVersionId   = $contentVersionId
        ContentFileId      = $file.FileId
        UpdatedExistingApp = ($mode -eq 'UpdateExisting')
        TenantId           = if ($context) { [string]$context.TenantId } else { '' }
        PortalUrl          = $portalUrl
        Assigned           = $false
        Changes            = $(if ($mode -eq 'UpdateExisting') { @($update.Changes) } else { @() })
        Warnings           = $(if ($mode -eq 'UpdateExisting') { @($update.Warnings) } else { @() })
        VerificationIssues = $verificationIssues
        VerifiedState      = $verifiedState
        Plan               = @()
    }
}

#endregion

Export-ModuleMember -Function @(
    'Get-IpGraphScopes', 'Set-IpGraphSettings', 'Get-IpGraphSettings', 'Assert-IpGraphModule', 'Get-IpGraphContext', 'Connect-IpGraph', 'Disconnect-IpGraph', 'Test-IpGraphBrokerLoaded', 'Get-IpGraphRuntimeInfo', 'Resolve-IpSignInMode',
    'Get-IpHttpErrorInfo', 'Get-IpRetryDelaySeconds', 'Invoke-IpGraphRequest', 'Get-IpGraphCollection',
    'Get-IpWin32App', 'Find-IpWin32App', 'Get-IpWin32AppDetails', 'ConvertTo-IpWin32AppSummary',
    'New-IpWin32App', 'New-IpContentVersion', 'New-IpContentFile', 'Wait-IpContentFileState', 'Request-IpSasRenewal',
    'Get-IpSasExpiryUtc', 'Test-IpSasRenewalNeeded', 'Join-IpSasQuery', 'New-IpBlockId', 'New-IpBlockListXml', 'Send-IpAzureBlobWithRetry', 'Send-IpIntuneWinContent',
    'Submit-IpContentFileCommit', 'Complete-IpAppContentVersion',
    'Get-IpPublishPlan', 'Publish-IpWin32App'
)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBwrH18DJvB+Mlr
# lIyu1uuA3okhPDbyU0DIHk08zCISRqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFGHHJ7Q+lrUi/+BeRozNECxkHZ03NNTrDwZhUOAOOqtMA0GCSqG
# SIb3DQEBAQUABIIBgK8rwxofbjNlyvmbl5Q51gfx+6ELuaURkUIaYe+5Eiu0oIy1
# /HCnXsaBT9I6ueVf9YH81zmTXe/6tCkn/LdBX6p0cZP3qpfwv7ncPXnlbDclh6YP
# jbyllMUCXrB0ZOS+zAWaPRC6qY6jhcRS1V/w2J5+TUH9aKmFEvZJnrRdnlL4bwG3
# pjNpCYp18JmLgLrRwzovEPVQJd5OHP4zqLyNPSnzTbqCnTBCd823CfutzR4xESDb
# Uh6Fe9S+vPcoRXkoiCN+18mOrMbRB22Ype7nyYRZ+80HANxYY/4m7JWjTZlDYdH+
# JpvvgIrAxkaP3opPWaV+Gzbetg7Lfs864hcm+5n8DZlAz578KngaZEDETSZV9WCO
# JEz/jMogJHdaVwq6iO1/TYEnAQkSYZFm6L4vgVsUE/kcKIK9ZgRDShjLL5OcN9/A
# ipMZ3m1f2YnCuD704dAwajL5PA0ct53dvSEyfVarZHZuvBkrP8imAt4+0fTtr3OU
# 0jtFJLdCCJ76NeIV5qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzVaMC8GCSqGSIb3DQEJBDEiBCClBw14KtwH1WkpOtt2LgjcdHcDNGCwdSYGu4sM
# pDm2qTANBgkqhkiG9w0BAQEFAASCAgAx1cSBGGx6ufuWnZ6S4TKaZgS3NSBSJVT7
# wrDo9VR++lKN55YC2FY39UBVXINGowF6nnIuptmEjVAclEzSAofzp9YSeMElNGFu
# ZRkfCaiFpUcTr5evgO9sn02jhlZoeguClEBvVbxaQ8xeY+cLBLQuFudovhJoj6pZ
# vBn4NIpMA07DKkmWD8UGnel684Ogr21jzAiY4sghG3tF9i1IY3e9B+c2YA/6ljXg
# XNycYczp22MPTgc51CmQZvc2Ab2Djl2AhrEunhuMlVX5hotF2GW/KrJsNmf9joIR
# aayJOHhEt3t5V30KMmlNYmYOaImup4osPnNGvw3QOuJa8VVy6oO3IbzhlAh8pxSI
# 4bFwCtFuQzfWuP2A8GJohUSPpOOEvnRGEjlM7svcsZKcB9Lekolvr1WbxwaxxQKX
# DguWhUk/go5sp+t+jPlv/HJz3+tMs3Oer0a4lK+xGkjTFB40SCpjbs1WO7TETtJH
# Upqy1hBwvCKUicyZvDKCkFpyvGeT53+0/2Jj2Uhy8F/hb85ePSoQNos7ThA2wScc
# cIT+WkDFCjd3pJDriUbcDjpMp34ksEvI6B1nkf7lrMJg4v0KSpGFCIhOEWrP5Fc6
# gatqVgqKDznn3BbgH4p/rTc4b1k8ImN+Oet9n5nypMiv7h9RbhDmjPkPPN6bQOp9
# 4jshT9j5rw==
# SIG # End signature block
