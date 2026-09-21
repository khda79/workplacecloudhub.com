<#
.SYNOPSIS
    Exports Microsoft Intune remediation scripts with Microsoft Graph PowerShell.

.DESCRIPTION
    Connects to Microsoft Graph with either app-only certificate authentication
    or delegated interactive authentication, lists Intune remediation
    scripts, exports metadata to CSV/JSON, and saves associated detection and
    remediation script contents as .ps1 files.

    Intune remediations are exposed by Microsoft Graph as deviceHealthScripts.
    The script uses Invoke-MgGraphRequest so it can call the beta endpoint while
    still relying on Microsoft.Graph.Authentication for authentication and token
    handling.

.REQUIREMENTS
    PowerShell module:
    - Microsoft.Graph.Authentication

    App-only Graph application permission for the custom app:
    - DeviceManagementScripts.Read.All

    Interactive delegated permission for Microsoft Graph PowerShell:
    - DeviceManagementScripts.Read.All
.VERSION
1.9

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version : 1.9
    Minimum application permissions: DeviceManagementScripts.Read.All, Group.Read.All
#>

[CmdletBinding()]
param(
    [string]$Tenant = "test",

    [string]$AppId = "00000000-0000-0000-0000-000000000000",

    [string]$TenantId = "00000000-0000-0000-0000-000000000000",

    [string]$Thumbprint = "0000000000000000000000000000000000000000",

    [string]$OrgDomain = "contoso.onmicrosoft.com",

    [string]$OutputRoot = "{{DataAllRootPath}}\Intune\Remediations\Exports",

    [ValidateSet("beta", "v1.0")]
    [string]$GraphApiVersion = "beta",

    [switch]$InteractiveAuth,

    [switch]$DeviceCodeAuth,

    [string[]]$InteractiveScopes = @("DeviceManagementScripts.Read.All"),

    [switch]$SkipAssignments,
    [int]$MaxItems = 0
)
if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
    foreach ($smartM365LimitName in @('TopUsers','TopMailboxes','MaxDevices','MaxSites','MaxTeams','MaxApps','MaxPolicies','Limit','MaxPages')) {
        $smartM365LimitVariable = Get-Variable -Name $smartM365LimitName -Scope Script -ErrorAction SilentlyContinue
        if ($smartM365LimitVariable -and -not $PSBoundParameters.ContainsKey($smartM365LimitName) -and $null -ne $smartM365LimitVariable.Value) {
            Set-Variable -Name $smartM365LimitName -Value ([int]$MaxItems) -Scope Script
        }
    }
}

$ErrorActionPreference = "Stop"
$ScriptVersion = "1.9"

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($candidate in $candidates) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365GlobalConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
            Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null
        }
        else {
            if (-not (Test-Path -LiteralPath $templatePath)) {
                $message = @(
                    "Local configuration file not found: $configPath",
                    "Template to copy is also missing: $templatePath",
                    'Create the .local.json file from a safe template, then run the script again.'
                ) -join [Environment]::NewLine
                throw $message
            }

            Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
            Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
            Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
        }
    }

    try {
        return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
}

function Resolve-SmartM365ConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    if ($Value -notmatch '\{\{[^}]+\}\}') {
        return $Value
    }

    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
            $scriptRootValue = Get-Variable -Name ScriptRoot -ValueOnly -ErrorAction SilentlyContinue
    $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($scriptRootValue) { $scriptRootValue } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }

            $tokenValue = Resolve-SmartM365ConfigValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}
function Get-ScriptLocalConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                return Resolve-SmartM365ConfigValue -Value $property.Value
            }
        }
        else {
            return Resolve-SmartM365ConfigValue -Value $property.Value
        }
    }


    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $PSCommandPath -Parent }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $globalProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) {
            return $DefaultValue
        }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

$ScriptLocalConfig = Get-ScriptLocalConfig

$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
foreach ($configName in @('AppId', 'TenantId', 'Thumbprint', 'OrgDomain', 'OutputRoot', 'GraphApiVersion', 'InteractiveScopes')) {
    Set-Variable -Name $configName -Value (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name $configName -DefaultValue (Get-Variable -Name $configName -ValueOnly)) -Scope Script
}

function Import-SmartM365CorePreflight {
    if (Get-Command Invoke-CoreSmartM365Preflight -ErrorAction SilentlyContinue) { return }

        $scriptRootValue = Get-Variable -Name ScriptRoot -ValueOnly -ErrorAction SilentlyContinue
    $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($scriptRootValue) { $scriptRootValue } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
    while ($searchRoot) {
        $modulePath = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $modulePath) {
            Import-Module -Name $modulePath -MinimumVersion '1.0.24' -Prefix Core -ErrorAction Stop
            return
        }

        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }

    throw 'SmartM365.Core module was not found. Preflight checks cannot run.'
}

function Import-RequiredGraphModule {
    $moduleName = "Microsoft.Graph.Authentication"

    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        throw @"
Required module '$moduleName' is not installed.

Install it with:
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

Or install the full SDK with:
Install-Module Microsoft.Graph -Scope CurrentUser
"@
    }

    Import-Module $moduleName -ErrorAction Stop
}

function Connect-IntuneGraph {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$Tenant,

        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint,

        [Parameter(Mandatory = $true)]
        [bool]$UseInteractiveAuth,

        [Parameter(Mandatory = $true)]
        [bool]$UseDeviceCodeAuth,

        [Parameter(Mandatory = $true)]
        [string[]]$Scopes
    )

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    if ($UseInteractiveAuth -or $UseDeviceCodeAuth) {
        if ($UseDeviceCodeAuth) {
            Write-Output "Connecting to Microsoft Graph with device-code authentication..."
            Connect-MgGraph `
                -TenantId $Tenant `
                -Scopes $Scopes `
                -UseDeviceCode `
                -NoWelcome `
                -ErrorAction Stop | Out-Null
        }
        else {
            Write-Output "Connecting to Microsoft Graph with interactive browser authentication..."
            Connect-MgGraph `
                -TenantId $Tenant `
                -Scopes $Scopes `
                -NoWelcome `
                -ErrorAction Stop | Out-Null
        }
    }
    else {
        Write-Output "Connecting to Microsoft Graph with app-only certificate authentication..."
        Connect-MgGraph `
            -TenantId $Tenant `
            -ClientId $ClientId `
            -CertificateThumbprint $CertificateThumbprint `
            -NoWelcome `
            -ErrorAction Stop | Out-Null
    }

    $context = Get-MgContext
    if ($null -eq $context) {
        throw "Microsoft Graph connection failed. Get-MgContext returned no context."
    }

    Write-Output "Connected to tenant: $($context.TenantId)"
}

function Get-ObjectValue {
    param(
        [AllowNull()]
        $InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        return $InputObject[$Name]
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }

    return $null
}

function Test-ObjectProperty {
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) { return $false }
    if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject.Contains($Name) }
    return $null -ne $InputObject.PSObject.Properties[$Name]
}

function Get-IntuneRemediationRetryDelaySeconds {
    param(
        [Parameter(Mandatory = $true)]$ErrorRecord,
        [int]$DefaultSeconds,
        [int]$MaximumSeconds = 300
    )

    $headers = $null
    try { $headers = $ErrorRecord.Exception.Response.Headers } catch {}
    $value = $null
    if ($null -ne $headers) {
        try { $value = @($headers.GetValues('Retry-After') | Select-Object -First 1)[0] } catch {}
        if ($null -eq $value -and $headers -is [Collections.IDictionary]) {
            foreach ($name in @('Retry-After', 'retry-after')) {
                if ($headers.Contains($name)) { $value = $headers[$name]; break }
            }
        }
    }
    if ($null -eq $value) {
        try { $value = $ErrorRecord.Exception.Data['Retry-After'] } catch {}
    }
    $seconds = 0
    if ($null -ne $value -and [int]::TryParse([string]$value, [ref]$seconds) -and $seconds -gt 0) {
        return [Math]::Min($seconds, $MaximumSeconds)
    }
    return [Math]::Min([Math]::Max(1, $DefaultSeconds), $MaximumSeconds)
}

function Invoke-GraphGet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [ValidateRange(1, 10)][int]$MaxAttempts = 5
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try { return Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop }
        catch {
            $statusCode = $null
            try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch {}
            if ($null -eq $statusCode) { try { $statusCode = [int]$_.Exception.Data['StatusCode'] } catch {} }
            $transient = $statusCode -in @(408, 409, 429, 500, 502, 503, 504) -or
                $_.Exception.Message -match '(?i)throttl|TooManyRequests|temporarily unavailable|timeout|timed out'
            if (-not $transient -or $attempt -ge $MaxAttempts) { throw }
            $fallback = [int][Math]::Min(300, [Math]::Pow(2, $attempt) * 5)
            $delay = Get-IntuneRemediationRetryDelaySeconds -ErrorRecord $_ -DefaultSeconds $fallback -MaximumSeconds 300
            Write-Warning ("Graph request transiently failed. Status={0}; attempt {1}/{2}; retry in {3}s." -f $statusCode, $attempt, $MaxAttempts, $delay)
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-GraphGetAllPages {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $items = New-Object System.Collections.Generic.List[object]
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $nextUri = $Uri
    $pageNumber = 0

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        if (-not $visited.Add($nextUri)) {
            throw "Graph pagination returned a repeated @odata.nextLink; collection is incomplete."
        }
        $pageNumber++
        $page = Invoke-GraphGet -Uri $nextUri
        if (-not (Test-ObjectProperty -InputObject $page -Name 'value')) {
            throw "Graph page $pageNumber returned an invalid collection response without a value property."
        }
        $value = Get-ObjectValue -InputObject $page -Name "value"

        foreach ($item in @($value)) {
            if ($null -ne $item) { [void]$items.Add($item) }
        }

        $nextUri = Get-ObjectValue -InputObject $page -Name "@odata.nextLink"
    }

    return $items.ToArray()
}

function ConvertFrom-GraphScriptContent {
    param(
        [AllowNull()]
        $Content
    )

    if ($null -eq $Content) {
        return $null
    }

    if ($Content -is [byte[]]) {
        return [Text.Encoding]::UTF8.GetString($Content)
    }

    $contentString = [string]$Content
    if ([string]::IsNullOrWhiteSpace($contentString)) {
        return $null
    }

    try {
        $bytes = [Convert]::FromBase64String($contentString)
        return [Text.Encoding]::UTF8.GetString($bytes)
    }
    catch {
        return $contentString
    }
}

function ConvertTo-SafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $invalidChars = [IO.Path]::GetInvalidFileNameChars()
    $builder = New-Object Text.StringBuilder

    foreach ($char in $Value.ToCharArray()) {
        if ($invalidChars -contains $char) {
            [void]$builder.Append("_")
        }
        else {
            [void]$builder.Append($char)
        }
    }

    $safe = $builder.ToString().Trim()
    $safe = $safe -replace "\s+", " "
    $safe = $safe -replace "[\\/:\*\?`"<>|]", "_"

    if ([string]::IsNullOrWhiteSpace($safe)) {
        return "UnnamedRemediation"
    }

    if ($safe.Length -gt 120) {
        return $safe.Substring(0, 120).Trim()
    }

    return $safe
}

function Save-Utf8NoBom {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [AllowNull()]
        [string]$Content
    )

    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $parentPath = Split-Path -Path $fullPath -Parent

    if (-not [string]::IsNullOrWhiteSpace($parentPath)) {
        [IO.Directory]::CreateDirectory($parentPath) | Out-Null
    }

    if ($null -eq $Content) {
        $Content = ""
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($fullPath, $Content, $utf8NoBom)
}

function Initialize-ExportFolders {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [string]$Timestamp
    )

    $rootPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputRoot)
    New-Item -Path $rootPath -ItemType Directory -Force | Out-Null

    $runPath = Join-Path -Path $rootPath -ChildPath $Timestamp
    $index = 1
    while (Test-Path -LiteralPath $runPath) {
        $runPath = Join-Path -Path $rootPath -ChildPath "$Timestamp-$index"
        $index++
    }

    New-Item -Path $runPath -ItemType Directory -Force | Out-Null

    return [pscustomobject]@{
        Root = $rootPath
        Run  = $runPath
    }
}

Import-SmartM365CorePreflight
$scriptName = 'SmartM365-Export-IntuneRemediations'
$logFileBaseName = 'SmartM365-Intune-RemediationsScripts'
if ($OutputRoot -match '\{\{[^}]+\}\}') {
    throw "OutputRoot contains unresolved configuration token(s): $OutputRoot"
}

$OutputRoot = CoreInitializeScriptEnvironment -OutputPathInit $OutputRoot -LogFileName $logFileBaseName
$script:SmartM365TranscriptStarted = $false
Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
$script:SmartM365TranscriptStarted = $true
function Stop-IntuneRemediationsTranscript {
    [CmdletBinding()]
    param()

    if (-not $script:SmartM365TranscriptStarted) { return }

    try { Stop-Transcript | Out-Null } catch { }
    try {
        $smartM365TranscriptPath = $null
        $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue
        if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
            $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
        }
        else {
            $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue
            if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
                $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
            }
        }
        if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath }
    }
    catch { }
    $script:SmartM365TranscriptStarted = $false
}

trap {
    $errorRecord = $_
    if ($script:SmartM365TranscriptStarted) {
        Stop-IntuneRemediationsTranscript
        Complete-CoreSmartM365ExecutionContext -Status Failed -ErrorRecord $errorRecord
    }
    throw $errorRecord
}

$tenantForConnection = if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $TenantId } else { $OrgDomain }
$graphBaseUri = "https://graph.microsoft.com/$GraphApiVersion"
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$exportFolders = Initialize-ExportFolders -OutputRoot $OutputRoot -Timestamp $timestamp
$exportRoot = $exportFolders.Run
$scriptsRoot = Join-Path -Path $exportRoot -ChildPath "Scripts"

New-Item -Path $scriptsRoot -ItemType Directory -Force | Out-Null

Import-RequiredGraphModule

Connect-IntuneGraph `
    -ClientId $AppId `
    -Tenant $tenantForConnection `
    -CertificateThumbprint $Thumbprint `
    -UseInteractiveAuth ([bool]$InteractiveAuth) `
    -UseDeviceCodeAuth ([bool]$DeviceCodeAuth) `
    -Scopes $InteractiveScopes

Invoke-CoreSmartM365Preflight -ScriptName $scriptName -RequiredModules @('Microsoft.Graph.Authentication') -OutputPaths @($exportFolders.Root, $exportFolders.Run, $scriptsRoot, $global:LogPath) -RequiredGraphApplicationPermissions @('DeviceManagementScripts.Read.All','Group.Read.All') -GraphProbeUris @("$graphBaseUri/deviceManagement/deviceHealthScripts?`$top=1") | Out-Null

Write-Output "Retrieving Intune remediation list..."
$listUri = "$graphBaseUri/deviceManagement/deviceHealthScripts"
$remediationList = Invoke-GraphGetAllPages -Uri $listUri
if ($MaxItems -gt 0) {
    $preCount = @($remediationList).Count
    $remediationList = @($remediationList | Sort-Object displayName | Select-Object -First $MaxItems)
    Write-Output ("MaxItems enabled: restricted remediations from {0} to {1}." -f $preCount, @($remediationList).Count)
}

$inventory = @()

foreach ($remediation in @($remediationList)) {
    $id = Get-ObjectValue -InputObject $remediation -Name "id"
    $displayNameValue = Get-ObjectValue -InputObject $remediation -Name "displayName"
    $displayName = if ($displayNameValue) { [string]$displayNameValue } else { "Unnamed remediation" }

    Write-Output "Exporting: $displayName ($id)"

    $detailUri = "$graphBaseUri/deviceManagement/deviceHealthScripts/$id"
    $detail = Invoke-GraphGet -Uri $detailUri

    $safeName = ConvertTo-SafeFileName -Value $displayName
    $scriptFolder = Join-Path -Path $scriptsRoot -ChildPath "$safeName`_$id"
    New-Item -Path $scriptFolder -ItemType Directory -Force | Out-Null

    $detectionScript = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name "detectionScriptContent")
    $remediationScript = ConvertFrom-GraphScriptContent -Content (Get-ObjectValue -InputObject $detail -Name "remediationScriptContent")

    $detectionPath = $null
    $remediationPath = $null

    if (-not [string]::IsNullOrWhiteSpace($detectionScript)) {
        $detectionPath = Join-Path -Path $scriptFolder -ChildPath "Detection.ps1"
        Save-Utf8NoBom -Path $detectionPath -Content $detectionScript
    }

    if (-not [string]::IsNullOrWhiteSpace($remediationScript)) {
        $remediationPath = Join-Path -Path $scriptFolder -ChildPath "Remediation.ps1"
        Save-Utf8NoBom -Path $remediationPath -Content $remediationScript
    }

    $assignmentsPath = $null
    $assignmentCount = $null

    if (-not $SkipAssignments) {
        try {
            $assignmentsUri = "$graphBaseUri/deviceManagement/deviceHealthScripts/$id/assignments"
            $assignments = Invoke-GraphGetAllPages -Uri $assignmentsUri
            $assignmentCount = @($assignments).Count
            $assignmentsPath = Join-Path -Path $scriptFolder -ChildPath "Assignments.json"
            Save-Utf8NoBom -Path $assignmentsPath -Content (ConvertTo-Json -InputObject @($assignments) -Depth 20)
        }
        catch {
            $assignmentCount = "Error"
            Write-Warning "Failed to retrieve assignments for $displayName ($id): $($_.Exception.Message)"
        }
    }

    $metadataPath = Join-Path -Path $scriptFolder -ChildPath "Metadata.json"
    Save-Utf8NoBom -Path $metadataPath -Content ($detail | ConvertTo-Json -Depth 20)

    $roleScopeTagIds = Get-ObjectValue -InputObject $detail -Name "roleScopeTagIds"

    $inventory += [pscustomobject]@{
        Id                    = $id
        DisplayName           = $displayName
        Description           = Get-ObjectValue -InputObject $detail -Name "description"
        Publisher             = Get-ObjectValue -InputObject $detail -Name "publisher"
        Version               = 1.4
        CreatedDateTime       = Get-ObjectValue -InputObject $detail -Name "createdDateTime"
        LastModifiedDateTime  = Get-ObjectValue -InputObject $detail -Name "lastModifiedDateTime"
        RunAsAccount          = Get-ObjectValue -InputObject $detail -Name "runAsAccount"
        RunAs32Bit            = Get-ObjectValue -InputObject $detail -Name "runAs32Bit"
        EnforceSignatureCheck = Get-ObjectValue -InputObject $detail -Name "enforceSignatureCheck"
        RoleScopeTagIds       = if ($roleScopeTagIds) { @($roleScopeTagIds) -join ";" } else { $null }
        AssignmentCount       = $assignmentCount
        Folder                = $scriptFolder
        DetectionScriptPath   = $detectionPath
        RemediationScriptPath = $remediationPath
        AssignmentsPath       = $assignmentsPath
        MetadataPath          = $metadataPath
    }
}

$csvPath = Join-Path -Path $exportRoot -ChildPath "IntuneRemediations.csv"
$jsonPath = Join-Path -Path $exportRoot -ChildPath "IntuneRemediations.json"
$exportInfoPath = Join-Path -Path $exportRoot -ChildPath "ExportInfo.json"
$inventoryColumns = @(
    "Id",
    "DisplayName",
    "Description",
    "Publisher",
    "Version",
    "CreatedDateTime",
    "LastModifiedDateTime",
    "RunAsAccount",
    "RunAs32Bit",
    "EnforceSignatureCheck",
    "RoleScopeTagIds",
    "AssignmentCount",
    "Folder",
    "DetectionScriptPath",
    "RemediationScriptPath",
    "AssignmentsPath",
    "MetadataPath"
)
$sortedInventory = @($inventory | Sort-Object DisplayName)

Publish-CoreSmartM365Csv -Data $sortedInventory -TimestampedPath $csvPath -Columns $inventoryColumns | Out-Null

Save-Utf8NoBom -Path $jsonPath -Content (ConvertTo-Json -InputObject $sortedInventory -Depth 20)

$exportInfo = [ordered]@{
    Timestamp             = $timestamp
    ExportedAtLocal       = (Get-Date).ToString("o")
    TenantId              = $tenantForConnection
    GraphApiVersion       = $GraphApiVersion
    RemediationCount      = $inventory.Count
    OutputRoot            = $exportFolders.Root
    ExportPath            = $exportFolders.Run
    ScriptsPath           = $scriptsRoot
    LogPath               = $global:LogPath
    TranscriptPath        = $global:logTranscriptFile
    AuthenticationMode    = if ($InteractiveAuth) { "Interactive" } elseif ($DeviceCodeAuth) { "DeviceCode" } else { "Certificate" }
}

Save-Utf8NoBom -Path $exportInfoPath -Content ($exportInfo | ConvertTo-Json -Depth 5)

Write-Output ""
Write-Output "Export completed."
Write-Output "Remediations exported: $($inventory.Count)"
Write-Output "CSV: $csvPath"
Write-Output "JSON: $jsonPath"
Write-Output "Export info: $exportInfoPath"
Write-Output "Scripts folder: $scriptsRoot"
Write-Output "Log folder: $global:LogPath"
Write-Output "Transcript: $global:logTranscriptFile"

if ($script:SmartM365TranscriptStarted) {
    Stop-IntuneRemediationsTranscript
    Complete-CoreSmartM365ExecutionContext -Status Auto
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA9ybN0L+Tr4sPz
# 0PpiixdPdzTPq7BXgpoLf6cuJQYmdKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILkXDUdSMXga0UdbYxLkauAdZHGBpuALs/0tdaDky3MYMA0GCSqG
# SIb3DQEBAQUABIIBgCY7r/nmJ9HfUPXenDKNMV1jRBHTeGVZ9Ga9FB0bm4N8gQ/T
# Ux7vRgVYi653SgWCaVerbEQOCc7+X3ypvVverMo9/zICIstJllD5dbaJy70vAXKc
# i1zYisU5CpdWsozDYCnLksOt4ct8gTd9nmfxLLnxVx4pnUYVixDw7fmvcJHjVOkV
# gUAKiSBhAZSIDGsdPJDNBiJ6Qvm3cBeHOiJAWsPQLHGDh3ihcah7hm0oZQmRQQ3g
# P+P/mebwP4l2WK0lUMvp+OOqEXxWADk1F6dh20W78VoJyf42N4TmjapDzLbRRKgq
# fse6gSu511yHaV+8FfQodaQrAmu0wu2DxpkfBmNKl4rsQUZgdwxwjaBgEjzTiuRg
# /4+kEoKt0NrehXyeNd4uD71upe71d+Bb9pwLdQQjngSg4QB2YpRk1li8VZoOdAX6
# 5jlNe984+BUaUXvdbWYRSaUnm6ikOuaCDHY/+5MCa/3XznCG8f/RqgWLRh1aDRFz
# wLVP+FdPyCfsJPrCcqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjExMDA0
# MzNaMC8GCSqGSIb3DQEJBDEiBCDrZXtXvAWfDj/Jn7kvEVSkRi5EwqbUQ0ZvLV75
# 3LXnWjANBgkqhkiG9w0BAQEFAASCAgB0hS6HzmrrhpdjibmtEuXkPEmA4ooTUcTu
# WepsPRAsnyskjK9NqiiUTXmGHoPtgO9c14oVDJhsVPvtW0UyMhA6VCl7VNmPfOjY
# iISCzXYZIdFh7az9Qed3O7c7xEzwfuuRED3YAoGFQ3cm8NXL6hSZ9hJuZjVVXRGM
# r+rE27g+TgDBDu9oN6kAdX0fjNid3BFaiFDGmoJbaJZnE40vPs+z8n43mfpTpYze
# BY7/43SxMKHiSf5gbz6Ci9f0jjTfrgAHMGbh1n9KBXrF5qTUOaNFJZTJ8h4jL6Aa
# KBeX/KTW7JgU6BlAnuUtKyYRcGWCoBGjlBgbYikJIx4nxV0DXddCY0civRAdlyZD
# jAsY+J6BqGvXwLb0WStlolqQa6rga7+faHpppcDHDl2tajOG5QwBrYhHn6prs8w1
# RHc2IYHP1sAXL8yoeFX1dy59gJqLqycxolD6ykAQRAb4L/wp94iu2FXfgxzpehyG
# Ml/EGy1Dm4qZfSQt5WAJO0VExKP2DokIVuknJOGy/Q5sn6puTX6ABFiMuzpMtbce
# wTcLOnQk5L9fTBMYceogtsSh1km5c8Xm7w6vozD8tqeNr9vfLmzRt1R/SkY4/Z4j
# 4v7l5RA/2vdqRCHggwFMb/ojOMPG1rsYnzPTY6zvUvuruqZqwdSkBfJCAq3NJWBx
# mydvatxpHw==
# SIG # End signature block
