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
1.5

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version : 1.4
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
$ScriptVersion = "1.5"

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
            Import-Module -Name $modulePath -MinimumVersion '1.0.23' -Prefix Core -ErrorAction Stop
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

function Invoke-GraphGet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
}

function Invoke-GraphGetAllPages {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    $items = @()
    $nextUri = $Uri

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $page = Invoke-GraphGet -Uri $nextUri
        $value = Get-ObjectValue -InputObject $page -Name "value"

        if ($null -ne $value) {
            foreach ($item in @($value)) {
                $items += $item
            }
        }

        $nextUri = Get-ObjectValue -InputObject $page -Name "@odata.nextLink"
    }

    return $items
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
if ($OutputRoot -match '\{\{[^}]+\}\}') {
    throw "OutputRoot contains unresolved configuration token(s): $OutputRoot"
}

$OutputRoot = CoreInitializeScriptEnvironment -OutputPathInit $OutputRoot -LogFileName $scriptName
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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDuIDzRrEKiuqfu
# e3acNJwDD1o+2fNKs76qlR++G5pqHqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDAG5STPDVdyfc+j32MXXACl3g7YmV7FPJcknnqgTfsODANBgkqhkiG9w0B
# AQEFAASCAYAWwe/T8n6ZcEXm3MxkhvTXDeZwHTgiGGcAbYrT57hYBF5WoTRefHnL
# 4yLEM2N7xO2XCdddsjwSFxEp7hU69iu15hUtx8NRMZFEdu+5OGRJN09jOJeZp1DR
# ReS59T40jsoVhDF1nuhdmnRQxmNSegHnaKsRoeRSnRRQmY8ObsW+Gxt8fYARH5XG
# kROHI2j0+yiRauvuCCqzGCt4D8QXJmoDmhpEwEuRWwmxRhm1uVAhQEJwWf4iP2Qm
# eOjmyJImlp759wJ+KmjGhJg8b9J/JrlMhO8nkYZt4CFk8kqmuE4q1ELzkGE+NSet
# Yo8k5N+JXt45XWq+xoNChE/KIjXDemJ5gVKhHcLqbruuQwwESzlYy0K25rSZ2Xgp
# Zp09AMaMQNACAH0ceiFXG7zt2juzsyifxB8mwlO7mJC8+IkvMhN2OkBsl3Gz1SYV
# b3v9uRA5pbJkJ6RPXvoq9QfYnq1pn7hYn1mtecFq9h6lvVSnJqYDc3yiLrliRnNd
# VQOQv1oT7eU=
# SIG # End signature block
