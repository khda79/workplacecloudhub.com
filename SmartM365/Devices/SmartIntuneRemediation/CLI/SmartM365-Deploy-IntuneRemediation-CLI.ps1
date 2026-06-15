#Requires -Version 7.0

<#
.SYNOPSIS
    Deploys one or more SmartM365 Intune remediation packages.

.DESCRIPTION
    Creates or updates Microsoft Intune remediation packages exposed by
    Microsoft Graph as deviceHealthScripts. The script accepts SmartM365
    detection scripts, remediation scripts, or folders that contain standard
    SmartM365 remediation script pairs.

    Authentication is delegated and interactive only. The script never accepts
    application IDs, certificate thumbprints, client secrets, or app-only
    authentication parameters.

.PARAMETER Path
    One or more files or folders to deploy. A folder can contain
    SmartM365-*-Detection.ps1, SmartM365-*-DetectionOnly.ps1, and optional
    SmartM365-*-Remediation.ps1 files.

.PARAMETER Tenant
    Tenant profile key. Defaults to test.

.PARAMETER TenantId
    Optional tenant ID or verified domain used for the interactive Graph sign-in.
    When omitted, the value is read from the selected SmartM365 tenant profile.

.PARAMETER Publisher
    Optional Author/Publisher value written to Intune. When omitted, the script
    uses the account connected through interactive Microsoft Graph authentication.

.PARAMETER Description
    Optional Intune package description. When omitted, the script uses the
    # Description: header from the detection and remediation scripts.

.PARAMETER UpdateExisting
    Updates an existing Intune remediation when a package with the same display
    name already exists. Without this switch, existing packages are skipped.

.PARAMETER AssignmentGroupId
    Optional Entra group IDs used to replace the package assignments through the
    Graph assign action. Omit this parameter to deploy packages without changing
    assignments.

.EXAMPLE
    .\SmartM365-Deploy-IntuneRemediation-CLI.ps1 -Path ..\Packages\WindowsUpdate\Cache-Health

.EXAMPLE
    .\SmartM365-Deploy-IntuneRemediation-CLI.ps1 -Path ..\Packages\WindowsUpdate -Recurse -UpdateExisting

.EXAMPLE
    .\SmartM365-Deploy-IntuneRemediation-CLI.ps1 -Path ..\Packages\WindowsUpdate\Cache-Health -AssignmentGroupId "00000000-0000-0000-0000-000000000000" -RunRemediationScript

.REQUIREMENTS
    PowerShell module:
    - Microsoft.Graph.Authentication

    Delegated Microsoft Graph permission:
    - DeviceManagementScripts.ReadWrite.All

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version: 1.0
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('PackagePath', 'LiteralPath')]
    [ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [string]$Tenant = 'test',

    [string]$TenantId = '',

    [ValidateSet('beta')]
    [string]$GraphApiVersion = 'beta',

    [string]$DisplayName = '',

    [string]$Publisher = '',

    [string]$Description = '',

    [ValidateNotNullOrEmpty()]
    [string]$Version = '1.0',

    [ValidateSet('system', 'user')]
    [string]$RunAsAccount = 'system',

    [switch]$RunAs32Bit,

    [switch]$EnforceSignatureCheck,

    [string[]]$RoleScopeTagIds = @('0'),

    [switch]$UpdateExisting,

    [switch]$Recurse,

    [string[]]$AssignmentGroupId = @(),

    [switch]$RunRemediationScript,

    [ValidateSet('Daily', 'Hourly')]
    [string]$AssignmentSchedule = 'Daily',

    [ValidateRange(1, 30)]
    [int]$AssignmentInterval = 1,

    [ValidatePattern('^\d{2}:\d{2}(:\d{2})?$')]
    [string]$AssignmentTime = '09:00:00',

    [switch]$AssignmentUseLocalTime
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $script:InputPaths = New-Object System.Collections.Generic.List[string]
    $script:RequiredScopes = @('DeviceManagementScripts.ReadWrite.All')

    $tenantContextPath = & {
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
    . $tenantContextPath
    Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

    function Get-SmartM365ConfigString {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$Name)

        if ($null -eq $script:SmartM365GlobalConfig) { return '' }
        $property = $script:SmartM365GlobalConfig.PSObject.Properties[$Name]
        if ($null -eq $property -or $null -eq $property.Value) { return '' }
        return [string]$property.Value
    }

    function Import-RequiredGraphModule {
        [CmdletBinding()]
        param()

        $moduleName = 'Microsoft.Graph.Authentication'
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

    function Connect-SmartM365InteractiveGraph {
        [CmdletBinding()]
        param(
            [AllowEmptyString()]
            [string]$TenantForConnection,

            [Parameter(Mandatory = $true)]
            [string[]]$Scopes
        )

        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

        $connectParams = @{
            Scopes      = $Scopes
            NoWelcome   = $true
            ErrorAction = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($TenantForConnection)) {
            $connectParams['TenantId'] = $TenantForConnection
        }

        Write-Output 'Connecting to Microsoft Graph with interactive delegated authentication...'
        Connect-MgGraph @connectParams | Out-Null

        $context = Get-MgContext
        if ($null -eq $context) {
            throw 'Microsoft Graph connection failed. Get-MgContext returned no context.'
        }

        Write-Output ("Connected to tenant: {0}" -f $context.TenantId)
        if (-not [string]::IsNullOrWhiteSpace($context.Account)) {
            Write-Output ("Connected account: {0}" -f $context.Account)
        }

        return $context
    }

    function Get-ObjectValue {
        [CmdletBinding()]
        param(
            [AllowNull()]
            $InputObject,

            [Parameter(Mandatory = $true)]
            [string]$Name
        )

        if ($null -eq $InputObject) { return $null }
        if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject[$Name] }

        $property = $InputObject.PSObject.Properties[$Name]
        if ($null -ne $property) { return $property.Value }
        return $null
    }

    function ConvertTo-GraphBinaryString {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$LiteralPath)

        $bytes = [System.IO.File]::ReadAllBytes($LiteralPath)
        return [Convert]::ToBase64String($bytes)
    }

    function ConvertTo-PackageDisplayName {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$DetectionScriptPath)

        $name = [System.IO.Path]::GetFileNameWithoutExtension($DetectionScriptPath)
        $name = $name -replace '-DetectionOnly$', ''
        $name = $name -replace '-Detection$', ''
        return $name
    }

    function Get-ScriptHelpDescription {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$LiteralPath)

        $headerLineCount = 0
        foreach ($line in [System.IO.File]::ReadLines($LiteralPath)) {
            $headerLineCount++
            if ($headerLineCount -gt 80) { break }

            if ($line -match '^\s*#\s*Description\s*:\s*(?<Text>.+?)\s*$') {
                return ($Matches['Text'] -replace '\s+', ' ').Trim()
            }
        }

        return ''
    }

    function Get-PackageDescription {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$DetectionScriptPath,

            [AllowNull()]
            [string]$RemediationScriptPath
        )

        $detectionDescription = Get-ScriptHelpDescription -LiteralPath $DetectionScriptPath
        $remediationDescription = ''
        if (-not [string]::IsNullOrWhiteSpace($RemediationScriptPath)) {
            $remediationDescription = Get-ScriptHelpDescription -LiteralPath $RemediationScriptPath
        }

        if (-not [string]::IsNullOrWhiteSpace($detectionDescription) -and
            -not [string]::IsNullOrWhiteSpace($remediationDescription) -and
            $detectionDescription -ne $remediationDescription) {
            return ("Detection: {0} Remediation: {1}" -f $detectionDescription, $remediationDescription)
        }

        if (-not [string]::IsNullOrWhiteSpace($detectionDescription)) {
            return $detectionDescription
        }

        if (-not [string]::IsNullOrWhiteSpace($remediationDescription)) {
            return $remediationDescription
        }

        return ''
    }

    function New-IntuneRemediationPackage {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$DetectionScriptPath)

        $detectionName = [System.IO.Path]::GetFileName($DetectionScriptPath)
        if ($detectionName -notmatch '^SmartM365-.+-(Detection|DetectionOnly)\.ps1$') {
            throw "Detection script name must follow SmartM365-*-Detection.ps1 or SmartM365-*-DetectionOnly.ps1: $DetectionScriptPath"
        }

        $isDetectionOnly = $detectionName -match '-DetectionOnly\.ps1$'
        $remediationPath = $null

        if (-not $isDetectionOnly) {
            $candidate = Join-Path -Path (Split-Path -Path $DetectionScriptPath -Parent) -ChildPath ($detectionName -replace '-Detection\.ps1$', '-Remediation.ps1')
            if (Test-Path -LiteralPath $candidate) {
                $remediationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($candidate)
            }
            else {
                Write-Warning "No remediation script found for '$detectionName'. The package will be deployed as detection-only."
            }
        }

        return [pscustomobject]@{
            DisplayName     = ConvertTo-PackageDisplayName -DetectionScriptPath $DetectionScriptPath
            Description     = Get-PackageDescription -DetectionScriptPath $DetectionScriptPath -RemediationScriptPath $remediationPath
            DetectionPath   = $DetectionScriptPath
            RemediationPath = $remediationPath
            DetectionOnly   = ($isDetectionOnly -or [string]::IsNullOrWhiteSpace($remediationPath))
        }
    }

    function Get-DetectionScriptsFromPath {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string]$InputPath)

        $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InputPath)
        if (-not (Test-Path -LiteralPath $fullPath)) {
            throw "Path not found: $InputPath"
        }

        $item = Get-Item -LiteralPath $fullPath -ErrorAction Stop
        if ($item.PSIsContainer) {
            $searchOption = if ($Recurse) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }
            return [System.IO.Directory]::EnumerateFiles($item.FullName, 'SmartM365-*-Detection*.ps1', $searchOption) |
                Where-Object { [System.IO.Path]::GetFileName($_) -match '^SmartM365-.+-(Detection|DetectionOnly)\.ps1$' } |
                Sort-Object
        }

        $fileName = $item.Name
        if ($fileName -match '^SmartM365-.+-(Detection|DetectionOnly)\.ps1$') {
            return @($item.FullName)
        }

        if ($fileName -match '^SmartM365-.+-Remediation\.ps1$') {
            $candidate = Join-Path -Path $item.DirectoryName -ChildPath ($fileName -replace '-Remediation\.ps1$', '-Detection.ps1')
            if (-not (Test-Path -LiteralPath $candidate)) {
                throw "Remediation script '$($item.FullName)' has no matching detection script."
            }
            return @($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($candidate))
        }

        throw "Unsupported path. Provide a SmartM365 detection script, remediation script, or package folder: $InputPath"
    }

    function Get-IntuneRemediationPackages {
        [CmdletBinding()]
        param([Parameter(Mandatory = $true)][string[]]$InputPath)

        $packagesByDetectionPath = [ordered]@{}
        foreach ($singlePath in $InputPath) {
            foreach ($detectionScript in Get-DetectionScriptsFromPath -InputPath $singlePath) {
                $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($detectionScript)
                $key = $resolvedPath.ToLowerInvariant()
                if (-not $packagesByDetectionPath.Contains($key)) {
                    $packagesByDetectionPath[$key] = New-IntuneRemediationPackage -DetectionScriptPath $resolvedPath
                }
            }
        }

        return @($packagesByDetectionPath.Values)
    }

    function Invoke-GraphJsonRequest {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('POST', 'PATCH')]
            [string]$Method,

            [Parameter(Mandatory = $true)]
            [string]$Uri,

            [Parameter(Mandatory = $true)]
            $Body
        )

        $json = $Body | ConvertTo-Json -Depth 20
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $json -ContentType 'application/json' -ErrorAction Stop
    }

    function Find-DeviceHealthScriptByDisplayName {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)][string]$GraphBaseUri,
            [Parameter(Mandatory = $true)][string]$Name
        )

        $escapedName = $Name.Replace("'", "''")
        $filter = [System.Uri]::EscapeDataString("displayName eq '$escapedName'")
        $uri = "$GraphBaseUri/deviceManagement/deviceHealthScripts?`$filter=$filter"
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        return @((Get-ObjectValue -InputObject $result -Name 'value'))
    }

    function New-DeviceHealthScriptBody {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]$Package,
            [Parameter(Mandatory = $true)][string]$PackageDisplayName
        )

        $body = [ordered]@{
            '@odata.type'             = '#microsoft.graph.deviceHealthScript'
            publisher                 = $script:EffectivePublisher
            version                   = $Version
            displayName               = $PackageDisplayName
            description               = $Package.Description
            detectionScriptContent    = ConvertTo-GraphBinaryString -LiteralPath $Package.DetectionPath
            remediationScriptContent  = ''
            runAsAccount             = $RunAsAccount
            enforceSignatureCheck     = [bool]$EnforceSignatureCheck
            runAs32Bit                = [bool]$RunAs32Bit
            roleScopeTagIds           = @($RoleScopeTagIds)
            deviceHealthScriptType    = 'deviceHealthScript'
        }

        if (-not [string]::IsNullOrWhiteSpace($Package.RemediationPath)) {
            $body['remediationScriptContent'] = ConvertTo-GraphBinaryString -LiteralPath $Package.RemediationPath
        }

        return $body
    }

    function New-AssignmentScheduleBody {
        [CmdletBinding()]
        param()

        if ($AssignmentSchedule -eq 'Hourly') {
            return [ordered]@{
                '@odata.type' = '#microsoft.graph.deviceHealthScriptHourlySchedule'
                interval      = $AssignmentInterval
            }
        }

        $timeValue = if ($AssignmentTime -match '^\d{2}:\d{2}$') { "$AssignmentTime`:00.0000000" } else { "$AssignmentTime.0000000" }
        return [ordered]@{
            '@odata.type' = '#microsoft.graph.deviceHealthScriptDailySchedule'
            interval      = $AssignmentInterval
            useUtc        = (-not [bool]$AssignmentUseLocalTime)
            time          = $timeValue
        }
    }

    function Set-DeviceHealthScriptAssignments {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            [Parameter(Mandatory = $true)][string]$GraphBaseUri,
            [Parameter(Mandatory = $true)][string]$ScriptId,
            [Parameter(Mandatory = $true)][string]$PackageDisplayName,
            [Parameter(Mandatory = $true)][string[]]$GroupIds,
            [Parameter(Mandatory = $true)][bool]$EnableRemediation
        )

        $assignments = @()
        foreach ($groupId in $GroupIds) {
            if ([string]::IsNullOrWhiteSpace($groupId)) { continue }
            $assignments += [ordered]@{
                '@odata.type'         = '#microsoft.graph.deviceHealthScriptAssignment'
                target                = [ordered]@{
                    '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                    groupId       = $groupId
                }
                runRemediationScript  = $EnableRemediation
                runSchedule           = New-AssignmentScheduleBody
            }
        }

        if ($assignments.Count -eq 0) { return }

        $body = [ordered]@{
            deviceHealthScriptAssignments = @($assignments)
        }
        $uri = "$GraphBaseUri/deviceManagement/deviceHealthScripts/$ScriptId/assign"

        if ($PSCmdlet.ShouldProcess($PackageDisplayName, 'Replace Intune remediation assignments')) {
            Invoke-GraphJsonRequest -Method POST -Uri $uri -Body $body | Out-Null
            Write-Output ("Assignments updated: {0}" -f $PackageDisplayName)
        }
    }

    function Publish-IntuneRemediationPackage {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            [Parameter(Mandatory = $true)]$Package,
            [Parameter(Mandatory = $true)][string]$GraphBaseUri,
            [Parameter(Mandatory = $true)][string]$PackageDisplayName
        )

        $existing = Find-DeviceHealthScriptByDisplayName -GraphBaseUri $GraphBaseUri -Name $PackageDisplayName
        if ($existing.Count -gt 1) {
            throw "More than one Intune remediation uses displayName '$PackageDisplayName'. Rename duplicates before deploying."
        }

        $body = New-DeviceHealthScriptBody -Package $Package -PackageDisplayName $PackageDisplayName
        $scriptId = $null

        if ($existing.Count -eq 1) {
            $scriptId = [string](Get-ObjectValue -InputObject $existing[0] -Name 'id')
            if (-not $UpdateExisting) {
                Write-Warning "Skipping existing Intune remediation '$PackageDisplayName'. Use -UpdateExisting to patch it."
                return $null
            }

            $uri = "$GraphBaseUri/deviceManagement/deviceHealthScripts/$scriptId"
            if ($PSCmdlet.ShouldProcess($PackageDisplayName, 'Update Intune remediation')) {
                Invoke-GraphJsonRequest -Method PATCH -Uri $uri -Body $body | Out-Null
                Write-Output ("Updated: {0} ({1})" -f $PackageDisplayName, $scriptId)
            }
        }
        else {
            $uri = "$GraphBaseUri/deviceManagement/deviceHealthScripts"
            if ($PSCmdlet.ShouldProcess($PackageDisplayName, 'Create Intune remediation')) {
                $created = Invoke-GraphJsonRequest -Method POST -Uri $uri -Body $body
                $scriptId = [string](Get-ObjectValue -InputObject $created -Name 'id')
                Write-Output ("Created: {0} ({1})" -f $PackageDisplayName, $scriptId)
            }
        }

        return [pscustomobject]@{
            Id              = $scriptId
            DisplayName     = $PackageDisplayName
            DetectionPath   = $Package.DetectionPath
            RemediationPath = $Package.RemediationPath
            DetectionOnly   = $Package.DetectionOnly
        }
    }
}

process {
    foreach ($item in $Path) {
        [void]$script:InputPaths.Add($item)
    }
}

end {
    Import-RequiredGraphModule

    $tenantForConnection = if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $TenantId
    }
    elseif (-not [string]::IsNullOrWhiteSpace((Get-SmartM365ConfigString -Name 'TenantId'))) {
        Get-SmartM365ConfigString -Name 'TenantId'
    }
    else {
        Get-SmartM365ConfigString -Name 'OrgDomain'
    }

    if ([string]::IsNullOrWhiteSpace($tenantForConnection)) {
        throw 'TenantId or OrgDomain was not found. Provide -TenantId or update the selected SmartM365 tenant profile.'
    }

    $packages = Get-IntuneRemediationPackages -InputPath @($script:InputPaths)
    if ($packages.Count -eq 0) {
        throw 'No SmartM365 Intune remediation detection scripts were found.'
    }

    if (-not [string]::IsNullOrWhiteSpace($DisplayName) -and $packages.Count -ne 1) {
        throw '-DisplayName can only be used when deploying exactly one package.'
    }

    $graphContext = Connect-SmartM365InteractiveGraph -TenantForConnection $tenantForConnection -Scopes $script:RequiredScopes
    $script:EffectivePublisher = if (-not [string]::IsNullOrWhiteSpace($Publisher)) {
        $Publisher
    }
    elseif ($null -ne $graphContext -and -not [string]::IsNullOrWhiteSpace($graphContext.Account)) {
        [string]$graphContext.Account
    }
    else {
        'Interactive Graph account'
    }
    Write-Output ("Intune Author/Publisher: {0}" -f $script:EffectivePublisher)

    $graphBaseUri = "https://graph.microsoft.com/$GraphApiVersion"
    $published = @()

    foreach ($package in $packages) {
        $packageDisplayName = if ([string]::IsNullOrWhiteSpace($DisplayName)) { $package.DisplayName } else { $DisplayName }
        Write-Output ("Processing: {0}" -f $packageDisplayName)
        if (-not [string]::IsNullOrWhiteSpace($Description)) {
            $package.Description = $Description
        }
        elseif ([string]::IsNullOrWhiteSpace($package.Description)) {
            $package.Description = $packageDisplayName
        }

        Write-Output ("Description: {0}" -f $package.Description)
        Write-Output ("Detection: {0}" -f $package.DetectionPath)
        if (-not [string]::IsNullOrWhiteSpace($package.RemediationPath)) {
            Write-Output ("Remediation: {0}" -f $package.RemediationPath)
        }

        $result = Publish-IntuneRemediationPackage -Package $package -GraphBaseUri $graphBaseUri -PackageDisplayName $packageDisplayName
        if ($null -ne $result) {
            $published += $result

            if ($AssignmentGroupId.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($result.Id)) {
                $enableRemediation = ([bool]$RunRemediationScript -and -not [bool]$result.DetectionOnly)
                if ($RunRemediationScript -and $result.DetectionOnly) {
                    Write-Warning "RunRemediationScript was requested, but '$packageDisplayName' has no remediation script. Assignment will run detection only."
                }

                Set-DeviceHealthScriptAssignments `
                    -GraphBaseUri $graphBaseUri `
                    -ScriptId $result.Id `
                    -PackageDisplayName $packageDisplayName `
                    -GroupIds $AssignmentGroupId `
                    -EnableRemediation $enableRemediation
            }
        }
    }

    Write-Output ''
    Write-Output 'Deployment completed.'
    Write-Output ("Packages found: {0}" -f $packages.Count)
    Write-Output ("Packages created or updated: {0}" -f $published.Count)

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
}
