#Requires -Version 7.0
<#
.SYNOPSIS
    Core helpers for Smart IntunePackager: logging, local configuration, app
    definition profiles, MSI and PSADT discovery, .intunewin metadata and
    Microsoft Graph payload construction.

.DESCRIPTION
    This module performs no network call. Every function is usable from the GUI,
    from background runspaces and from the CLI. Payload builders return ordered
    hashtables that are serialized as-is to Microsoft Graph.

.VERSION
    1.0.6

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:IpLogState = @{
    Path    = ''
    Sink    = $null
    Console = $true
}

$script:IpRunAsAccounts = @('system', 'user')
$script:IpRestartBehaviors = @('allow', 'basedOnReturnCode', 'suppress', 'force')
$script:IpArchitectures = @('x64', 'arm64', 'x86')
$script:IpReturnCodeTypes = @('success', 'softReboot', 'hardReboot', 'retry', 'failed')
# 'Preserved' keeps the detection rules of an imported app unchanged (mixed or unsupported rule types).
$script:IpDetectionTypes = @('MsiProductCode', 'File', 'Registry', 'Script', 'Preserved')
$script:IpDetectionOperators = @('notConfigured', 'equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual')
$script:IpFileDetectionTypes = @('exists', 'doesNotExist', 'modifiedDate', 'createdDate', 'version', 'sizeInMB')
$script:IpRegistryDetectionTypes = @('exists', 'doesNotExist', 'string', 'integer', 'version')
$script:IpPublishModes = @('Create', 'UpdateExisting')
# Values sent as minimumSupportedWindowsRelease. Windows 11 values follow the documented
# 'Windows11_23H2' format. Windows 10 releases up to 21H1 use the bare release name ('2H20' is 20H2);
# 21H2 and 22H2 use 'Windows10_<release>' (observed on a tenant: 'Windows10_22H2').
$script:IpWindowsReleases = @(
    '1607', '1703', '1709', '1803', '1809', '1903', '1909', '2004', '2H20', '21H1', 'Windows10_21H2', 'Windows10_22H2',
    'Windows11_21H2', 'Windows11_22H2', 'Windows11_23H2', 'Windows11_24H2'
)
# Values written by Smart IntunePackager 1.0.x profiles before the Windows 10 21H2/22H2 correction.
$script:IpLegacyWindowsReleases = @{ '21H2' = 'Windows10_21H2'; '22H2' = 'Windows10_22H2' }
$script:IpProfileSchemaVersion = 1

#region Logging

function Format-IpLogLines {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [AllowEmptyString()][string]$Message = '',
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO',
        [datetime]$Timestamp = (Get-Date)
    )

    $stamp = $Timestamp.ToString('yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    $prefix = if ($Level -eq 'INFO') { $stamp } else { '{0} [{1}]' -f $stamp, $Level }
    return @([regex]::Split([string]$Message, '\r?\n') | ForEach-Object { '{0} {1}' -f $prefix, $_ })
}

function Set-IpLogSink {
    <#
    .SYNOPSIS
        Registers a scriptblock that receives every formatted log line (GUI log panel).
    #>
    [CmdletBinding()]
    param(
        [AllowNull()][scriptblock]$Sink,
        [bool]$Console = ($null -eq $Sink)
    )

    $script:IpLogState.Sink = $Sink
    $script:IpLogState.Console = $Console
}

function Set-IpLogPath {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Path)
    $script:IpLogState.Path = $Path
}

function Get-IpLogPath {
    [CmdletBinding()]
    param()
    return [string]$script:IpLogState.Path
}

function Initialize-IpLog {
    <#
    .SYNOPSIS
        Creates the log directory, archives the previous log and keeps the newest archives.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [string]$BaseName = 'SmartM365-IntunePackager-GUI',
        [ValidateRange(1, 1000)][int]$MaxFiles = 10
    )

    [IO.Directory]::CreateDirectory($Directory) | Out-Null
    $logPath = Join-Path -Path $Directory -ChildPath ($BaseName + '.log')
    if (Test-Path -LiteralPath $logPath -PathType Leaf) {
        $current = Get-Item -LiteralPath $logPath
        if ($current.Length -gt 0) {
            $archive = Join-Path -Path $Directory -ChildPath ('{0}-{1}.log' -f $BaseName, (Get-Date -Format 'yyyyMMdd-HHmmss'))
            if (Test-Path -LiteralPath $archive) {
                $archive = Join-Path -Path $Directory -ChildPath ('{0}-{1}-{2}.log' -f $BaseName, (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8)))
            }
            Move-Item -LiteralPath $logPath -Destination $archive
        }
    }

    $archives = @(Get-ChildItem -LiteralPath $Directory -Filter ($BaseName + '-*.log') -File | Sort-Object LastWriteTime -Descending)
    if ($archives.Count -gt $MaxFiles) {
        $archives | Select-Object -Skip $MaxFiles | Remove-Item -Force -ErrorAction SilentlyContinue
    }

    Set-IpLogPath -Path $logPath
    return $logPath
}

function Write-IpLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO',
        # Writes only to the log file (technical details such as line numbers and stack traces).
        [switch]$FileOnly
    )

    $lines = Format-IpLogLines -Message $Message -Level $Level
    $path = [string]$script:IpLogState.Path
    if (-not [string]::IsNullOrWhiteSpace($path)) {
        $text = ($lines -join "`r`n") + "`r`n"
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                [IO.File]::AppendAllText($path, $text, [Text.UTF8Encoding]::new($false))
                break
            }
            catch [IO.IOException] {
                # Another runspace may be appending at the same moment.
                Start-Sleep -Milliseconds (50 * $attempt)
            }
            catch {
                Write-Verbose ("Unable to write log file '{0}': {1}" -f $path, $_.Exception.Message)
                break
            }
        }
    }

    if ($FileOnly) { return }
    if ($null -ne $script:IpLogState.Sink) {
        foreach ($line in $lines) {
            try { & $script:IpLogState.Sink $line $Level } catch { Write-Verbose $_.Exception.Message }
        }
    }
    if ($script:IpLogState.Console) {
        foreach ($line in $lines) { Microsoft.PowerShell.Utility\Write-Host $line }
    }
}

function Format-IpByteSize {
    <#
    .SYNOPSIS
        Formats a byte count with an adapted unit (B, KB, MB, GB, TB), invariant culture.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)][double]$Bytes)

    $units = @('B', 'KB', 'MB', 'GB', 'TB')
    $value = [Math]::Abs($Bytes)
    $index = 0
    while ($value -ge 1024 -and $index -lt $units.Count - 1) { $value /= 1024; $index++ }
    if ($index -eq 0) { return ('{0} B' -f [int64]$Bytes) }
    if ($Bytes -lt 0) { $value = -$value }
    return [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:0.0} {1}', $value, $units[$index])
}

function Format-IpLocalDateTime {
    <#
    .SYNOPSIS
        Converts a Graph timestamp (ISO string, DateTime or DateTimeOffset) to local time, 'yyyy-MM-dd HH:mm' by default.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Value, [string]$Format = 'yyyy-MM-dd HH:mm')

    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { return '' }
    $culture = [Globalization.CultureInfo]::InvariantCulture
    if ($Value -is [datetimeoffset]) { return $Value.LocalDateTime.ToString($Format, $culture) }
    if ($Value -is [datetime]) {
        # Graph timestamps are UTC; an unspecified kind comes from a UTC value without its marker.
        $utc = switch ($Value.Kind) { 'Local' { $Value.ToUniversalTime() } 'Unspecified' { [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) } default { $Value } }
        return $utc.ToLocalTime().ToString($Format, $culture)
    }
    $parsed = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse([string]$Value, $culture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { return $parsed.LocalDateTime.ToString($Format, $culture) }
    return [string]$Value
}

#endregion

#region JSON and configuration

function Copy-IpJsonValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value -or $Value -is [string] -or $Value -is [ValueType]) { return $Value }
    return (ConvertTo-Json -InputObject $Value -Depth 30 | ConvertFrom-Json -NoEnumerate)
}

function Merge-IpObjectFromTemplate {
    <#
    .SYNOPSIS
        Adds properties missing from Target using Template, recursively, without
        overwriting existing values. Returns the dotted paths of added keys.

    .DESCRIPTION
        Properties named in DictionaryProperties hold user-keyed entries (for example
        tenant profiles). Their first template entry is used as the schema for every
        existing entry; template entries are only copied when the target has none.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)][psobject]$Target,
        [Parameter(Mandatory = $true)][psobject]$Template,
        [string]$Path = '',
        [string[]]$DictionaryProperties = @('Tenants')
    )

    $added = [Collections.Generic.List[string]]::new()
    foreach ($property in $Template.PSObject.Properties) {
        $name = $property.Name
        $childPath = if ($Path) { '{0}.{1}' -f $Path, $name } else { $name }
        $existing = $Target.PSObject.Properties[$name]
        if ($null -eq $existing) {
            $Target | Add-Member -NotePropertyName $name -NotePropertyValue (Copy-IpJsonValue $property.Value)
            $added.Add($childPath)
            continue
        }

        $templateValue = $property.Value
        $targetValue = $existing.Value
        if ($templateValue -isnot [Management.Automation.PSCustomObject] -or $targetValue -isnot [Management.Automation.PSCustomObject]) {
            continue
        }

        if ($name -in $DictionaryProperties) {
            $entries = @($targetValue.PSObject.Properties)
            $prototypes = @($templateValue.PSObject.Properties)
            if ($entries.Count -eq 0) {
                foreach ($prototype in $prototypes) {
                    $targetValue | Add-Member -NotePropertyName $prototype.Name -NotePropertyValue (Copy-IpJsonValue $prototype.Value)
                    $added.Add(('{0}.{1}' -f $childPath, $prototype.Name))
                }
                continue
            }
            if ($prototypes.Count -eq 0 -or $prototypes[0].Value -isnot [Management.Automation.PSCustomObject]) { continue }
            foreach ($entry in $entries) {
                if ($entry.Value -is [Management.Automation.PSCustomObject]) {
                    foreach ($key in (Merge-IpObjectFromTemplate -Target $entry.Value -Template $prototypes[0].Value -Path ('{0}.{1}' -f $childPath, $entry.Name) -DictionaryProperties $DictionaryProperties)) {
                        $added.Add($key)
                    }
                }
            }
            continue
        }

        foreach ($key in (Merge-IpObjectFromTemplate -Target $targetValue -Template $templateValue -Path $childPath -DictionaryProperties $DictionaryProperties)) {
            $added.Add($key)
        }
    }

    return [string[]]$added.ToArray()
}

function Save-IpJsonFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$InputObject
    )

    $directory = Split-Path -Path $Path -Parent
    if ($directory) { [IO.Directory]::CreateDirectory($directory) | Out-Null }
    $json = ConvertTo-Json -InputObject $InputObject -Depth 30
    if ($PSCmdlet.ShouldProcess($Path, 'Write JSON file')) {
        [IO.File]::WriteAllText($Path, $json + "`r`n", [Text.UTF8Encoding]::new($false))
    }
}

function Read-IpJsonFile {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    try {
        return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        throw ("Invalid JSON file '{0}': {1}" -f $Path, $_.Exception.Message)
    }
}

function Import-IpConfiguration {
    <#
    .SYNOPSIS
        Loads the local configuration, creating it from the committed template when
        missing and adding any key missing from the template without overwriting
        local values. Added keys are logged and the file is saved.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [Parameter(Mandatory = $true)][string]$TemplatePath
    )

    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
        throw "Configuration template not found: $TemplatePath"
    }
    $template = Read-IpJsonFile -Path $TemplatePath

    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        Copy-Item -LiteralPath $TemplatePath -Destination $ConfigPath -ErrorAction Stop
        Write-IpLog "Created local configuration from template: $ConfigPath"
        return (Read-IpJsonFile -Path $ConfigPath)
    }

    $config = Read-IpJsonFile -Path $ConfigPath
    $added = @(Merge-IpObjectFromTemplate -Target $config -Template $template)
    if ($added.Count -gt 0) {
        Save-IpJsonFile -Path $ConfigPath -InputObject $config -Confirm:$false
        Write-IpLog ("Configuration keys added from template: {0}" -f ($added -join ', '))
    }
    return $config
}

function Resolve-IpPath {
    <#
    .SYNOPSIS
        Expands {{AppRoot}} and environment variables, and roots relative paths on BaseDirectory.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowEmptyString()][AllowNull()][string]$Path,
        [Parameter(Mandatory = $true)][string]$BaseDirectory
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $expanded = [Environment]::ExpandEnvironmentVariables($Path.Replace('{{AppRoot}}', $BaseDirectory))
    if (-not [IO.Path]::IsPathRooted($expanded)) { $expanded = Join-Path -Path $BaseDirectory -ChildPath $expanded }
    return [IO.Path]::GetFullPath($expanded)
}

function Get-IpTenantProfile {
    <#
    .SYNOPSIS
        Returns the validated tenant profile selected by -Tenant.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$Config,
        [Parameter(Mandatory = $true)][string]$Tenant
    )

    $tenants = $Config.PSObject.Properties['Tenants']
    if ($null -eq $tenants -or $null -eq $tenants.Value) { throw 'Configuration error: the Tenants section is missing.' }
    $entry = $tenants.Value.PSObject.Properties[$Tenant]
    if ($null -eq $entry) {
        $known = @($tenants.Value.PSObject.Properties.Name) -join ', '
        throw ("Configuration error: tenant profile '{0}' was not found. Known profiles: {1}" -f $Tenant, $known)
    }

    $value = $entry.Value
    $read = { param($name) $property = $value.PSObject.Properties[$name]; if ($null -eq $property -or $null -eq $property.Value) { '' } else { ([string]$property.Value).Trim() } }
    $tenantId = & $read 'TenantId'
    $clientId = & $read 'ClientId'
    $thumbprint = (& $read 'CertificateThumbprint').Replace(' ', '').ToUpperInvariant()
    $authMode = & $read 'AuthMode'
    $signIn = & $read 'DelegatedSignIn'
    if (-not $signIn) { $signIn = 'Auto' }

    if ($authMode -notin @('Delegated', 'Certificate')) {
        throw ("Configuration error: tenant '{0}' AuthMode must be Delegated or Certificate." -f $Tenant)
    }
    if ($signIn -notin @('Auto', 'Wam', 'Browser', 'DeviceCode')) {
        throw ("Configuration error: tenant '{0}' DelegatedSignIn must be Auto, Wam, Browser or DeviceCode." -f $Tenant)
    }
    if ($authMode -eq 'Certificate') {
        foreach ($required in @(@('TenantId', $tenantId), @('ClientId', $clientId), @('CertificateThumbprint', $thumbprint))) {
            if (-not $required[1]) { throw ("Configuration error: tenant '{0}' uses certificate authentication and requires {1}." -f $Tenant, $required[0]) }
        }
    }
    if ($authMode -eq 'Delegated' -and $signIn -eq 'Browser' -and -not $clientId) {
        throw ("Configuration error: tenant '{0}' uses DelegatedSignIn=Browser, which requires the ClientId of your own app registration." -f $Tenant)
    }

    $tenantKind = ''
    if ($tenantId) {
        $tenantIssue = Test-IpTenantIdentifier -Value $tenantId
        if ($tenantIssue) { throw ("Configuration error: tenant '{0}' TenantId '{1}' is invalid: {2}" -f $Tenant, $tenantId, $tenantIssue) }
        $tenantKind = if (Test-IpGuid $tenantId) { 'Guid' } else { 'Domain' }
    }
    if ($clientId) {
        $clientIssue = Test-IpIdentifierGuid -Value $clientId
        if ($clientIssue) { throw ("Configuration error: tenant '{0}' ClientId '{1}' is invalid: {2}" -f $Tenant, $clientId, $clientIssue) }
    }
    if ($thumbprint -and $thumbprint -notmatch '^[0-9A-F]{40}$') {
        throw ("Configuration error: tenant '{0}' CertificateThumbprint must contain 40 hexadecimal characters." -f $Tenant)
    }

    return [pscustomobject]@{
        Name                  = $Tenant
        TenantId              = $tenantId
        TenantIdKind          = $tenantKind
        AuthMode              = $authMode
        DelegatedSignIn       = $signIn
        ClientId              = $clientId
        CertificateThumbprint = $thumbprint
    }
}

function Test-IpIdentifierGuid {
    <#
    .SYNOPSIS
        Returns why a value is not a usable GUID identifier, or an empty string when it is.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string]$Value)

    if ($Value -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { return 'not a GUID in the form 00000000-0000-0000-0000-000000000000.' }
    # Placeholders such as the empty GUID or 11111111-1111-... use at most two distinct digits.
    if (@($Value.Replace('-', '').ToLowerInvariant().ToCharArray() | Select-Object -Unique).Count -le 2) { return 'placeholder GUID.' }
    return ''
}

function Test-IpTenantIdentifier {
    <#
    .SYNOPSIS
        Returns why a TenantId is unusable (not a GUID or a domain, or a placeholder), or an empty string.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string]$Value)

    $text = ([string]$Value).Trim()
    if (-not $text) { return 'empty value.' }
    if ($text -match '[<>{}\[\]\s]' -or $text -match '^(x+|\?+|tenant[-_ ]?id|your[-_ ].*|changeme|todo|tbd|placeholder|none|null)$') { return 'placeholder value.' }
    if ($text -match '^[0-9a-fA-F-]{36}$') { $issue = Test-IpIdentifierGuid -Value $text; if ($issue) { return $issue } else { return '' } }
    if ($text -notmatch '^(?=.{4,253}$)([A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,63}$') { return 'expected a tenant GUID or a domain such as tenant.onmicrosoft.com.' }
    $labels = $text.ToLowerInvariant().Split('.')
    if ($labels[-1] -in @('example', 'test', 'invalid', 'localhost') -or @($labels | Where-Object { $_ -in @('contoso', 'fabrikam', 'northwindtraders', 'adventure-works', 'example', 'yourtenant', 'tenant') }).Count -gt 0) {
        return 'fictional or documentation domain.'
    }
    return ''
}

#endregion

#region App definition profile

function Get-IpDefaultReturnCodes {
    [CmdletBinding()]
    param()
    return @(
        [pscustomobject]@{ ReturnCode = 0; Type = 'success' }
        [pscustomobject]@{ ReturnCode = 1707; Type = 'success' }
        [pscustomobject]@{ ReturnCode = 3010; Type = 'softReboot' }
        [pscustomobject]@{ ReturnCode = 1641; Type = 'hardReboot' }
        [pscustomobject]@{ ReturnCode = 1618; Type = 'retry' }
    )
}

function Get-IpObjectProperty {
    <#
    .SYNOPSIS
        Returns a property value, or $null when the object or the property is missing (StrictMode-safe).
    #>
    param($InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [Collections.IDictionary]) { if ($InputObject.Contains($Name)) { return $InputObject[$Name] }; return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Test-IpWindowsRelease {
    <#
    .SYNOPSIS
        True for a listed release or a Graph release format (imported apps can carry newer values).
    #>
    param([AllowEmptyString()][string]$Value)
    if ($Value -in $script:IpWindowsReleases) { return $true }
    return ($Value -match '^(\d{4}|\dH\d\d|\d\dH\d|Windows1[01]_\d\dH\d)$')
}

function Get-IpWindowsReleases {
    [CmdletBinding()]
    param()
    return @($script:IpWindowsReleases)
}

function New-IpAppDefinition {
    <#
    .SYNOPSIS
        Returns an empty app definition with every key and default value.
    #>
    [CmdletBinding()]
    param()

    $definition = [ordered]@{
        SchemaVersion = $script:IpProfileSchemaVersion
        Package       = [ordered]@{ SourceFolder = ''; SetupFile = ''; OutputFolder = '' }
        App           = [ordered]@{
            DisplayName = ''; Description = ''; Publisher = ''; Version = ''; Notes = ''
            InformationUrl = ''; PrivacyInformationUrl = ''; Developer = ''; Owner = ''; LogoPath = ''
        }
        Program       = [ordered]@{
            InstallCommandLine = ''; UninstallCommandLine = ''; RunAsAccount = 'system'
            DeviceRestartBehavior = 'suppress'; MaxRunTimeInMinutes = 60
        }
        ReturnCodes   = @(Get-IpDefaultReturnCodes)
        Requirements  = [ordered]@{ Architectures = @('x64'); MinimumWindowsRelease = '1607'; MinimumFreeDiskSpaceInMB = $null }
        Detection     = [ordered]@{
            Type     = 'MsiProductCode'
            Msi      = [ordered]@{ ProductCode = ''; ProductVersionOperator = 'notConfigured'; ProductVersion = '' }
            File     = @()
            Registry = @()
            Script   = [ordered]@{ ScriptPath = ''; EnforceSignatureCheck = $false; RunAs32Bit = $false }
        }
        Publish       = [ordered]@{ Mode = 'Create'; ExistingAppId = '' }
        # Set by Import definition from Intune; LastModifiedDateTime guards against concurrent portal edits.
        ImportedFrom  = [ordered]@{ AppId = ''; DisplayName = ''; ContentVersion = ''; LastModifiedDateTime = ''; ImportedAt = ''; SyncedAt = '' }
        # Informational list of what Intune keeps and this tool does not edit (never sent).
        Preserved     = @()
    }
    return (ConvertTo-Json -InputObject $definition -Depth 10 | ConvertFrom-Json)
}

function Import-IpAppProfile {
    <#
    .SYNOPSIS
        Loads a JSON app definition profile and adds any missing key with its default.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    $definition = Read-IpJsonFile -Path $Path
    $schema = $definition.PSObject.Properties['SchemaVersion']
    if ($null -ne $schema -and [int]$schema.Value -gt $script:IpProfileSchemaVersion) {
        throw ("Profile schema version {0} is newer than supported version {1}: {2}" -f $schema.Value, $script:IpProfileSchemaVersion, $Path)
    }
    [void](Merge-IpObjectFromTemplate -Target $definition -Template (New-IpAppDefinition) -DictionaryProperties @())
    $release = [string]$definition.Requirements.MinimumWindowsRelease
    if ($script:IpLegacyWindowsReleases.ContainsKey($release)) {
        $definition.Requirements.MinimumWindowsRelease = $script:IpLegacyWindowsReleases[$release]
        Write-IpLog ("Profile minimum Windows release '{0}' converted to '{1}' (Graph value for Windows 10 {0})." -f $release, $definition.Requirements.MinimumWindowsRelease) -Level WARN
    }
    return $definition
}

function Export-IpAppProfile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][psobject]$Definition,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $Definition.SchemaVersion = $script:IpProfileSchemaVersion
    Save-IpJsonFile -Path $Path -InputObject $Definition -WhatIf:$WhatIfPreference -Confirm:$false
}

function Save-IpProfileImportedFrom {
    <#
    .SYNOPSIS
        Writes only the ImportedFrom section into an existing profile file; the rest of the file is
        kept as it is (command line overrides such as -PublishMode are not persisted). The file is only
        changed when it was imported from the same app; returns $true when it was written.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)]$ImportedFrom)

    $stored = Read-IpJsonFile -Path $Path
    $storedAppId = [string](Get-IpObjectProperty (Get-IpObjectProperty $stored 'ImportedFrom') 'AppId')
    if (-not $storedAppId -or $storedAppId -ne [string](Get-IpObjectProperty $ImportedFrom 'AppId')) { return $false }
    $copy = $ImportedFrom | ConvertTo-Json -Depth 5 | ConvertFrom-Json
    $stored | Add-Member -NotePropertyName ImportedFrom -NotePropertyValue $copy -Force
    if (-not $PSCmdlet.ShouldProcess($Path, 'Update ImportedFrom')) { return $false }
    Save-IpJsonFile -Path $Path -InputObject $stored -Confirm:$false
    return $true
}

function Test-IpGuid {
    param([AllowNull()][string]$Value)
    $parsed = [guid]::Empty
    return [guid]::TryParse([string]$Value, [ref]$parsed)
}

function Test-IpAppDefinition {
    <#
    .SYNOPSIS
        Returns validation errors for an app definition. An empty result means valid.

    .PARAMETER BaseDirectory
        Base for relative paths. When given, referenced files (setup, logo, script) must exist.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)][psobject]$Definition,
        [string]$BaseDirectory = '',
        [switch]$ForPublish
    )

    $errors = [Collections.Generic.List[string]]::new()
    $app = $Definition.App
    $program = $Definition.Program

    if ([string]::IsNullOrWhiteSpace($app.DisplayName)) { $errors.Add('App.DisplayName is required.') }
    if ([string]::IsNullOrWhiteSpace($app.Publisher)) { $errors.Add('App.Publisher is required.') }
    foreach ($urlName in @('InformationUrl', 'PrivacyInformationUrl')) {
        $url = [string]$app.$urlName
        if ($url) {
            $uri = $null
            if (-not [Uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -notin @('http', 'https')) {
                $errors.Add("App.$urlName must be an absolute http or https URL.")
            }
        }
    }
    if ($app.LogoPath -and ([IO.Path]::GetExtension([string]$app.LogoPath).ToLowerInvariant() -notin @('.png', '.jpg', '.jpeg'))) {
        $errors.Add('App.LogoPath must be a .png, .jpg or .jpeg file.')
    }

    if ([string]::IsNullOrWhiteSpace($program.InstallCommandLine)) { $errors.Add('Program.InstallCommandLine is required.') }
    if ([string]::IsNullOrWhiteSpace($program.UninstallCommandLine)) { $errors.Add('Program.UninstallCommandLine is required.') }
    if ($program.RunAsAccount -notin $script:IpRunAsAccounts) { $errors.Add('Program.RunAsAccount must be system or user.') }
    if ($program.DeviceRestartBehavior -notin $script:IpRestartBehaviors) {
        $errors.Add(('Program.DeviceRestartBehavior must be one of: {0}.' -f ($script:IpRestartBehaviors -join ', ')))
    }
    $maxRun = 0
    if (-not [int]::TryParse([string]$program.MaxRunTimeInMinutes, [ref]$maxRun) -or $maxRun -lt 1 -or $maxRun -gt 1440) {
        $errors.Add('Program.MaxRunTimeInMinutes must be between 1 and 1440.')
    }

    $codes = @($Definition.ReturnCodes)
    if ($codes.Count -eq 0) { $errors.Add('ReturnCodes must contain at least one entry.') }
    $seen = @{}
    foreach ($code in $codes) {
        $number = 0
        if (-not [int]::TryParse([string]$code.ReturnCode, [ref]$number)) { $errors.Add("Return code '$($code.ReturnCode)' is not an integer."); continue }
        if ($seen.ContainsKey($number)) { $errors.Add("Return code $number is defined more than once.") }
        $seen[$number] = $true
        if ($code.Type -notin $script:IpReturnCodeTypes) { $errors.Add(("Return code {0} type must be one of: {1}." -f $number, ($script:IpReturnCodeTypes -join ', '))) }
    }
    if (-not @($codes | Where-Object { $_.Type -eq 'success' })) { $errors.Add('ReturnCodes must contain at least one success code.') }

    $architectures = @($Definition.Requirements.Architectures | Where-Object { $_ })
    if ($architectures.Count -eq 0) { $errors.Add('Requirements.Architectures must contain at least one architecture.') }
    if ('NotRestricted' -in $architectures) {
        if ($architectures.Count -gt 1) { $errors.Add('Requirements.Architectures: NotRestricted cannot be combined with an architecture.') }
    }
    else {
        foreach ($architecture in $architectures) {
            if ($architecture -notin $script:IpArchitectures) { $errors.Add("Unsupported architecture '$architecture'. Use x64, arm64 or x86, or NotRestricted.") }
        }
    }
    if (-not (Test-IpWindowsRelease -Value ([string]$Definition.Requirements.MinimumWindowsRelease))) {
        $errors.Add(('Requirements.MinimumWindowsRelease must be one of: {0}.' -f ($script:IpWindowsReleases -join ', ')))
    }
    $disk = Get-IpObjectProperty $Definition.Requirements 'MinimumFreeDiskSpaceInMB'
    if ($null -ne $disk -and [string]$disk -ne '') {
        $diskValue = 0
        if (-not [int]::TryParse([string]$disk, [ref]$diskValue) -or $diskValue -lt 1) { $errors.Add('Requirements.MinimumFreeDiskSpaceInMB must be a positive number of megabytes, or empty.') }
    }
    if ($Definition.Detection.Type -eq 'Preserved' -and $Definition.Publish.Mode -ne 'UpdateExisting') {
        $errors.Add('Detection.Type Preserved keeps the rules of an existing app: it is only valid with Publish.Mode UpdateExisting.')
    }

    foreach ($detectionError in (Get-IpDetectionValidationErrors -Detection $Definition.Detection -BaseDirectory $BaseDirectory)) {
        $errors.Add($detectionError)
    }

    if ($Definition.Publish.Mode -notin $script:IpPublishModes) { $errors.Add('Publish.Mode must be Create or UpdateExisting.') }
    if ($Definition.Publish.Mode -eq 'UpdateExisting' -and -not (Test-IpGuid $Definition.Publish.ExistingAppId)) {
        $errors.Add('Publish.ExistingAppId must be a valid GUID when Publish.Mode is UpdateExisting.')
    }

    if ($BaseDirectory) {
        $source = Resolve-IpPath -Path $Definition.Package.SourceFolder -BaseDirectory $BaseDirectory
        if (-not $ForPublish) {
            if (-not $source -or -not (Test-Path -LiteralPath $source -PathType Container)) { $errors.Add("Package.SourceFolder not found: $source") }
            elseif (-not $Definition.Package.SetupFile -or -not (Test-Path -LiteralPath (Join-Path $source $Definition.Package.SetupFile) -PathType Leaf)) {
                $errors.Add("Package.SetupFile not found in the source folder: $($Definition.Package.SetupFile)")
            }
        }
        if ($app.LogoPath) {
            $logo = Resolve-IpPath -Path $app.LogoPath -BaseDirectory $BaseDirectory
            if (-not (Test-Path -LiteralPath $logo -PathType Leaf)) { $errors.Add("App.LogoPath not found: $logo") }
        }
    }

    return [string[]]$errors.ToArray()
}

function ConvertTo-IpDetectionItem {
    <#
    .SYNOPSIS
        Returns a file or registry detection rule with every field present (missing fields defaulted).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowNull()]$Item,
        [Parameter(Mandatory = $true)][ValidateSet('File', 'Registry')][string]$Kind
    )

    $result = if ($Kind -eq 'File') {
        [ordered]@{ Path = ''; FileOrFolderName = ''; Check32BitOn64System = $false; DetectionType = 'exists'; Operator = 'notConfigured'; DetectionValue = '' }
    }
    else {
        [ordered]@{ KeyPath = ''; ValueName = ''; Check32BitOn64System = $false; DetectionType = 'exists'; Operator = 'notConfigured'; DetectionValue = '' }
    }
    if ($null -ne $Item) {
        foreach ($key in @($result.Keys)) {
            $value = $null
            if ($Item -is [Collections.IDictionary]) { if ($Item.Contains($key)) { $value = $Item[$key] } }
            else {
                $property = $Item.PSObject.Properties[$key]
                if ($null -ne $property) { $value = $property.Value }
            }
            if ($null -ne $value) { $result[$key] = $value }
        }
    }
    return [pscustomobject]$result
}

function Get-IpDetectionValidationErrors {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)][psobject]$Detection,
        [string]$BaseDirectory = ''
    )

    $errors = [Collections.Generic.List[string]]::new()
    switch ([string]$Detection.Type) {
        'MsiProductCode' {
            $msi = $Detection.Msi
            if (-not (Test-IpGuid $msi.ProductCode)) { $errors.Add('Detection.Msi.ProductCode must be a GUID such as {00000000-0000-0000-0000-000000000000}.') }
            if ($msi.ProductVersionOperator -notin $script:IpDetectionOperators) { $errors.Add('Detection.Msi.ProductVersionOperator is not a supported operator.') }
            if ($msi.ProductVersionOperator -ne 'notConfigured' -and [string]::IsNullOrWhiteSpace($msi.ProductVersion)) {
                $errors.Add('Detection.Msi.ProductVersion is required when a version operator is configured.')
            }
        }
        'File' {
            $rules = @(@($Detection.File) | ForEach-Object { ConvertTo-IpDetectionItem -Item $_ -Kind File })
            if ($rules.Count -eq 0) { $errors.Add('Detection.File must contain at least one rule.') }
            foreach ($rule in $rules) {
                if ([string]::IsNullOrWhiteSpace($rule.Path)) { $errors.Add('File detection Path is required.') }
                if ([string]::IsNullOrWhiteSpace($rule.FileOrFolderName)) { $errors.Add('File detection FileOrFolderName is required.') }
                if ($rule.DetectionType -notin $script:IpFileDetectionTypes) { $errors.Add(("File detection type must be one of: {0}." -f ($script:IpFileDetectionTypes -join ', '))) }
                elseif ($rule.DetectionType -notin @('exists', 'doesNotExist')) {
                    if ($rule.Operator -notin $script:IpDetectionOperators -or $rule.Operator -eq 'notConfigured') { $errors.Add('File detection with a value comparison requires an operator.') }
                    if ([string]::IsNullOrWhiteSpace([string]$rule.DetectionValue)) { $errors.Add('File detection with a value comparison requires DetectionValue.') }
                }
            }
        }
        'Registry' {
            $rules = @(@($Detection.Registry) | ForEach-Object { ConvertTo-IpDetectionItem -Item $_ -Kind Registry })
            if ($rules.Count -eq 0) { $errors.Add('Detection.Registry must contain at least one rule.') }
            foreach ($rule in $rules) {
                if ([string]$rule.KeyPath -notmatch '^(HKEY_LOCAL_MACHINE|HKLM|HKEY_CURRENT_USER|HKCU)\\') { $errors.Add('Registry detection KeyPath must start with HKEY_LOCAL_MACHINE\ or HKEY_CURRENT_USER\.') }
                if ($rule.DetectionType -notin $script:IpRegistryDetectionTypes) { $errors.Add(("Registry detection type must be one of: {0}." -f ($script:IpRegistryDetectionTypes -join ', '))) }
                elseif ($rule.DetectionType -notin @('exists', 'doesNotExist')) {
                    if ([string]::IsNullOrWhiteSpace($rule.ValueName)) { $errors.Add('Registry detection with a value comparison requires ValueName.') }
                    if ($rule.Operator -notin $script:IpDetectionOperators -or $rule.Operator -eq 'notConfigured') { $errors.Add('Registry detection with a value comparison requires an operator.') }
                    if ([string]::IsNullOrWhiteSpace([string]$rule.DetectionValue)) { $errors.Add('Registry detection with a value comparison requires DetectionValue.') }
                }
            }
        }
        'Script' {
            $scriptPath = [string]$Detection.Script.ScriptPath
            if ([string]::IsNullOrWhiteSpace($scriptPath)) { $errors.Add('Detection.Script.ScriptPath is required.') }
            elseif ([IO.Path]::GetExtension($scriptPath) -ne '.ps1') { $errors.Add('Detection.Script.ScriptPath must be a .ps1 file.') }
            elseif ($BaseDirectory -and -not (Test-Path -LiteralPath (Resolve-IpPath -Path $scriptPath -BaseDirectory $BaseDirectory) -PathType Leaf)) {
                $errors.Add("Detection script not found: $scriptPath")
            }
        }
        'Preserved' { }
        default { $errors.Add(('Detection.Type must be one of: {0}.' -f ($script:IpDetectionTypes -join ', '))) }
    }
    return [string[]]$errors.ToArray()
}

#endregion

#region MSI and PSADT discovery

function Get-IpMsiProperties {
    <#
    .SYNOPSIS
        Reads the Property table of an MSI database (read-only) through the Windows Installer COM API.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $installer = $null
    $database = $null
    $view = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        # 0 = msiOpenDatabaseModeReadOnly
        $database = $installer.OpenDatabase($resolved, 0)
        $view = $database.OpenView('SELECT `Property`, `Value` FROM `Property`')
        [void]$view.Execute()
        $properties = [ordered]@{}
        while ($true) {
            $record = $view.Fetch()
            if ($null -eq $record) { break }
            try { $properties[[string]$record.StringData(1)] = [string]$record.StringData(2) }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
        }
        [void]$view.Close()
        return $properties
    }
    catch {
        throw ("Unable to read MSI properties from '{0}': {1}" -f $resolved, $_.Exception.Message)
    }
    finally {
        foreach ($comObject in @($view, $database, $installer)) {
            if ($null -ne $comObject) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) }
        }
    }
}

function ConvertTo-IpMsiMetadata {
    <#
    .SYNOPSIS
        Maps MSI Property table values to app definition fields and default command lines.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][Collections.IDictionary]$Properties,
        [Parameter(Mandatory = $true)][string]$MsiFileName
    )

    $get = { param($name) if ($Properties.Contains($name)) { [string]$Properties[$name] } else { '' } }
    $productCode = & $get 'ProductCode'
    $allUsers = & $get 'ALLUSERS'
    $packageType = switch ($allUsers) { '1' { 'perMachine' } '2' { 'dualPurpose' } default { 'perUser' } }
    $uninstall = if ($productCode) { 'msiexec /x {0} /qn /norestart' -f $productCode } else { '' }

    return [pscustomobject]@{
        ProductCode          = $productCode
        ProductVersion       = & $get 'ProductVersion'
        Publisher            = & $get 'Manufacturer'
        Name                 = & $get 'ProductName'
        UpgradeCode          = & $get 'UpgradeCode'
        PackageType          = $packageType
        RunAsAccount         = if ($packageType -eq 'perUser') { 'user' } else { 'system' }
        InstallCommandLine   = 'msiexec /i "{0}" /qn /norestart' -f $MsiFileName
        UninstallCommandLine = $uninstall
    }
}

function Get-IpMsiMetadata {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)
    $properties = Get-IpMsiProperties -Path $Path
    return (ConvertTo-IpMsiMetadata -Properties $properties -MsiFileName ([IO.Path]::GetFileName($Path)))
}

function Get-IpPsadtCommands {
    <#
    .SYNOPSIS
        Detects a PSAppDeployToolkit package root (v4 Invoke-AppDeployToolkit.exe, v3
        Deploy-Application.exe) and returns suggested setup file and command lines.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$SourceFolder)

    $candidates = @(
        @{ File = 'Invoke-AppDeployToolkit.exe'; Version = 4 }
        @{ File = 'Deploy-Application.exe'; Version = 3 }
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath (Join-Path $SourceFolder $candidate.File) -PathType Leaf) {
            return [pscustomobject]@{
                ToolkitVersion       = $candidate.Version
                SetupFile            = $candidate.File
                InstallCommandLine   = '{0} -DeploymentType Install -DeployMode Silent' -f $candidate.File
                UninstallCommandLine = '{0} -DeploymentType Uninstall -DeployMode Silent' -f $candidate.File
            }
        }
    }
    return $null
}

#endregion

#region .intunewin metadata

function Get-IpXmlChildText {
    param([Xml.XmlNode]$Node, [string]$Name)
    if ($null -eq $Node) { return '' }
    $child = $Node.SelectSingleNode("*[local-name()='$Name']")
    if ($null -eq $child) { return '' }
    return ([string]$child.InnerText).Trim()
}

function ConvertFrom-IpDetectionXml {
    <#
    .SYNOPSIS
        Parses the Detection.xml produced by the Win32 Content Prep Tool using explicit element paths.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$XmlText)

    try { [xml]$xml = $XmlText }
    catch { throw ("Detection.xml is not valid XML: {0}" -f $_.Exception.Message) }

    $root = $xml.SelectSingleNode("/*[local-name()='ApplicationInfo']")
    if ($null -eq $root) { throw 'Detection.xml does not contain an ApplicationInfo root element.' }
    $encryption = $root.SelectSingleNode("*[local-name()='EncryptionInfo']")
    if ($null -eq $encryption) { throw 'Detection.xml does not contain an EncryptionInfo element.' }

    $sizeText = Get-IpXmlChildText $root 'UnencryptedContentSize'
    $size = [int64]0
    if (-not [int64]::TryParse($sizeText, [ref]$size)) { throw "Detection.xml UnencryptedContentSize is not a number: '$sizeText'" }

    $metadata = [ordered]@{
        ToolVersion            = [string]$root.GetAttribute('ToolVersion')
        Name                   = Get-IpXmlChildText $root 'Name'
        FileName               = Get-IpXmlChildText $root 'FileName'
        SetupFile              = Get-IpXmlChildText $root 'SetupFile'
        UnencryptedContentSize = $size
        EncryptionKey          = Get-IpXmlChildText $encryption 'EncryptionKey'
        MacKey                 = Get-IpXmlChildText $encryption 'MacKey'
        InitializationVector   = Get-IpXmlChildText $encryption 'InitializationVector'
        Mac                    = Get-IpXmlChildText $encryption 'Mac'
        ProfileIdentifier      = Get-IpXmlChildText $encryption 'ProfileIdentifier'
        FileDigest             = Get-IpXmlChildText $encryption 'FileDigest'
        FileDigestAlgorithm    = Get-IpXmlChildText $encryption 'FileDigestAlgorithm'
        MsiInfo                = $null
    }
    if (-not $metadata.ProfileIdentifier) { $metadata.ProfileIdentifier = 'ProfileVersion1' }
    if (-not $metadata.FileDigestAlgorithm) { $metadata.FileDigestAlgorithm = 'SHA256' }
    foreach ($required in @('FileName', 'SetupFile', 'EncryptionKey', 'MacKey', 'InitializationVector', 'Mac', 'FileDigest')) {
        if ([string]::IsNullOrWhiteSpace([string]$metadata[$required])) { throw "Detection.xml is missing required field: $required" }
    }

    $msiNode = $root.SelectSingleNode("*[local-name()='MsiInfo']")
    if ($null -ne $msiNode) {
        $metadata.MsiInfo = [pscustomobject]@{
            ProductCode    = Get-IpXmlChildText $msiNode 'MsiProductCode'
            ProductVersion = Get-IpXmlChildText $msiNode 'MsiProductVersion'
            UpgradeCode    = Get-IpXmlChildText $msiNode 'MsiUpgradeCode'
            Publisher      = Get-IpXmlChildText $msiNode 'MsiPublisher'
            RequiresReboot = (Get-IpXmlChildText $msiNode 'MsiRequiresReboot') -eq 'true'
            IsMachine      = (Get-IpXmlChildText $msiNode 'MsiIsMachineInstall') -eq 'true'
            IsUser         = (Get-IpXmlChildText $msiNode 'MsiIsUserInstall') -eq 'true'
        }
    }
    return [pscustomobject]$metadata
}

function Read-IpIntuneWinMetadata {
    <#
    .SYNOPSIS
        Reads Detection.xml and locates the encrypted payload inside a .intunewin package.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    if ([IO.Path]::GetExtension($resolved) -ne '.intunewin') { throw "Expected a .intunewin file: $resolved" }

    $archive = [IO.Compression.ZipFile]::OpenRead($resolved)
    try {
        $entries = @($archive.Entries)
        $detectionEntry = $entries | Where-Object { $_.FullName -match '(^|/)Metadata/Detection\.xml$' } | Select-Object -First 1
        if (-not $detectionEntry) { $detectionEntry = $entries | Where-Object { $_.FullName -match '(^|/)Detection\.xml$' } | Select-Object -First 1 }
        if (-not $detectionEntry) { throw 'Detection.xml was not found inside the .intunewin package.' }

        $reader = [IO.StreamReader]::new($detectionEntry.Open())
        try { $xmlText = $reader.ReadToEnd() }
        finally { $reader.Dispose() }
        $metadata = ConvertFrom-IpDetectionXml -XmlText $xmlText

        $contentEntry = $entries | Where-Object { $_.FullName -match ('(^|/)Contents/' + [regex]::Escape($metadata.FileName) + '$') } | Select-Object -First 1
        if (-not $contentEntry) {
            $contentEntry = $entries | Where-Object { $_.FullName -match '(^|/)Contents/[^/]+$' } | Sort-Object Length -Descending | Select-Object -First 1
        }
        if (-not $contentEntry) { throw 'Encrypted payload was not found under Contents/ inside the .intunewin package.' }

        $metadata | Add-Member -NotePropertyName ContentEntryName -NotePropertyValue ([string]$contentEntry.FullName)
        $metadata | Add-Member -NotePropertyName SizeEncrypted -NotePropertyValue ([int64]$contentEntry.Length)
        $metadata | Add-Member -NotePropertyName IntuneWinPath -NotePropertyValue $resolved
        return $metadata
    }
    finally {
        $archive.Dispose()
    }
}

#endregion

#region Graph payloads

function ConvertTo-IpBase64Utf8 {
    param([AllowNull()][string]$Value)
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$Value))
}

function ConvertTo-IpRegistryKeyPath {
    param([string]$KeyPath)
    return ($KeyPath -replace '^HKLM\\', 'HKEY_LOCAL_MACHINE\' -replace '^HKCU\\', 'HKEY_CURRENT_USER\')
}

function New-IpDetectionRules {
    <#
    .SYNOPSIS
        Builds the win32LobApp detectionRules collection from Definition.Detection.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$Detection,
        [string]$BaseDirectory = ''
    )

    $errors = @(Get-IpDetectionValidationErrors -Detection $Detection -BaseDirectory $BaseDirectory)
    if ($errors.Count -gt 0) { throw ('Invalid detection rules: ' + ($errors -join ' ')) }

    switch ([string]$Detection.Type) {
        'MsiProductCode' {
            $msi = $Detection.Msi
            $rule = [ordered]@{
                '@odata.type'          = '#microsoft.graph.win32LobAppProductCodeDetection'
                productCode            = ([string]$msi.ProductCode).ToUpperInvariant()
                productVersionOperator = [string]$msi.ProductVersionOperator
                productVersion         = $null
            }
            if ($msi.ProductVersionOperator -ne 'notConfigured') { $rule.productVersion = [string]$msi.ProductVersion }
            return @($rule)
        }
        'File' {
            return @(foreach ($item in @(@($Detection.File) | ForEach-Object { ConvertTo-IpDetectionItem -Item $_ -Kind File })) {
                    $hasValue = $item.DetectionType -notin @('exists', 'doesNotExist')
                    [ordered]@{
                        '@odata.type'        = '#microsoft.graph.win32LobAppFileSystemDetection'
                        path                 = [string]$item.Path
                        fileOrFolderName     = [string]$item.FileOrFolderName
                        check32BitOn64System = [bool]$item.Check32BitOn64System
                        detectionType        = [string]$item.DetectionType
                        operator             = if ($hasValue) { [string]$item.Operator } else { 'notConfigured' }
                        detectionValue       = if ($hasValue) { [string]$item.DetectionValue } else { $null }
                    }
                })
        }
        'Registry' {
            return @(foreach ($item in @(@($Detection.Registry) | ForEach-Object { ConvertTo-IpDetectionItem -Item $_ -Kind Registry })) {
                    $hasValue = $item.DetectionType -notin @('exists', 'doesNotExist')
                    [ordered]@{
                        '@odata.type'        = '#microsoft.graph.win32LobAppRegistryDetection'
                        check32BitOn64System = [bool]$item.Check32BitOn64System
                        keyPath              = ConvertTo-IpRegistryKeyPath ([string]$item.KeyPath)
                        valueName            = if ([string]::IsNullOrWhiteSpace($item.ValueName)) { $null } else { [string]$item.ValueName }
                        detectionType        = [string]$item.DetectionType
                        operator             = if ($hasValue) { [string]$item.Operator } else { 'notConfigured' }
                        detectionValue       = if ($hasValue) { [string]$item.DetectionValue } else { $null }
                    }
                })
        }
        'Script' {
            $scriptPath = Resolve-IpPath -Path $Detection.Script.ScriptPath -BaseDirectory $(if ($BaseDirectory) { $BaseDirectory } else { (Get-Location).ProviderPath })
            $content = Get-Content -LiteralPath $scriptPath -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($content)) { throw "Detection script is empty: $scriptPath" }
            return @([ordered]@{
                    '@odata.type'         = '#microsoft.graph.win32LobAppPowerShellScriptDetection'
                    enforceSignatureCheck = [bool]$Detection.Script.EnforceSignatureCheck
                    runAs32Bit            = [bool]$Detection.Script.RunAs32Bit
                    scriptContent         = ConvertTo-IpBase64Utf8 -Value $content
                })
        }
        # Imported rules kept as they are in Intune: nothing is sent.
        'Preserved' { return @() }
    }
}

function New-IpReturnCodes {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$ReturnCodes)
    return @(foreach ($code in $ReturnCodes) {
            [ordered]@{
                '@odata.type' = '#microsoft.graph.win32LobAppReturnCode'
                returnCode    = [int]$code.ReturnCode
                type          = [string]$code.Type
            }
        })
}

function New-IpLargeIcon {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    $mime = switch ($extension) { '.png' { 'image/png' } { $_ -in '.jpg', '.jpeg' } { 'image/jpeg' } default { throw "Unsupported logo format: $extension" } }
    return [ordered]@{
        '@odata.type' = '#microsoft.graph.mimeContent'
        type          = $mime
        value         = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Path))
    }
}

function New-IpMsiInformation {
    [CmdletBinding()]
    param([AllowNull()][psobject]$MsiInfo, [psobject]$Definition)

    if ($null -eq $MsiInfo -or [string]::IsNullOrWhiteSpace($MsiInfo.ProductCode)) { return $null }
    $packageType = if ($MsiInfo.IsMachine -and $MsiInfo.IsUser) { 'dualPurpose' } elseif ($MsiInfo.IsUser) { 'perUser' } else { 'perMachine' }
    return [ordered]@{
        '@odata.type'  = '#microsoft.graph.win32LobAppMsiInformation'
        productCode    = [string]$MsiInfo.ProductCode
        productVersion = [string]$MsiInfo.ProductVersion
        upgradeCode    = [string]$MsiInfo.UpgradeCode
        requiresReboot = [bool]$MsiInfo.RequiresReboot
        packageType    = $packageType
        productName    = [string]$Definition.App.DisplayName
        publisher      = [string]$MsiInfo.Publisher
    }
}

function New-IpWin32LobAppBody {
    <#
    .SYNOPSIS
        Builds the win32LobApp create body from an app definition and .intunewin metadata.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$Definition,
        [Parameter(Mandatory = $true)][psobject]$IntuneWinMetadata,
        [string]$BaseDirectory = ''
    )

    $errors = @(Test-IpAppDefinition -Definition $Definition -ForPublish)
    if ($errors.Count -gt 0) { throw ('Invalid app definition: ' + ($errors -join ' ')) }

    $app = $Definition.App
    $program = $Definition.Program
    $nullIfEmpty = { param($value) if ([string]::IsNullOrWhiteSpace([string]$value)) { $null } else { [string]$value } }

    $body = [ordered]@{
        '@odata.type'                   = '#microsoft.graph.win32LobApp'
        displayName                     = [string]$app.DisplayName
        description                     = if ([string]::IsNullOrWhiteSpace($app.Description)) { [string]$app.DisplayName } else { [string]$app.Description }
        publisher                       = [string]$app.Publisher
        displayVersion                  = & $nullIfEmpty $app.Version
        notes                           = & $nullIfEmpty $app.Notes
        informationUrl                  = & $nullIfEmpty $app.InformationUrl
        privacyInformationUrl           = & $nullIfEmpty $app.PrivacyInformationUrl
        developer                       = & $nullIfEmpty $app.Developer
        owner                           = & $nullIfEmpty $app.Owner
        isFeatured                      = $false
        fileName                        = [string]$IntuneWinMetadata.FileName
        setupFilePath                   = [string]$IntuneWinMetadata.SetupFile
        installCommandLine              = [string]$program.InstallCommandLine
        uninstallCommandLine            = [string]$program.UninstallCommandLine
        installExperience               = [ordered]@{
            '@odata.type'         = '#microsoft.graph.win32LobAppInstallExperience'
            runAsAccount          = [string]$program.RunAsAccount
            deviceRestartBehavior = [string]$program.DeviceRestartBehavior
            maxRunTimeInMinutes   = [int]$program.MaxRunTimeInMinutes
        }
        # applicableArchitectures is deprecated; Graph sets it to none when allowedArchitectures is provided.
        applicableArchitectures         = 'none'
        # NotRestricted: allowedArchitectures null with applicableArchitectures none (no architecture requirement).
        allowedArchitectures            = if ('NotRestricted' -in @($Definition.Requirements.Architectures)) { $null } else { (@($Definition.Requirements.Architectures | Where-Object { $_ }) -join ',') }
        minimumSupportedOperatingSystem = [ordered]@{
            '@odata.type' = '#microsoft.graph.windowsMinimumOperatingSystem'
            v10_1607      = $true
        }
        minimumSupportedWindowsRelease  = [string]$Definition.Requirements.MinimumWindowsRelease
        detectionRules                  = @(New-IpDetectionRules -Detection $Definition.Detection -BaseDirectory $BaseDirectory)
        returnCodes                     = @(New-IpReturnCodes -ReturnCodes @($Definition.ReturnCodes))
    }

    $msiInformation = New-IpMsiInformation -MsiInfo $IntuneWinMetadata.MsiInfo -Definition $Definition
    if ($null -ne $msiInformation) { $body.msiInformation = $msiInformation }
    $disk = Get-IpObjectProperty $Definition.Requirements 'MinimumFreeDiskSpaceInMB'
    if ($null -ne $disk -and [string]$disk -ne '') { $body.minimumFreeDiskSpaceInMB = [int]$disk }

    if ($app.LogoPath) {
        $logoPath = Resolve-IpPath -Path $app.LogoPath -BaseDirectory $(if ($BaseDirectory) { $BaseDirectory } else { (Get-Location).ProviderPath })
        $body.largeIcon = New-IpLargeIcon -Path $logoPath
    }
    return $body
}

function New-IpContentFileBody {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][psobject]$IntuneWinMetadata)
    return [ordered]@{
        '@odata.type' = '#microsoft.graph.mobileAppContentFile'
        name          = [string]$IntuneWinMetadata.FileName
        size          = [int64]$IntuneWinMetadata.UnencryptedContentSize
        sizeEncrypted = [int64]$IntuneWinMetadata.SizeEncrypted
        manifest      = $null
        isDependency  = $false
    }
}

function New-IpCommitBody {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][psobject]$IntuneWinMetadata)
    return [ordered]@{
        fileEncryptionInfo = [ordered]@{
            '@odata.type'        = '#microsoft.graph.fileEncryptionInfo'
            encryptionKey        = [string]$IntuneWinMetadata.EncryptionKey
            macKey               = [string]$IntuneWinMetadata.MacKey
            initializationVector = [string]$IntuneWinMetadata.InitializationVector
            mac                  = [string]$IntuneWinMetadata.Mac
            profileIdentifier    = [string]$IntuneWinMetadata.ProfileIdentifier
            fileDigest           = [string]$IntuneWinMetadata.FileDigest
            fileDigestAlgorithm  = [string]$IntuneWinMetadata.FileDigestAlgorithm
        }
    }
}

function New-IpFinalizePatchBody {
    <#
    .SYNOPSIS
        Builds the PATCH body that sets committedContentVersion and reapplies app metadata.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][Collections.IDictionary]$AppBody,
        [Parameter(Mandatory = $true)][string]$ContentVersionId
    )

    $patch = [ordered]@{}
    foreach ($key in $AppBody.Keys) {
        # Read-only after creation or deprecated: not resent on PATCH.
        if ($key -in @('applicableArchitectures')) { continue }
        $patch[$key] = $AppBody[$key]
    }
    $patch.committedContentVersion = $ContentVersionId
    return $patch
}

function Protect-IpPayloadForDisplay {
    <#
    .SYNOPSIS
        Returns a deep copy of a payload with secrets redacted and long base64 values truncated.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Payload, [int]$MaxValueLength = 96)

    if ($null -eq $Payload) { return $null }
    $copy = ConvertTo-Json -InputObject $Payload -Depth 30 | ConvertFrom-Json -AsHashtable -NoEnumerate
    $secretNames = @('encryptionKey', 'macKey', 'initializationVector', 'mac', 'fileDigest', 'azureStorageUri')
    $walk = {
        param($node)
        if ($node -is [Collections.IDictionary]) {
            foreach ($key in @($node.Keys)) {
                $value = $node[$key]
                if ($key -in $secretNames -and $value) { $node[$key] = '<redacted>' }
                elseif ($value -is [string] -and $value.Length -gt $MaxValueLength -and $key -in @('value', 'scriptContent')) {
                    $node[$key] = '{0}... <{1} characters>' -f $value.Substring(0, 32), $value.Length
                }
                else { & $walk $value }
            }
        }
        elseif ($node -is [Collections.IList]) { foreach ($item in $node) { & $walk $item } }
    }
    & $walk $copy
    return $copy
}

function Get-IpIntunePortalAppUrl {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)][string]$AppId)
    return ('https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/{0}' -f $AppId)
}

#endregion

#region Differential update and import (no network call)

function ConvertFrom-IpBase64Text {
    <#
    .SYNOPSIS
        Decodes base64 script content, honoring a UTF-8 or UTF-16 byte order mark.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    try { $bytes = [Convert]::FromBase64String($Value) } catch { return $Value }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return [Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3) }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2) }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2) }
    return [Text.Encoding]::UTF8.GetString($bytes)
}

function ConvertTo-IpNormalizedScript {
    <#
    .SYNOPSIS
        Normalizes script text for comparison: no BOM, LF line endings, no trailing spaces or blank lines.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $lines = $Text.Replace([string][char]0xFEFF, '').Replace("`r`n", "`n").Replace("`r", "`n").Split("`n") | ForEach-Object { $_.TrimEnd() }
    return ((@($lines) -join "`n").TrimEnd("`n"))
}

function Get-IpRuleKind {
    param([AllowNull()][string]$ODataType)
    switch -Regex ([string]$ODataType) {
        'ProductCode' { return 'ProductCode' }
        'FileSystem' { return 'FileSystem' }
        'Registry' { return 'Registry' }
        'PowerShellScript' { return 'PowerShellScript' }
        default { return 'Unknown' }
    }
}

function Get-IpRuleSignature {
    <#
    .SYNOPSIS
        Returns a normalized signature of a detection or requirement rule, in the legacy
        (detectionRules/requirementRules) or the rules-collection shape, so both compare equal.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)]$Rule, [string]$RuleType = '')

    $get = { param($name) [string](Get-IpObjectProperty $Rule $name) }
    $kind = Get-IpRuleKind (& $get '@odata.type')
    $type = & $get 'detectionType'; if (-not $type) { $type = & $get 'operationType' }
    $value = & $get 'detectionValue'; if (-not $value -and $null -ne (Get-IpObjectProperty $Rule 'comparisonValue')) { $value = & $get 'comparisonValue' }
    $operator = & $get 'operator'
    # Existence checks ignore leftover operator and value (portal artifacts).
    if ($type -eq 'notConfigured') { $type = '' }
    if ($operator -eq 'notConfigured') { $operator = '' }
    if ($type -in @('exists', 'doesNotExist', '')) { $operator = ''; $value = '' }
    $check32 = [string][bool](Get-IpObjectProperty $Rule 'check32BitOn64System')
    $prefix = if ($RuleType) { $RuleType + ':' } else { '' }
    switch ($kind) {
        'ProductCode' {
            $versionOperator = & $get 'productVersionOperator'
            $version = if ($versionOperator -in @('', 'notConfigured')) { '' } else { & $get 'productVersion' }
            return ('{0}ProductCode|{1}|{2}|{3}' -f $prefix, (& $get 'productCode').ToUpperInvariant(), $versionOperator, $version)
        }
        'FileSystem' { return ('{0}FileSystem|{1}|{2}|{3}|{4}|{5}|{6}' -f $prefix, (& $get 'path').TrimEnd('\').ToLowerInvariant(), (& $get 'fileOrFolderName').ToLowerInvariant(), $check32, $type, $operator, $value) }
        'Registry' { return ('{0}Registry|{1}|{2}|{3}|{4}|{5}|{6}' -f $prefix, (ConvertTo-IpRegistryKeyPath (& $get 'keyPath')).ToLowerInvariant(), (& $get 'valueName').ToLowerInvariant(), $check32, $type, $operator, $value) }
        'PowerShellScript' {
            $script = ConvertTo-IpNormalizedScript (ConvertFrom-IpBase64Text (& $get 'scriptContent'))
            $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($script)))
            return ('{0}PowerShellScript|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}' -f $prefix, $hash, [bool](Get-IpObjectProperty $Rule 'enforceSignatureCheck'), [bool](Get-IpObjectProperty $Rule 'runAs32Bit'), (& $get 'runAsAccount'), $type, $operator, $value, (& $get 'displayName'))
        }
        default { return ('{0}Unknown|{1}' -f $prefix, (ConvertTo-Json -InputObject $Rule -Depth 6 -Compress)) }
    }
}

function Get-IpAppRuleSet {
    <#
    .SYNOPSIS
        Returns the detection and requirement rules of an app. The rules collection is preferred;
        the legacy detectionRules/requirementRules are used when rules is missing.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$App)
    $rules = @(Get-IpObjectProperty $App 'rules' | Where-Object { $null -ne $_ })
    if ($rules.Count) {
        return [pscustomobject]@{
            Source       = 'rules'
            Detection    = @($rules | Where-Object { [string](Get-IpObjectProperty $_ 'ruleType') -eq 'detection' })
            Requirements = @($rules | Where-Object { [string](Get-IpObjectProperty $_ 'ruleType') -eq 'requirement' })
        }
    }
    return [pscustomobject]@{
        Source       = 'legacy'
        Detection    = @(Get-IpObjectProperty $App 'detectionRules' | Where-Object { $null -ne $_ })
        Requirements = @(Get-IpObjectProperty $App 'requirementRules' | Where-Object { $null -ne $_ })
    }
}

function ConvertTo-IpRulesCollectionItem {
    <#
    .SYNOPSIS
        Converts a legacy detection or requirement rule to the rules-collection shape with its ruleType.
        Items already in the rules shape are returned unchanged.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Rule, [Parameter(Mandatory = $true)][ValidateSet('detection', 'requirement')][string]$RuleType)

    $type = [string](Get-IpObjectProperty $Rule '@odata.type')
    if ($type -match 'Rule$' -and $null -ne (Get-IpObjectProperty $Rule 'ruleType')) { return $Rule }
    $kind = Get-IpRuleKind $type
    $item = [ordered]@{ '@odata.type' = ('#microsoft.graph.win32LobApp{0}Rule' -f $kind); ruleType = $RuleType }
    switch ($kind) {
        'ProductCode' {
            $item.productCode = Get-IpObjectProperty $Rule 'productCode'
            $item.productVersionOperator = Get-IpObjectProperty $Rule 'productVersionOperator'
            $item.productVersion = Get-IpObjectProperty $Rule 'productVersion'
        }
        'FileSystem' {
            $item.path = Get-IpObjectProperty $Rule 'path'
            $item.fileOrFolderName = Get-IpObjectProperty $Rule 'fileOrFolderName'
            $item.check32BitOn64System = [bool](Get-IpObjectProperty $Rule 'check32BitOn64System')
            $item.operationType = Get-IpObjectProperty $Rule 'detectionType'
            $item.operator = Get-IpObjectProperty $Rule 'operator'
            $item.comparisonValue = Get-IpObjectProperty $Rule 'detectionValue'
        }
        'Registry' {
            $item.check32BitOn64System = [bool](Get-IpObjectProperty $Rule 'check32BitOn64System')
            $item.keyPath = Get-IpObjectProperty $Rule 'keyPath'
            $item.valueName = Get-IpObjectProperty $Rule 'valueName'
            $item.operationType = Get-IpObjectProperty $Rule 'detectionType'
            $item.operator = Get-IpObjectProperty $Rule 'operator'
            $item.comparisonValue = Get-IpObjectProperty $Rule 'detectionValue'
        }
        'PowerShellScript' {
            $item.displayName = Get-IpObjectProperty $Rule 'displayName'
            $item.enforceSignatureCheck = [bool](Get-IpObjectProperty $Rule 'enforceSignatureCheck')
            $item.runAs32Bit = [bool](Get-IpObjectProperty $Rule 'runAs32Bit')
            $item.runAsAccount = Get-IpObjectProperty $Rule 'runAsAccount'
            $item.scriptContent = Get-IpObjectProperty $Rule 'scriptContent'
            $operation = Get-IpObjectProperty $Rule 'detectionType'
            $item.operationType = if ($operation) { $operation } else { 'notConfigured' }
            $operator = Get-IpObjectProperty $Rule 'operator'
            $item.operator = if ($operator) { $operator } else { 'notConfigured' }
            $item.comparisonValue = Get-IpObjectProperty $Rule 'detectionValue'
        }
        default { throw "Unsupported rule type for the rules collection: $type" }
    }
    return [pscustomobject]$item
}

function ConvertTo-IpIsoUtc {
    <#
    .SYNOPSIS
        Normalizes a Graph timestamp to 'yyyy-MM-ddTHH:mm:ssZ' (UTC, second precision) for comparisons.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Value)
    if ($null -eq $Value -or [string]$Value -eq '') { return '' }
    $culture = [Globalization.CultureInfo]::InvariantCulture
    $utc = $null
    if ($Value -is [datetimeoffset]) { $utc = $Value.UtcDateTime }
    elseif ($Value -is [datetime]) { $utc = if ($Value.Kind -eq 'Local') { $Value.ToUniversalTime() } else { [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) } }
    else {
        $parsed = [datetimeoffset]::MinValue
        if (-not [datetimeoffset]::TryParse([string]$Value, $culture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { return [string]$Value }
        $utc = $parsed.UtcDateTime
    }
    return $utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $culture)
}

function Test-IpImportConcurrency {
    <#
    .SYNOPSIS
        Returns a message when the app changed in Intune since the definition was imported, or ''.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)][psobject]$Definition, [Parameter(Mandatory = $true)]$Current, [string]$ExpectedLastModified = '')

    $expected = $ExpectedLastModified
    if (-not $expected) {
        $imported = Get-IpObjectProperty $Definition 'ImportedFrom'
        if ($null -eq $imported -or [string](Get-IpObjectProperty $imported 'AppId') -ne [string](Get-IpObjectProperty $Current 'id')) { return '' }
        $expected = [string](Get-IpObjectProperty $imported 'LastModifiedDateTime')
    }
    if (-not $expected) { return '' }
    $actual = ConvertTo-IpIsoUtc (Get-IpObjectProperty $Current 'lastModifiedDateTime')
    if ((ConvertTo-IpIsoUtc $expected) -eq $actual) { return '' }
    return ("[Concurrency] The app was modified in Intune after the definition was read (read: {0}; current: {1}). Re-import the definition, or force the publication to replace the portal changes." -f (ConvertTo-IpIsoUtc $expected), $actual)
}

function Update-IpImportedFrom {
    <#
    .SYNOPSIS
        After a successful publication, records in Definition.ImportedFrom the app state read again by
        the post-publication check (lastModifiedDateTime, content version, display name), so the next
        publication with the same profile is not blocked by its own previous update.

    .OUTPUTS
        $true when ImportedFrom was updated. Definitions not imported from Intune, or imported from
        another app, are left unchanged.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory = $true)][psobject]$Definition, [AllowNull()]$VerifiedState)

    if ($null -eq $VerifiedState -or -not [string](Get-IpObjectProperty $VerifiedState 'LastModifiedDateTime')) { return $false }
    $imported = Get-IpObjectProperty $Definition 'ImportedFrom'
    if ($null -eq $imported -or -not [string](Get-IpObjectProperty $imported 'AppId')) { return $false }
    if ([string](Get-IpObjectProperty $imported 'AppId') -ne [string](Get-IpObjectProperty $VerifiedState 'AppId')) { return $false }
    $values = [ordered]@{
        DisplayName          = [string](Get-IpObjectProperty $VerifiedState 'DisplayName')
        ContentVersion       = [string](Get-IpObjectProperty $VerifiedState 'ContentVersion')
        LastModifiedDateTime = ConvertTo-IpIsoUtc (Get-IpObjectProperty $VerifiedState 'LastModifiedDateTime')
        SyncedAt             = ConvertTo-IpIsoUtc ([datetime]::UtcNow)
    }
    foreach ($name in $values.Keys) {
        if ($imported -is [Collections.IDictionary]) { $imported[$name] = $values[$name] }
        else { $imported | Add-Member -NotePropertyName $name -NotePropertyValue $values[$name] -Force }
    }
    return $true
}

function Get-IpEffectiveArchitectures {
    param($InputObject)
    $allowed = [string](Get-IpObjectProperty $InputObject 'allowedArchitectures')
    $values = if ($allowed -and $allowed -ne 'none') { $allowed } else { [string](Get-IpObjectProperty $InputObject 'applicableArchitectures') }
    if (-not $values -or $values -eq 'none') { return 'NotRestricted' }
    return ((@($values.Split(',') | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)) -join ',')
}

function Format-IpChangeValue {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '(empty)' }
    $text = if ($Value -is [string]) { $Value } else { ConvertTo-Json -InputObject $Value -Depth 6 -Compress }
    $text = $text.Replace("`r", ' ').Replace("`n", ' ')
    if ($text -eq '') { return '(empty)' }
    if ($text.Length -gt 70) { return $text.Substring(0, 67) + '...' }
    return $text
}

function ConvertTo-IpComparableText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Replace("`r`n", "`n").Trim()
}

function New-IpUpdatePatchBody {
    <#
    .SYNOPSIS
        Builds the differential PATCH of an UpdateExisting publication: committedContentVersion, the
        package properties and only the editable properties that differ from the current app.

    .DESCRIPTION
        Everything this tool does not edit (isFeatured, allowAvailableUninstall, scope tags, memory,
        CPU and processor requirements, minimumSupportedOperatingSystem flags, install scripts,
        requirement rules...) is never sent and stays as it is. installExperience is merged into the
        current object. A changed detection is sent through the rules collection together with the
        current requirement rules, never through detectionRules alone.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][Collections.IDictionary]$AppBody,
        [Parameter(Mandatory = $true)]$Current,
        [Parameter(Mandatory = $true)][string]$ContentVersionId,
        [Parameter(Mandatory = $true)][psobject]$Definition
    )

    $patch = [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobApp' }
    $changes = [Collections.Generic.List[object]]::new()
    $warnings = [Collections.Generic.List[string]]::new()
    $record = { param($name, $before, $after) $changes.Add([pscustomobject]@{ Property = $name; Before = (Format-IpChangeValue $before); After = (Format-IpChangeValue $after) }) }

    # Package properties: always sent with the new content.
    foreach ($name in @('fileName', 'setupFilePath')) {
        $patch[$name] = $AppBody[$name]
        if ((ConvertTo-IpComparableText $AppBody[$name]) -ne (ConvertTo-IpComparableText (Get-IpObjectProperty $Current $name))) { & $record $name (Get-IpObjectProperty $Current $name) $AppBody[$name] }
    }
    $newMsi = if ($AppBody.Contains('msiInformation')) { $AppBody['msiInformation'] } else { $null }
    $currentMsi = Get-IpObjectProperty $Current 'msiInformation'
    if ($null -ne $newMsi -or $null -ne $currentMsi) {
        $patch.msiInformation = $newMsi
        $msiKey = { param($msi) if ($null -eq $msi) { '' } else { '{0}|{1}|{2}' -f ([string](Get-IpObjectProperty $msi 'productCode')).ToUpperInvariant(), (Get-IpObjectProperty $msi 'productVersion'), (Get-IpObjectProperty $msi 'packageType') } }
        if ((& $msiKey $newMsi) -ne (& $msiKey $currentMsi)) { & $record 'msiInformation' $(if ($currentMsi) { Get-IpObjectProperty $currentMsi 'productCode' } else { $null }) $(if ($newMsi) { $newMsi.productCode } else { $null }) }
        if ($null -eq $newMsi -and $null -ne $currentMsi) { $warnings.Add('The new package is not an MSI while the app was: its MSI information is cleared.') }
    }

    foreach ($name in @('displayName', 'description', 'publisher', 'displayVersion', 'notes', 'informationUrl', 'privacyInformationUrl', 'developer', 'owner', 'installCommandLine', 'uninstallCommandLine', 'minimumSupportedWindowsRelease', 'minimumFreeDiskSpaceInMB')) {
        $new = if ($AppBody.Contains($name)) { $AppBody[$name] } else { $null }
        $old = Get-IpObjectProperty $Current $name
        if ((ConvertTo-IpComparableText $new) -ne (ConvertTo-IpComparableText $old)) { $patch[$name] = $new; & $record $name $old $new }
    }

    $newArchitectures = Get-IpEffectiveArchitectures ([pscustomobject]@{ allowedArchitectures = $AppBody['allowedArchitectures']; applicableArchitectures = 'none' })
    $currentArchitectures = Get-IpEffectiveArchitectures $Current
    if ($newArchitectures -ne $currentArchitectures) { $patch.allowedArchitectures = $AppBody['allowedArchitectures']; & $record 'allowedArchitectures' $currentArchitectures $newArchitectures }

    $experience = [ordered]@{ '@odata.type' = '#microsoft.graph.win32LobAppInstallExperience' }
    $currentExperience = Get-IpObjectProperty $Current 'installExperience'
    if ($null -ne $currentExperience) { foreach ($property in $currentExperience.PSObject.Properties) { if ($property.Name -ne '@odata.type') { $experience[$property.Name] = $property.Value } } }
    $experienceChanged = $false
    foreach ($name in @('runAsAccount', 'deviceRestartBehavior', 'maxRunTimeInMinutes')) {
        $new = $AppBody['installExperience'][$name]
        if ((ConvertTo-IpComparableText $new) -ne (ConvertTo-IpComparableText $experience[$name])) { & $record ('installExperience.' + $name) $experience[$name] $new; $experience[$name] = $new; $experienceChanged = $true }
    }
    if ($experienceChanged) { $patch.installExperience = $experience }

    $codeKey = { param($codes) (@($codes | Where-Object { $_ } | ForEach-Object { '{0}={1}' -f (Get-IpObjectProperty $_ 'returnCode'), (Get-IpObjectProperty $_ 'type') } | Sort-Object)) -join ',' }
    $newCodes = & $codeKey @($AppBody['returnCodes'])
    $oldCodes = & $codeKey @(Get-IpObjectProperty $Current 'returnCodes')
    if ($newCodes -ne $oldCodes) { $patch.returnCodes = $AppBody['returnCodes']; & $record 'returnCodes' $oldCodes $newCodes }

    if ($AppBody.Contains('largeIcon')) {
        $currentIcon = [string](Get-IpObjectProperty (Get-IpObjectProperty $Current 'largeIcon') 'value')
        if ([string]$AppBody['largeIcon']['value'] -ne $currentIcon) { $patch.largeIcon = $AppBody['largeIcon']; & $record 'largeIcon' $(if ($currentIcon) { '<logo>' } else { $null }) '<new logo>' }
    }

    $ruleSet = Get-IpAppRuleSet -App $Current
    $detectionChanged = $false
    if ([string]$Definition.Detection.Type -ne 'Preserved') {
        $newSignatures = (@($AppBody['detectionRules'] | ForEach-Object { Get-IpRuleSignature -Rule $_ } | Sort-Object)) -join "`n"
        $oldSignatures = (@($ruleSet.Detection | ForEach-Object { Get-IpRuleSignature -Rule $_ } | Sort-Object)) -join "`n"
        if ($newSignatures -ne $oldSignatures) {
            $detectionChanged = $true
            # Send detection through rules with the current requirement rules, so they are kept.
            $rules = [Collections.Generic.List[object]]::new()
            foreach ($rule in @($AppBody['detectionRules'])) { $rules.Add((ConvertTo-IpRulesCollectionItem -Rule ([pscustomobject]$rule) -RuleType detection)) }
            foreach ($rule in $ruleSet.Requirements) { $rules.Add((ConvertTo-IpRulesCollectionItem -Rule $rule -RuleType requirement)) }
            $patch.rules = $rules.ToArray()
            & $record 'detection rules' ('{0} rule(s)' -f $ruleSet.Detection.Count) ('{0} rule(s), {1} requirement rule(s) kept' -f @($AppBody['detectionRules']).Count, $ruleSet.Requirements.Count)
        }
    }

    $patch.committedContentVersion = $ContentVersionId
    return [pscustomobject]@{
        Body                      = $patch
        Changes                   = $changes.ToArray()
        Warnings                  = $warnings.ToArray()
        DetectionChanged          = $detectionChanged
        ExpectedDetection         = $(if ($detectionChanged) { @($AppBody['detectionRules']) } else { @($ruleSet.Detection) })
        PreservedRequirementRules = $ruleSet.Requirements.Count
    }
}

function Get-IpPreservedState {
    <#
    .SYNOPSIS
        Returns the state of what an update must keep unchanged, as comparable strings.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$App)
    $ruleSet = Get-IpAppRuleSet -App $App
    $state = [ordered]@{
        requirementRules        = (@($ruleSet.Requirements | ForEach-Object { Get-IpRuleSignature -Rule $_ -RuleType requirement } | Sort-Object)) -join "`n"
        isFeatured              = [string](Get-IpObjectProperty $App 'isFeatured')
        allowAvailableUninstall = [string](Get-IpObjectProperty $App 'allowAvailableUninstall')
        roleScopeTagIds         = (@(Get-IpObjectProperty $App 'roleScopeTagIds' | Sort-Object)) -join ','
        minimumMemoryInMB       = [string](Get-IpObjectProperty $App 'minimumMemoryInMB')
        minimumNumberOfProcessors = [string](Get-IpObjectProperty $App 'minimumNumberOfProcessors')
        minimumCpuSpeedInMHz    = [string](Get-IpObjectProperty $App 'minimumCpuSpeedInMHz')
        inUseBehavior           = [string](Get-IpObjectProperty (Get-IpObjectProperty $App 'installExperience') 'inUseBehavior')
        activeInstallScript     = ConvertTo-Json -InputObject (Get-IpObjectProperty $App 'activeInstallScript') -Compress -Depth 4
        activeUninstallScript   = ConvertTo-Json -InputObject (Get-IpObjectProperty $App 'activeUninstallScript') -Compress -Depth 4
    }
    return $state
}

function Test-IpUpdateResult {
    <#
    .SYNOPSIS
        Compares the app after an update with the state before it and the PATCH sent. Returns issues.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)]$Before,
        [Parameter(Mandatory = $true)]$After,
        [Parameter(Mandatory = $true)][psobject]$Update,
        [Parameter(Mandatory = $true)][string]$ContentVersionId
    )
    $issues = [Collections.Generic.List[string]]::new()
    if ([string](Get-IpObjectProperty $After 'committedContentVersion') -ne $ContentVersionId) {
        $issues.Add(("committedContentVersion is {0}, expected {1}." -f (Get-IpObjectProperty $After 'committedContentVersion'), $ContentVersionId))
    }
    $beforeState = Get-IpPreservedState -App $Before
    $afterState = Get-IpPreservedState -App $After
    $beforeRequirements = @((Get-IpAppRuleSet -App $Before).Requirements).Count
    $afterRequirements = @((Get-IpAppRuleSet -App $After).Requirements).Count
    if ($beforeState.requirementRules -ne $afterState.requirementRules) {
        $issues.Add(("Requirement rules changed: {0} before, {1} after the update. Check the app in the Intune admin center." -f $beforeRequirements, $afterRequirements))
    }
    foreach ($key in @($beforeState.Keys | Where-Object { $_ -ne 'requirementRules' })) {
        if ($beforeState[$key] -ne $afterState[$key]) { $issues.Add(("{0} changed although it is not edited by this tool: '{1}' -> '{2}'." -f $key, $beforeState[$key], $afterState[$key])) }
    }
    $expected = (@($Update.ExpectedDetection | ForEach-Object { Get-IpRuleSignature -Rule $_ } | Sort-Object)) -join "`n"
    $actual = (@((Get-IpAppRuleSet -App $After).Detection | ForEach-Object { Get-IpRuleSignature -Rule $_ } | Sort-Object)) -join "`n"
    if ($expected -ne $actual) { $issues.Add('The detection rules in Intune differ from the expected rules after the update.') }
    foreach ($name in @('displayName', 'publisher', 'installCommandLine', 'uninstallCommandLine')) {
        if ($Update.Body.Contains($name) -and (ConvertTo-IpComparableText $Update.Body[$name]) -ne (ConvertTo-IpComparableText (Get-IpObjectProperty $After $name))) {
            $issues.Add(("{0} was not applied." -f $name))
        }
    }
    return [string[]]$issues.ToArray()
}

function ConvertFrom-IpWin32App {
    <#
    .SYNOPSIS
        Converts a Win32 app read from Graph into an app definition (UpdateExisting), the detection
        script and logo to save, the list of preserved properties and warnings. Package source and
        output folders stay empty. No network call.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$App)

    $get = { param($name) Get-IpObjectProperty $App $name }
    $text = { param($name) $value = & $get $name; if ($null -eq $value) { '' } else { [string]$value } }
    $definition = New-IpAppDefinition
    $warnings = [Collections.Generic.List[string]]::new()
    $preserved = [Collections.Generic.List[string]]::new()

    $definition.App.DisplayName = & $text 'displayName'
    $definition.App.Description = & $text 'description'
    $definition.App.Publisher = & $text 'publisher'
    $definition.App.Version = & $text 'displayVersion'
    $definition.App.Notes = & $text 'notes'
    $definition.App.InformationUrl = & $text 'informationUrl'
    $definition.App.PrivacyInformationUrl = & $text 'privacyInformationUrl'
    $definition.App.Developer = & $text 'developer'
    $definition.App.Owner = & $text 'owner'
    $definition.Package.SetupFile = & $text 'setupFilePath'
    $definition.Program.InstallCommandLine = & $text 'installCommandLine'
    $definition.Program.UninstallCommandLine = & $text 'uninstallCommandLine'
    $experience = & $get 'installExperience'
    foreach ($pair in @(@('RunAsAccount', 'runAsAccount'), @('DeviceRestartBehavior', 'deviceRestartBehavior'), @('MaxRunTimeInMinutes', 'maxRunTimeInMinutes'))) {
        $value = Get-IpObjectProperty $experience $pair[1]
        if ($null -ne $value -and [string]$value -ne '') { $definition.Program.($pair[0]) = $value }
    }
    $codes = @(& $get 'returnCodes' | Where-Object { $_ } | ForEach-Object { [pscustomobject]@{ ReturnCode = [int](Get-IpObjectProperty $_ 'returnCode'); Type = [string](Get-IpObjectProperty $_ 'type') } })
    if ($codes.Count) { $definition.ReturnCodes = $codes }

    $architectures = Get-IpEffectiveArchitectures $App
    if ($architectures -eq 'NotRestricted') { $definition.Requirements.Architectures = @('NotRestricted') }
    else {
        $known = @($architectures.Split(',') | Where-Object { $_ -in $script:IpArchitectures })
        $unknown = @($architectures.Split(',') | Where-Object { $_ -notin $script:IpArchitectures })
        if ($unknown.Count) { $warnings.Add(("Architectures not supported by this tool were ignored: {0}." -f ($unknown -join ', '))) }
        $definition.Requirements.Architectures = if ($known.Count) { $known } else { @('NotRestricted') }
    }
    $release = & $text 'minimumSupportedWindowsRelease'
    if ($release) { $definition.Requirements.MinimumWindowsRelease = $release } else { $warnings.Add('The app has no minimum Windows release; 1607 is proposed.') }
    $disk = & $get 'minimumFreeDiskSpaceInMB'
    if ($null -ne $disk -and [string]$disk -ne '' -and [int]$disk -gt 0) { $definition.Requirements.MinimumFreeDiskSpaceInMB = [int]$disk }

    $ruleSet = Get-IpAppRuleSet -App $App
    $detection = @(& $get 'detectionRules' | Where-Object { $_ })
    if (-not $detection.Count) { $detection = @($ruleSet.Detection) }
    $kinds = @($detection | ForEach-Object { Get-IpRuleKind (Get-IpObjectProperty $_ '@odata.type') } | Sort-Object -Unique)
    $scriptText = ''
    $detectionType = if ($kinds.Count -ne 1) { 'Preserved' }
    elseif ($kinds[0] -eq 'ProductCode' -and $detection.Count -eq 1) { 'MsiProductCode' }
    elseif ($kinds[0] -eq 'FileSystem') { 'File' }
    elseif ($kinds[0] -eq 'Registry') { 'Registry' }
    elseif ($kinds[0] -eq 'PowerShellScript' -and $detection.Count -eq 1) { 'Script' }
    else { 'Preserved' }
    $value = { param($rule, $legacy, $current) $v = Get-IpObjectProperty $rule $legacy; if ($null -eq $v) { $v = Get-IpObjectProperty $rule $current }; $v }
    switch ($detectionType) {
        'MsiProductCode' {
            $rule = $detection[0]
            $definition.Detection.Msi.ProductCode = [string](Get-IpObjectProperty $rule 'productCode')
            $operator = [string](Get-IpObjectProperty $rule 'productVersionOperator')
            $definition.Detection.Msi.ProductVersionOperator = if ($operator) { $operator } else { 'notConfigured' }
            $definition.Detection.Msi.ProductVersion = [string](Get-IpObjectProperty $rule 'productVersion')
        }
        'File' {
            $definition.Detection.File = @(foreach ($rule in $detection) {
                    [pscustomobject]@{ Path = [string](Get-IpObjectProperty $rule 'path'); FileOrFolderName = [string](Get-IpObjectProperty $rule 'fileOrFolderName'); Check32BitOn64System = [bool](Get-IpObjectProperty $rule 'check32BitOn64System'); DetectionType = [string](& $value $rule 'detectionType' 'operationType'); Operator = [string](Get-IpObjectProperty $rule 'operator'); DetectionValue = [string](& $value $rule 'detectionValue' 'comparisonValue') }
                })
        }
        'Registry' {
            $definition.Detection.Registry = @(foreach ($rule in $detection) {
                    [pscustomobject]@{ KeyPath = [string](Get-IpObjectProperty $rule 'keyPath'); ValueName = [string](Get-IpObjectProperty $rule 'valueName'); Check32BitOn64System = [bool](Get-IpObjectProperty $rule 'check32BitOn64System'); DetectionType = [string](& $value $rule 'detectionType' 'operationType'); Operator = [string](Get-IpObjectProperty $rule 'operator'); DetectionValue = [string](& $value $rule 'detectionValue' 'comparisonValue') }
                })
        }
        'Script' {
            $rule = $detection[0]
            $scriptText = ConvertFrom-IpBase64Text ([string](Get-IpObjectProperty $rule 'scriptContent'))
            $definition.Detection.Script.EnforceSignatureCheck = [bool](Get-IpObjectProperty $rule 'enforceSignatureCheck')
            $definition.Detection.Script.RunAs32Bit = [bool](Get-IpObjectProperty $rule 'runAs32Bit')
            if (-not $scriptText) { $warnings.Add('The detection script content was not returned by Graph; the detection is kept as in Intune.'); $detectionType = 'Preserved' }
        }
    }
    $definition.Detection.Type = $detectionType
    if ($detectionType -eq 'Preserved') {
        $preserved.Add(("Detection rules: {0} rule(s) of type {1}, kept as in Intune (not editable here)." -f $detection.Count, (@($kinds) -join ' + ')))
        if ($detection.Count) { $warnings.Add('The detection combines several rule types or rules this tool cannot edit: it is kept as in Intune.') }
    }

    if ($ruleSet.Requirements.Count) {
        $names = @($ruleSet.Requirements | ForEach-Object { $n = [string](Get-IpObjectProperty $_ 'displayName'); '{0}{1}' -f (Get-IpRuleKind (Get-IpObjectProperty $_ '@odata.type')), $(if ($n) { " '$n'" } else { '' }) })
        $preserved.Add(("Requirement rules: {0} ({1})." -f $ruleSet.Requirements.Count, ($names -join ', ')))
    }
    # Every preserved property that has a value is listed, booleans included (false is a value).
    foreach ($name in @('isFeatured', 'allowAvailableUninstall', 'minimumMemoryInMB', 'minimumNumberOfProcessors', 'minimumCpuSpeedInMHz')) {
        $v = & $get $name
        if ($null -ne $v -and [string]$v -ne '') { $preserved.Add(('{0}: {1}.' -f $name, $v)) }
    }
    $tags = @(& $get 'roleScopeTagIds' | Where-Object { $null -ne $_ -and [string]$_ -ne '' })
    if ($tags.Count) { $preserved.Add(('roleScopeTagIds: {0} ({1}).' -f $tags.Count, ($tags -join ', '))) }
    $inUse = Get-IpObjectProperty $experience 'inUseBehavior'
    if ($null -ne $inUse -and [string]$inUse -ne '') { $preserved.Add(('installExperience.inUseBehavior: {0}.' -f $inUse)) }
    $operatingSystem = & $get 'minimumSupportedOperatingSystem'
    if ($null -ne $operatingSystem) {
        $names = if ($operatingSystem -is [Collections.IDictionary]) { @($operatingSystem.Keys) } else { @($operatingSystem.PSObject.Properties.Name) }
        $flags = @($names | Where-Object { $_ -ne '@odata.type' -and (Get-IpObjectProperty $operatingSystem $_) -eq $true })
        if ($flags.Count) { $preserved.Add(('minimumSupportedOperatingSystem: {0}.' -f ($flags -join ', '))) }
    }
    foreach ($name in @('activeInstallScript', 'activeUninstallScript')) {
        if ($null -ne (& $get $name)) { $preserved.Add(('{0}: set.' -f $name)); $warnings.Add(("The app uses {0}: Intune may run that script instead of the command line." -f $name)) }
    }
    $preserved.Add('Assignments, dependencies, supersedence and categories (never read or changed).')

    $icon = Get-IpObjectProperty $App 'largeIcon'
    $iconValue = [string](Get-IpObjectProperty $icon 'value')
    $iconBytes = $null
    $iconExtension = ''
    if ($iconValue) {
        try { $iconBytes = [Convert]::FromBase64String($iconValue); $iconExtension = if ([string](Get-IpObjectProperty $icon 'type') -match 'jpe?g') { '.jpg' } else { '.png' } }
        catch { $warnings.Add('The app logo could not be decoded; it is kept as in Intune.') }
    }

    $definition.Publish.Mode = 'UpdateExisting'
    $definition.Publish.ExistingAppId = & $text 'id'
    $definition.ImportedFrom.AppId = & $text 'id'
    $definition.ImportedFrom.DisplayName = & $text 'displayName'
    $definition.ImportedFrom.ContentVersion = & $text 'committedContentVersion'
    $definition.ImportedFrom.LastModifiedDateTime = ConvertTo-IpIsoUtc (& $get 'lastModifiedDateTime')
    $definition.ImportedFrom.ImportedAt = ConvertTo-IpIsoUtc ([datetime]::UtcNow)
    $definition.Preserved = $preserved.ToArray()

    return [pscustomobject]@{
        Definition    = $definition
        ScriptText    = $scriptText
        IconBytes     = $iconBytes
        IconExtension = $iconExtension
        Preserved     = $preserved.ToArray()
        Warnings      = $warnings.ToArray()
    }
}

function Save-IpImportAssets {
    <#
    .SYNOPSIS
        Writes the imported detection script and logo next to the profile and sets their paths.
        RelativeTo makes the paths relative (profile folder); otherwise they are absolute.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][psobject]$Import,
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][string]$BaseName,
        [switch]$Relative
    )
    [IO.Directory]::CreateDirectory($Folder) | Out-Null
    $safe = ($BaseName -replace '[^A-Za-z0-9._-]+', '-').Trim('-')
    if (-not $safe) { $safe = 'ImportedApp' }
    $definition = $Import.Definition
    if ($Import.ScriptText) {
        $name = $safe + '-Detection.ps1'
        [IO.File]::WriteAllText((Join-Path $Folder $name), $Import.ScriptText, [Text.UTF8Encoding]::new($false))
        $definition.Detection.Script.ScriptPath = if ($Relative) { $name } else { Join-Path $Folder $name }
    }
    if ($Import.IconBytes) {
        $name = $safe + '-Logo' + $Import.IconExtension
        [IO.File]::WriteAllBytes((Join-Path $Folder $name), [byte[]]$Import.IconBytes)
        $definition.App.LogoPath = if ($Relative) { $name } else { Join-Path $Folder $name }
    }
    return $definition
}

#endregion

Export-ModuleMember -Function @(
    'Format-IpLogLines', 'Set-IpLogSink', 'Set-IpLogPath', 'Get-IpLogPath', 'Initialize-IpLog', 'Write-IpLog', 'Format-IpByteSize', 'Format-IpLocalDateTime',
    'Copy-IpJsonValue', 'Merge-IpObjectFromTemplate', 'Save-IpJsonFile', 'Read-IpJsonFile', 'Import-IpConfiguration', 'Resolve-IpPath', 'Get-IpTenantProfile', 'Test-IpIdentifierGuid', 'Test-IpTenantIdentifier',
    'Get-IpDefaultReturnCodes', 'Get-IpWindowsReleases', 'Get-IpObjectProperty', 'Test-IpWindowsRelease', 'New-IpAppDefinition', 'Import-IpAppProfile', 'Export-IpAppProfile', 'Save-IpProfileImportedFrom', 'Test-IpGuid', 'Test-IpAppDefinition', 'Get-IpDetectionValidationErrors', 'ConvertTo-IpDetectionItem',
    'Get-IpMsiProperties', 'ConvertTo-IpMsiMetadata', 'Get-IpMsiMetadata', 'Get-IpPsadtCommands',
    'ConvertFrom-IpDetectionXml', 'Read-IpIntuneWinMetadata',
    'ConvertTo-IpBase64Utf8', 'New-IpDetectionRules', 'New-IpReturnCodes', 'New-IpLargeIcon', 'New-IpMsiInformation', 'New-IpWin32LobAppBody',
    'New-IpContentFileBody', 'New-IpCommitBody', 'New-IpFinalizePatchBody', 'Protect-IpPayloadForDisplay', 'Get-IpIntunePortalAppUrl',
    'ConvertFrom-IpBase64Text', 'ConvertTo-IpNormalizedScript', 'Get-IpRuleKind', 'Get-IpRuleSignature', 'Get-IpAppRuleSet', 'ConvertTo-IpRulesCollectionItem',
    'ConvertTo-IpIsoUtc', 'Test-IpImportConcurrency', 'Update-IpImportedFrom', 'Get-IpEffectiveArchitectures', 'Format-IpChangeValue', 'New-IpUpdatePatchBody', 'Get-IpPreservedState',
    'Test-IpUpdateResult', 'ConvertFrom-IpWin32App', 'Save-IpImportAssets'
)
