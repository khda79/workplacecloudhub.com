#Requires -Version 7.0
<#
.SYNOPSIS
    Read-only diagnostic of existing Intune Win32 apps, to design "Import definition from Intune".

.DESCRIPTION
    Sends GET requests only (Microsoft Graph beta). Connects with the delegated scope
    DeviceManagementApps.Read.All (or the publication scope with -UsePublishScope when the app
    registration has no consent for the read scope).

    -Discover lists every Win32 app with its detection, rule and requirement types (one GET per
    app, nothing saved but a local summary) so that a representative sample can be chosen.
    -Apps exports a sample of up to 10 apps (names or app ids separated by ';'):
    - raw\        full GET responses, icons and detection/requirement scripts (local only, never share)
    - shared\     anonymized structure and report.md: property names, types, enum values and
                  presence only; no names, ids, commands, URLs or script content (safe to share)
    - mapping.csv sample labels (App01...) to names and ids (local only)

    Outputs go under %LOCALAPPDATA%\SmartM365\IntunePackager\Diagnostics by default; a folder
    inside the repository or the application folder is refused.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com

.EXAMPLE
    pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tools\SmartM365-IntunePackager-Win32AppDiagnostics.ps1 -Tenant test -Discover
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$Discover,
    # Display names (contains match, case-insensitive) or app ids, separated by ';'.
    [string]$Apps = '',
    [string]$OutputRoot = '%LOCALAPPDATA%\SmartM365\IntunePackager\Diagnostics',
    [switch]$UsePublishScope,
    [switch]$UseDeviceCode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:AppRoot = Split-Path -Path $PSScriptRoot -Parent
$script:MaxSample = 10
$script:ReadScope = 'DeviceManagementApps.Read.All'
# Values kept as-is in the shared output: types, enums, flags and numbers that carry no tenant data.
$script:KeptStringProperties = @(
    '@odata.type', 'runAsAccount', 'deviceRestartBehavior', 'detectionType', 'operator', 'operationType', 'ruleType',
    'productVersionOperator', 'applicableArchitectures', 'allowedArchitectures', 'minimumSupportedWindowsRelease',
    'publishingState', 'packageType', 'type', 'committedContentVersion', 'fileDigestAlgorithm'
)
$script:Base64Properties = @('scriptContent', 'value')

#region Helpers (tested offline)

function Split-IpDiagAppList {
    [CmdletBinding()]
    [OutputType([string[]])]
    param([AllowEmptyString()][string]$Text)
    return [string[]]@($Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
}

function Get-IpDiagRepositoryRoot {
    param([string]$Start)
    $current = $Start
    while ($current) {
        if (Test-Path -LiteralPath (Join-Path $current '.git')) { return $current }
        $parent = Split-Path -Path $current -Parent
        if ($parent -eq $current) { break }
        $current = $parent
    }
    return ''
}

function Test-IpDiagOutputRoot {
    <#
    .SYNOPSIS
        Returns why an output folder is refused (inside the application folder or a Git repository), or ''.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$AppRoot)
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\') + '\'
    foreach ($forbidden in @($AppRoot, (Get-IpDiagRepositoryRoot -Start $AppRoot), (Get-IpDiagRepositoryRoot -Start ([IO.Path]::GetFullPath($Path))))) {
        if (-not $forbidden) { continue }
        $root = [IO.Path]::GetFullPath($forbidden).TrimEnd('\') + '\'
        if ($full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { return "Diagnostic outputs must stay outside the application folder and any Git repository: $Path" }
    }
    return ''
}

function ConvertTo-IpDiagShape {
    <#
    .SYNOPSIS
        Returns an anonymized copy of a Graph object: structure, types, enums, flags and numbers only.
    #>
    [CmdletBinding()]
    param([AllowNull()]$Value, [string]$Name = '')

    if ($null -eq $Value) { return $null }
    if ($Value -is [bool] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return $Value }
    if ($Value -is [datetime] -or $Value -is [datetimeoffset]) { return '<datetime>' }
    if ($Value -is [string]) {
        if ($Value -eq '') { return '<empty>' }
        if ($Name -in $script:KeptStringProperties) { return $Value }
        if ($Name -in $script:Base64Properties) { return ('<base64:{0}>' -f $Value.Length) }
        $guid = [guid]::Empty
        if ([guid]::TryParse($Value, [ref]$guid)) { return '<guid>' }
        return ('<string:{0}>' -f $Value.Length)
    }
    if ($Value -is [Collections.IDictionary]) {
        $copy = [ordered]@{}
        foreach ($key in $Value.Keys) { $copy[[string]$key] = ConvertTo-IpDiagShape -Value $Value[$key] -Name ([string]$key) }
        return $copy
    }
    if ($Value -is [Collections.IEnumerable]) {
        return , @(foreach ($item in $Value) { ConvertTo-IpDiagShape -Value $item -Name $Name })
    }
    $object = [ordered]@{}
    foreach ($property in $Value.PSObject.Properties) { $object[$property.Name] = ConvertTo-IpDiagShape -Value $property.Value -Name $property.Name }
    return $object
}

function Get-IpDiagValue {
    param($InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-IpDiagShortType {
    param([AllowNull()][string]$ODataType)
    return ([string]$ODataType -replace '^#?microsoft\.graph\.win32LobApp', '' -replace '^#?microsoft\.graph\.', '')
}

function Get-IpDiagAppSummary {
    <#
    .SYNOPSIS
        Summarizes the detection, rule and requirement types of one Win32 app (no tenant data).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$App)

    $detection = @(Get-IpDiagValue $App 'detectionRules' | Where-Object { $_ })
    $rules = @(Get-IpDiagValue $App 'rules' | Where-Object { $_ })
    $requirements = @(Get-IpDiagValue $App 'requirementRules' | Where-Object { $_ })
    $extras = @(foreach ($name in @('minimumFreeDiskSpaceInMB', 'minimumMemoryInMB', 'minimumNumberOfProcessors', 'minimumCpuSpeedInMHz')) {
            $value = Get-IpDiagValue $App $name
            if ($null -ne $value -and [string]$value -ne '' -and [string]$value -ne '0') { $name }
        })
    $scripts = @(@($detection) + @($rules) + @($requirements) | Where-Object { (Get-IpDiagShortType (Get-IpDiagValue $_ '@odata.type')) -match 'PowerShellScript' })
    $scriptContents = @($scripts | Where-Object { -not [string]::IsNullOrWhiteSpace([string](Get-IpDiagValue $_ 'scriptContent')) })
    return [pscustomobject][ordered]@{
        DetectionRules       = (@($detection | ForEach-Object { Get-IpDiagShortType (Get-IpDiagValue $_ '@odata.type') }) -join ',')
        Rules                = (@($rules | ForEach-Object { '{0}:{1}' -f (Get-IpDiagValue $_ 'ruleType'), (Get-IpDiagShortType (Get-IpDiagValue $_ '@odata.type')) }) -join ',')
        RequirementRules     = (@($requirements | ForEach-Object { Get-IpDiagShortType (Get-IpDiagValue $_ '@odata.type') }) -join ',')
        ExtraRequirements    = ($extras -join ',')
        Scripts              = $scripts.Count
        ScriptContentPresent = ('{0}/{1}' -f $scriptContents.Count, $scripts.Count)
        ActiveInstallScript  = [bool](Get-IpDiagValue $App 'activeInstallScript')
        MsiInformation       = [bool](Get-IpDiagValue $App 'msiInformation')
        AllowedArchitectures = [string](Get-IpDiagValue $App 'allowedArchitectures')
        ApplicableArchitecture = [string](Get-IpDiagValue $App 'applicableArchitectures')
        MinimumWindowsRelease = [string](Get-IpDiagValue $App 'minimumSupportedWindowsRelease')
        LargeIconInFullGet   = [bool](Get-IpDiagValue (Get-IpDiagValue $App 'largeIcon') 'value')
    }
}

function Get-IpDiagScripts {
    <#
    .SYNOPSIS
        Returns the PowerShell scripts embedded in detection, rules and requirement rules (decoded).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$App)
    $sources = @(
        @{ Kind = 'detection'; Items = @(Get-IpDiagValue $App 'detectionRules' | Where-Object { $_ }) }
        @{ Kind = 'rule'; Items = @(Get-IpDiagValue $App 'rules' | Where-Object { $_ }) }
        @{ Kind = 'requirement'; Items = @(Get-IpDiagValue $App 'requirementRules' | Where-Object { $_ }) }
    )
    foreach ($source in $sources) {
        $index = 0
        foreach ($item in $source.Items) {
            $index++
            $content = [string](Get-IpDiagValue $item 'scriptContent')
            if ((Get-IpDiagShortType (Get-IpDiagValue $item '@odata.type')) -notmatch 'PowerShellScript' -or -not $content) { continue }
            $text = $content
            try { $text = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($content)) } catch { Write-Verbose 'Script content is not base64.' }
            $kind = $source.Kind
            if ($kind -eq 'rule') { $kind = 'rule-' + [string](Get-IpDiagValue $item 'ruleType') }
            [pscustomobject]@{ Kind = $kind; Index = $index; Text = $text }
        }
    }
}

function Get-IpDiagPropertyMatrix {
    <#
    .SYNOPSIS
        For each top-level property: in how many sampled apps it is present, non-null, and its kept values.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object[]]$Apps)
    $names = @($Apps | ForEach-Object { $_.PSObject.Properties.Name } | Sort-Object -Unique)
    foreach ($name in $names) {
        $present = 0; $set = 0; $values = [Collections.Generic.List[string]]::new()
        foreach ($app in $Apps) {
            $property = $app.PSObject.Properties[$name]
            if ($null -eq $property) { continue }
            $present++
            $value = $property.Value
            $isSet = $null -ne $value -and -not ($value -is [string] -and $value -eq '') -and -not ($value -is [Collections.ICollection] -and $value.Count -eq 0)
            if ($isSet) { $set++ }
            if ($isSet -and $name -in $script:KeptStringProperties -and $value -is [string] -and -not $values.Contains($value)) { $values.Add($value) }
            if ($isSet -and ($value -is [bool] -or $value -is [int] -or $value -is [long]) -and -not $values.Contains([string]$value)) { $values.Add([string]$value) }
        }
        [pscustomobject]@{ Property = $name; Present = $present; NonEmpty = $set; Values = ($values -join ', ') }
    }
}

function Write-IpDiagReport {
    param([Parameter(Mandatory = $true)][string]$Path, [object[]]$Summaries, [object[]]$Matrix, [object[]]$Notes)
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add('# Smart IntunePackager - Win32 app diagnostic (anonymized)')
    $lines.Add('')
    $lines.Add(('Generated: {0}; sampled apps: {1}. No name, id, command, URL or script content is included.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm'), @($Summaries).Count))
    $lines.Add('')
    $lines.Add('## Apps')
    $lines.Add('')
    $lines.Add('| App | Detection rules | Rules (ruleType:type) | Requirement rules | Extra requirements | Script content | MSI info | allowedArchitectures | applicableArchitectures | Min. release | Icon in GET | Icon with $select | Content versions |')
    $lines.Add('| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |')
    foreach ($s in $Summaries) {
        $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8} | {9} | {10} | {11} | {12} |' -f $s.Label, $s.DetectionRules, $s.Rules, $s.RequirementRules, $s.ExtraRequirements, $s.ScriptContentPresent, $s.MsiInformation, $s.AllowedArchitectures, $s.ApplicableArchitecture, $s.MinimumWindowsRelease, $s.LargeIconInFullGet, $s.LargeIconWithSelect, $s.ContentVersions))
    }
    $lines.Add('')
    $lines.Add('## Top-level properties returned by GET mobileApps/{id}')
    $lines.Add('')
    $lines.Add('| Property | Present | Non-empty | Observed values (enums and flags only) |')
    $lines.Add('| --- | --- | --- | --- |')
    foreach ($m in $Matrix) { $lines.Add(('| {0} | {1} | {2} | {3} |' -f $m.Property, $m.Present, $m.NonEmpty, $m.Values)) }
    if (@($Notes).Count) {
        $lines.Add('')
        $lines.Add('## Notes')
        $lines.Add('')
        foreach ($note in $Notes) { $lines.Add('- ' + $note) }
    }
    [IO.File]::WriteAllText($Path, ($lines -join "`r`n") + "`r`n", [Text.UTF8Encoding]::new($false))
}

#endregion

#region Graph access (GET only)

function Invoke-IpDiagGet {
    param([Parameter(Mandatory = $true)][string]$Uri, [string]$Step = 'Diagnostic')
    # The only Graph call of this script: always GET.
    return Invoke-IpGraphRequest -Method GET -Uri $Uri -Step $Step
}

function Invoke-IpDiagDiscover {
    param([Parameter(Mandatory = $true)][string]$RunFolder)
    $apps = @(Find-IpWin32App | Sort-Object displayName)
    Write-IpLog ("Discover: reading {0} Win32 app(s) one by one (GET only)." -f $apps.Count)
    $rows = [Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($app in $apps) {
        $index++
        $full = Invoke-IpDiagGet -Uri ("deviceAppManagement/mobileApps/{0}" -f $app.id) -Step 'DiscoverApp'
        $summary = Get-IpDiagAppSummary -App $full
        $row = [pscustomobject][ordered]@{ DisplayName = $app.displayName; AppId = $app.id }
        foreach ($property in $summary.PSObject.Properties) { $row | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value }
        $rows.Add($row)
        if ($index % 10 -eq 0) { Write-IpLog ("Discover: {0}/{1}" -f $index, $apps.Count) }
    }
    $csv = Join-Path $RunFolder 'discover-summary.local.csv'
    $rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8
    $rows | Format-Table DisplayName, DetectionRules, Rules, RequirementRules, ExtraRequirements, ScriptContentPresent, MsiInformation -AutoSize | Out-String -Width 240 | Microsoft.PowerShell.Utility\Write-Host
    Write-IpLog "Discover summary (local only, contains app names): $csv"
    Write-IpLog 'Pick 5 to 10 apps covering MSI, file, registry, script detection and requirement rules, then run again with -Apps "name1;name2;...".'
}

function Resolve-IpDiagSample {
    param([Parameter(Mandatory = $true)][string[]]$Entries)
    $all = $null
    $resolved = [Collections.Generic.List[object]]::new()
    foreach ($entry in $Entries) {
        if (Test-IpGuid $entry) { $resolved.Add([pscustomobject]@{ id = $entry; displayName = '' }); continue }
        if ($null -eq $all) { $all = @(Find-IpWin32App) }
        $exact = @($all | Where-Object { $_.displayName -eq $entry })
        $found = if ($exact.Count -eq 1) { $exact } else { @($all | Where-Object { $_.displayName -like ('*' + [WildcardPattern]::Escape($entry) + '*') }) }
        if ($found.Count -eq 1) { $resolved.Add($found[0]); continue }
        if ($found.Count -eq 0) { Write-IpLog "No Win32 app matches '$entry'; skipped." -Level WARN; continue }
        Write-IpLog ("'{0}' matches {1} apps; use a more precise name or the app id. Matches: {2}" -f $entry, $found.Count, (@($found | ForEach-Object { '{0} ({1})' -f $_.displayName, $_.id }) -join '; ')) -Level WARN
    }
    return @($resolved | Group-Object id | ForEach-Object { $_.Group[0] })
}

function Invoke-IpDiagSample {
    param([Parameter(Mandatory = $true)][object[]]$Sample, [Parameter(Mandatory = $true)][string]$RunFolder)
    $raw = Join-Path $RunFolder 'raw'
    $shared = Join-Path $RunFolder 'shared'
    [IO.Directory]::CreateDirectory($raw) | Out-Null
    [IO.Directory]::CreateDirectory($shared) | Out-Null
    $mapping = [Collections.Generic.List[object]]::new()
    $summaries = [Collections.Generic.List[object]]::new()
    $fullApps = [Collections.Generic.List[object]]::new()
    $notes = [Collections.Generic.List[string]]::new()
    $number = 0
    foreach ($entry in $Sample) {
        $number++
        $label = 'App{0:D2}' -f $number
        $app = Invoke-IpDiagGet -Uri ("deviceAppManagement/mobileApps/{0}" -f $entry.id) -Step 'GetApp'
        $type = [string](Get-IpDiagValue $app '@odata.type')
        if ($type -ne '#microsoft.graph.win32LobApp') { Write-IpLog ("{0}: not a Win32 app ({1}); skipped." -f $entry.id, $type) -Level WARN; $number--; continue }
        $fullApps.Add($app)
        $mapping.Add([pscustomobject]@{ Label = $label; DisplayName = [string]$app.displayName; AppId = [string]$app.id })

        $iconSelect = $null
        try { $iconSelect = Invoke-IpDiagGet -Uri ("deviceAppManagement/mobileApps/{0}?`$select=largeIcon" -f $entry.id) -Step 'GetIcon' }
        catch { $notes.Add(("{0}: GET with `$select=largeIcon failed: {1}" -f $label, $_.Exception.Message)) }
        $versions = @()
        try { $versions = @((Invoke-IpDiagGet -Uri ("deviceAppManagement/mobileApps/{0}/microsoft.graph.win32LobApp/contentVersions" -f $entry.id) -Step 'GetContentVersions').value) }
        catch { $notes.Add(("{0}: content versions unavailable: {1}" -f $label, $_.Exception.Message)) }

        # Raw files (local only).
        $icon = Get-IpDiagValue (Get-IpDiagValue $iconSelect 'largeIcon') 'value'
        if (-not $icon) { $icon = Get-IpDiagValue (Get-IpDiagValue $app 'largeIcon') 'value' }
        if ($icon) {
            $mime = [string](Get-IpDiagValue (Get-IpDiagValue $iconSelect 'largeIcon') 'type')
            $extension = if ($mime -match 'jpe?g') { '.jpg' } else { '.png' }
            try { [IO.File]::WriteAllBytes((Join-Path $raw ($label + '-icon' + $extension)), [Convert]::FromBase64String([string]$icon)) } catch { $notes.Add("${label}: icon could not be decoded.") }
        }
        foreach ($script in @(Get-IpDiagScripts -App $app)) {
            [IO.File]::WriteAllText((Join-Path $raw ('{0}-{1}-{2}.ps1' -f $label, $script.Kind, $script.Index)), $script.Text, [Text.UTF8Encoding]::new($false))
        }
        $rawCopy = $app | ConvertTo-Json -Depth 30 | ConvertFrom-Json
        if ($rawCopy.PSObject.Properties['largeIcon'] -and $rawCopy.largeIcon -and $rawCopy.largeIcon.PSObject.Properties['value'] -and $rawCopy.largeIcon.value) { $rawCopy.largeIcon.value = '<saved as icon file>' }
        [IO.File]::WriteAllText((Join-Path $raw ($label + '.json')), (ConvertTo-Json -InputObject $rawCopy -Depth 30), [Text.UTF8Encoding]::new($false))

        # Shared files (anonymized).
        [IO.File]::WriteAllText((Join-Path $shared ($label + '.shape.json')), (ConvertTo-Json -InputObject (ConvertTo-IpDiagShape -Value $app) -Depth 30), [Text.UTF8Encoding]::new($false))
        $summary = Get-IpDiagAppSummary -App $app
        $summary | Add-Member -NotePropertyName Label -NotePropertyValue $label
        $summary | Add-Member -NotePropertyName LargeIconWithSelect -NotePropertyValue ([bool](Get-IpDiagValue (Get-IpDiagValue $iconSelect 'largeIcon') 'value'))
        $summary | Add-Member -NotePropertyName ContentVersions -NotePropertyValue $versions.Count
        $summaries.Add($summary)
        Write-IpLog ("{0}: exported (detection={1}; rules={2}; requirements={3})." -f $label, $summary.DetectionRules, $summary.Rules, $summary.RequirementRules)
    }
    $mapping | Export-Csv -LiteralPath (Join-Path $RunFolder 'mapping.local.csv') -NoTypeInformation -Encoding utf8
    $matrix = @(Get-IpDiagPropertyMatrix -Apps $fullApps.ToArray())
    Write-IpDiagReport -Path (Join-Path $shared 'report.md') -Summaries $summaries.ToArray() -Matrix $matrix -Notes $notes.ToArray()
    return [pscustomobject]@{ Exported = $summaries.Count; Shared = $shared; Raw = $raw }
}

#endregion

#region Entry point

if ($MyInvocation.InvocationName -eq '.') { return }

foreach ($moduleName in @('SmartM365.IntunePackager.Core', 'SmartM365.IntunePackager.Graph')) {
    Import-Module (Join-Path $script:AppRoot ($moduleName + '.psm1')) -Force
}
Set-IpLogSink -Sink $null -Console $true

$entries = @(Split-IpDiagAppList -Text $Apps)
if (-not $Discover -and $entries.Count -eq 0) { throw 'Use -Discover to list the Win32 apps, or -Apps "name1;name2" to export a sample.' }
if ($entries.Count -gt $script:MaxSample) { throw ("Sample limited to {0} apps; {1} were given." -f $script:MaxSample, $entries.Count) }

$root = Resolve-IpPath -Path $OutputRoot -BaseDirectory $script:AppRoot
$issue = Test-IpDiagOutputRoot -Path $root -AppRoot $script:AppRoot
if ($issue) { throw $issue }
$runFolder = Join-Path $root (Get-Date -Format 'yyyyMMdd-HHmmss')
[IO.Directory]::CreateDirectory($runFolder) | Out-Null
Set-IpLogPath -Path (Join-Path $runFolder 'diagnostic.log')

$config = Import-IpConfiguration -ConfigPath (Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.config.json') -TemplatePath (Join-Path $script:AppRoot 'SmartM365-IntunePackager-GUI.config.json.template')
Set-IpGraphSettings -BaseUri ([string]$config.Graph.BaseUri) -MaxRetries ([int]$config.Graph.MaxRetries)
$tenantProfile = Get-IpTenantProfile -Config $config -Tenant $Tenant
if (-not $tenantProfile.TenantId) { throw "Tenant profile '$Tenant' has no TenantId; set it in the local configuration first." }
$scopes = if ($UsePublishScope) { Get-IpGraphScopes } else { @($script:ReadScope) }
Write-IpLog ("Read-only Win32 app diagnostic. Tenant profile={0}; Scopes={1}; Output={2}" -f $Tenant, ($scopes -join ', '), $runFolder)

$exitCode = 0
try {
    Connect-IpGraph -TenantProfile $tenantProfile -Scopes $scopes -UseDeviceCode:$UseDeviceCode | Out-Null
    if ($Discover) { Invoke-IpDiagDiscover -RunFolder $runFolder }
    if ($entries.Count) {
        $sample = @(Resolve-IpDiagSample -Entries $entries)
        if ($sample.Count -eq 0) { throw 'No app of the sample could be resolved.' }
        $result = Invoke-IpDiagSample -Sample $sample -RunFolder $runFolder
        Write-IpLog ("{0} app(s) exported. Share only the content of: {1}" -f $result.Exported, $result.Shared)
        Write-IpLog ("Local only (contains tenant data, do not share): {0} and {1}" -f $result.Raw, (Join-Path $runFolder 'mapping.local.csv'))
    }
}
catch {
    $exitCode = 1
    Write-IpLog $_.Exception.Message -Level ERROR
    Write-IpLog ("Error location: {0}" -f $_.ScriptStackTrace) -Level ERROR -FileOnly
}
finally {
    Disconnect-IpGraph
    Write-IpLog ("Diagnostic completed with exit code {0}." -f $exitCode)
}
exit $exitCode

#endregion
