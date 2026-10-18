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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA1FFmHrUBnqXTK
# 35v5VLq3zzyhCdsw4HYKW3cQsLlufaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKmT9aTL5NMaj4lTHV7urLCZZQxt4nDuOIXoySXwLipnMA0GCSqG
# SIb3DQEBAQUABIIBgB+ujbBXuKiygkr1ZrOmwRye0eWktvIVhjO9g0L7UyPzICmP
# WMVrW/nuPOW8szYcQD1GKtYKj2oUIoyA1xesB0K83a+mCffClvoJw+EJ5FfMN+So
# bR0povIxnRb6yxq+AVDDws/ECSCnyXkEsHMum0Loebz28EhAYCfsXKyIIqSLkSf2
# eFoV6o6IUIwfJXGI6DUlpPJmH3j2JrmUYsVPKweVi73AcgdoTUUtIusooMRCx0Yv
# /Nyni0RbVllvN9gR1HQUKTzGpEHBQi5P5tFM1mr0gR0rzH7R+xNePAqi0izG1wtx
# Lj8IVlSWwZh3FoPvd2xDjvIswkyUxZ2jIAjWL9dub6D/J/TozcCT6FudhrRPd+7r
# F4SL7apDYF24l51NoBURlN40pwVeN3WcDjeaQNoWQq7YaT8BEVMLjgYDlGLsoeeV
# 81Mz8AYRyEtAIg0lwDWsfQecmNRc3XGsqkvOJzE40BoH/9jpaDLWLe4Oq7BSGzt3
# q0U4ERiZiCqdZ4Tab6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMjQ5
# MzhaMC8GCSqGSIb3DQEJBDEiBCDVbRRxVgIBHdRYHM1jocgXbHQGWxLPdfDVKr3k
# lRmhLDANBgkqhkiG9w0BAQEFAASCAgAxdOdl3bu9t4zLKnOL1PnPlLblDDcYMs0z
# MOIU40DbDLl5TBsFYDsnJ/WmtTp4/bRdn9t3Btgm33Sn6DcB3SQ5VTHRde445uQz
# e6I3AklFeCdOV9uwGs2tHEyJUYfCGxBYHUpk37YYCsDYYjWD6pydZdoqlmi3umC7
# +fAfbNDT1BpMb0XbrTfUF1Z2pzcArGsFNzHm09Ot3jz7m9cNFjjDzL+kiPWR0Z/K
# i06lfBiUiCuzNnZ4P5ebw88BK/qmwyN/0x4/F6wB8Tr0/ANxBOq/4b1Bn8jHmusM
# K3h1Bb9oi/YJUK4SkEd2XoBviamnFBZPbYCN+q7vpbmXu7RlN9jx55E7CAhhu+5S
# UNtQCBJyktQWAJYITPWGGmoNN5XnK/8CzboQAK1XQlgyl60F7jQY8IYaqpxFIk2s
# YX1jFqbLwplrsCjokaqnjhpPsXuyHmTwUFHNnG62Cc4B/dnkCFhw5weQ/e/wat4t
# DJcBYBRkKUEAH1glUD5lZdQ1l7a1yEZFeSXOHO7wO0JlIv6Blhpij+sr0rDanyV4
# R7GMaQTmsCKeZHwEmcD1AwRvRteOY6I0BOA6yY+udDBy06Qpkze9eCZ6KotjBpVT
# mMPq+twE6XnmLfj2CFZK1Pw2Kfa8QZVQKMwlpq4mXyg+nPJrRFH7i5z1C+j24O6R
# IpS/WeFaCw==
# SIG # End signature block
