<#
.SYNOPSIS
    Verify ShareGate endpoint authentication routing and same-host boundaries offline.
.VERSION
    1.0.5
#>
#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ShareGateAuth-' + [guid]::NewGuid().ToString('N'))
try {
    $module = Join-Path $root 'ShareGate'
    New-Item -ItemType Directory -Path $module -Force | Out-Null
    "@{ RootModule='ShareGate.psm1'; ModuleVersion='99.0.0'; GUID='7a75c782-6f07-4a08-84ba-91e7a1037883'; FunctionsToExport=@('Connect-Site') }" | Set-Content -LiteralPath (Join-Path $module 'ShareGate.psd1') -Encoding UTF8
    @'
function Connect-Site { [CmdletBinding()] param([string]$Url,[switch]$Browser,[switch]$ModernAuth,$UseCredentialsFrom,$UseConnection) [pscustomobject]@{Url=$Url; Address=$Url; Browser=[bool]$Browser; ModernAuth=[bool]$ModernAuth; UseCredentialsFrom=$UseCredentialsFrom; UseConnection=$UseConnection} }
Export-ModuleMember -Function Connect-Site
'@ | Set-Content -LiteralPath (Join-Path $module 'ShareGate.psm1') -Encoding UTF8
    Import-Module -Name (Join-Path $module 'ShareGate.psd1') -Force
    . (Join-Path $PSScriptRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGateAuth.ps1')
    foreach ($mode in @('Default','Browser','ModernAuth')) {
        $site = Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode $mode
        if ($site.Browser -ne ($mode -eq 'Browser') -or $site.ModernAuth -ne ($mode -eq 'ModernAuth')) { throw "Wrong switches for $mode" }
    }
    $sourceConnection=[pscustomobject]@{Address='https://example.test/sites/a/';Marker='source'}
    $targetConnection=[pscustomobject]@{Address='https://example.test/sites/b';Marker='target'}
    Register-SmartM365ShareGateSessionConnection -Url 'https://example.test/sites/a' -Mode Browser -Site $sourceConnection
    Register-SmartM365ShareGateSessionConnection -Url 'https://example.test/sites/b/' -Mode Browser -Site $targetConnection
    if (-not [object]::ReferenceEquals((Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode Browser),$sourceConnection)) { throw 'Source session was not reused.' }
    if (-not [object]::ReferenceEquals((Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/b' -Mode Browser),$targetConnection)) { throw 'Target session was not isolated.' }
    $differentMode=Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode ModernAuth
    if (-not $differentMode.ModernAuth) { throw 'A session was reused for a different authentication mode.' }
    $rejected=$false
    try { Register-SmartM365ShareGateSessionConnection -Url 'https://example.test/sites/b' -Mode Browser -Site $sourceConnection } catch { $rejected=$true }
    if (-not $rejected) { throw 'A connection with the wrong site address was accepted.' }
    $rejected=$false
    try { Register-SmartM365ShareGateSessionConnection -Url 'https://example.test/sites/a' -Mode Browser -Site ([pscustomobject]@{Address='https://example.test/sites/a'}) } catch { $rejected=$true }
    if (-not $rejected) { throw 'A registered connection was silently replaced.' }
    $sourceConnection.Address='https://example.test/sites/b'
    $rejected=$false
    try { Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode Browser | Out-Null } catch { $rejected=$true }
    if (-not $rejected) { throw 'A changed session address was reused.' }
    Clear-SmartM365ShareGateSessionConnections
    $defaultAfterClear=Connect-SmartM365ShareGateEndpoint -Url 'https://example.test/sites/a' -Mode Browser
    if (-not $defaultAfterClear.Browser) { throw 'Clearing the session did not restore default behavior.' }
    $credentialSite=[pscustomobject]@{Address='https://tenant.sharepoint.com/sites/source'}
    $fromCredentials=Connect-SmartM365ShareGateEndpointWithExistingCredentials -Url 'https://tenant.sharepoint.com/sites/target' -Mode Browser -CredentialSite $credentialSite
    if (-not [object]::ReferenceEquals($fromCredentials.UseCredentialsFrom,$credentialSite) -or -not [object]::ReferenceEquals((Connect-SmartM365ShareGateEndpoint -Url $fromCredentials.Address -Mode Browser),$fromCredentials)) { throw 'Explicit credential reuse was not registered.' }
    $blocked=$false
    try { Connect-SmartM365ShareGateEndpointWithExistingCredentials -Url 'https://other.sharepoint.com/sites/target' -Mode Browser -CredentialSite $credentialSite | Out-Null } catch { $blocked=$true }
    if(-not $blocked){throw 'Cross-host credential reuse was not explicitly guarded.'}
    $blocked=$false
    try { Connect-SmartM365ShareGateEndpointWithExistingCredentials -Url 'http://tenant.sharepoint.com/sites/target' -Mode Browser -CredentialSite $credentialSite -AllowDifferentHost | Out-Null } catch { $blocked=$true }
    if(-not $blocked){throw 'Protocol changes must not reuse credentials.'}
    $other=Connect-SmartM365ShareGateEndpointWithExistingCredentials -Url 'https://other.sharepoint.com/sites/target' -Mode Browser -CredentialSite $credentialSite -AllowDifferentHost
    if(-not [object]::ReferenceEquals($other.UseCredentialsFrom,$credentialSite)){throw 'Explicit cross-host reuse failed.'}
    $saved=[pscustomobject]@{Id='synthetic'}
    $fromSaved=Connect-SmartM365ShareGateEndpointWithSavedConnection -Url 'https://third.sharepoint.com/sites/target' -Mode Browser -SavedConnection $saved
    if(-not [object]::ReferenceEquals($fromSaved.UseConnection,$saved) -or -not [object]::ReferenceEquals((Connect-SmartM365ShareGateEndpoint -Url $fromSaved.Address -Mode Browser),$fromSaved)){throw 'Explicit saved connection was not registered.'}
    Clear-SmartM365ShareGateSessionConnections
    & {
        function Get-Module { param([switch]$ListAvailable,[string]$Name) }
        function Get-ItemProperty { param([string]$Path,$ErrorAction) [pscustomobject]@{DisplayName='ShareGate Migrate';InstallLocation=$root} }
        $discovered = Get-SmartM365ShareGateModule
        if (-not $discovered -or $discovered.Path -ne (Join-Path $root 'ShareGate\Sharegate.psd1') -or $discovered.Version -ne [version]'99.0.0') {
            throw 'Registry fallback did not resolve and validate the installed module manifest.'
        }
    }
    & {
        function Get-Module { param([switch]$ListAvailable,[string]$Name) [pscustomobject]@{Name='ShareGate';Version=[version]'99.0.0';Path=(Join-Path $root 'ShareGate.exe')} }
        function Get-ItemProperty { param([string]$Path,$ErrorAction) [pscustomobject]@{DisplayName='ShareGate Migrate';InstallLocation=$root} }
        $discovered = Get-SmartM365ShareGateModule
        if (-not $discovered -or $discovered.Path -ne (Join-Path $root 'ShareGate\Sharegate.psd1')) { throw 'GUI executable must not mask a valid module manifest.' }
    }
    & {
        function Get-Module { param([switch]$ListAvailable,[string]$Name) }
        function Get-ItemProperty { param([string]$Path,$ErrorAction) [pscustomobject]@{DisplayName='ShareGate Migrate';InstallLocation=$root} }
        function Test-ModuleManifest { param([string]$Path,$ErrorAction,$WarningAction) throw 'Invalid manifest fixture' }
        $blocked=$false
        try { Get-SmartM365ShareGateModule | Out-Null } catch { $blocked=$_.Exception.Message -match 'manifest validation failed' }
        if(-not $blocked){throw 'Invalid registered module manifest must not be imported.'}
    }
    Assert-SmartM365ShareGateDistinctSites -SourceUrl 'https://tenant.sharepoint.com/sites/a' -DestinationUrl 'https://tenant.sharepoint.com/sites/b'
    Assert-SmartM365ShareGateDistinctSites -SourceUrl 'https://farm.example.test/sites/a' -DestinationUrl 'https://tenant.sharepoint.com/sites/b'
    foreach ($pair in @(
        @('https://tenant.sharepoint.com/sites/a', 'https://tenant.sharepoint.com/sites/a'),
        @('https://tenant.sharepoint.com/sites/a/sub', 'https://tenant.sharepoint.com/sites/a/other'),
        @('https://tenant.sharepoint.com/sites/a', 'https://tenant.sharepoint.com/sites/%61'),
        @('https://tenant.sharepoint.com/sites/a', 'https://tenant.sharepoint.com/sites/a%2fb'),
        @('https://tenant.sharepoint.com/', 'https://tenant.sharepoint.com/sites/b'),
        @('https://tenant.sharepoint.com/sites/a', 'https://tenant.sharepoint.com/sites/b?x=1')
    )) {
        $blocked = $false
        try { Assert-SmartM365ShareGateDistinctSites -SourceUrl $pair[0] -DestinationUrl $pair[1] } catch { $blocked = $true }
        if (-not $blocked) { throw 'Unsafe or ambiguous boundary was accepted.' }
    }
    # ShareGate 26.11 exports cmdlets from a nested module: -Module ShareGate
    # does not find them, while the parent-qualified command name does.
    Remove-Module ShareGate -ErrorAction SilentlyContinue
    $nested = Join-Path $root 'Nested\ShareGate'
    New-Item -ItemType Directory -Path $nested -Force | Out-Null
    "@{ NestedModules=@('Sharegate.Automation.psm1'); ModuleVersion='99.0.1'; GUID='1eb75763-57f8-4141-8a2c-7ddf4e632772'; FunctionsToExport=@('Connect-Site') }" | Set-Content -LiteralPath (Join-Path $nested 'ShareGate.psd1') -Encoding UTF8
    "function Connect-Site { param([string]`$Url) `$Url }; Export-ModuleMember -Function Connect-Site" | Set-Content -LiteralPath (Join-Path $nested 'Sharegate.Automation.psm1') -Encoding UTF8
    Import-Module -Name (Join-Path $nested 'ShareGate.psd1') -Force
    $qualified = @(Get-Command -Name 'ShareGate\Connect-Site' -ErrorAction SilentlyContinue)
    if ($qualified.Count -ne 1 -or $qualified[0].Source -ne 'Sharegate.Automation') { throw 'Nested ShareGate command was not found via parent-qualified name.' }
    'PASS: independent ShareGate modes and distinct-site guards; no tenant access.'
}
finally {
    Clear-SmartM365ShareGateSessionConnections
    Remove-Module ShareGate -ErrorAction SilentlyContinue
    $resolved=[IO.Path]::GetFullPath($root)
    $allowed=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if ($resolved.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-ShareGateAuth-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
