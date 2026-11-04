<#
.SYNOPSIS
    Connect one ShareGate endpoint with an explicitly selected authentication mode.
.VERSION
    1.0.4
#>
function Get-SmartM365ShareGateModule {
    # PowerShell can discover the installed GUI executable as a version 0 module.
    # It is not an importable PowerShell module and must not mask a valid manifest.
    $available = @(Get-Module -ListAvailable -Name ShareGate | Where-Object {
        $_.Path -and [IO.Path]::GetExtension([string]$_.Path) -in @('.psd1','.psm1','.dll')
    } | Sort-Object Version -Descending | Select-Object -First 1)
    if ($available.Count) { return $available[0] }
    foreach ($registryRoot in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
        $apps = @(Get-ItemProperty -Path (Join-Path $registryRoot '*') -ErrorAction SilentlyContinue | Where-Object {
            $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -match '^ShareGate( Migrate)?$' -and
            $_.PSObject.Properties['InstallLocation'] -and $_.InstallLocation
        })
        foreach ($app in $apps) {
            $candidate = Join-Path ([string]$app.InstallLocation) 'ShareGate\Sharegate.psd1'
            if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
            try {
                $manifest = Test-ModuleManifest -Path $candidate -ErrorAction Stop -WarningAction SilentlyContinue
                return [pscustomobject]@{ Version=$manifest.Version; Path=$candidate }
            } catch { throw "ShareGate module manifest validation failed: $($_.Exception.Message)" }
        }
    }
    $programFiles = [Environment]::GetFolderPath('ProgramFilesX86')
    if ($programFiles) {
        $candidate = Join-Path $programFiles 'Sharegate\ShareGate\Sharegate.psd1'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            try {
                $manifest = Test-ModuleManifest -Path $candidate -ErrorAction Stop -WarningAction SilentlyContinue
                return [pscustomobject]@{ Version=$manifest.Version; Path=$candidate }
            } catch { throw "ShareGate module manifest validation failed: $($_.Exception.Message)" }
        }
    }
}

function Get-SmartM365ShareGateSessionKey {
    param([Parameter(Mandatory = $true)][string]$Url, [Parameter(Mandatory = $true)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode)
    $uri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -notin @('http','https') -or $uri.UserInfo -or $uri.Query -or $uri.Fragment) {
        throw 'A session connection requires an absolute HTTP(S) site URL without credentials, query or fragment.'
    }
    return ('{0}|{1}' -f $Mode.ToLowerInvariant(), $uri.AbsoluteUri.TrimEnd('/').ToLowerInvariant())
}

function Register-SmartM365ShareGateSessionConnection {
    param([Parameter(Mandatory = $true)][string]$Url, [Parameter(Mandatory = $true)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode, [Parameter(Mandatory = $true)]$Site)
    $key = Get-SmartM365ShareGateSessionKey -Url $Url -Mode $Mode
    $address = $Site.PSObject.Properties['Address']
    if (-not $address -or -not $address.Value -or (Get-SmartM365ShareGateSessionKey -Url ([string]$address.Value) -Mode $Mode) -ne $key) {
        throw 'Cannot register a ShareGate connection whose site address differs from the requested URL.'
    }
    if (-not (Get-Variable -Name SmartM365ShareGateSessionConnections -Scope Global -ErrorAction SilentlyContinue)) {
        $global:SmartM365ShareGateSessionConnections = @{}
    }
    if ($global:SmartM365ShareGateSessionConnections.ContainsKey($key) -and
        -not [object]::ReferenceEquals($global:SmartM365ShareGateSessionConnections[$key], $Site)) {
        throw 'A different ShareGate connection is already registered for this URL and mode.'
    }
    $global:SmartM365ShareGateSessionConnections[$key] = $Site
}

function Clear-SmartM365ShareGateSessionConnections {
    Remove-Variable -Name SmartM365ShareGateSessionConnections -Scope Global -ErrorAction SilentlyContinue
}

function Connect-SmartM365ShareGateEndpointWithExistingCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode,
        [Parameter(Mandatory)]$CredentialSite,
        [switch]$AllowDifferentHost
    )
    $null = Get-SmartM365ShareGateSessionKey -Url $Url -Mode $Mode
    $address = $CredentialSite.PSObject.Properties['Address']
    if (-not $address -or -not $address.Value) { throw 'The credential source has no site address.' }
    $null = Get-SmartM365ShareGateSessionKey -Url ([string]$address.Value) -Mode $Mode
    $targetUri = [uri]$Url; $sourceUri = [uri]$address.Value
    if ($targetUri.Scheme -ne $sourceUri.Scheme -or $targetUri.Port -ne $sourceUri.Port) {
        throw 'Credential reuse requires the same protocol and port.'
    }
    if ($targetUri.DnsSafeHost -ine $sourceUri.DnsSafeHost -and -not $AllowDifferentHost) {
        throw 'Reusing credentials across different hosts requires explicit -AllowDifferentHost.'
    }
    $site = ShareGate\Connect-Site -Url $Url -UseCredentialsFrom $CredentialSite -ErrorAction Stop
    if (-not $site) { throw 'ShareGate returned no site for the reused credentials.' }
    Register-SmartM365ShareGateSessionConnection -Url $Url -Mode $Mode -Site $site
    return $site
}

function Connect-SmartM365ShareGateEndpointWithSavedConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode,
        [Parameter(Mandatory)]$SavedConnection
    )
    $null = Get-SmartM365ShareGateSessionKey -Url $Url -Mode $Mode
    $site = ShareGate\Connect-Site -Url $Url -UseConnection $SavedConnection -ErrorAction Stop
    if (-not $site) { throw 'ShareGate returned no site for the saved connection.' }
    Register-SmartM365ShareGateSessionConnection -Url $Url -Mode $Mode -Site $site
    return $site
}

function Connect-SmartM365ShareGateEndpoint {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][ValidateSet('Default','Browser','ModernAuth')][string]$Mode
    )
    $key = Get-SmartM365ShareGateSessionKey -Url $Url -Mode $Mode
    if ((Get-Variable -Name SmartM365ShareGateSessionConnections -Scope Global -ErrorAction SilentlyContinue) -and
        $global:SmartM365ShareGateSessionConnections.ContainsKey($key)) {
        $site = $global:SmartM365ShareGateSessionConnections[$key]
        $address = $site.PSObject.Properties['Address']
        if (-not $address -or -not $address.Value -or (Get-SmartM365ShareGateSessionKey -Url ([string]$address.Value) -Mode $Mode) -ne $key) {
            throw 'Registered ShareGate session connection no longer matches the requested URL.'
        }
        return $site
    }
    $arguments = @{ Url=$Url; ErrorAction='Stop' }
    if ($Mode -eq 'Browser') { $arguments.Browser = $true }
    if ($Mode -eq 'ModernAuth') { $arguments.ModernAuth = $true }
    $site = ShareGate\Connect-Site @arguments
    if (-not $site) { throw "ShareGate Connect-Site returned no site for $Url." }
    return $site
}

function Assert-SmartM365ShareGateDistinctSites {
    param([Parameter(Mandatory = $true)][string]$SourceUrl, [Parameter(Mandatory = $true)][string]$DestinationUrl)
    $source = $null; $destination = $null
    if (-not [uri]::TryCreate($SourceUrl, [UriKind]::Absolute, [ref]$source) -or
        -not [uri]::TryCreate($DestinationUrl, [UriKind]::Absolute, [ref]$destination) -or
        $source.UserInfo -or $source.Query -or $source.Fragment -or
        $destination.UserInfo -or $destination.Query -or $destination.Fragment -or
        $source.Scheme -notin @('http','https') -or
        $destination.Scheme -ne 'https' -or $destination.Port -ne 443 -or $destination.Host -notmatch '(?i)\.sharepoint\.(com|cn|de)$') {
        throw 'Source and destination must be valid site URLs and the destination must be HTTPS SharePoint Online.'
    }
    if ($source.Host -ine $destination.Host) { return }
    if ($source.Scheme -ne 'https' -or $source.Port -ne $destination.Port) {
        throw 'Same-host SharePoint sites must both use the same HTTPS endpoint.'
    }
    $collections = foreach ($site in @($source, $destination)) {
        $segments = @($site.AbsolutePath.Trim('/').Split('/') | Where-Object { $_ })
        if ($segments.Count -lt 2 -or $segments[0] -notin @('sites','teams')) {
            throw 'Same-host copy requires two explicit /sites/ or /teams/ site collections.'
        }
        $name = [uri]::UnescapeDataString($segments[1])
        if ([string]::IsNullOrWhiteSpace($name) -or $name -in @('.','..') -or $name.Contains('%') -or $name -match '[/\\\x00-\x1f]') {
            throw 'Invalid site collection segment in same-host copy.'
        }
        '{0}/{1}' -f $segments[0].ToLowerInvariant(), $name.ToLowerInvariant()
    }
    if ($collections[0] -eq $collections[1]) { throw 'Source and destination must be distinct site collections on the same host.' }
}
