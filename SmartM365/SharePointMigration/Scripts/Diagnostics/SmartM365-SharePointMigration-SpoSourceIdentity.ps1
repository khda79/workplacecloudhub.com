<#
.SYNOPSIS
    Verify a SharePoint Online source document's item ID and exact library path.
.VERSION
    1.0.0
#>
#Requires -Version 7.4

function Get-SmartM365SpoSourceIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$ListName,
        [Parameter(Mandatory = $true)][string]$SourceFilePath,
        [Parameter(Mandatory = $true)][ValidateRange(1, 2147483647)][int]$ExpectedItemId
    )

    $site = $null
    if (-not [uri]::TryCreate($SiteUrl, [UriKind]::Absolute, [ref]$site) -or
        $site.Scheme -ne 'https' -or $site.Port -ne 443 -or
        $site.Host -notmatch '(?i)\.sharepoint\.(com|cn|de)$' -or
        $site.UserInfo -or $site.Query -or $site.Fragment) {
        throw 'Source site must be an HTTPS SharePoint Online URL without credentials, query or fragment.'
    }
    $connectionUrl = $Connection.PSObject.Properties['Url']
    $connectionSite = $null
    if (-not $connectionUrl -or -not [uri]::TryCreate([string]$connectionUrl.Value, [UriKind]::Absolute, [ref]$connectionSite) -or
        $connectionSite.UserInfo -or $connectionSite.Query -or $connectionSite.Fragment -or
        -not [string]::Equals($connectionSite.AbsoluteUri.TrimEnd('/'), $site.AbsoluteUri.TrimEnd('/'), [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The PnP connection is not bound to the exact source web.'
    }
    if ([string]::IsNullOrWhiteSpace($ListName)) { throw 'Source list name is required.' }
    $raw = $SourceFilePath.Trim().Replace('\', '/')
    if (-not $raw -or $raw.StartsWith('/') -or $raw.Contains('://') -or $raw.Contains('?') -or $raw.Contains('#')) {
        throw 'Source file path must be a nonempty library-relative path.'
    }
    $decoded = @(foreach ($segment in $raw.Split('/')) {
        if (-not $segment -or $segment -match '%(?![0-9A-Fa-f]{2})') { throw 'Source file path contains an invalid segment.' }
        $value = [uri]::UnescapeDataString($segment)
        if (-not $value -or $value -in @('.', '..') -or $value -match '[/\\:%\x00-\x1f]') {
            throw 'Source file path contains an ambiguous or unsafe segment.'
        }
        $value
    })
    if (-not $decoded.Count) { throw 'Source file path is empty.' }
    $list = Get-PnPList -Identity $ListName -Includes RootFolder -Connection $Connection -ErrorAction Stop
    if (-not $list -or -not [string]::Equals([string]$list.Title, $ListName, [StringComparison]::OrdinalIgnoreCase) -or
        [string]$list.BaseType -ne 'DocumentLibrary' -or -not $list.Id -or -not $list.RootFolder.ServerRelativeUrl) {
        throw 'Source list is not the exact document library requested.'
    }
    $root = ([string]$list.RootFolder.ServerRelativeUrl).TrimEnd('/')
    $webPath = [uri]::UnescapeDataString($site.AbsolutePath).TrimEnd('/')
    if (-not $root.StartsWith($webPath + '/', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Source library root is outside the connected web.'
    }
    $expectedPath = $root + '/' + ($decoded -join '/')
    $item = Get-PnPFile -Url $expectedPath -AsListItem -Connection $Connection -ErrorAction Stop
    if (-not $item -or [int]$item.Id -ne $ExpectedItemId -or
        -not [string]::Equals([string]$item['FileRef'], $expectedPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Source item ID or server-relative path no longer matches the reviewed file.'
    }
    Get-PnPProperty -ClientObject $item -Property ParentList -Connection $Connection -ErrorAction Stop | Out-Null
    if (-not $item.ParentList -or [guid]$item.ParentList.Id -ne [guid]$list.Id) {
        throw 'Source item does not belong to the reviewed library.'
    }
    [pscustomobject]@{
        SiteUrl = $site.AbsoluteUri.TrimEnd('/')
        LibraryId = [string]$list.Id
        LibraryTitle = [string]$list.Title
        SourceItemId = [int]$item.Id
        SourceFilePath = ($decoded -join '/')
        ServerRelativePath = $expectedPath
        VerifiedUtc = [datetime]::UtcNow.ToString('o')
    }
}
