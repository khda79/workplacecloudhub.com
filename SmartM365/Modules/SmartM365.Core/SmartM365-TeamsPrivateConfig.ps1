<#
.SYNOPSIS
    Resolves profile-specific private Teams webhook configuration.
.VERSION
    1.0.0
#>
function Get-SmartM365PrivateTeamsWebhook {
    param([Parameter(Mandatory)][string]$Name)
    $config = Get-SmartM365EffectiveModuleGlobalConfig
    $profile = [string]$config.ProfileKey
    if ($profile -notmatch '^[A-Za-z0-9][A-Za-z0-9_-]*$') { throw 'Invalid Teams configuration profile key.' }
    $path = Join-Path ([string]$config.SmartM365RootPath) ("Config\Tenants\{0}.teams.local.json" -f $profile)
    try {
        $path = Resolve-SmartM365JsonConfigurationPath $path
        if (-not (Test-Path -LiteralPath $path -PathType Leaf -ErrorAction Stop)) { return $null }
        $document = (Read-SmartM365JsonDocument $path).Document
        if ($document -isnot [pscustomobject]) { throw 'Invalid object.' }
        $property = $document.PSObject.Properties[$Name]
        if ($null -eq $property) { return $null }
        if ($property.Value -isnot [string]) { throw 'Invalid value.' }
        return [string]$property.Value
    } catch { throw 'Private Teams configuration could not be read or is invalid.' }
}
