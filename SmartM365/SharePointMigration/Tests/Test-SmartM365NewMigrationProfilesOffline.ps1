<#
.SYNOPSIS
    Verify independent SPO profiles during migration creation without network access.
.VERSION
    1.0.3
#>
#requires -Version 7.4
$ErrorActionPreference = 'Stop'
$path = Join-Path $PSScriptRoot '..\Scripts\SmartM365-SharePointMigration-NewMigration.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'New migration script failed to parse.' }
foreach ($name in @('Assert-WebUrl', 'Get-CommonPermissionPath', 'Get-SPOAuthParameters', 'ConvertTo-Psd1Value', 'Assert-MigrationRequest', 'Invoke-NewMigration')) {
    $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if (-not $definition) { throw "Missing function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-NewMigrationProfiles-' + [guid]::NewGuid().ToString('N'))
try {
    $projectRoot = $fixture
    $migrationsRoot = Join-Path $fixture 'Migrations'
    $template = Join-Path $fixture 'Migrations_Template'
    New-Item -ItemType Directory -Path $migrationsRoot, $template -Force | Out-Null
    "@{ Name='NewMigration'; Source=@{}; Target=@{} }" | Set-Content -LiteralPath (Join-Path $template 'migration.config.psd1')
    $source = Join-Path $fixture 'source.local.psd1'; $target = Join-Path $fixture 'target.local.psd1'
    "@{ ClientId='source-app'; Tenant='source.example.test' }" | Set-Content -LiteralPath $source
    "@{ ClientId='target-app' }" | Set-Content -LiteralPath $target
    function Import-Module { param($Name, $ErrorAction) }
    $script:checks = [Collections.Generic.List[string]]::new()
    function Test-SPOWeb { param($Url, $Label, $AuthParameters) $script:checks.Add("$Label|$($AuthParameters.ClientId)"); if ($Label -eq 'Source web' -and -not $AuthParameters.DeviceLogin) { throw 'Source validation did not select DeviceLogin.' }; if ($Label -eq 'Target web' -and -not $AuthParameters.Interactive) { throw 'Target validation did not select Interactive.' } }
    function Test-SPOAdmin { param($Url, $AuthParameters, $Label = 'Target admin URL') $script:checks.Add("$Label|$($AuthParameters.ClientId)") }
    $request = [pscustomobject]@{
        Name='Synthetic'; SourceType='SPO'; TargetType='SPO'; AuthMode='Interactive'
        SourceAuthMode='DeviceLogin'; TargetAuthMode='Interactive'
        SourceAuthConfigPath=$source; TargetAuthConfigPath=$target
        SourceAdminUrl='https://source-admin.example.test'; TargetAdminUrl='https://target-admin.example.test'
        Mappings=@([pscustomobject]@{ SourceUrl='https://source.example.test/sites/a'; TargetUrl='https://target.example.test/sites/b' })
    }
    $result = Invoke-NewMigration $request
    if (-not $result.Success -or $result.ValidatedUrls -ne 4) { throw 'Migration was not validated.' }
    foreach ($expected in @('Source web|source-app', 'Target web|target-app', 'Source admin URL|source-app', 'Target admin URL|target-app')) {
        if ($script:checks -notcontains $expected) { throw "Missing side-specific validation: $expected" }
    }
    $saved = Import-PowerShellDataFile (Join-Path $result.Folder 'migration.config.psd1')
    if ($saved.Source.AuthConfigPath -ne $source -or $saved.Target.AuthConfigPath -ne $target -or $saved.Source.TenantAdminUrl -ne $request.SourceAdminUrl -or $saved.Source.AuthMode -ne 'DeviceLogin' -or $saved.Target.AuthMode -ne 'Interactive') {
        throw 'Side-specific configuration was not persisted.'
    }
    function Test-OnPremWeb { param($Url, $Label) }
    $legacyRequest = [pscustomobject]@{
        Name='Legacy'; SourceType='SP2019'; TargetType='SPO'; AuthMode='Interactive'
        TargetAdminUrl='https://target-admin.example.test'
        Mappings=@([pscustomobject]@{ SourceUrl='https://farm.example.test/sites/a'; TargetUrl='https://target.example.test/sites/b' })
    }
    $legacy = Invoke-NewMigration $legacyRequest
    $legacyConfig = Import-PowerShellDataFile (Join-Path $legacy.Folder 'migration.config.psd1')
    if (-not $legacy.Success -or $legacyConfig.Source.Type -ne 'SP2019' -or $legacyConfig.Target.AuthConfigPath) {
        throw 'Legacy request without side-specific fields failed.'
    }
    $script:checks.Clear()
    $request.Name = 'SitesOnly'
    $request.SourceAdminUrl = ''
    $request.TargetAdminUrl = ''
    $sitesOnly = Invoke-NewMigration $request
    $sitesConfig = Import-PowerShellDataFile (Join-Path $sitesOnly.Folder 'migration.config.psd1')
    if (-not $sitesOnly.Success -or $sitesOnly.ValidatedUrls -ne 2 -or $sitesConfig.Target.TenantAdminUrl -or
        @($script:checks | Where-Object { $_ -match 'admin' }).Count) {
        throw 'A site-only SPO migration must validate both webs without requiring a tenant admin endpoint.'
    }
    'PASS: independent source and target profile validation and configuration.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($fixture)
    $allowed = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-NewMigrationProfiles-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
