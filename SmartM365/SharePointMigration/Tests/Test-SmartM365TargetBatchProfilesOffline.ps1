<#
.SYNOPSIS
    Verify destination batch certificate profile selection without certificates or network access.
.VERSION
    1.0.0
#>
#requires -Version 7.4
$ErrorActionPreference = 'Stop'
$sourcePath = Join-Path $PSScriptRoot '..\Scripts\Launchers\Generic\SmartM365-SharePointMigration-TargetScanBatch.ps1'
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Batch script failed to parse.' }
$definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Assert-CertificateAuth' }, $true)
if (-not $definition) { throw 'Certificate preflight not found.' }
. ([scriptblock]::Create($definition.Extent.Text))

$fixture = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-BatchProfiles-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path (Join-Path $fixture 'Config'), (Join-Path $fixture 'Migration') -Force | Out-Null
    "@{ ClientId='app'; TenantId='tenant'; Thumbprint='LEGACY' }" | Set-Content (Join-Path $fixture 'Config\SPOAuth.local.psd1')
    "@{ ClientId='app'; TenantId='tenant'; Thumbprint='SEPARATE' }" | Set-Content (Join-Path $fixture 'Migration\target.local.psd1')
    function Get-ChildItem {
        param($Path, $ErrorAction)
        @('LEGACY', 'SEPARATE') | ForEach-Object { [pscustomobject]@{ Thumbprint=$_; HasPrivateKey=$true; NotAfter=(Get-Date).AddDays(1) } }
    }
    $migrations = @(
        [pscustomobject]@{ Name='Old'; Root=(Join-Path $fixture 'Migration'); AuthConfigPath='' },
        [pscustomobject]@{ Name='New'; Root=(Join-Path $fixture 'Migration'); AuthConfigPath='target.local.psd1' }
    )
    Assert-CertificateAuth -Root $fixture -Migrations $migrations
    $migrations[1].AuthConfigPath = 'missing.local.psd1'
    $failed = $false
    try { Assert-CertificateAuth -Root $fixture -Migrations $migrations } catch { $failed = $true }
    if (-not $failed) { throw 'Missing explicit profile incorrectly fell back to the shared profile.' }
    Write-Output 'PASS: shared and per-migration certificate profiles; missing explicit profile rejected.'
}
finally {
    Remove-Item function:Get-ChildItem -ErrorAction SilentlyContinue
    $resolved = [IO.Path]::GetFullPath($fixture)
    $allowed = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-BatchProfiles-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
