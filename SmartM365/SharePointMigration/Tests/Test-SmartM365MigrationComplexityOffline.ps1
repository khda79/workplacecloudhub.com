<#
.SYNOPSIS
    Verify the migration complexity request and presentation without SharePoint access.
.VERSION
    1.0.0
#>

#Requires -Version 7.4

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $projectRoot 'SmartM365-SharePointMigration-Summary.ps1')

$python = Join-Path $projectRoot 'Tools\Python\python.exe'
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { $python = (Get-Command python -ErrorAction Stop).Source }

$safeRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$testRoot = [IO.Path]::GetFullPath((Join-Path $safeRoot ('SharePointMigration-complexity-tests-' + [guid]::NewGuid().ToString('N'))))
if (-not $testRoot.StartsWith($safeRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Test directory is outside the temporary folder.' }

function Assert-Complexity {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Complexity test failed: $Message" }
}

try {
    $fileDir = Join-Path $testRoot 'scans\source\files'
    $permissionDir = Join-Path $testRoot 'scans\source\permissions'
    [void](New-Item -ItemType Directory -Path $fileDir, $permissionDir -Force)
    $fileScan = Join-Path $fileDir 'SP2019-FileInventory-Fixture-20261001-120000.csv'
    $permissionScan = Join-Path $permissionDir 'SP2019-PermissionInventory-Fixture-20261001-120000.csv'
    @(
        '"WebUrl";"LibraryUrl";"FileName";"ServerRelativeUrl";"SizeBytes";"VersionsCount";"CheckedOutBy"'
        '"https://source.test";"/Docs";"a.docx";"/Docs/a.docx";"100";"1";""'
    ) | Set-Content -LiteralPath $fileScan -Encoding utf8
    @(
        '"ObjectScope";"ObjectId";"HasUniqueRoleAssignments";"PrincipalType";"PrincipalLoginName"'
        '"Item";"1";"True";"User";"i:0e.t|adfs|user@contoso.test"'
    ) | Set-Content -LiteralPath $permissionScan -Encoding utf8

    $migration = [pscustomobject]@{
        Name = 'Fixture'; Root = $testRoot
        Config = @{ Name = 'Fixture'; Target = @{ SiteUrl = 'https://target.test/sites/fixture' }
            Comparison = @{ EntraUsersCachePath = 'missing-entra-cache.csv' } }
    }
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $null).State -eq 'Unavailable') 'no source scan must be unavailable'

    $request = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot `
        -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ($null -eq $request.EntraCache) 'a configured but missing Entra cache must not be used'
    $before = Get-SmartM365ComplexityPresentation -Request $request
    Assert-Complexity ($before.State -eq 'Missing' -and $before.Text -eq 'Computing') 'missing result must request a calculation'
    $readOnly = Get-SmartM365ComplexityPresentation -Request $request -ReadOnly
    Assert-Complexity ($readOnly.Caption -eq 'Not evaluated') 'read-only views must not announce a calculation'

    & $python @($request.Arguments) | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Complexity calculation failed with exit code $LASTEXITCODE." }
    $request = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot `
        -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    $current = Get-SmartM365ComplexityPresentation -Request $request
    Assert-Complexity ($current.State -eq 'Current' -and $current.Text -eq 'Low') 'calculated result must be current'
    Assert-Complexity ($current.Caption -like '*/100*partial') 'missing Entra cache must be shown as partial'
    Assert-Complexity (@($current.Criteria).Count -eq 11) 'all criteria must be listed'
    Assert-Complexity ((@($current.Criteria) | Where-Object Criterion -eq 'Disabled or unresolved accounts').Points -eq 'N/A') 'unevaluated criterion must show N/A'

    Add-Content -LiteralPath $fileScan -Value '"https://source.test";"/Docs";"b.docx";"/Docs/b.docx";"100";"0";""' -Encoding utf8
    $newer = Get-SmartM365ComplexityRequest -Migration $migration -ProjectRoot $projectRoot `
        -FileScan (Get-Item $fileScan) -PermissionScan (Get-Item $permissionScan)
    Assert-Complexity ($newer.Signature -ne $request.Signature) 'a changed scan must change the signature'
    $stale = Get-SmartM365ComplexityPresentation -Request $newer
    Assert-Complexity ($stale.State -eq 'Stale' -and $stale.Caption -eq 'Updating' -and $stale.Text -eq 'Low') 'stale result must keep the previous level'
    Assert-Complexity ((Get-SmartM365ComplexityPresentation -Request $newer -ReadOnly).Caption -eq 'Outdated') 'read-only stale result must be outdated'
    $failed = Get-SmartM365ComplexityPresentation -Request $newer -ErrorMessage 'boom'
    Assert-Complexity ($failed.State -eq 'Error' -and $failed.Tooltip -like '*boom*') 'failure must be visible'

    Write-Host ('{0} Migration complexity offline checks passed.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
}
finally {
    if ((Split-Path $testRoot -Leaf) -like 'SharePointMigration-complexity-tests-*' -and (Test-Path -LiteralPath $testRoot)) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
