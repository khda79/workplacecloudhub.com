#Requires -Version 5.1
<#
.SYNOPSIS
Offline tests of SmartInventoryConfigRootPath: path mapping, templates, fail-closed reads and the switch tool.
.VERSION
1.0.0
.NOTES
Builds a synthetic SmartM365 tree under TEMP (copies of the JSON transport module, the tenant context, the
account-classification module and the switch tool) and runs the checks in Windows PowerShell 5.1 and
PowerShell 7 when both are available. No real configuration, tenant, process or scheduled task is touched:
the switch tool copy sees no process, no task and an elevated session.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))

$inner = @'
param([string]$Root)
$ErrorActionPreference = 'Stop'
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Reject([scriptblock]$Action, [string]$Message) { $failed = $false; try { & $Action | Out-Null } catch { $failed = $true }; Check $failed $Message }
function Get-Long([string]$Path) { (Resolve-Path -LiteralPath $Path).ProviderPath }
$ws = Join-Path (Split-Path $Root -Parent) 'WS'
$code = Join-Path $Root 'SmartInventory'
$globalPath = Join-Path $Root 'Config\SmartM365.global.local.json.txt'
function Set-Global([string]$Value) {
    $document = [ordered]@{ WorkspaceRootPath = $ws; SmartInventoryConfigRootPath = $Value }
    [IO.File]::WriteAllText($globalPath, ($document | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
}
Set-Global ''
. (Join-Path $Root 'Config\SmartM365-TenantContext.ps1')
$foo = Join-Path $code 'Area\SmartM365-Foo.local.json'

Check ((ConvertTo-SmartM365InventoryConfigLocation $foo) -eq $foo) 'Empty setting changed the path.'
Check ((Resolve-SmartM365JsonConfigurationPath $foo) -eq "$foo.txt") 'Empty setting changed the resolution.'
Check ((Get-SmartM365JsonTemplateName "$foo.txt") -eq "$foo.template") 'Empty setting changed the template name.'

Set-Global '{{WorkspaceRootPath}}\SmartInventory\Config'
$target = Join-Path $ws 'SmartInventory\Config\Area\SmartM365-Foo.local.json'
Check ((ConvertTo-SmartM365InventoryConfigLocation $foo) -eq $target) 'Mapping to the configuration root is wrong.'
Check (Test-Path -LiteralPath (Split-Path $target -Parent) -PathType Container) 'Target folder was not created.'
Check ((ConvertTo-SmartM365InventoryConfigLocation $target) -eq $target) 'Mapping is not idempotent.'
Check ((ConvertFrom-SmartM365InventoryConfigLocation "$target.txt") -eq "$foo.txt") 'Inverse mapping is wrong.'
Check ((Get-SmartM365JsonTemplateName "$target.txt") -eq "$foo.template") 'Module template is not read from the code folder.'
Check ((Get-SmartM365JsonTemplatePath "$target.txt") -eq "$foo.template") 'Tenant context template is not read from the code folder.'
$config = Read-SmartM365JsonConfig -Path $foo -Required
Check ($config.Setting -eq 1) 'First run did not use the code folder template.'
Check ((Test-Path -LiteralPath "$target.txt") -and -not (Test-Path -LiteralPath "$foo.txt") -and -not (Test-Path -LiteralPath $foo)) 'First run created the file in the wrong folder.'
[IO.File]::WriteAllText("$foo.template", '{"Setting":1,"Added":2}')
$null = Read-SmartM365JsonConfig -Path $foo -Required
Check (((Get-Content -LiteralPath "$target.txt" -Raw | ConvertFrom-Json).Added) -eq 2) 'Template enrichment was not written to the configuration root.'
$bar = Join-Path $code 'Area\SmartM365-Bar.local.json'
[IO.File]::WriteAllText("$bar.txt", '{"A":1}')
Reject { Resolve-SmartM365JsonConfigurationPath $bar } 'A file left in the code folder was read from a second source.'
Remove-Item -LiteralPath "$bar.txt"
$tenantPath = Join-Path $Root 'Config\Tenants\x.local.json'
Check ((ConvertTo-SmartM365InventoryConfigLocation $tenantPath) -eq $tenantPath) 'A path outside SmartInventory was mapped.'
Set-Global ($code + '\Private'); Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'A root inside the code folder was accepted.'
Set-Global '{{DataAllRootPath}}\Config'; Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'An unsupported token was accepted.'
Set-Global 'Relative\Config'; Reject { ConvertTo-SmartM365InventoryConfigLocation $foo } 'A relative root was accepted.'

if ($PSVersionTable.PSVersion.Major -ge 7) {
    # The account-classification module requires PowerShell 7.
    Set-Global '{{WorkspaceRootPath}}\SmartInventory\Config'
    Import-Module (Join-Path $code 'Common\SmartM365.AccountClassification.psd1') -Force
    $classification = Join-Path $ws 'SmartInventory\Config\Config\AccountClassification.local.json.txt'
    New-Item -ItemType Directory -Path (Split-Path $classification -Parent) -Force | Out-Null
    [IO.File]::WriteAllText($classification, '{}')
    Check ((Resolve-SmartM365AccountClassificationPath) -eq (Get-Long $classification)) 'Account classification was not read from the configuration root.'
}

# Switch tool: preview, apply, rollback on a fresh tree.
Remove-Item -LiteralPath $ws -Recurse -Force
Set-Global ''
$area = Join-Path $code 'Area'
foreach ($name in 'SmartM365-Foo.local.json', 'SmartM365-Foo.Registry.local.json',
    'SmartM365-Old.local.json', 'SmartM365-Foo.local.json.txt.bak-20260101', 'SmartM365-Foo.local.json.txt.0123abcd.pending') {
    [IO.File]::WriteAllText((Join-Path $area $name), ('{"Name":"' + $name + '"}'))
}
# A real .json -> .json.txt migration creates the transport lock and journal that the tool must move.
$null = Resolve-SmartM365JsonConfigurationPath $foo
Check ((Test-Path -LiteralPath "$foo.txt") -and (Test-Path -LiteralPath "$foo.migration.log") -and (Test-Path -LiteralPath "$foo.transport.lock")) 'Fixture migration did not run.'
[IO.File]::WriteAllText((Join-Path $code 'Config\AccountClassification.local.json.txt'), '{"Rules":1}')
function global:Get-CimInstance { @() }
function global:Get-ScheduledTask { @() }
$tool = Join-Path $code 'Maintenance\SmartM365-Move-InventoryConfigToRoot.ps1'
$before = @(Get-ChildItem -LiteralPath $code -Recurse -File | ForEach-Object { $_.FullName + '|' + (Get-FileHash -LiteralPath $_.FullName).Hash })
& $tool -Tenant test *> $null
$after = @(Get-ChildItem -LiteralPath $code -Recurse -File | ForEach-Object { $_.FullName + '|' + (Get-FileHash -LiteralPath $_.FullName).Hash })
Check ((Compare-Object $before $after).Count -eq 0 -and -not (Test-Path -LiteralPath $ws)) 'Preview changed files.'
Reject { & $tool -Tenant test -Apply *> $null } 'Apply ran without -OtherServersStopped.'
& $tool -Tenant test -Apply -OtherServersStopped *> $null
$root = Join-Path $ws 'SmartInventory\Config'
foreach ($relative in 'Area\SmartM365-Foo.local.json.txt', 'Area\SmartM365-Foo.local.json.transport.lock', 'Area\SmartM365-Foo.local.json.migration.log', 'Area\SmartM365-Foo.Registry.local.json', 'Config\AccountClassification.local.json.txt') {
    Check (Test-Path -LiteralPath (Join-Path $root $relative) -PathType Leaf) "Not moved: $relative"
    Check (@(Get-ChildItem -LiteralPath (Split-Path (Join-Path $code $relative) -Parent) -Filter ((Split-Path $relative -Leaf) + '.moved-*')).Count -eq 1) "Original not renamed: $relative"
    Check (-not (Test-Path -LiteralPath (Join-Path $code $relative))) "Original still active: $relative"
}
Check (-not (Test-Path -LiteralPath (Join-Path $area 'SmartM365-Old.local.json')) -and -not (Test-Path -LiteralPath (Join-Path $root 'Area\SmartM365-Old.local.json'))) 'Obsolete configuration was not archived.'
Check (-not (Test-Path -LiteralPath (Join-Path $area 'SmartM365-Foo.local.json.txt.bak-20260101'))) 'Manual backup was not archived.'
Check (Test-Path -LiteralPath (Join-Path $area 'SmartM365-Foo.local.json.txt.0123abcd.pending')) 'Interrupted replacement was not kept in place.'
$zip = @(Get-ChildItem -LiteralPath (Join-Path $ws 'SmartInventory\Archive') -Filter '*.zip')
Check ($zip.Count -eq 1) 'ZIP backup missing.'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::OpenRead($zip[0].FullName)
try { Check ($archive.Entries.Count -eq 8) ('ZIP backup entry count: ' + $archive.Entries.Count) } finally { $archive.Dispose() }
Check ((Get-Content -LiteralPath $globalPath -Raw | ConvertFrom-Json).SmartInventoryConfigRootPath -eq '{{WorkspaceRootPath}}\SmartInventory\Config') 'Setting was not activated.'
Check ((Read-SmartM365JsonConfig -Path $foo -Required).Name -eq 'SmartM365-Foo.local.json') 'Moved configuration is not the one read.'
& $tool -Tenant test -Rollback -OtherServersStopped *> $null
Check ((Get-Content -LiteralPath $globalPath -Raw | ConvertFrom-Json).SmartInventoryConfigRootPath -eq '') 'Rollback did not clear the setting.'
Check ((Test-Path -LiteralPath "$foo.txt") -and @(Get-ChildItem -LiteralPath $code -Recurse -Filter '*.moved-*').Count -eq 0) 'Rollback did not restore the originals.'
Check ((Read-SmartM365JsonConfig -Path $foo -Required).Name -eq 'SmartM365-Foo.local.json') 'Rolled back configuration is not the one read.'
'PASSED {0}' -f $script:passed
'@

function New-Fixture {
    $base = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-InvCfgRoot-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $root = Join-Path $base 'SmartM365'
    $copies = @(
        'Modules\SmartM365.Core\SmartM365.JsonTransport.psd1', 'Modules\SmartM365.Core\SmartM365.JsonTransport.psm1',
        'Config\SmartM365-JsonTransport.policy.psd1', 'Config\SmartM365-TenantContext.ps1', 'Config\SmartM365.global.local.json.template',
        'Config\Tenants\tenant.local.json.template', 'Config\Tenants\tenant.mail.local.json.template',
        'SmartInventory\Common\SmartM365.AccountClassification.psd1', 'SmartInventory\Common\SmartM365.AccountClassification.psm1',
        'SmartInventory\Config\AccountClassification.local.json.template',
        'SmartInventory\Maintenance\SmartM365-Move-InventoryConfigToRoot.ps1')
    foreach ($relative in $copies) {
        $destination = Join-Path $root $relative
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $repoRoot $relative) -Destination $destination
    }
    # The copy runs unelevated: the elevation check is the only line changed.
    $tool = Join-Path $root 'SmartInventory\Maintenance\SmartM365-Move-InventoryConfigToRoot.ps1'
    $text = [IO.File]::ReadAllText($tool)
    $check = '$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)'
    if (-not $text.Contains($check)) { throw 'Switch tool elevation check not found.' }
    [IO.File]::WriteAllText($tool, $text.Replace($check, '$true'), (New-Object Text.UTF8Encoding($true)))
    [IO.File]::WriteAllText((Join-Path $root 'Config\Tenants\test.local.json.txt'), '{"ProfileKey":"test","OrganizationKey":"contoso","EnvironmentKey":"test","TenantKey":"contoso-test"}')
    $area = Join-Path $root 'SmartInventory\Area'
    New-Item -ItemType Directory -Path $area -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $area 'SmartM365-Foo.ps1'), '# owner script')
    [IO.File]::WriteAllText((Join-Path $area 'SmartM365-Foo.local.json.template'), '{"Setting":1}')
    $innerPath = Join-Path $base 'inner.ps1'
    [IO.File]::WriteAllText($innerPath, $inner, (New-Object Text.UTF8Encoding($true)))
    [pscustomobject]@{ Base = $base; Root = $root; Inner = $innerPath }
}

$engines = @()
$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (Test-Path -LiteralPath $windowsPowerShell) { $engines += $windowsPowerShell }
$pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
if ($pwsh) { $engines += $pwsh.Source }
if (-not $engines.Count) { throw 'No PowerShell engine found.' }
foreach ($engine in $engines) {
    $fixture = New-Fixture
    try {
        $output = & $engine -NoProfile -ExecutionPolicy Bypass -File $fixture.Inner -Root $fixture.Root 2>&1
        $text = ($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine
        if ($LASTEXITCODE -ne 0 -or $text -notmatch 'PASSED \d+') { throw ("{0} failed:{1}{2}" -f $engine, [Environment]::NewLine, $text) }
        Write-Host ('{0}: {1}' -f (Split-Path $engine -Leaf), ([regex]::Match($text, 'PASSED \d+').Value)) -ForegroundColor Green
    }
    finally { Remove-Item -LiteralPath $fixture.Base -Recurse -Force -ErrorAction SilentlyContinue }
}
