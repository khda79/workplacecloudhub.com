$ErrorActionPreference = 'Stop'
$generator = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/New-SmartM365-OrchestratorJobLaunchers.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($generator, [ref]$null, [ref]$null)
$loops = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.ForEachStatementAst] -and $node.Extent.Text -like 'foreach ($relative in @($existing*' }, $true))
if ($loops.Count -ne 1) { throw 'Cannot find the generator obsolete-launcher loop.' }
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-StandaloneLaunchers-' + [guid]::NewGuid().ToString('N'))
try {
    $paths = @('Launchers\OnPremises\Start-SmartM365-Windows11-HardwareReadiness-Inventory.cmd', 'Launchers\Orchestrator\Start-SmartM365-Inventory-OrchestratorScheduledTask-Uninstaller.cmd', 'Launchers\Orchestrator\Obsolete.cmd')
    foreach ($relative in $paths) {
        $path = Join-Path $root $relative
        $null = New-Item -ItemType Directory -Path (Split-Path $path -Parent) -Force
        [IO.File]::WriteAllText($path, 'synthetic')
    }
    & {
        param($Code, $Root, $Paths)
        $smartInventoryRoot = $Root
        $existing = $Paths
        $expected = [ordered]@{}
        $standalone = [ordered]@{ ($Paths[0]) = @{}; ($Paths[1]) = @{} }
        $differences = [Collections.Generic.List[string]]::new()
        $Check = $false
        . ([scriptblock]::Create($Code))
        if ((Test-Path -LiteralPath (Join-Path $Root $Paths[2])) -or $differences.Count -ne 1) { throw 'The obsolete generated launcher was not removed.' }
        foreach ($relative in $Paths[0..1]) {
            if (-not (Test-Path -LiteralPath (Join-Path $Root $relative))) { throw "Standalone source launcher removed: $relative" }
        }
    } $loops[0].Extent.Text $root $paths
    'PASS: generator removes obsolete aliases while preserving standalone launchers'
} finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }
