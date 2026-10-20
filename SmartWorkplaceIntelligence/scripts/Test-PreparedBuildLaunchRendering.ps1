[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Version 1.0.0: the generator build process runs with NO_COLOR=1, restored afterwards.
# Executes only the extracted build-launch statements against a synthetic build script; no preparation.
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ('build-launch-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture -Force | Out-Null
$script:checks = 0
function Check([bool]$Pass, [string]$Message) {
    if (-not $Pass) { throw $Message }; $script:checks++
}
$savedNoColor = $env:NO_COLOR
try {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1'), [ref]$tokens, [ref]$errors)
    Check ($errors.Count -eq 0) 'Pipeline parse failed'
    $launch = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.TryStatementAst] -and $n.Body.Extent.Text -match 'Invoke-PreparedEvidenceBuild\.ps1' -and $n.Finally -and $n.Finally.Extent.Text -match 'NO_COLOR' }, $true))
    Check ($launch.Count -eq 1) 'Expected one guarded build launch'
    Check ($launch[0].Finally.Extent.Text -match '\$previousNoColor') 'NO_COLOR is not restored in finally'
    $block = @($launch[0].Parent.Statements)
    $index = [array]::IndexOf($block, $launch[0])
    Check ($index -ge 2 -and $block[$index - 2].Extent.Text -match '^\$previousNoColor\s*=' -and $block[$index - 1].Extent.Text -match "^\`$env:NO_COLOR = '1'$") 'NO_COLOR is not set immediately before the build launch'
    $launchText = ($block[($index - 2)..$index] | ForEach-Object { $_.Extent.Text }) -join [Environment]::NewLine

    # Synthetic build: reports what the child sees and emits one warning with default rendering.
    @'
param($DataRoot, $OutputRoot, $AccountClassificationConfigPath, $AdCountryRulesPath, $MaxParallelGenerators)
Write-Host ('child NO_COLOR=' + $env:NO_COLOR + '; rendering=' + $PSStyle.OutputRendering)
Write-Warning 'Synthetic generator warning'
'@ | Set-Content -LiteralPath (Join-Path $fixture 'Invoke-PreparedEvidenceBuild.ps1')
    @(
        'param([string]$LogPath)'
        '$snapshot = "snapshot"; $staging = "staging"; $rules = "rules"; $countryRulesPath = "country"'
        '$MaxParallelGenerators = 3; $systemPolicyArgs = @(); $log = $LogPath'
        $launchText
        '$LASTEXITCODE'
    ) | Set-Content -LiteralPath (Join-Path $fixture 'launch.ps1')
    $launchScript = Join-Path $fixture 'launch.ps1'

    # Success: the child sees NO_COLOR=1, the retained log has the plain WARNING prefix, the caller value returns.
    $env:NO_COLOR = 'caller-value'
    $log = Join-Path $fixture 'preparation.log'
    $exitCode = & $launchScript -LogPath $log 6>$null
    $logged = Get-Content -LiteralPath $log -Raw
    Check ($exitCode -eq 0) 'Synthetic build failed'
    Check ($logged -match 'child NO_COLOR=1; rendering=PlainText') 'Build child did not run with NO_COLOR=1'
    Check ($logged -match '\] (?:WARNING|AVERTISSEMENT)\s*: Synthetic generator warning') 'Generator warning lost its plain WARNING prefix'
    Check (-not $logged.Contains([string][char]27)) 'Retained build log contains ANSI sequences'
    Check ($env:NO_COLOR -ceq 'caller-value') 'Caller NO_COLOR value was not restored'

    # Failure inside the relay: NO_COLOR is still restored (unset before, unset after).
    Remove-Item -LiteralPath Env:\NO_COLOR -ErrorAction SilentlyContinue
    $failed = $false
    try { & $launchScript -LogPath (Join-Path $fixture 'missing\preparation.log') 6>$null | Out-Null } catch { $failed = $true }
    Check $failed 'Relay failure was not raised'
    Check (-not (Test-Path -LiteralPath Env:\NO_COLOR)) 'NO_COLOR leaked after a relay failure'
    Write-Output "PASS: $script:checks build launch rendering checks. Synthetic build only."
} finally {
    if ($null -eq $savedNoColor) { Remove-Item -LiteralPath Env:\NO_COLOR -ErrorAction SilentlyContinue } else { $env:NO_COLOR = $savedNoColor }
    Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue
}
