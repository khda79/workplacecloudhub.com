<#
.SYNOPSIS
Offline tests of the completion banner for interrupted SmartInventory runs.

.DESCRIPTION
No collector, tenant, network or mail is used. Covers:
- Config\SmartM365-TenantContext.ps1 banner: Success, CompletedWithWarnings and Failed unchanged, with or
  without a lifecycle; Auto resolved from the counts as before; a run interrupted before its end (lifecycle
  in cleanup, terminal not reached) is CANCELLED with exit code 130, or FAILED when errors were counted;
- lifecycle finalizer: with both the TenantContext banner and the prefixed Core banner loaded, the deferred
  banner is the Core one (CANCELLED, no cleanup failure); with the TenantContext banner only, it still shows
  CANCELLED;
- entry scripts that call the TenantContext banner from their own finally: their real catch and finally
  blocks are run in a Windows PowerShell child process with stubs for every other command; an interruption
  gives CANCELLED and exit code 130, a completed run keeps SUCCESS and exit code 0.

.VERSION
1.0.0
#>
[CmdletBinding()]
param([string]$SourceRoot)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }

$script:Checks = 0
function Assert-Banner {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

function Get-ScriptAst {
    param([string]$Path)
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw "Parse failed: $Path" }
    return $ast
}

function Get-FunctionText {
    param($Ast, [string]$Name)
    $node = $Ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $node) { throw "Missing function: $Name" }
    return $node.Extent.Text
}

$tenantContextPath = Join-Path $SourceRoot 'Config/SmartM365-TenantContext.ps1'
$lifecyclePath = Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1'
$corePath = Join-Path $SourceRoot 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psd1'
$bannerText = Get-FunctionText (Get-ScriptAst $tenantContextPath) 'Write-SmartM365CompletionBanner'
$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Banner-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

function Invoke-ChildScript {
    param([string]$Name, [string[]]$Lines)
    $path = Join-Path $tempRoot "$Name.ps1"
    [IO.File]::WriteAllText($path, ($Lines -join "`r`n"), [Text.UTF8Encoding]::new($false))
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $output = (& $windowsPowerShell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $path 2>&1 | Out-String) }
    finally { $ErrorActionPreference = $previous }
    return [pscustomobject]@{ Output = $output; ExitCode = $LASTEXITCODE }
}

try {
    # 1. TenantContext banner in-process: completed statuses unchanged, interrupted run Cancelled.
    # The signed lifecycle file is dot-sourced, never concatenated: its signature block must stay last.
    $bannerModule = New-Module -ScriptBlock ([scriptblock]::Create((". '{0}'" -f $lifecyclePath.Replace("'", "''")) + "`n" + $bannerText))
    $render = {
        param([string]$Status, [int]$Warnings, [int]$Errors)
        Remove-Variable -Name SmartM365CompletionBannerRunKey -Scope Global -ErrorAction SilentlyContinue
        $start = [datetime]'2026-10-21T10:00:00'
        (& $bannerModule { param($s, $w, $e, $t) Write-SmartM365CompletionBanner -Status $s -ScriptName 'Probe' -StartedAt $t -EndedAt $t.AddSeconds(65) -WarningCount $w -ErrorCount $e -GeneratedCsvFiles 2 6>&1 } $Status $Warnings $Errors $start | Out-String)
    }
    $expected = @(
        @{ Status = 'Success'; Warnings = 0; Errors = 0; Title = 'Execution completed'; Label = 'SUCCESS' },
        @{ Status = 'CompletedWithWarnings'; Warnings = 1; Errors = 0; Title = 'Execution completed with warnings'; Label = 'COMPLETED WITH WARNINGS' },
        @{ Status = 'Failed'; Warnings = 0; Errors = 1; Title = 'Execution failed'; Label = 'FAILED' },
        @{ Status = 'Auto'; Warnings = 0; Errors = 0; Title = 'Execution completed'; Label = 'SUCCESS' },
        @{ Status = 'Auto'; Warnings = 2; Errors = 0; Title = 'Execution completed with warnings'; Label = 'COMPLETED WITH WARNINGS' },
        @{ Status = 'Auto'; Warnings = 0; Errors = 3; Title = 'Execution failed'; Label = 'FAILED' }
    )
    foreach ($lifecycleMode in @('None', 'Completed')) {
        foreach ($case in $expected) {
            & $bannerModule {
                param($mode)
                Remove-Variable -Name SmartM365InventoryRun, SmartM365ErrorCount -Scope Global -ErrorAction SilentlyContinue
                if ($mode -eq 'Completed') {
                    $state = Start-SmartM365InventoryRun -ScriptPath 'C:\probe.ps1' -RequireReturn
                    $state.ReachedTerminal = $true
                    $state.InCleanup = $true
                }
            } $lifecycleMode
            $text = & $render $case.Status $case.Warnings $case.Errors
            $expectedLines = @(
                (' SmartM365 by WorkplaceCloudHub - {0}' -f $case.Title),
                ' Script    : Probe',
                (' Status    : {0}' -f $case.Label),
                ' Duration  : 00:01:05',
                (' Warnings  : {0}' -f $case.Warnings),
                (' Errors    : {0}' -f $case.Errors),
                ' CSV files : 2'
            )
            $missing = @($expectedLines | Where-Object { -not $text.Contains($_) })
            Assert-Banner ($missing.Count -eq 0) ("{0} with {1} warning(s) and {2} error(s), lifecycle {3}: unchanged banner (missing: {4})." -f $case.Status, $case.Warnings, $case.Errors, $lifecycleMode, ($missing -join ' | '))
        }
    }

    $interrupted = & $bannerModule {
        Remove-Variable -Name SmartM365InventoryRun, SmartM365ErrorCount -Scope Global -ErrorAction SilentlyContinue
        $state = Start-SmartM365InventoryRun -ScriptPath 'C:\probe.ps1' -RequireReturn
        $state.InCleanup = $true
        $state
    }
    $cancelledText = & $render 'Auto' 0 0
    Assert-Banner ($cancelledText.Contains(' SmartM365 by WorkplaceCloudHub - Execution cancelled') -and $cancelledText.Contains(' Status    : CANCELLED') -and $cancelledText -cnotmatch 'SUCCESS' -and $interrupted.ExitCode -eq 130) "An interrupted run is CANCELLED with exit code 130 (exit $($interrupted.ExitCode))."
    & $bannerModule {
        Remove-Variable -Name SmartM365InventoryRun -Scope Global -ErrorAction SilentlyContinue
        $state = Start-SmartM365InventoryRun -ScriptPath 'C:\probe.ps1' -RequireReturn
        $state.InCleanup = $true
    }
    Assert-Banner ((& $render 'Success' 0 0).Contains(' Status    : CANCELLED')) 'A run that never reached its end is CANCELLED even when Success was passed.'
    & $bannerModule {
        Remove-Variable -Name SmartM365InventoryRun -Scope Global -ErrorAction SilentlyContinue
        $state = Start-SmartM365InventoryRun -ScriptPath 'C:\probe.ps1' -RequireReturn
        $state.InCleanup = $true
    }
    Assert-Banner ((& $render 'Failed' 0 1).Contains(' Status    : FAILED')) 'A failed run that did not reach its end stays FAILED.'
    & $bannerModule {
        Remove-Variable -Name SmartM365InventoryRun -Scope Global -ErrorAction SilentlyContinue
        $state = Start-SmartM365InventoryRun -ScriptPath 'C:\probe.ps1' -RequireReturn
        $state.InCleanup = $true
        $global:SmartM365ErrorCount = 2
    }
    Assert-Banner ((& $render 'Auto' 0 2).Contains(' Status    : FAILED')) 'An interrupted run with counted errors is FAILED.'
    & $bannerModule {
        Remove-Variable -Name SmartM365ErrorCount -Scope Global -ErrorAction SilentlyContinue
        Remove-Variable -Name SmartM365InventoryRun -Scope Global -ErrorAction SilentlyContinue
        $state = Start-SmartM365InventoryRun -ScriptPath 'C:\probe.ps1' -RequireReturn
    }
    Assert-Banner ((& $render 'Auto' 0 0).Contains(' Status    : SUCCESS')) 'Outside the cleanup the banner does not consult the lifecycle (no side effect).'
    & $bannerModule { Remove-Variable -Name SmartM365InventoryRun, SmartM365CompletionBannerRunKey -Scope Global -ErrorAction SilentlyContinue }
    Assert-Banner ((& $render 'Cancelled' 0 0).Contains(' Status    : CANCELLED')) 'Cancelled can be passed explicitly.'

    # 2. Lifecycle finalizer with both banners loaded, and with the TenantContext banner only.
    $lifecycleText = [IO.File]::ReadAllText($lifecyclePath)
    Assert-Banner ($lifecycleText.Contains('Get-Command Write-CoreSmartM365CompletionBanner,Write-SmartM365CompletionBanner')) 'The finalizer looks for the Core banner first.'
    foreach ($withCore in @($true, $false)) {
        $lines = @(
            ". '$($lifecyclePath.Replace("'", "''"))'",
            $bannerText
        )
        if ($withCore) { $lines += "Import-Module '$($corePath.Replace("'", "''"))' -Prefix Core -ErrorAction Stop -WarningAction SilentlyContinue" }
        $lines += @(
            '$guard = Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn',
            'try {',
            '    try { throw [Management.Automation.PipelineStoppedException]::new() }',
            '    catch { if ($_.Exception -is [Management.Automation.PipelineStoppedException]) { throw } }',
            '    finally {',
            '        $guard.InCleanup = $true',
            $(if ($withCore) { "        Write-CoreSmartM365CompletionBanner -Status 'Auto' -ScriptName 'Probe' -StartedAt (Get-Date)" } else { "        `$guard.PendingBanner = @{ Status = 'Auto'; ScriptName = 'Probe'; StartedAt = (Get-Date) }" }),
            '    }',
            '    $guard.ReachedTerminal = $true; $guard.ReturnedNormally = $true',
            '} catch { if ($_.Exception -isnot [Management.Automation.PipelineStoppedException]) { $guard.ErrorRecord = $_ } }',
            'finally { Complete-SmartM365InventoryRun -State $guard; if ($guard.ExitCode -eq 130) { exit 130 } }'
        )
        $result = Invoke-ChildScript -Name $(if ($withCore) { 'finalizer-both' } else { 'finalizer-tenantcontext' }) -Lines $lines
        $label = if ($withCore) { 'both banners loaded' } else { 'TenantContext banner only' }
        Assert-Banner ($result.ExitCode -eq 130 -and $result.Output -match 'Status\s+:\s+CANCELLED' -and $result.Output -notmatch 'Cleanup failed' -and $result.Output -cnotmatch 'SUCCESS') "Deferred banner with $label is CANCELLED, exit 130 (exit $($result.ExitCode)): $($result.Output.Trim())"
    }

    # 3. Entry scripts that call the TenantContext banner from their own finally.
    $entryScripts = @(
        'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1',
        'SmartInventory/ExchangeInventory/Hybrid/FreeBusy/SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1',
        'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1',
        'SmartInventory/ExchangeInventory/BackupProtection/SmartM365-M365-BackupPolicyScope-Inventory.ps1',
        'SmartInventory/ExchangeInventory/Migration/SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1',
        'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1'
    )
    $knownCommands = @('Write-SmartM365CompletionBanner', 'Start-SmartM365InventoryRun', 'Complete-SmartM365InventoryRun', 'Get-SmartM365InventoryCompletionStatus', 'Test-SmartM365InterruptedRun', 'Disable-SmartM365InterruptedPublication', 'Invoke-SmartM365RunCleanup', 'Close-SmartM365InventoryConnections', 'Stop-SmartM365RunTranscript', 'Start-SmartM365RunTranscript', 'Set-SmartM365InventoryConnectionOwnership')
    foreach ($relative in $entryScripts) {
        $entryAst = Get-ScriptAst (Join-Path $SourceRoot $relative)
        $outerTry = @($entryAst.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State') })[0]
        $innerTry = $outerTry.Body.Find({ param($n) $n -is [Management.Automation.Language.TryStatementAst] -and $n.Finally -and $n.Finally.Extent.Text.Contains('Write-SmartM365CompletionBanner') -and $n.CatchClauses.Count -gt 0 }, $true)
        Assert-Banner ($null -ne $innerTry) "$relative calls the TenantContext banner from its finally."
        $commands = @(@($innerTry.CatchClauses[0], $innerTry.Finally) | ForEach-Object { $_.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] }, $true) } | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ -and $_ -notin $knownCommands } | Sort-Object -Unique)
        $stubs = @($commands | ForEach-Object { "if (-not (Get-Command -Name '$($_.Replace("'", "''"))' -CommandType Cmdlet, Application, Alias -ErrorAction SilentlyContinue)) { function global:$_ { } }" })
        foreach ($scenario in @('Interrupted', 'Completed')) {
            $body = if ($scenario -eq 'Interrupted') { 'throw [Management.Automation.PipelineStoppedException]::new()' } else { "`$script:CompletionStatus = 'Success'; `$smartM365RunGuard.ReachedTerminal = `$true" }
            $lines = @(
                ". '$($lifecyclePath.Replace("'", "''"))'",
                $bannerText
            ) + $stubs + @(
                '$global:EnableSharePointUpload = $false; $global:EnableTeamsNotifications = $false',
                "`$ScriptName = 'Probe'; `$TaskName = 'Probe'; `$StartTime = Get-Date; `$RunStartedAt = `$StartTime; `$scriptStart = `$StartTime; `$script:StartedAt = `$StartTime; `$script:CompletionStatus = 'Auto'",
                '$smartM365RunGuard = Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn',
                'try {',
                "try { $body }",
                $innerTry.CatchClauses[0].Extent.Text,
                "finally $($innerTry.Finally.Extent.Text)",
                '$smartM365RunGuard.ReachedTerminal=$true',
                '$smartM365RunGuard.ReturnedNormally=$true',
                "} $($outerTry.CatchClauses[0].Extent.Text) finally $($outerTry.Finally.Extent.Text)",
                'if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}'
            )
            $result = Invoke-ChildScript -Name ("{0}-{1}" -f [IO.Path]::GetFileNameWithoutExtension($relative), $scenario) -Lines $lines
            $name = Split-Path -Path $relative -Leaf
            if ($scenario -eq 'Interrupted') {
                Assert-Banner ($result.ExitCode -eq 130 -and $result.Output -match 'Status\s+:\s+CANCELLED' -and $result.Output -cnotmatch 'Status\s+:\s+SUCCESS') "$name interrupted: CANCELLED and exit 130 (exit $($result.ExitCode)): $($result.Output.Trim())"
            }
            else {
                Assert-Banner ($result.ExitCode -eq 0 -and $result.Output -match 'Status\s+:\s+SUCCESS' -and $result.Output -notmatch 'CANCELLED') "$name completed: SUCCESS and exit 0 as before (exit $($result.ExitCode)): $($result.Output.Trim())"
            }
        }
    }
}
finally {
    Remove-Variable -Name SmartM365InventoryRun, SmartM365ErrorCount, SmartM365CompletionBannerRunKey -Scope Global -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ("PASS: {0} completion banner checks; PowerShell {1}." -f $script:Checks, $PSVersionTable.PSVersion)
