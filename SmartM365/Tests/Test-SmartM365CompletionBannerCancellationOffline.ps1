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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZQivEfjvwRzxD
# NKm44WYrV43wYvMYwv39U+Tt0G9vyKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIN+C/OgYhxB5oa6UmxCVfF4YFZ5ssI9/NStuMo3A8sLCMA0GCSqG
# SIb3DQEBAQUABIIBgD0tLrwEUYDJ5WtriobZVrql/GvLb7MAxI5XwyVUQn201EBx
# fRdvFKupS1k3tuzXFQ/DhZV7GghhcgRktFp69IEaH2YI+7JOWLiCaYcFCI0ZAwU2
# W6GRjhQmfgsewzdPxmRQBjozKICsVrYmL76mSZNJJvYoRJSEDH5mIkMSlD7d5rUm
# brfg45byzT67DofbDF/RD7UjH96rwGwHH0/+uYEpiBJy5X1bVo4+hLRp264V7XUP
# ibTSGJyEPK5CGq8dRY2AY2WF3W7Bg/KJW2PhMKRXX6DONBFCCUEURPPytFtPBI+F
# mPtyrt82lnRq3VpE6/Muula6z87GrKF1/MkvWT/nyVjdUHh01CYIZDLMOXB/3Ani
# MOWCl0BjdHgZSUenEQhT/ZKx+ktbTHvxpbo7PF0zvv8yoRcfgoFXZVGlt5ya+0mG
# PRppHO96JDX7VlLV2v7D0wWFIcll70HNvbTCN3UUbPM62XxnO6KQTM6ELGxKgJz5
# hj3P9rqcm5PUd51heKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEyMDMy
# NDVaMC8GCSqGSIb3DQEJBDEiBCAaklnEfRN0Ea+Gysru1Z1YwMR5YJiJIpPqWimy
# YH2trDANBgkqhkiG9w0BAQEFAASCAgBy7iEHNeNf1S1c51ZLtFZ+vgHV0yBBq/NF
# scAz/pD53715SWruXsmq6k4zSQJU4oQUlev2h7KUsGQiXWdKg8Aliv3JucQ1rv0c
# DvJ31yYH8vVKRaSYQojCkP6tBSPcTODUPd9n6+a8L4lLlNTlgic4IWqBe/2uifbV
# 9grtRnqmlYKKOs7Xt5oUu7mNYUSHoVZISam1YOVITiFDw9n1iXAgo6w0xiKYIaqn
# RmwkM1uDipZt+lt5+hju47ZSGzOVniuK/yp/QzTHunr6/q1M7AR9t7tDc83HDnmO
# mFoFr0RoECSt027BsPVdbNub0EF0I4w67eMJvF2YQbsuKi9wiaK5ig/MIVSeHONx
# 7CW+/m0epos8m4JRSmOgseTj40MTLc0AuVnt/o9ot9+4vyUKubQt+ftnxF4G+8M7
# mHzbkK2PambSlNa9iHxm5+KxnO2UQpJ9iNvkHE8mKlu5LQZjBciiEmGPtqeouKj8
# yV0xdM3oG1TC53oZZkNx1lkO/o5iPM6F3ziIfpTQ9hTxn1JS1oeJ3jPY7g/ZAhvQ
# 0Qf1Mb7P0KybJoxbnXmCBj5INx58IwaLhyYpLVtzdcwxdtsw4elBwpmpkSBD1blf
# tMT8dpNJmZdL0bO/2qv6E255f7G+guyd79tB5wThZRvL3JbYB2lx+R/ahNZg0b5u
# G+xhDGO/JA==
# SIG # End signature block
