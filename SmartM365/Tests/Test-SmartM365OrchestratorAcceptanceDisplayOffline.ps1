#Requires -Version 7.0
<#
.SYNOPSIS
Checks the present / functional capability columns of Servers and versions, offline.
.DESCRIPTION
Present capabilities come from the orchestrator probes (Orchestrator-Capabilities.json, used by
the election). Functional capabilities come from the last server acceptance test run as the
orchestrator service account (Orchestrator-Acceptance.json); a run under another account
(Orchestrator-Acceptance.OtherAccount.json) is reported but never validates. Display only.
Synthetic files in a temporary folder; no tenant, collector or publication.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
$script:Checks = 0
function Assert-Display {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}
function Write-TestJson([string]$Path, $Value) { [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false)) }
function New-Run([string]$Account, [bool]$ServiceAccount, [object[]]$Results) {
    [ordered]@{ SchemaVersion = 1; ServerName = 'X'; GeneratedAtUtc = '2026-10-30T10:15:00.0000000Z'; AcceptanceVersion = '1.0.12'; RunAccount = $Account; ServiceAccount = $ServiceAccount; Results = $Results }
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-AcceptanceDisplay-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
try {
    foreach ($server in @('SRV-A', 'SRV-B', 'SRV-C', 'SRV-D', 'SRV-E')) {
        $folder = Join-Path $root $server
        $null = New-Item -ItemType Directory -Path $folder -Force
        Write-TestJson (Join-Path $folder 'Orchestrator-Heartbeat.json') ([ordered]@{ Timestamp = (Get-Date).ToString('o'); ScriptVersion = '1.5.80' })
        Write-TestJson (Join-Path $folder 'Orchestrator-Capabilities.json') ([ordered]@{ ReadyCapabilities = @('Graph', 'Python', 'SharedRuntime', 'TeamsPowerShell') })
    }
    $results = @(
        [ordered]@{ Name = 'Graph'; Status = 'PASS'; Detail = '100 users' }
        [ordered]@{ Name = 'SharedRuntime'; Status = 'PASS'; Detail = 'ok' }
        [ordered]@{ Name = 'Python'; Status = 'PASS'; Detail = 'Python 3.12.1' }
        [ordered]@{ Name = 'TeamsPowerShell'; Status = 'WARN'; Detail = 'module copy not readable' }
        [ordered]@{ Name = 'EXO'; Status = 'FAIL'; Detail = 'Forbidden' }
        [ordered]@{ Name = 'ExchangeOnPrem'; Status = 'SKIPPED'; Detail = 'snap-in not registered' }
        [ordered]@{ Name = 'Network'; Status = 'FAIL'; Detail = 'not a capability' }
    )
    # SRV-A: service account run; a later admin run must not hide it.
    Write-TestJson (Join-Path $root 'SRV-A/Orchestrator-Acceptance.json') (New-Run 'CONTOSO\svc-orch' $true $results)
    Write-TestJson (Join-Path $root 'SRV-A/Orchestrator-Acceptance.OtherAccount.json') (New-Run 'CONTOSO\admin' $false @([ordered]@{ Name = 'Graph'; Status = 'FAIL'; Detail = 'x' }))
    # SRV-B: only an admin run.
    Write-TestJson (Join-Path $root 'SRV-B/Orchestrator-Acceptance.OtherAccount.json') (New-Run 'CONTOSO\admin' $false $results)
    # SRV-C: never tested. SRV-D: unreadable result. SRV-E: older document without optional fields.
    [IO.File]::WriteAllText((Join-Path $root 'SRV-D/Orchestrator-Acceptance.json'), '{ not json', [Text.UTF8Encoding]::new($false))
    Write-TestJson (Join-Path $root 'SRV-E/Orchestrator-Acceptance.json') ([ordered]@{ Results = @([ordered]@{ Name = 'AD'; Status = 'PASS' }) })

    $cluster = [pscustomobject]@{ ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-C', 'SRV-D', 'SRV-E'); ElectionWeightsByServer = [pscustomobject]@{}; ServerJobPolicies = [pscustomobject]@{}; PeerHeartbeatStaleMinutes = 5 }
    $rows = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $root -ClusterDocument $cluster)
    $row = { param($Name) @($rows | Where-Object Server -eq $Name)[0] }
    $localAt = ([datetimeoffset]::Parse('2026-10-30T10:15:00Z')).LocalDateTime.ToString('yyyy-MM-dd HH:mm')

    $a = & $row 'SRV-A'
    Assert-Display ($a.Capabilities -eq 'Graph, Python, SharedRuntime, TeamsPowerShell') "Present capabilities still come from the probes: $($a.Capabilities)"
    Assert-Display ($a.FunctionalCapabilities -eq 'Graph, Python, SharedRuntime') "Functional = PASS of the service account run: $($a.FunctionalCapabilities)"
    Assert-Display ($a.NotFunctionalCapabilities -eq 'EXO (FAIL), TeamsPowerShell (WARN)') "Not functional = WARN and FAIL capabilities only: $($a.NotFunctionalCapabilities)"
    Assert-Display ($a.Acceptance -eq "$localAt, service account CONTOSO\svc-orch, acceptance 1.0.12") "Acceptance time and account: $($a.Acceptance)"

    $b = & $row 'SRV-B'
    Assert-Display (-not $b.FunctionalCapabilities -and -not $b.NotFunctionalCapabilities -and $b.Acceptance -eq "Not validating: $localAt, run as CONTOSO\admin, not the service account, acceptance 1.0.12") "An admin run never validates: $($b.Acceptance)"
    $c = & $row 'SRV-C'
    Assert-Display (-not $c.FunctionalCapabilities -and $c.Acceptance -eq 'Never run as the service account') "Never tested: $($c.Acceptance)"
    $d = & $row 'SRV-D'
    Assert-Display (-not $d.FunctionalCapabilities -and $d.Acceptance -like 'Acceptance result unreadable: *' -and $d.Capabilities -like 'Graph*') "Unreadable result does not break the row: $($d.Acceptance)"
    $e = & $row 'SRV-E'
    Assert-Display ($e.FunctionalCapabilities -eq 'AD' -and $e.Acceptance -eq 'service account') "Older document without optional fields: $($e.Acceptance)"

    # GUI columns and the acceptance summary are display only: the election never reads them.
    $guiText = [IO.File]::ReadAllText((Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'))
    foreach ($column in @('Header="Present (probe)" Binding="{Binding Capabilities}"', 'Header="Functional (acceptance)" Binding="{Binding FunctionalCapabilities}"', 'Header="Not functional" Binding="{Binding NotFunctionalCapabilities}"', 'Header="Acceptance" Binding="{Binding Acceptance}"')) {
        Assert-Display ($guiText.Contains($column)) "GUI column: $column"
    }
    Assert-Display (-not $guiText.Contains('nothing kept')) 'GUI no longer says that the acceptance test keeps nothing.'
    foreach ($file in @('SmartM365.Orchestrator.Distributed.psm1', 'SmartM365-Inventory-Orchestrator.ps1')) {
        Assert-Display (-not ([IO.File]::ReadAllText((Join-Path $base $file))).Contains('Orchestrator-Acceptance')) "$file does not read the acceptance result (election unchanged)."
    }
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

"[{0}] PASS: present / functional capability display, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks
