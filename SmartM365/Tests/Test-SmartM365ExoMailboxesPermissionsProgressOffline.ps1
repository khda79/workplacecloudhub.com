<#
.SYNOPSIS
Synthetic, offline tests for the progress lines of the EXO mailboxes parallel permission collection (-PermissionsOnly).
.VERSION
1.0.0
.NOTES
Synthetic counters and a local ForEach-Object -Parallel sample only. No tenant, Exchange, API, mail or synchronized data access.
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$results = New-Object 'System.Collections.Generic.List[object]'

function Assert-Offline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Test-OfflineCase {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        [void]$results.Add([pscustomobject]@{ Name=$Name; Passed=$true; Error='' })
    }
    catch {
        [void]$results.Add([pscustomobject]@{ Name=$Name; Passed=$false; Error=$_.Exception.Message })
    }
}

function Get-FunctionText {
    param([string]$Path, [string[]]$Names)
    $tokens=$null; $parseErrors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    if($parseErrors.Count){throw "Source parse failed: $Path"}
    foreach($name in $Names){
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if($null -eq $node){throw "Function not found: $name"}
        $node.Extent.Text
    }
}

$producerPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1'
$functionText = @(Get-FunctionText -Path $producerPath -Names @('New-MailboxPermissionProgressTracker', 'Get-MailboxPermissionProgressLine')) -join "`n"
$module = New-Module -ScriptBlock ([scriptblock]::Create($functionText))

try {
    Test-OfflineCase 'A line is logged at each 1,000-mailbox step and on the last mailbox only' {
        $lines = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 2500
            for ($i = 1; $i -le 2500; $i++) {
                $tracker.Processed = $i
                Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds ($i * 0.05) -Now ([datetime]'2026-10-30 06:00:00')
            }
        }
        $lines = @($lines | Where-Object { $_ })
        Assert-Offline ($lines.Count -eq 3) ("Expected 3 count-based lines, got {0}." -f $lines.Count)
        Assert-Offline ($lines[0] -like 'Mailbox permissions progress: 1000/2500 mailboxes (40.0%);*') "Unexpected first line: $($lines[0])"
        Assert-Offline ($lines[1] -like '*: 2000/2500 mailboxes (80.0%);*') "Unexpected second line: $($lines[1])"
        Assert-Offline ($lines[2] -like '*: 2500/2500 mailboxes (100.0%);*ETA done.') "Unexpected last line: $($lines[2])"
    }

    Test-OfflineCase 'A slow collection logs at least every 5 minutes while mailboxes complete' {
        $lines = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 50
            for ($i = 1; $i -le 20; $i++) {
                $tracker.Processed = $i
                Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds ($i * 70.0) -Now ([datetime]'2026-10-30 06:00:00')
            }
        }
        $lines = @($lines | Where-Object { $_ })
        # 70 s per mailbox: lines after 5, 10, 15 and 20 mailboxes (350, 700, 1050, 1400 s).
        Assert-Offline ($lines.Count -eq 4) ("Expected 4 time-based lines, got {0}." -f $lines.Count)
        Assert-Offline ($lines[0] -like '*: 5/50 mailboxes (10.0%);*') "Unexpected first time-based line: $($lines[0])"
    }

    Test-OfflineCase 'No progress means no line, so the stalled-job alert still fires' {
        $observed = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 100
            $tracker.Processed = 10
            $first = Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 400
            $stalled = Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 7200
            [pscustomobject]@{ First=$first; Stalled=$stalled }
        }
        Assert-Offline (-not [string]::IsNullOrEmpty($observed.First)) 'The first time-based line was not produced.'
        Assert-Offline ($null -eq $observed.Stalled) 'A line was produced without new completed mailboxes.'
    }

    Test-OfflineCase 'Rate, ETA, counters and worker detail are exact and culture-invariant ASCII' {
        $previousCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR')
            $line = & $module {
                $tracker = New-MailboxPermissionProgressTracker -Total 4000 -WorkerTotals @(2000, 2000)
                $tracker.Processed = 1000
                $tracker.Errors = 3
                $tracker.ConnectionFailures = 1
                $tracker.WorkerProcessed[0] = 600
                $tracker.WorkerProcessed[1] = 400
                Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 500 -Now ([datetime]'2026-10-30 06:00:00')
            }
        }
        finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
        }
        $expected = 'Mailbox permissions progress: 1000/4000 mailboxes (25.0%); errors=3; connectionFailures=1; rate=2.00 mailbox/s; elapsed=8.3 min; ETA ~25 min (06:25); workers 600/2000 400/2000.'
        Assert-Offline ($line -ceq $expected) "Unexpected line: $line"
        Assert-Offline (-not ($line.ToCharArray() | Where-Object { [int]$_ -gt 127 })) 'Progress line contains non-ASCII characters.'
    }

    Test-OfflineCase 'Workers are omitted when not given; a single mailbox completes' {
        $line = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 1
            $tracker.Processed = 1
            Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 0
        }
        Assert-Offline ($line -ceq 'Mailbox permissions progress: 1/1 mailboxes (100.0%); errors=0; connectionFailures=0; rate=0.00 mailbox/s; elapsed=0.0 min; ETA done.') "Unexpected line: $line"
    }

    Test-OfflineCase 'An empty collection produces no line' {
        $line = & $module {
            $tracker = New-MailboxPermissionProgressTracker -Total 0
            Get-MailboxPermissionProgressLine -Tracker $tracker -ElapsedSeconds 900
        }
        Assert-Offline ($null -eq $line) "A line was produced for an empty collection: $line"
    }

    Test-OfflineCase 'The parent consumes worker results as a stream and keeps the publication guards' {
        $source = Get-Content -LiteralPath $producerPath -Raw
        Assert-Offline ($source -notmatch '\$parallelPermissionOutput\s*=\s*@\(\s*\$permissionWorkItems') 'Parallel output is still buffered before any progress line.'
        Assert-Offline ($source -match '(?s)\$permissionWorkItems \| ForEach-Object -ThrottleLimit \$workerCount -Parallel \{.*?\} \| ForEach-Object \{.*?\[void\]\$parallelPermissionOutput\.Add\(\$permissionOutput\).*?Get-MailboxPermissionProgressLine -Tracker \$permissionProgress') 'Parallel results are not piped into the progress consumer.'
        Assert-Offline ($source -match '(?s)Get-MailboxPermissionProgressLine -Tracker \$permissionProgress.*?refusing to publish blank permission data.*?row-count mismatch: dispatched=') 'Connection-failure or row-parity guards are missing after the progress consumer.'
    }

    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Test-OfflineCase 'ForEach-Object -Parallel results reach the parent before all workers finish' {
            $clock = [System.Diagnostics.Stopwatch]::StartNew()
            $received = New-Object 'System.Collections.Generic.List[double]'
            @(1, 2) | ForEach-Object -ThrottleLimit 2 -Parallel {
                if ($_ -eq 1) { 'fast' } else { Start-Sleep -Seconds 4; 'slow' }
            } | ForEach-Object { [void]$received.Add($clock.Elapsed.TotalSeconds) }
            Assert-Offline ($received.Count -eq 2) 'Parallel results were lost.'
            Assert-Offline ($received[0] -lt 3.0) ("First result arrived only after {0:F1} s: output is not streamed." -f $received[0])
        }
    }
}
finally {
    Remove-Module $module -Force
}

$results.ToArray() | Format-Table -AutoSize
if ($ResultPath) { $results.ToArray() | Export-Csv -LiteralPath $ResultPath -NoTypeInformation -Encoding UTF8 }
$failed = @($results | Where-Object { -not $_.Passed })
if ($failed.Count -gt 0) {
    throw ("{0} mailbox permission progress offline test(s) failed." -f $failed.Count)
}
