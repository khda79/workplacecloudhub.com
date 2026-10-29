<#
.SYNOPSIS
Synthetic, offline tests for the calendar-permission collector progress lines (Exchange Online and Exchange on-premises).
.VERSION
1.0.0
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

$exoPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/CalendarPermissions/SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1'
$onPremPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/CalendarPermissions/SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
$functionNames = @('New-CalendarProgressTracker', 'Get-CalendarProgressLine')
$exoText = @(Get-FunctionText -Path $exoPath -Names $functionNames) -join "`n"
$onPremText = @(Get-FunctionText -Path $onPremPath -Names $functionNames) -join "`n"
$module = New-Module -ScriptBlock ([scriptblock]::Create($exoText))

try {
    Test-OfflineCase 'Both collectors carry the same progress helpers' {
        Assert-Offline ($exoText -ceq $onPremText) 'EXO and on-premises progress helpers diverged.'
    }

    Test-OfflineCase 'A line is logged at each 1,000-mailbox step and on the last mailbox only' {
        $lines = & $module {
            $tracker = New-CalendarProgressTracker -Total 2500
            for ($i = 1; $i -le 2500; $i++) {
                $tracker.Processed = $i
                Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds ($i * 0.05) -Now ([datetime]'2026-10-29 20:00:00')
            }
        }
        $lines = @($lines | Where-Object { $_ })
        Assert-Offline ($lines.Count -eq 3) ("Expected 3 count-based lines, got {0}." -f $lines.Count)
        Assert-Offline ($lines[0] -like 'Calendar permissions progress: 1000/2500 mailboxes (40.0%);*') "Unexpected first line: $($lines[0])"
        Assert-Offline ($lines[1] -like '*: 2000/2500 mailboxes (80.0%);*') "Unexpected second line: $($lines[1])"
        Assert-Offline ($lines[2] -like '*: 2500/2500 mailboxes (100.0%);*ETA done.') "Unexpected last line: $($lines[2])"
    }

    Test-OfflineCase 'A slow collection logs at least every 5 minutes while mailboxes complete' {
        $lines = & $module {
            $tracker = New-CalendarProgressTracker -Total 50
            for ($i = 1; $i -le 20; $i++) {
                $tracker.Processed = $i
                Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds ($i * 70.0) -Now ([datetime]'2026-10-29 20:00:00')
            }
        }
        $lines = @($lines | Where-Object { $_ })
        # 70 s per mailbox: lines after 5, 10, 15 and 20 mailboxes (350, 700, 1050, 1400 s).
        Assert-Offline ($lines.Count -eq 4) ("Expected 4 time-based lines, got {0}." -f $lines.Count)
        Assert-Offline ($lines[0] -like '*: 5/50 mailboxes (10.0%);*') "Unexpected first time-based line: $($lines[0])"
    }

    Test-OfflineCase 'No progress means no line, so the stalled-job alert still fires' {
        $observed = & $module {
            $tracker = New-CalendarProgressTracker -Total 100
            $tracker.Processed = 10
            $first = Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds 400
            $stalled = Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds 7200
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
                $tracker = New-CalendarProgressTracker -Total 4000 -WorkerTotals @(2000, 2000) -TrackWarnings
                $tracker.Processed = 1000
                $tracker.Errors = 3
                $tracker.Warnings = 2
                $tracker.Rows = 12345
                $tracker.WorkerProcessed[0] = 600
                $tracker.WorkerProcessed[1] = 400
                Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds 500 -Now ([datetime]'2026-10-29 20:00:00')
            }
        }
        finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
        }
        $expected = 'Calendar permissions progress: 1000/4000 mailboxes (25.0%); errors=3; warnings=2; rows=12345; rate=2.00 mailbox/s; elapsed=8.3 min; ETA ~25 min (20:25); workers 600/2000 400/2000.'
        Assert-Offline ($line -ceq $expected) "Unexpected line: $line"
        Assert-Offline (-not ($line.ToCharArray() | Where-Object { [int]$_ -gt 127 })) 'Progress line contains non-ASCII characters.'
    }

    Test-OfflineCase 'Warnings and workers are omitted when not tracked; single mailbox completes' {
        $line = & $module {
            $tracker = New-CalendarProgressTracker -Total 1
            $tracker.Processed = 1
            Get-CalendarProgressLine -Tracker $tracker -ElapsedSeconds 0
        }
        Assert-Offline ($line -ceq 'Calendar permissions progress: 1/1 mailboxes (100.0%); errors=0; rows=0; rate=0.00 mailbox/s; elapsed=0.0 min; ETA done.') "Unexpected line: $line"
    }

    Test-OfflineCase 'EXO parent consumes worker results as a stream and logs progress' {
        $source = Get-Content -LiteralPath $exoPath -Raw
        Assert-Offline ($source -notmatch '\$parallelCalendarOutput\s*=\s*@\(\s*\$calendarWorkItems') 'Parallel output is still buffered before any progress line.'
        Assert-Offline ($source -match '(?s)ForEach-Object -ThrottleLimit \$workerCount -Parallel \{.*?\} \| ForEach-Object \{.*?Get-CalendarProgressLine -Tracker \$progressTracker') 'Parallel results are not piped into the progress consumer.'
        $onPremSource = Get-Content -LiteralPath $onPremPath -Raw
        Assert-Offline ($onPremSource -match '(?s)foreach \(\$mbx in \$mailboxes\) \{.*Get-CalendarProgressLine -Tracker \$progressTracker.*?\}\s*Write-Progress -Id 0 -Activity \$overallActivity -Completed') 'On-premises loop does not log progress per mailbox iteration.'
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
    throw ("{0} calendar-permission progress offline test(s) failed." -f $failed.Count)
}
