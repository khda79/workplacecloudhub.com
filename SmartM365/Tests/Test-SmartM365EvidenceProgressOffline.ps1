#Requires -Version 7.0
<#
.SYNOPSIS
Offline tests of the shared evidence progress line (New-/Write-SmartM365EvidenceProgress). No tenant calls.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $root 'SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1') -MinimumVersion '1.0.18' -Force
$script:checks = 0
function Check([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}
# Write-SmartM365EvidenceLog calls the Core WriteLog; captured here instead of writing a log file.
$global:SyntheticProgressLog = [Collections.Generic.List[object]]::new()
function global:WriteLog { param($Message, $Level) $global:SyntheticProgressLog.Add([pscustomobject]@{Message=$Message;Level=$Level}) }
$global:SyntheticNow = [datetime]::new(2026, 10, 29, 18, 0, 0, [DateTimeKind]::Utc)
$clock = { $global:SyntheticNow }
function New-Progress([int]$Total, [int]$EveryItems=1000, [double]$EveryMinutes=5) {
    $global:SyntheticProgressLog.Clear()
    $global:SyntheticNow = [datetime]::new(2026, 10, 29, 18, 0, 0, [DateTimeKind]::Utc)
    New-SmartM365EvidenceProgress -Activity 'Direct membership' -Total $Total -Unit 'groups' -EveryItems $EveryItems -EveryMinutes $EveryMinutes -Clock $clock
}
try {
    # Count trigger: one line per 1 000 items, none in between, none at 0.
    $progress = New-Progress -Total 3500 -EveryMinutes 0
    $written = @(0..3500 | Where-Object { $global:SyntheticNow = $global:SyntheticNow.AddMilliseconds(100); Write-SmartM365EvidenceProgress -Progress $progress -Processed $_ })
    Check ($written.Count -eq 3 -and ($written -join ',') -ceq '1000,2000,3000') "Count trigger wrote at: $($written -join ',')"
    Check ($global:SyntheticProgressLog.Count -eq 3 -and @($global:SyntheticProgressLog | Where-Object Level -ne 'INFO').Count -eq 0) 'Progress lines are not exactly three INFO lines.'

    # Line content: counter, percentage, elapsed, rate and ETA, culture independent.
    $progress = New-Progress -Total 4000
    $global:SyntheticNow = $global:SyntheticNow.AddMinutes(10)
    Check (Write-SmartM365EvidenceProgress -Progress $progress -Processed 1000 -Detail 'membership rows=52000') 'Due line was not written.'
    $line = $global:SyntheticProgressLog[0].Message
    Check ($line -match '^Direct membership progress: 1000/4000 groups \(25\.0%\); elapsed 10\.0 min; 1\.7 groups/s; ETA 30 min \(~\d\d:\d\d\); membership rows=52000$') "Unexpected progress line: $line"

    # Time trigger: a slow loop still writes every 5 minutes, with the count trigger reset by it.
    $progress = New-Progress -Total 100
    $global:SyntheticNow = $global:SyntheticNow.AddMinutes(4.9)
    Check (-not (Write-SmartM365EvidenceProgress -Progress $progress -Processed 3)) 'Line written before the time interval.'
    $global:SyntheticNow = $global:SyntheticNow.AddMinutes(0.1)
    Check (Write-SmartM365EvidenceProgress -Progress $progress -Processed 4) 'Time trigger did not write after 5 minutes.'
    $global:SyntheticNow = $global:SyntheticNow.AddMinutes(1)
    Check (-not (Write-SmartM365EvidenceProgress -Progress $progress -Processed 5)) 'Time trigger was not reset by the last line.'

    # Both triggers at once give a single line; no advance gives no line, even when overdue (a real stall stays visible).
    $progress = New-Progress -Total 5000
    $global:SyntheticNow = $global:SyntheticNow.AddMinutes(6)
    Check (Write-SmartM365EvidenceProgress -Progress $progress -Processed 1000) 'Combined trigger did not write.'
    Check (-not (Write-SmartM365EvidenceProgress -Progress $progress -Processed 1000)) 'Same count was written twice.'
    $global:SyntheticNow = $global:SyntheticNow.AddMinutes(30)
    Check (-not (Write-SmartM365EvidenceProgress -Progress $progress -Processed 1000)) 'A loop that did not advance hid a stall.'
    Check ($global:SyntheticProgressLog.Count -eq 1) 'Combined or repeated triggers wrote more than one line.'

    # Edge cases: start, empty total, more items than announced, forced line, disabled triggers.
    $progress = New-Progress -Total 0
    Check (-not (Write-SmartM365EvidenceProgress -Progress $progress -Processed 0)) 'A line was written at 0 processed.'
    Check (Write-SmartM365EvidenceProgress -Progress $progress -Processed 0 -Force) 'Forced line was not written.'
    Check ($global:SyntheticProgressLog[0].Message -match '0/0 groups \(n/a\); elapsed 0\.0 min; 0\.0 groups/s; ETA n/a\.$') "Empty total line: $($global:SyntheticProgressLog[0].Message)"
    $progress = New-Progress -Total 10 -EveryItems 5
    $global:SyntheticNow = $global:SyntheticNow.AddMinutes(1)
    Check (Write-SmartM365EvidenceProgress -Progress $progress -Processed 15) 'Overrun count was not written.'
    Check ($global:SyntheticProgressLog[0].Message -match '15/10 groups \(150\.0%\).*ETA n/a\.$') "Overrun line: $($global:SyntheticProgressLog[0].Message)"
    $progress = New-Progress -Total 10 -EveryItems 0 -EveryMinutes 0
    $global:SyntheticNow = $global:SyntheticNow.AddMinutes(60)
    Check (-not (Write-SmartM365EvidenceProgress -Progress $progress -Processed 9)) 'Disabled triggers still wrote a line.'

    # Log volume for the 29/10 case: 37 787 groups over about 3 h stays below a hundred lines.
    $progress = New-Progress -Total 37787
    $lines = 0
    foreach ($index in 1..37787) {
        $global:SyntheticNow = $global:SyntheticNow.AddSeconds(0.3)
        if (Write-SmartM365EvidenceProgress -Progress $progress -Processed $index) { $lines++ }
    }
    Check ($lines -ge 37 -and $lines -le 100) "Unexpected line count for a 37 787 group run: $lines"
    $bytes = ($global:SyntheticProgressLog | ForEach-Object { $_.Message.Length + 40 } | Measure-Object -Sum).Sum
    Check ($bytes -lt 20KB) "Progress lines exceed 20 KB: $bytes bytes."
}
finally {
    Remove-Item Function:\global:WriteLog -ErrorAction SilentlyContinue
    Remove-Variable SyntheticProgressLog, SyntheticNow -Scope Global -ErrorAction SilentlyContinue
}
Write-Output "PASS: $script:checks offline evidence progress checks. No collectors, APIs, mail or live writes."
