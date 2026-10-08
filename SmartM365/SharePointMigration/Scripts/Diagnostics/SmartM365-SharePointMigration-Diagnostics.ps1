<#
.SYNOPSIS
    Analyze local ShareGate migration reports without connecting to ShareGate.
.VERSION
    1.0.7
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [string]$InputPath = '',
    [string]$OutputDirectory = '',
    [string]$SessionId = '',
    [string]$ReportTimeZoneId = '',
    [string]$ActivityPath = '',
    [string]$PatternKey = '',
    [ValidateSet('To fix', 'Accepted', 'Fixed')][string]$SetPatternState = '',
    [string]$ExpectedState = '',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {
$toolRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).ProviderPath
$analyzer = Join-Path $PSScriptRoot 'analyze_sharegate_reports.py'
$activityHelper = Join-Path $toolRoot 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-GuiActivity.ps1'
if ($ActivityPath) { . $activityHelper }
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-ImportExcel.ps1')

function Write-DiagnosticPhase {
    param([string]$State, [string]$Message)
    foreach ($line in ($Message -split '\r?\n')) {
        Microsoft.PowerShell.Utility\Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $line)
    }
    if ($OutputDirectory -and -not $DryRun) {
        [void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
        @{ State=$State; Message=$Message; UpdatedUtc=[DateTimeOffset]::UtcNow.ToString('o') } |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'analysis.phase.json.txt') -Encoding utf8
    }
}

function Copy-StableShareGateWorkbook {
    param([string]$Source, [string]$Destination, [int]$Attempts = 3)
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        $sourceStream = $targetStream = $archive = $null
        try {
            $before = Get-Item -LiteralPath $Source -ErrorAction Stop
            Start-Sleep -Milliseconds 500
            $ready = Get-Item -LiteralPath $Source -ErrorAction Stop
            if ($before.Length -ne $ready.Length -or $before.LastWriteTimeUtc -ne $ready.LastWriteTimeUtc) { throw 'Workbook is still changing.' }
            # Deny concurrent writes/deletes while reading the original bytes.
            $sourceStream = [IO.File]::Open($Source, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            $targetStream = [IO.File]::Create($Destination)
            $sourceStream.CopyTo($targetStream)
            $targetStream.Dispose(); $targetStream = $null
            $sourceStream.Dispose(); $sourceStream = $null
            $after = Get-Item -LiteralPath $Source -ErrorAction Stop
            if ($ready.Length -ne $after.Length -or $ready.LastWriteTimeUtc -ne $after.LastWriteTimeUtc -or
                (Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash) {
                throw 'Workbook changed while its snapshot was being copied.'
            }
            $archive = [IO.Compression.ZipFile]::OpenRead($Destination)
            if (-not $archive.GetEntry('[Content_Types].xml') -or -not $archive.GetEntry('xl/workbook.xml')) { throw 'Workbook package is incomplete.' }
            foreach ($entry in $archive.Entries) {
                $stream = $entry.Open()
                try { $stream.CopyTo([IO.Stream]::Null) } finally { $stream.Dispose() }
            }
            return $Destination
        }
        catch {
            if ($attempt -eq $Attempts) { throw "Cannot obtain a stable workbook snapshot after $Attempts attempts: $Source. $($_.Exception.Message)" }
            Write-DiagnosticPhase 'WaitingForReport' ("Waiting for a stable ShareGate workbook ({0}/{1}): {2}" -f $attempt, $Attempts, $_.Exception.Message)
        }
        finally {
            if ($archive) { $archive.Dispose() }
            if ($targetStream) { $targetStream.Dispose() }
            if ($sourceStream) { $sourceStream.Dispose() }
        }
        Start-Sleep -Seconds $attempt
    }
}

function Write-DiagnosticEvent {
    param([string]$Status, [int]$ExitCode, [string]$Detail)
    if ($ActivityPath) {
        $logPath = if ($OutputDirectory) { Join-Path $OutputDirectory 'analysis.stdout.log' } else { '' }
        Write-SmartM365GuiActivityEvent -Path $ActivityPath -Status $Status -ExitCode $ExitCode -Detail $Detail -LogPath $logPath
    }
}

$converted = [System.Collections.Generic.List[string]]::new()
$snapshots = @{}
$snapshotRoot = ''
try {
    $project = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
    if (-not (Test-Path -LiteralPath $project -PathType Container)) { throw 'Project root is not a directory.' }
    $reportTimeZoneSource = if ($ReportTimeZoneId) { 'explicit parameter' } else { 'current analysis machine' }
    if (-not $ReportTimeZoneId) { $ReportTimeZoneId = [TimeZoneInfo]::Local.Id }
    [void][TimeZoneInfo]::FindSystemTimeZoneById($ReportTimeZoneId)
    $python = Join-Path $toolRoot 'Tools\Python\python.exe'
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        $python = (Get-Command python -ErrorAction Stop).Source
    }
    if ($SetPatternState) {
        if (-not $PatternKey -or $PatternKey -notmatch '^[0-9a-f]{64}$') { throw 'A valid pattern key is required.' }
        if ($DryRun) {
            Write-Output "DryRun: set pattern $PatternKey to '$SetPatternState' in $project"
            return
        }
        & $python $analyzer --project-root $project --pattern-key $PatternKey --set-state $SetPatternState --expected-state $ExpectedState
        if ($LASTEXITCODE -ne 0) { throw "Pattern state update failed with exit code $LASTEXITCODE." }
        Write-DiagnosticEvent -Status 'Succeeded' -ExitCode 0 -Detail "Pattern state set to $SetPatternState; key=$PatternKey"
        return
    }
    if (-not $InputPath) { $InputPath = Join-Path $project 'ShareGate\MigrationReport' }
    $inputItem = Get-Item -LiteralPath $InputPath -ErrorAction Stop
    if ($inputItem.PSIsContainer) {
        $files = @(Get-ChildItem -LiteralPath $inputItem.FullName -File | Where-Object { $_.Extension -in @('.csv', '.xlsx') })
    }
    else {
        if ($inputItem.Extension -notin @('.csv', '.xlsx')) { throw 'Select a CSV or XLSX report.' }
        $files = @($inputItem)
    }
    $csvFiles = @($files | Where-Object Extension -EQ '.csv' | Sort-Object FullName)
    $allXlsxFiles = @($files | Where-Object Extension -EQ '.xlsx' | Sort-Object FullName)
    $xlsxFiles = @($allXlsxFiles | Where-Object { $name = $_.BaseName; -not @($csvFiles | Where-Object BaseName -EQ $name).Count })
    if ($files.Count -eq 0) { throw 'No CSV or XLSX reports were found.' }
    $excelModule = if ($allXlsxFiles.Count) {
        Initialize-SmartM365ImportExcel -DryRun:$DryRun -Progress { param($state,$message) Write-DiagnosticPhase $state $message }
    } else { $null }
    if ($excelModule -and $allXlsxFiles.Count) {
        $snapshotRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ShareGate-' + [guid]::NewGuid().ToString('N'))
        [void](New-Item -ItemType Directory -Path $snapshotRoot)
        foreach ($file in $allXlsxFiles) {
            $destination = Join-Path $snapshotRoot ([guid]::NewGuid().ToString('N') + '.xlsx')
            $snapshots[$file.FullName] = Copy-StableShareGateWorkbook -Source $file.FullName -Destination $destination
        }
    }
    if (-not $DryRun) { Write-DiagnosticPhase 'InspectingReports' 'Inspecting the selected ShareGate report…' }
    $inventory = [System.Collections.Generic.List[object]]::new()
    $sessions = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $csvFiles) {
        $inspection = & $python $analyzer --inspect $file.FullName | ConvertFrom-Json
        if ($LASTEXITCODE -ne 0) { throw "Cannot inspect CSV: $($file.FullName)" }
        $record = @($inspection)[0]
        $status = if ($record.Error) { 'Error' } else { 'Used' }
        $reason = if ($record.Error) { $record.Error } else { 'CSV report selected.' }
        $inventory.Add([pscustomobject]@{ File=$file.FullName; Status=$status; Reason=$reason; Sessions=@($record.Sessions) })
        foreach ($session in @($record.Sessions)) { [void]$sessions.Add([string]$session) }
    }
    foreach ($file in $allXlsxFiles) {
        $matchingCsv = @($csvFiles | Where-Object BaseName -EQ $file.BaseName | Select-Object -First 1)
        if ($matchingCsv.Count) {
            if (-not $excelModule) {
                $inventory.Add([pscustomobject]@{ File=$file.FullName; Status='Skipped - ImportExcel not installed'; Reason="CSV counterpart is used: $($matchingCsv[0].Name). XLSX equivalence cannot be verified without ImportExcel."; Sessions=@() })
                continue
            }
            try {
                Import-Module ImportExcel -ErrorAction Stop
                $sheet = @(Get-ExcelSheetInfo -Path $snapshots[$file.FullName] -ErrorAction Stop | Where-Object Name -EQ 'Data' | Select-Object -First 1)
                $excelRows = if ($sheet.Count) { @(Import-Excel -Path $snapshots[$file.FullName] -WorksheetName 'Data' -ErrorAction Stop) } else { @(Import-Excel -Path $snapshots[$file.FullName] -ErrorAction Stop) }
                $csvRows = @(Import-Csv -LiteralPath $matchingCsv[0].FullName -ErrorAction Stop)
                if ($excelRows.Count -ne $csvRows.Count) { throw "Row count differs: CSV=$($csvRows.Count), XLSX=$($excelRows.Count)." }
                $checkColumns = @('Session ID','ID','Status','Type','Source site address','Destination site address')
                for ($rowIndex = 0; $rowIndex -lt $csvRows.Count; $rowIndex++) {
                    foreach ($column in $checkColumns) {
                        if ([string]$excelRows[$rowIndex].$column -ne [string]$csvRows[$rowIndex].$column) { throw "Row $($rowIndex + 1) differs in '$column'." }
                    }
                }
                $matchingSessions = @($inventory | Where-Object File -EQ $matchingCsv[0].FullName | Select-Object -First 1 | ForEach-Object Sessions)
                $inventory.Add([pscustomobject]@{ File=$file.FullName; Status='Masked - same row identities as CSV'; Reason="Same row count and matching session, row ID, status, type and endpoint columns as $($matchingCsv[0].Name); other cell values can differ between export formats."; Sessions=$matchingSessions })
            }
            catch { $inventory.Add([pscustomobject]@{ File=$file.FullName; Status='Error'; Reason=$_.Exception.Message; Sessions=@() }) }
        }
        elseif (-not $excelModule) {
            $inventory.Add([pscustomobject]@{ File=$file.FullName; Status='Requires ImportExcel'; Reason='ImportExcel will be installed automatically for analysis; DryRun does not install it.'; Sessions=@() })
        }
        elseif ($DryRun) {
            try {
                Import-Module ImportExcel -ErrorAction Stop
                $sheet = @(Get-ExcelSheetInfo -Path $snapshots[$file.FullName] -ErrorAction Stop | Where-Object Name -EQ 'Data' | Select-Object -First 1)
                $excelRows = if ($sheet.Count) { @(Import-Excel -Path $snapshots[$file.FullName] -WorksheetName 'Data' -ErrorAction Stop) } else { @(Import-Excel -Path $snapshots[$file.FullName] -ErrorAction Stop) }
                $found = @($excelRows | ForEach-Object { $_.'Session ID' } | Where-Object { $_ } | Sort-Object -Unique)
                foreach ($session in $found) { [void]$sessions.Add([string]$session) }
                $inventory.Add([pscustomobject]@{ File=$file.FullName; Status='Used'; Reason='XLSX report selected for conversion.'; Sessions=$found })
            }
            catch { $inventory.Add([pscustomobject]@{ File=$file.FullName; Status='Error'; Reason=$_.Exception.Message; Sessions=@() }) }
        }
        else { $inventory.Add([pscustomobject]@{ File=$file.FullName; Status='Used'; Reason='XLSX report selected for conversion.'; Sessions=@() }) }
    }
    $usedCsv = @($inventory | Where-Object { $_.Status -eq 'Used' -and $_.File -like '*.csv' })
    $usedXlsx = @($inventory | Where-Object { $_.Status -eq 'Used' -and $_.File -like '*.xlsx' })
    $errors = @($inventory | Where-Object Status -EQ 'Error')
    if ($errors.Count -and -not $DryRun) { throw ('Input validation failed: ' + (($errors | ForEach-Object { $_.File + ': ' + $_.Reason }) -join '; ')) }
    if (-not $usedCsv.Count -and -not $usedXlsx.Count -and -not ($DryRun -and $allXlsxFiles.Count)) { throw 'No usable CSV or XLSX report was found.' }
    if ($DryRun) {
        $moduleLabel = if ($excelModule) { "present v$($excelModule.Version)" } else { 'not installed' }
        Write-Output "DryRun: project=$project; input=$($inputItem.FullName); found CSV=$($csvFiles.Count), XLSX=$($allXlsxFiles.Count); used CSV=$($usedCsv.Count), XLSX=$($usedXlsx.Count); ImportExcel=$moduleLabel"
        Write-Output "ShareGate local report dates: $ReportTimeZoneId ($reportTimeZoneSource)."
        foreach ($entry in $inventory) { Write-Output ("{0}: {1} | {2} | Sessions: {3}" -f $entry.Status, $entry.File, $entry.Reason, $(if ($entry.Sessions.Count) { $entry.Sessions -join ', ' } else { '(none)' })) }
        Write-Output ("Detected Session IDs: {0}" -f $(if ($sessions.Count) { @($sessions | Sort-Object) -join ', ' } else { '(none)' }))
        Write-Output ("Selected Session ID: {0}" -f $(if ($SessionId) { $SessionId } else { 'all detected sessions' }))
        return
    }
    if (-not $OutputDirectory) {
        $id = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
        $OutputDirectory = Join-Path $project "ShareGate\Diagnostics\$id"
    }
    $OutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $inputs = [System.Collections.Generic.List[string]]::new()
    $sourceLabels = [System.Collections.Generic.List[string]]::new()
    foreach ($file in $csvFiles) {
        if (@($inventory | Where-Object { $_.File -eq $file.FullName -and $_.Status -eq 'Used' }).Count) { $inputs.Add($file.FullName); $sourceLabels.Add($file.FullName) }
    }
    if ($xlsxFiles.Count -gt 0) {
        if (-not $excelModule) {
            throw 'ImportExcel preparation did not complete; XLSX reports cannot be analyzed.'
        }
        else {
            Import-Module ImportExcel -ErrorAction Stop
            foreach ($file in $xlsxFiles) {
                $target = Join-Path $OutputDirectory ('.converted-' + [guid]::NewGuid().ToString('N') + '.csv')
                $sheet = @(Get-ExcelSheetInfo -Path $snapshots[$file.FullName] | Where-Object Name -EQ 'Data' | Select-Object -First 1)
                if ($sheet.Count -gt 0) { $rows = @(Import-Excel -Path $snapshots[$file.FullName] -WorksheetName 'Data') }
                else { $rows = @(Import-Excel -Path $snapshots[$file.FullName]) }
                if ($rows.Count -eq 0) { throw "XLSX report has no rows: $($file.FullName)" }
                $rows | Export-Csv -LiteralPath $target -NoTypeInformation -Encoding utf8
                $converted.Add($target)
                $inputs.Add($target)
                $sourceLabels.Add($file.FullName)
            }
        }
    }
    $arguments = [System.Collections.Generic.List[string]]::new()
    $arguments.Add($analyzer)
    $arguments.Add('--project-root'); $arguments.Add($project)
    $arguments.Add('--output-dir'); $arguments.Add($OutputDirectory)
    $arguments.Add('--report-time-zone'); $arguments.Add($ReportTimeZoneId)
    for ($index = 0; $index -lt $inputs.Count; $index++) {
        $arguments.Add('--input'); $arguments.Add($inputs[$index])
        $arguments.Add('--source-label'); $arguments.Add($sourceLabels[$index])
        $arguments.Add('--source-snapshot')
        $arguments.Add($(if ($snapshots.ContainsKey($sourceLabels[$index])) { $snapshots[$sourceLabels[$index]] } else { $sourceLabels[$index] }))
    }
    if ($SessionId) { $arguments.Add('--session'); $arguments.Add($SessionId) }
    Write-DiagnosticPhase 'AnalyzingReports' "Interpreting ShareGate report dates without UTC offsets in $ReportTimeZoneId ($reportTimeZoneSource)."
    Write-DiagnosticPhase 'AnalyzingReports' 'Analyzing the ShareGate report and generating the HTML summary…'
    Write-DiagnosticEvent -Status 'Running' -ExitCode 0 -Detail "Analyzing $($inputs.Count) local report(s); output=$OutputDirectory"
    $analyzerOutput = @(& $python $arguments.ToArray())
    $code = $LASTEXITCODE
    if ($code -ne 0) { throw "Analysis failed with exit code $code. See stderr for details." }
    foreach ($line in $analyzerOutput | Select-Object -SkipLast 1) { Write-Output $line }
    $summary = Get-Content -LiteralPath (Join-Path $OutputDirectory 'Summary.json.txt') -Raw | ConvertFrom-Json -AsHashtable
    $unknown = @($summary.Patterns | Where-Object RuleId -EQ 'UNKNOWN').Count
    Write-Output ("Summary: lines={0}; distinct items={1}; Success={2}; Warning={3}; Error={4}; To fix={5}; Accepted={6}; residual lines={7}%; residual items={8}%; Unknown patterns={9}" -f $summary.Lines, $summary.DistinctItems, [int]$summary.LineStatus['Success'], [int]$summary.LineStatus['Warning'], [int]$summary.LineStatus['Error'], [int]$summary.IssueLineState['To fix'], [int]$summary.IssueLineState['Accepted'], $summary.ResidualLineRate, $summary.ResidualItemRate, $unknown)
    Write-DiagnosticPhase 'Completed' 'ShareGate report analysis completed.'
    Write-Output $summary.ReportPath
    Write-DiagnosticEvent -Status 'Succeeded' -ExitCode 0 -Detail "ShareGate report analysis completed; output=$OutputDirectory"
}
catch {
    Write-DiagnosticPhase 'Failed' $_.Exception.Message
    Write-DiagnosticEvent -Status 'Failed' -ExitCode 1 -Detail $_.Exception.Message
    [Console]::Error.WriteLine($_.Exception.Message)
    $script:ConsoleLifecycleStatus = 'FAILED'
    exit 1
}
finally {
    foreach ($file in $converted) { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue }
    if ($snapshotRoot -and (Test-Path -LiteralPath $snapshotRoot)) {
        Get-ChildItem -LiteralPath $snapshotRoot -File | Remove-Item -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $snapshotRoot -ErrorAction SilentlyContinue
    }
}
}
catch {
    $script:ConsoleLifecycleFailure = $_
    throw
}
finally {
    Complete-SmartM365MigrationConsoleLifecycle -Context $script:ConsoleLifecycleContext -Failure $script:ConsoleLifecycleFailure -Status $script:ConsoleLifecycleStatus
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDoGdt5u667lUZ+
# j+bIwpaLbSH69dM+j6Pd7trNrnBgv6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAbYpl0NoDZldhM/SwqQlQ2ew+SJhYlAoRiOXZfJ2fUwMA0GCSqG
# SIb3DQEBAQUABIIBgKig21svzKvaWc3Eq6vaacQJBvA+RNsppKLjbdMd/0AIhSYT
# EiOGvVvHRd5AqbACnUE/KivIUj2Dv/VaO/yZHL9m5EM5htxb+tGKP3MnCW1g/ALB
# RBUCW4Z8LDbNsEp3i3NO5HhBDxWd+hyot3WKOpwD7JyuhPgb1/8dtqUq0WwdhdhS
# r0FP9LmfOtWcrdY4PspUSMTrKnKnCVAwpSPButr6V0oxlwL8mVF2GIZLixSO8qk+
# e6kVmtvHW85QuwsOHnuRWr2szdolYe+dw5nYJrRBlGTj8zsBLTWvFqhNuW6Lks74
# 4+fPqGvEois7h+DcBAIOOujkV2pFW8jCBNKvQHGL2DDf5O0YPhtXYRS2dVl2+g4K
# cqnDoXrnEwRVqcTrwRHrwT1chSd/YvPEdS5OBJNMO3hadNSOaxGXEQz4YjSA+SgT
# seU2+hIbeeXcopJu7EFg9DXoqtPcENX2qpgaAK91YxtVn6HocMVTcPFJAAawhXtX
# 5dL7Rc80LygkwLT4i6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgxNjEy
# MjRaMC8GCSqGSIb3DQEJBDEiBCAjdv0LRn2kN3xyfw+7lt/OWkWqaDDtxIt4f2Hh
# DbT3pzANBgkqhkiG9w0BAQEFAASCAgA+bn5bF7v5tT5Hy31GhtuLRt1menQWx5iT
# wxJvQJeKG5IUqmqiM3u2gEnqoS4NuleVJh3XqwbrOtP3xjBh5Rk560HCyuhDm4WG
# fv7UrKZVZxvx5b1mpE7eWzfyXWYVwupg58nt732RD867/IzyKJwvPgjW+IcEocBT
# w0/nDrK/HvSxNFC+Q6+obtH2CUyJq7LDgIaSB5gcbET6jJpOkn3Q2UG9Hjqi5th6
# nDYCVKjAc3EVkeDuOFdZasTjUKccuGOONbFq7gRPFinBDw6SNs2kNyDwwHmp2xoG
# BQ9ypHoIPtirQOZEVIeFku1/L0EKaalubIyJi3/rfgCUGALO9CiC4W0njFt3Zd23
# vSwgn9dCMOWLviY/lXicXkz52xUWw0k4Gnqa4LTu2v/Fjhd+pF14xDHZtukVQUHn
# t5Bh9jHs/bmpR/qa8Mhex+DkP7Iptx0tCExW+8utmGQ8i3cMO4TXZS0zVWIrRb8O
# n2rQPJCf2JNd1jTz3uzGUwnW5N7hdqvDn2lfedm5Jdm1vSOPgS03PBfxWRf/CT6P
# oq7e2RbuNbpVyy7BKC/uQE7Er1c7OwO9sWbglIm4yHEJx4AlBsmfxkYKG3s5+Zj1
# JLrUM3fsm6m4nvL83azYMWHCOZKkSGHhD+FEccyzmIv6am8vHvAbAcpP8S2WAuwn
# sDXZarMA8Q==
# SIG # End signature block
