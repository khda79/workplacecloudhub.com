<#
.SYNOPSIS
    Analyze local ShareGate migration reports without connecting to ShareGate.
.VERSION
    1.0.6
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [string]$InputPath = '',
    [string]$OutputDirectory = '',
    [string]$SessionId = '',
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
    for ($index = 0; $index -lt $inputs.Count; $index++) {
        $arguments.Add('--input'); $arguments.Add($inputs[$index])
        $arguments.Add('--source-label'); $arguments.Add($sourceLabels[$index])
        $arguments.Add('--source-snapshot')
        $arguments.Add($(if ($snapshots.ContainsKey($sourceLabels[$index])) { $snapshots[$sourceLabels[$index]] } else { $sourceLabels[$index] }))
    }
    if ($SessionId) { $arguments.Add('--session'); $arguments.Add($SessionId) }
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBXrnVi5z5Qxkgm
# t5j5vvDGjjk4ecAMhCsga/U7hiinwqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAOfwUAcH6CKSCe07cP//9S
# 5TC6XDLhsRofvLB2uDokjjANBgkqhkiG9w0BAQEFAASCAYAjYHyinsd7HHXYq+T/
# fhBYTwPJNWxjhY0sXr7CDKx237y5vhFG7wD3neK/dkJ4R9Vp6QJo7raKvlolExEj
# B92dZ8R4QTHWPvIJejFwa+MsAZzFB+6e0tRh2Pi8i0Ai2H4ICB8q3+jAPE56AdVs
# 3EdrTrg6hM9h7t+HaWKq1ry0Mo5Ymx3gkr4GtLEEq9S/W98Dx84SdKWntoppi5W6
# Gu9GgQ58ZZChKveWTEqwR/D3+R1/h9bGQ9GGVzaEoycwbDTpIFsi9kzk3eQNhzBz
# jjdc8es39lnBZrgJKHAd5pNov0SHdkF6qdpOAIbCXG9zSPP0UTWQBew1yz5D4ZPt
# fdkGEZN22DKFpWTL2VyTyEE45OmjWpIsgw9PaqAT7rjfZ8FPY1PuInRi9UByN+3Z
# +1LmoXQaI+1Hn5CPrR5lNCLAhLLAMB+hxUFVUzVO++pFcIZKvlIozYGPBSaiWmtM
# 2HyxfOdOcgBBYvdfVOHTPx52NgdstHxnabIuHTo4NBtbLDE=
# SIG # End signature block
