<#
.SYNOPSIS
    Offline checks for empty Microsoft 365 usage reports.
.DESCRIPTION
    Runs the collector's actual retry and fallback functions on synthetic CSVs
    with a mocked download. No Graph, tenant collection, or publication is
    performed.
.VERSION
1.0
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$collector = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-M365UserActivity-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
foreach ($name in @('Get-M365UsageReportDefinition','Assert-M365UsageReportRefreshDate','Invoke-M365UsageReportAcquisition','Resolve-M365UsageEmptyReportFallback')) {
    $definition = @($ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    },$true))
    if ($definition.Count -ne 1) { throw "Expected one $name function." }
    Invoke-Expression $definition[0].Extent.Text
}

$script:warnings = [System.Collections.Generic.List[string]]::new()
$script:sleeps = [System.Collections.Generic.List[int]]::new()
$script:responses = [System.Collections.Generic.Queue[string]]::new()
$script:downloads = 0
function WriteLog { param([string]$Message,[string]$Level) if ($Level -eq 'WARNING') { $script:warnings.Add($Message) | Out-Null } }
function Start-Sleep { param([int]$Seconds) $script:sleeps.Add($Seconds) | Out-Null }
function Invoke-M365UsageReportDownload {
    param($ReportDefinition,[string]$RawPath)
    $script:downloads++
    Set-Content -LiteralPath $RawPath -Value $script:responses.Dequeue() -Encoding UTF8
}
function Assert-Case { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw $Message } }
function Assert-Rejected { param([scriptblock]$Action,[string]$Pattern,[string]$Message)
    $caught = $false
    try { & $Action | Out-Null } catch { $caught = $_.Exception.Message -match $Pattern }
    Assert-Case $caught $Message
}
function Reset-Case { param([string[]]$Responses)
    $script:warnings.Clear(); $script:sleeps.Clear(); $script:responses.Clear(); $script:downloads = 0
    foreach ($response in $Responses) { $script:responses.Enqueue($response) }
}

$definitions = Get-M365UsageReportDefinition
$activation = $definitions['Office365ActivationUserDetail']
$activeUsers = $definitions['Office365ActiveUserDetail']
$refreshText = (Get-Date).Date.AddDays(-2).ToString('yyyy-MM-dd')
$header = 'Report Refresh Date,User Principal Name,Product Type'
$data = "$header`r`n$refreshText,user1@contoso.example,MICROSOFT 365 APPS FOR ENTERPRISE"
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-UsageEmpty-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    $rawPath = Join-Path $tempRoot 'raw.csv'
    $latest = Join-Path $tempRoot 'DATA-LAST'
    New-Item -ItemType Directory -Path $latest -Force | Out-Null

    Assert-Case (@($definitions.Values | Where-Object { $_.PSObject.Properties['Required'] -and $_.Required }).Name -join ',' -eq 'Office365ActiveUserDetail') 'Only the active user detail report may be required.'

    Set-Content -LiteralPath $rawPath -Value $header -Encoding UTF8
    Assert-Rejected { Assert-M365UsageReportRefreshDate -RawPath $rawPath -ReportName 'X' -WarningAgeDays 3 } 'Graph report is empty: X' 'Strict refresh-date check accepted an empty report.'

    # Empty, then populated: the retry recovers within the same run.
    Reset-Case @($header,$data)
    $date = Invoke-M365UsageReportAcquisition -ReportDefinition $activation -RawPath $rawPath -WarningAgeDays 3 -Attempts 3 -RetryDelaySeconds 120
    Assert-Case ($null -ne $date -and $date.ToString('yyyy-MM-dd') -eq $refreshText) 'Populated retry did not return its refresh date.'
    Assert-Case ($script:downloads -eq 2 -and ($script:sleeps -join ',') -eq '120') "Unexpected retry sequence: downloads=$script:downloads sleeps=$($script:sleeps -join ',')."
    Assert-Case ($script:warnings.Count -eq 1 -and $script:warnings[0] -match 'attempt 1/3') 'Retry warning missing.'

    # A populated first answer is neither retried nor warned.
    Reset-Case @($data)
    $null = Invoke-M365UsageReportAcquisition -ReportDefinition $activation -RawPath $rawPath -WarningAgeDays 3 -Attempts 3 -RetryDelaySeconds 120
    Assert-Case ($script:downloads -eq 1 -and $script:sleeps.Count -eq 0 -and $script:warnings.Count -eq 0) 'Populated report was retried.'

    # Persistently empty: every attempt is used, no sleep after the last one.
    Reset-Case @($header,$header,$header)
    $date = Invoke-M365UsageReportAcquisition -ReportDefinition $activation -RawPath $rawPath -WarningAgeDays 3 -Attempts 3 -RetryDelaySeconds 120
    Assert-Case ($null -eq $date -and $script:downloads -eq 3 -and $script:sleeps.Count -eq 2) 'Persistently empty report was not retried exactly three times.'

    # Non-required report: previous CSV kept byte for byte, its refresh date reported.
    $previous = Join-Path $latest 'M365_Apps_Activations.csv'
    'RunId,ReportName,ReportPeriodRequested,Report Refresh Date,User Principal Name','old,Office365ActivationUserDetail,D180,2026-10-18,user1@contoso.example' | Set-Content -LiteralPath $previous -Encoding UTF8
    $before = (Get-FileHash -LiteralPath $previous -Algorithm SHA256).Hash
    $script:warnings.Clear()
    $entry = Resolve-M365UsageEmptyReportFallback -ReportDefinition $activation -LatestFolderPath $latest -Attempts 3
    Assert-Case ($entry -eq 'Office365ActivationUserDetail=2026-10-18') "Unexpected carried-over entry: $entry"
    Assert-Case ((Get-FileHash -LiteralPath $previous -Algorithm SHA256).Hash -eq $before) 'Previous CSV was modified.'
    Assert-Case ($script:warnings.Count -eq 1 -and $script:warnings[0] -match 'kept unchanged \(Report Refresh Date=2026-10-18\).*not part of this run''s receipt') 'Carried-over warning missing.'

    Remove-Item -LiteralPath $previous -Force
    $script:warnings.Clear()
    $entry = Resolve-M365UsageEmptyReportFallback -ReportDefinition $activation -LatestFolderPath $latest -Attempts 3
    Assert-Case ($entry -eq 'Office365ActivationUserDetail=none' -and $script:warnings[0] -match 'No previous M365_Apps_Activations.csv exists') 'Missing previous CSV was not reported.'

    # Required report: an empty answer stays fatal.
    Assert-Rejected { Resolve-M365UsageEmptyReportFallback -ReportDefinition $activeUsers -LatestFolderPath $latest -Attempts 3 } 'Graph report is empty: Office365ActiveUserDetail after 3 attempt\(s\)' 'Required empty report did not fail the run.'

    # The run loop uses the retry and the fallback, and records the fallback in the receipt.
    $runText = @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] -and $_.Finally })[0].Body.Extent.Text
    Assert-Case ($runText -match 'Invoke-M365UsageReportAcquisition' -and $runText -match 'Resolve-M365UsageEmptyReportFallback' -and $runText -notmatch 'Assert-M365UsageReportRefreshDate') 'Run loop does not use the retry and fallback.'
    Assert-Case ($runText -match 'not receipted' -and $runText -match 'EmptyKeptPrevious') 'Carried-over reports are missing from the receipt qualifications or summary.'
    'PASS: empty reports retry then recover; persistent empties keep the previous CSV with a warning; the required report stays fatal.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    $base = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($base,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-UsageEmpty-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
    }
}
