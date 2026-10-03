<#
.SYNOPSIS
    Bounded, read-only ShareGate -WhatIf witness for two known warnings and three prior 401 items.
.VERSION
    1.0.0
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [Parameter(Mandatory)][string]$AnalysisDirectory,
    [Parameter(Mandatory)][string]$SessionId,
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedAnalysisHash = '',
    [switch]$DryRun,
    [switch]$Run,
    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {
if (-not $WhatIf) { throw 'Refusing to start: -WhatIf is mandatory for every ShareGate witness invocation.' }
if ($DryRun -and $Run) { throw 'Choose either -DryRun or -Run.' }
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) { throw 'Use Windows PowerShell 5.1 (powershell.exe) for ShareGate.' }
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-ShareGateReportReader.ps1')
$aliases = Get-SmartM365ShareGateReportAliases -ConfigRoot (Join-Path $PSScriptRoot '..\..\Config')
$project = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
$analysis = (Resolve-Path -LiteralPath $AnalysisDirectory -ErrorAction Stop).ProviderPath
if (-not (Test-Path -LiteralPath $project -PathType Container)) { throw 'ProjectRoot must be a directory.' }
if (-not $analysis.StartsWith((Join-Path $project 'ShareGate\Diagnostics') + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'AnalysisDirectory must be inside this project ShareGate\Diagnostics folder.'
}
$classifiedPath = Join-Path $analysis 'ClassifiedRows.csv'
if (-not (Test-Path -LiteralPath $classifiedPath -PathType Leaf)) { throw 'ClassifiedRows.csv is missing.' }
$rows = @(Import-Csv -LiteralPath $classifiedPath -Encoding UTF8 | Where-Object SessionId -EQ $SessionId)
if (-not $rows.Count) { throw "Session '$SessionId' was not found in ClassifiedRows.csv." }

function Get-EligibleItems {
    param([object[]]$Rows, [string]$RuleId)
    $byKey = @{}
    foreach ($row in $Rows) {
        if ($row.RuleId -ne $RuleId -or -not $row.ItemKey -or -not $row.SourceUrl -or -not $row.SourceList -or -not $row.DestinationUrl -or -not $row.DestinationList) { continue }
        $number = 0
        if (-not [int]::TryParse([string]$row.SourceItemId, [ref]$number) -or $number -le 0) { continue }
        if (-not $byKey.ContainsKey($row.ItemKey) -or ($byKey[$row.ItemKey].ObjectType -ne 'File' -and $row.ObjectType -eq 'File')) { $byKey[$row.ItemKey] = $row }
    }
    return @($byKey.Values)
}
function Get-GroupKey {
    param($Row)
    return (($Row.SourceUrl.TrimEnd('/') + '|' + $Row.SourceList + '|' + $Row.DestinationUrl.TrimEnd('/') + '|' + $Row.DestinationList).ToLowerInvariant())
}
function Get-FirstCandidate {
    param([object[]]$Candidates, [int[]]$ExcludedIds)
    $eligible = @($Candidates | Where-Object { [int]$_.SourceItemId -notin $ExcludedIds } |
        Sort-Object @{Expression={ if ($_.ObjectType -eq 'File') { 0 } else { 1 } }}, @{Expression={ [int]$_.SourceItemId }}, ItemKey | Select-Object -First 1)
    if (-not $eligible.Count) { return $null }
    return $eligible[0]
}
$access = @(Get-EligibleItems -Rows $rows -RuleId 'SG-ACCESS-SOURCE')
$shortcut = @(Get-EligibleItems -Rows $rows -RuleId 'SG-SHORTCUT')
$modern = @(Get-EligibleItems -Rows $rows -RuleId 'SG-MODERN-LINK')
if ($access.Count -lt 3 -or -not $shortcut.Count -or -not $modern.Count) { throw 'The selected session does not have three eligible source 401 items and both warning witness categories.' }
$groups = @($access | Group-Object -Property { Get-GroupKey -Row $_ } | Sort-Object @{Expression='Count';Descending=$true}, Name)
$dominant = @($groups[0].Group | Sort-Object @{Expression={ [int]$_.SourceItemId }}, ItemKey)
if ($dominant.Count -lt 2) { throw 'The dominant source 401 list has fewer than two distinct items.' }
$outside = @($access | Where-Object { (Get-GroupKey -Row $_) -ne $groups[0].Name } |
    Sort-Object SourceUrl, SourceList, @{Expression={ [int]$_.SourceItemId }}, ItemKey | Select-Object -First 1)
if (-not $outside.Count) { throw 'No source 401 item exists outside the dominant list.' }
$accessSelected = @($dominant[0], $dominant[-1], $outside[0])
$usedIds = @($accessSelected | ForEach-Object { [int]$_.SourceItemId })
$shortcutSelected = Get-FirstCandidate -Candidates $shortcut -ExcludedIds $usedIds
if (-not $shortcutSelected) { throw 'No distinct numeric shortcut witness ID is available.' }
$usedIds += [int]$shortcutSelected.SourceItemId
$modernSelected = Get-FirstCandidate -Candidates @($modern | Where-Object ItemKey -NE $shortcutSelected.ItemKey) -ExcludedIds $usedIds
if (-not $modernSelected) { throw 'No distinct numeric modern component witness ID is available.' }
$selection = @(
    [pscustomobject]@{ Role='WitnessShortcut'; Row=$shortcutSelected },
    [pscustomobject]@{ Role='WitnessModernLink'; Row=$modernSelected },
    [pscustomobject]@{ Role='AccessDominantFirst'; Row=$accessSelected[0] },
    [pscustomobject]@{ Role='AccessDominantLast'; Row=$accessSelected[1] },
    [pscustomobject]@{ Role='AccessOtherList'; Row=$accessSelected[2] }
)
if (@($selection | ForEach-Object { $_.Row.ItemKey } | Sort-Object -Unique).Count -ne 5) { throw 'The five selected item keys are not unique.' }
$analysisHash = (Get-FileHash -LiteralPath $classifiedPath -Algorithm SHA256).Hash
if ($ExpectedAnalysisHash -and $analysisHash -ne $ExpectedAnalysisHash.ToUpperInvariant()) { throw 'ClassifiedRows.csv changed after witness selection. Re-run DryRun and review the five IDs.' }
Write-Output ('{0} Mode={1}; WhatIf=True; Session={2}; Analysis={3}; SHA256={4}; selected=5' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $(if ($Run) { 'Run' } else { 'DryRun' }), $SessionId, $analysis, $analysisHash)
foreach ($entry in $selection) {
    $row = $entry.Row
    Write-Output ('{0} {1}: ID={2}; {3} | {4} -> {5} | {6}; prior={7}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $entry.Role, $row.SourceItemId, $row.SourceUrl, $row.SourceList, $row.DestinationUrl, $row.DestinationList, $row.RuleId)
}
if (-not $Run) { return }

$diagnosticsRoot = Join-Path $project 'ShareGate\Diagnostics'
$runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$output = Join-Path $diagnosticsRoot ('Witness-' + $runId)
$reports = Join-Path $output 'Reports'
$objects = Join-Path $output 'CopyResults'
New-Item -ItemType Directory -Path $reports,$objects -Force | Out-Null
$logPath = Join-Path $output 'Witness.log'
$results = [System.Collections.Generic.List[object]]::new()
$siteCache = @{}
$listCache = @{}

function Write-WitnessLog {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Write-Output $line
}
function Export-AtomicCsv {
    param([string]$Path, [object[]]$Rows, [string[]]$Columns)
    $temporary = Join-Path (Split-Path -Parent $Path) ('.' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        if ($Rows.Count) { $Rows | Select-Object -Property $Columns | Export-Csv -LiteralPath $temporary -NoTypeInformation -Encoding UTF8 }
        else { Set-Content -LiteralPath $temporary -Value (($Columns | ForEach-Object { '"' + $_.Replace('"','""') + '"' }) -join ',') -Encoding UTF8 }
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}
function Get-ExactList {
    param([string]$Side, [string]$Url, [string]$Name)
    $siteKey = $Side + '|' + $Url.TrimEnd('/').ToLowerInvariant()
    if (-not $siteCache.ContainsKey($siteKey)) {
        if ($Side -eq 'Source') { $siteCache[$siteKey] = ShareGate\Connect-Site -Url $Url -ErrorAction Stop }
        else { $siteCache[$siteKey] = ShareGate\Connect-Site -Url $Url -Browser -ErrorAction Stop }
        if (-not $siteCache[$siteKey]) { throw "Connect-Site returned no site for $Side $Url." }
    }
    $listKey = $siteKey + '|' + $Name.ToLowerInvariant()
    if (-not $listCache.ContainsKey($listKey)) {
        $found = @(ShareGate\Get-List -Site $siteCache[$siteKey] -Name $Name -ErrorAction Stop)
        $exact = @($found | Where-Object {
            ($_.PSObject.Properties['Title'] -and $_.Title -eq $Name) -or
            ($_.PSObject.Properties['Name'] -and $_.Name -eq $Name)
        })
        if ($exact.Count -ne 1) { throw "Get-List returned $($exact.Count) exact matches for $Side '$Name' at $Url." }
        $listCache[$listKey] = $exact[0]
    }
    return $listCache[$listKey]
}
function Get-PrecheckSessionId {
    param($CopyResult)
    foreach ($name in @('SessionId', 'SessionID', 'CopySessionId', 'Id')) {
        $property = $CopyResult.PSObject.Properties[$name]
        if ($property -and [string]$property.Value -match '^\d{6}-\d+$') { return [string]$property.Value }
    }
    return ''
}

try {
    Write-WitnessLog ('Actor={0}\{1}; Machine={2}; Module=ShareGate; WhatIf=True; Session={3}; AnalysisSHA256={4}' -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME, $SessionId, $analysisHash)
    $columns = @('Role','SessionId','ItemKey','SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList','PriorRuleId','Status','ReportRows','PrecheckSessionId','CopyResultPath','ReportPath','Error')
    $selectionRows = @($selection | ForEach-Object { [pscustomobject]@{ Role=$_.Role; ItemKey=$_.Row.ItemKey; SourceUrl=$_.Row.SourceUrl; SourceList=$_.Row.SourceList; SourceItemId=$_.Row.SourceItemId; DestinationUrl=$_.Row.DestinationUrl; DestinationList=$_.Row.DestinationList; PriorRuleId=$_.Row.RuleId } })
    Export-AtomicCsv -Path (Join-Path $output 'Witness-Selection.csv') -Rows $selectionRows -Columns @('Role','ItemKey','SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList','PriorRuleId')
    Export-AtomicCsv -Path (Join-Path $output 'Witness-Results.csv') -Rows @() -Columns $columns
    $module = @(Get-Module -ListAvailable -Name ShareGate | Sort-Object Version -Descending | Select-Object -First 1)
    if (-not $module.Count) { throw 'ShareGate module is not discoverable in Windows PowerShell 5.1.' }
    Import-Module -Name $module[0].Path -ErrorAction Stop
    Write-WitnessLog ('ShareGate module version={0}; path={1}' -f $module[0].Version, $module[0].Path)
    foreach ($name in @('Connect-Site','Get-List','Copy-Content','Export-Report')) {
        if (-not (Get-Command -Name $name -Module ShareGate -ErrorAction SilentlyContinue)) { throw "Required ShareGate cmdlet is missing: $name" }
    }
    $copy = Get-Command Copy-Content -Module ShareGate
    if (-not @($copy.ParameterSets | Where-Object { @($_.Parameters | Where-Object Name -EQ 'SourceItemId').Count -gt 0 -and @($_.Parameters | Where-Object Name -EQ 'WhatIf').Count -gt 0 }).Count) {
        throw 'Installed Copy-Content does not expose SourceItemId and WhatIf together.'
    }
    for ($index = 0; $index -lt $selection.Count; $index++) {
        $entry = $selection[$index]
        $row = $entry.Row
        $reportPath = Join-Path $reports ('Witness-{0:D2}.csv' -f ($index + 1))
        $objectPath = Join-Path $objects ('CopyResult-{0:D2}.txt' -f ($index + 1))
        $status = 'Undetermined'
        $errorText = ''
        $precheckSession = ''
        $reportRowCount = 0
        try {
            $sourceList = Get-ExactList -Side 'Source' -Url $row.SourceUrl -Name $row.SourceList
            $destinationList = Get-ExactList -Side 'Destination' -Url $row.DestinationUrl -Name $row.DestinationList
            # Hard safety boundary: the only Copy-* call always includes the literal -WhatIf switch.
            $copyResult = ShareGate\Copy-Content -SourceList $sourceList -DestinationList $destinationList -SourceItemId @([int]$row.SourceItemId) -WhatIf -ErrorAction Stop
            if (-not $copyResult -or @($copyResult).Count -ne 1) { throw 'Copy-Content did not return exactly one CopyResult.' }
            $precheckSession = Get-PrecheckSessionId -CopyResult $copyResult
            @('Type: ' + $copyResult.GetType().FullName, 'Session ID: ' + $(if ($precheckSession) { $precheckSession } else { '(not exposed)' }), '', ($copyResult | Format-List * -Force | Out-String -Width 4096)) |
                Set-Content -LiteralPath $objectPath -Encoding UTF8
            ShareGate\Export-Report -CopyResult $copyResult -Path $reportPath -ErrorAction Stop | Out-Null
            if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) { throw 'Export-Report did not create a CSV.' }
            $reportRows = @(Import-Csv -LiteralPath $reportPath -Encoding UTF8)
            $assessment = Get-SmartM365ShareGatePrecheckAssessment -Rows $reportRows -SourceItemId ([int]$row.SourceItemId) -SourceUrl $row.SourceUrl -DestinationUrl $row.DestinationUrl -Aliases $aliases
            $status = $assessment.Status
            $reportRowCount = $assessment.RowCount
        }
        catch {
            $status = 'Undetermined - witness call failed'
            $errorText = $_.Exception.Message
            if (-not (Test-Path -LiteralPath $objectPath -PathType Leaf)) { Set-Content -LiteralPath $objectPath -Value ('No CopyResult was captured. Error: ' + $errorText) -Encoding UTF8 }
        }
        $results.Add([pscustomobject]@{ Role=$entry.Role; SessionId=$SessionId; ItemKey=$row.ItemKey; SourceUrl=$row.SourceUrl; SourceList=$row.SourceList; SourceItemId=$row.SourceItemId; DestinationUrl=$row.DestinationUrl; DestinationList=$row.DestinationList; PriorRuleId=$row.RuleId; Status=$status; ReportRows=$reportRowCount; PrecheckSessionId=$precheckSession; CopyResultPath=$objectPath; ReportPath=$(if (Test-Path -LiteralPath $reportPath -PathType Leaf) { $reportPath } else { '' }); Error=$errorText })
        Export-AtomicCsv -Path (Join-Path $output 'Witness-Results.csv') -Rows $results.ToArray() -Columns $columns
        Write-WitnessLog ('Role={0}; ID={1}; status={2}; reportRows={3}; session={4}; object={5}; report={6}; error={7}' -f $entry.Role, $row.SourceItemId, $status, $reportRowCount, $(if ($precheckSession) { $precheckSession } else { '(not exposed)' }), $objectPath, $reportPath, $errorText)
    }
    $witnessRows = @($results | Where-Object { $_.Role -like 'Witness*' -and [int]$_.ReportRows -gt 0 }).Count
    $accessEmpty = @($results | Where-Object { $_.Role -like 'Access*' -and $_.Status -eq 'Undetermined - empty report' }).Count
    $interpretation = if ($witnessRows -eq 2 -and $accessEmpty -eq 3) { 'Both positive witnesses produced rows while all three prior 401 reports were empty: consistent with no issue detected by this pre-check; not proof that item access or a real migration will succeed.' }
                      elseif ($witnessRows -eq 0) { 'Both positive witnesses were empty or failed: this pre-check export method remains inconclusive.' }
                      else { 'Mixed witness results: the empty-report interpretation remains inconclusive.' }
    Write-WitnessLog ('Completed: calls={0}; witnessReportsWithRows={1}; accessReportsEmpty={2}; output={3}' -f $results.Count, $witnessRows, $accessEmpty, $output)
    Write-WitnessLog ('Interpretation: ' + $interpretation)
}
catch {
    Write-WitnessLog ('Failed: ' + $_.Exception.Message)
    throw
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBR/+fQpb566brP
# KojtI589a6NyI5hFc7HlKp8nSAIRd6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFJYo+xb2soXCOBgJAJ0SuXqe11OdkGgo8s/NhrkWfhXMA0GCSqG
# SIb3DQEBAQUABIIBgE/+PgEoCghAPp+DY+ERnBMxgGw1Zbdm+BwWeNI40B4Y2xlg
# rKc/FaT1QR4g5j/z7zbQol/Bc6xBmzv1ZLOFhlXT48seyPqLa4D1jbBqoW1iizCQ
# /BNV/73ZkOPXIa5Mw1LjoQcV3I3iQeWWVamWDOp45o/5b6Ataj6JqFg/Rbs4edU9
# OovKXW6ZozRd/BtOJo9QMfbOJ9O9Ykcaq51aEmHQkx05M7p6HnV+eWJftBIiisVc
# CkQqMAYhnY7D2XIMmAuJoBWQL5RQ0wa7dUxvcBqzOzDAd929TBTUMAWPIiq+9ITN
# lxxSfwKjPD1P8US/ermnl3D4gqbkBS74UBCvt//oaq7qGyqPFFAj8XvfmLOQ8E0E
# 4Cq9S3poa5sLAc/+pmXmVZLDZTkrY0CmD6o6GkzADWbfZRg6TlK6qeH2O5bDQ+0e
# FNUJm7oikE8L8DmXPRsT0Ja7NE5cJX+cTFYqrdtv5m+bq2X/B19qYzj2v6YlXyDF
# r3tgDYWme3exaSxTSaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNDA3
# NTlaMC8GCSqGSIb3DQEJBDEiBCCO+SA8V9THandqm81w9OxjpNL1FBD2UXqwPF1l
# 3l3ObTANBgkqhkiG9w0BAQEFAASCAgBn/850I4uJKwJ/UOsRHH180Lsup1N9iHZF
# yKxltYXRwMFlKhRY5bkkgtwtbGvGirnzNFizfECHJUYWSF2IeLxQsZdC6Sbr7idq
# vMyN8mB1y2M0MCC4hgsKlv3VVKoOyqqUzS1fbQwnZ+KgR9uemqLhGHS/m7FocvLM
# qfRXDvqwaoHVIx5O07RwXLoa3wcbU57xi1b7lkDDT3CD6BINA/nzTW6i/d41tlLU
# 8xg3Mt9r3hA2KzD827Tbj1AKYLjsvmxvh6EINx9EDb/XrGO5XPtjXyaGk9G9akv8
# layc1NYtJCPr/Yu1Jxvzjlcj627POSNDpvG5aoTc5upezzQeRzjE8Ywj9qcSyk12
# 6c/Ruj2EKYOAmHRvxRsXgGYVJRtlAbNpK0lzWa/WIKJVuxyssdpWUffoIjzYAnzy
# SsnndMRqJN5x70zOwgn8GitpxvKmEMN1Yoo3YBtqYbCHcI8kluNS87CdFUb63DtM
# +gKtWcS4yAS/aE1TZvDv9ey7F2tFhJQAWaHqR05pBcj3AL9Lkd/TSPUzr3y/h05X
# 5JbRl6BnNJz06P4uZEqZNcph6Vw4Gb0zhHJ3syBUnO0PtJ9Y0368dHzE9AaUfldf
# Lo9Ytjz2Hvxw8xPNywNH6MLYVi8Bolm44DH32prC7DX3JusfU1p6JD71bp9kqeMi
# MFY0FOdX9w==
# SIG # End signature block
