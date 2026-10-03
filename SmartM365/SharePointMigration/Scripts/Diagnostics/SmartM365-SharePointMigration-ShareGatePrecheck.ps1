<#
.SYNOPSIS
    Item-level ShareGate pre-check. Copy-Content is invoked only with -WhatIf.
.VERSION
    1.0.2
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [string]$AnalysisDirectory = '',
    [Parameter(Mandatory)][string]$SessionId,
    [switch]$DryRun,
    [switch]$Run,
    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Microsoft.PowerShell.Utility\Write-Host ('{0} Script  : {1} v{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $MyInvocation.MyCommand.Name, '1.0.2') -ForegroundColor Cyan
if (-not $WhatIf) { throw 'Refusing to start: -WhatIf is mandatory for every ShareGate pre-check invocation.' }
if ($DryRun -and $Run) { throw 'Choose either -DryRun or -Run.' }
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) { throw 'Use Windows PowerShell 5.1 (powershell.exe) for ShareGate.' }
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-ShareGateReportReader.ps1')
$reportAliases = Get-SmartM365ShareGateReportAliases -ConfigRoot (Join-Path $PSScriptRoot '..\..\Config')
$execute = [bool]$Run
$project = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
if (-not (Test-Path -LiteralPath $project -PathType Container)) { throw 'ProjectRoot must be a directory.' }
$diagnosticsRoot = Join-Path $project 'ShareGate\Diagnostics'
if (-not $AnalysisDirectory) {
    if (-not (Test-Path -LiteralPath $diagnosticsRoot -PathType Container)) { throw 'Run phase 2a analysis for this project first.' }
    $latest = Get-ChildItem -LiteralPath $diagnosticsRoot -Directory |
        Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'ClassifiedRows.csv') -PathType Leaf } |
        Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    if (-not $latest) { throw 'No phase 2a ClassifiedRows.csv was found.' }
    $AnalysisDirectory = $latest.FullName
}
$analysis = (Resolve-Path -LiteralPath $AnalysisDirectory -ErrorAction Stop).ProviderPath
$classifiedPath = Join-Path $analysis 'ClassifiedRows.csv'
if (-not (Test-Path -LiteralPath $classifiedPath -PathType Leaf)) { throw 'ClassifiedRows.csv is missing.' }
$allRows = @(Import-Csv -LiteralPath $classifiedPath -Encoding UTF8)
if (-not @($allRows | Where-Object SessionId -EQ $SessionId).Count) { throw "Session '$SessionId' was not found in ClassifiedRows.csv." }
$accessRows = @($allRows | Where-Object { $_.SessionId -eq $SessionId -and $_.RuleId -in @('SG-ACCESS-SOURCE','SG-ACCESS-TARGET','SG-ACCESS-UNKNOWN') })
$itemsByKey = @{}
$skipped = New-Object 'System.Collections.Generic.List[object]'
foreach ($row in $accessRows) {
    $number = 0
    $hasId = [int]::TryParse([string]$row.SourceItemId, [ref]$number) -and $number -gt 0
    if (-not $row.ItemKey -or -not $hasId -or -not $row.SourceUrl -or -not $row.SourceList -or -not $row.DestinationUrl -or -not $row.DestinationList) {
        $skipped.Add([pscustomobject]@{ SessionId=$row.SessionId; RowId=$row.RowId; SourceUrl=$row.SourceUrl; SourceList=$row.SourceList; SourceItemId=$row.SourceItemId; Reason='Missing positive item ID, item key, site URL or list name.' })
        continue
    }
    if (-not $itemsByKey.ContainsKey($row.ItemKey)) {
        $itemsByKey[$row.ItemKey] = [pscustomobject]@{ ItemKey=$row.ItemKey; SourceUrl=$row.SourceUrl; SourceList=$row.SourceList; SourceItemId=$number; DestinationUrl=$row.DestinationUrl; DestinationList=$row.DestinationList; RowIds=(New-Object 'System.Collections.Generic.List[string]'); Ambiguous=$false }
    }
    $item = $itemsByKey[$row.ItemKey]
    if ($item.DestinationUrl.TrimEnd('/') -ne $row.DestinationUrl.TrimEnd('/') -or $item.DestinationList -ne $row.DestinationList) { $item.Ambiguous = $true }
    $item.RowIds.Add([string]$row.RowId)
}
$items = @($itemsByKey.Values | Sort-Object SourceUrl,SourceList,DestinationUrl,DestinationList,SourceItemId)
$groups = @($items | Group-Object SourceUrl,SourceList,DestinationUrl,DestinationList)
Write-Output ('{0} Mode={1}; WhatIf=True; Session={2}; Analysis={3}; access lines={4}; distinct items={5}; groups={6}; skipped rows={7}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $(if ($execute) { 'Run' } else { 'DryRun' }), $SessionId, $analysis, $accessRows.Count, $items.Count, $groups.Count, $skipped.Count)
foreach ($group in $groups) {
    $first = $group.Group[0]
    Write-Output ('{0} Group: {1} | {2} -> {3} | {4}; items={5}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $first.SourceUrl, $first.SourceList, $first.DestinationUrl, $first.DestinationList, $group.Count)
}
if (-not $execute) { return }

$runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$output = Join-Path $diagnosticsRoot ('Precheck-' + $runId)
$reports = Join-Path $output 'Reports'
New-Item -ItemType Directory -Path $reports -Force | Out-Null
$logPath = Join-Path $output 'Precheck.log'
$itemResults = New-Object 'System.Collections.Generic.List[object]'
$siteCache = @{}
$listCache = @{}

function Write-PrecheckLog {
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
function Get-PrecheckStatus {
    param([string]$Text, [string]$SourceUrl, [string]$DestinationUrl)
    if ($Text -notmatch '(?i)(?:\(401\)|HTTP\s*401|401\s+Unauthorized)') { return 'No 401 observed in pre-check' }
    $failedHosts = @([regex]::Matches($Text, "(?i)(?:WebUri\s*:\s*'(?<web>https?://[^']+)'|URL\s*'(?<failure>https?://[^']+)'\s*was not authorized)") |
        ForEach-Object { $url = if ($_.Groups['web'].Success) { $_.Groups['web'].Value } else { $_.Groups['failure'].Value }; ([uri]$url).Host.ToLowerInvariant() } | Select-Object -Unique)
    $sourceHost = ([uri]$SourceUrl).Host.ToLowerInvariant()
    $targetHost = ([uri]$DestinationUrl).Host.ToLowerInvariant()
    if ($failedHosts.Count -eq 1 -and $failedHosts[0] -eq $sourceHost -and $sourceHost -ne $targetHost) { return 'Source 401 observed' }
    if ($failedHosts.Count -eq 1 -and $failedHosts[0] -eq $targetHost -and $sourceHost -ne $targetHost) { return 'Destination 401 observed' }
    return '401 observed - side undetermined'
}

try {
    Write-PrecheckLog ('Actor={0}\{1}; Machine={2}; Module=ShareGate; WhatIf=True; Session={3}; Analysis={4}' -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME, $SessionId, $analysis)
    Export-AtomicCsv -Path (Join-Path $output 'Precheck-SkippedRows.csv') -Rows $skipped.ToArray() -Columns @('SessionId','RowId','SourceUrl','SourceList','SourceItemId','Reason')
    Export-AtomicCsv -Path (Join-Path $output 'Precheck-Items.csv') -Rows @() -Columns @('SessionId','ItemKey','RowIds','SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList','Status','Reason','PrecheckSessionId','ReportPath')
    $module = @(Get-Module -ListAvailable -Name ShareGate | Sort-Object Version -Descending | Select-Object -First 1)
    if (-not $module.Count) { throw 'ShareGate module is not discoverable in Windows PowerShell 5.1.' }
    Import-Module -Name $module[0].Path -ErrorAction Stop
    Write-PrecheckLog ('ShareGate module version={0}; path={1}' -f $module[0].Version, $module[0].Path)
    foreach ($name in @('Connect-Site','Get-List','Copy-Content','Export-Report')) {
        if (-not (Get-Command -Name $name -Module ShareGate -ErrorAction SilentlyContinue)) { throw "Required ShareGate cmdlet is missing: $name" }
    }
    $copy = Get-Command Copy-Content -Module ShareGate
    $sourceItemSet = @($copy.ParameterSets | Where-Object { @($_.Parameters | Where-Object Name -EQ 'SourceItemId').Count -gt 0 -and @($_.Parameters | Where-Object Name -EQ 'WhatIf').Count -gt 0 })
    if (-not $sourceItemSet.Count) { throw 'Installed Copy-Content does not expose SourceItemId and WhatIf together.' }
    foreach ($group in $groups) {
        $first = $group.Group[0]
        $sourceList = $destinationList = $null
        $groupError = ''
        try {
            $sourceList = Get-ExactList -Side 'Source' -Url $first.SourceUrl -Name $first.SourceList
            $destinationList = Get-ExactList -Side 'Destination' -Url $first.DestinationUrl -Name $first.DestinationList
        }
        catch { $groupError = $_.Exception.Message; Write-PrecheckLog ('Group connection failed: ' + $groupError) }
        foreach ($item in $group.Group) {
            $status = 'Undetermined'
            $reason = ''
            $reportPath = ''
            $precheckSession = ''
            if ($item.Ambiguous) { $status = 'Skipped - ambiguous destination'; $reason = 'The same source item maps to multiple destination lists or sites.' }
            elseif ($groupError) { $status = 'Undetermined - connection failed'; $reason = $groupError }
            else {
                $reportPath = Join-Path $reports ('Precheck-' + [guid]::NewGuid().ToString('N') + '.csv')
                try {
                    # Hard safety boundary: every Copy-Content call passes the literal -WhatIf switch.
                    $copyResult = ShareGate\Copy-Content -SourceList $sourceList -DestinationList $destinationList -SourceItemId @([int]$item.SourceItemId) -WhatIf -ErrorAction Stop
                    if (-not $copyResult) { $status = 'Undetermined - no CopyResult'; $reason = 'ShareGate returned no result to export.'; $reportPath = '' }
                    else {
                        if (@($copyResult).Count -ne 1) { throw 'Copy-Content returned more than one CopyResult for one item.' }
                        if ($copyResult.PSObject.Properties['Id']) { $precheckSession = [string]$copyResult.Id }
                        ShareGate\Export-Report -CopyResult $copyResult -Path $reportPath -ErrorAction Stop | Out-Null
                        if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) { throw 'Export-Report did not create a CSV.' }
                        $reportRows = @(Import-Csv -LiteralPath $reportPath -Encoding UTF8)
                        $assessment = Get-SmartM365ShareGatePrecheckAssessment -Rows $reportRows -SourceItemId $item.SourceItemId -SourceUrl $item.SourceUrl -DestinationUrl $item.DestinationUrl -Aliases $reportAliases
                        $status = $assessment.Status
                        $reason = if ($reportRows.Count -eq 0) { 'Export-Report contained only a header; item-level access remains undetermined.' }
                                  else { 'Observed in the ShareGate -WhatIf pre-check report; no content was copied. Absence of 401 does not prove a full migration will succeed.' }
                    }
                }
                catch {
                    $reason = $_.Exception.Message
                    $status = Get-PrecheckStatus -Text $reason -SourceUrl $item.SourceUrl -DestinationUrl $item.DestinationUrl
                    if ($status -eq 'No 401 observed in pre-check') { $status = 'Undetermined - pre-check error' }
                    if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) { $reportPath = '' }
                }
            }
            $itemResults.Add([pscustomobject]@{ SessionId=$SessionId; ItemKey=$item.ItemKey; RowIds=($item.RowIds -join ';'); SourceUrl=$item.SourceUrl; SourceList=$item.SourceList; SourceItemId=$item.SourceItemId; DestinationUrl=$item.DestinationUrl; DestinationList=$item.DestinationList; Status=$status; Reason=$reason; PrecheckSessionId=$precheckSession; ReportPath=$reportPath })
            Export-AtomicCsv -Path (Join-Path $output 'Precheck-Items.csv') -Rows $itemResults.ToArray() -Columns @('SessionId','ItemKey','RowIds','SourceUrl','SourceList','SourceItemId','DestinationUrl','DestinationList','Status','Reason','PrecheckSessionId','ReportPath')
            Write-PrecheckLog ('Item={0}; list={1}; ID={2}; status={3}; report={4}' -f $item.ItemKey, $item.SourceList, $item.SourceItemId, $status, $reportPath)
        }
    }
    Write-PrecheckLog ('Completed: items={0}; statuses={1}; output={2}' -f $itemResults.Count, (($itemResults | Group-Object Status | ForEach-Object { $_.Name + '=' + $_.Count }) -join ', '), $output)
}
catch {
    Write-PrecheckLog ('Failed: ' + $_.Exception.Message)
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCFIisyh8WfrMD3
# sLtn2ktwwoNSk5D4IV7Q5LnO53B/caCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMf5UUdk5xGDEMXmcKnW+TaLoYTjPvxon84rFN1cbV2fMA0GCSqG
# SIb3DQEBAQUABIIBgKuTAyFBOc7O5f2oyPq0eFbDqfWxoNIYNRr1Rb0bIkn4TWMX
# i15Y1oemZSNR19LtpTc1w0ISHM/OnEEL3oV63qqq6KC1veqBBESorIc5wW6o1ZMC
# 58VZMSRRvmdMWtC7EZzlgSplC4Xo+a0fr8shM4NuuzSvdVdbfTukNp720DvOeRpC
# 3JkiA51gWEuSk1g4dHBvRC+CWxWEjFblE6JHYCdeRGtrmpD3n0kWgZaFvN1Uz2SH
# ibDdd5tPnD9j15i7UjPXpSbRwzraKrlk2Id3gteFlDITaITwv3nBUZVTF8y8LmR8
# vs68hTtvDXI6/F03dqJr2tLLlnqNmWVlgs9Gk7P5B4k7+BLJwf3gMreqXcarVv5R
# qCXDZR43KtFvt1pK14UNTv2gSRG2uc9sVYB4u6lCehu617xBjw6KFS/uNdvgDnaV
# dbSfmC4zc4lv2G6MUqgiGoHpkrumg/b1hfNB9szYK77o0rkTPdT6nBiiLr3Gn/pO
# keCbCmcvalD2I3JoDKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxMDQy
# NDVaMC8GCSqGSIb3DQEJBDEiBCC5nri8XeDq92nyRR6s2Bx7e9js7kirCV6KyN6L
# az+3xDANBgkqhkiG9w0BAQEFAASCAgB3/MYKZDlPXwZLD8eqeYTy9d13A2rX2rbr
# y95ZHAwee8YT23AueFqc967CaoTjXxBMdk4q8gEpwbTeno+ItYGpKazOY2h93wRZ
# LOxro9yxYGAjInOEJZCk3iKlikYW9V3xyOTRm6/SJTb4tyVSeMTpYyM25zNgMCG4
# lQcOLFvTuq1Omc8yEYSOES+vTah0cCWSI7fvefNsHZyWC37lTVfBlhlNYv93fdMd
# z0CeOMDO+MOmlvmWu8bb9+Oq88c1mAR1Jtb2bxMovh7kcpMr8a3E5RkiIxlPoP9P
# m6Qq2XvJGc7JiNwqb56Ahn5kL2UotLH2f9T9xPPnDvaNm3JuiBQk0DkODgyaPHNy
# ibcfyegv7bgVRmmrcp7oAc/HkXqR/dIi/o/yygUDQrobhAIOyHGReM2/EhW9JoC2
# xIpVsVq3/FfCpMva0quLoFJow5a8aeIAf9WTT1ccRMqLG4SCa/ovS5Jhsn7ioOCn
# 2F4Udc6L5srBDOrMKdisH6kpK6RtMJmhZ0SJHScMHIjqFlv7j6XsG2gEcURo/X56
# 3u9tkjPqeMgPgyzY78KxAIl8Gdoz4bKQlWO3QhVBEgutnnvjkNkAL+B9u/dyZpq8
# 7xfYgXwxAc3rSfoyTKiJ1IyxOcbSq9iwyLjsT59HzPfcGzfqEGUDBBE90YAEif8O
# p2jx7LmZ3g==
# SIG # End signature block
