<#
.SYNOPSIS
    Read-only review of an existing five-item ShareGate pilot and SPO item URLs.
.VERSION
    1.0.0
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [Parameter(Mandatory)][string]$AnalysisDirectory,
    [Parameter(Mandatory)][string]$PilotDirectory,
    [Parameter(Mandatory)][string]$SessionId,
    [ValidateNotNullOrEmpty()][string]$FarmTimeZoneId = 'W. Europe Standard Time',
    [switch]$DryRun,
    [switch]$Run
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$version = '1.0.0'
Microsoft.PowerShell.Utility\Write-Host ('{0} Script  : {1} v{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $MyInvocation.MyCommand.Name, $version) -ForegroundColor Cyan
if ($DryRun -and $Run) { throw 'Choose either -DryRun or -Run.' }
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) { throw 'Use Windows PowerShell 5.1 for ShareGate.' }
. (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-TransientEvidence.ps1')
$evidence = Get-SmartM365TransientEvidence -ProjectRoot $ProjectRoot -AnalysisDirectory $AnalysisDirectory -PilotDirectory $PilotDirectory -SessionId $SessionId
$farmZone = [TimeZoneInfo]::FindSystemTimeZoneById($FarmTimeZoneId)
$script:logPath = ''

function Write-ReviewLog {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    if ($script:logPath) { Add-Content -LiteralPath $script:logPath -Value $line -Encoding UTF8 }
    Write-Output $line
}

function Get-ExactReviewList {
    param([string]$Side, [string]$SiteUrl, [string]$ListName)
    $siteKey = $Side + '|' + $SiteUrl.TrimEnd('/').ToLowerInvariant()
    if (-not $script:siteCache.ContainsKey($siteKey)) {
        $site = if ($Side -eq 'Source') {
            ShareGate\Connect-Site -Url $SiteUrl -ErrorAction Stop
        }
        else { ShareGate\Connect-Site -Url $SiteUrl -Browser -ErrorAction Stop }
        if (-not $site) { throw "Connect-Site returned no $Side site: $SiteUrl" }
        $script:siteCache[$siteKey] = $site
    }
    $listKey = $siteKey + '|' + $ListName.ToLowerInvariant()
    if (-not $script:listCache.ContainsKey($listKey)) {
        $lists = @(ShareGate\Get-List -Site $script:siteCache[$siteKey] -Name $ListName -ErrorAction Stop)
        $exact = @($lists | Where-Object {
            ($_.PSObject.Properties['Title'] -and $_.Title -eq $ListName) -or
            ($_.PSObject.Properties['Name'] -and $_.Name -eq $ListName)
        })
        if ($exact.Count -ne 1) { throw "Get-List returned $($exact.Count) exact matches for $Side '$ListName'." }
        $script:listCache[$listKey] = $exact[0]
    }
    return $script:listCache[$listKey]
}

function Get-ReviewFile {
    param($List, [string]$RelativePath)
    if (-not $RelativePath) { return $null }
    $files = @(ShareGate\Get-File -List $List -Path $RelativePath.TrimStart('/') -ErrorAction Stop)
    if ($files.Count -ne 1) { return $null }
    return $files[0]
}

function Get-ReviewModifiedValue {
    param($Object)
    if ($null -eq $Object) { return $null }
    foreach ($name in @('Modified','LastModified','TimeLastModified','ModifiedDate','DateModified','LastWriteTime')) {
        $property = $Object.PSObject.Properties[$name]
        if ($property -and $property.Value) { return $property.Value }
    }
    foreach ($container in @('Properties','Fields','FieldValues')) {
        $property = $Object.PSObject.Properties[$container]
        if (-not $property -or -not $property.Value) { continue }
        foreach ($name in @('Modified','LastModified','TimeLastModified','ModifiedDate')) {
            if ($property.Value -is [System.Collections.IDictionary] -and $property.Value.Contains($name)) {
                return $property.Value[$name]
            }
            $nested = $property.Value.PSObject.Properties[$name]
            if ($nested -and $nested.Value) { return $nested.Value }
        }
    }
    return $null
}

function Convert-ReviewModifiedDate {
    param($Value, [ValidateSet('Source','Destination')][string]$Side)
    if ($null -eq $Value -or [string]$Value -eq '') { return [pscustomobject]@{ Raw='Unavailable'; UTC=''; Comparable=$false } }
    $raw = [string]$Value
    try {
        if ($Value -is [DateTimeOffset]) { $utc = $Value.UtcDateTime }
        else {
            $date = if ($Value -is [datetime]) { $Value } else { [datetime]::Parse($raw, [Globalization.CultureInfo]::InvariantCulture) }
            if ($date.Kind -eq [DateTimeKind]::Utc) { $utc = $date }
            elseif ($date.Kind -eq [DateTimeKind]::Local) { $utc = $date.ToUniversalTime() }
            elseif ($Side -eq 'Source') {
                if ($farmZone.IsAmbiguousTime($date) -or $farmZone.IsInvalidTime($date)) { throw 'Ambiguous farm-local time.' }
                $utc = [TimeZoneInfo]::ConvertTimeToUtc($date, $farmZone)
            }
            else { $utc = [DateTime]::SpecifyKind($date, [DateTimeKind]::Utc) }
        }
        return [pscustomobject]@{ Raw=$raw; UTC=$utc.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture); Comparable=$true }
    }
    catch { return [pscustomobject]@{ Raw=$raw; UTC=''; Comparable=$false } }
}

function Test-ReviewUrl {
    param([string]$Url, [string]$SiteUrl)
    $candidate = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$candidate) -or $candidate.Scheme -ne 'https') { return $false }
    $site = [uri]$SiteUrl
    return ($candidate.Host -eq $site.Host -and $candidate.AbsolutePath.StartsWith($site.AbsolutePath.TrimEnd('/') + '/', [StringComparison]::OrdinalIgnoreCase))
}

function Get-InferredReviewUrl {
    param($List, [string]$SiteUrl, [string]$RelativePath)
    if (-not $List -or -not $RelativePath -or -not $List.PSObject.Properties['RootFolder']) { return '' }
    $root = [string]$List.RootFolder
    if (-not $root) { return '' }
    try {
        $site = [uri]$SiteUrl
        $base = if ($root -match '^https://') { $root.TrimEnd('/') }
                elseif ($root.StartsWith('/')) { $site.GetLeftPart([UriPartial]::Authority) + $root.TrimEnd('/') }
                else { $SiteUrl.TrimEnd('/') + '/' + $root.Trim('/') }
        $encoded = (($RelativePath.TrimStart('/') -split '/' | ForEach-Object { [uri]::EscapeDataString([uri]::UnescapeDataString($_)) }) -join '/')
        $url = ([uri]($base + '/' + $encoded)).AbsoluteUri
        if (Test-ReviewUrl -Url $url -SiteUrl $SiteUrl) { return $url }
    }
    catch { return '' }
    return ''
}

Write-ReviewLog ('Mode={0}; Session={1}; pilot items={2}; analysis SHA256={3}; pilot manifest SHA256={4}; ShareGate writes=none' -f
    $(if ($Run) { 'ReadOnlyRun' } else { 'DryRun' }), $SessionId, @($evidence.PilotItems).Count, $evidence.AnalysisSHA256, $evidence.PilotManifestSHA256)
if (-not $Run) {
    foreach ($item in $evidence.PilotItems) {
        Write-ReviewLog ('ID={0}; Result={1}; CopySession={2}; Export={3}; DestinationPath={4} ({5}); SPO URL=unresolved (read-only -Run required)' -f
            $item.SourceItemId, $item.Result, $item.CopySessionId, $item.ReportPath, $item.DestinationPath, $item.DestinationPathEvidence)
    }
    return
}

$module = @(Get-Module -ListAvailable -Name ShareGate | Sort-Object Version -Descending | Select-Object -First 1)
if (-not $module.Count) { throw 'ShareGate module is not discoverable under Windows PowerShell 5.1.' }
Import-Module -Name $module[0].Path -ErrorAction Stop
foreach ($name in @('Connect-Site','Get-List','Get-File','Get-ListItem')) {
    if (-not (Get-Command -Name $name -Module ShareGate -ErrorAction SilentlyContinue)) { throw "Required read-only ShareGate cmdlet is missing: $name" }
}
$script:siteCache = @{}
$script:listCache = @{}
$output = Join-Path $evidence.DiagnosticsRoot ('PilotReview-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output -Force | Out-Null
$script:logPath = Join-Path $output 'PilotReview.log'
$itemsPath = Join-Path $output 'PilotReview-Items.csv'
$summaryPath = Join-Path $output 'PilotReview-Summary.json.txt'
Write-ReviewLog ('Actor={0}\{1}; Machine={2}; ShareGate module={3}; SourceAuth=CurrentWindowsIdentity; DestinationAuth=Browser; WritesToSharePoint=none' -f
    $env:USERDOMAIN,$env:USERNAME,$env:COMPUTERNAME,$module[0].Version)
$results = [System.Collections.Generic.List[object]]::new()
foreach ($item in $evidence.PilotItems) {
    $destinationList = $null
    $destinationFile = $null
    $destinationLookupPath = if ($item.DestinationPath) { $item.DestinationPath } else { $item.SourcePath }
    $url = ''
    $urlEvidence = 'Unavailable'
    $sourceModified = [pscustomobject]@{ Raw='Unavailable'; UTC=''; Comparable=$false }
    $destinationModified = [pscustomobject]@{ Raw='Unavailable'; UTC=''; Comparable=$false }
    $notes = [System.Collections.Generic.List[string]]::new()
    try {
        $destinationList = Get-ExactReviewList -Side Destination -SiteUrl $item.DestinationUrl -ListName $item.DestinationList
        try { $destinationFile = Get-ReviewFile -List $destinationList -RelativePath $destinationLookupPath }
        catch { $notes.Add('Destination Get-File: ' + $_.Exception.Message) }
        if ($destinationFile -and $destinationFile.PSObject.Properties['Address']) {
            $candidate = [string]$destinationFile.Address
            if (Test-ReviewUrl -Url $candidate -SiteUrl $item.DestinationUrl) {
                $url = $candidate
                $urlEvidence = 'Verified by ShareGate Get-File'
            }
        }
        if (-not $url) {
            $url = Get-InferredReviewUrl -List $destinationList -SiteUrl $item.DestinationUrl -RelativePath $destinationLookupPath
            if ($url) {
                $urlEvidence = if ($item.DestinationPath) { 'Inferred from prior destination path and list RootFolder; verify in SPO' }
                               else { 'Inferred from source path and destination list RootFolder; verify in SPO' }
            }
        }
        if ($item.Result -eq 'Skipped') {
            $destinationModified = Convert-ReviewModifiedDate -Value (Get-ReviewModifiedValue -Object $destinationFile) -Side Destination
        }
    }
    catch { $notes.Add('Destination list: ' + $_.Exception.Message) }
    if ($item.Result -eq 'Skipped') {
        try {
            $sourceList = Get-ExactReviewList -Side Source -SiteUrl $item.SourceUrl -ListName $item.SourceList
            $sourceFile = $null
            try { $sourceFile = Get-ReviewFile -List $sourceList -RelativePath $item.SourcePath }
            catch { $notes.Add('Source Get-File: ' + $_.Exception.Message) }
            if (-not $sourceFile) {
                $sourceItem = @(ShareGate\Get-ListItem -List $sourceList -Id $item.SourceItemId -ErrorAction Stop)
                if ($sourceItem.Count -eq 1) { $sourceFile = $sourceItem[0] }
            }
            $sourceModified = Convert-ReviewModifiedDate -Value (Get-ReviewModifiedValue -Object $sourceFile) -Side Source
        }
        catch { $notes.Add('Source item: ' + $_.Exception.Message) }
    }
    $skipAssessment = ''
    if ($item.Result -eq 'Skipped') {
        if ($sourceModified.Comparable -and $destinationModified.Comparable -and $urlEvidence -like 'Verified*') {
            if ([datetime]$destinationModified.UTC -ge [datetime]$sourceModified.UTC) {
                $skipAssessment = 'Consistent with IncrementalUpdate: destination exists and is not older than source.'
            }
            else { $skipAssessment = 'Not explained by modified dates; destination appears older than source.' }
        }
        else { $skipAssessment = 'Not established: item existence or comparable Modified dates unavailable.' }
    }
    $result = [pscustomobject]@{
        SourceItemId=$item.SourceItemId; ItemKey=$item.ItemKey; Result=$item.Result; CopySessionId=$item.CopySessionId;
        ReportPath=$item.ReportPath; DestinationPath=$item.DestinationPath;
        DestinationPathEvidence=$item.DestinationPathEvidence; DestinationItemUrl=$url; UrlEvidence=$urlEvidence;
        SourceModifiedRaw=$sourceModified.Raw; SourceModifiedUtc=$sourceModified.UTC;
        DestinationModifiedRaw=$destinationModified.Raw; DestinationModifiedUtc=$destinationModified.UTC;
        SkipAssessment=$skipAssessment; Notes=($notes -join ' | ')
    }
    $results.Add($result)
    Write-ReviewLog ('ID={0}; Result={1}; CopySession={2}; Export={3}; SPO item={4}; URL evidence={5}' -f
        $result.SourceItemId,$result.Result,$result.CopySessionId,$result.ReportPath,
        $(if ($url) { $url } else { '(unavailable)' }),$urlEvidence)
    if ($item.Result -eq 'Skipped') {
        Write-ReviewLog ('ID={0}; Source Modified={1} (UTC {2}); Destination Modified={3} (UTC {4}); Assessment={5}' -f
            $item.SourceItemId,$sourceModified.Raw,$sourceModified.UTC,$destinationModified.Raw,$destinationModified.UTC,$skipAssessment)
    }
}
$columns = @('SourceItemId','ItemKey','Result','CopySessionId','ReportPath','DestinationPath','DestinationPathEvidence','DestinationItemUrl','UrlEvidence','SourceModifiedRaw','SourceModifiedUtc','DestinationModifiedRaw','DestinationModifiedUtc','SkipAssessment','Notes')
Export-SmartM365TransientCsv -Path $itemsPath -Rows $results.ToArray() -Columns $columns
$summary = [ordered]@{ SchemaVersion=1; ScriptVersion=$version; SessionId=$SessionId; GeneratedAtUtc=([datetime]::UtcNow.ToString('o'));
    AnalysisSHA256=$evidence.AnalysisSHA256; PilotManifestSHA256=$evidence.PilotManifestSHA256;
    Success=@($results | Where-Object Result -EQ 'Success').Count; Skipped=@($results | Where-Object Result -EQ 'Skipped').Count;
    Error=@($results | Where-Object Result -EQ 'Error').Count; ItemsPath=$itemsPath; LogPath=$script:logPath }
Export-SmartM365TransientJson -Path $summaryPath -Value $summary
Write-ReviewLog ('Review completed. Items={0}; Summary={1}' -f $itemsPath,$summaryPath)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBW6jEBgFgRHBXg
# DY/rctzYZzK2M2GarNgB9WiDuEg/TqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJF49t66gy+1Lz7Gg2hWXrGpQrHhRwyx9K+C8CRWopFrMA0GCSqG
# SIb3DQEBAQUABIIBgJek9geTA5qWOBqsB+zWidnYqisK6/0fHTYfeazo/AX8ywao
# wSdvhWtQtDJ5ZM7uIOyeo21Lk//1MB1bsJTXXvPPGShyeOjo6y0a7lAogT0jcMjL
# Pg9k8kiMYc/PzFs7CSjklu5C+CQfpGovHTxRBoVCdgLGR9BVlPs21T64fqM6Y6VH
# bNsbaXiO6h3ApDwfVEvsf7a6u98Pw/q8e7pFgTnYgTPCfG1qkXWBZK1cb9SCPElW
# nQ8eZLg1qu/jUEKB94Rt8oZ1Hta2vULTJoxpYJK8CdzcJzCpFvkhQT9gR8XClrCT
# lRHPWFBlIHyARRsrVS2JGtnLWxvNhbUcST/EGhbZWagdY/tUhlFoEKyjGLa69U6F
# BXxxmdLucbmL1tmuCt/H3q9v1KQYBsCiDUH+0aswdB0/4Oz2gj5BDIQwrY9rLZWt
# /0wkC0LMRm9K0tN4woEzIrugXYIdDAOexpwa9TL3XEhPXNmPosmpu5vMUfyTWHKC
# SCTtpBcBfOLr2JuZxaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxMzMw
# MzRaMC8GCSqGSIb3DQEJBDEiBCDoa9xJOlBeInSVdg4su4TQeqQsM8wRh1kmKzhi
# i+sU7TANBgkqhkiG9w0BAQEFAASCAgCuMw+WU6gZaNafugAMEWrAiW/xCKv+LM02
# 3FLhW9NZRgTL1aNug/tC8YniyiHR+1txy/B5KuY/pWU7YQ2SLrb669kqFRZF9aNO
# Pht5uR1OrZ65NieC7/i0KN6rESX15NcxSeW8/galPJwypoTizLEvUD6Ki/F3q8Rm
# E7BJA/W28wLWm1a1B9T/63XsDYXDMHSVxD/tbwPsPq20FO7Ojwa+u8Fly00mRgh1
# N0jmu5ZXe0ZlXFYucSEjIqNWuF/iHtZo3b3/DVjx87OeFgNbJGN8xm8wpu1AptPp
# fC09HY6kGfGeDJywtr8UWyZnQL9GgIpKa/Op18SvF1QItNyLnY2Vi1aLKPB6HGCs
# /vANDtRQXMQYf7/kkoGrusDuY3/LXjxV7fkvJebUFh0ZEtCklySn0Nj91IXjTBU8
# nQ2EtNk0WW32NVV/Zk1msSxd0D7UFG1qSH1ou4ofSXQupWGaNGwOk+KeEznE6nWp
# 7Wn6fADMeZ09uVVEcg/K6tA104Wspymz23Bme8Pdu0H8MNDHXRLh8q6wAT/0UtuT
# BgTi0qP4D78EnHABcjdxftmkCjl29LEv+SY1qOmEB95wsHhXenUtdKwt5cNkxSMY
# jsNJoLPYb3gC6iJM8sjMika9x+TMjmSFyrJLWnQMfiKGj1XiqfwGOTrqxymJyY1X
# xMFOM85kCw==
# SIG # End signature block
