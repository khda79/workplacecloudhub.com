<#
.SYNOPSIS
Removes only approved, revalidated files from a server DATA second-pass report.

.DESCRIPTION
Preview creates a SHA256-pinned plan outside the DATA root. Execute requires that
plan, the original evidence CSVs, and explicit expected counts. Every source and
retained copy is checked again before any deletion. Directories are never removed.
#>
#requires -Version 7.2
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DataRootPath,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$EvidenceDirectory,
    [Parameter(Mandatory)][datetimeoffset]$CutoffUtc,
    [Parameter(Mandatory)][ValidateRange(0,100000)][int]$ExpectedUsageFiles,
    [Parameter(Mandatory)][ValidateRange(0,100000)][int]$ExpectedEntraScripts,
    [Parameter(Mandatory)][ValidateRange(0,100000)][int]$ExpectedTestOutputs,
    [Parameter(Mandatory)][ValidateRange(0,100000)][int]$ExpectedTransportLogs,
    [string]$ReportDirectory = '',
    [string]$PlanPath = '',
    [switch]$Execute
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$startedAt = Get-Date
$tenantContextPath = Join-Path $PSScriptRoot '..\..\Config\SmartM365-TenantContext.ps1'
if (Test-Path -LiteralPath $tenantContextPath -PathType Leaf) {
    . $tenantContextPath
    Write-SmartM365StartupBanner
}

$root = [IO.Path]::GetFullPath($DataRootPath).TrimEnd('\', '/')
if (-not (Test-Path -LiteralPath $root -PathType Container) -or
    -not [string]::Equals((Split-Path $root -Leaf), $Tenant, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'DataRootPath must be an existing tenant directory whose leaf matches -Tenant.'
}
if ((Get-Item -LiteralPath $root -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
    throw 'A reparse-point DATA root is not supported.'
}
$rootPrefix = $root + [IO.Path]::DirectorySeparatorChar
function Assert-OutsideData([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\', '/')
    if ($full.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        [string]::Equals($full, $root, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path must be outside the DATA root: $full"
    }
    return $full
}
$evidenceRoot = Assert-OutsideData $EvidenceDirectory
if (-not (Test-Path -LiteralPath $evidenceRoot -PathType Container)) { throw "Evidence directory not found: $evidenceRoot" }
if (-not $ReportDirectory) {
    $ReportDirectory = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ServerDataSecondPassCleanup_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$reportRoot = Assert-OutsideData $ReportDirectory
New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null
if ($Execute) {
    if (-not $PlanPath) { throw '-Execute requires the plan.csv produced by a previous Preview.' }
    $PlanPath = Assert-OutsideData $PlanPath
    if (-not (Test-Path -LiteralPath $PlanPath -PathType Leaf)) { throw "Plan not found: $PlanPath" }
} elseif ($PlanPath) { throw '-PlanPath is used only with -Execute.' }

function Get-DataFile([string]$Relative, [string]$Pattern) {
    if ($Relative -notmatch $Pattern) { throw "Unexpected DATA path: $Relative" }
    $path = [IO.Path]::GetFullPath((Join-Path $root $Relative))
    if (-not $path.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path escapes DATA root: $Relative"
    }
    $cursor = $path
    while (-not [string]::Equals($cursor, $root, [StringComparison]::OrdinalIgnoreCase)) {
        if (-not (Test-Path -LiteralPath $cursor)) { throw "Missing DATA path: $Relative" }
        if ((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Linked DATA path: $Relative"
        }
        $cursor = Split-Path -Path $cursor -Parent
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Not a DATA file: $Relative" }
    return Get-Item -LiteralPath $path -Force
}
function Assert-ReportedMetadata([object]$File, [object]$Row) {
    if ($File.Length -ne [int64]$Row.Bytes -or
        $File.LastWriteTimeUtc.ToString('o') -cne [string]$Row.LastWriteTimeUtc) {
        throw "File differs from the reviewed report: $($File.FullName)"
    }
}
function Get-Hash([object]$File) { (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256).Hash }
function Add-PlanRow([string]$Category, [string]$Relative, [object]$File,
    [string]$RetainedCopy = '', [string]$RetainedSHA256 = '') {
    $plan.Add([pscustomobject]@{
        Category = $Category
        RelativePath = $Relative
        Bytes = [int64]$File.Length
        LastWriteTimeUtc = $File.LastWriteTimeUtc.ToString('o')
        SHA256 = Get-Hash $File
        RetainedCopy = $RetainedCopy
        RetainedSHA256 = $RetainedSHA256
    })
}
function Import-Evidence([string]$Name) {
    $path = Join-Path $evidenceRoot $Name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing evidence CSV: $path" }
    return @(Import-Csv -LiteralPath $path)
}

$plan = [Collections.Generic.List[object]]::new()
$cutoff = $CutoffUtc.UtcDateTime
$usageRows = @(Import-Evidence 'usage-files.csv' | Where-Object {
    $_.Decision -in @('DuplicateOfWeekly','DuplicateOfLaterRun') -and
    [datetime]::ParseExact($_.Run, 'yyyyMMdd_HHmmss', [Globalization.CultureInfo]::InvariantCulture) -lt $cutoff
})
if ($usageRows.Count -ne $ExpectedUsageFiles) {
    throw "Usage evidence count $($usageRows.Count) differs from expected $ExpectedUsageFiles."
}
foreach ($row in $usageRows) {
    $relative = [string]$row.Path
    $file = Get-DataFile $relative '^DATA-ALL\\M365\\Usage\\\d{8}_\d{6}\\[^\\]+\.csv$'
    if ($file.Length -ne [int64]$row.Bytes) { throw "Usage source size changed: $relative" }
    if ($relative -cnotlike ('DATA-ALL\M365\Usage\' + $row.Run + '\*') -or
        [string]$row.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or
        (Get-Hash $file) -cne [string]$row.SHA256) { throw "Usage source mismatch: $relative" }
    $copyRelative = [string]$row.RetainedCopy
    $copy = Get-DataFile $copyRelative '^DATA-ALL\\M365\\Usage\\(?:WeeklyHistory\\\d{4}-W\d{2}|\d{8}_\d{6})\\[^\\]+\.csv$'
    if ($copy.FullName -eq $file.FullName -or $copy.Length -ne $file.Length -or
        (Get-Hash $copy) -cne [string]$row.SHA256) { throw "Usage retained copy mismatch: $relative" }
    if ($row.Decision -eq 'DuplicateOfWeekly' -and
        $copyRelative -cnotlike ('DATA-ALL\M365\Usage\WeeklyHistory\' + $row.Week + '\*')) {
        throw "Usage weekly copy mismatch: $relative"
    }
    if ($row.Decision -eq 'DuplicateOfLaterRun' -and
        ($copyRelative -notmatch '^DATA-ALL\\M365\\Usage\\(?<Later>\d{8}_\d{6})\\' -or
        $Matches.Later -cle $row.Run)) { throw "Usage later-run copy mismatch: $relative" }
    Add-PlanRow 'UsageExactDuplicate' $relative $file $copyRelative ([string]$row.SHA256)
}

$entraRows = @(Import-Evidence 'entra-generated-scripts.csv' | Where-Object {
    [datetime]::Parse([string]$_.LastWriteTimeUtc, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime() -lt $cutoff
})
if ($entraRows.Count -ne $ExpectedEntraScripts) {
    throw "Entra evidence count $($entraRows.Count) differs from expected $ExpectedEntraScripts."
}
foreach ($row in $entraRows) {
    if ($row.Decision -ne 'ManualOperatorUseReview') { throw "Unexpected Entra decision: $($row.Path)" }
    $file = Get-DataFile ([string]$row.Path) '^DATA-ALL\\M365\\Devices\\EntraDevices\\SmartM365-EntraDevices-RemoveCandidates_\d{8}_\d{6}\.ps1$'
    Assert-ReportedMetadata $file $row
    Add-PlanRow 'OldGeneratedEntraScript' ([string]$row.Path) $file
}

$lastRows = @(Import-Evidence 'data-last-old.csv')
$testRows = @($lastRows | Where-Object Decision -eq 'ReviewOldTestOutput')
$logRows = @($lastRows | Where-Object Decision -eq 'ReviewOldTransportLog')
if ($testRows.Count -ne $ExpectedTestOutputs -or $logRows.Count -ne $ExpectedTransportLogs) {
    throw "DATA-LAST evidence count differs: test=$($testRows.Count), log=$($logRows.Count)."
}
foreach ($row in @($testRows) + @($logRows)) {
    $name = [string]$row.Name
    if ($row.Decision -eq 'ReviewOldTestOutput') {
        if ($name -notmatch '_MAXITEMS[-_]\d+' -or $name -notmatch '\.(csv|json)$') {
            throw "Unexpected test output name: $name"
        }
    } elseif ($name -cne 'Intune_DiscoveredApps_AppDeviceRelations.cache.json.migration.log') {
        throw "Unexpected transport log name: $name"
    }
    $relative = 'DATA-LAST\' + $name
    $file = Get-DataFile $relative '^DATA-LAST\\[^\\/]+$'
    Assert-ReportedMetadata $file $row
    Add-PlanRow $(if ($row.Decision -eq 'ReviewOldTestOutput') { 'OldTestOutput' } else { 'OldTransportLog' }) $relative $file
}

$planned = @($plan.ToArray() | Sort-Object Category, RelativePath)
$paths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($item in $planned) {
    if (-not $paths.Add($item.RelativePath)) { throw "Duplicate planned path: $($item.RelativePath)" }
}
foreach ($item in $planned) {
    if ($item.RetainedCopy -and $paths.Contains($item.RetainedCopy)) {
        throw "Retained copy is also planned for deletion: $($item.RetainedCopy)"
    }
}

if (-not $Execute) {
    $outputPlan = Join-Path $reportRoot 'plan.csv'
    $planned | Export-Csv -LiteralPath $outputPlan -NoTypeInformation -Encoding utf8
    $result = [pscustomobject]@{Mode='Preview';Planned=$planned.Count;Removed=0;PlanPath=$outputPlan;ReportDirectory=$reportRoot}
} else {
    $approved = @(Import-Csv -LiteralPath $PlanPath)
    if ($approved.Count -ne $planned.Count) { throw 'The preview plan count differs from current evidence.' }
    $columns = @('Category','RelativePath','Bytes','LastWriteTimeUtc','SHA256','RetainedCopy','RetainedSHA256')
    for ($i = 0; $i -lt $planned.Count; $i++) {
        foreach ($column in $columns) {
            if ([string]$approved[$i].$column -cne [string]$planned[$i].$column) {
                throw "The preview plan differs from current evidence at row $($i + 1): $column"
            }
        }
    }
    $results = [Collections.Generic.List[object]]::new()
    try {
        foreach ($item in $planned) {
            $file = Get-DataFile $item.RelativePath '^DATA-(?:ALL|LAST)\\'
            Assert-ReportedMetadata $file $item
            if ((Get-Hash $file) -cne $item.SHA256) { throw "Changed before removal: $($item.RelativePath)" }
            if ($item.RetainedCopy) {
                $copy = Get-DataFile $item.RetainedCopy '^DATA-ALL\\M365\\Usage\\'
                if ((Get-Hash $copy) -cne $item.RetainedSHA256) {
                    throw "Retained copy changed: $($item.RetainedCopy)"
                }
            }
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            if (Test-Path -LiteralPath $file.FullName) { throw "File remains after removal: $($item.RelativePath)" }
            $results.Add([pscustomobject]@{Category=$item.Category;RelativePath=$item.RelativePath;Result='Removed'})
        }
    } finally {
        $results.ToArray() | Export-Csv -LiteralPath (Join-Path $reportRoot 'results.csv') -NoTypeInformation -Encoding utf8
    }
    $result = [pscustomobject]@{Mode='Execute';Planned=$planned.Count;Removed=$results.Count;PlanPath=$PlanPath;ReportDirectory=$reportRoot}
}
$result
if (Get-Command Write-SmartM365CompletionBanner -ErrorAction SilentlyContinue) {
    Write-SmartM365CompletionBanner -Status Success -ScriptName 'SmartM365-ServerDataSecondPassCleanup' -StartedAt $startedAt -EndedAt (Get-Date)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAEl0lLpdM9s3ZZ
# tqTdLUITVjkb94kFCte1y2dYJMu1qqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIG5Q8kTo6EtQKK0o9dIa4OvIoujZZPDJFEijwr31HnuIMA0GCSqG
# SIb3DQEBAQUABIIBgJM2FkDK5vu3uj6NwNU4UmWxT30EhBSeXzrKt1Xk917HfyuN
# N3HbFJks5xz4s3563cwxkcd4R2013yDG+UMAduUDg5zRZotOe+uwEcsyDsVxrjTv
# Nk3PZJOWBb7n65WRyo7CB+K1D+bmf/Rlx+NCMAtJFmk+Fd8NIe86UfsS+qQl9ax9
# W39s7tkn0bIb5nSmQa5Y2PM5PqeZrRLzvVs4S1xjAP3Mc0wOdoPzM2QLdqCDxJYs
# SjepLr/mVE6aAz7yJNERhINkwVRmWSjOw08tHG1IXr+so4plN3/bhwlbM3HLEoAG
# iMRRAI/t79fqPInyz3FOSY9ROHGOcth2CgDkNjMC1XzEs3+2SNWLxH1+aJFYvI66
# l7F6xNyM27CLg7KNXCAqcZysoRPHhh3nwz1wfA4ZNVQkUVPF+mVItWqdVzAOcQRS
# O55ZX8vZmK4Tlw01Sg9ZyvYALx5cQfTewpzzpjl4GBTZSiKJk80Rgppp44GKLQlP
# HH8OcaI7OZYwSOCGh6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxODQw
# MzVaMC8GCSqGSIb3DQEJBDEiBCDRwJregg7wlVEJN8aZvuiRKu01MC3oZdjDMuRS
# kE9GnjANBgkqhkiG9w0BAQEFAASCAgCTVShpijX30tlkM21tKeFrI/xptYjowgJb
# kt2Qv4yn9e3ba4TvJJYUejt69p2UJBjplHGrKyZ1Y6OPDKl3HyIn8RYT/qo6Tn4i
# l8R5sPGOWBgF8yOLGRYlfKhV6FCIeFLj85qjrOSRfwXJ73z2qU64LCBnmOqsKV4t
# BScsZVx4VTTShiFErjTVqasVpI1gUF9R3GGlJyszQQnZXM9zBvuaJ2ZLwIVWHkvn
# 6WrwHfFQDUt+BI1XzZ5qXx4flAOtsZoVeSZPKkVazNI8tB75+cS8hLZu/4zq5iLH
# 7wii0BcQySJRxyXr1ZbCzkDJU7x2UsjPzFGeYFHn6CiPV6gRbzIKvYwbzn0AJoYx
# ZVidH4z5H0JdM1cI0FjuWGFFM2ixqTUE49C8ylLGWMkMUl78K9VzkefZgi8GsBpr
# SL6dkp1LKhY6WyjDHfh08pcjC2k2munAHwXLr9kxClxqRBoXKvmax8jr1Cimr+ZK
# 1mxBgvOnvGbnw2WAxFMW8kbuRSla1YxwRsJks/ebO4YMYhKO1rGv30hiX5WcRdcW
# 51y4JALq8/8ZrWAFn1uvvCHUa0KXcOIdP+Mb9eSSgVQkG9dCXkkGIHuc789Lgo+3
# D1iq+e8Ss3EjL3GpirkMXJ8Erd36w03JV3jJxwRl/bHCTW+rNOIEcgC/CiEUwzwn
# +3gG4ipvxw==
# SIG # End signature block
