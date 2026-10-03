<#
.SYNOPSIS
    Recycles only audited root-level SharePoint Online file duplicates.
.DESCRIPTION
    Requires an exact audit hash and count. A real run rechecks the complete
    destination library, then checks each root item again immediately before
    recycling it. The expected file in its subfolder is never modified.
.VERSION
    1.0.0
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AuditCsv,
    [Parameter(Mandatory)][string]$WebUrl,
    [Parameter(Mandatory)][string]$LibraryTitle,
    [Parameter(Mandatory)][ValidateRange(1,1000000)][int]$ExpectedCount,
    [Parameter(Mandatory)][ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedAuditHash,
    [string]$ClientId = '',
    [switch]$Run,
    [switch]$ConfirmRecycle
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$version = '1.0.0'
if ($ConfirmRecycle -and -not $Run) { throw '-ConfirmRecycle applies only with -Run.' }
if ($Run -and -not $ConfirmRecycle) { throw 'Recycling requires -Run -ConfirmRecycle.' }

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$auditPath = (Resolve-Path -LiteralPath $AuditCsv -ErrorAction Stop).ProviderPath
$auditHash = (Get-FileHash -LiteralPath $auditPath -Algorithm SHA256).Hash
if ($auditHash -ne $ExpectedAuditHash.ToUpperInvariant()) { throw 'Audit SHA256 differs from the reviewed value.' }
$rows = @(Import-Csv -LiteralPath $auditPath -Delimiter ';' -Encoding UTF8)
if ($rows.Count -ne $ExpectedCount) { throw "Audit contains $($rows.Count) rows; expected $ExpectedCount." }
$required = @('SourceItemId','Run','CopySessionId','FileName','ExpectedTargetPath','RootDuplicatePath',
    'CorrectPathPresentInPreviousTargetScan','RootDuplicateAbsentInPreviousTargetScan',
    'CorrectPathPresentInNewTargetScan','SizeBytes','Version')
foreach ($column in $required) {
    if ($column -notin @($rows[0].PSObject.Properties.Name)) { throw "Audit column is missing: $column" }
}
$web = [uri]$WebUrl
if ($web.Scheme -ne 'https' -or -not $web.Host.EndsWith('.sharepoint.com', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'WebUrl must be an HTTPS SharePoint Online web.'
}
$webPath = $web.AbsolutePath.TrimEnd('/')
$seenRoots = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$seenIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($row in $rows) {
    $rootPath = [string]$row.RootDuplicatePath
    $correctPath = [string]$row.ExpectedTargetPath
    if (-not $rootPath.StartsWith($webPath + '/', [StringComparison]::OrdinalIgnoreCase) -or
        -not $correctPath.StartsWith($webPath + '/', [StringComparison]::OrdinalIgnoreCase) -or
        $rootPath -eq $correctPath -or
        [IO.Path]::GetFileName($rootPath) -ne [string]$row.FileName -or
        [IO.Path]::GetFileName($correctPath) -ne [string]$row.FileName -or
        -not $seenRoots.Add($rootPath) -or -not $seenIds.Add([string]$row.SourceItemId)) {
        throw "Invalid or duplicate audit path/ID for source item $($row.SourceItemId)."
    }
    foreach ($flag in @('CorrectPathPresentInPreviousTargetScan','RootDuplicateAbsentInPreviousTargetScan','CorrectPathPresentInNewTargetScan')) {
        if ([string]$row.$flag -ne 'True') { throw "Audit evidence is incomplete for ID $($row.SourceItemId): $flag" }
    }
    $size = 0L
    if (-not [long]::TryParse([string]$row.SizeBytes, [ref]$size) -or $size -lt 0 -or
        [string]::IsNullOrWhiteSpace([string]$row.Version)) {
        throw "Invalid size or version in audit for ID $($row.SourceItemId)."
    }
}

$stamp = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'),[guid]::NewGuid().ToString('N')
$output = Join-Path (Split-Path -Parent $auditPath) ('RootDuplicateCleanup-' + $stamp)
New-Item -ItemType Directory -Path $output -ErrorAction Stop | Out-Null
$logPath = Join-Path $output 'Cleanup.log'
$resultPath = Join-Path $output 'Results.csv'
$results = [Collections.Generic.List[object]]::new()
function Write-CleanupLog {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Microsoft.PowerShell.Utility\Write-Host $line
}
function Save-Results {
    $columns = @('SourceItemId','Run','CopySessionId','Status','RootDuplicatePath','ExpectedTargetPath','Detail','CompletedAtUtc')
    $temp = $resultPath + '.tmp'
    if ($results.Count) {
        $results.ToArray() | Select-Object $columns | Export-Csv -LiteralPath $temp -Delimiter ';' -Encoding utf8BOM -NoTypeInformation
    } else {
        [IO.File]::WriteAllText($temp, (($columns | ForEach-Object { '"' + $_ + '"' }) -join ';') + "`r`n", [Text.UTF8Encoding]::new($true))
    }
    Move-Item -LiteralPath $temp -Destination $resultPath -Force
}
function Add-Result {
    param($Row,[string]$Status,[string]$Detail)
    $results.Add([pscustomobject]@{
        SourceItemId=$Row.SourceItemId; Run=$Row.Run; CopySessionId=$Row.CopySessionId;
        Status=$Status; RootDuplicatePath=$Row.RootDuplicatePath;
        ExpectedTargetPath=$Row.ExpectedTargetPath; Detail=$Detail;
        CompletedAtUtc=[datetime]::UtcNow.ToString('o')
    })
    Save-Results
}
Write-CleanupLog ('Script=v{0}; mode={1}; actor={2}\{3}; machine={4}; count={5}; auditSHA256={6}; web={7}; list={8}' -f
    $version,$(if ($Run) { 'Recycle' } else { 'DryRun' }),$env:USERDOMAIN,$env:USERNAME,
    $env:COMPUTERNAME,$rows.Count,$auditHash,$WebUrl,$LibraryTitle)
if (-not $Run) {
    Write-CleanupLog 'Audit, count and hash verified. No SPO connection and no deletion.'
    Write-CleanupLog ('Output={0}' -f $output)
    return
}

$connection = $null
try {
    Import-Module PnP.PowerShell -ErrorAction Stop
    if (-not $ClientId) {
        $authPath = Join-Path $projectRoot 'Config\SPOAuth.local.psd1'
        if (Test-Path -LiteralPath $authPath -PathType Leaf) {
            $auth = Import-PowerShellDataFile -LiteralPath $authPath
            $ClientId = [string]$auth.ClientId
        }
    }
    if (-not $ClientId) { throw 'Interactive PnP authentication requires ClientId or Config\SPOAuth.local.psd1.' }
    $connection = Connect-PnPOnline -Url $WebUrl -ClientId $ClientId -Interactive -ForceAuthentication `
        -ValidateConnection -ReturnConnection -ErrorAction Stop
    if (-not $connection) { throw 'Interactive PnP connection returned no connection.' }
    $context = Get-PnPContext -Connection $connection
    $context.Load($context.Web.CurrentUser)
    $context.ExecuteQuery()
    Write-CleanupLog ('SPO account={0}' -f $context.Web.CurrentUser.LoginName)
    $library = Get-PnPList -Identity $LibraryTitle -Connection $connection -ErrorAction Stop
    if (-not $library -or $library.Title -ne $LibraryTitle) { throw 'The destination library was not found by exact title.' }
    $context.Load($library.RootFolder)
    $context.ExecuteQuery()
    $libraryRoot = [string]$library.RootFolder.ServerRelativeUrl
    if (-not $libraryRoot.StartsWith($webPath + '/', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Library root is outside the requested web: $libraryRoot"
    }
    foreach ($row in $rows) {
        if (-not ([string]$row.RootDuplicatePath).StartsWith($libraryRoot + '/', [StringComparison]::OrdinalIgnoreCase) -or
            ([string]$row.RootDuplicatePath).Substring($libraryRoot.Length + 1).Contains('/') -or
            -not ([string]$row.ExpectedTargetPath).StartsWith($libraryRoot + '/', [StringComparison]::OrdinalIgnoreCase) -or
            -not ([string]$row.ExpectedTargetPath).Substring($libraryRoot.Length + 1).Contains('/')) {
            throw "Audit paths are not root duplicate and subfolder original for ID $($row.SourceItemId)."
        }
    }
    $query = @"
<View Scope='RecursiveAll'>
  <ViewFields>
    <FieldRef Name='FSObjType' />
    <FieldRef Name='FileRef' />
    <FieldRef Name='File_x0020_Size' />
    <FieldRef Name='_UIVersionString' />
    <FieldRef Name='Modified' />
  </ViewFields>
  <RowLimit Paged='TRUE'>2000</RowLimit>
</View>
"@
    $items = @(Get-PnPListItem -List $library -Query $query -PageSize 2000 -Connection $connection -ErrorAction Stop)
    $files = @{}
    foreach ($item in $items) {
        if ([string]$item.FieldValues['FSObjType'] -ne '0') { continue }
        $path = [string]$item.FieldValues['FileRef']
        if (-not $path) { continue }
        if ($files.ContainsKey($path)) { throw "SPO returned a duplicate file path: $path" }
        $files[$path] = $item
    }
    foreach ($row in $rows) {
        $root = $files[[string]$row.RootDuplicatePath]
        $original = $files[[string]$row.ExpectedTargetPath]
        if (-not $root -or -not $original) { throw "Root copy or original missing for ID $($row.SourceItemId). No files were recycled." }
        foreach ($item in @($root,$original)) {
            if ([long]$item.FieldValues['File_x0020_Size'] -ne [long]$row.SizeBytes -or
                [string]$item.FieldValues['_UIVersionString'] -ne [string]$row.Version) {
                throw "Metadata changed for ID $($row.SourceItemId). No files were recycled."
            }
        }
    }
    Write-CleanupLog ('Live preflight passed for {0} exact root copies and their originals. Library root={1}' -f $rows.Count,$libraryRoot)
    foreach ($row in $rows) {
        $rootPath = [string]$row.RootDuplicatePath
        $snapshot = $files[$rootPath]
        try {
            $current = Get-PnPListItem -List $library -Id $snapshot.Id -Fields 'FileRef','File_x0020_Size','_UIVersionString','Modified' `
                -Connection $connection -ErrorAction Stop
            if (-not $current -or [string]$current.FieldValues['FileRef'] -cne $rootPath -or
                [long]$current.FieldValues['File_x0020_Size'] -ne [long]$row.SizeBytes -or
                [string]$current.FieldValues['_UIVersionString'] -ne [string]$row.Version -or
                [string]$current.FieldValues['Modified'] -ne [string]$snapshot.FieldValues['Modified']) {
                throw 'Root item changed since the live preflight.'
            }
            Remove-PnPFile -ServerRelativeUrl $rootPath -Recycle -Force -Connection $connection -ErrorAction Stop
            Add-Result -Row $row -Status 'Recycled' -Detail ('SPO list item ID {0}' -f $snapshot.Id)
            Write-CleanupLog ('Recycled {0}/{1}: Source ID={2}; SPO item ID={3}; path={4}' -f
                $results.Count,$rows.Count,$row.SourceItemId,$snapshot.Id,$rootPath)
        }
        catch {
            Add-Result -Row $row -Status 'Failed' -Detail $_.Exception.Message
            Write-CleanupLog ('FAILED for source ID {0}: {1}' -f $row.SourceItemId,$_.Exception.Message)
            throw
        }
    }
    $after = @(Get-PnPListItem -List $library -Query $query -PageSize 2000 -Connection $connection -ErrorAction Stop)
    $afterPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $after) {
        if ([string]$item.FieldValues['FSObjType'] -eq '0') { [void]$afterPaths.Add([string]$item.FieldValues['FileRef']) }
    }
    foreach ($row in $rows) {
        if ($afterPaths.Contains([string]$row.RootDuplicatePath) -or
            -not $afterPaths.Contains([string]$row.ExpectedTargetPath)) {
            throw "Post-recycle verification failed for source ID $($row.SourceItemId); inspect $resultPath."
        }
    }
    $summaryPath = Join-Path $output 'Summary.json.txt'
    [ordered]@{
        ScriptVersion=$version; Expected=$rows.Count; Recycled=$results.Count;
        AuditSHA256=$auditHash; WebUrl=$WebUrl; LibraryTitle=$LibraryTitle;
        SpoAccount=[string]$context.Web.CurrentUser.LoginName;
        OriginalsPreserved=$true; RecycleBinUsed=$true;
        ResultsCsv=$resultPath; CompletedAtUtc=[datetime]::UtcNow.ToString('o')
    } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath ($summaryPath + '.tmp') -Encoding UTF8
    Move-Item -LiteralPath ($summaryPath + '.tmp') -Destination $summaryPath -ErrorAction Stop
    Write-CleanupLog ('COMPLETE: recycled={0}/{1}; originals preserved; report={2}' -f $results.Count,$rows.Count,$resultPath)
}
catch {
    Write-CleanupLog ('FAILED: {0}; partial report={1}' -f $_.Exception.Message,$resultPath)
    throw
}
finally {
    if ($connection) { Disconnect-PnPOnline -ErrorAction SilentlyContinue }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB4aO6LY2S+WzfG
# zscvHim2XrRxvamro0NHIT93R+l/AqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII5XVpKlerfFGeUin0CojPd72FmIJqOOsbif7pvDslFlMA0GCSqG
# SIb3DQEBAQUABIIBgE08STiqMrLVPt9MNyxF0ziqNF68QWPg95eKUoTWjuGLk7Ad
# EChCW/ISELmPw/Bl/U37y6c1HV7kFgpZgbjWho6VJuKF9SE6MnvE4za0zj1OCuvO
# fsoncYKr0kEJesMTDuRLnSDfO5vShMs/k+rzTg4VHP0vFl25B9r0Aw2SQDbDgsT5
# k/RZNMnTsnx1dIDF5p80b1rFrt26+H+0fu2CUIxh7s1L2xDnPecHkVfMwsnoGgHF
# MpLHMYYTNf+pOpV6Ken3vMZ5KDufeOGQaDElVbnswqIPhl2Ks7UqSy2+duItsEvv
# ycN+uJ2zu0ZOS7U5JVP+1HQB5BxFV3FI6oyDCSH639UpviLerIvwjksAmfPUOA+F
# O7AMGx5ru5/d9E4o9vF06TCRCJ3Q/v+H/XA6eFhVHRZ9DBwlEGHxmQs2vBNdwfX+
# prtS2LvKrYW5Vfqv4X/QxTSv2rAEps6MDcjdfAK2WYzvOTmlnX28dlun6UmbEz5w
# z8Pf1nZkolxnzqAHCqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNTI3
# NDJaMC8GCSqGSIb3DQEJBDEiBCALQoPPQVJeAZ5AjJ4SMgaSQ0qPyLe3HAZpXQzv
# qvgMPjANBgkqhkiG9w0BAQEFAASCAgBBh2e/Ih46sOWXRfxKPQcghfPmZS0gwuqN
# CDycfOECwx4ei5jOkETSLseJgBFAf/Z969ZBCxMht0amwrnswxKeHgd9S8+MZzjb
# LlZ/duDNRXUh0OnbOi0IE1EHsEuN4gP+9eoYdD8OzWZYXuEd7VaEIWsjtHy6iF/O
# F9jywuaX51B9Ly+ZaWgD/+qrrlFOe21G0XQcsUcGyWC6xjpGkvWGJmC72F88A+bo
# tyuwMJLna1RPXWQh4a4WTe8sMQ4NNObwiNy0tGaYUp50/fr/W3SlQ5yErfHTYUBv
# uuz6DW7ADpvJ3ZsrlIOv3eNOY0oHdUWfpVAOb9DwFm6J4g/Jub5D3HItdYqZftwh
# 4KKM+UdO6dGcFetGWUM1jcLDGtd+t8QBjkn1zliI05Ty8SB4b+j4vY0AuiEc6pjz
# R2AE+YE+8t60ur0yGpb87ibWY9cJpadtOXYKLZDRoE0R2xd7j1ZWdzqEgupqJ9vy
# 7NKPG2txZvNg8M/o4sh7q8ZtfqzrRwJcutUgDkSpnG62ueuFGI1hWBdYmmZOC7z7
# roR82W2KgPQSVfUqK6oqjbG8AWu7sP/v1qyqpbOx3e1i5iYElyr0Rnfnv/+XH9O+
# eUHTUXZ+c5ux4qVncKN5x+1hQdhrXn+hg4xVArt0S3V8r1uz0H2qBlF1r7tADbDR
# gAWaZqCy/w==
# SIG # End signature block
