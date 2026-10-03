<#
.SYNOPSIS
    Qualifies one item-scoped ShareGate copy into an explicit SPO folder.
.VERSION
    1.0.1
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ProjectRoot,
    [Parameter(Mandatory)][string]$AnalysisDirectory,
    [Parameter(Mandatory)][string]$SessionId,
    [Parameter(Mandatory)][ValidateRange(1,2147483647)][int]$SourceItemId,
    [ValidateNotNullOrEmpty()][string]$FarmTimeZoneId = 'W. Europe Standard Time',
    [ValidatePattern('^[0-9A-Fa-f]{64}$')][string]$ExpectedAnalysisHash = '',
    [switch]$DryRun,
    [switch]$Run,
    [switch]$ConfirmQualification
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$version = '1.0.1'
. (Join-Path $PSScriptRoot '..\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1')
$script:ConsoleLifecycleContext = Start-SmartM365MigrationConsoleLifecycle -ScriptPath $PSCommandPath
$script:ConsoleLifecycleFailure = $null
$script:ConsoleLifecycleStatus = 'SUCCESS'
try {
    if ($DryRun -and $Run) { throw 'Choose either -DryRun or -Run.' }
    if ($ConfirmQualification -and -not $Run) { throw '-ConfirmQualification applies only with -Run.' }
    if ($Run -and (-not $ConfirmQualification -or -not $ExpectedAnalysisHash)) {
        throw 'A real qualification requires -Run, -ConfirmQualification, and the reviewed analysis SHA256.'
    }
    if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
        throw 'Use Windows PowerShell 5.1 (powershell.exe) for ShareGate.'
    }
    . (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-FarmMaintenance.ps1')
    . (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-DestinationPath.ps1')
    . (Join-Path $PSScriptRoot 'SmartM365-SharePointMigration-ShareGateReportReader.ps1')
    $reportAliases = Get-SmartM365ShareGateReportAliases -ConfigRoot (Join-Path $PSScriptRoot '..\..\Config')
    $farmZone = [TimeZoneInfo]::FindSystemTimeZoneById($FarmTimeZoneId)
    $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase 'path qualification preparation'
    $project = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
    $analysis = (Resolve-Path -LiteralPath $AnalysisDirectory -ErrorAction Stop).ProviderPath
    $diagnosticsRoot = Join-Path $project 'ShareGate\Diagnostics'
    if (-not $analysis.StartsWith($diagnosticsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'AnalysisDirectory must be under this project ShareGate\Diagnostics folder.'
    }
    $classifiedPath = Join-Path $analysis 'ClassifiedRows.csv'
    $hash = (Get-FileHash -LiteralPath $classifiedPath -Algorithm SHA256).Hash
    if ($ExpectedAnalysisHash -and $hash -ne $ExpectedAnalysisHash.ToUpperInvariant()) { throw 'ClassifiedRows.csv hash differs from the reviewed analysis.' }
    $matches = @(Import-Csv -LiteralPath $classifiedPath -Encoding UTF8 | Where-Object {
        $_.SessionId -eq $SessionId -and $_.SourceItemId -eq [string]$SourceItemId -and
        $_.RuleId -eq 'SG-ACCESS-SOURCE' -and $_.State -eq 'To fix' -and $_.AccessSide -eq 'Source' -and $_.ObjectType -eq 'File'
    })
    $keys = @($matches | ForEach-Object ItemKey | Sort-Object -Unique)
    if ($keys.Count -ne 1) { throw "Source ID $SourceItemId does not identify exactly one eligible source 401 file in session $SessionId." }
    $row = $matches[0]
    if (@($matches | Where-Object { $_.SourceUrl -ne $row.SourceUrl -or $_.SourceList -ne $row.SourceList -or
        $_.DestinationUrl -ne $row.DestinationUrl -or $_.DestinationList -ne $row.DestinationList -or
        $_.'Raw: Source path' -ne $row.'Raw: Source path' -or $_.'Raw: Destination path' -ne $row.'Raw: Destination path' }).Count) {
        throw 'Rows for this item disagree on source or destination routing.'
    }
    $route = Resolve-SmartM365ShareGateDestinationPath -Row $row
    if (-not $route.DestinationFolder) { throw 'Qualification requires an item in a destination subfolder.' }
    $sourceUri = [uri]$row.SourceUrl
    $destinationUri = [uri]$row.DestinationUrl
    if ($sourceUri.Host -eq $destinationUri.Host -or $destinationUri.Scheme -ne 'https' -or
        $destinationUri.Host -notmatch '(?i)\.sharepoint\.(com|cn|de)$') {
        throw 'Destination must be a distinct HTTPS SharePoint Online host.'
    }
    Write-Output ('{0} Mode={1}; item=1; source writes=none; destination write={2}; Session={3}; AnalysisSHA256={4}' -f
        (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $(if ($Run) { 'Run' } else { 'DryRun' }),
        $(if ($Run) { 'one overwrite after confirmation' } else { 'none' }), $SessionId, $hash)
    Write-Output ('ID={0}; source={1} | {2} | {3}; destination={4} | {5} | {6}; folder={7}' -f
        $SourceItemId, $row.SourceUrl, $row.SourceList, $route.SourceFilePath,
        $row.DestinationUrl, $row.DestinationList, $route.DestinationFilePath, $route.DestinationFolder)
    if (-not $Run) { return }

    $phrase = 'QUALIFY ONE ITEM ' + $SourceItemId + ' v' + $version
    $entered = Read-Host ('Type exactly "{0}" to authorize one SPO destination overwrite' -f $phrase)
    if ($entered -cne $phrase) { throw 'Qualification confirmation was not entered exactly. No copy was started.' }
    $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase 'path qualification startup'

    $runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
    $output = Join-Path $diagnosticsRoot ('PathQualification-' + $runId)
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    $logPath = Join-Path $output 'PathQualification.log'
    $reportPath = Join-Path $output 'ShareGate-Report.csv'
    $objectPath = Join-Path $output 'CopyResult.txt'
    $resultPath = Join-Path $output 'PathQualification-Result.json.txt'
    $result = [ordered]@{
        ScriptVersion=$version; Actor=($env:USERDOMAIN + '\' + $env:USERNAME); Machine=$env:COMPUTERNAME;
        SessionId=$SessionId; SourceItemId=$SourceItemId; AnalysisSHA256=$hash;
        SourceUrl=$row.SourceUrl; SourceList=$row.SourceList; SourceFilePath=$route.SourceFilePath;
        DestinationUrl=$row.DestinationUrl; DestinationList=$row.DestinationList;
        DestinationFilePath=$route.DestinationFilePath; DestinationFolder=$route.DestinationFolder;
        SourceRead='Not attempted'; DestinationBefore='Not attempted'; RootBefore='Not attempted';
        CopySessionId=''; ReportPath=''; ExportedDestinationPath=''; ShareGateResult='';
        DestinationAfter='Not attempted'; RootAfter='Not attempted'; Qualification='Not proven'; Error=''
    }
    function Write-QualificationLog {
        param([string]$Message)
        $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
        Write-Output $line
    }
    function Save-QualificationResult {
        $temporary = Join-Path $output ('.' + [guid]::NewGuid().ToString('N') + '.tmp')
        try {
            ($result | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $temporary -Encoding UTF8
            Move-Item -LiteralPath $temporary -Destination $resultPath -Force
        }
        finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
    }
    function Get-OneShareGateFile {
        param($List, [string]$Path, [switch]$AllowMissing)
        try { $files = @(ShareGate\Get-File -List $List -Path $Path -ErrorAction Stop) }
        catch {
            if ($AllowMissing -and $_.Exception.Message -match '(?i)(notfound|not found|could not be found|cannot find|does not exist|doesn.t exist|\b404\b|introuvable|n.existe pas|n.a pas .t. trouv.)') { return $null }
            throw
        }
        if ($files.Count -gt 1) { throw "Get-File returned multiple files for $Path." }
        if ($files.Count -eq 0) {
            if ($AllowMissing) { return $null }
            throw "Get-File returned no file for $Path."
        }
        return $files[0]
    }
    function Get-ExactShareGateList {
        param($Site, [string]$Name)
        $found = @(ShareGate\Get-List -Site $Site -Name $Name -ErrorAction Stop)
        $exact = @($found | Where-Object { ($_.PSObject.Properties['Title'] -and $_.Title -eq $Name) -or
            ($_.PSObject.Properties['Name'] -and $_.Name -eq $Name) })
        if ($exact.Count -ne 1) { throw "Get-List did not resolve exactly one list named '$Name'." }
        return $exact[0]
    }
    try {
        Write-QualificationLog ('Actor={0}; Machine={1}; SourceReadOnly=True; DestinationCopyLimit=1; Mode=Overwrite; AnalysisSHA256={2}' -f
            $result.Actor, $result.Machine, $hash)
        $module = @(Get-Module -ListAvailable -Name ShareGate | Sort-Object Version -Descending | Select-Object -First 1)
        if ($module.Count -ne 1) { throw 'ShareGate module is not installed.' }
        Import-Module -Name $module[0].Path -ErrorAction Stop
        Write-QualificationLog ('ShareGate module version={0}; path={1}' -f $module[0].Version, $module[0].Path)
        foreach ($name in @('Connect-Site','Get-List','Get-Folder','Get-File','Copy-Content','New-CopySettings','Export-Report')) {
            if (-not (Get-Command -Name $name -Module ShareGate -ErrorAction SilentlyContinue)) { throw "Required ShareGate cmdlet is missing: $name" }
        }
        $copy = Get-Command Copy-Content -Module ShareGate
        $required = @('SourceList','DestinationList','SourceItemId','DestinationFolder','CopySettings','TaskName')
        if (-not @($copy.ParameterSets | Where-Object {
            $names = @($_.Parameters | ForEach-Object Name)
            @($required | Where-Object { $_ -notin $names }).Count -eq 0
        }).Count) { throw 'Installed Copy-Content does not support the required item-scoped DestinationFolder parameter set.' }
        $sourceSite = ShareGate\Connect-Site -Url $row.SourceUrl -ErrorAction Stop
        $sourceList = Get-ExactShareGateList -Site $sourceSite -Name $row.SourceList
        $sourceFile = Get-OneShareGateFile -List $sourceList -Path $route.SourceFilePath
        $result.SourceRead = [string]$sourceFile.Address
        $destinationSite = ShareGate\Connect-Site -Url $row.DestinationUrl -Browser -ErrorAction Stop
        $destinationList = Get-ExactShareGateList -Site $destinationSite -Name $row.DestinationList
        Assert-SmartM365ShareGateDestinationFolder -DestinationList $destinationList -DestinationFolder $route.DestinationFolder
        $before = Get-OneShareGateFile -List $destinationList -Path $route.DestinationFilePath
        $result.DestinationBefore = [string]$before.Address
        $rootBefore = Get-OneShareGateFile -List $destinationList -Path $route.FileName -AllowMissing
        if ($rootBefore) { throw 'A file with the same name already exists at the destination library root; cannot use it as an unambiguous placement control.' }
        $result.RootBefore = 'Absent'
        $null = Assert-SmartM365OutsideFarmMaintenance -FarmTimeZone $farmZone -Phase 'one-item ShareGate copy'
        $settings = ShareGate\New-CopySettings -OnContentItemExists Overwrite -ErrorAction Stop
        $taskName = 'SmartM365 path qualification ' + $SessionId + ' ID=' + $SourceItemId + ' ' + $runId
        Write-QualificationLog ('Copy-Content starts: ID={0}; DestinationFolder={1}; TaskName={2}' -f $SourceItemId,$route.DestinationFolder,$taskName)
        $copyResult = ShareGate\Copy-Content -SourceList $sourceList -DestinationList $destinationList -SourceItemId @($SourceItemId) -DestinationFolder $route.DestinationFolder -CopySettings $settings -TaskName $taskName -ErrorAction Stop
        if (-not $copyResult -or @($copyResult).Count -ne 1) { throw 'Copy-Content did not return exactly one CopyResult.' }
        @('Type: ' + $copyResult.GetType().FullName, '', ($copyResult | Format-List * -Force | Out-String -Width 4096)) |
            Set-Content -LiteralPath $objectPath -Encoding UTF8
        foreach ($name in @('SessionId','SessionID','CopySessionId','Id')) {
            $property = $copyResult.PSObject.Properties[$name]
            if ($property -and [string]$property.Value -match '^\d{6}-\d+$') { $result.CopySessionId = [string]$property.Value; break }
        }
        ShareGate\Export-Report -CopyResult $copyResult -Path $reportPath -ErrorAction Stop | Out-Null
        if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) { throw 'Export-Report did not create a CSV.' }
        $result.ReportPath = $reportPath
        $reportRows = @(Import-Csv -LiteralPath $reportPath -Encoding UTF8)
        if (-not $reportRows.Count) { throw 'ShareGate export contains no item row; placement is not proven.' }
        $itemRows = $reportRows
        if (Test-SmartM365ShareGateReportField -Row $reportRows[0] -Field 'SourceItemId' -Aliases $reportAliases) {
            $itemRows = @($reportRows | Where-Object {
                (Get-SmartM365ShareGateReportValue -Row $_ -Field 'SourceItemId' -Aliases $reportAliases) -eq [string]$SourceItemId
            })
        }
        if (-not @($itemRows).Count) { throw 'ShareGate export contains no row for the selected source item.' }
        $statuses = @($itemRows | ForEach-Object {
            Get-SmartM365ShareGateReportValue -Row $_ -Field 'Status' -Aliases $reportAliases
        } | Where-Object { $_ } | Sort-Object -Unique)
        $result.ShareGateResult = $statuses -join '; '
        $paths = @($itemRows | ForEach-Object {
            Get-SmartM365ShareGateReportValue -Row $_ -Field 'DestinationPath' -Aliases $reportAliases
        } | Where-Object { $_ } | Sort-Object -Unique)
        if ($paths.Count -eq 1) { $result.ExportedDestinationPath = [string]$paths[0] }
        $after = Get-OneShareGateFile -List $destinationList -Path $route.DestinationFilePath
        $result.DestinationAfter = [string]$after.Address
        $rootAfter = Get-OneShareGateFile -List $destinationList -Path $route.FileName -AllowMissing
        $result.RootAfter = if ($rootAfter) { [string]$rootAfter.Address } else { 'Absent' }
        if ($rootAfter) { throw 'A file appeared at the destination library root; placement failed. No other item will be copied.' }
        if ($statuses.Count -ne 1 -or $statuses[0] -notmatch '(?i)^success$') { throw 'ShareGate did not report one unambiguous Success result.' }
        if ($result.ExportedDestinationPath -and
            (ConvertTo-SmartM365ShareGateRelativePath -Path $result.ExportedDestinationPath -Side ExportedDestination) -ne $route.DestinationFilePath) {
            throw 'ShareGate export destination path differs from the reviewed target path.'
        }
        $result.Qualification = if ($result.ExportedDestinationPath) { 'Passed' } else { 'Inconclusive - export path blank' }
        Write-QualificationLog ('Qualification={0}; ShareGate={1}; Expected={2}; Root={3}; Report={4}' -f
            $result.Qualification,$result.ShareGateResult,$result.DestinationAfter,$result.RootAfter,$reportPath)
    }
    catch {
        $result.Error = $_.Exception.Message
        Write-QualificationLog ('Qualification failed or inconclusive: ' + $result.Error)
        throw
    }
    finally {
        Save-QualificationResult
        Write-Output ('Result: {0}; Report: {1}; Log: {2}' -f $resultPath,$result.ReportPath,$logPath)
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCEuT4TAT37l2Yf
# 2ceNLCDbZ5D/A6BYdmRZulPVp6+JqKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIH2JEQ4BcOTYNumHOjkN2uakIXwIg0eYVRWb3qc8N92QMA0GCSqG
# SIb3DQEBAQUABIIBgFcpT9Z01PLnfahT6xCw2h86Wq7gE42+mKMLi9iB3bzsN/Ed
# DBMd4nDu5ErVbSSQFg2v4A3SCsC/szoLida57YhfPCMCWWJvnWcNS+mmkrosNB9W
# nQgpTgRftQePwLOzlM3gWiaFfE3iHFkn3jtymcqj22MGrsxZ0+7KYr2MUVJJRsz4
# Kct9hVB+Hp7MZ0kYcF7PiQ2QEaNdlHgB4bBws3JsoYyr6MJGl1YZ+KnVrmf+mGGM
# uuMdYCilTNQH8mygRyRU+Xeqxfg4SrSEQfs55Y12ibmPM5QoAl5Ga0etsyJGrhT6
# KrpA3NN/uxw3418/wgy4J7/XQ95IWuZJeyNjL2Yz4+Cl01/tjgeE7G82dbxF7wwp
# 4UAOwoWPRJmQe+MAPDVGRkk/5OXS1ywnZwkTCw9kkBzLsCLnx4/X43HwiV1uPGiI
# AwMrvLYZRxLzI95zT4Mh453kDqLWAKKM/GXsG2Yzw51FxXzD0nTkljuJIMnfTNEA
# NW4IHAVrSg0xLhn296GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxNzA2
# MTFaMC8GCSqGSIb3DQEJBDEiBCDKAWmgUdSwExKUnywiezURFVh5AmN/de9lJn/f
# fUE5PzANBgkqhkiG9w0BAQEFAASCAgAU9KRSMhxWlL6docc4PTKD1JGXXNTHx9A/
# A2MO7+d1/gHckjEYATTPXzK5/AlqXRh9pZ3VG03PFXHcPvt3pZTkqJ08Xe2D0bq/
# YUw3gRBlrHst5yuKVaW1Xb3u1IY6XmviKvWymy21xONJxgfhHnyhFsM4PVEI15r1
# HURCgP+/YV3uZ2XspzfJl/qiWEhEJKJDWCP/5/OLLE4WpPxYEyPQQAKhEBjs/t5p
# Nd7mp+FkugNDIM0B5lOZCEOQFaV8BFcCEaGAYY9fK/OKc+6CL/AKJNz3wnrn85b5
# kMa/9rFgNtycbvWh8eA0Ia5XT9mgvtUjgnbnmCJzZnA7tuh0vmtPXUJwFWhSI2p0
# CPsL3pbn8K0QZPhrsQXw6Zg9PPBdIR1+JyTgSbqH73u4vFCeSXZmODAnxNSaVF6z
# Y8Pg6EfoFxpW7+GWYkGLkJ6GQ+Sox4KLOAmt0ATrd+pmVtFrUgefQNU7FDebT6CR
# CDkhxztx8dKxDaOGQOhybsibUYGn13zNPOSvmkcwjcQpVUlXIq6iUNVK7qxMFB1J
# kPkg+B4T8QevRr2VfyL+PAOUMB0rTBoyGWqZmPS/90Xv7RGUOCvv8jz3/OcUvXLF
# OVf4RNWdZCy2i59bPaY0YrCuXtVFTIeyscZgunXQWp2GUwP0H/l54SKgKjRO3n9e
# byfVB5ctyg==
# SIG # End signature block
