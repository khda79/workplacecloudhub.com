#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged DATA-LAST publication for the Teams and SharePoint Online producers: with
the previous run's different CSVs in DATA-LAST, a running or failed run neither changes, uploads nor
historizes them, and a complete run promotes this run's CSVs with the receipt.
.VERSION
1.0.0
.NOTES
Synthetic TEMP data only. SharePoint upload and WeeklyHistory are replaced by recorders.
No tenant, Graph, SharePoint, API, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','',Justification='Script variables are read by the extracted producer functions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Synthetic recorders replace the producer SharePoint calls.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.95' -Force -PassThru
$teams = Join-Path $root 'SmartInventory/M365Inventory/Teams/SmartM365-Teams-Inventory.ps1'
$spo = Join-Path $root 'SmartInventory/M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1'
foreach ($source in @(@($teams, @('Write-TeamsCsvAtomically', 'Export-InventoryCsv')), @($spo, @('Export-SpoEntityCsv')))) {
    $ast = [Management.Automation.Language.Parser]::ParseFile($source[0], [ref]$null, [ref]$null)
    foreach ($name in $source[1]) {
        $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $definition) { throw "$name is missing." }
        . ([scriptblock]::Create($definition.Extent.Text))
    }
}
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-TeamsSpoStaged-' + [guid]::NewGuid().ToString('N'))
$latest = Join-Path $temporary 'DATA-LAST'
$output = Join-Path $temporary 'DATA-ALL/M365'
New-Item -ItemType Directory -Path $latest, $output | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey', 'SmartM365OrganizationKey', 'SmartM365EnvironmentKey', 'SmartM365TenantId', 'csvGeneratedPaths', 'RetentionMaxCSV')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'; $global:RetentionMaxCSV = 0
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
function Get-Hash([string]$Path) { if (Test-Path -LiteralPath $Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash } else { '' } }
function Get-LatestHashTable([string[]]$Files) { $result = @{}; foreach ($file in $Files) { $result[$file] = Get-Hash (Join-Path $latest $file) }; return $result }

# Producer context of the extracted functions; uploads made by the producer itself are recorded here.
$DryRun = $false
$AppendHistory = $false
$GeneratedCsvPaths = New-Object 'Collections.Generic.List[string]'
$script:DirectUploads = New-Object 'Collections.Generic.List[string]'
function Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:DirectUploads.Add($LocalFilePath); [pscustomobject]@{ LocalFilePath = $LocalFilePath } }
function Remove-SmartM365SharePointTimestampedCsvOlderThan { param($TimestampedPath, $RetentionDays) $null = $TimestampedPath, $RetentionDays }
function Assert-SmartM365CsvDataCompleteness { param($Data, $Columns, $TimestampedPath, $LatestPath) $null = $Data, $Columns, $TimestampedPath, $LatestPath }
function Add-SmartM365TenantKey { process { $_ } }
function WriteLog { param($Message, $Level) $null = $Message, $Level }
function Write-SpoLog { param($Message, $Level) $null = $Message, $Level }
function Add-SpoHistoryCsv { param($Data, $Path, $Columns) $null = $Data, $Path, $Columns }

# Uploads and WeeklyHistory made by Core: the publication helper and the deferred promotion actions.
& $core {
    $script:Recorded = New-Object 'Collections.Generic.List[object]'
    function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add([pscustomobject]@{ Kind = 'Upload'; Path = $LocalFilePath; SHA256 = (Get-FileHash -LiteralPath $LocalFilePath -Algorithm SHA256).Hash }) }
    function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param([string[]]$SourceFiles, [string]$TimestampedPath) $null = $TimestampedPath; $script:Recorded.Add([pscustomobject]@{ Kind = 'History'; Path = $SourceFiles[0]; SHA256 = (Get-FileHash -LiteralPath $SourceFiles[0] -Algorithm SHA256).Hash }) }
    function script:Remove-SmartM365SharePointTimestampedCsvOlderThan { param($TimestampedPath, $RetentionDays) $null = $TimestampedPath, $RetentionDays }
    function script:Assert-SmartM365CsvDataCompleteness { param($Data, $Columns, $TimestampedPath, $LatestPath) $null = $Data, $Columns, $TimestampedPath, $LatestPath }
    function script:WriteLog { param($Message, $Level) $null = $Message, $Level }
}
function Get-Recorded { & $core { $script:Recorded.ToArray() } }
function Clear-Recording { & $core { $script:Recorded.Clear() }; $script:DirectUploads.Clear(); $GeneratedCsvPaths.Clear() }
function Open-Run([string]$Producer, [switch]$Staged) {
    $global:csvGeneratedPaths = $null
    Start-SmartM365CmdbSourceReceipt -ScriptPath $Producer -SourceRootPath $latest
    if ($Staged) { Enable-SmartM365StagedSourcePublication }
}
function Close-Run([string]$Scope, [switch]$Failed) {
    if ($Failed) { Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null; return }
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $Scope
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
}
$stamp = 0
function Export-TeamsSet([string]$Label, [int]$Count) {
    $script:stamp++
    foreach ($file in @('M365_Teams_Teams', 'M365_Teams_Members', 'M365_Teams_Channels')) {
        $rows = @(1..$Count | ForEach-Object { [pscustomobject]@{ TenantKey = 'synthetic'; OrganizationKey = 'contoso'; EnvironmentKey = 'test'; TenantId = 'synthetic-tenant'; TeamId = "$Label-$_"; DisplayName = "$Label team $_" } })
        Export-InventoryCsv -Rows $rows -Columns @('TeamId', 'DisplayName') -TimestampedPath (Join-Path $output ('{0}_2026102{1}_010203.csv' -f $file, $script:stamp)) -LatestPath (Join-Path $latest "$file.csv") -HistoryPath ''
    }
}
function Export-SpoSet([string]$Label, [int]$Count) {
    $script:stamp++
    $rows = @(if ($Count -gt 0) { 1..$Count | ForEach-Object { [pscustomobject]@{ SiteId = "$Label-$_"; Url = "https://contoso.test/sites/$Label$_" } } })
    Export-SpoEntityCsv -BaseFileName 'M365_SPO_Sites' -Data $rows -Columns @('SiteId', 'Url') -TimestampedFolder $output -LatestFolder $latest -Timestamp ('2026102{0}_010203' -f $script:stamp)
}
try {
    Initialize-SmartM365DefaultCsvValidationRules
    # Static wiring of both producers: staging right after the receipt, staging Core, one version,
    # and WeeklyHistory fed with this run's copies, never the canonical DATA-LAST paths.
    $teamsText = [IO.File]::ReadAllText($teams)
    Check ($teamsText -match "(?m)^\s*Start-SmartM365CmdbSourceReceipt [^\r\n]*-ReadOnly:\`$DryRun\r?\n(\s*#[^\r\n]*\r?\n)*\s*Enable-SmartM365StagedSourcePublication") 'Teams does not enable staging right after its receipt.'
    Check ($teamsText -notmatch 'Write-TeamsCsvAtomically [^\r\n]*-Path \$LatestPath') 'Teams writes DATA-LAST outside the shared atomic copy.'
    Check ($teamsText -match 'Add-SmartM365WeeklyHistory -SourceCsvPaths \$GeneratedCsvPaths') 'Teams WeeklyHistory no longer reads the run copies list.'
    $spoText = [IO.File]::ReadAllText($spo)
    Check ($spoText -match "(?m)^\s*Start-SmartM365CmdbSourceReceipt [^\r\n]*-ReadOnly:\`$DryRun\r?\n(\s*#[^\r\n]*\r?\n)*\s*Enable-SmartM365StagedSourcePublication") 'SPO does not enable staging right after its receipt.'
    Check ($spoText -notmatch 'weeklyHistorySourcePaths\.Add\(\$exportResult\.LatestPath\)') 'SPO WeeklyHistory reads the previous run DATA-LAST file.'
    foreach ($text in @($teamsText, $spoText)) {
        $version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
        Check ($version -and $text -match ('\$ScriptVersion ?= ?"' + [regex]::Escape($version) + '"')) 'Header and runtime versions differ.'
        Check ($text -match "-MinimumVersion '1\.0\.(?:9[5-9]|[1-9]\d\d)'") 'A producer does not require the staging Core.'
    }

    # ---- Teams ----
    # Channels is outside the CMDB declaration and follows the same staging and promotion.
    $teamsFiles = @('M365_Teams_Teams.csv', 'M365_Teams_Members.csv', 'M365_Teams_Channels.csv')
    $receiptPath = Join-Path $latest 'SmartInventory_SmartM365-Teams-Inventory.current.json.txt'
    Open-Run $teams
    Export-TeamsSet 'yesterday' 2
    Close-Run 'CMDB:teams,team_members'
    $yesterdayReceipt = Read-Json $receiptPath
    $yesterday = Get-LatestHashTable $teamsFiles
    Check ($yesterdayReceipt.Status -eq 'Completed' -and @($yesterdayReceipt.Files | Where-Object File -in @('M365_Teams_Teams.csv', 'M365_Teams_Members.csv')).Count -eq 2) 'The previous Teams run did not publish.'

    # Running run with different data: DATA-LAST, receipt, uploads and the WeeklyHistory list keep
    # away from yesterday's files; the run copies hold this run's bytes.
    Clear-Recording
    Open-Run $teams -Staged
    Export-TeamsSet 'today' 4
    $during = Get-LatestHashTable $teamsFiles
    Check (@($teamsFiles | Where-Object { $during[$_] -ne $yesterday[$_] }).Count -eq 0) 'A running Teams run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) 'A running Teams run changed the receipt.'
    Check (@($script:DirectUploads | Where-Object { $_ -like "$latest*" }).Count -eq 0) 'A running Teams run uploaded the previous DATA-LAST file.'
    $historyInputs = @($GeneratedCsvPaths)
    Check (@($historyInputs | Where-Object { [IO.Path]::GetDirectoryName($_) -eq $latest }).Count -eq 0) 'Teams WeeklyHistory would read the previous DATA-LAST file.'
    $runCopies = @($historyInputs | Where-Object { $_ -like '*.staging*' })
    Check ($runCopies.Count -eq 3) 'Teams run copies are not staged.'
    foreach ($copy in $runCopies) {
        $timestamped = @($historyInputs | Where-Object { $_ -like "$output*" -and [IO.Path]::GetFileName($_) -like ([IO.Path]::GetFileNameWithoutExtension($copy) + '_*') })
        Check ($timestamped.Count -eq 1 -and (Get-Hash $copy) -eq (Get-Hash $timestamped[0])) "Teams run copy differs from this run export: $copy"
    }
    Check (@(Get-Recorded).Count -eq 0) 'A staged Teams copy was uploaded before promotion.'

    # Failed run: DATA-LAST and receipt unchanged, staging removed.
    Close-Run '' -Failed
    $after = Get-LatestHashTable $teamsFiles
    Check (@($teamsFiles | Where-Object { $after[$_] -ne $yesterday[$_] }).Count -eq 0) 'A failed Teams run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) 'A failed Teams run replaced the receipt.'
    Check (@(Get-ChildItem -LiteralPath (Join-Path $latest '.staging/SmartM365-Teams-Inventory') -Force).Count -eq 0) 'A failed Teams run left its staging folder.'

    # Complete run: promotion with the receipt, then canonical uploads of the promoted bytes.
    Clear-Recording
    Open-Run $teams -Staged
    Export-TeamsSet 'today' 4
    Close-Run 'CMDB:teams,team_members'
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable $teamsFiles
    Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $yesterdayReceipt.RunId) 'A complete Teams run did not publish its receipt.'
    foreach ($file in $teamsFiles) {
        $entry = @($receipt.Files | Where-Object File -eq $file)
        Check ($final[$file] -ne $yesterday[$file] -and ($entry.Count -eq 0 -or $entry[0].SHA256 -eq $final[$file])) "Promoted Teams CSV and receipt disagree: $file"
    }
    foreach ($file in @('M365_Teams_Teams.csv', 'M365_Teams_Members.csv')) {
        Check (@($receipt.Files | Where-Object File -eq $file).Count -eq 1) "Teams receipt misses $file"
    }
    $uploads = @(Get-Recorded | Where-Object Kind -eq 'Upload')
    Check ($uploads.Count -eq 3 -and @($uploads | Where-Object { $_.Path -like '*.staging*' -or $final[[IO.Path]::GetFileName($_.Path)] -ne $_.SHA256 }).Count -eq 0) 'Promoted Teams CSVs were not uploaded by canonical path.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
    $script:checks++

    # ---- SharePoint Online ----
    $spoFiles = @('M365_SPO_Sites.csv')
    $receiptPath = Join-Path $latest 'SmartInventory_SmartM365-SPO-Inventory.current.json.txt'
    Open-Run $spo
    $null = Export-SpoSet 'yesterday' 2
    Close-Run 'CMDB:sites'
    $yesterdayReceipt = Read-Json $receiptPath
    $yesterday = Get-LatestHashTable $spoFiles
    Check ($yesterdayReceipt.Status -eq 'Completed') 'The previous SPO run did not publish.'

    # Running run, populated then empty export: the run copy is staged and holds this run's bytes.
    foreach ($count in @(3, 0)) {
        Clear-Recording
        Open-Run $spo -Staged
        $result = Export-SpoSet 'today' $count
        $during = Get-LatestHashTable $spoFiles
        Check ($during['M365_SPO_Sites.csv'] -eq $yesterday['M365_SPO_Sites.csv']) "A running SPO run changed DATA-LAST (rows=$count)."
        Check ((Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) "A running SPO run changed the receipt (rows=$count)."
        Check ($result.PSObject.Properties['PublishedPath'] -and $result.PublishedPath -like '*.staging*' -and (Get-Hash $result.PublishedPath) -eq (Get-Hash $result.TimestampedPath)) "The SPO run copy for WeeklyHistory is not this run's staged file (rows=$count)."
        Check (@($script:DirectUploads | Where-Object { $_ -like "$latest*" }).Count -eq 0) "A running SPO run uploaded the previous DATA-LAST file (rows=$count)."
        Check (@(Get-Recorded | Where-Object { $_.Path -like "$latest*" }).Count -eq 0) "A staged SPO copy was uploaded or historized before promotion (rows=$count)."
        Close-Run '' -Failed
        Check ((Get-LatestHashTable $spoFiles)['M365_SPO_Sites.csv'] -eq $yesterday['M365_SPO_Sites.csv'] -and (Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) "A failed SPO run changed DATA-LAST (rows=$count)."
        Check (@(Get-ChildItem -LiteralPath (Join-Path $latest '.staging/SmartM365-SPO-Inventory') -Force).Count -eq 0) "A failed SPO run left its staging folder (rows=$count)."
    }

    # Complete run: promotion with the receipt, then canonical upload and WeeklyHistory of the promoted bytes.
    Clear-Recording
    Open-Run $spo -Staged
    $null = Export-SpoSet 'today' 3
    Close-Run 'CMDB:sites'
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable $spoFiles
    $entry = @($receipt.Files | Where-Object File -eq 'M365_SPO_Sites.csv')
    Check ($receipt.RunId -ne $yesterdayReceipt.RunId -and $final['M365_SPO_Sites.csv'] -ne $yesterday['M365_SPO_Sites.csv'] -and $entry.Count -eq 1 -and $entry[0].SHA256 -eq $final['M365_SPO_Sites.csv']) 'Promoted SPO CSV and receipt disagree.'
    $deferred = @(Get-Recorded | Where-Object { $_.Path -like "$latest*" })
    Check (@($deferred | Where-Object Kind -eq 'Upload').Count -eq 1 -and @($deferred | Where-Object Kind -eq 'History').Count -eq 1 -and @($deferred | Where-Object { $_.Path -like '*.staging*' -or $_.SHA256 -ne $final['M365_SPO_Sites.csv'] }).Count -eq 0) 'Promoted SPO CSV was not uploaded and historized by canonical path.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
    $script:checks++
}
finally {
    try { Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch { Write-Verbose $_.Exception.Message }
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks Teams and SPO staged publication checks; PowerShell $($PSVersionTable.PSVersion)."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCQyKh77IyCHuN7
# ByINis0ll+1XttiYKm2Dip1vXXIX3aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHOU+cL6vBWotmL5GHfD7c3LaO+Lgozz9SrreC5OVn1xMA0GCSqG
# SIb3DQEBAQUABIIBgDckfJardurn3Q+ONFiaoZrKRBuMYp8ZQ1kCHpV8moxYuAQB
# H7FRz/MJtjXFtAxykBO0rpWoNuTfE7TZd06/rsAA4SsFghqxHwFbqbWyur1sLiJ0
# laPEbrecso+IpBzGcVvtr3DEY+SKjfVjr3P7ZyfED8R8suO3UV1+Fa2Mm0Yg7awj
# VY1Qrd19xyFBMgjBhG9QHF/ogZVkXKTX+tZSWhdaJkHJ6Vs93LEn0MTiPrMot5fH
# 3sM8CpirhBG/59XnVgiBBL22IXVIw4BJ8rxOyi5QQyF2w9vyndVZpqEaDhc1QS33
# fxCKevO2++QzWE4z1JMfBEhQWDYWy5wliuiQ5gJEjGKC33kiM+hd1H70jq7Cnxhl
# 6Te+Ec5Et7WTczl3Dio1YNCsod8BjOP3++5eKwO5XfdIx2WQPPWrGhh6S1DfQeAD
# 2wbPsyoqBpQmr542/CH6HEhTuCME6aAukNRVl/0B6v4nO4PjPWX71qnc4AUwZ4Bl
# PG0owhDR8nAdJ18eHaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjEwNzA4
# MzdaMC8GCSqGSIb3DQEJBDEiBCCM3oIWQYUXGLwKorjcKImKugQUCB88PllCXgGA
# sV6p7DANBgkqhkiG9w0BAQEFAASCAgBpjYV8xqjIwaeRUPkmEdwrKcatUzwMT1vX
# aWBNZQmQNVdyxDWIMEAnipc1smfcoik46H0JBP7oqY07cFkv9SEOpHeTPZNUCufL
# DzRa/6d6UV5ur/Z29N8EWFngKBM/kcy9f7k+cyeYJZo+ZsLx4V3Y1KPPbHq4jgmO
# 20wwRKFhv9b6qZRGJPsIuUcGAoV7vosVntAJuiSkzVK0vbAmzIAqpJCpEoQg8aQB
# Xjj5AcPcxKg1TXfyGYAxX+l5xc3pHCCIAM0wakQShkgGSeV+FzGdNif+xngigXZV
# ody6kERmiftmfJP7MT9Tl4mC4lQX4ZSu0ahvLHPPHbA/768NIsyca+IakgxPTWsI
# xumsHRbN0x2O0FcotleZYp19fUQI0FH9iEFSD/enzQszXm5I0pREidPphz/dGiej
# ey8QwhfnXXy5IpR3sDf8n2RqW3tLGkR0gTlDMyYnyaj5kPv5S+IU6WZ0LG7+sMis
# +7Id8HbM9pkiG2NlfBX2kUBqY9GZzpBzRRAblsOynHSU5nUxQycoTdw39tEGwjvB
# uFRalwj6eeLyaUxGyGOYwph4kVSwlYHcNwa17BdfKGwcP/p1hbu7/YMpUmAu0QGy
# tVXaJNQ/31rv5curTUN1CogZhaSSGb+WxH84YPA62iJDL7kZLYvACSBP7hjMSDKw
# 38ObZpNwyg==
# SIG # End signature block
