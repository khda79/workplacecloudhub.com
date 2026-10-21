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
