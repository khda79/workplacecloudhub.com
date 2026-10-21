#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged DATA-LAST publication for the Exchange hybrid identity issues: with the
previous run's different CSVs in DATA-LAST, a running or failed run neither changes nor uploads them,
and a complete run promotes this run's CSVs with the receipt, receipt last.
.VERSION
1.0.0
.NOTES
Synthetic TEMP data only. SharePoint upload and WeeklyHistory are replaced by recorders.
No directory, Exchange, tenant, API, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','',Justification='Script variables are read by the extracted collector functions.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.95' -Force -PassThru
$producer = Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter 'SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1' | Select-Object -First 1
$ast = [Management.Automation.Language.Parser]::ParseFile($producer.FullName, [ref]$null, [ref]$null)
foreach ($name in @('FlattenPath', 'CopyCsv', 'PublishLatest', 'Remove-DetailedArchiveFilesOlderThan', 'ExportIssues')) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $definition) { throw "$name is missing." }
    . ([scriptblock]::Create($definition.Extent.Text))
}
function Log([string]$m) { $null = $m }
function Warn([string]$m) { $null = $m }
$files = @('Exchange_HybridIdentity_Issues.csv', 'Exchange_HybridIdentity_Issues_Summary.csv')
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-HybridStaged-' + [guid]::NewGuid().ToString('N'))
$LatestFolder = Join-Path $temporary 'DATA-LAST'
$OutputFolder = Join-Path $temporary 'DATA-ALL/Exchange/Issues/HybridIdentity'
$DetailedArchiveRetentionDays = 7
New-Item -ItemType Directory -Path $LatestFolder, $OutputFolder | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey', 'SmartM365OrganizationKey', 'SmartM365EnvironmentKey', 'SmartM365TenantId', 'csvGeneratedPaths')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$receiptPath = Join-Path $LatestFolder 'SmartInventory_SmartM365-Exchange-HybridIdentity-Issues-Inventory.current.json.txt'
$runPath = Join-Path $LatestFolder 'SmartInventory_SmartM365-Exchange-HybridIdentity-Issues-Inventory.run.json.txt'
$stagingProducer = Join-Path $LatestFolder '.staging/SmartM365-Exchange-HybridIdentity-Issues-Inventory'
function Get-LatestHashTable { $result = @{}; foreach ($file in $files) { $path = Join-Path $LatestFolder $file; $result[$file] = if (Test-Path -LiteralPath $path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } else { '' } }; return $result }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
function Get-IssueRowSet([string]$Label, [int]$Count) {
    @(1..$Count | ForEach-Object { [pscustomobject]@{ TenantKey = 'synthetic'; IssueNumber = $_; ObjectGUID = [guid]::NewGuid().ToString(); Potential_Issue = "$Label issue $_"; IssueCategory = '3.Medium'; RecommendedAction = 'Review' } })
}
function Open-HybridRun([switch]$Staged) {
    $global:csvGeneratedPaths = $null
    $script:RunStamp = (Get-Date).ToString('yyyyMMdd-HHmmss') + [guid]::NewGuid().ToString('N').Substring(0, 4)
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $LatestFolder
    if ($Staged) { Enable-SmartM365StagedSourcePublication }
}
try {
    Initialize-SmartM365DefaultCsvValidationRules
    # Static wiring: staging right after the receipt, inputs read from other producers only, one version.
    $text = [IO.File]::ReadAllText($producer.FullName)
    Check ($text -match "(?m)^\s*Start-SmartM365SourceReceipt -ScriptPath \`$PSCommandPath -SourceRootPath \`$LatestFolder\r?\n(\s*#[^\r\n]*\r?\n)*\s*Enable-SmartM365StagedSourcePublication") 'The collector does not enable staging right after its receipt.'
    $inputs = @([regex]::Matches($text, "Csv '([^']+)'") | ForEach-Object { $_.Groups[1].Value })
    Check ($inputs.Count -ge 10 -and @($inputs | Where-Object { $_ -in $files }).Count -eq 0) 'The collector reads its own DATA-LAST output during the run.'
    Check (@($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'CopyCsv' -and $n.Extent.Text -match '\$(latest|sl)\b' }, $true)).Count -eq 0) 'A DATA-LAST copy bypasses the staged helper.'
    $version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
    Check ($version -and $text -match ('\$ScriptVersion="' + [regex]::Escape($version) + '"')) 'Header and runtime versions differ.'
    Check ($text -match "-MinimumVersion '1\.0\.(?:9[5-9]|[1-9]\d\d)'") 'The collector does not require the staging Core.'

    & $core {
        $script:Recorded = New-Object 'Collections.Generic.List[string]'
        function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add($LocalFilePath) }
        function script:WriteLog { param($Message, $Level) $null = $Message, $Level }
    }
    function Get-DeferredUploadList { & $core { $script:Recorded.ToArray() } }

    # Previous run, published directly: DATA-LAST holds yesterday's issues and receipt.
    Open-HybridRun
    $null = ExportIssues -rows (Get-IssueRowSet 'yesterday' 3)
    Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null
    $yesterdayReceipt = Read-Json $receiptPath
    $yesterday = Get-LatestHashTable
    Check ($yesterdayReceipt.Status -eq 'Completed' -and @($yesterdayReceipt.Files).Count -eq 2) 'The previous hybrid identity run did not publish.'

    # Running run with different issues: DATA-LAST and the receipt keep yesterday's version, and the
    # upload list of the run never names a DATA-LAST file (it would upload yesterday's copy).
    Open-HybridRun -Staged
    $runFiles = @(ExportIssues -rows (Get-IssueRowSet 'today' 5))
    $during = Get-LatestHashTable
    Check (@($files | Where-Object { $during[$_] -ne $yesterday[$_] }).Count -eq 0) 'A running hybrid identity run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) 'A running hybrid identity run changed the receipt.'
    Check (@($runFiles | Where-Object { $_ -like "$LatestFolder*" }).Count -eq 0) 'The run uploads DATA-LAST files before their promotion.'
    Check (@($runFiles | Where-Object { $_ -like "$OutputFolder*" }).Count -eq 3) 'The run lost its DATA-ALL outputs.'
    Check (@(Get-DeferredUploadList).Count -eq 0) 'A staged copy was uploaded before promotion.'

    # Failed run: nothing changes, staging is removed.
    Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $yesterday[$_] }).Count -eq 0) 'A failed hybrid identity run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $yesterdayReceipt.RunId) 'A failed hybrid identity run replaced the receipt.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A failed hybrid identity run left its staging folder.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath)
    $script:checks++

    # Complete run: this run's CSVs replace yesterday's together with the receipt; uploads follow.
    Open-HybridRun -Staged
    $null = ExportIssues -rows (Get-IssueRowSet 'today' 5)
    Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable
    Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $yesterdayReceipt.RunId) 'A complete hybrid identity run did not publish its receipt.'
    foreach ($file in $files) {
        $entry = @($receipt.Files | Where-Object File -eq $file)
        Check ($final[$file] -ne $yesterday[$file] -and $entry.Count -eq 1 -and $entry[0].SHA256 -eq $final[$file]) "Promoted hybrid identity CSV and receipt disagree: $file"
    }
    Check (@(Import-Csv -LiteralPath (Join-Path $LatestFolder 'Exchange_HybridIdentity_Issues.csv') | Where-Object { $_.Potential_Issue -like 'today*' }).Count -eq 5) 'DATA-LAST does not hold this run issues.'
    $deferred = @(Get-DeferredUploadList)
    Check ($deferred.Count -eq 2 -and @($deferred | Where-Object { $_ -like '*.staging*' -or $_ -notlike "$LatestFolder*" }).Count -eq 0) 'Promoted copies were not uploaded by canonical DATA-LAST path.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication -and $run.StagedPublication) 'Completed hybrid identity run state is not qualified.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
    $script:checks++
}
finally {
    try { Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch { Write-Verbose $_.Exception.Message }
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks hybrid identity staged publication checks; PowerShell $($PSVersionTable.PSVersion)."
