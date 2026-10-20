#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged DATA-LAST publication for the Active Directory inventory: the combined
CSVs, duplicate-identity CSVs and daily copies reach DATA-LAST with the receipt, never before.
.VERSION
1.0.0
.NOTES
Synthetic TEMP data only. SharePoint upload and WeeklyHistory are replaced by recorders.
No directory, tenant, API, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.93' -Force -PassThru
$producer = Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter 'SmartM365-ActiveDirectory-Inventory.ps1' | Select-Object -First 1
$ast = [Management.Automation.Language.Parser]::ParseFile($producer.FullName, [ref]$null, [ref]$null)
$definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Publish-SmartM365AdLatestCopy' }, $true)
if (-not $definition) { throw 'Publish-SmartM365AdLatestCopy is missing.' }
. ([scriptblock]::Create($definition.Extent.Text))
$required = @('AD_Users_AllDomains.csv', 'AD_Computers_AllDomains.csv', 'AD_Domains_AllDomains.csv', 'AD_Groups_AllDomains.csv', 'AD_DirectoryObjects_AllDomains.csv', 'AD_GroupMemberships_AllDomains.csv')
$extra = @('AD_Users_DuplicateUPN.csv', 'AD_Users_DailyReport.csv')
$files = @($required + $extra)
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-AdStaged-' + [guid]::NewGuid().ToString('N'))
$latest = Join-Path $temporary 'DATA-LAST'
$output = Join-Path $temporary 'DATA-ALL/ActiveDirectory/Inventory'
New-Item -ItemType Directory -Path $latest, $output | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey', 'SmartM365OrganizationKey', 'SmartM365EnvironmentKey', 'SmartM365TenantId', 'csvGeneratedPaths')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$script:Uploads = New-Object 'Collections.Generic.List[string]'
function Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Uploads.Add($LocalFilePath); return [pscustomobject]@{ LocalFilePath = $LocalFilePath } }
$receiptPath = Join-Path $latest 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt'
$runPath = Join-Path $latest 'SmartInventory_SmartM365-ActiveDirectory-Inventory.run.json.txt'
$stagingProducer = Join-Path $latest '.staging/SmartM365-ActiveDirectory-Inventory'
$scope = 'CMDB:ad_users,ad_computers,ad_domains,ad_groups,ad_objects,ad_members'
function Get-LatestHashTable { $result = @{}; foreach ($file in $files) { $path = Join-Path $latest $file; $result[$file] = if (Test-Path -LiteralPath $path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } else { '' } }; return $result }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
# Same publication as the collector: combined and report CSVs in DATA-ALL, then their DATA-LAST copies.
function Publish-AdSet([string]$Label) {
    foreach ($file in $files) {
        $source = Join-Path $output $file
        [IO.File]::WriteAllText($source, "TenantKey,DomainName,Name`r`nsynthetic,contoso.test,$Label-$file`r`n", [Text.UTF8Encoding]::new($false))
        Publish-SmartM365AdLatestCopy -SourcePath $source -DestinationPath (Join-Path $latest $file) | Out-Null
        [void]$global:csvGeneratedPaths.Add((Join-Path $latest $file))
    }
}
function Open-AdRun([switch]$Staged) {
    $global:csvGeneratedPaths = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    Start-SmartM365CmdbSourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    if ($Staged) { Enable-SmartM365StagedSourcePublication }
}
try {
    # Static wiring: staging follows the receipt, and every DATA-LAST copy uses the staged helper.
    $text = [IO.File]::ReadAllText($producer.FullName)
    Check ($text -match "(?m)^\s*Start-SmartM365CmdbSourceReceipt [^\r\n]*-ReadOnly:\(\`$DomainWorker -or \`$ReportOnly -or \`$DuplicateAnalysisOnly\)\r?\n(\s*#[^\r\n]*\r?\n)*\s*Enable-SmartM365StagedSourcePublication") 'AD does not enable staging right after its receipt.'
    $copies = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Copy-SmartM365FileAtomically' }, $true))
    # Atomic copies stay in the staged helper, the DATA-ALL combine retry and the DATA-ALL WeeklyHistory.
    $directLatest = @($copies | Where-Object {
        $parent = $_.Parent; while ($parent -and $parent -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $parent = $parent.Parent }
        -not $parent -or $parent.Name -notin @('Publish-SmartM365AdLatestCopy', 'Copy-SmartM365AdFileWithRetry', 'Save-WeeklyInventoryHistory')
    })
    Check ($directLatest.Count -eq 0) 'AD copies a CSV to DATA-LAST outside the staged helper.'
    Check (@($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Publish-SmartM365AdLatestCopy' }, $true)).Count -ge 4) 'AD does not publish its DATA-LAST copies through the staged helper.'
    $version = [regex]::Match($text, '\.VERSION\r?\n\s*([0-9.]+)').Groups[1].Value
    Check ($version -and $text -match ('\$ScriptVersion = "' + [regex]::Escape($version) + '"')) 'Header and runtime versions differ.'
    Check ($text -match "-MinimumVersion '1\.0\.(?:9[3-9]|[1-9]\d\d)'") 'AD does not require the staging Core.'

    & $core {
        $script:Recorded = New-Object 'Collections.Generic.List[object]'
        function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add($LocalFilePath) }
        function script:WriteLog { param($Message, $Level) $null = $Message, $Level }
    }
    function Get-DeferredUploadList { & $core { $script:Recorded.ToArray() } }

    # Baseline: the previous run published directly.
    Open-AdRun
    Publish-AdSet 'baseline'
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $scope
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $baselineReceipt = Read-Json $receiptPath
    $baselineHashes = Get-LatestHashTable
    Check ($baselineReceipt.Status -eq 'Completed' -and @($baselineReceipt.Files).Count -eq $files.Count) 'Baseline AD receipt not completed.'
    Check ($script:Uploads.Count -eq $files.Count) 'Direct AD copies were not uploaded immediately.'
    $script:Uploads.Clear()

    # Running run: DATA-LAST and the receipt are unchanged; no DATA-LAST upload yet.
    Open-AdRun -Staged
    Publish-AdSet 'running'
    $during = Get-LatestHashTable
    Check (@($files | Where-Object { $during[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A running AD run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A running AD run changed the receipt.'
    Check ($script:Uploads.Count -eq 0 -and @(Get-DeferredUploadList).Count -eq 0) 'A staged AD copy was uploaded before promotion.'
    Check ((Read-Json $runPath).Status -eq 'Collecting') 'A running AD run left the Collecting state.'

    # A failed run (for example a swallowed copy failure or a missing required file) changes nothing.
    Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A failed AD run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A failed AD run replaced the receipt.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A failed AD run left its staging folder.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath)
    $script:checks++

    # A run that lost one required copy is never promoted, even when declared complete.
    Open-AdRun -Staged
    Publish-AdSet 'missing'
    Remove-Item -LiteralPath (Join-Path (Get-ChildItem -LiteralPath $stagingProducer -Directory)[0].FullName 'AD_Groups_AllDomains.csv') -Force
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $scope
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'An AD run with a lost required copy changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'An AD run with a lost required copy published a receipt.'

    # Complete run: all copies and the receipt are promoted together, uploads follow the promotion.
    Open-AdRun -Staged
    Publish-AdSet 'complete'
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $scope
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable
    Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $baselineReceipt.RunId) 'A complete AD run did not publish its receipt.'
    foreach ($file in $files) {
        $entry = @($receipt.Files | Where-Object File -eq $file)
        Check ($final[$file] -ne $baselineHashes[$file] -and $entry.Count -eq 1 -and $entry[0].SHA256 -eq $final[$file]) "Promoted AD CSV and receipt disagree: $file"
    }
    $deferred = @(Get-DeferredUploadList)
    Check ($deferred.Count -eq $files.Count -and @($deferred | Where-Object { $_ -like '*.staging*' }).Count -eq 0) 'Promoted AD copies were not uploaded by canonical path.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication -and $run.StagedPublication) 'Completed AD run state is not qualified.'
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
"PASS: $script:checks AD staged publication checks; PowerShell $($PSVersionTable.PSVersion)."
