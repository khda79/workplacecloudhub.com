#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged DATA-LAST publication for the EXO mailboxes producer (full, fast and
permissions-only jobs share one script).
.VERSION
1.0.0
.NOTES
Synthetic TEMP data only. SharePoint upload and WeeklyHistory are replaced by recorders.
No tenant, Exchange, API, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.93' -Force -PassThru
$producer = Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter 'SmartM365-EXO-Mailboxes-Inventory.ps1' | Select-Object -First 1
$files = @('Exchange_EXO_Mailboxes_AllDomains.csv', 'Exchange_EXO_Mailboxes_AllDomains_Archive.csv', 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv', 'Exchange_EXO_Mailboxes_AllDomains_Permissions.csv')
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-ExoStaged-' + [guid]::NewGuid().ToString('N'))
$latest = Join-Path $temporary 'DATA-LAST'
$output = Join-Path $temporary 'DATA-ALL/Exchange/Mailboxes'
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
$receiptPath = Join-Path $latest 'SmartInventory_SmartM365-EXO-Mailboxes-Inventory.current.json.txt'
$runPath = Join-Path $latest 'SmartInventory_SmartM365-EXO-Mailboxes-Inventory.run.json.txt'
$stagingProducer = Join-Path $latest '.staging/SmartM365-EXO-Mailboxes-Inventory'
function Get-LatestHashTable { $result = @{}; foreach ($file in $files) { $path = Join-Path $latest $file; $result[$file] = if (Test-Path -LiteralPath $path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } else { '' } }; return $result }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
# Same calls as the producer export stage: one ExportAndCopyCsv per mailbox CSV.
function Export-MailboxSet([string]$Label) {
    foreach ($file in $files) {
        $rows = @([pscustomobject]@{ UserPrincipalName = "$Label@contoso.test"; PrimarySmtpAddress = "$Label.$file@contoso.test" })
        ExportAndCopyCsv -BaseFileName ([IO.Path]::GetFileNameWithoutExtension($file)) -OutputPath $output -GlobalPath $latest -Data $rows -Encoding UTF8 -NoTypeInformation
        Start-Sleep -Milliseconds 1100
    }
}
function Open-MailboxRun([switch]$Staged) {
    $global:csvGeneratedPaths = $null
    Start-SmartM365CmdbSourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    if ($Staged) { Enable-SmartM365StagedSourcePublication }
}
try {
    # Static wiring: staging follows the receipt start, diagnostic selections and permissions-only
    # runs keep their behavior, and every DATA-LAST snapshot is read before the first export.
    $text = [IO.File]::ReadAllText($producer.FullName)
    $ast = [Management.Automation.Language.Parser]::ParseFile($producer.FullName, [ref]$null, [ref]$null)
    Check ($text -match "(?m)^\s*Start-SmartM365CmdbSourceReceipt [^\r\n]*-ReadOnly:\`$PermissionsOnly\r?\n(\s*#[^\r\n]*\r?\n)*\s*if \(-not \`$Top100 -and \[string\]::IsNullOrWhiteSpace\(\`$DebugUPN\)\) \{ Enable-SmartM365StagedSourcePublication \}") 'EXO mailboxes does not enable staging right after its receipt.'
    Check ($text -match "-MinimumVersion '1\.0\.(?:9[3-9]|[1-9]\d\d)'") 'EXO mailboxes does not require the staging Core.'
    $calls = @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] }, $true))
    $firstExport = @($calls | Where-Object { $_.GetCommandName() -eq 'ExportAndCopyCsv' } | ForEach-Object { $_.Extent.StartLineNumber } | Sort-Object)[0]
    $snapshotLoads = @($calls | Where-Object { $_.GetCommandName() -in @('Load-M365UsersCache', 'Load-AdUsersCache', 'Load-PermissionsSnapshot', 'Load-StatsSnapshot') } | ForEach-Object { $_.Extent.StartLineNumber })
    Check ($snapshotLoads.Count -ge 4 -and @($snapshotLoads | Where-Object { $_ -gt $firstExport }).Count -eq 0) 'EXO mailboxes reads a DATA-LAST snapshot after its first export.'
    $version = [regex]::Match($text, '\.VERSION\r?\n([0-9.]+)').Groups[1].Value
    Check ($version -and $text -match ('\$ScriptVersion = "' + [regex]::Escape($version) + '"')) 'Header and runtime versions differ.'

    & $core {
        $script:Recorded = New-Object 'Collections.Generic.List[object]'
        function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add([pscustomobject]@{ Kind = 'Upload'; Path = $LocalFilePath }) }
        function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param([string[]]$SourceFiles, [string]$TimestampedPath) $null = $TimestampedPath; $script:Recorded.Add([pscustomobject]@{ Kind = 'History'; Path = $SourceFiles[0] }) }
        function script:WriteLog {
            param($Message, $Level)
            $null = $Level
            if ([string]$Message -match 'promoted from source staging') {
                $run = Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName(($Message -replace '^CSV latest copy written to: (.+) \(promoted.*$', '$1'))) -Filter '*.run.json.txt' | Select-Object -First 1
                $receipt = Get-ChildItem -LiteralPath $run.DirectoryName -Filter '*.current.json.txt' | Select-Object -First 1
                $script:Recorded.Add([pscustomobject]@{ Kind = 'Promoted'; Path = $Message; RunStatus = (Get-Content -LiteralPath $run.FullName -Raw | ConvertFrom-Json).Status; ReceiptRunId = $(if ($receipt) { (Get-Content -LiteralPath $receipt.FullName -Raw | ConvertFrom-Json).RunId } else { '' }) })
            }
        }
    }
    function Get-Recorded { & $core { $script:Recorded.ToArray() } }
    function Clear-Recorded { & $core { $script:Recorded.Clear() } }

    # Baseline: the previous weekly run published directly.
    Open-MailboxRun
    Export-MailboxSet 'baseline'
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope 'CMDB:mailboxes'
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $baselineReceipt = Read-Json $receiptPath
    $baselineHashes = Get-LatestHashTable
    Check ($baselineReceipt.Status -eq 'Completed' -and @($baselineReceipt.Files).Count -eq 4) 'Baseline mailbox receipt not completed.'
    Clear-Recorded

    # Saturday morning: the weekly run is collecting and exporting; preparation reads DATA-LAST.
    Open-MailboxRun -Staged
    Export-MailboxSet 'running'
    $during = Get-LatestHashTable
    Check (@($files | Where-Object { $during[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A running mailbox run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A running mailbox run changed the receipt.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Collecting' -and $run.StagedPublication -eq $true -and -not $run.UnqualifiedPublication) 'A running mailbox run left the Collecting state.'
    Check (@(Get-Recorded | Where-Object { $_.Kind -eq 'Upload' -and $_.Path -like "$latest*" }).Count -eq 0) 'A staged mailbox CSV was uploaded before promotion.'

    # A run that stops after its exports (timeout, cancellation, error) leaves the previous set.
    Complete-SmartM365CmdbSourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A failed mailbox run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A failed mailbox run replaced the receipt.'
    Check (-not (Read-Json $runPath).UnqualifiedPublication) 'A failed mailbox run disqualified the previous set.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A failed mailbox run left its staging folder.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath)
    $script:checks++

    # A partial scope (MaxRunExceeded, manual stop) is never promoted.
    Open-MailboxRun -Staged
    Export-MailboxSet 'partial'
    Set-SmartM365CmdbSourceScope -CompleteScope $false -Scope 'CMDB:mailboxes'
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A partial mailbox run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A partial mailbox run replaced the receipt.'

    # Complete run: the four CSVs and the receipt are promoted together, receipt last.
    Clear-Recorded
    Open-MailboxRun -Staged
    Export-MailboxSet 'complete'
    Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope 'CMDB:mailboxes'
    Complete-SmartM365CmdbSourceReceipt -Status 'Success' | Out-Null
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable
    Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $baselineReceipt.RunId) 'A complete mailbox run did not publish its receipt.'
    foreach ($file in $files) {
        $entry = @($receipt.Files | Where-Object File -eq $file)
        Check ($final[$file] -ne $baselineHashes[$file] -and $entry.Count -eq 1 -and $entry[0].SHA256 -eq $final[$file]) "Promoted mailbox CSV and receipt disagree: $file"
    }
    $promoted = @(Get-Recorded | Where-Object Kind -eq 'Promoted')
    Check ($promoted.Count -eq 4 -and @($promoted | Where-Object { $_.RunStatus -ne 'Publishing' -or $_.ReceiptRunId -ne $baselineReceipt.RunId }).Count -eq 0) 'A mailbox CSV was promoted outside the fence or after the new receipt.'
    $uploads = @(Get-Recorded | Where-Object { $_.Kind -eq 'Upload' -and $_.Path -like "$latest*" })
    Check ($uploads.Count -eq 4 -and @($uploads | Where-Object { $_.Path -like '*.staging*' }).Count -eq 0) 'Promoted mailbox CSVs were not uploaded by canonical path.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication) 'Completed mailbox run state is not qualified.'
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
"PASS: $script:checks EXO mailbox staged publication checks; PowerShell $($PSVersionTable.PSVersion)."
