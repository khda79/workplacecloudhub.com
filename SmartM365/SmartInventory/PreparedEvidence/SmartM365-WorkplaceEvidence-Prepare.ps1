<#
.SYNOPSIS
    Prepare a validated, versioned SmartWorkplaceIntelligence batch after raw collectors.
.VERSION
    0.4.12
.NOTES
    PowerShell 7. SharePoint mapping reads use the tenant configuration unless Offline.
    Deployment must include the sibling SmartWorkplaceIntelligence/scripts, config and reference folders.
#>
[CmdletBinding()]
param([string]$Tenant='test', [switch]$ValidateOnly, [switch]$Offline, [switch]$WorkforceDiagnostic,
    [switch]$RepairLegacyHistory, [string]$RepairWeeks, [int]$ExpectedRepairFileCount=0, [switch]$ApplyRepair,
    [switch]$ConvertMetadata, [switch]$ApplyMetadataConversion, [switch]$TransferOnly, [string]$ExpectedBatchId)
# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
$ErrorActionPreference='Stop'
if ($RepairLegacyHistory -or $ConvertMetadata) { $Offline=$true }
$tenantContext = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'Config/SmartM365-TenantContext.ps1'
. $tenantContext
$effective = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$smartRoot = Find-SmartM365Root -StartPath $PSScriptRoot
$scriptName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$failure = $null
$coreLoaded = $false
$phase = 'Configuration'
$output = ''
$mappingRoot = ''
$batchUploadEnabled = $false
$transfer = $null
$transcriptStarted = $false
function Invoke-PreparedStep {
    param([Parameter(Mandatory)][scriptblock]$Action)
    # Native child processes flatten warning records into console text. Restore their
    # severity before the shared completion summary, keeping success output unchanged.
    & $Action 3>&1 6>&1 | ForEach-Object {
        if ($_ -is [Management.Automation.WarningRecord]) {
            WriteLog -Message $_.Message -Level WARNING
        } elseif ($_ -is [Management.Automation.InformationRecord]) {
            foreach ($line in ([string]$_.MessageData -split '\r?\n')) {
                # Child hosts attached to a VT console colour warnings; match the text without ANSI sequences.
                $plain = $line -replace "$([char]27)\[[0-?]*[ -/]*[@-~]", ''
                if ($plain -match '^(?:\[\d{4}-\d{2}-\d{2}[^\]]*\]\s*)*(?:WARNING|AVERTISSEMENT)\s*:\s*(.*)$') {
                    WriteLog -Message $Matches[1] -Level WARNING
                } elseif ($_.Tags -contains 'PSHOST') {
                    Write-Host $line
                } else {
                    Write-Information -MessageData $line -Tags $_.Tags -InformationAction Continue
                }
            }
        } else { $_ }
    }
}
function Register-PreparedRunCsvFiles {
    param([Parameter(Mandatory)]$PreparedResult)
    # The validation receipt lists prepared tables only, excluding Source/History artifacts.
    $receipt=Get-Content -LiteralPath (Join-Path $PreparedResult.BatchPath 'validation.json.txt') -Raw | ConvertFrom-Json
    $names=@($receipt.Files.File)
    if ($receipt.Passed -ne $true -or $receipt.SchemaOnly -ne $false -or $names.Count -lt 1 -or
        $names.Count -ne $PreparedResult.Files -or @($names | Sort-Object -Unique).Count -ne $names.Count -or
        @($names | Where-Object {$_ -cnotmatch '^[A-Za-z0-9_-]+\.csv$'}).Count) {
        throw 'Prepared CSV counter requires the matching successful validation receipt.'
    }
    if (-not $global:csvGeneratedPaths) {
        $global:csvGeneratedPaths=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    }
    foreach ($name in $names) { [void]$global:csvGeneratedPaths.Add((Join-Path $PreparedResult.BatchPath $name)) }
}
function ConvertTo-PreparedAgeOverrides {
    param([AllowNull()]$InputObject)
    $result = @{}
    if ($null -eq $InputObject) { return $result }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $keys = @($InputObject.Keys)
    } elseif ($InputObject -is [pscustomobject]) {
        $keys = @(foreach ($property in $InputObject.PSObject.Properties) { $property.Name })
    } else { throw 'PreparedSourceAgeOverrides must be a JSON object keyed by CSV filename.' }
    foreach ($key in $keys) {
        $value = if ($InputObject -is [System.Collections.IDictionary]) { $InputObject[$key] } else { $InputObject.PSObject.Properties[$key].Value }
        $hours = 0
        if ([string]::IsNullOrWhiteSpace([string]$key) -or -not [int]::TryParse([string]$value, [ref]$hours) -or $hours -le 0) {
            throw 'PreparedSourceAgeOverrides entries require a non-empty CSV filename and positive integer hours.'
        }
        $result[[string]$key] = $hours
    }
    return $result
}
function Invoke-PostTransferRetention($VerifiedTransfer) {
    # Retention failure must never enter the preparation/publication failure path.
    try {
        $retention = Invoke-PreparedStep {
            & (Join-Path $product 'scripts/Invoke-PreparedSharePointRetention.ps1') -Tenant $Tenant -Mode Configured `
                -ConfigurationPath (Get-SmartM365JsonReadPath (ConvertTo-SmartM365InventoryConfigLocation -Path (Join-Path $PSScriptRoot "$scriptName.local.json"))) `
                -ExpectedPublishedBatchId $VerifiedTransfer.BatchId -ExpectedPointerSHA256 $VerifiedTransfer.PointerSHA256 -NoCompletionBanner
        }
        $level=if($retention.Status -in @('Blocked','PartialOrBlocked')){'WARNING'}else{'INFO'}
        $reasons=if($retention.PSObject.Properties['Plan'] -and @($retention.Plan.BlockedReasons).Count){
            ' Reasons: '+(@($retention.Plan.BlockedReasons) -join '; ')
        }else{''}
        WriteLog -Message "Post-transfer retention: $($retention.Status); plan=$($retention.PlanPath); journal=$($retention.JournalPath).$reasons" -Level $level
    } catch {
        WriteLog -Message "Batch publication remains verified; retention incomplete: $($_.Exception.Message)" -Level WARNING
    }
}
function Send-PreparedSuccessMail {
    param([Parameter(Mandatory)][string]$BatchId,
        [Parameter(Mandatory)][ValidateRange(1,[int]::MaxValue)][int]$CsvFiles,
        [AllowNull()]$VerifiedTransfer)
    if ($Offline -or $ValidateOnly -or $WorkforceDiagnostic -or $RepairLegacyHistory -or $ConvertMetadata) { return }
    if ($batchUploadEnabled) {
        Assert-PreparedTransferResult -Result $VerifiedTransfer -ExpectedBatchId $BatchId
        if ($VerifiedTransfer.CsvFiles -ne $CsvFiles) { throw 'Success email requires consistent local and transferred CSV counts.' }
    }
    $ended=Get-Date
    $started=Get-Variable -Name SmartM365ExecutionStartTime -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    $logPath=Get-Variable -Name LogTextFile -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    $details=[ordered]@{
        'Tenant'=$effective.TenantKey
        'Batch ID'=$BatchId
        'Operation'=$(if($TransferOnly){'Existing-batch transfer (no CSV generation)'}else{'Preparation and validation'})
        'Prepared CSV files'=$CsvFiles
        'Completed at'=$ended.ToString('o')
        'Duration'=$(if($started){($ended-[datetime]$started).ToString('c')}else{'Not measured'})
        'Output root'=$output
        'SharePoint transfer'=$(if($batchUploadEnabled){'Batch files and current pointer read back and verified'}else{'Disabled; local batch only'})
        'Log'=$logPath
    }
    if ($batchUploadEnabled) { $details['Verified SharePoint files']=$VerifiedTransfer.VerifiedFiles }
    $rows=foreach($label in $details.Keys) {
        '<tr><th>{0}</th><td>{1}</td></tr>' -f [Net.WebUtility]::HtmlEncode($label),[Net.WebUtility]::HtmlEncode([string]$details[$label])
    }
    $body='<h2>Workplace evidence completed</h2><table>'+($rows -join '')+'</table><p>Review the execution log for source-quality warnings. Batch retention runs after this notification; its result and any warning are in the same log. Power BI refresh and publication are separate operations.</p>'
    # Use the shared sender/recipient and Graph/SMTP configuration; no attachments.
    SendEmailHtmlReport -Subject "$scriptName completed - batch $BatchId" -BodyHtml $body -MailPurpose Report -AllowAttachments:$false
}
try {
    if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }
    if ($TransferOnly -and ($Offline -or $ValidateOnly -or $RepairLegacyHistory -or $WorkforceDiagnostic -or $ConvertMetadata)) { throw 'TransferOnly cannot be combined with other modes.' }
    if ($ExpectedBatchId -and -not $TransferOnly) { throw 'ExpectedBatchId requires TransferOnly.' }
    if ($ConvertMetadata -and ($ValidateOnly -or $RepairLegacyHistory -or $WorkforceDiagnostic)) { throw 'ConvertMetadata is a separate manual mode.' }
    if ($ApplyMetadataConversion -and -not $ConvertMetadata) { throw 'ApplyMetadataConversion requires ConvertMetadata.' }
    if ($WorkforceDiagnostic -and ($ValidateOnly -or $RepairLegacyHistory)) { throw 'WorkforceDiagnostic cannot be combined with ValidateOnly or RepairLegacyHistory.' }
    if (($ApplyRepair -or $RepairWeeks -or $ExpectedRepairFileCount) -and -not $RepairLegacyHistory) { throw 'Repair arguments require the manual RepairLegacyHistory mode.' }
    if ($RepairLegacyHistory -and ($ValidateOnly -or -not $RepairWeeks -or $ExpectedRepairFileCount -lt 1)) {
        throw 'Manual repair requires RepairWeeks and ExpectedRepairFileCount, without ValidateOnly. Preview is default; ApplyRepair enables writes.'
    }
    $config = Read-SmartM365JsonConfig -Path (Join-Path $PSScriptRoot "$scriptName.local.json") -Required
    function ConfigValue([string]$Name) {
        $value = $config[$Name]
        if ($null -eq $value -or $value -in @('','__USE_GLOBAL__','USE_GLOBAL')) { $value = $effective.$Name }
        if ($value -isnot [string]) { return $value }
        for ($i=0; $i -lt 10 -and $value -match '\{\{'; $i++) {
            foreach ($match in [regex]::Matches($value,'\{\{([A-Za-z0-9_.-]+)\}\}')) {
                $replacement = $effective.($match.Groups[1].Value)
                if ($null -eq $replacement) { throw "Unresolved configuration token: $($match.Groups[1].Value)" }
                $value = $value.Replace($match.Value,[string]$replacement)
            }
        }
        if ($value -match '\{\{') { throw "Unresolved configuration: $Name" }
        $value
    }
    $latest = [IO.Path]::GetFullPath((ConfigValue 'LatestCsvFolderPath'))
    $all = [IO.Path]::GetFullPath((ConfigValue 'DataAllRootPath'))
    $data = Split-Path $latest -Parent
    if ((Split-Path $latest -Leaf) -ne 'DATA-LAST' -or $all.TrimEnd('\') -ne (Join-Path $data 'DATA-ALL').TrimEnd('\')) {
        throw 'Prepared generators require sibling DATA-LAST and DATA-ALL under the same tenant data root.'
    }
    $output = Join-Path $data 'DATA-POWERBI'
    $product = ConfigValue 'WorkplaceIntelligenceRootPath'
    if ([string]::IsNullOrWhiteSpace($product)) { $product = Join-Path (Split-Path $smartRoot -Parent) 'SmartWorkplaceIntelligence' }
    $pipeline = Join-Path $product 'scripts/PreparedEvidencePipeline.psm1'
    # Private account-classification rules: resolved once, the run fails early when they are missing.
    Import-Module (Join-Path $smartRoot 'SmartInventory/Common/SmartM365.AccountClassification.psd1') -MinimumVersion '1.0.0' -ErrorAction Stop
    $accountClassificationPath = Resolve-SmartM365AccountClassificationPath -Path (Join-Path $smartRoot 'SmartInventory/Config/AccountClassification.local.json')
    if (-not (Test-Path -LiteralPath $pipeline)) { throw 'Deploy SmartWorkplaceIntelligence scripts/config beside SmartM365, or configure WorkplaceIntelligenceRootPath.' }
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.104' -ErrorAction Stop
    $coreLoaded=$true
    # Refuse a stale sibling product before downloading mappings or preparing a batch.
    if (-not ($TransferOnly -or $ConvertMetadata -or $RepairLegacyHistory -or $WorkforceDiagnostic)) {
        $productContract = (Read-SmartM365JsonDocument (Join-Path $product 'config/prepared-evidence-contract.json')).Document
        $requiredEstateFiles = @('AllAccountsEvidence.csv','AllDevicesEvidence.csv','DeviceSystemEvidence.csv')
        foreach ($requiredFile in $requiredEstateFiles) {
            if (@($productContract.tables | Where-Object file -EQ $requiredFile).Count -ne 1) {
                throw "Prepared product contract is incomplete: $requiredFile. Deploy the reviewed SmartWorkplaceIntelligence scripts/config/reference alongside SmartM365. No preparation started."
            }
        }
    }
    # Prepared batch transfer: explicit TransferOnly, or a normal run with EnableSharePointUpload (default true).
    $batchUploadEnabled = $TransferOnly -or (-not ($Offline -or $ValidateOnly -or $WorkforceDiagnostic) -and [bool](ConfigValue 'EnableSharePointUpload'))
    # Run logs and transcript are uploaded by the shared completion step in every mode except Offline,
    # independently of the prepared batch transfer. Offline never contacts SharePoint.
    $global:EnableSharePointUpload = -not $Offline
    $global:SharePointSiteHostname = [string](ConfigValue 'SharePointSiteHostname')
    $global:SharePointSitePath = [string](ConfigValue 'SharePointSitePath')
    $global:SharePointLibraryDisplayName = [string](ConfigValue 'SharePointLibraryDisplayName')
    $global:SharePointTargetFolderPath = [string](ConfigValue 'SharePointTargetFolderPath')
    $global:AppId = [string](ConfigValue 'AppId')
    $global:TenantId = [string]$effective.TenantId
    $global:Thumb = [string](ConfigValue 'Thumb')
    $global:Thumbprint = $global:Thumb
    $work = ConfigValue 'PreparedWorkRootPath'
    if ([string]::IsNullOrWhiteSpace($work)) { $work=Join-Path ([IO.Path]::GetTempPath()) "SmartWorkplaceIntelligence/$($effective.ProfileKey)" }
    if ($WorkforceDiagnostic) { $output=Join-Path $work 'workforce-diagnostics' }
    InitializeScriptEnvironment -OutputPath $output -LogFileName $scriptName -CallerScriptPath $PSCommandPath | Out-Null
    Start-SmartM365RunTranscript -Path $global:logTranscriptFile -Append | Out-Null
    $transcriptStarted = $true
    if (-not ($TransferOnly -or $ConvertMetadata -or $RepairLegacyHistory -or $ValidateOnly)) {
        $phase = 'ImportExcel startup health'
        Initialize-SmartM365ImportExcel -SkipRepositoryRepair:$Offline | Out-Null
    }
    WriteLog -Message "Prepared batch SharePoint transfer enabled: $batchUploadEnabled; run log SharePoint upload enabled: $($global:EnableSharePointUpload)." -Level INFO
    Import-Module $pipeline -Force
    if ($ConvertMetadata) {
        $phase='Convert prepared metadata names'
        Import-Module (Join-Path $product 'scripts/PreparedMetadata.psm1') -Force
        $conversion=Convert-PreparedMetadataNames -OutputRoot $output -TenantKey $effective.TenantKey -Apply:$ApplyMetadataConversion
        WriteLog -Message "Metadata conversion: applied=$($conversion.Applied); batches=$($conversion.Batches); checked CSVs=$($conversion.CheckedCsvFiles); metadata files=$($conversion.MetadataFiles); current batch=$($conversion.BatchId). No CSV recalculation or SharePoint transfer." -Level INFO
        $conversion | ConvertTo-Json -Depth 5 | Write-Host
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    if (-not $ValidateOnly -and -not $RepairLegacyHistory -and -not $WorkforceDiagnostic) {
        Import-Module (Join-Path $product 'scripts/PreparedMetadata.psm1') -Force
        Initialize-PreparedMetadataNames -OutputRoot $output -TenantKey $effective.TenantKey
        if(Test-Path -LiteralPath (Join-Path $output 'current.json')){throw 'Legacy prepared metadata is preserved while JSON transport policy is Readers. Activate the approved JsonText deployment before preparation or transfer.'}
    }
    # Fail before downloads/copying/calculations when cloud publication is enabled but incomplete.
    if ($batchUploadEnabled) {
        Import-Module (Join-Path $product 'scripts/PreparedSharePointTransfer.psm1') -Force
        $dataCloudRoot=ConvertTo-SmartM365SharePointDataRootPath -TargetFolderPath (ConfigValue 'SharePointTargetFolderPath')
        $cloudRoot=Resolve-PreparedSharePointFolder -ConfiguredPath (ConfigValue 'PreparedSharePointFolderPath') -NormalizedDataRoot $dataCloudRoot
        foreach($key in 'SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','AppId','Thumb') {
            if ([string]::IsNullOrWhiteSpace([string](ConfigValue $key))) { throw "SharePoint publication configuration missing: $key. No prepared calculations started." }
        }
        $cloud=@{
            Enabled=$true;SiteHostname=(ConfigValue 'SharePointSiteHostname');SitePath=(ConfigValue 'SharePointSitePath')
            LibraryDisplayName=(ConfigValue 'SharePointLibraryDisplayName');AppId=(ConfigValue 'AppId')
            TenantId=$effective.TenantId;Thumbprint=(ConfigValue 'Thumb')
        }
        $transferParameters=@{
            OutputRoot=$output;TenantKey=$effective.TenantKey;WorkRoot=$work
            ContractPath=(Join-Path $product 'config/prepared-evidence-contract.json')
            UploadFile={
                param($localPath,$relativePath)
                if((Get-SmartM365SharePointRelativeFilePath $localPath) -cne [IO.Path]::GetFileName($localPath)){throw 'Prepared batch must not be nested under raw/log folders.'}
                $parent=$relativePath.LastIndexOf('/')
                $target=if($parent -lt 0){$cloudRoot}else{$cloudRoot+'/'+$relativePath.Substring(0,$parent)}
                SmartM365.Core\Invoke-SmartM365SharePointCsvUpload -LocalFilePath $localPath -TargetFolderPath $target -EnsureParentFolders @cloud
            }.GetNewClosure()
            DownloadFile={
                param($destination,$relativePath)
                SmartM365.Core\Invoke-SmartM365SharePointFileDownload -LocalFilePath $destination -SharePointRelativePath $relativePath -TargetFolderPath $cloudRoot -Force @cloud
            }.GetNewClosure()
        }
        WriteLog -Message "Prepared SharePoint target: $cloudRoot. Each uploaded file will be read back and SHA256-verified before publishing the root pointer." -Level INFO
    }
    if ($TransferOnly) {
        $phase='SharePoint existing-batch transfer'
        $transfer=Invoke-PreparedStep { Send-PreparedEvidenceBatch @transferParameters -ExpectedBatchId $ExpectedBatchId }
        Assert-PreparedTransferResult -Result $transfer -ExpectedBatchId $ExpectedBatchId
        WriteLog -Message "Transfer complete: batch=$($transfer.BatchId); CSVs=$($transfer.CsvFiles); verified files=$($transfer.VerifiedFiles); pointer verified=$($transfer.PointerVerified); audit=$($transfer.AuditPath). No collection, mapping download or CSV generation." -Level INFO
        $phase='Success email notification'
        Send-PreparedSuccessMail -BatchId $transfer.BatchId -CsvFiles $transfer.CsvFiles -VerifiedTransfer $transfer
        Invoke-PostTransferRetention $transfer
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    if ($RepairLegacyHistory) {
        $phase='Repair historical TenantKey columns'
        Import-Module (Join-Path $product 'scripts/Repair-PreparedHistoryTenantKeys.psm1') -Force
        $repair = Invoke-PreparedHistoryTenantRepair -DataRoot $data -TenantKey $effective.TenantKey -Weeks @($RepairWeeks.Split(',') | ForEach-Object {$_.Trim()}) -ExpectedFileCount $ExpectedRepairFileCount -Apply:$ApplyRepair
        WriteLog -Message "Historical TenantKey repair: $($repair.Status); files=$($repair.Files); applied=$($repair.Applied); backup=$($repair.BackupPath). No prepared CSV generation or SharePoint access." -Level INFO
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    $mappingRoot = ''
    if (-not $Offline) {
        $phase='Download SharePoint classification workbooks'
        $mappingFolder = [string](ConfigValue 'PreparedMappingSharePointFolderPath')
        if ([string]::IsNullOrWhiteSpace($mappingFolder)) { $mappingFolder = [string](ConfigValue 'SharePointTargetFolderPath') }
        $mappingConnection = @{
            Enabled=$true; SiteHostname=(ConfigValue 'SharePointSiteHostname'); SitePath=(ConfigValue 'SharePointSitePath')
            LibraryDisplayName=(ConfigValue 'SharePointLibraryDisplayName'); TargetFolderPath=$mappingFolder
            AppId=(ConfigValue 'AppId'); TenantId=$effective.TenantId; Thumbprint=(ConfigValue 'Thumb')
        }
        foreach ($key in 'SiteHostname','SitePath','LibraryDisplayName','TargetFolderPath','AppId','TenantId','Thumbprint') {
            if ([string]::IsNullOrWhiteSpace([string]$mappingConnection[$key])) { throw "Required SharePoint mapping connection value missing: $key" }
        }
        $downloadMapping = {
            param($destination, $name)
            SmartM365.Core\Invoke-SmartM365SharePointFileDownload -LocalFilePath $destination -SharePointRelativePath $name -Force @mappingConnection
        }.GetNewClosure()
        $mappingExceptions=[bool](ConfigValue 'PreparedMailboxMigrationExceptionsEnabled') -and -not $WorkforceDiagnostic
        $migrationMappingArgs=@{
            SharePointMigrationEnabled=([bool](ConfigValue 'PreparedSharePointMigrationEnabled') -and -not $WorkforceDiagnostic)
            SharePointMigrationWorkbookFileName=[string](ConfigValue 'PreparedSharePointMigrationWorkbookFileName')
            SharePointMigrationBindingsFileName=[string](ConfigValue 'PreparedSharePointMigrationBindingsFileName')
        }
        $mappingRoot = Invoke-PreparedStep { Receive-PreparedMappingWorkbooks -DataRoot $data -DownloadFile $downloadMapping -MailboxMigrationExceptionsEnabled:$mappingExceptions -PythonCommand ([string](ConfigValue 'PreparedPythonCommand')) @migrationMappingArgs }
        WriteLog -Message "Classification workbooks downloaded, validated and retained in DATA: $data. Mailbox exceptions enabled=$mappingExceptions; this run uses its isolated validated mapping set." -Level INFO
    }
    if ($WorkforceDiagnostic) {
        $phase='Workforce memory diagnostic'
        Import-Module (Join-Path $product 'scripts/WorkforceMemoryDiagnostic.psm1') -Force
        $diagnostic=Invoke-WorkforceMemoryDiagnostic -DataRoot $data -WorkRoot $work -MappingRoot $mappingRoot -AccountClassificationConfigPath $accountClassificationPath
        WriteLog -Message "Workforce diagnostic: exit=$($diagnostic.ExitCode); last stage=$($diagnostic.LastStage); samples=$($diagnostic.Samples); sample errors=$($diagnostic.SampleErrors); sampled peak private MiB=$([math]::Round($diagnostic.SampledPeakPrivateBytes/1MB,1)); private logs=$($diagnostic.RunRoot). No prepared publication." -Level INFO
        if ($diagnostic.SampleErrors -gt 0) { WriteLog -Message 'Some memory measurements were unavailable. Inspect memory.csv; this diagnostic is not a complete memory profile.' -Level WARNING }
        if ($diagnostic.ExitCode -ne 0) { throw "Workforce diagnostic worker failed (exit $($diagnostic.ExitCode)); last stage: $($diagnostic.LastStage). Logs: $($diagnostic.RunRoot)" }
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    $phase='Prepare and validate'
    $preparedPython=if($config['PreparedPythonCommand']){[string]$config['PreparedPythonCommand']}else{'python'}
    # Independent generators run concurrently (1 = sequential); the default keeps one vCPU free on a 4-vCPU server.
    $preparedParallel=ConfigValue 'PreparedMaxParallelGenerators'
    $preparedParallel=if($null -eq $preparedParallel -or [string]$preparedParallel -eq ''){3}else{[int]$preparedParallel}
    if($preparedParallel -lt 1 -or $preparedParallel -gt 8){throw 'PreparedMaxParallelGenerators must be between 1 and 8.'}
    # A kept Office activation report older than this (hours since its last acquisition) is warned about, never blocked.
    $officeActivationWarning=ConfigValue 'PreparedOfficeActivationWarningAgeHours'
    $officeActivationWarning=if($null -eq $officeActivationWarning -or [string]$officeActivationWarning -eq ''){72}else{[int]$officeActivationWarning}
    if($officeActivationWarning -lt 1 -or $officeActivationWarning -gt 8760){throw 'PreparedOfficeActivationWarningAgeHours must be between 1 and 8760.'}
    $params = @{
        DataRoot=$data;OutputRoot=$output;WorkRoot=$work;TenantKey=$effective.TenantKey
        AccountClassificationConfigPath=$accountClassificationPath
        AdCountryConfiguration=$effective
        MaxSourceAgeHours=[int](ConfigValue 'PreparedMaxSourceAgeHours')
        AgeOverrides=(ConvertTo-PreparedAgeOverrides (ConfigValue 'PreparedSourceAgeOverrides'))
        AllowLegacyTenantless=[bool](ConfigValue 'PreparedAllowLegacyTenantless')
        AllowEmptyTables=[string[]](ConfigValue 'PreparedAllowEmptyTables')
        ValidateOnly=[bool]$ValidateOnly
        MappingRoot=$mappingRoot
        PythonCommand=$preparedPython
        MaxParallelGenerators=$preparedParallel
        OfficeActivationWarningAgeHours=$officeActivationWarning
        # Weeks pruned by the WeeklyHistory ceiling leave the published history instead of blocking it.
        HistoryWindowWeeks=[int](Get-SmartM365WeeklyHistoryMaxRetentionWeeks)
        HybridAllowedTargetDomains=[string[]]@(ConfigValue 'PreparedHybridAllowedTargetDomains' | Where-Object { $_ })
        MailboxMigrationExceptionsEnabled=[bool](ConfigValue 'PreparedMailboxMigrationExceptionsEnabled')
        SharePointMigrationEnabled=[bool](ConfigValue 'PreparedSharePointMigrationEnabled')
        SharePointMigrationWorkbookFileName=[string](ConfigValue 'PreparedSharePointMigrationWorkbookFileName')
        SharePointMigrationBindingsFileName=[string](ConfigValue 'PreparedSharePointMigrationBindingsFileName')
        SharePointMigrationTargetTenantUrl=[string](ConfigValue 'PreparedSharePointMigrationTargetTenantUrl')
        SharePointMigrationReviewMaxAgeDays=[int](ConfigValue 'PreparedSharePointMigrationReviewMaxAgeDays')
        SharePointMigrationReportingTimeZoneId=[string](ConfigValue 'PreparedSharePointMigrationReportingTimeZoneId')
    }
    if ($params.MailboxMigrationExceptionsEnabled) {
        $bindingPath=[string](ConfigValue 'PreparedMailboxMigrationBindingsPath')
        if ([string]::IsNullOrWhiteSpace($bindingPath)) {
            $bindingPath=ConvertTo-SmartM365InventoryConfigLocation -Path (Join-Path $PSScriptRoot 'SmartM365-WorkplaceEvidence-Prepare.MailboxMigrationBindings.local.json')
            # An absent default registry leaves address-only rows as candidates, never approved.
            $bindingPath=[string](Get-SmartM365JsonReadPath -Path $bindingPath -Optional)
        } else {
            # Prefer .json.txt when present; explicit missing configuration is an error.
            $bindingPath=Get-SmartM365JsonReadPath -Path $bindingPath
        }
        if ($bindingPath) { $params.MailboxMigrationBindingsPath=$bindingPath }
    }
    $result = Invoke-PreparedStep { Invoke-PreparedEvidencePipeline @params }
    if (-not $ValidateOnly) { Register-PreparedRunCsvFiles -PreparedResult $result }
    if (-not $ValidateOnly) { WriteLog -Message "Local batch published and validated: $($result.BatchId); CSV files=$($result.Files); path=$($result.BatchPath). Cloud transfer is a separate step." -Level INFO }
    if (-not $ValidateOnly -and $batchUploadEnabled) {
        $phase='SharePoint batch transfer'
        $transfer=Invoke-PreparedStep { Send-PreparedEvidenceBatch @transferParameters -ExpectedBatchId $result.BatchId }
        Assert-PreparedTransferResult -Result $transfer -ExpectedBatchId $result.BatchId
        WriteLog -Message "SharePoint batch and pointer verified: $($transfer.BatchId); files=$($transfer.VerifiedFiles); audit=$($transfer.AuditPath)." -Level INFO
    }
    $recap = if ($ValidateOnly) { "Source preflight: $($result.SourceFiles) files; $($result.Identity.CsvFiles) CSVs and $($result.Identity.Rows) rows checked; no generation/publication." } else { "$($result.Files) prepared CSVs; observed history retained; batch $($result.BatchId)." }
    WriteLog -Message $recap -Level INFO
    # The verified pointer already exposes the batch, so notify before the long retention run.
    if (-not ($Offline -or $ValidateOnly -or $WorkforceDiagnostic)) {
        $phase='Success email notification'
        Send-PreparedSuccessMail -BatchId $result.BatchId -CsvFiles $result.Files -VerifiedTransfer $transfer
        $phase='Success Teams notification'
        Send-SmartM365TeamsNotification -Title $scriptName -Message 'Preparation completed.' -Level SUCCESS -Channel Infos -ResultSummary $recap -Facts @{Tenant=$effective.TenantKey;Output=$output} | Out-Null
    }
    if ($transfer) { Invoke-PostTransferRetention $transfer }

    $smartM365RunGuard.ReachedTerminal=$true
} catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_

    $failure=$_
    if ($coreLoaded) {
        WriteLog -Message "$phase failed: $($_.Exception.Message)" -Level ERROR
        $failedBatchId = [string]$failure.Exception.Data['PreparedFailureAuditBatchId']
        if ($phase -eq 'Prepare and validate' -and $batchUploadEnabled -and $failedBatchId -and $transferParameters) {
            try {
                $auditTransfer = Send-PreparedFailureAudit -OutputRoot $output -TenantKey $effective.TenantKey -BatchId $failedBatchId -WorkRoot $work -UploadFile $transferParameters.UploadFile -DownloadFile $transferParameters.DownloadFile
                WriteLog -Message "Failure audit uploaded and verified: batch=$($auditTransfer.BatchId); files=$($auditTransfer.VerifiedFiles). Current pointer unchanged." -Level INFO
            } catch {
                WriteLog -Message "Failure audit transfer failed for batch ${failedBatchId}: $($_.Exception.Message)" -Level ERROR
            }
        }
        if (-not ($Offline -or $ValidateOnly -or $WorkforceDiagnostic)) {
            $message=$_.Exception.Message
            $help='https://chatgpt.com/?q='+[uri]::EscapeDataString("Explain SmartM365 preparation failure in phase ${phase}: $message")
            Send-SmartM365TeamsNotification -Title $scriptName -Message $message -Level ERROR -Channel Alerts -HelpUrl $help -Facts @{Tenant=$effective.TenantKey;Phase=$phase;Output=$output;Log=$global:LogTextFile;InnerException=[string]$_.Exception.InnerException} | Out-Null
            SendEmailHtmlReport -Subject "$scriptName failed" -BodyHtml ([Net.WebUtility]::HtmlEncode("$phase : $message")) -MailPurpose Error
        }
    } else { Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date),$_.Exception.Message) }
} finally {
    $smartM365RunGuard.InCleanup=$true

    if ($mappingRoot) { try { Invoke-PreparedStep { Remove-PreparedMappingWorkbooks -DataRoot $data -MappingRoot $mappingRoot } | Out-Null } catch { WriteLog -Message "Downloaded mapping cleanup incomplete: $($_.Exception.Message)" -Level WARNING } }
    if ($transcriptStarted) { try { Stop-SmartM365RunTranscript | Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } catch { WriteLog -Message "Transcript finalization failed: $($_.Exception.Message)" -Level WARNING } }
    if ($coreLoaded) { Complete-SmartM365ExecutionContext -Status $(if($failure){'Failed'}else{'Auto'}) -ErrorRecord $failure -FailureStage $(if($failure){$phase}else{''}) }
    else { Write-SmartM365CompletionBanner -Status 'Failed' }
}
if ($failure) { $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode }
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCr4yHe1H40+oHu
# AavZfWO1UoBLyyOQrazWQyNt59QjH6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDKw3s5KfNmcWLjgW3tkRyz
# 09FzkZPuCBSLUtGvpsKLHTANBgkqhkiG9w0BAQEFAASCAYAQ2OA/L76BJAgDpsuV
# g493ZQeIrsqM8wc/9SZz/fmcHhUTK5QkIiInO69drKmOlB9LY+JJlGbcsgvg3k2F
# yEW2TWWZxpAbN6wEt2a0ZchyWjDMcgGTHcVGRhCYnIuJQtjLrAu6/VV6OjzhmHgw
# k4i2UT2m+YgJchwtMJou26eHwdcqugpSAR/Vs2MOGfBDay9S0FHPX00AGL6qDZPl
# HV/Ypk6WWcVv1tNAETsqgfi7+LMb7cvXOyChnF5xcaDnfYNfDS+s0DE3qxsaDNsF
# eV2WEpmGj+WDQ6BkWhP1J77fgN9HxK3jcutOKfWkgxHy3eRowCbbAZvKWGhFZG5K
# TaMJAjUQ9GiX6fQotC5h4h9eaO/ydA7MXP0ohuEMyljg/7vWBpRSwQtkiMB738jg
# 6T4eKsD1vXw7Q4k41bwtI9tfuxRbw+NYJMjQSPRtNLcSK+IK9EMr4yBetU4F3pUD
# FVZPTz4q3OVj+WRqJfy7UlelVYwm+aBU+4Kr/JIsNxhWo/U=
# SIG # End signature block
