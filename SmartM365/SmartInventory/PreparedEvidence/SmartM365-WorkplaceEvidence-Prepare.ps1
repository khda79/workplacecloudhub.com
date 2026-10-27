<#
.SYNOPSIS
    Prepare a validated, versioned SmartWorkplaceIntelligence batch after raw collectors.
.VERSION
    0.4.9
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
                -ConfigurationPath (Get-SmartM365JsonReadPath (Join-Path $PSScriptRoot "$scriptName.local.json")) `
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
    $body='<h2>Workplace evidence completed</h2><table>'+($rows -join '')+'</table><p>Review the execution log for source-quality and retention warnings. Power BI refresh and publication are separate operations.</p>'
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
        Invoke-PostTransferRetention $transfer
        WriteLog -Message "Transfer complete: batch=$($transfer.BatchId); CSVs=$($transfer.CsvFiles); verified files=$($transfer.VerifiedFiles); pointer verified=$($transfer.PointerVerified); audit=$($transfer.AuditPath). No collection, mapping download or CSV generation." -Level INFO
        $phase='Success email notification'
        Send-PreparedSuccessMail -BatchId $transfer.BatchId -CsvFiles $transfer.CsvFiles -VerifiedTransfer $transfer
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
        Invoke-PostTransferRetention $transfer
        WriteLog -Message "SharePoint batch and pointer verified: $($transfer.BatchId); files=$($transfer.VerifiedFiles); audit=$($transfer.AuditPath)." -Level INFO
    }
    $recap = if ($ValidateOnly) { "Source preflight: $($result.SourceFiles) files; $($result.Identity.CsvFiles) CSVs and $($result.Identity.Rows) rows checked; no generation/publication." } else { "$($result.Files) prepared CSVs; observed history retained; batch $($result.BatchId)." }
    WriteLog -Message $recap -Level INFO
    if (-not ($Offline -or $ValidateOnly -or $WorkforceDiagnostic)) {
        $phase='Success email notification'
        Send-PreparedSuccessMail -BatchId $result.BatchId -CsvFiles $result.Files -VerifiedTransfer $transfer
        $phase='Success Teams notification'
        Send-SmartM365TeamsNotification -Title $scriptName -Message 'Preparation completed.' -Level SUCCESS -Channel Infos -ResultSummary $recap -Facts @{Tenant=$effective.TenantKey;Output=$output} | Out-Null
    }

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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC+dW1qw5ZJkZGy
# pTPO7XPwqV0dLsV8s9PDkyo4dbqeiKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEP6CG8fAruEwlrsM6gHy7W0RS4TYz2UEeP3BUMfdUcWMA0GCSqG
# SIb3DQEBAQUABIIBgBpoNkbDtg0zVcb4rF8P7N/zOghUOoZsJ87/iprUnYR/2Z0X
# wAwGpPgUnD4CLJTw5xN9icGvKJDRb+8UiLEKcWeS6h1LDgPtiYHEzIrQ49AfZHnS
# EWBKSjePyN3dCAL6CHtURH57N2leXQivwk+K2g/D6NxCKAdl3pwvMd+KvKtATgFz
# gl5djJbdUISmmvTQOnW2IerbQuXz3ZMyz5Br97fWZ3G2VTwwoNN8n67AL2vWWa8p
# 2INbzUZG/v7a7hgJ4D7WEBkOV28/wTtPk8oaTZx2Z4evZHR0rT7j7dKhMYqAhAAN
# 1rYZffDYOZA39avZdO4ob27DKhJXnYT6vvDVaXIpwZFawJi0OyyzxI9QwkFyHETd
# EBNUeBjFD+3FyfrFHOKk+anVGWmRAjEiCRS4E2drZDFmjD5AYMLEavDBOAJcASjt
# 4ZD36h9IfBU5ZsEqUCDgKGp+JR1ST4z4i2/173FaBnfTtzMAoq3Vzz3bl0td/X2t
# WN0hZtnfOU98hRKllKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcyMDUy
# NTVaMC8GCSqGSIb3DQEJBDEiBCA3jYedfCToPdaXNwpxHItJLi4f9nrbTrp7yhmi
# QEcbzzANBgkqhkiG9w0BAQEFAASCAgARRYup7KiScslIxCQnHJK0MySNoHVmJ2XF
# /AC8ASXgcxACwvWApiNxREqCJ/yrqXNNKN6Iis9j/ium0YM2/JVhM8gXnobIugvZ
# Pi4f6395YlKM2+f03N0ROpR0c9qyqYpDOOcao8zLNaWtZBByQqJGnjSrkvHv1uqI
# +T+cokQIkAFEnDgA6gP9RPmLqIO6H2HjzBQ1rXBhHZFkexbF0oX7f+L1oMZYiM4Y
# VQgZPfHzaoR6s5V/JvnL/mYQDU/oA1Y7hLchLQQ5l+xkxG1vlPgTw233/ha0gtjo
# B2VlNP/IfQoeLGyCl2IaExFST6yzZHd2+a/Tc6iLQ8D1SPvAUwYu/LM6Fnmjzlsa
# E0+8AuDoBxxBG8bKFl4SdmAM+wajqwf6gnofIQcwtgveVtC0vE+m9zwVhs/f30wP
# qAjfELH5K+qbHgPxNlHZ9BPCiLrBduSO2PWA+YTvu8H6FK63bpOGynjOxIFJzHDM
# 21vFJTMVHwtdhMnlrVTdmna26sx+nIxUY31EwhkZq1EnWi+3drbNXRruO/sXmDX7
# seEB2b6xhf6/n5A4EVV8+msBq4WETRuYvKnaH2FsZp9i7O2DAR6hz5Jm21fHG3BZ
# msAAHveghHraQNF66yyai6/jR8IE9GzxF++U2tMP/+T0qkWWKpcxtyFQIf5Dq2Iu
# x86l67ngWw==
# SIG # End signature block
