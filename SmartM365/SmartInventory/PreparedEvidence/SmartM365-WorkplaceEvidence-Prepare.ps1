<#
.SYNOPSIS
    Prepare a validated, versioned SmartWorkplaceIntelligence batch after raw collectors.
.VERSION
    0.4.17
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
. (Join-Path $smartM365LifecycleRoot 'SmartInventory/Common/SmartM365.ReportRenderers.ps1')
. (Join-Path $smartM365LifecycleRoot 'SmartInventory/Common/SmartM365.ReportInputs.ps1')
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
                -ExpectedPublishedBatchId $VerifiedTransfer.BatchId -ExpectedPointerSHA256 $VerifiedTransfer.PointerSHA256 -NoCompletionBanner -Simple
        }
        $level=if($retention.Status -in @('Blocked','PartialOrBlocked')){'WARNING'}else{'INFO'}
        $reasons=if($retention.PSObject.Properties['Plan'] -and @($retention.Plan.BlockedReasons).Count){
            ' Reasons: '+(@($retention.Plan.BlockedReasons) -join '; ')
        }else{''}
        WriteLog -Message "Post-transfer retention: $($retention.Status); protected=$(@($retention.Protected) -join ', '); actions=$(@($retention.Actions) -join '; ').$reasons" -Level $level
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
    # Warnings this run logged so far (source quality, qualification); post-transfer retention logs after this mail.
    $warnings=[Collections.Generic.List[string]]::new()
    if ($logPath -and (Test-Path -LiteralPath $logPath -PathType Leaf)) {
        try {
            $stream=[IO.File]::Open($logPath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
            try {
                $reader=[IO.StreamReader]::new($stream)
                while ($null -ne ($line=$reader.ReadLine())) {
                    if ($line -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[WARN(?:ING)?\]\s*(.+)$') { $warnings.Add($Matches[1].Trim()) }
                }
            } finally { $stream.Dispose() }
        } catch { WriteLog -Message "Run warnings could not be read for the completion email: $($_.Exception.Message)" -Level INFO }
    }
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
        'Warnings'=$warnings.Count
    }
    if ($batchUploadEnabled) { $details['Verified SharePoint files']=$VerifiedTransfer.VerifiedFiles }
    try {
        $batchManifestPath = Join-Path $output "batches/$BatchId/batch.json.txt"
        $reportManifestHash = (Get-FileHash -LiteralPath $batchManifestPath -Algorithm SHA256 -ErrorAction Stop).Hash
        Save-SmartM365ReportInputs -Producer 'SmartM365-WorkplaceEvidence-Prepare' -LatestFolder $latest -TenantKey $effective.TenantKey -BatchId $BatchId -BatchManifestHash $reportManifestHash -InputData @{
            Details=$details; Warnings=$warnings.ToArray()
        }
    } catch { Write-Warning "Could not associate the completion mail with its prepared batch: $($_.Exception.Message)" }
    $body=New-SmartM365PreparedCompletionMailBody -Details $details -Warnings $warnings.ToArray()
    # Use the shared sender/recipient and Graph/SMTP configuration; no attachments.
    SendEmailHtmlReport -Subject "$scriptName completed - batch $BatchId" -BodyHtml $body -MailPurpose Report -AllowAttachments:$false -ReportCadenceKey 'WorkplaceEvidence.Completion'
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
    $config = Read-SmartM365ScriptConfiguration -Path (Join-Path $PSScriptRoot "$scriptName.local.json") -AsHashtable
    function ConfigValue {
    param([string]$Name)
    Get-SmartM365ConfigurationValue -Name $Name -LocalConfig $config -EffectiveConfig $effective -Strict -ResolveDefault
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
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.130' -ErrorAction Stop
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
    . (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-SharePointUploadPolicy.ps1')
    # Preserve the historic independent defaults for the batch and logs when no category setting exists.
    $uploadPolicy = Resolve-SmartM365PreparedUploadPolicy -LocalConfig $config -GlobalConfig $effective -LegacyBatchEnabled ([bool](ConfigValue 'EnableSharePointUpload')) -Offline:$Offline -ValidateOnly:$ValidateOnly -WorkforceDiagnostic:$WorkforceDiagnostic -TransferOnly:$TransferOnly
    $batchUploadEnabled = $uploadPolicy.BatchEnabled
    $global:EnableSharePointUpload = $uploadPolicy.LogEnabled
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
    WriteLog -Message "Prepared batch SharePoint transfer enabled: $batchUploadEnabled (policy: $($uploadPolicy.BatchSource)); run log SharePoint upload enabled: $($global:EnableSharePointUpload) (policy: $($uploadPolicy.LogSource))." -Level INFO
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
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAeMSIiCLprMKRJ
# Ez2v5lvmBauzlZOoN69ElSMq25zboKCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg1xxIip7iR/J4
# Ra3AhHfzWukPPY5CXsoWLNwNy8NIK5gwDQYJKoZIhvcNAQEBBQAEggGAi8sfvH6s
# WnTDENDVXqTFbX902t5rz6eQuMrBt9m3NWSnRDuAaHQ/TSjx5keow9UG8ijt5Mpn
# E+/RhA+Ytgzt0QgoTwvFV9NkUtIXmqnZLOEpwwd4xBmvtPGdCH7xHmv0WZXXWgd/
# O1kkWoLjj8Hn6epL8GGg+y2DVu/ei8fXinD7iJuGPWxz+teGXiWP5plPSoWMv/xm
# PlRo1kb09XQXlzrPqKk4ksyMyTMGVI0ibM6J1j/I5MoT3IXO56D8Tb2drthXnzbf
# pkht1Dv7PqDyDwZ1bbZzrDarkb8MUColmqnwLULyTiG3vhB1FN1yd6tiVv1RXQ5L
# 3PzIj4x5o+jWxF5vU3yNzQPC3M/KkjOGT1bJj1hdONT7X8YWPmyVqQfWfCCXqYSp
# b5DG9n4Kh9sPg4m3e8bDUD3N7AFTZtuOMjMc8wk31fltd2WjXh7Wr9zugV1bbMAV
# ah1jcZqjjXZU7wfyTsLbVCfEc/oNPkAaTX4X95UxURk3DRmhJ0T44A2OoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNTAxWjA/BgkqhkiG9w0BCQQxMgQwiQJZzO/vnBMTSdcSDQ31GDdW
# 3QbvhcV6GQD/Ywx8cl5OPcQy4nAME7ZTY/zPNOgfMA0GCSqGSIb3DQEBAQUABIIC
# ABceDLz03qg+3S5YOWp8p2jjoXJpx8UdFo36z7r8cIUDYfzl1wwN0S3EN/TT0H9B
# +UvQpqtxwqyj5sqblo1gefRJV8scTNjjiNmSztWnH9s1LTa4hrZTNIqoDw8PAieF
# i/RB5iAYq+WzrC/wWCbzXg74w1W64LTu2jjn31zZeeNatOjZujD3YxP9LARSlKkx
# SS9poSs29xfOkFN8SCopSPu9KRYkwRuoFHwpP/mlLbLzID51YvQrfxIyC7onKCBi
# pGdUwORhDlWdLK93iaPbWtlh0Mxxvka+KXQw3+s9iZ0YWH2+h4fex5AUgciFGqLt
# sGu89268SPUtOQ5d3HcrRgfSyGIjzh5aOq+yEyczHevyx+9qSdkz4JghNFjUDnMV
# ZJND1tFQsY6EaY13U/P6c0nKj3xy3ZnTEsNPraDBTaYLAaG9R8v7BprHWJftzEcS
# he/X2ZqVJ87bxuHGC4cu3opVM0FCZahRznlyuTWNRyHvcC+YbZ46sCICbKEgLrsn
# iQhanQnIBNrksY4xqW+GtyW1awXFqiYsELYwX4OexuT+paQLJNbsdSKViX6x4wcP
# ttkfQJq3uaz+XDWkFxajtonK2k7CnzcUXUwvAVmEBvaUf6uCuR4/Q9T2EB6OHdKi
# YToeyDkr+EO2qmnBXWRzd7o6r5QjKd4stN8lwnQ9Jo/F
# SIG # End signature block
