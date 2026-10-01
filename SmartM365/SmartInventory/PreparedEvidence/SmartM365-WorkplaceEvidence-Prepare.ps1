<#
.SYNOPSIS
    Prepare a validated, versioned SmartWorkplaceIntelligence batch after raw collectors.
.VERSION
    0.1.13
.NOTES
    PowerShell 7. SharePoint mapping reads use the tenant configuration unless Offline.
    Deployment must include the sibling SmartWorkplaceIntelligence/scripts and config folders.
#>
[CmdletBinding()]
param([string]$Tenant='test', [switch]$ValidateOnly, [switch]$Offline, [switch]$WorkforceDiagnostic,
    [switch]$RepairLegacyHistory, [string]$RepairWeeks, [int]$ExpectedRepairFileCount=0, [switch]$ApplyRepair,
    [switch]$ConvertMetadata, [switch]$ApplyMetadataConversion, [switch]$TransferOnly, [string]$ExpectedBatchId)
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
$transcriptStarted = $false
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
    Import-Module (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.58' -ErrorAction Stop
    $coreLoaded=$true
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
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    $transcriptStarted = $true
    WriteLog -Message "Prepared batch SharePoint transfer enabled: $batchUploadEnabled; run log SharePoint upload enabled: $($global:EnableSharePointUpload)." -Level INFO
    Import-Module $pipeline -Force
    if ($ConvertMetadata) {
        $phase='Convert prepared metadata names'
        Import-Module (Join-Path $product 'scripts/PreparedMetadata.psm1') -Force
        $conversion=Convert-PreparedMetadataNames -OutputRoot $output -TenantKey $effective.TenantKey -Apply:$ApplyMetadataConversion
        WriteLog -Message "Metadata conversion: applied=$($conversion.Applied); batches=$($conversion.Batches); checked CSVs=$($conversion.CheckedCsvFiles); metadata files=$($conversion.MetadataFiles); current batch=$($conversion.BatchId). No CSV recalculation or SharePoint transfer." -Level INFO
        $conversion | ConvertTo-Json -Depth 5 | Write-Host
        return
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
        $transfer=Send-PreparedEvidenceBatch @transferParameters -ExpectedBatchId $ExpectedBatchId
        WriteLog -Message "Transfer complete: batch=$($transfer.BatchId); CSVs=$($transfer.CsvFiles); verified files=$($transfer.VerifiedFiles); pointer verified=$($transfer.PointerVerified); audit=$($transfer.AuditPath). No collection, mapping download or CSV generation." -Level INFO
        return
    }
    if ($RepairLegacyHistory) {
        $phase='Repair historical TenantKey columns'
        Import-Module (Join-Path $product 'scripts/Repair-PreparedHistoryTenantKeys.psm1') -Force
        $repair = Invoke-PreparedHistoryTenantRepair -DataRoot $data -TenantKey $effective.TenantKey -Weeks @($RepairWeeks.Split(',') | ForEach-Object {$_.Trim()}) -ExpectedFileCount $ExpectedRepairFileCount -Apply:$ApplyRepair
        WriteLog -Message "Historical TenantKey repair: $($repair.Status); files=$($repair.Files); applied=$($repair.Applied); backup=$($repair.BackupPath). No prepared CSV generation or SharePoint access." -Level INFO
        return
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
        $mappingRoot = Receive-PreparedMappingWorkbooks -WorkRoot $work -DownloadFile $downloadMapping
        WriteLog -Message 'Both classification workbooks downloaded and structurally validated from configured SharePoint source.' -Level INFO
    }
    if ($WorkforceDiagnostic) {
        $phase='Workforce memory diagnostic'
        Import-Module (Join-Path $product 'scripts/WorkforceMemoryDiagnostic.psm1') -Force
        $diagnostic=Invoke-WorkforceMemoryDiagnostic -DataRoot $data -WorkRoot $work -MappingRoot $mappingRoot -AccountClassificationConfigPath $accountClassificationPath
        WriteLog -Message "Workforce diagnostic: exit=$($diagnostic.ExitCode); last stage=$($diagnostic.LastStage); samples=$($diagnostic.Samples); sample errors=$($diagnostic.SampleErrors); sampled peak private MiB=$([math]::Round($diagnostic.SampledPeakPrivateBytes/1MB,1)); private logs=$($diagnostic.RunRoot). No prepared publication." -Level INFO
        if ($diagnostic.SampleErrors -gt 0) { WriteLog -Message 'Some memory measurements were unavailable. Inspect memory.csv; this diagnostic is not a complete memory profile.' -Level WARNING }
        if ($diagnostic.ExitCode -ne 0) { throw "Workforce diagnostic worker failed (exit $($diagnostic.ExitCode)); last stage: $($diagnostic.LastStage). Logs: $($diagnostic.RunRoot)" }
        return
    }
    $phase='Prepare and validate'
    $params = @{
        DataRoot=$data;OutputRoot=$output;WorkRoot=$work;TenantKey=$effective.TenantKey
        AccountClassificationConfigPath=$accountClassificationPath
        MaxSourceAgeHours=[int](ConfigValue 'PreparedMaxSourceAgeHours')
        AgeOverrides=(ConvertTo-PreparedAgeOverrides (ConfigValue 'PreparedSourceAgeOverrides'))
        AllowLegacyTenantless=[bool](ConfigValue 'PreparedAllowLegacyTenantless')
        AllowEmptyTables=[string[]](ConfigValue 'PreparedAllowEmptyTables')
        ValidateOnly=[bool]$ValidateOnly
        MappingRoot=$mappingRoot
    }
    $result = Invoke-PreparedEvidencePipeline @params
    if (-not $ValidateOnly) { WriteLog -Message "Local batch published and validated: $($result.BatchId); CSV files=$($result.Files); path=$($result.BatchPath). Cloud transfer is a separate step." -Level INFO }
    if (-not $ValidateOnly -and $batchUploadEnabled) {
        $phase='SharePoint batch transfer'
        $transfer=Send-PreparedEvidenceBatch @transferParameters -ExpectedBatchId $result.BatchId
        WriteLog -Message "SharePoint batch and pointer verified: $($transfer.BatchId); files=$($transfer.VerifiedFiles); audit=$($transfer.AuditPath)." -Level INFO
    }
    $recap = if ($ValidateOnly) { "Source preflight: $($result.SourceFiles) files; $($result.Identity.CsvFiles) CSVs and $($result.Identity.Rows) rows checked; no generation/publication." } else { "$($result.Files) prepared CSVs; observed history retained; batch $($result.BatchId)." }
    WriteLog -Message $recap -Level INFO
    if (-not ($Offline -or $ValidateOnly -or $WorkforceDiagnostic)) {
        Send-SmartM365TeamsNotification -Title $scriptName -Message 'Preparation completed.' -Level SUCCESS -Channel Infos -ResultSummary $recap -Facts @{Tenant=$effective.TenantKey;Output=$output} | Out-Null
    }
} catch {
    $failure=$_
    if ($coreLoaded) {
        WriteLog -Message "$phase failed: $($_.Exception.Message)" -Level ERROR
        if (-not ($Offline -or $ValidateOnly -or $WorkforceDiagnostic)) {
            $message=$_.Exception.Message
            $help='https://chatgpt.com/?q='+[uri]::EscapeDataString("Explain SmartM365 preparation failure in phase ${phase}: $message")
            Send-SmartM365TeamsNotification -Title $scriptName -Message $message -Level ERROR -Channel Alerts -HelpUrl $help -Facts @{Tenant=$effective.TenantKey;Phase=$phase;Output=$output;Log=$global:LogTextFile;InnerException=[string]$_.Exception.InnerException} | Out-Null
            SendEmailHtmlReport -Subject "$scriptName failed" -BodyHtml ([Net.WebUtility]::HtmlEncode("$phase : $message"))
        }
    } else { Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date),$_.Exception.Message) }
} finally {
    if ($mappingRoot) { try { Remove-PreparedMappingWorkbooks -WorkRoot $work -MappingRoot $mappingRoot } catch { Write-Warning "Downloaded mapping cleanup incomplete: $($_.Exception.Message)" } }
    if ($transcriptStarted) { try { Stop-Transcript | Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } catch { Write-Warning "Transcript finalization failed: $($_.Exception.Message)" } }
    if ($coreLoaded) { Complete-SmartM365ExecutionContext -Status $(if($failure){'Failed'}else{'Success'}) -ErrorRecord $failure -FailureStage $(if($failure){$phase}else{''}) }
    else { Write-SmartM365CompletionBanner -Status 'Failed' }
}
if ($failure) { exit 1 }
