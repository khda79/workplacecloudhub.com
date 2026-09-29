<#
.SYNOPSIS
    Prepare a validated, versioned SmartWorkplaceIntelligence batch after raw collectors.
.VERSION
    0.1.12
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
        $diagnostic=Invoke-WorkforceMemoryDiagnostic -DataRoot $data -WorkRoot $work -MappingRoot $mappingRoot -AccountClassificationConfigPath (Join-Path $smartRoot 'SmartInventory/Config/AccountClassification.psd1')
        WriteLog -Message "Workforce diagnostic: exit=$($diagnostic.ExitCode); last stage=$($diagnostic.LastStage); samples=$($diagnostic.Samples); sample errors=$($diagnostic.SampleErrors); sampled peak private MiB=$([math]::Round($diagnostic.SampledPeakPrivateBytes/1MB,1)); private logs=$($diagnostic.RunRoot). No prepared publication." -Level INFO
        if ($diagnostic.SampleErrors -gt 0) { WriteLog -Message 'Some memory measurements were unavailable. Inspect memory.csv; this diagnostic is not a complete memory profile.' -Level WARNING }
        if ($diagnostic.ExitCode -ne 0) { throw "Workforce diagnostic worker failed (exit $($diagnostic.ExitCode)); last stage: $($diagnostic.LastStage). Logs: $($diagnostic.RunRoot)" }
        return
    }
    $phase='Prepare and validate'
    $params = @{
        DataRoot=$data;OutputRoot=$output;WorkRoot=$work;TenantKey=$effective.TenantKey
        AccountClassificationConfigPath=(Join-Path $smartRoot 'SmartInventory/Config/AccountClassification.psd1')
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBYgOHkKr0nMQjC
# RQ5SvEgoFztt4PJbowJXP917DbS/CqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMw8X8ZuqYheak1khpjVhQ9tWjB597OInxfhUIOf1W2IMA0GCSqG
# SIb3DQEBAQUABIIBgKRXzN1pw5d7l6v/UwiDxCRD+WrJGDaSGwo6AZDURzOpz3gT
# CkFmwWnO+qLWyp4NH9be16PqMOWtIlrFGhPrwapS/LFSroTOBgVpJCIMlpwUOiA9
# vPAIQE6/mZcQHEF8NWu3Q37MPM1Zx1DblvR5hSqm3q/jodNBs8kPrSrESqnkiNio
# wZstllH3uxgbFeo9UPxMSER4tYDIM+5jS4ibEV+0g5xF9MiTIlicKDFBASoWiHkJ
# HyPl3aOAhl+XVzcr2gY0Dwam4SfGHdaPibWa+HWHRex8pUVW6y0MyumRY+o+DzU5
# bDhDPn/EPqtuuIG33c+ERfpszBqK7Mezwmf8KkTTywckG24c/cZEgQFrOG/dwy/b
# KzEM5WY1yFChBEKSDiF7dz/mM1fTESn9/22TYIMXuhmM46cCuMgpxllralvIW/XX
# 1tssSqBpwUWVbGJOCJYwoCe4E6nDXK4t5wA3/iqonO7brETnEJPOKhe7hwWczbNM
# TKBPP4aqapSog9hpf6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU3
# NTRaMC8GCSqGSIb3DQEJBDEiBCBn4Rwh/sQAGkYrK50avMlTWNuZ9Y3gmuqZPbKO
# Fi/GujANBgkqhkiG9w0BAQEFAASCAgA3S1+jXHD235BXcHYGoArZpVjQ37FGyJRK
# ZIAvytEMqgAS7Z99/6aw4Sw2Ac8MXlL++ji+5U3n+WOfy9QjlX00KWq18Xn9jenx
# 23KxeeOo2FzaTlB9EJ7lkJ3/J2XalYMm+g70/fKYOaAfl2+sibm3DeJvPqCPsoX2
# POhz43Fn+ZBE8sSPgWjG+zxHdvwisKCMByv8XkQhhjqFSVTnUBPzqLLgCCj76FnF
# U7O3y/6UQFA/MJvVagw/yIKSEouGUb/xNaG5iKGNYWmLtZsdbgi3Kg32DkEcEQpj
# 2V8Gv0na+MdH2es6mkt5xPiQ3BRG5lngx4i7dDAY7orj+xt9bf2JeEZ4kXUHieQ2
# HCV7Yc6Krs+7v9y/HdnW9Llf80Nt0Yo8GQUpTaTZB6XHQCdPq1di6ZOipG7TdqXw
# qiJITv9xVwUea+5DMop8swri5drr/XMUCOksI0ACvO77MqwUadl0ETnbAcifIsnh
# 4CzUyHG0YWOBjI3PMXtUPt+bDPHiwpuIwkRboqcMwopsM9MZ/KxTrPlB076KPxAR
# A9eD1JXPi/+1zvsT81mc+fVTEO7Cq3KBjhfb0U7cSahNm18xdWEnr4ANFIQc2Dox
# 83W+XGM2jN+0dIdt5lWfl+enkflGDYM7RnsVdVxgUUxK62NW7RThTslc1NALdYiF
# ZOMruPLZzg==
# SIG # End signature block
