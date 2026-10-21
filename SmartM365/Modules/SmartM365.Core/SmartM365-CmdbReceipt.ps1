<#
.SYNOPSIS
Producer-owned, current-only SmartInventory source receipts. Dot-sourced helper.
.VERSION
1.4.2
.NOTES
Compatible with Windows PowerShell 5.1. Optional daily capture maintenance is
delegated to Core only for explicitly opted-in producers; no collection or report refresh.
Opted-in producers stage their DATA-LAST CSVs (and declared artifacts) under
DATA-LAST\.staging\<producer>\<run> and promote them, receipt last, only when the run
completes; DATA-LAST is untouched otherwise.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Core identity and current-run CSV publication contract; this helper does not change these globals.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Internal producer transaction: its start and scope phases must not be individually skipped while publishing completion proof.')]
param()
$script:SmartM365CmdbSourceContext=$null
$script:SmartM365SourceReceiptUploadPaths=@()
$script:SmartM365CmdbRegistryPath=Join-Path $PSScriptRoot 'SmartM365-CmdbSources.json.txt'
if($PSScriptRoot -like '*WindowsPowerShell5'){$script:SmartM365CmdbRegistryPath=Join-Path $PSScriptRoot '../../SmartM365-CmdbSources.json.txt'}

function Write-SmartM365CmdbReceiptDocument {
    param([Parameter(Mandatory)]$Context,[Parameter(Mandatory)]$Document,[switch]$RunState)
    $path=if($RunState){$Context.RunPath}else{$Context.ReceiptPath}
    if((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Linked source metadata is not accepted.'}
    if(-not $RunState -and $Context.Started){
        $currentHash=if(Test-Path -LiteralPath $path){(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}else{''}
        if($currentHash -cne $Context.PreviousProofHash){throw 'Source completion proof changed during collection.'}
    }
    if(Test-Path -LiteralPath $path){
        if((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked CMDB receipt is not accepted.'}
        $existing=[IO.File]::ReadAllText($path) | ConvertFrom-Json
        $knownVersion=if($RunState){$existing.Owner -eq 'SmartInventory-SourceRun' -and $existing.ContractVersion -eq '1.0'}else{($existing.Owner -eq 'SmartInventory-CmdbSourceReceipt' -and $existing.ContractVersion -eq '1.1') -or ($existing.Owner -eq 'SmartInventory-SourceReceipt' -and $existing.ContractVersion -eq '1.2')}
        if(-not $knownVersion -or $existing.TenantKey -ne $Context.TenantKey){throw 'Source receipt is not owned by the active tenant or has an unsupported version.'}
        foreach($field in @('OrganizationKey','EnvironmentKey','TenantId','Producer')){if($existing.$field -cne $Context[$field]){throw 'CMDB receipt identity changed.'}}
        if($RunState -and $Context.Started -and $existing.RunId -ne $Context.RunId){throw 'Source run identity changed.'}
    }
    $temporary=[IO.Path]::Combine([IO.Path]::GetDirectoryName($path), ('~sm' + [guid]::NewGuid().ToString('N').Substring(0, 12) + '.tmp'))
    try{
        [IO.File]::WriteAllText($temporary,($Document | ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
        if(Test-Path -LiteralPath $path){[IO.File]::Replace($temporary,$path,[Management.Automation.Language.NullString]::Value,$true)}else{[IO.File]::Move($temporary,$path)}
    }finally{if(Test-Path -LiteralPath $temporary){Remove-Item -LiteralPath $temporary -Force}}
}

# The current proof is immutable during collection; only canonical replacement fences it.
function Write-SmartM365SourceRunState {
    param($Context,[string]$Status,[string]$ErrorMessage='',[int]$Errors=0)
    $document=@{Owner='SmartInventory-SourceRun';ContractVersion='1.0';Status=$Status;Errors=$Errors;Error=$ErrorMessage;PublicationStarted=[bool]$Context.PublicationStarted;UnqualifiedPublication=[bool]$Context.UnqualifiedPublication;Files=@();IsPartialInventory=($Status -ne 'Completed');ConsumerScopeQualified=$false;Scope=$Context.Scope}
    # Read by the orchestrator: a staged run leaves DATA-LAST untouched until Publishing.
    if($Context.ContainsKey('StagedPublication') -and $Context.StagedPublication){$document.StagedPublication=$true}
    foreach($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer','ScriptVersion','RunId','StartedAtUtc')){$document[$field]=$Context[$field]}
    if($Context.ContainsKey('RunNotes') -and $Context.RunNotes.Count){$document.Notes=@($Context.RunNotes)}
    $document.UpdatedAtUtc=[datetime]::UtcNow.ToString('o')
    if($Status -in @('Completed','Failed')){$document.CompletedAtUtc=$document.UpdatedAtUtc}
    Write-SmartM365CmdbReceiptDocument $Context $document -RunState
}

function Assert-SmartM365SourcePublication {
    param([Parameter(Mandatory)][string]$ReceiptPath,[Parameter(Mandatory)]$Proof)
    $runPath=$ReceiptPath -replace '\.current\.json\.txt$','.run.json.txt'
    $protocol=$Proof.PSObject.Properties['PublicationProtocol']
    if($protocol -and (($protocol.Value -isnot [int] -and $protocol.Value -isnot [long]) -or $protocol.Value -ne 1)){throw 'Unsupported source publication protocol.'}
    if(-not $protocol -and -not(Test-Path -LiteralPath $runPath)){return}
    if($runPath -eq $ReceiptPath -or -not(Test-Path -LiteralPath $runPath -PathType Leaf)){throw 'Source run state is missing.'}
    if((Get-Item -LiteralPath $runPath).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked source run state is not accepted.'}
    $run=[IO.File]::ReadAllText($runPath) | ConvertFrom-Json
    if($run.Owner -cne 'SmartInventory-SourceRun' -or $run.ContractVersion -cne '1.0' -or $run.Status -notin @('Collecting','Publishing','Completed','Failed') -or $run.PublicationStarted -isnot [bool] -or $run.UnqualifiedPublication -isnot [bool] -or -not $run.RunId){throw 'Invalid source run state.'}
    foreach($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer')){if($run.$field -cne $Proof.$field){throw 'Source run identity differs from completion proof.'}}
    if($run.Status -eq 'Publishing' -or $run.UnqualifiedPublication){throw 'Canonical source publication is in progress or remains unqualified.'}
}

function Start-SmartM365SourcePublication {
    param([Parameter(Mandatory)][string]$Path)
    $context=$script:SmartM365CmdbSourceContext
    if(-not $context -or [IO.Path]::GetExtension($Path) -ine '.csv'){return}
    $full=[IO.Path]::GetFullPath($Path)
    if([IO.Path]::GetDirectoryName($full).TrimEnd([char[]]'\/') -ine $context.SourceRoot){return}
    # A direct DATA-LAST write voids the staging guarantee for the rest of this run.
    if($context.StagedPublication){
        $context.StagedPublication=$false
        WriteLog -Message "Staged source publication bypassed by a direct DATA-LAST write: $full" -Level WARNING
    }
    if(-not $context.PublicationStarted){
        $context.PublicationStarted=$true;$context.UnqualifiedPublication=$true
        try{Write-SmartM365SourceRunState $context Publishing}
        catch{$context.PublicationStarted=$false;throw}
    }
}

function Start-SmartM365CmdbSourceReceipt {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ScriptPath,[string]$SourceRootPath,[switch]$ReadOnly,[switch]$ConfiguredOutputs,[hashtable]$ScopeParameters=@{})
    if($script:SmartM365CmdbSourceContext){throw 'A CMDB source receipt is already active in this module.'}
    if (Get-Command Test-SmartM365DailyCapturePath -ErrorAction SilentlyContinue) { $script:SmartM365DailyCapture=$null }
    $script:SmartM365SourceReceiptUploadPaths=@()
    if($ReadOnly -or (Test-SmartM365MaxItemsMode)){return}
    $registryPath=if($ConfiguredOutputs){Join-Path (Split-Path $script:SmartM365CmdbRegistryPath) 'SmartM365-SourceReceipts.json.txt'}else{$script:SmartM365CmdbRegistryPath}
    $registry=[IO.File]::ReadAllText($registryPath) | ConvertFrom-Json
    $name=[IO.Path]::GetFileName($ScriptPath)
    $producers=@($registry.Producers | Where-Object Script -eq $name)
    if($producers.Count -ne 1){throw 'CMDB producer is not registered exactly once.'}
    if($ConfiguredOutputs){$global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)}
    if(-not $SourceRootPath){$SourceRootPath=[string](Get-SmartM365EffectiveModuleGlobalConfig).LatestCsvFolderPath}
    if(-not [IO.Path]::IsPathRooted($SourceRootPath) -or $SourceRootPath -match '\{\{|__USE_GLOBAL__'){throw 'Source receipt root must be an absolute resolved path.'}
    $effectiveScope=@{}
    foreach($key in $ScopeParameters.Keys){
        $value=$ScopeParameters[$key]
        $effectiveScope[$key]=if($value -is [Management.Automation.SwitchParameter]){[bool]$value}else{$value}
    }
    $root=[IO.Path]::GetFullPath($SourceRootPath).TrimEnd([char[]]'\/')
    if(-not(Test-Path -LiteralPath $root -PathType Container)){New-Item -ItemType Directory -Path $root | Out-Null}
    $path=Join-Path $root $producers[0].Receipt
    $context=@{
        Owner='SmartInventory-SourceReceipt';ContractVersion='1.2'
        TenantKey=[string]$global:SmartM365TenantKey;OrganizationKey=[string]$global:SmartM365OrganizationKey
        EnvironmentKey=[string]$global:SmartM365EnvironmentKey;TenantId=[string]$global:SmartM365TenantId
        Producer=$name;ScriptVersion=Get-SmartM365ScriptVersionFromFile -Path $ScriptPath
        RunId=[guid]::NewGuid().ToString('N');StartedAtUtc=[datetime]::UtcNow.ToString('o')
        SourceRoot=$root;ReceiptPath=$path;RunPath=($path -replace '\.current\.json\.txt$','.run.json.txt');ExpectedFiles=@($producers[0].Files)
        PublicationStarted=$false;UnqualifiedPublication=$false;PreviousProofHash=''
        RequiredScope=[string]$producers[0].Scope;ScopeQualified=[bool]$ConfiguredOutputs;Scope=$(if($ConfiguredOutputs){'ConfiguredOutputs'}else{'NotDeclared'});Started=$false;Lock=$null
        ConfiguredOutputs=[bool]$ConfiguredOutputs;PublishedHashes=@{};OptionalFiles=@(if($producers[0].PSObject.Properties['OptionalFiles']){$producers[0].OptionalFiles})
        ScopeParameters=$effectiveScope
        StagedPublication=$false;StagingRoot='';StagedFiles=@{};StagedArtifacts=@{};StagedActions=New-Object 'Collections.Generic.List[object]'
    }
    foreach($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','ScriptVersion')){if(-not $context[$field]){throw "Missing producer identity: $field"}}
    try{
        if((Test-Path -LiteralPath ($path+'.collection.lock')) -and ((Get-Item -LiteralPath ($path+'.collection.lock')).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Linked producer lock is not accepted.'}
        $context.Lock=[IO.File]::Open($path+'.collection.lock',[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        # Validate ownership without rewriting or reconstructing previous evidence.
        if(Test-Path -LiteralPath $path){
            if((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked source receipt is not accepted.'}
            $context.PreviousProofHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
            $previous=[IO.File]::ReadAllText($path) | ConvertFrom-Json
            if(-not (($previous.Owner -eq 'SmartInventory-CmdbSourceReceipt' -and $previous.ContractVersion -eq '1.1') -or ($previous.Owner -eq 'SmartInventory-SourceReceipt' -and $previous.ContractVersion -eq '1.2'))){throw 'Foreign source receipt ownership.'}
            foreach($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer')){if($previous.$field -cne $context[$field]){throw 'Source receipt identity differs from active producer.'}}
        }
        if(Test-Path -LiteralPath $context.RunPath){
            if((Get-Item -LiteralPath $context.RunPath).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked source run state is not accepted.'}
            $previousRun=[IO.File]::ReadAllText($context.RunPath) | ConvertFrom-Json
            if($previousRun.UnqualifiedPublication -isnot [bool] -or $previousRun.PublicationStarted -isnot [bool] -or $previousRun.Status -notin @('Collecting','Publishing','Completed','Failed')){throw 'Invalid previous source run state.'}
            $context.UnqualifiedPublication=$previousRun.UnqualifiedPublication -or $previousRun.Status -eq 'Publishing'
        }
        $script:SmartM365SourceReceiptUploadPaths=@()
        Write-SmartM365SourceRunState $context Collecting
        $context.Started=$true;$script:SmartM365CmdbSourceContext=$context
    }catch{if($context.Lock){$context.Lock.Dispose()};throw}
}

# Public generic entry point. CMDB wrappers retain their independently declared scope.
function Start-SmartM365SourceReceipt {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ScriptPath,[string]$SourceRootPath,[switch]$ReadOnly,[hashtable]$ScopeParameters=@{})
    Start-SmartM365CmdbSourceReceipt -ScriptPath $ScriptPath -SourceRootPath $SourceRootPath -ReadOnly:$ReadOnly -ConfiguredOutputs -ScopeParameters $ScopeParameters
}

# Register only a successful write/copy in the exact current source root, never scan DATA-LAST.
function Register-SmartM365SourceCsv {
    param([Parameter(Mandatory)][string]$Path)
    $context=$script:SmartM365CmdbSourceContext
    if(-not $context -or [IO.Path]::GetExtension($Path) -ine '.csv'){return}
    $full=[IO.Path]::GetFullPath($Path)
    $directory=[IO.Path]::GetDirectoryName($full).TrimEnd([char[]]'\/')
    if($context.StagedPublication -and $directory -ieq $context.StagingRoot){
        # Recorded under its canonical name; the staged copy is promoted at completion.
        $canonical=Join-Path $context.SourceRoot ([IO.Path]::GetFileName($full))
        $context.StagedFiles[$canonical]=$full
        $context.PublishedHashes[$canonical]=(Get-FileHash -LiteralPath $full -Algorithm SHA256 -ErrorAction Stop).Hash
        if(-not $global:csvGeneratedPaths){$global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)}
        if($global:csvGeneratedPaths -is [array]){$global:csvGeneratedPaths=@($global:csvGeneratedPaths)+@($canonical)}else{[void]$global:csvGeneratedPaths.Add($canonical)}
        return
    }
    if($directory -ine $context.SourceRoot){return}
    Start-SmartM365SourcePublication -Path $full
    $context.PublishedHashes[$full]=(Get-FileHash -LiteralPath $full -Algorithm SHA256 -ErrorAction Stop).Hash
    if(-not $global:csvGeneratedPaths){$global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)}
    if($global:csvGeneratedPaths -is [array]){$global:csvGeneratedPaths=@($global:csvGeneratedPaths)+@($full)}else{[void]$global:csvGeneratedPaths.Add($full)}
}

# Opt-in, after the receipt starts: this run writes its DATA-LAST CSVs to a private staging folder.
function Enable-SmartM365StagedSourcePublication {
    [CmdletBinding()]
    param()
    $context=$script:SmartM365CmdbSourceContext
    if(-not $context){return}
    if($context.PublicationStarted){throw 'Staged source publication must be enabled before the first DATA-LAST write.'}
    $producerRoot=Join-Path (Join-Path $context.SourceRoot '.staging') ([IO.Path]::GetFileNameWithoutExtension($context.Producer))
    foreach($folder in @((Join-Path $context.SourceRoot '.staging'),$producerRoot)){
        if((Test-Path -LiteralPath $folder) -and ((Get-Item -LiteralPath $folder -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Linked source staging folder is not accepted.'}
    }
    if(-not(Test-Path -LiteralPath $producerRoot -PathType Container)){New-Item -ItemType Directory -Path $producerRoot -Force | Out-Null}
    # The producer lock is held: every other run folder of this producer is an orphan of a crash.
    foreach($orphan in @(Get-ChildItem -LiteralPath $producerRoot -Force -ErrorAction Stop)){
        if($orphan.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "Linked source staging entry is not accepted: $($orphan.FullName)"}
        Remove-Item -LiteralPath $orphan.FullName -Recurse -Force -ErrorAction Stop
        WriteLog -Message "Orphaned source staging removed: $($orphan.FullName)" -Level INFO
    }
    $context.StagingRoot=[IO.Path]::GetFullPath((Join-Path $producerRoot $context.RunId)).TrimEnd([char[]]'\/')
    New-Item -ItemType Directory -Path $context.StagingRoot | Out-Null
    $context.StagedPublication=$true
    Write-SmartM365SourceRunState $context Collecting
    WriteLog -Message "Staged source publication enabled; DATA-LAST is updated only after a complete run: $($context.StagingRoot)" -Level INFO
}

# Write target for a canonical DATA-LAST CSV: its staged copy while staging is active, else unchanged.
# -Artifact declares a non-CSV DATA-LAST file (report snapshot) promoted with the CSVs, before the receipt.
function Resolve-SmartM365SourcePublicationPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path,[switch]$Artifact)
    $context=$script:SmartM365CmdbSourceContext
    if(-not $context -or -not $context.StagedPublication -or ([IO.Path]::GetExtension($Path) -ine '.csv' -and -not $Artifact)){return $Path}
    $full=[IO.Path]::GetFullPath($Path)
    if([IO.Path]::GetDirectoryName($full).TrimEnd([char[]]'\/') -ine $context.SourceRoot){return $Path}
    $staged=Join-Path $context.StagingRoot ([IO.Path]::GetFileName($full))
    if($Artifact){
        if([IO.Path]::GetExtension($full) -ieq '.csv'){throw 'A CSV is published through the receipt, not as an artifact.'}
        $context.StagedArtifacts[$full]=$staged
    }
    return $staged
}

# Follow-up of a staged CSV (SharePoint copy, WeeklyHistory, producer callback), run only after its promotion.
function Add-SmartM365StagedPublicationAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path,[string]$TimestampedPath='',[switch]$SharePointUpload,[switch]$WeeklyHistory,[scriptblock]$Callback,[object[]]$ArgumentList=@())
    $context=$script:SmartM365CmdbSourceContext
    if(-not $context -or -not $context.StagedPublication){throw 'No staged source publication is active.'}
    $context.StagedActions.Add([pscustomobject]@{Path=[IO.Path]::GetFullPath($Path);TimestampedPath=$TimestampedPath;SharePointUpload=[bool]$SharePointUpload;WeeklyHistory=[bool]$WeeklyHistory;Callback=$Callback;ArgumentList=@($ArgumentList)})
}

function Remove-SmartM365SourceStaging {
    param([Parameter(Mandatory)]$Context)
    if(-not $Context.StagingRoot -or -not(Test-Path -LiteralPath $Context.StagingRoot)){return}
    try{Remove-Item -LiteralPath $Context.StagingRoot -Recurse -Force -ErrorAction Stop}
    catch{WriteLog -Message ("Source staging could not be removed; the next run removes it: {0}" -f $_.Exception.Message) -Level WARNING}
}

# Moves every staged CSV over its canonical file. The caller writes the receipt afterwards.
function Publish-SmartM365StagedSourceSet {
    param([Parameter(Mandatory)]$Context,[Parameter(Mandatory)][object[]]$Records)
    $Context.PublicationStarted=$true;$Context.UnqualifiedPublication=$true
    Write-SmartM365SourceRunState $Context Publishing
    foreach($record in $Records){
        $canonical=Join-Path $Context.SourceRoot $record.File
        if(-not $Context.StagedFiles.ContainsKey($canonical)){continue}
        $staged=$Context.StagedFiles[$canonical]
        if((Test-Path -LiteralPath $canonical) -and ((Get-Item -LiteralPath $canonical).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw "Linked canonical CSV is not accepted: $canonical"}
        if(Test-Path -LiteralPath $canonical){[IO.File]::Replace($staged,$canonical,[Management.Automation.Language.NullString]::Value,$true)}else{[IO.File]::Move($staged,$canonical)}
        if((Get-FileHash -LiteralPath $canonical -Algorithm SHA256).Hash -cne $record.SHA256){throw "Promoted CSV differs from its staged copy: $canonical"}
        WriteLog -Message "CSV latest copy written to: $canonical (promoted from source staging)" -Level INFO
    }
    # Declared artifacts follow the CSVs they describe; an artifact the run did not write keeps its previous copy.
    foreach($canonical in @($Context.StagedArtifacts.Keys | Sort-Object)){
        $staged=$Context.StagedArtifacts[$canonical]
        if(-not(Test-Path -LiteralPath $staged -PathType Leaf)){continue}
        if((Test-Path -LiteralPath $canonical) -and ((Get-Item -LiteralPath $canonical).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw "Linked canonical artifact is not accepted: $canonical"}
        $hash=(Get-FileHash -LiteralPath $staged -Algorithm SHA256).Hash
        if(Test-Path -LiteralPath $canonical){[IO.File]::Replace($staged,$canonical,[Management.Automation.Language.NullString]::Value,$true)}else{[IO.File]::Move($staged,$canonical)}
        if((Get-FileHash -LiteralPath $canonical -Algorithm SHA256).Hash -cne $hash){throw "Promoted artifact differs from its staged copy: $canonical"}
        WriteLog -Message "Artifact latest copy written to: $canonical (promoted from source staging)" -Level INFO
    }
}

function Invoke-SmartM365StagedPublicationFollowUp {
    param([Parameter(Mandatory)]$Context)
    foreach($action in $Context.StagedActions.ToArray()){
        try{
            if($action.SharePointUpload -and (Test-Path -LiteralPath $action.Path)){Invoke-SmartM365SharePointCsvUpload -LocalFilePath $action.Path | Out-Null}
            if($action.WeeklyHistory){Invoke-SmartM365WeeklyInventoryHistoryForCsv -SourceFiles @($action.Path) -TimestampedPath $action.TimestampedPath}
            if($action.Callback){$arguments=@($action.ArgumentList);& $action.Callback @arguments}
        }catch{WriteLog -Message ("Post-publication step failed for {0}; DATA-LAST and its receipt are published: {1}" -f $action.Path,$_.Exception.Message) -Level WARNING}
    }
}

# Read-only view of the active run: lets a producer derive outputs from its own current publication.
function Get-SmartM365CurrentSourcePublication {
    [CmdletBinding()]
    param()
    $context=$script:SmartM365CmdbSourceContext
    if(-not $context){return $null}
    $hashes=[hashtable]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($key in $context.PublishedHashes.Keys){$hashes[$key]=$context.PublishedHashes[$key]}
    return [pscustomobject]@{Producer=[string]$context.Producer;RunId=[string]$context.RunId;StartedAtUtc=[string]$context.StartedAtUtc;SourceRoot=[string]$context.SourceRoot;Hashes=$hashes}
}

# Records a qualification of the active run in its run state, for example an output that was not republished.
function Add-SmartM365SourceRunNote {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)
    $context=$script:SmartM365CmdbSourceContext
    if(-not $context){return}
    if(-not $context.ContainsKey('RunNotes')){$context.RunNotes=[System.Collections.Generic.List[string]]::new()}
    $context.RunNotes.Add($Message)
    Write-SmartM365SourceRunState $context $(if($context.PublicationStarted){'Publishing'}else{'Collecting'})
}

function Complete-SmartM365SourceReceipt {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Status,[int]$ErrorCount=0)
    Complete-SmartM365CmdbSourceReceipt -Status $Status -ErrorCount $ErrorCount
}

function Set-SmartM365CmdbSourceScope {
    [CmdletBinding()]
    param([Parameter(Mandatory)][bool]$CompleteScope,[Parameter(Mandatory)][string]$Scope,[string[]]$Qualifications=@(),[hashtable]$DomainCoverage)
    if(-not $script:SmartM365CmdbSourceContext){return}
    $context=$script:SmartM365CmdbSourceContext
    $context.ScopeQualified=$false;$context.Scope=$Scope;$context.Qualifications=@($Qualifications)
    $context.DomainCoverage=$null
    if($DomainCoverage){
        if($CompleteScope -or $context.ConfiguredOutputs -or $context.Producer -cne 'SmartM365-ActiveDirectory-Inventory.ps1' -or $Scope -cne $context.RequiredScope){throw 'Partial domain coverage is restricted to the AD CMDB producer.'}
        $sets=@{}
        foreach($field in @('ExpectedDomains','CollectedDomains','UnavailableDomains','NonBlockingDomainErrors')){
            $values=@($DomainCoverage[$field])
            $set=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            foreach($value in $values){
                if($value -isnot [string] -or $value -notmatch '^[a-zA-Z0-9](?:[a-zA-Z0-9.-]*[a-zA-Z0-9])?$' -or -not $set.Add($value)){throw "Invalid or repeated AD coverage domain: $field"}
            }
            $sets[$field]=$set
        }
        if(-not $sets.ExpectedDomains.Count -or -not $sets.CollectedDomains.Count -or -not $sets.UnavailableDomains.Count -or
           $sets.CollectedDomains.Overlaps($sets.UnavailableDomains) -or -not $sets.UnavailableDomains.IsSubsetOf($sets.NonBlockingDomainErrors)){throw 'Partial AD coverage is empty, overlapping or not explicitly tolerated.'}
        $union=New-Object 'Collections.Generic.HashSet[string]' ($sets.CollectedDomains,[StringComparer]::OrdinalIgnoreCase)
        $union.UnionWith($sets.UnavailableDomains)
        if(-not $union.SetEquals($sets.ExpectedDomains)){throw 'Partial AD coverage does not account for every expected domain.'}
        $coverage=@{Kind='ADDomainCoverage';Reason='NonBlockingDomainErrors';Status='PartialAccepted'}
        foreach($field in $sets.Keys){$coverage[$field]=@($sets[$field] | Sort-Object)}
        $context.DomainCoverage=$coverage;$context.ScopeQualified=$true
    }else{$context.ScopeQualified=$CompleteScope}
}

function Get-SmartM365CmdbCsvReceipt {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string]$TenantKey,[hashtable]$Identity=@{})
    if((Get-Item -LiteralPath $Path).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked CSV source is not accepted.'}
    $before=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
    $parser=$null
    $delimiter=$null
    $columns=@()
    foreach($candidate in @(',', ';')){
        $candidateParser=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path,[Text.UTF8Encoding]::new($false),$true)
        try{
            $candidateParser.SetDelimiters($candidate);$candidateParser.HasFieldsEnclosedInQuotes=$true;$candidateParser.TrimWhiteSpace=$false
            try{$candidateColumns=@($candidateParser.ReadFields())}
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException]{
                if($candidate -eq ','){continue}
                throw
            }
            if($candidateColumns -contains 'TenantKey'){
                $parser=$candidateParser;$columns=$candidateColumns;$delimiter=$candidate
                break
            }
        }finally{if($candidateParser -ne $parser){$candidateParser.Dispose()}}
    }
    if(-not $parser){throw 'CSV TenantKey header missing.'}
    try{
        if(-not $columns.Count -or @($columns | Select-Object -Unique).Count -ne $columns.Count){throw 'Invalid CSV header.'}
        $tenantIndex=[array]::IndexOf($columns,'TenantKey')
        if($tenantIndex -lt 0){throw 'CSV TenantKey header missing.'}
        $identityIndices=@{}
        foreach($field in @('OrganizationKey','EnvironmentKey','TenantId')){if($Identity.ContainsKey($field) -and $columns -contains $field){$identityIndices[$field]=[array]::IndexOf($columns,$field)}}
        [int64]$count=0
        while(-not $parser.EndOfData){
            $fields=@($parser.ReadFields())
            if($fields.Count -ne $columns.Count -or $fields[$tenantIndex] -cne $TenantKey){throw 'Malformed CSV or foreign tenant row.'}
            foreach($field in $identityIndices.Keys){if($fields[$identityIndices[$field]] -cne $Identity[$field]){throw 'CSV tenant identity differs from the producer.'}}
            $count++
        }
    }finally{$parser.Close();$parser.Dispose()}
    if((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -ne $before){throw 'CSV changed while recording acquisition receipt.'}
    return @{File=[IO.Path]::GetFileName($Path);Rows=$count;SHA256=$before;Delimiter=$delimiter}
}

function Complete-SmartM365CmdbSourceReceipt {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Status,[int]$ErrorCount=0)
    if (Get-Command Get-SmartM365InventoryCompletionStatus -ErrorAction SilentlyContinue) {
        $Status = Get-SmartM365InventoryCompletionStatus -Status $Status -ErrorRecord $null
    }
    $context=$script:SmartM365CmdbSourceContext
    if(-not $context){return}
    try{
        $complete=$Status -in @('Success','Completed','CompletedWithWarnings') -and $ErrorCount -eq 0 -and $context.ScopeQualified -and $context.Scope -ceq $context.RequiredScope
        $files=@();$reason=''
        if($complete){
            try{
                $published=@($global:csvGeneratedPaths | Where-Object {
                    [IO.Path]::GetExtension($_) -ieq '.csv' -and [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($_)).TrimEnd([char[]]'\/') -ieq $context.SourceRoot
                } | ForEach-Object {[IO.Path]::GetFileName($_)} | Sort-Object -Unique)
                if($published.Count -eq 0){throw 'No canonical CSV was published in this run.'}
                foreach($name in $context.ExpectedFiles){if($published -notcontains $name){throw "Required canonical CSV was not published in this run: $name"}}
                foreach($name in $published){
                    $canonical=Join-Path $context.SourceRoot $name
                    $path=if($context.StagedFiles.ContainsKey($canonical)){$context.StagedFiles[$canonical]}else{$canonical}
                    $record=Get-SmartM365CmdbCsvReceipt -Path $path -TenantKey $context.TenantKey -Identity $context
                    $record.File=$name
                    if($context.PublishedHashes.ContainsKey($canonical) -and $context.PublishedHashes[$canonical] -cne $record.SHA256){throw 'CSV changed after current-run publication.'}
                    foreach($field in @('Producer','ScriptVersion','RunId','StartedAtUtc')){$record[$field]=$context[$field]}
                    $record.CompletedAtUtc=[datetime]::UtcNow.ToString('o');$record.Status='Success'
                    $record.Errors=0;$record.IsPartialInventory=$(if($context.ConfiguredOutputs){$null}else{$context.ContainsKey('DomainCoverage') -and $null -ne $context.DomainCoverage});$record.Scope=$context.Scope
                    if($context.ContainsKey('DomainCoverage') -and $context.DomainCoverage){$record.DomainCoverage=$context.DomainCoverage}
                    $record.Required=$context.ExpectedFiles -contains $name;$files+=$record
                }
                foreach($record in $files){
                    $canonical=Join-Path $context.SourceRoot $record.File
                    $current=if($context.StagedFiles.ContainsKey($canonical)){$context.StagedFiles[$canonical]}else{$canonical}
                    if((Get-FileHash -LiteralPath $current -Algorithm SHA256).Hash -ne $record.SHA256){throw 'CSV changed while completing the producer batch.'}
                }
            }catch{$complete=$false;$reason=$_.Exception.Message;$files=@()}
            # Staged CSVs replace DATA-LAST only now, under the Publishing fence; the receipt follows them.
            if($complete -and ($context.StagedFiles.Count -or $context.StagedArtifacts.Count)){
                try{Publish-SmartM365StagedSourceSet -Context $context -Records $files}
                catch{$complete=$false;$reason='Staged source promotion failed: '+$_.Exception.Message;$files=@()}
            }
        }else{$reason=if($Status -eq 'Cancelled'){'Producer cancelled by operator before completion.'}else{'Producer failed, restricted/unavailable scope, or full-scope declaration missing.'}}
        $document=@{}
        foreach($field in @('Owner','ContractVersion','TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer','ScriptVersion','RunId','StartedAtUtc')){$document[$field]=$context[$field]}
        $document.CompletedAtUtc=[datetime]::UtcNow.ToString('o')
        $document.Status=if($complete){'Completed'}else{'Failed'}
        $document.IsPartialInventory=if(-not $complete){$true}elseif($context.ConfiguredOutputs){$null}else{$context.ContainsKey('DomainCoverage') -and $null -ne $context.DomainCoverage}
        if($context.ContainsKey('DomainCoverage') -and $context.DomainCoverage){$document.DomainCoverage=$context.DomainCoverage}
        $document.ScopeQualification=if($context.ConfiguredOutputs){'ConfiguredOutputsOnly'}else{'ConsumerScope'}
        $document.RequiredFiles=@($context.ExpectedFiles)
        $document.ScopeParameters=$context.ScopeParameters
        $document.OptionalFiles=@($context.OptionalFiles | Where-Object {$_} | ForEach-Object {@{File=$_;Published=(@($files | ForEach-Object {$_.File}) -contains $_)}})
        $document.FullInventoryQualified=$false
        $document.ConsumerScopeQualified=$complete -and -not $context.ConfiguredOutputs
        $document.Errors=if($complete){0}else{[math]::Max(1,$ErrorCount)};$document.Scope=$context.Scope
        $document.Qualifications=@(if($context.ContainsKey('Qualifications')){$context.Qualifications})
        $document.Error=$reason;$document.Files=@($files)
        if($complete){
            $document.PublicationProtocol=1
            Write-SmartM365CmdbReceiptDocument $context $document
            $context.UnqualifiedPublication=$false
            $script:SmartM365SourceReceiptUploadPaths=@($context.ReceiptPath)
        }
        Write-SmartM365SourceRunState $context $document.Status $reason $document.Errors
        $script:SmartM365SourceReceiptUploadPaths+=@($context.RunPath)
        if ($complete -and (Get-Command Invoke-SmartM365DailyCaptureRetention -ErrorAction SilentlyContinue)) {
            try { Invoke-SmartM365DailyCaptureRetention -Context $context -Proof ([pscustomobject]$document) }
            catch { WriteLog -Message ('Daily capture retention refused; source CSVs and receipt preserved: {0}' -f $_.Exception.Message) -Level WARNING }
        }
        if($complete -and $context.StagedPublication){Invoke-SmartM365StagedPublicationFollowUp -Context $context}
        if(-not $complete){WriteLog -Message "Source qualification rejected for '$($context.Producer)': $reason Native CSVs remain preserved." -Level WARNING}
        if($complete){return $context.ReceiptPath}else{return $context.RunPath}
    }finally{Remove-SmartM365SourceStaging -Context $context;if($context.Lock){$context.Lock.Dispose()};$script:SmartM365CmdbSourceContext=$null}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA9h5SzNmSgLFR6
# G6LowQ24eLqL8b1mObtkK0H+e5Ue6aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHA81uqeJCJaY0LimYK6pk0zqjUDCiGC/EmtU55iyj40MA0GCSqG
# SIb3DQEBAQUABIIBgC7m4CmAWoQ/ZrRHnlTk9oL7k51UfzIPl2NU7NedR4TUxBOH
# QzrszsyelpsGxrXQ9BiC9eqGkYYrTUuKfOUylfUrxPQ0dQH70p1Dcr0pMaA8mYkP
# ZD3enMli8vO60+dc6lDn4a5P6Qu1Eca5/3k4IvFloGE3scwl+HmyBh2/FglxWaWB
# CHNPB0yTemKSkZNCz9uS3X7NI2Ff0mRDpkmwuPokF6jJBbGA8WQuqsySCcfuVLme
# D/4x2XdYOFOxx6QIlNs0zgZvMAbaUjfQzvRMeGBOE+n3lYh0XB023Haa/xs40TmS
# HqKXVkxnMk/tJLnPKL1USlUM1yJJieZ1mkZ3zLkvDzFt4/4ysHzTi14EdQSnuMQS
# iL2fzXa/ajXGEhhNQQipIuNsloYg8Y8+KYYSkPhadJFVvVjsO8j9FQf+9+tKa7tl
# tkETlpAi2/ZdAxahaMHhY/RpOHzIKARFyA0swL1++DB3iL0Lue/KfCSvRcbasoNa
# n7gQb9XNZIiM+Lvr+qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjExNzE2
# MjJaMC8GCSqGSIb3DQEJBDEiBCCdbH/hcPgyxaxqPcVZV+CG61lBDKXi+QgvZb1l
# x5XAtTANBgkqhkiG9w0BAQEFAASCAgCiKMwlsBu8A9THYmKclpIqWw7H0mjLmfQX
# A5ryKlsavpjqhKKy4XxcvFSwNagTO7JHY9m5h76QU4Adyu+eRMAlwhs/gvUgwRRf
# fw9V/dzZAArXu3Nno978Y5V/MZ5jvFdJS11SwzKy5cLwShY2kyPZfzVYjWUj5FNw
# 8YCyM8cesrXZHaMfyRGR7OaO+kFet6jbp8w+OKKWzMb4QLqbXL3X6stO/QHVgWmD
# E43zN8Oa65IoHbHCi1PfGzUe3l5D5n6sxFWrtcQ4yF70EIiKmBDroP0NDM1fdIZP
# SNViMkBWF7GDt7bsm1cqfF1qE3No4jMttAJOtFZmSDgLFBJIk0gBy07BtB7LSIyb
# gfHrI1IH+oGZ12V6hATyu1nru7LaW28ufOM5cNqJzidGKGuBInAS04YpCDNu1ZIT
# CBxRHYRk9wZIjpYbSl2ol6m8e73lN4+7wvfblFdiYnzmV+yRE0aVfmZHWlAJbkqd
# n1BEdK9QPWYuRT8iPkgrhJ/npT+bJTsBs1Vl7SLIV1XK6vdt9o7qYPKeSda/Eh6r
# bQTuSaT4Jp3uuO3GNkPxmuFGzZFCEqr/0oSqg7ypd42RNciMetAzovvlrzRf52xm
# aEVr05ZtGfD4V9r0WdbB2/ISs2QcMQacr1aVjw0XF8+RixqXwSkap6CCpjdteuns
# 5SDQ7t44+Q==
# SIG # End signature block
