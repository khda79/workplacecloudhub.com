<#
.SYNOPSIS
Producer-owned, current-only SmartInventory source receipts. Dot-sourced helper.
.VERSION
1.2.1
.NOTES
Compatible with Windows PowerShell 5.1. No APIs, history or report refresh.
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
    $temporary=$path+'.'+[guid]::NewGuid().ToString('N')+'.tmp'
    try{
        [IO.File]::WriteAllText($temporary,($Document | ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
        if(Test-Path -LiteralPath $path){[IO.File]::Replace($temporary,$path,[Management.Automation.Language.NullString]::Value,$true)}else{[IO.File]::Move($temporary,$path)}
    }finally{if(Test-Path -LiteralPath $temporary){Remove-Item -LiteralPath $temporary -Force}}
}

# The current proof is immutable during collection; only canonical replacement fences it.
function Write-SmartM365SourceRunState {
    param($Context,[string]$Status,[string]$ErrorMessage='',[int]$Errors=0,[bool]$PartialInventory=$false)
    $document=@{Owner='SmartInventory-SourceRun';ContractVersion='1.0';Status=$Status;Errors=$Errors;Error=$ErrorMessage;PublicationStarted=[bool]$Context.PublicationStarted;UnqualifiedPublication=[bool]$Context.UnqualifiedPublication;Files=@();IsPartialInventory=($PartialInventory -or $Status -ne 'Completed');ConsumerScopeQualified=$false;Scope=$Context.Scope}
    foreach($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer','ScriptVersion','RunId','StartedAtUtc')){$document[$field]=$Context[$field]}
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
    if([IO.Path]::GetDirectoryName($full).TrimEnd([char[]]'\/') -ine $context.SourceRoot){return}
    Start-SmartM365SourcePublication -Path $full
    $context.PublishedHashes[$full]=(Get-FileHash -LiteralPath $full -Algorithm SHA256 -ErrorAction Stop).Hash
    if(-not $global:csvGeneratedPaths){$global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)}
    if($global:csvGeneratedPaths -is [array]){$global:csvGeneratedPaths=@($global:csvGeneratedPaths)+@($full)}else{[void]$global:csvGeneratedPaths.Add($full)}
}

function Complete-SmartM365SourceReceipt {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Status,[int]$ErrorCount=0,[switch]$PartialInventory)
    Complete-SmartM365CmdbSourceReceipt -Status $Status -ErrorCount $ErrorCount -PartialInventory:$PartialInventory
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
    param([Parameter(Mandatory)][string]$Status,[int]$ErrorCount=0,[switch]$PartialInventory)
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
                    $path=Join-Path $context.SourceRoot $name
                    $record=Get-SmartM365CmdbCsvReceipt -Path $path -TenantKey $context.TenantKey -Identity $context
                    if($context.PublishedHashes.ContainsKey($path) -and $context.PublishedHashes[$path] -cne $record.SHA256){throw 'CSV changed after current-run publication.'}
                    foreach($field in @('Producer','ScriptVersion','RunId','StartedAtUtc')){$record[$field]=$context[$field]}
                    $record.CompletedAtUtc=[datetime]::UtcNow.ToString('o');$record.Status='Success'
                    $record.Errors=0;$record.IsPartialInventory=$(if($PartialInventory){$true}elseif($context.ConfiguredOutputs){$null}else{$context.ContainsKey('DomainCoverage') -and $null -ne $context.DomainCoverage});$record.Scope=$context.Scope
                    if($context.ContainsKey('DomainCoverage') -and $context.DomainCoverage){$record.DomainCoverage=$context.DomainCoverage}
                    $record.Required=$context.ExpectedFiles -contains $name;$files+=$record
                }
                foreach($record in $files){if((Get-FileHash -LiteralPath (Join-Path $context.SourceRoot $record.File) -Algorithm SHA256).Hash -ne $record.SHA256){throw 'CSV changed while completing the producer batch.'}}
            }catch{$complete=$false;$reason=$_.Exception.Message;$files=@()}
        }else{$reason='Producer failed, restricted/unavailable scope, or full-scope declaration missing.'}
        $document=@{}
        foreach($field in @('Owner','ContractVersion','TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer','ScriptVersion','RunId','StartedAtUtc')){$document[$field]=$context[$field]}
        $document.CompletedAtUtc=[datetime]::UtcNow.ToString('o')
        $document.Status=if($complete){'Completed'}else{'Failed'}
        $document.IsPartialInventory=if(-not $complete -or $PartialInventory){$true}elseif($context.ConfiguredOutputs){$null}else{$context.ContainsKey('DomainCoverage') -and $null -ne $context.DomainCoverage}
        if($context.ContainsKey('DomainCoverage') -and $context.DomainCoverage){$document.DomainCoverage=$context.DomainCoverage}
        $document.ScopeQualification=if($context.ConfiguredOutputs){'ConfiguredOutputsOnly'}else{'ConsumerScope'}
        $document.RequiredFiles=@($context.ExpectedFiles)
        $document.ScopeParameters=$context.ScopeParameters
        $document.OptionalFiles=@($context.OptionalFiles | Where-Object {$_} | ForEach-Object {@{File=$_;Published=(@($files | ForEach-Object {$_.File}) -contains $_)}})
        $document.FullInventoryQualified=$false
        $document.ConsumerScopeQualified=$complete -and -not $context.ConfiguredOutputs
        $document.Errors=if($complete){0}else{[math]::Max(1,$ErrorCount)};$document.Scope=$context.Scope
        $document.Qualifications=@(if($context.ContainsKey('Qualifications')){$context.Qualifications})
        if($complete -and $PartialInventory){$document.Qualifications+=@('PartialCoverage')}
        $document.Error=$reason;$document.Files=@($files)
        if($complete){
            $document.PublicationProtocol=1
            Write-SmartM365CmdbReceiptDocument $context $document
            $context.UnqualifiedPublication=$false
            $script:SmartM365SourceReceiptUploadPaths=@($context.ReceiptPath)
        }
        Write-SmartM365SourceRunState -Context $context -Status $document.Status -ErrorMessage $reason -Errors $document.Errors -PartialInventory ([bool]($document.IsPartialInventory -eq $true))
        $script:SmartM365SourceReceiptUploadPaths+=@($context.RunPath)
        if(-not $complete){WriteLog -Message "Source qualification rejected for '$($context.Producer)': $reason Native CSVs remain preserved." -Level WARNING}
        if($complete){return $context.ReceiptPath}else{return $context.RunPath}
    }finally{if($context.Lock){$context.Lock.Dispose()};$script:SmartM365CmdbSourceContext=$null}
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCb1IcBKkq7C6PN
# o/E77FF6sN1hiHyQD2AI6IkgqHa4S6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBqZQbGMq7SfTp7WCBFj+2r
# Y0cqehNheSAM8oFnono7NTANBgkqhkiG9w0BAQEFAASCAYCQ8+zt2f01Ut1oCZtD
# zT9CPQ2u9dKu1vK7aAqOUHlDPcTmd6FKqEc5nRiImMpqeKHsCpHYp9jk4OdPk8FV
# cb31n/U0dGvAqc1r4WzXTqv4Oxkx2x40DyQdgc4wiZRkO9aeAS35Vz59+DTe7xN2
# w5QZmU/lEBUf7+7ygL9dGBadKpxNYQS42J9GpJi9JDTXbZeefiZp2SeH7pmafvrj
# Wd1kxCJeeghxmx7GrzxHdX1+KRGDgH8tTzVf93p/J4U4O30vndQJrTX4SdeV6G1e
# JZEHKFCBi0S8Kr1LgxGiv/94aHMFc2bShYGXaAsPJz6GMelrOAR4eyf1c9WEFZib
# lzBHHW8P3Hm1VRvnEDySuwhb2gvLw+nss4sNKdsuFWvqK8+6Mn+XgPm+W0xS2i0q
# vr6QrRYsvfydZYycc7rM3FEV4ZPMAZ8Lxq7G4zcagJZ5Gm0HwYFQRqJuvkDWD0TS
# gTbYxZN6Rqn56uJvS7aIANBtw4wjWC0nOvPHvoAkbZ6KX8s=
# SIG # End signature block
