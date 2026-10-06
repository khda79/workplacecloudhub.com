<#
.SYNOPSIS
Producer-owned, current-only SmartInventory source receipts. Dot-sourced helper.
.VERSION
1.1.2
.NOTES
Compatible with Windows PowerShell 5.1. No APIs, history or report refresh.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Core identity and current-run CSV publication contract; this helper does not change these globals.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Internal producer transaction: its start and scope phases must not be individually skipped while publishing completion proof.')]
param()
$script:SmartM365CmdbSourceContext=$null
$script:SmartM365CmdbRegistryPath=Join-Path $PSScriptRoot 'SmartM365-CmdbSources.json.txt'
if($PSScriptRoot -like '*WindowsPowerShell5'){$script:SmartM365CmdbRegistryPath=Join-Path $PSScriptRoot '../../SmartM365-CmdbSources.json.txt'}

function Write-SmartM365CmdbReceiptDocument {
    param([Parameter(Mandatory)]$Context,[Parameter(Mandatory)]$Document)
    $path=$Context.ReceiptPath
    if(Test-Path -LiteralPath $path){
        if((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked CMDB receipt is not accepted.'}
        $existing=[IO.File]::ReadAllText($path) | ConvertFrom-Json
        $knownVersion=($existing.Owner -eq 'SmartInventory-CmdbSourceReceipt' -and $existing.ContractVersion -eq '1.1') -or ($existing.Owner -eq 'SmartInventory-SourceReceipt' -and $existing.ContractVersion -eq '1.2')
        if(-not $knownVersion -or $existing.TenantKey -ne $Context.TenantKey){throw 'Source receipt is not owned by the active tenant or has an unsupported version.'}
        foreach($field in @('OrganizationKey','EnvironmentKey','TenantId','Producer')){if($existing.$field -cne $Context[$field]){throw 'CMDB receipt identity changed.'}}
        if($Context.Started -and $existing.RunId -ne $Context.RunId){throw 'CMDB receipt run identity changed.'}
    }
    $temporary=$path+'.'+[guid]::NewGuid().ToString('N')+'.tmp'
    try{
        [IO.File]::WriteAllText($temporary,($Document | ConvertTo-Json -Depth 12),[Text.UTF8Encoding]::new($false))
        if(Test-Path -LiteralPath $path){[IO.File]::Replace($temporary,$path,[Management.Automation.Language.NullString]::Value,$true)}else{[IO.File]::Move($temporary,$path)}
    }finally{if(Test-Path -LiteralPath $temporary){Remove-Item -LiteralPath $temporary -Force}}
}

function Start-SmartM365CmdbSourceReceipt {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ScriptPath,[string]$SourceRootPath,[switch]$ReadOnly,[switch]$ConfiguredOutputs,[hashtable]$ScopeParameters=@{})
    if($script:SmartM365CmdbSourceContext){throw 'A CMDB source receipt is already active in this module.'}
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
        SourceRoot=$root;ReceiptPath=$path;ExpectedFiles=@($producers[0].Files)
        RequiredScope=[string]$producers[0].Scope;ScopeQualified=[bool]$ConfiguredOutputs;Scope=$(if($ConfiguredOutputs){'ConfiguredOutputs'}else{'NotDeclared'});Started=$false;Lock=$null
        ConfiguredOutputs=[bool]$ConfiguredOutputs;PublishedHashes=@{};OptionalFiles=@(if($producers[0].PSObject.Properties['OptionalFiles']){$producers[0].OptionalFiles})
        ScopeParameters=$effectiveScope
    }
    foreach($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','ScriptVersion')){if(-not $context[$field]){throw "Missing producer identity: $field"}}
    try{
        if((Test-Path -LiteralPath ($path+'.collection.lock')) -and ((Get-Item -LiteralPath ($path+'.collection.lock')).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Linked producer lock is not accepted.'}
        $context.Lock=[IO.File]::Open($path+'.collection.lock',[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        $document=@{}
        foreach($field in @('Owner','ContractVersion','TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer','ScriptVersion','RunId','StartedAtUtc')){$document[$field]=$context[$field]}
        $document.Status='Running';$document.IsPartialInventory=$true;$document.Files=@();$document.RequiredFiles=@($context.ExpectedFiles)
        $document.ScopeQualification=if($ConfiguredOutputs){'ConfiguredOutputsOnly'}else{'ConsumerScope'}
        $document.ScopeParameters=$effectiveScope
        Write-SmartM365CmdbReceiptDocument $context $document
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
    $context.PublishedHashes[$full]=(Get-FileHash -LiteralPath $full -Algorithm SHA256 -ErrorAction Stop).Hash
    if(-not $global:csvGeneratedPaths){$global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)}
    if($global:csvGeneratedPaths -is [array]){$global:csvGeneratedPaths=@($global:csvGeneratedPaths)+@($full)}else{[void]$global:csvGeneratedPaths.Add($full)}
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
                    $record.Errors=0;$record.IsPartialInventory=$(if($context.ConfiguredOutputs){$null}else{$context.ContainsKey('DomainCoverage') -and $null -ne $context.DomainCoverage});$record.Scope=$context.Scope
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
        Write-SmartM365CmdbReceiptDocument $context $document
        if(-not $complete){WriteLog -Message "Source qualification rejected for '$($context.Producer)': $reason Native CSVs remain preserved." -Level WARNING}
        return $context.ReceiptPath
    }finally{if($context.Lock){$context.Lock.Dispose()};$script:SmartM365CmdbSourceContext=$null}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCeF3QjPgFjO8V1
# 9z/kaS3fNX0b1XyAcSw8NdX7ehUcUKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILOn2nt9EnrYKpzXyTEmcHMGfXd1M0Nl5/XtMD6p4GWhMA0GCSqG
# SIb3DQEBAQUABIIBgA2kuc1jYzB8PzCJ2NbNuSYrFAy1hWJuprbrvjME0zdgkRqC
# r1XFYRuQa3wq4JmuUlvTE/TDdW6Y4DYYkXWu08UsEeiv6Laf4Cgs3x8D2hxCaTj7
# sA1kffyTPfr5Uwewa7i1/RqcU9SEOYeyg1f55ZGYdKgP2O2Oi2I7QRBkx7I4M/kJ
# BSP0LScWJTw5d/6Rt8gcR2qA9qSKmVMlt+EZMDGgaumXrdKqjm+UKXSx9q74RjtV
# i/puUBuAzSI13BqLzR9gv3JupB+wQJQv5D7/WNykOFRxDsG2op3U9IPwRhvE+8As
# duErEj0HYwSTQljehsyqH30SwpnpTi6aAfJjNEFI0+wlEG7JWjLJ75YIg+OgvtCl
# 1eR/nEmkRbNGtoGYBAa9mRV+g0/frQS09JbL34C/vCbKcKTHtd0TM7r8rVNnLTxo
# 1CJvQzRjcVIQHcexrFf/cxz4gN9HbUOMwm+PLI2pP40y+SlCmsFOa3QsyJXwguaR
# 5ActiTqgl30q2KoFn6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDYxODA3
# MDdaMC8GCSqGSIb3DQEJBDEiBCBddPAGTazEbYT7ZEAScj3p9X4xAPOrpa8mtdxG
# tyfRQjANBgkqhkiG9w0BAQEFAASCAgBVIc6FSzB4TDwSA+ZyyTXZu/dEx79AUJaw
# dYUq0g7a1mSOjbp0Zj5BZ9NA5TLPS/V56U8FwsQe0to0q9fJS7TopF0cuyVoUc13
# TLNRwVNTbiUgXm7njh9BTUs+hmKnt8IfoFmHjy3VLE41ah4XxLeSAPrstqCZ2Vdi
# 2uCfwfoy6FZAB+DNnVXCx79GnNv4R3Xi2Op/CFqRXaT0mELTNhB4EkUQGAl7+NaN
# 8BZoTVn2e5N/apiZIpxsDUZewmDTHv5ia4jzlVtiIup3YJ7OW06x8+dc5vvEvQPG
# q+87crlkPd8oumd4X9+A/b83W7lPjYvdqga7vyF9HImDKznhfUw2tpqBfDgStj1U
# zwvcQ82kH/bZqI3Zle5DcbyKJebj3QSIDSKiOTDFb4LIqgoybcv/qadpvIBWLpl3
# i+wwUZ8tAy1iScym3Nr6oO4/n9f85wMDuEMwdIKCpPsU2R2wejbwHWasPocj7jon
# eOmDGe2pc+KSwgb58F02XfJZom7+lZ94CwTiYTosktkmzVfSTCTyDTrv+I2KQQjJ
# 4BzZ3dqim3p5CUI3tql9lTMwLe0dwVbopd5wYv2dpruSAIKj31v8LZz/KsR6iDv7
# rwkKM4xIWp3EV5Xilrz7znLP+yMHWJIEhP2gGlb5rLMFz7cILqYkjGUCD6yCMwrt
# 7+Lc+0BWZA==
# SIG # End signature block
