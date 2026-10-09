#Requires -Version 5.1
<#
.SYNOPSIS
Synthetic shared source-receipt qualification. No tenant, API, mail or synchronized data access.
.VERSION
1.0.2
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core globals are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs','',Justification='Begin/Finish name fixture phases, not public commands.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Offline validation mock retains the publisher API; data is validated by the real receipt parser.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$helper=Join-Path $root 'Modules/SmartM365.Core/SmartM365-CmdbReceipt.ps1'
$registry=Get-Content (Join-Path (Split-Path $helper) 'SmartM365-SourceReceipts.json.txt') -Raw | ConvertFrom-Json
$temporary=Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-SourceReceiptTests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
$saved=@{}
foreach($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')){
    $v=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name]=@{Exists=($null -ne $v);Value=$(if($v){$v.Value}else{$null})}
}
$global:SmartM365TenantKey='synthetic';$global:SmartM365OrganizationKey='test'
$global:SmartM365EnvironmentKey='test';$global:SmartM365TenantId='synthetic-tenant'
$global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$mock=New-Module -Name SyntheticSourceReceipts -ScriptBlock {
    param($path)
    . $path
    $script:Limited=$false
    function Test-SmartM365MaxItemsMode { return $script:Limited }
    function Get-SmartM365ScriptVersionFromFile { param($Path) $null=$Path;return 'synthetic-test' }
    function WriteLog { param($Message,$Level) $null=$Message;$null=$Level }
    Export-ModuleMember -Function *
} -ArgumentList $helper
$script:checks=0
function Check {param([bool]$Condition,[string]$Message) if(-not $Condition){throw $Message};$script:checks++}
function Begin-Fixture {param($Definition,[string]$Label,[switch]$ReadOnly)
    $folder=Join-Path $temporary $Label
    & $mock {param($p,$r,$ro) Start-SmartM365SourceReceipt -ScriptPath $p -SourceRootPath $r -ReadOnly:$ro -ScopeParameters @{PrimaryOnly=[Management.Automation.SwitchParameter]::new($true)}} $Definition.Script $folder $ReadOnly.IsPresent
    return @{Folder=$folder;Receipt=(Join-Path $folder $Definition.Receipt);Definition=$Definition}
}
function Write-FixtureCsv {param($Fixture,[string]$Name,[string]$Delimiter=',',[switch]$Empty,[switch]$NoRegister)
    $path=Join-Path $Fixture.Folder $Name
    $body='TenantKey'+$Delimiter+'Name'+"`r`n"
    if(-not $Empty){$body+='synthetic'+$Delimiter+'"multi'+"`r`n"+'line"'+"`r`n"}
    [IO.File]::WriteAllText($path,$body,[Text.UTF8Encoding]::new($false))
    if(-not $NoRegister){& $mock {param($p) Register-SmartM365SourceCsv -Path $p} $path}
    return $path
}
function Finish-Fixture {param($Fixture,[string]$Status='Success',[int]$Errors=0)
    $result=& $mock {param($s,$e) Complete-SmartM365SourceReceipt -Status $s -ErrorCount $e} $Status $Errors
    return (Get-Content -LiteralPath $result -Raw | ConvertFrom-Json)
}
try {
    foreach($definition in $registry.Producers){
        $fixture=Begin-Fixture $definition ([IO.Path]::GetFileNameWithoutExtension($definition.Script))
        $running=Get-Content ($fixture.Receipt -replace '\.current\.json\.txt$','.run.json.txt') -Raw | ConvertFrom-Json
        Check ($running.Status -eq 'Collecting' -and -not(Test-Path $fixture.Receipt)) 'First acquisition fabricated completion proof.'
        $delimiter=if($definition.Script -like '*InfrastructureAndReadiness*'){';'}else{','}
        foreach($name in $definition.Files){$null=Write-FixtureCsv $fixture $name $delimiter}
        $null=Write-FixtureCsv $fixture 'AdditionalPublished.csv' $delimiter
        $receipt=Finish-Fixture $fixture
        Check ($receipt.Owner -eq 'SmartInventory-SourceReceipt' -and $receipt.ContractVersion -eq '1.2') 'Generic ownership/version missing.'
        Check ($receipt.Status -eq 'Completed' -and $receipt.Files.Count -eq $definition.Files.Count+1) 'A current published CSV was omitted.'
        Check ($receipt.ScopeQualification -eq 'ConfiguredOutputsOnly' -and $null -eq $receipt.IsPartialInventory -and -not $receipt.FullInventoryQualified -and -not $receipt.ConsumerScopeQualified) 'Configured success was misrepresented as complete tenant inventory.'
        Check ($receipt.ScopeParameters.PrimaryOnly -eq $true) 'Effective scope evidence lost.'
        foreach($record in $receipt.Files){
            Check ($record.Rows -eq 1 -and $record.Delimiter -eq $delimiter -and $record.RunId -eq $receipt.RunId -and $record.SHA256 -eq (Get-FileHash (Join-Path $fixture.Folder $record.File)).Hash) 'Logical rows, delimiter, hash or lineage differ.'
        }
        Check (@($receipt.Files | Where-Object Required).Count -eq $definition.Files.Count) 'Required and additional outputs confused.'
        $scriptFile=Get-ChildItem (Join-Path $root SmartInventory) -Recurse -File -Filter $definition.Script | Select-Object -First 1
        $t=$null;$e=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($scriptFile.FullName,[ref]$t,[ref]$e)
        Check ($e.Count -eq 0) ('Collector syntax failed: '+$definition.Script)
        $commands=$ast.FindAll({param($a) $a -is [Management.Automation.Language.CommandAst]},$true)
        Check (@($commands | Where-Object {$_.GetCommandName() -match '^Start-(Core)?SmartM365SourceReceipt$'}).Count -eq 1) ('Startup coverage missing: '+$definition.Script)
        Check ($ast.Extent.Text -match "1\.0\.72|1\.0\.5\d|1\.0\.8") ('Required module guard missing: '+$definition.Script)
        Check (@($commands | Where-Object {$_.GetCommandName() -match '^Complete-(Core)?SmartM365(ExecutionContext|EvidenceRuntime|SourceReceipt)$'}).Count -gt 0) ('Completion coverage missing: '+$definition.Script)
    }
    # Repeated attempts must retain proof bytes and acquisition dates until a qualified replacement.
    $d=$registry.Producers[0]
    $f=Begin-Fixture $d 'separated-run'
    foreach($name in $d.Files){$null=Write-FixtureCsv $f $name}
    $proof=Finish-Fixture $f
    $hash=(Get-FileHash $f.Receipt).Hash
    $f=Begin-Fixture $d 'separated-run'
    Check ((Get-FileHash $f.Receipt).Hash -ceq $hash) 'Collecting changed previous proof bytes or acquisition timestamps.'
    & $mock {param($path,$proof) Assert-SmartM365SourcePublication -ReceiptPath $path -Proof $proof} $f.Receipt $proof
    $blocked=$false
    try {& $mock {param($path) Start-SmartM365SourceReceipt -ScriptPath 'SmartM365-AD-HealthCheck.ps1' -SourceRootPath $path} $f.Folder}catch{$blocked=$true}
    Check $blocked 'Overlapping collection was accepted.'
    $r=Finish-Fixture $f Failed 1
    Check ($r.Status -eq 'Failed' -and -not $r.PublicationStarted -and (Get-FileHash $f.Receipt).Hash -ceq $hash) 'Pre-publication failure erased previous validated proof.'
    & $mock {param($path,$proof) Assert-SmartM365SourcePublication -ReceiptPath $path -Proof $proof} $f.Receipt $proof
    $f=Begin-Fixture $d 'separated-run'
    & $mock {param($path) Start-SmartM365SourcePublication -Path $path} (Join-Path $f.Folder $d.Files[0])
    $blocked=$false
    try {& $mock {param($path,$proof) Assert-SmartM365SourcePublication -ReceiptPath $path -Proof $proof} $f.Receipt $proof}catch{$blocked=$true}
    Check $blocked 'In-progress canonical replacement was accepted.'
    $r=Finish-Fixture $f Failed 1
    Check ($r.UnqualifiedPublication -and (Get-FileHash $f.Receipt).Hash -ceq $hash) 'Failed replacement was not fenced or previous proof was overwritten.'
    $f=Begin-Fixture $d 'separated-run'
    $blocked=$false
    try {& $mock {param($path,$proof) Assert-SmartM365SourcePublication -ReceiptPath $path -Proof $proof} $f.Receipt $proof}catch{$blocked=$true}
    Check $blocked 'Retry cleared an unqualified replacement fence.'
    foreach($name in $d.Files){$null=Write-FixtureCsv $f $name}
    $newProof=Finish-Fixture $f
    Check ($newProof.RunId -cne $proof.RunId -and $newProof.PublicationProtocol -eq 1) 'Qualified replacement did not advance the proof.'
    & $mock {param($path,$proof) Assert-SmartM365SourcePublication -ReceiptPath $path -Proof $proof} $f.Receipt $newProof
    $runPath=$f.Receipt -replace '\.current\.json\.txt$','.run.json.txt'
    $run=Get-Content $runPath -Raw | ConvertFrom-Json
    Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication -and $run.RunId -ceq $newProof.RunId) 'Successful publication did not release its fence.'
    $f=Begin-Fixture $d 'separated-run'
    $qualifiedHash=(Get-FileHash $f.Receipt).Hash
    $null=Finish-Fixture $f Failed 1
    $uploads=@(& $mock {$script:SmartM365SourceReceiptUploadPaths})
    Check ($uploads.Count -eq 1 -and $uploads[0] -ceq $runPath) 'Failed attempt re-uploaded a stale completion proof.'
    Check ((Get-FileHash $f.Receipt).Hash -ceq $qualifiedHash -and @(Get-ChildItem $f.Folder -Filter '*.json.txt').Count -eq 2) 'Source receipt history or duplicate snapshots were created.'
    $f=Begin-Fixture $d 'fence-write-failure'
    $runPath=$f.Receipt -replace '\.current\.json\.txt$','.run.json.txt'
    $savedRun=$runPath+'.fixture-save'
    Move-Item -LiteralPath $runPath -Destination $savedRun
    New-Item -ItemType Directory -Path $runPath | Out-Null
    foreach($attempt in 1..2){
        $blocked=$false
        try{& $mock {param($path) Start-SmartM365SourcePublication -Path $path} (Join-Path $f.Folder $d.Files[0])}catch{$blocked=$true}
        Check ($blocked -and -not (& $mock {$script:SmartM365CmdbSourceContext.PublicationStarted})) 'A failed fence write allowed a subsequent canonical replacement.'
    }
    [IO.Directory]::Delete($runPath)
    Move-Item -LiteralPath $savedRun -Destination $runPath
    $null=Finish-Fixture $f Failed 1
    $rbac=$registry.Producers | Where-Object Script -eq 'SmartM365-Intune-RBAC-GroupMembers.ps1'
    $f=Begin-Fixture $rbac 'rbac-quoted-semicolon'
    $csv=Join-Path $f.Folder $rbac.Files[0]
    [IO.File]::WriteAllText($csv,(('"TenantKey";"Name"'+"`r`n"+'"synthetic";"name, with comma"'+"`r`n")),[Text.UTF8Encoding]::new($false))
    & $mock {param($p) Register-SmartM365SourceCsv -Path $p} $csv
    $r=Finish-Fixture $f
    Check ($r.Status -eq 'Completed' -and $r.Files[0].Rows -eq 1 -and $r.Files[0].Delimiter -eq ';') 'Quoted semicolon RBAC CSV was not qualified.'
    $f=Begin-Fixture $rbac 'rbac-malformed-semicolon'
    $csv=Join-Path $f.Folder $rbac.Files[0]
    [IO.File]::WriteAllText($csv,(('"TenantKey";"Name"'+"`r`n"+'"synthetic";"unclosed'+"`r`n")),[Text.UTF8Encoding]::new($false))
    & $mock {param($p) Register-SmartM365SourceCsv -Path $p} $csv
    $r=Finish-Fixture $f
    Check ($r.Status -eq 'Failed' -and $r.Files.Count -eq 0) 'Malformed semicolon RBAC CSV was qualified.'
    $definition=$registry.Producers[0]
    # Exercise actual atomic publishers without running module initialization or external actions.
    foreach($publisher in @(
        @{Path=(Join-Path (Split-Path $helper) 'SmartM365.Core.psm1');Writer='Write-SmartM365PreparedCsvAtomically';Label='core'},
        @{Path=(Join-Path (Split-Path $helper) 'Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1');Writer='Write-SmartM365CsvAtomically';Label='ps5'}
    )){
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($publisher.Path,[ref]$tokens,[ref]$errors)
        Check ($errors.Count -eq 0) 'Atomic publisher module syntax failed.'
        foreach($functionName in @($publisher.Writer,'Copy-SmartM365FileAtomically')){
            $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $functionName},$true)
            Check ($null -ne $node) 'Atomic publisher function missing.'
            $definitionText=$node.Extent.Text.Replace(('function '+$functionName+' {'),('function script:'+$functionName+' {'))
            & $mock ([scriptblock]::Create($definitionText))
        }
        & $mock {
            $script:CanonicalMoveChecks=0
            function script:Move-Item {
                [CmdletBinding()]
                param([string]$LiteralPath,[string]$Destination,[switch]$Force)
                $ctx=$script:SmartM365CmdbSourceContext
                if($ctx -and [IO.Path]::GetDirectoryName($Destination) -ieq $ctx.SourceRoot -and [IO.Path]::GetExtension($Destination) -ieq '.csv'){
                    $state=Get-Content $ctx.RunPath -Raw | ConvertFrom-Json
                    if($state.Status -ne 'Publishing' -or -not $state.UnqualifiedPublication){throw 'Canonical rename was not fenced before mutation.'}
                    $script:CanonicalMoveChecks++
                }
                Microsoft.PowerShell.Management\Move-Item @PSBoundParameters
            }
            function script:Add-SmartM365TenantKeyToCsvData {param($Data,$Columns) return @{Data=$Data;Columns=$Columns}}
            function script:Assert-SmartM365CsvDataCompleteness {param($Data,$Columns,$TimestampedPath,$LatestPath)}
        }
        $f=Begin-Fixture $definition ($publisher.Label+'-atomic-write')
        $csv=Join-Path $f.Folder $definition.Files[0]
        & $mock {param($writer,$p) & $writer -Data @([pscustomobject]@{TenantKey='synthetic';Name="multi`nline"}) -Columns @('TenantKey','Name') -Path $p -Encoding UTF8} $publisher.Writer $csv
        $r=Finish-Fixture $f
        Check ($r.Status -eq 'Completed' -and $r.Files[0].Rows -eq 1) 'Actual atomic writer did not register current output.'
        $f=Begin-Fixture $definition ($publisher.Label+'-atomic-copy')
        & $mock {param($s,$d) Copy-SmartM365FileAtomically -SourcePath $s -DestinationPath $d} $csv (Join-Path $f.Folder $definition.Files[0])
        $r=Finish-Fixture $f
        Check ($r.Status -eq 'Completed' -and $r.Files[0].SHA256 -eq (Get-FileHash $csv).Hash) 'Actual atomic copy did not register current output.'
        Check ((& $mock {$script:CanonicalMoveChecks}) -eq 2) 'Canonical writer and copier did not each set a pre-mutation fence.'
    }
    $f=Begin-Fixture $definition 'missing-required'
    $old=Write-FixtureCsv $f $definition.Files[0] -NoRegister
    $oldHash=(Get-FileHash $old).Hash
    $r=Finish-Fixture $f
    Check ($r.Status -eq 'Failed' -and $r.Files.Count -eq 0 -and (Get-FileHash $old).Hash -eq $oldHash) 'An old CSV rescued a missing acquisition or was deleted.'
    $f=Begin-Fixture $definition 'empty'
    $null=Write-FixtureCsv $f $definition.Files[0] -Empty
    $r=Finish-Fixture $f
    Check ($r.Status -eq 'Completed' -and $r.Files[0].Rows -eq 0) 'A valid empty output was rejected.'
    $f=Begin-Fixture $definition 'mutation'
    $csv=Write-FixtureCsv $f $definition.Files[0]
    [IO.File]::AppendAllText($csv,"synthetic,changed`r`n")
    $r=Finish-Fixture $f
    Check ($r.Status -eq 'Failed' -and $r.Error -match 'changed after') 'Publication hash mutation was ignored.'
    foreach($case in @('failed','errors')){
        $f=Begin-Fixture $definition $case
        $null=Write-FixtureCsv $f $definition.Files[0]
        $r=if($case -eq 'failed'){Finish-Fixture $f Failed}else{Finish-Fixture $f Success 1}
        Check ($r.Status -eq 'Failed' -and $r.Files.Count -eq 0) 'Failed acquisition was qualified.'
    }
    $optional=$registry.Producers | Where-Object {$_.OptionalFiles.Count -gt 0} | Select-Object -First 1
    $f=Begin-Fixture $optional 'optional-old'
    foreach($name in $optional.Files){$null=Write-FixtureCsv $f $name}
    $old=Write-FixtureCsv $f $optional.OptionalFiles[0] -NoRegister
    $r=Finish-Fixture $f
    Check ($r.Status -eq 'Completed' -and @($r.Files | Where-Object File -eq $optional.OptionalFiles[0]).Count -eq 0 -and -not $r.OptionalFiles[0].Published) 'An old optional file was reported as current.'
    $f=Begin-Fixture $definition 'readonly' -ReadOnly
    Check (-not (Test-Path $f.Folder)) 'Read-only execution created canonical artifacts.'
    & $mock {$script:Limited=$true}
    $f=Begin-Fixture $definition 'limited'
    & $mock {$script:Limited=$false}
    Check (-not (Test-Path $f.Folder)) 'MAXITEMS replaced canonical metadata.'
    $legacy=Join-Path $temporary 'legacy-owner'
    New-Item -ItemType Directory -Path $legacy | Out-Null
    $legacyPath=Join-Path $legacy $definition.Receipt
    $legacyDocument=@{Owner='SmartInventory-CmdbSourceReceipt';ContractVersion='1.1';TenantKey='synthetic';OrganizationKey='test';EnvironmentKey='test';TenantId='synthetic-tenant';Producer=$definition.Script;RunId='old-run'}
    [IO.File]::WriteAllText($legacyPath,($legacyDocument | ConvertTo-Json))
    $f=Begin-Fixture $definition 'legacy-owner'
    $null=Write-FixtureCsv $f $definition.Files[0]
    $r=Finish-Fixture $f
    Check ($r.Owner -eq 'SmartInventory-SourceReceipt' -and $r.RunId -ne 'old-run') 'The next actual run did not transition legacy ownership.'
    $legacyDocument.Owner='UnknownOwner'
    [IO.File]::WriteAllText($legacyPath,($legacyDocument | ConvertTo-Json))
    $hash=(Get-FileHash $legacyPath).Hash;$rejected=$false
    try {$null=Begin-Fixture $definition 'legacy-owner'}catch{$rejected=$true}
    Check ($rejected -and (Get-FileHash $legacyPath).Hash -eq $hash) 'Foreign metadata was overwritten.'
    $cmdb=Get-Content (Join-Path (Split-Path $helper) 'SmartM365-CmdbSources.json.txt') -Raw | ConvertFrom-Json
    $scripts=@($cmdb.Producers.Script)+@($registry.Producers.Script)
    Check ($scripts.Count -eq 47 -and @($scripts | Select-Object -Unique).Count -eq 47) 'DATA-LAST receipt coverage is not 47 distinct producers.'
    "PASS: $script:checks shared receipt checks; PowerShell $($PSVersionTable.PSVersion). No production access."
} finally {
    & $mock {if($script:SmartM365CmdbSourceContext -and $script:SmartM365CmdbSourceContext.Lock){$script:SmartM365CmdbSourceContext.Lock.Dispose()}}
    Remove-Module $mock -ErrorAction SilentlyContinue
    foreach($name in $saved.Keys){if($saved[$name].Exists){Set-Variable -Name $name -Scope Global -Value $saved[$name].Value}else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}}
    $full=[IO.Path]::GetFullPath($temporary)
    if([IO.Path]::GetDirectoryName($full).TrimEnd('\') -eq [IO.Path]::GetTempPath().TrimEnd('\') -and [IO.Path]::GetFileName($full) -like 'SmartInventory-SourceReceiptTests-*'){Remove-Item -LiteralPath $full -Recurse -Force}
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCh4PQLd5hG2x0/
# 7mGK+IBf6yZ8WAMwMkC/1YaWEdvy7qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCEVB32JmrJHaAe31H7f8MO
# FIhX+IGeNl6zlLs4HSU72jANBgkqhkiG9w0BAQEFAASCAYCrHEJWl5fPIk4aCnVd
# jnY2NjQ0nYJ3rp4GPs0wUTwjrlDdeamug1AQ7XMlUZy7JCMFd0K18H+2dclstpuX
# eMuLgDZ86jXLUHCNcbAaM2VLGXq+HqQQLB7TA9JrGfd79gEowr7gbdWN0JJkNdHc
# uizzh/ZSu0yh+UOF8CjBLHQPqII4k7fUl/x8FWvMsCvEXbBgB+SLdx/tjAOHiU2X
# BYOQRnXmcwZ8sDJwMQbvik5utzBiStQGBxgyi1inVCV5ZN/3t1RxEryUCdctFuvH
# 18AgyZFpg/68L7BRjmRSIJ5S0xEpFIM+LcJKYGH3AA8lXciLY2TespYBeDaymGPE
# YCEpJ19qYdY4upGC5IEoKj5k24xIUYHh6NsFbRJ7VKYsceWceIe/Uv/gqDYVrShQ
# nwS7Sn3KChpM0osQWzspqQxyF+tT1lMm6f9fkC0W3tIWPo5jjycHy1Qa6HDoJc+/
# qyWfXJkCXpuWDCufzRodRdm5KXBhQ2lUUfQEzNv96trSZNU=
# SIG # End signature block
