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
        Check ($ast.Extent.Text -match "1\.0\.72|1\.0\.50|1\.0\.54|1\.0\.55|1\.0\.8") ('Required module guard missing: '+$definition.Script)
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDZjin+rmcyqy/k
# XZX63lXuZ7f+jReOspr2ITTydz3G/qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIILqzwiO2te7j45m+p3kHR3OjPrfSSIMCbLHQPgvFshSMA0GCSqG
# SIb3DQEBAQUABIIBgGqi/u9hRQAeztI8tHFLy9rQHQi4b6BMVsXsZlf+W0iViw+/
# IoW+jVGhCAecVxgaNpM9XJlSV+0PNkwlNoCJJzOTq2cagNkanSof5rWOy6y7UnS5
# /Z/TLG/cHtis59bgsYvsX2c2cdxW1F6Gu0JXXHRYCzP3mQOmsxVT9du1umEk1OFX
# MifOgT+P+xqIj8pVvlolBz7c3yJHLbIbMhZURQTjhSdaQZlBadyFC0aD+97Ol1hZ
# jUNpFv+6237lBnoA+Rwz/TpXMZjrRQmz3Ry9iysa8HPu1WtLxAAlgdEmEg80GH2X
# nMZQl/SYLXTDgOLHRFQPtNztdMsvJT1emI3fLC//qi9+g9JLnAERdCimK4ya+Mhc
# 9v/WVufMdgVlr3XWPNDtu02tkclRLrMP0S04J44DxwZq1ocLY+XiNWkNOuhFvP1J
# ul1uPc4TTUsw92/sqsiT+E9T0Aq7XO48r8JMgi+ZDPaH1xnOBv/RU/GqAHQscXuf
# FoOcCPyl1KEk79jBNaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDkyMjA0
# MTRaMC8GCSqGSIb3DQEJBDEiBCDiYxx906DydnlLWgSSXcxcycMaKR2ij9T3DJ4u
# sDhSgzANBgkqhkiG9w0BAQEFAASCAgCxIyvimVAPxZfjxsFZ1hMbMujPJpjTJNgi
# yD8VA9FwBTlyAc+X/lNLcn+jlwaXeGgws43DmFaqVN8C2BjF4a3W/vZUfKvdHpWH
# Oe8gieBa3UP+HAXOmGal2mq5H0ySESWFjFOjGE5tx/Ai68pc0oODq6dMxrZr4wD6
# h5MiNuxxTDvb5dz9spBUWudwpmVeAuauPAZSx7UKyBGbNx7+QZVMNjVyTM0vYnyS
# HOMWyZR++e5KvgkeYTe0zpg28Zgi4WQlwjFID+lRdXLZFUKQ2xO6yt3tD1vAivt7
# LPw4mm6A5wZeAr/xiZ4SCqDnaO19MdNIdonPBXrKfHlbPGQS9c3xSuEGlFlvgf86
# DMLF+RMxrVF3oVnn56Bl7GjxTs2QL//bWYOzpFI9b7WveECh2dwQZdElc+zQoSsF
# u5B34iBbWXxUy+McSHdpyjhbdpnkLUW8bIEisHiAaESXdhc33Wuk7rOtUQkR4ax2
# tFgapCeVV6rj3411Up9ZOJDboNxI6wW8qasLGii/1hQefs91rj5l1UQfsl7Wpdx4
# ytZZn4kmQ5ka/fNa+ECgBedlKkDl2nbi9l7O8IBp9sp/UmwOQjWASJ8T87fsCz9S
# 6lBOKPC3ArYQlMyYKUWLMFHdJaWGnW2fxUBZ/Xb2zrOokT/Q7inqpqt4vLIbfVjN
# FdZ3q18X8Q==
# SIG # End signature block
