# Version 0.1.0. Synthetic mapping rollback, contract and transfer guards only.
[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot, [string]$PythonCommand='python')
$ErrorActionPreference='Stop'
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ('mailbox-exceptions-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PreparedSharePointTransfer.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'NativeEvidence.psm1') -Force
$script:checks=0
function Check([bool]$Pass,[string]$Message){if(-not $Pass){throw $Message};$script:checks++}
function Reject([scriptblock]$Action){$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Check $failed 'Expected fail-closed rejection'}
$fixtures=Join-Path $root 'fixtures';$data=Join-Path $root 'data'
New-Item -ItemType Directory -Path $fixtures,$data | Out-Null
function MinimalWorkbook([string]$Path,[string[]]$Sheets){
    $archive=[IO.Compression.ZipFile]::Open($Path,[IO.Compression.ZipArchiveMode]::Create)
    try{$writer=[IO.StreamWriter]::new($archive.CreateEntry('xl/workbook.xml').Open());try{$writer.Write('<workbook><sheets>'+ (($Sheets|ForEach-Object {'<sheet name="'+$_+'"/>'}) -join '')+'</sheets></workbook>')}finally{$writer.Dispose()}}finally{$archive.Dispose()}
}
$persona='SmartWorkplaceIntelligence-PersonaClassification.xlsx';$site='SmartWorkplaceIntelligence-SiteClassification.xlsx';$exceptions='SmartWorkplaceIntelligence-MailboxMigrationExceptions.xlsx'
MinimalWorkbook (Join-Path $fixtures $persona) @('Personas','Rules','Exclusions')
MinimalWorkbook (Join-Path $fixtures $site) @('Settings','SiteTypes','Sites')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "../config/templates/$exceptions") -Destination (Join-Path $fixtures $exceptions)
$download={param($destination,$name) Copy-Item -LiteralPath (Join-Path $fixtures $name) -Destination $destination;Get-Item -LiteralPath $destination}.GetNewClosure()
$accepted=Receive-PreparedMappingWorkbooks -DataRoot $data -DownloadFile $download -MailboxMigrationExceptionsEnabled -PythonCommand $PythonCommand
Check (@(Get-ChildItem -LiteralPath $accepted -File).Count -eq 3) 'Enabled mapping set contains all three fresh workbooks'
$hashes=@{};foreach($name in @($persona,$site,$exceptions)){$hashes[$name]=(Get-FileHash -LiteralPath (Join-Path $data $name)).Hash}
$failThird={param($destination,$name) if($name -eq $exceptions){return $null}; & $download $destination $name}.GetNewClosure()
Reject {Receive-PreparedMappingWorkbooks -DataRoot $data -DownloadFile $failThird -MailboxMigrationExceptionsEnabled -PythonCommand $PythonCommand}
foreach($name in $hashes.Keys){Check ((Get-FileHash -LiteralPath (Join-Path $data $name)).Hash -eq $hashes[$name]) ('Third download failure preserves '+$name)}
$invalidThird={param($destination,$name) if($name -eq $exceptions){Copy-Item -LiteralPath (Join-Path $fixtures $site) -Destination $destination;return Get-Item $destination}; & $download $destination $name}.GetNewClosure()
Reject {Receive-PreparedMappingWorkbooks -DataRoot $data -DownloadFile $invalidThird -MailboxMigrationExceptionsEnabled -PythonCommand $PythonCommand}
foreach($name in $hashes.Keys){Check ((Get-FileHash -LiteralPath (Join-Path $data $name)).Hash -eq $hashes[$name]) ('Invalid third workbook preserves '+$name)}
$sourceContract=Join-Path $root 'source-contract.json.txt'
@{currentFiles=@();mappingFiles=@($persona,$site);dailyFiles=@();history=@();mailboxMigrationExceptionsFile=$exceptions}|ConvertTo-Json|Set-Content -LiteralPath $sourceContract
$disabled=@(Get-PreparedSourcePlan -DataRoot $data -SourceContractPath $sourceContract -MappingRoot $accepted)
$enabled=@(Get-PreparedSourcePlan -DataRoot $data -SourceContractPath $sourceContract -MappingRoot $accepted -MailboxMigrationExceptionsEnabled)
Check ($disabled.Count -eq 2 -and $enabled.Count -eq 3) 'Feature flag controls capture, with no cached third-file fallback when disabled'
Check (@($enabled|Where-Object Relative -EQ $exceptions).SHA256 -eq $hashes[$exceptions]) 'Workbook hash is frozen in capture provenance'
Remove-PreparedMappingWorkbooks -DataRoot $data -MappingRoot $accepted
foreach($name in $hashes.Keys){Check (Test-Path -LiteralPath (Join-Path $data $name)) ('Cleanup retains operational mapping '+$name)}
# Transfer old/new immutable cohorts; no filesystem transport or production data.
$product=Split-Path $PSScriptRoot -Parent
$base=Get-Content -LiteralPath (Join-Path $product 'config/prepared-evidence-contract.json.txt') -Raw|ConvertFrom-Json
$base=Merge-NativeEvidenceContract $base
$extension=Get-Content -LiteralPath (Join-Path $product 'config/mailbox-migration-exceptions-contract.json.txt') -Raw|ConvertFrom-Json
$intunePath=Join-Path $product 'config/intune-operational-contract.json.txt'
$nativePath=Join-Path $product 'config/native-output-contract.json.txt'
$exceptionPath=Join-Path $product 'config/mailbox-migration-exceptions-contract.json.txt'
$transferModule=Get-Module PreparedSharePointTransfer
function Manifest($Contract,[bool]$Enabled){
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($Contract|ConvertTo-Json -Depth 20))
    $hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
    $proof=@{NativeEvidence=@{ContractSHA256=(Get-FileHash $nativePath).Hash};IntuneOperations=@{ContractSHA256=(Get-FileHash $intunePath).Hash}}
    if($Enabled){$proof.MailboxMigrationExceptions=@{ContractSHA256=(Get-FileHash $exceptionPath).Hash}}
    @{Files=@($Contract.tables|ForEach-Object {@{File=$_.file}});Provenance=$proof;ContractSHA256=$hash}|ConvertTo-Json -Depth 20|ConvertFrom-Json
}
$path=Join-Path $root 'base.json.txt';[IO.File]::WriteAllText($path,($base|ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))
$old=Manifest $base $false
$resolved=& $transferModule {param($m,$p) Resolve-PreparedTransferContract $m $p} $old $path
Check ($resolved.tables.Count -eq $base.tables.Count) 'Existing immutable batch contract remains unchanged'
$expanded=$base|ConvertTo-Json -Depth 30|ConvertFrom-Json;$expanded.tables=@($expanded.tables)+@($extension.tables)
$new=Manifest $expanded $true
$resolved=& $transferModule {param($m,$p) Resolve-PreparedTransferContract $m $p} $new $path
Check ($resolved.tables.Count -eq $expanded.tables.Count) 'New exception tables participate in the exact unified contract hash'
$partial=Manifest $expanded $true;$partial.Files=@($partial.Files|Where-Object File -NE 'MailboxMigrationExceptions.csv')
Reject {& $transferModule {param($m,$p) Resolve-PreparedTransferContract $m $p} $partial $path}
$wrong=Manifest $expanded $true;$wrong.Provenance.MailboxMigrationExceptions.ContractSHA256=('0'*64)
Reject {& $transferModule {param($m,$p) Resolve-PreparedTransferContract $m $p} $wrong $path}
# Exercise the production config loader on a private synthetic file, never initialize a tenant.
. (Join-Path $PSScriptRoot '../../SmartM365/Config/SmartM365-TenantContext.ps1')
$localConfig=Join-Path $root 'fixture.local.json'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot '../../SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.local.json.template') -Destination ($localConfig+'.template')
@{PreparedPythonCommand='keep-python';Nested=@{Keep='private-value'}}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $localConfig
$loaded=Read-SmartM365JsonConfig -Path $localConfig -Required
Check ($loaded.PreparedMailboxMigrationExceptionsEnabled -eq $false -and $loaded.PreparedMailboxMigrationBindingsPath -eq '') 'New default keys are added by the real runtime loader'
Check ($loaded.PreparedPythonCommand -eq 'keep-python' -and $loaded.Nested.Keep -eq 'private-value') 'Runtime template merge preserves existing local and nested values'
$saved=(Read-SmartM365JsonDocument -Path $localConfig).Document
Check ($saved.PSObject.Properties['PreparedMailboxMigrationExceptionsEnabled'] -and $saved.PreparedPythonCommand -eq 'keep-python') 'Merged configuration is saved without overwriting local values'
# Execute the real Data Trust adapter against a synthetic audit and summary.
$auditPath=Join-Path $root 'MailboxMigrationExceptionsAudit.json.txt'
@{'Source Rows'=0;'Snapshot Date'='2026-10-19';'Workbook SHA256'=('0'*64);'Captured UTC'='2026-10-19T08:00:00Z'}|ConvertTo-Json|Set-Content -LiteralPath $auditPath
[pscustomobject]@{'Domain Key'='__all__';'Review Count'=2}|Export-Csv -LiteralPath (Join-Path $root 'MailboxMigrationExceptionSummary.csv') -NoTypeInformation
$trustOutput=Join-Path $root 'trust.csv'
& (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'New-DataTrustEvidence.ps1') -DataRoot $data -OutputPath $trustOutput -AsOfDate '2026-10-19' -MailboxMigrationExceptionsAuditPath $auditPath | Out-Null
Check ($LASTEXITCODE -eq 0) 'Data Trust adapter executes successfully'
$trust=@(Import-Csv -LiteralPath $trustOutput|Where-Object 'Source Name' -EQ 'Mailbox migration exception register')
Check ($trust.Count -eq 1 -and $trust[0].'Decision Readiness' -eq 'Review required' -and $trust[0].'Evidence Age Days' -eq '0') 'Data Trust exposes one dated, review-qualified registry row'
[pscustomobject]@{Status='Passed';Checks=$script:checks;TestRoot=$root}
