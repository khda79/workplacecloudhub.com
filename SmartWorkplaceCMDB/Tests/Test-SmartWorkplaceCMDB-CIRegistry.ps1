# Stable V1. Synthetic local tests only; no configuration or tenant connections.
[CmdletBinding()]param()
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
$projectRoot=Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $projectRoot 'Modules\SmartWorkplaceCMDB.CI\SmartWorkplaceCMDB.CI.psd1') -Force
$contract=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json') | ConvertFrom-Json
$catalog=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.ci.catalog.json') | ConvertFrom-Json
$tempParent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/')
$testRoot=Join-Path $tempParent ('cmdb-ci-tests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
$passed=0; $failed=0
function Assert-True([bool]$Value,[string]$Message) { if(-not $Value){throw $Message} }
function Assert-Throws([scriptblock]$Action,[string]$Pattern) {
    $caught=$null
    try { & $Action | Out-Null } catch {$caught=$_}
    if($null -eq $caught -or $caught.Exception.Message -notmatch $Pattern){throw "Expected error: $Pattern; received: $caught"}
}
function Write-Row([string]$File,[object[]]$Rows) {
    $table=@($contract.tables | Where-Object name -eq $File)[0]
    $path=Join-Path $script:inputPath $File
    if($Rows.Count){$Rows | Select-Object $table.columns | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8}
    else {($table.columns -join ',') | Set-Content -LiteralPath $path -Encoding UTF8}
}
function New-Fixture {
    $case=Join-Path $testRoot ([guid]::NewGuid().ToString('N'))
    $script:inputPath=Join-Path $case 'CMDB'
    $script:outputPath=Join-Path $case 'registry'
    New-Item -ItemType Directory -Path $script:inputPath -Force | Out-Null
    foreach($d in $catalog.types) {
        $row=[ordered]@{TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId='';SourceSystem='ExampleInventory';SourceCollectedDateTime='2026-01-01T00:00:00Z'}
        $row[$d.idColumn]='example-test|'+$d.name.ToLowerInvariant()+'|internal-1'
        $row[$d.nameColumn]=$d.name+' example'
        $row[$d.sourceIdColumn]='source-'+$d.name.ToLowerInvariant()+'-1'
        Write-Row $d.file @([pscustomobject]$row)
    }
    $relationship=[pscustomobject]@{TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId='';CmdbRelationshipId='example-test|relationship|1';FromEntityType='User';FromEntityId='example-test|user|internal-1';ToEntityType='Device';ToEntityId='example-test|device|internal-1';RelationshipType='PrimaryUser';SourceSystem='ExampleInventory';ConfidenceScore='0.5';SourceCollectedDateTime='2026-01-01T00:00:00Z'}
    Write-Row 'CMDB_Relationships.csv' @($relationship)
    $script:argsCI=@{InputRootPath=$script:inputPath;OutputDirectory=$script:outputPath;OrganizationKey='example';EnvironmentKey='test';TenantKey='example-test'}
}
function Test-Case([string]$Name,[scriptblock]$Action) {
    try {New-Fixture; & $Action; $script:passed++; Write-Output "PASS $Name"}
    catch {$script:failed++; Write-Output "FAIL $Name : $($_.Exception.Message)"}
}
function New-RawFixture {
    $script:rawPath=Join-Path (Split-Path $outputPath) 'Raw'
    $script:rawContract=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json') | ConvertFrom-Json
    $specs=@(
        @('User','Entra_Users.csv','MicrosoftEntraID','SourceUserId','source-user-1','SourceUserId','source-user-1'),
        @('Group','Entra_Groups.csv','MicrosoftEntraID','SourceGroupId','source-group-1','SourceGroupId','source-group-1'),
        @('Device','Entra_Devices.csv','MicrosoftEntraID','SourceObjectId','entra-object-1','SourceDeviceId','source-device-1'),
        @('Device','Intune_ManagedDevices.csv','MicrosoftIntune','ManagedDeviceId','intune-managed-1','AzureAdDeviceId','source-device-1'),
        @('License','M365_SubscribedSkus.csv','MicrosoftGraph','SubscribedSkuId','subscription-1','SkuId','source-license-1'),
        @('Mailbox','ExchangeOnline_Mailboxes.csv','ExchangeOnline','SourceMailboxId','mailbox-object-1','ExternalDirectoryObjectId','source-mailbox-1')
    )
    foreach ($s in $specs) {
        $definition=@($catalog.types | Where-Object name -eq $s[0])[0]
        $ci=Import-Csv (Join-Path $inputPath $definition.file)
        $ci.SourceSystem=if ($s[0] -eq 'Device') {'MicrosoftEntraID+MicrosoftIntune'} else {$s[2]}
        if ($s[0] -eq 'Mailbox') { $ci.CmdbMailboxId='example-test|exo-mailbox|mailbox-object-1' }
        Write-Row $definition.file @($ci)
        $row=[ordered]@{TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId='';SourceSystem=$s[2];SourceCollectedDateTime='2026-01-01T00:00:00Z'}
        $row[$s[3]]=$s[4]; $row[$s[5]]=$s[6]
        Write-Raw $s[1] @([pscustomobject]$row)
    }
}
function Get-RawPath([string]$File) {
    $table=@($rawContract.tables | Where-Object name -eq $File)[0]
    Join-Path (Join-Path $rawPath $table.area.Substring(4)) $File
}
function Write-Raw([string]$File,[object[]]$Rows) {
    $table=@($rawContract.tables | Where-Object name -eq $File)[0]
    $path=Get-RawPath $File
    New-Item -ItemType Directory -Path (Split-Path $path) -Force | Out-Null
    if ($Rows.Count) {$Rows | Select-Object $table.columns | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8}
    else {($table.columns -join ',') | Set-Content -LiteralPath $path -Encoding UTF8}
}
function New-Change([string]$Id,[string]$Field,[string]$Old,[string]$New,[string]$Time='2026-01-01T00:00:00Z') {
    [pscustomobject][ordered]@{TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId='';ChangeId=$Id;CI_ID='example-test|device|internal-1';Field=$Field;OldValue=$Old;NewValue=$New;EffectiveUtcDateTime=$Time;ChangedBy='example-test|user|internal-1';Reason='Synthetic reviewed change'}
}
function Write-Journal([object[]]$Events) {
    $script:journalPath=Join-Path (Split-Path $outputPath) 'governance.csv'
    if ($Events.Count) {$Events | Select-Object $catalog.governanceColumns | Export-Csv -LiteralPath $journalPath -NoTypeInformation -Encoding UTF8}
    else {($catalog.governanceColumns -join ',') | Set-Content -LiteralPath $journalPath -Encoding UTF8}
}
function New-ActiveJournal {
    @(
        (New-Change '1' 'BusinessOwner' '' 'example-test|user|internal-1'),
        (New-Change '2' 'TechnicalOwner' '' 'example-test|group|internal-1'),
        (New-Change '3' 'SupportGroup' '' 'example-test|group|internal-1'),
        (New-Change '4' 'LifecycleStatus' 'Unknown' 'Planned'),
        (New-Change '5' 'LifecycleStatus' 'Planned' 'Active')
    )
}
function New-OrganizationFixture {
    $script:organizationPath=Join-Path (Split-Path $outputPath) 'organization.csv'
    $script:organizationRows=@(
        [pscustomobject]@{TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId='';ReferenceId='example-test|country|1';ReferenceType='Country';Name='Example country';ParentReferenceId=''},
        [pscustomobject]@{TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId='';ReferenceId='example-test|entity|1';ReferenceType='Entity';Name='Example entity';ParentReferenceId='example-test|country|1'},
        [pscustomobject]@{TenantKey='example-test';OrganizationKey='example';EnvironmentKey='test';TenantId='';ReferenceId='example-test|site|1';ReferenceType='Site';Name='Example site';ParentReferenceId='example-test|entity|1'}
    )
    Write-Organization $organizationRows
}
function Write-Organization([object[]]$Rows) {
    $cols=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','ReferenceId','ReferenceType','Name','ParentReferenceId')
    if ($Rows.Count) {$Rows | Select-Object $cols | Export-Csv -LiteralPath $organizationPath -NoTypeInformation -Encoding UTF8}
    else {($cols -join ',') | Set-Content -LiteralPath $organizationPath -Encoding UTF8}
}
try {
    Test-Case 'ValidateOnly reads but does not write' {
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -ValidateOnly
        Assert-True ($r.CICount -eq 5 -and $r.RelationshipCount -eq 1) 'Counts'
        Assert-True (-not (Test-Path $outputPath)) 'Validation wrote output'
    }
    Test-Case 'Registry preserves IDs, source identity and unknown governance' {
        $before=@{}; Get-ChildItem $inputPath -File | ForEach-Object {$before[$_.Name]=(Get-FileHash $_.FullName).Hash}
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
        $ci=@(Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv'))
        $user=@($ci | Where-Object CI_Type -eq User)[0]
        Assert-True ($user.CI_ID -eq 'example-test|user|internal-1' -and $user.SourceID -eq 'source-user-1') 'IDs conflated'
        Assert-True ($user.BusinessOwner -eq '' -and $user.LifecycleStatus -eq 'Unknown') 'Governance inferred'
        Assert-True ($r.SourceMappingsRequiringReview -eq 1) 'Device correlation treated as source object'
        Get-ChildItem $inputPath -File | ForEach-Object {Assert-True ((Get-FileHash $_.FullName).Hash -eq $before[$_.Name]) 'Input changed'}
        Assert-True ((Get-FileHash (Join-Path $outputPath 'CMDB_Relationships.csv')).Hash -eq $before['CMDB_Relationships.csv']) 'Relationships rewritten'
    }
    Test-Case 'Name update and multiline CSV retain internal identity' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Users.csv'); $row.DisplayName="Renamed, `"user`"`r`nsecond line"
        Write-Row 'CMDB_Users.csv' @($row)
        Export-SmartWorkplaceCMDBCIRegistry @argsCI | Out-Null
        $user=Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv') | Where-Object CI_Type -eq User
        Assert-True ($user.CI_Name -ceq $row.DisplayName -and $user.CI_ID -ceq $row.CmdbUserId) 'Roundtrip/identity'
    }
    Test-Case 'Reject cross-tenant entity without partial destination' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Users.csv'); $row.TenantKey='foreign-test'
        Write-Row 'CMDB_Users.csv' @($row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'tenant identity mismatch'
        Assert-True (-not (Test-Path $outputPath)) 'Partial destination'
        Assert-True (@(Get-ChildItem (Split-Path $outputPath) -Force -Directory | Where-Object Name -like '.cmdb-ci-*').Count -eq 0) 'Stage leaked'
    }
    Test-Case 'Reject mismatched TenantId and invalid context' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Groups.csv'); $row.TenantId='11111111-1111-1111-1111-111111111111'
        Write-Row 'CMDB_Groups.csv' @($row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -ValidateOnly} 'identity mismatch'
        $argsCI.TenantKey='wrong-test'
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -ValidateOnly} 'TenantKey must match'
    }
    Test-Case 'Reject duplicate internal IDs' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Users.csv'); Write-Row 'CMDB_Users.csv' @($row,$row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'Duplicate CI'
    }
    Test-Case 'Reject duplicate source mappings with different internal IDs' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Users.csv'); $other=$row.PSObject.Copy(); $other.CmdbUserId='example-test|user|internal-2'
        Write-Row 'CMDB_Users.csv' @($row,$other)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'Duplicate source'
    }
    Test-Case 'Reject orphan relationship atomically' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Relationships.csv'); $row.ToEntityId='example-test|device|missing'
        Write-Row 'CMDB_Relationships.csv' @($row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'Orphan'
        Assert-True (-not (Test-Path $outputPath)) 'Partial registry published'
    }
    Test-Case 'Reject wrong relationship endpoint type' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Relationships.csv'); $row.ToEntityType='User'
        Write-Row 'CMDB_Relationships.csv' @($row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'type mismatch'
    }
    Test-Case 'Reject foreign relationship and unknown verb' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Relationships.csv'); $row.OrganizationKey='foreign'
        Write-Row 'CMDB_Relationships.csv' @($row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'Relationship tenant'
        $row.OrganizationKey='example'; $row.RelationshipType='owned_by'
        Write-Row 'CMDB_Relationships.csv' @($row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'Unknown relationship'
    }
    Test-Case 'Reject duplicate relationship ID' {
        $row=Import-Csv (Join-Path $inputPath 'CMDB_Relationships.csv'); Write-Row 'CMDB_Relationships.csv' @($row,$row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'Duplicate relationship'
    }
    Test-Case 'Relationship removed from snapshot is not resurrected' {
        Write-Row 'CMDB_Relationships.csv' @()
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
        Assert-True ($r.RelationshipCount -eq 0 -and $r.CICount -eq 5) 'Removed edge retained'
    }
    Test-Case 'Missing file fails; empty tables are valid only without orphan edges' {
        Remove-Item -LiteralPath (Join-Path $inputPath 'CMDB_Groups.csv')
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'Missing required input'
        foreach($d in $catalog.types){Write-Row $d.file @()}; Write-Row 'CMDB_Relationships.csv' @()
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
        Assert-True ($r.CICount -eq 0 -and $r.RelationshipCount -eq 0) 'Empty contract failed'
        Assert-True ((Get-Content (Join-Path $outputPath 'CMDB_ConfigurationItems.csv') -TotalCount 1) -eq ($catalog.columns -join ',')) 'Missing header'
    }
    Test-Case 'Incompatible header fails before writes' {
        'bad,header' | Set-Content (Join-Path $inputPath 'CMDB_Users.csv')
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'Incompatible input header'
    }
    Test-Case 'Existing output and nested input/output are preserved' {
        New-Item -ItemType Directory -Path $outputPath | Out-Null
        'keep' | Set-Content (Join-Path $outputPath 'sentinel.txt')
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'already exists'
        Assert-True ((Get-Content (Join-Path $outputPath 'sentinel.txt')) -eq 'keep') 'Existing output overwritten'
        $argsCI.OutputDirectory=Join-Path $inputPath 'child'
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI} 'separate from the input'
    }
    Test-Case 'Missing/composite provenance is explicit, not fabricated' {
        $mail=Import-Csv (Join-Path $inputPath 'CMDB_Mailboxes.csv'); $mail.ExternalDirectoryObjectId=''; Write-Row 'CMDB_Mailboxes.csv' @($mail)
        $device=Import-Csv (Join-Path $inputPath 'CMDB_Devices.csv'); $device.SourceSystem='Entra+Intune'; Write-Row 'CMDB_Devices.csv' @($device)
        Export-SmartWorkplaceCMDBCIRegistry @argsCI | Out-Null
        $rows=@(Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv'))
        Assert-True (($rows | Where-Object CI_Type -eq Mailbox).SourceMappingStatus -eq 'MissingSourceIdentity') 'Missing provenance hidden'
        Assert-True (($rows | Where-Object CI_Type -eq Device).SourceMappingStatus -eq 'RequiresSourceEvidence') 'Composite split guessed'
    }
    Test-Case 'Catalog is stable and cannot alias a foreign file or duplicate types' {
        $localCatalog=Join-Path (Split-Path $outputPath) 'catalog.json'
        $changed=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.ci.catalog.json') | ConvertFrom-Json
        $changed.channel='BETA'; $changed | ConvertTo-Json -Depth 10 | Set-Content $localCatalog
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -CatalogPath $localCatalog} 'frozen V1 stable'
        $changed.channel='stable'; $changed.types[0].file='../foreign.csv'; $changed | ConvertTo-Json -Depth 10 | Set-Content $localCatalog
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -CatalogPath $localCatalog} 'existing CMDB entity'
    }
    Test-Case 'Existing full fixture pipeline feeds the registry without adapters' {
        $orchestrator=Join-Path $projectRoot 'Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1'
        $fixtureRoot=Join-Path $PSScriptRoot 'Fixtures'
        $runtimeRoot=Join-Path (Split-Path $outputPath) 'runtime'
        # The existing orchestrator requires PS7 (SharePoint module), whereas
        # the registry consumer is tested in the current PS5.1 or PS7 process.
        $pwsh=(Get-Command pwsh -ErrorAction Stop).Source
        & $pwsh -NoProfile -ExecutionPolicy Bypass -File $orchestrator -Tenant ci-test -OrganizationKey example -EnvironmentKey test -TenantKey example-test -TenantId aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa -NoConfigWrite -DataRootPath $runtimeRoot -FixtureRootPath $fixtureRoot -DisableSharePointUpload | Out-Null
        Assert-True ($LASTEXITCODE -eq 0) 'PS7 fixture pipeline failed'
        $argsCI.InputRootPath=Join-Path $runtimeRoot 'DATA-LAST\CMDB'
        $argsCI.TenantId='aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
        $expected=0
        foreach($d in $catalog.types){$expected+=@(Import-Csv (Join-Path $argsCI.InputRootPath $d.file)).Count}
        $expectedRelationships=@(Import-Csv (Join-Path $argsCI.InputRootPath 'CMDB_Relationships.csv')).Count
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI
        Assert-True ($r.CICount -eq $expected -and $r.RelationshipCount -eq $expectedRelationships -and $expectedRelationships -gt 0) 'Native pipeline/registry parity failed'
        $rawRoot=Join-Path $runtimeRoot 'DATA-LAST\Raw'
        $e=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawRoot -ValidateOnly
        $expectedSources=0
        foreach ($relative in @('Entra\Entra_Users.csv','Entra\Entra_Groups.csv','Entra\Entra_Devices.csv','Intune\Intune_ManagedDevices.csv','M365\M365_SubscribedSkus.csv','ExchangeOnline\ExchangeOnline_Mailboxes.csv')) {
            $expectedSources+=@(Import-Csv (Join-Path $rawRoot $relative)).Count
        }
        Assert-True ($e.SourceEvidence.MappingCount -eq $expectedSources -and $e.SourceEvidence.CIsWithoutEvidence -eq 0) 'Native raw evidence parity failed'
        $context=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawRoot -IncludeContext -ValidateOnly
        Assert-True ($context.Context.Counts['CMDB_CIUserContext.csv'] -eq $r.CITypes.User -and $context.Context.Counts['CMDB_CIDeviceContext.csv'] -eq $r.CITypes.Device) 'Native context counts failed'
    }
    Test-Case 'Native source IDs remain distinct from CI and correlation IDs' {
        New-RawFixture
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath
        $rows=@(Import-Csv (Join-Path $outputPath 'CMDB_CISources.csv'))
        $devices=@($rows | Where-Object CI_ID -eq 'example-test|device|internal-1')
        Assert-True ($rows.Count -eq 6 -and $devices.Count -eq 2 -and $r.SourceEvidence.CIsWithoutEvidence -eq 0) 'Source mapping counts'
        Assert-True ($devices.SourceID -contains 'entra-object-1' -and $devices.SourceID -contains 'intune-managed-1') 'Native IDs conflated'
        Assert-True ($r.SourceCollectionStatus -eq 'NotAssessed') 'Freshness certified incorrectly'
    }
    Test-Case 'All Intune candidates and subscribed SKU objects retain mappings' {
        New-RawFixture
        foreach ($s in @(@('Intune_ManagedDevices.csv','ManagedDeviceId'),@('M365_SubscribedSkus.csv','SubscribedSkuId'))) {
            $a=Import-Csv (Get-RawPath $s[0]); $b=$a.PSObject.Copy(); $b.($s[1])='other-native-object'
            Write-Raw $s[0] @($a,$b)
        }
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath -ValidateOnly
        Assert-True ($r.SourceEvidence.MappingCount -eq 8 -and $r.CICount -eq 5) 'Candidates or subscriptions collapsed'
    }
    Test-Case 'Intune zero correlation uses existing fallback; mailbox needs no external ID' {
        New-RawFixture
        Write-Raw 'Entra_Devices.csv' @()
        $raw=Import-Csv (Get-RawPath 'Intune_ManagedDevices.csv'); $raw.AzureAdDeviceId='00000000-0000-0000-0000-000000000000'; Write-Raw 'Intune_ManagedDevices.csv' @($raw)
        $ci=Import-Csv (Join-Path $inputPath 'CMDB_Devices.csv'); $ci.SourceDeviceId='intune:intune-managed-1'; $ci.SourceSystem='MicrosoftIntune'; Write-Row 'CMDB_Devices.csv' @($ci)
        $mail=Import-Csv (Join-Path $inputPath 'CMDB_Mailboxes.csv'); $mail.ExternalDirectoryObjectId=''; Write-Row 'CMDB_Mailboxes.csv' @($mail)
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath -ValidateOnly
        Assert-True ($r.SourceEvidence.CIsWithoutEvidence -eq 0) 'Fallback identity lost'
    }
    Test-Case 'Reject foreign raw identity, duplicate source and orphan without partial output' {
        New-RawFixture
        $raw=Import-Csv (Get-RawPath 'Entra_Users.csv'); $raw.TenantKey='foreign-test'; Write-Raw 'Entra_Users.csv' @($raw)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'Source evidence tenant'
        $raw.TenantKey='example-test'; Write-Raw 'Entra_Users.csv' @($raw,$raw)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'Duplicate native source'
        $raw.SourceUserId='missing'; Write-Raw 'Entra_Users.csv' @($raw)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'Unmapped native'
        Assert-True (-not (Test-Path $outputPath)) 'Partial source output'
        Assert-True (@(Get-ChildItem (Split-Path $outputPath) -Force -Directory | Where-Object Name -like '.cmdb-ci-*').Count -eq 0) 'Source stage leaked'
    }
    Test-Case 'Raw missing file differs from empty evidence and missing native ID' {
        New-RawFixture
        Write-Raw 'Entra_Groups.csv' @()
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath -ValidateOnly
        Assert-True ($r.SourceEvidence.CIsWithoutEvidence -eq 1) 'Missing CI evidence hidden'
        $raw=Import-Csv (Get-RawPath 'Entra_Devices.csv'); $raw.SourceObjectId=''; Write-Raw 'Entra_Devices.csv' @($raw)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'Missing native source'
        Remove-Item -LiteralPath (Get-RawPath 'Entra_Users.csv')
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'Missing required evidence file'
    }
    Test-Case 'Reject ambiguous curated correlation before assigning source objects' {
        New-RawFixture
        $a=Import-Csv (Join-Path $inputPath 'CMDB_Devices.csv'); $b=$a.PSObject.Copy(); $b.CmdbDeviceId='example-test|device|internal-2'; Write-Row 'CMDB_Devices.csv' @($a,$b)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'Ambiguous curated'
    }
    Test-Case 'Journal replays ownership and lifecycle and preserves evidence bytes' {
        Write-Journal (New-ActiveJournal)
        $before=(Get-FileHash $journalPath).Hash
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath
        $device=Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv') | Where-Object CI_Type -eq Device
        Assert-True ($device.LifecycleStatus -eq 'Active' -and $device.OwnershipStatus -eq 'DeclaredComplete' -and $device.TechnicalOwner -eq 'example-test|group|internal-1') 'Governance state failed'
        Assert-True ($r.Governance.EventCount -eq 5 -and $r.Governance.ActorAuthentication -eq 'NotPerformed') 'False authorization claim'
        Assert-True ((Get-FileHash (Join-Path $outputPath 'CMDB_CIGovernanceChanges.csv')).Hash -eq $before) 'Journal changed'
        $again=Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath -ValidateOnly
        Assert-True ($again.Governance.EventCount -eq 5) 'Replay duplicated history'
    }
    Test-Case 'Retirement keeps CI and relationships; lifecycle alone does not declare ownership' {
        Write-Journal @((New-Change '1' 'LifecycleStatus' 'Unknown' 'Planned'),(New-Change '2' 'LifecycleStatus' 'Planned' 'Retired'))
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath
        $device=Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv') | Where-Object CI_Type -eq Device
        Assert-True ($device.LifecycleStatus -eq 'Retired' -and $device.OwnershipStatus -eq 'NotCollected' -and $r.RelationshipCount -eq 1) 'Retirement deleted/inferred data'
    }
    Test-Case 'Reject stale previous value, forbidden transition and no-op' {
        Write-Journal @((New-Change '1' 'LifecycleStatus' 'Active' 'Retired'))
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'OldValue'
        Write-Journal @((New-Change '1' 'LifecycleStatus' 'Unknown' 'Retired'))
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'Forbidden lifecycle'
        Write-Journal @((New-Change '1' 'LifecycleStatus' 'Unknown' 'Unknown'))
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'No-op'
    }
    Test-Case 'Reject activation without owners and clearing required owner after activation' {
        Write-Journal @((New-Change '1' 'LifecycleStatus' 'Unknown' 'Planned'),(New-Change '2' 'LifecycleStatus' 'Planned' 'Active'))
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'ownership requirement'
        Write-Journal @((New-ActiveJournal) + (New-Change '6' 'SupportGroup' 'example-test|group|internal-1' ''))
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'ownership requirement'
    }
    Test-Case 'Reject missing foreign and wrong-type governance references' {
        foreach ($target in @('foreign-test|group|1','example-test|group|missing','example-test|user|internal-1')) {
            Write-Journal @((New-Change '1' 'SupportGroup' '' $target))
            Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'Governance reference'
        }
        $e=New-Change '1' 'BusinessOwner' '' 'example-test|user|internal-1'; $e.CI_ID='example-test|device|missing'; Write-Journal @($e)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'missing or foreign CI'
        Assert-True (-not (Test-Path $outputPath)) 'Partial governed registry'
    }
    Test-Case 'Reject governance foreign identity missing actor reason duplicate events and unknown field' {
        $base=New-Change '1' 'BusinessOwner' '' 'example-test|user|internal-1'
        foreach ($s in @(@('TenantId','foreign'),@('ChangedBy',''),@('Reason',''),@('Field','PrimaryUser'))) {
            $e=$base.PSObject.Copy(); $e.($s[0])=$s[1]; Write-Journal @($e)
            Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'Governance tenant|Missing governance|Unknown governance'
        }
        Write-Journal @($base,$base)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'Duplicate governance'
    }
    Test-Case 'Reject ambiguous future invalid and decreasing governance dates' {
        foreach ($time in @('01/02/2026','2026-01-01T00:00:00','2099-01-01T00:00:00Z','2026-02-30T00:00:00Z')) {
            Write-Journal @((New-Change '1' 'LifecycleStatus' 'Unknown' 'Planned' $time))
            Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'governance UTC date'
        }
        Write-Journal @((New-Change '1' 'LifecycleStatus' 'Unknown' 'Planned' '2026-01-02T00:00:00Z'),(New-Change '2' 'LifecycleStatus' 'Planned' 'Retired'))
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'out of order'
    }
    Test-Case 'Custom lifecycle remains catalog driven' {
        $custom=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.ci.catalog.json') | ConvertFrom-Json
        $custom.lifecycleStatuses+=@('HeldForReview')
        $custom.lifecycleTransitions+=@([pscustomobject]@{from='Unknown';to='HeldForReview'})
        $custom.lifecycleRequirements+=@([pscustomobject]@{status='HeldForReview';fields=@()})
        $path=Join-Path (Split-Path $outputPath) 'custom.json'; $custom | ConvertTo-Json -Depth 10 | Set-Content $path
        Write-Journal @((New-Change '1' 'LifecycleStatus' 'Unknown' 'HeldForReview'))
        Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath -CatalogPath $path | Out-Null
        $device=Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv') | Where-Object CI_Type -eq Device
        Assert-True ($device.LifecycleStatus -eq 'HeldForReview') 'Custom status rejected'
    }
    Test-Case 'Empty journal and ValidateOnly preserve unknown state and create no artifacts' {
        New-RawFixture; Write-Journal @()
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath -GovernanceJournalPath $journalPath -ValidateOnly
        Assert-True ($r.Governance.EventCount -eq 0 -and $r.SourceEvidence.MappingCount -eq 6 -and -not (Test-Path $outputPath)) 'Read-only evidence validation wrote files'
    }
    Test-Case 'Reject forged raw provenance, incompatible headers and nested raw output' {
        New-RawFixture
        $row=Import-Csv (Get-RawPath 'Entra_Users.csv'); $row.SourceSystem='ExchangeOnline'; Write-Raw 'Entra_Users.csv' @($row)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'Unexpected native'
        'wrong,header' | Set-Content (Get-RawPath 'Entra_Users.csv')
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'Incompatible evidence header'
        $argsCI.OutputDirectory=Join-Path $rawPath 'output'
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath} 'separate from the raw input'
    }
    Test-Case 'Governance actor must be a known user and header must match contract' {
        $e=New-Change '1' 'BusinessOwner' '' 'example-test|user|internal-1'; $e.ChangedBy='example-test|group|internal-1'; Write-Journal @($e)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'Governance reference'
        'wrong,header' | Set-Content $journalPath
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'Incompatible evidence header'
    }
    Test-Case 'Lifecycle policy rejects unknown states endpoints and required fields' {
        $path=Join-Path (Split-Path $outputPath) 'invalid-catalog.json'
        foreach ($case in @('transition','field','missing')) {
            $custom=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.ci.catalog.json') | ConvertFrom-Json
            if ($case -eq 'transition') {$custom.lifecycleTransitions[0].to='Undefined'}
            elseif ($case -eq 'field') {$custom.lifecycleRequirements[2].fields=@('PrimaryUser')}
            else {$custom.lifecycleRequirements=@($custom.lifecycleRequirements | Where-Object status -ne 'Active')}
            $custom | ConvertTo-Json -Depth 10 | Set-Content $path
            Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -CatalogPath $path -ValidateOnly} 'Invalid lifecycle|Invalid required|explicit requirement'
        }
    }
    Test-Case 'Context preserves associated user department without declaring device organization' {
        $user=Import-Csv (Join-Path $inputPath 'CMDB_Users.csv'); $user.Department="Example, unit`r`nsecond line"; $user.JobTitle='Example role'; Write-Row 'CMDB_Users.csv' @($user)
        $device=Import-Csv (Join-Path $inputPath 'CMDB_Devices.csv'); $device.PrimaryUserId=$user.SourceUserId; Write-Row 'CMDB_Devices.csv' @($device)
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext
        $row=Import-Csv (Join-Path $outputPath 'CMDB_CIDeviceContext.csv')
        Assert-True ($row.AssociatedUserCI_ID -ceq $user.CmdbUserId -and $row.AssociatedUserDepartment -ceq $user.Department -and $row.AssociationStatus -ceq 'Resolved') 'Context association/roundtrip failed'
        Assert-True ($r.Context.Counts['CMDB_CIOrganization.csv'] -eq 0 -and $r.Context.NativeEvidence -eq 'NotProvided') 'Organization or native evidence inferred'
        $ci=Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv') | Where-Object CI_Type -eq Device
        Assert-True ($ci.BusinessOwner -eq '' -and $ci.OwnershipStatus -eq 'NotCollected') 'Association became ownership'
    }
    Test-Case 'Missing user and unknown association remain explicit' {
        $device=Import-Csv (Join-Path $inputPath 'CMDB_Devices.csv'); $device.PrimaryUserId='not-collected-user'; Write-Row 'CMDB_Devices.csv' @($device)
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext
        $row=Import-Csv (Join-Path $outputPath 'CMDB_CIDeviceContext.csv')
        Assert-True ($row.AssociationStatus -eq 'Unresolved' -and $row.AssociatedUserCI_ID -eq '' -and $r.Context.UsersMissingDepartment -eq 1) 'Missing association hidden'
    }
    Test-Case 'No primary user is distinct from unresolved user' {
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext -ValidateOnly
        Assert-True ($r.Context.DeviceAssociations.NotProvided -eq 1 -and -not (Test-Path $outputPath)) 'Missing association or ValidateOnly incorrect'
    }
    Test-Case 'Raw device context retains candidates, dates and ambiguity' {
        New-RawFixture
        $entra=Import-Csv (Get-RawPath 'Entra_Devices.csv'); $entra.ApproximateLastSignInDateTime='09/01/2026 12:00:00'; Write-Raw 'Entra_Devices.csv' @($entra)
        $intune=Import-Csv (Get-RawPath 'Intune_ManagedDevices.csv'); $intune.LastSyncDateTime='2026-01-01T12:00:00+02:00'; $intune.EnrolledDateTime='2025-02-01T10:00:00Z'; $intune.ManagementAgent='mdm'; $intune.DeviceEnrollmentType='exampleEnrollment'
        $candidate=$intune.PSObject.Copy(); $candidate.ManagedDeviceId='other-candidate'; Write-Raw 'Intune_ManagedDevices.csv' @($intune,$candidate)
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath -IncludeContext
        $rows=@(Import-Csv (Join-Path $outputPath 'CMDB_CIDeviceSourceContext.csv'))
        $e=$rows | Where-Object SourceSystem -eq MicrosoftEntraID
        $i=@($rows | Where-Object SourceSystem -eq MicrosoftIntune)
        Assert-True ($rows.Count -eq 3 -and $r.Context.Counts['CMDB_CIDeviceSourceContext.csv'] -eq 3) 'Candidate discarded'
        Assert-True ($e.ActivityStatus -eq 'NotOffsetQualified' -and $e.ActivityUtcDateTime -eq '' -and $e.ActivityRaw -eq $entra.ApproximateLastSignInDateTime -and $e.EnrollmentStatus -eq 'NotApplicable') 'Ambiguous date guessed'
        Assert-True ($i[0].ActivityUtcDateTime -eq '2026-01-01T10:00:00.0000000Z' -and $i[0].ManagementAgent -eq 'mdm' -and $i[0].EnrollmentType -eq 'exampleEnrollment') 'Native detail/date lost'
    }
    Test-Case 'Invalid qualified and empty activity dates are not promoted to UTC' {
        New-RawFixture
        $intune=Import-Csv (Get-RawPath 'Intune_ManagedDevices.csv'); $intune.LastSyncDateTime='2026-02-30T12:00:00Z'; Write-Raw 'Intune_ManagedDevices.csv' @($intune)
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -RawRootPath $rawPath -IncludeContext -ValidateOnly
        Assert-True ($r.Context.ActivityQualification['MicrosoftIntune:Invalid'] -eq 1 -and $r.Context.ActivityQualification['MicrosoftEntraID:Missing'] -eq 1) 'Invalid/missing dates hidden'
    }
    Test-Case 'Declared site resolves its hierarchy through the existing journal' {
        New-OrganizationFixture
        Write-Journal @((New-Change 'org1' 'OrganizationRefId' '' 'example-test|site|1'))
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext -OrganizationReferencePath $organizationPath -GovernanceJournalPath $journalPath
        $row=Import-Csv (Join-Path $outputPath 'CMDB_CIOrganization.csv')
        Assert-True ($row.CountryRefId -eq 'example-test|country|1' -and $row.EntityRefId -eq 'example-test|entity|1' -and $row.SiteRefId -eq 'example-test|site|1') 'Hierarchy lost'
        Assert-True ($r.Context.OrganizationReferenceCount -eq 3 -and $r.Governance.EventCount -eq 1) 'Governed counts incorrect'
        $ci=Import-Csv (Join-Path $outputPath 'CMDB_ConfigurationItems.csv') | Where-Object CI_Type -eq Device
        Assert-True ($ci.OwnershipStatus -eq 'NotCollected') 'Organization falsely declared ownership'
        Assert-True ((Get-FileHash $organizationPath).Hash -eq (Get-FileHash (Join-Path $outputPath 'CMDB_OrganizationReferences.csv')).Hash) 'Reference bytes changed'
    }
    Test-Case 'Clearing organization retains journal and removes only current assignment' {
        New-OrganizationFixture
        Write-Journal @((New-Change 'org1' 'OrganizationRefId' '' 'example-test|site|1'),(New-Change 'org2' 'OrganizationRefId' 'example-test|site|1' ''))
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext -OrganizationReferencePath $organizationPath -GovernanceJournalPath $journalPath
        Assert-True ($r.Context.Counts['CMDB_CIOrganization.csv'] -eq 0 -and $r.Governance.EventCount -eq 2 -and $r.RelationshipCount -eq 1) 'Clear erased history/relationship'
    }
    Test-Case 'Organization declarations cannot be silently omitted without context' {
        New-OrganizationFixture
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -OrganizationReferencePath $organizationPath} 'require IncludeContext'
        Write-Journal @((New-Change 'org1' 'OrganizationRefId' '' 'example-test|site|1'))
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -GovernanceJournalPath $journalPath} 'require IncludeContext'
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext -GovernanceJournalPath $journalPath} 'organization reference is missing'
        Assert-True (-not (Test-Path $outputPath)) 'Partial organizational output'
        Assert-True (@(Get-ChildItem (Split-Path $outputPath) -Force -Directory | Where-Object Name -like '.cmdb-ci-*').Count -eq 0) 'Context stage leaked'
    }
    Test-Case 'Reject foreign duplicate orphan cyclic and wrong-type organization references' {
        foreach ($case in @('foreign','duplicate','orphan','cycle','parentType','emptyName','type')) {
            New-OrganizationFixture
            switch ($case) {
                foreign {$organizationRows[0].TenantKey='other-test'}
                duplicate {$script:organizationRows+=@($script:organizationRows[0])}
                orphan {$organizationRows[2].ParentReferenceId='example-test|entity|missing'}
                cycle {$organizationRows[0].ParentReferenceId='example-test|site|1'}
                parentType {$organizationRows[2].ParentReferenceId='example-test|country|1'}
                emptyName {$organizationRows[0].Name=''}
                type {$organizationRows[0].ReferenceType='Department'}
            }
            Write-Organization $organizationRows
            try { Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext -OrganizationReferencePath $organizationPath} 'reference tenant|Duplicate organization|organization parent|cannot have a parent|identity or name|Unknown organization' }
            catch { throw "Organization case '$case': $($_.Exception.Message)" }
        }
    }
    Test-Case 'Historical organizational targets must remain resolvable' {
        New-OrganizationFixture
        Write-Journal @((New-Change 'org1' 'OrganizationRefId' '' 'example-test|site|1'),(New-Change 'org2' 'OrganizationRefId' 'example-test|site|1' ''))
        Write-Organization @($organizationRows | Where-Object ReferenceType -ne Site)
        Assert-Throws {Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext -OrganizationReferencePath $organizationPath -GovernanceJournalPath $journalPath} 'organization reference is missing'
    }
    Test-Case 'Empty context contracts produce headers, no invented entities or references' {
        foreach ($d in $catalog.types) {Write-Row $d.file @()}; Write-Row 'CMDB_Relationships.csv' @()
        New-OrganizationFixture; Write-Organization @()
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -IncludeContext -OrganizationReferencePath $organizationPath
        $schema=Get-Content -Raw (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.ci.context.json') | ConvertFrom-Json
        foreach ($table in $schema.tables) {Assert-True ((Get-Content (Join-Path $outputPath $table.name) -TotalCount 1) -eq ($table.columns -join ',')) 'Empty context header mismatch'}
        Assert-True ($r.Context.OrganizationReferenceCount -eq 0) 'Invented references'
    }
    . (Join-Path $PSScriptRoot 'CIHardware.Cases.ps1')
    Test-Case 'Two thousand rows do not change graph or source-key semantics' {
        $base=Import-Csv (Join-Path $inputPath 'CMDB_Users.csv')
        $rows=@(foreach($i in 1..2000){$r=$base.PSObject.Copy(); $r.CmdbUserId='example-test|user|internal-'+$i; $r.SourceUserId='source-user-'+$i; $r})
        Write-Row 'CMDB_Users.csv' $rows
        $r=Export-SmartWorkplaceCMDBCIRegistry @argsCI -ValidateOnly
        Assert-True ($r.CICount -eq 2004 -and $r.RelationshipCount -eq 1) 'Volume counts mismatch'
    }
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    if(-not $resolved.StartsWith($tempParent+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'cmdb-ci-tests-*'){throw 'Unsafe test cleanup path'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
Write-Output "CI registry tests: $passed passed, $failed failed."
if($failed){exit 1}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAIlH01jEJKvyq7
# eg6Dh1Z+K/XaH5G5DdS2cBx74z5knaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJjQ4OqkHnyVGlJ/xBWFb8W6pDTaY8mtskeNvBRsAs2RMA0GCSqG
# SIb3DQEBAQUABIIBgCk1dXmDw0KBYjfUfqO3FB5XSiomhnDb/mn5mfQjPYREoY+E
# GkTpKMiClY0ObOuzG8GtkvdrBNAS9RFlNOdWMyfSjTOBzzpAY04pkwlQ3T+8zVXK
# xYYT2lhxHC+vje+R+74kOOtCghnHkQXnEKYwtcDI2YeRzhUBeRLK9mwn6IsDLd7h
# 1tOXbCxPb+zFpfRRcAnd7GIufiJ9iU6n1Fo+kNzoA2CEbqOnmaER+WL50Vym+JF1
# UrXYTsJOLJMhBGc5mPajpJ8a03SDmcII/u83n6Agqm1gCXQFc3Tuz8rUC0sfuC0B
# QKgQ6XuDjLzcBimNxwoaPlq/iRg+d/ebOFIESwN699P51nQ2T3IUhUAo9qL63VdY
# 1bZBN90E0/WWfITaDwZOUyQZebt1v/TcPQKeI9soz4DO4QWuImlviWNocXGbQdAQ
# KnhoINaNAqPq1DxItqP4REZDL1UaPDhms78uqPwm4EiWFgTnkbNliHlQChsUk1kF
# Q7QCj4AKHQVJdGqj8aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MTIxMzI1
# MjRaMC8GCSqGSIb3DQEJBDEiBCCcfgwJDneUCdHyWJPqRg06rknvxZUfI/zc5qYO
# 4JSP2jANBgkqhkiG9w0BAQEFAASCAgABfqP74oGN0L3wflQtZNN3YF/9jwuIIJBp
# XdwxlgkIey6pCcCfrhCN12wZAi8RGg2d1hYdRh1s7PhvRfU4MnC5t4UpB79nHF+K
# qUM+XHv+EOo2QTpOXkW3Scu/LhiHca3HysWdlDDPvD1zQjHEaH4keBfrZhL6riqI
# LQC9js+Cg7xrX2dX9xuM3X/GhcMxeDOYCTZ1H7wb+DZJbfT66Vy33nEP45sdYuQ2
# Y/3yzyfkfgG0Qam1jeNaenmOftYXLcOZqi4TbwkGE7xls5+XgHxQIUKAmF2UouEl
# +Hk06wUocigIfNbiZ9FjfiTRPUas82fo6vs0X1PcGbBYj5SpXqhwcvII6dNXNhXg
# AqhU5aCsCgF7OW6pRhUGFVbqhJOG9k+oAYUQyEqDxuDU8IPIp+U6W3/KTTRid+S8
# hPs/MssK+qhrvR5BA/9Jj5nNuHQRjs3NgApc7GO+erkEgjxNlvUtMK6KM+cIdKnP
# rJMbTz8PzmdlCkjAiyO0zP/eqDN+qdsULHxabIT/6AAHCjTyZ4NtbOlMsXQRV7U0
# NAQ4TMpweG4Sged2xiwdCVaYxTAqqdwuUOM/gfTtDZapxcCSeehF8kSsRcSSZ/MF
# icMFD80u/Q9Crke5oeiQoBM/q0WgnOMEl2jz92aPj57TRVF2XKTEOH+65AVXAA8S
# W/b8Avri6Q==
# SIG # End signature block
