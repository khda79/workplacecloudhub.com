# Stable V1 additive registry; never collects, infers owners, or rewrites source IDs.
Set-StrictMode -Version 2.0
. (Join-Path $PSScriptRoot 'CI.Governance.ps1')
. (Join-Path $PSScriptRoot 'CI.Sources.ps1')
. (Join-Path $PSScriptRoot 'CI.Context.ps1')
. (Join-Path $PSScriptRoot 'CI.Hardware.ps1')

function Export-SmartWorkplaceCMDBCIRegistry {
    <#
    .SYNOPSIS
    Validates or exports a common CI registry from existing curated CSVs.
    .DESCRIPTION
    InputRootPath is the CMDB directory, not a Power BI ReportData directory.
    All five current entity files and CMDB_Relationships.csv are required.
    Writes only to a new OutputDirectory. ValidateOnly performs no writes.
    Existing CMDB relationships are validated and copied without rewriting.
    RawRootPath optionally supplies the corresponding Raw snapshot directory.
    GovernanceJournalPath optionally supplies a complete local change journal.
    These inputs declare evidence, not authenticated approval or source freshness.
    IncludeContext adds typed device/user context. OrganizationReferencePath
    supplies a controlled hierarchy for journal OrganizationRefId declarations.
    HardwareInputPath optionally supplies the separate Intune hardware snapshot
    and required state sidecar. RawRootPath supplies its native ID mapping.
    .EXAMPLE
    Export-SmartWorkplaceCMDBCIRegistry -InputRootPath C:\Example\CMDB -OrganizationKey example -EnvironmentKey test -TenantKey example-test -ValidateOnly
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InputRootPath,
        [Parameter(Mandatory)][string]$OrganizationKey,
        [Parameter(Mandatory)][string]$EnvironmentKey,
        [Parameter(Mandatory)][string]$TenantKey,
        [string]$TenantId = '',
        [string]$OutputDirectory,
        [string]$CatalogPath,
        [string]$RawRootPath,
        [string]$GovernanceJournalPath,
        [switch]$IncludeContext,
        [string]$OrganizationReferencePath,
        [string]$HardwareInputPath,
        [switch]$ValidateOnly
    )
    $ErrorActionPreference = 'Stop'
    if ($HardwareInputPath -and -not $RawRootPath) { throw 'HardwareInputPath requires RawRootPath for native ID mapping.' }
    $projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    if (-not $CatalogPath) { $CatalogPath = Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.ci.catalog.json' }
    $catalogHash=(Get-FileHash -LiteralPath $CatalogPath -Algorithm SHA256).Hash
    $catalog = Get-Content -Raw -LiteralPath $CatalogPath | ConvertFrom-Json
    $contract = Get-Content -Raw -LiteralPath (Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.tables.json') | ConvertFrom-Json
    if ($catalog.channel -cne 'stable' -or $catalog.contractVersion -cne '1.0.0') { throw 'CI catalog must be the frozen V1 stable contract.' }
    foreach ($component in @($OrganizationKey,$EnvironmentKey)) {
        if ($component.Length -gt 64 -or $component -cnotmatch '^[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$') { throw 'Invalid tenant identity component.' }
    }
    if ($TenantKey -cne ($OrganizationKey+'-'+$EnvironmentKey)) { throw 'TenantKey must match organization and environment.' }
    if ($TenantId) {
        $parsedGuid = [guid]::Empty
        if (-not [guid]::TryParse($TenantId,[ref]$parsedGuid)) { throw 'Invalid TenantId.' }
    }
    $identity = [ordered]@{TenantKey=$TenantKey;OrganizationKey=$OrganizationKey;EnvironmentKey=$EnvironmentKey;TenantId=$TenantId}
    $inputRoot = [IO.Path]::GetFullPath($InputRootPath).TrimEnd('\','/')
    if (-not (Test-Path -LiteralPath $inputRoot -PathType Container)) { throw 'Input CMDB directory does not exist.' }
    $output = $null
    if (-not $ValidateOnly) {
        if (-not $OutputDirectory) { throw 'A new OutputDirectory is required.' }
        $output = [IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\','/')
        if ($output -eq $inputRoot -or $output.StartsWith($inputRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or
            $inputRoot.StartsWith($output+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Output must be separate from the input tree.' }
        if (Test-Path -LiteralPath $output) { throw 'Output directory already exists; nothing will be replaced.' }
        if (-not (Test-Path -LiteralPath (Split-Path $output -Parent) -PathType Container)) { throw 'Output parent must already exist.' }
        if ($RawRootPath) {
            $rawRoot=[IO.Path]::GetFullPath($RawRootPath).TrimEnd('\','/')
            if ($output -eq $rawRoot -or $output.StartsWith($rawRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) -or
                $rawRoot.StartsWith($output+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Output must be separate from the raw input tree.' }
        }
    }
    $types=@{}; $files=@{}; $definitions=@{}; $inputHashes=@{}
    foreach ($definition in @($catalog.types)) {
        $type=[string]$definition.name; $file=[string]$definition.file
        if ([string]::IsNullOrWhiteSpace($type) -or $types.ContainsKey($type) -or $files.ContainsKey($file)) { throw 'Duplicate or empty CI type/file.' }
        $table=@($contract.tables | Where-Object { $_.name -ceq $file -and $_.area -ceq 'CMDB' -and $_.role -ceq 'entity' })
        if ($table.Count -ne 1) { throw 'CI type must reference one existing CMDB entity contract.' }
        foreach ($field in @($definition.idColumn,$definition.nameColumn,$definition.sourceIdColumn,'SourceSystem','SourceCollectedDateTime')) {
            if ($field -cnotin $table[0].columns) { throw 'CI projection references an undeclared source column.' }
        }
        if ($definition.sourceMappingStatus -cnotin @('SourceIdentity','RequiresSourceEvidence') -or [string]::IsNullOrWhiteSpace($definition.sourceIdKind)) { throw 'Invalid source mapping classification.' }
        $types[$type]=$definition; $files[$file]=$true; $definitions[$file]=$table[0]
    }
    # First-slice adapters must cover exactly the current five entities.
    $entityFiles=@($contract.tables | Where-Object { $_.area -ceq 'CMDB' -and $_.role -ceq 'entity' } | ForEach-Object { $_.name })
    if ($types.Count -ne $entityFiles.Count -or @($entityFiles | Where-Object { -not $files.ContainsKey($_) }).Count) { throw 'CI catalog must cover every current entity exactly once.' }
    $relationsFile='CMDB_Relationships.csv'
    $definitions[$relationsFile]=@($contract.tables | Where-Object name -ceq $relationsFile)[0]
    $relationTypes=@{}
    foreach ($r in @($catalog.relationshipTypes)) {
        if ([string]::IsNullOrWhiteSpace($r.name) -or $relationTypes.ContainsKey([string]$r.name) -or -not $types.ContainsKey([string]$r.fromType) -or -not $types.ContainsKey([string]$r.toType)) { throw 'Invalid or duplicate relationship type.' }
        $relationTypes[[string]$r.name]=$r
    }
    if ($catalog.unknownLifecycleStatus -cnotin $catalog.lifecycleStatuses -or $catalog.unknownLifecycleStatus -cne 'Unknown' -or $catalog.unknownOwnershipStatus -cne 'NotCollected') { throw 'Missing governance must remain explicitly unknown.' }
    $expectedColumns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','CI_ID','CI_Name','CI_Type','Environment','BusinessOwner','TechnicalOwner','SupportGroup','OwnershipStatus','LifecycleStatus','SourceSystem','SourceID','SourceIDKind','SourceMappingStatus','SourceCollectedDateTime')
    if (($catalog.columns -join ',') -cne ($expectedColumns -join ',')) { throw 'Unsupported registry column contract.' }
    foreach ($file in $definitions.Keys) {
        $path=Join-Path $inputRoot $file
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing required input: $file" }
        $header=Get-Content -LiteralPath $path -TotalCount 1
        $actual=@($header.Split(',') | ForEach-Object { $_.Trim().Trim('"') })
        if (($actual -join ',') -cne ($definitions[$file].columns -join ',')) { throw "Incompatible input header: $file" }
        $inputHashes[$file]=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
    $sourceColumns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','CI_ID','SourceSystem','SourceID','SourceIDKind','CorrelationKey','MappingStatus','SourceCollectedDateTime')
    if (($catalog.sourceColumns -join ',') -cne ($sourceColumns -join ',')) { throw 'Unsupported source evidence column contract.' }
    $governance=Read-CIGovernanceJournal -Path $GovernanceJournalPath -Catalog $catalog -Identity $identity
    if (-not $IncludeContext -and ($OrganizationReferencePath -or $governance.OrganizationReferences.Count)) { throw 'Organization references require IncludeContext.' }
    $context=$null; $contextContract=$null; $contextHash=''; $organization=$null
    $contextPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.ci.context.json'
    if ($IncludeContext) {
        $contextHash=(Get-FileHash -LiteralPath $contextPath -Algorithm SHA256).Hash
        $contextContract=Get-Content -LiteralPath $contextPath -Raw | ConvertFrom-Json
        $referenceColumns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId','ReferenceId','ReferenceType','Name','ParentReferenceId')
        if (($contextContract.referenceColumns -join ',') -cne ($referenceColumns -join ',')) { throw 'Unsupported organization reference contract.' }
        $organization=Read-CIOrganizationReference -Path $OrganizationReferencePath -Identity $identity -Columns $referenceColumns
    }
    $evidence=@{Count=0;ByCI=@{};Hashes=@{}}
    $hardware=[ordered]@{Status='NotProvided'}
    if ($HardwareInputPath) { $evidence.HardwareIndex=@{} }
    $rawContract=$null; $rawContractPath=$null; $rawContractHash=''
    if ($RawRootPath) {
        $rawContractPath=Join-Path $projectRoot 'Schema\SmartWorkplaceCMDB.raw.tables.json'
        $rawContractHash=(Get-FileHash -LiteralPath $rawContractPath -Algorithm SHA256).Hash
        $rawContract=Get-Content -LiteralPath $rawContractPath -Raw | ConvertFrom-Json
    }
    # Retain indexes and sparse governance state, not complete source/entity rows.
    $ciIndex=@{}; $sourceIndex=@{}; $relationIds=@{}; $counts=[ordered]@{}
    $correlationIndex=@{}; $ciSystems=@{}
    $stage=$null; $writer=$null; $sourceWriter=$null; $relationCount=0; $unresolvedSources=0
    try {
        if (-not $ValidateOnly) {
            $stage=Join-Path (Split-Path $output -Parent) ('.cmdb-ci-'+[guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $stage | Out-Null
            $writer=[IO.StreamWriter]::new((Join-Path $stage 'CMDB_ConfigurationItems.csv'),$false,[Text.UTF8Encoding]::new($true))
            $writer.WriteLine($expectedColumns -join ',')
        }
        if ($IncludeContext) { $context=Initialize-CIContext -Contract $contextContract -Identity $identity -Stage $stage -ValidateOnly ([bool]$ValidateOnly) }
        foreach ($definition in @($catalog.types)) {
            $type=[string]$definition.name; $counts[$type]=0
            Import-Csv -LiteralPath (Join-Path $inputRoot $definition.file) | ForEach-Object {
                $row=$_
                foreach ($field in $identity.Keys) {
                    if ([string]$row.$field -cne [string]$identity[$field]) { throw 'CI tenant identity mismatch.' }
                }
                $id=[string]$row.($definition.idColumn)
                if ([string]::IsNullOrWhiteSpace($id) -or -not $id.StartsWith($TenantKey+'|',[StringComparison]::Ordinal)) { throw 'CI internal ID is empty or outside its tenant namespace.' }
                if ($ciIndex.ContainsKey($id)) { throw 'Duplicate CI internal ID.' }
                $ciIndex[$id]=$type
                $system=[string]$row.SourceSystem; $sourceId=[string]$row.($definition.sourceIdColumn)
                if ($RawRootPath) {
                    $ciSystems[$id]=$system
                    if ($type -cne 'Mailbox' -and -not [string]::IsNullOrWhiteSpace($sourceId)) {
                        $joinKey=ConvertTo-Json -InputObject @($type,$sourceId.Trim().ToLowerInvariant()) -Compress
                        if ($correlationIndex.ContainsKey($joinKey)) { throw 'Ambiguous curated source correlation.' }
                        $correlationIndex[$joinKey]=$id
                    }
                }
                $mapping=[string]$definition.sourceMappingStatus
                if ([string]::IsNullOrWhiteSpace($system) -or [string]::IsNullOrWhiteSpace($sourceId)) { $mapping='MissingSourceIdentity' }
                elseif ($system.Contains('+')) { $mapping='RequiresSourceEvidence' }
                if ($mapping -ceq 'SourceIdentity') {
                    $sourceKey=ConvertTo-Json -InputObject @($TenantKey,$system,$definition.sourceIdKind,$sourceId) -Compress
                    if ($sourceIndex.ContainsKey($sourceKey)) { throw 'Duplicate source identity mapping.' }
                    $sourceIndex[$sourceKey]=$id
                } else { $unresolvedSources++ }
                $ci=[ordered]@{}
                foreach ($field in $identity.Keys) { $ci[$field]=$identity[$field] }
                $ci.CI_ID=$id; $ci.CI_Name=[string]$row.($definition.nameColumn); $ci.CI_Type=$type
                $ci.Environment=$EnvironmentKey
                $ci.BusinessOwner=''; $ci.TechnicalOwner=''; $ci.SupportGroup=''
                $ci.OwnershipStatus=[string]$catalog.unknownOwnershipStatus
                $ci.LifecycleStatus=[string]$catalog.unknownLifecycleStatus
                if ($governance.Items.ContainsKey($id)) {
                    foreach ($field in @('BusinessOwner','TechnicalOwner','SupportGroup','OwnershipStatus','LifecycleStatus')) {
                        $ci[$field]=$governance.Items[$id][$field]
                    }
                }
                $ci.SourceSystem=$system; $ci.SourceID=$sourceId; $ci.SourceIDKind=[string]$definition.sourceIdKind
                $ci.SourceMappingStatus=$mapping; $ci.SourceCollectedDateTime=[string]$row.SourceCollectedDateTime
                if ($null -ne $writer) { $writer.WriteLine((@([pscustomobject]$ci | ConvertTo-Csv -NoTypeInformation)[1])) }
                $counts[$type]++
            }
        }
        Test-CIGovernanceReference -Governance $governance -CIIndex $ciIndex
        if ($IncludeContext) { Write-CICuratedContext -Context $context -InputRoot $inputRoot -CIIndex $ciIndex -Governance $governance -Organization $organization }
        if ($RawRootPath) {
            if (-not $ValidateOnly) {
                $sourceWriter=[IO.StreamWriter]::new((Join-Path $stage 'CMDB_CISources.csv'),$false,[Text.UTF8Encoding]::new($true))
                $sourceWriter.WriteLine($sourceColumns -join ',')
            }
            Get-CISourceEvidence -RawRootPath $RawRootPath -RawContract $rawContract -Identity $identity -CIIndex $ciIndex -CISystems $ciSystems -CorrelationIndex $correlationIndex -State $evidence -Context $context | ForEach-Object {
                if ($null -ne $sourceWriter) { $sourceWriter.WriteLine((@($_ | ConvertTo-Csv -NoTypeInformation)[1])) }
            }
            if ($null -ne $sourceWriter) { $sourceWriter.Dispose(); $sourceWriter=$null }
        }
        if ($HardwareInputPath) {
            $hardware=Export-CIHardwareContext -Path $HardwareInputPath -Identity $identity -Mapping $evidence.HardwareIndex -ProjectRoot $projectRoot -Stage $stage -ValidateOnly ([bool]$ValidateOnly)
            $hardware.DeviceCIsWithoutHardware=$counts.Device-$hardware.CIsWithHardware
        }
        Import-Csv -LiteralPath (Join-Path $inputRoot $relationsFile) | ForEach-Object {
            $row=$_
            foreach ($field in $identity.Keys) { if ([string]$row.$field -cne [string]$identity[$field]) { throw 'Relationship tenant identity mismatch.' } }
            $id=[string]$row.CmdbRelationshipId
            if ([string]::IsNullOrWhiteSpace($id) -or -not $id.StartsWith($TenantKey+'|',[StringComparison]::Ordinal)) { throw 'Invalid relationship ID namespace.' }
            if ($relationIds.ContainsKey($id)) { throw 'Duplicate relationship ID.' }
            $relationIds[$id]=$true
            $verb=[string]$row.RelationshipType
            if (-not $relationTypes.ContainsKey($verb)) { throw 'Unknown relationship type.' }
            $rule=$relationTypes[$verb]
            foreach ($side in @('From','To')) {
                $endpoint=[string]$row.PSObject.Properties[$side+'EntityId'].Value
                $kind=[string]$row.PSObject.Properties[$side+'EntityType'].Value
                if (-not $ciIndex.ContainsKey($endpoint)) { throw 'Orphan relationship endpoint.' }
                $expectedType=[string]$rule.PSObject.Properties[$side.ToLowerInvariant()+'Type'].Value
                if ($ciIndex[$endpoint] -cne $kind -or $kind -cne $expectedType) { throw 'Relationship endpoint type mismatch.' }
            }
            $relationCount++
        }
        if ($null -ne $writer) { $writer.Dispose(); $writer=$null }
        Close-CIContext -Context $context
        if ($IncludeContext -and (Get-FileHash -LiteralPath $contextPath -Algorithm SHA256).Hash -cne $contextHash) { throw 'Context contract changed during build.' }
        if ($OrganizationReferencePath) {
            if (-not $ValidateOnly) { Copy-Item -LiteralPath $OrganizationReferencePath -Destination (Join-Path $stage 'CMDB_OrganizationReferences.csv') }
            if ((Get-FileHash -LiteralPath $OrganizationReferencePath -Algorithm SHA256).Hash -cne $organization.Hash -or
                (-not $ValidateOnly -and (Get-FileHash -LiteralPath (Join-Path $stage 'CMDB_OrganizationReferences.csv') -Algorithm SHA256).Hash -cne $organization.Hash)) { throw 'Organization references changed during build.' }
        }
        if (-not $ValidateOnly) { Copy-Item -LiteralPath (Join-Path $inputRoot $relationsFile) -Destination (Join-Path $stage $relationsFile) }
        if ($GovernanceJournalPath) {
            if (-not $ValidateOnly) { Copy-Item -LiteralPath $GovernanceJournalPath -Destination (Join-Path $stage 'CMDB_CIGovernanceChanges.csv') }
            if ((Get-FileHash -LiteralPath $GovernanceJournalPath -Algorithm SHA256).Hash -cne $governance.Hash -or
                (-not $ValidateOnly -and (Get-FileHash -LiteralPath (Join-Path $stage 'CMDB_CIGovernanceChanges.csv') -Algorithm SHA256).Hash -cne $governance.Hash)) { throw 'Governance journal changed during build.' }
        }
        foreach ($path in $evidence.Hashes.Keys) {
            if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $evidence.Hashes[$path]) { throw 'Raw evidence changed during build.' }
        }
        if ($HardwareInputPath) {
            foreach ($path in $hardware.InputHashes.Keys) {
                if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $hardware.InputHashes[$path]) { throw 'Hardware evidence changed during registry build.' }
            }
        }
        if ($RawRootPath -and (Get-FileHash -LiteralPath $rawContractPath -Algorithm SHA256).Hash -cne $rawContractHash) { throw 'Raw contract changed during build.' }
        foreach ($file in $inputHashes.Keys) {
            if ((Get-FileHash -LiteralPath (Join-Path $inputRoot $file) -Algorithm SHA256).Hash -cne $inputHashes[$file]) { throw 'Input changed during registry build.' }
        }
        if (-not $ValidateOnly -and (Get-FileHash -LiteralPath (Join-Path $stage $relationsFile) -Algorithm SHA256).Hash -cne $inputHashes[$relationsFile]) { throw 'Copied relationships differ from validated input.' }
        if ((Get-FileHash -LiteralPath $CatalogPath -Algorithm SHA256).Hash -cne $catalogHash) { throw 'Catalog changed during registry build.' }
        $summary=[pscustomobject][ordered]@{Status='Validated';Channel='stable';ContractVersion=$catalog.contractVersion;TenantKey=$TenantKey;CICount=$ciIndex.Count;CITypes=$counts;RelationshipCount=$relationCount;SourceMappingsRequiringReview=$unresolvedSources;SourceCollectionStatus='NotAssessed';InputHashes=$inputHashes;CatalogSHA256=$catalogHash}
        $summary | Add-Member -NotePropertyName Hardware -NotePropertyValue $hardware
        $summary | Add-Member -NotePropertyName SourceEvidence -NotePropertyValue ([ordered]@{
            Status=$(if ($RawRootPath) {'Validated'} else {'NotProvided'});MappingCount=$evidence.Count
            CIsWithEvidence=$evidence.ByCI.Count;CIsWithoutEvidence=$(if ($RawRootPath) {$ciIndex.Count-$evidence.ByCI.Count} else {$null})
            InputHashes=$evidence.Hashes;RawContractSHA256=$rawContractHash
        })
        $summary | Add-Member -NotePropertyName Governance -NotePropertyValue ([ordered]@{
            Status=$(if ($GovernanceJournalPath) {'ValidatedDeclarations'} else {'NotProvided'})
            EventCount=$governance.Count;CICount=$governance.Items.Count;JournalSHA256=$governance.Hash;ActorAuthentication='NotPerformed'
        })
        if ($IncludeContext) {
            $summary | Add-Member -NotePropertyName Context -NotePropertyValue ([ordered]@{
                Status='Validated';ContractVersion=$contextContract.contractVersion;ContractSHA256=$contextHash;Counts=$context.Counts
                DeviceAssociations=$context.AssociationCounts;UsersMissingDepartment=$context.DepartmentMissing
                ActivityQualification=$context.ActivityQualification;EnrollmentQualification=$context.EnrollmentQualification
                NativeEvidence=$(if ($RawRootPath) {'Provided'} else {'NotProvided'})
                OrganizationReferenceCount=$organization.Items.Count;OrganizationReferenceSHA256=$organization.Hash
            })
        }
        if (-not $ValidateOnly) {
            $summary.Status='Exported'
            $summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $stage 'CIRegistry.manifest.json.txt') -Encoding UTF8
            # Directory.Move never replaces an existing destination, including races.
            [IO.Directory]::Move($stage,$output)
            $stage=$null
        }
        return $summary
    } finally {
        Close-CIContext -Context $context
        if ($null -ne $writer) { $writer.Dispose() }
        if ($null -ne $sourceWriter) { $sourceWriter.Dispose() }
        if ($stage -and (Test-Path -LiteralPath $stage)) {
            # Remove only files created by this invocation; never recurse.
            foreach ($name in @('CMDB_ConfigurationItems.csv','CMDB_Relationships.csv','CIRegistry.manifest.json.txt','CMDB_CISources.csv','CMDB_CIGovernanceChanges.csv','CMDB_CIUserContext.csv','CMDB_CIDeviceContext.csv','CMDB_CIDeviceSourceContext.csv','CMDB_CIOrganization.csv','CMDB_OrganizationReferences.csv','CMDB_CIDeviceHardware.csv')) {
                $ownFile=Join-Path $stage $name
                if (Test-Path -LiteralPath $ownFile -PathType Leaf) { Remove-Item -LiteralPath $ownFile -Force }
            }
            [IO.Directory]::Delete($stage,$false)
        }
    }
}

Export-ModuleMember -Function Export-SmartWorkplaceCMDBCIRegistry

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAOL8YJdV/CCGy5
# TecuiubsYNITj68YQV6DVQX8/FXGfqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAowcpnH+tzEIZanbtUShLm
# gK3C8mi6AfWaenhtO8Da8zANBgkqhkiG9w0BAQEFAASCAYA6GVA4Rt5X0pe8yGwp
# iAP3vYdN8lxATRonwC1uCFS8miBdiBb+xC3t8PHnSVnPN82UxWvTAlrAKW9dZwL/
# 6d4deM8e0L1XabY05KT7K+BSUcrFBh3PIlDFdud0lALpluIZ6gqoLx7gDrl+Ucq2
# G0uyJFNCxbVIjP+RdJmTNwgRQTYfqF4ARX5s66gcd9+IBHndDRr5X+vK+vipcrE/
# kzJwIu3Ug6xzBOd4Yt6gbSxI0BbseR8/4RgUbEvucLbY9tb8jsk/kG6X+mbx3gzU
# nUmG9GNb3Xqo8k/9tDWOoqmBoiPZlT4qHRnmc62AV9tYR7iMXzCigSPPEhmoGhP/
# Xtp1tiwekF+nmqBq2YDzDd/Ixb8X7VWnV0OBPm43SY5YM9+90zE4vx2kCRn7gvoz
# sKSCXa1sM6097ND+u9d+E5Ka3RoOLo28+pN0GOB1Ez2Dse44VMEdj+K8i6sInZL0
# uuFzt5JN55GX0ULfUbsrII3eJHzanrzeNbVnJIA82uu+T4g=
# SIG # End signature block
