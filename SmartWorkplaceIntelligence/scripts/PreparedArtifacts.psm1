# Version 1.2.3: receipt-bound snapshot integrity and explicit historical cohorts.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Get-PreparedArtifactContract {
    Get-Content -LiteralPath (Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt') -Raw | ConvertFrom-Json
}

function Get-PreparedLegacyArtifactNames {
    param([string]$ContractSHA256)
    if (-not $ContractSHA256) { return @() }
    $contract=Get-PreparedArtifactContract
    if (-not (Test-PreparedArtifactContractCompatible $ContractSHA256)) { throw 'Artifact contract mismatch.' }
    if ($contract.PSObject.Properties['legacyCurrentFiles']) {
        foreach ($entry in $contract.legacyCurrentFiles) {
            if ($ContractSHA256.ToUpperInvariant() -cin @($entry.contractSHA256)) { 'Source-'+$entry.file }
        }
    }
}

function Get-PreparedOptionalArtifactFiles {
    param([Parameter(Mandatory)]$Contract)
    if(-not $Contract.PSObject.Properties['optionalCurrentFiles']){return @()}
    $names=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($entry in @($Contract.optionalCurrentFiles)){
        if([string]::IsNullOrWhiteSpace($entry.file) -or [IO.Path]::GetFileName([string]$entry.file) -cne [string]$entry.file -or
           -not $names.Add($entry.file) -or $entry.file -cin @($Contract.currentFiles) -or [int]$entry.maxAgeHours -lt 1 -or [int]$entry.maxAgeHours -gt 8760){
            throw 'Invalid or duplicate optional artifact file.'
        }
        if($entry.PSObject.Properties['delimiter'] -and [string]$entry.delimiter -cnotin @(',', ';')){
            throw 'Invalid optional artifact delimiter.'
        }
    }
    @($Contract.optionalCurrentFiles)
}

function Assert-PreparedArtifactTenant {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)]$Specification)
    $delimiter=if($Specification.PSObject.Properties['delimiter']){[string]$Specification.delimiter}else{','}
    $parser=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try{
        $parser.SetDelimiters($delimiter)
        $parser.HasFieldsEnclosedInQuotes=$true
        $parser.TrimWhiteSpace=$false
        $header=$parser.ReadFields()
        if(-not $header -or @($header | Sort-Object -Unique).Count -ne $header.Length -or
           @($header | Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count){throw 'Empty or duplicate optional artifact header.'}
        $tenantIndex=[array]::IndexOf($header,'TenantKey')
        if($tenantIndex -lt 0){throw "Optional artifact TenantKey missing: $($Specification.file)"}
        while(-not $parser.EndOfData){
            $fields=$parser.ReadFields()
            if($fields.Length -ne $header.Length){throw "Malformed optional artifact record: $($Specification.file)"}
            if([string]::IsNullOrWhiteSpace($fields[$tenantIndex]) -or $fields[$tenantIndex] -cne $TenantKey){
                throw "Optional artifact tenant mismatch: $($Specification.file)"
            }
        }
    }finally{$parser.Dispose()}
}

function Get-PreparedArtifactSources {
    param([Parameter(Mandatory)][string]$DataRoot)
    $contract=Get-PreparedArtifactContract
    foreach($name in $contract.currentFiles){@{Relative="DATA-LAST/$name";Kind='Current'}}
    # Optional current files: missing or stale inputs are optional source issues, never a failed capture.
    foreach($entry in @(Get-PreparedOptionalArtifactFiles $contract)){@{Relative="DATA-LAST/$($entry.file)";Kind='Current';IsOptional=$true;MaxAgeHours=[int]$entry.maxAgeHours}}
    foreach($spec in $contract.history){
        $root=Join-Path $DataRoot $spec.root
        $weeks=@(Get-ChildItem -LiteralPath $root -Directory | Where-Object Name -Match '^\d{4}-W\d{2}$' | Sort-Object Name)
        if(-not $weeks.Count){throw "Required migration history missing: $($spec.kind)"}
        foreach($week in $weeks){
            if([int]$week.Name.Substring(6) -notin 1..53){throw 'Invalid history week.'}
            $relative="$($spec.root)/$($week.Name)"
            @{Relative="$relative/$($spec.file)";Kind='History'}
            $metadata=if(Test-Path -LiteralPath (Join-Path $week.FullName 'manifest.json.txt')){'manifest.json.txt'}else{'manifest.json'}
            @{Relative="$relative/$metadata";Kind='History'}
            if($spec.PSObject.Properties['optionalFile'] -and (Test-Path -LiteralPath (Join-Path $week.FullName $spec.optionalFile))){
                @{Relative="$relative/$($spec.optionalFile)";Kind='History'}
            }
        }
    }
}

function Test-PreparedArtifactName {
    param([Parameter(Mandatory)][string]$Name,[string]$ContractSHA256)
    $contract=Get-PreparedArtifactContract
    if ($Name -cin @(Get-PreparedLegacyArtifactNames $ContractSHA256)) { return $true }
    if($contract.PSObject.Properties['nativeSemanticArtifacts'] -and $Name -cin @($contract.nativeSemanticArtifacts | ForEach-Object {"Source-$($_.file)"})){return $true}
    if($Name -ceq 'PreparedHistoryIndex.json.txt'){return $true}
    if($Name -cin @($contract.currentFiles | ForEach-Object {"Source-$_"})){return $true}
    if($Name -cin @(Get-PreparedOptionalArtifactFiles $contract | ForEach-Object {"Source-$($_.file)"})){return $true}
    foreach($spec in $contract.history){
        $prefix='History-'+$spec.kind+'-\d{4}-W(?:0[1-9]|[1-4][0-9]|5[0-3])-'
        $names=@($spec.file,'manifest.json.txt')
        if($spec.PSObject.Properties['optionalFile']){$names+=@($spec.optionalFile)}
        foreach($leaf in $names){if($Name -cmatch ('^'+$prefix+[regex]::Escape($leaf)+'$')){return $true}}
    }
    return $false
}

function Copy-PreparedArtifacts {
    # Only the already verified capture is read. No second live capture or publication.
    param([Parameter(Mandatory)][string]$SnapshotRoot,[Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$TenantKey)
    $contract=Get-PreparedArtifactContract
    $entries=[Collections.Generic.List[object]]::new()
    $index=[Collections.Generic.List[object]]::new()
    function Copy-One([string]$Relative,[string]$Name){
        if(-not (Test-PreparedArtifactName $Name)){throw "Unapproved artifact: $Name"}
        $source=Join-Path $SnapshotRoot $Relative
        $target=Join-Path $Destination $Name
        if(Test-Path -LiteralPath $target){throw "Artifact output already exists: $Name"}
        $hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
        Copy-Item -LiteralPath $source -Destination $target -ErrorAction Stop
        if((Get-FileHash -LiteralPath $target).Hash -ne $hash -or (Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Captured artifact changed during publication.'}
        $entries.Add([pscustomobject]@{File=$Name;Bytes=(Get-Item -LiteralPath $target).Length;SHA256=$hash;SourceRelative=$Relative})
    }
    $snapshot=Get-Content -LiteralPath (Join-Path $SnapshotRoot 'DATA-LAST/M365_Licenses_ReportSnapshot.json.txt') -Raw | ConvertFrom-Json
    if($snapshot.TenantKey -cne $TenantKey -or $snapshot.SchemaVersion -notin @(1,2) -or $snapshot.SnapshotId -notmatch '^[a-fA-F0-9]{32}$'){throw 'Invalid license report snapshot identity.'}
    # Match the existing Power Query reader before committing the unified pointer.
    if($snapshot.Products -isnot [array] -or $snapshot.Products.Count -ne 4 -or
       $snapshot.OtherProducts -isnot [array] -or $snapshot.OtherProducts.Count -ne 3){throw 'Invalid license report snapshot structure.'}
    if($snapshot.SchemaVersion -eq 2){
        $issues=$snapshot.AssignmentIssues
        if($issues -isnot [pscustomobject] -or $issues.Available -isnot [bool] -or $issues.Rows -isnot [array] -or
           ($issues.Available -and $issues.Counts -isnot [pscustomobject]) -or
           (-not $issues.Available -and $issues.Rows.Count -ne 0)){throw 'Invalid license assignment snapshot structure.'}
    }
    foreach($name in $contract.currentFiles){
        if($name -like 'M365_Licenses_Report*.csv'){
            # Another tenant is fatal. A row of another snapshot is a lineage defect: the report is
            # copied and declared Not qualified by Get-PreparedLicenseReportQualification.
            foreach($row in Import-Csv -LiteralPath (Join-Path $SnapshotRoot "DATA-LAST/$name")){
                if($row.TenantKey -cne $TenantKey){throw "License report tenant mismatch: $name"}
            }
        }
        Copy-One "DATA-LAST/$name" "Source-$name"
    }
    foreach($entry in @(Get-PreparedOptionalArtifactFiles $contract)){
        $relative="DATA-LAST/$($entry.file)"
        if(-not (Test-Path -LiteralPath (Join-Path $SnapshotRoot $relative) -PathType Leaf)){continue}
        Assert-PreparedArtifactTenant -Path (Join-Path $SnapshotRoot $relative) -TenantKey $TenantKey -Specification $entry
        Copy-One $relative "Source-$($entry.file)"
    }
    if($contract.PSObject.Properties['nativeSemanticArtifacts']){
        foreach($spec in $contract.nativeSemanticArtifacts){
            $receiptPath=Join-Path $SnapshotRoot ("DATA-LAST/"+$spec.receipt)
            if(-not(Test-Path -LiteralPath $receiptPath -PathType Leaf)){continue}
            $proof=Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
            if(-not $proof.PSObject.Properties['RequiredFiles'] -or $spec.file -cnotin $proof.RequiredFiles){continue}
            if($proof.Producer -cne $spec.producer -or $proof.TenantKey -cne $TenantKey -or $proof.Status -cne 'Completed' -or $proof.Errors -ne 0 -or -not $proof.ConsumerScopeQualified){throw 'Unqualified native semantic artifact receipt.'}
            $records=@($proof.Files | Where-Object File -CEQ $spec.file)
            $relative='DATA-LAST/'+$spec.file
            $source=Join-Path $SnapshotRoot $relative
            if($records.Count -ne 1 -or $records[0].RunId -cne $proof.RunId -or (Get-FileHash -LiteralPath $source).Hash -cne $records[0].SHA256){throw 'Native semantic artifact changed or belongs to another run.'}
            Assert-PreparedArtifactTenant -Path $source -TenantKey $TenantKey -Specification $spec
            Copy-One $relative ("Source-"+$spec.file)
            $entries[$entries.Count-1] | Add-Member -NotePropertyName RunId -NotePropertyValue $proof.RunId
            $entries[$entries.Count-1] | Add-Member -NotePropertyName TenantKey -NotePropertyValue $TenantKey
            $entries[$entries.Count-1] | Add-Member -NotePropertyName SemanticVersion -NotePropertyValue $spec.contractVersion
        }
    }
    foreach($spec in $contract.history){
        foreach($week in Get-ChildItem -LiteralPath (Join-Path $SnapshotRoot $spec.root) -Directory | Where-Object Name -Match '^\d{4}-W\d{2}$' | Sort-Object Name){
            $relative="$($spec.root)/$($week.Name)"
            $prefix="History-$($spec.kind)-$($week.Name)-"
            $metadata=if(Test-Path -LiteralPath (Join-Path $week.FullName 'manifest.json.txt')){'manifest.json.txt'}else{'manifest.json'}
            $manifest=Get-Content -LiteralPath (Join-Path $week.FullName $metadata) -Raw | ConvertFrom-Json
            if($manifest.PSObject.Properties['Week'] -and $manifest.Week -cne $week.Name){throw 'History manifest week mismatch.'}
            Copy-One "$relative/$metadata" "${prefix}manifest.json.txt"
            Copy-One "$relative/$($spec.file)" "$prefix$($spec.file)"
            $issues=$null
            if($spec.PSObject.Properties['optionalFile'] -and (Test-Path -LiteralPath (Join-Path $week.FullName $spec.optionalFile))){
                $issues="$prefix$($spec.optionalFile)"; Copy-One "$relative/$($spec.optionalFile)" $issues
            }
            $index.Add(@{Kind=$spec.kind;Week=$week.Name;SourceFile=$spec.file;File="$prefix$($spec.file)";Manifest="${prefix}manifest.json.txt";Issues=$issues})
        }
    }
    $indexPath=Join-Path $Destination 'PreparedHistoryIndex.json.txt'
    if(Test-Path -LiteralPath $indexPath){throw 'History index already exists.'}
    @{SchemaVersion=1;TenantKey=$TenantKey;Entries=$index.ToArray()} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $indexPath -Encoding utf8NoBOM
    $entries.Add([pscustomobject]@{File='PreparedHistoryIndex.json.txt';Bytes=(Get-Item -LiteralPath $indexPath).Length;SHA256=(Get-FileHash -LiteralPath $indexPath).Hash;SourceRelative='GeneratedFromVerifiedCapture'})
    $entries.ToArray()
}

function Assert-PreparedArtifacts {
    param([Parameter(Mandatory)][string]$BatchRoot,[Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Entries,
        [Parameter(Mandatory)][string]$TenantKey,[string]$ContractSHA256)
    $legacy=@(Get-PreparedLegacyArtifactNames $ContractSHA256)
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $optional=@(Get-PreparedOptionalArtifactFiles (Get-PreparedArtifactContract))
    foreach($entry in $Entries){
        if(-not (Test-PreparedArtifactName ([string]$entry.File) -ContractSHA256 $ContractSHA256) -or -not $seen.Add($entry.File)){throw 'Unsafe or duplicate prepared artifact.'}
        $path=Join-Path $BatchRoot $entry.File
        if((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Artifact links are forbidden.'}
        if($entry.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or (Get-Item -LiteralPath $path).Length -ne $entry.Bytes -or (Get-FileHash -LiteralPath $path).Hash -ne $entry.SHA256){throw "Artifact integrity mismatch: $($entry.File)"}
        $specification=@($optional | Where-Object { "Source-$($_.file)" -ceq $entry.File })
        if($specification.Count -eq 1){Assert-PreparedArtifactTenant -Path $path -TenantKey $TenantKey -Specification $specification[0]}
    }
    $required=@('PreparedHistoryIndex.json.txt')+$legacy+@((Get-PreparedArtifactContract).currentFiles | ForEach-Object {"Source-$_"})
    foreach($name in $required){if(-not $seen.Contains($name)){throw "Required artifact missing: $name"}}
    $index=Get-Content -LiteralPath (Join-Path $BatchRoot 'PreparedHistoryIndex.json.txt') -Raw | ConvertFrom-Json
    if($index.SchemaVersion -ne 1 -or $index.TenantKey -cne $TenantKey){throw 'Invalid prepared history index.'}
    $referenced=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($name in $required){$null=$referenced.Add($name)}
    # Optional current files are declared by the contract, not by the history index.
    foreach($entry in @(Get-PreparedOptionalArtifactFiles (Get-PreparedArtifactContract))){if($seen.Contains("Source-$($entry.file)")){$null=$referenced.Add("Source-$($entry.file)")}}
    $artifactContract=Get-PreparedArtifactContract
    if($artifactContract.PSObject.Properties['nativeSemanticArtifacts']){foreach($spec in $artifactContract.nativeSemanticArtifacts){
        $name='Source-'+$spec.file
        if($seen.Contains($name)){
            $entry=@($Entries | Where-Object File -CEQ $name)[0]
            if(-not $entry.PSObject.Properties['RunId'] -or -not $entry.RunId -or $entry.TenantKey -cne $TenantKey -or $entry.SemanticVersion -cne $spec.contractVersion){throw 'Native semantic artifact lineage missing.'}
            Assert-PreparedArtifactTenant -Path (Join-Path $BatchRoot $name) -TenantKey $TenantKey -Specification $spec
            $null=$referenced.Add($name)
        }
    }}
    $weeks=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($entry in $index.Entries){
        if($entry.Kind -cnotin @('Mailbox','Windows') -or $entry.Week -notmatch '^\d{4}-W(?:0[1-9]|[1-4][0-9]|5[0-3])$' -or -not $weeks.Add("$($entry.Kind)/$($entry.Week)")){throw 'Invalid or duplicate history index entry.'}
        $spec=@((Get-PreparedArtifactContract).history | Where-Object kind -CEQ $entry.Kind)[0]
        $prefix="History-$($entry.Kind)-$($entry.Week)-"
        if($entry.SourceFile -cne $spec.file -or $entry.File -cne "$prefix$($spec.file)" -or $entry.Manifest -cne "${prefix}manifest.json.txt"){throw 'History index scope mismatch.'}
        if($entry.Issues -and (-not $spec.PSObject.Properties['optionalFile'] -or $entry.Issues -cne "$prefix$($spec.optionalFile)")){throw 'Invalid history issue artifact.'}
        foreach($name in @($entry.File,$entry.Manifest,$entry.Issues) | Where-Object {$_}){
            if(-not $seen.Contains($name)){throw 'History index references an undeclared artifact.'}
            $null=$referenced.Add($name)
        }
    }
    foreach($kind in 'Mailbox','Windows'){if(-not @($index.Entries | Where-Object Kind -CEQ $kind).Count){throw "Required indexed history missing: $kind"}}
    if($referenced.Count -ne $seen.Count){throw 'Unindexed historical artifact.'}
}

function Assert-PreparedLicenseSnapshotSource {
    # Only frozen capture paths are read. Never repair a snapshot from current live DATA-LAST.
    param([Parameter(Mandatory)][string]$Last,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)]$Snapshot)
    function Read-Utc($Value) {
        # ConvertFrom-Json yields Utc for 'Z', Local for an explicit offset such as '+00:00'
        # and Unspecified when no timezone is written; only the last one is ambiguous.
        if ($Value -is [datetimeoffset]) { return $Value.ToUniversalTime() }
        if ($Value -is [datetime]) {
            if ($Value.Kind -eq [DateTimeKind]::Unspecified) { throw 'License source timestamp has no explicit timezone.' }
            return [datetimeoffset]::new($Value.ToUniversalTime())
        }
        if ($Value -notmatch '(Z|[+-]\d{2}:\d{2})$') { throw 'License source timestamp has no explicit timezone.' }
        [datetimeoffset]::Parse($Value,[Globalization.CultureInfo]::InvariantCulture).ToUniversalTime()
    }
    if (-not $Snapshot.PSObject.Properties['Sources'] -or $Snapshot.Sources -isnot [array]) { throw 'License snapshot contains no source qualification evidence.' }
    $generated=Read-Utc $Snapshot.GeneratedAtUtc
    $groups=@(
        @{Producer='SmartM365-M365UserActivity-Inventory.ps1';Files=@('M365_Users_Activity.csv','M365_Mailbox_Usage.csv','M365_Email_Activity.csv','M365_Apps_Usage_180D.csv','M365_OneDrive_Usage.csv')},
        @{Producer='SmartM365-Devices-Inventory.ps1';Files=@('Intune_Devices_Inventory.csv')}
    )
    foreach ($group in $groups) {
        $stem='SmartInventory_'+[IO.Path]::GetFileNameWithoutExtension($group.Producer)
        $proofPath=Join-Path $Last ($stem+'.current.json.txt')
        if (-not (Test-Path -LiteralPath $proofPath -PathType Leaf)) { throw "License dependency receipt not captured: $($group.Producer)" }
        $proof=Get-Content -LiteralPath $proofPath -Raw | ConvertFrom-Json
        if ($proof.Owner -cne 'SmartInventory-SourceReceipt' -or $proof.TenantKey -cne $TenantKey -or $proof.Producer -cne $group.Producer -or
            $proof.Status -cne 'Completed' -or $proof.Errors -ne 0 -or $proof.IsPartialInventory -ne $false -or
            $proof.ConsumerScopeQualified -ne $true -or -not $proof.RunId) { throw "License dependency receipt is not qualified: $($group.Producer)" }
        $completed=Read-Utc $proof.CompletedAtUtc
        if ($generated -lt $completed) { throw "License snapshot predates the completed dependency: $($group.Producer)" }
        $runPath=Join-Path $Last ($stem+'.run.json.txt')
        if ($proof.PSObject.Properties['PublicationProtocol'] -or (Test-Path -LiteralPath $runPath)) {
            if (-not (Test-Path -LiteralPath $runPath -PathType Leaf)) { throw "License dependency run state not captured: $($group.Producer)" }
            $run=Get-Content -LiteralPath $runPath -Raw | ConvertFrom-Json
            if ($run.Owner -cne 'SmartInventory-SourceRun' -or $run.Status -cne 'Completed' -or $run.UnqualifiedPublication -ne $false -or $run.RunId -cne $proof.RunId -or
                $run.TenantKey -cne $TenantKey -or $run.Producer -cne $group.Producer) { throw "License dependency publication is not complete: $($group.Producer)" }
        }
        foreach ($name in $group.Files) {
            $source=@($Snapshot.Sources | Where-Object { $_.Name -ceq $name })
            if ($source.Count -ne 1 -or $source[0].Ready -isnot [bool] -or -not $source[0].Ready -or ($source[0].PSObject.Properties['Provisional'] -and $source[0].Provisional -ne $false)) { throw "License snapshot did not qualify source: $name" }
            $record=@($proof.Files | Where-Object { $_.File -ceq $name })
            $path=Join-Path $Last $name
            if ($record.Count -ne 1 -or $record[0].Status -cne 'Success' -or $record[0].IsPartialInventory -ne $false -or
                $record[0].RunId -cne $proof.RunId -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "License source completion proof is invalid: $name" }
            $hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
            if ([string]$source[0].SHA256 -cnotmatch '^[A-Fa-f0-9]{64}$' -or
                $hash -ine [string]$source[0].SHA256 -or $hash -ine [string]$record[0].SHA256) { throw "License snapshot uses a different source capture: $name" }
        }
    }
}

function Get-PreparedLicenseReportQualification {
    # The license report is consumed only when the current licensing receipt records it:
    # completed receipt, same tenant and producer, every report CSV hashed in Files and declared
    # Published, snapshot ReceiptRun.RunId equal to the receipt RunId. A lineage defect yields
    # "Not qualified" (empty tables in the report), never a failed batch or older values.
    # New Prepare/Workforce consumption additionally requires frozen Usage/Devices sources to
    # match the report's qualified SHA256s and completed producer receipts. Historical callers
    # may omit RequireCurrentSources; existing batches are never rewritten or repaired.
    param([Parameter(Mandatory)][string]$SnapshotRoot,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Artifacts,[switch]$RequireCurrentSources)
    $spec=(Get-PreparedArtifactContract).licenseReport
    $last=Join-Path $SnapshotRoot 'DATA-LAST'
    $hashes=@{}
    foreach($entry in $Artifacts){$hashes[[string]$entry.File]=[string]$entry.SHA256}
    $snapshot=Get-Content -LiteralPath (Join-Path $last $spec.snapshot) -Raw | ConvertFrom-Json
    $result=[ordered]@{Version=1;State='Qualified';Reason='';SnapshotId=[string]$snapshot.SnapshotId;ReceiptRunId='';ReceiptSHA256='';Files=@()}
    $receipt=$null;$receiptReason=''
    $receiptPath=Join-Path $last $spec.receipt
    if(-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)){$receiptReason='Licensing receipt is not part of the capture.'}
    else{
        try{
            $result.ReceiptSHA256=(Get-FileHash -LiteralPath $receiptPath -Algorithm SHA256).Hash
            $receipt=Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
            $result.ReceiptRunId=[string]$receipt.RunId
            if([string]$receipt.Producer -cne $spec.producer -or [string]$receipt.TenantKey -cne $TenantKey){$receiptReason='Licensing receipt belongs to another producer or tenant.'}
            elseif([string]$receipt.Status -cne 'Completed'){$receiptReason=('Licensing receipt is {0}, not Completed.' -f $receipt.Status)}
            elseif(-not $snapshot.PSObject.Properties['ReceiptRun'] -or $null -eq $snapshot.ReceiptRun){$receiptReason='Report snapshot names no licensing receipt run (produced outside a receipt run).'}
            elseif([string]$snapshot.ReceiptRun.Producer -cne $spec.producer -or [string]$snapshot.ReceiptRun.RunId -cne [string]$receipt.RunId){$receiptReason='Report snapshot belongs to another licensing receipt run.'}
        }catch{$receiptReason=('Licensing receipt is unreadable: {0}' -f $_.Exception.Message)}
    }
    if ($RequireCurrentSources -and -not $receiptReason) {
        try { Assert-PreparedLicenseSnapshotSource -Last $last -TenantKey $TenantKey -Snapshot $snapshot }
        catch { $receiptReason=('License snapshot dependency mismatch: '+$_.Exception.Message) }
    }
    $files=[Collections.Generic.List[object]]::new()
    $integrityName='M365_Licenses_ReportSnapshotIntegrity.csv'
    foreach($name in @($spec.reportFiles)+@($spec.decisionFiles)+@($integrityName)){
        $decision=$name -cin @($spec.decisionFiles)
        $integrity=$name -ceq $integrityName
        $artifact="Source-$name"
        $state='Qualified';$reason=''
        if(-not $hashes.ContainsKey($artifact)){$state='Absent';$reason='Not part of the licensing output used by this batch.'}
        elseif($receiptReason){$state='Not qualified';$reason=$receiptReason}
        else{
            $record=@($receipt.Files | Where-Object {[string]$_.File -ceq $name})
            $declared=@($receipt.OptionalFiles | Where-Object {[string]$_.File -ceq $name -and $_.Published -eq $true})
            if($record.Count -ne 1 -or $declared.Count -ne 1){$state='Not qualified';$reason='File is not recorded as published by the current licensing receipt.'}
            elseif($integrity -and (-not $receipt.PSObject.Properties['Errors'] -or $receipt.Errors -ne 0 -or
                -not $receipt.PSObject.Properties['IsPartialInventory'] -or $receipt.IsPartialInventory -ne $false -or
                -not $receipt.PSObject.Properties['ConsumerScopeQualified'] -or $receipt.ConsumerScopeQualified -ne $true -or
                [string]$record[0].Status -cne 'Success' -or [string]$record[0].RunId -cne [string]$receipt.RunId)){$state='Not qualified';$reason='Licensing receipt or integrity file is not complete and qualified.'}
            elseif([string]$record[0].SHA256 -ine $hashes[$artifact]){$state='Not qualified';$reason='File hash differs from the current licensing receipt.'}
            else{
                $rows=@(Import-Csv -LiteralPath (Join-Path $last $name))
                if($integrity -and ($rows.Count -ne 1 -or @($rows[0].PSObject.Properties.Name).Count -ne 4 -or
                    @('SnapshotId','TenantKey','ReceiptRunId','SnapshotSHA256' | Where-Object { $_ -cnotin $rows[0].PSObject.Properties.Name }).Count -ne 0)) {
                    $state='Not qualified';$reason='Snapshot integrity file has an invalid schema or row count.';$rows=@()
                }
                foreach($row in $rows){
                    if([string]$row.SnapshotId -cne [string]$snapshot.SnapshotId){$state='Not qualified';$reason='Rows belong to another report snapshot.';break}
                    if(($decision -or $integrity) -and [string]$row.ReceiptRunId -cne [string]$receipt.RunId){$state='Not qualified';$reason='Rows belong to another licensing receipt run.';break}
                    if($integrity -and ([string]$row.TenantKey -cne $TenantKey -or [string]$row.SnapshotSHA256 -cnotmatch '^[A-Fa-f0-9]{64}$' -or [string]$row.SnapshotSHA256 -ine $hashes["Source-$($spec.snapshot)"])){$state='Not qualified';$reason='Snapshot integrity evidence differs from the captured JSON.';break}
                }
            }
        }
        $files.Add([ordered]@{File=$name;Kind=$(if($integrity){'Integrity'}elseif($decision){'Decision'}else{'Report'});State=$state;Reason=$reason;SHA256=$(if($hashes.ContainsKey($artifact)){$hashes[$artifact]}else{''})})
    }
    $core=@($files | Where-Object {$_.Kind -eq 'Report' -and $_.State -ne 'Qualified'})
    if($core.Count -gt 0){
        $result.State='Not qualified'
        $result.Reason=$(if($receiptReason){$receiptReason}else{'{0}: {1}' -f $core[0].File,$core[0].Reason})
        # Decision files are never shown without their qualified report.
        foreach($file in $files){if($file.Kind -in @('Decision','Integrity') -and $file.State -eq 'Qualified'){$file.State='Not qualified';$file.Reason='License report is not qualified.'}}
    }
    $result.Files=$files.ToArray()
    [pscustomobject]$result
}

function Test-PreparedArtifactContractCompatible {
    # A batch built with a previous artifact contract stays transferable when the current
    # contract lists the exact previous hash; retired files use a separate read-only cohort.
    param([Parameter(Mandatory)][string]$ContractSHA256)
    $path=Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt'
    if($ContractSHA256 -ieq (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash){return $true}
    $contract=Get-PreparedArtifactContract
    return $contract.PSObject.Properties['compatibleContractSHA256'] -and $ContractSHA256 -cin @($contract.compatibleContractSHA256 | ForEach-Object {([string]$_).ToUpperInvariant()})
}

Export-ModuleMember -Function Get-PreparedArtifactSources,Copy-PreparedArtifacts,Assert-PreparedArtifacts,Test-PreparedArtifactName,Get-PreparedLicenseReportQualification,Test-PreparedArtifactContractCompatible

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD3yBuvMB/EBDGS
# dTAm0T2qgBCakLZqXXCDjClI85qaXKCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgjqCCa3RxbFn4
# yfo/nhs2NdZjqnEyv0NFJY3GBoOfe3YwDQYJKoZIhvcNAQEBBQAEggGAgyZcb3LL
# 45CsfdND6AUPwoLXhJ2P9FeI0QULYblX5NzBaY+afEL0LSgCQ2zhUIkcf7XWknTG
# Xw50I3m2m55e3DgmQ1rIZ9lgUnw1oj+1VePIvMPibwjge0kLcRHJu9SV2g33c0+R
# vIibItKLUKhjcg00Z0C/o9hSAuaJHzUA9/tbmoVCfQCkMQKJ8fUM6/wULop3q1SK
# dAeq2ewaRzIxs+rDpJ8K9lOZUNzMVevnSS8Up5Rah/W8xJGtJFFG8dz714mZtEUm
# s9LB29O21sH7h7O+euJ9CFkZxWdEXTHSvXNXlsdb1KChuNuhMPpK/p9TotKYwuK7
# IljPyXxdOE5YnF4DmUZJS4g8OgFgQYVcnqsfgdbeSvNiLX6tEbZ7kshEQD/Mi0C6
# ySkI7+NMtHSQfgGVLmfEnjdBa6Tj6ItqJVM+db6cu3rCBYnC077QJnKcUE3yEUII
# TPxJZKA3/34dHrRqp1PXq3s0I7PGxo1+zmnbz1SxFbROLiVrPDzvz5MloYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTA1MTU0NzI3WjA/BgkqhkiG9w0BCQQxMgQwkzKrfwcNqbgrvZZZtCKM8Zs5
# Bx77PKLFgZwq5QoD292opOmREl5qtwGApg/DTZzgMA0GCSqGSIb3DQEBAQUABIIC
# AKWaD5iWlzf3zyzMPyhtSDleehHSD1r9pHgY/24OEUHEkHKwpPQGgAvAtHRrVrv/
# Qh24hMQxSE+EY/9hdE3cJvtc5r42zj4K8NM7IsBpgbBATp25CpUL9JYqOXYfsLMS
# ZGFRbbX5TgU5q6V431wR+yjb5/MhsLeNAorFdsYtlHq7gisCztjb3nYZVF1MBIkw
# XZfE17tK/CWDGGZ/DeY/Yu96Ut5mxhR5W/1DC1gmKPGUkw6+IJbV5wfBLCFgaIRl
# i+3m4rZlYJGYE8fkVNmOr/AUbssCfpA+lho+Ly9e5rsP9fH5iGe83GqzwLFd65IK
# qDhKqAQOjSSAPis+2FYOR/iNDxXW+FCFz2rzeGf5D/5Zy0MzUMLW5Hbzf+BNDgui
# 6i2aTeh9P5/2rijMz0tO9gsm0XAjoGC69pg+BU6brvbDC5NJseL/nViaRCJyQbl4
# Al3eNaXPa0+R5f+nOM/q5CTvFwsAgnqKg9Oq8D35SwS04/Iv6xgOTOLwSKjQcjAb
# q/5ghe7nigplW7bR+D4IwsDxLUl8IvLqIpivkpQNzLce7IiK8MiggamxM2Lf/As3
# ottKtf39zQGgt+IHbJjKAkUsPTlmdxT6qbBhWW+PAx1MmO+eeHTsl7tW/SlVaqGP
# H5B02sTLOjIZBEAp7dr46fN9whSrPF4M0yZvmitLMkXd
# SIG # End signature block
