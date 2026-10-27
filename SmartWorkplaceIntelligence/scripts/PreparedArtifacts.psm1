# Version 1.1.0: explicit historical artifact cohorts.
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

function Get-PreparedLicenseReportQualification {
    # The license report is consumed only when the current licensing receipt records it:
    # completed receipt, same tenant and producer, every report CSV hashed in Files and declared
    # Published, snapshot ReceiptRun.RunId equal to the receipt RunId. A lineage defect yields
    # "Not qualified" (empty tables in the report), never a failed batch or older values.
    param([Parameter(Mandatory)][string]$SnapshotRoot,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Artifacts)
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
    $files=[Collections.Generic.List[object]]::new()
    foreach($name in @($spec.reportFiles)+@($spec.decisionFiles)){
        $decision=$name -cin @($spec.decisionFiles)
        $artifact="Source-$name"
        $state='Qualified';$reason=''
        if(-not $hashes.ContainsKey($artifact)){$state='Absent';$reason='Not part of the licensing output used by this batch.'}
        elseif($receiptReason){$state='Not qualified';$reason=$receiptReason}
        else{
            $record=@($receipt.Files | Where-Object {[string]$_.File -ceq $name})
            $declared=@($receipt.OptionalFiles | Where-Object {[string]$_.File -ceq $name -and $_.Published -eq $true})
            if($record.Count -ne 1 -or $declared.Count -ne 1){$state='Not qualified';$reason='File is not recorded as published by the current licensing receipt.'}
            elseif([string]$record[0].SHA256 -ine $hashes[$artifact]){$state='Not qualified';$reason='File hash differs from the current licensing receipt.'}
            else{
                foreach($row in Import-Csv -LiteralPath (Join-Path $last $name)){
                    if([string]$row.SnapshotId -cne [string]$snapshot.SnapshotId){$state='Not qualified';$reason='Rows belong to another report snapshot.';break}
                    if($decision -and [string]$row.ReceiptRunId -cne [string]$receipt.RunId){$state='Not qualified';$reason='Rows belong to another licensing receipt run.';break}
                }
            }
        }
        $files.Add([ordered]@{File=$name;Kind=$(if($decision){'Decision'}else{'Report'});State=$state;Reason=$reason;SHA256=$(if($hashes.ContainsKey($artifact)){$hashes[$artifact]}else{''})})
    }
    $core=@($files | Where-Object {$_.Kind -eq 'Report' -and $_.State -ne 'Qualified'})
    if($core.Count -gt 0){
        $result.State='Not qualified'
        $result.Reason=$(if($receiptReason){$receiptReason}else{'{0}: {1}' -f $core[0].File,$core[0].Reason})
        # Decision files are never shown without their qualified report.
        foreach($file in $files){if($file.Kind -eq 'Decision' -and $file.State -eq 'Qualified'){$file.State='Not qualified';$file.Reason='License report is not qualified.'}}
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
