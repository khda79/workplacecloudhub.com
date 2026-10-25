# Version: 1.1.2
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedMetadata.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'PreparedArtifacts.psm1')

function Resolve-PreparedSharePointFolder {
    param([string]$ConfiguredPath,[string]$NormalizedDataRoot)
    $path=if([string]::IsNullOrWhiteSpace($ConfiguredPath)){
        if([string]::IsNullOrWhiteSpace($NormalizedDataRoot)){throw 'SharePoint data root is missing.'}
        $NormalizedDataRoot.TrimEnd('/','\')+'/DATA-POWERBI'
    }else{$ConfiguredPath}
    $path=($path -replace '\\','/').Trim('/')
    if($path -match '[:%?#]' -or @($path.Split('/') | Where-Object {$_ -in @('','.','..','DATA-LAST','DATA-ALL','LOG-ALL')}).Count -or $path.Split('/')[-1] -cne 'DATA-POWERBI'){
        throw 'Prepared SharePoint folder must be a library-relative DATA-POWERBI path, outside raw/log folders.'
    }
    $path
}

function Assert-TransferPath([string]$Path){
    $cursor=[IO.Path]::GetFullPath($Path);$volume=[IO.Path]::GetPathRoot($cursor).TrimEnd('\','/')
    while($cursor){
        if((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Transfer refuses linked paths.'}
        if($cursor.TrimEnd('\','/') -eq $volume){break};$cursor=Split-Path $cursor -Parent
    }
}

function Resolve-PreparedTransferContract {
    param([Parameter(Mandatory)]$Manifest,[Parameter(Mandatory)][string]$ContractPath)
    $contractReceipt=Read-SmartM365JsonDocument $ContractPath
    $contract=$contractReceipt.Document
    # Old immutable batches remain transferable after the optional Intune adapter
    # is deployed. Do not force a preparation solely to add new optional tables.
    $intuneContract=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/intune-operational-contract.json.txt')).Document
    $intuneFiles=@($intuneContract.tables.file)
    $intunePresent=@($Manifest.Files | Where-Object File -in $intuneFiles)
    $contractFiltered=$false
    # Batches built before New-LicensingEvidence was retired carry its two tables: restore their
    # reviewed definitions at their original positions so the recorded contract hash still matches.
    $legacyLicense=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/legacy-license-evidence-tables.json.txt')).Document
    $legacyLicenseFiles=@($legacyLicense.Tables | ForEach-Object {[string]$_.table.file})
    $legacyLicensePresent=@($Manifest.Files | Where-Object File -in $legacyLicenseFiles)
    if($legacyLicensePresent.Count){
        if($legacyLicensePresent.Count -ne $legacyLicenseFiles.Count -or @($legacyLicensePresent.File | Sort-Object -Unique).Count -ne $legacyLicenseFiles.Count){throw 'Incomplete legacy license evidence file set; transfer not started.'}
        if(@($contract.tables | Where-Object file -in $legacyLicenseFiles).Count -eq 0){
            $contract=$contract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
            $tables=[Collections.Generic.List[object]]::new()
            foreach($table in @($contract.tables)){$tables.Add($table)}
            foreach($entry in @($legacyLicense.Tables | Sort-Object {[int]$_.index})){$tables.Insert([int]$entry.index,$entry.table)}
            $contract.tables=$tables.ToArray()
            $contractFiltered=$true
        }
    }
    # Retain transfer of immutable pre-DailyTrends cohorts. A declared new
    # history cohort or a partial file pair must never fall back to legacy.
    $dailyFiles=@('DailyDirectoryTrends.csv','DailyMailboxTrends.csv')
    $dailyTables=@('Daily Directory Trends','Daily Mailbox Trends')
    $dailyPresent=@($Manifest.Files | Where-Object File -in $dailyFiles)
    $dailySupplied=@($contract.tables | Where-Object file -in $dailyFiles)
    if($dailyPresent.Count -eq 0){
        $dailyDeclared=$Manifest.PSObject.Properties['HistoryKeyVersions'] -and
            @($dailyTables | Where-Object {$Manifest.HistoryKeyVersions.PSObject.Properties[$_]}).Count -gt 0
        if($dailyDeclared){throw 'Incomplete daily trends file set; transfer not started.'}
        if($dailySupplied.Count){
            if($dailySupplied.Count -ne 2){throw 'Incomplete supplied daily trends contract; transfer not started.'}
            $contract=$contract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
            $contract.tables=@($contract.tables | Where-Object file -notin $dailyFiles)
            $contractFiltered=$true
        }
    }elseif($dailyPresent.Count -ne 2 -or @($dailyPresent.File | Sort-Object -Unique).Count -ne 2){
        throw 'Incomplete daily trends file set; transfer not started.'
    }elseif($dailySupplied.Count -ne 2){
        throw 'Incomplete supplied daily trends contract; transfer not started.'
    }
    # Batches prepared before the hybrid identity findings and device readiness gaps carry neither
    # file and stay transferable; a batch with only one of them is incomplete.
    $readinessFiles=@('HybridIdentityFindings.csv','DeviceReadinessGapEvidence.csv')
    $readinessPresent=@($Manifest.Files | Where-Object File -in $readinessFiles)
    if($readinessPresent.Count -eq 0){
        if(@($contract.tables | Where-Object file -in $readinessFiles).Count){
            $contract=$contract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
            $contract.tables=@($contract.tables | Where-Object file -notin $readinessFiles)
            $contractFiltered=$true
        }
    }elseif($readinessPresent.Count -ne 2 -or @($readinessPresent.File | Sort-Object -Unique).Count -ne 2){
        throw 'Incomplete readiness findings file set; transfer not started.'
    }
    if($intunePresent.Count -eq 0){
        if($Manifest.PSObject.Properties['Provenance'] -and $Manifest.Provenance.PSObject.Properties['IntuneOperations']){throw 'Incomplete Intune operational file set; transfer not started.'}
        if(@($contract.tables | Where-Object file -in $intuneFiles).Count){
            $contract=$contract | ConvertTo-Json -Depth 30 | ConvertFrom-Json
            $contract.tables=@($contract.tables | Where-Object file -notin $intuneFiles)
            $contractFiltered=$true
        }
    }elseif($intunePresent.Count -ne $intuneFiles.Count -or @($intunePresent.File | Sort-Object -Unique).Count -ne $intuneFiles.Count){
        throw 'Incomplete Intune operational file set; transfer not started.'
    }else{
        $provenance=$null
        if($Manifest.PSObject.Properties['Provenance'] -and $Manifest.Provenance.PSObject.Properties['IntuneOperations']){$provenance=$Manifest.Provenance.IntuneOperations}
        if($null -eq $provenance -or -not $provenance.PSObject.Properties['ContractSHA256'] -or
            $provenance.ContractSHA256 -ine (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot '../config/intune-operational-contract.json.txt')).Hash){
            throw 'Intune operational batch lacks matching reviewed contract provenance; transfer not started.'
        }
    }
    $nativePath=Join-Path $PSScriptRoot '../config/native-output-contract.json.txt'
    $nativeContract=(Read-SmartM365JsonDocument $nativePath).Document
    $nativeFiles=@($nativeContract.tables.file)
    $nativePresent=@($Manifest.Files | Where-Object File -in $nativeFiles)
    $nativeProvenance=$null
    if($Manifest.PSObject.Properties['Provenance'] -and $Manifest.Provenance.PSObject.Properties['NativeEvidence']){$nativeProvenance=$Manifest.Provenance.NativeEvidence}
    if($nativePresent.Count -or $null -ne $nativeProvenance){
        if($nativePresent.Count -ne $nativeFiles.Count -or @($nativePresent.File | Sort-Object -Unique).Count -ne $nativeFiles.Count){throw 'Incomplete unified native evidence file set; transfer not started.'}
        if($null -eq $nativeProvenance -or -not $nativeProvenance.PSObject.Properties['ContractSHA256'] -or $nativeProvenance.ContractSHA256 -ine (Get-FileHash -LiteralPath $nativePath).Hash){throw 'Native evidence contract provenance mismatch; transfer not started.'}
        $nativeSupplied=@($contract.tables | Where-Object file -in $nativeFiles)
        if($nativeSupplied.Count -eq 0){
            Import-Module (Join-Path $PSScriptRoot 'NativeEvidence.psm1') -Force
            $contract=Merge-NativeEvidenceContract -BaseContract $contract
            $contractFiltered=$true
        }elseif($nativeSupplied.Count -ne $nativeFiles.Count){throw 'Incomplete supplied native contract; transfer not started.'}
    }
    # Match preparation order: native, SharePoint migration, mailbox exceptions, endpoint.
    # Resolve from immutable files/provenance, including disabled typed-empty cohorts.
    $migrationPath=Join-Path $PSScriptRoot '../config/sharepoint-migration-contract.json.txt'
    $migrationContract=(Read-SmartM365JsonDocument $migrationPath).Document
    $migrationFiles=@($migrationContract.tables.file)
    $migrationPresent=@($Manifest.Files | Where-Object File -in $migrationFiles)
    $migrationProvenance=$null
    if($Manifest.PSObject.Properties['Provenance'] -and $Manifest.Provenance.PSObject.Properties['SharePointMigration']){$migrationProvenance=$Manifest.Provenance.SharePointMigration}
    if($migrationPresent.Count -or $null -ne $migrationProvenance){
        if($migrationPresent.Count -ne $migrationFiles.Count -or @($migrationPresent.File | Sort-Object -Unique).Count -ne $migrationFiles.Count){throw 'Incomplete SharePoint migration file set; transfer not started.'}
        if($null -eq $migrationProvenance -or -not $migrationProvenance.PSObject.Properties['ContractSHA256'] -or $migrationProvenance.ContractSHA256 -ine (Get-FileHash -LiteralPath $migrationPath).Hash){throw 'SharePoint migration contract provenance mismatch; transfer not started.'}
        $migrationSupplied=@($contract.tables | Where-Object file -in $migrationFiles)
        if($migrationSupplied.Count -eq 0){
            $contract.tables=@($contract.tables)+@($migrationContract.tables)
            $contractFiltered=$true
        }elseif($migrationSupplied.Count -ne $migrationFiles.Count -or @($migrationSupplied.file | Sort-Object -Unique).Count -ne $migrationFiles.Count){throw 'Incomplete supplied SharePoint migration contract; transfer not started.'}
    }
    $exceptionPath=Join-Path $PSScriptRoot '../config/mailbox-migration-exceptions-contract.json.txt'
    $exceptionContract=(Read-SmartM365JsonDocument $exceptionPath).Document
    $exceptionPresent=@($Manifest.Files | Where-Object File -in @($exceptionContract.tables.file))
    $exceptionProof=$null
    if($Manifest.PSObject.Properties['Provenance'] -and $Manifest.Provenance.PSObject.Properties['MailboxMigrationExceptions']){$exceptionProof=$Manifest.Provenance.MailboxMigrationExceptions}
    if($exceptionPresent.Count -or $exceptionProof){
        if($exceptionPresent.Count -ne 2 -or @($exceptionPresent.File | Sort-Object -Unique).Count -ne 2 -or -not $exceptionProof -or $exceptionProof.ContractSHA256 -ine (Get-FileHash -LiteralPath $exceptionPath).Hash){throw 'Mailbox migration exception file set or provenance is incomplete; transfer not started.'}
        $supplied=@($contract.tables | Where-Object file -in @($exceptionContract.tables.file))
        if($supplied.Count -eq 0){$contract.tables=@($contract.tables)+@($exceptionContract.tables);$contractFiltered=$true}
        elseif($supplied.Count -ne 2){throw 'Incomplete supplied mailbox exception contract.'}
    }
    $endpointFiles=@('EndpointProtectionPolicy.csv','EndpointProtectionEvidence.csv')
    $present=@($Manifest.Files | Where-Object { $_.File -in $endpointFiles })
    $endpointProvenance=$null
    if($Manifest.PSObject.Properties['Provenance'] -and $null -ne $Manifest.Provenance -and
        $Manifest.Provenance.PSObject.Properties['EndpointProtectionPolicy']){
        $endpointProvenance=$Manifest.Provenance.EndpointProtectionPolicy
    }
    # Use the immutable batch, never the tenant's current Enabled setting.
    if($present.Count -eq 0 -and $null -eq $endpointProvenance){
        if($nativePresent.Count){
            $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($contract | ConvertTo-Json -Depth 20))
            $hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
            if($hash -ine $Manifest.ContractSHA256){throw 'Unified native contract hash mismatch; transfer not started.'}
        }
        return $contract
    }
    if($present.Count -ne 2 -or @($present.File | Sort-Object -Unique).Count -ne 2){
        throw 'Incomplete endpoint protection file set; transfer not started.'
    }
    if($null -eq $endpointProvenance -or -not $endpointProvenance.PSObject.Properties['PolicyVersion'] -or
        [string]::IsNullOrWhiteSpace([string]$endpointProvenance.PolicyVersion) -or
        -not $endpointProvenance.PSObject.Properties['SHA256'] -or $endpointProvenance.SHA256 -notmatch '^[A-Fa-f0-9]{64}$'){
        throw 'Endpoint protection batch lacks qualified policy provenance; transfer not started.'
    }
    if(-not $Manifest.PSObject.Properties['ContractSHA256'] -or $Manifest.ContractSHA256 -notmatch '^[A-Fa-f0-9]{64}$'){
        throw 'Endpoint protection batch lacks a contract hash; transfer not started.'
    }
    $alreadyExtended=@($contract.tables | Where-Object { $_.file -in $endpointFiles })
    if($alreadyExtended.Count -eq 0){
        $extension=(Read-SmartM365JsonDocument (Join-Path $PSScriptRoot '../config/endpoint-protection-contract.json.txt')).Document
        if(@($extension.tables).Count -ne 2 -or @($extension.tables.file | Where-Object { $_ -in $endpointFiles } | Sort-Object -Unique).Count -ne 2){
            throw 'Invalid reviewed endpoint protection contract extension; transfer not started.'
        }
        Import-Module (Join-Path $PSScriptRoot 'EndpointProtectionEvidence.psm1') -Force
        $contract=Merge-EndpointProtectionContract -BaseContract $contract -ExtensionContract $extension
        # Match the exact UTF-8, no-BOM serialization used by Write-PreparedJson.
        $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($contract | ConvertTo-Json -Depth 20))
        $sha=[Security.Cryptography.SHA256]::Create()
        try{$hash=[BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-','')}finally{$sha.Dispose()}
    }elseif($alreadyExtended.Count -eq 2 -and @($alreadyExtended.file | Sort-Object -Unique).Count -eq 2){
        if($contractFiltered){
            $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($contract | ConvertTo-Json -Depth 20))
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-','')}finally{$sha.Dispose()}
        }else{$hash=(Get-FileHash -LiteralPath $contractReceipt.Path -Algorithm SHA256).Hash}
    }else{
        throw 'Incomplete supplied endpoint protection contract; transfer not started.'
    }
    if($hash -ine $Manifest.ContractSHA256){
        throw 'Endpoint protection contract hash mismatch; deploy the matching reviewed contract. Transfer not started.'
    }
    return $contract
}

function Get-PreparedFileHashMap([string[]]$Paths){
    # Parallel reads only (the batch usually sits on a network share); callers still compare
    # every hash individually and in their own order. The first read error is rethrown.
    $results=@($Paths | Select-Object -Unique | ForEach-Object -ThrottleLimit 4 -Parallel {
        $path=$_
        try{[pscustomobject]@{Path=$path;Hash=(Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash;Error=$null}}
        catch{[pscustomobject]@{Path=$path;Hash=$null;Error=$_.Exception.Message}}
    })
    $hashes=@{}
    foreach($path in $Paths){
        $result=@($results | Where-Object Path -CEQ $path)[0]
        if($result.Error){throw $result.Error}
        $hashes[$path]=$result.Hash
    }
    $hashes
}

function Get-PreparedTransferPlan {
    param([string]$Root,[string]$TenantKey,[string]$ContractPath,[string]$ExpectedBatchId)
    $pointerPath=Get-PreparedMetadataPath $Root 'current';Assert-TransferPath $pointerPath
    $pointer=Get-Content -LiteralPath $pointerPath -Raw | ConvertFrom-Json
    $id=[string]$pointer.BatchId
    if($id -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$' -or $pointer.SchemaVersion -ne 1 -or $pointer.TenantKey -ne $TenantKey){throw 'Invalid local pointer identity.'}
    if($ExpectedBatchId -and $id -cne $ExpectedBatchId){throw 'Current batch differs from ExpectedBatchId; transfer not started.'}
    $folder=Join-Path $Root ('batches/'+$id);Assert-TransferPath $folder
    $metadata=@{}
    foreach($name in 'batch','validation','current'){
        $path=Get-PreparedMetadataPath $folder $name;Assert-TransferPath $path
        $metadata[$name]=[pscustomobject]@{Path=$path;Bytes=(Get-Item -LiteralPath $path).Length;SHA256=(Get-FileHash -LiteralPath $path).Hash;Value=(Get-Content -LiteralPath $path -Raw | ConvertFrom-Json)}
    }
    $manifest=$metadata.batch.Value;$validation=$metadata.validation.Value;$receipt=$metadata.current.Value
    if($manifest.SchemaVersion -ne 1 -or $manifest.TenantKey -ne $TenantKey -or $manifest.BatchId -ne $id -or $metadata.batch.SHA256 -ne $pointer.ManifestSHA256 -or $metadata.current.SHA256 -ne (Get-FileHash -LiteralPath $pointerPath).Hash){throw 'Manifest or publication receipt mismatch.'}
    if($receipt.SchemaVersion -ne 1 -or $validation.Passed -ne $true -or $validation.SchemaOnly -ne $false){throw 'Batch lacks a successful full validation receipt.'}
    $contract=Resolve-PreparedTransferContract -Manifest $manifest -ContractPath $ContractPath
    $expected=@($contract.tables.file)
    if(@($manifest.Files).Count -ne $expected.Count -or @($validation.Files).Count -ne $expected.Count -or $expected.Count -eq 0){throw 'Prepared file count differs from the contract.'}
    $files=[Collections.Generic.List[object]]::new();$seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $csvEntries=[Collections.Generic.List[object]]::new()
    foreach($entry in $manifest.Files){
        $name=[string]$entry.File
        if($name -notmatch '^[A-Za-z0-9_-]+\.csv$' -or $name -cnotin $expected -or -not $seen.Add($name)){throw 'Unexpected, unsafe or duplicate CSV name.'}
        $path=Join-Path $folder $name;Assert-TransferPath $path
        $validated=@($validation.Files | Where-Object File -CEQ $name)
        if($entry.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or $validated.Count -ne 1 -or $validated[0].SHA256 -ne $entry.SHA256 -or $validated[0].Rows -ne $entry.Rows -or (Get-Item -LiteralPath $path).Length -ne $entry.Bytes){throw "Local CSV integrity mismatch: $name"}
        $csvEntries.Add([pscustomobject]@{Name=$name;Path=$path;Entry=$entry})
    }
    $csvHashes=Get-PreparedFileHashMap @($csvEntries.Path)
    foreach($csv in $csvEntries){
        if($csvHashes[$csv.Path] -ne $csv.Entry.SHA256){throw "Local CSV integrity mismatch: $($csv.Name)"}
        $files.Add([pscustomobject]@{Path=$csv.Path;Relative="batches/$id/$($csv.Name)";Bytes=$csv.Entry.Bytes;SHA256=$csv.Entry.SHA256})
    }
    if($manifest.PSObject.Properties['Artifacts']){
        if(-not $manifest.PSObject.Properties['ArtifactContractSHA256'] -or -not (Test-PreparedArtifactContractCompatible -ContractSHA256 ([string]$manifest.ArtifactContractSHA256))){throw 'Artifact contract mismatch.'}
        Assert-PreparedArtifacts -BatchRoot $folder -Entries @($manifest.Artifacts) -TenantKey $TenantKey
        foreach($entry in $manifest.Artifacts){
            if(-not $seen.Add($entry.File)){throw 'Artifact collides with a prepared table.'}
            $files.Add([pscustomobject]@{Path=(Join-Path $folder $entry.File);Relative="batches/$id/$($entry.File)";Bytes=$entry.Bytes;SHA256=$entry.SHA256})
        }
    } elseif($manifest.Provenance.PSObject.Properties['NativeEvidence']) {throw 'Unified batch is missing its governed source artifacts.'}
    foreach($name in 'batch','validation','current'){
        $m=$metadata[$name];$files.Add([pscustomobject]@{Path=$m.Path;Relative="batches/$id/$name.json.txt";Bytes=$m.Bytes;SHA256=$m.SHA256})
    }
    # Immutable batch receipt is also the source for the final root pointer.
    $files.Add([pscustomobject]@{Path=$metadata.current.Path;Relative='current.json.txt';Bytes=$metadata.current.Bytes;SHA256=$metadata.current.SHA256})
    [pscustomobject]@{BatchId=$id;CsvCount=$expected.Count;PointerPath=$pointerPath;PointerHash=$metadata.current.SHA256;Files=@($files)}
}

function Send-PreparedEvidenceBatch {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputRoot,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][string]$ContractPath,[Parameter(Mandatory)][string]$WorkRoot,
        [Parameter(Mandatory)][scriptblock]$UploadFile,[Parameter(Mandatory)][scriptblock]$DownloadFile,
        [string]$ExpectedBatchId)
    $root=(Resolve-Path -LiteralPath $OutputRoot).ProviderPath.TrimEnd('\','/')
    if((Split-Path $root -Leaf) -ne 'DATA-POWERBI'){throw 'Transfer requires a dedicated DATA-POWERBI directory.'}
    Assert-TransferPath $root
    $scratch=[IO.Path]::GetFullPath($WorkRoot).TrimEnd('\','/')
    $data=Split-Path $root -Parent
    if($scratch -eq $data -or $scratch.StartsWith($data+'\',[StringComparison]::OrdinalIgnoreCase) -or $data.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)){
        throw 'Transfer scratch must be outside the tenant DATA tree and not its ancestor.'
    }
    Initialize-PreparedMetadataNames -OutputRoot $root -TenantKey $TenantKey
    $lock=[IO.File]::Open((Join-Path $root '.publication.lock'),'OpenOrCreate','ReadWrite','None')
    $run=$null;$temp=$null;$journal=$null
    try {
        Convert-PreparedAuditNames -Root $scratch -Family transfers -TenantKey $TenantKey
        $clock=[Diagnostics.Stopwatch]::StartNew()
        $plan=Get-PreparedTransferPlan $root $TenantKey $ContractPath $ExpectedBatchId
        Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] SharePoint transfer plan verified: {1} files; {2:F1}s' -f (Get-Date),$plan.Files.Count,$clock.Elapsed.TotalSeconds)
        $run=Join-Path $scratch ('transfers/'+[guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $run -Force | Out-Null;Assert-TransferPath $run
        $auditPath=Resolve-SmartM365OwnedJsonPath -Path (Join-Path $run 'transfer.json') -Owner 'WorkplaceEvidence-Prepare/transfer' -Validate {param($document) if($document.TenantKey -ne $TenantKey){throw 'Transfer audit tenant mismatch.'}}
        $temp=Join-Path $run 'readback.tmp'
        $journal=[ordered]@{TenantKey=$TenantKey;BatchId=$plan.BatchId;StartedUtc=[datetime]::UtcNow.ToString('O');Status='Transferring';PointerAttempted=$false;VerifiedFiles=[Collections.Generic.List[string]]::new();Error=$null}
        foreach($file in $plan.Files){
            if($file.Relative -eq 'current.json.txt'){
                # Verify local inputs again before activating the cloud batch.
                $clock=[Diagnostics.Stopwatch]::StartNew()
                if((Get-FileHash -LiteralPath $plan.PointerPath).Hash -ne $plan.PointerHash){throw 'Local current pointer changed during transfer.'}
                $currentHashes=Get-PreparedFileHashMap @($plan.Files.Path)
                foreach($source in $plan.Files){if($currentHashes[$source.Path] -ne $source.SHA256){throw 'Local batch changed during transfer; cloud pointer not advanced.'}}
                Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Local batch re-verified before the root pointer: {1} files; {2:F1}s' -f (Get-Date),$plan.Files.Count,$clock.Elapsed.TotalSeconds)
                $journal.PointerAttempted=$true
            }
            $clock=[Diagnostics.Stopwatch]::StartNew()
            # The plan already hashed every file and the read-back below compares the uploaded bytes with
            # that SHA256; the full local re-verification before the root pointer covers later changes.
            # A size check keeps the early failure without a third read of each file from the share.
            if((Get-Item -LiteralPath $file.Path).Length -ne $file.Bytes){throw "Local file changed: $($file.Relative)"}
            $localSeconds=$clock.Elapsed.TotalSeconds;$clock.Restart()
            if(-not (& $UploadFile $file.Path $file.Relative)){throw "Upload failed: $($file.Relative)"}
            $uploadSeconds=$clock.Elapsed.TotalSeconds;$clock.Restart()
            if(-not (& $DownloadFile $temp $file.Relative)){throw "Read-back failed: $($file.Relative)"}
            $downloadSeconds=$clock.Elapsed.TotalSeconds;$clock.Restart()
            Assert-TransferPath $temp
            if((Get-Item -LiteralPath $temp).Length -ne $file.Bytes -or (Get-FileHash -LiteralPath $temp).Hash -ne $file.SHA256){throw "Remote SHA256/size mismatch: $($file.Relative)"}
            Remove-Item -LiteralPath $temp -Force
            $journal.VerifiedFiles.Add($file.Relative)
            Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] SharePoint verified {1}/{2}: {3}; bytes={4}; localCheck={5:F1}s; upload={6:F1}s; readBack={7:F1}s; verify={8:F1}s' -f (Get-Date),$journal.VerifiedFiles.Count,$plan.Files.Count,$file.Relative,$file.Bytes,$localSeconds,$uploadSeconds,$downloadSeconds,$clock.Elapsed.TotalSeconds)
        }
        $journal.Status='Success'
        [pscustomobject]@{Status='PublishedAndReadBack';BatchId=$plan.BatchId;CsvFiles=$plan.CsvCount;VerifiedFiles=$journal.VerifiedFiles.Count;ExpectedFiles=$plan.Files.Count;ManifestSHA256=($plan.Files|Where-Object Relative -eq "batches/$($plan.BatchId)/batch.json.txt").SHA256;PointerSHA256=$plan.PointerHash;PointerVerified=$true;CsvRecalculated=$false;AuditPath=$auditPath}
    } catch {
        if($journal){$journal.Status='Failed';$journal.Error=$_.Exception.Message}
        if($journal -and $journal.PointerAttempted){throw "Cloud pointer publication was attempted but completion is not confirmed. Do not refresh Power BI; retry the same batch. $($_.Exception.Message)"}
        throw
    } finally {
        try{
            if($temp -and (Test-Path -LiteralPath $temp)){Assert-TransferPath $temp;Remove-Item -LiteralPath $temp -Force}
            if($journal){$journal['EndedUtc']=[datetime]::UtcNow.ToString('O');Write-SmartM365JsonBytesAtomically -Path $auditPath -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($journal | ConvertTo-Json -Depth 6))) -Validate {param($document) if($document.TenantKey -ne $TenantKey){throw 'Transfer audit tenant mismatch.'}} | Out-Null}
        }finally{$lock.Dispose()}
    }
}
function Assert-PreparedTransferResult {
    param([Parameter(Mandatory)]$Result,[Parameter(Mandatory)][string]$ExpectedBatchId)
    foreach($field in 'Status','BatchId','CsvFiles','VerifiedFiles','ExpectedFiles','ManifestSHA256','PointerVerified','CsvRecalculated'){
        if(-not $Result.PSObject.Properties[$field]){throw 'Verified transfer receipt is missing required fields.'}
    }
    if($Result.Status -cne 'PublishedAndReadBack' -or $Result.BatchId -cne $ExpectedBatchId -or
       $Result.PointerVerified -isnot [bool] -or -not $Result.PointerVerified -or
       $Result.CsvRecalculated -isnot [bool] -or $Result.CsvRecalculated -or
       $Result.CsvFiles -le 0 -or $Result.ExpectedFiles -lt ($Result.CsvFiles+4) -or
       $Result.VerifiedFiles -ne $Result.ExpectedFiles -or $Result.ManifestSHA256 -notmatch '^[A-Fa-f0-9]{64}$'){
        throw 'Transfer did not return a complete verified receipt for the expected batch.'
    }
}

function Send-PreparedFailureAudit {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputRoot,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][string]$BatchId,[Parameter(Mandatory)][string]$WorkRoot,
        [Parameter(Mandatory)][scriptblock]$UploadFile,[Parameter(Mandatory)][scriptblock]$DownloadFile)
    if ($BatchId -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$') { throw 'Invalid failure audit batch ID.' }
    $root=(Resolve-Path -LiteralPath $OutputRoot).ProviderPath.TrimEnd('\','/')
    if ((Split-Path $root -Leaf) -ne 'DATA-POWERBI') { throw 'Failure audit transfer requires DATA-POWERBI.' }
    Assert-TransferPath $root
    $folder=Join-Path $root ('failed/'+$BatchId)
    Assert-TransferPath $folder
    $failurePath=Join-Path $folder 'failure.json.txt'
    Assert-TransferPath $failurePath
    $failure=Get-Content -LiteralPath $failurePath -Raw | ConvertFrom-Json
    if ($failure.BatchId -cne $BatchId -or $failure.TenantKey -cne $TenantKey) { throw 'Failure audit identity mismatch.' }
    $files=@($failurePath)
    $validationPath=Join-Path $folder 'validation.json.txt'
    if (Test-Path -LiteralPath $validationPath -PathType Leaf) { Assert-TransferPath $validationPath; $files+= $validationPath }
    $scratch=[IO.Path]::GetFullPath($WorkRoot).TrimEnd('\','/')
    $data=Split-Path $root -Parent
    if ($scratch -eq $data -or $scratch.StartsWith($data+'\',[StringComparison]::OrdinalIgnoreCase) -or $data.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)) {
        throw 'Failure audit scratch must be outside the tenant DATA tree.'
    }
    New-Item -ItemType Directory -Path $scratch -Force | Out-Null
    Assert-TransferPath $scratch
    $temp=Join-Path $scratch ('failure-readback-'+[guid]::NewGuid().ToString('N')+'.tmp')
    try {
        foreach ($path in $files) {
            $relative='failed/'+$BatchId+'/'+[IO.Path]::GetFileName($path)
            $hash=(Get-FileHash -LiteralPath $path).Hash
            $bytes=(Get-Item -LiteralPath $path).Length
            if (-not (& $UploadFile $path $relative)) { throw "Failure audit upload failed: $relative" }
            if (-not (& $DownloadFile $temp $relative)) { throw "Failure audit read-back failed: $relative" }
            Assert-TransferPath $temp
            if ((Get-Item -LiteralPath $temp).Length -ne $bytes -or (Get-FileHash -LiteralPath $temp).Hash -ne $hash) { throw "Failure audit read-back mismatch: $relative" }
            Remove-Item -LiteralPath $temp -Force
        }
        [pscustomobject]@{BatchId=$BatchId;VerifiedFiles=$files.Count;PointerChanged=$false}
    } finally {
        if (Test-Path -LiteralPath $temp) { Assert-TransferPath $temp; Remove-Item -LiteralPath $temp -Force }
    }
}

Export-ModuleMember -Function Resolve-PreparedSharePointFolder,Send-PreparedEvidenceBatch,Send-PreparedFailureAudit,Assert-PreparedTransferResult

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBDSKSUogq612mi
# iw0ITwdiAdLY66tIOn3s0OUk7wriNKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPhKLaONVXqfr4EbLvHO+SFXK01T6xMnWkzqNh2fOtglMA0GCSqG
# SIb3DQEBAQUABIIBgI4zElLlWdaWwh8Eb4vbIgCRQ/oulPfQHnNkr42o2rtqRjIe
# Yvmt8CnZ2nojPPQ+pjzyqFuMunqR7MML9ajEXZkfBIrbjwVo2Ln4sbT95C5QcdDA
# Nq1RR7smj+3PCVI4kZfdAb783BZh6fH+BPEg9kg+KyjsVsR16uDj1RnCKGKLs+X5
# wSgrz5Rju+Nll+G4lq03E6htnapxtTz+hNHk3LIfVOWgyIZRMylbgUXsiFSe58zT
# H+Y33lrSHFkpYkQZkQBzMzFNWylzihxcr4yHunrT8WQnLRakrvoVCoZuyHZxl41C
# 2HkxKaWtoowNcJiQF6Jnlwf0Q2apqq7wv909bLPerGA+6WXvid7aidVO3VP3xQ7F
# pTci9IET7WqH72EbMZJXXgTVs7QsxLnJTknYgsP+WjYVf+wo+UDsYIX/wiKxeJ4m
# tDw5pptQfhysBzn52WS+uoSW2UENFj2396DMC8XB7uhZN0zx+CXYidAeEQeM3v9F
# JnhlEMqpJWSTNp72YaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkyMjI0
# MDhaMC8GCSqGSIb3DQEJBDEiBCBLajd3MAlbCChL0rRHvOfufqvYKzDPHoJdDrSc
# t6x2NzANBgkqhkiG9w0BAQEFAASCAgCteXiJ+UAHkvbKYeL+nCMDhj7tPbRhy/bU
# rHQK/MF5jMYQ72oDxW3zTFwdUOKdr+7WYT2VbByuGy3bLlrCr5wqB90R6y4ua5TY
# /9qqF6YUjTvhCHaReDjPDEqVDppEiwkwH17tKE4vPRNMuifSrZD5MCkUqLpbKDav
# 8z49wQPqpK/WFGPW6Xr+LhxscvQnrHkbNohK0gcgu5hZy3vf+dejFr+9NSzTGUJ3
# eIoBIa6gtVEJ32rHQSbtGVww7rcBZJbs16ZlJk6JkNLL1ND/hi1sho9jmkBiRBC3
# HRSL3nnNvY8xyyL1VoVC92YW5RUSFuorR2CI1vGp3Oczcqx3RBpj/5AiqKWLar0q
# U/sk8PzhIiBKRenX5eMeUeGhuL7xu3VZwFPYc7inf1mNG46p7YYWlB9/0cEfUmcY
# vh5cQE25QWFY+7tKe3MSvCkyxcG3rmM6JO/pCis+eXEbpsjsjth6d8o9AmQBg3FE
# qTq1c77+WecRvfC0eANXrGZriG+uNySSozWuxfxUtwIsWftdPZcWBCnw9dVsxIyI
# L/7EAqx1q//FUj12DchYZOkUiHbppJd8z0uaNH/mt9mNfAecTSs85a3Svxkz0qZM
# TtBwbSD89UGe3qtsHB5vUJKiPpuqmtI0z5uRXkKMRPSDQBHVAyyUrNUtw6sTJf4o
# htLFvCNgPw==
# SIG # End signature block
