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
    # Retain transfer of immutable pre-DailyTrends cohorts. A declared new
    # history cohort or a partial file pair must never fall back to legacy.
    $dailyFiles=@('DailyDirectoryTrends.csv','DailyMailboxTrends.csv')
    $dailyTables=@('Daily Directory Trends','Daily Mailbox Trends')
    $dailyPresent=@($Manifest.Files | Where-Object File -in $dailyFiles)
    $dailySupplied=@($contract.tables | Where-Object file -in $dailyFiles)
    if($dailyPresent.Count -eq 0){
        $dailyDeclared=$Manifest.PSObject.Properties['HistoryKeyVersions'] -and
            @($Manifest.HistoryKeyVersions.PSObject.Properties.Name | Where-Object {$_ -in $dailyTables}).Count -gt 0
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
    foreach($entry in $manifest.Files){
        $name=[string]$entry.File
        if($name -notmatch '^[A-Za-z0-9_-]+\.csv$' -or $name -cnotin $expected -or -not $seen.Add($name)){throw 'Unexpected, unsafe or duplicate CSV name.'}
        $path=Join-Path $folder $name;Assert-TransferPath $path
        $validated=@($validation.Files | Where-Object File -CEQ $name)
        if($entry.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or $validated.Count -ne 1 -or $validated[0].SHA256 -ne $entry.SHA256 -or $validated[0].Rows -ne $entry.Rows -or (Get-Item -LiteralPath $path).Length -ne $entry.Bytes -or (Get-FileHash -LiteralPath $path).Hash -ne $entry.SHA256){throw "Local CSV integrity mismatch: $name"}
        $files.Add([pscustomobject]@{Path=$path;Relative="batches/$id/$name";Bytes=$entry.Bytes;SHA256=$entry.SHA256})
    }
    if($manifest.PSObject.Properties['Artifacts']){
        $artifactContract=Join-Path $PSScriptRoot '../config/prepared-artifact-contract.json.txt'
        if(-not $manifest.PSObject.Properties['ArtifactContractSHA256'] -or $manifest.ArtifactContractSHA256 -ne (Get-FileHash -LiteralPath $artifactContract).Hash){throw 'Artifact contract mismatch.'}
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
                foreach($source in $plan.Files){if((Get-FileHash -LiteralPath $source.Path).Hash -ne $source.SHA256){throw 'Local batch changed during transfer; cloud pointer not advanced.'}}
                Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Local batch re-verified before the root pointer: {1} files; {2:F1}s' -f (Get-Date),$plan.Files.Count,$clock.Elapsed.TotalSeconds)
                $journal.PointerAttempted=$true
            }
            $clock=[Diagnostics.Stopwatch]::StartNew()
            if((Get-FileHash -LiteralPath $file.Path).Hash -ne $file.SHA256){throw "Local file changed: $($file.Relative)"}
            $localSeconds=$clock.Elapsed.TotalSeconds;$clock.Restart()
            if(-not (& $UploadFile $file.Path $file.Relative)){throw "Upload failed: $($file.Relative)"}
            $uploadSeconds=$clock.Elapsed.TotalSeconds;$clock.Restart()
            if(-not (& $DownloadFile $temp $file.Relative)){throw "Read-back failed: $($file.Relative)"}
            $downloadSeconds=$clock.Elapsed.TotalSeconds;$clock.Restart()
            Assert-TransferPath $temp
            if((Get-Item -LiteralPath $temp).Length -ne $file.Bytes -or (Get-FileHash -LiteralPath $temp).Hash -ne $file.SHA256){throw "Remote SHA256/size mismatch: $($file.Relative)"}
            Remove-Item -LiteralPath $temp -Force
            $journal.VerifiedFiles.Add($file.Relative)
            Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] SharePoint verified {1}/{2}: {3}; bytes={4}; localHash={5:F1}s; upload={6:F1}s; readBack={7:F1}s; verify={8:F1}s' -f (Get-Date),$journal.VerifiedFiles.Count,$plan.Files.Count,$file.Relative,$file.Bytes,$localSeconds,$uploadSeconds,$downloadSeconds,$clock.Elapsed.TotalSeconds)
        }
        $journal.Status='Success'
        [pscustomobject]@{Status='PublishedAndReadBack';BatchId=$plan.BatchId;CsvFiles=$plan.CsvCount;VerifiedFiles=$journal.VerifiedFiles.Count;ExpectedFiles=$plan.Files.Count;ManifestSHA256=($plan.Files|Where-Object Relative -eq "batches/$($plan.BatchId)/batch.json.txt").SHA256;PointerVerified=$true;CsvRecalculated=$false;AuditPath=$auditPath}
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAD8YFgcSOLIDSm
# CClW+ruQAyMAGyw3hTAVwqOBnFkrPqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBqpIpp3ocH/Rbv3Ofxkz8d
# 71sAwRsEQgA+7McgjI4tGDANBgkqhkiG9w0BAQEFAASCAYBXA1zFTbzoDLaH3xKy
# +21BZ4WbNWlndEr9zJhrnfXfuC2AJznqP7aNdH7ZA4b+1KkSkWhIYLDTWb4s/hlu
# sHuxpt4iqg/04/aVWYHJHYoKQELdw6sDrmXkgCvsZb03eRqduYWyGRtYK+5Oh9ss
# y161FW6oZ0cyj/x7KOibUs//uzOj0vSMnCqDhrJ/6IQMOVC2acFL1ipoLIDzn8Zg
# PcZyl5CzIfsLOxiJlVhYSXROSrs3h2+KMM6+bxQF/grONKnxt0CbdIdzoSrTeF8B
# 6L8norwLODdxJmDZujyeO600PdAvG+v2kpXqGO4FKoOlq9JM3DKFx4DZ27XJz0FV
# R58HqJQeExGn3i3ePo1aLMRMYbLFT8rNKc/fdFbZJYjWXrUL7wOApnkeN+FUp2w1
# JGiUdZFZNgabsAN5lJFQyJK9mimnuwEvg2osU8b0t3OWZeWBEoFvZ3FiuFrpospG
# 9UWY/wZv3l4TLMqqvtTbOKMxzRDkA5RvwgIVPCw+RcvtasU=
# SIG # End signature block
