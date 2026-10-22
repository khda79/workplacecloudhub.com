<#
.SYNOPSIS
    Qualified, coupled server/SharePoint batch retention. Preview is the default.
.VERSION
    1.4.0
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../../SmartM365/Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.6'
$script:RetentionPolicyVersion = 'PreparedBatchRetention/v1'
$script:RetentionModulePath = $PSCommandPath
$script:RetentionModuleSHA256 = (Get-FileHash -LiteralPath $script:RetentionModulePath).Hash

function Get-RetentionHash([byte[]]$Bytes) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-','')}
    finally {$sha.Dispose()}
}
function Get-RetentionObjectHash($Value) {
    Get-RetentionHash ([Text.Encoding]::UTF8.GetBytes(($Value | ConvertTo-Json -Depth 70 -Compress)))
}
function Get-RetentionValue($Value,[string]$Name,$Default=$null) {
    if ($Value -is [Collections.IDictionary]) { if($Value.Contains($Name)){return $Value[$Name]} }
    elseif ($null -ne $Value -and $Value.PSObject.Properties[$Name]) { return $Value.$Name }
    $Default
}
function Get-RetentionHttpStatus($ErrorRecord) {
    $exception=$ErrorRecord.Exception
    for($depth=0;$depth -lt 8 -and $exception;$depth++){
        foreach($key in 'ResponseStatusCode','StatusCode'){
            $value=Get-RetentionValue $exception $key
            if($null -ne $value){return [int]$value}
        }
        $response=Get-RetentionValue $exception 'Response'
        if($null -ne $response){$value=Get-RetentionValue $response 'StatusCode';if($null -ne $value){return [int]$value}}
        $exception=$exception.InnerException
    }
    0
}
function Assert-RetentionPath([string]$Path) {
    $cursor=[IO.Path]::GetFullPath($Path);$volume=[IO.Path]::GetPathRoot($cursor).TrimEnd('\','/')
    while($cursor){
        if(Test-Path -LiteralPath $cursor){
            if((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Retention refuses linked or synchronized placeholder paths.'}
        }
        if($cursor.TrimEnd('\','/') -eq $volume){break}
        $parent=Split-Path $cursor -Parent;if($parent -eq $cursor){break};$cursor=$parent
    }
}
function Get-RetentionChild([string]$Root,[string]$Relative) {
    if($Relative -match '[:%?#]' -or @($Relative.Replace('\','/').Split('/') | Where-Object {$_ -in @('','.','..')}).Count){throw 'Unsafe retention relative path.'}
    $base=[IO.Path]::GetFullPath($Root).TrimEnd('\','/')
    $path=[IO.Path]::GetFullPath((Join-Path $base $Relative))
    if(-not $path.StartsWith($base+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Retention path escapes its exact root.'}
    Assert-RetentionPath $path
    $path
}
function Assert-RetentionBatchId([string]$Id) {
    if($Id -cnotmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$'){throw 'Invalid retention batch ID.'}
    $null=[datetime]::ParseExact($Id.Substring(0,19),"yyyyMMdd'T'HHmmssfff'Z'",[Globalization.CultureInfo]::InvariantCulture,
        ([Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal))
}
function Assert-RetentionLock([IO.FileStream]$Stream) {
    if(-not $Stream.CanWrite -or $Stream.SafeFileHandle.IsClosed -or $Stream.SafeFileHandle.IsInvalid){throw 'Server publication lock was lost.'}
    # A durable round trip detects an unusable SMB handle before cloud mutations.
    $pulse=[guid]::NewGuid().ToByteArray();$read=[byte[]]::new($pulse.Length)
    $Stream.Position=0;$Stream.Write($pulse,0,$pulse.Length);$Stream.Flush($true);$Stream.Position=0
    if($Stream.Read($read,0,$read.Length) -ne $read.Length -or [Convert]::ToBase64String($pulse) -cne [Convert]::ToBase64String($read)){throw 'Server publication lock round-trip failed.'}
}
function Assert-RetentionReviewedUtc($Value) {
    if($Value -is [datetime]){
        if($Value.Kind -ne [DateTimeKind]::Utc){throw 'Review must explicitly use UTC.'}
        $review=[DateTimeOffset]$Value
    } else {
        if([string]$Value -notmatch '(Z|\+00:00)$'){throw 'Review must explicitly use UTC.'}
        $review=[DateTimeOffset]::Parse([string]$Value)
    }
    if($review -gt [DateTimeOffset]::UtcNow){throw 'Review cannot be in the future.'}
}
function Write-RetentionDocument([string]$Path,$Value) {
    Assert-RetentionPath $Path
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($Value | ConvertTo-Json -Depth 70))
    Write-SmartM365JsonBytesAtomically -Path $Path -Bytes $bytes -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Retention document must be an object.'}} | Out-Null
}
function Read-RetentionLockedDocument([string]$Path) {
    Assert-RetentionPath $Path
    # Deny writes and rename for the entire plan/application operation, including on SMB.
    $stream=[IO.File]::Open($Path,'Open','Read','Read')
    try {
        $memory=[IO.MemoryStream]::new();$stream.CopyTo($memory);$bytes=$memory.ToArray();$memory.Dispose()
        $value=[Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF) | ConvertFrom-Json -Depth 70
        [pscustomobject]@{Stream=$stream;Value=$value;SHA256=(Get-RetentionHash $bytes)}
    } catch {$stream.Dispose();throw}
}
function Assert-RetentionRegistry($Registry,[string]$TenantKey,[string]$TenantId,[switch]$ReferencesOnly) {
    if($Registry.SchemaVersion -is [string] -or $Registry.SchemaVersion -ne 1 -or $Registry.TenantKey -cne $TenantKey -or $Registry.TenantId -ine $TenantId){throw 'Consumer registry schema/tenant mismatch.'}
    if(-not $ReferencesOnly){
    $qualification=$Registry.Qualification
    if($qualification.Status -cne 'Qualified' -or $qualification.DeclaredConsumersComplete -isnot [bool] -or $qualification.DeclaredConsumersComplete -ne $true -or
       [string]::IsNullOrWhiteSpace($qualification.ReviewedBy)){throw 'Consumer registry is not explicitly qualified.'}
    Assert-RetentionReviewedUtc $qualification.ReviewedUtc
    }
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($entry in @($Registry.Consumers)){
        Assert-RetentionBatchId $entry.BatchId
        if([string]::IsNullOrWhiteSpace($entry.ConsumerId) -or [string]::IsNullOrWhiteSpace($entry.Reason) -or
           -not $seen.Add($entry.ConsumerId+'|'+$entry.BatchId)){throw 'Invalid or duplicate consumer reference.'}
        Assert-RetentionReviewedUtc $entry.ReviewedUtc
    }
}
function Assert-RetentionTarget($Actual,$Expected) {
    foreach($key in 'TenantId','ClientId','SiteId','DriveId','RootId','BatchesId','RootPath'){
        if([string]::IsNullOrWhiteSpace([string](Get-RetentionValue $Expected $key)) -or
           (Get-RetentionValue $Actual $key) -cne (Get-RetentionValue $Expected $key)){throw "Retention target mismatch: $key"}
    }
    $path=[string]$Actual.RootPath
    if($path -match '[:%?#\\]' -or $path.Split('/')[-1] -cne 'DATA-POWERBI' -or
       @($path.Split('/') | Where-Object {$_ -in @('','.','..','DATA-LAST','DATA-ALL','LOG-ALL')}).Count){throw 'Invalid exact SharePoint DATA-POWERBI root.'}
    if($Actual.RootId -ceq $Actual.BatchesId){throw 'Batches must be a distinct child of DATA-POWERBI.'}
}
function Assert-RetentionQualification($Settings,$Target,[string]$RegistryHash,[string]$OutputRoot) {
    $q=$Settings.Qualification
    if($q.Status -cne 'Qualified' -or $q.PolicyVersion -cne $script:RetentionPolicyVersion -or
       [string]::IsNullOrWhiteSpace($q.ReviewedBy) -or $q.ExclusiveCloudWritersConfirmed -isnot [bool] -or $q.ExclusiveCloudWritersConfirmed -ne $true -or
       $q.Coordination -cne 'SharedServerPublicationLock' -or $q.RecoverabilityConfirmed -isnot [bool] -or $q.RecoverabilityConfirmed -ne $true){throw 'Retention policy/coordination is not qualified.'}
    Assert-RetentionReviewedUtc $q.ReviewedUtc
    if($q.RegistrySHA256 -cne $RegistryHash -or $q.ModuleSHA256 -cne $script:RetentionModuleSHA256 -or
       (Get-FileHash -LiteralPath $script:RetentionModulePath).Hash -cne $script:RetentionModuleSHA256){throw 'Registry or loaded retention implementation requires qualification.'}
    if([IO.Path]::GetFullPath($q.ServerOutputRoot).TrimEnd('\','/') -cne $OutputRoot){throw 'Server authority differs from qualification.'}
    if($q.AccessEvidenceSHA256 -cnotmatch '^[A-F0-9]{64}$' -or $q.AccessEvidenceSHA256 -cne $Target.AccessEvidenceSHA256){throw 'Live writer access evidence differs from qualification.'}
    # The review must explain exclusions of other apps, inherited writers and sync clients.
    if([string]::IsNullOrWhiteSpace($q.WriterExclusionEvidence)){throw 'Writer exclusion evidence is required; a boolean is insufficient.'}
}
function Read-RetentionRemoteFile($Transport,$Item,[string]$Scratch) {
    $temp=Get-RetentionChild $Scratch ([guid]::NewGuid().ToString('N')+'.readback')
    try {
        if((& $Transport.Download $Item.Id $temp) -ne $true){throw 'Retention remote read-back failed.'}
        Assert-RetentionPath $temp
        if((Get-Item -LiteralPath $temp).Length -ne [long]$Item.Bytes){throw 'Retention remote size mismatch.'}
        $hash=(Get-FileHash -LiteralPath $temp).Hash
        $value=$null
        if($Item.Name -in @('current.json.txt','batch.json.txt','validation.json.txt')){$value=Get-Content -LiteralPath $temp -Raw | ConvertFrom-Json -Depth 70}
        [pscustomobject]@{Name=$Item.Name;Id=$Item.Id;ETag=$Item.ETag;Bytes=[long]$Item.Bytes;SHA256=$hash;Value=$value}
    } finally {if(Test-Path -LiteralPath $temp){Assert-RetentionPath $temp;[IO.File]::Delete($temp)}}
}
function Get-RetentionRemoteChildren($Transport,[string]$ParentId) {
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $ids=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($item in @(& $Transport.List $ParentId | Sort-Object Name)){
        if($item.ParentId -cne $ParentId -or [string]::IsNullOrWhiteSpace($item.Id) -or
           $item.Name -notmatch '^[A-Za-z0-9_.-]+$' -or $item.Name -in @('.','..') -or
           $item.Linked -or -not $seen.Add($item.Name) -or -not $ids.Add($item.Id)){throw 'Unexpected/linked/duplicate remote item or parent mismatch.'}
        $item
    }
}
function Get-RetentionTechnicalLockOwner([string]$Name,$ExpectedNames) {
    # Only sidecars emitted by JsonTransport / SharePointJsonTransition, for a declared JSON.
    foreach($owner in $ExpectedNames){
        if($owner -cmatch '\.json\.txt$' -and
           ($Name -ceq ($owner.Substring(0,$owner.Length-4)+'.transport.lock') -or
            $Name -ceq ($owner+'.sharepoint-transition.lock'))){return $owner}
    }
    $null
}
function Open-RetentionTechnicalLock($File) {
    Assert-RetentionPath $File.Id
    if((Split-Path $File.Id -Leaf) -cne $File.Name -or $File.Bytes -ne 0){throw 'Invalid technical lock identity/size.'}
    $handle=$null
    try {
        # Open existing only; deny helper writers, allow hashing and final directory removal.
        $handle=[IO.File]::Open($File.Id,'Open','Read',([IO.FileShare]::Read -bor [IO.FileShare]::Delete))
        if($handle.Length -ne 0){throw 'Technical lock must be empty.'}
        $handle
    } catch {if($handle){$handle.Dispose()};throw "Technical lock unavailable or nonempty: $($File.Name)"}
}
function Get-RetentionFileEvidence($Files) {
    # Raw receipt SHA256 binds every original byte, including unprojected provenance.
    foreach($file in $Files){$file | Select-Object Name,Id,ETag,Bytes,SHA256}
}
function Get-RetentionSnapshotEvidence($Snapshot) {
    $batches=@(foreach($batch in $Snapshot.Batches){
        [pscustomobject]@{Storage=$batch.Storage;BatchId=$batch.BatchId;Id=$batch.Id;ETag=$batch.ETag;
            Invalid=$batch.Invalid;Verification=$batch.Verification;Files=@(Get-RetentionFileEvidence $batch.Files)}
    })
    [pscustomobject]@{Verification=$Snapshot.Verification;Target=$Snapshot.Target;LocalPointer=$Snapshot.LocalPointer;
        LocalPointerSHA256=$Snapshot.LocalPointerSHA256;RemotePointer=$Snapshot.RemotePointer;
        RemotePointerSHA256=$Snapshot.RemotePointerSHA256;Batches=$batches;ArchivedLinks=$Snapshot.ArchivedLinks}
}
function Test-RetentionBatch($Batch,[string]$TenantKey) {
    try {
        $metadataOnly=(Get-RetentionValue $Batch 'Verification' 'Full') -ceq 'Metadata'
        Assert-RetentionBatchId $Batch.BatchId
        if($Batch.Invalid){throw $Batch.Invalid}
        if($Batch.Storage -ceq 'SharePoint' -and [string]::IsNullOrWhiteSpace($Batch.ETag)){throw 'Remote batch lacks a conditional removal eTag.'}
        $map=@{};foreach($file in @($Batch.Files)){if($map.ContainsKey($file.Name)){throw 'Duplicate batch file.'};$map[$file.Name]=$file}
        foreach($name in 'current','batch','validation'){if(-not $map.ContainsKey($name+'.json.txt')){throw 'Incomplete legacy metadata; manual review required.'}}
        $p=$map['current.json.txt'].Value;$m=$map['batch.json.txt'].Value;$v=$map['validation.json.txt'].Value
        if($p.SchemaVersion -ne 1 -or $m.SchemaVersion -ne 1 -or $p.BatchId -cne $Batch.BatchId -or $m.BatchId -cne $Batch.BatchId -or
           $p.TenantKey -cne $TenantKey -or $m.TenantKey -cne $TenantKey -or $p.ManifestSHA256 -ine $map['batch.json.txt'].SHA256){throw 'Batch ownership/manifest mismatch.'}
        if($p.PreviousBatchId){Assert-RetentionBatchId $p.PreviousBatchId;if($p.PreviousBatchId -ceq $Batch.BatchId){throw 'Self-referencing batch.'}}
        if($v.Passed -isnot [bool] -or $v.Passed -ne $true -or $v.SchemaOnly -isnot [bool] -or $v.SchemaOnly -ne $false -or @(Get-RetentionValue $v 'Errors' @()).Count){throw 'Full successful validation is required.'}
        $expected=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach($name in 'current.json.txt','batch.json.txt','validation.json.txt'){$null=$expected.Add($name)}
        if(@($m.Files).Count -eq 0 -or @($v.Files).Count -ne @($m.Files).Count){throw 'Incomplete CSV validation records.'}
        foreach($entry in @($m.Files)+@(Get-RetentionValue $m 'Artifacts' @())){
            $name=[string]$entry.File
            if($name -cnotmatch '^[A-Za-z0-9_-]+\.(csv|json\.txt)$' -or -not $expected.Add($name) -or -not $map.ContainsKey($name)){throw 'Unexpected/unsafe/missing declared file.'}
            if(($entry.Bytes -isnot [int] -and $entry.Bytes -isnot [long]) -or $entry.Bytes -lt 0 -or $entry.SHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or [long]$entry.Bytes -ne $map[$name].Bytes -or
               (-not $metadataOnly -and $entry.SHA256 -ine $map[$name].SHA256)){throw 'Declared file hash/size mismatch.'}
        }
        foreach($entry in @($m.Files)){
            $validated=@($v.Files | Where-Object File -CEQ $entry.File)
            if(($entry.Rows -isnot [int] -and $entry.Rows -isnot [long]) -or $entry.Rows -lt 0 -or $validated.Count -ne 1 -or $validated[0].SHA256 -ine $entry.SHA256 -or $validated[0].Rows -ne $entry.Rows){throw 'CSV validation record mismatch.'}
        }
        $technicalLocks=@(foreach($file in $Batch.Files){
            if($expected.Contains($file.Name)){continue}
            if($Batch.Storage -cne 'Local' -or -not (Get-RetentionTechnicalLockOwner $file.Name $expected)){throw 'Unexpected batch content.'}
            if($file.Id -cne (Get-RetentionChild $Batch.Id $file.Name)){throw 'Technical lock escaped batch.'}
            $handle=Open-RetentionTechnicalLock $file
            try {$file.Name} finally {$handle.Dispose()}
        })
        [pscustomobject]@{Valid=$true;Reason=$(if($metadataOnly){'Receipts and file metadata verified; payload SHA256 not measured'}else{'Manifest, receipts and all payload hashes verified'});PreviousBatchId=$p.PreviousBatchId;TechnicalLockNames=$technicalLocks}
    } catch {[pscustomobject]@{Valid=$false;Reason=$_.Exception.Message;PreviousBatchId=$null;TechnicalLockNames=@()}}
}
function Get-RetentionSnapshot([string]$OutputRoot,$Transport,[string]$Scratch,[string]$TenantKey,[ValidateSet('Metadata','Full')][string]$Verification='Full') {
    $target=& $Transport.Describe
    $pointerPath=Get-RetentionChild $OutputRoot 'current.json.txt'
    $localPointer=Get-Content -LiteralPath $pointerPath -Raw | ConvertFrom-Json
    $localHash=(Get-FileHash -LiteralPath $pointerPath).Hash
    $remoteChildren=@(Get-RetentionRemoteChildren $Transport $target.RootId)
    $remotePointer=@($remoteChildren | Where-Object Name -CEQ 'current.json.txt')
    $remoteBatches=@($remoteChildren | Where-Object Name -CEQ 'batches')
    if($remotePointer.Count -ne 1 -or $remotePointer[0].Kind -cne 'File' -or $remoteBatches.Count -ne 1 -or
       $remoteBatches[0].Kind -cne 'Folder' -or $remoteBatches[0].Id -cne $target.BatchesId){throw 'Remote pointer/batches root identity mismatch.'}
    $cloudPointer=Read-RetentionRemoteFile $Transport $remotePointer[0] $Scratch
    foreach($p in @($localPointer,$cloudPointer.Value)){
        Assert-RetentionBatchId $p.BatchId
        if($p.PreviousBatchId){Assert-RetentionBatchId $p.PreviousBatchId}
        if($p.SchemaVersion -ne 1 -or $p.TenantKey -cne $TenantKey -or $p.ManifestSHA256 -notmatch '^[A-Fa-f0-9]{64}$'){throw 'Invalid retention pointer.'}
    }
    $batches=[Collections.Generic.List[object]]::new()
    $localBatches=Get-RetentionChild $OutputRoot 'batches'
    foreach($directory in Get-ChildItem -LiteralPath $localBatches -Force | Sort-Object Name){
        $files=[Collections.Generic.List[object]]::new();$invalid=''
        try {
            Assert-RetentionPath $directory.FullName
            if(-not $directory.PSIsContainer){throw 'Unexpected item in batches root.'}
            foreach($file in Get-ChildItem -LiteralPath $directory.FullName -Force | Sort-Object Name){
                Assert-RetentionPath $file.FullName
                if($file.PSIsContainer){throw 'Unexpected nested folder in batch.'}
                $value=$null
                if($file.Name -in @('current.json.txt','batch.json.txt','validation.json.txt')){$value=Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json -Depth 70}
                $hash=if($Verification -ceq 'Full' -or $file.Name -in @('current.json.txt','batch.json.txt','validation.json.txt')){(Get-FileHash -LiteralPath $file.FullName).Hash}else{$null}
                $files.Add([pscustomobject]@{Name=$file.Name;Id=$file.FullName;ETag='';Bytes=$file.Length;SHA256=$hash;Value=$value})
            }
        } catch {$invalid=$_.Exception.Message}
        $batches.Add([pscustomobject]@{Storage='Local';BatchId=$directory.Name;Id=$directory.FullName;ETag='';Files=$files.ToArray();Invalid=$invalid;Verification=$Verification})
    }
    foreach($directory in @(Get-RetentionRemoteChildren $Transport $target.BatchesId)){
        Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] Retention {1}: SharePoint batch {2}' -f (Get-Date),$Verification,$directory.Name)
        $files=[Collections.Generic.List[object]]::new();$invalid=''
        try {
            if($directory.Kind -cne 'Folder'){throw 'Unexpected remote batch item.'}
            foreach($file in @(Get-RetentionRemoteChildren $Transport $directory.Id)){
                if($file.Kind -cne 'File'){throw 'Unexpected nested remote batch folder.'}
                if($Verification -ceq 'Full' -or $file.Name -in @('current.json.txt','batch.json.txt','validation.json.txt')){
                    $files.Add((Read-RetentionRemoteFile $Transport $file $Scratch))
                }else{
                    $files.Add([pscustomobject]@{Name=$file.Name;Id=$file.Id;ETag=$file.ETag;Bytes=$file.Bytes;SHA256=$null;Value=$null})
                }
            }
        } catch {$invalid=$_.Exception.Message}
        $batches.Add([pscustomobject]@{Storage='SharePoint';BatchId=$directory.Name;Id=$directory.Id;ETag=$directory.ETag;Files=$files.ToArray();Invalid=$invalid;Verification=$Verification})
    }
    $archives=[Collections.Generic.List[object]]::new()
    $retired=Get-RetentionChild $OutputRoot 'retired'
    if(Test-Path -LiteralPath $retired){
        foreach($dir in Get-ChildItem -LiteralPath $retired -Directory -Force | Sort-Object Name){
            try {
                Assert-RetentionBatchId $dir.Name;Assert-RetentionPath $dir.FullName
                $cp=Get-RetentionChild $dir.FullName 'current.json.txt';$mp=Get-RetentionChild $dir.FullName 'batch.json.txt';$vp=Get-RetentionChild $dir.FullName 'validation.json.txt'
                $c=Get-Content -LiteralPath $cp -Raw | ConvertFrom-Json;$m=Get-Content -LiteralPath $mp -Raw | ConvertFrom-Json;$v=Get-Content -LiteralPath $vp -Raw | ConvertFrom-Json
                if($c.BatchId -cne $dir.Name -or $m.BatchId -cne $dir.Name -or $c.TenantKey -cne $TenantKey -or $m.TenantKey -cne $TenantKey -or
                   $c.SchemaVersion -ne 1 -or $m.SchemaVersion -ne 1 -or $c.ManifestSHA256 -ine (Get-FileHash -LiteralPath $mp).Hash -or $v.Passed -ne $true -or $v.SchemaOnly -ne $false){throw 'Invalid archived publication link.'}
                if($c.PreviousBatchId){Assert-RetentionBatchId $c.PreviousBatchId}
                $archives.Add([pscustomobject]@{BatchId=$dir.Name;PreviousBatchId=$c.PreviousBatchId;CurrentSHA256=(Get-FileHash -LiteralPath $cp).Hash;ManifestSHA256=$c.ManifestSHA256})
            } catch {Write-Verbose "Archived chain link ignored: $($_.Exception.Message)"}
        }
    }
    [pscustomobject]@{Verification=$Verification;Target=$target;LocalPointer=$localPointer;LocalPointerSHA256=$localHash;RemotePointer=$cloudPointer.Value;RemotePointerSHA256=$cloudPointer.SHA256;Batches=$batches.ToArray();ArchivedLinks=$archives.ToArray()}
}
function Get-RetentionIntegerSetting($Settings,[string]$Name,[int]$Default,[int]$Minimum,[int]$Maximum,[int[]]$Allowed=@()) {
    $value=Get-RetentionValue $Settings $Name $null
    if($null -eq $value){return $Default}
    if(($value -isnot [int] -and $value -isnot [long]) -or (($value -lt $Minimum -or $value -gt $Maximum) -and $value -notin $Allowed)){
        throw "PreparedRetention.$Name must be an integer in $Minimum..$Maximum$(if($Allowed.Count){' or '+($Allowed -join ', ')})."
    }
    [int]$value
}
function Get-RetentionPolicy($Settings) {
    # Optional keys with code defaults: adding them to the template would rewrite the private configuration.
    $sharePoint=Get-RetentionIntegerSetting $Settings 'SharePointKeepBatches' 2 2 50
    $server=Get-RetentionIntegerSetting $Settings 'ServerKeepBatches' 0 2 100000 @(0)
    if($server -ne 0 -and $server -lt $sharePoint){throw 'PreparedRetention.ServerKeepBatches must be 0 (keep the full server history) or at least SharePointKeepBatches.'}
    $removals=Get-RetentionIntegerSetting $Settings 'MaxSharePointRemovalsPerRun' 1 1 50
    [pscustomobject]@{SharePointKeepBatches=$sharePoint;ServerKeepBatches=$server;MaxSharePointRemovalsPerRun=$removals}
}
function Get-RetentionMinimumBatches($Snapshot,$Validation,$Policy) {
    foreach($storage in 'Local','SharePoint'){
        # A full server history still verifies its newest copies before every SharePoint removal.
        $keep=if($storage -ceq 'SharePoint' -or $Policy.ServerKeepBatches -eq 0){$Policy.SharePointKeepBatches}else{$Policy.ServerKeepBatches}
        # IDs carry the validated UTC creation timestamp; file modification times are irrelevant.
        $qualified=@($Snapshot.Batches | Where-Object {$_.Storage -ceq $storage -and $Validation[$storage+'|'+$_.BatchId].Valid} |
            Sort-Object @{Expression={$_.BatchId.Substring(0,19)};Descending=$true},BatchId)
        $ids=@()
        if($qualified.Count){
            $boundary=$qualified[[Math]::Min($keep-1,$qualified.Count-1)].BatchId.Substring(0,19)
            # Keep every tie at the boundary timestamp instead of arbitrarily dropping a valid copy.
            $ids=@($qualified | Where-Object {[string]::CompareOrdinal($_.BatchId.Substring(0,19),$boundary) -ge 0} | ForEach-Object BatchId)
        }
        [pscustomobject]@{Storage=$storage;Keep=$keep;BatchIds=$ids}
    }
}
function New-RetentionPlan($Snapshot,$Registry,[string]$RegistryHash,[string]$ConfigurationHash,[string]$OutputRoot,$Settings,[string]$TenantKey) {
    $blocked=[Collections.Generic.List[string]]::new();$protected=@{}
    function Protect([string]$Id,[string]$Reason){if($Id){if(-not $protected.ContainsKey($Id)){$protected[$Id]=[Collections.Generic.List[string]]::new()};$protected[$Id].Add($Reason)}}
    foreach($store in 'Local','Remote'){
        $p=$Snapshot.($store+'Pointer');Protect $p.BatchId ($store+' current');Protect $p.PreviousBatchId ($store+' previous')
    }
    if($Snapshot.LocalPointerSHA256 -cne $Snapshot.RemotePointerSHA256){$blocked.Add('Local and remote publication pointers differ.')}
    try {Assert-RetentionRegistry $Registry $TenantKey $Snapshot.Target.TenantId -ReferencesOnly;foreach($entry in @($Registry.Consumers)){Protect $entry.BatchId ('Consumer: '+$entry.ConsumerId+'; '+$entry.Reason)};Assert-RetentionRegistry $Registry $TenantKey $Snapshot.Target.TenantId}
    catch {$blocked.Add($_.Exception.Message)}
    try {Assert-RetentionTarget $Snapshot.Target $Settings.Target;Assert-RetentionQualification $Settings $Snapshot.Target $RegistryHash $OutputRoot}
    catch {$blocked.Add($_.Exception.Message)}
    try {$policy=Get-RetentionPolicy $Settings}
    catch {$blocked.Add($_.Exception.Message);$policy=[pscustomobject]@{SharePointKeepBatches=2;ServerKeepBatches=0;MaxSharePointRemovalsPerRun=1}}
    $rows=[Collections.Generic.List[object]]::new();$valid=@{};$localProtected=@{}
    foreach($b in @($Snapshot.Batches)){$valid[$b.Storage+'|'+$b.BatchId]=Test-RetentionBatch $b $TenantKey}
    $minimum=@(Get-RetentionMinimumBatches $Snapshot $valid $policy)
    foreach($store in $minimum){
        foreach($id in $store.BatchIds){
            $reason="Newest valid $($store.Storage) batches; keep $($store.Keep); $($Snapshot.Verification)"
            # A SharePoint selection protects both copies; a longer server history protects only the server copy.
            if($store.Storage -ceq 'SharePoint'){Protect $id $reason}elseif(-not $localProtected.ContainsKey($id)){$localProtected[$id]=$reason}
        }
        if($store.BatchIds.Count -lt 2){$blocked.Add("Fewer than two complete validated $($store.Storage) batches; retention minimum cannot be established.")}
    }
    foreach($storage in 'Local','SharePoint'){
        $byId=@{};foreach($b in @($Snapshot.Batches | Where-Object Storage -CEQ $storage)){$byId[$b.BatchId]=$b}
        $chain=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $next=$Snapshot.LocalPointer.BatchId
        while($next){
            if(-not $chain.Add($next)){$blocked.Add("Cyclic $storage publication chain.");break}
            if(-not $byId.ContainsKey($next)){
                $link=@($Snapshot.ArchivedLinks | Where-Object BatchId -CEQ $next)
                if($link.Count -ne 1){break};$next=$link[0].PreviousBatchId;continue
            }
            $result=$valid[$storage+'|'+$next]
            if(-not $result.Valid){break};$next=$result.PreviousBatchId
        }
        $current=$Snapshot.LocalPointer.BatchId
        if(-not $byId.ContainsKey($current) -or -not $valid[$storage+'|'+$current].Valid){$blocked.Add("Current $storage batch is not fully qualified.")}
        elseif(($byId[$current].Files | Where-Object Name -CEQ 'batch.json.txt').SHA256 -ine $Snapshot.LocalPointer.ManifestSHA256){$blocked.Add("Current $storage manifest differs from pointer.")}
        foreach($b in @($byId.Values | Sort-Object BatchId)){
            $check=$valid[$storage+'|'+$b.BatchId];$state='Blocked';$reason=$check.Reason
            if($protected.ContainsKey($b.BatchId)){$state='Protected';$reason=$protected[$b.BatchId].ToArray() -join '; '}
            elseif($storage -ceq 'Local' -and $localProtected.ContainsKey($b.BatchId)){$state='Protected';$reason=$localProtected[$b.BatchId]}
            elseif($storage -ceq 'Local' -and $policy.ServerKeepBatches -eq 0){$state='Retained';$reason='Server share keeps its full batch history (ServerKeepBatches=0).'}
            elseif($check.Valid -and $chain.Contains($b.BatchId)){$state=if($Snapshot.Verification -ceq 'Metadata'){'Candidate'}else{'Eligible'}}
            elseif($check.Valid){$reason='No verified published-chain membership.'}
            $logicalBytes=0L;foreach($file in $b.Files){$logicalBytes+=[long]$file.Bytes}
            $rows.Add([pscustomobject]@{Storage=$storage;BatchId=$b.BatchId;Id=$b.Id;ETag=$b.ETag;State=$state;Reason=$reason;Valid=$check.Valid;ValidationReason=$check.Reason;Verification=$Snapshot.Verification;PayloadSHA256Verified=($check.Valid -and $Snapshot.Verification -ceq 'Full');Files=@(Get-RetentionFileEvidence $b.Files);TechnicalLockNames=@($check.TechnicalLockNames);FileCount=$b.Files.Count;LogicalBytes=$logicalBytes})
        }
    }
    foreach($id in $protected.Keys){if(-not @($Snapshot.Batches | Where-Object BatchId -CEQ $id).Count){$blocked.Add("Protected batch is absent from both stores: $id")}}
    foreach($group in @($rows.ToArray() | Group-Object BatchId)){
        if($group.Count -ne 2){continue}
        $copies=@($group.Group)
        $left=@($copies[0].Files | Where-Object Name -CNotIn $copies[0].TechnicalLockNames | Sort-Object Name | Select-Object Name,Bytes,SHA256)
        $right=@($copies[1].Files | Where-Object Name -CNotIn $copies[1].TechnicalLockNames | Sort-Object Name | Select-Object Name,Bytes,SHA256)
        if(-not $copies[0].Valid -or -not $copies[1].Valid -or (Get-RetentionObjectHash $left) -cne (Get-RetentionObjectHash $right)){
            foreach($row in $copies){if($row.State -cin @('Eligible','Candidate')){$row.State='Blocked';$row.Reason='Local/SharePoint immutable batch contents disagree.'}}
            if($protected.ContainsKey($group.Name)){$blocked.Add("Protected batch copies disagree: $($group.Name)")}
        }
    }
    # At most the configured number of SharePoint folders per publication, oldest first.
    $cloudActions=@($rows.ToArray() | Where-Object {$_.Storage -ceq 'SharePoint' -and $_.State -cin @('Eligible','Candidate')} | Sort-Object BatchId)
    foreach($row in @($cloudActions | Select-Object -Skip $policy.MaxSharePointRemovalsPerRun)){
        $row.State='Deferred';$row.Reason="Per-run SharePoint removal limit ($($policy.MaxSharePointRemovalsPerRun)) reached; a later publication retries it."
        foreach($localRow in @($rows.ToArray() | Where-Object {$_.Storage -ceq 'Local' -and $_.BatchId -ceq $row.BatchId -and $_.State -cin @('Eligible','Candidate')})){
            $localRow.State='Deferred';$localRow.Reason='Server copy waits for its deferred SharePoint retirement.'
        }
    }
    [pscustomobject]@{SchemaVersion=3;RetentionPolicy=$policy;Verification=$Snapshot.Verification;ApplySupported=($Snapshot.Verification -ceq 'Full');PolicyVersion=$script:RetentionPolicyVersion;PlanId=[guid]::NewGuid().ToString('N');CreatedUtc=[datetime]::UtcNow.ToString('O');TenantKey=$TenantKey;OutputRoot=$OutputRoot;RegistrySHA256=$RegistryHash;ConfigurationSHA256=$ConfigurationHash;ModuleSHA256=$script:RetentionModuleSHA256;SnapshotSHA256=(Get-RetentionObjectHash (Get-RetentionSnapshotEvidence $Snapshot));Target=$Snapshot.Target;LocalPointer=$Snapshot.LocalPointer;RemotePointer=$Snapshot.RemotePointer;ProtectedBatchIds=@($protected.Keys | Sort-Object);MinimumRetainedBatches=$minimum;BlockedReasons=$blocked.ToArray();Rows=$rows.ToArray()}
}
function Assert-RetentionPlan($Plan,$Fresh,[int]$MaxAgeHours) {
    foreach($candidate in $Plan,$Fresh){if($candidate.SchemaVersion -ne 3 -or $candidate.Verification -cne 'Full' -or $candidate.ApplySupported -isnot [bool] -or -not $candidate.ApplySupported){throw 'Application requires a current Full verification plan; metadata previews cannot authorize removal.'}}
    if($MaxAgeHours -lt 1 -or $MaxAgeHours -gt 168){throw 'Plan maximum age must be 1..168 hours.'}
    $created=if($Plan.CreatedUtc -is [datetime]){$Plan.CreatedUtc.ToUniversalTime()}else{[DateTimeOffset]::Parse([string]$Plan.CreatedUtc).UtcDateTime}
    $age=[datetime]::UtcNow-$created
    if($age.TotalHours -gt $MaxAgeHours -or $age.TotalSeconds -lt 0){throw 'Retention plan is stale.'}
    foreach($key in 'SchemaVersion','Verification','ApplySupported','PolicyVersion','TenantKey','OutputRoot','RegistrySHA256','ConfigurationSHA256','ModuleSHA256','SnapshotSHA256'){
        if($Plan.$key -cne $Fresh.$key){throw "Retention plan changed/stale: $key"}
    }
    if((Get-RetentionObjectHash $Plan.Rows) -cne (Get-RetentionObjectHash $Fresh.Rows) -or
       (Get-RetentionObjectHash $Plan.ProtectedBatchIds) -cne (Get-RetentionObjectHash $Fresh.ProtectedBatchIds) -or
       (Get-RetentionObjectHash (Get-RetentionValue $Plan 'MinimumRetainedBatches' @())) -cne (Get-RetentionObjectHash $Fresh.MinimumRetainedBatches)){throw 'Retention plan actions/protections differ from fresh validation.'}
    if(@($Plan.BlockedReasons).Count -or @($Fresh.BlockedReasons).Count){throw 'Retention application blocked by qualification.'}
}
function Save-RetentionReceipts($Row,[string]$OutputRoot,$Transport,[string]$Scratch) {
    $archive=Get-RetentionChild $OutputRoot ('retired/'+$Row.BatchId)
    New-Item -ItemType Directory -Path $archive -Force | Out-Null
    foreach($name in 'current.json.txt','batch.json.txt','validation.json.txt'){
        $source=@($Row.Files | Where-Object Name -CEQ $name)[0];$path=Get-RetentionChild $archive $name
        if(Test-Path -LiteralPath $path){if((Get-FileHash -LiteralPath $path).Hash -cne $source.SHA256){throw 'Archived receipt conflicts with source.'}}
        elseif($Row.Storage -ceq 'Local'){[IO.File]::Copy($source.Id,$path,$false)}
        elseif((& $Transport.Download $source.Id $path) -ne $true){throw 'Receipt archive download failed.'}
        if((Get-Item -LiteralPath $path).Length -ne $source.Bytes -or (Get-FileHash -LiteralPath $path).Hash -cne $source.SHA256){throw 'Local receipt archive verification failed.'}
        $relative='retired/'+$Row.BatchId+'/'+$name
        if((& $Transport.Archive $path $relative) -ne $true){throw 'Remote receipt archive/read-back failed.'}
    }
}
function Get-RetentionJournalFolder([string[]]$JournalRoots) {
    # The shared root is authoritative; the legacy per-account work root is still read so old incidents keep blocking.
    foreach($root in @($JournalRoots | Select-Object -Unique)){
        $history=Get-RetentionChild $root 'retention'
        if(Test-Path -LiteralPath $history){Get-ChildItem -LiteralPath $history -Directory -Force | Sort-Object Name}
    }
}
function Test-RetentionPriorRemoteConfirmation($Row,[string[]]$JournalRoots,[string]$RegistryHash,[string]$ConfigurationHash) {
    foreach($dir in @(Get-RetentionJournalFolder $JournalRoots)){
        $path=Get-RetentionChild $dir.FullName 'journal.json.txt'
        if(-not (Test-Path -LiteralPath $path)){continue}
        try {
            $journal=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -Depth 70
            if(-not @($journal.Events | Where-Object {$_.Storage -ceq 'SharePoint' -and $_.BatchId -ceq $Row.BatchId -and $_.State -ceq 'Recycled'}).Count){continue}
            $originalPath=[IO.Path]::GetFullPath($journal.PlanPath);Assert-RetentionPath $originalPath
            if((Get-FileHash -LiteralPath $originalPath).Hash -cne $journal.PlanSHA256){continue}
            $original=Get-Content -LiteralPath $originalPath -Raw | ConvertFrom-Json -Depth 70
            if($original.RegistrySHA256 -cne $RegistryHash -or $original.ConfigurationSHA256 -cne $ConfigurationHash -or $original.ModuleSHA256 -cne $script:RetentionModuleSHA256){continue}
            $prior=@($original.Rows | Where-Object {$_.Storage -ceq 'Local' -and $_.BatchId -ceq $Row.BatchId -and $_.State -ceq 'Eligible'})
            if($prior.Count -eq 1 -and $prior[0].Id -ceq $Row.Id -and (Get-RetentionObjectHash $prior[0].Files) -ceq (Get-RetentionObjectHash $Row.Files)){return $true}
        } catch {Write-Verbose 'An incomplete or changed prior journal cannot authorize local recovery.'}
    }
    $false
}
function Get-RetentionUnresolvedIncidents([string[]]$JournalRoots) {
    foreach($dir in @(Get-RetentionJournalFolder $JournalRoots)){
        $path=Get-RetentionChild $dir.FullName 'journal.json.txt'
        if(-not(Test-Path -LiteralPath $path)){continue}
        $prior=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -Depth 70
        if($prior.SchemaVersion -ne 1){throw "Unknown retention journal schema requires reconciliation: $path"}
        foreach($entry in @($prior.Events)){
            if($entry.State -cin @('Uncertain','RemovalRequested')){[string]$entry.BatchId}
        }
    }
}
function Assert-RetentionBeforeRemoval($Row,$Snapshot,$Settings,$Transport,[string]$RegistryHash,[string]$Root,[string]$Scratch) {
    $target=& $Transport.Describe
    Assert-RetentionTarget $target $Snapshot.Target
    Assert-RetentionQualification $Settings $target $RegistryHash $Root
    if((Get-FileHash -LiteralPath (Join-Path $Root 'current.json.txt')).Hash -cne $Snapshot.LocalPointerSHA256){throw 'Local pointer changed before removal.'}
    $p=@(Get-RetentionRemoteChildren $Transport $target.RootId | Where-Object Name -CEQ 'current.json.txt')
    if($p.Count -ne 1 -or (Read-RetentionRemoteFile $Transport $p[0] $Scratch).SHA256 -cne $Snapshot.RemotePointerSHA256){throw 'Remote pointer changed before removal.'}
    Assert-RetentionBatchUnchanged $Row $target $Transport $Root $Scratch
}
function Assert-RetentionBatchUnchanged($Row,$Target,$Transport,[string]$Root,[string]$Scratch) {
    if($Row.Storage -ceq 'SharePoint'){
        $folder=@(Get-RetentionRemoteChildren $Transport $Target.BatchesId | Where-Object Id -CEQ $Row.Id)
        if($folder.Count -ne 1 -or $folder[0].Name -cne $Row.BatchId -or $folder[0].Kind -cne 'Folder' -or $folder[0].ETag -cne $Row.ETag){throw 'Remote batch changed after archival.'}
        $files=@(Get-RetentionRemoteChildren $Transport $Row.Id)
        if($files.Count -ne $Row.FileCount){throw 'Remote content changed after archival.'}
        foreach($file in $files){
            $expected=@($Row.Files | Where-Object Name -CEQ $file.Name)
            if($file.Kind -cne 'File' -or $expected.Count -ne 1 -or $file.Id -cne $expected[0].Id -or $file.ETag -cne $expected[0].ETag -or
               (Read-RetentionRemoteFile $Transport $file $Scratch).SHA256 -cne $expected[0].SHA256){throw 'Remote payload changed after archival.'}
        }
    } else {
        $path=Get-RetentionChild $Root ('batches/'+$Row.BatchId)
        if($path -cne $Row.Id){throw 'Local batch identity changed after archival.'}
        $files=@(Get-ChildItem -LiteralPath $path -Force)
        if($files.Count -ne $Row.FileCount){throw 'Local content changed after archival.'}
        foreach($file in $files){
            Assert-RetentionPath $file.FullName;$expected=@($Row.Files | Where-Object Name -CEQ $file.Name)
            if($file.PSIsContainer -or $expected.Count -ne 1 -or $file.Length -ne $expected[0].Bytes -or (Get-FileHash -LiteralPath $file.FullName).Hash -cne $expected[0].SHA256){throw 'Local payload changed after archival.'}
        }
    }
}
function Assert-RetentionMinimumBeforeRemoval($Row,$Snapshot,$Minimum,$Transport,[string]$Root,[string]$Scratch) {
    if($Snapshot.Verification -cne 'Full'){throw 'Retention minimum requires full payload verification.'}
    foreach($storage in 'Local','SharePoint'){
        $sets=@($Minimum | Where-Object Storage -CEQ $storage)
        if($sets.Count -ne 1 -or @($sets[0].BatchIds | Select-Object -Unique).Count -lt 2){throw "Retention minimum missing for $storage."}
        # SharePoint selections bind both copies; a server-only selection binds only the server copy.
        if($Row.BatchId -cin $sets[0].BatchIds -and ($storage -ceq 'SharePoint' -or $Row.Storage -ceq $storage)){throw "A newest retained $storage batch cannot be retired: $($Row.BatchId)."}
        foreach($id in $sets[0].BatchIds){
            $copies=@($Snapshot.Batches | Where-Object {$_.Storage -ceq $storage -and $_.BatchId -ceq $id})
            if($copies.Count -ne 1){throw "Retention minimum copy is ambiguous or absent: $storage/$id"}
            $copy=$copies[0]
            $retained=[pscustomobject]@{Storage=$storage;BatchId=$id;Id=$copy.Id;ETag=$copy.ETag;Files=$copy.Files;FileCount=$copy.Files.Count}
            try {Assert-RetentionBatchUnchanged $retained $Snapshot.Target $Transport $Root $Scratch}
            catch {throw "Retention minimum changed or unavailable: $storage/$id. $($_.Exception.Message)"}
        }
    }
}
function Assert-RetentionOutsideData([string]$Path,[string]$Data,[string]$Label) {
    if($Path -eq $Data -or $Path.StartsWith($Data+'\',[StringComparison]::OrdinalIgnoreCase) -or $Data.StartsWith($Path+'\',[StringComparison]::OrdinalIgnoreCase)){throw "$Label must be outside DATA and its ancestors."}
}
function Resolve-RetentionJournalRoot($Settings,[string]$Data,[string]$WorkRoot,[bool]$Apply) {
    $configured=[string](Get-RetentionValue $Settings 'JournalRootPath' '')
    if([string]::IsNullOrWhiteSpace($configured)){
        if($Apply){throw 'PreparedRetention.JournalRootPath is required for Apply and Auto: set an absolute shared folder outside DATA so every server and account writes and reads the same retention journals.'}
        return $WorkRoot
    }
    if(-not [IO.Path]::IsPathFullyQualified($configured)){throw "PreparedRetention.JournalRootPath must be an absolute path: '$configured'."}
    $path=[IO.Path]::GetFullPath($configured).TrimEnd('\','/')
    Assert-RetentionOutsideData $path $Data 'PreparedRetention.JournalRootPath'
    Assert-RetentionPath $path
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    Assert-RetentionPath $path
    $path
}
function Get-RetentionRemovalObservation($Row,$Target,$Transport) {
    # Read-only observation after an unconfirmed DELETE; any read failure keeps the outcome unknown.
    try {
        if((& $Transport.IsAbsent $Row.Id) -eq $true){return 'Absent'}
        $items=@(Get-RetentionRemoteChildren $Transport $Target.BatchesId | Where-Object Id -CEQ $Row.Id)
        if($items.Count -eq 1 -and $items[0].Name -ceq $Row.BatchId -and $items[0].ETag -ceq $Row.ETag){return 'Unchanged'}
    } catch {Write-Verbose "Removal outcome observation failed: $($_.Exception.Message)"}
    'Unknown'
}
function Invoke-RetentionSharePointRemoval($Row,$JournalEvent,$Transport,$Target,$Request) {
    $result=& $Transport.Recycle $Request
    $outcome=[string](Get-RetentionValue $result 'Outcome' '')
    $JournalEvent.HttpStatus=Get-RetentionValue $result 'HttpStatus'
    $detail=[string](Get-RetentionValue $result 'Error' '')
    if($outcome -ceq 'Recycled'){
        if((& $Transport.IsAbsent $Row.Id) -ne $true){throw "SharePoint batch $($Row.BatchId): Graph returned HTTP 204 but the folder is still visible; outcome uncertain. Check DATA-POWERBI/batches and the site recycle bin, then reconcile this journal."}
        $JournalEvent.Outcome='Recycled';$JournalEvent.State='Recycled';return
    }
    $observed=Get-RetentionRemovalObservation $Row $Target $Transport
    $what=switch($outcome){
        'NotFound' {'Graph returned HTTP 404'}
        'PreconditionFailed' {'Graph returned HTTP 412: the folder eTag changed after review'}
        default {if($detail){$detail}else{"unexpected transport outcome '$outcome'"}}
    }
    if($observed -ceq 'Absent'){
        $JournalEvent.Outcome=if($outcome -ceq 'NotFound'){'AlreadyAbsent'}else{'AbsentAfterError'}
        if($JournalEvent.Outcome -ceq 'AbsentAfterError'){Write-Warning "SharePoint batch $($Row.BatchId) is absent after an unconfirmed removal ($what); recorded as recycled."}
        $JournalEvent.State='Recycled';return
    }
    # A folder still listed with its reviewed eTag was not removed, unless the request may still be in flight.
    if($observed -ceq 'Unchanged' -and (Get-RetentionValue $result 'MayHaveReachedService' $true) -ne $true){
        $JournalEvent.Outcome='NotRemoved';$JournalEvent.State='NotRemoved'
        throw "SharePoint batch $($Row.BatchId) was not removed ($what). The folder is still present with its reviewed eTag; nothing was deleted and a later publication plans it again."
    }
    throw "SharePoint batch $($Row.BatchId) removal outcome is uncertain ($what; folder observation: $observed). Check DATA-POWERBI/batches and the site recycle bin, then reconcile this journal before Auto continues."
}
function Invoke-PreparedBatchRetention {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputRoot,[Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][string]$ConfigurationPath,[Parameter(Mandatory)][string]$ConsumersPath,
        [Parameter(Mandatory)][string]$WorkRoot,[Parameter(Mandatory)]$Transport,
        [ValidateSet('Configured','Preview','Apply')][string]$Mode='Preview',
        [ValidateSet('Metadata','Full')][string]$Verification='Metadata',[string]$PlanPath,[string]$PlanSHA256,
        [string]$ExpectedPublishedBatchId,[string]$ExpectedPointerSHA256)
    $root=[IO.Path]::GetFullPath($OutputRoot).TrimEnd('\','/');Assert-RetentionPath $root
    if((Split-Path $root -Leaf) -cne 'DATA-POWERBI'){throw 'Retention requires the exact server DATA-POWERBI root.'}
    $work=[IO.Path]::GetFullPath($WorkRoot).TrimEnd('\','/');Assert-RetentionPath $work
    $data=Split-Path $root -Parent
    Assert-RetentionOutsideData $work $data 'Retention work'
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $lock=$null;$config=$null;$registry=$null;$planLock=$null;$run=$null;$journal=$null;$journalPath=$null;$scratch=$null
    try {
        $lock=[IO.File]::Open((Join-Path $root '.publication.lock'),'OpenOrCreate','ReadWrite','None')
        Assert-RetentionLock $lock
        $config=Read-RetentionLockedDocument $ConfigurationPath
        $settings=$config.Value.PreparedRetention
        if($settings.Mode -cnotin @('Preview','Auto')){throw 'PreparedRetention.Mode must be Preview or Auto.'}
        if($Mode -ceq 'Apply' -and (-not $PlanPath -or $PlanSHA256 -cnotmatch '^[A-Fa-f0-9]{64}$')){throw 'Manual Apply requires PlanPath and PlanSHA256.'}
        $apply=$Mode -ceq 'Apply' -or ($Mode -ceq 'Configured' -and $settings.Mode -ceq 'Auto')
        # A caller cannot weaken application checks with the Preview setting.
        $effectiveVerification=if($apply){'Full'}else{$Verification}
        if($Mode -ceq 'Apply'){
            $planLock=Read-RetentionLockedDocument $PlanPath
            if($planLock.SHA256 -ine $PlanSHA256){throw 'Approved retention plan hash mismatch.'}
            $approved=$planLock.Value
            if((Get-RetentionValue $approved 'SchemaVersion') -ne 3 -or (Get-RetentionValue $approved 'Verification') -cne 'Full' -or
               (Get-RetentionValue $approved 'ApplySupported') -isnot [bool] -or -not $approved.ApplySupported){throw 'Application requires a Full verification plan; metadata previews cannot authorize removal.'}
        }
        # Plans and journals live under the shared journal root; read-back temporaries stay in local work.
        $journalRoot=Resolve-RetentionJournalRoot $settings $data $work $apply
        $journalRoots=@($journalRoot,$work)
        $runId=[guid]::NewGuid().ToString('N')
        $run=Get-RetentionChild $journalRoot ('retention/'+$runId)
        New-Item -ItemType Directory -Path $run -Force | Out-Null
        $scratch=Get-RetentionChild $work ('retention-scratch/'+$runId)
        New-Item -ItemType Directory -Path $scratch -Force | Out-Null
        $journalPath=Join-Path $run 'journal.json.txt'
        $journal=[ordered]@{SchemaVersion=1;Status='Planning';StartedUtc=[datetime]::UtcNow.ToString('O');Machine=[Environment]::MachineName;ModuleSHA256=$script:RetentionModuleSHA256;PlanPath='';PlanSHA256='';Events=[Collections.Generic.List[object]]::new();Error=$null}
        Write-RetentionDocument $journalPath $journal
        # A missing registry is never synthesized from the neutral template.
        try {$registry=Read-RetentionLockedDocument $ConsumersPath;$registryValue=$registry.Value;$registryHash=$registry.SHA256}
        catch {$registryValue=[pscustomobject]@{};$registryHash='MissingOrInvalid'}
        $snapshot=Get-RetentionSnapshot $root $Transport $scratch $TenantKey -Verification $effectiveVerification
        $fresh=New-RetentionPlan $snapshot $registryValue $registryHash $config.SHA256 $root $settings $TenantKey
        foreach($incident in @(Get-RetentionUnresolvedIncidents $journalRoots)){
            foreach($row in @($fresh.Rows | Where-Object {$_.BatchId -ceq $incident -and $_.State -cin @('Eligible','Candidate')})){
                $row.State='Blocked';$row.Reason='Unresolved prior removal intent/outcome; operator reconciliation required.'
            }
        }
        if(($ExpectedPublishedBatchId -and $snapshot.LocalPointer.BatchId -cne $ExpectedPublishedBatchId) -or
           ($ExpectedPointerSHA256 -and $snapshot.RemotePointerSHA256 -ine $ExpectedPointerSHA256)){$fresh.BlockedReasons+= 'Published transfer receipt no longer matches current pointers.'}
        $newPath=Join-Path $run 'plan.json.txt';Write-RetentionDocument $newPath $fresh
        $newHash=(Get-FileHash -LiteralPath $newPath).Hash
        $plan=$fresh
        if($Mode -ceq 'Apply'){
            $plan=$planLock.Value
        } else {$PlanPath=$newPath;$PlanSHA256=$newHash}
        $journal.PlanPath=$PlanPath;$journal.PlanSHA256=$PlanSHA256
        $journal.Status=if($fresh.BlockedReasons.Count){'Blocked'}else{'Preview'}
        Write-RetentionDocument $journalPath $journal
        if(-not $apply){return [pscustomobject]@{Status=$journal.Status;PlanPath=$newPath;PlanSHA256=$newHash;JournalPath=$journalPath;Plan=$fresh}}
        Assert-RetentionPlan $plan $fresh ([int]$settings.PlanMaxAgeHours)
        $journal.Status='Applying';Write-RetentionDocument $journalPath $journal
        # Fresh state was read under the common lock. Recheck cloud target/access/pointers before each action.
        foreach($row in @($plan.Rows | Where-Object State -CEQ 'Eligible' | Sort-Object @{Expression={if($_.Storage -ceq 'SharePoint'){0}else{1}}},BatchId)){
            $event=[ordered]@{Storage=$row.Storage;BatchId=$row.BatchId;Id=$row.Id;State='Checking';Utc=[datetime]::UtcNow.ToString('O');RequestUri=$null;HttpStatus=$null;Outcome=$null;Error=$null}
            $journal.Events.Add($event);Write-RetentionDocument $journalPath $journal
            $technicalHandles=[Collections.Generic.List[IO.FileStream]]::new()
            try {
                Assert-RetentionLock $lock
                $liveTarget=& $Transport.Describe;Assert-RetentionTarget $liveTarget $plan.Target
                Assert-RetentionQualification $settings $liveTarget $registryHash $root
                if((Get-FileHash -LiteralPath (Join-Path $root 'current.json.txt')).Hash -cne $snapshot.LocalPointerSHA256){throw 'Local pointer changed during application.'}
                $pointerItems=@(Get-RetentionRemoteChildren $Transport $liveTarget.RootId | Where-Object Name -CEQ 'current.json.txt')
                if($pointerItems.Count -ne 1 -or (Read-RetentionRemoteFile $Transport $pointerItems[0] $scratch).SHA256 -cne $snapshot.RemotePointerSHA256){throw 'Remote pointer changed during application.'}
                if($row.BatchId -cin @($plan.ProtectedBatchIds)){throw 'Protected batch cannot be retired.'}
                # Also guard local sidecars before cloud retirement, retaining handles through removal.
                foreach($localRow in @($plan.Rows | Where-Object {$_.Storage -ceq 'Local' -and $_.BatchId -ceq $row.BatchId})){
                    foreach($name in $localRow.TechnicalLockNames){
                        $file=@($localRow.Files | Where-Object Name -CEQ $name)
                        if($file.Count -ne 1){throw 'Technical lock missing from approved inventory.'}
                        $technicalHandles.Add((Open-RetentionTechnicalLock $file[0]))
                    }
                }
                if($row.Storage -ceq 'Local'){
                    $remoteRow=@($plan.Rows | Where-Object { $_.Storage -ceq 'SharePoint' -and $_.BatchId -ceq $row.BatchId })
                    $confirmed=@($journal.Events | Where-Object {$_.Storage -ceq 'SharePoint' -and $_.BatchId -ceq $row.BatchId -and $_.State -ceq 'Recycled'})
                    $priorConfirmed=$false
                    if($remoteRow.Count -eq 0){$priorConfirmed=Test-RetentionPriorRemoteConfirmation $row $journalRoots $registryHash $config.SHA256}
                    if(-not ($remoteRow.Count -eq 1 -and $confirmed.Count -eq 1) -and -not $priorConfirmed){throw 'Local retirement requires a fingerprint-matched confirmed remote retirement journal.'}
                    $path=Get-RetentionChild $root ('batches/'+$row.BatchId)
                    if($path -cne $row.Id){throw 'Local batch identity changed.'}
                    $actual=@(Get-ChildItem -LiteralPath $path -Force)
                    if($actual.Count -ne $row.FileCount){throw 'Local batch content changed.'}
                    foreach($file in $actual){
                        Assert-RetentionPath $file.FullName;$expected=@($row.Files | Where-Object Name -CEQ $file.Name)
                        if($file.PSIsContainer -or $expected.Count -ne 1 -or $file.Length -ne $expected[0].Bytes -or (Get-FileHash -LiteralPath $file.FullName).Hash -cne $expected[0].SHA256){throw 'Local batch changed before retirement.'}
                    }
                } else {
                    $parentItems=@(Get-RetentionRemoteChildren $Transport $liveTarget.BatchesId | Where-Object Id -CEQ $row.Id)
                    if($parentItems.Count -ne 1 -or $parentItems[0].Name -cne $row.BatchId -or $parentItems[0].ETag -cne $row.ETag){throw 'Remote batch identity changed.'}
                    $actual=@(Get-RetentionRemoteChildren $Transport $row.Id)
                    if($actual.Count -ne $row.FileCount){throw 'Remote batch content changed.'}
                    foreach($file in $actual){
                        $expected=@($row.Files | Where-Object Name -CEQ $file.Name)
                        if($file.Kind -cne 'File' -or $expected.Count -ne 1 -or $file.Id -cne $expected[0].Id -or $file.ETag -cne $expected[0].ETag -or
                           (Read-RetentionRemoteFile $Transport $file $scratch).SHA256 -cne $expected[0].SHA256){throw 'Remote payload changed before retirement.'}
                    }
                }
                Save-RetentionReceipts $row $root $Transport $scratch
                $event.State='Archived';Write-RetentionDocument $journalPath $journal
                Assert-RetentionBeforeRemoval $row $snapshot $settings $Transport $registryHash $root $scratch
                # Recheck all retained payloads on BOTH stores, including during partial-run recovery.
                Assert-RetentionMinimumBeforeRemoval $row $snapshot $fresh.MinimumRetainedBatches $Transport $root $scratch
                Assert-RetentionLock $lock
                $removalRequest=$null
                if($row.Storage -ceq 'SharePoint'){
                    # Build and validate the exact request before recording intent: a refused URI is Blocked, never Uncertain.
                    $prepare=Get-RetentionValue $Transport 'PrepareRecycle'
                    if($prepare -isnot [scriptblock]){throw 'Retention transport lacks PrepareRecycle; SharePoint removal refused before any request.'}
                    $removalRequest=& $prepare $row.Id $row.ETag
                    if($null -eq $removalRequest -or [string]::IsNullOrWhiteSpace([string]$removalRequest.Uri)){throw "SharePoint batch $($row.BatchId): the transport returned no validated removal URI."}
                    $event.RequestUri=[string]$removalRequest.Uri
                }
                # Persist intent BEFORE the irreversible request; uncertain results are never blindly replayed.
                $event.State='RemovalRequested';Write-RetentionDocument $journalPath $journal
                if($row.Storage -ceq 'SharePoint'){
                    Invoke-RetentionSharePointRemoval $row $event $Transport $liveTarget $removalRequest
                } else {
                    # All paths were resolved and verified against the approved inventory above.
                    Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
                    $event.State='Retired'
                }
                Write-RetentionDocument $journalPath $journal
            } catch {
                $event.Error=$_.Exception.Message
                $event.State=switch -CaseSensitive ($event.State){'RemovalRequested' {'Uncertain'};'NotRemoved' {'NotRemoved'};default {'Blocked'}}
                Write-RetentionDocument $journalPath $journal
                throw
            } finally {foreach($handle in $technicalHandles){$handle.Dispose()}}
        }
        $journal.Status='Completed';Write-RetentionDocument $journalPath $journal
        [pscustomobject]@{Status='Completed';PlanPath=$PlanPath;PlanSHA256=$PlanSHA256;JournalPath=$journalPath;Plan=$plan}
    } catch {
        if($journal){$journal.Error=$_.Exception.Message;$journal.Status='PartialOrBlocked';Write-RetentionDocument $journalPath $journal}
        throw
    } finally {
        if($planLock){$planLock.Stream.Dispose()};if($registry){$registry.Stream.Dispose()};if($config){$config.Stream.Dispose()};if($lock){$lock.Dispose()}
        # Only this run's own read-back folder is removed; journals and plans are kept.
        if($scratch -and (Test-Path -LiteralPath $scratch)){Assert-RetentionPath $scratch;Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue}
    }
}

function Assert-RetentionGraphDeleteSupport {
    $command=Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue
    if(-not $command){throw 'Invoke-MgGraphRequest is unavailable; SharePoint batch removal refused before any request.'}
    foreach($name in 'Method','Uri','Headers','OutputType','SkipHttpErrorCheck','StatusCodeVariable'){
        if(-not $command.Parameters.ContainsKey($name)){
            $source=if($command.Version){"$($command.Source) $($command.Version)"}else{[string]$command.Source}
            throw "Invoke-MgGraphRequest ($source) lacks -$name; update Microsoft.Graph.Authentication. SharePoint batch removal refused before any request."
        }
    }
}
function Invoke-RetentionGraphDelete([string]$Uri,[string]$ETag) {
    Assert-RetentionGraphDeleteSupport
    # Never Set-MgRequestContext here: SDK 2.30 then rebuilds its client from a session environment that an
    # earlier absolute-URI request left without an authority, and the certificate credential fails with Invalid URI.
    # The status comes from -StatusCodeVariable, never from a disposable HttpResponseMessage.
    $deleteStatus=$null
    try {
        $null=Invoke-MgGraphRequest -Method DELETE -Uri $Uri -Headers @{'If-Match'=$ETag} -OutputType PSObject -SkipHttpErrorCheck -StatusCodeVariable 'deleteStatus' -ErrorAction Stop
    } catch {
        $types=[Collections.Generic.List[string]]::new();$exception=$_.Exception
        for($depth=0;$depth -lt 8 -and $exception;$depth++){$types.Add($exception.GetType().FullName);$exception=$exception.InnerException}
        # Connection, timeout and cancellation failures may follow an accepted request; their outcome stays uncertain.
        $inFlight=@($types | Where-Object {$_ -in @('System.Net.Http.HttpRequestException','System.Threading.Tasks.TaskCanceledException','System.OperationCanceledException','System.TimeoutException','System.IO.IOException','System.Net.Sockets.SocketException','System.Net.WebException')}).Count -gt 0
        $code=if($null -ne $deleteStatus){[int]$deleteStatus}else{$null}
        return [pscustomobject]@{Outcome='RequestFailed';HttpStatus=$code;MayHaveReachedService=$inFlight;Error=('Graph DELETE failed with {0}: {1}' -f ($types -join ' <- '),$_.Exception.Message)}
    }
    if($null -eq $deleteStatus){return [pscustomobject]@{Outcome='HttpError';HttpStatus=$null;MayHaveReachedService=$true;Error='Graph DELETE completed without a status code.'}}
    $code=[int]$deleteStatus
    $outcome=switch($code){204 {'Recycled'};404 {'NotFound'};412 {'PreconditionFailed'};default {'HttpError'}}
    # A refusal (4xx, or 503 throttling) changed nothing; other 2xx (202 = asynchronous) and 5xx may have acted.
    $mayHaveActed=$outcome -ceq 'HttpError' -and -not (($code -ge 400 -and $code -le 499) -or $code -eq 503)
    [pscustomobject]@{Outcome=$outcome;HttpStatus=$code;MayHaveReachedService=$mayHaveActed;Error=$(if($outcome -ceq 'HttpError'){"Graph DELETE returned HTTP $code; only 204 confirms removal."}else{$null})}
}
function New-PreparedRetentionGraphTransport {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Connection,[Parameter(Mandatory)][string]$CloudRoot,
        [Parameter(Mandatory)][string]$OutputRoot,[Parameter(Mandatory)][string]$WorkRoot,
        [Parameter(Mandatory)][scriptblock]$UploadFile,[Parameter(Mandatory)][scriptblock]$DownloadFile)
    # GetNewClosure creates a dynamic module; retain bound private helpers explicitly.
    $readValue=${function:Get-RetentionValue}
    $objectHash=${function:Get-RetentionObjectHash}
    $assertPath=${function:Assert-RetentionPath}
    $httpStatus=${function:Get-RetentionHttpStatus}
    $assertDeleteSupport=${function:Assert-RetentionGraphDeleteSupport}
    $graphDelete=${function:Invoke-RetentionGraphDelete}
    # Reuse the public app-only connector. No private Core functions or cached tenant identity.
    if(-not (Connect-SmartM365GraphAppOnly -AppId $Connection.AppId -TenantId $Connection.TenantId -Thumbprint $Connection.Thumbprint -Purpose 'Prepared batch retention')){throw 'Retention Graph connection failed.'}
    $context=Get-MgContext
    if($context.TenantId -ine $Connection.TenantId -or $context.ClientId -ine $Connection.AppId -or $context.AuthType -cne 'AppOnly'){throw 'Retention requires the expected app-only Graph context.'}
    if($Connection.SiteHostname -notmatch '^[A-Za-z0-9.-]+$' -or $Connection.SitePath -notmatch '^/(sites|teams)/[A-Za-z0-9_-]+$'){throw 'Invalid configured SharePoint site.'}
    if($CloudRoot -match '[:%?#\\]' -or @($CloudRoot.Split('/') | Where-Object {$_ -in @('','.','..')}).Count -or $CloudRoot.Split('/')[-1] -cne 'DATA-POWERBI'){throw 'Invalid retention cloud root.'}
    $request={param([string]$Method,[string]$Uri,$Headers)
        $ctx=Get-MgContext
        if($ctx.TenantId -ine $Connection.TenantId -or $ctx.ClientId -ine $Connection.AppId -or $ctx.AuthType -cne 'AppOnly'){throw 'Graph context changed during retention.'}
        if(-not $Uri.StartsWith('https://graph.microsoft.com/v1.0/',[StringComparison]::Ordinal)){throw 'Unexpected Graph endpoint or pagination link.'}
        # The only mutation is the validated recycle request below.
        if($Method -cne 'GET'){throw "Retention reads use GET only; $Method refused."}
        $arguments=@{Method=$Method;Uri=$Uri;OutputType='PSObject';ErrorAction='Stop'}
        if($Headers){$arguments.Headers=$Headers}
        Invoke-MgGraphRequest @arguments
    }.GetNewClosure()
    $pages={param([string]$Uri)
        $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        while($Uri){
            if(-not $seen.Add($Uri)){throw 'Cyclic Graph pagination.'}
            $page=& $request 'GET' $Uri $null
            if(-not $page.PSObject.Properties['value']){throw 'Incomplete Graph enumeration.'}
            foreach($entry in @($page.value)){$entry}
            $Uri=[string](& $readValue $page '@odata.nextLink' '')
        }
    }.GetNewClosure()
    $site=& $request 'GET' ('https://graph.microsoft.com/v1.0/sites/'+$Connection.SiteHostname+':'+$Connection.SitePath) $null
    $drives=@(& $pages ('https://graph.microsoft.com/v1.0/sites/'+$site.id+'/drives') | Where-Object name -CEQ $Connection.LibraryDisplayName)
    if($drives.Count -ne 1){throw 'Retention requires one exact document library.'}
    $drive=$drives[0];$driveBase='https://graph.microsoft.com/v1.0/drives/'+[uri]::EscapeDataString($drive.id)
    $encoded=($CloudRoot.Split('/') | ForEach-Object {[uri]::EscapeDataString($_)}) -join '/'
    $rootItem=& $request 'GET' ($driveBase+'/root:/'+$encoded) $null
    $batchesItem=& $request 'GET' ($driveBase+'/items/'+[uri]::EscapeDataString($rootItem.id)+':/batches') $null
    foreach($item in $rootItem,$batchesItem){if(-not (Get-RetentionValue $item 'folder') -or (Get-RetentionValue $item 'remoteItem')){throw 'Retention root is not a native SharePoint folder.'}}
    if($batchesItem.parentReference.id -cne $rootItem.id -or $batchesItem.parentReference.driveId -cne $drive.id){throw 'Remote batches root escaped the configured authority.'}
    $describe={
        $nativeRoot=& $request 'GET' ($driveBase+'/root:/'+$encoded) $null
        $nativeBatches=& $request 'GET' ($driveBase+'/items/'+[uri]::EscapeDataString($nativeRoot.id)+':/batches') $null
        if($nativeRoot.id -cne $rootItem.id -or $nativeBatches.id -cne $batchesItem.id -or
           $nativeBatches.parentReference.id -cne $rootItem.id -or (& $readValue $nativeRoot 'remoteItem') -or (& $readValue $nativeBatches 'remoteItem')){throw 'Retention root identity changed.'}
        # Bind approval to live ACL and permission evidence, not a switch claiming exclusivity.
        # A site owner must additionally review inherited groups, other app grants and sync writers.
        $permissions=@(& $pages ($driveBase+'/items/'+[uri]::EscapeDataString($rootItem.id)+'/permissions') | Sort-Object id)
        $batchPermissions=@(& $pages ($driveBase+'/items/'+[uri]::EscapeDataString($batchesItem.id)+'/permissions') | Sort-Object id)
        $acl=(Get-Acl -LiteralPath $OutputRoot).Sddl
        if(-not $acl){throw 'Server writer access evidence cannot be read; retention is blocked.'}
        # A successfully enumerated empty sharing list is valid evidence, not proof of no inherited writers.
        [pscustomobject]@{TenantId=[string]$context.TenantId;ClientId=[string]$context.ClientId;SiteId=[string]$site.id;DriveId=[string]$drive.id;RootId=[string]$rootItem.id;BatchesId=[string]$batchesItem.id;RootPath=$CloudRoot;AccessEvidenceSHA256=(& $objectHash ([ordered]@{ServerAcl=$acl;RootPermissions=$permissions;BatchesPermissions=$batchPermissions}))}
    }.GetNewClosure()
    $list={param([string]$ParentId)
        foreach($item in @(& $pages ($driveBase+'/items/'+[uri]::EscapeDataString($ParentId)+'/children'))){
            $file=& $readValue $item 'file';$folder=& $readValue $item 'folder'
            [pscustomobject]@{Id=[string]$item.id;Name=[string]$item.name;ParentId=[string]$item.parentReference.id;Kind=$(if($null -ne $folder){'Folder'}elseif($null -ne $file){'File'}else{'Unknown'});Linked=[bool]((& $readValue $item 'remoteItem') -or $item.parentReference.driveId -cne $drive.id);Bytes=[long]$item.size;ETag=[string](& $readValue $item 'eTag' '')}
        }
    }.GetNewClosure()
    $download={param([string]$Id,[string]$Destination)
        $ctx=Get-MgContext
        if($ctx.TenantId -ine $Connection.TenantId -or $ctx.ClientId -ine $Connection.AppId -or $ctx.AuthType -cne 'AppOnly'){throw 'Graph context changed before file download.'}
        # Stream authenticated content to disk; never buffer large evidence files as JSON or text.
        Invoke-MgGraphRequest -Method GET -Uri ($driveBase+'/items/'+[uri]::EscapeDataString($Id)+'/content') -OutputFilePath $Destination -ErrorAction Stop | Out-Null
        Test-Path -LiteralPath $Destination -PathType Leaf
    }.GetNewClosure()
    $archive={param([string]$LocalPath,[string]$Relative)
        if($Relative -cnotmatch '^retired/\d{8}T\d{9}Z-[a-f0-9]{8}/(current|batch|validation)\.json\.txt$'){throw 'Invalid receipt archive path.'}
        # Existing different receipts are never overwritten, including after an interrupted run.
        $existing=$null
        $archiveEncoded=(($CloudRoot+'/'+$Relative).Split('/') | ForEach-Object {[uri]::EscapeDataString($_)}) -join '/'
        try {$existing=& $request 'GET' ($driveBase+'/root:/'+$archiveEncoded) $null}
        catch {if((& $httpStatus $_) -ne 404){throw}}
        $temp=Join-Path $WorkRoot ([guid]::NewGuid().ToString('N')+'.archive-readback')
        try {
            if($existing){
                if((& $download $existing.id $temp) -ne $true -or (Get-FileHash -LiteralPath $temp).Hash -cne (Get-FileHash -LiteralPath $LocalPath).Hash){throw 'Remote archive already contains a different receipt.'}
                return $true
            }
            if((& $UploadFile $LocalPath $Relative) -ne $true -or (& $DownloadFile $temp $Relative) -ne $true){return $false}
            (Get-Item -LiteralPath $LocalPath).Length -eq (Get-Item -LiteralPath $temp).Length -and (Get-FileHash -LiteralPath $LocalPath).Hash -ceq (Get-FileHash -LiteralPath $temp).Hash
        } finally {if(Test-Path -LiteralPath $temp){& $assertPath $temp;[IO.File]::Delete($temp)}}
    }.GetNewClosure()
    $prepareRecycle={param([string]$Id,[string]$ETag)
        # Runs before RemovalRequested is journaled: any refusal here means no request was attempted.
        & $assertDeleteSupport
        if($Id -cnotmatch '^[A-Za-z0-9!._-]{1,400}$'){throw "Recycle refused before any request: SharePoint item ID '$Id' is empty or contains unexpected characters."}
        if([string]::IsNullOrWhiteSpace($ETag) -or $ETag.Length -gt 512 -or $ETag -match '[\r\n]'){throw "Recycle refused before any request: the reviewed folder eTag of item '$Id' is missing or malformed."}
        $text=$driveBase+'/items/'+[uri]::EscapeDataString($Id)
        $parsed=$null
        if(-not [uri]::TryCreate($text,[UriKind]::Absolute,[ref]$parsed) -or $parsed.Scheme -cne 'https' -or $parsed.Host -cne 'graph.microsoft.com' -or
           $parsed.Query -or $parsed.Fragment -or $parsed.AbsolutePath -cnotmatch '^/v1\.0/drives/[^/]+/items/[^/]+$'){throw "Recycle refused before any request: invalid Graph item URI '$text'."}
        [pscustomobject]@{Id=$Id;ETag=$ETag;Uri=$text}
    }.GetNewClosure()
    $recycle={param($Request)
        if($null -eq $Request -or [string]$Request.Uri -cne ($driveBase+'/items/'+[uri]::EscapeDataString([string]$Request.Id)) -or [string]::IsNullOrWhiteSpace([string]$Request.ETag)){
            throw 'Recycle requires the validated request prepared before RemovalRequested.'
        }
        $ctx=Get-MgContext
        if($ctx.TenantId -ine $Connection.TenantId -or $ctx.ClientId -ine $Connection.AppId -or $ctx.AuthType -cne 'AppOnly'){throw 'Graph context changed before recycle; no request was sent.'}
        & $graphDelete ([string]$Request.Uri) ([string]$Request.ETag)
    }.GetNewClosure()
    $absent={param([string]$Id)
        try {$null=& $request 'GET' ($driveBase+'/items/'+[uri]::EscapeDataString($Id)) $null;return $false}
        catch {if((& $httpStatus $_) -eq 404){return $true};throw}
    }.GetNewClosure()
    @{Describe=$describe;List=$list;Download=$download;Archive=$archive;PrepareRecycle=$prepareRecycle;Recycle=$recycle;IsAbsent=$absent}
}

Export-ModuleMember -Function Invoke-PreparedBatchRetention,New-PreparedRetentionGraphTransport


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCgNyr8eXBw8Nsv
# jXjQ0NkDSSCICqhooUmS+SyV2/y2/6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBBG2qJEj/PmeLc19YRmx2PlEDMO9VNAGdxilyoxuLvrMA0GCSqG
# SIb3DQEBAQUABIIBgFyEftI9Lzq/rljfIFZfTfjEvMIsSHcnnYryM+NkQXiL3bbK
# bMQv1K/azC4BqWomSkMSPt2dKctYDa1uF7w4yOvwVmbafZihoFfpy8E63df6eKnh
# N+hw0ki6M12raj+z48fj4B4hx33pSGOhXtGMCMhNssKC0/ZfhTAnzjV/FrnBnVsB
# IgV/v6uh0RiuL2sckudvaZLizi7pLsLtOzcOOF4I16m14SJgDdmk20e6i+vMJEHV
# u0UvVE1amb/MWacK9qIWQLYFb65V2vNgyjgPVnCSkyC+UuD572wDd20G2UOgdIZP
# 8RulUDsyDj4pcrTp31KkDpUknOE9xCmxY7CjpmIlvLosHsQD4CYl4pFetf16XXkb
# iCQCPqrj91n5PjUMHSQZ0mkYOO4mTXPNU5kRqPO9CEL1hebJUY+qsC2VgGiABIbl
# 9vOLZtcSO61XHm1CyGh9keD4+DpAaTgeTTfCc7L2u1baMYEgMdfUeRSsVaW+9Q1x
# U6MhGqzJdY+qiI6yhqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIxNzI3
# MjBaMC8GCSqGSIb3DQEJBDEiBCC4sHIJzjhXzQw7RGR4pCywk43tiMOicBckgWzz
# c8j5VDANBgkqhkiG9w0BAQEFAASCAgBQcmhzOUFNFpHIqYQYa6bEjq27FsmhuQJa
# D1WBCEWBxgknOFKtOHvKfofnv3mHFblDbr+gQWBxhlJsLhX8PtT7q55CGcUr7A4E
# HQDzyiFcL42G/pn66C5YESuJuHTiceee5op0XpNzE1Koo4eXkp0aXdZAvWPV0BDS
# 15Y7AJWFy/3L2/fJnSi2PeSVGhC4Ydo8SchoXatUA3Y+I9MQlTPoPHapn5LRP5uB
# hKq7rdVFmlasd0u5nPMRuLU1zDDZf1zu/e8n9m9coWJPr+M8JtW3T5UoEIdH2tZd
# blb+fv0P2MKjKDyDLH7X2Hsv2HlANIpTAlPq5BdoU6azhA2CAnl9eyt0HcBD39sl
# VLXo6aQlhZztODQvOXXE2EUO9LwOH+VPk11VXHZonX8jxWpb7QeERGodPCV0im9e
# mOf1lP45XGLXMAgOJRlPQR3/qM+4mkgrEalfDI8S9aHxhrs7APtI9qLPcgdAyfVb
# z7uhsy2VMlxMBNkQvK+hCQtwfzfS7N+wUbVhvceKsTHZvJq44mrJnpvRehyWdlas
# Pd3G7hnqRLdA1Zpsl0qjk2bDKe7Anb/D1MOhq5o/hcjG0KxA3oKIZaW0ter0KwJs
# fKawbn3Y+GFlUfFl6kJCfpHoE/VdWk+GLgjvMHx7Q0/NFxuGSHGyUETeQOmwDIt8
# MzRz276Vpw==
# SIG # End signature block
