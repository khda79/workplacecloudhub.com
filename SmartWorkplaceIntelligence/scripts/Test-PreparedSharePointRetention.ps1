<#
.SYNOPSIS
    Synthetic retention regression tests. File-backed fake SharePoint; no tenant or network.
.VERSION
    1.4.0
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot,[switch]$AdapterOnly)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
trap {Write-Host $_.Exception.Message;Write-Host $_.ScriptStackTrace;throw $_}
$modulePath=Join-Path $PSScriptRoot 'PreparedSharePointRetention.psm1'
Import-Module $modulePath -Force
$module=Get-Module PreparedSharePointRetention
$base=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ('retention-tests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $base -Force | Out-Null
$script:checks=0
function Check([bool]$Pass,[string]$Message){if(-not $Pass){throw $Message};$script:checks++}
function Reject([scriptblock]$Action,[string]$Pattern){$message='';try{& $Action | Out-Null}catch{$message=$_.Exception.Message};Check ($message -match $Pattern) "Expected '$Pattern', got '$message'"}
function Json([string]$Path,$Value){[IO.File]::WriteAllText($Path,($Value | ConvertTo-Json -Depth 70))}
function Fixture {
    $root=Join-Path $base ([guid]::NewGuid().ToString('N'));$output=Join-Path $root 'DATA/DATA-POWERBI';$remote=Join-Path $root 'cloud';$work=Join-Path $root 'work'
    New-Item -ItemType Directory -Path $output,$remote,$work -Force | Out-Null
    $ids=@(1..6 | ForEach-Object {'202610{0:00}T000000000Z-{0:x8}' -f $_})
    $previous=$null
    foreach($id in $ids){
        $folder=Join-Path $output ('batches/'+$id);New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $csv=Join-Path $folder 'Example.csv';[IO.File]::WriteAllText($csv,"TenantKey,Name`nsynthetic,fixture`n")
        $hash=(Get-FileHash $csv).Hash
        $entry=[ordered]@{File='Example.csv';Rows=1;Bytes=(Get-Item $csv).Length;SHA256=$hash}
        Json (Join-Path $folder 'batch.json.txt') ([ordered]@{SchemaVersion=1;BatchId=$id;TenantKey='synthetic';Files=@($entry)})
        Json (Join-Path $folder 'validation.json.txt') ([ordered]@{Passed=$true;SchemaOnly=$false;Errors=@();Files=@($entry)})
        $pointer=[ordered]@{SchemaVersion=1;BatchId=$id;TenantKey='synthetic';PreviousBatchId=$previous;ManifestSHA256=(Get-FileHash (Join-Path $folder 'batch.json.txt')).Hash}
        Json (Join-Path $folder 'current.json.txt') $pointer
        $previous=$id
    }
    Json (Join-Path $output 'current.json.txt') $pointer
    Copy-Item -LiteralPath (Join-Path $output 'batches') -Destination $remote -Recurse
    Copy-Item -LiteralPath (Join-Path $output 'current.json.txt') -Destination $remote
    $registry=Join-Path $root 'consumers.json.txt';$config=Join-Path $root 'prepare.local.json.txt'
    Json $registry ([ordered]@{SchemaVersion=1;TenantKey='synthetic';TenantId='synthetic-tenant';Qualification=@{Status='Qualified';ReviewedBy='Synthetic operator';ReviewedUtc='2026-01-01T00:00:00Z';DeclaredConsumersComplete=$true};Consumers=@(
        @{ConsumerId='Pinned desktop';BatchId=$ids[2];Reason='Synthetic old middle pin';ReviewedUtc='2026-01-01T00:00:00Z'},
        @{ConsumerId='Other consumer';BatchId=$ids[0];Reason='Synthetic oldest pin';ReviewedUtc='2026-01-01T00:00:00Z'})})
    $target=[ordered]@{TenantId='synthetic-tenant';ClientId='synthetic-app';SiteId='synthetic-site';DriveId='synthetic-drive';RootId='root';BatchesId='batches';RootPath='Example/DATA/DATA-POWERBI';AccessEvidenceSHA256=('A'*64)}
    $q=[ordered]@{Status='Qualified';PolicyVersion='PreparedBatchRetention/v1';ReviewedBy='Synthetic operator';ReviewedUtc='2026-01-01T00:00:00Z';Coordination='SharedServerPublicationLock';ServerOutputRoot=$output;RegistrySHA256=(Get-FileHash $registry).Hash;ModuleSHA256=(Get-FileHash $modulePath).Hash;AccessEvidenceSHA256=('A'*64);WriterExclusionEvidence='Synthetic isolated file-backed transport, no other writers';ExclusiveCloudWritersConfirmed=$true;RecoverabilityConfirmed=$true}
    # Legacy-equivalent policy keeps older coverage; dedicated cases below exercise the code defaults.
    Json $config @{PreparedRetention=@{Mode='Preview';PlanMaxAgeHours=24;JournalRootPath=$work;ServerKeepBatches=2;MaxSharePointRemovalsPerRun=10;Target=$target;Qualification=$q}}
    $state=@{Remote=$remote;Target=$target;Events=[Collections.Generic.List[string]]::new();FailArchive=$false;FailArchiveAt=0;Archives=0;FailDownload=$false;FailPayloadDownload=$false;PayloadDownloads=0;ReceiptDownloads=0;FailRecycleAt=0;Recycles=0;Uncertain=$false;ChangeAfterArchive=$false;ChangePayloadAfterArchive=$false;Linked=$false;WrongParent=$false;Recycled=@{};Hook=$null;ArchiveHook=$null;Prepares=0;PrepareFailure=$null;RecycleResult=$null;RemoveOnResult=$false}
    $describe={if($state.Hook){& $state.Hook};[pscustomobject]$state.Target}.GetNewClosure()
    $list={param($parent)
        $folder=if($parent -eq 'root'){$state.Remote}else{Join-Path $state.Remote $parent}
        if(-not(Test-Path $folder)){throw 'Synthetic missing remote parent'}
        foreach($item in Get-ChildItem -LiteralPath $folder -Force){
            $relative=[IO.Path]::GetRelativePath($state.Remote,$item.FullName).Replace('\','/')
            [pscustomobject]@{Id=$relative;Name=$item.Name;ParentId=$(if($state.WrongParent){'wrong'}else{$parent});Kind=$(if($item.PSIsContainer){'Folder'}else{'File'});Linked=$state.Linked;Bytes=$(if($item.PSIsContainer){0}else{$item.Length});ETag=$(if($item.PSIsContainer){'stable-folder-etag'}else{[string]$item.Length+'-'+$item.LastWriteTimeUtc.Ticks})}
        }
    }.GetNewClosure()
    $download={param($id,$destination)
        if([IO.Path]::GetFileName($id) -in @('current.json.txt','batch.json.txt','validation.json.txt')){$state.ReceiptDownloads++}
        else {$state.PayloadDownloads++;if($state.FailPayloadDownload){throw 'Synthetic forbidden payload download'}}
        if($state.FailDownload){return $false}
        Copy-Item -LiteralPath (Join-Path $state.Remote $id) -Destination $destination
        $true
    }.GetNewClosure()
    $archive={param($path,$relative)
        $state.Archives++;$state.Events.Add('Archive:'+$relative)
        if($state.ArchiveHook){& $state.ArchiveHook}
        if($state.FailArchive -or ($state.FailArchiveAt -and $state.Archives -eq $state.FailArchiveAt)){return $false}
        $destination=Join-Path $state.Remote $relative
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $path -Destination $destination
        if($state.ChangeAfterArchive){Add-Content (Join-Path $state.Remote 'current.json.txt') ' '}
        if($state.ChangePayloadAfterArchive){Add-Content (Join-Path $state.Remote ('batches/'+$ids[1]+'/Example.csv')) 'changed'}
        (Get-FileHash $path).Hash -eq (Get-FileHash $destination).Hash
    }.GetNewClosure()
    $prepare={param($id,$etag)
        $state.Prepares++
        if($state.PrepareFailure){throw $state.PrepareFailure}
        [pscustomobject]@{Id=$id;ETag=$etag;Uri=('https://graph.microsoft.com/v1.0/drives/synthetic-drive/items/'+[uri]::EscapeDataString($id))}
    }.GetNewClosure()
    $recycle={param($request)
        $id=$request.Id
        $state.Recycles++;$state.Events.Add('Recycle:'+$id)
        if($state.FailRecycleAt -and $state.Recycles -eq $state.FailRecycleAt){throw 'Synthetic recycle failure'}
        if($state.Uncertain){throw 'Synthetic timeout after request'}
        $result=[pscustomobject]@{Outcome='Recycled';HttpStatus=204;MayHaveReachedService=$false;Error=$null}
        if($state.RecycleResult){$result=[pscustomobject]$state.RecycleResult;if(-not $state.RemoveOnResult){return $result}}
        $source=Join-Path $state.Remote $id;$destination=Join-Path $root ('recycle/'+[IO.Path]::GetFileName($source))
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Move-Item -LiteralPath $source -Destination $destination
        $state.Recycled[$id]=$true
        $result
    }.GetNewClosure()
    $absent={param($id) -not(Test-Path -LiteralPath (Join-Path $state.Remote $id))}.GetNewClosure()
    [pscustomobject]@{Root=$root;Output=$output;Remote=$remote;Work=$work;Config=$config;Registry=$registry;Ids=$ids;State=$state;Args=@{OutputRoot=$output;TenantKey='synthetic';ConfigurationPath=$config;ConsumersPath=$registry;WorkRoot=$work;Verification='Full';Transport=@{Describe=$describe;List=$list;Download=$download;Archive=$archive;PrepareRecycle=$prepare;Recycle=$recycle;IsAbsent=$absent}}}
}
function SetPolicy($f,[hashtable]$Values){
    $c=Get-Content $f.Config -Raw | ConvertFrom-Json -Depth 70
    foreach($key in $Values.Keys){
        $c.PreparedRetention.PSObject.Properties.Remove($key)
        if($null -ne $Values[$key]){$c.PreparedRetention | Add-Member $key $Values[$key]}
    }
    Json $f.Config $c
}
function LastJournal([string]$Root){@(Get-ChildItem (Join-Path $Root 'retention') -Recurse -Filter journal.json.txt | Sort-Object LastWriteTimeUtc)[-1].FullName | ForEach-Object {Get-Content $_ -Raw | ConvertFrom-Json -Depth 70}}
function SetAuto($f){$c=Get-Content $f.Config -Raw | ConvertFrom-Json -Depth 70;$c.PreparedRetention.Mode='Auto';Json $f.Config $c}
function Requalify($f){$c=Get-Content $f.Config -Raw | ConvertFrom-Json -Depth 70;$c.PreparedRetention.Qualification.RegistrySHA256=(Get-FileHash $f.Registry).Hash;Json $f.Config $c}
function AdvanceFixture($f,[int]$Day){
    $old=Get-Content (Join-Path $f.Output 'current.json.txt') -Raw | ConvertFrom-Json
    $id='202610{0:00}T000000000Z-{0:x8}' -f $Day
    $folder=Join-Path $f.Output ('batches/'+$id)
    Copy-Item -LiteralPath (Join-Path $f.Output ('batches/'+$old.BatchId)) -Destination $folder -Recurse
    $manifest=Get-Content (Join-Path $folder 'batch.json.txt') -Raw | ConvertFrom-Json -Depth 70
    $manifest.BatchId=$id;Json (Join-Path $folder 'batch.json.txt') $manifest
    $pointer=[ordered]@{SchemaVersion=1;BatchId=$id;TenantKey='synthetic';PreviousBatchId=$old.BatchId;ManifestSHA256=(Get-FileHash (Join-Path $folder 'batch.json.txt')).Hash}
    Json (Join-Path $folder 'current.json.txt') $pointer
    Copy-Item -LiteralPath $folder -Destination (Join-Path $f.Remote 'batches') -Recurse
    Json (Join-Path $f.Output 'current.json.txt') $pointer
    Json (Join-Path $f.Remote 'current.json.txt') $pointer
}
function UpdateFixtureManifest($f,[string]$Id,$Manifest){
    $folder=Join-Path $f.Output ('batches/'+$Id)
    Json (Join-Path $folder 'batch.json.txt') $Manifest
    $pointer=Get-Content (Join-Path $folder 'current.json.txt') -Raw|ConvertFrom-Json
    $pointer.ManifestSHA256=(Get-FileHash (Join-Path $folder 'batch.json.txt')).Hash
    Json (Join-Path $folder 'current.json.txt') $pointer
    Copy-Item -LiteralPath $folder -Destination (Join-Path $f.Remote 'batches') -Recurse -Force
    if($Id -ceq $f.Ids[-1]){Json (Join-Path $f.Output 'current.json.txt') $pointer;Json (Join-Path $f.Remote 'current.json.txt') $pointer}
}
function AddFixtureTechnicalLocks($f){
    foreach($id in $f.Ids){
        $folder=Join-Path $f.Output ('batches/'+$id);$artifact=Join-Path $folder 'Source-Example.json.txt'
        Json $artifact @{Source='synthetic'}
        $manifest=Get-Content (Join-Path $folder 'batch.json.txt') -Raw|ConvertFrom-Json
        $manifest | Add-Member Artifacts @(@{File='Source-Example.json.txt';Bytes=(Get-Item $artifact).Length;SHA256=(Get-FileHash $artifact).Hash})
        UpdateFixtureManifest $f $id $manifest
        foreach($owner in 'current.json.txt','batch.json.txt','validation.json.txt','Source-Example.json.txt'){
            foreach($name in ($owner.Substring(0,$owner.Length-4)+'.transport.lock'),($owner+'.sharepoint-transition.lock')){
                [IO.File]::WriteAllBytes((Join-Path $folder $name),[byte[]]@())
            }
        }
    }
}
function RemoveFixtureBatchCopy($f,[string]$Storage,[string]$Id){
    $root=if($Storage -ceq 'Local'){$f.Output}else{$f.Remote}
    $path=[IO.Path]::GetFullPath((Join-Path $root ('batches/'+$Id)))
    $expectedRoot=[IO.Path]::GetFullPath((Join-Path $root 'batches')).TrimEnd('\')+'\'
    if($Id -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$' -or -not $path.StartsWith($expectedRoot,[StringComparison]::OrdinalIgnoreCase) -or
       -not $path.StartsWith(([IO.Path]::GetFullPath($base).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)){throw 'Unsafe synthetic fixture removal.'}
    Remove-Item -LiteralPath $path -Recurse -Force
}
function AddUnpointedFixtureBatch($f,[string]$Id){
    $source=Join-Path $f.Output ('batches/'+$f.Ids[-1]);$folder=Join-Path $f.Output ('batches/'+$Id)
    Copy-Item -LiteralPath $source -Destination $folder -Recurse
    $manifest=Get-Content (Join-Path $folder 'batch.json.txt') -Raw|ConvertFrom-Json
    $manifest.BatchId=$Id
    $pointer=Get-Content (Join-Path $folder 'current.json.txt') -Raw|ConvertFrom-Json
    $pointer.BatchId=$Id;Json (Join-Path $folder 'current.json.txt') $pointer
    UpdateFixtureManifest $f $Id $manifest
}

if(-not $AdapterOnly){
# The minimum is independent of pointer adjacency and is additive across both stores.
foreach($verification in 'Metadata','Full'){
    $f=Fixture;$a=@{}+$f.Args;$a.Verification=$verification
    RemoveFixtureBatchCopy $f 'SharePoint' $f.Ids[-2]
    $p=Invoke-PreparedBatchRetention @a
    $local=@($p.Plan.MinimumRetainedBatches|Where-Object Storage -CEQ 'Local')[0]
    $cloud=@($p.Plan.MinimumRetainedBatches|Where-Object Storage -CEQ 'SharePoint')[0]
    Check (($local.BatchIds -join ',') -ceq ($f.Ids[-1..-2] -join ',')) "$verification local minimum changed with remote predecessor absence"
    Check (($cloud.BatchIds -join ',') -ceq (@($f.Ids[-1],$f.Ids[-3]) -join ',')) "$verification failed to protect the second available SharePoint batch"
    Check (@($p.Plan.Rows|Where-Object {$_.BatchId -ceq $f.Ids[-3] -and $_.State -ceq 'Protected'}).Count -eq 2) 'A per-store minimum was not propagated to both copies'
    foreach($id in $f.Ids[0],$f.Ids[2],$f.Ids[-2],$f.Ids[-1]){Check ($p.Plan.ProtectedBatchIds -ccontains $id) 'Existing pointer/consumer protection lost'}
    Check ($f.State.Recycles -eq 0) 'Minimum planning removed data'
    if($verification -ceq 'Metadata'){Check ($f.State.PayloadDownloads -eq 0) 'Metadata minimum downloaded payloads'}
}
# A complete newer batch outside the current pointer and timestamp ties are also retained.
foreach($case in 'Newer','Tie'){
    $f=Fixture;$a=$f.Args
    $id=if($case -ceq 'Newer'){'20261007T000000000Z-00000007'}else{'20261005T000000000Z-00000009'}
    AddUnpointedFixtureBatch $f $id
    # Reverse modification-time order: the batch ID timestamp must remain authoritative.
    foreach($root in $f.Output,$f.Remote){(Get-Item (Join-Path $root ('batches/'+$f.Ids[1]))).LastWriteTimeUtc=[datetime]::UtcNow.AddDays(30)}
    $p=Invoke-PreparedBatchRetention @a
    Check ($p.Plan.ProtectedBatchIds -ccontains $id) "$case batch was not retained"
    Check (@($p.Plan.Rows|Where-Object {$_.BatchId -ceq $f.Ids[1] -and $_.State -ceq 'Eligible'}).Count -eq 2) 'Modification time incorrectly affected chronology'
    $expected=if($case -ceq 'Tie'){3}else{2}
    Check (@($p.Plan.MinimumRetainedBatches|Where-Object {$_.BatchIds.Count -ne $expected}).Count -eq 0) 'Minimum timestamp boundary/tie failed'
    $r=Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256
    Check ($r.Status -ceq 'Completed' -and $f.State.Recycles -eq 2) 'New minimum broke eligible older retirement'
    foreach($root in $f.Output,$f.Remote){Check (Test-Path (Join-Path $root ('batches/'+$id))) 'Newest/tied copy was deleted'}
}
# A newer incomplete, schema-only, failed, corrupt or invalid-date batch cannot displace valid copies.
foreach($case in 'MissingPayload','MissingReceipt','SchemaOnly','Failed','SameSizeCorruption','InvalidTimestamp'){
    $f=Fixture;$a=$f.Args
    $id=if($case -ceq 'InvalidTimestamp'){'20261307T000000000Z-00000007'}else{'20261007T000000000Z-00000007'}
    AddUnpointedFixtureBatch $f $id
    foreach($root in $f.Output,$f.Remote){
        $folder=Join-Path $root ('batches/'+$id)
        switch($case){
            MissingPayload {[IO.File]::Delete((Join-Path $folder 'Example.csv'))}
            MissingReceipt {[IO.File]::Delete((Join-Path $folder 'validation.json.txt'))}
            SchemaOnly {$v=Get-Content (Join-Path $folder 'validation.json.txt') -Raw|ConvertFrom-Json;$v.SchemaOnly=$true;Json (Join-Path $folder 'validation.json.txt') $v}
            Failed {$v=Get-Content (Join-Path $folder 'validation.json.txt') -Raw|ConvertFrom-Json;$v.Passed=$false;Json (Join-Path $folder 'validation.json.txt') $v}
            SameSizeCorruption {$path=Join-Path $folder 'Example.csv';[IO.File]::WriteAllText($path,([IO.File]::ReadAllText($path).Replace('fixture','fixturX')))}
        }
    }
    $p=Invoke-PreparedBatchRetention @a
    foreach($store in $p.Plan.MinimumRetainedBatches){Check (($store.BatchIds -join ',') -ceq ($f.Ids[-1..-2] -join ',')) "$case displaced valid retained copies"}
    Check (@($p.Plan.Rows|Where-Object {$_.BatchId -ceq $id -and $_.State -ceq 'Blocked'}).Count -eq 2) "$case became eligible"
}
# Missing/invalid second copies block all application; zero/one valid copies are never enough.
foreach($count in 0,1,2){
    $f=Fixture;$a=$f.Args
    foreach($id in $f.Ids|Select-Object -First (6-$count)){RemoveFixtureBatchCopy $f 'SharePoint' $id}
    $p=Invoke-PreparedBatchRetention @a
    Check ((@($p.Plan.BlockedReasons|Where-Object {$_ -match 'Fewer than two'}).Count -gt 0) -eq ($count -lt 2)) 'Wrong minimum-count decision'
    if($count -lt 2){
        Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256} 'blocked by qualification'
        Check ($f.State.Recycles -eq 0 -and $f.State.Archives -eq 0) 'Insufficient minimum reached archival/removal'
    }
}
# Same-size corruption is provisional in Metadata; Full must select the older valid fallback.
$f=Fixture;$a=@{}+$f.Args;$a.Verification='Metadata';$payload=Join-Path $f.Remote ('batches/'+$f.Ids[-2]+'/Example.csv')
[IO.File]::WriteAllText($payload,([IO.File]::ReadAllText($payload).Replace('fixture','fixturX')))
$p=Invoke-PreparedBatchRetention @a
Check (@($p.Plan.Rows|Where-Object PayloadSHA256Verified).Count -eq 0 -and $f.State.PayloadDownloads -eq 0) 'Metadata minimum claimed full verification'
$a.Verification='Full';$p=Invoke-PreparedBatchRetention @a
$cloud=@($p.Plan.MinimumRetainedBatches|Where-Object Storage -CEQ 'SharePoint')[0]
Check (($cloud.BatchIds -join ',') -ceq (@($f.Ids[-1],$f.Ids[-3]) -join ',')) 'Full did not replace corrupt second-newest with a valid older copy'
Check ($p.Status -ceq 'Blocked') 'Corrupt existing predecessor lost its global protection'
# Changed or removed retained payloads after planning/archival stop BEFORE the next deletion.
foreach($storage in 'Local','SharePoint'){
    foreach($case in 'Missing','Corrupt','DownloadFailure','BeforeLocalRemoval'){
        if($storage -ceq 'Local' -and $case -ceq 'DownloadFailure'){continue}
        $f=Fixture;SetAuto $f;$a=$f.Args;$state=$f.State
        $root=if($storage -ceq 'Local'){$f.Output}else{$f.Remote}
        $payload=Join-Path $root ('batches/'+$f.Ids[-1]+'/Example.csv')
        $trigger=if($case -ceq 'BeforeLocalRemoval'){7}else{1}
        $f.State.ArchiveHook={
            if($state.Archives -eq $trigger){
                if($case -ceq 'DownloadFailure'){$state.FailPayloadDownload=$true}
                elseif($case -ceq 'Corrupt'){[IO.File]::WriteAllText($payload,([IO.File]::ReadAllText($payload).Replace('fixture','fixturX')))}
                else {[IO.File]::Delete($payload)}
            }
        }.GetNewClosure()
        Reject {Invoke-PreparedBatchRetention @a -Mode Configured} 'minimum changed|forbidden payload download'
        $expectedRecycles=if($case -ceq 'BeforeLocalRemoval'){2}else{0}
        Check ($f.State.Recycles -eq $expectedRecycles) "$storage/$case failed to stop before the next cloud recycle"
        Check (@(Get-ChildItem (Join-Path $f.Output 'batches') -Directory).Count -eq 6) "$storage/$case removed a local batch"
    }
}
# The new per-store protection evidence is part of the reviewed plan contract.
$f=Fixture;$a=$f.Args;$p=Invoke-PreparedBatchRetention @a
$tampered=Get-Content $p.PlanPath -Raw|ConvertFrom-Json -Depth 70;$tampered.MinimumRetainedBatches=@();Json $p.PlanPath $tampered
Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 (Get-FileHash $p.PlanPath).Hash} 'actions/protections'
Check ($f.State.Recycles -eq 0) 'Tampered minimum authorized removal'
# Known idle locks belong to local infrastructure, remain in physical evidence, and do not split copies.
foreach($verification in 'Metadata','Full'){
    $f=Fixture;AddFixtureTechnicalLocks $f;$a=@{}+$f.Args;$a.Verification=$verification
    $p=Invoke-PreparedBatchRetention @a
    Check ($p.Status -ceq 'Preview' -and @($p.Plan.Rows|Where-Object {-not $_.Valid}).Count -eq 0) "$verification rejected known idle JSON locks"
    Check (@($p.Plan.Rows|Where-Object State -CEQ 'Protected').Count -eq 8) "$verification lost protected batches with locks"
    Check (@($p.Plan.Rows|Where-Object {$_.Storage -ceq 'Local' -and $_.TechnicalLockNames.Count -eq 8 -and $_.FileCount -eq 13}).Count -eq 6) 'Physical lock inventory was dropped'
    Check (@($p.Plan.Rows.Files|Where-Object {$_.PSObject.Properties['Value']}).Count -eq 0) 'Plan embeds parsed receipt bodies'
    Check ($p.Plan.SchemaVersion -eq 3 -and (Get-Item $p.PlanPath).Length -lt 150000) 'Compact plan schema/size regression'
    if($verification -ceq 'Metadata'){
        Check ($f.State.PayloadDownloads -eq 0 -and $f.State.Recycles -eq 0) 'Lock-aware metadata Preview downloaded payloads or recycled'
    }else{
        $lockPath=Join-Path $f.Output ('batches/'+$f.Ids[1]+'/batch.json.transport.lock')
        $probe=@{Calls=0}
        $f.State.ArchiveHook={
            if(Test-Path $lockPath){
                $denied=$false;try{$h=[IO.File]::Open($lockPath,'Open','ReadWrite','None');$h.Dispose()}catch{$denied=$true}
                if(-not $denied){throw 'Technical helper writer was not excluded during archive/removal'}
                $probe.Calls++
            }
        }.GetNewClosure()
        # Check only this batch's archival window; unrelated batch locks need not be held.
        $hook=$f.State.ArchiveHook
        $fixtureState=$f.State;$fixtureId=$f.Ids[1]
        $f.State.ArchiveHook={if($fixtureState.Events[-1] -like ('Archive:retired/'+$fixtureId+'/*')){& $hook}}.GetNewClosure()
        $r=Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256
        Check ($r.Status -ceq 'Completed' -and $f.State.Recycles -eq 2 -and $probe.Calls -eq 6) 'Full compact plan/lock-guarded retirement failed'
        Check ((Test-Path (Join-Path $f.Output ('batches/'+$f.Ids[2]))) -and (Test-Path (Join-Path $f.Remote ('batches/'+$f.Ids[2])))) 'Full compact application removed consumer pin'
    }
}
foreach($verification in 'Metadata','Full'){
    foreach($case in 'Unknown','Orphan','CsvOwner','WrongCase','Nonempty','Active','Remote'){
        $f=Fixture;$a=@{}+$f.Args;$a.Verification=$verification
        $folder=Join-Path $f.Output ('batches/'+$f.Ids[1]);$held=$null
        $name=switch($case){Unknown {'other.lock'};Orphan {'missing.json.transport.lock'};CsvOwner {'Example.csv.sharepoint-transition.lock'};WrongCase {'Batch.json.transport.lock'};default {'batch.json.transport.lock'}}
        if($case -ceq 'Remote'){$folder=Join-Path $f.Remote ('batches/'+$f.Ids[1])}
        $path=Join-Path $folder $name;[IO.File]::WriteAllBytes($path,[byte[]]@())
        if($case -ceq 'Nonempty'){[IO.File]::WriteAllText($path,'owner data')}
        if($case -ceq 'Active'){$held=[IO.File]::Open($path,'Open','ReadWrite','None')}
        try {$p=Invoke-PreparedBatchRetention @a}finally{if($held){$held.Dispose()}}
        Check (@($p.Plan.Rows|Where-Object {$_.BatchId -ceq $f.Ids[1] -and $_.State -ceq 'Blocked'}).Count -eq 2) "$verification $case lock did not block both copies"
        Check ($f.State.Recycles -eq 0 -and (Test-Path $path)) "$verification $case lock was modified or removed"
    }
}
# Active locks on protected batches stay visible as a validation failure, never become deletion candidates.
$f=Fixture;$a=@{}+$f.Args;$a.Verification='Metadata'
$path=Join-Path $f.Output ('batches/'+$f.Ids[2]+'/batch.json.transport.lock');[IO.File]::WriteAllBytes($path,[byte[]]@())
$held=[IO.File]::Open($path,'Open','ReadWrite','None')
try {$p=Invoke-PreparedBatchRetention @a}finally{$held.Dispose()}
$row=@($p.Plan.Rows|Where-Object {$_.Storage -ceq 'Local' -and $_.BatchId -ceq $f.Ids[2]})[0]
Check ($row.State -ceq 'Protected' -and -not $row.Valid -and $row.ValidationReason -match 'Technical lock unavailable') 'Protected active-lock diagnostic was hidden'
Check ($p.Status -ceq 'Blocked' -and $f.State.Recycles -eq 0) 'Protected active-lock disagreement did not block application'
# Recheck availability after a reviewed plan, before cloud mutation.
$f=Fixture;AddFixtureTechnicalLocks $f;$a=$f.Args;$p=Invoke-PreparedBatchRetention @a
$path=Join-Path $f.Output ('batches/'+$f.Ids[1]+'/batch.json.transport.lock');$held=[IO.File]::Open($path,'Open','ReadWrite','None')
try {Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256} 'changed|differ|blocked'}finally{$held.Dispose()}
Check ($f.State.Recycles -eq 0) 'Newly active technical lock permitted cloud retirement'
$f=Fixture;AddFixtureTechnicalLocks $f;$a=$f.Args;$p=Invoke-PreparedBatchRetention @a
$race=@{Calls=0;Handle=$null};$path=Join-Path $f.Output ('batches/'+$f.Ids[1]+'/batch.json.transport.lock')
$f.State.Hook={
    $race.Calls++
    if($race.Calls -eq 2){$race.Handle=[IO.File]::Open($path,'Open','ReadWrite','None')}
}.GetNewClosure()
try {Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256} 'Technical lock unavailable'}finally{if($race.Handle){$race.Handle.Dispose()}}
Check ($f.State.Recycles -eq 0 -and $f.State.Archives -eq 0) 'Lock activated after snapshot was missed before archival/cloud retirement'
# Large provenance remains bound by raw hashes but never copied into plans or snapshot serialization.
$f=Fixture;$a=@{}+$f.Args;$a.Verification='Metadata';$id=$f.Ids[-1]
$manifest=Get-Content (Join-Path $f.Output ('batches/'+$id+'/batch.json.txt')) -Raw|ConvertFrom-Json
$manifest | Add-Member Provenance ('SYNTHETIC-PROVENANCE-' + ('x'*2000000));UpdateFixtureManifest $f $id $manifest
$p=Invoke-PreparedBatchRetention @a;$text=Get-Content $p.PlanPath -Raw
Check ($text.Length -lt 60000 -and $text -notmatch 'SYNTHETIC-PROVENANCE|"Value"') 'Large provenance leaked into compact plan'
$receipt=@($p.Plan.Rows|Where-Object {$_.Storage -ceq 'Local' -and $_.BatchId -ceq $id})[0].Files|Where-Object Name -CEQ 'batch.json.txt'
Check ($receipt.SHA256 -ceq (Get-FileHash (Join-Path $f.Output ('batches/'+$id+'/batch.json.txt'))).Hash) 'Compact plan lost original manifest fingerprint'
$manifest.Provenance=$manifest.Provenance.Replace('xxxx','yyyy');UpdateFixtureManifest $f $id $manifest
$changed=Invoke-PreparedBatchRetention @a
Check ($p.Plan.SnapshotSHA256 -cne $changed.Plan.SnapshotSHA256) 'Omitted provenance changed without changing snapshot fingerprint'
# Old schemas and modifications to compact evidence cannot authorize removal.
$f=Fixture;$a=$f.Args;$p=Invoke-PreparedBatchRetention @a
$old=Get-Content $p.PlanPath -Raw|ConvertFrom-Json -Depth 70;$old.SchemaVersion=2;Json $p.PlanPath $old
Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 (Get-FileHash $p.PlanPath).Hash} 'Full verification plan'
$p=Invoke-PreparedBatchRetention @a;$tampered=Get-Content $p.PlanPath -Raw|ConvertFrom-Json -Depth 70;$tampered.Rows[0].Files[0].SHA256='F'*64;Json $p.PlanPath $tampered
Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 (Get-FileHash $p.PlanPath).Hash} 'actions/protections'
Check ($f.State.Recycles -eq 0) 'Changed compact evidence authorized removal'
# Fast Preview must never request payload bytes, even when the registry is not yet qualified.
$f=Fixture;$a=@{}+$f.Args;$a.Remove('Verification');$f.State.FailPayloadDownload=$true
$held=[IO.File]::Open((Join-Path $f.Output ('batches/'+$f.Ids[0]+'/Example.csv')),'Open','ReadWrite','None')
try {$p=Invoke-PreparedBatchRetention @a}finally{$held.Dispose()}
Check ($p.Status -ceq 'Preview' -and $p.Plan.Verification -ceq 'Metadata' -and -not $p.Plan.ApplySupported) 'Default Preview must be metadata-only and non-applicable'
Check ($f.State.PayloadDownloads -eq 0 -and $f.State.ReceiptDownloads -eq 19) 'Fast Preview downloaded payloads or unexpected receipts'
Check (@($p.Plan.Rows | Where-Object State -CEQ 'Protected').Count -eq 8) 'Fast Preview lost current/previous/consumer protections'
Check (@($p.Plan.Rows | Where-Object State -CEQ 'Candidate').Count -eq 4 -and @($p.Plan.Rows | Where-Object State -CEQ 'Eligible').Count -eq 0) 'Metadata candidates were presented as fully verified eligible actions'
Check (@($p.Plan.Rows | Where-Object PayloadSHA256Verified).Count -eq 0) 'Fast Preview claimed payload SHA256 evidence'
Check ($f.State.Events.Count -eq 0 -and @(Get-ChildItem (Join-Path $f.Output 'batches')).Count -eq 6) 'Fast Preview mutated batches'
$reads=$f.State.ReceiptDownloads
Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256} 'Full verification plan'
Check ($f.State.ReceiptDownloads -eq $reads -and $f.State.Recycles -eq 0) 'Metadata plan Apply did not fail before remote scanning'

$f=Fixture;$a=@{}+$f.Args;$a.Verification='Metadata';$r=Get-Content $f.Registry -Raw|ConvertFrom-Json -Depth 70
$r.Qualification.Status='Unqualified';$r.Qualification.DeclaredConsumersComplete=$false;Json $f.Registry $r
$p=Invoke-PreparedBatchRetention @a
Check ($p.Status -ceq 'Blocked' -and $p.Plan.ProtectedBatchIds -ccontains $f.Ids[2]) 'Unqualified but structurally valid consumer pins must remain visible and fail closed'
Check (@($p.Plan.Rows|Where-Object {$_.BatchId -ceq $f.Ids[2] -and $_.State -ceq 'Protected'}).Count -eq 2) 'Unqualified registry hid the old pinned batch'

foreach($case in 'MissingPayload','WrongSize','MissingReceipt','DifferentReceipts'){
    $f=Fixture;$a=@{}+$f.Args;$a.Verification='Metadata';$folder=Join-Path $f.Remote ('batches/'+$f.Ids[1])
    switch($case){
        MissingPayload {Remove-Item (Join-Path $folder 'Example.csv')}
        WrongSize {Add-Content (Join-Path $folder 'Example.csv') 'changed'}
        MissingReceipt {Remove-Item (Join-Path $folder 'validation.json.txt')}
        DifferentReceipts {Add-Content (Join-Path $folder 'validation.json.txt') ' '}
    }
    $p=Invoke-PreparedBatchRetention @a
    Check (@($p.Plan.Rows|Where-Object {$_.BatchId -ceq $f.Ids[1] -and $_.State -ceq 'Blocked'}).Count -gt 0) "$case was accepted by metadata Preview"
    Check ($f.State.PayloadDownloads -eq 0 -and $f.State.Recycles -eq 0) "$case triggered payload download or removal"
}

# A same-size payload corruption is explicitly unmeasured in Metadata, then rejected by Full.
$f=Fixture;$a=@{}+$f.Args;$a.Verification='Metadata';$payload=Join-Path $f.Remote ('batches/'+$f.Ids[1]+'/Example.csv')
$text=[IO.File]::ReadAllText($payload);[IO.File]::WriteAllText($payload,$text.Replace('fixture','fixturX'))
$p=Invoke-PreparedBatchRetention @a
Check (@($p.Plan.Rows|Where-Object {$_.BatchId -ceq $f.Ids[1] -and $_.State -ceq 'Candidate'}).Count -eq 2) 'Same-size corruption must remain unmeasured, never fully verified in metadata Preview'
$a.Verification='Full';$full=Invoke-PreparedBatchRetention @a
Check (@($full.Plan.Rows|Where-Object {$_.Storage -ceq 'SharePoint' -and $_.BatchId -ceq $f.Ids[1] -and $_.State -ceq 'Blocked'}).Count -eq 1) 'Full verification missed same-size corruption'
$a.Verification='Metadata'
$r=Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $full.PlanPath -PlanSHA256 $full.PlanSHA256
Check ($r.Status -ceq 'Completed' -and (Test-Path $payload)) 'Manual Apply weakened full verification or removed a corrupt batch'
Check ($r.Plan.Verification -ceq 'Full' -and $f.State.PayloadDownloads -gt 0) 'Apply accepted Metadata instead of performing full verification'

$f=Fixture;$a=@{}+$f.Args;$a.Verification='Metadata';SetAuto $f
$p=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($p.Status -ceq 'Completed' -and $p.Plan.Verification -ceq 'Full' -and $f.State.PayloadDownloads -gt 0) 'Configured Auto weakened full SHA256 verification'
Check ((Test-Path (Join-Path $f.Output ('batches/'+$f.Ids[2]))) -and (Test-Path (Join-Path $f.Remote ('batches/'+$f.Ids[2])))) 'Auto lost the old consumer pin'

$f=Fixture;$a=$f.Args;$p=Invoke-PreparedBatchRetention @a
Check ($p.Status -ceq 'Preview') 'Qualified preview failed'
Check ($f.State.Events.Count -eq 0) 'Preview called archive/recycle'
Check (@($p.Plan.Rows | Where-Object State -CEQ 'Protected').Count -eq 8) 'Current, previous and two consumers must protect both stores'
Check (@($p.Plan.Rows | Where-Object State -CEQ 'Eligible').Count -eq 4) 'Old pin stopped traversal of other old batches'
Check ((Get-ChildItem (Join-Path $f.Output 'batches')).Count -eq 6) 'Preview removed local payload'
$r=Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256
Check ($r.Status -ceq 'Completed') 'Manual plan apply failed'
Check ($f.State.Recycles -eq 2) 'Wrong number of cloud retirements'
Check ((Get-ChildItem (Join-Path $f.Output 'batches')).Count -eq 4) 'Wrong local retained set'
Check ((Get-ChildItem (Join-Path $f.Remote 'batches')).Count -eq 4) 'Local/cloud retained sets differ'
Check ((Get-ChildItem (Join-Path $f.Output 'retired') -File -Recurse).Count -eq 6) 'Verified receipt archives missing'
Check (-not @(Get-ChildItem (Join-Path $f.Output 'retired') -Filter '*.csv' -Recurse).Count) 'Archives must not retain payload CSVs'
Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256} 'changed/stale|actions/protections'

$f=Fixture;SetAuto $f;$a=$f.Args
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($r.Status -ceq 'Completed' -and $f.State.Recycles -eq 2) 'Configured Auto did not retire automatically'
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($r.Status -ceq 'Completed' -and $f.State.Recycles -eq 2) 'Auto was not idempotent'
$policyHash=(Get-FileHash $f.Config).Hash
foreach($day in 7,8){AdvanceFixture $f $day;$r=Invoke-PreparedBatchRetention @a -Mode Configured;Check ($r.Status -ceq 'Completed') 'A new batch unexpectedly required policy requalification'}
Check ((Get-FileHash $f.Config).Hash -ceq $policyHash -and $f.State.Recycles -eq 4) 'Automatic multi-publication retention changed policy or missed a retirement'
# Releasing a former pin still follows archived links through the two already retired middle lots.
$registry=Get-Content $f.Registry -Raw | ConvertFrom-Json -Depth 70;$registry.Consumers=@($registry.Consumers | Where-Object ConsumerId -NE 'Pinned desktop');Json $f.Registry $registry;Requalify $f
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($f.State.Recycles -eq 5 -and -not(Test-Path (Join-Path $f.Output ('batches/'+$f.Ids[2])))) 'Archived links did not permit former consumer retirement'

foreach($case in 'MissingRegistry','BrokenRegistry','UnqualifiedRegistry','BadReference','WrongTenant','ChangedAccess','UnqualifiedPolicy','WrongRoot','WrongClient'){
    $f=Fixture;$a=$f.Args
    switch($case){
        MissingRegistry {[IO.File]::Delete($f.Registry)}
        BrokenRegistry {[IO.File]::WriteAllText($f.Registry,'broken')}
        UnqualifiedRegistry {$v=Get-Content $f.Registry -Raw|ConvertFrom-Json -Depth 70;$v.Qualification.Status='Unqualified';Json $f.Registry $v;Requalify $f}
        BadReference {$v=Get-Content $f.Registry -Raw|ConvertFrom-Json -Depth 70;$v.Consumers[0].BatchId='../bad';Json $f.Registry $v;Requalify $f}
        WrongTenant {$f.State.Target.TenantId='foreign'}
        ChangedAccess {$f.State.Target.AccessEvidenceSHA256='B'*64}
        UnqualifiedPolicy {$v=Get-Content $f.Config -Raw|ConvertFrom-Json -Depth 70;$v.PreparedRetention.Qualification.Status='Unqualified';Json $f.Config $v}
        WrongRoot {$f.State.Target.RootId='foreign-root'}
        WrongClient {$f.State.Target.ClientId='foreign-app'}
    }
    if($case -ceq 'WrongRoot'){Reject {Invoke-PreparedBatchRetention @a} 'missing remote parent'}
    else {$r=Invoke-PreparedBatchRetention @a;Check ($r.Status -ceq 'Blocked') "$case must block preview/application"}
    Check ($f.State.Recycles -eq 0) "$case made a delete call"
}
foreach($case in 'Pointer','Registry','Payload','Unexpected','Stale','TamperedPlan'){
    $f=Fixture;$a=$f.Args;$p=Invoke-PreparedBatchRetention @a
    switch($case){
        Pointer {Add-Content (Join-Path $f.Remote 'current.json.txt') ' '}
        Registry {Add-Content $f.Registry ' '}
        Payload {Add-Content (Join-Path $f.Remote ('batches/'+$f.Ids[1]+'/Example.csv')) 'changed'}
        Unexpected {[IO.File]::WriteAllText((Join-Path $f.Remote ('batches/'+$f.Ids[1]+'/unexpected.txt')),'keep')}
        Stale {$v=Get-Content $p.PlanPath -Raw|ConvertFrom-Json -Depth 70;$v.CreatedUtc='2020-01-01T00:00:00Z';Json $p.PlanPath $v;$p.PlanSHA256=(Get-FileHash $p.PlanPath).Hash}
        TamperedPlan {Add-Content $p.PlanPath ' '}
    }
    Reject {Invoke-PreparedBatchRetention @a -Mode Apply -PlanPath $p.PlanPath -PlanSHA256 $p.PlanSHA256} 'stale|changed|hash mismatch|blocked|actions/protections'
    Check ($f.State.Recycles -eq 0) "$case made a delete call"
}
foreach($case in 'IncompleteLegacy','BadManifest','UnexpectedFile','InvalidId','StringBoolean','NegativeRows','MissingRows'){
    $f=Fixture;$a=$f.Args
    $folder=Join-Path $f.Remote ('batches/'+$f.Ids[1])
    switch($case){
        IncompleteLegacy {[IO.File]::Delete((Join-Path $folder 'batch.json.txt'))}
        BadManifest {Add-Content (Join-Path $folder 'batch.json.txt') ' '}
        UnexpectedFile {[IO.File]::WriteAllText((Join-Path $folder 'keep.txt'),'keep')}
        InvalidId {$folder=Join-Path $f.Remote 'batches/invalid';New-Item -ItemType Directory $folder|Out-Null}
        StringBoolean {$v=Get-Content (Join-Path $folder 'validation.json.txt') -Raw|ConvertFrom-Json -Depth 70;$v.Passed='true';Json (Join-Path $folder 'validation.json.txt') $v}
        NegativeRows {$v=Get-Content (Join-Path $folder 'validation.json.txt') -Raw|ConvertFrom-Json -Depth 70;$v.Files[0].Rows=-1;Json (Join-Path $folder 'validation.json.txt') $v}
        MissingRows {$v=Get-Content (Join-Path $folder 'validation.json.txt') -Raw|ConvertFrom-Json -Depth 70;$v.Files[0].PSObject.Properties.Remove('Rows');Json (Join-Path $folder 'validation.json.txt') $v}
    }
    $r=Invoke-PreparedBatchRetention @a
    Check (@($r.Plan.Rows | Where-Object {$_.Storage -ceq 'SharePoint' -and $_.State -ceq 'Blocked'}).Count -gt 0) "$case was accepted"
    Check ($f.State.Recycles -eq 0) 'Preview called recycle'
}
foreach($case in 'Archive','Download','Uncertain','Partial'){
    $f=Fixture;SetAuto $f;$a=$f.Args
    switch($case){Archive {$f.State.FailArchive=$true};Download {$f.State.FailDownload=$true};Uncertain {$f.State.Uncertain=$true};Partial {$f.State.FailRecycleAt=2}}
    Reject {Invoke-PreparedBatchRetention @a -Mode Configured} 'archive|read-back|timeout|recycle failure'
    Check ((Get-ChildItem (Join-Path $f.Output 'batches')).Count -eq 6) "$case removed local data without cloud confirmation"
    $logs=@(Get-ChildItem $f.Work -Recurse -Filter journal.json.txt)
    Check ($logs.Count -eq 1 -and (Get-Content $logs[0].FullName -Raw | ConvertFrom-Json).Status -ceq 'PartialOrBlocked') 'Exact failure journal missing'
    if($case -ceq 'Archive'){Check ($f.State.Recycles -eq 0) 'Archive failure made a recycle call'}
    if($case -ceq 'Partial'){
        $f.State.FailRecycleAt=0;$priorCalls=$f.State.Recycles;$r=Invoke-PreparedBatchRetention @a -Mode Configured
        Check ($r.Status -ceq 'Completed' -and (Get-ChildItem (Join-Path $f.Output 'batches')).Count -eq 5) 'Confirmed first remote result did not permit controlled local recovery'
        Check ($f.State.Recycles -eq $priorCalls) 'Uncertain second request was blindly repeated'
        Check (@($r.Plan.Rows | Where-Object Reason -Match 'Unresolved prior').Count -eq 2) 'Uncertain lot must block both stores'
    }
}
foreach($case in 'PointerDuringArchive','PayloadDuringArchive'){
    $f=Fixture;SetAuto $f;$a=$f.Args
    if($case -ceq 'PointerDuringArchive'){$f.State.ChangeAfterArchive=$true}else{$f.State.ChangePayloadAfterArchive=$true}
    Reject {Invoke-PreparedBatchRetention @a -Mode Configured} 'pointer changed|payload changed'
    Check ($f.State.Recycles -eq 0 -and (Get-ChildItem (Join-Path $f.Output 'batches')).Count -eq 6) "$case was not stopped before removal"
}
$f=Fixture;$a=$f.Args;$f.State.Linked=$true;Reject {Invoke-PreparedBatchRetention @a} 'linked';Check ($f.State.Recycles -eq 0) 'Shortcut caused deletion'
$f=Fixture;$a=$f.Args;$f.State.WrongParent=$true;Reject {Invoke-PreparedBatchRetention @a} 'parent mismatch'
$f=Fixture;$a=$f.Args
$lock=[IO.File]::Open((Join-Path $f.Output '.publication.lock'),'OpenOrCreate','ReadWrite','None')
try{Reject {Invoke-PreparedBatchRetention @a} 'being used|utilis|Open'}finally{$lock.Dispose()}
# A second publication cannot acquire the same server lock while the retention reader is active.
$f.State.Hook={
    $blocked=$false;try{$handle=[IO.File]::Open((Join-Path $f.Output '.publication.lock'),'OpenOrCreate','ReadWrite','None');$handle.Dispose()}catch{$blocked=$true}
    if(-not $blocked){throw 'Concurrent publication acquired the retention lock.'}
    $denied=$false;try{$h=[IO.File]::Open($f.Registry,'Open','Write','ReadWrite');$h.Dispose()}catch{$denied=$true}
    if(-not $denied){throw 'Consumer registry was writable during retention.'}
}.GetNewClosure()
$null=Invoke-PreparedBatchRetention @a;Check $true 'Concurrent publisher and registry writer excluded'
$f=Fixture;SetAuto $f;$a=$f.Args
& $module {
    $script:originalLockCheck=${function:Assert-RetentionLock};$script:lockChecks=0
    function script:Assert-RetentionLock([IO.FileStream]$Stream){
        $script:lockChecks++
        if($script:lockChecks -eq 3){$Stream.Dispose()}
        & $script:originalLockCheck $Stream
    }
}
try {Reject {Invoke-PreparedBatchRetention @a -Mode Configured} 'lock was lost|disposed|ferm'}
finally {& $module {Set-Item Function:script:Assert-RetentionLock $script:originalLockCheck}}
Check ($f.State.Recycles -eq 0) 'Lost publication handle did not stop before remote mutation'
& $module {param($root) try{Get-RetentionChild $root '../outside';throw 'Traversal accepted'}catch{if($_.Exception.Message -eq 'Traversal accepted'){throw}}} $f.Output
Check $true 'Root traversal rejected'
$outside=Join-Path $f.Root 'outside';New-Item -ItemType Directory $outside|Out-Null
$link=Join-Path $f.Output 'linked';New-Item -ItemType Junction -Path $link -Target $outside|Out-Null
Reject {& $module {param($root) Get-RetentionChild $root 'linked/file'} $f.Output} 'linked'
$f=Fixture;$a=$f.Args
$r=Invoke-PreparedBatchRetention @a -ExpectedPublishedBatchId '20200101T000000000Z-aaaaaaaa'
Check ($r.Status -ceq 'Blocked' -and $f.State.Recycles -eq 0) 'Stale transfer receipt did not block retention'
$f=Fixture;$a=$f.Args
$loadedHash=& $module {$script:RetentionModuleSHA256}
& $module {$script:RetentionModuleSHA256='F'*64}
try {
    $r=Invoke-PreparedBatchRetention @a
    Check ($r.Status -ceq 'Blocked' -and $f.State.Recycles -eq 0) 'An older loaded implementation was allowed under a different disk qualification'
} finally {& $module {param($hash) $script:RetentionModuleSHA256=$hash} $loadedHash}
# Code defaults: SharePoint keeps current+previous, the server keeps its full history, one SharePoint removal per publication.
$f=Fixture;SetAuto $f;SetPolicy $f @{ServerKeepBatches=$null;MaxSharePointRemovalsPerRun=$null};$a=$f.Args
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($r.Status -ceq 'Completed' -and $f.State.Recycles -eq 1 -and $f.State.Events -ccontains ('Recycle:batches/'+$f.Ids[1])) 'Default policy did not remove exactly the oldest eligible SharePoint batch'
Check ($r.Plan.RetentionPolicy.SharePointKeepBatches -eq 2 -and $r.Plan.RetentionPolicy.ServerKeepBatches -eq 0 -and $r.Plan.RetentionPolicy.MaxSharePointRemovalsPerRun -eq 1) 'Plan does not expose the default policy'
Check (@($r.Plan.Rows|Where-Object {$_.Storage -ceq 'SharePoint' -and $_.BatchId -ceq $f.Ids[3] -and $_.State -ceq 'Deferred'}).Count -eq 1) 'Second SharePoint candidate was not deferred'
Check (@($r.Plan.Rows|Where-Object {$_.Storage -ceq 'Local' -and $_.State -cin @('Eligible','Candidate','Deferred')}).Count -eq 0) 'Full server history produced a local action'
Check (@($r.Plan.Rows|Where-Object {$_.Storage -ceq 'Local' -and $_.State -ceq 'Retained'}).Count -eq 2) 'Unprotected server batches were not reported as retained history'
Check (@(Get-ChildItem (Join-Path $f.Output 'batches') -Directory).Count -eq 6) 'Default policy removed a server batch'
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($f.State.Recycles -eq 2 -and -not (Test-Path (Join-Path $f.Remote ('batches/'+$f.Ids[3])))) 'Deferred SharePoint batch was not retired by the next publication'
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($r.Status -ceq 'Completed' -and $f.State.Recycles -eq 2 -and @(Get-ChildItem (Join-Path $f.Remote 'batches') -Directory).Count -eq 4) 'SharePoint did not settle on current, previous and consumer pins'
Check (@(Get-ChildItem (Join-Path $f.Output 'batches') -Directory).Count -eq 6) 'Server history changed under the default policy'
# A longer server history protects only server copies.
$f=Fixture;SetAuto $f;SetPolicy $f @{ServerKeepBatches=4};$a=$f.Args
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($r.Status -ceq 'Completed' -and $f.State.Recycles -eq 2) 'Server history count changed SharePoint retirement'
Check (-not (Test-Path (Join-Path $f.Output ('batches/'+$f.Ids[1]))) -and (Test-Path (Join-Path $f.Output ('batches/'+$f.Ids[3])))) 'Server keep count was not applied to the server copies only'
Check (@($r.Plan.MinimumRetainedBatches|Where-Object Storage -CEQ 'Local')[0].BatchIds.Count -eq 4) 'Server minimum does not follow ServerKeepBatches'
foreach($case in @(@{ServerKeepBatches=1},@{ServerKeepBatches='4'},@{SharePointKeepBatches=1},@{MaxSharePointRemovalsPerRun=0},@{SharePointKeepBatches=3;ServerKeepBatches=2})){
    $f=Fixture;SetPolicy $f $case;$a=$f.Args
    $r=Invoke-PreparedBatchRetention @a
    Check ($r.Status -ceq 'Blocked' -and @($r.Plan.BlockedReasons|Where-Object {$_ -match 'PreparedRetention\.'}).Count -eq 1) "Invalid retention policy was not blocked: $(($case.Keys|ForEach-Object {$_+'='+$case[$_]}) -join ',')"
}
# Journals: required shared root for Auto/Apply, outside DATA, absolute; legacy work journals still block.
foreach($case in 'Missing','InsideData','Relative'){
    $f=Fixture;SetAuto $f;$a=$f.Args
    $value=switch($case){Missing {$null};InsideData {Join-Path $f.Root 'DATA/journals'};Relative {'journals'}}
    SetPolicy $f @{JournalRootPath=$value}
    Reject {Invoke-PreparedBatchRetention @a -Mode Configured} 'JournalRootPath (is required|must be outside DATA|must be an absolute)'
    Check ($f.State.Recycles -eq 0 -and $f.State.Archives -eq 0) "$case journal root reached removal"
}
$f=Fixture;$shared=Join-Path $f.Root 'shared-journals';SetAuto $f;SetPolicy $f @{JournalRootPath=$shared};$a=$f.Args
$legacy=Join-Path $f.Work ('retention/'+[guid]::NewGuid().ToString('N'));New-Item -ItemType Directory $legacy|Out-Null
Json (Join-Path $legacy 'journal.json.txt') @{SchemaVersion=1;Status='PartialOrBlocked';Events=@(@{Storage='SharePoint';BatchId=$f.Ids[1];State='Uncertain'})}
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($r.JournalPath.StartsWith($shared+'\',[StringComparison]::OrdinalIgnoreCase) -and $r.PlanPath.StartsWith($shared+'\',[StringComparison]::OrdinalIgnoreCase)) 'Plan/journal were not written under the shared journal root'
Check (@($r.Plan.Rows|Where-Object {$_.BatchId -ceq $f.Ids[1] -and $_.Reason -match 'Unresolved prior'}).Count -eq 2) 'Legacy work-root incident no longer blocks its batch'
Check (-not (Test-Path (Join-Path $f.Work 'retention-scratch/*'))) 'Run read-back folder was left behind'
Check ((Get-Content $r.JournalPath -Raw|ConvertFrom-Json).Machine -ceq [Environment]::MachineName) 'Shared journal lacks its writer machine'
# Regression for the 2026-10-22 incident: a pre-send Invalid URI failure leaves the folder intact and is NotRemoved, not Uncertain.
$f=Fixture;SetAuto $f;$a=$f.Args
$f.State.RecycleResult=@{Outcome='RequestFailed';HttpStatus=$null;MayHaveReachedService=$false;Error='Graph DELETE failed with System.UriFormatException: Invalid URI: The format of the URI could not be determined.'}
Reject {Invoke-PreparedBatchRetention @a -Mode Configured} 'was not removed \(Graph DELETE failed with System\.UriFormatException'
$journal=LastJournal $f.Work;$journalEvent=@($journal.Events)[-1]
Check ($journalEvent.State -ceq 'NotRemoved' -and $journalEvent.Outcome -ceq 'NotRemoved' -and $journalEvent.RequestUri -match '/items/' -and $journal.Status -ceq 'PartialOrBlocked') 'Invalid URI failure was not journaled as NotRemoved with its validated URI'
Check ((Test-Path (Join-Path $f.Remote ('batches/'+$f.Ids[1]))) -and @(Get-ChildItem (Join-Path $f.Output 'batches') -Directory).Count -eq 6) 'Invalid URI failure changed data'
$f.State.RecycleResult=$null
$r=Invoke-PreparedBatchRetention @a -Mode Configured
Check ($r.Status -ceq 'Completed' -and -not (Test-Path (Join-Path $f.Remote ('batches/'+$f.Ids[1])))) 'A confirmed NotRemoved outcome wrongly required incident reconciliation'
# URI/ID validation happens before RemovalRequested; a refusal never reaches the DELETE.
$f=Fixture;SetAuto $f;$a=$f.Args;$f.State.PrepareFailure='Recycle refused before any request: invalid Graph item URI'
Reject {Invoke-PreparedBatchRetention @a -Mode Configured} 'refused before any request'
$journalEvent=@((LastJournal $f.Work).Events)[-1]
Check ($f.State.Recycles -eq 0 -and $journalEvent.State -ceq 'Blocked' -and $null -eq $journalEvent.RequestUri) 'Refused removal URI was journaled as an intent or sent'
# Explicit HTTP outcomes.
$outcomes=@(
    @{Name='404Absent';Result=@{Outcome='NotFound';HttpStatus=404;MayHaveReachedService=$false;Error=$null};Remove=$true;State='Recycled';Outcome='AlreadyAbsent';Pattern=$null},
    @{Name='404Present';Result=@{Outcome='NotFound';HttpStatus=404;MayHaveReachedService=$false;Error=$null};Remove=$false;State='NotRemoved';Outcome='NotRemoved';Pattern='HTTP 404'},
    @{Name='412';Result=@{Outcome='PreconditionFailed';HttpStatus=412;MayHaveReachedService=$false;Error=$null};Remove=$false;State='NotRemoved';Outcome='NotRemoved';Pattern='HTTP 412'},
    @{Name='500Present';Result=@{Outcome='HttpError';HttpStatus=500;MayHaveReachedService=$true;Error='Graph DELETE returned HTTP 500; only 204 confirms removal.'};Remove=$false;State='Uncertain';Outcome=$null;Pattern='outcome is uncertain'},
    @{Name='TimeoutAbsent';Result=@{Outcome='RequestFailed';HttpStatus=$null;MayHaveReachedService=$true;Error='timeout'};Remove=$true;State='Recycled';Outcome='AbsentAfterError';Pattern=$null},
    @{Name='204Present';Result=@{Outcome='Recycled';HttpStatus=204;MayHaveReachedService=$false;Error=$null};Remove=$false;State='Uncertain';Outcome=$null;Pattern='HTTP 204 but the folder is still visible'}
)
foreach($case in $outcomes){
    $f=Fixture;SetAuto $f;SetPolicy $f @{MaxSharePointRemovalsPerRun=1};$a=$f.Args
    $f.State.RecycleResult=$case.Result;$f.State.RemoveOnResult=$case.Remove
    if($case.Pattern){Reject {Invoke-PreparedBatchRetention @a -Mode Configured} $case.Pattern}else{$r=Invoke-PreparedBatchRetention @a -Mode Configured -WarningAction SilentlyContinue;Check ($r.Status -ceq 'Completed') "$($case.Name) did not complete"}
    $journalEvent=@((LastJournal $f.Work).Events | Where-Object Storage -CEQ 'SharePoint')[0]
    Check ($journalEvent.State -ceq $case.State -and $journalEvent.Outcome -ceq $case.Outcome -and $journalEvent.HttpStatus -eq $case.Result.HttpStatus) "$($case.Name) journaled $($journalEvent.State)/$($journalEvent.Outcome) instead of $($case.State)/$($case.Outcome)"
    $incidents=@(& $module {param($roots) Get-RetentionUnresolvedIncidents $roots} @($f.Work))
    Check (($incidents.Count -gt 0) -eq ($case.State -ceq 'Uncertain')) "$($case.Name) incident classification is wrong"
}
# Exercise the shared template merge function without tenant initialization or runtime files.
$tenantPath=Join-Path $PSScriptRoot '../../SmartM365/Config/SmartM365-TenantContext.ps1'
$tokens=$null;$errors=$null;$tenantAst=[Management.Automation.Language.Parser]::ParseFile($tenantPath,[ref]$tokens,[ref]$errors)
$merge=$tenantAst.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Add-SmartM365MissingJsonTemplateProperties'},$true)
. ([scriptblock]::Create($merge.Extent.Text))
$template=Get-Content (Join-Path $PSScriptRoot '../../SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.local.json.template') -Raw | ConvertFrom-Json
$syntheticConfig=[pscustomobject]@{PreparedRetention=[pscustomobject]@{Mode='Auto';ConsumersPath='synthetic-custom';Qualification=[pscustomobject]@{Status='Qualified'}}}
$added=@(Add-SmartM365MissingJsonTemplateProperties -Target $syntheticConfig -Template $template)
Check ($added.Count -gt 0 -and $syntheticConfig.PreparedRetention.PlanMaxAgeHours -eq 24) 'Missing runtime parameters were not added'
Check ($syntheticConfig.PreparedRetention.Mode -ceq 'Auto' -and $syntheticConfig.PreparedRetention.ConsumersPath -ceq 'synthetic-custom' -and $syntheticConfig.PreparedRetention.Qualification.Status -ceq 'Qualified') 'Template merge overwrote explicit local settings'
}

# Exercise the real Graph adapter with a local SDK stub. No Graph module or network call.
# The real cmdlet call shape is exercised separately by Test-PreparedRetentionGraphTransport.ps1.
$global:RetentionGraphStub=@{Events=[Collections.Generic.List[object]]::new();Tenant='synthetic-tenant';ContextChanges=0;DeleteStatus=204;DeleteThrow=$null;Pagination='Normal';EmptyPermissions=$false}
function global:Connect-SmartM365GraphAppOnly {param($AppId,$TenantId,$Thumbprint,$Purpose) $true}
function global:Get-MgContext {[pscustomobject]@{TenantId=$global:RetentionGraphStub.Tenant;ClientId='synthetic-app';AuthType='AppOnly'}}
function global:Set-MgRequestContext {param($MaxRetry) $global:RetentionGraphStub.ContextChanges++}
function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param($Method,$Uri,$OutputType,$Headers,$OutputFilePath,[switch]$SkipHttpErrorCheck,$StatusCodeVariable)
    $global:RetentionGraphStub.Events.Add(@{Method=$Method;Uri=[string]$Uri;Headers=$Headers;OutputType=$OutputType;SkipHttpErrorCheck=[bool]$SkipHttpErrorCheck;StatusCodeVariable=$StatusCodeVariable})
    if($Method -ceq 'DELETE'){
        if($global:RetentionGraphStub.DeleteThrow){throw $global:RetentionGraphStub.DeleteThrow}
        # Like the cmdlet, set the variable in the caller's (module) session state.
        if($StatusCodeVariable){$PSCmdlet.SessionState.PSVariable.Set($StatusCodeVariable,$global:RetentionGraphStub.DeleteStatus)}
        return
    }
    if($OutputFilePath){[IO.File]::WriteAllText($OutputFilePath,'synthetic bytes');return}
    switch -Regex ([string]$Uri){
        '/permissions$' {
            if($global:RetentionGraphStub.EmptyPermissions){return [pscustomobject]@{value=@()}}
            return [pscustomobject]@{value=@([pscustomobject]@{id='permission';roles=@('write')})}
        }
        '/children$' {
            if($global:RetentionGraphStub.Pagination -ceq 'Cycle'){return [pscustomobject]@{value=@();'@odata.nextLink'=[string]$Uri}}
            return [pscustomobject]@{value=@([pscustomobject]@{id='batch-id';name='20261001T000000000Z-00000001';folder=@{};parentReference=@{id='batches';driveId='drive'};size=0;eTag='folder-etag'})}
        }
        '/drives$' {return [pscustomobject]@{value=@([pscustomobject]@{id='b!drive';name='Documents'})}}
        ':/batches$' {return [pscustomobject]@{id='batches';folder=[pscustomobject]@{childCount=1};parentReference=@{id='root';driveId='b!drive'}}}
        '/root:/' {return [pscustomobject]@{id='root';folder=[pscustomobject]@{childCount=2};parentReference=@{id='drive-root';driveId='b!drive'}}}
        '/sites/' {return [pscustomobject]@{id='site'}}
        default {throw [Net.Http.HttpRequestException]::new('Synthetic missing item',$null,[Net.HttpStatusCode]::NotFound)}
    }
}
try {
    $f=Fixture
    $transport=New-PreparedRetentionGraphTransport -Connection @{AppId='synthetic-app';TenantId='synthetic-tenant';Thumbprint='synthetic';SiteHostname='example.sharepoint.com';SitePath='/sites/Synthetic';LibraryDisplayName='Documents'} -CloudRoot 'Example/DATA/DATA-POWERBI' -OutputRoot $f.Output -WorkRoot $f.Work -UploadFile {throw 'Unexpected real upload'} -DownloadFile {throw 'Unexpected real download'}
    $target=& $transport.Describe
    Check ($target.RootId -ceq 'root' -and $target.BatchesId -ceq 'batches' -and $target.AccessEvidenceSHA256 -match '^[A-F0-9]{64}$') 'Graph target/access resolution failed'
    $items=@(& $transport.List 'batches');Check ($items.Count -eq 1 -and $items[0].ParentId -ceq 'batches') 'Graph item adapter failed'
    $prepared=& $transport.PrepareRecycle 'batch-id' '"{6A5B},3"'
    Check ($prepared.Uri -ceq 'https://graph.microsoft.com/v1.0/drives/b%21drive/items/batch-id') 'Validated recycle URI is not the exact item URI'
    $result=& $transport.Recycle $prepared
    $request=$global:RetentionGraphStub.Events[-1]
    Check ($result.Outcome -ceq 'Recycled' -and $result.HttpStatus -eq 204 -and $request.Method -ceq 'DELETE' -and $request.Uri -ceq $prepared.Uri -and $request.Headers['If-Match'] -ceq '"{6A5B},3"') 'Graph recycle lost its validated conditional request'
    Check ($request.OutputType -ceq 'PSObject' -and $request.SkipHttpErrorCheck -and $request.StatusCodeVariable) 'DELETE does not use -SkipHttpErrorCheck/-StatusCodeVariable'
    Check ($global:RetentionGraphStub.ContextChanges -eq 0) 'DELETE changed the SDK request context (SDK 2.30 Invalid URI trigger)'
    foreach($status in @(@{Code=404;Outcome='NotFound';Reached=$false},@{Code=412;Outcome='PreconditionFailed';Reached=$false},@{Code=429;Outcome='HttpError';Reached=$false},@{Code=503;Outcome='HttpError';Reached=$false},@{Code=500;Outcome='HttpError';Reached=$true},@{Code=202;Outcome='HttpError';Reached=$true})){
        $global:RetentionGraphStub.DeleteStatus=$status.Code
        $result=& $transport.Recycle $prepared
        Check ($result.Outcome -ceq $status.Outcome -and $result.HttpStatus -eq $status.Code -and $result.MayHaveReachedService -eq $status.Reached) "HTTP $($status.Code) was not classified explicitly"
    }
    $global:RetentionGraphStub.DeleteThrow=[UriFormatException]::new('Invalid URI: The format of the URI could not be determined.')
    $result=& $transport.Recycle $prepared
    Check ($result.Outcome -ceq 'RequestFailed' -and -not $result.MayHaveReachedService -and $result.Error -match 'System\.UriFormatException: Invalid URI') 'Pre-send Invalid URI failure was not reported explicitly as not sent'
    $global:RetentionGraphStub.DeleteThrow=[Threading.Tasks.TaskCanceledException]::new('Synthetic timeout')
    $result=& $transport.Recycle $prepared
    Check ($result.Outcome -ceq 'RequestFailed' -and $result.MayHaveReachedService) 'Timeout was not kept uncertain'
    $global:RetentionGraphStub.DeleteThrow=$null;$global:RetentionGraphStub.DeleteStatus=204
    foreach($id in '','a/b','x%2Fy','id with space','..','id?x=1'){Reject {& $transport.PrepareRecycle $id 'etag'} 'refused before any request'}
    foreach($etag in '',"a`r`nb",('e'*513)){Reject {& $transport.PrepareRecycle 'batch-id' $etag} 'eTag .* missing or malformed'}
    $deletes=@($global:RetentionGraphStub.Events | Where-Object Method -CEQ 'DELETE').Count
    Reject {& $transport.Recycle ([pscustomobject]@{Id='batch-id';ETag='etag';Uri='https://graph.microsoft.com/v1.0/drives/other/items/batch-id'})} 'validated request'
    Check (@($global:RetentionGraphStub.Events | Where-Object Method -CEQ 'DELETE').Count -eq $deletes) 'An unvalidated recycle request reached Graph'
    Check (& $transport.IsAbsent 'missing') 'Graph 404 adapter failed'
    $global:RetentionGraphStub.EmptyPermissions=$true;$emptyTarget=& $transport.Describe
    Check ($emptyTarget.AccessEvidenceSHA256 -match '^[A-F0-9]{64}$' -and $emptyTarget.AccessEvidenceSHA256 -cne $target.AccessEvidenceSHA256) 'A valid empty permissions listing must remain reviewable and fingerprinted'
    $global:RetentionGraphStub.Pagination='Cycle';Reject {& $transport.List 'batches'} 'Cyclic Graph pagination'
    $global:RetentionGraphStub.Tenant='foreign';Reject {& $transport.List 'batches'} 'context changed'
    Reject {& $transport.Recycle $prepared} 'context changed before recycle'
    $global:RetentionGraphStub.Tenant='synthetic-tenant'
    # A cmdlet without -SkipHttpErrorCheck fails before any DELETE, with an explicit message.
    $deletes=@($global:RetentionGraphStub.Events | Where-Object Method -CEQ 'DELETE').Count
    function global:Invoke-MgGraphRequest {[CmdletBinding()] param($Method,$Uri,$OutputType,$Headers,$StatusCodeVariable) $global:RetentionGraphStub.Events.Add(@{Method=$Method})}
    Reject {& $transport.PrepareRecycle 'batch-id' 'etag'} 'lacks -SkipHttpErrorCheck'
    Check (@($global:RetentionGraphStub.Events | Where-Object Method -CEQ 'DELETE').Count -eq $deletes) 'Missing SDK parameter reached DELETE'
    # Static guard: the module never changes the SDK request context nor reads a raw HttpResponseMessage.
    $astTokens=$null;$astErrors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($modulePath,[ref]$astTokens,[ref]$astErrors)
    $commands=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.CommandAst]},$true) | ForEach-Object {$_.GetCommandName()})
    $strings=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.StringConstantExpressionAst]},$true) | ForEach-Object Value)
    Check ($commands -notcontains 'Set-MgRequestContext' -and $strings -notcontains 'HttpResponseMessage') 'Module still calls Set-MgRequestContext or requests HttpResponseMessage'
} finally {
    foreach($name in 'Connect-SmartM365GraphAppOnly','Get-MgContext','Set-MgRequestContext','Invoke-MgGraphRequest'){if(Test-Path ('Function:global:'+$name)){Remove-Item ('Function:global:'+$name)}}
    Remove-Variable RetentionGraphStub -Scope Global
}
Write-Host "PASS: $script:checks synthetic retention checks. No tenant, network, production data or runtime configuration accessed. Fixtures: $base"


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA5y2DPTSNXuyFu
# uIGJxCrk+NWwngdZ82Vsb9E7+z2QmqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIw+8KtpkRUnhiWJFexYuiZ4bXT+Nu6uuIC3KJIQiWU2MA0GCSqG
# SIb3DQEBAQUABIIBgH0KAEQRu6K4qi50FVjmsC3+tPJBlRgt40jWJ54Di+5y1BpV
# PxHAg3F19iS2sEttBaqXly48lPp9jY49ZAk+r+gzKiSnvRMYbV0bugsBQZSvhgXy
# az4fn5TOjhoyM8Vui+QmHroTax4L+Vpm+hk/80LHiPLQTQNn0sE9MBIGqZWHH91F
# b33Waa1hrL96rUwT7j/aqqkaUAEaNZ8RlhR40dtJsNJ+9LcAyuveyYLjwWMJuACm
# LZAhqJNj0sTV6/HBfh6Jy45VNcSTuPKN9jBkwbWxq36toL03RrM2tLl6nHnjzyVD
# VLZxg2wqTenc0yU8VWdlXsW7YItxUC48CBR0i9aApzl3DNShFppIkGXkzuSHygex
# aXFvB1LhsYUiOdLWGIF02mPbA5mctzpFHCCfUXHL6p5j1f8xPk36dXVnUi80SK/L
# xCvw+OGdhx1QBufXcpQWvGtWRHrGUT+aqD0Ik8AwV02yhT9afKJeNQHYoybiCWbZ
# bXoZ1n/q9xYjqgLxSqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjAwMDQ4
# NDRaMC8GCSqGSIb3DQEJBDEiBCCvrwQc8L9FY4HgUHh2A2/56PHDvoFW3oEQ/IZN
# ECKkLjANBgkqhkiG9w0BAQEFAASCAgAFquZy8+kze/1hS0wn7dgQKATH88C2LLIp
# aRhETepTVqW5JJSv7a1ccuIj4l496//lOli1yXpRd0PXztxBiAE0IlYgt0yVpJ+h
# 1O6QxOUmDr78vLOsyDh0KhO2JbyL+eXqeSm5i+TcVEAfy4eRZzHYX5ulhoWq/qYi
# KR5ziTm29gojsx7QSC59T0LDlFyRrmLKXe1sqZXPT9oJG5v4cdwUEoer2qS/JpVG
# frANHJ9ghyUj8UeFupid/3KQdadXclGmUSyX28yRIlgExqnnzKZAC/JaR2bWgJCW
# eG148OvpYyZelL3sb+8ysmsEYimGk3/1jeFqOr81Pb1lNB27ZF0iil5qgPjCqyM1
# O+70Sv6zSxCAiBhGqpuCrtpRhtRxbNsrQVXeNINSQCryeweDHuyrkynih7KPE7LA
# 5tPDQhyFq0UF3WQauKRlCgCODH0XactqQKMCoCpfQ9eKSIHJaGWSr6WtKZKOuUF2
# 6y7WI75RKcMXc9s13swSAXQX6CWK99lVEC4+gAps4hjrQSqs3lK0x1C5RrOIw4Gg
# DdJTKwSx/GSCfFphQiY6fDZxT4Rm425LEagpEo9OmWdVTSvzv3Oje3VSTIk+8cPV
# EGnn/g07E2LgsuMFL2o1/b30rAR97MlMIU0B3hhp81KV/FNTdeWOEW4uuykKJhao
# 27sKRndeNw==
# SIG # End signature block
