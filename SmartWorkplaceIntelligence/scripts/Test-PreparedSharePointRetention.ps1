<#
.SYNOPSIS
    Synthetic retention regression tests. File-backed fake SharePoint; no tenant or network.
.VERSION
    1.0.0
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
    Json $config @{PreparedRetention=@{Mode='Preview';PlanMaxAgeHours=24;Target=$target;Qualification=$q}}
    $state=@{Remote=$remote;Target=$target;Events=[Collections.Generic.List[string]]::new();FailArchive=$false;FailArchiveAt=0;Archives=0;FailDownload=$false;FailRecycleAt=0;Recycles=0;Uncertain=$false;ChangeAfterArchive=$false;ChangePayloadAfterArchive=$false;Linked=$false;WrongParent=$false;Recycled=@{};Hook=$null}
    $describe={if($state.Hook){& $state.Hook};[pscustomobject]$state.Target}.GetNewClosure()
    $list={param($parent)
        $folder=if($parent -eq 'root'){$state.Remote}else{Join-Path $state.Remote $parent}
        if(-not(Test-Path $folder)){throw 'Synthetic missing remote parent'}
        foreach($item in Get-ChildItem -LiteralPath $folder -Force){
            $relative=[IO.Path]::GetRelativePath($state.Remote,$item.FullName).Replace('\','/')
            [pscustomobject]@{Id=$relative;Name=$item.Name;ParentId=$(if($state.WrongParent){'wrong'}else{$parent});Kind=$(if($item.PSIsContainer){'Folder'}else{'File'});Linked=$state.Linked;Bytes=$(if($item.PSIsContainer){0}else{$item.Length});ETag=$(if($item.PSIsContainer){'stable-folder-etag'}else{(Get-FileHash $item.FullName).Hash})}
        }
    }.GetNewClosure()
    $download={param($id,$destination)
        if($state.FailDownload){return $false}
        Copy-Item -LiteralPath (Join-Path $state.Remote $id) -Destination $destination
        $true
    }.GetNewClosure()
    $archive={param($path,$relative)
        $state.Archives++;$state.Events.Add('Archive:'+$relative)
        if($state.FailArchive -or ($state.FailArchiveAt -and $state.Archives -eq $state.FailArchiveAt)){return $false}
        $destination=Join-Path $state.Remote $relative
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $path -Destination $destination
        if($state.ChangeAfterArchive){Add-Content (Join-Path $state.Remote 'current.json.txt') ' '}
        if($state.ChangePayloadAfterArchive){Add-Content (Join-Path $state.Remote ('batches/'+$ids[1]+'/Example.csv')) 'changed'}
        (Get-FileHash $path).Hash -eq (Get-FileHash $destination).Hash
    }.GetNewClosure()
    $recycle={param($id,$etag)
        $state.Recycles++;$state.Events.Add('Recycle:'+$id)
        if($state.FailRecycleAt -and $state.Recycles -eq $state.FailRecycleAt){throw 'Synthetic recycle failure'}
        if($state.Uncertain){throw 'Synthetic timeout after request'}
        $source=Join-Path $state.Remote $id;$destination=Join-Path $root ('recycle/'+[IO.Path]::GetFileName($source))
        New-Item -ItemType Directory -Path (Split-Path $destination -Parent) -Force | Out-Null
        Move-Item -LiteralPath $source -Destination $destination
        $state.Recycled[$id]=$true
        'Recycled'
    }.GetNewClosure()
    $absent={param($id) -not(Test-Path -LiteralPath (Join-Path $state.Remote $id))}.GetNewClosure()
    [pscustomobject]@{Root=$root;Output=$output;Remote=$remote;Work=$work;Config=$config;Registry=$registry;Ids=$ids;State=$state;Args=@{OutputRoot=$output;TenantKey='synthetic';ConfigurationPath=$config;ConsumersPath=$registry;WorkRoot=$work;Transport=@{Describe=$describe;List=$list;Download=$download;Archive=$archive;Recycle=$recycle;IsAbsent=$absent}}}
}
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

if(-not $AdapterOnly){
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
$global:RetentionGraphStub=@{Events=[Collections.Generic.List[object]]::new();Tenant='synthetic-tenant';Retry=3;DeleteStatus=204;Pagination='Normal';EmptyPermissions=$false}
function global:Connect-SmartM365GraphAppOnly {param($AppId,$TenantId,$Thumbprint,$Purpose) $true}
function global:Get-MgContext {[pscustomobject]@{TenantId=$global:RetentionGraphStub.Tenant;ClientId='synthetic-app';AuthType='AppOnly'}}
function global:Get-MgRequestContext {[pscustomobject]@{MaxRetry=$global:RetentionGraphStub.Retry}}
function global:Set-MgRequestContext {param($MaxRetry) $global:RetentionGraphStub.Retry=$MaxRetry}
function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param($Method,$Uri,$OutputType,$Headers,$OutputFilePath,$StatusCodeVariable)
    $global:RetentionGraphStub.Events.Add(@{Method=$Method;Uri=[string]$Uri;Headers=$Headers;Retry=$global:RetentionGraphStub.Retry})
    if($Method -ceq 'DELETE'){
        return [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]$global:RetentionGraphStub.DeleteStatus)
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
        '/drives$' {return [pscustomobject]@{value=@([pscustomobject]@{id='drive';name='Documents'})}}
        ':/batches$' {return [pscustomobject]@{id='batches';folder=[pscustomobject]@{childCount=1};parentReference=@{id='root';driveId='drive'}}}
        '/root:/' {return [pscustomobject]@{id='root';folder=[pscustomobject]@{childCount=2};parentReference=@{id='drive-root';driveId='drive'}}}
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
    $null=& $transport.Recycle 'batch-id' 'folder-etag'
    $request=$global:RetentionGraphStub.Events[-1]
    Check ($request.Method -ceq 'DELETE' -and $request.Headers['If-Match'] -ceq 'folder-etag' -and $request.Retry -eq 0) 'Graph recycle lost conditional request/retry guard'
    Check ($global:RetentionGraphStub.Retry -eq 3) 'SDK retry setting was not restored'
    Check (& $transport.IsAbsent 'missing') 'Graph 404 adapter failed'
    $global:RetentionGraphStub.DeleteStatus=200
    Reject {& $transport.Recycle 'batch-id' 'folder-etag'} 'HTTP 204'
    Check ($global:RetentionGraphStub.Retry -eq 3) 'SDK retry restoration failed after error'
    $global:RetentionGraphStub.EmptyPermissions=$true;$emptyTarget=& $transport.Describe
    Check ($emptyTarget.AccessEvidenceSHA256 -match '^[A-F0-9]{64}$' -and $emptyTarget.AccessEvidenceSHA256 -cne $target.AccessEvidenceSHA256) 'A valid empty permissions listing must remain reviewable and fingerprinted'
    $global:RetentionGraphStub.Pagination='Cycle';Reject {& $transport.List 'batches'} 'Cyclic Graph pagination'
    $global:RetentionGraphStub.Tenant='foreign';Reject {& $transport.List 'batches'} 'context changed'
} finally {
    foreach($name in 'Connect-SmartM365GraphAppOnly','Get-MgContext','Get-MgRequestContext','Set-MgRequestContext','Invoke-MgGraphRequest'){Remove-Item ('Function:global:'+$name)}
    Remove-Variable RetentionGraphStub -Scope Global
}
Write-Host "PASS: $script:checks synthetic retention checks. No tenant, network, production data or runtime configuration accessed. Fixtures: $base"

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC1CletTpvLRa5w
# FZ4pQM+Orev8bRa0M5wwDRF5QpW2gqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINfnONaWOFhnY2yUKCnTORdrzDsGwVp1/VvdoJdLkqCeMA0GCSqG
# SIb3DQEBAQUABIIBgKXEUe0dfStZb02S1aotWXoQBUUgyA4A0MDNSZhInMH56ChS
# MadM3txDudJBVlH6E7Svg8gCBX6ij7Wx9/G6F9OJPqjQwzZM2BfNO1LVszMKW4ja
# kJ6nVytf2vN0gHyrdckpcxNEkHqEvvm8WY5RB8lb9BTBhmzuxvV0Xe9tuMLh7SwK
# VSDMi9DPwLEh5Nte7vadSGX4epz7Uc3bQs0M9s/qgW2m3zgccmBL45ZcWBQmf+Wh
# ciHUE02KkLz8o0n1YiH7CTiSOX/daf4WFN+7jvI6jUjQ5hWRbKXZMSvaQoJe7D87
# Gx3fXk01Pkx+NZ9yaQsO1NOFWm87uUQ2FJcYEP4TzNkfdwwWH0DXratshRzE4lPA
# p+SIDMxoJt+qnaA9grVY60sVpgUOVBUcytcavcO696jIUluuwcdBx1TYJoUEqoTp
# BOPmTgtBcKXVFyUJfT1ZooO/J2pH56mjwO17F7ZjOOI/uy77YmhdD2ulbrtmvc/o
# Q6ZfJbvFsMZGVX2UiaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxOTQ3
# MTVaMC8GCSqGSIb3DQEJBDEiBCCHCqO7WnganHw6p+9wt0WULSeZkZ1tvqtvkJwR
# Y5bHdzANBgkqhkiG9w0BAQEFAASCAgBSNropIW/qIma0kQT52g0T5/e64PVSlv+p
# OLXyCmiGvl1JVf1NoaPXZAawJ/ZAn9eXPL0D4hFmfsSLFsx9XSqk5T9XrP/sBIIH
# mRJHzP7msCBTjxSusZiGHtJULTvkxKCju2OmtvVBaO9wIVbI8Lm7MR0ZiZ8uCVWy
# fkP9/hIT5k1AYZ9dpWBlEEGLuyBN5xmJMI5mmQ5JwSPakvvfMK5YzVxCaveIlC7o
# xerqCY/6SiWCZsFbX9fpOg8zShR9h17AIgtxl/BxZf+09SUoCf3TGs5YzynAnD0+
# fjkYIKnBP8H08C2PVM2sr32DWFegDrLx23E8MarvXA8qtH7HqFWAohICGB8/u/Q2
# 9t1ZI50XG6vPvFWMM+UswmIthf2eneEGzS4/V9EBLuA2kfIGfjHgRncLTZXP3REU
# 7+vNpm/7ra7/y6H70hopIvFOChM2sa3QFKHnviC2Oux+mmr70oUIltwV+nQDzZcU
# NR9B6jtEvK4R8oaq0xEknqsIQzd0hmrfX2JKMdSfEnoglTCh7zH1IJSZExB6mEHd
# 2/FVBnPqOlHHRpt2sZ5i5xCqe1uVU3gi3/1Mg7GZ7IXSqAaz2Q9jj0eK4mffQJpO
# etJ3ElPZP/Q+V4ZAhWScazG/U5qsNY45T7E2T3vVzBJOv8ukkrbyTqPX7KMpTRlI
# rrYImUM1qQ==
# SIG # End signature block
