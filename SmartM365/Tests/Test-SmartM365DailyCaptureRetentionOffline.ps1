#requires -Version 7.2
<#
.SYNOPSIS
Synthetic daily capture, preservation and conditional SharePoint deletion tests.
.VERSION
1.0.4
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core globals are saved and restored in finally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs','',Justification='Short internal fixture/assertion names, not exported commands.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Offline transport mocks retain the actual API parameter signatures.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$module=Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.92' -Force -PassThru
$temporary=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DailyCaptureOffline-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
$savedUpload=Get-Variable EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue
$savedPaths=Get-Variable csvGeneratedPaths -Scope Global -ErrorAction SilentlyContinue
$savedIdentity=@{}
foreach($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId')){
    $savedIdentity[$name]=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
}
try {
    $result=& $module {
        param($Temporary)
        $script:Checks=0;$script:Logs=[Collections.Generic.List[string]]::new()
        $script:RemoteCalls=[Collections.Generic.List[object]]::new();$script:RemoteRoot=Join-Path $Temporary 'fake-remote'
        $script:BadUpload=$false;$script:MutateOnDelete=$false;$script:UploadLog=[Collections.Generic.List[string]]::new()
        $script:Limited=$false;$script:ProbeProtectedPath='';$script:ProtectedWriteBlocked=$false;$script:Downloads=0
        function script:WriteLog {param($Message,$Level) $script:Logs.Add("$Level|$Message")}
        function script:Check {param([bool]$Condition,[string]$Message) if(-not $Condition){throw "FAIL: $Message"};$script:Checks++}
        function script:Reject {param([scriptblock]$Action,[string]$Pattern)
            $caught=$null;try{& $Action}catch{$caught=$_.Exception.Message}
            Check ($caught -and $caught -match $Pattern) "expected rejection $Pattern, got $caught"
        }
        function script:Test-SmartM365MaxItemsMode {return $script:Limited}
        function script:RemotePath {
            param($Path)
            $sha=[Security.Cryptography.SHA256]::Create()
            try {$key=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Path))).Replace('-','')} finally {$sha.Dispose()}
            return Join-Path $script:RemoteRoot ($key+'.bin')
        }
        function script:Invoke-SmartM365SharePointCsvUpload {
            param($LocalFilePath,[switch]$EnsureParentFolders)
            Check ([bool]$EnsureParentFolders) 'capture uploads create remote parent folders explicitly'
            New-Item -ItemType Directory -Path $script:RemoteRoot -Force | Out-Null
            $script:UploadLog.Add($LocalFilePath)
            [IO.File]::Copy($LocalFilePath,(RemotePath $LocalFilePath),$true)
            if($script:BadUpload -and $LocalFilePath -like '*.csv'){[IO.File]::AppendAllText((RemotePath $LocalFilePath),'corrupt')}
            return [pscustomobject]@{LocalFilePath=$LocalFilePath;SharePointPath='synthetic'}
        }
        function script:Get-SmartM365DailyCaptureRemoteItem {
            param($LocalPath)
            $remotePath=RemotePath $LocalPath
            if(-not(Test-Path -LiteralPath $remotePath)){throw 'synthetic missing remote file'}
            return @{Uri=$remotePath;ETag=(Get-FileHash -LiteralPath $remotePath -Algorithm SHA256).Hash;Bytes=(Get-Item -LiteralPath $remotePath).Length}
        }
        function script:Invoke-SmartM365GraphFileDownloadWithRetry {
            param($Uri,$OutputFilePath,$Operation)
            $null=$Operation;$script:Downloads++;[IO.File]::Copy(($Uri -replace '/content$',''),$OutputFilePath,$true)
        }
        function script:Invoke-SmartM365GraphRestWithRetry {
            param($Method,$Uri,$AdditionalHeaders,$Operation)
            $null=$Operation
            if($Method -ne 'DELETE'){throw 'Unexpected network method in offline test.'}
            if($script:ProbeProtectedPath){
                try{[IO.File]::AppendAllText($script:ProbeProtectedPath,'unexpected concurrent write')}
                catch{$script:ProtectedWriteBlocked=$true}
            }
            if($script:MutateOnDelete){[IO.File]::AppendAllText($Uri,'concurrent mutation')}
            if($AdditionalHeaders['If-Match'] -cne (Get-FileHash -LiteralPath $Uri -Algorithm SHA256).Hash){throw 'synthetic HTTP 412'}
            $script:RemoteCalls.Add(@{Method=$Method;Uri=$Uri;IfMatch=$AdditionalHeaders['If-Match']})
            Remove-Item -LiteralPath $Uri -Force
        }
        function script:New-Fixture {param([string]$Name)
            $base=Join-Path $Temporary $Name
            $all=Join-Path $base 'all';$last=Join-Path $base 'last'
            New-Item -ItemType Directory -Path $all,$last -Force | Out-Null
            foreach($relative in @('WeeklyHistory/2026-W42/One.csv','Corrections/One.csv','unindexed_20200101_010101.csv','manifest.json.txt','receipt.json.txt')){
                $p=Join-Path $all $relative;New-Item -ItemType Directory -Path (Split-Path $p -Parent) -Force | Out-Null;[IO.File]::WriteAllText($p,'protected')
            }
            return @{All=$all;Last=$last;Base=$base}
        }
        function script:New-Run {
            param($Fixture,[datetimeoffset]$At,[string]$Label,[string[]]$Names=@('One.csv','Two.csv'),[switch]$ExistingTimestamp,[string]$Status='Completed')
            $id=[guid]::NewGuid().ToString('N')
            $context=@{Producer='SmartM365-IntuneConfigurations-Inventory.ps1';TenantKey='synthetic';OrganizationKey='example';EnvironmentKey='test';TenantId='synthetic-tenant';RunId=$id;SourceRoot=$Fixture.Last}
            $script:SmartM365CmdbSourceContext=$context
            Enable-SmartM365DailyCaptureRetention -OutputRoot $Fixture.All
            $global:csvGeneratedPaths=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $files=@()
            foreach($name in $Names){
                $path=Join-Path $Fixture.Last $name
                [IO.File]::WriteAllText($path,"TenantKey,Value`r`nsynthetic,$Label`r`n")
                $files+=@{File=$name;SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
                [void]$global:csvGeneratedPaths.Add($path)
                if($ExistingTimestamp){
                    $dated=Join-Path $Fixture.All ([IO.Path]::GetFileNameWithoutExtension($name)+'_'+$At.ToString('yyyyMMdd_HHmmss')+'.csv')
                    [IO.File]::Copy($path,$dated,$false);[void]$global:csvGeneratedPaths.Add($dated)
                }
            }
            $proof=[pscustomobject]@{Status=$Status;RunId=$id;CompletedAtUtc=$At.ToString('o');Files=$files;RequiredFiles=$Names;ScopeParameters=@{Mode='All'};ScriptVersion='1.0.0'}
            return @{Context=$context;Proof=$proof;At=$At;Fixture=$Fixture}
        }
        function script:Finish {param($Run) Invoke-SmartM365DailyCaptureRetention -Context $Run.Context -Proof $Run.Proof -ReferenceUtc $Run.At -Confirm:$false}
        function script:Records {param($Fixture) return @(Get-ChildItem -LiteralPath $Fixture.All -Recurse -File -Filter '*.json.txt' | Where-Object {$_.DirectoryName -match 'DailyCaptureRecords'} | ForEach-Object {[IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json})}
        function script:Payloads {param($Fixture) return @(Get-ChildItem -LiteralPath (Join-Path $Fixture.All 'DailyCaptures') -Recurse -File -ErrorAction SilentlyContinue)}
        $global:EnableSharePointUpload=$false
        $f=New-Fixture 'calendar'
        # Ten days, two full acquisitions per day. Only the last run of each of the final seven days survives.
        foreach($day in 1..10){foreach($hour in @(10,15)){
            $at=[datetimeoffset]::new(2026,10,$day,$hour,0,0,[timespan]::Zero)
            $run=New-Run $f $at "day${day}hour$hour";Finish $run
            $all=Records $f
            Check ($all.Count -eq (($day-1)*2+[int]($hour -eq 15)+1)) 'every metadata record remains available'
            $retained=@($all | Where-Object {@($_.RemovedFiles).Count -eq 0})
            Check ($retained.Count -eq [math]::Min($day,7)) 'one complete retained capture per calendar day'
        }}
        $records=Records $f
        Check ((Payloads $f).Count -eq 14) 'seven coherent two-file captures remain'
        foreach($rec in @($records|Where-Object {@($_.RemovedFiles).Count -eq 0})){
            Check ($rec.Day -ge '2026-10-04') 'calendar cutoff includes today and six preceding days'
            foreach($file in $rec.Files){Check ([IO.File]::ReadAllText((Join-Path $f.All $file.Relative)) -match 'hour15') 'both files come from final successful daily run'}
        }
        foreach($path in @('WeeklyHistory/2026-W42/One.csv','Corrections/One.csv','unindexed_20200101_010101.csv','manifest.json.txt','receipt.json.txt')){Check ([IO.File]::ReadAllText((Join-Path $f.All $path)) -ceq 'protected') 'protected and unindexed evidence unchanged'}
        Check ([IO.File]::ReadAllText((Join-Path $f.Last 'One.csv')) -match 'day10hour15') 'current source remains unchanged by retention'
        Check ((Get-SmartM365DailyCaptureDay ([datetimeoffset]'2026-10-19T22:30:00Z')) -ceq '2026-10-20') 'completion date uses Paris, including UTC midnight offset'
        Check ((Get-SmartM365DailyCaptureDay ([datetimeoffset]'2026-10-25T23:30:00Z')) -ceq '2026-10-26') 'Paris winter-time transition'
        $before=(Payloads $f).Count
        $partial=New-Run $f ([datetimeoffset]'2026-10-10T16:00:00Z') 'partial' @('One.csv')
        Reject {Finish $partial} 'fewer exports'
        Check ((Payloads $f).Count -eq $before) 'a smaller selection cannot replace full same-day evidence'
        $failed=New-Run $f ([datetimeoffset]'2026-10-10T17:00:00Z') 'failed' -Status Failed
        Finish $failed;Check ((Payloads $f).Count -eq $before) 'failed run never purges'
        $whatif=New-Run $f ([datetimeoffset]'2026-10-10T18:00:00Z') 'whatif'
        Invoke-SmartM365DailyCaptureRetention -Context $whatif.Context -Proof $whatif.Proof -ReferenceUtc $whatif.At -WhatIf
        Check ((Records $f).Count -eq 20 -and (Payloads $f).Count -eq $before) 'WhatIf creates neither payloads nor metadata'
        $script:SmartM365DailyCapture=$null
        $script:SmartM365CmdbSourceContext=$null
        Enable-SmartM365DailyCaptureRetention -OutputRoot (Join-Path $Temporary 'readonly')
        Check (-not(Test-Path (Join-Path $Temporary 'readonly'))) 'read-only/no-receipt mode creates no output'
        $sample=New-Fixture 'sample';$script:Limited=$true
        $sampleRun=New-Run $sample ([datetimeoffset]'2026-10-19T08:00:00Z') 'sample'
        Finish $sampleRun
        Check (-not(Test-Path (Join-Path $sample.All 'DailyCaptureRecords'))) 'bounded sampling cannot initialize daily capture retention'
        $script:Limited=$false
        $script:SmartM365CmdbSourceContext=@{Producer='Unrelated-Producer.ps1'}
        Reject {Enable-SmartM365DailyCaptureRetention -OutputRoot $sample.All} 'not enabled'

        $f=New-Fixture 'existing-exports'
        $run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') 'existing' -ExistingTimestamp
        Finish $run
        Check (-not(Test-Path (Join-Path $f.All 'DailyCaptures'))) 'existing matching current-run timestamped payloads are reused'
        $old=Records $f
        Check ($old[0].Files[0].Relative -match '_20261019_080000.csv$') 'metadata records exact existing publication paths'
        $lock=[IO.File]::Open((Join-Path $f.All $old[0].Files[0].Relative),'Open','ReadWrite','None')
        try{
            $run=New-Run $f ([datetimeoffset]'2026-10-19T09:00:00Z') 'locked'
            Reject {Finish $run} 'being used|another process|utilis|acc.s|access'
            Check (Test-Path (Join-Path $f.All $old[0].Files[0].Relative)) 'locked obsolete payload is preserved'
        }finally{$lock.Dispose()}

        $f=New-Fixture 'changed-retained';$run=New-Run $f ([datetimeoffset]'2026-10-18T08:00:00Z') 'old';Finish $run
        $old=Records $f;[IO.File]::AppendAllText((Join-Path $f.All $old[0].Files[0].Relative),'changed')
        $run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') 'new'
        Reject {Finish $run} 'Retained capture is incomplete or changed'
        Check (Test-Path (Join-Path $f.All $old[0].Files[1].Relative)) 'changed retained evidence prevents every deletion'

        # A relocated workspace keeps its tenant-relative records; another tenant identity is still refused.
        $savedUploadState=$global:EnableSharePointUpload;$global:EnableSharePointUpload=$false
        $f=New-Fixture 'relocated-old/Tenants/synthetic/DATA-ALL';$run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') 'old-share';Finish $run
        $movedBase=Join-Path $Temporary 'relocated-new/Tenants/synthetic/DATA-ALL'
        New-Item -ItemType Directory -Path (Split-Path $movedBase -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $f.Base -Destination $movedBase -Recurse
        $moved=@{All=(Join-Path $movedBase 'all');Last=(Join-Path $movedBase 'last');Base=$movedBase}
        Check ((Records $moved)[0].OutputRoot -ceq [IO.Path]::GetFullPath($f.All)) 'relocated record keeps the former output root'
        $run=New-Run $moved ([datetimeoffset]'2026-10-20T08:00:00Z') 'new-share';Finish $run
        Check ((Records $moved).Count -eq 2 -and (Payloads $moved).Count -eq 4) 'relocated history remains eligible for retention'
        $foreignBase=Join-Path $Temporary 'relocated-foreign/Tenants/other/DATA-ALL'
        New-Item -ItemType Directory -Path (Split-Path $foreignBase -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $movedBase -Destination $foreignBase -Recurse
        $foreign=@{All=(Join-Path $foreignBase 'all');Last=(Join-Path $foreignBase 'last');Base=$foreignBase}
        $run=New-Run $foreign ([datetimeoffset]'2026-10-21T08:00:00Z') 'foreign'
        Reject {Finish $run} 'Invalid capture metadata'
        Check ((Payloads $foreign).Count -eq 4) 'foreign tenant records block retention before copying or deleting'
        $script:SmartM365DailyCapture=$null;$script:SmartM365CmdbSourceContext=$null
        $global:EnableSharePointUpload=$savedUploadState

        $f=New-Fixture 'metadata';$run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') 'first';Finish $run
        $meta=Get-ChildItem -LiteralPath (Join-Path $f.All 'DailyCaptureRecords') -Recurse -File -Filter '*.json.txt' | Select-Object -First 1
        $value=[IO.File]::ReadAllText($meta.FullName)|ConvertFrom-Json;$value.Files[0].Relative='../outside.csv'
        [IO.File]::WriteAllText($meta.FullName,($value|ConvertTo-Json -Depth 12))
        $run=New-Run $f ([datetimeoffset]'2026-10-19T09:00:00Z') 'second'
        Reject {Finish $run} 'Unsafe capture path'
        Check ((Payloads $f).Count -eq 2) 'traversal metadata blocks retention before copying or deleting'
        foreach($path in @('WeeklyHistory/One.csv','DATA-LAST/One.csv','Corrections/One.csv','../One.csv','C:\outside.csv','manifest.csv','receipt.csv')){
            Reject {Resolve-SmartM365DailyCapturePath $f.All $path} 'Unsafe|eligible'
        }
        $link=Join-Path $f.All 'linked';$linkTarget=Join-Path $Temporary 'synthetic-link-target'
        New-Item -ItemType Directory -Path $linkTarget -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $linkTarget 'One.csv'),'protected synthetic target')
        New-Item -ItemType Junction -Path $link -Target $linkTarget | Out-Null
        try {Reject {Resolve-SmartM365DailyCapturePath $f.All 'linked/One.csv'} 'Linked capture'}
        finally {Remove-Item -LiteralPath $link -Force}
        Check ([IO.File]::ReadAllText((Join-Path $linkTarget 'One.csv')) -ceq 'protected synthetic target') 'junction rejection does not alter its target'

        $f=New-Fixture 'linked-output';$linkedRoot=Join-Path $f.Base 'linked-all'
        New-Item -ItemType Junction -Path $linkedRoot -Target $f.All | Out-Null
        try {
            $f.All=$linkedRoot
            $run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') 'linked'
            Check ([bool]$script:SmartM365DailyCapture.Disabled) 'linked output disables retention without failing collection'
            Check (Test-SmartM365DailyCapturePath $linkedRoot) 'linked output still blocks legacy cleanup'
            Finish $run
            Check (-not(Test-Path (Join-Path $linkedRoot 'DailyCaptureRecords'))) 'linked output receives no retention writes'
        } finally {Remove-Item -LiteralPath $linkedRoot -Force}

        # Actual remote adapter verification runs against local fake remote bytes; no network is reachable.
        $global:EnableSharePointUpload=$true
        $f=New-Fixture 'remote'
        $run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') 'first';Finish $run
        $old=Records $f
        $run=New-Run $f ([datetimeoffset]'2026-10-19T09:00:00Z') 'second'
        $script:ProbeProtectedPath=Join-Path $f.All ('DailyCaptures/'+$run.Proof.RunId+'/One.csv')
        Finish $run
        Check $script:ProtectedWriteBlocked 'retained payload cannot be modified concurrently during obsolete-file deletion'
        Check ([IO.File]::ReadAllText($script:ProbeProtectedPath) -notmatch 'unexpected concurrent write') 'retained read lease preserves the verified bytes'
        $script:ProbeProtectedPath=''
        Check ($script:RemoteCalls.Count -eq 2) 'remote obsolete files removed only after all read-back checks'
        foreach($call in $script:RemoteCalls){Check ($call.IfMatch -match '^[a-fA-F0-9]{64}$') 'conditional deletion carries the verified remote ETag'}
        Check ((Payloads $f).Count -eq 2) 'local files removed only after matching remote deletion'
        Check ((Records $f).Count -eq 2) 'remote retention also preserves metadata records'

        $f=New-Fixture 'remote-cas'
        $run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') 'first';Finish $run;$old=Records $f
        $script:MutateOnDelete=$true
        $run=New-Run $f ([datetimeoffset]'2026-10-19T09:00:00Z') 'second'
        Reject {Finish $run} 'HTTP 412'
        Check (Test-Path (Join-Path $f.All $old[0].Files[0].Relative)) 'remote concurrent change prevents local deletion'
        $script:MutateOnDelete=$false
        $f=New-Fixture 'remote-upload-failure';$script:BadUpload=$true
        $run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') 'first'
        Reject {Finish $run} 'size differs|SHA256 differs'
        Check ((Records $f).Count -eq 0) 'failed read-back cannot create completed capture proof'
        $script:BadUpload=$false

        # A capture this run already uploaded (Publish-SmartM365Csv) is verified by SHA-256 read-back, not resent.
        foreach($case in @('Verified','StaleRemote','FailedOutcome')){
            $f=New-Fixture "remote-uploaded-$case"
            $run=New-Run $f ([datetimeoffset]'2026-10-19T08:00:00Z') $case -ExistingTimestamp
            $dated=@($global:csvGeneratedPaths | Where-Object {$_ -match '_\d{8}_\d{6}\.csv$'})
            Check ($dated.Count -eq 2) 'fixture publishes two timestamped captures'
            $global:SmartM365SharePointUploadedFiles=[Collections.ArrayList]::new()
            foreach($path in $dated){
                $null=Invoke-SmartM365SharePointCsvUpload -LocalFilePath $path -EnsureParentFolders
                [void]$global:SmartM365SharePointUploadedFiles.Add([pscustomobject]@{LocalFilePath=$path;Size=(Get-Item -LiteralPath $path).Length})
            }
            if($case -eq 'StaleRemote'){[IO.File]::AppendAllText((RemotePath $dated[0]),'changed')}
            if($case -eq 'FailedOutcome'){Set-SmartM365SharePointUploadOutcome -Path $dated[0] -Failed}
            $script:UploadLog.Clear();$script:Logs.Clear()
            Finish $run
            $resent=@($script:UploadLog | Where-Object {$_ -like '*.csv'})
            $skipped=@($script:Logs | Where-Object {$_ -match 'upload skipped'})
            if($case -eq 'Verified'){Check ($resent.Count -eq 0 -and $skipped.Count -eq 2) 'verified captures are not uploaded twice'}
            else{Check ($resent.Count -eq 1 -and $resent[0] -ieq $dated[0] -and $skipped.Count -eq 1) "$case capture is uploaded again"}
            if($case -eq 'StaleRemote'){Check (@($script:Logs | Where-Object {$_ -match '^WARNING\|.*uploading again'}).Count -eq 1) 'stale remote copy is reported'}
            if($case -eq 'FailedOutcome'){Set-SmartM365SharePointUploadOutcome -Path $dated[0]}
            Check ((Records $f).Count -eq 1) "$case capture proof is completed"
            foreach($path in $dated){Check ((Get-FileHash -LiteralPath (RemotePath $path)).Hash -ceq (Get-FileHash -LiteralPath $path).Hash) "$case remote bytes match"}
        }
        $global:SmartM365SharePointUploadedFiles=$null;$global:EnableSharePointUpload=$false
        # Verification memory (1.0.4): retained files fully verified within seven days get a light check only.
        $global:EnableSharePointUpload=$true
        $f=New-Fixture 'light-check'
        foreach($day in 19..21){$run=New-Run $f ([datetimeoffset]"2026-10-$($day)T08:00:00Z") "day$day";Finish $run}
        $memoryPath=Join-Path $f.All 'DailyCaptureRecords/SmartM365-IntuneConfigurations-Inventory/.retention.verification.json'
        Check ((Test-Path -LiteralPath $memoryPath) -and @((Records $f)).Count -eq 3) 'the verification memory sits beside the records and is not read as a record'
        $script:Downloads=0;$script:Logs.Clear()
        $run=New-Run $f ([datetimeoffset]'2026-10-22T08:00:00Z') 'day22';Finish $run
        Check ($script:Downloads -eq 3) "retained files already verified are not downloaded again, only the two new files and the record ($($script:Downloads) downloads)"
        Check (@($script:Logs | Where-Object {$_ -match 'checking 8 retained file\(s\) of 4 day\(s\)'}).Count -eq 1 -and @($script:Logs | Where-Object {$_ -match '8 light check\(s\).*0 full verification\(s\)'}).Count -eq 1) 'progress and light-check summary are logged'
        $memory=[IO.File]::ReadAllText($memoryPath)|ConvertFrom-Json
        Check (@($memory.Files.PSObject.Properties).Count -eq 8) 'the memory keeps the retained files only'
        # A rewritten local file (same bytes, new write time) gets a full check that passes.
        $retained=@(Records $f | Sort-Object Day)[0]
        (Get-Item -LiteralPath (Join-Path $f.All $retained.Files[0].Relative)).LastWriteTimeUtc=[datetime]'2026-10-20T00:00:00Z'
        $script:Downloads=0;$script:Logs.Clear()
        $run=New-Run $f ([datetimeoffset]'2026-10-23T08:00:00Z') 'day23';Finish $run
        Check ($script:Downloads -eq 4 -and @($script:Logs | Where-Object {$_ -match 'full check 1/10 \(2026-10-19'}).Count -eq 1) 'a changed write time falls back to the full verification with its progress line'
        # A SharePoint copy changed with the same size gets a new ETag: the full check refuses it.
        $retained=@(Records $f | Sort-Object Day)[1]
        $remoteCopy=RemotePath (Join-Path $f.All $retained.Files[0].Relative)
        $bytes=[IO.File]::ReadAllBytes($remoteCopy);$bytes[$bytes.Length-3]=[byte][char]'X';[IO.File]::WriteAllBytes($remoteCopy,$bytes)
        $run=New-Run $f ([datetimeoffset]'2026-10-24T08:00:00Z') 'day24'
        Reject {Finish $run} 'SharePoint capture SHA256 differs'
        [IO.File]::Copy((Join-Path $f.All $retained.Files[0].Relative),$remoteCopy,$true)
        # A full check older than seven days, or an unreadable memory, means a full verification again.
        $memory=[IO.File]::ReadAllText($memoryPath)|ConvertFrom-Json
        foreach($property in @($memory.Files.PSObject.Properties)){$property.Value.VerifiedAtUtcTicks=([datetimeoffset]'2026-10-10T00:00:00Z').UtcTicks}
        [IO.File]::WriteAllText($memoryPath,($memory|ConvertTo-Json -Depth 6))
        $script:Downloads=0;$script:Logs.Clear()
        $run=New-Run $f ([datetimeoffset]'2026-10-24T09:00:00Z') 'day24b';Finish $run
        Check (@($script:Logs | Where-Object {$_ -match ': 12 retained file\(s\) checked .* 2 light check\(s\).* 10 full verification\(s\)'}).Count -eq 1 -and $script:Downloads -ge 12) "a full check older than seven days is repeated, only this run's new files are light ($(@($script:Logs) -join ' // '))"
        [IO.File]::WriteAllText($memoryPath,'{ not json')
        $script:Logs.Clear()
        $run=New-Run $f ([datetimeoffset]'2026-10-24T10:00:00Z') 'day24c';Finish $run
        Check (@($script:Logs | Where-Object {$_ -match '^WARNING\|Daily capture verification memory unreadable'}).Count -eq 1 -and @($script:Logs | Where-Object {$_ -match ' 2 light check\(s\).* 10 full verification\(s\)'}).Count -eq 1) 'an unreadable memory gives full verifications, never a failure'
        $global:EnableSharePointUpload=$false

        # Exercise the actual receipt finalizer: the source lock encloses the capture handoff.
        $script:SmartM365CmdbSourceContext=$null;$script:SmartM365DailyCapture=$null
        $global:SmartM365TenantKey='synthetic';$global:SmartM365OrganizationKey='example'
        $global:SmartM365EnvironmentKey='test';$global:SmartM365TenantId='synthetic-tenant'
        $f=New-Fixture 'real-receipt'
        $producer=Join-Path (Split-Path $script:SmartM365CmdbRegistryPath -Parent) '../../SmartInventory/M365Inventory/WorkplaceScope/SmartM365-IntuneConfigurations-Inventory.ps1'
        $producer=[IO.Path]::GetFullPath($producer)
        Start-SmartM365SourceReceipt -ScriptPath $producer -SourceRootPath $f.Last
        Enable-SmartM365DailyCaptureRetention -OutputRoot $f.All
        $expected=@($script:SmartM365CmdbSourceContext.ExpectedFiles)
        foreach($name in $expected){
            $path=Join-Path $f.Last $name
            [IO.File]::WriteAllText($path,"TenantKey,OrganizationKey,EnvironmentKey,TenantId,Value`r`nsynthetic,example,test,synthetic-tenant,1`r`n")
            Register-SmartM365SourceCsv $path
        }
        $proofPath=Complete-SmartM365CmdbSourceReceipt -Status Success
        $proof=[IO.File]::ReadAllText($proofPath)|ConvertFrom-Json
        Check ($proof.Status -eq 'Completed' -and $proof.Files.Count -eq 7) 'source receipt completes through the actual Core finalizer'
        Check ((Records $f).Count -eq 1 -and (Payloads $f).Count -eq 7) 'receipt finalizer creates one complete seven-file capture'
        Check ($null -eq $script:SmartM365CmdbSourceContext) 'source lock context released after capture completion'
        $oldPath=Join-Path $f.All 'One_20200101_010101.csv';[IO.File]::WriteAllText($oldPath,'protected old unindexed payload')
        RemoveOldFiles -Path $f.All -Filter '*.csv' -KeepCount 0
        Remove-SmartM365TimestampedFilesOlderThan -FolderPath $f.All -FilePattern '*.csv' -RetentionDays 1
        $r=Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath $oldPath -Enabled $true
        Check (Test-Path $oldPath) 'per-file counted and age cleanups are bypassed for this opted-in root'
        Check ($r.Deleted -eq 0) 'per-file SharePoint cleanup is bypassed before any API call'
        Start-SmartM365SourceReceipt -ScriptPath $producer -SourceRootPath $f.Last
        Enable-SmartM365DailyCaptureRetention -OutputRoot $f.All
        $null=Complete-SmartM365CmdbSourceReceipt -Status Failed
        Check ((Records $f).Count -eq 1 -and (Payloads $f).Count -eq 7) 'failed actual receipt completion cannot purge the last complete capture'
        # Static wiring checks: selected producers only, correct parent for Usage and version floors.
        $script:SmartM365DailyCapture=$null
        return [pscustomobject]@{Checks=$script:Checks;NetworkCalls=0;DataRootAccess=0;TemporaryRoot=$Temporary}
    } $temporary
    $sources=@(
        'M365Inventory/Usage/SmartM365-M365UserActivity-Inventory.ps1','M365Inventory/Usage/SmartM365-CopilotUsage-Inventory.ps1',
        'M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.ps1','M365Inventory/WorkplaceScope/SmartM365-IntuneConfigurations-Inventory.ps1',
        'M365Inventory/WorkplaceScope/SmartM365-IntuneManagedApps-Inventory.ps1','M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1',
        'M365Inventory/Licensing/SmartM365-LicensePricing-Inventory.ps1','M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'
    )
    foreach($relative in $sources){
        $text=[IO.File]::ReadAllText((Join-Path $root ('SmartInventory/'+$relative)))
        if($text -notmatch 'Enable-(?:Core)?SmartM365DailyCaptureRetention -OutputRoot \$(?:OutputPath|OutputFolder|script:Runtime.OutputPath)'){throw "Missing daily capture wiring: $relative"}
        # Daily capture needs Core 1.0.92; a later floor (staged publication, 1.0.93) also qualifies.
        if($text -notmatch "-MinimumVersion '1\.0\.(?:9[2-9]|[1-9]\d\d)'"){throw "Missing Core module floor: $relative"}
        $result.Checks++
    }
    "PASS: Daily capture retention $($result.Checks) checks; synthetic local files and mocked SharePoint only."
} finally {
    if($savedUpload){Set-Variable EnableSharePointUpload -Scope Global -Value $savedUpload.Value}else{Remove-Variable EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue}
    if($savedPaths){Set-Variable csvGeneratedPaths -Scope Global -Value $savedPaths.Value}else{Remove-Variable csvGeneratedPaths -Scope Global -ErrorAction SilentlyContinue}
    foreach($name in $savedIdentity.Keys){
        if($savedIdentity[$name]){Set-Variable -Name $name -Scope Global -Value $savedIdentity[$name].Value}else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}
    }
    Remove-Module $module -Force -ErrorAction SilentlyContinue
    $full=[IO.Path]::GetFullPath($temporary)
    if(-not $full.StartsWith([IO.Path]::GetTempPath(),[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($full) -notlike 'SmartM365-DailyCaptureOffline-*'){throw 'Unsafe synthetic cleanup root.'}
    Remove-Item -LiteralPath $full -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCz0VOXnnd1Xqvm
# enjgP2YFn9MaBlR1Dc2mzFx8lY2IzqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEVGO1MLSmxUWuOOwNcK8/HFDpIvNLtee9uafYVFZWetMA0GCSqG
# SIb3DQEBAQUABIIBgCnm9wbVQGTRXad6DxI/BIrrXqGxXN0UJuV4GmadgtaONUDa
# /saufNj6Mrr2nWz7PzeYTYhZPgIIU1kQ+svbfqhRZRkI4ej2RJZEE3JW46BvzD2L
# eJUOcWaI5a7xLu12/l7bLIUkJwSnTW9ulynpDoWIM+QFbgQ+tefUxnO+U7ih098t
# kKDCWlorlC02aGYKNq4bI/B7ux1zROsZgHqzF/gpYT3fAcx/+qYxm7xovvfbX985
# gc1cPVkvvFyz+918ILekxnUpv8192m9t30lQ9Fol0JBPf8NEdEeS1EpAN3jHQFts
# JOooZ/z58Mvcang/XATfLhHPfyWO+RJAgJeRqL5SVMgDPw2A9ibu2fekzGFLQR0q
# /XvmDpQefDjOITEZi/h163Tinf0INnAurRhj2ae+EjP+GbGH5LNn4+wCWxyPUMCP
# OrYZAHsPQl/enhblZQgDgwf9xJlS5SD2816LCHKgDdOVN7LUW08zcAhlvDis3ySg
# QZzkHUKpxKbQhzJvOaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwODMw
# MjlaMC8GCSqGSIb3DQEJBDEiBCA0/WqK6PcmH+DheHAVk8OIgXAUaDlYiz/NV3sO
# kXpyWDANBgkqhkiG9w0BAQEFAASCAgBC/JQGekSNj72SpnLHXE4oQEy1UuMqKAvA
# fVxHoHngS2oAwVevd3fnlMidQQIhnwbXG6weZizDvOggXyUjKlVVZjl0wiefb/8l
# SSaZLKtCb1aU8s6YdAoxYhlHQPEiusnVu9H34hNMIRVNSJgGt9gYBqVLu8wZzWcR
# Orv8etj9sFJBoW8swExLte33LqsYbjejBk71468W2B1JXx8J9ISIQDQX25G40gYB
# Fvw5Mx1MD9osa6dzJrW5t+ziw1qbIs5VKKJoA7E2t+V0lt2pv0dz75HHHG2mFyZ7
# gyNfYQJYsDTscwMt4nttmviw6cAduvmB3cptBNeEPJM0ru1Mr0penrIKNPqyIUPF
# 2ZUazW3wCWYcx0SikOb/kx2elcwL0LDT8pmAvFniPkfVls+Hg7uK2zSe4eLtYvlA
# Na7q5mC5jW5FcFR+/Fo0sceZlqSWoP2wLBqqwIXqn1YT+O6RCGjPq932E8B/min2
# tCchJWnGXRF7mu7TeM5Ndtb829WEERMkzKWOvn1KuiOsoAunc4bOVenDgvmdJ90e
# vGSUp49z1dN2RxOdL8YA81zGqHexyA8bHwLyoJEtCfeoaCHFlEy8DfPmT3dL/T7g
# /+9iGL4PbGntz1+r7dhq6jLHgRAcODekgWAUADwyuqW0Z+23B14t+82Ja6OEsIxk
# i4e3cT1irQ==
# SIG # End signature block
