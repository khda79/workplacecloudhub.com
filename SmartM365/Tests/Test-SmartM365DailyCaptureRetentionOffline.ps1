#requires -Version 7.2
<#
.SYNOPSIS
Synthetic daily capture, preservation and conditional SharePoint deletion tests.
.VERSION
1.0.0
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
        $script:BadUpload=$false;$script:MutateOnDelete=$false
        $script:Limited=$false;$script:ProbeProtectedPath='';$script:ProtectedWriteBlocked=$false
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
            $null=$Operation;[IO.File]::Copy(($Uri -replace '/content$',''),$OutputFilePath,$true)
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
        $script:BadUpload=$false;$global:EnableSharePointUpload=$false

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
        if($text -notmatch "-MinimumVersion '1.0.92'"){throw "Missing Core module floor: $relative"}
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
