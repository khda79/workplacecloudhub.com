#Requires -Version 5.1
<# .SYNOPSIS Offline certificate budgets, workers, projections and infrastructure continuation. .VERSION 1.1.2 #>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$smartRoot=Split-Path $PSScriptRoot -Parent
$collector=Join-Path $smartRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
$worker=Join-Path (Split-Path $collector) 'SmartM365-Exchange-OnPrem-CertificateWorker.ps1'
. $worker
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$errors)
$reportAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $smartRoot 'SmartInventory/Common/SmartM365.ReportRenderers.ps1'),[ref]$null,[ref]$null)
$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
Check ($errors.Count -eq 0) 'Collector parser failed'
foreach($name in @('Write-Log','Get-ExchangeReadinessInventory','ConvertTo-ReadinessText','Get-ObjectPropertyValue','Add-ExchangeReadinessRow','Add-ExchangeReadinessProperties','Invoke-ExchangeReadinessCollector','Stop-SmartM365InfrastructureCertificateWorker','Add-SmartM365InfrastructureCertificateProperties','Invoke-SmartM365InfrastructureCertificateInventory')){
    $node=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    if (-not $node) {
        $node=$reportAst.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    }
    Check ($null -ne $node) ('Missing production function '+$name)
    . ([scriptblock]::Create($node.Extent.Text))
}
$realLogger=(Get-Command Write-Log).ScriptBlock
function Write-Log {param($Message,$Level='INFO') $script:messages.Add($Level+': '+$Message);& $realLogger -Message $Message -Level $Level}
$script:InfrastructureCertificateWorkerPath=$worker
$smartM365LifecycleRoot=$smartRoot
$fixtureRoot=Join-Path ([IO.Path]::GetTempPath()) ('CertificateBudget-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixtureRoot
$LogFile=Join-Path $fixtureRoot 'synthetic.log'
$realStart=(Get-Command Start-SmartM365CertificateWorker).ScriptBlock
$realRead=(Get-Command Invoke-SmartM365CertificateWorkerRead).ScriptBlock
$realStop=(Get-Command Stop-SmartM365CertificateWorker).ScriptBlock
$hostExe=Join-Path $PSHOME $(if($PSVersionTable.PSEdition -eq 'Desktop'){'powershell.exe'}else{'pwsh.exe'})
function Reset-Fixture([string]$Case='Complete',[string]$Mode='LocalFirst',[object]$Budget=120){
    $script:case=$Case;$script:mode=$Mode;$script:budget=$Budget
    $script:messages=New-Object 'System.Collections.Generic.List[string]'
    $script:ServersAndStorageWarningCount=0;$script:ServersAndStorageErrorCount=0
    $script:starts=New-Object 'System.Collections.Generic.List[object]'
    $script:reads=New-Object 'System.Collections.Generic.List[object]'
    $script:states=New-Object 'System.Collections.Generic.List[object]'
}
function Get-SmartM365EffectiveConfigValue {
    param($Name,$DefaultValue)
    if($Name -eq 'ExchangeConnectionMode'){if($null -eq $script:mode){$DefaultValue}else{$script:mode}}
    elseif($Name -eq 'CertificateTimeoutSecondsPerServer'){if($null -eq $script:budget){$DefaultValue}else{$script:budget}}
    else{$DefaultValue}
}
function New-SyntheticCertificate {
    [pscustomobject]@{Thumbprint=('1'*40);Subject='CN=mail.example.test';Services='IIS, SMTP';Status='Valid';IsSelfSigned=$false;CertificateDomains=@('mail.example.test');NotBefore=[datetime]'2026-01-01';NotAfter=[datetime]'2027-01-01'}
}
try{
    Reset-Fixture
    # Compare native snapshot values with the actual pre-IPC renderer, including enum/zero/false/null/arrays/dates.
    foreach($value in @('IIS, SMTP',[IO.FileAttributes]'ReadOnly, Hidden',0,$false,$null,@(),@('a','b'),[datetime]'2027-01-02')){
        $c=New-SyntheticCertificate;$c.Services=$value
        $snapshot=ConvertTo-SmartM365CertificateSnapshot -Certificate $c -Mode Native
        $wire=[Management.Automation.PSSerializer]::Deserialize([Management.Automation.PSSerializer]::Serialize($snapshot))
        Check ((ConvertTo-ReadinessText $value) -ceq (ConvertTo-ReadinessText $wire.Services)) 'Native value changed during snapshot/IPC'
    }
    $enum=New-SyntheticCertificate
    $enum.Services=[IO.FileAttributes]'ReadOnly, Hidden'
    $enumSnapshot=ConvertTo-SmartM365CertificateSnapshot -Certificate $enum -Mode Remote
    $enumWire=$enumSnapshot|ConvertTo-Json -Depth 8|ConvertFrom-Json
    Check ($enumWire.Services -ceq [string]$enum.Services) 'Remote enum turned into numeric JSON instead of its known text'
    $missing=[pscustomobject]@{Thumbprint=('2'*40);Subject='CN=mail.example.test';DnsNameList=@([pscustomobject]@{Unicode='mail.example.test'});NotBefore='raw-date';NotAfter='raw-date'}
    $snapshot=ConvertTo-SmartM365CertificateSnapshot -Certificate $missing -Mode Remote
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Add-SmartM365InfrastructureCertificateProperties -Rows $rows -InputObject $snapshot -ObjectName synthetic -Mode Remote
    Check (@($rows|Where-Object Value -eq Unknown).Count -eq 3) 'Remote missing metadata became known'
    Check (@($rows|Where-Object {$_.Setting -eq 'CertificateDomains' -and $_.Value -eq 'mail.example.test'}).Count -eq 1) 'Remote DNS fallback lost'
    Check (@($rows|Where-Object {$_.Setting -eq 'NotAfter' -and $_.Value -eq 'raw-date'}).Count -eq 1) 'Remote date reinterpreted'
    $known=[pscustomobject]@{Thumbprint=('3'*40);Subject='s';CertificateDomains=@();Services=0;IsSelfSigned=$false;Status='Valid';NotBefore='x';NotAfter='x'}
    $snapshot=ConvertTo-SmartM365CertificateSnapshot -Certificate $known -Mode Remote
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Add-SmartM365InfrastructureCertificateProperties -Rows $rows -InputObject $snapshot -ObjectName synthetic -Mode Remote
    Check (@($rows|Where-Object {$_.Setting -eq 'Services' -and $_.Value -eq '0'}).Count -eq 1) 'Remote zero lost'
    Check (@($rows|Where-Object {$_.Setting -eq 'IsSelfSigned' -and $_.Value -eq 'False'}).Count -eq 1) 'Remote false lost'
    Check (@($rows|Where-Object {$_.Setting -eq 'CertificateDomains' -and $_.Value -eq '(none)'}).Count -eq 1) 'Remote empty list lost'

    # Real child processes run only synthetic Exchange primitives, using the production worker loop.
    $wrapper=Join-Path $fixtureRoot 'synthetic-worker.ps1'
    $quotedWorker=$worker.Replace("'","''")
    $wrapperCode=@(
        'param([switch]$RunWorker,[string]$WorkerMode,[string]$WorkerRoot)',
        '$savedMode=$WorkerMode;$savedRoot=$WorkerRoot',
        ('. ''{0}''' -f $quotedWorker),
        'function Add-PSSnapin {param($Name)}',
        'function Get-SyntheticNative { [CmdletBinding()]param($Server) if($Server -like ''*HANG*''){Start-Sleep -Seconds 30}; if($Server -like ''*DENY*''){throw ''Synthetic native denied''}; [pscustomobject]@{Thumbprint=(''1''*40);Subject=''CN=synthetic'';Services=''IIS, SMTP'';Status=''Valid'';IsSelfSigned=$false;CertificateDomains=@(''mail.example.test'');NotAfter=''raw-date'';NotBefore=''raw-date''} }',
        'function Get-Command { [CmdletBinding()]param($Name,$CommandType) if($Name -eq ''Microsoft.Exchange.Management.PowerShell.SnapIn\Get-ExchangeCertificate''){return Microsoft.PowerShell.Core\Get-Command Get-SyntheticNative}; Microsoft.PowerShell.Core\Get-Command @PSBoundParameters }',
        'Invoke-SmartM365CertificateChildLoop -Mode $savedMode -Root $savedRoot'
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($wrapper,$wrapperCode,[Text.UTF8Encoding]::new($true))
    $core=Join-Path $fixtureRoot 'Modules/SmartM365.Core'
    $null=New-Item -ItemType Directory -Path $core -Force
    $coreCode=@(
        'function Start-SmartM365ExchangeOnPremSession { param($RequiredCommands)',
        ' $case=Get-Content (Join-Path $Root ''scenario.txt'') -Raw',
        ' Add-Content (Join-Path $Root ''sessions.txt'') ''start''',
        ' if($case -eq ''InitHang''){Start-Sleep -Seconds 30}; if($case -eq ''ImportHang''){Start-Sleep -Seconds 30}; if($case -eq ''InitDenied''){throw ''Synthetic import denied''}',
        ' $module=New-Module -Name SyntheticCertificates -ScriptBlock { function Get-ExchangeCertificate { [CmdletBinding()]param($Server) if($Server -like ''*HANG*''){Start-Sleep -Seconds 30};if($Server -like ''*DENY*''){throw ''Synthetic remote denied''};[pscustomobject]@{Thumbprint=(''2''*40);Subject=''CN=synthetic'';DnsNameList=@(''mail.example.test'');NotAfter=''raw-date'';NotBefore=''raw-date''} } }',
        ' Import-Module $module -Global -Force',
        ' [pscustomobject]@{Session=[pscustomobject]@{State=''Opened''};Module=$module}',
        '}',
        'function Stop-SmartM365ExchangeOnPremSession {param($State) Add-Content (Join-Path $Root ''closed.txt'') ''closed'';Remove-Module $State.Module -Force}'
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText((Join-Path $core 'SmartM365-ExchangeOnPremSession.ps1'),$coreCode,[Text.UTF8Encoding]::new($true))
    Set-Content (Join-Path $fixtureRoot 'scenario.txt') 'Complete' -NoNewline
    $state=$null
    try{
        $state=Start-SmartM365CertificateWorker -Mode Native -Root $fixtureRoot -WorkerPath $wrapper -PowerShellPath $hostExe
        $childId=$state.Process.Id
        $reply=Invoke-SmartM365CertificateWorkerRead $state SYNTHETIC-OK 10
        Check ($reply.Status -eq 'Success' -and @($reply.Items).Count -eq 1 -and $reply.Items[0].Services -eq 'IIS, SMTP') 'Real native worker lost metadata'
    }finally{Stop-SmartM365CertificateWorker $state}
    Check ($null -eq (Get-Process -Id $childId -ErrorAction SilentlyContinue)) 'Native worker orphaned'
    $state=$null
    try{
        $state=Start-SmartM365CertificateWorker -Mode Remote -Root $fixtureRoot -WorkerPath $wrapper -PowerShellPath $hostExe
        $childId=$state.Process.Id
        foreach($server in @('SYNTHETIC-OK1','SYNTHETIC-OK2')){
            $reply=Invoke-SmartM365CertificateWorkerRead $state $server 10
            Check ($reply.Status -eq 'Success' -and $reply.SessionReady) 'Real remote worker failed'
            Check ($null -eq $reply.Items[0].PSObject.Properties['Services']) 'Worker invented remote Services'
        }
        Check (@(Get-Content (Join-Path $fixtureRoot 'sessions.txt')).Count -eq 1) 'Real worker opened more than one session'
    }finally{Stop-SmartM365CertificateWorker $state}
    Check ((Test-Path (Join-Path $fixtureRoot 'closed.txt')) -and $null -eq (Get-Process -Id $childId -ErrorAction SilentlyContinue)) 'Remote finally/session/process cleanup failed'
    foreach($case in @('NativeHang','RemoteHang','InitHang','ImportHang')){
        Set-Content (Join-Path $fixtureRoot 'scenario.txt') $case -NoNewline
        $state=$null;$watch=[Diagnostics.Stopwatch]::StartNew()
        try{
            $state=Start-SmartM365CertificateWorker -Mode $(if($case -eq 'NativeHang'){'Native'}else{'Remote'}) -Root $fixtureRoot -WorkerPath $wrapper -PowerShellPath $hostExe
            $childId=$state.Process.Id
            $reply=Invoke-SmartM365CertificateWorkerRead $state $(if($case -like '*Hang' -and $case -notin @('InitHang','ImportHang')){'SYNTHETIC-HANG'}else{'SYNTHETIC-OK'}) 1
            Check ($reply.Status -eq 'Timeout' -and @($reply.Items).Count -eq 0) ('Hung worker returned success/partial data: '+$case)
        }finally{Stop-SmartM365CertificateWorker $state -Force;$watch.Stop()}
        Check ($watch.Elapsed.TotalSeconds -lt 6 -and $null -eq (Get-Process -Id $childId -ErrorAction SilentlyContinue)) ('Hung worker not bounded/cleaned: '+$case)
    }
    foreach($payload in @('malformed','{"RequestId":"wrong","Status":"Success","Items":[]}')){
        $bad=Join-Path $fixtureRoot 'bad-worker.ps1'
        $badCode='[Console]::Out.WriteLine('''+$payload.Replace("'","''")+''');Start-Sleep -Seconds 30'
        [IO.File]::WriteAllText($bad,$badCode,[Text.UTF8Encoding]::new($true))
        $state=$null
        try{
            $state=Start-SmartM365CertificateWorker -Mode Native -Root $fixtureRoot -WorkerPath $bad -PowerShellPath $hostExe
            $reply=Invoke-SmartM365CertificateWorkerRead $state SYNTHETIC-OK 5
            Check ($reply.Status -eq 'Failed' -and @($reply.Items).Count -eq 0) 'Malformed/truncated response promoted to success'
        }finally{Stop-SmartM365CertificateWorker $state -Force}
    }

    Reset-Fixture Complete LocalFirst 6
    Set-Content (Join-Path $fixtureRoot 'scenario.txt') 'Complete' -NoNewline
    $script:realWorkerStarts=New-Object 'System.Collections.Generic.List[object]'
    function Start-SmartM365CertificateWorker {
        param($Mode,$Root,$WorkerPath)
        $s=& $realStart -Mode $Mode -Root $fixtureRoot -WorkerPath $wrapper -PowerShellPath $hostExe
        $script:realWorkerStarts.Add([pscustomobject]@{Mode=$Mode;Id=$s.Process.Id})
        $s
    }
    $watch=[Diagnostics.Stopwatch]::StartNew()
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-1-HANG','SYNTHETIC-2-HANG','SYNTHETIC-3-OK')
    $watch.Stop()
    Check ($watch.Elapsed.TotalSeconds -lt 15) 'Combined native/remote budget doubled, cleanup unbounded, or next budget not independent'
    Check (@($script:realWorkerStarts|Where-Object Mode -eq Remote).Count -eq 1) 'Remote worker recreated after real timeout'
    Check (@($rows|Where-Object {$_.Setting -eq 'Collection' -and $_.CollectionStatus -eq 'ERROR'}).Count -eq 2) 'Real timeout failed diagnostics missing'
    Check (@($rows|Where-Object {$_.ObjectName -like 'SYNTHETIC-3-OK:*' -and $_.Setting -eq 'Services' -and $_.Value -eq 'IIS, SMTP'}).Count -eq 1) 'Real next native target lost metadata'
    Check (@($rows|Where-Object {$_.Setting -eq 'Collection' -and $_.ErrorMessage -match 'Local: Timeout.*; Remote: Timeout'}).Count -eq 1) 'Real combined timeout messages missing'
    foreach($s in $script:realWorkerStarts){Check ($null -eq (Get-Process -Id $s.Id -ErrorAction SilentlyContinue)) 'Real phase orphaned a child'}

    # The production dispatcher now uses deterministic transport stubs for its complete mode/error matrix.
    function Start-SmartM365CertificateWorker {
        param($Mode,$Root,$WorkerPath)
        $s=[pscustomobject]@{Mode=$Mode;Closed=$false;ReadyLogged=$false}
        $script:starts.Add($s);$script:states.Add($s);$s
    }
    function Stop-SmartM365CertificateWorker {param($State,[switch]$Force) if($null -ne $State){$State.Closed=$true}}
    function Invoke-SmartM365CertificateWorkerRead {
        param($State,$Server,$TimeoutSeconds)
        $script:reads.Add([pscustomobject]@{Mode=$State.Mode;Server=$Server;Limit=$TimeoutSeconds})
        if($script:case -eq 'LocalCancel' -and $State.Mode -eq 'Native'){throw [OperationCanceledException]::new('Synthetic cancellation')}
        if($script:case -eq 'RemoteCancel' -and $State.Mode -eq 'Remote'){throw [OperationCanceledException]::new('Synthetic cancellation')}
        $r=[pscustomobject]@{Status='Success';Items=@((New-SyntheticCertificate));Stage='Read';Message='';FatalSession=$false;SessionReady=($State.Mode -eq 'Remote')}
        if($script:case -eq 'Empty'){$r.Items=@()}
        if($State.Mode -eq 'Native' -and $script:case -ne 'Complete' -and $script:case -ne 'Empty' -and ($Server -ne 'SYNTHETIC-EX03')){
            $r.Status='Failed';$r.Items=@();$r.Message='Synthetic native denied'
            if($script:case -in @('NativeTimeout','BothTimeout')){$r.Status='Timeout';$r.Stage='Transport';$r.Message='Synthetic native timeout'}
        }
        if($State.Mode -eq 'Remote'){
            if($script:case -ne 'Empty'){$r.Items=@([pscustomobject]@{Thumbprint=('2'*40);Subject='s';DnsNameList=@('mail.example.test');NotBefore='raw-date';NotAfter='raw-date'})}
            if($script:case -in @('RemoteDenied','InitDenied','BothTimeout','RemoteTimeout')){
                $r.Items=@();$r.Status='Failed';$r.Message='Synthetic remote denied'
                if($script:case -eq 'InitDenied'){$r.Stage='Initialize';$r.FatalSession=$true;$r.SessionReady=$false}
                if($script:case -in @('BothTimeout','RemoteTimeout')){$r.Status='Timeout';$r.Stage='Transport';$r.FatalSession=$true;$r.Message='Synthetic remote timeout';$r.SessionReady=$false}
            }
        }
        $r
    }
    foreach($mode in @('LocalFirst','LocalSnapin','Remote')){
        foreach($case in @('Complete','Empty','Retry','RemoteDenied','InitDenied','NativeTimeout','BothTimeout','RemoteTimeout','LocalCancel','RemoteCancel')){
            Reset-Fixture $case $mode
            $rows=New-Object 'System.Collections.Generic.List[object]';$caught=$null
            try{Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-EX01','SYNTHETIC-EX02','SYNTHETIC-EX03')}catch{$caught=$_}
            $cancel=($case -eq 'LocalCancel' -and $mode -ne 'Remote') -or ($case -eq 'RemoteCancel' -and $mode -ne 'LocalSnapin')
            Check (($null -ne $caught) -eq $cancel) ('Wrong cancellation/continuation: '+$mode+'/'+$case)
            Check (@($script:states|Where-Object {-not $_.Closed}).Count -eq 0) 'Dispatcher leaked worker on success/error/cancellation'
            Check (@($script:starts|Where-Object Mode -eq Remote).Count -le 1) 'Dispatcher recreated remote session'
            if($cancel){Check (@($rows|Where-Object Setting -eq Collection).Count -eq 0) 'Cancellation converted to Failed';continue}
            $failedTargets=@($rows|Where-Object {$_.Setting -eq 'Collection' -and $_.Value -eq 'Failed' -and $_.CollectionStatus -eq 'ERROR'})
            if($failedTargets.Count){
                Check ($script:ServersAndStorageErrorCount -ge 1) 'Failed certificate targets must increment the general error counter.'
            } else {
                Check ($script:ServersAndStorageErrorCount -eq 0) 'Successful certificate targets were counted as errors.'
            }

            if($mode -eq 'LocalFirst' -and $case -in @('Complete','Empty')){Check (@($script:starts|Where-Object Mode -eq Remote).Count -eq 0) 'Native success opened remote session'}
            if($mode -eq 'LocalSnapin'){Check (@($script:reads|Where-Object Mode -eq Remote).Count -eq 0) 'LocalSnapin used Remote'}
            if($mode -eq 'Remote'){Check (@($script:reads|Where-Object Mode -eq Native).Count -eq 0) 'Remote used native fallback'}
            if($case -in @('BothTimeout','RemoteTimeout','InitDenied') -and $mode -ne 'LocalSnapin'){
                Check (@($script:reads|Where-Object Mode -eq Remote).Count -eq 1) 'Unavailable remote session still used'
                Check (@($rows|Where-Object {$_.Setting -eq 'Collection' -and $_.Value -eq 'Failed' -and $_.CollectionStatus -eq 'ERROR'}).Count -ge 1) 'Timeout diagnostics lost'
                if($mode -eq 'LocalFirst'){Check (@($rows|Where-Object {$_.ObjectName -like 'SYNTHETIC-EX03:*' -and $_.Setting -eq 'Services' -and $_.Value -eq 'IIS, SMTP'}).Count -eq 1) 'Native reads after remote loss failed'}
            }
            if($mode -eq 'LocalFirst' -and $case -in @('RemoteDenied','BothTimeout')){
                Check (@($rows|Where-Object {$_.Setting -eq 'Collection' -and $_.ErrorMessage -match 'Local: .*; Remote: '}).Count -eq 2) 'Combined local/remote diagnostic missing'
            }
            foreach($read in $script:reads){
                $max=if($mode -eq 'LocalFirst' -and $read.Mode -eq 'Native'){60}else{120}
                Check ($read.Limit -gt 0 -and $read.Limit -le $max) 'Per-server budget was doubled or not reset'
            }
            if($case -eq 'Empty'){Check ($rows.Count -eq 0) 'Known empty inventory became failure'}
        }
    }
    foreach($bad in @(0,-1,3601,'abc',1.5)){
        Reset-Fixture Complete LocalFirst $bad;$caught=$null
        try{Invoke-SmartM365InfrastructureCertificateInventory -Rows (New-Object 'System.Collections.Generic.List[object]')}catch{$caught=$_}
        Check ($null -ne $caught -and $script:starts.Count -eq 0) 'Invalid budget accepted'
    }
    Reset-Fixture Complete Invalid;$caught=$null
    try{Invoke-SmartM365InfrastructureCertificateInventory -Rows (New-Object 'System.Collections.Generic.List[object]')}catch{$caught=$_}
    Check ($null -ne $caught -and $script:starts.Count -eq 0) 'Invalid mode accepted'
    Reset-Fixture Complete
    $script:mode=$null;$script:budget=$null
    Invoke-SmartM365InfrastructureCertificateInventory -Rows (New-Object 'System.Collections.Generic.List[object]') -ExchangeServerNames @('SYNTHETIC-EX01')
    Check ($script:reads[0].Limit -le 60 -and $script:starts[0].Mode -eq 'Native') 'Default LocalFirst/120 budget lost'
    $template=Get-Content ([IO.Path]::ChangeExtension($collector,'local.json.template')) -Raw|ConvertFrom-Json
    Check ($template.CertificateTimeoutSecondsPerServer -eq 120 -and $template.ExchangeConnectionMode -eq 'LocalFirst') 'Template defaults mismatch'

    Reset-Fixture
    $script:InfrastructureCertificateWorkerPath=Join-Path $fixtureRoot 'missing-worker.ps1'
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-EX01','SYNTHETIC-EX02')
    Check ($rows.Count -eq 2 -and @($rows|Where-Object CollectionStatus -eq ERROR).Count -eq 2 -and $script:starts.Count -eq 0) 'Missing cached worker blocked infrastructure or became known zero'
    $script:InfrastructureCertificateWorkerPath=$worker
    # The local cache copies the whole script folder (/MIR), so the worker next to the collector is cached.
    $catalogJob=@((Get-Content (Join-Path $smartRoot 'SmartInventory/Orchestrator/Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json).Jobs | Where-Object Name -eq 'ExchangeOnPrem-Infrastructure-Inventory')
    $launchModuleText=Get-Content (Join-Path $smartRoot 'SmartInventory/Orchestrator/SmartM365.Inventory.Launch.psm1') -Raw
    Check ($catalogJob.Count -eq 1 -and $catalogJob[0].Execution.CacheName -and $launchModuleText -match "Label = 'script'; Source = \`$scriptFolder" -and $launchModuleText -match "-Options @\('/MIR'\) -LogOption \`$copy\.Log") 'Worker not covered by launcher cache copy'

    # Cleanup failure stays explicit but cannot abort unrelated infrastructure phases.
    $fakeStop=(Get-Command Stop-SmartM365CertificateWorker).ScriptBlock
    function Stop-SmartM365CertificateWorker {param($State,[switch]$Force) if($null -ne $State){$State.Closed=$true;throw 'Synthetic cleanup failure'}}
    Reset-Fixture Retry
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-EX01','SYNTHETIC-EX03')
    Check ($rows.Count -gt 0 -and $script:ServersAndStorageErrorCount -gt 0 -and ($script:messages -join '|') -match 'cleanup failed') 'Cleanup failure stopped infrastructure or was silent'
    Set-Item Function:Stop-SmartM365CertificateWorker $fakeStop

    # Run the actual readiness call/export and publishing tail, with I/O and external services stubbed.
    $source=Get-Content $collector -Raw
    $originalReadinessCollector=(Get-Command Invoke-ExchangeReadinessCollector).ScriptBlock
    function Invoke-ExchangeReadinessCollector {
        param($Rows,$CollectorName,$ScriptBlock)
        if($CollectorName -eq 'HybridConfiguration'){$script:events.Add('PostCertificateReadiness');Add-ExchangeReadinessRow -Rows $Rows -Category SyntheticOther -ObjectName synthetic -Setting Collection -Value Preserved -MigrationFocus synthetic}
    }
    function Export-ServersAndStorageCsv {param($Data,$Path) $Data|Export-Csv $Path -Delimiter ';' -NoTypeInformation;$script:events.Add('CSV:'+([IO.Path]::GetFileName($Path)))}
    function New-HtmlExecutiveSummary {param($Summary,$PerServerSummary,$ReadinessInventory,$HealthCheckerSummary,$HealthCheckerFindings,$Path) Set-Content $Path '<html>synthetic</html>';$script:events.Add('HTML')}
    function Invoke-ServersAndStorageSharePointUpload {param($LocalFilePath) $script:events.Add('SharePoint')}
    function Save-ServersAndStorageWeeklyHistory {}
    function Remove-CoreSmartM365TimestampedDirectoriesOlderThan {param($RootPath,$RetentionDays,$ExcludeDirectories,$LogFile)}
    function Send-ServersAndStorageHtmlReport {param($HtmlReportPath,$Subject,$Attachments,$Summary,$PerServerSummary,$ReadinessInventory) $script:events.Add('Mail')}
    $a=$source.IndexOf('    $exchangeReadinessInventory = @(Get-ExchangeReadinessInventory')
    $b=$source.IndexOf('    $logicalDiskRows =',$a)
    $c=$source.LastIndexOf('    $summaryPath = Join-Path $OutputFolder')
    $d=$source.IndexOf('    $completionStatus = if',$c)
    Check ($a -gt 0 -and $b -gt $a -and $d -gt $c) 'Publication boundaries changed'
    $readinessCode=[scriptblock]::Create($source.Substring($a,$b-$a))
    $publishCode=[scriptblock]::Create($source.Substring($c,$d-$c))
    $statusNode=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$completionStatus' -and $n.Extent.Text -match 'ServersAndStorageWarningCount'},$true))[-1]
    foreach($case in @('InitDenied','BothTimeout','RemoteTimeout','RemoteDenied')){
        Reset-Fixture $case
        $script:events=New-Object 'System.Collections.Generic.List[string]'
        $exchangeServers=@([pscustomobject]@{Name='SYNTHETIC-EX01'},[pscustomobject]@{Name='SYNTHETIC-EX02'},[pscustomobject]@{Name='SYNTHETIC-EX03'})
        $OutputFolder=$fixtureRoot;$OutputRoot=$fixtureRoot;$RunId='synthetic';$summary=[pscustomobject]@{OtherInventory='Preserved'};$perServerSummary=@();$perServerSummaryPath='synthetic.csv';$healthCheckerSummary=@();$healthCheckerFindings=@()
        . $readinessCode
        . $publishCode
        . ([scriptblock]::Create($statusNode.Extent.Text))
        Check ($completionStatus -eq 'CompletedWithWarnings') 'Timeout reported silent success'
        foreach($event in @('PostCertificateReadiness','CSV:Exchange_OnPrem_MigrationReadiness_Config.csv','CSV:Exchange_OnPrem_Servers_Inventory_Summary.csv','HTML','SharePoint','Mail')){Check ($script:events.Contains($event)) ('Output phase not reached: '+$event)}
        $exported=@(Import-Csv $exchangeReadinessPath -Delimiter ';')
        Check (@($exported|Where-Object {$_.Setting -eq 'Collection' -and $_.CollectionStatus -eq 'ERROR'}).Count -eq 2) 'Failed target rows lost during export'
    }
    $exitNode=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '^if \(\$completionStatus -eq ''CompletedWithWarnings''\).*ExitCode=\(3\)'},$true))[-1]
    $probe=Join-Path $fixtureRoot 'completion-probe.ps1'
    $code='$script:ServersAndStorageWarningCount=1;$smartM365RunGuard=[pscustomobject]@{ReachedTerminal=$false;ExplicitExit=$false;ExitCode=0};'+[Environment]::NewLine+$statusNode.Extent.Text+[Environment]::NewLine+$exitNode.Extent.Text
    [IO.File]::WriteAllText($probe,$code,[Text.UTF8Encoding]::new($true))
    & $hostExe -NoProfile -ExecutionPolicy Bypass -File $probe
    Check ($LASTEXITCODE -eq 3) 'Warning completion exit changed'
    Check ($source -notmatch '\bInvoke-Command\s+-') 'Remote scriptblock introduced'
    Check ($source -match 'Test-ExchangeShell\s+Invoke-CoreSmartM365Preflight[^\r\n]+-RequireExchangeOnPrem') 'Existing loading/preflight changed'
}
finally{
    if([IO.Path]::GetFullPath($fixtureRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)){Remove-Item $fixtureRoot -Recurse -Force}
}
Write-Host ('PASS: {0} certificate budget/worker checks; {1}; no Exchange connection or collector.' -f $checks,$PSVersionTable.PSVersion)

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAd33MbEOT6wrwz
# NujXN/x07FXVXDgnDgOYiUiOTB0kH6CCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgRhvsNSOSFujs
# /vGxef3vIUMVvQVRYoQSmSn/c5LUma4wDQYJKoZIhvcNAQEBBQAEggGAAgmn2X0r
# woyFuoio9poVJgHEkJippnvVYw/2bGBn+qgjf/nwF9G2fU2uJ2vGtF1+ScK2RrUw
# jXX81GbON0ME3MkUd1363Sb4EJ9dAz6yumG1Y7e7MW4kzF/0JHxIU23CQnmhL/mr
# oE2/NhcCVTKkFpJlzyDSiPdfprsPtDZ4iZzPeqe3uKahG5RPvaBPieooaIbCf/cb
# Ty+zakbUvpdBuRR+MfocD7lEF7+P/5viPBZ+wp0T01EIvQZYoLxovJu8lJ5qWSnZ
# SiSLd+UCEqgHZFmUtLmYSy506U/9vZ3mP1o/a2mEe1G7TJRoJFHLb6m0Zr10hsxe
# srRzb1nzec5Rm5YyV3pjlalqKmeHLwkFiOiFlJA/U4GGWPgbi8kEYPRpIzPEutsZ
# UpvNvSwGVpVdoUjVj5IEDkTZIcWXKV2Hlb9dr5W94kyw1J5uQzg97hFqUiDtvudW
# 4+HQUyrmqn1RGE+93Za88v4xWD26Mb6qGUkfGyK66zI/Wa8Dhmz5eU4RoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNTQ2WjA/BgkqhkiG9w0BCQQxMgQw5IVB8+I9DLEne0fjzPFDJE+D
# lYabomABjA9ZK/2sxY+SvlFqQjA5Jsv03w0UT+Y+MA0GCSqGSIb3DQEBAQUABIIC
# AEmRH9cFmk7XCCqr7TyqbOV2ojJcZ4+6u5i/v6NRXsjAiRDAcZVjUezMYZruPbDZ
# ad+Y+K1pdwD3rrUPsixv2qMd4mXb3R+doaXXETt5AXhODgW7gyN/QVb4OVGmTTRU
# gEKGZxEtI4+yCR20mJEPkYrGbpi/zwZrmJXx5bvx5tyZ+Rv4QtSHO8bM67lnMK4m
# Vi/wMWdP/wIcb/DqrgKzUUVljHOAria9Mid3Dq0Z9/nJDzJmeGcpuJneJFwMf3Au
# PdJjm94UyHruNUolpVYo0l5KyjwjtqaQca6plE3B9LGFIvZOozvzPmaZSOQMtCnJ
# C8A1JFo4Sk243JrLusSiYv1PzlwwLwsAJd7xC+1u/QYfD6/NQ7MFPujAymVKuL7F
# vwD3ltODYOs8+48APtUJA27M5QYYHeG6Wovx5vAfgHwG+On/9vgxX1RCSFb373xY
# qYgcorLIDwpPpeg2Q8wZWlZanF8Q8aLvBMYTg8UF16K9tfpgEgc/deKXphrDhKPK
# bKy6GvHEnZI1FimGj9ZQx+H3ZRcQbVBkom8Qf8Bv0qmQI4KGIR8r+kRAts4r6dvU
# guIcLFcstqqBnUBqD/6z1cjdnA/zGS7EOxClwHwUHgSzDqYzq88LwQb2yIydveSr
# 5uSPE7u3+WbsKWABF2/BsvEGCAfdvYJ4nwx2D6CAy+KZ
# SIG # End signature block
