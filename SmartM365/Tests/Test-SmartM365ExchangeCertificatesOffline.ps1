#Requires -Version 5.1
<# .SYNOPSIS Offline certificate budgets, workers, projections and infrastructure continuation. .VERSION 1.1.0 #>
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
$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
Check ($errors.Count -eq 0) 'Collector parser failed'
foreach($name in @('Write-Log','Get-ExchangeReadinessInventory','ConvertTo-ReadinessText','Get-ObjectPropertyValue','Add-ExchangeReadinessRow','Add-ExchangeReadinessProperties','Invoke-ExchangeReadinessCollector','Stop-SmartM365InfrastructureCertificateWorker','Add-SmartM365InfrastructureCertificateProperties','Invoke-SmartM365InfrastructureCertificateInventory')){
    $node=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
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
    $launcher=Join-Path $smartRoot 'SmartInventory/Launchers/OnPremises/Start-SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.cmd'
    Check ((Get-Content $launcher -Raw) -match 'robocopy "%SCRIPT_DIR%" "%LOCAL_SCRIPT_DIR%" /MIR') 'Worker not covered by launcher cache copy'

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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAnt2uBc1KTBoDa
# 8693gdHB9kNswisNR6HjBD9dtRWnAKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICxvtCcOFosxKkaTiIYwVylgdHlIfn6MSHT4CGHF3W++MA0GCSqG
# SIb3DQEBAQUABIIBgFanFIhnmW+OA6AF+yeKUddQVvlcgSRfcgEX/88rCESppA33
# JcfsgOdNokfExjprR/4yErqIniIAp/zMaqE7GtCs40BYiqSXUbjCw5H1fvBu/1KA
# 1Yo4JtecBATkYtukON1L6u7mToRPAh79NEd28JBla1fPIrlARAzr7n0ioqbShauk
# VjEb8iKSazPbRNWI2hKcnKOaCclOH88W1SsyAIwHWbDA9RHkhDQJcac4uZKlBCPg
# VHNjZx9/BIjd7/XnYYYfbOTrjzeYzMDuxnQdEqG2BYWKN25rjJsM22we33SZOsL3
# twNkh46vfRw7ihHIpon3xq8IR1c2IIXTaIgIsG5+O+K/rYMMKWx5DtPRhNpQM61L
# HKubaoPQD5ZODH595DdevUbGdfBd4uLMMOfz0LyfAVuymvBj8DT9nig7y5rwmi4d
# 4/zdT8Q+jRBEjc6aWT5/5aK7DEYZgcAn6qEdty33WyHJhzuM3KThQuHqXlSoWdgW
# 95Botz0UWmETjK2TdaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMyMzE2
# MDhaMC8GCSqGSIb3DQEJBDEiBCAUW/f5YwOhxbOKDXaG4XV7lz0fHwrcXqZy+Bl/
# wbDqqDANBgkqhkiG9w0BAQEFAASCAgB8ohi9Kk+ECsMBwJL/ElLl9Y5arBFD5XTh
# 3xYlbSMdVNZyDYuZVWePNo+soOpzxH0qlhaDDSRSfjFAuvzNrv7wwjuN+GknzKJ6
# ok9byCO24agdDLjnJC6v9BLX+ESw+ImJK3vdUirKkrlnEI/7xSqzhdktrqTCkCtt
# 4n4S6HV1r9CHzScEKuDw2Wf8hjjlFxTuV32m75tSngtgAYEVcBjl5ZF3iARKvRPH
# ZPDeM8NYbesyIDGf2glWbuJ7mzKSZ8J0leLTK24eaXHAYnAs0Paknt027S8th1lH
# KN4jC1rqtKaohA5Re/9YI71i6lTuxaSnVEA25Cvx3tvMBwIuQAzizgaUgl3+Q0u+
# ej4zRUGei1GVzT/7jd3W7ixQc1CQi8doiMlP4hsHBZ2ZllsoMIIx8O3RN6i2Gwh+
# +JA5BUH91ei8dB0hOZo54mwhmZm5mEP4WsXXt5uitAzWSTRGtG7nZo4ZHtfYqawI
# iWoP0uQ0vP/aEOcOyK9rOOKwzuQ6c/sRg3Ol/lt9OBPIMNRhzHshuwrcZJbZr+S5
# /aMuz4/b1I7DTFOs2h8FkbfjP1Jq5gmfQMbkgEvCDOXIfYwu2lMSZokK5DoqONHe
# mmGgEaWoF6d0vtpYQixtWwRX7mo0qe024eryqNZ6FbS5desgRS3mfirENL8q+X3L
# n55Uxvii7A==
# SIG # End signature block
