#Requires -Version 5.1
<# .SYNOPSIS Offline infrastructure certificate session and wire-property tests. .VERSION 1.0.1 #>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$smartRoot=Split-Path $PSScriptRoot -Parent
$collector=Join-Path $smartRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
. (Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-ExchangeOnPremSession.ps1')
$tokens=$null;$parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($collector,[ref]$tokens,[ref]$parseErrors)
$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
Check ($parseErrors.Count -eq 0) 'Collector parser failed'
foreach($name in @('Write-Log','Get-ExchangeReadinessInventory','ConvertTo-ReadinessText','Get-ObjectPropertyValue','Add-ExchangeReadinessRow','Add-ExchangeReadinessProperties',
    'Invoke-SmartM365InfrastructureCertificatesLocalFirst','Invoke-ExchangeReadinessCollector','Initialize-SmartM365InfrastructureCertificateConnection','Get-SmartM365InfrastructureCertificates',
    'Add-SmartM365InfrastructureCertificateProperties','Invoke-SmartM365InfrastructureCertificateInventory')) {
    $node=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    Check ($null -ne $node) ('Missing production function: '+$name)
    . ([scriptblock]::Create($node.Extent.Text))
}
function Wire($Value){[Management.Automation.PSSerializer]::Deserialize([Management.Automation.PSSerializer]::Serialize($Value,10))}
$script:collectorWriteLog=(Get-Command Write-Log).ScriptBlock
function Write-Log {param($Message,$Level='INFO') $script:messages.Add(($Level+': '+$Message)); & $script:collectorWriteLog -Message $Message -Level $Level}
function Get-SmartM365EffectiveConfigValue {param($Name,$DefaultValue) if($Name -eq 'ExchangeConnectionMode'){if($null -eq $script:mode){$DefaultValue}else{$script:mode}}else{$DefaultValue}}
function Get-Command {
    [CmdletBinding()]param([string]$Name,[string[]]$CommandType)
    if($Name -eq 'Microsoft.Exchange.Management.PowerShell.SnapIn\Get-ExchangeCertificate'){
        Check ($CommandType -contains 'Cmdlet') 'LocalFirst did not bind the native snap-in cmdlet'
        return Microsoft.PowerShell.Core\Get-Command Get-SyntheticNativeCertificate
    }
    Microsoft.PowerShell.Core\Get-Command @PSBoundParameters
}
function Get-SyntheticNativeCertificate {
    [CmdletBinding()]param([string]$Server)
    $script:nativeCalls++
    $script:fixture.NativeTargets.Add($Server)
    if($script:fixture.Scenario -eq 'LocalCancel'){throw [OperationCanceledException]::new('Synthetic local cancellation')}
    if(($script:fixture.Scenario -in @('RetryOne','RetryRemoteDenied','ConnectFailure','ImportDenied','RemoteCancel','RetryBroken')) -and $Server -eq 'SYNTHETIC-EX01'){
        throw 'Synthetic native access denied'
    }
    if($script:fixture.Scenario -eq 'RetryMany' -and $Server -ne 'SYNTHETIC-EX03'){throw 'Synthetic native access denied'}
    if($script:fixture.Scenario -eq 'RetryUnavailable' -and $Server -ne 'SYNTHETIC-EX03'){throw 'Synthetic native access denied'}
    $script:fixture.NativeItems
}
function New-PSSession {
    [CmdletBinding()]param($ConfigurationName,$ConnectionUri,$Authentication)
    $script:created++
    Check ($ConfigurationName -eq 'Microsoft.Exchange' -and $Authentication -eq 'Kerberos') 'Wrong Exchange endpoint/authentication'
    if($script:fixture.Scenario -in @('ConnectFailure','RetryUnavailable')){throw 'Synthetic connection failure'}
    $script:fixture.Session=[pscustomobject]@{State='Opened';Id=42}
    $script:fixture.Session
}
function Import-PSSession {
    [CmdletBinding()]param($Session,[string[]]$CommandName,[switch]$AllowClobber,[switch]$DisableNameChecking)
    Check ($CommandName.Count -eq 1 -and $CommandName[0] -eq 'Get-ExchangeCertificate') 'Remote import exceeded certificate scope'
    if($script:fixture.Scenario -eq 'ImportDenied'){throw 'Synthetic certificate import denied'}
    if($script:fixture.Scenario -eq 'MissingRequired'){return New-Module -Name ('SyntheticMissing'+[guid]::NewGuid().ToString('N')) -ScriptBlock {function Get-SyntheticMissing {}}}
    $module=New-Module -Name ('SyntheticCertificates'+[guid]::NewGuid().ToString('N')) -ArgumentList $script:fixture -ScriptBlock {
        param($Fixture)
        $script:Fixture=$Fixture
        function Get-ExchangeCertificate {
            [CmdletBinding()]param([string]$Server)
            $script:Fixture.Calls.Add($Server)
            if($script:Fixture.Scenario -eq 'RemoteCancel'){throw [OperationCanceledException]::new('Synthetic remote cancellation')}
            if($script:Fixture.Scenario -eq 'RetryBroken'){$script:Fixture.Session.State='Broken';throw 'Synthetic retry session broken'}
            if($script:Fixture.Scenario -eq 'RetryRemoteDenied'){throw 'Synthetic remote target denied'}
            if($script:Fixture.Scenario -eq 'BreakAfterFirst' -and $Server -eq 'SYNTHETIC-EX02'){$script:Fixture.Session.State='Broken';throw 'Synthetic remote session broken after one target'}
            if($script:Fixture.Scenario -eq 'BrokenSession'){$script:Fixture.Session.State='Broken';throw 'Synthetic remote session broken'}
            if($script:Fixture.Scenario -eq 'TargetDenied' -and $Server -eq 'SYNTHETIC-EX01'){throw 'Synthetic target access is denied'}
            if($script:Fixture.Scenario -eq 'Empty'){return}
            if($script:Fixture.Scenario -eq 'ShadowProxy'){Set-Item -Path Function:global:Get-ExchangeCertificate -Value $script:Fixture.NativeScriptBlock}
            $script:Fixture.Items
        }
    }
    $script:fixture.ModuleName=$module.Name
    $module
}
function Remove-PSSession {[CmdletBinding()]param($Session) Check ($Session.Id -eq 42) 'Removed unrelated session';$script:removed++}
function Reset-Fixture([string]$Scenario='Complete',[string]$Mode='Remote') {
    $script:mode=$Mode;$script:created=0;$script:removed=0;$script:nativeCalls=0
    $script:ServersAndStorageWarningCount=0;$script:ServersAndStorageErrorCount=0
    $script:ExchangeCertificateContext=$null;$script:messages=New-Object 'System.Collections.Generic.List[string]'
    $item=[pscustomobject]@{Thumbprint=('1'*40);Subject='CN=mail.example.test';CertificateDomains=@('mail.example.test');Services='IIS, SMTP';NotBefore=[datetime]'2026-01-01T00:00:00Z';NotAfter=[datetime]'2027-01-01T00:00:00Z';IsSelfSigned=$false;Status='Valid'}
    if($Scenario -eq 'MissingMetadata') {
        $item=[pscustomobject]@{Thumbprint=('2'*40);Subject='CN=mail.example.test';DnsNameList=@([pscustomobject]@{Unicode='mail.example.test'},[pscustomobject]@{Unicode='autodiscover.example.test'});NotBefore='01/02/2026 12:00:00';NotAfter='01/02/2027 12:00:00'}
    }
    $nativeItem=$item
    if($Mode -eq 'LocalFirst' -or $null -eq $Mode){$item=[pscustomobject]@{Thumbprint=('2'*40);Subject='CN=mail.example.test';DnsNameList=@('mail.example.test');NotBefore='raw-date';NotAfter='raw-date'}}
    $script:fixture=[pscustomobject]@{NativeItems=@($nativeItem);NativeTargets=(New-Object 'System.Collections.Generic.List[string]');Scenario=$Scenario;Session=$null;ModuleName='';Items=@((Wire $item));Calls=(New-Object 'System.Collections.Generic.List[string]');NativeScriptBlock=(Get-Command Get-ExchangeCertificate).ScriptBlock}
}
$savedCertificate=Get-Command Get-ExchangeCertificate -CommandType Function -ErrorAction SilentlyContinue
$savedServer=Get-Command Get-ExchangeServer -CommandType Function -ErrorAction SilentlyContinue
function global:Get-ExchangeCertificate {[CmdletBinding()]param($Server) $script:nativeCalls++;$script:fixture.Items}
function global:Get-ExchangeServer {'SyntheticNativeExchangeServer'}
$fixtureRoot=Join-Path ([IO.Path]::GetTempPath()) ('InfrastructureCertificates-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $fixtureRoot
$LogFile=Join-Path $fixtureRoot 'synthetic.log'
try {
    foreach($mode in @('Remote','LocalSnapin')) {
        Reset-Fixture -Mode $mode
        $rows=New-Object 'System.Collections.Generic.List[object]'
        Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-EX01','SYNTHETIC-EX02')
        Check ($rows.Count -eq 16) ('Existing certificate schema changed in '+$mode)
        Check (@($rows | Where-Object CollectionStatus -ne 'OK').Count -eq 0) 'Known native/wire properties were not preserved'
        Check (@($rows | Where-Object {$_.Setting -eq 'IsSelfSigned' -and $_.Value -eq 'False'}).Count -eq 2) 'False became unknown'
        Check ((Get-ExchangeServer) -eq 'SyntheticNativeExchangeServer') 'Non-certificate Exchange command changed'
        Check ($null -eq $script:ExchangeCertificateContext) 'Context retained after certificate phase'
        if($mode -eq 'Remote') {
            Check ($script:created -eq 1 -and $script:removed -eq 1 -and $script:nativeCalls -eq 0) 'Remote session count or fallback incorrect'
            Check ($script:fixture.Calls.Count -eq 2 -and -not(Get-Module $script:fixture.ModuleName)) 'Remote server calls or cleanup incorrect'
        }
        else {Check ($script:created -eq 0 -and $script:nativeCalls -eq 2) 'LocalSnapin behavior changed'}
        Check (($script:messages -join '|') -match 'duration seconds=') 'Certificate duration missing'
    }
    Reset-Fixture 'MissingMetadata'
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-EX01')
    Check (@($rows|Where-Object {$_.Value -eq 'Unknown' -and $_.CollectionStatus -eq 'WARNING'}).Count -eq 3) 'Missing Services/Status/IsSelfSigned must be Unknown'
    Check (@($rows|Where-Object {$_.Setting -eq 'CertificateDomains' -and $_.Value -eq 'mail.example.test; autodiscover.example.test'}).Count -eq 1) 'Wire DNS names were lost or invented'
    Check (@($rows|Where-Object {$_.Setting -eq 'NotAfter' -and $_.Value -eq '01/02/2027 12:00:00'}).Count -eq 1) 'Display date text was interpreted instead of preserved'
    Check (($script:messages -join '|') -match 'metadata unavailable') 'Remote property loss was not logged'

    Reset-Fixture 'Complete'
    $item=Wire ([pscustomobject]@{Subject='CN=synthetic';CertificateDomains=@();Services=0;IsSelfSigned=$false;NotBefore='Unknown';NotAfter='Unknown';Status='Unknown'})
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Add-SmartM365InfrastructureCertificateProperties -Rows $rows -InputObject $item -ObjectName 'synthetic' -Mode Remote
    Check (@($rows|Where-Object {$_.Setting -eq 'Services' -and $_.Value -eq '0' -and $_.CollectionStatus -eq 'OK'}).Count -eq 1) 'Known zero became missing'
    Check (@($rows|Where-Object {$_.Setting -eq 'CertificateDomains' -and $_.Value -eq '(none)'}).Count -eq 1) 'Known empty domain list became an error'

    $fallback=Wire ([pscustomobject]@{Subject='CN=synthetic';CertificateDomains=@('mail.example.test');ServicesStringForm='None';IsSelfSigned=$true;NotBefore='2026-01-01T00:00:00Z';NotAfter='2027-01-01T00:00:00Z';Status='Valid'})
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Add-SmartM365InfrastructureCertificateProperties -Rows $rows -InputObject $fallback -ObjectName 'synthetic' -Mode Remote
    Check (@($rows|Where-Object {$_.Setting -eq 'Services' -and $_.Value -eq 'None' -and $_.CollectionStatus -eq 'OK'}).Count -eq 1) 'Explicit ServicesStringForm was lost'
    Check (@($rows|Where-Object {$_.Setting -eq 'IsSelfSigned' -and $_.Value -eq 'True'}).Count -eq 1) 'True self-signed flag was lost'

    foreach($scenario in @('Empty','TargetDenied','BrokenSession','BreakAfterFirst','ShadowProxy','ConnectFailure','ImportDenied','MissingRequired')) {
        Reset-Fixture $scenario
        $rows=New-Object 'System.Collections.Generic.List[object]';$caught=$null;$continued=$false
        try{
            Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-EX01','SYNTHETIC-EX02','SYNTHETIC-EX03')
            $continued=$true
        }catch{$caught=$_}
        Check ($null -eq $caught -and $continued) ('Certificate failure stopped infrastructure: '+$scenario)
        Check ($script:nativeCalls -eq 0 -and $null -eq $script:ExchangeCertificateContext) 'Failure used fallback or leaked context'
        Check ($script:created -eq 1 -and $script:removed -eq [int]($scenario -ne 'ConnectFailure')) ('Failure leaked owned session: '+$scenario)
        $failed=@($rows | Where-Object {$_.Category -eq 'ExchangeCertificate' -and $_.Setting -eq 'Collection' -and $_.Value -eq 'Failed' -and $_.CollectionStatus -eq 'ERROR'})
        if($scenario -eq 'Empty'){Check ($rows.Count -eq 0 -and ($script:messages -join '|') -match 'Certificates: 0') 'Known empty result was not distinguished';continue}
        $expectedFailed=if($scenario -eq 'TargetDenied'){@('SYNTHETIC-EX01')}elseif($scenario -in @('ShadowProxy','BreakAfterFirst')){@('SYNTHETIC-EX02','SYNTHETIC-EX03')}else{@('SYNTHETIC-EX01','SYNTHETIC-EX02','SYNTHETIC-EX03')}
        Check (($failed.ObjectName -join ',') -eq ($expectedFailed -join ',')) ('Failed diagnostics missing/duplicated or successful target invalidated: '+$scenario)
        Check (($script:messages -join '|') -notmatch 'Certificates: 0') 'Failed targets became a known zero'
        if($scenario -eq 'TargetDenied') {Check ($rows.Count -eq 17 -and $script:fixture.Calls.Count -eq 3) 'Target-specific failure prevented other targets';continue}
        Check (@($failed | Where-Object {$_.ErrorMessage -notmatch 'Remote certificate session unavailable; no LocalSnapin fallback'}).Count -eq 0) 'Session failure diagnostic omitted no-fallback reason'
        Check ($script:ServersAndStorageErrorCount -eq 1 -and $script:ServersAndStorageWarningCount -gt 0) 'Session failure was not counted by real collector logging'
        Check (($script:messages -join '|') -match 'ERROR: ExchangeConnectionMode=Remote; certificate state=Failed; no fallback') 'Missing explicit session ERROR'
        if($scenario -in @('BrokenSession','ShadowProxy')){Check ($script:fixture.Calls.Count -eq 1) 'Unavailable session/proxy continued target calls'}
        if($scenario -eq 'BreakAfterFirst'){Check ($script:fixture.Calls.Count -eq 2 -and $rows.Count -eq 10) 'Successful certificate rows were lost or unread target called'}
        if($scenario -in @('ConnectFailure','ImportDenied','MissingRequired')){Check ($script:fixture.Calls.Count -eq 0) 'Startup failure still called certificates'}
    }
    Reset-Fixture -Mode Invalid
    $caught=$null;try{Invoke-SmartM365InfrastructureCertificateInventory -Rows (New-Object 'System.Collections.Generic.List[object]')}catch{$caught=$_}
    Check ($null -ne $caught -and $script:created -eq 0) 'Invalid mode accepted'

    # LocalFirst preserves native metadata and retries only failed targets, with one lazy session.
    foreach($scenario in @('Complete','RetryOne','RetryMany','RetryRemoteDenied','ConnectFailure','ImportDenied','RetryUnavailable','RetryBroken','LocalCancel','RemoteCancel')) {
        Reset-Fixture $scenario LocalFirst
        $rows=New-Object 'System.Collections.Generic.List[object]';$caught=$null;$continued=$false
        try{Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-EX01','SYNTHETIC-EX02','SYNTHETIC-EX03');$continued=$true}catch{$caught=$_}
        $cancelled=$scenario -in @('LocalCancel','RemoteCancel')
        Check ($continued -eq (-not $cancelled)) ('LocalFirst continuation/cancellation incorrect: '+$scenario)
        if($cancelled){
            Check ($null -ne $caught -and $caught.Exception -is [OperationCanceledException]) 'Cancellation was converted into a certificate failure'
            Check (@($rows|Where-Object Setting -eq Collection).Count -eq 0) 'Cancellation produced Failed/ERROR'
            Check ($script:fixture.NativeTargets.Count -eq 1) 'Cancellation continued to later servers'
            Check ($script:created -eq [int]($scenario -eq 'RemoteCancel') -and $script:removed -eq $script:created) 'Cancellation leaked/opened a session'
            continue
        }
        Check ($null -eq $caught -and $script:nativeCalls -eq 3) ('LocalFirst skipped local reads after importing proxy: '+$scenario)
        Check (($script:fixture.NativeTargets -join ',') -eq 'SYNTHETIC-EX01,SYNTHETIC-EX02,SYNTHETIC-EX03') 'Native per-target order changed'
        $expectedRetries=if($scenario -eq 'Complete'){0}elseif($scenario -in @('RetryMany','RetryUnavailable')){2}else{1}
        $paths=@($rows|Where-Object {$_.Setting -eq 'CollectionPath' -and $_.Value -eq 'RemoteRetry'})
        Check ($paths.Count -eq $expectedRetries) 'RemoteRetry provenance missing or native success retried'
        Check ($script:created -eq [int]($expectedRetries -gt 0)) 'Session was not lazy or was opened more than once'
        Check ($script:removed -eq [int]($expectedRetries -gt 0 -and $scenario -notin @('ConnectFailure','RetryUnavailable'))) 'Session cleanup incorrect'
        Check ($null -eq $script:ExchangeCertificateContext) 'LocalFirst retained context'
        $failed=@($rows|Where-Object {$_.Setting -eq 'Collection' -and $_.Value -eq 'Failed' -and $_.CollectionStatus -eq 'ERROR'})
        $expectedFailures=if($scenario -eq 'RetryUnavailable'){2}elseif($scenario -in @('RetryRemoteDenied','ConnectFailure','ImportDenied','RetryBroken')){1}else{0}
        Check ($failed.Count -eq $expectedFailures) ('LocalFirst failure diagnostics incorrect: '+$scenario)
        Check (@($failed|Where-Object {$_.ErrorMessage -notmatch 'Local: Synthetic native access denied; Remote: .+'}).Count -eq 0) 'Combined local/remote error lost'
        $knownNative=@($rows|Where-Object {$_.Setting -eq 'Services' -and $_.Value -eq 'IIS, SMTP'})
        Check ($knownNative.Count -eq (3-$expectedRetries)) 'Native metadata lost after proxy import'
        $expectedRemote=if($scenario -eq 'RetryMany'){2}elseif($scenario -eq 'RetryOne'){1}else{0}
        Check (@($rows|Where-Object {$_.Setting -eq 'Services' -and $_.Value -eq 'Unknown'}).Count -eq $expectedRemote) 'Remote metadata rendering changed or native metadata became Unknown'
        Check ($script:fixture.Calls.Count -eq $(if($scenario -in @('ConnectFailure','RetryUnavailable','ImportDenied')){0}else{$expectedRetries})) 'Unexpected remote calls'
    }
    Reset-Fixture Complete
    $script:mode=$null
    $rows=New-Object 'System.Collections.Generic.List[object]'
    Invoke-SmartM365InfrastructureCertificateInventory -Rows $rows -ExchangeServerNames @('SYNTHETIC-EX01')
    Check ($script:nativeCalls -eq 1 -and $script:created -eq 0) 'Missing config key did not default to LocalFirst'
    $template=Get-Content -LiteralPath ([IO.Path]::ChangeExtension($collector,'local.json.template')) -Raw|ConvertFrom-Json
    Check ($template.ExchangeConnectionMode -eq 'LocalFirst') 'Template default differs from code'

    # Exercise the actual readiness call/export and publishing tail with offline I/O stubs.
    $source=Get-Content -LiteralPath $collector -Raw
    $script:productionReadinessCollector=(Get-Command Invoke-ExchangeReadinessCollector).ScriptBlock
    function Invoke-ExchangeReadinessCollector {
        param($Rows,$CollectorName,$ScriptBlock)
        if($CollectorName -eq 'ExchangeCertificate'){& $script:productionReadinessCollector -Rows $Rows -CollectorName $CollectorName -ScriptBlock $ScriptBlock}
        elseif($CollectorName -eq 'HybridConfiguration'){
            $script:events.Add('PostCertificateReadiness')
            Add-ExchangeReadinessRow -Rows $Rows -Category 'SyntheticOtherPhase' -ObjectName 'synthetic' -Setting 'Collection' -Value 'Preserved' -MigrationFocus 'Offline continuation marker.'
        }
    }
    function Export-ServersAndStorageCsv {param($Data,$Path) $Data | Export-Csv -LiteralPath $Path -Delimiter ';' -NoTypeInformation; $script:events.Add(('CSV:'+([IO.Path]::GetFileName($Path))))}
    function New-HtmlExecutiveSummary {param($Summary,$PerServerSummary,$ReadinessInventory,$HealthCheckerSummary,$HealthCheckerFindings,$Path) Set-Content -LiteralPath $Path -Value '<html>synthetic</html>'; $script:events.Add('HTML')}
    function Invoke-ServersAndStorageSharePointUpload {param($LocalFilePath) $script:events.Add('SharePoint')}
    function Save-ServersAndStorageWeeklyHistory {$script:events.Add('WeeklyHistory')}
    function Remove-CoreSmartM365TimestampedDirectoriesOlderThan {param($RootPath,$RetentionDays,$ExcludeDirectories,$LogFile) $script:events.Add('Retention')}
    function Send-ServersAndStorageHtmlReport {param($HtmlReportPath,$Subject,$Attachments,$Summary,$PerServerSummary,$ReadinessInventory) $script:events.Add('Mail')}
    $readinessStart=$source.IndexOf('    $exchangeReadinessInventory = @(Get-ExchangeReadinessInventory')
    $readinessEnd=$source.IndexOf('    $logicalDiskRows =', $readinessStart)
    $publishStart=$source.LastIndexOf('    $summaryPath = Join-Path $OutputFolder')
    $publishEnd=$source.IndexOf('    $completionStatus = if', $publishStart)
    Check ($readinessStart -gt 0 -and $readinessEnd -gt $readinessStart -and $publishEnd -gt $publishStart) 'Publication boundaries changed'
    $readinessCode=[scriptblock]::Create($source.Substring($readinessStart,$readinessEnd-$readinessStart))
    $publishCode=[scriptblock]::Create($source.Substring($publishStart,$publishEnd-$publishStart))
    $statusNode=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$completionStatus' -and $n.Extent.Text -match 'ServersAndStorageWarningCount'},$true))[-1]
    $statusCode=[scriptblock]::Create($statusNode.Extent.Text)
    foreach($case in @(
        @{Scenario='ConnectFailure';Mode='Remote'},@{Scenario='ImportDenied';Mode='Remote'},@{Scenario='BrokenSession';Mode='Remote'},@{Scenario='BreakAfterFirst';Mode='Remote'},@{Scenario='ShadowProxy';Mode='Remote'},
        @{Scenario='RetryRemoteDenied';Mode='LocalFirst'},@{Scenario='ConnectFailure';Mode='LocalFirst'},@{Scenario='RetryUnavailable';Mode='LocalFirst'}
    )) {
        $scenario=$case.Scenario
        Reset-Fixture $scenario $case.Mode
        $script:events=New-Object 'System.Collections.Generic.List[string]'
        $exchangeServers=@([pscustomobject]@{Name='SYNTHETIC-EX01'},[pscustomobject]@{Name='SYNTHETIC-EX02'},[pscustomobject]@{Name='SYNTHETIC-EX03'})
        $OutputFolder=$fixtureRoot;$OutputRoot=$fixtureRoot;$RunId='synthetic';$summary=[pscustomobject]@{OtherInventory='Preserved'};$perServerSummary=@();$perServerSummaryPath='synthetic.csv';$healthCheckerSummary=@();$healthCheckerFindings=@()
        . $readinessCode
        . $publishCode
        . $statusCode
        Check ($completionStatus -eq 'CompletedWithWarnings' -and $script:ServersAndStorageErrorCount -eq [int]($scenario -ne 'RetryRemoteDenied')) 'Existing completion selector or target-only warning accounting changed'
        foreach($event in @('PostCertificateReadiness','CSV:Exchange_OnPrem_MigrationReadiness_Config.csv','CSV:Exchange_OnPrem_Servers_Inventory_Summary.csv','HTML','SharePoint','Mail')){Check ($script:events.Contains($event)) ('Infrastructure output phase not reached: '+$scenario+' / '+$event)}
        $exported=@(Import-Csv -LiteralPath $exchangeReadinessPath -Delimiter ';')
        Check (@($exported|Where-Object {$_.Category -eq 'ExchangeCertificate' -and $_.CollectionStatus -eq 'ERROR'}).Count -eq $(if($case.Mode -eq 'LocalFirst'){if($scenario -eq 'RetryUnavailable'){2}else{1}}elseif($scenario -in @('BreakAfterFirst','ShadowProxy')){2}else{3})) 'Failed targets did not survive readiness CSV export'
        Check ($script:nativeCalls -eq $(if($case.Mode -eq 'LocalFirst'){3}else{0}) -and $script:removed -eq [int]($scenario -notin @('ConnectFailure','RetryUnavailable'))) 'Continuation leaked resources or changed native calls'
    }
    # Execute the unchanged completion exit branch in a child host; it must return 3.
    $exitNode=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text -match '^if \(\$completionStatus -eq ''CompletedWithWarnings''\).*ExitCode=\(3\)'},$true))[-1]
    $probe=Join-Path $fixtureRoot 'completion-probe.ps1'
    $probeCode='$script:ServersAndStorageWarningCount=1; $script:ServersAndStorageErrorCount=1; $smartM365RunGuard=[pscustomobject]@{ReachedTerminal=$false;ExplicitExit=$false;ExitCode=0};'+[Environment]::NewLine+$statusNode.Extent.Text+[Environment]::NewLine+$exitNode.Extent.Text
    [IO.File]::WriteAllText($probe,$probeCode,[Text.UTF8Encoding]::new($true))
    $hostExe=Join-Path $PSHOME $(if($PSVersionTable.PSEdition -eq 'Desktop'){'powershell.exe'}else{'pwsh.exe'})
    & $hostExe -NoProfile -ExecutionPolicy Bypass -File $probe
    Check ($LASTEXITCODE -eq 3) 'Existing completion exit code changed'

    Check ($source -notmatch '\bInvoke-Command\s+-') 'Server-side scriptblock introduced'
    Check ($source -match 'Test-ExchangeShell\s+Invoke-CoreSmartM365Preflight[^\r\n]+-RequireExchangeOnPrem') 'Existing Exchange loading/preflight path changed'
    $inventory=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-ExchangeReadinessInventory'},$true)
    Check ([regex]::Matches($inventory.Extent.Text,'Invoke-SmartM365InfrastructureCertificateInventory -Rows').Count -eq 1) 'Certificate phase not called exactly once'
}
finally {
    if($null -ne $script:ExchangeCertificateContext){Stop-SmartM365ExchangeOnPremSession $script:ExchangeCertificateContext}
    if($savedCertificate){Set-Item Function:global:Get-ExchangeCertificate $savedCertificate.ScriptBlock}else{Remove-Item Function:global:Get-ExchangeCertificate -ErrorAction SilentlyContinue}
    if($savedServer){Set-Item Function:global:Get-ExchangeServer $savedServer.ScriptBlock}else{Remove-Item Function:global:Get-ExchangeServer -ErrorAction SilentlyContinue}
    if([IO.Path]::GetFullPath($fixtureRoot).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase)){Remove-Item -LiteralPath $fixtureRoot -Recurse -Force}
}
Write-Host ("PASS: {0} infrastructure certificate checks; {1}; no Exchange connection or collector." -f $checks,$PSVersionTable.PSVersion)
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA8npZB/mtaUoB0
# xK4EJqGPyiCOG2A0TqFGsjzWSXcT4aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJCw+9560JjiVj9wXrXdR/Poa4EQ6AB4kTjb6X21hyg3MA0GCSqG
# SIb3DQEBAQUABIIBgGScjyIxq5pFaieue6RRw++EvzC9HXUw0ME2N2qgTOvm526L
# +rs2j0WxiJCoeuSIz6MYK2CyRO7WlPvdAsXcWLhqGgAO+ZbhQmO51osCuM7jb1C0
# xyeKiXu70WsdVn2EXSGp01e8ECdEJsU4kL9DrsevuOukFchWIO14MFaJycFPuQbm
# SwVgiz7ugCnHmxfprunI4BHlXtBfq4wg2qs130XCxuZp6VsYCj2bb0OpwVNmZ2tB
# byBGfUeRk/MeQPPyuzGm1BpUHQExL7sFkkX0+hB1IH/fnwAZTdF9eNT5HVSXAz5Q
# dh6HfGRUsmN5+2sNQhGICCguRVmBzcDxbZIkWvJ0PrgJY3FtWnPNkSljqsLFnhBC
# sHrnDELVTMO+7YIGKthXUbKMFiWTpTLS0CadO++MOdNd4kseIzNE0uBqkQn8DhRk
# OPvof/gT5QOTFrL1Nnw8wac4aE4nMXBZooxGLtWteY2m7Q9NGr181wmN/LLLbxsZ
# xP1fhGtgns2Wrc6Je6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMyMTA1
# MDVaMC8GCSqGSIb3DQEJBDEiBCDbNzyqzB7eeS4Gli2baVoNJzqXX+f6MG8tFfaO
# mGRWWTANBgkqhkiG9w0BAQEFAASCAgBQbblu7VdnUzBTunChZtxHYOyrqLVgXQ60
# 1zYhl8YLEMTWC/J8YCMnqSzcH5X/FDZpv5fSENUoRuaJjjWDRlkWQxjdAWN8sclX
# pDKk/X1BH/XNAtKSU1LLgP9ziy9DzBEB6bU5CNv5RXxGR3SBkFgZyfG2vK6hQtA2
# IbRcmn8ct6ab2KoJJFElSsNVAKjh+VM3/Tvxd1aIyIwXFwNxGj01LLUzylw14jWA
# nyf0I827bhBAowx5JoMxx7qJnSf8wcew948ynHS/QukFVHUW+Mx9eO5u+R3g64a9
# PfgewRJPk8R5QLSUz1jsZOKTjuwsY6NR/AGiZfn5XAgYXIkusisSQI26yITnoMGy
# +SuHMLVlG9s987CQc+fjqlGGRMbnrjYePogrSASypuZe5YCvIiTC5AqExhpvK4zJ
# LRpTaFfu/PjIM3FIR7adsxDXzWs/tFqklnDDlklGMcQdx/Zd08ao9LkNrrcM96GD
# JYzFn0HkITE+ccUTdJYgNJFcAR9Ydxs5KrgerVi/LrtcBtBe72gMZXYeuEBd32J1
# 2rTfC5ynOIgAIQFNVgL2JHNMi13uhF8q13DbzPYZ7+m09Rqp1Qb9q7JGjygZuC8U
# mFDFFE2xHwsT9s/SIE5X1L5CnHVmUaNhU6ykb7nZLADmimN32HtwUWsT0m4/supC
# yqnPK123Hw==
# SIG # End signature block
