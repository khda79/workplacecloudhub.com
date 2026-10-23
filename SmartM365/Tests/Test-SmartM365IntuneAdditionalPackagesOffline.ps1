<#
.SYNOPSIS
    Executes additional Intune remediation packages with synthetic endpoint interfaces.
.VERSION
    1.0.0
.NOTES
    No real registry, service, task, certificate, BITS, printer or network operations.
    Synthetic observations and logs stay under TestRoot.
#>
[CmdletBinding()]
param([string]$TestRoot = (Join-Path $env:TEMP ('SmartM365-IntunePackages-' + [guid]::NewGuid().ToString('N'))), [string]$CaseFilter = '*')
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$packageRoot = Join-Path $repoRoot 'SmartM365/Intune/Remediation/Packages'
$childHost = Join-Path $PSHOME $(if ($PSVersionTable.PSEdition -eq 'Desktop') { 'powershell.exe' } else { 'pwsh.exe' })
[IO.Directory]::CreateDirectory($TestRoot) | Out-Null
$script:results = [Collections.Generic.List[object]]::new()

function Invoke-PackageCase {
    param([string]$Name,[string]$Folder,[string]$Package,[string]$Role,[string]$Setup,[int]$Exit,[string]$Pattern,[string]$Arguments='',[string]$Forbidden='OP:')
    if ($Name -notlike $CaseFilter) { return }
    $path = Join-Path (Join-Path (Join-Path $packageRoot $Folder) $Package) ('SmartM365-' + $Package + '-' + $Role + '.ps1')
    $scratch = Join-Path $TestRoot $Name
    [IO.Directory]::CreateDirectory($scratch) | Out-Null
    $guard = @'
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
foreach($name in @('Get-Service','Get-CimInstance','Get-WinEvent','Get-ItemProperty','Get-MpComputerStatus','Update-MpSignature',
 'Get-BitsTransfer','Resume-BitsTransfer','Remove-BitsTransfer','Get-Printer','Get-PrintJob','Remove-PrintJob',
 'Start-Service','Restart-Service','Stop-Service','Set-Service','Start-ScheduledTask','Register-ScheduledTask',
 'Get-ScheduledTask','Get-AuthenticodeSignature','Invoke-LapsPolicyProcessing','w32tm.exe','msg.exe','certutil.exe',
 'Restart-Computer','Start-Process','Invoke-WebRequest','Invoke-RestMethod','Remove-Item','Set-ItemProperty')) {
    Set-Item ('Function:global:'+$name) ([scriptblock]::Create("[Console]::WriteLine('UNEXPECTED:$name'); throw 'Blocked operational command'"))
}
function Test-Path {
    [CmdletBinding()]param($LiteralPath,$Path)
    $p=if($LiteralPath){$LiteralPath}else{$Path}
    if($p -match '^(HKLM:|Cert:)'){[Console]::WriteLine('UNEXPECTED:ProviderRead');throw 'Blocked provider read'}
    Microsoft.PowerShell.Management\Test-Path -LiteralPath $p
}
function Get-ChildItem { [CmdletBinding()]param($LiteralPath) [Console]::WriteLine('UNEXPECTED:ProviderEnumeration');throw 'Blocked enumeration' }
function Start-Sleep { param($Seconds,$Milliseconds) }
'@
    $code = $guard + "`n`$env:ProgramData='" + $scratch.Replace("'","''") + "'`n" + $Setup +
        "`n& ([scriptblock]::Create([IO.File]::ReadAllText('" + $path.Replace("'","''") + "'))) " + $Arguments
    $start = [Diagnostics.ProcessStartInfo]::new($childHost)
    $start.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
    $start.UseShellExecute=$false; $start.CreateNoWindow=$true
    $start.RedirectStandardOutput=$true; $start.RedirectStandardError=$true
    $process=[Diagnostics.Process]::Start($start)
    try {
        $stdout=$process.StandardOutput.ReadToEndAsync(); $stderr=$process.StandardError.ReadToEndAsync()
        if(-not $process.WaitForExit(30000)){ $process.Kill(); throw "Synthetic test timed out: $Name" }
        $outputText=$stdout.GetAwaiter().GetResult()+$stderr.GetAwaiter().GetResult()
        $ok=$process.ExitCode -eq $Exit -and $outputText -match $Pattern -and $outputText -notmatch 'UNEXPECTED:'
        if($Forbidden -and $outputText -match $Forbidden){$ok=$false}
        $lines=@($outputText -split '\r?\n' | Where-Object {$_ -match ' Status='})
        if($lines.Count -ne 1 -or $lines[0].Length -gt 2048){$ok=$false}
        $script:results.Add([pscustomobject]@{Case=$Name;Passed=$ok;Exit=$process.ExitCode;ExpectedExit=$Exit;Output=$outputText.Trim()})
        if(-not $ok){Write-Output "FAIL $Name : $outputText"}
    }
    finally{$process.Dispose()}
}

$laps=@'
function Test-Path { [CmdletBinding()]param($LiteralPath) $LiteralPath -eq 'HKLM:\SOFTWARE\Microsoft\Policies\LAPS' }
function Get-ItemProperty { [CmdletBinding()]param($LiteralPath) [pscustomobject]@{BackupDirectory=1} }
function Get-WinEvent {
 [CmdletBinding()]param($ListLog,$FilterHashtable,$MaxEvents)
 if($ListLog){[pscustomobject]@{IsEnabled=$true}}else{[pscustomobject]@{Id=10004;RecordId=10;TimeCreated=(Get-Date)}}
}
'@
Invoke-PackageCase 'laps-healthy' Security-LAPS Windows-LAPS-Health Detection $laps 0 'RecentLocalPolicySuccess'
Invoke-PackageCase 'laps-disabled-log' Security-LAPS Windows-LAPS-Health Detection ($laps.Replace('IsEnabled=$true','IsEnabled=$false')) 2 'Status=Error'
Invoke-PackageCase 'laps-future-outcome' Security-LAPS Windows-LAPS-Health Detection ($laps.Replace('TimeCreated=(Get-Date)','TimeCreated=(Get-Date).AddDays(1)')) 2 'Status=Error'
Invoke-PackageCase 'laps-unconfigured' Security-LAPS Windows-LAPS-Health Remediation ($laps+"`nfunction Get-ItemProperty { [CmdletBinding()]param(`$LiteralPath) [pscustomobject]@{BackupDirectory=0} }") 0 'NotApplicable'
Invoke-PackageCase 'laps-failed-detection' Security-LAPS Windows-LAPS-Health Detection ($laps+"`nfunction Get-WinEvent { [CmdletBinding()]param(`$ListLog,`$FilterHashtable,`$MaxEvents) [pscustomobject]@{IsEnabled=1;Id=10005;RecordId=10;TimeCreated=(Get-Date)} }") 1 'RemediationRequired'
$lapsRetry=$laps+@'

$script:retried=$false
function Invoke-LapsPolicyProcessing { [CmdletBinding()]param() $script:retried=$true; [Console]::WriteLine('OP:Laps') }
function Get-WinEvent { [CmdletBinding()]param($ListLog,$FilterHashtable,$MaxEvents)
 if($ListLog){return [pscustomobject]@{IsEnabled=$true}}
 [pscustomobject]@{Id=$(if($script:retried){10004}else{10005});RecordId=$(if($script:retried){11}else{10});TimeCreated=(Get-Date)}
}
'@
Invoke-PackageCase 'laps-retry-success' Security-LAPS Windows-LAPS-Health Remediation $lapsRetry 0 'NewLocalPolicySuccess.*RemoteBackupVerified=False' '' 'OP:(?!Laps)'
Invoke-PackageCase 'laps-no-new-outcome' Security-LAPS Windows-LAPS-Health Remediation ($lapsRetry+"`nfunction Get-WinEvent { [CmdletBinding()]param(`$ListLog,`$FilterHashtable,`$MaxEvents) [pscustomobject]@{IsEnabled=1;Id=10005;RecordId=10;TimeCreated=(Get-Date)} }") 1 'NoNewPolicyOutcome' '' 'OP:(?!Laps)'

$time=@'
function Get-Service { [CmdletBinding()]param($Name) [pscustomobject]@{Status='Running'} }
function Get-CimInstance { [CmdletBinding()]param($ClassName,$Filter) [pscustomobject]@{StartMode='Auto'} }
function w32tm.exe { if($args -contains '/source'){'time.contoso.example'}else{'12:00:00, +00.0500000s';'12:00:01, -00,0400000s';'12:00:02, +00.0300000s'}; $global:LASTEXITCODE=0 }
'@
Invoke-PackageCase 'time-healthy' System-Time Windows-Time-Sync Detection $time 0 'OffsetWithinThreshold'
Invoke-PackageCase 'time-disabled' System-Time Windows-Time-Sync Remediation ($time+"`nfunction Get-CimInstance { [CmdletBinding()]param(`$ClassName,`$Filter) [pscustomobject]@{StartMode='Disabled'} }") 0 'Deferred.*Disabled'
Invoke-PackageCase 'time-offline' System-Time Windows-Time-Sync Remediation ($time+"`nfunction w32tm.exe { 'not reachable'; `$global:LASTEXITCODE=1 }") 0 'Deferred.*Unreachable'
Invoke-PackageCase 'time-drift' System-Time Windows-Time-Sync Detection ($time+"`nfunction w32tm.exe { if(`$args -contains '/source'){'time.contoso.example'}else{1..3|ForEach-Object {'12:00:00, +180.0000000s'}}; `$global:LASTEXITCODE=0 }") 1 'RemediationRequired'
$timeRetry=$time+@'

$script:resynced=$false
function w32tm.exe {
 if($args -contains '/source'){'time.contoso.example'}
 elseif($args -contains '/resync'){$script:resynced=$true;[Console]::WriteLine('OP:Resync')}
 else{1..3|ForEach-Object {if($script:resynced){'12:00:00, +0.050s'}else{'12:00:00, +180.000s'}}}
 $global:LASTEXITCODE=0
}
'@
Invoke-PackageCase 'time-resync' System-Time Windows-Time-Sync Remediation $timeRetry 0 'OffsetVerified' '' 'OP:(?!Resync)'

$defender=@'
function Get-MpComputerStatus { [CmdletBinding()]param() [pscustomobject]@{AMRunningMode='Normal';AMServiceEnabled=$true;AntivirusEnabled=$true;AntivirusSignatureLastUpdated=(Get-Date)} }
'@
Invoke-PackageCase 'defender-fresh' Security-Defender Defender-Signature-Freshness Remediation $defender 0 'SignaturesFresh'
Invoke-PackageCase 'defender-unknown-mode' Security-Defender Defender-Signature-Freshness Remediation ($defender.Replace("AMRunningMode='Normal'","AMRunningMode=''")) 1 'Status=Error'
Invoke-PackageCase 'defender-passive' Security-Defender Defender-Signature-Freshness Remediation "function Get-MpComputerStatus { [CmdletBinding()]param() [pscustomobject]@{AMRunningMode='Passive Mode'} }" 0 'NotApplicable'
Invoke-PackageCase 'defender-stale' Security-Defender Defender-Signature-Freshness Detection ($defender.Replace('(Get-Date)','(Get-Date).AddDays(-3)')) 1 'SignaturesStale'
$defenderRetry=@'
$script:updated=$false
function Update-MpSignature { [CmdletBinding()]param() $script:updated=$true; [Console]::WriteLine('OP:UpdateSignature') }
function Get-MpComputerStatus { [CmdletBinding()]param() [pscustomobject]@{AMRunningMode='Normal';AMServiceEnabled=$true;AntivirusEnabled=$true;AntivirusSignatureLastUpdated=$(if($script:updated){Get-Date}else{(Get-Date).AddDays(-3)})} }
'@
Invoke-PackageCase 'defender-update' Security-Defender Defender-Signature-Freshness Remediation $defenderRetry 0 'FreshSignaturesVerified' '' 'OP:(?!UpdateSignature)'
Invoke-PackageCase 'defender-update-unresolved' Security-Defender Defender-Signature-Freshness Remediation ($defenderRetry+"`nfunction Update-MpSignature { [CmdletBinding()]param() [Console]::WriteLine('OP:UpdateSignature') }") 1 'FreshnessNotVerified' '' 'OP:(?!UpdateSignature)'

$certificate=@'
function Get-ChildItem { [CmdletBinding()]param($LiteralPath) [pscustomobject]@{HasPrivateKey=$true;NotAfter=(Get-Date).AddDays(3);EnhancedKeyUsageList=@([pscustomobject]@{ObjectId='1.3.6.1.5.5.7.3.2'})} }
'@
Invoke-PackageCase 'certificate-expiring' Certificates Network-Certificate-Expiry Detection $certificate 1 'CertificatesExpiring.*Count=1'
Invoke-PackageCase 'certificate-fresh' Certificates Network-Certificate-Expiry Remediation ($certificate.Replace('AddDays(3)','AddDays(30)')) 0 'NoExpiring'
Invoke-PackageCase 'certificate-unrelated-eku' Certificates Network-Certificate-Expiry Remediation ($certificate.Replace('1.3.6.1.5.5.7.3.2','1.3.6.1.5.5.7.3.1')) 0 'NoExpiring'
Invoke-PackageCase 'certificate-no-private-key' Certificates Network-Certificate-Expiry Remediation ($certificate.Replace('HasPrivateKey=$true','HasPrivateKey=$false')) 0 'NoExpiring'
Invoke-PackageCase 'certificate-notify' Certificates Network-Certificate-Expiry Remediation ($certificate+"`nfunction msg.exe { [Console]::WriteLine('OP:Notify'); `$global:LASTEXITCODE=0 }") 0 'Notified.*RenewalActionRequired' '' 'OP:(?!Notify)'
Invoke-PackageCase 'certificate-no-session' Certificates Network-Certificate-Expiry Remediation ($certificate+"`nfunction msg.exe { `$global:LASTEXITCODE=1 }") 0 'Deferred.*ActionRequired'
$certRenew=$certificate+@'

function Get-CimInstance { [CmdletBinding()]param($ClassName) [pscustomobject]@{PartOfDomain=$true} }
function Test-Path { [CmdletBinding()]param($LiteralPath) $true }
function Get-ItemProperty { [CmdletBinding()]param($LiteralPath) [pscustomobject]@{AEPolicy=1} }
function certutil.exe { [Console]::WriteLine('OP:AutoEnroll'); $global:LASTEXITCODE=0 }
'@
Invoke-PackageCase 'certificate-auto-enroll' Certificates Network-Certificate-Expiry Remediation $certRenew 0 'AutoEnrollmentRequested.*RenewalVerified=False' '-RequestAutoEnrollment' 'OP:(?!AutoEnroll)'
Invoke-PackageCase 'certificate-no-domain' Certificates Network-Certificate-Expiry Remediation ($certRenew+"`nfunction Get-CimInstance { [CmdletBinding()]param(`$ClassName) [pscustomobject]@{PartOfDomain=`$false} }") 0 'Deferred.*NotConfigured' '-RequestAutoEnrollment'

$bits=@'
$script:retried=$false
function Get-BitsTransfer { [CmdletBinding()]param($JobId) [pscustomobject]@{JobId=[guid]'11111111-2222-4333-8444-555555555555';JobState=$(if($script:retried){'Queued'}else{'Error'});ModificationTime=(Get-Date).AddDays(-2)} }
function Resume-BitsTransfer { [CmdletBinding()]param($BitsJob,[switch]$Asynchronous) if(-not $Asynchronous){throw 'Synchronous retry forbidden'}; $script:retried=$true;[Console]::WriteLine('OP:ResumeBits') }
function Remove-BitsTransfer { [CmdletBinding(SupportsShouldProcess=$true)]param($BitsJob) [Console]::WriteLine('OP:CancelBits') }
'@
Invoke-PackageCase 'bits-stale' BITS BITS-Stale-Jobs Detection $bits 1 'StaleErroredBitsJobs.*Count=1'
Invoke-PackageCase 'bits-retry' BITS BITS-Stale-Jobs Remediation $bits 0 'RetryAccepted=1 Cancelled=0 Unresolved=0' '' 'OP:(?!ResumeBits)'
Invoke-PackageCase 'bits-recent' BITS BITS-Stale-Jobs Remediation ($bits.Replace('AddDays(-2)','AddMinutes(-2)')) 0 'NoStaleErroredBitsJobs'
Invoke-PackageCase 'bits-suspended' BITS BITS-Stale-Jobs Remediation ($bits.Replace("'Error'","'Suspended'")) 0 'NoStaleErroredBitsJobs'
Invoke-PackageCase 'bits-explicit-cancel' BITS BITS-Stale-Jobs Remediation $bits 0 'Cancelled=1' '-CancelJobIds 11111111-2222-4333-8444-555555555555' 'OP:(?!CancelBits)'
Invoke-PackageCase 'bits-unresolved' BITS BITS-Stale-Jobs Remediation ($bits+"`nfunction Resume-BitsTransfer { [CmdletBinding()]param(`$BitsJob,[switch]`$Asynchronous) [Console]::WriteLine('OP:ResumeBits') }") 1 'BitsErrorsPersist' '' 'OP:(?!ResumeBits)'

$printing=@'
function Get-Service { [CmdletBinding()]param($Name) [pscustomobject]@{Status='Running'} }
function Get-CimInstance { [CmdletBinding()]param($ClassName,$Filter) [pscustomobject]@{StartMode='Auto'} }
function Get-Printer { [CmdletBinding()]param() [pscustomobject]@{Name='SyntheticPrinter'} }
function Get-PrintJob { [CmdletBinding()]param($PrinterName,$ID) [pscustomobject]@{ID=42;JobStatus='Error';SubmittedTime=(Get-Date).AddHours(-3)} }
function Restart-Service { [CmdletBinding()]param($Name) [Console]::WriteLine('OP:RestartSpooler') }
'@
Invoke-PackageCase 'print-blocked' Printing Print-Spooler-Health Detection $printing 1 'RemediationRequired.*Count=1'
Invoke-PackageCase 'print-disabled' Printing Print-Spooler-Health Remediation ($printing+"`nfunction Get-CimInstance { [CmdletBinding()]param(`$ClassName,`$Filter) [pscustomobject]@{StartMode='Disabled'} }") 0 'NotApplicable'
Invoke-PackageCase 'print-recent' Printing Print-Spooler-Health Remediation ($printing.Replace('AddHours(-3)','AddMinutes(-3)')) 0 'Healthy'
Invoke-PackageCase 'print-no-purge-default' Printing Print-Spooler-Health Remediation $printing 1 'BlockedJobsRemain' '' 'OP:(?!RestartSpooler)'
$purge=$printing+@'

$script:purged=$false
function Get-PrintJob { [CmdletBinding()]param($PrinterName,$ID) if(-not $script:purged){[pscustomobject]@{ID=42;JobStatus='Error';SubmittedTime=(Get-Date).AddHours(-3)}} }
function Remove-PrintJob { [CmdletBinding(SupportsShouldProcess=$true)]param($PrinterName,$ID) if($PrinterName -ne 'SyntheticPrinter' -or $ID -ne 42){throw 'Unexpected purge'}; $script:purged=$true;[Console]::WriteLine('OP:Purge42') }
'@
Invoke-PackageCase 'print-explicit-purge' Printing Print-Spooler-Health Remediation $purge 0 'SpoolerAndQueuesVerified' '-AllowPurge' 'OP:(?!(RestartSpooler|Purge42))'

$reboot=@'
$script:boot=(Get-Date).AddDays(-10)
function Get-CimInstance { [CmdletBinding()]param($ClassName) [pscustomobject]@{LastBootUpTime=$script:boot} }
function Test-Path { [CmdletBinding()]param($LiteralPath)
 if($LiteralPath -like 'HKLM:*'){return ($LiteralPath -like '*RebootPending')}
 Microsoft.PowerShell.Management\Test-Path -LiteralPath $LiteralPath
}
function Initialize-Observation {
 param($Package,[int]$Hours=80,[int]$ReminderHours=100)
 $dir=Join-Path $env:ProgramData ('SmartM365\IntuneRemediation\Output\'+$Package)
 [IO.Directory]::CreateDirectory($dir)|Out-Null
 $observation=[ordered]@{SchemaVersion=1;BootUtc=([datetimeoffset]$script:boot).ToUniversalTime().ToString('o');FirstObservedUtc=[datetimeoffset]::UtcNow.AddHours(-$Hours).ToString('o');Reasons=@('CBS');LastReminderUtc=[datetimeoffset]::UtcNow.AddHours(-$ReminderHours).ToString('o')}
 [IO.File]::WriteAllText((Join-Path $dir 'observation.json'),(ConvertTo-Json $observation -Compress))
}
'@
Invoke-PackageCase 'reboot-first-observation' Reboot Pending-Reboot-Age Detection $reboot 0 'Observing.*ObservedAgeHours=0'
Invoke-PackageCase 'reboot-aged' Reboot Pending-Reboot-Age Detection ($reboot+"`nInitialize-Observation 'Pending-Reboot-Age'") 1 'PendingRebootAged'
Invoke-PackageCase 'reboot-future-observation' Reboot Pending-Reboot-Age Detection ($reboot+"`nInitialize-Observation 'Pending-Reboot-Age' -24") 2 'Status=Error'
Invoke-PackageCase 'reboot-new-boot' Reboot Pending-Reboot-Age Detection ($reboot+"`nInitialize-Observation 'Pending-Reboot-Age'; `$script:boot=Get-Date") 0 'Observing.*ObservedAgeHours=0'
Invoke-PackageCase 'reboot-cleared' Reboot Pending-Reboot-Age Remediation ($reboot+"`nInitialize-Observation 'Pending-Reboot-Age'; function Test-Path { [CmdletBinding()]param(`$LiteralPath) if(`$LiteralPath -like 'HKLM:*'){return `$false}; Microsoft.PowerShell.Management\Test-Path -LiteralPath `$LiteralPath }") 0 'NoPendingReboot'
$corrupt=$reboot+"`nInitialize-Observation 'Pending-Reboot-Age'; [IO.File]::WriteAllText((Join-Path `$env:ProgramData 'SmartM365\IntuneRemediation\Output\Pending-Reboot-Age\observation.json'),'invalid-json')"
Invoke-PackageCase 'reboot-corrupt-json' Reboot Pending-Reboot-Age Detection $corrupt 2 'Status=Error'
Invoke-PackageCase 'reboot-cooldown' Reboot Pending-Reboot-Age Remediation ($reboot+"`nInitialize-Observation 'Pending-Reboot-Age' 80 1") 0 'ReminderCooldown'
Invoke-PackageCase 'reboot-reminder' Reboot Pending-Reboot-Age Remediation ($reboot+"`nInitialize-Observation 'Pending-Reboot-Age'; function msg.exe { [Console]::WriteLine('OP:Notify'); `$global:LASTEXITCODE=0 }") 0 'UserReminderSent.*RebootVerified=False' '' 'OP:(?!Notify)'
Invoke-PackageCase 'reboot-no-session' Reboot Pending-Reboot-Age Remediation ($reboot+"`nInitialize-Observation 'Pending-Reboot-Age'; function msg.exe { `$global:LASTEXITCODE=1 }") 0 'NoReminderDelivered'

$drm=$reboot+@'

Initialize-Observation 'Pending-Reboot-Age-DeviceRebootManager'
function Get-ScheduledTask { [CmdletBinding()]param($TaskPath,$TaskName)
 $gui=Join-Path $env:ProgramData 'SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1'
 [pscustomobject]@{State='Ready';Principal=[pscustomobject]@{RunLevel='Limited';GroupId='S-1-5-32-545';LogonType='Group'};Actions=@([pscustomobject]@{Execute=(Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe');Arguments=('-STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "'+$gui+'"')})}
}
function Get-AuthenticodeSignature { [CmdletBinding()]param($LiteralPath) [pscustomobject]@{Status='Valid'} }
function Start-ScheduledTask { [CmdletBinding()]param($InputObject) [Console]::WriteLine('OP:StartRebootManager') }
'@
Invoke-PackageCase 'drm-task-start' Reboot Pending-Reboot-Age-DeviceRebootManager Remediation $drm 0 'RebootManagerTaskStarted.*RebootVerified=False' '' 'OP:(?!StartRebootManager)'
Invoke-PackageCase 'drm-elevated-task' Reboot Pending-Reboot-Age-DeviceRebootManager Remediation ($drm.Replace("RunLevel='Limited'","RunLevel='Highest'")) 1 'Status=Error'
Invoke-PackageCase 'drm-invalid-signature' Reboot Pending-Reboot-Age-DeviceRebootManager Remediation ($drm.Replace("Status='Valid'","Status='HashMismatch'")) 1 'Status=Error'
Invoke-PackageCase 'drm-injected-action' Reboot Pending-Reboot-Age-DeviceRebootManager Remediation ($drm.Replace('-STA -NoProfile','-Command evil -STA -NoProfile')) 1 'Status=Error'
Invoke-PackageCase 'drm-running-task' Reboot Pending-Reboot-Age-DeviceRebootManager Remediation ($drm.Replace("State='Ready'","State='Running'")) 0 'AlreadyRunning'
Invoke-PackageCase 'drm-missing-task' Reboot Pending-Reboot-Age-DeviceRebootManager Remediation ($drm+"`nfunction Get-ScheduledTask { [CmdletBinding()]param(`$TaskPath,`$TaskName) throw 'Synthetic task missing' }") 1 'Status=Error'

$resultPath=Join-Path $TestRoot 'results.json'
if ($CaseFilter -eq '*' -and $PSVersionTable.PSVersion.Major -ge 7) {
    # Exercise the manager's actual discovery functions, without launching WPF or Graph.
    $guiPath=Join-Path $repoRoot 'SmartM365/Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1'
    $parseErrors=$null
    $guiAst=[Management.Automation.Language.Parser]::ParseFile($guiPath,[ref]$null,[ref]$parseErrors)
    if($parseErrors.Count){throw 'Manager source does not parse.'}
    foreach($functionName in @('Set-ObjectNoteProperty','Get-NormalizedScriptContent','Get-ContentSha256',
        'Get-ScriptContentSignature','Get-LocalPackageContentSignature','ConvertTo-CompareLines',
        'Get-ScriptHelpDescription','ConvertTo-PackageDisplayName','Get-PackageDescription',
        'Get-MatchingLocalRemediationPath','New-LocalPackageFromDetection','Get-DefaultRemediationRoot','Get-LocalPackages')) {
        $definition=$guiAst.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName},$true)
        if(-not $definition){throw "Manager function missing: $functionName"}
        . ([scriptblock]::Create($definition.Extent.Text))
    }
    Set-Variable -Name RemediationRoot -Scope Script -Value $packageRoot
    $localPackages=@(Get-LocalPackages)
    foreach($name in @('Windows-LAPS-Health','Windows-Time-Sync','Defender-Signature-Freshness',
        'Network-Certificate-Expiry','BITS-Stale-Jobs','Print-Spooler-Health',
        'Pending-Reboot-Age','Pending-Reboot-Age-DeviceRebootManager')) {
        $item=@($localPackages | Where-Object {$_.DisplayName -eq ('SmartM365-'+$name)})
        $ok=$item.Count -eq 1 -and -not $item[0].DetectionOnly -and
            -not [string]::IsNullOrWhiteSpace($item[0].RemediationPath) -and
            -not [string]::IsNullOrWhiteSpace($item[0].Description)
        $script:results.Add([pscustomobject]@{Case=('discovery-'+$name);Passed=$ok;Exit=0;ExpectedExit=0;Output='Manager package discovery'})
    }
}
[IO.File]::WriteAllText($resultPath,(ConvertTo-Json -InputObject $script:results.ToArray() -Depth 5))
$failed=@($script:results | Where-Object {-not $_.Passed})
Write-Output ('Cases={0} Passed={1} Failed={2} PowerShell={3} Evidence={4}' -f $script:results.Count,($script:results.Count-$failed.Count),$failed.Count,$PSVersionTable.PSVersion,$resultPath)
if($failed.Count){exit 1}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDeQfGNYiHt3bjc
# /N1VKTQbe7ESsgecZLONOB2q6rcmC6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINGuQ8TTPmDWzCp7tdzpdfoVuud4BETZrdzMbHsH+AC3MA0GCSqG
# SIb3DQEBAQUABIIBgEmdCh4PV+meeR7E7/yKEGHV2HdMEpyZB9y6fYCnSHJXCtVg
# IMhmQ8uM+ycabdt+j+Q12TMsD2OzmAQt9w9gZCzRieMOW9LIx7VaNb69s/6qr1zK
# 0bJBY6TIZhxYTygS2IDTfCYmK15pvaYaxAibdivDhhTKmh9IaAygCFLty4UqQ03X
# EXf2rMntCBdKh5Ddc5o9l7RXBfMY2ptYwYu23VI6+c5hHWVWxRmg5Pjs90KHnmKS
# 4uheIcvug4IRfj1DpU4d5F3+sZSej6PJTJl7qPVf+ISarM18uLtXZ4H0zW6xumwg
# 00rqKz0Wt14fVTWJrTBAtOTlwV5eIUDM4ds4ttZ7UGN5GLplTfsRmvgTSa7wKQwI
# gvBQvXIs30m5DFhE+RpLNWP0wuipngzKvH+c90ecB5r3ViEzx6OsUoCakbY7F5QC
# u6FHH8Dtl1iF+taeKjvfQ2vAkQKsUjuie+528cM21SkDsKfDxLz4JWqVAUa6sN2O
# cG9qlHxiIsk5L2Sf1KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzQ1
# MzFaMC8GCSqGSIb3DQEJBDEiBCD5VNvGLK+hW1N0M7D9rRfRycVyqDxXGhTr8l7d
# RRm4jTANBgkqhkiG9w0BAQEFAASCAgCtXWLVke0feL+neIDFyGj1BY4XQ41pma7b
# 5LaLnvrq5CFmfGxo+SbPRtKuH5/A6+4UuTWAfkg5oXV7b0xTwockIegS8oiL5NPp
# 5QGSlQ402tFwc5z+oaUcdgvcKKBzsFS/FGj7/wfjkjJYegl70Ew1U2lg3FB9iSSh
# mFh2sD9X2SvFfwoqa8gxXK+WS3sDRkFFTDuRCIiW169QL8MHlgpgMyJqux2VDc1C
# hZPTZ8WEEoowqn3Jn5MsTDZUKb/MWg7LF4Qqj4BoHv2ecK9xqrO7G6aTxZdADRIC
# pZ3cfYmi6b41V8stqpP4RcdBLODMGmjJjgAYx2FXtttqbv2dL4mz1V94ivDRVIbf
# yXJtFpadaxqjAq0wYJIDt1ggCNqkTen1ZdgdEQypR9K7GoR8U96D/W4wZY7d/4HL
# 0VKxdIRQGsWlxbCbb6pzq4ME1wRWvsA6qTrRlpih7NQ5JAFG9zeNzsMqtJxFKsMJ
# oL8TAzagWMy6ifMHBEcNVB2dpb3o9Nif+MrPuYrTeygqNDbw51LiFS/e9Dii3s68
# cU2vLcV5Hr/0aDCQQQqmaNdW4j981Wd90XwptWEs3rXKks2npeI3qo1oOG3svA5j
# B5GRwZL4fTtekasZWYLvUHuu5QUMAgfruyulSl/wsqCdNlCfpZYEgnYhuSgf7MqA
# nlf3aag5zA==
# SIG # End signature block
