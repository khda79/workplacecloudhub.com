<# .VERSION 1.4.2
Read-only diagnostics in the caller's own user context. No privileged helper or repair.
#>

function Get-DrtNamedRegistryValues {
    param([string[]]$Names,[scriptblock]$ValueReader)
    $values=@{}
    foreach ($name in $Names) {
        # The unnamed registry default is not a policy field. PS5.1 JSON cannot read an empty property name.
        if (-not [string]::IsNullOrEmpty($name)) { $values[$name]=& $ValueReader $name }
    }
    return $values
}

function Read-DrtRegistryEvidence {
    param([string]$Path)
    $base = $null; $key = $null
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        $key = $base.OpenSubKey($Path, $false)
        if ($null -eq $key) { return [pscustomobject]@{ State='Missing'; Path=$Path; Values=@{}; Children=@(); Detail='Key not present.' } }
        $values = Get-DrtNamedRegistryValues -Names $key.GetValueNames() -ValueReader {param($name);$key.GetValue($name)}
        return [pscustomobject]@{ State='Available'; Path=$Path; Values=$values; Children=@($key.GetSubKeyNames()); Detail='' }
    }
    catch {
        $state = if ($_.Exception -is [UnauthorizedAccessException] -or $_.Exception -is [Security.SecurityException]) { 'AccessDenied' } else { 'Unavailable' }
        return [pscustomobject]@{ State=$state; Path=$Path; Values=@{}; Children=@(); Detail=$_.Exception.Message }
    }
    finally { if ($key) { $key.Dispose() }; if ($base) { $base.Dispose() } }
}

function Read-DrtTaskEvidence {
    param([string]$Path,[switch]$Folder)
    try {
        $service = New-Object -ComObject Schedule.Service
        $service.Connect()
        if ($Folder) { $folderObject=$service.GetFolder($Path); return [pscustomobject]@{State='Available';Path=$Path;HasTasks=($folderObject.GetTasks(1).Count -gt 0);Detail='Only tasks visible to this account are inspected.'} }
        $task = $service.GetFolder((Split-Path $Path)).GetTask((Split-Path $Path -Leaf))
        return [pscustomobject]@{ State='Available'; Path=$Path; Enabled=[bool]$task.Enabled; Detail='' }
    }
    catch {
        $errorObject = $_.Exception
        while ($errorObject.InnerException) { $errorObject = $errorObject.InnerException }
        $code = $errorObject.HResult -band 0xffff
        # Scheduler can hide a task/folder rather than return AccessDenied to a standard user.
        $admin=([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $state = if ($code -in @(2,3) -and $admin) { 'Missing' } elseif ($code -in @(2,3)) { 'NotVisible' } elseif ($code -eq 5) { 'AccessDenied' } else { 'Unavailable' }
        return [pscustomobject]@{ State=$state; Path=$Path; Enabled=$null; Detail=$errorObject.Message }
    }
}

function Test-DrtMicrosoftLinkedEnrollment {
    param([string]$Provider,[string]$DiscoveryUrl,[string]$EnrollmentType)
    # Microsoft WinDC enrollment documentation: exact Microsoft endpoint plus provider/type.
    # This identifies a linked enrollment; it does not establish its health or primary Intune state.
    if (([string]$Provider).Trim() -ine 'Microsoft Device Management' -or ([string]$EnrollmentType).Trim() -notin @('26','0x1a')) { return $false }
    $uri=$null
    if (-not [Uri]::TryCreate($DiscoveryUrl,[UriKind]::Absolute,[ref]$uri)) { return $false }
    return ($uri.Scheme -eq 'https' -and $uri.DnsSafeHost -ieq 'discovery.dm.microsoft.com' -and $uri.IsDefaultPort -and -not $uri.UserInfo -and $uri.AbsolutePath.TrimEnd('/') -ieq '/EnrollmentConfiguration')
}

function Get-DrtMdmClassification {
    param([string]$Provider,[string]$DiscoveryUrl,[string]$EnrollmentType,[string]$Upn,[bool]$Omadm,[bool]$Task)
    $providerValue = ([string]$Provider).Trim()
    if ($providerValue -in @('Deploy Authority','Cloud Authority','Local Authority')) { return 'WindowsInternal' }
    if ($providerValue -eq 'WMI_Bridge_SCCM_Server') { return 'ManagementBridge' }
    if ($providerValue -eq 'MS DM Server') { return 'Intune' }
    if (Test-DrtMicrosoftLinkedEnrollment -Provider $providerValue -DiscoveryUrl $DiscoveryUrl -EnrollmentType $EnrollmentType) { return 'MicrosoftLinkedEnrollment' }
    $intuneUrl = $DiscoveryUrl -match '(?i)^https://enrollment\.manage\.microsoft\.com(?:/|$)'
    if ($intuneUrl) {
        if (-not $providerValue -and $EnrollmentType -eq '6' -and $Upn -match '(?i)^fooUser@[^@\s]+$' -and -not $Omadm -and -not $Task) { return 'IncompleteIntuneCandidate' }
        return 'UnconfirmedIntuneCandidate'
    }
    if (($DiscoveryUrl -and $DiscoveryUrl -match '^https?://') -or ($providerValue -and ($Omadm -or $Task))) { return 'OtherMdm' }
    return 'UnconfirmedOrInternal'
}

function Get-DrtEnrollmentEvidence {
    $root = Read-DrtRegistryEvidence 'SOFTWARE\Microsoft\Enrollments'
    $entries = @(); $limitations = @()
    if ($root.State -notin @('Available','Missing')) { $limitations += $root }
    foreach ($child in $root.Children) {
        if ($child -notmatch '^[{]?[0-9a-fA-F-]{36}[}]?$') { continue }
        $item = Read-DrtRegistryEvidence ($root.Path+'\'+$child)
        if ($item.State -ne 'Available') { $limitations += $item; continue }
        $v = $item.Values
        $provider = [string]$v['ProviderID']; $url = [string]$v['DiscoveryServiceFullURL']
        $oma = $false; $task = $false; $entryCoverage = 'Complete'
        if (($provider -and $provider -notin @('Deploy Authority','Cloud Authority','Local Authority','WMI_Bridge_SCCM_Server')) -or $url) {
            $account = Read-DrtRegistryEvidence ('SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\'+$child)
            $oma = $account.State -eq 'Available'
            # Task visibility is limited for standard users; never turn a refusal into absence.
            $taskEvidence = Read-DrtTaskEvidence ('\Microsoft\Windows\EnterpriseMgmt\'+$child) -Folder
            $task = $taskEvidence.State -eq 'Available' -and $taskEvidence.HasTasks
            if ($taskEvidence.State -eq 'Available' -and -not $task) {
                $taskEvidence=[pscustomobject]@{State='NotVisible';Path=$taskEvidence.Path;Detail='No task visible in the folder; this does not prove absence.'}
            }
            foreach ($e in @($account,$taskEvidence)) {
                if ($e.State -notin @('Available','Missing')) { $limitations += $e; $entryCoverage = 'Partial' }
            }
        }
        $classification = Get-DrtMdmClassification -Provider $provider -DiscoveryUrl $url -EnrollmentType ([string]$v['EnrollmentType']) -Upn ([string]$v['UPN']) -Omadm $oma -Task $task
        if ($classification -eq 'IncompleteIntuneCandidate' -and $entryCoverage -ne 'Complete') { $classification = 'UnconfirmedIntuneCandidate' }
        $entries += [pscustomobject]@{ EnrollmentId=$child; Provider=$provider; DiscoveryUrl=$url; EnrollmentType=[string]$v['EnrollmentType']; Classification=$classification; Coverage=$entryCoverage; OmadmAccount=$oma; TaskVisible=$task }
    }
    $state = if (@($entries | Where-Object Classification -eq 'Intune').Count) { 'Detected' }
        elseif ($limitations.Count) { 'Unknown' }
        elseif (@($entries | Where-Object Classification -match 'IntuneCandidate|OtherMdm|MicrosoftLinkedEnrollment').Count) { 'Unconfirmed' }
        else { 'NotDetected' }
    return [pscustomobject]@{ State=$state; Coverage=if ($limitations.Count) {'Partial'} else {'Complete'}; Entries=$entries; Limitations=$limitations; Source='Local registry and visible tasks; no cloud confirmation' }
}

function Convert-DrtUtcDate {
    param([string]$Value)
    if ($Value -notmatch '(?i)(UTC|GMT|Z|[+-]\d\d:\d\d)\s*$') { return $null }
    $date = [DateTimeOffset]::MinValue
    $text = $Value.Trim() -replace '\s+(UTC|GMT)$','Z'
    if ([DateTimeOffset]::TryParse($text,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AllowWhiteSpaces,[ref]$date)) { return $date.ToUniversalTime() }
    return $null
}

function ConvertFrom-DrtUserStatus {
    param([string[]]$Lines,[string]$CurrentUpn,[string]$CurrentAccount,[DateTimeOffset]$Now=[DateTimeOffset]::UtcNow)
    $fields=@{}; $attempts=@(); $attempt=$null
    foreach ($line in $Lines) {
        if ($line -match '^\s*\|\s*([^|]+)\s*\|\s*$') { $attempt=$null; continue }
        if ($line -notmatch '^\s*([^:]+?)\s*:\s*(.*?)\s*$') { continue }
        $key=$matches[1].Trim(); $value=$matches[2].Trim()
        if ($key -in @('AcquirePrtDiagnostics','RefreshPrtDiagnostics')) {
            $attempt=[ordered]@{ Kind=$key; Presence=$value }; $attempts += $attempt; continue
        }
        if ($attempt -and $key -in @('Previous Prt Attempt','Attempt Status','User Identity','HTTP Error','HTTP status','Server Error Code','Server Error Description','Correlation ID')) { $attempt[$key]=$value }
        elseif ($key -in @('AzureAdPrt','AzureAdPrtUpdateTime','AzureAdPrtExpiryTime','EnterprisePrt','CloudTgt','OnPremTgt','WorkplaceJoined','WamDefaultSet','WamDefaultAuthority','NgcSet','IsUserAzureAD','SessionIsNotRemote')) { $fields[$key]=$value }
    }
    $diagnostics=@()
    foreach ($a in $attempts) {
        $time=Convert-DrtUtcDate ([string]$a['Previous Prt Attempt'])
        $sameUser=([string]$a['User Identity'] -and ([string]$a['User Identity'] -ieq $CurrentUpn -or [string]$a['User Identity'] -ieq $CurrentAccount))
        $recent=$null -ne $time -and $time -ge $Now.AddHours(-24) -and $time -le $Now.AddMinutes(5)
        $failed=([string]$a['Attempt Status'] -and [string]$a['Attempt Status'] -notmatch '^(0x)?0+$') -or ([string]$a['HTTP Error'] -and [string]$a['HTTP Error'] -notmatch '^(0x)?0+$') -or ([string]$a['HTTP status'] -and [string]$a['HTTP status'] -ne '200')
        $diagnostics += [pscustomobject]@{ Kind=$a['Kind']; User=$a['User Identity']; Timestamp=$a['Previous Prt Attempt']; SameUser=[bool]$sameUser; Recent=[bool]$recent; Failed=[bool]$failed; RelevantFailure=[bool]($sameUser -and $recent -and $failed); AttemptStatus=$a['Attempt Status']; HttpStatus=$a['HTTP status']; HttpError=$a['HTTP Error']; ServerError=$a['Server Error Code'] }
    }
    $expiry=Convert-DrtUtcDate ([string]$fields['AzureAdPrtExpiryTime'])
    $state=if (-not $fields['AzureAdPrt']) {'Unknown'} elseif ($fields['AzureAdPrt'] -ne 'YES') {'Missing'} elseif ($expiry -and $expiry -le $Now) {'Expired'} elseif ($expiry -and $expiry -le $Now.AddHours(4)) {'Expiring'} elseif (@($diagnostics|Where-Object RelevantFailure).Count) {'AttentionRequired'} else {'Present'}
    return [pscustomobject]@{ State=$state; Fields=$fields; Attempts=$diagnostics; CurrentUser=$CurrentAccount; CurrentUpn=$CurrentUpn; Context='Current process user; no SYSTEM helper'; RefreshAttempted=$false }
}

function Test-DrtEndpointConnectivity {
    param([string]$HostName,[int]$TimeoutMs=1500)
    $dns=$null; $client=$null; $tcp=$null; $addresses=$null
    try {
        $dns=[Net.Dns]::BeginGetHostAddresses($HostName,$null,$null)
        if (-not $dns.AsyncWaitHandle.WaitOne($TimeoutMs)) { return [pscustomobject]@{ Host=$HostName; Dns='Timeout'; Tcp443='NotTested'; Detail='DNS timeout in user context.' } }
        $addresses=[Net.Dns]::EndGetHostAddresses($dns)
        if (-not $addresses.Count) { throw 'No DNS address returned.' }
        # Use resolved addresses: BeginConnect(hostname) could start a second unbounded DNS lookup.
        $client=New-Object Net.Sockets.TcpClient
        $tcp=$client.BeginConnect($addresses,443,$null,$null)
        if (-not $tcp.AsyncWaitHandle.WaitOne($TimeoutMs)) { return [pscustomobject]@{ Host=$HostName; Dns='Resolved'; Tcp443='Timeout'; Detail='TCP 443 timeout in user context.' } }
        $client.EndConnect($tcp)
        return [pscustomobject]@{ Host=$HostName; Dns='Resolved'; Tcp443='Reachable'; Detail='TCP only; does not validate TLS, proxy, authentication or SYSTEM connectivity.' }
    }
    catch { return [pscustomobject]@{ Host=$HostName; Dns=if ($addresses) {'Resolved'} else {'Unavailable'}; Tcp443=if ($addresses) {'Unavailable'} else {'NotTested'}; Detail=$_.Exception.Message } }
    finally { if ($client) {$client.Close()}; if ($dns) {$dns.AsyncWaitHandle.Close()}; if ($tcp) {$tcp.AsyncWaitHandle.Close()} }
}

function Get-DrtRecentEventEvidence {
    param([string]$LogName,[datetime]$Since)
    try {
        $events=@(Get-WinEvent -FilterHashtable @{LogName=$LogName;StartTime=$Since;Level=@(2,3)} -MaxEvents 40 -ErrorAction Stop)
        $items=foreach ($e in $events) {
            $message=[string]$e.Message
            $code=if ($message -match '(?i)AADSTS5000611|Derivation Function version .KDFV1. is invalid') {'KdfUpdateRequired'} elseif ($message -match '(?i)0x80180018|\(UserLicense\)') {'LicenseRejectionReported'} elseif ($message -match '(?i)0x8018000A') {'EnrollmentConflictReported'} else {'DiagnosticEvent'}
            [pscustomobject]@{ Id=$e.Id; RecordId=$e.RecordId; TimeUtc=$e.TimeCreated.ToUniversalTime().ToString('o'); UserSid=[string]$e.UserId; Category=$code; Message=$message }
        }
        return [pscustomobject]@{ Log=$LogName; State='Available'; Events=@($items); Detail='Recent local events; may concern the machine or another user. Not proof of current license assignment.' }
    }
    catch {
        if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { return [pscustomobject]@{Log=$LogName;State='Available';Events=@();Detail='No matching events in the selected window.'} }
        return [pscustomobject]@{Log=$LogName;State='Unavailable';Events=@();Detail=$_.Exception.Message}
    }
}

function Get-DrtPolicyEvidence {
    $mdm=Read-DrtRegistryEvidence 'SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM'
    $workplace=Read-DrtRegistryEvidence 'SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin'
    $task=Read-DrtTaskEvidence '\Microsoft\Windows\Workplace Join\Automatic-Device-Join'
    $mdmState=if ($mdm.State -notin @('Available','Missing')) {'Unknown'} elseif ([string]$mdm.Values['AutoEnrollMDM'] -eq '1') {'Enabled'} else {'MissingOrDisabled'}
    $workState=if ($workplace.State -notin @('Available','Missing')) {'Unknown'} elseif ([string]$workplace.Values['autoWorkplaceJoin'] -eq '1') {'Enabled'} else {'MissingOrDisabled'}
    return [pscustomobject]@{ MdmState=$mdmState; WorkplaceState=$workState; Mdm=$mdm; Workplace=$workplace; AutomaticJoinTask=$task }
}

function Get-DrtUserDiagnostics {
    param([psobject]$Result,[string[]]$StatusLines,[DateTimeOffset]$Now=[DateTimeOffset]::UtcNow,[psobject]$EnrollmentEvidence)
    $sso=ConvertFrom-DrtUserStatus -Lines $StatusLines -CurrentUpn $Result.CurrentUserUpn -CurrentAccount $Result.CurrentUser -Now $Now
    $enrollment=if ($EnrollmentEvidence) {$EnrollmentEvidence} else {Get-DrtEnrollmentEvidence}
    $policies=Get-DrtPolicyEvidence
    $hosts=@('login.microsoftonline.com','device.login.microsoftonline.com','enterpriseregistration.windows.net','enrollment.manage.microsoft.com')
    foreach ($url in @($Result.MdmUrl)) { $uri=$null; if ([Uri]::TryCreate([string]$url,[UriKind]::Absolute,[ref]$uri) -and $uri.Scheme -eq 'https' -and $uri.DnsSafeHost -and $uri.DnsSafeHost -notin $hosts) { $hosts += $uri.DnsSafeHost } }
    $network=@(foreach ($hostName in ($hosts|Select-Object -First 6)) {Test-DrtEndpointConnectivity $hostName})
    $events=@(foreach ($log in @('Microsoft-Windows-AAD/Operational','Microsoft-Windows-User Device Registration/Admin','Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin')) {Get-DrtRecentEventEvidence -LogName $log -Since $Now.UtcDateTime.AddHours(-48)})
    $registration=if ($Result.AzureAdJoined -eq 'YES' -and ($Result.DeviceAuthStatus -like '*SUCCESS*' -or (-not $Result.DeviceAuthStatus -and $Result.KeySignTest -eq 'PASSED'))) {'Healthy'} elseif (-not $Result.AzureAdJoined) {'Unknown'} else {'AttentionRequired'}
    $partial=$enrollment.Coverage -eq 'Partial' -or $policies.MdmState -eq 'Unknown' -or $policies.WorkplaceState -eq 'Unknown' -or $policies.AutomaticJoinTask.State -notin @('Available','Missing') -or @($events|Where-Object State -ne 'Available').Count -gt 0
    return [pscustomobject]@{ SchemaVersion=1; CollectedUtc=$Now.ToString('o'); Context='User'; RegistrationState=$registration; Sso=$sso; Enrollment=$enrollment; Policies=$policies; Connectivity=$network; Events=$events; Coverage=if ($partial) {'Partial'} else {'Complete'}; Advice=@('User: review account, PRT dates and network results; share evidence with support if needed.','Support: verify Intune licensing, MDM scope and machine policies. No automatic enrollment or repair is requested.') }
}

function Add-DrtUserDiagnostics {
    param([psobject]$Result,[string]$OutputDir,[psobject]$EnrollmentEvidence)
    $lines=@()
    if ($Result.DsregSnapshotPath -and (Test-Path -LiteralPath $Result.DsregSnapshotPath)) { $lines=@(Get-Content -LiteralPath $Result.DsregSnapshotPath -Encoding UTF8) }
    $d=Get-DrtUserDiagnostics -Result $Result -StatusLines $lines -EnrollmentEvidence $EnrollmentEvidence
    $path=Join-Path $OutputDir ('UserDiagnostics_'+$Result.RunId+'.json')
    $d|ConvertTo-Json -Depth 12|Set-Content -LiteralPath $path -Encoding UTF8
    foreach ($pair in @(@('UserDiagnostics',$d),@('UserDiagnosticsPath',$path),@('RegistrationHealth',$d.RegistrationState),@('UserSsoState',$d.Sso.State),@('LocalMdmState',$d.Enrollment.State),@('DiagnosticCoverage',$d.Coverage))) { $Result|Add-Member NoteProperty $pair[0] $pair[1] -Force }
    $Result.IntuneEnrolled=if ($d.Enrollment.State -eq 'Detected') {$true} elseif ($d.Enrollment.State -in @('Unknown','Unconfirmed')) {$null} else {$false}
    $Result.MdmNonIntuneEnrollmentDetected=@($d.Enrollment.Entries|Where-Object Classification -eq 'OtherMdm').Count -gt 0
    $Result.MdmNonIntuneEnrollmentIds=(@($d.Enrollment.Entries|Where-Object Classification -eq 'OtherMdm'|ForEach-Object EnrollmentId) -join ';')
    $linkedIds=(@($d.Enrollment.Entries|Where-Object Classification -eq 'MicrosoftLinkedEnrollment'|ForEach-Object EnrollmentId) -join ';')
    $Result|Add-Member NoteProperty MdmMicrosoftLinkedEnrollmentIds $linkedIds -Force
    $Result.MdmUnconfirmedIntuneEnrollmentIds=(@($d.Enrollment.Entries|Where-Object Classification -match 'IntuneCandidate'|ForEach-Object EnrollmentId) -join ';')
    if ($d.Enrollment.State -eq 'Unknown') {
        $Result.IntuneEnrolled=$null; $Result.EnrollmentConfidence='Unknown'; $Result.EnrollmentStatusMessage='Not verifiable with this account; local evidence is incomplete.'
        if ($d.RegistrationState -eq 'Healthy') { $Result.Message='Device registration is healthy. Local Intune enrollment cannot be verified with this account.' }
    }
    elseif ($d.Enrollment.State -eq 'Unconfirmed') { $Result.EnrollmentConfidence='Unconfirmed'; $Result.EnrollmentStatusMessage='Unconfirmed local MDM traces; support review required.' }
    if ($d.Policies.MdmState -eq 'Unknown') { $Result.MdmPolicyStatus='Unknown'; $Result.MdmPolicyStatusMessage='MDM policy is not verifiable with this account.' }
    if ($d.Policies.AutomaticJoinTask.State -notin @('Available','Missing')) { $Result.AutomaticDeviceJoinTaskPresent=$null; $Result.AutomaticDeviceJoinTaskState='Unknown' }
    $Result.NextAction=if (@($d.Connectivity|Where-Object Tcp443 -ne 'Reachable').Count) {'CHECK_USER_NETWORK'} elseif ($d.Sso.State -in @('Missing','Expired','Expiring','AttentionRequired')) {'REVIEW_USER_SIGNIN'} elseif ($d.Enrollment.State -in @('Unknown','Unconfirmed')) {'CONTACT_SUPPORT_VERIFY_EVIDENCE'} elseif ($d.Enrollment.State -ne 'Detected') {'CONTACT_SUPPORT_VERIFY_INTUNE'} else {'NONE'}
}

function Sync-DrtConfigDefaults {
    param([psobject]$Config,[psobject]$Template,[Collections.Generic.List[string]]$Added,[string]$Prefix='')
    foreach ($p in $Template.PSObject.Properties) {
        $name=$Prefix+$p.Name
        $existing=$Config.PSObject.Properties[$p.Name]
        if (-not $existing) { $Config|Add-Member NoteProperty $p.Name $p.Value; $Added.Add($name) }
        elseif ($p.Value -is [pscustomobject] -and $existing.Value -is [pscustomobject]) { Sync-DrtConfigDefaults -Config $existing.Value -Template $p.Value -Added $Added -Prefix ($name+'.') }
    }
}

function Format-DrtUserDiagnostics {
    param([psobject]$Diagnostics)
    if (-not $Diagnostics) { return '' }
    $d=$Diagnostics
    $lines=@('','--- User diagnostics / read-only ---',('Registration: {0} | User SSO: {1} | Local Intune: {2}' -f $d.RegistrationState,$d.Sso.State,$d.Enrollment.State),('Coverage: {0}; inaccessible evidence is unknown, not absent.' -f $d.Coverage),'','[User PRT / SSO]')
    foreach ($key in @('AzureAdPrt','AzureAdPrtUpdateTime','AzureAdPrtExpiryTime','EnterprisePrt','CloudTgt','OnPremTgt','WorkplaceJoined','WamDefaultSet','WamDefaultAuthority','NgcSet')) { $lines += ('{0}: {1}' -f $key,$d.Sso.Fields[$key]) }
    foreach ($a in $d.Sso.Attempts) { $lines+=('{0}: User={1}; Time={2}; SameUser={3}; Recent={4}; RelevantFailure={5}; Status={6}; HTTP={7}; Error={8}' -f $a.Kind,$a.User,$a.Timestamp,$a.SameUser,$a.Recent,$a.RelevantFailure,$a.AttemptStatus,$a.HttpStatus,$a.ServerError) }
    $lines+=@('','[Connectivity: current user, DNS and TCP 443 only]')
    foreach ($c in $d.Connectivity) { $lines+=('{0}: DNS={1}; TCP443={2}; {3}' -f $c.Host,$c.Dns,$c.Tcp443,$c.Detail) }
    $lines+=@('','[Local MDM classification]')
    foreach ($e in $d.Enrollment.Entries) { $lines+=('{0}: {1}; Provider={2}; Coverage={3}' -f $e.EnrollmentId,$e.Classification,$e.Provider,$e.Coverage) }
    foreach ($e in $d.Enrollment.Limitations) { $lines+=('Not verifiable: {0}; {1}; {2}' -f $e.Path,$e.State,$e.Detail) }
    $lines+=@('','[Read-only policy checks]',('MDM={0}; WorkplaceJoin={1}; AutomaticJoinTask={2}' -f $d.Policies.MdmState,$d.Policies.WorkplaceState,$d.Policies.AutomaticJoinTask.State),'','[Recent accessible events / last 48 hours]')
    foreach ($log in $d.Events) {
        $lines+=('{0}: {1}; {2} event(s); {3}' -f $log.Log,$log.State,@($log.Events).Count,$log.Detail)
        foreach ($e in $log.Events) { $lines+=('  {0}; Id={1}; Record={2}; UserSid={3}; {4}' -f $e.TimeUtc,$e.Id,$e.RecordId,$e.UserSid,$e.Category) }
    }
    $lines+=@('','[Next steps - guidance only]')+$d.Advice
    return $lines -join [Environment]::NewLine
}

function Get-DrtSummaryLines {
    param([psobject]$Diagnostics)
    $d=$Diagnostics
    if (-not $d) { return @() }
    $fr=$script:SelectedLanguage -like 'fr*'
    $lines=@()
    if ($fr) {
        $lines+='[INFO] Enregistrement : '+$d.RegistrationState+' | SSO utilisateur : '+$d.Sso.State+' | Intune local : '+$d.Enrollment.State
        $lines+='[INFO] Couverture : '+$d.Coverage+'. Toute preuve inaccessible reste non verifiable.'
    } else {
        $lines+='[INFO] Registration: '+$d.RegistrationState+' | User SSO: '+$d.Sso.State+' | Local Intune: '+$d.Enrollment.State
        $lines+='[INFO] Coverage: '+$d.Coverage+'. Inaccessible evidence is unknown, not absent.'
    }
    foreach ($a in @($d.Sso.Attempts|Where-Object Failed)) { $lines+=if ($a.RelevantFailure) {'[WARNING] Recent PRT failure for the current account: '+$a.AttemptStatus} else {'[INFO] Historical or other-account PRT failure; not attributed to this session: '+$a.User} }
    if (@($d.Enrollment.Entries|Where-Object Classification -eq 'MicrosoftLinkedEnrollment').Count) {
        $lines+=if($fr){'[INFO] Inscription Microsoft liee WinDC/MMP-C reconnue. Presence uniquement ; synchronisation et bon fonctionnement non verifies.'}else{'[INFO] Microsoft linked enrollment (WinDC/MMP-C) recognized. Presence only; synchronization and health not checked.'}
    }
    if ($d.Sso.State -in @('Missing','Expired','Expiring','AttentionRequired')) { $lines+='[WARNING] User PRT: '+$d.Sso.State+'. Review sign-in with support; no automatic refresh.' }
    foreach ($c in $d.Connectivity) { $lines+=if ($c.Tcp443 -eq 'Reachable') {'[INFO] User DNS/TCP443 reachable: '+$c.Host} else {'[WARNING] User connectivity: '+$c.Host+'; DNS='+$c.Dns+'; TCP443='+$c.Tcp443} }
    foreach ($log in $d.Events) {
        if ($log.State -ne 'Available') { $lines+='[INFO] Event log not verifiable with this account: '+$log.Log }
        foreach ($e in @($log.Events|Where-Object Category -ne 'DiagnosticEvent')) { $lines+='[WARNING] Local event '+$e.Category+' at '+$e.TimeUtc+'; ask support to correlate account and current state.' }
    }
    return $lines
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC8gbgwQe0v5iFW
# DCWBX+FTnsryFzCnygstI7sHjI8YK6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILVkDcO1ZWCcJlrf1UIEUmuiUALaEWCQk4e/lJaCEl7EMA0GCSqG
# SIb3DQEBAQUABIIBgKZZSPa1n5PJpGpujh45Q0rN8JhWYzIt1FoK96vXRuiXsXN1
# J0GBGOjGWNDUKwf/1DHUqjX1+iuL2pOsGWCiRqIha2LRkt4DOFj/RDOPDmFjmQI9
# vbRfWSXJL8n3veC2Z++r2L5phuBXxjRiuXooBHRAhPHyfFjXJ+opMwuFNIj8pHkv
# 75zJ1iA2O2nGcHtpWO6dzkQvozx/EXLmQZsamWLAZiJxCNfakf7f3aTnSM9Zuz1L
# 0canm9OBTgTlKesXp0Lu+Zp0W2O4u38DQjsVG5sH8fscZcLJ/cae32QlF41ZIv1r
# Et3fNqDb6TyUqFkBJwtud3T6Vyh42snzta7E8Xum7S5bpOu1d/A6R6PlJKfJZRAl
# ggs1mm9LggqEJ7h14wxOUO0MQIYOMuBgLrrtnGFJbR4/caRZRWc50rstCooV0ltK
# G/619Nn6tS5JL8XE+0h7il7TeEGcNdJQwBgSbEU6rpANLNUUnYndH7zvIWhltfyR
# yfc/Qwhbsndj61fzlaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTIyMDAy
# MTdaMC8GCSqGSIb3DQEJBDEiBCAaViQRBKX25ZY5PVvYrUsC3cQuY/G0/tbyvvj0
# AHNU1zANBgkqhkiG9w0BAQEFAASCAgCaMIIg/lqKbcBkUnPchQxDtbnqQOMh6B+B
# 8u4FqyrdE1x0bgGBPveB+9TjxBJfOCZ7SR/WFqncDhrKRl5cRTVyXNR5vjHAaG8S
# ItkoMnxisjkScdbhpqEJBxbcbyz3csKC9JtJS+llWecNt/uQA/1jpze4e6unlJvX
# l4577a99DGxsAfkouAgwnyf5wjz1DjqpNl5LqinO08LMmULVZ0JZtnsgZ9vuYTYs
# zDbU6wVVD9UH1nz3nYiNrWUPmCf0Gfnx9tUm/mz+GmD/ywblNMFe9cejFu0n5/Fs
# D6WUP4nqFQTdyX5hqGGwcn402jmfjf3Ab5limflNdGngOPVVTJfyGQIqRCdg4yRn
# CQMzufr1i3FwINwU3o5QBsj79cZVDUqhoJ9sEKffDok1MrF76v59Qii5PE2Eotdt
# fVtWTrvoFBfKQGF5UcSLQEdoIK7y3xruXyDrl7cM8Ak3roqCzcCRhrbLuRkjPxRI
# lkwmSDmtbcsELSPvJrS9HAU+BfQaVXBFCAwqZq9jk8smllEC2BDyISCclK+KPtUF
# WEjXU+JpyKzSTdw37PDD9VyAgN8G2RT8+7L/DvUR9rWMUtSoU0t8RSeeqJGLjMUQ
# 3kkuiuBw8qtnmOGxgZeg58tQGRrfm5V8XgDitIiQvgpihlFFGMJX9yXPzdsY8MCj
# eIn6Dn42yg==
# SIG # End signature block
