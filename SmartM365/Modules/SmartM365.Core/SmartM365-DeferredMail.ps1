<# Shared deferred report delivery. Disabled unless explicitly configured. #>
# Version: 1.4.0
# A scheduled report is eligible for its current Paris calendar day only.
function Get-SmartM365ReportMailCadenceDay {
    param($Config, [datetimeoffset]$Now)
    if ($null -eq $Config) { return }
    if ($Config -is [Collections.IDictionary]) {
        if (-not $Config.Contains('ReportMailCadence')) { return }
        $policy = $Config['ReportMailCadence']
    } else {
        $property = $Config.PSObject.Properties['ReportMailCadence']
        if (-not $property) { return }
        $policy = $property.Value
    }
    if ($policy -isnot [pscustomobject] -and $policy -isnot [Collections.IDictionary]) {
        throw 'ReportMailCadence must be an object.'
    }
    $mode = if ($policy -is [Collections.IDictionary]) { $policy['Mode'] } else {
        $property = $policy.PSObject.Properties['Mode']
        if ($property) { $property.Value }
    }
    $local = [TimeZoneInfo]::ConvertTime($Now, [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time'))
    if ($mode -notin @('Daily','SelectedDays')) { throw 'ReportMailCadence.Mode must be Daily or SelectedDays.' }
    if ($mode -eq 'SelectedDays') {
        $configuredDays = if ($policy -is [Collections.IDictionary]) { ,$policy['Days'] } else {
            $property = $policy.PSObject.Properties['Days']
            if ($property) { ,$property.Value }
        }
        $days = @($configuredDays)
        $names = [Enum]::GetNames([DayOfWeek])
        if ($configuredDays -isnot [array] -or $days.Count -lt 1 -or $days.Count -ne @($days | Sort-Object -Unique).Count -or @($days | Where-Object { $_ -isnot [string] -or $_ -notin $names }).Count) {
            throw 'ReportMailCadence.Days must be a non-empty array of unique English day names.'
        }
        if ([string]$local.DayOfWeek -notin $days) { return '' }
    }
    return $local.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}

function Test-SmartM365ReportMailDeliveryDate {
    param([string]$Date, [datetimeoffset]$Now)
    if (-not $Date) { return $true }
    return $Date -ceq (Get-SmartM365ReportMailCadenceDay -Config @{ReportMailCadence=@{Mode='Daily'}} -Now $Now)
}

function Enter-SmartM365ReportMailCadence {
    param($Config, [string]$Key, [datetimeoffset]$Now = (Get-SmartM365MailUtcNow))
    $day = Get-SmartM365ReportMailCadenceDay -Config $Config -Now $Now
    if ($null -eq $day) { return }
    if (-not $day) {
        WriteLog -Message ("Report mail {0} skipped: today is outside ReportMailCadence (Europe/Paris)." -f $Key) -Level INFO
        return @{ Skip = $true }
    }

    # The log root is shared by collectors of the same tenant. Lock the scoped
    # state through submission so concurrent hosts cannot send the same report.
    $effective = Get-SmartM365EffectiveModuleGlobalConfig
    $tenant = [string]$effective.ProfileKey
    $logRoot = [string](Resolve-SmartM365ConfigValue $effective.LogAllRootPath)
    if ($tenant -notmatch '^[a-zA-Z0-9][a-zA-Z0-9._-]*$' -or -not [IO.Path]::IsPathRooted($logRoot) -or $logRoot -match '\{\{') {
        throw 'Report mail cadence requires a tenant profile and its configured absolute LogAllRootPath.'
    }
    if ($Key -notmatch '^[a-zA-Z0-9][a-zA-Z0-9._:-]{0,319}$') { throw 'Invalid report mail cadence key.' }
    $root = Join-Path $logRoot 'ReportMailCadence'
    [void][IO.Directory]::CreateDirectory($root)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $id = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(("$tenant|$Key").ToLowerInvariant())))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
    $path = Join-Path $root "$id.state"
    try { $stream = [IO.File]::Open($path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
    catch [IO.IOException] {
        if (($_.Exception.HResult -band 0xFFFF) -notin @(32,33)) { throw }
        WriteLog -Message ("Report mail {0} skipped: another process holds its submission lock." -f $Key) -Level INFO
        return @{ Skip = $true; Busy = $true }
    }
    try {
        if ($stream.Length -gt 0) {
            if ($stream.Length -ne 10) { throw 'Report mail cadence state is invalid.' }
            $buffer = New-Object byte[] 10
            if ($stream.Read($buffer, 0, $buffer.Length) -ne $buffer.Length) { throw 'Report mail cadence state could not be read.' }
            $last = [Text.Encoding]::ASCII.GetString($buffer)
            if ($last -notmatch '^\d{4}-\d{2}-\d{2}$') { throw 'Report mail cadence state is invalid.' }
            if ($last -ge $day) {
                WriteLog -Message ("Report mail {0} skipped: already submitted or pending today (Europe/Paris)." -f $Key) -Level INFO
                return @{ Skip = $true }
            }
        }
        $result = @{ Skip = $false; Date = $day; Stream = $stream }
        $stream = $null
        return $result
    } finally {
        if ($stream) { $stream.Dispose() }
    }
}

function Set-SmartM365ReportMailSubmission {
    param($Cadence)
    # Reserve before transport to avoid sending twice after an ambiguous
    # transport failure or an interruption immediately after acceptance.
    if (-not $Cadence) { return }
    $bytes = [Text.Encoding]::ASCII.GetBytes($Cadence.Date)
    $Cadence.Stream.Position = 0
    $Cadence.Stream.Write($bytes, 0, $bytes.Length)
    $Cadence.Stream.Flush($true)
}

function Get-SmartM365MailUtcNow { return [datetimeoffset]::UtcNow }
function Get-SmartM365DeferredMailSettings {
    $config = Get-SmartM365EffectiveModuleGlobalConfig
    $p = $config.PSObject.Properties['ReportMailSchedule']
    if (-not $p -or -not $p.Value) { return $null }
    $settings = $p.Value
    if ($settings.Enabled -isnot [bool]) { throw 'ReportMailSchedule.Enabled must be Boolean.' }
    if (-not $settings.Enabled) { return $null }
    if (-not $config.ProfileKey -or -not $settings.TrustedSids) { throw 'Report mail queue configuration is incomplete.' }
    $queuePath = if ($settings -is [Collections.IDictionary]) { [string]$settings['QueuePath'] }
        else { $property = $settings.PSObject.Properties['QueuePath']; if ($property) { [string]$property.Value } else { '' } }
    if ([string]::IsNullOrWhiteSpace($queuePath) -or $queuePath.Trim() -in @('__USE_GLOBAL__','USE_GLOBAL')) {
        $queuePath = '{{WorkspaceRootPath}}\Tenants\{{TenantKey}}\MAIL'
    }
    $queuePath = [string](Resolve-SmartM365ConfigValue $queuePath)
    if ([string]::IsNullOrWhiteSpace($queuePath) -or $queuePath -match '\{\{') { throw 'Report mail QueuePath could not be resolved.' }
    return [pscustomobject]@{ Policy=$settings; Profile=[string]$config.ProfileKey; Path=$queuePath }
}

function Get-SmartM365ReportStorageAcl {
    param($Item)
    if ($PSVersionTable.PSEdition -eq 'Desktop') { return $Item.GetAccessControl() }
    return [IO.FileSystemAclExtensions]::GetAccessControl($Item)
}

function Test-SmartM365DeferredMailDirectory {
    param([string]$Path,[string[]]$TrustedSids)
    # Queue provisioning is an explicit operator action. Never change share ACLs here.
    $allowed = @('S-1-5-18','S-1-5-32-544') + $TrustedSids
    $directory = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if (-not $directory.PSIsContainer) { return $false }
    $cursor = $directory
    while ($cursor) {
        if ($cursor.Attributes -band [IO.FileAttributes]::ReparsePoint) { return $false }
        # Parent access belongs to the existing share administration boundary.
        # Only the queue and its payload files must restrict access to TrustedSids.
        $cursor = $cursor.Parent
    }
    $acl = Get-SmartM365ReportStorageAcl -Item $directory
    if (-not $acl.AreAccessRulesProtected -or $acl.GetOwner([Security.Principal.SecurityIdentifier]).Value -notin $allowed) { return $false }
    foreach ($rule in $acl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier])) {
        if ($rule.AccessControlType -eq 'Allow' -and $rule.IdentityReference.Value -notin $allowed) { return $false }
    }
    foreach ($item in Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop) {
        if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { return $false }
        $childAcl = Get-SmartM365ReportStorageAcl -Item $item
        if ($childAcl.GetOwner([Security.Principal.SecurityIdentifier]).Value -notin $allowed) { return $false }
        foreach ($rule in $childAcl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier])) {
            if ($rule.AccessControlType -eq 'Allow' -and $rule.IdentityReference.Value -notin $allowed) { return $false }
        }
    }
    return $true
}

function Initialize-SmartM365DeferredMailRuntime {
    $folder = $PSScriptRoot
    Import-Module (Join-Path $folder 'SmartM365.Orchestrator.ReportMailSchedule.psm1') -Global -ErrorAction Stop
    Import-Module (Join-Path $folder 'SmartM365.Orchestrator.ReportMailQueue.psm1') -Global -ErrorAction Stop
}

function Add-SmartM365DeferredReport {
    param([hashtable]$Message,[string[]]$Attachments,[string]$MailPurpose,[string]$DeliveryClass)
    $global:SmartM365LastMailDelivery = 'Sent'
    if ($MailPurpose -in @('Error','Maintenance') -or $DeliveryClass -eq 'Operational') { return $false }
    # Limit interception to a SmartInventory caller; shared Core has other consumers.
    if (-not @(Get-PSCallStack | Where-Object { $_.ScriptName -match '[\\/]SmartInventory[\\/]' }).Count) { return $false }
    $cfg = Get-SmartM365DeferredMailSettings
    if (-not $cfg) { return $false }
    Initialize-SmartM365DeferredMailRuntime
    $now = Get-SmartM365MailUtcNow
    $supersedeKey = ''
    $window = $null
    try { $window = Get-SmartM365MailDeliveryWindow }
    catch { WriteLog -Message ('Mail delivery window is invalid; report is sent to administrators only: {0}' -f $_.Exception.Message) -Level WARNING; return $false }
    if ($window) {
        # Business hours: administrators receive the report now (routing restricts the
        # immediate send to them); the other recipients get a queued copy at the next opening.
        try { $state = Get-SmartM365MailDeliveryWindowState -Window $window -Now $now }
        catch { WriteLog -Message ('Mail delivery window is invalid; report is sent to administrators only: {0}' -f $_.Exception.Message) -Level WARNING; return $false }
        if ($state.IsOpen) { return $false }
        $due = $state.NextOpenUtc
        $Message.ExcludeAdministrators = $true
        $supersedeKey = Get-SmartM365DeferredReportSupersedeKey -Message $Message
    } else {
        $due = Get-SmartM365ReportMailNotBeforeUtc -Now $now -NotBefore $cfg.Policy.NotBefore -TimeZoneId $cfg.Policy.TimeZone
        if ($due -le $now) { return $false }
    }
    if ($Message['ReportCadenceDate'] -and -not (Test-SmartM365ReportMailDeliveryDate -Date $Message['ReportCadenceDate'] -Now $due)) {
        WriteLog -Message 'Report mail copy cannot be queued: the next delivery window is on a later calendar day (Europe/Paris).' -Level INFO
        # In business-hours mode the existing policy still sends the administrators' copy now.
        if ($window) { return $false }
        $global:SmartM365LastMailDelivery = 'Skipped'
        return $true
    }
    $files = @()
    foreach ($path in $Attachments) {
        $bytes = [IO.File]::ReadAllBytes($path)
        $files += @{ Name=[IO.Path]::GetFileName($path); Content=[Convert]::ToBase64String($bytes) }
    }
    $Message.Files = $files
    $Message.MailPurpose = 'Report'
    $trusted = @($cfg.Policy.TrustedSids)
    $verify = { param($path) Test-SmartM365DeferredMailDirectory -Path $path -TrustedSids $trusted }
    $queueParameters = @{ QueuePath=$cfg.Path; AssertTrustedDirectory=$verify; TenantKey=$cfg.Profile; DeduplicationKey=([guid]::NewGuid().ToString('N')); NotBeforeUtc=$due; Message=$Message }
    if ($supersedeKey) { $queueParameters.SupersedeKey = $supersedeKey }
    $item = Add-SmartM365ReportMailQueueItem @queueParameters
    $global:SmartM365LastMailDelivery = 'Queued'
    if ($window) {
        WriteLog -Message ('Outside the mail delivery window: administrators receive the report now; copy for the other recipients queued: {0}; eligible {1}' -f $item.Id,$due.ToString('o')) -Level INFO
        return $false
    }
    WriteLog -Message ('Report mail queued: {0}; eligible {1}' -f $item.Id,$due.ToString('o')) -Level INFO
    return $true
}

function Get-SmartM365DeferredReportSupersedeKey {
    # Only the latest queued copy per collector script, recipients and subject (digits
    # ignored, so dates do not matter) is delivered when the window opens.
    param([Parameter(Mandatory)][hashtable]$Message)
    $frames = @(Get-PSCallStack | Where-Object { $_.ScriptName -match '[\\/]SmartInventory[\\/]' })
    $script = if ($frames.Count) { [IO.Path]::GetFileName($frames[-1].ScriptName).ToLowerInvariant() } else { '' }
    $recipients = @(([string]$Message.To + ';' + [string]$Message.Cc) -split '[;,]' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique) -join ';'
    $subject = ([string]$Message.Subject -replace '\d', '').Trim().ToLowerInvariant()
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($script + '|' + $recipients + '|' + $subject)))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Invoke-SmartM365DeferredMail {
    $cfg = Get-SmartM365DeferredMailSettings
    if (-not $cfg) { return }
    Initialize-SmartM365DeferredMailRuntime
    $now = Get-SmartM365MailUtcNow
    $window = Get-SmartM365MailDeliveryWindow
    if ($window) {
        if (-not (Get-SmartM365MailDeliveryWindowState -Window $window -Now $now).IsOpen) { return }
    } elseif ((Get-SmartM365ReportMailNotBeforeUtc -Now $now -NotBefore $cfg.Policy.NotBefore -TimeZoneId $cfg.Policy.TimeZone) -gt $now) { return }
    $trusted = @($cfg.Policy.TrustedSids)
    $verify = { param($path) Test-SmartM365DeferredMailDirectory -Path $path -TrustedSids $trusted }
    $isExpired = {
        param($message)
        $date = $message.PSObject.Properties['ReportCadenceDate']
        return ($date -and -not (Test-SmartM365ReportMailDeliveryDate -Date $date.Value -Now (Get-SmartM365MailUtcNow)))
    }
    $send = {
        param($message)
        # A protected temporary subdirectory cannot be placed inside the flat queue.
        $temp = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DeferredMail-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $temp
        try {
            $acl = New-Object Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true,$false)
            foreach ($sid in @('S-1-5-18','S-1-5-32-544',[Security.Principal.WindowsIdentity]::GetCurrent().User.Value)) {
                $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier($sid)),'FullControl','ContainerInherit,ObjectInherit','None','Allow')))
            }
            Set-Acl -LiteralPath $temp -AclObject $acl
            $paths = @()
            foreach ($file in $message.Files) {
                if ([IO.Path]::GetFileName($file.Name) -cne $file.Name -or $file.Name -in @('.','..') -or $file.Name -match '[:\\/]') { throw 'Invalid queued attachment name.' }
                $target = Join-Path $temp $file.Name
                [IO.File]::WriteAllBytes($target,[Convert]::FromBase64String($file.Content)); $paths += $target
            }
            $args = @{ From=$message.From; To=$message.To; Cc=$message.Cc; Subject=$message.Subject; BodyHtml=$message.BodyHtml; Attachments=$paths; AllowAttachments=$true; MailPurpose='Report'; DeliveryClass='Operational'; SendMailMode=$message.SendMailMode; SmtpServer=$message.SmtpServer; SmtpPort=[int]$message.SmtpPort }
            $excludeProperty = $message.PSObject.Properties['ExcludeAdministrators']
            if ($excludeProperty -and $excludeProperty.Value -eq $true) {
                # Administrators already received this report when it was produced.
                $global:SmartM365MailExcludeAdministrators = $true
                $route = Resolve-SmartM365UnifiedMailRoute -To ([string]$message.To) -Cc ([string]$message.Cc) -MailPurpose Report
                if ([string]::IsNullOrWhiteSpace($route.To)) {
                    WriteLog -Message ('Queued report has no recipient besides administrators; nothing more to send: {0}' -f $message.Subject) -Level INFO
                    return $true
                }
            }
            SendEmailHtmlReport @args
            return $true
        } finally {
            Remove-Variable -Name SmartM365MailExcludeAdministrators -Scope Global -ErrorAction SilentlyContinue
            $resolved = [IO.Path]::GetFullPath($temp)
            if ((Split-Path -Leaf $resolved) -notmatch '^SmartM365-DeferredMail-[a-f0-9]{32}$' -or (Split-Path -Parent $resolved).TrimEnd('\\') -ne ([IO.Path]::GetTempPath()).TrimEnd('\\')) { throw 'Unexpected mail staging path.' }
            Remove-Item -LiteralPath $resolved -Recurse -Force
        }
    }
    Invoke-SmartM365ReportMailQueueBatch -QueuePath $cfg.Path -AssertTrustedDirectory $verify -TenantKey $cfg.Profile -Now $now -SendAction $send -IsExpired $isExpired -MaximumItems 10
}


# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD4nHpzZKxjT+QT
# mBCCO+dXw2EP6SnppCUz7MauvGyAC6CCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgw3+gyWPevugs
# OV2pPaJMfRSKrnVc0NIt2S/dMLMKMvUwDQYJKoZIhvcNAQEBBQAEggGAUaoaAnZZ
# dQC2v8N/gg8jg3HigxOiIo4MHJ2rr6Nio+q14WrqIRuzqRB9TyfTIKzIuPypwbfy
# PXokl7b8ITAq0aH0DsBNbvTPeXQEIxX5Wde3P6WFfEEYnh9cBFRqYa4mGQW/fEXG
# gwvcxfDjbZ2355Y3dlrNS2KlvItgUOP8A5Gz/fOJnB5UqPHMIZK4iU7vd802WaKj
# WtiYakjVCSN3u+APB4DlOezLs/H1dClCMcfjrPIxaZEiSd//7SGhgRRVfar8kYMG
# I7RcUA1YkpOF4v7FbXdRlqIDLdwqu2sER+uF8GEQHltVdr922EvXPV8vu7hGHruT
# 8p/p1MrjgFT7ucZGEJl78kFkLAbjAnrRj4sKOloWvlzTQho06TkP/Sbbk6pptcU8
# 2tBvBI5h8o5POhRnnKKoEAeaamvfiiFx+JqqONF/pAyrYsTkvIYJGRgfQ4m/LmYs
# hZ9RNIUfia7taQJSK6RNAAmn2CeJYL5eDvLJEEgj1pu4NB5yQRmeSDbOoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDgzNDU2WjA/BgkqhkiG9w0BCQQxMgQwsdcjO7hX/Ek4oW9edgOXEJtd
# z5HJOTq1Ou3Hs67H/pT2fH7sI3cs9KQuB3527LojMA0GCSqGSIb3DQEBAQUABIIC
# AIDS/kIcY+jAICSus2q+oRtruIavUhYsnjpsahbRv67ZSQOwp8ZW5x4mgEwBreh/
# YZ4N1wrlb/GXSqyHD+NqdP3FnAGtxiMW+GTJ44K5BA88HW1K/UxBJbi8Qgf8/vxp
# KX2G16On0fTehKQ0VnfD4lu6BMZWTjc8d8u3Yj42d/AmSHHfglrsiV5mwfnGO2CJ
# eKWg1w0Yiv6HF7qAQmjFWq54og8mTfGAR6wqDCXiHGhcGsQta3I+r7f/o/dk4QZo
# NOathjw+3khMLNRXhrIzliLyp0KK+z88OkJiSbFj2cBTkS8v+VMwjXm1LzRzMtv+
# sb3lynwCOTxHA9sC+RKNr4Gnt2c00VeUOnInOhaXG+uMNUyLVz/pQ2OhYwgc87fu
# WGusxpXvHP5xmFb71y1bCHS7tPppgFZEqQkALW6WnOg1FU7eoQjalRnB94NlQFyp
# axGWZwKPghONyLaazrwkD3pmUUkahYyn4CSdYv4IZff/bHKMUZghvY86/AQ1j1nj
# x9zb+lLIqm1o9RX9YPqiZZntVDd0DlLB1T5BCh/ouim//4nt5DuJQlYLkB37Of9a
# V8ZTRORUodXA7LbTo/wzCf79J9YvzhZkMqZtZlemyhKqQY2yu4WaYHkjGxhxPsGr
# NSBt6iF+yOBnstkGwrDITyZg+T+310U7GD0AOVzYbSzF
# SIG # End signature block
