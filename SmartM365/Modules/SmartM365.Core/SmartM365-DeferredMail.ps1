<# Shared deferred report delivery. Disabled unless explicitly configured. #>
# Version: 1.1.0
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
    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
    if (-not $acl.AreAccessRulesProtected -or $acl.GetOwner([Security.Principal.SecurityIdentifier]).Value -notin $allowed) { return $false }
    foreach ($rule in $acl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier])) {
        if ($rule.AccessControlType -eq 'Allow' -and $rule.IdentityReference.Value -notin $allowed) { return $false }
    }
    foreach ($item in Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop) {
        if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { return $false }
        $childAcl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
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
    Invoke-SmartM365ReportMailQueueBatch -QueuePath $cfg.Path -AssertTrustedDirectory $verify -TenantKey $cfg.Profile -Now $now -SendAction $send -MaximumItems 10
}


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDsMxTjE5Q/BJRE
# XafkAfTXr6YMJ+LvRepuiuSdhLYrC6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAarUluemkAygMj26lV4rGy
# u/Qb5ha1bQCwuklUlEJlWjANBgkqhkiG9w0BAQEFAASCAYARps7aAWUpsWmx5tZ4
# KdBQJBG1pXR8HykzrLYqvfOtEw6QKTa+oiv2RQQiTAs7po3V7+sy6oh//DDOp6Er
# KjWIK2drJDbx49O1CVGDsJ51+6X2sgwl0rDbs0bJg6viXyN+tO3FkBIgj8YlDxrq
# 2ICnbuL98g1YsZMR+DsuWosxc8mvIuk7W/Js/K3/dMdHS4NnTz+wNDQ5aKWbR2QR
# FBQTGf5Yxt9BKuxTvvQCGxdEE4zRdYLWZJEKFe8GGVxCltpDcK+BO/ksEF5ckEhc
# gmLHPkrESijYhGkCAeuH8GNd+BlI5OYZfEv28NTBP+ayih9jTCS42LspSY2Fde3s
# QxKKqThbDZq21EyU+bpTH0WSlyPR/VwDr7o7SKJuuyKLqM+cvjBzZBWdXjKGQpXK
# n6TKW3Nogj1Oskxs4pNXv/jv7dQdI22Ae0Dh72CKAwhKZ0cWWiE3K4faXXCQN1/H
# BeoaQfQ1tsaKcTvrKafJpb0ybFnsVSv0JG7AaoT66FizZIk=
# SIG # End signature block
