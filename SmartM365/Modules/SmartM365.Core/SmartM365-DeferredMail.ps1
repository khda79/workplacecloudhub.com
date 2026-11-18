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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAGCsvC2Bd+17NI
# GuZg32VGwqM9CEvGtCAnDjR/1jZRLKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBl0TJWiOnFNq7+GiEsBz3EFK6xnADF4N11AsQnqy5wYMA0GCSqG
# SIb3DQEBAQUABIIBgKMxT702/fevCagYm95aLn899C1NlJfk+5Fpg5PJCgoBgfWE
# 6hPmOqPY9VFHyQMsKWx8gWZ4i16tt+Hnie/gjCQ4jgPXiwPyBA7jlrqgUP8KnkOm
# AzU5z7nqU0PLPMy+Nc2pRUWpSsC+I5RBCnP4NmNWJBirskVzPfBGlY7VxDO/n0we
# TjqZsAkW9W7YKu6dBu7bYE2ALa2fFsG4CBL6kgIpTETF2TOiuWixY30IimeE8BVf
# Mm2dEmWOsZVCHXyR0MvmTLtlxKU7kQ7xv1jTnUPb3vgHg89v+kGta9QF5aGrw5tw
# vWtf+oefUe7G8/oATX4EXLWS+FEpS7eHurGw8BMQO9u9U5ivr2kjrJrd4pJwqDG+
# NromnKI3zLRkYsjGoIfmCqtBsaBwkZbJzxVKtry3vx1PDpxWUyzU5w3rSQtDIXCN
# YVvWiVdgDrmGcGwzTwvKrQfpZ5XEV6gUTkHzGK3JaJ/C50fphoo+NmL++uIH4VVz
# 1KYAY7Lph0Yk59l4mKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTgxMzAx
# NDZaMC8GCSqGSIb3DQEJBDEiBCBq1SsjF2o/E31xWLEpSKVGjtCxRpdaRiOaOjB0
# xCSOEDANBgkqhkiG9w0BAQEFAASCAgC2V4VPdul3D1aYk6BjLl/mXk5qP9ROOFi/
# ELFUFxW/v6oJA/w4UZ3HXIfiY6/5lbxKAzZzqa3+rrZ09YbJll0L7uzteJTgSWq8
# KFYACukBSIWM+CmIu+W1wiKJSa+P3LwHO0PiDZhb4Xz29rBsZzVsJMrbweg1Sbc/
# 1L8xSnqb+E9PQAy7Z9Z+LnC+EPH1C479KQ2J6tXy7t/6gKBt8IhtxtNJYKLseSDG
# deXYo4LkLr+WX2bxugcYPeykHy9L6lEXPIOzww2uu/cL510dArOob3RNXCUbvZML
# eq92YWo5BBMJhXcu01S5ul1VnRt0KBR4RlOitFmIUYR7znAP7MnE0eWIEiKUD4DU
# 57WrqR2UOXQxs5BnvCLQALWxLMai29SP/hjf+SCVhvOS7A9j08O316X1gvbcSeNC
# /ElRQlctP+qerllju59tPcHVlvBR07riRKbT3vOr0UencWnes7/O4rYhYJEGUexm
# NayQp+8X7gHP4amNg9smoBtFzCe2MDiEXU3QdQ8PFW/XmkY1mPPbSJs+kBRJ8EOx
# JMUtWN6viod27eLO93nv1dbAl4OmsIihfskFmdFv+rB3xh9UtHbYkzX0lWbT2+pH
# WppripUrYLKbkJ2S0mPLA83W+9ezd009H72oz+N7H3Rp6UwDJ0EktFJM3wpN/gXc
# LRyaOkOmig==
# SIG # End signature block
