<# Shared deferred report delivery. Disabled unless explicitly configured. #>
function Get-SmartM365MailUtcNow { return [datetimeoffset]::UtcNow }
function Get-SmartM365DeferredMailSettings {
    $config = Get-SmartM365EffectiveModuleGlobalConfig
    $p = $config.PSObject.Properties['ReportMailSchedule']
    if (-not $p -or -not $p.Value) { return $null }
    $settings = $p.Value
    if ($settings.Enabled -isnot [bool]) { throw 'ReportMailSchedule.Enabled must be Boolean.' }
    if (-not $settings.Enabled) { return $null }
    if (-not $config.ProfileKey -or -not $settings.QueuePath -or -not $settings.TrustedSids) { throw 'Report mail queue configuration is incomplete.' }
    return [pscustomobject]@{ Policy=$settings; Profile=[string]$config.ProfileKey; Path=[string](Resolve-SmartM365ConfigValue $settings.QueuePath) }
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
        $parentAcl = Get-Acl -LiteralPath $cursor.FullName -ErrorAction Stop
        if ($parentAcl.GetOwner([Security.Principal.SecurityIdentifier]).Value -notin $allowed) { return $false }
        foreach ($rule in $parentAcl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier])) {
            if ($rule.PropagationFlags -band [Security.AccessControl.PropagationFlags]::InheritOnly) { continue }
            $writeMask = [long]0x500D0156
            if ($rule.AccessControlType -eq 'Allow' -and ([long]$rule.FileSystemRights -band $writeMask) -and $rule.IdentityReference.Value -notin $allowed) { return $false }
        }
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
    $due = Get-SmartM365ReportMailNotBeforeUtc -Now $now -NotBefore $cfg.Policy.NotBefore -TimeZoneId $cfg.Policy.TimeZone
    if ($due -le $now) { return $false }
    $files = @()
    foreach ($path in $Attachments) {
        $bytes = [IO.File]::ReadAllBytes($path)
        $files += @{ Name=[IO.Path]::GetFileName($path); Content=[Convert]::ToBase64String($bytes) }
    }
    $Message.Files = $files
    $Message.MailPurpose = 'Report'
    $trusted = @($cfg.Policy.TrustedSids)
    $verify = { param($path) Test-SmartM365DeferredMailDirectory -Path $path -TrustedSids $trusted }
    $item = Add-SmartM365ReportMailQueueItem -QueuePath $cfg.Path -AssertTrustedDirectory $verify -TenantKey $cfg.Profile -DeduplicationKey ([guid]::NewGuid().ToString('N')) -NotBeforeUtc $due -Message $Message
    $global:SmartM365LastMailDelivery = 'Queued'
    WriteLog -Message ('Report mail queued: {0}; eligible {1}' -f $item.Id,$due.ToString('o')) -Level INFO
    return $true
}

function Invoke-SmartM365DeferredMail {
    $cfg = Get-SmartM365DeferredMailSettings
    if (-not $cfg) { return }
    Initialize-SmartM365DeferredMailRuntime
    $now = Get-SmartM365MailUtcNow
    if ((Get-SmartM365ReportMailNotBeforeUtc -Now $now -NotBefore $cfg.Policy.NotBefore -TimeZoneId $cfg.Policy.TimeZone) -gt $now) { return }
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
            SendEmailHtmlReport @args
            return $true
        } finally {
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDvMuOuCFghR4QA
# qomMuJNYh8DC4K7MbThughJX9i67+aCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCACjqCezZxkOY3V5b+Ix/Bq
# ayh+xmISz0twD3oJcDjGQzANBgkqhkiG9w0BAQEFAASCAYAQq9LwFA/xVAErbzBW
# khsgxhjayqqp59hVqmjP3ca3IGyL2HuC7dUNh4GZ2GH4M47FnXh/vrxLvHCf2PFB
# sVw+VSATtLSYvLYjClzLLw0xSDe9RcalfCEeTIRc+4X97d0a5/nkWX7g+OtMrPlA
# /WTAgw+FTuV0QlP7e9wMPP3k2OorTukzpu/bI12t8AHPILS5Y6tzRm/7qEQ8g70i
# Io0fpDle+XdD3GaIj/nzUtXU+p62QYWCmhYQXorrL2EtrYCzmzLVSWJOTShR8pe3
# HXXqcpAt1dEp5PXJjkrcVMul6ol/VhwBK/Co1n/TJ/U7LEKOqwm9TpqndAv/2Ycn
# /pm0hN5YUu/qhh0xdAFHfoffMINAi0dDKImsoFpLzvlnZI0tYrjH7X9HN2v6xTab
# fnP3MFX1ONBGt668X+GwZ5n+oQLVmnMMI0I7V79AJzTEIkVA1GBBpzzT2DQah722
# +zvHQ6vTp6CcosdjGW1eYECoqs8QWWBdOOPqwVwMeOLO46Q=
# SIG # End signature block
