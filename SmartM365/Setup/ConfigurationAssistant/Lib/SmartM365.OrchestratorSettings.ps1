# Explicit-file local orchestrator settings editor, version 0.4.0.
function Get-SiOrchestratorSettingNames {
    @('EnableSharePointUploadDataAll','EnableSharePointUploadDataLast','EnableSharePointUploadDataPowerBI','EnableSharePointUploadLogAll','WorkspaceRootPath','DataAllRootPath','LatestCsvFolderPath','LogAllRootPath','ErrorMailTo','To','Cc','Bcc','From','SmtpServer','SendMailMode','JobMailMode','SmtpPort','UseIntegratedAuth','UseSsl','RelayIp','SendDailySummaryEmail','DailySummaryTime','AllowedServers','PeerMonitoringEnabled','PeerJobMonitoringEnabled','ExpectedOrchestratorServers','PeerMonitoringCheckIntervalSeconds','PeerHeartbeatStaleMinutes','PeerMonitoringConfirmationChecks','PeerJobStartGraceMinutes','PeerRecycleGraceMinutes','PeerAlertReminderMinutes','PeerAlertMailRetryMinutes','PeerRecoveryEmailEnabled','MaxConcurrency','MaxConcurrencyByServer','MaxLifetimeHours','AtomicWriteRetrySeconds','CentralConfigurationEnabled','AutoRecycleOnRuntimeUpdate','RuntimeUpdateCheckIntervalSeconds','RuntimeUpdateStableChecks','RuntimeUpdateCooldownMinutes','MonitorCoreModuleVersion','SendRuntimeUpdateEmail','DistributedSchedulingEnabled','CapabilityProbeMode','CapabilityProbeTimeoutSeconds','CapabilityRefreshMinutes','CapabilityMaxAgeMinutes','LaunchPrerequisiteChecksEnabled','PrerequisiteCacheMinutes','PrerequisitePathTimeoutSeconds','PrerequisiteAlertRepeatHours','CertificateExpiryWarningDays','ElectionPlanRefreshSeconds','ElectionClaimGraceMinutes','ElectionHistoryDays','ElectionClaimRetentionDays','ElectionClaimRetentionMaxSecondsPerRun','ElectionWeight','ElectionWeightsByServer','ServerJobPolicies','ExchangeOnlineOrganization','TickSeconds','DependencyWaitLogIntervalMinutes','OrchestratorHeartbeatLogIntervalMinutes','OrchestratorSharePointUploadIntervalMinutes','AuthenticodeValidationEnabled','AuthenticodeValidationMode','AuthenticodeAllowedThumbprints','AuthenticodeCheckCoreModule','AuthenticodeCheckWindowsPowerShellModule','OrchestratorDataFolderPath','OrchestratorLogFolderPath','OrchestratorLogRetentionDays','JobLogRetentionDays','JobRunsCsvRetentionDays','OrchestratorTranscriptEnabled','OrchestratorTranscriptRetentionDays','PipelineRunArchiveDays','PipelineRunArchiveRetentionDays','MailHtmlCopyRetentionHours','AuthenticodeInstallTrustedCertificates','AuthenticodeTrustedCertificatePaths','AuthenticodeInstallTrustedRoot','AuthenticodeInstallTrustedPublisher','DependencyWaitTimeoutMinutes','OrchestratorRunsCsvLockTimeoutSeconds','FeatureUpdatePolicyId')
}
function Get-SiOrchestratorSettingKind {
    param([string]$Field)
    if ($Field -like 'EnableSharePointUpload*') { return 'InheritBoolean' }
    if ($Field -in @('AllowedServers','ExpectedOrchestratorServers','AuthenticodeAllowedThumbprints','AuthenticodeTrustedCertificatePaths')) { return 'Array' }
    if ($Field -in @('MaxConcurrencyByServer','ElectionWeightsByServer','ServerJobPolicies')) { return 'Object' }
    if ($Field -eq 'ElectionWeight') { return 'Decimal' }
    if ($Field -in @('WorkspaceRootPath','DataAllRootPath','LatestCsvFolderPath','LogAllRootPath','ErrorMailTo','To','Cc','Bcc','From','SmtpServer','SendMailMode','JobMailMode','RelayIp','DailySummaryTime','CapabilityProbeMode','ExchangeOnlineOrganization','AuthenticodeValidationMode','OrchestratorDataFolderPath','OrchestratorLogFolderPath','FeatureUpdatePolicyId')) { return 'String' }
    if ($Field -match 'Enabled$' -or $Field -in @('UseIntegratedAuth','UseSsl','SendDailySummaryEmail','AutoRecycleOnRuntimeUpdate','MonitorCoreModuleVersion','SendRuntimeUpdateEmail','AuthenticodeCheckCoreModule','AuthenticodeCheckWindowsPowerShellModule','AuthenticodeInstallTrustedCertificates','AuthenticodeInstallTrustedRoot','AuthenticodeInstallTrustedPublisher')) { return 'Boolean' }
    return 'Integer'
}
function Get-SiOrchestratorSettingSnapshot {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string]$Field,[string]$TemplatePath='')
    if ($Field -notin @(Get-SiOrchestratorSettingNames)) { throw 'Setting is not supported by this editor.' }
    if (-not [IO.Path]::IsPathFullyQualified($Path) -or [IO.Path]::GetFileName($Path) -inotmatch '^SmartM365-Inventory-Orchestrator\.local\.json(?:\.txt)?$') { throw 'Select an existing orchestrator local JSON configuration file.' }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'The selected file does not exist.' }
    $read=Read-SmartM365JsonDocument -Path $Path
    if ($read.Coexistence -ne 'Single' -or [IO.Path]::GetFullPath($read.Path) -ine [IO.Path]::GetFullPath($Path)) { throw 'Conflicting configuration names require review.' }
    if ($read.Document -isnot [pscustomobject]) { throw 'Configuration must be an object.' }
    $property=$read.Document.PSObject.Properties[$Field]; $present=($null -ne $property); $templateHash=''
    if (-not $present) {
        if (-not [IO.Path]::IsPathFullyQualified($TemplatePath)) { throw 'The setting is absent and the trusted template is unavailable.' }
        $bytes=[IO.File]::ReadAllBytes($TemplatePath); $sha=[Security.Cryptography.SHA256]::Create()
        try { $templateHash=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','') } finally { $sha.Dispose() }
        $template=[Text.UTF8Encoding]::new($false,$true).GetString($bytes).TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop
        if ($template -isnot [pscustomobject]) { throw 'Orchestrator template must be an object.' }
        $property=$template.PSObject.Properties[$Field]
        if ($null -eq $property) { throw 'The setting is not present in the selected template.' }
    }
    $kind=Get-SiOrchestratorSettingKind $Field; $current=$property.Value
    if (($kind -eq 'String' -and $current -isnot [string]) -or ($kind -eq 'Boolean' -and $current -isnot [bool]) -or ($kind -eq 'Integer' -and $current -isnot [int] -and $current -isnot [long]) -or ($kind -eq 'Decimal' -and ($current -isnot [ValueType] -or $current -is [bool] -or $current -is [datetime])) -or ($kind -eq 'InheritBoolean' -and $current -isnot [bool] -and $current -notin @('__USE_GLOBAL__','USE_GLOBAL')) -or ($kind -eq 'Array' -and $current -isnot [array] -and $current -isnot [string]) -or ($kind -eq 'Object' -and $current -isnot [pscustomobject])) { throw 'The selected setting has an unexpected type.' }
    $value=if ($kind -in @('Array','Object')) { ConvertTo-Json -InputObject $current -Depth 20 -Compress } elseif ($kind -eq 'Decimal') { [Convert]::ToString($current,[Globalization.CultureInfo]::InvariantCulture) } else { [string]$current }
    [pscustomobject]@{Path=$read.Path;Hash=$read.SHA256;Field=$Field;Kind=$kind;Value=$value;Present=$present;TemplatePath=$TemplatePath;TemplateHash=$templateHash}
}
function ConvertTo-SiOrchestratorSettingValue {
    param([Parameter(Mandatory)]$Snapshot,[Parameter(Mandatory)][AllowEmptyString()][string]$Value)
    if ($Value.Length -gt 65536 -or $Value -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') { throw 'The new value is too large or contains unsupported direction controls.' }
    $fresh=$Snapshot
    switch ($fresh.Kind) {
        'Boolean' { if ($Value -notin @('true','false')) { throw 'Enter true or false.' }; $newValue=($Value -ieq 'true') }
        'InheritBoolean' { if ($Value -in @('__USE_GLOBAL__','USE_GLOBAL')) { $newValue='__USE_GLOBAL__' } elseif ($Value -in @('true','false')) { $newValue=($Value -ieq 'true') } else { throw 'Enter true, false or __USE_GLOBAL__.' } }
        'Integer' {
            $number=0
            if ($Value -notmatch '^\d{1,7}$' -or -not [int]::TryParse($Value,[ref]$number) -or $number -gt 1000000) { throw 'Enter a whole number between 0 and 1000000.' }
            if ($fresh.Field -in @('MaxConcurrency','TickSeconds','RuntimeUpdateCheckIntervalSeconds','RuntimeUpdateStableChecks') -and $number -lt 1) { throw 'This setting must be at least 1.' }
            if ($fresh.Field -eq 'MaxConcurrency' -and $number -gt 128) { throw 'MaxConcurrency must not exceed 128 in this editor.' }
            if ($fresh.Field -eq 'SmtpPort' -and ($number -lt 1 -or $number -gt 65535)) { throw 'SmtpPort must be between 1 and 65535.' }
            $newValue=$number
        }
        'Decimal' {
            $weight=0.0
            if (-not [double]::TryParse($Value,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$weight) -or [double]::IsNaN($weight) -or [double]::IsInfinity($weight) -or $weight -lt 0.1 -or $weight -gt 1000) { throw 'ElectionWeight must be between 0.1 and 1000.' }
            $newValue=$weight
        }
        'Array' {
            try { $newValue=ConvertFrom-Json -InputObject $Value -NoEnumerate -ErrorAction Stop } catch { throw 'Enter a JSON array of text values.' }
            if ($newValue -isnot [array] -or @($newValue | Where-Object { $_ -isnot [string] -or [string]::IsNullOrWhiteSpace($_) -or $_ -match '[\x00-\x1f]' }).Count -gt 0) { throw 'Enter a JSON array of nonempty text values.' }
            if ($fresh.Field -eq 'AuthenticodeAllowedThumbprints' -and @($newValue | Where-Object { $_ -notmatch '^[0-9a-fA-F]{40}$' }).Count -gt 0) { throw 'Each certificate thumbprint must contain 40 hexadecimal characters.' }
        }
        'Object' {
            try { $newValue=ConvertFrom-Json -InputObject $Value -ErrorAction Stop } catch { throw 'Enter a JSON object.' }
            if ($newValue -isnot [pscustomobject]) { throw 'Enter a JSON object.' }
            foreach ($property in $newValue.PSObject.Properties) {
                if ([string]::IsNullOrWhiteSpace($property.Name) -or $property.Name -match '[\x00-\x1f\u200e\u200f\u202a-\u202e\u2066-\u2069]') { throw 'Server names cannot be empty or contain control characters.' }
                if ($fresh.Field -eq 'MaxConcurrencyByServer' -and ($property.Value -isnot [int] -and $property.Value -isnot [long] -or $property.Value -lt 1 -or $property.Value -gt 128)) { throw 'Each per-server concurrency must be an integer from 1 to 128.' }
                if ($fresh.Field -eq 'ElectionWeightsByServer' -and ($property.Value -isnot [ValueType] -or $property.Value -is [bool] -or [double]$property.Value -lt 0.1 -or [double]$property.Value -gt 1000)) { throw 'Each server weight must be a number from 0.1 to 1000.' }
                if ($fresh.Field -eq 'ServerJobPolicies') {
                    $policy=$property.Value; $allowed=@('SharedRuntime','Graph','EXO','AD','ExchangeOnPrem','SharePointOnPrem','TeamsPowerShell','Adaxes','Python')
                    if ($policy -isnot [pscustomobject] -or $policy.OnlyJobsRequiring -isnot [array] -or $policy.OnlyJobsRequiring.Count -eq 0 -or @($policy.OnlyJobsRequiring | Where-Object { $_ -isnot [string] -or $_ -notin $allowed }).Count -gt 0) { throw 'Each server policy requires a nonempty OnlyJobsRequiring array of supported capabilities.' }
                }
            }
        }
        'String' {
            if ($fresh.Field -eq 'JobMailMode' -and $Value -cnotin @('Always','OnError','Never')) { throw 'JobMailMode must be Always, OnError or Never.' }
            if ($fresh.Field -eq 'DailySummaryTime' -and $Value -notmatch '^([01]\d|2[0-3]):[0-5]\d$') { throw 'DailySummaryTime must be HH:mm.' }
            if ($fresh.Field -eq 'SendMailMode' -and $Value -notin @('Graph','SMTP','Both','__USE_GLOBAL__','USE_GLOBAL')) { throw 'SendMailMode must be Graph, SMTP, Both or __USE_GLOBAL__.' }
            if ($fresh.Field -eq 'CapabilityProbeMode' -and $Value -notin @('Static','ReadOnly')) { throw 'CapabilityProbeMode must be Static or ReadOnly.' }
            if ($fresh.Field -eq 'AuthenticodeValidationMode' -and $Value -notin @('Audit','Enforce')) { throw 'AuthenticodeValidationMode must be Audit or Enforce.' }
            if ($Value.Length -gt 4096 -or $Value -match '[\x00-\x1f\u200e\u200f\u202a-\u202e\u2066-\u2069]') { throw 'Text contains unsupported characters or is too long.' }
            $newValue=$Value
        }
    }
    return ,$newValue
}
function Set-SiOrchestratorSettingSnapshot {
    param([Parameter(Mandatory)]$Snapshot,[Parameter(Mandatory)][AllowEmptyString()][string]$Value)
    if ($Value.Length -gt 65536 -or $Value -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') { throw 'The new value is too large or contains unsupported direction controls.' }
    $fresh=Get-SiOrchestratorSettingSnapshot -Path $Snapshot.Path -Field $Snapshot.Field -TemplatePath $Snapshot.TemplatePath
    if ($fresh.Hash -cne $Snapshot.Hash -or $fresh.Kind -cne $Snapshot.Kind -or $fresh.Value -cne $Snapshot.Value -or $fresh.Present -ne $Snapshot.Present -or $fresh.TemplateHash -cne $Snapshot.TemplateHash) { throw 'Configuration or template changed; reload before saving.' }
    $newValue=ConvertTo-SiOrchestratorSettingValue -Snapshot $fresh -Value $Value
    $read=Read-SmartM365JsonDocument -Path $fresh.Path
    if ($read.Coexistence -ne 'Single' -or $read.SHA256 -cne $fresh.Hash) { throw 'Configuration changed; reload before saving.' }
    if ($fresh.Present) { $read.Document.($fresh.Field)=$newValue }
    else {
        if ($null -ne $read.Document.PSObject.Properties[$fresh.Field]) { throw 'Setting was created concurrently; reload before saving.' }
        Add-Member -InputObject $read.Document -NotePropertyName $fresh.Field -NotePropertyValue $newValue
    }
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($read.Document | ConvertTo-Json -Depth 100)+[Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $bytes -ExpectedSHA256 $fresh.Hash -Validate {param($d) if ($d -isnot [pscustomobject]) {throw 'Configuration must be an object.'}} | Out-Null
    [pscustomobject]@{Field=$fresh.Field;Path=$fresh.Path}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCADuukJGlhPZhYa
# lkxBeD0qhQiiOx8TrbVe362I098fvqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEupOUPLunazixYceiAtHejwTSDmrjxyb282b7TqnageMA0GCSqG
# SIb3DQEBAQUABIIBgAYY3pOD2iFIZDUbMvwAiWUpHMY8TCzGZW65vOVtLPzmyltC
# Auxnt1pETeTLo2bbnSxGrfqAgFhkxs8aNXIA4h7h+Q4P5fTKqaXBY20HnTTU3WXO
# VWS4PsBoY2cVZDrKgLMf2cfQMGTs5xu0myC3bqtr8PcMVyWmjacS9Lfc8ssTS+Ny
# ptgcXIATPG9y5he0CL3AzAsrKYaNzIfIb5DrFtsJhGz9INmnPEO0TSjKa90UoSdM
# Dy8O8979dsmzS4XInN/whltkayVQ+GoAbkUk7E8RAFnyE4FijTdPkVEy+bUYvPdS
# YFanhvK06Z3/vhixfJL6JDMsorH4R7a8RprBfQye2k52tdzQLRuaSbuVKoC5DOEN
# Yv8qGK9Bi09/DS9SMyATFEqbRDrVA4NJnd0Q+bV9+3XYGc9aQYPxuX70S1hu8pKu
# LEqXh4JLp7LEJVjrOqB5/VmIEJ+dMj+bZKPIrvnWuLl5KP/LA/hoaIum3EwUn+Iu
# wCHndhzR4HbVCb6y4qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjRaMC8GCSqGSIb3DQEJBDEiBCBQQOtUWl5exr8d30stpJcw1rPRP1o1olXlT55h
# eojSQzANBgkqhkiG9w0BAQEFAASCAgCYpoqR2Sqnl1P5TSkormfcnXiMUAe3V5NG
# ErRNFgB32ntj56DP/IE9XOCOdPj6SKhy4vuBgwWCEAM6UllA97l6h+G/WDYuNMen
# REsfpdfPh0tzZrTRnmJgLcSGiZ5XdzVizHgeslOn9sG+TSgKwlf3n9SaTXYfE7Ae
# n0O4YUZulGSGqBxSR4U1MSoB7mXOvoF6rLKYxtqpfeliUx+Cfdbmzkfqq3ibulbP
# CQbFViMoJY1FEmhCc8UbvlIeWq6zRjsb72cRP94CnQpUaYjqZO6YJPp15mJgPFl2
# qXCdqgs4kc1iAhhkVstsOteXfQ7A9Glst2EYgYHMj3g0ekk9vEgRn+IklV99AyFE
# kF8CbqwF7RRHLRB0oLBuZZ8nPmnWRGRtIr4eZvYnBeeBOZPBzQ2Y9IwnZMd49f7w
# ky4hbSZYzKdINOYtnUbJuqTYAFrDO0naNA56OBfFsdtjSQSTJujyZNh7f1uDitsQ
# lMDYfGgMDBdWb513wvqIIIv1z2SrlxHNkjitpNsHEKTEKgFSpZt+sZfkLNA5k8H3
# PoAor/Ik+qlKqbfjA4c3/SWpB9JvrBveMeRwCT0A/rc4+l/KMJmHgtZ11glojiuz
# FlMmvKTOB/Ty3iT/1FpB9tDWo0sNC/Lzb4HAj2XwUfryEqUGbrY3JtHSPhIHK+Jt
# 5jeIFbQO7A==
# SIG # End signature block
