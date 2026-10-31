# Shared local lifecycle for SmartInventory entry points. Compatible with PowerShell 5.1.
<#
.VERSION
1.0.4
#>
function Start-SmartM365InventoryRun {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ScriptPath,[switch]$RequireReturn)
    $previousVariable = Get-Variable SmartM365InventoryRun -Scope Global -ErrorAction SilentlyContinue
    $previous = if ($previousVariable) { @{Value=$previousVariable.Value} } else { $null }
    $previousUploadVariable = Get-Variable EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue
    $previousUpload = if ($previousUploadVariable) { @{Value=$previousUploadVariable.Value} } else { $null }
    $policyVariable = Get-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue
    $previousPolicy = if ($policyVariable) { @{Value=$policyVariable.Value} } else { $null }
    Remove-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue
    $rootsVariable = Get-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue
    $previousRoots = if ($rootsVariable) { @{Value=$rootsVariable.Value} } else { $null }
    Remove-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue
    $state = [pscustomobject]@{
        ScriptPath=$ScriptPath; ScriptName=[IO.Path]::GetFileNameWithoutExtension($ScriptPath)
        StartedAt=Get-Date; ReachedTerminal=$false; InCleanup=$false; ErrorRecord=$null
        Initialized=$false; TranscriptOwned=$false; Finalized=$false; Finalizing=$false; ExitCode=0
        PendingCompletion=$null; PendingBanner=$null
        RequireReturn=[bool]$RequireReturn; ReturnedNormally=$false; ExplicitExit=$false
        GraphOwned=$false; ExchangeOwned=$false; TeamsOwned=$false
        Previous=$previous; PreviousEnableSharePointUpload=$previousUpload; PreviousCollectorUploadPolicy=$previousPolicy; PreviousCollectorUploadRoots=$previousRoots; ManagedCollectorLegacyUpload=$null; PublicationMask=@{}; CleanupWarnings=New-Object 'Collections.Generic.List[string]'
    }
    $global:SmartM365InventoryRun = $state
    return $state
}

function Disable-SmartM365InterruptedPublication {
    param([Parameter(Mandatory)]$State)
    foreach ($name in @('EnableSharePointUpload','EnableTeamsNotifications')) {
        if (-not $State.PublicationMask.ContainsKey($name)) {
            $value = Get-Variable $name -Scope Global -ErrorAction SilentlyContinue
            $State.PublicationMask[$name] = @{Exists=($null -ne $value);Value=$(if($value){$value.Value}else{$null})}
        }
        Set-Variable $name -Value $false -Scope Global
    }
}

function Test-SmartM365InterruptedRun {
    $state=Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    return ($null -ne $state -and ($state.ExitCode -eq 130 -or ($state.InCleanup -and -not $state.ErrorRecord -and
        (-not $state.ReachedTerminal -or ($state.Finalizing -and $state.RequireReturn -and -not $state.ReturnedNormally -and -not $state.ExplicitExit)))))
}

function Invoke-SmartM365RunCleanup {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$State,[Parameter(Mandatory)][string]$Name,[Parameter(Mandatory)][scriptblock]$Action)
    try { & $Action | Out-Null }
    catch {
        $message = "Cleanup failed ($Name): $($_.Exception.Message)"
        $State.CleanupWarnings.Add($message)
        if ($State.Initialized -and (Get-Command WriteLog -ErrorAction SilentlyContinue)) {
            try { WriteLog -Message $message -Level WARNING } catch {}
        }
        # Cleanup reporting must never prevent the other resources being released.
        try { [Console]::Error.WriteLine($message) } catch {}
    }
}

function Set-SmartM365InventoryConnectionOwnership {
    param([ValidateSet('Graph','Exchange','Teams')][string]$Connection,[bool]$Owned)
    $state=Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if($state){$state.($Connection+'Owned')=$Owned}
}

function Close-SmartM365InventoryConnections {
    param([Parameter(Mandatory)]$State)
    foreach($entry in @(@('Graph','Disconnect-MgGraph'),@('Exchange','Disconnect-ExchangeOnline'),@('Teams','Disconnect-MicrosoftTeams'))){
        $kind=$entry[0];$command=$entry[1]
        if($State.($kind+'Owned') -and (Get-Command $command -ErrorAction SilentlyContinue)){
            Invoke-SmartM365RunCleanup $State $kind {
                if($kind -eq 'Exchange'){ & $command -Confirm:$false -ErrorAction Stop | Out-Null }
                else{ & $command -ErrorAction Stop | Out-Null }
                $State.($kind+'Owned')=$false
            }
        }
    }
}

function Start-SmartM365RunTranscript {
    [CmdletBinding()]
    param([string]$Path,[string]$LiteralPath,[switch]$Append,[switch]$Force,[switch]$NoClobber,
        [switch]$IncludeInvocationHeader,[switch]$UseMinimalHeader,[string]$OutputDirectory)
    $result = Microsoft.PowerShell.Host\Start-Transcript @PSBoundParameters
    $state = Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ($state) { $state.TranscriptOwned=$true }
    return $result
}

function Stop-SmartM365RunTranscript {
    [CmdletBinding()]
    param()
    $state = Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ($state -and -not $state.TranscriptOwned) { return }
    $result = Microsoft.PowerShell.Host\Stop-Transcript @PSBoundParameters
    if ($state) { $state.TranscriptOwned=$false }
    return $result
}

function Get-SmartM365InventoryCompletionStatus {
    param([string]$Status,[AllowNull()]$ErrorRecord)
    $state = Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if (-not $state) { return $Status }
    if ($Status -eq 'Cancelled' -or $state.ExitCode -eq 130) { $state.ExitCode=130; Disable-SmartM365InterruptedPublication $state; return 'Cancelled' }
    if (-not $state.InCleanup) {
        $state.ReachedTerminal=$true
        if ($Status -eq 'Failed') { $state.ExitCode=1 }
        return $Status
    }
    if ($state.ErrorRecord -or $ErrorRecord -or $state.ExitCode -notin @(0,3,130)) { if($state.ExitCode -in @(0,3)){ $state.ExitCode=1 }; return 'Failed' }
    if (-not $state.ReachedTerminal -or ($state.Finalizing -and $state.RequireReturn -and -not $state.ReturnedNormally -and -not $state.ExplicitExit)) {
        $errorCount = Get-Variable SmartM365ErrorCount -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        if ($errorCount -gt 0 -or $Status -eq 'Failed') { $state.ExitCode=1; return 'Failed' }
        $state.ExitCode=130
        Disable-SmartM365InterruptedPublication $state
        return 'Cancelled'
    }
    if ($state.Finalizing -and $state.CleanupWarnings.Count -gt 0 -and $Status -in @('Auto','Success')) { return 'CompletedWithWarnings' }
    return $Status
}

function Complete-SmartM365InventoryRun {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$State,[scriptblock]$Cleanup)
    if ($State.Finalized) { return }
    $State.Finalized=$true; $State.InCleanup=$true; $State.Finalizing=$true
    try {
        if (Test-SmartM365InterruptedRun) { Disable-SmartM365InterruptedPublication $State }
        if ($Cleanup) { Invoke-SmartM365RunCleanup $State 'entry resources' $Cleanup }
        Close-SmartM365InventoryConnections $State
        $completion = $State.PendingCompletion
        if ($State.TranscriptOwned -and -not ($completion -and $completion.CloseTranscriptBeforeUpload)) {
            Invoke-SmartM365RunCleanup $State 'transcript' { Stop-SmartM365RunTranscript -ErrorAction Stop }
        }
        $status = Get-SmartM365InventoryCompletionStatus -Status 'Auto' -ErrorRecord $State.ErrorRecord
        $completionCommand=Get-Command Complete-SmartM365ExecutionContext,Complete-CoreSmartM365ExecutionContext -ErrorAction SilentlyContinue | Select-Object -First 1
        # Core banner first: Get-Command keeps the order of the names, and a dot-sourced TenantContext banner
        # would otherwise be chosen even when the prefixed Core one is loaded.
        $bannerCommand=Get-Command Write-CoreSmartM365CompletionBanner,Write-SmartM365CompletionBanner -ErrorAction SilentlyContinue | Select-Object -First 1
        $written = Get-Variable SmartM365ExecutionSummaryWritten -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        if ($State.PendingBanner -and -not $completion) {
            $banner = @{}; foreach($key in $State.PendingBanner.Keys){ $banner[$key]=$State.PendingBanner[$key] }
            $banner.Status = Get-SmartM365InventoryCompletionStatus -Status $banner.Status -ErrorRecord $State.ErrorRecord
            if($State.CleanupWarnings.Count){$banner.WarningCount=[Math]::Max([int]$banner.WarningCount,$State.CleanupWarnings.Count)}
            if($bannerCommand){Invoke-SmartM365RunCleanup $State 'terminal banner' { & $bannerCommand @banner }}
        }
        # Bespoke entry initialization still owns an explicitly deferred completion.
        elseif (($State.Initialized -or $completion) -and -not $written -and $completionCommand) {
            Invoke-SmartM365RunCleanup $State 'terminal summary' {
                $options = @{Status=$status; ErrorRecord=$State.ErrorRecord; FailureStage=$(if($status -eq 'Cancelled'){'Interrupted'}else{'Inventory'})}
                if($completion){ foreach($key in $completion.Keys){$options[$key]=$completion[$key]} }
                if($State.ErrorRecord){$options.ErrorRecord=$State.ErrorRecord}
                $options.Status=Get-SmartM365InventoryCompletionStatus -Status $options.Status -ErrorRecord $options.ErrorRecord
                & $completionCommand @options
            }
        }
        $finalStatus = Get-Variable SmartM365ExecutionStatus -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        if ($State.ErrorRecord -or (($State.Initialized -or $completion) -and $finalStatus -eq 'Failed')) {
            if ($State.ExitCode -in @(0,3)) { $State.ExitCode=1 }
        }
        if ($State.TranscriptOwned) { Invoke-SmartM365RunCleanup $State 'remaining transcript' { Stop-SmartM365RunTranscript -ErrorAction Stop } }
        Close-SmartM365InventoryConnections $State
        if ($State.ExitCode) { [Environment]::ExitCode=$State.ExitCode }
    }
    finally {
        foreach ($name in $State.PublicationMask.Keys) {
            $previous=$State.PublicationMask[$name]
            if ($previous.Exists) { Set-Variable $name -Value $previous.Value -Scope Global }
            else { Remove-Variable $name -Scope Global -ErrorAction SilentlyContinue }
        }
        if ($State.Previous) { Set-Variable SmartM365InventoryRun -Value $State.Previous.Value -Scope Global }
        else { Remove-Variable SmartM365InventoryRun -Scope Global -ErrorAction SilentlyContinue }
        if ($State.PreviousCollectorUploadPolicy) { Set-Variable SmartM365CollectorUploadPolicies -Scope Global -Value $State.PreviousCollectorUploadPolicy.Value }
        else { Remove-Variable SmartM365CollectorUploadPolicies -Scope Global -ErrorAction SilentlyContinue }
        if ($State.PreviousCollectorUploadRoots) { Set-Variable SmartM365CollectorUploadRoots -Scope Global -Value $State.PreviousCollectorUploadRoots.Value }
        else { Remove-Variable SmartM365CollectorUploadRoots -Scope Global -ErrorAction SilentlyContinue }
        if ($State.Previous -and $null -ne $State.Previous.Value) {
            if ($State.PreviousEnableSharePointUpload) { Set-Variable EnableSharePointUpload -Scope Global -Value $State.PreviousEnableSharePointUpload.Value }
            else { Remove-Variable EnableSharePointUpload -Scope Global -ErrorAction SilentlyContinue }
        }
        elseif ($State.ManagedCollectorLegacyUpload) { $global:EnableSharePointUpload = [bool]$State.ManagedCollectorLegacyUpload.Value }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBMidk47PzinVBm
# ZBPRuLyEwVGpXML7VEZfCSo7wynimKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJLM1i2w3YlX3w1wxDN411tuLLkpGts/ELhPVbUI3sC1MA0GCSqG
# SIb3DQEBAQUABIIBgCCoE4TYyx9JF0SZUntpRYjEJ2pyEB/nW/PjH3ZbVxf75OkT
# xT9xqv1hq55NanWMUuMpKLo8zMFRAZAufcHOFtfLJY0s1FUDA9ix49sX2fOvNDcJ
# FRFS+QCm2kSbVvzYrotrakbji97OwPvU96HJOi8QZSQjaJ4EvxE8zHJyEV5CNbgp
# HzWyGYaIQv/yg0to0xfGSATAA4tpnPiMMq9YTK+nx/XNpcBXK7yuXuq4Q4BHe+HZ
# kO6Tv5l/XiK7nj7jR5uHBqHM+QDmBIQeaS58VX12vH2511zfCel76SqV3S1QGkgN
# CIXR5wt7LMmWD0vISJZ37VLLCaOIS6u1Ud05bNnmC2gplhrWJH2JlhXRufovSHJg
# tZUpVOnxaKHQNzHTCJUk+gYUFsisunJ62/k20nBbiUdjMVJzlTvkLXZckTJLvalA
# FWGFj3s2KmxT2NGlBRA8KhlN2qx1opPANW6gXxhN5IrsYSOMKUK7WWINrCnAPxDB
# gNRspQbXl69YBlQUjaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExNDM3
# MjNaMC8GCSqGSIb3DQEJBDEiBCDfPcqE3ExmNRBfcXkfCE+YJ3hcXmCopls+j9mS
# nthOEzANBgkqhkiG9w0BAQEFAASCAgBcXF+ESGCHmVa1cKUUEYugzmHmQTGk9jxU
# Nupl9i50xMQP/BFh+uSmtSqM5M1ulGg5tLnqgfVs2QWkfaHglI/fHrGilMbirzHE
# UDH75haP55aeQujCuhpp9NXlK2NIBPV7ArBOCPoUMjdee4wrqMqOvXC63cqcdPsW
# p5HabKFokyj2OjSQr5Oa3VFsh5vBYMAVzCtw9mIhRRaKMXHXY5RSVdBALdlqGvU1
# D/u4uo1hA4Pf8kBip8wk9UhyiBPPvofUj7xKaPvAuTiPCb4+AByM0+rlW1n2uHNI
# p3oxOYc1EyxygIRk8M34PK5yyhwkNJtuPxh73qAfqOGXGBmGxTDmHqVTgvq9qZ2C
# NzfbnWloCHzMjujyZf/O/9bxx2ce/l33sLu6Us3rv8lQ6OYXYN/LQYz4nQ1IKAfq
# uhTQRCYpI6R9g8ZCoFje27IHnBnys5eJsxW0zRbjr30srdCpn3YsQ1l4rMEfrzAa
# tFQmKtLNm//FTfMUEblalxKpH3moMJbUZaRnyWpMgwFhCbxaJgcYVcTYhnJkqSMN
# XTL6IRA2cK1ltC1W/xlZ+m3C2DwRj5Dg5i9vg9Ob3/Z1tjhci5TchRO+zulla0a3
# mTWc1b9bswfQyVjVznzHeNuUEYiDuh7JXVMU1vhFkr2m2y+pApAQqJFRpW8nlBtp
# Z7Wwf5DZgw==
# SIG # End signature block
