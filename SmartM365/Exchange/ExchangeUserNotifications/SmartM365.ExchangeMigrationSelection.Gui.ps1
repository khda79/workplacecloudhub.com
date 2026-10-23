<#
.SYNOPSIS
WPF migration batch selection handlers. Only explicit connect starts the EXO worker.
.VERSION
1.0.0
#>
$script:ExoMigration = [pscustomobject]@{Process=$null;Timer=$null;Output=$null;Error=$null;Connected=$false;Busy=$false
    TenantId='';TenantKey='';Account='';Batches=@();Rows=@();ReadAtUtc='';SelectionReadAtUtc=''}
$script:MigrationSelectionPath = ''
$script:MigrationCampaignRunning = $false
function Get-MigrationSelectionContext {
    $key = Get-SelectedTenantProfile
    $root = Get-SmartM365RootPath
    $map = @{}
    foreach ($path in @((Join-Path $root 'Config/SmartM365.global.local.json.template'),(Join-Path $root 'Config/SmartM365.global.local.json'),(Join-Path $root "Config/Tenants/$key.local.json"))) {
        Add-GuiConfigValue -Map $map -Path $path
    }
    $id = [string]$map['TenantId']; $guid = [guid]::Empty
    if (-not [guid]::TryParse($id,[ref]$guid) -or $guid -eq [guid]::Empty) { throw 'Selected profile must supply a valid TenantId GUID' }
    return [pscustomobject]@{TenantKey=$key;TenantId=$id}
}
function Test-MigrationSelectionSource {
    return $script:SelectedCampaignIndex -eq 0 -and $script:MigrationRecipientSourceCombo.SelectedIndex -eq 1
}
function Clear-MigrationSelection {
    $script:MigrationSelectionPath=''
    $script:ExoMigration.Rows=@(); $script:ExoMigration.SelectionReadAtUtc=''
    $script:ExoRecipientsGrid.ItemsSource=@()
}
function Show-MigrationBatchList {
    $script:ExoBatchesGrid.ItemsSource=@($script:ExoMigration.Batches | Where-Object {
        $script:ExoShowCompletedCheck.IsChecked -or $_.Status -notmatch '^Completed|^Removed|^Removing'
    })
}
function Show-MigrationWorkerState {
    $state=$script:ExoMigration
    $campaignRunning=$script:MigrationCampaignRunning -or $null -ne $script:CurrentProcess
    $script:ExoConnectButton.IsEnabled=(-not $state.Busy -and -not $campaignRunning)
    $script:ExoRefreshButton.IsEnabled=($state.Connected -and -not $state.Busy -and -not $campaignRunning)
    $script:ExoLoadButton.IsEnabled=$script:ExoRefreshButton.IsEnabled
    $script:ExoUseSelectionButton.IsEnabled=(-not $state.Busy -and -not $campaignRunning -and $state.Rows.Count -gt 0)
    $script:ExoApplyDateButton.IsEnabled=$script:ExoUseSelectionButton.IsEnabled
    $script:ExoCancelButton.IsEnabled=$null -ne $state.Process
    $script:ExoBatchesGrid.IsEnabled=(-not $state.Busy -and -not $campaignRunning)
    $script:ExoRecipientsGrid.IsEnabled=$script:ExoBatchesGrid.IsEnabled
    $script:ExoTimeZoneCombo.IsEnabled=$script:ExoBatchesGrid.IsEnabled
    $script:ExoOverrideDate.IsEnabled=$script:ExoBatchesGrid.IsEnabled
    $script:ExoShowCompletedCheck.IsEnabled=$script:ExoBatchesGrid.IsEnabled
    $script:MigrationRecipientSourceCombo.IsEnabled=-not $campaignRunning
}
function Stop-MigrationSelectionWorker {
    $state=$script:ExoMigration
    if ($state.Timer) { $state.Timer.Stop(); $state.Timer=$null }
    if ($state.Process) {
        try {
            if (-not $state.Process.HasExited) {
                $state.Process.StandardInput.WriteLine('{"Operation":"Close"}')
                if (-not $state.Process.WaitForExit(1000)) { $state.Process.Kill($true); $state.Process.WaitForExit() }
            }
        }
        catch { Write-Verbose $_.Exception.Message }
        finally { $state.Process.Dispose(); $state.Process=$null }
    }
    $state.Connected=$false; $state.Busy=$false; $state.Batches=@(); $state.Account=''
    $script:ExoBatchesGrid.ItemsSource=@()
    Clear-MigrationSelection
    $script:ExoConnectionText.Text='Disconnected. No selection is active.'
    Show-MigrationWorkerState
}
function Send-MigrationWorkerRequest {
    param([ValidateSet('Batches','Recipients')][string]$Operation)
    $state=$script:ExoMigration
    if (-not $state.Connected -or $state.Busy -or $script:CurrentProcess) { return }
    $context=Get-MigrationSelectionContext
    if ($context.TenantKey -ne $state.TenantKey -or $context.TenantId -ine $state.TenantId) { Stop-MigrationSelectionWorker; throw 'Tenant changed; connect again' }
    $request=@{Operation=$Operation}
    if ($Operation -eq 'Recipients') {
        [void]$script:ExoBatchesGrid.CommitEdit(); [void]$script:ExoBatchesGrid.CommitEdit()
        $ids=@($script:ExoBatchesGrid.ItemsSource | Where-Object Selected | ForEach-Object Identity)
        if (-not $ids.Count) { throw 'Select at least one active batch' }
        $request.BatchIds=$ids; $request.TimeZoneId=[string]$script:ExoTimeZoneCombo.SelectedItem
    }
    Clear-MigrationSelection
    $state.Busy=$true
    $script:ExoConnectionText.Text='Reading EXO. Cancel remains available; no emails are sent.'
    $state.Process.StandardInput.WriteLine(($request | ConvertTo-Json -Compress))
    $state.Process.StandardInput.Flush()
    Show-MigrationWorkerState
}
function Receive-MigrationWorkerResponse {
    param($Response)
    $state=$script:ExoMigration
    switch ($Response.Kind) {
        'Connected' {
            if ([string]$Response.Data.TenantId -ine $state.TenantId) { Stop-MigrationSelectionWorker; throw 'Worker tenant mismatch' }
            $state.Connected=$true; $state.Busy=$false; $state.Account=[string]$Response.Data.Account
            Send-MigrationWorkerRequest 'Batches'
        }
        'Batches' {
            $state.Batches=@($Response.Data.Rows); $state.ReadAtUtc=[string]$Response.Data.ReadAtUtc
            $state.Busy=$false; Show-MigrationBatchList
            $script:ExoConnectionText.Text="Account: $($state.Account) | Tenant: $($state.TenantId) | Read: $($state.ReadAtUtc) | $($state.Batches.Count) batch(es)"
        }
        'Recipients' {
            $state.Rows=@($Response.Data.Rows); $state.SelectionReadAtUtc=[string]$Response.Data.ReadAtUtc
            $state.Busy=$false
            $campaign=$script:Campaigns[0]
            $registryPath=Resolve-GuiTokenizedPath -Value (Get-CampaignConfigValue -Campaign $campaign -Name 'SentRegistryPath') -Campaign $campaign
            if ([string]::IsNullOrWhiteSpace($registryPath) -or $registryPath -match '\{\{') { Clear-MigrationSelection; throw 'Cannot resolve sent registry; recipient eligibility is unknown' }
            $registry=Load-SmartM365CommunicationSentRegistry -Path $registryPath
            foreach ($row in $state.Rows) {
                if ($registry.ContainsKey([string]$row.EmailAddress)) {
                    $row.AlreadySent=$true; $row.Selected=$false; $row.Exclusion='Already notified (Force send required)'
                }
            }
            $script:ExoRecipientsGrid.ItemsSource=$state.Rows
            $script:ExoConnectionText.Text="$($state.Rows.Count) unique recipient(s) | Read: $($state.SelectionReadAtUtc) | $($script:ExoTimeZoneCombo.SelectedItem) | Review dates and exclusions"
        }
        'Error' {
            $state.Busy=$false; Clear-MigrationSelection
            $script:ExoConnectionText.Text='EXO read failed: ' + [string]$Response.Data.Message
            Add-OutputLine $script:ExoConnectionText.Text
        }
        default { throw 'Invalid worker response' }
    }
    Show-MigrationWorkerState
}
function Update-MigrationSelectionWorker {
    $state=$script:ExoMigration
    try {
        foreach ($pipe in @($state.Output,$state.Error)) {
            $lines=0
            while ($pipe -and -not $pipe.Closed -and $pipe.Pending.IsCompleted -and $lines -lt 100) {
                $text=$pipe.Pending.GetAwaiter().GetResult(); $lines++
                if ($null -eq $text) { $pipe.Closed=$true; break }
                $pipe.Pending=$pipe.Reader.ReadLineAsync()
                if ($text.StartsWith('EXOBATCH:')) { Receive-MigrationWorkerResponse ($text.Substring(9) | ConvertFrom-Json) }
                elseif ($text) { Add-OutputLine ('EXO worker: ' + $text) }
            }
        }
        if ($state.Process -and $state.Process.HasExited -and $state.Output.Closed -and $state.Error.Closed) {
            Stop-MigrationSelectionWorker
            $script:ExoConnectionText.Text='EXO worker closed. Connect again to retrieve a fresh selection.'
        }
    }
    catch {
        $message=$_.Exception.Message
        Stop-MigrationSelectionWorker
        $script:ExoConnectionText.Text='EXO worker failed: ' + $message
        Add-OutputLine $script:ExoConnectionText.Text
    }
}
function Start-MigrationSelectionWorker {
    if ($script:CurrentProcess) { return }
    Stop-MigrationSelectionWorker
    $context=Get-MigrationSelectionContext
    $state=$script:ExoMigration; $state.TenantId=$context.TenantId; $state.TenantKey=$context.TenantKey
    $psi=[Diagnostics.ProcessStartInfo]::new(); $psi.FileName=Get-PowerShellPath
    foreach ($arg in @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',(Join-Path $script:GuiScriptRoot 'SmartM365-ExchangeMigration-SelectionWorker.ps1'),'-TenantId',$context.TenantId)) { $psi.ArgumentList.Add($arg) }
    $psi.UseShellExecute=$false; $psi.CreateNoWindow=$true
    $psi.RedirectStandardInput=$true; $psi.RedirectStandardOutput=$true; $psi.RedirectStandardError=$true
    $state.Process=[Diagnostics.Process]::Start($psi)
    $state.Output=[pscustomobject]@{Reader=$state.Process.StandardOutput;Pending=$state.Process.StandardOutput.ReadLineAsync();Closed=$false}
    $state.Error=[pscustomobject]@{Reader=$state.Process.StandardError;Pending=$state.Process.StandardError.ReadLineAsync();Closed=$false}
    $state.Busy=$true
    $state.Timer=[Windows.Threading.DispatcherTimer]::new(); $state.Timer.Interval=[timespan]::FromMilliseconds(150)
    $state.Timer.Add_Tick({Update-MigrationSelectionWorker}); $state.Timer.Start()
    $script:ExoConnectionText.Text='Complete the interactive EXO sign-in. The selected tenant will be verified.'
    Show-MigrationWorkerState
}
function Set-MigrationRecipientDate {
    [void]$script:ExoRecipientsGrid.CommitEdit(); [void]$script:ExoRecipientsGrid.CommitEdit()
    $date=$script:ExoOverrideDate.SelectedDate
    if ($null -eq $date) { throw 'Choose an explicit notification date first' }
    foreach ($row in $script:ExoMigration.Rows | Where-Object Selected) {
        $row.EffectiveDate=([datetime]$date).ToString('yyyy-MM-dd'); $row.DateConfirmed=$true
        $row.DateSource='Operator override'; $row.ScheduledUtc=''; $row.ScheduledLocal=''
        $row.ScheduleReason='Date explicitly supplied by the operator; migration schedule was not modified'
    }
    $script:MigrationSelectionPath=''; $script:ExoRecipientsGrid.Items.Refresh()
}
function Save-MigrationRecipientSelection {
    if ($script:MigrationRecipientSourceCombo.SelectedIndex -ne 1) {
        $script:MigrationRecipientSourceCombo.SelectedIndex=1
    }
    $state=$script:ExoMigration
    if ($state.Busy) { throw 'Wait for the read operation to finish' }
    $context=Get-MigrationSelectionContext
    if ($context.TenantKey -ne $state.TenantKey -or $context.TenantId -ine $state.TenantId) { throw 'Tenant changed; reconnect and reload recipients' }
    $readAt=[datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse($state.SelectionReadAtUtc,[ref]$readAt) -or ([datetimeoffset]::UtcNow-$readAt).TotalMinutes -gt 15) { throw 'Recipient snapshot expired; reload recipients (maximum age: 15 minutes)' }
    [void]$script:ExoRecipientsGrid.CommitEdit(); [void]$script:ExoRecipientsGrid.CommitEdit()
    $rows=@($state.Rows | Where-Object Selected)
    $registryPath=Resolve-GuiTokenizedPath -Value (Get-CampaignConfigValue -Campaign $script:Campaigns[0] -Name 'SentRegistryPath') -Campaign $script:Campaigns[0]
    if (-not $registryPath -or $registryPath -match '\{\{') { throw 'Sent registry path is unresolved' }
    $registry=Load-SmartM365CommunicationSentRegistry -Path $registryPath
    foreach($row in $rows) { $row.AlreadySent=$registry.ContainsKey([string]$row.EmailAddress) }
    $root=Resolve-GuiTokenizedPath -Value (Get-CampaignConfigValue -Campaign $script:Campaigns[0] -Name 'RunOutputPath') -Campaign $script:Campaigns[0]
    if (-not $root -or $root -match '\{\{') { throw 'Selection output path is unresolved' }
    $path=Join-Path $root ('Selections/MigrationSelection-' + [guid]::NewGuid().ToString('N') + '.csv')
    Export-ExoMigrationSelection -Rows $rows -TenantId $state.TenantId -TenantKey $state.TenantKey -Path $path -SourceReadAtUtc $state.SelectionReadAtUtc -ForceSend:$script:ForceSendCheck.IsChecked | Out-Null
    $script:MigrationSelectionPath=$path; $script:RecipientPathBox.Text=$path
    $script:ExoConnectionText.Text="$($rows.Count) recipient(s) reviewed. CSV prepared; no email sent. Dry Run remains the default."
    Show-CommandPreview
}
function Prepare-MigrationSelectionRun {
    if (-not (Test-MigrationSelectionSource)) { return $true }
    try { Save-MigrationRecipientSelection; return $true }
    catch { Add-OutputLine $_.Exception.Message; Show-Status -Text $_.Exception.Message -Color '#B42318'; return $false }
}
function Show-MigrationSelectionSource {
    $migration=$script:SelectedCampaignIndex -eq 0
    $script:MigrationRecipientSourcePanel.Visibility=if($migration){'Visible'}else{'Collapsed'}
    $script:ExoMigrationTab.Visibility=if($migration){'Visible'}else{'Collapsed'}
    $exo=Test-MigrationSelectionSource
    if ($migration) {
        $script:RecipientPathBox.IsEnabled=-not $exo
        $script:BrowsePathButton.IsEnabled=-not $exo
        $script:BrowseFolderButton.IsEnabled=-not $exo
        $script:EffectiveDatePicker.IsEnabled=-not $exo
        $script:ExchangeManagementModeCombo.IsEnabled=-not $exo
    }
    if ($exo) { $script:RecipientPathHelpText.Text='Review EXO batches and recipients in the EXO migration batches tab. Dates are preserved per recipient.' }
    Show-MigrationWorkerState
    Show-CommandPreview
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDHbmJ5Tnm3FIVU
# jcArG8a1RttMmGh1pCO80NhlgNdBvKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHp4pz5cmdeNROGN0+URHZw3cP7B67rwwzIZKi+R9fBFMA0GCSqG
# SIb3DQEBAQUABIIBgKnNvQjAeg/pETNUJZXhExaf4DSYV0tTUKlqZzsYfZhuXT35
# VgJXsWcNDtvaQX3xdwiRu1QPvaXiQmDVI/rOgdWrWgmAtXaEFRm7s67EPGaS8Cmf
# T3h5H4B/b9zQWvsHenObdqsqMOGTGFoMhcoKk8RMZqFyHbFxAdQmYDvnCEQoix9p
# 6x4GaTE6HxA2iKcucSXJtc5r1bYui9R6ETuWrdq7U7wR2dcYuCsEoCUR5swObKC4
# iYsIFI9FeriaNM0XbIniVpD+gUcp1sEOOSAptuV7FzyiK1UJvA3pBDLepgIRQGC1
# 3xbe9DoRtnDubr8nr6v+HEzSCg8Ays+MVVvGtKXJHwzM/bddQ3L7I0K2HxFfo/x+
# AKT6Wyp9+2JbQjXypmj3izqhIFGkctUADC2JZfIwM1FV45stR8F2EFfbU3ZCuFXs
# vb40qvMz9Pc866djqHWG/h8FLPQQAU1YC3gz++XgGIJ97Tsrl8ZT8OKR3MHV5UL2
# bWoDYLZK/ShzkVYe3KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjMxNzU4
# MThaMC8GCSqGSIb3DQEJBDEiBCCh/+0t4FnCmIFeeBDbkQtesDRGXugEc1fVuE1E
# WMj6ZzANBgkqhkiG9w0BAQEFAASCAgAZGm9OpCVyqtkFVvemxJGnbJLX8q8APlCV
# 3e4X7BhXjYhUqAdb9i3J7ikW9I99qyCgZDFSUhA5tZEl/jija5KDLJCcM87ZiMrs
# 0SxlWa6z4cUb6r6O7iDQsyHCLIchVxLeexAGZWoEogJ3/zF0fRGOSaFXt+eGh+9f
# W755mNbgADbT4KtxVC8JKwiN7uAXwX1udW3C0RxE9yubkeognDZ5xfTyYp4BGJDL
# oSGC4KnnSKFM5mYth32g/hzt1u1U8u6uE1sR/nMP4bqm3CMFTrmHGSOqr7/mPTd3
# 9l8rwbcksTo2VF9E9c6rGaf+kSZueE0t48vhECqc1Z2sf3z8veqMw/vD+cc4GAom
# cD72j+2ME681Uf/wdEQTWbTwAt/ltluEqNNpjzDSIFO9Huqpzkq6NUavTRIScfX2
# vcBG7TehOA0f43u+bs5Lv1JI2jKHopVRMQo24dCTbGvCLgzHD88zeb7HFKao3WUW
# yfRsO4MzoikysuMCzPIpjbvdolGhUHjTd+0KYZmLxUaDY10cOqHg92sIE3YgC3Iz
# Z4c380mtI0dVyYslhhh2sQ1kCXWmuivHQGOjJf+rheldi9V2vTcfkuJ/ln4umpgP
# vP7uB/LUP8h2drngCdkVwxV6yuKAbymx7cajHYNBkPj2MgPlm2u4w+y5t4jLUiPz
# xAW+zlCxBQ==
# SIG # End signature block
