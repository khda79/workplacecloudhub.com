function Stop-SiTenantSignInProcess {
    param([Parameter(Mandatory)]$Process)
    if (-not $Process.HasExited) {$Process.Kill();if (-not $Process.WaitForExit(5000)) {throw 'Sign-in process termination is uncertain.'}}
}
function Initialize-SiTenantSignInDialog {
    param($Dialog,$State,[string]$LibraryRoot,[scriptblock]$Changed,[scriptblock]$InstallPrerequisites,[bool]$OfflineFixture=$false,[scriptblock]$StartWorker,[scriptblock]$ConfirmTenant,[scriptblock]$ModuleAvailable)
    $Dialog.Add_Closing({param($sender,$eventArgs)
        if ($null -ne $State.Process) {
            try {
                Stop-SiTenantSignInProcess -Process $State.Process
                if ($null -ne $State.Timer) {$State.Timer.Stop();$State.Timer=$null}
                $State.Process.Dispose();$State.Process=$null
                if ($State.Path) {Remove-Item -LiteralPath $State.Path -Force -ErrorAction SilentlyContinue};if ($State.ProgressPath) {Remove-Item -LiteralPath $State.ProgressPath -Force -ErrorAction SilentlyContinue}
            } catch {$eventArgs.Cancel=$true;$Dialog.FindName('TenantStatus').Text='Sign-in could not be stopped yet.'}
        }
    }.GetNewClosure())
    $Dialog.FindName('DetectTenant').Add_Click({
        try {
            $moduleFound=if ($null -ne $ModuleAvailable) {[bool](& $ModuleAvailable)} else {[bool](Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Select-Object -First 1)}
            if (-not $OfflineFixture -and -not $moduleFound) {
                if ([Windows.MessageBox]::Show($Dialog,'Microsoft Graph Authentication is required for sign-in. Install the machine prerequisites now?','Microsoft 365 sign-in','YesNo','Question') -ne 'Yes') {return}
                & $InstallPrerequisites
                if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Select-Object -First 1)) {$Dialog.FindName('TenantStatus').Text='Microsoft Graph Authentication is still unavailable. Install prerequisites and try again.';return}
            }
            $State.Tenant='';$State.Account='';$State.Mail='';$State.SiteHost='';$State.InitialDomain='';$State.MissingScopes=@();$State.PreflightReady=$false;$State.PreflightMessage='';& $Changed
            if ($OfflineFixture) {$State.Tenant='11111111-2222-3333-4444-555555555555';$State.Account='user@synthetic.invalid';$State.Mail='admin@synthetic.invalid';$State.SiteHost='synthetic.sharepoint.invalid';$State.InitialDomain='synthetic.onmicrosoft.com';$State.MissingScopes=@();$State.PreflightReady=$true;$State.PreflightMessage='Synthetic preflight passed.';$Dialog.FindName('TenantStatus').Text="Tenant ID: $($State.Tenant)`nSigned in as: $($State.Account)";& $Changed;return}
            $State.Path=[IO.Path]::GetTempFileName();$State.ProgressPath=[IO.Path]::GetTempFileName()
            $worker=Join-Path $LibraryRoot 'SmartM365.TenantDiscoveryWorker.InitialDomain.ps1'
            $quote={param([string]$v) "'"+$v.Replace("'","''")+"'"}
            $command='& '+(& $quote $worker)+' -ResultPath '+(& $quote $State.Path)+' -ProgressPath '+(& $quote $State.ProgressPath)+'; exit $LASTEXITCODE'
            $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
            if ($null -ne $StartWorker) {$State.Process=& $StartWorker $worker $State.Path $State.ProgressPath} else {$State.Process=Start-Process -FilePath (Join-Path $PSHOME 'pwsh.exe') -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru -ErrorAction Stop}
            if ($null -eq $State.Process) {throw 'Sign-in process could not be started.'}
            $Dialog.FindName('DetectTenant').IsEnabled=$false
            $Dialog.FindName('TenantStatus').Text='Preparing Microsoft 365 sign-in...'
            $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(500);$State.Timer=$timer
            $stateRef=$State;$dialogRef=$Dialog;$changedRef=$Changed;$confirmRef=$ConfirmTenant
            $timer.Add_Tick({
                if (-not $stateRef.Process.HasExited) {try {$progress=Get-Content -LiteralPath $stateRef.ProgressPath -Raw -ErrorAction Stop;if ($progress) {$dialogRef.FindName('TenantStatus').Text=$progress}} catch {};return}
                $stateRef.Timer.Stop()
                try {
                    if ($stateRef.Process.ExitCode -ne 0) {throw 'Microsoft 365 sign-in or tenant detection failed.'}
                    $result=Get-Content -LiteralPath $stateRef.Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    $id=[guid]::Empty
                    if (-not [guid]::TryParse([string]$result.TenantId,[ref]$id) -or $id -eq [guid]::Empty -or [string]::IsNullOrWhiteSpace([string]$result.Account)) {throw 'The authenticated tenant could not be determined.'}
                    $detected=$id.ToString();$account=[string]$result.Account
                    $dialogRef.FindName('TenantStatus').Text="Tenant ID: $detected`nSigned in as: $account"
                    if (($null -ne $confirmRef -and (& $confirmRef $detected $account)) -or ($null -eq $confirmRef -and [Windows.MessageBox]::Show($dialogRef,"Tenant ID: $detected`nSigned in as: $account`n`nUse this Microsoft 365 tenant?",'Confirm Microsoft 365 tenant','YesNo','Question') -eq 'Yes')) {$stateRef.Tenant=$detected;$stateRef.Account=$account;$stateRef.Mail=if ($result.Mail -is [string]) {[string]$result.Mail} else {''};$stateRef.InitialDomain=if ($result.InitialDomain -is [string]) {[string]$result.InitialDomain} else {''};$stateRef.MissingScopes=@($result.MissingScopes);$stateRef.PreflightReady=($null -ne $result.Readiness -and $result.Readiness.Ready -eq $true);$stateRef.PreflightMessage=if ($stateRef.PreflightReady) {'Directory roles and group creation policy checked. Bootstrap reconnects and can still encounter restrictions.'} else {(@($result.Readiness.Messages) -join ' ')};$stateRef.SiteHost=if ($result.SiteHost -is [string]) {[string]$result.SiteHost} else {''};if (-not $stateRef.InitialDomain) {$dialogRef.FindName('TenantStatus').Text+="`nInitial Microsoft 365 domain could not be detected. Reconnect and check directory permissions."};if (-not $stateRef.SiteHost) {$dialogRef.FindName('TenantStatus').Text+="`nSharePoint root site could not be read. Reconnect and grant the requested permissions."};if ($stateRef.MissingScopes.Count) {$dialogRef.FindName('TenantStatus').Text+="`nMissing Graph permissions: $($stateRef.MissingScopes -join ', ')"} else {$dialogRef.FindName('TenantStatus').Text+="`nGraph delegated permissions granted. $($stateRef.PreflightMessage)"}}
                    else {$dialogRef.FindName('TenantStatus').Text+='`nTenant not confirmed. Connect again to select another tenant.'}
                } catch {$dialogRef.FindName('TenantStatus').Text='Tenant detection failed. Connect again and check the browser or Windows sign-in window.'}
                finally {Remove-Item -LiteralPath $stateRef.Path,$stateRef.ProgressPath -Force -ErrorAction SilentlyContinue;$stateRef.Process.Dispose();$stateRef.Process=$null;$stateRef.Timer=$null;$dialogRef.FindName('DetectTenant').IsEnabled=$true;$dialogRef.FindName('Cancel').IsEnabled=$true;& $changedRef}
            }.GetNewClosure())
            $timer.Start()
        } catch {if ($State.Path) {Remove-Item -LiteralPath $State.Path -Force -ErrorAction SilentlyContinue};if ($State.ProgressPath) {Remove-Item -LiteralPath $State.ProgressPath -Force -ErrorAction SilentlyContinue};$Dialog.FindName('TenantStatus').Text='Could not start Microsoft 365 sign-in: '+$_.Exception.Message;& $Changed}
    }.GetNewClosure())
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDxzyHcTekX+rq+
# +yPKZ9rZMynDOtsnRR4l7HthgvjeE6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIuH41kWQMH4VuPRN5dNzH2haa3mbJ7YPLxMVJauDn2XMA0GCSqG
# SIb3DQEBAQUABIIBgGumj8TSldD8PTBnhf4+ExPSHM0tzAXxRLSeqyY4LqG7tiMF
# U09FC1ZHYil/BpXd/m2wbQx7a2sYIbaYd+BES9vdAmuvvSrH2LpA4AKZhtq0nTCd
# XQ8xWzLX7QBmuKv+sp0X0367bDa93a6PWN7xv5xRY3PvoxI/CtvUrZrIhrE0yBVf
# ScvoraFrcavmNB/NZuYfmnPZGoB7+y5gtkYAMOq3OOLSZq5kkKvHRAFfxHi8eRJv
# XmsKgrCCnvt4pC/b0wjVh/jhaFD8X1AubhSC15VMEBQCyOkBHKvNQn00hTABrcKn
# rIEzs08IgnVkLa1YiZ4rUL0Uh9YZbiLd4umuD0X/HpT9BnMRCvLtUS1QmVpZTHvp
# R34c2ufIO5ZldtwZZ0weHvuSbiimwiouxZC/+tYlNW73lHV58DOw5nLWqANk3+KK
# hpEFiIA7R1qGz8d/lkIpw/6cZrEVILIhUHv9/awUzguD0HP1MGc4pVUifvfPVFBv
# XbWrdcd5U8k9r7RJvKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjBaMC8GCSqGSIb3DQEJBDEiBCCY3Gi9I4+kcuaJFZyk729zXtfNeC//BhGGMLIU
# tg0mtzANBgkqhkiG9w0BAQEFAASCAgB1OoQ83ZRvP9JehLO+aSPr+PxF0Gcbxy1i
# aApeDWAns5ms0XpSSV63XiRevt/l1IApT0Vi0Y2AX3KOz0vmpA0YwK/g4mKTGhcW
# gTsOxuXpXbKvutogKDMuHTJsZHRIPK55m91qlos7I4C4d6x152cc+0qckAmtgN7s
# GnRdAaMH8PZjTU/sr4vbrK8laijkefPaxG6TnWH0WzzHBKwutfWLkhkvFJ2lQ/If
# MKrvoq8HTbAajEgChZFwwe5XWv1LtOxRxiGGHE+y9CbhVGhsPZvEh86omdMea7OA
# G15C4jP0/SFlPescsnnRstdy4divsWsYgCTNAM3VDjVXm1ohM13xJWTO1jQPP8UT
# Emo3aQ9TFz6ol4Wtjyu/yKUXAf1iyDoYrSQUdWvpsRNTxakcLfsfH97mlMU5IbVy
# cnN4B9btSG21S8V/QpRKmzX61pCi5bxLK2e8GtYqF95oqf+ey9KbLDgbdmwwUSQt
# 9YGK2hjhFEskVcnBddRsHywE/x8DhN6zj77rGK0z6/eWzuR74vF0pscxK049l7lR
# fJPLIsCa031rsUchFMChXfvLFh5U6NQYGANUBJ9kN6zQPXko+RyxiLyMcQCfm9eS
# /sxvfWBTkmtGnvPhcGworK3MTweY0W9g2rJFn/AqRj8hbOhs+Ky0AJ6xdJ6uHiXC
# 8986LKT27A==
# SIG # End signature block
