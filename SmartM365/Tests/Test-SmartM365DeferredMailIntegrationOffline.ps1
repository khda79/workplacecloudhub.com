# Extracts the actual mail entrypoints. Authentication and transport are mocked.
$ErrorActionPreference='Stop'
$core=Join-Path $PSScriptRoot '../Modules/SmartM365.Core'
. (Join-Path $core 'SmartM365-MailAdminRouting.ps1')
. (Join-Path $core 'SmartM365-MailMaintenance.ps1')
. (Join-Path $core 'SmartM365-DeferredMail.ps1')
$aclVerifier=${function:Test-SmartM365DeferredMailDirectory}
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DeferredIntegration-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $root
$script:cfg=[pscustomobject]@{ProfileKey='demo';LogAllRootPath=$root;ReportMailSchedule=[pscustomobject]@{Enabled=$true;NotBefore='07:30';TimeZone='Romance Standard Time';QueuePath=$root;TrustedSids=@('synthetic')}}
$script:sent=0
function Get-SmartM365MailUtcNow {$script:clock}
function Get-SmartM365EffectiveModuleGlobalConfig {$script:cfg}
function Resolve-SmartM365ConfigValue {param($Value) $Value.Replace('{{WorkspaceRootPath}}','C:\SyntheticWorkspace').Replace('{{TenantKey}}','synthetic-demo')}
function Test-SmartM365DeferredMailDirectory {param($Path,$TrustedSids) $Path -eq $root}
function Get-PSCallStack {[pscustomobject]@{ScriptName='C:\synthetic\SmartInventory\Collector.ps1'}}
function Get-ModuleLocalConfig {[pscustomobject]@{}}
$script:callerConfig=@{}
function Get-SmartM365CallerLocalConfig { $script:callerConfig }
function Get-ModuleLocalConfigValue {param($Config,$Name,$DefaultValue) $DefaultValue}
function Test-SmartM365InterruptedRun {$false}
function WriteLog {param($Message,$Level)}
function Add-SmartM365MaxItemsSubjectPrefix {param($Subject) $Subject}
function Format-SmartM365MailSubject {param($Subject) $Subject}
function Add-SmartM365MaxItemsMailBanner {param($BodyHtml) $BodyHtml}
function ConvertToRecipientArray {param($Recipients) @($Recipients -split ';' | Where-Object {$_})}
function Add-SmartM365MailFilesSection {param($BodyHtml,$Files) $BodyHtml}
function Convert-SmartM365MailBodyLocalPathsToSharePointLinks {param($BodyHtml) $BodyHtml}
function ConvertTo-SmartM365EmailBody {param($BodyHtml,$Subject,$Category) $BodyHtml}
function Save-SmartM365MailHtmlCopy {param($Subject,$BodyHtml)}
function Resolve-SmartM365MaintenanceMailRecipientRoute {param($To,$Cc,$MailPurpose) [pscustomobject]@{To=$To;Cc=$Cc}}
function Set-SmartM365SmtpMaintenanceRecipient {param($MailParameters,$MailPurpose)}
function Send-SmartM365GraphMail {param($From,$To,$Cc,$Subject,$BodyHtml,$Attachments,$AllowAttachments,$SuppressAttachmentLinks,[switch]$SkipHtmlCopy,$MailPurpose)
 if($script:failGraph){throw 'Simulated Graph failure'}
 if ($Attachments.Count -and [IO.File]::ReadAllText($Attachments[0]) -ne 'snapshot') {throw 'Attachment snapshot changed'}
 $script:sent++
}
function Send-MailMessage {param($SmtpServer,$Port,$From,$To,$Cc,$Subject,$Body,$BodyAsHtml,$Attachments,$ErrorAction)
 if($Attachments.Count -and [IO.File]::ReadAllText($Attachments[0]) -ne 'snapshot'){throw 'SMTP attachment changed'}
 $script:sent++
}
try {
 foreach ($configuredPath in @('', ' ', '__USE_GLOBAL__', 'USE_GLOBAL', '{{WorkspaceRootPath}}\Tenants\{{TenantKey}}\MAIL')) {
  $script:cfg.ReportMailSchedule.QueuePath=$configuredPath
  if((Get-SmartM365DeferredMailSettings).Path -ne 'C:\SyntheticWorkspace\Tenants\synthetic-demo\MAIL'){throw 'Default queue path did not resolve under WorkspaceRootPath'}
 }
 $script:cfg.ReportMailSchedule.PSObject.Properties.Remove('QueuePath')
 if((Get-SmartM365DeferredMailSettings).Path -ne 'C:\SyntheticWorkspace\Tenants\synthetic-demo\MAIL'){throw 'Missing QueuePath did not use the default'}
 $script:cfg.ReportMailSchedule | Add-Member -NotePropertyName QueuePath -NotePropertyValue $root
 if((Get-SmartM365DeferredMailSettings).Path -ne $root){throw 'Explicit QueuePath was overwritten'}
 if(& $aclVerifier -Path $root -TrustedSids @([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)){throw 'Unprotected queue was accepted'}
 foreach ($edition in @('SmartM365.Core.psm1','Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
  $tokens=$null;$errors=$null
  $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $core $edition),[ref]$tokens,[ref]$errors)
  if($errors.Count){throw ($errors.Message -join ';')}
  $fn=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'SendEmailHtmlReport'},$true)
  . ([scriptblock]::Create($fn.Extent.Text))
  foreach($mode in @('Graph','SMTP','Both')) {
   $script:failGraph=$mode -eq 'Both'
   $script:clock=[datetimeoffset]::Parse('2026-11-03T06:00:00Z')
   $attachment=Join-Path $root 'snapshot.txt';[IO.File]::WriteAllText($attachment,'snapshot')
   $before=$script:sent
   $params=@{From='sender@example.invalid';To='report@example.invalid';Cc='';Subject='Synthetic';BodyHtml='<p>Report</p>';Attachments=@($attachment);AllowAttachments=$true;SendMailMode=$mode;SmtpServer='synthetic.invalid';MailPurpose='Report'}
   SendEmailHtmlReport @params
   if($script:sent -ne $before -or $global:SmartM365LastMailDelivery -ne 'Queued'){throw 'Report was not deferred'}
   [IO.File]::WriteAllText($attachment,'changed')
   $null=Invoke-SmartM365DeferredMail
   if($script:sent -ne $before){throw 'Worker sent early'}
   $script:clock=[datetimeoffset]::Parse('2026-11-03T06:30:00Z')
   $result=Invoke-SmartM365DeferredMail
   if($result.Sent -ne 1 -or $script:sent -ne ($before+1)){throw ('Worker failed: '+($result|ConvertTo-Json -Compress))}
   $null=Invoke-SmartM365DeferredMail
   if($script:sent -ne ($before+1)){throw 'Worker repeated a delivered message'}
   $params.Attachments=@();$params.DeliveryClass='Operational';$script:clock=[datetimeoffset]::Parse('2026-11-03T06:00:00Z')
   SendEmailHtmlReport @params
   if($script:sent -ne ($before+2)){throw 'Operational notice was deferred'}
   $params.DeliveryClass='Report';$params.MailPurpose='Error'
   SendEmailHtmlReport @params
   if($script:sent -ne ($before+3)){throw 'Error mail was deferred'}
   $params.MailPurpose='Report';$script:cfg.ReportMailSchedule.Enabled=$false
   SendEmailHtmlReport @params
   if($script:sent -ne ($before+4)){throw 'Disabled setting changed behavior'}
   $script:cfg.ReportMailSchedule.Enabled=$true
  }
  $script:failGraph=$false
  $key=if ($edition -eq 'SmartM365.Core.psm1') { 'Test.Core' } else { 'Test.WindowsPowerShell' }
  $cadenceParams=@{From='sender@example.invalid';To='report@example.invalid';BodyHtml='<p>Report</p>';SendMailMode='Graph';MailPurpose='Report';ReportCadenceKey=$key}

  $script:callerConfig=@{ReportMailCadence=@{Mode='SelectedDays';Days=@('Monday')}}
  $before=$script:sent
  # Sunday UTC crosses midnight in Paris one hour before the UTC date changes.
  $script:clock=[datetimeoffset]::Parse('2026-11-01T22:59:00Z')
  SendEmailHtmlReport @cadenceParams
  if ($global:SmartM365LastMailDelivery -ne 'Skipped' -or $script:sent -ne $before) {throw 'Sunday report passed Monday-only cadence'}
  $script:clock=[datetimeoffset]::Parse('2026-11-01T23:00:00Z')
  SendEmailHtmlReport @cadenceParams
  if ($global:SmartM365LastMailDelivery -ne 'Queued' -or $script:sent -ne $before) {throw 'Monday report was not queued at midnight in Paris'}
  SendEmailHtmlReport @cadenceParams
  if ($global:SmartM365LastMailDelivery -ne 'Skipped') {throw 'Duplicate Monday submission was accepted'}
  $script:clock=[datetimeoffset]::Parse('2026-11-02T06:30:00Z')
  $null=Invoke-SmartM365DeferredMail
  if ($script:sent -ne ($before+1)) {throw 'Monday queue did not send exactly one report'}

  $cadenceParams.ReportCadenceKey="$key.expired"
  $script:clock=[datetimeoffset]::Parse('2026-11-02T06:00:00Z')
  SendEmailHtmlReport @cadenceParams
  if ($global:SmartM365LastMailDelivery -ne 'Queued') {throw 'Expiration fixture was not queued'}
  $script:clock=[datetimeoffset]::Parse('2026-11-03T06:30:00Z')
  $expired=Invoke-SmartM365DeferredMail
  if ($expired.Expired -ne 1 -or $expired.Sent -ne 0) {throw 'Expired report was not recorded separately from sent mail'}
  SendEmailHtmlReport @cadenceParams
  if ($script:sent -ne ($before+1)) {throw 'Monday report was sent on Tuesday'}
  $cadenceParams.MailPurpose='Error'
  SendEmailHtmlReport @cadenceParams
  if ($script:sent -ne ($before+2)) {throw 'Report cadence blocked an error alert'}

  $cadenceParams.MailPurpose='Report'
  $script:callerConfig=@{ReportMailCadence=@{Mode='Daily'}}
  $cadenceParams.ReportCadenceKey="$key.scope1"
  SendEmailHtmlReport @cadenceParams
  $cadenceParams.ReportCadenceKey="$key.scope2"
  SendEmailHtmlReport @cadenceParams
  if ($script:sent -ne ($before+4)) {throw 'Distinct scopes blocked one another'}
  $script:clock=$script:clock.AddDays(1)
  SendEmailHtmlReport @cadenceParams
  if ($script:sent -ne ($before+5)) {throw 'Daily cadence did not reset on the next day'}

  $script:callerConfig=[pscustomobject]@{ReportMailCadence=[pscustomobject]@{Mode='SelectedDays';Days=@('TUESDAY','Wednesday')}}
  $cadenceParams.ReportCadenceKey="$key.selected"
  SendEmailHtmlReport @cadenceParams
  if ($script:sent -ne ($before+6)) {throw 'Selected weekdays did not allow a configured Wednesday'}
  $script:callerConfig.ReportMailCadence.Days=@('Monday','monday')
  $cadenceParams.ReportCadenceKey="$key.invalid"
  SendEmailHtmlReport @cadenceParams
  if ($global:SmartM365LastMailDelivery -ne 'Skipped' -or $script:sent -ne ($before+6)) {throw 'Invalid cadence allowed delivery'}
  $script:callerConfig=@{}
 }
 Write-Host 'PASS: Paris report cadence, distinct scopes, deduplication, queue expiry, mocked Graph/SMTP, snapshots and operational notices.'
} finally {
 $resolved=[IO.Path]::GetFullPath($root)
 if((Split-Path -Leaf $resolved) -notmatch '^SmartM365-DeferredIntegration-[a-f0-9]{32}$'){throw 'Unexpected test path'}
 Remove-Item -LiteralPath $resolved -Recurse -Force
}


# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAwaB/S0Wp1K4V8
# lMg4LntucPfea0J95NU5XU72fk2FJaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgIi9quYKSiaK7
# 7FaQM/AWKlaDxLObaDaODFioGchH/J4wDQYJKoZIhvcNAQEBBQAEggGAF0/yKGfr
# JD9vEpNZAHG7LnOBWdrXz3vuyZEiFsqKUVnrruhU85GjoNeS9GH0rG7+DMvmElKG
# OofLKJOGt78SRqauu1/eGOt4QgBrFQ7NTZKknkKIuyDv4ysLbRFNSCSUfNlyNMh8
# 42SAlSdJhM+224WmrIqZPYtA5H0WZqRk7Kzvqx1l8gwHmFhKoRmjaDg0/7MbeVo3
# ed5wEqZlZYIp5V+/Xm8xXjpkTdPOHxQnIA7xR/fxF9NDhBmWRwwzbSZ/uP1jdEar
# 52+NLg0NDkj0vx/F6HwHUqBzh15YBx2ZZ/l0PV4nXyZ8Mqf4VjeS4+Z6MeBOEA6p
# kJZ1TE3z3u5ThfdPZvCenA9gRlp+nUv8d37NdnuP+JAfW99vgMjCWXVgWog63o9Y
# UXKnXU4OkB6d79P/MygIDyvWw5vI8/Pl+t0QKhXAXNxJUYtEtzDw7iICqyaNuu7i
# KkgNbi23zZbFzHYCj+u4hWnQ7hW2UwbESsHhD2DuRo6oHak/jZWJW9EpoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDgwNzA0WjA/BgkqhkiG9w0BCQQxMgQwgn5Iz0d4IQcdr9TNIqtw6N9v
# IC7T9gFWNeoVc7I6KfKqzBUoj+CKIRJpMF4Y/WRmMA0GCSqGSIb3DQEBAQUABIIC
# AFV+n4M/rLqvRVH/5U1rsNxqZlFrZ3Q/fFYqJa0BbKxyPF+itkpjKAkpDSzlpAx4
# jYg8H4FduRqlvrNWizCuu/BDOirdEQLaSSAFbIYolMZ6rPyP30QT9icTn9OgiYbm
# wdL8YPI1IwCW20neSSOyp3cwhCCj73wbalI3HFbK4XE3J0iOMTsBcpem543zfJ+P
# R0GFKmusxb5omWhX9SzOuzKA6JedlioRr5HuHgnma4FPYmq27KFVCGHc+IEHaPMI
# r10SHzc5YILfNsl9sRGB41On0/VRsXgwEEVtK2D5jUY6nagCy/XpHnMEuKJD9rMa
# 90fteeVZ7cZ/x/DjJn+KSaSW82F9kd+KYq8pvs713H4Hhwm43xuR5d/7JHeAkKZ1
# MaFu1vqowd6dHOlk/ohGlY1CV9Ky4mWWLZ01dYhB2aziXkFif0JaWzoMa8WHPFn3
# gICH5z6T91hyV320hK+9uPViOtre+zknISjmWYLgFV6tG8na30s5ezrL+SqxlOGu
# mFrJBJkHKSyZ6mWPN4sWt8gzb7H/7Pfqkyrpvoqe/16bURv6heVofZaQn+EAEnrX
# DdYA+ypNEQBDdiK+ancsSAxmaBwhCgqZTPO50hLjTJkD16xgQ2ZoeJOSb3jwCyV4
# TwP0nhqRpWk7aOHc/hprtFYdqBqQAisyMfrIPZjg3N4x
# SIG # End signature block
