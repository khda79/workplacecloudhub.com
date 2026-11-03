# Extracts the actual mail entrypoints. Authentication and transport are mocked.
$ErrorActionPreference='Stop'
$core=Join-Path $PSScriptRoot '../Modules/SmartM365.Core'
. (Join-Path $core 'SmartM365-DeferredMail.ps1')
$aclVerifier=${function:Test-SmartM365DeferredMailDirectory}
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DeferredIntegration-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $root
$script:cfg=[pscustomobject]@{ProfileKey='demo';ReportMailSchedule=[pscustomobject]@{Enabled=$true;NotBefore='07:30';TimeZone='Romance Standard Time';QueuePath=$root;TrustedSids=@('synthetic')}}
$script:sent=0
function Get-SmartM365MailUtcNow {$script:clock}
function Get-SmartM365EffectiveModuleGlobalConfig {$script:cfg}
function Resolve-SmartM365ConfigValue {param($Value) $Value.Replace('{{WorkspaceRootPath}}','C:\SyntheticWorkspace').Replace('{{TenantKey}}','synthetic-demo')}
function Test-SmartM365DeferredMailDirectory {param($Path,$TrustedSids) $Path -eq $root}
function Get-PSCallStack {[pscustomobject]@{ScriptName='C:\synthetic\SmartInventory\Collector.ps1'}}
function Get-ModuleLocalConfig {[pscustomobject]@{}}
function Get-SmartM365CallerLocalConfig {[pscustomobject]@{}}
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
 }
 Write-Host 'PASS: actual mail entrypoints -> queue -> worker -> mocked Graph/SMTP, snapshots, operational notices.'
} finally {
 $resolved=[IO.Path]::GetFullPath($root)
 if((Split-Path -Leaf $resolved) -notmatch '^SmartM365-DeferredIntegration-[a-f0-9]{32}$'){throw 'Unexpected test path'}
 Remove-Item -LiteralPath $resolved -Recurse -Force
}


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBpX/0lReiQIOWk
# gWJKXHxYiItgBDg4iU8TtQSmoOJMgqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCOWluq3BcnAiSHQ8kUu5A3
# dM0lVF++U1L2pZCYWRqPCjANBgkqhkiG9w0BAQEFAASCAYCa8NcBM/qzR9uAYM36
# ySvKLNW5szmJK+z91QWVgZ+WQtIlhVYstse0mq08F4OZeVAUHRO1UpZHQxqEV1bq
# 6Hu/zArSA4AZyeti9kL3OanVO1T1z/hTXEub8b8QLlP13wHP8s0i5I/nNrBuQPcn
# CJa0eLj/HGUaoZUEpvLsLxyY3GaNvjE+XtMLBGy1oPq+j2m4lqW+9urkAcH8QYKe
# piDaPPl1qJgSGZjoaG+WvMCuhpiCD7lRUeZEnqKaTlcDhZW/8LS7laXvV71P28nR
# Gl7mZLfdEBkd1JDVqu8qiLcqoleg52ZIVsY9KRxJ+js9jo2yGKrUajzM/CmAhK7x
# 4f52q5BiNtAd702cJa9PusoQ9I2cm/4R2y/d8b+JQvIT4mASYuNoj8q74LGyIPLr
# YM9I+u4nLNc2ZUAfM/HMHe7zQAK2cujKkCJaPDjTVpWtWCZJV80NUlOVgHxwg7U/
# Wk32ziDGlZu1bE6dUxn4OiErHe8fXhNOqP5ulQM784iSnlc=
# SIG # End signature block
