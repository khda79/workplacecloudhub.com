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
function Resolve-SmartM365ConfigValue {param($Value) $Value}
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCQgikMVwM+IYPK
# VHG5DWlnX7dK2N96mVPpDAn/7yMqRKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBMbTlwjGv9N11vSizjmWfS
# /BP00wPa2PpJ3qiswweLVzANBgkqhkiG9w0BAQEFAASCAYA7WxcaPjCoF/WfDEfE
# i+YLmXNf3y6p0sWmtUohpy5FxEkdbMkh1zCGPqiQDlUuWlMDppUnKSjkqWLE742x
# PYvb2lr8lCsmqe0yAeo0T3pJ5uEGsKsnm2EO/1vpiy7QXEKgS/4aHpZv9UlfqUkr
# Y8Bv03UsBkh0V9DIWtb3BhjjTzNbFdS8IopGeprlzaGU4bPqDefGF8poFb0yZZ+b
# oeDMf++Etj9yQdqUeBOaGKT49ZmgrjL4t1F50cC9ml/7/s4SpOZA2n0170ICIkr1
# LgAgzbae6ItaAt97NealtpAjOZ1gBtO5rmAkKT6StvhS2NQ6qH107kW8qpsrGuYW
# TMOVnRjWk7esPHuG3gQOWAOPrIsSFCWWLZr5etcfv83TRpI5TXEZ5jNFER/syOtj
# 2cfRdc88Q4fdELmSqiw0cLy36FV+RmBTI/oKJpoTIJXPxShddpkflwZoMymVdYKJ
# ZnQBv5vY8cIrQkPczWMCVzpM+w8PLLH5KB7eHInsm7dvvFc=
# SIG # End signature block
