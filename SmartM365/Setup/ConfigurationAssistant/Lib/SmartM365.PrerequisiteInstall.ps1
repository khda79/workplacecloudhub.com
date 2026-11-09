# SmartInventory prerequisite installation adapter, version 0.5.0.
function New-SiPrerequisitePlan {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][ValidateSet('CurrentUser','AllUsers')][string]$Scope,[bool]$TrustRepository=$false,[bool]$InstallDirectoryFeature=$false)
    $path=Join-Path $Root 'Setup/Install-SmartM365-SmartInventoryPrerequisites.ps1'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Prerequisite installer is missing.' }
    [pscustomobject]@{Path=$path;Scope=$Scope;TrustRepository=$TrustRepository;InstallDirectoryFeature=$InstallDirectoryFeature;SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}
}
function Get-SiDirectoryFeatureInstallCommand {
    @'
$ErrorActionPreference='Stop'
$env:PSModulePath=@((Join-Path $PSHOME 'Modules'),(Join-Path $env:ProgramFiles 'WindowsPowerShell/Modules')) -join [IO.Path]::PathSeparator
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    $os=Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    if ($null -eq $os -or $os.ProductType -notin @(1,2,3)) {throw 'Windows edition could not be determined.'}
    if ($os.ProductType -eq 1) {
        $result=Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ErrorAction Stop
        if ($null -eq $result) {throw 'Active Directory RSAT installation returned no result.'}
    } else {
        Import-Module ServerManager -ErrorAction Stop
        $result=Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeManagementTools -ErrorAction Stop
        if ($null -eq $result -or -not $result.Success) {throw 'Active Directory RSAT installation failed.'}
    }
}
'@
}
function Start-SiPrerequisiteInstallation {
    param([Parameter(Mandatory)]$Plan,[string]$Executable=(Join-Path $PSHOME 'pwsh.exe'))
    if ($Plan.Scope -notin @('CurrentUser','AllUsers') -or $Plan.TrustRepository -isnot [bool] -or $Plan.InstallDirectoryFeature -isnot [bool] -or ($Plan.InstallDirectoryFeature -and $Plan.Scope -ne 'AllUsers')) { throw 'Invalid installation plan.' }
    $bytes=[IO.File]::ReadAllBytes($Plan.Path)
    $sha=[Security.Cryptography.SHA256]::Create()
    try { $hash=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','') } finally { $sha.Dispose() }
    if ($hash -ne $Plan.SHA256) { throw 'Installer changed after review; prepare a new plan.' }
    if (-not [IO.Path]::IsPathFullyQualified($Executable) -or -not (Test-Path -LiteralPath $Executable -PathType Leaf)) { throw 'PowerShell executable not found.' }
    $dir=Join-Path ([IO.Path]::GetTempPath()) ('smartinventory-prereq-'+[guid]::NewGuid().ToString('N'))
    $copy=Join-Path $dir 'Install-SmartM365-SmartInventoryPrerequisites.ps1'
    $proc=$null; $copyLock=$null; $started=$false
    try {
        $null=New-Item -ItemType Directory -Path $dir -ErrorAction Stop
        [IO.File]::WriteAllBytes($copy,$bytes)
        if ((Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash -ne $Plan.SHA256) { throw 'Installer copy verification failed.' }
        $copyLock=[IO.File]::Open($copy,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        if ((Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash -ne $Plan.SHA256) { throw 'Installer copy changed before protection was acquired.' }
        $info=[Diagnostics.ProcessStartInfo]::new()
        $info.FileName=$Executable; $info.UseShellExecute=$false; $info.CreateNoWindow=$true
        $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
        $modulePath=if ($Plan.Scope -eq 'AllUsers') {@((Join-Path $PSHOME 'Modules'),(Join-Path $env:ProgramFiles 'PowerShell/Modules'),(Join-Path $env:ProgramFiles 'WindowsPowerShell/Modules'),(Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules')) -join [IO.Path]::PathSeparator} else {$env:PSModulePath}
        $info.Environment['PSModulePath']=$modulePath
        if ($Plan.Scope -eq 'AllUsers') {
            $info.Environment['SI_MACHINE_MODULE_PATH']=$modulePath;$info.Environment['SI_PREREQ_FILE']=$copy;$info.Environment['SI_PREREQ_TRUST']=if($Plan.TrustRepository){'1'}else{'0'}
            $info.Environment['SI_INSTALL_AD']=if($Plan.InstallDirectoryFeature){'1'}else{'0'}
            $info.Environment['SI_AD_COMMAND']=Get-SiDirectoryFeatureInstallCommand
            $command='$ErrorActionPreference="Stop"; $env:PSModulePath=$env:SI_MACHINE_MODULE_PATH; if ($env:SI_INSTALL_AD -eq "1") { $legacy=Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"; & $legacy -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $env:SI_AD_COMMAND; if ($LASTEXITCODE -ne 0) { throw "Active Directory RSAT installation failed." } }; if ($env:SI_PREREQ_TRUST -eq "1") { & $env:SI_PREREQ_FILE -Scope AllUsers -Force:$false -AllowClobber:$false -SkipImportValidation -TrustRepository } else { & $env:SI_PREREQ_FILE -Scope AllUsers -Force:$false -AllowClobber:$false -SkipImportValidation }; if (-not $?) { exit 1 }'
            foreach ($arg in @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-Command',$command)) {$info.ArgumentList.Add($arg)}
        } else {
            foreach ($arg in @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$copy,'-Scope',$Plan.Scope,'-Force:$false','-AllowClobber:$false','-SkipImportValidation')) { $info.ArgumentList.Add($arg) }
            if ($Plan.TrustRepository) { $info.ArgumentList.Add('-TrustRepository') }
        }
        $proc=[Diagnostics.Process]::new(); $proc.StartInfo=$info
        if (-not $proc.Start()) { throw 'Installer process could not start.' }
        $started=$true
        [pscustomobject]@{Process=$proc;Stdout=$proc.StandardOutput.ReadToEndAsync();Stderr=$proc.StandardError.ReadToEndAsync();Directory=$dir;Copy=$copy;Executable=$Executable;SHA256=$Plan.SHA256;CopyLock=$copyLock;ModulePath=$modulePath;Scope=$Plan.Scope}
    } catch {
        $stopUncertain=$false
        if ($null -ne $proc -and $started) { try { if (-not $proc.HasExited) { $proc.Kill($true); if (-not $proc.WaitForExit(3000)) {$stopUncertain=$true} } } catch {$stopUncertain=$true} }
        if ($null -ne $proc) { $proc.Dispose() }
        if ($null -ne $copyLock) { $copyLock.Dispose() }
        if ($stopUncertain) {throw 'Installer start failed and its process may still be running; inspect the machine before retrying.'}
        if (Test-Path -LiteralPath $copy -PathType Leaf) { [IO.File]::Delete($copy) }
        if (Test-Path -LiteralPath $dir -PathType Container) { [IO.Directory]::Delete($dir,$false) }
        throw
    }
}
function Invoke-SiIsolatedModuleImport {
    param([Parameter(Mandatory)][string]$Name,[Parameter(Mandatory)][string]$Executable,[Parameter(Mandatory)][string]$ModulePath,[int]$TimeoutMilliseconds=600000)
    if ($Name -cnotmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,127}\z') {throw 'Invalid module name in audit output.'}
    $info=[Diagnostics.ProcessStartInfo]::new();$info.FileName=$Executable;$info.UseShellExecute=$false;$info.CreateNoWindow=$true
    $info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
    $info.Environment['SI_MACHINE_MODULE_PATH']=$ModulePath;$info.Environment['SI_IMPORT_MODULE']=$Name
    $command='$ErrorActionPreference="Stop"; $env:PSModulePath=$env:SI_MACHINE_MODULE_PATH; try { Import-Module -Name $env:SI_IMPORT_MODULE -ErrorAction Stop; exit 0 } catch { [Console]::Error.WriteLine($_.Exception.ToString()); exit 1 }'
    foreach ($arg in @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-Command',$command)) {$info.ArgumentList.Add($arg)}
    $process=[Diagnostics.Process]::Start($info)
    try {
        $out=$process.StandardOutput.ReadToEndAsync();$err=$process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutMilliseconds)) {try {$process.Kill($true);$stopped=$process.WaitForExit(3000)} catch {$stopped=$false};return [pscustomobject]@{Name=$Name;Success=$false;Detail=$(if ($stopped) {'Import validation timed out.'} else {'Import validation timed out; process termination is uncertain.'})}}
        $streams=Read-SiBoundedProcessStreams -Stdout $out -Stderr $err
        $detail=($streams.Stdout+[Environment]::NewLine+$streams.Stderr).Trim()
        [pscustomobject]@{Name=$Name;Success=($process.ExitCode -eq 0);Detail=$detail}
    } finally {$process.Dispose()}
}
function Complete-SiPrerequisiteInstallation {
    param([Parameter(Mandatory)]$Run,[scriptblock]$ImportCheck,[ValidateRange(1,3600)][int]$AuditTimeoutSeconds=300)
    if (-not $Run.Process.HasExited) { throw 'Installation is still running.' }
    try {
        $installExit=$Run.Process.ExitCode
        $streams=Read-SiBoundedProcessStreams -Stdout $Run.Stdout -Stderr $Run.Stderr
        $installOutput=$streams.Stdout; $installError=$streams.Stderr
        if ((Get-FileHash -LiteralPath $Run.Copy -Algorithm SHA256).Hash -ne $Run.SHA256) { throw 'Isolated installer changed during execution; result is uncertain.' }
        $auditInfo=[Diagnostics.ProcessStartInfo]::new();$auditInfo.FileName=$Run.Executable;$auditInfo.UseShellExecute=$false;$auditInfo.CreateNoWindow=$true
        $auditInfo.RedirectStandardOutput=$true;$auditInfo.RedirectStandardError=$true;$auditInfo.Environment['PSModulePath']=$Run.ModulePath
        if ($Run.Scope -eq 'AllUsers') {
            $auditInfo.Environment['SI_MACHINE_MODULE_PATH']=$Run.ModulePath;$auditInfo.Environment['SI_PREREQ_FILE']=$Run.Copy
            $auditCommand='$env:PSModulePath=$env:SI_MACHINE_MODULE_PATH; & $env:SI_PREREQ_FILE -AuditOnly'
            foreach ($arg in @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-Command',$auditCommand)) {$auditInfo.ArgumentList.Add($arg)}
        } else {foreach ($arg in @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$Run.Copy,'-AuditOnly')) {$auditInfo.ArgumentList.Add($arg)}}
        $audit=[Diagnostics.Process]::Start($auditInfo)
        try {
            $out=$audit.StandardOutput.ReadToEndAsync();$err=$audit.StandardError.ReadToEndAsync()
            if (-not $audit.WaitForExit($AuditTimeoutSeconds * 1000)) {
                try {$audit.Kill($true);$stopped=$audit.WaitForExit(3000)} catch {$stopped=$false}
                if (-not $stopped) {throw 'Prerequisite audit timed out and process termination is uncertain. Inspect the machine before retrying.'}
                throw 'Prerequisite audit timed out. Installation may be partial; inspect the machine before retrying.'
            }
            $streams=Read-SiBoundedProcessStreams -Stdout $out -Stderr $err
            $auditExit=$audit.ExitCode;$auditOutput=@($streams.Stdout,$streams.Stderr)
        } finally {$audit.Dispose()}
        $state=if ($installExit -eq 0 -and $auditExit -eq 0) { 'InstalledAuditPassed' } elseif ($installExit -ne 0) { 'InstallationFailedOrPartial' } else { 'AuditNotPassed' }
        $checks=[Collections.Generic.List[string]]::new()
        if ($state -eq 'InstalledAuditPassed') {
            $names=@($(foreach ($line in (($auditOutput -join [Environment]::NewLine) -split '\r?\n')) {if ($line -match '^AVAILABLE ([A-Za-z0-9][A-Za-z0-9._-]{0,127}) ') {$Matches[1]}}) | Select-Object -Unique)
            if (-not $names.Count) {$state='AuditNotPassed';$checks.Add('No module results were returned by the prerequisite audit.')}
            else {
                if ($null -eq $ImportCheck) {$ImportCheck={param($name,$run) Invoke-SiIsolatedModuleImport -Name $name -Executable $run.Executable -ModulePath $run.ModulePath}}
                foreach ($name in $names) {
                    $check=& $ImportCheck $name $Run
                    if ($null -eq $check -or -not $check.Success) {$state='ImportValidationFailed';$checks.Add('IMPORT FAILED '+$name+': '+[string]$check.Detail);break}
                    $checks.Add('IMPORT OK '+$name)
                }
            }
        }
        [pscustomobject]@{State=$state;InstallExit=$installExit;AuditExit=$auditExit;Detail=((($checks -join [Environment]::NewLine),$installError,$installOutput,($auditOutput -join [Environment]::NewLine)) -join [Environment]::NewLine)}
    } finally {
        $Run.Process.Dispose()
        $Run.CopyLock.Dispose()
        if (Test-Path -LiteralPath $Run.Copy -PathType Leaf) { [IO.File]::Delete($Run.Copy) }
        if (Test-Path -LiteralPath $Run.Directory -PathType Container) { [IO.Directory]::Delete($Run.Directory,$false) }
    }
}
function Stop-SiPrerequisiteInstallation {
    param([Parameter(Mandatory)]$Run)
    if (-not $Run.Process.HasExited) { $Run.Process.Kill($true) }
}
. (Join-Path $PSScriptRoot 'SmartM365.BoundedStreams.ps1')

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDrpNDiCjZcWzmL
# 5ERVcwlykflJKxqCdQu+Xv1l7tZjQKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAtBx/oPdfUa5UUX0OK6jmdm6hWFhOJKuiml0QBXTbV3MA0GCSqG
# SIb3DQEBAQUABIIBgHA+CwCRG2q7spVMoaRt9cf6hDucGrGS+nJzbo73M6A9uKGD
# 8nzbhjwdb4fMmLV8vdXkFS0CbbsXuXe/1Q5fDBG/M6iY2Yw72OyCtN0murQXTKjl
# skjy9jQnI0nmCWAh7Mgrk68q11GEe/HuSFn4FBcxf6jQCrS/ySWLNYjGrGOe6fhI
# o67vCkkoYkHlEiPpUKcXJ8YmLJkiwurdm3YzgCa064tPRxyfA7l+7RtGIZHgh8h8
# Qj3dVzfYwfY6jFmkLcJRmbxAmS1gWCsZxiSt9DibM5Sbai6sIQKr03Ndp9T975kU
# QXx+OSG/InYbE/DTDWAxhwtk6u9iUkL9aadD3xZr6qJDbmc20skkqvjfuiC2V5aq
# P98O4ZL+djcWhyYE35fhGI04M8WFLpnQ7Ui7Gcp8oGv6HC7Y38ZqZmhjYSYAZIT2
# +80DaMLmdS1JvEOOBHaSiGFuPPZaymIIpqsBqw59aN9tegPUovwVgFM6yhpjrEZh
# WFSg0RjTy9b1OcKHwqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjJaMC8GCSqGSIb3DQEJBDEiBCD0HrY2PU2FtRkx4DmZYw8hYPxGxPUB8ZyoznHH
# kE17XjANBgkqhkiG9w0BAQEFAASCAgAg8rz9pnU6+SRvSmgQhk0vC8nTaZTRSgLG
# ++hDTwJf81GsEaOnRuB08861KuOcC4Vx+365m/DBlizpCM9wmhLJH9zoyW7Miphi
# NyyeWJfdNsPjF4aQa14k9+jtkzPkC6uuMLbOeP4bxrCzvVlqdVK0558zZv2CpK+o
# bpDdPJsF55krmxhqlmod2CNPRnSddFIxHrTBWxlgCQ0Rz06/QVHccEGieew62E+f
# Sf9btiA3cySc+LMHpKtiGBlsNFx/KU/93oyv2vuQJ8V7broALWPjajSf3m+aOtxw
# TWUCI71Bkp6W30tYEt0R3Tq/zF7OduYEKBhtPkdO4cU3265BiyZQXwttHz2KexQr
# bWuTzEx1q2LByZnvCkbHtRGaBahWjx298i7tNd/iT75iPRlK06nrhdljpXZHUGrS
# 1nasER3M3HjQlAeft/4aZo/TzeXj/8i1d24bmRnWsQaYGIg+iIfQRe4/rF9iHrRq
# 0J6cBogYdx0Zmr1qN4BqwASpcJUoFIYgQamX03EZRCgio8S8ZYupz1rE70PaqBTR
# EZ9NfTHH3NrKzu2lIkcQVNGWM7hYPQdL0qBYF3+ZYeScG2HymlRcV0nnDYvSIj+i
# 0tDnhlfE2yvJH6t51CaCcFWis3xoL6YEzfIDFQ5bizVLTMO2F9DECbx4lyYM8sdC
# QJur6ec6zQ==
# SIG # End signature block
