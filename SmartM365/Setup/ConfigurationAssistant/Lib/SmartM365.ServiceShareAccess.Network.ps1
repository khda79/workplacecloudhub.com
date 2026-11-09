# Service-account access and least-privilege grant orchestration.
function ConvertTo-SiServiceShareTarget {
    param([Parameter(Mandatory)][string]$Path)
    if ($Path -notmatch '^\\\\(?<server>[a-zA-Z0-9][a-zA-Z0-9.-]*)\\(?<share>[^\\/:*?"<>|]+)(?<tail>(?:\\[^\\/:*?"<>|]+)*)\\?$') {throw 'Enter a valid UNC share path.'}
    $server=$Matches.server;$share=$Matches.share;$tail=$Matches.tail
    if (@($tail -split '\\' | Where-Object {$_ -in @('.','..')}).Count) {throw 'Relative path segments are not supported.'}
    [pscustomobject]@{Server=$server;Share=$share;Relative=$tail.TrimStart('\');Path=$Path.TrimEnd('\')}
}
function Get-SiServiceShareGrantReview {
    param([Parameter(Mandatory)]$SourceTarget,[Parameter(Mandatory)]$DataTarget,[Parameter(Mandatory)][string]$Account)
    $source=ConvertTo-SiServiceShareTarget -Path ('\\'+$SourceTarget.Server+'\'+$SourceTarget.Share)
    $data=ConvertTo-SiServiceShareTarget -Path ('\\'+$DataTarget.Server+'\'+$DataTarget.Share)
    $paths=@($source.Path); if ($data.Path -ine $source.Path) {$paths+= $data.Path}
    $message="Apply permissions for $Account to these entire share roots?`n`n"+($paths -join "`n")+"`n`nThis grants SMB Change and NTFS Modify, disables inherited NTFS permissions on each root, removes the local Users entry and resets descendant ACLs, including protected ACLs. Prior NTFS ACLs are backed up on the file server. Existing access may change. If an operation fails, changes already made are not rolled back automatically."
    [pscustomobject]@{Paths=$paths;Message=$message}
}
function Test-SiServiceShareCredential {
    param([Parameter(Mandatory)]$Target,[Parameter(Mandatory)][string]$Sid,[Parameter(Mandatory)][pscredential]$Credential,[scriptblock]$Launch,[int]$TimeoutSeconds=45,[bool]$OfflineFixtureProbe=$false)
    if ($Sid -notmatch '^S-1-(?:\d+-)+\d+$') {throw 'Invalid service identity.'}
    $markerName='.si-access-'+[guid]::NewGuid().ToString('N')+'.tmp'
    $markerPath=Join-Path $Target.Path $markerName
    if ($null -eq $Launch) {
        $probePath=Join-Path $PSScriptRoot 'SmartM365.ServiceShareNetworkProbe.ps1'
        $Launch={param($target,$expectedSid,$credential,$timeout,$markerName,$offlineFixture)
            $probe=$probePath
            $probeHash=(Get-FileHash -LiteralPath $probe -Algorithm SHA256 -ErrorAction Stop).Hash
            $lock=[IO.File]::Open($probe,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
            try {
                if ((Get-FileHash -LiteralPath $probe -Algorithm SHA256).Hash -cne $probeHash) {throw 'Service-account check changed before launch.'}
                $workDirectory=[Environment]::GetFolderPath([Environment+SpecialFolder]::System)
                $info=[Diagnostics.ProcessStartInfo]::new();$info.FileName=Join-Path $PSHOME 'pwsh.exe';$info.WorkingDirectory=$workDirectory;$info.UseShellExecute=$false;$info.CreateNoWindow=$true;$info.RedirectStandardInput=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
                foreach ($value in @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$probe,'-MarkerPath',(Join-Path $target.Path $markerName),'-ExpectedSid',$expectedSid)) {$info.ArgumentList.Add([string]$value)}
                if ($offlineFixture) {$info.ArgumentList.Add('-OfflineFixture')}
                $network=$credential.GetNetworkCredential()
                $payload=@{User=[string]$network.UserName;Domain=[string]$network.Domain;EncryptedPassword=(ConvertFrom-SecureString -SecureString $credential.Password)} | ConvertTo-Json -Compress
                $process=[Diagnostics.Process]::new();$process.StartInfo=$info;$started=$false
                try {
                    if (-not $process.Start()) {throw 'Service-account check could not be started.'}
                    $started=$true
                    $output=$process.StandardOutput.ReadToEndAsync();$errors=$process.StandardError.ReadToEndAsync()
                    $process.StandardInput.WriteLine($payload);$process.StandardInput.Close();$payload=$null
                    if (-not $process.WaitForExit($timeout*1000)) {try {$process.Kill($true);$stopped=$process.WaitForExit(3000)} catch {$stopped=$false};return [pscustomobject]@{Code=14;ProcessStopped=$stopped}}
                    $streams=Read-SiBoundedProcessStreams -Stdout $output -Stderr $errors; $text=$streams.Stdout
                    try {$result=$text | ConvertFrom-Json -ErrorAction Stop} catch {throw 'Service-account check returned an invalid response.'}
                    if ($null -eq $result -or $null -eq $result.Code) {throw 'Service-account check returned no result.'}
                    return $result
                } finally {$payload=$null;if ($started -and -not $process.HasExited) {try {$process.Kill($true);$null=$process.WaitForExit(3000)} catch {}};$process.Dispose()}
            } finally {$lock.Dispose()}
        }.GetNewClosure()
    }
    $outcome=& $Launch $Target $Sid $Credential $TimeoutSeconds $markerName $OfflineFixtureProbe
    $code=if ($outcome -is [int]) {$outcome} else {[int]$outcome.Code}
    switch ($code) {
        0 {[pscustomobject]@{State='Verified';Detail='Read, write and cleanup succeeded under the selected account.'}}
        10 {[pscustomobject]@{State='Denied';Detail='The selected account was denied access.'}}
        13 {[pscustomobject]@{State='Uncertain';Detail=('Cleanup could not be confirmed. Inspect '+$markerPath+' before retrying.')}}
        14 {[pscustomobject]@{State='Uncertain';Detail=('The test timed out; process termination and cleanup may be uncertain. Inspect '+$markerPath+' before retrying.')}}
        20 {[pscustomobject]@{State='Uncertain';Detail=('Network sign-in failed (Windows error '+[string]$outcome.NativeError+'): '+[string]$outcome.Detail)}}
        default {[pscustomobject]@{State='Uncertain';Detail='Access could not be determined. Check the account credentials, logon rights, connectivity and share before retrying.'}}
    }
}
function Grant-SiServiceShareAccess {
    param([Parameter(Mandatory)]$Target,[Parameter(Mandatory)][string]$Sid,[scriptblock]$Apply)
    if ($Sid -notmatch '^S-1-(?:\d+-)+\d+$') {throw 'Invalid service identity.'}
    if ($null -eq $Apply) {
        $Apply={param($target,$serviceSid)
            $operation={param($shareName,$relative,$sidText)
                $ErrorActionPreference='Stop'
                $share=Get-SmbShare -Name $shareName -ErrorAction Stop
                if (@($share).Count -ne 1 -or -not $share.Path -or -not (Test-Path -LiteralPath $share.Path -PathType Container)) {throw 'The selected Windows file share could not be resolved.'}
                $root=[IO.Path]::GetFullPath([string]$share.Path)
                if ($root.Length -gt [IO.Path]::GetPathRoot($root).Length) {$root=$root.TrimEnd('\')}
                $rootItem=Get-Item -LiteralPath $root -Force -ErrorAction Stop
                if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {throw 'A reparse-point share root requires manual review.'}
                $folder=if ($relative) {[IO.Path]::GetFullPath((Join-Path $root $relative))} else {$root}
                if ($folder -ine $root -and -not $folder.StartsWith($root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) {throw 'Share path escaped its root.'}
                if ($folder.TrimEnd('\').Length -le [IO.Path]::GetPathRoot($folder).Length) {throw 'A drive root cannot be reconfigured.'}
                $cursor=$root
                foreach ($part in @($relative -split '\\' | Where-Object {$_})) {$cursor=Join-Path $cursor $part;$item=Get-Item -LiteralPath $cursor -Force -ErrorAction Stop;if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {throw 'Reparse points require manual review.'}}
                $sid=[Security.Principal.SecurityIdentifier]::new($sidText);$account=$sid.Translate([Security.Principal.NTAccount]).Value
                $entries=@(Get-SmbShareAccess -Name $shareName -ErrorAction Stop)
                $direct=@($entries | Where-Object {try {([Security.Principal.NTAccount]::new($_.AccountName)).Translate([Security.Principal.SecurityIdentifier]).Value -eq $sidText} catch {$false}})
                if (@($direct | Where-Object AccessControlType -eq 'Deny').Count) {throw 'A direct SMB deny requires manual review.'}
                $acl=Get-Acl -LiteralPath $folder -ErrorAction Stop
                $rules=@($acl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier]) | Where-Object {$_.IdentityReference.Value -eq $sidText})
                if (@($rules | Where-Object AccessControlType -eq 'Deny').Count) {throw 'A direct NTFS deny requires manual review.'}
                $smbReady=@($direct | Where-Object {$_.AccessControlType -eq 'Allow' -and $_.AccessRight -in @('Change','Full')}).Count -gt 0
                $ntfsReady=@($rules | Where-Object {-not $_.IsInherited -and $_.AccessControlType -eq 'Allow' -and ($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::Modify) -eq [Security.AccessControl.FileSystemRights]::Modify -and ($_.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ContainerInherit) -and ($_.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ObjectInherit) -and $_.PropagationFlags -eq [Security.AccessControl.PropagationFlags]::None}).Count -gt 0
                $usersSid=[Security.Principal.SecurityIdentifier]::new('S-1-5-32-545')
                $usersRules=@($acl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier]) | Where-Object {$_.IdentityReference.Value -eq $usersSid.Value})
                $changeAcl=(-not $acl.AreAccessRulesProtected -or $usersRules.Count -gt 0 -or -not $ntfsReady)
                $children=@(Get-ChildItem -LiteralPath $folder -Recurse -Force -ErrorAction Stop)
                if (@($children | Where-Object {($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0}).Count) {throw 'A descendant reparse point requires manual review before permissions are changed.'}
                $needsReset=$false
                foreach($child in $children) {$childAcl=Get-Acl -LiteralPath $child.FullName -ErrorAction Stop;if($childAcl.AreAccessRulesProtected -or @($childAcl.GetAccessRules($true,$false,[Security.Principal.SecurityIdentifier])).Count){$needsReset=$true;break}}
                $backupFile=''
                if ($changeAcl -or $needsReset) {
                    $backupDir=Join-Path ([Environment]::GetFolderPath('CommonApplicationData')) ('SmartM365-AclBackup-'+[guid]::NewGuid().ToString('N'))
                    $null=New-Item -ItemType Directory -Path $backupDir -ErrorAction Stop
                    $secure=[Security.AccessControl.DirectorySecurity]::new();$secure.SetAccessRuleProtection($true,$false)
                    $inherit=[Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [Security.AccessControl.InheritanceFlags]::ObjectInherit
                    foreach($trusted in @('S-1-5-18','S-1-5-32-544',[Security.Principal.WindowsIdentity]::GetCurrent().User.Value) | Select-Object -Unique) {$trustedSid=[Security.Principal.SecurityIdentifier]::new($trusted);$secure.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($trustedSid,[Security.AccessControl.FileSystemRights]::FullControl,$inherit,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Allow))}
                    Set-Acl -LiteralPath $backupDir -AclObject $secure -ErrorAction Stop
                    $backupFile=Join-Path $backupDir 'permissions.acl'
                    $icacls=Join-Path ([Environment]::GetFolderPath('System')) 'icacls.exe'
                    & $icacls $folder /save $backupFile /T /L /Q | Out-Null
                    if($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $backupFile -PathType Leaf)) {throw 'Could not save the existing ACLs; no share permission was changed.'}
                }
                if (-not $smbReady) {Grant-SmbShareAccess -Name $shareName -AccountName $account -AccessRight Change -Force -ErrorAction Stop | Out-Null}
                $verifiedSmb=@(Get-SmbShareAccess -Name $shareName -ErrorAction Stop | Where-Object {try {([Security.Principal.NTAccount]::new($_.AccountName)).Translate([Security.Principal.SecurityIdentifier]).Value -eq $sidText -and $_.AccessControlType -eq 'Allow' -and $_.AccessRight -in @('Change','Full')} catch {$false}})
                if (-not $verifiedSmb.Count) {throw 'The explicit SMB permission could not be verified.'}
                if ($changeAcl) {
                    if (-not $acl.AreAccessRulesProtected) {
                        & $icacls $folder /inheritance:d /L /Q | Out-Null
                        if($LASTEXITCODE -ne 0) {throw "Could not protect inherited permissions; inspect $backupFile."}
                        $acl=Get-Acl -LiteralPath $folder -ErrorAction Stop
                    }
                    & $icacls $folder /remove ('*'+$usersSid.Value) /L /Q | Out-Null
                    if($LASTEXITCODE -ne 0) {throw "Could not remove Users from the folder ACL; inspect $backupFile."}
                    if (-not $ntfsReady) {
                        & $icacls $folder /grant ('*'+$sidText+':(OI)(CI)M') /L /Q | Out-Null
                        if($LASTEXITCODE -ne 0) {throw "Could not grant the service account NTFS Modify; inspect $backupFile."}
                    }
                }
                if ($changeAcl -or $needsReset) {
                    foreach($child in @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop)) {
                        & $icacls $child.FullName /reset /T /L /Q | Out-Null
                        if($LASTEXITCODE -ne 0) {throw "ACL propagation failed. Review the folder and ACL backup: $backupFile"}
                    }
                }
                $verified=Get-Acl -LiteralPath $folder -ErrorAction Stop
                $verifiedRules=@($verified.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier]))
                if (-not $verified.AreAccessRulesProtected -or @($verifiedRules | Where-Object {$_.IdentityReference.Value -eq $usersSid.Value}).Count -or -not @($verifiedRules | Where-Object {$_.IdentityReference.Value -eq $sidText -and -not $_.IsInherited -and $_.AccessControlType -eq 'Allow' -and ($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::Modify) -eq [Security.AccessControl.FileSystemRights]::Modify -and ($_.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ContainerInherit) -and ($_.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ObjectInherit) -and $_.PropagationFlags -eq [Security.AccessControl.PropagationFlags]::None}).Count) {throw 'NTFS permissions could not be verified; inspect the share before retrying.'}
                foreach($child in @(Get-ChildItem -LiteralPath $folder -Recurse -Force -ErrorAction Stop)) {
                    if (($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {throw "A reparse point appeared during ACL propagation; inspect $backupFile."}
                    $childAcl=Get-Acl -LiteralPath $child.FullName -ErrorAction Stop
                    if($childAcl.AreAccessRulesProtected -or @($childAcl.GetAccessRules($true,$false,[Security.Principal.SecurityIdentifier])).Count) {throw "A descendant ACL remains explicit or protected; inspect $backupFile."}
                }
                [pscustomobject]@{ShareGranted=(-not $smbReady);FolderGranted=($changeAcl -or $needsReset);Backup=$backupFile}
            }
            $localNames=@($env:COMPUTERNAME,([Net.Dns]::GetHostEntry('localhost').HostName))
            if ($target.Server -in $localNames) {& $operation $target.Share $target.Relative $serviceSid}
            else {Invoke-Command -ComputerName $target.Server -ScriptBlock $operation -ArgumentList $target.Share,$target.Relative,$serviceSid -ErrorAction Stop}
        }
    }
    & $Apply $Target $Sid
}
. (Join-Path $PSScriptRoot 'SmartM365.BoundedStreams.ps1')

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDI0K6MsxvSlTYq
# tgZT5inxTBtnOn0pBMtofM6bSM8UP6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB4aSkOdOLlCj3rSzKjGrIYxQhtxSq3HsBZOCXULFITEMA0GCSqG
# SIb3DQEBAQUABIIBgENrckAwKFKWa/1iM3zIIurLhM2PG9+tmWG9cTSzK/IbnDwg
# q4M0mHFldwbzRCZI0U9j4x0J/CZNzORWjKc3RynJl9yDHyF7+zyT1U9pOVZzI+bZ
# Os7sPvBkLJzqtEKJQWw0c3/WTuGTI4ksCjtkOFklqLsBvhtwB+NrW7124DjJRCpW
# SQT8H24bgqQmv2IDbyE0iUCSLPBpShf38hmUrpkbbq/3ej3i+UqBkRUsiB53QcyK
# EzXBu0QWZ8x5MBF1DX85QZlpUpznRzEg8XmbAw8ftp7ofVBgXmM/PQDfVej+dpZ2
# tYlsHdt1ePEFyP2GURoKGNkB0tfxOuxjDGV2DPPj0sczarOnqK3fdFS8NL7jnw5O
# vlEnIt8+blxbZoWkoyk5CGA9lU9vAmFIiVZMzgiwOPU0Krs0DX21BJ3eJv4yQsyG
# LiCvMFlVQB1PTUlI+G+xNFR7bnaxUpx6foqlEZC6sZvoqx9s9eeLzOBSnnB38Pzb
# 0IZ/FvRfhBZAjghwXaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjNaMC8GCSqGSIb3DQEJBDEiBCDk4+O8PFsI+nkNwsG46hL7CQfRiqZqxC6Vy3MT
# dZSpmDANBgkqhkiG9w0BAQEFAASCAgBZl1YYYuEKjuqcuVzBlPYlGNK4er/OJ75V
# e32iZXr/tFW6pB7IYZplIJHmuTsgUT7wkd7FB5NY2sGe/xmXI7NhqGYJ6xRuDig7
# n6AhvcTYIbZOc+B8zwzDezIZe1J0zzPqxudsXK8mxfu9CSLyfQqMukdd/L1saH/D
# qBWDDD8VLipV0uKl/fZ5A7FJVFqF1mlNE0SrW5ZGkb+h5aiF2FkvHJCvW3l1I9IM
# G/Yg7jcdg3dGU+vcJgMfXlHaev8eibu4QA9Jq/kFXDtDuXJ5jxfueyqF5yw7idaQ
# d96WjT5inl7+O3ZmElzqosSMXjZ9T3KLxMrmuqcGcJ8MfXtB0AkMWZ+9Z0O82G6e
# kH6NWnilNZ3PVFCOiuYAAbOn4ew/jCl/g/Q8u5uIQRxrR0hKLNlBzH4TRpOiK8Es
# 2OGwcaZ1ykWSHC5vhw/AGsCQLw/RRI5hs4tFf1AwmIxYZBa6JciAG8bLdPHOZAU4
# EMCUQP+NDzwYB1BjY+j+WH2G4L+EUOya+ZrQeNncpOswfdNf5fw9c4iwtSZ2t9R5
# lIfzhnB8SU5oUiPGb5RBuMrdU91t04uIXULeqlM5F/DwrzoBgq6n92EnT2WcLLey
# oReJQcnCzbuHlIuK0kGWdj5X7AThP8VtPA8gR9TNIoBfgPGKEN4t/geftSMlzqTk
# m1AGwL1Mxw==
# SIG # End signature block
