#Requires -Version 7.0
# Current-only transport. Callbacks provide the shared Core upload/download APIs.
# Version 0.2.0: isolate persistent transition state from the validated cohort.
Set-StrictMode -Version Latest

function Assert-SmartM365CmdbTransitionPath {
    param([Parameter(Mandatory)][string]$Path)
    $cursor=[IO.Path]::GetFullPath($Path)
    while($cursor){
        if(Test-Path -LiteralPath $cursor){
            if((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){
                throw 'Linked CMDB transition state path refused.'
            }
        }
        $parent=[IO.Path]::GetDirectoryName($cursor)
        if($parent -eq $cursor){break}
        $cursor=$parent
    }
}

function Resolve-SmartM365CmdbTransitionStateRoot {
    param([string]$PreparedRoot,[string]$StateRoot)
    $prepared=[IO.Path]::GetFullPath($PreparedRoot).TrimEnd('\','/')
    $state=[IO.Path]::GetFullPath($StateRoot).TrimEnd('\','/')
    if((Split-Path $prepared -Leaf) -cne 'DATA-POWERBI-CMDB' -or
       $state -ieq $prepared -or $state.StartsWith($prepared+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){
        throw 'CMDB transition state must be outside the prepared cohort.'
    }
    Assert-SmartM365CmdbTransitionPath $prepared
    Assert-SmartM365CmdbTransitionPath $state
    return $state
}

function Move-SmartM365CmdbTransitionState {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PreparedRoot,
          [Parameter(Mandatory)][string]$StateRoot,
          [Parameter(Mandatory)][hashtable]$Identity)
    if(-not $IsWindows){throw 'CMDB transition recovery requires Windows locking.'}
    $state=Resolve-SmartM365CmdbTransitionStateRoot $PreparedRoot $StateRoot
    if(-not(Test-Path -LiteralPath $PreparedRoot)){return 0}
    $lock=$null;$locked=$false;$handles=@();$moved=@()
    try {
        $lockPath=Join-Path (Split-Path ([IO.Path]::GetFullPath($PreparedRoot)) -Parent) '.cmdb-preparation.lock'
        Assert-SmartM365CmdbTransitionPath $lockPath
        $lock=[IO.File]::Open($lockPath,'OpenOrCreate','ReadWrite','ReadWrite')
        $lock.Lock(0,1);$locked=$true
        $names=@('current.json.txt.sharepoint-transition.lock','current.json.txt.sharepoint-transition.log')
        $legacy=@($names|Where-Object{Test-Path -LiteralPath (Join-Path $PreparedRoot $_)})
        if(-not $legacy.Count){return 0}
        $manifestPath=Join-Path $PreparedRoot 'current.json.txt'
        Assert-SmartM365CmdbTransitionPath $manifestPath
        $manifest=Get-Content -LiteralPath $manifestPath -Raw|ConvertFrom-Json
        if($manifest.Owner -cne 'SmartInventory-CMDB-Prepared' -or $manifest.Status -cne 'Validated'){
            throw 'Transition recovery requires an owned validated CMDB manifest.'
        }
        foreach($field in 'TenantKey','OrganizationKey','EnvironmentKey','TenantId'){
            if([string]::IsNullOrWhiteSpace([string]$Identity[$field]) -or $manifest.Identity.$field -cne $Identity[$field]){
                throw 'Transition recovery reporting identity mismatch.'
            }
        }
        # Keep handles open through relocation: FileShare.Delete permits our
        # rename, but prevents a legacy exclusive transition from acquiring them.
        foreach($name in $legacy){
            $source=Join-Path $PreparedRoot $name
            $destination=Join-Path $state $name
            Assert-SmartM365CmdbTransitionPath $source
            if(Test-Path -LiteralPath $destination){throw 'Transition recovery destination already exists; nothing overwritten.'}
            $handle=[IO.File]::Open($source,'Open','ReadWrite','Delete')
            $handles+=,$handle
            if($name.EndsWith('.lock') -and $handle.Length -ne 0){throw 'Nonempty legacy transition lock refused.'}
        }
        $null=New-Item -Path $state -ItemType Directory -Force
        foreach($name in $legacy){
            [IO.File]::Move((Join-Path $PreparedRoot $name),(Join-Path $state $name))
            $moved+=,$name
        }
        return $moved.Count
    } catch {
        [array]::Reverse($moved)
        foreach($name in $moved){
            [IO.File]::Move((Join-Path $state $name),(Join-Path $PreparedRoot $name))
        }
        throw
    } finally {
        foreach($handle in $handles){$handle.Dispose()}
        if($lock){if($locked){$lock.Unlock(0,1)};$lock.Dispose()}
    }
}

$ErrorActionPreference='Stop'

function Resolve-SmartM365CmdbSharePointFolder {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$NormalizedDataRoot)
    $root=($NormalizedDataRoot -replace '\\','/').Trim('/')
    if($root -notmatch '(^|/)DATA$' -or $root -match '(^|/)\.\.?(/|$)' -or $root -match '//|:|\{\{|__USE_GLOBAL__'){
        throw 'CMDB publication requires a resolved SharePoint DATA root.'
    }
    return $root+'/DATA-POWERBI-CMDB'
}

function Get-SmartM365CmdbTransferPlan {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PreparedRoot,
          [Parameter(Mandatory)][hashtable]$Identity,
          [Parameter(Mandatory)][string]$PythonPath,
          [string]$ExpectedManifestSHA256)
    $arguments=@('-B',(Join-Path $PSScriptRoot 'cmdb_transfer.py'),'--root',$PreparedRoot)
    foreach($pair in @(@('tenant-key','TenantKey'),@('organization-key','OrganizationKey'),
                        @('environment-key','EnvironmentKey'),@('tenant-id','TenantId'))){
        if([string]::IsNullOrWhiteSpace([string]$Identity[$pair[1]])){throw 'Complete reporting identity is required.'}
        $arguments+=@(('--'+$pair[0]),([string]$Identity[$pair[1]]))
    }
    if($ExpectedManifestSHA256){$arguments+=@('--expected-manifest-sha256',$ExpectedManifestSHA256)}
    $result=& $PythonPath @arguments 2>&1
    if($LASTEXITCODE -ne 0){throw ('CMDB transfer verification failed: '+(@($result | Select-Object -Last 1) -join ''))}
    $plan=($result -join [Environment]::NewLine) | ConvertFrom-Json
    Assert-SmartM365CmdbTransferFreshness -Plan $plan
    return $plan
}

function Assert-SmartM365CmdbTransferFile {
    param([string]$Path,[object]$Record)
    $item=Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if($item.PSIsContainer -or $item.Attributes -band [IO.FileAttributes]::ReparsePoint -or
       $item.Length -ne $Record.Bytes -or
       (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -cne $Record.SHA256){
        throw "CMDB transfer byte mismatch: $($Record.Name)."
    }
}

function Assert-SmartM365CmdbTransferFreshness {
    param([object]$Plan)
    $expiry=[DateTimeOffset]::Parse($Plan.EarliestSourceExpiryUtc,[Globalization.CultureInfo]::InvariantCulture)
    if([DateTimeOffset]::UtcNow -gt $expiry){throw 'CMDB source evidence expired during transfer; regenerate from fresh acquisitions.'}
}

function Send-SmartM365CmdbPreparedSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$PreparedRoot,
          [Parameter(Mandatory)][hashtable]$Identity,
          [Parameter(Mandatory)][string]$PythonPath,
          [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$ExpectedManifestSHA256,
          [Parameter(Mandatory)][scriptblock]$UploadFile,
          [Parameter(Mandatory)][scriptblock]$DownloadFile,
          [scriptblock]$Progress)
    if(-not $IsWindows){throw 'CMDB publication requires Windows byte-range locking compatible with the generator.'}
    $root=[IO.Path]::GetFullPath($PreparedRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
    if((Split-Path $root -Leaf) -cne 'DATA-POWERBI-CMDB' -or -not (Test-Path -LiteralPath $root -PathType Container)){
        throw 'Use the existing current DATA-POWERBI-CMDB directory.'
    }
    if((Get-Item -LiteralPath $root -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Linked CMDB prepared root refused.'}
    $lock=$null; $locked=$false; $scratch=$null
    try {
        # Same byte-range lock as cmdb_prepare.PublicationLock (msvcrt on Windows).
        # The persistent one-byte lock file is coordination, not a batch backup.
        $lockPath=Join-Path (Split-Path $root -Parent) '.cmdb-preparation.lock'
        if((Test-Path -LiteralPath $lockPath) -and
           (Get-Item -LiteralPath $lockPath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){
            throw 'Linked CMDB publication lock refused.'
        }
        $lock=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::ReadWrite)
        $lock.Lock(0,1); $locked=$true
        if($lock.Length -eq 0){$lock.WriteByte(48);$lock.Flush()}
        $parameters=@{PreparedRoot=$root;Identity=$Identity;PythonPath=$PythonPath;ExpectedManifestSHA256=$ExpectedManifestSHA256}
        $plan=Get-SmartM365CmdbTransferPlan @parameters
        $scratch=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CmdbTransfer-'+[guid]::NewGuid().ToString('N'))
        $null=New-Item -Path $scratch -ItemType Directory
        $destination=Join-Path $scratch 'readback.bin'
        $verified=0
        foreach($record in @($plan.Files)){
            if($record.Name -ceq 'current.json.txt'){
                # No pointer until every CSV has been uploaded and read back.
                # Recheck the whole local cohort and current source age at this boundary.
                $null=Get-SmartM365CmdbTransferPlan @parameters
            }
            Assert-SmartM365CmdbTransferFreshness -Plan $plan
            $localPath=Join-Path $root $record.Name
            Assert-SmartM365CmdbTransferFile -Path $localPath -Record $record
            if($Progress){& $Progress $record.Name ($verified+1) @($plan.Files).Count | Out-Null}
            $uploaded=& $UploadFile $localPath $record.Name
            if(-not $uploaded){throw "SharePoint upload was not confirmed: $($record.Name)."}
            if(Test-Path -LiteralPath $destination){Remove-Item -LiteralPath $destination -Force}
            $downloaded=& $DownloadFile $destination $record.Name
            if(-not $downloaded){throw "SharePoint read-back was not confirmed: $($record.Name)."}
            Assert-SmartM365CmdbTransferFile -Path $destination -Record $record
            Assert-SmartM365CmdbTransferFile -Path $localPath -Record $record
            Assert-SmartM365CmdbTransferFreshness -Plan $plan
            Remove-Item -LiteralPath $destination -Force
            $verified++
        }
        return [pscustomobject]@{Status='PublishedAndReadBack';CsvFiles=$plan.CsvFiles;VerifiedFiles=$verified;
            ManifestSHA256=$plan.ManifestSHA256;ManifestVerified=$true;
            EarliestSourceExpiryUtc=$plan.EarliestSourceExpiryUtc;FreshnessWarnings=@($plan.FreshnessWarnings)}
    } finally {
        try {
            if($scratch -and (Test-Path -LiteralPath $scratch)){
                $scratchFull=[IO.Path]::GetFullPath($scratch)
                if((Split-Path $scratchFull -Parent).TrimEnd('\','/') -cne ([IO.Path]::GetTempPath()).TrimEnd('\','/') -or
                   (Split-Path $scratchFull -Leaf) -notmatch '^SmartM365-CmdbTransfer-[a-f0-9]{32}$'){
                    throw 'Unsafe transient read-back cleanup target refused.'
                }
                Remove-Item -LiteralPath $scratchFull -Recurse -Force
            }
        } finally {
            if($lock){try {if($locked){$lock.Unlock(0,1)}} finally {$lock.Dispose()}}
        }
    }
}

Export-ModuleMember -Function Resolve-SmartM365CmdbSharePointFolder,Get-SmartM365CmdbTransferPlan,Send-SmartM365CmdbPreparedSnapshot,Move-SmartM365CmdbTransitionState

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDqkqmu8DmIyhXp
# ggxFO/wCb5UIxQNq2yUzFfkWfb6RZqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP6icJYk5Pki9Ts9oRJk9sJ8yqPHx4zPrriC0UhXjQMjMA0GCSqG
# SIb3DQEBAQUABIIBgB9ooy8m1rnhkh9ZDZMULCswVpX3sGHhA5sY0daYjATags93
# YbkHv/U0ZFVTQ5GLNaVaMmL9qgzRhEvFjWk4ALX3ZViGDs7dktlDVMiLI6UI9ize
# AjijxwjHPOOBZMgdvl2U2j76enrdGpNaLWZjtcGteFPsXNNUhg2Ofdx53CxFu/9c
# qbN35aU4r7Vlb7RN5pDahuwyNQXNt6z4xMFwFU45GgHCGWclYbv0jSJhCmw7aYKS
# pXpCTS0ghm7CS9q4R+bJ7YF9eiz+v0JA3/Xv+xpeonauWiLpg6tZarm+x6O/GrCj
# CHSJ+Z3wKE9cuRm8FSr+JtC7Iy998LrFTiyZzh9/a85RCM2ClDTLhct9vt4IGAEk
# vgAR1oYmJU8efHYb9UCCBTpzobTz54KzZIJeyhH2gaYStPwkxvERBTMo5w4xtYHW
# tR8GsZrdLzR31HqlCmPMYNQ0vnvROE+j7prNF8eQHPNeANHKxNymQhK3PUb+2ccB
# 3MnHzZuUQ9ozufAgdqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxMTM2
# NTFaMC8GCSqGSIb3DQEJBDEiBCCVc1VeGWK00qrNdozm0RVcRTOXUWlUJa2gbSv7
# OPP5FTANBgkqhkiG9w0BAQEFAASCAgCpsgKyiOcQUcgqR8UzjKktsQ5CSJsck15K
# RikXrT6IGZH7AFttngUDkVJ9BfY+OrhGF5c3UgYzUbfS3lXXiZYgx3/jKUL8oki1
# 6tUyONUdLsIFLOmpOIghrouVrQ6SNn/uOirft4hsqSSPsBHVbWO5GI80QcswQgoB
# Xn66bNAfg5CN1SJx7q5OMSCBa++Cb8uUuWaa9V77qTKDqOFZew/kN+pFjRl1UEFb
# 4tA0A8GmcI3f8oY7+k5Fermj5XRdJEw/End/uOSk1LoAmVs5ttaJ5gbxt4NXQZmW
# zDibFga5FvjwBC1K2rY33ekqVBjTd+LTwhvbcoG5+P5rYjmkiIxdV6bf6F4VXmg/
# bVaqQrFPYu3TNRE9nLD/3Ni2hpDpFoAr8Job4jyRhlcGoSBp2m7zxLxhRqfky5Gs
# L7gYJw4GCcr27D1aqvY1QFzlHNVMOlParpNSHWMRUbw6inrV8WEkO2PxLRJsh1GZ
# zmLaQ7IT8do3Vwxg3sK7OBFWjgYE4dqje2eENUesFWxuSp9vdDnVqjaK6HrVarxY
# I4BPJvyU2XOZZQk4GEoGp6e+XhZD3jd6o3ZtnPwAlKPpRK4uFNl8KzSTct4g8Lmw
# xEf+U9gwQfTflHA15ctjwGD5/wN0sNgaTP8S4dZYHgNFOwNG/da+EvazO28j/rHh
# 1Qawi7PW7A==
# SIG # End signature block
