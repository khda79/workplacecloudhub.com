Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot 'SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop

function Invoke-SmartM365SharePointJsonNameTransition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LocalFilePath,
        [Parameter(Mandatory)][string]$DriveId,
        [Parameter(Mandatory)][string]$EncodedTargetPath,
        [Parameter(Mandatory)][scriptblock]$Request,
        [Parameter(Mandatory)][scriptblock]$Download
    )
    if (-not $EncodedTargetPath.EndsWith('.json.txt',[StringComparison]::OrdinalIgnoreCase)) { throw 'SharePoint JSON transition requires an exact preferred target.' }
    $policy = Get-SmartM365JsonTransportPolicy
    if ($policy.Mode -ne 'JsonText') { return }
    if (-not $policy.ContainsKey('QualifiedSharePointDrives') -or $DriveId -notin @($policy.QualifiedSharePointDrives)) { throw 'SharePoint JSON transition requires prior drive qualification.' }
    $baseUri = 'https://graph.microsoft.com/v1.0/drives/' + [uri]::EscapeDataString($DriveId)
    $legacyPath = $EncodedTargetPath.Substring(0,$EncodedTargetPath.Length-4)
    $newName = [uri]::UnescapeDataString(($EncodedTargetPath -split '/')[-1])
    $journalPath = $LocalFilePath + '.sharepoint-transition.log'
    $lockPath = $LocalFilePath + '.sharepoint-transition.lock'
    $journalParent = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($journalPath))
    if (-not (Test-Path -LiteralPath $journalParent -PathType Container)) { throw 'Local owner directory missing for SharePoint transition journal.' }
    $transport = Get-Module SmartM365.JsonTransport
    foreach ($candidate in @($LocalFilePath,$journalPath,$lockPath)) {
        & $transport { param($path) Assert-SmartM365JsonUnlinkedPath $path } $candidate
    }
    $lock = [IO.File]::Open($lockPath,'OpenOrCreate','ReadWrite','None')
    $temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-RemoteJson-' + [guid]::NewGuid().ToString('N') + '.download')
    function Get-RemoteItem([string]$Uri) {
        try { $item = & $Request 'GET' $Uri $null @{}; if ($item -is [Collections.IDictionary]) { return [pscustomobject]$item }; return $item }
        catch {
            if ($_.Exception.Data.Contains('StatusCode') -and [int]$_.Exception.Data['StatusCode'] -eq 404) { return $null }
            throw
        }
    }
    function Get-RemoteHash([string]$ItemId) {
        $null = & $Download ($baseUri + '/items/' + [uri]::EscapeDataString($ItemId) + '/content') $temporary
        if (-not [IO.File]::Exists($temporary)) { throw 'Remote content download did not produce bytes.' }
        $hash = (Get-FileHash -LiteralPath $temporary -Algorithm SHA256 -ErrorAction Stop).Hash
        [IO.File]::Delete($temporary)
        return $hash
    }
    function Get-RemoteVersions([string]$ItemId) {
        $uri = $baseUri + '/items/' + [uri]::EscapeDataString($ItemId) + '/versions'
        $versions = @()
        $seen = @{}
        while ($uri) {
            if ($seen.ContainsKey($uri) -or -not $uri.StartsWith($baseUri + '/', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected or repeated versions continuation link.' }
            $seen[$uri]=$true
            $page = & $Request 'GET' $uri $null @{}
            if ($page -is [Collections.IDictionary]) { $page = [pscustomobject]$page }
            if ($null -eq $page -or -not $page.PSObject.Properties['value']) { throw 'Invalid remote versions response.' }
            foreach ($version in @($page.value)) { $versions += [pscustomobject]@{Id=[string]$version.id;Size=[int64]$version.size} }
            $uri = if ($page.PSObject.Properties['@odata.nextLink']) { [string]$page.'@odata.nextLink' } else { '' }
        }
        return $versions
    }
    function Write-TransitionEvent([string]$Phase,$Record) {
        $event = [ordered]@{Utc=[datetime]::UtcNow.ToString('o');Phase=$Phase;DriveId=$DriveId;Target=$EncodedTargetPath;Record=$Record}
        # Separate any interrupted final record from this complete, flushed event.
        $bytes=[Text.UTF8Encoding]::new($false).GetBytes([Environment]::NewLine+($event|ConvertTo-Json -Depth 12 -Compress)+[Environment]::NewLine)
        $stream=[IO.File]::Open($journalPath,'Append','Write','Read')
        try {$stream.Write($bytes,0,$bytes.Length);$stream.Flush($true)} finally {$stream.Dispose()}
    }
    function Confirm-RenamedItem($Record) {
        $item = Get-RemoteItem ($baseUri + '/items/' + [uri]::EscapeDataString([string]$Record.Id))
        if (-not $item -or $item.id -ne $Record.Id -or $item.name -ne $Record.Name) { throw 'Remote rename identity/name readback failed.' }
        if ((Get-RemoteHash $item.id) -ne $Record.SHA256) { throw 'Remote content changed during rename; no upload or deletion is permitted.' }
        $versions=@(Get-RemoteVersions $item.id)
        foreach ($before in @($Record.Versions)) {
            if (@($versions|Where-Object {$_.Id -eq $before.Id -and $_.Size -eq $before.Size}).Count -ne 1) { throw 'Remote version history verification failed after rename.' }
        }
        return $item
    }
    try {
        $last = $null
        if ([IO.File]::Exists($journalPath)) {
            $lines = @(Get-Content -LiteralPath $journalPath -Tail 32 -ErrorAction Stop)
            for ($i=$lines.Count-1; $i -ge 0; $i--) {
                if ([string]::IsNullOrWhiteSpace($lines[$i])) { continue }
                try { $event=$lines[$i]|ConvertFrom-Json -ErrorAction Stop }
                catch { Write-Warning 'Incomplete SharePoint transition journal tail retained; recovering from the last durable event.'; continue }
                if (-not $event -or -not $event.PSObject.Properties['Phase'] -or $event.Phase -notin @('Prepared','Completed') -or -not $event.PSObject.Properties['Record']) { throw 'Unrecognized SharePoint transition event; owner review required.' }
                $last=$event; break
            }
            if (-not $last) { throw 'No durable SharePoint transition event is available; owner review required.' }
        }
        if ($last -and ($last.DriveId -ne $DriveId -or $last.Target -ne $EncodedTargetPath)) { throw 'SharePoint transition journal target mismatch.' }
        $legacy = Get-RemoteItem ($baseUri + '/root:/' + $legacyPath)
        $preferred = Get-RemoteItem ($baseUri + '/root:/' + $EncodedTargetPath)
        if ($last -and $last.Phase -eq 'Prepared') {
            $byId = Get-RemoteItem ($baseUri + '/items/' + [uri]::EscapeDataString([string]$last.Record.Id))
            if ($byId -and $byId.name -eq $last.Record.Name) {
                $confirmed = Confirm-RenamedItem $last.Record
                Write-TransitionEvent 'Completed' $last.Record
                return [pscustomobject]@{Status='Recovered';Item=$confirmed}
            }
            if (-not $legacy -or $legacy.id -ne $last.Record.Id -or $legacy.eTag -ne $last.Record.ETag) { throw 'Prepared remote rename changed externally; qualification required.' }
        }
        if (-not $legacy) { return [pscustomobject]@{Status='NoLegacy';Item=$preferred} }
        if (-not $legacy.PSObject.Properties['file'] -or [string]::IsNullOrWhiteSpace([string]$legacy.eTag)) { throw 'Remote legacy target is not a versioned file with an eTag.' }
        $oldHash = Get-RemoteHash $legacy.id
        $destinationName = $newName
        if ($preferred) {
            if ((Get-RemoteHash $preferred.id) -ne $oldHash) { throw 'Remote old/new JSON copies differ; both items and histories preserved for explicit conflict resolution.' }
            # Identical bytes do not imply identical item history. Keep the legacy item
            # and its ID under an explicit archival name; never delete it or merge histories.
            $sha=[Security.Cryptography.SHA256]::Create()
            try {$idSuffix=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes([string]$legacy.id)))).Replace('-','').Substring(0,16)} finally {$sha.Dispose()}
            $destinationName = $newName.Substring(0,$newName.Length-9) + '.legacy-' + $idSuffix + '.json.txt'
            $slash=$EncodedTargetPath.LastIndexOf('/')
            $archivePath=if($slash -ge 0){$EncodedTargetPath.Substring(0,$slash+1)+[uri]::EscapeDataString($destinationName)}else{[uri]::EscapeDataString($destinationName)}
            if (Get-RemoteItem ($baseUri + '/root:/' + $archivePath)) { throw 'Remote legacy archive name already exists; no item overwritten.' }
        }
        $record=[pscustomobject]@{Id=[string]$legacy.id;ETag=[string]$legacy.eTag;Name=$destinationName;SHA256=$oldHash;Versions=@(Get-RemoteVersions $legacy.id)}
        Write-TransitionEvent 'Prepared' $record
        $body=@{name=$destinationName;'@microsoft.graph.conflictBehavior'='fail'}|ConvertTo-Json -Compress
        $null = & $Request 'PATCH' ($baseUri + '/items/' + [uri]::EscapeDataString([string]$legacy.id)) $body @{'If-Match'=[string]$legacy.eTag}
        $confirmed=Confirm-RenamedItem $record
        Write-TransitionEvent 'Completed' $record
        [pscustomobject]@{Status='Renamed';Item=$confirmed}
    } finally {
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
        $lock.Dispose()
    }
}

function Receive-SmartM365SharePointJsonBytes {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DriveId,[Parameter(Mandatory)][string]$EncodedTargetPath,
        [Parameter(Mandatory)][scriptblock]$Request,[Parameter(Mandatory)][scriptblock]$Download)
    if($EncodedTargetPath -notmatch '\.json(?:\.txt)?$'){throw 'An exact JSON remote name is required.'}
    $legacy=$EncodedTargetPath -replace '\.json\.txt$','.json'
    $base='https://graph.microsoft.com/v1.0/drives/'+[uri]::EscapeDataString($DriveId)
    $folder=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonRead-'+[guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($folder) | Out-Null
    $selected=$null
    try{
        foreach($suffix in '.txt',''){
            try{$item=& $Request 'GET' ($base+'/root:/'+$legacy+$suffix) $null @{}}
            catch{if($_.Exception.Data.Contains('StatusCode') -and [int]$_.Exception.Data['StatusCode'] -eq 404){continue};throw}
            if($item -is [Collections.IDictionary]){$item=[pscustomobject]$item}
            if(-not $item -or -not $item.PSObject.Properties['file'] -or -not $item.PSObject.Properties['eTag']){throw 'Remote JSON target is not a file with an eTag.'}
            $id=[uri]::EscapeDataString([string]$item.id)
            $path=Join-Path $folder ('payload.json'+$suffix)
            $null=& $Download ($base+'/items/'+$id+'/content') $path
            $after=& $Request 'GET' ($base+'/items/'+$id) $null @{}
            if($after.id -ne $item.id -or $after.eTag -ne $item.eTag){throw 'Remote JSON changed during download.'}
            if(-not $selected){$selected=$path}
            # Validate the preferred bytes before even considering the old name.
            Read-SmartM365JsonDocument $path | Out-Null
        }
        if(-not $selected){throw 'Both remote JSON names are absent.'}
        $document=Read-SmartM365JsonDocument (Join-Path $folder 'payload.json')
        return ,([IO.File]::ReadAllBytes($document.Path))
    }finally{
        foreach($name in 'payload.json','payload.json.txt'){
            $path=Join-Path $folder $name
            if([IO.File]::Exists($path)){[IO.File]::Delete($path)}
        }
        [IO.Directory]::Delete($folder)
    }
}

Export-ModuleMember -Function Invoke-SmartM365SharePointJsonNameTransition,Receive-SmartM365SharePointJsonBytes

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCApJO2v+77qoEaZ
# aK7ftP6wXrqOgx1x8Pi8xtVFZOkI/aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFSdtbveyYUydP0ZYcduVOCiXDqYFMwgo6lpvS53Qm/GMA0GCSqG
# SIb3DQEBAQUABIIBgGZkiFc3JjqSy+FRl9EjsvdwcRdweod00cnmxlWSI9z8M3tN
# 9tuxkMvBBP8jpUXSCqIrKUMagLTIGzjl2yQWwDb7bO36kC5Qdb4S2CRnyC9j1neJ
# SWGll4YW+OhD5RlHVyccqArfkqskIBCLHvH0BdntQd2xFhuR8LKaN9wLoMl4f5uy
# MfEpmbKR/UmDbGp2wiyZVVrSSUqT//AsXR0VJbByOe03Bt1vPfdMC3fKQFV08pHj
# 7cCeebNJYdiEPBwfo9E/6tW6xWFszI4kbybVe1E5hBD1lZar0WBLuu+WBrv6ujXV
# H3UQTFvRGO/MZ9WPrSZNDLCszZ4uJumeqgv07MtvJ+WhGBAQfuRpv6IgHj4RwXcL
# 0bnJ3q//WkbZ6kcfjpdI35kLcE9ftPm5qlfs/a/OAkZ1JUkOUmOS1RP0tjREeago
# lX6IeV8rS3BIJmkT3yCDNvAGa58unc/fr8YMYgAh0U0PZbUoiuALYLqVkpaPWo4q
# WgDOWqByoUbfM+JZ06GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU3
# MzhaMC8GCSqGSIb3DQEJBDEiBCAQUXjVZq1nvUP3vRLSAUhFtmAySyYV0FQgHeHQ
# czz7UTANBgkqhkiG9w0BAQEFAASCAgAqstdBjpBVSztuhbfacBaT+FQEDGPLXNd7
# y0Pd1/JK/ak4J3GppuHItK6ZGlzH00VPuOYd28N6ZVjeKTvz/RjjamZ8SOnJHUXz
# MVAV/lqPR1xbG3virBMjdEeMDlXj7DaTY60JWFEH2ADIDFrtCy9bMSy2+VDTOZ2E
# 8guvWpNOnjIj4Tk4QwaPzB4eTMGOabk56xq4Bhi+PiQB1YuTy3JODbIRHxNRQqUB
# 2JjFAc21h/BtZ653R0GAQ048e/1y9sAtnoMXaOv1TqwUmr9RMtlCP5THKfOYhtQT
# RnNTVTEL6dHfHYjat37jlTj/d+Q/pQwkVP2Wo+1ZVEJk6aYrjzPjuR0nRuuZs0Yw
# l5PsWRdWNa2pD0N88j6JsX3ybw8AQ81rzlMSjqukOwith47hYTAVaAQ6DO37b9Bn
# dVVCh3wYpS6HzfJxfTvX2b9pWv7v+AS7vLqojWjThQ/6f0mausipU1/0DYkuPjut
# FT3iyGbh6kAJEtlb598m20kl2MIRHPqggWhe/D81ArMhgeBnLqrdn3JM6OsFvlEB
# iiI4ny8o7S1oNkdPfGizV+tXI5Nl8HujhSZPIl9RmFJ/8cMj06eS+huZgJ2Q3wBD
# E6SR3OjL/Mg8mMDcZOpOJ/mfq/JvA5I1f+BE+aG+AOm478ZZH02RPhD0Jwf6r84f
# 36JcMQ2Now==
# SIG # End signature block
