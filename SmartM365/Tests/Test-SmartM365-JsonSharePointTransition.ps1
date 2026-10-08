[CmdletBinding()]
param()
<#
.VERSION
1.0.1
#>
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.SharePointJsonTransition.psd1') -Force
$module=Get-Module SmartM365.SharePointJsonTransition
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonSharePoint-' + [guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory $root
$script:passed=0
function Check([bool]$Value,[string]$Message){if(!$Value){throw $Message};$script:passed++}
function Reject([scriptblock]$Action,[string]$Message){$failed=$false;try{& $Action|Out-Null}catch{$failed=$true};Check $failed $Message}
function Reset-Fixture([string]$Name) {
    $script:items=@{};$script:patches=0;$script:interrupt=$false;$script:etagConflict=$false;$script:forbidden=$false
    $script:pagination=$false;$script:badPageUri='';$script:versionPages=0;$script:alterAfterDownload=$false;$script:nullOnMissing=$false
    $script:local=Join-Path $root ($Name+'.json.txt')
    [IO.File]::WriteAllText($script:local,'{"Generation":2}')
}
function Add-Remote([string]$Id,[string]$Name,[string]$Content) {
    $script:items[$Id]=[pscustomobject]@{id=$Id;name=$Name;eTag='etag-1';file=[pscustomobject]@{};Bytes=[Text.Encoding]::UTF8.GetBytes($Content);Versions=@([pscustomobject]@{id='1.0';size=17},[pscustomobject]@{id='2.0';size=17})}
}
$request={
    param($method,$uri,$body,$headers)
    if($method -notin @('GET','PATCH')){throw 'Unexpected remote mutation; no deletion is allowed.'}
    if($script:forbidden){$e=[IO.IOException]::new('Synthetic forbidden');$e.Data['StatusCode']=403;throw $e}
    $item=$null
    if($uri -match '/versions(?:\?|$)' -and $script:pagination) {
        $script:versionPages++
        $id=if($uri -match "/items\('([^']+)'\)/versions"){$Matches[1]}elseif($uri -match '/items/([^/]+)/versions'){$Matches[1]}else{throw 'Unexpected paged mock route'}
        $item=$script:items[$id]
        if($uri -match 'skiptoken=page2'){return [pscustomobject]@{value=@($item.Versions[1])}}
        $next=if($script:badPageUri){$script:badPageUri}else{"https://graph.microsoft.com/v1.0/drives('synthetic')/items('$id')/versions?`$skiptoken=page2"}
        return [pscustomobject]@{value=@($item.Versions[0]);'@odata.nextLink'=$next}
    }
    if($uri -match '/root:/(.+)$') {
        $name=[uri]::UnescapeDataString(($Matches[1] -split '/')[-1])
        $item=@($script:items.Values|Where-Object name -eq $name)|Select-Object -First 1
    } elseif($uri -match '/items/([^/]+)(/versions)?$') {
        $item=$script:items[$Matches[1]]
        if($item -and $Matches[2]){return [pscustomobject]@{value=@($item.Versions)}}
    } else {throw "Unexpected mock URI: $uri"}
    if(!$item){if($script:nullOnMissing){return $null};$e=[IO.IOException]::new('Synthetic missing');$e.Data['StatusCode']=404;throw $e}
    if($method -eq 'GET'){return [pscustomobject]@{id=$item.id;name=$item.name;eTag=$item.eTag;file=$item.file}}
    if($script:etagConflict){$item.eTag='changed'}
    if($headers['If-Match'] -ne $item.eTag){$e=[IO.IOException]::new('Synthetic stale eTag');$e.Data['StatusCode']=412;throw $e}
    $change=$body|ConvertFrom-Json
    Check ($change.'@microsoft.graph.conflictBehavior' -eq 'fail') 'Rename did not forbid overwriting another remote item.'
    if(@($script:items.Values|Where-Object {$_.name -eq $change.name -and $_.id -ne $item.id}).Count){throw 'Synthetic name conflict'}
    $script:patches++
    $item.name=$change.name;$item.eTag='etag-2'
    if($script:interrupt){$script:interrupt=$false;throw 'Synthetic response lost after rename'}
    return [pscustomobject]@{id=$item.id;name=$item.name;eTag=$item.eTag;file=$item.file}
}
$download={
    param($uri,$destination)
    if($uri -notmatch '/items/([^/]+)/content$'){throw 'Unexpected download URI'}
    [IO.File]::WriteAllBytes($destination,$script:items[$Matches[1]].Bytes)
    if ($script:alterAfterDownload) { $script:items[$Matches[1]].eTag='etag-changed' }
}
function Invoke-Fixture([switch]$CompareLocalContent,[string]$StateFolderPath) {
    Invoke-SmartM365SharePointJsonNameTransition -LocalFilePath $script:local -DriveId synthetic -EncodedTargetPath 'Folder/state.json.txt' -Request $request -Download $download -CompareLocalContent:$CompareLocalContent -StateFolderPath $StateFolderPath
}
try {
    & $module { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedSharePointDrives=@('synthetic')} } }
    foreach($uri in @('https://graph.microsoft.com/v1.0/drives/b!synthetic/items/item/versions?$skiptoken=opaque','https://graph.microsoft.com/v1.0/drives/b%21synthetic/items/item/versions?$skiptoken=opaque',"https://graph.microsoft.com/v1.0/drives('b!synthetic')/items('item')/versions?`$skiptoken=opaque")) {
        & $module {param($u)Assert-SmartM365VersionsPageUri $u 'b!synthetic' 'item'} $uri
        Check $true 'Equivalent Graph route rejected.'
    }
    foreach($uri in @('http://graph.microsoft.com/v1.0/drives/b!synthetic/items/item/versions','https://example.invalid/v1.0/drives/b!synthetic/items/item/versions','https://graph.microsoft.com/v1.0/drives/other/items/item/versions','https://graph.microsoft.com/v1.0/drives/b!synthetic/items/other/versions','https://graph.microsoft.com/v1.0/drives/b!synthetic/items/item/content')) {
        Reject {& $module {param($u)Assert-SmartM365VersionsPageUri $u 'b!synthetic' 'item'} $uri} 'Unsafe versions route accepted.'
    }
    Reset-Fixture paged;Add-Remote old 'state.json' '{"Generation":1}';$script:pagination=$true
    $result=Invoke-Fixture
    Check ($result.Status -eq 'Renamed' -and $script:versionPages -eq 4) 'Both version pages were not checked before and after rename.'
    Reset-Fixture repeated;Add-Remote old 'state.json' '{"Generation":1}';$script:pagination=$true;$script:badPageUri='https://graph.microsoft.com/v1.0/drives/synthetic/items/old/versions'
    Reject {Invoke-Fixture} 'Repeated continuation was accepted.'
    Check ($script:patches -eq 0) 'Repeated continuation performed a rename.'
    Reset-Fixture onlyold;Add-Remote old 'state.json' '{"Generation":1}'
    $result=Invoke-Fixture
    Check ($result.Status -eq 'Renamed' -and $script:items.old.name -eq 'state.json.txt') 'Old remote item not renamed.'
    Check ($script:items.old.id -eq 'old' -and $script:items.old.Versions.Count -eq 2) 'Remote identity or history changed.'
    Check ([Text.Encoding]::UTF8.GetString($script:items.old.Bytes) -eq '{"Generation":1}') 'Rename modified remote bytes.'
    $null=Invoke-Fixture
    Check ($script:patches -eq 1) 'Remote transition was not idempotent.'
    Reset-Fixture comparemissing
    $result=Invoke-Fixture -CompareLocalContent
    Check ($result.Status -eq 'NoLegacy' -and -not $result.ContentMatchesLocal) 'Absent preferred item was treated as unchanged.'
    Reset-Fixture compareequal;Add-Remote new 'state.json.txt' '{"Generation":2}'
    $result=Invoke-Fixture -CompareLocalContent
    Check ($result.Status -eq 'NoLegacy' -and $result.ContentMatchesLocal -and $script:patches -eq 0) 'Identical preferred bytes were not recognized.'
    Reset-Fixture comparedifferent;Add-Remote new 'state.json.txt' '{"Generation":1}'
    $result=Invoke-Fixture -CompareLocalContent
    Check ($result.Status -eq 'NoLegacy' -and -not $result.ContentMatchesLocal) 'Different preferred bytes were treated as unchanged.'
    Reset-Fixture comparechanged;Add-Remote new 'state.json.txt' '{"Generation":2}';$script:alterAfterDownload=$true
    Reject {Invoke-Fixture -CompareLocalContent} 'Preferred item changing during comparison was accepted.'
    Check ($script:patches -eq 0) 'Content comparison changed a remote item.'
    Reset-Fixture interrupted;Add-Remote old 'state.json' '{"Generation":1}';$script:interrupt=$true
    Reject {Invoke-Fixture} 'Lost rename response not exercised.'
    [IO.File]::AppendAllText(($script:local+'.sharepoint-transition.log'),'{"Phase":"Comp')
    $result=Invoke-Fixture
    Check ($result.Status -eq 'Recovered' -and $script:patches -eq 1) 'Rename response loss did not recover by stable item ID.'
    $null=Invoke-Fixture
    Check ($script:patches -eq 1) 'Truncated journal recovery was not idempotent.'
    Reset-Fixture same;Add-Remote old 'state.json' '{"Generation":1}';Add-Remote new 'state.json.txt' '{"Generation":1}'
    $null=Invoke-Fixture
    Check ($script:items.old.name -match '^state\.legacy-[A-F0-9]{16}\.json\.txt$' -and $script:items.new.name -eq 'state.json.txt') 'Identical pair did not retain both remote histories.'
    Check ($script:items.Count -eq 2 -and $script:items.old.Versions.Count -eq 2) 'Identical pair discarded a remote item or history.'
    Reset-Fixture different;Add-Remote old 'state.json' '{"Generation":1}';Add-Remote new 'state.json.txt' '{"Generation":2}'
    Reject {Invoke-Fixture} 'Divergent remote files were silently reconciled.'
    Check ($script:patches -eq 0 -and $script:items.Count -eq 2) 'Divergent remote files were modified.'
    Reset-Fixture stale;Add-Remote old 'state.json' '{"Generation":1}';$script:etagConflict=$true
    Reject {Invoke-Fixture} 'Remote concurrent modification ignored.'
    Check ($script:patches -eq 0 -and $script:items.old.name -eq 'state.json') 'Stale eTag changed the remote item.'
    Reset-Fixture denied;$script:forbidden=$true
    Reject {Invoke-Fixture} 'Authorization failure treated as file absence.'
    Check ($script:patches -eq 0) 'Authorization failure triggered mutation.'
    & $module { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedSharePointDrives=@()} } }
    Reset-Fixture unqualified
    Reject {Invoke-Fixture} 'Unqualified SharePoint drive accepted.'
    function Read-Fixture {Receive-SmartM365SharePointJsonBytes -DriveId synthetic -EncodedTargetPath 'Folder/state.json' -Request $request -Download $download}
    Reset-Fixture readold;Add-Remote old 'state.json' '{"Generation":1}'
    Check ([Text.Encoding]::UTF8.GetString((Read-Fixture)) -eq '{"Generation":1}') 'Absent new name did not read old bytes.'
    Add-Remote new 'state.json.txt' '{"Generation":1}'
    Check ([Text.Encoding]::UTF8.GetString((Read-Fixture)) -eq '{"Generation":1}') 'Identical pair read failed.'
    $script:items.new.Bytes=[Text.Encoding]::UTF8.GetBytes('{"Generation":2}')
    Reject {Read-Fixture} 'Divergent download pair accepted.'
    $script:items.new.Bytes=[Text.Encoding]::UTF8.GetBytes('incomplete{')
    Reject {Read-Fixture} 'Invalid new remote file fell back.'
    Reset-Fixture readmissing
    Reject {Read-Fixture} 'Absent download succeeded.'
    $script:nullOnMissing=$true
    Reject {Read-Fixture} 'Two quietly absent names were accepted.'
    Add-Remote new 'state.json.txt' '{"Generation":2}'
    Check ([Text.Encoding]::UTF8.GetString((Read-Fixture)) -eq '{"Generation":2}') 'Quiet legacy absence blocked preferred download.'
    Reset-Fixture quietold;$script:nullOnMissing=$true;Add-Remote old 'state.json' '{"Generation":1}'
    Check ([Text.Encoding]::UTF8.GetString((Read-Fixture)) -eq '{"Generation":1}') 'Quiet preferred absence blocked legacy download.'
    Reset-Fixture quiettransition;$script:nullOnMissing=$true
    & $module { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedSharePointDrives=@('synthetic')} } }
    Check ((Invoke-Fixture).Status -eq 'NoLegacy') 'Quiet absent names blocked new upload transition.'
    $script:forbidden=$true
    Reject {Read-Fixture} 'Read authorization error treated as absence.'
    Check ($script:patches -eq 0) 'Reading performed remote mutations.'
    & $module {function script:Get-SmartM365JsonTransportPolicy {@{Mode='JsonText';QualifiedSharePointDrives=@('synthetic')}}}
    Reset-Fixture external
    $externalState=Join-Path $root 'external-state'
    $result=Invoke-Fixture -StateFolderPath $externalState
    Check ($result.Status -eq 'NoLegacy') 'Explicit external state changed transition behavior.'
    Check (Test-Path -LiteralPath (Join-Path $externalState 'external.json.txt.sharepoint-transition.lock')) 'External lock missing.'
    Check (-not(Test-Path -LiteralPath ($script:local+'.sharepoint-transition.lock'))) 'External state polluted local owner directory.'
    Reject {Invoke-Fixture -StateFolderPath 'relative-state'} 'Relative transition state accepted.'
    Reset-Fixture externalrecover
    Add-Remote old 'state.json' '{"Generation":1}';$script:interrupt=$true
    Reject {Invoke-Fixture -StateFolderPath $externalState} 'Interrupted external-state rename succeeded.'
    Check (Test-Path -LiteralPath (Join-Path $externalState 'externalrecover.json.txt.sharepoint-transition.log')) 'External durable journal missing.'
    Check (-not(Test-Path -LiteralPath ($script:local+'.sharepoint-transition.log'))) 'Recovery journal polluted local owner directory.'
    $result=Invoke-Fixture -StateFolderPath $externalState
    Check ($result.Status -eq 'Recovered' -and $script:patches -eq 1) 'External-state recovery repeated or lost remote rename.'
    [pscustomobject]@{Passed=$script:passed;FixtureRoot=$root;Evidence='Mock Graph metadata, versions and bytes; no SharePoint connection or request'}
} finally { & $module {Remove-Item Function:script:Get-SmartM365JsonTransportPolicy -ErrorAction SilentlyContinue} }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAFSw4lduIzmC48
# DEas/99Hi5gCTWoqjAoA86JPumJgvKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILGG8U9XZRDJXxbt0pRUz4QvV1tO7jV5U61pZONN0CIfMA0GCSqG
# SIb3DQEBAQUABIIBgETK5dgL2PvQRZdJ/P2rimN7WcerHbP8/Q+4JileWQcrCjCw
# 6abOpnPzLaalre2yr3g8QPaZkushuyaMNNsu2MfzI5XwDMCoX0YJ6SabBuvvJ0Mg
# hLRTyZbymPzmZqSD3SVGjoAWWCDfi3cKU5NVlI5qKsY/717hFWf2Z20Xl1T3CN0y
# IOGRuE+KwA9pHEVpLdv/duVjVMJMM0E3gPEek9m2mM58TvioxeY74agYYKIFb5DZ
# 8X+yg2hjC/ZA1dL2ONyezu/iCgeMNJHxZClJg81w2aW97PYlcau6TDzNIRP/wJhy
# yhD4jU/SehXdzw7lHHrXE4UrX47Iw8TG4beKpz2Nzx7FllEkhRKO1O1SgshoOjwg
# QcH2tmsfE+5vzKV9FeFdR08OLl9V1AjyI9hssBTPTVXRvUukIGnkfG3U4Ks2dREc
# 3nV5o2lxqPhAwf8kLJqvcBsKI8HvESizMQCMwRWC0yzNQ459BfwASaDjTtU5hRX3
# u0X3wvymmfL8XeBOqKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgxMzIw
# MTRaMC8GCSqGSIb3DQEJBDEiBCBVtINqWOE0LBUwmYfdX2P+q4/yc+bMSRI4a/Dh
# CwhllDANBgkqhkiG9w0BAQEFAASCAgCkcAJpv4hTCVOKXLspg1zddI4EzjneugTr
# Iqn2+Qj77A0SZNwkMNic1fM+63iF2xByukVQN8hgtaNKOGKBV1r7zb4drPbjm2qN
# HFCtUzF/WvAeBwxJKbiK6eifJ9klDqxjPslGExZtwn91iiJlf4q0xKeW7m+6SjHz
# zUXXgCUATZRVf9lZdOvEApilb+PmqdGT93/nlCzZIcr+5uJWWExn1wuQppZHHMlQ
# 0ZRTLP8X3UBOzwb7bBg30mIDMuGthzr0Qtao1KFtlcAEm/zvcEsSmRdP2atpuh8f
# DDcKxQw7pQ6wrBL94R4RM7m3zK96z8o4gEGz6aFuwOSj1XeikoLWDOfczr5CcVfZ
# a1n2xyf9dsIM0R6YLUpZ6YKrhuz+eO5yBnJ7hzxypqBu6OQyNt6YJNrsoUGMMckN
# aZtbsYriJthuLoWddvvBGsu4UQEZdjhMfGRox68UdsiGcSwmt3WedrMeNZUSey2q
# j8Yl0Hw1Q/Q8O/rdXxMRk4wtE8CuVA90WiXtmhMuRP5q4/gewBUa7HP0oBhVOzcI
# lkHNgri++lWRBxge+AWhsnkRp4seChl8nQvGspd7HpCuat+1qRriXFvtmn8xHmQa
# +xi6CHGpzqiNWWSymbbhdT6SE9EJOclJdMNJ320LRHwbiD5zLDmMX37FxQZ6Yz9z
# xjzoO2fcmw==
# SIG # End signature block
