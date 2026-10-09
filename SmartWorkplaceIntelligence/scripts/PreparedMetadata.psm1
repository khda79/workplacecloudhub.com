Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../../SmartM365/Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Global

function Get-PreparedMetadataPath {
    param([string]$Folder,[ValidateSet('current','batch','validation')][string]$Name,[switch]$Optional)
    $path=Get-SmartM365JsonReadPath (Join-Path $Folder ($Name+'.json')) -Optional:$Optional
    if($path){return (Read-SmartM365JsonDocument $path).Path}
}

function Initialize-PreparedMetadataNames {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputRoot,[Parameter(Mandatory)][string]$TenantKey)
    if((Get-SmartM365JsonTransportPolicy).Mode -ne 'JsonText'){return}
    if(-not(Get-SmartM365JsonReadPath (Join-Path $OutputRoot 'current.json') -Optional)){return}
    # An already converted feed needs no rename operation. Publication validates
    # the current batch separately; do not run the legacy migration on cloud
    # placeholder directories when none of its owned legacy filenames remain.
    $needsConversion=Test-Path -LiteralPath (Join-Path $OutputRoot 'current.json')
    foreach($family in 'batches','retired','failed'){
        $folder=Join-Path $OutputRoot $family
        if(-not(Test-Path -LiteralPath $folder -PathType Container)){continue}
        foreach($entry in Get-ChildItem -LiteralPath $folder -Directory){
            if($entry.Name -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$'){continue}
            foreach($name in 'batch','current','validation','failure'){
                if(Test-Path -LiteralPath (Join-Path $entry.FullName ($name+'.json'))){$needsConversion=$true}
            }
        }
    }
    if($needsConversion){Convert-PreparedMetadataNames -OutputRoot $OutputRoot -TenantKey $TenantKey -Apply | Out-Null}
}

function Assert-PreparedUnlinkedPath([string]$Path){
    $cursor=[IO.Path]::GetFullPath($Path)
    $volumeRoot=[IO.Path]::GetPathRoot($cursor).TrimEnd('\','/')
    while($cursor){
        if((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Metadata conversion refuses linked paths.'}
        # A UNC share is the filesystem root; never probe the server namespace above it.
        if($cursor.TrimEnd('\','/') -eq $volumeRoot){break}
        $parent=Split-Path $cursor -Parent
        if($parent -eq $cursor){break};$cursor=$parent
    }
}

function Convert-PreparedMetadataNames {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$OutputRoot,[Parameter(Mandatory)][string]$TenantKey,[switch]$Apply)
    $qualifiedUnc=$false
    if($Apply -and [IO.Path]::GetFullPath($OutputRoot).StartsWith('\\')){
        $policy=Get-SmartM365JsonTransportPolicy
        $qualifiedUnc=@($policy.QualifiedUncRoots | Where-Object { [IO.Path]::GetFullPath($OutputRoot).StartsWith(([IO.Path]::GetFullPath($_).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
        if(-not $qualifiedUnc){throw 'Prepared metadata UNC migration requires qualification before activation.'}
    }
    $root=(Resolve-Path -LiteralPath $OutputRoot).ProviderPath.TrimEnd('\','/')
    if((Split-Path $root -Leaf) -ne 'DATA-POWERBI'){throw 'Conversion requires a dedicated DATA-POWERBI directory.'}
    Assert-PreparedUnlinkedPath $root
    $lock=[IO.File]::Open((Join-Path $root '.publication.lock'),'OpenOrCreate','ReadWrite','None')
    try {
        $plan=[Collections.Generic.List[object]]::new()
        $checkedCsv=0
        function Plan-Name([string]$folder,[string]$name){
            $path=Get-PreparedMetadataPath $folder $name
            Assert-PreparedUnlinkedPath $path
            $legacy=Join-Path $folder ($name+'.json');$new=Join-Path $folder ($name+'.json.txt')
            if(Test-Path -LiteralPath $legacy){$plan.Add([pscustomobject]@{From=$legacy;To=$new;SHA256=(Get-FileHash -LiteralPath $legacy).Hash})}
            return $path
        }
        $rootPointerPath=Get-PreparedMetadataPath $root 'current'
        Assert-PreparedUnlinkedPath $rootPointerPath
        $pointer=Get-Content -LiteralPath $rootPointerPath -Raw | ConvertFrom-Json
        if($pointer.SchemaVersion -ne 1 -or $pointer.TenantKey -ne $TenantKey){throw 'Root pointer schema/tenant mismatch.'}
        $expected=$pointer;$seen=[Collections.Generic.HashSet[string]]::new()
        $chainSeen=[Collections.Generic.HashSet[string]]::new()
        $pending=[Collections.Generic.Queue[string]]::new()
        foreach($batchFolder in Get-ChildItem -LiteralPath (Join-Path $root 'batches') -Directory){
            if($batchFolder.Name -match '^\d{8}T\d{9}Z-[a-f0-9]{8}$'){$pending.Enqueue($batchFolder.FullName)}
        }
        while($null -ne $expected){
            $id=[string]$expected.BatchId
            if($id -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$' -or -not $seen.Add($id)){throw 'Invalid or cyclic batch ID.'}
            $null=$chainSeen.Add($id)
            $folder=Join-Path $root ('batches/'+$id)
            Assert-PreparedUnlinkedPath $folder
            $manifestPath=Plan-Name $folder 'batch'
            $validationPath=Plan-Name $folder 'validation'
            $receiptPath=Plan-Name $folder 'current'
            $manifest=Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
            $receipt=Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
            $validation=Get-Content -LiteralPath $validationPath -Raw | ConvertFrom-Json
            $manifestHash=(Get-FileHash -LiteralPath $manifestPath).Hash
            if($receipt.SchemaVersion -ne 1 -or $manifest.SchemaVersion -ne 1 -or $receipt.BatchId -ne $id -or $manifest.BatchId -ne $id -or $receipt.TenantKey -ne $TenantKey -or $manifest.TenantKey -ne $TenantKey -or $receipt.ManifestSHA256 -ne $manifestHash -or $expected.ManifestSHA256 -ne $manifestHash -or $receipt.PreviousBatchId -ne $expected.PreviousBatchId){throw 'Batch identity or manifest hash mismatch.'}
            if($validation.Passed -ne $true -or $validation.SchemaOnly -ne $false -or @($manifest.Files).Count -eq 0 -or @($validation.Files).Count -ne @($manifest.Files).Count){throw 'Batch validation receipt is incomplete or failed.'}
            $fileNames=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach($file in $manifest.Files){
                $name=[string]$file.File
                if($name -notmatch '^[A-Za-z0-9_-]+\.csv$' -or -not $fileNames.Add($name)){throw 'Invalid or duplicate CSV filename.'}
                $path=Join-Path $folder $name;Assert-PreparedUnlinkedPath $path
                $validationRows=@($validation.Files | Where-Object File -CEQ $name)
                if($validationRows.Count -ne 1 -or $validationRows[0].SHA256 -ne $file.SHA256 -or $validationRows[0].Rows -ne $file.Rows -or (Get-Item -LiteralPath $path).Length -ne $file.Bytes -or (Get-FileHash -LiteralPath $path).Hash -ne $file.SHA256){throw 'CSV integrity or validation receipt mismatch.'}
                $checkedCsv++
            }
            $expected=$null
            if($receipt.PreviousBatchId){
                $previousId=[string]$receipt.PreviousBatchId
                if($previousId -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$'){throw 'Invalid previous batch ID.'}
                if($chainSeen.Contains($previousId)){throw 'Cyclic previous batch identity.'}
                $previousFolder=Join-Path $root ('batches/'+$previousId)
                # Older retired payloads are not recreated or traversed.
                if(-not $seen.Contains($previousId) -and (Test-Path -LiteralPath $previousFolder)){
                    Assert-PreparedUnlinkedPath $previousFolder
                    $previousReceipt=Get-PreparedMetadataPath $previousFolder 'current'
                    Assert-PreparedUnlinkedPath $previousReceipt
                    $expected=Get-Content -LiteralPath $previousReceipt -Raw | ConvertFrom-Json
                    if($expected.BatchId -ne $previousId){throw 'Previous batch identity mismatch.'}
                }
            }
            while($null -eq $expected -and $pending.Count -gt 0){
                $other=$pending.Dequeue()
                if($seen.Contains((Split-Path $other -Leaf))){continue}
                $otherReceipt=Get-PreparedMetadataPath $other 'current' -Optional
                if(-not $otherReceipt){Write-Warning "Uncommitted prepared batch preserved; owner receipt missing: $other";continue}
                $expected=(Read-SmartM365JsonDocument $otherReceipt).Document
                $chainSeen.Clear()
                if($expected.BatchId -ne (Split-Path $other -Leaf)){throw 'Unreferenced batch ownership mismatch.'}
            }
        }
        foreach($kind in 'retired','failed'){
            $archiveRoot=Join-Path $root $kind
            if(-not(Test-Path -LiteralPath $archiveRoot -PathType Container)){continue}
            foreach($archive in Get-ChildItem -LiteralPath $archiveRoot -Directory){
                if($archive.Name -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$'){continue}
                Assert-PreparedUnlinkedPath $archive.FullName
                if($kind -eq 'retired'){
                    $archivedReceipt=(Read-SmartM365JsonDocument (Plan-Name $archive.FullName 'current')).Document
                    $archivedManifest=Read-SmartM365JsonDocument (Plan-Name $archive.FullName 'batch')
                    $archivedValidation=(Read-SmartM365JsonDocument (Plan-Name $archive.FullName 'validation')).Document
                    if($archivedReceipt.TenantKey -ne $TenantKey -or $archivedManifest.Document.TenantKey -ne $TenantKey -or $archivedReceipt.BatchId -ne $archive.Name -or $archivedManifest.Document.BatchId -ne $archive.Name -or $archivedReceipt.ManifestSHA256 -ne $archivedManifest.SHA256 -or $archivedValidation.Passed -ne $true){throw 'Retired prepared receipts do not establish ownership and integrity.'}
                }else{
                    $failurePath=Get-SmartM365JsonReadPath (Join-Path $archive.FullName 'failure.json') -Optional
                    if(-not $failurePath){Write-Warning "Failed batch without owner receipt preserved: $($archive.FullName)";continue}
                    $failure=(Read-SmartM365JsonDocument $failurePath).Document
                    if($failure.TenantKey -ne $TenantKey -or $failure.BatchId -ne $archive.Name){throw 'Failed prepared receipt ownership mismatch.'}
                    $legacy=Join-Path $archive.FullName 'failure.json'
                    if(Test-Path -LiteralPath $legacy){$plan.Add([pscustomobject]@{From=$legacy;To=$legacy+'.txt';SHA256=(Get-FileHash -LiteralPath $legacy).Hash})}
                    if(Get-PreparedMetadataPath $archive.FullName 'validation' -Optional){$null=Plan-Name $archive.FullName 'validation'}
                }
            }
        }
        $null=Plan-Name $root 'current' # Publish the renamed root pointer LAST.
        foreach($item in $plan){
            if((Read-SmartM365JsonDocument $item.From).SHA256 -ne $item.SHA256){throw 'Metadata changed during conversion planning.'}
        }
        if($Apply){foreach($item in $plan){
            if((Get-FileHash -LiteralPath $item.From).Hash -ne $item.SHA256){throw 'Metadata changed before rename.'}
            $validate={param($document) if($document -isnot [pscustomobject]){throw 'Prepared metadata must be a JSON object.'}}
            # Publication lock spans all receipts; the shared per-file lock and journal
            # make a partly completed chain resumable without recalculation.
            Move-SmartM365OwnedJsonFile -Root (Split-Path $item.From -Parent) -RelativePath (Split-Path $item.From -Leaf) -Owner 'WorkplaceEvidence-Prepare' -Validate $validate -QualifiedUnc:$qualifiedUnc -RemoveIdenticalLegacy | Out-Null
            if((Get-FileHash -LiteralPath $item.To).Hash -ne $item.SHA256){throw 'Renamed metadata hash mismatch.'}
        }}
        [pscustomobject]@{Applied=[bool]$Apply;BatchId=$pointer.BatchId;Batches=$seen.Count;CheckedCsvFiles=$checkedCsv;MetadataFiles=$plan.Count;CsvRecalculated=$false;Files=@($plan)}
    } finally {$lock.Dispose()}
}
function Convert-PreparedAuditNames {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][ValidateSet('transfers','workforce-diagnostics','DATA-REPAIR-BACKUPS')][string]$Family,[string]$TenantKey)
    if((Get-SmartM365JsonTransportPolicy).Mode -ne 'JsonText'){return}
    $folder=Join-Path $Root $Family
    if(-not(Test-Path -LiteralPath $folder -PathType Container)){return}
    Assert-PreparedUnlinkedPath $folder
    $plans=[Collections.Generic.List[object]]::new()
    foreach($entry in Get-ChildItem -LiteralPath $folder -Directory){
        $pattern=if($Family -eq 'transfers'){'^[a-f0-9]{32}$'}else{'^\d{8}T\d{9}Z-[a-f0-9]{8}$'}
        if($entry.Name -notmatch $pattern){continue}
        Assert-PreparedUnlinkedPath $entry.FullName
        $markerName=switch($Family){'transfers'{'transfer'} 'DATA-REPAIR-BACKUPS'{'repair'} default{'environment'}}
        $markerPath=Get-SmartM365JsonReadPath (Join-Path $entry.FullName ($markerName+'.json')) -Optional
        if(-not $markerPath){Write-Warning "Audit folder without owner receipt preserved: $($entry.FullName)";continue}
        $marker=(Read-SmartM365JsonDocument $markerPath).Document
        if($Family -eq 'workforce-diagnostics'){
            if($marker.Publication -ne $false -or $marker.WorkerSHA256 -notmatch '^[A-Fa-f0-9]{64}$' -or $marker.MonitorSHA256 -notmatch '^[A-Fa-f0-9]{64}$'){throw 'Diagnostic audit ownership incomplete.'}
            $names=@('environment','result')
        }else{
            if([string]::IsNullOrWhiteSpace($TenantKey) -or $marker.TenantKey -ne $TenantKey){throw 'Audit tenant ownership mismatch.'}
            if($Family -eq 'transfers' -and $marker.BatchId -notmatch '^\d{8}T\d{9}Z-[a-f0-9]{8}$'){throw 'Invalid transfer batch identity.'}
            if($Family -eq 'DATA-REPAIR-BACKUPS' -and ($marker.SchemaVersion -ne 1 -or -not $marker.PSObject.Properties['Files'])){throw 'Repair audit ownership incomplete.'}
            $names=@($markerName)
        }
        foreach($name in $names){
            $path=Get-SmartM365JsonReadPath (Join-Path $entry.FullName ($name+'.json')) -Optional
            if(-not $path){continue}
            $receipt=Read-SmartM365JsonDocument $path
            if($name -eq 'result' -and ([IO.Path]::GetFullPath([string]$receipt.Document.RunRoot) -ne $entry.FullName -or $receipt.Document.Publication -ne $false)){throw 'Diagnostic result belongs to another run.'}
            $plans.Add([pscustomobject]@{Path=$path;SHA256=$receipt.SHA256})
        }
    }
    foreach($plan in $plans){
        if((Read-SmartM365JsonDocument $plan.Path).SHA256 -ne $plan.SHA256){throw 'Audit changed during migration planning.'}
        Resolve-SmartM365OwnedJsonPath -Path $plan.Path -Owner ('WorkplaceEvidence-Prepare/'+$Family) -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Audit must be an object.'}} | Out-Null
    }
}
Export-ModuleMember -Function Get-PreparedMetadataPath,Convert-PreparedMetadataNames,Initialize-PreparedMetadataNames,Convert-PreparedAuditNames

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAgPM31T0gdyIde
# 6ZDZGYS0BOfoXw20sozxXJfO3vQv2KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILHwg8lOsyCLRya4HS5hIugnreSrJ2YFFlkrbcTYO68/MA0GCSqG
# SIb3DQEBAQUABIIBgEhkXKnhmSQVowgAFXL0NIbO8mfMQ8KVH+3ioSgXu2Coh+Lt
# hR4sQna+IcSZz2ns98hOt8+AZ2ysAbPJTdVGwgJJemcGIWrYcQxqhNz6b+AA+YYF
# G769is1EtIwynblPW6kCQhQcEdBaLFjfx2diDJA9Gnb5Ys6Epb/FWhLjyECLwZHQ
# jwzAhDWdAywMcWDPP5Utte00xWjajJU6Dye6br6p+qC+b5axmn7Zb2noilWG/plF
# B5ZC6RygJb2yDUrNovq3kuv4FJspmzaMj6SuOuFUdXGbbkS6yF0pr0VTGI5QK/TU
# EBJBgLAVnsgVty6AnWZbBTJSbP7btwGnUR3SJLG3Yctq+ifEKf6lHXmhOPfFIoxZ
# iNN0EWek2B/qYmiEA2gWqmTzR8jtopBC2mkO6TuO22UCfsWstYN85LW5ABneSk7J
# dm7eEZuqCrm7HjlYKh7Pjo6G1HLE2sktvCUMTQmxkJXliTKQ8809Y/SD/nj9D/xq
# BLhPutdnZ3ERSLvyKaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDkxNDQx
# MDJaMC8GCSqGSIb3DQEJBDEiBCBu5LCRoksq4N0N9O1bXYvY2Cy6WmE6pt63tN+C
# 60FRfDANBgkqhkiG9w0BAQEFAASCAgBJEJF+QzW5w8FaxcwkUsEkB4371KBuuDZ7
# 6udghfAdhIMFY3qRf0bbg7JPvohCMt5+mzcpuHalr7A3XWWEjg6F9ZkLjLPx4jd2
# SuUdYOZUYlo8UHswqhp3WH+apCx9ib/AyzuUPrtQ4D1t7v3kVfHKak3gaTECJ7H1
# ik7b8IGF7mZABjpcf3O97JPCUlZx5DXfLo/3bgruEAm6Sh9cRxVMMzMNMU42biaa
# XAdCcPnrePINzVU9+BorIukcgW5NeitkOR/hJLDPqYAkVfOu/N1H621mxnJLcHlb
# 1nyfy9ZOUTj27Z3Hi7Kcdwyjh/0nqQbI0WrkAKBmY/wIkLo+ZwSchJa59/mIERaE
# eBoJier+Fo0PN/k7NawuLe9h6D68m5ehx6jarVBOGJCM6wftOUYy8czl30I7VliV
# agI0lUYi3qSydSYupvmyUoweYMQBHvLUYN8Xes3WtCCCf6hwVzFrbwdFhhANAulE
# 8Us5tj37vIXBsQ1aLsOgMsaSnd/r1JMwOzW8KGxjmQLxUCnQV1FNfQQdvZW+ruuo
# NgtguLjNXCuc7hRu0t1qu761qL1Vo4wylxy++DB89ikbcasDoVRZobbxfC4HtifN
# jDZD1vxdq1Ff/y/VoErzsSqekaI8Ag8VokeMyAEeX7zKzCw3FTMvIpxJN1uqfxNA
# QpwYsSR5Dw==
# SIG # End signature block
