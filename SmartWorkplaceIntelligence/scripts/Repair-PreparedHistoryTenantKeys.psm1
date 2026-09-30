# One-time, explicitly scoped historical source repair. No network APIs or deletions.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'PreparedMetadata.psm1') -Force
function Write-PreparedRepairAudit([string]$Path,$Document){
    Write-SmartM365JsonBytesAtomically -Path $Path -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($Document | ConvertTo-Json -Depth 8))) -Validate {param($value) if($value.SchemaVersion -ne 1 -or -not $value.PSObject.Properties['Files']){throw 'Invalid repair audit.'}} | Out-Null
}
function Get-RepairChildPath([string]$Root,[string]$Relative) {
    $base=[IO.Path]::GetFullPath($Root).TrimEnd('\','/')
    $path=[IO.Path]::GetFullPath((Join-Path $base $Relative))
    if(-not $path.StartsWith($base+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Repair path escapes its data root.'}
    $path
}
function New-HistoryCsvParser([string]$Path) {
    $parser=[Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path,[Text.UTF8Encoding]::new($false,$true),$true)
    $parser.SetDelimiters(',');$parser.HasFieldsEnclosedInQuotes=$true;$parser.TrimWhiteSpace=$false
    $parser
}
function Read-HistoryCsvIdentity([string]$Path,[string]$TenantKey) {
    $parser=New-HistoryCsvParser $Path
    try {
        $header=$parser.ReadFields()
        if(-not $header -or @($header | Sort-Object -Unique).Count -ne $header.Count -or @($header | Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count){throw 'Empty or duplicate CSV header.'}
        $index=[array]::IndexOf($header,'TenantKey');$rows=0L
        while(-not $parser.EndOfData) {
            $values=$parser.ReadFields();$rows++
            if($values.Count -ne $header.Count){throw "Malformed record $rows."}
            if($index -ge 0 -and ([string]::IsNullOrWhiteSpace($values[$index]) -or $values[$index] -cne $TenantKey)){throw "Empty or incompatible TenantKey at record $rows."}
        }
        [pscustomobject]@{HasTenantKey=($index -ge 0);Rows=$rows;Columns=$header.Count}
    } finally {$parser.Dispose()}
}
function Add-HistoryCsvTenantKey([string]$Original,[string]$Destination,[string]$TenantKey) {
    $parser=New-HistoryCsvParser $Original
    $writer=[IO.StreamWriter]::new($Destination,$false,[Text.UTF8Encoding]::new($true))
    $writer.NewLine="`r`n"
    try {
        $header=$parser.ReadFields()
        $writer.WriteLine('"TenantKey",'+(($header | ForEach-Object {'"'+$_.Replace('"','""')+'"'}) -join ','))
        while(-not $parser.EndOfData) {
            $values=$parser.ReadFields()
            if($values.Count -ne $header.Count){throw 'CSV changed or malformed during conversion.'}
            $writer.WriteLine('"'+$TenantKey+'",'+(($values | ForEach-Object {'"'+$_.Replace('"','""')+'"'}) -join ','))
        }
    } finally {$writer.Dispose();$parser.Dispose()}
    # Compare all decoded values, in order, including headers, whitespace and multiline fields.
    $left=New-HistoryCsvParser $Original;$right=New-HistoryCsvParser $Destination
    $records=0L
    try {
        while(-not $left.EndOfData) {
            if($right.EndOfData){throw 'Repaired file lost records.'}
            $a=$left.ReadFields();$b=$right.ReadFields()
            $expected=if($records -eq 0){'TenantKey'}else{$TenantKey}
            if($b.Count -ne ($a.Count+1) -or $b[0] -cne $expected){throw 'Repaired tenant field/shape mismatch.'}
            for($i=0;$i -lt $a.Count;$i++){if(-not [string]::Equals($a[$i],$b[$i+1],[StringComparison]::Ordinal)){throw "Existing value changed at record $records, column $i."}}
            $records++
        }
        if(-not $right.EndOfData){throw 'Repaired file gained records.'}
    } finally {$left.Dispose();$right.Dispose()}
    if($records -lt 1){throw 'Missing CSV header.'}
    $records-1
}
function Invoke-PreparedHistoryTenantRepair {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DataRoot,
        [Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_.-]+$')][string]$TenantKey,
        [Parameter(Mandatory)][string[]]$Weeks,
        [Parameter(Mandatory)][ValidateRange(1,10000)][int]$ExpectedFileCount,
        [string]$SourceContractPath=(Join-Path (Split-Path $PSScriptRoot -Parent) 'config/prepared-source-contract.json'),
        [switch]$Apply)
    if(-not $Weeks.Count -or @($Weeks | Where-Object {$_ -notmatch '^\d{4}-W(0[1-9]|[1-4][0-9]|5[0-3])$'}).Count){throw 'Explicit valid YYYY-Www weeks are required.'}
    $root=(Resolve-Path -LiteralPath $DataRoot).ProviderPath
    $contract=(Read-SmartM365JsonDocument $SourceContractPath).Document
    $lock=[IO.File]::Open((Join-Path $root '.prepared-source.lock'),'OpenOrCreate','ReadWrite','None')
    $manifest=$null;$auditPath=$null
    try {
        if($Apply){Convert-PreparedAuditNames -Root $root -Family DATA-REPAIR-BACKUPS -TenantKey $TenantKey}
        $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $candidates=[Collections.Generic.List[object]]::new()
        foreach($family in $contract.history) {
            if($family.root.Replace('\','/') -notlike 'DATA-ALL/*' -or [IO.Path]::GetFileName($family.file) -ne $family.file){throw 'Invalid historical source contract.'}
            $folder=Get-RepairChildPath $root $family.root
            foreach($file in Get-ChildItem -LiteralPath $folder -File -Filter $family.file -Recurse) {
                $relative=[IO.Path]::GetRelativePath($root,$file.FullName).Replace('\','/')
                if($relative -notmatch '/WeeklyHistory/(\d{4}-W\d{2})/[^/]+\.csv$' -or $Matches[1] -notin $Weeks -or -not $seen.Add($relative)){continue}
                # Reject junctions/symlinks rather than reaching outside the selected tree.
                $cursor=$file
                while($cursor.FullName.TrimEnd('\') -ne $root.TrimEnd('\')) {
                    if($cursor.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "Linked repair path is not supported: $relative"}
                    $cursor=Get-Item -LiteralPath (Split-Path $cursor.FullName -Parent)
                }
                try {$state=Read-HistoryCsvIdentity $file.FullName $TenantKey} catch {throw "${relative}: $($_.Exception.Message)"}
                if(-not $state.HasTenantKey) {
                    $candidates.Add([pscustomobject]@{Relative=$relative;Source=$file.FullName;Rows=$state.Rows;OriginalSHA256=(Get-FileHash -LiteralPath $file.FullName).Hash;OriginalLastWriteUtc=$file.LastWriteTimeUtc.ToString('O');RepairedSHA256='';Status='Planned'})
                }
            }
        }
        foreach($item in $candidates){Write-Host ('[{0:yyyy-MM-dd HH:mm:ss}] TenantKey repair candidate: {1}; rows={2}' -f (Get-Date),$item.Relative,$item.Rows)}
        if($seen.Count -eq 0){throw 'No contract-selected historical files found in the requested weeks. No CSV modified.'}
        if($candidates.Count -eq 0) {return [pscustomobject]@{Status='NoMissingTenantKey';Files=0;Applied=$false;BackupPath=''}}
        if($candidates.Count -ne $ExpectedFileCount){throw "Repair count differs: expected $ExpectedFileCount, found $($candidates.Count). No CSV modified."}
        if(-not $Apply){return [pscustomobject]@{Status='PreviewOnly';Files=$candidates.Count;Applied=$false;BackupPath=''}}
        $backupRoot=Get-RepairChildPath $root ('DATA-REPAIR-BACKUPS/'+[datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8))
        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
        $auditPath=Resolve-SmartM365OwnedJsonPath -Path (Join-Path $backupRoot 'repair.json') -Owner 'WorkplaceEvidence-Prepare/repair' -Validate {param($document) if($document.TenantKey -ne $TenantKey){throw 'Repair audit tenant mismatch.'}}
        $manifest=@{SchemaVersion=1;TenantKey=$TenantKey;Weeks=$Weeks;Operator=[Environment]::UserName;StartedUtc=[datetime]::UtcNow.ToString('O');Status='Preparing';Files=@($candidates.ToArray())}
        Write-PreparedRepairAudit -Path $auditPath -Document $manifest
        # Stage and validate ALL files before replacing any source.
        foreach($item in $candidates) {
            $original=Get-RepairChildPath $backupRoot ('originals/'+$item.Relative)
            $repaired=Get-RepairChildPath $backupRoot ('repaired/'+$item.Relative)
            foreach($p in $original,$repaired){New-Item -ItemType Directory -Path (Split-Path $p -Parent) -Force | Out-Null}
            Copy-Item -LiteralPath $item.Source -Destination $original
            if((Get-FileHash $original).Hash -ne $item.OriginalSHA256){throw "Source changed before backup: $($item.Relative)"}
            $rows=Add-HistoryCsvTenantKey $original $repaired $TenantKey
            if($rows -ne $item.Rows){throw 'Row count changed during repair.'}
            $item.RepairedSHA256=(Get-FileHash $repaired).Hash;$item.Status='Verified'
        }
        foreach($item in $candidates){if((Get-FileHash $item.Source).Hash -ne $item.OriginalSHA256){throw "Source changed before replacement: $($item.Relative)"}}
        $manifest.Status='Replacing';Write-PreparedRepairAudit -Path $auditPath -Document $manifest
        foreach($item in $candidates) {
            $repaired=Get-RepairChildPath $backupRoot ('repaired/'+$item.Relative)
            $replaced=Get-RepairChildPath $backupRoot ('replaced/'+$item.Relative)
            New-Item -ItemType Directory -Path (Split-Path $replaced -Parent) -Force | Out-Null
            if((Get-FileHash $item.Source).Hash -ne $item.OriginalSHA256){throw "Source changed during replacement: $($item.Relative)"}
            # Atomic per-file replacement; original bytes remain in the verified backup.
            [IO.File]::Replace($repaired,$item.Source,$replaced)
            if((Get-FileHash $replaced).Hash -ne $item.OriginalSHA256){throw 'Concurrent source change detected at replacement; replaced bytes retained in backup.'}
            [IO.File]::SetLastWriteTimeUtc($item.Source,[datetime]::Parse($item.OriginalLastWriteUtc).ToUniversalTime())
            if((Get-FileHash $item.Source).Hash -ne $item.RepairedSHA256){throw 'Repaired source hash mismatch.'}
            $item.Status='Repaired'
            Write-PreparedRepairAudit -Path $auditPath -Document $manifest
        }
        $manifest.Status='Completed';$manifest.CompletedUtc=[datetime]::UtcNow.ToString('O');Write-PreparedRepairAudit -Path $auditPath -Document $manifest
        [pscustomobject]@{Status='Completed';Files=$candidates.Count;Applied=$true;BackupPath=$backupRoot}
    } catch {
        if($manifest -and $auditPath){$manifest.Status='Failed';$manifest.Error=$_.Exception.Message;Write-PreparedRepairAudit -Path $auditPath -Document $manifest}
        throw
    } finally {$lock.Dispose()}
}
Export-ModuleMember -Function Invoke-PreparedHistoryTenantRepair

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAK5cPmd84CYp8B
# 6jzdmx5TfkuP3kT5QQIDAz+C4ViUNqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINRdJrtjdikPFb4CV4VixWnA7bD/9Oq+pJ8byeneAUasMA0GCSqG
# SIb3DQEBAQUABIIBgK4L+R7CGmfLpdW6k4+dhHzgEA+DY/X3tUmroj6n+IUbbOoC
# t9Y4b3qDN/CpI01odVOkzM0KyCSbgMKuA0gjxbPeW3XTnMsQMWNcO8Zm7wDeRhf/
# hrWsxYGcrdNh8HT79Q5EB0tZoAtXpYN9RYsKyL2DxRtvU2A1cW2YrdajLYGDkRex
# d7mWGJCMPZRWP3d1DCELlSNY5/RCVxk5QA+Pf2OgmdlZEImIXaisYqAyVtM/sqKd
# HDzNkjDENxCrl8HTuTZiKx5CJBVg7n586upre8KejrQqUXXUe5UjyAd5OH2Vgu/t
# vooIrH+LEV4+vsVDGnxLFbolh437X5WHqpKN60yURfBGxXkYGmOXVDqZDr7mZaTV
# 8xF7N2xpIb74U0kFCNQ6U+r2u4CTMO2wMgUGTL6ubSLw2WUfIts+ZcjvbtLvZcVl
# EicmLbM2z+1LIFQ4v3swmtWUe7boEfngp095nxDftihDxXBDMiU4btbjxfmE3pVc
# byf/VxNJ1DpSqJTmGKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcxNjU4
# MDBaMC8GCSqGSIb3DQEJBDEiBCD4VhBNU0ERbNTYFZqTtRlA/LMx8M6+qcj1Wi8k
# 9BMArTANBgkqhkiG9w0BAQEFAASCAgATvcfuCTwqyKH3ywi3ZCJJLAk78VnqYFQ+
# WNIPhM5JSWtD7eT/HUN/HmgdgoYvrj0xoE0lv5LuAfkxUM8WSLxWStkkM4NKSctz
# LKcZd/+UtxnTzjLRv60tJiprNa09Q34sGtaht7wZBhziTFhiQFLfdqr6r8yLoI10
# rXPNS4jGN87d5wfwvE8TYnGQxkZ+PZuAzFmVqLqQI+h5Op0l5b1tXiFvemASNaBy
# iLECkNGFexHLrn8PMIS0Zr8NkFoXpNMmVtLnjR82CJpTxNjJam93Ed5iDaJLofWo
# 02ra7uongLE10nLPvh/WasB1t6XO7Eb5o3s0xY+u9GuNKbysE0zAUoFzwynkdhQ5
# uDZ2W3uCmewkBh2sUz7Px2np8ycN0hsEycokw35EW5TiTDTyqMjILmZL2yLpdh94
# 9xaPzB5TtFEzijzTAyWFOKkzQODBH4s48ePFcgJ//BT1ncgLB4nHonMjBsanhLlv
# ak6bZ4LKW28g2fTrJS42nAhgWLYyWw+65qFzMUZ8ojrQvykSPecEayFJy4je1YN7
# YhNQfgIl6jePWQlijNEFlUtzgJwgBJ6GV26qfhGNs7Ma6syxmGNJkE/VtDVSbKeq
# nNkBF80RjA1SWn/SUkrva94lzQ53iuYFp/yMg/hF3h7eIMY4wwcs7ytyolTdQ+JW
# AMBgBMUt5w==
# SIG # End signature block
