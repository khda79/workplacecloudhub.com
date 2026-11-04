<#
.SYNOPSIS
Detailed managed-device hardware acquisition helpers. No execution on import.
.VERSION
1.1.0
#>
function Test-SmartM365DeviceNotFound {
    param([Parameter(Mandatory)]$Record)
    $exception=$Record.Exception
    while($null -ne $exception){
        foreach($name in @('ResponseStatusCode','StatusCode')){
            $property=$exception.PSObject.Properties[$name]
            if($property -and $null -ne $property.Value){
                try { $code=[int]$property.Value; if($code -gt 0){return ($code -eq 404)} } catch { }
            }
        }
        if([string]$exception.Message -match '(?i)Status:\s*404\b|^\s*\[ResourceNotFound\]\s*:'){return $true}
        $exception=$exception.InnerException
    }
    return $false
}

function Confirm-SmartM365MissingDevices {
    param([object[]]$Devices,[hashtable]$Hardware,[string[]]$CandidateIds)
    if(-not $CandidateIds.Count){return ,@($Devices)}
    # A full, successful enumeration is required even for preview runs.
    $current=@(Get-MgDeviceManagementManagedDevice -All -ErrorAction Stop)
    $currentIds=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($item in $current){
        $currentId=[string](Get-SafeProperty $item 'Id')
        if([string]::IsNullOrWhiteSpace($currentId) -or -not $currentIds.Add($currentId)){throw 'Invalid identity in full managed-device verification.'}
    }
    $excluded=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach($id in $CandidateIds){
        if($currentIds.Contains($id)){throw 'A device with a 404 remains in the full managed-device list.'}
        try {
            $null=Get-MgDeviceManagementManagedDevice -ManagedDeviceId $id -ErrorAction Stop
            throw 'A device with a 404 became visible during final verification.'
        } catch {
            if(-not (Test-SmartM365DeviceNotFound $_)){throw}
        }
        [void]$excluded.Add($id)
    }
    foreach($id in $excluded){
        [void]$Hardware.Remove($id)
        WriteLog -Message ("Managed device absent during collection: $id; full list and individual 404 confirmed at $([datetimeoffset]::UtcNow.ToString('o')).") -Level WARNING
    }
    return ,@($Devices | Where-Object {-not $excluded.Contains([string](Get-SafeProperty $_ 'Id'))})
}

function ConvertTo-SmartM365HardwareRecord {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$ManagedDeviceId, [Parameter(Mandatory)]$Device)
    $actualId = [string](Get-SafeProperty $Device 'id')
    if (-not $actualId -or $actualId -ine $ManagedDeviceId) { throw 'Hardware response identity mismatch.' }
    $result = [ordered]@{ManagedDeviceId=$ManagedDeviceId;CollectionStatus='Collected';CollectedAtUtc=[datetimeoffset]::UtcNow.ToString('o')}
    foreach ($field in @('azureADDeviceId','serialNumber','manufacturer','model')) {
        $result[$field] = [string](Get-SafeProperty $Device $field)
    }
    foreach ($field in @('totalStorageSpaceInBytes','freeStorageSpaceInBytes','physicalMemoryInBytes')) {
        $raw = Get-SafeProperty $Device $field
        $parsed = 0L
        if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) {
            $result[$field] = ''; $result[$field+'Status']='Missing'
        } elseif (-not [int64]::TryParse([string]$raw,[ref]$parsed) -or $parsed -lt 0) {
            throw ('Invalid hardware byte value: '+$field)
        } else {
            $result[$field]=$parsed
            $result[$field+'Status']=if($parsed -eq 0){'ZeroReported'}else{'Reported'}
        }
    }
    if ($result.totalStorageSpaceInBytesStatus -eq 'Reported' -and $result.freeStorageSpaceInBytesStatus -eq 'Reported' -and
        $result.freeStorageSpaceInBytes -gt $result.totalStorageSpaceInBytes) { throw 'Free storage exceeds total storage.' }
    return [pscustomobject]$result
}

function Get-SmartM365ManagedDeviceHardware {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([AllowEmptyCollection()][string[]]$ManagedDeviceIds, [ValidateRange(1,10)][int]$MaxAttempts=6)
    $ids=@($ManagedDeviceIds | Sort-Object -Unique)
    if (@($ids | Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count) { throw 'Missing managed-device identity.' }
    $result=@{}
    $select='id,azureADDeviceId,serialNumber,manufacturer,model,totalStorageSpaceInBytes,freeStorageSpaceInBytes,physicalMemoryInBytes'
    for($offset=0;$offset -lt $ids.Count;$offset+=20){
        $pending=@($ids[$offset..([math]::Min($offset+19,$ids.Count-1))])
        for($attempt=1;$pending.Count -and $attempt -le $MaxAttempts;$attempt++){
            $batchPending=@($pending); $requestIndex=@{}; $requests=@()
            foreach($id in $pending){
                $rid=[string]($requests.Count+1);$requestIndex[$rid]=$id
                $requests+=@{id=$rid;method='GET';url=('/deviceManagement/managedDevices/'+[uri]::EscapeDataString($id)+'?$select='+$select)}
            }
            $retry=[Collections.Generic.List[string]]::new();$notFound=[Collections.Generic.List[string]]::new();$waitSeconds=5*$attempt
            try {
                $response=Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/$batch' -Body (@{requests=$requests}|ConvertTo-Json -Depth 6) -ContentType 'application/json' -ErrorAction Stop
                $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
                foreach($part in @(Get-SafeProperty $response 'responses')){
                    $rid=[string](Get-SafeProperty $part 'id')
                    if(-not $requestIndex.ContainsKey($rid) -or -not $seen.Add($rid)){throw 'Unmapped or duplicate hardware batch response.'}
                    $id=$requestIndex[$rid];$status=[int](Get-SafeProperty $part 'status')
                    if($status -eq 200){
                        try {$result[$id]=ConvertTo-SmartM365HardwareRecord -ManagedDeviceId $id -Device (Get-SafeProperty $part 'body')}
                        catch {$retry.Add($id);WriteLog -Message ('Hardware response rejected: '+$_.Exception.Message) 'WARNING'}
                    } elseif($status -eq 404){
                        $notFound.Add($id)
                    } elseif($status -in @(408,429) -or $status -ge 500){
                        $retry.Add($id)
                        $headers=Get-SafeProperty $part 'headers';$delay=0
                        if([int]::TryParse([string](Get-SafeProperty $headers 'Retry-After'),[ref]$delay) -and $delay -gt $waitSeconds){$waitSeconds=$delay}
                    } else { $retry.Add($id);WriteLog -Message ("Hardware request HTTP $status; unit SDK fallback will verify the record.") 'WARNING' }
                }
                foreach($rid in $requestIndex.Keys){if(-not $seen.Contains($rid)){$retry.Add($requestIndex[$rid])}}
                $pending=@($retry | Sort-Object -Unique)
                foreach($id in $notFound){
                    try {
                        $device=Get-MgDeviceManagementManagedDevice -ManagedDeviceId $id -Property $select -ErrorAction Stop
                        $result[$id]=ConvertTo-SmartM365HardwareRecord -ManagedDeviceId $id -Device $device
                    } catch {
                        if(Test-SmartM365DeviceNotFound $_){$result[$id]=[pscustomobject]@{ManagedDeviceId=$id;CollectionStatus='NotFound';CollectedAtUtc=[datetimeoffset]::UtcNow.ToString('o')}}
                        else {$retry.Add($id)}
                    }
                }
                $pending=@($retry | Sort-Object -Unique)
            } catch {
                WriteLog -Message ('Hardware batch could not be qualified: '+$_.Exception.Message) 'WARNING'
                # Reject the entire unqualified response envelope, not just its tail.
                foreach($id in $batchPending){[void]$result.Remove($id)}
                $pending=$batchPending
            }
            if($pending.Count -and $attempt -lt $MaxAttempts){Start-Sleep -Seconds $waitSeconds}
        }
        foreach($id in $pending){
            try {
                $device=Get-MgDeviceManagementManagedDevice -ManagedDeviceId $id -Property $select -ErrorAction Stop
                $result[$id]=ConvertTo-SmartM365HardwareRecord -ManagedDeviceId $id -Device $device
            } catch {
                $status=if(Test-SmartM365DeviceNotFound $_){'NotFound'}else{'Failed'}
                $result[$id]=[pscustomobject]@{ManagedDeviceId=$id;CollectionStatus=$status;CollectedAtUtc=[datetimeoffset]::UtcNow.ToString('o')}
                WriteLog -Message ("Detailed hardware collection failed for native device $id : "+$_.Exception.Message) 'WARNING'
            }
        }
        WriteLog -Message ('Detailed hardware progress: {0}/{1}' -f ([math]::Min($offset+20,$ids.Count)),$ids.Count) 'INFO'
    }
    return $result
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBAzQRiLfFomJZh
# MGCPJKflo3JnQ+/6ZB/1NbPNkpbwxqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIASoimEQ0otQ/7jp1llTSWhyJHX0CJ55YL9EElXjSoaaMA0GCSqG
# SIb3DQEBAQUABIIBgAZ+pk9qTda6SEtPUvMHwplFVk/qn0BqXzm546qnpF1mc8dB
# tiJmCSG51zkfPIieXnM02RHTEf9PcHRI3HmY+tUYUJNkT3K9GB57ZFZuesUe37JU
# qSJAoefhrNx0Qe9BgzxSEv72Cui9fJA+A7iSjCMc7zMchrtxfiw67P/17q3xFwVH
# BwlWuPnXk3WaF3e7Ud/D/M6kLTCkP0Yx6++fqvOYXslh/I6pBxgp2deAhOdVjGuo
# 7sQuC4sJ0qzE7pMdOvOUO2hSH2zirIV6c2/IISGrPCgU6Nc6Af92bDyAntcKTsZK
# Z/y3hqJph255nf+PUZGditTlTseOYPm6MTNhJFqpwqaEgFbUtxFFhA8ofw8jTnl/
# b5/ScBWnoUpVN286vtghrIM9fZzDxh0qAcAGIs2J3SfiaZjiMd9Mjf4MlalZ/zm5
# cwh+vqXnQgQZvF3kyYeKqZgV4c5KEOyOdkq/fEjig/3WPEdbdCAXGcL9dElPqIxQ
# K72LmPAjt6OTxKlbtKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDQxNTIy
# NDVaMC8GCSqGSIb3DQEJBDEiBCCFqzQRL+ZPgF1j4+qsujpogLbIPkonitWU1vKS
# wZdolTANBgkqhkiG9w0BAQEFAASCAgAjpiasokaBVHW+RKPjKHQQbstOtxHmxzJM
# 82VaAba6l7zXQwgefR1k+srXZQSHQaS9bZ7DOqgBVQEbkBNJDUfx5gPjTypzxOSD
# KPsQ80EANLO0lmFii0ItjXEzS3USUX0LdnQ3hXYK7HC1nURNWxgH9yfxMsuzmcMl
# jCpuC9UhOcd/ZvJx1mo82iCRCkfA4PFF3orBI1/sBp7sEtjmqm9yd1wPGlOtB+hI
# Y8Ch5Bsw4EPsr88jujTet4R5JpBk6s2Wo4O5V6SGdFWmquA+w6yTWprWf/sBu2rC
# S3jGVaEGd0AYtC8D8Hltjp9NMFA2RPrEN/7N63DdMqc2bg90t5aYlwk9HHGwxxR3
# gM/6fuN3pwcMozJsFvm4TXF23CN8sLgmfz/qS4T+3SfvuP2Lb8NDuBkztzC8kpBT
# 8XdX7q0dg3ip67eTPkXVJi348DdFP2yWs63lQRSt88txplIswqO9xGWakMtPUb1u
# 8duSj4zqzwoopKOsvwbOvjjyWEGrr92JW0ZKomK9nD9Y+nm+631bNnD4l6YYbR5J
# rMKDGq8tOO3HXnYSmZ+uKmDDdszhesXdFpw8Me+cmTogr3Fc9vEjt+gSaoz2w22d
# rkyK19e5R5Jl/5ENJycoRYKRu+IGgudP9aO56LLY2tZUUIJFbQ/jn9TvmfSOtFcI
# ZpDRduwoBw==
# SIG # End signature block
