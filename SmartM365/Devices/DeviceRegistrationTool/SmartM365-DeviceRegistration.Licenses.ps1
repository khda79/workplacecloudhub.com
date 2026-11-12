# Version: 1.1.0
# Optional delegated Graph check. Dot-sourcing never connects to Microsoft 365.
[CmdletBinding()]
param([switch]$Check, [string]$ExpectedUpn, [string]$ResultPath)

function New-DrtLicenseResult {
    param([string]$Status='Unverifiable', [string]$Reason='Unavailable', [string]$Account='', [string[]]$SkuIds=@())
    [pscustomobject]@{SchemaVersion=1;Status=$Status;Reason=$Reason;Account=$Account;CheckedAtUtc=[DateTime]::UtcNow.ToString('o');SkuIds=@($SkuIds);ReadOnly=$true}
}

function ConvertTo-DrtLicenseResult {
    param([object]$Profile, [string]$ExpectedUpn)
    if ([string]::IsNullOrWhiteSpace($ExpectedUpn) -or [string]$Profile.userPrincipalName -ine $ExpectedUpn) {
        return New-DrtLicenseResult -Reason AccountMismatch
    }
    if (-not $Profile.PSObject.Properties['assignedLicenses'] -or $null -eq $Profile.assignedLicenses -or $Profile.assignedLicenses -isnot [array]) {
        return New-DrtLicenseResult -Reason IncompleteResponse -Account $ExpectedUpn
    }
    $ids = @()
    foreach ($license in @($Profile.assignedLicenses)) {
        $guid = [guid]::Empty
        if (-not [guid]::TryParse([string]$license.skuId,[ref]$guid) -or $guid -eq [guid]::Empty) {
            return New-DrtLicenseResult -Reason IncompleteResponse -Account $ExpectedUpn
        }
        $ids += $guid.ToString()
    }
    $status = if ($ids.Count) {'Assigned'} else {'None'}
    New-DrtLicenseResult -Status $status -Reason VerifiedProfile -Account $ExpectedUpn -SkuIds @($ids | Select-Object -Unique)
}

function Invoke-DrtLicenseCheck {
    param([string]$ExpectedUpn)
    if ($ExpectedUpn -notmatch '^[^\s@]+@[^\s@]+$') {return New-DrtLicenseResult -Reason AccountUnavailable}
    $phase='LoadModule'
    try {
        $local = Join-Path $PSScriptRoot 'Modules\Microsoft.Graph.Authentication\2.41.2\Microsoft.Graph.Authentication.psd1'
        if (Test-Path -LiteralPath $local) {Import-Module $local -ErrorAction Stop}
        else {Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.41.2 -ErrorAction Stop}
    } catch {return New-DrtLicenseFailure -ErrorRecord $_ -Phase $phase -Account $ExpectedUpn}
    try {
        # This function runs in its own short-lived process. Never reuse a cached account.
        try {Disconnect-MgGraph -ErrorAction Stop | Out-Null} catch {}
        $phase='Authentication'; Write-DrtLicenseProgress $phase
        Connect-MgGraph -Scopes @('User.Read') -ContextScope Process -LoginHint $ExpectedUpn -NoWelcome -ClientTimeout 120 -ErrorAction Stop | Out-Null
        $phase='ValidateContext'; Write-DrtLicenseProgress $phase
        $context = Get-MgContext -ErrorAction Stop
        if ($context.AuthType -ne 'Delegated' -or $context.ContextScope -ne 'Process' -or $context.Account -ine $ExpectedUpn) {
            return New-DrtLicenseResult -Reason AccountMismatch -Account $ExpectedUpn
        }
        $phase='ReadProfile'; Write-DrtLicenseProgress $phase
        $profile = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,assignedLicenses' -OutputType PSObject -ErrorAction Stop
        ConvertTo-DrtLicenseResult -Profile $profile -ExpectedUpn $ExpectedUpn
    } catch {
        # SDK exceptions may include tokens, URLs or tenant data. Do not persist their text.
        New-DrtLicenseFailure -ErrorRecord $_ -Phase $phase -Account $ExpectedUpn
    } finally {try {Disconnect-MgGraph -ErrorAction Stop | Out-Null} catch {}}
}

function Write-DrtLicenseProgress {
    param([string]$Phase)
    if ($script:DrtLicenseProgressPath) {
        [IO.File]::WriteAllText($script:DrtLicenseProgressPath,([pscustomobject]@{Phase=$Phase;Utc=[DateTime]::UtcNow.ToString('o')}|ConvertTo-Json),[Text.UTF8Encoding]::new($true))
    }
}

function New-DrtLicenseFailure {
    param($ErrorRecord,[string]$Phase,[string]$Account)
    $exception=$ErrorRecord.Exception
    $message=[string]$exception.Message
    $status=$null
    foreach($candidate in @($exception.ResponseStatusCode,$exception.StatusCode,$exception.Response.StatusCode)){
        $number=0;if([int]::TryParse([string]$candidate,[ref]$number) -and $number -ge 100 -and $number -le 599){$status=$number;break}
    }
    $reason=if($Phase -eq 'LoadModule'){'DependencyUnavailable'}elseif($status -in @(401,403)){'AccessDenied'}elseif($message -match '(?i)timed?\s*out|timeout|inactivity'){'AuthenticationTimeout'}elseif($message -match '(?i)cancel'){ 'AuthenticationCancelled' }elseif($message -match '(?i)window handle'){ 'AuthenticationWindowUnavailable' }else{'AccessOrConnectionUnavailable'}
    $aadsts=[regex]::Match($message,'AADSTS([0-9]{4,10})').Groups[1].Value
    $r=New-DrtLicenseResult -Reason $reason -Account $Account
    # Allowlisted technical facts only; no exception text, headers, URLs or tokens.
    $r|Add-Member NoteProperty Diagnostic ([pscustomobject]@{Phase=$Phase;ExceptionType=$exception.GetType().FullName;HResult=$exception.HResult;HttpStatus=$status;AadstsCode=$aadsts})
    $r
}

function Resolve-DrtLicenseProducts {
    param([string[]]$SkuIds,[string]$CatalogPath=(Join-Path $PSScriptRoot 'SmartM365-DeviceRegistration.LicenseCatalog.json'))
    $map=@{}
    try {
        $catalog=Get-Content -LiteralPath $CatalogPath -Raw -ErrorAction Stop|ConvertFrom-Json -ErrorAction Stop
        if($catalog.SchemaVersion -eq 1){
            foreach($product in $catalog.Products){
                $id=[guid]::Empty
                $names=@($product.Names|Where-Object {$_ -is [string] -and $_.Length -gt 0 -and $_.Length -le 250 -and $_ -notmatch '[\x00-\x1f]'})
                if([guid]::TryParse([string]$product.SkuId,[ref]$id) -and $names.Count){$map[$id.ToString()]=$names -join ' / '}
            }
        }
    }catch{}
    foreach($sku in $SkuIds){
        $key=$sku.ToLowerInvariant();$known=$map.ContainsKey($key)
        [pscustomobject]@{SkuId=$sku;DisplayName=$(if($known){$map[$key]}else{'Unrecognized Microsoft product'});CatalogMatch=$known}
    }
}

function Get-DrtLicenseText {
    param([object]$Result, [string]$Language='en')
    $fr = $Language -like 'fr*'
    if (-not $Result) {if($fr){return 'Non vérifié. Connexion Microsoft facultative.'}else{return 'Not checked. Optional Microsoft sign-in.'}}
    $label = switch ($Result.Status) {
        'Assigned' {if($fr){'Licences Microsoft attribuées : '+@($Result.SkuIds).Count}else{'Assigned Microsoft licenses: '+@($Result.SkuIds).Count}}
        'None' {if($fr){'Aucune licence Microsoft attribuée.'}else{'No Microsoft license assigned.'}}
        default {if($fr){'Non vérifiable. Connexion Microsoft ou vérification incomplète.'}else{'Unverifiable. Microsoft sign-in or the check did not complete.'}}
    }
    $details = if($Result.Account){' ('+$Result.Account+'; '+$Result.CheckedAtUtc+')'}else{''}
    $sku = if(@($Result.SkuIds).Count){
        $products=@(Resolve-DrtLicenseProducts -SkuIds $Result.SkuIds)
        $names=@($products|ForEach-Object {if($fr -and -not $_.CatalogMatch){'Produit Microsoft non identifie'}else{$_.DisplayName}})
        "`r`n"+($names -join '; ')+"`r`nSKU: "+(@($Result.SkuIds) -join ', ')
    }else{''}
    $label+$details+$sku
}

if ($Check) {
    $ErrorActionPreference='Stop'
    $script:DrtLicenseProgressPath=$ResultPath+'.progress.json'
    Write-DrtLicenseProgress 'Started'
    # The GUI supplies a unique path in this user's existing output directory.
    try {$result = Invoke-DrtLicenseCheck -ExpectedUpn $ExpectedUpn} catch {$result = New-DrtLicenseFailure -ErrorRecord $_ -Phase Worker -Account $ExpectedUpn}
    $temp = $ResultPath+'.tmp'
    [IO.File]::WriteAllText($temp,($result|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($true))
    Move-Item -LiteralPath $temp -Destination $ResultPath -Force
    Write-DrtLicenseProgress 'Completed'
}




# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcavXUk6C5KENk
# OLesMe6hzHiYyMSrZVHn/FrBsVuwcqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMWWSbLpkyfLnK5uuTzbDqaIxzWqLbivFYiRXhcCQVYRMA0GCSqG
# SIb3DQEBAQUABIIBgGCepj2dmV50T9Bt/XNMcgEA4nFbf/n9NHHkftvOSReNIVOU
# RPDQVmTSDIPPe4lgorW/yWL1g2vh9UsjUETn/gxp9Frz5RAJN6Ohxx1mDfZximTI
# d1EcyLVKGMvO25adnHv6x8sgZ2xMGpb9rytYZklQqtigdx1t0Ih7TVDhUQxw6QgO
# azmMgFBTZ/j5Z8zhM7cjQ7YGdyUwR9p2BuSb3oRlDFBfNbM4C1bjeNi1DfcYX6/4
# eHR0ZcVEIuRO9ZSXS0vlxrTZ+yKXMP6jNtdOxjllgwNn85e7SJjjk3RVduubCgZA
# ppW3eAlnN2VF0ARTcdeM1O20R8JdA1YcmAEAdv8nE9qW8NunE3YcPtxLU4pe81jG
# aS2RqOzgtXiTtyqP4OfT+zRlbW9Mbe2BayoDy96I0dYfdnwIBq84GcV1gxHeH7UO
# T4Xm36+HHdTtwkWvC2ez7lScVZKurM+QhMbUm/+ycD294Eqh1WuNj4ejyDnz+3bl
# d8unfMtyuB8lOhMWgqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTIyMTIy
# MDBaMC8GCSqGSIb3DQEJBDEiBCAeP96mZjV+Q6zxiuaDVn2iwe4KHSfn3Hp8kWY8
# BuO4tDANBgkqhkiG9w0BAQEFAASCAgBaZtIBaEWuKfpaUSkV5JZYJap4BQMFhITE
# EOtpshy/rXttQAXaMOFRA0g4nhqNCiwOklFb2iI3qgrS4UqdKxISqbzN8Q/GKK/x
# 7t02l9zBApauHCLtfEoMUKtYFgQUeXgOZMWnzTm062PeDMC2p8zyugGMyinyhu5y
# dam9BuL8g2LJ8/OjcdoU9bkjiIN0sggccDfsJcwGr84qdLD+QQo9PJ1tFwh57WoS
# KYzqP3HhwRL46x2ntFke+yGGpl15CgECeN2PTJ49ZnTQ+X5JRWPP1Zv5mdZhNRmK
# vvSEISjAx6WWUr9JzCdIgG5Oo4Rz7tw6Dqm8dG8ziLZmIGx1JUcwWIhT2x9W8raE
# sdBWT9rerQKIc83qcTgFSbBbe2Fwi7meOlYwqwkFtfkLqrcq665Y8UucOQys1VVF
# AzlmOynsEsJDwt9wM8S043iXnMpiweqc6LtoAthDFclBKsYL9gHCJpbqdxEdFh1v
# Wxn8eQf9OtxcL27+XiJ+uqg/lTFkTJkXGEXqc/F2vE00gBeULlQZet3fu0p43WW6
# 38YTo1+NRGWDi9fzu/oHYfyn3qbSR4nkaKZDOTsooAej1L4ig+s8kLZzcktYj0ic
# Idgb3i+gghX5jPr31BW3ykgITU1qndjAaRxJF6FVSmWvXP0YApcZbxJbGqFfEfIr
# Vghk1CaOpg==
# SIG # End signature block
