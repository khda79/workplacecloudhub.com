#Requires -Version 5.1
<#
.SYNOPSIS
Offline AD coverage receipt and failed-domain fragment tests. No tenant calls.
.VERSION
1.0.1
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic receipt globals are saved and restored.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$helper=Join-Path $root 'Modules/SmartM365.Core/SmartM365-CmdbReceipt.ps1'
$registry=Get-Content (Join-Path (Split-Path $helper) 'SmartM365-CmdbSources.json.txt') -Raw | ConvertFrom-Json
$producer=@($registry.Producers | Where-Object Script -eq 'SmartM365-ActiveDirectory-Inventory.ps1')[0]
$temporary=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-AdCoverage-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
$saved=@{}
foreach($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')){
    $v=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name]=@{Exists=($null -ne $v);Value=$(if($v){$v.Value}else{$null})}
}
$global:SmartM365TenantKey='synthetic';$global:SmartM365OrganizationKey='test'
$global:SmartM365EnvironmentKey='test';$global:SmartM365TenantId='synthetic-tenant'
$mock=New-Module -Name SyntheticAdCoverage -ScriptBlock {
    param($path)
    . $path
    function Test-SmartM365MaxItemsMode {return $false}
    function Get-SmartM365ScriptVersionFromFile {param($Path) $null=$Path;return 'synthetic'}
    function WriteLog {param($Message,$Level) $null=$Message;$null=$Level}
    Export-ModuleMember -Function *
} -ArgumentList $helper
$script:checks=0
$script:InvalidScopeStayedQualified=0
$merge=$null
function Check {param([bool]$Condition,[string]$Message) if(-not $Condition){throw $Message};$script:checks++}
function Coverage {return @{ExpectedDomains=@('good.synthetic.invalid','missing.synthetic.invalid');CollectedDomains=@('good.synthetic.invalid');UnavailableDomains=@('missing.synthetic.invalid');NonBlockingDomainErrors=@('missing.synthetic.invalid')}}
function Run-Receipt {
    param([hashtable]$Coverage,[int]$Errors=0,[string]$Status='CompletedWithWarnings',[switch]$MissingFile,[switch]$OtherProducer,[switch]$ClaimFull)
    $directory=Join-Path $temporary ([guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $directory | Out-Null
    $selected=if($OtherProducer){$registry.Producers[0]}else{$producer}
    $global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach($file in $selected.Files){
        if($MissingFile -and $file -eq $selected.Files[0]){continue}
        $path=Join-Path $directory $file
        [IO.File]::WriteAllText($path,"TenantKey,DomainName`r`nsynthetic,good.synthetic.invalid`r`n",[Text.UTF8Encoding]::new($false))
        [void]$global:csvGeneratedPaths.Add($path)
    }
    & $mock {param($scriptPath,$folder) Start-SmartM365CmdbSourceReceipt -ScriptPath $scriptPath -SourceRootPath $folder} $selected.Script $directory
    try {
        & $mock {param($scope,$coverage,$full) Set-SmartM365CmdbSourceScope -CompleteScope $full -Scope $scope -DomainCoverage $coverage} $selected.Scope $Coverage ([bool]$ClaimFull)
    } catch {
        if(& $mock {$script:SmartM365CmdbSourceContext.ScopeQualified}){$script:InvalidScopeStayedQualified++}
        & $mock {Complete-SmartM365CmdbSourceReceipt -Status Failed -ErrorCount 1 | Out-Null}
        throw
    }
    $result=& $mock {param($status,$errors) Complete-SmartM365CmdbSourceReceipt -Status $status -ErrorCount $errors} $Status $Errors
    return (Get-Content $result -Raw | ConvertFrom-Json)
}
try {
    $rejected=$false
    try{$null=Run-Receipt -Coverage (Coverage) -ClaimFull}catch{$rejected=$true}
    Check $rejected 'Partial domain metadata was accepted with a full-coverage claim.'
    $receipt=Run-Receipt -Coverage (Coverage)
    Check ($receipt.Status -eq 'Completed' -and $receipt.IsPartialInventory -and $receipt.ConsumerScopeQualified -and -not $receipt.FullInventoryQualified) 'Accepted partial coverage was represented as full coverage.'
    Check ($receipt.Errors -eq 0 -and $receipt.Files.Count -eq 6) 'Accepted partial receipt lost current required exports.'
    foreach($file in $receipt.Files){Check ($file.IsPartialInventory -and $file.DomainCoverage.Status -eq 'PartialAccepted') 'File coverage not propagated.'}
    foreach($mode in @('Errors','Failed','MissingFile')){
        $args=@{Coverage=(Coverage)}
        if($mode -eq 'Errors'){$args.Errors=1}elseif($mode -eq 'Failed'){$args.Status='Failed'}else{$args.MissingFile=$true}
        $receipt=Run-Receipt @args
        Check ($receipt.Status -eq 'Failed' -and -not $receipt.ConsumerScopeQualified -and $receipt.Files.Count -eq 0) 'A real failure or missing export was accepted.'
    }
    foreach($case in @('Untolerated','AllMissing','Overlap','Unaccounted','Duplicate','Blank','OtherProducer')){
        $coverage=Coverage
        switch($case){
            'Untolerated' {$coverage.NonBlockingDomainErrors=@('different.synthetic.invalid')}
            'AllMissing' {$coverage.CollectedDomains=@()}
            'Overlap' {$coverage.CollectedDomains+=@('missing.synthetic.invalid')}
            'Unaccounted' {$coverage.ExpectedDomains+=@('third.synthetic.invalid')}
            'Duplicate' {$coverage.CollectedDomains+=@('GOOD.synthetic.invalid')}
            'Blank' {$coverage.UnavailableDomains=@('')}
        }
        $rejected=$false
        try{$null=Run-Receipt -Coverage $coverage -OtherProducer:($case -eq 'OtherProducer')}catch{$rejected=$true}
        Check $rejected "Invalid partial case accepted: $case"
    }
    Check ($script:InvalidScopeStayedQualified -eq 0) 'Invalid scope left the receipt qualified.'
    # Execute the actual combine function against current-run fragments only.
    $adPath=Join-Path $root 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1'
    $tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($adPath,[ref]$tokens,[ref]$errors)
    Check ($errors.Count -eq 0) 'AD script parser failed.'
    $combine=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Combine-CsvFiles'},$true)
    $merge=New-Module -Name SyntheticAdMerge -ScriptBlock {
        param($definition)
        . ([scriptblock]::Create($definition))
        $script:FailedInventoryDomains=@('missing.synthetic.invalid')
        $script:DomainsToProcess=@('good.synthetic.invalid','missing.synthetic.invalid')
        function WriteLog {param($Message,$Level) $null=$Message;$null=$Level}
        function Invoke-SmartM365AdCsvReadWithRetry {param($Path,$ReadAction) & $ReadAction}
        function Add-SmartM365TenantKey {process {$_}}
        function Assert-SmartM365CsvDataCompleteness {param($Data,$TimestampedPath,$LatestPath) $null=$Data;$null=$TimestampedPath;$null=$LatestPath}
        function Copy-SmartM365AdFileWithRetry {param($SourcePath,$DestinationPath) Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath}
        function Add-SmartM365AdGeneratedCsvPath {param($Path) $null=$Path}
        function Remove-SmartM365AdFileWithRetry {param($Path) if(Test-Path -LiteralPath $Path){Remove-Item -LiteralPath $Path}}
        Export-ModuleMember -Function Combine-CsvFiles
    } -ArgumentList $combine.Extent.Text
    $fragments=Join-Path $temporary 'fragments';New-Item -ItemType Directory -Path $fragments | Out-Null
    foreach($domain in @('good.synthetic.invalid','missing.synthetic.invalid','old.synthetic.invalid')){
        [IO.File]::WriteAllText((Join-Path $fragments "AD_Users_$domain.csv"),"TenantKey,DomainName`r`nsynthetic,$domain`r`n",[Text.UTF8Encoding]::new($false))
    }
    $combined=Join-Path $temporary 'combined.csv'
    & $merge {param($source,$destination) Combine-CsvFiles -SourceFolder $source -Filter 'AD_Users_*.csv' -DestinationFile $destination} $fragments $combined
    $rows=@(Import-Csv -LiteralPath $combined)
    Check ($rows.Count -eq 1 -and $rows[0].DomainName -eq 'good.synthetic.invalid') 'Failed-domain or previous-cohort fragments leaked into the combined export.'
    Write-Output "PASS: $script:checks offline AD partial-coverage checks. No collectors or live writes."
} finally {
    if($merge){Remove-Module $merge -Force -ErrorAction SilentlyContinue}
    Remove-Module $mock -Force -ErrorAction SilentlyContinue
    foreach($name in $saved.Keys){if($saved[$name].Exists){Set-Variable -Name $name -Scope Global -Value $saved[$name].Value}else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}}
    if([IO.Path]::GetDirectoryName($temporary) -ne [IO.Path]::GetTempPath().TrimEnd([char[]]'\/') -or [IO.Path]::GetFileName($temporary) -notlike 'SmartM365-AdCoverage-*'){throw 'Unsafe test cleanup root.'}
    Remove-Item -LiteralPath $temporary -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCjWqFF/tzA0rUe
# HO1E9GFhTH1F6kreII5RO8y1Kqf11aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIF8KVGzeSpBbHvDhejghPkVZM0it0IqYViJZubiTY0k4MA0GCSqG
# SIb3DQEBAQUABIIBgFJBDliLUEkuQDZPPeGHvgb6pa2lGmwBsTZG+Fa7TLk7que5
# P+jojGQjL3St93L9p7hOidm6zSTWY02TJmYbzb5TJDuUFAseLAS8Wq3O7zjrEdDP
# x8LT1D9n+6uV0iy6aMeGTQSXBn6OT6oK0Hlz4a20qgaIonVxptgkxt1EAr55ebe0
# QqTKzsYErqLQ4aFfXAKTnWYOz9dVUYH333okamiW+pfCTijpj1Cg+k5G2LQuvUBc
# beIzwQJtW7WWvYC80ZZV2YLKeulSBSn5j90WDkRaHmfUhrHByE5yfuhP5ucAHGIS
# 5an9qX6X3jq8+PFKTVjyVWyjhz3JOiHqMZqYaDVyFrL73vzhe2DPcafG9trIxIBD
# 09O8vTlnmOxHqgyeZ6j6dYAtWLUzhjzTpdcNr/qVHci6VhxjvL5Y4tlQS2hG5LAS
# LTi9WOFiNuwx+s/WgsndwksNIgT8qXqTaZUL4768ITYtvRtjhxpvnAE9U0Garbzj
# Xs1TpoHv81FODhCLQaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxMDMx
# MjRaMC8GCSqGSIb3DQEJBDEiBCDjtvplVbnjKrvSOrSp6G6eggHSOwj7ej8pUsfT
# wgfUwzANBgkqhkiG9w0BAQEFAASCAgCGzeRpllx3nLXyhRJfIPbGcWZoWsq/j0+a
# gGJQHdrzgOJWyMWsMP4IAGB1lMfDfTNdAiyKD/psAiCtTOTm8sqP1YILGzKrhAMQ
# jbZQogNZNMy9YNs3ptG7hg0AU1TBII6bthDJ+7ohRJn546ISD48iFzo1Bqd/NVbE
# GNy0EvA30eSaNrQHQEGLLorL2m2heyNr/Gm4nU6jFiee5K3o6zrsp7YQhvLjQTFH
# Nv3ldurcm/NjY4jfufO/XRAO6W7wSIhBt/JsJcUfnXZBKbmErS2t3NOYM1xHLfSP
# FC9d+GqOectyOV62orfhmVyugRozL76yNseKrmaPWbOGRZuID3oqvKswrSnqlxD8
# JD1X9FTGE/SjjG5s3lf/PDGWnwTtxvRaZrZRO96B4PNIV9jmB5ZavTKmSAZ6BWTt
# t8055YSd3PGgmC12+h6FjpUQURV9fuz2rG4ve/n4A+evYRYf3S+ShVUJMNWVI3EQ
# 2fJLMwIncE7v5FkVn1m853opgf8zN9nx88a2eroOpGMz1H+YspV4OnDCeOv0QzOX
# g1oOIjoXHCay4FORTMzPO5VnYC3z814wMB9c2jJWBtRP69heIm6lwcEaxS4EZYpb
# 9IgQ8PPymlYvBuy1RaQ29N332MYw5nV8R34iL918yRnGYfou0OQbVbkzAQ2Jek0z
# hAJ5pqIKqg==
# SIG # End signature block
