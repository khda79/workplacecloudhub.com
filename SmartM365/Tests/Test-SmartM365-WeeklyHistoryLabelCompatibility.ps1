<#
.SYNOPSIS
Offline tests for weekly history label compatibility.
.VERSION
1.0.2
#>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$transport=Get-Module SmartM365.JsonTransport
$policy=& $transport {(Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock}
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-WeeklyLabels-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory $root
$script:passed=0
function Check($Value,$Message){if(-not $Value){throw $Message};$script:passed++}
function Reject($Action,$Message){$rejected=$false;try{& $Action|Out-Null}catch{$rejected=$true};Check $rejected $Message}
function WriteLog {param($Message,$Level)}
function Get-SmartM365IsoWeekName {'2026-W40'}
function Get-SmartM365WeeklyHistoryMaxRetentionWeeks {0}
function Get-SmartM365WeeklyHistoryFileName {param($Path)[IO.Path]::GetFileName($Path)}
function Copy-SmartM365FileAtomically {param($SourcePath,$DestinationPath)Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath}
function Invoke-SmartM365SharePointCsvUpload {param($LocalFilePath)$script:uploads+=$LocalFilePath}
$families=@(
    @{Path='Exchange/OnPrem/ServersAndStorage';Labels=@('Exchange on-prem infrastructure and readiness')}
    @{Path='Exchange/OnPrem/Mailboxes';Labels=@('Exchange on-prem mailboxes','Exchange on-prem remote mailboxes','Exchange on-prem mailbox daily stats')}
    @{Path='Exchange/OnPrem/CalendarPermissions';Labels=@('Exchange on-premises calendar permissions')}
    @{Path='M365/Licensing/Licenses';Labels=@('M365 licenses inventory')}
    @{Path='M365/Backup/PolicyScope';Labels=@('M365 Backup policy scope inventory')}
    @{Path='M365/Teams/Inventory';Labels=@('Microsoft Teams inventory')}
    @{Path='M365/SharePoint/Inventory';Labels=@('SmartM365 SharePoint Online inventory')}
)
function Fixture($Name,$Relative,$Label) {
    $history=Join-Path $root "$Name/Tenants/fixture/DATA-ALL/$Relative/WeeklyHistory"
    $week=Join-Path $history '2026-W29';$null=New-Item -ItemType Directory $week -Force
    $csv=Join-Path $week 'Inventory.csv';[IO.File]::WriteAllText($csv,"Id`r`n42`r`n")
    $doc=[ordered]@{Week='2026-W29';HistoryRootPath=$history;HistoryLabel=$Label;Files=@('Inventory.csv');UpdatedAt='2026-07-15T00:00:00Z'}
    $legacy=Join-Path $week 'manifest.json';[IO.File]::WriteAllText($legacy,($doc|ConvertTo-Json),[Text.UTF8Encoding]::new($true))
    @{Root=$history;Legacy=$legacy;Document=$doc;Hash=(Get-FileHash $legacy).Hash;Csv=$csv;CsvHash=(Get-FileHash $csv).Hash}
}
try {
    & $transport {function script:Get-SmartM365JsonTransportPolicy {@{Mode='JsonText';QualifiedUncRoots=@();QualifiedLocalDataRoots=@()}}}
    # Start with the production failure: a named historical manifest read by the generic CSV helper.
    foreach($family in $families) {
        foreach($stored in @($family.Labels)+@('SmartM365 inventory')) {
            foreach($requested in @('SmartM365 inventory')+@($family.Labels)) {
                $f=Fixture ([guid]::NewGuid().ToString('N')) $family.Path $stored
                $null=Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $f.Root -HistoryLabel $requested
                Check ((Get-FileHash ($f.Legacy+'.txt')).Hash -eq $f.Hash -and -not(Test-Path $f.Legacy)) "Label compatibility failed: $stored -> $requested"
                Check ((Get-FileHash $f.Csv).Hash -eq $f.CsvHash) 'Historical CSV changed.'
            }
        }
    }
    foreach($scenario in @('wrongFolder','wrongLabel','wrongTenant','incomplete','invalid','different','resume')) {
        $relative=if($scenario -eq 'wrongFolder'){'M365/Licensing/Licenses'}else{'Exchange/OnPrem/ServersAndStorage'}
        $f=Fixture $scenario $relative 'Exchange on-prem infrastructure and readiness'
        if($scenario -eq 'wrongLabel'){$f.Document.HistoryLabel='Microsoft Teams inventory'}
        if($scenario -eq 'wrongTenant'){$f.Document.HistoryRootPath=$f.Root.Replace('\fixture\','\foreign\').Replace('/fixture/','/foreign/')}
        if($scenario -in 'wrongLabel','wrongTenant'){[IO.File]::WriteAllText($f.Legacy,($f.Document|ConvertTo-Json));$f.Hash=(Get-FileHash $f.Legacy).Hash}
        if($scenario -eq 'incomplete'){[IO.File]::Move($f.Csv,($f.Csv+'.retained'))}
        if($scenario -eq 'invalid'){[IO.File]::WriteAllText(($f.Legacy+'.txt'),'{')}
        if($scenario -eq 'different'){[IO.File]::WriteAllText(($f.Legacy+'.txt'),(($f.Document|ConvertTo-Json)+"`n"))}
        if($scenario -eq 'resume') {
            Copy-Item -LiteralPath $f.Legacy -Destination ($f.Legacy+'.txt')
            $journal=@{Owner='WeeklyHistory:Exchange on-prem infrastructure and readiness';Phase='Prepared';SHA256=$f.Hash}|ConvertTo-Json -Compress
            [IO.File]::WriteAllText(($f.Legacy+'.migration.log'),($journal+"`n"))
            $null=Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $f.Root -HistoryLabel 'SmartM365 inventory'
            Check ((Get-FileHash ($f.Legacy+'.txt')).Hash -eq $f.Hash) 'Journal recovery changed payload.'
        } else {
            Reject {Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $f.Root -HistoryLabel 'SmartM365 inventory'} "Unsafe $scenario accepted."
            Check ((Get-FileHash $f.Legacy).Hash -eq $f.Hash) 'Rejected history lost its bytes.'
        }
    }
    # Exercise alternating automatic and explicit saves in both shared helper implementations.
    foreach($source in @('../Modules/SmartM365.Core/SmartM365.Core.psm1','../Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $source),[ref]$tokens,[ref]$errors)
        if($errors.Count){throw ($errors|Out-String)}
        $fn=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Save-SmartM365WeeklyInventoryHistory'},$true)
        . ([scriptblock]::Create($fn.Extent.Text))
        $f=Fixture ([guid]::NewGuid().ToString('N')) 'Exchange/OnPrem/ServersAndStorage' 'Exchange on-prem infrastructure and readiness'
        $script:uploads=@();$step=0
        foreach($label in @('SmartM365 inventory','Exchange on-prem infrastructure and readiness','SmartM365 inventory')) {
            $step++;$csv=Join-Path $root "Source-$step.csv";[IO.File]::WriteAllText($csv,"Id`r`n$step`r`n")
            Save-SmartM365WeeklyInventoryHistory -SourceFiles $csv -HistoryRootPath $f.Root -HistoryLabel $label -RetentionWeeks 0
            Check ((Get-FileHash ($f.Legacy+'.txt')).Hash -eq $f.Hash) 'Alternating saves rewrote an older week.'
        }
        $current=Join-Path $f.Root '2026-W40/manifest.json.txt'
        $doc=Get-Content $current -Raw|ConvertFrom-Json
        Check (@($doc.Files).Count -eq 3) 'Alternating saves lost current-week CSVs.'
        Check ($script:uploads -contains ($f.Legacy+'.txt')) 'Historical manifest omitted from publication.'
    }
    [pscustomobject]@{Passed=$script:passed;FixtureRoot=$root;Evidence='Synthetic local history only; uploads mocked; no collector or remote access'}
} finally {& $transport {param($p)Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $p} $policy}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAjiQIBVQUZ/32w
# oVNZ5LYfE4vJ8kxBL9lfQkh1eoobFKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIN8fdQ7o2E5mKJUpWHQq3RXMUyzSovFclvHbM7TT3UdnMA0GCSqG
# SIb3DQEBAQUABIIBgCnHHS5f3DOYWMtlPM0ALLH95Z28jM2SYfM1+EEVwgUkvmb1
# 6tCNIhHwevgHIUJqBq3H+FlX9y9tlwlW27kGnP/IK26UV7O0FEvVnbbTcvbvCYBX
# Mlocw3edf+HVzbfvE9aBmRE/m7kj9IEXDuMMa3J5zkaFa8Kf8OH256KdJhmrToOK
# 9LalZjy5bwC7P0IuuWAufqkUW8oJlgDbDHEsK1vaaXDHcx9Siuhjuq0bs2BGJLuA
# ZdvXjDbO1lC13mSlW4DhSc1dw669X5E1YreZpSvhlxXzqWjxClGaLmbbRctl3lsQ
# XzdRdcqqjSjDPttSeGUtWlvxKLIzjamDR7RudxbYdckPf0RIg8CWm2G8McXivLs/
# AXiXk2Nz9Fq3ZtDicLQeNIwn4o84WxtgrRQj2ElUUXX1N8xvyTl6GbL/IxPdGvv/
# V6CJV3civVXfiGjHZZldYjSXYN/z+RkLgzl50fXt+aic3QnC3tlr1wXeGQofPA/R
# O4zOOMJGtCIw11lkXKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxNTI2
# MDFaMC8GCSqGSIb3DQEJBDEiBCAK/V7JL8wUzYq79oY6zCJ8y459WXRjRYK55kbf
# xx7TijANBgkqhkiG9w0BAQEFAASCAgAv0EGYJtLRP+kIuzflLJZwZBhj/B8mjW7a
# tQSAAlK6Fh97SrgH4FKPIxJ3LycpnyJ1Rx+pPN9+MSae94iXo7uNYnqlBfUpFzdl
# boG1iRgGBqrGVnEtr2pfUDYWC1qomfyyjDuKDFtrjqZQTAYKIks53+mjAeRdey1t
# my/90DkiaEeVwTh8k4jo6Ss2q1TgV0jEKjXkp0yxQyu5b47Ecrkbs/Fa0U+PPvv5
# mN2uGVABH9jpBYkUGy0pIujuh3mvaa+lO2JOi8qul+OMhKOCIfKTBytcCdXbZ8hn
# y2spSpZuFR0zfV8nQqH28i/OULmA3ECA/k8sBLClXOKXc3y5+1OSIneeu19EKyJr
# 7PDbB6ql9z45CRzoOw51BY9Wg0SKBP/od5qYhQZvkmNL8GvVpJI4cohY+q4fTCym
# 3pLm5M0FDnuJ4/eyj7C8ByJo+282pteCXE9NCKHzCLUTAKkRUgo4zMw4wR2mTc8l
# 5a0BAvp/2J2rOst39wu7RABcHAFSHwIEY4sb2r5QdlJAavruKhGcdY5YXPsF080w
# HZ5DnC0vFQa6+qmtNeIwRHwBM3ibpHDRIvQX6IDd8ZE0wGQUgikkb2yESxppfsp8
# 01QZFqpo6QePd1HXLKc8XmM9LJa8VgJWgD1TSzKI+bhr30PXThcV+MNwoULoefOr
# kqqQLeVoQA==
# SIG # End signature block
