[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Synthetic file/lock fixtures and isolated runspace cancellation only. No preparation entry point.
$ErrorActionPreference='Stop'
$script:Checks=0
function Check([bool]$Value,[string]$Message){$script:Checks++;if(-not $Value){throw $Message}}
$modulePath=Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1'
Import-Module $modulePath -Force
$module=Get-Module PreparedEvidencePipeline
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ('lifecycle-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force|Out-Null
$persona='SmartWorkplaceIntelligence-PersonaClassification.xlsx';$site='SmartWorkplaceIntelligence-SiteClassification.xlsx'
function Workbook([string]$Path,[string[]]$Sheets){
    $archive=[IO.Compression.ZipFile]::Open($Path,[IO.Compression.ZipArchiveMode]::Create)
    try{$entry=$archive.CreateEntry('xl/workbook.xml');$writer=[IO.StreamWriter]::new($entry.Open());try{$writer.Write('<workbook><sheets>'+ (($Sheets|ForEach-Object {'<sheet name="'+$_+'" />'}) -join '')+'</sheets></workbook>')}finally{$writer.Dispose()}}finally{$archive.Dispose()}
}
$fixtures=Join-Path $root 'fixtures';New-Item -ItemType Directory $fixtures|Out-Null
Workbook (Join-Path $fixtures $persona) @('Personas','Rules','Exclusions','NewRevision')
Workbook (Join-Path $fixtures $site) @('Settings','SiteTypes','Sites','NewRevision')
foreach($stage in @('Download','FirstRename','SecondRename')){
    $data=Join-Path $root $stage;New-Item -ItemType Directory $data|Out-Null
    Workbook (Join-Path $data $persona) @('Personas','Rules','Exclusions')
    Workbook (Join-Path $data $site) @('Settings','SiteTypes','Sites')
    $hashes=@{};foreach($name in @($persona,$site)){$hashes[$name]=(Get-FileHash (Join-Path $data $name)).Hash}
    $control=[hashtable]::Synchronized(@{Ready=$false;Stage=$stage})
    $actor=[powershell]::Create()
    try {
        $code={
            param($modulePath,$data,$fixtures,$control,$persona,$site)
            Import-Module $modulePath -Force
            $module=Get-Module PreparedEvidencePipeline
            if($control.Stage -ne 'Download'){
                & $module {param($control,$persona,$site)
                    $script:pauseControl=$control;$script:pauseName=if($control.Stage -eq 'FirstRename'){$persona}else{$site}
                    function script:Move-PreparedFile {param($Source,$Destination,[switch]$Overwrite)
                        [IO.File]::Move($Source,$Destination,[bool]$Overwrite)
                        if((Split-Path $Source -Parent) -like '*promotion' -and (Split-Path $Destination -Leaf) -eq $script:pauseName){
                            $script:pauseControl.Ready=$true
                            while($true){[Threading.Thread]::Sleep(10)}
                        }
                    }
                } $control $persona $site
            }
            $download={param($destination,$name)
                Copy-Item -LiteralPath (Join-Path $fixtures $name) -Destination $destination
                if($control.Stage -eq 'Download'){$control.Ready=$true;while($true){[Threading.Thread]::Sleep(10)}}
                Get-Item -LiteralPath $destination
            }.GetNewClosure()
            Receive-PreparedMappingWorkbooks -DataRoot $data -DownloadFile $download
        }
        $null=$actor.AddScript($code.ToString()).AddArgument($modulePath).AddArgument($data).AddArgument($fixtures).AddArgument($control).AddArgument($persona).AddArgument($site)
        $running=$actor.BeginInvoke();$deadline=[datetime]::UtcNow.AddSeconds(15)
        while(-not $control.Ready -and -not $running.IsCompleted -and [datetime]::UtcNow -lt $deadline){[Threading.Thread]::Sleep(10)}
        Check $control.Ready "Fixture did not reach interruption point: $stage"
        $actor.Stop()
        foreach($name in @($persona,$site)){Check ((Get-FileHash (Join-Path $data $name)).Hash -eq $hashes[$name]) "Rollback preserves original DATA pair: $stage/$name"}
        Check (@(Get-ChildItem (Join-Path $data 'mappings') -Directory).Count -eq 0) "Interruption leaves no owned mapping folder: $stage"
        $lock=[IO.File]::Open((Join-Path $data '.prepared-source.lock'),'Open','ReadWrite','None');$lock.Dispose()
        Check $true "Source lock released: $stage"
    } finally { $actor.Dispose() }
}
# Invoke only the owned-payload cleanup helper, never Invoke-PreparedEvidencePipeline.
$work=Join-Path $root 'work';$runId=[guid]::NewGuid().ToString('N');$run=Join-Path $work $runId
New-Item -ItemType Directory $run -Force|Out-Null
[IO.File]::WriteAllText((Join-Path $run 'run.json.txt'),(@{Owner='PreparedEvidencePipeline/v1';RunId=$runId;DataRoot=$root;TenantKey='synthetic'}|ConvertTo-Json))
[IO.File]::WriteAllText((Join-Path $run 'capture.json.txt'),'{}')
[IO.File]::WriteAllText((Join-Path $run 'preparation.log'),'synthetic audit')
[IO.File]::WriteAllText((Join-Path $run 'IntuneSourcePolicy.json.txt'),'{}')
New-Item -ItemType Directory (Join-Path $run 'source'),(Join-Path $run 'prepared')|Out-Null
& $module {param($work,$id,$data) Clear-PreparedRunPayload $work $id $data 'synthetic'} $work $runId $root
Check (-not(Test-Path (Join-Path $run 'IntuneSourcePolicy.json.txt'))) 'Intune source policy is removed'
Check (-not(Test-Path (Join-Path $run 'source')) -and -not(Test-Path (Join-Path $run 'prepared'))) 'Captured payload is removed'
Check ((Test-Path (Join-Path $run 'run.json.txt')) -and (Test-Path (Join-Path $run 'capture.json.txt')) -and (Test-Path (Join-Path $run 'preparation.log'))) 'Audits remain preserved'
$tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($modulePath,[ref]$tokens,[ref]$errors)
$publish=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Publish-PreparedEvidenceBatch'},$true)
$publishTry=@($publish.Body.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally})[0]
$finalText=$publishTry.Finally.Extent.Text;$cleanupText=$finalText.Substring(1,$finalText.Length-2)
$catchText=$publishTry.CatchClauses[0].Body.Extent.Text;$catchText=$catchText.Substring(1,$catchText.Length-2)
Check (-not $catchText.Contains('Remove-PreparedOwnedPath')) 'Error handler defers deletion until the pointer is checked'
foreach($mode in @('Interruption','Error')){
foreach($reference in @('Old','Committed','Unreadable')){
    $output=Join-Path $root ('pointer-'+$mode+'-'+$reference);$batchId='20261014T000000000Z-00000001'
    $batch=Join-Path $output ('batches/'+$batchId);New-Item -ItemType Directory $batch -Force|Out-Null
    [IO.File]::WriteAllText((Join-Path $batch 'payload.txt'),'synthetic data')
    $pointer=Join-Path $output 'current.json.txt'
    $pointerText=if($reference -eq 'Unreadable'){'broken'}else{@{BatchId=$(if($reference -eq 'Committed'){$batchId}else{'previous'})}|ConvertTo-Json}
    [IO.File]::WriteAllText($pointer,$pointerText)
    [IO.File]::WriteAllText((Join-Path $output "current.$batchId.tmp"),'synthetic temporary pointer')
    & $module {
        param($cleanupText,$output,$batch,$batchId,$currentPath,$catchText,$mode)
        $committed=$false;$TenantKey='synthetic';$lock=[IO.File]::Open((Join-Path $output 'fixture.lock'),'Create','ReadWrite','None')
        try {
            if($mode -eq 'Error'){
                $combined="try { throw 'Synthetic post-rename failure' } catch {"+$catchText+'} finally {'+$cleanupText+'}'
                & ([scriptblock]::Create($combined))
            } else { & ([scriptblock]::Create($cleanupText)) }
        } catch {if($_.Exception.Message -ne 'Synthetic post-rename failure'){throw}}
    } $cleanupText $output $batch $batchId $pointer $catchText $mode
    Check ([IO.File]::ReadAllText($pointer) -eq $pointerText) "Current pointer is never rewritten by cleanup: $reference"
    Check (-not(Test-Path (Join-Path $output "current.$batchId.tmp"))) "Temporary pointer removed: $reference"
    Check ((Test-Path $batch) -eq ($reference -ne 'Old')) "Referenced/uncertain batches are preserved: $reference"
    $lock=[IO.File]::Open((Join-Path $output 'fixture.lock'),'Open','ReadWrite','None');$lock.Dispose()
    Check $true "Publication lock released: $reference"
}
}
Write-Host "PASS: $script:Checks interruption, rollback, payload and pointer cleanup checks. Synthetic fixtures only; no preparation, collection or transfer."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA6/vhT1eIfipat
# VKEmPU+mUvnBUOVUAlXZDvkQWtXh7qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILCgKRQMkBFklJ2y7j4kSbcTrc4pUTpxg0Opq//1j8+HMA0GCSqG
# SIb3DQEBAQUABIIBgIl9TTU70O00sL2nBgyUIPIuUES0uvqhB6xlflvfNs6K973i
# LnhbigPDlB+KNVxSOFIRr7KtlHzXidJji4nZ8WkIUw0dazYHl2NMUPsWhugjzDu1
# zv/tfD4y4l6TbTAKTMsmQd6Wv6aiGtMS/A51jWDUcgqPULUwwb8kld37/CuX/+Lx
# pJV8X087B3tDHna8rWZ7kPFgBkNlgzp7vWk2afoba4yQCv11q2t3qpwUAh98EBbW
# r3ZSvjiqOwKwZ2C/XZ6L1y1t0aOU3lBQHf2MyDtv+xpU6/OuwYQQwosa5uXrlmGp
# bbj994x+g2OBohgJf361wJmTyznbcTkr3/TQQ3Mhueq+PSXHydQNssIWX3Nyo6Dl
# 3Ca70tGf54wqrjDWXmjXi+XQb5ZT0U/JCOP9CYwmiHKpEzyMxAfy59pUjCkapXsV
# mSBCgYYhijEpTtHOAP4/WPEx07AiSpe5kMCkpYXXNE8q7MGvbEwRBP0lyY5wr5V2
# 15SWrQ+Vi3x1VqPxlqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTQxNTMw
# NDVaMC8GCSqGSIb3DQEJBDEiBCBhmkcVApS3XYo72dxKzm0R9zWNkQ05zPLMXQqG
# OY2fvjANBgkqhkiG9w0BAQEFAASCAgBVjYmR4hWPY+BCzI+DAguBUbvRICfGR40m
# gnHvrfkgj1s7M1pnERIh9Da+irp8GvLHJCoSIJlFz8/B0WgoyhXf2pBppEc7dAj2
# /R82dioEQOsffBaQGrjF3LEmKPonSiTX7LaEa9WB+1HYkmLC4nc4C81ZEH1Cj8+Z
# JA+iTuKNOq+OwdbyTv3jHcUyovZTPLSMrQzCVcXHhyaGXbWaW+fnQwnay55lKVp2
# eVqu9AJOeu5Ql7OGsNgecluaPB0htwP+2Y24/nXxDQsr3H2tNNl1f/cUDpQF5qyp
# cnf++rSoczZ0NFbfrheTa4k1u9aBXc8KfIsIJWtfVMKMGnZQShO+q1IoLtNl2SAd
# tea6QA6aOxpy6LS4ZomhsqMwsuyjIgKC7hSJbw1x0KKR6xsc5sTdWiGJyu5xIB+R
# Bs1YZRILSDnEiYVTmlCUEM8xHqEWg226Q7jaCdikZGzklTUstZJHRtUbK3GAXgAR
# HPpF5W9gWvkNUWMgHUWxXcEGKGTbBKJD46Z9B4Opc6fw0yHIH5fD+4c//xsg22me
# cpeNgRXNtukb/daDiC9BrgAUeVs5o02IZEYS7GEdhRirO6WHS8/wJP9hyraxiI/w
# tjR9db4ymC5IbRZuAUgFXtMCzTUj4ph73Y0zzFCJ99pu+edW+G2iKCglAiw6vyHT
# ZcYobZdr3g==
# SIG # End signature block
