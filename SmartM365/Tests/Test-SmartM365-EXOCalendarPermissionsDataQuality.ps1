<#
.SYNOPSIS
Synthetic, offline data-quality tests for the Exchange Online calendar-permission collector.
.VERSION
1.0.0
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$results = [System.Collections.Generic.List[object]]::new()

function Assert-Offline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Test-OfflineCase {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        [void]$results.Add([pscustomobject]@{ Name=$Name; Passed=$true; Error='' })
    }
    catch {
        [void]$results.Add([pscustomobject]@{ Name=$Name; Passed=$false; Error=$_.Exception.Message })
    }
}

function Get-FunctionText {
    param([string]$Path, [string[]]$Names)
    $tokens=$null; $parseErrors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$parseErrors)
    if($parseErrors.Count){throw "Source parse failed: $Path"}
    foreach($name in $Names){
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if($null -eq $node){throw "Function not found: $name"}
        $node.Extent.Text
    }
}

$collectorPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/CalendarPermissions/SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1'
$functionNames = @(
    'Get-CalendarPermissionPropertyValue',
    'ConvertTo-CalendarPermissionKeyComponent',
    'New-CalendarPermissionRow',
    'New-CalendarStatusRow',
    'Get-MailboxCoverageKey',
    'Resolve-CalendarPermissionExportRows',
    'Assert-CalendarPermissionCoverage',
    'Get-WorkerPermissionPropertyValue',
    'ConvertTo-WorkerPermissionKeyComponent',
    'New-WorkerCalendarPermissionRow',
    'New-WorkerCalendarStatusRow'
)
$module = New-Module -ScriptBlock ([scriptblock]::Create((@(Get-FunctionText -Path $collectorPath -Names $functionNames) -join "`n")))

try {
    Test-OfflineCase 'Stable principal identity is exported and exact stable-key duplicates are removed' {
        $observed = & $module {
            $principal = [pscustomobject]@{
                ADRecipient = [pscustomobject]@{
                    ExternalDirectoryObjectId = '00000000-0000-0000-0000-000000000123'
                    PrimarySmtpAddress = 'delegate@example.test'
                    RecipientTypeDetails = 'UserMailbox'
                }
            }
            $permission = [pscustomobject]@{ User=$principal; AccessRights=@('Reviewer','Editor','Reviewer') }
            $row = New-CalendarPermissionRow -Mailbox 'owner@example.test' -UPN 'owner@example.test' -CalendarFolder 'Calendar' -Permission $permission
            $resolved = Resolve-CalendarPermissionExportRows -Rows @($row,$row)
            [pscustomobject]@{ Row=$row; Resolved=$resolved }
        }
        Assert-Offline ($observed.Row.PermissionPrincipalId -ceq '00000000-0000-0000-0000-000000000123') 'Stable Exchange principal ID was not exported.'
        Assert-Offline ($observed.Row.PermissionPrincipalSmtpAddress -ceq 'delegate@example.test') 'Principal SMTP address was not exported.'
        Assert-Offline ($observed.Row.PermissionKeySource -ceq 'StableId') 'Stable ID was not preferred for the permission key.'
        Assert-Offline ($observed.Row.AccessRights -ceq 'Editor,Reviewer') 'Access rights were not normalized deterministically.'
        Assert-Offline ($observed.Resolved.Rows.Count -eq 1 -and $observed.Resolved.DeduplicatedStableRowCount -eq 1) 'Exact stable-key duplicate was not removed exactly once.'
        Assert-Offline ($observed.Resolved.Diagnostics.Count -eq 0) 'Stable-key duplicate was incorrectly reported as ambiguous.'
    }

    Test-OfflineCase 'Display-name-only collisions are preserved and explicitly diagnosed' {
        $observed = & $module {
            $firstPermission = [pscustomobject]@{ User='Same Display Name'; AccessRights=@('Reviewer') }
            $secondPermission = [pscustomobject]@{ User='Same Display Name'; AccessRights=@('Editor') }
            $first = New-CalendarPermissionRow -Mailbox 'owner@example.test' -UPN 'owner@example.test' -CalendarFolder 'Calendar' -Permission $firstPermission
            $second = New-CalendarPermissionRow -Mailbox 'owner@example.test' -UPN 'owner@example.test' -CalendarFolder 'Calendar' -Permission $secondPermission
            Resolve-CalendarPermissionExportRows -Rows @($first,$second)
        }
        Assert-Offline ($observed.Rows.Count -eq 2) 'Ambiguous rows were silently deduplicated.'
        Assert-Offline (@($observed.Rows | Where-Object CollectionStatus -eq 'AmbiguousPrincipalCollision').Count -eq 2) 'Ambiguous rows were not explicitly marked.'
        Assert-Offline ($observed.Diagnostics.Count -eq 2 -and $observed.AmbiguousPermissionKeyCount -eq 1) 'Ambiguous collision diagnostic is incomplete.'
    }

    Test-OfflineCase 'Parallel worker rows retain the same schema and stable-key semantics' {
        $observed = & $module {
            $principal = [pscustomobject]@{
                ADRecipient = [pscustomobject]@{
                    ExternalDirectoryObjectId = '00000000-0000-0000-0000-000000000456'
                    PrimarySmtpAddress = 'parallel@example.test'
                    RecipientTypeDetails = 'UserMailbox'
                }
            }
            $permission = [pscustomobject]@{User=$principal;AccessRights=@('Reviewer')}
            $parent = New-CalendarPermissionRow -Mailbox 'owner@example.test' -UPN 'owner@example.test' -CalendarFolder 'Calendar' -Permission $permission
            $worker = New-WorkerCalendarPermissionRow -Mailbox 'owner@example.test' -UPN 'owner@example.test' -CalendarFolder 'Calendar' -Permission $permission
            $workerMissing = New-WorkerCalendarStatusRow -Mailbox 'missing@example.test' -UPN 'missing@example.test' -CalendarFolder '' `
                -Status 'NoCalendarFolder' -Message 'Synthetic.' -UseLegacyPlaceholder $false
            [pscustomobject]@{Parent=$parent;Worker=$worker;WorkerMissing=$workerMissing}
        }
        Assert-Offline ((@($observed.Parent.PSObject.Properties.Name) -join '|') -ceq (@($observed.Worker.PSObject.Properties.Name) -join '|')) 'Parallel and sequential permission schemas diverge.'
        Assert-Offline ($observed.Parent.PermissionKey -ceq $observed.Worker.PermissionKey -and $observed.Worker.PermissionKeySource -ceq 'StableId') 'Parallel worker builds a different stable permission key.'
        Assert-Offline ($observed.WorkerMissing.CollectionStatus -ceq 'NoCalendarFolder' -and -not $observed.WorkerMissing.User) 'Parallel worker does not represent a missing calendar honestly.'
    }

    Test-OfflineCase 'Missing calendar folder remains an explicit mailbox coverage row' {
        $row = & $module {
            New-CalendarStatusRow -Mailbox 'missing@example.test' -UPN 'missing@example.test' -Status 'NoCalendarFolder' -Message 'Synthetic missing calendar.'
        }
        Assert-Offline ($row.CollectionStatus -ceq 'NoCalendarFolder') 'Missing calendar status was not exported.'
        Assert-Offline ($row.CollectionMessage -ceq 'Synthetic missing calendar.') 'Missing calendar diagnostic message was lost.'
        Assert-Offline ([string]::IsNullOrWhiteSpace($row.User) -and [string]::IsNullOrWhiteSpace($row.AccessRights)) 'Unavailable permission fields were presented as measured.'
    }

    Test-OfflineCase 'Coverage accepts main or structured error evidence and refuses silent loss' {
        $observed = & $module {
            $mailboxes = @(
                [pscustomobject]@{PrimarySmtpAddress='one@example.test';UserPrincipalName='one@example.test'},
                [pscustomobject]@{PrimarySmtpAddress='two@example.test';UserPrincipalName='two@example.test'}
            )
            $mainRow = New-CalendarStatusRow -Mailbox 'one@example.test' -UPN 'one@example.test' -Status 'NoCalendarFolder' -Message 'Synthetic.'
            $errorRow = [pscustomobject]@{Mailbox='two@example.test';UPN='two@example.test'}
            $accepted = Assert-CalendarPermissionCoverage -Mailboxes $mailboxes -ExportRows @($mainRow) -ErrorRows @($errorRow) -ProcessedCount 2
            $refused = $false
            try {
                Assert-CalendarPermissionCoverage -Mailboxes $mailboxes -ExportRows @($mainRow) -ErrorRows @() -ProcessedCount 2 | Out-Null
            }
            catch { $refused = $true }
            [pscustomobject]@{Accepted=$accepted;Refused=$refused}
        }
        Assert-Offline ($observed.Accepted.ExpectedCount -eq 2 -and $observed.Accepted.RepresentedCount -eq 2) 'Main/error union did not satisfy coverage.'
        Assert-Offline $observed.Refused 'Coverage barrier accepted a silently missing mailbox.'
    }
}
finally {
    Remove-Module $module -Force
}

$results.ToArray() | Format-Table -AutoSize
if ($ResultPath) { $results.ToArray() | Export-Csv -LiteralPath $ResultPath -NoTypeInformation -Encoding UTF8 }
$failed = @($results | Where-Object { -not $_.Passed })
if ($failed.Count -gt 0) {
    throw ("{0} EXO calendar-permission offline test(s) failed." -f $failed.Count)
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPHtz7Pj1WsPtB
# dcif1hdUbJ6huh/JTct7af3EnQH2E6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILy/A+QRoiQrH7Be66dAVI/cOEIY5mJVBZNuDp/rFuCzMA0GCSqG
# SIb3DQEBAQUABIIBgK7nqh4Bwr3lFg6gJIVV4KfRyNAs9vKylSEApxaoTLZxoX51
# g4sxPKSaNAjDL+wwHQFKKqJcIYvwCcC8yUbvBCVcLQXgQ2GH3TFHKbGB0GMepwwg
# zrPPY7Exie4xNm9nkNoRO0jsh8AOE/eYXhtfYNGJy1HYUUGx0i6v5IZal8un/MhZ
# O5u8JWqKbURlQvpFdjdjLUHZYKF4sZ55+E2H76rOT8vv84dKULpyQ6e2rgrzE1zP
# 5jGf3OhXlOUB66TPLCi4oO17Lf4jtHM1NrPKiYs5t1yJuuZ0EeU/7wdwWSEde2dx
# zZ9lwvZoa9u/rGSLJaNBb85hB6itEVDDOAdi8GG18kqBmE3NxmncFATdPJVwr7HM
# CuEWp+zaPKJswMtjsGryJGxq5/fajs9WH0fPxAQqMU29QxolNCaCg5vWQbE3Vo/M
# oDykmeojBfpOnYAbCj0Zilb2Utr+exW0gY7KLqRiDrVKcug5sE8a7SUxUvkeom7d
# J3QkgeBQ6VlOJnZV4qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjMxODE4
# MzhaMC8GCSqGSIb3DQEJBDEiBCDFV9delvKbPvYbfcsHfP2cTErWxiekLRpA3x1C
# lmNNDTANBgkqhkiG9w0BAQEFAASCAgCRg75Xe1Kg1ed1d7HSJsHn1SWTWWqwBk/v
# K16F6eaXN3sUdDvHQ34V/C1DbiuTcpLKqk0OgscCUL0tA1hQYQp5YSTB4cL0TDUa
# PfuaZdT/HwRxHeArLiy3UUHGhrZ8KcuPDbGxa+nwniF+cpV4u+e2kMomd6chOmQw
# 0Oy+j/g3laHvv/OKcdHB2JGPFYAwZPDNWOx8wabTM3P1DaGgYBdQJlTgPhf9kzTw
# /3u4UbIFgHE5B3mepNsa/QuDobRqK7SpiMPiOxSCQUTV6Oagwt49LAOjoux19/bY
# B2fwYkwmf/enlmUWrKZhmCWToU2kCxI/AWYXdYqR+HKt2KR5hkj7CY8loN1MywrU
# as5u/t5r+6rC/j1NxhANXyt66SMLK/xnuVwOHhGMaWLmpGLS5zXRgy026k9bmBgv
# vdUsMQ5pdbwzia2d25it9YMznGJUGlx5eJ96qavHwPYBErZvSKF8GIW99FWRCBJ0
# lAqfGYZTxR4IMh3gYCzT49izz2K18/q+CQEZGnLDRNdg8zT7Mgk1Vt92KQ76RSG0
# l4eNISDglGl+dW7Xig4RU2abrlBJSzqHtgfe6ROBdy65SaALA2tJTq0VeqCaWbY/
# 7swLICn23QoRlX0PmcZWhEtgYK0U20s3HQt6E0mOC0ojlUw13RZRmzV74bvi/INs
# K0qE5uZFJg==
# SIG # End signature block
