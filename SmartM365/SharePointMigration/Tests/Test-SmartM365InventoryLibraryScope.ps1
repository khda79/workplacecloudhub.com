<#
.SYNOPSIS
    Verify file and permission scan library scope without SharePoint connections.
.VERSION
    1.0.2
#>
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$script:SPOPermissionConnection = 'offline'

# Minimal SharePoint types let the real source helpers bind synthetic lists.
if ('Microsoft.SharePoint.SPList' -as [type]) { throw 'Run in a fresh offline PowerShell process without SharePoint assemblies.' }
Add-Type -TypeDefinition @'
namespace Microsoft.SharePoint {
    public enum SPBaseType { GenericList, DocumentLibrary }
    public class SPList {
        public string Title;
        public object RootFolder;
        public SPBaseType BaseType;
        public int BaseTemplate;
        public bool Hidden;
        public int ItemCount;
    }
    public class SPWeb { public SPList[] Lists; }
}
'@

function New-TestList {
    param([string]$Name, [int]$Items = 1, [bool]$Hidden = $false, [string]$BaseType = 'DocumentLibrary', [int]$Template = 101)
    $list = New-Object Microsoft.SharePoint.SPList
    $list.Title = $Name
    $list.RootFolder = [pscustomobject]@{ Url = $Name; Name = $Name; ServerRelativeUrl = '/project/' + $Name }
    $list.ItemCount = $Items
    $list.Hidden = $Hidden
    $list.BaseType = [Microsoft.SharePoint.SPBaseType]$BaseType
    $list.BaseTemplate = $Template
    return $list
}
function Get-PnPList { param($Connection) return $script:FixtureLists }
$script:TestDeniedTitle = ''
function Get-PnPProperty { param($ClientObject, $Property) if ($ClientObject.Title -eq $script:TestDeniedTitle) { throw 'Synthetic root access denied' }; return $ClientObject.$Property }
function Write-Host { param($Object) [void]$script:Messages.Add([string]$Object) }
function Write-Info { param($Color, $Message) [void]$script:Messages.Add([string]$Message) }
function Write-ConsoleMessage { param($Message) [void]$script:Messages.Add([string]$Message) }
function Assert-Names {
    param($Actual, [string[]]$Expected, [string]$Label)
    $names = @($Actual | ForEach-Object { $_.Title } | Sort-Object)
    if (($names -join '|') -ne (($Expected | Sort-Object) -join '|')) { throw "$Label selected unexpected lists: $($names -join ', ')" }
}

$script:FixtureLists = @(
    (New-TestList 'SiteAssets' -Items 16),
    (New-TestList 'SitePages' -Template 119),
    (New-TestList 'Documents' -Items 0),
    (New-TestList '_catalogs/masterpage'),
    (New-TestList 'HiddenLibrary' -Hidden $true),
    (New-TestList 'Events' -BaseType 'GenericList' -Template 106)
)
foreach ($side in @('Source', 'Target')) {
    foreach ($kind in @('File', 'Permission')) {
        $script:Messages = New-Object 'System.Collections.Generic.List[string]'
        $name = "SmartM365-SharePoint$side-${kind}Inventory.ps1"
        $path = Join-Path $PSScriptRoot "..\Scripts\Inventory\$name"
        $tokens = $null; $parseErrors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors)
        if (@($parseErrors).Count) { throw "Parser errors: $name" }
        $systemFunction = if ($kind -eq 'File') { 'Test-SystemLibrary' } else { 'Test-SystemList' }
        $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $systemFunction }, $true)
        if (-not $definition) { throw "Missing system filter in $name" }
        Invoke-Expression $definition.Extent.Text
        foreach ($list in $script:FixtureLists[0..1]) {
            if (& $systemFunction -List $list) { throw "$name classified $($list.Title) as system content" }
        }
        if (-not (& $systemFunction -List $script:FixtureLists[3])) { throw "$name no longer excludes the master-page catalog" }

        if ($kind -eq 'File') {
            if ($side -eq 'Target') {
                $retry = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-SPORead' }, $true)
                Invoke-Expression $retry.Extent.Text
            }
            $definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-DocumentLibraries' }, $true)
            Invoke-Expression $definition.Extent.Text
            $args = @{}
            if ($side -eq 'Source') {
                $web = New-Object Microsoft.SharePoint.SPWeb
                $web.Lists = $script:FixtureLists
                $args.Web = $web
            } else { $args.Connection = 'offline' }
            Assert-Names @(Get-DocumentLibraries @args) @('SiteAssets', 'SitePages') "$name default"
            Assert-Names @(Get-DocumentLibraries @args -IncludeHidden -IncludeSystem) @('SiteAssets', 'SitePages', '_catalogs/masterpage', 'HiddenLibrary') "$name opt-ins"
            if ($side -eq 'Target') {
                $script:TestDeniedTitle = 'HiddenLibrary'
                Assert-Names @(Get-DocumentLibraries @args) @('SiteAssets', 'SitePages') "$name inaccessible excluded hidden library"
                $blocked = $false
                try { [void]@(Get-DocumentLibraries @args -IncludeHidden -IncludeSystem) } catch { $blocked = $_.Exception.Message -match 'Synthetic root access denied' }
                if (-not $blocked) { throw 'An inaccessible included hidden library was not rejected.' }
                $script:TestDeniedTitle = 'SitePages'; $blocked = $false
                try { [void]@(Get-DocumentLibraries @args) } catch { $blocked = $_.Exception.Message -match 'Synthetic root access denied' }
                if (-not $blocked) { throw 'An inaccessible visible library was not rejected.' }
                $script:TestDeniedTitle = ''
            }
            foreach ($reason in @('hidden', 'system', 'zero items')) {
                if (-not @($script:Messages | Where-Object { $_ -like "*: $reason*" }).Count) { throw "$name did not log $reason exclusions" }
            }
        } else {
            # Exercise the actual three pre-query filter blocks, preserving their continue behavior.
            $filters = @($ast.FindAll({ param($node)
                $node -is [Management.Automation.Language.IfStatementAst] -and
                ($node.Clauses[0].Item1.Extent.Text -match '^(-not \$IncludeHidden.*\$list.Hidden|-not \$IncludeSystem.*Test-SystemList|\$(LibrariesOnly|DocumentLibrariesOnly) -and)')
            }, $true))
            if ($filters.Count -ne 3) { throw "Expected three permission filters in $name; got $($filters.Count)" }
            $body = 'param([switch]$IncludeHidden,[switch]$IncludeSystem,[switch]$IncludeHiddenLists,[switch]$IncludeSystemLists,[switch]$LibrariesOnly,[switch]$DocumentLibrariesOnly) foreach ($list in $script:FixtureLists) { $listTitle=$list.Title; $listUrl=$list.RootFolder.ServerRelativeUrl; ' +
                (($filters | ForEach-Object { $_.Extent.Text }) -join "`n") + '; $list }'
            $select = [scriptblock]::Create($body)
            Assert-Names @(& $select) @('SiteAssets', 'SitePages', 'Documents', 'Events') "$name default"
            Assert-Names @(& $select -LibrariesOnly -DocumentLibrariesOnly) @('SiteAssets', 'SitePages', 'Documents') "$name library-only"
            Assert-Names @(& $select -IncludeHidden -IncludeHiddenLists -IncludeSystem -IncludeSystemLists) @('SiteAssets', 'SitePages', 'Documents', 'Events', '_catalogs/masterpage', 'HiddenLibrary') "$name opt-ins"
            foreach ($reason in @('hidden', 'system', 'not a document library')) {
                if (-not @($script:Messages | Where-Object { $_ -like "*: $reason*" }).Count) { throw "$name did not log $reason exclusions" }
            }
        }
        Microsoft.PowerShell.Utility\Write-Host "$name scope and exclusion logs: OK"
    }
}
Microsoft.PowerShell.Utility\Write-Host 'Offline library scope tests passed for all four inventory scripts.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAlpTiN2fHvtBaL
# MqFTBJ4X/L84dfORYIlIMSZPgNs8lqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJBPCVPi71SPO5q7b4+GAZV03K2D2WlO6Ya5VVxSIGldMA0GCSqG
# SIb3DQEBAQUABIIBgAsrcUeM23OkFppiNzmJeWCTLkCEqjy52+teVp/gh7GzeuGL
# N94Rsh4iwuYK4eLkVtUC3M9v1zuSpbCW6xU6+U+YIVEs7LDKoNTvT4h0vSiqZdKh
# 6odnkMKcAQelx6mTx4GbH/DRmr9QOu+BEUEBluhX28DW/3FsGLCymIBB/AZZmAz+
# MOIxdr2Qo81XuZV3aL/DXJ0FxVqgjKJNpKtGJLGpFf6IfkCo7hV3/NMq6jm70B9V
# tGEZ60B0zR/S8DFi4SCNrEBsbxXdT2StbOc/vr7ON6cA7CqUCs7r0eENVzh4Zr9z
# dbJ8mHrH2E4AJ29uLaLqhaEV/ssVBolgqa7rMsnSuhhRiMvJkB0eW6qNgFPBTWns
# ejZVk+XWwCX3twHPSntCsKGFa1a5jWhO4+Ol4vqQBwFknUMfHvkO+LZkOnjhjAoN
# tu/n9Qnrw0EF5Jz0CSDVaxaVr8NbyocOhE5VfP6Ziq5OKF81palTe86Mv/KXdnHc
# d+0dJvHu62NY/+oxu6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDQyMjM4
# MzBaMC8GCSqGSIb3DQEJBDEiBCDgNFdQmdkY3bFAj9KLydEoWR1FAB6eM2loJmiL
# DSkfIzANBgkqhkiG9w0BAQEFAASCAgCUpUp+pqRDadE0O4Sd0kRSqW4+uOPzLJ9X
# HZsXwpgNy1odd+v+AiBhyNSpeqppOg7S4ShWlUslcY+hH+3v44afZJ/StPvJWQad
# Rbg4Qtg2wZk0zzWqQQwCXkND5gsKNDGoe2jkEopZffqu9RtDAniu6bEok/phZQ5w
# 3mpWnqfuJrbou9N/XQpfLPq0bGDZGIPDRCr2oJnDch8YrOZW7Q3P7K/cx6HVC7m/
# rjMW7u2/mGPUpHyhc3I/vnTueod36TBli9+Z84SwVb0mDEA68vMhUzxTGVkvfr26
# 85JOQOEclAe6bjZ3JhaTGF/iIQHNHiiN3X1yZfqbdnwR4Fn+8pFQjk+Ff7kPutr3
# N0/Ud/JQihLkdZAZAqFcNOKQ3F78NGcOiZx1TYkcFyDeI+ZjIYhjXGxxsNEQpSJc
# L80N0lgyxp+fWd0EHbWiWHWSOVt/s6YU6T7PAULjUH1B3ySztp06+RCbDGkrmYpD
# IORClEmU98MSRZa+LoL2f4lxzM/MyhO+Z1AqxGdTqyMwH9yDFlKAmLGWlqPO8yOx
# bg9mQv8PIWVXYX5j3Mp3z0U8v45FyBj0lMh2S1sJRdFMSUa+/2u8ZCgFwm/c2/Ai
# GVQT2sRQM/t1wZDov3rFkEsAjeD8JeyPVNIpGR10NMiSZwobMoCUSJTUeNdYHK+F
# 5KYZX1jiSQ==
# SIG # End signature block
