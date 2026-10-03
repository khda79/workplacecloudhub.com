#Requires -Version 5.1
<#
.SYNOPSIS
Offline producer-receipt tests for PowerShell 5.1 and 7. No operational imports.
.VERSION
1.0.5
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core globals are saved and restored; no operational module is imported.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='Offline fixtures only in a validated unique temporary directory, removed in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$helper=Join-Path $root 'Modules/SmartM365.Core/SmartM365-CmdbReceipt.ps1'
$registry=Get-Content (Join-Path (Split-Path $helper) 'SmartM365-CmdbSources.json.txt') -Raw | ConvertFrom-Json
$temporary=Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-Cmdb-ReceiptTests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
$saved=@{}
foreach($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')){
    $v=Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name]=@{Exists=($null -ne $v);Value=$(if($v){$v.Value}else{$null})}
}
$global:SmartM365TenantKey='synthetic';$global:SmartM365OrganizationKey='test'
$global:SmartM365EnvironmentKey='test';$global:SmartM365TenantId='synthetic-tenant'
$global:csvGeneratedPaths=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$mock=New-Module -Name SyntheticCmdbReceiptTests -ScriptBlock {
    param($path)
    . $path
    $script:Limited=$false;$script:Warnings=0
    function Test-SmartM365MaxItemsMode { return $script:Limited }
    function Get-SmartM365ScriptVersionFromFile { param($Path) $null=$Path;return 'synthetic-test' }
    function WriteLog { param($Message,$Level) $null=$Message;if($Level -eq 'WARNING'){$script:Warnings++} }
    Export-ModuleMember -Function *
} -ArgumentList $helper
$script:checks=0
function Assert-ReceiptTest {param([bool]$Condition,[string]$Message) if(-not $Condition){throw $Message};$script:checks++}
function New-SyntheticProducer {
    param([string]$Folder='current')
    $directory=Join-Path $temporary $Folder
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $producer=$registry.Producers[0]
    $file=Join-Path $directory $producer.Files[0]
    [IO.File]::WriteAllText($file,"TenantKey,DisplayName`r`nsynthetic,`"multi`r`nline`"`r`n",[Text.UTF8Encoding]::new($false))
    [void]$global:csvGeneratedPaths.Add($file)
    return @{Directory=$directory;File=$file;Producer=$producer;Receipt=(Join-Path $directory $producer.Receipt)}
}
function Start-SyntheticReceipt {param($Fixture)
    & $mock {param($p,$r) Start-SmartM365CmdbSourceReceipt -ScriptPath $p -SourceRootPath $r} $Fixture.Producer.Script $Fixture.Directory
}
function Complete-SyntheticReceipt {param($Fixture,[bool]$Scope=$true,[string]$Status='Success',[int]$Errors=0)
    & $mock {param($scope,$qualified,$status,$errors)
        Set-SmartM365CmdbSourceScope -CompleteScope $qualified -Scope $scope
        Complete-SmartM365CmdbSourceReceipt -Status $status -ErrorCount $errors
    } $Fixture.Producer.Scope $Scope $Status $Errors | Out-Null
    return (Get-Content $Fixture.Receipt -Raw | ConvertFrom-Json)
}
try {
    $fixture=New-SyntheticProducer
    Start-SyntheticReceipt $fixture
    $running=Get-Content $fixture.Receipt -Raw | ConvertFrom-Json
    Assert-ReceiptTest ($running.Status -eq 'Running' -and $running.IsPartialInventory) 'Start must invalidate old completion evidence.'
    $receipt=Complete-SyntheticReceipt $fixture
    Assert-ReceiptTest ($receipt.Status -eq 'Completed' -and -not $receipt.IsPartialInventory) 'Complete producer rejected.'
    Assert-ReceiptTest ($receipt.Files[0].Rows -eq 1) 'Logical multiline CSV row count is wrong.'
    Assert-ReceiptTest ($receipt.Files[0].SHA256 -eq (Get-FileHash $fixture.File -Algorithm SHA256).Hash) 'Exact native hash missing.'
    Assert-ReceiptTest ($receipt.Files[0].RunId -eq $receipt.RunId -and $receipt.Scope -eq $fixture.Producer.Scope) 'Lineage missing.'
    Assert-ReceiptTest (@(Get-ChildItem $fixture.Directory -Filter '*.tmp').Count -eq 0) 'Receipt temporary file retained.'
    $old=$receipt.RunId
    Start-SyntheticReceipt $fixture
    $receipt=Complete-SyntheticReceipt $fixture -Scope $false
    Assert-ReceiptTest ($receipt.Status -eq 'Failed' -and $receipt.Files.Count -eq 0 -and $receipt.RunId -ne $old) 'Partial scope reused old receipt.'
    Start-SyntheticReceipt $fixture
    $receipt=Complete-SyntheticReceipt $fixture -Status Failed
    Assert-ReceiptTest ($receipt.Status -eq 'Failed') 'Failed acquisition was declared complete.'
    Start-SyntheticReceipt $fixture
    $receipt=Complete-SyntheticReceipt $fixture -Errors 1
    Assert-ReceiptTest ($receipt.Status -eq 'Failed') 'Error count was ignored.'
    $global:csvGeneratedPaths.Clear()
    Start-SyntheticReceipt $fixture
    $receipt=Complete-SyntheticReceipt $fixture
    Assert-ReceiptTest ($receipt.Status -eq 'Failed' -and $receipt.Error -match 'this run') 'Old CSV file rescued missing current export.'
    [void]$global:csvGeneratedPaths.Add($fixture.File)
    [IO.File]::WriteAllText($fixture.File,"TenantKey,Name`r`nforeign,test`r`n")
    Start-SyntheticReceipt $fixture
    $receipt=Complete-SyntheticReceipt $fixture
    Assert-ReceiptTest ($receipt.Status -eq 'Failed' -and $receipt.Error -match 'foreign tenant') 'Foreign tenant accepted.'
    [IO.File]::WriteAllText($fixture.File,"TenantKey,OrganizationKey,Name`r`nsynthetic,foreign,test`r`n")
    Start-SyntheticReceipt $fixture
    $receipt=Complete-SyntheticReceipt $fixture
    Assert-ReceiptTest ($receipt.Status -eq 'Failed' -and $receipt.Error -match 'identity differs') 'Foreign source organization accepted.'
    [IO.File]::WriteAllText($fixture.File,"TenantKey,Name`r`n")
    Start-SyntheticReceipt $fixture
    $receipt=Complete-SyntheticReceipt $fixture
    Assert-ReceiptTest ($receipt.Status -eq 'Completed' -and $receipt.Files[0].Rows -eq 0) 'Successful empty CSV rejected.'
    $readonly=New-SyntheticProducer 'readonly'
    & $mock {param($p,$r) Start-SmartM365CmdbSourceReceipt -ScriptPath $p -SourceRootPath $r -ReadOnly} $readonly.Producer.Script $readonly.Directory
    Assert-ReceiptTest (-not (Test-Path $readonly.Receipt)) 'Read-only run wrote a receipt.'
    & $mock {$script:Limited=$true}
    Start-SyntheticReceipt $readonly
    & $mock {$script:Limited=$false}
    Assert-ReceiptTest (-not (Test-Path $readonly.Receipt)) 'MAXITEMS run wrote a canonical receipt.'
    $lock=[IO.File]::Open($fixture.Receipt+'.collection.lock',[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $failed=$false
    try { Start-SyntheticReceipt $fixture } catch {$failed=$true} finally {$lock.Dispose()}
    Assert-ReceiptTest $failed 'Concurrent producer lock was ignored.'
    $foreign=Get-Content $fixture.Receipt -Raw | ConvertFrom-Json
    $foreign.TenantId='another-tenant'
    [IO.File]::WriteAllText($fixture.Receipt,($foreign | ConvertTo-Json -Depth 12))
    $failed=$false;try{Start-SyntheticReceipt $fixture}catch{$failed=$true}
    Assert-ReceiptTest $failed 'Foreign receipt ownership was overwritten.'
    Assert-ReceiptTest ($registry.Producers.Count -eq 17 -and @($registry.Producers | ForEach-Object {$_.Files}).Count -eq 33) 'Registry coverage changed.'
    $scopeDefaults=@{
        MaxItems=0;DomainWorker=$false;ReportOnly=$false;DuplicateAnalysisOnly=$false;TargetDomains=@()
        DeviceDetailMode='All';MaxApps=0;TopUsers=0;Top100=$false;PermissionsOnly=$false;DebugUPN=''
        DetectAllDomains=$true;ForceOverwriteCSV=$true;IncludeRemoteMailboxes=$true;RemoteMailboxesOnly=$false
        OnlyADPermission=$false;IncludedOrganizationalUnit=@();MaxSites=0;MaxTeams=0;Filter='';MaxDevices=0
        IncludeFeatureUpdates=$true;IncludeQualityUpdates=$true
        'script:Stat_DetailAppsFromCache'=0;'script:Stat_DetailAppsSkippedByResume'=0
        'script:CmdbReadinessDeviceExportComplete'=$true
        'script:LocalMailboxIssues'=@()
        'script:CmdbLocalMailboxSourceComplete'=$true
        hardwareRows=@([pscustomobject]@{CollectionStatus='Collected'})
        'script:DataQualityRows'=@([pscustomobject]@{ReportName='EADevicePerformanceV2';Status='Collected'},[pscustomobject]@{ReportName='EADeviceScoresV2';Status='Collected'})
        selectedReports=@([pscustomobject]@{Name='Office365ActiveUserDetail'})
    }
    $restricted=@{
        'SmartM365-Devices-Inventory.ps1'=@{hardwareRows=@([pscustomobject]@{CollectionStatus='Failed'})}
        'SmartM365-ActiveDirectory-Inventory.ps1'=@{TargetDomains=@('one-domain')}
        'SmartM365-Intune-DiscoveredApps-Inventory.ps1'=@{DeviceDetailMode='Top';'script:Stat_DetailAppsFromCache'=1;'script:Stat_DetailAppsSkippedByResume'=1;MaxApps=1}
        'SmartM365-Licences-Inventory.ps1'=@{TopUsers=1}
        'SmartM365-EXO-Mailboxes-Inventory.ps1'=@{Top100=$true;PermissionsOnly=$true;DebugUPN='one-user'}
        'SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'=@{IncludeRemoteMailboxes=$false;ForceOverwriteCSV=$false;DetectAllDomains=$false;IncludedOrganizationalUnit=@('one-ou');TargetDomains=@('one-domain');'script:CmdbLocalMailboxSourceComplete'=$false}
        'SmartM365-SPO-Inventory.ps1'=@{MaxSites=1}
        'SmartM365-Teams-Inventory.ps1'=@{MaxTeams=1}
        'SmartM365-EndpointAnalytics-Inventory.ps1'=@{'script:DataQualityRows'=@([pscustomobject]@{ReportName='EADevicePerformanceV2';Status='Collected'},[pscustomobject]@{ReportName='EADeviceScoresV2';Status='UnavailableInTenant'})}
        'SmartM365-Devices-UpgradeEligibility.ps1'=@{Filter='one-device';MaxDevices=1;'script:CmdbReadinessDeviceExportComplete'=$false}
        'SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'=@{IncludeFeatureUpdates=$false;IncludeQualityUpdates=$false}
        'SmartM365-M365UserActivity-Inventory.ps1'=@{selectedReports=@([pscustomobject]@{Name='M365AppUserDetail'})}
    }
    $scopeModule=New-Module -Name SyntheticCmdbScopeTests -ScriptBlock {
        function Test-CmdbScopeExpression {param($Expression,$Values)
            foreach($name in $Values.Keys){Set-Variable -Name ($name -replace '^script:','') -Scope Script -Value $Values[$name]}
            return (& ([scriptblock]::Create($Expression)))
        }
    }
    foreach($producer in $registry.Producers){
        $file=Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter $producer.Script | Select-Object -First 1
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors)
        Assert-ReceiptTest ($errors.Count -eq 0) ('Parser failure: '+$producer.Script)
        $commands=$ast.FindAll({param($a) $a -is [Management.Automation.Language.CommandAst]},$true)
        Assert-ReceiptTest (@($commands | Where-Object {$_.GetCommandName() -match '^Start-(Core)?SmartM365CmdbSourceReceipt$'}).Count -eq 1) ('Missing start: '+$producer.Script)
        Assert-ReceiptTest (@($commands | Where-Object {$_.GetCommandName() -match '^Set-(Core)?SmartM365CmdbSourceScope$' -and $_.Extent.Text.Contains($producer.Scope)}).Count -eq 1) ('Missing scope: '+$producer.Script)
        $scopeCommand=$commands | Where-Object {$_.GetCommandName() -match '^Set-(Core)?SmartM365CmdbSourceScope$'} | Select-Object -First 1
        $expression=$null
        for($index=0;$index -lt $scopeCommand.CommandElements.Count;$index++){
            $element=$scopeCommand.CommandElements[$index]
            if($element -is [Management.Automation.Language.CommandParameterAst] -and $element.ParameterName -eq 'CompleteScope'){
                $expression=$scopeCommand.CommandElements[$index+1].Extent.Text;break
            }
        }
        Assert-ReceiptTest ([bool](& $scopeModule {param($e,$v) Test-CmdbScopeExpression $e $v} $expression $scopeDefaults)) ('Complete fixture scope rejected: '+$producer.Script)
        if($restricted.ContainsKey($producer.Script)){
            foreach($name in $restricted[$producer.Script].Keys){
                $values=$scopeDefaults.Clone();$values[$name]=$restricted[$producer.Script][$name]
                Assert-ReceiptTest (-not [bool](& $scopeModule {param($e,$v) Test-CmdbScopeExpression $e $v} $expression $values)) ('Restricted scope accepted: '+$producer.Script+' / '+$name)
            }
        }
        $minimum=if($producer.Script -match 'Exchange-Local'){'1.0.48'}elseif($producer.Script -eq 'SmartM365-ActiveDirectory-Inventory.ps1'){'1.0.66'}elseif($producer.Script -eq 'SmartM365-SPO-Inventory.ps1'){'1.0.67'}else{'1.0.65'}
        Assert-ReceiptTest ($ast.Extent.Text -match ([regex]::Escape("-MinimumVersion '$minimum'"))) ('Missing module guard: '+$producer.Script)
    }
    "PASS: $script:checks offline producer receipt checks; PowerShell $($PSVersionTable.PSVersion)."
} finally {
    & $mock {if($script:SmartM365CmdbSourceContext -and $script:SmartM365CmdbSourceContext.Lock){$script:SmartM365CmdbSourceContext.Lock.Dispose()}}
    Remove-Module $mock -ErrorAction SilentlyContinue
    if(Get-Variable scopeModule -ErrorAction SilentlyContinue){Remove-Module $scopeModule -ErrorAction SilentlyContinue}
    foreach($name in $saved.Keys){if($saved[$name].Exists){Set-Variable -Name $name -Scope Global -Value $saved[$name].Value}else{Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue}}
    $resolved=[IO.Path]::GetFullPath($temporary)
    if((Split-Path $resolved -Parent) -eq [IO.Path]::GetTempPath().TrimEnd('\') -and (Split-Path $resolved -Leaf) -like 'SmartInventory-Cmdb-ReceiptTests-*'){
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBqzjFikZIkD7gb
# kGb5+4qNuujv8Gy4gsYro1krXzv6U6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBHVTRqUle5AyKycJJnqeqky76kqRfxyY2dad9f+XWixMA0GCSqG
# SIb3DQEBAQUABIIBgKSB5YbEQRjqvq/9CdxKxfXFil/djPQ2KBdUDIKPixUtD/e0
# CfTVdjeVcKmT8p5EBf2vMbxG9OS+unT4p9pflmbpfeaMbH+7aiAQ2cFF/jvojFm9
# K57OUxFoI3AapYtbu0G7Dg1muk6kpcotAa00m+8FMrSI4FYpVvitNycIyM0jLZrX
# vQDoSeauqnpbzu2BS+X9UScpj79NbVKf5r/r1P/dG0oJ8mtPiU5uiyz31+d7WCdC
# sjPRDVQHqDcGwAduHiQ1zNRi7dvZgCYcr7vKaJqkLoZOKAplcrC6WhuP6F9JoTwA
# vLHmYqK5JDxqlItXgBZoABpJ0L5gea9toR1UTkrHNyKj2Ea5+a0caKZZP5KiorPu
# FwLe4Z+p5OaF7FxirUv6eZuDblipC7CLvKazY/N6ixeRZnRCCu/BF8o/lNSmhV91
# YP7EKLVA+mD0KITkZXJJOKZGTKT3UQ3+wmUzzLcBeCr1wIO1c5tRZ/364zIDnazr
# rfDADKIVacVgFiM3D6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMyMjQ1
# NDFaMC8GCSqGSIb3DQEJBDEiBCDvBillSDgQPqfE/+0GXMv37YupZLQzWOrvy/8O
# TXixrTANBgkqhkiG9w0BAQEFAASCAgA7t9HXfbSJU7wyrV0JsqJLDbqNZFqhDzCB
# UqbPUnVqx9oKI52WTWzkjCjm9aBHzVXxu0Qt5hGjxHjG9ZAv358wncCN9mqMHPq2
# nOTKDNmK1xsrzOORGrVCTg9fYfiwwTORAOlL6uzVX2J5+UXkEvMUfD8AEuOihUNp
# pWtCr2sXw85yWzz9Db95nwwucSptQ9JyClQE5M+KGKe/OMMVy/+jufSuHsFJmD2Z
# aiKrMpNMg/kjhUnwXjsyfhJHc9wusBTdST39UpmvjQt2zJg/ECgkLNLgfn3yVYTo
# gXspp/991Ftr8Zdt3AdCP++/XSYyCXLvATsnV2mJ7mQfyKDL/GqSFbNpuQDimEvb
# vhCfHeMZpxEmEEr2q0qYgw0MELBQpshs++IKXSgHtH9nkNOtTRY6xERFQOv8tvZX
# creGjW+b0wvKqcao9DvkchwNG+eKNOqQEB/5jKgVgCKb3FX/n4Nsxci9Z9m6V0SF
# VLrx+8LLRP07+D+ThTy2HG+YH2iqIto76uVn1Bbe/fKBU52DDXTS7VcLgUzaqEj7
# CufdgS/b8QyhFsg0ISZFNZ0YnUOGnCe47ebfmkS6lTrbzmdZrmWIjTySc3ansIaC
# CwWRtMG4ZDn3BQv9i7ZI4HWSm6G6dkbWUFrGbE+hIKWF8XNyxMBBJnrVjcg3qibw
# /eZO5sY2XA==
# SIG # End signature block
