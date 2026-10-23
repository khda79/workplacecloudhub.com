<#
.SYNOPSIS
    Exercise automatic assessment with fictional in-memory SharePoint objects only.
.VERSION
    1.0.1
#>
[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
$project=Split-Path $PSScriptRoot -Parent
. (Join-Path $project 'Scripts/Inventory/SmartM365-SharePointSource-Assessment.ps1')
$baselineReader=${function:Get-SPSAStandardRoleMasks}
$pageAdapter=${function:Get-SPSAPageWebParts}
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('SPAssessment-offline-'+[guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $testRoot)
$rulesDir=Join-Path $testRoot 'Config';[void](New-Item -ItemType Directory -Path $rulesDir)
Copy-Item -LiteralPath (Join-Path $project 'Config/source-assessment.rules.json.template') -Destination $rulesDir
$script:checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
# Every adapter that would access the farm is replaced before use.
function Get-SPSAWorkflowSubscriptions {param($Web) return $Web.MockSubscriptions}
function Get-SPSAPageItems {param($List) return $List.MockPageItems}
function Get-SPSAStandardRoleMasks {param($Context) return @{Reader='1';Contributor='3'}}
function Get-SPSAPageWebParts {param($Web,$Url) return $Web.MockWebParts}
foreach($name in @('Get-SPSite','Get-SPWeb','Get-SPFarm','Add-PSSnapin','Connect-PnPOnline','Invoke-WebRequest','Invoke-RestMethod','Send-MailMessage','Start-Process')){
    Set-Item ('Function:'+$name) {throw 'A live transport was reached by an offline assessment test.'}
}
function New-FakeList {
    [pscustomobject]@{RootFolder=[pscustomobject]@{ServerRelativeUrl='/sites/demo/HiddenList'};BaseTemplate=100;BaseType='GenericList';Hidden=$true;WorkflowAssociations=@();ContentTypes=@();UserCustomActions=@();EventReceivers=@();Fields=@();Forms=@();MockPageItems=@()}
}
function New-FakeWeb {
    param($Lists=@())
    [pscustomobject]@{Url='https://source.contoso.example/sites/demo';WorkflowAssociations=@();MockSubscriptions=@();MockWebParts=@();UserCustomActions=@();EventReceivers=@();Features=@();Lists=@($Lists);Site=[pscustomobject]@{Url='https://source.contoso.example/sites/demo';Features=@();UserCustomActions=@();EventReceivers=@()}}
}
try {
    $ctx=New-SPSAAssessment File $rulesDir
    $permissionNames=@($ctx.Rules.StandardRolePermissions.PSObject.Properties.Value|ForEach-Object {$_}|Sort-Object -Unique|Where-Object {$_ -ne 'FullMask'})
    $members=@(for($i=0;$i -lt $permissionNames.Count;$i++){'{0} = 1UL << {1}' -f $permissionNames[$i],$i})
    Add-Type ('namespace Microsoft.SharePoint { [System.Flags] public enum SPBasePermissions : ulong { '+($members -join ', ')+', FullMask=ulong.MaxValue } }')
    $baseline=& $baselineReader $ctx
    Check ($baseline.Keys.Count -eq 6) 'Documented core default-role profiles must resolve using installed enum bits.'
    Check (([uint64]$baseline.Contributor -band [uint64]$baseline.Reader) -eq [uint64]$baseline.Reader) 'Contributor must include the Reader baseline.'
    $ctx.Rules.StandardRoleMasks|Add-Member -NotePropertyName Reader -NotePropertyValue '999'
    $overridden=& $baselineReader $ctx
    Check ($overridden.Reader -eq '999') 'Reviewed standard-mask overrides must be preserved.'
    $ctx.Rules.StandardRoleMasks.PSObject.Properties.Remove('Reader')
    Check (Test-Path (Join-Path $rulesDir 'source-assessment.rules.json.txt')) 'Rules runtime was not self-created.'
    $ctx.Rules.FunctionalRules.Workflow2010.Effort=2
    Write-SPSAJson (Join-Path $rulesDir 'source-assessment.rules.json.txt') $ctx.Rules
    $ctx=New-SPSAAssessment File $rulesDir
    Check ($ctx.Rules.FunctionalRules.Workflow2010.Effort -eq 2) 'Local rule override was overwritten.'
    $list=New-FakeList;$list.WorkflowAssociations=@([pscustomobject]@{Id='wf-list'})
    $list.BaseTemplate=115
    $web=New-FakeWeb @($list)
    $web.MockSubscriptions=@([pscustomobject]@{Id='wf2013'})
    Add-SPSAWeb $ctx $web
    Check ($ctx.Items.Count -eq 3) 'Hidden non-document list/workflow/InfoPath detection is incomplete.'
    Check (@($ctx.Items.Values|Where-Object Technology -eq 'Workflow2013').Count -eq 1) 'Workflow 2013 subscription missing.'
    Check (@($ctx.Items.Values|Where-Object Technology -eq 'InfoPath').Count -eq 1) 'InfoPath library missing.'
    Add-SPSAWeb $ctx $web
    Check ($ctx.Items.Count -eq 3) 'Repeated web processing double-counted components.'
    $csv=Join-Path $testRoot 'source-files.csv'
    [IO.File]::WriteAllText($csv,'"WebUrl";"LibraryUrl";"SizeBytes";"VersionsCount"'+"`r`n"+'"https://source.contoso.example/sites/demo";"/Docs";"100";"0"')
    Complete-SPSAAssessment $ctx $csv
    $result=Get-Content ($csv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check ($result.FunctionalComponents.Complete) 'A fully covered and classified source should be complete.'
    Check ($result.FunctionalComponents.CollectionComplete -and $result.FunctionalComponents.QualificationComplete) 'Coverage and qualification must both be complete.'
    Check ($result.Inputs.FileInventorySha256 -eq (Get-FileHash $csv -Algorithm SHA256).Hash) 'Source hash binding missing.'
    Check ($result.DetailsSha256 -eq (Get-FileHash (Join-Path $testRoot 'AssessmentDetails-source-files.csv') -Algorithm SHA256).Hash) 'Detail hash binding missing.'
    $page=[pscustomobject]@{ServerRelativeUrl='/sites/demo/Pages/Home.aspx';CustomizedPageStatus='Uncustomized'}
    $web.MockWebParts=@([pscustomobject]@{Id='wp1';TypeName='Microsoft.SharePoint.WebPartPages.ListFormWebPart';AssemblyName='Microsoft.SharePoint'})
    Read-SPSAPage $ctx $web $page
    Check (@($ctx.Items.Values|Where-Object Technology -eq StandardListView).Count -eq 1) 'Standard list form must be recognized.'
    $web.MockWebParts=@([pscustomobject]@{Id='wp2';TypeName='Contoso.Demo.CustomPart';AssemblyName='Contoso.Demo'})
    Read-SPSAPage $ctx $web $page
    Check (@($ctx.Items.Values|Where-Object Technology -eq ServerCode).Count -eq 1) 'Custom server Web Part must be detected.'
    $page.CustomizedPageStatus='Customized'
    Read-SPSAPage $ctx $web $page
    Check (@($ctx.Items.Values|Where-Object Technology -eq CustomizedPage).Count -eq 1) 'Customized page must remain visible for qualification.'
    $unknown=New-SPSAAssessment File $rulesDir
    Add-SPSAWeb $unknown (New-FakeWeb)
    Add-SPSAComponent $unknown 'custom-form' Form UnknownTechnology 'DEMO custom form without qualification'
    Add-SPSAComponent $unknown 'known-wf' Workflow Workflow2013 'DEMO known replacement risk'
    Complete-SPSAAssessment $unknown $csv
    $partial=Get-Content ($csv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check ($partial.FunctionalComponents.CollectionComplete -and -not $partial.FunctionalComponents.QualificationComplete) 'Unknown component must distinguish collection from qualification.'
    Check (-not $partial.FunctionalComponents.Complete) 'Unknown component must not silently score zero.'
    Check (@($partial.FunctionalComponents.Items|Where-Object Technology -eq Workflow2013).Count -eq 1) 'Known risk must survive partial qualification.'
    function Get-SPSAWorkflowSubscriptions {param($Web) throw 'DEMO service unavailable'}
    $unavailable=New-SPSAAssessment File $rulesDir
    Add-SPSAWeb $unavailable (New-FakeWeb)
    Complete-SPSAAssessment $unavailable $csv
    $failed=Get-Content ($csv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check (-not $failed.FunctionalComponents.CollectionComplete) 'Unavailable workflow service must not prove absence.'
    $permission=New-SPSAAssessment Permission $rulesDir
    Add-SPSAWeb $permission (New-FakeWeb)
    $bindings=@([pscustomobject]@{RoleDefinitionBindings=@(
        [pscustomobject]@{Id=1;Type='Reader';BasePermissions=1;Name='Lecture'},
        [pscustomobject]@{Id=2;Type='Reader';BasePermissions=5;Name='Modified standard'},
        [pscustomobject]@{Id=3;Type='None';BasePermissions=7;Name='Custom A'},
        [pscustomobject]@{Id=4;Type='None';BasePermissions=7;Name='Custom B'})})
    Add-SPSARoleBindings $permission $bindings '/sites/demo/Docs'
    Add-SPSARoleBindings $permission $bindings '/sites/demo/Docs/child'
    Check ($permission.Items.Count -eq 4) 'Distinct role definitions must retain their provenance.'
    Check (@($permission.Items.Values|Where-Object Occurrences -eq 2).Count -eq 4) 'Repeated assignments must update occurrences, not create duplicate definitions.'
    Check (@($permission.Items.Values|Where-Object IsBuiltIn).Count -eq 1) 'Standard role baseline was not respected.'
    Check (@($permission.Items.Values|Where-Object {-not $_.IsBuiltIn}).Count -eq 3) 'Custom and modified standard roles must both be detected.'
    $pcsv=Join-Path $testRoot 'source-permissions.csv';[IO.File]::WriteAllText($pcsv,'"ObjectScope";"ObjectId"')
    Complete-SPSAAssessment $permission $pcsv
    $roles=Get-Content ($pcsv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check ($roles.CustomPermissions.Complete) 'Fully covered permission assessment should be complete.'
    $permission.StandardMasks=@{}
    $permission.Items=@{}
    Add-SPSARoleBindings $permission $bindings '/sites/demo/Docs'
    Complete-SPSAAssessment $permission $pcsv
    $unclassifiedRoles=Get-Content ($pcsv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check (-not $unclassifiedRoles.CustomPermissions.QualificationComplete -and $unclassifiedRoles.CustomPermissions.CollectionComplete) 'Missing standard-role baseline must leave qualification partial.'
    Add-SPSAError $permission FilteredPermissionScope Options 'DEMO item scope excluded'
    Complete-SPSAAssessment $permission $pcsv
    $filtered=Get-Content ($pcsv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check (-not $filtered.CustomPermissions.CollectionComplete) 'Filtered permission scope must not claim full coverage.'
    function Get-SPSAWorkflowSubscriptions {param($Web) return $Web.MockSubscriptions}
    $ctx=New-SPSAAssessment File $rulesDir
    Add-SPSAWeb $ctx (New-FakeWeb)
    Complete-SPSAAssessment $ctx $csv
    $empty=Get-Content ($csv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check ($empty.FunctionalComponents.Complete -and @($empty.FunctionalComponents.Items).Count -eq 0) 'A complete empty assessment must establish absence.'
    # Non-regression: custom rendering on standard Web Parts/fields is not zero effort.
    $render=New-SPSAAssessment File $rulesDir
    $web=New-FakeWeb
    $web.MockWebParts=@([pscustomobject]@{Id='standard-with-script';TypeName='Microsoft.SharePoint.WebPartPages.XsltListViewWebPart';AssemblyName='Microsoft.SharePoint';JSLink='clienttemplates.js|~site/Style Library/custom.js';XslLink='';Xsl=''})
    Read-SPSAPage $render $web ([pscustomobject]@{ServerRelativeUrl='/Pages/View.aspx';CustomizedPageStatus='Uncustomized'})
    Check (@($render.Items.Values|Where-Object { $_.Technology -eq 'ClientScript' -and $_.Effort -eq 2 }).Count -eq 1) 'Custom Web Part JSLink must require adaptation.'
    $list=New-FakeList
    $list.Fields=@([pscustomobject]@{Id='field';TypeAsString='Text';JSLink='~site/Style Library/field.js'})
    Read-SPSAList $render $web $list
    Check (@($render.Items.Values|Where-Object Technology -eq ClientScript).Count -eq 2) 'Standard field JSLink must be assessed.'
    $standard=New-SPSAAssessment File $rulesDir
    Read-SPSARenderingOverrides $standard ([pscustomobject]@{JSLink=' clienttemplates.js ';XslLink='';Xsl=''}) 'standard'
    Check ($standard.Items.Count -eq 0) 'The built-in renderer must not be classified as custom.'
    foreach($property in @('XslLink','Xsl')) {
        $owner=[pscustomobject]@{JSLink='';XslLink='';Xsl=''}
        $owner.$property='DEMO override'
        Read-SPSARenderingOverrides $render $owner ('override-'+$property)
    }
    Add-SPSAWeb $render (New-FakeWeb)
    Complete-SPSAAssessment $render $csv
    $renderResult=Get-Content ($csv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check ($renderResult.FunctionalComponents.CollectionComplete -and -not $renderResult.FunctionalComponents.QualificationComplete) 'Unreviewed XSL overrides must leave qualification partial.'
    Check (@($renderResult.FunctionalComponents.Items|Where-Object Technology -eq CustomXslt).Count -eq 2) 'Both inline XSL and XslLink must be retained.'
    if(-not ('System.Web.UI.WebControls.WebParts.PersonalizationScope' -as [type])) {
        Add-Type 'namespace System.Web.UI.WebControls.WebParts { public enum PersonalizationScope { User, Shared } }'
    }
    $adapterWeb=[pscustomobject]@{Manager=[pscustomobject]@{WebParts=@([pscustomobject]@{ID='adapter-part';JSLink='custom.js';XslLink='custom.xsl';Xsl='<xsl/>'})}}
    $adapterWeb|Add-Member -MemberType ScriptMethod -Name GetLimitedWebPartManager -Value {param($Url,$Scope) $this.Manager}
    $metadata=@(& $pageAdapter $adapterWeb '/DEMO.aspx')
    Check ($metadata[0].JSLink -eq 'custom.js' -and $metadata[0].XslLink -eq 'custom.xsl' -and $metadata[0].Xsl -eq '<xsl/>') 'The real metadata adapter must retain rendering overrides.'
    # Non-regression: SPSite receivers are read independently and only once per site.
    $siteContext=New-SPSAAssessment File $rulesDir
    $web=New-FakeWeb
    $web.Site.EventReceivers=@([pscustomobject]@{Id='site-receiver';Assembly='Contoso.Demo.ServerCode'})
    Add-SPSAWeb $siteContext $web
    $web.Url+='/child'
    Add-SPSAWeb $siteContext $web
    Check (@($siteContext.Items.Values|Where-Object Technology -eq ServerCode).Count -eq 1) 'Site receiver must be detected without duplicate subsite records.'
    Check (@($siteContext.Detectors|Where-Object Detector -eq SiteEventReceivers).Count -eq 1) 'Site receivers must be enumerated only once.'
    $badSite=New-SPSAAssessment File $rulesDir
    $web=New-FakeWeb
    $web.Site.PSObject.Properties.Remove('EventReceivers')
    Add-SPSAWeb $badSite $web
    Complete-SPSAAssessment $badSite $csv
    $badResult=Get-Content ($csv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check (-not $badResult.FunctionalComponents.CollectionComplete) 'Unreadable site receivers must leave collection partial.'
    # Non-regression: a bad list does not hide risks in later lists.
    $continueContext=New-SPSAAssessment File $rulesDir
    $badList=New-FakeList
    $badList.PSObject.Properties.Remove('RootFolder')
    $badList|Add-Member -MemberType ScriptProperty -Name RootFolder -Value {throw 'DEMO list metadata unreadable'}
    $goodList=New-FakeList
    $goodList.WorkflowAssociations=@([pscustomobject]@{Id='workflow-after-bad-list'})
    Add-SPSAWeb $continueContext (New-FakeWeb @($badList,$goodList))
    Complete-SPSAAssessment $continueContext $csv
    $continued=Get-Content ($csv+'.assessment.json.txt') -Raw|ConvertFrom-Json
    Check (-not $continued.FunctionalComponents.CollectionComplete) 'Bad list must leave collection partial.'
    Check (@($continued.FunctionalComponents.Items|Where-Object Technology -eq Workflow2010).Count -eq 1) 'A later list workflow must survive an earlier list failure.'
    Check (@($continued.FunctionalComponents.Errors|Where-Object { $_.Detector -eq 'ListAssessment' -and $_.Scope -like '*list-index:1' }).Count -eq 1) 'Bad list must retain individual error provenance.'
    Complete-SPSAAssessment $ctx $csv
    if($OutputDirectory){
        [void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
        Copy-Item -LiteralPath $csv,($csv+'.assessment.json.txt'),(Join-Path $testRoot 'AssessmentDetails-source-files.csv'),$pcsv,($pcsv+'.assessment.json.txt'),(Join-Path $testRoot 'AssessmentDetails-source-permissions.csv'),(Join-Path $rulesDir 'source-assessment.rules.json.txt') -Destination $OutputDirectory -Force
        foreach($name in @('source-files.csv.assessment.json.txt','source-permissions.csv.assessment.json.txt')){
            $path=Join-Path $OutputDirectory $name
            $demo=Get-Content -LiteralPath $path -Raw|ConvertFrom-Json
            $demo|Add-Member -NotePropertyName Fictional -NotePropertyValue $true
            $demo|Add-Member -NotePropertyName ValidationMode -NotePropertyValue 'DEMO - offline test doubles; permission masks are fictional'
            Write-SPSAJson $path $demo
        }
    }
    'PASS: '+$script:checks+' automatic source assessment checks; fake objects only.'
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $safe=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $resolved.StartsWith($safe,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SPAssessment-offline-*'){throw 'Unsafe test cleanup path.'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAE/b6wBbcPHoMq
# bOJGgxm7w8tLEv47iqDwR8kVGdDBqaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDX05Q/m3wU2KhlBX2A66Mq
# 0PIE0hbCFTCcZOEF4KuuazANBgkqhkiG9w0BAQEFAASCAYCtdlKr/Zg3iQvCqVZn
# 61qUJKmy0sndO/pnHnWBF4jpQVZnuNS4KmL0I0mW/acR9Xhfki6RzF9dBCcus2g0
# 0rOWImw/72qr0Qiy2xPKf+ghPOvH69xAWqcyt9O3vHnjL9mSrSToodhgx3J8fyS/
# TzstaSh2rKzNHQ/MMOMhaupdD8BOP3HDiEN9hvHRl4Kq24ta//Cvw6B64PE3LSqX
# N67Kei2/x/DELFYb7iz6R1WPE8T2IsywIt2/+eMhduJntuPD+Tc61lsHRNngQQpa
# 3hl4nGUBnvP2mK02LD+Zv0ibW5CubfHtzYkO3KSkIxl5YH239mMfkwmtxPNL1Utx
# chlotjV7pNgBgjWG+PvREoB5ioc1VPK57p9Xahf+qi+nIU6P8DTEQ9j8Cqm9+I1S
# lqi5tFJgcXgpSdv8JPueG+xKe46dOyZ4P5Y3R53e90PD8g4bOrfz6xafJXo1+Z1d
# Quu3HMdclgGEMZYYOMB4Pd7PHLYWx7rtSOHyvx1ej3jkl2g=
# SIG # End signature block
