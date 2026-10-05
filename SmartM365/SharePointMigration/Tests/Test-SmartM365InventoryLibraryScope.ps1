<#
.SYNOPSIS
    Verify file and permission scan library scope without SharePoint connections.
.VERSION
    1.0.1
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
function Get-PnPProperty { param($ClientObject, $Property) return $ClientObject.$Property }
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBT61+tP+4BAqzB
# IDJKMwvh1+iXOUKM27MDuMujPk39S6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAr3P9lJl4BzilVykPnqq5T
# v+WqAdZZ2IQ5JMjEplsQ4zANBgkqhkiG9w0BAQEFAASCAYCMAjqHqmPA42comCVv
# qta/jDik5PORrAd8XwpLZYJinNamKsiGJgakT+3NFfpW+Fp9OCY2YwfrkoI3HeqM
# gB4F0Zv0e2q3glZbfv4fUhQ96dCmgJk18PIeJZaEt4kI2SXg4TAUN9M98lq1Td1+
# vABRw8UdU/pnUgpOSHttWJJC2+MKwcZTJn4SjlKUnWVWlS4MKSzvfP6zYBeqDxou
# eHyHDI7cnzbB0ydAzsYyYphhIPc5jueBy/g8sKgKBPTl9IJpYyA1b4VTUTWIoKEN
# 380LM1CSRb6WbFzvTb5vF5yr8Pvyh/dUSVXYw2yQuu6uHzoDPvNAQAdfOmcmcdXh
# ABXzdpSqdYx8tCNdvgT18HEzhZOdshY3Id5HosdmHnH/nClviuBSEq2ON953I5fB
# SZwrmUV61VCx1sn6yENap+ePuPIt3x4R/nmxDwNTL6IGD1mq3WJ/ZAVkFDMing7y
# ufn7R4LDpjETAmmMWHP2H2/jrS9NI1IptNHV/XcsLWmPinw=
# SIG # End signature block
