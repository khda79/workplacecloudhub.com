<#
.SYNOPSIS
Synthetic regression tests for non-Graph SmartInventory CSV publication paths.
.VERSION
1.0.1
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$ResultPath
)
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$results = New-Object 'System.Collections.Generic.List[object]'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-NonGraph-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)

function Assert-Offline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}
function Test-OfflineCase {
    param([string]$Name, [scriptblock]$Body)
    try { & $Body; $results.Add([pscustomobject]@{ Name=$Name; Passed=$true; Error='' }) }
    catch { $results.Add([pscustomobject]@{ Name=$Name; Passed=$false; Error=$_.Exception.Message }) }
}
function Get-FunctionText {
    param([string]$Path, [string[]]$Names)
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw "Source parse failed: $Path"}
    foreach($name in $Names){
        $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
        if($null -eq $node){throw "Function not found: $name"}
        $node.Extent.Text
    }
}
function New-IssueExporterModule {
    param([string]$CollectorPath)
    $corePath=Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365.Core.psm1'
    $coreNames=@(
        'Get-SmartM365CoreContextValue','Get-SmartM365CsvValidationBaseName','Get-SmartM365CsvValidationRule',
        'Assert-SmartM365CsvDataCompleteness','Add-SmartM365TenantKeyToCsvData',
        'Write-SmartM365CsvAtomically','Copy-SmartM365FileAtomically'
    )
    $definitions=@(Get-FunctionText -Path $corePath -Names $coreNames)
    $definitions+=@(Get-FunctionText -Path $CollectorPath -Names @('FlattenPath','CopyCsv','ExportIssues'))
    $module=New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
    & $module {
        $script:SmartM365CoreTenantKey='synthetic-a'
        $script:SmartM365CoreOrganizationKey='synthetic-org'
        $script:SmartM365CoreEnvironmentKey='test'
        $script:SmartM365CoreTenantId='00000000-0000-0000-0000-000000000001'
    }
    $module
}

try {
    $collectors=@(
        [pscustomobject]@{
            Name='Exchange hybrid identity issues'
            RelativePath='SmartInventory/ExchangeInventory/Migration/SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1'
            Detail='Exchange_HybridIdentity_Issues.csv'
            Summary='Exchange_HybridIdentity_Issues_Summary.csv'
            Data=@([pscustomobject]@{IssueNumber=101;ObjectGUID='synthetic-object-1';Potential_Issue='Synthetic mismatch';IssueCategory='Identity'})
        },
        [pscustomobject]@{
            Name='Windows 11 readiness issues'
            RelativePath='SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1'
            Detail='Intune_Windows11_Readiness_Issues.csv'
            Summary='Intune_Windows11_Readiness_Issues_Summary.csv'
            Data=@([pscustomobject]@{IssueCode='SYN-001';Area='Readiness';ObjectGUID_Norm='synthetic-device-1';Potential_Issue='Synthetic blocker';IssueCategory='Hardware';PriorityScore=10;IsBlocking=$true;RecommendedAction='Test only';ImpactMigration='Test only'})
        }
    )

    foreach($collector in $collectors){
        $path=Join-Path $SourceRoot $collector.RelativePath
        Test-OfflineCase "$($collector.Name) uses atomic final publication" {
            $copyText=@(Get-FunctionText -Path $path -Names @('CopyCsv'))[0]
            $exportText=@(Get-FunctionText -Path $path -Names @('ExportIssues'))[0]
            Assert-Offline ($copyText -match 'Copy-SmartM365FileAtomically' -and $copyText -notmatch '\bCopy-Item\b') 'CopyCsv still writes a final path directly.'
            Assert-Offline ($exportText -match 'Write-SmartM365CsvAtomically' -and $exportText -notmatch '\bExport-Csv\b') 'ExportIssues still writes a final path directly.'
        }
        Test-OfflineCase "$($collector.Name) DATA-ALL DATA-LAST schema and byte parity" {
            $output=Join-Path $testRoot (($collector.Name -replace '[^A-Za-z0-9]','-')+'-all')
            $latest=Join-Path $testRoot (($collector.Name -replace '[^A-Za-z0-9]','-')+'-last')
            [void][IO.Directory]::CreateDirectory($output);[void][IO.Directory]::CreateDirectory($latest)
            $module=New-IssueExporterModule -CollectorPath $path
            try {
                & $module {param($o,$l,$d) $script:OutputFolder=$o;$script:LatestFolder=$l;$script:RunStamp='20260101-000000';ExportIssues $d} $output $latest $collector.Data | Out-Null
                $allDetail=Join-Path $output $collector.Detail; $lastDetail=Join-Path $latest $collector.Detail
                $allSummary=Join-Path $output $collector.Summary; $lastSummary=Join-Path $latest $collector.Summary
                Assert-Offline ((Get-FileHash $allDetail).Hash -eq (Get-FileHash $lastDetail).Hash) 'Detail DATA-ALL/DATA-LAST bytes differ.'
                Assert-Offline ((Get-FileHash $allSummary).Hash -eq (Get-FileHash $lastSummary).Hash) 'Summary DATA-ALL/DATA-LAST bytes differ.'
                $header=(Import-Csv -LiteralPath $lastDetail | Select-Object -First 1).PSObject.Properties.Name
                Assert-Offline (($header | Select-Object -First 4) -join ',' -eq 'TenantKey,OrganizationKey,EnvironmentKey,TenantId') 'Identity-first schema changed.'
            } finally {Remove-Module $module}
        }
        Test-OfflineCase "$($collector.Name) locked DATA-LAST preserves last valid export" {
            $output=Join-Path $testRoot (($collector.Name -replace '[^A-Za-z0-9]','-')+'-locked-all')
            $latest=Join-Path $testRoot (($collector.Name -replace '[^A-Za-z0-9]','-')+'-locked-last')
            [void][IO.Directory]::CreateDirectory($output);[void][IO.Directory]::CreateDirectory($latest)
            $lastDetail=Join-Path $latest $collector.Detail
            [IO.File]::WriteAllText($lastDetail,'LAST VALID SYNTHETIC EXPORT')
            $lock=[IO.File]::Open($lastDetail,'Open','Read','Read')
            $module=New-IssueExporterModule -CollectorPath $path
            try {
                $caught=$false
                try {& $module {param($o,$l,$d) $script:OutputFolder=$o;$script:LatestFolder=$l;$script:RunStamp='20260101-000000';ExportIssues $d} $output $latest $collector.Data | Out-Null}
                catch {$caught=$true}
                Assert-Offline $caught 'Locked DATA-LAST publication was acknowledged.'
            } finally {$lock.Dispose();Remove-Module $module}
            Assert-Offline ([IO.File]::ReadAllText($lastDetail) -eq 'LAST VALID SYNTHETIC EXPORT') 'Locked DATA-LAST was modified.'
        }
    }

    Test-OfflineCase 'AD HealthCheck uses atomic current and append paths' {
        $path=Join-Path $SourceRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1'
        $source=Get-Content -LiteralPath $path -Raw
        Assert-Offline ($source -match 'Add-SmartM365CsvRowsAtomically\s+-Data\s+\$all') 'AD history append is not atomic.'
        Assert-Offline ($source -match 'Write-SmartM365CsvAtomically\s+-Data\s+\$all\s+-Path\s+\$latestCsv') 'AD latest export is not atomic.'
        Assert-Offline ($source -notmatch 'Export-Csv\s+\$latestCsv') 'AD latest direct Export-Csv remains.'
        Assert-Offline ($source -match '-Encoding\s+utf8BOM') 'AD HealthCheck encoding contract changed.'
    }
    Test-OfflineCase 'Exchange infrastructure uses paired atomic publication' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
        $text=@(Get-FunctionText -Path $path -Names @('Export-ServersAndStorageCsv'))[0]
        Assert-Offline ($text -match 'Publish-CoreSmartM365Csv[\s\S]+-LatestPath\s+\$latestPath') 'LatestPath is not part of the paired Core publication.'
        Assert-Offline ($text -notmatch '\bCopy-Item\b') 'Infrastructure latest path still uses direct Copy-Item.'
        Assert-Offline ($text -match "-Delimiter\s+';'" ) 'Semicolon delimiter contract changed.'
    }
    Test-OfflineCase 'AD full inventory blocks incomplete sequential domains' {
        $path=Join-Path $SourceRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1'
        $source=Get-Content -LiteralPath $path -Raw
        foreach($label in @('OU','Computer','User','Group','Contact')){
            $failurePattern=[regex]::Escape("$label inventory failed for domain")+'[\s\S]{0,250}\bthrow\b'
            Assert-Offline ($source -match $failurePattern) "$label domain failure is still non-blocking."
        }
        $copyText=@(Get-FunctionText -Path $path -Names @('Copy-SmartM365AdFileWithRetry'))[0]
        Assert-Offline ($copyText -match 'Copy-SmartM365FileAtomically' -and $copyText -notmatch '\bCopy-Item\b') 'AD publication retry does not use atomic copy.'
        Assert-Offline ($source -match 'Add-SmartM365CsvRowsAtomically[\s\S]+AD daily summary snapshot') 'AD daily history append is not atomic.'
    }
    Test-OfflineCase 'AD enrichment generic list publication is atomic and PowerShell 7 safe' {
        $syntheticRows=New-Object System.Collections.Generic.List[object]
        [void]$syntheticRows.Add([pscustomobject]@{Id='synthetic'})
        $syntheticArray=$syntheticRows.ToArray()
        Assert-Offline ($syntheticArray.Count-eq1 -and $syntheticArray[0].Id-eq'synthetic') 'Generic object list ToArray conversion failed.'
        foreach($relative in @(
            'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Enrichment.ps1',
            'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-UsersEnrichment.ps1'
        )){
            $source=Get-Content -LiteralPath (Join-Path $SourceRoot $relative) -Raw
            Assert-Offline ($source -match 'Write-SmartM365CsvAtomically\s+-Data\s+\$enrichedRows\.ToArray\(\)') "$relative does not publish a PowerShell 7-safe object array atomically."
            Assert-Offline ($source -notmatch '-Data\s+@\(\$enrichedRows\)') "$relative still uses the failing generic List[object] array-subexpression conversion."
            Assert-Offline ($source -notmatch '\$enrichedRows[\s\S]{0,80}\bExport-Csv\b') "$relative still exports directly."
        }
    }
    Test-OfflineCase 'Exchange local mailbox current copies are atomic and blocking' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        $publishText=@(Get-FunctionText -Path $path -Names @('Publish-SmartM365ExchangeLocalMailboxCsv'))[0]
        Assert-Offline ($publishText -match 'Copy-SmartM365FileAtomically') 'Mailbox DATA-LAST copy is not atomic.'
        Assert-Offline ($publishText -match "Failed to publish latest CSV copy[\s\S]+-Level 'ERROR'[\s\S]+throw") 'Mailbox DATA-LAST failure is still acknowledged as a warning.'
        $source=Get-Content -LiteralPath $path -Raw
        Assert-Offline ($source -notmatch 'Copy-Item\s+-LiteralPath\s+\$(dailyCsv|summaryCsv)') 'Mailbox daily or summary latest copy remains direct.'
    }
    Test-OfflineCase 'Mailbox inventory summary sends at most once per local day' {
        $path=Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        $definition=@(Get-FunctionText -Path $path -Names @('Invoke-SmartM365MailboxDailySummaryMail'))[0]
        $module=New-Module -ScriptBlock ([scriptblock]::Create($definition))
        $marker=Join-Path $testRoot 'mailbox-summary.sent'
        try {
            &$module { function script:WriteLog { param($Message,$Level) }; $script:SendCount=0 }
            $first=&$module {param($p)Invoke-SmartM365MailboxDailySummaryMail -MarkerPath $p -SendAction {$script:SendCount++}} $marker
            $second=&$module {param($p)Invoke-SmartM365MailboxDailySummaryMail -MarkerPath $p -SendAction {$script:SendCount++}} $marker
            $count=&$module {$script:SendCount}
            Assert-Offline ($first -and -not $second -and $count -eq 1) 'Mailbox summary was sent more than once on the same day.'
            Assert-Offline ((Get-Content -LiteralPath $marker -Raw).Trim() -eq (Get-Date).ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)) 'Mailbox summary marker does not store the local send date.'
        } finally {Remove-Module $module -Force}
    }
    Test-OfflineCase 'Exchange warnings and alert mail subjects retain their operational contracts' {
        $exchange=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1') -Raw
        $mailbox=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1') -Raw
        $winUpdate=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1') -Raw
        $adHealth=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1') -Raw
        $exchangeCompletionLine = '$completionStatus = if ($script:ServersAndStorageWarningCount -gt 0) { ''CompletedWithWarnings'' } else { ''Success'' }'
        Assert-Offline ([regex]::Matches($exchange,[regex]::Escape($exchangeCompletionLine)).Count -eq 2) 'Exchange completion does not propagate collector warnings.'
        Assert-Offline ($exchange -match '\$reportedWarningCount\s*=\s+\$exchangeReadinessWarningRows\.Count\s*\+\s*\$lowSpaceRows\.Count') 'Exchange readiness and capacity warnings are not included in completion reporting.'
        $warningExitPattern = [regex]::Escape('$script:CompletionStatus = $completionStatus') + '[\s\S]{0,180}' + [regex]::Escape('exit 3')
        Assert-Offline ([regex]::Matches($exchange,$warningExitPattern).Count -eq 2) 'Exchange warnings do not return the orchestrator warning exit code.'
        Assert-Offline ($mailbox -match 'Invoke-SmartM365MailboxDailySummaryMail\s+-MarkerPath') 'Mailbox summary email is not protected by the daily guard.'
        Assert-Offline (($winUpdate | Select-String -Pattern 'SMART365 - \[\$reportStatus\] WinUpdate Feature Update' -AllMatches).Matches.Count -eq 2) 'WinUpdate alert subject lacks the SMART365 prefix.'
        Assert-Offline (($adHealth | Select-String -Pattern 'SMART365 - \[\$\(\$worst\.ToUpperInvariant\(\)\)\] Active Directory Health Check' -AllMatches).Matches.Count -eq 1) 'AD Health alert subject lacks the SMART365 prefix.'
        Assert-Offline ($adHealth -match 'SMART365 - \[CRITICAL\] Active Directory Health Check failed') 'AD Health failure subject lacks the SMART365 prefix.'
    }
    Test-OfflineCase 'Weekly CSV histories and manifests use atomic primitives' {
        $teams=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/Teams/SmartM365-Teams-Inventory.ps1') -Raw
        $spo=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1') -Raw
        $licenses=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1') -Raw
        Assert-Offline ($teams -match 'Add-SmartM365CsvRowsAtomically[\s\S]{0,200}-Encoding\s+utf8BOM') 'Teams append history is not atomic or lost its BOM contract.'
        Assert-Offline ($spo -match 'Add-SmartM365CsvRowsAtomically[\s\S]{0,200}-Encoding\s+utf8BOM') 'SharePoint append history is not atomic or lost its BOM contract.'
        Assert-Offline ($licenses -match 'Write-SmartM365TextAtomically\s+-Path\s+\$manifestPath') 'Licensing weekly manifest is not atomic.'
    }
    Test-OfflineCase 'Entra empty canonical current copy is atomic' {
        $source=Get-Content -LiteralPath (Join-Path $SourceRoot 'SmartInventory/M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1') -Raw
        Assert-Offline ($source -match 'Copy-SmartM365FileAtomically\s+-SourcePath\s+\$emptyPendingExport\.TimestampedPath\s+-DestinationPath\s+\$currentPendingPath') 'Entra header-only current file is not promoted atomically.'
    }
}
finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $expectedParent=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $resolved.StartsWith($expectedParent,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SmartInventory-NonGraph-*'){throw 'Unsafe cleanup root.'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
if($ResultPath){$results|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $ResultPath -Encoding UTF8}
$results|Format-Table Name,Passed,Error -AutoSize
$failed=@($results|Where-Object{-not $_.Passed}).Count
Write-Output ("Cases={0}; Passed={1}; Failed={2}; PowerShell={3}" -f $results.Count,($results.Count-$failed),$failed,$PSVersionTable.PSVersion)
if($failed){exit 1}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBs2tzTY6VbJG5f
# 4+/meUkEodWTpt+vbbHyW0DoMJOofaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIA9Hmb5rDw/SwW9d3au3ETsvYccRvvBHqxSGF2wQ3j8hMA0GCSqG
# SIb3DQEBAQUABIIBgIl2hHjHv1JDNHKCR6bATTgqKSY7l/paKwUvi2OmaJ5bwUZe
# Q5qfVCcSHmaPA5kk29PsOkypT3SEPi0XfVmAczwpqdkxQj483cgP0sHdOxIp966+
# Y/dDQhrEEZ37FkCv3aDfyhepup6X48roC9ZdUsYtEeHpGNboJrhtMMcIckBdudx3
# j0pCqrxqtEF13gUNdG16J2C6SiIRQaIvjFS2IH7YDNWy/pNnhoA7TxmD5t89Jo00
# sFs5ZH8EaTJcWaNYWvJw/7DtrayCckdc7KzOSBPvajWjY833Z4HcirMQvVjWaGBN
# k8reYnIvOfLaVfgjSc9gnvAgl0MTFMCfKi8igFs/Rcm6G8AmnA7G4H1M9vHnjjPr
# ix7pIviQ2eeB2qJaTf4bxTk8GBQP5TsYIjU7a86EPptqh7bY5eMdK/eG0xHJL1ya
# NvoD1OwjijnXeaOdEh8qo4nixWusokObgpBgJu+pNoM2d+MjUi84rk7xaHIi2bPH
# GGtf+mYZXyh65eqJgKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjExNzE1
# MjJaMC8GCSqGSIb3DQEJBDEiBCB7oR8wnqPJzGDU8yK2EjZkkRpBO1PyxpFp2rRt
# 0k9V9zANBgkqhkiG9w0BAQEFAASCAgBfJpJhAWcVmv8z09SL7748GXZOoMZjbspk
# v4t6Z0toUSfT2TllKsh3SLuWreHdT3jnYgZ2d/791Id0jPMh/0ZdyL7M1SHAI5ZJ
# ivXGS0EtBQSL0Gdd6Gn680+UyakL5MqCd6JqCkXw9buftQGNWwTVVVUJBVBjc03u
# 1Xne1WYQtvORovUrgHp7XVM+Le9W4leuMslM3Fp2mKTUx8EAH1jJTvGHrpZ6rlq5
# /F0ZDkj1Bdu24mvUjI0FQ2Gz2Tyuqo/lxUQCmJYf3fa3eDBtoUfrbGWcxAGMaqv3
# /O1W/MLoO2UZlMkUraxmqP1L1fAdEXi7VA3EwXmu8b1ZABnX5x9TTcNkd/+BdHyF
# QV/s30YEeoaNDKbTcWtSM2tG5NvD571cuZYo0nJyofV96mHW7cx1diuGsg6ZiPOe
# XyDTVPCubsVGkuTfzCrYRYkxf38/bICkXmvXvYDZzNWfJP9HPf8vpLIMSprz2jeD
# 676ctgB6soY0QmCJ7ZYWB2N+6qAe5km8Qp02ywu4DsN44Z08iteioAu0VsBiUe0s
# p5j3IAKM76H6Nw5aIhoJsx/3b8oMfTCQEeIfY0GNKXUnNP/e7/xmMkWptJeC7BmO
# 2xpU2UFkVtFha4Q8RqNeE818pnMgcQpA3kk1zXCx2ktYhzcdJ6HqA/AsrhYBcBPc
# mJ8ynXcRWA==
# SIG # End signature block
