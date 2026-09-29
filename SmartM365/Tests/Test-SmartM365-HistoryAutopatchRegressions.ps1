[CmdletBinding()]
param([ValidateSet('History','Issue','Autopatch','Mailbox','All')][string]$Case='All')
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$module=Get-Module SmartM365.JsonTransport
$originalPolicy=& $module { (Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock }
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-HistoryRegression-' + [guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory $root
$script:passed=0
function Check($Value,$Message) {if(-not $Value){throw $Message};$script:passed++}
function Reject($Action,$Message) {$failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Check $failed $Message}
function Definition($Path,$Name) {
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $Path),[ref]$tokens,[ref]$errors)
    if($errors.Count){throw ($errors|Out-String)}
    $node=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name},$true)
    if(-not $node){throw "Missing function $Name"}
    $node.Extent.Text
}
try {
    & $module {function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@()} }}
    if($Case -in 'History','All') {
        $mailboxLabels=@('Exchange on-prem mailboxes','Exchange on-prem remote mailboxes','Exchange on-prem mailbox daily stats')
        foreach($label in @('SmartM365 inventory') + $mailboxLabels) {
            $history=Join-Path $root (([guid]::NewGuid().ToString('N')) + '/Tenants/fixture/DATA-ALL/Exchange/OnPrem/Mailboxes/WeeklyHistory')
            $week=Join-Path $history '2026-W29';$null=New-Item -ItemType Directory $week -Force
            $legacy=Join-Path $week 'manifest.json'
            $doc=[ordered]@{Week='2026-W29';HistoryLabel=$label;HistoryRootPath=$history;Files=@('Exchange_OnPrem_Mailboxes_AllDomains.csv')}
            [IO.File]::WriteAllText($legacy,($doc|ConvertTo-Json),[Text.UTF8Encoding]::new($true))
            [IO.File]::WriteAllText((Join-Path $week $doc.Files[0]),"Id`r`n42`r`n")
            $hash=(Get-FileHash $legacy).Hash
            foreach($request in $mailboxLabels) {
                $null=Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel $request
                Check ((Get-FileHash "$legacy.txt").Hash -eq $hash) 'Compatible label rewrote history.'
            }
            Check (-not (Test-Path $legacy)) 'Legacy name retained after verified transition.'
            $doc.HistoryLabel='Foreign collector'
            [IO.File]::WriteAllText("$legacy.txt",($doc|ConvertTo-Json))
            Reject {Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel $request} 'Foreign label accepted.'
        }
        foreach($scenario in @('generic','relocated','foreignTenant','foreignRoot','incomplete','foreignLabel','invalid','identical','different','resume')) {
            $history=Join-Path $root "$scenario/Tenants/fixture/DATA-ALL/M365/Licensing/Licenses/WeeklyHistory"
            $week=Join-Path $history '2026-W28';$null=New-Item -ItemType Directory $week -Force
            $legacy=Join-Path $week 'manifest.json'
            $doc=[ordered]@{Week='2026-W28';HistoryLabel='SmartM365 inventory';HistoryRootPath=$history;Files=@('M365_Licenses_Users.csv')}
            if($scenario -eq 'relocated'){$doc.HistoryRootPath='C:\old\Tenants\fixture\DATA-ALL\M365\Licensing\Licenses\WeeklyHistory'}
            if($scenario -eq 'foreignTenant'){$doc.HistoryRootPath=$history.Replace('\fixture\','\another\').Replace('/fixture/','/another/')}
            if($scenario -eq 'foreignRoot'){$doc.HistoryRootPath=$history.Replace('Licenses','Different')}
            if($scenario -eq 'foreignLabel'){$doc.HistoryLabel='Foreign'}
            [IO.File]::WriteAllText($legacy,($doc|ConvertTo-Json),[Text.UTF8Encoding]::new($true))
            [IO.File]::WriteAllText((Join-Path $week $doc.Files[0]),"Id`r`n42`r`n")
            if($scenario -eq 'incomplete'){[IO.File]::Move((Join-Path $week $doc.Files[0]),(Join-Path $week 'unrelated.csv'))}
            $hash=(Get-FileHash $legacy).Hash
            if($scenario -in 'identical','resume'){Copy-Item $legacy "$legacy.txt"}
            if($scenario -eq 'invalid'){[IO.File]::WriteAllText("$legacy.txt",'{')}
            if($scenario -eq 'different'){[IO.File]::WriteAllText("$legacy.txt",(($doc|ConvertTo-Json) + "`n"))}
            if($scenario -eq 'resume') {
                [IO.File]::WriteAllText("$legacy.migration.log",((@{Owner='WeeklyHistory:SmartM365 inventory';Phase='Prepared';SHA256=$hash}|ConvertTo-Json -Compress) + "`n"))
            }
            if($scenario -in 'generic','relocated','identical','resume') {
                $null=Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel 'M365 licenses inventory'
                Check ((Get-FileHash "$legacy.txt").Hash -eq $hash -and -not (Test-Path $legacy)) "Migration failed: $scenario"
            } else {
                Reject {Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel 'M365 licenses inventory'} "Unsafe $scenario accepted"
                Check ((Get-FileHash $legacy).Hash -eq $hash) "Unsafe $scenario destroyed old bytes"
            }
        }
    }
    if($Case -in 'Issue','All') {
        foreach($file in @('../SmartInventory/ExchangeInventory/Migration/SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1','../SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1')) {
            $definition=Definition $file 'PublishWeeklyHistory'
            # Private invocation scope reproduces script variables without executing an unsigned fixture file.
            $fixtureScript=Join-Path $root (([IO.Path]::GetFileNameWithoutExtension($file)) + '.fixture.ps1')
            $body=@(
                'param($OutputFolder,$ScriptName,$Tenant,$Definition)'
                '$ScriptVersion="fixture";$lc=@{}'
                'function CB {param($Config,$Name,$Default)$Default}'
                'function Cfg {param($Config,$Name,$Default)$Default}'
                'function Log {param($Message)}'
                'function CopyCsv {param($Source,$Destination)Copy-Item -LiteralPath $Source -Destination $Destination}'
                '. ([scriptblock]::Create($Definition))'
                'PublishWeeklyHistory @((Join-Path $OutputFolder "Inventory.csv"))'
            ) -join "`r`n"
            [IO.File]::WriteAllText($fixtureScript,$body)
            $fixtureInvocation=[scriptblock]::Create($body)
            $output=Join-Path $root ([IO.Path]::GetFileNameWithoutExtension($file));$null=New-Item -ItemType Directory $output
            [IO.File]::WriteAllText((Join-Path $output 'Inventory.csv'),"Id`r`n42`r`n")
            $scriptName=[IO.Path]::GetFileNameWithoutExtension($file)
            $result=@(& $fixtureInvocation $output $scriptName 'fixture' $definition)
            $manifest=@($result|Where-Object {$_ -like '*.json.txt'})[0]
            $hash=(Get-FileHash $manifest).Hash
            $null=& $fixtureInvocation $output $scriptName 'fixture' $definition
            Check ((Get-FileHash $manifest).Hash -eq $hash) 'Private script scope lost owner or rewrote history.'
            $doc=Get-Content $manifest -Raw|ConvertFrom-Json
            $doc.Tenant='foreign';[IO.File]::WriteAllText($manifest,($doc|ConvertTo-Json))
            Reject {& $fixtureInvocation $output $scriptName 'fixture' $definition} 'Foreign tenant accepted in private scope.'
            $doc.Tenant='fixture';$doc.ScriptName='Foreign';[IO.File]::WriteAllText($manifest,($doc|ConvertTo-Json))
            Reject {& $fixtureInvocation $output $scriptName 'fixture' $definition} 'Foreign collector accepted in private scope.'
        }
    }
    if($Case -in 'Autopatch','All') {
        Set-StrictMode -Version Latest
        $file='../SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'
        foreach($name in @('Test-AutopatchAlertMessage','Convert-FeatureRowsToAlertDetails','Convert-QualityRowsToAlertDetails','Convert-QualityErrorRowsToAlertDetails','Group-AlertSummary')) { . ([scriptblock]::Create((Definition $file $name))) }
        foreach($name in @('Convert-FeatureRowsToAlertDetails','Convert-QualityRowsToAlertDetails','Convert-QualityErrorRowsToAlertDetails')) {
            Check (@(& $name -Rows @() -PolicyMap @{}).Count -eq 0) 'Empty completed report failed.'
            Reject {& $name -Rows $null -PolicyMap @{}} 'Unnormalized null accepted.'
            $row=[pscustomobject]@{LatestAlertMessage='0';AlertMessage='0';DeviceId='device';DeviceName='fixture';PolicyId='policy';EventDateTimeUTC='2026-01-01';LastWUScanTimeUTC='2026-01-01';AggregateState='Error';CurrentDeviceUpdateStatus='Error';ExpediteQUReleaseDate='2026-01-01'}
            Check (@(& $name -Rows @($row) -PolicyMap @{}).Count -eq 0) 'No-alert sentinel emitted a row.'
            $row.LatestAlertMessage='UnknownAlert';$row.AlertMessage='UnknownAlert'
            $alerts=@(& $name -Rows @($row,$row) -PolicyMap @{policy='Fixture'})
            Check ($alerts.Count -eq 2 -and $alerts[0].DeviceId -eq 'device') 'Alert rows lost.'
            $summary=@(Group-AlertSummary -Details $alerts)
            Check ($summary.Count -eq 1 -and $summary[0].Impact -eq 2) 'Summary lost alert count.'
        }
        Check (@(Group-AlertSummary -Details @()).Count -eq 0) 'Empty alert summary failed.'

        # Exercise the actual feature/quality collection and publication phase with completed synthetic reports.
        $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot $file))
        $start=$source.IndexOf('    $detailRows = New-Object')
        $end=$source.IndexOf('    Remove-CoreSmartM365TimestampedFilesOlderThan', $start)
        Check ($start -ge 0 -and $end -gt $start) 'Autopatch collection phase not found.'
        $phase=[scriptblock]::Create($source.Substring($start,$end-$start))
        function Write-Log {param($Message,$Level)}
        function Get-FeatureUpdatePolicyMap { @{empty='Empty';alert='Alert'} }
        function Get-QualityUpdatePolicyMap { @{} }
        function Import-ExportedCsv {
            param($ReportName,$Select,$Filter)
            $script:requested += $ReportName
            if($script:reportCase -eq 'failed'){throw 'Synthetic Graph report failure'}
            if($ReportName -eq 'FeatureUpdatePolicyStatusSummary'){return @([pscustomobject]@{PolicyId='empty'},[pscustomobject]@{PolicyId='alert'})}
            if($ReportName -eq 'FeatureUpdateDeviceState' -and $Filter -match "'alert'" -and $script:reportCase -eq 'alerts') {return $script:alertFixture}
            # Completed header-only CSV: PowerShell emits no pipeline objects.
        }
        function Publish-CoreSmartM365Csv {
            param($Data,$TimestampedPath,$LatestPath,$Columns)
            $script:published += [pscustomobject]@{Path=$TimestampedPath;Rows=@($Data).Count;Columns=$Columns}
        }
        $script:alertFixture=$row
        $IncludeFeatureUpdates=$true;$IncludeQualityUpdates=$true;$MaxItems=0
        $PolicyCsvPath='policy.csv';$PolicyLatestCsvPath='latest-policy.csv';$PolicyColumns=@('PolicyId')
        $DetailCsvPath='details.csv';$DetailLatestCsvPath='latest-details.csv';$DetailColumns=@('DeviceId','AlertName')
        $SummaryCsvPath='summary.csv';$SummaryLatestCsvPath='latest-summary.csv';$SummaryColumns=@('AlertName','Impact')
        foreach($script:reportCase in @('empty','alerts','failed')) {
            $script:published=@();$script:requested=@()
            if($script:reportCase -eq 'failed') {
                Reject {& $phase} 'Graph failure converted to an empty successful report.'
                Check ($script:published.Count -eq 0) 'Failed report published partial output.'
            } else {
                & $phase
                Check ($script:requested.Count -eq 4) 'Empty feature report prevented subsequent policy/quality requests.'
                Check ($script:published.Count -eq 3) 'Missing policy, detail or summary publication.'
                $expected=if($script:reportCase -eq 'alerts'){1}else{0}
                Check (($script:published|Where-Object Path -eq 'details.csv').Rows -eq $expected) 'Detail publication count changed.'
                Check (($script:published|Where-Object Path -eq 'summary.csv').Rows -eq $expected) 'Summary publication count changed.'
                Check (@($script:published|Where-Object {$_.Columns.Count -eq 0}).Count -eq 0) 'Empty publication lost CSV schemas.'
            }
        }
        # Actual CSV importer, synthetic ZIP files and a mocked download; no Graph request.
        . ([scriptblock]::Create((Definition $file 'Import-ExportedCsv')))
        function Start-ExportJob {param($ReportName,$Select,$Filter) @{id=[guid]::NewGuid().ToString('N')} }
        function Wait-ExportJob {param($JobId) @{url='https://example.test/synthetic'} }
        function Invoke-WebRequest {param($Uri,$OutFile) Copy-Item -LiteralPath $script:fixtureZip -Destination $OutFile }
        function Ensure-Folder {param($Path)$null=New-Item -ItemType Directory -Path $Path -Force}
        # Isolate ZIP plumbing from the locally installed Archive module; test the collector importer.
        function Expand-Archive {param($Path,$DestinationPath,[switch]$Force) [IO.Compression.ZipFile]::ExtractToDirectory($Path,$DestinationPath,$true)}
        foreach($payload in @('empty','row','noHeader','missingCsv')) {
            $folder=Join-Path $root "zip-$payload";$null=New-Item -ItemType Directory $folder
            $csv=Join-Path $folder 'report.csv'
            $content=switch($payload){'empty'{"DeviceId,PolicyId`r`n"};'row'{"DeviceId,PolicyId`r`ndevice,policy`r`n"};default{''}}
            if($payload -eq 'missingCsv'){$csv=Join-Path $folder 'missing.txt'}
            [IO.File]::WriteAllText($csv,$content)
            $script:fixtureZip=Join-Path $root "$payload.zip"
            [IO.Compression.ZipFile]::CreateFromDirectory($folder,$script:fixtureZip)
            if($payload -in 'noHeader','missingCsv') {
                Reject {Import-ExportedCsv -ReportName 'Synthetic'} 'Malformed or missing CSV was accepted.'
            } else {
                $rows=@(Import-ExportedCsv -ReportName 'Synthetic')
                $expected=if($payload -eq 'row'){1}else{0}
                Check ($rows.Count -eq $expected) 'CSV importer changed row count.'
            }
        }

    }
    if($Case -in 'Mailbox','All') {
        $path=Join-Path $PSScriptRoot '../SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        $ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$null)
        $failure=$ast.Find({param($n)$n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$InventoryCompletedSuccessfully -eq $false'},$true)
        Check ($null -ne $failure) 'Mailbox final failure branch missing.'
        foreach($scenario in @('record','fallback','success')) {
            $fixture=Join-Path $root ("Mailbox-$scenario.ps1")
            $lines=@(
                'param([string]$Scenario)'
                '$StartTime=Get-Date'
                '$InventoryCompletedSuccessfully=$Scenario -eq "success"'
                '$InventoryFailureRecord=if($Scenario -eq "record"){try{throw "Synthetic inventory failure"}catch{$_}}else{$null}'
                'function Stop-SmartM365TranscriptSafely {}'
                'function Complete-SmartM365ExecutionContext {param($Status,$FailureStage,$ErrorRecord)}'
                ('if ('+$failure.Clauses[0].Item1.Extent.Text+') '+$failure.Clauses[0].Item2.Extent.Text)
            )
            [IO.File]::WriteAllText($fixture,($lines -join "`r`n"))
            foreach($engine in @((Join-Path $PSHOME 'pwsh.exe'), "$env:SystemRoot/System32/WindowsPowerShell/v1.0/powershell.exe")) {
                if(-not (Test-Path $engine)){continue}
                $result=& $engine -NoProfile -ExecutionPolicy Bypass -File $fixture $scenario 2>&1
                $expected=if($scenario -eq 'success'){0}else{1}
                Check ($LASTEXITCODE -eq $expected) "Mailbox $scenario returned false status with $engine."
                if($scenario -eq 'record'){Check (($result|Out-String) -match 'Synthetic inventory failure') 'Original failure message lost.'}
            }
        }
    }
    [pscustomobject]@{Passed=$script:passed;Case=$Case;FixtureRoot=$root;Evidence='Synthetic local fixtures only; no collection or remote access'}
} finally {& $module {param($p)Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $p} $originalPolicy}



# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVIWOHE9dgvkLI
# zcn8VkI0RNyMLZIkY+hdvz7Gc/hgMqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP2LoW623kuHH4kiUQyTd3ztee7vwzd7xFZq1uztJDA5MA0GCSqG
# SIb3DQEBAQUABIIBgB4KNe4wtqtah4I/rZ+wU3q92Ns3GqPljVuqTwXHcVjgIKIV
# X5VCtccMd2d+sZjT6p9Z8TOY5Rr1mzZE//gQaA5VA94/6+WVOeanUvwceCN4jB+P
# L4misUlf08sZLuCKGdH85snUDyqoRFUrOUw1bNorFGFDHjqc4YBIKZLoPdsfQfo/
# uRKMwVeX7fMf+NSR9SUVZ8iQOGk/zCZUPuRHb0jgM8BIfefeNFYB31rwSBfrPvTF
# XE3aqgXC55hJXuO346MyNEdv7QR+yR476ywqkspSx1/kKuDUhw19CXsvDe+Y4mIo
# 08XE9y4R7tsvNGEmPCYmSqXS6ymz3FuV36V1WTcPm/L+WtRH/d8pFFQ1Rzl1qQfK
# N8hfYmBzZ4yhdYhmXOl294/mcPHRIXAUtnHqJQYQnK7CdIZiz8ewwC9y88PLJGom
# /37JEKNkC9n8dywlj026XlFvxLMpUBDmQ5v2PbLD3R+QRlcSYkHnWdjSjd41Q8Ke
# 545vbLTxXQV+/kyVWKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjkwNzQw
# MzdaMC8GCSqGSIb3DQEJBDEiBCAKg6lqZRFFKQIEsa6k9520Cv5kKX0kKqfQt0ip
# RNJijDANBgkqhkiG9w0BAQEFAASCAgB3OJWVPD5VWqcX3cGvygEs92zzdew9zDJi
# 73JatlubEuKRQklsyCkH8sQoQJ0/U6HOo7tVf0rpHr1zL9Cvf2kogTHrP2vL88UT
# G8ZwUgVk9ePvlKE4I/eUWMZXcpvNMuOQzsVTQAqNybAk8nhFgnaJOxrlf7j16xuF
# l5BXlAC0CW3vztlgRXXO5WLmmMgQwBoGGtszqtVnmqwcUjEJWD6kvWup9srFhnuT
# oOK6unSDXp1NfF+ZSCl17/XZxSwC98vUX4DKhf6MqYRj2G6qB8QgvdVv75FzvKK9
# dzK+UiSwC3Tdrw4Iz4TT3ZkooZR9lGgsbX23wviG4KsIu658cRIKNfTjTqmTtc7n
# o5UCsOiYAc3yJ76GvcszMuWqssY9em0ILRDInZokSOVAVvaixY9TV+6FglQUoN7C
# PA5T2SRRv5k0e5GccEV88BEowcindp4Qi3mFod5pVYGExJFVk5WRhYofronc+5Mu
# 0+4lKL2TFtnFXFb+Fw0gGk/Okg1xgTV3LJrtiJ29epAwpttG5qapwa7LcL4KbXJu
# POh6DRsd5kXUqAPHoraa0HQ9I3REIOUspmFcPUX4kwnV2dRgHqkmO8F6k6uyF1Vs
# 3Mq7TGmePqJ1oWIOgedZEpnCBlHSq4O5XDloffu+CwxScWa3eqvAcS3hSomKCT4j
# y6jQgskIaA==
# SIG # End signature block
