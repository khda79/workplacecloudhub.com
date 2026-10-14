[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Synthetic XLSX metadata and download callbacks only; no Graph or SharePoint connection.
$ErrorActionPreference='Stop'
$script:checks=0
function Check([bool]$Condition,[string]$Message){$script:checks++;if(-not $Condition){throw "Assertion failed: $Message"}}
Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force
$root=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
function Write-WorkbookFixture([string]$Path,[string[]]$Sheets) {
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $archive=[IO.Compression.ZipFile]::Open($Path,[IO.Compression.ZipArchiveMode]::Create)
    try {
        $entry=$archive.CreateEntry('xl/workbook.xml')
        $writer=[IO.StreamWriter]::new($entry.Open())
        try {
            $nodes=($Sheets | ForEach-Object {'<sheet name="'+$_+'" />'}) -join ''
            $writer.Write('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets>'+$nodes+'</sheets></workbook>')
        } finally {$writer.Dispose()}
    } finally {$archive.Dispose()}
}
$fixtures=Join-Path $root 'fixtures'
New-Item -ItemType Directory -Path $fixtures | Out-Null
$persona='SmartWorkplaceIntelligence-PersonaClassification.xlsx'
$site='SmartWorkplaceIntelligence-SiteClassification.xlsx'
Write-WorkbookFixture (Join-Path $fixtures $persona) @('Personas','Rules','Exclusions')
Write-WorkbookFixture (Join-Path $fixtures $site) @('Settings','SiteTypes','Sites')
$download={param($destination,$name) Copy-Item -LiteralPath (Join-Path $fixtures $name) -Destination $destination; Get-Item -LiteralPath $destination}.GetNewClosure()
$work=Join-Path $root 'data'
New-Item -ItemType Directory -Path $work | Out-Null
$accepted=Receive-PreparedMappingWorkbooks -DataRoot $work -DownloadFile $download
if(@(Get-ChildItem -LiteralPath $accepted -File).Count -ne 2){throw 'Expected two downloaded workbooks'}
$originalHashes=@{}
foreach($name in @($persona,$site)){
    $originalHashes[$name]=(Get-FileHash -LiteralPath (Join-Path $accepted $name)).Hash
    Check ((Get-FileHash -LiteralPath (Join-Path $work $name)).Hash -eq $originalHashes[$name]) ('Retained DATA copy matches validated download '+$name)
}
Check ($accepted.StartsWith($work+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) 'Download staging lives below DATA, not WorkRoot/Temp'
Write-WorkbookFixture (Join-Path $fixtures $persona) @('Personas','Rules','Exclusions','Revision2')
Write-WorkbookFixture (Join-Path $fixtures $site) @('Settings','SiteTypes','Sites','Revision2')
$second=Receive-PreparedMappingWorkbooks -DataRoot $work -DownloadFile $download
if($accepted -eq $second){throw 'Run mappings must be isolated'}
foreach($name in @($persona,$site)){
    Check ((Get-FileHash -LiteralPath (Join-Path $work $name)).Hash -eq (Get-FileHash -LiteralPath (Join-Path $second $name)).Hash) ('Next run replaces retained copy '+$name)
    Check ((Get-FileHash -LiteralPath (Join-Path $work $name)).Hash -ne $originalHashes[$name]) ('Retained revision changed '+$name)
    Check ((Get-FileHash -LiteralPath (Join-Path $accepted $name)).Hash -eq $originalHashes[$name]) ('First run keeps its own immutable pair '+$name)
}
function Expect-Rejection([string]$Name,[scriptblock]$Download) {
    $before=@{};foreach($file in @($persona,$site)){$before[$file]=(Get-FileHash -LiteralPath (Join-Path $work $file)).Hash}
    $folderCount=@(Get-ChildItem -LiteralPath (Join-Path $work 'mappings') -Directory).Count
    $failed=$false
    try {$null=Receive-PreparedMappingWorkbooks -DataRoot $work -DownloadFile $Download} catch {$failed=$true}
    if(-not $failed){throw "Expected rejection: $Name"}
    if(@(Get-ChildItem -LiteralPath $accepted -File).Count -ne 2){throw 'Previous files changed'}
    foreach($file in @($persona,$site)){Check ((Get-FileHash -LiteralPath (Join-Path $work $file)).Hash -eq $before[$file]) ('Failure preserves retained workbook '+$Name+'/'+$file)}
    Check (@(Get-ChildItem -LiteralPath (Join-Path $work 'mappings') -Directory).Count -eq $folderCount) ('Failure removes only its intermediate folder '+$Name)
}
Expect-Rejection 'null receipt, even with previous cache' {param($destination,$name) $null}
Expect-Rejection 'receipt without file' {param($destination,$name) $true}
Expect-Rejection 'invalid workbook/HTML' {param($destination,$name) [IO.File]::WriteAllText($destination,'<html>sign in</html>'); Get-Item -LiteralPath $destination}
$wrong=Join-Path $fixtures 'wrong.xlsx'
Write-WorkbookFixture $wrong @('Other')
Expect-Rejection 'missing required worksheet' ({param($destination,$name) Copy-Item -LiteralPath $wrong -Destination $destination; Get-Item -LiteralPath $destination}.GetNewClosure())
Expect-Rejection 'second download failed' ({
    param($destination,$name)
    if($name -eq $site){return $null}
    Copy-Item -LiteralPath (Join-Path $fixtures $name) -Destination $destination
    Get-Item -LiteralPath $destination
}.GetNewClosure())
$data=Join-Path $root 'source-selection'
New-Item -ItemType Directory -Path $data | Out-Null
# An unrelated raw-root workbook must never override the freshly downloaded mapping.
[IO.File]::WriteAllText((Join-Path $data $persona),'stale fixture')
$contract=Join-Path $root 'source-contract.json'
@{currentFiles=@();mappingFiles=@($persona,$site);dailyFiles=@();history=@()} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $contract
$plan=@(Get-PreparedSourcePlan -DataRoot $data -SourceContractPath $contract -MappingRoot $accepted)
if($plan.Count -ne 2){throw 'Source plan omitted mapping files'}
foreach($entry in $plan){
    if($entry.SourcePath -ne (Join-Path $accepted $entry.Relative)){throw 'Source plan used raw-root fallback'}
    if($entry.SHA256 -ne (Get-FileHash -LiteralPath $entry.SourcePath).Hash){throw 'Mapping provenance hash mismatch'}
}
# Simulate deterministic replacement errors without filesystem timing or any live service.
$module=Get-Module PreparedEvidencePipeline
$moveOriginal=& $module { (Get-Item Function:Move-PreparedFile).ScriptBlock }
try {
    & $module {param($siteName) $script:failMappingSite=$siteName
        function script:Move-PreparedFile {param($Source,$Destination,[switch]$Overwrite)
            if((Split-Path $Source -Parent) -like '*promotion' -and (Split-Path $Destination -Leaf) -eq $script:failMappingSite){throw 'Synthetic second mapping move failed'}
            [IO.File]::Move($Source,$Destination,[bool]$Overwrite)
        }
    } $site
    Write-WorkbookFixture (Join-Path $fixtures $persona) @('Personas','Rules','Exclusions','Revision3')
    Write-WorkbookFixture (Join-Path $fixtures $site) @('Settings','SiteTypes','Sites','Revision3')
    Expect-Rejection 'second replacement rolls back first' $download
    $firstInstall=Join-Path $root 'first-install';New-Item -ItemType Directory -Path $firstInstall | Out-Null
    $rejected=$false;try{$null=Receive-PreparedMappingWorkbooks -DataRoot $firstInstall -DownloadFile $download}catch{$rejected=$true}
    Check $rejected 'Failed initial pair installation is rejected'
    foreach($name in @($persona,$site)){Check (-not(Test-Path -LiteralPath (Join-Path $firstInstall $name))) ('Initial replacement failure leaves no partial retained pair '+$name)}
} finally { & $module {param($original) Set-Item Function:script:Move-PreparedFile -Value $original} $moveOriginal }

# A live capture lock refuses replacement before any download or workbook change.
$captureLock=[IO.File]::Open((Join-Path $work '.prepared-source.lock'),'OpenOrCreate','ReadWrite','None')
$blockedDownload=[pscustomobject]@{Called=$false}
try {Expect-Rejection 'capture lock' ({param($destination,$name) $blockedDownload.Called=$true;throw 'Download callback must not be invoked'}.GetNewClosure())} finally {$captureLock.Dispose()}
Check (-not $blockedDownload.Called) 'Capture lock rejects before download callback'
$nested=[pscustomobject]@{Rejected=$false}
$nestedDownload={param($destination,$name)
    if(-not $nested.Rejected){try{$null=Receive-PreparedMappingWorkbooks -DataRoot $work -DownloadFile $download}catch{$nested.Rejected=$true}}
    & $download $destination $name
}.GetNewClosure()
$third=Receive-PreparedMappingWorkbooks -DataRoot $work -DownloadFile $nestedDownload
Check $nested.Rejected 'Concurrent nested mapping run refused'
foreach($name in @($persona,$site)){Check ((Get-FileHash -LiteralPath (Join-Path $work $name)).Hash -eq (Get-FileHash -LiteralPath (Join-Path $third $name)).Hash) ('Concurrent attempt does not corrupt final pair '+$name)}
# Successful online acquisition also installs both files when DATA has no prior pair.
$fresh=Join-Path $root 'fresh-data';New-Item -ItemType Directory -Path $fresh | Out-Null
$freshPair=Receive-PreparedMappingWorkbooks -DataRoot $fresh -DownloadFile $download
Remove-PreparedMappingWorkbooks -DataRoot $fresh -MappingRoot $freshPair
foreach($name in @($persona,$site)){Check (Test-Path -LiteralPath (Join-Path $fresh $name) -PathType Leaf) ('Persistent workbook survives first-run cleanup '+$name)}
foreach($mapping in @($accepted,$second,$third)){Remove-PreparedMappingWorkbooks -DataRoot $work -MappingRoot $mapping}
foreach($name in @($persona,$site)){Check ((Get-FileHash -LiteralPath (Join-Path $work $name)).Hash -eq (Get-FileHash -LiteralPath (Join-Path $fixtures $name)).Hash) ('Persistent workbook survives all run cleanup '+$name)}
$cleanupRejected=$false;try{Remove-PreparedMappingWorkbooks -DataRoot $work -MappingRoot $work}catch{$cleanupRejected=$true}
Check $cleanupRejected 'Mapping cleanup refuses DATA root'
Check (@(Get-ChildItem -LiteralPath (Join-Path $work 'mappings') -Directory).Count -eq 0) 'No completed-run workbook history left in DATA'
# If rollback itself fails, preserve recovery bytes and stop instead of deleting the backups.
try {
    & $module {param($siteName) $script:failMappingSite=$siteName
        function script:Move-PreparedFile {param($Source,$Destination,[switch]$Overwrite)
            if((Split-Path $Source -Parent) -like '*rollback' -or ((Split-Path $Source -Parent) -like '*promotion' -and (Split-Path $Destination -Leaf) -eq $script:failMappingSite)){throw 'Synthetic move/rollback failure'}
            [IO.File]::Move($Source,$Destination,[bool]$Overwrite)
        }
    } $site
    $prior=@{};foreach($name in @($persona,$site)){$prior[$name]=(Get-FileHash -LiteralPath (Join-Path $work $name)).Hash}
    Write-WorkbookFixture (Join-Path $fixtures $persona) @('Personas','Rules','Exclusions','Revision4')
    Write-WorkbookFixture (Join-Path $fixtures $site) @('Settings','SiteTypes','Sites','Revision4')
    $recoveryFailure=$null;try{$null=Receive-PreparedMappingWorkbooks -DataRoot $work -DownloadFile $download}catch{$recoveryFailure=$_}
    Check ($null -ne $recoveryFailure -and $recoveryFailure.Exception.Data['KeepMappingRecovery']) 'Failed rollback is explicitly fatal and retains recovery files'
    $recoveryFolders=@(Get-ChildItem -LiteralPath (Join-Path $work 'mappings') -Directory)
    Check ($recoveryFolders.Count -eq 1) 'Only failed rollback folder retained'
    foreach($name in @($persona,$site)){
        $backup=Join-Path $recoveryFolders[0].FullName ('rollback/'+$name)
        Check ((Get-FileHash -LiteralPath $backup).Hash -eq $prior[$name]) ('Recovery backup preserves previous bytes '+$name)
        # Restore synthetic fixtures only, proving the retained rollback bytes are usable.
        Copy-Item -LiteralPath $backup -Destination (Join-Path $work $name) -Force
        Check ((Get-FileHash -LiteralPath (Join-Path $work $name)).Hash -eq $prior[$name]) ('Synthetic recovery byte verification '+$name)
    }
    Remove-PreparedMappingWorkbooks -DataRoot $work -MappingRoot $recoveryFolders[0].FullName
} finally { & $module {param($original) Set-Item Function:script:Move-PreparedFile -Value $original} $moveOriginal }
# Check shared CSV-to-DATA root normalization with neutral fixture values.
$repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$tokens=$null;$errors=$null
$entryAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'SmartM365/SmartInventory/PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1'),[ref]$tokens,[ref]$errors)
Check ($errors.Count -eq 0) 'Preparation entry point parses without execution'
$mappingCalls=@($entryAst.FindAll({param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Receive-PreparedMappingWorkbooks'},$true))
Check ($mappingCalls.Count -eq 1 -and $mappingCalls[0].Extent.Text -match '-DataRoot \$data' -and $mappingCalls[0].Extent.Text -notmatch '-WorkRoot') 'Preparation passes tenant DATA to mapping acquisition'
$cleanupCalls=@($entryAst.FindAll({param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Remove-PreparedMappingWorkbooks'},$true))
Check ($cleanupCalls.Count -eq 1 -and $cleanupCalls[0].Extent.Text -match '-DataRoot \$data') 'Preparation cleans DATA intermediates only'
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'SmartM365/Modules/SmartM365.Core/SmartM365.Core.psm1'),[ref]$tokens,[ref]$errors)
$fn=$ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$true) | Where-Object Name -eq 'ConvertTo-SmartM365SharePointDataRootPath'
. ([scriptblock]::Create($fn.Extent.Text))
if((ConvertTo-SmartM365SharePointDataRootPath 'Example/CSV') -ne 'Example/DATA'){throw 'Shared source root normalization changed'}
Write-Host "PASS: $script:checks persistent DATA mapping, replacement rollback, concurrency, cleanup and provenance checks; synthetic workbooks/callbacks only. No network access."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDeclaGpB3Y37qT
# alKyp7dj4HCQtE4CQTPpN4H/UyZEPaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINJeoU8LrW+ENAehQ0YWUJiPPy3JFsPVFW54RwckaHDdMA0GCSqG
# SIb3DQEBAQUABIIBgBUN4bAXDn6DXSVOWumJPwOjeAT03mT8+92BmTia8FlUIJFv
# Vt+ZmH2cqKOsDdlHwJa9lg/ggi7aGJN8t5EeJEvDBKW2/cg5ZgkdBlDqosQzWdzG
# j/FDKXCT2X5/ex8x8EyrbDERBVfcTBEWuUV+XRTZv21mQTe+VhQtFdW5E61LDDp2
# 1SgD7j95NCVkblScqN/uSzmNxKNRdyan+3XFkSR/ZMfP8XPHvuq0KfOWFw1h+jO+
# ciURZaZ69u9oCA7rxJLLad1yYY8ovLM+r0tk5AsBWXCNEGHJJRc2GM1Mcp0IQj3C
# C4/W0ywfPNB47pCXJia/59lhHb1vihvRbLDg6L6hrTkEwM8koyG/jjmuwltjIEA9
# v/6NMT1FAPFPdDm4FSlY7JyDBFAcqC3RmmM0Qda49yvcffghhOzfQJg3TsfOFYw/
# Ywmyk0v5q0pXF/9tazGl5/6kMHV7apsX2hJB7CXdMSw6ss5Pa0ha4R2/tWsPKKDA
# VUXSLj46D/l9XX41aKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTQxMjIz
# MzRaMC8GCSqGSIb3DQEJBDEiBCAv5tngytZ1uClvqSoQuVSPLqmqcMmrgYW6Tpmh
# tTI4NDANBgkqhkiG9w0BAQEFAASCAgCRY55JjxnFtLu95sU3mXKhcq9xQ/B+nbIU
# YgXb5vxgJPK4RUHZ0iRsBKkgToAMcfsxYyaz3TQnXe1VhlNQX2QcQIAUV4jI7lGu
# +eoMy6jPeiG6trYMEdW2g/6IMW/q4by/R7miaeNJ6z5V40SmRzxbgqlCbT98Uuw9
# 41tVDOpkOHwAXNq8fX5o36dyRNHDV4588GCxyj6Hc9gPylXAso+efnBEEWigpEHa
# O7DuQAkX7jqX9aiZe/dfw8eD3j8odcvrX5O8Hn8CGLE8q7BFgU3B0jChB/DOI8Jn
# YSWsdGwHjiv24+JXHQLqcnCRSnXHJW1ehk95awSeq+ewOiVf9OseN2mfCF30rs3O
# C5vwL9+kp0+P6/856/BzP3MXNaM5Ah2N7MNJpNzPUvPjv2govtr0bNGUa+cr0pFa
# FKemPZF3ueagcar2roMwrt4FZI9iTMdD6K/BhXwCFEWlzKnHabU/o892rrmN/fw1
# N4t8f4RvfnkDikj089+P42aw2AqlokxsvQ7mEmZ3tm4HfYUDWrpKmghL3XQex8a0
# DncbqH47qGsmKNLGQlFLiw7w5zt6dvQ0JXFnD+U9O38ppf2olzFM9+zgCO3uFulM
# kQq+EJJ98T2xIMa734m4PYKmaRvr2e8r/PFwbJUje6WXsdOqZ/eqksa6SphdvP/s
# cbzdN4KAIQ==
# SIG # End signature block
