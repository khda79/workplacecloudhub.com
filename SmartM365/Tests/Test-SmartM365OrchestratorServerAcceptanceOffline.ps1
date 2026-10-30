<#
.SYNOPSIS
    Offline checks of the orchestrator server acceptance script: dataset checks, statuses,
    module assessment, child output parsing, Graph roles and job eligibility. Nothing is read
    from a directory, a tenant or a share.
.VERSION
    1.0.1
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartInventory\Maintenance\SmartM365-OrchestratorServerAcceptance.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) { throw 'The acceptance script has PowerShell parse errors.' }
foreach ($definition in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and ($node.Name -like '*-SmartM365Acceptance*' -or $node.Name -eq 'Invoke-AcceptanceChildProcess') }, $true))) {
    . ([scriptblock]::Create($definition.Extent.Text))
}
foreach ($assignment in @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -in @('$script:GuidPattern', '$script:SmtpPattern', '$script:ResultMarker', '$script:ExpectedModules', '$script:ChildPreamble') }, $false))) {
    . ([scriptblock]::Create($assignment.Extent.Text))
}

$script:Checks = 0
function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "Assertion failed: $Message" }
    $script:Checks++
}

# Worst status and latency/clock thresholds.
Assert-True ((Get-SmartM365AcceptanceWorstStatus -Status @('PASS', 'WARN', 'SKIPPED')) -eq 'WARN') 'WARN wins over PASS and SKIPPED.'
Assert-True ((Get-SmartM365AcceptanceWorstStatus -Status @('WARN', 'FAIL')) -eq 'FAIL') 'FAIL wins.'
Assert-True ((Get-SmartM365AcceptanceWorstStatus -Status @()) -eq 'PASS') 'No status is PASS.'
Assert-True ((Get-SmartM365AcceptancePercentile -Values @(1, 2, 3, 4, 100) -Percent 50) -eq 3) 'Median of five values.'
Assert-True ((Get-SmartM365AcceptancePercentile -Values @(1, 2, 3, 4, 100) -Percent 95) -eq 100) 'p95 of five values.'
Assert-True ($null -eq (Get-SmartM365AcceptancePercentile -Values @())) 'No sample has no percentile.'
Assert-True ((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds 9 -WarningMilliseconds 15 -FailureMilliseconds 50) -eq 'PASS') 'Fast share passes.'
Assert-True ((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds 23 -WarningMilliseconds 15 -FailureMilliseconds 50) -eq 'WARN') 'The 23 ms incident share warns.'
Assert-True ((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds 50 -WarningMilliseconds 15 -FailureMilliseconds 50) -eq 'FAIL') 'Very slow share fails.'
Assert-True ((Get-SmartM365AcceptanceLatencyStatus -MedianMilliseconds $null -WarningMilliseconds 15 -FailureMilliseconds 50) -eq 'FAIL') 'Unmeasured latency fails.'
Assert-True ((Get-SmartM365AcceptanceClockStatus -OffsetSeconds -12) -eq 'PASS') 'Small clock offset passes.'
Assert-True ((Get-SmartM365AcceptanceClockStatus -OffsetSeconds 90) -eq 'WARN') 'Clock offset of 90 s warns.'
Assert-True ((Get-SmartM365AcceptanceClockStatus -OffsetSeconds -301) -eq 'FAIL') 'Clock offset beyond 5 minutes fails.'

# Dataset checks.
function New-AdRow([string]$Guid, [string]$Sam, [string]$Upn, [string]$Created = '2015-03-01T08:00:00.0000000Z') {
    [pscustomobject]@{ ObjectGUID = $Guid; SamAccountName = $Sam; UserPrincipalName = $Upn; WhenCreated = $Created }
}
$spec = Get-SmartM365AcceptanceDatasetSpec -Name 'ADUsers'
$good = @(1..5 | ForEach-Object { New-AdRow ([guid]::NewGuid().ToString()) "user$_" "user$_@example.test" })
$check = Test-SmartM365AcceptanceDataset -Rows $good -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'PASS' -and $check.Count -eq 5) "Clean AD rows pass ($($check.Issues -join '; '))."
$check = Test-SmartM365AcceptanceDataset -Rows @() -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL') 'No AD row fails.'
$check = Test-SmartM365AcceptanceDataset -Rows $good -SampleSize 4 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL' -and ($check.Issues -join ' ') -match 'limit of 4') 'More rows than the limit fails.'
$duplicate = @($good + (New-AdRow $good[0].ObjectGUID 'other' 'other@example.test'))
$check = Test-SmartM365AcceptanceDataset -Rows $duplicate -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL' -and ($check.Issues -join ' ') -match 'ObjectGUID has 1 duplicated') 'Duplicated ObjectGUID fails.'
$emptyGuid = @($good + (New-AdRow '' 'nogid' 'nogid@example.test'))
$check = Test-SmartM365AcceptanceDataset -Rows $emptyGuid -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL' -and ($check.Issues -join ' ') -match 'ObjectGUID empty in 1') 'Empty ObjectGUID fails.'
$badGuid = @($good + (New-AdRow 'not-a-guid' 'badgid' 'badgid@example.test'))
$check = Test-SmartM365AcceptanceDataset -Rows $badGuid -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'FAIL') 'Malformed required GUID fails.'
$fewUpn = @(1..4 | ForEach-Object { New-AdRow ([guid]::NewGuid().ToString()) "svc$_" '' }) + $good[0]
$check = Test-SmartM365AcceptanceDataset -Rows $fewUpn -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'WARN' -and ($check.Issues -join ' ') -match 'UserPrincipalName filled in 20 %') 'Mostly empty optional UPN warns.'
$future = @($good + (New-AdRow ([guid]::NewGuid().ToString()) 'future' 'future@example.test' '2099-01-01T00:00:00Z'))
$check = Test-SmartM365AcceptanceDataset -Rows $future -SampleSize 100 -Columns $spec.Columns
Assert-True ($check.Status -eq 'WARN' -and ($check.Issues -join ' ') -match 'implausible date') 'Future creation date warns.'
$sites = Get-SmartM365AcceptanceDatasetSpec -Name 'SharePointSites'
Assert-True ((Test-SmartM365AcceptanceDataset -Rows @() -SampleSize 10 -Columns $sites.Columns -AllowEmpty:([bool]$sites.AllowEmpty)).Status -eq 'WARN') 'An empty SharePoint farm only warns.'
$mailboxes = Get-SmartM365AcceptanceDatasetSpec -Name 'EXOMailboxes'
$mailboxRows = @([pscustomobject]@{ Guid = [guid]::NewGuid().ToString(); PrimarySmtpAddress = 'no-at-sign'; ExternalDirectoryObjectId = [guid]::NewGuid().ToString(); RecipientTypeDetails = 'UserMailbox' })
Assert-True ((Test-SmartM365AcceptanceDataset -Rows $mailboxRows -SampleSize 100 -Columns $mailboxes.Columns).Status -eq 'FAIL') 'Malformed SMTP address fails.'
foreach ($name in @('ADUsers', 'GraphUsers', 'EntraDevices', 'IntuneDevices', 'EXOMailboxes', 'OnPremMailboxes', 'SharePointSites', 'TeamsUsers', 'AdaxesUsers')) {
    Assert-True (@((Get-SmartM365AcceptanceDatasetSpec -Name $name).Columns | Where-Object { $_.PSObject.Properties['Required'] -and $_.Required }).Count -ge 1) "$name declares a required key column."
}

# Child output parsing and assessment.
$json = ConvertTo-Json -Compress -Depth 8 -InputObject @{ Ready = $true; Detail = 'forest example.test'; Datasets = @(
        @{ Name = 'ADUsers'; Scope = 'a.example.test'; Error = ''; DurationSeconds = 1.2; Rows = @($good) }
        @{ Name = 'ADUsers'; Scope = 'b.example.test'; Error = 'The server is not operational'; DurationSeconds = 21; Rows = @() }) }
$child = ConvertFrom-SmartM365AcceptanceChildOutput -StandardOutput ("WARNING: noise`r`n" + $script:ResultMarker + $json + "`r`n")
Assert-True ([bool]$child.Ready -and @($child.Datasets).Count -eq 2) 'Marker line is parsed among other output.'
$assessment = Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize 100
Assert-True ($assessment.Status -eq 'WARN' -and $assessment.Detail -match 'b\.example\.test: read failed') 'One failing domain out of two warns.'
$child.Datasets[0].Error = 'denied'
Assert-True ((Get-SmartM365AcceptanceChildAssessment -Child $child -SampleSize 100).Status -eq 'FAIL') 'Every domain failing fails.'
$missing = ConvertFrom-SmartM365AcceptanceChildOutput -StandardOutput ($script:ResultMarker + '{"Ready":false,"ComponentMissing":true,"Detail":"Adaxes PowerShell module is not installed on this server."}')
Assert-True ((Get-SmartM365AcceptanceChildAssessment -Child $missing -SampleSize 100).Status -eq 'SKIPPED') 'Missing component is skipped.'
Assert-True ((Get-SmartM365AcceptanceChildAssessment -Child $missing -SampleSize 100 -Requested).Status -eq 'FAIL') 'Missing component fails when requested.'
$broken = ConvertFrom-SmartM365AcceptanceChildOutput -StandardOutput '' -StandardError 'Import-Module : boom' -ExitCode 1
Assert-True (-not [bool]$broken.Ready -and $broken.Detail -match 'boom') 'No marker reports stderr.'
Assert-True ((Get-SmartM365AcceptanceChildAssessment -Child $broken -SampleSize 100).Status -eq 'FAIL') 'Child without result fails.'

# Steps announced by a child, and the step reported after a real time-out.
$stdout = "SMARTM365_ACCEPTANCE_STAGE:importing the Adaxes module (09:40:01)`r`nnoise`r`nSMARTM365_ACCEPTANCE_STAGE:Get-AdmUser through service adx1 on a.example.test (09:40:03)`r`n"
Assert-True ((Get-SmartM365AcceptanceLastStage -StandardOutput $stdout) -eq 'Get-AdmUser through service adx1 on a.example.test (09:40:03)') 'Last announced step is found.'
Assert-True ((Get-SmartM365AcceptanceLastStage -StandardOutput 'no stage') -eq '') 'No step gives an empty text.'
$hung = Invoke-AcceptanceChildProcess -EnginePath (Get-Process -Id $PID).Path -ScriptText "Write-AcceptanceStage 'waiting on a hung call'; Start-Sleep -Seconds 60" -TimeoutSeconds 8
Assert-True (-not [bool]$hung.Ready -and [bool]$hung.TimedOut -and $hung.Detail -match 'during: waiting on a hung call') "Time-out names the hung step ($($hung.Detail))."
$quick = Invoke-AcceptanceChildProcess -EnginePath (Get-Process -Id $PID).Path -ScriptText "Invoke-AcceptanceDataset 'X' '' { [pscustomobject]@{ A = 1 } } | Out-Null; Write-AcceptanceResult @{ Ready = `$true; Detail = 'ok'; Datasets = @(Invoke-AcceptanceDataset 'TeamsUsers' '' { [pscustomobject]@{ Identity = 'i'; UserPrincipalName = 'u@example.test' } }) }" -TimeoutSeconds 60
Assert-True ([bool]$quick.Ready -and @($quick.Datasets[0].Rows).Count -eq 1) 'Step lines do not leak into the rows of a data set.'

# Modules.
$installed = @(
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = '2.25.0' }
    [pscustomobject]@{ Name = 'Microsoft.Graph.Users'; Version = '2.37.0' }
    [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = '3.9.0' }
    [pscustomobject]@{ Name = 'ImportExcel'; Version = '7.8.10' }
    [pscustomobject]@{ Name = 'MicrosoftTeams'; Version = '7.1.0' }
    [pscustomobject]@{ Name = 'ActiveDirectory'; Version = '1.0.1.0' }
    [pscustomobject]@{ Name = 'Adaxes'; Version = '2025.1' }
)
Assert-True ((Get-SmartM365AcceptanceModuleAssessment -Installed $installed -Expected $script:ExpectedModules).Status -eq 'PASS') 'Expected modules pass.'
$oldExo = @($installed | Where-Object Name -ne 'ExchangeOnlineManagement') + [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = '3.4.0' }
$assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $oldExo -Expected $script:ExpectedModules
Assert-True ($assessment.Status -eq 'FAIL' -and $assessment.Detail -match 'ExchangeOnlineManagement 3\.4\.0 below 3\.7\.2') 'Old required module fails.'
$noAdaxes = @($installed | Where-Object Name -ne 'Adaxes')
Assert-True ((Get-SmartM365AcceptanceModuleAssessment -Installed $noAdaxes -Expected $script:ExpectedModules).Status -eq 'WARN') 'Missing optional module warns.'
$mixedGraph = @($installed) + [pscustomobject]@{ Name = 'Microsoft.Graph.Groups'; Version = '2.30.0' }
$assessment = Get-SmartM365AcceptanceModuleAssessment -Installed $mixedGraph -Expected $script:ExpectedModules
Assert-True ($assessment.Status -eq 'WARN' -and $assessment.Detail -match 'Microsoft\.Graph\.Groups 2\.30\.0') 'Graph submodule at another version warns.'

# Graph roles and job eligibility.
$jobs = @(
    [pscustomobject]@{ Name = 'Domains'; Enabled = $true; RequiredCapabilities = @('Graph', 'SharedRuntime'); RequiredGraphAppRoles = 'Directory.Read.All' }
    [pscustomobject]@{ Name = 'Devices'; Enabled = $true; RequiredCapabilities = @('Graph', 'SharedRuntime'); RequiredGraphAppRoles = @('Device.Read.All', 'Directory.Read.All') }
    [pscustomobject]@{ Name = 'Exchange2016'; Enabled = $true; RequiredCapabilities = @('AD', 'ExchangeOnPrem', 'SharedRuntime'); RequiredGraphAppRoles = @() }
    [pscustomobject]@{ Name = 'Disabled'; Enabled = $false; RequiredCapabilities = @('Graph'); RequiredGraphAppRoles = @('Policy.Read.All') }
)
$required = @(Get-SmartM365AcceptanceRequiredGraphRole -Jobs $jobs -AdditionalRoles @('Mail.Send'))
Assert-True ((@($required.Role) -join ',') -eq 'Device.Read.All,Directory.Read.All,Mail.Send') 'Roles of enabled jobs plus configuration roles.'
Assert-True ((@($required | Where-Object Role -eq 'Directory.Read.All')[0].Jobs -join ',') -eq 'Domains,Devices') 'A role lists the jobs that need it.'
$eligibility = @(Get-SmartM365AcceptanceJobEligibility -Jobs $jobs -ReadyCapabilities @('Graph', 'SharedRuntime') -GraphRoles @('Directory.Read.All'))
Assert-True ($eligibility.Count -eq 3) 'Disabled jobs are ignored.'
Assert-True ([bool]@($eligibility | Where-Object Name -eq 'Domains')[0].Eligible) 'Job with every capability and role is eligible.'
$devices = @($eligibility | Where-Object Name -eq 'Devices')[0]
Assert-True (-not $devices.Eligible -and ($devices.MissingGraphAppRoles -join ',') -eq 'Device.Read.All') 'Missing role is reported.'
$exchange = @($eligibility | Where-Object Name -eq 'Exchange2016')[0]
Assert-True (-not $exchange.Eligible -and ($exchange.MissingCapabilities -join ',') -eq 'AD,ExchangeOnPrem') 'Missing capabilities are reported.'

# Tokens.
$tokenValues = @{ WorkspaceRootPath = 'D:\Data'; TenantKey = 'contoso-test'; DataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL' }
Assert-True ((Resolve-SmartM365AcceptanceTokens -Value '{{DataAllRootPath}}' -Tokens $tokenValues) -eq 'D:\Data\Data\Tenants\contoso-test\DATA-ALL') 'Nested tokens are resolved.'
$rejected = $false
try { $null = Resolve-SmartM365AcceptanceTokens -Value '{{Unknown}}\x' -Tokens $tokenValues } catch { $rejected = $true }
Assert-True $rejected 'Unknown token is refused.'

# The script must never export data.
$text = Get-Content -LiteralPath $scriptPath -Raw
Assert-True ($text -notmatch '(?i)\bExport-Csv\b|Export-SmartM365Csv|Publish-SmartM365Csv|Out-File') 'The acceptance script writes no data file.'
Assert-True ($text -match '\$global:EnableSharePointUpload = \$false') 'The SharePoint log upload is disabled.'

# Catalog entry used by Manual launches and the GUI buttons.
$catalog = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'SmartInventory\Orchestrator\Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json
$entry = @($catalog.Launches | Where-Object Name -eq 'Orchestrator-ServerAcceptance')
Assert-True ($entry.Count -eq 1 -and $entry[0].Manual -eq $true -and $entry[0].ScriptPath -eq 'Maintenance\SmartM365-OrchestratorServerAcceptance.ps1') 'Catalog manual entry points to the script.'
Assert-True (-not $entry[0].PSObject.Properties['HideInGui']) 'The entry is visible in the GUI Manual launches.'

Write-Host ("Orchestrator server acceptance offline test passed ({0} checks)." -f $script:Checks) -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCALeUO7OP1cB87x
# EIG/80Hw6Bai0bBrE3Dv9EqTzIMmOqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILo1p8SKJz3hrLdwwrTvnI8sXpG9BpWTvIczYKiRg/WgMA0GCSqG
# SIb3DQEBAQUABIIBgHOvFaO1rpNckNUCj2NHTUr08OxFJ/C+NOHuAi8uywb7aHKv
# Y2Wz6VW7ElmlAE/a0j5OVHthPI+oqkOpZEMn4B6yYQvVa+hf6XC57l+2+/N/KDst
# q4PE3Od/cMLcVwfPw7l7boyKN2vpjiZ3JE2P73TMIxp7Uu1H3TGVbgMPesk3B0U8
# HHbFnh0bAwk91pocv8VvzmW3Vs3hn+lmMSaTQA6TQfLvu2VoAjjgG6uGsPVoJuCN
# 2CpSEOhLTsQIeUPtlBQYnYMIkUwTh3t3U1Uc3MzixeH4ciAxQzXWrUAx9LzqavTn
# UAlB5eUC5yUWQWc6PJlxfwLmAiT2q25HW4r4ndSN+UUP9EOWg8MWSIuGtq3CjC3L
# sDWHSZEy/oEhFP7s5JD+0Cmd8d9qJPkS7W+6t3i2Kv7P26DLgwic22e8JN1GVqKz
# eISLZ0ezzClC5UU2keJS3JThpnIf0f6QlWESm/vuD6FGMrp6YdQouW4/s3UxQCQw
# xMZzrb2TYbCOsRB6wKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwODM1
# MjVaMC8GCSqGSIb3DQEJBDEiBCAmlytFLO8M+icukIm87impfhv/Qe9vN8t53Erq
# HUmMUTANBgkqhkiG9w0BAQEFAASCAgBatweGlux93RD7t3EOiRVWh6zh0sciW+PD
# bzt1Im1dk4zmeKAd/k4gz0IDiPCLTf7JiTxocbMnj2VaMa1Zj4K2v2a8YD88tYrv
# rh7ZXlMAZ1MwfszF+9zIzvmzgQwIg7A+nD18X6qRym8nYbfwzzmhRYIsKSJdxEnY
# K9i27uO6xNUGsyUnI0bU7fK9YUx7EKfxEhlAyYhstjuyKUSsCpW+RVVBZ0VJzA0v
# mPvfkEwlQW+Ly6/jLPq3Eh4u5N+IK6YgePFHKpAQ7eBRhmbyblj08gKKS25yu/FA
# SY37XSgg3Pf/sbqC/xuoRsEGpUg+PHrApIRHHvPNvPDEbQf9S9YJ9ShzLcnqSh4A
# /nFdqV/le133sA3zjpUr9N2ubM+qNOi+lfzZ0eLtWlLmjHukSJSl8er64N0zKngo
# hAy46nsK4adBGMF0hCyHBp6XqEfzawZYds2SvRlsbvHV8lLm46n7qMzHe1k3eAsv
# pd/+DQ7z9IorSv4CX6nVT0rsD9R/+RkSjNiQOKNWqnZEC8U/lc8OirTZBUh+UzF8
# Gb6zYaglsz9p5gS5UTGW4RRmZfqbpm5Sa+7pPyP+N5j7nF4ortmRaA+R9QPTLdw2
# FEVStsBYA51NGzA0j/gvOeqvPJW1Nj0/yX4Crs2f8b5ufq1RTXg8FUFoviV7oX1u
# ww55T62zJw==
# SIG # End signature block
