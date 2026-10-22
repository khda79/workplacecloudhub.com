<#
.SYNOPSIS
    Synthetic identity date export tests. Never invokes a collector or imports its runtime.
.VERSION
    1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$adPath = Join-Path $root 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1'
$userPath = Join-Path $root 'SmartInventory/M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1'
$autopilotPath = Join-Path $root 'SmartInventory/M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1'
$scopePath = Join-Path $root 'SmartInventory/M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.ps1'
$checks = 0
function Assert-DateTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}
function Read-TestAst([string]$Path) {
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors)
    Assert-DateTest ($parseErrors.Count -eq 0) "Parser failed: $Path : $parseErrors"
    return $ast
}
function Import-TestFunction($Ast, [string]$Name) {
    $node = $Ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name}, $true)
    if (-not $node) { throw "Missing isolated function: $Name" }
    # Only the requested pure function is loaded, never the collector body.
    . ([scriptblock]::Create($node.Extent.Text.Replace("function $Name", "function global:$Name")))
}
$ad = Read-TestAst $adPath
$usersAst = Read-TestAst $userPath
$autopilot = Read-TestAst $autopilotPath
$scope = Read-TestAst $scopePath
$cultureBefore = [Globalization.CultureInfo]::CurrentCulture
try {
    $utc = [datetime]::new(2026, 10, 5, 9, 8, 7, [DateTimeKind]::Utc)
    $naive = [datetime]::SpecifyKind($utc, [DateTimeKind]::Unspecified)
    $offset = [datetimeoffset]::new(2026, 10, 5, 9, 8, 7, [timespan]::FromHours(2))
    $local = $utc.ToLocalTime()
    foreach ($ast in @($ad, $usersAst, $autopilot, $scope)) {
        Import-TestFunction $ast 'ConvertTo-InventoryIsoDate'
        foreach ($culture in @('fr-FR', 'en-US')) {
            [Globalization.CultureInfo]::CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($culture)
            Assert-DateTest ((ConvertTo-InventoryIsoDate $utc) -ceq '2026-10-05T09:08:07.0000000Z') 'UTC kind lost'
            Assert-DateTest ((ConvertTo-InventoryIsoDate $offset) -ceq '2026-10-05T09:08:07.0000000+02:00') 'Offset lost'
            Assert-DateTest ((ConvertTo-InventoryIsoDate $naive) -ceq '2026-10-05T09:08:07.0000000') 'Timezone invented'
            Assert-DateTest ((ConvertTo-InventoryIsoDate $local) -ceq $local.ToString('o', [Globalization.CultureInfo]::InvariantCulture)) 'Local kind changed'
            Assert-DateTest ((ConvertTo-InventoryIsoDate $null) -ceq '') 'Null changed'
            Assert-DateTest ((ConvertTo-InventoryIsoDate '05/10/2026 09:08:07') -ceq '') 'Legacy string promoted'
        }
    }
    Import-TestFunction $ad 'ConvertFrom-AdFileTimeIso'
    Assert-DateTest ((ConvertFrom-AdFileTimeIso $utc.ToFileTimeUtc()) -ceq '2026-10-05T09:08:07.0000000Z') 'FILETIME depends on host timezone'
    Assert-DateTest ((ConvertFrom-AdFileTimeIso 0) -ceq '') 'Zero FILETIME became a date'
    Assert-DateTest ((ConvertFrom-AdFileTimeIso $null) -ceq '') 'Missing FILETIME became a date'
    $invalidRejected = $false
    try { ConvertFrom-AdFileTimeIso -1 | Out-Null } catch { $invalidRejected = $true }
    Assert-DateTest $invalidRejected 'Invalid FILETIME accepted'

    Import-TestFunction $ad 'Get-SmartM365AdNativeColumns'
    $sourceUser = [pscustomobject]@{ WhenCreated=$utc; WhenChanged=$offset; AccountExpirationDate=$null;
        LastLogonTimestamp=$utc.ToFileTimeUtc(); pwdLastSet=0; badPasswordTime=$utc.ToFileTimeUtc(); PasswordExpired=$false }
    $select = $ad.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and
        $n.GetCommandName() -eq 'Select-Object' -and $n.Extent.Text.Contains("Name = 'WhenChangedIso8601'")}, $true)
    Assert-DateTest ($null -ne $select) 'AD user projection missing'
    $dateExpressions = @($select.FindAll({param($n)
        $n -is [Management.Automation.Language.HashtableAst] -and $n.Extent.Text -match "Name = '\w+Iso8601'"
    }, $true) | ForEach-Object { & ([scriptblock]::Create($_.Extent.Text)) })
    Assert-DateTest ($dateExpressions.Count -eq 6) 'AD date projection incomplete'
    $projected = $sourceUser | Select-Object -Property (@('WhenCreated','WhenChanged','PasswordExpired') + $dateExpressions)
    $csv = @($projected | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)[0]
    Assert-DateTest ($csv.WhenChangedIso8601 -ceq '2026-10-05T09:08:07.0000000+02:00') 'AD offset lost in CSV'
    Assert-DateTest ($csv.PasswordExpired -ceq 'False') 'Observed false lost'
    Assert-DateTest ($csv.AccountExpirationDateIso8601 -ceq '' -and $csv.PasswordLastSetDateIso8601 -ceq '') 'Missing/sentinel evidence changed'
    $userColumns = @(Get-SmartM365AdNativeColumns Users)
    foreach ($name in @($dateExpressions.Name) + 'PasswordExpired') {
        Assert-DateTest ($userColumns -contains $name) "AD empty schema drops $name"
    }
    Assert-DateTest ($select.Extent.Text -match '(?m)^\s+PasswordExpired\s*$') 'AD selection drops PasswordExpired'
    $getAdUser = $ad.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-ADUser' -and $n.Extent.Text.Contains('LastLogonTimestamp')}, $true)
    Assert-DateTest ($getAdUser.Extent.Text -match '\bPasswordExpired\b') 'AD query does not request PasswordExpired'

    $computer = $sourceUser
    $computerRow = [ordered]@{}
    $pcAssignments = $ad.FindAll({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left.Extent.Text -match "^\`$computerRow\['(?:\w+Iso8601|LastLogonTimestamp)'\]$"}, $true)
    Assert-DateTest ($pcAssignments.Count -eq 5) 'Computer evidence projection incomplete'
    foreach ($assignment in $pcAssignments) { . ([scriptblock]::Create($assignment.Extent.Text)) }
    Assert-DateTest ($computerRow.LastLogonTimestamp -eq $sourceUser.LastLogonTimestamp) 'Computer FILETIME dropped'
    Assert-DateTest ($computerRow.WhenCreatedIso8601.EndsWith('Z')) 'Computer creation kind lost'
    $pcColumns = @(Get-SmartM365AdNativeColumns Computers)
    foreach ($name in $computerRow.Keys) { Assert-DateTest ($pcColumns -contains $name) "Computer empty schema drops $name" }

    # Execute the actual ActiveUsers export object on a synthetic Graph-shaped user.
    $table = $usersAst.Find({param($n) $n -is [Management.Automation.Language.HashtableAst] -and
        $n.Extent.Text.Contains('"When createdIso8601"')}, $true)
    Assert-DateTest ($null -ne $table) 'ActiveUsers export object missing'
    $graph = [pscustomobject]@{CreatedDateTime=$offset;DeletedDateTime=$null;OnPremisesLastSyncDateTime=$utc;
        LastPasswordChangeDateTime=$naive;SignInActivity=[pscustomobject]@{LastSignInDateTime=$offset;LastNonInteractiveSignInDateTime=$null;LastSuccessfulSignInDateTime=$utc}}
    $licenseNames = ''; $usersCollectedAtUtc = $utc.ToString('o')
    $graphRow = $graph | ForEach-Object { [pscustomobject](& ([scriptblock]::Create($table.Extent.Text))) }
    $graphCsv = @($graphRow | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)[0]
    Assert-DateTest ($graphCsv.'When createdIso8601' -ceq '2026-10-05T09:08:07.0000000+02:00') 'Graph creation offset lost'
    Assert-DateTest ($graphCsv.'Last password change time stampIso8601' -ceq '2026-10-05T09:08:07.0000000') 'Graph unknown zone invented'
    Assert-DateTest ($graphRow.'When created' -is [datetimeoffset]) 'Legacy value was replaced'
    $emptyExport = $usersAst.Find({param($n) $n -is [Management.Automation.Language.CommandAst] -and
        $n.GetCommandName() -eq 'Export-SmartM365Csv' -and $n.Extent.Text.Contains('-Columns')}, $true)
    foreach ($name in @($graphRow.PSObject.Properties.Name | Where-Object { $_.EndsWith('Iso8601') })) {
        Assert-DateTest ($emptyExport.Extent.Text.Contains("'$name'")) "ActiveUsers empty schema drops $name"
    }
    Import-TestFunction $autopilot 'Get-SafeProperty'
    Import-TestFunction $autopilot 'Get-InventoryColumns'
    $apTable = $autopilot.Find({param($n) $n -is [Management.Automation.Language.HashtableAst] -and
        $n.Extent.Text.Contains('"Last contactedIso8601"')}, $true)
    Assert-DateTest ($null -ne $apTable) 'Autopilot date projection missing'
    $entraInfo = [pscustomobject]@{Id='synthetic-object';DeviceId='synthetic-device';DisplayName='Synthetic';ApproximateLastSignInDateTime=$offset}
    $aadDeviceId = 'synthetic-device'
    $apRow = [pscustomobject]@{Id='synthetic-autopilot';LastContactedDateTime=$utc} | ForEach-Object {
        [pscustomobject](& ([scriptblock]::Create($apTable.Extent.Text)))
    }
    $apCsv = @($apRow | Select-Object -Property (Get-InventoryColumns) | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)[0]
    Assert-DateTest ($apCsv.'Last contactedIso8601' -ceq '2026-10-05T09:08:07.0000000Z') 'Autopilot export drops contact timezone'
    Assert-DateTest ($apCsv.'Entra ApproximateLastSignInDateTimeIso8601' -ceq '2026-10-05T09:08:07.0000000+02:00') 'Autopilot export drops Entra offset'
    Assert-DateTest ($apRow.'Last contacted' -is [datetime]) 'Autopilot legacy value replaced'
    $sourceAst = Read-TestAst (Join-Path $root 'SmartInventory/Common/SmartM365.WorkplaceSource.psm1')
    Import-TestFunction $sourceAst 'Get-WorkplaceSourceValue'
    Import-TestFunction $sourceAst 'ConvertTo-WorkplacePolicyEvidence'
    $policyEvidence = ConvertTo-WorkplacePolicyEvidence -Policy @{id='synthetic-policy';createdDateTime=$offset;lastModifiedDateTime=$utc} -Family SettingsCatalog
    $dateLoop = $scope.Find({param($n) $n -is [Management.Automation.Language.ForEachStatementAst] -and
        $n.Variable.Extent.Text -eq '$dateField'}, $true)
    Assert-DateTest ($null -ne $dateLoop) 'Policy date companion projection missing'
    . ([scriptblock]::Create($dateLoop.Extent.Text))
    $policyExport = $scope.Find({param($n) $n -is [Management.Automation.Language.HashtableAst] -and
        $n.Extent.Text.Contains("Name='Intune_Policies_All'")}, $true)
    $policies = [Collections.Generic.List[object]]::new(); $policies.Add($policyEvidence)
    $policySpec = & ([scriptblock]::Create($policyExport.Extent.Text))
    $policyCsv = @($policySpec.Rows | Select-Object -Property $policySpec.Columns | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)[0]
    Assert-DateTest ($policyCsv.CreatedDateTimeIso8601 -ceq '2026-10-05T09:08:07.0000000+02:00') 'Policy creation offset lost'
    Assert-DateTest ($policyCsv.LastModifiedDateTimeIso8601 -ceq '2026-10-05T09:08:07.0000000Z') 'Policy modification offset lost'
    Assert-DateTest ($policyEvidence.CreatedDateTime -is [datetimeoffset]) 'Legacy policy date replaced'
    Write-Output "PASS: $checks identity-date checks; isolated AST functions/projections and synthetic CSV only."
} finally {
    [Globalization.CultureInfo]::CurrentCulture = $cultureBefore
    foreach ($name in @('ConvertTo-InventoryIsoDate','ConvertFrom-AdFileTimeIso','Get-SmartM365AdNativeColumns','Get-SafeProperty','Get-InventoryColumns','Get-WorkplaceSourceValue','ConvertTo-WorkplacePolicyEvidence')) {
        Remove-Item -LiteralPath "Function:\$name" -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDYGLpkkUO1k2VU
# htflbAyvorvAMn1EMdSiz1IcQeGqOaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILbNPQi08v6mqHwsPI7C1fNaLxl23R2SSHsOhIWibh8LMA0GCSqG
# SIb3DQEBAQUABIIBgGZiihknM3DKG6aDcJHHvMqTumMoBH8MapCYioUqBRyi+C0u
# XSPlswVyMXfCaoakuqY5d/ho/rjsHTzIJJrKiB8YgBvEO/kV7Tin2lzOlAdXA8wJ
# Msi7tJ3P2sY+gx+VG4lTFIKxnzXWrxokP/N8qnDCwlmL+SEKVDCKYUZp+1CSqjuZ
# izPNIh+r5LziVzy5ltLaAn9uFMBkq0qh9wvWDWBfpwU3Ag6pf8l3nqO9uSnuudSs
# zpfS3r4HSyUi/liyWfkzOhkT7Z2+RfXEiAXitzVIqoi78fLJeVtU2OGIBTIC/ng/
# pePwe2DtmmlTjsO121WzQVwQ+dO+8Qctd77f+4GrYkfiBnKYO5hlaHkrkUrR5lK7
# NAueGpoyhAeCUQcD62vqcySVZW4/W2qL7EkfN6UgzvBVi1jBEyMFlpffDgaJ3/lF
# ItoKiVFatBv/UR1ov+zo8lng6IYfleZYoJG9XbDdIT4An8elCmxgPjRyra8b7bNC
# 5hv+pcdKuLbfGX5pLKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIyMTI2
# MjlaMC8GCSqGSIb3DQEJBDEiBCB6LR/ZPItI0edR2vPNTYPHh/zDhtxf+V4t8qYq
# f+IrtjANBgkqhkiG9w0BAQEFAASCAgAZxkhQ0l2vP47YeP4efki0g7ZffTli7Io/
# /pE60MpicQ+77K7QHTr6zCK9gHqucuGSECMrk2P5J3ylST4auFtTEfyw+1xBnujT
# MVYd3qUZXiaiGf4yVQxHwCmd9MwZgd/JEWzYGRNQpKaok+vG2MmlunNXwgHdI0lK
# 2QcI5SC2vVDyQboLLckYfUozrlcGl6pbcNXJlk/8Lx8StRUJNQE6knrzTcnpeH4w
# uNHOTbVzT5mkcVwoX4GwYtt9OTiotE/HTf8DZHwCh4bTAauIGZxJ+1WpfSLBivrZ
# jBlAMpt04CtvMjwfinl0JL6imPu+njvbpRPYQFE6iPaeiU2XdoQLAXGMXMdbmjDM
# tJTyqBlbAH3TfpZp5FEK+ht2AH6Xpjti7GQ3+48cLlySPZ7lyRAVGM+JCVGBOoer
# 4aaAcsLESDKzToCcXSc4bqr2/FsingLrXTnaYh2VmISXajP9L3VV/S1GGg/8Lgxw
# X5pAtAoKJBPtOqYb/BK394Q6R7D+XTTFnyELdsTYhadhN5TffijPKjQp0FXFaCbo
# bBRRtp0I1YYOSOSSlsieCzNM7qPGOciK6Vj7Pqk4PbCvh0gCU5MKluL72hvCsisn
# o5MLu2m0DE4UZ+aFCKw4zqFSislTMJ5jqaQO2+8n1/00ESHB7k/HoNnjPEirYlYO
# m5hUkbtkcQ==
# SIG # End signature block
