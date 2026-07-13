<#
.SYNOPSIS
    Checks SmartM365 PowerShell script standards before commit/push.

.DESCRIPTION
    By default, this guard checks only changed SmartM365 scripts so legacy issues do not hide new regressions.
    Use -Scope All to scan all SmartM365 scripts, including SmartInventory, Setup, Communications,
    SharePointMigration, Devices, Intune Remediation, and future SmartM365 folders.

.VERSION
1.5
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [ValidateSet('Changed', 'All')]
    [string]$Scope = 'Changed',

    [switch]$WarningsAsErrors
)

$ErrorActionPreference = 'Stop'

function Find-SmartM365Root {
    [CmdletBinding()]
    param([string]$StartPath)

    $current = if ([string]::IsNullOrWhiteSpace($StartPath)) { (Get-Location).Path } else { $StartPath }
    while ($current) {
        if ((Test-Path -LiteralPath (Join-Path -Path $current -ChildPath 'Config\SmartM365.global.local.json.template')) -and
            (Test-Path -LiteralPath (Join-Path -Path $current -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'))) {
            return $current
        }

        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }

    throw "SmartM365 root not found from '$StartPath'."
}

function Get-RepositoryRoot {
    [CmdletBinding()]
    param([string]$StartPath)

    $output = & git -C $StartPath rev-parse --show-toplevel 2>$null
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace([string]$output)) {
        return ''
    }

    return [string]$output
}

function ConvertTo-RelativePath {
    [CmdletBinding()]
    param(
        [string]$RootPath,
        [string]$Path
    )

    $rootFull = [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\', '/')
    $pathFull = [System.IO.Path]::GetFullPath($Path)
    if ($pathFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $pathFull.Substring($rootFull.Length).TrimStart('\', '/')
    }

    return $Path
}

function Get-SmartM365ChangedFiles {
    [CmdletBinding()]
    param(
        [string]$RepositoryRoot,
        [string]$SmartM365Root
    )

    $changes = @{}
    if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) { return $changes }

    $smartRelative = ConvertTo-RelativePath -RootPath $RepositoryRoot -Path $SmartM365Root
    $statusLines = @(& git -C $RepositoryRoot status --porcelain --untracked-files=all -- $smartRelative 2>$null)
    if ($LASTEXITCODE -ne 0) { return $changes }

    foreach ($line in $statusLines) {
        if ([string]::IsNullOrWhiteSpace($line) -or $line.Length -lt 4) { continue }

        $status = $line.Substring(0, 2)
        $relativePath = $line.Substring(3).Trim()
        if ($relativePath -match ' -> ') {
            $parts = $relativePath -split ' -> '
            $relativePath = $parts[$parts.Count - 1]
        }
        $relativePath = $relativePath.Trim('"')
        if ([string]::IsNullOrWhiteSpace($relativePath)) { continue }

        $fullPath = Join-Path -Path $RepositoryRoot -ChildPath $relativePath
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) { continue }

        $changes[[System.IO.Path]::GetFullPath($fullPath).ToLowerInvariant()] = [pscustomobject]@{
            Status       = $status
            RelativePath = $relativePath
            FullPath     = [System.IO.Path]::GetFullPath($fullPath)
        }
    }

    return $changes
}

function Get-HeaderVersion {
    [CmdletBinding()]
    param([string]$Content)

    $match = [regex]::Match($Content, '(?im)^[ \t]*\.VERSION(?:[ \t]+(?<InlineVersion>[^\r\n]+)|[ \t]*\r?\n[ \t]*(?<Version>[^\r\n]+))')
    if ($match.Success -and -not [string]::IsNullOrWhiteSpace($match.Groups['InlineVersion'].Value)) {
        return $match.Groups['InlineVersion'].Value.Trim()
    }
    if ($match.Success) { return $match.Groups['Version'].Value.Trim() }
    return ''
}

function Get-RuntimeVersion {
    [CmdletBinding()]
    param([string]$Content)

    $match = [regex]::Match($Content, '(?m)^\s*\$ScriptVersion\s*=\s*["''](?<Version>[^"'']+)["'']')
    if ($match.Success) { return $match.Groups['Version'].Value.Trim() }
    return ''
}

function Get-HeadFileContent {
    [CmdletBinding()]
    param(
        [string]$RepositoryRoot,
        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($RepositoryRoot) -or [string]::IsNullOrWhiteSpace($RelativePath)) { return $null }

    $gitPath = $RelativePath -replace '\\', '/'
    $output = @(& git -C $RepositoryRoot show ("HEAD:{0}" -f $gitPath) 2>$null)
    if ($LASTEXITCODE -ne 0) { return $null }

    return ($output -join [Environment]::NewLine)
}

function Add-StandardResult {
    [CmdletBinding()]
    param(
        [System.Collections.Generic.List[object]]$Results,
        [ValidateSet('ERROR', 'WARNING')]
        [string]$Severity,
        [string]$Rule,
        [string]$Path,
        [string]$Message
    )

    $Results.Add([pscustomobject]@{
        Severity = $Severity
        Rule     = $Rule
        Path     = $Path
        Message  = $Message
    }) | Out-Null
}

$script:SmartM365Root = Find-SmartM365Root -StartPath $PSScriptRoot
$repositoryRoot = Get-RepositoryRoot -StartPath $script:SmartM365Root
$changedFiles = Get-SmartM365ChangedFiles -RepositoryRoot $repositoryRoot -SmartM365Root $script:SmartM365Root

if ($Scope -eq 'Changed') {
    $scriptFiles = @(
        foreach ($entry in $changedFiles.Values) {
            if ($entry.FullPath -like '*.ps1') { Get-Item -LiteralPath $entry.FullPath }
        }
    ) | Sort-Object FullName -Unique
}
else {
    $scriptFiles = Get-ChildItem -LiteralPath $script:SmartM365Root -Recurse -Filter '*.ps1' -File | Sort-Object FullName
}

$results = New-Object 'System.Collections.Generic.List[object]'

foreach ($scriptFile in $scriptFiles) {
    $content = [System.IO.File]::ReadAllText($scriptFile.FullName)
    $relativePath = if ($repositoryRoot) { ConvertTo-RelativePath -RootPath $repositoryRoot -Path $scriptFile.FullName } else { ConvertTo-RelativePath -RootPath $script:SmartM365Root -Path $scriptFile.FullName }
    $changeKey = [System.IO.Path]::GetFullPath($scriptFile.FullName).ToLowerInvariant()
    $change = if ($changedFiles.ContainsKey($changeKey)) { $changedFiles[$changeKey] } else { $null }
    $isChanged = $null -ne $change
    $isNew = $isChanged -and ($change.Status -eq '??' -or $change.Status -match 'A')

    $headerVersion = Get-HeaderVersion -Content $content
    $runtimeVersion = Get-RuntimeVersion -Content $content

    if ([string]::IsNullOrWhiteSpace($headerVersion)) {
        $severity = if ($Scope -eq 'Changed' -and $isChanged) { 'ERROR' } else { 'WARNING' }
        Add-StandardResult -Results $results -Severity $severity -Rule 'MissingVersionHeader' -Path $relativePath -Message 'PowerShell scripts should declare a .VERSION block.'
    }

    if (-not [string]::IsNullOrWhiteSpace($headerVersion) -and -not [string]::IsNullOrWhiteSpace($runtimeVersion) -and $headerVersion -ne $runtimeVersion) {
        Add-StandardResult -Results $results -Severity 'ERROR' -Rule 'VersionMismatch' -Path $relativePath -Message (".VERSION '$headerVersion' does not match `$ScriptVersion '$runtimeVersion'.")
    }

    if ($Scope -eq 'Changed' -and $isChanged -and -not $isNew -and -not [string]::IsNullOrWhiteSpace($headerVersion)) {
        $headContent = Get-HeadFileContent -RepositoryRoot $repositoryRoot -RelativePath $change.RelativePath
        if ($null -ne $headContent) {
            $headVersion = Get-HeaderVersion -Content $headContent
            if (-not [string]::IsNullOrWhiteSpace($headVersion) -and $headVersion -eq $headerVersion) {
                Add-StandardResult -Results $results -Severity 'ERROR' -Rule 'VersionNotBumped' -Path $relativePath -Message ("Script changed but .VERSION stayed at '$headerVersion'.")
            }
        }
    }

    $usesInitializeScriptEnvironment = $content -match '\bInitializeScriptEnvironment\b'
    $usesCoreInitializeScriptEnvironment = $content -match '\bCoreInitializeScriptEnvironment\b'
    $usesCoreContextInitialization = $content -match '\bSet-SmartM365CoreContext\b'
    $usesSharedScriptInitialization = $usesInitializeScriptEnvironment -or $usesCoreInitializeScriptEnvironment -or $usesCoreContextInitialization
    $usesExecutionSummary = $content -match '\bComplete-SmartM365ExecutionContext\b'
    $usesSmartM365Logging = $content -match '\bWriteLog\b|\bWriteLogSmartM365\b|\bStart-Transcript\b|\$global:LogTextFile'

    $normalizedRelativePath = $relativePath.Replace('/', '\').ToLowerInvariant()
    $isSmartInventoryScript = $normalizedRelativePath.StartsWith('smartm365\smartinventory\')
    $startupBannerExclusions = @(
        'smartm365\smartinventory\activedirectoryinventory\smartm365-activedirectory-enrichedcolumns.ps1',
        'smartm365\smartinventory\activedirectoryinventory\smartm365-activedirectory-enrichment.ps1',
        'smartm365\smartinventory\activedirectoryinventory\smartm365-activedirectory-usersenrichment.ps1',
        'smartm365\smartinventory\m365inventory\intuneinventory\devices\smartm365-detect-devicesysteminfo.ps1'
    )
    $isStartupBannerExcluded = $normalizedRelativePath -in $startupBannerExclusions
    if ($isSmartInventoryScript -and -not $isStartupBannerExcluded) {
        $usesTenantContext = $content -match '\bInitialize-SmartM365TenantContext\b'
        $usesExplicitStartupBanner = $content -match '\bWrite-SmartM365StartupBanner\b'
        $allowsEarlyExplicitStartupBanner = $normalizedRelativePath -eq 'smartm365\smartinventory\orchestrator\smartm365-inventory-orchestrator.ps1'
        $severity = if ($Scope -eq 'Changed' -and $isChanged) { 'ERROR' } else { 'WARNING' }

        if (-not $usesTenantContext -and -not $usesExplicitStartupBanner) {
            Add-StandardResult -Results $results -Severity $severity -Rule 'MissingStartupBrandBanner' -Path $relativePath -Message 'Human-facing SmartInventory entry scripts must initialize the tenant context or call Write-SmartM365StartupBanner explicitly.'
        }
        if ($usesTenantContext -and $usesExplicitStartupBanner -and -not $allowsEarlyExplicitStartupBanner) {
            Add-StandardResult -Results $results -Severity $severity -Rule 'RedundantStartupBrandBanner' -Path $relativePath -Message 'Initialize-SmartM365TenantContext already displays the startup banner; remove the explicit banner call.'
        }
        if ($content -match 'SmartM365 by WorkplaceCloudHub') {
            Add-StandardResult -Results $results -Severity $severity -Rule 'HardcodedStartupBrandBanner' -Path $relativePath -Message 'Do not hard-code the WorkplaceCloudHub banner in an entry script; use the shared startup banner.'
        }
    }
    if ($usesSmartM365Logging -and -not $usesSharedScriptInitialization) {
        Add-StandardResult -Results $results -Severity 'WARNING' -Rule 'LoggingWithoutSharedInitialization' -Path $relativePath -Message 'Script uses SmartM365-style logging but does not call a shared SmartM365 script initialization helper.'
    }
    if ($usesInitializeScriptEnvironment -and -not $usesExecutionSummary) {
        $severity = if ($Scope -eq 'Changed' -and $isChanged) { 'ERROR' } else { 'WARNING' }
        Add-StandardResult -Results $results -Severity $severity -Rule 'MissingExecutionSummary' -Path $relativePath -Message 'Script calls InitializeScriptEnvironment but does not call Complete-SmartM365ExecutionContext in its cleanup path.'
    }

    foreach ($line in ($content -split "\r?\n")) {
        if ($line -notmatch '\bImport-Module\b') { continue }
        $isSmartM365InternalImport = $line -match 'SmartM365\.Core\.psd1|SmartM365-WindowsPowerShell5\.psd1|\$modulePath|\$ModulePath|\$CoreModulePath|\$coreModulePath|\$candidate|Join-ModulePath'
        if ($isSmartM365InternalImport -and $line -notmatch '-MinimumVersion') {
            $severity = if ($Scope -eq 'Changed' -and $isChanged) { 'ERROR' } else { 'WARNING' }
            Add-StandardResult -Results $results -Severity $severity -Rule 'MissingSmartM365ModuleMinimumVersion' -Path $relativePath -Message 'SmartM365 internal module imports must use -MinimumVersion so copied scripts fail fast with stale modules.'
        }
    }
    $expectsScriptLocalJson = ($content -match '\.local\.json') -and ($content -match 'GetFileNameWithoutExtension\(\$PSCommandPath\)')
    if ($expectsScriptLocalJson) {
        $expectedTemplatePath = Join-Path -Path $scriptFile.DirectoryName -ChildPath ("{0}.local.json.template" -f $scriptFile.BaseName)
        if (-not (Test-Path -LiteralPath $expectedTemplatePath -PathType Leaf)) {
            $severity = if ($Scope -eq 'Changed' -and $isChanged) { 'ERROR' } else { 'WARNING' }
            Add-StandardResult -Results $results -Severity $severity -Rule 'MissingScriptLocalJsonTemplate' -Path $relativePath -Message ("Expected template is missing: {0}" -f (ConvertTo-RelativePath -RootPath $repositoryRoot -Path $expectedTemplatePath))
        }
    }

    if ($usesInitializeScriptEnvironment) {
        $manualContextPatterns = @(
            'PowerShell Version\s*:',
            'WindowsIdentity\]::GetCurrent',
            '\$env:COMPUTERNAME',
            '\$env:USERNAME',
            '\[Environment\]::MachineName',
            'RunAsAdmin'
        )
        $manualContextLogLine = $false
        foreach ($line in ($content -split "\r?\n")) {
            if ($line -notmatch '\bWriteLog\b|\bWriteLogSmartM365\b') { continue }
            foreach ($pattern in $manualContextPatterns) {
                if ($line -match $pattern) {
                    $manualContextLogLine = $true
                    break
                }
            }
            if ($manualContextLogLine) { break }
        }
        if ($manualContextLogLine) {
            Add-StandardResult -Results $results -Severity 'WARNING' -Rule 'RedundantExecutionContextLog' -Path $relativePath -Message 'Manual execution context logging found in WriteLog; InitializeScriptEnvironment already logs execution context.'
        }
    }
}

$errorCount = @($results | Where-Object { $_.Severity -eq 'ERROR' }).Count
$warningCount = @($results | Where-Object { $_.Severity -eq 'WARNING' }).Count

if ($results.Count -gt 0) {
    $results | Sort-Object Severity, Path, Rule | Format-Table -AutoSize
    Write-Host ("SmartM365 script standards: {0} error(s), {1} warning(s), scope: {2}." -f $errorCount, $warningCount, $Scope)
}
else {
    Write-Host ("SmartM365 script standards: OK, scope: {0}." -f $Scope)
}

if ($errorCount -gt 0 -or ($WarningsAsErrors -and $warningCount -gt 0)) {
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC8Bqn7LHJL2uBE
# O9/4kbq+JeVtwPnm18qzKdAzT+3ELqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEILGfZp/S0FaXH9MdksR07lcWxtR0h8CeHXU7h1LfSrvMMA0GCSqG
# SIb3DQEBAQUABIIBgAM8/YLfOJrpPPNk2TbmQmDj1LudYW52dnr8vD1SjKDFZzlx
# lXPUYNiaR7IpchJ/egwoErgFK934r6jsA1LFQlG32f9TCbBDrUTLVaynItlBrN9v
# OFTAeQqVmruP/90rc3eBNcCuRuFEskvFeqRyarxPHCzCutCtb4pYMeQI2SNlccCF
# 90D3gHMOiGP68tVlwzyyAbSbvzxFICaIMjkWLFUmTzqe0HDY+63LRcEaSCBsRd8i
# OIkeDGJZWQY0RFoyXMkjD9nDWZJfkipDdnov5Y9J+dAQTy8ejZLW2u8Ov+xIf26N
# qEMp4YZBs01A6Auw/vVVOFkjnXpnT9UZa0RBKJd+UbIk6MtUNQGj/xft1v2idWje
# dFRyvqHjb/AZkjDg0/GcZXPxwQ06iYkqp8SS+ptmHEJBdtA4RZHQqIRY6ccHc7yW
# N4e3Js/6CLIc6Po9ZAMxZp3hSZI2vvsH5fR5bTvcsUYLQ1+1nowlzu6jeFnx+hyY
# rYZpv6s1q+PieXHyI6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMyMzA2
# MjRaMC8GCSqGSIb3DQEJBDEiBCCjAu8avDVK6MW8wpQXzzqR+MCkLJGF8EEVz/fe
# xKb36DANBgkqhkiG9w0BAQEFAASCAgB+7YezxDxq7u9TcpXKBNstv7oKfykWVgq4
# MXeRRKz/K4ZEBbAo6i5o+Xu3IMDCWVM89ePQQdJ771gqoH1Uyp64DUNeCiTWZ02t
# jt8zo+nvEc2z7iu6ssDuyMQFKxHBfGMlmXC5zKccEn3PtLFoQBvlwxfQRVBvdNW9
# qEy6vmFdBu9wZvCqxlbH9R3AxEOSjdeF+oTqZSMTnGlocUHlG5SO5EZSGJGM25CP
# rcBgBjRTlMk93TBJtKroWlOhx7oPwYnTiB+zP+abw8fhEeL+nY4VJz5+3DK61ytt
# 5Q8E0QG2T5fHtqyIl2NjkTiVcPEoqss74QibcdLA1i3Tv/IuHJqNj+IPe+mZDZNZ
# Vz3b/FqkO7xr94NJPCyDqATPcT5Sju1GrFKbksmLb2Uu2vNAsPuRwoLvzTwB5J6F
# IS2t0IBpfqgL2a68+lBV/xrpgPEBzs4v4tmW1OD9C0s54W5/5l1IHe4mXZAyqtpO
# 5FrtvBb7eHGM3ufbGUy+aGxoKOaaz+9c7ifBnl3kz/B4DwQZkaaOAuHB0sa+7Ql6
# Aq2nvPGQUP9cfmLQoss1e6EIta4cmMPi5IvEgegFhIXBpHoTzJJJaHgrDUPzFXHj
# j8k1GFZJ2dmT/Sz0oKK4x3RMa2SmDssKuCe8UmPBDHIl9jZAzbHDExy6GU58RVej
# WLeP+uYpZg==
# SIG # End signature block
