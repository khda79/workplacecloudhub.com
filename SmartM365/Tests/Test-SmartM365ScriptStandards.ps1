<#
.SYNOPSIS
    Checks SmartM365 PowerShell script standards before commit/push.

.DESCRIPTION
    By default, this guard checks only changed SmartM365 scripts so legacy issues do not hide new regressions.
    Use -Scope All to scan all SmartM365 scripts, including SmartInventory, Setup, Communications,
    SharePointMigration, Devices, Intune Remediation, and future SmartM365 folders.

.VERSION
1.4
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

    $match = [regex]::Match($Content, '(?im)^[ \t]*\.VERSION[ \t]*\r?\n[ \t]*(?<Version>[^\r\n]+)')
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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJ2lg/jazcOVqi
# KhDmcitLrBh33QuNA1jUGVMFPUWS/aCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCBfNroUj4JTYPAOpP1fWqplDeEVa2+4PlFHFMKBalM+GTANBgkqhkiG9w0B
# AQEFAASCAYAHayH/DNA0KYAQfytyJUw6EkRTBN9V2PBTUZT7Ujrz0JdD8F7doIfu
# HzuHl4cfyZtZNl1eSfIffGTt6cnCgnqFTJpvlYKfZbF54yxi+nPZDGLj/Q3+ia4C
# YBQhCLtvtKvxQ79MdGqw5tZysyReEQVGFtRURYyx/BIyLFl7Vq5RbAp/lQVgVFG+
# XAWxzpld8nU0EovEYjmvEbqM2GiFtCRK/Oifx2OAIqweSAt2nY1++ucBvhL03/vV
# Tx4SAM7EckMYqzx/qXZbl8Jh+S8JuYTbO19HC3slJerJNexDVq3r+jiC6P30FWM1
# pcNAJoeKQS44hPO4bDvFkE8Rnl4rQeOAcLqrJtBf7dDKaUkz68rw8BSCAyH057Vw
# 3o90wPY6phYdr/+rMCwDw1/GQkt9PXURlkQffEiEwkj6Wb4hwaz57GXSkXNyyft/
# L8ojgB1e+dLupZ4+aP9i3VBksXC/AWNU4gli4p8smRpXHaXrrQAw4kGl7U40BqaJ
# 9UMrMntLIxc=
# SIG # End signature block
