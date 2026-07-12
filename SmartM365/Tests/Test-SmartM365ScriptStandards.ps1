<#
.SYNOPSIS
    Checks SmartM365 PowerShell script standards before commit/push.

.DESCRIPTION
    By default, this guard checks only changed SmartM365 scripts so legacy issues do not hide new regressions.
    Use -Scope All for a wider migration scan.

.VERSION
1.2
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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB4rA4LiXeDN5jD
# mr9KRUOdB3rKWI4mExFta2yRG4EQ9aCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCnnJPafEy/+89WJ6hT
# LZKR2jC/Zq05z5wZz1KGUq1BxTANBgkqhkiG9w0BAQEFAASCAYCs2ZOqT98OBY1k
# UtcrrhLmOieEX6V90DTU6+LHx6eYks9Sn6CqgJ1xEadGTRP8PRXUXN8YvfRslCPj
# qSO/DpSnCOlBoMQezVSgFb+fVUkMJ/DW06mZ+XQ5d2CBI2VZIbHw1H9Y8PEe6Tvj
# 4Gk6kk8pxZ3oa0kjMHYSRcUIqSmPUpiSAozpjXse2ruinS3gvoh5lLh4ylPf/xMm
# UBd7kuo+Lt/3zZz9v0Lz4onUnkXblqCtO1QlJBjLpdQWEP8YPFtTSs0yID9T7jl9
# 5a7VBrIYd1v3mrEWj9vAay7jIX97aaPysGYsRlLN2IdJnvqhxaOxHqAt7OvYVCm6
# fQWFC9fLkDRcipda2LjOEyf83cXtMbONsUIHTWlJLqp2hdd2Xa4L68HXvlci8Fjh
# MGfdqkcsfjKmcTh5fIgokpjdv2rihW8PXIJvJyPOpfVi1t2GSlRcgZH/6itH/UjU
# E4Pka5X+bwjmvolw0/98Y+ProrphofyLyknTmV5KyHX5Ac8rrqw=
# SIG # End signature block
