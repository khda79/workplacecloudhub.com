<#
    Name: SmartM365-Upgrade-Diagnostics-Detection.ps1
    Version: 1.0
    Description: Consolidated SetupDiag and Windows upgrade blocker detection.
#>

[CmdletBinding()]
param(
    [int]$LookbackDays = 7,
    [int]$MaxIssuesToDisplay = 5,
    [int]$MaxIssueLength = 180,
    [int]$MaxAnalysisAgeDays = 7
)

$ErrorActionPreference = "Stop"
$SetupDiagResultPath = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Output\SetupDiag\SetupDiagResults.xml"
$SetupDiagToolPaths = @(
    (Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Tools\SetupDiag\SetupDiag.exe"),
    (Join-Path $env:ProgramData "SetupDiag\SetupDiag.exe")
)
$EnterpriseRegistryPath = "HKLM:\SOFTWARE\SmartM365\IntuneRemediation\SetupDiag"

$BlockingCodes = @(
    "0xC1900208",
    "0xC1900101",
    "0x80070070",
    "0x80240020",
    "0x80242016",
    "0x800F0922",
    "0xC1900200",
    "0x8007042B",
    "0x8007001F"
)

$RollbackPaths = @(
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
    (Join-Path $env:WINDIR "Panther\NewOS\Rollback"),
    (Join-Path $env:SystemDrive 'Windows.old\$Windows.~BT\Sources\Rollback')
)

$SetupLogRoots = @(
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Rollback'),
    (Join-Path $env:WINDIR "Panther"),
    (Join-Path $env:SystemDrive '$Windows.~BT\Sources\Panther'),
    (Join-Path $env:SystemDrive 'Windows.old\$Windows.~BT\Sources\Rollback')
)

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Issues.Contains($Message)) {
        $Issues.Add($Message)
    }
}

function Add-EventIssue {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[string]]$Issues,
        [Parameter(Mandatory = $true)][string]$Source,
        [AllowNull()]$Events,
        [Parameter(Mandatory = $true)][string[]]$Codes,
        [Parameter(Mandatory = $true)][int]$MaxIssueLength
    )

    foreach ($eventRecord in @($Events)) {
        if ($null -eq $eventRecord -or [string]::IsNullOrWhiteSpace($eventRecord.Message)) {
            continue
        }

        foreach ($code in $Codes) {
            if ($eventRecord.Message -match [regex]::Escape($code)) {
                $message = $eventRecord.Message -replace "`r|`n", " "
                if ($message.Length -gt $MaxIssueLength) {
                    $message = $message.Substring(0, $MaxIssueLength) + "..."
                }
                Add-Issue -Issues $Issues -Message "$Source Id=$($eventRecord.Id) Code=$code Message=$message"
            }
        }
    }
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $cutoffDate = (Get-Date).AddDays(-1 * $LookbackDays)
    $hasSetupDiagResult = Test-Path -LiteralPath $SetupDiagResultPath -PathType Leaf
    $hasSetupDiagTool = $false
    foreach ($setupDiagToolPath in $SetupDiagToolPaths) {
        if (Test-Path -LiteralPath $setupDiagToolPath -PathType Leaf) {
            $hasSetupDiagTool = $true
            break
        }
    }
    $hasSetupLogs = $false
    $hasRollbackLogs = $false

    foreach ($setupLogRoot in $SetupLogRoots) {
        if (Test-Path -LiteralPath $setupLogRoot) {
            $hasSetupLogs = $true
            break
        }
    }

    foreach ($rollbackPath in $RollbackPaths) {
        if (Test-Path -LiteralPath $rollbackPath) {
            $hasRollbackLogs = $true
            Add-Issue -Issues $issues -Message "RollbackDetected=$rollbackPath"
        }
    }

    if ($hasSetupLogs -and -not $hasSetupDiagResult) {
        if ($hasSetupDiagTool) {
            Add-Issue -Issues $issues -Message "SetupLogsDetectedWithoutSetupDiagAnalysis"
        }
        else {
            Add-Issue -Issues $issues -Message "SetupLogsDetectedSetupDiagToolMissing"
        }
    }

    if ($hasRollbackLogs -and -not $hasSetupDiagResult) {
        Add-Issue -Issues $issues -Message "RollbackLogsDetectedWithoutSetupDiagAnalysis"
    }

    if ($hasSetupDiagResult) {
        $xmlAgeDays = (New-TimeSpan -Start (Get-Item -LiteralPath $SetupDiagResultPath).LastWriteTimeUtc -End (Get-Date).ToUniversalTime()).TotalDays
        if ($xmlAgeDays -gt $MaxAnalysisAgeDays) {
            Add-Issue -Issues $issues -Message ("SetupDiagResultOutdated AgeDays={0:N1}" -f $xmlAgeDays)
        }

        $xmlContent = Get-Content -LiteralPath $SetupDiagResultPath -Raw -ErrorAction SilentlyContinue
        foreach ($code in $BlockingCodes) {
            if ($xmlContent -match [regex]::Escape($code)) {
                Add-Issue -Issues $issues -Message "SetupDiagDetected=$code"
            }
        }

        if ($xmlContent -match "Matching Profile found:\s*(.+)") {
            Add-Issue -Issues $issues -Message "SetupDiagProfile=$($matches[1].Trim())"
        }
    }
    elseif (Test-Path -Path $EnterpriseRegistryPath) {
        $setupDiagState = Get-ItemProperty -Path $EnterpriseRegistryPath -Name LastRunUtc -ErrorAction SilentlyContinue
        if ($setupDiagState -and -not [string]::IsNullOrWhiteSpace($setupDiagState.LastRunUtc)) {
            $analysisAgeDays = (New-TimeSpan -Start ([datetime]::Parse($setupDiagState.LastRunUtc)) -End (Get-Date).ToUniversalTime()).TotalDays
            if ($analysisAgeDays -gt $MaxAnalysisAgeDays) {
                Add-Issue -Issues $issues -Message ("SetupDiagRegistryStateOutdated AgeDays={0:N1}" -f $analysisAgeDays)
            }
        }
    }

    $wuEvents = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 500 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $cutoffDate }
    Add-EventIssue -Issues $issues -Source "WUEvent" -Events $wuEvents -Codes $BlockingCodes -MaxIssueLength $MaxIssueLength

    $setupEvents = Get-WinEvent -LogName "Setup" -MaxEvents 300 -ErrorAction SilentlyContinue |
        Where-Object { $_.TimeCreated -ge $cutoffDate }
    Add-EventIssue -Issues $issues -Source "SetupEvent" -Events $setupEvents -Codes $BlockingCodes -MaxIssueLength $MaxIssueLength

    if ($issues.Count -gt 0) {
        Write-Output "Status=UpgradeDiagnosticsRequired"
        Write-Output "IssueCount=$($issues.Count)"
        $issues | Select-Object -First $MaxIssuesToDisplay | ForEach-Object { Write-Output $_ }
        exit 1
    }

    Write-Output "Status=Healthy"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKUu/if6iIZ8nq
# I/UAxiYk2/wExXsyT1opD/MGuw06eaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD5pJK6aAvjh9llLqUV
# LSlQ2kw+H4Y+MXjDnDS4MA4K/zANBgkqhkiG9w0BAQEFAASCAYAwyhs13Ebow4XG
# 1IWJPyFjsyfiBj/tB84AQD26NZyo4W+FG89mL9WI1jxQXWunTEDiBG80a/5wX6zW
# 6uuWwKLzxHN+5HcDuWROOEIoY7KDIsryN8+khauBmWPamEr1LbxrbXtby927Tq8T
# DODAG5QpsjHq+FNO8bfgFc9jv0uZu1abuwROLMCWFJjHadRvikiAN/zwsV/WadkV
# v8jpe+RP5/k9JpvAO8lGwJYr9YSScC+ACdHg97GoycG2chplDN3HmHDJuNlZnAUD
# ODg3hD5bTERR34Ouu8fbpjaagkgqJE4R8lBYpxtn1Iy5yReBUve3MyzgrYCKBVuL
# MBwrr7zbh3SkVXRVR9tpek14EPxcgsbQCq8fBjD90K/CMhZVTYIQnNSJ7aTasJ/v
# u7HZF/P1v1ck8n0VxJavesHiGH+TA1395U/ilrtaFo0/BMXmhmzhQyzI+8sGkOYb
# J7ByD90DLyJWJKZw3aChqZqira6XmD5DIQf8LGTqmEiFb6Fsd30=
# SIG # End signature block
