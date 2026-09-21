<#
.SYNOPSIS
Synthetic regression tests for the bounded Teams Graph collector audit.
.VERSION
1.0.2
#>
[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$ResultPath
)

$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$results = New-Object 'System.Collections.Generic.List[object]'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-TeamsGraph-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)

function Assert-Offline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Test-OfflineCase {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        $results.Add([pscustomobject]@{ Name = $Name; Passed = $true; Error = '' })
    }
    catch {
        $results.Add([pscustomobject]@{ Name = $Name; Passed = $false; Error = $_.Exception.Message })
    }
}

function Import-OfflineFunctions {
    param([string]$Path, [string[]]$Names)
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "Source parse failed: $Path" }
    $definitions = foreach ($name in $Names) {
        $node = $ast.Find({
            param($candidate)
            $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq $name
        }, $true)
        if ($null -eq $node) { throw "Function not found: $name" }
        $node.Extent.Text
    }
    # Only named definitions run. Collector initialization, authentication,
    # Graph, mail, upload and scheduled-task entry points are never executed.
    New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
}

function Get-LiteralArrayAssignment {
    param([string]$Path, [string]$VariableName)
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw "Source parse failed: $Path" }
    $assignment = $ast.Find({
        param($candidate)
        $candidate -is [Management.Automation.Language.AssignmentStatementAst] -and
        $candidate.Left -is [Management.Automation.Language.VariableExpressionAst] -and
        $candidate.Left.VariablePath.UserPath -eq $VariableName
    }, $true)
    if ($null -eq $assignment) { throw "Assignment not found: $VariableName" }
    return @(& ([scriptblock]::Create($assignment.Right.Extent.Text)))
}

$teamsPath = Join-Path $SourceRoot 'SmartInventory/M365Inventory/Teams/SmartM365-Teams-Inventory.ps1'
$teamsPhonePath = Join-Path $SourceRoot 'SmartInventory/M365Inventory/Teams/SmartM365-TeamsPhonePstnUsage-Inventory.ps1'
$dashboardSchemaPath = Join-Path (Split-Path $SourceRoot -Parent) 'SmartWorkplaceDashboard/source-schema.json'
$dashboardSelectionPath = Join-Path (Split-Path $SourceRoot -Parent) 'SmartWorkplaceDashboard/source-selection.json'
$finOpsContractPath = Join-Path (Split-Path $SourceRoot -Parent) 'SmartFinOps/Config/SmartFinOps-Workplace-SourceContracts.json'

$teamsModule = Import-OfflineFunctions -Path $teamsPath -Names @(
    'IsoUtc',
    'Get-TeamsRetryDelay',
    'Invoke-Graph',
    'Test-TeamsGraphProperty',
    'Get-TeamsGraphPropertyValue',
    'Get-GraphCollection',
    'Invoke-TeamsGraphBatch',
    'Get-ReportRow',
    'Write-TeamsCsvAtomically',
    'Export-InventoryCsv'
)
function Reset-TeamsOfflineModule {
    if ($script:teamsModule) { Remove-Module $script:teamsModule -Force -ErrorAction SilentlyContinue }
    $script:teamsModule = Import-OfflineFunctions -Path $teamsPath -Names @(
        'IsoUtc',
        'Get-TeamsRetryDelay',
        'Invoke-Graph',
        'Test-TeamsGraphProperty',
        'Get-TeamsGraphPropertyValue',
        'Get-GraphCollection',
        'Invoke-TeamsGraphBatch',
        'Get-ReportRow',
        'Write-TeamsCsvAtomically',
        'Export-InventoryCsv'
    )
}
$teamsPhoneModule = Import-OfflineFunctions -Path $teamsPhonePath -Names @(
    'ConvertTo-TeamsPhoneUtcText',
    'Get-TeamsPhoneRetryAfterSeconds',
    'Invoke-TeamsPhoneGraphRequest',
    'Get-SmartM365GraphCollection',
    'New-TeamsPhoneDateWindow'
)

try {
    Test-OfflineCase 'Teams collection follows every nextLink page' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            $script:PageRequestCount = 0
            function script:Invoke-Graph {
                param($Uri, $Operation, $Headers)
                $script:PageRequestCount++
                if ($script:PageRequestCount -eq 1) {
                    return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'team-1' }); '@odata.nextLink' = 'synthetic-page-2' }
                }
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'team-2' }, [pscustomobject]@{ id = 'team-3' }) }
            }
        }
        $items = @(& $teamsModule { Get-GraphCollection -Uri 'synthetic-page-1' -Operation 'fixture' })
        $requestCount = & $teamsModule { $script:PageRequestCount }
        Assert-Offline ($items.Count -eq 3 -and $requestCount -eq 2) 'Teams pagination did not return all three synthetic rows in two pages.'
    }

    Test-OfflineCase 'Teams collection accepts Hashtable Graph pages' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            $script:PageRequestCount = 0
            function script:Invoke-Graph {
                param($Uri, $Operation, $Headers)
                $script:PageRequestCount++
                if ($script:PageRequestCount -eq 1) {
                    return @{ value = @([pscustomobject]@{ id = 'team-hash-1' }); '@odata.nextLink' = 'synthetic-hash-page-2' }
                }
                return @{ value = @([pscustomobject]@{ id = 'team-hash-2' }) }
            }
        }
        $items = @(& $teamsModule { Get-GraphCollection -Uri 'synthetic-hash-page-1' -Operation 'fixture' })
        Assert-Offline ($items.Count -eq 2) 'Teams rejected or truncated Hashtable Graph pages.'
    }

    Test-OfflineCase 'Teams malformed Graph page is rejected' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            function script:Invoke-Graph { [pscustomobject]@{ unexpected = 'missing value property' } }
        }
        $caught = $false
        try { & $teamsModule { Get-GraphCollection -Uri 'synthetic-malformed' -Operation 'fixture' } | Out-Null }
        catch { $caught = $_.Exception.Message -match '(?i)value|malformed|invalid' }
        Assert-Offline $caught 'A Graph collection response without value was accepted as a complete empty page.'
    }

    Test-OfflineCase 'Teams later page failure exposes no partial result' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            $script:PageRequestCount = 0
            function script:Invoke-Graph {
                $script:PageRequestCount++
                if ($script:PageRequestCount -eq 1) {
                    return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'partial-team' }); '@odata.nextLink' = 'synthetic-page-2' }
                }
                throw 'Synthetic Teams page 2 failure'
            }
        }
        $items = $null
        $caught = $false
        try { $items = & $teamsModule { Get-GraphCollection -Uri 'synthetic-page-1' -Operation 'fixture' } }
        catch { $caught = $_.Exception.Message -match 'Synthetic Teams page 2 failure' }
        Assert-Offline ($caught -and $null -eq $items) 'Teams returned the first page after a later page failed.'
    }

    Test-OfflineCase 'Teams repeated nextLink is rejected without looping' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            $script:CycleRequestCount = 0
            function script:Invoke-Graph {
                $script:CycleRequestCount++
                if ($script:CycleRequestCount -gt 3) { throw 'Synthetic loop guard reached.' }
                [pscustomobject]@{ value = @([pscustomobject]@{ id = "team-$script:CycleRequestCount" }); '@odata.nextLink' = 'synthetic-repeat' }
            }
        }
        $caught = $false
        try { & $teamsModule { Get-GraphCollection -Uri 'synthetic-repeat' -Operation 'fixture' } | Out-Null }
        catch { $caught = $_.Exception.Message -match '(?i)repeat|cycle' }
        $requestCount = & $teamsModule { $script:CycleRequestCount }
        Assert-Offline ($caught -and $requestCount -le 2) 'A repeated nextLink was not rejected deterministically.'
    }

    Test-OfflineCase 'Teams retry honors Retry-After' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            $script:GraphAttempts = 0
            $script:SleepSeconds = New-Object 'System.Collections.Generic.List[int]'
            function script:WriteLog { param($Message, $Level) }
            function script:Start-Sleep { param([int]$Seconds) [void]$script:SleepSeconds.Add($Seconds) }
            function script:Invoke-MgGraphRequest {
                $script:GraphAttempts++
                if ($script:GraphAttempts -eq 1) {
                    $exception = [InvalidOperationException]::new('Synthetic throttled response')
                    $exception.Data['Retry-After'] = '17'
                    throw $exception
                }
                [pscustomobject]@{ value = @() }
            }
        }
        & $teamsModule { Invoke-Graph -Uri 'synthetic-retry' -Operation 'fixture' } | Out-Null
        $state = & $teamsModule { [pscustomobject]@{ Attempts = $script:GraphAttempts; Sleeps = @($script:SleepSeconds) } }
        Assert-Offline ($state.Attempts -eq 2 -and $state.Sleeps.Count -eq 1 -and $state.Sleeps[0] -eq 17) 'Retry-After=17 was not honored before the successful retry.'
        $httpHeaders = [pscustomobject]@{}
        $httpHeaders | Add-Member -MemberType ScriptMethod -Name GetValues -Value { param($Name) if ($Name -eq 'Retry-After') { @('19') } }
        $httpDelay = & $teamsModule { param($Headers) Get-TeamsRetryDelay -Headers $Headers -DefaultSeconds 10 } $httpHeaders
        Assert-Offline ($httpDelay -eq 19) 'Retry-After from HttpResponseHeaders.GetValues was not honored.'
    }

    Test-OfflineCase 'Teams non-transient Graph error is not retried' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            $script:GraphAttempts = 0
            $script:SleepSeconds = New-Object 'System.Collections.Generic.List[int]'
            function script:WriteLog { param($Message, $Level) }
            function script:Start-Sleep { param([int]$Seconds) [void]$script:SleepSeconds.Add($Seconds) }
            function script:Invoke-MgGraphRequest { $script:GraphAttempts++; throw [UnauthorizedAccessException]::new('Synthetic forbidden') }
        }
        $caught = $false
        try { & $teamsModule { Invoke-Graph -Uri 'synthetic-forbidden' -Operation 'fixture' } | Out-Null }
        catch { $caught = $_.Exception.Message -match 'Synthetic forbidden' }
        $state = & $teamsModule { [pscustomobject]@{ Attempts = $script:GraphAttempts; Sleeps = @($script:SleepSeconds) } }
        Assert-Offline ($caught -and $state.Attempts -eq 1 -and $state.Sleeps.Count -eq 0) 'A non-transient Graph error was retried or swallowed.'
    }

    Test-OfflineCase 'Teams batch transport retry honors Retry-After' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            $script:BatchAttempts = 0
            $script:SleepSeconds = New-Object 'System.Collections.Generic.List[int]'
            function script:WriteLog { param($Message, $Level) }
            function script:Start-Sleep { param([int]$Seconds) [void]$script:SleepSeconds.Add($Seconds) }
            function script:Invoke-MgGraphRequest {
                $script:BatchAttempts++
                if ($script:BatchAttempts -eq 1) {
                    $exception = [InvalidOperationException]::new('Synthetic throttled batch transport')
                    $exception.Data['Retry-After'] = '23'
                    throw $exception
                }
                [pscustomobject]@{ responses = @([pscustomobject]@{ id = '1'; status = 200; body = [pscustomobject]@{} }) }
            }
        }
        $responses = @(& $teamsModule { Invoke-TeamsGraphBatch -Requests @(@{ id = '1'; method = 'GET'; url = '/synthetic' }) })
        $state = & $teamsModule { [pscustomobject]@{ Attempts = $script:BatchAttempts; Sleeps = @($script:SleepSeconds) } }
        Assert-Offline ($responses.Count -eq 1 -and $state.Attempts -eq 2 -and $state.Sleeps[0] -eq 23) 'Batch transport Retry-After=23 was not honored.'
    }

    Test-OfflineCase 'Teams batch subrequest retry honors Retry-After' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            $script:BatchAttempts = 0
            $script:SleepSeconds = New-Object 'System.Collections.Generic.List[int]'
            function script:WriteLog { param($Message, $Level) }
            function script:Start-Sleep { param([int]$Seconds) [void]$script:SleepSeconds.Add($Seconds) }
            function script:Invoke-MgGraphRequest {
                $script:BatchAttempts++
                if ($script:BatchAttempts -eq 1) {
                    return [pscustomobject]@{ responses = @([pscustomobject]@{ id = '1'; status = 429; headers = @{ 'Retry-After' = '11' } }) }
                }
                [pscustomobject]@{ responses = @([pscustomobject]@{ id = '1'; status = 200; body = [pscustomobject]@{} }) }
            }
        }
        $responses = @(& $teamsModule { Invoke-TeamsGraphBatch -Requests @(@{ id = '1'; method = 'GET'; url = '/synthetic' }) })
        $state = & $teamsModule { [pscustomobject]@{ Attempts = $script:BatchAttempts; Sleeps = @($script:SleepSeconds) } }
        Assert-Offline ($responses.Count -eq 1 -and $state.Attempts -eq 2 -and $state.Sleeps[0] -eq 11) 'Batch subrequest Retry-After=11 was not honored.'
    }

    Test-OfflineCase 'Teams activity report failure blocks collection' {
        Reset-TeamsOfflineModule
        & $teamsModule {
            function script:WriteLog { param($Message, $Level) }
            function script:Invoke-Graph { throw [InvalidOperationException]::new('Synthetic report failure') }
        }
        $caught = $false
        try { & $teamsModule { Get-ReportRow -ReportName 'getTeamsTeamActivityDetail' -Period D180 } | Out-Null }
        catch { $caught = $_.Exception.Message -match 'Synthetic report failure|could not be loaded' }
        Assert-Offline $caught 'The mandatory Teams activity report failure was converted to an empty successful result.'
    }

    Test-OfflineCase 'Teams DATA-LAST survives failed promotion' {
        Reset-TeamsOfflineModule
        $timestamped = Join-Path $testRoot 'M365_Teams_Teams_20260911_010203.csv'
        $latest = Join-Path $testRoot 'M365_Teams_Teams.csv'
        [IO.File]::WriteAllText($latest, 'LAST VALID SYNTHETIC EXPORT')
        $before = (Get-FileHash -LiteralPath $latest).Hash
        & $teamsModule {
            param($latestPath)
            $script:GeneratedCsvPaths = New-Object 'System.Collections.Generic.List[string]'
            $script:DryRun = $true
            $script:AppendHistory = $false
            function script:Assert-SmartM365CsvDataCompleteness { param($Data, $Columns, $TimestampedPath, $LatestPath) }
            function script:Add-SmartM365TenantKey { process { $_ } }
            function script:Move-Item {
                param($LiteralPath, $Destination, [switch]$Force, $ErrorAction)
                if ($Destination -eq $latestPath) { throw 'Synthetic promotion failure before replacement.' }
                [IO.File]::Move($LiteralPath, $Destination)
            }
        } $latest
        $caught = $false
        try {
            & $teamsModule {
                param($timestampedPath, $latestPath)
                Export-InventoryCsv -Rows @() -Columns @('RunId', 'TeamId') -TimestampedPath $timestampedPath -LatestPath $latestPath -HistoryPath ''
            } $timestamped $latest
        }
        catch { $caught = $_.Exception.Message -match 'Synthetic .*failure' }
        Assert-Offline $caught 'Synthetic DATA-LAST promotion failure did not terminate publication.'
        Assert-Offline ((Get-FileHash -LiteralPath $latest).Hash -eq $before) 'Failed publication changed the last valid Teams DATA-LAST file.'
        $bytes = [IO.File]::ReadAllBytes($timestamped)
        Assert-Offline ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) 'Atomic Teams CSV publication no longer preserves the UTF-8 BOM contract.'
    }

    Test-OfflineCase 'Teams UTC timestamps remain normalized' {
        Reset-TeamsOfflineModule
        $value = & $teamsModule { IsoUtc ([datetimeoffset]'2026-09-11T03:15:00+02:00') }
        Assert-Offline ($value -eq '2026-09-11T01:15:00Z') 'Teams date normalization no longer produces invariant UTC text.'
    }

    Test-OfflineCase 'Teams alert mail subject retains its operational payload' {
        $source = Get-Content -LiteralPath $teamsPath -Raw
        Assert-Offline ($source -match '\$subject="\[\$\(\$worst\.ToUpperInvariant\(\)\)\] Microsoft Teams Inventory - \$RunDateUtc"') 'Teams alert subject payload changed.'
        Assert-Offline ($source -notmatch '\$subject="(?:SMART ?365|\[SmartM365\])') 'Teams alert subject still hard-codes a legacy product prefix.'
    }

    Test-OfflineCase 'Teams CSV schemas match Dashboard contracts' {
        $schema = Get-Content -LiteralPath $dashboardSchemaPath -Raw | ConvertFrom-Json
        $identity = @('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')
        $pairs = @(
            @{ Variable = 'teamColumns'; File = 'M365_Teams_Teams.csv' },
            @{ Variable = 'memberColumns'; File = 'M365_Teams_Members.csv' },
            @{ Variable = 'channelColumns'; File = 'M365_Teams_Channels.csv' },
            @{ Variable = 'guestColumns'; File = 'M365_Teams_Guests.csv' }
        )
        foreach ($pair in $pairs) {
            $actual = $identity + @(Get-LiteralArrayAssignment -Path $teamsPath -VariableName $pair.Variable)
            $expected = @($schema.files.($pair.File).columns)
            Assert-Offline (($actual -join '|') -ceq ($expected -join '|')) "Schema mismatch for $($pair.File)."
        }
    }

    Test-OfflineCase 'Teams required detail and channel failures are not swallowed' {
        $source = Get-Content -LiteralPath $teamsPath -Raw
        Assert-Offline ($source -notmatch 'catch\{WriteLog -Message \(\"Team details unavailable') 'A team-details failure is still logged and converted to a default value.'
        Assert-Offline ($source -notmatch 'catch\{WriteLog -Message \(\"Channels unavailable') 'A channel failure is still logged and converted to an empty collection.'
    }

    Test-OfflineCase 'Teams Phone pagination returns all pages' {
        $requestCount = 0
        $result = & $teamsPhoneModule {
            param([ref]$counter)
            Get-SmartM365GraphCollection -InitialUri 'phone-page-1' -Operation 'fixture' -RequestInvoker {
                param($uri)
                $counter.Value++
                if ($uri -eq 'phone-page-1') {
                    return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'call-1' }); '@odata.nextLink' = 'phone-page-2' }
                }
                return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'call-2' }) }
            }
        } ([ref]$requestCount)
        Assert-Offline ($result.PageCount -eq 2 -and @($result.Items).Count -eq 2 -and $requestCount -eq 2) 'Teams Phone did not retain every paged call row.'
    }

    Test-OfflineCase 'Teams Phone persistent page failure exposes no partial result' {
        $result = $null
        $caught = $false
        try {
            $result = & $teamsPhoneModule {
                Get-SmartM365GraphCollection -InitialUri 'phone-page-1' -Operation 'fixture' -RequestInvoker {
                    param($uri)
                    if ($uri -eq 'phone-page-1') {
                        return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'partial-call' }); '@odata.nextLink' = 'phone-page-2' }
                    }
                    throw 'Synthetic page 2 failure'
                }
            }
        }
        catch { $caught = $_.Exception.Message -match 'Synthetic page 2 failure' }
        Assert-Offline ($caught -and $null -eq $result) 'Teams Phone returned a partial collection after a later page failed.'
    }

    Test-OfflineCase 'Teams Phone splits ranges into bounded windows' {
        $windows = @(& $teamsPhoneModule {
            New-TeamsPhoneDateWindow -FromUtc ([datetime]'2026-01-01T00:00:00Z') -ToUtc ([datetime]'2026-09-11T00:00:00Z')
        })
        Assert-Offline ($windows.Count -eq 3) 'Expected three Teams Phone windows for the synthetic date range.'
        foreach ($window in $windows) {
            Assert-Offline (([datetimeoffset]$window.ToUtc - [datetimeoffset]$window.FromUtc).TotalDays -le 90) 'A Teams Phone request window exceeds 90 days.'
        }
    }

    Test-OfflineCase 'Teams consumers keep selected source contracts' {
        $selection = Get-Content -LiteralPath $dashboardSelectionPath -Raw | ConvertFrom-Json
        foreach ($file in @('M365_Teams_Teams.csv', 'M365_Teams_Members.csv', 'M365_Teams_Channels.csv', 'M365_Teams_Guests.csv', 'M365_Teams_PhoneUserUsage.csv')) {
            Assert-Offline ($selection.includedFiles -contains $file) "Dashboard no longer selects $file."
        }
        $contract = Get-Content -LiteralPath $finOpsContractPath -Raw | ConvertFrom-Json
        $phone = @($contract.Sources | Where-Object Key -eq 'M365TeamsPhoneUserUsage')
        Assert-Offline ($phone.Count -eq 1 -and $phone[0].FileNames -contains 'M365_Teams_PhoneUserUsage.csv') 'SmartFinOps Teams Phone source contract changed or disappeared.'
    }
}
finally {
    Remove-Module $teamsModule, $teamsPhoneModule -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$failed = @($results | Where-Object { -not $_.Passed })
$summary = [pscustomobject]@{
    PowerShell = $PSVersionTable.PSVersion.ToString()
    Total = $results.Count
    Passed = $results.Count - $failed.Count
    Failed = $failed.Count
    Results = $results.ToArray()
}
if ($ResultPath) {
    $parent = Split-Path -Path $ResultPath -Parent
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }
    $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ResultPath -Encoding utf8
}
$summary | ConvertTo-Json -Depth 6
if ($failed.Count -gt 0) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDNeUKpO0bw8IBw
# t4dn1ywzeRepcSQYHL0lk9uRJ748hqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICPDLSZpoM5XDROBosBxVtl2ypzZgYEnQzXitTMM7Uz2MA0GCSqG
# SIb3DQEBAQUABIIBgCqmnanrM6omAMqGRwHiPIEdA35ewMA6cZ4Sx2PsieCI1Ub5
# qoUyOHO6XOnFIoZtVZmDDQ6Y3UL1KSfptkyR9MHgj7hubUOcvyqaT/TnxYcvp9Xp
# +xAbgTLXwBnrfjOpVIwwhzRJgpbNsCOvNh+9He/1JgLuVLiUt5LIMDg18rg1K0hH
# R4xDJjyy6AgUVzvxBaqrnUDbPBkKlDzVmp4P6iBQhWLCypTde5Pj1N18oYRnUmyB
# 69BYhtanaohaOct635OCViOBQ0Z4JgSDtBmoekkYcIQ13DNF0jowKmGTVY+INnc3
# D0ukDFn8dVahqRDj1eBid2gB9yY0Fbk524YVIJHzt0FarZLM52IMx6bVLQxnr2vH
# A5ncIziGFs9m8RAa3OL+zjmlxN3cMOm7Mt8hkXZHZL/O0fQgwOzKgkokSM9LVDRU
# 62K6Cl5ohBg4HTQuSmAnp8u6Q60YcJ+hqooT2xMq8YYnG/twUppyU7uX1hQi5PFU
# r7mvxjItXzHtpRdYD6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjExNDE5
# NTdaMC8GCSqGSIb3DQEJBDEiBCA0UrXPD7ZhuctJlQC5TmRP+vLicZJzaJdPj1Ge
# DOEvVjANBgkqhkiG9w0BAQEFAASCAgBrggxqbHR5ujMIvVS7+T8CWeJpytNaw3y5
# 1E6pYlUg+zLpHIgvSj0/5PUfjhJRR8IdnRrM0vWkV7pihWLyG8pAzHzPBVx00s8G
# ZiqWsFsSBYODR+PRIcnyMrwkQhZOP/y6P2GC+tp4sQ1wPvthJNT2+ENjvO+rKTxQ
# aEhnWi60hYVdPEt41NhLdz1tPBoEtlRPASfkbSAEsmIYstc7mfor2blnWkmHgKnu
# cEZ1sLsGpaTnguSmbj4QP6z9jn8X8oJUADDUx6ZbkvfEp2yHSD+gdJqgWQwoIGhP
# iLwfr5XeRwAN4wqNCRWND2Wl0ik2lDe614Hc0Uwa8On9u0RgvWh1aUpA6HDMGB0h
# DZ+vg5/XHKImwJi3JcPXguYkGauZxDiuuOWYMz81kk2M+ozMBcyCtiGXE9IVb+Wn
# n5kfqFdJBohfNMgmKeFL1b325xjoPpXl2kHSNTzZhgB2iTcOC23q/hGaNf2f2kVJ
# sF2RsM/hhHP2MJONq74mvfUh0bJqx8ykUfTObRHwJWkKNYP08sM6V91X7b05lhWi
# 1qFeknRtypCU/hlRTFIK5HRb8elwWyeVZB6B4BmUBzgA2/kfPQbfFkCQCgUy0lQ4
# sD9FtR8ngsknGsiJDWYHHPE9qTFQ0R6gm2WarcijA6+emhRzN7r50eZvlr2RycNU
# ITyXvIq5Ig==
# SIG # End signature block
