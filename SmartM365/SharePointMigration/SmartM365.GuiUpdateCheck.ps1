<#
.SYNOPSIS
Reusable SmartM365 GUI update checker.

.VERSION
1.0.0
#>

Set-StrictMode -Version 2.0

function Start-SmartM365GuiUpdateCheck {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseOutputTypeCorrectly','')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Owner,
        [Parameter(Mandatory = $true)][string]$ManifestPath,
        [Parameter(Mandatory = $true)][string]$AppRoot,
        [scriptblock]$OnStatus
    )

    $null = $Owner
    $null = $OnStatus

    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        return $null
    }

    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        Add-Type -AssemblyName WindowsBase -ErrorAction Stop
    }
    catch {
        return $null
    }

    $jobScript = {
        param(
            [string]$JobManifestPath,
            [string]$JobAppRoot
        )

        Set-StrictMode -Version 2.0
        $ErrorActionPreference = 'Stop'

        function Test-SmartM365UpdateCheckDisabled {
            param([hashtable]$Manifest)

            $names = New-Object System.Collections.Generic.List[string]
            [void]$names.Add('SMARTM365_GUI_UPDATE_CHECK')
            if ($Manifest.ContainsKey('DisableEnvironmentVariable') -and -not [string]::IsNullOrWhiteSpace([string]$Manifest['DisableEnvironmentVariable'])) {
                [void]$names.Add([string]$Manifest['DisableEnvironmentVariable'])
            }

            foreach ($name in @($names)) {
                $value = [Environment]::GetEnvironmentVariable($name, 'Process')
                if ([string]::IsNullOrWhiteSpace($value)) {
                    $value = [Environment]::GetEnvironmentVariable($name, 'User')
                }
                if ([string]::IsNullOrWhiteSpace($value)) {
                    $value = [Environment]::GetEnvironmentVariable($name, 'Machine')
                }

                if ($value -and ($value.Trim().ToLowerInvariant() -in @('0','false','no','off','disabled'))) {
                    return $true
                }
            }

            return $false
        }

        function ConvertTo-SmartM365RawGitHubUrl {
            param(
                [string]$RepositoryUrl,
                [string]$Branch,
                [string]$RemotePath
            )

            $repo = $RepositoryUrl.TrimEnd('/')
            if ($repo -notmatch '^https://github\.com/(?<Owner>[^/]+)/(?<Repo>[^/]+)$') {
                throw "Unsupported repository URL: $RepositoryUrl"
            }

            $owner = $matches.Owner
            $name = $matches.Repo
            if ($name.EndsWith('.git')) {
                $name = $name.Substring(0, $name.Length - 4)
            }

            $path = $RemotePath.Replace('\','/').TrimStart('/')
            return ('https://raw.githubusercontent.com/{0}/{1}/{2}/{3}' -f $owner,$name,$Branch,$path)
        }

        function Get-SmartM365GuiVersionFromContent {
            param(
                [string]$Content,
                [hashtable]$Component
            )

            $source = 'Header'
            if ($Component.ContainsKey('VersionSource')) {
                $source = [string]$Component['VersionSource']
            }

            if ($source -eq 'Header') {
                $match = [regex]::Match($Content, '(?m)^\s*\.VERSION\s*\r?\n\s*(?<Version>[^\r\n]+)')
                if ($match.Success) {
                    return $match.Groups['Version'].Value.Trim()
                }
            }

            if ($source -eq 'Variable') {
                $variable = [string]$Component['VersionVariable']
                if ([string]::IsNullOrWhiteSpace($variable)) {
                    throw "VersionVariable is required for $($Component['Name'])."
                }

                $escaped = [regex]::Escape($variable)
                $match = [regex]::Match($Content, ('(?m)^\s*\${0}\s*=\s*[''"](?<Version>[^''"]+)[''"]' -f $escaped))
                if ($match.Success) {
                    return $match.Groups['Version'].Value.Trim()
                }
            }

            return ''
        }

        function Compare-SmartM365GuiVersion {
            param(
                [string]$Left,
                [string]$Right
            )

            $leftParts = @([regex]::Matches($Left, '\d+') | ForEach-Object { [int]$_.Value })
            $rightParts = @([regex]::Matches($Right, '\d+') | ForEach-Object { [int]$_.Value })
            $count = [Math]::Max($leftParts.Count, $rightParts.Count)
            for ($i = 0; $i -lt $count; $i++) {
                $leftValue = if ($i -lt $leftParts.Count) { $leftParts[$i] } else { 0 }
                $rightValue = if ($i -lt $rightParts.Count) { $rightParts[$i] } else { 0 }
                if ($leftValue -lt $rightValue) { return -1 }
                if ($leftValue -gt $rightValue) { return 1 }
            }

            return 0
        }

        function Get-SmartM365GuiUpdateCachePath {
            param([hashtable]$Manifest)

            $cacheRoot = ''
            if ($Manifest.ContainsKey('CacheRoot')) {
                $cacheRoot = [string]$Manifest['CacheRoot']
            }

            if ([string]::IsNullOrWhiteSpace($cacheRoot)) {
                $programData = [Environment]::GetFolderPath('CommonApplicationData')
                $cacheRoot = Join-Path $programData 'SmartM365\GuiUpdateCheck'
            }

            try {
                if (-not (Test-Path -LiteralPath $cacheRoot -PathType Container)) {
                    New-Item -Path $cacheRoot -ItemType Directory -Force | Out-Null
                }
            }
            catch {
                $cacheRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'SmartM365\GuiUpdateCheck'
                if (-not (Test-Path -LiteralPath $cacheRoot -PathType Container)) {
                    New-Item -Path $cacheRoot -ItemType Directory -Force | Out-Null
                }
            }

            $appId = 'SmartM365'
            if ($Manifest.ContainsKey('AppId') -and -not [string]::IsNullOrWhiteSpace([string]$Manifest['AppId'])) {
                $appId = [string]$Manifest['AppId']
            }

            $safeAppId = [regex]::Replace($appId, '[^A-Za-z0-9._-]+', '-').Trim('-._')
            if ([string]::IsNullOrWhiteSpace($safeAppId)) {
                $safeAppId = 'SmartM365'
            }

            return Join-Path $cacheRoot ("{0}.json" -f $safeAppId)
        }

        function Get-SmartM365GuiUpdateAlertKey {
            param(
                [string]$AppId,
                [object[]]$Updates
            )

            $text = $AppId + '|' + (@($Updates | Sort-Object Name | ForEach-Object {
                '{0}:{1}->{2}' -f $_.Name,$_.LocalVersion,$_.RemoteVersion
            }) -join '|')

            $sha = [System.Security.Cryptography.SHA256]::Create()
            try {
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
                return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
            }
            finally {
                $sha.Dispose()
            }
        }

        function Test-SmartM365GuiUpdateShouldAlert {
            param(
                [hashtable]$Manifest,
                [string]$AlertKey
            )

            $cachePath = Get-SmartM365GuiUpdateCachePath -Manifest $Manifest
            $lastAlertKey = ''
            if (Test-Path -LiteralPath $cachePath -PathType Leaf) {
                try {
                    $cache = Get-Content -LiteralPath $cachePath -Raw | ConvertFrom-Json
                    if ($cache.LastAlertKey) {
                        $lastAlertKey = [string]$cache.LastAlertKey
                    }
                }
                catch { [void]$_.Exception }
            }

            $payload = [pscustomobject]@{
                LastAlertKey = $AlertKey
                LastCheckedUtc = (Get-Date).ToUniversalTime().ToString('o')
            }
            try {
                $payload | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $cachePath -Encoding UTF8
            }
            catch { [void]$_.Exception }

            return ($lastAlertKey -ne $AlertKey)
        }

        try {
            if (-not (Test-Path -LiteralPath $JobManifestPath -PathType Leaf)) {
                return [pscustomobject]@{ HasUpdate = $false; ShouldAlert = $false; Message = ''; GitHubUrl = ''; StatusMessage = ''; ErrorMessage = '' }
            }

            $manifest = Import-PowerShellDataFile -LiteralPath $JobManifestPath
            if (Test-SmartM365UpdateCheckDisabled -Manifest $manifest) {
                return [pscustomobject]@{ HasUpdate = $false; ShouldAlert = $false; Message = ''; GitHubUrl = ''; StatusMessage = 'GitHub update check disabled.'; ErrorMessage = '' }
            }

            $repositoryUrl = [string]$manifest['RepositoryUrl']
            $branch = [string]$manifest['Branch']
            if ([string]::IsNullOrWhiteSpace($branch)) {
                $branch = 'main'
            }

            $timeoutSeconds = 4
            if ($manifest.ContainsKey('TimeoutSeconds')) {
                $timeoutSeconds = [int]$manifest['TimeoutSeconds']
            }

            $updates = New-Object System.Collections.Generic.List[object]
            foreach ($component in @($manifest['Components'])) {
                $localPath = Join-Path $JobAppRoot ([string]$component['LocalPath'])
                if (-not (Test-Path -LiteralPath $localPath -PathType Leaf)) {
                    continue
                }

                $localContent = Get-Content -LiteralPath $localPath -Raw
                $localVersion = Get-SmartM365GuiVersionFromContent -Content $localContent -Component $component
                if ([string]::IsNullOrWhiteSpace($localVersion)) {
                    continue
                }

                $rawUrl = ConvertTo-SmartM365RawGitHubUrl -RepositoryUrl $repositoryUrl -Branch $branch -RemotePath ([string]$component['RemotePath'])
                $remoteContent = (Invoke-WebRequest -Uri $rawUrl -UseBasicParsing -TimeoutSec $timeoutSeconds).Content
                $remoteVersion = Get-SmartM365GuiVersionFromContent -Content $remoteContent -Component $component
                if ([string]::IsNullOrWhiteSpace($remoteVersion)) {
                    continue
                }

                if ((Compare-SmartM365GuiVersion -Left $localVersion -Right $remoteVersion) -lt 0) {
                    [void]$updates.Add([pscustomobject]@{
                        Name = [string]$component['Name']
                        LocalVersion = $localVersion
                        RemoteVersion = $remoteVersion
                    })
                }
            }

            if ($updates.Count -eq 0) {
                return [pscustomobject]@{ HasUpdate = $false; ShouldAlert = $false; Message = ''; GitHubUrl = ''; StatusMessage = ''; ErrorMessage = '' }
            }

            $productName = [string]$manifest['ProductName']
            if ([string]::IsNullOrWhiteSpace($productName)) {
                $productName = 'SmartM365 GUI'
            }

            $lines = New-Object System.Collections.Generic.List[string]
            [void]$lines.Add(("A newer version of {0} is available on GitHub." -f $productName))
            [void]$lines.Add('')
            foreach ($update in @($updates)) {
                [void]$lines.Add(("- {0}: local {1}, GitHub {2}" -f $update.Name,$update.LocalVersion,$update.RemoteVersion))
            }
            [void]$lines.Add('')
            [void]$lines.Add('Open GitHub now?')

            $appId = if ($manifest.ContainsKey('AppId')) { [string]$manifest['AppId'] } else { 'SmartM365' }
            $alertKey = Get-SmartM365GuiUpdateAlertKey -AppId $appId -Updates @($updates)
            $shouldAlert = Test-SmartM365GuiUpdateShouldAlert -Manifest $manifest -AlertKey $alertKey
            $githubUrl = if ($manifest.ContainsKey('GitHubPathUrl') -and -not [string]::IsNullOrWhiteSpace([string]$manifest['GitHubPathUrl'])) {
                [string]$manifest['GitHubPathUrl']
            } else {
                $repositoryUrl
            }

            return [pscustomobject]@{
                HasUpdate = $true
                ShouldAlert = $shouldAlert
                Message = ($lines -join [Environment]::NewLine)
                GitHubUrl = $githubUrl
                StatusMessage = ("GitHub update available for {0}." -f $productName)
                ErrorMessage = ''
            }
        }
        catch {
            return [pscustomobject]@{
                HasUpdate = $false
                ShouldAlert = $false
                Message = ''
                GitHubUrl = ''
                StatusMessage = ''
                ErrorMessage = $_.Exception.Message
            }
        }
    }

    try {
        $job = Start-Job -ScriptBlock $jobScript -ArgumentList $ManifestPath,$AppRoot
    }
    catch {
        return $null
    }

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromSeconds(2)
    $timer.Add_Tick({
        try {
            if ($job.State -eq 'Running' -or $job.State -eq 'NotStarted') {
                return
            }

            $timer.Stop()
            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue | Select-Object -Last 1
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

            if (-not $result) {
                return
            }

            if ($result.HasUpdate -and $result.StatusMessage -and $OnStatus) {
                & $OnStatus $result.StatusMessage 'Update available'
            }

            if ($result.HasUpdate -and $result.ShouldAlert) {
                $answer = [System.Windows.MessageBox]::Show(
                    $Owner,
                    [string]$result.Message,
                    'Update available on GitHub',
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Information
                )
                if ($answer -eq [System.Windows.MessageBoxResult]::Yes -and $result.GitHubUrl) {
                    $psi = [System.Diagnostics.ProcessStartInfo]::new([string]$result.GitHubUrl)
                    $psi.UseShellExecute = $true
                    [System.Diagnostics.Process]::Start($psi) | Out-Null
                }
            }
        }
        catch {
            try {
                $timer.Stop()
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }
            catch { [void]$_.Exception }
        }
    }.GetNewClosure())
    $timer.Start()
    return $timer
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC/kfYAgCF2ETWv
# zZz/O+D+Qr0Zr0ERAga03nTm4WbAHKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAsa+dDmr4Bd8/R4x+3
# n5GZWcrgXMEHIbLnVaBTInI3BjANBgkqhkiG9w0BAQEFAASCAYB749Cu+7NaNTFO
# Z5EqK/96PzH7++GOeunFXpKoVGlFkzMHb5kIkkQ2fK2t+bar46ChJMOTqZ54CKxX
# H0yfFpAl5W9aWmv5Ot1HNqXoUe4shstjuiENTCottTALANrNrS0ucHeA3q5VaFYa
# bVVT9wxwh2qo9m3sBKvKNT8aYMQq7YqVEfs0xgNe3OCzrS8CoeYggRc1JVrKkpim
# SmF3i9Z8o1ia/ZDjW8F6KjnZhB1nEliq1a8pcHHTLW/yyMzHQYb8HPf7FfrcC936
# dPtZf0tz+NAWFQLWo89wX0wizECNws7aypKaqF3rB3/MLiHI/oEeb9nIfWLsOr0D
# nheL0XhMu9L0H85xkF+iHixZmxCnG+iIHsh2+BBPVRQwOYfCgpven47cGzdopO+f
# D/q7LwhtGn1WUKMWPC8DVc7TNM3bln4bvje6wR8seKp5RMkHH/5/jnrIr06ZKvQh
# 064CRP73o5i/bmU5k6284ShRWVXbFouU/V+sVUZDTlM8l80Tdlc=
# SIG # End signature block
