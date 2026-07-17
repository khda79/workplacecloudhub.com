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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfNuPuXTrdL61Y
# D0QWkxXKUuvfgSJwDUUdfEZSItuUXqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDujmeV4b90arTfwfwRnz5IM5l8x/EN1SbBxobzH9iMrMA0GCSqG
# SIb3DQEBAQUABIIBgFYPBFKCERdiP+d50New9KxJfDon7gZhxLOf2a2hoebjOdaX
# uimSb6ac4Db36QaJTFaMMuZBNbAD4O/5VcF84IuITRKDFiFNvQhP+R5NcszxoYVt
# Sjd57wjbalw1kKrXaPfqdmhjk+AjGsEpAC8zR8wLmr2Ei1/wWcc2HUlZ9naKI8h3
# eIcu6ke0IdvQEH5ZTRgoHWA14klxmfYY4TM+Ua3gX73ljvdO/0dtBmqFwQk464Dg
# DU4UDM3ziq7bOcH2zrXPOuqPc+5JwQPnoX+y7JnjqGj+b+t62F736JUtuK0OKuw5
# XXLby8kaLQWIoTzvjZqqGHqE8FXchgPc+S6vuPYl6qj7N0BxwVW7mzyKyIP2UFhh
# qu4VfsJ++K2YkC+0fyD1HSkGmonHDL/FkjZmvTe16j6MM/K961KPROpOOdBJ9Nkz
# j3ZCKYWMVZu3J0wpPGrT4HoVd5Fdrt5pN9lEU/VhJjwaXhMnFQSf6pL/oQo/+kO5
# EMpisLEXZl2j070LgKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MjBaMC8GCSqGSIb3DQEJBDEiBCAw7wgJEq5jmJjG8vI8ctZiLaYATBB6//VI6BC5
# ELNh5DANBgkqhkiG9w0BAQEFAASCAgDNKfkiQTLBj5GqOVlUZBPU66P7MogIF5Ff
# y7GTofJ5/t9wtk/T9nNksYzXAz6A4Ph1NYQVTP3nOU7ti+9BAyObnQ8pUrbVbVgH
# WK+t5f8wE5byUSKJ3cgHKAkPaYhir/W3uKhfKgYuwqMIGmHMFmT2RfPxhlXQR558
# vpbGioXZfmD3m3RljZfbZOeb5u8p2cNDSoG14C5+IYewaWghLk61te1Hve/kFYJC
# wgJfb36zW/R01sIm0ApaemQeaYyUTes+By5IW7dMt0GvUS/NqrE4U+EySJ7P3fr+
# 92W35GpBNX5GxJ2K+wfd14G0uZd35I0th/BNxQ8PGNCIDfEmSVjxQwy0CrRT7W5N
# 3pqy4emr/bM+mrw3zmsmzxaBkCTz5oXvfOx7UizTS4f6QOx5YOwqDR/bt70hQ0pU
# PQrZMx3cfSwCV1B8qyt7B6aMTjHWYCp/CZd3T+UHFpFCe5tXOICh7oB3Yu+sfEHf
# UJPyYbrgCYu0wES0g7SIgbDNUcEfXH24hndUxPgWnbGjIJB1xhvKNQFwzqu+kO0Y
# K5CKOGJjmzhRu8huQqlMK+BtWTZrENRK89q54U5jm4PzAb/iwXxq+BlP/ccET0XN
# uWmpXE2p3ce28MrtRBPfgqWLYqlEF4exFPwzzoce0eArGXg8Gd/2SKnUrWNUvOQg
# BnpzhN/xNQ==
# SIG # End signature block
