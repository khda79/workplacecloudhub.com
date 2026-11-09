# Bootstrap output reconciliation, version 0.5.0.
function Read-SiProposal {
    param([Parameter(Mandatory)][string]$Path)
    if (-not [IO.Path]::IsPathFullyQualified($Path)) { throw 'An absolute proposal path is required.' }
    $bytes=[IO.File]::ReadAllBytes($Path)
    $sha=[Security.Cryptography.SHA256]::Create()
    try { $hash=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','') } finally { $sha.Dispose() }
    $text=[Text.UTF8Encoding]::new($false,$true).GetString($bytes).TrimStart([char]0xfeff)
    $document=$text | ConvertFrom-Json -ErrorAction Stop
    if ($document -isnot [pscustomobject]) { throw 'Proposal must be a JSON object.' }
    [pscustomobject]@{Document=$document;Hash=$hash}
}
function New-SiBootstrapReconciliation {
    param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$Tenant,[Parameter(Mandatory)][string]$ProposalPath,$NewProfilePlan=$null)
    $legacy=Get-SiEditLocation $Root $Tenant 'identity'
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Target profile missing, invalid or conflicting.' }
    $read=Read-SmartM365JsonDocument -Path $legacy
    if ($read.Coexistence -ne 'Single') { throw 'Coexisting configuration files require review.' }
    $proposal=Read-SiProposal -Path $ProposalPath
    if ([string]$read.Document.ProfileKey -ine $Tenant -or [string]$proposal.Document.ProfileKey -ine $Tenant) { throw 'Profile keys do not match the selected tenant.' }
    $baseline=$null;$editedFields=@()
    if ($null -ne $NewProfilePlan) {
        if ($NewProfilePlan.ProfileSHA256 -cne 'ABSENT' -or $NewProfilePlan.ProfileKey -cne $Tenant -or $null -eq $NewProfilePlan.PendingProfile) {throw 'New-profile provenance is invalid.'}
        $templatePath=Join-Path $Root 'Config/Tenants/tenant.local.json.template'
        if ((Get-FileHash -LiteralPath $templatePath -Algorithm SHA256).Hash -cne $NewProfilePlan.PendingProfile.TemplateHash) {throw 'The new-profile template changed; review existing values instead.'}
        $resolved=New-SiBootstrapStagePlan -Root $Root -BootstrapPlan $NewProfilePlan
        if ($resolved.ProfileHash -cne $read.SHA256) {throw 'The new profile changed; review existing values instead.'}
        $baseline=$NewProfilePlan.PendingProfile.Model
        $editedFields=@($NewProfilePlan.PendingProfile.PendingEdits | ForEach-Object Field)
    }
    $groups=[ordered]@{
        Identity=@('AppId','TenantId','Thumb','Thumbprint')
        Mail=@('From','SmtpServer','MailSendAccessPolicyGroup')
        SharePoint=@('SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath')
    }
    $additions=[Collections.Generic.List[object]]::new();$replacements=[Collections.Generic.List[object]]::new(); $conflicts=[Collections.Generic.List[string]]::new(); $conflictValues=[Collections.Generic.List[object]]::new(); $blocked=[Collections.Generic.List[string]]::new()
    foreach ($group in $groups.Keys) {
        $pending=[Collections.Generic.List[object]]::new();$pendingReplacements=[Collections.Generic.List[object]]::new(); $groupConflict=$false
        $proposedCount=0
        foreach ($name in $groups[$group]) { if ($null -ne $proposal.Document.PSObject.Properties[$name]) { $proposedCount++ } }
        if ($proposedCount -gt 0 -and $proposedCount -ne $groups[$group].Count) { $groupConflict=$true; $conflicts.Add($group+' (incomplete proposal)') }
        foreach ($name in $groups[$group]) {
            $proposed=$proposal.Document.PSObject.Properties[$name]
            if ($null -eq $proposed) { continue }
            if ($proposed.Value -isnot [string]) { throw ('Proposal field must be text: '+$name) }
            if ([string]$proposed.Value -match '[\x00-\x1f\x7f\u200e\u200f\u202a-\u202e\u2066-\u2069]' -or ([string]$proposed.Value).Length -gt 2048) { throw ('Proposal field contains unsupported characters: '+$name) }
            $existing=$read.Document.PSObject.Properties[$name]
            if ($null -ne $existing) {
                if ($existing.Value -isnot [string] -or [string]$existing.Value -cne [string]$proposed.Value) {
                    $defaultProperty=if ($null -ne $baseline) {$baseline.PSObject.Properties[$name]} else {$null}
                    if ($null -ne $defaultProperty -and $name -notin $editedFields -and $existing.Value -is [string] -and $defaultProperty.Value -is [string] -and [string]$existing.Value -ceq [string]$defaultProperty.Value) {
                        $pendingReplacements.Add([pscustomobject]@{Field=$name;Current=[string]$existing.Value;Value=[string]$proposed.Value});continue
                    }
                    $groupConflict=$true; $conflicts.Add($name)
                    $currentJson=if ($existing.Value -is [string]) { ConvertTo-Json -InputObject $existing.Value -Depth 10 -Compress } else { '<non-text value>' }
                    if ($currentJson.Length -gt 4096 -or $currentJson -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') { throw ('Existing value cannot be safely displayed: '+$name) }
                    $proposedJson=ConvertTo-Json -InputObject $proposed.Value -Depth 10 -Compress
                    $conflictValues.Add([pscustomobject]@{Field=$name;CurrentJson=$currentJson;ProposedJson=$proposedJson})
                }
            } elseif (-not [string]::IsNullOrWhiteSpace([string]$proposed.Value) -and [string]$proposed.Value -notin @('__USE_GLOBAL__','USE_GLOBAL')) {
                $pending.Add([pscustomobject]@{Field=$name;Value=[string]$proposed.Value})
            }
        }
        if ($group -eq 'Identity') {
            $a=$proposal.Document.PSObject.Properties['Thumb']; $b=$proposal.Document.PSObject.Properties['Thumbprint']
            if ($null -ne $a -and $null -ne $b -and [string]$a.Value -ine [string]$b.Value) { $groupConflict=$true; $conflicts.Add('Thumb / Thumbprint') }
            foreach ($idName in @('AppId','TenantId')) {
                $id=$proposal.Document.PSObject.Properties[$idName]
                if ($null -ne $id -and -not [string]::IsNullOrWhiteSpace([string]$id.Value)) { $parsed=[guid]::Empty; if (-not [guid]::TryParse([string]$id.Value,[ref]$parsed) -or $parsed -eq [guid]::Empty) { $groupConflict=$true; $conflicts.Add($idName+' (invalid)') } }
            }
            foreach ($thumbName in @('Thumb','Thumbprint')) {
                $thumb=$proposal.Document.PSObject.Properties[$thumbName]
                if ($null -ne $thumb -and -not [string]::IsNullOrWhiteSpace([string]$thumb.Value) -and ([string]$thumb.Value -notmatch '^[0-9A-Fa-f]{40}$' -or [string]$thumb.Value -match '^0+$')) { $groupConflict=$true; $conflicts.Add($thumbName+' (invalid)') }
            }
        }
        if ($proposedCount -gt 0) {
            $required=if ($group -eq 'Mail') { @('From','MailSendAccessPolicyGroup') } else { $groups[$group] }
            foreach ($name in $required) { $p=$proposal.Document.PSObject.Properties[$name]; if ($null -eq $p -or [string]::IsNullOrWhiteSpace([string]$p.Value) -or [string]$p.Value -in @('__USE_GLOBAL__','USE_GLOBAL')) { $groupConflict=$true; $conflicts.Add($name+' (missing or inherited value)') } }
            if ($group -eq 'Mail') {
                $smtp=$proposal.Document.PSObject.Properties['SmtpServer']
                if ($null -ne $smtp -and [string]$smtp.Value -in @('__USE_GLOBAL__','USE_GLOBAL')) { $groupConflict=$true; $conflicts.Add('SmtpServer (inherited value)') }
                foreach ($name in @('From','MailSendAccessPolicyGroup')) {
                    $p=$proposal.Document.PSObject.Properties[$name]
                    if ($null -ne $p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
                        try { $email=[Net.Mail.MailAddress]::new([string]$p.Value); if ($email.Address -ine [string]$p.Value) { throw 'Invalid address.' } }
                        catch { $groupConflict=$true; $conflicts.Add($name+' (invalid address)') }
                    }
                }
            }
        }
        if ($groupConflict) { $blocked.Add($group) } else { foreach ($item in $pending) { $additions.Add($item) };foreach ($item in $pendingReplacements) {$replacements.Add($item)} }
    }
    $templatePath=Join-Path $Root 'Config/Tenants/tenant.local.json.template'
    $template=Read-SiProposal -Path $templatePath
    $templateKeys=@(); $templateAdditions=@()
    if ($additions.Count -gt 0 -or $replacements.Count -gt 0) {
        $shadow=$read.Document | ConvertTo-Json -Depth 100 | ConvertFrom-Json
        foreach ($item in $additions) { Add-Member -InputObject $shadow -NotePropertyName $item.Field -NotePropertyValue $item.Value }
        $templateKeys=@(Add-SmartM365MissingJsonTemplateProperties -Target $shadow -Template $template.Document)
        $templateAdditions=@(foreach ($key in $templateKeys) {
            $node=$shadow
            foreach ($part in ($key -split '\.')) {
                $property=$node.PSObject.Properties[$part]
                if ($null -eq $property) { throw 'A template addition could not be represented for review.' }
                $node=$property.Value
            }
            $valueJson=ConvertTo-Json -InputObject $node -Depth 100 -Compress
            if ($valueJson -match '[\u200e\u200f\u202a-\u202e\u2066-\u2069]') { throw 'A template addition contains unsupported direction controls.' }
            [pscustomobject]@{Field=$key;ValueJson=$valueJson}
        })
    }
    $blockedTemplateKeys=@(foreach ($group in $blocked) { foreach ($name in $groups[$group]) { if ($name -in $templateKeys) { $name } } })
    [pscustomobject]@{Root=$Root;Tenant=$Tenant;TargetPath=$read.Path;TargetHash=$read.SHA256;ProposalPath=$ProposalPath;ProposalHash=$proposal.Hash;TemplateHash=$template.Hash;NewProfilePlan=$NewProfilePlan;Additions=$additions.ToArray();Replacements=$replacements.ToArray();Conflicts=$conflicts.ToArray();ConflictValues=$conflictValues.ToArray();BlockedGroups=$blocked.ToArray();TemplateKeys=$templateKeys;TemplateAdditions=$templateAdditions;BlockedTemplateKeys=$blockedTemplateKeys}
}
function Invoke-SiBootstrapReconciliation {
    param([Parameter(Mandatory)]$Plan)
    $fresh=New-SiBootstrapReconciliation -Root $Plan.Root -Tenant $Plan.Tenant -ProposalPath $Plan.ProposalPath -NewProfilePlan $Plan.NewProfilePlan
    if ($fresh.TargetPath -cne $Plan.TargetPath -or $fresh.TargetHash -cne $Plan.TargetHash -or $fresh.ProposalHash -cne $Plan.ProposalHash -or $fresh.TemplateHash -cne $Plan.TemplateHash) { throw 'Profile, proposal or template changed; review the plan again.' }
    if (($fresh.Additions | ConvertTo-Json -Depth 10 -Compress) -cne ($Plan.Additions | ConvertTo-Json -Depth 10 -Compress)) { throw 'The reviewed additions changed.' }
    if (($fresh.Replacements | ConvertTo-Json -Depth 10 -Compress) -cne ($Plan.Replacements | ConvertTo-Json -Depth 10 -Compress)) {throw 'The reviewed replacements changed.'}
    if ($fresh.BlockedTemplateKeys.Count -gt 0) { throw 'Template enrichment would change a blocked group; resolve the conflicts first.' }
    if ($fresh.Additions.Count -eq 0 -and $fresh.Replacements.Count -eq 0) { return [pscustomobject]@{Added=@();TemplateKeys=@()} }
    $read=Read-SmartM365JsonDocument -Path $fresh.TargetPath
    if ($read.Coexistence -ne 'Single' -or $read.SHA256 -cne $fresh.TargetHash) { throw 'Target profile changed; review the plan again.' }
    foreach ($item in $fresh.Additions) {
        if ($null -ne $read.Document.PSObject.Properties[$item.Field]) { throw 'A proposed field is no longer absent.' }
        Add-Member -InputObject $read.Document -NotePropertyName $item.Field -NotePropertyValue $item.Value
    }
    foreach ($item in $fresh.Replacements) {if ($null -eq $read.Document.PSObject.Properties[$item.Field] -or [string]$read.Document.($item.Field) -cne $item.Current) {throw 'A default value changed.'};$read.Document.($item.Field)=$item.Value}
    $template=Read-SiProposal -Path (Join-Path $Plan.Root 'Config/Tenants/tenant.local.json.template')
    if ($template.Hash -cne $fresh.TemplateHash) { throw 'Template changed; review the plan again.' }
    $templateKeys=@(Add-SmartM365MissingJsonTemplateProperties -Target $read.Document -Template $template.Document)
    if (($templateKeys -join "`n") -cne ($fresh.TemplateKeys -join "`n")) { throw 'Template additions changed; review the plan again.' }
    $bytes=[Text.UTF8Encoding]::new($false).GetBytes(($read.Document | ConvertTo-Json -Depth 100)+[Environment]::NewLine)
    Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $bytes -ExpectedSHA256 $fresh.TargetHash -Validate {param($d) if ($d -isnot [pscustomobject]) {throw 'Expected JSON object.'}} | Out-Null
    $legacy=Get-SiEditLocation $Plan.Root $Plan.Tenant 'identity'
    if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') { throw 'Write completed but profile names now conflict; inspect both files before retrying.' }
    [pscustomobject]@{Added=@($fresh.Additions | ForEach-Object {$_.Field});TemplateKeys=$templateKeys}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBtJKGC+7ZTJfew
# TZ/aAf++EZBgoimXXI8+lZbrjZNGi6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINx1R19f5tHaQV/rZObki+rC++ox+kDvG3Vq2vTHzmL5MA0GCSqG
# SIb3DQEBAQUABIIBgDLXs4YmI5KxzRVrfqXCMr015CNM5yKsdE3k5uIY48h31xQF
# 6uUlNe+/6qevFSLTj2GHoRkn96TA/GEplUGn+w1c6C8++QPWQcp7DVPjqvu3sxO2
# vTasUTK+7wuW1s0lQvI/CoLFH8LLV2QTN2b2wGiLnIIKP+FikcWXDFWl2vRdxMsb
# Kra+0jQ1be7dGxfhnF6saLROQg38W09wQgOmgSsudHNCK9Fx1B4HuTbL0WQYhnzc
# JyExHSOVL+9tH44t2bs8vdPVgQPIlqVpZGZYRsIBSj5KfGB9PhKeoCcs6Oypxc7o
# 1rka37NXg7jEfkhUiDS8/hQ+zpOOnsxx1H8L4qcXDzmMIB+y1qla63cCgeNn0uqr
# OrvE6iEf+XzSxhDA8ACSWTmrL4gPKITYlQovZkOIJPzh24ZHKByfUly7iniVDdpo
# wgeaVQ3pEVpU/I3KlMs+LfOnWvSdVm/Xaz8T0BZizHGKlJ4NnuRirhiLFgfTSFaY
# ozPVaLd76fdR/smY2aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MTZaMC8GCSqGSIb3DQEJBDEiBCBSaInxLHLKBkDLrNmzChi9C2dUr4nVWiD6+GAo
# zhBW6TANBgkqhkiG9w0BAQEFAASCAgB4sLx4Fzk9hr7HvWpysyv1J7LaPKbWYsv2
# iQt/8SGG6plFTjr2Jwo8jb+vhxZsmpHAFwve9thgUZ7HHzPwEmWXo7AkVWe8eqSY
# rfQxocF718iw2Jvz6kKHZtI3DfSHCzvYif5HvwuBxKyo6Jkas7NgXuBOl0U09vvh
# GxA0pzZMDGTIZjuXuOLNFSAnqAz979iN2SDbjnDaAptgImMpJY+7Y9fLDlXK6Vp/
# gN5iRpcOk8FRxrMMb1IC6wmUvo5J6+/V1ztPIkqufyPqqv792o7bQm/38LN8cd0Y
# 8cYqlgDO3Bo3j08INtb1WQ7+e8fzAg3qnFh2ATpkJ87e0sI1hsEkNfiaAHnm/uxw
# hwSz89mdsYIAeE1r55nJwGLfZVSndIF7kwpGPIomur17Fq658A37wwmT+UmZWcBg
# RG3XGKBjNekVghzB8xrjNsVVELIib9tXimNitsLGt4c67peT7+Fv1YIvIx+ZeK54
# sOvGJU7CT1bkA9UDzZRLSMqSPwIQzaT7mPD1hNkRbrVeesgj/dfmuyMkVa1X9Z3v
# ZDg+2KZdS2U0xxxbx0HFKrTxHlhXu5Ov6FBoQIbd0PgR76rV3GYB2oG9jiJfc/VL
# qOSH/xRfDn+Usv0uYa2D/FL1kIQ/EMesPqf8YwChKQ9NmeSKemN1i2HjvWtoFqDV
# GUmJyISAOQ==
# SIG # End signature block
