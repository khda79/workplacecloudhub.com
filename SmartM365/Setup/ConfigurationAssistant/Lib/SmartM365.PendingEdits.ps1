# Reviewed configuration edits, version 0.5.0. No writes occur during preparation.
function New-SiPendingEdit {
    param([Parameter(Mandatory)]$Snapshot,[AllowEmptyString()][string]$Value)
    $null=ConvertTo-SiEditValue -Snapshot $Snapshot -Value $Value
    $preview=Get-SiEditTemplatePreview -Snapshot $Snapshot
    $additions=@(foreach ($addition in $preview.Additions) {
        if ($addition.Field -match '(?i)password|secret|token|webhook|credential|private.?key') {
            if ($addition.ValueJson -notin @('null','""','"__USE_GLOBAL__"','"USE_GLOBAL"')) {throw 'Sensitive template additions require separate review.'}
            [pscustomobject]@{Field=$addition.Field;ValueJson='(empty or inherit)'}
        } else {
            $valueObject=ConvertFrom-Json -InputObject $addition.ValueJson -NoEnumerate
            if (Test-SiConfigSensitiveObject $valueObject) {throw 'Sensitive template additions require separate review.'}
            $addition
        }
    })
    [pscustomobject]@{Kind='Field';Snapshot=$Snapshot;Value=$Value;TemplateHash=$preview.Hash;AddedKeys=@($preview.Keys);Additions=$additions}
}
function New-SiPendingIdentityEdit {
    param([Parameter(Mandatory)]$Plan)
    $fresh=New-SiTenantIdentityEditPlan -Snapshot $Plan.Snapshot -OrganizationKey $Plan.OrganizationKey -EnvironmentKey $Plan.EnvironmentKey
    if ($fresh.TemplateHash -cne $Plan.TemplateHash -or (($fresh.AddedKeys -join "`n") -cne ($Plan.AddedKeys -join "`n"))) {throw 'Identity or template changed; review again.'}
    $snapshot=[pscustomobject]@{Root=$Plan.Snapshot.Root;Tenant=$Plan.Snapshot.Tenant;Step='tenant';Field='TenantIdentity';Path=$Plan.Snapshot.Path;Hash=$Plan.Snapshot.Hash;Present=$true;Value=$Plan.Snapshot.Current.TenantKey}
    [pscustomobject]@{Kind='Identity';Snapshot=$snapshot;Value=$Plan.TenantKey;Plan=$Plan;TemplateHash=$Plan.TemplateHash;AddedKeys=@($Plan.AddedKeys);Additions=@($Plan.Additions)}
}
function New-SiPendingReconciliationEdit {
    param([Parameter(Mandatory)]$Plan)
    $fresh=New-SiBootstrapReconciliation -Root $Plan.Root -Tenant $Plan.Tenant -ProposalPath $Plan.ProposalPath -NewProfilePlan $Plan.NewProfilePlan
    foreach ($key in @('TargetPath','TargetHash','ProposalHash','TemplateHash')) {if ($fresh.$key -cne $Plan.$key) {throw 'Profile, proposal or template changed; review again.'}}
    if (($fresh.Additions | ConvertTo-Json -Depth 10 -Compress) -cne ($Plan.Additions | ConvertTo-Json -Depth 10 -Compress) -or ($fresh.Replacements | ConvertTo-Json -Depth 10 -Compress) -cne ($Plan.Replacements | ConvertTo-Json -Depth 10 -Compress) -or $fresh.BlockedTemplateKeys.Count -gt 0 -or ($fresh.Additions.Count+$fresh.Replacements.Count) -eq 0) {throw 'The bootstrap changes cannot be applied safely.'}
    $read=Read-SmartM365JsonDocument -Path $fresh.TargetPath
    if ($read.Coexistence -ne 'Single' -or $read.SHA256 -cne $fresh.TargetHash) {throw 'Profile changed; review again.'}
    $template=Read-SiProposal -Path (Join-Path $Plan.Root 'Config/Tenants/tenant.local.json.template')
    $keys=@(Add-SmartM365MissingJsonTemplateProperties -Target $read.Document -Template $template.Document)
    $snapshot=[pscustomobject]@{Root=$Plan.Root;Tenant=$Plan.Tenant;Step='tenant';Field='BootstrapReconciliation';Path=$fresh.TargetPath;Hash=$fresh.TargetHash;Present=$true;Value='(missing fields only)'}
    $display=@($fresh.Additions | ForEach-Object {$_.Field+' = '+$_.Value})+@($fresh.Replacements | ForEach-Object {$_.Field+': '+$_.Current+' -> '+$_.Value})
    [pscustomobject]@{Kind='Reconciliation';Snapshot=$snapshot;Value=($display -join ', ');Plan=$Plan;TemplateHash=$fresh.TemplateHash;AddedKeys=$keys;Additions=@($fresh.TemplateAdditions)}
}
function New-SiPendingCertificateEdit {
    param([Parameter(Mandatory)]$Snapshot,[Parameter(Mandatory)][string]$Thumbprint)
    $value=$Thumbprint.Replace(' ','').ToUpperInvariant()
    if ($value -notmatch '^[0-9A-F]{40}$' -or $value -match '^0+$') {throw 'Enter a valid nonzero certificate thumbprint.'}
    $fresh=Get-SiCertificateEditSnapshot -Root $Snapshot.Root -Tenant $Snapshot.Tenant
    if ($fresh.Path -cne $Snapshot.Path -or $fresh.Hash -cne $Snapshot.Hash -or $fresh.TemplateHash -cne $Snapshot.TemplateHash -or (($fresh.AddedKeys -join "`n") -cne ($Snapshot.AddedKeys -join "`n"))) {throw 'Profile or template changed; review the certificate again.'}
    $reviewSnapshot=[pscustomobject]@{Root=$fresh.Root;Tenant=$fresh.Tenant;Step='identity';Field='CertificateThumbprints';Path=$fresh.Path;Hash=$fresh.Hash;Present=$true;Value=('Thumb='+$fresh.Thumb+'; Thumbprint='+$fresh.Thumbprint)}
    [pscustomobject]@{Kind='Certificate';Snapshot=$reviewSnapshot;Value=$value;TemplateHash=$fresh.TemplateHash;AddedKeys=@($fresh.AddedKeys);Additions=@($fresh.Additions)}
}
function Clear-SiPendingEdits {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue)
    $Queue.Clear()
}
function Add-SiPendingEdit {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue,[Parameter(Mandatory)]$Edit)
    for ($i=$Queue.Count-1;$i -ge 0;$i--) {
        $other=$Queue[$i]
        if ($other.Snapshot.Path -ieq $Edit.Snapshot.Path -and $other.Snapshot.Hash -cne $Edit.Snapshot.Hash) {throw 'Configuration changed; discard pending edits for this file and review again.'}
        if ($other.Snapshot.Path -ieq $Edit.Snapshot.Path -and $other.Snapshot.Field -ieq $Edit.Snapshot.Field) {$Queue.RemoveAt($i)}
    }
    $Queue.Add($Edit)
}
function Get-SiPendingEditReview {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue)
    $lines=[Collections.Generic.List[string]]::new()
    foreach ($edit in $Queue) {
        $before=if ($edit.Snapshot.Present) {$edit.Snapshot.Value} else {'(missing; template default shown previously)'}
        $after=$edit.Value
        if ($edit.Snapshot.Field -match '(?i)url') {$before='(masked)';$after=if ($after) {'(configured; masked)'} else {'(empty)'}}
        $lines.Add($edit.Snapshot.Field+': '+$before+' -> '+$after)
        foreach ($addition in $edit.Additions) {$lines.Add('  Template addition: '+$addition.Field+' = '+$addition.ValueJson)}
    }
    if (-not $lines.Count) {return 'No configuration changes are pending.'}
    return ($lines -join [Environment]::NewLine)
}
function Get-SiConcisePendingEditReview {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue)
    $lines=[Collections.Generic.List[string]]::new()
    foreach ($edit in $Queue) {
        if ($edit.Kind -eq 'Reconciliation') {
            foreach ($item in @($edit.Plan.Additions)+@($edit.Plan.Replacements)) {$lines.Add($item.Field+': '+[string]$item.Value)}
        } else {
            $value=if ($edit.Snapshot.Field -match '(?i)url|password|secret|token|credential|private.?key') {'(masked)'} else {[string]$edit.Value}
            $lines.Add($edit.Snapshot.Field+': '+$value)
        }
        if ($null -ne $edit.Additions -and @($edit.Additions).Count) {$lines.Add('  Additional missing settings will use template defaults.')}
    }
    if (-not $lines.Count) {return 'No configuration changes are pending.'}
    return ($lines -join [Environment]::NewLine)
}
function Invoke-SiPendingEdits {
    param([Parameter(Mandatory)][AllowEmptyCollection()][Collections.Generic.List[object]]$Queue,[scriptblock]$BeforeWrite)
    if (-not $Queue.Count) {return [pscustomobject]@{Applied=0;Files=0}}
    $groups=@($Queue | Group-Object {$_.Snapshot.Path})
    # Preflight every document and template before the first write.
    foreach ($group in $groups) {
        $first=$group.Group[0]
        foreach ($edit in $group.Group) {
            $snapshot=$edit.Snapshot
            if ($snapshot.Hash -cne $first.Snapshot.Hash -or $edit.TemplateHash -cne $first.TemplateHash -or (($edit.AddedKeys -join "`n") -cne ($first.AddedKeys -join "`n"))) {throw 'Pending changes no longer share the same reviewed configuration.'}
            if ($edit.Kind -eq 'Reconciliation') {
                $verified=New-SiPendingReconciliationEdit -Plan $edit.Plan
                if ($verified.Snapshot.Hash -cne $snapshot.Hash -or $verified.Value -cne $edit.Value -or (($verified.AddedKeys -join "`n") -cne ($edit.AddedKeys -join "`n"))) {throw 'Bootstrap additions changed; review again.'}
            } elseif ($edit.Kind -eq 'Identity') {
                $verified=New-SiPendingIdentityEdit -Plan $edit.Plan
                if ($verified.Snapshot.Hash -cne $snapshot.Hash -or $verified.Value -cne $edit.Value -or $verified.TemplateHash -cne $edit.TemplateHash) {throw 'Tenant identity changed; review again.'}
            } elseif ($edit.Kind -eq 'Certificate') {
                $fresh=Get-SiCertificateEditSnapshot -Root $snapshot.Root -Tenant $snapshot.Tenant
                $verified=New-SiPendingCertificateEdit -Snapshot $fresh -Thumbprint $edit.Value
                if ($verified.Snapshot.Path -cne $snapshot.Path -or $verified.Snapshot.Hash -cne $snapshot.Hash -or $verified.Snapshot.Value -cne $snapshot.Value -or $verified.TemplateHash -cne $edit.TemplateHash -or (($verified.AddedKeys -join "`n") -cne ($edit.AddedKeys -join "`n"))) {throw 'Certificate configuration or template changed; review again.'}
            } else {
                $fresh=Get-SiEditSnapshot -Root $snapshot.Root -Tenant $snapshot.Tenant -Step $snapshot.Step -Field $snapshot.Field
                if ($fresh.Path -cne $snapshot.Path -or $fresh.Hash -cne $snapshot.Hash -or $fresh.Present -ne $snapshot.Present -or $fresh.Value -cne $snapshot.Value) {throw 'Configuration changed; discard and review the pending changes.'}
                $preview=Get-SiEditTemplatePreview -Snapshot $fresh
                if ($preview.Hash -cne $edit.TemplateHash -or (($preview.Keys -join "`n") -cne ($edit.AddedKeys -join "`n"))) {throw 'Template changed; discard and review the pending changes.'}
                $null=ConvertTo-SiEditValue -Snapshot $snapshot -Value $edit.Value
            }
        }
    }
    foreach ($group in $groups) {
        $targets=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $blocked=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($reconciliation in @($group.Group | Where-Object Kind -eq 'Reconciliation')) {
            foreach ($name in $reconciliation.Plan.BlockedGroups) {
                $fields=switch ($name) {'Identity' {@('AppId','TenantId','Thumb','Thumbprint')} 'Mail' {@('From','SmtpServer','MailSendAccessPolicyGroup')} 'SharePoint' {@('SharePointSiteHostname','SharePointSitePath','SharePointLibraryDisplayName','SharePointTargetFolderPath')}}
                foreach ($field in $fields) {[void]$blocked.Add($field)}
            }
        }
        foreach ($edit in $group.Group) {
            $fields=switch ($edit.Kind) {'Reconciliation' {@($edit.Plan.Additions | ForEach-Object Field)+@($edit.Plan.Replacements | ForEach-Object Field)} 'Identity' {@('ProfileKey','OrganizationKey','EnvironmentKey','TenantKey')} 'Certificate' {@('Thumb','Thumbprint')} default {@($edit.Snapshot.Field)}}
            foreach ($field in $fields) {if ($blocked.Contains($field)) {throw ('Bootstrap reconciliation blocked changes to '+$field+'. Review the conflict first.')};if (-not $targets.Add($field)) {throw ('Pending changes conflict on field '+$field+'. Review the selected values.') }}
        }
    }
    $applied=0;$files=0
    foreach ($group in $groups) {
        try {
            $first=$group.Group[0];$snapshot=$first.Snapshot
            $legacy=Get-SiEditLocation $snapshot.Root $snapshot.Tenant $snapshot.Step
            if ((Read-SiPrivateJson ($legacy.Substring(0,$legacy.Length-5))).State -ne 'Present') {throw 'Configuration changed or conflicts.'}
            $read=Read-SmartM365JsonDocument -Path $legacy
            if ($read.Coexistence -ne 'Single' -or $read.Path -cne $snapshot.Path -or $read.SHA256 -cne $snapshot.Hash) {throw 'Configuration changed.'}
            $templatePath=if ($snapshot.Step -eq 'share') {Join-Path $snapshot.Root 'Config/SmartM365.global.local.json.template'} else {Join-Path $snapshot.Root 'Config/Tenants/tenant.local.json.template'}
            $bytes=[IO.File]::ReadAllBytes($templatePath);$sha=[Security.Cryptography.SHA256]::Create()
            try {$hash=([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','')} finally {$sha.Dispose()}
            if ($hash -cne $first.TemplateHash) {throw 'Template changed.'}
            $model=[Text.UTF8Encoding]::new($false,$true).GetString($bytes).TrimStart([char]0xfeff) | ConvertFrom-Json -ErrorAction Stop
            if ($model -isnot [pscustomobject]) {throw 'Invalid configuration template.'}
            $reconciliation=@($group.Group | Where-Object Kind -eq 'Reconciliation')
            $proposalFields=@()
            foreach ($item in $reconciliation) {foreach ($addition in $item.Plan.Additions) {if ($null -ne $read.Document.PSObject.Properties[$addition.Field]) {throw 'A proposed field is no longer absent.'};Add-Member -InputObject $read.Document -NotePropertyName $addition.Field -NotePropertyValue $addition.Value;$proposalFields+= $addition.Field};foreach ($replacement in $item.Plan.Replacements) {if ($null -eq $read.Document.PSObject.Properties[$replacement.Field] -or [string]$read.Document.($replacement.Field) -cne $replacement.Current) {throw 'A reviewed default changed.'};$read.Document.($replacement.Field)=$replacement.Value}}
            $added=@(Add-SmartM365MissingJsonTemplateProperties -Target $read.Document -Template $model)
            $expected=@($first.AddedKeys | Where-Object {$_ -notin $proposalFields})
            if (($added -join "`n") -cne ($expected -join "`n")) {throw 'Template additions changed.'}
            $secretPointers=[Collections.Generic.List[IntPtr]]::new()
            try {
            foreach ($edit in $group.Group) {
                if ($edit.Kind -eq 'Reconciliation') {continue}
                elseif ($edit.Kind -eq 'Identity') {
                    foreach ($key in @('ProfileKey','OrganizationKey','EnvironmentKey','TenantKey')) {if ($null -eq $read.Document.PSObject.Properties[$key]) {throw 'Identity field is missing.'}}
                    $read.Document.ProfileKey=$edit.Plan.Snapshot.Tenant;$read.Document.OrganizationKey=$edit.Plan.OrganizationKey;$read.Document.EnvironmentKey=$edit.Plan.EnvironmentKey;$read.Document.TenantKey=$edit.Plan.TenantKey
                } elseif ($edit.Kind -eq 'Certificate') {
                    if ($null -eq $read.Document.PSObject.Properties['Thumb'] -or $null -eq $read.Document.PSObject.Properties['Thumbprint']) {throw 'Certificate fields are missing.'}
                    $read.Document.Thumb=$edit.Value;$read.Document.Thumbprint=$edit.Value
                } else {
                    if ($null -eq $read.Document.PSObject.Properties[$edit.Snapshot.Field]) {throw 'A selected setting is missing.'}
                    $read.Document.($edit.Snapshot.Field)=ConvertTo-SiEditValue -Snapshot $edit.Snapshot -Value $edit.Value
                }
            }
            $output=[Text.UTF8Encoding]::new($false).GetBytes(($read.Document | ConvertTo-Json -Depth 100)+[Environment]::NewLine)
            if ($null -ne $BeforeWrite) {& $BeforeWrite $read.Path}
            Write-SmartM365JsonBytesAtomically -Path $read.Path -Bytes $output -ExpectedSHA256 $snapshot.Hash -Validate {param($d) if ($d -isnot [pscustomobject]) {throw 'Expected JSON object.'}} | Out-Null
            } finally {foreach ($ptr in $secretPointers) {[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)};if ($null -ne $output) {[Array]::Clear($output,0,$output.Length)};$read=$null}
            foreach ($edit in @($group.Group)) {[void]$Queue.Remove($edit);$applied++}
            $files++
        } catch {throw ('Apply stopped after '+$files+' file(s) and '+$applied+' change(s): '+$_.Exception.Message+'. Review the remaining changes before retrying.')}
    }
    [pscustomobject]@{Applied=$applied;Files=$files}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAYZt8HAbj+RKLV
# /aGvXP6xynfxht6LesJ6gGJMPtSltaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFit/gUpCncWMLUTBcMeTRC+Y8rYp07fbq6NelsIiKKWMA0GCSqG
# SIb3DQEBAQUABIIBgJJVe+/MtZ4Q854zGEKy8ZyFGSzhBssci2NZ2eHTiCsXuMmT
# ImaJjRlrsI4ZwxJSRiojnxRcE2j28qCqxtSahjEUMOaR1rYVCv9+u5Q3IDDLVWVg
# 8TzJ63UQdXyOKeGsxtvFtRMyxkAhqtTOcn9dOVbKDhmHOAipAeR9HPSF3p2DZTM9
# rzjwjV835uOAsTmQYT9hxroPWzYvrRKd5OQbffHlonZ0R2a14irNvJ2WmciU0PJK
# j8uMB4R36YAngKPfrwSAydBzh0j8kPOOldvOK+vjvcyLdgGmwTMbt0evQz19tQiq
# ACJJeNNWLXsvfR6FifBknnhYW3JPxrAMl2nbxq6QuQS2/cNTuew8lEbTYoXXR55h
# gH7FAg7agYZYo/6V+CM1ut0kt5wn6D1sYecYNuXiUmmKM8+QyZjai6Fft90allGd
# h1hhfb3W7TEnAQ+WCA08X+C4AhFLQn4CmrK6ql6myk516fp5wEPTYyD5mbZj2ZLQ
# jr/FrrWJv7x685KPuqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MjdaMC8GCSqGSIb3DQEJBDEiBCD6rOd6riLQplZMjI9xIuOQ8EyqVdBPiS0n0ymF
# Bk4KmTANBgkqhkiG9w0BAQEFAASCAgAr0/n3p1SuSVJCxVBmytH1aNnGzo3PEDxB
# NOI01bA6ZX4kxE0RlPnfHsuARzjzoQ9qmsv/PQ7+iMnGz8rjNVZ982Zl58WdKixv
# 1DkoaY1yC8Zph/lKRSeSw1nzUuiiZomwLCf6FVir9raTDGJZSy8VkGdTmjx71Pyr
# rLxQnMjKOuSYOGCkTt76wzf2/lmZNjMnI4X5QeslwBdbBlKMkrIXuq3GKe35EFzL
# amwGxnnmRL8B8K7HP9dDbfo4xq+7gWmv0LEaqXeO43FpXu35G1m8fpPjl/SycBi/
# b1N+5oi9Q2+a59FMtRsD6jRQMONdgUdVyB8CIGffHANmT9OYhymk7fpcudvqAJnf
# WF/slCl8Q17tGMDYkQ+3XrwJXns6FLGeMq05RR/a6i95KQ6yS9rHPcNQj5sJ30Vm
# L+mMqhVw/x5glBO/s9IZl2aonEho8vejrHEDEPHEyT8om3eqzJ5jxP/psHFEj00P
# /qz4n5et929HnOZYSEmnae416Fgt2u87vmUcUr4KJ5zSdZDBb3V7k8zlWsxlUEnV
# SPKsz+NS36CifphATSTBKSoeMY/IPa3PliY7pk5+gM025eJRzyKrKPonPDNNiOOH
# yPZzHC8vwLHPXjFlVMLXZ0t1hzQ6DEERivvuWYGHGLvo7CtxSQ6QVDh+MmfIer0I
# 5uSKOjlfMA==
# SIG # End signature block
