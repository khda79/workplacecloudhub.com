<#
.SYNOPSIS
Compiled CSV validation and serialization for large datasets. Dot-sourced Core helper.
.VERSION
1.0.0
.NOTES
PowerShell 7 only; the type is compiled on first use, for datasets of SmartM365CsvFastMinimumRows rows
or more. Each helper reproduces one PowerShell path exactly (Assert-SmartM365CsvDataCompleteness loops,
Select-Object -Property | Export-Csv with UTF8 or utf8BOM). Input outside that exact behaviour returns
$null or $false and the caller runs its original PowerShell path, so results never depend on the path.
#>
[CmdletBinding()]
param()
$script:SmartM365CsvFastMinimumRows = 5000
$script:SmartM365CsvFastState = $null

function Initialize-SmartM365CsvFast {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    if ($null -ne $script:SmartM365CsvFastState) { return $script:SmartM365CsvFastState }
    $script:SmartM365CsvFastState = $false
    if ($PSVersionTable.PSEdition -ne 'Core') { return $false }
    try {
        if (-not ('SmartM365.Core.CsvFast' -as [type])) {
            Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;

namespace SmartM365.Core
{
    public sealed class CsvRowScan
    {
        public bool Supported;
        public string MissingIdentityColumn;
        public bool DuplicateKey;
        public int MissingRows;
        public bool ThresholdReached;
        public string[] Examples = new string[0];
    }

    public static class CsvFast
    {
        // $row.PSObject.Properties[$name]: PowerShell returns $null for an empty name, C# throws.
        private static PSPropertyInfo GetProperty(PSObject item, string name)
        {
            return string.IsNullOrEmpty(name) ? null : item.Properties[name];
        }

        // Same text as a PowerShell [string] cast.
        private static string AsText(object value)
        {
            if (value == null) { return string.Empty; }
            string text = value as string;
            if (text != null) { return text; }
            object converted = LanguagePrimitives.ConvertTo(value, typeof(string), CultureInfo.InvariantCulture);
            return converted == null ? string.Empty : (string)converted;
        }

        // Unique-key and critical-field loops of Assert-SmartM365CsvDataCompleteness, in the same order.
        public static CsvRowScan Scan(object[] rows, string[] uniqueFields, string[] criticalFields, int failMinRows, double failPercent)
        {
            CsvRowScan result = new CsvRowScan();
            foreach (object row in rows) { if (row == null) { return result; } }
            result.Supported = true;
            if (uniqueFields.Length > 0)
            {
                HashSet<string> keys = new HashSet<string>(StringComparer.Ordinal);
                StringBuilder key = new StringBuilder();
                foreach (object row in rows)
                {
                    PSObject item = PSObject.AsPSObject(row);
                    key.Length = 0;
                    foreach (string field in uniqueFields)
                    {
                        PSPropertyInfo property = GetProperty(item, field);
                        if (property == null) { result.MissingIdentityColumn = field; return result; }
                        string value = AsText(property.Value).Trim().ToLowerInvariant();
                        // Length-prefixed parts: as injective as the JSON array key of the PowerShell path.
                        key.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value).Append('|');
                    }
                    if (!keys.Add(key.ToString())) { result.DuplicateKey = true; return result; }
                }
            }
            List<string> examples = new List<string>();
            int rowNumber = 1;
            foreach (object row in rows)
            {
                PSObject item = PSObject.AsPSObject(row);
                bool rowMissing = false;
                foreach (string field in criticalFields)
                {
                    PSPropertyInfo property = GetProperty(item, field);
                    object value = property == null ? null : property.Value;
                    if (value == null || string.IsNullOrWhiteSpace(AsText(value)))
                    {
                        rowMissing = true;
                        if (examples.Count < 10) { examples.Add(string.Format(CultureInfo.CurrentCulture, "row {0}, field {1}", rowNumber, field)); }
                    }
                }
                if (rowMissing)
                {
                    result.MissingRows++;
                    double percent = Math.Round(100.0 * result.MissingRows / rows.Length, 4);
                    bool rowThreshold = failMinRows > 0 && result.MissingRows >= failMinRows;
                    bool percentThreshold = failPercent > 0 && percent >= failPercent;
                    if (rowThreshold || percentThreshold) { result.ThresholdReached = true; break; }
                }
                rowNumber++;
            }
            result.Examples = examples.ToArray();
            return result;
        }

        private static void AppendQuoted(StringBuilder line, string value)
        {
            // Export-Csv (UseQuotes Always): a null value is an empty unquoted field.
            if (value == null) { return; }
            line.Append('"');
            if (value.IndexOf('"') >= 0) { line.Append(value.Replace("\"", "\"\"")); } else { line.Append(value); }
            line.Append('"');
        }

        // Select-Object -Property <columns> | Export-Csv -NoTypeInformation for PSObject rows holding note
        // properties. Returns false, with no file left, when a row is outside that exact behaviour.
        public static bool WriteQuoted(object[] rows, string[] columns, string path, char delimiter, bool byteOrderMark)
        {
            if (rows.Length == 0 || columns.Length == 0) { return false; }
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string column in columns)
            {
                if (string.IsNullOrEmpty(column) || column.IndexOfAny(new char[] { '*', '?', '[', ']' }) >= 0 || !seen.Add(column)) { return false; }
            }
            PSObject first = rows[0] as PSObject;
            if (first == null) { return false; }
            // Select-Object names each property after the matching property of the row; Export-Csv takes the
            // header from the first row.
            string[] header = new string[columns.Length];
            for (int index = 0; index < columns.Length; index++)
            {
                PSPropertyInfo property = first.Properties[columns[index]];
                header[index] = property == null ? columns[index] : property.Name;
            }
            bool written = false;
            try
            {
                using (StreamWriter writer = new StreamWriter(path, false, new UTF8Encoding(byteOrderMark), 1 << 20))
                {
                    StringBuilder line = new StringBuilder(1024);
                    for (int index = 0; index < header.Length; index++)
                    {
                        if (index > 0) { line.Append(delimiter); }
                        AppendQuoted(line, header[index]);
                    }
                    writer.WriteLine(line.ToString());
                    foreach (object row in rows)
                    {
                        PSObject item = row as PSObject;
                        if (item == null || !(item.BaseObject is PSCustomObject)) { return false; }
                        line.Length = 0;
                        for (int index = 0; index < columns.Length; index++)
                        {
                            if (index > 0) { line.Append(delimiter); }
                            PSPropertyInfo property = item.Properties[columns[index]];
                            if (property == null) { continue; }
                            if (!(property is PSNoteProperty)) { return false; }
                            object value = property.Value;
                            AppendQuoted(line, value == null ? null : value.ToString());
                        }
                        writer.WriteLine(line.ToString());
                    }
                }
                written = true;
                return true;
            }
            finally
            {
                if (!written && File.Exists(path)) { File.Delete(path); }
            }
        }
    }
}
'@
        }
        $script:SmartM365CsvFastState = $true
    }
    catch {
        WriteLog -Message ("Compiled CSV helpers unavailable; PowerShell CSV path used: {0}" -f $_.Exception.Message) -Level 'WARNING'
    }
    return $script:SmartM365CsvFastState
}

# Unique-key and critical-field scan; $null means the caller runs its PowerShell loops.
function Invoke-SmartM365CsvRowScan {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [AllowEmptyCollection()][string[]]$UniqueFields = @(),
        [AllowEmptyCollection()][string[]]$CriticalFields = @(),
        [int]$FailMinRows,
        [double]$FailPercent
    )
    if ($null -eq $Rows -or $Rows.Count -lt $script:SmartM365CsvFastMinimumRows -or -not (Initialize-SmartM365CsvFast)) { return $null }
    if ($null -eq $UniqueFields) { $UniqueFields = @() }
    if ($null -eq $CriticalFields) { $CriticalFields = @() }
    $scan = [SmartM365.Core.CsvFast]::Scan($Rows, $UniqueFields, $CriticalFields, $FailMinRows, $FailPercent)
    if (-not $scan.Supported) { return $null }
    return $scan
}

# Writes the rows as Export-Csv would; $false means nothing was written and the caller runs Export-Csv.
function Write-SmartM365CsvRowsFast {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)][string]$Path,
        [string]$Encoding = 'UTF8',
        [string]$Delimiter = ','
    )
    if ($null -eq $Rows -or $Rows.Count -lt $script:SmartM365CsvFastMinimumRows -or $Delimiter.Length -ne 1) { return $false }
    if ($Encoding -notin @('UTF8','utf8BOM') -or -not (Initialize-SmartM365CsvFast)) { return $false }
    return [SmartM365.Core.CsvFast]::WriteQuoted($Rows, $Columns, $Path, $Delimiter[0], ($Encoding -eq 'utf8BOM'))
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAqjlKeFlUtFLaq
# ebUirQ8ji7Q7t5GsjHivIelBmZ3AVKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIA9/YP8NK88kBb/XLAhM371p//PHpkXqxssY74DrzPJcMA0GCSqG
# SIb3DQEBAQUABIIBgAqnR1SUIbW4H3PgzZNOMxY9v/FonKL6oS7rvmAI9tfC9v/x
# xNp5o3AUY7EkJ7pw6hKAXo+gBt7hPUIhiHYddpez+8qD8DUmQTOTst3XSGgln7M3
# zLWmXiykbV1LlrZltt/yulO73EVUycMP8+JMV9qPcYnWUj1g6ffY+kz6veBP1/k+
# aucEVekjkANLfyOHOD7VwmFMX+ZdJ6glucIy5MVVKMstL0BlF1HWg/bzByrtATL3
# m6rFCV5V4OUnQuCNYqIlLueClTvEDE3K+9ODxpgUiYZEprC19sG5sQ7ti+uJWbPm
# D2tpm6h+lSZxXFSrBj6vXl0OM1vKIH1yH5jkvWxI/UnZKR1Y4A8f/cTzCh2G5sQ7
# xmMDDDhxR3e27zIxAkyF1gOxpXifc0lyQUuLe03xEYCNwsmkd0thageYfgb9kbbD
# IuZ1kyhFlqHuqJ3d/CMAoT/BAXqW3Lj2Rc3g+WFsT+1HF6C28M2V1xUZntr2pFH2
# ehex+PRkC6P9ffepraGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxMDEy
# NTNaMC8GCSqGSIb3DQEJBDEiBCCT/Y3nOlLpTksJXmlMhxw00WVlTnCVx/hLroTs
# 34RdKDANBgkqhkiG9w0BAQEFAASCAgAKsi4fimsPZycVdc4pYsAB2/1b7aDVIuBY
# AAubyN8mYMpHXBULQF3jI3hf56dZ9yPe4x4TGf+T0APym0SC921RJTEXKCESIgtr
# wnV7bnsUudDts7NIbxniwIia8WVdfAvoWWgSUqUdiTAYC43trd7aICwxlLOqCost
# 1jO703BwcMDzeDBj1wCaq52JzCZxLiozLr7mgduFFDu1OLQnMCS+/mEGdejwuAsF
# S7TwQmJwKjFy+bLFzOftmyAr3Osj4pXmV2mC7H73/26/plEB5SSNwiKlY5a4NUkt
# oqvPOHZTgjTsPpdXiVaXfjKyAgQcJibIKTnaBdT05+zAbIY6n/KvUXtQLkz/bmL4
# qwuAI1bu2s8ebCqdppVM7v9Z1/fCFlmlTYaXWC/cZ24JP3kDpjkPox4HmqZPG7cR
# N/eRL8eL1Ls5WTD9xe6Xqoz8+IEQWUgBlA/M0dD6U7pBHu8c/0/DtvES7r37Vjll
# SSx5NmfPr9DvmfuBfmByRLIJznJLFvkvMNzX4cm/BGnQBvfQ2Y0gz4+YehJ8eBYl
# xWrPiZvJ4pTF82KkeUzMZ9Y0frOvvkri4zFXb/2O5+JrZTnDcG2HMc3eYAjQxlvc
# 5aCp6YC/h/veQ610TTNyiWroJw3szI0/UzlS1/vJcTsgxqN+pso017wCRHZBKuNA
# ii0sLTBDTg==
# SIG # End signature block
