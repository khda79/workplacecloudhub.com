// Streaming CSV checks with the same invariant .NET conversion rules as the
// PowerShell validator. Values and identities are never included in diagnostics.
using System;
using System.IO;
using System.Text;
using System.Globalization;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Cryptography;
using Microsoft.VisualBasic.FileIO;

namespace Workplace.Prepared {
    public sealed class Issue {
        public string Column { get; set; }
        public string Rule { get; set; }
        public long Count { get; set; }
        public List<long> SampleRecords { get; set; } = new List<long>();
    }
    public sealed class CsvResult {
        public long Rows { get; set; }
        public int Columns { get; set; }
        public long DuplicateKeys { get; set; }
        public long InvalidValues { get; set; }
        public string SHA256 { get; set; }
        public double ScanSeconds { get; set; }
        public double HashSeconds { get; set; }
        public List<Issue> Issues { get; set; } = new List<Issue>();
    }
    public static class CsvValidation {
        static void IssueAt(CsvResult result, string column, string rule, long record) {
            var issue = result.Issues.Find(i => i.Column == column && i.Rule == rule);
            if (issue == null) { issue = new Issue { Column = column, Rule = rule }; result.Issues.Add(issue); }
            issue.Count++;
            if (issue.SampleRecords.Count < 3) issue.SampleRecords.Add(record);
        }
        static bool Valid(string value, string type) {
            long integer; double number; decimal amount; DateTime date; bool boolean;
            switch(type) {
                case "int64": return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out integer);
                case "double": return double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number) && !double.IsNaN(number) && !double.IsInfinity(number);
                case "decimal": return decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out amount);
                case "dateTime": return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
                case "boolean": return bool.TryParse(value, out boolean);
                case "string": return true;
                default: throw new InvalidDataException("Unknown contract type: " + type);
            }
        }
        public static CsvResult Check(string path, string[] names, string[] types, string[] keyNames, string tenant, bool schemaOnly) {
            if (names.Length != types.Length) throw new InvalidDataException("Invalid column contract.");
            var result = new CsvResult();
            var clock = Stopwatch.StartNew();
            // Prevent concurrent writes while scanning and hashing the same bytes.
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                using (var parser = new TextFieldParser(stream, Encoding.UTF8, true, true)) {
                    parser.SetDelimiters(","); parser.HasFieldsEnclosedInQuotes = true; parser.TrimWhiteSpace = false;
                    var header = parser.ReadFields();
                    if (header == null) throw new InvalidDataException("No header.");
                    result.Columns = header.Length;
                    var indexes = new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < header.Length; i++) {
                        if (String.IsNullOrWhiteSpace(header[i]) || indexes.ContainsKey(header[i])) throw new InvalidDataException("Empty or duplicate CSV header.");
                        indexes.Add(header[i], i);
                    }
                    var typedIndexes = new List<int>(); var typedNames = new List<string>(); var typedTypes = new List<string>();
                    for (int i = 0; i < names.Length; i++) {
                        if (!indexes.ContainsKey(names[i])) throw new InvalidDataException("Missing column: " + names[i]);
                        Valid("", types[i]); // Validate type even for an empty file.
                        if (types[i] != "string") { typedIndexes.Add(indexes[names[i]]); typedNames.Add(names[i]); typedTypes.Add(types[i]); }
                    }
                    var keyIndexes = new List<int>();
                    foreach (var name in keyNames) {
                        if (!indexes.ContainsKey(name)) throw new InvalidDataException("Missing key column: " + name);
                        keyIndexes.Add(indexes[name]);
                    }
                    int tenantIndex = -1;
                    if (!String.IsNullOrEmpty(tenant) && !indexes.TryGetValue("TenantKey", out tenantIndex)) throw new InvalidDataException("TenantKey column missing.");
                    var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var key = new StringBuilder();
                    while (!schemaOnly && !parser.EndOfData) {
                        string[] row;
                        try { row = parser.ReadFields(); }
                        catch (MalformedLineException) { throw new InvalidDataException("Malformed CSV near record " + (result.Rows + 1)); }
                        if (row == null) break;
                        result.Rows++;
                        if (row.Length != header.Length) throw new InvalidDataException("CSV field count mismatch at record " + result.Rows);
                        if (tenantIndex >= 0 && !String.Equals(row[tenantIndex], tenant, StringComparison.OrdinalIgnoreCase)) {
                            result.InvalidValues++; IssueAt(result, "TenantKey", "TenantIdentity", result.Rows);
                        }
                        if (keyIndexes.Count > 0) {
                            key.Clear(); bool blank = false;
                            foreach (int index in keyIndexes) {
                                string value = row[index].Trim(); blank |= value.Length == 0;
                                if (key.Length > 0) key.Append('|');
                                key.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value);
                            }
                            if (blank) { result.InvalidValues++; IssueAt(result, String.Join("|", keyNames), "RequiredKey", result.Rows); }
                            if (!keys.Add(key.ToString())) { result.DuplicateKeys++; IssueAt(result, String.Join("|", keyNames), "UniqueKey", result.Rows); }
                        }
                        for (int i = 0; i < typedIndexes.Count; i++) {
                            string value = row[typedIndexes[i]];
                            if (!String.IsNullOrWhiteSpace(value) && !Valid(value, typedTypes[i])) {
                                result.InvalidValues++; IssueAt(result, typedNames[i], typedTypes[i], result.Rows);
                            }
                        }
                    }
                }
                result.ScanSeconds = clock.Elapsed.TotalSeconds; clock.Restart();
                stream.Position = 0;
                using (var hash = SHA256.Create()) result.SHA256 = BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "");
                result.HashSeconds = clock.Elapsed.TotalSeconds;
            }
            return result;
        }
    }
}
