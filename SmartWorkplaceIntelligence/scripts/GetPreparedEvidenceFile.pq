(fileName as text) as binary =>
let
    SafeSegment = (s as text) as text => if s = "" or Text.Contains(s, "/") or Text.Contains(s, "\") or Text.Contains(s, "..") then error "Invalid prepared path segment" else s,
    ReadFolder = (segments as list) as table =>
        if PreparedSourceMode = "Local" then
            Folder.Contents(Text.TrimEnd(PreparedEvidenceRoot, {"\", "/"}) & (if List.IsEmpty(segments) then "" else "\" & Text.Combine(List.Transform(segments, SafeSegment), "\")))
        else if PreparedSourceMode = "SharePoint" then
            let
                Root = SharePoint.Contents(SharePointSiteUrl, [ApiVersion=15]),
                Path = {SharePointLibraryName} & List.Select(Text.Split(SharePointDataFolder & "/" & PreparedSharePointFolder, "/"), each _ <> "") & List.Transform(segments, SafeSegment),
                Content = List.Accumulate(Path, Root, (parent, segment) =>
                    let found = Table.SelectRows(parent, each [Name] = segment)
                    in if Table.RowCount(found) = 1 then found{0}[Content] else error Error.Record("Prepared evidence", "Missing or ambiguous folder; wait for a complete batch", segment))
            in Content
        else error "PreparedSourceMode must be Local or SharePoint",
    ReadBinary = (segments as list) as binary =>
        if PreparedSourceMode = "Local" then
            File.Contents(Text.TrimEnd(PreparedEvidenceRoot, {"\", "/"}) & "\" & Text.Combine(List.Transform(segments, SafeSegment), "\"))
        else
            let
                Parent = ReadFolder(List.RemoveLastN(segments, 1)),
                Found = Table.SelectRows(Parent, each [Name] = SafeSegment(List.Last(segments)))
            in if Table.RowCount(Found) = 1 then Found{0}[Content] else error "Missing or ambiguous prepared file",
    ReadMetadata = (segments as list, name as text) as binary =>
        let
            Parent = ReadFolder(segments),
            Preferred = Table.SelectRows(Parent, each [Name] = name & ".json.txt"),
            Legacy = Table.SelectRows(Parent, each [Name] = name & ".json"),
            // Fallback only on ABSENCE, never after a read or JSON parse error.
            Selected = if Table.RowCount(Preferred) = 0 then Legacy else Preferred
        in if Table.RowCount(Selected) = 1 then Binary.Buffer(Selected{0}[Content]) else error "Missing or ambiguous prepared metadata",
    PointerBytes = ReadMetadata({}, "current"),
    Pointer = Json.Document(PointerBytes),
    BatchId = SafeSegment(if PreparedBatchId = "" then Text.From(Pointer[BatchId]) else PreparedBatchId),
    Manifest = Json.Document(ReadMetadata({"batches", BatchId}, "batch")),
    Selected = List.Select(Manifest[Files], each [File] = SafeSegment(fileName)),
    Entry = if Manifest[SchemaVersion] <> 1 or Manifest[BatchId] <> BatchId or Manifest[TenantKey] <> Pointer[TenantKey] or List.Count(Selected) <> 1 then error "Invalid prepared batch manifest" else Selected{0},
    Content = Binary.Buffer(ReadBinary({"batches", BatchId, SafeSegment(fileName)})),
    // Size guard rejects incomplete synchronized files. SHA256 is checked by the publisher, not by M.
    Checked = if Binary.Length(Content) <> Entry[Bytes] then error "Prepared CSV transfer incomplete; previous imported model must be retained" else Content,
    StablePointer = if PreparedBatchId <> "" or Binary.ToText(PointerBytes, BinaryEncoding.Base64) = Binary.ToText(ReadMetadata({}, "current"), BinaryEncoding.Base64) then Checked else error "Prepared batch changed during read; refresh after preparation finishes"
in
    StablePointer
