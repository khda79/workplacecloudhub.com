// Internal transport; callers validate path segments. phase distinguishes selection and verification in synthetic tests.
(segments as list, optional phase as nullable text) as table =>
        if List.IsEmpty(segments) then
            let
                // A fresh HTTP read is required: connector navigation can reuse cached content.
                Library = Json.Document(Web.Contents(SharePointSiteUrl, [
                    RelativePath="_api/web/lists/GetByTitle('" & Uri.EscapeDataString(Text.Replace(SharePointLibraryName, "'", "''")) & "')/RootFolder",
                    Query=[#"$select"="ServerRelativeUrl"], Headers=[Accept="application/json;odata=nometadata"], IsRetry=true])),
                RootPath = Library[ServerRelativeUrl] & "/" & Text.Trim(SharePointDataFolder & "/" & PreparedSharePointFolder, "/"),
                Fetch = (name as text) as nullable binary =>
                    let
                        Response = Web.Contents(SharePointSiteUrl, [
                            RelativePath="_api/web/GetFileByServerRelativeUrl('" & Uri.EscapeDataString(Text.Replace(RootPath & "/" & name, "'", "''")) & "')/$value",
                            ManualStatusHandling={404}, IsRetry=true]),
                        Status = Record.FieldOrDefault(Value.Metadata(Response), "Response.Status", 200)
                    in if Status = 404 then null else Binary.Buffer(Response),
                Preferred = Fetch("current.json.txt"),
                Name = if Preferred = null then "current.json" else "current.json.txt",
                Content = if Preferred = null then Fetch("current.json") else Preferred
            in if Content = null then #table({"Name", "Content"}, {}) else #table({"Name", "Content"}, {{Name, Content}})
        else
            let
                Root = SharePoint.Contents(SharePointSiteUrl, [ApiVersion=15]),
                Path = {SharePointLibraryName} & List.Select(Text.Split(SharePointDataFolder & "/" & PreparedSharePointFolder, "/"), each _ <> "") & segments,
                Content = List.Accumulate(Path, Root, (parent, segment) =>
                    let found = Table.SelectRows(parent, each [Name] = segment)
                    in if Table.RowCount(found) = 1 then found{0}[Content] else error Error.Record("Prepared evidence", "Missing or ambiguous folder; wait for a complete batch", segment))
            in Content
