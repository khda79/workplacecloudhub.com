// Internal transport; callers validate path segments. phase distinguishes selection and verification in synthetic tests.
(segments as list, optional phase as nullable text) as table =>
        if List.IsEmpty(segments) then
            let
                // A fresh HTTP read is required: connector navigation can reuse cached content.
                LibrariesResponse = Json.Document(Web.Contents(SharePointSiteUrl, [
                    RelativePath="_api/web/lists",
                    Query=[#"$select"="Title,RootFolder/Name,RootFolder/ServerRelativeUrl", #"$expand"="RootFolder", #"$filter"="BaseType eq 1"],
                    Headers=[Accept="application/json;odata=nometadata"], IsRetry=true])),
                Libraries = if Record.HasFields(LibrariesResponse, "odata.nextLink") or Record.HasFields(LibrariesResponse, "@odata.nextLink") then
                    error Error.Record("Prepared evidence", "Library discovery is incomplete", null) else LibrariesResponse[value],
                Matches = List.Select(Libraries, each [Title] = SharePointLibraryName or [RootFolder][Name] = SharePointLibraryName),
                Library = if List.Count(Matches) = 1 then Matches{0}[RootFolder] else
                    error Error.Record("Prepared evidence", "Missing or ambiguous document library", SharePointLibraryName),
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
