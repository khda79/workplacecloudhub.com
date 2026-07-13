// Smart Workplace Dashboard - Power Query function library
// Copy these expressions into the semantic model if Power BI Desktop needs manual repair.

TenantKey = "organization-default";
OrganizationKey = "organization";
EnvironmentKey = "default";
TenantId = "";
HistoryMonths = 12;
SmartInventoryLatestPath = "C:\\Temp\\WORKPLACE\\SMART-M365\\DATA-LAST";
SmartInventoryHistoryPath = "C:\\Temp\\WORKPLACE\\SMART-M365\\DATA-ALL";
CMDBPowerBIPath = "..\\SmartWorkplaceCMDB\\Data\\Tenants\\default\\DATA-LAST\\PowerBI";

fnLoadCsv = (folderPath as text, fileName as text) as table =>
let
    Files = try Folder.Files(folderPath) otherwise #table(type table [Name = text, Content = binary, Extension = text, Folder Path = text, Date modified = datetime], {}),
    Match = Table.SelectRows(Files, each [Name] = fileName),
    Output = if Table.RowCount(Match) = 0 then
        #table({}, {})
    else
        let
            Content = Match{0}[Content],
            Csv = Csv.Document(Content, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
            Promoted = Table.PromoteHeaders(Csv, [PromoteAllScalars = true])
        in
            Promoted
in
    Output;

fnLoadCmdbTable = (fileName as text) as table =>
    fnLoadCsv(CMDBPowerBIPath, fileName);

fnLoadHistoryFolder = (historyRootPath as text, optional fileName as nullable text) as table =>
let
    Files = try Folder.Files(historyRootPath) otherwise #table(type table [Name = text, Content = binary, Extension = text, Folder Path = text, Date modified = datetime], {}),
    CsvFiles = Table.SelectRows(Files, each Text.Lower([Extension]) = ".csv" and (fileName = null or [Name] = fileName)),
    Tables = List.Transform(
        Table.ToRecords(CsvFiles),
        (r as record) =>
            let
                Csv = Csv.Document(r[Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                Promoted = Table.PromoteHeaders(Csv, [PromoteAllScalars = true]),
                FolderParts = List.RemoveItems(Text.Split(Text.TrimEnd(r[Folder Path], {"\\", "/"}), "\\"), {""}),
                SnapshotWeek = if List.Count(FolderParts) = 0 then null else List.Last(FolderParts),
                WithWeek = Table.AddColumn(Promoted, "SnapshotWeek", each SnapshotWeek, type text),
                WithFile = Table.AddColumn(WithWeek, "SnapshotFile", each r[Name], type text),
                WithModified = Table.AddColumn(WithFile, "SnapshotModifiedDateTime", each r[Date modified], type datetime)
            in
                WithModified
    ),
    Output = if List.Count(Tables) = 0 then #table({}, {}) else Table.Combine(Tables)
in
    Output;

fnSourceFreshness = (folderPath as text, fileName as text, sourceName as text) as table =>
let
    Files = try Folder.Files(folderPath) otherwise #table(type table [Name = text, Content = binary, Extension = text, Folder Path = text, Date modified = datetime], {}),
    Match = Table.SelectRows(Files, each [Name] = fileName),
    Exists = Table.RowCount(Match) > 0,
    Modified = if Exists then Match{0}[Date modified] else null,
    Rows = if Exists then try Table.RowCount(fnLoadCsv(folderPath, fileName)) otherwise null else null,
    Output = #table(
        type table [SourceName = text, FileName = text, Exists = logical, LastModifiedDateTime = nullable datetime, RowCount = nullable number],
        {{sourceName, fileName, Exists, Modified, Rows}}
    )
in
    Output;
