// Smart Workplace Dashboard BETA 3.8.0-beta.1
(relativeFolder as text) as table =>
let
 Valid = List.Contains({"LocalFolder","SharePoint"},SourceMode)
  and Text.Trim(ExpectedTenantKey)<>""
  and List.Contains({"DATA-LAST","DATA-ALL"},relativeFolder)
  and HistoryMonths>=0 and HistoryMonths<=120 and Number.RoundDown(HistoryMonths)=HistoryMonths
  and HistoryMinRowRatio>=0 and HistoryMinRowRatio<=1
  and (if SourceMode="LocalFolder" then Text.Trim(DataRootPath)<>"" else Text.StartsWith(SharePointSiteUrl,"https://",Comparer.OrdinalIgnoreCase) and Text.StartsWith(SharePointDataFolderUrl,Text.TrimEnd(SharePointSiteUrl,{"/"})&"/",Comparer.OrdinalIgnoreCase)),
 Root = if Valid then Text.TrimEnd(DataRootPath,{"\","/"}) else error Error.Record("SmartInventoryConfiguration","Select a source mode, one tenant root and ExpectedTenantKey; check history parameters.",null),
 LocalPath=Root&"\"&relativeFolder,
 Empty=#table({"Name","Content","Extension","Folder Path","Date modified","Size"},{}),
 LocalFiles=if SourceMode<>"LocalFolder" then Empty else if relativeFolder="DATA-LAST" then Folder.Contents(LocalPath) else Folder.Files(LocalPath),
 SharePointFiles=if SourceMode<>"SharePoint" then Empty else if not Valid then error Error.Record("SmartInventoryConfiguration","Invalid SharePoint or tenant settings.",null) else let
  All=SharePoint.Files(SharePointSiteUrl,[ApiVersion=15]), Prefix=Text.TrimEnd(SharePointDataFolderUrl,{"/"})&"/"&relativeFolder&"/"
 in Table.SelectRows(All,each if relativeFolder="DATA-LAST" then Comparer.Equals(Comparer.OrdinalIgnoreCase,[Folder Path],Prefix) else Text.StartsWith([Folder Path],Prefix,Comparer.OrdinalIgnoreCase)),
 Selected=if SourceMode="SharePoint" then SharePointFiles else LocalFiles
in if not Valid then error Error.Record("SmartInventoryConfiguration","Invalid source mode, tenant key, source root or history parameters.",null) else Table.SelectColumns(Table.SelectRows(Selected,each [Extension]<>null and Text.Lower([Extension])=".csv"),{"Name","Content","Extension","Folder Path","Date modified","Size"},MissingField.UseNull)
