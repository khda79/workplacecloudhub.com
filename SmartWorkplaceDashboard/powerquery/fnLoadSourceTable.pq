// Smart Workplace Dashboard BETA 3.8.0-beta.1
(fileName as text,delimiter as text,expectedColumns as list,typeMap as list,includeHistory as logical) as table =>
let
 Now=DateTime.FixedLocalNow(),
 CurrentFiles=Table.SelectRows(fnGetSourceFiles("DATA-LAST"),each Comparer.Equals(Comparer.OrdinalIgnoreCase,[Name],fileName)),
 Current=if Table.RowCount(CurrentFiles)=0 then error Error.Record("SmartInventorySourceMissing","Required CSV not found in the selected tenant DATA-LAST.",[FileName=fileName]) else if Table.RowCount(CurrentFiles)<>1 then error Error.Record("SmartInventorySourceAmbiguous","More than one current CSV matches; select one tenant root.",[FileName=fileName]) else CurrentFiles{0},
 CurrentModified=fnToDateTime(Current[Date modified]),
 CurrentWeek=if CurrentModified=null then null else Date.StartOfWeek(Date.From(CurrentModified),Day.Monday),
 Parse=(item as record,isCurrent as logical) as table => let
  Content=Binary.Buffer(item[Content]),
  HeaderLine=try Lines.FromBinary(Content,null,null,65001){0} otherwise "",
  HeaderWith=(sep as text) as list => try Table.ToRows(Csv.Document(Text.ToBinary(HeaderLine),[Delimiter=sep,Encoding=65001,QuoteStyle=QuoteStyle.Csv])){0} otherwise {},
  Matches=(header as list) as logical => List.Count(header)=List.Count(expectedColumns) and List.Count(List.Difference(expectedColumns,header))=0 and List.Count(List.Distinct(header))=List.Count(header),
  ActualDelimiter=if Matches(HeaderWith(delimiter)) then delimiter else if Matches(HeaderWith(if delimiter="," then ";" else ",")) then (if delimiter="," then ";" else ",") else error Error.Record("SmartInventorySchemaMismatch","CSV header differs from the pinned schema; regenerate and review the contract.",[FileName=fileName]),
  Csv=Csv.Document(Content,[Delimiter=ActualDelimiter,Columns=List.Count(expectedColumns),Encoding=65001,QuoteStyle=QuoteStyle.Csv,ExtraValues=ExtraValues.Error]),
  Promoted=Table.PromoteHeaders(Csv,[PromoteAllScalars=true]),
  Base0=Table.SelectColumns(Promoted,expectedColumns),
  TenantRows=if not Table.HasColumns(Base0,"TenantKey") then Base0 else if Text.Trim(ExpectedTenantKey)="" or Table.RowCount(Table.SelectRows(Base0,each [TenantKey]=null or Text.Lower(Text.Trim(Text.From([TenantKey])))<>Text.Lower(Text.Trim(ExpectedTenantKey))))>0 then error Error.Record("SmartInventoryTenantMismatch","CSV contains missing or unexpected TenantKey values.",[FileName=fileName]) else Base0,
  Base=if Table.HasColumns(TenantRows,"IsPartialInventory") and Table.RowCount(Table.SelectRows(TenantRows,each fnToLogical([IsPartialInventory])<>false))>0 then error Error.Record("SmartInventoryPartialSource","Partial or unknown inventory completeness; totals are not qualified.",[FileName=fileName]) else TenantRows,
  FileStamp=fnToDateTime(item[Date modified]),
  DateFields=List.Select({"ReportRefreshDate","Report Refresh Date","RunDateUtc","ExportDateTime","ExecutionDate"},each List.Contains(expectedColumns,_)),
  Dates=if List.IsEmpty(DateFields) then {} else List.Transform(Table.Column(Base,DateFields{0}),each fnToDateTime(_)),
  UsableDates=List.RemoveNulls(Dates),
  Stamp0=if List.IsEmpty(DateFields) then FileStamp else if List.IsEmpty(Dates) or List.Count(UsableDates)<>List.Count(Dates) then null else List.Min(UsableDates),
  Stamp=if Stamp0<>null and Stamp0>Now then error Error.Record("SmartInventoryFutureTimestamp","Source timestamp is in the future; correct source/clock before refresh.",[FileName=fileName]) else Stamp0,
  ValidatedBase=if (FileStamp<>null and FileStamp>Now) or List.AnyTrue(List.Transform(UsableDates,each _>Now)) then error Error.Record("SmartInventoryFutureTimestamp","Source timestamp is in the future; correct source/clock before refresh.",[FileName=fileName]) else Base,
  A=Table.AddColumn(ValidatedBase,"__SnapshotDate",each if FileStamp=null then null else DateTime.From(Date.From(FileStamp)),type nullable datetime),
  B=Table.AddColumn(A,"__SnapshotDateTime",each Stamp,type nullable datetime),
  C=Table.AddColumn(B,"__SnapshotPeriod",each if FileStamp=null then null else Date.ToText(Date.StartOfWeek(Date.From(FileStamp),Day.Monday),"yyyy-MM-dd"),type nullable text),
  D=Table.AddColumn(C,"__IsCurrent",each isCurrent,type logical),
  E=Table.AddColumn(D,"__SourceFile",each fileName,type text),
  F=Table.AddColumn(E,"__SourceFolder",each item[Folder Path],type nullable text)
 in F,
 Historical=if not includeHistory or CurrentModified=null then #table({}, {}) else let
  Earliest=Date.AddMonths(Date.From(CurrentModified),-Number.RoundDown(HistoryMonths)),
  Candidates=Table.SelectRows(fnGetSourceFiles("DATA-ALL"),each Comparer.Equals(Comparer.OrdinalIgnoreCase,[Name],fileName) and [Date modified]<>null and Date.From([Date modified])>=Earliest and Date.From([Date modified])<CurrentWeek and Text.Contains(Text.Replace([Folder Path],"\","/"),"/WeeklyHistory/",Comparer.OrdinalIgnoreCase)),
  WithPeriod=Table.AddColumn(Candidates,"__Period",each Date.ToText(Date.StartOfWeek(Date.From([Date modified]),Day.Monday),"yyyy-MM-dd"),type text),
  Grouped=Table.Group(WithPeriod,{"__Period"},{{"Latest",each let Sorted=Table.Sort(_,{{"Date modified",Order.Descending}}),Top=Sorted{0}[Date modified] in if Table.RowCount(Table.SelectRows(Sorted,each [Date modified]=Top))>1 then error Error.Record("SmartInventoryHistoryAmbiguous","Historical timestamps tie; remove ambiguity.",[FileName=fileName]) else Table.FirstN(Sorted,1)}})
 in if Table.RowCount(Grouped)=0 then Candidates else Table.Combine(Grouped[Latest]),
 CurrentTable=Parse(Current,true),
 MinimumHistoryRows=Number.RoundUp(Table.RowCount(CurrentTable)*HistoryMinRowRatio),
 CandidateHistoryTables=if includeHistory and Table.RowCount(Historical)>0 then List.Transform(Table.ToRecords(Historical),each Parse(_,false)) else {},
 HistoryTables=List.Select(CandidateHistoryTables,each Table.RowCount(_)>=MinimumHistoryRows),
 Combined=Table.Combine(List.Combine({HistoryTables,{CurrentTable}})),
 Typed=List.Accumulate(typeMap,Combined,(state as table,item as list)=>if item{1}="logical" then Table.TransformColumns(state,{{item{0},fnToLogical,type nullable logical}}) else if item{1}="datetime" then Table.TransformColumns(state,{{item{0},fnToDateTime,type nullable datetime}}) else if item{1}="number" or item{1}="double" then Table.TransformColumns(state,{{item{0},fnToNumber,type nullable number}}) else if item{1}="int64" then Table.TransformColumns(state,{{item{0},fnToInt64,Int64.Type}}) else state)
in Typed
