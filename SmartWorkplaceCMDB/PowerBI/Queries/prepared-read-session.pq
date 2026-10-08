// Reusable loopback function; installed in the privately migrated model.
// Every partition must receive the same private session parameters.
(BaseUrl as text, SessionToken as text, BatchSHA256 as text,
 ExpectedIdentity as record, TableName as text) as table =>
let
    PortText = if Text.StartsWith(BaseUrl, "http://127.0.0.1:")
        then Text.Range(BaseUrl, Text.Length("http://127.0.0.1:")) else "",
    Port = try Number.FromText(PortText, "en-US") otherwise null,
    ValidBase = PortText <> "" and Text.Select(PortText, {"0".."9"}) = PortText
        and Port <> null and Port >= 1 and Port <= 65535,
    Prefix = if ValidBase then "session/" & SessionToken & "/" & BatchSHA256 & "/"
        else error "A loopback read session is required.",
    // The session serves immutable validated bytes. Stream large CSVs instead
    // of duplicating them in a Power Query binary buffer.
    Read = (Name as text) as binary => Web.Contents(BaseUrl,
        [RelativePath = Prefix & Name, IsRetry = true, Timeout = #duration(0, 0, 2, 0)]),
    Contract = Json.Document(Read("model-contract.json.txt")),
    IdentityFields = {"TenantKey", "OrganizationKey", "EnvironmentKey", "TenantId"},
    IdentityMatches = List.Sort(Record.FieldNames(ExpectedIdentity)) = List.Sort(IdentityFields)
        and List.AllTrue(List.Transform(IdentityFields, (Field) =>
            Value.Is(Record.Field(ExpectedIdentity, Field), type text)
            and Text.Trim(Record.Field(ExpectedIdentity, Field)) <> ""
            and Record.Field(Contract[Identity], Field) = Record.Field(ExpectedIdentity, Field))),
    Name = TableName & ".csv",
    Definition = if Contract[ProtocolVersion] = "0.1.2"
        and Contract[Owner] = "SmartInventory-CMDB-Prepared"
        and Contract[Status] = "Validated" and Contract[BatchSHA256] = BatchSHA256
        and IdentityMatches and Record.HasFields(Contract[Tables], Name)
        then Record.Field(Contract[Tables], Name)
        else error "Read-session identity, batch or table mismatch.",
    Raw = Csv.Document(Read(Name), [Delimiter = ",", Columns = List.Count(Definition[Columns]),
        Encoding = 65001, QuoteStyle = QuoteStyle.Csv, ExtraValues = ExtraValues.Error]),
    Rows = Table.PromoteHeaders(Raw, [PromoteAllScalars = true]),
    Checked = if Table.ColumnNames(Rows) = Definition[Columns]
        and Table.RowCount(Rows) = Definition[Rows]
        then Rows else error "Read-session CSV schema or row count mismatch."
in
    Checked
