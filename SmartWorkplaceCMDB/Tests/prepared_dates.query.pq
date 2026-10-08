// Execute the actual candidate function with PQTest's parameterQueryFile option.
// Synthetic in-memory data only: no CSV, reader session, or tenant access.
(Convert as function) as table =>
let
    Definition = (Kind as text, Wall as logical, EmptyContact as logical) as record =>
        [name = "Value", dataType = Kind, allowSourceWallClock = Wall,
            allowEmptyAutopilotContact = EmptyContact],
    Read = (Value as any, D as record) as any =>
        Convert(#table({"Value"}, {{Value}}), {D}){0}[Value],
    D = Definition("DateTime", false, false),
    W = Definition("DateTime", true, false),
    A = Definition("DateTime", true, true),
    Cases = {
        [Name="French day first", Actual=() => Read("31/07/2026 08:00:34", W), Expected=#datetime(2026,7,31,8,0,34)],
        [Name="French ambiguous month", Actual=() => Read("07/10/2026 09:02:38", W), Expected=#datetime(2026,10,7,9,2,38)],
        [Name="No invented DST zone", Actual=() => Read("25/10/2026 02:30:00", W), Expected=#datetime(2026,10,25,2,30,0)],
        [Name="Leap day", Actual=() => Read("29/02/2024 12:00:00", W), Expected=#datetime(2024,2,29,12,0,0)],
        [Name="UTC unchanged", Actual=() => Read("2026-10-07T09:02:38Z", D), Expected=#datetime(2026,10,7,9,2,38)],
        [Name="Positive offset UTC", Actual=() => Read("2026-10-07T11:02:38+02:00", W), Expected=#datetime(2026,10,7,9,2,38)],
        [Name="Negative offset UTC", Actual=() => Read("2026-10-07T07:02:38-02:00", A), Expected=#datetime(2026,10,7,9,2,38)],
        [Name="Autopilot unknown sentinel", Actual=() => Read("01/01/0001 00:00:00", A), Expected=null],
        [Name="Valid Autopilot contact", Actual=() => Read("06/10/2026 12:34:56", A), Expected=#datetime(2026,10,6,12,34,56)],
        [Name="Null unchanged", Actual=() => Read(null, W), Expected=null],
        [Name="Blank unchanged", Actual=() => Read("", A), Expected=null],
        [Name="Boolean true", Actual=() => Read("TRUE", Definition("Boolean",false,false)), Expected=true],
        [Name="Boolean false", Actual=() => Read("false", Definition("Boolean",false,false)), Expected=false],
        [Name="Double unchanged", Actual=() => Read("1234.5", Definition("Double",false,false)), Expected=1234.5],
        [Name="Integer unchanged", Actual=() => Read("42", Definition("Int64",false,false)), Expected=42],
        [Name="DateOnly unchanged", Actual=() => Read("2026-10-07", Definition("DateOnly",false,false)), Expected=#datetime(2026,10,7,0,0,0)],
        [Name="French not globally allowed", Actual=() => Read("07/10/2026 09:02:38", D), Error=true],
        [Name="Sentinel not globally allowed", Actual=() => Read("01/01/0001 00:00:00", D), Error=true],
        [Name="Sentinel not a normal wall clock", Actual=() => Read("01/01/0001 00:00:00", W), Error=true],
        [Name="Other ancient date rejected", Actual=() => Read("02/01/0001 00:00:00", A), Error=true],
        [Name="Invalid leap day", Actual=() => Read("29/02/2026 12:00:00", W), Error=true],
        [Name="Invalid month", Actual=() => Read("07/13/2026 09:02:38", W), Error=true],
        [Name="Invalid hour", Actual=() => Read("07/10/2026 25:02:38", W), Error=true],
        [Name="Unpadded date rejected", Actual=() => Read("7/10/2026 09:02:38", W), Error=true],
        [Name="Whitespace rejected", Actual=() => Read("07/10/2026 09:02:38 ", W), Error=true],
        [Name="ISO without zone rejected", Actual=() => Read("2026-10-07T09:02:38", W), Error=true],
        [Name="Malformed date not suppressed", Actual=() => Read("not a date", A), Error=true],
        [Name="Invalid Boolean still rejected", Actual=() => Read("yes", Definition("Boolean",false,false)), Error=true],
        [Name="Nonfinite still rejected", Actual=() => Read("NaN", Definition("Double",false,false)), Error=true],
        [Name="Fractional integer rejected", Actual=() => Read("1.5", Definition("Int64",false,false)), Error=true],
        [Name="Inexact integer rejected", Actual=() => Read("9007199254740992", Definition("Int64",false,false)), Error=true],
        [Name="Missing column remains an error", Actual=() => Table.RowCount(Convert(#table({"Other"},{{1}}), {D})), Error=true],
        [Name="Default flags remain strict", Actual=() => Read("07/10/2026 09:02:38", [name="Value",dataType="DateTime"]), Error=true],
        [Name="Rows and null preserved", Actual=() => Table.RowCount(Convert(#table({"Value"},{{"07/10/2026 09:02:38"},{null}}), {W})), Expected=2]
    },
    Results = List.Transform(Cases, (Case) =>
        let Result = try Case[Actual](), WantError = Record.FieldOrDefault(Case,"Error",false),
            Pass = if WantError then Result[HasError]
                else if Result[HasError] then false else Result[Value] = Case[Expected]
        in [Name=Case[Name], Passed=Pass]),
    Failures = List.Select(Results, each not [Passed]),
    Output = if List.IsEmpty(Failures) then Table.FromRecords(Results)
        else error Error.Record("PreparedDateTestFailure", "Native M conversion checks failed.", Failures)
in
    Output
