// Strict typed projection used by the migrated model; pass the complete column map.
// No missing-field suppression, locale inference, or blank-to-zero conversion.
(Rows as table, Types as list) as table =>
let
    // Wall-clock support is opt-in per source column, never inferred globally.
    Parse = (Value as any, Kind as text, SourceWallClock as logical, EmptyAutopilotContact as logical) as any =>
        if Value = null or Value = "" then null
        else if Kind = "String" then Text.From(Value)
        else if Kind = "Boolean" then
            if Text.Lower(Text.From(Value)) = "true" then true
            else if Text.Lower(Text.From(Value)) = "false" then false
            else error "Invalid Boolean evidence."
        else if Kind = "Double" then
            let Number = Number.FromText(Text.From(Value), "en-US")
            in if Number.IsNaN(Number) or Number = #infinity or Number = -#infinity
                then error "Non-finite numeric evidence." else Number
        else if Kind = "Int64" then
            let Number = Number.FromText(Text.From(Value), "en-US")
            in if not Number.IsNaN(Number) and Number <> #infinity and Number <> -#infinity
                and Number.Abs(Number) <= 9007199254740991 and Number.RoundTowardZero(Number) = Number
                then Int64.From(Number) else error "Invalid or inexact integer evidence."
        else if Kind = "DateOnly" then DateTime.From(
            Date.FromText(Text.From(Value), [Format = "yyyy-MM-dd", Culture = "en-US"]))
        else if Kind = "DateTime" then
            let TextValue = Text.From(Value), Offset = Text.End(TextValue, 6),
                ExplicitZone = Text.EndsWith(TextValue, "Z")
                    or (Text.Length(Offset) = 6
                        and List.Contains({"+", "-"}, Text.Start(Offset, 1))
                        and Text.At(Offset, 3) = ":"
                        and Text.Length(Text.Select(Offset, {"0".."9"})) = 4),
                WallClock = () as datetime =>
                    let Parsed = DateTime.FromText(TextValue,
                            [Format = "dd/MM/yyyy HH:mm:ss", Culture = "fr-FR"]),
                        Exact = DateTime.ToText(Parsed,
                            [Format = "dd/MM/yyyy HH:mm:ss", Culture = "fr-FR"]) = TextValue
                    in if Exact and Date.Year(DateTime.Date(Parsed)) >= 1900 then Parsed
                        else error "Invalid source wall-clock date evidence."
            in if EmptyAutopilotContact and TextValue = "01/01/0001 00:00:00" then null
                else if ExplicitZone then DateTimeZone.RemoveZone(
                DateTimeZone.SwitchZone(DateTimeZone.FromText(TextValue, [Culture = "en-US"]), 0))
                else if SourceWallClock then WallClock()
                else error "Date evidence requires an explicit timezone."
        else error "Unsupported candidate type.",
    Converted = Table.TransformColumns(Rows, List.Transform(Types, (Definition) =>
        {Definition[name], (Value) => Parse(Value, Definition[dataType],
            Record.FieldOrDefault(Definition, "allowSourceWallClock", false),
            Record.FieldOrDefault(Definition, "allowEmptyAutopilotContact", false)),
            if Definition[dataType] = "String" then type nullable text
            else if Definition[dataType] = "Boolean" then type nullable logical
            else if Definition[dataType] = "Double" then type nullable number
            else if Definition[dataType] = "Int64" then Int64.Type
            else type nullable datetime}), null, MissingField.Error)
in
    Converted
