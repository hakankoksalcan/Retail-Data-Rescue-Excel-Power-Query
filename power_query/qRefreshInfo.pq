let
    Failed = Table.SelectRows(qModel[Controls], each [Status] = "FAIL"),
    Result = if Table.RowCount(Failed) > 0 then
        error Error.Record("FailedControls", "Review qControls before using this refresh", Failed)
    else #table(type table [ExecutedAtUTC=text,Mode=text], {
        {DateTimeZone.ToText(DateTimeZone.FixedUtcNow()),"Excel Power Query; controls evaluated"}
    })
in
    Result
