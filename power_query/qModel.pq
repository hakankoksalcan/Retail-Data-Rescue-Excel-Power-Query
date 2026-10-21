let
    Config = Excel.CurrentWorkbook(){[Name="tblSource"]}[Content]{0},
    Result = fnRetailModel(Text.From(Config[SourcePath]), Logical.From(Config[ValidateBaseline]))
in
    Result
