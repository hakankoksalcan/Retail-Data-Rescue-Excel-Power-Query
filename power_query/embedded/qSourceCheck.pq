let
    Ayarlar = Excel.CurrentWorkbook(){[Name="tblSource"]}[Content]{0},
    DosyaYolu = Text.From(Ayarlar[SourcePath]),

    Kaynak = Excel.Workbook(
        File.Contents(DosyaYolu),
        null,
        true
    ),

    Sayfa = Kaynak{[Item="Online Retail", Kind="Sheet"]}[Data],

    Basliklar = Table.PromoteHeaders(
        Sayfa,
        [PromoteAllScalars=true]
    ),

    Kontrol = #table(
        {"Kontrol", "Sonuç"},
        {
            {"Veri satırı sayısı", Table.RowCount(Basliklar)},
            {"Sütun sayısı", Table.ColumnCount(Basliklar)}
        }
    )
in
    Kontrol
