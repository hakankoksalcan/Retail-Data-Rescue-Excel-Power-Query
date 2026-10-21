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

    KaynakSatiri = Table.AddIndexColumn(
        Basliklar,
        "SourceRow",
        2,
        1,
        Int64.Type
    )
in
    KaynakSatiri
