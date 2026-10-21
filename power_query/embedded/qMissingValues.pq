let
    Alanlar = {
        "InvoiceNo", "StockCode", "Description",
        "Quantity", "InvoiceDate", "UnitPrice",
        "CustomerID", "Country"
    },

    Kaynak = Table.SelectColumns(qRawSource, Alanlar),
    SatirSayisi = Table.RowCount(Kaynak),

    Kayitlar = List.Transform(
        Alanlar,
        (Alan) => [
            Field = Alan,
            MissingRows =
                SatirSayisi -
                List.NonNullCount(Table.Column(Kaynak, Alan))
        ]
    ),

    Sonuc = Table.FromRecords(Kayitlar),

    Tipler = Table.TransformColumnTypes(
        Sonuc,
        {
            {"Field", type text},
            {"MissingRows", Int64.Type}
        }
    )
in
    Tipler
