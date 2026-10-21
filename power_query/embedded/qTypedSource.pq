let
    Kaynak = qRawSource,

    HamFatura = Table.DuplicateColumn(
        Kaynak,
        "InvoiceNo",
        "InvoiceNo_Raw"
    ),

    HamUrun = Table.DuplicateColumn(
        HamFatura,
        "StockCode",
        "StockCode_Raw"
    ),

    HamMusteri = Table.DuplicateColumn(
        HamUrun,
        "CustomerID",
        "CustomerID_Raw"
    ),

    Tipler = Table.TransformColumnTypes(
        HamMusteri,
        {
            {"InvoiceNo", type text},
            {"StockCode", type text},
            {"CustomerID", type text},
            {"Quantity", type number},
            {"UnitPrice", type number},
            {"InvoiceDate", type datetime},
            {"Country", type text}
        },
        "en-GB"
    )
in
    Tipler
