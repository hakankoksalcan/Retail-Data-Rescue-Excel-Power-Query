# Veri sözlüğü ve göstergeler

| Alan | Tanım | Köken |
| --- | --- | --- |
| SourceFile/SourceSheet/SourceRow | Özgün dosya, sayfa ve Excel satır numarası; başlık 1 | Kaynak |
| RecordID | Dosya / sayfa / dokuz haneli satır; dosya değişince kimlik de değişir | Türetilmiş |
| *_Raw | Sekiz özgün alan; impute veya yeniden yazma yok | Kaynak kopyası |
| InvoiceNo/StockCode/CustomerID | Metin kimlikler; boş müşteri null kalır | Türetilmiş |
| DescriptionTrim | Yalnız dış boşlukları kırpılmış metin; metin olmayan değer null | Türetilmiş; özgün değer korunur |
| Quantity/UnitPrice/InvoiceDate/Country | Kaynakla aynı anlamsal değerler; tipler kontrollü | Kaynak |
| PriceUnits4 | Birim fiyat ×10000; tam sayı | Türetilmiş |
| AmountUnits4 | Quantity × PriceUnits4; tam sayı | Türetilmiş |
| LineAmount | AmountUnits4 /10000; görünüm iki ondalık olabilir | Türetilmiş |
| Month | InvoiceDate üzerinden yyyy-MM | Türetilmiş |
| DescriptionVariantCount | Kodun dolu ham metin açıklamalarının farklı sayısı | Kaynak tablodan türetilmiş sözlük |
| CodeCategory | Açık kod sözlüğü, genel ürün veya eşlenmemiş kod | Belgelenmiş iş kuralı |
| TransactionClass | Tek, öncelik kurallı uzlaştırma sınıfı | Belgelenmiş iş kuralı |
| ReportScope | MerchandiseSale veya MerchandiseCancellation | Belgelenmiş kapsam |
| CustomerAnalysisEligible | MerchandiseSale ve CustomerID mevcut | Belgelenmiş kapsam |
| NeedsReview/ReviewReasons | Örtüşebilen tanı bayraklarının birleşimi ve nedenleri | Türetilmiş; dışlama anlamına gelmez |

| Metric | Hesap | Kapsam | Sınırlama |
| --- | --- | --- | --- |
| ReportableSalesGBP | MerchandiseSale satırlarının AmountUnits4 toplamı /10000 | Pozitif miktar/fiyat; özel işlemler ayrılmış; tam tekrarlar dahil | Ödeme veya muhasebe geliri değildir |
| MerchandiseCancellationGBP | MerchandiseCancellation işaretli toplamının negatifi /10000 | Ürün kapsamındaki C iptalleri | Fiziksel iade değildir; 9288 C kaydının hepsi dahil değil |
| MerchandiseBalanceGBP | Satış tutarı - ürün iptal tutarı | Aynı ürün kapsamının işaretli dengesi | Net sales veya kâr olarak etiketlenmez |
| EligibleSalesInvoiceIDs | Satış kapsamındaki farklı InvoiceNo sayısı | Satır sayısından ayrı | Satış kalemli fatura; ödeme/teslimat teyidi yok |
| KnownCustomerRowCoverage | Bilinen CustomerID satış satırı / bütün satış satırları | Satır ağırlıklı kapsama | Tutar kapsamı ile aynı ölçü değildir |
| KnownCustomerAmountCoverage | Bilinen CustomerID satış tutarı / bütün satış tutarı | Tutar ağırlıklı kapsama | Tam müşteri evrenini temsil ettiği varsayılmaz |
| ReviewRows | NeedsReview=true farklı RecordID sayısı | Bayrakların OR birleşimi | Satış kapsamıyla örtüşür; dışlanan kayıt sayısı değildir |
| RawSignedAmountGBP | Bütün kaynak Quantity × UnitPrice işaretli toplamı | Özel işlemler ve tekrarlar dahil | Kontrol toplamıdır |
| DuplicateScenario | Ana sonuç ve her grupta ilk SourceRow korunmuş alternatif | Sekiz özgün alan tam eşitliği | Alternatif ana sonucu değiştirmez |

ProductDictionary, Customers, DashboardData ve diğer rapor tabloları tek kaynak işlem tablosundan türetilmiştir. Bağımsız ürün, müşteri, ödeme, maliyet, stok veya teslimat sistemi değildir. Ürün grafiklerinde açıklama belirsizliği gizlenmeden StockCode kullanılır.

Country/Month küpünde uygun fatura sayısı toplanır. Aynı satış faturasının farklı ülke/ay bölümlerine dağılmadığı kontrol edilir. Bu kontrol yeni kaynakta başarısızsa küp fatura sayısını kullanmadan önce modeli güncelle; benzersiz sayılar genel durumda toplamsal değildir.
