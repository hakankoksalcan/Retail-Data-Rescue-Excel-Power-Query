# Veri kalitesi ve uzlaştırma sonuçları

Doğrulama motoru: Python referans. Tam raporlama modelinin Excel Power Query yürütmesi bekliyor. Kaynak/eksik değer denetimleri Excel 2021 içinde doğrulandı; validation/README.md kapsamı açıklar. Kaynak checksum'u özgün Aşama 1 sonucu ile eşleşti. Önceki sayıların tümü yeniden okuma ile karşılaştırıldı; fark yok. Ayrıntılı özgün denetim: stage1-VERI_UYGUNLUK_DENETIMI.md.

## Kaynak ve kapsam

541909 satır, 8 sütun; 25900 farklı InvoiceNo. 135080 CustomerID, 1454 Description eksik. 4879 tam tekrar grubu, 10147 üye, ilk kayıt dışındaki 5268 ek tekrar. 9288 C kaydı, 1336 C olmayan negatif miktar, 2515 sıfır fiyat, 2 negatif fiyat. 650 kodun birden fazla açıklaması, 113452 açıklamanın çevresel boşluğu var. Gerçek bozuk tarih bulunmadı. Kaynak metadata 'eksik değer yok' diyor; dosya eksiklerini esas aldık.

## Kaynağın tamamı açıklanıyor

| TransactionClass | Rows | InvoiceIDs | AmountUnits4 | SignedAmountGBP | ReviewFlagRows |
| --- | --- | --- | --- | --- | --- |
| MerchandiseSale | 527761 | 19773 | 102714330630 | 10271433.063 | 226068 |
| MerchandiseCancellation | 8704 | 3422 | -4787241800 | -478724.18 | 2572 |
| SpecialOperation | 2946 | 2769 | -449609490 | -44960.949 | 1043 |
| ZeroPriceAdjustment | 1335 | 1335 | 0 | 0.0 | 1335 |
| UnresolvedReview | 1163 | 805 | 0 | 0.0 | 1163 |

Sınıf satırları toplamı **541909**, işaretli tutarlar toplamı **9747747.934 GBP**. InvoiceIDs sütununu sınıflar arasında toplama; aynı InvoiceNo birden fazla sınıfta kalem içerebilir. 1336 C olmayan negatif miktardan 1335'i düzeltme sınıfındadır; sayısal Description içeren satır 420393 öncelik gereği UnresolvedReview'dır. 9288 C kaydının 8704'ü ürün iptali; diğerleri özel işlemler kapsamında ayrılır.

Ham toplamdan **SpecialOperation = -44960.949 GBP** ayrılınca ürün işaretli dengesi **9792708.883 GBP** olur. Diğer iki ayrı sınıfın tutarı sıfırdır; satırlarının incelenmesi yine önemlidir. Özel toplamın negatif olması nedeniyle bu ayrıştırma ham toplamdan daha yüksek ürün dengesi üretir; bu bir muhasebe geliri düzeltmesi iddiası değildir.

## Tam tekrar duyarlılığı

| Scenario | SourceRows | RawAmountUnits4 | RawSignedGBP | SalesGBP | CancellationGBP | BalanceGBP | SalesInvoiceIDs | CustomerSalesIDs | ReviewRows |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Main: retain all rows | 541909 | 97477479340 | 9747747.934 | 10271433.063 | 478724.18 | 9792708.883 | 19773 | 4334 | 232181 |
| Alternative: first exact record only | 536641 | 97260069540 | 9726006.954 | 10247219.323 | 475901.16 | 9771318.163 | 19773 | 4334 | 226913 |

İlk kayıt korunsaydı 5268 satır ve ham işaretli tutardan **21740.98 GBP** ayrılırdı. Satış **24213.74 GBP**, ürün iptal büyüklüğü **2823.02 GBP**, ürün dengesi **21390.72 GBP** azalırdı. Uygun fatura ve bilinen satış müşteri sayıları değişmedi. Bu alternatif ana sonuç değildir.

## Kalite bayrakları

| Flag | Kind | Rows | SourceRowExample |
| --- | --- | --- | --- |
| HasMissingCustomer | review | 135080 | 624.0 |
| HasMissingDescription | review | 1454 | 624.0 |
| HasNonTextDescription | review | 1 | 420393.0 |
| HasOuterWhitespace | format | 113452 | 12.0 |
| IsExactDuplicateMember | review | 10147 | 487.0 |
| IsDuplicateBeyondFirst | scenario | 5268 | 519.0 |
| IsDescriptionVariantDuplicate | review | 4 | 199126.0 |
| HasMultipleDescriptions | review | 113465 | 2.0 |
| IsNonCNegative | review | 1336 | 2408.0 |
| IsZeroPrice | review | 2515 | 624.0 |
| IsNegativePrice | review | 2 | 299985.0 |
| IsLargeAmount | review | 17 | 15018.0 |
| IsUnmappedCode | review | 0 | nan |
| IsInvoicePatternException | context | 3 | 299984.0 |
| IsCancellationCode | context | 9288 | 143.0 |

İnceleme listesi **232181 farklı kaynak kaydı** içerir; bu sayı bayrakların toplamı değildir. Listeyle raporlama kapsamı örtüşür. Her kaydın kaynak kimliği ve gerekçeleri korunur. CSV örnekler ve BeforeAfter sayfası aynı kaynak satırlarına bağlıdır. Büyük tutar, açıklama varyantı veya müşteri eksikliği otomatik silme nedeni değildir.

## Kontroller

reference/Controls.csv beklenen ve gerçekleşen değerleri içerir. Temel referans 29 kontrolü geçti; dashboard küpünün fatura bölümü için ek kontrol geçti. Gerçek Excel çalışması için tek PENDING satırı vardır. validation/refresh-tests.json içindeki 19 Python tekrar/ekleme kontrolü geçti. M kaynaklarında syntax PASS; semantik ve gerçek Excel yenileme PENDING.

Finansal hesaplar 1/10000 GBP tamsayı birimleriyle yapılır. Sınıf ve senaryo uzlaştırmasında tamsayı farkı sıfır; gösterge/formül kıyaslarında GBP toleransı 0.0001. Excel ekranındaki iki ondalık ile tam hassasiyetin görünüm farkı veri değişimi değildir.

Bu uzlaştırma yalnız kaynak dosyanın kendi içindedir; banka, ödeme, muhasebe, stok veya teslimat sistemiyle çapraz doğrulama yapılmadı.
