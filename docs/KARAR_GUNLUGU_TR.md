# Karar günlüğü

| DecisionID | Konu | Karar | Kaynak/Etkisi | Kontrol |
| --- | --- | --- | --- | --- |
| D01 | Kaynak bütünlüğü | Özgün XLSX korunur; önce/sonra SHA-256 doğrulanır | 541909 kaynak satırı | Kaynak dosyası hash kontrolü |
| D02 | Kayıt kimliği | Dosya + sayfa + özgün Excel satırı; indeks filtre/sıralama öncesi | 541909 benzersiz RecordID | RecordID benzersizliği ve kayıpsız sınıflama |
| D03 | Kimlik tipleri | InvoiceNo/StockCode/CustomerID metin; özgün alanlar _Raw ile korunur | Harfli kimlikler geçerli | Özgün değer ve türetilmiş alan karşılaştırması |
| D04 | Açıklama boşlukları | Yalnız çevresel boşluklar ayrı DescriptionTrim alanında kırpılır | 113452 kayıt | Satır 12 önce/sonra; özgün metin korunur |
| D05 | Açıklama anlamı | 650 kodda farklı ad/notlar tek etikete zorlanmaz; sözlük varyantları görünür | 650 kod / 113465 satır | Kod bazında ham/trim/normalize farklılık sayısı |
| D06 | Eksik müşteri | Uygun satışta kalır; müşteri analizinde hariç; yapay kimlik yok | 135080 ham kayıt; satış kapsamı ayrıca hesaplanır | Satış satırı ve tutar kapsama oranları |
| D07 | Tam tekrar | Ana sonuçta bütün kayıtlar kalır; ilk SourceRow alternatif senaryosu ayrı | 5268 ek kayıt | Satır/tutar/gösterge senaryo farkları |
| D08 | İptaller | C öneki ve negatif miktar + pozitif fiyat; özel işlemler öncelikle ayrılır | 9288 C satırından 8704 ürün iptali | Fiziksel iade iddiası yok; sınıf köprüsü |
| D09 | Özel işlemler | 11 operasyon kodu ve 5 voucher kodu açık sözlükle ayrılır | 2946 satır | CodeRules ve SpecialOperations tablosu |
| D10 | Biçim dışı ürün | 16 DCGS kodu ve PADS ürün istisnası; bilinmeyen biçim dışı kod inceleme | 45 DCGS + 4 PADS satırı | Biçim dışı kod otomatik hizmet değildir |
| D11 | Negatif fiyat | B / Adjust bad debt muhasebe düzeltmesi adayı olarak özel sınıf | 2 negatif fiyat satırı | 299985/299986 kaynak satırları |
| D12 | C olmayan negatif | Fiyat 0 ve başka öncelik yoksa ZeroPriceAdjustment | 1335 sınıf satırı; 1 sayısal açıklama Review sınıfında | Kaynakta 1336 bayrak; sınıf farkı gerekçeli |
| D13 | Sıfır fiyat | Fiyat uydurulmaz; özel işlem/düzeltme veya UnresolvedReview | 2515 bayraklı satır | Bilinen müşteri/pozitif miktar 40 kayıt |
| D14 | Büyük tutar | Mutlak tutar >=10000 GBP inceleme bayrağı; otomatik silme yok | 17 satır | 540423/540424 ürün/tutar karşıtlığı |
| D15 | Hassasiyet | PriceUnits4 ve AmountUnits4, 1 GBP =10000 tamsayı birimi | 0.001 fiyatlı 4 satır korunur | Tamsayı uzlaştırma sıfır fark; GBP tolerans 0.0001 |
| D16 | Tarihler | Gerçek datetime değerleri korunur; bozuk tarih temizliği iddiası yok | 541909 geçerli tarih | Min/max ve fatura içi farklı zamanlar |
| D17 | İnceleme kapsamı | Bayraklı kayıt satılırken de inceleme listesinde bulunabilir | 232181 benzersiz kayıt | Bayrak toplamı değil OR birleşimi; sınıf dışlama değildir |

## Birbirini dışlayan sınıf önceliği

1. Açık operasyon/voucher sözlüğü → SpecialOperation.
2. Sayısal açıklama, eşlenmemiş biçim dışı kod veya desteklenmeyen fatura biçimi → UnresolvedReview.
3. C öneki, negatif miktar, pozitif fiyat → MerchandiseCancellation.
4. C olmayan negatif miktar, sıfır fiyat → ZeroPriceAdjustment.
5. Normal altı rakamlı fatura, pozitif miktar, pozitif fiyat → MerchandiseSale.
6. Geri kalanlar → UnresolvedReview.

Kritik şema/tarih/sayı hatası veya dört ondalıktan fazla fiyat hassasiyeti, yeni kaynağın desteklenen sözleşmesini aşar; kaynak değiştirilmeden yükleme durdurulur. Bu engel sıradan eksik müşteri kimliği değildir. Yeni fiyat hassasiyeti gerekiyorsa kural açıkça genişletilmelidir.

İnceleme listesi bir tanı kaydıdır. Her anonim müşteri satırına tek tek kimlik bulmayı veya ürün adlarını zorla değiştirmeyi gerektirmez. Birlikte ele alınabilecek nedenler gruplanır; belirsizlik korunur. Format bayrakları, inceleme bayrakları, işlem bağlamı ve alternatif senaryo işaretleri ayrı tanımlanır.
