> Arşiv notu: Bu belge ilk veri denetiminin kaydıdır. Güncel Excel kurulum ve doğrulama durumu için [doğrulama kaydına](../validation/README.md) bakınız.

# Retail Data Rescue — Aşama 1: Veri uygunluk denetimi

**Sonuç:** UCI Online Retail, hedeflenen portfolyo için güçlü ve koşullu olarak uygun bir adaydır. Gerçek eksikler, tam tekrarlar, açıklama değişimleri, iptaller ve muhasebe/operasyon satırları vardır. Bu sorunları görünür kılmak için yapay hata eklemek gerekmiyor. Güçlü teslimat, bütün kayıtları “temiz satış” saymak yerine ölçülerin kapsamını açıkça tanımlayıp kaynakla uzlaştıran bir süreç olmalıdır.

Bu rapor yalnızca Aşama 1'i kapsar. Ham dosyada temizlik yapılmadı; nihai dönüşüm politikaları kararlaştırılmadı. Dashboard üretilmedi, mevcut Git deposu değiştirilmedi, GitHub'a gönderim yapılmadı. Masaüstü Excel çalıştırılmadı; sonuçlar Python ile özgün çalışma kitabından elde edildi. Power Query'nin Excel'de çalışması sonraki aşamada ayrıca doğrulanmalıdır.

## 1. Kaynak, lisans ve dosya yapısı

- Kaynak: [UCI Online Retail](https://archive.ics.uci.edu/dataset/352/online+retail).
- İndirilen özgün arşiv: `https://archive.ics.uci.edu/static/public/352/online+retail.zip`.
- Atıf: Chen, D. (2015). Online Retail [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5BW33.
- Kaynak sayfasındaki lisans: **Creative Commons Attribution 4.0 International (CC BY 4.0)**. Kaynağa atıf verilmeli ve yapılan dönüşümler belirtilmelidir. Lisans bağlantısı: https://creativecommons.org/licenses/by/4.0/legalcode.
- Arşivde tek dosya: `Online Retail.xlsx`, **23.715.344 bayt**.
- Tek çalışma sayfası: **Online Retail**.
- **541.909 veri satırı**, başlık dahil **541.910 çalışma sayfası satırı**, **8 sütun**; tamamen boş veri satırı **0**.
- Başlıklar sırasıyla: `InvoiceNo`, `StockCode`, `Description`, `Quantity`, `InvoiceDate`, `UnitPrice`, `CustomerID`, `Country`.
- Dosyadan hesaplanan tarih aralığı: **1 Aralık 2010 08:26 — 9 Aralık 2011 12:50**. Kaynakta saat dilimi belirtilmediğinden saatler Türkiye saatine çevrilmedi.
- En erken kayıt örneği: Excel satırı **2**, fatura `536365`, ürün `85123A`. En geç zaman örneği: satır **541896**, fatura `581587`, ürün `22631`.
- Fiyat birimi, UCI alan tanımında sterlin olarak belirtilmiştir; raporda **GBP** kullanılır. Dosyanın içinde ayrı para birimi sütunu yoktur.
- Ham XLSX SHA-256: `43465a06f2ccf7c8b5bd2892bc7defb52f97487934fe93b16ae4c3936424676d`.

Kaynak sayfası ve metadata yanıtı arşivlendi. Metadata'daki “eksik değer yok” ifadesi dosyayla çelişmektedir; aşağıdaki eksikler dosya içeriğinden doğrulanmıştır. Metadata'nın “6 özellik” sayısı iki kimlik alanını ayrı değerlendirmektedir; çalışma kitabının fiziksel sütun sayısı **8**'dir.

**Yeniden üretme:** `download_source.py` kaynak indirmesini ve dosya özetlerini kaydeder. `audit.py` çalışma sayfalarını, hücre türlerini, satırları ve tarih uçlarını okur. İkinci okuma `deep_audit.py` ile yapılır. İndirme ve denetim betikleri özgün XLSX'i kaydetmez veya değiştirmez. SHA-256 denetim öncesi/sonrası eşleşmiştir.

## 2. Bir satırın anlamı

Bir satır bir fatura/işlem içindeki ürün veya işlem kodu kalemidir; bir fatura, bir müşteri veya bir adet ürün değildir. `Quantity` bu kalemin miktarıdır. Bazı `StockCode` değerleri fiziksel ürün yerine kargo, indirim, komisyon veya manuel işlem açıklamaları taşır.

Örnek: `536365` faturası Excel satırları **2–8** arasında **7 farklı ürün kaleminden** oluşur. İlk iki kalem `85123A` ve `71053`'tür. Bunları aynı fatura numarasına sahip oldukları için mükerrer saymak yanlıştır.

| Ölçü | Dosyadan sonuç | Kontrol |
|---|---:|---|
| İşlem satırı | 541.909 | Veri satırlarını say; başlığı dışarıda bırak |
| Benzersiz InvoiceNo | 25.900 | `InvoiceNo.nunique()` |
| C önekli benzersiz InvoiceNo | 3.836 | `C` ile başlayan satırlarda farklı kimlik sayısı |
| C öneki olmayan benzersiz InvoiceNo | 22.064 | Diğer satırlarda farklı kimlik sayısı |
| Benzersiz StockCode | 4.070 | Boş olmayan `StockCode.nunique()`; hepsi fiziksel ürün değildir |
| Bilinen müşteri kimliği | 4.372 | Boşları dışarıda bırakarak `CustomerID.nunique()` |
| Country değeri | 38 | `Country.nunique()` |

**25.900 benzersiz kimlik, 25.900 tamamlanmış müşteri siparişi anlamına gelmez.** İptal ve düzeltme kimlikleri de sayılır. C olmayan 22.064 kimliğin içinde `A` önekli 3 muhasebe/düzeltme kaydı vardır. Ticari sipariş sayısı, iş kuralları belirlenince ayrıca tanımlanmalıdır.

## 3. Alan bazında eksikler

Eksik tanımı: özgün çalışma kitabında boş/null hücre. Yalnızca boşluklardan oluşan, boş olmayan metin hücreleri ayrıca kontrol edildi; bütün alanlarda bu ikinci ölçütün sonucu **0**.

| Alan | Eksik satır | Oran | Örnek / kontrol |
|---|---:|---:|---|
| InvoiceNo | 0 | %0 | Eşleşen eksik kayıt yok; `InvoiceNo.isna()` |
| StockCode | 0 | %0 | Eşleşen eksik kayıt yok; `StockCode.isna()` |
| Description | 1.454 | %0,2683 | Satır 624: `536414 / 22139`, açıklama boş; `Description.isna()` |
| Quantity | 0 | %0 | Eşleşen eksik kayıt yok; `Quantity.isna()` |
| InvoiceDate | 0 | %0 | Eşleşen eksik kayıt yok; `InvoiceDate.isna()` |
| UnitPrice | 0 | %0 | Eşleşen eksik kayıt yok; `UnitPrice.isna()` |
| CustomerID | 135.080 | %24,9267 | Satır 1445: `536544 / 21773`, miktar 1, fiyat 2,51, müşteri boş; `CustomerID.isna()` |
| Country | 0 | %0 | Eşleşen eksik kayıt yok; `Country.isna()` |

1.454 açıklama eksiğinin **tamamında** fiyat 0 ve müşteri kimliği boş. Bu, sorunların örtüştüğünü gösterir; eksik sayılarını toplamak hatalıdır. Yeniden üretme: `Description.isna() & (UnitPrice != 0)` sonucu **0**; açıklaması eksik satırlarda `CustomerID.isna().all()` sonucu **True**.

Müşterisi eksik ama tutarı pozitif olan kayıtlar vardır: satır **1445**, `Quantity × UnitPrice = 2,51 GBP`. Müşteri kimliği bulunmaması, bu tutarı otomatik olarak yanlış yapmaz. Ticari kapsamı uygun kayıt satış raporuna dahil edilebilir; kimlik gerektiren müşteri analizine dahil edilmemeli ve tek bir yapay müşteri kimliği altında birleştirilmemelidir.

## 4. Doğrulanmış kalite ve işlem bulguları

`SourceRow`, özgün Excel satır numarasıdır; başlık 1. satırdır. Aşağıdaki sayılar örtüşür. **Bunların toplamı “sorunlu satır sayısı” değildir.** Ayrıca bir bayrağın bulunması her zaman hata bulunduğu anlamına gelmez.

| Bulgu | Sayı | Özgün örnek | Yeniden üretilebilir kontrol | Ön değerlendirme |
|---|---:|---|---|---|
| Sekiz alanı tamamen aynı tekrarlar | 4.879 grup; 10.147 grup üyesi; ilk kayıt dışındaki 5.268 satır | Satırlar 487 ve 541: `536409 / 22111`, miktar 1, fiyat 4,95, aynı tarih/müşteri/ülke/açıklama | Sekiz özgün alanla `duplicated(keep=False)`; ek tekrar için `duplicated(keep='first')`; SourceRow ve hesaplanan tutarı anahtara katma | Tam eşitlik kanıtlandı; yanlışlıkla çoğaltıldıkları henüz kanıtlanmadı. Kaynakta satır kimliği yok; otomatik silme kararı verilmedi |
| Açıklama dışında diğer yedi alanı aynı kayıtlar | 2 grup, 4 satır | Satırlar 199126 ve 199146: `554084 / 23298`, miktar 3, fiyat 4,95; `BUNTING , SPOTTY ` ve `SPOTTY BUNTING` | Açıklama hariç yedi alana göre grupla; grup içindeki farklı açıklama sayısı >1 | Olası semantik tekrar; açıklama değişikliği veya iki gerçek kalem olma ihtimali ayrıştırılmalı |
| Aynı fatura–ürün çiftinin tekrar görünmesi | 9.694 grup, 20.378 satır | Satırlar 115 ve 127: `536381 / 71270`, miktarlar 1 ve 3 | `InvoiceNo, StockCode` grubunun boyutu >1 | Bu ölçüt tek başına mükerrer tespit etmez; miktar/fiyat/tarih farklı olabilir |
| C önekli iptal satırları | 9.288 satır, 3.836 farklı InvoiceNo | Satır 156: `C536383 / 35004C`, miktar -1, fiyat 4,65 | InvoiceNo metni `C` ile başlıyor mu? | UCI tanımı iptal diyor; fiziksel iade demiyor. Korunmalı ve ayrı sınıflanmalı |
| Negatif Quantity | 10.624 | Satır 156 yukarıdaki iptal; satır 2408 C olmayan kayıt | `Quantity < 0` | 9.288 C satırı + 1.336 C olmayan satır olarak ayrışır; her negatif hata değildir |
| C öneki olmayan negatif miktar | 1.336; tamamında fiyat 0 | Satır 2408: `536589 / 21777`, miktar -10, fiyat 0, açıklama/müşteri boş | `(Quantity < 0) & ~C`; bunun içinde `UnitPrice != 0` sonucu 0 | Stok/operasyon düzeltmesi adayı; işlem nedeni her kayıtta kesin değil |
| Sıfır fiyat | 2.515 | Satır 624: miktar 56, fiyat 0. Satır 9304: `537197 / 22841`, miktar 1, müşteri 12647, fiyat 0 | `UnitPrice == 0` | Fiyat doldurma yok; ücretsiz ürün, operasyon veya eksik fiyat ihtimalleri farklıdır |
| Sıfır fiyat, bilinen müşteri, pozitif miktar | 40 | Satır 9304 yukarıdaki kayıt | `(UnitPrice == 0) & (Quantity > 0) & CustomerID.notna()` | Ücretsiz gönderim/hediye veya fiyat sorunu olabilir; dosya tek başına açıklamıyor |
| Negatif fiyat | 2 | Satırlar 299985–299986: `A563186/A563187`, kod B, açıklama `Adjust bad debt`, miktar 1, fiyat -11.062,06 | `UnitPrice < 0` | Muhasebe/düzeltme işlemi güçlü aday; sıradan ürün fiyat yazım hatası diye düzeltme |
| Büyük mutlak satır tutarı | 17 | Satırlar 540423–540424: `581483/C581484`, ürün 23843, miktar ±80.995, fiyat 2,08, tutar ±168.469,60 | `abs(Quantity * UnitPrice) >= 10000` | 10.000 GBP bir inceleme eşiğidir; hata sınırı değildir. Bu çift aynı müşteri/ürünü içerir ve 12 dakika aralıklıdır; özgün sipariş bağlantısı kesinleştirilmiş değildir |
| Bir koda birden fazla açıklama | 650 kod; bu kodlara ait 113.465 satır | `85123A`: satır 2 `WHITE HANGING...`; satır 537623 `CREAM HANGING...`; ayrıca `?` ve karton notları | Boş olmayan **metin** açıklamalarla StockCode başına farklı açıklama sayısı >1 | Ürün adları, yeniden adlandırma ve operasyon notları karışıyor; tek açıklamayı körlemesine bütün satırlara yazma |
| Normalize açıklamada da çokluk | 643 kod | 23298: `SPOTTY BUNTING` ve `BUNTING , SPOTTY ` yalnızca boşluk/büyük harf farkı değil | Yalnız analiz kopyasında büyük harf ve tekrarlı boşlukları birleştir; farklı açıklama sayısını tekrar say | Hafif metin standardizasyonu bütün anlam farklarını çözmez |
| Açıklama başı/sonu boşluk | 113.452 satır | Satır 12: `22745`, `POPPY'S PLAYHOUSE BEDROOM ` sonunda boşluk | Metin Description için `value != value.strip()` | Görüntüleme/eşleştirme için düzeltilmesi uygun biçim sorunu; özgün açıklama korunmalı |
| Metin olmayan, dolu Description | 1 | Satır 420393: `572891 / 23343`, Description sayısal **20713**, miktar -400, fiyat 0 | Açıklama dolu ve Python hücre değeri `str` değil | Beklenen açıklama türüne aykırı hücre; başka ürün koduna not olabilir, ürün adı uydurulmamalı |
| İki ondalıktan daha hassas fiyat | 4 | Satır 157197: PADS, fiyat 0,001; satır 361743: BANK CHARGES, fiyat 0,001 | `Decimal(str(UnitPrice)).normalize()` üssü <-2 | Üç PADS ve bir BANK CHARGES satırı; kaynağı iki ondalığa yuvarlayıp gizleme |
| Aynı InvoiceNo içinde birden fazla zaman | 43 kimlik; bunlara ait 3.347 satır | Fatura 536591: satır 2423 16:57, satır 2457 16:58; tarih 01.12.2010 | InvoiceNo başına farklı InvoiceDate sayısı >1 | Tarih ayrıştırma hatası değildir; fatura düzeyinde zaman seçimi tanımlanmalı |

Sıfır fiyatların 1.179'unda miktar pozitif, 1.336'sında negatiftir; C önekli sıfır fiyat satırı **0**. Sıfır fiyatlı 2.515 satırın 2.475'inde müşteri kimliği eksik, 40'ında doludur. Bu alt gruplar kendi içinde ayrışıktır; diğer bulgularla örtüşebilir.

Tam tekrarların ilk kayıt dışındaki 5.268 satırının işaretli tutarı **21.740,98 GBP**'dir. Bu, olası bir silme kuralının kaynak toplamına etkisini ölçer; tutarın kesin yanlış olduğunu veya silme kararının onaylandığını göstermez. Hiçbir kayıt silinmedi.

Ek uç değer özeti: Quantity aralığı **-80.995 ile +80.995**; UnitPrice aralığı **-11.062,06 ile 38.970 GBP**; satır tutarı aralığı **-168.469,60 ile +168.469,60 GBP**. En yüksek UnitPrice örneği satır **222683**, `C556445 / M`, açıklama `Manual`, miktar -1 ve fiyat 38.970'dir. Bu da fiyat uçlarının otomatik olarak normal ürün fiyatı diye yorumlanamayacağını gösterir. Yeniden üretme: `Quantity.min()/max()`, `UnitPrice.min()/max()` ve `Quantity * UnitPrice` uçlarını hesapla; uç kayıtları SourceRow ile getir.

### Özel işlem kodları ve ürün kodu biçimi

`StockCode` için “5 rakam + isteğe bağlı harf” biçimine uymayan **33 farklı kod / 2.995 satır** vardır. Bu kontrol yalnızca aday üretir. Bütün biçim dışı kodları hizmet işlemi saymak doğru değildir.

| Kod | Satır | Dosyada görülen açıklama | Örnek Excel satırı |
|---|---:|---|---:|
| POST | 1.256 | POSTAGE | 47 |
| D | 77 | Discount | 143 |
| C2 | 144 | CARRIAGE | 1425 |
| DOT | 710 | DOTCOM POSTAGE | 1816 |
| M | 571 | Manual | 2241 |
| m | 1 | Manual | 40385 |
| BANK CHARGES | 37 | Bank Charges | 361743 |
| S | 63 | SAMPLES | 14438 |
| AMAZONFEE | 34 | AMAZON FEE | 15018 |
| B | 3 | Adjust bad debt | 299984–299986 |
| CRUK | 16 | CRUK Commission | 317510 |

Yukarıdaki açık liste **2.912 satır** kapsar; açıklamalar hizmet, ücret, indirim, numune veya düzeltme ayrımı gerektirir. Satış kapsamı ayrıca kararlaştırılmalıdır. `M` ile `m` otomatik birleştirilmedi.

`gift_` ile başlayan hediye çeki kodlarında **34 satır** vardır; örnek satır **38250**, `gift_0001_40`, `Dotcomgiftshop Gift Voucher £40.00`. Hediye çeki muhasebesi bu dosyadan kesinleştirilemez.

`DCGS` ile başlayan kodlarda **45 satır** bulunur; örnek satır **21328**, `DCGS0076`, `SUNJAR LED NIGHT NIGHT LIGHT`. Bazıları açıkça ürün açıklaması taşır. Bu nedenle biçim dışı kod ≠ ürün olmayan kod.

`PADS` kodunda **4 satır** vardır; açıklaması `PADS TO MATCH ALL CUSHIONS`. Kod biçimi sıra dışı olsa da ürün olma ihtimali kuvvetlidir; üç kayıtta 0,001, birinde 0 fiyat görülür. İstisna tablosu için adaydır.

## 5. Tarih ve hücre tipi sorunları gerçekten var mı?

**Bozuk tarih veya karışık metin tarih bulunduğu söylenemez.** 541.909 InvoiceDate hücresinin tamamı Excel'den gerçek datetime olarak okunmuştur. Eksik/ayrıştırılamayan tarih **0**. `01/12` gibi bir görüntüyü Türkiye/ABD tarih tahminiyle yeniden yorumlamak gereksiz ve risklidir.

| Alan | Gerçek hücre değer türü | Yorum |
|---|---|---|
| InvoiceNo | 532.618 sayı, 9.291 metin | Metinler 9.288 C ve 3 A önekli kayıttır. Kimlik alanını metin olarak modellemek uygundur; hepsini tam sayıya dönüştürme |
| StockCode | 487.036 sayı, 54.873 metin | Harfli ürün ve işlem kodları geçerlidir. Örnekler: 71053 (satır 3), 85123A (satır 2), POST (satır 47) |
| Description | 540.454 metin, 1.454 boş, 1 sayı | Sayısal istisna satır 420393'tür |
| Quantity | 541.909 tam sayı | Sayısal olmayan 0; kesirli 0; sıfır miktar 0 |
| InvoiceDate | 541.909 datetime | Metin tarih 0; eksik/geçersiz 0 |
| UnitPrice | 541.909 sayısal hücre | Python'da int/float farkı iki farklı iş türü değildir; sayısal olmayan 0 |
| CustomerID | 406.829 tam sayı, 135.080 boş | Kimliktir; miktar gibi toplama yapılmamalıdır |
| Country | 541.909 metin | Eksik 0 |

InvoiceNo/StockCode/CustomerID alanlarında başta veya sonda metin boşluğu **0**. C önekli olup miktarı sıfır/pozitif olan satır **0**, küçük `c` önekli satır **0**. Aynı InvoiceNo içinde birden fazla bilinen müşteri kimliği **0**, birden fazla ülke **0**, hem boş hem dolu müşteri kimliği olan fatura **0**. Bu sonuçlar bütün ticari işlemlerin doğru olduğunu kanıtlamaz; ilgili kontrol kapsamındaki tutarlılığı gösterir.

## 6. Ölçüler ve kaynakla uzlaştırma

**Şu anda güvenle hesaplanabilenler:** kaynak satırları; farklı InvoiceNo/StockCode; bilinen müşteri kimliği sayısı; eksik oranları; işaretli miktar/fiyat çarpımı; C önekli kayıtların ayrı tutarı; kayıtların ay/ülke bazında dağılımı. Bunlar kaynakta görünenlerin teknik ölçüleridir.

| Ham tutar bölümü | GBP, kaynak hassasiyeti korunarak |
|---|---:|
| Bütün satırlarda Σ(Quantity × UnitPrice) | 9.747.747,934 |
| C önekli satırlar | -896.812,49 |
| C öneki olmayan satırlar | 10.644.560,424 |
| Müşteri kimliği bilinen satırlar | 8.300.065,814 |
| Müşteri kimliği eksik satırlar | 1.447.682,12 |

Kontroller: `C + C olmayan = bütün satırlar` ve `bilinen müşteri + eksik müşteri = bütün satırlar`. Hesaplamalar `Decimal(str(Quantity)) × Decimal(str(UnitPrice))` ile yapılır. İki okumanın satır, eksik, tekrar ve tutar sonuçları eşleşmiştir; `verification.json` içindeki **10 kontrol geçti**.

**9.747.747,934 GBP, bu aşamada “net satış geliri” değildir.** Ham işaretli işlem tutarıdır; özel işlem kodları, tekrarlar ve düzeltmeler dahil edilmiştir. C olmayan tutar da tek başına satış toplamı sayılmaz. Kaynakta 0,001 fiyatlar bulunduğu için üçüncü ondalık korunmuştur; gösterim yuvarlaması ayrı bir karar olacaktır.

Şimdilik şu iddialar desteklenmez:

- Kesin ürün satış geliri: kargo, indirim, hediye çeki, komisyon ve muhasebe kayıtlarının kapsamı çözülmeli.
- Fiziksel iade miktarı/oranı: C, iptal işaretidir; fiziksel iade teyidi değildir.
- Tamamlanmış müşteri siparişi ve ortalama sipariş değeri: ticari işlem kapsamı ve payda tanımlanmalı.
- Bütün müşteriler için müşteri analizi: satırların %24,93'ünde CustomerID eksik; yalnız tanınan müşterilerin kapsamı raporlanabilir.
- Kâr, marj veya KDV sonrası gelir: maliyet, vergi ve işletme kuralları yok.
- İptalin özgün faturaya kesin bağlantısı: doğrudan bağlantı alanı yok; miktar/tutar benzerliği eşleştirme adayıdır.
- Tam yıl büyüme karşılaştırması: 2010 yalnız Aralık'ı, 2011 Aralık ise 9'una kadar kayıtları kapsıyor.

## 7. İlk karar çerçevesi — henüz temizlik uygulanmadı

| Durum | Hata / geçerli işlem / belirsizlik | Önerilen yaklaşım |
|---|---|---|
| Açıklama çevresindeki boşluklar | Biçim sorunu | Özgün Description yanında standartlaştırılmış bir alan; kaynak değişmez |
| Kimliklerin sayı/metin karışımı | Kimliklerin doğal kod çeşitliliği | InvoiceNo/StockCode'u metin olarak modelle; A/C ve harfli ürün kodlarını koru |
| Eksik CustomerID | Müşteri analizinde eksik bilgi; satış tutarında otomatik hata değil | Uygun işlem kapsamındaki satış satırını koru; müşteri analizinde hariç tut ve kapsama oranını göster; kimlik uydurma |
| C önekli negatif miktar | Kaynak tanımına göre iptal | Ayrı sınıf ve işaretli tutar; fiziksel iade iddiası yok |
| C olmayan, sıfır fiyatlı negatif miktar | Operasyon/stok düzeltmesi adayı | İnceleme; ürün satışı/adet göstergelerinden ayrılma politikası tanımlanmalı |
| B / Adjust bad debt ve A kimlikleri | Muhasebe/düzeltme güçlü adayı | Satışla karıştırma; ayrı kategoriyle kaynak tutar uzlaştırmasına dahil et |
| Tam tekrarlar | Eşit kayıtlar kesin; kazara çoğaltma belirsiz | Önce bayrakla; silme kararı ancak belgelenmiş varsayım ve satır/tutar etkisiyle |
| Sıfır fiyat, bilinen müşteri | Geçerli hediye veya fiyat eksikliği olabilir | Müşteriye açıklanacak inceleme sınıfı; ortalama fiyatla doldurma yok |
| Kod–açıklama farkları | Bazısı biçim farkı, bazısı işlem notu/yeniden adlandırma | Kontrollü ürün sözlüğü; notları koru; her kod için tek etiketi otomatik dayatma |
| Büyük miktar/tutar | Büyük ticari işlem, iptal/düzeltme veya veri hatası olabilir | İnceleme eşiğiyle görünür kıl; otomatik silme yok |

## 8. Portfolyo uygunluğu, senaryo ve sonraki aşamalar

**Önerim: bu veriyi ana aday olarak kabul edelim.** Veri temizliği ve ölçüm kararlarını gösterecek kadar gerçek sorun barındırıyor; 8 alanlı yapısı öğrenme sürecini yönetilebilir tutuyor. Nihai uygunluk kararını birlikte vereceğiz. Verinin gerçek kapsamı bir Birleşik Krallık çevrimiçi perakendecisidir; Türkiye işletmesi diye sunulmamalıdır.

Önerilen müşteri senaryosu: “Birleşik Krallık merkezli çevrimiçi hediye perakendecisinin işlem dışa aktarımında satış, iptal ve operasyon satırları karışıyor. İşletme sahibi, her ay yenilenebilen ve hangi kayıtların neden kullanıldığını açıklayan bir rapor istiyor.” Bu bir **portfolyo müşteri senaryosudur**, gerçek bir freelance müşteriye teslim edildiği iddia edilmez.

| Aşama | Amaç | İşlem | Doğrulama |
|---|---|---|---|
| 1 — Bu denetim | Veriyi tanımak ve uygunluk kararı | Ham dosyayı incele, sayı/örnek/kontrol üret | SHA-256, satır sayısı, iki okuma ve tutar bölümleri |
| 2 — Karar sözleşmesi | Müşteriye açıklanabilen kurallar | Satış/iptal/operasyon kapsamı, tekrar ve ürün sözlüğü karar günlüğü | Örnek kayıtlar üzerinde beklenen sınıf ve rapora etkisi |
| 3 — Excel Power Query | Yenilenebilir dönüşüm | Kaynak bağlantısı, SourceRow, kontrollü tipler ve kalite bayrakları | Gerçek Excel'de 541.909 kaynak satırı, hata kayıtları ve Python sayılarıyla karşılaştırma |
| 4 — İnceleme ve uzlaştırma | Kayıt kaybını görünür kılmak | İnceleme tablosu, birbirini dışlayan sonuç sınıfları, kaynak/çıktı kontrolleri | Satır sayıları ve işaretli tutarlar kaynakla uzlaşır; dışlamalar açıklanır |
| 5 — Dashboard ve yenileme | Güncellenebilir rapor teslimi | Tanımlı göstergeler, kaynak dosyası yenileme ve önce/sonra kanıtı | Aynı kurallar yeni kaynakta çalışır; kontroller tekrar geçer |

Kaynağın 541.909 satırlık boyutu Excel çalışma sayfası sınırının altındadır. Power Query performansı için Excel sürümü, 32/64 bit bilgisi ve bilgisayar belleği ayrıca değerlendirilmelidir; yerel performans bu bulut denetimiyle doğrulanmış değildir.

## 9. Kullanıcının ilk küçük adımı

**Amaç:** ham kaynağı koruyup fatura ile kalem ayrımını kendin görmek.

1. Özgün `Online Retail.xlsx` dosyasını yerel `raw` klasörüne koy; temizlik veya dosyanın üzerine kaydetme yapma.
2. Dosyayı salt okunur aç ve `Online Retail` sayfasındaki **A1:H8** aralığına bak.
3. Sekiz başlığı ve `536365` faturasına ait satırlar **2–8** arasındaki yedi ürün kalemini kontrol et. İlk satırın kodu `85123A`, miktarı 6, birim fiyatı 2,55 ve müşteri kimliği 17850 olmalı.
4. Öğrenme notuna şu cümleyi yaz: “Bu kaynakta bir satır bir fatura kalemidir; aynı InvoiceNo birden fazla üründe tekrar eder.”

**Doğrulama:** tek InvoiceNo (`536365`), yedi veri satırı ve yedi farklı StockCode görmelisin. Mükerrer silme veya tutar düzeltme yapma. Tarihin Excel'de görünümü yerel ayara göre değişebilir; bu tek başına bozuk tarih anlamına gelmez.

Excel sürümü, işletim sistemi ve Veri → Veri Al erişimi yanıtı henüz alınmadı. Bu ilk adım Power Query gerektirmez. Veri uygunluğu birlikte değerlendirildikten sonra sürümüne uygun, küçük Power Query adımlarıyla ilerlenebilir.

## 10. Yeniden üretme ve kanıt dosyaları

Denetim klasöründe, Python 3.12 / pandas 2.2.3 / openpyxl 3.1.5 ile:

```bash
python3 download_source.py
python3 audit.py 'raw/Online Retail.xlsx'
python3 deep_audit.py
python3 verify_results.py
```

- `download-manifest.json`: kaynak URL'leri, boyutlar, SHA-256 özetleri.
- `audit-results.json`: tüm temel sayılar, sıfır sonuçlar, alan tipleri, dağılımlar ve örnek kayıtlar.
- `deep-audit-results.json`: ek alan/işlem kontrolleri, tam tekrar grupları, açıklama eşleştirmeleri ve örnekler.
- `findings.csv`: temel kontrol adları, sayıları ve koşulları.
- `examples.csv`: temel bulguların sekiz özgün alanı ve SourceRow numarasıyla örnekleri.
- `verification.json`: iki okuma/hesaplama ve kaynak koruma kontrolleri.

Ek bulguların kayıt örnekleri `deep-audit-results.json` içindedir. Her sıfır sonuç için eşleşen örnek olmadığı özellikle korunmuştur. CSV/JSON dosyaları bir temiz veri seti veya nihai müşteri teslimi değildir; Aşama 1'in denetim kanıtlarıdır.
