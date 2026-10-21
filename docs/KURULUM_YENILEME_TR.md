# Excel kurulumu, yenileme ve kullanım — Türkçe

## Bu sürümün durumu

Ana dosya `Retail_Data_Rescue.xlsx`, Windows 11 / Excel 2021'de oluşturulmuş dört gömülü Power Query sorgusu içerir: qSourceCheck, qRawSource, qMissingValues ve qTypedSource. Kaynak sayımları ve eksik değer sonuçları gerçek Excel'de yenilenip doğrulandı; veri tipi sorgusunun önizlemesi kullanıcı tarafından onaylandı. Tam veri tipi hata taraması henüz ayrı bir kabul kanıtına sahip değildir.

Dashboard, Python referansıyla üretilmiş uzlaştırılmış tablolara bağlı Excel formülleri ve grafiklerini kullanır. Ülke/ay filtreli altı gösterge Excel 2021'de kontrol edildi. Gömülü denetim sorgularını yenilemek, dashboard'un referans tablolarını yeniden üretmez.

Bu rehberin 2–5. bölümleri, **ayrı sağlanan tam M modelinin** entegrasyonu içindir; mevcut dört sorgunun kurulumunu tekrar yapman gerekmez. Tam modelin uçtan uca Excel yenileme kabulü henüz tamamlanmadı. Ana kitabı koruyup ayrı kopyada çalış. Excel 2021 masaüstü kullanıldı; büyük model için 64 bit sürüm önerilir.

### Mevcut sürümü açmak ve denetimi yenilemek

1. Ana Excel dosyasını aç; Dashboard filtrelerini kullanabilirsin.
2. `raw/Online Retail.xlsx` dosyasını yerel bilgisayarında tut; Settings A6'ya tam yolunu yaz.
3. Veri → Tümünü Yenile. SourceCheck: 541909 veri satırı / 8 kaynak sütunu. qMissingValues: CustomerID 135080, Description 1454, diğer altı alan 0.
4. Ayarlar tablosunun adı tblSource olarak kalsın. Python referansıyla oluşturulmuş Controls sayfasını gerçek M yürütme sonucu olarak yorumlama.

## 1. Kaynak parametresi

**Amaç:** yolu kodda tekrar tekrar değiştirmemek.

1. GitHub'dan indirdiğin proje ZIP'ini tek klasöre çıkar. Ham dosyayı düzenleme.
2. Çalışma kitabını aç; Settings sayfasındaki **A6** hücresini kendi tam yerel XLSX yolunla değiştir. Örnek: C:\RetailDataRescue\raw\Online Retail.xlsx.
3. **B6** gerçek Excel TRUE/DOĞRU değeri olarak kalsın. Bu değer orijinal UCI kaynağının referans sayılarını denetler.

**Kontrol:** tablo adı tblSource; başlıklar SourcePath ve ValidateBaseline. Kaynak sayfa adı Online Retail. Kaynak yolunu değiştirmek tek başına statik referans tablolarını yenilemez.

Windows'ta ham dosya bütünlüğü için PowerShell'de `Get-FileHash -Algorithm SHA256 ".\raw\Online Retail.xlsx"` çalıştır. Beklenen SHA-256: **43465a06f2ccf7c8b5bd2892bc7defb52f97487934fe93b16ae4c3936424676d**. Bu harici dosya bütünlüğü kontrolüdür; M sorgusunun hash hesapladığı iddia edilmez.

## 2. Fonksiyon ve model sorgusu

**Amaç:** dönüşümlerin tek kural setiyle başlaması.

Veri → Veri Al → Diğer Kaynaklardan → Boş Sorgu. Gelişmiş Düzenleyici'yi aç ve içeriği tamamen değiştir:

- power_query/fnRetailModel.pq içeriğini yapıştır; sorgu adını **fnRetailModel** yap. Fonksiyon için Invoke çağrısı yapma.
- İkinci boş sorguya qModel.pq içeriğini yapıştır; adını **qModel** yap.
- qModel bir record döndürür. Bunu ve fonksiyonu **yalnızca bağlantı** olarak tut.

Gizlilik düzeyi sorulursa kaynakların gerçek sınıfına uygun seçim yap. Ham UCI dosyası herkese açık veridir; başka müşteri verisi için aynı varsayımı yapma. Gizlilik kontrollerini kapatmak kuruluma dahil değildir.

**Kontrol:** qModel önizlemesinde Facts, KPI, Reconciliation vb. alanlar görünür. Hata varsa önce kaynağın yolu, sayfa adı, sekiz başlık ve native tarih/sayı sözleşmesini kontrol et; hatalı satırları silerek geçme.

## 3. Sonuç sorgularını çalışma sayfalarına bağla

Her .pq dosyası için bir boş sorgu aç, içeriğini yapıştır ve dosya uzantısı hariç aynı sorgu adını ver. İlk olarak qControls'u çalıştır; bütün satırlar PASS olmalı. Aşağıdaki çıktıların yükleme adresi ilgili sayfanın **A1** hücresidir:

| Sorgu | Sayfa | Amaç |
|---|---|---|
| qControls | Controls | Gerçek Power Query kontrol sonucu |
| qKPI | KPI | Genel göstergeler |
| qDashboardData | DashboardData | Filtreli kartların ülke/ay küpü |
| qMonthly | Monthly | Aylık tablolar ve grafiğin ay listesi |
| qCountry | Country | Kaynak genelinde ülke sıralaması |
| qProduct | Product | Kaynak genelinde ürün kodu sıralaması |
| qReconciliation | Reconciliation | Birbirini dışlayan sınıflar |
| qBridge | Bridge | Ham tutardan ürün dengesine köprü |
| qDuplicateScenario | DuplicateScenario | Ana ve alternatif sonuç |
| qQuality | Quality | Örtüşebilen kalite bayrakları |
| qBeforeAfter | BeforeAfter | Aynı kaynak satırlarında önce/sonra |
| qProductDictionary | ProductDictionary | Açıklama varyantları |
| qSpecialOperations | SpecialOperations | Özel işlem dökümü |
| qCustomers | Customers | Yalnız uygun, kimliği bilinen satışlar |
| qReviewExamples | ReviewExamples | İlk 200 inceleme kaydı; referans örnek seçimiyle aynı olmak zorunda değil |
| qRefreshInfo | RefreshInfo | Kontroller değerlendirildikten sonra UTC yenileme bilgisi |

Mevcut sayfalarda statik referans tabloları vardır. Bunları canlı sonuçlarla değiştirirken **sayfanın kendisini silme**: tablonun içine tıkla → Tablo Tasarımı → Aralığa Dönüştür → mevcut içerikleri temizle. Ardından Kapat ve Yükle → Şuraya Yükle → Tablo → Var olan çalışma sayfası → doğru sayfanın A1 hücresi. qRefreshInfo'yu en son yükle. Aynı sayfada eski veri aşağıda kalmasın; qCountry ve qProduct başlıklarının ilk sütunda olması sıralama grafiklerinin bağlantısını korur.

Dashboard formülleri sayfa/sütun konumlarına bağlıdır; çıktı sütunlarını elle yeniden sıralama. Power Query'nin verdiği tablo adına bağlı olmadıkları için eski tablo adını korumak gerekmez. Settings/tblSource tablosunu dönüştürme veya silme.

qFacts ve qReviewRecords büyük çıktılardır: varsayılan **yalnızca bağlantı**. Bütün kaynağı çalışma sayfasına çoğaltmak gerekmez. Tam inceleme kaydı ayrıca review/ReviewRecords.csv.gz içindedir. CSV.GZ dosyaları sıkıştırılmış CSV'dir; örneğin 7-Zip ile açılabilir. Statik CSV'ler gerçek Power Query yenilemesiyle kendiliğinden yeniden yazılmaz; onları Python yeniden üretir.

**Kontrol:** qKPI ve Reconciliation sonuçlarını reference/ içindeki referanslarla eşleştir. qControls kapsamı Excel Power Query olmalı. Sadece fonksiyon önizlemesi veya PASS yazan statik referans tablo gerçek Excel çalışmasını kanıtlamaz.

## 4. Dashboard'u kullan

- B5: ülke; F5/J5: başlangıç/bitiş ayı. Başlangıç bitişten sonra olmasın.
- Kartlar ve aylık grafik bu filtrelerden etkilenir. Ülke/ürün sıralamaları açıkça **full source** kapsamındadır.
- Tutarlar görünümde iki ondalıktır; kontrol tabloları dört ondalık gösterir. Kaynak değerleri yuvarlanmaz.
- Merchandise balance, muhasebe net geliri veya kâr değildir. Cancellations, fiziksel iade değildir.
- Review rows, satıştan dışlanan satır sayısı değildir.
- Tarih seçiminde 2010 Aralık gözlenen kayıtları 1–23; 2011 Aralık 1–9 aralığındadır. İlk dönem sonunda işlem olmaması dosyadan kesin veri kaybı veya kapalı işletme olarak açıklanamaz.
- Listeler/ay grafiği 1000 ay, ülke listesi 300 satır için hazırdır. Bu sınırlar aşılırsa Lists/ChartData formüllerini genişlet; grafik adı aralıkları ay sayısına göre otomatik büyür.

## 5. Gerçek Excel kabul ve yenileme testi

1. Orijinal kaynak, ValidateBaseline=TRUE: Veri → Tümünü Yenile. Bütün sorgular bitene kadar bekle; hata/yenileme başarısızlığı olmadığını kontrol et. validation/EXCEL_ACCEPTANCE_CHECKLIST.csv içindeki beklenen sonuçları girip durumları doldur.
2. Aynı kaynakla tekrar Tümünü Yenile: satır/sınıf/tutar/gösterge değerleri değişmemeli; yalnız yenileme zamanı değişebilir.
3. Çalışma kitabının ayrı bir test kopyasında Settings A6'yı **tests/TEST_ONLY_Online_Retail_plus_one.xlsx** dosyasına yönlendir; B6=FALSE. Bu dosya açıkça testtir ve özgün kaynak değildir.
4. Tümünü Yenile: kaynak satırı **541910**; satış ve ham tutar **+7 GBP**; uygun satış faturası **19774**; iptal tutarı aynı; ek tam tekrar sayısı **5268**. Eklenen satırın SourceRow değeri **541911**, InvoiceNo değeri **990001** olmalı.
5. Asıl çalışma kitabında orijinal yolu ve TRUE değerini koru. Asıl teslimi test verisiyle bırakma.
6. Filtre kontrolü: B5=United Kingdom, tüm dönem → satış kartı **8747561.663 GBP**, iptal kartı **445843.38 GBP**; ardından All'a dön.

Sonuçlar eşleşmezse Excel sürümünü, sorgu hata metnini ve fark veren kontrol satırını kaydet. Gerçek Excel onayı ancak bu testler yapıldıktan sonra tamamlanır. Kullanım sırasında yeni gerçek kaynak için B6=FALSE; genel uzlaştırma kontrolleri yine çalışır. İş kuralına uymayan yeni kodlar görünür inceleme sınıfına gider.

## 6. Python referansını yeniden üret

Python 3.12 ve requirements-validation.txt içindeki paketlerle, proje klasöründe:

```bash
python3 scripts/build_reference.py
python3 scripts/test_refresh.py
python3 scripts/build_reference_workbook.py
python3 scripts/make_manifest.py
python3 scripts/validate_publication.py
```

M sözdizimi için proje klasöründe Node 22+ ile `npm ci --ignore-scripts` ve `npm run validate:m` çalıştır. Sözdizimi kontrolü çalıştırma/yenileme doğrulaması değildir. pandas 2.2.3 ve openpyxl 3.1.5 referans denetim için kullanıldı. Çalıştırma süreleri Excel performansını göstermez.

Referans çalışma kitabı üretimi ana Excel dosyasını değiştirmez; ayrı `reference/Reference_Workbook.xlsx` üretir. TEST_ONLY XLSX dosyası depoya eklenmez, `scripts/test_refresh.py` tarafından oluşturulur. Tam sınıflanmış CSV.GZ çıktısı da yerel yeniden üretimle elde edilir.
