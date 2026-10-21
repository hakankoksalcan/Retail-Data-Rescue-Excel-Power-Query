# Power Query kaynakları

## Excel dosyasında bulunan dört sorgu

`embedded/` klasörü, kullanıcının Excel 2021'de kaydettiği çalışma kitabındaki gerçek M kodlarının aynısını içerir:

- qSourceCheck: kaynak satır/sütun sayıları → SourceCheck sayfası.
- qRawSource: kaynak okuma ve filtre/sıralama öncesinde SourceRow → yalnızca bağlantı.
- qMissingValues: sekiz kaynak alanında null sayımı → qMissingValues sayfası.
- qTypedSource: kimlik alanlarının ham kopyaları ve veri tipi tanımları → yalnızca bağlantı.

`SourceRow=2`, ham dosyada ilk veri kaydının gerçek Excel satırıdır. Müşteri kimliği eksikleri doldurulmaz. Description içindeki sayısal değerler ham kaynakta korunur.

## Ayrı sağlanan tam raporlama modeli

Bu klasörün doğrudan altındaki 20 `.pq` dosyası çalışma kitabına henüz gömülü değildir. `fnRetailModel` → `qModel` → çıktı sorguları akışını tanımlar. M sözdizimi kontrol edilmiştir; tam modelin Excel'de çalıştırılması ve dashboard tablolarının canlı çıktılarla değiştirilmesi ayrı kabul adımıdır.

`qRefreshInfo`, qControls sonuçları değerlendirildikten sonra zaman bilgisi üretir; FAIL varsa hata verir. Tam model Facts/ReviewRecords çıktılarını varsayılan yalnızca bağlantı olarak tutar. Kurulum ve adresler: [Türkçe rehber](../docs/KURULUM_YENILEME_TR.md).
