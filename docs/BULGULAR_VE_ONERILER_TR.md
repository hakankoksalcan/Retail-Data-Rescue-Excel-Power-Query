# Veriden doğrulanan bulgular ve sınırlı öneriler

1. **UK ürün satış kapsamının %85,16'sı**: 8747561.663 / 10271433.063 GBP. Ülke kapsamı Country alanının anlamı ile sınırlıdır; Türkiye pazarı sonucu çıkarılmaz. Öneri: ülke raporunda UK ve diğer ülkeleri ayrı göster; veri dışı pazar potansiyeli iddiası kurma.
2. **23843 kodunun dengesi sıfır**: satış ve iptal tutarı ayrı ayrı 168469.60 GBP; bu ürünün iptali ürün iptal tutarının %35,19'u. Satırlar 540423/540424 aynı ürün/müşteriyi ve karşıt miktarları gösterir; doğrudan orijinal fatura bağlantısı alanı yoktur. Öneri: yalnız satış sıralamasıyla başarılı ürün demeden iptal etkisini göster.
3. **Satış müşteri kapsamı satırda %75,10, tutarda %85,30**. Kaynağın bütününde eksik CustomerID oranı %24,93. Öneri: satış raporunu anonim kayıtlardan arındırma; müşteri analizi kapsamını ayrı sun. Anonim kimlik üretme.
4. **Tam tekrar alternatifinin etkisi**: satış -24213.74 GBP, iptal büyüklüğü -2823.02 GBP, ürün dengesi -21390.72 GBP; uygun fatura sayısı aynı. Öneri: kaynak satır kimliği olmadığı için temel sonucu koru; tekilleştirme talebini kanıt ve etkisiyle tartış.
5. **Dönem sınırı karşılaştırmayı değiştirir**: 2010 Aralık gözlenen tarihler 1–23, 2011 Aralık 1–9. Öneri: Aralıkları tam ay/yıllık büyüme diye kıyaslama. Neden işlem olmayan tarihler bulunduğu dosyadan kesinleştirilemez.
