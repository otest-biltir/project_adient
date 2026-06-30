# Sled Test Analyzer 🚀

Sled Test Analyzer, araç çarpışma testi (Sled) verilerinin analiz edilmesi ve grafik çıktılarının PNG resim olarak alınması için geliştirilmiş bir veri analiz uygulamasıdır. Bu sürüm yalnızca SPUL/grafik akışını içerir; genel bilgi formu ve Word şablon raporu kaldırılmıştır.

## Gereksinimler

Program çalışırken Python'ın ihtiyaç duyduğu kütüphaneler:
- `pandas` (Veri okuma)
- `numpy` (Matematiksel işlemler)
- `PyQt5` (Kullanıcı arayüzü)
- `matplotlib` (Grafikleri çizmek)
- `openpyxl` (Excel `.xlsx` dosyalarını işlemek)
- `xlrd` (Excel `.xls` dosyalarını işlemek)

## Nasıl Kullanılır? (Adım Adım Kılavuz)

### 1- Tek Excel Dosyasının Yüklenmesi (Sol Üst Köşe)
- **Excel Veri Dosyası Yükle:** Target ve actual verileri aynı Excel dosyasından okunur. Ayrı ayrı Actual/Target dosyası seçilmez.
- Uygulama ilk 2 satırı atlar ve 3. satırdan itibaren A:E sütunlarını veri olarak okur.
- Beklenen sütun sırası:
  - **A:** Time (`s`)
  - **B:** Target Acceleration (`g`)
  - **C:** Target hız (`m/s`)
  - **D:** Actual Acceleration (`g`)
  - **E:** Actual hız (`m/s`)

### 2- Grafikleri Görme
- Veriyi seçtikten sonra **"Oluştur / Güncelle"** butonuna basın. Uygulama verileri işler ve grafikleri ana ekranda çizer.
- **⬅** ve **➡** tuşlarıyla 3 grafik arasında geçiş yapabilirsiniz: `Spul`, `Acceleration vs Velocity`, `Actual vs Target Acceleration`.

### 3- Offset (Kayıklık) Ayarları
- Sağ üstteki **Actual Satır Offset Ayarları** tablosunda actual hız/ivme verisini satır bazlı yukarı/aşağı kaydırabilirsiniz.
- Her 1 satır, Excel verisindeki **0.0004 saniye** örnek aralığına karşılık gelir; uygulama seçilen satır offsetinin ms karşılığını tabloda gösterir.
- Pozitif satır offseti actual veriyi aşağı, negatif satır offseti yukarı kaydırır.
- Target veri offsetten etkilenmez; orijinal zaman ekseninde sabit kalır.
- **Evrensel Offset** alanı actual grafiklere aynı satır offsetini uygular.
- **14 ms Sabit Offset** seçeneği actual grafikleri **35 satır** kaydırır.
- Grafiklerin zaman ekseni en fazla **0.15 saniye** gösterir; bir seri daha erken biterse çizgi orada kesilir ve eksik bölüm 0 olarak tamamlanmaz.

### 4- PNG Resim Çıktısı Alma
- **Kayıt Dizini:** PNG dosyalarının kaydedileceği klasörü seçin.
- **Tüm Grafikleri Kaydet (.png):** Uygulama üç grafiği yüksek çözünürlüklü PNG dosyaları olarak seçilen klasöre kaydeder: `Spul.png`, `Acc_vs_Vel.png`, `Acc_vs_Targetacc.png`.

Bol analizler!
- *Created by Efe Nakcı* 🚀
