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

### 1- Verilerin Yüklenmesi (Sol Üst Köşe)
- **Actual Data Yükle:** Sisteminize dışarıdan aldığınız gerçekleşen test verisini uygulamaya tanıtın. Excel dosyasında `Time`, `Velocity` ve `Acceleration` verilerinin bulunması gerekir.
- **Target Data Yükle:** Eğer varsa, hedeflenen çarpışma verilerini bu butonla uygulamaya tanıtın. İçerisinde `Target Velocity` ve `Target Acceleration` gibi değerler bulunmalıdır. Girilmesi zorunlu değildir.

### 2- Grafikleri Görme
- Veriyi seçtikten sonra **"Oluştur / Güncelle"** butonuna basın. Uygulama verileri işler ve grafikleri ana ekranda çizer.
- **⬅** ve **➡** tuşlarıyla 3 grafik arasında geçiş yapabilirsiniz: `Spul`, `Acceleration vs Velocity`, `Actual vs Target Acceleration`.

### 3- Offset (Kayıklık) Ayarları
- Sağ üstteki **Grafik Offset Ayarları** tablosunda her grafik için milisaniye (`ms`) cinsinden offset verebilirsiniz.
- **Evrensel Offset** alanı tüm grafiklere aynı offset değerini uygular.
- **14 ms Sabit Offset** seçeneği tüm grafikleri 14 ms değerine sabitler.

### 4- PNG Resim Çıktısı Alma
- **Kayıt Dizini:** PNG dosyalarının kaydedileceği klasörü seçin.
- **Tüm Grafikleri Kaydet (.png):** Uygulama üç grafiği yüksek çözünürlüklü PNG dosyaları olarak seçilen klasöre kaydeder: `Spul.png`, `Acc_vs_Vel.png`, `Acc_vs_Targetacc.png`.

Bol analizler!
- *Created by Efe Nakcı* 🚀
