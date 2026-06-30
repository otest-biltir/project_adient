# Sled Test Analyzer (SPUL)

Bu sürümde uygulama yalnızca SPUL modülüyle çalışacak şekilde sadeleştirildi. Genel bilgi formu, Kapak, Photo Report ve Word template raporu kaldırıldı; tek Excel dosyasından çıktı olarak PNG grafik dosyaları üretilir. Excel formatı: 3. satırdan itibaren A=Time(s), B=Target Acceleration(g), C=Target hız(m/s), D=Actual Acceleration(g), E=Actual hız(m/s). Target verisi offsetten etkilenmez; actual hız/ivme offseti ms cinsinden girilir; hassasiyet 0.4 ms’dir ve her 0.4 ms 1 satıra karşılık gelir. Grafikler en fazla 0.15 saniyeye kadar çizilir.

## Kurulum

```bash
python -m pip install -r requirements.txt
```

## Çalıştırma

```bash
python app.py
```

> Not: Uygulama açılışında eksik paketleri otomatik kurmayı dener. Ağ/pip kısıtı varsa yukarıdaki kurulum komutunu manuel çalıştırın.
