# Sled Test Analyzer (SPUL)

Bu sürümde uygulama yalnızca SPUL modülüyle çalışacak şekilde sadeleştirildi. Genel bilgi formu, Kapak, Photo Report ve Word template raporu kaldırıldı; tek Excel dosyasından çıktı olarak PNG grafik dosyaları üretilir. Excel formatı: 3. satırdan itibaren A=Time(s), B=Target Acceleration(g), C=Target hız(m/s), D=Actual Acceleration(g), E=Actual hız(m/s).

## Kurulum

```bash
python -m pip install -r requirements.txt
```

## Çalıştırma

```bash
python app.py
```

> Not: Uygulama açılışında eksik paketleri otomatik kurmayı dener. Ağ/pip kısıtı varsa yukarıdaki kurulum komutunu manuel çalıştırın.
