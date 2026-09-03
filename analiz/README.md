# TÜPRAŞ (TUPRS) Yatırım Analizi — 02.09.2026

Bu klasör, borsa MCP konnektörü üzerinden toplanan verilerle hazırlanan TÜPRAŞ analizini içerir.

| Dosya | Açıklama |
|---|---|
| `TUPRAS_Analiz_2026-09-02.xlsx` | 10 sayfalık çalışma kitabı (canlı formüllerle) |
| `tupras-rapor.html` | Detaylı yatırım notu (Artifact olarak yayımlandı) |
| `build_tuprs.py` | Çalışma kitabını üreten script — veri güncellenirse yeniden çalıştırılabilir |

## Excel sayfaları

1. **1-Ozet** — Temel göstergeler, kısa/orta vade kanaati
2. **2-Bilanco** — 2026Q2 / 2026Q1 / 2025Q4 / 2025Q2 karşılaştırmalı bilanço + türetilmiş rasyolar
3. **3-Gelir-Tablosu** — 6 aylık kümülatif + yıllık (2023-2025) + 2Ç26 tek çeyrek, marj analizi
4. **4-Nakit-Akim** — İşletme/yatırım/finansman nakit akımları, serbest nakit akım verimi
5. **5-Rasyo-Sektor** — Değerleme rasyoları ve XUSIN emsal karşılaştırması
6. **6-Teknik** — Hareketli ortalamalar, RSI/MACD, pivot seviyeleri, fiyat performansı
7. **7-Senaryo** — FD/FAVÖK senaryo matrisi (sarı hücreler değiştirilebilir girdilerdir)
8. **8-Haberler** — KAP bildirimleri ve basın akışı, etki/vade etiketleriyle
9. **9-Analist** — Konsensüs, hedef fiyatlar, kazanç sürprizleri, şirket rehberliği
10. **10-Fiyat-Verisi** — Günlük (son 1 ay) ve aylık (son 1 yıl) OHLCV

## Ana sonuç

- **Kısa vade (0-3 ay): BEKLE.** Fiyat (392,25 TL) konsensüs hedefinin (396,98 TL) üzerine oturdu; Koç blok satışı ve 30.09 temettü kesintisi kısa vadede baskı unsuru.
- **Orta vade (6-18 ay): KADEMELİ AL.** 385-378 / 375-370 / 368-362 TL bandında üç kademe.
- **Ana risk:** Net rafineri marjının 13-15 $/varilden orta çevrim seviyesi olan 7-8 $/varile normalleşmesi — bu senaryoda adil değer ~317 TL.

## Yeniden üretim

```bash
pip install openpyxl
python build_tuprs.py
```

Veri tarihi: 02.09.2026. Yatırım danışmanlığı değildir.
