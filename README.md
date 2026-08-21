# Komite Nabız

İki komite sisteminin arka uçlarını **12:00–22:00 (Türkiye saati)** arasında uyanık
tutar ve aynı anda ayakta olup olmadıklarını denetler.

## Neden gerekli

Arka uçlar Render'ın ücretsiz katmanında çalışıyor. Render **15 dakika hareketsizlikten
sonra servisi durduruyor**; sonraki ilk istek servisi sıfırdan ayağa kaldırıyor ve
**30 saniyenin üzerinde** sürebiliyor. Bu bekleme, kullanıcı tarafında zaman zaman
hatalı davranışa yol açıyordu.

Bu depo, çalışma saatleri boyunca 10 dakikada bir iki arka ucu da yokluyor. Böylece
servisler uyumuyor ve kimse 30 saniye beklemiyor.

## Neden 10 saat, daha fazlası değil

Render ücretsiz katmanı **ayda 750 örnek-saati** veriyor ve bu kota servis başına
değil, **hesabın tamamı için**. Kota bitince Render bütün ücretsiz servisleri ay
sonuna kadar askıya alıyor.

| Pencere | 2 servis × 30 gün | Kalan marj |
|---|---|---|
| **10 sa/gün** | 600 saat | 150 saat ✓ |
| 12 sa/gün | 720 saat | 30 saat — çok dar |
| 18 sa/gün | 1080 saat | **kota aşımı** |

Marj önemli: pencere dışında gerçek kullanıcı girdiğinde servis yine uyanıyor ve o
süre de kotadan düşüyor.

## Uyarı nasıl gelir

Bir arka uç yanıt vermezse iş **başarısız olur** ve GitHub bunu e-posta ile bildirir.
Yani bu depo hem uyanık tutucu hem de kesinti alarmıdır.

## Neden ayrı ve public bir depo

Public depolarda GitHub Actions dakikası sınırsız; özel depolarda aylık 2000 dakika
sınırı var ve bu iş tek başına onun %90'ını yerdi. Depoda **hiçbir sır yok** —
yalnızca iki API adresi, onlar da zaten sitelerin tarayıcıya gönderdiği JavaScript
paketinin içinde açıkta duruyor.

## Elle çalıştırma

Actions sekmesi → **Nabız** → **Run workflow**.

## Saat penceresini değiştirmek

`.github/workflows/nabiz.yml` içindeki `cron` satırı. GitHub UTC kullanır,
Türkiye UTC+3'tür: `9-18` UTC = 12:00–21:59 Türkiye saati.
Pencereyi genişletmeden önce yukarıdaki kota tablosuna bak.
