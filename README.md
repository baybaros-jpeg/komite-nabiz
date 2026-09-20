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

## Neden GitHub'ın saatine güvenmiyoruz (2026-09-20)

İlk sürüm `*/10 9-18 * * *` yazıp işi GitHub'ın 10 dakikada bir tetiklemesine
bırakıyordu. **GitHub zamanlanmış işleri yoğunlukta sessizce düşürüyor.** 27 Ağustos
2026'dan itibaren günde beklenen 60 koşu yerine yalnızca **2-4** tanesi çalıştı:

| Tarih | Koşu sayısı |
|---|---|
| 22-25 Ağustos | 23, 23, 15, 15 |
| 27 Ağustos → 19 Eylül | **günde 2-4** |

Sonuç: arka uçlar pencere içinde bile 2,5 saat uyudu. Üstelik 19 Eylül'de bir koşu
**23:57**'de, yani pencere kapandıktan sonra çalışıp boşa kota yaktı.

Artık tetiklemenin kaç kez geldiği önemli değil. İş hangi saatte başlarsa başlasın
**kendi içinde döngü kurup** 10 dakikada bir yokluyor ve pencere bitene kadar ayakta
kalıyor. Cron saat başı deniyor; biri tutarsa o gün için yeter, kaçarsa bir sonraki
devralır. Pencere kapandıktan sonra başlayan koşu hiç yoklama yapmadan çıkar.

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

İki yer birden değişir:

1. `.github/workflows/nabiz.yml` içindeki `cron` satırı — işin ne zaman
   **başlayabileceği**. GitHub UTC kullanır, Türkiye UTC+3'tür:
   `9-18` UTC = 12:00–21:59 Türkiye saati.
2. Aynı dosyadaki `bitis=$(TZ=Europe/Istanbul date -d 'today 22:00' ...)` satırı —
   döngünün **ne zaman duracağı**. Asıl pencereyi bu belirler.

Pencereyi genişletmeden önce yukarıdaki kota tablosuna bak.

## Bir koşu neden saatlerce sürüyor

Sürmesi gerekiyor: iş uyanık tutmak için bekliyor. GitHub'ın iş başına sınırı 6
saattir, biz 5 sa 50 dk'da kendimiz çıkıyoruz. Depo public olduğu için Actions
dakikası sınırsız, bu bekleme hiçbir kotayı yemiyor.
