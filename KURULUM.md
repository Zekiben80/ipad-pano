# iPad Pano – Kurulum

Ekranda görünen her şey Almanca (Lisi için). Bu rehber Türkçe.

Sayfa (index.html) herkese açık GitHub Pages'te durur ve kişisel veri içermez.
Takvim, notlar, geri sayımlar ve fotoğraflar senin Google hesabındaki Apps Script'ten, gizli token ile gelir.

## 1. Google tarafı (bir kez, ~10 dk)

1. drive.google.com → Yeni → Google E-Tablolar. Adı: **iPad Pano**.
2. Menü: Uzantılar → Apps Script. `Code.gs` içeriğini sil, `apps-script/Code.gs` içeriğini yapıştır, kaydet.
3. Proje Ayarları (dişli) → Saat dilimi: **Europe/Berlin**.
4. Üstte fonksiyon olarak `einrichten` seç → Çalıştır → izinleri onayla (Takvim, Drive, E-Tablo).
   Tabloda 3 sayfa oluşur: **Notizen**, **Countdown**, **Einstellungen**.
5. Fotoğraflar: Drive'da bir klasör aç (ör. "Pano Fotoğraflar"), fotoğrafları yükle.
   Klasörü aç, adres çubuğunda `/folders/` sonrasındaki kimliği kopyala → Einstellungen sayfasında `foto_ordner_id` alanına yapıştır.
6. Einstellungen → `stadt`: ör. `Köln` (boş bırakırsan iPad konumu kullanılır).
7. Dağıt → Yeni dağıtım → Tür: **Web uygulaması**.
   - Şu kullanıcı olarak yürüt: **Ben**
   - Erişimi olanlar: **Herkes** (token olmadan veri vermez)
   - Dağıt → **Web uygulaması URL'sini** kopyala (`.../exec` ile biter).
8. Test: tarayıcıda `URL?k=TOKEN` aç → JSON görmelisin.

Kodu değiştirirsen: Dağıt → Dağıtımları yönet → düzenle → Sürüm: Yeni sürüm. (URL aynı kalır.)

## 2. iPad tarafı

1. Safari → GitHub Pages adresini aç.
2. Kurulum ekranına web uygulaması URL'sini ve token'ı gir → Kaydet.
3. Paylaş → **Ana Ekrana Ekle** → "Pano". Bundan sonra ana ekrandaki simgeden aç (tam ekran).
   Not: Ana ekran uygulaması Safari'den ayrı hafıza kullanır. Simgeden ilk açılışta kurulumu bir kez daha gir.
4. Ayarlar → Ekran ve Parlaklık → Otomatik Kilit: **Hiçbir Zaman**.
5. Ayarlar → Erişilebilirlik → Yönlendirilmiş Erişim: **Açık** (kod belirle).
   Panoyu aç, ana ekran düğmesine 3 kez bas → Başlat. iPad bu ekranda kilitli kalır.
6. Şarj kablosu takılı kalsın.

## Kullanım

- Geri sayım: Google Tablosunu (Countdown) telefondan veya PC'den düzenle. Pano 5 dakikada bir yenilenir.
  - Countdown: "Jährlich" işaretliyse (doğum günü gibi) her yıl tekrarlar.
- Fotoğraf: Drive klasörüne ekle. Liste 10 dakikada bir güncellenir, fotoğraf dakikada bir değişir.
- Ekrana dokun: 30 sn yalnızca fotoğraf. Gece (23:00–06:00) ekran kararır, dokununca 1 dk açılır.
- Ayarları sıfırlamak: adrese `?api=YENI_URL&k=YENI_TOKEN` ekleyerek aç.
- Demo: adrese `?demo=1` ekle.
