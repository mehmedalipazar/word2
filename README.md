# SerbestSayfa

Resimleri sürükle-bırakla **istediğiniz yere koyabildiğiniz** ve metin değişse bile **yerinden kaymayan** küçük bir kelime işlemci. Kurulum gerektirmez.

## Çalıştırma

`index.html` dosyasına çift tıklayın (Chrome veya Edge). İnternet bağlantısı gerekmez.

## Neden Word'deki gibi kaymıyor?

Word'de resim bir paragrafa **bağlıdır** (çapa). Metin değişince paragraf yer değiştirir, resim de onunla gider.
Burada resim doğrudan **sayfaya ve (x, y) konumuna** bağlıdır. Metin, resimlerin etrafından akar ama resmi hiçbir zaman itmez.

Teknik olarak: metin tek bir `contenteditable` akışıdır. Resimlerin ve sayfa aralarının kapladığı alanlar,
metin akışının içine konan görünmez `float` şeritleriyle kapatılır (`js/core.js` → `computeBands`).

## Özellikler

- A4/A5/A3/Letter, dikey/yatay, kenar boşlukları, sayfa numarası
- Metin: yazı tipi, boyut, kalın/italik/altı çizili/üstü çizili, renk, vurgu, hizalama, satır aralığı, başlıklar, listeler
- Resim: dosya seç / sürükle-bırak / yapıştır; taşı, boyutlandır (köşeler oranı korur, Shift serbest), döndür (Shift: 15°)
- Metin kaydırma: **Kare**, **Üst-Alt**, **Metnin önünde**, **Metnin arkasında**
- Akıllı kılavuzlar: sayfa kenarına, kenar boşluğuna, sayfa ortasına ve diğer resimlere yapışma (Alt ile geçici kapatılır); isteğe bağlı 0,5 cm ızgara
- Hizala/dağıt, öne/arkaya, kilitle, çoğalt, cm cinsinden X/Y/Genişlik/Yükseklik/Açı
- **Şekil yazısı** ("Şekil 1. …"): resme bağlıdır, resimle birlikte taşınır; numaralar belge sırasına göre
  kendiliğinden güncellenir (Şekil, Harita, Fotoğraf, Grafik… her etiket ayrı sayılır); resmin altında ya da üstünde
- **Sayfa sonu** (Ctrl+Enter): sonraki metin yeni sayfadan başlar
- Tek geri alma geçmişi (metin + resim işlemleri birlikte)
- Kaydet/Aç (`.sayfa` dosyası), tarayıcıda otomatik kayıt
- **Word belgesi açma (.docx)**: metin, başlıklar (numaralarıyla), biçimler, listeler, satır aralıkları,
  sayfa ayarları, sayfa numarası, resimler (kırpma/döndürme/metin kaydırma dahil). Resmin hemen altındaki
  "Şekil 3. …" paragrafı ya da resimle gruplanmış yazı kutusu, o resmin şekil yazısı olarak gelir.
  "Aç" düğmesiyle ya da dosyayı sayfaya sürükleyerek.
- **Word'e aktar (.docx)**: resimler sayfaya göre mutlak konumla yazılır, Word'de de aynı yerde durur;
  şekil yazıları Word'ün "Resim Yazısı" stiliyle ve otomatik numara alanıyla yazılır (Word'de "Şekiller Tablosu" eklenebilir)
- **PDF / Yazdır**: ekrandaki sayfaların birebir aynısı ("PDF olarak kaydet" seçin)

## Kısayollar

| Tuş | İş |
| --- | --- |
| Ctrl+Z / Ctrl+Y | Geri al / Yinele |
| Ctrl+S / Ctrl+O / Ctrl+P | Kaydet / Aç / Yazdır |
| Oklar (Shift: 10 px) | Seçili resmi ince ayarla kaydır |
| Delete, Ctrl+D, Esc | Sil, çoğalt, seçimi bırak |
| Enter / F2 / çift tık | Seçili resmin şekil yazısını yaz/düzenle (Enter: bitir, Esc: vazgeç) |
| Ctrl+Enter | Sayfa sonu |
| Alt+tık | Metnin arkasındaki resmi yazının üstünden seç |
| Ctrl+tekerlek, Ctrl + / − / 0 | Yakınlaştırma |

## Dosyalar

| Dosya | İçerik |
| --- | --- |
| `js/core.js` | Durum, sayfa geometrisi, metin akış şeritleri, geri alma, dosya biçimi |
| `js/text.js` | Biçimlendirme komutları, Word/web yapıştırma temizliği |
| `js/objects.js` | Resim seçimi, sürükleme, boyut/döndürme, kılavuzlar, hizalama, pano |
| `js/docx.js` | Word (.docx) dışa aktarma |
| `js/docximport.js` | Word (.docx) içe aktarma |
| `js/zip.js` | Bağımlılıksız ZIP yazıcı |
| `js/main.js` | Düğmeler, kısayollar, yakınlaştırma, sayfa yapısı, yazdırma, otomatik kayıt |

## Bilinen sınırlamalar

- Tablo, dipnot, içindekiler, üst bilgi yok. Word'den açılan tablolar düz metne (satır başına bir paragraf) dönüşür;
  dipnotlar, grafikler ve EMF/WMF biçimli resimler alınamaz (açılışta kaç tanesinin atlandığı gösterilir).
- Eski `.doc` biçimi açılamaz; Word'de ".docx" olarak kaydedin.
- Şekil numaraları düz sayıdır (1, 2, 3…); Word'deki "Şekil 2.3" gibi bölüm numaralı yazılar düz numaraya döner.
- Word/LibreOffice satırları tarayıcıdan biraz farklı kırabilir; uzun belgelerde sayfa sonları birkaç satır kayabilir (resimler yerinde kalır).
- Kare kaydırmada metin resmin yalnızca geniş tarafından akar (Word'deki "En geniş taraf").
- Chrome ve Edge için yazıldı.
