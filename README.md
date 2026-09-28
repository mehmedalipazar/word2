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

- A4/A5/A3/Letter, dikey/yatay, kenar boşlukları
- **Üst ve alt bilgi** (Sayfa yapısı): her bantta sol/orta/sağ yuva; `{sayfa}` ve `{toplam}` alanlarıyla "Sayfa 1 / 5";
  "İlk sayfada gösterme" (kapak sayfası). Sayfanın üst/alt boşluğuna çift tıklayarak ya da araç çubuğundaki
  üst/alt bilgi düğmesiyle de açılır.
- Metin: yazı tipi, boyut, kalın/italik/altı çizili/üstü çizili, üst/alt simge, renk, vurgu, hizalama, başlıklar.
  Yazı tipi ve boyut kutularına Word'deki gibi yazılabilir (ör. 11,5 nk); yazı tipi listesi bilgisayarda yüklü
  yaygın yazı tiplerini (Garamond, Book Antiqua, Arial Narrow… yüklüyse) kendi görünümleriyle gösterir.
- **Stiller** (stil kutusunda "Stili değiştir…"): Normal ve Başlık 1–3'ün yazı tipi, punto, kalın/italik, renk,
  hizalama, önce/sonra aralığı ve satır aralığı; o stildeki bütün paragraflar güncellenir. Normal, belgenin varsayılan
  yazı tipidir. "Seçimden al" imlecin paragrafının biçimini stile alır; "Varsayılan olarak ayarla" stilleri yeni
  belgeler için de saklar. Word'deki gibi stil uygulamak, paragrafın tamamına verilmiş yazı tipi/punto/rengi kaldırır.
  Stiller Word'e styles.xml olarak gider, Word belgesinin Normal ve Başlık 1–3 stilleri de açılışta alınır.
- **Biçim boyacısı** (fırça düğmesi; Ctrl+Shift+C / Ctrl+Shift+V): karakter biçimini (ve imleç tek başınaysa paragraf
  biçimini) başka metne uygular; tek tık bir kez, çift tık sürekli (Esc: bitir); tıklanan sözcüğe de uygulanır.
- **Köprü (bağlantı)**: Ctrl+K ya da köprü düğmesi (ekle, düzenle, kaldır; `www.…` ve e-posta adresi de olur);
  Ctrl+tık bağlantıyı açar. Yazılan web/e-posta adresi boşluk ya da Enter'dan sonra kendiliğinden köprü olur (hemen
  Ctrl+Z geri alır). Web'den/Word'den yapıştırılan köprüler korunur; Word'e köprü olarak aktarılır, Word'deki
  köprüler de köprü olarak gelir.
- Word gibi paragraf işareti: üç tıkla seçip silmek paragrafı bütünüyle siler, sonraki paragrafın stili ve ardındaki
  sayfa sonu korunur. Belge içinde kopyalanan paragraflar stilleriyle, sözcükler yazı tipi ve boyutuyla yapıştırılır;
  liste maddesine yapıştırılan paragraf listeyi bölerek ayrı paragraf olarak girer. Sürükle-bırakla taşıma tek geri
  alma adımıdır.
- **Satır aralığı Word'deki gibi satır cinsinden** (1,0 / 1,08 / 1,15 / 1,5 / 2,0…; "1,5" Word'ün 1,5 satırıdır);
  varsayılan Word'ün Normal stili gibi 1,08 satır + 8 nk paragraf sonrası
- **Paragraf penceresi** (¶ düğmesi): sol/sağ girinti, ilk satır ya da asılı girinti, önce/sonra aralığı, satır aralığı
  (tek, 1,5, çift, birden çok, en az); satır aralığı kutusunda "Paragraftan önce/sonra boşluk ekle/kaldır"
- **Cetvel**: imlecin paragrafının girintileri; ilk satır, asılı, sol ve sağ girinti işaretleri sürüklenir
  (0,25 cm'ye yapışır, Alt ile serbest)
- Sekme durakları Word'deki gibi her 1,27 cm'de
- Listeler Word'ün asılı girintisiyle (işaret 0,63 cm, metin 1,27 cm; alt düzeyler 1,27 cm içeride, •/o/▪ ve 1./a./i.);
  maddenin başında Backspace önce işareti kaldırır (metin yerinde kalır), numara sonraki maddelerde sürer.
  Liste düğmesiyle kaldırılan her madde ayrı paragraf olur; seçilmeyen maddeler liste olarak kalır ve yeniden numaralanır.
- Paragraf başında Backspace: önce ilk satır girintisini, sonra sol girintiyi azaltır; sayfa sonundan sonra sayfa sonunu siler
- **Sayfalama Word gibi**: dul/öksüz satır denetimi (paragrafın ilk ya da son satırı sayfada tek kalmaz), başlıklar
  bölünmez ve sonraki paragrafla aynı sayfada kalır. Uzun belgede yazarken arka planda güncellenir.
- **Bul ve Değiştir** (Ctrl+F / Ctrl+H): büyük/küçük harf, tam sözcük; Türkçe I/ı–İ/i kuralı; "Tümünü değiştir" şekil yazılarını da kapsar ve tek adımda geri alınır
- Resim: dosya seç / sürükle-bırak / yapıştır; taşı, boyutlandır (köşeler oranı korur, Shift serbest), döndür (Shift: 15°).
  Eklenen resim imlecin satırının altına gelir (paragrafın başındaysa satırın yerine); birlikte eklenen resimler alt alta
  dizilir, sığmayan sonraki sayfaya geçer.
- **Kırp** (bağlam çubuğu): siyah tutamaklarla kenarları kesin, resmi sürükleyerek çerçevenin içinde kaydırın
  (Esc/Enter: bitir); **Kırpmayı sıfırla** resmin tamamını geri getirir. Resim dosyası değişmez; Word'e kırpma olarak
  aktarılır, Word'de kırpılmış resimlerin kırpması da değiştirilebilir.
- Resmi başka sayfaya taşıma: sürüklerken pencerenin üst/alt kenarında belge kayar; kes → metinde bir yere tıkla → yapıştır
  resmi imlecin sayfasına ve satırına koyar (imleç yoksa görünen sayfaya); bağlam çubuğunda **Sayfa** alanı
- Metin kaydırma: **Kare**, **Üst-Alt**, **Metnin önünde**, **Metnin arkasında**
- Akıllı kılavuzlar: sayfa kenarına, kenar boşluğuna, sayfa ortasına ve diğer resimlere yapışma (Alt ile geçici kapatılır); isteğe bağlı 0,5 cm ızgara
- Hizala/dağıt, öne/arkaya, kilitle, çoğalt, cm cinsinden X/Y/Genişlik/Yükseklik/Açı
- **Şekil yazısı** ("Şekil 1. …"): resme bağlıdır, resimle birlikte taşınır; numaralar belge sırasına göre
  kendiliğinden güncellenir (Şekil, Harita, Fotoğraf, Grafik… her etiket ayrı sayılır); resmin altında ya da üstünde
- **Sayfa sonu** (Ctrl+Enter): sonraki metin yeni sayfadan başlar; liste maddesinde listeyi böler (numara sürer)
- Tek geri alma geçmişi (metin + resim işlemleri birlikte)
- Kaydet/Aç (`.sayfa` dosyası), tarayıcıda otomatik kayıt: her sekmenin kendi kurtarma kaydı var, sayfa kapanırken
  son yazılanlar da korunur; kaydedilmemiş değişiklik varken kapatma/yenilemede tarayıcı uyarır. Kapanan sekmenin
  belgesi, sonra açılan yeni sekmede geri gelir.
- **Word belgesi açma (.docx)**: metin, başlıklar (numaralarıyla), biçimler, listeler (düzeyleriyle; araya paragraf girse
  de numara sürer), satır aralıkları, girintiler, sayfa ayarları, üst/alt bilgi (metin ve sayfa numarası), resimler
  (kırpma/döndürme/metin kaydırma dahil). Resmin hemen altındaki "Şekil 3. …" paragrafı ya da resimle gruplanmış
  yazı kutusu, o resmin şekil yazısı olarak gelir. "Aç" düğmesiyle ya da dosyayı sayfaya sürükleyerek.
- **Word'e aktar (.docx)**: resimler sayfaya göre mutlak konumla yazılır; çapa, resmin sayfasında başlayan bir
  paragrafa konur, böylece Word'de de aynı sayfada ve yerde durur. Şekil yazıları Word'ün "Resim Yazısı" stiliyle ve
  otomatik numara alanıyla yazılır (Word'de "Şekiller Tablosu" eklenebilir). Satır aralığı Word'ün satır birimiyle,
  dul/öksüz satır denetimi ve başlıklarda "sonrakiyle birlikte tut" açık yazılır.
- **PDF / Yazdır**: ekrandaki sayfaların birebir aynısı ("PDF olarak kaydet" seçin)

## Kısayollar

| Tuş | İş |
| --- | --- |
| Ctrl+Z / Ctrl+Y | Geri al / Yinele |
| Ctrl+S / Ctrl+O / Ctrl+P | Kaydet / Aç / Yazdır |
| Ctrl+F / Ctrl+H | Bul / Bul ve Değiştir (Enter: sonraki, Shift+Enter: önceki, Esc: kapat) |
| Ctrl+B / Ctrl+I / Ctrl+U | Kalın / italik / altı çizili |
| Ctrl+K | Köprü ekle / düzenle (Ctrl+tık: bağlantıyı aç) |
| Ctrl+L / Ctrl+E / Ctrl+R / Ctrl+J | Sola / ortala / sağa / iki yana (resim seçiliyse resmi sola / ortaya / sağa hizalar) |
| Ctrl+Shift+> / Ctrl+Shift+< | Yazıyı bir boyut büyüt / küçült |
| Ctrl+] / Ctrl+[ | Yazıyı 1 nk büyüt / küçült |
| Ctrl+= / Ctrl+Shift+= | Alt simge / üst simge |
| Ctrl+Boşluk | Karakter biçimini temizle |
| Ctrl+Shift+C / Ctrl+Shift+V | Biçimi kopyala / uygula (biçim kopyalanmadıysa Ctrl+Shift+V düz metin yapıştırır) |
| Ctrl+Alt+1 / 2 / 3 | Başlık 1 / 2 / 3 |
| Ctrl+1 / Ctrl+5 / Ctrl+2 | Satır aralığı tek / 1,5 / çift |
| Ctrl+M / Ctrl+Shift+M | Girintiyi artır / azalt |
| Ctrl+Shift+L | Madde işaretleri |
| Tab | Metnin içinde sekme; maddenin başında alt düzeye in; birden çok paragraf seçiliyse girinti |
| Shift+Tab | Maddede üst düzeye çık; paragrafın başında girintiyi azalt |
| Oklar (Shift: 10 px) | Seçili resmi ince ayarla kaydır |
| Delete, Ctrl+D, Esc | Sil, çoğalt, seçimi bırak (Delete ve Esc sonrası yazma metinde, son imleç yerinde sürer) |
| Esc / Enter (kırparken) | Kırpmayı bitir |
| Enter / F2 / çift tık | Seçili resmin şekil yazısını yaz/düzenle (Enter: bitir, Esc: vazgeç) |
| Ctrl+Enter | Sayfa sonu (sonraki paragrafın başında Backspace siler) |
| Alt+tık | Metnin arkasındaki resmi yazının üstünden seç |
| Ctrl+tekerlek, Ctrl+Num+ / Ctrl+− / Ctrl+0 | Yakınlaştırma |

## Dosyalar

| Dosya | İçerik |
| --- | --- |
| `js/core.js` | Durum, sayfa geometrisi, metin akış şeritleri, geri alma, dosya biçimi |
| `js/text.js` | Biçimlendirme komutları, paragraf penceresi, köprüler, pano, Word/web yapıştırma temizliği |
| `js/find.js` | Bul ve değiştir |
| `js/ruler.js` | Cetvel (girinti işaretleri) |
| `js/objects.js` | Resim seçimi, sürükleme, boyut/döndürme, kırpma, kılavuzlar, hizalama, pano |
| `js/docx.js` | Word (.docx) dışa aktarma |
| `js/docximport.js` | Word (.docx) içe aktarma |
| `js/zip.js` | Bağımlılıksız ZIP yazıcı |
| `js/main.js` | Düğmeler, kısayollar, yakınlaştırma, sayfa yapısı, yazdırma, otomatik kayıt |

## Bilinen sınırlamalar

- Tablo, dipnot, içindekiler yok. Word'den açılan tablolar düz metne (satır başına bir paragraf) dönüşür;
  dipnotlar, grafikler ve EMF/WMF biçimli resimler alınamaz (açılışta kaç tanesinin atlandığı gösterilir).
- Üst/alt bilgi tek satır düz metindir (biçim, resim, tablo, tek/çift sayfa ayrımı yok); Word'den yalnızca ilk dolu
  satırı alınır, kalanı açılışta bildirilir.
- Cetvelde sekme durağı yoktur. Word'deki özel sekme durakları alınmaz (varsayılan 1,27 cm durakları geçerlidir).
  Word'ün "Tam" satır aralığı "En az" olarak alınır.
- Çok düzeyli listelerde numaralar düzey başına ayrıdır (1., a., i.); Word'deki "1.1." gibi birleşik numaralar yoktur.
- Eski `.doc` biçimi açılamaz; Word'de ".docx" olarak kaydedin.
- Köprüler yalnızca web (`http`, `https`) ve e-posta (`mailto`) adresleri içindir; belge içi (yer imi) köprüler yok,
  Word'den gelenlerin yalnızca metni alınır.
- Stil listesi Normal ve Başlık 1–3'tür; Word'ün "Konu Başlığı" gibi diğer stilleri biçimleriyle Normal paragraf olur.
  Stilde girinti ve "En az/Tam" satır aralığı yoktur (bunlar paragraf biçimi olarak kalır).
- Kırpmada oran seçenekleri (1:1, 4:3…), sayıyla kırpma ve dışa doğru kırpma yok.
- Şekil numaraları düz sayıdır (1, 2, 3…); Word'deki "Şekil 2.3" gibi bölüm numaralı yazılar düz numaraya döner.
- Word/LibreOffice satırları tarayıcıdan biraz farklı kırabilir; uzun belgelerde sayfa sonları birkaç satır kayabilir (resimler yerinde kalır).
- Kare kaydırmada metin resmin yalnızca geniş tarafından akar (Word'deki "En geniş taraf").
- Chrome ve Edge için yazıldı.
