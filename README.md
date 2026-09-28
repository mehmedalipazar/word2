# SerbestSayfa

Resimleri sürükle-bırakla **istediğiniz yere koyabildiğiniz** ve metin değişse bile **yerinden kaymayan** küçük bir kelime işlemci. Kurulum gerektirmez.

## Çalıştırma

`index.html` dosyasına çift tıklayın (Chrome veya Edge). İnternet bağlantısı gerekmez.

## Web'de yayın

Adres: **https://mehmedalipazar.github.io/word2/**

Uygulama GitHub Pages'te yayınlanır; `main` dalına her gönderimde (git push) kendiliğinden güncellenir
(`.github/workflows/pages.yml`). GitHub Pages dosyaları 10 dakika önbellekte tuttuğu için güncelleme
tarayıcılarda en geç birkaç dakika içinde görünür. Yayına yalnızca uygulama dosyaları girer
(`build.sh`: `index.html`, `css/`, `js/`). Belgeler hiçbir sunucuya gitmez; her kullanıcının kendi tarayıcısında kalır.

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
  (tek, 1,5, çift, birden çok, en az); satır aralığı kutusunda "Paragraftan önce/sonra boşluk ekle/kaldır".
  Liste maddesinde Sol maddenin metninin, Asılı / İlk satır işaretin yeridir
- **Cetvel**: imlecin paragrafının girintileri; ilk satır, asılı, sol ve sağ girinti işaretleri sürüklenir
  (0,25 cm'ye yapışır, Alt ile serbest). Liste maddesinde ilk satır işaretin, asılı metnin yeridir; listenin bütün
  maddeleri seçiliyse bütün düzeye, yoksa seçili maddelere uygulanır (Word gibi)
- **Sekme durakları**: sola, ortaya, sağa ve ondalık (virgülde hizalar) özel duraklar; dolgu `. . .`, `- - -`, `___`
  (içindekiler satırı gibi). Cetvelin solundaki kutu durağın türünü seçer, cetvele tıklamak seçili paragraflara durak
  koyar; durak sürüklenerek taşınır, cetvelin dışına sürüklenince kalkar, çift tık **Sekmeler** penceresini açar
  (Paragraf penceresinde de "Sekmeler…"). Son durağın ötesinde Word'deki gibi her 1,27 cm'de bir durak vardır.
  Ondalık durak virgüle hizalar; virgülsüz sayı ("3.250", "150") durakta biter.
  Resmî yazıdaki "Sayı … ⇥ Tarih" (sağa durak) ve ortalı imza bloğu Word'den böyle gelir, Word'e böyle gider.
- **Başlık numaralandırması** (liste düğmelerinin yanındaki kutu; Word'ün Çok Düzeyli Liste'si): "1. / 1.1. / 1.1.1.",
  "1 / 1.1 / 1.1.1" ya da "I. / A. / 1." Başlık 1–3'e bağlanır. Numara metne yazılmaz: başlık eklenince, silinince ya
  da taşınınca kendiliğinden güncellenir. "Bu başlığı numaralandırma" seçili başlığı numarasız yapar (Önsöz,
  Kaynaklar…); numaralı başlığın başında Backspace de önce numarayı kaldırır. Word'ün başlık stillerine bağlı ya da
  başlıklara verilmiş numaraları böyle açılır, Word'e başlık stillerine bağlı liste olarak gider.
- Listeler Word'ün asılı girintisiyle (işaret 0,63 cm, metin 1,27 cm; alt düzeyler 1,27 cm içeride, •/o/▪ ve 1./a./i.);
  maddenin başında Backspace önce işareti kaldırır (metin yerinde kalır), numara sonraki maddelerde sürer.
- Sayfa sonuna düşen numaralı başlık, liste maddesi ya da asılı girintili paragraf (Kaynaklar) Word'deki gibi numarası,
  madde işareti ve ilk sözcüğüyle birlikte yeni sayfada başlar; ekranda da baskıda da. Resmin yanından akan paragrafın
  ve maddenin girintisi resmin kenarından ölçülür (Word gibi; madde işareti resmin altında kalmaz).
- **Liste girintisi** Word'deki gibi sol kenardan ölçülür ve liste (düzey) ya da madde başına ayarlanabilir (cetvel,
  Paragraf penceresi). Word'den açılan ya da yapıştırılan listenin kendi girintisi korunur (kurumsal şablonlardaki
  sol kenara dayalı madde işareti, "MADDE 1 –" gibi uzun numaralar için geniş asılı girinti), Word'e aynen gider.
  Tab / Shift+Tab ile düzeyi değişen madde yeni düzeyin girintisini alır.
- **Liste biçimi** (liste düğmelerinin yanındaki kutu; Word'ün Madde İşareti / Numaralandırma kitaplığı):
  • – ▪ ➢ ✓ ❖ o madde işaretleri, 1. / 1) / a) / a. / A. / I. / i. / (1) numaraları ve çok düzeyli 1. / 1.1. / 1.1.1.
  Biçim seçili maddelerin düzeyine uygulanır, alt düzeyler (Tab) çok düzeyli biçimi alır. "Numaralandırma değerini
  ayarla…" listeyi istenen numaradan başlatır (ortadaki maddede listeyi böler). Word'den açılan a) / I. / 1) / 1.1. /
  tire listeleri biçimleriyle gelir, ayrı listeler ayrı kalır; Word'e de aynı biçimle gider.
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
  kendiliğinden güncellenir (Şekil, Harita, Fotoğraf, Grafik… her etiket ayrı sayılır); resmin altında ya da üstünde.
  **Bölüme göre numara** (bağlam çubuğundaki kutu): "Şekil 2.1" (Başlık 1'in numarası + sıra, her bölümde yeniden
  başlar); Word'ün "Bölüm numarasını ekle"li yazıları böyle açılır, Word'e alanlarıyla gider
- **Sayfa sonu** (Ctrl+Enter): sonraki metin yeni sayfadan başlar; liste maddesinde listeyi böler (numara sürer)
- Tek geri alma geçmişi (metin + resim işlemleri birlikte). Word'deki gibi Ctrl+Z seçimi geri alınan değişikliğin
  yerine koyar (üzerine yazılan sözcük yeniden seçili olur, imleç yazımın başladığı yerde durur) ve o yer görünmüyorsa
  oraya kaydırır; Ctrl+Y değişikliğin sonuna döner. Her komut (Shift+Tab, başlık, punto, seçimin üzerine yapıştırma…)
  tek adımdır; imleç başka yere götürülüp yazılınca yeni adım başlar.
- **Simge** (Ω düğmesi): son kullanılanlar, birim ve matematik işaretleri (° ± × ≤ ≥ ‰ µ Ω …), Yunan harfleri, oklar,
  noktalama ve para işaretleri; karakter kodu da yazılabilir. Word'deki gibi metinde kodu yazıp Alt+X (00B0 → °).
  İkisi de şekil yazısında (resmin bağlam çubuğunda da Ω var), üst/alt bilgi kutularında (Sayfa yapısı penceresinde
  "Ω Simge…") ve Bul/Değiştir'de de çalışır; simge ekleme ve Alt+X ayrı geri alma adımıdır.
- **Sözcük sayısı**: durum çubuğunda; metin seçiliyken "seçili / toplam" (ör. "7 / 1.250 sözcük"). Sayıya tıklamak
  (ya da Ctrl+Shift+G) sayfa, sözcük, karakter (boşluklu/boşluksuz), paragraf ve satır sayısını gösterir.
- Kaydet/Aç (`.sayfa` dosyası), tarayıcıda otomatik kayıt: her sekmenin kendi kurtarma kaydı var, sayfa kapanırken
  son yazılanlar da korunur; kaydedilmemiş değişiklik varken kapatma/yenilemede tarayıcı uyarır. Kapanan sekmenin
  belgesi, sonra açılan yeni sekmede geri gelir.
- **Kapat**: belgeyi kapatır. Kaydedilmemiş değişiklik varsa Word'deki gibi Kaydet / Kaydetme / Vazgeç sorar; Yeni, Aç
  ve dosyayı sayfaya sürükleyip bırakma da aynı soruyu sorar. Kaydet seçilip kayıt penceresinden vazgeçilirse belge
  açık kalır. Kapatınca boş belge açılır ve sekmenin kurtarma kaydı silinir. Ctrl+W tarayıcıya ayrıldığı için kısayolu yok.
- **Word belgesi açma (.docx)**: metin, başlıklar (numaralarıyla), biçimler, listeler (düzeyleri ve biçimleriyle: a), I.,
  1.1., –…; araya paragraf girse de numara sürer), satır aralıkları, girintiler, sekme durakları, sayfa ayarları, üst/alt bilgi (metin ve sayfa numarası), resimler
  (kırpma/döndürme/metin kaydırma dahil). Resmin hemen altındaki "Şekil 3. …" paragrafı ya da resimle gruplanmış
  yazı kutusu, o resmin şekil yazısı olarak gelir. "Aç" düğmesiyle ya da dosyayı sayfaya sürükleyerek.
- **Word'e aktar (.docx)**: resimler sayfaya göre mutlak konumla yazılır; çapa, resmin sayfasında başlayan bir
  paragrafa konur, böylece Word'de de aynı sayfada ve yerde durur. Şekil yazıları Word'ün "Resim Yazısı" stiliyle ve
  otomatik numara alanıyla yazılır (Word'de "Şekiller Tablosu" eklenebilir). Satır aralığı Word'ün satır birimiyle,
  dul/öksüz satır denetimi ve başlıklarda "sonrakiyle birlikte tut" açık yazılır.
- **PDF / Yazdır**: ekrandaki sayfaların birebir aynısı ("PDF olarak kaydet" seçin); uzun belgede de her sayfa kenar
  boşluklarına uyar ve belge sonuna kadar basılır

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
| Ctrl+Shift+G | Sözcük sayısı |
| Alt+X | İmlecin önündeki onaltılık kodu karaktere çevir (00B0 → °) ya da karakteri koduna |
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
| `js/ruler.js` | Cetvel (girinti işaretleri, sekme durakları) |
| `js/objects.js` | Resim seçimi, sürükleme, boyut/döndürme, kırpma, kılavuzlar, hizalama, pano |
| `js/docx.js` | Word (.docx) dışa aktarma |
| `js/docximport.js` | Word (.docx) içe aktarma |
| `js/zip.js` | Bağımlılıksız ZIP yazıcı |
| `js/main.js` | Düğmeler, kısayollar, yakınlaştırma, sayfa yapısı, yazdırma, otomatik kayıt, sözcük sayısı |

## Bilinen sınırlamalar

- Tablo, dipnot, içindekiler yok. Word'den açılan tablolar düz metne (satır başına bir paragraf) dönüşür;
  dipnotlar, grafikler ve EMF/WMF biçimli resimler alınamaz (açılışta kaç tanesinin atlandığı gösterilir).
- Üst/alt bilgi tek satır düz metindir (biçim, resim, tablo, tek/çift sayfa ayrımı yok); Word'den yalnızca ilk dolu
  satırı alınır, kalanı açılışta bildirilir.
- Çubuk sekme durağı ve belgeye özgü varsayılan durak aralığı yoktur (varsayılan duraklar her 1,27 cm'de).
  Word'ün "Tam" satır aralığı "En az" olarak alınır.
- "A.1." gibi düzeyleri farklı biçimde gösteren numaralar editörde tek biçimle çizilir ("1.1."; Word'e aynen gider).
  Numaradan sonra sekme yerine boşluk ya da hiçbir şey (Word'ün `suff`'u) yoktur. Word'ün "Liste Girintilerini Ayarla"
  penceresi yoktur; liste girintisi cetvelden ya da Paragraf penceresinden ayarlanır. Başlık numaralandırması üç düzeylidir (Word'ün Başlık 4–9'u Başlık 3 olur, numarası metne
  yazılır); Word'de ortada yeniden başlatılmış başlık numaraları editöre metin olarak gelir.
- Eski `.doc` biçimi açılamaz; Word'de ".docx" olarak kaydedin.
- Köprüler yalnızca web (`http`, `https`) ve e-posta (`mailto`) adresleri içindir; belge içi (yer imi) köprüler yok,
  Word'den gelenlerin yalnızca metni alınır.
- Stil listesi Normal ve Başlık 1–3'tür; Word'ün "Konu Başlığı" gibi diğer stilleri biçimleriyle Normal paragraf olur.
  Stilde girinti ve "En az/Tam" satır aralığı yoktur (bunlar paragraf biçimi olarak kalır).
- Kırpmada oran seçenekleri (1:1, 4:3…), sayıyla kırpma ve dışa doğru kırpma yok.
- Şekil yazısında bölüm numarası yalnızca Başlık 1'e göredir (Word'ün Başlık 2–9 seçeneği yok).
- Word/LibreOffice satırları tarayıcıdan biraz farklı kırabilir; uzun belgelerde sayfa sonları birkaç satır kayabilir (resimler yerinde kalır).
- Kare kaydırmada metin resmin yalnızca geniş tarafından akar (Word'deki "En geniş taraf").
- Chrome ve Edge için yazıldı.
