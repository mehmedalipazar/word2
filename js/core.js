/* Çekirdek: durum, sayfa geometrisi, metin akışı düzeni, geri alma geçmişi, belge serileştirme.
 *
 * Temel fikir: Resimler metnin İÇİNDE durmaz. Her resim bir sayfaya ve o sayfadaki
 * (x, y) konumuna sabitlenir. Metin tek bir contenteditable akışıdır; resimlerin ve sayfa
 * aralarının kapladığı yerler görünmez "float" şeritleriyle metne kapatılır. Böylece metin
 * resimlerin etrafından akar ama resimler metin değiştikçe asla yer değiştirmez.
 */
(function () {
  'use strict';
  const SS = window.SS;
  const U = SS.units;
  const $ = (id) => document.getElementById(id);

  const app = (SS.app = {});
  const els = (app.els = {
    workspace: $('workspace'),
    scaler: $('scaler'),
    doc: $('doc'),
    pages: $('pagesLayer'),
    behind: $('behindLayer'),
    front: $('frontLayer'),
    overlay: $('overlay'),
    flow: $('flow'),
    editor: $('editor'),
  });

  const defaultPage = () => {
    const m = U.cmToPx(2.5);
    return { size: 'A4', orient: 'portrait', margins: { t: m, r: m, b: m, l: m } };
  };
  app.defaultPage = defaultPage;

  // Üst/alt bilgi: her bantta sol/orta/sağ üç yuva; {sayfa} ve {toplam} alanları yazılabilir
  const emptyHF = () => ({ header: ['', '', ''], footer: ['', '', ''], firstPage: false });
  app.emptyHF = emptyHF;

  // ---------- Belge stilleri (Word'ün Normal ve Başlık 1–3'ü) ----------
  // font: yazı tipi adı, size: punto, color: #rrggbb, align: left|center|right|justify, before/after: paragraf
  // öncesi/sonrası aralık (nk), line: satır aralığı (Word satır katı). Editörün CSS'i bunlardan üretilir
  // (applyStyles); stildeki paragraflarda satır içi biçim yalnızca stilden farklı olan için vardır. Yerleşik
  // değerler css/app.css'teki varsayılanlarla aynı.
  // Word'ün "1,08 satır"ı aslında 259/240'tır (w:line="259"); Calibri'de 1,3177 (css: .editor)
  const L108 = 259 / 240;
  const BUILTIN_STYLES = {
    p: { font: 'Calibri', size: 11, bold: false, italic: false, color: '#000000', align: 'left', before: 0, after: 8, line: L108 },
    h1: { font: 'Calibri', size: 18, bold: true, italic: false, color: '#000000', align: 'left', before: 12, after: 6, line: L108 },
    h2: { font: 'Calibri', size: 14, bold: true, italic: false, color: '#000000', align: 'left', before: 10, after: 5, line: L108 },
    h3: { font: 'Calibri', size: 12, bold: true, italic: false, color: '#000000', align: 'left', before: 8, after: 4, line: L108 },
    num: null,
  };
  app.STYLE_KEYS = ['p', 'h1', 'h2', 'h3'];
  app.STYLE_NAMES = { p: 'Normal', h1: 'Başlık 1', h2: 'Başlık 2', h3: 'Başlık 3' };
  app.builtinStyles = () => JSON.parse(JSON.stringify(BUILTIN_STYLES));
  const inRange = (v, lo, hi, d) => (typeof v === 'number' && isFinite(v) && v >= lo && v <= hi ? v : d);

  // Başlık numaralandırması (Word: Çok Düzeyli Liste, Başlık 1–3'e bağlı): styles.num = null ya da
  // [Başlık 1, Başlık 2, Başlık 3] düzeylerinin tanımı (numarasız düzey null). text: Word'ün lvlText'i (%1–%3: Başlık
  // 1–3'ün sayacı), fmt: sayı biçimi, start: ilk değer, ind/hang: sol girinti ve asılı girinti (nk; hang < 0 ilk satır
  // girintisi), suff: numaradan sonra sekme/boşluk/hiçbiri, lgl: bütün sayılar ondalık (Word: isLgl). Numara metne
  // yazılmaz, CSS sayacıyla çizilir (applyStyles): başlık eklenince, silinince, taşınınca kendiliğinden güncellenir.
  // data-num="0" olan başlık numarasızdır (Word: numId 0).
  app.NUM_FMTS = { decimal: 'decimal', decimalZero: 'decimal-leading-zero', upperRoman: 'upper-roman', lowerRoman: 'lower-roman',
    upperLetter: 'upper-alpha', lowerLetter: 'lower-alpha' };
  const numLevel = (text, fmt, ind) => ({ text, fmt, start: 1, ind, hang: ind, suff: 'tab', lgl: false });
  // Word'ün başlıklara bağlı hazır listeleri (girintiler Word'ünkü: 0,76 / 1,02 / 1,27 cm)
  app.HEADING_NUMS = {
    '1.': [numLevel('%1.', 'decimal', 21.6), numLevel('%1.%2.', 'decimal', 28.8), numLevel('%1.%2.%3.', 'decimal', 36)],
    1: [numLevel('%1', 'decimal', 21.6), numLevel('%1.%2', 'decimal', 28.8), numLevel('%1.%2.%3', 'decimal', 36)],
    'I.': [numLevel('%1.', 'upperRoman', 21.6), numLevel('%2.', 'upperLetter', 28.8), numLevel('%3.', 'decimal', 36)],
  };
  function cleanNum(src) {
    if (!Array.isArray(src) || src.length !== 3) return null;
    const out = src.map((L, i) => {
      if (!L || typeof L !== 'object' || typeof L.text !== 'string' || L.text.length > 60) return null;
      // eslint-disable-next-line no-control-regex
      const text = L.text.replace(/[\u0000-\u001f]/g, '');
      // Yalnızca kendi ve üst düzeylerin sayacı (%1 … %düzey) kullanılabilir
      if (/%(?![1-3])/.test(text) || [...text.matchAll(/%([1-3])/g)].some((m) => +m[1] > i + 1)) return null;
      return {
        text,
        fmt: app.NUM_FMTS[L.fmt] ? L.fmt : 'decimal',
        start: Math.round(inRange(L.start, 0, 32767, 1)),
        ind: inRange(L.ind, -1584, 1584, 0),
        hang: inRange(L.hang, -1584, 1584, 0),
        suff: /^(tab|space|nothing)$/.test(L.suff) ? L.suff : 'tab',
        lgl: L.lgl === true,
      };
    });
    return out.some(Boolean) ? out : null;
  }
  // Numaranın örneği (ör. "1. / 1.1. / 1.1.1."; ilk numara 1 değilse "(ilk: 3)"): araç çubuğunda gösterilir
  app.numSample = function (N) {
    const ONE = { decimal: '1', decimalZero: '01', upperRoman: 'I', lowerRoman: 'i', upperLetter: 'A', lowerLetter: 'a' };
    const levels = (N || []).filter(Boolean);
    const s = levels.map((L) => L.text.replace(/%([1-3])/g, (_, d) => (L.lgl || !N[d - 1] ? '1' : ONE[N[d - 1].fmt]))).join(' / ');
    return levels.length && levels[0].start !== 1 ? `${s} (ilk: ${levels[0].start})` : s;
  };

  // Dışarıdan gelen (dosya, yerel kayıt) stilleri denetle; eksik ya da bozuk alan yerleşik değeri alır
  app.cleanStyles = function (src, base = BUILTIN_STYLES) {
    const out = {};
    for (const k of app.STYLE_KEYS) {
      const s = (src && typeof src[k] === 'object' && src[k]) || {};
      const b = base[k] || BUILTIN_STYLES[k];
      out[k] = {
        font: typeof s.font === 'string' && /^[\p{L}\p{N} ._-]{1,48}$/u.test(s.font.trim()) ? s.font.trim() : b.font,
        size: inRange(s.size, 1, 1638, b.size),
        bold: typeof s.bold === 'boolean' ? s.bold : b.bold,
        italic: typeof s.italic === 'boolean' ? s.italic : b.italic,
        color: typeof s.color === 'string' && /^#[0-9a-f]{6}$/i.test(s.color) ? s.color.toLowerCase() : b.color,
        align: /^(left|center|right|justify)$/.test(s.align) ? s.align : b.align,
        before: inRange(s.before, 0, 1584, b.before),
        after: inRange(s.after, 0, 1584, b.after),
        line: inRange(s.line, 0.06, 132, b.line),
      };
    }
    out.num = cleanNum(src && src.num);
    return out;
  };
  // Yeni belgelerin stilleri: "Varsayılan olarak ayarla" ile saklananlar, yoksa yerleşik
  const DEFAULT_STYLES_KEY = 'serbestsayfa-varsayilan-stiller';
  app.defaultStyles = function () {
    try {
      const s = JSON.parse(localStorage.getItem(DEFAULT_STYLES_KEY) || 'null');
      if (s) return app.cleanStyles(s);
    } catch (_) { /* bozuk kayıt: yerleşik */ }
    return app.builtinStyles();
  };
  app.saveDefaultStyles = function (styles) {
    try {
      localStorage.setItem(DEFAULT_STYLES_KEY, JSON.stringify(app.cleanStyles(styles)));
      return true;
    } catch (_) {
      return false;
    }
  };
  // Yazı tipinin CSS yığını: bulunmazsa editörün (ve SS.lineFactor'ün) yedek sırası kullanılır
  app.fontStack = (f) => (/^calibri$/i.test(f) ? 'Calibri, Carlito, "Segoe UI", sans-serif' : `"${String(f).replace(/["\\]/g, '')}", Calibri, Carlito, sans-serif`);
  // Stilin birimsiz CSS satır yüksekliği: Word satır katı × yazı tipinin tek satır katsayısı (bkz. lineSpacing)
  app.styleLineHeight = (s) => +(s.line * SS.lineFactor(s.font)).toFixed(4);

  const state = (app.state = {
    page: defaultPage(),
    hf: emptyHF(),
    styles: app.builtinStyles(),
    minPages: 1,
    pageCount: 1,
    images: [], // { id, asset, page, x, y, w, h, rot, wrap, locked, z }  (x,y: sayfa sol-üst köşesine göre px)
    assets: {}, // { [assetId]: { src, mime, nw, nh } }
    selection: [],
    zoom: 1,
    gap: 24,
    snap: true,
    grid: false,
    fileName: 'Adsız belge',
  });

  app.WRAP_DIST = U.cmToPx(0.3); // resim ile metin arası boşluk
  app.MIN_TEXT_W = U.cmToPx(1.8); // bundan dar aralıklara metin sokulmaz
  app.FOOTER_DIST = U.cmToPx(1.25);
  app.HEADER_DIST = U.cmToPx(1.25); // Word'ün varsayılanı (w:header="709")
  app.CAP_GAP = U.cmToPx(0.15); // resim ile şekil yazısı arası
  app.CAP_MIN_W = U.cmToPx(4); // şekil yazısının en küçük genişliği
  const EXTRA_BANDS = 3; // metnin taşabileceği ileri sayfalar için önceden kurulan şerit sayısı (fazlası düzeni yavaşlatır)

  app.geom = function () {
    // "Custom": içe aktarılan Word belgesindeki standart dışı kağıt (dikey yöndeki cm ölçüleri)
    const [wcm, hcm] =
      state.page.size === 'Custom' && state.page.custom ? state.page.custom : SS.PAGE_SIZES[state.page.size] || SS.PAGE_SIZES.A4;
    let PW = U.cmToPx(wcm);
    let PH = U.cmToPx(hcm);
    if (state.page.orient === 'landscape') [PW, PH] = [PH, PW];
    const m = state.page.margins;
    return { PW, PH, m, cw: PW - m.l - m.r, ch: PH - m.t - m.b, stride: PH + state.gap, gap: state.gap };
  };

  app.getImage = (id) => state.images.find((i) => i.id === id);
  app.wrapsText = (img) => img.wrap === 'square' || img.wrap === 'topbottom';

  // Döndürülmüş resmin sayfa koordinatlarındaki eksen hizalı sınır kutusu
  app.aabb = function (img) {
    const r = ((img.rot || 0) * Math.PI) / 180;
    const c = Math.abs(Math.cos(r));
    const s = Math.abs(Math.sin(r));
    const w = img.w * c + img.h * s;
    const h = img.w * s + img.h * c;
    return { x: img.x + img.w / 2 - w / 2, y: img.y + img.h / 2 - h / 2, w, h };
  };

  // ---------- Şekil yazısı (resme bağlı, resimle birlikte taşınır) ----------
  // img.caption = { text, label: 'Şekil', pos: 'below' | 'above' }
  // Yazı döndürülmez; resmin sınır kutusunun altında/üstünde, ortalanmış durur.
  const capH = new Map(); // ölçülen yazı yükseklikleri (px)
  app.capRect = function (img) {
    if (!img.caption) return null;
    const bb = app.aabb(img);
    const w = Math.max(bb.w, app.CAP_MIN_W);
    const h = capH.get(img.id) || 18;
    const y = img.caption.pos === 'above' ? bb.y - app.CAP_GAP - h : bb.y + bb.h + app.CAP_GAP;
    return { x: bb.x + bb.w / 2 - w / 2, y, w, h };
  };
  // Resim + yazısı birlikte ("şekil"): kaydırma, hizalama ve yapışma bu kutuyla yapılır
  app.objBox = function (img) {
    const bb = app.aabb(img);
    const c = app.capRect(img);
    if (!c) return bb;
    const x0 = Math.min(bb.x, c.x);
    const y0 = Math.min(bb.y, c.y);
    return { x: x0, y: y0, w: Math.max(bb.x + bb.w, c.x + c.w) - x0, h: Math.max(bb.y + bb.h, c.y + c.h) - y0 };
  };
  // Numaralar belge sırasına göre (sayfa, üstten alta, soldan sağa), her etiket kendi içinde. caption.chapter
  // (ayırıcı: "." "-" "–" "—") bölüm numaralıdır (Word: Resim Yazısı > Numaralandırma > Bölüm numarasını ekle):
  // numara, resmin üstünde kalan son Başlık 1'in numarası + ayırıcı + o bölümdeki sıra ("Şekil 2.1"); sıra her
  // Başlık 1'de yeniden başlar. Döner: id → { n: sıra, ch: bölüm numarası ya da null, sep, text: görünen numara }
  app.captionNumbers = function () {
    const list = state.images.filter((i) => i.caption).sort((a, b) => a.page - b.page || a.y - b.y || a.x - b.x);
    const sepOf = (c) => (/^[.:\-–—]$/.test(c.chapter) ? c.chapter : null);
    const marks = list.some((i) => sepOf(i.caption)) ? chapterMarks() : [];
    const counters = {};
    const nums = new Map();
    for (const img of list) {
      const l = img.caption.label || 'Şekil';
      const sep = sepOf(img.caption);
      let ch = null;
      let key = l;
      if (sep) {
        const top = app.aabb(img).y;
        let m = null;
        for (const h of marks) if (h.page < img.page || (h.page === img.page && h.y <= top)) m = h;
        ch = m ? m.text : chapterText(0);
        key = `${l}\u0000${m ? m.idx : -1}`;
      }
      counters[key] = (counters[key] || 0) + 1;
      nums.set(img.id, { n: counters[key], ch, sep, text: sep ? `${ch}${sep}${counters[key]}` : String(counters[key]) });
    }
    return nums;
  };
  // Başlık 1'lerin yeri (sayfa, sayfanın üstünden y) ve numarası: başlık numaralandırması varsa onun sayacı ve
  // biçimi (numarasız başlık yeni bölüm başlatır, numarası bir öncekininkidir), yoksa başlık sırası
  const romanUp = (n) => {
    let s = '';
    for (const [v, r] of [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']])
      for (; n >= v; n -= v) s += r;
    return s;
  };
  function chapterText(c) {
    const L = state.styles.num && state.styles.num[0];
    const f = !L || L.lgl ? 'decimal' : L.fmt;
    if (c < 1 || f === 'decimal') return String(c);
    if (f === 'decimalZero') return String(c).padStart(2, '0');
    if (/Roman$/.test(f)) return f === 'upperRoman' ? romanUp(c) : romanUp(c).toLowerCase();
    let s = '';
    for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(97 + ((n - 1) % 26)) + s;
    return f === 'upperLetter' ? s.toUpperCase() : s;
  }
  function chapterMarks() {
    const L = state.styles.num && state.styles.num[0];
    const g = app.geom();
    const fr = els.flow.getBoundingClientRect();
    const z = state.zoom;
    let c = L ? L.start - 1 : 0;
    return [...els.editor.querySelectorAll('h1')].map((h, idx) => {
      if (!L || h.getAttribute('data-num') !== '0') c++;
      const y = (h.getBoundingClientRect().top - fr.top) / z;
      const page = Math.max(0, Math.floor(y / g.stride));
      return { page, y: y - page * g.stride + g.m.t, idx, text: chapterText(c) };
    });
  }

  // ---------- Metin dışlama şeritleri ----------
  // Akış koordinatları: (0,0) = ilk sayfanın metin alanının sol-üst köşesi.
  // k. sayfanın metin alanı: y ∈ [k*stride, k*stride + ch]
  const keep = []; // keep[k]: sayfalama kuralları için k. sayfanın metin alanından alttan kısılan yükseklik (px)
  function computeBands() {
    const { cw, ch, stride, m } = app.geom();
    const full = [];
    const left = [];
    const right = [];
    // Yazdırırken yalnızca gerçek sayfa sınırları: akış kutusunun dışına taşan float'lar
    // tarayıcının sayfalara bölme hesabını bozuyor (son sayfanın üst boşluğu kayboluyordu).
    const bandCount = app.printMode ? state.pageCount - 1 : state.pageCount + EXTRA_BANDS;
    for (let k = 0; k < bandCount; k++) full.push([k * stride + ch, (k + 1) * stride]);
    // Sayfalama kuralları (dul/öksüz satır, sonrakiyle birlikte tut): sayfanın metin alanı alttan kısalır
    for (let k = 0; k < Math.min(keep.length, state.pageCount - 1); k++)
      if (keep[k] > 0.5) full.push([k * stride + ch - keep[k], k * stride + ch]);

    const d = app.WRAP_DIST;
    for (const img of state.images) {
      if (!app.wrapsText(img)) continue;
      const bb = app.objBox(img);
      const top = img.page * stride;
      const y0 = Math.max(top + bb.y - m.t - d, top);
      const y1 = Math.min(top + bb.y + bb.h - m.t + d, top + ch);
      const x0 = Math.max(bb.x - m.l - d, 0);
      const x1 = Math.min(bb.x + bb.w - m.l + d, cw);
      if (y1 - y0 < 0.5 || x1 - x0 < 0.5) continue;
      if (img.wrap === 'topbottom') {
        full.push([y0, y1]);
        continue;
      }
      // Kare kaydırma: metin daha geniş olan tarafa akar (Word'deki "En geniş taraf")
      const spaceL = x0;
      const spaceR = cw - x1;
      if (Math.max(spaceL, spaceR) < app.MIN_TEXT_W) full.push([y0, y1]);
      else if (spaceR >= spaceL) left.push([y0, y1, x1]);
      else right.push([y0, y1, x0]);
    }

    const ys = new Set([0]);
    for (const list of [full, left, right]) for (const it of list) ys.add(it[0]).add(it[1]);
    // Yazdırırken hiçbir float fiziksel sayfa sınırını aşmasın: sınırlarda böl, birleştirme
    const cuts = [];
    if (app.printMode) for (let k = 1; k < state.pageCount; k++) cuts.push(k * stride - m.t);
    cuts.forEach((y) => ys.add(y));
    const isCut = (y) => cuts.some((c) => Math.abs(c - y) < 0.01);
    const sorted = [...ys].sort((a, b) => a - b);
    const bands = [];
    for (let i = 0; i < sorted.length - 1; i++) {
      const a = sorted[i];
      const b = sorted[i + 1];
      if (b - a < 0.01) continue;
      const mid = (a + b) / 2;
      const inside = (it) => it[0] <= mid && mid < it[1];
      let xL = 0;
      let xR = cw;
      if (full.some(inside)) {
        xL = xR = cw;
      } else {
        for (const it of left) if (inside(it)) xL = Math.max(xL, it[2]);
        for (const it of right) if (inside(it)) xR = Math.min(xR, it[2]);
        if ((xL > 0 || xR < cw) && xR - xL < app.MIN_TEXT_W) xL = xR = cw;
      }
      if (xL === 0 && xR === cw) continue;
      const last = bands[bands.length - 1];
      if (last && Math.abs(last.b - a) < 0.01 && last.xL === xL && last.xR === xR && !isCut(a)) last.b = b;
      else bands.push({ a, b, xL, xR });
    }
    return bands;
  }

  function mkFloat(side, clear, w, h, mt) {
    const d = document.createElement('div');
    d.className = 'ex';
    d.style.cssText = `float:${side};clear:${clear};width:${w}px;height:${h}px;margin-top:${mt}px`;
    return d;
  }

  // Her şerit bir sol ve/veya sağ float'tur. "clear" ile alt alta dizilir, margin-top ile
  // doğru yüksekliğe itilir; shape-outside:border-box sayesinde margin alanı metni itmez.
  // Float'lar doğrudan #flow'un çocuğudur: #flow tüm sayfalar boyunca uzandığı için yazdırırken
  // tarayıcı şeritleri sayfalara doğru böler (sıfır yükseklikli bir kapsayıcı içinde bölünmüyordu).
  let sideBands = []; // metnin yanından aktığı şeritler (resmin yanı); fitTabs bunların yanındaki paragrafı yeniden ölçer
  function renderExclusions() {
    const { cw } = app.geom();
    const frag = document.createDocumentFragment();
    let prev = 0;
    const bands = computeBands();
    sideBands = bands.filter((bd) => bd.xL < cw);
    for (const bd of bands) {
      const mt = bd.a - prev;
      const h = bd.b - bd.a;
      const hasLeft = bd.xL > 0;
      // Not: float kaptan geniş olmamalı; Chrome o zaman margin-top alanını da metne kapatıyor.
      if (hasLeft) frag.appendChild(mkFloat('left', 'both', bd.xL, h, mt));
      if (bd.xR < cw) frag.appendChild(mkFloat('right', hasLeft ? 'right' : 'both', cw - bd.xR, h, mt));
      prev = bd.b;
    }
    els.flow.querySelectorAll(':scope > .ex').forEach((el) => el.remove());
    els.flow.insertBefore(frag, els.editor);
  }

  // Tarayıcının düzenleme komutları (kalın, punto, hizalama, liste…) her adımda satırları yeniden dizer; şerit
  // float'ları dururken bu, uzun belgede belge boyuyla karesel büyür (47 sayfada Tümünü seç + Kalın: 14 sn).
  // Komut süresince şeritler düzenden çıkarılır (css: .flow.nobands), ardından düzen bir kez yeniden kurulur.
  app.withoutBands = function (fn) {
    const ws = els.workspace;
    const top = ws.scrollTop;
    const was = els.flow.classList.contains('nobands');
    els.flow.classList.add('nobands');
    try {
      return fn();
    } finally {
      if (!was) els.flow.classList.remove('nobands');
      ws.scrollTop = top; // şeritsiz kısa düzende yapılan kaydırma geri alınsın
    }
  };

  // Akış koordinatına çevirici (ekran -> akış)
  app.clientToFlow = function (clientX, clientY) {
    const f = els.flow.getBoundingClientRect();
    return { x: (clientX - f.left) / state.zoom, y: (clientY - f.top) / state.zoom };
  };

  function textBottom() {
    const last = els.editor.lastElementChild || els.editor;
    return app.clientToFlow(0, last.getBoundingClientRect().bottom).y;
  }
  app.textBottom = textBottom;

  function neededPages() {
    const { stride } = app.geom();
    const textPages = Math.floor(Math.max(0, textBottom() - 0.5) / stride) + 1;
    let imgPages = 0;
    for (const img of state.images) imgPages = Math.max(imgPages, img.page + 1);
    return Math.max(textPages, imgPages, state.minPages, 1);
  }

  // ---------- Sayfa ve resim çizimi ----------
  function updateDocSize() {
    const g = app.geom();
    const h = state.pageCount * g.stride - g.gap;
    const ds = els.doc.style;
    ds.width = g.PW + 'px';
    ds.height = h + 'px';
    ds.transform = `scale(${state.zoom})`;
    ds.setProperty('--z', state.zoom);
    els.scaler.style.width = g.PW * state.zoom + 'px';
    els.scaler.style.height = h * state.zoom + 'px';
    const fs = els.flow.style;
    fs.left = g.m.l + 'px';
    fs.top = g.m.t + 'px';
    fs.width = g.cw + 'px';
    fs.height = Math.max(0, h - g.m.t) + 'px';
  }
  app.updateDocSize = updateDocSize;

  app.fillHF = (s, k) => s.replace(/\{sayfa\}/g, String(k + 1)).replace(/\{toplam\}/g, String(state.pageCount));

  // Stillerin CSS'i (css/app.css'teki varsayılanların üstüne; bu <style> ondan sonra gelir). Başlıklar kendi yazı
  // tipinin satır katsayısıyla satır yüksekliği alır. Üst/alt bilgi ve şekil yazısı Word'deki gibi Normal'in yazı tipinde.
  const docStyleEl = document.head.appendChild(document.createElement('style'));
  docStyleEl.id = 'docStyles';
  app.applyStyles = function () {
    const S = state.styles;
    const font = (s) =>
      `font-family: ${app.fontStack(s.font)}; font-size: ${s.size}pt; font-weight: ${s.bold ? 'bold' : 'normal'}; ` +
      `font-style: ${s.italic ? 'italic' : 'normal'}; color: ${s.color}; text-align: ${s.align}; line-height: ${app.styleLineHeight(s)};`;
    const css = [
      `.editor { ${font(S.p)} }`,
      `.editor p, .editor div, .editor ul, .editor ol { margin: ${S.p.before}pt 0 ${S.p.after}pt; }`,
      ...['h1', 'h2', 'h3'].map((k) => `.editor ${k} { ${font(S[k])} margin: ${S[k].before}pt 0 ${S[k].after}pt; }`),
      `.page .hf, .cap { font-family: ${app.fontStack(S.p.font)}; }`,
      ...headingNumCSS(S.num),
    ].join('\n');
    if (docStyleEl.textContent !== css) docStyleEl.textContent = css;
    if (app.onStyles) app.onStyles();
  };
  // Başlık numaraları: sayaçlar (hn1–hn3) editörde kurulur; Başlık k, hn<k>'yi artırır ve alt düzeylerinkini
  // sıfırlar (counter-set: kardeş başlıklarda counter-reset yeni sayaç açıp alt düzeyi sıfırlamıyor). Numara
  // ::before'da; Word'deki gibi numaradan sonraki sekme asılı girintiye (yoksa sonraki 1,27 cm durağına) gider,
  // numara oradan uzunsa kısa bir boşluk kalır.
  function headingNumCSS(N) {
    if (!N) return [];
    const q = (s) => `"${s.replace(/[\\"]/g, '\\$&')}"`;
    const reset = (from) => N.slice(from).map((L, j) => `hn${from + j + 1} ${L ? L.start - 1 : 0}`).join(' ');
    const out = [`.editor { counter-reset: ${reset(0)}; }`];
    N.forEach((L, i) => {
      if (!L) return;
      const sel = `.editor h${i + 1}:not([data-num="0"])`;
      const content = L.text.split(/(%[1-3])/).filter(Boolean)
        .map((t) => (/^%[1-3]$/.test(t) ? `counter(hn${t[1]}, ${app.NUM_FMTS[L.lgl || !N[t[1] - 1] ? 'decimal' : N[t[1] - 1].fmt]})` : q(t)));
      if (L.suff === 'space') content.push('" "');
      const first = L.ind - L.hang; // ilk satırın (numaranın) başladığı yer
      const tabW = L.hang > 0 ? L.hang : 36 - (((first % 36) + 36) % 36);
      out.push(`${sel} { counter-increment: hn${i + 1};${i < 2 ? ` counter-set: ${reset(i + 1)};` : ''} margin-left: ${L.ind}pt; text-indent: ${-L.hang}pt; }`);
      out.push(`${sel}::before { content: ${content.join(' ') || '""'}; text-indent: 0;` +
        `${L.suff === 'tab' ? ` display: inline-block; box-sizing: border-box; min-width: ${tabW}pt; padding-right: 0.3em;` : ''} }`);
    });
    return out;
  }
  // Başlık numaralı mı (düzeyi numaralı ve data-num="0" değil)
  app.isNumbered = (b) => {
    const N = state.styles.num;
    return !!(N && /^H[123]$/.test(b.tagName) && b.getAttribute('data-num') !== '0' && N[b.tagName[1] - 1]);
  };
  // Paragrafın girintisini yaz (prop: marginLeft, marginRight, textIndent; px). Stildeki değerle (numaralı başlıkta
  // numaralandırmanın girintisi, diğerlerinde 0) aynıysa satır içi biçim kalkar; farklıysa 0 da olsa yazılır.
  app.setIndent = function (b, prop, px) {
    const L = app.isNumbered(b) && state.styles.num[b.tagName[1] - 1];
    const base = !L || prop === 'marginRight' ? 0 : ((prop === 'marginLeft' ? L.ind : -L.hang) * 4) / 3;
    b.style[prop] = Math.abs(px - base) > 0.5 ? px + 'px' : '';
    if (!b.getAttribute('style')) b.removeAttribute('style');
  };
  app.applyStyles();
  // Stil değişikliği: bütün sayfalar yeniden dizilir, tek geri alma adımı
  app.setStyles = function (styles) {
    if (JSON.stringify(styles) === JSON.stringify(state.styles)) return;
    state.styles = styles;
    app.applyStyles();
    app.relayoutAll();
    app.markDirty();
    app.commit('edit');
  };

  // ---------- Liste biçimleri (Word: Madde İşareti / Numaralandırma kitaplığı, Çok Düzeyli Liste) ----------
  // Listenin (ol/ul) data-lf özniteliği düzeylerinin biçimidir: "düzey1;düzey2;…", her düzey "biçim|metin". Biçim
  // Word'ün numFmt'si (decimal, decimalZero, lowerLetter, upperLetter, lowerRoman, upperRoman) ya da bullet; metin
  // Word'ün lvlText'i: %1 listenin kendi maddeleri, %2 bir kat içerideki liste… (madde işaretinde işaretin kendisi).
  // Düzey 1 listenin kendi maddelerine, düzey k kendi biçimi olmayan k-1 kat içerideki listelere uygulanır (türü
  // tutuyorsa: numaralı düzey ol'a, madde işareti ul'ye). Verilmeyen düzey css/app.css'teki varsayılandır (1. a. i. /
  // • o ▪). İşaretler CSS sayacıyla çizilir (listFmtCSS); %1.%2. gibi art arda düzeyler counters() ile; ol[start]
  // numarayı başlatır.
  const LF_FMTS = [...Object.keys(app.NUM_FMTS), 'bullet'];
  app.parseLF = (s) =>
    String(s || '').split(';').map((t) => {
      const m = /^([a-zA-Z]+)\|([^;|]{1,30})$/u.exec(t);
      // eslint-disable-next-line no-control-regex
      return m && LF_FMTS.includes(m[1]) && !/[\u0000-\u001f]/.test(m[2]) ? { fmt: m[1], text: m[2] } : null;
    });
  app.formatLF = (levels) => {
    const out = levels.map((L) => (L ? `${L.fmt}|${L.text}` : ''));
    while (out.length && !out[out.length - 1]) out.pop();
    return out.join(';');
  };
  // css/app.css'teki varsayılan işaret (k: 0'dan başlayan düzey)
  app.defaultLevel = (k, ordered) =>
    ordered ? { fmt: ['decimal', 'lowerLetter', 'lowerRoman'][Math.min(k, 2)], text: `%${k + 1}.` } : { fmt: 'bullet', text: ['•', 'o', '▪'][Math.min(k, 2)] };
  // Düzey metnindeki yer tutucuları kaydır (ör. kitaplıktaki "%1)" 2. düzeye "%2)" olarak konur)
  app.shiftLevel = (L, by) => L && { ...L, text: L.fmt === 'bullet' ? L.text : L.text.replace(/%([1-9])/g, (_, d) => '%' + Math.min(9, Math.max(1, +d + by))) };
  const cssStr = (s) => `"${String(s).replace(/[\\"]/g, '\\$&')}"`;
  function markerContent(L) {
    const ph = L.fmt === 'bullet' ? [] : [...L.text.matchAll(/%[1-9]/g)];
    if (!ph.length) return `${cssStr(L.text)} "\\9"`;
    const f = app.NUM_FMTS[L.fmt];
    const a = ph[0].index;
    const b = ph[ph.length - 1].index + 2;
    const num = ph.length === 1 ? `counter(list-item, ${f})` : `counters(list-item, ${cssStr(L.text.slice(a + 2, ph[1].index))}, ${f})`;
    return [a ? cssStr(L.text.slice(0, a)) : '', num, b < L.text.length ? cssStr(L.text.slice(b)) : '', '"\\9"'].filter(Boolean).join(' ');
  }
  function listFmtCSS(v) {
    const attr = `[data-lf=${cssStr(v)}]`;
    return app.parseLF(v).map((L, k) => {
      if (!L) return '';
      const tag = L.fmt === 'bullet' ? 'ul' : 'ol';
      const sel = k ? `.editor ${attr} ${':is(ol, ul) '.repeat(k - 1)}${tag}:not([data-lf])` : `.editor ${tag}${attr}`;
      return `${sel} > li::marker { content: ${markerContent(L)}; }`;
    }).filter(Boolean).join('\n');
  }
  const lfStyleEl = document.head.appendChild(document.createElement('style'));
  lfStyleEl.id = 'listStyles';
  const lfDone = new Set(); // CSS'i yazılmış biçimler (belgede kalmasa da durur)
  function applyListStyles(ed) {
    let add = false;
    ed.querySelectorAll('[data-lf]').forEach((el) => {
      // Dosyadan ya da panodan gelen değer doğrulanır; liste olmayan öğede anlamı yok
      const v = /^(OL|UL)$/.test(el.tagName) ? app.formatLF(app.parseLF(el.getAttribute('data-lf'))) : '';
      if (!v) return el.removeAttribute('data-lf');
      if (v !== el.getAttribute('data-lf')) el.setAttribute('data-lf', v);
      if (!lfDone.has(v)) {
        lfDone.add(v);
        add = true;
      }
    });
    if (add) lfStyleEl.textContent = [...lfDone].map(listFmtCSS).join('\n');
  }

  function renderPages() {
    const g = app.geom();
    const mk = (cls, css) => {
      const d = document.createElement('div');
      d.className = cls;
      if (css) d.style.cssText = css;
      return d;
    };
    const pf = document.createDocumentFragment();
    const bf = document.createDocumentFragment();
    const ff = document.createDocumentFragment();
    for (let k = 0; k < state.pageCount; k++) {
      const box = `top:${k * g.stride}px;width:${g.PW}px;height:${g.PH}px`;
      const page = mk('page' + (state.grid ? ' grid' : ''), box);
      page.dataset.page = k;
      page.appendChild(mk('mguide', `left:${g.m.l}px;top:${g.m.t}px;width:${g.cw}px;height:${g.ch}px`));
      // Üst/alt bilgi (kenar boşluğunda, metin genişliğinde; "İlk sayfada gösterme" seçiliyse kapakta yok)
      for (const band of k === 0 && state.hf.firstPage ? [] : ['header', 'footer']) {
        const slots = state.hf[band];
        if (!slots.some((s) => s.trim())) continue;
        const pos = band === 'header' ? `top:${app.HEADER_DIST}px` : `bottom:${app.FOOTER_DIST}px`;
        const el = mk('hf ' + band, `${pos};left:${g.m.l}px;width:${g.cw}px`);
        for (const s of slots) el.appendChild(document.createElement('span')).textContent = app.fillHF(s, k);
        page.appendChild(el);
      }
      pf.appendChild(page);
      bf.appendChild(mk('pclip', box));
      ff.appendChild(mk('pclip', box));
    }
    els.pages.replaceChildren(pf);
    els.behind.replaceChildren(bf);
    els.front.replaceChildren(ff);
    app.renderImages();
    updateDocSize();
  }
  app.renderPages = renderPages;

  const clipFor = (img) => (img.wrap === 'behind' ? els.behind : els.front).children[img.page];

  function placeImageEl(el, img) {
    el.style.cssText = `left:${img.x}px;top:${img.y}px;width:${img.w}px;height:${img.h}px;transform:rotate(${img.rot || 0}deg)`;
    el.classList.toggle('locked', !!img.locked);
    el.classList.toggle('cropped', !!img.crop);
    el.firstChild.style.cssText = img.crop ? app.cropStyle(img.crop) : '';
  }
  // Kırpılmış resim (img.crop = [sol, üst, sağ, alt], 0–1): tam resim çerçeveden taşar, çerçeve keser (objects.js: Kırpma)
  app.cropStyle = function ([l, t, r, b]) {
    const fw = 1 - l - r;
    const fh = 1 - t - b;
    const pc = (v) => SS.round(v * 100, 4) + '%';
    return `position:absolute;left:${pc(-l / fw)};top:${pc(-t / fh)};width:${pc(1 / fw)};height:${pc(1 / fh)}`;
  };

  // Şekil yazısı elemanı: resmin sayfa kırpma kutusunda, resmin hemen üstündeki katmanda
  function placeCaption(img, clip, nums) {
    let el = els.doc.querySelector(`.cap[data-id="${img.id}"]`);
    if (!img.caption || !clip) {
      if (el) el.remove();
      capH.delete(img.id);
      return;
    }
    if (!el) {
      el = document.createElement('div');
      el.className = 'cap';
      el.dataset.id = img.id;
      el.innerHTML = '<span class="cap-num"></span> <span class="cap-text"></span>';
    }
    if (el.parentElement !== clip || el.previousElementSibling?.dataset.id !== img.id) {
      const imgEl = clip.querySelector(`.img-obj[data-id="${img.id}"]`);
      if (imgEl) imgEl.after(el);
      else clip.appendChild(el);
    }
    const bb = app.aabb(img);
    el.style.width = Math.max(bb.w, app.CAP_MIN_W) + 'px';
    el.firstChild.textContent = `${img.caption.label || 'Şekil'} ${(nums.get(img.id) || { text: '1' }).text}.`;
    const t = el.lastChild;
    if (!t.isContentEditable && t.textContent !== img.caption.text) t.textContent = img.caption.text || '';
    capH.set(img.id, el.offsetHeight);
    const r = app.capRect(img);
    el.style.left = r.x + 'px';
    el.style.top = r.y + 'px';
  }
  app.measureCaption = (img) => placeCaption(img, clipFor(img), app.captionNumbers());

  // Metni etkileyen resim değişiklikleri (konum, boyut, kaydırma, yazı) o sayfaların sayfalamasını yeniler
  const imgSig = new Map(); // id -> { page, sig }
  function markImageChanges() {
    const seen = new Set();
    let lo = Infinity;
    let hi = -1;
    const touch = (p) => {
      lo = Math.min(lo, p);
      hi = Math.max(hi, p);
    };
    for (const img of state.images) {
      seen.add(img.id);
      const sig = [img.page, img.x, img.y, img.w, img.h, img.rot, img.wrap, img.caption && JSON.stringify(img.caption)].join('|');
      const old = imgSig.get(img.id);
      if (old && old.sig === sig) continue;
      if (old) touch(old.page);
      touch(img.page);
      imgSig.set(img.id, { page: img.page, sig });
    }
    for (const [id, v] of imgSig)
      if (!seen.has(id)) {
        touch(v.page);
        imgSig.delete(id);
      }
    if (hi >= 0 && app.markDirty) app.markDirty(lo - 1, hi + 1);
  }

  app.renderImages = function () {
    markImageChanges();
    const existing = new Map();
    els.doc.querySelectorAll('.img-obj').forEach((el) => existing.set(el.dataset.id, el));
    const nums = app.captionNumbers();
    for (const img of [...state.images].sort((a, b) => a.z - b.z)) {
      let el = existing.get(img.id);
      if (el) existing.delete(img.id);
      else {
        el = document.createElement('div');
        el.className = 'img-obj';
        el.dataset.id = img.id;
        const im = document.createElement('img');
        im.draggable = false;
        im.alt = '';
        el.appendChild(im);
      }
      if (el.dataset.asset !== img.asset) {
        el.dataset.asset = img.asset;
        el.firstChild.src = state.assets[img.asset]?.src || '';
      }
      placeImageEl(el, img);
      const clip = clipFor(img);
      if (clip) clip.appendChild(el);
      placeCaption(img, clip, nums);
    }
    existing.forEach((el) => el.remove());
    els.doc.querySelectorAll('.cap').forEach((c) => {
      const img = app.getImage(c.dataset.id);
      if (!img || !img.caption) c.remove();
    });
  };

  // Sürükleme sırasında tek resmi hızlıca güncelle
  app.updateImageEl = function (img) {
    const el = els.doc.querySelector(`.img-obj[data-id="${img.id}"]`);
    const clip = clipFor(img);
    if (!el || !clip) return app.renderImages();
    if (el.parentElement !== clip) clip.appendChild(el);
    placeImageEl(el, img);
    if (img.caption) placeCaption(img, clip, app.captionNumbers());
  };

  // ---------- Sayfa sonu ----------
  // <div class="pb"> yüksekliği, bulunduğu sayfanın metin alanı sonuna kadar uzatılır;
  // böylece arkasından gelen metin bir sonraki sayfanın başından başlar.
  function fitPageBreaks() {
    const pbs = els.editor.getElementsByClassName('pb');
    if (!pbs.length) return;
    const { stride, ch } = app.geom();
    for (const pb of pbs) {
      const top = app.clientToFlow(0, pb.getBoundingClientRect().top).y;
      const k = Math.floor((top + 0.5) / stride);
      const h = Math.max(0, k * stride + ch - top);
      if (Math.abs((parseFloat(pb.style.height) || 0) - h) > 0.5) pb.style.height = h + 'px';
    }
  }

  // ---------- Sayfalama kuralları (Word gibi) ----------
  // Dul/öksüz satır denetimi: paragrafın ilk satırı sayfa sonunda, son satırı sayfa başında tek başına
  // kalmaz. Başlıklar bölünmez ve sonraki paragrafla aynı sayfada kalır (sonrakiyle birlikte tut).
  // Kural gerektiğinde sayfanın metin alanı alttan kısaltılır (keep[k], computeBands). Word'e aktarılan
  // belgede de aynı kurallar açık (docx.js), böylece sayfa sonları örtüşür.
  // Sayfa sınırları sırayla, tarayıcının gerçek düzeni ölçülerek karara bağlanır (her değişiklik bir düzen
  // hesabı: uzun belgede pahalı). Bu yüzden yalnızca değişen sayfalardan başlanır ve bir sayfanın başı
  // öncekiyle aynı çıkınca durulur (sonrası değişmemiştir). Yazarken iş Word'deki gibi arka planda, küçük
  // parçalar hâlinde yapılır; dışa aktarma, yazdırma ve içe aktarma öncesinde app.paginateNow() tamamlar.
  const UNIT = 'p,h1,h2,h3,h4,h5,h6,li,div:not(.pb)';
  const pageKey = []; // pageKey[k]: k. sayfanın ilk satırı { u: paragraf, i: satır no } (erken bitiş için)
  const dirty = { from: Infinity, to: -1 }; // yeniden karara bağlanacak sayfa sınırları
  const isHeading = (u) => /^H[1-6]$/.test(u.tagName);

  // Satırları olan paragraf birimleri, belge sırasıyla (madde içindeki alt liste ayrı birimdir)
  function unitsOf() {
    return [...els.editor.querySelectorAll(UNIT)].filter(
      (u) => (u.tagName === 'LI' || !u.parentElement.closest('li,p,h1,h2,h3,h4,h5,h6')) && !(u.tagName === 'DIV' && u.querySelector(UNIT))
    );
  }
  // Birimin satırları (akış koordinatında metin kutuları; aynı satırdaki parçalar birleştirilir)
  function unitLines(u) {
    const r = document.createRange();
    r.selectNodeContents(u);
    const sub = u.tagName === 'LI' && u.querySelector(':scope > ul, :scope > ol');
    if (sub) r.setEndBefore(sub);
    const f = els.flow.getBoundingClientRect().top;
    const z = state.zoom;
    const rects = [...r.getClientRects()]
      .filter((x) => x.height > 0.5)
      .map((x) => ({ top: (x.top - f) / z, bottom: (x.bottom - f) / z }))
      .sort((a, b) => a.top - b.top);
    const lines = [];
    for (const x of rects) {
      const last = lines[lines.length - 1];
      if (last && x.top < last.bottom - 1) last.bottom = Math.max(last.bottom, x.bottom);
      else lines.push({ ...x });
    }
    if (!lines.length) {
      const b = u.getBoundingClientRect();
      lines.push({ top: (b.top - f) / z, bottom: (b.bottom - f) / z });
    }
    return lines;
  }
  // Paragraftan hemen sonra sayfa sonu (.pb) var mı?
  function breakAfter(u) {
    let t = u;
    while (t.parentElement !== els.editor) {
      if (t.nextElementSibling) return false;
      t = t.parentElement;
    }
    return !!(t.nextElementSibling && t.nextElementSibling.classList.contains('pb'));
  }

  // k. sayfa sınırında kural gerekiyorsa sonraki sayfaya itilecek ilk satırın üst kenarı (akış y), yoksa null
  function boundaryCut(k, units, L) {
    const { stride, ch } = app.geom();
    const top = k * stride;
    const bottom = top + ch;
    // Sayfa sonundan önce başlayan son birim
    let lo = 0;
    let hi = units.length - 1;
    let idx = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (L(units[mid])[0].top < bottom - 0.5) {
        idx = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    if (idx < 0) return null;
    const U = units[idx];
    const lines = L(U);
    const onK = lines.filter((l) => l.top >= top - 0.5 && l.top < bottom);
    const after = lines.filter((l) => l.top >= bottom);
    if (!onK.length) return null;
    const firstOnK = lines[0].top >= top - 0.5;
    let at = null; // U'nun itilecek ilk satırı
    if (after.length) {
      const nOn = onK.length;
      if (isHeading(U)) at = firstOnK ? 0 : null; // başlık bölünmez
      else if (firstOnK && nOn < 2) at = 0; // öksüz: ilk satır sayfa sonunda tek başına
      else if (after.length < 2) {
        // dul: son satır sayfa başında tek başına; bir satır daha götür (ilk satır yalnız kalacaksa hepsini)
        const need = 2 - after.length;
        if (firstOnK && nOn - need < 2) at = 0;
        else if (nOn - need >= 1) at = lines.indexOf(onK[nOn - need]);
      }
      if (at === null) return null;
      if (at > 0) return (lines[at - 1].bottom + lines[at].top) / 2;
    } else {
      // U sayfada bitiyor: başlıksa ve sonraki paragraf sonraki sayfada başlıyorsa başlığı da götür
      const N = units[idx + 1];
      if (!N || !isHeading(U) || !firstOnK || breakAfter(U) || L(N)[0].top < bottom) return null;
    }
    // Sonrakiyle birlikte tut: itilen paragrafın hemen önündeki başlık(lar) da gider
    let j = idx;
    while (j > 0 && isHeading(units[j - 1]) && !breakAfter(units[j - 1]) && L(units[j - 1])[0].top >= top - 0.5) j--;
    const first = L(units[j])[0];
    if (first.top <= top + 1) return null; // sayfa boşalırdı: kural uygulanamaz (Word de vazgeçer)
    const prev = j > 0 ? L(units[j - 1]) : null;
    const pb = prev && prev[prev.length - 1].bottom;
    return pb && pb > top && pb <= first.top ? (pb + first.top) / 2 : first.top - 1;
  }

  // p. sayfanın ilk satırı: { u, i }
  function pageStartKey(p, units, L) {
    const y = p * app.geom().stride - 0.5;
    let lo = 0;
    let hi = units.length - 1;
    let idx = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (L(units[mid])[0].top < y) {
        idx = mid;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    for (const j of [idx, idx + 1]) {
      const u = units[j];
      const i = u ? L(u).findIndex((l) => l.top >= y) : -1;
      if (i >= 0) return { u, i };
    }
    return null;
  }

  function fitPageCount() {
    const n = neededPages();
    if (n === state.pageCount) return;
    state.pageCount = n;
    keep.length = Math.min(keep.length, n - 1);
    renderPages();
    renderExclusions();
    fitPageBreaks();
  }

  // Bir sayfa sınırını karara bağla; sonraki sayfanın başı öncekiyle aynıysa true
  function paginateBoundary(k, units) {
    const { stride, ch } = app.geom();
    const had = keep[k] || 0;
    if (had) {
      keep[k] = 0; // kuralsız doğal düzene bak
      renderExclusions();
      fitPageBreaks();
    }
    let cache = new Map();
    const L = (u) => cache.get(u) || cache.set(u, unitLines(u)).get(u);
    const cut = boundaryCut(k, units, L);
    const K = cut === null ? 0 : k * stride + ch - cut;
    keep[k] = K > 0.5 ? K : 0;
    if (had || keep[k]) {
      renderExclusions();
      fitPageBreaks();
      cache = new Map();
    }
    fitPageCount();
    const key = pageStartKey(k + 1, units, L);
    const old = pageKey[k + 1];
    pageKey[k + 1] = key;
    return !!(key && old && key.u === old.u && key.i === old.i);
  }

  // Bekleyen sayfalama işi. upto: yalnızca bu sayfa sınırına kadar (içe aktarmada resim yerleştirme)
  let pagTimer = 0;
  let passGen = 0;
  function paginate(sync, upto = Infinity) {
    clearTimeout(pagTimer);
    const gen = ++passGen;
    let units = null;
    const step = () => {
      if (gen !== passGen || app.printMode) return;
      const t0 = performance.now();
      units = units || unitsOf();
      while (dirty.from < Math.min(state.pageCount - 1, upto)) {
        const k = dirty.from;
        const same = paginateBoundary(k, units);
        dirty.from = k + 1;
        if (same && k >= dirty.to) dirty.from = Infinity; // sonrası değişmedi
        if (!sync && performance.now() - t0 > 24 && dirty.from < state.pageCount - 1) {
          pagTimer = setTimeout(step, 0); // arayüz donmasın
          break;
        }
      }
      if (dirty.from >= state.pageCount - 1) {
        dirty.from = Infinity;
        dirty.to = -1;
        keep.length = Math.min(keep.length, Math.max(0, state.pageCount - 1));
        app.storeKeep();
      }
      updateDocSize();
      app.onLayout && app.onLayout();
    };
    step();
  }
  // from..to sayfalarının metni/resimleri değişti: sayfa sınırları arka planda yeniden karara bağlanır
  app.markDirty = function (from = 0, to = Infinity) {
    dirty.from = Math.min(dirty.from, Math.max(0, Math.floor(from)));
    dirty.to = Math.max(dirty.to, to);
    clearTimeout(pagTimer);
    pagTimer = setTimeout(() => paginate(false), 300);
  };
  // Seçimin (imlecin) bulunduğu sayfalar değişti. Sayfa, zaten düzen hesaplayan bir sonraki app.layout()
  // içinde ölçülür (girdi anında ölçmek her tuşta fazladan bir düzen hesabı demek)
  let selDirty = false;
  app.markDirtyAtSelection = () => (selDirty = true);
  function markSelectionPages() {
    const sel = window.getSelection();
    if (!sel.rangeCount || !els.editor.contains(sel.anchorNode)) return app.markDirty();
    const { stride } = app.geom();
    const f = els.flow.getBoundingClientRect().top;
    const r = sel.getRangeAt(0);
    const rects = r.getClientRects();
    const host = (n) => (n.nodeType === 1 ? n : n.parentElement).getBoundingClientRect();
    const y0 = rects.length ? rects[0].top : host(r.startContainer).top;
    const y1 = rects.length ? rects[rects.length - 1].bottom : host(r.endContainer).bottom;
    const page = (y) => Math.floor(Math.max(0, (y - f) / state.zoom) / stride);
    app.markDirty(page(y0) - 1, page(y1) + 1);
  }
  // Bekleyen sayfalamayı hemen bitir (dışa aktarma, yazdırma, içe aktarma)
  app.paginateNow = (upto) => dirty.from < Infinity && paginate(true, upto);
  app.paginationPending = () => dirty.from < Infinity;

  // ---------- Düzen ----------
  // Yalnızca <br> içeren boş satırlar tarayıcıda float şeritlerini yok sayar (sayfa arasına düşer).
  // "el" sınıfı, CSS ile görünmez bir boşluk ekleyerek bu satırların da şeritlerden kaçmasını sağlar.
  function markEmptyLines() {
    for (const el of els.editor.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,div:not(.pb)')) {
      const empty = !el.textContent && !el.querySelector('p,li,ul,ol,div');
      if (el.classList.contains('el') === empty) continue;
      el.classList.toggle('el', empty);
      if (!el.classList.length) el.removeAttribute('class'); // HTML'de class="" kalmasın
    }
  }

  // ---------- Blok yapısı ----------
  // Chrome bazı komutlarda listeyi ya da başlığı paragrafın İÇİNE koyar: <p><ol>…</ol><p>…</p></p>.
  // Bu HTML yeniden ayrıştırılınca (geri alma, dosya açma, otomatik kayıt) tarayıcı dış paragrafı kapatır,
  // boş paragraflar doğar ve metin kayar. İçteki blokları dışarı alıp paragrafı parçalara böleriz.
  const NESTED = ':is(p,h1,h2,h3,h4,h5,h6) > :is(p,h1,h2,h3,h4,h5,h6,ul,ol,div,blockquote,pre,table)';
  const isBlockEl = (n) => n.nodeType === 1 && /^(P|H[1-6]|UL|OL|DIV|BLOCKQUOTE|PRE|TABLE)$/.test(n.tagName);

  function splitBlock(E, runs) {
    const parent = E.parentNode;
    let run = null;
    for (const c of [...E.childNodes]) {
      if (isBlockEl(c)) {
        run = null;
        parent.insertBefore(c, E);
        if (/^(P|H[1-6])$/.test(c.tagName) && [...c.children].some(isBlockEl)) splitBlock(c, runs);
      } else {
        if (!run) runs.push((run = parent.insertBefore(E.cloneNode(false), E)));
        run.appendChild(c);
      }
    }
    E.remove();
  }

  // Seçim uçlarını, düğümler taşınınca da geçerli kalacak biçimde (metin düğümü ya da komşu düğüm) sakla
  function pinSelection() {
    const sel = window.getSelection();
    if (!sel.rangeCount || !els.editor.contains(sel.getRangeAt(0).commonAncestorContainer)) return null;
    const r = sel.getRangeAt(0);
    const pin = (n, o) => (n.nodeType === 3 ? { n, o } : n.childNodes[o] ? { before: n.childNodes[o] } : { after: n.lastChild || n });
    return [pin(r.startContainer, r.startOffset), pin(r.endContainer, r.endOffset)];
  }
  app.pinSelection = () => pinSelection();
  app.unpinSelection = (pins) => unpinSelection(pins);
  function unpinSelection(pins) {
    if (!pins) return;
    const r = document.createRange();
    const ok = pins.every((p, i) => {
      const n = p.n || p.before || p.after;
      if (!els.editor.contains(n) || n === els.editor) return false;
      const set = i ? ['setEnd', 'setEndBefore', 'setEndAfter'] : ['setStart', 'setStartBefore', 'setStartAfter'];
      if (p.n) r[set[0]](n, Math.min(p.o, n.length));
      else r[p.before ? set[1] : set[2]](n);
      return true;
    });
    if (!ok) return;
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }

  // ---------- Satır aralığı (Word'ün "satır" birimi) ----------
  // Paragrafın satır aralığı CSS line-height'ta durur:
  //  - birimsiz sayı = Word katı × satırı belirleyen yazı tipinin tek satır katsayısı (SS.lineFactor);
  //    ör. Calibri'de "1,5 satır" = 1,5 × 1,221 = 1,83
  //  - birimli (pt) değer = Word'ün "En az" aralığı.
  // Varsayılan (css: .editor) Word'ün Normal stili gibi 1,08 satırdır.
  const family1 = (f) => (f || '').split(',')[0].replace(/["']/g, '').trim() || 'Calibri';
  // Satır yüksekliğini belirleyen yazı tipi: paragraftaki en büyük puntolu metnin (eşitse en uzununun) yazı tipi
  app.lineFont = function (block) {
    const acc = new Map();
    let max = 0;
    const w = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(), i = 0; n && i < 80; n = w.nextNode()) {
      const len = n.nodeValue.trim().length;
      if (!len) continue;
      const sub = n.parentElement.closest('ul,ol');
      if (sub && block.contains(sub)) continue; // alt listenin metni ayrı paragraftır
      i++;
      const cs = getComputedStyle(n.parentElement);
      const size = parseFloat(cs.fontSize);
      if (size > max + 0.01) {
        max = size;
        acc.clear();
      }
      if (size > max - 0.01) acc.set(cs.fontFamily, (acc.get(cs.fontFamily) || 0) + len);
    }
    let best = getComputedStyle(block).fontFamily;
    let most = 0;
    acc.forEach((v, f) => v > most && ((most = v), (best = f)));
    return family1(best);
  };
  // { multiple: Word katı } ya da { px: "En az" yüksekliği }
  app.lineSpacing = function (block) {
    let v = '';
    for (let n = block; n && n !== els.editor && !v; n = n.parentElement) v = n.style.lineHeight;
    const cs = getComputedStyle(block);
    const px = parseFloat(cs.lineHeight);
    const fs = parseFloat(cs.fontSize);
    if (/[a-z%]$/i.test(v)) return { px: isFinite(px) ? px : fs * 1.2 };
    return { multiple: (isFinite(px) ? px / fs : 1.2) / SS.lineFactor(app.lineFont(block)) };
  };
  // Word katını bu paragraf için CSS line-height değerine çevir
  app.cssLineHeight = (block, multiple) => String(+(multiple * SS.lineFactor(app.lineFont(block))).toFixed(4));

  // Düzenleme sonrası: iç içe blokları düzelt. Yükleme sonrası (fromParse): ayrıştırıcının ürettiği
  // çocuksuz paragrafları da at (düzenleyici boş paragrafa her zaman <br> koyar; <p></p> yalnızca
  // bozuk iç içe yapının yeniden ayrıştırılmasından doğar).
  app.normalizeBlocks = function (fromParse) {
    const ed = els.editor;
    if (fromParse) ed.querySelectorAll('p:empty, h1:empty, h2:empty, h3:empty, h4:empty, h5:empty, h6:empty').forEach((e) => e.remove());
    // Liste maddesinin asılı girintisini (css: .editor li, -0,6351 cm) Chrome listeden çıkan paragrafa ya da
    // taşınan metne satır içi stil olarak kopyalar; ilk satır kenar boşluğunun dışına kayar. Satır içi
    // öğelerde girinti zaten anlamsız.
    ed.querySelectorAll('[style*="text-indent"]').forEach((e) => {
      if (/^(P|H[1-6]|DIV)$/.test(e.tagName) && !/^-(0\.6351cm|24\.003\d*px)$/.test(e.style.textIndent)) return;
      e.style.textIndent = '';
      if (!e.getAttribute('style')) e.removeAttribute('style');
    });
    // Seçim boşken simge komutu verilip yazılınca Chrome <span style="vertical-align: super|sub"> üretir:
    // gerçek <sup>/<sub> olsun (küçük boyut, Word'e vertAlign). Simge kapatılınca içeride kalan
    // "baseline" parçası simgenin dışına alınır.
    const vas = ed.querySelectorAll('span[style*="vertical-align"]');
    if (vas.length) {
      const pins = pinSelection();
      vas.forEach((s) => {
        const v = s.style.verticalAlign;
        s.style.verticalAlign = '';
        const rest = (s.getAttribute('style') || '').trim();
        if (v === 'super' || v === 'sub') {
          const el = document.createElement(v === 'super' ? 'sup' : 'sub');
          s.replaceWith(el);
          if (rest) el.appendChild(s);
          else while (s.firstChild) el.appendChild(s.firstChild);
        } else {
          const outer = s.parentElement.closest('sup, sub');
          if (outer && ed.contains(outer)) {
            // <sup>ön<span baseline>metin</span>arka</sup> → <sup>ön</sup>metin<sup>arka</sup>
            const tail = document.createRange();
            tail.setStartAfter(s);
            tail.setEndAfter(outer.lastChild);
            const after = outer.cloneNode(false);
            after.appendChild(tail.extractContents());
            outer.after(after);
            if (!after.textContent) after.remove();
            outer.after(s);
            if (!outer.textContent) outer.remove();
          }
          if (!rest) {
            while (s.firstChild) s.before(s.firstChild);
            s.remove();
          }
        }
      });
      unpinSelection(pins);
    }
    // Word'den gelen üst/alt simgede boyut <sup> içindeki span'daydı ve küçültmeyi eziyordu: boyutu dışarı al
    ed.querySelectorAll('sup > span[style*="font-size"]:only-child, sub > span[style*="font-size"]:only-child').forEach((s) => {
      const v = s.parentNode;
      v.replaceWith(s);
      while (s.firstChild) v.appendChild(s.firstChild);
      s.appendChild(v);
    });
    let changed = false;
    if (ed.querySelector(NESTED)) {
      const pins = pinSelection();
      const runs = [];
      let E;
      while ((E = ed.querySelector(NESTED)) && E.parentElement) splitBlock(E.parentElement, runs);
      runs.forEach((p) => !p.textContent.trim() && p.remove()); // bölünmeden kalan boş parçalar
      unpinSelection(pins);
      changed = true;
    }
    const hoisted = hoistPageBreaks(ed);
    applyListStyles(ed); // liste biçimlerinin CSS'i
    return wrapTabs(ed) || hoisted || changed;
  };

  // ---------- Özel sekme durakları (Word: w:tabs) ----------
  // Paragrafın data-tabs özniteliği: "l56.69 r604.72." gibi; tür (l sola, c ortaya, r sağa, d ondalık) + metin alanının
  // sol kenarından uzaklık (px; Word'deki gibi girintiden değil) + isteğe bağlı dolgu (. - _). Durak konan paragrafta
  // her sekme karakteri kendi span.tab'ında sıfır genişlikte durur (css); genişliği fitTabs bir sonraki durağa göre
  // padding-left olarak verir. Son özel duraktan sonra varsayılan 1,27 cm durakları, asılı girintide girinti de durak.
  const TAB_RE = /^([lcrd])(-?\d+(?:\.\d+)?)([._-]?)$/;
  app.parseTabs = (s) =>
    String(s || '').split(/\s+/).map((t) => TAB_RE.exec(t)).filter(Boolean)
      .map((m) => ({ type: m[1], pos: +m[2], leader: m[3] || '' }))
      .sort((a, b) => a.pos - b.pos);
  app.formatTabs = (list) =>
    list.slice().sort((a, b) => a.pos - b.pos).map((t) => t.type + +(+t.pos).toFixed(2) + (t.leader || '')).join(' ');
  // Blokta imlecin karakter uzaklığı (düğümler yeniden kurulurken imleç korunsun)
  function caretOffsetIn(b) {
    const sel = window.getSelection();
    if (!sel.rangeCount || !b.contains(sel.getRangeAt(0).startContainer)) return null;
    const pre = document.createRange();
    pre.setStart(b, 0);
    pre.setEnd(sel.getRangeAt(0).startContainer, sel.getRangeAt(0).startOffset);
    return pre.toString().length;
  }
  // Sekme span'ının sonuna düşen imleç ardındaki metne konur: yazılan metin span'ın içine girmesin
  function setCaretOffsetIn(b, off) {
    const w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
    const r = document.createRange();
    let last = null;
    let after = null;
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      last = t;
      if (off < t.length || (off === t.length && !t.parentElement.classList.contains('tab'))) break;
      if (off === t.length) {
        const next = w.nextNode();
        if (next) {
          last = next;
          off = 0;
        } else after = t.parentElement;
        break;
      }
      off -= t.length;
    }
    if (!last) return;
    if (after) r.setStartAfter(after);
    else r.setStart(last, Math.min(off, last.length));
    r.collapse(true);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }
  function wrapTabs(ed) {
    let changed = false;
    for (const b of ed.querySelectorAll('[data-tabs]')) {
      // Chrome'un sekme için açtığı <span style="white-space:pre"> sekme span'ı olur
      b.querySelectorAll('span[style*="white-space"]').forEach((s) => {
        if (s.textContent !== '\t' || s.classList.contains('tab')) return;
        s.removeAttribute('style');
        s.className = 'tab';
        changed = true;
      });
      // Sekme span'ına yazılmış metin (Chrome yazmayı span'ın içinde sürdürebilir): span açılır, sekme yeniden sarılır
      const bad = [...b.querySelectorAll('.tab')].filter((t) => t.textContent !== '\t');
      if (bad.length) {
        const off = caretOffsetIn(b);
        bad.forEach((t) => t.replaceWith(...t.childNodes));
        if (off !== null) setCaretOffsetIn(b, off);
        changed = true;
      }
      const texts = [];
      const w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
      for (let t = w.nextNode(); t; t = w.nextNode())
        if (t.nodeValue.includes('\t') && !(t.nodeValue === '\t' && t.parentElement.classList.contains('tab'))) texts.push(t);
      if (!texts.length) continue;
      const off = caretOffsetIn(b);
      for (const t of texts) {
        const frag = document.createDocumentFragment();
        for (const part of t.nodeValue.split(/(\t)/)) {
          if (!part) continue;
          if (part !== '\t') frag.appendChild(document.createTextNode(part));
          else frag.appendChild(Object.assign(document.createElement('span'), { className: 'tab', textContent: '\t' }));
        }
        t.replaceWith(frag);
      }
      if (off !== null) setCaretOffsetIn(b, off);
      changed = true;
    }
    return changed;
  }
  const DEFAULT_TAB = U.cmToPx(1.27);
  const TAB_LEADER = { '.': 'lead-dot', '-': 'lead-hyphen', _: 'lead-line' };
  const tabSig = new WeakMap(); // blok → en son hesaplandığı durum (değişmeyen paragraf yeniden ölçülmez)
  function fitTabs() {
    const ed = els.editor;
    ed.querySelectorAll('.tab[style]').forEach((t) => !t.closest('[data-tabs]') && t.removeAttribute('style')); // durağı kalkan paragraf
    const blocks = [...ed.querySelectorAll('[data-tabs]')].filter((b) => b.querySelector('.tab'));
    if (!blocks.length) return;
    const z = state.zoom;
    const { cw } = app.geom();
    const fr = els.flow.getBoundingClientRect();
    const fx = fr.left;
    const sigOf = (b) => {
      const r = b.getBoundingClientRect();
      // Yanındaki resim şeridi satırların başını kaydırır: şerit ya da paragrafın ona göre yeri değişince yeniden ölçülür
      const top = (r.top - fr.top) / z;
      const bottom = (r.bottom - fr.top) / z;
      const near = sideBands.filter((bd) => bd.a < bottom && bd.b > top).map((bd) => [bd.a - top, bd.b - top, bd.xL, bd.xR].map(Math.round).join(','));
      return [b.dataset.tabs, b.textContent, b.innerHTML.length, b.getAttribute('style'), Math.round(r.left), Math.round(r.width), z, near.join(';')].join('|');
    };
    for (const b of blocks) {
      if (tabSig.get(b) === sigOf(b)) continue;
      const stops = app.parseTabs(b.dataset.tabs);
      const cs = getComputedStyle(b);
      const indent = (b.getBoundingClientRect().left - fx) / z + (parseFloat(cs.paddingLeft) || 0);
      const hanging = parseFloat(cs.textIndent) < -0.5;
      // Soldan sağa: sonraki sekmelerin eski genişliği ölçülen satırı kırmasın diye önce hepsi sıfırlanır
      const tabs = [...b.querySelectorAll('.tab')];
      tabs.forEach((t) => (t.style.paddingLeft = '0px'));
      tabs.forEach((t, ti) => {
        const tr = t.getBoundingClientRect();
        const x = (tr.left - fx) / z;
        let stop = stops.find((s) => s.pos > x + 0.5);
        if (hanging && indent > x + 0.5 && (!stop || indent < stop.pos)) stop = { type: 'l', pos: indent, leader: '' };
        if (!stop) {
          // Sağ kenarın ötesinde durak yok: sekme ilerlemez, ardındaki metin alt satıra geçer
          const pos = (Math.floor((x + 0.5) / DEFAULT_TAB) + 1) * DEFAULT_TAB;
          stop = { type: 'l', pos: pos > cw + 0.5 ? x : pos, leader: '' };
        }
        let w = stop.pos - x;
        if (stop.type !== 'l') {
          // Sekmeden sonraki metin (sonraki sekmeye ya da satır sonuna kadar, aynı satırdaki kısmı)
          const seg = document.createRange();
          seg.setStartAfter(t);
          let end = null;
          const nextTab = tabs[ti + 1];
          if (nextTab) seg.setEndBefore(nextTab);
          else seg.setEnd(b, b.childNodes.length);
          if (stop.type === 'd') {
            const txt = seg.toString();
            // Word bölge ayarındaki ondalık ayırıcıya hizalar; Türkçede virgül. Virgül yoksa sayının sonu durakta
            // biter ("3.250": binlik ayırıcılı tam sayı), sayı da yoksa bütün parça (sağa hizalı gibi)
            const num = /\d(?:[\d.]*\d)?/.exec(txt);
            const i = txt.includes(',') ? txt.indexOf(',') : num ? num.index + num[0].length : -1;
            if (i >= 0) {
              end = document.createRange();
              end.setStart(seg.startContainer, seg.startOffset);
              // ondalık ayıracına kadar: metin düğümlerinde say
              let left = i;
              const w2 = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
              for (let tx = w2.nextNode(); tx; tx = w2.nextNode()) {
                if (!seg.intersectsNode(tx)) continue;
                const from = tx === seg.startContainer ? seg.startOffset : 0;
                if (left <= tx.length - from) {
                  end.setEnd(tx, from + left);
                  break;
                }
                left -= tx.length - from;
              }
            }
          }
          const rs = [...(end || seg).getClientRects()].filter((q) => q.width > 0 && q.top < tr.bottom && q.bottom > tr.top);
          const segW = rs.reduce((s, q) => s + q.width, 0) / z;
          w -= (stop.type === 'c' ? segW / 2 : segW) + 0.2; // küçük pay: yuvarlama satırı taşırıp kırmasın
        }
        t.style.paddingLeft = Math.max(0, +w.toFixed(2)) + 'px';
        t.className = 'tab' + (stop.leader ? ' ' + TAB_LEADER[stop.leader] : '');
      });
      tabSig.set(b, sigOf(b));
    }
  }

  // Sayfa sonu (.pb) her zaman editörün doğrudan çocuğudur. Liste maddesinde (Ctrl+Enter) ya da başka bir blokta
  // kalan sayfa sonu, içinde bulunduğu üst düzey bloğu böler: ardındaki içerik (listede kalan maddeler, numara
  // sürerek) sayfa sonundan sonra gelir. Liste içindeki sayfa sonu sayfalamayı durmadan büyütüyordu (sayfa sayısı
  // sınırsız artıyor, sekme donuyordu).
  function hoistPageBreaks(ed) {
    const stray = [...ed.querySelectorAll('.pb')].filter((pb) => pb.parentElement !== ed);
    if (!stray.length) return false;
    const pins = pinSelection();
    for (const pb of stray) {
      const cut = app.splitTop(pb.parentNode, Array.prototype.indexOf.call(pb.parentNode.childNodes, pb) + 1);
      cut.top.after(pb);
      if (cut.rest) pb.after(cut.rest);
      cut.dropEmptyTop();
    }
    unpinSelection(pins);
    return true;
  }

  // Konumun (container, offset) üst düzey bloğunu ikiye böler: konumdan sonrası bloğun sığ kopyasına (rest) taşınır.
  // Listede bölünen maddenin boş yarısı atılır, numaralı listede numara sürer (ol start). rest yerleştirilmez;
  // içerik kalmadıysa null. dropEmptyTop(): baştaki parça boş kaldıysa onu kaldırır.
  app.splitTop = function (container, offset) {
    const ed = els.editor;
    let top = container.nodeType === 1 ? container : container.parentNode;
    while (top.parentElement !== ed) top = top.parentElement;
    const tail = document.createRange();
    tail.setStart(container, offset);
    tail.setEnd(ed, Array.prototype.indexOf.call(ed.childNodes, top) + 1);
    let rest = tail.extractContents().firstElementChild;
    const hasText = (el) => !!el.textContent.trim() || !!el.querySelector('img, ul, ol');
    for (const part of [top, rest]) {
      if (!part || !/^(UL|OL)$/.test(part.tagName)) continue;
      part.querySelectorAll('li').forEach((li) => !hasText(li) && !li.querySelector('br') && li.remove());
    }
    if (rest && rest.tagName === 'OL' && top.tagName === 'OL')
      rest.setAttribute('start', (parseInt(top.getAttribute('start'), 10) || 1) + top.querySelectorAll(':scope > li').length);
    if (rest && !hasText(rest) && !rest.querySelector('li')) rest = null;
    return { top, rest, dropEmptyTop: () => !hasText(top) && !top.querySelector('li, br') && top.remove() };
  };

  app.layout = function () {
    markEmptyLines();
    renderExclusions();
    fitTabs();
    fitPageBreaks();
    for (let i = 0; i < 25; i++) {
      const n = neededPages();
      if (n === state.pageCount) break;
      state.pageCount = n;
      renderPages();
      renderExclusions();
      fitTabs();
      fitPageBreaks();
    }
    keep.length = Math.min(keep.length, Math.max(0, state.pageCount - 1));
    if (selDirty) {
      selDirty = false;
      markSelectionPages(); // düzen güncel: ölçüm ek maliyetsiz
    }
    refreshChapterCaptions();
    updateDocSize();
    app.onLayout && app.onLayout();
  };
  // Bölüm numaralı şekil yazıları metin değişince (başlık eklenir, silinir, resmin önüne ya da ardına geçer) yeniden
  // numaralanır
  function refreshChapterCaptions() {
    if (!state.images.some((i) => i.caption && i.caption.chapter)) return;
    const nums = app.captionNumbers();
    for (const img of state.images) {
      const el = img.caption && els.doc.querySelector(`.cap[data-id="${img.id}"] .cap-num`);
      if (!el || el.textContent === `${img.caption.label || 'Şekil'} ${nums.get(img.id).text}.`) continue;
      const h = capH.get(img.id);
      placeCaption(img, clipFor(img), nums);
      if (capH.get(img.id) !== h) {
        // Yazının yüksekliği değişti (satıra sığmadı): o sayfanın şeritleri yeniden kurulsun
        app.markDirty(img.page, img.page);
        app.scheduleLayout();
      }
    }
  }

  let layoutQueued = false;
  app.scheduleLayout = function () {
    if (layoutQueued) return;
    layoutQueued = true;
    requestAnimationFrame(() => {
      layoutQueued = false;
      app.layout();
    });
  };

  app.relayoutAll = function () {
    renderPages();
    app.layout();
    app.renderOverlay && app.renderOverlay();
  };

  // ---------- Metin seçimi: kaydet / geri yükle (geri alma için) ----------
  const nodeLen = (n) => (n.nodeType === 3 ? n.length : n.childNodes.length);
  function pathOf(node) {
    const path = [];
    while (node && node !== els.editor) {
      const p = node.parentNode;
      if (!p) return null;
      path.unshift(Array.prototype.indexOf.call(p.childNodes, node));
      node = p;
    }
    return node === els.editor ? path : null;
  }
  function nodeAt(path) {
    let n = els.editor;
    for (const i of path) {
      n = n.childNodes[i];
      if (!n) return null;
    }
    return n;
  }
  function saveSel() {
    const sel = window.getSelection();
    if (!sel.rangeCount) return null;
    const r = sel.getRangeAt(0);
    const a = pathOf(r.startContainer);
    const b = pathOf(r.endContainer);
    return a && b ? { a, ao: r.startOffset, b, bo: r.endOffset } : null;
  }
  function restoreSel(s) {
    if (!s) return;
    const a = nodeAt(s.a);
    const b = nodeAt(s.b);
    if (!a || !b) return;
    try {
      const r = document.createRange();
      r.setStart(a, Math.min(s.ao, nodeLen(a)));
      r.setEnd(b, Math.min(s.bo, nodeLen(b)));
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
    } catch (_) { /* yol artık geçersiz */ }
  }

  // ---------- Geri alma / yineleme (metin + resimler birlikte) ----------
  const hist = { stack: [], index: -1, kind: null, time: 0 };
  function snapshot() {
    return {
      html: els.editor.innerHTML,
      images: JSON.stringify(state.images),
      meta: JSON.stringify({ page: state.page, hf: state.hf, minPages: state.minPages, styles: state.styles }),
      sel: saveSel(),
      keep: keep.slice(), // sayfalama sonucu: geri alınınca düzen hemen doğru kurulsun
    };
  }
  // Sayfalama bitince güncel geçmiş adımının sonucunu da güncelle
  app.storeKeep = () => hist.stack[hist.index] && (hist.stack[hist.index].keep = keep.slice());
  function resetPagination(saved) {
    keep.length = 0;
    (saved || []).forEach((v, i) => (keep[i] = v));
    pageKey.length = 0;
    imgSig.clear();
  }

  // kind: 'typing' | 'deleting' | 'nudge' ardışık olanlar tek adımda birleşir; 'edit' birleşmez
  app.commit = function (kind = 'edit') {
    const s = snapshot();
    const now = Date.now();
    const top = hist.stack[hist.index];
    if (top && top.html === s.html && top.images === s.images && top.meta === s.meta) {
      top.sel = s.sel;
      return;
    }
    const merge =
      top && kind !== 'edit' && kind === hist.kind && now - hist.time < 1500 && hist.index === hist.stack.length - 1;
    if (merge) hist.stack[hist.index] = s;
    else {
      hist.stack.length = hist.index + 1;
      hist.stack.push(s);
      if (hist.stack.length > 300) hist.stack.shift();
      hist.index = hist.stack.length - 1;
    }
    hist.kind = kind;
    hist.time = now;
    app.onChange && app.onChange();
  };

  function restore(s) {
    hist.kind = null;
    els.editor.innerHTML = s.html;
    app.normalizeBlocks(true);
    state.images = JSON.parse(s.images);
    Object.assign(state, JSON.parse(s.meta));
    app.applyStyles();
    state.selection = state.selection.filter((id) => app.getImage(id));
    resetPagination(s.keep);
    app.relayoutAll();
    app.markDirty();
    restoreSel(s.sel);
    app.updateCtxBar && app.updateCtxBar();
    app.onChange && app.onChange();
  }

  app.undo = () => {
    if (hist.index > 0) restore(hist.stack[--hist.index]);
  };
  app.redo = () => {
    if (hist.index < hist.stack.length - 1) restore(hist.stack[++hist.index]);
  };
  app.canUndo = () => hist.index > 0;
  app.canRedo = () => hist.index < hist.stack.length - 1;
  app.resetHistory = () => {
    hist.stack = [];
    hist.index = -1;
    hist.kind = null;
    app.commit('init');
  };

  // ---------- Belge (dosya biçimi: .sayfa = JSON) ----------
  app.serialize = function () {
    const assets = {};
    for (const img of state.images) if (state.assets[img.asset]) assets[img.asset] = state.assets[img.asset];
    return {
      app: 'SerbestSayfa',
      version: 1,
      title: state.fileName,
      page: state.page,
      hf: state.hf,
      styles: state.styles,
      minPages: state.minPages,
      // yardımcı sınıf ve sekme genişlikleri düzen sırasında yeniden hesaplanır
      html: els.editor.innerHTML.replace(/ class="(el)?"/g, '').replace(/(<span class="tab[^"]*") style="[^"]*"/g, '$1'),
      images: state.images,
      assets,
      savedAt: new Date().toISOString(),
    };
  };

  app.load = function (data) {
    // Yapı, var olan belgeye dokunmadan önce denetlenir: bozuk dosya yarım yüklenip açık belgeyi bozmasın
    const obj = (v) => v == null || (typeof v === 'object' && !Array.isArray(v));
    const ok = data && data.app === 'SerbestSayfa' && (data.html == null || typeof data.html === 'string') &&
      (data.images == null || Array.isArray(data.images)) && obj(data.assets) && obj(data.page) && obj(data.hf) && obj(data.styles);
    if (!ok) throw new Error('Bu dosya bir SerbestSayfa belgesi değil ya da bozuk.');
    state.page = data.page || defaultPage();
    // Eski dosyalardaki "Sayfa numarası (alt orta)" seçeneği alt bilginin orta yuvasına çevrilir
    const hf = data.hf || emptyHF();
    if (!data.hf && data.pageNumbers) hf.footer[1] = '{sayfa}';
    state.hf = { header: [0, 1, 2].map((i) => String((hf.header || [])[i] || '')), footer: [0, 1, 2].map((i) => String((hf.footer || [])[i] || '')), firstPage: !!hf.firstPage };
    // Stil tanımı olmayan (eski) belgeler yerleşik stillerle yazılmıştı: kullanıcının varsayılanı değil, onlar
    state.styles = app.cleanStyles(data.styles);
    app.applyStyles();
    state.minPages = data.minPages || 1;
    state.images = data.images || [];
    state.assets = data.assets || {};
    state.fileName = data.title || 'Adsız belge';
    state.selection = [];
    state.pageCount = 1;
    els.editor.innerHTML = data.html || '<p><br></p>';
    app.normalizeBlocks(true);
    if (!els.editor.firstElementChild) els.editor.innerHTML = '<p><br></p>';
    resetPagination();
    app.relayoutAll();
    app.markDirty();
    app.resetHistory();
    app.updateCtxBar && app.updateCtxBar();
    app.onLoad && app.onLoad();
  };

  app.newDocument = function () {
    app.load({ app: 'SerbestSayfa', page: defaultPage(), styles: app.defaultStyles(), html: '<p><br></p>', images: [], assets: {} });
  };
})();
