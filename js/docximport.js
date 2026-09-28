/* Word (.docx) içe aktarma.
 * Metin, başlıklar (numaralarıyla), biçimler, listeler, satır aralıkları, sayfa ayarları ve resimler alınır.
 * - Satır içi resimler Word'deki satırlarına sabitlenir (metin üstünden/altından akar).
 * - Yüzen resimler Word'deki konum ve metin kaydırma ayarlarıyla gelir; kırpma ve döndürme korunur.
 * - Resmin hemen altındaki/üstündeki "Şekil 3. ..." paragrafı ya da gruptaki yazı kutusu, o resmin
 *   şekil yazısı olur.
 * Tablolar düz metne dönüşür (satır başına bir paragraf, hücreler sekmeyle ayrılır).
 */
(function () {
  'use strict';
  const SS = window.SS;
  const U = SS.units;
  const app = SS.app;

  const NS = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    v: 'urn:schemas-microsoft-com:vml',
  };
  const W = NS.w;
  const EMU = 9525; // 1 px (96 DPI)
  const TWIP_PER_CM = 566.929;

  // ---------- ZIP okuma (DecompressionStream ile, bağımlılıksız) ----------
  function readZip(buf) {
    const v = new DataView(buf);
    const u8 = new Uint8Array(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
      if (v.getUint32(i, true) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error('Bu dosya geçerli bir Word (.docx) dosyası değil.');
    const count = v.getUint16(eocd + 10, true);
    let p = v.getUint32(eocd + 16, true);
    const dec = new TextDecoder();
    const files = new Map();
    for (let i = 0; i < count && v.getUint32(p, true) === 0x02014b50; i++) {
      const method = v.getUint16(p + 10, true);
      const csize = v.getUint32(p + 20, true);
      const nlen = v.getUint16(p + 28, true);
      const xlen = v.getUint16(p + 30, true);
      const clen = v.getUint16(p + 32, true);
      const off = v.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      const start = off + 30 + v.getUint16(off + 26, true) + v.getUint16(off + 28, true);
      files.set(name, { method, data: u8.subarray(start, start + csize) });
      p += 46 + nlen + xlen + clen;
    }
    return files;
  }

  async function bytesOf(entry) {
    if (!entry) return null;
    if (entry.method === 0) return entry.data;
    if (entry.method !== 8) throw new Error('Desteklenmeyen sıkıştırma yöntemi.');
    const stream = new Blob([entry.data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function resolvePath(dir, target) {
    const parts = (target.startsWith('/') ? target.slice(1) : dir + target).split('/');
    const out = [];
    for (const p of parts) {
      if (p === '..') out.pop();
      else if (p && p !== '.') out.push(decodeURIComponent(p));
    }
    return out.join('/');
  }

  async function openPackage(file) {
    const buf = await file.arrayBuffer();
    const sig = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
    if (sig[0] === 0xd0 && sig[1] === 0xcf)
      throw new Error('Eski Word biçimi (.doc) açılamıyor. Word\'de "Farklı Kaydet → Word Belgesi (.docx)" ile kaydedip tekrar deneyin.');
    const files = readZip(buf);
    const lower = new Map([...files.keys()].map((k) => [k.toLowerCase(), k]));
    const entry = (path) => files.get(path) || files.get(lower.get(path.toLowerCase()));
    const xml = async (path) => {
      const b = await bytesOf(entry(path));
      return b ? new DOMParser().parseFromString(new TextDecoder().decode(b), 'application/xml') : null;
    };
    const relsOf = async (partPath) => {
      const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
      const doc = await xml(dir + '_rels/' + partPath.slice(dir.length) + '.rels');
      const map = new Map();
      if (doc)
        for (const r of doc.getElementsByTagName('Relationship')) {
          const external = r.getAttribute('TargetMode') === 'External';
          const target = r.getAttribute('Target') || '';
          map.set(r.getAttribute('Id'), { type: r.getAttribute('Type') || '', external, target: external ? target : resolvePath(dir, target) });
        }
      return map;
    };
    let docPath = 'word/document.xml';
    for (const r of (await relsOf('')).values()) if (/\/officeDocument$/.test(r.type)) docPath = r.target;
    return { entry, xml, relsOf, docPath };
  }

  // ---------- XML yardımcıları ----------
  const kids = (el, ns, name) => (el ? [...el.children].filter((c) => c.namespaceURI === ns && (!name || c.localName === name)) : []);
  const kid = (el, ns, name) => kids(el, ns, name)[0] || null;
  const all = (el, ns, name) => (el ? [...el.getElementsByTagNameNS(ns, name)] : []);
  const wa = (el, name) => (el ? el.getAttributeNS(W, name) : null);
  const onOff = (el) => {
    const v = wa(el, 'val');
    return v === null || !/^(0|false|off|none)$/i.test(v);
  };
  const num = (el, a) => +((el && el.getAttribute(a)) || 0);
  const textOf = (el) => all(el, W, 't').map((t) => t.textContent).join('');

  // ---------- Stil, tema, numaralandırma ----------
  function parseStyles(doc) {
    const map = new Map();
    let defPara = null;
    const dd = all(doc, W, 'docDefaults')[0];
    for (const s of all(doc, W, 'style')) {
      const id = wa(s, 'styleId');
      map.set(id, {
        id,
        name: (wa(kid(s, W, 'name'), 'val') || '').toLowerCase(),
        basedOn: wa(kid(s, W, 'basedOn'), 'val'),
        pPr: kid(s, W, 'pPr'),
        rPr: kid(s, W, 'rPr'),
      });
      if (wa(s, 'type') === 'paragraph' && wa(s, 'default') === '1') defPara = id;
    }
    const chain = (id) => {
      const out = [];
      const seen = new Set();
      while (id && map.has(id) && !seen.has(id)) {
        seen.add(id);
        out.unshift(map.get(id));
        id = map.get(id).basedOn;
      }
      return out;
    };
    return {
      defPara,
      defR: kid(kid(dd, W, 'rPrDefault'), W, 'rPr'),
      defP: kid(kid(dd, W, 'pPrDefault'), W, 'pPr'),
      chain,
      byName: (n) => [...map.values()].find((s) => s.name === n)?.id || null, // yerleşik adlar küçük harf İngilizce (Türkçe Word'de de)
    };
  }

  function parseTheme(doc) {
    const font = (tag) => {
      const latin = kid(all(doc, NS.a, tag)[0], NS.a, 'latin');
      return latin ? latin.getAttribute('typeface') : null;
    };
    return { minor: font('minorFont'), major: font('majorFont') };
  }

  const roman = (n) => {
    let s = '';
    for (const [v, r] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']])
      while (n >= v) {
        s += r;
        n -= v;
      }
    return s;
  };
  const letter = (n) => {
    let s = '';
    while (n > 0) {
      n--;
      s = String.fromCharCode(97 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  };
  function fmtNum(n, fmt) {
    if (fmt === 'lowerLetter') return letter(n);
    if (fmt === 'upperLetter') return letter(n).toUpperCase();
    if (fmt === 'lowerRoman') return roman(n);
    if (fmt === 'upperRoman') return roman(n).toUpperCase();
    return String(n);
  }

  // Word'ün liste sayaçlarını taklit eder: numaralı başlıkların ("1.2 Amaç") numarasını metne yazmak için
  function parseNumbering(doc) {
    const abs = new Map();
    const nums = new Map();
    for (const a of all(doc, W, 'abstractNum')) {
      const lv = new Map();
      for (const l of kids(a, W, 'lvl')) {
        const ind = kid(kid(l, W, 'pPr'), W, 'ind');
        const left = ind && (wa(ind, 'left') ?? wa(ind, 'start'));
        lv.set(+wa(l, 'ilvl'), {
          fmt: wa(kid(l, W, 'numFmt'), 'val') || 'decimal',
          text: wa(kid(l, W, 'lvlText'), 'val') || '',
          start: +(wa(kid(l, W, 'start'), 'val') || 1),
          left: left !== null && left !== undefined ? +left / 20 : null, // pt
        });
      }
      abs.set(wa(a, 'abstractNumId'), lv);
    }
    for (const n of all(doc, W, 'num')) {
      const starts = new Map();
      for (const o of kids(n, W, 'lvlOverride')) {
        const so = kid(o, W, 'startOverride');
        if (so) starts.set(+wa(o, 'ilvl'), +wa(so, 'val'));
      }
      nums.set(wa(n, 'numId'), { abs: wa(kid(n, W, 'abstractNumId'), 'val'), starts });
    }
    const counters = new Map(); // abstractNumId -> [sayaç...]
    const seenNum = new Set();
    return {
      fmt(numId, ilvl) {
        const lv = abs.get(nums.get(numId)?.abs);
        return lv ? (lv.get(ilvl) || lv.get(0) || { fmt: 'bullet' }).fmt : null;
      },
      // Düzeyin sol girintisi (pt) ve son next() çağrısındaki sayaç değeri
      left(numId, ilvl) {
        const lv = abs.get(nums.get(numId)?.abs);
        return lv && lv.get(ilvl) ? lv.get(ilvl).left : null;
      },
      count(numId, ilvl) {
        const c = counters.get(nums.get(numId)?.abs);
        return (c && c[ilvl]) || 1;
      },
      next(numId, ilvl) {
        const n = nums.get(numId);
        const lv = n && abs.get(n.abs);
        if (!lv) return '';
        if (!counters.has(n.abs)) counters.set(n.abs, []);
        const c = counters.get(n.abs);
        if (!seenNum.has(numId)) {
          seenNum.add(numId);
          n.starts.forEach((s, l) => (c[l] = s - 1));
        }
        const L = lv.get(ilvl) || { fmt: 'decimal', text: '', start: 1 };
        c[ilvl] = (c[ilvl] === undefined ? L.start - 1 : c[ilvl]) + 1;
        for (let i = ilvl + 1; i < 9; i++) c[i] = undefined; // alt seviyeler baştan başlar
        if (L.fmt === 'bullet' || L.fmt === 'none') return '';
        return L.text.replace(/%([1-9])/g, (_, d) => {
          const i = +d - 1;
          const li = lv.get(i) || { fmt: 'decimal', start: 1 };
          return fmtNum(c[i] === undefined ? li.start : c[i], li.fmt);
        });
      },
    };
  }

  // ---------- Biçim özellikleri ----------
  const HL = { yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
    darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
    darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000' };

  function applyRPr(rPr, o, theme) {
    if (!rPr) return o;
    for (const c of rPr.children) {
      if (c.namespaceURI !== W) continue;
      const v = wa(c, 'val');
      switch (c.localName) {
        case 'b': o.b = onOff(c); break;
        case 'i': o.i = onOff(c); break;
        case 'strike': case 'dstrike': o.s = onOff(c); break;
        case 'u': o.u = v !== 'none'; break;
        case 'caps': o.caps = onOff(c); break;
        case 'vanish': case 'webHidden': o.hidden = onOff(c); break;
        case 'color': o.color = v && v !== 'auto' ? '#' + v.toLowerCase() : null; break;
        case 'sz': if (v) o.sz = +v / 2; break;
        case 'highlight': o.hl = HL[v] || null; break;
        case 'shd': {
          const f = wa(c, 'fill');
          if (f && f !== 'auto') o.hl = /^f{6}$/i.test(f) ? null : '#' + f.toLowerCase();
          break;
        }
        case 'vertAlign': o.va = v === 'superscript' ? 'sup' : v === 'subscript' ? 'sub' : null; break;
        case 'rFonts': {
          const f = wa(c, 'ascii') || wa(c, 'hAnsi');
          const th = wa(c, 'asciiTheme') || wa(c, 'hAnsiTheme');
          if (f) o.font = f;
          else if (th) o.font = (/major/i.test(th) ? theme.major : theme.minor) || o.font;
          break;
        }
        default:
      }
    }
    return o;
  }

  function applyPPr(pPr, o) {
    if (!pPr) return o;
    for (const c of pPr.children) {
      if (c.namespaceURI !== W) continue;
      switch (c.localName) {
        case 'jc': o.jc = wa(c, 'val'); break;
        case 'spacing': {
          const b = wa(c, 'before');
          const a = wa(c, 'after');
          const l = wa(c, 'line');
          if (b !== null) o.before = +b / 20;
          if (a !== null) o.after = +a / 20;
          if (wa(c, 'beforeAutospacing') === '1') o.before = 14;
          if (wa(c, 'afterAutospacing') === '1') o.after = 14;
          if (l !== null) {
            o.line = +l;
            o.lineRule = wa(c, 'lineRule') || 'auto';
          }
          break;
        }
        case 'ind': {
          const l = wa(c, 'left') ?? wa(c, 'start');
          if (l !== null) o.left = +l / 20;
          const rt = wa(c, 'right') ?? wa(c, 'end');
          if (rt !== null) o.right = +rt / 20;
          const fl = wa(c, 'firstLine');
          const hg = wa(c, 'hanging');
          if (fl !== null) o.first = +fl / 20;
          if (hg !== null) o.first = -hg / 20;
          break;
        }
        case 'numPr': {
          const id = wa(kid(c, W, 'numId'), 'val');
          const lvl = wa(kid(c, W, 'ilvl'), 'val');
          if (id !== null) o.numId = id;
          if (lvl !== null) o.ilvl = +lvl;
          break;
        }
        case 'outlineLvl': o.outline = +wa(c, 'val'); break;
        case 'pageBreakBefore': o.pbBefore = onOff(c); break;
        default:
      }
    }
    return o;
  }

  // Word'ün stilleri → belgenin stilleri (core.js: state.styles). Normal = docDefaults + varsayılan paragraf stili;
  // Başlık 1–3 = "heading 1–3" stilleri (belgede yoksa yerleşik olan). "En az"/"Tam" satır aralıklı stil 1 satır
  // sayılır; o paragraflara aralık satır içi yazılır (paraCSS).
  function importStyles(ctx) {
    const def = app.builtinStyles();
    const JCMAP = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' };
    const conv = (id) => {
      const { pp, rb } = styleBase(ctx, id);
      return {
        font: rb.font,
        size: rb.sz,
        bold: !!rb.b,
        italic: !!rb.i,
        color: rb.color && /^#[0-9a-f]{6}$/.test(rb.color) ? rb.color : '#000000',
        align: JCMAP[pp.jc] || 'left',
        before: pp.before || 0,
        after: pp.after || 0,
        line: pp.line && (pp.lineRule || 'auto') === 'auto' ? pp.line / 240 : 1,
      };
    };
    const out = { p: conv(ctx.styles.defPara) };
    for (const [k, n] of [['h1', 'heading 1'], ['h2', 'heading 2'], ['h3', 'heading 3']]) {
      const id = ctx.styles.byName(n);
      out[k] = id ? conv(id) : def[k];
    }
    return app.cleanStyles(out);
  }
  // Belgenin stilleri (bunlardan farklı olanlar satır içi stil olarak yazılır); lh: stilin CSS satır yüksekliği
  let TAG = null;
  const tagFrom = (styles) =>
    Object.fromEntries(app.STYLE_KEYS.map((k) => {
      const s = styles[k];
      return [k, { font: s.font, sz: s.size, b: s.bold, i: s.italic, color: s.color, align: s.align, mt: s.before, mb: s.after, lh: app.styleLineHeight(s) }];
    }));

  const CAPTION_RE = /^\s*(Şekil|Sekil|Figure|Fig\.?|Harita|Fotoğraf|Foto|Resim|Grafik|Tablo|Table|Çizelge|Levha)\s*[-–]?\s*([0-9]+(?:[.\-–][0-9]+)*)\s*[.:\-–)]?\s*/i;
  const normLabel = (l) => {
    const s = l.replace(/\.$/, '');
    if (/^(sekil|figure|fig)$/i.test(s)) return 'Şekil';
    if (/^foto$/i.test(s)) return 'Fotoğraf';
    if (/^table$/i.test(s)) return 'Tablo';
    return s.charAt(0).toLocaleUpperCase('tr-TR') + s.slice(1).toLocaleLowerCase('tr-TR');
  };

  function styleBase(ctx, styleId) {
    if (ctx.cache.has(styleId)) return ctx.cache.get(styleId);
    const S = ctx.styles;
    const pp = applyPPr(S.defP, {});
    const rb = applyRPr(S.defR, { font: 'Times New Roman', sz: 10 }, ctx.theme);
    let heading = 0;
    let caption = false;
    for (const s of S.chain(styleId)) {
      applyPPr(s.pPr, pp);
      applyRPr(s.rPr, rb, ctx.theme);
      const m = /^heading ([1-9])$/.exec(s.name);
      if (m) heading = +m[1]; // "Title" (Konu Başlığı) başlık düzeyi değildir: biçimiyle Normal paragraf olur
      if (s.name === 'caption') caption = true;
    }
    if (!heading && pp.outline !== undefined && pp.outline < 9) heading = pp.outline + 1;
    const res = { pp, rb, heading, caption };
    ctx.cache.set(styleId, res);
    return res;
  }

  function paraInfo(ctx, p) {
    const pPr = kid(p, W, 'pPr');
    const styleId = wa(kid(pPr, W, 'pStyle'), 'val') || ctx.styles.defPara;
    const base = styleBase(ctx, styleId);
    const pp = applyPPr(pPr, { ...base.pp });
    let heading = base.heading;
    const ol = kid(pPr, W, 'outlineLvl');
    if (ol) heading = +wa(ol, 'val') < 9 ? +wa(ol, 'val') + 1 : 0;
    let list = null;
    let numText = '';
    if (pp.numId && pp.numId !== '0') {
      const ilvl = pp.ilvl || 0;
      numText = ctx.numbering.next(pp.numId, ilvl);
      const fmt = ctx.numbering.fmt(pp.numId, ilvl);
      // indent: görsel düzey için (renderBlocks); value: numaralı listenin bu maddedeki sayacı (start için)
      const lvLeft = ctx.numbering.left(pp.numId, ilvl);
      if (fmt && fmt !== 'none')
        list = { ordered: fmt !== 'bullet', ilvl, indent: pp.left !== undefined ? pp.left : lvLeft, value: ctx.numbering.count(pp.numId, ilvl) };
    }
    const tag = heading ? 'h' + Math.min(heading, 3) : 'p';
    return { pp, rb: { ...base.rb }, tag, list: heading ? null : list, numText: heading ? numText : '', caption: base.caption, sectPr: kid(pPr, W, 'sectPr') };
  }

  // lineFont: satır yüksekliğini belirleyen yazı tipi (paragraftaki en büyük puntolu metnin)
  function paraCSS(pp, tag, lineFont, inList) {
    const d = TAG[tag];
    const st = [];
    const jc = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[pp.jc] || 'left';
    if (jc !== d.align) st.push(`text-align: ${jc}`);
    if (!inList) {
      const mt = pp.before || 0;
      const mb = pp.after || 0;
      if (Math.abs(mt - d.mt) > 0.5) st.push(`margin-top: ${+mt.toFixed(1)}pt`);
      if (Math.abs(mb - d.mb) > 0.5) st.push(`margin-bottom: ${+mb.toFixed(1)}pt`);
      if (pp.left > 0.5) st.push(`margin-left: ${+pp.left.toFixed(1)}pt`);
      if (pp.right > 0.5) st.push(`margin-right: ${+pp.right.toFixed(1)}pt`);
      if (pp.first && Math.abs(pp.first) > 0.5) st.push(`text-indent: ${+pp.first.toFixed(1)}pt`);
    }
    // Satır aralığı hiçbir yerde belirtilmemişse Word'de tek satırdır
    if (!pp.line || pp.lineRule === 'auto') {
      const r = +(((pp.line || 240) / 240) * SS.lineFactor(lineFont)).toFixed(4);
      if (Math.abs(r - d.lh) > 0.002) st.push(`line-height: ${r}`); // küçük farklar sayfalarda birikir
    } else st.push(`line-height: ${+(pp.line / 20).toFixed(1)}pt`);
    return st.join('; ');
  }

  function lineFontOf(b, base) {
    const acc = new Map();
    let max = 0;
    for (const x of b.parts) {
      if (x.t !== 'text' || !x.text.trim()) continue;
      const sz = x.props.sz || base.sz;
      const f = x.props.font || base.font;
      if (sz > max + 0.01) {
        max = sz;
        acc.clear();
      }
      if (sz > max - 0.01) acc.set(f, (acc.get(f) || 0) + x.text.length);
    }
    let best = base.font;
    let most = 0;
    acc.forEach((v, f) => v > most && ((most = v), (best = f)));
    return best;
  }

  // ---------- Paragraf içeriği ----------
  function makePara() {
    const parts = [];
    return {
      parts,
      text(s, props) {
        if (!s || props.hidden) return;
        if (props.caps) s = s.toLocaleUpperCase('tr-TR');
        const key = JSON.stringify([props.font, props.sz, props.b, props.i, props.u, props.s, props.color, props.hl, props.va, props.href]);
        const last = parts[parts.length - 1];
        if (last && last.t === 'text' && last.key === key) last.text += s;
        else parts.push({ t: 'text', text: s, key, props });
      },
      br: () => parts.push({ t: 'br' }),
      ph: (id) => parts.push({ t: 'ph', id }),
      pb: () => parts.push({ t: 'pb' }),
    };
  }

  function altContent(ac, fn) {
    const choice = kid(ac, NS.mc, 'Choice');
    const fb = kid(ac, NS.mc, 'Fallback');
    const req = (choice && choice.getAttribute('Requires')) || '';
    if (choice && /^(wps|wpg|wp14|w14|w15|a14)\b/.test(req)) fn(choice);
    else if (fb) fn(fb);
    else if (choice) fn(choice);
  }

  function walkInline(parent, ctx, out, rb) {
    for (const n of parent.children) {
      if (n.namespaceURI === NS.mc && n.localName === 'AlternateContent') {
        altContent(n, (el) => walkInline(el, ctx, out, rb));
        continue;
      }
      if (n.namespaceURI !== W) continue;
      switch (n.localName) {
        case 'r': runContent(n, ctx, out, rb); break;
        case 'hyperlink': case 'fldSimple': {
          // Köprü: dış ilişkili w:hyperlink ya da HYPERLINK alanı (iç yer imleri düz metin kalır)
          const rel = n.localName === 'hyperlink' && ctx.rels.get(n.getAttributeNS(NS.r, 'id'));
          const fld = n.localName === 'fldSimple' && /HYPERLINK\s+"([^"]+)"/i.exec(wa(n, 'instr') || '');
          const prev = ctx.href;
          const href = rel && rel.external ? app.safeHref(rel.target) : fld ? app.safeHref(fld[1]) : null;
          if (href) ctx.href = href;
          walkInline(n, ctx, out, rb);
          ctx.href = prev;
          break;
        }
        case 'smartTag': case 'customXml': case 'ins': case 'moveTo': case 'dir': case 'bdo':
          walkInline(n, ctx, out, rb);
          break;
        case 'sdt': walkInline(kid(n, W, 'sdtContent') || n, ctx, out, rb); break;
        default: // pPr, bookmark, del, moveFrom, proofErr...
      }
    }
  }

  function runProps(ctx, r, rb, href) {
    const rPr = kid(r, W, 'rPr');
    const o = { ...rb };
    const rStyle = wa(kid(rPr, W, 'rStyle'), 'val');
    // Köprüdeki "Hyperlink" stilinin mavisi ve alt çizgisi düzenleyicide köprünün kendi görünümü: biçim olarak alınmaz
    const chain = rStyle ? ctx.styles.chain(rStyle) : [];
    if (!(href && chain.some((s) => /^(followed)?hyperlink$/.test(s.name)))) for (const s of chain) applyRPr(s.rPr, o, ctx.theme);
    return applyRPr(rPr, o, ctx.theme);
  }

  function runContent(r, ctx, out, rb) {
    const href = ctx.href || ctx.fldHref || null;
    const props = runProps(ctx, r, rb, href);
    props.href = href;
    const handle = (c) => {
      if (c.namespaceURI === NS.mc && c.localName === 'AlternateContent') return altContent(c, (el) => [...el.children].forEach(handle));
      if (c.namespaceURI !== W) return;
      switch (c.localName) {
        case 't': if (!ctx.inInstr) out.text(c.textContent, props); break;
        case 'tab': case 'ptab': out.text('\t', props); break;
        case 'br': {
          const t = wa(c, 'type');
          if (t === 'page') out.pb();
          else if (t !== 'column') out.br();
          break;
        }
        case 'cr': out.br(); break;
        case 'noBreakHyphen': out.text('-', props); break;
        case 'fldChar': {
          const t = wa(c, 'fldCharType');
          if (t === 'begin') {
            ctx.inInstr = true;
            ctx.instr = '';
          } else {
            ctx.inInstr = false; // separate / end: alan sonucu (görünen metin) okunur
            const m = t === 'separate' && /HYPERLINK\s+"([^"]+)"/i.exec(ctx.instr || '');
            if (m) ctx.fldHref = app.safeHref(m[1]);
            if (t === 'end') ctx.fldHref = null;
          }
          break;
        }
        case 'instrText': if (ctx.inInstr) ctx.instr += c.textContent; break;
        case 'drawing': drawing(c, ctx, out); break;
        case 'pict': case 'object': vml(c, ctx, out); break;
        case 'footnoteReference': case 'endnoteReference': ctx.report.notes++; break;
        default:
      }
    };
    [...r.children].forEach(handle);
  }

  // ---------- Çizimler: resim, metin kutusu, grup ----------
  function readPos(p) {
    if (!p) return null;
    const off = kid(p, NS.wp, 'posOffset');
    const al = kid(p, NS.wp, 'align');
    return { rel: p.getAttribute('relativeFrom'), off: off ? +off.textContent / EMU : null, align: al ? al.textContent.trim() : null };
  }

  function readAnchor(a) {
    const behind = a.getAttribute('behindDoc') === '1';
    const wrapEl = [...a.children].find((c) => c.namespaceURI === NS.wp && /^wrap/.test(c.localName));
    const wn = wrapEl ? wrapEl.localName : 'wrapNone';
    const res = {
      relHeight: num(a, 'relativeHeight'),
      wrap: wn === 'wrapTopAndBottom' ? 'topbottom' : /^wrap(Square|Tight|Through)$/.test(wn) ? 'square' : behind ? 'behind' : 'front',
      posH: readPos(kid(a, NS.wp, 'positionH')),
      posV: readPos(kid(a, NS.wp, 'positionV')),
    };
    if (a.getAttribute('simplePos') === '1') {
      const sp = kid(a, NS.wp, 'simplePos');
      res.posH = { rel: 'page', off: num(sp, 'x') / EMU };
      res.posV = { rel: 'page', off: num(sp, 'y') / EMU };
    }
    return res;
  }

  function drawing(d, ctx, out) {
    const box = kid(d, NS.wp, 'inline') || kid(d, NS.wp, 'anchor');
    if (!box) return;
    const ext = kid(box, NS.wp, 'extent');
    const w = num(ext, 'cx') / EMU;
    const h = num(ext, 'cy') / EMU;
    const rec = { inline: box.localName === 'inline', w, h, gw: w, gh: h, dx: 0, dy: 0, ph: 'ph' + ++ctx.seq };
    if (!rec.inline) Object.assign(rec, readAnchor(box));
    const gd = all(box, NS.a, 'graphicData')[0];
    const uri = (gd && gd.getAttribute('uri')) || '';
    const before = ctx.images.length;
    if (uri.endsWith('/picture')) picture(kid(gd, NS.pic, 'pic'), rec, ctx);
    else if (uri.endsWith('/wordprocessingShape')) shape(kid(gd, NS.wps, 'wsp'), rec, ctx);
    else if (uri.endsWith('/wordprocessingGroup')) group(kid(gd, NS.wpg, 'wgp'), rec, ctx, null);
    else if (/chart/.test(uri)) ctx.report.charts++;
    else ctx.report.other++;
    if (ctx.images.length > before) out.ph(rec.ph);
  }

  function picture(el, rec, ctx) {
    if (!el) return;
    const blip = all(el, NS.a, 'blip')[0];
    const rid = blip && (blip.getAttributeNS(NS.r, 'embed') || blip.getAttributeNS(NS.r, 'link'));
    const rel = rid && ctx.rels.get(rid);
    if (!rel || rel.external) {
      ctx.report.skipped++;
      return;
    }
    const sr = all(el, NS.a, 'srcRect')[0];
    const crop = sr ? ['l', 't', 'r', 'b'].map((k) => Math.max(0, num(sr, k) / 100000)) : null; // dışa genişletme (eksi) yok sayılır
    const xf = all(el, NS.a, 'xfrm')[0];
    ctx.images.push({
      ...rec,
      id: 'im' + ++ctx.seq,
      path: rel.target,
      crop: crop && crop.some((v) => v > 0.001) && crop[0] + crop[2] < 0.99 && crop[1] + crop[3] < 0.99 ? crop : null,
      rot: xf ? num(xf, 'rot') / 60000 : 0,
    });
    ctx.lastImage = ctx.images[ctx.images.length - 1];
  }

  function shape(wsp, rec, ctx) {
    if (!wsp) return;
    const tx = all(wsp, W, 'txbxContent')[0];
    if (tx) return textBox(tx, rec, ctx);
    if (all(wsp, NS.a, 'blip').length) return picture(wsp, rec, ctx); // resimle doldurulmuş şekil
    ctx.report.other++;
  }

  // Yazı kutusu: bir resmin şekil yazısıysa ona bağlanır, değilse paragrafları ana metne eklenir
  function textBox(tx, rec, ctx) {
    const text = kids(tx, W, 'p').map(textOf).join(' ').replace(/\s+/g, ' ').trim();
    const styled = all(tx, W, 'pStyle').some((s) => ctx.styles.chain(wa(s, 'val')).some((st) => st.name === 'caption'));
    const m = CAPTION_RE.exec(text);
    const target = ctx.lastImage;
    if (text && (m || styled) && target && !target.caption) {
      // Aynı gruptaysa grup içi konuma, ayrı yüzen nesnelerse dikey konumlarına bak
      const yBox = (rec.posV && rec.posV.off) || 0;
      const yImg = (target.posV && target.posV.off) || 0;
      const above = target.ph === rec.ph ? rec.dy + 1 < target.dy : !rec.inline && !target.inline && yBox < yImg;
      target.caption = {
        label: m ? normLabel(m[1]) : 'Şekil',
        text: m ? text.slice(m[0].length).trim() : text,
        pos: above ? 'above' : 'below',
      };
      ctx.report.captions++;
      return;
    }
    if (!text) return;
    ctx.report.textboxes++;
    ctx.pendingBoxes.push(tx);
  }

  // Grup: çocukların konumu grubun kendi koordinat sisteminden çerçeveye (px) çevrilir
  function group(g, rec, ctx, parent) {
    if (!g) return;
    const xf = all(kid(g, NS.wpg, 'grpSpPr'), NS.a, 'xfrm')[0];
    const off = kid(xf, NS.a, 'off');
    const ext = kid(xf, NS.a, 'ext');
    const chOff = kid(xf, NS.a, 'chOff');
    const chExt = kid(xf, NS.a, 'chExt');
    let T;
    if (!parent) {
      T = {
        sx: num(chExt, 'cx') ? rec.w / num(chExt, 'cx') : 1 / EMU,
        sy: num(chExt, 'cy') ? rec.h / num(chExt, 'cy') : 1 / EMU,
        cx: num(chOff, 'x'),
        cy: num(chOff, 'y'),
        bx: 0,
        by: 0,
      };
    } else {
      const kx = num(chExt, 'cx') ? num(ext, 'cx') / num(chExt, 'cx') : 1;
      const ky = num(chExt, 'cy') ? num(ext, 'cy') / num(chExt, 'cy') : 1;
      T = {
        sx: parent.sx * kx,
        sy: parent.sy * ky,
        cx: num(chOff, 'x'),
        cy: num(chOff, 'y'),
        bx: parent.bx + (num(off, 'x') - parent.cx) * parent.sx,
        by: parent.by + (num(off, 'y') - parent.cy) * parent.sy,
      };
    }
    for (const c of g.children) {
      const cxf = all(c, NS.a, 'xfrm')[0];
      const coff = kid(cxf, NS.a, 'off');
      const cext = kid(cxf, NS.a, 'ext');
      const crec = {
        ...rec,
        dx: T.bx + (num(coff, 'x') - T.cx) * T.sx,
        dy: T.by + (num(coff, 'y') - T.cy) * T.sy,
        w: num(cext, 'cx') * T.sx,
        h: num(cext, 'cy') * T.sy,
      };
      if (c.namespaceURI === NS.pic && c.localName === 'pic') picture(c, crec, ctx);
      else if (c.namespaceURI === NS.wps && c.localName === 'wsp') shape(c, crec, ctx);
      else if (c.namespaceURI === NS.wpg && c.localName === 'grpSp') group(c, rec, ctx, T);
    }
  }

  const cssLen = (style, prop) => {
    const m = new RegExp('(?:^|;)\\s*' + prop + '\\s*:\\s*([\\d.]+)(pt|px|in|cm|mm)?', 'i').exec(style);
    return m ? +m[1] * { pt: 4 / 3, px: 1, in: 96, cm: 37.795, mm: 3.7795 }[(m[2] || 'px').toLowerCase()] : 0;
  };

  // Eski (VML) çizimler
  function vml(el, ctx, out) {
    const im = all(el, NS.v, 'imagedata')[0];
    const tx = all(el, W, 'txbxContent')[0];
    if (im) {
      const rid = im.getAttributeNS(NS.r, 'id') || im.getAttribute('r:id');
      const rel = rid && ctx.rels.get(rid);
      if (!rel || rel.external) return void ctx.report.skipped++;
      const st = (im.parentElement && im.parentElement.getAttribute('style')) || '';
      const w = cssLen(st, 'width') || 200;
      const h = cssLen(st, 'height') || 150;
      const ph = 'ph' + ++ctx.seq;
      ctx.images.push({ id: 'im' + ++ctx.seq, ph, inline: true, w, h, gw: w, gh: h, dx: 0, dy: 0, path: rel.target, crop: null, rot: 0 });
      ctx.lastImage = ctx.images[ctx.images.length - 1];
      out.ph(ph);
    } else if (tx) textBox(tx, { inline: true, dx: 0, dy: 0 }, ctx);
    else ctx.report.other++;
  }

  // ---------- Gövde ----------
  function paragraph(p, ctx) {
    const info = paraInfo(ctx, p);
    const out = makePara();
    ctx.lastImage = null;
    ctx.inInstr = false;
    ctx.fldHref = null;
    const boxes = (ctx.pendingBoxes = []);
    if (info.numText) out.text(info.numText + ' ', info.rb);
    walkInline(p, ctx, out, info.rb);
    const res = [];
    if (info.pp.pbBefore) res.push({ type: 'pb' });
    const mk = (parts) => ({ type: 'p', tag: info.tag, list: info.list, rb: info.rb, pp: info.pp, parts, caption: info.caption });
    let cur = [];
    const hasContent = (parts) => parts.some((x) => x.t !== 'br');
    for (const part of out.parts) {
      if (part.t !== 'pb') {
        cur.push(part);
        continue;
      }
      if (hasContent(cur)) res.push(mk(cur));
      res.push({ type: 'pb' });
      cur = [];
    }
    if (hasContent(cur) || !out.parts.some((x) => x.t === 'pb')) res.push(mk(cur));
    // Bölüm sonu (sonraki sayfadan başlayan) -> sayfa sonu
    if (info.sectPr && !/^continuous$/i.test(wa(kid(info.sectPr, W, 'type'), 'val') || '')) res.push({ type: 'pb' });
    for (const tx of boxes) for (const bp of kids(tx, W, 'p')) res.push(...paragraph(bp, ctx));
    return res;
  }

  function table(tbl, ctx) {
    ctx.report.tables++;
    const blocks = [];
    for (const tr of kids(tbl, W, 'tr')) {
      const out = makePara();
      let rb = null;
      kids(tr, W, 'tc').forEach((tc, ci) => {
        let firstP = true;
        for (const p of all(tc, W, 'p')) {
          const info = paraInfo(ctx, p);
          rb = rb || info.rb;
          if (ci > 0 && firstP) out.text('\t', info.rb);
          else if (!firstP) out.text(' ', info.rb);
          firstP = false;
          ctx.lastImage = null;
          ctx.inInstr = false;
          ctx.pendingBoxes = [];
          walkInline(p, ctx, out, info.rb);
        }
      });
      blocks.push({ type: 'p', tag: 'p', list: null, rb: rb || { font: 'Calibri', sz: 11 }, pp: {}, parts: out.parts.filter((x) => x.t !== 'pb') });
    }
    return blocks;
  }

  function walkBody(parent, ctx) {
    const blocks = [];
    for (const n of parent.children) {
      if (n.namespaceURI === NS.mc && n.localName === 'AlternateContent') {
        altContent(n, (el) => blocks.push(...walkBody(el, ctx)));
        continue;
      }
      if (n.namespaceURI !== W) continue;
      if (n.localName === 'p') blocks.push(...paragraph(n, ctx));
      else if (n.localName === 'tbl') blocks.push(...table(n, ctx));
      else if (n.localName === 'sdt') blocks.push(...walkBody(kid(n, W, 'sdtContent') || n, ctx));
      else if (n.localName === 'customXml') blocks.push(...walkBody(n, ctx));
    }
    return blocks;
  }

  // Resimden oluşan paragrafın hemen altındaki/üstündeki "Şekil N. ..." paragrafını şekil yazısı yap
  function attachCaptions(blocks, ctx) {
    const isImageOnly = (b) =>
      b && b.type === 'p' && b.parts.some((x) => x.t === 'ph') && b.parts.every((x) => x.t === 'ph' || x.t === 'br' || (x.t === 'text' && !x.text.trim()));
    const capInfo = (b) => {
      if (!b || b.type !== 'p' || b.list || b.used || b.parts.some((x) => x.t === 'ph')) return null;
      const text = b.parts.filter((x) => x.t === 'text').map((x) => x.text).join('').replace(/\s+/g, ' ').trim();
      const m = CAPTION_RE.exec(text);
      if (m) return { label: normLabel(m[1]), text: text.slice(m[0].length).trim() };
      return b.caption && text ? { label: 'Şekil', text } : null;
    };
    const lastImg = (b) => {
      const phs = b.parts.filter((x) => x.t === 'ph').map((x) => x.id);
      for (let i = phs.length - 1; i >= 0; i--) {
        const rec = ctx.images.filter((r) => r.ph === phs[i]).pop();
        if (rec && !rec.caption) return rec;
      }
      return null;
    };
    for (let i = 0; i < blocks.length; i++) {
      if (!isImageOnly(blocks[i])) continue;
      for (const [j, pos] of [[i + 1, 'below'], [i - 1, 'above']]) {
        const c = capInfo(blocks[j]);
        const rec = c && lastImg(blocks[i]);
        if (!rec) continue;
        rec.caption = { ...c, pos };
        blocks[j].used = true;
        ctx.report.captions++;
        break;
      }
    }
    return blocks.filter((b) => !b.used);
  }

  // ---------- HTML üretimi ----------
  const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

  function runHTML(text, p, base) {
    let h = esc(text);
    // Üst/alt simge en içte: dıştaki punto span'ı <sup>'nun küçültmesini ezmesin (css: .editor sup)
    if (p.va) h = `<${p.va}>${h}</${p.va}>`;
    const st = [];
    if (p.font && p.font.toLowerCase() !== base.font.toLowerCase()) st.push(`font-family: '${p.font.replace(/['"]/g, '')}', Calibri, sans-serif`);
    if (p.sz && Math.abs(p.sz - base.sz) > 0.01) st.push(`font-size: ${p.sz}pt`);
    // Renk, kalınlık ve italik paragrafın stiline göre (stil renkliyse "otomatik" metin siyah olarak yazılır)
    if ((p.color || '#000000') !== (base.color || '#000000')) st.push(`color: ${p.color || '#000000'}`);
    if (p.hl) st.push(`background-color: ${p.hl}`);
    if (!p.b && base.b) st.push('font-weight: normal');
    if (!p.i && base.i) st.push('font-style: normal');
    if (st.length) h = `<span style="${st.join('; ')}">${h}</span>`;
    if (p.b && !base.b) h = `<b>${h}</b>`;
    if (p.i && !base.i) h = `<i>${h}</i>`;
    if (p.u) h = `<u>${h}</u>`;
    if (p.s) h = `<s>${h}</s>`;
    if (p.href) h = `<a href="${esc(p.href)}">${h}</a>`;
    return h;
  }

  function renderBlocks(blocks) {
    // Görsel liste düzeyi = art arda gelen liste paragraflarında girintinin sırası. Word'ün "Liste Madde
    // İşareti 2" gibi tek düzeyli ama daha içeride duran listeleri de böylece alt düzey olur; girinti
    // bilinmiyorsa ilvl kullanılır.
    const key = (b) => Math.round(b.list.indent ?? b.list.ilvl * 36) + b.list.ilvl / 100;
    let run = [];
    const flush = () => {
      const keys = [...new Set(run.map(key))].sort((a, b) => a - b);
      run.forEach((b) => (b.list.level = keys.indexOf(key(b))));
      run = [];
    };
    for (const b of blocks) {
      if (b.type === 'p' && b.list) run.push(b);
      else flush();
    }
    flush();

    const root = document.createElement('div');
    let stack = [];
    const fill = (el, b) => {
      const tagBase = TAG[b.tag] || TAG.p;
      const rb = b.rb || tagBase;
      // Paragraf stilinin yazı tipi ve boyutu bloğun kendisine yazılır (Word'deki paragraf işareti gibi):
      // aynı biçimdeki metin span'sız kalır, boş satırlar ve liste işaretleri de doğru boyutta olur
      const base = { ...tagBase };
      const st = [];
      if (rb.font && rb.font.toLowerCase() !== tagBase.font.toLowerCase()) {
        st.push(`font-family: '${rb.font.replace(/['"]/g, '')}', Calibri, sans-serif`);
        base.font = rb.font;
      }
      if (rb.sz && Math.abs(rb.sz - tagBase.sz) > 0.01) {
        st.push(`font-size: ${rb.sz}pt`);
        base.sz = rb.sz;
      }
      let html = '';
      for (const x of b.parts) {
        if (x.t === 'text') html += runHTML(x.text, x.props, base);
        else if (x.t === 'br') html += '<br>';
        else if (x.t === 'ph') html += `<span class="ph" data-ph="${x.id}"></span>`;
      }
      el.innerHTML = html || '<br>';
      const css = paraCSS(b.pp || {}, b.tag, lineFontOf(b, base), !!b.list);
      if (css) st.push(css);
      if (st.length) el.setAttribute('style', st.join('; '));
    };
    for (const b of blocks) {
      if (b.type === 'pb') {
        stack = [];
        const pb = document.createElement('div');
        pb.className = 'pb';
        pb.contentEditable = 'false';
        root.appendChild(pb);
        continue;
      }
      if (!b.list) {
        stack = [];
        const el = document.createElement(b.tag);
        fill(el, b);
        root.appendChild(el);
        continue;
      }
      const lvl = Math.min(b.list.level ?? b.list.ilvl, 8);
      while (stack.length > lvl + 1) stack.pop();
      if (stack.length === lvl + 1 && stack[lvl].ordered !== b.list.ordered) stack.pop();
      while (stack.length < lvl + 1) {
        const L = document.createElement(b.list.ordered ? 'ol' : 'ul');
        // Araya paragraf girip devam eden Word listesi kaldığı numaradan sürer
        if (b.list.ordered && stack.length === lvl && b.list.value > 1) L.setAttribute('start', b.list.value);
        const top = stack[stack.length - 1];
        (top ? top.lastLi || top.el : root).appendChild(L);
        stack.push({ el: L, ordered: b.list.ordered, lastLi: null });
      }
      const li = document.createElement('li');
      fill(li, b);
      stack[stack.length - 1].el.appendChild(li);
      stack[stack.length - 1].lastLi = li;
    }
    return root.innerHTML;
  }

  // ---------- Sayfa ayarları ----------
  function pageFromSect(sect) {
    const pgSz = kid(sect, W, 'pgSz');
    const pgMar = kid(sect, W, 'pgMar');
    const tw = (v, d) => (v === null || v === '' || isNaN(+v) ? d : +v);
    const wcm = tw(wa(pgSz, 'w'), 11906) / TWIP_PER_CM;
    const hcm = tw(wa(pgSz, 'h'), 16838) / TWIP_PER_CM;
    const orient = wcm > hcm ? 'landscape' : 'portrait';
    const [a, b] = orient === 'landscape' ? [hcm, wcm] : [wcm, hcm];
    let size = 'Custom';
    for (const [k, [sw, sh]] of Object.entries(SS.PAGE_SIZES)) if (Math.abs(sw - a) < 0.3 && Math.abs(sh - b) < 0.3) size = k;
    const m = (k) => U.cmToPx(Math.abs(tw(wa(pgMar, k), 1417)) / TWIP_PER_CM);
    return { size, orient, custom: [+a.toFixed(2), +b.toFixed(2)], margins: { t: m('top'), r: m('right'), b: m('bottom'), l: m('left') } };
  }

  // ---------- Üst/alt bilgi ----------
  // Varsayılan üst/alt bilginin ilk dolu paragrafı sol/orta/sağ yuvalara bölünür (sekmeye ya da hizalamaya
  // göre); PAGE / NUMPAGES alanları {sayfa} / {toplam} olur. Alınamayanlar (resim, tablo, ek paragraf,
  // ilk sayfaya özel içerik) sayılır ve açılışta bildirilir.
  const pageToken = (instr) => (/\bPAGE\b/.test(instr) ? '{sayfa}' : /\b(NUMPAGES|SECTIONPAGES)\b/.test(instr) ? '{toplam}' : null);

  function hfSlots(x) {
    const root = x.documentElement;
    let slots = null;
    let lost = kids(root, W, 'tbl').length;
    for (const p of kids(root, W, 'p')) {
      const segs = [''];
      let fld = null; // karmaşık alan: { instr, sep, tok }
      const put = (s) => (segs[segs.length - 1] += s);
      const run = (r) => {
        for (const c of r.children) {
          if (c.namespaceURI !== W) {
            if (c.localName === 'AlternateContent') lost++;
            continue;
          }
          switch (c.localName) {
            case 't': if (!fld || (fld.sep && !fld.tok)) put(c.textContent); break;
            case 'tab': segs.push(''); break;
            case 'ptab': {
              const at = { left: 0, center: 1, right: 2 }[wa(c, 'alignment')] ?? segs.length;
              while (segs.length - 1 < at) segs.push('');
              break;
            }
            case 'instrText': if (fld) fld.instr += c.textContent; break;
            case 'fldChar': {
              const t = wa(c, 'fldCharType');
              if (t === 'begin') fld = { instr: '', sep: false, tok: null };
              else if (t === 'separate' && fld) {
                fld.sep = true;
                fld.tok = pageToken(fld.instr);
                if (fld.tok) put(fld.tok);
              } else if (t === 'end') {
                if (fld && !fld.sep && pageToken(fld.instr)) put(pageToken(fld.instr));
                fld = null;
              }
              break;
            }
            case 'drawing': case 'pict': case 'object': lost++; break;
            default:
          }
        }
      };
      const walk = (el) => {
        for (const c of el.children) {
          if (c.namespaceURI === NS.mc && c.localName === 'AlternateContent') lost++;
          if (c.namespaceURI !== W) continue;
          if (c.localName === 'r') run(c);
          else if (c.localName === 'fldSimple') {
            const t = pageToken(wa(c, 'instr') || '');
            if (t) put(t);
            else walk(c);
          } else if (/^(hyperlink|smartTag|ins|customXml)$/.test(c.localName)) walk(c);
          else if (c.localName === 'sdt') walk(kid(c, W, 'sdtContent') || c);
        }
      };
      walk(p);
      if (!segs.join('').trim()) continue;
      if (slots) {
        lost++; // yalnızca ilk dolu paragraf alınır
        continue;
      }
      if (segs.length === 1) {
        const jc = wa(kid(kid(p, W, 'pPr'), W, 'jc'), 'val') || '';
        slots = ['', '', ''];
        slots[jc === 'center' ? 1 : /right|end/.test(jc) ? 2 : 0] = segs[0];
      } else slots = [segs[0], segs[1], segs.slice(2).join(' ')];
    }
    return { slots: slots && slots.map((s) => s.replace(/\s+/g, ' ').trim()), lost };
  }

  async function readHF(sect, ctx) {
    const hf = app.emptyHF();
    const tp = kid(sect, W, 'titlePg');
    hf.firstPage = !!tp && onOff(tp);
    let lost = 0;
    const part = async (tag, type) => {
      const ref = kids(sect, W, tag).find((r) => (wa(r, 'type') || 'default') === type);
      const rel = ref && ctx.rels.get(ref.getAttributeNS(NS.r, 'id'));
      return rel ? ctx.pkg.xml(rel.target) : null;
    };
    for (const [tag, band] of [['headerReference', 'header'], ['footerReference', 'footer']]) {
      const x = await part(tag, 'default');
      if (x) {
        const r = hfSlots(x);
        if (r.slots) hf[band] = r.slots;
        lost += r.lost;
      }
      // İlk sayfaya özel içerik alınamaz: kapakta üst/alt bilgi gösterilmez
      const first = hf.firstPage && (await part(tag, 'first'));
      if (first && hfSlots(first).slots) lost++;
    }
    return { hf, lost };
  }

  // ---------- Resimler ----------
  const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };

  // Word'ün kırpması resmi değiştirmez: tam resim alınır, kırpma img.crop olarak kalır (değiştirilebilir, sıfırlanabilir)
  async function loadAssets(ctx) {
    const assets = {};
    const byKey = new Map();
    for (const r of ctx.images) {
      const key = r.path;
      if (byKey.has(key)) {
        r.asset = byKey.get(key);
        continue;
      }
      const mime = MIME[r.path.split('.').pop().toLowerCase()];
      const e = ctx.pkg.entry(r.path);
      if (!mime || !e) {
        ctx.report.unsupported++; // EMF/WMF/TIFF gibi tarayıcının gösteremediği biçimler
        byKey.set(key, null);
        continue;
      }
      try {
        const a = await SS.prepareImage(new Blob([await bytesOf(e)], { type: mime }));
        const id = SS.uid('a');
        assets[id] = a;
        byKey.set(key, id);
        r.asset = id;
      } catch (_) {
        ctx.report.unsupported++;
        byKey.set(key, null);
      }
    }
    return assets;
  }

  function hPos(r, g, m) {
    const p = r.posH || { rel: 'column', off: 0 };
    let x0 = g.m.l;
    let x1 = g.PW - g.m.r;
    if (p.rel === 'page') [x0, x1] = [0, g.PW];
    else if (p.rel === 'leftMargin' || p.rel === 'insideMargin') [x0, x1] = [0, g.m.l];
    else if (p.rel === 'rightMargin' || p.rel === 'outsideMargin') [x0, x1] = [g.PW - g.m.r, g.PW];
    else if (p.rel === 'character') x0 = x1 = m.phX;
    const w = r.gw || r.w;
    if (p.align) return p.align === 'center' ? (x0 + x1 - w) / 2 : /right|outside/.test(p.align) ? x1 - w : x0;
    return x0 + (p.off || 0);
  }

  function vPos(r, g, m) {
    const p = r.posV || { rel: 'paragraph', off: 0 };
    let page = m.paraPage;
    let y0 = m.paraY;
    let y1 = m.paraY;
    if (p.rel === 'page') [y0, y1] = [0, g.PH];
    else if (p.rel === 'margin') [y0, y1] = [g.m.t, g.PH - g.m.b];
    else if (p.rel === 'topMargin') [y0, y1] = [0, g.m.t];
    else if (/bottomMargin|insideMargin|outsideMargin/.test(p.rel)) [y0, y1] = [g.PH - g.m.b, g.PH];
    else if (p.rel === 'line') {
      page = m.page;
      y0 = y1 = m.lineY;
    }
    const h = r.gh || r.h;
    if (p.align) return { page, y: p.align === 'center' ? (y0 + y1 - h) / 2 : /bottom|outside/.test(p.align) ? y1 - h : y0 };
    return { page, y: y0 + (p.off || 0) };
  }

  function placeOne(r, m, z) {
    const g = app.geom();
    let w = r.w;
    let h = r.h;
    let page;
    let x;
    let y;
    let wrap;
    if (r.inline) {
      // Satır içi resim: satırın olduğu yere, paragraf hizasına göre; sığmazsa sonraki sayfaya
      const s = Math.min(1, g.cw / (r.gw || w), g.ch / (r.gh || h));
      const gw = (r.gw || w) * s;
      const gh = (r.gh || h) * s;
      w *= s;
      h *= s;
      const fx = m.align === 'center' ? g.m.l + (g.cw - gw) / 2 : /right|end/.test(m.align) ? g.PW - g.m.r - gw : g.m.l;
      // Resmin metin boşluğu (WRAP_DIST) bir üstteki satırı itmesin; üstte yazı varsa onun yeri de ayrılır
      const capTop = r.caption && r.caption.pos === 'above' ? 20 + app.CAP_GAP : 0;
      const capBottom = r.caption && r.caption.pos !== 'above' ? 20 + app.CAP_GAP : 0;
      let fy = m.lineY + app.WRAP_DIST + capTop;
      page = m.page;
      if (fy + gh + capBottom > g.PH - g.m.b + 1) {
        page += 1;
        fy = g.m.t + capTop;
      }
      x = fx + (r.dx || 0) * s;
      y = fy + (r.dy || 0) * s;
      wrap = 'topbottom';
    } else {
      const v = vPos(r, g, m);
      page = v.page;
      x = hPos(r, g, m) + (r.dx || 0);
      y = v.y + (r.dy || 0);
      wrap = r.wrap;
    }
    const img = { id: SS.uid('img'), asset: r.asset, page: Math.max(0, page), x, y, w, h, rot: r.rot || 0, wrap, locked: false, z };
    if (r.crop) img.crop = r.crop.slice();
    if (r.caption) img.caption = { ...r.caption };
    app.clampToPage(img);
    app.state.images.push(img);
  }

  // Her yer tutucu belge sırasıyla ölçülür, resmi konur ve düzen yenilenir (sonraki ölçüm doğru olsun)
  function placeAll(ctx) {
    const els = app.els;
    const state = app.state;
    const recsByPh = new Map();
    for (const r of ctx.images) {
      if (!r.asset) continue;
      if (!recsByPh.has(r.ph)) recsByPh.set(r.ph, []);
      recsByPh.get(r.ph).push(r);
    }
    const order = ctx.images.filter((r) => r.asset).sort((a, b) => (a.inline ? 0 : 1) - (b.inline ? 0 : 1) || (a.relHeight || 0) - (b.relHeight || 0));
    const zOf = new Map(order.map((r, i) => [r.id, i + 1]));
    const blockSel = 'p,h1,h2,h3,li,div';
    // Word'deki satır konumları dul/öksüz satır kuralıyla dizilmiş düzene göredir: ölçmeden önce yer
    // tutucunun sayfasına (ve sonundaki sınıra) kadar sayfalamayı kesinleştir
    const settle = (ph) => {
      for (let i = 0; i < 4 && app.paginationPending(); i++) {
        const p = app.pageAtY((ph.getBoundingClientRect().top - els.doc.getBoundingClientRect().top) / state.zoom);
        app.paginateNow(p + 1);
        if (app.pageAtY((ph.getBoundingClientRect().top - els.doc.getBoundingClientRect().top) / state.zoom) === p) break;
      }
    };
    for (const ph of [...els.editor.querySelectorAll('span.ph')]) {
      const recs = recsByPh.get(ph.dataset.ph) || [];
      const block = ph.parentElement.closest(blockSel) || els.editor;
      if (recs.length) {
        settle(ph);
        const g = app.geom();
        const d = els.doc.getBoundingClientRect();
        const toDoc = (v, o) => (v - o) / state.zoom;
        const pr = ph.getBoundingClientRect();
        const rng = document.createRange();
        rng.selectNodeContents(block);
        const first = rng.getClientRects()[0] || block.getBoundingClientRect();
        const lineDocY = toDoc(pr.top, d.top);
        const paraDocY = toDoc(first.top, d.top);
        const page = app.pageAtY(lineDocY);
        const paraPage = app.pageAtY(paraDocY);
        const m = {
          page,
          lineY: lineDocY - page * g.stride,
          paraPage,
          paraY: paraDocY - paraPage * g.stride,
          phX: toDoc(pr.left, d.left),
          align: getComputedStyle(block).textAlign,
        };
        for (const r of recs) placeOne(r, m, zOf.get(r.id));
        ctx.report.images += recs.length;
      }
      ph.remove();
      const onlyImages = recs.some((r) => r.inline) && block !== els.editor && !block.textContent.trim() && !block.querySelector('span.ph');
      if (onlyImages) {
        const parent = block.parentElement;
        block.remove();
        if (parent && parent !== els.editor && /^(UL|OL)$/.test(parent.tagName) && !parent.children.length) parent.remove();
      } else if (block !== els.editor && !block.firstChild) block.innerHTML = '<br>';
      if (recs.length) {
        app.renderImages();
        app.layout();
      }
    }
    if (!els.editor.firstElementChild) els.editor.innerHTML = '<p><br></p>';
  }

  // ---------- Ana giriş ----------
  SS.importDocx = async function (file) {
    const pkg = await openPackage(file);
    const doc = await pkg.xml(pkg.docPath);
    const body = doc && all(doc, W, 'body')[0];
    if (!body) throw new Error('Belgenin içeriği okunamadı.');
    const rels = await pkg.relsOf(pkg.docPath);
    const partOf = async (type) => {
      const rel = [...rels.values()].find((r) => r.type.endsWith('/' + type));
      return rel ? pkg.xml(rel.target) : null;
    };
    const ctx = {
      pkg,
      rels,
      theme: parseTheme(await partOf('theme')),
      numbering: parseNumbering(await partOf('numbering')),
      styles: parseStyles(await partOf('styles')),
      cache: new Map(),
      images: [],
      seq: 0,
      inInstr: false,
      instr: '',
      href: null, // açık köprü (w:hyperlink)
      fldHref: null, // HYPERLINK alanının adresi
      pendingBoxes: [],
      lastImage: null,
      report: { images: 0, captions: 0, tables: 0, textboxes: 0, charts: 0, other: 0, skipped: 0, unsupported: 0, notes: 0 },
    };
    const blocks = attachCaptions(walkBody(body, ctx), ctx);
    const styles = importStyles(ctx);
    TAG = tagFrom(styles);
    const html = renderBlocks(blocks);
    const sect = kid(body, W, 'sectPr');
    const page = sect ? pageFromSect(sect) : app.defaultPage();
    const hfRes = sect ? await readHF(sect, ctx) : { hf: app.emptyHF(), lost: 0 };
    ctx.report.hf = [...hfRes.hf.header, ...hfRes.hf.footer].some(Boolean);
    ctx.report.hfLost = hfRes.lost;
    const assets = await loadAssets(ctx);
    // Belge adı Word'deki gibi dosya adıdır (belge özelliklerindeki "Başlık" değil): kaydetme ve Word'e aktarma bu adı önerir
    const title = file.name.replace(/\.docx$/i, '');

    app.load({ app: 'SerbestSayfa', title, page, hf: hfRes.hf, styles, html, images: [], assets });
    placeAll(ctx);
    app.relayoutAll();
    app.resetHistory();
    app.onLoad && app.onLoad();
    return ctx.report;
  };

  // Kullanıcıya gösterilecek özet
  SS.importSummary = function (rep) {
    const done = [];
    if (rep.images) done.push(`${rep.images} resim`);
    if (rep.captions) done.push(`${rep.captions} şekil yazısı`);
    if (rep.hf) done.push('üst/alt bilgi');
    const notes = [];
    if (rep.hfLost) notes.push(`üst/alt bilgideki ${rep.hfLost} öğe (resim, tablo, ek satır ya da kapağa özel içerik) alınamadı`);
    if (rep.tables) notes.push(`${rep.tables} tablo düz metne dönüştürüldü`);
    if (rep.textboxes) notes.push(`${rep.textboxes} metin kutusu paragrafa dönüştürüldü`);
    if (rep.unsupported) notes.push(`${rep.unsupported} resim desteklenmeyen biçimde (EMF/WMF/TIFF) olduğu için alınamadı`);
    if (rep.charts) notes.push(`${rep.charts} grafik alınamadı`);
    if (rep.other + rep.skipped) notes.push(`${rep.other + rep.skipped} şekil/nesne atlandı`);
    if (rep.notes) notes.push(`${rep.notes} dipnot alınamadı`);
    return 'Word belgesi açıldı' + (done.length ? ` (${done.join(', ')})` : '') + '.' + (notes.length ? ' ' + notes.join('; ') + '.' : '');
  };
})();
