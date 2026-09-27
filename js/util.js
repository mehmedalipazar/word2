/* Ortak yardımcılar: birimler, ikonlar, IndexedDB, dosya ve görsel işlemleri */
(function () {
  'use strict';
  const SS = (window.SS = window.SS || {});

  // ---------- Birimler (ekran 96 DPI kabul edilir) ----------
  const PX_PER_CM = 96 / 2.54;
  SS.units = {
    PX_PER_CM,
    cmToPx: (cm) => cm * PX_PER_CM,
    pxToCm: (px) => px / PX_PER_CM,
    pxToTwip: (px) => Math.round(px * 15), // 1px = 0,75pt = 15 twip
    pxToEmu: (px) => Math.round(px * 9525),
  };
  SS.PAGE_SIZES = { A4: [21, 29.7], A5: [14.8, 21], A3: [29.7, 42], Letter: [21.59, 27.94] };

  // ---------- İkonlar (24x24, çizgi tabanlı) ----------
  const ICONS = {
    file: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/>',
    open: '<path d="m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2"/>',
    save: '<path d="M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7"/><path d="M7 3v4a1 1 0 0 0 1 1h7"/>',
    download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/>',
    printer: '<path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M6 9V3a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v6"/><rect x="6" y="14" width="12" height="8" rx="1"/>',
    undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
    redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
    bold: '<path d="M6 12h9a4 4 0 0 1 0 8H7a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h7a4 4 0 0 1 0 8"/>',
    italic: '<path d="M19 4h-9M14 20H5M15 4 9 20"/>',
    underline: '<path d="M6 4v6a6 6 0 0 0 12 0V4"/><path d="M4 20h16"/>',
    strike: '<path d="M16 4H9a3 3 0 0 0-2.83 4"/><path d="M14 12a4 4 0 0 1 0 8H6"/><path d="M4 12h16"/>',
    subscript: '<path d="m4 5 8 8M12 5l-8 8"/><path d="M20 19h-4c0-1.5.44-2 1.5-2.5S20 15.33 20 14c0-.47-.17-.93-.48-1.29a2.11 2.11 0 0 0-2.62-.44c-.42.24-.74.62-.9 1.07"/>',
    superscript: '<path d="m4 19 8-8M12 19l-8-8"/><path d="M20 12h-4c0-1.5.44-2 1.5-2.5S20 8.33 20 7c0-.47-.17-.93-.48-1.29a2.11 2.11 0 0 0-2.62-.44c-.42.24-.74.62-.9 1.07"/>',
    search: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>',
    chevUp: '<path d="m18 15-6-6-6 6"/>',
    chevDown: '<path d="m6 9 6 6 6-6"/>',
    close: '<path d="M18 6 6 18M6 6l12 12"/>',
    paragraph: '<path d="M13 4v16M17 4v16M19 4H9.5a4.5 4.5 0 0 0 0 9H13"/>',
    highlight: '<path d="m9 11-6 6v3h9l3-3"/><path d="m22 12-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4"/>',
    eraser: '<path d="m7 21-4.3-4.3c-1-1-1-2.5 0-3.4l9.6-9.6c1-1 2.5-1 3.4 0l5.6 5.6c1 1 1 2.5 0 3.4L13 21"/><path d="M22 21H7"/><path d="m5 11 9 9"/>',
    alignLeft: '<path d="M21 6H3M15 12H3M17 18H3"/>',
    alignCenter: '<path d="M21 6H3M17 12H7M19 18H5"/>',
    alignRight: '<path d="M21 6H3M21 12H9M21 18H7"/>',
    alignJustify: '<path d="M3 6h18M3 12h18M3 18h18"/>',
    list: '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
    listOrdered: '<path d="M10 6h11M10 12h11M10 18h11M4 6h1v4M4 10h2M6 18H4c0-1 2-2 2-3s-1-1.5-2-1"/>',
    indent: '<path d="m3 8 4 4-4 4M21 12H11M21 6H11M21 18H11"/>',
    outdent: '<path d="m7 8-4 4 4 4M21 12H11M21 6H11M21 18H11"/>',
    image: '<rect width="18" height="18" x="3" y="3" rx="2"/><circle cx="9" cy="9" r="2"/><path d="m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21"/>',
    filePlus: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4M9 15h6M12 18v-6"/>',
    fileMinus: '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4M9 15h6"/>',
    pageSetup: '<rect x="4" y="2" width="16" height="20" rx="2"/><path d="M8 6h8M8 10h8M8 14h5"/>',
    wrapSquare: '<rect x="3" y="8" width="8" height="8" rx="1"/><path d="M14 9h7M14 12h7M14 15h7M3 4h18M3 20h18"/>',
    wrapTopBottom: '<rect x="6" y="8" width="12" height="8" rx="1"/><path d="M3 4h18M3 20h18"/>',
    wrapFront: '<path d="M3 5h18M3 10h18M3 15h18M3 20h18" opacity=".45"/><rect x="7" y="7" width="10" height="10" rx="1" fill="currentColor" fill-opacity=".3"/>',
    wrapBehind: '<rect x="7" y="7" width="10" height="10" rx="1" stroke-dasharray="2 2"/><path d="M3 5h18M3 10h18M3 15h18M3 20h18"/>',
    objLeft: '<path d="M3 2v20"/><rect x="7" y="5" width="14" height="5" rx="1.5"/><rect x="7" y="14" width="8" height="5" rx="1.5"/>',
    objCenter: '<path d="M12 2v20"/><rect x="4" y="5" width="16" height="5" rx="1.5"/><rect x="7" y="14" width="10" height="5" rx="1.5"/>',
    objRight: '<path d="M21 2v20"/><rect x="3" y="5" width="14" height="5" rx="1.5"/><rect x="9" y="14" width="8" height="5" rx="1.5"/>',
    objTop: '<path d="M2 3h20"/><rect x="5" y="7" width="5" height="14" rx="1.5"/><rect x="14" y="7" width="5" height="8" rx="1.5"/>',
    objMiddle: '<path d="M2 12h20"/><rect x="5" y="4" width="5" height="16" rx="1.5"/><rect x="14" y="7" width="5" height="10" rx="1.5"/>',
    objBottom: '<path d="M2 21h20"/><rect x="5" y="3" width="5" height="14" rx="1.5"/><rect x="14" y="9" width="5" height="8" rx="1.5"/>',
    distH: '<path d="M3 3v18M21 3v18"/><rect x="9" y="7" width="6" height="10" rx="1.5"/>',
    distV: '<path d="M3 3h18M3 21h18"/><rect x="7" y="9" width="10" height="6" rx="1.5"/>',
    fitWidth: '<path d="M3 12h18M7 8l-4 4 4 4M17 8l4 4-4 4"/>',
    front: '<rect x="8" y="8" width="8" height="8" rx="2"/><path d="M4 10a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2M14 20a2 2 0 0 0 2 2h4a2 2 0 0 0 2-2v-4a2 2 0 0 0-2-2"/>',
    back: '<rect x="14" y="14" width="8" height="8" rx="2"/><rect x="2" y="2" width="8" height="8" rx="2"/><path d="M7 14v1a2 2 0 0 0 2 2h1M14 7h1a2 2 0 0 1 2 2v1"/>',
    lock: '<rect width="18" height="11" x="3" y="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
    copy: '<rect width="14" height="14" x="8" y="8" rx="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>',
    trash: '<path d="M3 6h18M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>',
    magnet: '<path d="m6 15-4-4 6.75-6.77a7.79 7.79 0 0 1 11 11L13 22l-4-4 6.39-6.36a2.14 2.14 0 0 0-3-3L6 15"/><path d="m5 8 4 4M12 15l4 4"/>',
    grid: '<rect width="18" height="18" x="3" y="3" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>',
    caption: '<rect x="3" y="3" width="18" height="12" rx="2"/><path d="m3 12 5-4 4 3 3-2 6 4M6 19h12M9 22h6"/>',
    capPos: '<path d="M12 3v18M8 7l4-4 4 4M8 17l4 4 4-4"/>',
    pageBreak: '<path d="M5 3v5h14V3M5 21v-5h14v5M2 12h3M9.5 12h5M19 12h3"/>',
    zoomIn: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35M11 8v6M8 11h6"/>',
    zoomOut: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35M8 11h6"/>',
  };
  SS.icon = (name) =>
    `<svg class="ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name] || ''}</svg>`;
  SS.applyIcons = (root = document) => {
    root.querySelectorAll('[data-icon]').forEach((el) => {
      if (el.dataset.iconDone) return;
      el.insertAdjacentHTML('afterbegin', SS.icon(el.dataset.icon));
      el.dataset.iconDone = '1';
    });
  };

  // ---------- Küçük yardımcılar ----------
  SS.uid = (p = 'i') => p + Math.random().toString(36).slice(2, 8) + Date.now().toString(36).slice(-4);
  SS.clamp = (v, a, b) => Math.min(b, Math.max(a, v));
  SS.round = (v, d = 2) => Math.round(v * 10 ** d) / 10 ** d;

  let toastTimer = 0;
  SS.toast = (msg, ms = 2800) => {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), ms);
  };

  // ---------- IndexedDB (otomatik kayıt için; localStorage resimler için küçük kalır) ----------
  let dbPromise = null;
  function db() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        const req = indexedDB.open('serbestsayfa', 1);
        req.onupgradeneeded = () => req.result.createObjectStore('kv');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    return dbPromise;
  }
  function tx(mode, fn) {
    return db().then(
      (d) =>
        new Promise((resolve, reject) => {
          const t = d.transaction('kv', mode);
          const req = fn(t.objectStore('kv'));
          t.oncomplete = () => resolve(req && req.result);
          t.onerror = () => reject(t.error);
        })
    );
  }
  SS.idb = {
    get: (k) => tx('readonly', (s) => s.get(k)),
    set: (k, v) => tx('readwrite', (s) => s.put(v, k)),
    del: (k) => tx('readwrite', (s) => s.delete(k)),
    keys: () => tx('readonly', (s) => s.getAllKeys()),
  };

  // ---------- Dosya yardımcıları ----------
  SS.readFile = (file, as = 'dataURL') =>
    new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(r.error);
      if (as === 'text') r.readAsText(file);
      else if (as === 'buffer') r.readAsArrayBuffer(file);
      else r.readAsDataURL(file);
    });

  SS.dataURLtoBytes = (url) => {
    const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(url);
    if (!m) throw new Error('Geçersiz data URL');
    const mime = m[1];
    if (!m[2]) return { mime, bytes: new TextEncoder().encode(decodeURIComponent(m[3])) };
    const bin = atob(m[3]);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return { mime, bytes };
  };

  SS.downloadBlob = (blob, name) => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
  };

  SS.safeFileName = (s) => (s || 'belge').replace(/[\\/:*?"<>|]+/g, '_').trim() || 'belge';

  // ---------- Yazı tipinin "tek satır" yüksekliği ----------
  // Word'ün tek satır aralığı yazı tipinin doğal satır yüksekliğidir (üst + alt + satır arası boşluğu):
  // Calibri'de puntonun 1,22 katı, Times New Roman'da 1,15 katı. Tarayıcıdaki line-height: normal aynı
  // ölçüleri kullanır; bu katsayıyı yazı tipi başına bir kez ölçeriz. "1,5 satır" = 1,5 × bu katsayı.
  const lineFactors = new Map();
  SS.lineFactor = (family) => {
    const key = String(family || 'Calibri').toLowerCase();
    if (lineFactors.has(key)) return lineFactors.get(key);
    const d = document.createElement('div');
    d.style.cssText = 'position:absolute;left:-9999px;top:0;visibility:hidden;white-space:nowrap;line-height:normal;font-size:1000px';
    d.style.fontFamily = `"${String(family || 'Calibri').replace(/["']/g, '')}", Calibri, Carlito, sans-serif`; // editördeki yedek sırası
    d.textContent = 'Hg';
    document.body.appendChild(d);
    const f = d.getBoundingClientRect().height / 1000;
    d.remove();
    const v = f > 0.8 && f < 2 ? Math.round(f * 10000) / 10000 : 1.17;
    lineFactors.set(key, v);
    return v;
  };

  // ---------- Görsel hazırlama ----------
  // Word ve PDF'te sorunsuz görünmesi için: yalnızca PNG/JPEG tutulur, telefon fotoğraflarındaki
  // EXIF döndürme bilgisi piksele işlenir (Word bu bilgiyi her zaman dikkate almaz).
  function jpegOrientation(bytes) {
    try {
      const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      if (v.getUint16(0) !== 0xffd8) return 1;
      let off = 2;
      while (off + 4 <= v.byteLength) {
        const marker = v.getUint16(off);
        const len = v.getUint16(off + 2);
        if (marker === 0xffe1 && v.getUint32(off + 4) === 0x45786966) {
          const tiff = off + 10;
          const le = v.getUint16(tiff) === 0x4949;
          const ifd = tiff + v.getUint32(tiff + 4, le);
          const n = v.getUint16(ifd, le);
          for (let i = 0; i < n; i++) {
            const e = ifd + 2 + i * 12;
            if (v.getUint16(e, le) === 0x0112) return v.getUint16(e + 8, le);
          }
          return 1;
        }
        if ((marker & 0xff00) !== 0xff00 || marker === 0xffda) break;
        off += 2 + len;
      }
    } catch (_) { /* bozuk EXIF: yok say */ }
    return 1;
  }

  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('Görsel okunamadı'));
      img.src = src;
    });
  }
  SS.loadImage = loadImage;

  function rasterize(img, mime, scale = 1) {
    const w = Math.max(1, Math.round((img.naturalWidth || 800) * scale));
    const h = Math.max(1, Math.round((img.naturalHeight || 600) * scale));
    const c = document.createElement('canvas');
    c.width = w;
    c.height = h;
    const g = c.getContext('2d');
    if (mime === 'image/jpeg') {
      g.fillStyle = '#fff';
      g.fillRect(0, 0, w, h);
    }
    g.drawImage(img, 0, 0, w, h);
    return c.toDataURL(mime, 0.92);
  }

  // Dosya/Blob -> { src (dataURL), mime, nw, nh }
  SS.prepareImage = async (blob) => {
    let src = await SS.readFile(blob);
    let mime = (blob.type || /^data:([^;,]+)/.exec(src)?.[1] || '').toLowerCase();
    let img = await loadImage(src);
    if (mime === 'image/jpeg' || mime === 'image/jpg') {
      mime = 'image/jpeg';
      if (jpegOrientation(SS.dataURLtoBytes(src).bytes) > 1) {
        src = rasterize(img, 'image/jpeg');
        img = await loadImage(src);
      }
    } else if (mime !== 'image/png') {
      const scale = mime === 'image/svg+xml' ? 2 : 1;
      src = rasterize(img, 'image/png', scale);
      mime = 'image/png';
      img = await loadImage(src);
      if (scale !== 1) return { src, mime, nw: img.naturalWidth / scale, nh: img.naturalHeight / scale };
    }
    return { src, mime, nw: img.naturalWidth || 400, nh: img.naturalHeight || 300 };
  };
})();
