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

  const state = (app.state = {
    page: defaultPage(),
    pageNumbers: false,
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
  app.CAP_GAP = U.cmToPx(0.15); // resim ile şekil yazısı arası
  app.CAP_MIN_W = U.cmToPx(4); // şekil yazısının en küçük genişliği
  const EXTRA_BANDS = 30; // metnin taşabileceği ileri sayfalar için önceden kurulan şerit sayısı

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
  // Numaralar belge sırasına göre (sayfa, üstten alta, soldan sağa), her etiket kendi içinde
  app.captionNumbers = function () {
    const list = state.images.filter((i) => i.caption).sort((a, b) => a.page - b.page || a.y - b.y || a.x - b.x);
    const counters = {};
    const nums = new Map();
    for (const img of list) {
      const l = img.caption.label || 'Şekil';
      counters[l] = (counters[l] || 0) + 1;
      nums.set(img.id, counters[l]);
    }
    return nums;
  };

  // ---------- Metin dışlama şeritleri ----------
  // Akış koordinatları: (0,0) = ilk sayfanın metin alanının sol-üst köşesi.
  // k. sayfanın metin alanı: y ∈ [k*stride, k*stride + ch]
  function computeBands() {
    const { cw, ch, stride, m } = app.geom();
    const full = [];
    const left = [];
    const right = [];
    // Yazdırırken yalnızca gerçek sayfa sınırları: akış kutusunun dışına taşan float'lar
    // tarayıcının sayfalara bölme hesabını bozuyor (son sayfanın üst boşluğu kayboluyordu).
    const bandCount = app.printMode ? state.pageCount - 1 : state.pageCount + EXTRA_BANDS;
    for (let k = 0; k < bandCount; k++) full.push([k * stride + ch, (k + 1) * stride]);

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
  function renderExclusions() {
    const { cw } = app.geom();
    const frag = document.createDocumentFragment();
    let prev = 0;
    for (const bd of computeBands()) {
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
      if (state.pageNumbers) {
        const pn = mk('pnum', `bottom:${app.FOOTER_DIST}px`);
        pn.textContent = k + 1;
        page.appendChild(pn);
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
  }

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
    el.firstChild.textContent = `${img.caption.label || 'Şekil'} ${nums.get(img.id) || 1}.`;
    const t = el.lastChild;
    if (!t.isContentEditable && t.textContent !== img.caption.text) t.textContent = img.caption.text || '';
    capH.set(img.id, el.offsetHeight);
    const r = app.capRect(img);
    el.style.left = r.x + 'px';
    el.style.top = r.y + 'px';
  }
  app.measureCaption = (img) => placeCaption(img, clipFor(img), app.captionNumbers());

  app.renderImages = function () {
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

  // ---------- Düzen ----------
  // Yalnızca <br> içeren boş satırlar tarayıcıda float şeritlerini yok sayar (sayfa arasına düşer).
  // "el" sınıfı, CSS ile görünmez bir boşluk ekleyerek bu satırların da şeritlerden kaçmasını sağlar.
  function markEmptyLines() {
    for (const el of els.editor.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,div:not(.pb)')) {
      const empty = !el.textContent && !el.querySelector('p,li,ul,ol,div');
      if (el.classList.contains('el') !== empty) el.classList.toggle('el', empty);
    }
  }

  app.layout = function () {
    markEmptyLines();
    renderExclusions();
    fitPageBreaks();
    for (let i = 0; i < 25; i++) {
      const n = neededPages();
      if (n === state.pageCount) break;
      state.pageCount = n;
      renderPages();
      renderExclusions();
      fitPageBreaks();
    }
    updateDocSize();
    app.onLayout && app.onLayout();
  };

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
      meta: JSON.stringify({ page: state.page, pageNumbers: state.pageNumbers, minPages: state.minPages }),
      sel: saveSel(),
    };
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
    state.images = JSON.parse(s.images);
    Object.assign(state, JSON.parse(s.meta));
    state.selection = state.selection.filter((id) => app.getImage(id));
    app.relayoutAll();
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
      pageNumbers: state.pageNumbers,
      minPages: state.minPages,
      html: els.editor.innerHTML.replace(/ class="el"/g, ''), // yardımcı sınıf düzen sırasında yeniden eklenir
      images: state.images,
      assets,
      savedAt: new Date().toISOString(),
    };
  };

  app.load = function (data) {
    if (!data || data.app !== 'SerbestSayfa') throw new Error('Bu dosya bir SerbestSayfa belgesi değil.');
    state.page = data.page || defaultPage();
    state.pageNumbers = !!data.pageNumbers;
    state.minPages = data.minPages || 1;
    state.images = data.images || [];
    state.assets = data.assets || {};
    state.fileName = data.title || 'Adsız belge';
    state.selection = [];
    state.pageCount = 1;
    els.editor.innerHTML = data.html || '<p><br></p>';
    app.relayoutAll();
    app.resetHistory();
    app.updateCtxBar && app.updateCtxBar();
    app.onLoad && app.onLoad();
  };

  app.newDocument = function () {
    app.load({ app: 'SerbestSayfa', page: defaultPage(), html: '<p><br></p>', images: [], assets: {} });
  };
})();
