/* Uygulama kabuğu: komutlar, dosya işlemleri, yakınlaştırma, sayfa yapısı, yazdırma, otomatik kayıt */
(function () {
  'use strict';
  const SS = window.SS;
  const U = SS.units;
  const app = SS.app;
  const els = app.els;
  const state = app.state;
  const $ = (id) => document.getElementById(id);

  SS.applyIcons();

  // ---------- Komutlar ----------
  const commands = {
    new: () => newDoc(false),
    open: openDoc,
    save: () => saveDoc(false),
    close: () => newDoc(true),
    docx: exportDocx,
    print: printDoc,
    undo: () => app.undo(),
    redo: () => app.redo(),
    indent: () => app.indent(1),
    outdent: () => app.indent(-1),
    insertImage: () => $('fileImage').click(),
    addPage,
    removePage,
    pageSetup,
    headerFooter: () => pageSetup('h0'),
    fitWidth: () => app.fitWidth(),
    crop: () => app.toggleCrop(),
    resetCrop: () => app.resetCrop(),
    bringFront: () => app.bringFront(),
    sendBack: () => app.sendBack(),
    toggleLock: () => app.toggleLock(),
    duplicate: () => app.duplicate(),
    deleteImage: () => app.deleteSelected(true),
    toggleCaption: () => app.toggleCaption(),
    captionPos: () => app.toggleCaptionPos(),
    pageBreak: () => app.insertPageBreak(),
    find: () => app.openFind(true),
    symbol: () => app.symbolPanel(),
    paragraph: () => app.paragraphDialog(),
    link: () => app.linkDialog(),
    formatPainter: () => app.formatPainter(),
    zoomIn: () => zoomStep(1),
    zoomOut: () => zoomStep(-1),
    zoomReset: () => setZoom(1),
  };
  for (const c of ['bold', 'italic', 'underline', 'strikeThrough', 'removeFormat', 'justifyLeft', 'justifyCenter',
    'justifyRight', 'justifyFull', 'insertUnorderedList', 'insertOrderedList']) commands[c] = () => app.exec(c);
  commands.subscript = () => app.setScript('subscript');
  commands.superscript = () => app.setScript('superscript');

  // Düğmeler odağı çalmasın: metindeki seçim korunur
  document.addEventListener('mousedown', (e) => {
    if (e.target.closest('.chrome button')) e.preventDefault();
  });
  document.addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (!b || b.closest('dialog')) return;
    if (b.dataset.cmd && commands[b.dataset.cmd]) commands[b.dataset.cmd]();
    else if (b.dataset.wrap) app.setWrap(b.dataset.wrap);
    else if (b.dataset.align) app.alignSelected(b.dataset.align);
  });

  $('fileImage').addEventListener('change', (e) => {
    if (e.target.files.length) app.insertFiles(e.target.files, null);
    e.target.value = '';
  });
  $('fileOpen').addEventListener('change', (e) => {
    if (e.target.files[0]) loadFile(e.target.files[0]);
    e.target.value = '';
  });
  $('optSnap').addEventListener('change', (e) => (state.snap = e.target.checked));
  $('optGrid').addEventListener('change', (e) => {
    state.grid = e.target.checked;
    app.renderPages();
    app.layout();
  });

  // ---------- Klavye kısayolları ----------
  // Word'ün metin kısayolları. Tarayıcının aynı tuşlara bağlı işleri (Ctrl+R yenile, Ctrl+E/L adres çubuğu,
  // Ctrl+J indirilenler, Ctrl+H geçmiş, Ctrl+K arama) her durumda engellenir. İş, imleç metindeyken ya da
  // hiçbir alan odakta değilken yapılır; resim seçiliyken Ctrl+L/E/R resmi kenar boşluklarına göre hizalar.
  const TEXT_KEYS = {
    b: () => app.exec('bold'),
    i: () => app.exec('italic'),
    u: () => app.exec('underline'),
    e: () => app.exec('justifyCenter'),
    l: () => app.exec('justifyLeft'),
    r: () => app.exec('justifyRight'),
    j: () => app.exec('justifyFull'),
    ' ': () => app.exec('removeFormat'),
    m: () => app.indent(1),
    1: () => app.setLineHeight(1),
    2: () => app.setLineHeight(2),
    5: () => app.setLineHeight(1.5),
    '=': () => app.setScript('subscript'),
    '+': () => app.setScript('superscript'),
    '>': () => app.growFont(1),
    '<': () => app.growFont(-1),
    ']': () => app.growFont(1, true),
    '[': () => app.growFont(-1, true),
    k: () => app.linkDialog(),
  };
  const SHIFT_KEYS = { l: () => app.exec('insertUnorderedList'), m: () => app.indent(-1), '=': TEXT_KEYS['='], '+': TEXT_KEYS['+'], '>': TEXT_KEYS['>'], '<': TEXT_KEYS['<'],
    g: () => app.wordCountDialog() };
  const IMAGE_ALIGN = { l: 'left', e: 'center', r: 'right' };

  document.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    if (e.getModifierState && e.getModifierState('AltGraph')) return; // AltGr ile yazılan karakterler (@, #, [ …)
    const k = e.key.toLowerCase();
    const a = document.activeElement || {};
    // Form alanları ve şekil yazısı düzenlenirken kendi geri alma/yineleme işlemleri çalışsın
    const inField = /^(INPUT|SELECT|TEXTAREA)$/.test(a.tagName) || (a.isContentEditable && a !== els.editor);
    if (e.altKey) {
      // Ctrl+Alt+1/2/3: Başlık 1/2/3 (tuş başka bir karakter üretiyorsa dokunma)
      if (/^[123]$/.test(e.key) && !inField) {
        e.preventDefault();
        app.setBlock('h' + e.key);
      }
      return;
    }
    if (k === 's') {
      e.preventDefault();
      saveDoc(e.shiftKey);
    } else if (k === 'o') {
      e.preventDefault();
      openDoc();
    } else if (k === 'p') {
      e.preventDefault();
      printDoc();
    } else if (k === 'f' || k === 'h') {
      e.preventDefault();
      app.openFind && app.openFind(k === 'h');
    } else if (e.code === 'NumpadAdd' || e.code === 'NumpadSubtract' || k === '-' || k === '0') {
      e.preventDefault();
      if (k === '0') setZoom(1);
      else zoomStep(e.code === 'NumpadAdd' ? 1 : -1);
    } else if (e.shiftKey && (k === 'c' || k === 'v') && !inField && !app.objectsFocused()) {
      // Word: Ctrl+Shift+C / Ctrl+Shift+V biçimi kopyala / uygula. Biçim kopyalanmadıysa Ctrl+Shift+V tarayıcının
      // düz metin yapıştırması olarak kalır
      if (k === 'c') {
        e.preventDefault();
        app.copyFormat();
      } else if (app.hasFormat()) {
        e.preventDefault();
        app.pasteFormat();
      }
    } else if ((e.shiftKey ? SHIFT_KEYS : TEXT_KEYS)[k]) {
      e.preventDefault();
      if (inField) return;
      if (app.objectsFocused()) {
        if (IMAGE_ALIGN[k] && !e.shiftKey) app.alignSelected(IMAGE_ALIGN[k]);
        return;
      }
      (e.shiftKey ? SHIFT_KEYS : TEXT_KEYS)[k]();
    } else if (inField) {
      /* alanların kendi geri alması çalışsın */
    } else if (k === 'z' && !e.shiftKey) {
      e.preventDefault();
      app.undo();
    } else if (k === 'y' || (k === 'z' && e.shiftKey)) {
      e.preventDefault();
      app.redo();
    }
  });

  // ---------- Yakınlaştırma ----------
  const ZOOMS = [0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2];
  function setZoom(z) {
    const ws = els.workspace;
    const mid = (ws.scrollTop + ws.clientHeight / 2 - els.scaler.offsetTop) / state.zoom;
    state.zoom = z;
    app.updateDocSize();
    app.renderOverlay();
    ws.scrollTop = els.scaler.offsetTop + mid * z - ws.clientHeight / 2;
    $('zoomVal').textContent = Math.round(z * 100) + '%';
    app.updateRuler();
  }
  function zoomStep(dir) {
    const z = state.zoom;
    const next = dir > 0 ? ZOOMS.find((v) => v > z + 0.001) : [...ZOOMS].reverse().find((v) => v < z - 0.001);
    if (next) setZoom(next);
  }
  els.workspace.addEventListener(
    'wheel',
    (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      zoomStep(e.deltaY < 0 ? 1 : -1);
    },
    { passive: false }
  );

  // ---------- Sayfalar ----------
  function scrollToPage(k) {
    els.workspace.scrollTop = els.scaler.offsetTop + k * app.geom().stride * state.zoom - 20;
  }
  function addPage() {
    state.minPages = state.pageCount + 1;
    app.layout();
    app.commit('edit');
    scrollToPage(state.pageCount - 1);
  }
  function removePage() {
    const last = state.pageCount - 1;
    if (last === 0) return SS.toast('Belgede en az bir sayfa olmalı.');
    const before = state.pageCount;
    state.minPages = Math.min(state.minPages, last);
    app.layout();
    if (state.pageCount === before)
      SS.toast(state.images.some((i) => i.page === last) ? 'Son sayfada resim var; önce resmi taşıyın ya da silin.' : 'Son sayfada metin var.');
    else app.commit('edit');
  }

  // @page boyu cm olarak (3 basamak). Chrome basarken sayfaları, bu boyun 1/64 px'e yuvarlanıp tam piksele yukarı
  // yuvarlanmış aralığıyla böler (A4: 29,7 cm = 1122,52 px, sayfalar 1123 px'te bir; Letter 1056 px'te bir).
  const pageCm = (px) => SS.round(U.pxToCm(px), 3);
  const printPitch = (px) => Math.ceil(Math.round(U.cmToPx(pageCm(px)) * 64) / 64);
  function updatePageStyle() {
    const g = app.geom();
    const css = `@page { size: ${pageCm(g.PW)}cm ${pageCm(g.PH)}cm; margin: 0; }`;
    if ($('pageStyle').textContent !== css) $('pageStyle').textContent = css;
  }

  const dlg = $('pageDialog');
  // focus: açılınca imlecin gideceği alan (ör. "h1": üst bilginin orta yuvası)
  function pageSetup(focus) {
    const f = dlg.querySelector('form');
    const m = state.page.margins;
    f.elements.size.value = state.page.size;
    f.elements.orient.value = state.page.orient;
    const cmText = (px) => String(SS.round(U.pxToCm(px), 2)).replace('.', ','); // Türkçe ondalık virgül
    f.elements.mt.value = cmText(m.t);
    f.elements.mb.value = cmText(m.b);
    f.elements.ml.value = cmText(m.l);
    f.elements.mr.value = cmText(m.r);
    for (let i = 0; i < 3; i++) {
      f.elements['h' + i].value = state.hf.header[i];
      f.elements['f' + i].value = state.hf.footer[i];
    }
    f.elements.firstPage.checked = state.hf.firstPage;
    f.elements.pnPreset.value = '';
    dlg.returnValue = '';
    dlg.showModal();
    if (!$('symbolPanel').hidden) app.symbolPanel(dlg); // açık simge paneli pencerede de kullanılabilsin
    const inp = typeof focus === 'string' && f.elements[focus];
    if (inp) {
      inp.focus();
      inp.select();
    }
  }

  // Word gibi: sayfanın üst/alt kenar boşluğuna çift tıklamak üst/alt bilgiyi (tıklanan yuvayı) düzenlemeye açar
  function hfSlotAt(e) {
    if (e.target.closest && e.target.closest('.img-obj, .cap, .sel-box, .handle')) return null;
    const g = app.geom();
    const p = app.toDoc(e);
    const k = app.pageAtY(p.y);
    const y = p.y - k * g.stride;
    if (p.x < 0 || p.x > g.PW || y < 0 || y > g.PH) return null;
    const band = y < g.m.t ? 'h' : y > g.PH - g.m.b ? 'f' : '';
    return band && band + SS.clamp(Math.floor(((p.x - g.m.l) / g.cw) * 3), 0, 2);
  }
  els.doc.addEventListener('mousedown', (e) => e.detail > 1 && hfSlotAt(e) && e.preventDefault(), true); // sözcük seçilmesin
  els.doc.addEventListener('dblclick', (e) => {
    const slot = hfSlotAt(e);
    if (!slot) return;
    e.preventDefault();
    pageSetup(slot);
  });
  els.doc.addEventListener('mousemove', (e) => {
    const slot = document.body.classList.contains('dragging') ? null : hfSlotAt(e);
    const tip = slot ? (slot[0] === 'h' ? 'Üst' : 'Alt') + ' bilgi: düzenlemek için çift tıklayın' : '';
    if (els.doc.title !== tip) els.doc.title = tip;
  });
  // "Vazgeç" gönder düğmesi değil: Enter örtük gönderimde "Uygula"yı seçsin (ilk gönder düğmesi)
  document.querySelectorAll('dialog button[value="cancel"]').forEach((b) => b.addEventListener('click', () => b.closest('dialog').close('cancel')));
  // Simge paneli pencerenin içinde açılır (kalıcı pencerenin dışı tıklanamaz); düğme odağı almaz: simge imlecin
  // olduğu kutuya girer (text.js: insertSymbol)
  const hfSym = dlg.querySelector('[data-hf-sym]');
  hfSym.addEventListener('mousedown', (e) => e.preventDefault());
  hfSym.addEventListener('click', () => app.symbolPanel(dlg));
  // Hazır sayfa numarası seçeneği ilgili yuvaya yazılır
  dlg.querySelector('[name="pnPreset"]').addEventListener('change', (e) => {
    const [slot, text] = e.target.value.split('|');
    if (slot) dlg.querySelector('form').elements[slot].value = text;
    e.target.value = '';
  });
  dlg.addEventListener('close', () => {
    if (dlg.returnValue !== 'ok') return;
    const f = dlg.querySelector('form').elements;
    const cm = (v) => {
      const n = parseFloat(String(v).replace(',', '.'));
      return U.cmToPx(isFinite(n) ? SS.clamp(n, 0, 10) : 2.5);
    };
    const prev = state.page;
    state.page = {
      size: f.size.value,
      orient: f.orient.value,
      margins: { t: cm(f.mt.value), r: cm(f.mr.value), b: cm(f.mb.value), l: cm(f.ml.value) },
    };
    const g = app.geom();
    if (g.cw < 80 || g.ch < 80) {
      state.page = prev;
      return SS.toast('Kenar boşlukları kağıt için çok büyük.');
    }
    state.hf = {
      header: [0, 1, 2].map((i) => f['h' + i].value),
      footer: [0, 1, 2].map((i) => f['f' + i].value),
      firstPage: f.firstPage.checked,
    };
    state.images.forEach(app.clampToPage);
    updatePageStyle();
    app.relayoutAll();
    app.markDirty(); // sayfa boyutu/kenar boşluğu değişti: tüm sayfa sonları yeniden
    app.commit('edit');
  });

  // ---------- Yazdırma / PDF ----------
  // Yazdırırken sayfalar Chrome'un baskıdaki sayfa aralığıyla dizilir (printPitch; sayfa arası boşluk yalnızca kağıt
  // boyunun tam piksele tamamlanan kısmı, o da sayfanın altında basılmayan yerde); metin sayfa sınırlarına göre aynen
  // yeniden dizilir. Aralık kağıt boyu olunca her sayfada ~0,5 px kayma birikiyordu: sayfa başları yukarı kayıyor,
  // sayfa sınırını aşan itmeler (core.js: pushHanging) gerçek sınırın önünde kalıp metni aşağı itiyor, uzun belgenin
  // sonu son sayfadan taşıp basılmıyordu.
  let printing = false;
  let savedZoom = 1;
  let savedGap = 24;
  function beforePrint() {
    if (printing) return;
    app.paginateNow(); // sayfa sonları (dul/öksüz satır) kesinleşsin
    printing = true;
    savedZoom = state.zoom;
    savedGap = state.gap;
    state.zoom = 1;
    const { PH } = app.geom();
    state.gap = printPitch(PH) - PH;
    app.printMode = true;
    app.clearSelection();
    app.relayoutAll();
  }
  function afterPrint() {
    if (!printing) return;
    printing = false;
    app.printMode = false;
    state.zoom = savedZoom;
    state.gap = savedGap;
    app.relayoutAll();
  }
  window.addEventListener('beforeprint', beforePrint);
  window.addEventListener('afterprint', afterPrint);
  function printDoc() {
    beforePrint();
    window.print();
    afterPrint();
  }
  app.beforePrint = beforePrint;
  app.afterPrint = afterPrint;

  // ---------- Durum çubuğu ----------
  app.visiblePage = function () {
    const ws = els.workspace;
    const y = (ws.scrollTop + ws.clientHeight * 0.35 - els.scaler.offsetTop) / state.zoom;
    return SS.clamp(Math.floor(y / app.geom().stride), 0, state.pageCount - 1);
  };
  function updateStatus() {
    $('stPage').textContent = `Sayfa ${app.visiblePage() + 1} / ${state.pageCount}`;
  }
  els.workspace.addEventListener('scroll', updateStatus, { passive: true });

  // Sözcük sayısı (Word gibi): boşlukla ayrılan her parça bir sözcük; metin seçiliyken "seçili / toplam". Başlık ve
  // liste numaraları (CSS'le çizilir), şekil yazıları ve üst/alt bilgi sayılmaz.
  const countWords = (s) => (s.match(/\S+/g) || []).length;
  const trInt = (n) => n.toLocaleString('tr-TR');
  let totalWords = 0;
  function editorSelection() {
    const sel = window.getSelection();
    return sel.rangeCount && !sel.isCollapsed && els.editor.contains(sel.getRangeAt(0).commonAncestorContainer) ? sel : null;
  }
  function showWords() {
    const sel = editorSelection();
    $('stWords').textContent = (sel ? trInt(countWords(sel.toString())) + ' / ' : '') + trInt(totalWords) + ' sözcük';
  }
  let wordTimer = 0;
  function updateWords() {
    clearTimeout(wordTimer);
    wordTimer = setTimeout(() => {
      totalWords = countWords(els.editor.innerText);
      els.editor.classList.toggle('empty', !totalWords && !state.images.length);
      showWords();
    }, 250);
  }
  let selWordTimer = 0;
  document.addEventListener('selectionchange', () => {
    clearTimeout(selWordTimer);
    selWordTimer = setTimeout(showWords, 150);
  });

  // Sözcük Sayısı penceresi (Word: Gözden Geçir > Sözcük Sayısı, Ctrl+Shift+G): seçim varsa seçimin, yoksa belgenin
  // sayfa, sözcük, karakter, paragraf ve satır sayısı. Satır: metnin ekrandaki satırları (boş paragraflar hariç).
  app.wordCountDialog = function () {
    const sel = editorSelection();
    const r = sel && sel.getRangeAt(0);
    const text = sel ? sel.toString() : els.editor.innerText;
    const paras = new Set();
    const rects = [];
    const w = document.createTreeWalker(els.editor, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      if (!/\S/.test(t.nodeValue) || (r && !r.intersectsNode(t))) continue;
      const tr = document.createRange();
      tr.selectNodeContents(t);
      if (r && r.startContainer === t) tr.setStart(t, r.startOffset);
      if (r && r.endContainer === t) tr.setEnd(t, r.endOffset);
      if (!/\S/.test(tr.toString())) continue;
      const b = t.parentElement.closest('p,h1,h2,h3,li,div');
      if (b) paras.add(b);
      for (const q of tr.getClientRects()) if (q.width > 0) rects.push(q);
    }
    // Aynı satırdaki parçalar (farklı punto) dikeyde örtüşür
    rects.sort((a, b) => a.top - b.top);
    let lines = 0;
    let bottom = -Infinity;
    for (const q of rects) {
      if (q.top >= bottom - 1) {
        lines++;
        bottom = q.bottom;
      } else bottom = Math.max(bottom, q.bottom);
    }
    const chars = text.replace(/\n/g, '');
    $('wcScope').textContent = sel ? 'Seçili metin' : 'Belgenin tamamı';
    $('wcPages').textContent = trInt(state.pageCount);
    $('wcWords').textContent = trInt(countWords(text));
    $('wcChars').textContent = trInt(chars.replace(/\s/g, '').length);
    $('wcCharsSp').textContent = trInt(chars.length);
    $('wcParas').textContent = trInt(paras.size);
    $('wcLines').textContent = trInt(lines);
    $('wordsDialog').showModal();
  };
  $('stWords').addEventListener('mousedown', (e) => e.preventDefault()); // metindeki seçim korunur
  $('stWords').addEventListener('click', () => app.wordCountDialog());
  $('wordsDialog').addEventListener('close', () => app.focusEditor());
  function updateUndo() {
    document.querySelectorAll('[data-cmd="undo"]').forEach((b) => (b.disabled = !app.canUndo()));
    document.querySelectorAll('[data-cmd="redo"]').forEach((b) => (b.disabled = !app.canRedo()));
  }

  app.onLayout = () => {
    updateStatus();
    updatePageStyle(); // geri alma sayfa yapısını da değiştirebilir
    app.updateRuler();
  };
  let dirty = false; // son kayıt/açmadan beri değişiklik var mı
  app.onChange = () => {
    dirty = true;
    updateWords();
    updateUndo();
    scheduleAutosave();
  };
  app.onSelectionChange = (imgs) => {
    $('stHint').textContent = app.isCropping()
      ? 'Kırpma: siyah tutamaklar kenarları keser · Resmi sürükle: çerçevede kaydır · Esc/Enter: bitir'
      : imgs.length
        ? 'Sürükle: taşı · Köşe: boyut (Shift: serbest) · Üst tutamaç: döndür · Alt: yapışmasız · Oklar: ince ayar'
        : '';
  };

  const title = $('docTitle');
  function syncTitle() {
    title.value = state.fileName;
    document.title = state.fileName + ' — SerbestSayfa';
  }
  title.addEventListener('input', () => {
    state.fileName = title.value.trim() || 'Adsız belge';
    document.title = state.fileName + ' — SerbestSayfa';
    scheduleAutosave();
  });
  title.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      app.focusEditor();
    }
  });
  app.onLoad = () => {
    syncTitle();
    updatePageStyle();
    updateWords();
    updateUndo();
  };

  // ---------- Otomatik kayıt (tarayıcı içinde, IndexedDB) ----------
  // Her sekmenin kendi kurtarma kaydı var ('autosave:<sekme>'). Sekme kimliği sessionStorage'da durur:
  // sayfa yenilenince aynı kayıt geri gelir. Açık sekmeler Web Locks ile kilit tutar; yeni açılan sekme
  // yalnızca sahibi kapanmış ("sahipsiz") kayıtları devralır, başka bir sekmenin belgesini almaz.
  // IndexedDB yazımı zaman uyumsuzdur ve sayfa kapanırken yetişmeyebilir: kapanırken ya da gizlenirken
  // bekleyen değişiklikler localStorage'a eşzamanlı "acil kayıt" olarak da yazılır; açılışta yenisi seçilir.
  const AS = 'autosave:';
  const EMG = 'serbestsayfa-acil:';
  const LOCK = 'serbestsayfa-sekme:';
  const TAB = 'serbestsayfa-sekme';
  const recKey = (id) => (id ? AS + id : 'autosave'); // '' = eski sürümün tek kaydı
  const ls = (fn) => {
    try {
      return fn();
    } catch (_) {
      return null;
    }
  };
  let tabId = null;
  let saveTimer = 0;
  let changeSeq = 0; // her değişiklikte artar
  let savedSeq = 0; // IndexedDB'ye yazılmış son değişiklik
  const persisted = new Set(); // IndexedDB kaydında zaten bulunan resim varlıkları

  function autosaveNow() {
    clearTimeout(saveTimer);
    if (!tabId) return Promise.resolve();
    const seq = changeSeq;
    const data = app.serialize();
    const write = hasContent() ? SS.idb.set(recKey(tabId), data) : SS.idb.del(recKey(tabId));
    return write
      .then(() => {
        savedSeq = Math.max(savedSeq, seq);
        Object.keys(data.assets).forEach((a) => persisted.add(a));
        const e = ls(() => JSON.parse(localStorage.getItem(EMG + tabId)));
        if (!e || (e.savedAt || '') <= data.savedAt) ls(() => localStorage.removeItem(EMG + tabId));
      })
      .catch(() => {});
  }
  function scheduleAutosave() {
    changeSeq++;
    clearTimeout(saveTimer);
    saveTimer = setTimeout(autosaveNow, 1200);
  }
  // Eşzamanlı: yalnızca IndexedDB'de henüz olmayan resimler yazılır (sığmazsa resimsiz)
  function emergencySave() {
    if (!tabId || savedSeq === changeSeq) return;
    const data = app.serialize();
    const fresh = {};
    for (const [id, a] of Object.entries(data.assets)) if (!persisted.has(id)) fresh[id] = a;
    for (const assets of [fresh, {}]) if (ls(() => (localStorage.setItem(EMG + tabId, JSON.stringify({ ...data, assets })), true))) return;
  }
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState !== 'hidden') return;
    emergencySave();
    autosaveNow();
  });
  window.addEventListener('pagehide', () => {
    emergencySave();
    autosaveNow();
  });
  // Kaydedilmemiş değişiklik varken kapatma/yenileme: tarayıcının "Sayfadan ayrılınsın mı?" sorusu
  window.addEventListener('beforeunload', (e) => {
    emergencySave();
    if (dirty && hasContent()) {
      e.preventDefault();
      e.returnValue = '';
    }
  });

  // Sekme kilidi: sekme açık kaldıkça tutulur. Yenilemede eski sayfanın kilidi bir an sürebilir: kısa bekle;
  // alınamazsa kimlik başka bir açık sekmede (çoğaltılmış sekme) kullanılıyor demektir.
  function holdLock(id) {
    if (!navigator.locks) return Promise.resolve(true);
    return new Promise((resolve) => {
      const ac = new AbortController();
      const t = setTimeout(() => ac.abort(), 1500);
      navigator.locks
        .request(LOCK + id, { signal: ac.signal }, () => {
          clearTimeout(t);
          resolve(true);
          return new Promise(() => {});
        })
        .catch(() => resolve(false));
    });
  }

  // Kayıt = IndexedDB kaydı ile acil kayıttan yenisi (acil kayıttaki eksik resimler IndexedDB'den tamamlanır)
  async function readRecord(id) {
    const rec = await SS.idb.get(recKey(id)).catch(() => null);
    const emg = id ? ls(() => JSON.parse(localStorage.getItem(EMG + id))) : null;
    if (emg && emg.app === 'SerbestSayfa' && (!rec || (emg.savedAt || '') > (rec.savedAt || '')))
      return { data: { ...emg, assets: { ...(rec && rec.assets), ...emg.assets } }, stored: rec ? Object.keys(rec.assets || {}) : [] };
    return rec ? { data: rec, stored: Object.keys(rec.assets || {}) } : null;
  }
  const dropRecord = (id) => {
    ls(() => localStorage.removeItem(EMG + id));
    return SS.idb.del(recKey(id)).catch(() => {});
  };

  // Sahibi kapanmış kayıtlar (Web Locks yoksa: bu sekmeninki dışındaki tüm kayıtlar)
  async function orphanIds() {
    const ids = new Set();
    for (const k of (await SS.idb.keys().catch(() => [])) || [])
      if (k === 'autosave') ids.add('');
      else if (typeof k === 'string' && k.startsWith(AS)) ids.add(k.slice(AS.length));
    ls(() => Object.keys(localStorage).forEach((k) => k.startsWith(EMG) && ids.add(k.slice(EMG.length))));
    ids.delete(tabId);
    if (!navigator.locks) return [...ids];
    const held = new Set(((await navigator.locks.query()).held || []).map((l) => l.name));
    return [...ids].filter((id) => !held.has(LOCK + id));
  }

  const hasText = (data) =>
    data && data.app === 'SerbestSayfa' && ((data.html || '').replace(/<[^>]+>/g, '').trim() || (data.images && data.images.length));

  // Açılış: yenilenen sekme kendi kaydını, yeni sekme en yeni sahipsiz kaydı, çoğaltılmış sekme kaynağın kopyasını açar
  async function recoverWork() {
    let id = ls(() => sessionStorage.getItem(TAB));
    let from = null;
    let adopt = false;
    if (id && (await holdLock(id))) from = id;
    else {
      if (id) from = id; // çoğaltılmış sekme: kaynağın kaydını kopyala, kendine yeni kimlik al
      id = SS.uid('s');
      await holdLock(id);
    }
    tabId = id;
    ls(() => sessionStorage.setItem(TAB, id));
    let others = [];
    if (from === null) {
      others = await orphanIds();
      let best = null;
      for (const o of others) {
        const r = await readRecord(o);
        if (r && hasText(r.data) && (!best || (r.data.savedAt || '') > (best.r.data.savedAt || ''))) best = { o, r };
      }
      if (!best) return;
      others = others.filter((o) => o !== best.o);
      from = best.o;
      adopt = true;
    }
    const rec = await readRecord(from);
    if (!rec || !hasText(rec.data)) return;
    app.load(rec.data); // geri yüklenen çalışma bir dosyaya kaydedilmemiş olabilir: "değişmiş" sayılır
    if (from === tabId) rec.stored.forEach((a) => persisted.add(a));
    changeSeq++;
    await autosaveNow(); // devralınan/kopyalanan kayıt artık bu sekmenin
    if (adopt) await dropRecord(from);
    const more = (await Promise.all(others.map(readRecord))).filter((r) => r && hasText(r.data)).length;
    SS.toast('Son çalışmanız geri yüklendi.' + (more ? ` Kurtarılabilecek ${more} belge daha var; yeni bir sekmede açılır.` : ''), more ? 6000 : 2800);
  }

  // ---------- Dosya: yeni / aç / kaydet / Word'e aktar ----------
  let fileHandle = null;
  const hasContent = () => els.editor.innerText.trim().length > 0 || state.images.length > 0;
  const FILE_TYPES = [{ description: 'SerbestSayfa belgesi', accept: { 'application/json': ['.sayfa'] } }];
  const OPEN_TYPES = [
    {
      description: 'SerbestSayfa ya da Word belgesi',
      accept: {
        'application/json': ['.sayfa', '.json'],
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
      },
    },
  ];
  // Word gibi: kaydedilmemiş değişiklik varsa "Kaydet / Kaydetme / Vazgeç" sorulur; true: devam edilebilir.
  // Kaydet seçilip kayıt penceresinden vazgeçilirse belge açık kalır.
  const saveDlg = $('saveDialog');
  function askSave() {
    if (saveDlg.open) return Promise.resolve(''); // pencere açıkken gelen ikinci istek (ör. Ctrl+O)
    $('saveAsk').textContent = `"${state.fileName}" belgesinde kaydedilmemiş değişiklikler var. Kaydetmezseniz bu değişiklikler kaybolur.`;
    saveDlg.returnValue = '';
    saveDlg.showModal();
    return new Promise((resolve) => saveDlg.addEventListener('close', () => resolve(saveDlg.returnValue), { once: true }));
  }
  async function confirmDiscard() {
    if (!dirty || !hasContent()) return true;
    const answer = await askSave();
    if (answer === 'discard') return true;
    if (answer !== 'save') return false;
    await saveDoc(false);
    return !dirty;
  }

  // Yeni belge ya da belgeyi kapat (Word: Dosya > Kapat). Tek belgelik uygulamada kapatınca boş belge kalır; boş
  // belgenin kaydı tutulmadığı için bu sekmenin kurtarma kaydı da silinir. Ctrl+W ve Ctrl+F4 tarayıcıya ayrılmış.
  async function newDoc(closing) {
    if (!(await confirmDiscard())) return;
    const name = state.fileName;
    fileHandle = null;
    app.newDocument();
    dirty = false;
    autosaveNow();
    if (closing) SS.toast(`"${name}" kapatıldı.`);
    app.focusEditor();
  }

  // .sayfa ya da .docx aç; başarılıysa true
  async function loadFile(file) {
    try {
      if (/\.docx$/i.test(file.name)) {
        SS.toast('Word belgesi açılıyor…', 20000);
        const rep = await SS.importDocx(file);
        fileHandle = null; // kaydederken .sayfa olarak yeni yer sorulsun
        SS.toast(SS.importSummary(rep), 8000);
      } else {
        let data = null;
        try {
          data = JSON.parse(await SS.readFile(file, 'text'));
        } catch (_) { /* JSON değil: app.load Türkçe iletiyle reddeder */ }
        app.load(data);
        SS.toast('Açıldı: ' + file.name);
      }
      dirty = false;
      autosaveNow();
      return true;
    } catch (err) {
      // Uygulamanın kendi iletileri (Error) Türkçe; tarayıcının çözümleme/çalışma hataları (TypeError, SyntaxError…)
      // kullanıcıya İngilizce ham ileti olarak gösterilmez, yalnızca konsola yazılır
      const own = err && err.constructor === Error;
      if (!own) console.error(err);
      SS.toast('Dosya açılamadı: ' + (own ? err.message : 'dosya bozuk ya da desteklenmeyen biçimde.'), 7000);
      return false;
    }
  }
  app.openFile = async (file) => (await confirmDiscard()) && loadFile(file);

  async function openDoc() {
    if (!(await confirmDiscard())) return;
    // Kayıt penceresi uzun sürdüyse tıklamanın verdiği izin (geçici etkinlik) bitmiş olur ve Chrome dosya seçiciyi
    // açmaz: sessizce hiçbir şey olmaması yerine yeniden tıklamayı iste
    if (navigator.userActivation && !navigator.userActivation.isActive)
      return void SS.toast("Açılacak dosyayı seçmek için Aç'a yeniden tıklayın.", 5000);
    if (window.showOpenFilePicker) {
      try {
        const [h] = await window.showOpenFilePicker({ types: OPEN_TYPES });
        const file = await h.getFile();
        if (await loadFile(file)) fileHandle = /\.(sayfa|json)$/i.test(file.name) ? h : null;
        return;
      } catch (err) {
        if (err.name === 'AbortError') return;
      }
    }
    $('fileOpen').click();
  }

  async function saveDoc(asNew) {
    const data = JSON.stringify(app.serialize());
    const name = SS.safeFileName(state.fileName) + '.sayfa';
    if (window.showSaveFilePicker) {
      try {
        if (!fileHandle || asNew) fileHandle = await window.showSaveFilePicker({ suggestedName: name, types: FILE_TYPES });
        const w = await fileHandle.createWritable();
        await w.write(data);
        await w.close();
        dirty = false;
        return SS.toast('Kaydedildi: ' + fileHandle.name);
      } catch (err) {
        if (err.name === 'AbortError') return;
        fileHandle = null;
      }
    }
    SS.downloadBlob(new Blob([data], { type: 'application/json' }), name);
    dirty = false;
    SS.toast('"İndirilenler" klasörüne kaydedildi: ' + name);
  }

  async function exportDocx() {
    try {
      app.clearSelection();
      const blob = await SS.exportDocx(app);
      SS.downloadBlob(blob, SS.safeFileName(state.fileName) + '.docx');
      dirty = false; // çalışma bir dosyada: kapatırken sorulmasın
      SS.toast('Word dosyası hazır (İndirilenler klasörü).');
    } catch (err) {
      console.error(err);
      SS.toast('Dışa aktarılamadı: ' + err.message, 5000);
    }
  }

  // ---------- Başlangıç ----------
  async function init() {
    syncTitle();
    updatePageStyle();
    state.styles = app.defaultStyles(); // ilk boş belge de "Varsayılan olarak ayarla" ile saklanan stillerle
    app.applyStyles();
    app.relayoutAll();
    app.resetHistory();
    updateUndo();
    updateWords();
    updateStatus();
    const fit = (els.workspace.clientWidth - 48) / app.geom().PW;
    if (fit < 1) setZoom(Math.max(0.5, Math.floor(fit * 20) / 20));
    dirty = false;
    try {
      await recoverWork();
    } catch (_) { /* IndexedDB yoksa (ör. gizli pencere) sorun değil */ }
    els.editor.focus({ preventScroll: true });
  }
  init();
})();
