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
    new: newDoc,
    open: openDoc,
    save: () => saveDoc(false),
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
    fitWidth: () => app.fitWidth(),
    bringFront: () => app.bringFront(),
    sendBack: () => app.sendBack(),
    toggleLock: () => app.toggleLock(),
    duplicate: () => app.duplicate(),
    deleteImage: () => app.deleteSelected(),
    toggleCaption: () => app.toggleCaption(),
    captionPos: () => app.toggleCaptionPos(),
    pageBreak: () => app.insertPageBreak(),
    zoomIn: () => zoomStep(1),
    zoomOut: () => zoomStep(-1),
    zoomReset: () => setZoom(1),
  };
  for (const c of ['bold', 'italic', 'underline', 'strikeThrough', 'removeFormat', 'justifyLeft', 'justifyCenter',
    'justifyRight', 'justifyFull', 'insertUnorderedList', 'insertOrderedList']) commands[c] = () => app.exec(c);

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
  document.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    const k = e.key.toLowerCase();
    const a = document.activeElement || {};
    // Form alanları ve şekil yazısı düzenlenirken kendi geri alma/yineleme işlemleri çalışsın
    const inField = /^(INPUT|SELECT|TEXTAREA)$/.test(a.tagName) || (a.isContentEditable && a !== els.editor);
    if (k === 's') {
      e.preventDefault();
      saveDoc(e.shiftKey);
    } else if (k === 'o') {
      e.preventDefault();
      openDoc();
    } else if (k === 'p') {
      e.preventDefault();
      printDoc();
    } else if (inField) {
      /* alanların kendi geri alması çalışsın */
    } else if (k === 'z' && !e.shiftKey) {
      e.preventDefault();
      app.undo();
    } else if (k === 'y' || (k === 'z' && e.shiftKey)) {
      e.preventDefault();
      app.redo();
    } else if (k === '=' || k === '+') {
      e.preventDefault();
      zoomStep(1);
    } else if (k === '-') {
      e.preventDefault();
      zoomStep(-1);
    } else if (k === '0') {
      e.preventDefault();
      setZoom(1);
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

  function updatePageStyle() {
    const g = app.geom();
    const css = `@page { size: ${SS.round(U.pxToCm(g.PW), 3)}cm ${SS.round(U.pxToCm(g.PH), 3)}cm; margin: 0; }`;
    if ($('pageStyle').textContent !== css) $('pageStyle').textContent = css;
  }

  const dlg = $('pageDialog');
  function pageSetup() {
    const f = dlg.querySelector('form');
    const m = state.page.margins;
    f.elements.size.value = state.page.size;
    f.elements.orient.value = state.page.orient;
    f.elements.mt.value = SS.round(U.pxToCm(m.t), 2);
    f.elements.mb.value = SS.round(U.pxToCm(m.b), 2);
    f.elements.ml.value = SS.round(U.pxToCm(m.l), 2);
    f.elements.mr.value = SS.round(U.pxToCm(m.r), 2);
    f.elements.pageNumbers.checked = state.pageNumbers;
    dlg.returnValue = '';
    dlg.showModal();
  }
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
    state.pageNumbers = f.pageNumbers.checked;
    state.images.forEach(app.clampToPage);
    updatePageStyle();
    app.relayoutAll();
    app.commit('edit');
  });

  // ---------- Yazdırma / PDF ----------
  // Yazdırırken sayfa arası boşluk 0 yapılır; metin sayfa sınırlarına göre aynen yeniden dizilir.
  let printing = false;
  let savedZoom = 1;
  let savedGap = 24;
  function beforePrint() {
    if (printing) return;
    printing = true;
    savedZoom = state.zoom;
    savedGap = state.gap;
    state.zoom = 1;
    state.gap = 0;
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

  let wordTimer = 0;
  function updateWords() {
    clearTimeout(wordTimer);
    wordTimer = setTimeout(() => {
      const n = (els.editor.innerText.match(/\S+/g) || []).length;
      $('stWords').textContent = n.toLocaleString('tr-TR') + ' kelime';
      els.editor.classList.toggle('empty', !n && !state.images.length);
    }, 250);
  }
  function updateUndo() {
    document.querySelectorAll('[data-cmd="undo"]').forEach((b) => (b.disabled = !app.canUndo()));
    document.querySelectorAll('[data-cmd="redo"]').forEach((b) => (b.disabled = !app.canRedo()));
  }

  app.onLayout = () => {
    updateStatus();
    updatePageStyle(); // geri alma sayfa yapısını da değiştirebilir
  };
  let dirty = false; // son kayıt/açmadan beri değişiklik var mı
  app.onChange = () => {
    dirty = true;
    updateWords();
    updateUndo();
    scheduleAutosave();
  };
  app.onSelectionChange = (imgs) => {
    $('stHint').textContent = imgs.length
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
  let saveTimer = 0;
  function autosaveNow() {
    clearTimeout(saveTimer);
    return SS.idb.set('autosave', app.serialize()).catch(() => {});
  }
  function scheduleAutosave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(autosaveNow, 1200);
  }
  document.addEventListener('visibilitychange', () => document.visibilityState === 'hidden' && autosaveNow());
  window.addEventListener('pagehide', autosaveNow);

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
  const confirmDiscard = () =>
    !dirty || !hasContent() || window.confirm('Açık belgedeki kaydedilmemiş değişiklikler kaybolacak. Devam edilsin mi?');

  function newDoc() {
    if (!confirmDiscard()) return;
    fileHandle = null;
    app.newDocument();
    dirty = false;
    autosaveNow();
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
        app.load(JSON.parse(await SS.readFile(file, 'text')));
        SS.toast('Açıldı: ' + file.name);
      }
      dirty = false;
      autosaveNow();
      return true;
    } catch (err) {
      console.error(err);
      SS.toast('Dosya açılamadı: ' + err.message, 7000);
      return false;
    }
  }
  app.openFile = (file) => confirmDiscard() && loadFile(file);

  async function openDoc() {
    if (!confirmDiscard()) return;
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
    app.relayoutAll();
    app.resetHistory();
    updateUndo();
    updateWords();
    updateStatus();
    const fit = (els.workspace.clientWidth - 48) / app.geom().PW;
    if (fit < 1) setZoom(Math.max(0.5, Math.floor(fit * 20) / 20));
    dirty = false;
    try {
      const data = await SS.idb.get('autosave');
      const text = data && data.html ? data.html.replace(/<[^>]+>/g, '').trim() : '';
      if (data && data.app === 'SerbestSayfa' && (text || (data.images && data.images.length))) {
        app.load(data); // geri yüklenen çalışma bir dosyaya kaydedilmemiş olabilir: "değişmiş" sayılır
        SS.toast('Son çalışmanız geri yüklendi.');
      }
    } catch (_) { /* IndexedDB yoksa (ör. gizli pencere) sorun değil */ }
    els.editor.focus({ preventScroll: true });
  }
  init();
})();
