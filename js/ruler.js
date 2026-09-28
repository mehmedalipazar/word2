/* Cetvel: imlecin bulunduğu paragrafın girintileri ve sekme durakları; işaretler sürüklenerek değiştirilir (Word gibi).
 * İlk satır (üst üçgen): yalnızca ilk satır girintisi. Asılı (alt üçgen): sol girinti, ilk satır yerinde kalır.
 * Sol (kutu): sol girinti ilk satırla birlikte. Sağ (sağdaki üçgen): sağ girinti.
 * Sekme durakları (core.js: data-tabs): cetvele tıklamak soldaki kutuda seçili türde durak koyar, durak sürüklenerek
 * taşınır, cetvelin dışına (aşağı/yukarı) sürüklenince kalkar; çift tık Sekmeler penceresini açar.
 * 0,25 cm'ye yapışır (Alt: serbest). Liste maddelerinde girinti düzeye bağlıdır: işaretler yalnızca gösterilir.
 */
(function () {
  'use strict';
  const SS = window.SS;
  const U = SS.units;
  const app = SS.app;
  const els = app.els;
  const state = app.state;
  const ruler = document.getElementById('ruler');
  const pageEl = document.getElementById('rulerPage');
  const marks = {};
  pageEl.querySelectorAll('.rm').forEach((m) => (marks[m.dataset.rm] = m));
  const STEP = U.cmToPx(0.25);
  const CM = U.cmToPx(1);
  const typeEl = document.getElementById('rulerTabType');
  const TYPES = ['l', 'c', 'r', 'd'];
  const TYPE_NAMES = { l: 'Sola', c: 'Ortaya', r: 'Sağa', d: 'Ondalık' };
  let tabType = 'l';
  typeEl.addEventListener('click', () => {
    tabType = TYPES[(TYPES.indexOf(tabType) + 1) % TYPES.length];
    typeEl.className = 'ruler-tabtype ' + tabType;
    typeEl.title = typeEl.title.replace(/^Sekme türü: \S+/, 'Sekme türü: ' + TYPE_NAMES[tabType]);
  });
  let geomKey = '';
  let ticks = null;

  // Ölçek çizgileri: 0 sol kenar boşluğunda, her 0,25 cm'de bir çizgi, her cm'de sayı
  function build() {
    const g = app.geom();
    const z = state.zoom;
    const key = [g.PW, g.m.l, g.m.r, z].join('|');
    if (key === geomKey) return;
    geomKey = key;
    pageEl.style.width = g.PW * z + 'px';
    pageEl.style.setProperty('--ml', g.m.l * z + 'px');
    pageEl.style.setProperty('--mr', g.m.r * z + 'px');
    const t = document.createElement('div');
    t.className = 'ruler-ticks';
    for (let i = -Math.floor((g.m.l / CM) * 4); ((i * CM) / 4) <= g.PW - g.m.l; i++) {
      if (!i) continue;
      const s = document.createElement('span');
      s.style.left = (g.m.l + (i * CM) / 4) * z + 'px';
      if (i % 4 === 0) {
        s.className = 'num';
        s.textContent = Math.abs(i / 4);
      } else s.className = i % 2 ? 'min' : 'mid';
      t.appendChild(s);
    }
    if (ticks) ticks.remove();
    pageEl.prepend((ticks = t));
  }

  // Cetvel sayfayla yatayda hizalı kalır (yakınlaştırma, yatay kaydırma, pencere boyutu)
  function place() {
    pageEl.style.left = els.doc.getBoundingClientRect().left - ruler.getBoundingClientRect().left + 'px';
  }

  function caretBlock() {
    const r = app.getCaretRange && app.getCaretRange();
    if (!r || !els.editor.contains(r.startContainer)) return null;
    const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    const b = n && n.closest('p,h1,h2,h3,h4,h5,h6,li,div:not(.pb)');
    return b && els.editor.contains(b) ? b : null;
  }

  // İşaretleri metin alanına göre konumlar (px): F ilk satır, L diğer satırlar, R sağ girinti
  function setMarks(F, L, R) {
    const g = app.geom();
    const z = state.zoom;
    marks.first.style.left = (g.m.l + F) * z + 'px';
    marks.hang.style.left = marks.left.style.left = (g.m.l + L) * z + 'px';
    marks.right.style.left = (g.PW - g.m.r - R) * z + 'px';
  }

  // İmlecin paragrafının sekme durakları
  function showTabs(b) {
    pageEl.querySelectorAll('.rt').forEach((r) => r.remove());
    if (!b || !b.dataset.tabs) return;
    const g = app.geom();
    for (const t of app.parseTabs(b.dataset.tabs)) {
      const d = document.createElement('div');
      d.className = 'rt ' + t.type;
      d.dataset.pos = t.pos;
      d.style.left = (g.m.l + t.pos) * state.zoom + 'px';
      d.title = `${TYPE_NAMES[t.type]} sekme durağı: ${String(SS.round(U.pxToCm(t.pos), 2)).replace('.', ',')} cm (sürükle: taşı, cetvelin dışına sürükle: kaldır, çift tık: Sekmeler)`;
      pageEl.appendChild(d);
    }
  }

  function showMarks() {
    const b = state.selection.length ? null : caretBlock();
    pageEl.classList.toggle('off', !b);
    showTabs(b);
    if (!b) return;
    const z = state.zoom;
    const cs = getComputedStyle(b);
    const er = els.editor.getBoundingClientRect();
    const br = b.getBoundingClientRect();
    const L = (br.left - er.left) / z + (parseFloat(cs.paddingLeft) || 0);
    const R = (er.right - br.right) / z + (parseFloat(cs.paddingRight) || 0);
    setMarks(L + (parseFloat(cs.textIndent) || 0), L, R);
    pageEl.classList.toggle('list', b.tagName === 'LI');
  }

  app.updateRuler = function () {
    build();
    place();
    showMarks();
  };
  document.addEventListener('selectionchange', () => requestAnimationFrame(showMarks));
  ruler.addEventListener('mousedown', (e) => e.preventDefault()); // odak metinden kaçmasın
  els.workspace.addEventListener('scroll', place, { passive: true });
  window.addEventListener('resize', place);

  const tabBlocks = () => app.selectedBlocks().filter((b) => !b.classList.contains('pb'));
  function setTabs(blocks, fn) {
    for (const b of blocks) {
      const list = fn(app.parseTabs(b.dataset.tabs));
      if (list.length) b.dataset.tabs = app.formatTabs(list);
      else b.removeAttribute('data-tabs');
    }
    app.normalizeBlocks(); // sekmeler span'a (core.js: wrapTabs)
    app.scheduleLayout();
  }
  // Durak koy / taşı / kaldır
  function tabPointer(e, stopEl) {
    app.focusEditor();
    const blocks = tabBlocks();
    if (!blocks.length) return;
    const g = app.geom();
    const pr = pageEl.getBoundingClientRect();
    const at = (ev) => {
      const v = (ev.clientX - pr.left) / state.zoom - g.m.l;
      return SS.clamp(ev.altKey ? v : Math.round(v / STEP) * STEP, 0, g.cw);
    };
    let from = stopEl ? +stopEl.dataset.pos : null;
    if (from === null) {
      // Boş yere tıklama: seçili türde yeni durak
      const pos = at(e);
      setTabs(blocks, (l) => l.filter((t) => Math.abs(t.pos - pos) > 0.5).concat({ type: tabType, pos, leader: '' }));
      app.afterFormat();
      showTabs(blocks[0]);
      return;
    }
    const type = stopEl.className.split(' ')[1];
    const leader = (app.parseTabs(blocks[0].dataset.tabs).find((t) => Math.abs(t.pos - from) < 0.5) || {}).leader || '';
    let moved = false;
    let gone = false;
    const move = (ev) => {
      const off = Math.abs(ev.clientY - (pr.top + pr.height / 2)) > 24; // cetvelin dışı: kaldır
      const pos = at(ev);
      if (!moved && Math.abs(ev.clientX - e.clientX) < 2 && !off) return;
      moved = true;
      setTabs(blocks, (l) => {
        const rest = l.filter((t) => Math.abs(t.pos - from) > 0.5 && Math.abs(t.pos - pos) > 0.5);
        return off ? rest : rest.concat({ type, pos, leader });
      });
      gone = off;
      if (!off) from = pos;
      showTabs(blocks[0]);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', up);
      if (moved) app.afterFormat();
      if (gone) SS.toast('Sekme durağı kaldırıldı.', 1500);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', up);
  }
  ruler.addEventListener('dblclick', (e) => {
    if (e.target.closest('.rt') && app.tabsDialog) app.tabsDialog();
  });

  ruler.addEventListener('pointerdown', (e) => {
    const m = e.target.closest('.rm');
    e.preventDefault(); // metindeki seçim korunur
    if (e.target === typeEl || typeEl.contains(e.target) || pageEl.classList.contains('off')) return;
    const stop = e.target.closest('.rt');
    if (stop || (!m && pageEl.contains(e.target))) return tabPointer(e, stop);
    if (!m || pageEl.classList.contains('list')) return;
    app.focusEditor();
    const blocks = app.selectedBlocks().filter((b) => b.tagName !== 'LI' && !b.classList.contains('pb'));
    if (!blocks.length) return;
    const cs = getComputedStyle(blocks[0]);
    const L0 = parseFloat(cs.marginLeft) || 0;
    const T0 = parseFloat(cs.textIndent) || 0;
    const R0 = parseFloat(cs.marginRight) || 0;
    const x0 = e.clientX;
    const kind = m.dataset.rm;
    const g = app.geom();
    let moved = false;
    const move = (ev) => {
      const d = (ev.clientX - x0) / state.zoom;
      const snap = (v) => (ev.altKey ? v : Math.round(v / STEP) * STEP);
      let L = L0;
      let T = T0;
      let R = R0;
      if (kind === 'first') T = Math.max(0, snap(L0 + T0 + d)) - L0;
      else if (kind === 'hang') {
        L = Math.max(0, snap(L0 + d));
        T = L0 + T0 - L; // ilk satır yerinde kalır
      } else if (kind === 'left') L = Math.max(Math.max(0, -T0), snap(L0 + d));
      else R = Math.max(0, snap(R0 - d));
      if (g.cw - Math.max(L, L + T) - R < U.cmToPx(2)) return; // metne en az 2 cm kalsın
      for (const b of blocks) {
        app.setIndent(b, 'marginLeft', L);
        app.setIndent(b, 'textIndent', T);
        app.setIndent(b, 'marginRight', R);
      }
      moved = true;
      setMarks(L + T, L, R);
      app.scheduleLayout();
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', up);
      if (moved) app.afterFormat();
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', up);
  });
})();
