/* Bul ve değiştir (Ctrl+F / Ctrl+H).
 * Sonuçlar belgeye dokunmadan CSS Custom Highlight API ile boyanır; panel kapanınca seçili sonuç metin
 * seçimi olur. "Değiştir" ve "Tümünü değiştir" metin düğümlerini doğrudan düzenler (biçim korunur); her biri
 * tek geri alma adımıdır. "Tümünü değiştir" şekil yazılarını da kapsar (Word'deki metin kutuları gibi).
 * Büyük/küçük harf eşleştirilmezken Türkçe kurallar geçerlidir (I↔ı, İ↔i).
 */
(function () {
  'use strict';
  const SS = window.SS;
  const app = SS.app;
  const ed = app.els.editor;
  const $ = (id) => document.getElementById(id);
  const panel = $('findPanel');
  const qIn = $('findText');
  const rIn = $('replaceText');
  const caseBox = $('findCase');
  const wordBox = $('findWord');
  const countEl = $('findCount');
  const BLOCK = 'p,h1,h2,h3,h4,h5,h6,li,div';
  const WORD_CH = /[\p{L}\p{N}_]/u;
  const HL = !!(window.CSS && CSS.highlights && window.Highlight);

  let matches = []; // { nodes: [[metinDüğümü, baş, son], …], range }
  let current = -1;

  // Türkçe küçük harfe çevirme; karakter sayısı korunur (konumlar metinle örtüşsün)
  const fold = (s) => {
    let out = '';
    for (const c of s) {
      const l = c.toLocaleLowerCase('tr-TR');
      out += l.length === c.length ? l : c;
    }
    return out;
  };

  // Belgenin metni paragraf paragraf: { text, map: [[düğüm, metindeki başlangıcı], …] }. <br> satır sonu
  // olarak yazılır: sonuç satırlar arasında birleşmesin.
  function segments() {
    const out = [];
    let cur = null;
    let block = null;
    const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
      acceptNode: (n) => (n.nodeType === 3 || n.nodeName === 'BR' ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP),
    });
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      const b = n.parentElement.closest(BLOCK);
      if (!cur || b !== block) {
        cur = { text: '', map: [] };
        out.push(cur);
        block = b;
      }
      if (n.nodeType === 1) cur.text += '\n';
      else {
        cur.map.push([n, cur.text.length]);
        cur.text += n.nodeValue;
      }
    }
    return out;
  }

  // Bir metinde sonuçların [baş, son) konumları
  function find(text, q) {
    const res = [];
    const mc = caseBox.checked;
    const hay = mc ? text : fold(text);
    const needle = mc ? q : fold(q);
    let i = hay.indexOf(needle);
    while (i >= 0) {
      const j = i + needle.length;
      if (!wordBox.checked || (!WORD_CH.test(text[i - 1] || '') && !WORD_CH.test(text[j] || ''))) {
        res.push([i, j]);
        i = hay.indexOf(needle, j);
      } else i = hay.indexOf(needle, i + 1);
    }
    return res;
  }

  function toMatch(seg, i, j) {
    const nodes = [];
    for (const [n, s] of seg.map) {
      const e = s + n.length;
      if (e <= i || s >= j) continue;
      nodes.push([n, Math.max(i, s) - s, Math.min(j, e) - s]);
    }
    const range = document.createRange();
    range.setStart(nodes[0][0], nodes[0][1]);
    const last = nodes[nodes.length - 1];
    range.setEnd(last[0], last[2]);
    return { nodes, range };
  }

  function search() {
    const q = qIn.value;
    matches = [];
    if (q)
      for (const seg of segments())
        for (const [i, j] of find(seg.text, q)) matches.push(toMatch(seg, i, j));
    if (current >= matches.length) current = matches.length - 1;
    paint();
  }

  function paint() {
    if (HL) {
      if (matches.length) CSS.highlights.set('ss-bul', new Highlight(...matches.map((m) => m.range)));
      else CSS.highlights.delete('ss-bul');
      if (matches[current]) CSS.highlights.set('ss-bul-secili', new Highlight(matches[current].range));
      else CSS.highlights.delete('ss-bul-secili');
    }
    countEl.textContent = !qIn.value ? '' : !matches.length ? 'Sonuç yok' : current < 0 ? `${matches.length} sonuç` : `${current + 1} / ${matches.length}`;
  }

  function reveal(range) {
    const rc = range.getBoundingClientRect();
    const ws = app.els.workspace;
    const wr = ws.getBoundingClientRect();
    if (rc.top < wr.top + 40 || rc.bottom > wr.bottom - 40) ws.scrollTop += rc.top - (wr.top + wr.height / 3);
  }

  // İmlecin (ya da son sonucun) ardındaki ilk sonuç
  function go(dir) {
    if (!matches.length) return search();
    if (current < 0) {
      const caret = app.getCaretRange && app.getCaretRange();
      const after = (m) => caret && m.range.compareBoundaryPoints(Range.START_TO_START, caret) >= 0;
      current = dir > 0 ? Math.max(0, matches.findIndex(after)) : (matches.findLastIndex((m) => !after(m)) + matches.length) % matches.length;
    } else current = (current + dir + matches.length) % matches.length;
    paint();
    reveal(matches[current].range);
  }

  function replaceIn(m, text) {
    const [n0, a0, b0] = m.nodes[0];
    n0.nodeValue = n0.nodeValue.slice(0, a0) + text + n0.nodeValue.slice(b0);
    for (const [n, a, b] of m.nodes.slice(1)) n.nodeValue = n.nodeValue.slice(0, a) + n.nodeValue.slice(b);
    // Boşalan paragraf çökmesin
    const block = n0.parentElement.closest(BLOCK);
    if (block && !block.textContent && !block.querySelector('br')) block.appendChild(document.createElement('br'));
    return { node: n0, offset: a0 + text.length };
  }

  function afterEdit() {
    app.markDirty();
    app.scheduleLayout();
    app.commit('edit');
  }

  function replaceOne() {
    if (!matches[current]) return go(1);
    const at = replaceIn(matches[current], rIn.value);
    afterEdit();
    search();
    // Değiştirilen metnin ARDINDAN başlayan ilk sonuca geç (yeni metin aranan metni içerse de döngüye girmesin)
    const point = document.createRange();
    point.setStart(at.node, at.offset);
    const i = matches.findIndex((m) => m.range.compareBoundaryPoints(Range.START_TO_START, point) >= 0);
    current = matches.length ? (i < 0 ? 0 : i) : -1;
    paint();
    if (matches[current]) reveal(matches[current].range);
  }

  function replaceAll() {
    const q = qIn.value;
    if (!q) return;
    search();
    let n = matches.length;
    for (let k = matches.length - 1; k >= 0; k--) replaceIn(matches[k], rIn.value); // sondan: konumlar kaymasın
    let caps = 0;
    for (const img of app.state.images) {
      const t = img.caption && img.caption.text;
      if (!t) continue;
      const hits = find(t, q);
      if (!hits.length) continue;
      caps += hits.length;
      let s = t;
      for (const [i, j] of hits.reverse()) s = s.slice(0, i) + rIn.value + s.slice(j);
      img.caption.text = s;
    }
    n += caps;
    if (!n) return SS.toast('Değiştirilecek sonuç yok.');
    if (caps) app.renderImages();
    afterEdit();
    current = -1;
    search();
    SS.toast(`${n.toLocaleString('tr-TR')} yerde değiştirildi` + (caps ? ` (${caps} tanesi şekil yazısında).` : '.'));
  }

  // Panel açıkken belge değişirse sonuçlar yenilenir
  let timer = 0;
  const observer = new MutationObserver(() => {
    clearTimeout(timer);
    timer = setTimeout(search, 200);
  });

  app.openFind = function (replace) {
    const tb = document.querySelector('.toolbars').getBoundingClientRect();
    panel.style.top = tb.bottom + 8 + 'px';
    const wasHidden = panel.hidden;
    panel.hidden = false;
    // Seçili kısa metin aranacak metin olur (Word gibi)
    const r = app.getCaretRange && app.getCaretRange();
    const sel = r && !r.collapsed ? r.toString() : '';
    if (sel && sel.length < 100 && !/\n/.test(sel)) qIn.value = sel;
    (replace && qIn.value ? rIn : qIn).focus();
    (replace && qIn.value ? rIn : qIn).select();
    if (wasHidden) {
      current = -1;
      observer.observe(ed, { subtree: true, childList: true, characterData: true });
    }
    search();
  };

  function close() {
    panel.hidden = true;
    observer.disconnect();
    clearTimeout(timer);
    if (HL) {
      CSS.highlights.delete('ss-bul');
      CSS.highlights.delete('ss-bul-secili');
    }
    const m = matches[current];
    ed.focus({ preventScroll: true });
    if (m && ed.contains(m.range.startContainer)) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(m.range);
    }
    matches = [];
    current = -1;
  }

  qIn.addEventListener('input', () => {
    current = -1;
    search();
  });
  caseBox.addEventListener('change', search);
  wordBox.addEventListener('change', search);
  qIn.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      go(e.shiftKey ? -1 : 1);
    }
  });
  rIn.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      replaceOne();
    }
  });
  panel.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      close();
    }
  });
  panel.addEventListener('click', (e) => {
    const b = e.target.closest('[data-find]');
    if (!b) return;
    ({ next: () => go(1), prev: () => go(-1), close, replace: replaceOne, replaceAll })[b.dataset.find]();
  });
})();
