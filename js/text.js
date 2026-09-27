/* Metin düzenleme: biçimlendirme komutları, yapıştırma temizliği, araç çubuğu durumu */
(function () {
  'use strict';
  const SS = window.SS;
  const app = SS.app;
  const ed = app.els.editor;
  const $ = (id) => document.getElementById(id);
  const fontFamilySel = $('fontFamily');
  const fontSizeSel = $('fontSize');
  const blockSel = $('blockStyle');
  const lineHeightSel = $('lineHeight');

  document.execCommand('defaultParagraphSeparator', false, 'p');
  document.execCommand('styleWithCSS', false, true);

  const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,li,div,blockquote,pre';
  const FONTS = [...fontFamilySel.options].map((o) => o.text);

  // ---------- Seçim takibi ----------
  // Açılır listeler odağı editörden alır; komut uygulamadan önce son metin seçimini geri yükleriz.
  let savedRange = null;
  document.addEventListener('selectionchange', () => {
    const sel = window.getSelection();
    if (sel.rangeCount && ed.contains(sel.getRangeAt(0).commonAncestorContainer)) {
      savedRange = sel.getRangeAt(0).cloneRange();
      updateToolbarState();
    }
  });
  app.getCaretRange = () => savedRange;

  function focusEditor() {
    if (document.activeElement !== ed) ed.focus({ preventScroll: true });
    if (savedRange && ed.contains(savedRange.startContainer)) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
  }
  app.focusEditor = focusEditor;

  function afterFormat() {
    fixFontTags();
    app.scheduleLayout();
    app.commit('edit');
    updateToolbarState();
  }

  app.exec = function (cmd, value = null) {
    app.clearSelection && app.clearSelection();
    focusEditor();
    document.execCommand(cmd, false, value);
    afterFormat();
  };

  function selectedBlocks() {
    const sel = window.getSelection();
    if (!sel.rangeCount) return [];
    const r = sel.getRangeAt(0);
    let list = [...ed.querySelectorAll(BLOCK_SEL)].filter((b) => r.intersectsNode(b));
    list = list.filter((b) => !list.some((o) => o !== b && b.contains(o)));
    if (!list.length) {
      const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      const b = n && n.closest(BLOCK_SEL);
      if (b && ed.contains(b)) list = [b];
    }
    return list;
  }

  // execCommand('fontSize') yalnızca 1-7 arası değer alır: 7 ile işaretleyip gerçek pt değeriyle değiştiriyoruz
  let pendingSize = null;
  function fixFontTags() {
    const size = pendingSize || fontSizeSel.value;
    ed.querySelectorAll('font[size="7"]').forEach((f) => {
      f.removeAttribute('size');
      f.style.fontSize = size + 'pt';
    });
    ed.querySelectorAll('span[style*="xxx-large"]').forEach((s) => (s.style.fontSize = size + 'pt'));
  }

  app.setFontSize = function (pt) {
    focusEditor();
    pendingSize = pt;
    document.execCommand('styleWithCSS', false, false);
    document.execCommand('fontSize', false, '7');
    document.execCommand('styleWithCSS', false, true);
    afterFormat();
  };

  app.setBlock = function (tag) {
    focusEditor();
    document.execCommand('formatBlock', false, tag);
    afterFormat();
  };

  app.setLineHeight = function (v) {
    focusEditor();
    selectedBlocks().forEach((b) => (b.style.lineHeight = v));
    afterFormat();
  };

  app.indent = function (dir) {
    focusEditor();
    const blocks = selectedBlocks();
    if (blocks.some((b) => b.tagName === 'LI')) document.execCommand(dir > 0 ? 'indent' : 'outdent');
    else
      blocks.forEach((b) => {
        const next = Math.max(0, (parseFloat(b.style.marginLeft) || 0) + dir * 48);
        b.style.marginLeft = next ? next + 'px' : '';
      });
    afterFormat();
  };

  // ---------- Araç çubuğu durumu ----------
  function setSelectValue(sel, v) {
    sel.querySelectorAll('option[data-temp]').forEach((o) => o.value !== v && o.remove());
    if (![...sel.options].some((o) => o.value === v)) {
      const o = new Option(v, v);
      o.dataset.temp = '1';
      sel.add(o);
    }
    sel.value = v;
  }

  function updateToolbarState() {
    const cmds = ['bold', 'italic', 'underline', 'strikeThrough', 'insertUnorderedList', 'insertOrderedList',
      'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull'];
    for (const cmd of cmds) {
      const btn = document.querySelector(`#toolbar [data-cmd="${cmd}"]`);
      let on = false;
      try { on = document.queryCommandState(cmd); } catch (_) { /* desteklenmiyor */ }
      if (btn) btn.classList.toggle('active', on);
    }
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    let n = sel.getRangeAt(0).startContainer;
    if (n.nodeType === 3) n = n.parentElement;
    if (!n || !ed.contains(n)) return;
    const cs = getComputedStyle(n);
    setSelectValue(fontFamilySel, cs.fontFamily.split(',')[0].replace(/["']/g, '').trim());
    setSelectValue(fontSizeSel, String(Math.round(parseFloat(cs.fontSize) * 1.5) / 2));
    const block = n.closest('h1,h2,h3,p,li,div');
    blockSel.value = block && /^H[123]$/.test(block.tagName) ? block.tagName.toLowerCase() : 'p';
    if (block) {
      const ratio = parseFloat(getComputedStyle(block).lineHeight) / parseFloat(getComputedStyle(block).fontSize);
      const opts = [...lineHeightSel.options].map((o) => parseFloat(o.value));
      const near = opts.reduce((a, b) => (Math.abs(b - ratio) < Math.abs(a - ratio) ? b : a), opts[0]);
      lineHeightSel.value = String(near);
    }
  }
  app.updateToolbarState = updateToolbarState;

  fontFamilySel.addEventListener('change', () => app.exec('fontName', fontFamilySel.value));
  fontSizeSel.addEventListener('change', () => app.setFontSize(fontSizeSel.value));
  blockSel.addEventListener('change', () => app.setBlock(blockSel.value));
  lineHeightSel.addEventListener('change', () => app.setLineHeight(lineHeightSel.value));
  const colorInput = (id, barId, cmd) => {
    const inp = $(id);
    const bar = $(barId);
    bar.style.background = inp.value;
    inp.addEventListener('input', () => (bar.style.background = inp.value));
    inp.addEventListener('change', () => app.exec(cmd, inp.value));
    // Etikete tıklamak odağı alır ama seçim saklı olduğu için sorun olmaz
  };
  colorInput('foreColor', 'foreBar', 'foreColor');
  colorInput('backColor', 'backBar', 'hiliteColor');

  // ---------- Girdi ----------
  function ensureContent() {
    if (!ed.firstChild || (!ed.firstElementChild && !ed.textContent)) {
      ed.innerHTML = '<p><br></p>';
      const r = document.createRange();
      r.setStart(ed.firstChild, 0);
      r.collapse(true);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
      return;
    }
    // Üst düzeyde kalmış çıplak metin/satır içi öğeleri paragrafa al (imleç korunur)
    const isLoose = (n) =>
      (n.nodeType === 3 && n.nodeValue.length) || (n.nodeType === 1 && !/^(P|H[1-6]|UL|OL|DIV|BLOCKQUOTE|PRE)$/.test(n.tagName));
    if (![...ed.childNodes].some(isLoose)) return;
    const sel = window.getSelection();
    const r = sel.rangeCount ? sel.getRangeAt(0) : null;
    const saved = r && r.startContainer !== ed ? [r.startContainer, r.startOffset, r.endContainer, r.endOffset] : null;
    let p = null;
    for (const n of [...ed.childNodes]) {
      if (isLoose(n)) {
        if (!p) {
          p = document.createElement('p');
          ed.insertBefore(p, n);
        }
        p.appendChild(n);
      } else p = null;
    }
    if (saved && ed.contains(saved[0]) && ed.contains(saved[2])) {
      const nr = document.createRange();
      nr.setStart(saved[0], saved[1]);
      nr.setEnd(saved[2], saved[3]);
      sel.removeAllRanges();
      sel.addRange(nr);
    }
  }

  ed.addEventListener('input', (e) => {
    fixFontTags();
    ensureContent();
    app.scheduleLayout();
    const t = e.inputType || '';
    const kind = /^insert(Text|CompositionText|ReplacementText)$/.test(t) ? 'typing' : /^delete/.test(t) ? 'deleting' : 'edit';
    app.commit(kind);
    if (/^format/.test(t)) updateToolbarState(); // Ctrl+B/I/U sonrası düğme durumları
  });

  // Tarayıcının kendi geri alması yerine bizimki (resim hareketleri de dahil)
  ed.addEventListener('beforeinput', (e) => {
    if (e.inputType === 'historyUndo' || e.inputType === 'historyRedo') {
      e.preventDefault();
      e.inputType === 'historyUndo' ? app.undo() : app.redo();
    }
  });

  // Sayfa sonu: sonraki metin yeni sayfadan başlar (yüksekliğini core.js ayarlar)
  app.insertPageBreak = function () {
    app.clearSelection && app.clearSelection();
    focusEditor();
    const known = new Set(ed.querySelectorAll('.pb'));
    document.execCommand('insertHTML', false, '<div class="pb" contenteditable="false"></div>');
    const pb = [...ed.querySelectorAll('.pb')].find((p) => !known.has(p));
    const sel = window.getSelection();
    if (pb && !pb.nextElementSibling) pb.after(Object.assign(document.createElement('p'), { innerHTML: '<br>' }));
    // İmleci sayfa sonundan sonraki paragrafın başına taşı
    if (pb && pb.nextElementSibling) {
      const r = document.createRange();
      r.setStart(pb.nextElementSibling, 0);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
    }
    afterFormat();
  };

  ed.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      app.insertPageBreak();
      return;
    }
    if (e.key === 'Tab' && !e.ctrlKey && !e.altKey && !e.metaKey) {
      e.preventDefault();
      if (selectedBlocks().some((b) => b.tagName === 'LI')) {
        document.execCommand(e.shiftKey ? 'outdent' : 'indent');
      } else if (!e.shiftKey) document.execCommand('insertText', false, '\t');
    }
  });

  // ---------- Yapıştırma: Word/web içeriğini sadeleştir ----------
  const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  const DROP = new Set(['SCRIPT', 'STYLE', 'META', 'LINK', 'TITLE', 'HEAD', 'XML', 'NOSCRIPT', 'TEMPLATE', 'IFRAME',
    'OBJECT', 'EMBED', 'SVG', 'CANVAS', 'VIDEO', 'AUDIO', 'BUTTON', 'INPUT', 'SELECT', 'TEXTAREA']);
  const BLOCKS = { P: 'p', DIV: 'p', H1: 'h1', H2: 'h2', H3: 'h3', H4: 'h3', H5: 'h3', H6: 'h3', BLOCKQUOTE: 'p', PRE: 'p',
    TR: 'p', DT: 'p', DD: 'p', ADDRESS: 'p', FIGCAPTION: 'p', CENTER: 'p', CAPTION: 'p', UL: 'ul', OL: 'ol', LI: 'li' };
  const INLINE = { B: 'b', STRONG: 'b', I: 'i', EM: 'i', U: 'u', INS: 'u', S: 's', STRIKE: 's', DEL: 's', SUB: 'sub', SUP: 'sup',
    SPAN: 'span', FONT: 'span', A: 'span', MARK: 'span', CODE: 'span', SMALL: 'span', BIG: 'span', ABBR: 'span', LABEL: 'span' };

  function copyInlineStyle(src, el) {
    const st = src.style;
    const out = el.style;
    if (/bold|[6-9]00/.test(st.fontWeight)) out.fontWeight = 'bold';
    if (st.fontStyle === 'italic') out.fontStyle = 'italic';
    const td = st.textDecorationLine || st.textDecoration || '';
    const deco = [/underline/.test(td) && 'underline', /line-through/.test(td) && 'line-through'].filter(Boolean);
    if (deco.length) out.textDecoration = deco.join(' ');
    const color = st.color || (src.tagName === 'FONT' && src.getAttribute('color')) || '';
    if (color && !/^(windowtext|black|#000(000)?|rgb\(0, 0, 0\))$/i.test(color.trim())) out.color = color;
    const bg = st.backgroundColor || (src.tagName === 'MARK' ? 'yellow' : '');
    if (bg && !/^(transparent|white|#fff(fff)?|rgb\(255, 255, 255\)|rgba\(0, 0, 0, 0\))$/i.test(bg.trim())) out.backgroundColor = bg;
    if (/^\d+(\.\d+)?pt$/.test(st.fontSize)) out.fontSize = st.fontSize; // web'den gelen px boyutları atılır
    const fam = (st.fontFamily || src.getAttribute('face') || '').split(',')[0].replace(/["']/g, '').trim();
    if (FONTS.includes(fam)) out.fontFamily = fam;
  }

  function convertChildren(src, dst, ctx) {
    for (const n of [...src.childNodes]) convertNode(n, dst, ctx);
  }
  function convertNode(n, dst, ctx) {
    if (n.nodeType === 3) {
      const t = n.nodeValue.replace(/[ \t\r\n\f]+/g, ' ');
      if (t) dst.appendChild(document.createTextNode(t));
      return;
    }
    if (n.nodeType !== 1) return;
    const tag = n.tagName.toUpperCase();
    if (DROP.has(tag) || (tag.includes(':') && tag !== 'O:P')) return;
    if (/mso-list:\s*ignore/i.test(n.getAttribute('style') || '')) return; // Word'ün sahte madde işaretleri
    if (tag === 'IMG') {
      const src = n.getAttribute('src') || '';
      if (/^(data:image|https?:|blob:)/i.test(src)) ctx.images.push(src);
      else if (src) ctx.skipped++;
      return;
    }
    if (tag === 'BR') return void dst.appendChild(document.createElement('br'));
    if (tag === 'DIV' && n.classList.contains('pb')) {
      const pb = document.createElement('div');
      pb.className = 'pb';
      pb.contentEditable = 'false';
      return void dst.appendChild(pb);
    }
    if (tag === 'TD' || tag === 'TH') {
      convertChildren(n, dst, ctx);
      if (n.nextElementSibling) dst.appendChild(document.createTextNode('\t'));
      return;
    }
    if (BLOCKS[tag]) {
      // Word, listeleri "MsoListParagraph" sınıflı paragraf + sahte madde işareti olarak kopyalar
      const listy = BLOCKS[tag] === 'p' &&
        (/MsoListParagraph/i.test(n.getAttribute('class') || '') || /mso-list:\s*l\d/i.test(n.getAttribute('style') || ''));
      const el = document.createElement(listy ? 'li' : BLOCKS[tag]);
      if (listy) {
        const marker = [...n.querySelectorAll('span')].find((s) => /mso-list:\s*ignore/i.test(s.getAttribute('style') || ''));
        el.dataset.wordList = marker && /^\s*([0-9]+|[a-z]|[ivx]+)[.)]/i.test(marker.textContent) ? 'ol' : 'ul';
      }
      const align = n.style.textAlign || n.getAttribute('align');
      if (/^(left|center|right|justify)$/i.test(align || '')) el.style.textAlign = align.toLowerCase();
      convertChildren(n, el, ctx);
      dst.appendChild(el);
      return;
    }
    if (INLINE[tag]) {
      const el = document.createElement(INLINE[tag]);
      copyInlineStyle(n, el);
      convertChildren(n, el, ctx);
      if (!el.childNodes.length) return;
      if (el.tagName === 'SPAN' && !el.getAttribute('style')) while (el.firstChild) dst.appendChild(el.firstChild);
      else dst.appendChild(el);
      return;
    }
    convertChildren(n, dst, ctx); // bilinmeyen/şeffaf kapsayıcı (tablo, section, o:p...): yalnızca içeriği
  }

  const isBlock = (n) => n.nodeType === 1 && (/^(P|H1|H2|H3|UL|OL|LI)$/.test(n.tagName) || n.classList.contains('pb'));

  // İç içe blokları düzleştir: <p> içinde <p> olmasın, listeler temiz olsun
  function fixList(list) {
    const L = document.createElement(list.tagName.toLowerCase());
    for (const n of [...list.childNodes]) {
      if (n.nodeType === 1 && n.tagName === 'LI') {
        const li = document.createElement('li');
        for (const c of [...n.childNodes]) {
          if (c.nodeType === 1 && /^(UL|OL)$/.test(c.tagName)) li.appendChild(fixList(c));
          else if (c.nodeType === 1 && /^(P|H1|H2|H3|LI)$/.test(c.tagName)) {
            if (li.childNodes.length && li.lastChild.nodeName !== 'BR') li.appendChild(document.createElement('br'));
            while (c.firstChild) li.appendChild(c.firstChild);
          } else li.appendChild(c);
        }
        if (!li.childNodes.length) li.appendChild(document.createElement('br'));
        L.appendChild(li);
      } else if (n.nodeType === 1 && /^(UL|OL)$/.test(n.tagName)) {
        (L.lastElementChild || L.appendChild(document.createElement('li'))).appendChild(fixList(n));
      } else if (n.nodeType === 3 && !n.nodeValue.trim()) {
        /* boşluk */
      } else {
        const li = document.createElement('li');
        li.appendChild(n);
        L.appendChild(li);
      }
    }
    return L;
  }

  function flatten(root, out) {
    let cur = null;
    const loose = (n) => {
      if (!cur && n.nodeType === 3 && /^ *$/.test(n.nodeValue)) return; // bloklar arası kaynak boşluğu
      if (!cur) cur = out.appendChild(document.createElement('p'));
      cur.appendChild(n);
    };
    for (const n of [...root.childNodes]) {
      if (!isBlock(n)) {
        loose(n);
        continue;
      }
      cur = null;
      if (n.classList.contains('pb')) out.appendChild(n);
      else if (/^(UL|OL)$/.test(n.tagName)) out.appendChild(fixList(n));
      else if (n.tagName === 'LI') {
        // Tek başına gelen maddeler (ör. Word listeleri): art arda olanları aynı listede topla
        const type = n.dataset.wordList || 'ul';
        const tmp = document.createElement(type);
        tmp.appendChild(n);
        const fixed = fixList(tmp);
        const prev = out.lastElementChild;
        if (prev && prev.dataset.loose === type) while (fixed.firstChild) prev.appendChild(fixed.firstChild);
        else {
          fixed.dataset.loose = type;
          out.appendChild(fixed);
        }
      } else {
        // P/H: içinde blok varsa parçalara böl
        let part = null;
        const wasEmpty = !n.childNodes.length; // alt düğümler aşağıda taşınacağı için önceden bak
        const style = n.getAttribute('style');
        const newPart = () => {
          part = out.appendChild(document.createElement(n.tagName.toLowerCase()));
          if (style) part.setAttribute('style', style);
          return part;
        };
        for (const c of [...n.childNodes]) {
          if (isBlock(c)) {
            part = null;
            const tmp = document.createElement('div');
            tmp.appendChild(c);
            flatten(tmp, out);
          } else (part || newPart()).appendChild(c);
        }
        if (wasEmpty) newPart().appendChild(document.createElement('br'));
      }
    }
    // Bloklar arasındaki kaynak boşluklarından doğan paragrafları at; Word'ün &nbsp; ile yaptığı
    // kasıtlı boş satırları ise boş paragraf olarak koru
    out.querySelectorAll('p').forEach((p) => {
      if (p.textContent.trim() || p.querySelector('br')) return;
      if (p.textContent.includes(' ')) p.replaceChildren(document.createElement('br'));
      else p.remove();
    });
  }

  app.sanitizeHTML = function (html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const ctx = { images: [], skipped: 0 };
    const tmp = document.createElement('div');
    convertChildren(doc.body, tmp, ctx);
    const out = document.createElement('div');
    flatten(tmp, out);
    out.querySelectorAll('[data-loose],[data-word-list]').forEach((e) => {
      e.removeAttribute('data-loose');
      e.removeAttribute('data-word-list');
    });
    return { html: out.innerHTML, text: out.textContent.trim(), images: ctx.images, skipped: ctx.skipped };
  };

  function insertPlain(text) {
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    if (lines.length === 1) document.execCommand('insertText', false, text);
    else document.execCommand('insertHTML', false, lines.map((l) => `<p>${esc(l) || '<br>'}</p>`).join(''));
  }

  app.insertSanitized = function (res) {
    if (res.text || /<br>/.test(res.html)) document.execCommand('insertHTML', false, res.html);
    if (res.images.length) app.insertImageURLs && app.insertImageURLs(res.images);
    if (res.skipped) SS.toast('Word\'den gelen görseller metinle yapıştırılamaz; resimleri sürükleyip bırakın ya da tek tek yapıştırın.', 4500);
    ensureContent();
    app.scheduleLayout();
    app.commit('edit');
  };

  document.addEventListener('paste', (e) => {
    const dt = e.clipboardData;
    if (!dt) return;
    const a = document.activeElement;
    if (a && a.classList.contains('cap-text')) return; // şekil yazısı: tarayıcı düz metin olarak yapıştırır
    if (app.pasteObjects && app.pasteObjects(e)) return; // resimler (dosya/pano) nesne olarak eklenir
    const inEditor = document.activeElement === ed || ed.contains(document.activeElement);
    if (!inEditor) return;
    e.preventDefault();
    const html = dt.getData('text/html');
    if (html) app.insertSanitized(app.sanitizeHTML(html));
    else {
      insertPlain(dt.getData('text/plain'));
      ensureContent();
      app.scheduleLayout();
      app.commit('edit');
    }
  });
})();
