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
    // Editör zaten odaktaysa ve seçim içindeyse canlı seçim günceldir; saklı seçim hızlı yazımda
    // (selectionchange henüz gelmemişken) eskide kalabilir ve komutu eski imleç yerine uygulatır
    const live = window.getSelection();
    if (document.activeElement === ed && live.rangeCount && ed.contains(live.getRangeAt(0).startContainer)) return;
    ed.focus({ preventScroll: true });
    if (savedRange && ed.contains(savedRange.startContainer)) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
  }
  app.focusEditor = focusEditor;

  function afterFormat() {
    fixFontTags();
    ensureContent();
    app.markDirtyAtSelection();
    app.scheduleLayout();
    app.commit('edit');
    updateToolbarState();
  }

  app.exec = function (cmd, value = null) {
    app.clearSelection && app.clearSelection();
    focusEditor();
    app.withoutBands(() => document.execCommand(cmd, false, value));
    afterFormat();
  };

  app.afterFormat = () => afterFormat();
  app.selectedBlocks = () => selectedBlocks();
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
    app.withoutBands(() => {
      document.execCommand('styleWithCSS', false, false);
      document.execCommand('fontSize', false, '7');
      document.execCommand('styleWithCSS', false, true);
    });
    afterFormat();
  };

  // Üst/alt simge: styleWithCSS açıkken Chrome <span style="vertical-align"> üretir; sözlükteki <sup>/<sub> için kapatılır
  app.setScript = function (cmd) {
    app.clearSelection && app.clearSelection();
    focusEditor();
    app.withoutBands(() => {
      document.execCommand('styleWithCSS', false, false);
      document.execCommand(cmd);
      document.execCommand('styleWithCSS', false, true);
    });
    afterFormat();
  };

  // Ctrl+Shift+> / <: listedeki sonraki/önceki boyut; Ctrl+] / [: 1 nk büyüt/küçült (Word gibi)
  app.growFont = function (dir, step) {
    const cur = parseFloat(fontSizeSel.value) || 11;
    const sizes = [...fontSizeSel.options].filter((o) => !o.dataset.temp).map((o) => parseFloat(o.value));
    let next = step ? cur + dir : dir > 0 ? sizes.find((v) => v > cur + 0.01) : [...sizes].reverse().find((v) => v < cur - 0.01);
    if (!next) next = dir > 0 ? cur + 1 : cur - 1;
    if (next >= 1 && next <= 1638) app.setFontSize(String(next));
  };

  app.setBlock = function (tag) {
    focusEditor();
    app.withoutBands(() => document.execCommand('formatBlock', false, tag));
    afterFormat();
  };

  // v: Word'ün satır katı (1, 1.15, 1.5, 2 …); CSS değeri paragrafın yazı tipine göre hesaplanır
  app.setLineHeight = function (v) {
    focusEditor();
    selectedBlocks().forEach((b) => (b.style.lineHeight = app.cssLineHeight(b, +v)));
    afterFormat();
  };

  app.indent = function (dir) {
    focusEditor();
    const blocks = selectedBlocks();
    if (blocks.some((b) => b.tagName === 'LI')) app.withoutBands(() => document.execCommand(dir > 0 ? 'indent' : 'outdent'));
    else
      blocks.forEach((b) => {
        const next = Math.max(0, (parseFloat(getComputedStyle(b).marginLeft) || 0) + dir * 48);
        b.style.marginLeft = next ? next + 'px' : '';
        if (!b.getAttribute('style')) b.removeAttribute('style');
      });
    afterFormat();
  };

  // ---------- Araç çubuğu durumu ----------
  const trNum = (x) => String(x).replace('.', ',');
  function setSelectValue(sel, v, label = v) {
    sel.querySelectorAll('option[data-temp]').forEach((o) => o.value !== v && o.remove());
    if (![...sel.options].some((o) => o.value === v)) {
      const o = new Option(label, v);
      o.dataset.temp = '1';
      sel.add(o);
    }
    sel.value = v;
  }

  function updateToolbarState() {
    const cmds = ['bold', 'italic', 'underline', 'strikeThrough', 'subscript', 'superscript', 'insertUnorderedList', 'insertOrderedList',
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
      const bcs = getComputedStyle(block);
      lineHeightSel.querySelector('[value="before"]').text = parseFloat(bcs.marginTop) > 0.5 ? 'Paragraftan önce boşluğu kaldır' : 'Paragraftan önce boşluk ekle';
      lineHeightSel.querySelector('[value="after"]').text = parseFloat(bcs.marginBottom) > 0.5 ? 'Paragraftan sonra boşluğu kaldır' : 'Paragraftan sonra boşluk ekle';
      // Word'deki gibi gerçek değer: listede yoksa geçici seçenek olarak gösterilir (ör. "1,32", "En az 18 nk")
      const ls = app.lineSpacing(block);
      if (ls.px) setSelectValue(lineHeightSel, 'pt' + SS.round(ls.px * 0.75, 1), `En az ${trNum(SS.round(ls.px * 0.75, 1))} nk`);
      else {
        const m = SS.round(ls.multiple, 2);
        const opt = [...lineHeightSel.options].find((o) => /^[\d.]+$/.test(o.value) && Math.abs(+o.value - m) < 0.015);
        setSelectValue(lineHeightSel, opt ? opt.value : String(m), trNum(m));
      }
    }
  }
  app.updateToolbarState = updateToolbarState;

  // ---------- Paragraf ayarları (Word'ün Paragraf penceresi) ----------
  // Girinti: margin-left (Word'ün "Sol"u), margin-right, text-indent (+ ilk satır / − asılı).
  // Aralık: margin-top/bottom (nk) ve satır aralığı (core.js: lineSpacing). Yalnızca değiştirilen alanlar uygulanır.
  const U = SS.units;
  const paraDlg = $('paraDialog');
  const paraForm = paraDlg.querySelector('form').elements;
  let paraShown = null; // pencere açılırken gösterilen değerler
  const numOf = (v) => {
    const n = parseFloat(String(v).replace(',', '.'));
    return isFinite(n) ? n : null;
  };
  function paraValues(b) {
    const cs = getComputedStyle(b);
    const ti = parseFloat(cs.textIndent) || 0;
    const ls = app.lineSpacing(b);
    const m = ls.px ? null : SS.round(ls.multiple, 2);
    return {
      left: trNum(SS.round(U.pxToCm(parseFloat(cs.marginLeft) || 0), 2)),
      right: trNum(SS.round(U.pxToCm(parseFloat(cs.marginRight) || 0), 2)),
      special: ti > 0.5 ? 'first' : ti < -0.5 ? 'hanging' : 'none',
      by: trNum(SS.round(U.pxToCm(Math.abs(ti)), 2)),
      before: trNum(SS.round((parseFloat(cs.marginTop) || 0) * 0.75, 1)),
      after: trNum(SS.round((parseFloat(cs.marginBottom) || 0) * 0.75, 1)),
      rule: ls.px ? 'atLeast' : [1, 1.5, 2].includes(m) ? String(m) : 'multiple',
      lineVal: trNum(ls.px ? SS.round(ls.px * 0.75, 1) : m),
    };
  }
  app.paragraphDialog = function () {
    focusEditor();
    const blocks = selectedBlocks().filter((b) => !b.classList.contains('pb'));
    if (!blocks.length) return;
    const inList = blocks.some((b) => b.tagName === 'LI');
    paraShown = paraValues(blocks[0]);
    for (const [k, v] of Object.entries(paraShown)) paraForm[k].value = v;
    for (const k of ['left', 'right', 'special', 'by']) paraForm[k].disabled = inList;
    $('paraListNote').hidden = !inList;
    paraDlg.returnValue = '';
    paraDlg.showModal();
  };
  paraForm.rule.addEventListener('change', () => {
    const r = paraForm.rule.value;
    if (/^[\d.]+$/.test(r)) paraForm.lineVal.value = trNum(r);
    else if (r === 'atLeast' && numOf(paraForm.lineVal.value) < 6) paraForm.lineVal.value = '12';
  });
  paraDlg.addEventListener('close', () => {
    if (paraDlg.returnValue !== 'ok' || !paraShown) return;
    const f = paraForm;
    const changed = (...ks) => ks.some((k) => String(f[k].value) !== paraShown[k]);
    focusEditor();
    const blocks = selectedBlocks().filter((b) => !b.classList.contains('pb'));
    for (const b of blocks) {
      if (b.tagName !== 'LI') {
        let left = numOf(f.left.value) || 0;
        const by = numOf(f.by.value) || 0;
        // Asılı girintide ilk satır kenar boşluğunun dışına taşmasın (Word'de Ctrl+T gibi: sol = asılı)
        if (f.special.value === 'hanging' && left < by) left = by;
        if (changed('left') || (f.special.value === 'hanging' && changed('special', 'by'))) b.style.marginLeft = left > 0 ? U.cmToPx(left) + 'px' : '';
        if (changed('right')) b.style.marginRight = numOf(f.right.value) > 0 ? U.cmToPx(numOf(f.right.value)) + 'px' : '';
        if (changed('special', 'by')) b.style.textIndent = f.special.value === 'none' || !by ? '' : (f.special.value === 'hanging' ? -1 : 1) * U.cmToPx(by) + 'px';
      }
      if (changed('before') && numOf(f.before.value) !== null) b.style.marginTop = numOf(f.before.value) + 'pt';
      if (changed('after') && numOf(f.after.value) !== null) b.style.marginBottom = numOf(f.after.value) + 'pt';
      if (changed('rule', 'lineVal')) {
        const r = f.rule.value;
        const v = numOf(f.lineVal.value);
        if (r === 'atLeast') b.style.lineHeight = (v || 12) + 'pt';
        else b.style.lineHeight = app.cssLineHeight(b, r === 'multiple' ? v || 1 : +r);
      }
      if (!b.getAttribute('style')) b.removeAttribute('style');
    }
    paraShown = null;
    afterFormat();
  });

  // Word'ün "Satır ve paragraf aralığı" menüsündeki hızlı seçenekler
  function toggleParaSpace(which) {
    focusEditor();
    const blocks = selectedBlocks().filter((b) => !b.classList.contains('pb'));
    const prop = which === 'before' ? 'marginTop' : 'marginBottom';
    const has = blocks.length && blocks.every((b) => parseFloat(getComputedStyle(b)[prop]) > 0.5);
    blocks.forEach((b) => (b.style[prop] = has ? '0pt' : which === 'before' ? '12pt' : '8pt'));
    afterFormat();
  }

  fontFamilySel.addEventListener('change', () => app.exec('fontName', fontFamilySel.value));
  fontSizeSel.addEventListener('change', () => app.setFontSize(fontSizeSel.value));
  blockSel.addEventListener('change', () => app.setBlock(blockSel.value));
  lineHeightSel.addEventListener('change', () => {
    const v = lineHeightSel.value;
    if (v === 'dialog') app.paragraphDialog();
    else if (v === 'before' || v === 'after') toggleParaSpace(v);
    else if (isFinite(+v)) app.setLineHeight(v);
    updateToolbarState(); // kutu yeniden gerçek değeri göstersin
  });
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
    app.normalizeBlocks(); // <p> içine konmuş liste/başlık (Chrome'un liste komutu böyle yapar)
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
    bandsOn();
    fixFontTags();
    ensureContent();
    app.markDirtyAtSelection(); // bu sayfaların sayfa sonları (dul/öksüz satır) arka planda yenilenir
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
      return;
    }
    // Birden çok paragrafa yayılan seçimde yerel düzenleme (silme, üzerine yazma, biçim) şeritsiz yapılır;
    // şeritler 'input' olayında (olmazsa zamanlayıcıyla) geri gelir. Tek paragraf içi yazma etkilenmez.
    const sel = window.getSelection();
    if (bandsOff || app.els.flow.classList.contains('nobands') || !sel.rangeCount || sel.isCollapsed) return;
    const r = sel.getRangeAt(0);
    const blockOf = (n) => (n.nodeType === 1 ? n : n.parentElement).closest(BLOCK_SEL);
    if (blockOf(r.startContainer) === blockOf(r.endContainer)) return;
    bandsOff = { top: app.els.workspace.scrollTop };
    app.els.flow.classList.add('nobands');
    setTimeout(bandsOn, 0);
  });
  let bandsOff = null;
  function bandsOn() {
    if (!bandsOff) return;
    app.els.flow.classList.remove('nobands');
    app.els.workspace.scrollTop = bandsOff.top;
    bandsOff = null;
  }

  // Sayfa sonu: sonraki metin yeni sayfadan başlar (yüksekliğini core.js ayarlar)
  app.insertPageBreak = function () {
    app.clearSelection && app.clearSelection();
    focusEditor();
    const known = new Set(ed.querySelectorAll('.pb'));
    app.withoutBands(() => document.execCommand('insertHTML', false, '<div class="pb" contenteditable="false"></div>'));
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

  // ---------- Paragraf başında Backspace (Word gibi) ----------
  // Sırayla: madde işareti/numara kalkar (metin yerinde, girintili paragraf olur) → ilk satır girintisi
  // kalkar → sol girinti bir durak (1,27 cm) azalır → önceki sayfa sonu silinir. Hiçbiri yoksa tarayıcı
  // paragrafı öncekiyle birleştirir.
  function caretBlockAtStart() {
    const sel = window.getSelection();
    if (!sel.rangeCount || !sel.isCollapsed) return null;
    const r = sel.getRangeAt(0);
    const n = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    const block = n && n.closest('p,h1,h2,h3,h4,h5,h6,li');
    if (!block || !ed.contains(block)) return null;
    const pre = document.createRange();
    pre.setStart(block, 0);
    pre.setEnd(r.startContainer, r.startOffset);
    const f = pre.cloneContents();
    return f.textContent.length || f.querySelector('br') ? null : block;
  }
  function pageBreakBefore(block) {
    let top = block;
    while (top.parentElement !== ed && !top.previousElementSibling) top = top.parentElement;
    const prev = top.parentElement === ed && top.previousElementSibling;
    return prev && prev.classList.contains('pb') ? prev : null;
  }
  // Üst düzey madde → girintili paragraf; sonraki maddeler ve alt listeler yeni listede, numara devam eder
  function unlistItem(li) {
    const list = li.parentElement;
    if (list.parentElement !== ed) {
      // alt düzey madde: bir üst düzeye çıkar (Shift+Tab gibi)
      app.withoutBands(() => document.execCommand('outdent'));
      return;
    }
    const items = [...list.children].filter((c) => c.tagName === 'LI');
    const p = document.createElement('p');
    if (li.getAttribute('style')) p.setAttribute('style', li.getAttribute('style'));
    p.style.marginLeft = '48px'; // madde metninin yeri (css: .editor li)
    const nested = [...li.children].filter((c) => /^(UL|OL)$/.test(c.tagName));
    for (const c of [...li.childNodes]) if (!nested.includes(c)) p.appendChild(c);
    if (!p.firstChild) p.appendChild(document.createElement('br'));
    const tail = list.cloneNode(false); // alt listeler (düzeyi korunur) ve sonraki maddeler
    nested.forEach((n) => tail.appendChild(n));
    while (li.nextSibling) tail.appendChild(li.nextSibling);
    if (list.tagName === 'OL' && tail.querySelector(':scope > li')) tail.setAttribute('start', (parseInt(list.getAttribute('start'), 10) || 1) + items.indexOf(li));
    li.remove();
    list.after(p);
    if (tail.children.length) p.after(tail);
    if (!list.querySelector('li')) list.remove();
    const r = document.createRange();
    r.setStart(p, 0);
    r.collapse(true);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
  }
  function backspaceAtStart() {
    const block = caretBlockAtStart();
    if (!block) return false;
    const cs = getComputedStyle(block);
    if (block.tagName === 'LI') unlistItem(block);
    else if (parseFloat(cs.textIndent) > 0.5) block.style.textIndent = '';
    else if (parseFloat(cs.marginLeft) > 0.5) {
      const v = Math.max(0, parseFloat(cs.marginLeft) - 48);
      block.style.marginLeft = v > 0.5 ? v + 'px' : '';
    } else {
      const pb = pageBreakBefore(block);
      if (!pb) return false;
      pb.remove();
    }
    if (block.isConnected && !block.getAttribute('style')) block.removeAttribute('style');
    afterFormat();
    return true;
  }

  ed.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      app.insertPageBreak();
      return;
    }
    if (e.key === 'Backspace' && !e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey && backspaceAtStart()) {
      e.preventDefault();
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
      // Word, listeleri "MsoListParagraph" sınıflı paragraf + sahte madde işareti olarak kopyalar;
      // düzey "mso-list:l0 level2" içinde yazar
      const style = n.getAttribute('style') || '';
      const listy = BLOCKS[tag] === 'p' && (/MsoListParagraph/i.test(n.getAttribute('class') || '') || /mso-list:\s*l\d/i.test(style));
      const el = document.createElement(listy ? 'li' : BLOCKS[tag]);
      if (listy) {
        const marker = [...n.querySelectorAll('span')].find((s) => /mso-list:\s*ignore/i.test(s.getAttribute('style') || ''));
        const num = marker && /^\s*([0-9]+|[a-z]|[ivx]+)[.)]/i.exec(marker.textContent);
        el.dataset.wordList = num ? 'ol' : 'ul';
        el.dataset.wordLevel = (/mso-list:\s*l\d+\s+level(\d+)/i.exec(style) || [0, 1])[1];
        if (num && /^\d+$/.test(num[1])) el.dataset.wordStart = num[1];
      }
      if (tag === 'OL' && +n.getAttribute('start') > 1) el.setAttribute('start', n.getAttribute('start'));
      const st = n.style;
      const align = st.textAlign || n.getAttribute('align');
      if (/^(left|center|right|justify)$/i.test(align || '')) el.style.textAlign = align.toLowerCase();
      // Word'ün paragraf ayarları (pt/cm birimli olanlar; web sayfalarının px değerleri alınmaz)
      const wordLen = (v) => /^-?[\d.]+(pt|cm|mm|in)$/.test(v || '');
      if (!listy && el.tagName !== 'LI') {
        if (wordLen(st.textIndent)) el.style.textIndent = st.textIndent;
        if (wordLen(st.marginLeft) && parseFloat(st.marginLeft) > 0) el.style.marginLeft = st.marginLeft;
        if (wordLen(st.marginRight) && parseFloat(st.marginRight) > 0) el.style.marginRight = st.marginRight;
        if (wordLen(st.marginTop)) el.style.marginTop = st.marginTop;
        if (wordLen(st.marginBottom)) el.style.marginBottom = st.marginBottom;
      }
      convertChildren(n, el, ctx);
      // Satır aralığı: Word'ün "%150"si 1,5 satırdır (yazı tipine göre çevrilir); "18pt" en az
      if (/^[\d.]+%$/.test(st.lineHeight)) {
        const f = el.querySelector('[style*="font-family"]');
        el.style.lineHeight = String(+((parseFloat(st.lineHeight) / 100) * SS.lineFactor(f ? f.style.fontFamily.split(',')[0].replace(/["']/g, '') : 'Calibri')).toFixed(4));
      } else if (/^[\d.]+pt$/.test(st.lineHeight)) el.style.lineHeight = st.lineHeight;
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
    if (list.getAttribute('start')) L.setAttribute('start', list.getAttribute('start'));
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
    let run = null; // art arda gelen tek başına maddeler: { root, stack: [düzey 1 listesi, düzey 2 listesi…] }
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
        // Tek başına gelen maddeler (ör. Word listeleri): art arda olanları düzeylerine göre iç içe listelerde topla
        const type = n.dataset.wordList || 'ul';
        const level = SS.clamp(+n.dataset.wordLevel || 1, 1, 9);
        const start = +n.dataset.wordStart || 0;
        const tmp = document.createElement(type);
        tmp.appendChild(n);
        const li = fixList(tmp).firstElementChild;
        const newList = (parent) => {
          const L = parent.appendChild(document.createElement(type));
          if (type === 'ol' && start > 1) L.setAttribute('start', start);
          return L;
        };
        let st = run && out.lastElementChild === run.root ? run.stack : null;
        if (!st) {
          const first = newList(out);
          run = { root: first, stack: (st = [first]) };
        }
        while (st.length > level) st.pop();
        while (st.length < level) {
          const parent = st[st.length - 1];
          st.push(newList(parent.lastElementChild || parent.appendChild(document.createElement('li'))));
        }
        let L = st[st.length - 1];
        if (L.tagName.toLowerCase() !== type) {
          // aynı düzeyde tür değişti (ör. madde → numara): yeni liste
          L = newList(L.parentElement);
          st[st.length - 1] = L;
          if (st.length === 1) run.root = L;
        }
        L.appendChild(li);
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
    out.querySelectorAll('[data-word-list],[data-word-level],[data-word-start]').forEach((e) => {
      e.removeAttribute('data-word-list');
      e.removeAttribute('data-word-level');
      e.removeAttribute('data-word-start');
    });
    return { html: out.innerHTML, text: out.textContent.trim(), images: ctx.images, skipped: ctx.skipped };
  };

  function insertPlain(text) {
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    app.withoutBands(() => {
      if (lines.length === 1) document.execCommand('insertText', false, text);
      else document.execCommand('insertHTML', false, lines.map((l) => `<p>${esc(l) || '<br>'}</p>`).join(''));
    });
    revealCaretSoon();
  }

  // Yapıştırmadan sonra imleç görünür olsun (düzen şeritlerle yeniden kurulduktan sonra)
  function revealCaretSoon() {
    requestAnimationFrame(() => {
      const sel = window.getSelection();
      if (!sel.rangeCount || !ed.contains(sel.anchorNode)) return;
      const r = sel.getRangeAt(0).cloneRange();
      r.collapse(false);
      const host = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      const rc = r.getClientRects()[0] || host.getBoundingClientRect();
      const ws = app.els.workspace;
      const wr = ws.getBoundingClientRect();
      if (rc.bottom > wr.bottom - 20) ws.scrollTop += rc.bottom - wr.bottom + 60;
      else if (rc.top < wr.top + 20) ws.scrollTop -= wr.top - rc.top + 60;
    });
  }

  app.insertSanitized = function (res) {
    if (res.text || /<br>/.test(res.html)) app.withoutBands(() => document.execCommand('insertHTML', false, res.html));
    if (res.images.length) app.insertImageURLs && app.insertImageURLs(res.images);
    if (res.skipped) SS.toast('Word\'den gelen görseller metinle yapıştırılamaz; resimleri sürükleyip bırakın ya da tek tek yapıştırın.', 4500);
    ensureContent();
    app.scheduleLayout();
    revealCaretSoon();
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
