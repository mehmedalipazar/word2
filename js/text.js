/* Metin düzenleme: biçimlendirme komutları, yapıştırma temizliği, araç çubuğu durumu */
(function () {
  'use strict';
  const SS = window.SS;
  const app = SS.app;
  const ed = app.els.editor;
  const $ = (id) => document.getElementById(id);
  const fontFamilySel = $('fontFamily');
  const fontSizeSel = $('fontSize');
  const famText = $('fontFamilyText'); // yazılabilir kutuların metin alanları (combo)
  const sizeText = $('fontSizeText');
  const blockSel = $('blockStyle');
  const lineHeightSel = $('lineHeight');
  const headNumSel = $('headingNum');

  document.execCommand('defaultParagraphSeparator', false, 'p');
  document.execCommand('styleWithCSS', false, true);

  const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,li,div,blockquote,pre';
  // Windows/Office'in ve Linux'un yaygın yazı tipleri. Yapıştırmada bu adlar korunur (yüklü olmasalar da Word'e
  // adıyla gider); açılır listede yalnızca bu bilgisayarda yüklü olanlar, kendi görünümleriyle yer alır.
  const FONTS = ['Aptos', 'Arial', 'Arial Black', 'Arial Narrow', 'Bahnschrift', 'Book Antiqua', 'Bookman Old Style', 'Caladea',
    'Calibri', 'Calibri Light', 'Cambria', 'Candara', 'Carlito', 'Century', 'Century Gothic', 'Comic Sans MS', 'Consolas',
    'Constantia', 'Corbel', 'Courier New', 'DejaVu Sans', 'DejaVu Serif', 'Franklin Gothic Book', 'Franklin Gothic Medium',
    'Garamond', 'Georgia', 'Gill Sans MT', 'Impact', 'Liberation Mono', 'Liberation Sans', 'Liberation Serif', 'Lucida Console',
    'Lucida Sans Unicode', 'Noto Sans', 'Noto Serif', 'Palatino Linotype', 'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold',
    'Sitka Text', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
  // Yüklü mü: yazı tipiyle ölçülen genişlik, üç yedek aileden en az biriyle ölçülenden farklıysa yüklüdür
  const fontCtx = document.createElement('canvas').getContext('2d');
  function fontInstalled(name) {
    const probe = 'mmmmmmmmmmlliWWwwıİşğ0123456789';
    return ['monospace', 'serif', 'sans-serif'].some((fb) => {
      fontCtx.font = `72px ${fb}`;
      const w = fontCtx.measureText(probe).width;
      fontCtx.font = `72px "${name}", ${fb}`;
      return fontCtx.measureText(probe).width !== w;
    });
  }
  const INSTALLED = FONTS.filter(fontInstalled);
  for (const f of INSTALLED) {
    const o = new Option(f, f);
    o.style.fontFamily = `"${f}"`;
    fontFamilySel.add(o);
  }

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
    if (!dragMove) app.commit('edit'); // sürükle-bırakla taşımanın silme yarısı ayrı adım olmaz
    updateToolbarState();
  }
  // Sürükle-bırakla taşıma Word'deki gibi tek geri alma adımı: kaynağın silinmesi (deleteByDrag) kaydedilmez,
  // bırakma (insertFromDrop) ikisini birlikte kaydeder; bırakma editörün dışına olduysa dragend kaydeder
  let dragMove = false;
  document.addEventListener('dragend', () => {
    if (!dragMove) return;
    dragMove = false;
    app.commit('edit');
  });

  app.exec = function (cmd, value = null) {
    app.clearSelection && app.clearSelection();
    focusEditor();
    if (/^insert(Ordered|Unordered)List$/.test(cmd) && removeList(cmd === 'insertOrderedList' ? 'OL' : 'UL')) return afterFormat();
    app.withoutBands(() => document.execCommand(cmd, false, value));
    afterFormat();
  };

  // ---------- Numarayı / madde işaretini kaldırma (Word gibi) ----------
  // Chrome listeyi kaldırırken maddeleri <br> ile tek paragrafta birleştiriyor, alt düzeyi numaralı
  // bırakıyordu. Seçili maddelerin hepsi bu türden bir listedeyse: her seçili madde (alt düzeyler dahil)
  // ayrı bir Normal paragraf olur; seçilmeyenler kendi düzeylerinde liste olarak kalır ve Word'deki gibi
  // yeniden numaralanır.
  const ownRange = (li) => {
    const r = document.createRange();
    r.selectNodeContents(li);
    const sub = li.querySelector(':scope > ul, :scope > ol');
    if (sub) r.setEndBefore(sub);
    return r;
  };
  function selectedItems() {
    const sel = window.getSelection();
    if (!sel.rangeCount) return [];
    const r = sel.getRangeAt(0);
    return [...ed.querySelectorAll('li')].filter((li) => {
      const o = ownRange(li);
      return r.compareBoundaryPoints(Range.START_TO_END, o) >= 0 && r.compareBoundaryPoints(Range.END_TO_START, o) <= 0;
    });
  }
  function removeList(type) {
    const items = selectedItems();
    const blocks = selectedBlocks().filter((b) => !b.classList.contains('pb'));
    if (!items.length || blocks.some((b) => b.tagName !== 'LI' && !b.closest('li')) || items.some((li) => li.parentElement.tagName !== type))
      return false;
    const chosen = new Set(items);
    const pins = app.pinSelection();
    const tops = new Set(items.map((li) => {
      let t = li.parentElement;
      while (t.parentElement !== ed) t = t.parentElement;
      return t;
    }));
    app.withoutBands(() => tops.forEach((L) => rebuildList(L, chosen)));
    app.unpinSelection(pins);
    return true;
  }
  function rebuildList(L, chosen) {
    const items = []; // belge sırasıyla { li, depth, list }
    const walk = (list, depth) => {
      for (const c of [...list.children]) {
        if (c.tagName === 'LI') {
          items.push({ li: c, depth, list });
          for (const sub of [...c.children]) if (/^(UL|OL)$/.test(sub.tagName)) walk(sub, depth + 1);
        } else if (/^(UL|OL)$/.test(c.tagName)) walk(c, depth + 1); // Chrome'un girinti biçimi
      }
    };
    walk(L, 0);
    const out = [];
    const kept = new Map(); // özgün liste -> kalan madde sayısı (numara sürsün)
    let stack = []; // açık listeler: { el, src }
    const own = (li, into) => {
      for (const c of [...li.childNodes]) if (!(c.nodeType === 1 && /^(UL|OL)$/.test(c.tagName))) into.appendChild(c);
      if (!into.firstChild) into.appendChild(document.createElement('br'));
      if (li.getAttribute('style')) into.setAttribute('style', li.getAttribute('style'));
      return into;
    };
    for (const it of items) {
      if (chosen.has(it.li)) {
        out.push(own(it.li, document.createElement('p')));
        stack = [];
        continue;
      }
      while (stack.length > it.depth + 1) stack.pop();
      if (stack.length === it.depth + 1 && stack[it.depth].src !== it.list) stack.pop();
      while (stack.length < it.depth + 1) {
        // üst düzey madde paragrafa döndüyse alt liste listenin içine doğrudan (Chrome biçimi) açılır
        const d = stack.length;
        const src = d === it.depth ? it.list : items.find((x) => x.depth === d && x.list.contains(it.li))?.list || it.list;
        const el = document.createElement(src.tagName.toLowerCase());
        const n = kept.get(src) || 0;
        const start = (parseInt(src.getAttribute('start'), 10) || 1) + n;
        if (src.tagName === 'OL' && start > 1) el.setAttribute('start', start);
        const parent = stack[stack.length - 1];
        if (!parent) out.push(el);
        else (parent.el.lastElementChild && parent.el.lastElementChild.tagName === 'LI' ? parent.el.lastElementChild : parent.el).appendChild(el);
        stack.push({ el, src });
      }
      stack[it.depth].el.appendChild(own(it.li, document.createElement('li')));
      kept.set(it.list, (kept.get(it.list) || 0) + 1);
    }
    L.replaceWith(...out);
  }

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

  // Word gibi: paragraf stili uygulanınca paragrafın tamamını kaplayan doğrudan yazı tipi, punto ve renk kalkar;
  // paragraf stilin yazı tipini alır (ör. Ctrl+A ile TNR 12 yapılmış metnin ardından eklenen başlık Başlık 1 olur).
  // Kalın/italik gibi vurgular ve paragraf biçimi (girinti, aralık, hizalama) kalır.
  function dropWholeParagraphFont(block) {
    const all = block.textContent.replace(/\s+/g, '');
    const props = ['fontFamily', 'fontSize', 'color'];
    props.forEach((p) => (block.style[p] = ''));
    if (!block.getAttribute('style')) block.removeAttribute('style');
    block.querySelectorAll('span, font').forEach((el) => {
      if (el.closest('ul, ol') !== block.closest('ul, ol')) return; // alt listenin metni ayrı paragraftır
      if (el.textContent.replace(/\s+/g, '') !== all) return;
      props.forEach((p) => (el.style[p] = ''));
      ['face', 'size', 'color'].forEach((a) => el.removeAttribute(a));
      if (!el.getAttribute('style')) el.removeAttribute('style');
      if (!el.attributes.length) el.replaceWith(...el.childNodes);
    });
  }
  app.setBlock = function (tag) {
    focusEditor();
    app.withoutBands(() => document.execCommand('formatBlock', false, tag));
    const pins = app.pinSelection();
    selectedBlocks().forEach((b) => {
      dropWholeParagraphFont(b);
      b.removeAttribute('data-num'); // stil uygulanınca stilin numarası gelir (numarasız başlık yeniden numaralanır)
    });
    app.unpinSelection(pins);
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
    if (blocks.some((b) => b.tagName === 'LI')) {
      const known = new Set(ed.querySelectorAll('span'));
      app.withoutBands(() => document.execCommand(dir > 0 ? 'indent' : 'outdent'));
      // Chrome üst düzeye çıkardığı maddenin metnini aynı puntoyla span'a sarıyor; stil sonradan değişince eski
      // punto kalmasın diye bu yeni ve gereksiz span'lar açılır
      const extra = [...ed.querySelectorAll('span[style]')].filter((s) => !known.has(s) && s.style.length === 1 && s.style.fontSize &&
        getComputedStyle(s).fontSize === getComputedStyle(s.parentElement).fontSize);
      if (extra.length) {
        const pins = app.pinSelection();
        extra.forEach((s) => s.replaceWith(...s.childNodes));
        app.unpinSelection(pins);
      }
    } else blocks.forEach((b) => app.setIndent(b, 'marginLeft', Math.max(0, (parseFloat(getComputedStyle(b).marginLeft) || 0) + dir * 48)));
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
    const N = syncHeadingNum();
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    let n = sel.getRangeAt(0).startContainer;
    if (n.nodeType === 3) n = n.parentElement;
    if (!n || !ed.contains(n)) return;
    const cs = getComputedStyle(n);
    setSelectValue(fontFamilySel, cs.fontFamily.split(',')[0].replace(/["']/g, '').trim());
    const pt = String(Math.round(parseFloat(cs.fontSize) * 1.5) / 2);
    setSelectValue(fontSizeSel, pt, trNum(pt));
    syncCombos();
    const block = n.closest('h1,h2,h3,p,li,div');
    blockSel.value = block && /^H[123]$/.test(block.tagName) ? block.tagName.toLowerCase() : 'p';
    const skip = headNumSel.querySelector('[value="skip"]');
    skip.disabled = !(N && block && /^H[123]$/.test(block.tagName) && N[block.tagName[1] - 1]);
    skip.text = !skip.disabled && block.getAttribute('data-num') === '0' ? 'Bu başlığı numaralandır' : 'Bu başlığı numaralandırma';
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
  // Belgenin başlık numaralandırması: hazır biçimlerden değilse (ör. Word'den gelen) örneğiyle gösterilir
  function syncHeadingNum() {
    const N = app.state.styles.num;
    const preset = N ? Object.keys(app.HEADING_NUMS).find((k) => JSON.stringify(app.HEADING_NUMS[k]) === JSON.stringify(N)) : '';
    const sample = app.numSample(N);
    setSelectValue(headNumSel, preset ?? '~' + sample, sample);
    return N;
  }
  app.onStyles = syncHeadingNum; // stiller değişince (açma, geri alma): core.js applyStyles

  // ---------- Köprü (Ctrl+K) ----------
  // Yalnızca http, https ve mailto adresleri kabul edilir (yapıştırılan ya da dosyadan gelen diğerleri düz metin olur).
  app.safeHref = (u) => {
    const s = String(u || '').trim();
    return /^(https?:\/\/|mailto:)\S+$/i.test(s) ? s : null;
  };
  // Pencereye yazılan adres: "www.…" ya da alan adı → https://, e-posta → mailto:
  const normalizeUrl = (u) => {
    const s = String(u || '').trim();
    if (/^[^@\s/]+@[^@\s]+\.[^@\s]+$/.test(s)) return 'mailto:' + s;
    if (!/^[a-z][a-z0-9+.-]*:/i.test(s) && /^[\w-]+(\.[\w-]+)+([/?#]\S*)?$/.test(s)) return 'https://' + s;
    return app.safeHref(s);
  };
  const linkDlg = $('linkDialog');
  const linkForm = linkDlg.querySelector('form').elements;
  let linkCtx = null;
  const linkAt = (n) => {
    const a = n && (n.nodeType === 1 ? n : n.parentElement).closest('a');
    return a && ed.contains(a) ? a : null;
  };
  // Çift tıklama sözcükten sonraki boşluğu da seçer: köprü baştaki/sondaki boşluklar dışında kurulur
  function trimSpaces(r) {
    const s = r.toString();
    if (!s.trim()) return;
    let lead = s.length - s.trimStart().length;
    let trail = s.length - s.trimEnd().length;
    const nodes = [];
    const w = document.createTreeWalker(r.commonAncestorContainer, NodeFilter.SHOW_TEXT);
    for (let n = w.currentNode.nodeType === 3 ? w.currentNode : w.nextNode(); n; n = w.nextNode()) if (r.intersectsNode(n)) nodes.push(n);
    const span = (n) => [n === r.startContainer ? r.startOffset : 0, n === r.endContainer ? r.endOffset : n.length];
    for (const n of nodes) {
      if (!lead) break;
      const [from, to] = span(n);
      const k = Math.min(lead, to - from);
      lead -= k;
      if (!lead) r.setStart(n, from + k);
    }
    for (const n of nodes.reverse()) {
      if (!trail) break;
      const [from, to] = span(n);
      const k = Math.min(trail, to - from);
      trail -= k;
      if (!trail) r.setEnd(n, to - k);
    }
  }
  app.linkDialog = function () {
    focusEditor();
    const sel = window.getSelection();
    if (!sel.rangeCount || !ed.contains(sel.getRangeAt(0).startContainer)) return;
    app.trimParagraphMark();
    const r = sel.getRangeAt(0);
    trimSpaces(r);
    const a = linkAt(r.startContainer) || linkAt(r.endContainer);
    linkCtx = { range: r.cloneRange(), a, text0: a ? a.textContent : r.toString().replace(/\s+/g, ' ') };
    linkForm.text.value = linkCtx.text0;
    linkForm.url.value = a ? a.getAttribute('href') : '';
    $('linkRemove').hidden = !a;
    linkDlg.returnValue = '';
    linkDlg.showModal();
    (linkForm.text.value ? linkForm.url : linkForm.text).focus();
  };
  $('linkRemove').addEventListener('click', () => linkDlg.close('remove'));
  linkDlg.addEventListener('close', () => {
    const c = linkCtx;
    linkCtx = null;
    if (!c || !/^(ok|remove)$/.test(linkDlg.returnValue)) return;
    const sel = window.getSelection();
    ed.focus({ preventScroll: true });
    sel.removeAllRanges();
    sel.addRange(c.range);
    if (linkDlg.returnValue === 'remove') {
      if (c.a) c.a.replaceWith(...c.a.childNodes);
      return afterFormat();
    }
    const href = normalizeUrl(linkForm.url.value);
    if (!href) return SS.toast('Geçerli bir adres yazın (https://…, www.… ya da e-posta adresi).', 4000);
    const text = linkForm.text.value.replace(/\s+/g, ' ').trim();
    if (c.a) {
      c.a.setAttribute('href', href);
      if (text && text !== c.text0) c.a.textContent = text;
    } else if (!c.range.collapsed && (!text || text === c.text0)) {
      app.withoutBands(() => document.execCommand('createLink', false, href)); // seçili metin biçimiyle köprü olur
    } else {
      const a = document.createElement('a');
      a.setAttribute('href', href);
      a.textContent = text || href.replace(/^mailto:/i, '');
      c.range.deleteContents();
      c.range.insertNode(a);
      const r = document.createRange();
      r.setStartAfter(a);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
    }
    afterFormat();
  });
  // Word gibi: Ctrl+tık bağlantıyı açar; üzerine gelince adres ipucu
  ed.addEventListener('click', (e) => {
    const a = linkAt(e.target);
    if (!a || !(e.ctrlKey || e.metaKey)) return;
    e.preventDefault();
    window.open(a.getAttribute('href'), '_blank', 'noopener');
  });
  ed.addEventListener('mouseover', (e) => {
    const a = linkAt(e.target);
    if (a) ed.title = a.getAttribute('href') + '\nCtrl+tık: bağlantıyı aç';
    else if (ed.title) ed.removeAttribute('title');
  });

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
  // "Sekmeler…": Word'deki gibi paragraf ayarları uygulanır, ardından Sekmeler penceresi açılır
  $('paraTabs').addEventListener('click', () => paraDlg.close('tabs'));
  paraDlg.addEventListener('close', () => {
    if (paraDlg.returnValue === 'tabs') setTimeout(() => app.tabsDialog(), 0);
    if (!/^(ok|tabs)$/.test(paraDlg.returnValue) || !paraShown) return;
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
        if (changed('left') || (f.special.value === 'hanging' && changed('special', 'by'))) app.setIndent(b, 'marginLeft', U.cmToPx(Math.max(0, left)));
        if (changed('right')) app.setIndent(b, 'marginRight', U.cmToPx(Math.max(0, numOf(f.right.value) || 0)));
        if (changed('special', 'by')) app.setIndent(b, 'textIndent', f.special.value === 'none' ? 0 : (f.special.value === 'hanging' ? -1 : 1) * U.cmToPx(by));
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

  // ---------- Sekmeler penceresi (Word: Paragraf > Sekmeler) ----------
  // Seçili paragrafların özel sekme durakları (core.js: data-tabs). Ayarla: yazılan konumda durak (varsa değiştirir);
  // Temizle: seçili durak; Tamam: yazılı konum da ayarlanır (Word gibi) ve duraklar seçili paragraflara uygulanır.
  const tabsDlg = $('tabsDialog');
  const tabsForm = tabsDlg.querySelector('form').elements;
  const TAB_NAMES = { l: 'Sola', c: 'Ortaya', r: 'Sağa', d: 'Ondalık' };
  const LEAD_NAMES = { '': '', '.': ', dolgu . . .', '-': ', dolgu - - -', _: ', dolgu ___' };
  let tabsWork = null;
  const cmOf = (px) => trNum(SS.round(U.pxToCm(px), 2));
  function fillTabs(sel) {
    tabsForm.list.replaceChildren(...tabsWork.map((t, i) => new Option(`${cmOf(t.pos)} cm · ${TAB_NAMES[t.type]}${LEAD_NAMES[t.leader]}`, i)));
    if (sel !== undefined) tabsForm.list.value = String(sel);
  }
  function setTabFromFields() {
    const v = numOf(tabsForm.pos.value);
    if (v === null || v < 0 || v > 55) return false;
    const pos = SS.round(U.cmToPx(v), 2);
    tabsWork = tabsWork.filter((t) => Math.abs(t.pos - pos) > 0.5);
    tabsWork.push({ type: tabsForm.type.value, pos, leader: tabsForm.leader.value });
    tabsWork.sort((a, b) => a.pos - b.pos);
    fillTabs(tabsWork.findIndex((t) => t.pos === pos));
    return true;
  }
  app.tabsDialog = function () {
    focusEditor();
    const blocks = selectedBlocks().filter((b) => !b.classList.contains('pb'));
    if (!blocks.length) return;
    tabsWork = app.parseTabs(blocks[0].dataset.tabs);
    fillTabs();
    tabsForm.pos.value = '';
    tabsForm.type.value = 'l';
    tabsForm.leader.value = '';
    tabsDlg.returnValue = '';
    tabsDlg.showModal();
    tabsForm.pos.focus();
  };
  tabsForm.list.addEventListener('change', () => {
    const t = tabsWork[+tabsForm.list.value];
    if (!t) return;
    tabsForm.pos.value = cmOf(t.pos);
    tabsForm.type.value = t.type;
    tabsForm.leader.value = t.leader;
  });
  $('tabSet').addEventListener('click', () => setTabFromFields() || SS.toast('Geçerli bir konum yazın (cm, ör. 8 ya da 12,5).'));
  $('tabClear').addEventListener('click', () => {
    let i = +tabsForm.list.value;
    if (!(i >= 0)) i = tabsWork.findIndex((t) => cmOf(t.pos) === tabsForm.pos.value.trim());
    if (i < 0 || !tabsWork[i]) return;
    tabsWork.splice(i, 1);
    fillTabs();
    tabsForm.pos.value = '';
  });
  $('tabClearAll').addEventListener('click', () => {
    tabsWork = [];
    fillTabs();
  });
  tabsDlg.addEventListener('close', () => {
    const work = tabsWork;
    if (tabsDlg.returnValue !== 'ok' || !work) return void (tabsWork = null);
    if (tabsForm.pos.value.trim()) setTabFromFields(); // Word: Tamam yazılı konumu (ve türünü) de ayarlar
    const val = app.formatTabs(tabsWork);
    tabsWork = null;
    focusEditor();
    for (const b of selectedBlocks().filter((x) => !x.classList.contains('pb'))) {
      if (val) b.dataset.tabs = val;
      else b.removeAttribute('data-tabs');
    }
    app.normalizeBlocks(); // sekmeler span'a (core.js: wrapTabs)
    afterFormat();
    app.updateRuler && app.updateRuler();
  });

  // ---------- Stiller (Word: Giriş > Stiller > Değiştir) ----------
  // Normal ve Başlık 1–3'ün yazı tipi, punto, kalın/italik, renk, hizalama, önce/sonra aralığı ve satır aralığı.
  // Değişiklik o stildeki bütün paragraflara CSS'ten uygulanır (core.js: applyStyles); paragraflara ayrıca verilmiş
  // biçim korunur. Pencere dört stilin kopyasını düzenler; yalnızca değiştirilen alanlar yazılır.
  const styleDlg = $('styleDialog');
  const styleForm = styleDlg.querySelector('form').elements;
  let styleWork = null;
  let styleKey = 'p';
  let styleShown = null;
  const hexOf = (c) => {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c || '');
    return m ? '#' + [m[1], m[2], m[3]].map((v) => (+v).toString(16).padStart(2, '0')).join('') : '#000000';
  };
  const styleFields = (s) => ({ font: s.font, size: trNum(s.size), color: s.color, bold: s.bold, italic: s.italic, align: s.align,
    before: trNum(SS.round(s.before, 1)), after: trNum(SS.round(s.after, 1)), line: trNum(SS.round(s.line, 2)) });
  function fillStyleForm(v) {
    setSelectValue(styleForm.font, v.font);
    for (const k of ['size', 'color', 'align', 'before', 'after', 'line']) styleForm[k].value = v[k];
    styleForm.bold.checked = v.bold;
    styleForm.italic.checked = v.italic;
    stylePreview();
  }
  function showStyle(k) {
    styleShown = styleFields(styleWork[k]);
    fillStyleForm(styleShown);
  }
  function readStyle(k) {
    const f = styleForm;
    const s = styleWork[k];
    const v = (name) => (String(f[name].value) !== String(styleShown[name]) ? numOf(f[name].value) : null);
    if (f.font.value !== styleShown.font) s.font = f.font.value;
    if (v('size') >= 1 && v('size') <= 1638) s.size = Math.round(v('size') * 2) / 2;
    if (v('before') !== null) s.before = SS.clamp(v('before'), 0, 1584);
    if (v('after') !== null) s.after = SS.clamp(v('after'), 0, 1584);
    if (v('line') > 0) s.line = SS.clamp(v('line'), 0.06, 132);
    s.color = f.color.value.toLowerCase();
    s.bold = f.bold.checked;
    s.italic = f.italic.checked;
    s.align = f.align.value;
  }
  function stylePreview() {
    const f = styleForm;
    const pv = $('stylePreview');
    pv.style.fontFamily = app.fontStack(f.font.value);
    pv.style.fontSize = Math.min(36, numOf(f.size.value) || 11) + 'pt';
    pv.style.fontWeight = f.bold.checked ? 'bold' : 'normal';
    pv.style.fontStyle = f.italic.checked ? 'italic' : 'normal';
    pv.style.color = f.color.value;
    pv.style.textAlign = f.align.value;
  }
  app.styleDialog = function (k) {
    focusEditor();
    styleWork = JSON.parse(JSON.stringify(app.state.styles));
    styleKey = app.STYLE_KEYS.includes(k) ? k : 'p';
    styleForm.style.value = styleKey;
    if (styleForm.font.options.length < INSTALLED.length)
      for (const f of INSTALLED) {
        const o = new Option(f, f);
        o.style.fontFamily = `"${f}"`;
        styleForm.font.add(o);
      }
    showStyle(styleKey);
    styleDlg.returnValue = '';
    styleDlg.showModal();
  };
  styleForm.style.addEventListener('change', () => {
    readStyle(styleKey);
    styleKey = styleForm.style.value;
    showStyle(styleKey);
  });
  styleDlg.querySelector('form').addEventListener('input', stylePreview);
  // Word'ün "Seçimle eşleşecek şekilde güncelle"si: imlecin paragrafının biçimi alanlara gelir
  $('styleFromSel').addEventListener('click', () => {
    const r = app.getCaretRange && app.getCaretRange();
    const n = r && (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement);
    const b = n && ed.contains(n) && n.closest('p,h1,h2,h3,h4,h5,h6,li');
    if (!b) return SS.toast('Önce metinde bir paragrafa tıklayın.');
    const w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
    let t = w.nextNode();
    while (t && !t.nodeValue.trim()) t = w.nextNode();
    const cs = getComputedStyle(t ? t.parentElement : b);
    const bcs = getComputedStyle(b);
    // Satır aralığı yalnızca paragrafa ayrıca verilmişse alınır: birimsiz CSS satır yüksekliği yazı tipiyle
    // çarpıldığından, stilin yazı tipi değişince hesaplanan kat kayar (Word'de kat yazı tipinden bağımsızdır)
    const ls = b.style.lineHeight ? app.lineSpacing(b) : { px: true };
    fillStyleForm({
      font: app.lineFont(b),
      size: trNum(Math.round(parseFloat(cs.fontSize) * 1.5) / 2),
      color: hexOf(cs.color),
      bold: parseInt(cs.fontWeight, 10) >= 600,
      italic: cs.fontStyle === 'italic',
      align: { center: 'center', right: 'right', end: 'right', justify: 'justify' }[bcs.textAlign] || 'left',
      before: trNum(SS.round((parseFloat(bcs.marginTop) || 0) * 0.75, 1)),
      after: trNum(SS.round((parseFloat(bcs.marginBottom) || 0) * 0.75, 1)),
      line: ls.px ? styleForm.line.value : trNum(SS.round(ls.multiple, 2)),
    });
  });
  $('styleDefault').addEventListener('click', () => styleDlg.close('default'));
  styleDlg.addEventListener('close', () => {
    const rv = styleDlg.returnValue;
    const work = styleWork;
    styleWork = null;
    if (!work || (rv !== 'ok' && rv !== 'default')) return;
    styleWork = work;
    readStyle(styleKey);
    styleWork = null;
    const next = app.cleanStyles(work, app.state.styles);
    if (rv === 'default')
      SS.toast(app.saveDefaultStyles(next) ? 'Stiller bu belgeye uygulandı; yeni belgeler de bu stillerle açılacak.' : 'Varsayılan kaydedilemedi (tarayıcı depolaması kapalı).', 4000);
    focusEditor();
    app.setStyles(next);
    updateToolbarState();
  });

  // ---------- Biçim boyacısı (Word: Giriş > Biçim Boyacısı, Ctrl+Shift+C / Ctrl+Shift+V) ----------
  // Kopyalanan: seçimin başındaki karakter biçimi (yazı tipi, punto, kalın, italik, altı/üstü çizili, renk, vurgu,
  // üst/alt simge). İmleç tek başınaysa, paragraf işareti ya da birden çok paragraf seçiliyse paragraf biçimi de
  // (stil, hizalama, girinti, aralık, satır aralığı). Uygulama: karakter biçimi seçime (tıklanınca sözcüğe),
  // paragraf biçimi seçili paragraflara; tek geri alma adımı. Düğme: tek tık bir kez, çift tık sürekli (Esc: bitir).
  let fmt = null;
  let painting = 0; // 0 kapalı, 1 bir kez, 2 sürekli
  let batching = false; // true iken girdi olayları geçmişe yazılmaz (toplu biçim tek adım)
  const PARA_PROPS = ['textAlign', 'marginLeft', 'marginRight', 'textIndent', 'marginTop', 'marginBottom', 'lineHeight'];
  const fam1 = (f) => f.split(',')[0].replace(/["']/g, '').trim();
  const ptOf = (cs) => Math.round(parseFloat(cs.fontSize) * 1.5) / 2;
  function firstTextIn(r) {
    if (r.startContainer.nodeType === 3) return r.startContainer;
    const root = r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentNode;
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) if (r.intersectsNode(t) && t.nodeValue.trim()) return t;
    return null;
  }
  function readFormat() {
    const sel = window.getSelection();
    const r = sel.rangeCount && ed.contains(sel.getRangeAt(0).startContainer) ? sel.getRangeAt(0) : savedRange;
    if (!r || !ed.contains(r.startContainer)) return null;
    const t = firstTextIn(r);
    const el = t ? t.parentElement : r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    const block = el.closest(BLOCK_SEL);
    const cs = getComputedStyle(el);
    const c = { font: fam1(cs.fontFamily), size: ptOf(cs), bold: parseInt(cs.fontWeight, 10) >= 600, italic: cs.fontStyle === 'italic',
      underline: false, strike: false, color: hexOf(cs.color), bg: null, va: null };
    for (let x = el; x && x !== ed; x = x.parentElement) {
      const s = getComputedStyle(x);
      const d = s.textDecorationLine || '';
      if (d.includes('underline') && x.tagName !== 'A') c.underline = true;
      if (d.includes('line-through')) c.strike = true;
      if (!c.bg && x !== block && !/^(transparent|rgba\(0, 0, 0, 0\))$/.test(s.backgroundColor)) c.bg = hexOf(s.backgroundColor);
      if (!c.va && /^(SUP|SUB)$/.test(x.tagName)) {
        c.va = x.tagName === 'SUP' ? 'superscript' : 'subscript';
        c.size = ptOf(getComputedStyle(x.parentElement)); // simge küçültülmüş çizilir: asıl punto
      }
      if (x === block) break;
    }
    const withPara = sel.isCollapsed || !!paragraphMarkSelection() || selectedBlocks().length > 1;
    const p = withPara && block && ed.contains(block)
      ? { tag: block.tagName, style: Object.fromEntries(PARA_PROPS.map((k) => [k, block.style[k]])) }
      : null;
    return { c, p };
  }
  // Tıklanan sözcük (Word'de biçim boyacısıyla tıklamak sözcüğe uygular)
  function selectWordAtCaret(sel) {
    const r = sel.getRangeAt(0);
    const n = r.startContainer;
    if (n.nodeType !== 3) return;
    const s = n.nodeValue;
    const W = /[\p{L}\p{N}_'’-]/u;
    let a = r.startOffset;
    let b = r.startOffset;
    while (a > 0 && W.test(s[a - 1])) a--;
    while (b < s.length && W.test(s[b])) b++;
    if (a === b) return;
    const rr = document.createRange();
    rr.setStart(n, a);
    rr.setEnd(n, b);
    sel.removeAllRanges();
    sel.addRange(rr);
  }
  function applyFormat() {
    if (!fmt) return;
    app.clearSelection && app.clearSelection();
    focusEditor();
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    if (sel.isCollapsed) selectWordAtCaret(sel);
    const { c, p } = fmt;
    batching = true; // birçok komut, tek geri alma adımı
    try {
      app.withoutBands(() => {
        // Önce paragrafın stili: karakter biçiminden yalnızca yeni stilden farklı olan doğrudan biçim olarak kalır.
        // Paragraf biçimi (hizalama, girinti, aralık) en sonda: Chrome'un removeFormat'ı hizalamayı da siliyor.
        const blocks = () => selectedBlocks().filter((x) => !x.classList.contains('pb'));
        if (p) {
          const pins = app.pinSelection();
          for (const b of blocks()) {
            if (b.tagName === p.tag || b.tagName === 'LI' || !/^(P|H[1-3])$/.test(p.tag) || !/^(P|H[1-6]|DIV)$/.test(b.tagName)) continue;
            const n = document.createElement(p.tag);
            if (b.getAttribute('style')) n.setAttribute('style', b.getAttribute('style'));
            while (b.firstChild) n.appendChild(b.firstChild);
            b.replaceWith(n);
          }
          app.unpinSelection(pins);
        }
        const paraProps = () => {
          if (!p) return;
          for (const b of blocks()) {
            for (const k of PARA_PROPS) if (b.tagName !== 'LI' || !/^(margin|textIndent)/.test(k)) b.style[k] = p.style[k] || '';
            if (!b.getAttribute('style')) b.removeAttribute('style');
          }
        };
        if (sel.isCollapsed) return paraProps();
        const now = () => {
          const t = firstTextIn(sel.getRangeAt(0));
          return getComputedStyle(t ? t.parentElement : ed);
        };
        const exec = (cmd, v = null, css = true) => {
          document.execCommand('styleWithCSS', false, css);
          document.execCommand(cmd, false, v);
          document.execCommand('styleWithCSS', false, true);
        };
        const isOn = (cmd) => {
          try {
            return document.queryCommandState(cmd);
          } catch (_) {
            return false;
          }
        };
        exec('removeFormat');
        if (fam1(now().fontFamily).toLowerCase() !== c.font.toLowerCase()) exec('fontName', c.font);
        if (Math.abs(ptOf(now()) - c.size) > 0.01) {
          pendingSize = c.size; // fixFontTags gerçek puntoyu yazar
          exec('fontSize', '7', false);
        }
        for (const [cmd, want] of [['bold', c.bold], ['italic', c.italic], ['underline', c.underline], ['strikeThrough', c.strike]])
          if (isOn(cmd) !== want) exec(cmd);
        if (hexOf(now().color) !== c.color) exec('foreColor', c.color);
        if (c.bg) exec('hiliteColor', c.bg);
        if (c.va) exec(c.va, null, false);
        paraProps();
      });
    } finally {
      batching = false;
    }
    afterFormat();
  }
  const painterBtn = document.querySelector('#toolbar [data-cmd="formatPainter"]');
  function setPainting(mode) {
    painting = mode;
    ed.classList.toggle('painting', !!mode);
    painterBtn.classList.toggle('active', !!mode);
    $('stHint').textContent = mode ? 'Biçim boyacısı: biçimin uygulanacağı metni seçin ya da sözcüğe tıklayın' + (mode === 2 ? ' (sürekli; Esc: bitir)' : ' (Esc: vazgeç)') : '';
  }
  // Düğme: tek tık bir kez, ikinci tık kapatır; çift tık sürekli kip
  app.formatPainter = function (sticky) {
    if (painting && !sticky) return setPainting(0);
    fmt = readFormat();
    if (!fmt) return SS.toast('Önce biçimi alınacak metne tıklayın.');
    setPainting(sticky ? 2 : 1);
  };
  painterBtn.addEventListener('dblclick', () => app.formatPainter(true));
  ed.addEventListener('mouseup', () => {
    if (!painting) return;
    setTimeout(() => {
      applyFormat();
      if (painting === 1) setPainting(0);
    }, 0);
  });
  document.addEventListener('keydown', (e) => {
    if (painting && e.key === 'Escape') setPainting(0);
  });
  app.copyFormat = function () {
    fmt = readFormat();
    if (fmt) SS.toast('Biçim kopyalandı. Uygulamak için metni seçip Ctrl+Shift+V.', 2500);
  };
  app.hasFormat = () => !!fmt;
  app.pasteFormat = () => applyFormat();

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

  // Yazılabilir kutular (Word gibi): yazı tipi adı ya da punto (ör. 11,5) yazılıp Enter ile uygulanır; Esc ya da
  // başka yere tıklamak vazgeçer ve gerçek değer geri gelir. Liste, kutunun okundan (alttaki <select>) açılır.
  function syncCombos() {
    if (document.activeElement !== famText) famText.value = fontFamilySel.value;
    if (document.activeElement !== sizeText) sizeText.value = trNum(fontSizeSel.value);
  }
  function combo(input, apply) {
    input.addEventListener('focus', () => input.select());
    input.addEventListener('blur', syncCombos);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        apply(input.value.trim());
      } else if (e.key === 'Escape') {
        e.preventDefault();
        focusEditor();
      }
    });
  }
  combo(famText, (t) => {
    if (!t) return;
    const name = FONTS.find((f) => f.toLocaleLowerCase('tr-TR') === t.toLocaleLowerCase('tr-TR')) || t;
    if (!/^[\p{L}\p{N} ._-]{1,48}$/u.test(name)) return SS.toast('Geçerli bir yazı tipi adı yazın.');
    if (!fontInstalled(name)) SS.toast(`"${name}" bu bilgisayarda yüklü değil: benzer bir yazı tipiyle gösterilir, Word'e bu adla aktarılır.`, 5000);
    app.exec('fontName', name);
  });
  // Yazarken yüklü yazı tiplerinden ilk uyan tamamlanır (Word gibi); tamamlanan kısım seçili kalır
  famText.addEventListener('input', (e) => {
    if (!/^insert/.test(e.inputType || '')) return;
    const t = famText.value;
    const hit = t && INSTALLED.find((f) => f.toLocaleLowerCase('tr-TR').startsWith(t.toLocaleLowerCase('tr-TR')));
    if (!hit || hit.length === t.length) return;
    famText.value = hit;
    famText.setSelectionRange(t.length, hit.length);
  });
  combo(sizeText, (t) => {
    const v = Math.round(parseFloat(t.replace(',', '.')) * 2) / 2; // Word'deki gibi yarım nokta adımlarıyla
    if (!(v >= 1 && v <= 1638)) return SS.toast('Yazı boyutu 1 ile 1638 arasında bir sayı olmalı (ör. 11,5).');
    app.setFontSize(String(v));
  });
  syncCombos();
  blockSel.addEventListener('change', () => {
    if (blockSel.value !== 'edit') return app.setBlock(blockSel.value);
    // Pencere imlecin paragrafının stiliyle açılır; kutu yeniden o stili gösterir
    const r = savedRange;
    const n = r && (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement);
    const h = n && ed.contains(n) && n.closest('h1,h2,h3');
    blockSel.value = h ? h.tagName.toLowerCase() : 'p';
    app.styleDialog(blockSel.value);
  });
  // Başlık numaralandırması (Word: Çok Düzeyli Liste, Başlık 1–3'e bağlı): hazır biçim Başlık 1–3 stillerine
  // uygulanır (core.js: styles.num). "Bu başlığı numaralandırma" seçili başlıkları numarasız yapar ya da yeniden
  // numaralar (data-num="0"; Word: numId 0).
  headNumSel.addEventListener('change', () => {
    const v = headNumSel.value;
    focusEditor();
    if (v === 'skip') {
      const heads = selectedBlocks().filter((b) => /^H[123]$/.test(b.tagName));
      const off = heads.some((h) => h.getAttribute('data-num') !== '0');
      heads.forEach((h) => (off ? h.setAttribute('data-num', '0') : h.removeAttribute('data-num')));
      if (heads.length) afterFormat();
    } else if (v === '' || app.HEADING_NUMS[v]) app.setStyles(app.cleanStyles({ ...app.state.styles, num: v ? app.HEADING_NUMS[v] : null }));
    updateToolbarState();
  });
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
    if (!batching && !dragMove) app.commit(kind);
    if (/^format/.test(t)) updateToolbarState(); // Ctrl+B/I/U sonrası düğme durumları
    if (!batching && (t === 'insertText' || t === 'insertParagraph') && autoLink(e)) app.commit('edit'); // ayrı adım: Ctrl+Z köprüyü geri alır
  });

  // Yazarken otomatik köprü (Word: "Yazarken Otomatik Biçimlendir > İnternet ve ağ yollarını köprülerle"): boşluk,
  // sekme ya da Enter'dan sonra önceki sözcük web ya da e-posta adresiyse köprü olur. Sondaki noktalama köprüye
  // girmez. Hemen ardından Ctrl+Z yalnızca köprüyü geri alır, yazılan kalır.
  const AUTO_URL = /^(?:(?:https?:\/\/|www\.)[^\s<>"]+\.[^\s<>"]+|mailto:[^\s@<>"]+@[^\s@<>"]+|[\w.+-]+@[\w-]+(?:\.[\w-]+)+)$/i;
  function autoLink(e) {
    const sel = window.getSelection();
    if (!sel.rangeCount || !sel.isCollapsed) return false;
    let node = null;
    let end = 0;
    if (e.inputType === 'insertText') {
      if (!/^[ \t\u00a0]$/.test(e.data || '')) return false;
      const r = sel.getRangeAt(0);
      node = r.startContainer;
      end = r.startOffset - 1; // yazılan boşluğun yeri
      if (node.nodeType !== 3 || end < 1) return false;
    } else {
      // Enter: önceki paragrafın son metni
      const cur = unitOf(sel.anchorNode);
      const units = [...ed.querySelectorAll(UNIT_SEL)];
      const prev = cur && units[units.indexOf(cur) - 1];
      if (!prev) return false;
      const w = document.createTreeWalker(prev, NodeFilter.SHOW_TEXT);
      for (let t = w.nextNode(); t; t = w.nextNode()) if (t.nodeValue.length) node = t;
      if (!node) return false;
      end = node.nodeValue.length;
    }
    const m = /(\S+)$/.exec(node.nodeValue.slice(0, end));
    if (!m || node.parentElement.closest('a')) return false;
    const trail = (/[.,;:!?)\]}"'»]+$/.exec(m[1]) || [''])[0];
    const word = m[1].slice(0, m[1].length - trail.length);
    if (!AUTO_URL.test(word)) return false;
    const href = normalizeUrl(word);
    if (!href) return false;
    const r = document.createRange();
    const start = end - m[1].length;
    r.setStart(node, start);
    r.setEnd(node, start + word.length);
    const a = document.createElement('a');
    a.setAttribute('href', href);
    r.surroundContents(a);
    if (e.inputType === 'insertText' && a.nextSibling && a.nextSibling.nodeType === 3) {
      const c = document.createRange(); // imleç yazılan boşluğun ardında kalır
      c.setStart(a.nextSibling, Math.min(a.nextSibling.length, trail.length + 1));
      sel.removeAllRanges();
      sel.addRange(c);
    }
    return true;
  }

  // ---------- Paragrafın tamamını silme / üzerine yazma (Word gibi) ----------
  // Üç tıklama ya da Shift+↓ seçimi sonraki paragrafın başında (ofset 0) biter: Word'de paragraf işareti de
  // seçilmiştir ve sonraki paragraf hiç etkilenmez. Chrome ise iki paragrafı birleştirip ilkinin etiketini
  // korur; sonraki başlık "Normal" paragrafa döner, biçimi span'a taşınır. Bu yüzden:
  //  - silme/kesme: seçim bir paragrafın başından başlıyorsa seçili paragraflar bütünüyle kaldırılır;
  //  - yazma, Enter, yapıştırma ve ortadan başlayan silme: seçimin sonu önceki paragrafın sonuna çekilir.
  const UNIT_SEL = 'p,h1,h2,h3,h4,h5,h6,li';
  const unitOf = (n) => {
    const el = n && (n.nodeType === 1 ? n : n.parentElement);
    const u = el && el.closest(UNIT_SEL);
    return u && ed.contains(u) ? u : null;
  };
  const atStartOf = (u, node, off) => {
    const r = document.createRange();
    r.setStart(u, 0);
    r.setEnd(node, off);
    const f = r.cloneContents();
    return !f.textContent.length && !f.querySelector('br');
  };
  const isPB = (n) => !!n && n.nodeType === 1 && n.classList.contains('pb');
  // Seçimin sonu başında durduğu (içinden hiçbir şey seçilmemiş) paragraf ya da sayfa sonu
  function blockAtRangeEnd(r) {
    const n = r.endContainer;
    if (n.nodeType === 1 && !unitOf(n)) {
      // editör ya da liste düzeyinde, iki blok arası: sonraki blok (üç tıklama sayfa sonunun önünde biter)
      let c = n.childNodes[r.endOffset];
      while (c && c.nodeType === 1 && !c.matches(UNIT_SEL) && !isPB(c)) c = c.firstElementChild;
      return c && c.nodeType === 1 ? c : null;
    }
    const u = unitOf(n);
    return u && atStartOf(u, n, r.endOffset) ? u : null;
  }
  // B'nin hemen önündeki sayfa sonlarının ilki (B listenin ilk maddesiyse listenin önündekiler)
  function pageBreaksBefore(B) {
    let top = B;
    while (top.parentElement && top.parentElement !== ed && !top.previousElementSibling) top = top.parentElement;
    let p = top.parentElement === ed ? top.previousElementSibling : null;
    let first = null;
    for (; isPB(p); p = p.previousElementSibling) first = p;
    return first;
  }
  // Seçim sonraki paragrafın başında bitiyorsa: { A: ilk paragraf, B: sonraki paragraf ya da sayfa sonu, whole: A'nın başından mı }
  // Word'de Ctrl+Enter'ın sayfa sonu ayrı bir öğedir: seçim hemen ardındaki paragrafın başında bitse de (imleç sayfa
  // sonuna konamadığı için Shift+↓ onu atlar) sayfa sonu seçime sayılmaz, B o olur.
  function paragraphMarkSelection() {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return null;
    const r = sel.getRangeAt(0);
    if (!ed.contains(r.commonAncestorContainer) && r.commonAncestorContainer !== ed) return null;
    let B = blockAtRangeEnd(r);
    const A = unitOf(r.startContainer);
    if (!A || !B || A === B || A.contains(B) || !(A.compareDocumentPosition(B) & Node.DOCUMENT_POSITION_FOLLOWING)) return null;
    const pb = !isPB(B) && pageBreaksBefore(B);
    if (pb && A.compareDocumentPosition(pb) & Node.DOCUMENT_POSITION_FOLLOWING) B = pb;
    return { r, A, B, whole: atStartOf(A, r.startContainer, r.startOffset) };
  }
  // B'den önceki son paragraf (B bir alt listenin ilk maddesiyse üst madde)
  function unitBefore(B) {
    const units = [...ed.querySelectorAll(UNIT_SEL)];
    if (!isPB(B)) return units[units.indexOf(B) - 1];
    return units.filter((u) => u.compareDocumentPosition(B) & Node.DOCUMENT_POSITION_FOLLOWING).pop();
  }
  // Seçimin sonunu B'den önceki paragrafın sonuna çek (B ayrı kalır)
  function trimToParagraphEnd(m) {
    const prev = unitBefore(m.B);
    if (!prev) return;
    const end = document.createRange();
    end.selectNodeContents(prev);
    const sub = prev.querySelector(':scope > ul, :scope > ol');
    if (sub) end.setEndBefore(sub);
    end.collapse(false);
    m.r.setEnd(end.endContainer, end.endOffset);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(m.r);
  }
  // Word gibi sil: tüm paragraflar seçiliyse onları kaldırır (true döner); değilse seçimi düzeltip tarayıcıya bırakır
  function deleteParagraphs() {
    const m = paragraphMarkSelection();
    if (!m) return false;
    if (!m.whole) {
      trimToParagraphEnd(m);
      return false;
    }
    const rr = document.createRange();
    rr.setStartBefore(m.A);
    rr.setEndBefore(m.B);
    app.withoutBands(() => rr.deleteContents());
    ed.querySelectorAll('ul, ol').forEach((l) => !l.querySelector('li') && l.remove());
    // İmleç B'nin başına; B sayfa sonuysa (imleç konamaz) önceki paragrafın sonuna
    const c = document.createRange();
    const prev = isPB(m.B) && unitBefore(m.B);
    if (prev) {
      c.selectNodeContents(prev);
      const sub = prev.querySelector(':scope > ul, :scope > ol');
      if (sub) c.setEndBefore(sub);
      c.collapse(false);
    } else if (isPB(m.B) && m.B.nextElementSibling) c.setStart(m.B.nextElementSibling, 0);
    else c.setStart(m.B, 0);
    c.collapse(false);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(c);
    afterFormat();
    return true;
  }
  app.deleteSelection = () => deleteParagraphs() || app.withoutBands(() => document.execCommand('delete'));
  // Tam paragraf seçiliyken Enter (Word gibi): seçimin yerinde tek boş paragraf kalır, imleç orada
  function enterOverParagraphs() {
    const m = paragraphMarkSelection();
    if (!m || !m.whole || m.A.tagName === 'LI') return false;
    const rr = document.createRange();
    rr.setStartAfter(m.A);
    rr.setEndBefore(m.B);
    app.withoutBands(() => rr.deleteContents());
    ed.querySelectorAll('ul, ol').forEach((l) => !l.querySelector('li') && l.remove());
    m.A.innerHTML = '<br>';
    const c = document.createRange();
    c.setStart(m.A, 0);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(c);
    afterFormat();
    return true;
  }
  app.trimParagraphMark = () => {
    const m = paragraphMarkSelection();
    if (m) trimToParagraphEnd(m);
  };

  // Tarayıcının kendi geri alması yerine bizimki (resim hareketleri de dahil)
  ed.addEventListener('beforeinput', (e) => {
    if (e.inputType === 'historyUndo' || e.inputType === 'historyRedo') {
      e.preventDefault();
      e.inputType === 'historyUndo' ? app.undo() : app.redo();
      return;
    }
    if (e.inputType === 'deleteByDrag') dragMove = true;
    else if (e.inputType === 'insertFromDrop') dragMove = false;
    if ((/^delete/.test(e.inputType) && deleteParagraphs()) || (e.inputType === 'insertParagraph' && enterOverParagraphs())) {
      e.preventDefault();
      return;
    }
    if (/^insert(Text|ReplacementText|Paragraph|LineBreak|FromPaste|FromDrop)$/.test(e.inputType)) app.trimParagraphMark();
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
    app.normalizeBlocks(); // liste maddesinde: sayfa sonu listenin dışına çıkar, liste bölünür
    const sel = window.getSelection();
    if (pb && !pb.nextElementSibling) pb.after(Object.assign(document.createElement('p'), { innerHTML: '<br>' }));
    // İmleci sayfa sonundan sonraki paragrafın (liste ise ilk maddenin) başına taşı
    let next = pb && pb.nextElementSibling;
    while (next && !next.matches(UNIT_SEL) && next.firstElementChild) next = next.firstElementChild;
    if (next) {
      const r = document.createRange();
      r.setStart(next, 0);
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
  // Girintiyi bir adım azalt: önce ilk satır girintisi kalkar, sonra sol girinti bir durak (1,27 cm) azalır.
  // Numaralı başlığın girintisi numaralandırmanındır: azaltılmaz
  function outdentBlock(block) {
    if (app.isNumbered(block)) return false;
    const cs = getComputedStyle(block);
    if (parseFloat(cs.textIndent) > 0.5) block.style.textIndent = '';
    else if (parseFloat(cs.marginLeft) > 0.5) {
      const v = Math.max(0, parseFloat(cs.marginLeft) - 48);
      block.style.marginLeft = v > 0.5 ? v + 'px' : '';
    } else return false;
    if (!block.getAttribute('style')) block.removeAttribute('style');
    return true;
  }
  function backspaceAtStart() {
    const block = caretBlockAtStart();
    if (!block) return false;
    if (block.tagName === 'LI') unlistItem(block);
    else if (app.isNumbered(block)) block.setAttribute('data-num', '0'); // Word gibi: numaralı başlıkta önce numara kalkar
    else if (!outdentBlock(block)) {
      const pb = pageBreakBefore(block);
      if (!pb) return false;
      pb.remove();
    }
    if (block.isConnected && !block.getAttribute('style')) block.removeAttribute('style');
    afterFormat();
    return true;
  }

  // Tab (Word gibi): birden çok paragraf ya da paragrafın tamamı seçiliyse girinti/madde düzeyi artar. Maddenin
  // başında düzey iner, metnin içinde sekme karakteri girer ("Terim⇥açıklama" hizaları için). Shift+Tab maddede
  // düzeyi çıkarır, paragrafın başında girintiyi azaltır.
  function tabKey(shift) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    app.trimParagraphMark(); // üç tıkla seçilen paragrafın işareti sonraki paragrafı katmasın
    const r = sel.getRangeAt(0);
    const blocks = selectedBlocks();
    const u = unitOf(r.startContainer);
    if (!u) return;
    const li = u.closest('li');
    const unit = li && ed.contains(li) ? li : u;
    const atStart = atStartOf(unit, r.startContainer, r.startOffset);
    const whole = blocks.length > 1 || (atStart && !sel.isCollapsed && atEndOf(unit, r.endContainer, r.endOffset));
    if (whole || (li && (shift || atStart))) return app.indent(shift ? -1 : 1);
    if (shift) {
      if (atStart && outdentBlock(unit)) afterFormat();
      return;
    }
    document.execCommand('insertText', false, '\t');
  }
  // Konumdan birimin sonuna (alt listesi hariç) metin kalmıyor mu
  function atEndOf(u, node, off) {
    const rr = document.createRange();
    rr.setStart(node, off);
    rr.setEnd(u, u.childNodes.length);
    const f = rr.cloneContents();
    f.querySelectorAll('ul, ol').forEach((l) => l.remove());
    return !f.textContent.length;
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
      tabKey(e.shiftKey);
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

  // Word'ün HTML panosundaki sekme durakları, ör. "tab-stops:center 226.8pt right dotted 453.6pt" → data-tabs
  function wordTabStops(style) {
    const m = /tab-stops:\s*([^;"]+)/i.exec(style || '');
    if (!m) return '';
    const TYPE = { left: 'l', center: 'c', right: 'r', decimal: 'd' };
    const LEAD = { dotted: '.', dashed: '-', lined: '_', heavy: '_' };
    const UNIT = { pt: 96 / 72, cm: 96 / 2.54, mm: 96 / 25.4, in: 96, px: 1 };
    const out = [];
    let type = 'l';
    let leader = '';
    for (const tok of m[1].trim().toLowerCase().split(/\s+/)) {
      const len = /^(-?[\d.]+)(pt|cm|mm|in|px)$/.exec(tok);
      if (TYPE[tok]) type = TYPE[tok];
      else if (LEAD[tok]) leader = LEAD[tok];
      else if (tok === 'list' || tok === 'bar') type = null; // liste ve çubuk durakları alınmaz
      else if (len) {
        if (type) out.push({ type, pos: +len[1] * UNIT[len[2]], leader });
        type = 'l';
        leader = '';
      }
    }
    return out.length ? app.formatTabs(out) : '';
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
      const stops = wordTabStops(style); // Word'ün özel sekme durakları
      if (stops) el.dataset.tabs = stops;
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
    const tabCount = /mso-tab-count:\s*(\d+)/i.exec(n.getAttribute('style') || '');
    if (tabCount) return void dst.appendChild(document.createTextNode('\t'.repeat(Math.min(+tabCount[1], 20)))); // Word'ün sekmesi
    if (INLINE[tag]) {
      // Köprü korunur (yalnızca http/https/mailto); diğer bağlantılar düz metin olur
      const href = tag === 'A' && app.safeHref(n.getAttribute('href'));
      const el = document.createElement(href ? 'a' : INLINE[tag]);
      if (href) el.setAttribute('href', href);
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
        if (n.getAttribute('style')) li.setAttribute('style', n.getAttribute('style'));
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
        const tabs = n.getAttribute('data-tabs');
        const noNum = n.getAttribute('data-num') === '0';
        const newPart = () => {
          part = out.appendChild(document.createElement(n.tagName.toLowerCase()));
          if (style) part.setAttribute('style', style);
          if (tabs) part.setAttribute('data-tabs', tabs);
          if (noNum) part.setAttribute('data-num', '0');
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
    app.trimParagraphMark(); // seçim sonraki paragrafın başında bitiyorsa o paragraf korunur
    if (res.text || /<br>/.test(res.html)) app.withoutBands(() => document.execCommand('insertHTML', false, res.html));
    if (res.images.length) app.insertImageURLs && app.insertImageURLs(res.images);
    if (res.skipped) SS.toast('Word\'den gelen görseller metinle yapıştırılamaz; resimleri sürükleyip bırakın ya da tek tek yapıştırın.', 4500);
    ensureContent();
    app.scheduleLayout();
    revealCaretSoon();
    app.commit('edit');
  };

  // ---------- Pano: uygulama içi kopyala / kes / yapıştır ----------
  // Tarayıcının panoya yazdığı HTML paragraf biçimini px cinsinden ve hesaplanmış stillerle karışık verir;
  // Word/web temizliği (sanitizeHTML) bunları atar ve yazı tipi, punto, girinti, satır aralığı kaybolurdu.
  // Bu yüzden seçimin kendi HTML'imiz (sözlükteki stilleriyle) işaretli olarak panoya konur; yapıştırırken
  // işaretli içerik biçimiyle alınır. Seçim sonraki paragrafın başında bitiyorsa (paragraf işareti seçili)
  // içerik "tam paragraf"tır: Word'deki gibi ayrı paragraf olarak girer, boş paragraf oluşmaz.
  const CLIP_ATTR = 'data-serbestsayfa';
  function copyHTML() {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return null;
    const r = sel.getRangeAt(0).cloneRange();
    if (!ed.contains(r.startContainer) || !ed.contains(r.endContainer)) return null;
    const m = paragraphMarkSelection();
    if (m) {
      // B bir listenin ilk maddesiyse seçim listenin önünde biter: panoya boş liste kabuğu gitmesin
      let e = m.B;
      while (e.parentElement && e.parentElement !== ed && !e.previousSibling) e = e.parentElement;
      r.setEndBefore(e);
    }
    const box = document.createElement('div');
    box.appendChild(r.cloneContents());
    box.querySelectorAll('.el').forEach((e) => e.classList.remove('el'));
    box.querySelectorAll('[class=""]').forEach((e) => e.removeAttribute('class'));
    let html = box.innerHTML;
    if (!m) {
      // Paragraf içinden kopya: paragrafın kendi yazı tipi/boyutu (içe aktarmada bloğa yazılır) metinle gitsin
      const u = unitOf(r.startContainer);
      const st = u && ['fontFamily', 'fontSize'].filter((k) => u.style[k]).map((k) => `${k === 'fontFamily' ? 'font-family' : 'font-size'}: ${u.style[k]}`);
      if (st && st.length) html = `<span style="${esc(st.join('; '))}">${html}</span>`;
    }
    return `<div ${CLIP_ATTR}="${m ? 'paragraf' : 'metin'}">${html}</div>`;
  }
  function onCopy(e) {
    if (app.objectsFocused && app.objectsFocused()) return; // resim panosu: objects.js
    const html = copyHTML();
    if (!html) return;
    e.clipboardData.setData('text/html', html);
    e.clipboardData.setData('text/plain', window.getSelection().toString());
    e.preventDefault();
    if (e.type === 'cut') app.deleteSelection();
  }
  document.addEventListener('copy', onCopy);
  document.addEventListener('cut', onCopy);

  // Kendi panomuzdaki HTML: yalnızca sözlükteki öğeler ve stilleri (ayrıca sayfa sonu, ol[start], a[href])
  const KEEP = new Set(['P', 'H1', 'H2', 'H3', 'UL', 'OL', 'LI', 'B', 'I', 'U', 'S', 'SUB', 'SUP', 'SPAN', 'BR', 'A']);
  function cleanInternal(src, dst) {
    for (const n of [...src.childNodes]) {
      if (n.nodeType === 3) dst.appendChild(document.createTextNode(n.nodeValue));
      if (n.nodeType !== 1) continue;
      let tag = n.tagName.toUpperCase();
      if (tag === 'DIV' && n.classList.contains('pb')) {
        const pb = dst.appendChild(document.createElement('div'));
        pb.className = 'pb';
        pb.contentEditable = 'false';
        continue;
      }
      if (/^H[4-6]$/.test(tag)) tag = 'H3';
      if (tag === 'FONT') tag = 'SPAN';
      if (!KEEP.has(tag)) {
        cleanInternal(n, dst);
        continue;
      }
      const el = dst.appendChild(document.createElement(tag.toLowerCase()));
      if (n.getAttribute('style')) el.setAttribute('style', n.getAttribute('style'));
      if (n.getAttribute('data-tabs')) el.setAttribute('data-tabs', app.formatTabs(app.parseTabs(n.getAttribute('data-tabs'))));
      if (/^H[1-3]$/.test(tag) && n.getAttribute('data-num') === '0') el.setAttribute('data-num', '0'); // numarasız başlık
      if (tag === 'SPAN' && n.classList.contains('tab')) el.className = 'tab';
      if (tag === 'OL' && +n.getAttribute('start') > 1) el.setAttribute('start', n.getAttribute('start'));
      if (tag === 'A' && app.safeHref && app.safeHref(n.getAttribute('href'))) el.setAttribute('href', app.safeHref(n.getAttribute('href')));
      cleanInternal(n, el);
    }
    return dst;
  }

  // Tam paragrafları imlecin olduğu paragrafın önüne/arkasına (ortadaysa paragrafı bölerek) ayrı paragraf olarak koy
  function insertParagraphs(box) {
    const sel = window.getSelection();
    const r = sel.getRangeAt(0);
    const C = unitOf(r.startContainer);
    const nodes = [...box.childNodes].filter((n) => n.nodeType === 1);
    const lists = nodes.every((n) => /^(UL|OL)$/.test(n.tagName));
    if (!C || !nodes.length) return false;
    const inList = C.tagName === 'LI';
    const li = C.closest('li');
    if (li && ed.contains(li) && !(inList && lists && nodes.length === 1 && nodes[0].tagName === C.parentElement.tagName))
      return insertParagraphsInList(li, r, nodes);
    const blocks = inList ? [...nodes[0].children] : nodes; // listeye liste maddeleri girer
    let ref;
    const after = document.createRange();
    after.setStart(r.startContainer, r.startOffset);
    after.setEnd(C, C.childNodes.length);
    const sub = C.querySelector(':scope > ul, :scope > ol');
    if (sub) after.setEndBefore(sub);
    const f = after.cloneContents();
    if (atStartOf(C, r.startContainer, r.startOffset)) ref = C;
    else if (!f.textContent.length && !f.querySelector('br')) ref = C.nextSibling;
    else {
      const C2 = C.cloneNode(false); // paragrafın ortası: böl
      C2.appendChild(after.extractContents());
      C.after(C2);
      if (!C.textContent && !C.querySelector('br')) C.appendChild(document.createElement('br'));
      ref = C2;
    }
    const parent = C.parentNode;
    let last = null;
    app.withoutBands(() => blocks.forEach((b) => (last = parent.insertBefore(b, ref))));
    const c = document.createRange();
    if (ref && ref.nodeType === 1) c.setStart(ref, 0); // Word: imleç yapıştırılanın ardında
    else {
      c.selectNodeContents(last);
      c.collapse(false);
    }
    c.collapse(true);
    sel.removeAllRanges();
    sel.addRange(c);
    return true;
  }

  // Liste maddesine liste olmayan tam paragraflar: paragraflar listenin dışında, kendi biçimleriyle girer; liste
  // imlecin yerinden bölünür (maddenin başı: önüne, metninin sonu: ardına, ortası: madde ikiye), kalan maddeler
  // yeni listede numarası sürerek devam eder. Chrome'un insertHTML'i paragrafı maddenin içine koyuyordu.
  function insertParagraphsInList(li, r, nodes) {
    let pos = { node: r.startContainer, off: r.startOffset };
    if (atStartOf(li, r.startContainer, r.startOffset)) pos = { node: li.parentNode, off: Array.prototype.indexOf.call(li.parentNode.childNodes, li) };
    else {
      const tail = document.createRange(); // maddenin kendi metninin sonu mu (alt listesi hariç)
      tail.setStart(r.startContainer, r.startOffset);
      tail.setEnd(li, li.childNodes.length);
      const sub = li.querySelector(':scope > ul, :scope > ol');
      if (sub) tail.setEndBefore(sub);
      const f = tail.cloneContents();
      if (!f.textContent.length && !f.querySelector('br') && !sub) pos = { node: li.parentNode, off: Array.prototype.indexOf.call(li.parentNode.childNodes, li) + 1 };
    }
    let last = null;
    let next = null;
    app.withoutBands(() => {
      const cut = app.splitTop(pos.node, pos.off);
      const ref = cut.top.nextSibling;
      nodes.forEach((n) => (last = ed.insertBefore(n, ref)));
      if (cut.rest) next = ed.insertBefore(cut.rest, ref);
      cut.dropEmptyTop();
    });
    // Word: imleç yapıştırılanın ardında (kalan ilk maddenin başı ya da son paragrafın sonu)
    const c = document.createRange();
    let first = next;
    while (first && !first.matches(UNIT_SEL) && first.firstElementChild) first = first.firstElementChild;
    if (first) c.setStart(first, 0);
    else {
      c.selectNodeContents(last);
      c.collapse(false);
    }
    c.collapse(true);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(c);
    return true;
  }

  // Satır içi içerik doğrudan imlecin yerine konur (Chrome'un insertHTML'i stil span'larını ayıklıyor)
  function insertInline(box) {
    const sel = window.getSelection();
    const r = sel.getRangeAt(0);
    const frag = document.createDocumentFragment();
    while (box.firstChild) frag.appendChild(box.firstChild);
    const last = frag.lastChild;
    if (!last) return;
    r.insertNode(frag);
    const c = document.createRange();
    c.setStartAfter(last);
    c.collapse(true);
    sel.removeAllRanges();
    sel.addRange(c);
  }

  function pasteInternal(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const root = doc.querySelector(`[${CLIP_ATTR}]`);
    const box = cleanInternal(root, document.createElement('div'));
    box.querySelectorAll('ul, ol').forEach((l) => !l.querySelector('li') && l.remove()); // eski panolardaki boş kabuk
    const sel = window.getSelection();
    if (sel.rangeCount && !sel.isCollapsed) app.deleteSelection(); // seçimin yerine
    if (!box.querySelector('p,h1,h2,h3,ul,ol,li,.pb')) insertInline(box);
    else if (!(root.getAttribute(CLIP_ATTR) === 'paragraf' && insertParagraphs(box)))
      app.withoutBands(() => document.execCommand('insertHTML', false, box.innerHTML));
    ensureContent();
    afterFormat();
    revealCaretSoon();
  }

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
    if (html && html.includes(CLIP_ATTR)) pasteInternal(html);
    else if (html) app.insertSanitized(app.sanitizeHTML(html));
    else {
      app.trimParagraphMark();
      insertPlain(dt.getData('text/plain'));
      ensureContent();
      app.scheduleLayout();
      app.commit('edit');
    }
  });
})();
