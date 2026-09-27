/* Word (.docx) dışa aktarma.
 * Metin paragraf/biçim olarak, resimler SAYFAYA GÖRE MUTLAK konumla (wp:anchor relativeFrom="page",
 * locked, allowOverlap) yazılır. Her resmin çapası, editörde o sayfanın ilk karakterine konur;
 * böylece Word resmi aynı sayfada, aynı noktada gösterir.
 */
(function () {
  'use strict';
  const SS = window.SS;
  const U = SS.units;
  const twip = U.pxToTwip;
  const emu = U.pxToEmu;
  const DEFAULT_FONT = 'Calibri';

  const esc = (s) => s.replace(/[<>&"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' })[c]);
  // eslint-disable-next-line no-control-regex
  const clean = (s) => s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '').replace(/ /g, ' ');
  const firstFamily = (f) => (f || '').split(',')[0].replace(/["']/g, '').trim() || DEFAULT_FONT;

  function rgbToHex(c) {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(c || '');
    if (!m || (m[4] !== undefined && parseFloat(m[4]) === 0)) return null;
    return [m[1], m[2], m[3]].map((v) => (+v).toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  // ---------- Sayfa başlarını ölç ----------
  // starts[k] = { node, offset, top } : akışta k. sayfanın üstünden sonra gelen ilk metin konumu
  function findPageStarts(app) {
    const { stride } = app.geom();
    const items = [];
    const walker = document.createTreeWalker(app.els.editor, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
      acceptNode: (n) =>
        (n.nodeType === 3 && n.length) || n.nodeName === 'BR' ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP,
    });
    for (let n = walker.nextNode(); n; n = walker.nextNode()) items.push(n);

    const rng = document.createRange();
    const flowY = (r) => app.clientToFlow(0, r.top).y;
    const firstVisible = (rects) => {
      for (const r of rects) if (r.width || r.height) return flowY(r);
      return NaN;
    };
    const charTop = (node, i) => {
      rng.setStart(node, i);
      rng.setEnd(node, i + 1);
      return firstVisible(rng.getClientRects());
    };
    const lastTop = (node) => {
      if (node.nodeType !== 3) return firstVisible(node.getClientRects());
      rng.selectNodeContents(node);
      const rects = [...rng.getClientRects()].reverse();
      return firstVisible(rects);
    };

    const starts = [];
    let idx = 0;
    for (let k = 0; k < app.state.pageCount; k++) {
      const pageTop = k * stride - 1;
      while (idx < items.length && !(lastTop(items[idx]) >= pageTop)) idx++;
      if (idx >= items.length) break;
      const node = items[idx];
      if (node.nodeType !== 3) {
        starts[k] = { node, offset: 0, top: lastTop(node) };
        continue;
      }
      const tAt = (i) => {
        for (let j = i; j < node.length; j++) {
          const t = charTop(node, j);
          if (!isNaN(t)) return t;
        }
        return Infinity;
      };
      let lo = 0;
      let hi = node.length - 1;
      let ans = node.length - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (tAt(mid) >= pageTop) {
          ans = mid;
          hi = mid - 1;
        } else lo = mid + 1;
      }
      starts[k] = { node, offset: ans, top: tAt(ans) };
    }
    return starts;
  }

  // ---------- Numaralandırma ----------
  function makeNumbering() {
    const nums = []; // { id, ordered, start }
    return {
      add(ordered, start = 1) {
        const id = nums.length + 1;
        nums.push({ id, ordered, start });
        return id;
      },
      xml() {
        // Word'ün varsayılan liste girintisi (editör css'i de aynı): metin her düzeyde 1,27 cm içeride,
        // işaret 0,635 cm solunda
        const lvls = (ordered) => {
          let s = '';
          for (let i = 0; i < 9; i++) {
            const left = 720 * (i + 1);
            if (ordered) {
              const fmt = ['decimal', 'lowerLetter', 'lowerRoman'][i % 3];
              s += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
            } else {
              const [chr, font] = [['&#xF0B7;', 'Symbol'], ['o', 'Courier New'], ['&#xF0A7;', 'Wingdings']][i % 3];
              s += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${chr}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr></w:lvl>`;
            }
          }
          return s;
        };
        return (
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvls(false)}</w:abstractNum>` +
          `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvls(true)}</w:abstractNum>` +
          nums
            .map((n) =>
              n.ordered
                ? `<w:num w:numId="${n.id}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="${n.start}"/></w:lvlOverride></w:num>`
                : `<w:num w:numId="${n.id}"><w:abstractNumId w:val="0"/></w:num>`
            )
            .join('') +
          '</w:numbering>'
        );
      },
    };
  }

  // ---------- Ana dışa aktarma ----------
  SS.exportDocx = async function (app) {
    if (app.paginateNow) app.paginateNow(); // sayfa sonları kesinleşsin: çapalar canlı düzenden ölçülür
    const state = app.state;
    const g = app.geom();
    const ed = app.els.editor;
    const numbering = makeNumbering();

    // Görseller (her varlık bir kez)
    const media = new Map();
    const mediaFiles = [];
    const mediaFor = (assetId) => {
      if (media.has(assetId)) return media.get(assetId);
      const a = state.assets[assetId];
      const { mime, bytes } = SS.dataURLtoBytes(a.src);
      const ext = mime === 'image/jpeg' ? 'jpeg' : 'png';
      const n = mediaFiles.length + 1;
      const m = { rId: 'rIdImg' + n, name: `image${n}.${ext}`, ext };
      mediaFiles.push({ name: 'word/media/' + m.name, data: bytes, rId: m.rId });
      media.set(assetId, m);
      return m;
    };

    const byZ = [...state.images].sort((a, b) => a.z - b.z);
    const relHeight = new Map(byZ.map((img, i) => [img.id, 251658240 + i * 1024]));
    const capNums = app.captionNumbers();
    const imgsOnPage = [];
    for (const img of byZ) (imgsOnPage[img.page] = imgsOnPage[img.page] || []).push(img);
    let docPrId = 0;

    function drawing(img, noWrap) {
      const m = mediaFor(img.asset);
      const id = ++docPrId;
      const cx = emu(img.w);
      const cy = emu(img.h);
      const rot = ((Math.round((img.rot || 0) * 60000) % 21600000) + 21600000) % 21600000;
      const bb = app.aabb(img);
      const ex = emu(Math.max(0, (bb.w - img.w) / 2));
      const ey = emu(Math.max(0, (bb.h - img.h) / 2));
      const dist = emu(app.WRAP_DIST);
      const mode = noWrap ? (img.wrap === 'behind' ? 'behind' : 'front') : img.wrap;
      const wrap =
        mode === 'square' ? '<wp:wrapSquare wrapText="largest"/>' : mode === 'topbottom' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapNone/>';
      return (
        '<w:r><w:drawing>' +
        `<wp:anchor distT="${dist}" distB="${dist}" distL="${dist}" distR="${dist}" simplePos="0" relativeHeight="${relHeight.get(img.id)}" behindDoc="${mode === 'behind' ? 1 : 0}" locked="1" layoutInCell="1" allowOverlap="1">` +
        '<wp:simplePos x="0" y="0"/>' +
        `<wp:positionH relativeFrom="page"><wp:posOffset>${emu(img.x)}</wp:posOffset></wp:positionH>` +
        `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(img.y)}</wp:posOffset></wp:positionV>` +
        `<wp:extent cx="${cx}" cy="${cy}"/>` +
        `<wp:effectExtent l="${ex}" t="${ey}" r="${ex}" b="${ey}"/>` +
        wrap +
        `<wp:docPr id="${id}" name="Resim ${id}"/>` +
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
        `<pic:nvPicPr><pic:cNvPr id="${id}" name="${m.name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${m.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm${rot ? ` rot="${rot}"` : ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
      );
    }

    // Şekil yazısı: resmin altına/üstüne sabitlenmiş kenarlıksız metin kutusu.
    // "Caption" (Resim Yazısı) stili ve SEQ alanı sayesinde Word'de "Şekiller Tablosu" da oluşturulabilir.
    function captionBox(img, noWrap) {
      const c = app.capRect(img);
      if (!c) return '';
      const id = ++docPrId;
      const label = img.caption.label || 'Şekil';
      const cx = emu(c.w);
      const cy = emu(c.h + 6); // Word'ün yazı ölçüsü biraz farklı olabilir: küçük pay
      const dist = emu(app.WRAP_DIST);
      const mode = noWrap ? (img.wrap === 'behind' ? 'behind' : 'front') : img.wrap;
      const wrap =
        mode === 'square' ? '<wp:wrapSquare wrapText="largest"/>' : mode === 'topbottom' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapNone/>';
      const bold = '<w:rPr><w:b/><w:bCs/></w:rPr>';
      const text = clean(img.caption.text || '');
      const para =
        '<w:p><w:pPr><w:pStyle w:val="Caption"/><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>' +
        `<w:r>${bold}<w:t xml:space="preserve">${esc(label)} </w:t></w:r>` +
        `<w:fldSimple w:instr=" SEQ ${esc(label.replace(/\s+/g, '_'))} \\* ARABIC "><w:r>${bold}<w:t>${capNums.get(img.id) || 1}</w:t></w:r></w:fldSimple>` +
        `<w:r>${bold}<w:t xml:space="preserve">.</w:t></w:r>` +
        (text ? `<w:r><w:t xml:space="preserve"> ${esc(text)}</w:t></w:r>` : '') +
        '</w:p>';
      return (
        '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
        `<wp:anchor distT="${dist}" distB="${dist}" distL="${dist}" distR="${dist}" simplePos="0" relativeHeight="${relHeight.get(img.id) + 1}" behindDoc="${mode === 'behind' ? 1 : 0}" locked="1" layoutInCell="1" allowOverlap="1">` +
        '<wp:simplePos x="0" y="0"/>' +
        `<wp:positionH relativeFrom="page"><wp:posOffset>${emu(c.x)}</wp:posOffset></wp:positionH>` +
        `<wp:positionV relativeFrom="page"><wp:posOffset>${emu(c.y)}</wp:posOffset></wp:positionV>` +
        `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
        wrap +
        `<wp:docPr id="${id}" name="Şekil yazısı ${id}"/><wp:cNvGraphicFramePr/>` +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
        '<wps:cNvSpPr txBox="1"/>' +
        `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
        `<wps:txbx><w:txbxContent>${para}</w:txbxContent></wps:txbx>` +
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0" upright="1"><a:spAutoFit/></wps:bodyPr>' +
        '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>'
      );
    }

    // ---- Çapa planı ----
    const starts = findPageStarts(app);
    const events = new Map(); // node -> [{ offset, type: 'holder'|'anchor', page }]
    const addEvent = (node, offset, type, page) => {
      if (!events.has(node)) events.set(node, []);
      events.get(node).push({ offset, type, page });
    };
    const trailing = [];
    const anchorPages = []; // metinli ve resimli sayfalar: çapa yeri paragraflar toplanınca seçilir (placeAnchors)
    for (let k = 0; k < state.pageCount; k++) {
      const s = starts[k];
      const onPage = s && s.top < k * g.stride + g.ch + 1;
      if (onPage) {
        if (imgsOnPage[k]) anchorPages.push(k);
      } else if (s) addEvent(s.node, s.offset, 'holder', k); // metinsiz sayfa (ör. tam sayfa resim)
      else trailing.push(k); // metnin bittiği yerden sonraki sayfalar
    }
    // Word ve LibreOffice nesneyi çapa PARAGRAFININ sayfasına koyar. Çapa bir paragrafın ortasına düşerse
    // (sayfa paragraf ortasında başlıyorsa) resim paragrafın başladığı önceki sayfaya kayar. Bu yüzden çapa,
    // o sayfada BAŞLAYAN paragraflardan sayfa ortasına en yakın olanın başına konur: satır kırılımı farkıyla
    // metin birkaç satır kaysa da paragraf aynı sayfada kalır. Sayfada başlayan paragraf yoksa (sayfayı aşan
    // uzun paragraf) sayfanın ilk karakteri kullanılır.
    function placeAnchors(units) {
      const rng = document.createRange();
      const visibleTop = (rects) => {
        for (const r of rects) if (r.width || r.height) return app.clientToFlow(0, r.top).y;
        return NaN;
      };
      const topOf = (n) => {
        if (n.nodeType !== 3) return visibleTop(n.getClientRects());
        for (let i = 0; i < n.length; i++) {
          rng.setStart(n, i);
          rng.setEnd(n, i + 1);
          const t = visibleTop(rng.getClientRects());
          if (!isNaN(t)) return t;
        }
        return NaN;
      };
      const heads = units.map((u) => {
        const first = collectInline(u.el)[0];
        return first ? { first, top: topOf(first) } : null;
      }).filter((h) => h && !isNaN(h.top));
      for (const k of anchorPages) {
        const top = k * g.stride;
        const mid = top + g.ch / 2;
        let best = null;
        for (const h of heads)
          if (h.top >= top - 1 && h.top < top + g.ch + 1 && (!best || Math.abs(h.top - mid) < Math.abs(best.top - mid))) best = h;
        if (best) addEvent(best.first, 0, 'anchor', k);
        else addEvent(starts[k].node, starts[k].offset, 'anchor', k);
      }
      for (const list of events.values())
        list.sort((a, b) => a.offset - b.offset || (a.type === b.type ? 0 : a.type === 'holder' ? -1 : 1));
    }
    const consumed = new Set();
    const anchorsFor = (k, noWrap) => {
      consumed.add(k);
      return (imgsOnPage[k] || []).map((img) => drawing(img, noWrap) + captionBox(img, noWrap)).join('');
    };
    const holderPara = (k) =>
      `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr>${anchorsFor(k, true)}</w:p>`;

    // ---- Biçim ----
    const rPrCache = new Map();
    function rPr(el, blockEl) {
      const key = el;
      if (rPrCache.has(key)) return rPrCache.get(key);
      const cs = getComputedStyle(el);
      const p = [];
      const font = firstFamily(cs.fontFamily);
      if (font !== DEFAULT_FONT) p.push(`<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:eastAsia="${esc(font)}" w:cs="${esc(font)}"/>`);
      if (parseInt(cs.fontWeight, 10) >= 600) p.push('<w:b/><w:bCs/>');
      else if (/^H[1-6]$/.test(blockEl.tagName)) p.push('<w:b w:val="0"/><w:bCs w:val="0"/>'); // başlık stili kalın: normal metin açıkça
      if (cs.fontStyle === 'italic' || cs.fontStyle === 'oblique') p.push('<w:i/><w:iCs/>');
      let underline = false;
      let strike = false;
      let bg = null;
      let va = null;
      for (let n = el; n; n = n.parentElement) {
        const s = getComputedStyle(n);
        const d = s.textDecorationLine || s.textDecoration || '';
        if (d.includes('underline')) underline = true;
        if (d.includes('line-through')) strike = true;
        if (!bg && n !== blockEl) bg = rgbToHex(s.backgroundColor);
        if (!va && (n.tagName === 'SUP' || n.tagName === 'SUB')) va = n.tagName === 'SUP' ? 'superscript' : 'subscript';
        if (n === blockEl || n === ed) break;
      }
      if (strike) p.push('<w:strike/>');
      const color = rgbToHex(cs.color);
      if (color && color !== '000000') p.push(`<w:color w:val="${color}"/>`);
      let px = parseFloat(cs.fontSize);
      if (va) px = parseFloat(getComputedStyle(el.closest('sup,sub').parentElement).fontSize);
      const hp = Math.max(2, Math.round(px * 1.5));
      p.push(`<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`);
      if (underline) p.push('<w:u w:val="single"/>');
      if (bg) p.push(`<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`);
      if (va) p.push(`<w:vertAlign w:val="${va}"/>`);
      const xml = `<w:rPr>${p.join('')}</w:rPr>`;
      rPrCache.set(key, xml);
      return xml;
    }

    function textRuns(node, from, to, blockEl) {
      const text = clean(node.nodeValue.slice(from, to));
      if (!text) return '';
      const props = rPr(node.parentElement, blockEl);
      const body = text
        .split(/(\t|\n)/)
        .map((t) => (t === '\t' ? '<w:tab/>' : t === '\n' ? '<w:br/>' : t ? `<w:t xml:space="preserve">${esc(t)}</w:t>` : ''))
        .join('');
      return `<w:r>${props}${body}</w:r>`;
    }

    let prevAfter = 0;
    function pPr(blockEl, o) {
      const cs = getComputedStyle(blockEl);
      const p = [];
      if (o.style) p.push(`<w:pStyle w:val="${o.style}"/>`);
      if (o.pageBreakBefore) p.push('<w:pageBreakBefore/>');
      if (o.numId && !o.cont) p.push(`<w:numPr><w:ilvl w:val="${o.ilvl}"/><w:numId w:val="${o.numId}"/></w:numPr>`);
      const fontPx = parseFloat(cs.fontSize);
      // Birimsiz satır yüksekliği Word'ün "satır katı"dır (auto); birimli değer "En az"
      const ls = app.lineSpacing(blockEl);
      const line = ls.px ? `w:line="${twip(ls.px)}" w:lineRule="atLeast"` : `w:line="${Math.round(240 * ls.multiple)}" w:lineRule="auto"`;
      const before = o.cont ? 0 : Math.max(0, o.mt - prevAfter);
      p.push(`<w:spacing w:before="${twip(before)}" w:after="${twip(o.mb)}" ${line}/>`);
      if (o.numId && o.cont) p.push(`<w:ind w:left="${720 * (o.ilvl + 1)}"/>`);
      else if (!o.numId) {
        const ml = parseFloat(cs.marginLeft) || 0;
        const mr = parseFloat(cs.marginRight) || 0;
        const ti = o.cont ? 0 : parseFloat(cs.textIndent) || 0;
        const ind = [];
        if (ml > 0) ind.push(`w:left="${twip(ml)}"`);
        if (mr > 0) ind.push(`w:right="${twip(mr)}"`);
        if (ti > 0) ind.push(`w:firstLine="${twip(ti)}"`);
        else if (ti < 0) ind.push(`w:hanging="${twip(-ti)}"`);
        if (ind.length) p.push(`<w:ind ${ind.join(' ')}/>`);
      }
      const jc = { center: 'center', right: 'right', end: 'right', justify: 'both' }[cs.textAlign];
      if (jc) p.push(`<w:jc w:val="${jc}"/>`);
      const hp = Math.round(fontPx * 1.5);
      p.push(`<w:rPr><w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/></w:rPr>`);
      prevAfter = o.mb;
      return `<w:pPr>${p.join('')}</w:pPr>`;
    }

    function collectInline(el) {
      const out = [];
      const walk = (n) => {
        if (n.nodeType === 3) {
          if (n.length) out.push(n);
        } else if (n.nodeType === 1) {
          if (n.tagName === 'BR') out.push(n);
          else if (n.tagName !== 'UL' && n.tagName !== 'OL') n.childNodes.forEach(walk);
        }
      };
      if (el.nodeType === 3) out.push(el);
      else el.childNodes.forEach(walk);
      return out;
    }

    const body = [];
    const units = []; // yazılacak paragraflar, belge sırasıyla: { el, o }
    let pendingBreak = false; // editördeki sayfa sonu -> sonraki paragrafa "öncesinde sayfa sonu"
    function unit(el, o) {
      if (pendingBreak) {
        o = { ...o, pageBreakBefore: true };
        pendingBreak = false;
      }
      units.push({ el, o });
    }
    // Bir blok -> bir ya da daha fazla w:p (metinsiz sayfa araya girerse paragraf bölünür)
    function emitBlock(el, o) {
      const styleEl = el.nodeType === 3 ? el.parentElement : el;
      const items = collectInline(el);
      const paras = [];
      let cur = { o, runs: [] };
      const onEvent = (ev) => {
        if (ev.type === 'anchor') return void cur.runs.push(anchorsFor(ev.page, false));
        const atStart = !cur.runs.length && !paras.length;
        if (!atStart) paras.push(cur);
        paras.push({ holder: ev.page });
        cur = { o: { ...o, pageBreakBefore: true, cont: !atStart }, runs: [] };
      };
      items.forEach((node, i) => {
        const evs = events.get(node) || [];
        if (node.nodeType === 3) {
          let pos = 0;
          for (const ev of evs) {
            if (ev.offset > pos) cur.runs.push(textRuns(node, pos, ev.offset, styleEl));
            pos = Math.max(pos, ev.offset);
            onEvent(ev);
          }
          cur.runs.push(textRuns(node, pos, node.length, styleEl));
        } else {
          evs.forEach(onEvent);
          if (i < items.length - 1) cur.runs.push('<w:r><w:br/></w:r>'); // bloğun sonundaki <br> satır üretmez
        }
      });
      paras.push(cur);
      for (const p of paras) body.push(p.holder !== undefined ? holderPara(p.holder) : `<w:p>${pPr(styleEl, p.o)}${p.runs.join('')}</w:p>`);
    }

    const blockMargins = (el) => {
      if (el.nodeType === 3) return { mt: 0, mb: 0 };
      const cs = getComputedStyle(el);
      return { mt: parseFloat(cs.marginTop) || 0, mb: parseFloat(cs.marginBottom) || 0 };
    };

    function walk(nodes, listCtx) {
      for (const n of nodes) {
        if (n.nodeType === 3) {
          if (n.nodeValue.trim()) unit(n, { mt: 0, mb: 0 });
          continue;
        }
        if (n.nodeType !== 1) continue;
        const tag = n.tagName;
        if (tag === 'UL' || tag === 'OL') {
          const ordered = tag === 'OL';
          const depth = listCtx ? listCtx.depth + 1 : 0;
          const start = ordered ? Math.max(1, parseInt(n.getAttribute('start'), 10) || 1) : 1; // bölünen listede numara devam eder
          const numId = listCtx && listCtx.ordered === ordered ? listCtx.numId : numbering.add(ordered, start);
          const ctx = { ordered, depth, numId };
          // Chrome'un girinti komutu alt listeyi maddenin içine değil listenin kendisine koyar: <ul><li/><ul>…</ul></ul>
          const items = [...n.children].filter((c) => /^(LI|UL|OL)$/.test(c.tagName));
          const lcs = getComputedStyle(n);
          items.forEach((li, i) => {
            if (li.tagName !== 'LI') return void walk([li], ctx);
            const mt = i === 0 && !listCtx ? parseFloat(lcs.marginTop) || 0 : 0;
            const nested = [...li.children].filter((c) => c.tagName === 'UL' || c.tagName === 'OL');
            const mb = i === items.length - 1 && !listCtx && !nested.length ? parseFloat(lcs.marginBottom) || 0 : 0;
            unit(li, { numId, ilvl: Math.min(depth, 8), mt, mb });
            walk(nested, ctx);
          });
        } else if (/^(P|H[1-6]|DIV|BLOCKQUOTE|PRE)$/.test(tag)) {
          if (n.classList.contains('pb')) pendingBreak = true;
          else if (n.querySelector('p,h1,h2,h3,h4,h5,h6,div,ul,ol,li,blockquote,pre')) walk([...n.childNodes], listCtx);
          else {
            const lvl = /^H([1-6])$/.exec(tag);
            unit(n, { style: lvl ? 'Heading' + Math.min(3, +lvl[1]) : null, ...blockMargins(n) });
          }
        } else unit(n, blockMargins(n));
      }
    }
    walk([...ed.childNodes], null);
    placeAnchors(units);
    units.forEach((u) => emitBlock(u.el, u.o));
    if (pendingBreak) body.push('<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>');
    for (const k of trailing) body.push(holderPara(k));
    // Güvenlik: herhangi bir nedenle çapası yazılamamış sayfa kalırsa resimler kaybolmasın
    imgsOnPage.forEach((list, k) => {
      if (list && !consumed.has(k)) body.push(holderPara(k));
    });

    // ---- Üst/alt bilgi ----
    // Sol/orta/sağ yuvalar tek paragrafta sekmeyle ayrılır (orta ve sağ sekme durağı); {sayfa} ve {toplam}
    // PAGE ve NUMPAGES alanı olur. "İlk sayfada gösterme": titlePg + boş ilk sayfa üst/alt bilgisi.
    const hf = state.hf;
    const hfUsed = (slots) => slots.some((s) => s.trim());
    const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const field = (instr) =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const hfRuns = (s) =>
      s.split(/(\{sayfa\}|\{toplam\})/).map((t) =>
        t === '{sayfa}' ? field('PAGE') : t === '{toplam}' ? field('NUMPAGES') : t ? `<w:r><w:t xml:space="preserve">${esc(clean(t))}</w:t></w:r>` : ''
      ).join('');
    const hfXml = (tag, style, slots) => {
      const last = slots[2].trim() ? 2 : slots[1].trim() ? 1 : 0;
      const runs = slots.slice(0, last + 1).map((s, i) => (i ? '<w:r><w:tab/></w:r>' : '') + hfRuns(s)).join('');
      return (
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} ${W_NS}><w:p><w:pPr><w:pStyle w:val="${style}"/>` +
        `<w:tabs><w:tab w:val="center" w:pos="${twip(g.cw / 2)}"/><w:tab w:val="right" w:pos="${twip(g.cw)}"/></w:tabs></w:pPr>${runs}</w:p></w:${tag}>`
      );
    };
    const hfParts = [];
    if (hfUsed(hf.header)) hfParts.push({ kind: 'header', ref: 'default', file: 'header1.xml', data: hfXml('hdr', 'Header', hf.header) });
    if (hfUsed(hf.footer)) hfParts.push({ kind: 'footer', ref: 'default', file: 'footer1.xml', data: hfXml('ftr', 'Footer', hf.footer) });
    const titlePg = hf.firstPage && hfParts.length > 0;
    if (titlePg) {
      hfParts.push({ kind: 'header', ref: 'first', file: 'header2.xml', data: hfXml('hdr', 'Header', ['', '', '']) });
      hfParts.push({ kind: 'footer', ref: 'first', file: 'footer2.xml', data: hfXml('ftr', 'Footer', ['', '', '']) });
    }
    hfParts.forEach((p, i) => (p.rId = 'rIdHf' + (i + 1)));

    // ---- Bölüm (sayfa) ayarları ----
    const m = g.m;
    const sectPr =
      '<w:sectPr>' +
      hfParts.map((p) => `<w:${p.kind}Reference w:type="${p.ref}" r:id="${p.rId}"/>`).join('') +
      `<w:pgSz w:w="${twip(g.PW)}" w:h="${twip(g.PH)}"${state.page.orient === 'landscape' ? ' w:orient="landscape"' : ''}/>` +
      `<w:pgMar w:top="${twip(m.t)}" w:right="${twip(m.r)}" w:bottom="${twip(m.b)}" w:left="${twip(m.l)}" w:header="${twip(app.HEADER_DIST)}" w:footer="${twip(app.FOOTER_DIST)}" w:gutter="0"/>` +
      '<w:cols w:space="708"/>' +
      (titlePg ? '<w:titlePg/>' : '') +
      '</w:sectPr>';

    const documentXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      `<w:body>${body.join('')}${sectPr}</w:body></w:document>`;

    const heading = (id, name, lvl, hp, before, after) =>
      `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
      `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="${after}"/><w:outlineLvl w:val="${lvl}"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/></w:rPr></w:style>`;
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>' +
      '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="tr-TR" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
      // Dul/öksüz satır denetimi açık: editörün sayfalaması da aynı kuralı uygular (core.js: boundaryCut)
      '<w:pPrDefault><w:pPr><w:widowControl/><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
      heading('Heading1', 'heading 1', 0, 36, 240, 120) +
      heading('Heading2', 'heading 2', 1, 28, 200, 100) +
      heading('Heading3', 'heading 3', 2, 24, 160, 80) +
      '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>' +
      '</w:styles>';

    const settingsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
      '</w:settings>';

    const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const docRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdNumbering" Type="${REL}/numbering" Target="numbering.xml"/>` +
      `<Relationship Id="rIdSettings" Type="${REL}/settings" Target="settings.xml"/>` +
      hfParts.map((p) => `<Relationship Id="${p.rId}" Type="${REL}/${p.kind}" Target="${p.file}"/>`).join('') +
      mediaFiles.map((f) => `<Relationship Id="${f.rId}" Type="${REL}/image" Target="${f.name.slice(5)}"/>`).join('') +
      '</Relationships>';

    const OD = 'application/vnd.openxmlformats-officedocument';
    const contentTypes =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
      `<Override PartName="/word/document.xml" ContentType="${OD}.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${OD}.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="${OD}.wordprocessingml.numbering+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${OD}.wordprocessingml.settings+xml"/>` +
      hfParts.map((p) => `<Override PartName="/word/${p.file}" ContentType="${OD}.wordprocessingml.${p.kind}+xml"/>`).join('') +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      `<Override PartName="/docProps/app.xml" ContentType="${OD}.extended-properties+xml"/>` +
      '</Types>';

    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const coreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${esc(state.fileName || '')}</dc:title><dc:creator>SerbestSayfa</dc:creator>` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
      '</cp:coreProperties>';
    const appXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>SerbestSayfa</Application></Properties>';
    const rootRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>` +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>` +
      '</Relationships>';

    const files = [
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'docProps/core.xml', data: coreXml },
      { name: 'docProps/app.xml', data: appXml },
      { name: 'word/document.xml', data: documentXml },
      { name: 'word/styles.xml', data: stylesXml },
      { name: 'word/numbering.xml', data: numbering.xml() },
      { name: 'word/settings.xml', data: settingsXml },
      { name: 'word/_rels/document.xml.rels', data: docRels },
      ...hfParts.map((p) => ({ name: 'word/' + p.file, data: p.data })),
      ...mediaFiles.map((f) => ({ name: f.name, data: f.data })),
    ];
    return SS.makeZip(files, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  };
})();
