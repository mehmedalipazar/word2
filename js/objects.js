/* Resim nesneleri: seçim, sürükle-bırak, akıllı kılavuzlar, boyutlandırma, döndürme, hizalama, pano */
(function () {
  'use strict';
  const SS = window.SS;
  const U = SS.units;
  const app = SS.app;
  const els = app.els;
  const state = app.state;

  const SNAP_PX = 7; // ekran pikseli cinsinden yapışma mesafesi
  const MIN = 12; // en küçük kenar (px)
  const GRID = U.cmToPx(0.5);
  const HANDLES = [['nw', 0, 0], ['n', 0.5, 0], ['ne', 1, 0], ['e', 1, 0.5], ['se', 1, 1], ['s', 0.5, 1], ['sw', 0, 1], ['w', 0, 0.5]];
  let guides = [];

  const selImages = () => state.selection.map(app.getImage).filter(Boolean);
  app.selImages = selImages;
  const maxZ = () => state.images.reduce((m, i) => Math.max(m, i.z || 0), 0);
  const minZ = () => state.images.reduce((m, i) => Math.min(m, i.z || 0), 0);

  app.toDoc = function (e) {
    const r = els.doc.getBoundingClientRect();
    return { x: (e.clientX - r.left) / state.zoom, y: (e.clientY - r.top) / state.zoom };
  };
  function pageAtY(docY) {
    const g = app.geom();
    return SS.clamp(Math.floor((docY + g.gap / 2) / g.stride), 0, state.pageCount - 1);
  }
  app.pageAtY = pageAtY;

  // Belge koordinatındaki konuma taşı; resmin merkezi hangi sayfadaysa o sayfaya bağlanır
  function setDocPos(img, x, docY) {
    const g = app.geom();
    const page = pageAtY(docY + img.h / 2);
    img.page = page;
    img.x = x;
    img.y = docY - page * g.stride;
  }

  function clampToPage(img) {
    const g = app.geom();
    const bb = app.objBox(img);
    const keep = 16; // en az bu kadarı sayfada kalsın
    img.x += SS.clamp(bb.x, keep - bb.w, g.PW - keep) - bb.x;
    img.y += SS.clamp(bb.y, keep - bb.h, g.PH - keep) - bb.y;
  }
  app.clampToPage = clampToPage;

  // ---------- Seçim ----------
  function dropTextFocus() {
    window.getSelection().removeAllRanges();
    if (document.activeElement === els.editor) els.editor.blur();
  }

  app.select = function (ids, toggle = false) {
    ids = [].concat(ids);
    if (toggle) {
      const set = new Set(state.selection);
      ids.forEach((id) => (set.has(id) ? set.delete(id) : set.add(id)));
      state.selection = [...set];
    } else state.selection = ids;
    if (state.selection.length) dropTextFocus();
    app.renderOverlay();
    app.updateCtxBar();
  };

  app.clearSelection = function () {
    if (!state.selection.length) return;
    state.selection = [];
    app.renderOverlay();
    app.updateCtxBar();
  };

  app.renderOverlay = function () {
    const g = app.geom();
    const frag = document.createDocumentFragment();
    const imgs = selImages();
    const single = imgs.length === 1;
    for (const img of imgs) {
      const box = document.createElement('div');
      box.className = 'sel-box' + (img.locked ? ' locked' : single ? '' : ' multi');
      box.style.cssText = `left:${img.x}px;top:${img.page * g.stride + img.y}px;width:${img.w}px;height:${img.h}px;transform:rotate(${img.rot || 0}deg)`;
      if (single && !img.locked) {
        for (const [h, x, y] of HANDLES) {
          const d = document.createElement('div');
          d.className = 'handle';
          d.dataset.h = h;
          d.style.left = x * 100 + '%';
          d.style.top = y * 100 + '%';
          box.appendChild(d);
        }
        box.insertAdjacentHTML('beforeend', '<div class="rot-stem"></div><div class="handle rot" data-h="rot" style="left:50%;top:calc(-22px / var(--z))" title="Döndür (Shift: 15°)"></div>');
      }
      if (img.locked) box.insertAdjacentHTML('beforeend', `<span class="lock-badge">${SS.icon('lock')}</span>`);
      frag.appendChild(box);
      const c = app.capRect(img);
      if (c) {
        const cb = document.createElement('div');
        cb.className = 'cap-box';
        cb.style.cssText = `left:${c.x}px;top:${img.page * g.stride + c.y}px;width:${c.w}px;height:${c.h}px`;
        frag.appendChild(cb);
      }
    }
    for (const gd of guides) {
      const d = document.createElement('div');
      d.className = 'guide ' + gd.type;
      d.style.cssText =
        gd.type === 'v'
          ? `left:${gd.pos}px;top:${gd.from}px;height:${gd.to - gd.from}px`
          : `top:${gd.pos}px;left:${gd.from}px;width:${gd.to - gd.from}px`;
      frag.appendChild(d);
    }
    els.overlay.replaceChildren(frag);
  };

  // ---------- İsabet testi ----------
  function caretRangeAt(x, y) {
    if (document.caretRangeFromPoint) return document.caretRangeFromPoint(x, y);
    const p = document.caretPositionFromPoint && document.caretPositionFromPoint(x, y);
    if (!p) return null;
    const r = document.createRange();
    r.setStart(p.offsetNode, p.offset);
    return r;
  }

  function isOverText(cx, cy) {
    const r = caretRangeAt(cx, cy);
    const node = r && r.startContainer;
    if (!node || node.nodeType !== 3 || !els.editor.contains(node)) return false;
    const off = r.startOffset;
    const rng = document.createRange();
    for (const [a, b] of [[off, off + 1], [off - 1, off]]) {
      if (a < 0 || b > node.length) continue;
      rng.setStart(node, a);
      rng.setEnd(node, b);
      const rc = rng.getBoundingClientRect();
      if (cx >= rc.left - 1 && cx <= rc.right + 1 && cy >= rc.top && cy <= rc.bottom) return true;
    }
    return false;
  }

  function hitImage(e) {
    for (const el of document.elementsFromPoint(e.clientX, e.clientY)) {
      const obj = el.closest && el.closest('.img-obj, .cap');
      if (!obj) continue;
      if (obj.parentElement.parentElement === els.front) return obj.dataset.id;
      // Metnin arkasındaki resim: üzerinde yazı yoksa, zaten seçiliyse ya da Alt basılıysa tutulur
      const id = obj.dataset.id;
      if (e.altKey || state.selection.includes(id) || !isOverText(e.clientX, e.clientY)) return id;
      return null;
    }
    return null;
  }

  // Sayfanın metin olmayan bir yerine tıklanınca imleci en yakın metin konumuna koy
  function placeCaretNear(e) {
    const g = app.geom();
    const f = els.flow.getBoundingClientRect();
    const p = app.clientToFlow(e.clientX, e.clientY);
    let range = null;
    if (p.y < app.textBottom()) {
      const cx = SS.clamp(e.clientX, f.left + 2, f.left + g.cw * state.zoom - 2);
      range = caretRangeAt(cx, e.clientY);
    }
    if (!range || !els.editor.contains(range.startContainer)) {
      const last = els.editor.lastElementChild || els.editor;
      range = document.createRange();
      if (last.lastChild && last.lastChild.nodeName === 'BR') range.setStartBefore(last.lastChild);
      else {
        range.selectNodeContents(last);
        range.collapse(false);
      }
    }
    els.editor.focus({ preventScroll: true });
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  // ---------- Yapışma (akıllı kılavuzlar) ----------
  function snapTargets(page, excludeIds) {
    const g = app.geom();
    const top = page * g.stride;
    const xs = [0, g.m.l, g.PW / 2, g.PW - g.m.r, g.PW].map((v) => ({ v }));
    const ys = [0, g.m.t, g.PH / 2, g.PH - g.m.b, g.PH].map((v) => ({ v: v + top }));
    for (const o of state.images) {
      if (o.page !== page || excludeIds.includes(o.id)) continue;
      const bb = app.objBox(o);
      xs.push({ v: bb.x }, { v: bb.x + bb.w / 2 }, { v: bb.x + bb.w });
      ys.push({ v: top + bb.y }, { v: top + bb.y + bb.h / 2 }, { v: top + bb.y + bb.h });
    }
    return { xs, ys, page };
  }

  function snapAxis(edges, targets, thr) {
    let best = null;
    for (const e of edges)
      for (const t of targets) {
        const d = t.v - e;
        if (Math.abs(d) <= thr && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, t };
      }
    return best;
  }

  function addGuides(b, T) {
    const g = app.geom();
    const top = T.page * g.stride;
    const xe = [b.x0, (b.x0 + b.x1) / 2, b.x1];
    const ye = [b.y0, (b.y0 + b.y1) / 2, b.y1];
    const seen = new Set();
    for (const t of T.xs)
      if (xe.some((e) => Math.abs(e - t.v) < 0.5) && !seen.has('v' + t.v.toFixed(1))) {
        seen.add('v' + t.v.toFixed(1));
        guides.push({ type: 'v', pos: t.v, from: top, to: top + g.PH });
      }
    for (const t of T.ys)
      if (ye.some((e) => Math.abs(e - t.v) < 0.5) && !seen.has('h' + t.v.toFixed(1))) {
        seen.add('h' + t.v.toFixed(1));
        guides.push({ type: 'h', pos: t.v, from: 0, to: g.PW });
      }
  }

  function snapBox(box, page, excludeIds, noSnap) {
    const T = snapTargets(page, excludeIds);
    let dx = 0;
    let dy = 0;
    let sx = null;
    let sy = null;
    if (state.snap && !noSnap) {
      const thr = SNAP_PX / state.zoom;
      sx = snapAxis([box.x0, (box.x0 + box.x1) / 2, box.x1], T.xs, thr);
      sy = snapAxis([box.y0, (box.y0 + box.y1) / 2, box.y1], T.ys, thr);
      if (sx) dx = sx.d;
      if (sy) dy = sy.d;
    }
    if (state.grid && !noSnap) {
      const top = page * app.geom().stride;
      if (!sx) dx = Math.round(box.x0 / GRID) * GRID - box.x0;
      if (!sy) dy = Math.round((box.y0 - top) / GRID) * GRID - (box.y0 - top);
    }
    if (state.snap && !noSnap) addGuides({ x0: box.x0 + dx, x1: box.x1 + dx, y0: box.y0 + dy, y1: box.y1 + dy }, T);
    return { dx, dy };
  }

  // ---------- Sürükleme ----------
  // İşaretçi çalışma alanının üst/alt kenarına (ya da dışına) gelince belge kayar; sürüklenen nesne
  // işaretçinin altında kalır. Böylece resim başka bir sayfaya sürüklenebilir.
  const EDGE = 40;
  function track(onMove, onEnd) {
    let last = null;
    let raf = 0;
    const scrollStep = () => {
      raf = 0;
      if (!last) return;
      const ws = els.workspace;
      const r = ws.getBoundingClientRect();
      const d = last.clientY < r.top + EDGE ? last.clientY - (r.top + EDGE) : last.clientY > r.bottom - EDGE ? last.clientY - (r.bottom - EDGE) : 0;
      if (!d) return;
      const before = ws.scrollTop;
      ws.scrollTop += SS.clamp(Math.round(d / 2), -40, 40) || Math.sign(d);
      if (ws.scrollTop !== before) onMove(last);
      raf = requestAnimationFrame(scrollStep);
    };
    const move = (ev) => {
      last = ev;
      onMove(ev);
      if (!raf) raf = requestAnimationFrame(scrollStep);
    };
    const up = () => {
      last = null;
      cancelAnimationFrame(raf);
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', up);
      document.body.classList.remove('dragging');
      guides = [];
      onEnd();
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', up);
  }

  function startMove(e) {
    const g = app.geom();
    const movers = selImages().filter((i) => !i.locked);
    if (!movers.length) return;
    const start = app.toDoc(e);
    const ids = movers.map((i) => i.id);
    const orig = movers.map((img) => {
      const bb = app.objBox(img);
      return { img, x: img.x, dy: img.page * g.stride + img.y, bx: bb.x - img.x, by: bb.y - img.y, bw: bb.w, bh: bb.h };
    });
    const reflow = movers.some(app.wrapsText);
    let moved = false;
    track(
      (ev) => {
        const p = app.toDoc(ev);
        let dx = p.x - start.x;
        let dy = p.y - start.y;
        if (!moved) {
          if (Math.hypot(dx, dy) * state.zoom < 3) return;
          moved = true;
          document.body.classList.add('dragging');
        }
        const box = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
        for (const o of orig) {
          box.x0 = Math.min(box.x0, o.x + o.bx + dx);
          box.y0 = Math.min(box.y0, o.dy + o.by + dy);
          box.x1 = Math.max(box.x1, o.x + o.bx + o.bw + dx);
          box.y1 = Math.max(box.y1, o.dy + o.by + o.bh + dy);
        }
        guides = [];
        const s = snapBox(box, pageAtY((box.y0 + box.y1) / 2), ids, ev.altKey);
        dx += s.dx;
        dy += s.dy;
        for (const o of orig) setDocPos(o.img, o.x + dx, o.dy + dy);
        movers.forEach(app.updateImageEl);
        if (reflow) app.scheduleLayout();
        app.renderOverlay();
        app.updateCtxBar();
      },
      () => {
        if (moved) {
          movers.forEach(clampToPage);
          app.renderImages();
          app.layout();
          app.commit('edit');
        }
        app.renderOverlay();
        app.updateCtxBar();
      }
    );
  }

  function startHandle(e, h) {
    const img = selImages()[0];
    if (!img || img.locked || state.selection.length !== 1) return;
    const g = app.geom();
    const pageTop = img.page * g.stride;
    const th = ((img.rot || 0) * Math.PI) / 180;
    const cos = Math.cos(th);
    const sin = Math.sin(th);
    const w0 = img.w;
    const h0 = img.h;
    const c0 = { x: img.x + w0 / 2, y: pageTop + img.y + h0 / 2 };
    let moved = false;
    const update = () => {
      moved = true;
      document.body.classList.add('dragging');
      app.updateImageEl(img);
      if (app.wrapsText(img)) app.scheduleLayout();
      app.renderOverlay();
      app.updateCtxBar();
    };
    let onMove;
    if (h === 'rot') {
      onMove = (ev) => {
        const p = app.toDoc(ev);
        let a = ((Math.atan2(p.y - c0.y, p.x - c0.x) * 180) / Math.PI + 90 + 360) % 360;
        if (ev.shiftKey) a = Math.round(a / 15) * 15;
        else for (const s of [0, 90, 180, 270, 360]) if (Math.abs(a - s) < 3) a = s;
        if (a > 180) a -= 360;
        img.rot = SS.round(a, 1);
        update();
      };
    } else {
      const hx = h.includes('e') ? 1 : h.includes('w') ? -1 : 0;
      const hy = h.includes('s') ? 1 : h.includes('n') ? -1 : 0;
      // Sabit kalan karşı nokta (döndürülmüş resimde de doğru çalışır)
      const ax = (-hx * w0) / 2;
      const ay = (-hy * h0) / 2;
      const A = { x: c0.x + ax * cos - ay * sin, y: c0.y + ax * sin + ay * cos };
      const T = snapTargets(img.page, [img.id]);
      onMove = (ev) => {
        const p = app.toDoc(ev);
        const vx = p.x - A.x;
        const vy = p.y - A.y;
        const lx = vx * cos + vy * sin;
        const ly = -vx * sin + vy * cos;
        let nw = hx ? Math.max(MIN, hx * lx) : w0;
        let nh = hy ? Math.max(MIN, hy * ly) : h0;
        const keep = hx !== 0 && hy !== 0 && !ev.shiftKey; // köşeler oranı korur, Shift ile serbest
        if (keep) {
          const s = Math.max(MIN / Math.min(w0, h0), (hx * lx * w0 + hy * ly * h0) / (w0 * w0 + h0 * h0));
          nw = w0 * s;
          nh = h0 * s;
        }
        guides = [];
        const snapping = th === 0 && state.snap && !ev.altKey;
        if (snapping) {
          const thr = SNAP_PX / state.zoom;
          const sx = hx ? snapAxis([A.x + hx * nw], T.xs, thr) : null;
          const sy = hy ? snapAxis([A.y + hy * nh], T.ys, thr) : null;
          if (keep) {
            if (sx && (!sy || Math.abs(sx.d) <= Math.abs(sy.d))) {
              nw = Math.max(MIN, nw + hx * sx.d);
              nh = (nw * h0) / w0;
            } else if (sy) {
              nh = Math.max(MIN, nh + hy * sy.d);
              nw = (nh * w0) / h0;
            }
          } else {
            if (sx) nw = Math.max(MIN, nw + hx * sx.d);
            if (sy) nh = Math.max(MIN, nh + hy * sy.d);
          }
        }
        const ox = (hx * nw) / 2;
        const oy = (hy * nh) / 2;
        const cx = A.x + ox * cos - oy * sin;
        const cy = A.y + ox * sin + oy * cos;
        img.w = nw;
        img.h = nh;
        img.x = cx - nw / 2;
        img.y = cy - nh / 2 - pageTop;
        if (snapping) addGuides({ x0: img.x, x1: img.x + nw, y0: pageTop + img.y, y1: pageTop + img.y + nh }, T);
        update();
      };
    }
    track(onMove, () => {
      if (moved) {
        clampToPage(img);
        app.renderImages();
        app.layout();
        app.commit('edit');
      }
      app.renderOverlay();
      app.updateCtxBar();
    });
  }

  // ---------- Fare olayları ----------
  // pointerdown'da işi yapıyoruz; odak/metin seçimi değişmesin diye ardından gelen mousedown'ı engelliyoruz.
  let blockMouse = false;
  els.doc.addEventListener(
    'pointerdown',
    (e) => {
      blockMouse = false;
      if (e.button !== 0) return;
      const capText = e.target.closest && e.target.closest('.cap-text');
      if (capText && capText.isContentEditable) return; // şekil yazısı düzenleniyor: imleç normal çalışsın
      const handle = e.target.closest && e.target.closest('.handle');
      if (handle) {
        blockMouse = true;
        e.stopPropagation();
        startHandle(e, handle.dataset.h);
        return;
      }
      const id = hitImage(e);
      if (id) {
        blockMouse = true;
        e.stopPropagation();
        if (e.shiftKey || e.ctrlKey || e.metaKey) return app.select(id, true);
        if (!state.selection.includes(id)) app.select(id);
        else dropTextFocus();
        startMove(e);
        return;
      }
      app.clearSelection();
      if (e.target === els.editor || els.editor.contains(e.target)) return;
      blockMouse = true;
      placeCaretNear(e);
    },
    true
  );
  els.doc.addEventListener(
    'mousedown',
    (e) => {
      if (blockMouse) {
        e.preventDefault();
        blockMouse = false;
      }
    },
    true
  );
  // Sayfa dışındaki gri alana tıklanınca seçimi bırak
  els.workspace.addEventListener('pointerdown', (e) => {
    if (!els.doc.contains(e.target)) app.clearSelection();
  });

  // ---------- Klavye ----------
  const objectsFocused = () => {
    const a = document.activeElement;
    return !(a && (a === els.editor || a.isContentEditable || /^(INPUT|SELECT|TEXTAREA)$/.test(a.tagName)));
  };
  app.objectsFocused = () => state.selection.length > 0 && objectsFocused();

  document.addEventListener('keydown', (e) => {
    if (!app.objectsFocused()) return;
    const k = e.key;
    if (k === 'Delete' || k === 'Backspace') {
      e.preventDefault();
      app.deleteSelected();
    } else if (k === 'Escape') app.clearSelection();
    else if (k.startsWith('Arrow') && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      const step = e.shiftKey ? 10 : 1;
      const dx = k === 'ArrowLeft' ? -step : k === 'ArrowRight' ? step : 0;
      const dy = k === 'ArrowUp' ? -step : k === 'ArrowDown' ? step : 0;
      const g = app.geom();
      for (const img of selImages()) if (!img.locked) setDocPos(img, img.x + dx, img.page * g.stride + img.y + dy);
      app.renderImages();
      app.layout();
      app.renderOverlay();
      app.updateCtxBar();
      app.commit('nudge');
    } else if ((e.ctrlKey || e.metaKey) && k.toLowerCase() === 'd') {
      e.preventDefault();
      app.duplicate();
    } else if ((k === 'Enter' || k === 'F2') && state.selection.length === 1) {
      e.preventDefault();
      app.editCaption(state.selection[0]);
    }
  });

  // ---------- Komutlar ----------
  function finishEdit() {
    app.renderImages();
    app.layout();
    app.renderOverlay();
    app.updateCtxBar();
    app.commit('edit');
  }

  app.setWrap = function (wrap) {
    selImages().forEach((i) => (i.wrap = wrap));
    finishEdit();
  };

  app.alignSelected = function (mode) {
    const imgs = selImages().filter((i) => !i.locked);
    if (!imgs.length) return;
    const g = app.geom();
    const bs = imgs.map((img) => {
      const bb = app.objBox(img);
      const top = img.page * g.stride;
      return { img, top, x0: bb.x, x1: bb.x + bb.w, y0: top + bb.y, y1: top + bb.y + bb.h };
    });
    if (mode === 'distH' || mode === 'distV') {
      if (bs.length < 3) return SS.toast('Dağıtmak için en az 3 resim seçin (Shift+tık).');
      const h = mode === 'distH';
      const lo = h ? 'x0' : 'y0';
      const hi = h ? 'x1' : 'y1';
      bs.sort((a, b) => a[lo] + a[hi] - (b[lo] + b[hi]));
      const span = bs[bs.length - 1][hi] - bs[0][lo];
      const gap = (span - bs.reduce((s, b) => s + b[hi] - b[lo], 0)) / (bs.length - 1);
      let pos = bs[0][lo];
      for (const b of bs) {
        const d = pos - b[lo];
        if (h) b.img.x += d;
        else setDocPos(b.img, b.img.x, b.top + b.img.y + d);
        pos += b[hi] - b[lo] + gap;
      }
      return finishEdit();
    }
    let ref;
    if (bs.length === 1) {
      const t = bs[0].top; // tek resim: kenar boşluklarına göre
      ref = { x0: g.m.l, x1: g.PW - g.m.r, y0: t + g.m.t, y1: t + g.PH - g.m.b };
    } else {
      ref = {
        x0: Math.min(...bs.map((b) => b.x0)),
        x1: Math.max(...bs.map((b) => b.x1)),
        y0: Math.min(...bs.map((b) => b.y0)),
        y1: Math.max(...bs.map((b) => b.y1)),
      };
    }
    for (const b of bs) {
      let dx = 0;
      let dy = 0;
      if (mode === 'left') dx = ref.x0 - b.x0;
      else if (mode === 'right') dx = ref.x1 - b.x1;
      else if (mode === 'center') dx = (ref.x0 + ref.x1 - b.x0 - b.x1) / 2;
      else if (mode === 'top') dy = ref.y0 - b.y0;
      else if (mode === 'bottom') dy = ref.y1 - b.y1;
      else if (mode === 'middle') dy = (ref.y0 + ref.y1 - b.y0 - b.y1) / 2;
      setDocPos(b.img, b.img.x + dx, b.top + b.img.y + dy);
    }
    finishEdit();
  };

  app.fitWidth = function () {
    const g = app.geom();
    for (const img of selImages()) {
      if (img.locked) continue;
      const ratio = img.h / img.w;
      img.rot = 0;
      img.w = g.cw;
      img.h = g.cw * ratio;
      img.x = g.m.l;
      clampToPage(img);
    }
    finishEdit();
  };

  app.bringFront = function () {
    let z = maxZ();
    selImages().sort((a, b) => a.z - b.z).forEach((i) => (i.z = ++z));
    finishEdit();
  };
  app.sendBack = function () {
    let z = minZ();
    selImages().sort((a, b) => b.z - a.z).forEach((i) => (i.z = --z));
    finishEdit();
  };
  app.toggleLock = function () {
    const imgs = selImages();
    const lock = imgs.some((i) => !i.locked);
    imgs.forEach((i) => (i.locked = lock));
    finishEdit();
  };
  app.deleteSelected = function () {
    const del = new Set(state.selection);
    if (!del.size) return;
    state.images = state.images.filter((i) => !del.has(i.id));
    state.selection = [];
    finishEdit();
  };

  // target (yapıştırma): { page, y } → grup o sayfaya, üst kenarı y'de olacak biçimde taşınır (x korunur)
  function cloneImages(list, offset, target) {
    let z = maxZ();
    const g = app.geom();
    const ids = [];
    let shift = null;
    if (target) {
      const top = Math.min(...list.map((i) => i.page * g.stride + app.objBox(i).y));
      const bottom = Math.max(...list.map((i) => i.page * g.stride + app.objBox(i).y + app.objBox(i).h));
      // Grup sayfaya sığsın: alttan taşarsa yukarı al
      const y = SS.clamp(target.y, g.m.t, Math.max(g.m.t, g.PH - g.m.b - (bottom - top)));
      shift = target.page * g.stride + y - top;
    }
    for (const src of list) {
      const img = { ...src, id: SS.uid('img'), z: ++z, locked: false };
      if (src.caption) img.caption = { ...src.caption };
      if (shift === null) setDocPos(img, img.x + offset, img.page * g.stride + img.y + offset);
      else {
        img.page = target.page; // grup hedef sayfaya konur
        img.y = src.page * g.stride + src.y + shift - target.page * g.stride;
      }
      clampToPage(img);
      state.images.push(img);
      ids.push(img.id);
    }
    app.renderImages();
    app.layout();
    app.select(ids);
    app.commit('edit');
  }
  app.duplicate = () => selImages().length && cloneImages(selImages(), 16);

  // ---------- Şekil yazısı ----------
  const capLabelSel = document.getElementById('capLabel');
  let defaultLabel = 'Şekil';
  const newCaption = () => ({ text: '', label: defaultLabel, pos: 'below' });

  // Yazıyı sayfa üzerinde düzenle (Enter: bitir, Esc: vazgeç, başka yere tıklamak: bitir)
  app.editCaption = function (id) {
    const img = app.getImage(id);
    if (!img) return;
    if (!img.caption) {
      img.caption = newCaption();
      app.renderImages();
      app.layout();
      app.renderOverlay();
      app.updateCtxBar();
    }
    const t = els.doc.querySelector(`.cap[data-id="${id}"] .cap-text`);
    if (!t) return;
    const before = img.caption.text || '';
    try {
      t.contentEditable = 'plaintext-only';
    } catch (_) {
      t.contentEditable = 'true';
    }
    t.focus({ preventScroll: true });
    const r = document.createRange();
    r.selectNodeContents(t);
    r.collapse(false);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    const onInput = () => {
      app.measureCaption(img);
      app.scheduleLayout();
      app.renderOverlay();
    };
    const onKey = (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        t.blur();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        t.textContent = before;
        t.blur();
      }
    };
    t.addEventListener('input', onInput);
    t.addEventListener('keydown', onKey);
    t.addEventListener(
      'blur',
      () => {
        t.removeEventListener('input', onInput);
        t.removeEventListener('keydown', onKey);
        t.removeAttribute('contenteditable');
        if (img.caption) img.caption.text = t.textContent.replace(/\s+/g, ' ').trim();
        finishEdit();
      },
      { once: true }
    );
  };

  app.toggleCaption = function () {
    const imgs = selImages();
    if (!imgs.length) return;
    if (imgs.length === 1 && !imgs[0].caption) return app.editCaption(imgs[0].id);
    const add = imgs.some((i) => !i.caption);
    imgs.forEach((i) => (add ? (i.caption = i.caption || newCaption()) : delete i.caption));
    finishEdit();
  };

  app.toggleCaptionPos = function () {
    const imgs = selImages().filter((i) => i.caption);
    if (!imgs.length) return SS.toast('Önce "Şekil yazısı" ekleyin.');
    const pos = imgs.every((i) => i.caption.pos === 'above') ? 'below' : 'above';
    imgs.forEach((i) => (i.caption.pos = pos));
    finishEdit();
  };

  capLabelSel.addEventListener('change', () => {
    defaultLabel = capLabelSel.value;
    const imgs = selImages().filter((i) => i.caption);
    if (!imgs.length) return;
    imgs.forEach((i) => (i.caption.label = defaultLabel));
    finishEdit();
  });

  els.doc.addEventListener('dblclick', (e) => {
    const cap = e.target.closest && e.target.closest('.cap');
    if (!cap) return;
    e.preventDefault();
    app.select(cap.dataset.id);
    app.editCaption(cap.dataset.id);
  });

  // ---------- Bağlam çubuğu ----------
  const numInputs = [...document.querySelectorAll('[data-num]')];
  const keepRatio = document.getElementById('keepRatio');
  app.updateCtxBar = function () {
    const imgs = selImages();
    document.body.classList.toggle('img-selected', imgs.length > 0);
    app.onSelectionChange && app.onSelectionChange(imgs);
    if (!imgs.length) return;
    const bar = document.getElementById('ctxbar');
    bar.querySelectorAll('[data-wrap]').forEach((b) => b.classList.toggle('active', imgs.every((i) => i.wrap === b.dataset.wrap)));
    bar.querySelector('[data-cmd="toggleLock"]').classList.toggle('active', imgs.every((i) => i.locked));
    bar.querySelectorAll('[data-align^="dist"]').forEach((b) => (b.disabled = imgs.length < 3));
    const caps = imgs.filter((i) => i.caption);
    bar.querySelector('[data-cmd="toggleCaption"]').classList.toggle('active', caps.length === imgs.length);
    bar.querySelector('[data-cmd="captionPos"]').classList.toggle('active', caps.length > 0 && caps.every((i) => i.caption.pos === 'above'));
    if (caps.length) {
      const l = caps[0].caption.label || 'Şekil';
      if (![...capLabelSel.options].some((o) => o.value === l)) capLabelSel.add(new Option(l, l));
      capLabelSel.value = l;
    }
    const one = imgs.length === 1 ? imgs[0] : null;
    for (const inp of numInputs) {
      inp.disabled = !one;
      if (!one) inp.value = '';
      if (!one || document.activeElement === inp) continue;
      inp.value = fieldValue(one, inp.dataset.num);
    }
  };
  function fieldValue(img, k) {
    if (k === 'rot') return String(SS.round(img.rot || 0, 1));
    if (k === 'page') return String(img.page + 1);
    return String(SS.round(U.pxToCm(img[k]), 2));
  }

  // Elle girilen konumda şeklin tamamı sayfada kalsın (sürüklemedeki gibi yalnızca bir kenarı değil)
  function clampInside(img) {
    const g = app.geom();
    const bb = app.objBox(img);
    img.x += SS.clamp(bb.x, 0, Math.max(0, g.PW - bb.w)) - bb.x;
    img.y += SS.clamp(bb.y, 0, Math.max(0, g.PH - bb.h)) - bb.y;
  }

  numInputs.forEach((inp) => {
    inp.addEventListener('change', () => {
      const imgs = selImages();
      const img = imgs.length === 1 ? imgs[0] : null;
      const v = parseFloat(String(inp.value).replace(',', '.'));
      if (!img || !isFinite(v)) return app.updateCtxBar();
      const k = inp.dataset.num;
      if (k === 'rot') img.rot = ((((v + 180) % 360) + 360) % 360) - 180;
      else if (k === 'page') img.page = SS.clamp(Math.round(v), 1, state.pageCount + 1) - 1; // en çok bir yeni sayfa
      else {
        const px = U.cmToPx(v);
        if (k === 'x') img.x = px;
        else if (k === 'y') img.y = px;
        else if (k === 'w') {
          const nw = Math.max(MIN, px);
          if (keepRatio.checked) img.h *= nw / img.w;
          img.w = nw;
        } else if (k === 'h') {
          const nh = Math.max(MIN, px);
          if (keepRatio.checked) img.w *= nh / img.h;
          img.h = nh;
        }
      }
      if (k === 'x' || k === 'y' || k === 'page') clampInside(img);
      else clampToPage(img);
      finishEdit();
      inp.value = fieldValue(img, k); // odaktaki alan da gerçek (sınırlanmış) değeri göstersin
      if (k === 'page') scrollToImage(img);
    });
    // Enter: uygula ve belgeye dön (Ctrl+Z, oklar resme işlesin); Esc: vazgeç
    inp.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' && e.key !== 'Escape') return;
      e.preventDefault();
      e.stopPropagation(); // odak belgeye dönünce aynı Enter "şekil yazısı düzenle" sayılmasın
      if (e.key === 'Escape') {
        const imgs = selImages();
        inp.value = imgs.length === 1 ? fieldValue(imgs[0], inp.dataset.num) : '';
      }
      inp.blur();
    });
  });

  function scrollToImage(img) {
    const g = app.geom();
    const ws = els.workspace;
    const y = els.scaler.offsetTop + (img.page * g.stride + img.y) * state.zoom;
    if (y < ws.scrollTop + 20 || y > ws.scrollTop + ws.clientHeight - 60) ws.scrollTop = y - 80;
  }

  // ---------- Resim ekleme ----------
  function caretSpot() {
    const g = app.geom();
    const r = app.getCaretRange && app.getCaretRange();
    if (r && els.editor.contains(r.startContainer)) {
      const rects = r.getClientRects();
      const host = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      const rc = rects.length ? rects[0] : host.getBoundingClientRect();
      const docY = (rc.top - els.doc.getBoundingClientRect().top) / state.zoom;
      const page = pageAtY(docY);
      return { page, y: docY - page * g.stride };
    }
    return { page: app.visiblePage ? app.visiblePage() : 0, y: g.m.t };
  }

  function makeImage(assetId, a, at, offset) {
    const g = app.geom();
    let w = Math.min(a.nw, g.cw * 0.8);
    let h = (w * a.nh) / a.nw;
    if (h > g.ch * 0.5) {
      h = g.ch * 0.5;
      w = (h * a.nw) / a.nh;
    }
    let page;
    let x;
    let y;
    let wrap;
    if (at) {
      // Bırakıldığı yere, fare ortada olacak şekilde
      page = at.page;
      x = at.x - w / 2 + offset;
      y = at.y - h / 2 + offset;
      wrap = w > g.cw * 0.6 ? 'topbottom' : 'square';
    } else {
      // İmlecin olduğu satıra, ortalanmış (Word'deki "metinle aynı hizada" eklemeye benzer)
      const c = caretSpot();
      page = c.page;
      x = g.m.l + (g.cw - w) / 2 + offset;
      y = c.y + offset;
      wrap = 'topbottom';
    }
    y = SS.clamp(y, g.m.t, Math.max(g.m.t, g.PH - g.m.b - h));
    x = SS.clamp(x, 0, Math.max(0, g.PW - w));
    return { id: SS.uid('img'), asset: assetId, page, x, y, w, h, rot: 0, wrap, locked: false, z: maxZ() + 1 };
  }

  app.insertFiles = async function (files, at) {
    const list = [...files].filter((f) => /^image\//.test(f.type));
    if (!list.length) return SS.toast('Yalnızca resim dosyaları eklenebilir.');
    const ids = [];
    for (const f of list) {
      try {
        const a = await SS.prepareImage(f);
        const assetId = SS.uid('a');
        state.assets[assetId] = a;
        const img = makeImage(assetId, a, at, ids.length * 18);
        state.images.push(img);
        ids.push(img.id);
      } catch (err) {
        SS.toast('Resim eklenemedi: ' + (f.name || 'pano'));
      }
    }
    if (!ids.length) return;
    app.renderImages();
    app.layout();
    app.select(ids);
    app.commit('edit');
  };

  app.insertImageURLs = async function (urls, at) {
    const files = [];
    for (const u of urls) {
      try {
        const res = await fetch(u);
        const blob = await res.blob();
        if (blob.type.startsWith('image/')) files.push(blob);
      } catch (_) { /* CORS vb. */ }
    }
    if (files.length) app.insertFiles(files, at);
    if (files.length < urls.length) SS.toast('Bazı resimler alınamadı; dosyayı indirip sürükleyerek ekleyin.', 4000);
  };

  // ---------- Pano (kopyala / kes / yapıştır) ----------
  let clip = null;
  const CLIP_MARK = 'serbestsayfa-nesne:';
  function copySel(e, cut) {
    if (!app.objectsFocused()) return;
    const imgs = selImages();
    clip = { id: SS.uid('c'), items: imgs.map((i) => ({ ...i })), cut };
    e.clipboardData.setData('text/plain', CLIP_MARK + clip.id);
    // Başka programlara (ör. Word) yapıştırılabilsin diye HTML olarak da koy
    e.clipboardData.setData(
      'text/html',
      imgs.map((i) => `<img src="${state.assets[i.asset].src}" width="${Math.round(i.w)}" height="${Math.round(i.h)}">`).join('')
    );
    e.preventDefault();
    if (cut) app.deleteSelected();
  }
  document.addEventListener('copy', (e) => copySel(e, false));
  document.addEventListener('cut', (e) => copySel(e, true));

  // Yapıştırma yeri (Word'deki gibi): imleç metindeyse imlecin sayfası ve satırı; değilse ekranda görünen
  // sayfa. Görünen sayfa resimlerin kendi sayfasıysa eski yerleri (kopyada biraz kaydırılmış) kullanılır.
  function pasteTarget(items) {
    if (document.activeElement === els.editor && app.getCaretRange && app.getCaretRange()) return caretSpot();
    const vp = app.visiblePage ? app.visiblePage() : 0;
    if (items.some((i) => i.page === vp)) return null;
    return { page: vp, y: Math.min(...items.map((i) => app.objBox(i).y)) };
  }

  app.pasteObjects = function (e) {
    const dt = e.clipboardData;
    if (clip && dt.getData('text/plain') === CLIP_MARK + clip.id) {
      e.preventDefault();
      const target = pasteTarget(clip.items);
      const offset = clip.cut || target ? 0 : 16;
      clip.cut = false;
      cloneImages(clip.items, offset, target);
      return true;
    }
    const files = [...dt.files].filter((f) => f.type.startsWith('image/'));
    if (!files.length)
      for (const it of dt.items || [])
        if (it.kind === 'file' && it.type.startsWith('image/')) {
          const f = it.getAsFile();
          if (f) files.push(f);
        }
    if (!files.length) return false;
    // Excel/Word gibi uygulamalar metnin yanında bir de resim koyar: metin varsa metni tercih et
    const html = dt.getData('text/html');
    if (html && app.sanitizeHTML(html).text) return false;
    e.preventDefault();
    app.insertFiles(files, null);
    return true;
  };

  // ---------- Dosya sürükle-bırak ----------
  let internalDrag = false;
  els.editor.addEventListener('dragstart', () => (internalDrag = true));
  document.addEventListener('dragend', () => (internalDrag = false));
  els.workspace.addEventListener('dragover', (e) => {
    if (internalDrag) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  });
  els.workspace.addEventListener(
    'drop',
    (e) => {
      if (internalDrag) return;
      e.preventDefault();
      e.stopPropagation();
      const dt = e.dataTransfer;
      const p = app.toDoc(e);
      const page = pageAtY(p.y);
      const at = { page, x: p.x, y: p.y - page * app.geom().stride };
      const docFile = [...dt.files].find((f) => /\.(docx|sayfa)$/i.test(f.name));
      if (docFile) return void (app.openFile && app.openFile(docFile)); // belge bırakıldı: aç
      const files = [...dt.files].filter((f) => f.type.startsWith('image/'));
      if (files.length) return void app.insertFiles(files, at);
      const html = dt.getData('text/html');
      const res = html ? app.sanitizeHTML(html) : null;
      if (res && res.images.length && !res.text) return void app.insertImageURLs(res.images, at);
      const range = caretRangeAt(e.clientX, e.clientY);
      els.editor.focus({ preventScroll: true });
      if (range && els.editor.contains(range.startContainer)) {
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
      }
      if (res) app.insertSanitized(res);
      else if (dt.getData('text/plain')) {
        document.execCommand('insertText', false, dt.getData('text/plain'));
        app.commit('edit');
      }
    },
    true
  );
})();
