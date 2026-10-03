/* ============================================================
   محرر الجدول: يعرض ملف Excel كاملًا ويسمح بالتحديد والحذف والدمج والتعديل.
   التعديلات تتم على نسخة داخل التطبيق فقط — ملفك الأصلي لا يتغير أبدًا.
   ============================================================ */
(function () {
  'use strict';

  const ROW_H = 40, HEAD_H = 40, NUM_W = 56, BUF = 8, MAX_ROWS = 400000;
  const App = () => window.__stockApp;
  const letter = c => App().colLetter(c);
  const esc = s => App().esc(s);

  let root = null, el = {}, key = null, st = null, pushed = false;
  let sel = null, rangeMode = false, msg = '', editTarget = null;
  let undoStack = [], redoStack = [];
  let nrows = 1, ncols = 1, colW = [], colX = [], totalW = 0;
  let lastF = -1, lastL = -2, ticking = false;

  /* ---------- أدوات ---------- */
  const rangeArr = (a, b) => { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; };
  const hit = (a, b) => !(a.r2 < b.r1 || a.r1 > b.r2 || a.c2 < b.c1 || a.c1 > b.c2);
  const getCell = (r, c) => { const row = st.rows[r]; return row ? row[c] : undefined; };
  const disp = v => (v === undefined || v === null) ? '' : (typeof v === 'number' ? App().fmt(v) : String(v));
  const raw = v => (v === undefined || v === null) ? '' : String(v);
  const cellName = (r, c) => letter(c) + (r + 1);
  function countLess(arr, x) { let lo = 0, hi = arr.length; while (lo < hi) { const m = (lo + hi) >> 1; if (arr[m] < x) lo = m + 1; else hi = m; } return lo; }

  function mergeAt(r, c) {
    for (const m of st.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m;
    return null;
  }
  function expand(rect) {
    let changed = true, guard = 0;
    while (changed && guard++ < 50) {
      changed = false;
      for (const m of st.merges) {
        if (hit(m, rect) && (m.r1 < rect.r1 || m.c1 < rect.c1 || m.r2 > rect.r2 || m.c2 > rect.c2)) {
          rect = { r1: Math.min(rect.r1, m.r1), c1: Math.min(rect.c1, m.c1), r2: Math.max(rect.r2, m.r2), c2: Math.max(rect.c2, m.c2) };
          changed = true;
        }
      }
    }
    return rect;
  }

  /* ---------- بناء الواجهة ---------- */
  function build() {
    root = document.createElement('div');
    root.id = 'editorView';
    root.hidden = true;
    root.innerHTML = `
      <div class="ed-top">
        <div class="ttl" id="ed-title"></div>
        <button class="btn primary" id="ed-done" type="button">✓ تم</button>
      </div>
      <div class="ed-fbar">
        <span class="ed-addr" id="ed-addr">—</span>
        <input id="ed-inp" type="text" disabled placeholder="اختر خلية لتعديل محتواها" autocomplete="off">
        <button class="btn" id="ed-ok" type="button">✓</button>
      </div>
      <div class="ed-grid" id="ed-grid">
        <div class="ed-scene" id="ed-scene">
          <div class="ed-hdr" id="ed-hdr"></div>
          <div class="ed-body"><div class="ed-nums" id="ed-nums"></div><div class="ed-cells" id="ed-cells"></div></div>
        </div>
      </div>
      <div class="ed-status" id="ed-status" role="status"></div>
      <div class="ed-tools" id="ed-tools">
        <button class="btn" data-act="undo" type="button">↶ تراجع</button>
        <button class="btn" data-act="redo" type="button">↷ إعادة</button>
        <button class="btn" data-act="range" type="button">تحديد نطاق: إيقاف</button>
        <button class="btn danger" data-act="delcols" type="button">🗑 حذف الأعمدة</button>
        <button class="btn danger" data-act="delrows" type="button">🗑 حذف الصفوف</button>
        <button class="btn" data-act="clear" type="button">🧹 مسح المحتوى</button>
        <button class="btn" data-act="merge" type="button">⛓ دمج الخلايا</button>
        <button class="btn" data-act="unmerge" type="button">✂ فك الدمج</button>
        <button class="btn" data-act="insrow" type="button">➕ إدراج صف</button>
        <button class="btn" data-act="inscol" type="button">➕ إدراج عمود</button>
        <button class="btn" data-act="asname" type="button">🏷 هذا عمود الاسم</button>
        <button class="btn" data-act="asqty" type="button">🔢 هذا عمود الكمية</button>
        <button class="btn" data-act="ashead" type="button">📌 هذا صف العناوين</button>
        <button class="btn" data-act="save" type="button">⬇ حفظ نسخة معدّلة</button>
        <button class="btn" data-act="reload" type="button">↺ الملف الأصلي</button>
      </div>`;
    document.body.appendChild(root);
    const q = s => root.querySelector(s);
    el = {
      title: q('#ed-title'), done: q('#ed-done'), addr: q('#ed-addr'), inp: q('#ed-inp'), ok: q('#ed-ok'),
      grid: q('#ed-grid'), scene: q('#ed-scene'), hdr: q('#ed-hdr'), nums: q('#ed-nums'), cells: q('#ed-cells'),
      status: q('#ed-status'), tools: q('#ed-tools')
    };

    el.done.addEventListener('click', () => closeEditor());
    el.grid.addEventListener('scroll', () => {
      if (ticking) return;
      ticking = true;
      requestAnimationFrame(() => { ticking = false; renderWindow(false); });
    }, { passive: true });
    el.grid.addEventListener('click', onGridClick);
    el.tools.addEventListener('click', e => {
      const b = e.target.closest('[data-act]');
      if (b && !b.disabled) act(b.dataset.act);
    });
    el.inp.addEventListener('change', applyInput);
    el.inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); el.inp.blur(); } });
    el.ok.addEventListener('click', () => { applyInput(); el.inp.blur(); });
    window.addEventListener('resize', () => renderWindow(true));
    window.addEventListener('popstate', () => { if (root && !root.hidden) { pushed = false; closeEditor(true); } });
    document.addEventListener('keydown', e => {
      if (!root || root.hidden || e.target === el.inp) return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); doUndo(); }
      else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') { e.preventDefault(); doRedo(); }
      else if (e.key === 'Delete' && sel) { e.preventDefault(); act('clear'); }
    });
  }

  /* ---------- فتح وإغلاق ---------- */
  function openEditor(k) {
    const A = App();
    if (!A) return;
    const s = A.state[k];
    if (!s || !s.rows) return;
    key = k; st = s;
    if (!root) build();
    undoStack = []; redoStack = []; sel = null; rangeMode = false; msg = '';
    el.title.textContent = st.def.name + ' — ' + st.fileName;
    root.hidden = false;
    document.body.classList.add('noscroll');
    if (!pushed) { try { history.pushState({ ed: 1 }, ''); pushed = true; } catch (_) { /* لا بأس */ } }
    el.grid.scrollTop = 0; el.grid.scrollLeft = 0;
    if (st.rows.length > MAX_ROWS) msg = 'الملف كبير جدًا: يُعرض أول ' + MAX_ROWS + ' صف فقط.';
    afterStructure();
  }

  function closeEditor(fromPop) {
    if (!root || root.hidden) return;
    applyInput();
    root.hidden = true;
    document.body.classList.remove('noscroll');
    if (pushed && !fromPop) { pushed = false; try { history.back(); } catch (_) { /* لا بأس */ } }
    const k = key;
    key = null;
    undoStack = []; redoStack = [];
    if (k) App().afterEdit(k);
  }

  /* ---------- القياسات ---------- */
  function computeSizes() {
    nrows = Math.max(1, Math.min(st.rows.length, MAX_ROWS));
    let n = 1;
    for (let r = 0; r < st.rows.length; r++) { const row = st.rows[r]; if (row && row.length > n) n = row.length; }
    for (const m of st.merges) if (m.c2 + 1 > n) n = m.c2 + 1;
    ncols = n;
    const lens = new Array(ncols).fill(3);
    const lim = Math.min(st.rows.length, 300);
    for (let r = 0; r < lim; r++) {
      const row = st.rows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) { const v = row[c]; if (v !== undefined && v !== null) { const l = String(v).length; if (l > lens[c]) lens[c] = l; } }
    }
    colW = lens.map(l => Math.max(76, Math.min(260, l * 9 + 24)));
    colX = []; let x = 0;
    for (let c = 0; c < ncols; c++) { colX.push(x); x += colW[c]; }
    totalW = x;
  }

  function afterStructure() {
    computeSizes();
    const totalH = nrows * ROW_H;
    el.scene.style.width = (NUM_W + totalW) + 'px';
    el.nums.style.height = totalH + 'px';
    el.cells.style.height = totalH + 'px';
    el.cells.style.width = totalW + 'px';
    renderHeader();
    renderWindow(true);
    updateUI();
  }

  /* ---------- الرسم ---------- */
  const roleCols = () => ({
    n: st.nameCol === '' ? -1 : parseInt(st.nameCol, 10),
    q: st.qtyCol === '' ? -1 : parseInt(st.qtyCol, 10)
  });

  function colSel(c) {
    if (!sel) return false;
    if (sel.type === 'cols') return sel.cols.has(c);
    if (sel.type === 'cells') return c >= sel.c1 && c <= sel.c2;
    return false;
  }
  function rowSel(r) {
    if (!sel) return false;
    if (sel.type === 'rows') return sel.rows.has(r);
    if (sel.type === 'cells') return r >= sel.r1 && r <= sel.r2;
    return false;
  }
  function areaSel(r1, c1, r2, c2) {
    if (!sel) return false;
    if (sel.type === 'cols') { for (let c = c1; c <= c2; c++) if (sel.cols.has(c)) return true; return false; }
    if (sel.type === 'rows') { for (let r = r1; r <= r2; r++) if (sel.rows.has(r)) return true; return false; }
    return !(r2 < sel.r1 || r1 > sel.r2 || c2 < sel.c1 || c1 > sel.c2);
  }

  function renderHeader() {
    const rc = roleCols();
    let h = '<div class="ed-corner" data-corner="1" title="تحديد الكل">▦</div>';
    for (let c = 0; c < ncols; c++) {
      const badge = c === rc.n ? '<i>اسم</i>' : (c === rc.q ? '<i class="q">كمية</i>' : '');
      h += `<div class="ed-colh${colSel(c) ? ' sel' : ''}" data-cn="${c}" style="width:${colW[c]}px">${letter(c)}${badge}</div>`;
    }
    el.hdr.innerHTML = h;
  }

  function renderWindow(force) {
    if (!root || root.hidden || !st) return;
    const top = el.grid.scrollTop, vh = el.grid.clientHeight || 400;
    let f = Math.floor((top - HEAD_H) / ROW_H) - BUF; if (f < 0) f = 0;
    let l = Math.ceil((top + vh - HEAD_H) / ROW_H) + BUF; if (l > nrows - 1) l = nrows - 1;
    if (!force && f === lastF && l === lastL) return;
    lastF = f; lastL = l;

    const K = ncols + 1, covered = new Set(), tl = new Set(), inWin = [];
    for (const m of st.merges) {
      if (m.r2 < f || m.r1 > l || m.r1 >= nrows) continue;
      inWin.push(m); tl.add(m.r1 * K + m.c1);
      const ra = Math.max(f, m.r1), rb = Math.min(l, m.r2);
      for (let r = ra; r <= rb; r++) for (let c = m.c1; c <= m.c2; c++) if (!(r === m.r1 && c === m.c1)) covered.add(r * K + c);
    }

    const rc = roleCols(), hr = st.headerRow;
    const roleCls = (r, c) => (c === rc.n ? ' cn' : '') + (c === rc.q ? ' cq' : '') + (r === hr ? ' hr' : '');
    const out = [], nums = [];
    for (let r = f; r <= l; r++) {
      nums.push(`<div class="ed-n${rowSel(r) ? ' sel' : ''}${r === hr ? ' hr' : ''}" data-rn="${r}" style="top:${r * ROW_H}px;height:${ROW_H}px">${r + 1}</div>`);
      const row = st.rows[r];
      for (let c = 0; c < ncols; c++) {
        const k = r * K + c;
        if (covered.has(k) || tl.has(k)) continue;
        const v = row ? row[c] : undefined;
        const act = sel && sel.type === 'cells' && sel.ar === r && sel.ac === c;
        out.push(`<div class="ed-c${roleCls(r, c)}${areaSel(r, c, r, c) ? ' sel' : ''}${act ? ' act' : ''}" data-r="${r}" data-c="${c}" style="top:${r * ROW_H}px;right:${colX[c]}px;width:${colW[c]}px;height:${ROW_H}px"><span dir="auto">${esc(disp(v))}</span></div>`);
      }
    }
    for (const m of inWin) {
      let w = 0; for (let c = m.c1; c <= m.c2; c++) w += colW[c];
      const act = sel && sel.type === 'cells' && sel.ar === m.r1 && sel.ac === m.c1;
      out.push(`<div class="ed-c mg${roleCls(m.r1, m.c1)}${areaSel(m.r1, m.c1, m.r2, m.c2) ? ' sel' : ''}${act ? ' act' : ''}" data-r="${m.r1}" data-c="${m.c1}" style="top:${m.r1 * ROW_H}px;right:${colX[m.c1]}px;width:${w}px;height:${(m.r2 - m.r1 + 1) * ROW_H}px"><span dir="auto">${esc(disp(getCell(m.r1, m.c1)))}</span></div>`);
    }
    el.cells.innerHTML = out.join('');
    el.nums.innerHTML = nums.join('');
  }

  /* ---------- التحديد ---------- */
  function onGridClick(e) {
    const t = e.target.closest('[data-r],[data-rn],[data-cn],[data-corner]');
    if (!t) return;
    applyInput();
    if (t.dataset.corner) {
      sel = { type: 'cells', ar: 0, ac: 0, r1: 0, c1: 0, r2: nrows - 1, c2: ncols - 1 };
    } else if (t.dataset.cn !== undefined) {
      const c = parseInt(t.dataset.cn, 10);
      if (sel && sel.type === 'cols') {
        if (rangeMode && sel.anchor !== undefined) {
          const a = sel.anchor;
          sel = { type: 'cols', cols: new Set(rangeArr(Math.min(a, c), Math.max(a, c))), anchor: a };
        } else if (sel.cols.has(c)) { sel.cols.delete(c); if (!sel.cols.size) sel = null; else sel.anchor = c; }
        else { sel.cols.add(c); sel.anchor = c; }
      } else sel = { type: 'cols', cols: new Set([c]), anchor: c };
    } else if (t.dataset.rn !== undefined) {
      const r = parseInt(t.dataset.rn, 10);
      if (sel && sel.type === 'rows') {
        if (rangeMode && sel.anchor !== undefined) {
          const a = sel.anchor;
          sel = { type: 'rows', rows: new Set(rangeArr(Math.min(a, r), Math.max(a, r))), anchor: a };
        } else if (sel.rows.has(r)) { sel.rows.delete(r); if (!sel.rows.size) sel = null; else sel.anchor = r; }
        else { sel.rows.add(r); sel.anchor = r; }
      } else sel = { type: 'rows', rows: new Set([r]), anchor: r };
    } else {
      let r = parseInt(t.dataset.r, 10), c = parseInt(t.dataset.c, 10);
      const m = mergeAt(r, c);
      if (m) { r = m.r1; c = m.c1; }
      if ((e.shiftKey || rangeMode) && sel && sel.type === 'cells') {
        const rect = expand({ r1: Math.min(sel.ar, r), c1: Math.min(sel.ac, c), r2: Math.max(sel.ar, r), c2: Math.max(sel.ac, c) });
        sel = Object.assign({ type: 'cells', ar: sel.ar, ac: sel.ac }, rect);
      } else {
        const rect = m ? { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 } : { r1: r, c1: c, r2: r, c2: c };
        sel = Object.assign({ type: 'cells', ar: r, ac: c }, rect);
      }
    }
    msg = '';
    renderHeader();
    renderWindow(true);
    updateUI();
  }

  const sortedNum = s => [...s].sort((a, b) => a - b);
  function rowsSpanned() {
    if (!sel) return [];
    if (sel.type === 'rows') return sortedNum(sel.rows);
    if (sel.type === 'cells') return rangeArr(sel.r1, Math.min(sel.r2, st.rows.length ? st.rows.length - 1 : 0));
    return [];
  }
  function colsSpanned() {
    if (!sel) return [];
    if (sel.type === 'cols') return sortedNum(sel.cols);
    if (sel.type === 'cells') return rangeArr(sel.c1, sel.c2);
    return [];
  }

  function describeSel() {
    if (!sel) return 'المس رأس عمود (A, B…) أو رقم صف أو خلية لتحديدها. يمكنك تحديد أكثر من عمود أو صف.';
    if (sel.type === 'cols') {
      const a = sortedNum(sel.cols);
      return a.length === 1 ? 'تم تحديد العمود ' + letter(a[0]) : 'تم تحديد ' + a.length + ' أعمدة: ' + a.slice(0, 8).map(letter).join('، ') + (a.length > 8 ? '…' : '');
    }
    if (sel.type === 'rows') {
      const a = sortedNum(sel.rows);
      return a.length === 1 ? 'تم تحديد الصف ' + (a[0] + 1) : 'تم تحديد ' + a.length + ' صفوف: ' + a.slice(0, 8).map(x => x + 1).join('، ') + (a.length > 8 ? '…' : '');
    }
    const n = (sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1);
    if (n === 1) return 'الخلية ' + cellName(sel.r1, sel.c1);
    return 'النطاق ' + cellName(sel.r1, sel.c1) + ':' + cellName(sel.r2, sel.c2) + ' (' + n + ' خلية)';
  }

  function updateUI() {
    // شريط تعديل الخلية
    let target = null;
    if (sel && sel.type === 'cells') {
      const m = mergeAt(sel.ar, sel.ac);
      const single = (sel.r1 === sel.r2 && sel.c1 === sel.c2) || (m && m.r1 === sel.r1 && m.c1 === sel.c1 && m.r2 === sel.r2 && m.c2 === sel.c2);
      if (single) target = { r: sel.ar, c: sel.ac };
    }
    editTarget = target;
    el.inp.disabled = !target;
    el.addr.textContent = target ? cellName(target.r, target.c) : '—';
    el.inp.value = target ? raw(getCell(target.r, target.c)) : '';

    el.status.textContent = msg || describeSel();

    const rows = rowsSpanned(), cols = colsSpanned();
    const btn = a => el.tools.querySelector('[data-act="' + a + '"]');
    const set = (a, enabled, label) => { const b = btn(a); b.disabled = !enabled; if (label) b.textContent = label; };
    const cnt = n => n > 1 ? ' (' + n + ')' : '';

    set('undo', undoStack.length > 0);
    set('redo', redoStack.length > 0);
    btn('range').textContent = 'تحديد نطاق: ' + (rangeMode ? 'تشغيل' : 'إيقاف');
    btn('range').classList.toggle('on', rangeMode);
    set('delcols', cols.length > 0, '🗑 حذف ' + (cols.length > 1 ? 'الأعمدة' : 'العمود') + cnt(cols.length));
    set('delrows', rows.length > 0, '🗑 حذف ' + (rows.length > 1 ? 'الصفوف' : 'الصف') + cnt(rows.length));
    set('clear', !!sel);
    const exactMerge = sel && sel.type === 'cells' && st.merges.some(m => m.r1 === sel.r1 && m.c1 === sel.c1 && m.r2 === sel.r2 && m.c2 === sel.c2);
    set('merge', !!(sel && sel.type === 'cells' && (sel.r2 > sel.r1 || sel.c2 > sel.c1) && !exactMerge));
    set('unmerge', !!(sel && st.merges.some(m => areaSel(m.r1, m.c1, m.r2, m.c2))));
    set('insrow', rows.length > 0);
    set('inscol', cols.length > 0);
    set('asname', cols.length === 1, cols.length === 1 && st.nameCol === String(cols[0]) ? '🏷 إلغاء عمود الاسم' : '🏷 هذا عمود الاسم');
    set('asqty', cols.length === 1, cols.length === 1 && st.qtyCol === String(cols[0]) ? '🔢 إلغاء عمود الكمية' : '🔢 هذا عمود الكمية');
    set('ashead', rows.length === 1, rows.length === 1 && st.headerRow === rows[0] ? '📌 إلغاء صف العناوين' : '📌 هذا صف العناوين');
  }

  /* ---------- التعديل والتراجع ---------- */
  function snapshot() {
    return {
      rows: st.rows.map(r => (r ? r.slice() : r)),
      merges: st.merges.map(m => Object.assign({}, m)),
      headerRow: st.headerRow, nameCol: st.nameCol, qtyCol: st.qtyCol
    };
  }
  const maxUndo = () => (st.rows.length > 50000 ? 3 : 15);
  function pushUndo() { undoStack.push(snapshot()); if (undoStack.length > maxUndo()) undoStack.shift(); }
  function restore(s) {
    st.rows = s.rows; st.merges = s.merges; st.headerRow = s.headerRow; st.nameCol = s.nameCol; st.qtyCol = s.qtyCol;
    sel = null;
    afterStructure();
  }
  function doUndo() { const s = undoStack.pop(); if (!s) return; redoStack.push(snapshot()); msg = 'تم التراجع'; restore(s); }
  function doRedo() { const s = redoStack.pop(); if (!s) return; undoStack.push(snapshot()); msg = 'تمت الإعادة'; restore(s); }

  function edit(fn, note, dirty) {
    pushUndo();
    fn();
    if (dirty !== false) st.edited = true;
    redoStack = [];
    msg = note || '';
    afterStructure();
  }

  function applyInput() {
    if (!editTarget || !root || root.hidden) return;
    const { r, c } = editTarget;
    const s = el.inp.value;
    if (s === raw(getCell(r, c))) return;
    pushUndo();
    const t = s.trim();
    let v;
    if (t === '') v = undefined;
    else if (/^-?(0|[1-9]\d*)(\.\d+)?$/.test(t)) v = Number(t);
    else v = t;
    if (!st.rows[r]) st.rows[r] = [];
    st.rows[r][c] = v;
    st.edited = true; redoStack = []; msg = 'تم تعديل ' + cellName(r, c);
    renderWindow(true);
    updateUI();
  }

  function remapMerges(axis, removedSorted) {
    const out = [];
    for (const m of st.merges) {
      const a1 = axis === 'r' ? m.r1 : m.c1, a2 = axis === 'r' ? m.r2 : m.c2;
      const span = a2 - a1 + 1;
      const inSpan = countLess(removedSorted, a2 + 1) - countLess(removedSorted, a1);
      if (inSpan >= span) continue;
      const n1 = a1 - countLess(removedSorted, a1), n2 = n1 + (span - inSpan) - 1;
      const nm = axis === 'r' ? { r1: n1, r2: n2, c1: m.c1, c2: m.c2 } : { r1: m.r1, r2: m.r2, c1: n1, c2: n2 };
      if (nm.r1 === nm.r2 && nm.c1 === nm.c2) continue;
      out.push(nm);
    }
    st.merges = out;
  }
  function remapRole(val, removedSorted) {
    if (val === '') return '';
    const x = parseInt(val, 10);
    if (removedSorted.includes(x)) return '';
    return String(x - countLess(removedSorted, x));
  }

  function removeCols(cols) {
    const asc = cols.slice().sort((a, b) => a - b), desc = asc.slice().reverse();
    for (let r = 0; r < st.rows.length; r++) {
      const row = st.rows[r];
      if (!row) continue;
      for (const c of desc) if (c < row.length) row.splice(c, 1);
    }
    remapMerges('c', asc);
    st.nameCol = remapRole(st.nameCol, asc);
    st.qtyCol = remapRole(st.qtyCol, asc);
  }
  function removeRows(rows) {
    const asc = rows.slice().sort((a, b) => a - b), set = new Set(asc);
    const out = [];
    for (let r = 0; r < st.rows.length; r++) if (!set.has(r)) out.push(st.rows[r]);
    st.rows = out;
    remapMerges('r', asc);
    if (st.headerRow >= 0) {
      const wasDeleted = set.has(st.headerRow);
      st.headerRow = st.headerRow - countLess(asc, st.headerRow) - (wasDeleted ? 1 : 0);
    }
  }

  async function act(a) {
    const rows = rowsSpanned(), cols = colsSpanned();
    switch (a) {
      case 'undo': doUndo(); return;
      case 'redo': doRedo(); return;
      case 'range': rangeMode = !rangeMode; msg = ''; updateUI(); return;
      case 'delcols':
        if (!cols.length) return;
        edit(() => { removeCols(cols); sel = null; }, 'تم حذف ' + cols.length + ' عمود — يمكنك التراجع');
        return;
      case 'delrows':
        if (!rows.length) return;
        edit(() => { removeRows(rows); sel = null; }, 'تم حذف ' + rows.length + ' صف — يمكنك التراجع');
        return;
      case 'clear':
        if (!sel) return;
        edit(() => {
          if (sel.type === 'cells') {
            for (let r = sel.r1; r <= sel.r2; r++) { const row = st.rows[r]; if (row) for (let c = sel.c1; c <= sel.c2; c++) row[c] = undefined; }
          } else if (sel.type === 'cols') {
            for (let r = 0; r < st.rows.length; r++) { const row = st.rows[r]; if (row) for (const c of sel.cols) row[c] = undefined; }
          } else {
            for (const r of sel.rows) st.rows[r] = undefined;
          }
        }, 'تم مسح المحتوى — يمكنك التراجع');
        return;
      case 'merge': {
        if (!sel || sel.type !== 'cells') return;
        const rect = expand({ r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 });
        edit(() => {
          st.merges = st.merges.filter(m => !hit(m, rect));
          for (let r = rect.r1; r <= rect.r2; r++) {
            const row = st.rows[r];
            if (!row) continue;
            for (let c = rect.c1; c <= rect.c2; c++) if (!(r === rect.r1 && c === rect.c1)) row[c] = undefined;
          }
          st.merges.push({ r1: rect.r1, c1: rect.c1, r2: rect.r2, c2: rect.c2 });
          sel = Object.assign({ type: 'cells', ar: rect.r1, ac: rect.c1 }, rect);
        }, 'تم الدمج (تبقى قيمة الخلية الأولى فقط) — يمكنك التراجع');
        return;
      }
      case 'unmerge':
        edit(() => { st.merges = st.merges.filter(m => !areaSel(m.r1, m.c1, m.r2, m.c2)); }, 'تم فك الدمج');
        return;
      case 'insrow': {
        if (!rows.length) return;
        const r = rows[0];
        edit(() => {
          st.rows.splice(r, 0, undefined);
          for (const m of st.merges) { if (m.r1 >= r) { m.r1++; m.r2++; } else if (m.r2 >= r) m.r2++; }
          if (st.headerRow >= r) st.headerRow++;
          sel = { type: 'rows', rows: new Set([r]), anchor: r };
        }, 'تم إدراج صف فارغ');
        return;
      }
      case 'inscol': {
        if (!cols.length) return;
        const c = cols[0];
        edit(() => {
          for (let r = 0; r < st.rows.length; r++) { const row = st.rows[r]; if (row && c < row.length) row.splice(c, 0, undefined); }
          for (const m of st.merges) { if (m.c1 >= c) { m.c1++; m.c2++; } else if (m.c2 >= c) m.c2++; }
          if (st.nameCol !== '' && parseInt(st.nameCol, 10) >= c) st.nameCol = String(parseInt(st.nameCol, 10) + 1);
          if (st.qtyCol !== '' && parseInt(st.qtyCol, 10) >= c) st.qtyCol = String(parseInt(st.qtyCol, 10) + 1);
          sel = { type: 'cols', cols: new Set([c]), anchor: c };
        }, 'تم إدراج عمود فارغ');
        return;
      }
      case 'asname':
        if (cols.length !== 1) return;
        edit(() => {
          if (st.nameCol === String(cols[0])) st.nameCol = '';
          else { st.nameCol = String(cols[0]); if (st.qtyCol === st.nameCol) st.qtyCol = ''; }
        }, st.nameCol === String(cols[0]) ? 'تم تحديد عمود الاسم: ' + letter(cols[0]) : 'تم إلغاء عمود الاسم', false);
        return;
      case 'asqty':
        if (cols.length !== 1) return;
        edit(() => {
          if (st.qtyCol === String(cols[0])) st.qtyCol = '';
          else { st.qtyCol = String(cols[0]); if (st.nameCol === st.qtyCol) st.nameCol = ''; }
        }, st.qtyCol === String(cols[0]) ? 'تم تحديد عمود الكمية: ' + letter(cols[0]) : 'تم إلغاء عمود الكمية', false);
        return;
      case 'ashead':
        if (rows.length !== 1) return;
        edit(() => { st.headerRow = (st.headerRow === rows[0]) ? -1 : rows[0]; },
          'المواد تبدأ من الصف ' + (rows[0] + 2) + ' فما بعد (الصف ' + (rows[0] + 1) + ' عناوين)', false);
        if (st.headerRow === -1) { msg = 'لا يوجد صف عناوين: المواد تبدأ من الصف 1'; updateUI(); }
        return;
      case 'save': {
        msg = 'جارٍ تجهيز الملف…'; updateUI();
        try {
          const blob = await buildXlsx();
          // اسم إنكليزي ثابت حتى لا يضيع الامتداد على بعض الأجهزة
          download(blob, (key === 'gh' ? 'ghadeer' : 'hirfiyyin') + '-edited.xlsx');
          msg = 'تم حفظ نسخة معدّلة (الأصل لم يتغير)';
        } catch (e) { msg = 'تعذّر إنشاء الملف.'; }
        updateUI();
        return;
      }
      case 'reload': {
        if (!confirm('سيتم إلغاء كل تعديلاتك واسترجاع الملف الأصلي. هل أنت متأكد؟')) return;
        msg = 'جارٍ إعادة تحميل الملف…'; updateUI();
        let err = '';
        try { await App().reloadOriginal(key); } catch (e) { err = e.message || ''; }
        if (!st.rows) { closeEditor(); return; }
        undoStack = []; redoStack = []; sel = null;
        msg = err || 'تم استرجاع الملف الأصلي';
        afterStructure();
        return;
      }
    }
  }

  /* ---------- حفظ نسخة xlsx (محليًا بدون إنترنت) ---------- */
  async function buildXlsx() {
    const A = App();
    const xesc = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const letters = [];
    const L = c => letters[c] || (letters[c] = letter(c));
    const parts = [];
    for (let r = 0; r < st.rows.length; r++) {
      const row = st.rows[r];
      if (!row) continue;
      let cells = '';
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (v === undefined || v === null || v === '') continue;
        const ref = L(c) + (r + 1);
        if (typeof v === 'number' && isFinite(v)) cells += `<c r="${ref}"><v>${v}</v></c>`;
        else cells += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
      }
      if (cells) parts.push(`<row r="${r + 1}">${cells}</row>`);
      if (r % 5000 === 4999) await A.tick();
    }
    const mg = st.merges.length
      ? `<mergeCells count="${st.merges.length}">${st.merges.map(m => `<mergeCell ref="${L(m.c1)}${m.r1 + 1}:${L(m.c2)}${m.r2 + 1}"/>`).join('')}</mergeCells>` : '';
    const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews><sheetData>' +
      parts.join('') + '</sheetData>' + mg + '</worksheet>';
    let name = (st.wb && st.wb.names && st.wb.names[st.sheetIdx]) || 'Sheet1';
    name = name.replace(/[\[\]:*?\/\\]/g, ' ').trim().slice(0, 31) || 'Sheet1';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + xesc(name) + '" sheetId="1" r:id="rId1"/></sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    zip.file('xl/worksheets/sheet1.xml', sheet);
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 4000);
  }

  window.openEditor = openEditor;
})();
