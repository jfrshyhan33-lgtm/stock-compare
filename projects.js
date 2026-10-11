/* ============================================================
   حفظ المشاريع: كل شيء (الملفات المستوردة + تعديلاتك + الإعدادات) يُحفظ تلقائيًا
   داخل ذاكرة المتصفح على جهازك (IndexedDB) — لا يُرفع شيء إلى الإنترنت.
   ============================================================ */
(function () {
  'use strict';

  const App = () => window.__stockApp;
  const KEYS = ['gh', 'hr'];
  const esc = s => App().esc(s);
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  let db = null, cur = null, dirty = false, timer = null, saving = false, again = false, restoring = false;
  let savedBuf = { gh: null, hr: null }, listView = null, pushed = false;

  /* ---------- قاعدة البيانات ---------- */
  function openDB() {
    return new Promise((res, rej) => {
      if (!window.indexedDB) return rej(new Error('no indexedDB'));
      const r = indexedDB.open('stockCompareDB', 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        d.createObjectStore('meta', { keyPath: 'id' });
        d.createObjectStore('state', { keyPath: 'id' });
        d.createObjectStore('files', { keyPath: 'fid' });
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
      r.onblocked = () => rej(new Error('blocked'));
    });
  }
  const rq = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  const getOne = (store, key) => rq(db.transaction(store).objectStore(store).get(key));
  const getAll = store => rq(db.transaction(store).objectStore(store).getAll());
  function write(ops) {
    return new Promise((res, rej) => {
      const names = [...new Set(ops.map(o => o.store))];
      const t = db.transaction(names, 'readwrite');
      t.oncomplete = () => res();
      t.onerror = () => rej(t.error);
      t.onabort = () => rej(t.error || new Error('abort'));
      for (const o of ops) {
        const os = t.objectStore(o.store);
        if (o.del !== undefined) os.delete(o.del); else os.put(o.value);
      }
    });
  }

  /* ---------- الواجهة العلوية ---------- */
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const stamp = d => App().dateStr(d) + ' ' + App().timeStr(d);
  const defaultName = () => 'مشروع ' + stamp(new Date());

  function buildBar() {
    const top = document.querySelector('.top');
    const bar = document.createElement('div');
    bar.className = 'projbar';
    bar.innerHTML = `
      <div class="pj-row">
        <button type="button" id="pj-list">📁 المشاريع</button>
        <button type="button" id="pj-new">➕ مشروع جديد</button>
        <button type="button" id="pj-save">💾 حفظ الآن</button>
      </div>
      <button type="button" class="pj-name" id="pj-name"></button>
      <div class="pj-state" id="pj-state" role="status"></div>`;
    top.appendChild(bar);
    document.getElementById('pj-list').addEventListener('click', showList);
    document.getElementById('pj-new').addEventListener('click', newProject);
    document.getElementById('pj-save').addEventListener('click', async () => {
      if (!cur && !hasData()) { setState('لا يوجد شيء لحفظه بعد — استورد ملفًا أولًا.'); return; }
      dirty = true; clearTimeout(timer); await save();
    });
    document.getElementById('pj-name').addEventListener('click', renameCurrent);
    updateBar();
  }

  function setState(text, cls) {
    const e = document.getElementById('pj-state');
    if (!e) return;
    e.textContent = text || '';
    e.className = 'pj-state' + (cls ? ' ' + cls : '');
  }
  function updateBar() {
    const b = document.getElementById('pj-name');
    if (!b) return;
    b.textContent = cur ? '📌 ' + cur.name + '  ✎' : '📌 مشروع بدون اسم — يُحفظ تلقائيًا عند استيراد ملف';
  }

  const hasData = () => { const S = App().state; return !!(S.gh.rows || S.hr.rows); };

  /* ---------- الحفظ ---------- */
  function schedule(ms) {
    if (restoring || !db) return;
    dirty = true;
    clearTimeout(timer);
    timer = setTimeout(save, ms === undefined ? 1800 : ms);
  }

  async function save() {
    if (!db || restoring) return;
    if (saving) { again = true; return; }
    if (!cur) {
      if (!hasData()) { dirty = false; return; }
      cur = { id: newId(), name: defaultName(), created: Date.now() };
      updateBar();
    }
    saving = true; dirty = false;
    setState('جارٍ الحفظ…', 'busy');
    try {
      const A = App(), S = A.state, ops = [], stores = {}, info = {}, newSaved = Object.assign({}, savedBuf);
      for (const k of KEYS) {
        const st = S[k];
        if (st.rows && st.srcBuf) {
          stores[k] = { fileName: st.fileName, sheetIdx: st.sheetIdx, rows: st.rows, merges: st.merges, headerRow: st.headerRow,
            nameCol: st.nameCol, qtyCol: st.qtyCol, edited: st.edited };
          info[k] = { name: st.fileName, count: st.items ? st.items.size : 0, edited: !!st.edited };
          if (savedBuf[k] !== st.srcBuf) { ops.push({ store: 'files', value: { fid: cur.id + ':' + k, buf: st.srcBuf } }); newSaved[k] = st.srcBuf; }
        } else {
          stores[k] = null; info[k] = null;
          if (savedBuf[k] !== null) { ops.push({ store: 'files', del: cur.id + ':' + k }); newSaved[k] = null; }
        }
      }
      const now = Date.now();
      ops.push({ store: 'meta', value: { id: cur.id, name: cur.name, created: cur.created, updated: now, info } });
      ops.push({ store: 'state', value: { id: cur.id, settings: A.getSettings(), review: window.__review ? window.__review.getState() : null, stores } });
      await write(ops);
      savedBuf = newSaved;
      try { localStorage.setItem('stockLastProject', cur.id); } catch (_) { /* لا بأس */ }
      setState('✓ تم الحفظ ' + A.timeStr(new Date(now)), 'ok');
    } catch (e) {
      dirty = true;
      const full = e && (e.name === 'QuotaExceededError' || /quota/i.test(String(e.message)));
      setState(full ? '✕ تعذّر الحفظ: مساحة الجهاز ممتلئة. احذف مشاريع قديمة.' : '✕ تعذّر الحفظ', 'err');
    } finally {
      saving = false;
      if (again) { again = false; schedule(300); }
    }
  }

  async function flush() {
    clearTimeout(timer);
    while (saving) await sleep(60);
    if (dirty) await save();
  }

  /* ---------- فتح مشروع ---------- */
  async function openProject(id, silent) {
    const A = App();
    restoring = true;
    setState('جارٍ فتح المشروع…', 'busy');
    try {
      const [meta, st] = await Promise.all([getOne('meta', id), getOne('state', id)]);
      if (!meta || !st) throw new Error('missing');
      A.resetAll();
      cur = { id: meta.id, name: meta.name, created: meta.created };
      savedBuf = { gh: null, hr: null };
      for (const k of KEYS) {
        const snap = st.stores && st.stores[k];
        if (!snap) continue;
        const fr = await getOne('files', id + ':' + k);
        if (!fr) continue;
        await A.restoreStore(k, Object.assign({}, snap, { buf: fr.buf }));
        savedBuf[k] = fr.buf;
      }
      A.setSettings(st.settings);
      if (window.__review) window.__review.setState(st.review);
      if (st.settings && st.settings.compared && A.state.gh.items && A.state.hr.items) await A.runCompare();
      try { localStorage.setItem('stockLastProject', id); } catch (_) { /* لا بأس */ }
      updateBar();
      setState(silent ? '✓ تم فتح آخر مشروع: ' + cur.name : '✓ تم فتح المشروع', 'ok');
      window.scrollTo(0, 0);
      return true;
    } catch (e) {
      setState('✕ تعذّر فتح المشروع', 'err');
      return false;
    } finally {
      restoring = false; dirty = false;
    }
  }

  /* ---------- مشروع جديد / تسمية ---------- */
  async function newProject() {
    const name = prompt('اسم المشروع الجديد:', 'مشروع جديد');
    if (name === null) return;
    await flush();
    restoring = true;
    App().resetAll();
    App().setSettings({ threshold: 5, A: true, B: true, C: true });
    if (window.__review) window.__review.setState({});
    restoring = false;
    savedBuf = { gh: null, hr: null };
    cur = { id: newId(), name: name.trim() || defaultName(), created: Date.now() };
    updateBar();
    dirty = true;
    await save();
    window.scrollTo(0, 0);
  }

  async function renameProject(id, current) {
    const name = prompt('اسم المشروع:', current);
    if (name === null || !name.trim()) return false;
    const meta = await getOne('meta', id);
    if (!meta) return false;
    meta.name = name.trim();
    await write([{ store: 'meta', value: meta }]);
    if (cur && cur.id === id) { cur.name = meta.name; updateBar(); }
    return true;
  }

  async function renameCurrent() {
    if (!db) return;
    if (!cur) {
      const name = prompt('اسم المشروع:', defaultName());
      if (name === null || !name.trim()) return;
      cur = { id: newId(), name: name.trim(), created: Date.now() };
      updateBar(); dirty = true; await save();
      return;
    }
    await flush();
    await renameProject(cur.id, cur.name);
  }

  /* ---------- قائمة المشاريع ---------- */
  function buildList() {
    listView = document.createElement('div');
    listView.id = 'projectsView';
    listView.hidden = true;
    listView.innerHTML = `
      <div class="ed-top"><div class="ttl">📁 المشاريع المحفوظة على هذا الجهاز</div><button class="btn primary" id="pj-close" type="button">✕ إغلاق</button></div>
      <div class="pj-body" id="pj-body"></div>`;
    document.body.appendChild(listView);
    listView.querySelector('#pj-close').addEventListener('click', () => closeList());
    listView.querySelector('#pj-body').addEventListener('click', onListClick);
    window.addEventListener('popstate', () => { if (listView && !listView.hidden) { pushed = false; closeList(true); } });
  }

  async function showList() {
    if (!db) { setState('الحفظ غير متاح في هذا المتصفح.', 'err'); return; }
    if (!listView) buildList();
    await flush();
    await renderList();
    listView.hidden = false;
    document.body.classList.add('noscroll');
    if (!pushed) { try { history.pushState({ pj: 1 }, ''); pushed = true; } catch (_) { /* لا بأس */ } }
  }

  function closeList(fromPop) {
    if (!listView || listView.hidden) return;
    listView.hidden = true;
    document.body.classList.remove('noscroll');
    if (pushed && !fromPop) { pushed = false; try { history.back(); } catch (_) { /* لا بأس */ } }
  }

  async function renderList() {
    const metas = (await getAll('meta')).sort((a, b) => b.updated - a.updated);
    const body = listView.querySelector('#pj-body');
    if (!metas.length) { body.innerHTML = '<p class="empty">لا توجد مشاريع محفوظة بعد. استورد ملف Excel وسيُحفظ مشروعك تلقائيًا.</p>'; return; }
    const line = (label, i) => label + ': ' + (i ? esc(i.name) + ' (' + i.count + ' مادة' + (i.edited ? ' — معدّل' : '') + ')' : 'لا يوجد ملف');
    body.innerHTML = metas.map(m => {
      const isCur = cur && cur.id === m.id;
      return `<div class="pj-item${isCur ? ' cur' : ''}">
        <div class="pj-t">${esc(m.name)}${isCur ? ' <i>● مفتوح الآن</i>' : ''}</div>
        <div class="pj-d">آخر حفظ: ${esc(stamp(new Date(m.updated)))}</div>
        <div class="pj-i">${line('معرض الغدير', m.info && m.info.gh)}</div>
        <div class="pj-i">${line('الحرفيين', m.info && m.info.hr)}</div>
        <div class="pj-b">
          <button type="button" class="btn primary" data-a="open" data-id="${m.id}">فتح</button>
          <button type="button" class="btn" data-a="ren" data-id="${m.id}" data-n="${esc(m.name)}">✎ تغيير الاسم</button>
          <button type="button" class="btn danger" data-a="del" data-id="${m.id}" data-n="${esc(m.name)}">🗑 حذف</button>
        </div></div>`;
    }).join('');
  }

  async function onListClick(e) {
    const b = e.target.closest('[data-a]');
    if (!b) return;
    const id = b.dataset.id;
    if (b.dataset.a === 'open') {
      closeList();
      await flush();
      await openProject(id, false);
    } else if (b.dataset.a === 'ren') {
      if (await renameProject(id, b.dataset.n)) await renderList();
    } else if (b.dataset.a === 'del') {
      if (!confirm('حذف المشروع «' + b.dataset.n + '» نهائيًا من هذا الجهاز؟')) return;
      await write([{ store: 'meta', del: id }, { store: 'state', del: id }, { store: 'files', del: id + ':gh' }, { store: 'files', del: id + ':hr' }]);
      if (cur && cur.id === id) {
        restoring = true; App().resetAll(); restoring = false;
        cur = null; savedBuf = { gh: null, hr: null }; updateBar();
        try { localStorage.removeItem('stockLastProject'); } catch (_) { /* لا بأس */ }
        setState('تم حذف المشروع');
      }
      await renderList();
    }
  }

  /* ---------- تشغيل ---------- */
  async function init() {
    if (!window.__stockApp) return;
    buildBar();
    try { db = await openDB(); }
    catch (_) { setState('الحفظ غير متاح في هذا المتصفح (الوضع الخاص؟).', 'err'); return; }
    try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (_) { /* لا بأس */ }

    for (const ev of ['stores', 'settings', 'results', 'edit']) App().on(ev, () => schedule());
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && dirty) flush(); });
    window.addEventListener('pagehide', () => { if (dirty) flush(); });

    let last = null;
    try { last = localStorage.getItem('stockLastProject'); } catch (_) { /* لا بأس */ }
    if (last) {
      const meta = await getOne('meta', last).catch(() => null);
      if (meta) await openProject(last, true);
    }
  }

  window.__projects = { save, flush, openProject };
  init();
})();
