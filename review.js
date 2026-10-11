/* ============================================================
   ملخص المخزنين + استعراض جميع المواد (مع بحث واستثناء وإخفاء/إظهار الصفري)
   كل شيء يُحسب تلقائيًا من الملفات المستوردة، وبدون إنترنت.
   ============================================================ */
(function () {
  'use strict';

  const App = () => window.__stockApp;
  const PAGE = 200, PDF_CAP = 5000;
  const collator = new Intl.Collator('ar');
  const f = { q: '', x: '', z: '0', s: 'name', tab: 'all' };
  let card = null, el = {}, limit = PAGE, timer = null, memo = new Map(), view = null;

  const TABS = {
    all:   { num: '①', title: 'جميع المواد' },
    onlyA: { num: '②', title: 'موجودة في معرض الغدير وغير موجودة في الحرفيين' },
    onlyB: { num: '③', title: 'موجودة في الحرفيين وغير موجودة في معرض الغدير' },
    both:  { num: '④', title: 'موجودة في المخزنين' }
  };
  const ZERO = { '0': 'إظهار الكل', '1': 'إخفاء المواد الصفرية (صفر في كل مخزن موجودة فيه)', '2': 'إخفاء أي مادة فيها صفر في أي مخزن', '3': 'إظهار المواد الصفرية فقط' };
  const SORT = { name: 'الاسم', gh: 'كمية الغدير (الأكثر أولًا)', hr: 'كمية الحرفيين (الأكثر أولًا)', diff: 'الأكبر فرقًا بين المخزنين' };

  /* ---------- أدوات ---------- */
  const SMAP = { 'أ': 'ا', 'إ': 'ا', 'آ': 'ا', 'ٱ': 'ا', 'ة': 'ه', 'ى': 'ي', 'ئ': 'ي', 'ؤ': 'و' };
  function sNorm(v) {
    return String(v).replace(/[ً-ْٰـ]|[أإآٱةىئؤ]|[٠-٩]|[۰-۹]/g, m => {
      const c = m.charCodeAt(0);
      if (c >= 0x660 && c <= 0x669) return String(c - 0x660);
      if (c >= 0x6F0 && c <= 0x6F9) return String(c - 0x6F0);
      return SMAP[m] || '';
    }).toLowerCase();
  }
  function nameNorm(key, name) {
    let v = memo.get(key);
    if (v === undefined) { v = sNorm(name); if (memo.size > 400000) memo.clear(); memo.set(key, v); }
    return v;
  }
  const esc = s => App().esc(s);
  const fmt = n => App().fmt(n);
  const num = v => (v === null || v === undefined) ? '<span class="na">غير موجودة</span>' : `<span class="${v <= 0 ? 'zq' : ''}">${fmt(v)}</span>`;
  const numTxt = v => (v === null || v === undefined) ? 'غير موجودة' : fmt(v);

  /* ---------- الحساب ---------- */
  function totals() {
    const S = App().state, g = S.gh.items, h = S.hr.items;
    if (!g || !h) return null;
    let onlyA = 0, onlyB = 0, both = 0;
    for (const k of g.keys()) { if (h.has(k)) both++; else onlyA++; }
    for (const k of h.keys()) if (!g.has(k)) onlyB++;
    return { gh: g.size, hr: h.size, onlyA, onlyB, both };
  }

  function storeStats(items) {
    if (!items) return null;
    let pos = 0, zero = 0;
    for (const v of items.values()) { if (v.qty > 0) pos++; else zero++; }
    return { n: items.size, pos, zero };
  }

  function passZero(r) {
    if (f.z === '0') return true;
    const a = r.gh, b = r.hr;
    if (f.z === '1') return (a !== null && a > 0) || (b !== null && b > 0);
    if (f.z === '2') return (a === null || a > 0) && (b === null || b > 0);
    return (a !== null && a <= 0) || (b !== null && b <= 0);
  }

  function compute() {
    const S = App().state, g = S.gh.items, h = S.hr.items;
    if (!g && !h) return null;
    const inc = sNorm(f.q).split(/\s+/).filter(Boolean);
    const exc = sNorm(f.x).split(/[\s,،؛;]+/).filter(Boolean);
    const keys = new Set();
    if (g) for (const k of g.keys()) keys.add(k);
    if (h) for (const k of h.keys()) keys.add(k);

    const rows = [];
    let mg = 0, mh = 0;
    for (const k of keys) {
      const a = g ? g.get(k) : undefined, b = h ? h.get(k) : undefined;
      const name = (a || b).name;
      if (inc.length || exc.length) {
        const n = nameNorm(k, name);
        let ok = true;
        for (const w of inc) if (!n.includes(w)) { ok = false; break; }
        if (ok) for (const w of exc) if (n.includes(w)) { ok = false; break; }
        if (!ok) continue;
      }
      if (a) mg++;
      if (b) mh++;
      rows.push({ name, gh: a ? a.qty : null, hr: b ? b.qty : null });
    }

    const lists = { all: [], onlyA: [], onlyB: [], both: [] };
    for (const r of rows) {
      if (!passZero(r)) continue;
      lists.all.push(r);
      if (r.gh !== null && r.hr === null) lists.onlyA.push(r);
      else if (r.hr !== null && r.gh === null) lists.onlyB.push(r);
      else lists.both.push(r);
    }
    const cmpName = (x, y) => collator.compare(x.name, y.name);
    let cmp = cmpName;
    if (f.s === 'gh') cmp = (x, y) => ((y.gh === null ? -1 : y.gh) - (x.gh === null ? -1 : x.gh)) || cmpName(x, y);
    else if (f.s === 'hr') cmp = (x, y) => ((y.hr === null ? -1 : y.hr) - (x.hr === null ? -1 : x.hr)) || cmpName(x, y);
    else if (f.s === 'diff') cmp = (x, y) => (Math.abs((y.gh || 0) - (y.hr || 0)) - Math.abs((x.gh || 0) - (x.hr || 0))) || cmpName(x, y);
    for (const k of Object.keys(lists)) lists[k].sort(cmp);

    return { lists, mg, mh, filtered: inc.length > 0 || exc.length > 0, g, h };
  }

  function sums(rows) {
    let a = 0, b = 0;
    for (const r of rows) { if (r.gh !== null) a += r.gh; if (r.hr !== null) b += r.hr; }
    return { a, b };
  }

  /* ---------- الرسم ---------- */
  function columns(tab) {
    if (tab === 'onlyA') return ['gh'];
    if (tab === 'onlyB') return ['hr'];
    return ['gh', 'hr'];
  }
  const HEAD = { gh: 'معرض الغدير', hr: 'الحرفيين' };

  function tableHtml(tab, rows, lim, forReport) {
    if (!rows.length) return '<p class="empty">لا توجد مواد في هذه القائمة.</p>';
    const cols = columns(tab);
    const shown = lim >= rows.length ? rows : rows.slice(0, lim);
    let h = '<div class="table-wrap"><table class="tbl"><thead><tr><th class="n">#</th><th>المادة</th>' +
      cols.map(c => `<th class="n">${HEAD[c]}</th>`).join('') + '</tr></thead><tbody>';
    shown.forEach((r, i) => {
      h += `<tr><td class="idx n">${i + 1}</td><td class="name">${esc(r.name)}</td>` + cols.map(c => `<td class="n">${num(r[c])}</td>`).join('') + '</tr>';
    });
    const t = sums(rows);
    h += '</tbody><tfoot><tr class="tot"><td></td><td>إجمالي الكميات (' + rows.length + ' مادة)</td>' +
      cols.map(c => `<td class="n">${fmt(c === 'gh' ? t.a : t.b)}</td>`).join('') + '</tr></tfoot></table></div>';
    return h;
  }

  function render() {
    const S = App().state;
    const g = S.gh.items, h = S.hr.items;
    card.hidden = !(g || h);
    if (card.hidden) { view = null; return; }
    view = compute();
    const both = !!(g && h);
    const gs = storeStats(g), hs = storeStats(h);
    const t = totals();

    const storeCard = (name, st, matched) => st
      ? `<div class="sumcard"><h3>${name}</h3><div class="big">${st.n}</div><div class="sub">مادة</div>
          <div class="ln">بكمية أكبر من صفر: <b>${st.pos}</b></div><div class="ln">كميتها صفر: <b>${st.zero}</b></div>
          ${view.filtered ? `<div class="ln hl">المطابقة للبحث: <b>${matched}</b></div>` : ''}</div>`
      : `<div class="sumcard off"><h3>${name}</h3><div class="sub">لم يُستورد بعد</div></div>`;
    el.sum.innerHTML = storeCard('معرض الغدير', gs, view.mg) + storeCard('الحرفيين', hs, view.mh);

    el.cross.innerHTML = t
      ? `موجودة في المخزنين: <b>${t.both}</b> · في معرض الغدير فقط: <b>${t.onlyA}</b> · في الحرفيين فقط: <b>${t.onlyB}</b>`
      : 'استورد ملف المخزنين معًا لتظهر المواد الموجودة في مخزن وغير الموجودة في الآخر.';

    if (!both && f.tab !== 'all') f.tab = 'all';
    el.tabs.innerHTML = Object.keys(TABS).map(id => {
      const dis = !both && id !== 'all';
      const c = TABS[id];
      return `<button type="button" class="cat${f.tab === id ? ' active' : ''}${dis ? ' off' : ''}" data-rv="${id}"${dis ? ' disabled' : ''}>
        <span class="cat-title">${c.num} ${c.title}</span>
        <span class="cat-count">${dis ? '—' : 'عدد النتائج: ' + view.lists[id].length + ' مادة'}</span></button>`;
    }).join('');

    const rows = view.lists[f.tab];
    let html = `<h3 class="table-title">${TABS[f.tab].title} — عدد النتائج: ${rows.length} مادة</h3>` + tableHtml(f.tab, rows, limit);
    if (rows.length > limit) html += `<button type="button" class="btn more" id="rv-more">عرض المزيد (المتبقي ${rows.length - limit})</button>`;
    el.table.innerHTML = html;
  }

  const schedule = (ms) => { clearTimeout(timer); timer = setTimeout(render, ms === undefined ? 120 : ms); };

  /* ---------- تقرير PDF للقائمة ---------- */
  function reportHtml() {
    if (!view) return '';
    const now = new Date(), A = App();
    const rows = view.lists[f.tab];
    const t = totals();
    let h = `<h1>تقرير قائمة المواد</h1><p class="rdate">تاريخ التقرير: ${A.dateStr(now)} — ${A.timeStr(now)}</p>
      <table class="rmeta"><tbody>
      <tr><th>القائمة</th><td>${esc(TABS[f.tab].title)}</td></tr>
      <tr><th>عدد مواد معرض الغدير</th><td>${view.g ? view.g.size : 'لم يُستورد'}</td></tr>
      <tr><th>عدد مواد الحرفيين</th><td>${view.h ? view.h.size : 'لم يُستورد'}</td></tr>`;
    if (t) h += `<tr><th>موجودة في الغدير وغير موجودة في الحرفيين</th><td>${t.onlyA}</td></tr><tr><th>موجودة في الحرفيين وغير موجودة في الغدير</th><td>${t.onlyB}</td></tr>`;
    if (f.q.trim()) h += `<tr><th>البحث</th><td>${esc(f.q.trim())}</td></tr>`;
    if (f.x.trim()) h += `<tr><th>الاستثناء</th><td>${esc(f.x.trim())}</td></tr>`;
    h += `<tr><th>المواد الصفرية</th><td>${esc(ZERO[f.z])}</td></tr><tr><th>عدد النتائج</th><td>${rows.length} مادة</td></tr></tbody></table>`;
    h += tableHtml(f.tab, rows, PDF_CAP, true);
    if (rows.length > PDF_CAP) h += `<p class="foot">تم عرض أول ${PDF_CAP} مادة فقط من أصل ${rows.length}. ضيّق البحث لعرض الباقي.</p>`;
    return h;
  }

  /* ---------- البناء ---------- */
  function build() {
    card = document.createElement('section');
    card.className = 'card';
    card.id = 'review';
    card.hidden = true;
    card.innerHTML = `
      <h2>ملخص المخزنين</h2>
      <div class="sumgrid" id="rv-sum"></div>
      <p class="cross" id="rv-cross"></p>
      <h2 class="rv-h2">استعراض جميع المواد</h2>
      <label class="field"><span>🔍 ابحث (تظهر المواد التي فيها كل هذه الكلمات)</span>
        <input id="rv-q" type="search" placeholder="مثلاً: كاوه  أو  النور" autocomplete="off"></label>
      <label class="field"><span>🚫 استثناء (تختفي المواد التي فيها أي من هذه الكلمات)</span>
        <input id="rv-x" type="search" placeholder="مثلاً: كاوه" autocomplete="off"></label>
      <div class="two">
        <label class="field"><span>المواد الصفرية</span><select id="rv-z">${Object.keys(ZERO).map(k => `<option value="${k}">${ZERO[k]}</option>`).join('')}</select></label>
        <label class="field"><span>الترتيب</span><select id="rv-s">${Object.keys(SORT).map(k => `<option value="${k}">${SORT[k]}</option>`).join('')}</select></label>
      </div>
      <div class="cats" id="rv-tabs"></div>
      <div id="rv-table"></div>
      <button type="button" class="btn big" id="rv-pdf">📄 تقرير هذه القائمة (PDF)</button>`;
    const ref = document.getElementById('settings');
    ref.parentNode.insertBefore(card, ref);
    const q = s => card.querySelector(s);
    el = { sum: q('#rv-sum'), cross: q('#rv-cross'), tabs: q('#rv-tabs'), table: q('#rv-table'), q: q('#rv-q'), x: q('#rv-x'), z: q('#rv-z'), s: q('#rv-s') };

    const changed = () => { f.q = el.q.value; f.x = el.x.value; f.z = el.z.value; f.s = el.s.value; limit = PAGE; schedule(250); App().emit('settings'); };
    el.q.addEventListener('input', changed);
    el.x.addEventListener('input', changed);
    el.z.addEventListener('change', changed);
    el.s.addEventListener('change', changed);
    el.tabs.addEventListener('click', e => {
      const b = e.target.closest('[data-rv]');
      if (!b || b.disabled) return;
      f.tab = b.dataset.rv; limit = PAGE; render(); App().emit('settings');
    });
    el.table.addEventListener('click', e => { if (e.target.id === 'rv-more') { limit += PAGE; render(); } });
    q('#rv-pdf').addEventListener('click', () => { if (view) App().showReport(reportHtml()); });
  }

  function getState() { return { q: f.q, x: f.x, z: f.z, s: f.s, tab: f.tab }; }
  function setState(s) {
    if (!s) return;
    f.q = s.q || ''; f.x = s.x || ''; f.z = s.z || '0'; f.s = s.s || 'name'; f.tab = s.tab || 'all';
    el.q.value = f.q; el.x.value = f.x; el.z.value = f.z; el.s.value = f.s;
    limit = PAGE;
    render();
  }

  function init() {
    if (!window.__stockApp) return;
    build();
    App().on('stores', () => schedule(60));
    window.__review = { totals, getState, setState, refresh: render };
    render();
  }
  init();
})();
