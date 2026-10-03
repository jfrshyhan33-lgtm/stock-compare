/* ============================================================
   مقارنة المخزون — كل المعالجة تتم على الجهاز، لا شيء يُرفع للإنترنت
   الأقسام:
   1) أدوات صغيرة
   2) قراءة Excel (xlsx) و CSV
   3) تنظيف الأسماء وقراءة الكميات
   4) حالة التطبيق وواجهة المخزنين
   5) المقارنة
   6) عرض النتائج
   7) التقرير و PDF
   ============================================================ */
(function () {
  'use strict';

  /* ---------- 1) أدوات صغيرة ---------- */
  const $ = (s, r = document) => r.querySelector(s);
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const tick = () => new Promise(r => setTimeout(r, 0));
  const fmt = n => String(Math.round(n * 1000) / 1000);
  const collator = new Intl.Collator('ar');

  class AppError extends Error {}

  function colLetter(i) {
    let s = '', n = i + 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function colIndex(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
    return n - 1;
  }

  /* ---------- 2) قراءة Excel (xlsx) ---------- */
  function unxml(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e) => {
      switch (e) {
        case 'amp': return '&';
        case 'lt': return '<';
        case 'gt': return '>';
        case 'quot': return '"';
        case 'apos': return "'";
      }
      const code = (e[1] === 'x' || e[1] === 'X') ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(code); } catch (_) { return ''; }
    });
  }
  function decodeText(s) {
    return unxml(s).replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
  }
  function attr(tag, name) {
    const m = new RegExp('(?:^|\\s)' + name + '="([^"]*)"').exec(tag);
    return m ? m[1] : null;
  }
  function collectT(s) {
    let out = '';
    for (const m of s.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) out += m[1];
    return decodeText(out);
  }
  async function zipText(zip, path) {
    let f = zip.file(path);
    if (!f) {
      const low = path.toLowerCase();
      const alt = Object.keys(zip.files).find(k => k.toLowerCase() === low);
      if (!alt) return null;
      f = zip.file(alt);
    }
    return f.async('string');
  }

  async function loadShared(zip) {
    const xml = await zipText(zip, 'xl/sharedStrings.xml');
    if (xml == null) return [];
    const out = [];
    const re = /<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g;
    let m, n = 0;
    while ((m = re.exec(xml)) !== null) {
      const body = (m[1] || '').replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
      out.push(collectT(body));
      if (++n % 20000 === 0) await tick();
    }
    return out;
  }

  async function parseSheet(xml, shared) {
    const rows = [];
    const rowRe = /<row\b[^>]*?(?:\/>|>([\s\S]*?)<\/row>)/g;
    let m, n = 0, seqRow = 0;
    while ((m = rowRe.exec(xml)) !== null) {
      const inner = m[1];
      const rp = /^<row\b[^>]*?\sr="(\d+)"/.exec(m[0]);
      const rowPos = rp ? parseInt(rp[1], 10) - 1 : seqRow;
      seqRow = rowPos + 1;
      if (inner && rowPos >= 0 && rowPos < 1048576) {
        const arr = [];
        let any = false, seq = 0, c;
        const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
        while ((c = cellRe.exec(inner)) !== null) {
          const attrs = c[1], body = c[2];
          const rm = /\sr="([A-Z]+)\d*"/.exec(attrs);
          const ci = rm ? colIndex(rm[1]) : seq;
          seq = ci + 1;
          if (!body) continue;
          const tm = /\st="([^"]*)"/.exec(attrs);
          const t = tm ? tm[1] : 'n';
          let val;
          if (t === 'inlineStr') {
            const im = /<is>([\s\S]*?)<\/is>/.exec(body);
            val = im ? collectT(im[1]) : '';
          } else {
            const vm = /<v>([\s\S]*?)<\/v>/.exec(body);
            if (!vm) continue;
            const raw = vm[1];
            if (t === 's') val = shared[parseInt(raw, 10)];
            else if (t === 'str') val = decodeText(raw);
            else if (t === 'e') continue;
            else if (t === 'b') val = raw === '1' ? 'TRUE' : 'FALSE';
            else { val = Number(raw); if (!isFinite(val)) val = decodeText(raw); }
          }
          if (val === undefined || val === '') continue;
          arr[ci] = val;
          any = true;
        }
        if (any) rows[rowPos] = arr;
      }
      if (++n % 4000 === 0) await tick();
    }
    return rows;
  }

  function parseMerges(xml) {
    const out = [];
    const re = /<mergeCell\b[^>]*?\sref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g;
    let m;
    while ((m = re.exec(xml)) !== null && out.length < 200000) {
      const c1 = colIndex(m[1]), r1 = parseInt(m[2], 10) - 1, c2 = colIndex(m[3]), r2 = parseInt(m[4], 10) - 1;
      if (r1 >= 0 && r2 >= r1 && c2 >= c1 && (r2 > r1 || c2 > c1)) out.push({ r1, c1, r2, c2 });
    }
    return out;
  }

  async function openXlsx(buf) {
    if (typeof JSZip === 'undefined') throw new AppError('تعذّر تحميل مكتبة قراءة Excel. أعد فتح التطبيق.');
    let zip;
    try { zip = await JSZip.loadAsync(buf); }
    catch (_) { throw new AppError('الملف غير صالح أو تالف. اختر ملف Excel بصيغة xlsx.'); }

    const wbXml = await zipText(zip, 'xl/workbook.xml');
    if (wbXml == null) throw new AppError('الملف ليس ملف Excel صالحًا (xlsx).');
    const relsXml = (await zipText(zip, 'xl/_rels/workbook.xml.rels')) || '';

    const rels = {};
    for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = attr(m[0], 'Id'), tg = attr(m[0], 'Target');
      if (id && tg) rels[id] = tg;
    }
    const sheets = [];
    for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
      const name = decodeText(attr(m[0], 'name') || '');
      const tg = rels[attr(m[0], 'r:id')];
      if (!tg) continue;
      const path = tg.startsWith('/') ? tg.slice(1) : 'xl/' + tg.replace(/^\.?\//, '');
      sheets.push({ name: name || ('Sheet' + (sheets.length + 1)), path });
    }
    if (!sheets.length) throw new AppError('لم يتم العثور على أوراق عمل داخل الملف.');

    let shared = null;
    return {
      names: sheets.map(s => s.name),
      async getRows(i) {
        if (shared === null) shared = await loadShared(zip);
        const xml = await zipText(zip, sheets[i].path);
        if (xml == null) throw new AppError('تعذّر قراءة الورقة المحددة.');
        const rows = await parseSheet(xml, shared);
        return { rows, merges: parseMerges(xml) };
      }
    };
  }

  function openCsv(buf) {
    let text = new TextDecoder('utf-8').decode(buf);
    if (text.includes('�')) {
      try { text = new TextDecoder('windows-1256').decode(buf); } catch (_) { /* نبقي utf-8 */ }
    }
    text = text.replace(/^﻿/, '');
    const first = text.split(/\r?\n/, 1)[0] || '';
    const counts = { ',': 0, ';': 0, '\t': 0 };
    for (const ch of first) if (ch in counts) counts[ch]++;
    let delim = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
    if (!counts[delim]) delim = ',';

    const rows = [];
    let row = [], cell = '', q = false, any = false;
    const endCell = () => { row.push(cell); if (cell.trim() !== '') any = true; cell = ''; };
    const endRow = () => {
      endCell();
      if (any) {
        const arr = [];
        row.forEach((v, i) => { if (v.trim() !== '') arr[i] = v.trim(); });
        rows.push(arr);
      }
      row = []; any = false;
    };
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === delim) endCell();
      else if (ch === '\n') endRow();
      else if (ch !== '\r') cell += ch;
    }
    if (cell !== '' || row.length) endRow();
    return { names: ['CSV'], getRows: async () => ({ rows, merges: [] }) };
  }

  async function openWorkbook(file) {
    const buf = await file.arrayBuffer();
    if (!buf.byteLength) throw new AppError('الملف فارغ.');
    const h = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
    if (h[0] === 0x50 && h[1] === 0x4B) return openXlsx(buf);               // ZIP => xlsx
    if (h[0] === 0xD0 && h[1] === 0xCF && h[2] === 0x11 && h[3] === 0xE0)  // OLE => xls قديم أو محمي
      throw new AppError('هذا الملف بصيغة xls القديمة أو محمي بكلمة مرور. افتحه في Excel واحفظه بصيغة xlsx (عادي) ثم استورده.');
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if (['csv', 'tsv', 'txt'].includes(ext)) return openCsv(buf);
    throw new AppError('الملف غير صالح. اختر ملف Excel بصيغة xlsx.');
  }

  /* ---------- 3) تنظيف الأسماء وقراءة الكميات ---------- */
  // تنظيف محافظ: مسافات زائدة، تشكيل، تطويل، أرقام عربية، حروف لاتينية صغيرة.
  // لا نوحّد الهمزات أو التاء المربوطة حتى لا ندمج أسماء مختلفة عشوائيًا.
  function normName(s) {
    return String(s)
      .normalize('NFKC')
      .replace(/[ً-ٰٟـ]/g, '')
      .replace(/[​-‏‪-‮⁦-⁩﻿]/g, '')
      .replace(/[٠-٩]/g, d => String(d.charCodeAt(0) - 0x660))
      .replace(/[۰-۹]/g, d => String(d.charCodeAt(0) - 0x6F0))
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }
  function cleanDisplay(s) { return String(s).replace(/[​-‏‪-‮⁦-⁩﻿]/g, '').replace(/\s+/g, ' ').trim(); }

  function parseQty(v) {
    if (v === undefined || v === null) return { ok: true, n: 0 };
    if (typeof v === 'number') return isFinite(v) ? { ok: true, n: Math.round(v * 1e6) / 1e6 } : { ok: false };
    let s = String(v).trim();
    if (s === '') return { ok: true, n: 0 };
    s = s.replace(/[٠-٩]/g, d => String(d.charCodeAt(0) - 0x660))
         .replace(/[۰-۹]/g, d => String(d.charCodeAt(0) - 0x6F0))
         .replace(/٫/g, '.').replace(/[٬,\s]/g, '');
    if (/^[-+]?\d+(\.\d+)?$/.test(s)) return { ok: true, n: Math.round(Number(s) * 1e6) / 1e6 };
    return { ok: false };
  }
  const isNumericCell = v => typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && parseQty(v).ok);

  /* ---------- 4) حالة التطبيق وواجهة المخزنين ---------- */
  const STORE_DEFS = [
    { key: 'gh', ordinal: 'المخزن الأول', name: 'معرض الغدير' },
    { key: 'hr', ordinal: 'المخزن الثاني', name: 'الحرفيين' }
  ];
  const CATS = {
    A: { short: 'معرض الغدير قليل ← الحرفيين متوفر', title: 'مواد قليلة في معرض الغدير ومتوفرة في الحرفيين', num: '①', diff: true },
    B: { short: 'الحرفيين قليل ← معرض الغدير متوفر', title: 'مواد قليلة في الحرفيين ومتوفرة في معرض الغدير', num: '②', diff: true },
    C: { short: 'غير متوفر في المخزنين', title: 'مواد غير متوفرة في المخزنين', num: '③', diff: false }
  };
  const PAGE = 200;

  function newStore(def) {
    return { def, fileName: '', wb: null, sheetIdx: 0, rows: null, headerRow: 0, cols: [], nameCol: '', qtyCol: '', items: null, skipped: 0, file: null, merges: [], edited: false };
  }
  const state = {
    gh: newStore(STORE_DEFS[0]),
    hr: newStore(STORE_DEFS[1]),
    results: null,
    active: 'A',
    limit: { A: PAGE, B: PAGE, C: PAGE }
  };

  function renderStoreCards() {
    $('#stores').innerHTML = STORE_DEFS.map(d => `
      <section class="card" id="card-${d.key}">
        <div class="card-head"><span class="ordinal">${d.ordinal}</span><h2>${d.name}</h2></div>
        <input type="file" id="${d.key}-file" hidden
               accept=".xlsx,.xlsm,.xls,.csv,.tsv,.txt,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv">
        <label class="btn primary big" for="${d.key}-file">📁 استيراد مخزن ${d.name}</label>
        <p class="status" id="${d.key}-status">لم يتم اختيار ملف بعد.</p>
        <button type="button" class="btn" id="${d.key}-edit" hidden>✏️ عرض وتعديل الملف (حذف وتحديد ودمج)</button>
        <div class="cols" id="${d.key}-cols" hidden>
          <label class="field" id="${d.key}-sheetWrap" hidden><span>الورقة (Sheet)</span><select id="${d.key}-sheet"></select></label>
          <label class="field"><span>عمود اسم المادة</span><select id="${d.key}-name"></select></label>
          <label class="field"><span>عمود الكمية</span><select id="${d.key}-qty"></select></label>
          <p class="sample" id="${d.key}-sample"></p>
        </div>
      </section>`).join('');

    for (const d of STORE_DEFS) {
      const k = d.key;
      $('#' + k + '-file').addEventListener('change', e => {
        const f = e.target.files && e.target.files[0];
        e.target.value = '';           // للسماح باختيار نفس الملف مرة أخرى
        if (f) loadFile(k, f);
      });
      $('#' + k + '-edit').addEventListener('click', () => { if (window.openEditor) window.openEditor(k); });
      $('#' + k + '-sheet').addEventListener('change', e => changeSheet(k, parseInt(e.target.value, 10)));
      $('#' + k + '-name').addEventListener('change', e => { state[k].nameCol = e.target.value; recompute(k); });
      $('#' + k + '-qty').addEventListener('change', e => { state[k].qtyCol = e.target.value; recompute(k); });
    }
  }

  function setStatus(key, html, cls) {
    const el = $('#' + key + '-status');
    el.className = 'status' + (cls ? ' ' + cls : '');
    el.innerHTML = html;
  }
  function hideCols(key) { $('#' + key + '-cols').hidden = true; $('#' + key + '-edit').hidden = true; }

  function errText(err) {
    return err instanceof AppError ? err.message : 'تعذّرت قراءة الملف. تأكد أنه ملف Excel صالح بصيغة xlsx.';
  }

  async function loadFile(key, file) {
    const st = state[key];
    Object.assign(st, { fileName: file.name, file, merges: [], edited: false, wb: null, rows: null, items: null, nameCol: '', qtyCol: '', skipped: 0 });
    hideCols(key);
    setStatus(key, 'جارٍ قراءة الملف…', 'busy');
    await tick();
    try {
      st.wb = await openWorkbook(file);
      st.sheetIdx = 0;
      await selectSheet(key, 0);
    } catch (err) {
      st.wb = null; st.rows = null; st.items = null;
      setStatus(key, '✕ ' + esc(errText(err)), 'err');
      hideCols(key);
      afterStoreChange();
    }
  }

  async function changeSheet(key, idx) {
    const st = state[key];
    setStatus(key, 'جارٍ قراءة الورقة…', 'busy');
    await tick();
    try { await selectSheet(key, idx); }
    catch (err) {
      st.rows = null; st.items = null;
      setStatus(key, '✕ ' + esc(errText(err)), 'err');
      afterStoreChange();
    }
  }

  async function selectSheet(key, idx) {
    const st = state[key];
    const res = await st.wb.getRows(idx);
    const rows = res.rows;
    st.merges = res.merges || [];
    st.edited = false;
    st.sheetIdx = idx;
    if (!rows.length) { st.rows = null; throw new AppError('الورقة المحددة فارغة. اختر ورقة أخرى.'); }
    st.rows = rows;
    setupColumns(key);
    recompute(key);
  }

  const countFilled = row => (row ? row.filter(v => v !== undefined && v !== '').length : 0);

  function computeCols(st) {
    const rows = st.rows;
    const upto = Math.min(rows.length, 200);
    let nc = 0;
    for (let r = 0; r < upto; r++) nc = Math.max(nc, (rows[r] || []).length);
    st.cols = [];
    for (let i = 0; i < nc; i++) {
      let used = false;
      for (let r = 0; r < upto && !used; r++) { const v = rows[r] && rows[r][i]; if (v !== undefined && v !== '') used = true; }
      if (used) st.cols.push(i);
    }
  }

  function setupColumns(key) {
    const st = state[key], rows = st.rows;

    let h = 0;
    for (let r = 0; r < Math.min(rows.length, 30); r++) if (countFilled(rows[r]) >= 2) { h = r; break; }
    st.headerRow = h;

    computeCols(st);

    const g = guessColumns(st);
    st.nameCol = g.name >= 0 ? String(g.name) : '';
    st.qtyCol = g.qty >= 0 ? String(g.qty) : '';

    renderColumnSelects(key);
  }

  function renderColumnSelects(key) {
    const st = state[key], rows = st.rows;
    const header = rows[st.headerRow] || [];
    const label = i => {
      const t = header[i] === undefined ? 'بدون عنوان' : cleanDisplay(header[i]);
      return colLetter(i) + ' — ' + (t.length > 32 ? t.slice(0, 32) + '…' : t);
    };
    const opts = '<option value="">— اختر العمود —</option>' +
      st.cols.map(i => `<option value="${i}">${esc(label(i))}</option>`).join('');
    const selName = $('#' + key + '-name'), selQty = $('#' + key + '-qty');
    selName.innerHTML = opts; selQty.innerHTML = opts;
    selName.value = st.nameCol; selQty.value = st.qtyCol;

    const sheetSel = $('#' + key + '-sheet');
    const names = st.wb.names;
    $('#' + key + '-sheetWrap').hidden = names.length < 2;
    sheetSel.innerHTML = names.map((n, i) => `<option value="${i}">${esc(n)}</option>`).join('');
    sheetSel.value = String(st.sheetIdx);

    $('#' + key + '-cols').hidden = false;
    $('#' + key + '-edit').hidden = false;
  }

  function guessColumns(st) {
    const h = st.rows[st.headerRow] || [];
    const sample = st.rows.slice(st.headerRow + 1, st.headerRow + 201).filter(Boolean);
    const hdr = i => String(h[i] === undefined ? '' : h[i]).toLowerCase();

    let name = -1, best = 0;
    for (const i of st.cols) {
      const t = hdr(i);
      let s = 0;
      if (/اسم|name|بيان|descr|تفاصيل|منتج|product/.test(t)) s = 3;
      else if (/صنف|item|مادة|ماده|material|بضاعة|goods|article/.test(t)) s = 2;
      if (s && /رقم|كود|code|barcode|باركود|\bno\.?\b|#|\bid\b|serial|تسلسل/.test(t)) s -= 2.5;
      if (s > best) { best = s; name = i; }
    }
    if (name < 0) {
      let top = 0;
      for (const i of st.cols) {
        let c = 0;
        for (const r of sample) { const v = r[i]; if (typeof v === 'string' && v.trim().length > 1 && !parseQty(v).ok) c++; }
        if (c > top) { top = c; name = i; }
      }
    }

    let qty = -1;
    for (const re of [/كمي|رصيد|موجود|مخزون|qty|quant|balance|stock|on.?hand|pcs/, /عدد|count|total|مجموع/]) {
      for (const i of st.cols) { if (i !== name && re.test(hdr(i))) { qty = i; break; } }
      if (qty >= 0) break;
    }
    if (qty < 0) {
      let top = 0;
      for (const i of st.cols) {
        if (i === name) continue;
        let c = 0;
        for (const r of sample) if (isNumericCell(r[i])) c++;
        if (c > top) { top = c; qty = i; }
      }
    }
    return { name, qty };
  }

  function buildItems(st) {
    const map = new Map();
    const ni = parseInt(st.nameCol, 10), qi = parseInt(st.qtyCol, 10);
    let skipped = 0;
    for (let r = st.headerRow + 1; r < st.rows.length; r++) {
      const row = st.rows[r];
      if (!row) continue;
      const raw = row[ni];
      if (raw === undefined || raw === null) continue;
      const disp = cleanDisplay(raw);
      if (!disp) continue;
      const q = parseQty(row[qi]);
      if (!q.ok) { skipped++; continue; }
      const key = normName(disp);
      const e = map.get(key);
      if (e) e.qty += q.n; else map.set(key, { name: disp, qty: q.n });
    }
    st.skipped = skipped;
    return map;
  }

  function recompute(key) {
    const st = state[key], d = st.def;
    st.items = null;
    const sample = $('#' + key + '-sample');
    sample.hidden = true;
    if (!st.rows) { afterStoreChange(); return; }

    if (st.nameCol === '') { setStatus(key, 'يرجى تحديد عمود اسم المادة.', 'warn'); afterStoreChange(); return; }
    if (st.qtyCol === '') { setStatus(key, 'يرجى تحديد عمود الكمية.', 'warn'); afterStoreChange(); return; }
    if (st.nameCol === st.qtyCol) { setStatus(key, 'عمود الاسم وعمود الكمية متطابقان. اختر عمودين مختلفين.', 'warn'); afterStoreChange(); return; }

    st.items = buildItems(st);
    if (st.items.size === 0) {
      setStatus(key, '✕ لم يتم العثور على مواد صالحة. تأكد من اختيار العمودين الصحيحين.', 'err');
      afterStoreChange();
      return;
    }
    let html = `✓ تم تحميل ملف ${d.name}<span class="file">${esc(st.fileName)} — ${st.items.size} مادة</span>`;
    if (st.skipped) html += `<span class="file">تم تجاهل ${st.skipped} صف لأن الكمية فيه ليست رقمًا.</span>`;
    if (st.edited) html += '<span class="file">✎ تم تعديل الملف داخل التطبيق (ملفك الأصلي لم يتغير)</span>';
    setStatus(key, html, 'ok');

    const first = st.items.values().next().value;
    sample.textContent = `تأكد من الأعمدة — مثال: «${first.name}» ← ${fmt(first.qty)}`;
    sample.hidden = false;
    afterStoreChange();
  }

  function afterStoreChange() {
    if (!state.results) return;
    if (state.gh.items && state.hr.items) runCompare(true);
    else hideResults();
  }

  /* ---------- 5) المقارنة ---------- */
  const optOn = id => $('#opt-' + id).checked;

  function readThreshold() {
    const v = $('#threshold').value.trim();
    if (v === '') return null;
    const n = Number(v.replace(',', '.'));
    return isFinite(n) && n >= 0 ? n : null;
  }

  function validate() {
    for (const d of STORE_DEFS) {
      const st = state[d.key];
      if (!st.rows) return `يرجى اختيار ملف Excel لمخزن ${d.name}.`;
      if (st.nameCol === '') return `يرجى تحديد عمود اسم المادة لمخزن ${d.name}.`;
      if (st.qtyCol === '') return `يرجى تحديد عمود الكمية لمخزن ${d.name}.`;
      if (st.nameCol === st.qtyCol) return `عمود اسم المادة وعمود الكمية متطابقان في مخزن ${d.name}. اختر عمودين مختلفين.`;
      if (!st.items || st.items.size === 0) return `لم يتم العثور على مواد صالحة في ملف ${d.name}. تأكد من اختيار الأعمدة الصحيحة.`;
    }
    if (readThreshold() === null) return 'يرجى إدخال رقم صحيح للكمية القليلة (0 أو أكثر).';
    if (!optOn('A') && !optOn('B') && !optOn('C')) return 'يرجى اختيار نوع واحد على الأقل من النتائج.';
    return null;
  }

  function compare() {
    const T = readThreshold();
    const opts = { A: optOn('A'), B: optOn('B'), C: optOn('C') };
    const g = state.gh.items, h = state.hr.items;
    const A = [], B = [], C = [];

    const keys = new Set(g.keys());
    for (const k of h.keys()) keys.add(k);

    for (const k of keys) {
      const a = g.get(k), b = h.get(k);
      const gq = a ? a.qty : 0, hq = b ? b.qty : 0;
      const row = { name: (a || b).name, gh: a ? a.qty : null, hr: b ? b.qty : null };
      if (gq <= 0 && hq <= 0) { if (opts.C) C.push(row); continue; }
      if (gq <= T && hq > T) { if (opts.A) { row.diff = hq - gq; A.push(row); } }
      else if (hq <= T && gq > T) { if (opts.B) { row.diff = gq - hq; B.push(row); } }
    }
    const byDiff = (x, y) => (y.diff - x.diff) || collator.compare(x.name, y.name);
    A.sort(byDiff); B.sort(byDiff);
    C.sort((x, y) => collator.compare(x.name, y.name));
    return { T, opts, A, B, C, date: new Date() };
  }

  function showMsg(text, info) {
    const el = $('#msg');
    el.textContent = text;
    el.className = 'msg' + (info ? ' info' : '');
    el.hidden = false;
  }
  function hideMsg() { $('#msg').hidden = true; }

  async function runCompare(auto) {
    const err = validate();
    if (err) { showMsg(err); if (auto) hideResults(); return; }
    hideMsg();
    const btn = $('#compareBtn');
    btn.disabled = true; btn.textContent = '⏳ جارٍ المقارنة…';
    await tick();
    try {
      state.results = compare();
      const firstEnabled = ['A', 'B', 'C'].find(c => state.results.opts[c]);
      if (!state.results.opts[state.active]) state.active = firstEnabled;
      if (!auto) state.limit = { A: PAGE, B: PAGE, C: PAGE };
      renderResults();
      if (!auto) $('#results').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
      showMsg('حدث خطأ أثناء المقارنة. أعد المحاولة.');
      hideResults();
    } finally {
      btn.disabled = false; btn.textContent = '🔍 إجراء المقارنة';
    }
  }

  function hideResults() {
    state.results = null;
    $('#results').hidden = true;
    $('#reportView').hidden = true;
    document.body.classList.remove('noscroll');
  }

  /* ---------- 6) عرض النتائج ---------- */
  function tableHtml(cat, rows, limit) {
    if (!rows.length) return '<p class="empty">لا توجد مواد في هذه القائمة.</p>';
    const c = CATS[cat];
    const shown = limit >= rows.length ? rows : rows.slice(0, limit);
    const val = v => (v === null || v === undefined) ? '—' : fmt(v);
    let h = '<div class="table-wrap"><table class="tbl"><thead><tr><th class="n">#</th><th>المادة</th><th class="n">معرض الغدير</th><th class="n">الحرفيين</th>';
    if (c.diff) h += '<th class="n">الفرق</th>';
    h += '</tr></thead><tbody>';
    shown.forEach((r, i) => {
      h += `<tr><td class="idx n">${i + 1}</td><td class="name">${esc(r.name)}</td><td class="n">${val(r.gh)}</td><td class="n">${val(r.hr)}</td>`;
      if (c.diff) h += `<td class="n">${fmt(r.diff)}</td>`;
      h += '</tr>';
    });
    return h + '</tbody></table></div>';
  }

  function renderResults() {
    const r = state.results;
    $('#results').hidden = false;

    $('#cards').innerHTML = ['A', 'B', 'C'].map(id => {
      const c = CATS[id];
      if (!r.opts[id]) {
        return `<div class="cat off"><span class="cat-title">${c.num} ${c.short}</span><span class="cat-count">غير مفعّل</span></div>`;
      }
      const active = state.active === id;
      return `<button type="button" class="cat${active ? ' active' : ''}" data-cat="${id}">
        <span class="cat-title">${c.num} ${c.short}</span>
        <span class="cat-count">عدد النتائج: ${r[id].length} مادة</span>
        <span class="cat-go">عرض النتائج</span></button>`;
    }).join('');

    const id = state.active, rows = r[id], lim = state.limit[id];
    let html = `<h3 class="table-title">${CATS[id].title} — عدد النتائج: ${rows.length} مادة</h3>` + tableHtml(id, rows, lim);
    if (rows.length > lim) {
      html += `<button type="button" class="btn more" id="moreBtn">عرض المزيد (المتبقي ${rows.length - lim})</button>`;
    }
    $('#tableArea').innerHTML = html;
  }

  /* ---------- 7) التقرير و PDF ---------- */
  const pad = n => String(n).padStart(2, '0');
  const dateStr = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const timeStr = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;

  function buildReport() {
    const r = state.results;
    let h = `<h1>تقرير مقارنة المخزون</h1>
      <p class="rdate">تاريخ التقرير: ${dateStr(r.date)} — ${timeStr(r.date)}</p>
      <table class="rmeta"><tbody>
        <tr><th>المخزن الأول</th><td>معرض الغدير</td></tr>
        <tr><th>المخزن الثاني</th><td>الحرفيين</td></tr>
        <tr><th>الحد المحدد للكمية القليلة</th><td>${fmt(r.T)} (من 0 إلى ${fmt(r.T)} = كمية قليلة)</td></tr>`;
    for (const id of ['A', 'B', 'C']) {
      if (r.opts[id]) h += `<tr><th>${CATS[id].title}</th><td>${r[id].length} مادة</td></tr>`;
    }
    h += '</tbody></table>';
    for (const id of ['A', 'B', 'C']) {
      if (!r.opts[id]) continue;
      h += `<h2>${CATS[id].num} ${CATS[id].title} <span class="cnt">— عدد المواد: ${r[id].length}</span></h2>` + tableHtml(id, r[id], Infinity);
    }
    $('#report').innerHTML = h;
  }

  function openPreview() {
    if (!state.results) { showMsg('يرجى إجراء المقارنة أولًا.'); return; }
    buildReport();
    $('#reportView').hidden = false;
    $('#reportView').scrollTop = 0;
    document.body.classList.add('noscroll');
  }
  function closePreview() {
    $('#reportView').hidden = true;
    document.body.classList.remove('noscroll');
  }

  function exportPdf() {
    if (!state.results) { showMsg('يرجى إجراء المقارنة أولًا.'); return; }
    buildReport();
    const old = document.title;
    document.title = 'تقرير-مقارنة-المخزون-' + dateStr(state.results.date);
    const restore = () => { document.title = old; window.removeEventListener('afterprint', restore); };
    window.addEventListener('afterprint', restore);
    setTimeout(() => window.print(), 60);
  }

  /* ---------- تشغيل ---------- */
  function init() {
    renderStoreCards();

    $('#compareBtn').addEventListener('click', () => runCompare(false));

    let timer = null;
    $('#threshold').addEventListener('input', () => {
      clearTimeout(timer);
      timer = setTimeout(() => { if (state.results) runCompare(true); }, 250);
    });
    for (const id of ['A', 'B', 'C']) {
      $('#opt-' + id).addEventListener('change', () => { if (state.results) runCompare(true); });
    }

    $('#cards').addEventListener('click', e => {
      const b = e.target.closest('[data-cat]');
      if (!b || !state.results) return;
      state.active = b.dataset.cat;
      renderResults();
    });
    $('#tableArea').addEventListener('click', e => {
      if (e.target.id === 'moreBtn') { state.limit[state.active] += PAGE; renderResults(); }
    });

    $('#previewBtn').addEventListener('click', openPreview);
    $('#pdfBtn').addEventListener('click', exportPdf);
    $('#rvBack').addEventListener('click', closePreview);
    $('#rvPrint').addEventListener('click', exportPdf);
    document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('#reportView').hidden) closePreview(); });

    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
      const hadController = !!navigator.serviceWorker.controller;
      let reloaded = false;
      navigator.serviceWorker.addEventListener('controllerchange', () => {
        // نسخة جديدة: نعيد التحميل مرة واحدة فقط إذا لم يستورد المستخدم شيئًا بعد
        if (!hadController || reloaded || state.gh.rows || state.hr.rows) return;
        reloaded = true; location.reload();
      });
      navigator.serviceWorker.register('sw.js').catch(() => { /* يعمل التطبيق حتى بدون تسجيل */ });
    }
  }

  function afterEdit(key) {
    const st = state[key];
    if (!st.rows) return;
    computeCols(st);
    renderColumnSelects(key);
    recompute(key);
  }
  async function reloadOriginal(key) {
    const st = state[key];
    if (!st.file) throw new AppError('الملف الأصلي غير متاح. اختره من جديد.');
    await loadFile(key, st.file);
  }
  window.__stockApp = { state, $, esc, fmt, colLetter, tick, afterEdit, reloadOriginal };

  init();
})();
