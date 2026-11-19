/* Recon OS — screen logic. Talks to OSEngine (engine.js) and OSExport (export.js). */
(function () {
'use strict';
const E = window.OSEngine, ST = E.ST;
const $ = id => document.getElementById(id);
const esc = s => String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const tick = () => new Promise(r => setTimeout(r, 0));
const LABEL = E.PRODUCT_LABEL;

/* ── formatting ─────────────────────────────────────────────────────────── */
const fmt = n => { if (n === null || n === undefined || n === '') return ''; const a = Math.abs(n).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return n < 0 ? '(' + a + ')' : a; };
const money = n => n === null || n === undefined || n === '' ? '' : '<span class="' + (n < 0 ? 'neg' : '') + '">' + fmt(n) + '</span>';
const fd = E.fmtDMY;
const PILL = { 'MATCHED': 'p-ok', 'MATCHED – REFERENCE MISMATCH': 'p-ref', 'DIFFERENCE': 'p-diff', 'PENDING': 'p-pen', 'NOT IN BANK': 'p-nb', 'ZERO – NO ENTRY EXPECTED': 'p-zero',
  'BANK ONLY': 'p-bo', 'PRIOR PERIOD': 'p-per', 'NEXT PERIOD': 'p-per', 'NOT IN OPS': 'p-no' };
const pill = s => s ? '<span class="pill ' + (PILL[s] || 'p-oth') + '">' + esc(s) + '</span>' : '';

/* ── state ───────────────────────────────────────────────────────────────── */
const S = { intake: fresh(), seen: new Set(), result: null, tab: 'overview', view: {}, busy: false };
function fresh() { return { bank: null, bankFile: null, ops: [], netc: [], nfs: [], rupay: [], upi: [], files: [], warnings: [] }; }

/* ── building entries from the browser's File objects ──────────────────── */
const fileEntry = (f, path) => ({ name: f.name, path: path || f.webkitRelativePath || f.name, size: f.size, getBuffer: async () => new Uint8Array(await f.arrayBuffer()) });
async function walkEntry(e, out) {
  if (e.isFile) { const f = await new Promise((res, rej) => e.file(res, rej)); out.push(fileEntry(f, (e.fullPath || f.name).replace(/^\//, ''))); }
  else if (e.isDirectory) {
    const rd = e.createReader(); let batch;
    do { batch = await new Promise((res, rej) => rd.readEntries(res, rej)); for (const c of batch) await walkEntry(c, out); } while (batch.length);   // readEntries returns ≤100 at a time
  }
}
async function collectDropped(dt) {
  const out = [], items = [...(dt.items || [])], ents = items.map(i => i.webkitGetAsEntry && i.webkitGetAsEntry()).filter(Boolean);   // read before any await
  if (ents.length) for (const e of ents) await walkEntry(e, out); else for (const f of dt.files) out.push(fileEntry(f));
  return out;
}

/* ── intake (incremental) ───────────────────────────────────────────────── */
function progress(i, n, name) { $('pbar').classList.add('vis'); $('pfill').style.width = Math.round(i / Math.max(n, 1) * 100) + '%'; $('ptxt').textContent = 'Reading ' + i + ' / ' + n + (name ? ' — ' + name : ''); }
function endProgress(msg) { $('pfill').style.width = '100%'; $('ptxt').textContent = msg || ''; setTimeout(() => $('pbar').classList.remove('vis'), 500); }
function merge(add, replaceBank) {
  const A = S.intake;
  ['ops', 'netc', 'nfs', 'rupay', 'upi'].forEach(k => A[k].push(...add[k]));
  A.files.push(...add.files);
  if (add.bank) {
    if (A.bank) A.files.forEach(f => { if (f.product === 'BANK' && f.status === 'ok' && f.name === A.bankFile && !add.files.includes(f)) { f.status = 'replaced'; f.note = 'replaced by a newer statement'; } });
    A.bank = add.bank; A.bankFile = add.bankFile;
  }
}
async function addEntries(entries, opt) {
  opt = opt || {};
  const fresh1 = entries.filter(e => { const k = (e.path || e.name) + '|' + e.size; if (S.seen.has(k)) return false; S.seen.add(k); return true; });
  if (!fresh1.length) { endProgress('Those files are already loaded.'); return; }
  S.busy = true; hideResults(); $('runBtn').disabled = true;
  try {
    const add = await E.intake(fresh1, { onProgress: progress, yield: tick, onExpand: (n, q) => { $('ptxt').textContent = 'Opening zips… ' + n + ' files found' + (q ? ' (' + q + ' queued)' : ''); } });
    if (opt.bankOnly) {
      if (!add.bank) { $('bankStat').className = 'fstat er'; $('bankStat').textContent = '✗ ' + (fresh1[0] && fresh1[0].name) + ' does not look like a bank statement (needs Narrative/Transaction Details, Debit and Credit columns).'; $('bankDrop').classList.add('bad'); endProgress(''); renderIntake(); S.busy = false; return; }
      add.ops = []; add.netc = []; add.nfs = []; add.rupay = []; add.upi = [];
    }
    merge(add);
    const err = add.files.filter(f => f.status === 'error').length, ign = add.files.filter(f => f.status === 'ignored').length;
    endProgress('Read ' + add.files.length + ' files' + (err ? ' · ' + err + ' could not be read' : '') + (ign ? ' · ' + ign + ' ignored' : ''));
  } catch (e) { console.error(e); endProgress('Error: ' + e.message); }
  S.busy = false; renderIntake();
}

/* ── intake screen ───────────────────────────────────────────────────────── */
function counts() {
  const c = { NETC: 0, NFS: 0, RUPAY: 0, UPI: 0, OPS: 0, ignored: 0, error: 0, duplicate: 0 };
  S.intake.files.forEach(f => {
    if (f.status === 'ignored') c.ignored++; else if (f.status === 'error') c.error++; else if (f.status === 'duplicate') { c.duplicate++; if (c[f.product] !== undefined) c[f.product]++; }
    else if (f.status === 'ok' && c[f.product] !== undefined) c[f.product]++;
  });
  return c;
}
function renderIntake() {
  const A = S.intake, c = counts();
  // bank card
  const bs = $('bankStat'), bd = $('bankDrop');
  if (A.bank) {
    const m = A.bank.meta; bd.classList.remove('bad'); bd.classList.add('ok'); bs.className = 'fstat ok';
    bs.innerHTML = '✓ <b>' + esc(A.bankFile) + '</b><br>' + A.bank.lines.length.toLocaleString('en-IN') + ' transactions · ' + fd(m.firstDate) + ' → ' + fd(m.lastDate) +
      (m.opening !== null ? '<br>Opening ₹' + fmt(m.opening) + ' · Closing ₹' + fmt(m.closing) : '') + (m.accountName ? '<br>' + esc(m.accountName) : '');
  } else { bd.classList.remove('ok'); if (!bd.classList.contains('bad')) { bs.className = 'fstat'; bs.textContent = ''; } }
  // chips
  const chip = (t, n, cls) => '<span class="chip ' + (cls || (n ? '' : 'z')) + '">' + t + ' <b>' + n + '</b></span>';
  $('chips').innerHTML = chip('FASTag (NETC)', c.NETC) + chip('NFS', c.NFS) + chip('RuPay', c.RUPAY) + chip('UPI', c.UPI) + chip('Add Money Ops', c.OPS) +
    (c.duplicate ? chip('duplicates', c.duplicate, 'w') : '') + (c.ignored ? chip('ignored', c.ignored, 'z') : '') + (c.error ? chip('unreadable', c.error, 'e') : '');
  // coverage
  const any = c.NETC + c.NFS + c.RUPAY + c.UPI + c.OPS;
  $('covCard').classList.toggle('hidden', !any);
  if (any) {
    const cov = E.coverage(A); let html = '', tot = 0, exp = 0, miss = 0;
    for (const p of E.PRODUCTS) {
      const v = cov[p]; if (!v) { html += '<div class="nm">' + LABEL[p] + '</div><div class="rt">no files loaded</div><div></div>'; continue; }
      tot += v.found; exp += v.expected; miss += v.expected - v.found;
      const cells = v.days.map(d => '<span class="cell ' + (d.found === d.expected ? 'f' : d.found ? 'p' : 'n') + '" title="' + fd(d.date) + ' — ' + d.found + '/' + d.expected + (d.missing.length ? ' (missing ' + d.missing.join(', ') + ')' : '') + '">' + d.date.slice(8) + '</span>').join('');
      html += '<div class="nm">' + LABEL[p] + '</div><div class="strip">' + cells + '</div><div class="rt"><b>' + v.found + '/' + v.expected + '</b> cycles · ' + fd(v.first).slice(0, 5) + ' → ' + fd(v.last).slice(0, 5) + '</div>';
      if (v.missing.length) html += '<div class="miss">Missing: ' + v.missing.slice(0, 8).map(m => fd(m.date).slice(0, 5) + ' ' + m.cycles.join('/')).join(' · ') + (v.missing.length > 8 ? ' … +' + (v.missing.length - 8) + ' more days' : '') + '</div>';
    }
    html += '<div class="nm">Add Money</div><div class="rt">' + (A.ops.length ? '✓ Ops file: ' + esc(A.ops.map(o => o.name).join(', ')) : 'no Ops file — optional; add it to reconcile the Add Money credits') + '</div><div></div>';
    $('cov').innerHTML = html;
    $('covSum').textContent = tot + ' of ' + exp + ' expected cycles present' + (miss ? ' · ' + miss + ' missing' : ' ✓');
    $('covNote').textContent = 'Each square is one day (green = all cycles present, amber = some missing, red = none). Missing cycles are listed as exceptions only if the bank shows a settlement for them.';
  }
  // run button
  const ready = !!A.bank && any > 0;
  $('runBtn').disabled = !ready || S.busy;
  $('runHint').textContent = ready ? 'Ready — ' + any + ' source files × ' + A.bank.lines.length.toLocaleString('en-IN') + ' bank lines' : (!A.bank && !any ? 'Add the bank statement and the settlement files to begin' : !A.bank ? 'Add the bank statement' : 'Add the settlement files (folder, files or zip)');
}
function hideResults() { $('results').classList.add('hidden'); S.result = null; }

/* ── pickers & drops ─────────────────────────────────────────────────────── */
$('bankIn').addEventListener('change', async e => { const f = e.target.files[0]; if (!f) return; S.seen.delete((f.name) + '|' + f.size); $('bankDrop').classList.remove('bad'); await addEntries([fileEntry(f)], { bankOnly: true }); e.target.value = ''; });
$('folderIn').addEventListener('change', async e => { const l = [...e.target.files].map(f => fileEntry(f)); e.target.value = ''; await addEntries(l); });
$('filesIn').addEventListener('change', async e => { const l = [...e.target.files].map(f => fileEntry(f)); e.target.value = ''; await addEntries(l); });
let dragDepth = 0;
document.addEventListener('dragenter', e => { e.preventDefault(); dragDepth++; document.body.classList.add('dragging'); });
document.addEventListener('dragleave', e => { e.preventDefault(); if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dragging'); } });
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop', async e => {
  e.preventDefault(); dragDepth = 0; document.body.classList.remove('dragging');
  const onBank = e.target.closest && e.target.closest('#bankDrop');
  const l = await collectDropped(e.dataTransfer);
  await addEntries(l, onBank ? { bankOnly: l.length === 1 } : {});
});
$('clearBtn').addEventListener('click', () => { S.intake = fresh(); S.seen = new Set(); hideResults(); $('bankDrop').classList.remove('bad'); $('ptxt').textContent = ''; renderIntake(); });

/* tolerances */
(function () {
  const names = { NETC: 'FASTag (NETC)', NFS: 'NFS', RUPAY: 'RuPay', UPI: 'UPI', ADDMONEY: 'Add Money' };
  $('tols').innerHTML = Object.keys(names).map(k => '<label>' + names[k] + '<input type="number" step="0.01" min="0" id="tol_' + k + '" value="' + E.DEFAULT_TOL[k] + '"></label>').join('');
})();
const readTol = () => { const t = {}; for (const k of Object.keys(E.DEFAULT_TOL)) { const v = parseFloat($('tol_' + k).value); t[k] = isNaN(v) ? E.DEFAULT_TOL[k] : v; } return t; };

/* ── run ─────────────────────────────────────────────────────────────────── */
$('runBtn').addEventListener('click', async () => {
  $('runBtn').disabled = true; $('runBtn').textContent = '⏳ Reconciling…'; await tick();
  try {
    S.result = E.reconcileAll(S.intake, { tol: readTol() }); S.tab = 'overview'; S.view = {};
    renderResults(); $('results').classList.remove('hidden'); $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (e) { console.error(e); $('runHint').textContent = '✗ Reconciliation failed: ' + e.message; }
  $('runBtn').textContent = '⚡ Reconcile everything'; renderIntake();
});

/* ── results ─────────────────────────────────────────────────────────────── */
function present() { return [...E.PRODUCTS, 'ADDMONEY'].filter(p => S.result.products[p]); }
function pstatus(p) {
  const t = S.result.products[p].totals;
  if (p === 'ADDMONEY') return t.diff + t.notBank ? 'er' : t.notOps ? 'wn' : 'ok';
  if (t.diff + t.notBank + t.bankOnly > 0) return 'er';
  if (t.pending + t.refmis > 0) return 'wn';
  return t.cycles === t.zero ? 'zz' : 'ok';
}
const STLAB = { ok: 'Reconciled', wn: 'Check', er: 'Action needed', zz: 'Nothing due' };
function renderResults() {
  const R = S.result, ps = present();
  // banner
  const att = R.overall.attention, ci = R.bank.cover.integrity;
  const nops = R.overall.notOps || 0;
  let cls = att ? 'wn' : nops ? 'wn' : 'ok', head = att ? att + ' item' + (att > 1 ? 's' : '') + ' need attention' : nops ? 'All cycles reconcile ✓' : 'Everything reconciles ✓';
  if (ps.some(p => pstatus(p) === 'er')) cls = 'er';
  const carry = ps.filter(p => p !== 'ADDMONEY').reduce((s, p) => s + R.products[p].totals.prior + R.products[p].totals.next, 0);
  const per = ps.filter(p => p !== 'ADDMONEY').map(p => R.products[p].period);
  const pStart = per.map(x => x.start).sort()[0], pEnd = per.map(x => x.end).sort().pop();
  $('banner').innerHTML = '<div class="banner ' + cls + '"><div><div class="big1">' + head + '</div><div class="sm">' + R.overall.matched.toLocaleString('en-IN') + ' of ' + R.overall.cycles.toLocaleString('en-IN') + ' cycles matched' +
    (nops ? ' · ' + nops + ' bank credits are not in the Ops file' : '') + (carry ? ' · ' + carry + ' bank lines belong to other months (carry-over)' : '') + ' · period ' + (pStart ? fd(pStart) + ' → ' + fd(pEnd) : '—') + ' · statement ' + fd(R.bank.meta.firstDate) + ' → ' + fd(R.bank.meta.lastDate) +
    (ci.ok === false ? ' · <b>statement balances do not add up</b>' : '') + '</div></div><div class="sp btns"><button class="btn btn-g" id="dlBtn">⬇ Download full report (Excel)</button>' +
    (E.PRODUCTS.some(p => R.products[p]) ? '<button class="btn btn-o" id="csBtn" title="The month\'s control sheets in the layout of the FASTag, NFS, RuPay and UPI control-sheet files">⬇ Control sheet download</button>' : '') + '</div></div>';
  $('dlBtn').addEventListener('click', downloadReport);
  if ($('csBtn')) $('csBtn').addEventListener('click', downloadControl);
  // product cards
  $('pcards').innerHTML = ps.map(p => {
    const P = R.products[p], t = P.totals, st = pstatus(p);
    if (p === 'ADDMONEY') return '<div class="pc ' + st + '" data-go="ADDMONEY"><div class="h">' + LABEL[p] + '<span class="st ' + st + '">' + STLAB[st] + '</span></div><div class="big2">' + t.matched + ' <small>/ ' + t.rows + ' Ops UTRs matched</small></div>' +
      '<div class="ln"><span>Ops matched</span><b>' + fmt(t.opsMatched) + '</b></div><div class="ln"><span>Differences / not in bank</span><b>' + t.diff + ' / ' + t.notBank + '</b></div><div class="ln"><span>Bank credits not in Ops</span><b>' + t.notOps + ' · ' + fmt(t.notOpsAmt) + '</b></div></div>';
    return '<div class="pc ' + st + '" data-go="' + p + '"><div class="h">' + LABEL[p] + '<span class="st ' + st + '">' + STLAB[st] + '</span></div><div class="big2">' + (t.matched + t.refmis) + ' <small>/ ' + (t.cycles - t.zero) + ' cycles' + (t.zero ? ' (+' + t.zero + ' zero)' : '') + '</small></div>' +
      '<div class="ln"><span>Source total</span><b>' + fmt(t.expected) + '</b></div><div class="ln"><span>Bank (this period)</span><b>' + fmt(t.bank) + '</b></div>' +
      '<div class="ln"><span>Open difference</span><b class="' + (t.open ? 'neg' : '') + '">' + fmt(t.open) + '</b></div><div class="ln"><span>Other-month bank lines</span><b>' + (t.prior + t.next) + '</b></div></div>';
  }).join('');
  document.querySelectorAll('.pc').forEach(el => el.addEventListener('click', () => showTab(el.dataset.go)));
  // tabs
  const ex = exceptionsList(), co = carryList();
  const tabs = [['overview', 'Overview', ''], ['exceptions', 'Exceptions', ex.length ? '<span class="n r">' + ex.length + '</span>' : '<span class="n">0</span>'], ['carry', 'Carry-over', '<span class="n">' + co.length + '</span>']]
    .concat(ps.map(p => [p, LABEL[p], ''])).concat([['bank', 'Bank lines', '<span class="n">' + R.bank.lines.length + '</span>'], ['files', 'Files', '<span class="n">' + S.intake.files.length + '</span>']]);
  $('tabs').innerHTML = tabs.map(t => '<button class="tab" data-t="' + t[0] + '">' + t[1] + t[2] + '</button>').join('');
  document.querySelectorAll('#tabs .tab').forEach(b => b.addEventListener('click', () => showTab(b.dataset.t)));
  showTab(S.tab);
}
function showTab(t) {
  S.tab = t; document.querySelectorAll('#tabs .tab').forEach(b => b.classList.toggle('act', b.dataset.t === t));
  const R = S.result, P = $('panel');
  if (t === 'overview') P.innerHTML = overviewHtml();
  else if (t === 'exceptions') filtered('exceptions', exceptionsList(), exCols(), { prod: true, status: ['DIFFERENCE', 'NOT IN BANK', 'PENDING', 'BANK ONLY', 'NOT IN OPS', 'MATCHED – REFERENCE MISMATCH'], noteHtml: 'Everything that needs a look. Carry-over bank lines (other months) are on their own tab.' });
  else if (t === 'carry') filtered('carry', carryList(), carryCols(), { prod: true, status: ['PRIOR PERIOD', 'NEXT PERIOD'], noteHtml: 'These bank lines settle cycles <b>outside</b> the loaded files (earlier or later month). They are not errors; they reconcile in that month\'s run.' });
  else if (t === 'bank') filtered('bank', R.bank.lines, bankCols(), { prod: true, status: ['MATCHED', 'MATCHED – REFERENCE MISMATCH', 'DIFFERENCE', 'PENDING', 'NOT IN BANK', 'BANK ONLY', 'PRIOR PERIOD', 'NEXT PERIOD', 'NOT IN OPS', '—'], statusKey: l => l.result || '—', prodKey: l => l.claimedBy || 'OTHER', noteHtml: 'Every line of the bank statement and what happened to it.' });
  else if (t === 'files') filtered('files', S.intake.files, fileCols(), { status: ['ok', 'duplicate', 'ignored', 'error', 'replaced'], statusKey: f => f.status, noteHtml: 'Every file found (including files inside zips) and how it was read.' });
  else productTab(t);
}

/* ── generic filtered table ─────────────────────────────────────────────── */
function filtered(id, rows, cols, o) {
  const v = S.view[id] || (S.view[id] = { q: '', st: '', pr: '' });
  const prodKey = o.prodKey || (r => r.product), stKey = o.statusKey || (r => r.type || r.status);
  $('panel').innerHTML = '<div class="note">' + (o.noteHtml || '') + '</div><div class="fbar"><input type="text" id="fq" placeholder="Search…" value="' + esc(v.q) + '">' +
    (o.status ? '<select id="fst"><option value="">All statuses</option>' + o.status.map(s => '<option' + (v.st === s ? ' selected' : '') + '>' + esc(s) + '</option>').join('') + '</select>' : '') +
    (o.prod ? '<select id="fpr"><option value="">All products</option>' + [...E.PRODUCTS, 'ADDMONEY', 'OTHER'].map(p => '<option value="' + p + '"' + (v.pr === p ? ' selected' : '') + '>' + (LABEL[p] || 'Other') + '</option>').join('') + '</select>' : '') +
    '<span class="cnt" id="fcnt"></span></div><div class="tw" id="tt"></div>';
  const draw = () => {
    const q = v.q.toLowerCase();
    const rs = rows.filter(r => (!v.st || stKey(r) === v.st) && (!v.pr || prodKey(r) === v.pr) && (!q || JSON.stringify(Object.values(r).filter(x => typeof x !== 'object')).toLowerCase().includes(q) || cols.some(c => String(c.t ? c.t(r) : '').toLowerCase().includes(q))));
    const MAX = 3000;
    $('tt').innerHTML = tableHtml(cols, rs.slice(0, MAX), o.foot === false ? null : rs);
    $('fcnt').textContent = rs.length.toLocaleString('en-IN') + ' of ' + rows.length.toLocaleString('en-IN') + (rs.length > MAX ? ' — first ' + MAX + ' shown (the Excel has all)' : '');
  };
  $('fq').addEventListener('input', e => { v.q = e.target.value; draw(); });
  if ($('fst')) $('fst').addEventListener('change', e => { v.st = e.target.value; draw(); });
  if ($('fpr')) $('fpr').addEventListener('change', e => { v.pr = e.target.value; draw(); });
  draw();
}
function tableHtml(cols, rows, footRows) {
  const head = cols.map(c => '<th class="' + (c.num ? 'num' : '') + '">' + c.h + '</th>').join('');
  const body = rows.map(r => '<tr>' + cols.map(c => '<td class="' + (c.cls || '') + (c.num ? ' num' : '') + '">' + (c.f ? c.f(r) : esc(r[c.k])) + '</td>').join('') + '</tr>').join('');
  let foot = '';
  if (footRows && cols.some(c => c.sum)) foot = '<tfoot><tr>' + cols.map((c, i) => '<td>' + (i === 0 ? 'TOTAL' : c.sum ? fmt(E.r2(footRows.reduce((s, r) => s + (c.sum(r) || 0), 0))) : '') + '</td>').join('') + '</tr></tfoot>';
  return '<table><thead><tr>' + head + '</tr></thead><tbody>' + body + '</tbody>' + foot + '</table>';
}

/* ── lists & columns ─────────────────────────────────────────────────────── */
function exceptionsList() {
  const R = S.result, out = [];
  for (const p of E.PRODUCTS) {
    const P = R.products[p]; if (!P) continue;
    P.rows.forEach(r => { if ([ST.DIFF, ST.NOTBANK, ST.PENDING, ST.REFMIS].includes(r.status)) out.push({ product: p, date: r.date, cycle: r.cycle, type: r.status, expected: r.expected, bank: r.bank, diff: r.diff, note: r.note, file: r.file }); });
    P.bankOnly.filter(b => b.cls === 'BANK ONLY').forEach(b => out.push({ product: p, date: b.date, cycle: b.cycle, type: 'BANK ONLY', expected: null, bank: b.amount, diff: -b.amount, note: b.note, file: '' }));
  }
  const am = R.products.ADDMONEY;
  if (am) {
    am.rows.filter(r => r.status === ST.DIFF || r.status === ST.NOTBANK).forEach(r => out.push({ product: 'ADDMONEY', date: r.date, cycle: r.utr, type: r.status, expected: r.expected, bank: r.bank, diff: r.diff, note: r.name, file: '' }));
    am.bankOnly.forEach(b => out.push({ product: 'ADDMONEY', date: b.date, cycle: b.utr, type: 'NOT IN OPS', expected: null, bank: b.amount, diff: -b.amount, note: b.narr, file: '' }));
  }
  return out.sort((a, b) => (a.date || '') < (b.date || '') ? -1 : 1);
}
function carryList() {
  const out = [];
  for (const p of E.PRODUCTS) { const P = S.result.products[p]; if (!P) continue; P.bankOnly.filter(b => b.cls !== 'BANK ONLY').forEach(b => out.push({ product: p, date: b.date, cycle: b.cycle, type: b.cls, amount: b.amount, bankDate: b.lines[0].date, narr: b.lines.map(l => l.narr).join(' + '), note: b.note })); }
  return out.sort((a, b) => a.product < b.product ? -1 : a.product > b.product ? 1 : a.date < b.date ? -1 : 1);
}
const exCols = () => [
  { h: 'Product', f: r => esc(LABEL[r.product]) }, { h: 'Date', f: r => fd(r.date), t: r => fd(r.date) }, { h: 'Cycle / UTR', cls: 'mono', f: r => esc(r.cycle) },
  { h: 'Type', f: r => pill(r.type), t: r => r.type }, { h: 'Source', num: true, f: r => money(r.expected), sum: r => r.expected }, { h: 'Bank', num: true, f: r => money(r.bank), sum: r => r.bank },
  { h: 'Difference', num: true, f: r => money(r.diff), sum: r => r.diff }, { h: 'Note', cls: 'wrap', f: r => esc(r.note) } ];
const carryCols = () => [
  { h: 'Product', f: r => esc(LABEL[r.product]) }, { h: 'Settles cycle of', f: r => fd(r.date), t: r => fd(r.date) }, { h: 'Cycle', f: r => esc(r.cycle) }, { h: 'Type', f: r => pill(r.type), t: r => r.type },
  { h: 'Posted', f: r => fd(r.bankDate) }, { h: 'Amount', num: true, f: r => money(r.amount), sum: r => r.amount }, { h: 'Bank narrative', cls: 'wrap', f: r => esc(r.narr) } ];
const bankCols = () => [
  { h: '#', f: r => r.row }, { h: 'Date', f: r => fd(r.date), t: r => fd(r.date) }, { h: 'Narrative', cls: 'wrap', f: r => esc(r.narr) }, { h: 'Debit', num: true, f: r => r.debit ? fmt(r.debit) : '', sum: r => r.debit },
  { h: 'Credit', num: true, f: r => r.credit ? fmt(r.credit) : '', sum: r => r.credit }, { h: 'Product', f: r => esc(LABEL[r.claimedBy] || 'Other'), t: r => LABEL[r.claimedBy] || 'Other' },
  { h: 'Refers to', f: r => r.cls ? fd(r.cls.date) + ' ' + esc(r.cls.cycle) : '' }, { h: 'Result', f: r => r.result ? pill(r.result) : '<span class="pill p-oth">not part of a reconciliation</span>', t: r => r.result || '' } ];
const fileCols = () => [
  { h: 'File', cls: 'wrap', f: r => esc(r.path) }, { h: 'Detected as', f: r => esc(r.product ? (LABEL[r.product] || (r.product === 'OPS' ? 'Add Money Ops' : r.product === 'BANK' ? 'Bank statement' : r.product)) : '—') },
  { h: 'Date', f: r => fd(r.date) }, { h: 'Cycle', f: r => esc(r.cycle === null || r.cycle === undefined ? '' : r.cycle) }, { h: 'Status', f: r => '<span class="pill ' + ({ ok: 'p-ok', duplicate: 'p-diff', ignored: 'p-oth', error: 'p-nb', replaced: 'p-oth' }[r.status] || 'p-oth') + '">' + r.status + '</span>' }, { h: 'Note', cls: 'wrap', f: r => esc(r.note) } ];

/* ── product tabs ────────────────────────────────────────────────────────── */
function productCols(p) {
  const base = [{ h: 'Date', f: r => fd(r.date), t: r => fd(r.date) }, { h: 'Cycle', f: r => esc(r.cycle) }];
  const tail = [{ h: 'Bank', num: true, f: r => money(r.bank), sum: r => r.bank }, { h: 'Bank date', f: r => fd(r.bankDate) }, { h: 'Difference', num: true, f: r => money(r.diff), sum: r => (r.status === ST.MATCHED || r.status === ST.ZERO) ? 0 : r.diff },
    { h: 'Status', f: r => pill(r.status), t: r => r.status }, { h: 'Note', cls: 'wrap', f: r => esc(r.note) }];
  const d = r => r.detail, n = (h, k) => ({ h, num: true, f: r => fmt(d(r)[k]), sum: r => d(r)[k] });
  if (p === 'NETC') return base.concat([{ h: 'File', cls: 'wrap', f: r => esc(r.file) }, { h: 'Txns', num: true, f: r => d(r).count }, n('Txn amount', 'txnAmt'), n('Taxable', 'taxable'), n('GST', 'gst'), n('Service fee', 'svcTotal'),
    { h: 'Expected (DSR net)', num: true, f: r => money(r.expected), sum: r => r.expected }]).concat(tail);
  if (p === 'NFS') return base.concat([{ h: 'ATM txns', num: true, f: r => d(r).atmCount }, n('ATM amount', 'atmAmt'), n('ATM fee+GST', 'atmFeeGst'), n('BI fee+GST', 'biFeeGst'), n('Micro amount', 'microAmt'), n('Micro fee+GST', 'microFeeGst'),
    n('Switching fee+GST', 'swFeeGst'), n('Disputes', 'disputes'), n('Computed total', 'computedTotal'), { h: 'Expected (NTSL final)', num: true, f: r => money(r.expected), sum: r => r.expected }]).concat(tail);
  if (p === 'RUPAY') return base.concat([n('Sales', 'sales'), n('Cash at POS', 'cash'), n('Refunds', 'refunds'), n('qSPARC', 'qsparc'), n('Chargebacks', 'chargebacks'), n('Re-present.', 'represent'), n('NPCI fees', 'fees'), n('Interchange', 'interchange'),
    n('GST', 'gst'), n('Gross', 'gross'), { h: 'Expected (DSR net)', num: true, f: r => money(r.expected), sum: r => r.expected }]).concat(tail);
  return base.concat([n('NTSL debit', 'debit'), n('NTSL credit', 'credit'), { h: 'Expected (NTSL net)', num: true, f: r => money(r.expected), sum: r => r.expected }]).concat(tail);          // UPI
}
function productTab(p) {
  const P = S.result.products[p];
  if (p === 'ADDMONEY') return addMoneyTab(P);
  const t = P.totals;
  const rows = P.rows;
  const info = '<div class="kv"><div><div class="l">Cycles</div><div class="v">' + t.cycles + '</div></div><div><div class="l">Matched</div><div class="v">' + (t.matched + t.refmis) + '</div></div><div><div class="l">Differences / not in bank</div><div class="v">' + t.diff + ' / ' + t.notBank + '</div></div>' +
    '<div><div class="l">Pending</div><div class="v">' + t.pending + '</div></div><div><div class="l">Source total</div><div class="v">' + fmt(t.expected) + '</div></div><div><div class="l">Bank total</div><div class="v">' + fmt(t.bank) + '</div></div>' +
    '<div><div class="l">Other-month bank lines</div><div class="v">' + t.prior + ' prior · ' + t.next + ' next</div></div></div>' +
    '<div class="note">Amounts are from the bank\'s point of view: <b>(brackets) = debit from the account</b>, plain = credit. Tolerance ₹' + P.tol + '. Typical posting lag seen: ' + P.lag + ' day(s).</div>';
  filtered('p_' + p, rows, productCols(p), { status: Object.values(ST), statusKey: r => r.status, noteHtml: '' });
  $('panel').insertAdjacentHTML('afterbegin', info);
}
function addMoneyTab(P) {
  const t = P.totals;
  const rows = P.rows.map(r => ({ utr: r.utr, name: r.name, expected: r.expected, bank: r.bank, diff: r.diff, status: r.status }))
    .concat(P.bankOnly.map(b => ({ utr: b.utr, name: b.narr, expected: null, bank: b.amount, diff: -b.amount, status: 'NOT IN OPS' })));
  const cols = [{ h: 'Final UTR', cls: 'mono', f: r => esc(r.utr) }, { h: 'Enterprise / narrative', cls: 'wrap', f: r => esc(r.name) }, { h: 'Ops amount', num: true, f: r => money(r.expected), sum: r => r.expected },
    { h: 'Bank amount', num: true, f: r => money(r.bank), sum: r => r.bank }, { h: 'Difference', num: true, f: r => money(r.diff), sum: r => r.diff }, { h: 'Status', f: r => pill(r.status), t: r => r.status }];
  filtered('p_ADDMONEY', rows, cols, { status: ['MATCHED', 'DIFFERENCE', 'NOT IN BANK', 'NOT IN OPS'], statusKey: r => r.status, noteHtml: '' });
  $('panel').insertAdjacentHTML('afterbegin', '<div class="kv"><div><div class="l">Ops UTRs</div><div class="v">' + t.rows + '</div></div><div><div class="l">Matched</div><div class="v">' + t.matched + '</div></div><div><div class="l">Amount mismatch</div><div class="v">' + t.diff + '</div></div>' +
    '<div><div class="l">Not in bank</div><div class="v">' + t.notBank + '</div></div><div><div class="l">Not in Ops</div><div class="v">' + t.notOps + '</div></div><div><div class="l">Ops / Bank matched</div><div class="v">' + fmt(t.opsMatched) + ' / ' + fmt(t.bankMatched) + '</div></div></div>' +
    '<div class="note">Rules of the Add Money macro: UTR read from the bank narrative; rows with the same UTR are added together. Only bank <b>credits</b> not already claimed by NETC/NFS/RuPay/UPI are considered.</div>');
}

/* ── overview ────────────────────────────────────────────────────────────── */
function overviewHtml() {
  const R = S.result, ps = E.PRODUCTS.filter(p => R.products[p]), c = R.bank.cover, ci = c.integrity;
  const rows = ps.map(p => { const t = R.products[p].totals; return '<tr><td><b>' + LABEL[p] + '</b></td><td class="num">' + t.cycles + '</td><td class="num">' + (t.matched + t.refmis) + '</td><td class="num">' + t.diff + '</td><td class="num">' + t.pending + '</td><td class="num">' + t.notBank + '</td><td class="num">' + t.zero + '</td>' +
    '<td class="num">' + money(t.expected) + '</td><td class="num">' + money(t.bank) + '</td><td class="num">' + money(t.open) + '</td><td class="num">' + t.prior + '</td><td class="num">' + t.next + '</td><td class="num">' + t.bankOnly + '</td></tr>'; }).join('');
  const sums = ps.reduce((a, p) => { const t = R.products[p].totals; a.cy += t.cycles; a.m += t.matched + t.refmis; a.d += t.diff; a.p += t.pending; a.n += t.notBank; a.z += t.zero; a.e += t.expected; a.b += t.bank; a.o += t.open; a.pr += t.prior; a.nx += t.next; a.bo += t.bankOnly; return a; }, { cy: 0, m: 0, d: 0, p: 0, n: 0, z: 0, e: 0, b: 0, o: 0, pr: 0, nx: 0, bo: 0 });
  let h = '<div class="sec">Reconciliation by product</div><div class="tw"><table><thead><tr><th>Product</th><th class="num">Cycles</th><th class="num">Matched</th><th class="num">Differences</th><th class="num">Pending</th><th class="num">Not in bank</th><th class="num">Zero</th><th class="num">Source total</th><th class="num">Bank total</th><th class="num">Open difference</th><th class="num">Prior lines</th><th class="num">Next lines</th><th class="num">Bank only</th></tr></thead><tbody>' + rows +
    '</tbody><tfoot><tr><td>TOTAL</td><td>' + sums.cy + '</td><td>' + sums.m + '</td><td>' + sums.d + '</td><td>' + sums.p + '</td><td>' + sums.n + '</td><td>' + sums.z + '</td><td>' + fmt(E.r2(sums.e)) + '</td><td>' + fmt(E.r2(sums.b)) + '</td><td>' + fmt(E.r2(sums.o)) + '</td><td>' + sums.pr + '</td><td>' + sums.nx + '</td><td>' + sums.bo + '</td></tr></tfoot></table></div>' +
    '<div class="note">Period = the calendar month(s) of the loaded files. Source/Bank totals cover that period only; bank lines of other months are carry-over (never false differences), and lines inside the period with no source file show as <b>Bank only</b> (usually a missing file).</div>';
  const am = R.products.ADDMONEY;
  if (am) h += '<div class="sec">Add money to Pool</div><div class="note">' + am.totals.matched + ' of ' + am.totals.rows + ' Ops UTRs matched · ' + am.totals.diff + ' amount mismatch · ' + am.totals.notBank + ' not in bank · ' + am.totals.notOps + ' bank credits not in Ops (₹' + fmt(am.totals.notOpsAmt) + ').</div>';
  else if (R.addMoneyPending) h += '<div class="sec">Add money to Pool</div><div class="note">No Ops file loaded. ' + R.addMoneyPending.lines + ' bank credits (₹' + fmt(R.addMoneyPending.amount) + ') are waiting to be reconciled: add the Add Money Ops file (drop it with the other files) and run again.</div>';
  h += '<div class="sec">Bank statement</div><div class="kv"><div><div class="l">Opening balance</div><div class="v">' + (ci.opening !== null ? fmt(ci.opening) : '—') + '</div></div><div><div class="l">+ Credits</div><div class="v">' + fmt(c.credit) + '</div></div><div><div class="l">− Debits</div><div class="v">' + fmt(c.debit) + '</div></div>' +
    '<div><div class="l">Closing balance (statement)</div><div class="v">' + (ci.closing !== null ? fmt(ci.closing) : '—') + '</div></div><div><div class="l">Closing balance (computed)</div><div class="v">' + (ci.computedClosing !== null ? fmt(ci.computedClosing) : '—') + '</div></div>' +
    '<div><div class="l">Statement check</div><div class="v ' + (ci.ok === false ? 'cross' : 'tick') + '">' + (ci.ok === null ? 'no balances in file' : ci.ok ? '✓ adds up' : '✗ does not add up') + '</div></div></div>';
  const cov = Object.entries(c.byProduct).sort((a, b) => b[1].lines - a[1].lines).map(([k, v]) => '<tr><td>' + (LABEL[k] || 'Not part of any reconciliation') + '</td><td class="num">' + v.lines + '</td><td class="num">' + fmt(v.debit) + '</td><td class="num">' + fmt(v.credit) + '</td></tr>').join('');
  h += '<div class="sec">Where every bank line went</div><div class="tw"><table><thead><tr><th>Claimed by</th><th class="num">Lines</th><th class="num">Debits</th><th class="num">Credits</th></tr></thead><tbody>' + cov + '</tbody></table></div>';
  const oth = Object.entries(c.other);
  if (oth.length) h += '<div class="sec">Lines not part of any reconciliation</div><div class="tw"><table><thead><tr><th>Type</th><th class="num">Lines</th><th class="num">Debits</th><th class="num">Credits</th></tr></thead><tbody>' + oth.sort((a, b) => b[1].lines - a[1].lines).map(([k, v]) => '<tr><td>' + esc(k) + '</td><td class="num">' + v.lines + '</td><td class="num">' + fmt(v.debit) + '</td><td class="num">' + fmt(v.credit) + '</td></tr>').join('') + '</tbody></table></div>' +
    '<div class="note">' + (R.products.ADDMONEY ? 'Credits here were not found in the Ops file or are not Add Money (e.g. gateway settlements, internal transfers).' : 'IMPS / RTGS / NEFT credits are Add Money receipts and gateway settlements; internal transfers and payouts are not reconciled by this module.') + '</div>';
  if (R.warnings.length) h += '<div class="sec">Notes from the run</div><div class="warns">' + R.warnings.map(w => '• ' + esc(w)).join('<br>') + '</div>';
  return h;
}

/* ── download ────────────────────────────────────────────────────────────── */
async function downloadReport() {
  const b = $('dlBtn'), old = b.textContent; b.disabled = true; b.textContent = 'Building Excel…'; await tick();
  try { await window.OSExport.download(S.result, S.intake); } catch (e) { console.error(e); $('ptxt').textContent = '✗ Excel error: ' + e.message; }
  b.disabled = false; b.textContent = old;
}

async function downloadControl() {
  const b = $('csBtn'), old = b.textContent; b.disabled = true; b.textContent = 'Building control sheets…'; await tick();
  try { await window.OSControlSheet.download(S.result); } catch (e) { console.error(e); $('ptxt').textContent = '✗ Control sheet error: ' + e.message; }
  b.disabled = false; b.textContent = old;
}

renderIntake();
window.__OS = S;        // for debugging / tests
})();
