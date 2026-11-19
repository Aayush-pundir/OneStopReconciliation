/* Recon OS — Excel report. One workbook: Summary · Exceptions · Carry-over · one sheet per product · Bank Lines · Files. */
(function () {
'use strict';
const E = window.OSEngine, ST = E.ST, LABEL = E.PRODUCT_LABEL;
const NAVY = 'FF0F2540', MONEY = '#,##0.00;[Red](#,##0.00)', DATE = 'DD-MM-YYYY';
const FILL = {
  'MATCHED': 'FFC6EFCE', 'MATCHED – REFERENCE MISMATCH': 'FFD7F0E6', 'DIFFERENCE': 'FFFFFF99', 'PENDING': 'FFDBEAFE', 'NOT IN BANK': 'FFFFC7CE',
  'ZERO – NO ENTRY EXPECTED': 'FFECEFF3', 'BANK ONLY': 'FFFFE6DD', 'PRIOR PERIOD': 'FFE7E9F5', 'NEXT PERIOD': 'FFE7E9F5', 'NOT IN OPS': 'FFFFE6DD' };
const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const isoToDate = iso => iso ? new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10))) : null;
const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

/* cols: {h, w, v:(row)=>value, t:'money'|'date'|'int'|'text', sum:bool}; opt: {statusCol: index, title} */
function addSheet(wb, name, cols, rows, opt) {
  opt = opt || {};
  const ws = wb.addWorksheet(name.slice(0, 31));
  let r0 = 1;
  if (opt.title) { ws.getCell(1, 1).value = opt.title; ws.getCell(1, 1).font = { bold: true, size: 12 }; r0 = 3; }
  const hr = ws.getRow(r0); cols.forEach((c, i) => { const cell = hr.getCell(i + 1); cell.value = c.h; cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; cell.fill = fill(NAVY); cell.alignment = { vertical: 'middle', wrapText: true, horizontal: c.t === 'money' ? 'right' : 'left' }; ws.getColumn(i + 1).width = c.w || 14; });
  hr.height = 30;
  rows.forEach((row, k) => {
    const xr = ws.getRow(r0 + 1 + k);
    cols.forEach((c, i) => {
      let v = c.v(row); if (c.t === 'date') v = isoToDate(v); if (v === '' || v === undefined) v = null;
      const cell = xr.getCell(i + 1); cell.value = v;
      if (c.t === 'money') cell.numFmt = MONEY; else if (c.t === 'date') cell.numFmt = DATE; else if (c.t === 'int') cell.numFmt = '0';
    });
    if (opt.statusCol !== undefined) { const s = cols[opt.statusCol].v(row); if (FILL[s]) xr.getCell(opt.statusCol + 1).fill = fill(FILL[s]); }
  });
  const last = r0 + rows.length;
  if (rows.length && cols.some(c => c.sum)) {
    const tr = ws.getRow(last + 1);
    tr.getCell(1).value = 'TOTAL'; tr.getCell(1).font = { bold: true };
    cols.forEach((c, i) => { if (c.sum) { const col = colName(i + 1); const total = E.r2(rows.reduce((s, x) => s + (Number(c.v(x)) || 0), 0)); tr.getCell(i + 1).value = { formula: 'SUBTOTAL(9,' + col + (r0 + 1) + ':' + col + last + ')', result: total }; tr.getCell(i + 1).numFmt = MONEY; } });
    tr.eachCell({ includeEmpty: true }, c => { c.font = { bold: true }; c.fill = fill('FFD9E1F2'); });
  }
  ws.views = [{ state: 'frozen', ySplit: r0 }];
  if (rows.length) ws.autoFilter = { from: { row: r0, column: 1 }, to: { row: last, column: cols.length } };
  return ws;
}

function productCols(p) {
  const d = (k) => r => r.detail[k];
  const base = [{ h: 'Date', w: 12, t: 'date', v: r => r.date }, { h: 'Cycle', w: 8, v: r => r.cycle }];
  const tail = [{ h: 'Bank', w: 17, t: 'money', v: r => r.bank, sum: true }, { h: 'Bank date', w: 12, t: 'date', v: r => r.bankDate ? r.bankDate.slice(0, 10) : null },
    { h: 'Difference', w: 15, t: 'money', v: r => (r.status === ST.MATCHED || r.status === ST.ZERO) ? 0 : r.diff, sum: true }, { h: 'Status', w: 30, v: r => r.status }, { h: 'Note', w: 50, v: r => r.note }];
  const m = (h, k) => ({ h, w: 15, t: 'money', v: d(k), sum: true });
  const exp = h => ({ h, w: 17, t: 'money', v: r => r.expected, sum: true });
  if (p === 'NETC') return base.concat([{ h: 'File', w: 34, v: r => r.file }, { h: 'Txn count', w: 10, t: 'int', v: d('count') }, m('Txn amount', 'txnAmt'), m('Taxable value', 'taxable'), m('GST @18%', 'gst'), m('Total service fee', 'svcTotal'), exp('Expected (DSR net)')]).concat(tail);
  if (p === 'NFS') return base.concat([{ h: 'File', w: 24, v: r => r.file }, { h: 'ATM txns', w: 9, t: 'int', v: d('atmCount') }, m('ATM amount', 'atmAmt'), m('ATM fee+GST', 'atmFeeGst'), m('BI fee+GST', 'biFeeGst'), m('Micro-ATM amount', 'microAmt'), m('Micro-ATM fee+GST', 'microFeeGst'),
    m('Switching fee+GST', 'swFeeGst'), m('Disputes', 'disputes'), m('Computed settlement total', 'computedTotal'), exp('Expected (NTSL final, as debit)')]).concat(tail);
  if (p === 'RUPAY') return base.concat([{ h: 'File', w: 34, v: r => r.file }, m('Sales', 'sales'), m('Cash at POS', 'cash'), m('Refunds', 'refunds'), m('qSPARC load', 'qsparc'), m('Chargebacks', 'chargebacks'), m('Re-presentments', 'represent'), m('NPCI fees', 'fees'), m('Interchange income', 'interchange'),
    m('GST', 'gst'), m('Gross total', 'gross'), exp('Expected (DSR final net)')]).concat(tail);
  return base.concat([{ h: 'File', w: 30, v: r => r.file }, m('NTSL debit', 'debit'), m('NTSL credit', 'credit'), exp('Expected (NTSL net)')]).concat(tail);
}

async function build(R, intake) {
  const wb = new ExcelJS.Workbook(); wb.created = new Date(); wb.creator = 'Recon OS';
  const ps = E.PRODUCTS.filter(p => R.products[p]);
  const cover = R.bank.cover, ci = cover.integrity, m = R.bank.meta;

  /* ── Summary ── */
  const ws = wb.addWorksheet('Summary');
  ws.columns = [{ width: 34 }, { width: 12 }, { width: 12 }, { width: 13 }, { width: 11 }, { width: 13 }, { width: 11 }, { width: 19 }, { width: 19 }, { width: 18 }, { width: 14 }, { width: 14 }, { width: 11 }];
  let r = 1; const put = (row, col, v, o) => { const c = ws.getCell(row, col); c.value = v; if (o) { if (o.b) c.font = { bold: true, color: o.white ? { argb: 'FFFFFFFF' } : undefined, size: o.size }; if (o.f) c.fill = fill(o.f); if (o.n) c.numFmt = o.n; if (o.al) c.alignment = { horizontal: o.al }; } return c; };
  put(r, 1, 'RECON OS — RECONCILIATION SUMMARY', { b: true, size: 14 }); r += 1;
  put(r, 1, 'Generated'); put(r, 2, new Date().toLocaleString('en-IN')); r++;
  put(r, 1, 'Bank statement'); put(r, 2, R.bank.file + ' (' + (m.accountName || m.account || '') + ')'); r++;
  put(r, 1, 'Statement period'); put(r, 2, E.fmtDMY(m.firstDate) + ' to ' + E.fmtDMY(m.lastDate)); r++;
  put(r, 1, 'Overall'); put(r, 2, R.overall.matched + ' of ' + R.overall.cycles + ' cycles matched; ' + R.overall.attention + ' item(s) need attention', { b: true }); r += 2;
  put(r, 1, 'RECONCILIATION BY PRODUCT', { b: true, white: true, f: NAVY }); for (let c = 2; c <= 13; c++) put(r, c, null, { f: NAVY }); r++;
  ['Product', 'Cycles', 'Matched', 'Differences', 'Pending', 'Not in bank', 'Zero', 'Source total', 'Bank total', 'Open difference', 'Prior-month bank lines', 'Next-month bank lines', 'Bank only'].forEach((h, i) => put(r, i + 1, h, { b: true, f: 'FFD9E1F2', al: i ? 'right' : 'left' })); r++;
  const first = r;
  ps.forEach(p => { const t = R.products[p].totals; [LABEL[p], t.cycles, t.matched + t.refmis, t.diff, t.pending, t.notBank, t.zero, t.expected, t.bank, t.open, t.prior, t.next, t.bankOnly].forEach((v, i) => { const c = put(r, i + 1, v); if (i >= 7 && i <= 9) c.numFmt = MONEY; }); r++; });
  if (ps.length) { put(r, 1, 'TOTAL', { b: true, f: 'FFD9E1F2' }); for (let c = 2; c <= 13; c++) { const col = colName(c); const tot = ps.reduce((s, p) => { const t = R.products[p].totals; return s + [0, t.cycles, t.matched + t.refmis, t.diff, t.pending, t.notBank, t.zero, t.expected, t.bank, t.open, t.prior, t.next, t.bankOnly][c - 1]; }, 0); const cell = put(r, c, { formula: 'SUM(' + col + first + ':' + col + (r - 1) + ')', result: E.r2(tot) }, { b: true, f: 'FFD9E1F2' }); if (c >= 8 && c <= 10) cell.numFmt = MONEY; } r++; }
  r++;
  const am = R.products.ADDMONEY;
  put(r, 1, 'ADD MONEY TO POOL', { b: true, white: true, f: NAVY }); for (let c = 2; c <= 13; c++) put(r, c, null, { f: NAVY }); r++;
  if (am) { const t = am.totals; [['Ops UTRs', t.rows], ['Matched', t.matched], ['Amount mismatch', t.diff], ['Not in bank', t.notBank], ['Not in Ops', t.notOps], ['Total Ops amount (matched)', t.opsMatched], ['Total bank amount (matched)', t.bankMatched], ['Net unreconciled difference', t.netDiff]].forEach(([k, v]) => { put(r, 1, k, { b: true }); const c = put(r, 2, v); if (typeof v === 'number' && Math.abs(v) > 999) c.numFmt = MONEY; r++; }); }
  else { put(r, 1, 'No Ops file loaded — ' + (R.addMoneyPending ? R.addMoneyPending.lines + ' bank credits (' + E.r2(R.addMoneyPending.amount).toLocaleString('en-IN') + ') waiting' : '')); r++; }
  r++;
  put(r, 1, 'BANK STATEMENT', { b: true, white: true, f: NAVY }); for (let c = 2; c <= 13; c++) put(r, c, null, { f: NAVY }); r++;
  [['Opening balance', ci.opening], ['+ Credits', cover.credit], ['− Debits', cover.debit], ['Closing balance (statement)', ci.closing], ['Closing balance (computed)', ci.computedClosing], ['Statement check', ci.ok === null ? 'no balances in file' : ci.ok ? 'OK — adds up' : 'DOES NOT ADD UP']].forEach(([k, v]) => { put(r, 1, k, { b: true }); const c = put(r, 2, v); if (typeof v === 'number') c.numFmt = MONEY; r++; });
  r++;
  ['Claimed by', 'Lines', 'Debits', 'Credits'].forEach((h, i) => put(r, i + 1, h, { b: true, f: 'FFD9E1F2' })); r++;
  Object.entries(cover.byProduct).sort((a, b) => b[1].lines - a[1].lines).forEach(([k, v]) => { put(r, 1, LABEL[k] || 'Not part of any reconciliation'); put(r, 2, v.lines); put(r, 3, v.debit).numFmt = MONEY; put(r, 4, v.credit).numFmt = MONEY; r++; });
  r++;
  if (Object.keys(cover.other).length) { ['Lines not part of any reconciliation', 'Lines', 'Debits', 'Credits'].forEach((h, i) => put(r, i + 1, h, { b: true, f: 'FFD9E1F2' })); r++; Object.entries(cover.other).forEach(([k, v]) => { put(r, 1, k); put(r, 2, v.lines); put(r, 3, v.debit).numFmt = MONEY; put(r, 4, v.credit).numFmt = MONEY; r++; }); r++; }
  if (R.warnings.length) { put(r, 1, 'NOTES FROM THE RUN', { b: true, f: 'FFFFF4E0' }); r++; R.warnings.forEach(w => { put(r, 1, w); r++; }); }

  /* ── Exceptions ── */
  const ex = [];
  ps.forEach(p => { const P = R.products[p];
    P.rows.forEach(x => { if ([ST.DIFF, ST.NOTBANK, ST.PENDING, ST.REFMIS].includes(x.status)) ex.push({ product: LABEL[p], date: x.date, cycle: x.cycle, type: x.status, expected: x.expected, bank: x.bank, diff: x.diff, note: x.note, file: x.file }); });
    P.bankOnly.filter(b => b.cls === 'BANK ONLY').forEach(b => ex.push({ product: LABEL[p], date: b.date, cycle: b.cycle, type: 'BANK ONLY', expected: null, bank: b.amount, diff: -b.amount, note: b.note, file: '' })); });
  if (am) { am.rows.filter(x => x.status === ST.DIFF || x.status === ST.NOTBANK).forEach(x => ex.push({ product: LABEL.ADDMONEY, date: x.date, cycle: x.utr, type: x.status, expected: x.expected, bank: x.bank, diff: x.diff, note: x.name, file: '' }));
    am.bankOnly.forEach(b => ex.push({ product: LABEL.ADDMONEY, date: b.date, cycle: b.utr, type: 'NOT IN OPS', expected: null, bank: b.amount, diff: -b.amount, note: b.narr, file: '' })); }
  ex.sort((a, b) => (a.date || '') < (b.date || '') ? -1 : 1);
  addSheet(wb, 'Exceptions', [{ h: 'Product', w: 18, v: r => r.product }, { h: 'Date', w: 12, t: 'date', v: r => r.date }, { h: 'Cycle / UTR', w: 20, v: r => r.cycle }, { h: 'Type', w: 30, v: r => r.type },
    { h: 'Source', w: 17, t: 'money', v: r => r.expected, sum: true }, { h: 'Bank', w: 17, t: 'money', v: r => r.bank, sum: true }, { h: 'Difference', w: 17, t: 'money', v: r => r.diff, sum: true }, { h: 'Note', w: 70, v: r => r.note }], ex, { statusCol: 3 });

  /* ── Carry-over ── */
  const co = [];
  ps.forEach(p => R.products[p].bankOnly.filter(b => b.cls !== 'BANK ONLY').forEach(b => co.push({ product: LABEL[p], date: b.date, cycle: b.cycle, type: b.cls, amount: b.amount, posted: b.lines[0].date, narr: b.lines.map(l => l.narr).join(' + ') })));
  addSheet(wb, 'Carry-over', [{ h: 'Product', w: 18, v: r => r.product }, { h: 'Settles cycle of', w: 14, t: 'date', v: r => r.date }, { h: 'Cycle', w: 8, v: r => r.cycle }, { h: 'Type', w: 16, v: r => r.type },
    { h: 'Posted on', w: 12, t: 'date', v: r => r.posted }, { h: 'Amount', w: 17, t: 'money', v: r => r.amount, sum: true }, { h: 'Bank narrative', w: 60, v: r => r.narr }], co, { statusCol: 3, title: 'Bank lines that settle cycles outside the loaded files (earlier or later month) — not errors' });

  /* ── product sheets ── */
  const sheetName = { NETC: 'FASTag (NETC)', NFS: 'NFS', RUPAY: 'RuPay', UPI: 'UPI' };
  ps.forEach(p => { const cols = productCols(p); addSheet(wb, sheetName[p], cols, R.products[p].rows, { statusCol: cols.findIndex(c => c.h === 'Status'), title: LABEL[p] + ' — amounts from the bank\'s point of view: (red brackets) = debit from the account' }); });
  if (am) {
    const rows = am.rows.map(x => ({ utr: x.utr, name: x.name, ops: x.expected, bank: x.bank, diff: x.diff, status: x.status })).concat(am.bankOnly.map(b => ({ utr: b.utr, name: b.narr, ops: null, bank: b.amount, diff: -b.amount, status: 'NOT IN OPS' })));
    addSheet(wb, 'Add Money', [{ h: 'Final UTR', w: 22, v: r => r.utr }, { h: 'Enterprise name / bank narrative', w: 50, v: r => r.name }, { h: 'Ops amount', w: 17, t: 'money', v: r => r.ops, sum: true }, { h: 'Bank amount', w: 17, t: 'money', v: r => r.bank, sum: true },
      { h: 'Difference', w: 17, t: 'money', v: r => r.diff, sum: true }, { h: 'Status', w: 18, v: r => r.status }], rows, { statusCol: 5, title: 'Add money to Pool (rules of the Reconciliation Automation Macro v3)' });
  }

  /* ── bank lines ── */
  addSheet(wb, 'Bank Lines', [{ h: '#', w: 7, t: 'int', v: r => r.row }, { h: 'Date', w: 12, t: 'date', v: r => r.date }, { h: 'Narrative', w: 70, v: r => r.narr }, { h: 'Debit', w: 17, t: 'money', v: r => r.debit || null, sum: true }, { h: 'Credit', w: 17, t: 'money', v: r => r.credit || null, sum: true },
    { h: 'Claimed by', w: 18, v: r => LABEL[r.claimedBy] || 'Not part of any reconciliation' }, { h: 'Refers to date', w: 14, t: 'date', v: r => r.cls ? r.cls.date : null }, { h: 'Cycle', w: 8, v: r => r.cls ? r.cls.cycle : '' }, { h: 'Result', w: 30, v: r => r.result || '' }], R.bank.lines, { statusCol: 8 });

  /* ── files ── */
  addSheet(wb, 'Files', [{ h: 'File (inside zips: outer › inner)', w: 80, v: r => r.path }, { h: 'Detected as', w: 18, v: r => r.product ? (LABEL[r.product] || (r.product === 'OPS' ? 'Add Money Ops' : r.product === 'BANK' ? 'Bank statement' : r.product)) : '' },
    { h: 'Date', w: 12, t: 'date', v: r => r.date }, { h: 'Cycle', w: 8, v: r => r.cycle }, { h: 'Status', w: 12, v: r => r.status }, { h: 'Note', w: 50, v: r => r.note }], intake.files);
  return wb.xlsx.writeBuffer();
}
async function download(R, intake) {
  const buf = await build(R, intake), d = new Date(), p = n => String(n).padStart(2, '0');
  const per = E.PRODUCTS.filter(x => R.products[x]).map(x => R.products[x].period.start).sort()[0] || R.bank.meta.firstDate || '';
  saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'Recon_OS_' + (per ? per.slice(0, 7) : 'report') + '_' + p(d.getDate()) + p(d.getMonth() + 1) + String(d.getFullYear()).slice(2) + '.xlsx');
}
window.OSExport = { build, download };
})();
