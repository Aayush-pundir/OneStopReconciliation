/* Recon OS — "Control sheet download".
   One workbook that holds the month's control sheets in the layout of the finance team's own files:
     FASTag : control sheet (30 days × 3 cycles) + Accounting Entry
     NFS    : control sheet (days × 4 cycles, fee formulas) + Workingsheet
     RuPay  : control sheet (wide, days × 4 cycles) + Working Sheet
     UPI    : Summary + Pivot Final + Working
   Source figures come from the DSR/NTSL files, the bank column from the matched bank lines; every
   calculation column is a live Excel formula (results are also cached so the file reads without recalculation). */
(function () {
'use strict';
const E = window.OSEngine, ST = E.ST;

/* ── small helpers ─────────────────────────────────────────────────────── */
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const r2 = E.r2, n0 = v => (typeof v === 'number' && isFinite(v)) ? v : 0;
const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
const isoOf = (y, m, d) => y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
const toDate = iso => iso ? new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10))) : null;
const serial = iso => Math.round((Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000);
const colL = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const colN = s => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
const q = name => "'" + name.replace(/'/g, "''") + "'";                       // quoted sheet name for formulas
const THIN = { style: 'thin', color: { argb: 'FF000000' } }, BORD = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const CAL = { name: 'Calibri', size: 11 }, CALB = { name: 'Calibri', size: 11, bold: true }, AR9 = { name: 'Arial', size: 9 }, AR9B = { name: 'Arial', size: 9, bold: true };
const ACC2 = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', ACC0 = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)';
const TOT0 = '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ ', TOT2 = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ ';
const CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };

/* write a cell: v = value | {f: formula, r: cached result}; o = {font, fill, fmt, al, border} */
function W(ws, ref, v, o) {
  const c = ws.getCell(ref); o = o || {};
  if (v !== undefined && v !== null) c.value = (v && typeof v === 'object' && 'f' in v) ? (window.__OS_NOCACHE ? { formula: v.f } : { formula: v.f, result: v.r }) : v;   // __OS_NOCACHE: test hook
  if (o.font) c.font = o.font; if (o.fill) c.fill = o.fill; if (o.fmt) c.numFmt = o.fmt; if (o.al) c.alignment = o.al; if (o.border) c.border = o.border;
  return c;
}
function merge(ws, a, b, v, o) { ws.mergeCells(a + ':' + b); return W(ws, a, v, o); }
function sumRange(vals) { return r2(vals.reduce((s, x) => s + n0(x), 0)); }

/* months present in the reconciled data, per product */
function monthsOf(P) { return [...new Set((P ? P.rows : []).map(r => r.date.slice(0, 7)))].sort(); }
const monthLabel = ym => MON[+ym.slice(5, 7) - 1] + ' ' + ym.slice(2, 4);          // "Oct 26"
const monthLabel2 = ym => MON[+ym.slice(5, 7) - 1] + '-' + ym.slice(2, 4);         // "Oct-26"
function rowMap(P) { const m = new Map(); P.rows.forEach(r => m.set(r.date + '|' + r.cycle, r)); return m; }

/* ═══════════════════════════ FASTag ═══════════════════════════════════════ */
function fastag(wb, P, ym, accum) {
  const y = +ym.slice(0, 4), m = +ym.slice(5, 7), nd = daysIn(y, m), name = 'FASTag Control Sheet_' + monthLabel(ym);
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', xSplit: 3, ySplit: 6, zoomScale: 85 }] });
  [['A', 12.8], ['B', 12.8], ['C', 7], ['D', 12.6], ['E', 15.4], ['F', 13.7], ['G', 13.7], ['H', 16.8], ['I', 16], ['J', 14.2], ['K', 16], ['L', 4], ['M', 18], ['N', 18]].forEach(([c, w]) => { ws.getColumn(c).width = w; });
  ws.getRow(5).height = 42;
  const hd = (ref, v, o) => W(ws, ref, v, Object.assign({ font: CALB, al: CENTER, border: BORD }, o || {}));
  merge(ws, 'A5', 'A6', 'Date', { font: CALB, al: CENTER, border: BORD }); merge(ws, 'B5', 'B6', 'Date', { font: CALB, al: CENTER, border: BORD }); merge(ws, 'C5', 'C6', 'Cycle', { font: CALB, al: CENTER, border: BORD });
  merge(ws, 'D5', 'E5', 'Transactions', { font: CALB, al: CENTER, border: BORD, fill: fill('FFE6B9B8') }); ws.getCell('E5').border = BORD;
  merge(ws, 'F5', 'G5', 'Interchange Income', { font: CALB, al: CENTER, border: BORD, fill: fill('FF92D050') }); ws.getCell('G5').border = BORD;
  merge(ws, 'H5', 'H6', 'Net Settlememt', { font: CALB, al: CENTER, border: BORD, fill: fill('FFFFFF00') });
  merge(ws, 'I5', 'I6', 'Actual amount settled from Bank', { font: CALB, al: CENTER, border: BORD, fill: fill('FFFFFF00') });
  merge(ws, 'J5', 'J6', 'Date of Settlement in Bank', { font: CALB, al: CENTER, border: BORD }); merge(ws, 'K5', 'K6', 'Difference if any', { font: CALB, al: CENTER, border: BORD });
  ['Count', 'Amount', 'Taxable Value', 'GST'].forEach((h, i) => hd(colL(4 + i) + '6', h));
  [['M5', 'DSR txn amount (Dr, toll & parking)'], ['N5', 'Other DSR items in Amount (adjustments, chargebacks)']].forEach(([a, t]) => { merge(ws, a, a.replace('5', '6'), t, { font: { name: 'Calibri', size: 9, bold: true }, al: CENTER, border: BORD, fill: fill('FFD9D9D9') }); });
  const first = 7, last = first + nd * 3 - 1, rm = rowMap(P), sums = { D: 0, E: 0, F: 0, G: 0, H: 0, I: 0, K: 0 };
  for (let d = 1; d <= nd; d++) {
    const iso = isoOf(y, m, d), r0 = first + (d - 1) * 3, ser = serial(iso);
    ws.mergeCells('A' + r0 + ':A' + (r0 + 2));
    W(ws, 'A' + r0, d === 1 ? toDate(iso) : { f: 'A' + (r0 - 3) + '+1', r: ser }, { font: CAL, fmt: 'mm-dd-yy', al: CENTER, border: BORD });
    for (let c = 1; c <= 3; c++) {
      const r = r0 + (c - 1), row = rm.get(iso + '|' + c), det = row && row.detail;
      ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'M', 'N'].forEach(L => { ws.getCell(L + r).border = BORD; ws.getCell(L + r).font = CAL; });
      W(ws, 'B' + r, { f: c === 1 ? 'A' + r : 'B' + (r - 1), r: ser }, { fmt: 'mm-dd-yy', al: { vertical: 'middle', wrapText: true } });
      W(ws, 'C' + r, 'C' + c, { al: { horizontal: 'center' } });
      if (!det) { ws.getCell('C' + r).fill = fill('FFFFC7CE'); continue; }                                      // no DSR for this cycle
      const svc = det.svcTotal, E_ = r2(-row.expected + svc), F_ = det.taxable, G_ = F_ * 0.18, H_ = E_ - (F_ + G_), I_ = row.bank === null ? null : r2(-row.bank);
      W(ws, 'D' + r, det.count, { fmt: ACC2 }); W(ws, 'E' + r, E_, { fmt: ACC2 }); W(ws, 'F' + r, F_, { fmt: ACC2 });
      W(ws, 'G' + r, { f: 'F' + r + '*0.18', r: F_ * 0.18 }, { fmt: ACC2 }); W(ws, 'H' + r, { f: 'E' + r + '-(F' + r + '+G' + r + ')', r: H_ }, { fmt: ACC2 });
      if (I_ !== null) W(ws, 'I' + r, I_, { fmt: ACC2, al: { vertical: 'middle', wrapText: true } });
      if (row.bankDate) W(ws, 'J' + r, toDate(row.bankDate.slice(0, 10)), { fmt: 'd-mmm-yy', al: { vertical: 'middle', wrapText: true } });
      W(ws, 'K' + r, { f: 'H' + r + '-I' + r, r: H_ - n0(I_) }, { fmt: ACC0 });
      W(ws, 'M' + r, det.txnAmt, { fmt: ACC2 }); W(ws, 'N' + r, { f: 'E' + r + '-M' + r, r: r2(E_ - det.txnAmt) }, { fmt: ACC2 });
      sums.D += det.count; sums.E += E_; sums.F += F_; sums.G += F_ * 0.18; sums.H += H_; sums.I += n0(I_); sums.K += H_ - n0(I_);
    }
  }
  ['D', 'E', 'F', 'G', 'H', 'I', 'K'].forEach(L => W(ws, L + '4', { f: 'SUM(' + L + first + ':' + L + last + ')', r: sums[L] }, { font: CALB, fmt: TOT0 }));
  W(ws, 'I3', { f: 'H4-I4', r: sums.H - sums.I }, { font: CAL, fmt: TOT2 });
  accum.fastag = accum.fastag || []; accum.fastag.push({ ym, name, E: sums.E, I: sums.I, F: sums.F, G: sums.G });
  return ws;
}
function fastagEntry(wb, accum) {
  const ws = wb.addWorksheet('FASTag Accounting Entry'); ws.getColumn('A').width = 33.1;
  const hdr = fill('FFB2B1A9');
  W(ws, 'A2', 'Particulars', { font: CALB, fill: hdr, border: BORD });
  accum.fastag.forEach((s, i) => {
    const L = colL(2 + i); ws.getColumn(L).width = 14;
    W(ws, L + '2', toDate(s.ym + '-01'), { font: CALB, fill: hdr, border: BORD, fmt: 'mmm-yy' });
    W(ws, L + '3', { f: q(s.name) + '!E4', r: s.E }, { font: CAL, border: BORD, fmt: ACC0 });
    W(ws, L + '4', { f: q(s.name) + '!I4', r: s.I }, { font: CAL, border: BORD, fmt: ACC0 });
    W(ws, L + '5', { f: q(s.name) + '!F4+' + q(s.name) + '!G4', r: s.F + s.G }, { font: CAL, border: BORD, fmt: ACC0 });
    W(ws, L + '10', { f: L + '3-' + L + '4-' + L + '5', r: s.E - s.I - (s.F + s.G) }, { font: CAL, fmt: ACC0 });
    s.entryCol = L;
  });
  [['A3', 'Customer Wallet (Dr)'], ['A4', 'NETC A/c (Cr)'], ['A5', 'Omnicard B2C Uttar Pradesh (Cr)']].forEach(([a, t]) => W(ws, a, t, { font: CAL, border: BORD }));
  W(ws, 'A10', 'Diff', { font: CAL });
  return ws;
}

/* ═══════════════════════════ NFS ══════════════════════════════════════════ */
function nfs(wb, P, ym, accum) {
  const y = +ym.slice(0, 4), m = +ym.slice(5, 7), nd = daysIn(y, m), name = 'NFS ' + monthLabel2(ym), ws = wb.addWorksheet(name);
  [9.7, 5.3, 13.3, 14.7, 12.9, 8, 9, 13.3, 12.9, 6, 7.1, 6.4, 12.7, 11.1, 12.9, 7, 8, 8.7, 12.9, 7, 8, 6, 8.3, 14.7, 19, 14.3, 13.6, 10.4].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const first = 10, last = first + nd * 4 - 1, rm = rowMap(P);
  const hb = (ref, v, bg, white, size) => W(ws, ref, v, { font: { name: 'Calibri', size: size || 11, bold: true, color: white ? { argb: 'FFFFFFFF' } : undefined }, fill: bg ? fill(bg) : undefined, al: CENTER, border: BORD });
  const mh = (a, b, v, bg, white, bold) => { merge(ws, a, b, v, { font: { name: 'Calibri', size: 11, bold: bold !== false, color: white ? { argb: 'FFFFFFFF' } : undefined }, fill: bg ? fill(bg) : undefined, al: CENTER, border: BORD }); };
  mh('A7', 'A9', 'Date', null, false, false); mh('B7', 'B9', 'Cycle', null, false, false);
  mh('C7', 'Q7', 'Issuer WDL Approved Fee', 'FF548235', true); mh('R7', 'U8', 'Switiching Fee', 'FFBF9000', true); mh('V7', 'W8', 'Disputes', 'FFD9D9D9', false);
  mh('X7', 'X9', 'Settlement Total', 'FFFFD966', false, false); mh('Y7', 'Y9', 'Settlement from Bank', 'FFC6DEB5', false, false); mh('Z7', 'Z9', 'Settlement Date', 'FFC6DEB5', false, false); mh('AA7', 'AA9', 'Difference', null, false, false);
  const a8 = (a, b, v, bg) => merge(ws, a, b, v, { font: { name: 'Arial', size: 11, bold: true }, fill: fill(bg), al: { horizontal: 'center' }, border: BORD });
  a8('C8', 'G8', 'ATM', 'FFFFE699'); a8('H8', 'K8', 'Balance Inquiry', 'FFFFE699'); a8('L8', 'Q8', 'Micro ATM', 'FF8F98A6');
  ['Transaction Count', 'Transaction Amount', 'Fee(without GST)', 'GST', 'Total Fee', 'Transaction Count', 'Fee(without GST)', 'GST', 'Total Fee', 'Count', 'Total Transaction Amount', 'Transaction Amount(>100)', 'Fee(without GST)', 'GST', 'Total Fee', 'Total Count', 'Fee(without GST)', 'GST', 'Total Fee', 'Count', 'Amount']
    .forEach((h, i) => W(ws, colL(3 + i) + '9', h, { font: { name: 'Calibri', size: 9 }, border: BORD, al: { wrapText: true } }));
  ws.getRow(7).height = 32.7; ws.getRow(8).height = 40.5; ws.getRow(9).height = 40.5;
  const BLUE = fill('FFDAE3F3'), tot = {};
  const add = (L, v) => { tot[L] = (tot[L] || 0) + n0(v); };
  for (let d = 1; d <= nd; d++) {
    const iso = isoOf(y, m, d), r0 = first + (d - 1) * 4;
    ws.mergeCells('A' + r0 + ':A' + (r0 + 3)); W(ws, 'A' + r0, toDate(iso), { font: CAL, fmt: 'd-mmm-yy', al: CENTER, border: BORD });
    for (let c = 1; c <= 4; c++) {
      const r = r0 + c - 1, row = rm.get(iso + '|C' + c), det = row && row.detail;
      for (let k = 1; k <= 28; k++) { const cell = ws.getCell(r, k); cell.border = BORD; if (!cell.font || !cell.font.name) cell.font = CAL; }
      W(ws, 'B' + r, 'C' + c, { al: { vertical: 'middle' } });
      if (!det) { ws.getCell('B' + r).fill = fill('FFFFC7CE'); continue; }
      const C_ = det.atmCount, D_ = det.atmAmt, E_ = C_ * 19, F_ = E_ * 0.18, G_ = E_ + F_, H_ = det.biCount, I_ = H_ * 6, J_ = I_ * 0.18, K_ = I_ + J_, L_ = det.microCount, M_ = det.microAmt, N_ = M_,
        O_ = N_ * 0.005, P_ = O_ * 0.18, Q_ = O_ + P_, R_ = det.swCount, S_ = R_ * 0.3, T_ = S_ * 0.18, U_ = S_ + T_, V_ = det.disputeCount || 0, W_ = det.disputes || 0,
        X_ = D_ + G_ + M_ + Q_ + U_ - W_ + K_, Y_ = row.bank === null ? null : r2(-row.bank);
      const F = (L, f, res, fm) => W(ws, L + r, { f, r: res }, { fmt: fm || ACC2, fill: BLUE });
      W(ws, 'C' + r, C_); W(ws, 'D' + r, D_);
      F('E', 'C' + r + '*19', E_); F('F', 'E' + r + '*18%', F_); F('G', 'SUM(E' + r + ':F' + r + ')', G_);
      W(ws, 'H' + r, H_); F('I', 'H' + r + '*6', I_); F('J', 'I' + r + '*18%', J_); F('K', 'I' + r + '+J' + r, K_);
      W(ws, 'L' + r, L_, { fmt: TOT0 }); W(ws, 'M' + r, M_); W(ws, 'N' + r, { f: 'M' + r, r: N_ }, { fmt: ACC2 });
      F('O', 'N' + r + '*0.5%', O_); F('P', 'O' + r + '*18%', P_); F('Q', 'SUM(O' + r + ':P' + r + ')', Q_);
      W(ws, 'R' + r, R_); F('S', 'R' + r + '*0.3', S_); F('T', 'S' + r + '*18%', T_); F('U', 'S' + r + '+T' + r, U_);
      W(ws, 'V' + r, V_); W(ws, 'W' + r, W_);
      F('X', 'D' + r + '+G' + r + '+M' + r + '+Q' + r + '+U' + r + '-W' + r + '+K' + r, X_);
      if (Y_ !== null) W(ws, 'Y' + r, Y_, { fmt: '#,##0.00' });
      if (row.bankDate) W(ws, 'Z' + r, toDate(row.bankDate.slice(0, 10)), { fmt: 'd-mmm-yy', al: CENTER });
      F('AA', 'X' + r + '-Y' + r, X_ - n0(Y_)); W(ws, 'AB' + r, { f: '+X' + r + '-Y' + r, r: X_ - n0(Y_) }, { fmt: ACC2, border: BORD });
      Object.entries({ C: C_, D: D_, E: E_, F: F_, G: G_, H: H_, I: I_, J: J_, K: K_, L: L_, M: M_, N: N_, O: O_, P: P_, Q: Q_, R: R_, S: S_, T: T_, U: U_, V: V_, W: W_, X: X_, Y: n0(Y_), AA: X_ - n0(Y_) }).forEach(([L, v]) => add(L, v));
    }
  }
  const sumCols = [...'CDEFGHIJKLMNOPQRSTUVWXY'.split(''), 'AA'];
  [3, 6].forEach(rr => sumCols.forEach(L => W(ws, L + rr, { f: 'SUBTOTAL(9,' + L + first + ':' + L + last + ')', r: tot[L] || 0 }, { font: CAL })));
  W(ws, 'N1', { f: 'D6+N6', r: (tot.D || 0) + (tot.N || 0) }, { font: CAL });
  ws.getColumn('AB').width = 10.4;
  accum.nfs = accum.nfs || []; accum.nfs.push({ ym, name, N1: (tot.D || 0) + (tot.N || 0), E3: tot.E || 0, O3: tot.O || 0, S3: tot.S || 0, W3: tot.W || 0, Y3: tot.Y || 0 });
  return ws;
}
function nfsWorking(wb, accum) {
  const ws = wb.addWorksheet('NFS Workingsheet'); ws.getColumn('A').width = 37.4; ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 2, zoomScale: 115 }];
  const H1 = { font: { name: 'Calibri', size: 14, bold: true }, fill: fill('FFF8CBAD') };
  W(ws, 'A1', 'Customer wallet effect', H1); W(ws, 'A4', 'Bank level', H1);
  const rows = { 5: 'Presentment as per NTSL', 6: 'Adjustment', 7: 'Issuer WDL Fee-ATM', 8: 'Issuer WDL Fee-Micro ATM', 9: 'Issuer WDL Fee-Balance Inquiry', 10: 'Switching Fees', 11: 'GST', 12: 'Amount to be settled', 13: 'Opening Balance',
    14: 'Net Settlement', 15: 'Amount Actually deducted from bank', 16: 'Testing Transaction', 17: 'Balance settlements (2)', 19: 'Net balance in NFS settlement (1+2)', 20: 'Opening Balance', 21: 'Closing balance', 23: 'Amount deducted from Bank', 24: 'Date', 29: 'Difference' };
  Object.entries(rows).forEach(([r, t]) => W(ws, 'A' + r, t, { font: +r === 19 ? { name: 'Calibri', size: 12, bold: true } : CAL, border: BORD, fill: +r === 19 ? fill('FFA9CE91') : undefined }));
  let prevBal = 0;
  accum.nfs.forEach((s, i) => {
    const L = colL(2 + i), sh = q(s.name); ws.getColumn(L).width = 15.6;
    W(ws, L + '2', toDate(s.ym + '-01'), { font: CALB, fill: fill('FFFFE699'), border: BORD, fmt: 'mmm-yy' });
    const adj = -s.W3, gst = r2((s.E3 + s.O3 + 0 + s.S3) * 0.18), toSettle = r2(s.N1 + adj + s.E3 + s.O3 + 0 + s.S3 + gst), net = toSettle + prevBal, ded = s.Y3, bal = net - ded, closing = toSettle - ded;
    const c = (r, v, fm, bg, bold) => W(ws, L + r, v, { font: bold ? CALB : CAL, fmt: fm || TOT2, border: BORD, fill: bg ? fill(bg) : undefined });
    c(5, { f: sh + '!N1', r: s.N1 }); c(6, { f: '-' + sh + '!W3', r: adj }); /* net disputes: a credit adjustment reduces the settlement */ c(7, { f: sh + '!E3', r: s.E3 }); c(8, { f: sh + '!O3', r: s.O3 }); c(9, 0); c(10, { f: sh + '!S3', r: s.S3 });
    c(11, { f: '(' + L + '7+' + L + '8+' + L + '9+' + L + '10)*0.18', r: gst }); c(12, { f: 'SUM(' + L + '5:' + L + '11)', r: toSettle });
    c(13, i === 0 ? 0 : { f: colL(1 + i) + '19', r: prevBal }); c(14, { f: L + '12+' + L + '13', r: net }, null, 'FF00B0F0');
    c(15, { f: sh + '!Y3', r: ded }, null, 'FF00B0F0'); c(16, null); c(17, { f: L + '14-' + L + '15', r: bal }, null, 'FF00B0F0', true);
    c(19, { f: L + '12-' + L + '15', r: closing }); s.balCell = L + '17'; prevBal = closing; s.close = closing; s.bal = bal; s.net = net; s.ded = ded; s.toSettle = toSettle;
  });
  return ws;
}

/* ═══════════════════════════ RuPay ════════════════════════════════════════ */
const RU_WIDTH = { A: 12, B: 12, C: 6, D: 13, E: 7, F: 14, G: 7, H: 14, I: 10.4, J: 11, K: 6, L: 11, M: 6, N: 11, O: 11, P: 11, Q: 9, R: 9, S: 9, T: 9, U: 10, V: 11, W: 11, X: 9, Y: 9, Z: 11, AA: 11, AB: 13, AC: 11, AD: 11, AE: 11, AF: 11, AG: 13, AH: 11, AI: 13, AJ: 13, AK: 13, AL: 11, AM: 0.4, AN: 0.4, AO: 12, AP: 6, AQ: 6, AR: 13, AS: 6, AT: 13, AU: 6, AV: 10, AW: 6, AX: 10, AY: 6, AZ: 10, BA: 0.4, BB: 0.4, BC: 0.4, BD: 9, BE: 9, BF: 9, BG: 9, BH: 9 };
function rupay(wb, P, ym, accum) {
  const y = +ym.slice(0, 4), m = +ym.slice(5, 7), nd = daysIn(y, m), name = 'RuPay Control Sheet_' + monthLabel(ym);
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', xSplit: 4, ySplit: 7 }] });
  Object.entries(RU_WIDTH).forEach(([c, w]) => { ws.getColumn(c).width = w; });
  ws.getRow(5).height = 40; ws.getRow(6).height = 24; ws.getRow(7).height = 15;
  const FH = fill('FFBDD7EE'), FS = fill('FFDDEBF7'), HB = { font: AR9B, al: CENTER, border: BORD };
  const h5 = (a, b, v) => { if (a !== b) ws.mergeCells(a + ':' + b); W(ws, a, v, Object.assign({ fill: FH }, HB)); for (let k = colN(a.replace(/\d/g, '')); k <= colN(b.replace(/\d/g, '')); k++) { const cell = ws.getCell(5, k); cell.fill = FH; cell.border = BORD; } };
  [['A', 'Date'], ['B', 'Date'], ['C', 'Cycle'], ['D', 'Concat'], ['AG', 'Gross Total'], ['AH', 'GST'], ['AI', 'Net Settlememt'], ['AJ', 'Actual amount settled from Bank'], ['AK', 'Date of Settlement in Bank'], ['AL', 'Difference if any'], ['AO', 'Date'], ['AP', 'Cycle']].forEach(([c, t]) => {
    ws.mergeCells(c + '5:' + c + '7'); W(ws, c + '5', t, Object.assign({ fill: FH }, HB)); for (let r = 6; r <= 7; r++) { ws.getCell(c + r).border = BORD; ws.getCell(c + r).fill = FH; } });
  h5('E5', 'N5', 'DMS presentment received'); h5('O5', 'Q5', 'NPCI Assessment Fees'); h5('R5', 'T5', 'NPCI Processing Fees'); h5('U5', 'U5', 'Interchange Fees'); h5('V5', 'Y5', 'Interchange Income'); h5('Z5', 'AB5', 'Refund');
  h5('AC5', 'AD5', 'Chargeback Raised'); h5('AE5', 'AF5', 'Re-Presentment Raised'); h5('AQ5', 'AZ5', 'DMS open items (Authorise)'); h5('BD5', 'BH5', 'NPCI Processing Fees');
  [['E6', 'F6', 'POS'], ['G6', 'H6', 'ECOM'], ['I6', 'J6', 'Cash at POS'], ['K6', 'L6', 'POS Gift Card'], ['M6', 'N6', 'ECOM Gift Card'], ['AQ6', 'AR6', 'POS'], ['AS6', 'AT6', 'ECOM'], ['AU6', 'AV6', 'Cash at POS'], ['AW6', 'AX6', 'POS - GC'], ['AY6', 'AZ6', 'ECOM-GC']]
    .forEach(([a, b, t]) => { ws.mergeCells(a + ':' + b); W(ws, a, t, Object.assign({ fill: FS }, HB)); ws.getCell(b).border = BORD; ws.getCell(b).fill = FS; });
  [['O6', 'POS'], ['P6', 'ECOM'], ['Q6', 'Cash at POS'], ['R6', 'POS'], ['S6', 'ECOM'], ['T6', 'Cash at POS'], ['U6', 'Cash at POS'], ['V6', 'POS'], ['W6', 'ECOM'], ['X6', 'POS - GC'], ['Y6', 'ECOM - GC'], ['Z6', 'POS'], ['AA6', 'ECOM'], ['AB6', 'qSPARC Load  Cash'],
    ['AC6', 'POS'], ['AD6', 'ECOM'], ['AE6', 'POS'], ['AF6', 'ECOM'], ['BD6', 'POS'], ['BE6', 'ECOM'], ['BF6', 'Cash at POS'], ['BG6', 'POS GC'], ['BH6', 'ECOM GC']].forEach(([a, t]) => W(ws, a, t, Object.assign({ fill: FS }, HB)));
  ['Count', 'Amount', 'Count', 'Amount', 'Count', 'Amount', 'Count', 'Amount', 'Count', 'Amount'].forEach((t, i) => { W(ws, colL(5 + i) + '7', t, Object.assign({ fill: FS }, HB)); W(ws, colL(43 + i) + '7', t, Object.assign({ fill: FS }, HB)); });
  [15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 56, 57, 58, 59, 60].forEach(c => W(ws, colL(c) + '7', 'Amount', Object.assign({ fill: FS }, HB)));
  const first = 8, last = first + nd * 4 - 1, rm = rowMap(P), tot = {};
  const NUMC = []; for (let c = 5; c <= 60; c++) NUMC.push(c);
  const sumAdd = (c, v) => { tot[c] = (tot[c] || 0) + n0(v); };
  for (let d = 1; d <= nd; d++) {
    const iso = isoOf(y, m, d), r0 = first + (d - 1) * 4, ser = serial(iso);
    ws.mergeCells('A' + r0 + ':A' + (r0 + 3)); W(ws, 'A' + r0, toDate(iso), { font: AR9, fmt: 'dd\\-mm\\-yyyy', al: { horizontal: 'center', vertical: 'middle' }, border: BORD });
    ws.mergeCells('AO' + r0 + ':AO' + (r0 + 3)); W(ws, 'AO' + r0, toDate(iso), { font: AR9, fmt: 'dd\\-mm\\-yyyy', al: { horizontal: 'center', vertical: 'middle' }, border: BORD });
    for (let c = 1; c <= 4; c++) {
      const r = r0 + c - 1, row = rm.get(iso + '|C' + c), dd = row && row.src && row.src.data;
      for (let k = 1; k <= 60; k++) { if (k === 39 || k === 40 || (k >= 53 && k <= 55)) continue; const cell = ws.getCell(r, k); cell.font = AR9; cell.border = BORD; if (k >= 5 && k !== 37) cell.numFmt = '#,##0.00'; }
      W(ws, 'B' + r, toDate(iso), { fmt: 'dd\\-mm\\-yyyy' }); W(ws, 'C' + r, 'C' + c); W(ws, 'D' + r, ser + 'C' + c); W(ws, 'AP' + r, 'C' + c);
      if (!dd) { ws.getCell('C' + r).fill = fill('FFFFC7CE'); continue; }
      const v = {
        5: dd.pos_count, 6: dd.pos_amt, 7: dd.ecom_count, 8: dd.ecom_amt, 9: dd.cash_pos_count, 10: dd.cash_pos_amt, 11: Math.round(dd.pos_gc_count), 12: dd.pos_gc_amt, 13: dd.ecom_gc_count, 14: dd.ecom_gc_amt,
        15: dd.npci_assess_pos + dd.npci_assess_pos_gc, 16: dd.npci_assess_ecom + dd.npci_assess_ecom_gc, 17: dd.npci_assess_cash, 18: dd.npci_proc_ecom, 19: dd.npci_proc_pos, 20: 0,                   // same column mapping as the RuPay tool
        21: dd.int_income_cash, 22: dd.int_income_pos, 23: dd.int_income_ecom, 24: dd.int_income_pos_gc, 25: dd.int_income_ecom_gc, 26: dd.refund_pos, 27: dd.refund_ecom, 28: dd.qsparc_settled,
        29: dd.chargeback_pos, 30: dd.chargeback_ecom, 31: dd.repres_pos, 32: dd.repres_ecom, 33: dd.gross, 34: -dd.gst, 35: dd.net_settlement,
        43: dd.dms_open_pos_cnt, 44: dd.dms_open_pos_amt, 45: dd.dms_open_ecom_cnt, 46: dd.dms_open_ecom_amt, 47: 0, 48: 0, 49: 0, 50: 0, 51: 0, 52: 0, 56: dd.open_npci_proc_ecom, 57: dd.open_npci_proc_pos, 58: 0, 59: 0, 60: 0 };
      Object.entries(v).forEach(([c2, val]) => { const isCnt = [5, 7, 9, 11, 13, 43, 45, 47, 49, 51].includes(+c2); W(ws, colL(+c2) + r, isCnt ? Math.round(val) : r2(val), isCnt ? { fmt: 'General' } : {}); sumAdd(+c2, isCnt ? Math.round(val) : r2(val)); });
      const bank = row.bank;                                                                           // signed: credit − debit
      if (bank !== null) { W(ws, 'AJ' + r, r2(bank)); sumAdd(36, bank); }
      W(ws, 'AK' + r, toDate((row.bankDate || E.addDaysISO(iso, 1)).slice(0, 10)), { fmt: 'dd\\-mm\\-yyyy' });
      const diff = r2(dd.net_settlement - n0(bank)); sumAdd(38, diff);
      W(ws, 'AL' + r, { f: 'AI' + r + '-AJ' + r, r: diff }, { fill: fill(Math.abs(diff) < 1 ? 'FFC6EFCE' : 'FFFFC7CE') });
    }
  }
  // totals row 4 (SUBTOTAL) + the working formulas of rows 1-2, as in the finance team's sheet
  for (let c = 5; c <= 60; c++) { if (c === 37 || (c >= 39 && c <= 42) || (c >= 53 && c <= 55)) continue; const L = colL(c); W(ws, L + '4', { f: 'SUBTOTAL(9,' + L + first + ':' + L + last + ')', r: tot[c] || 0 }, { font: CAL, fmt: TOT2 }); }
  const T = L => tot[colN(L)] || 0;
  const row1 = [['I1', 'F4+H4+L4+N4+J4', T('F') + T('H') + T('L') + T('N') + T('J')], ['P1', 'SUM(O4:Q4)', T('O') + T('P') + T('Q')], ['S1', 'SUM(R4:T4)', T('R') + T('S') + T('T')], ['X1', 'SUM(V4:Y4)', T('V') + T('W') + T('X') + T('Y')],
    ['AA1', 'AA4+Z4+AB4', T('AA') + T('Z') + T('AB')], ['AD1', '+AC4+AD4', T('AC') + T('AD')], ['AF1', '+AE4+AF4', T('AE') + T('AF')]];
  row1.forEach(([a, f, r]) => W(ws, a, { f, r }, { font: CALB, fmt: TOT0 })); W(ws, 'AI2', { f: '+AI4', r: T('AI') }, { font: CALB, fmt: TOT2 });
  accum.rupay = accum.rupay || []; accum.rupay.push({ ym, name, I1: row1[0][2], P1: row1[1][2], S1: row1[2][2], X1: row1[3][2], AA1: row1[4][2], AD1: row1[5][2], AF1: row1[6][2], AJ4: T('AJ') });
  return ws;
}
function rupayWorking(wb, accum) {
  const ws = wb.addWorksheet('RuPay Working Sheet'); ws.getColumn('A').width = 37.4; ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];
  W(ws, 'A2', 'Customer wallet effect', { font: { name: 'Calibri', size: 14, bold: true }, fill: fill('FFF8CBAD') });
  const rows = { 4: 'Presentment as per DSR', 5: 'Refund', 6: 'Adjustment', 7: 'Net Wallet balance to be reduced', 8: 'NPCI charges', 9: 'Interchange Fees', 10: 'GST', 11: 'Interchange Income', 12: 'Rupay Chargeback Raised', 13: 'Representment received', 14: 'Amount to be settled', 15: 'Amount actually deducted from escrow', 16: 'Balance in Rupay settlements (2)' };
  Object.entries(rows).forEach(([r, t]) => W(ws, 'A' + r, t, { font: CAL, border: BORD }));
  accum.rupay.forEach((s, i) => {
    const L = colL(2 + i), sh = q(s.name); ws.getColumn(L).width = 16.3;
    W(ws, L + '3', toDate(s.ym + '-01'), { font: CALB, fill: fill('FFFFE699'), border: BORD, fmt: 'mmm-yy' });
    const npci = s.P1 + s.S1, gst = npci * 0.18, toSettle = s.I1 + npci + 0 + gst - s.X1 - s.AA1 - s.AD1 + s.AF1, ded = -s.AJ4, bal = toSettle - ded;
    const c = (r, v, bg, bold) => W(ws, L + r, v, { font: bold ? CALB : CAL, fmt: TOT2, border: BORD, fill: bg ? fill(bg) : undefined });
    c(4, { f: sh + '!I1', r: s.I1 }); c(5, { f: sh + '!AA1', r: s.AA1 }); c(6, null); c(7, { f: L + '4-' + L + '5-' + L + '6', r: s.I1 - s.AA1 }); c(8, { f: sh + '!P1+' + sh + '!S1', r: npci }); c(9, 0);
    c(10, { f: L + '8*0.18', r: gst }); c(11, { f: '+' + sh + '!X1', r: s.X1 }); c(12, { f: sh + '!AD1', r: s.AD1 }); c(13, { f: sh + '!AF1', r: s.AF1 });
    c(14, { f: L + '4+' + L + '8+' + L + '9+' + L + '10-' + L + '11-' + L + '5-' + L + '12+' + L + '13', r: toSettle }); c(15, { f: '-' + sh + '!AJ4', r: ded }); c(16, { f: L + '14-' + L + '15', r: bal }, 'FFC6EFCE', true);
    s.bal = bal; s.toSettle = toSettle; s.ded = ded;
  });
  return ws;
}

/* ═══════════════════════════ UPI ══════════════════════════════════════════ */
const UPI_REMARKS = ['Approved Fee', 'Approved Fee Gst', 'GST on Penalty', 'Net Adjusted Amount', 'Penalty', 'Switching Fee', 'Switching Fee Gst', 'TDS on GST', 'Transaction Amount'];
function upi(wb, P, accum) {
  const months = monthsOf(P), WK = 'UPI Working';
  // Working: every classified NTSL row, in date/cycle order
  const lines = [];
  P.rows.forEach(r => { if (r.src && r.src.lines) r.src.lines.forEach(l => lines.push({ date: r.date, cycle: r.cycle, ...l })); });
  lines.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : E.UPI_ORDER.indexOf(a.cycle) - E.UPI_ORDER.indexOf(b.cycle));
  // build sheets in the order Summary, Pivot(s), Working (the Working sheet is created last but referenced by name)
  const sum = wb.addWorksheet('UPI Summary'); sum.getColumn('A').width = 37.4; sum.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];
  W(sum, 'A2', 'Customer wallet effect', { font: { name: 'Calibri', size: 14, bold: true }, fill: fill('FFF8CBAD') });
  const labels = { 4: 'Presentment as per NTSL', 5: 'Net Wallet balance to be reduced', 6: 'NPCI charges', 7: 'Approver Fees - Expense', 8: 'GST - NPCI charges', 9: 'GST - Approver Fees', 10: 'Approver Fees - Income', 11: 'GST - Output', 12: 'Short & Excess', 13: 'Chargeback raised', 14: 'Chargeback reversed', 15: 'Amount to be settled', 16: 'Amount actually deducted from escrow', 17: 'Balance in Rupay settlements (2)' };
  Object.entries(labels).forEach(([r, t]) => W(sum, 'A' + r, t, { font: CAL, border: BORD }));
  const pivots = [];
  months.forEach((ym, idx) => {
    const y = +ym.slice(0, 4), m = +ym.slice(5, 7), lastMonth = idx === months.length - 1, pname = lastMonth ? 'UPI Pivot Final' : 'UPI Pivot_' + monthLabel(ym);
    const pv = wb.addWorksheet(pname); [18.6, 17, 17].forEach((w, i) => { pv.getColumn(i + 1).width = w; });
    const mine = lines.filter(l => l.date.slice(0, 7) === ym), agg = {};
    mine.forEach(l => { const a = agg[l.remarks] || (agg[l.remarks] = { d: 0, c: 0 }); a.d += l.debit; a.c += l.credit; });
    ['Row Labels', 'Sum of Debit (₹)', 'Sum of Credit (₹)'].forEach((h, i) => W(pv, colL(i + 1) + '1', h, { font: CALB, fill: fill('FFDDEBF7'), border: BORD }));
    const crit = 'SUMIFS(' + q(WK) + '!$@:$@,' + q(WK) + '!$H:$H,$A#,' + q(WK) + '!$A:$A,">="&DATE(' + y + ',' + m + ',1),' + q(WK) + '!$A:$A,"<"&DATE(' + y + ',' + (m + 1) + ',1))';
    let td = 0, tc = 0;
    UPI_REMARKS.forEach((rem, i) => {
      const r = 2 + i, a = agg[rem] || { d: 0, c: 0 }; td += a.d; tc += a.c;
      W(pv, 'A' + r, rem, { font: CAL, border: BORD });
      W(pv, 'B' + r, { f: crit.replace(/@/g, 'F').replace('#', r), r: r2(a.d) }, { font: CAL, fmt: '#,##0.00', border: BORD });
      W(pv, 'C' + r, { f: crit.replace(/@/g, 'G').replace('#', r), r: r2(a.c) }, { font: CAL, fmt: '#,##0.00', border: BORD });
    });
    const gt = 2 + UPI_REMARKS.length;
    W(pv, 'A' + gt, 'Grand Total', { font: CALB, border: BORD, fill: fill('FFDDEBF7') });
    W(pv, 'B' + gt, { f: 'SUM(B2:B' + (gt - 1) + ')', r: r2(td) }, { font: CALB, fmt: '#,##0.00', border: BORD, fill: fill('FFDDEBF7') });
    W(pv, 'C' + gt, { f: 'SUM(C2:C' + (gt - 1) + ')', r: r2(tc) }, { font: CALB, fmt: '#,##0.00', border: BORD, fill: fill('FFDDEBF7') });
    W(pv, 'B' + (gt + 2), { f: 'B' + gt + '-C' + gt, r: r2(td - tc) }, { font: CAL, fmt: '#,##0.00' });
    pivots.push({ ym, name: pname, agg });
    // Summary column for the month
    const L = colL(2 + idx); sum.getColumn(L).width = 16.7;
    W(sum, L + '3', toDate(ym + '-01'), { font: CALB, fill: fill('FFFFE699'), border: BORD, fmt: 'mmm-yy' });
    const rng = c => q(pname) + '!$' + c + '$2:$' + c + '$' + (gt - 1), lab = n => q(pname) + '!$A$2:$A$' + (gt - 1);
    const D = rem => '(SUMIF(' + lab() + ',"' + rem + '",' + rng('B') + '))', C = rem => '(SUMIF(' + lab() + ',"' + rem + '",' + rng('C') + '))';
    const d = rem => (agg[rem] || { d: 0 }).d, c = rem => (agg[rem] || { c: 0 }).c;
    const v4 = d('Transaction Amount') - c('Transaction Amount'), v6 = d('Switching Fee'), v7 = d('Approved Fee'), v8 = v6 * 0.18, v9 = v7 * 0.18, v10 = c('Approved Fee'), v11 = v10 * 0.18;
    const v12 = c('GST on Penalty') + c('Penalty') + c('TDS on GST') + c('Switching Fee') - d('TDS on GST') - d('Penalty') - d('GST on Penalty');
    const v13 = d('Net Adjusted Amount'), v14 = c('Net Adjusted Amount'), v15 = v4 + v6 + v7 + v8 + v9 - v10 - v11 - v12 + v13 - v14;
    const bankRows = P.rows.filter(r => r.date.slice(0, 7) === ym && r.bank !== null), ded = r2(-bankRows.reduce((s, r) => s + r.bank, 0));
    const cc = (r, v, bold) => W(sum, L + r, v, { font: bold ? CALB : CAL, fmt: TOT2, border: BORD });
    cc(4, { f: D('Transaction Amount') + '-' + C('Transaction Amount'), r: v4 }); cc(5, { f: 'SUM(' + L + '4)', r: v4 }); cc(6, { f: D('Switching Fee'), r: v6 }); cc(7, { f: D('Approved Fee'), r: v7 });
    cc(8, { f: '(' + L + '6)*0.18', r: v8 }); cc(9, { f: '(' + L + '7)*0.18', r: v9 }); cc(10, { f: C('Approved Fee'), r: v10 }); cc(11, { f: L + '10*0.18', r: v11 });
    cc(12, { f: C('GST on Penalty') + '+' + C('Penalty') + '+' + C('TDS on GST') + '+' + C('Switching Fee') + '-' + D('TDS on GST') + '-' + D('Penalty') + '-' + D('GST on Penalty'), r: v12 });
    cc(13, { f: D('Net Adjusted Amount'), r: v13 }); cc(14, { f: C('Net Adjusted Amount'), r: v14 });
    cc(15, { f: L + '4+' + L + '6+' + L + '7+' + L + '8+' + L + '9-' + L + '10-' + L + '11-' + L + '12+' + L + '13-' + L + '14', r: v15 });
    cc(16, ded); cc(17, { f: L + '15-' + L + '16', r: v15 - ded }, true);
    sum.getCell(L + '17').fill = fill('FFC6EFCE'); (accum.upi = accum.upi || []).push({ ym, L, bal: v15 - ded, toSettle: v15, ded });
  });
  // Working sheet
  const wk = wb.addWorksheet(WK); [11, 7, 14, 78, 10, 16, 16, 20, 16].forEach((w, i) => { wk.getColumn(i + 1).width = w; });
  ['Date', 'Cycle', 'Unique Code', 'Description', 'No of Txns', 'Debit', 'Credit', 'Remarks', 'Amount'].forEach((h, i) => W(wk, colL(i + 1) + '1', h, { font: CALB, fill: fill('FFDDEBF7'), border: BORD }));
  wk.views = [{ state: 'frozen', ySplit: 1 }];
  lines.forEach((l, i) => {
    const r = i + 2, row = wk.getRow(r), ser = serial(l.date);
    row.getCell(1).value = toDate(l.date); row.getCell(1).numFmt = 'dd-mm-yyyy'; row.getCell(2).value = l.cycle;
    row.getCell(3).value = { formula: 'A' + r + '&B' + r, result: ser + l.cycle }; row.getCell(4).value = l.desc; row.getCell(5).value = l.txns; row.getCell(6).value = l.debit; row.getCell(7).value = l.credit; row.getCell(8).value = l.remarks;
    row.getCell(9).value = { formula: 'F' + r + '-G' + r, result: r2(l.debit - l.credit) };
  });
  wk.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(2, lines.length + 1), column: 9 } };
  return wk;
}

/* ═══════════════════════════ summary sheet + driver ══════════════════════ */
function controlSummary(ws, accum, R) {
  [34, 14, 26, 20, 20, 18].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  W(ws, 'A1', 'CONTROL SHEET DOWNLOAD', { font: { name: 'Calibri', size: 14, bold: true } });
  W(ws, 'A2', 'Generated ' + new Date().toLocaleString('en-IN') + ' · bank statement ' + R.bank.file, { font: { name: 'Calibri', size: 10, italic: true } });
  ['Product / month', 'Period', 'Working sheet', 'Amount to be settled', 'Deducted from bank', 'Balance (should be ≈ 0)'].forEach((h, i) => W(ws, colL(i + 1) + '4', h, { font: { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } }, fill: fill('FF0F2540'), border: BORD, al: CENTER }));
  let r = 5; const line = (label, ym, sheet, fSet, rSet, fDed, rDed, fBal, rBal) => {
    W(ws, 'A' + r, label, { font: CAL, border: BORD }); W(ws, 'B' + r, toDate(ym + '-01'), { font: CAL, border: BORD, fmt: 'mmm-yy' }); W(ws, 'C' + r, sheet, { font: CAL, border: BORD });
    W(ws, 'D' + r, { f: fSet, r: rSet }, { font: CAL, border: BORD, fmt: TOT2 }); W(ws, 'E' + r, { f: fDed, r: rDed }, { font: CAL, border: BORD, fmt: TOT2 }); W(ws, 'F' + r, { f: fBal, r: rBal }, { font: CALB, border: BORD, fmt: TOT2, fill: fill(Math.abs(rBal) < 5 ? 'FFC6EFCE' : 'FFFFC7CE') }); r++;
  };
  (accum.fastag || []).forEach(s => { const e = q('FASTag Accounting Entry'), L = s.entryCol; line('FASTag (NETC) — accounting entry', s.ym, 'FASTag Accounting Entry', e + '!' + L + '3', s.E, e + '!' + L + '4+' + e + '!' + L + '5', s.I + s.F + s.G, e + '!' + L + '10', s.E - s.I - (s.F + s.G)); });
  (accum.nfs || []).forEach((s, i) => { const w = q('NFS Workingsheet'), L = colL(2 + i); line('NFS', s.ym, 'NFS Workingsheet', w + '!' + L + '12', s.toSettle, w + '!' + L + '15', s.ded, w + '!' + L + '17', s.bal); });
  (accum.rupay || []).forEach((s, i) => { const w = q('RuPay Working Sheet'), L = colL(2 + i); line('RuPay', s.ym, 'RuPay Working Sheet', w + '!' + L + '14', s.toSettle, w + '!' + L + '15', s.ded, w + '!' + L + '16', s.bal); });
  (accum.upi || []).forEach(s => { const w = q('UPI Summary'); line('UPI', s.ym, 'UPI Summary', w + '!' + s.L + '15', s.toSettle, w + '!' + s.L + '16', s.ded, w + '!' + s.L + '17', s.bal); });
  r += 1;
  [['Notes', ''], ['•', 'Source columns come from the DSR/NTSL files; the bank column from the matched bank statement lines. Calculation columns are live Excel formulas.'],
   ['•', 'FASTag: "Amount" is the DSR final net payable plus the service fee, so Net Settlement ties to the DSR; columns M–N show the DSR toll/parking amount and the other items.'],
   ['•', 'A cycle with no source file is marked red in its Cycle cell; a cycle with no bank line has an empty bank column (its Difference shows the open amount).'],
   ['•', 'The balance column is the control: it should be ≈ 0 (paise rounding), exactly like the "Balance" rows of your working sheets.']].forEach(([a, b], i) => { W(ws, 'A' + (r + i), a, { font: i ? CAL : CALB }); if (b) W(ws, 'B' + (r + i), b, { font: CAL }); });
  return ws;
}

async function build(R) {
  const wb = new ExcelJS.Workbook(); wb.creator = 'Recon OS'; wb.created = new Date();
  const accum = {}, P = R.products;
  const cs = wb.addWorksheet('Control Summary', { views: [{ showGridLines: false }] });                  // first sheet; filled at the end
  if (P.NETC) { monthsOf(P.NETC).forEach(ym => fastag(wb, P.NETC, ym, accum)); fastagEntry(wb, accum); }
  if (P.NFS) { monthsOf(P.NFS).forEach(ym => nfs(wb, P.NFS, ym, accum)); nfsWorking(wb, accum); }
  if (P.RUPAY) { monthsOf(P.RUPAY).forEach(ym => rupay(wb, P.RUPAY, ym, accum)); rupayWorking(wb, accum); }
  if (P.UPI) upi(wb, P.UPI, accum);
  controlSummary(cs, accum, R);
  wb.calcProperties = { fullCalcOnLoad: true };
  return wb.xlsx.writeBuffer();
}
async function download(R) {
  const buf = await build(R), d = new Date(), p = n => String(n).padStart(2, '0');
  const per = E.PRODUCTS.filter(x => R.products[x]).map(x => R.products[x].period.start).sort()[0] || '';
  saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'Control_Sheets_' + (per ? per.slice(0, 7) : 'report') + '_' + p(d.getDate()) + p(d.getMonth() + 1) + String(d.getFullYear()).slice(2) + '.xlsx');
}
window.OSControlSheet = { build, download };
})();
