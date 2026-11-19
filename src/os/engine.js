/* ════════════════════════════════════════════════════════════════════════════
   RECON OS — engine (no DOM). One bank statement + any mix of settlement files/folders/zips
   → every reconciliation at once.

   Product rules are the finalised per-tool rules (FASTag/NETC, NFS, RuPay, UPI, Add Money macro v3),
   ported function-for-function. What is new here is the INPUT side: auto-detecting file types,
   unpacking nested zips, reading one combined bank statement and splitting it by product, and
   classifying bank lines that belong to another month.

   Dates are ISO strings 'YYYY-MM-DD' everywhere inside the engine. Amounts use the bank-perspective
   sign: credit to the account = +, debit from the account = −.
   Needs globals: XLSX (SheetJS) and, for zips, JSZip.
   ════════════════════════════════════════════════════════════════════════════ */
(function (root) {
'use strict';

const X = () => root.XLSX;
const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
const r4 = n => Math.round((n + Number.EPSILON) * 10000) / 10000;
const sf = v => { const f = parseFloat(v); return isNaN(f) ? 0 : f; };
const isBlank = v => v === null || v === undefined || v === '';
const pad2 = n => String(n).padStart(2, '0');
const PRODUCTS = ['NETC', 'NFS', 'RUPAY', 'UPI'];
const PRODUCT_LABEL = { NETC: 'FASTag (NETC)', NFS: 'NFS', RUPAY: 'RuPay', UPI: 'UPI', ADDMONEY: 'Add money to Pool' };
const DEFAULT_TOL = { NETC: 0.01, NFS: 1, RUPAY: 1, UPI: 1, ADDMONEY: 0.01 };
const INCLUSIVE = { NFS: true };                       // NFS tool: |diff| <= threshold; all others: |diff| < tolerance

/* ── dates ─────────────────────────────────────────────────────────────── */
const isoOf = (y, m, d) => y + '-' + pad2(m) + '-' + pad2(d);
function toISO(v) {                                   // Date | 'DD-MM-YYYY' | 'DD/MM/YYYY' | 'YYYY-MM-DD' | 'DD-MM-YYYY hh:mm:ss'
  if (isBlank(v)) return null;
  if (v instanceof Date) return isNaN(v) ? null : isoOf(v.getFullYear(), v.getMonth() + 1, v.getDate());
  if (typeof v === 'number') {                        // excel serial
    if (v > 20000 && v < 80000) { const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000); return isoOf(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
    return null;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/); if (m) return isoOf(+m[3], +m[2], +m[1]);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return isoOf(+m[1], +m[2], +m[3]);
  return null;
}
const addDaysISO = (iso, n) => { const [y, m, d] = iso.split('-').map(Number); const t = new Date(Date.UTC(y, m - 1, d + n)); return isoOf(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); };
const diffDays = (a, b) => Math.round((Date.parse(a + 'T00:00:00Z') - Date.parse(b + 'T00:00:00Z')) / 86400000);
const fmtDMY = iso => iso ? iso.slice(8, 10) + '-' + iso.slice(5, 7) + '-' + iso.slice(0, 4) : '';
const monthOf = iso => iso ? iso.slice(0, 7) : '';
const ddmmyy = s => '20' + s.slice(4, 6) + '-' + s.slice(2, 4) + '-' + s.slice(0, 2);

/* ── amounts (Indian grouping, trailing CR/DR) ─────────────────────────── */
function parseAmount(v) {
  if (isBlank(v)) return 0;
  if (typeof v === 'number') return v;
  let s = String(v).trim(); if (!s) return 0;
  let sign = 1;
  if (/\bDR\b\s*$/i.test(s)) sign = -1;
  s = s.replace(/\b(CR|DR)\b\s*$/i, '').replace(/[₹,\s]/g, '');
  if (/^\(.*\)$/.test(s)) { sign = -sign; s = s.slice(1, -1); }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : sign * n;
}

/* ── workbook reading ──────────────────────────────────────────────────── */
const toU8 = b => b instanceof Uint8Array ? b : new Uint8Array(b);
function readWB(buf, opt) { return X().read(toU8(buf), Object.assign({ type: 'array', raw: true }, opt || {})); }
function sheetRows(wb, name) {
  const ws = wb.Sheets[name || wb.SheetNames[0]];
  if (!ws || !ws['!ref']) return [];
  const rg = X().utils.decode_range(ws['!ref']);
  return X().utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: true, range: { s: { r: 0, c: 0 }, e: rg.e } });
}
function decodeText(u8) { return new TextDecoder('utf-8', { fatal: false }).decode(u8); }
function magic(u8) {
  if (u8[0] === 0x50 && u8[1] === 0x4B) return 'zip';                 // xlsx / zip
  if (u8[0] === 0xD0 && u8[1] === 0xCF) return 'ole';                 // legacy xls
  return 'text';
}

/* ═════════════════════════ BANK STATEMENT ═════════════════════════════ */
function detectBankHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 80); i++) {
    const r = (rows[i] || []).map(c => isBlank(c) ? '' : String(c).trim().toLowerCase());
    const narr = r.findIndex(c => /^(transaction details|narrative|description|particulars|transaction remarks)$/.test(c));
    const deb = r.findIndex(c => /^(debit|withdrawal|withdrawals|debit amount)$/.test(c));
    const cre = r.findIndex(c => /^(credit|deposit|deposits|credit amount)$/.test(c));
    if (narr >= 0 && deb >= 0 && cre >= 0) {
      const dt = r.findIndex(c => /^(trans(action)?\.? date( and time)?|txn date|tran date|date)$/.test(c));
      const val = r.findIndex(c => /^(value date|payment date|post(ing)? date)$/.test(c));
      const ref = r.findIndex(c => /(ref|cheque|chq)/.test(c));
      const bal = r.findIndex(c => /balance/.test(c));
      return { row: i, narr, deb, cre, dt, val, ref, bal };
    }
  }
  return null;
}
function parseBalanceText(s) { return parseAmount(s); }
function parseBankStatement(rows, fileName) {
  const h = detectBankHeader(rows);
  if (!h) return null;
  const lines = []; let skipped = 0;
  for (let i = h.row + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const first = isBlank(r[0]) ? '' : String(r[0]);
    if (/end of the statement/i.test(first) || /end of the statement/i.test(String(r[h.narr] || ''))) break;
    const dtIso = toISO(h.dt >= 0 ? r[h.dt] : null), valIso = toISO(h.val >= 0 ? r[h.val] : null);
    const narr = isBlank(r[h.narr]) ? '' : String(r[h.narr]).trim();
    if (!dtIso && !valIso) { if (narr || first) skipped++; continue; }            // opening-balance row, notes, blanks
    if (/^opening balance$/i.test(narr)) continue;
    const deb = parseAmount(r[h.deb]), cre = parseAmount(r[h.cre]);
    lines.push({ row: i + 1, date: dtIso || valIso, valueDate: valIso || dtIso, narr, ref: h.ref >= 0 && !isBlank(r[h.ref]) ? String(r[h.ref]) : '',
                 debit: r2(deb), credit: r2(cre), balance: h.bal >= 0 ? r[h.bal] : null });
  }
  // statement meta from the preamble
  const meta = { opening: null, closing: null, totalDebits: null, totalCredits: null, periodStart: null, periodEnd: null, account: '', accountName: '' };
  for (let i = 0; i < Math.min(h.row, 60); i++) {
    const r = (rows[i] || []).map(c => isBlank(c) ? '' : String(c).trim());
    const low = r.map(c => c.toLowerCase());
    if (low[0] === 'opening balance' && low.includes('closing balance')) {
      const v = rows[i + 1] || []; const ci = k => low.indexOf(k);
      meta.opening = parseBalanceText(v[ci('opening balance')]); meta.closing = parseBalanceText(v[ci('closing balance')]);
      meta.totalDebits = parseAmount(v[ci('total debits')]); meta.totalCredits = parseAmount(v[ci('total credits')]);
    }
    if (low[0] === 'statement period') { const m = String(r[1]).match(/(\d{1,2}\/\d{1,2}\/\d{4})\s*TO\s*(\d{1,2}\/\d{1,2}\/\d{4})/i); if (m) { meta.periodStart = toISO(m[1]); meta.periodEnd = toISO(m[2]); } if (low[2] === 'account name') meta.accountName = r[3]; }
    if (low[0] === 'account no') meta.account = r[1];
  }
  const dates = lines.map(l => l.date).sort();
  meta.firstDate = dates[0] || null; meta.lastDate = dates[dates.length - 1] || null;
  return { lines, meta, fileName, skipped };
}

/* ── bank narrative → product ──────────────────────────────────────────── */
// UPI narrative patterns (same five as the UPI tool)
const CYCLE_MAP = { '01': '1C', '02': '2C', '03': '3C', '04': '4C', '05': '5C', '06': '6C', '07': '7C', '08': '8C', '09': '9C', '10': '10C' };
function parseUpiNarrative(raw) {
  const n = String(raw || '').trim();
  let m = n.match(/Eroute UPI Setl\s+(\S+)\s+(\d{6})/i);
  if (m) { const iso = ddmmyy(m[2]); if (toISO(iso)) return { date: iso, cycle: m[1].toUpperCase().replace('DC01', 'DC1').replace('DC02', 'DC2') }; }
  m = n.match(/DC0([12])\s+UPI EROUTE TECH PVT LTD_NPCI/i);
  if (m) { const dm = n.match(/NPCI00(\d{4})(\d{2})(\d{2})/i); if (dm) return { date: isoOf(+dm[1], +dm[2], +dm[3]), cycle: 'DC' + m[1] }; }
  m = n.match(/NPCI00(\d{4})(\d{2})(\d{2})(\d{2})$/i);
  if (m) return { date: isoOf(+m[1], +m[2], +m[3]), cycle: CYCLE_MAP[m[4]] || m[4] };
  const mDC = n.match(/DC([12])/i), mDt = n.match(/NPCI00(\d{4})(\d{2})(\d{2})/i);
  if (mDC && mDt) return { date: isoOf(+mDt[1], +mDt[2], +mDt[3]), cycle: 'DC' + mDC[1] };
  return null;
}
function classifyBankLine(l) {
  const n = l.narr.trim(), u = n.toUpperCase();
  let m;
  if (/^NETC\b/.test(u) && (m = n.match(/_00(\d{4})(\d{2})(\d{2})(\d{2})\s*$/))) return { product: 'NETC', date: isoOf(+m[1], +m[2], +m[3]), cycle: String(+m[4]) };
  if (/\bNFS\b/.test(u) && (m = n.match(/_00(\d{4})(\d{2})(\d{2})(\d{2})\s*$/))) return { product: 'NFS', date: isoOf(+m[1], +m[2], +m[3]), cycle: 'C' + (+m[4]) };
  if (/\bRUPAY\b/.test(u) && !/REV_/.test(u) && (m = n.match(/_00(\d{4})(\d{2})(\d{2})(\d{2})\s*$/))) return { product: 'RUPAY', date: isoOf(+m[1], +m[2], +m[3]), cycle: 'C' + (+m[4]) };
  const up = parseUpiNarrative(n); if (up) return { product: 'UPI', date: up.date, cycle: up.cycle };
  return null;
}
function bankFamily(l) {
  const u = l.narr.toUpperCase();
  const f = /^IMPS-CIB/.test(u) ? 'IMPS-CIB (payout)' : /^IMPS/.test(u) ? 'IMPS' : /^(INEFT|NEFT)/.test(u) ? 'NEFT' : /^(IRTGS|RTGS)/.test(u) ? 'RTGS'
    : /^IFT/.test(u) ? 'IFT (internal transfer)' : u ? 'Other' : '(blank)';
  return f + (l.credit ? ' – credit' : l.debit ? ' – debit' : '');
}

/* ═════════════════════════ SOURCE FILE PARSERS ════════════════════════ */

/* ── NETC / FASTag DSR (parseSingleDSR) ─────────────────────────────────── */
function parseNetcDsr(rows, filename) {
  let hdrIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 5); i++) if (rows[i] && rows[i][0] === 'Settlement Date') { hdrIdx = i; break; }
  if (hdrIdx === -1) return null;
  let blockDate = null;
  for (let i = hdrIdx + 1; i < rows.length; i++) { const v = rows[i] && rows[i][0]; if (v && v !== 'Settlement Date' && String(v).trim()) { blockDate = toISO(v); break; } }
  if (!blockDate) return null;
  let grandTotal = null, gstAmt = 0, txnCount = 0, txnAmt = 0;
  for (let k = hdrIdx + 1; k < rows.length; k++) {
    const r = rows[k]; if (!r) continue;
    if (String(r[5] || '').trim() === 'INWARD GST') gstAmt = parseFloat(r[18] || 0);
    if (r[1] === 'Total' && r[4] === 'Total') grandTotal = r;
    if (r[8] === 'DEBIT' && (r[9] === 'TOLL' || r[9] === 'PARKING')) { txnCount += (parseInt(r[10]) || 0); txnAmt += (parseFloat(r[12]) || 0); }
  }
  if (!grandTotal) return null;
  const svcTotal = parseFloat(grandTotal[18] || 0);
  const cm = String(filename || '').match(/ISSUER[_ ]\d{4}-\d{2}-\d{2}-(\d+)/i);
  return { date: blockDate, cycleFromName: cm ? +cm[1] : null, filename, cnt: txnCount, txnAmt: r2(txnAmt), taxable: r2(svcTotal - gstAmt),
           gst: r2(gstAmt), svcTotal: r2(svcTotal), net: r2(parseFloat(grandTotal[27] || 0)) };
}

/* ── RuPay DSR (parseDsr) ───────────────────────────────────────────────── */
const ECOM_GC_BIN = 818014, RU_PRES = ['Presentment (With Auth)', 'Offline Presentment(Without Auth)'], RU_NO_FEE = ['Balance Update'];
function parseRupayDsr(rows, filename) {
  const r = { pos_count: 0, pos_amt: 0, ecom_count: 0, ecom_amt: 0, pos_gc_count: 0, pos_gc_amt: 0, ecom_gc_count: 0, ecom_gc_amt: 0, cash_pos_count: 0, cash_pos_amt: 0,
    npci_assess_pos: 0, npci_assess_ecom: 0, npci_assess_pos_gc: 0, npci_assess_ecom_gc: 0, npci_assess_cash: 0, npci_proc_ecom: 0, npci_proc_pos: 0,
    int_income_cash: 0, int_income_pos: 0, int_income_ecom: 0, int_income_pos_gc: 0, int_income_ecom_gc: 0, refund_pos: 0, refund_ecom: 0, qsparc_settled: 0,
    chargeback_pos: 0, chargeback_ecom: 0, repres_pos: 0, repres_ecom: 0, gst: 0, final_net: 0 };
  let cB = null, cS = null, cC = null, found = false;
  for (let row of rows) {
    if (!row || row.length < 10) continue;
    row = row.slice(); while (row.length < 30) row.push(null);
    if (row.every(v => v === null)) continue;
    if (row[4] !== null && row[4] !== 'Total' && row[4] !== 'Acq ID / ISS Bin') cB = row[4];
    if (row[6] !== null && row[6] !== 'Status(Approved/Declined)') cS = row[6];
    if (row[7] !== null && row[7] !== 'Transaction Cycle') cC = row[7];
    const ch = row[9], cnt = sf(row[10]), txnCR = sf(row[13]), setDR = sf(row[15]), setCR = sf(row[16]), intCR = sf(row[18]), othDR = sf(row[23]), tt = row[8], fCR = sf(row[27]);
    if (row[1] === 'Total' && row[2] === 'Total' && row[3] === 'Total' && row[4] === 'Total' && row[5] === 'Total') { r.final_net = sf(row[29]); found = true; }
    if (row[5] === 'INWARD GST') r.gst = sf(row[25]);
    if (sf(cB) === ECOM_GC_BIN) {
      if (RU_PRES.includes(cC) && tt === 'Purchase' && cS !== 'D') {
        if (ch === 'ECOM' && setDR > 0) { r.ecom_gc_count += cnt; r.ecom_gc_amt += setDR; r.int_income_ecom_gc += intCR; r.npci_assess_ecom_gc += othDR; }
        else if (ch === 'POS' && setDR > 0) { r.pos_gc_count += cnt; r.pos_gc_amt += setDR; r.int_income_pos_gc += intCR; r.npci_assess_pos_gc += othDR; }
      }
      if (cC === 'Refund') { if (ch === 'ECOM') r.refund_ecom += setCR > 0 ? setCR : txnCR; else if (ch === 'POS') r.refund_pos += setCR > 0 ? setCR : txnCR; }
    }
    if (cC === 'Chargeback Raise' || cC === 'Chargeback Acceptance') { const a = Math.max(setCR, fCR); if (ch === 'POS' && a > 0) r.chargeback_pos += a; if (ch === 'ECOM' && a > 0) r.chargeback_ecom += a; }
    if (cC === 'Re-Presentment Raise') { if (ch === 'POS' && setDR > 0) r.repres_pos += setDR; if (ch === 'ECOM' && setDR > 0) r.repres_ecom += setDR; }
    if (cC === 'NPCI Fee Collection' && ch === 'POS' && othDR > 0) { r.cash_pos_count += cnt; r.cash_pos_amt += othDR; }
    if (sf(cB) !== ECOM_GC_BIN) {
      if (RU_PRES.includes(cC) && tt === 'Purchase' && cS !== 'D') {
        if (ch === 'POS' && setDR > 0) { r.pos_count += cnt; r.pos_amt += setDR; r.int_income_pos += intCR; r.npci_assess_pos += othDR; }
        else if (ch === 'ECOM' && setDR > 0) { r.ecom_count += cnt; r.ecom_amt += setDR; r.int_income_ecom += intCR; r.npci_assess_ecom += othDR; }
      }
      if (tt === 'qSPARC Money Load through Cash' && setCR > 0) r.qsparc_settled += setCR;
      if (cC === 'Refund') { if (ch === 'ECOM') r.refund_ecom += setCR > 0 ? setCR : txnCR; else if (ch === 'POS') r.refund_pos += setCR > 0 ? setCR : txnCR; }
    }
    if (cC === 'DMS Auth Transaction' && cS === 'A' && tt !== null) {
      const iq = tt && (tt.includes('qSPARC') || tt.includes('qSparc'));
      if (!RU_NO_FEE.includes(tt) && !iq && setCR === 0 && othDR > 0) {
        if (ch === 'ECOM') r.npci_proc_ecom += othDR; else if (ch === 'POS') r.npci_proc_pos += othDR;
      }
    }
  }
  if (!found) return null;
  const totalAssess = r.npci_assess_pos + r.npci_assess_ecom + r.npci_assess_pos_gc + r.npci_assess_ecom_gc + r.npci_assess_cash;
  const totalProc = r.npci_proc_ecom + r.npci_proc_pos;
  r.gross = r2(-(r.pos_amt + r.pos_gc_amt + r.ecom_amt + r.ecom_gc_amt + r.cash_pos_amt) + r.refund_pos + r.refund_ecom + r.qsparc_settled - totalAssess - totalProc
    - r.int_income_cash + r.int_income_pos + r.int_income_ecom + r.int_income_pos_gc + r.int_income_ecom_gc + r.chargeback_pos + r.chargeback_ecom - r.repres_pos - r.repres_ecom);
  r.net_settlement = r2(r.final_net);
  r.fees = r2(totalAssess + totalProc);
  // first data row carries the settlement date
  let date = null;
  for (const row of rows.slice(1, 6)) if (row && row[0] && /\d/.test(String(row[0]))) { date = toISO(row[0]); break; }
  const nm = String(filename || '').match(/ISSUER[_ ](\d{4})-(\d{2})-(\d{2})-(\d+)/i);
  return { date: nm ? isoOf(+nm[1], +nm[2], +nm[3]) : date, fileDate: date, cycleFromName: nm ? +nm[4] : null, filename, data: r,
           // sanity checks of the RuPay tool (V1-V3), kept as data-quality warnings
           warn: ruValidate(r) };
}
function ruValidate(d) {
  const w = [];
  const tf = d.npci_assess_pos + d.npci_assess_ecom + d.npci_assess_pos_gc + d.npci_assess_ecom_gc + d.npci_assess_cash + d.npci_proc_ecom + d.npci_proc_pos;
  const g = Math.round(tf * 0.18 * 100) / 100;
  if (Math.abs(g - d.gst) > 0.10) w.push('V1 fees×18% = ' + g.toFixed(2) + ' vs INWARD GST ' + d.gst.toFixed(2));
  const settled = d.pos_amt + d.ecom_amt + d.qsparc_settled;
  if (d.final_net === 0 && settled > 0) w.push('V3 final net is 0 but settlements = ' + settled.toFixed(2));
  return w;
}

/* ── UPI NTSL (parseNTSL / classifyRemarks) ────────────────────────────── */
const UPI_ORDER = ['10C', '1C', '2C', '3C', '4C', '5C', '6C', '7C', '8C', '9C', 'DC1', 'DC2'];
function classifyRemarks(desc) {
  const d = String(desc || '');
  if (/TDS/i.test(d)) return 'TDS on GST';
  if (/gst\s+on/i.test(d)) return 'GST on Penalty';
  if (/penalty/i.test(d)) return 'Penalty';
  if (/Net Adjusted Amount/i.test(d)) return 'Net Adjusted Amount';
  if (/Transaction Amount/i.test(d)) return 'Transaction Amount';
  if (/Switching Fee\s*Gst|Switching Fee\s*GST/i.test(d)) return 'Switching Fee Gst';
  if (/Switching Fee/i.test(d)) return 'Switching Fee';
  if (/Fee\s*Gst|Fee\s*GST/i.test(d)) return 'Approved Fee Gst';
  if (/Fee/i.test(d)) return 'Approved Fee';
  return null;
}
function parseUpiNtsl(rows, filename) {
  if (rows.length < 5) return null;
  const hdr = String(rows[2] && rows[2][0] || '');
  let date = null, cycle = '';
  const dm = hdr.match(/as on (\d{2})-(\d{2})-(\d{4})/i), cm = hdr.match(/\(\s*([A-Z0-9]+)\s+\d{2}:\d{2}:\d{2}/i);
  if (dm && cm) { date = isoOf(+dm[3], +dm[2], +dm[1]); cycle = cm[1].toUpperCase(); }
  else { const fm = String(filename).match(/NTSLERT(\d{2})(\d{2})(\d{2})_(\w+)\.(xls|xlsx|xlsm|xlsb)$/i); if (!fm) return null; date = isoOf(2000 + +fm[3], +fm[2], +fm[1]); cycle = fm[4].toUpperCase(); }
  cycle = cycle.replace('DC01', 'DC1').replace('DC02', 'DC2');
  const isDC = cycle === 'DC1' || cycle === 'DC2';
  const ch = (rows[4] || []).map(c => String(c || '').toLowerCase());
  const fi = (...kws) => { for (const kw of kws) { const i = ch.findIndex(h => h.includes(kw)); if (i >= 0) return i; } return -1; };
  const colDesc = fi('description', 'particular', 'remark') < 0 ? 0 : fi('description', 'particular', 'remark');
  const colTxn = fi('no. of', 'no of', 'count', 'txn') < 0 ? 1 : fi('no. of', 'no of', 'count', 'txn');
  const colDebit = fi('debit') < 0 ? 2 : fi('debit'), colCred = fi('credit') < 0 ? 3 : fi('credit');
  const SKIP = /^\s*(total|grand|sub.?total|net settlement|settlement amount|opening|closing)/i;
  let debit = 0, credit = 0, txnDebit = 0, txnCredit = 0, nrows = 0;
  for (let ri = 5; ri < rows.length; ri++) {
    const r = rows[ri]; if (!r || !r[colDesc]) continue;
    const desc = String(r[colDesc]).trim(); if (!desc || SKIP.test(desc)) continue;
    const rem = classifyRemarks(desc); if (!rem) continue;
    const deb = sf(r[colDebit]), crd = sf(r[colCred]);
    if (deb === 0 && crd === 0) continue;
    if (isDC && !['Net Adjusted Amount', 'Penalty', 'GST on Penalty', 'TDS on GST'].includes(rem)) continue;
    debit += deb; credit += crd; nrows++;
    if (rem === 'Transaction Amount') { txnDebit += deb; txnCredit += crd; }
  }
  return { date, cycle, filename, rows: nrows, debit: r2(debit), credit: r2(credit), net: r2(credit - debit), txnDebit: r2(txnDebit), txnCredit: r2(txnCredit) };
}

/* ── NFS NTSL (HTML-xls or xlsx) ───────────────────────────────────────── */
const toN = v => { if (v === null || v === undefined || v === '') return 0; const n = parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? 0 : n; };
const toI = v => Math.round(toN(v));
function nfsBuild(cur) {
  const atmFee = (cur.atmFee != null && cur.atmFee !== 0) ? cur.atmFee : (cur.atmCount * 19);
  const atmGST = r2(atmFee * 0.18), atmTot = r2(atmFee + atmGST);
  const biFee = r2(cur.biCount * 6), biGST = r2(biFee * 0.18), biTot = r2(biFee + biGST);
  const mFee = r2(cur.microAmt * 0.005), mGST = r2(mFee * 0.18), mTot = r2(mFee + mGST);
  const swCnt = cur.atmCount + cur.biCount + cur.microCount;
  const swFee = r4(swCnt * 0.30), swGST = r4(swFee * 0.18), swTot = r4(swFee + swGST);
  const disp = cur.disputeAmt || 0;
  const settTot = r4(cur.atmAmt + atmTot + cur.microAmt + mTot + swTot - disp + biTot);       // control-sheet formula X = D+G+M+Q+U−W+K
  return { date: cur.date, finalAmt: cur.finalAmt, settTot, atmCount: cur.atmCount, atmAmt: cur.atmAmt, atmTotFee: atmTot, biCount: cur.biCount, biTot,
           microCount: cur.microCount, microAmt: cur.microAmt, mTot, swCnt, swTot, disp, cycleFromName: cur.cycleFromName || null };
}
function nfsCycleFromName(fname) { const m = String(fname).match(/_(\d+)C\./i); return m ? +m[1] : null; }
function nfsDateFromName(fname) {
  const m = String(fname).match(/NTSLERO(\d{2})(\d{2})(\d{2,4})_\d+C/i); if (!m) return null;
  const yy = +m[3], yr = yy < 100 ? (yy < 50 ? 2000 + yy : 1900 + yy) : yy; return isoOf(yr, +m[2], +m[1]);
}
function nfsParseHtml(text, fname) {
  let date = null, m = text.match(/Daily Settlement Statement[\s\S]{0,100}?as on\s+(\d{2})\/(\d{2})\/(\d{4})/i);
  if (m) date = isoOf(+m[3], +m[2], +m[1]);
  if (!date) { m = text.match(/Daily Settlement Statement[\s\S]{0,100}?as on\s+(\d{2})-(\d{2})-(\d{4})/i); if (m) date = isoOf(+m[3], +m[2], +m[1]); }
  if (!date) date = nfsDateFromName(fname);
  const fsaIdx = text.indexOf('Final Settlement Amount');
  if (!date || fsaIdx < 0) return [];
  let finalAmt = 0;
  const tdVals = [...text.slice(fsaIdx, fsaIdx + 400).matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)].map(x => x[1].replace(/<[^>]+>/g, '').trim());
  for (const v of tdVals) { const n = parseFloat(v); if (n > 0) { finalAmt = n; break; } }             // 0 = a zero cycle (no debit expected)
  let atmCount = 0, atmAmt = 0, atmFee = null, microAmt = 0, microCount = 0, biCount = 0, disputeAmt = 0;
  m = text.match(/Issuer WDL Transaction Amount<\/td>\s*<td[^>]*>(\d+)<\/td>\s*<td[^>]*>([\d.]+)/i); if (m) { atmCount = +m[1]; atmAmt = +m[2]; }
  m = text.match(/Issuer WDL Approved Fee<\/td>\s*<td[^>]*>(\d+)<\/td>\s*<td[^>]*>([\d.]+)/i); if (m) atmFee = +m[2];
  m = text.match(/Issuer WDL Transaction Amount \(Micro-ATM\)<\/td>\s*<td[^>]*>\d+<\/td>\s*<td[^>]*>([\d.]+)/i); if (m) microAmt = +m[1];
  m = text.match(/Issuer WDLTransaction \(Micro ATM\)<\/td>\s*<td[^>]*>Rs\. 100 and above<\/td>\s*<td[^>]*>(\d+)/i); if (m) microCount = +m[1];
  m = text.match(/Issuer BI Approved Fee<\/td>\s*<td[^>]*>(\d+)/i); if (m) biCount = +m[1];
  const dispIdx = text.indexOf('Dispute Adjustments');
  if (dispIdx >= 0) {
    const chunk = text.slice(dispIdx, dispIdx + 800), adjIdx = chunk.indexOf('Adjustment Sub Totals');
    if (adjIdx >= 0) {
      const t = [...chunk.slice(adjIdx, adjIdx + 300).matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)].map(x => x[1].replace(/<[^>]+>/g, '').trim());
      const deb = parseFloat(t[2]) || 0, cre = parseFloat(t[3]) || 0; if (deb || cre) disputeAmt = deb - cre;
    }
  }
  return [nfsBuild({ date, finalAmt, atmCount, atmAmt, atmFee, biCount, microAmt, microCount, disputeAmt, cycleFromName: nfsCycleFromName(fname) })];
}
function nfsExtractDate(s, row) {
  let m = s.match(/as on\s+(\d{2})\/(\d{2})\/(\d{4})/i); if (m) return isoOf(+m[3], +m[2], +m[1]);
  m = s.match(/as on\s+(\d{2})-(\d{2})-(\d{4})/i); if (m) return isoOf(+m[3], +m[2], +m[1]);
  m = s.match(/as on\s+(\d{4})-(\d{2})-(\d{2})/i); if (m) return isoOf(+m[1], +m[2], +m[3]);
  m = s.match(/as on\s+(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})/i);
  if (m) { const mo = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 }[m[2].toLowerCase()]; if (mo) return isoOf(+m[3], mo, +m[1]); }
  if (row) for (const c of row) if (c instanceof Date && c.getFullYear() > 2000) return toISO(c);
  return null;
}
function nfsParseRows(rows, fname) {
  const blocks = []; let cur = null;
  for (const row of rows) {
    if (!row) continue;
    let c0 = ''; if (row[0] !== null && row[0] !== undefined) c0 = (typeof row[0] === 'string') ? row[0].trim() : (row[0] instanceof Date ? '' : String(row[0]).trim());
    if (c0.toLowerCase().includes('daily settlement statement') && c0.toLowerCase().includes('as on')) {
      if (cur && cur.seenFinal) blocks.push(nfsBuild(cur));
      const d = nfsExtractDate(c0, row); if (!d) { cur = null; continue; }
      cur = { date: d, finalAmt: 0, seenFinal: false, atmCount: 0, atmAmt: 0, atmFee: null, biCount: 0, microAmt: 0, microCount: 0, disputeAmt: 0, inDispute: false, cycleFromName: null };
      continue;
    }
    if (!cur) continue;
    if (c0.toLowerCase().includes('dispute adjustment')) { cur.inDispute = true; continue; }
    if (c0.toLowerCase().includes('summary for micro')) { cur.inDispute = false; continue; }
    if (c0 === 'Final Settlement Amount') { cur.seenFinal = true; for (let i = 1; i < Math.min(row.length, 6); i++) { const v = toN(row[i]); if (v > 0) { cur.finalAmt = v; break; } } }
    else if (c0 === 'Issuer WDL Transaction Amount') { cur.atmCount = toI(row[1]); cur.atmAmt = toN(row[2]); }
    else if (c0 === 'Issuer WDL Approved Fee' && !c0.includes('GST') && !c0.includes('Micro')) cur.atmFee = toN(row[2]);
    else if (c0 === 'Issuer WDL Transaction Amount (Micro-ATM)') cur.microAmt = toN(row[2]);
    else if (c0 === 'Issuer WDLTransaction (Micro ATM)' && row[1] && String(row[1]).includes('100 and above')) cur.microCount += toI(row[2]);
    else if (c0 === 'Issuer BI Approved Fee' && !c0.includes('GST')) cur.biCount = toI(row[1]);
    else if (cur.inDispute && c0 === 'Adjustment Sub Totals') { const d = toN(row[2]), cr = toN(row[3]); if (d || cr) cur.disputeAmt = d - cr; }
    else if (cur.inDispute && c0 === 'Net Adjusted Amount') { const v = toN(row[2]); if (v !== 0) cur.disputeAmt = v; }
  }
  if (cur && cur.seenFinal) blocks.push(nfsBuild(cur));
  return blocks;
}

/* ── Ops file (Add Money macro, steps 1–3) ─────────────────────────────── */
const vtrim = s => String(s).replace(/^ +| +$/g, '');
const cstr = v => v === null || v === undefined ? '' : (typeof v === 'boolean' ? (v ? 'True' : 'False') : String(v));
function isNumericV(v) {
  if (v === null || v === undefined) return true;
  if (typeof v === 'number') return Number.isFinite(v);
  if (typeof v === 'boolean') return true;
  const s = String(v);
  return /\d/.test(s) && /^\s*[+-]?(\d{1,3}(,\d{3})+|\d*)(\.\d*)?([eE][+-]?\d+)?\s*$/.test(s);
}
const cdblV = v => v === null || v === undefined ? 0 : typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? -1 : 0) : parseFloat(String(v).replace(/,/g, '').trim());
const ascUpper = s => s.replace(/[a-z]/g, c => c.toUpperCase());
const extractNumbers = s => { let r = ''; for (const c of s) if (c >= '0' && c <= '9') r += c; return r; };
const firstNDigits = (s, n) => { let r = ''; for (const c of s) { if (c >= '0' && c <= '9') { r += c; if (r.length === n) break; } } return r; };
function extractBetweenTilde(s) { const u = ascUpper(s); let p = u.indexOf('~EROUTE~'); if (p < 0) return ''; p += 8; const e = u.indexOf('~', p); return extractNumbers(vtrim(e < 0 ? s.slice(p) : s.slice(p, e))); }
function extractBetweenSlashes(s) { const p1 = s.indexOf('/'); if (p1 < 0) return ''; const p2 = s.indexOf('/', p1 + 1); return extractNumbers(vtrim(p2 < 0 ? s.slice(p1 + 1) : s.slice(p1 + 1, p2))); }
function addMoneyBankUtr(narrRaw) {                                           // macro step 2 – narrative → Final UTR Bank
  const narr = vtrim(cstr(narrRaw)), n = ascUpper(narr);
  if (n.slice(0, 5) === 'IMPS/') return firstNDigits(narr, 12);
  if (n.slice(0, 4) === 'IFT/') return firstNDigits(narr, 11);
  if (n.slice(0, 5) === 'INEFT' || n.slice(0, 5) === 'IRTGS') return extractBetweenTilde(narr);
  if (['NEFT', 'RTGS', 'IMPS'].includes(n.slice(0, 4))) return extractBetweenSlashes(narr);
  return narr !== '' ? extractNumbers(narr) : '';
}
function opsColumns(rows) {
  const h = (rows[0] || []).map(v => vtrim(cstr(v)));
  const idx = n => { let k = -1; h.forEach((v, j) => { if (v === n) k = j; }); return k; };
  return { cat: idx('TransactionCategory'), type: idx('TransactionType'), utr: idx('UtrNo'), amt: idx('TotalTransactionAmount'), ent: idx('EnterpriseName'), date: idx('TransactionDate') };
}
function isOpsFile(rows) { const c = opsColumns(rows); return c.cat >= 0 && c.type >= 0 && c.utr >= 0 && c.amt >= 0; }
function addMoneyOpsRows(rows) {                                               // macro step 1 → [{utr, amt, ent, row}]
  const c = opsColumns(rows); let last = 0;
  for (let i = rows.length - 1; i >= 1; i--) if (rows[i] && !isBlank(rows[i][0])) { last = i; break; }
  const out = [];
  for (let i = 1; i <= last; i++) {
    const r = rows[i] || [];
    const cat = vtrim(cstr(r[c.cat])), typ = vtrim(cstr(r[c.type])), utr = vtrim(cstr(r[c.utr])), amt = r[c.amt];
    if (!(['ADD MONEY', 'OPS ADJUSTMENT', 'OPS ADJUSTEMENT'].includes(cat.toUpperCase()) && typ.toUpperCase() === 'CREDIT')) continue;
    if (ascUpper(utr).slice(0, 4) === 'TRXV') continue;                        // TRXV records ignored
    const fin = ascUpper(utr).slice(0, 3) === 'RE1' ? utr.slice(3) : extractNumbers(utr);
    out.push({ row: i + 1, utr: fin, amt: isNumericV(amt) ? cdblV(amt) : null, ent: c.ent >= 0 ? vtrim(cstr(r[c.ent])) : '', date: c.date >= 0 ? toISO(r[c.date]) : null });
  }
  return out;
}

/* ═════════════════════════ FILE INTAKE ════════════════════════════════ */
const DATA_EXT = /\.(xls|xlsx|xlsm|xlsb|csv|html?)$/i;
function baseName(p) { return String(p).split(/[\\/]/).pop(); }
function nameKind(name) {
  const n = baseName(name);
  if (/\.zip$/i.test(n)) return 'zip';
  if (!DATA_EXT.test(n)) return 'ignore';
  if (/^~\$/.test(n) || /^\./.test(n)) return 'ignore';
  if (/UPI_NTSLERT\d{6}_/i.test(n)) return 'UPI';
  if (/NTSLERO\d{6,8}_\d+C/i.test(n)) return 'NFS';
  if (/DSRSummaryReport/i.test(n)) return 'RUPAY';
  if (/^DSRReport/i.test(n)) return 'NETC';
  return 'sniff';
}
async function expandEntries(entries, onProgress) {
  const out = []; const queue = entries.slice();
  while (queue.length) {
    const e = queue.shift();
    const kind = nameKind(e.name);
    if (kind === 'zip') {
      try {
        const zip = await root.JSZip.loadAsync(await e.getBuffer());
        const names = Object.keys(zip.files).filter(k => !zip.files[k].dir);
        for (const k of names) { const f = zip.files[k]; queue.push({ name: baseName(k), path: (e.path || e.name) + ' › ' + k, size: f._data && f._data.uncompressedSize || 0, getBuffer: () => f.async('uint8array') }); }
      } catch (err) { out.push({ entry: e, kind: 'error', note: 'Cannot open zip: ' + err.message }); }
    } else out.push({ entry: e, kind });
    if (onProgress) onProgress(out.length, queue.length);
  }
  return out;
}
function sniffRows(rows, text) {
  const head = rows.slice(0, 12);
  if (rows[0] && rows[0][0] === 'Settlement Date') {
    const g = String(rows[0][6] || '');
    if (/Status\(Approved\/Declined\)/i.test(g)) return 'RUPAY';
    if (/Txn_Status/i.test(g)) return 'NETC';
    return 'NETC';
  }
  if (head.some(r => r && /Unified Payments Interface/i.test(String(r[0] || '')))) return 'UPI';
  if (head.some(r => r && /Daily Settlement Statement/i.test(String(r[0] || '')))) return 'NFS';
  if (isOpsFile(rows)) return 'OPS';
  if (detectBankHeader(rows)) return 'BANK';
  return null;
}

/* Reads every entry, decides its product, parses it. Returns the whole intake. */
async function intake(entries, opt) {
  opt = opt || {}; const log = [];
  const expanded = await expandEntries(entries, opt.onExpand);
  const res = { bank: null, bankFile: null, ops: [], netc: [], nfs: [], rupay: [], upi: [], files: [], warnings: [] };
  let i = 0;
  for (const x of expanded) {
    i++; if (opt.onProgress) opt.onProgress(i, expanded.length, x.entry.name);
    if (opt.yield && i % 8 === 0) await opt.yield();
    const rec = { name: x.entry.name, path: x.entry.path || x.entry.name, kind: x.kind, product: null, date: null, cycle: null, status: 'ok', note: '' };
    res.files.push(rec);
    if (x.kind === 'ignore') { rec.status = 'ignored'; rec.note = 'not a data file'; continue; }
    if (x.kind === 'error') { rec.status = 'error'; rec.note = x.note; continue; }
    try {
      const buf = toU8(await x.entry.getBuffer());
      let kind = x.kind, rows = null, wb = null, text = null;
      const mg = magic(buf);
      const loadRows = () => { if (!rows) { wb = readWB(buf); rows = sheetRows(wb); } return rows; };
      if (mg === 'text' && !/\.csv$/i.test(rec.name)) text = decodeText(buf);
      if (kind === 'sniff') {
        if (text !== null) kind = /Issuer WDL|Daily Settlement Statement/i.test(text) ? (/Unified Payments Interface/i.test(text) ? 'UPI' : 'NFS') : 'unknown';
        else { kind = sniffRows(loadRows()) || 'unknown'; }
      }
      rec.product = kind;
      if (kind === 'unknown') { rec.status = 'ignored'; rec.note = 'not recognised as a settlement, bank or Ops file'; continue; }
      if (kind === 'BANK') {
        const b = parseBankStatement(loadRows(), rec.name);
        if (!b) { rec.status = 'error'; rec.note = 'bank header not found'; continue; }
        if (res.bank) { rec.status = 'ignored'; rec.note = 'a bank statement is already loaded'; continue; }
        res.bank = b; res.bankFile = rec.name; rec.note = b.lines.length + ' lines'; continue;
      }
      if (kind === 'OPS') { res.ops.push({ rows: loadRows(), name: rec.name }); rec.note = 'Add Money Ops file'; continue; }
      if (kind === 'NETC') {
        const p = parseNetcDsr(loadRows(), rec.name);
        if (!p) { rec.status = 'error'; rec.note = 'no DSR grand-total found'; continue; }
        rec.date = p.date; rec.cycle = p.cycleFromName; p.fileRec = rec; res.netc.push(p); continue;
      }
      if (kind === 'RUPAY') {
        const p = parseRupayDsr(loadRows(), rec.name);
        if (!p || !p.date) { rec.status = 'error'; rec.note = 'no DSR grand-total / date found'; continue; }
        rec.date = p.date; rec.cycle = p.cycleFromName; p.fileRec = rec; res.rupay.push(p); continue;
      }
      if (kind === 'UPI') {
        const p = parseUpiNtsl(loadRows(), rec.name);
        if (!p) { rec.status = 'error'; rec.note = 'cannot read date/cycle'; continue; }
        rec.date = p.date; rec.cycle = p.cycle; p.fileRec = rec; res.upi.push(p); continue;
      }
      if (kind === 'NFS') {
        let blocks;
        if (text !== null) blocks = nfsParseHtml(text, rec.name);
        else blocks = (() => { const w = wb || readWB(buf); const out = []; for (const sn of w.SheetNames) out.push(...nfsParseRows(sheetRows(w, sn), rec.name)); return out; })();
        if (!blocks.length) { rec.status = 'error'; rec.note = 'no settlement block found'; continue; }
        blocks.forEach(b => { if (!b.cycleFromName) b.cycleFromName = nfsCycleFromName(rec.name); b.filename = rec.name; b.fileRec = rec; });
        rec.date = blocks[0].date; rec.cycle = blocks[0].cycleFromName; res.nfs.push(...blocks); continue;
      }
    } catch (err) { rec.status = 'error'; rec.note = err.message; }
  }
  return res;
}

/* Each product → list of cycles {date, cycle, ...}; duplicates removed (file records flagged, idempotent). */
function dedupe(list, keyFn, valFn, label, warnings) {
  const seen = new Map(), out = [];
  for (const it of list) {
    const k = keyFn(it);
    if (seen.has(k)) {
      const prev = seen.get(k);
      if (Math.abs(valFn(prev) - valFn(it)) < 0.005) warnings.push(label + ' ' + k.replace('|', ' cycle ') + ': duplicate file (identical) ignored — ' + it.filename);
      else warnings.push(label + ' ' + k.replace('|', ' cycle ') + ': TWO DIFFERENT FILES for the same cycle (' + prev.filename + ' vs ' + it.filename + ') — first one used');
      if (it.fileRec) { it.fileRec.status = 'duplicate'; it.fileRec.note = 'same date+cycle already loaded'; }
      continue;
    }
    if (it.fileRec && it.fileRec.status === 'duplicate') { it.fileRec.status = 'ok'; it.fileRec.note = ''; }
    seen.set(k, it); out.push(it);
  }
  return out;
}
function assignOrderCycles(list, getDate, setCycle, getHave) {                 // files without a cycle in the name: order within the day
  const ctr = {};
  list.filter(x => !getHave(x)).sort((a, b) => String(a.filename).localeCompare(String(b.filename))).forEach(x => { const d = getDate(x); ctr[d] = (ctr[d] || 0) + 1; setCycle(x, ctr[d]); });
}
/* Source cycles per product, normalised to {key,date,cycle,expected,file,detail}. */
function prepSources(intakeRes, warnings) {
  const prep = {};
  { // NETC
    const L = intakeRes.netc.slice();
    assignOrderCycles(L, x => x.date, (x, c) => { x.cycleFromName = c; }, x => x.cycleFromName);
    const dd = dedupe(L, x => x.date + '|' + x.cycleFromName, x => x.net, 'NETC', warnings);
    prep.NETC = dd.map(x => ({ key: x.date + '|' + x.cycleFromName, date: x.date, cycle: String(x.cycleFromName), expected: x.net, file: x.filename,
      detail: { count: x.cnt, txnAmt: x.txnAmt, taxable: x.taxable, gst: x.gst, svcTotal: x.svcTotal, netPayable: Math.abs(x.net) } }));
  }
  { // NFS
    const L = intakeRes.nfs.slice();
    assignOrderCycles(L, x => x.date, (x, c) => { x.cycleFromName = c; }, x => x.cycleFromName);
    const dd = dedupe(L, x => x.date + '|C' + x.cycleFromName, x => x.finalAmt, 'NFS', warnings);
    prep.NFS = dd.map(x => ({ key: x.date + '|C' + x.cycleFromName, date: x.date, cycle: 'C' + x.cycleFromName, expected: -r2(x.finalAmt), file: x.filename,
      detail: { atmCount: x.atmCount, atmAmt: x.atmAmt, atmFeeGst: x.atmTotFee, biCount: x.biCount, biFeeGst: x.biTot, microCount: x.microCount, microAmt: x.microAmt, microFeeGst: x.mTot,
                swCount: x.swCnt, swFeeGst: x.swTot, disputes: x.disp, computedTotal: x.settTot, finalAmt: x.finalAmt } }));
  }
  { // RuPay
    const L = intakeRes.rupay.slice();
    assignOrderCycles(L, x => x.date, (x, c) => { x.cycleFromName = c; }, x => x.cycleFromName);
    const dd = dedupe(L, x => x.date + '|C' + x.cycleFromName, x => x.data.net_settlement, 'RUPAY', warnings);
    prep.RUPAY = dd.map(x => ({ key: x.date + '|C' + x.cycleFromName, date: x.date, cycle: 'C' + x.cycleFromName, expected: x.data.net_settlement, file: x.filename,
      detail: { sales: r2(x.data.pos_amt + x.data.ecom_amt + x.data.pos_gc_amt + x.data.ecom_gc_amt), cash: r2(x.data.cash_pos_amt), refunds: r2(x.data.refund_pos + x.data.refund_ecom), qsparc: r2(x.data.qsparc_settled),
                chargebacks: r2(x.data.chargeback_pos + x.data.chargeback_ecom), represent: r2(x.data.repres_pos + x.data.repres_ecom), fees: x.data.fees,
                interchange: r2(x.data.int_income_pos + x.data.int_income_ecom + x.data.int_income_pos_gc + x.data.int_income_ecom_gc + x.data.int_income_cash), gst: x.data.gst, gross: x.data.gross, finalNet: x.data.final_net }, warn: x.warn }));
  }
  { // UPI
    const dd = dedupe(intakeRes.upi.slice(), x => x.date + '|' + x.cycle, x => x.net, 'UPI', warnings);
    prep.UPI = dd.map(x => ({ key: x.date + '|' + x.cycle, date: x.date, cycle: x.cycle, expected: x.net, file: x.filename,
      detail: { debit: x.debit, credit: x.credit, net: x.net, txnDebit: x.txnDebit, txnCredit: x.txnCredit, rows: x.rows } }));
  }
  return prep;
}
/* Which cycles are present / missing, per product and day (for the intake screen). */
const EXPECT = { NETC: ['1', '2', '3'], NFS: ['C1', 'C2', 'C3', 'C4'], RUPAY: ['C1', 'C2', 'C3', 'C4'], UPI: UPI_ORDER };
function coverage(intakeRes) {
  const prep = prepSources(intakeRes, []);
  const out = {};
  for (const p of PRODUCTS) {
    const src = prep[p]; if (!src.length) continue;
    const dates = [...new Set(src.map(s => s.date))].sort();
    const first = dates[0], last = dates[dates.length - 1], days = [];
    for (let d = first, n = 0; d <= last && n < 400; d = addDaysISO(d, 1), n++) {
      const have = new Set(src.filter(s => s.date === d).map(s => s.cycle));
      const missing = EXPECT[p].filter(c => !have.has(c));
      days.push({ date: d, found: EXPECT[p].length - missing.length, expected: EXPECT[p].length, missing, extra: [...have].filter(c => !EXPECT[p].includes(c)) });
    }
    out[p] = { files: src.length, first, last, days, found: days.reduce((a, x) => a + x.found, 0), expected: days.reduce((a, x) => a + x.expected, 0),
               missing: days.filter(x => x.missing.length).map(x => ({ date: x.date, cycles: x.missing })) };
  }
  return out;
}

/* ═════════════════════════ RECONCILIATION ═════════════════════════════ */
const ST = { MATCHED: 'MATCHED', REFMIS: 'MATCHED – REFERENCE MISMATCH', DIFF: 'DIFFERENCE', PENDING: 'PENDING', NOTBANK: 'NOT IN BANK', ZERO: 'ZERO – NO ENTRY EXPECTED' };

/* Generic matcher. sources: [{key,date,cycle,expected,...}]; bank: Map key → {amount, lines:[line]} */
function matchProduct(product, sources, bankLines, ctx) {
  const tol = ctx.tol[product];
  const within = d => INCLUSIVE[product] ? Math.abs(d) <= tol : Math.abs(d) < tol;
  // group bank lines of this product by key
  const bankMap = new Map();
  for (const bl of bankLines) {
    const k = bl.cls.date + '|' + bl.cls.cycle;
    if (!bankMap.has(k)) bankMap.set(k, { amount: 0, lines: [], date: bl.cls.date, cycle: bl.cls.cycle });
    const g = bankMap.get(k); g.amount = r2(g.amount + (bl.credit - bl.debit)); g.lines.push(bl);
  }
  const used = new Set(); const rows = [];
  const lagObs = [];
  // pass 1: exact key
  for (const s of sources) {
    const k = s.date + '|' + s.cycle; const row = Object.assign({}, s, { key: k, bank: null, diff: null, status: null, note: '', bankLines: [] });
    const g = bankMap.get(k);
    if (g) {
      used.add(k); row.bank = g.amount; row.bankLines = g.lines; row.diff = r2(s.expected - g.amount);
      row.bankDate = g.lines[0].date;
      if (g.lines.length > 1) row.note = g.lines.length + ' bank lines added together';
      row.status = within(row.diff) ? ST.MATCHED : ST.DIFF;
      lagObs.push(diffDays(g.lines[g.lines.length - 1].date, s.date));
    }
    rows.push(row);
  }
  // pass 2: reference mismatch – an unused bank entry with the same cycle and the same amount
  for (const row of rows) {
    if (row.status || row.expected === 0) continue;
    for (const [k, g] of bankMap) {
      if (used.has(k) || g.cycle !== row.cycle) continue;
      if (within(r2(row.expected - g.amount))) {
        used.add(k); row.bank = g.amount; row.bankLines = g.lines; row.diff = r2(row.expected - g.amount); row.bankDate = g.lines[0].date;
        row.status = ST.REFMIS; row.note = 'bank narrative carries ' + fmtDMY(g.date) + ' ' + g.cycle; break;
      }
    }
  }
  // observed posting lag (days after the source date), used to decide pending vs missing
  const lag = Math.max(1, ...lagObs.filter(n => n >= 0 && n <= 5));
  const stmtEnd = ctx.statementEnd;
  for (const row of rows) {
    if (row.status) continue;
    if (row.expected === 0) { row.status = ST.ZERO; continue; }
    if (stmtEnd && diffDays(stmtEnd, row.date) < lag) { row.status = ST.PENDING; row.note = 'bank statement ends ' + fmtDMY(stmtEnd) + '; posting expected by ' + fmtDMY(addDaysISO(row.date, lag)); }
    else row.status = ST.NOTBANK;
    row.diff = r2(row.expected);                      // nothing in the bank: the whole expected amount is open
  }
  // bank-only lines
  // reconciliation period = the whole calendar month(s) covered by the loaded files, so a missing day shows up as BANK ONLY
  const dates = sources.map(s => s.date).sort(); const lastDay = iso => addDaysISO(addDaysISO(iso.slice(0, 7) + '-01', 31).slice(0, 7) + '-01', -1);
  const pStart = dates[0].slice(0, 7) + '-01', pEnd = lastDay(dates[dates.length - 1]);
  const bankOnly = [];
  for (const [k, g] of bankMap) {
    if (used.has(k)) continue;
    const cls = g.date < pStart ? 'PRIOR PERIOD' : g.date > pEnd ? 'NEXT PERIOD' : 'BANK ONLY';
    bankOnly.push({ key: k, date: g.date, cycle: g.cycle, amount: g.amount, lines: g.lines, cls,
      note: cls === 'PRIOR PERIOD' ? 'settles a cycle before this period (' + fmtDMY(pStart) + ') — reconciled in that month\'s run' : cls === 'NEXT PERIOD' ? 'cycle is after this period (' + fmtDMY(pEnd) + ') — belongs to the next month\'s run' : 'inside the period but there is no source file for this cycle — a file may be missing' });
  }
  for (const row of rows) row.bankLines.forEach(l => { l.result = row.status; l.resultKey = row.key; });
  for (const b of bankOnly) b.lines.forEach(l => { l.result = b.cls; l.resultKey = b.key; });
  rows.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : cycleOrder(product, a.cycle) - cycleOrder(product, b.cycle));
  bankOnly.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : cycleOrder(product, a.cycle) - cycleOrder(product, b.cycle));
  return { rows, bankOnly, period: { start: pStart, end: pEnd }, lag, tol };
}
function cycleOrder(product, c) {
  if (product === 'UPI') { const i = UPI_ORDER.indexOf(String(c)); return i < 0 ? 99 : i; }
  const n = parseInt(String(c).replace(/\D/g, ''), 10); return isNaN(n) ? 99 : n;
}
function totalsOf(r) {
  const rows = r.rows, cnt = s => rows.filter(x => x.status === s).length;
  const matched = cnt(ST.MATCHED), refmis = cnt(ST.REFMIS);
  return {
    cycles: rows.length, matched, refmis, diff: cnt(ST.DIFF), pending: cnt(ST.PENDING), notBank: cnt(ST.NOTBANK), zero: cnt(ST.ZERO),
    expected: r2(rows.reduce((s, x) => s + x.expected, 0)), bank: r2(rows.reduce((s, x) => s + (x.bank || 0), 0)),
    open: r2(rows.filter(x => [ST.DIFF, ST.PENDING, ST.NOTBANK].includes(x.status)).reduce((s, x) => s + (x.diff || 0), 0)),
    prior: r.bankOnly.filter(b => b.cls === 'PRIOR PERIOD').length, priorAmt: r2(r.bankOnly.filter(b => b.cls === 'PRIOR PERIOD').reduce((s, b) => s + b.amount, 0)),
    next: r.bankOnly.filter(b => b.cls === 'NEXT PERIOD').length, nextAmt: r2(r.bankOnly.filter(b => b.cls === 'NEXT PERIOD').reduce((s, b) => s + b.amount, 0)),
    bankOnly: r.bankOnly.filter(b => b.cls === 'BANK ONLY').length, bankOnlyAmt: r2(r.bankOnly.filter(b => b.cls === 'BANK ONLY').reduce((s, b) => s + b.amount, 0)),
  };
}

/* ── ADD MONEY: macro step 3 on Ops vs the bank credits nobody else claimed ── */
function reconAddMoney(opsFiles, creditLines, ctx) {
  const tol = ctx.tol.ADDMONEY;
  const opsAmt = new Map(), opsFirst = new Map(), ent = new Map(), opsDate = new Map();
  for (const f of opsFiles) for (const o of addMoneyOpsRows(f.rows)) {
    if (o.utr === '' || !isNumericV(o.amt === null ? null : o.amt)) continue;
    const a = o.amt === null ? 0 : o.amt;
    if (opsAmt.has(o.utr)) opsAmt.set(o.utr, opsAmt.get(o.utr) + a); else { opsAmt.set(o.utr, a); opsFirst.set(o.utr, o.row); }
    if (!ent.has(o.utr)) { ent.set(o.utr, o.ent); opsDate.set(o.utr, o.date); }
  }
  const bankAmt = new Map(), bankLine = new Map(), bankLinesOf = new Map();
  for (const l of creditLines) {
    const u = addMoneyBankUtr(l.narr); if (u === '') continue;
    const a = l.credit;
    if (bankAmt.has(u)) { bankAmt.set(u, bankAmt.get(u) + a); bankLinesOf.get(u).push(l); } else { bankAmt.set(u, a); bankLine.set(u, l); bankLinesOf.set(u, [l]); }
  }
  const rows = [], matched = new Set();
  for (const [k, oa] of opsAmt) {
    let ba = null, diff, status;
    if (bankAmt.has(k)) { ba = bankAmt.get(k); diff = r2(oa - ba); status = Math.abs(diff) < tol ? ST.MATCHED : ST.DIFF; matched.add(k); }
    else { diff = r2(oa); status = ST.NOTBANK; }
    rows.push({ utr: k, name: ent.get(k) || '', date: opsDate.get(k) || (bankLine.get(k) && bankLine.get(k).date) || null, expected: r2(oa), bank: ba === null ? null : r2(ba), diff, status, opsRow: opsFirst.get(k), bankRow: bankLine.get(k) && bankLine.get(k).row });
    (bankLinesOf.get(k) || []).forEach(l => { l.result = status; });
  }
  const bankOnly = [];
  for (const [k, ba] of bankAmt) if (!matched.has(k)) { const l = bankLine.get(k); bankOnly.push({ utr: k, amount: r2(ba), date: l.date, narr: l.narr, row: l.row, lines: bankLinesOf.get(k), cls: 'NOT IN OPS' }); bankLinesOf.get(k).forEach(x => { x.result = 'NOT IN OPS'; }); }
  return { rows, bankOnly };
}

/* ═════════════════════════ ORCHESTRATOR ═══════════════════════════════ */
function reconcileAll(intakeRes, options) {
  options = options || {};
  const tol = Object.assign({}, DEFAULT_TOL, options.tol || {});
  const warnings = intakeRes.warnings.slice();
  const bank = intakeRes.bank;
  const out = { products: {}, warnings, bank: null, tol, generated: new Date().toISOString() };
  if (!bank) { warnings.push('No bank statement found.'); return out; }

  // classify every bank line
  const lines = bank.lines.map(l => Object.assign({}, l, { cls: classifyBankLine(l), claimedBy: null }));
  const stmtEnd = lines.reduce((m, l) => l.date > m ? l.date : m, '') || bank.meta.periodEnd;      // last transaction date actually present
  const ctx = { tol, statementEnd: stmtEnd };

  // ── sources per product (cycle numbers, dedupe)
  const files = intakeRes.files;
  const prep = prepSources(intakeRes, warnings);
  for (const x of intakeRes.rupay) (x.warn || []).forEach(w => warnings.push('RuPay ' + x.date + ' C' + x.cycleFromName + ': ' + w));

  // ── per product reconciliation
  for (const p of PRODUCTS) {
    const src = prep[p]; if (!src.length) continue;
    const bl = lines.filter(l => l.cls && l.cls.product === p);
    const r = matchProduct(p, src, bl, ctx);
    r.product = p; r.totals = totalsOf(r); r.bankLineCount = bl.length;
    bl.forEach(l => { l.claimedBy = p; });
    // coverage: expected cycles per day
    out.products[p] = r;
  }
  // bank lines of a product that has NO source files loaded: still that product's lines, never Add Money
  out.noSource = {};
  for (const p of PRODUCTS) if (!out.products[p]) {
    const bl = lines.filter(l => l.cls && l.cls.product === p);
    if (bl.length) { bl.forEach(l => { l.claimedBy = p; }); out.noSource[p] = { lines: bl.length, amount: r2(bl.reduce((s, l) => s + l.credit - l.debit, 0)) }; }
    if (bl.length) warnings.push(PRODUCT_LABEL[p] + ': the bank statement has ' + bl.length + ' line(s) for this product but no ' + PRODUCT_LABEL[p] + ' files were loaded, so they are not reconciled.');
  }

  // ── Add Money: only the credits nobody else claimed
  const free = lines.filter(l => !l.claimedBy);
  const creditsFree = free.filter(l => l.credit > 0);
  if (intakeRes.ops.length) {
    const am = reconAddMoney(intakeRes.ops, creditsFree, ctx);
    am.totals = {
      rows: am.rows.length, matched: am.rows.filter(r => r.status === ST.MATCHED).length, diff: am.rows.filter(r => r.status === ST.DIFF).length,
      notBank: am.rows.filter(r => r.status === ST.NOTBANK).length, notOps: am.bankOnly.length,
      opsMatched: r2(am.rows.filter(r => r.status === ST.MATCHED).reduce((s, r) => s + r.expected, 0)),
      bankMatched: r2(am.rows.filter(r => r.status === ST.MATCHED).reduce((s, r) => s + (r.bank || 0), 0)),
      netDiff: r2(am.rows.filter(r => r.status === ST.DIFF).reduce((s, r) => s + r.diff, 0)),
      notOpsAmt: r2(am.bankOnly.reduce((s, b) => s + b.amount, 0)), notBankAmt: r2(am.rows.filter(r => r.status === ST.NOTBANK).reduce((s, r) => s + r.expected, 0)),
    };
    out.products.ADDMONEY = am;
    const used = new Set(am.rows.filter(r => r.bank !== null).map(r => r.utr));
    creditsFree.forEach(l => { l.claimedBy = 'ADDMONEY'; });
  } else {
    out.addMoneyPending = { lines: creditsFree.length, amount: r2(creditsFree.reduce((s, l) => s + l.credit, 0)) };
  }

  // ── bank coverage / statement integrity
  const sumD = r2(lines.reduce((s, l) => s + l.debit, 0)), sumC = r2(lines.reduce((s, l) => s + l.credit, 0));
  const cover = { lines: lines.length, debit: sumD, credit: sumC, byProduct: {}, other: {} };
  for (const l of lines) {
    const k = l.claimedBy || 'OTHER';
    const g = cover.byProduct[k] || (cover.byProduct[k] = { lines: 0, debit: 0, credit: 0 });
    g.lines++; g.debit = r2(g.debit + l.debit); g.credit = r2(g.credit + l.credit);
    if (k === 'OTHER') { const f = bankFamily(l); const o = cover.other[f] || (cover.other[f] = { lines: 0, debit: 0, credit: 0 }); o.lines++; o.debit = r2(o.debit + l.debit); o.credit = r2(o.credit + l.credit); }
  }
  const m = bank.meta;
  cover.integrity = {
    opening: m.opening, closing: m.closing, declaredDebits: m.totalDebits, declaredCredits: m.totalCredits,
    computedClosing: m.opening === null ? null : r2(m.opening + sumC - sumD),
    ok: m.opening === null ? null : Math.abs((m.opening + sumC - sumD) - m.closing) < 0.05,
    debitsOk: m.totalDebits === null ? null : Math.abs(sumD - m.totalDebits) < 0.05,
    creditsOk: m.totalCredits === null ? null : Math.abs(sumC - m.totalCredits) < 0.05,
  };
  out.bank = { meta: m, lines, cover, file: intakeRes.bankFile };
  out.files = files; out.ctx = { statementEnd: stmtEnd };
  // overall
  const allT = Object.entries(out.products).filter(([, v]) => v).map(([k, v]) => [k, v.totals]);
  out.overall = {
    cycles: allT.reduce((s, [k, t]) => s + (t.cycles || t.rows || 0), 0),
    matched: allT.reduce((s, [k, t]) => s + (t.matched || 0) + (t.refmis || 0), 0),
    attention: allT.reduce((s, [k, t]) => s + (t.diff || 0) + (t.notBank || 0) + (t.pending || 0) + (t.bankOnly || 0), 0),
    notOps: allT.reduce((s, [k, t]) => s + (t.notOps || 0), 0),
  };
  return out;
}

root.OSEngine = { coverage, prepSources, EXPECT, PRODUCTS, PRODUCT_LABEL, DEFAULT_TOL, ST, UPI_ORDER, intake, reconcileAll, parseBankStatement, classifyBankLine, nameKind, toISO, fmtDMY, addDaysISO, diffDays, parseAmount, r2,
  parsers: { parseNetcDsr, parseRupayDsr, parseUpiNtsl, nfsParseHtml, nfsParseRows, addMoneyBankUtr, addMoneyOpsRows } };
})(typeof window !== 'undefined' ? window : globalThis);
