/* Help content shown by the "Rules & parameters" and "How to use" icons in each tab.
   Derived from the logic of each tool in src/modules/. If a tool's rules change, update the text here. */
const HELP = {

/* ───────────────────────────── FASTAG ───────────────────────────── */
fastag: {
  rules: [
    ['Purpose', `<p>Reconciles NETC FASTag settlement: each <b>DSR cycle</b> (one DSR file) against the <b>bank debit</b> for that cycle.</p>`],
    ['Inputs', `<ul>
      <li><b>DSR folder</b> – every .xlsx / .xls inside is read; <b>one file = one DSR cycle</b>. Files are sorted by file name, and the cycle number is the order of the file within the same date (1st file for a date = Cy1, 2nd = Cy2 …).</li>
      <li><b>Bank statement</b> – first sheet of an .xlsx / .xls (e.g. <code>Bank_Fastag.xlsx</code>).</li></ul>`],
    ['DSR parsing', `<ul>
      <li>Header row = a row (within the first 5) whose first column is <code>Settlement Date</code>. Block date = first date below it.</li>
      <li><b>Txn count / Txn amount</b> = sum of rows where Type = <code>DEBIT</code> and Category = <code>TOLL</code> or <code>PARKING</code>.</li>
      <li><b>GST</b> = the <code>INWARD GST</code> row. <b>Total service fee</b> = the grand <code>Total</code> row. <b>Taxable</b> = Total service fee − GST.</li>
      <li><b>Net payable</b> = final net on the grand <code>Total</code> row.</li>
      <li>Files without a recognisable DSR layout are skipped and listed as “skipped”.</li></ul>`],
    ['Bank parsing', `<ul>
      <li>Only rows whose narrative contains <code>_00</code> followed by <code>YYYYMMDD</code> + cycle number (e.g. <code>…_00202604011</code>).</li>
      <li>Debit amount and payment date are taken from the bank row. If the same date+cycle appears twice, the <b>last</b> one wins.</li></ul>`],
    ['Matching rules', `<table class="rt"><tr><th>Status</th><th>Rule</th></tr>
      <tr><td>Matched</td><td>Bank key (DSR date + cycle) found and |bank − DSR net| &lt; ₹0.01</td></tr>
      <tr><td>Amt Diff ₹x</td><td>Key found but amount differs</td></tr>
      <tr><td>Matched – Narrative Mismatch*</td><td>No key match, but an unused bank entry with the <b>same cycle number</b> and the <b>same amount</b> exists (bank narrative carries the wrong DSR date)</td></tr>
      <tr><td>Zero DSR – No Debit</td><td>DSR net is 0; no debit expected</td></tr>
      <tr><td>Pending – Next Month Debit</td><td>No bank entry, and the DSR is on the <b>last date</b> loaded</td></tr>
      <tr><td>Unmatched – Investigate</td><td>No bank entry for an earlier date → escalate</td></tr></table>`],
    ['Reconciliation bridge', `<p><b>Bank total + Pending next month = DSR net payable.</b> Net difference under ₹0.05 is shown as NIL / full recon achieved.</p>`],
    ['Output', `<p>Download Excel with sheets <b>Recon Detail</b>, <b>Summary</b>, <b>Exceptions</b>. File name <code>Fastag_Recon_MM_YYYY.xlsx</code>.</p>`],
  ],
  how: [
    ['Steps', `<ol>
      <li>Click the <b>DSR Folder</b> box and select the folder holding the whole month's DSR files (or drag the folder in).</li>
      <li>Wait for “✓ n DSR files parsed”. The preview lists each date and its number of cycles.</li>
      <li>Drop or select the <b>Bank Statement</b> file.</li>
      <li>Click <b>Run Reconciliation</b>.</li>
      <li>Review the <b>Summary</b>, <b>Recon Detail</b> and <b>Exceptions</b> tabs.</li>
      <li>Click <b>Download Recon Excel</b> to save the report.</li></ol>`],
    ['Reading the result', `<ul>
      <li>Green rows = matched, yellow = narrative mismatch / amount difference, orange = pending next month, grey = zero cycle, red = unmatched.</li>
      <li><b>Exceptions</b>: Section 0 unmatched (escalate), A narrative mismatch, B pending next month, C zero-net cycles.</li></ul>`],
    ['Tips', `<ul><li>Load the <b>full month</b> of DSRs; the last date in the folder is treated as the one whose debit may fall in next month's statement.</li>
      <li>Files that are not DSRs are skipped, not failed.</li></ul>`],
  ]
},

/* ───────────────────────────── NFS ───────────────────────────── */
nfs: {
  rules: [
    ['Purpose', `<p>Reconciles NFS (ATM) settlement: each <b>NTSL cycle</b> (4 cycles a day, T+1 settlement) against the <b>bank debit</b> for that cycle.</p>`],
    ['Inputs', `<ul>
      <li><b>NTSL files</b> – any of: NPCI HTML-style .xls, real .xls, .xlsx / .xlsm, individual files or one merged workbook. Each upload <b>adds</b> to what is loaded; use <b>✕ Clear</b> to start over.</li>
      <li><b>Bank statement</b> – .xlsx with a header row (within the first 8 rows) containing <code>Narrative</code> and <code>Debit</code>; <code>Payment date</code> and <code>Remarks</code> columns are used if present.</li></ul>`],
    ['NTSL parsing', `<ul>
      <li>Date from the header “Daily Settlement Statement … as on DD/MM/YYYY” (or from filename <code>NTSLERODDMMYY_NC</code>).</li>
      <li>Cycle from the filename suffix <code>_NC</code> (→ C<i>N</i>); otherwise assigned C1, C2, C3, C4 by block order within the date.</li>
      <li><b>Final Settlement Amount</b> = first positive value on that row. A block without it is ignored.</li></ul>`],
    ['Bank parsing', `<ul>
      <li>If a <code>Remarks</code> column exists it must be <code>NFS</code>; narrative must contain <code>NFS</code>.</li>
      <li>Narrative must end with <code>_00YYYYMMDDCC</code> → transaction date and cycle (e.g. <code>…_0020260401<b>02</b></code> = C2).</li>
      <li>Only debits &gt; 0 are used. Settlement date = bank payment date.</li></ul>`],
    ['Derived amounts (control-sheet formulas)', `<table class="rt">
      <tr><th>Item</th><th>Rule</th></tr>
      <tr><td>ATM fee</td><td>Issuer WDL Approved Fee from NTSL; if zero/missing = ATM count × ₹19. GST 18%</td></tr>
      <tr><td>BI fee</td><td>BI approved count × ₹6; GST 18%</td></tr>
      <tr><td>Micro-ATM fee</td><td>Micro-ATM amount × 0.5%; GST 18%</td></tr>
      <tr><td>Switching fee</td><td>(ATM + BI + Micro count) × ₹0.30; GST 18%</td></tr>
      <tr><td>Dispute amount</td><td>Dispute Adjustments: debit − credit</td></tr>
      <tr><td>Settlement Total</td><td>ATM amt + ATM fee&amp;GST + Micro amt + Micro fee&amp;GST + Switching fee&amp;GST − Disputes + BI fee&amp;GST</td></tr></table>`],
    ['Matching rules', `<p>Key = transaction date + cycle. <b>Diff = NTSL Final Settlement Amount − Bank debit.</b></p>
      <table class="rt"><tr><th>Status</th><th>Rule</th></tr>
      <tr><td>MATCHED</td><td>|Diff| ≤ <b>threshold</b> (default ₹1, editable)</td></tr>
      <tr><td>EXCESS DEBIT</td><td>Diff &lt; 0 — bank debited more than NTSL</td></tr>
      <tr><td>DEFICIT</td><td>Diff &gt; 0 — bank debited less than NTSL</td></tr>
      <tr><td>PENDING</td><td>No bank debit found; expected next day (T+1). Expected settlement = txn date + 1</td></tr></table>
      <p>Daily Summary adds status <b>PARTIAL</b> (some cycles of the day still pending).</p>`],
    ['Scope & carry-over', `<ul><li><b>Month</b> filter restricts NTSL to one month; bank debits whose transaction date falls outside the NTSL months are listed under <b>Carry-over</b> (audit trail only, not in totals).</li></ul>`],
    ['Output', `<p><b>Export Excel</b> → <code>NFS_Recon_&lt;period&gt;_&lt;date&gt;.xlsx</code> with sheets NFS_Recon_Cycle, Daily_Summary, Exceptions, Carry_Over. <b>Load Recon Excel</b> re-opens a previously exported file for review without re-running.</p>`],
  ],
  how: [
    ['Steps', `<ol>
      <li>In the <b>NTSL Files</b> card choose <b>Upload Folder</b> (Chrome/Edge) or <b>Multiple Files</b>, or drag files onto the drop area. Add more batches any time.</li>
      <li>Drop or select the <b>Bank Statement</b>.</li>
      <li>Optionally set <b>Month</b>, <b>Threshold (₹)</b> and the <b>Show carry-over</b> tick.</li>
      <li>Click <b>▶ Run Reconciliation</b>.</li>
      <li>Work through the tabs: <b>Cycle Detail</b>, <b>Daily Summary</b>, <b>Differences</b>, <b>Pending</b>, <b>Carry-over</b>, <b>Summary</b>. Use the search, status and date filters; click column headings to sort.</li>
      <li>Click <b>⬇ Export Excel</b> to save the report.</li></ol>`],
    ['Other buttons', `<ul><li><b>↺ Reset</b> clears everything.</li><li><b>📂 Load Recon Excel</b> loads a previously exported NFS_Recon file.</li></ul>`],
    ['Tips', `<ul><li>The KPI cards (Matched, Pending, Differences, totals) update when you re-run with a new threshold.</li>
      <li>Pending items are normal for the latest day; verify them in the next bank statement.</li></ul>`],
  ]
},

/* ───────────────────────────── RUPAY ───────────────────────────── */
rupay: {
  rules: [
    ['Purpose', `<p>Builds the <b>RuPay control sheet</b>: for every DSR cycle it computes gross and net settlement from the DSR and compares the net with the amount actually credited in the bank.</p>`],
    ['Inputs', `<ul>
      <li><b>DSR files</b> (.xls / .xlsx, up to 150). The <b>file name must contain</b> <code>ISSUER_YYYY-MM-DD-C</code> (e.g. <code>…ISSUER_2026-04-01-2.xls</code>); the cycle is C1–C4. Files not matching are ignored.</li>
      <li><b>Bank statement</b> – header row with <code>Narrative</code>/<code>Description</code>, <code>Credit</code>, <code>Debit</code> columns. Narrative must end with <code>_00YYYYMMDDCC</code>. Narratives containing <code>Rev_</code> are skipped.</li></ul>`],
    ['DSR classification', `<ul>
      <li>Sales = <i>Presentment (With Auth)</i> and <i>Offline Presentment (Without Auth)</i>, type <i>Purchase</i>, status not <code>D</code>, settlement debit &gt; 0, split by channel <b>POS</b> / <b>ECOM</b>. BIN <b>818014</b> is reported separately as <b>Gift Card</b> (POS-GC / ECOM-GC).</li>
      <li>Also captured: Refunds, Chargebacks (Raise / Acceptance), Re-presentments, Cash at POS (NPCI Fee Collection), qSPARC Money Load through Cash.</li>
      <li><b>NPCI assessment &amp; processing fees</b> are read <b>directly from the DSR</b> (not recalculated). Processing fees come from DMS Auth Transactions (status A, excluding Balance Update and qSPARC).</li>
      <li><b>Interchange income</b> by channel is read from the DSR.</li>
      <li><b>GST</b> = the <code>INWARD GST</code> row. <b>Net settlement</b> = final net on the DSR grand-total row (exact, no rounding).</li></ul>`],
    ['Gross total', `<p>Gross = −(POS + ECOM + GC + Cash sales) + Refunds + qSPARC − NPCI assessment − NPCI processing − cash interchange + interchange income + Chargebacks − Re-presentments.</p>`],
    ['Matching rules', `<ul>
      <li>Key = date + cycle. Bank amount for a key = Σ(credit − debit) of all entries with that key.</li>
      <li><b>Difference = Net settlement − Actual bank amount</b>. |Diff| &lt; ₹1 → green (counted as zero-diff), otherwise red.</li>
      <li>If a cycle has <b>no bank entry</b> and its DSR net is not zero, the DSR net is used as the “actual” (so it shows zero difference) and the cycle is counted as <b>Unmatched</b>.</li>
      <li>Date of settlement in bank = DSR date + 1.</li></ul>`],
    ['Triple validation (shown in the log)', `<ol>
      <li><b>V1</b> total NPCI fees × 18% ≈ INWARD GST (tolerance ₹0.10).</li>
      <li><b>V2</b> net settlement equals DSR final net.</li>
      <li><b>V3</b> final net is non-zero for cycles that have sales.</li></ol>`],
    ['Output', `<p>One Excel sheet <b>“Rupay Control Sheet_&lt;Mon YY&gt;”</b>, one row per cycle, with grouped headers, colour-coded difference and SUM totals row. File: <code>Rupay_Recon_YYYY-MM.xlsx</code>.</p>`],
  ],
  how: [
    ['Steps', `<ol>
      <li><b>Step 1</b> – click the DSR box and select all DSR files for the month (C1–C4 of every day), or drag them in. A ✓ shows next to each valid file name.</li>
      <li><b>Step 2</b> – select the bank statement.</li>
      <li><b>Step 3</b> – click <b>⚡ Generate Recon Excel</b>.</li>
      <li>Watch the log; the Excel file downloads automatically when finished.</li>
      <li>Check the result tiles: <b>Matched</b>, <b>Unmatched Cycles</b>, <b>Zero-Diff Rows</b>, <b>Diff &gt; ₹1 Rows</b>.</li></ol>`],
    ['Tips', `<ul><li>Red <b>Diff &gt; ₹1</b> rows in the Excel need investigation.</li>
      <li>Warnings in the log (V1/V2/V3) point to cycles whose DSR looks inconsistent.</li>
      <li>This tool produces the Excel directly — there is no on-screen result table.</li></ul>`],
  ]
},

/* ───────────────────────────── UPI ───────────────────────────── */
upi: {
  rules: [
    ['Purpose', `<p>Reconciles UPI settlement: the <b>net of each NTSL cycle</b> (Credit − Debit) against the <b>bank net</b> for that cycle, with three validation passes.</p>`],
    ['Inputs', `<ul>
      <li><b>NTSL files</b> (up to 500) named <code>UPI_NTSLERT{DDMMYY}_{CYCLE}.xls / .xlsx / .xlsm / .xlsb</code>. Only matching names are loaded.</li>
      <li><b>Bank statement</b> – .xlsx / .xls / .xlsm / .xlsb / .csv. First row is the header; columns: A Narrative, B Eroute remarks, C Customer ref, E Debit, F Credit.</li></ul>`],
    ['Cycles & settlement basis', `<table class="rt"><tr><th>Basis</th><th>Cycles</th><th>Bank entry date</th></tr>
      <tr><td>T</td><td>3C – 9C</td><td>Same day</td></tr>
      <tr><td>T+1</td><td>10C, 1C, 2C, DC1, DC2</td><td>NTSL date + 1 day</td></tr></table>
      <p>Full coverage is 12 cycles a day: 10C, 1C–9C, DC1, DC2.</p>`],
    ['NTSL parsing', `<ul>
      <li>Date and cycle from the row-3 header “as on DD-MM-YYYY ( CYCLE hh:mm:ss”, else from the file name. DC01/DC02 are treated as DC1/DC2.</li>
      <li>Description rows are classified: Transaction Amount, Switching Fee (+GST), Approved Fee (+GST), Net Adjusted Amount, Penalty, GST on Penalty, TDS on GST. Total/opening/closing rows and zero rows are skipped.</li>
      <li>For <b>DC1 / DC2</b> cycles only Net Adjusted Amount, Penalty, GST on Penalty and TDS on GST are kept.</li>
      <li><b>NTSL net = total Credit − total Debit</b> for the cycle.</li></ul>`],
    ['Bank narrative patterns', `<ol>
      <li><code>Eroute UPI Setl {CYCLE} {DDMMYY}</code></li>
      <li><code>DC01/DC02 UPI EROUTE TECH PVT LTD_NPCI00YYYYMMDD…</code></li>
      <li><code>…NPCI00YYYYMMDDNN</code> (NN = 2-digit cycle: 01→1C … 10→10C)</li>
      <li><code>DC1/DC2</code> in the narrative + <code>NPCI00YYYYMMDD</code></li>
      <li><code>…NPCI00YYYYMMDDNN</code> with 4-digit MMDD form</li></ol>
      <p>Bank entries that match none are flagged <b>Parse Error</b>.</p>`],
    ['Matching rules', `<p>Match tolerance <b>₹1.00</b>. Bank net = Bank credit − Bank debit; Diff = Bank net − NTSL net.</p>
      <table class="rt"><tr><th>Status</th><th>Rule</th></tr>
      <tr><td>Matched</td><td>|Diff| &lt; ₹1</td></tr>
      <tr><td>Discrepancy</td><td>|Diff| ≥ ₹1 — real reconciling item</td></tr>
      <tr><td>Matched / Discrepancy (Prior Month)</td><td>Bank entry's NTSL date is earlier than the earliest NTSL date loaded, and the cycle was found</td></tr>
      <tr><td>Prior Month Carryover</td><td>Prior-month entry with no NTSL cycle loaded</td></tr>
      <tr><td>Not Found in NTSL</td><td>Current-period bank entry with no NTSL cycle</td></tr>
      <tr><td>Parse Error</td><td>Narrative not recognised</td></tr></table>`],
    ['Triple validation', `<ol>
      <li><b>Pass 1 – Input integrity:</b> files present (warn above 500), name pattern, 12-cycle coverage per date, duplicate names, bank file selected. <i>Errors stop the run.</i></li>
      <li><b>Pass 2 – Data integrity:</b> rows parsed, no zero-value rows, cycle nets re-computed and agree, grand-total cross-check (tolerance ₹1), remarks classification. <i>Errors stop the run.</i></li>
      <li><b>Pass 3 – Recon integrity:</b> all narratives parsed, re-match on a second pass, discrepancy count, bank-net vs NTSL-net grand total, match rate.</li></ol>`],
    ['Output', `<p>Excel with 6 sheets: <b>Working</b>, <b>Settlement Recon</b>, <b>Bank Recon</b>, <b>Daily Summary</b>, <b>Pivot</b>, <b>Insights</b>. File: <code>UPI_Settlement_Recon_&lt;period&gt;.xlsx</code>.</p>`],
  ],
  how: [
    ['Steps', `<ol>
      <li>Click the <b>NTSL</b> box and select the folder with all NTSL files (or use “Select files individually” if the folder picker isn't available).</li>
      <li>Select the <b>Bank Statement</b>.</li>
      <li>Click <b>▶ RUN RECONCILIATION WITH TRIPLE VALIDATION</b>.</li>
      <li>Watch the three validation passes turn ✔ (green), ⚠ (warning) or ✘ (error). Fix errors and run again if a pass fails.</li>
      <li>Read the <b>Reconciliation Summary</b> and the <b>Reconciling Items</b> table (Bank vs NTSL differences).</li>
      <li>Click <b>⬇ Download Excel Report (.xlsx) — 6 Sheets</b>.</li></ol>`],
    ['Tips', `<ul><li>Load a full month so that all 12 cycles of every date are present; missing cycles are reported in Pass 1.</li>
      <li>“Prior Month” rows are normal at the start of a month (T+1 carry-over from last month).</li></ul>`],
  ]
},
/* ───────────────────────────── ADD MONEY TO POOL ───────────────────────────── */
addmoney: {
  rules: [
    ['Purpose', `<p>Checks that every <b>Add Money</b> / <b>Ops Adjustment</b> credit given to a customer in the Ops file has a matching <b>credit in the bank statement</b> (customer pool account). Both sides are matched on the <b>UTR</b> number. This is the logic of the <i>Reconciliation Automation Macro v3</i>, unchanged.</p>`],
    ['Inputs', `<ul>
      <li><b>Ops file</b> – sheet named <code>Ops File</code> (or the first sheet). Required headers: <code>TransactionCategory</code>, <code>TransactionType</code>, <code>UtrNo</code>, <code>TotalTransactionAmount</code>. <code>EnterpriseName</code> is used for the name column. Header names must match exactly (extra spaces at the ends are ignored).</li>
      <li><b>Bank statement</b> – sheet named <code>Bank Statement</code> (or the first sheet). Required headers: <code>Narrative</code>, <code>Credit</code>.</li>
      <li>If a file is missing a required header, that step is skipped and a warning is shown, exactly as the macro does.</li></ul>`],
    ['Step 1 – Ops file (Final UTR Ops, Ops Amount)', `<ul>
      <li>A row counts only if <b>TransactionCategory</b> is <code>ADD MONEY</code>, <code>OPS ADJUSTMENT</code> or <code>OPS ADJUSTEMENT</code> (any capitalisation) <b>and</b> <b>TransactionType</b> is <code>CREDIT</code>. Every other row gets a blank UTR and amount.</li>
      <li>UTRs that start with <code>TRXV</code> are <b>ignored</b> (blank).</li>
      <li>UTRs that start with <code>RE1</code> → <b>Final UTR</b> = the text after <code>RE1</code>.</li>
      <li>Any other UTR → <b>Final UTR</b> = only its digits (all other characters removed).</li>
      <li><b>Ops Amount</b> = TotalTransactionAmount (blank if it is not a number).</li></ul>`],
    ['Step 2 – Bank statement (Final UTR Bank, Bank Amount)', `<p>The narrative decides how the UTR is read (checked in this order):</p>
      <table class="rt"><tr><th>Narrative starts with</th><th>Final UTR</th></tr>
      <tr><td><code>IMPS/</code></td><td>first <b>12 digits</b> found in the narrative</td></tr>
      <tr><td><code>IFT/</code></td><td>first <b>11 digits</b> found in the narrative</td></tr>
      <tr><td><code>INEFT</code> or <code>IRTGS</code></td><td>digits between <code>~EROUTE~</code> and the next <code>~</code></td></tr>
      <tr><td><code>NEFT</code>, <code>RTGS</code> or <code>IMPS</code> (no slash)</td><td>digits between the first <code>/</code> and the next <code>/</code></td></tr>
      <tr><td>anything else (not empty)</td><td>all digits in the narrative</td></tr></table>
      <p><b>Bank Amount</b> = the Credit amount; blank when the credit is zero or empty.</p>`],
    ['Step 3 – Matching rules', `<ul>
      <li>Rows with the <b>same Final UTR are added together</b> on each side (Ops and Bank).</li>
      <li>Every Ops UTR is listed first, in file order; then every Bank UTR that was not matched, in file order.</li></ul>
      <table class="rt"><tr><th>Status</th><th>Rule</th></tr>
      <tr><td>MATCHED</td><td>UTR is on both sides and |Ops − Bank| &lt; ₹0.01</td></tr>
      <tr><td>AMOUNT MISMATCH</td><td>UTR is on both sides but the amounts differ (Difference = Ops − Bank)</td></tr>
      <tr><td>NOT IN BANK</td><td>Ops UTR with no bank entry (Difference = Ops amount)</td></tr>
      <tr><td>NOT IN OPS</td><td>Bank UTR with no Ops entry (Difference = −Bank amount). This includes bank debit rows whose narrative contains digits; they show a bank amount of 0.00</td></tr></table>`],
    ['Summary block', `<p>Counts of the four statuses, <b>Total Ops Amount (Matched)</b>, <b>Total Bank Amount (Matched)</b> and <b>Net Unreconciled Difference</b> (sum of Difference for AMOUNT MISMATCH rows).</p>`],
    ['Output', `<p>Download Excel with three sheets like the macro leaves behind: <b>Ops File</b> and <b>Bank Statement</b> (your data plus the Final UTR and Amount columns, with the first row of each reconciled UTR coloured) and <b>Reconciliation</b> (table + summary). File: <code>Add_Money_Recon_DDMMYY.xlsx</code>.</p>
      <p style="color:#5a6b82">Not part of the macro (so not done here): manual remarks, moving rows, and the “Add Condition” notes in the workbook (removing bank debits, bank-transfer checks).</p>`],
  ],
  how: [
    ['Steps', `<ol>
      <li><b>Step 1</b> – select or drag in the <b>Add Money Ops file</b> (.xlsx / .xls / .xlsm / .csv). You will see the sheet used and the number of data rows, or the headers that are missing.</li>
      <li><b>Step 2</b> – select or drag in the <b>Bank statement</b>.</li>
      <li><b>Step 3</b> – click <b>⚡ Run Reconciliation</b>.</li>
      <li>Read the result: the four status tiles, the matched totals and the table. Use <b>search</b> (UTR or enterprise name) and the <b>status</b> filter to find items.</li>
      <li>Click <b>⬇ Download Recon Excel</b> to save the full workbook.</li></ol>`],
    ['What to look at first', `<ul>
      <li><b>AMOUNT MISMATCH</b> – same UTR, different amounts; the Difference column shows by how much.</li>
      <li><b>NOT IN BANK</b> – credited in Ops but no matching credit in the bank statement.</li>
      <li><b>NOT IN OPS</b> – in the bank statement but not credited through Add Money / Ops Adjustment (internal transfers, other agent loads, etc.).</li></ul>`],
    ['Tips', `<ul>
      <li>Use the raw Ops and bank files; the Final UTR / Amount columns are added for you. If the file already has them, they are recalculated.</li>
      <li>The table on screen shows the first 5,000 rows; the downloaded Excel has all of them.</li>
      <li>Large Ops files (30,000+ rows) are fine; reading and the Excel download can take a few seconds.</li></ul>`],
  ]
}
};
