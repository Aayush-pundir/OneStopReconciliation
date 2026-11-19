# One Stop Reconciliation

A single dashboard: **Recon OS** (upload the bank statement + the month's files/folder/zip → everything reconciled in one click) plus one tab per reconciliation: **FASTag**, **NFS**, **RuPay**, **UPI**, **Add money to Pool**.

**Use it:** open `index.html` in Chrome/Edge. That one file is the whole platform.

## How it works
Each reconciliation tool lives in `src/modules/` exactly as it was built (logic, rules and Excel output untouched)
and runs in its own isolated frame inside its tab. At build time a styling-only stylesheet from `src/themes/` is appended to the
embedded copy so every tool matches the dashboard theme. Isolation means the tools cannot interfere with each other
and behave identically to standalone. A tab's uploaded files and results are kept while you switch to other tabs.

## Layout
- `index.html` – built single-file dashboard (generated, commit it)
- `src/os/` – **Recon OS**: `engine.js` (file detection, bank-statement reading, all matching rules; no screen code), `ui.js`, `export.js` (full Excel report), `controlsheet.js` (Control sheet download: the month's control sheets in the FASTag/NFS/RuPay/UPI control-sheet layouts), `os.html` (page). Assembled into one tab by `build.py`
- `src/shell.html` – dashboard shell (sidebar, per-tab toolbar, help dialog)
- `src/help.js` – text behind each tab's **Rules & parameters** and **How to use** icons (edit here if a tool's rules change)
- `src/modules/*.html` – the individual reconciliation tools (original files, never edited by the build). `addmoney.html` is a port of the Recon-181026.xlsm macro (Module1 v3)
- `src/themes/*.css` – per-tool theme (colours/layout only) injected by the build
- `build.py` – `python3 build.py` regenerates `index.html`

## Changing a tool or adding a new reconciliation
1. Edit/add the file in `src/modules/`.
2. For a new one, add a line to `MODULES` in `build.py`.
3. Run `python3 build.py`.

Note: FASTag, NFS and RuPay load SheetJS/ExcelJS/FileSaver from a CDN (internet needed); UPI embeds its libraries.

## Recon OS in one minute
1. Choose the **bank statement** (one file for all products).
2. Drop the month's **folder, files or zips** (nested zips are fine). Files are recognised by name and content.
3. Click **Reconcile everything**; download the Excel report.

The engine can be tested without a browser: it only needs `XLSX` (SheetJS) and `JSZip` as globals (see `src/os/engine.js`).
3. Click **⬇ Control sheet download** for the month's control sheets (FASTag, NFS, RuPay, UPI) in the layout of the finance team's control-sheet files; its first sheet lists each product's balance (should be ≈ 0).
