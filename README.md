# One Stop Reconciliation

A single dashboard with one tab per reconciliation: **FASTag**, **NFS**, **RuPay**, **UPI**.

**Use it:** open `index.html` in Chrome/Edge. That one file is the whole platform.

## How it works
Each reconciliation tool lives in `src/modules/` exactly as it was built (logic, rules and Excel output untouched)
and runs in its own isolated frame inside its tab. At build time a styling-only stylesheet from `src/themes/` is appended to the
embedded copy so every tool matches the dashboard theme. Isolation means the tools cannot interfere with each other
and behave identically to standalone. A tab's uploaded files and results are kept while you switch to other tabs.

## Layout
- `index.html` – built single-file dashboard (generated, commit it)
- `src/shell.html` – dashboard shell (sidebar, per-tab toolbar, help dialog)
- `src/help.js` – text behind each tab's **Rules & parameters** and **How to use** icons (edit here if a tool's rules change)
- `src/modules/*.html` – the individual reconciliation tools (original files, never edited by the build)
- `src/themes/*.css` – per-tool theme (colours/layout only) injected by the build
- `build.py` – `python3 build.py` regenerates `index.html`

## Changing a tool or adding a new reconciliation
1. Edit/add the file in `src/modules/`.
2. For a new one, add a line to `MODULES` in `build.py`.
3. Run `python3 build.py`.

Note: FASTag, NFS and RuPay load SheetJS/ExcelJS/FileSaver from a CDN (internet needed); UPI embeds its libraries.
