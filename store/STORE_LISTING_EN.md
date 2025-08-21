# HTML Table to Excel

Export HTML tables from any web page to real Excel (.xlsx) with one click. Right‑click near a table to export. Handles split header/body table patterns used by modern frameworks by merging them before export.

## Short description (max 132 chars)
Export HTML tables to XLSX from context menu. Merges header/body tables. All processing happens locally.

## Full description
HTML Table to Excel helps you quickly export `<table>` elements on web pages to genuine Excel `.xlsx` files.

- Right‑click near the target table → "Export table to Excel"
- Automatically picks the nearest `<table>`; falls back to the first table on the page
- Supports split header/body setups by merging them on the fly
- Files are generated and downloaded locally in your browser; no server involved

### Highlights
- True XLSX: generates OOXML and packages a ZIP in the browser; no external deps
- Smart targeting: pick the nearest table based on the right‑click location
- Header/body merge: works with UI frameworks that render header and body separately
- Privacy‑friendly: we do not collect or transmit any data (see Privacy Policy)

### How to use
1. Install the extension and refresh the page you want to export
2. Move the mouse near the desired table and use the context‑menu action
3. The browser downloads the `.xlsx` file

### Permissions
- contextMenus: to register the right‑click item
- activeTab + scripting: to inject/export only when you invoke the action
- Note: no background scraping; all processing runs locally when you trigger the action

### Privacy
- We do not collect, store, or share user data
- All processing happens in local browser memory
- See `PRIVACY.md` for details

### Known limitations
- Exports plain text cells for now (no rich styles, formulas, or merged cells)
- Virtualized tables export only rendered rows
- Cross‑origin iframes cannot be accessed due to browser security

### Changelog
- 1.0.0: Initial release with context‑menu export, header/body merging, and XLSX output

---
Suggested screenshots: at least one 1280×800 shot showing the context menu and the downloaded result.
