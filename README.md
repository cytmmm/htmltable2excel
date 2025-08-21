## HTML Table to Excel (Chrome Extension, MV3)

[中文说明](README.zh-CN.md)

Export HTML tables from any web page to real Excel (.xlsx) with one click. Right‑click near a table to export. Handles split header/body table patterns by merging them before export. All processing is local.

### Features
- Context‑menu action: export the nearest `<table>` at right‑click
- Auto‑merge header/body when rendered as separate tables
- Real `.xlsx` file generated in the browser (no servers, no libs)
- Privacy‑friendly: we do not collect or transmit data

### Install (developer mode)
1. Clone or download this folder.
2. Open Chrome → `chrome://extensions/` → enable Developer Mode.
3. Click “Load unpacked” and select this folder.

### Usage
- Open a page with an HTML table.
- Right‑click near the target table → “导出此处的表格为 Excel”.
- The table is highlighted briefly and an `.xlsx` is downloaded.

### Permissions
- `contextMenus`: register the right‑click item
- `activeTab` + `scripting`: inject the content script only when the user triggers export
- Content scripts are declared for convenience; no background scraping

### Privacy
- See `PRIVACY.md`. No data collection, storage, or sharing. All processing happens in local memory.

### Development
- Background SW: `background.js`
- Content script: `content.js`
- Manifest: `manifest.json`
- Simple icon tool: open `tools/icon_generator.html` to generate PNGs → place into `icons/`

Hot reload flow:
- After edits, refresh the extension in `chrome://extensions/`, then refresh your target page.
- Content‑script logs are prefixed with `[Table2Excel]` in DevTools Console.

### License
MIT — see `LICENSE`.
