(function () {
	if (window.__HTML_TABLE_TO_EXCEL_INSTALLED) return;
	window.__HTML_TABLE_TO_EXCEL_INSTALLED = true;

	let lastContextTarget = null;

	const DEBUG = true;
	let lastHighlightedTable = null;

	function log() {
		if (!DEBUG) return;
		try { console.log('[Table2Excel]', ...arguments); } catch (e) {}
	}

	function ensureHighlightStyle() {
		if (document.getElementById('htmltabletoexcel-highlight-style')) return;
		const style = document.createElement('style');
		style.id = 'htmltabletoexcel-highlight-style';
		style.textContent = '.htmltabletoexcel-highlight{outline:3px solid #ff4d4f !important;outline-offset:2px !important;}';
		document.documentElement.appendChild(style);
	}

	function highlightTable(tableElement) {
		try {
			ensureHighlightStyle();
			if (lastHighlightedTable && lastHighlightedTable !== tableElement) {
				lastHighlightedTable.classList.remove('htmltabletoexcel-highlight');
			}
			tableElement.classList.add('htmltabletoexcel-highlight');
			if (typeof tableElement.scrollIntoView === 'function') {
				tableElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
			}
			lastHighlightedTable = tableElement;
			setTimeout(() => {
				try { tableElement.classList.remove('htmltabletoexcel-highlight'); } catch (e) {}
			}, 1500);
		} catch (e) {
			log('highlight error', e);
		}
	}

	function ensureUiStyle() {
		if (document.getElementById('htmltabletoexcel-ui-style')) return;
		const style = document.createElement('style');
		style.id = 'htmltabletoexcel-ui-style';
		style.textContent = [
			'.htmltabletoexcel-overlay{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:2147483646;display:flex;align-items:center;justify-content:center}',
			'.htmltabletoexcel-dialog{background:#fff;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,.2);width:min(900px,92vw);max-height:86vh;display:flex;flex-direction:column;overflow:hidden;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif}',
			'.htmltabletoexcel-header{padding:12px 16px;border-bottom:1px solid #f0f0f0;display:flex;align-items:center;justify-content:space-between}',
			'.htmltabletoexcel-title{font-size:16px;font-weight:600}',
			'.htmltabletoexcel-body{display:flex;gap:12px;padding:12px;overflow:auto}',
			'.htmltabletoexcel-side{flex:0 0 240px;display:flex;flex-direction:column;gap:10px}',
			'.htmltabletoexcel-section{border:1px solid #f0f0f0;border-radius:6px;padding:10px}',
			'.htmltabletoexcel-section h4{margin:0 0 8px 0;font-size:13px}',
			'.htmltabletoexcel-cols{display:flex;flex-direction:column;gap:6px;max-height:260px;overflow:auto}',
			'.htmltabletoexcel-preview{flex:1 1 auto;border:1px solid #f0f0f0;border-radius:6px;padding:8px;overflow:auto}',
			'.htmltabletoexcel-preview table{border-collapse:collapse;width:100%}',
			'.htmltabletoexcel-preview th,.htmltabletoexcel-preview td{border:1px solid #ddd;padding:6px 8px;font-size:12px}',
			'.htmltabletoexcel-footer{padding:12px 16px;border-top:1px solid #f0f0f0;display:flex;justify-content:flex-end;gap:8px}',
			'.htmltabletoexcel-btn{appearance:none;border:1px solid #d9d9d9;background:#fff;border-radius:6px;padding:6px 12px;font-size:13px;cursor:pointer}',
			'.htmltabletoexcel-btn.primary{background:#1677ff;border-color:#1677ff;color:#fff}',
			'.htmltabletoexcel-toast{position:fixed;left:50%;transform:translateX(-50%);bottom:24px;z-index:2147483647;background:rgba(0,0,0,.85);color:#fff;padding:8px 12px;border-radius:6px;font-size:12px;max-width:80vw;white-space:pre-wrap}'
		].join('');
		document.documentElement.appendChild(style);
	}

	function showToast(message, durationMs) {
		try {
			ensureUiStyle();
			const el = document.createElement('div');
			el.className = 'htmltabletoexcel-toast';
			el.textContent = String(message || '');
			document.body.appendChild(el);
			setTimeout(() => { try { el.remove(); } catch (e) {} }, Math.max(1000, durationMs || 2000));
		} catch (e) {}
	}

	function cloneTableWithSelectedColumns(table, selectedIndexes) {
		const idxSet = new Set(selectedIndexes || []);
		const clone = document.createElement('table');
		clone.setAttribute('border', '1');
		clone.style.borderCollapse = 'collapse';
		// clone colgroup if present and filter cols
		const cg = table.querySelector('colgroup');
		if (cg) {
			const newCg = document.createElement('colgroup');
			const cols = Array.from(cg.children || []);
			for (let i = 0; i < cols.length; i++) {
				if (idxSet.has(i)) newCg.appendChild(cols[i].cloneNode(true));
			}
			if (newCg.children.length) clone.appendChild(newCg);
		}
		const processSection = (section) => {
			if (!section) return null;
			const newSection = document.createElement(section.tagName.toLowerCase());
			for (const row of Array.from(section.rows || [])) {
				const newRow = document.createElement('tr');
				const cells = Array.from(row.cells || []);
				for (let i = 0; i < cells.length; i++) {
					if (!idxSet.has(i)) continue;
					newRow.appendChild(cells[i].cloneNode(true));
				}
				if (newRow.cells.length > 0) newSection.appendChild(newRow);
			}
			return newSection;
		};
		const thead = processSection(table.querySelector('thead'));
		if (thead) clone.appendChild(thead);
		const tbodyElts = table.tBodies && table.tBodies.length ? Array.from(table.tBodies) : [table.querySelector('tbody')].filter(Boolean);
		for (const tb of tbodyElts) {
			const newTb = processSection(tb);
			if (newTb && newTb.rows.length) clone.appendChild(newTb);
		}
		return clone;
	}

	function buildDialogForSingle(table, onConfirm, onCancel) {
		ensureUiStyle();
		const overlay = document.createElement('div');
		overlay.className = 'htmltabletoexcel-overlay';
		const dialog = document.createElement('div');
		dialog.className = 'htmltabletoexcel-dialog';

		const header = document.createElement('div');
		header.className = 'htmltabletoexcel-header';
		const title = document.createElement('div');
		title.className = 'htmltabletoexcel-title';
		title.textContent = '导出表格';
		header.appendChild(title);
		const closeBtn = document.createElement('button');
		closeBtn.className = 'htmltabletoexcel-btn';
		closeBtn.textContent = '关闭';
		closeBtn.onclick = () => { try { overlay.remove(); } catch (e) {} onCancel && onCancel(); };
		header.appendChild(closeBtn);

		const body = document.createElement('div');
		body.className = 'htmltabletoexcel-body';
		const side = document.createElement('div');
		side.className = 'htmltabletoexcel-side';
		const secFormat = document.createElement('div');
		secFormat.className = 'htmltabletoexcel-section';
		secFormat.innerHTML = '<h4>导出格式</h4>'
			+ '<label><input type="radio" name="hte-format" value="xlsx" checked> XLSX</label><br>'
			+ '<label><input type="radio" name="hte-format" value="csv"> CSV</label><br>'
			+ '<label><input type="radio" name="hte-format" value="json"> JSON</label>';
		side.appendChild(secFormat);

		const secCols = document.createElement('div');
		secCols.className = 'htmltabletoexcel-section';
		const colsWrap = document.createElement('div');
		colsWrap.className = 'htmltabletoexcel-cols';
		const headerRow = table.querySelector('thead tr') || table.querySelector('tr');
		const colCount = headerRow ? headerRow.cells.length : 0;
		secCols.innerHTML = '<h4>选择列</h4>';
		for (let i = 0; i < colCount; i++) {
			const label = document.createElement('label');
			const cb = document.createElement('input');
			cb.type = 'checkbox';
			cb.value = String(i);
			cb.checked = true;
			label.appendChild(cb);
			label.appendChild(document.createTextNode(' 第' + (i + 1) + '列'));
			colsWrap.appendChild(label);
		}
		secCols.appendChild(colsWrap);
		side.appendChild(secCols);

		const preview = document.createElement('div');
		preview.className = 'htmltabletoexcel-preview';
		const previewTable = table.cloneNode(true);
		preview.appendChild(previewTable);

		const updatePreview = () => {
			const selected = Array.from(colsWrap.querySelectorAll('input[type="checkbox"]')).filter(c => c.checked).map(c => Number(c.value));
			const filtered = cloneTableWithSelectedColumns(table, selected);
			preview.innerHTML = '';
			preview.appendChild(filtered);
		};
		colsWrap.addEventListener('change', updatePreview);

		body.appendChild(side);
		body.appendChild(preview);

		const footer = document.createElement('div');
		footer.className = 'htmltabletoexcel-footer';
		const cancelBtn = document.createElement('button');
		cancelBtn.className = 'htmltabletoexcel-btn';
		cancelBtn.textContent = '取消';
		cancelBtn.onclick = () => { try { overlay.remove(); } catch (e) {} onCancel && onCancel(); };
		const okBtn = document.createElement('button');
		okBtn.className = 'htmltabletoexcel-btn primary';
		okBtn.textContent = '导出';
		okBtn.onclick = () => {
			const format = (secFormat.querySelector('input[name="hte-format"]:checked') || {}).value || 'xlsx';
			const selected = Array.from(colsWrap.querySelectorAll('input[type="checkbox"]')).filter(c => c.checked).map(c => Number(c.value));
			try { overlay.remove(); } catch (e) {}
			onConfirm && onConfirm({ format, selectedColumns: selected });
		};
		footer.appendChild(cancelBtn);
		footer.appendChild(okBtn);

		dialog.appendChild(header);
		dialog.appendChild(body);
		dialog.appendChild(footer);
		overlay.appendChild(dialog);
		document.body.appendChild(overlay);
		updatePreview();
	}

	function buildDialogForBatch(onConfirm, onCancel) {
		ensureUiStyle();
		const overlay = document.createElement('div');
		overlay.className = 'htmltabletoexcel-overlay';
		const dialog = document.createElement('div');
		dialog.className = 'htmltabletoexcel-dialog';
		const header = document.createElement('div');
		header.className = 'htmltabletoexcel-header';
		const title = document.createElement('div');
		title.className = 'htmltabletoexcel-title';
		title.textContent = '批量导出 - 选择格式';
		header.appendChild(title);
		const closeBtn = document.createElement('button');
		closeBtn.className = 'htmltabletoexcel-btn';
		closeBtn.textContent = '关闭';
		closeBtn.onclick = () => { try { overlay.remove(); } catch (e) {} onCancel && onCancel(); };
		header.appendChild(closeBtn);

		const body = document.createElement('div');
		body.className = 'htmltabletoexcel-body';
		const side = document.createElement('div');
		side.className = 'htmltabletoexcel-side';
		const secFormat = document.createElement('div');
		secFormat.className = 'htmltabletoexcel-section';
		secFormat.innerHTML = '<h4>导出格式</h4>'
			+ '<label><input type="radio" name="hte-format-batch" value="xlsx" checked> XLSX</label><br>'
			+ '<label><input type="radio" name="hte-format-batch" value="csv"> CSV</label><br>'
			+ '<label><input type="radio" name="hte-format-batch" value="json"> JSON</label>';
		side.appendChild(secFormat);
		body.appendChild(side);

		const footer = document.createElement('div');
		footer.className = 'htmltabletoexcel-footer';
		const cancelBtn = document.createElement('button');
		cancelBtn.className = 'htmltabletoexcel-btn';
		cancelBtn.textContent = '取消';
		cancelBtn.onclick = () => { try { overlay.remove(); } catch (e) {} onCancel && onCancel(); };
		const okBtn = document.createElement('button');
		okBtn.className = 'htmltabletoexcel-btn primary';
		okBtn.textContent = '开始导出';
		okBtn.onclick = () => {
			const format = (secFormat.querySelector('input[name="hte-format-batch"]:checked') || {}).value || 'xlsx';
			try { overlay.remove(); } catch (e) {}
			onConfirm && onConfirm({ format });
		};
		footer.appendChild(cancelBtn);
		footer.appendChild(okBtn);

		dialog.appendChild(header);
		dialog.appendChild(body);
		dialog.appendChild(footer);
		overlay.appendChild(dialog);
		document.body.appendChild(overlay);
	}

	document.addEventListener(
		"contextmenu",
		function (event) {
			lastContextTarget = event.target;
			try { log('contextmenu target', lastContextTarget); } catch (e) {}
		},
		true
	);

	function findNearestTable(startElement) {
		if (!startElement || !(startElement instanceof Element)) return null;
		if (typeof startElement.closest === "function") {
			const found = startElement.closest("table");
			if (found) return found;
		}
		let current = startElement;
		while (current && current !== document.body) {
			if (current.tagName && current.tagName.toLowerCase() === "table") return current;
			current = current.parentElement;
		}
		return null;
	}

	function getTableStats(table) {
		const thead = table.querySelector('thead');
		const tbodies = Array.from(table.tBodies || []);
		const theadRows = thead ? thead.rows.length : 0;
		let tbodyRows = 0;
		for (const t of tbodies) tbodyRows += (t && t.rows ? t.rows.length : 0);
		return { theadRows, tbodyRows };
	}

	function getColumnCount(table) {
		const headRow = table.querySelector('thead tr');
		const row = headRow || table.querySelector('tbody tr');
		if (!row) return 0;
		let count = 0;
		for (const cell of Array.from(row.cells || [])) {
			const span = Number(cell.colSpan || 1);
			count += isNaN(span) ? 1 : span;
		}
		return count;
	}

	function findRelatedBodyTable(headerTable) {
		const headerCols = getColumnCount(headerTable);
		let ancestor = headerTable.parentElement;
		for (let depth = 0; depth < 3 && ancestor; depth++, ancestor = ancestor.parentElement) {
			const tables = Array.from(ancestor.querySelectorAll('table'));
			for (const t of tables) {
				if (t === headerTable) continue;
				const { theadRows, tbodyRows } = getTableStats(t);
				if (tbodyRows > 0) {
					const cols = getColumnCount(t);
					if (headerCols === 0 || cols === 0 || Math.abs(cols - headerCols) <= 1) {
						return t;
					}
				}
			}
		}
		return null;
	}

	function buildUnifiedTableForExport(baseTable) {
		const stats = getTableStats(baseTable);
		if (stats.tbodyRows > 0) {
			return { elementToExport: baseTable, highlightElement: baseTable, merged: false };
		}
		const bodyTable = findRelatedBodyTable(baseTable);
		if (!bodyTable) {
			return { elementToExport: baseTable, highlightElement: baseTable, merged: false };
		}
		const exportTable = document.createElement('table');
		exportTable.setAttribute('border', '1');
		exportTable.style.borderCollapse = 'collapse';

		const bodyColgroup = bodyTable.querySelector('colgroup');
		const headColgroup = baseTable.querySelector('colgroup');
		const colgroup = bodyColgroup || headColgroup;
		if (colgroup) exportTable.appendChild(colgroup.cloneNode(true));

		const thead = baseTable.querySelector('thead');
		if (thead) exportTable.appendChild(thead.cloneNode(true));

		const tbody = document.createElement('tbody');
		for (const tb of Array.from(bodyTable.tBodies || [])) {
			for (const row of Array.from(tb.rows || [])) {
				const clonedRow = row.cloneNode(true);
				tbody.appendChild(clonedRow);
			}
		}
		exportTable.appendChild(tbody);

		log('merged header/body tables for export');
		return { elementToExport: exportTable, highlightElement: bodyTable, merged: true };
	}

	function sanitizeHtml(html) {
		return html
			.replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, "")
			.replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, "");
	}

	function generateFileName(ext) {
		const base = (document.title || "table").trim().replace(/[\/\\:*?"<>|]/g, "_");
		const timestamp = new Date()
			.toISOString()
			.replace(/[-:]/g, "")
			.replace("T", "_")
			.replace(/\..+/, "");
		return `${base}_${timestamp}.${ext || 'xlsx'}`;
	}

	// ===== Minimal XLSX generator (ZIP with stored method) =====
	function escapeXml(value) {
		return String(value)
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;')
			.replace(/'/g, '&apos;');
	}

	function escapeForFormula(value) {
		return String(value).replace(/"/g, '""');
	}

	function numberToColumnName(index) {
		let name = '';
		let i = index;
		while (i >= 0) {
			name = String.fromCharCode((i % 26) + 65) + name;
			i = Math.floor(i / 26) - 1;
		}
		return name;
	}

	function toAbsoluteUrl(href) {
		try { return new URL(href, location.href).toString(); } catch (e) { return href || ''; }
	}

	function parseCellContent(cell) {
		const baseText = (cell.innerText || cell.textContent || '').replace(/\s+/g, ' ').trim();
		const anchors = Array.from(cell.querySelectorAll('a[href]'));
		const images = Array.from(cell.querySelectorAll('img[src]'));

		// If single clean hyperlink equals base text, emit as HYPERLINK formula
		if (anchors.length === 1) {
			const a = anchors[0];
			const linkText = (a.textContent || '').replace(/\s+/g, ' ').trim();
			const href = toAbsoluteUrl(a.getAttribute('href'));
			if (href && (baseText === linkText || baseText === '' || baseText === href)) {
				const display = linkText || href;
				const formula = `HYPERLINK("${escapeForFormula(href)}","${escapeForFormula(display)}")`;
				return { formula };
			}
		}

		let text = baseText;
		// Append hyperlinks as text suffix if present
		if (anchors.length > 0) {
			const urls = anchors.map(a => toAbsoluteUrl(a.getAttribute('href'))).filter(Boolean);
			if (urls.length) {
				text = text ? `${text} (${urls.join(', ')})` : urls.join(', ');
			}
		}
		// Append images info
		if (images.length > 0) {
			const parts = images.map(img => {
				const alt = (img.getAttribute('alt') || img.getAttribute('title') || '').trim();
				const src = toAbsoluteUrl(img.getAttribute('src') || '');
				return alt ? `img:${alt}` : (src ? `img:${src}` : 'img');
			});
			text = text ? `${text} [${parts.join('; ')}]` : `[${parts.join('; ')}]`;
		}
		return { text };
	}

	function extractTableData(table) {
		const rows = [];
		const rowElements = table.querySelectorAll('tr');
		for (const tr of rowElements) {
			const cells = [];
			for (const cell of Array.from(tr.cells || [])) {
				cells.push(parseCellContent(cell));
			}
			if (cells.length > 0) rows.push(cells);
		}
		return rows;
	}

	function getCellText(cellElement) {
		const baseText = (cellElement.innerText || cellElement.textContent || '').replace(/\s+/g, ' ').trim();
		const anchors = Array.from(cellElement.querySelectorAll('a[href]'));
		const images = Array.from(cellElement.querySelectorAll('img[src]'));
		let text = baseText;
		if (anchors.length > 0) {
			const urls = anchors.map(a => toAbsoluteUrl(a.getAttribute('href'))).filter(Boolean);
			if (urls.length) {
				text = text ? `${text} (${urls.join(', ')})` : urls.join(', ');
			}
		}
		if (images.length > 0) {
			const parts = images.map(img => {
				const alt = (img.getAttribute('alt') || img.getAttribute('title') || '').trim();
				const src = toAbsoluteUrl(img.getAttribute('src') || '');
				return alt ? `img:${alt}` : (src ? `img:${src}` : 'img');
			});
			text = text ? `${text} [${parts.join('; ')}]` : `[${parts.join('; ')}]`;
		}
		return text;
	}

	function extractPlainRows(table) {
		const rows = [];
		const rowElements = table.querySelectorAll('tr');
		for (const tr of rowElements) {
			const cells = [];
			for (const cell of Array.from(tr.cells || [])) {
				cells.push(getCellText(cell));
			}
			if (cells.length > 0) rows.push(cells);
		}
		return rows;
	}

	function csvEscape(field) {
		const s = String(field == null ? '' : field);
		const needsQuote = /[",\n\r]/.test(s);
		const escaped = s.replace(/"/g, '""');
		return needsQuote ? `"${escaped}"` : escaped;
	}

	function exportTableToCsv(table) {
		try {
			const rows = extractPlainRows(table);
			const csvLines = rows.map(r => r.map(csvEscape).join(','));
			const csv = csvLines.join('\r\n');
			const blob = new Blob(["\uFEFF", csv], { type: 'text/csv;charset=utf-8' });
			const url = URL.createObjectURL(blob);
			const fileName = generateFileName('csv');
			log('exporting as', fileName);
			const a = document.createElement('a');
			a.href = url;
			a.download = fileName;
			document.body.appendChild(a);
			a.click();
			document.body.removeChild(a);
			setTimeout(() => URL.revokeObjectURL(url), 5000);
		} catch (e) {
			log('csv export error', e);
			alert('导出 CSV 失败，请在控制台查看错误信息');
		}
	}

	function exportTableToJson(table) {
		try {
			const rows = extractPlainRows(table);
			if (!rows.length) {
				alert('表格为空，无法导出 JSON');
				return;
			}
			const headers = (rows[0] || []).map((h, i) => {
				const name = String(h || '').trim();
				return name || `field${i + 1}`;
			});
			const dataRows = rows.slice(1);
			const objects = dataRows.map(r => {
				const obj = {};
				for (let i = 0; i < headers.length; i++) {
					obj[headers[i]] = r[i] == null ? '' : r[i];
				}
				return obj;
			});
			const json = JSON.stringify(objects, null, 2);
			const blob = new Blob([json], { type: 'application/json;charset=utf-8' });
			const url = URL.createObjectURL(blob);
			const fileName = generateFileName('json');
			log('exporting as', fileName);
			const a = document.createElement('a');
			a.href = url;
			a.download = fileName;
			document.body.appendChild(a);
			a.click();
			document.body.removeChild(a);
			setTimeout(() => URL.revokeObjectURL(url), 5000);
		} catch (e) {
			log('json export error', e);
			alert('导出 JSON 失败，请在控制台查看错误信息');
		}
	}

	function buildSheetXml(data) {
		let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
		xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		xml += '<sheetData>';
		for (let r = 0; r < data.length; r++) {
			const row = data[r];
			xml += `<row r="${r + 1}">`;
			for (let c = 0; c < row.length; c++) {
				const colName = numberToColumnName(c);
				const cellRef = `${colName}${r + 1}`;
				const cell = row[c] || {};
				if (cell.formula) {
					xml += `<c r="${cellRef}"><f>${escapeXml(cell.formula)}</f></c>`;
				} else {
					const value = escapeXml(cell.text == null ? '' : cell.text);
					xml += `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${value}</t></is></c>`;
				}
			}
			xml += '</row>';
		}
		xml += '</sheetData>';
		xml += '</worksheet>';
		return xml;
	}

	function buildWorkbookXml() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			+ '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
			+ '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
			+ '</workbook>';
	}

	function buildWorkbookRels() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
			+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
			+ '</Relationships>';
	}

	function buildRootRels() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
			+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
			+ '</Relationships>';
	}

	function buildContentTypes() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			+ '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			+ '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
			+ '<Default Extension="xml" ContentType="application/xml"/>'
			+ '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
			+ '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
			+ '</Types>';
	}

	// CRC32
	const CRC32_TABLE = (() => {
		const table = new Uint32Array(256);
		for (let i = 0; i < 256; i++) {
			let c = i;
			for (let k = 0; k < 8; k++) {
				c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
			}
			table[i] = c >>> 0;
		}
		return table;
	})();

	function crc32(bytes) {
		let c = 0 ^ -1;
		for (let i = 0; i < bytes.length; i++) {
			c = (c >>> 8) ^ CRC32_TABLE[(c ^ bytes[i]) & 0xFF];
		}
		return (c ^ -1) >>> 0;
	}

	function writeUint16LE(arr, value) {
		arr.push(value & 0xFF, (value >>> 8) & 0xFF);
	}
	function writeUint32LE(arr, value) {
		arr.push(value & 0xFF, (value >>> 8) & 0xFF, (value >>> 16) & 0xFF, (value >>> 24) & 0xFF);
	}

	function zipFiles(files) {
		// files: [{ name: string, data: Uint8Array }]
		const encoder = new TextEncoder();
		const localHeaders = [];
		const centralHeaders = [];
		let offset = 0;
		const out = [];

		for (const file of files) {
			const nameBytes = encoder.encode(file.name);
			const data = file.data;
			const crc = crc32(data);
			const compSize = data.length; // stored
			const uncompSize = data.length;
			const gpBitFlag = 1 << 11; // UTF-8 filenames
			const modTime = 0; // we will write zeros
			const modDate = 0;

			// Local file header
			const local = [];
			writeUint32LE(local, 0x04034b50);
			writeUint16LE(local, 20);
			writeUint16LE(local, gpBitFlag);
			writeUint16LE(local, 0); // method = stored
			writeUint16LE(local, modTime);
			writeUint16LE(local, modDate);
			writeUint32LE(local, crc);
			writeUint32LE(local, compSize);
			writeUint32LE(local, uncompSize);
			writeUint16LE(local, nameBytes.length);
			writeUint16LE(local, 0);
			localHeaders.push({ bytes: local, nameBytes, data });

			// Central directory header
			const central = [];
			writeUint32LE(central, 0x02014b50);
			writeUint16LE(central, 20);
			writeUint16LE(central, 20);
			writeUint16LE(central, gpBitFlag);
			writeUint16LE(central, 0);
			writeUint16LE(central, modTime);
			writeUint16LE(central, modDate);
			writeUint32LE(central, crc);
			writeUint32LE(central, compSize);
			writeUint32LE(central, uncompSize);
			writeUint16LE(central, nameBytes.length);
			writeUint16LE(central, 0);
			writeUint16LE(central, 0);
			writeUint16LE(central, 0);
			writeUint16LE(central, 0);
			writeUint32LE(central, 0);
			writeUint32LE(central, offset);
			centralHeaders.push({ bytes: central, nameBytes });

			offset += local.length + nameBytes.length + data.length;
		}

		// Write locals
		for (let i = 0; i < localHeaders.length; i++) {
			const l = localHeaders[i];
			out.push(...l.bytes, ...l.nameBytes, ...l.data);
		}

		const centralOffset = offset;
		for (let i = 0; i < centralHeaders.length; i++) {
			const c = centralHeaders[i];
			out.push(...c.bytes, ...c.nameBytes);
			offset += c.bytes.length + c.nameBytes.length;
		}
		const centralSize = offset - centralOffset;

		// End of central directory
		const end = [];
		writeUint32LE(end, 0x06054b50);
		writeUint16LE(end, 0);
		writeUint16LE(end, 0);
		writeUint16LE(end, files.length);
		writeUint16LE(end, files.length);
		writeUint32LE(end, centralSize);
		writeUint32LE(end, centralOffset);
		writeUint16LE(end, 0);
		out.push(...end);

		return new Uint8Array(out);
	}

	function createXlsxBlobFromTable(table) {
		const zipBytes = createXlsxBytesFromTable(table);
		return new Blob([zipBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
	}
	// ===== End XLSX generator =====

	function createXlsxBytesFromTable(table) {
		const data = extractTableData(table);
		const encoder = new TextEncoder();
		const files = [
			{ name: '[Content_Types].xml', data: encoder.encode(buildContentTypes()) },
			{ name: '_rels/.rels', data: encoder.encode(buildRootRels()) },
			{ name: 'xl/workbook.xml', data: encoder.encode(buildWorkbookXml()) },
			{ name: 'xl/_rels/workbook.xml.rels', data: encoder.encode(buildWorkbookRels()) },
			{ name: 'xl/worksheets/sheet1.xml', data: encoder.encode(buildSheetXml(data)) }
		];
		return zipFiles(files);
	}

	function exportTableToXlsx(table) {
		try {
			const blob = createXlsxBlobFromTable(table);
			const url = URL.createObjectURL(blob);
			const fileName = generateFileName('xlsx');
			log('exporting as', fileName);
			const a = document.createElement('a');
			a.href = url;
			a.download = fileName;
			document.body.appendChild(a);
			a.click();
			document.body.removeChild(a);
			setTimeout(() => URL.revokeObjectURL(url), 5000);
		} catch (e) {
			log('xlsx export error', e);
			alert('导出 XLSX 失败，请在控制台查看错误信息');
		}
	}

	chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
		if (message && message.type === "EXPORT_TABLE_AT_CONTEXT") {
			log('message received: EXPORT_TABLE_AT_CONTEXT');
			let baseTable = null;
			if (lastContextTarget) {
				baseTable = findNearestTable(lastContextTarget);
			}
			if (!baseTable) {
				const allTables = document.querySelectorAll("table");
				baseTable = allTables.length ? allTables[0] : null;
			}
			if (!baseTable) {
				alert("未找到可导出的表格");
				log('no table found');
				sendResponse && sendResponse({ ok: false, error: "NO_TABLE" });
				return;
			}

			const { elementToExport, highlightElement, merged } = buildUnifiedTableForExport(baseTable);
			try {
				const stats = getTableStats(elementToExport);
				log('export table stats', stats, { merged });
			} catch (e) {}
			highlightTable(highlightElement || baseTable);
			const format = (message && message.format) || 'xlsx';
			if (format === 'ask') {
				buildDialogForSingle(elementToExport, ({ format, selectedColumns }) => {
					const filtered = Array.isArray(selectedColumns) ? cloneTableWithSelectedColumns(elementToExport, selectedColumns) : elementToExport;
					try {
						if (format === 'csv') exportTableToCsv(filtered);
						else if (format === 'json') exportTableToJson(filtered);
						else exportTableToXlsx(filtered);
						showToast('导出成功');
						sendResponse && sendResponse({ ok: true });
					} catch (e) {
						showToast('导出失败');
						sendResponse && sendResponse({ ok: false });
					}
				}, () => {
					sendResponse && sendResponse({ ok: false, error: 'CANCELLED' });
				});
				return true;
			} else {
				try {
					if (format === 'csv') exportTableToCsv(elementToExport);
					else if (format === 'json') exportTableToJson(elementToExport);
					else exportTableToXlsx(elementToExport);
					showToast('导出成功');
					sendResponse && sendResponse({ ok: true });
				} catch (e) {
					showToast('导出失败');
					sendResponse && sendResponse({ ok: false });
				}
			}
		}
		if (message && message.type === "EXPORT_ALL_TABLES") {
			try {
				let format = (message && message.format) || 'xlsx';
				if (format === 'ask') {
					buildDialogForBatch(({ format: f }) => {
						chrome.runtime.sendMessage({ type: 'EXPORT_ALL_TABLES', format: f });
					}, () => {});
					sendResponse && sendResponse({ ok: true });
					return true;
				}
				const allTables = Array.from(document.querySelectorAll('table'));
				const seen = new Set();
				const encoder = new TextEncoder();
				const files = [];
				let index = 1;
				for (const table of allTables) {
					const { elementToExport, highlightElement } = buildUnifiedTableForExport(table);
					const key = highlightElement || elementToExport;
					if (seen.has(key)) continue;
					seen.add(key);
					const rows = extractPlainRows(elementToExport);
					if (!rows || rows.length === 0) continue;
					try { highlightTable(highlightElement || table); } catch (e) {}
					let name = '';
					let dataBytes = null;
					if (format === 'csv') {
						const csvLines = rows.map(r => r.map(csvEscape).join(','));
						const csv = csvLines.join('\r\n');
						const withBom = "\uFEFF" + csv;
						dataBytes = encoder.encode(withBom);
						name = `table_${index}.csv`;
					} else if (format === 'json') {
						const headers = (rows[0] || []).map((h, i) => {
							const nm = String(h || '').trim();
							return nm || `field${i + 1}`;
						});
						const dataRows = rows.slice(1);
						const objects = dataRows.map(r => {
							const obj = {};
							for (let i = 0; i < headers.length; i++) obj[headers[i]] = r[i] == null ? '' : r[i];
							return obj;
						});
						const json = JSON.stringify(objects, null, 2);
						dataBytes = encoder.encode(json);
						name = `table_${index}.json`;
					} else {
						const bytes = createXlsxBytesFromTable(elementToExport);
						dataBytes = bytes;
						name = `table_${index}.xlsx`;
					}
					files.push({ name, data: dataBytes });
					index++;
				}
				if (files.length === 0) {
					alert('未找到有效表格');
					sendResponse && sendResponse({ ok: false, error: 'NO_VALID_TABLES' });
					return;
				}
				const zipBytes = zipFiles(files);
				const blob = new Blob([zipBytes], { type: 'application/zip' });
				const url = URL.createObjectURL(blob);
				const base = (document.title || "tables").trim().replace(/[\/\\:*?"<>|]/g, "_");
				const timestamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "_").replace(/\..+/, "");
				const fileName = `${base}_${timestamp}_all_${format}.zip`;
				log('exporting batch as', fileName, 'files:', files.length);
				const a = document.createElement('a');
				a.href = url;
				a.download = fileName;
				document.body.appendChild(a);
				a.click();
				document.body.removeChild(a);
				setTimeout(() => URL.revokeObjectURL(url), 5000);
				showToast('批量导出完成：' + files.length + ' 个文件');
				sendResponse && sendResponse({ ok: true, count: files.length });
			} catch (e) {
				log('batch export error', e);
				showToast('批量导出失败');
				sendResponse && sendResponse({ ok: false, error: 'BATCH_ERROR' });
			}
		}
	});
})();
