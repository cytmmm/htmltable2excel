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
		const data = extractTableData(table);
		const encoder = new TextEncoder();
		const files = [
			{ name: '[Content_Types].xml', data: encoder.encode(buildContentTypes()) },
			{ name: '_rels/.rels', data: encoder.encode(buildRootRels()) },
			{ name: 'xl/workbook.xml', data: encoder.encode(buildWorkbookXml()) },
			{ name: 'xl/_rels/workbook.xml.rels', data: encoder.encode(buildWorkbookRels()) },
			{ name: 'xl/worksheets/sheet1.xml', data: encoder.encode(buildSheetXml(data)) }
		];
		const zipBytes = zipFiles(files);
		return new Blob([zipBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
	}
	// ===== End XLSX generator =====

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
			exportTableToXlsx(elementToExport);
			sendResponse && sendResponse({ ok: true });
		}
	});
})();
