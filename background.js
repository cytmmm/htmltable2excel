/*
Simple background service worker for context menu integration
*/

function isInjectableUrl(url) {
	if (!url) return false;
	if (/^https?:/i.test(url)) return true;
	if (/^file:/i.test(url)) return true; // requires user to enable "Allow access to file URLs"
	return false;
}


chrome.runtime.onInstalled.addListener(() => {
	chrome.contextMenus.create({
		id: "export_table_parent",
		title: "导出此处的表格",
		contexts: ["all"]
	});
	chrome.contextMenus.create({
		id: "export_table_xlsx",
		title: "为 Excel (.xlsx)",
		parentId: "export_table_parent",
		contexts: ["all"]
	});
	chrome.contextMenus.create({
		id: "export_table_csv",
		title: "为 CSV (.csv)",
		parentId: "export_table_parent",
		contexts: ["all"]
	});
	chrome.contextMenus.create({
		id: "export_table_json",
		title: "为 JSON (.json)",
		parentId: "export_table_parent",
		contexts: ["all"]
	});

	chrome.contextMenus.create({
		id: "export_all_tables_parent",
		title: "导出页面所有表格",
		contexts: ["all"]
	});
	chrome.contextMenus.create({
		id: "export_all_tables_xlsx",
		title: "为 Excel (.xlsx)",
		parentId: "export_all_tables_parent",
		contexts: ["all"]
	});
	chrome.contextMenus.create({
		id: "export_all_tables_csv",
		title: "为 CSV (.csv)",
		parentId: "export_all_tables_parent",
		contexts: ["all"]
	});
	chrome.contextMenus.create({
		id: "export_all_tables_json",
		title: "为 JSON (.json)",
		parentId: "export_all_tables_parent",
		contexts: ["all"]
	});
});

function sendExportMessageWithFallback(tab, frameId, format, type) {
	if (!tab || !tab.id) return;
	const options = typeof frameId === 'number' ? { frameId } : undefined;
	const messageType = type || "EXPORT_TABLE_AT_CONTEXT";
	chrome.tabs.sendMessage(tab.id, { type: messageType, format }, options, () => {
		const err = chrome.runtime.lastError;
		if (err) {
			if (!isInjectableUrl(tab.url)) {
				console.warn("[Table2Excel] Cannot inject content script into this URL:", tab && tab.url);
				return;
			}
			const target = typeof frameId === 'number' ? { tabId: tab.id, frameIds: [frameId] } : { tabId: tab.id, allFrames: false };
			chrome.scripting.executeScript(
				{ target, files: ["content.js"] },
				() => {
					const injErr = chrome.runtime.lastError;
					if (injErr) {
						console.warn("[Table2Excel] Failed to inject content script:", injErr.message);
						return;
					}
					chrome.tabs.sendMessage(tab.id, { type: messageType, format }, options, () => {
						const retryErr = chrome.runtime.lastError;
						if (retryErr) {
							console.warn("[Table2Excel] Send message failed after injection:", retryErr.message);
						}
					});
				}
			);
		}
	});
}

chrome.contextMenus.onClicked.addListener((info, tab) => {
	if (!tab || !tab.id) return;
	if (!isInjectableUrl(tab.url)) {
		console.warn("[Table2Excel] Unsupported page (cannot run content scripts):", tab && tab.url);
		return;
	}
	const frameId = typeof info.frameId === 'number' ? info.frameId : undefined;
	let format = 'xlsx';
	let type = 'EXPORT_TABLE_AT_CONTEXT';
	if (info.menuItemId === 'export_table_csv') format = 'csv';
	else if (info.menuItemId === 'export_table_json') format = 'json';
	else if (info.menuItemId === 'export_table_xlsx') format = 'xlsx';
	else if (info.menuItemId === 'export_table_parent') format = 'ask';
	else if (info.menuItemId === 'export_all_tables_xlsx') { format = 'xlsx'; type = 'EXPORT_ALL_TABLES'; }
	else if (info.menuItemId === 'export_all_tables_parent') { format = 'ask'; type = 'EXPORT_ALL_TABLES'; }
	else if (info.menuItemId === 'export_all_tables_csv') { format = 'csv'; type = 'EXPORT_ALL_TABLES'; }
	else if (info.menuItemId === 'export_all_tables_json') { format = 'json'; type = 'EXPORT_ALL_TABLES'; }
	sendExportMessageWithFallback(tab, frameId, format, type);
});
