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
		id: "export_table_to_excel",
		title: "导出此处的表格为 Excel",
		contexts: ["all"]
	});
});

function sendExportMessageWithFallback(tab, frameId) {
	if (!tab || !tab.id) return;
	const options = typeof frameId === 'number' ? { frameId } : undefined;
	chrome.tabs.sendMessage(tab.id, { type: "EXPORT_TABLE_AT_CONTEXT" }, options, () => {
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
					chrome.tabs.sendMessage(tab.id, { type: "EXPORT_TABLE_AT_CONTEXT" }, options, () => {
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
	if (info.menuItemId !== "export_table_to_excel" || !tab || !tab.id) return;
	if (!isInjectableUrl(tab.url)) {
		console.warn("[Table2Excel] Unsupported page (cannot run content scripts):", tab && tab.url);
		return;
	}
	const frameId = typeof info.frameId === 'number' ? info.frameId : undefined;
	sendExportMessageWithFallback(tab, frameId);
});
