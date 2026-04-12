const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('desktopApi', {
	openExcelFiles: () => ipcRenderer.invoke('dialog:openExcelFiles'),
	importFilesByPath: (filePaths) => ipcRenderer.invoke('dialog:importFilesByPath', filePaths),
	getPathForFile: (file) => {
		try {
			return webUtils.getPathForFile(file);
		} catch (_error) {
			return '';
		}
	},
	loadWorkspace: () => ipcRenderer.invoke('workspace:load'),
	saveWorkspace: (workspace) => ipcRenderer.invoke('workspace:save', workspace),
	exportDataset: (dataset) => ipcRenderer.invoke('dataset:export', dataset),
	queryOne: (job) => ipcRenderer.invoke('geocode:queryOne', job),
	openCleanupScriptLibrary: () => ipcRenderer.invoke('cleanupScripts:openWindow'),
	applyCleanupScriptFromLibrary: (scriptId) => ipcRenderer.invoke('cleanupScripts:applyToMain', scriptId),
	notifyCleanupScriptsUpdated: () => ipcRenderer.invoke('cleanupScripts:notifyUpdated'),
	onCleanupScriptApplyRequest: (handler) => {
		if (typeof handler !== 'function') {
			return () => {};
		}

		const listener = (_event, scriptId) => handler(scriptId);
		ipcRenderer.on('cleanupScripts:applyRequest', listener);
		return () => {
			ipcRenderer.removeListener('cleanupScripts:applyRequest', listener);
		};
	},
	onCleanupScriptsUpdated: (handler) => {
		if (typeof handler !== 'function') {
			return () => {};
		}

		const listener = () => handler();
		ipcRenderer.on('cleanupScripts:updated', listener);
		return () => {
			ipcRenderer.removeListener('cleanupScripts:updated', listener);
		};
	},
});
