const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('desktopApi', {
	openExcelFiles: () => ipcRenderer.invoke('dialog:openExcelFiles'),
	loadWorkspace: () => ipcRenderer.invoke('workspace:load'),
	saveWorkspace: (workspace) => ipcRenderer.invoke('workspace:save', workspace),
	exportDataset: (dataset) => ipcRenderer.invoke('dataset:export', dataset),
	queryOne: (job) => ipcRenderer.invoke('geocode:queryOne', job),
});
