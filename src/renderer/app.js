const HOME_TAB_ID = 'home';

const state = {
	datasets: [],
	activeDatasetId: HOME_TAB_ID,
	fileViewMode: 'list',
	activeSidebarTool: 'cleanup',
	cleanupScripts: [],
	scrollPositions: {
		sidebarTop: 0,
		tabContentTop: 0,
		tabContentLeft: 0,
		homeFileListTop: 0,
		homeFileListLeft: 0,
		composeScrollLeftByDataset: {},
		tableByDataset: {},
	},
};

let composeFormatSerial = 0;
let cleanupScriptSerial = 0;
const workspacePersistence = {
	isHydrating: true,
	saveTimer: null,
	lastSavedSnapshot: '',
};
const datasetHistoryState = {};
let homeDragAbortController = null;

const SYSTEM_COLUMNS = [
	'candidate_address',
	'matched_address',
	'geocode_status',
	'tgos_candidate_address',
	'coord_x',
	'coord_y',
	'coord_system',
	'geocode_match_type',
	'geocode_result_count',
	'geocode_error',
];

const tabs = document.getElementById('tabs');
const tabContent = document.getElementById('tabContent');
const sidebarPanel = document.getElementById('sidebarPanel');
const datasetCount = document.getElementById('datasetCount');
const saveDatasetButton = document.getElementById('saveDatasetButton');
const undoDatasetButton = document.getElementById('undoDatasetButton');
const redoDatasetButton = document.getElementById('redoDatasetButton');
const exportDatasetButton = document.getElementById('exportDatasetButton');
const tabPanelTemplate = document.getElementById('tabPanelTemplate');
const sidebarTemplate = document.getElementById('sidebarTemplate');
const homeSidebarTemplate = document.getElementById('homeSidebarTemplate');
const homePanelTemplate = document.getElementById('homePanelTemplate');
const emptyStateTemplate = document.getElementById('emptyStateTemplate');
const textInputDialogTemplate = document.getElementById('textInputDialogTemplate');
const confirmDialogTemplate = document.getElementById('confirmDialogTemplate');

function escapeHtml(value) {
	return String(value)
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&#39;');
}

function getDatasetById(datasetId) {
	return state.datasets.find((dataset) => dataset.id === datasetId);
}

function formatDatasetLabel(dataset) {
	return `${dataset.fileName} / ${dataset.sheetName}`;
}

function safeString(value) {
	return value === undefined || value === null ? '' : String(value);
}

function isFileDragEvent(event) {
	const types = Array.from(event.dataTransfer?.types || []);
	return types.includes('Files');
}

function promptForText({
	title = '請輸入內容',
	description = '請確認內容後送出。',
	label = '內容',
	defaultValue = '',
	placeholder = '',
	confirmText = '確認',
	cancelText = '取消',
} = {}) {
	if (!textInputDialogTemplate?.content?.firstElementChild) {
		const fallbackValue = window.prompt(title, defaultValue);
		return Promise.resolve(fallbackValue);
	}

	return new Promise((resolve) => {
		const dialogRoot = textInputDialogTemplate.content.firstElementChild.cloneNode(true);
		const titleElement = dialogRoot.querySelector('.dialog-title');
		const descriptionElement = dialogRoot.querySelector('.dialog-description');
		const labelElement = dialogRoot.querySelector('.dialog-field-label');
		const inputElement = dialogRoot.querySelector('.dialog-input');
		const cancelButton = dialogRoot.querySelector('.dialog-cancel-button');
		const confirmButton = dialogRoot.querySelector('.dialog-confirm-button');

		titleElement.textContent = title;
		descriptionElement.textContent = description;
		labelElement.textContent = label;
		inputElement.value = defaultValue;
		inputElement.placeholder = placeholder;
		cancelButton.textContent = cancelText;
		confirmButton.textContent = confirmText;

		const cleanup = (result) => {
			dialogRoot.remove();
			resolve(result);
		};

		cancelButton.addEventListener('click', () => cleanup(null));
		confirmButton.addEventListener('click', () => cleanup(inputElement.value));
		dialogRoot.addEventListener('click', (event) => {
			if (event.target === dialogRoot) {
				cleanup(null);
			}
		});
		inputElement.addEventListener('keydown', (event) => {
			if (event.key === 'Enter') {
				event.preventDefault();
				cleanup(inputElement.value);
				return;
			}

			if (event.key === 'Escape') {
				event.preventDefault();
				cleanup(null);
			}
		});

		document.body.appendChild(dialogRoot);
		inputElement.focus();
		inputElement.select();
	});
}

function confirmAction({
	title = '請再次確認',
	description = '這個操作需要你的確認。',
	confirmText = '確認',
	cancelText = '取消',
} = {}) {
	if (!confirmDialogTemplate?.content?.firstElementChild) {
		return Promise.resolve(window.confirm(description));
	}

	return new Promise((resolve) => {
		const dialogRoot = confirmDialogTemplate.content.firstElementChild.cloneNode(true);
		const titleElement = dialogRoot.querySelector('.dialog-title');
		const descriptionElement = dialogRoot.querySelector('.dialog-description');
		const cancelButton = dialogRoot.querySelector('.dialog-cancel-button');
		const confirmButton = dialogRoot.querySelector('.dialog-confirm-button');

		titleElement.textContent = title;
		descriptionElement.textContent = description;
		cancelButton.textContent = cancelText;
		confirmButton.textContent = confirmText;

		const cleanup = (result) => {
			dialogRoot.remove();
			resolve(result);
		};

		cancelButton.addEventListener('click', () => cleanup(false));
		confirmButton.addEventListener('click', () => cleanup(true));
		dialogRoot.addEventListener('click', (event) => {
			if (event.target === dialogRoot) {
				cleanup(false);
			}
		});
		dialogRoot.addEventListener('keydown', (event) => {
			if (event.key === 'Escape') {
				event.preventDefault();
				cleanup(false);
			}
		});

		document.body.appendChild(dialogRoot);
		confirmButton.focus();
	});
}

function safeComposeFormatId(value) {
	const normalized = safeString(value).trim();
	if (normalized) {
		return normalized;
	}

	composeFormatSerial += 1;
	return `compose-format-${Date.now()}-${composeFormatSerial}`;
}

function createComposeFormat(name, segments = []) {
	return {
		id: safeComposeFormatId(''),
		name,
		segments: Array.isArray(segments) ? segments.map((segment) => ({
			type: segment?.type === 'field' ? 'field' : 'text',
			value: safeString(segment?.value),
		})) : [],
	};
}

function safeCleanupScriptId(value) {
	const normalized = safeString(value).trim();
	if (normalized) {
		return normalized;
	}

	cleanupScriptSerial += 1;
	return `cleanup-script-${Date.now()}-${cleanupScriptSerial}`;
}

function cloneCleanupOptions(options = {}) {
	return {
		trim: options.trim !== undefined ? Boolean(options.trim) : true,
		collapseSpace: options.collapseSpace !== undefined ? Boolean(options.collapseSpace) : true,
		fullwidthSpace: options.fullwidthSpace !== undefined ? Boolean(options.fullwidthSpace) : true,
		fullwidthChar: Boolean(options.fullwidthChar),
		removeLinebreak: Boolean(options.removeLinebreak),
	};
}

function cloneCleanupRegexRules(regexRules = []) {
	return Array.isArray(regexRules)
		? regexRules.map((rule) => ({
			pattern: safeString(rule?.pattern),
			flags: safeString(rule?.flags),
			replacement: safeString(rule?.replacement),
		}))
		: [];
}

function createCleanupScript(name, config = {}) {
	const now = new Date().toISOString();
	return {
		id: safeCleanupScriptId(config.id),
		name: safeString(name).trim() || '未命名腳本',
		cleanupOptions: cloneCleanupOptions(config.cleanupOptions),
		cleanupRegexRules: cloneCleanupRegexRules(config.cleanupRegexRules),
		createdAt: config.createdAt || now,
		updatedAt: config.updatedAt || now,
	};
}

function normalizeCleanupScripts(scripts) {
	if (!Array.isArray(scripts)) {
		return [];
	}

	return scripts.map((script, index) => createCleanupScript(
		script?.name || `清洗腳本 ${index + 1}`,
		script,
	));
}

function getCleanupScriptById(scriptId) {
	return state.cleanupScripts.find((script) => script.id === scriptId);
}

function getCleanupFlowConfig(dataset, sidebarRoot) {
	const trimCheckbox = sidebarRoot?.querySelector('.cleanup-trim');
	const collapseSpaceCheckbox = sidebarRoot?.querySelector('.cleanup-collapse-space');
	const fullwidthSpaceCheckbox = sidebarRoot?.querySelector('.cleanup-fullwidth-space');
	const fullwidthCharCheckbox = sidebarRoot?.querySelector('.cleanup-fullwidth-char');
	const removeLinebreakCheckbox = sidebarRoot?.querySelector('.cleanup-remove-linebreak');

	return {
		cleanupOptions: cloneCleanupOptions({
			trim: trimCheckbox ? trimCheckbox.checked : dataset.cleanupOptions.trim,
			collapseSpace: collapseSpaceCheckbox ? collapseSpaceCheckbox.checked : dataset.cleanupOptions.collapseSpace,
			fullwidthSpace: fullwidthSpaceCheckbox ? fullwidthSpaceCheckbox.checked : dataset.cleanupOptions.fullwidthSpace,
			fullwidthChar: fullwidthCharCheckbox ? fullwidthCharCheckbox.checked : dataset.cleanupOptions.fullwidthChar,
			removeLinebreak: removeLinebreakCheckbox ? removeLinebreakCheckbox.checked : dataset.cleanupOptions.removeLinebreak,
		}),
		cleanupRegexRules: cloneCleanupRegexRules(dataset.cleanupRegexRules),
	};
}

function saveCleanupScript(name, config) {
	const normalizedName = safeString(name).trim();
	if (!normalizedName) {
		return { ok: false, reason: 'empty-name' };
	}

	const existing = state.cleanupScripts.find((script) => script.name === normalizedName);
	const now = new Date().toISOString();
	const nextScript = createCleanupScript(normalizedName, config);

	if (existing) {
		existing.cleanupOptions = nextScript.cleanupOptions;
		existing.cleanupRegexRules = nextScript.cleanupRegexRules;
		existing.updatedAt = now;
		return { ok: true, script: existing, replaced: true };
	}

	nextScript.createdAt = now;
	nextScript.updatedAt = now;
	state.cleanupScripts.unshift(nextScript);
	return { ok: true, script: nextScript, replaced: false };
}

function applyCleanupScriptToDataset(dataset, script) {
	ensureCleanupState(dataset);
	dataset.cleanupOptions = cloneCleanupOptions(script.cleanupOptions);
	dataset.cleanupRegexRules = cloneCleanupRegexRules(script.cleanupRegexRules);
}

function ensureSystemColumns(dataset) {
	if (!Array.isArray(dataset.sourceColumnNames)) {
		dataset.sourceColumnNames = dataset.columnNames.filter((name) => name !== '__rowId');
	}

	for (const column of SYSTEM_COLUMNS) {
		if (!dataset.columnNames.includes(column)) {
			dataset.columnNames.push(column);
			for (const row of dataset.rows) {
				row[column] = row[column] || '';
			}
		}
	}

	const nonSystemColumns = dataset.columnNames.filter((name) => name !== '__rowId' && !SYSTEM_COLUMNS.includes(name));
	dataset.columnNames = ['__rowId', ...nonSystemColumns, ...SYSTEM_COLUMNS];
}

function ensureCleanupState(dataset) {
	if (!Array.isArray(dataset.cleanupSelectedColumns)) {
		dataset.cleanupSelectedColumns = [];
	}

	if (!dataset.cleanupOptions) {
		dataset.cleanupOptions = {
			trim: true,
			collapseSpace: true,
			fullwidthSpace: true,
			fullwidthChar: false,
			removeLinebreak: false,
		};
	}

	if (!Array.isArray(dataset.cleanupRegexRules)) {
		dataset.cleanupRegexRules = [];
	}
}

function ensureColumnVisibilityState(dataset) {
	if (!Array.isArray(dataset.hiddenColumns)) {
		dataset.hiddenColumns = [];
	}

	dataset.hiddenColumns = dataset.hiddenColumns.filter((column) => dataset.columnNames.includes(column));
}

function ensureSortState(dataset) {
	if (!dataset.sortState) {
		dataset.sortState = {
			columnName: '',
			direction: '',
		};
	}

	if (!dataset.columnNames.includes(dataset.sortState.columnName)) {
		dataset.sortState.columnName = '';
		dataset.sortState.direction = '';
	}
}

function ensureComposeState(dataset) {
	if (Array.isArray(dataset.composeSegments)) {
		dataset.composeFormats = [
			createComposeFormat('主要格式', dataset.composeSegments),
		];
		delete dataset.composeSegments;
	}

	if (!Array.isArray(dataset.composeFormats) || dataset.composeFormats.length === 0) {
		dataset.composeFormats = [createComposeFormat('主要格式')];
	}

	dataset.composeFormats = dataset.composeFormats.map((format, index) => ({
		id: safeComposeFormatId(format?.id),
		name: safeString(format?.name).trim() || (index === 0 ? '主要格式' : `備選格式 ${index}`),
		segments: Array.isArray(format?.segments) ? format.segments : [],
	}));

	if (!dataset.activeComposeFormatId || !dataset.composeFormats.some((format) => format.id === dataset.activeComposeFormatId)) {
		dataset.activeComposeFormatId = dataset.composeFormats[0].id;
	}
}

function ensureRowSelectionState(dataset) {
	if (!Array.isArray(dataset.selectedRowIds)) {
		dataset.selectedRowIds = [];
	}

	if (!dataset.lastSelectedRowId) {
		dataset.lastSelectedRowId = '';
	}
}

function getActiveComposeFormat(dataset) {
	ensureComposeState(dataset);
	return dataset.composeFormats.find((format) => format.id === dataset.activeComposeFormatId) || dataset.composeFormats[0];
}

function buildCandidateAddressFromSegments(row, segments) {
	return segments.map((segment) => {
		if (segment.type === 'field') {
			const value = row[segment.value];
			return value === undefined || value === null ? '' : String(value);
		}

		return segment.value || '';
	}).join('');
}

function buildGeocodeCandidateAddresses(dataset, row) {
	ensureComposeState(dataset);
	const seen = new Set();
	const addresses = [];

	for (const format of dataset.composeFormats) {
		const candidateAddress = buildCandidateAddressFromSegments(row, format.segments).trim();
		if (!candidateAddress || seen.has(candidateAddress)) {
			continue;
		}

		seen.add(candidateAddress);
		addresses.push(candidateAddress);
	}

	if (addresses.length === 0) {
		const fallbackAddress = safeString(row.candidate_address).trim();
		if (fallbackAddress) {
			addresses.push(fallbackAddress);
		}
	}

	return addresses;
}

function ensureGeocodeState(dataset) {
	if (!dataset.geocodeRuntime) {
		dataset.geocodeRuntime = {
			isRunning: false,
			isPaused: false,
			currentRowId: '',
			pendingRowIds: [],
			totalCount: 0,
			completedCount: 0,
			loopToken: 0,
			scope: 'all',
			forceReprocess: false,
			reprocessUnmatched: false,
		};
	}

	return dataset.geocodeRuntime;
}

function createSerializableGeocodeRuntime(runtime = {}) {
	return {
		isRunning: false,
		isPaused: false,
		currentRowId: '',
		pendingRowIds: [],
		totalCount: 0,
		completedCount: 0,
		loopToken: 0,
		scope: runtime.scope || 'all',
		forceReprocess: Boolean(runtime.forceReprocess),
		reprocessUnmatched: Boolean(runtime.reprocessUnmatched),
	};
}

function createSerializableDataset(dataset) {
	return {
		id: dataset.id,
		fileName: dataset.fileName,
		filePath: dataset.filePath || '',
		sheetName: dataset.sheetName,
		columnNames: Array.isArray(dataset.columnNames) ? [...dataset.columnNames] : [],
		sourceColumnNames: Array.isArray(dataset.sourceColumnNames) ? [...dataset.sourceColumnNames] : [],
		rows: Array.isArray(dataset.rows) ? dataset.rows.map((row) => ({ ...row })) : [],
		importedAt: dataset.importedAt || '',
		cleanupSelectedColumns: Array.isArray(dataset.cleanupSelectedColumns) ? [...dataset.cleanupSelectedColumns] : [],
		cleanupOptions: dataset.cleanupOptions ? { ...dataset.cleanupOptions } : undefined,
		cleanupRegexRules: Array.isArray(dataset.cleanupRegexRules) ? dataset.cleanupRegexRules.map((rule) => ({ ...rule })) : [],
		hiddenColumns: Array.isArray(dataset.hiddenColumns) ? [...dataset.hiddenColumns] : [],
		sortState: dataset.sortState ? { ...dataset.sortState } : undefined,
		composeFormats: Array.isArray(dataset.composeFormats) ? dataset.composeFormats.map((format) => ({
			id: format.id,
			name: format.name,
			segments: Array.isArray(format.segments) ? format.segments.map((segment) => ({ ...segment })) : [],
		})) : undefined,
		activeComposeFormatId: dataset.activeComposeFormatId || '',
		selectedRowIds: Array.isArray(dataset.selectedRowIds) ? [...dataset.selectedRowIds] : [],
		lastSelectedRowId: dataset.lastSelectedRowId || '',
		geocodeRuntime: createSerializableGeocodeRuntime(dataset.geocodeRuntime),
	};
}

function createExportPayload(dataset) {
	return {
		fileName: dataset.fileName,
		sheetName: dataset.sheetName,
		columnNames: Array.isArray(dataset.columnNames) ? [...dataset.columnNames] : [],
		rows: Array.isArray(dataset.rows) ? dataset.rows.map((row) => ({ ...row })) : [],
	};
}

function createDatasetHistoryPayload(dataset) {
	const payload = createSerializableDataset(dataset);
	delete payload.geocodeRuntime;
	return payload;
}

function getDatasetHistorySnapshot(dataset) {
	return JSON.stringify(createDatasetHistoryPayload(dataset));
}

function ensureDatasetHistory(dataset) {
	if (!datasetHistoryState[dataset.id]) {
		const snapshot = getDatasetHistorySnapshot(dataset);
		datasetHistoryState[dataset.id] = {
			undoStack: [],
			redoStack: [],
			presentSnapshot: snapshot,
			savedSnapshot: snapshot,
		};
	}

	return datasetHistoryState[dataset.id];
}

function initializeDatasetHistory(dataset, options = {}) {
	const snapshot = getDatasetHistorySnapshot(dataset);
	datasetHistoryState[dataset.id] = {
		undoStack: [],
		redoStack: [],
		presentSnapshot: snapshot,
		savedSnapshot: options.markSaved === false ? '' : snapshot,
	};
}

function recordDatasetUndoPoint(dataset) {
	const history = ensureDatasetHistory(dataset);
	const currentSnapshot = getDatasetHistorySnapshot(dataset);
	history.presentSnapshot = currentSnapshot;
	if (history.undoStack[history.undoStack.length - 1] !== currentSnapshot) {
		history.undoStack.push(currentSnapshot);
		if (history.undoStack.length > 100) {
			history.undoStack.shift();
		}
	}
	history.redoStack = [];
}

function finalizeDatasetHistory(dataset) {
	const history = ensureDatasetHistory(dataset);
	history.presentSnapshot = getDatasetHistorySnapshot(dataset);
}

function restoreDatasetSnapshot(datasetId, snapshot) {
	const datasetIndex = state.datasets.findIndex((dataset) => dataset.id === datasetId);
	if (datasetIndex < 0) {
		return null;
	}

	const restored = JSON.parse(snapshot);
	ensureSystemColumns(restored);
	ensureCleanupState(restored);
	ensureColumnVisibilityState(restored);
	ensureSortState(restored);
	ensureComposeState(restored);
	ensureRowSelectionState(restored);
	restored.geocodeRuntime = createSerializableGeocodeRuntime(state.datasets[datasetIndex].geocodeRuntime);
	state.datasets[datasetIndex] = restored;
	return restored;
}

function undoDataset(datasetId) {
	const dataset = getDatasetById(datasetId);
	if (!dataset) {
		return;
	}

	const history = ensureDatasetHistory(dataset);
	if (history.undoStack.length === 0) {
		return;
	}

	const currentSnapshot = getDatasetHistorySnapshot(dataset);
	history.redoStack.push(currentSnapshot);
	const previousSnapshot = history.undoStack.pop();
	if (!restoreDatasetSnapshot(datasetId, previousSnapshot)) {
		return;
	}

	history.presentSnapshot = previousSnapshot;
	render();
}

function redoDataset(datasetId) {
	const dataset = getDatasetById(datasetId);
	if (!dataset) {
		return;
	}

	const history = ensureDatasetHistory(dataset);
	if (history.redoStack.length === 0) {
		return;
	}

	const currentSnapshot = getDatasetHistorySnapshot(dataset);
	history.undoStack.push(currentSnapshot);
	const nextSnapshot = history.redoStack.pop();
	if (!restoreDatasetSnapshot(datasetId, nextSnapshot)) {
		return;
	}

	history.presentSnapshot = nextSnapshot;
	render();
}

async function saveCurrentDataset(datasetId) {
	const dataset = getDatasetById(datasetId);
	if (!dataset) {
		return;
	}

	const workspace = serializeWorkspaceState();
	const result = await window.desktopApi.saveWorkspace(workspace);
	if (!result?.ok) {
		window.alert(`存檔失敗：${result?.error || '未知錯誤'}`);
		return;
	}

	workspacePersistence.lastSavedSnapshot = JSON.stringify(workspace);
	const history = ensureDatasetHistory(dataset);
	history.savedSnapshot = getDatasetHistorySnapshot(dataset);
	history.presentSnapshot = history.savedSnapshot;
	window.alert(`已儲存「${formatDatasetLabel(dataset)}」`);
	render();
}

function serializeWorkspaceState() {
	return {
		version: 2,
		activeDatasetId: state.activeDatasetId,
		fileViewMode: state.fileViewMode,
		activeSidebarTool: state.activeSidebarTool,
		cleanupScripts: state.cleanupScripts.map((script) => createCleanupScript(script.name, script)),
		datasets: state.datasets.map(createSerializableDataset),
	};
}

function normalizeLoadedWorkspace(workspace) {
	if (!workspace || typeof workspace !== 'object') {
		return null;
	}

	const datasets = Array.isArray(workspace.datasets) ? workspace.datasets : [];
	for (const dataset of datasets) {
		ensureSystemColumns(dataset);
		ensureCleanupState(dataset);
		ensureColumnVisibilityState(dataset);
		ensureSortState(dataset);
		ensureComposeState(dataset);
		ensureRowSelectionState(dataset);
		dataset.geocodeRuntime = createSerializableGeocodeRuntime(dataset.geocodeRuntime);
		initializeDatasetHistory(dataset);
	}

	return {
		activeDatasetId: workspace.activeDatasetId || HOME_TAB_ID,
		fileViewMode: workspace.fileViewMode === 'grid' ? 'grid' : 'list',
		activeSidebarTool: workspace.activeSidebarTool || 'cleanup',
		cleanupScripts: normalizeCleanupScripts(workspace.cleanupScripts),
		datasets,
	};
}

function scheduleWorkspaceSave() {
	if (workspacePersistence.isHydrating) {
		return;
	}

	const snapshot = JSON.stringify(serializeWorkspaceState());
	if (snapshot === workspacePersistence.lastSavedSnapshot) {
		return;
	}

	if (workspacePersistence.saveTimer) {
		window.clearTimeout(workspacePersistence.saveTimer);
	}

	workspacePersistence.saveTimer = window.setTimeout(async () => {
		const latestSnapshot = JSON.stringify(serializeWorkspaceState());
		if (latestSnapshot === workspacePersistence.lastSavedSnapshot) {
			return;
		}

		const result = await window.desktopApi.saveWorkspace(JSON.parse(latestSnapshot));
		if (result?.ok) {
			workspacePersistence.lastSavedSnapshot = latestSnapshot;
		}
	}, 500);
}

function updateSelectedRows(dataset, rowId, event) {
	ensureRowSelectionState(dataset);
	const rowIds = dataset.rows.map((row) => row.__rowId);
	const selectedRowSet = new Set(dataset.selectedRowIds);
	const isRangeSelect = event.shiftKey && dataset.lastSelectedRowId;
	const isToggleSelect = event.ctrlKey || event.metaKey;

	if (isRangeSelect) {
		const anchorIndex = rowIds.indexOf(dataset.lastSelectedRowId);
		const targetIndex = rowIds.indexOf(rowId);
		if (anchorIndex >= 0 && targetIndex >= 0) {
			const [start, end] = anchorIndex < targetIndex ? [anchorIndex, targetIndex] : [targetIndex, anchorIndex];
			if (!isToggleSelect) {
				selectedRowSet.clear();
			}

			for (let index = start; index <= end; index += 1) {
				selectedRowSet.add(rowIds[index]);
			}
		}
	} else if (isToggleSelect) {
		if (selectedRowSet.has(rowId)) {
			selectedRowSet.delete(rowId);
		} else {
			selectedRowSet.add(rowId);
		}
	} else {
		selectedRowSet.clear();
		selectedRowSet.add(rowId);
	}

	dataset.selectedRowIds = rowIds.filter((currentRowId) => selectedRowSet.has(currentRowId));
	dataset.lastSelectedRowId = rowId;
}

function getGeocodeTargetRows(dataset, scope, forceReprocess, reprocessUnmatched) {
	ensureRowSelectionState(dataset);
	const selectedRowIds = new Set(dataset.selectedRowIds);
	const baseRows = scope === 'selected'
		? dataset.rows.filter((row) => selectedRowIds.has(row.__rowId))
		: dataset.rows;

	return baseRows.filter((row) => {
		const status = row.geocode_status || '';
		if (!status || status === 'error') {
			return true;
		}

		if (status === 'success') {
			return forceReprocess;
		}

		return reprocessUnmatched;
	});
}

function toggleAllRowsSelected(dataset) {
	ensureRowSelectionState(dataset);

	if (dataset.selectedRowIds.length === dataset.rows.length) {
		dataset.selectedRowIds = [];
		dataset.lastSelectedRowId = '';
		return;
	}

	dataset.selectedRowIds = dataset.rows.map((row) => row.__rowId);
	dataset.lastSelectedRowId = dataset.selectedRowIds[dataset.selectedRowIds.length - 1] || '';
}

function invertRowSelection(dataset) {
	ensureRowSelectionState(dataset);
	const selectedRowSet = new Set(dataset.selectedRowIds);
	dataset.selectedRowIds = dataset.rows
		.map((row) => row.__rowId)
		.filter((rowId) => !selectedRowSet.has(rowId));
	dataset.lastSelectedRowId = dataset.selectedRowIds[dataset.selectedRowIds.length - 1] || '';
}

function toggleColumnHidden(dataset, columnName) {
	ensureColumnVisibilityState(dataset);
	const index = dataset.hiddenColumns.indexOf(columnName);
	if (index >= 0) {
		dataset.hiddenColumns.splice(index, 1);
		return;
	}

	dataset.hiddenColumns.push(columnName);
}

function setSortState(dataset, columnName, direction) {
	ensureSortState(dataset);
	dataset.sortState.columnName = columnName;
	dataset.sortState.direction = direction;
}

function toggleSortState(dataset, columnName) {
	ensureSortState(dataset);
	if (dataset.sortState.columnName !== columnName || dataset.sortState.direction === '') {
		setSortState(dataset, columnName, 'asc');
		return;
	}

	if (dataset.sortState.direction === 'asc') {
		setSortState(dataset, columnName, 'desc');
		return;
	}

	setSortState(dataset, '', '');
}

function compareCellValues(leftValue, rightValue) {
	const left = safeString(leftValue).trim();
	const right = safeString(rightValue).trim();

	if (!left && !right) {
		return 0;
	}

	if (!left) {
		return 1;
	}

	if (!right) {
		return -1;
	}

	const leftNumber = Number(left);
	const rightNumber = Number(right);
	if (!Number.isNaN(leftNumber) && !Number.isNaN(rightNumber)) {
		return leftNumber - rightNumber;
	}

	return left.localeCompare(right, 'zh-Hant', { numeric: true, sensitivity: 'base' });
}

function getSortedRows(dataset) {
	ensureSortState(dataset);
	const rows = [...dataset.rows];
	if (!dataset.sortState.columnName || !dataset.sortState.direction) {
		return rows;
	}

	const multiplier = dataset.sortState.direction === 'desc' ? -1 : 1;
	return rows.sort((leftRow, rightRow) => {
		const result = compareCellValues(leftRow[dataset.sortState.columnName], rightRow[dataset.sortState.columnName]);
		if (result !== 0) {
			return result * multiplier;
		}

		return compareCellValues(leftRow.__rowId, rightRow.__rowId);
	});
}

function moveItem(array, fromIndex, toIndex) {
	if (toIndex < 0 || toIndex >= array.length) {
		return;
	}

	const [item] = array.splice(fromIndex, 1);
	array.splice(toIndex, 0, item);
}

function captureScrollPositions() {
	state.scrollPositions.sidebarTop = sidebarPanel.scrollTop;
	state.scrollPositions.tabContentTop = tabContent.scrollTop;
	state.scrollPositions.tabContentLeft = tabContent.scrollLeft;

	if (state.activeDatasetId === HOME_TAB_ID) {
		const homeList = tabContent.querySelector('.file-browser-wrap');
		if (homeList) {
			state.scrollPositions.homeFileListTop = homeList.scrollTop;
			state.scrollPositions.homeFileListLeft = homeList.scrollLeft;
		}
		return;
	}

	const tableWrap = tabContent.querySelector('.table-wrap');
	if (tableWrap) {
		state.scrollPositions.tableByDataset[state.activeDatasetId] = {
			top: tableWrap.scrollTop,
			left: tableWrap.scrollLeft,
		};
	}

	const composeList = sidebarPanel.querySelector('.compose-segment-list');
	if (composeList) {
		state.scrollPositions.composeScrollLeftByDataset[state.activeDatasetId] = composeList.scrollLeft;
	}
}

function restoreHomeScrollPositions() {
	sidebarPanel.scrollTop = state.scrollPositions.sidebarTop;
	tabContent.scrollTop = state.scrollPositions.tabContentTop;
	tabContent.scrollLeft = state.scrollPositions.tabContentLeft;

	const homeList = tabContent.querySelector('.file-browser-wrap');
	if (homeList) {
		homeList.scrollTop = state.scrollPositions.homeFileListTop;
		homeList.scrollLeft = state.scrollPositions.homeFileListLeft;
	}
}

function restoreDatasetScrollPositions(datasetId) {
	sidebarPanel.scrollTop = state.scrollPositions.sidebarTop;
	tabContent.scrollTop = state.scrollPositions.tabContentTop;
	tabContent.scrollLeft = state.scrollPositions.tabContentLeft;

	const tableWrap = tabContent.querySelector('.table-wrap');
	const tablePosition = state.scrollPositions.tableByDataset[datasetId];
	if (tableWrap && tablePosition) {
		tableWrap.scrollTop = tablePosition.top;
		tableWrap.scrollLeft = tablePosition.left;
	}

	const composeList = sidebarPanel.querySelector('.compose-segment-list');
	const composeScrollLeft = state.scrollPositions.composeScrollLeftByDataset[datasetId];
	if (composeList && typeof composeScrollLeft === 'number') {
		composeList.scrollLeft = composeScrollLeft;
	}
}

function blurWorkspaceFocus() {
	const activeElement = document.activeElement;
	if (!(activeElement instanceof HTMLElement)) {
		return;
	}

	if (sidebarPanel.contains(activeElement) || tabContent.contains(activeElement) || tabs.contains(activeElement)) {
		activeElement.blur();
	}
}

function restoreActivePanelScrollPositions() {
	if (state.activeDatasetId === HOME_TAB_ID) {
		restoreHomeScrollPositions();
		return;
	}

	const dataset = getDatasetById(state.activeDatasetId);
	if (!dataset) {
		restoreHomeScrollPositions();
		return;
	}

	restoreDatasetScrollPositions(dataset.id);
}

function scheduleScrollRestore() {
	restoreActivePanelScrollPositions();

	window.requestAnimationFrame(() => {
		restoreActivePanelScrollPositions();
		window.requestAnimationFrame(() => {
			restoreActivePanelScrollPositions();
		});
	});
}

function swapComposeSegment(dataset, fromIndex, toIndex) {
	const activeFormat = getActiveComposeFormat(dataset);
	if (!activeFormat || !Array.isArray(activeFormat.segments)) {
		return;
	}

	if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0) {
		return;
	}

	moveItem(activeFormat.segments, fromIndex, toIndex);
}

function renderTabs() {
	const homeTab = `
		<button class="tab-button ${state.activeDatasetId === HOME_TAB_ID ? 'active' : ''}" data-dataset-id="${HOME_TAB_ID}">
			檔案中心
		</button>
	`;

	const datasetTabs = state.datasets.map((dataset) => `
		<button class="tab-button ${dataset.id === state.activeDatasetId ? 'active' : ''}" data-dataset-id="${dataset.id}">
			${escapeHtml(formatDatasetLabel(dataset))}
		</button>
	`).join('');

	tabs.innerHTML = homeTab + datasetTabs;

	for (const button of tabs.querySelectorAll('.tab-button')) {
		button.addEventListener('click', () => {
			setActiveDataset(button.dataset.datasetId);
		});
	}
}

function renderTable(tableElement, dataset) {
	ensureColumnVisibilityState(dataset);
	ensureSortState(dataset);
	const headers = dataset.columnNames.filter((header) => header !== '__rowId' && !dataset.hiddenColumns.includes(header));
	const selectedColumns = new Set(dataset.cleanupSelectedColumns || []);
	ensureRowSelectionState(dataset);
	const selectedRowIds = new Set(dataset.selectedRowIds);
	const isAllRowsSelected = dataset.rows.length > 0 && dataset.selectedRowIds.length === dataset.rows.length;
	const geocodeState = ensureGeocodeState(dataset);
	const pendingRowIds = new Set(geocodeState.pendingRowIds || []);
	const sortedRows = getSortedRows(dataset);

	tableElement.innerHTML = `
		<thead>
			<tr>
				<th class="row-selector-header">
					<div class="row-selector-header-actions">
						<button
							class="row-selector-button row-selector-toggle ${isAllRowsSelected ? 'is-selected' : ''}"
							type="button"
							aria-label="${isAllRowsSelected ? '全不選列' : '全選列'}"
							title="${isAllRowsSelected ? '全不選列' : '全選列'}"
						></button>
						<button
							class="row-selector-invert-button"
							type="button"
							aria-label="反選列"
							title="反選列"
						>反</button>
					</div>
				</th>
				${headers.map((header) => {
				const isSelected = selectedColumns.has(header);
				const isSortedAsc = dataset.sortState.columnName === header && dataset.sortState.direction === 'asc';
				const isSortedDesc = dataset.sortState.columnName === header && dataset.sortState.direction === 'desc';
				const sortLabel = isSortedAsc ? '↑' : isSortedDesc ? '↓' : '↕';
				const sortTitle = isSortedAsc ? '目前升冪，按一下改成降冪' : isSortedDesc ? '目前降冪，按一下改成升冪' : '排序';
				return `
					<th class="${isSelected ? 'selected-column' : ''}" data-column-name="${escapeHtml(header)}">
						<div class="column-header">
							<div class="column-header-main">
								<input type="checkbox" class="column-picker" data-column-name="${escapeHtml(header)}" ${isSelected ? 'checked' : ''}>
								<span>${escapeHtml(header)}</span>
							</div>
							<div class="column-header-actions">
								<button
									class="column-sort-button ${(isSortedAsc || isSortedDesc) ? 'is-active' : ''}"
									type="button"
									data-column-name="${escapeHtml(header)}"
									title="${escapeHtml(sortTitle)}"
									aria-label="${escapeHtml(sortTitle)}"
								>${sortLabel}</button>
								<button
									class="column-visibility-button"
									type="button"
									data-column-name="${escapeHtml(header)}"
									title="隱藏欄位"
									aria-label="隱藏欄位"
								>👁</button>
							</div>
						</div>
					</th>
				`;
			}).join('')}
			</tr>
		</thead>
		<tbody>
			${sortedRows.map((row) => `
			<tr class="${selectedRowIds.has(row.__rowId) ? 'selected-row' : ''}" data-row-id="${escapeHtml(row.__rowId)}">
					<td class="row-selector-cell">
						<button
							class="row-selector-button ${selectedRowIds.has(row.__rowId) ? 'is-selected' : ''}"
							type="button"
							data-row-id="${escapeHtml(row.__rowId)}"
							aria-label="選取這一列"
						></button>
					</td>
					${headers.map((header) => {
						const selectedClass = selectedColumns.has(header) ? 'selected-column' : '';
						const geocodeClass = header === 'candidate_address'
							? geocodeState.currentRowId === row.__rowId
								? ' geocode-cell geocode-cell-running'
								: pendingRowIds.has(row.__rowId)
									? ' geocode-cell geocode-cell-pending'
									: ''
							: '';
						return `<td class="${selectedClass}${geocodeClass} editable-cell" data-row-id="${escapeHtml(row.__rowId)}" data-column-name="${escapeHtml(header)}">${renderTableCellContent(row, header)}</td>`;
					}).join('')}
				</tr>
			`).join('')}
		</tbody>
	`;

	for (const checkbox of tableElement.querySelectorAll('.column-picker')) {
		checkbox.addEventListener('change', () => {
			recordDatasetUndoPoint(dataset);
			toggleCleanupColumn(dataset, checkbox.dataset.columnName);
			finalizeDatasetHistory(dataset);
			render();
		});
	}

	for (const button of tableElement.querySelectorAll('.column-visibility-button')) {
		button.addEventListener('click', () => {
			recordDatasetUndoPoint(dataset);
			toggleColumnHidden(dataset, button.dataset.columnName);
			finalizeDatasetHistory(dataset);
			render();
		});
	}

	for (const button of tableElement.querySelectorAll('.column-sort-button')) {
		button.addEventListener('click', () => {
			recordDatasetUndoPoint(dataset);
			toggleSortState(dataset, button.dataset.columnName);
			finalizeDatasetHistory(dataset);
			render();
		});
	}

	for (const button of tableElement.querySelectorAll('.row-selector-button')) {
		button.addEventListener('click', (event) => {
			if (button.classList.contains('row-selector-toggle')) {
				toggleAllRowsSelected(dataset);
				render();
				return;
			}

			updateSelectedRows(dataset, button.dataset.rowId, event);
			render();
		});
	}

	for (const button of tableElement.querySelectorAll('.row-selector-invert-button')) {
		button.addEventListener('click', () => {
			invertRowSelection(dataset);
			render();
		});
	}

	for (const cell of tableElement.querySelectorAll('.editable-cell')) {
		cell.addEventListener('click', () => {
			if (cell.querySelector('.table-cell-editor')) {
				return;
			}

			const row = dataset.rows.find((item) => item.__rowId === cell.dataset.rowId);
			const columnName = cell.dataset.columnName;
			if (!row || !columnName) {
				return;
			}

			const originalValue = safeString(row[columnName]);
			cell.classList.add('is-editing');
			cell.innerHTML = `<input class="table-cell-editor" type="text" value="${escapeHtml(originalValue)}">`;
			const input = cell.querySelector('.table-cell-editor');
			if (!input) {
				return;
			}

			input.focus();
			input.select();

			const finishEdit = (shouldSave) => {
				if (!cell.isConnected) {
					return;
				}

				if (shouldSave) {
					recordDatasetUndoPoint(dataset);
					row[columnName] = input.value;
					finalizeDatasetHistory(dataset);
					render();
					return;
				}

				cell.classList.remove('is-editing');
				cell.textContent = originalValue;
			};

			input.addEventListener('keydown', (event) => {
				if (event.key === 'Enter') {
					event.preventDefault();
					finishEdit(true);
				}

				if (event.key === 'Escape') {
					event.preventDefault();
					finishEdit(false);
				}
			});

			input.addEventListener('blur', () => {
				finishEdit(true);
			});
		});
	}
}

function hydrateColumnSelectors(rootElement, dataset) {
	const selectableColumns = dataset.columnNames.filter((name) => name !== '__rowId');

	for (const selector of rootElement.querySelectorAll('.column-selector')) {
		selector.innerHTML = selectableColumns
			.map((column) => `<option value="${escapeHtml(column)}">${escapeHtml(column)}</option>`)
			.join('');
	}
}

function buildRegex(pattern, flags) {
	try {
		return new RegExp(pattern, flags);
	} catch (_error) {
		return null;
	}
}

function toHalfwidth(value) {
	return value
		.replace(/[\uFF01-\uFF5E]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xFEE0))
		.replace(/\u3000/g, ' ');
}

function applyCleanup(dataset, selectedColumns, cleanupOptions, regexRules) {
	for (const row of dataset.rows) {
		for (const column of selectedColumns) {
			let value = row[column];
			if (value === undefined || value === null) {
				value = '';
			}

			value = String(value);

			if (cleanupOptions.fullwidthSpace) {
				value = value.replaceAll('\u3000', ' ');
			}

			if (cleanupOptions.fullwidthChar) {
				value = toHalfwidth(value);
			}

			if (cleanupOptions.removeLinebreak) {
				value = value.replace(/[\r\n]+/g, ' ');
			}

			if (cleanupOptions.collapseSpace) {
				value = value.replace(/\s+/g, ' ');
			}

			if (cleanupOptions.trim) {
				value = value.trim();
			}

			for (const rule of regexRules) {
				if (!rule.pattern) {
					continue;
				}

				const regex = buildRegex(rule.pattern, rule.flags || '');
				if (!regex) {
					continue;
				}

				value = value.replace(regex, rule.replacement || '');
			}

			row[column] = value;
		}
	}
}

function toggleCleanupColumn(dataset, columnName) {
	ensureCleanupState(dataset);
	const index = dataset.cleanupSelectedColumns.indexOf(columnName);

	if (index >= 0) {
		dataset.cleanupSelectedColumns.splice(index, 1);
		return;
	}

	dataset.cleanupSelectedColumns.push(columnName);
}

function composeCandidateAddress(dataset) {
	ensureSystemColumns(dataset);
	ensureComposeState(dataset);
	const primaryFormat = dataset.composeFormats[0];

	for (const row of dataset.rows) {
		row.candidate_address = primaryFormat ? buildCandidateAddressFromSegments(row, primaryFormat.segments) : '';
	}
}

function composePreviewValue(dataset, format = getActiveComposeFormat(dataset)) {
	ensureComposeState(dataset);

	if (!format || format.segments.length === 0) {
		return '尚未建立候選格式';
	}

	const preview = format.segments.map((segment) => {
		if (segment.type === 'field') {
			return `{{${segment.value}}}`;
		}

		return segment.value || '';
	}).join('');

	return preview || '預覽為空白';
}

function updateComposePreview(previewElement, dataset) {
	if (!previewElement) {
		return;
	}

	previewElement.textContent = composePreviewValue(dataset);
}

function renderStatusPill(status) {
	if (!status) {
		return '';
	}

	return `<span class="status-pill status-${status}">${escapeHtml(status)}</span>`;
}

function renderTableCellContent(row, columnName) {
	if (columnName === 'geocode_status') {
		return renderStatusPill(row[columnName] || '');
	}

	return escapeHtml(row[columnName] || '');
}

function getActiveDatasetTableElement(datasetId = state.activeDatasetId) {
	if (datasetId === HOME_TAB_ID) {
		return null;
	}

	return tabContent.querySelector('.dataset-table');
}

function updateRenderedGeocodeRow(row) {
	const tableElement = getActiveDatasetTableElement();
	if (!tableElement) {
		return;
	}

	for (const cell of tableElement.querySelectorAll(`td[data-row-id="${CSS.escape(row.__rowId)}"][data-column-name]`)) {
		cell.classList.remove('is-editing');
		cell.innerHTML = renderTableCellContent(row, cell.dataset.columnName);
	}
}

function updateRenderedGeocodeHighlight(dataset) {
	const tableElement = getActiveDatasetTableElement(dataset.id);
	if (!tableElement) {
		return;
	}

	const geocodeState = ensureGeocodeState(dataset);
	const pendingRowIds = new Set(geocodeState.pendingRowIds || []);
	for (const cell of tableElement.querySelectorAll('td[data-column-name="candidate_address"]')) {
		const rowId = cell.dataset.rowId;
		cell.classList.remove('geocode-cell', 'geocode-cell-running', 'geocode-cell-pending');
		if (geocodeState.currentRowId === rowId) {
			cell.classList.add('geocode-cell', 'geocode-cell-running');
			continue;
		}

		if (pendingRowIds.has(rowId)) {
			cell.classList.add('geocode-cell', 'geocode-cell-pending');
		}
	}
}

function updateRenderedGeocodeSidebar(dataset) {
	if (state.activeDatasetId !== dataset.id) {
		return;
	}

	const sidebarRoot = sidebarPanel.firstElementChild;
	if (!sidebarRoot) {
		return;
	}

	const geocodeState = ensureGeocodeState(dataset);
	const pendingCount = geocodeState.pendingRowIds.length;
	const selectedCount = Array.isArray(dataset.selectedRowIds) ? dataset.selectedRowIds.length : 0;
	const forceReprocessCheckbox = sidebarRoot.querySelector('.geocode-force-reprocess');
	const reprocessUnmatchedCheckbox = sidebarRoot.querySelector('.geocode-reprocess-unmatched');
	const runGeocodeAllButton = sidebarRoot.querySelector('.run-geocode-all');
	const runGeocodeSelectedButton = sidebarRoot.querySelector('.run-geocode-selected');
	const pauseGeocodeButton = sidebarRoot.querySelector('.pause-geocode');
	const stopGeocodeButton = sidebarRoot.querySelector('.stop-geocode');
	const geocodeQueueHint = sidebarRoot.querySelector('.geocode-queue-hint');
	const progressFill = sidebarRoot.querySelector('.progress-fill');
	const progressText = sidebarRoot.querySelector('.progress-text');
	if (!forceReprocessCheckbox || !reprocessUnmatchedCheckbox || !runGeocodeAllButton || !runGeocodeSelectedButton || !pauseGeocodeButton || !stopGeocodeButton || !geocodeQueueHint || !progressFill || !progressText) {
		return;
	}

	forceReprocessCheckbox.checked = geocodeState.forceReprocess;
	reprocessUnmatchedCheckbox.checked = geocodeState.reprocessUnmatched;
	const pendingAllCount = getGeocodeTargetRows(dataset, 'all', forceReprocessCheckbox.checked, reprocessUnmatchedCheckbox.checked).length;
	const pendingSelectedCount = getGeocodeTargetRows(dataset, 'selected', forceReprocessCheckbox.checked, reprocessUnmatchedCheckbox.checked).length;
	const pendingUnmatchedCount = dataset.rows.filter((row) => {
		const status = row.geocode_status || '';
		return status && status !== 'success' && status !== 'error';
	}).length;
	const progressPercent = geocodeState.totalCount === 0
		? 0
		: Math.round((geocodeState.completedCount / geocodeState.totalCount) * 100);

	progressFill.style.width = `${progressPercent}%`;
	progressText.textContent = geocodeState.isPaused
		? `已暫停 ${geocodeState.completedCount}/${geocodeState.totalCount}`
		: geocodeState.isRunning
			? `查詢中 ${geocodeState.completedCount}/${geocodeState.totalCount}`
			: geocodeState.totalCount > 0
				? `已完成 ${geocodeState.completedCount}/${geocodeState.totalCount}`
				: '尚未開始';
	geocodeQueueHint.textContent = geocodeState.totalCount === 0 && !geocodeState.isRunning && !geocodeState.isPaused
		? `全部可處理 ${pendingAllCount} 筆，已勾選 ${selectedCount} 列，其中可處理 ${pendingSelectedCount} 筆，未匹配成功 ${pendingUnmatchedCount} 筆。`
		: `目前模式：${geocodeState.scope === 'selected' ? '勾選列' : '全部'}，待查詢 ${pendingCount} 筆，進行中欄位會在右側以不同底色標記。`;
	runGeocodeAllButton.textContent = geocodeState.isPaused && geocodeState.scope === 'all' ? '繼續處理全部' : '批次處理全部';
	runGeocodeSelectedButton.textContent = geocodeState.isPaused && geocodeState.scope === 'selected' ? '繼續處理勾選列' : '批次處理勾選列';
	runGeocodeAllButton.disabled = geocodeState.isRunning;
	runGeocodeSelectedButton.disabled = geocodeState.isRunning || (!geocodeState.isPaused && selectedCount === 0);
	pauseGeocodeButton.disabled = !geocodeState.isRunning;
	stopGeocodeButton.disabled = !geocodeState.isRunning && !geocodeState.isPaused;
}

function updateRenderedGeocodeState(dataset, row = null) {
	updateRenderedGeocodeHighlight(dataset);
	updateRenderedGeocodeSidebar(dataset);
	if (row) {
		updateRenderedGeocodeRow(row);
	}
}

function sleep(ms) {
	return new Promise((resolve) => setTimeout(resolve, ms));
}

function isCompletedGeocodeRow(row) {
	return ['success', 'skipped', 'no_result'].includes(row.geocode_status || '');
}

function applyGeocodeResult(row, result) {
	row.candidate_address = result.inputCandidateAddress || row.candidate_address || '';
	row.tgos_candidate_address = result.tgosCandidateAddress || '';
	row.matched_address = result.matchedAddress || '';
	row.coord_x = result.x || '';
	row.coord_y = result.y || '';
	row.coord_system = result.coordSystem || '';
	row.geocode_match_type = result.matchType || '';
	row.geocode_result_count = result.resultCount === undefined || result.resultCount === null
		? ''
		: String(result.resultCount);
	row.geocode_status = result.status || '';
	row.geocode_error = result.error || '';
}

async function queryGeocodeWithFallback(dataset, row) {
	const candidateAddresses = buildGeocodeCandidateAddresses(dataset, row);

	if (candidateAddresses.length === 0) {
		return {
			rowId: row.__rowId,
			status: 'skipped',
			inputCandidateAddress: '',
			tgosCandidateAddress: '',
			matchedAddress: '',
			x: '',
			y: '',
			coordSystem: '',
			matchType: '',
			resultCount: 0,
			error: '候選地址為空白',
		};
	}

	let lastResult = null;

	for (const candidateAddress of candidateAddresses) {
		const result = await window.desktopApi.queryOne({
			rowId: row.__rowId,
			candidateAddress,
		});

		lastResult = result;
		if (result.status === 'success') {
			return result;
		}

		if (result.status !== 'no_result') {
			return result;
		}
	}

	return lastResult;
}

async function waitForGeocodeDelay(datasetId, delayMs, loopToken) {
	const stepMs = 50;
	let elapsed = 0;

	while (elapsed < delayMs) {
		await sleep(Math.min(stepMs, delayMs - elapsed));
		elapsed += stepMs;
		const dataset = getDatasetById(datasetId);
		if (!dataset) {
			return false;
		}

		const geocodeState = ensureGeocodeState(dataset);
		if (geocodeState.loopToken !== loopToken || geocodeState.isPaused) {
			return false;
		}
	}

	return true;
}

async function runGeocode(datasetId, nextOptions = null) {
	const dataset = getDatasetById(datasetId);
	if (!dataset) {
		return;
	}

	ensureSystemColumns(dataset);
	ensureRowSelectionState(dataset);
	const geocodeState = ensureGeocodeState(dataset);

	if (geocodeState.isRunning && !geocodeState.isPaused) {
		return;
	}

	if (!geocodeState.isPaused) {
		const scope = nextOptions?.scope || 'all';
		const forceReprocess = Boolean(nextOptions?.forceReprocess);
		const reprocessUnmatched = Boolean(nextOptions?.reprocessUnmatched);
		const pendingRows = getGeocodeTargetRows(dataset, scope, forceReprocess, reprocessUnmatched);
		if (pendingRows.length > 0) {
			recordDatasetUndoPoint(dataset);
		}
		geocodeState.pendingRowIds = pendingRows.map((row) => row.__rowId);
		geocodeState.totalCount = geocodeState.pendingRowIds.length;
		geocodeState.completedCount = 0;
		geocodeState.scope = scope;
		geocodeState.forceReprocess = forceReprocess;
		geocodeState.reprocessUnmatched = reprocessUnmatched;
	}

	if (geocodeState.pendingRowIds.length === 0) {
		geocodeState.isRunning = false;
		geocodeState.isPaused = false;
		geocodeState.currentRowId = '';
		render();
		return;
	}

	geocodeState.isRunning = true;
	geocodeState.isPaused = false;
	geocodeState.loopToken += 1;
	const loopToken = geocodeState.loopToken;
	updateRenderedGeocodeState(dataset);

	while (true) {
		const currentDataset = getDatasetById(datasetId);
		if (!currentDataset) {
			return;
		}

		const currentState = ensureGeocodeState(currentDataset);
		if (currentState.loopToken !== loopToken || currentState.isPaused) {
			currentState.isRunning = false;
			currentState.currentRowId = '';
			updateRenderedGeocodeState(currentDataset);
			return;
		}

		const nextRowId = currentState.pendingRowIds[0];
		if (!nextRowId) {
			currentState.isRunning = false;
			currentState.isPaused = false;
			currentState.currentRowId = '';
			updateRenderedGeocodeState(currentDataset);
			return;
		}

		currentState.currentRowId = nextRowId;
		updateRenderedGeocodeState(currentDataset);

		const row = currentDataset.rows.find((item) => item.__rowId === nextRowId);
		if (!row) {
			currentState.pendingRowIds.shift();
			continue;
		}

		const result = await queryGeocodeWithFallback(currentDataset, row);
		const latestDatasetAfterQuery = getDatasetById(datasetId);
		if (!latestDatasetAfterQuery) {
			return;
		}

		const latestStateAfterQuery = ensureGeocodeState(latestDatasetAfterQuery);
		if (latestStateAfterQuery.loopToken !== loopToken || latestStateAfterQuery.isPaused) {
			latestStateAfterQuery.isRunning = false;
			latestStateAfterQuery.currentRowId = '';
			updateRenderedGeocodeState(latestDatasetAfterQuery);
			return;
		}

		applyGeocodeResult(row, result);
		currentState.pendingRowIds.shift();
		currentState.completedCount += 1;
		currentState.currentRowId = '';
		finalizeDatasetHistory(currentDataset);
		updateRenderedGeocodeState(currentDataset, row);

		if (currentState.pendingRowIds.length === 0) {
			currentState.isRunning = false;
			currentState.isPaused = false;
			updateRenderedGeocodeState(currentDataset);
			return;
		}

		const shouldContinue = await waitForGeocodeDelay(datasetId, 300, loopToken);
		if (!shouldContinue) {
			const latestDataset = getDatasetById(datasetId);
			if (!latestDataset) {
				return;
			}

			const latestState = ensureGeocodeState(latestDataset);
			latestState.isRunning = false;
			latestState.currentRowId = '';
			updateRenderedGeocodeState(latestDataset);
			return;
		}
	}
}

function pauseGeocode(datasetId) {
	const dataset = getDatasetById(datasetId);
	if (!dataset) {
		return;
	}

	const geocodeState = ensureGeocodeState(dataset);
	geocodeState.isPaused = true;
	geocodeState.isRunning = false;
	geocodeState.currentRowId = '';
	render();
}

function stopGeocode(datasetId) {
	const dataset = getDatasetById(datasetId);
	if (!dataset) {
		return;
	}

	const geocodeState = ensureGeocodeState(dataset);
	geocodeState.loopToken += 1;
	geocodeState.isPaused = false;
	geocodeState.isRunning = false;
	geocodeState.currentRowId = '';
	geocodeState.pendingRowIds = [];
	geocodeState.totalCount = 0;
	geocodeState.completedCount = 0;
	updateRenderedGeocodeState(dataset);
}

function renderHomeSidebar() {
	sidebarPanel.innerHTML = '';
	sidebarPanel.appendChild(homeSidebarTemplate.content.firstElementChild.cloneNode(true));
}

function renderHomePanel() {
	if (homeDragAbortController) {
		homeDragAbortController.abort();
	}

	homeDragAbortController = new AbortController();
	const { signal } = homeDragAbortController;
	const panel = homePanelTemplate.content.firstElementChild.cloneNode(true);
	const fileList = panel.querySelector('#fileList');
	const fileCountBadge = panel.querySelector('#fileCountBadge');
	const summary = panel.querySelector('.home-summary');
	const viewButtons = panel.querySelectorAll('.view-mode-button');
	const homeImportButton = panel.querySelector('.home-import-button');
	const homeDropzone = panel.querySelector('.home-dropzone');
	const homeDropCard = panel.querySelector('.file-center-card');
	let dragDepth = 0;
	let isHandlingDrop = false;

	function setHomeDropzoneActive(isActive) {
		homeDropzone.classList.toggle('is-drag-over', isActive);
	}

	function extractDroppedFilePaths(event) {
		const files = Array.from(event.dataTransfer?.files || []);
		return files
			.map((file) => window.desktopApi.getPathForFile(file))
			.filter((filePath) => typeof filePath === 'string' && filePath.trim() !== '');
	}

	fileCountBadge.textContent = String(state.datasets.length);
	summary.textContent = `${state.datasets.length} 份資料表`;
	homeImportButton.addEventListener('click', importFiles);

	const handleDragEnter = (event) => {
		if (!isFileDragEvent(event)) {
			return;
		}

		event.preventDefault();
		dragDepth += 1;
		setHomeDropzoneActive(true);
	};

	const handleDragOver = (event) => {
		if (!isFileDragEvent(event)) {
			return;
		}

		event.preventDefault();
		if (event.dataTransfer) {
			event.dataTransfer.dropEffect = 'copy';
		}
		setHomeDropzoneActive(true);
	};

	const handleDragLeave = (event) => {
		if (!isFileDragEvent(event)) {
			return;
		}

		event.preventDefault();
		dragDepth = Math.max(0, dragDepth - 1);
		if (dragDepth === 0 || event.target === homeDropzone || event.target === homeDropCard) {
			setHomeDropzoneActive(false);
		}
	};

	const handleDrop = async (event) => {
		if (!isFileDragEvent(event)) {
			return;
		}

		event.preventDefault();
		event.stopPropagation();
		if (isHandlingDrop) {
			return;
		}

		isHandlingDrop = true;
		dragDepth = 0;
		setHomeDropzoneActive(false);
		try {
			await importFilesFromDrop(extractDroppedFilePaths(event));
		} finally {
			isHandlingDrop = false;
		}
	};

	for (const target of [window, homeDropCard, homeDropzone]) {
		target.addEventListener('dragenter', handleDragEnter, { signal });
		target.addEventListener('dragover', handleDragOver, { signal });
		target.addEventListener('dragleave', handleDragLeave, { signal });
		target.addEventListener('drop', handleDrop, { signal });
	}

	for (const button of viewButtons) {
		button.classList.toggle('active', button.dataset.viewMode === state.fileViewMode);
		button.addEventListener('click', () => {
			state.fileViewMode = button.dataset.viewMode;
			render();
		});
	}

	if (state.datasets.length === 0) {
		fileList.className = 'file-browser empty-state';
		fileList.textContent = '尚未匯入任何檔案';
		tabContent.innerHTML = '';
		tabContent.appendChild(panel);
		return;
	}

	fileList.className = `file-browser ${state.fileViewMode === 'grid' ? 'grid-mode' : 'list-mode'}`;
	fileList.innerHTML = state.datasets.map((dataset) => `
		<article class="file-card" data-dataset-id="${dataset.id}">
			<div class="file-card-main">
				<div class="file-card-icon">${state.fileViewMode === 'grid' ? 'XLS' : '表'}</div>
				<div class="file-card-copy">
					<h4>${escapeHtml(dataset.fileName)}</h4>
					<p>${escapeHtml(dataset.sheetName)}</p>
					<p class="file-meta">${dataset.rows.length} 筆資料</p>
				</div>
			</div>
			<div class="file-card-actions">
				<button class="mini-button open-dataset" data-dataset-id="${dataset.id}">開啟</button>
				<button class="mini-button rename-dataset" data-dataset-id="${dataset.id}">重新命名</button>
				<button class="mini-button danger-button delete-dataset" data-dataset-id="${dataset.id}">刪除</button>
			</div>
		</article>
	`).join('');

	for (const button of fileList.querySelectorAll('.open-dataset')) {
		button.addEventListener('click', () => {
			setActiveDataset(button.dataset.datasetId);
		});
	}

	for (const button of fileList.querySelectorAll('.rename-dataset')) {
		button.addEventListener('click', () => {
			const dataset = getDatasetById(button.dataset.datasetId);
			if (!dataset) {
				return;
			}

			const nextName = window.prompt('請輸入新的顯示名稱', dataset.fileName);
			if (!nextName) {
				return;
			}

			dataset.fileName = nextName.trim() || dataset.fileName;
			render();
		});
	}

	for (const button of fileList.querySelectorAll('.delete-dataset')) {
		button.addEventListener('click', () => {
			const dataset = getDatasetById(button.dataset.datasetId);
			if (!dataset) {
				return;
			}

			const confirmed = window.confirm(`確定要刪除「${formatDatasetLabel(dataset)}」嗎？`);
			if (!confirmed) {
				return;
			}

			state.datasets = state.datasets.filter((item) => item.id !== dataset.id);
			if (state.activeDatasetId === dataset.id) {
				state.activeDatasetId = HOME_TAB_ID;
			}
			render();
		});
	}

	tabContent.innerHTML = '';
	tabContent.appendChild(panel);
}

function renderEmptyWorkspace() {
	if (homeDragAbortController) {
		homeDragAbortController.abort();
		homeDragAbortController = null;
	}

	sidebarPanel.innerHTML = `
		<div class="sidebar-stack">
			<div>
				<p class="eyebrow">Geocoding Demo</p>
				<h1>檔案中心</h1>
				<p class="sidebar-copy">
					先匯入資料表，接著到檔案中心管理分頁，再進入表格分頁使用左側工具。
				</p>
			</div>
			<div class="panel">
				<div class="panel-title-row">
					<h2>匯入資料</h2>
				</div>
				<p class="sidebar-copy compact-copy">
					目前還沒有任何資料分頁。匯入 Excel 或 CSV 後，右側會顯示檔案中心頁面。
				</p>
				<button class="primary-button empty-import-button">匯入第一份資料表</button>
			</div>
		</div>
	`;

	sidebarPanel.querySelector('.empty-import-button').addEventListener('click', importFiles);
	sidebarPanel.scrollTop = state.scrollPositions.sidebarTop;

	tabContent.innerHTML = '';
	tabContent.appendChild(emptyStateTemplate.content.firstElementChild.cloneNode(true));
}

function renderSidebar(dataset, tableElement) {
	const sidebarRoot = sidebarTemplate.content.firstElementChild.cloneNode(true);
	const composeFieldTags = sidebarRoot.querySelector('.compose-field-tags');
	const composeFormatList = sidebarRoot.querySelector('.compose-format-list');
	const addComposeFormatButton = sidebarRoot.querySelector('.add-compose-format');
	const promoteComposeFormatButton = sidebarRoot.querySelector('.promote-compose-format');
	const removeComposeFormatButton = sidebarRoot.querySelector('.remove-compose-format');
	const composeSegmentList = sidebarRoot.querySelector('.compose-segment-list');
	const composeLiteralInput = sidebarRoot.querySelector('.compose-literal-input');
	const addComposeLiteralButton = sidebarRoot.querySelector('.add-compose-literal');
	const composePreviewValueElement = sidebarRoot.querySelector('.compose-preview-value');
	const applyCleanupButton = sidebarRoot.querySelector('.apply-cleanup');
	const composeAddressButton = sidebarRoot.querySelector('.compose-address');
	const runGeocodeAllButton = sidebarRoot.querySelector('.run-geocode-all');
	const runGeocodeSelectedButton = sidebarRoot.querySelector('.run-geocode-selected');
	const pauseGeocodeButton = sidebarRoot.querySelector('.pause-geocode');
	const stopGeocodeButton = sidebarRoot.querySelector('.stop-geocode');
	const geocodeForceReprocessCheckbox = sidebarRoot.querySelector('.geocode-force-reprocess');
	const geocodeReprocessUnmatchedCheckbox = sidebarRoot.querySelector('.geocode-reprocess-unmatched');
	const geocodeQueueHint = sidebarRoot.querySelector('.geocode-queue-hint');
	const geocodeState = ensureGeocodeState(dataset);
	const toolTabButtons = sidebarRoot.querySelectorAll('.tool-tab-button');
	const toolPanels = sidebarRoot.querySelectorAll('.tool-panel');
	const selectionSummary = sidebarRoot.querySelector('.selection-summary');
	const saveCleanupScriptButton = sidebarRoot.querySelector('.save-cleanup-script');
	const toggleCleanupLibraryButton = sidebarRoot.querySelector('.toggle-cleanup-library');
	const regexRuleList = sidebarRoot.querySelector('.regex-rule-list');
	const addRegexRuleButton = sidebarRoot.querySelector('.add-regex-rule');
	let draggingToken = null;
	let composeDragMoved = false;
	let composeAutoScrollTimer = null;

	ensureCleanupState(dataset);
	ensureComposeState(dataset);
	ensureRowSelectionState(dataset);
	hydrateColumnSelectors(sidebarRoot, dataset);
	sidebarRoot.querySelector('.cleanup-trim').checked = dataset.cleanupOptions.trim;
	sidebarRoot.querySelector('.cleanup-collapse-space').checked = dataset.cleanupOptions.collapseSpace;
	sidebarRoot.querySelector('.cleanup-fullwidth-space').checked = dataset.cleanupOptions.fullwidthSpace;
	sidebarRoot.querySelector('.cleanup-fullwidth-char').checked = dataset.cleanupOptions.fullwidthChar;
	sidebarRoot.querySelector('.cleanup-remove-linebreak').checked = dataset.cleanupOptions.removeLinebreak;
	selectionSummary.textContent = dataset.cleanupSelectedColumns.length === 0
		? '尚未選取欄位'
		: `已選取 ${dataset.cleanupSelectedColumns.length} 個欄位：${dataset.cleanupSelectedColumns.join('、')}`;
	updateComposePreview(composePreviewValueElement, dataset);
	const pendingCount = geocodeState.pendingRowIds.length;
	const selectedCount = dataset.selectedRowIds.length;
	geocodeForceReprocessCheckbox.checked = geocodeState.forceReprocess;
	geocodeReprocessUnmatchedCheckbox.checked = geocodeState.reprocessUnmatched;
	const pendingAllCount = getGeocodeTargetRows(dataset, 'all', geocodeForceReprocessCheckbox.checked, geocodeReprocessUnmatchedCheckbox.checked).length;
	const pendingSelectedCount = getGeocodeTargetRows(dataset, 'selected', geocodeForceReprocessCheckbox.checked, geocodeReprocessUnmatchedCheckbox.checked).length;
	const pendingUnmatchedCount = dataset.rows.filter((row) => {
		const status = row.geocode_status || '';
		return status && status !== 'success' && status !== 'error';
	}).length;
	const progressPercent = geocodeState.totalCount === 0
		? 0
		: Math.round((geocodeState.completedCount / geocodeState.totalCount) * 100);
	sidebarRoot.querySelector('.progress-fill').style.width = `${progressPercent}%`;
	sidebarRoot.querySelector('.progress-text').textContent = geocodeState.isPaused
		? `已暫停 ${geocodeState.completedCount}/${geocodeState.totalCount}`
		: geocodeState.isRunning
			? `查詢中 ${geocodeState.completedCount}/${geocodeState.totalCount}`
			: geocodeState.totalCount > 0
				? `已完成 ${geocodeState.completedCount}/${geocodeState.totalCount}`
				: '尚未開始';
	geocodeQueueHint.textContent = geocodeState.totalCount === 0 && !geocodeState.isRunning && !geocodeState.isPaused
		? `全部可處理 ${pendingAllCount} 筆，已勾選 ${selectedCount} 列，其中可處理 ${pendingSelectedCount} 筆，未匹配成功 ${pendingUnmatchedCount} 筆。`
		: `目前模式：${geocodeState.scope === 'selected' ? '勾選列' : '全部'}，待查詢 ${pendingCount} 筆，進行中欄位會在右側以不同底色標記。`;
	runGeocodeAllButton.textContent = geocodeState.isPaused && geocodeState.scope === 'all' ? '繼續處理全部' : '批次處理全部';
	runGeocodeSelectedButton.textContent = geocodeState.isPaused && geocodeState.scope === 'selected' ? '繼續處理勾選列' : '批次處理勾選列';
	runGeocodeAllButton.disabled = geocodeState.isRunning;
	runGeocodeSelectedButton.disabled = geocodeState.isRunning || (!geocodeState.isPaused && selectedCount === 0);
	pauseGeocodeButton.disabled = !geocodeState.isRunning;
	stopGeocodeButton.disabled = !geocodeState.isRunning && !geocodeState.isPaused;

	regexRuleList.innerHTML = dataset.cleanupRegexRules.length === 0
		? '<div class="regex-empty">尚未設定替換規則</div>'
		: dataset.cleanupRegexRules.map((rule, index) => `
			<div class="regex-rule-item" data-rule-index="${index}">
				<label class="field">
					<span>正則</span>
					<input type="text" class="regex-pattern" value="${escapeHtml(rule.pattern || '')}" placeholder="例如 \\s+">
				</label>
				<label class="field">
					<span>旗標</span>
					<input type="text" class="regex-flags" value="${escapeHtml(rule.flags || '')}" placeholder="例如 g">
				</label>
				<label class="field">
					<span>替換為</span>
					<input type="text" class="regex-replacement" value="${escapeHtml(rule.replacement || '')}" placeholder="替換文字">
				</label>
				<button class="mini-button danger-button remove-regex-rule" type="button">刪除規則</button>
			</div>
		`).join('');

	for (const button of toolTabButtons) {
		const isActive = button.dataset.tool === state.activeSidebarTool;
		button.classList.toggle('active', isActive);
		button.addEventListener('click', () => {
			state.activeSidebarTool = button.dataset.tool;
			render();
		});
	}

	for (const panel of toolPanels) {
		panel.classList.toggle('hidden', panel.dataset.toolPanel !== state.activeSidebarTool);
	}

	toggleCleanupLibraryButton.addEventListener('click', async () => {
		const result = await window.desktopApi.openCleanupScriptLibrary();
		if (!result?.ok) {
			window.alert(`無法開啟腳本庫：${result?.error || '未知錯誤'}`);
		}
	});

	saveCleanupScriptButton.addEventListener('click', async () => {
		const persistAndRefresh = async () => {
			const result = await window.desktopApi.saveWorkspace(serializeWorkspaceState());
			if (!result?.ok) {
				window.alert(`儲存腳本失敗：${result?.error || '未知錯誤'}`);
				return false;
			}

			workspacePersistence.lastSavedSnapshot = JSON.stringify(serializeWorkspaceState());
			await window.desktopApi.notifyCleanupScriptsUpdated();
			return true;
		};

		const currentConfig = getCleanupFlowConfig(dataset, sidebarRoot);
		const proposedName = `清洗腳本 ${state.cleanupScripts.length + 1}`;
		const scriptName = await promptForText({
			title: '儲存清洗腳本',
			description: '這個腳本只會保存清洗規則與正則設定，不會保存欄位勾選。',
			label: '腳本名稱',
			defaultValue: proposedName,
			placeholder: '例如：地址標準化',
			confirmText: '儲存腳本',
		});
		if (scriptName === null) {
			return;
		}

		const existing = state.cleanupScripts.find((script) => script.name === scriptName.trim());
		if (existing) {
			const confirmed = window.confirm(`「${existing.name}」已存在，要用目前流程覆蓋嗎？`);
			if (!confirmed) {
				return;
			}
		}

		const result = saveCleanupScript(scriptName, currentConfig);
		if (!result.ok) {
			window.alert('腳本名稱不能是空白。');
			return;
		}

		const saved = await persistAndRefresh();
		if (saved) {
			render();
		}
	});

	function refreshComposeTokenIndices() {
		for (const [index, token] of Array.from(composeSegmentList.querySelectorAll('.compose-token-button')).entries()) {
			token.dataset.segmentIndex = String(index);
		}
	}

	function swapComposeTokenNodes(firstToken, secondToken) {
		if (!firstToken || !secondToken || firstToken === secondToken) {
			return;
		}

		const placeholder = document.createElement('span');
		firstToken.replaceWith(placeholder);
		secondToken.replaceWith(firstToken);
		placeholder.replaceWith(secondToken);
		refreshComposeTokenIndices();
	}

	function stopComposeAutoScroll() {
		if (!composeAutoScrollTimer) {
			return;
		}

		window.clearInterval(composeAutoScrollTimer);
		composeAutoScrollTimer = null;
	}

	function startComposeAutoScroll(direction) {
		stopComposeAutoScroll();
		composeAutoScrollTimer = window.setInterval(() => {
			composeSegmentList.scrollLeft += direction * 14;
			state.scrollPositions.composeScrollLeftByDataset[state.activeDatasetId] = composeSegmentList.scrollLeft;
		}, 16);
	}

	composeSegmentList.addEventListener('dragover', (event) => {
		event.preventDefault();
		const rect = composeSegmentList.getBoundingClientRect();
		const threshold = 48;
		if (event.clientX < rect.left + threshold) {
			stopComposeAutoScroll();
			startComposeAutoScroll(-1);
		} else if (event.clientX > rect.right - threshold) {
			stopComposeAutoScroll();
			startComposeAutoScroll(1);
		} else {
			stopComposeAutoScroll();
		}
	});

	composeSegmentList.addEventListener('dragleave', () => {
		stopComposeAutoScroll();
	});

	composeSegmentList.addEventListener('scroll', () => {
		state.scrollPositions.composeScrollLeftByDataset[state.activeDatasetId] = composeSegmentList.scrollLeft;
	});

	function bindComposeTokenEvents() {
		for (const token of composeSegmentList.querySelectorAll('.compose-token-button')) {
			token.addEventListener('click', () => {
				if (composeDragMoved) {
					return;
				}

				const index = Number(token.dataset.segmentIndex);
				recordDatasetUndoPoint(dataset);
				const currentFormat = getActiveComposeFormat(dataset);
				currentFormat.segments.splice(index, 1);
				finalizeDatasetHistory(dataset);
				renderComposeEditor();
			});

			token.addEventListener('dragstart', (event) => {
				draggingToken = token;
				composeDragMoved = false;
				event.dataTransfer?.setData('text/plain', token.dataset.segmentIndex);
				event.dataTransfer.effectAllowed = 'move';
				token.classList.add('is-dragging');
			});

			token.addEventListener('dragover', (event) => {
				event.preventDefault();
				const rect = composeSegmentList.getBoundingClientRect();
				const threshold = 48;
				if (event.clientX < rect.left + threshold) {
					stopComposeAutoScroll();
					startComposeAutoScroll(-1);
				} else if (event.clientX > rect.right - threshold) {
					stopComposeAutoScroll();
					startComposeAutoScroll(1);
				} else {
					stopComposeAutoScroll();
				}
			});

			token.addEventListener('dragenter', (event) => {
				event.preventDefault();
				if (!draggingToken || draggingToken === token) {
					return;
				}

				const fromIndex = Number(draggingToken.dataset.segmentIndex);
				const toIndex = Number(token.dataset.segmentIndex);
				if (Number.isNaN(fromIndex) || Number.isNaN(toIndex) || fromIndex === toIndex) {
					return;
				}

				swapComposeSegment(dataset, fromIndex, toIndex);
				finalizeDatasetHistory(dataset);
				swapComposeTokenNodes(draggingToken, token);
				updateComposePreview(composePreviewValueElement, dataset);
				composeDragMoved = true;
				if (event.dataTransfer) {
					event.dataTransfer.setData('text/plain', String(toIndex));
				}
			});

			token.addEventListener('drop', (event) => {
				event.preventDefault();
				stopComposeAutoScroll();
			});

			token.addEventListener('dragend', () => {
				stopComposeAutoScroll();
				token.classList.remove('is-dragging');
				draggingToken = null;
				window.setTimeout(() => {
					composeDragMoved = false;
				}, 0);
			});
		}
	}

	function renderComposeEditor() {
		const composeScrollLeft = composeSegmentList.scrollLeft;
		const activeComposeFormat = getActiveComposeFormat(dataset);

		composeFieldTags.innerHTML = (dataset.sourceColumnNames || dataset.columnNames)
			.filter((column) => column !== '__rowId' && !SYSTEM_COLUMNS.includes(column))
			.map((column) => `
				<button class="field-tag-button" type="button" data-column-name="${escapeHtml(column)}">${escapeHtml(column)}</button>
			`)
			.join('');

		composeFormatList.innerHTML = dataset.composeFormats.map((format, index) => `
			<button
				class="compose-format-button ${format.id === activeComposeFormat.id ? 'active' : ''}"
				type="button"
				data-format-id="${escapeHtml(format.id)}"
			>${escapeHtml(index === 0 ? '主要格式' : `備選格式 ${index}`)}</button>
		`).join('');

		promoteComposeFormatButton.disabled = dataset.composeFormats.length <= 1 || dataset.composeFormats[0].id === activeComposeFormat.id;
		removeComposeFormatButton.disabled = dataset.composeFormats.length <= 1;
		composeSegmentList.innerHTML = activeComposeFormat.segments.length === 0
			? '<div class="regex-empty">尚未建立候選格式</div>'
			: activeComposeFormat.segments.map((segment, index) => `
				<button
					class="compose-token-button ${segment.type === 'text' ? 'compose-text-token' : ''}"
					type="button"
					draggable="true"
					data-segment-index="${index}"
					title="拖曳排序，點擊刪除"
				>${escapeHtml(segment.value)}</button>
			`).join('');

		updateComposePreview(composePreviewValueElement, dataset);

		for (const button of composeFieldTags.querySelectorAll('.field-tag-button')) {
			button.addEventListener('click', () => {
				recordDatasetUndoPoint(dataset);
				const currentFormat = getActiveComposeFormat(dataset);
				currentFormat.segments.push({
					type: 'field',
					value: button.dataset.columnName,
				});
				finalizeDatasetHistory(dataset);
				renderComposeEditor();
			});
		}

		for (const button of composeFormatList.querySelectorAll('.compose-format-button')) {
			button.addEventListener('click', () => {
				dataset.activeComposeFormatId = button.dataset.formatId;
				renderComposeEditor();
			});
		}

		bindComposeTokenEvents();

		const restoredScrollLeft = typeof state.scrollPositions.composeScrollLeftByDataset[state.activeDatasetId] === 'number'
			? state.scrollPositions.composeScrollLeftByDataset[state.activeDatasetId]
			: composeScrollLeft;
		composeSegmentList.scrollLeft = restoredScrollLeft;
	}

	renderComposeEditor();

	addComposeLiteralButton.addEventListener('click', () => {
		if (!composeLiteralInput.value.trim()) {
			return;
		}

		recordDatasetUndoPoint(dataset);
		const currentFormat = getActiveComposeFormat(dataset);
		currentFormat.segments.push({
			type: 'text',
			value: composeLiteralInput.value,
		});
		finalizeDatasetHistory(dataset);
		renderComposeEditor();
	});

	addComposeFormatButton.addEventListener('click', () => {
		recordDatasetUndoPoint(dataset);
		const nextIndex = dataset.composeFormats.length;
		const nextFormat = createComposeFormat(`備選格式 ${nextIndex}`);
		dataset.composeFormats.push(nextFormat);
		dataset.activeComposeFormatId = nextFormat.id;
		finalizeDatasetHistory(dataset);
		renderComposeEditor();
	});

	promoteComposeFormatButton.addEventListener('click', () => {
		const activeIndex = dataset.composeFormats.findIndex((format) => format.id === dataset.activeComposeFormatId);
		if (activeIndex <= 0) {
			return;
		}

		recordDatasetUndoPoint(dataset);
		moveItem(dataset.composeFormats, activeIndex, 0);
		finalizeDatasetHistory(dataset);
		renderComposeEditor();
	});

	removeComposeFormatButton.addEventListener('click', () => {
		if (dataset.composeFormats.length <= 1) {
			return;
		}

		const activeIndex = dataset.composeFormats.findIndex((format) => format.id === dataset.activeComposeFormatId);
		if (activeIndex < 0) {
			return;
		}

		recordDatasetUndoPoint(dataset);
		dataset.composeFormats.splice(activeIndex, 1);
		dataset.activeComposeFormatId = dataset.composeFormats[Math.max(0, activeIndex - 1)].id;
		finalizeDatasetHistory(dataset);
		renderComposeEditor();
	});

	composeLiteralInput.addEventListener('keydown', (event) => {
		if (event.key !== 'Enter') {
			return;
		}

		event.preventDefault();
		addComposeLiteralButton.click();
	});

	addRegexRuleButton.addEventListener('click', () => {
		recordDatasetUndoPoint(dataset);
		dataset.cleanupRegexRules.push({
			pattern: '',
			flags: 'g',
			replacement: '',
		});
		finalizeDatasetHistory(dataset);
		render();
	});

	for (const ruleItem of regexRuleList.querySelectorAll('.regex-rule-item')) {
		const index = Number(ruleItem.dataset.ruleIndex);
		const rule = dataset.cleanupRegexRules[index];
		if (!rule) {
			continue;
		}

		ruleItem.querySelector('.regex-pattern').addEventListener('input', (event) => {
			rule.pattern = event.target.value;
		});

		ruleItem.querySelector('.regex-flags').addEventListener('input', (event) => {
			rule.flags = event.target.value;
		});

		ruleItem.querySelector('.regex-replacement').addEventListener('input', (event) => {
			rule.replacement = event.target.value;
		});

		ruleItem.querySelector('.remove-regex-rule').addEventListener('click', () => {
			recordDatasetUndoPoint(dataset);
			dataset.cleanupRegexRules.splice(index, 1);
			finalizeDatasetHistory(dataset);
			render();
		});
	}

	applyCleanupButton.addEventListener('click', () => {
		if (dataset.cleanupSelectedColumns.length === 0) {
			return;
		}

		recordDatasetUndoPoint(dataset);
		dataset.cleanupOptions = {
			trim: sidebarRoot.querySelector('.cleanup-trim').checked,
			collapseSpace: sidebarRoot.querySelector('.cleanup-collapse-space').checked,
			fullwidthSpace: sidebarRoot.querySelector('.cleanup-fullwidth-space').checked,
			fullwidthChar: sidebarRoot.querySelector('.cleanup-fullwidth-char').checked,
			removeLinebreak: sidebarRoot.querySelector('.cleanup-remove-linebreak').checked,
		};

		applyCleanup(dataset, dataset.cleanupSelectedColumns, {
			trim: dataset.cleanupOptions.trim,
			collapseSpace: dataset.cleanupOptions.collapseSpace,
			fullwidthSpace: dataset.cleanupOptions.fullwidthSpace,
			fullwidthChar: dataset.cleanupOptions.fullwidthChar,
			removeLinebreak: dataset.cleanupOptions.removeLinebreak,
		}, dataset.cleanupRegexRules);

		finalizeDatasetHistory(dataset);
		render();
	});

	composeAddressButton.addEventListener('click', () => {
		if (dataset.composeFormats[0].segments.length === 0) {
			return;
		}

		recordDatasetUndoPoint(dataset);
		composeCandidateAddress(dataset);
		finalizeDatasetHistory(dataset);
		render();
	});

	geocodeForceReprocessCheckbox.addEventListener('change', () => {
		recordDatasetUndoPoint(dataset);
		geocodeState.forceReprocess = geocodeForceReprocessCheckbox.checked;
		finalizeDatasetHistory(dataset);
		render();
	});

	geocodeReprocessUnmatchedCheckbox.addEventListener('change', () => {
		recordDatasetUndoPoint(dataset);
		geocodeState.reprocessUnmatched = geocodeReprocessUnmatchedCheckbox.checked;
		finalizeDatasetHistory(dataset);
		render();
	});

	runGeocodeAllButton.addEventListener('click', async () => {
		await runGeocode(dataset.id, {
			scope: 'all',
			forceReprocess: geocodeForceReprocessCheckbox.checked,
			reprocessUnmatched: geocodeReprocessUnmatchedCheckbox.checked,
		});
	});

	runGeocodeSelectedButton.addEventListener('click', async () => {
		await runGeocode(dataset.id, {
			scope: 'selected',
			forceReprocess: geocodeForceReprocessCheckbox.checked,
			reprocessUnmatched: geocodeReprocessUnmatchedCheckbox.checked,
		});
	});

	pauseGeocodeButton.addEventListener('click', () => {
		pauseGeocode(dataset.id);
	});

	stopGeocodeButton.addEventListener('click', () => {
		stopGeocode(dataset.id);
	});

	sidebarPanel.innerHTML = '';
	sidebarPanel.appendChild(sidebarRoot);
}

function renderDatasetPanel(dataset) {
	if (homeDragAbortController) {
		homeDragAbortController.abort();
		homeDragAbortController = null;
	}

	const panel = tabPanelTemplate.content.firstElementChild.cloneNode(true);
	const tableElement = panel.querySelector('.dataset-table');
	const summary = panel.querySelector('.dataset-summary');
	const visibilityBar = panel.querySelector('.column-visibility-bar');

	ensureColumnVisibilityState(dataset);
	summary.textContent = `${dataset.sheetName} • ${dataset.rows.length} 筆資料`;
	visibilityBar.innerHTML = dataset.hiddenColumns.length === 0
		? ''
		: `
			<span class="column-visibility-copy">已隱藏欄位</span>
			${dataset.hiddenColumns.map((column) => `
				<button class="column-visibility-chip" type="button" data-column-name="${escapeHtml(column)}">
					👁 ${escapeHtml(column)}
				</button>
			`).join('')}
			<button class="column-visibility-chip" type="button" data-action="show-all-columns">全部顯示</button>
		`;
	renderTable(tableElement, dataset);
	renderSidebar(dataset, tableElement);

	for (const button of visibilityBar.querySelectorAll('.column-visibility-chip')) {
		button.addEventListener('click', () => {
			if (button.dataset.action === 'show-all-columns') {
				dataset.hiddenColumns = [];
				render();
				return;
			}

			toggleColumnHidden(dataset, button.dataset.columnName);
			render();
		});
	}

	tabContent.innerHTML = '';
	tabContent.appendChild(panel);
}

function renderActivePanel() {
	if (state.datasets.length === 0) {
		renderEmptyWorkspace();
		return;
	}

	if (state.activeDatasetId === HOME_TAB_ID) {
		renderHomeSidebar();
		renderHomePanel();
		restoreHomeScrollPositions();
		return;
	}

	const dataset = getDatasetById(state.activeDatasetId);
	if (!dataset) {
		state.activeDatasetId = HOME_TAB_ID;
		renderHomeSidebar();
		renderHomePanel();
		restoreHomeScrollPositions();
		return;
	}

	renderDatasetPanel(dataset);
}

function setActiveDataset(datasetId) {
	state.activeDatasetId = datasetId;
	if (datasetId !== HOME_TAB_ID) {
		state.activeSidebarTool = state.activeSidebarTool || 'cleanup';
	}
	render();
}

function render() {
	captureScrollPositions();
	blurWorkspaceFocus();

	if (state.datasets.length === 0) {
		datasetCount.textContent = '尚未匯入資料表';
	} else if (state.activeDatasetId === HOME_TAB_ID) {
		datasetCount.textContent = `共 ${state.datasets.length} 個資料分頁`;
	} else {
		const activeDataset = getDatasetById(state.activeDatasetId);
		datasetCount.textContent = activeDataset
			? `目前分頁：${formatDatasetLabel(activeDataset)}`
			: `共 ${state.datasets.length} 個資料分頁`;
	}

	const activeDataset = getDatasetById(state.activeDatasetId);
	const activeHistory = activeDataset ? ensureDatasetHistory(activeDataset) : null;
	const isDatasetActive = state.activeDatasetId !== HOME_TAB_ID && Boolean(activeDataset);
	const isDirty = activeDataset && activeHistory
		? getDatasetHistorySnapshot(activeDataset) !== activeHistory.savedSnapshot
		: false;
	saveDatasetButton.disabled = !isDatasetActive || !isDirty;
	saveDatasetButton.textContent = isDirty ? '存檔 *' : '存檔';
	undoDatasetButton.disabled = !isDatasetActive || !activeHistory || activeHistory.undoStack.length === 0;
	redoDatasetButton.disabled = !isDatasetActive || !activeHistory || activeHistory.redoStack.length === 0;
	exportDatasetButton.disabled = state.activeDatasetId === HOME_TAB_ID || !activeDataset;

	renderTabs();
	renderActivePanel();
	scheduleScrollRestore();
	scheduleWorkspaceSave();
}

async function importFiles() {
	const imported = await window.desktopApi.openExcelFiles();
	if (!imported || imported.length === 0) {
		return;
	}

	for (const dataset of imported) {
		ensureSystemColumns(dataset);
		initializeDatasetHistory(dataset);
	}

	state.datasets.push(...imported);
	render();
}

async function importFilesFromDrop(filePaths) {
	const normalizedPaths = Array.isArray(filePaths)
		? filePaths.filter((filePath) => typeof filePath === 'string' && filePath.trim() !== '')
		: [];
	if (normalizedPaths.length === 0) {
		return;
	}

	const imported = await window.desktopApi.importFilesByPath(normalizedPaths);
	if (!imported || imported.length === 0) {
		window.alert('沒有可匯入的 Excel 或 CSV 檔案。');
		return;
	}

	for (const dataset of imported) {
		ensureSystemColumns(dataset);
		initializeDatasetHistory(dataset);
	}

	state.datasets.push(...imported);
	render();
}

exportDatasetButton.addEventListener('click', async () => {
	const dataset = getDatasetById(state.activeDatasetId);
	if (!dataset) {
		return;
	}

	const result = await window.desktopApi.exportDataset(createExportPayload(dataset));
	if (result?.error) {
		window.alert(`匯出失敗：${result.error}`);
	}
});

saveDatasetButton.addEventListener('click', async () => {
	const dataset = getDatasetById(state.activeDatasetId);
	if (!dataset) {
		return;
	}

	const confirmed = await confirmAction({
		title: '確認存檔',
		description: `確定要儲存目前分頁「${formatDatasetLabel(dataset)}」的變更嗎？`,
		confirmText: '確認存檔',
	});
	if (!confirmed) {
		return;
	}

	await saveCurrentDataset(state.activeDatasetId);
});

undoDatasetButton.addEventListener('click', () => {
	undoDataset(state.activeDatasetId);
});

redoDatasetButton.addEventListener('click', () => {
	redoDataset(state.activeDatasetId);
});

window.desktopApi.onCleanupScriptApplyRequest((scriptId) => {
	const dataset = getDatasetById(state.activeDatasetId);
	if (!dataset) {
		window.alert('請先切換到一個資料分頁，再調用清洗腳本。');
		return;
	}

	const script = getCleanupScriptById(scriptId);
	if (!script) {
		window.alert('找不到指定的清洗腳本。');
		return;
	}

	recordDatasetUndoPoint(dataset);
	applyCleanupScriptToDataset(dataset, script);
	finalizeDatasetHistory(dataset);
	render();
});

window.desktopApi.onCleanupScriptsUpdated(async () => {
	const workspace = await window.desktopApi.loadWorkspace();
	if (!workspace || typeof workspace !== 'object') {
		state.cleanupScripts = [];
		render();
		return;
	}

	state.cleanupScripts = normalizeCleanupScripts(workspace.cleanupScripts);
	render();
});

async function hydrateWorkspace() {
	const workspace = normalizeLoadedWorkspace(await window.desktopApi.loadWorkspace());
	if (workspace) {
		state.datasets = workspace.datasets;
		state.fileViewMode = workspace.fileViewMode;
		state.activeSidebarTool = workspace.activeSidebarTool;
		state.cleanupScripts = workspace.cleanupScripts;
		state.activeDatasetId = workspace.activeDatasetId === HOME_TAB_ID
			? HOME_TAB_ID
			: workspace.datasets.some((dataset) => dataset.id === workspace.activeDatasetId)
				? workspace.activeDatasetId
				: workspace.datasets.length > 0
					? workspace.datasets[0].id
					: HOME_TAB_ID;
	} else {
		state.cleanupScripts = [];
		state.activeDatasetId = HOME_TAB_ID;
	}

	workspacePersistence.lastSavedSnapshot = JSON.stringify(serializeWorkspaceState());
	workspacePersistence.isHydrating = false;
	render();
}

hydrateWorkspace();
