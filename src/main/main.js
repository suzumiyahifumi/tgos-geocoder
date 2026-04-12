const path = require('path');
const fs = require('fs/promises');
const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const XLSX = require('xlsx');
const { queryTgosAddress } = require('../services/tgos');

function getWorkspaceFilePath() {
	return path.join(app.getPath('userData'), 'workspace.json');
}

function sanitizeFileNamePart(value, fallback = 'untitled') {
	const normalized = normalizeCellValue(value).trim().replace(/[<>:"/\\|?*\x00-\x1F]+/g, '-');
	return normalized || fallback;
}

function createWindow() {
	const window = new BrowserWindow({
		width: 1480,
		height: 960,
		minWidth: 1180,
		minHeight: 760,
		backgroundColor: '#efe4d4',
		webPreferences: {
			preload: path.join(__dirname, 'preload.js'),
			contextIsolation: true,
			nodeIntegration: false,
		},
	});

	window.loadFile(path.join(__dirname, '../renderer/index.html'));
}

function normalizeCellValue(value) {
	if (value === undefined || value === null) {
		return '';
	}

	if (typeof value === 'string') {
		return value;
	}

	return String(value);
}

function sheetToRecords(sheet) {
	const rows = XLSX.utils.sheet_to_json(sheet, {
		header: 1,
		defval: '',
		raw: false,
	});

	const nonEmptyRows = rows.filter((row) => row.some((cell) => `${cell}`.trim() !== ''));

	if (nonEmptyRows.length === 0) {
		return { columns: [], rows: [] };
	}

	const headers = nonEmptyRows[0].map((value, index) => {
		const text = normalizeCellValue(value).trim();
		return text || `欄位${index + 1}`;
	});

	const records = nonEmptyRows.slice(1).map((row, rowIndex) => {
		const record = { __rowId: `row-${rowIndex + 1}` };

		headers.forEach((header, index) => {
			record[header] = normalizeCellValue(row[index]);
		});

		return record;
	});

	return {
		columns: headers,
		rows: records,
	};
}

function workbookFileToDatasets(filePath) {
	const workbook = XLSX.readFile(filePath, {
		cellDates: false,
		raw: false,
	});

	if (workbook.SheetNames.length === 0) {
		throw new Error(`檔案沒有可讀取的工作表: ${path.basename(filePath)}`);
	}

	return workbook.SheetNames.map((sheetName, index) => {
		const { columns, rows } = sheetToRecords(workbook.Sheets[sheetName]);

		return {
			id: `${path.basename(filePath)}-${index}-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
			fileName: path.basename(filePath),
			filePath,
			sheetName,
			columnNames: columns,
			rows,
			importedAt: new Date().toISOString(),
		};
	});
}

function safeString(value) {
	if (value === undefined || value === null) {
		return '';
	}

	return typeof value === 'string' ? value : JSON.stringify(value);
}

function findFirstValueByKeys(payload, keys) {
	const normalizedKeys = new Set(keys.map((key) => key.toLowerCase()));
	const stack = [payload];
	const seen = new Set();

	while (stack.length > 0) {
		const current = stack.pop();

		if (!current || typeof current !== 'object' || seen.has(current)) {
			continue;
		}

		seen.add(current);

		if (Array.isArray(current)) {
			for (const item of current) {
				stack.push(item);
			}
			continue;
		}

		for (const [key, value] of Object.entries(current)) {
			if (normalizedKeys.has(key.toLowerCase()) && value !== '' && value !== null && value !== undefined) {
				return value;
			}

			if (value && typeof value === 'object') {
				stack.push(value);
			}
		}
	}

	return '';
}

function inferCoordSystem(payload) {
	const coordSys = findFirstValueByKeys(payload, [
		'coordSys',
		'coordsys',
		'coordinateSystem',
		'prj',
		'epsg',
	]);

	return coordSys ? safeString(coordSys) : '未知';
}

function normalizeAddressForCompare(value) {
	return safeString(value)
		.replace(/\s+/g, '')
		.replaceAll('臺', '台')
		.trim();
}

function findArrayByKeys(payload, keys) {
	const normalizedKeys = new Set(keys.map((key) => key.toLowerCase()));
	const stack = [payload];
	const seen = new Set();

	while (stack.length > 0) {
		const current = stack.pop();

		if (!current || typeof current !== 'object' || seen.has(current)) {
			continue;
		}

		seen.add(current);

		if (Array.isArray(current)) {
			for (const item of current) {
				stack.push(item);
			}
			continue;
		}

		for (const [key, value] of Object.entries(current)) {
			if (normalizedKeys.has(key.toLowerCase()) && Array.isArray(value)) {
				return value;
			}

			if (value && typeof value === 'object') {
				stack.push(value);
			}
		}
	}

	return null;
}

function inferResultCount(payload) {
	const explicitCount = Number(findFirstValueByKeys(payload, [
		'count',
		'total',
		'totalCount',
		'resultCount',
		'candidateCount',
		'cnt',
	]));

	if (Number.isFinite(explicitCount) && explicitCount >= 0) {
		return explicitCount;
	}

	const candidateList = findArrayByKeys(payload, [
		'AddressList',
		'addressList',
		'results',
		'Result',
		'result',
		'items',
		'data',
		'candidates',
		'candidateList',
	]);

	if (candidateList) {
		return candidateList.length;
	}

	const matchedAddress = findFirstValueByKeys(payload, [
		'FULL_ADDR',
		'full_addr',
		'FULLADDRESS',
		'ADDRESS',
		'address',
		'formattedAddress',
		'MatchAddr',
		'matchAddr',
		'LOC_ADDR',
	]);

	return matchedAddress ? 1 : 0;
}

function inferMatchType(candidateAddress, payload, matchedAddress, resultCount) {
	const explicitType = safeString(findFirstValueByKeys(payload, [
		'matchType',
		'MatchType',
		'queryType',
		'searchType',
		'precision',
		'matchLevel',
	]));

	const normalizedExplicitType = explicitType.toLowerCase();
	if (normalizedExplicitType.includes('exact') || explicitType.includes('完全') || explicitType.includes('精確')) {
		return '完全匹配';
	}

	if (normalizedExplicitType.includes('fuzzy') || explicitType.includes('模糊')) {
		return '模糊匹配';
	}

	if (resultCount === 0) {
		return '無結果';
	}

	if (resultCount > 1) {
		return '多筆候選';
	}

	if (normalizeAddressForCompare(candidateAddress) === normalizeAddressForCompare(matchedAddress)) {
		return '完全匹配';
	}

	return matchedAddress ? '模糊匹配' : '無結果';
}

function normalizeTgosResult(candidateAddress, payload) {
	const matchedAddress = findFirstValueByKeys(payload, [
		'FULL_ADDR',
		'full_addr',
		'FULLADDRESS',
		'ADDRESS',
		'address',
		'formattedAddress',
		'MatchAddr',
		'matchAddr',
		'LOC_ADDR',
	]) || '';

	const secondCandidate = findFirstValueByKeys(payload, [
		'CAND_ADDR',
		'candAddr',
		'candidateAddress',
		'cand_address',
	]) || '';

	const x = findFirstValueByKeys(payload, ['X', 'x', 'TWD97_X', 'twd97x', 'lng', 'lon', 'longitude']) || '';
	const y = findFirstValueByKeys(payload, ['Y', 'y', 'TWD97_Y', 'twd97y', 'lat', 'latitude']) || '';
	const resultCount = inferResultCount(payload);
	const matchType = inferMatchType(candidateAddress, payload, matchedAddress, resultCount);
	const status = resultCount === 0 && !matchedAddress && !secondCandidate && !x && !y ? 'no_result' : 'success';

	return {
		raw: payload,
		status,
		inputCandidateAddress: candidateAddress,
		tgosCandidateAddress: safeString(secondCandidate),
		matchedAddress: safeString(matchedAddress),
		x: safeString(x),
		y: safeString(y),
		coordSystem: inferCoordSystem(payload),
		matchType,
		resultCount,
	};
}

async function loadWorkspaceFromDisk() {
	try {
		const filePath = getWorkspaceFilePath();
		const content = await fs.readFile(filePath, 'utf8');
		return JSON.parse(content);
	} catch (error) {
		if (error.code === 'ENOENT') {
			return null;
		}

		console.log('[workspace:load:error]', error.message);
		return null;
	}
}

async function saveWorkspaceToDisk(workspace) {
	const filePath = getWorkspaceFilePath();
	await fs.mkdir(path.dirname(filePath), { recursive: true });
	await fs.writeFile(filePath, JSON.stringify(workspace, null, 2), 'utf8');
}

async function exportDatasetToFile(dataset) {
	const headers = Array.isArray(dataset?.columnNames)
		? dataset.columnNames.filter((column) => column !== '__rowId')
		: [];
	const rows = Array.isArray(dataset?.rows) ? dataset.rows : [];
	const exportRows = rows.map((row) => {
		const record = {};
		for (const header of headers) {
			record[header] = row?.[header] ?? '';
		}
		return record;
	});

	const defaultBaseName = `${sanitizeFileNamePart(dataset?.fileName, 'dataset')}-${sanitizeFileNamePart(dataset?.sheetName, 'sheet')}`;
	const saveResult = await dialog.showSaveDialog({
		defaultPath: `${defaultBaseName}.xlsx`,
		filters: [
			{ name: 'Excel Workbook', extensions: ['xlsx'] },
			{ name: 'CSV', extensions: ['csv'] },
		],
	});

	if (saveResult.canceled || !saveResult.filePath) {
		return { canceled: true };
	}

	const filePath = saveResult.filePath;
	const extension = path.extname(filePath).toLowerCase();
	const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: headers });

	if (extension === '.csv') {
		const csv = XLSX.utils.sheet_to_csv(worksheet);
		await fs.writeFile(filePath, csv, 'utf8');
		return { canceled: false, filePath, format: 'csv' };
	}

	const workbook = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(workbook, worksheet, sanitizeFileNamePart(dataset?.sheetName, 'Sheet1').slice(0, 31));
	XLSX.writeFile(workbook, filePath);
	return { canceled: false, filePath, format: 'xlsx' };
}

ipcMain.handle('dialog:openExcelFiles', async () => {
	const result = await dialog.showOpenDialog({
		properties: ['openFile', 'multiSelections'],
		filters: [
			{ name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm', 'csv'] },
		],
	});

	if (result.canceled) {
		return [];
	}

	return result.filePaths.flatMap(workbookFileToDatasets);
});

ipcMain.handle('workspace:load', async () => loadWorkspaceFromDisk());

ipcMain.handle('workspace:save', async (_event, workspace) => {
	try {
		await saveWorkspaceToDisk(workspace);
		return { ok: true };
	} catch (error) {
		console.log('[workspace:save:error]', error.message);
		return {
			ok: false,
			error: error.message,
		};
	}
});

ipcMain.handle('dataset:export', async (_event, dataset) => {
	try {
		return await exportDatasetToFile(dataset);
	} catch (error) {
		console.log('[dataset:export:error]', error.message);
		return {
			canceled: false,
			error: error.message,
		};
	}
});

ipcMain.handle('geocode:queryOne', async (_event, job) => {
	const candidateAddress = safeString(job.candidateAddress).trim();

	if (!candidateAddress) {
		return {
			rowId: job.rowId,
			status: 'skipped',
			inputCandidateAddress: '',
			tgosCandidateAddress: '',
			matchedAddress: '',
			x: '',
			y: '',
			coordSystem: '',
			error: '候選地址為空白',
		};
	}

	try {
		const payload = await queryTgosAddress(candidateAddress);
		const normalizedResult = normalizeTgosResult(candidateAddress, payload);
		console.log('[geocode:queryOne]', {
			rowId: job.rowId,
			candidateAddress,
			status: normalizedResult.status,
			matchType: normalizedResult.matchType,
			resultCount: normalizedResult.resultCount,
			matchedAddress: normalizedResult.matchedAddress,
			tgosCandidateAddress: normalizedResult.tgosCandidateAddress,
		});
		if (normalizedResult.status !== 'success' || normalizedResult.matchType !== '完全匹配' || normalizedResult.resultCount > 1) {
			console.log('[geocode:queryOne:raw]', safeString(payload));
		}
		return {
			rowId: job.rowId,
			...normalizedResult,
		};
	} catch (error) {
		console.log('[geocode:queryOne:error]', {
			rowId: job.rowId,
			candidateAddress,
			error: error.message,
		});
		return {
			rowId: job.rowId,
			status: 'error',
			inputCandidateAddress: candidateAddress,
			tgosCandidateAddress: '',
			matchedAddress: '',
			x: '',
			y: '',
			coordSystem: '',
			matchType: '',
			resultCount: 0,
			error: error.message,
		};
	}
});

app.whenReady().then(() => {
	createWindow();

	app.on('activate', () => {
		if (BrowserWindow.getAllWindows().length === 0) {
			createWindow();
		}
	});
});

app.on('window-all-closed', () => {
	if (process.platform !== 'darwin') {
		app.quit();
	}
});
