function escapeHtml(value) {
	return String(value)
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&#39;');
}

function safeString(value) {
	return value === undefined || value === null ? '' : String(value);
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
	const template = document.getElementById('libraryTextInputDialogTemplate');
	return new Promise((resolve) => {
		const dialogRoot = template.content.firstElementChild.cloneNode(true);
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
	const template = document.getElementById('libraryConfirmDialogTemplate');
	return new Promise((resolve) => {
		const dialogRoot = template.content.firstElementChild.cloneNode(true);
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

function summarizeCleanupOptions(options) {
	const labels = [];
	if (options.trim) labels.push('去除前後空白');
	if (options.collapseSpace) labels.push('壓縮連續空白');
	if (options.fullwidthSpace) labels.push('全形空白轉半形');
	if (options.fullwidthChar) labels.push('全形轉半形');
	if (options.removeLinebreak) labels.push('移除換行');
	return labels.length > 0 ? labels.join('、') : '未啟用基礎規則';
}

const scriptCount = document.getElementById('scriptCount');
const scriptList = document.getElementById('scriptList');

async function loadWorkspaceScripts() {
	const workspace = await window.desktopApi.loadWorkspace();
	return Array.isArray(workspace?.cleanupScripts) ? workspace.cleanupScripts : [];
}

async function saveWorkspaceScripts(scripts) {
	const workspace = await window.desktopApi.loadWorkspace();
	const nextWorkspace = {
		...(workspace && typeof workspace === 'object' ? workspace : {}),
		cleanupScripts: scripts,
	};
	const result = await window.desktopApi.saveWorkspace(nextWorkspace);
	if (!result?.ok) {
		window.alert(`儲存腳本庫失敗：${result?.error || '未知錯誤'}`);
		return false;
	}

	await window.desktopApi.notifyCleanupScriptsUpdated();

	return true;
}

async function render() {
	const scripts = await loadWorkspaceScripts();
	scriptCount.textContent = `${scripts.length} 筆腳本`;
	scriptList.innerHTML = scripts.length === 0
		? '<div class="cleanup-library-empty">尚未儲存任何清洗腳本。</div>'
		: scripts.map((script) => `
			<article class="cleanup-library-card" data-script-id="${escapeHtml(script.id || '')}">
				<div class="cleanup-library-card-header">
					<div>
						<h2>${escapeHtml(script.name || '未命名腳本')}</h2>
						<p>${escapeHtml(summarizeCleanupOptions(cloneCleanupOptions(script.cleanupOptions)))}</p>
						<p>正則規則 ${cloneCleanupRegexRules(script.cleanupRegexRules).length} 條</p>
					</div>
				</div>
				<div class="cleanup-library-card-actions">
					<button class="mini-button apply-script" type="button">調用到主視窗</button>
					<button class="mini-button rename-script" type="button">重新命名</button>
					<button class="mini-button danger-button delete-script" type="button">刪除</button>
				</div>
			</article>
		`).join('');

	for (const card of scriptList.querySelectorAll('.cleanup-library-card')) {
		const scriptId = card.dataset.scriptId;
		const script = scripts.find((item) => item.id === scriptId);
		if (!script) {
			continue;
		}

		card.querySelector('.apply-script').addEventListener('click', async () => {
			const result = await window.desktopApi.applyCleanupScriptFromLibrary(script.id);
			if (!result?.ok) {
				window.alert(`調用失敗：${result?.error || '未知錯誤'}`);
				return;
			}
		});

		card.querySelector('.rename-script').addEventListener('click', async () => {
			const nextName = await promptForText({
				title: '重新命名清洗腳本',
				description: '請輸入新的腳本名稱。',
				label: '腳本名稱',
				defaultValue: script.name || '',
				placeholder: '例如：地址標準化',
				confirmText: '儲存名稱',
			});
			if (nextName === null) {
				return;
			}

			const normalizedName = safeString(nextName).trim();
			if (!normalizedName) {
				window.alert('腳本名稱不能是空白。');
				return;
			}

			if (scripts.some((item) => item.id !== script.id && item.name === normalizedName)) {
				window.alert('已有同名腳本，請改用其他名稱。');
				return;
			}

			const nextScripts = scripts.map((item) => item.id === script.id
				? { ...item, name: normalizedName, updatedAt: new Date().toISOString() }
				: item);
			const saved = await saveWorkspaceScripts(nextScripts);
			if (saved) {
				await render();
			}
		});

		card.querySelector('.delete-script').addEventListener('click', async () => {
			const confirmed = await confirmAction({
				title: '刪除清洗腳本',
				description: `確定要刪除腳本「${script.name}」嗎？此操作無法復原。`,
				confirmText: '刪除',
			});
			if (!confirmed) {
				return;
			}

			const nextScripts = scripts.filter((item) => item.id !== script.id);
			const saved = await saveWorkspaceScripts(nextScripts);
			if (saved) {
				await render();
			}
		});
	}
}

render();

window.desktopApi.onCleanupScriptsUpdated(() => {
	render();
});
