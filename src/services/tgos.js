const axios = require('axios');
const cheerio = require('cheerio');
const { DatabaseSync } = require('node:sqlite');

const baseUrl = 'https://map.tgos.tw/TGOSCloudMap';
const postUrl = `${baseUrl}/reqcontroller.go`;
const db = new DatabaseSync('./tgos-cache.sqlite');

const commonHeaders = {
	'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36',
	'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
	'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
	'Connection': 'keep-alive',
};

db.exec(`
	CREATE TABLE IF NOT EXISTS tgos_headers (
		cache_key TEXT PRIMARY KEY,
		cookie TEXT NOT NULL,
		request_verification_token TEXT,
		gs_request_token TEXT,
		updated_at TEXT NOT NULL
	)
`);

function getCachedHeaderParams(cacheKey = 'default') {
	const stmt = db.prepare(`
		SELECT cookie, request_verification_token, gs_request_token, updated_at
		FROM tgos_headers
		WHERE cache_key = ?
	`);
	const row = stmt.get(cacheKey);

	if (!row) {
		return null;
	}

	return {
		cookie: row.cookie,
		requestVerificationToken: row.request_verification_token || '',
		gsRequestToken: row.gs_request_token || '',
		updatedAt: row.updated_at,
	};
}

function saveHeaderParams(headerParams, cacheKey = 'default') {
	const stmt = db.prepare(`
		INSERT INTO tgos_headers (
			cache_key,
			cookie,
			request_verification_token,
			gs_request_token,
			updated_at
		)
		VALUES (?, ?, ?, ?, ?)
		ON CONFLICT(cache_key) DO UPDATE SET
			cookie = excluded.cookie,
			request_verification_token = excluded.request_verification_token,
			gs_request_token = excluded.gs_request_token,
			updated_at = excluded.updated_at
	`);

	stmt.run(
		cacheKey,
		headerParams.cookie,
		headerParams.requestVerificationToken || '',
		headerParams.gsRequestToken || '',
		new Date().toISOString()
	);
}

async function fetchHeaderParams() {
	const getResp = await axios.get(baseUrl, {
		headers: commonHeaders,
		validateStatus: () => true,
	});

	if (getResp.status >= 400) {
		throw new Error(`GET ${baseUrl} 失敗，status=${getResp.status}`);
	}

	const setCookie = getResp.headers['set-cookie'] || [];
	const cookie = setCookie.map((c) => c.split(';')[0]).join('; ');
	const $ = cheerio.load(getResp.data);

	const requestVerificationToken =
		$('input[name="__RequestVerificationToken"]').val() || '';

	const gsRequestToken =
		$('gs-request-token').attr('value') ||
		$('gs-request-token').text() ||
		'';

	if (!requestVerificationToken) {
		console.warn('找不到 __RequestVerificationToken');
	}

	if (!gsRequestToken) {
		console.warn('找不到 gs-request-token');
	}

	return {
		cookie,
		requestVerificationToken,
		gsRequestToken,
	};
}

function buildPostHeaders(address, headerParams) {
	const postHeaders = {
		...commonHeaders,
		'Accept': '*/*',
		'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
		'Origin': 'https://map.tgos.tw',
		'Referer': `${baseUrl}?addr=${encodeURIComponent(address)}`,
		'X-Requested-With': 'XMLHttpRequest',
		'Cookie': headerParams.cookie,
	};

	if (headerParams.requestVerificationToken) {
		postHeaders['gs-csrf-token'] = headerParams.requestVerificationToken;
	}

	if (headerParams.gsRequestToken) {
		postHeaders['gs-request-token'] = headerParams.gsRequestToken;
	}

	return postHeaders;
}

async function postAddressQuery(address, headerParams) {
	const formData = new URLSearchParams();
	formData.append('method', 'index.QueryAddr');
	formData.append('oAddress', address);

	const postResp = await axios.post(postUrl, formData.toString(), {
		headers: buildPostHeaders(address, headerParams),
		validateStatus: () => true,
	});

	if (postResp.status >= 400) {
		throw new Error(`POST ${postUrl} 失敗，status=${postResp.status}`);
	}

	return postResp.data;
}

async function queryTgosAddress(address) {
	let headerParams = getCachedHeaderParams();

	if (headerParams) {
		try {
			return await postAddressQuery(address, headerParams);
		} catch (error) {
			console.warn(`快取 header 失效，重新抓取後重試: ${address}`);
		}
	}

	headerParams = await fetchHeaderParams();
	saveHeaderParams(headerParams);

	return postAddressQuery(address, headerParams);
}

function sleep(ms) {
	return new Promise((resolve) => setTimeout(resolve, ms));
}

async function queryTgosAddressQueue(addresses, delayMs = 300) {
	const results = [];

	for (const address of addresses) {
		try {
			const result = await queryTgosAddress(address);
			results.push({ address, result });
		} catch (error) {
			results.push({ address, error: error.message });
		}

		await sleep(delayMs);
	}

	return results;
}

module.exports = {
	queryTgosAddress,
	queryTgosAddressQueue,
};
