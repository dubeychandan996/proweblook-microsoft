/* global Office, Excel, CustomFunctions */
/**
 * ProWebLook for Microsoft Excel — Shared Runtime & Custom Functions
 * Exposes window.PWL_Runtime for taskpane.js and registers =PROWEBLOOK.* formulas + Ribbon commands.
 */
(function () {
	'use strict';

	const STORAGE_KEY_API = 'pwl_excel_api_key';
	const STORAGE_KEY_INSTALL = 'pwl_excel_install_id';

	const SAMPLE_HEADERS = ['Name', 'Phone', 'Email', 'Country', 'IP Address'];
	const SAMPLE_ROWS = [
		['Aarav Mehta', '+919876543210', 'aarav@infosys.com', 'IN', '49.36.128.10'],
		['Sarah Jenkins', '+14155552671', 'sarah.jenkins@stripe.com', 'US', '8.8.8.8'],
		['Fake Buyer', '+12065550199', 'temp99@mailinator.com', 'US', '185.220.101.4'],
		["Liam O'Connor", '+447700900123', 'liam.oconnor@gmail.com', 'GB', '103.21.244.8'],
		['Spam Bot', '+19990000000', 'promo@bad-mailer-spam.biz', 'US', '45.155.205.9'],
		['Fatima Al Mansoori', '+971501234567', 'fatima@emaar.ae', 'AE', '94.200.200.10'],
	];

	const memoryCache = new Map();
	let creditsSaved = 0;

	function safeGetStorage(key) {
		try {
			return window.localStorage.getItem(key) || '';
		} catch (_) {
			return '';
		}
	}

	function safeSetStorage(key, val) {
		try {
			if (!val) window.localStorage.removeItem(key);
			else window.localStorage.setItem(key, val);
		} catch (_) {
			// Ignore storage restrictions in private webviews
		}
	}

	function getInstallId() {
		let id = safeGetStorage(STORAGE_KEY_INSTALL);
		if (!id) {
			const rnd =
				typeof crypto !== 'undefined' && crypto.randomUUID
					? crypto.randomUUID()
					: 'pwl-' + Math.random().toString(36).slice(2) + Date.now().toString(36);
			id = rnd;
			safeSetStorage(STORAGE_KEY_INSTALL, id);
		}
		return id;
	}

	function getApiKey() {
		return safeGetStorage(STORAGE_KEY_API).trim();
	}

	function setApiKey(nextKey) {
		const clean = String(nextKey || '').trim();
		safeSetStorage(STORAGE_KEY_API, clean);
		memoryCache.clear();
	}

	function clearCache() {
		memoryCache.clear();
		creditsSaved = 0;
	}

	function buildHeaders(overrideKey) {
		const headers = {
			'Content-Type': 'application/json',
			Accept: 'application/json',
			'X-PWL-Install-Id': getInstallId(),
		};
		const key = overrideKey !== undefined ? String(overrideKey || '').trim() : getApiKey();
		if (key) {
			headers['X-API-Key'] = key;
		}
		return headers;
	}

	async function apiPost(path, body, cacheKey) {
		const authScope = getApiKey() ? 'byok' : 'sandbox';
		const fullKey = cacheKey ? authScope + ':' + cacheKey : '';
		if (fullKey && memoryCache.has(fullKey)) {
			creditsSaved += 1;
			return memoryCache.get(fullKey);
		}

		const pending = fetch(path, {
			method: 'POST',
			headers: buildHeaders(),
			body: JSON.stringify(body || {}),
		}).then(async (res) => {
			const json = await res.json().catch(() => null);
			if (!res.ok || !json || json.ok === false) {
				const msg = (json && (json.error || json.message)) || 'HTTP ' + res.status;
				const err = new Error(msg);
				err.code = (json && json.code) || 'request_failed';
				err.status = res.status;
				throw err;
			}
			return json;
		});

		if (fullKey) {
			memoryCache.set(fullKey, pending);
			pending.catch(() => memoryCache.delete(fullKey));
		}
		return pending;
	}

	async function getStatus() {
		const res = await fetch('/api/excel/v1/status', {
			method: 'GET',
			headers: buildHeaders(),
		});
		return res.json();
	}

	async function getBalances(overrideKey) {
		const res = await fetch('/api/excel/v1/balances', {
			method: 'GET',
			headers: buildHeaders(overrideKey),
		});
		const json = await res.json().catch(() => null);
		if (!res.ok || !json || json.ok === false) {
			throw new Error((json && json.error) || 'Invalid API key or balance lookup failed.');
		}
		return json.data;
	}

	async function callSingle(signal, params) {
		const keyParts = [
			signal,
			String(params.phoneNumber || '').trim(),
			String(params.email || '').trim().toLowerCase(),
			String(params.target || '').trim().toLowerCase(),
			String(params.ip || '').trim(),
			String(params.countryCode || params.expectedCountry || '').trim().toUpperCase(),
			String(params.leadName || '').trim().toLowerCase(),
		].join('|');
		const resp = await apiPost('/api/excel/v1/single', { signal, ...params }, 'single:' + keyParts);
		return resp.data;
	}

	async function callVerifyLead(params) {
		const flags = params.signalFlags || {};
		const flagSig = [
			flags.phone !== false ? '1' : '0',
			flags.whatsapp !== false ? '1' : '0',
			flags.email !== false ? '1' : '0',
			flags.callerId !== false ? '1' : '0',
			flags.blacklist !== false ? '1' : '0',
			flags.ipGeo !== false ? '1' : '0',
		].join('');
		const keyParts = [
			flagSig,
			String(params.phoneNumber || '').trim(),
			String(params.email || '').trim().toLowerCase(),
			String(params.countryCode || '').trim().toUpperCase(),
			String(params.leadName || '').trim().toLowerCase(),
			String(params.ip || '').trim(),
			String(params.domain || '').trim().toLowerCase(),
		].join('|');
		const resp = await apiPost('/api/excel/v1/verify', params, 'verify:' + keyParts);
		return resp;
	}

	async function insertSampleSheet() {
		if (typeof Excel === 'undefined' || !Excel.run) {
			throw new Error('Excel host API is not active in this browser tab.');
		}
		await Excel.run(async (context) => {
			const sheets = context.workbook.worksheets;
			let sheet = sheets.getItemOrNullObject('ProWebLook Sample');
			await context.sync();

			if (sheet.isNullObject) {
				sheet = sheets.add('ProWebLook Sample');
			}
			sheet.activate();

			const headerRange = sheet.getRange('A1:E1');
			headerRange.values = [SAMPLE_HEADERS];
			headerRange.format.fill.color = '#0F172A';
			headerRange.format.font.color = '#FFFFFF';
			headerRange.format.font.bold = true;

			const dataRange = sheet.getRange('A2:E7');
			dataRange.values = SAMPLE_ROWS;
			// Force text formatting on Phone column so + prefixes stay intact in Excel
			const phoneCol = sheet.getRange('B2:B7');
			phoneCol.numberFormat = SAMPLE_ROWS.map(() => ['@']);

			sheet.getUsedRange().format.autofitColumns();
			await context.sync();
		});
	}

	function cleanCellStr(v) {
		if (v === null || v === undefined) return '';
		return String(v).trim();
	}

	function formatFormulaError(err) {
		if (!err) return 'Error: Verification failed';
		if (err.code === 'sandbox_exhausted' || err.code === 'starter_key_required') {
			return 'Connect free API key in ProWebLook pane';
		}
		if (err.code === 'invalid_credentials') {
			return 'Invalid API key (check ProWebLook pane)';
		}
		return 'Error: ' + (err.message || 'Lookup failed');
	}

	async function fnPhone(phone, country, field) {
		const p = cleanCellStr(phone);
		if (!p) return '';
		try {
			const d = await callSingle('phone', {
				phoneNumber: p,
				countryCode: cleanCellStr(country),
			});
			const f = cleanCellStr(field).toLowerCase() || 'summary';
			if (f === 'status' || f === 'verdict') return d.verdict.toUpperCase();
			if (f === 'linetype' || f === 'line_type' || f === 'type') return d.lineType || 'Unknown';
			if (f === 'carrier') return d.carrier || 'Unknown';
			if (f === 'formatted' || f === 'e164') return d.formattedNumber || d.phoneNumber;
			if (f === 'voip') return Boolean(d.isVoip);
			if (f === 'country') return d.countryCode || '';
			if (!d.valid) return 'INVALID';
			const tag = d.isVoip ? 'VOIP' : (d.lineType || 'VALID').toUpperCase();
			return d.carrier ? tag + ' (' + d.carrier + ')' : tag;
		} catch (err) {
			return formatFormulaError(err);
		}
	}

	async function fnWhatsapp(phone, country, field) {
		const p = cleanCellStr(phone);
		if (!p) return '';
		try {
			const d = await callSingle('whatsapp', {
				phoneNumber: p,
				countryCode: cleanCellStr(country),
			});
			const f = cleanCellStr(field).toLowerCase() || 'summary';
			if (f === 'registered') return Boolean(d.registered);
			if (f === 'business') return Boolean(d.isBusiness);
			if (f === 'type') return d.accountType || 'none';
			if (!d.registered) return 'NO_WHATSAPP';
			return d.isBusiness ? 'WHATSAPP_BUSINESS' : 'ON_WHATSAPP';
		} catch (err) {
			return formatFormulaError(err);
		}
	}

	async function fnEmail(email, field) {
		const e = cleanCellStr(email);
		if (!e) return '';
		try {
			const d = await callSingle('email', { email: e });
			const f = cleanCellStr(field).toLowerCase() || 'verdict';
			if (f === 'deliverable') return Boolean(d.deliverable);
			if (f === 'disposable') return Boolean(d.isDisposable);
			if (f === 'type') return d.emailType || '';
			if (f === 'domain') return d.domain || '';
			if (f === 'reason') return d.reasonCode || '';
			return String(d.verdict || 'unknown').toUpperCase();
		} catch (err) {
			return formatFormulaError(err);
		}
	}

	async function fnCallerId(phone, leadName, country) {
		const p = cleanCellStr(phone);
		if (!p) return '';
		try {
			const d = await callSingle('callerId', {
				phoneNumber: p,
				leadName: cleanCellStr(leadName),
				countryCode: cleanCellStr(country),
			});
			if (!d.found || !d.callerName) return 'NOT_FOUND';
			if (cleanCellStr(leadName) && d.nameMatch && d.nameMatch !== 'not_checked') {
				return d.callerName + ' [' + d.nameMatch.toUpperCase() + ']';
			}
			return d.callerName;
		} catch (err) {
			return formatFormulaError(err);
		}
	}

	async function fnDnsbl(domainOrIp, field) {
		const t = cleanCellStr(domainOrIp);
		if (!t) return '';
		try {
			const d = await callSingle('blacklist', { target: t });
			const f = cleanCellStr(field).toLowerCase() || 'summary';
			if (f === 'listed') return Boolean(d.listed);
			if (f === 'count') return Number(d.listedCount || 0);
			if (d.skippedFreeWebmail) return 'CLEAN (Webmail)';
			return d.listed ? 'LISTED (' + d.listedCount + ' DNSBLs)' : 'CLEAN';
		} catch (err) {
			return formatFormulaError(err);
		}
	}

	async function fnIpGeo(ip, countryHint) {
		const cleanIp = cleanCellStr(ip);
		if (!cleanIp) return '';
		try {
			const d = await callSingle('ipGeo', {
				ip: cleanIp,
				expectedCountry: cleanCellStr(countryHint),
			});
			if (d.skippedPrivateIp) return 'PRIVATE_IP';
			const loc = [d.city, d.countryCode].filter(Boolean).join(', ') || d.countryName || 'Unknown';
			const flags = [];
			if (d.countryMismatch) flags.push('MISMATCH vs ' + d.expectedCountry);
			if (d.isHostingOrProxy) flags.push('VPN/HOSTING');
			return flags.length ? loc + ' [' + flags.join(', ') + ']' : loc + (d.isp ? ' (' + d.isp + ')' : '');
		} catch (err) {
			return formatFormulaError(err);
		}
	}

	async function fnTrustScore(phone, email, country, leadName, ip) {
		const p = cleanCellStr(phone);
		const e = cleanCellStr(email);
		const cleanIp = cleanCellStr(ip);
		if (!p && !e && !cleanIp) return '';
		try {
			const resp = await callVerifyLead({
				phoneNumber: p,
				email: e,
				countryCode: cleanCellStr(country),
				leadName: cleanCellStr(leadName),
				ip: cleanIp,
			});
			const d = resp.data;
			return Number(d.trustScore);
		} catch (err) {
			return formatFormulaError(err);
		}
	}

	async function insertSampleLeadsCommand(event) {
		try {
			await insertSampleSheet();
		} catch (_) {
			// Ignore inside ribbon command
		} finally {
			if (event && typeof event.completed === 'function') {
				event.completed();
			}
		}
	}

	window.insertSampleLeadsCommand = insertSampleLeadsCommand;

	window.PWL_Runtime = {
		getApiKey,
		setApiKey,
		getInstallId,
		clearCache,
		getCacheStats: () => ({ entries: memoryCache.size, creditsSaved }),
		getStatus,
		getBalances,
		callSingle,
		callVerifyLead,
		insertSampleSheet,
		SAMPLE_HEADERS,
		SAMPLE_ROWS,
	};

	if (typeof CustomFunctions !== 'undefined' && CustomFunctions.associate) {
		CustomFunctions.associate('PHONE', fnPhone);
		CustomFunctions.associate('WHATSAPP', fnWhatsapp);
		CustomFunctions.associate('EMAIL', fnEmail);
		CustomFunctions.associate('CALLERID', fnCallerId);
		CustomFunctions.associate('DNSBL', fnDnsbl);
		CustomFunctions.associate('IPGEO', fnIpGeo);
		CustomFunctions.associate('TRUSTSCORE', fnTrustScore);
	}

	if (typeof Office !== 'undefined') {
		Office.onReady(() => {
			if (Office.actions && typeof Office.actions.associate === 'function') {
				Office.actions.associate('insertSampleLeadsCommand', insertSampleLeadsCommand);
			}
		});
	}
})();
