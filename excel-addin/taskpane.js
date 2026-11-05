/* global Office, Excel */
/**
 * ProWebLook for Microsoft Excel — Fluent 2 Taskpane Controller
 */
(function () {
	'use strict';

	const ENRICH_HEADERS = [
		'PWL Trust Score',
		'PWL Risk Level',
		'Phone Status',
		'Line Type',
		'Carrier',
		'WhatsApp',
		'Caller ID',
		'Email Verdict',
		'Domain DNSBL',
		'IP Geo',
		'Verification Summary',
	];

	function $(id) {
		return document.getElementById(id);
	}

	function colIndexToLetter(idx) {
		let n = idx + 1;
		let str = '';
		while (n > 0) {
			const rem = (n - 1) % 26;
			str = String.fromCharCode(65 + rem) + str;
			n = Math.floor((n - 1) / 26);
		}
		return str;
	}

	function showBanner(id, type, message) {
		const el = $(id);
		if (!el) return;
		if (!message) {
			el.className = 'pwl-banner';
			el.textContent = '';
			return;
		}
		el.className = 'pwl-banner is-visible ' + type;
		el.textContent = message;
	}

	function escapeHtml(s) {
		return String(s || '')
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;');
	}

	function isExcelHostActive() {
		return typeof Excel !== 'undefined' && typeof Excel.run === 'function';
	}

	async function refreshHeaderStatus() {
		const rt = window.PWL_Runtime;
		const pill = $('pwlModePill');
		if (!rt || !pill) return;

		const key = rt.getApiKey();
		if (key) {
			pill.textContent = 'Pro API Key';
			pill.classList.add('is-live');
		} else {
			pill.classList.remove('is-live');
			try {
				const st = await rt.getStatus();
				if (st && st.sandbox && st.sandbox.liveEnabled) {
					pill.textContent = 'Sandbox (' + st.sandbox.remaining + '/' + st.sandbox.limit + ')';
				} else {
					pill.textContent = 'Sample & Starter Mode';
				}
			} catch (_) {
				pill.textContent = 'Sample & Starter Mode';
			}
		}

		const stats = rt.getCacheStats();
		const cacheEl = $('pwlCacheStat');
		if (cacheEl) {
			cacheEl.textContent =
				stats.entries + ' cached lookups · ' + stats.creditsSaved + ' duplicate credits saved';
		}
	}

	function populateSelectOptions(headers) {
		const mapTargets = [
			{ id: 'mapPhone', match: /(phone|mobile|tel|cell|whatsapp|contact_num)/i },
			{ id: 'mapEmail', match: /(email|e-mail|mail)/i },
			{ id: 'mapName', match: /(name|customer|buyer|lead|contact)/i },
			{ id: 'mapCountry', match: /(country|iso|nation)/i },
			{ id: 'mapIp', match: /(\bip\b|ip_address|ipaddress|client_ip)/i },
		];

		mapTargets.forEach(({ id, match }) => {
			const sel = $(id);
			if (!sel) return;
			sel.innerHTML = '';
			const noneOpt = document.createElement('option');
			noneOpt.value = '-1';
			noneOpt.textContent = '— Not Mapped —';
			sel.appendChild(noneOpt);

			let matchedIdx = -1;
			headers.forEach((h, idx) => {
				const label = String(h || '').trim();
				if (!label || ENRICH_HEADERS.includes(label)) return;
				const opt = document.createElement('option');
				opt.value = String(idx);
				opt.textContent = colIndexToLetter(idx) + ': ' + label;
				sel.appendChild(opt);
				if (matchedIdx === -1 && match.test(label)) {
					matchedIdx = idx;
				}
			});
			sel.value = String(matchedIdx);
		});
	}

	async function scanWorksheetColumns(silent) {
		if (!isExcelHostActive()) {
			populateSelectOptions(window.PWL_Runtime.SAMPLE_HEADERS);
			if (!silent) {
				showBanner(
					'bulkBanner',
					'info',
					'Previewing outside Excel desktop/web: sample column mapping loaded. Use Quick Check tab to test live lookups here!'
				);
			}
			return;
		}

		try {
			await Excel.run(async (context) => {
				const sheet = context.workbook.worksheets.getActiveWorksheet();
				const used = sheet.getUsedRangeOrNullObject(true);
				used.load(['values', 'rowCount', 'columnCount']);
				await context.sync();

				if (used.isNullObject || !used.values || used.rowCount < 1) {
					populateSelectOptions([]);
					if (!silent) {
						showBanner(
							'bulkBanner',
							'info',
							'Active sheet is empty. Click "Insert Sample Sheet" to create a 6-row test table!'
						);
					}
					return;
				}

				const firstRow = used.values[0] || [];
				populateSelectOptions(firstRow);
				if (!silent) {
					showBanner(
						'bulkBanner',
						'success',
						'Detected ' + Math.max(0, used.rowCount - 1) + ' data row(s) across ' + used.columnCount + ' column(s).'
					);
				}
			});
		} catch (err) {
			showBanner('bulkBanner', 'error', 'Could not inspect worksheet: ' + (err.message || err));
		}
	}

	function getSelectedSignalFlags() {
		return {
			phone: $('sigPhone') ? $('sigPhone').checked : true,
			whatsapp: $('sigWhatsapp') ? $('sigWhatsapp').checked : true,
			email: $('sigEmail') ? $('sigEmail').checked : true,
			callerId: $('sigCallerId') ? $('sigCallerId').checked : true,
			blacklist: $('sigBlacklist') ? $('sigBlacklist').checked : true,
			ipGeo: $('sigIpGeo') ? $('sigIpGeo').checked : true,
		};
	}

	function buildRowOutputCells(result) {
		const s = (result && result.signals) || {};
		const p = s.phone;
		const w = s.whatsapp;
		const c = s.callerId;
		const e = s.email;
		const b = s.blacklist;
		const g = s.ipGeo;

		const phoneStatus = p ? p.verdict.toUpperCase() : '';
		const lineType = p ? (p.isVoip ? 'VoIP (High Risk)' : p.lineType || '') : '';
		const carrier = p ? p.carrier || '' : '';
		const waText = w ? (w.registered ? (w.isBusiness ? 'WhatsApp Business' : 'On WhatsApp') : 'No WhatsApp') : '';
		const cnamText = c
			? c.found
				? c.callerName + (c.nameMatch && c.nameMatch !== 'not_checked' ? ' [' + c.nameMatch.toUpperCase() + ']' : '')
				: 'Not Found'
			: '';
		const emailText = e ? e.verdict.toUpperCase() + (e.emailType ? ' (' + e.emailType + ')' : '') : '';
		const dnsblText = b
			? b.skippedFreeWebmail
				? 'Skipped (Webmail)'
				: b.listed
					? 'LISTED (' + b.listedCount + ')'
					: 'CLEAN'
			: '';
		const ipText = g
			? g.skippedPrivateIp
				? 'Private IP'
				: [g.city, g.countryCode].filter(Boolean).join(', ') +
					(g.isHostingOrProxy ? ' [VPN/Hosting]' : '') +
					(g.countryMismatch ? ' [Mismatch]' : '')
			: '';

		return [
			result.trustScore,
			String(result.riskLevel || '').toUpperCase(),
			phoneStatus,
			lineType,
			carrier,
			waText,
			cnamText,
			emailText,
			dnsblText,
			ipText,
			result.summary || '',
		];
	}

	async function runBulkEnrichment() {
		if (!isExcelHostActive()) {
			showBanner(
				'bulkBanner',
				'info',
				'Open this add-in inside Microsoft Excel (Desktop or Web) to enrich worksheet rows, or test right now in the Quick Check tab.'
			);
			return;
		}

		const idxPhone = Number($('mapPhone').value);
		const idxEmail = Number($('mapEmail').value);
		const idxName = Number($('mapName').value);
		const idxCountry = Number($('mapCountry').value);
		const idxIp = Number($('mapIp').value);

		if (idxPhone < 0 && idxEmail < 0 && idxIp < 0) {
			showBanner('bulkBanner', 'error', 'Map at least Phone, Email, or IP Address column before running verification.');
			return;
		}

		const btn = $('btnRunBulk');
		const progWrap = $('bulkProgressWrap');
		const progFill = $('bulkProgressFill');
		const progLabel = $('bulkProgressLabel');
		const progCount = $('bulkProgressCount');

		btn.disabled = true;
		showBanner('bulkBanner', '', '');
		progWrap.classList.add('is-visible');
		progFill.style.width = '2%';
		progLabel.textContent = 'Reading Excel worksheet...';
		progCount.textContent = '';

		try {
			let rowsData = [];
			let startOutCol = 0;
			let startRowIndex = 0;
			let startColIndex = 0;

			await Excel.run(async (context) => {
				const sheet = context.workbook.worksheets.getActiveWorksheet();
				const used = sheet.getUsedRangeOrNullObject(true);
				used.load(['values', 'rowCount', 'columnCount', 'rowIndex', 'columnIndex']);
				await context.sync();

				if (used.isNullObject || used.rowCount < 2) {
					throw new Error('Add at least 1 data row under your header row (or click Insert Sample Sheet).');
				}

				startRowIndex = used.rowIndex;
				startColIndex = used.columnIndex;
				const headers = used.values[0].map((v) => String(v || '').trim());
				const existingIdx = headers.indexOf(ENRICH_HEADERS[0]);
				startOutCol = existingIdx >= 0 ? startColIndex + existingIdx : startColIndex + used.columnCount;

				const outHeaderRange = sheet.getRangeByIndexes(
					startRowIndex,
					startOutCol,
					1,
					ENRICH_HEADERS.length
				);
				outHeaderRange.values = [ENRICH_HEADERS];
				outHeaderRange.format.fill.color = '#0F172A';
				outHeaderRange.format.font.color = '#FFFFFF';
				outHeaderRange.format.font.bold = true;
				await context.sync();

				rowsData = used.values.slice(1);
			});

			const flags = getSelectedSignalFlags();
			const rt = window.PWL_Runtime;
			let trustedCount = 0;
			let reviewCount = 0;
			let highRiskCount = 0;

			const total = rowsData.length;
			const CHUNK_SIZE = 4;

			for (let offset = 0; offset < total; offset += CHUNK_SIZE) {
				const slice = rowsData.slice(offset, offset + CHUNK_SIZE);
				const chunkResults = await Promise.all(
					slice.map(async (row) => {
						const phoneNumber = idxPhone >= 0 ? String(row[idxPhone] || '').trim() : '';
						const email = idxEmail >= 0 ? String(row[idxEmail] || '').trim() : '';
						const leadName = idxName >= 0 ? String(row[idxName] || '').trim() : '';
						const countryCode = idxCountry >= 0 ? String(row[idxCountry] || '').trim() : '';
						const ip = idxIp >= 0 ? String(row[idxIp] || '').trim() : '';

						if (!phoneNumber && !email && !ip) {
							return null;
						}
						const resp = await rt.callVerifyLead({
							phoneNumber,
							email,
							leadName,
							countryCode,
							ip,
							signalFlags: flags,
						});
						return resp.data;
					})
				);

				await Excel.run(async (context) => {
					const sheet = context.workbook.worksheets.getActiveWorksheet();
					const outMatrix = chunkResults.map((res) =>
						res ? buildRowOutputCells(res) : ENRICH_HEADERS.map(() => '')
					);

					const writeRange = sheet.getRangeByIndexes(
						startRowIndex + 1 + offset,
						startOutCol,
						slice.length,
						ENRICH_HEADERS.length
					);
					writeRange.values = outMatrix;

					chunkResults.forEach((res, idxInChunk) => {
						if (!res) return;
						const scoreRange = sheet.getRangeByIndexes(
							startRowIndex + 1 + offset + idxInChunk,
							startOutCol,
							1,
							2
						);
						if (res.riskLevel === 'trusted') {
							trustedCount += 1;
							scoreRange.format.fill.color = '#DFF6DD';
							scoreRange.format.font.color = '#107C10';
						} else if (res.riskLevel === 'review') {
							reviewCount += 1;
							scoreRange.format.fill.color = '#FFF4CE';
							scoreRange.format.font.color = '#8A5700';
						} else {
							highRiskCount += 1;
							scoreRange.format.fill.color = '#FDE7E9';
							scoreRange.format.font.color = '#A80000';
						}
						scoreRange.format.font.bold = true;
					});

					await context.sync();
				});

				const done = Math.min(total, offset + slice.length);
				const pct = Math.round((done / total) * 100);
				progFill.style.width = pct + '%';
				progLabel.textContent = 'Enriching worksheet...';
				progCount.textContent = done + ' / ' + total;
			}

			await Excel.run(async (context) => {
				const sheet = context.workbook.worksheets.getActiveWorksheet();
				sheet.getUsedRange().format.autofitColumns();
				await context.sync();
			});

			await refreshHeaderStatus();
			const cacheStats = rt.getCacheStats();
			showBanner(
				'bulkBanner',
				'success',
				'Enriched ' +
					total +
					' row(s): ' +
					trustedCount +
					' Trusted, ' +
					reviewCount +
					' Review, ' +
					highRiskCount +
					' High Risk (' +
					cacheStats.creditsSaved +
					' duplicate credits saved).'
			);
		} catch (err) {
			showBanner('bulkBanner', 'error', err.message || 'Bulk verification failed.');
		} finally {
			btn.disabled = false;
		}
	}

	function renderQuickResult(data) {
		const container = $('quickResultBox');
		if (!container) return;

		const s = data.signals || {};
		const rows = [];
		if (s.phone) {
			rows.push([
				'Phone & Carrier',
				s.phone.verdict.toUpperCase() +
					' · ' +
					(s.phone.isVoip ? 'VoIP' : s.phone.lineType) +
					(s.phone.carrier ? ' (' + s.phone.carrier + ')' : ''),
			]);
		}
		if (s.whatsapp) {
			rows.push([
				'Silent WhatsApp',
				s.whatsapp.registered
					? s.whatsapp.isBusiness
						? 'Active (Business)'
						: 'Active (Personal)'
					: 'Not on WhatsApp',
			]);
		}
		if (s.callerId) {
			rows.push([
				'Caller ID (CNAM)',
				s.callerId.found
					? s.callerId.callerName +
						(s.callerId.nameMatch && s.callerId.nameMatch !== 'not_checked'
							? ' [' + s.callerId.nameMatch.toUpperCase() + ']'
							: '')
					: 'Not Found',
			]);
		}
		if (s.email) {
			rows.push(['Email Verdict', s.email.verdict.toUpperCase() + ' (' + s.email.emailType + ')']);
		}
		if (s.blacklist) {
			rows.push([
				'Domain DNSBL',
				s.blacklist.listed ? 'LISTED (' + s.blacklist.listedCount + ' lists)' : 'CLEAN',
			]);
		}
		if (s.ipGeo && !s.ipGeo.skippedPrivateIp) {
			rows.push([
				'Edge GeoIP',
				[s.ipGeo.city, s.ipGeo.countryCode].filter(Boolean).join(', ') +
					(s.ipGeo.isHostingOrProxy ? ' · VPN/Hosting' : ''),
			]);
		}

		const kvHtml = rows
			.map(
				([k, v]) =>
					'<div class="pwl-kv-row"><span>' + escapeHtml(k) + '</span><span>' + escapeHtml(v) + '</span></div>'
			)
			.join('');

		container.innerHTML =
			'<div class="pwl-score-card ' +
			escapeHtml(data.riskLevel) +
			'">' +
			'<div class="pwl-score-top">' +
			'<div><div style="font-size:10.5px;color:#475569;font-weight:600;">Composite Trust Score</div>' +
			'<div class="pwl-score-num">' +
			Number(data.trustScore) +
			' / 100</div></div>' +
			'<span class="pwl-risk-badge">' +
			escapeHtml(String(data.riskLevel || '').replace('_', ' ')) +
			'</span></div>' +
			'<div style="font-size:11.5px;color:#334155;">' +
			escapeHtml(data.summary) +
			'</div>' +
			'<div class="pwl-kv-list">' +
			kvHtml +
			'</div></div>';
	}

	async function runQuickCheck(ev) {
		if (ev) ev.preventDefault();
		const btn = $('btnQuickVerify');
		btn.disabled = true;
		showBanner('quickBanner', '', '');

		try {
			const resp = await window.PWL_Runtime.callVerifyLead({
				phoneNumber: $('qPhone').value,
				email: $('qEmail').value,
				leadName: $('qName').value,
				countryCode: $('qCountry').value,
				ip: $('qIp').value,
			});
			renderQuickResult(resp.data);
			await refreshHeaderStatus();
		} catch (err) {
			showBanner('quickBanner', 'error', err.message || 'Verification failed.');
		} finally {
			btn.disabled = false;
		}
	}

	function renderBalanceTiles(bal) {
		const grid = $('balanceGrid');
		if (!grid) return;
		if (!bal) {
			grid.innerHTML = '';
			return;
		}
		const items = [
			['Phone Credits', bal.phoneCredits],
			['WhatsApp Credits', bal.whatsappCredits],
			['Email Credits', bal.emailCredits],
			['Caller ID Credits', bal.callerIdCredits],
			['DNSBL Credits', bal.blacklistCredits],
			['Edge GeoIP / mo', bal.ipGeoMonthlyFreeQuota || 20000],
		];
		grid.innerHTML = items
			.map(
				([label, val]) =>
					'<div class="pwl-bal-tile"><div class="pwl-bal-label">' +
					escapeHtml(label) +
					'</div><div class="pwl-bal-val">' +
					Number(val || 0).toLocaleString() +
					'</div></div>'
			)
			.join('');
	}

	async function handleSaveApiKey(ev) {
		if (ev) ev.preventDefault();
		const input = $('apiKeyInput');
		const raw = String(input.value || '').trim();
		const rt = window.PWL_Runtime;

		if (!raw) {
			rt.setApiKey('');
			renderBalanceTiles(null);
			await refreshHeaderStatus();
			showBanner('settingsBanner', 'info', 'Switched back to Sample & Free Starter Sandbox mode.');
			return;
		}

		const btn = $('btnSaveKey');
		btn.disabled = true;
		showBanner('settingsBanner', '', '');
		try {
			const bal = await rt.getBalances(raw);
			rt.setApiKey(raw);
			renderBalanceTiles(bal);
			await refreshHeaderStatus();
			showBanner(
				'settingsBanner',
				'success',
				'Connected ProWebLook API key' + (bal.accountEmail ? ' (' + bal.accountEmail + ')' : '') + '!'
			);
		} catch (err) {
			showBanner('settingsBanner', 'error', err.message || 'Invalid ProWebLook API key.');
		} finally {
			btn.disabled = false;
		}
	}

	function wireTabs() {
		const tabs = document.querySelectorAll('.pwl-tab');
		const panels = document.querySelectorAll('.pwl-panel');
		tabs.forEach((tab) => {
			tab.addEventListener('click', () => {
				const target = tab.getAttribute('data-tab');
				tabs.forEach((t) => t.classList.toggle('is-active', t === tab));
				panels.forEach((p) => p.classList.toggle('is-active', p.id === 'panel-' + target));
			});
		});
	}

	function wireFormulaCopyButtons() {
		document.querySelectorAll('[data-copy-formula]').forEach((btn) => {
			btn.addEventListener('click', async () => {
				const text = btn.getAttribute('data-copy-formula') || '';
				try {
					await navigator.clipboard.writeText(text);
					const prev = btn.textContent;
					btn.textContent = 'Copied!';
					setTimeout(() => {
						btn.textContent = prev;
					}, 1300);
				} catch (_) {
					$('qPhone').focus();
				}
			});
		});
	}

	async function initTaskpane() {
		wireTabs();
		wireFormulaCopyButtons();

		const rt = window.PWL_Runtime;
		const existingKey = rt.getApiKey();
		if (existingKey && $('apiKeyInput')) {
			$('apiKeyInput').value = existingKey;
			rt.getBalances(existingKey)
				.then(renderBalanceTiles)
				.catch(() => {});
		}

		$('btnInsertSample').addEventListener('click', async () => {
			try {
				await rt.insertSampleSheet();
				await scanWorksheetColumns(true);
				showBanner(
					'bulkBanner',
					'success',
					'Inserted 6-row "ProWebLook Sample" sheet and auto-mapped columns! Click Verify & Enrich Sheet below.'
				);
			} catch (err) {
				showBanner('bulkBanner', 'info', err.message || 'Could not insert sample sheet.');
			}
		});

		$('btnScanCols').addEventListener('click', () => scanWorksheetColumns(false));
		$('btnRunBulk').addEventListener('click', runBulkEnrichment);

		$('quickForm').addEventListener('submit', runQuickCheck);
		$('btnFillSampleQuick').addEventListener('click', () => {
			$('qPhone').value = '+919876543210';
			$('qEmail').value = 'aarav@infosys.com';
			$('qName').value = 'Aarav Mehta';
			$('qCountry').value = 'IN';
			$('qIp').value = '49.36.128.10';
			runQuickCheck();
		});

		$('settingsForm').addEventListener('submit', handleSaveApiKey);
		$('btnClearKey').addEventListener('click', async () => {
			$('apiKeyInput').value = '';
			await handleSaveApiKey();
		});
		$('btnClearCache').addEventListener('click', async () => {
			rt.clearCache();
			await refreshHeaderStatus();
			showBanner('settingsBanner', 'info', 'Cleared formula & bulk deduplication cache.');
		});

		await refreshHeaderStatus();
		await scanWorksheetColumns(true);
	}

	if (typeof Office !== 'undefined' && Office.onReady) {
		Office.onReady(() => {
			initTaskpane();
		});
	} else if (document.readyState === 'loading') {
		document.addEventListener('DOMContentLoaded', initTaskpane);
	} else {
		initTaskpane();
	}
})();
