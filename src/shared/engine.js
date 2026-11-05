/**
 * Unified ProWebLook 6-Signal Verification & Trust-Scoring Engine
 * Powering both:
 *  - ProWebLook for Excel (Office.js Custom Functions + Bulk Taskpane)
 *  - Microsoft Power Automate, Logic Apps & Copilot Studio Custom Connector Gateway
 */

import {
	FREE_WEBMAIL_DOMAINS,
	HOSTING_PROXY_KEYWORDS,
	compareCallerIdWithLead,
	extractEmailDomain,
	isDisposableEmailDomain,
	isPrivateOrReservedIp,
	isValidIpAddress,
	normalizeDomainInput,
	normalizeEmailInput,
	normalizePhoneInput,
} from './parser.js';

export const DEFAULT_LOOKUP_BASE = 'https://lookup.proweblook.com/api/v1';
export const DEFAULT_MONOLITH_BASE = 'https://proweblook.com/api/v1';
export const DEFAULT_IPAPI_BASE = 'https://ipapi.proweblook.com';

export function createServiceError(message, code, httpStatus = 400) {
	const err = new Error(message);
	err.code = code;
	err.httpStatus = httpStatus;
	return err;
}

async function fetchJsonWithTimeout(url, init = {}, fetchImpl = fetch, timeoutMs = 16000) {
	const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
	const timer = controller ? setTimeout(() => controller.abort(), timeoutMs) : null;
	try {
		const res = await fetchImpl(url, controller ? { ...init, signal: controller.signal } : init);
		const data = await res.json().catch(() => null);
		return { status: res.status, ok: res.ok, data };
	} finally {
		if (timer) clearTimeout(timer);
	}
}

export async function validatePhone({
	phoneNumber,
	countryCode = '',
	apiKey,
	lookupBaseUrl = DEFAULT_LOOKUP_BASE,
	fetchImpl = fetch,
}) {
	const cleanKey = String(apiKey || '').trim();
	if (!cleanKey) throw createServiceError('Missing ProWebLook API key.', 'invalid_credentials', 401);

	const cleanPhone = normalizePhoneInput(phoneNumber);
	if (!cleanPhone) throw createServiceError('Enter a valid phone number (7–15 digits).', 'invalid_input', 400);

	const body = { api_key: cleanKey, phone_number: cleanPhone };
	const cc = String(countryCode || '').trim().toUpperCase();
	if (cc && /^[A-Z]{2}$/.test(cc) && !cleanPhone.startsWith('+')) {
		body.country_code = cc;
	}

	const { status, ok, data } = await fetchJsonWithTimeout(
		`${lookupBaseUrl.replace(/\/+$/, '')}/phonevalidation`,
		{
			method: 'POST',
			headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
			body: JSON.stringify(body),
		},
		fetchImpl,
	);

	if (status === 401 || status === 403) throw createServiceError('Invalid ProWebLook API key.', 'invalid_credentials', 401);
	if (status === 402) throw createServiceError('Insufficient phone validation credits.', 'insufficient_credit', 402);
	if (!ok || !data) {
		throw createServiceError(data?.error || `Phone validation HTTP ${status}.`, 'provider_error', status || 502);
	}

	const payload = data.data && typeof data.data === 'object' ? data.data : data;
	const rawStatus = String(payload.status || '').toUpperCase();
	const isValid = rawStatus === 'VALID_CONFIRMED' || rawStatus === 'VALID_UNCONFIRMED';
	const verdict = isValid ? 'valid' : rawStatus === 'INVALID' ? 'invalid' : 'unknown';
	const lineType = String(payload.linetype || payload.line_type || 'Unknown');
	const lowerLine = lineType.toLowerCase();
	const isVoip = lowerLine.includes('voip') || lowerLine.includes('virtual');
	const isLandline = lowerLine === 'fixed line' || lowerLine === 'landline';

	return {
		valid: isValid,
		status: rawStatus || verdict.toUpperCase(),
		verdict,
		phoneNumber: cleanPhone,
		formattedNumber: String(payload.formatted_number || payload.formatinternational || cleanPhone),
		nationalFormat: String(payload.formatnational || ''),
		lineType,
		isVoip,
		isLandline,
		carrier: String(payload.carrier || ''),
		location: String(payload.location || ''),
		countryCode: String(payload.countrycode || cc || ''),
		countryName: String(payload.countryname || ''),
		timezone: Array.isArray(payload.timezone) ? String(payload.timezone[0] || '') : String(payload.timezone || ''),
	};
}

export async function checkWhatsapp({
	phoneNumber,
	countryCode = '',
	apiKey,
	lookupBaseUrl = DEFAULT_LOOKUP_BASE,
	fetchImpl = fetch,
}) {
	const cleanKey = String(apiKey || '').trim();
	if (!cleanKey) throw createServiceError('Missing ProWebLook API key.', 'invalid_credentials', 401);

	const cleanPhone = normalizePhoneInput(phoneNumber);
	if (!cleanPhone) throw createServiceError('Enter a valid phone number (7–15 digits).', 'invalid_input', 400);

	const body = { api_key: cleanKey, phone_number: cleanPhone };
	const cc = String(countryCode || '').trim().toUpperCase();
	if (cc && /^[A-Z]{2}$/.test(cc) && !cleanPhone.startsWith('+')) {
		body.country_code = cc;
	}

	const { status, ok, data } = await fetchJsonWithTimeout(
		`${lookupBaseUrl.replace(/\/+$/, '')}/wavalidation`,
		{
			method: 'POST',
			headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
			body: JSON.stringify(body),
		},
		fetchImpl,
	);

	if (status === 401 || status === 403) throw createServiceError('Invalid ProWebLook API key.', 'invalid_credentials', 401);
	if (status === 402) throw createServiceError('Insufficient WhatsApp verification credits.', 'insufficient_credit', 402);
	if (!ok || !data) {
		throw createServiceError(data?.error || `WhatsApp verification HTTP ${status}.`, 'provider_error', status || 502);
	}

	const payload = data.data && typeof data.data === 'object' ? data.data : data;
	const registered = payload.numberstatus === true || payload.registered === true || payload.is_whatsapp === true;
	const isBusiness = payload.businessnumber === true || payload.is_business === true;
	const accountType = registered ? (isBusiness ? 'business' : 'personal') : 'none';

	return {
		phoneNumber: cleanPhone,
		registered,
		isBusiness,
		accountType,
		verdict: registered ? 'on_whatsapp' : 'not_on_whatsapp',
		countryCode: String(payload.countrycode || cc || ''),
	};
}

export async function verifyEmail({
	email,
	apiKey,
	lookupBaseUrl = DEFAULT_LOOKUP_BASE,
	fetchImpl = fetch,
}) {
	const cleanKey = String(apiKey || '').trim();
	if (!cleanKey) throw createServiceError('Missing ProWebLook API key.', 'invalid_credentials', 401);

	const cleanEmail = normalizeEmailInput(email);
	if (!cleanEmail) throw createServiceError('Enter a syntactically valid email address.', 'invalid_input', 400);

	const localDisposable = isDisposableEmailDomain(cleanEmail);
	const emailDomain = extractEmailDomain(cleanEmail);
	const isFreeWebmail = emailDomain ? FREE_WEBMAIL_DOMAINS.has(emailDomain) : false;

	const params = new URLSearchParams({ api_key: cleanKey, email: cleanEmail });
	const { status, ok, data } = await fetchJsonWithTimeout(
		`${lookupBaseUrl.replace(/\/+$/, '')}/emailvalidation?${params.toString()}`,
		{ method: 'GET', headers: { Accept: 'application/json' } },
		fetchImpl,
	);

	if (status === 401 || status === 403) throw createServiceError('Invalid ProWebLook API key.', 'invalid_credentials', 401);
	if (status === 402) throw createServiceError('Insufficient email verification credits.', 'insufficient_credit', 402);
	if (!ok || !data) {
		throw createServiceError(data?.error || `Email verification HTTP ${status}.`, 'provider_error', status || 502);
	}

	const payload = data.data && typeof data.data === 'object' ? data.data : data;
	const reasonCode = String(payload.reason_code || '');
	const isDisposable =
		localDisposable || reasonCode.toLowerCase().includes('disposable') || payload.disposable === true || payload.is_disposable === true;

	let verdict = String(payload.verdict || '').toLowerCase();
	if (!['deliverable', 'undeliverable', 'unknown', 'risky'].includes(verdict)) {
		if (payload.emailexist === false || payload.emailformat === false) {
			verdict = 'undeliverable';
		} else if (payload.emailexist === true) {
			verdict = 'deliverable';
		} else {
			verdict = 'unknown';
		}
	}

	const finalVerdict = isDisposable ? 'disposable' : verdict;
	return {
		email: cleanEmail,
		domain: emailDomain,
		deliverable: finalVerdict === 'deliverable',
		verdict: finalVerdict,
		isDisposable,
		isFreeWebmail,
		emailType: isDisposable ? 'disposable' : isFreeWebmail ? 'free_webmail' : 'corporate',
		reasonCode: reasonCode || (isDisposable ? 'disposable_domain' : 'completed'),
		emailFormatValid: payload.emailformat !== false,
		mailboxExists: Boolean(payload.emailexist),
	};
}

export async function lookupCallerId({
	phoneNumber,
	countryCode = '',
	leadName = '',
	companyName = '',
	apiKey,
	lookupBaseUrl = DEFAULT_LOOKUP_BASE,
	fetchImpl = fetch,
}) {
	const cleanKey = String(apiKey || '').trim();
	if (!cleanKey) throw createServiceError('Missing ProWebLook API key.', 'invalid_credentials', 401);

	const cleanPhone = normalizePhoneInput(phoneNumber);
	if (!cleanPhone) throw createServiceError('Enter a valid phone number (7–15 digits).', 'invalid_input', 400);

	const params = new URLSearchParams({ api_key: cleanKey, phone_number: cleanPhone });
	const cc = String(countryCode || '').trim().toUpperCase();
	if (cc && /^[A-Z]{2}$/.test(cc) && !cleanPhone.startsWith('+')) {
		params.set('country_code', cc);
	}

	const { status, ok, data } = await fetchJsonWithTimeout(
		`${lookupBaseUrl.replace(/\/+$/, '')}/calleridvalidation?${params.toString()}`,
		{
			method: 'GET',
			headers: { Accept: 'application/json', 'X-API-KEY': cleanKey },
		},
		fetchImpl,
	);

	if (status === 401 || status === 403) throw createServiceError('Invalid ProWebLook API key.', 'invalid_credentials', 401);
	if (status === 402) throw createServiceError('Insufficient Caller ID credits.', 'insufficient_credit', 402);
	if (!ok || !data) {
		throw createServiceError(data?.error || `Caller ID lookup HTTP ${status}.`, 'provider_error', status || 502);
	}

	const payload = data.data && typeof data.data === 'object' ? data.data : data;
	const callerName = String(payload.caller_name || payload.name || '').trim();
	const found = Boolean(callerName);
	const nameMatch = compareCallerIdWithLead(callerName, leadName, companyName);

	return {
		phoneNumber: cleanPhone,
		found,
		verdict: found ? 'found' : 'not_found',
		callerName,
		callerType: String(payload.caller_type || ''),
		nameMatch,
		carrier: String(payload.carrier || ''),
		lineType: String(payload.linetype || payload.line_type || ''),
		location: String(payload.location || ''),
		countryCode: String(payload.countrycode || cc || ''),
	};
}

export async function checkBlacklist({
	target,
	targetType = 'auto',
	skipFreeWebmail = true,
	apiKey,
	lookupBaseUrl = DEFAULT_LOOKUP_BASE,
	fetchImpl = fetch,
}) {
	const cleanKey = String(apiKey || '').trim();
	if (!cleanKey) throw createServiceError('Missing ProWebLook API key.', 'invalid_credentials', 401);

	const raw = String(target || '').trim();
	if (!raw) throw createServiceError('Enter a domain name, email address, or IP address.', 'invalid_input', 400);

	let resolvedTarget = raw;
	let resolvedType = ['ip', 'domain', 'url', 'auto'].includes(targetType) ? targetType : 'auto';

	if (isValidIpAddress(raw)) {
		resolvedTarget = raw;
		resolvedType = 'ip';
	} else {
		const dom = normalizeDomainInput(raw);
		if (!dom) {
			throw createServiceError('Enter a valid domain name or IP address.', 'invalid_input', 400);
		}
		resolvedTarget = dom;
		if (resolvedType === 'auto') resolvedType = 'domain';

		if (skipFreeWebmail && FREE_WEBMAIL_DOMAINS.has(dom)) {
			return {
				target: dom,
				targetType: 'domain',
				listed: false,
				listedCount: 0,
				checkedCount: 0,
				skippedFreeWebmail: true,
				verdict: 'clean_free_webmail',
			};
		}
	}

	const params = new URLSearchParams({
		api_key: cleanKey,
		target: resolvedTarget,
		target_type: resolvedType,
	});
	const { status, ok, data } = await fetchJsonWithTimeout(
		`${lookupBaseUrl.replace(/\/+$/, '')}/blacklistchecker?${params.toString()}`,
		{ method: 'GET', headers: { Accept: 'application/json' } },
		fetchImpl,
	);

	if (status === 401 || status === 403) throw createServiceError('Invalid ProWebLook API key.', 'invalid_credentials', 401);
	if (status === 402) throw createServiceError('Insufficient blacklist credits.', 'insufficient_credit', 402);
	if (!ok || !data) {
		throw createServiceError(data?.error || `Blacklist checker HTTP ${status}.`, 'provider_error', status || 502);
	}

	const payload = data.data && typeof data.data === 'object' ? data.data : data;
	const listedCount = Number(payload.listed_count ?? payload.listedCount ?? (payload.listed === true ? 1 : 0));
	const listed = listedCount > 0 || payload.listed === true;

	return {
		target: resolvedTarget,
		targetType: resolvedType,
		listed,
		listedCount,
		checkedCount: Number(payload.total_checked || payload.sources_checked || 0),
		skippedFreeWebmail: false,
		verdict: listed ? 'listed' : 'clean',
	};
}

export async function lookupIpGeo({
	ip,
	expectedCountry = '',
	apiKey,
	ipapiBaseUrl = DEFAULT_IPAPI_BASE,
	fetchImpl = fetch,
}) {
	const cleanKey = String(apiKey || '').trim();
	if (!cleanKey) throw createServiceError('Missing ProWebLook API key.', 'invalid_credentials', 401);

	const cleanIp = String(ip || '').trim();
	if (!cleanIp || !isValidIpAddress(cleanIp)) {
		throw createServiceError('Enter a valid IPv4 or IPv6 address.', 'invalid_input', 400);
	}

	if (isPrivateOrReservedIp(cleanIp)) {
		return {
			ip: cleanIp,
			countryCode: '',
			countryName: '',
			region: '',
			city: '',
			isp: 'Private / Local Network',
			asn: '',
			timezone: '',
			isHostingOrProxy: false,
			expectedCountry: String(expectedCountry || '').trim().toUpperCase(),
			countryMismatch: false,
			skippedPrivateIp: true,
			verdict: 'private_ip_skipped',
		};
	}

	const params = new URLSearchParams({ api_key: cleanKey, ip: cleanIp });
	const { status, ok, data } = await fetchJsonWithTimeout(
		`${ipapiBaseUrl.replace(/\/+$/, '')}/?${params.toString()}`,
		{ method: 'GET', headers: { Accept: 'application/json' } },
		fetchImpl,
		10000,
	);

	if (status === 401 || status === 403) throw createServiceError('Invalid ProWebLook API key.', 'invalid_credentials', 401);
	if (!ok || !data) {
		throw createServiceError(data?.error || `Edge GeoIP HTTP ${status}.`, 'provider_error', status || 502);
	}

	const payload = data.data && typeof data.data === 'object' ? data.data : data;
	const rawCc = String(payload.country_code || payload.countryCode || payload.country || '')
		.trim()
		.toUpperCase()
		.slice(0, 2);
	const countryCode = /^[A-Z]{2}$/.test(rawCc) ? rawCc : '';
	const isp = String(payload.isp || payload.org || payload.as_org || '');
	const asn = String(payload.asn || payload.as || '');
	const combinedOrg = `${isp} ${asn}`.toLowerCase();
	const isHostingOrProxy =
		Boolean(payload.proxy || payload.hosting || payload.vpn) ||
		HOSTING_PROXY_KEYWORDS.some((kw) => combinedOrg.includes(kw));

	const expCc = String(expectedCountry || '').trim().toUpperCase().slice(0, 2);
	const countryMismatch = Boolean(countryCode && /^[A-Z]{2}$/.test(expCc) && countryCode !== expCc);

	return {
		ip: String(payload.ip || cleanIp),
		countryCode,
		countryName: String(payload.country_name || payload.country || ''),
		region: String(payload.region || payload.region_name || ''),
		city: String(payload.city || ''),
		isp,
		asn,
		timezone: String(payload.timezone || ''),
		isHostingOrProxy,
		expectedCountry: expCc,
		countryMismatch,
		skippedPrivateIp: false,
		verdict: isHostingOrProxy ? 'hosting_or_vpn' : countryMismatch ? 'country_mismatch' : 'residential_isp',
	};
}

export function calculateLeadTrustScore({ signals = {} }) {
	let score = 65;
	const reasons = [];
	const positives = [];

	const p = signals.phone;
	if (p) {
		if (p.verdict === 'invalid') {
			score -= 35;
			reasons.push('Invalid phone number');
		} else if (p.isVoip) {
			score -= 18;
			reasons.push(`VoIP / virtual phone line (${p.carrier || p.lineType})`);
		} else if (p.verdict === 'valid') {
			score += 12;
			positives.push(`Verified ${p.lineType || 'phone'}${p.carrier ? ` (${p.carrier})` : ''}`);
		}
	}

	const w = signals.whatsapp;
	if (w) {
		if (w.registered) {
			score += w.isBusiness ? 12 : 10;
			positives.push(w.isBusiness ? 'Active WhatsApp Business' : 'Active WhatsApp');
		} else {
			score -= 10;
			reasons.push('Not registered on WhatsApp');
		}
	}

	const e = signals.email;
	if (e) {
		if (e.isDisposable || e.verdict === 'disposable') {
			score -= 38;
			reasons.push(`Disposable email domain (${e.domain})`);
		} else if (e.verdict === 'undeliverable') {
			score -= 28;
			reasons.push(`Undeliverable email (${e.reasonCode || 'bounces'})`);
		} else if (e.verdict === 'deliverable') {
			score += e.emailType === 'corporate' ? 12 : 10;
			positives.push(`Deliverable ${e.emailType === 'corporate' ? 'corporate ' : ''}email`);
		}
	}

	const c = signals.callerId;
	if (c) {
		if (c.nameMatch === 'match') {
			score += 12;
			positives.push(`Caller ID matches "${c.callerName}"`);
		} else if (c.nameMatch === 'partial') {
			score += 5;
			positives.push(`Caller ID partial match "${c.callerName}"`);
		} else if (c.nameMatch === 'mismatch') {
			score -= 20;
			reasons.push(`Caller ID mismatch (registered to "${c.callerName}")`);
		} else if (c.found) {
			score += 4;
			positives.push(`Caller ID found (${c.callerName})`);
		}
	}

	const b = signals.blacklist;
	if (b && !b.skippedFreeWebmail) {
		if (b.listed) {
			score -= 25;
			reasons.push(`${b.target} listed on ${b.listedCount} DNSBL(s)`);
		} else {
			score += 4;
			positives.push(`${b.target} clean across DNSBLs`);
		}
	}

	const g = signals.ipGeo;
	if (g && !g.skippedPrivateIp) {
		if (g.countryMismatch) {
			score -= 16;
			reasons.push(`IP country (${g.countryCode}) differs from expected (${g.expectedCountry})`);
		} else if (g.countryCode) {
			score += 5;
			positives.push(`IP country matches (${g.countryCode})`);
		}
		if (g.isHostingOrProxy) {
			score -= 14;
			reasons.push(`VPN / Datacenter ISP (${g.isp || g.ip})`);
		}
	}

	const trustScore = Math.max(0, Math.min(100, Math.round(score)));
	const riskLevel = trustScore >= 75 ? 'trusted' : trustScore >= 45 ? 'review' : 'high_risk';
	const summaryParts = [...reasons, ...positives];

	return {
		trustScore,
		riskLevel,
		summary: summaryParts.length > 0 ? summaryParts.slice(0, 4).join(' · ') : 'Verification completed.',
		reasons,
		positives,
	};
}

export async function verifyLeadMultiSignal({
	phoneNumber = '',
	email = '',
	countryCode = '',
	leadName = '',
	companyName = '',
	ip = '',
	domain = '',
	signalFlags = {},
	apiKey,
	lookupBaseUrl = DEFAULT_LOOKUP_BASE,
	ipapiBaseUrl = DEFAULT_IPAPI_BASE,
	fetchImpl = fetch,
}) {
	const cleanKey = String(apiKey || '').trim();
	if (!cleanKey) throw createServiceError('Missing ProWebLook API key.', 'invalid_credentials', 401);

	const cleanPhone = normalizePhoneInput(phoneNumber);
	const cleanEmail = normalizeEmailInput(email);
	const cleanIp = isValidIpAddress(ip) ? String(ip).trim() : '';
	const explicitDomain = normalizeDomainInput(domain);
	const emailDomain = extractEmailDomain(cleanEmail);
	const targetDomain =
		explicitDomain ||
		(emailDomain && !FREE_WEBMAIL_DOMAINS.has(emailDomain) && !isDisposableEmailDomain(cleanEmail)
			? emailDomain
			: '');

	if (!cleanPhone && !cleanEmail && !cleanIp && !targetDomain) {
		throw createServiceError(
			'Provide at least one of phoneNumber, email, domain, or ip to verify.',
			'invalid_input',
			400,
		);
	}

	const tasks = {};
	if (cleanPhone) {
		if (signalFlags.phone !== false) {
			tasks.phone = validatePhone({ phoneNumber: cleanPhone, countryCode, apiKey: cleanKey, lookupBaseUrl, fetchImpl });
		}
		if (signalFlags.whatsapp !== false) {
			tasks.whatsapp = checkWhatsapp({ phoneNumber: cleanPhone, countryCode, apiKey: cleanKey, lookupBaseUrl, fetchImpl });
		}
		if (signalFlags.callerId !== false) {
			tasks.callerId = lookupCallerId({
				phoneNumber: cleanPhone,
				countryCode,
				leadName,
				companyName,
				apiKey: cleanKey,
				lookupBaseUrl,
				fetchImpl,
			});
		}
	}

	if (cleanEmail && signalFlags.email !== false) {
		tasks.email = verifyEmail({ email: cleanEmail, apiKey: cleanKey, lookupBaseUrl, fetchImpl });
	}

	if (targetDomain && signalFlags.blacklist !== false) {
		tasks.blacklist = checkBlacklist({
			target: targetDomain,
			targetType: 'domain',
			skipFreeWebmail: true,
			apiKey: cleanKey,
			lookupBaseUrl,
			fetchImpl,
		});
	}

	if (cleanIp && signalFlags.ipGeo !== false) {
		tasks.ipGeo = lookupIpGeo({
			ip: cleanIp,
			expectedCountry: countryCode,
			apiKey: cleanKey,
			ipapiBaseUrl,
			fetchImpl,
		});
	}

	const keys = Object.keys(tasks);
	const settled = await Promise.allSettled(keys.map((k) => tasks[k]));
	const signals = {};
	const warnings = [];
	let fatalAuthError = null;

	keys.forEach((k, idx) => {
		const res = settled[idx];
		if (res.status === 'fulfilled') {
			signals[k] = res.value;
		} else {
			const err = res.reason;
			if (err?.code === 'invalid_credentials') {
				fatalAuthError = err;
			} else {
				warnings.push(`${k}: ${err?.message || 'unavailable'}`);
			}
		}
	});

	if (fatalAuthError && Object.keys(signals).length === 0) {
		throw fatalAuthError;
	}

	// If countryCode wasn't passed explicitly, compare IP Geo against Phone Country when both exist
	if (!countryCode && signals.ipGeo && !signals.ipGeo.skippedPrivateIp && signals.phone?.countryCode) {
		const pCc = signals.phone.countryCode;
		if (/^[A-Z]{2}$/.test(pCc) && signals.ipGeo.countryCode) {
			signals.ipGeo.expectedCountry = pCc;
			signals.ipGeo.countryMismatch = signals.ipGeo.countryCode !== pCc;
		}
	}

	const scoring = calculateLeadTrustScore({ signals });

	return {
		trustScore: scoring.trustScore,
		riskLevel: scoring.riskLevel,
		summary: scoring.summary,
		reasons: scoring.reasons,
		positives: scoring.positives,
		warnings,
		verifiedAt: new Date().toISOString(),
		signals: {
			phone: signals.phone || null,
			whatsapp: signals.whatsapp || null,
			email: signals.email || null,
			callerId: signals.callerId || null,
			blacklist: signals.blacklist || null,
			ipGeo: signals.ipGeo || null,
		},
	};
}

export async function getAccountBalances({
	apiKey,
	monolithBaseUrl = DEFAULT_MONOLITH_BASE,
	fetchImpl = fetch,
}) {
	const cleanKey = String(apiKey || '').trim();
	if (!cleanKey) throw createServiceError('Missing ProWebLook API key.', 'invalid_credentials', 401);

	const { status, ok, data } = await fetchJsonWithTimeout(
		`${monolithBaseUrl.replace(/\/+$/, '')}/zapier/get_balance`,
		{
			method: 'POST',
			headers: {
				Authorization: `Bearer ${cleanKey}`,
				'Content-Type': 'application/json',
				Accept: 'application/json',
			},
			body: '{}',
		},
		fetchImpl,
		12000,
	);

	if (status === 401 || status === 403) {
		throw createServiceError('Invalid ProWebLook API key.', 'invalid_credentials', 401);
	}

	if (ok && data?.ok && data?.data) {
		return {
			accountEmail: String(data.data.account_email || ''),
			phoneCredits: Number(data.data.phone?.amount ?? 0),
			whatsappCredits: Number(data.data.whatsapp?.amount ?? 0),
			emailCredits: Number(data.data.email?.amount ?? 0),
			callerIdCredits: Number(data.data.caller?.amount ?? 0),
			blacklistCredits: Number(data.data.blacklist?.amount ?? 0),
			ipGeoMonthlyFreeQuota: 20000,
		};
	}

	const fallback = await fetchJsonWithTimeout(
		`${monolithBaseUrl.replace(/\/+$/, '')}/checkbalance?api_key=${encodeURIComponent(cleanKey)}&service=phonevalidation`,
		{ method: 'GET', headers: { Accept: 'application/json' } },
		fetchImpl,
		10000,
	);
	if (fallback.status === 401 || fallback.data?.status === false) {
		throw createServiceError('Invalid ProWebLook API key.', 'invalid_credentials', 401);
	}

	return {
		accountEmail: '',
		phoneCredits: Number(fallback.data?.balance ?? 0),
		whatsappCredits: 0,
		emailCredits: 0,
		callerIdCredits: 0,
		blacklistCredits: 0,
		ipGeoMonthlyFreeQuota: 20000,
	};
}
