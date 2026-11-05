/**
 * ProWebLook Microsoft 365 & Power Platform Edge Worker (`proweblook-microsoft-edge`)
 * Custom Domains:
 *  - https://excel.proweblook.com  (Microsoft Excel Add-in + Custom Functions + Bulk Taskpane)
 *  - https://m365.proweblook.com   (Power Automate, Azure Logic Apps & Copilot Studio Gateway)
 */

import { EMBEDDED_PNG_BASE64 } from './generated-icons.js';
import { EXCEL_STATIC_FILES } from './generated-static.js';
import { SWAGGER_JSON_TEXT } from './generated-swagger.js';
import { renderHomePage, renderPrivacyPage, renderSupportPage, renderTermsPage } from './pages.js';
import {
	DEFAULT_IPAPI_BASE,
	DEFAULT_LOOKUP_BASE,
	DEFAULT_MONOLITH_BASE,
	checkBlacklist,
	checkWhatsapp,
	createServiceError,
	getAccountBalances,
	lookupCallerId,
	lookupIpGeo,
	validatePhone,
	verifyEmail,
	verifyLeadMultiSignal,
} from './shared/engine.js';
import {
	resolveReviewerBalances,
	resolveReviewerMultiSignal,
	resolveReviewerSingle,
	tryResolveSampleMultiSignal,
	tryResolveSampleSingle,
} from './shared/samples.js';

const CORS_HEADERS = {
	'Access-Control-Allow-Origin': '*',
	'Access-Control-Allow-Methods': 'GET, HEAD, POST, OPTIONS',
	'Access-Control-Allow-Headers': 'Content-Type, Accept, Authorization, X-API-Key, X-PWL-Install-Id',
};

function jsonResponse(payload, status = 200) {
	return new Response(JSON.stringify(payload), {
		status,
		headers: {
			'Content-Type': 'application/json; charset=utf-8',
			'Cache-Control': 'no-store',
			...CORS_HEADERS,
		},
	});
}

function textAssetResponse(content, contentType, method = 'GET', extraHeaders = {}) {
	return new Response(method === 'HEAD' ? null : content, {
		status: 200,
		headers: {
			'Content-Type': contentType,
			'Cache-Control': 'public, max-age=300',
			...CORS_HEADERS,
			...extraHeaders,
		},
	});
}

function base64ToUint8Array(b64) {
	const bin = atob(b64);
	const bytes = new Uint8Array(bin.length);
	for (let i = 0; i < bin.length; i++) {
		bytes[i] = bin.charCodeAt(i);
	}
	return bytes;
}

function extractApiKeyFromHeaders(request) {
	const direct = String(request.headers.get('X-API-Key') || request.headers.get('x-api-key') || '').trim();
	if (direct) return direct;
	const auth = String(request.headers.get('Authorization') || '').trim();
	if (/^Bearer\s+/i.test(auth)) {
		return auth.replace(/^Bearer\s+/i, '').trim();
	}
	return '';
}

function isReviewerToken(candidateKey, env = {}) {
	const configured = String(env.CERT_REVIEWER_TOKEN || '').trim();
	const provided = String(candidateKey || '').trim();
	if (!configured || configured.length < 16 || provided.length !== configured.length) {
		return false;
	}
	let diff = 0;
	for (let i = 0; i < configured.length; i++) {
		diff |= configured.charCodeAt(i) ^ provided.charCodeAt(i);
	}
	return diff === 0;
}

async function sha256Short(text) {
	const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(String(text || '')));
	return Array.from(new Uint8Array(buf), (b) => b.toString(16).padStart(2, '0'))
		.join('')
		.slice(0, 24);
}

async function readSandboxQuota(kv, request, limit = 20) {
	if (!kv) return { used: 0, limit, remaining: limit, ipKey: '', idKey: '' };
	const installId = String(request.headers.get('X-PWL-Install-Id') || '').trim().slice(0, 64);
	const clientIp = String(request.headers.get('CF-Connecting-IP') || '0.0.0.0').trim();
	const ipHash = await sha256Short(`pwl_excel_ip:${clientIp}`);
	const idHash = installId ? await sha256Short(`pwl_excel_id:${installId}`) : ipHash;
	const ipKey = `sb_ip:${ipHash}`;
	const idKey = `sb_id:${idHash}`;
	const [ipUsedRaw, idUsedRaw] = await Promise.all([kv.get(ipKey), kv.get(idKey)]);
	const used = Math.max(Number(ipUsedRaw || 0), Number(idUsedRaw || 0));
	return {
		used,
		limit,
		remaining: Math.max(0, limit - used),
		ipKey,
		idKey,
	};
}

async function incrementSandboxQuota(kv, quotaInfo) {
	const nextUsed = quotaInfo.used + 1;
	if (kv && quotaInfo.ipKey && quotaInfo.idKey) {
		const ttl = 30 * 24 * 3600;
		await Promise.all([
			kv.put(quotaInfo.ipKey, String(nextUsed), { expirationTtl: ttl }),
			kv.put(quotaInfo.idKey, String(nextUsed), { expirationTtl: ttl }),
		]);
	}
	return {
		used: nextUsed,
		limit: quotaInfo.limit,
		remaining: Math.max(0, quotaInfo.limit - nextUsed),
	};
}

async function executeSingleSignal(signal, body, apiKey, env, fetchImpl) {
	if (isReviewerToken(apiKey, env)) {
		return resolveReviewerSingle(signal, body);
	}
	const lookupBaseUrl = env.LOOKUP_BASE_URL || DEFAULT_LOOKUP_BASE;
	const ipapiBaseUrl = env.IPAPI_BASE_URL || DEFAULT_IPAPI_BASE;

	switch (signal) {
		case 'phone':
			return validatePhone({
				phoneNumber: body.phoneNumber,
				countryCode: body.countryCode || '',
				apiKey,
				lookupBaseUrl,
				fetchImpl,
			});
		case 'whatsapp':
			return checkWhatsapp({
				phoneNumber: body.phoneNumber,
				countryCode: body.countryCode || '',
				apiKey,
				lookupBaseUrl,
				fetchImpl,
			});
		case 'email':
			return verifyEmail({
				email: body.email,
				apiKey,
				lookupBaseUrl,
				fetchImpl,
			});
		case 'callerId':
			return lookupCallerId({
				phoneNumber: body.phoneNumber,
				countryCode: body.countryCode || '',
				leadName: body.leadName || '',
				companyName: body.companyName || '',
				apiKey,
				lookupBaseUrl,
				fetchImpl,
			});
		case 'blacklist':
			return checkBlacklist({
				target: body.target,
				targetType: body.targetType || 'auto',
				skipFreeWebmail: body.skipFreeWebmail !== false,
				apiKey,
				lookupBaseUrl,
				fetchImpl,
			});
		case 'ipGeo':
			return lookupIpGeo({
				ip: body.ip,
				expectedCountry: body.expectedCountry || body.countryCode || '',
				apiKey,
				ipapiBaseUrl,
				fetchImpl,
			});
		default:
			throw createServiceError(`Unsupported signal "${signal}".`, 'invalid_input', 400);
	}
}

function errorJsonResponse(err) {
	const status = Number(err?.httpStatus) || 500;
	const code = String(err?.code || 'provider_error');
	const message = String(err?.message || 'Verification request failed.');
	return jsonResponse({ ok: false, code, error: message }, status);
}

export async function handleMicrosoftEdgeRequest(request, env = {}, fetchImpl = fetch) {
	const url = new URL(request.url);
	const pathname = url.pathname.replace(/\/+$/, '') || '/';
	const method = request.method.toUpperCase();
	const isGetOrHead = method === 'GET' || method === 'HEAD';

	if (method === 'OPTIONS') {
		return new Response(null, { status: 204, headers: CORS_HEADERS });
	}

	if (isGetOrHead && pathname === '/health') {
		return jsonResponse({
			ok: true,
			service: 'proweblook-microsoft-edge',
			version: '1.0.0',
			host: url.hostname,
			timestamp: new Date().toISOString(),
		});
	}

	// Public HTML & Legal Pages
	if (isGetOrHead && pathname === '/') {
		return textAssetResponse(renderHomePage(), 'text/html; charset=utf-8', method);
	}
	if (isGetOrHead && (pathname === '/support' || pathname === '/setup')) {
		return textAssetResponse(renderSupportPage(), 'text/html; charset=utf-8', method);
	}
	if (isGetOrHead && pathname === '/privacy') {
		return textAssetResponse(renderPrivacyPage(), 'text/html; charset=utf-8', method);
	}
	if (isGetOrHead && pathname === '/terms') {
		return textAssetResponse(renderTermsPage(), 'text/html; charset=utf-8', method);
	}

	// Swagger 2.0 OpenAPI Spec
	if (isGetOrHead && (pathname === '/swagger.json' || pathname === '/apiDefinition.swagger.json')) {
		return textAssetResponse(SWAGGER_JSON_TEXT, 'application/json; charset=utf-8', method);
	}

	// Excel Add-in Static Assets
	if (isGetOrHead && pathname === '/manifest.xml') {
		return textAssetResponse(EXCEL_STATIC_FILES['manifest.xml'], 'application/xml; charset=utf-8', method, {
			'Content-Disposition': 'inline; filename="proweblook-excel-manifest.xml"',
		});
	}
	if (isGetOrHead && pathname === '/functions.json') {
		return textAssetResponse(EXCEL_STATIC_FILES['functions.json'], 'application/json; charset=utf-8', method);
	}
	if (isGetOrHead && pathname === '/functions.js') {
		return textAssetResponse(EXCEL_STATIC_FILES['functions.js'], 'application/javascript; charset=utf-8', method);
	}
	if (isGetOrHead && pathname === '/taskpane.html') {
		return textAssetResponse(EXCEL_STATIC_FILES['taskpane.html'], 'text/html; charset=utf-8', method, {
			'Content-Security-Policy': 'frame-ancestors *;',
		});
	}
	if (isGetOrHead && pathname === '/taskpane.css') {
		return textAssetResponse(EXCEL_STATIC_FILES['taskpane.css'], 'text/css; charset=utf-8', method);
	}
	if (isGetOrHead && pathname === '/taskpane.js') {
		return textAssetResponse(EXCEL_STATIC_FILES['taskpane.js'], 'application/javascript; charset=utf-8', method);
	}

	if (isGetOrHead && pathname.startsWith('/assets/')) {
		const assetName = pathname.slice('/assets/'.length);
		const b64 = EMBEDDED_PNG_BASE64[assetName];
		if (b64) {
			const bytes = base64ToUint8Array(b64);
			return new Response(method === 'HEAD' ? null : bytes, {
				status: 200,
				headers: {
					'Content-Type': 'image/png',
					'Cache-Control': 'public, max-age=86400',
					...CORS_HEADERS,
				},
			});
		}
	}

	// =========================================================================
	// PART 1: EXCEL ADD-IN RUNTIME API (`/api/excel/v1/*`)
	// =========================================================================
	const sandboxLimit = Math.max(1, Number(env.SANDBOX_LIMIT || 20));
	const starterKey = String(env.PROWEBLOOK_STARTER_API_KEY || '').trim();

	if (method === 'GET' && pathname === '/api/excel/v1/status') {
		const quota = await readSandboxQuota(env.PROWEBLOOK_EXCEL_KV, request, sandboxLimit);
		return jsonResponse({
			ok: true,
			sandbox: {
				liveEnabled: Boolean(starterKey),
				used: quota.used,
				limit: quota.limit,
				remaining: quota.remaining,
			},
		});
	}

	if (method === 'GET' && pathname === '/api/excel/v1/balances') {
		const userKey = extractApiKeyFromHeaders(request);
		try {
			if (isReviewerToken(userKey, env)) {
				return jsonResponse({ ok: true, data: resolveReviewerBalances() });
			}
			const balances = await getAccountBalances({
				apiKey: userKey,
				monolithBaseUrl: env.MONOLITH_BASE_URL || DEFAULT_MONOLITH_BASE,
				fetchImpl,
			});
			return jsonResponse({ ok: true, data: balances });
		} catch (err) {
			return errorJsonResponse(err);
		}
	}

	if (method === 'POST' && pathname === '/api/excel/v1/single') {
		const body = await request.json().catch(() => null);
		if (!body || typeof body !== 'object') {
			return jsonResponse({ ok: false, code: 'invalid_input', error: 'Invalid JSON request body.' }, 400);
		}

		const signal = String(body.signal || '').trim();
		const userKey = extractApiKeyFromHeaders(request);

		try {
			if (userKey) {
				const data = await executeSingleSignal(signal, body, userKey, env, fetchImpl);
				return jsonResponse({ ok: true, mode: 'byok', data });
			}

			const sampleHit = tryResolveSampleSingle(signal, body);
			if (sampleHit) {
				return jsonResponse({ ok: true, mode: 'sample', data: sampleHit });
			}

			if (!starterKey) {
				return jsonResponse(
					{
						ok: false,
						code: 'starter_key_required',
						error:
							'Connect your free ProWebLook API key in the API Key tab (50 free credits per product at proweblook.com/auth/signup) or test with Sample Sheet rows.',
					},
					401
				);
			}

			const quota = await readSandboxQuota(env.PROWEBLOOK_EXCEL_KV, request, sandboxLimit);
			if (quota.remaining <= 0) {
				return jsonResponse(
					{
						ok: false,
						code: 'sandbox_exhausted',
						error: `You have used all ${sandboxLimit} free Starter Sandbox lookups. Get 50 free credits per product at proweblook.com/auth/signup and paste your API key in the API Key tab.`,
					},
					429
				);
			}

			const data = await executeSingleSignal(signal, body, starterKey, env, fetchImpl);
			const nextQuota = await incrementSandboxQuota(env.PROWEBLOOK_EXCEL_KV, quota);
			return jsonResponse({ ok: true, mode: 'sandbox', sandbox: nextQuota, data });
		} catch (err) {
			return errorJsonResponse(err);
		}
	}

	if (method === 'POST' && pathname === '/api/excel/v1/verify') {
		const body = await request.json().catch(() => null);
		if (!body || typeof body !== 'object') {
			return jsonResponse({ ok: false, code: 'invalid_input', error: 'Invalid JSON request body.' }, 400);
		}

		const userKey = extractApiKeyFromHeaders(request);
		const lookupBaseUrl = env.LOOKUP_BASE_URL || DEFAULT_LOOKUP_BASE;
		const ipapiBaseUrl = env.IPAPI_BASE_URL || DEFAULT_IPAPI_BASE;

		try {
			if (userKey) {
				if (isReviewerToken(userKey, env)) {
					return jsonResponse({ ok: true, mode: 'byok', data: resolveReviewerMultiSignal(body) });
				}
				const data = await verifyLeadMultiSignal({
					...body,
					apiKey: userKey,
					lookupBaseUrl,
					ipapiBaseUrl,
					fetchImpl,
				});
				return jsonResponse({ ok: true, mode: 'byok', data });
			}

			const sampleMulti = tryResolveSampleMultiSignal(body);
			if (sampleMulti) {
				return jsonResponse({ ok: true, mode: 'sample', data: sampleMulti });
			}

			if (!starterKey) {
				return jsonResponse(
					{
						ok: false,
						code: 'starter_key_required',
						error:
							'Connect your free ProWebLook API key in the API Key tab (50 free credits per product at proweblook.com/auth/signup) or click Insert Sample Sheet to test right away.',
					},
					401
				);
			}

			const quota = await readSandboxQuota(env.PROWEBLOOK_EXCEL_KV, request, sandboxLimit);
			if (quota.remaining <= 0) {
				return jsonResponse(
					{
						ok: false,
						code: 'sandbox_exhausted',
						error: `You have used all ${sandboxLimit} free Starter Sandbox lookups. Get 50 free credits per product at proweblook.com/auth/signup and paste your API key in the API Key tab.`,
					},
					429
				);
			}

			const data = await verifyLeadMultiSignal({
				...body,
				apiKey: starterKey,
				lookupBaseUrl,
				ipapiBaseUrl,
				fetchImpl,
			});
			const nextQuota = await incrementSandboxQuota(env.PROWEBLOOK_EXCEL_KV, quota);
			return jsonResponse({ ok: true, mode: 'sandbox', sandbox: nextQuota, data });
		} catch (err) {
			return errorJsonResponse(err);
		}
	}

	// =========================================================================
	// PART 2: POWER AUTOMATE / LOGIC APPS / COPILOT STUDIO CONNECTOR GATEWAY
	// Base Path: `/api/connector/v1` — returns flat schema-matching JSON objects
	// =========================================================================
	if (pathname.startsWith('/api/connector/v1/')) {
		const apiKey = extractApiKeyFromHeaders(request);
		if (!apiKey) {
			return jsonResponse(
				{
					ok: false,
					code: 'invalid_credentials',
					error: 'Missing ProWebLook API key. Provide your key in the X-API-Key header.',
				},
				401
			);
		}

		const subRoute = pathname.slice('/api/connector/v1'.length);
		const lookupBaseUrl = env.LOOKUP_BASE_URL || DEFAULT_LOOKUP_BASE;
		const ipapiBaseUrl = env.IPAPI_BASE_URL || DEFAULT_IPAPI_BASE;
		const monolithBaseUrl = env.MONOLITH_BASE_URL || DEFAULT_MONOLITH_BASE;

		try {
			if (method === 'GET' && subRoute === '/balances') {
				if (isReviewerToken(apiKey, env)) {
					return jsonResponse(resolveReviewerBalances());
				}
				const balances = await getAccountBalances({ apiKey, monolithBaseUrl, fetchImpl });
				return jsonResponse(balances);
			}

			if (method !== 'POST') {
				return jsonResponse({ ok: false, code: 'method_not_allowed', error: 'Use HTTP POST for this operation.' }, 405);
			}

			const body = await request.json().catch(() => null);
			if (!body || typeof body !== 'object') {
				return jsonResponse({ ok: false, code: 'invalid_input', error: 'Request body must be valid JSON.' }, 400);
			}

			if (subRoute === '/verify-lead') {
				if (isReviewerToken(apiKey, env)) {
					const { warnings, sampleDemo, ...reviewerPayload } = resolveReviewerMultiSignal(body);
					return jsonResponse(reviewerPayload);
				}
				const res = await verifyLeadMultiSignal({
					phoneNumber: body.phoneNumber || '',
					email: body.email || '',
					countryCode: body.countryCode || '',
					leadName: body.leadName || '',
					companyName: body.companyName || '',
					ip: body.ip || '',
					domain: body.domain || '',
					apiKey,
					lookupBaseUrl,
					ipapiBaseUrl,
					fetchImpl,
				});
				const { warnings, ...connectorPayload } = res;
				return jsonResponse(connectorPayload);
			}

			if (subRoute === '/validate-phone') {
				return jsonResponse(await executeSingleSignal('phone', body, apiKey, env, fetchImpl));
			}
			if (subRoute === '/check-whatsapp') {
				return jsonResponse(await executeSingleSignal('whatsapp', body, apiKey, env, fetchImpl));
			}
			if (subRoute === '/verify-email') {
				return jsonResponse(await executeSingleSignal('email', body, apiKey, env, fetchImpl));
			}
			if (subRoute === '/lookup-caller-id') {
				return jsonResponse(await executeSingleSignal('callerId', body, apiKey, env, fetchImpl));
			}
			if (subRoute === '/check-blacklist') {
				return jsonResponse(await executeSingleSignal('blacklist', body, apiKey, env, fetchImpl));
			}
			if (subRoute === '/lookup-ip-geo') {
				return jsonResponse(await executeSingleSignal('ipGeo', body, apiKey, env, fetchImpl));
			}

			return jsonResponse({ ok: false, code: 'not_found', error: `Unknown connector endpoint ${subRoute}.` }, 404);
		} catch (err) {
			return errorJsonResponse(err);
		}
	}

	return jsonResponse({ ok: false, code: 'not_found', error: 'Route not found.' }, 404);
}

export default {
	async fetch(request, env) {
		return handleMicrosoftEdgeRequest(request, env, fetch);
	},
};
