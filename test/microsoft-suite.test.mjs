import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { handleMicrosoftEdgeRequest } from '../src/worker.js';
import { compareCallerIdWithLead, normalizePhoneInput } from '../src/shared/parser.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

function createMockKv() {
	const store = new Map();
	return {
		async get(k) {
			return store.has(k) ? store.get(k) : null;
		},
		async put(k, v) {
			store.set(k, String(v));
		},
	};
}

function createMockUpstreamFetch() {
	return async (urlStr, init = {}) => {
		const u = new URL(urlStr);
		const body = init.body ? JSON.parse(init.body) : {};
		const apiKey =
			body.api_key ||
			u.searchParams.get('api_key') ||
			(init.headers?.Authorization || '').replace(/^Bearer\s+/i, '');

		if (apiKey === 'bad_key') {
			return new Response(JSON.stringify({ error: 'Unauthorized' }), { status: 401 });
		}
		if (apiKey === 'no_credit_key') {
			return new Response(JSON.stringify({ error: 'Insufficient balance' }), { status: 402 });
		}

		if (u.pathname.endsWith('/phonevalidation')) {
			return new Response(
				JSON.stringify({
					status: 'VALID_CONFIRMED',
					formatted_number: '+1 415-555-9999',
					formatnational: '(415) 555-9999',
					linetype: 'Mobile',
					carrier: 'T-Mobile USA',
					location: 'California',
					countrycode: 'US',
					countryname: 'United States',
					timezone: ['America/Los_Angeles'],
				}),
				{ status: 200 }
			);
		}

		if (u.pathname.endsWith('/wavalidation')) {
			return new Response(
				JSON.stringify({
					numberstatus: true,
					businessnumber: true,
					countrycode: 'US',
				}),
				{ status: 200 }
			);
		}

		if (u.pathname.endsWith('/emailvalidation')) {
			return new Response(
				JSON.stringify({
					verdict: 'deliverable',
					emailexist: true,
					emailformat: true,
					reason_code: 'mailbox_ok',
				}),
				{ status: 200 }
			);
		}

		if (u.pathname.endsWith('/calleridvalidation')) {
			return new Response(
				JSON.stringify({
					caller_name: 'ACME CLOUD INC',
					caller_type: 'BUSINESS',
					carrier: 'T-Mobile USA',
					linetype: 'Mobile',
					location: 'San Jose, CA',
					countrycode: 'US',
				}),
				{ status: 200 }
			);
		}

		if (u.pathname.endsWith('/blacklistchecker')) {
			return new Response(
				JSON.stringify({
					listed: false,
					listed_count: 0,
					total_checked: 52,
				}),
				{ status: 200 }
			);
		}

		if (u.hostname === 'ipapi.proweblook.com') {
			return new Response(
				JSON.stringify({
					ip: u.searchParams.get('ip') || '1.1.1.1',
					country_code: 'US',
					country_name: 'United States',
					region: 'California',
					city: 'San Jose',
					isp: 'Comcast Cable',
					asn: 'AS7922',
					timezone: 'America/Los_Angeles',
				}),
				{ status: 200 }
			);
		}

		if (u.pathname.endsWith('/zapier/get_balance')) {
			return new Response(
				JSON.stringify({
					ok: true,
					data: {
						account_email: 'ops@acme.io',
						phone: { amount: 1250 },
						whatsapp: { amount: 900 },
						email: { amount: 3400 },
						caller: { amount: 420 },
						blacklist: { amount: 800 },
					},
				}),
				{ status: 200 }
			);
		}

		return new Response(JSON.stringify({ error: 'Not mocked: ' + urlStr }), { status: 404 });
	};
}

test('Parser & Caller ID fuzzy matching', () => {
	assert.equal(normalizePhoneInput('+1 (415) 555-2671'), '+14155552671');
	assert.equal(normalizePhoneInput('0091-98765-43210'), '+919876543210');
	assert.equal(compareCallerIdWithLead('Jenkins Sarah', 'Sarah Jenkins'), 'match');
	assert.equal(compareCallerIdWithLead('Burner Relay Pool', 'Fake Buyer'), 'mismatch');
});

test('Static & Legal routes serve valid XML, JSON, JS, HTML, CSS, and PNG assets', async () => {
	const paths = [
		['/health', 'application/json'],
		['/', 'text/html'],
		['/support', 'text/html'],
		['/privacy', 'text/html'],
		['/terms', 'text/html'],
		['/manifest.xml', 'application/xml'],
		['/functions.json', 'application/json'],
		['/functions.js', 'application/javascript'],
		['/taskpane.html', 'text/html'],
		['/taskpane.css', 'text/css'],
		['/taskpane.js', 'application/javascript'],
		['/swagger.json', 'application/json'],
		['/assets/icon-16.png', 'image/png'],
		['/assets/icon-32.png', 'image/png'],
		['/assets/icon-64.png', 'image/png'],
		['/assets/icon-80.png', 'image/png'],
		['/assets/icon-128.png', 'image/png'],
		['/assets/store-logo-300.png', 'image/png'],
	];

	for (const [p, expectedMime] of paths) {
		const res = await handleMicrosoftEdgeRequest(new Request(`https://excel.proweblook.com${p}`));
		assert.equal(res.status, 200, `Expected 200 on GET ${p}`);
		assert.ok(
			(res.headers.get('Content-Type') || '').includes(expectedMime),
			`Unexpected Content-Type on ${p}: ${res.headers.get('Content-Type')}`
		);
	}

	// Verify taskpane.html permits Office Online iframe embedding
	const tp = await handleMicrosoftEdgeRequest(new Request('https://excel.proweblook.com/taskpane.html'));
	assert.equal(tp.headers.get('Content-Security-Policy'), 'frame-ancestors *;');

	// Verify CORS preflight
	const opt = await handleMicrosoftEdgeRequest(
		new Request('https://m365.proweblook.com/api/connector/v1/verify-lead', { method: 'OPTIONS' })
	);
	assert.equal(opt.status, 204);
});

test('Excel Zero-Signup Sample Sheet rows work immediately with zero API key or KV secret', async () => {
	const trustedRes = await handleMicrosoftEdgeRequest(
		new Request('https://excel.proweblook.com/api/excel/v1/verify', {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				phoneNumber: '+919876543210',
				email: 'aarav@infosys.com',
				leadName: 'Aarav Mehta',
				countryCode: 'IN',
				ip: '49.36.128.10',
			}),
		})
	);
	assert.equal(trustedRes.status, 200);
	const trustedJson = await trustedRes.json();
	assert.equal(trustedJson.ok, true);
	assert.equal(trustedJson.mode, 'sample');
	assert.equal(trustedJson.data.riskLevel, 'trusted');
	assert.ok(trustedJson.data.trustScore >= 85);

	const fraudRes = await handleMicrosoftEdgeRequest(
		new Request('https://excel.proweblook.com/api/excel/v1/verify', {
			method: 'POST',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({
				phoneNumber: '+12065550199',
				email: 'temp99@mailinator.com',
				leadName: 'Fake Buyer',
				countryCode: 'US',
				ip: '185.220.101.4',
			}),
		})
	);
	const fraudJson = await fraudRes.json();
	assert.equal(fraudJson.data.riskLevel, 'high_risk');
	assert.ok(fraudJson.data.trustScore < 25);
});

test('Excel Starter Sandbox meters KV quota and enforces limit on non-sample rows', async () => {
	const env = {
		SANDBOX_LIMIT: '2',
		PROWEBLOOK_STARTER_API_KEY: 'starter_valid_key',
		PROWEBLOOK_EXCEL_KV: createMockKv(),
	};
	const mockFetch = createMockUpstreamFetch();

	for (let i = 1; i <= 2; i++) {
		const r = await handleMicrosoftEdgeRequest(
			new Request('https://excel.proweblook.com/api/excel/v1/single', {
				method: 'POST',
				headers: { 'Content-Type': 'application/json', 'X-PWL-Install-Id': 'inst-1' },
				body: JSON.stringify({ signal: 'phone', phoneNumber: '+14155559999' }),
			}),
			env,
			mockFetch
		);
		assert.equal(r.status, 200);
		const j = await r.json();
		assert.equal(j.mode, 'sandbox');
		assert.equal(j.sandbox.used, i);
	}

	const exhausted = await handleMicrosoftEdgeRequest(
		new Request('https://excel.proweblook.com/api/excel/v1/single', {
			method: 'POST',
			headers: { 'Content-Type': 'application/json', 'X-PWL-Install-Id': 'inst-1' },
			body: JSON.stringify({ signal: 'phone', phoneNumber: '+14155559999' }),
		}),
		env,
		mockFetch
	);
	assert.equal(exhausted.status, 429);
	const exJson = await exhausted.json();
	assert.equal(exJson.code, 'sandbox_exhausted');
});

test('Power Platform Custom Connector Gateway: all 8 operations + auth/error contracts + Swagger parity', async () => {
	const swagger = JSON.parse(
		fs.readFileSync(
			path.join(ROOT, 'power-platform-connector', 'certified-connectors', 'ProWebLook', 'apiDefinition.swagger.json'),
			'utf8'
		)
	);
	const mockFetch = createMockUpstreamFetch();

	// 1. Missing X-API-Key -> 401
	const unauth = await handleMicrosoftEdgeRequest(
		new Request('https://m365.proweblook.com/api/connector/v1/balances'),
		{},
		mockFetch
	);
	assert.equal(unauth.status, 401);

	// 2. Insufficient credits -> 402
	const lowCred = await handleMicrosoftEdgeRequest(
		new Request('https://m365.proweblook.com/api/connector/v1/validate-phone', {
			method: 'POST',
			headers: { 'Content-Type': 'application/json', 'X-API-Key': 'no_credit_key' },
			body: JSON.stringify({ phoneNumber: '+14155559999' }),
		}),
		{},
		mockFetch
	);
	assert.equal(lowCred.status, 402);

	// 3. Exercise all 8 operations and verify returned keys exist in Swagger 200 schema
	const cases = [
		['POST', '/verify-lead', { phoneNumber: '+14155559999', email: 'ops@acme.io', companyName: 'Acme Cloud', ip: '1.1.1.1' }, 'VerifyLeadResponse'],
		['POST', '/validate-phone', { phoneNumber: '+14155559999', countryCode: 'US' }, 'PhoneValidationResult'],
		['POST', '/check-whatsapp', { phoneNumber: '+14155559999' }, 'WhatsAppCheckResult'],
		['POST', '/verify-email', { email: 'ops@acme.io' }, 'EmailVerificationResult'],
		['POST', '/lookup-caller-id', { phoneNumber: '+14155559999', companyName: 'Acme Cloud' }, 'CallerIdResult'],
		['POST', '/check-blacklist', { target: 'acme.io' }, 'BlacklistCheckResult'],
		['POST', '/lookup-ip-geo', { ip: '1.1.1.1', expectedCountry: 'US' }, 'IpGeoResult'],
		['GET', '/balances', null, 'AccountBalancesResponse'],
	];

	for (const [method, subRoute, payload, defName] of cases) {
		const res = await handleMicrosoftEdgeRequest(
			new Request(`https://m365.proweblook.com/api/connector/v1${subRoute}`, {
				method,
				headers: {
					'Content-Type': 'application/json',
					'X-API-Key': 'valid_user_key',
				},
				body: payload ? JSON.stringify(payload) : undefined,
			}),
			{},
			mockFetch
		);
		assert.equal(res.status, 200, `Expected 200 on ${method} ${subRoute}`);
		const json = await res.json();
		const schemaProps = Object.keys(swagger.definitions[defName].properties);
		for (const k of Object.keys(json)) {
			assert.ok(schemaProps.includes(k), `Unexpected key "${k}" in ${subRoute} not declared in ${defName}`);
		}
		for (const expectedKey of schemaProps) {
			assert.ok(expectedKey in json, `Missing declared key "${expectedKey}" in ${subRoute} (${defName})`);
		}
	}
});
