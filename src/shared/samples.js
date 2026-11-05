/**
 * Deterministic Sample Sheet Dataset & Reviewer Demo Fixtures
 * Guarantees 100% zero-config testability for Microsoft AppSource certification reviewers
 * (Policy 1100.3) and instant 1-click onboarding in Excel without burning live API credits.
 */

import { calculateLeadTrustScore } from './engine.js';
import { compareCallerIdWithLead, normalizeDomainInput, normalizeEmailInput, normalizePhoneInput } from './parser.js';

export const SAMPLE_SHEET_HEADERS = ['Name', 'Phone', 'Email', 'Country', 'IP Address'];

export const SAMPLE_SHEET_ROWS = [
	['Aarav Mehta', '+919876543210', 'aarav@infosys.com', 'IN', '49.36.128.10'],
	['Sarah Jenkins', '+14155552671', 'sarah.jenkins@stripe.com', 'US', '8.8.8.8'],
	['Fake Buyer', '+12065550199', 'temp99@mailinator.com', 'US', '185.220.101.4'],
	["Liam O'Connor", '+447700900123', 'liam.oconnor@gmail.com', 'GB', '103.21.244.8'],
	['Spam Bot', '+19990000000', 'promo@bad-mailer-spam.biz', 'US', '45.155.205.9'],
	['Fatima Al Mansoori', '+971501234567', 'fatima@emaar.ae', 'AE', '94.200.200.10'],
];

const DEMO_PHONES = {
	'+919876543210': {
		valid: true,
		status: 'VALID_CONFIRMED',
		verdict: 'valid',
		phoneNumber: '+919876543210',
		formattedNumber: '+91 98765 43210',
		nationalFormat: '098765 43210',
		lineType: 'Mobile',
		isVoip: false,
		isLandline: false,
		carrier: 'Reliance Jio',
		location: 'Mumbai, Maharashtra',
		countryCode: 'IN',
		countryName: 'India',
		timezone: 'Asia/Kolkata',
		whatsapp: { registered: true, isBusiness: false, accountType: 'personal', verdict: 'on_whatsapp', countryCode: 'IN' },
		callerId: { found: true, verdict: 'found', callerName: 'Aarav Mehta', callerType: 'CONSUMER', carrier: 'Reliance Jio', lineType: 'Mobile', location: 'Mumbai', countryCode: 'IN' },
	},
	'+14155552671': {
		valid: true,
		status: 'VALID_CONFIRMED',
		verdict: 'valid',
		phoneNumber: '+14155552671',
		formattedNumber: '+1 415-555-2671',
		nationalFormat: '(415) 555-2671',
		lineType: 'Mobile',
		isVoip: false,
		isLandline: false,
		carrier: 'AT&T Mobility',
		location: 'San Francisco, CA',
		countryCode: 'US',
		countryName: 'United States',
		timezone: 'America/Los_Angeles',
		whatsapp: { registered: true, isBusiness: true, accountType: 'business', verdict: 'on_whatsapp', countryCode: 'US' },
		callerId: { found: true, verdict: 'found', callerName: 'Jenkins Sarah', callerType: 'CONSUMER', carrier: 'AT&T Mobility', lineType: 'Mobile', location: 'San Francisco, CA', countryCode: 'US' },
	},
	'+12065550199': {
		valid: true,
		status: 'VALID_UNCONFIRMED',
		verdict: 'valid',
		phoneNumber: '+12065550199',
		formattedNumber: '+1 206-555-0199',
		nationalFormat: '(206) 555-0199',
		lineType: 'VoIP',
		isVoip: true,
		isLandline: false,
		carrier: 'Twilio Burner VoIP',
		location: 'Seattle, WA',
		countryCode: 'US',
		countryName: 'United States',
		timezone: 'America/Los_Angeles',
		whatsapp: { registered: false, isBusiness: false, accountType: 'none', verdict: 'not_on_whatsapp', countryCode: 'US' },
		callerId: { found: true, verdict: 'found', callerName: 'Burner Relay Pool', callerType: 'BUSINESS', carrier: 'Twilio', lineType: 'VoIP', location: 'Seattle, WA', countryCode: 'US' },
	},
	'+447700900123': {
		valid: true,
		status: 'VALID_CONFIRMED',
		verdict: 'valid',
		phoneNumber: '+447700900123',
		formattedNumber: '+44 7700 900123',
		nationalFormat: '07700 900123',
		lineType: 'Mobile',
		isVoip: false,
		isLandline: false,
		carrier: 'Vodafone UK',
		location: 'London',
		countryCode: 'GB',
		countryName: 'United Kingdom',
		timezone: 'Europe/London',
		whatsapp: { registered: true, isBusiness: false, accountType: 'personal', verdict: 'on_whatsapp', countryCode: 'GB' },
		callerId: { found: true, verdict: 'found', callerName: 'Liam OConnor', callerType: 'CONSUMER', carrier: 'Vodafone UK', lineType: 'Mobile', location: 'London', countryCode: 'GB' },
	},
	'+19990000000': {
		valid: false,
		status: 'INVALID',
		verdict: 'invalid',
		phoneNumber: '+19990000000',
		formattedNumber: '+19990000000',
		nationalFormat: '',
		lineType: 'Unknown',
		isVoip: false,
		isLandline: false,
		carrier: '',
		location: '',
		countryCode: 'US',
		countryName: 'United States',
		timezone: '',
		whatsapp: { registered: false, isBusiness: false, accountType: 'none', verdict: 'not_on_whatsapp', countryCode: 'US' },
		callerId: { found: false, verdict: 'not_found', callerName: '', callerType: '', carrier: '', lineType: '', location: '', countryCode: 'US' },
	},
	'+971501234567': {
		valid: true,
		status: 'VALID_CONFIRMED',
		verdict: 'valid',
		phoneNumber: '+971501234567',
		formattedNumber: '+971 50 123 4567',
		nationalFormat: '050 123 4567',
		lineType: 'Mobile',
		isVoip: false,
		isLandline: false,
		carrier: 'Etisalat (e&)',
		location: 'Dubai',
		countryCode: 'AE',
		countryName: 'United Arab Emirates',
		timezone: 'Asia/Dubai',
		whatsapp: { registered: true, isBusiness: true, accountType: 'business', verdict: 'on_whatsapp', countryCode: 'AE' },
		callerId: { found: true, verdict: 'found', callerName: 'Fatima Al Mansoori', callerType: 'BUSINESS', carrier: 'Etisalat', lineType: 'Mobile', location: 'Dubai', countryCode: 'AE' },
	},
};

const DEMO_EMAILS = {
	'aarav@infosys.com': {
		email: 'aarav@infosys.com',
		domain: 'infosys.com',
		deliverable: true,
		verdict: 'deliverable',
		isDisposable: false,
		isFreeWebmail: false,
		emailType: 'corporate',
		reasonCode: 'mailbox_verified',
		emailFormatValid: true,
		mailboxExists: true,
	},
	'sarah.jenkins@stripe.com': {
		email: 'sarah.jenkins@stripe.com',
		domain: 'stripe.com',
		deliverable: true,
		verdict: 'deliverable',
		isDisposable: false,
		isFreeWebmail: false,
		emailType: 'corporate',
		reasonCode: 'mailbox_verified',
		emailFormatValid: true,
		mailboxExists: true,
	},
	'temp99@mailinator.com': {
		email: 'temp99@mailinator.com',
		domain: 'mailinator.com',
		deliverable: false,
		verdict: 'disposable',
		isDisposable: true,
		isFreeWebmail: false,
		emailType: 'disposable',
		reasonCode: 'disposable_domain',
		emailFormatValid: true,
		mailboxExists: false,
	},
	'liam.oconnor@gmail.com': {
		email: 'liam.oconnor@gmail.com',
		domain: 'gmail.com',
		deliverable: true,
		verdict: 'deliverable',
		isDisposable: false,
		isFreeWebmail: true,
		emailType: 'free_webmail',
		reasonCode: 'mailbox_verified',
		emailFormatValid: true,
		mailboxExists: true,
	},
	'promo@bad-mailer-spam.biz': {
		email: 'promo@bad-mailer-spam.biz',
		domain: 'bad-mailer-spam.biz',
		deliverable: false,
		verdict: 'undeliverable',
		isDisposable: false,
		isFreeWebmail: false,
		emailType: 'corporate',
		reasonCode: 'smtp_rejected',
		emailFormatValid: true,
		mailboxExists: false,
	},
	'fatima@emaar.ae': {
		email: 'fatima@emaar.ae',
		domain: 'emaar.ae',
		deliverable: true,
		verdict: 'deliverable',
		isDisposable: false,
		isFreeWebmail: false,
		emailType: 'corporate',
		reasonCode: 'mailbox_verified',
		emailFormatValid: true,
		mailboxExists: true,
	},
};

const DEMO_DNSBL = {
	'infosys.com': { target: 'infosys.com', targetType: 'domain', listed: false, listedCount: 0, checkedCount: 54, skippedFreeWebmail: false, verdict: 'clean' },
	'stripe.com': { target: 'stripe.com', targetType: 'domain', listed: false, listedCount: 0, checkedCount: 54, skippedFreeWebmail: false, verdict: 'clean' },
	'emaar.ae': { target: 'emaar.ae', targetType: 'domain', listed: false, listedCount: 0, checkedCount: 54, skippedFreeWebmail: false, verdict: 'clean' },
	'gmail.com': { target: 'gmail.com', targetType: 'domain', listed: false, listedCount: 0, checkedCount: 0, skippedFreeWebmail: true, verdict: 'clean_free_webmail' },
	'bad-mailer-spam.biz': { target: 'bad-mailer-spam.biz', targetType: 'domain', listed: true, listedCount: 6, checkedCount: 54, skippedFreeWebmail: false, verdict: 'listed' },
};

const DEMO_IPS = {
	'49.36.128.10': {
		ip: '49.36.128.10',
		countryCode: 'IN',
		countryName: 'India',
		region: 'Maharashtra',
		city: 'Mumbai',
		isp: 'Reliance Jio Infocomm',
		asn: 'AS55836',
		timezone: 'Asia/Kolkata',
		isHostingOrProxy: false,
		skippedPrivateIp: false,
	},
	'8.8.8.8': {
		ip: '8.8.8.8',
		countryCode: 'US',
		countryName: 'United States',
		region: 'California',
		city: 'Mountain View',
		isp: 'Google Public DNS',
		asn: 'AS15169',
		timezone: 'America/Los_Angeles',
		isHostingOrProxy: false,
		skippedPrivateIp: false,
	},
	'185.220.101.4': {
		ip: '185.220.101.4',
		countryCode: 'DE',
		countryName: 'Germany',
		region: 'Berlin',
		city: 'Berlin',
		isp: 'Tor Exit / Hosting Relay GmbH',
		asn: 'AS205100',
		timezone: 'Europe/Berlin',
		isHostingOrProxy: true,
		skippedPrivateIp: false,
	},
	'103.21.244.8': {
		ip: '103.21.244.8',
		countryCode: 'SG',
		countryName: 'Singapore',
		region: 'Singapore',
		city: 'Singapore',
		isp: 'APNIC Transit Cloud',
		asn: 'AS132203',
		timezone: 'Asia/Singapore',
		isHostingOrProxy: true,
		skippedPrivateIp: false,
	},
	'45.155.205.9': {
		ip: '45.155.205.9',
		countryCode: 'NL',
		countryName: 'Netherlands',
		region: 'North Holland',
		city: 'Amsterdam',
		isp: 'Bulletproof Datacenter VPS',
		asn: 'AS49870',
		timezone: 'Europe/Amsterdam',
		isHostingOrProxy: true,
		skippedPrivateIp: false,
	},
	'94.200.200.10': {
		ip: '94.200.200.10',
		countryCode: 'AE',
		countryName: 'United Arab Emirates',
		region: 'Dubai',
		city: 'Dubai',
		isp: 'Emirates Integrated Telecommunications',
		asn: 'AS15802',
		timezone: 'Asia/Dubai',
		isHostingOrProxy: false,
		skippedPrivateIp: false,
	},
};

export function tryResolveSampleSingle(signal, params = {}) {
	if (signal === 'phone') {
		const p = DEMO_PHONES[normalizePhoneInput(params.phoneNumber)];
		if (!p) return null;
		const { whatsapp, callerId, ...phoneOnly } = p;
		return phoneOnly;
	}
	if (signal === 'whatsapp') {
		const p = DEMO_PHONES[normalizePhoneInput(params.phoneNumber)];
		if (!p) return null;
		return { phoneNumber: p.phoneNumber, ...p.whatsapp };
	}
	if (signal === 'callerId') {
		const p = DEMO_PHONES[normalizePhoneInput(params.phoneNumber)];
		if (!p) return null;
		const nameMatch = compareCallerIdWithLead(p.callerId.callerName, params.leadName || '', params.companyName || '');
		return { phoneNumber: p.phoneNumber, ...p.callerId, nameMatch };
	}
	if (signal === 'email') {
		const e = DEMO_EMAILS[normalizeEmailInput(params.email)];
		return e ? { ...e } : null;
	}
	if (signal === 'blacklist') {
		const dom = normalizeDomainInput(params.target);
		const d = DEMO_DNSBL[dom] || DEMO_DNSBL[String(params.target || '').trim().toLowerCase()];
		return d ? { ...d } : null;
	}
	if (signal === 'ipGeo') {
		const g = DEMO_IPS[String(params.ip || '').trim()];
		if (!g) return null;
		const expCc = String(params.expectedCountry || '').trim().toUpperCase().slice(0, 2);
		const countryMismatch = Boolean(g.countryCode && /^[A-Z]{2}$/.test(expCc) && g.countryCode !== expCc);
		return {
			...g,
			expectedCountry: expCc,
			countryMismatch,
			verdict: g.isHostingOrProxy ? 'hosting_or_vpn' : countryMismatch ? 'country_mismatch' : 'residential_isp',
		};
	}
	return null;
}

export function tryResolveSampleMultiSignal(params = {}) {
	const cleanPhone = normalizePhoneInput(params.phoneNumber);
	const cleanEmail = normalizeEmailInput(params.email);
	const cleanIp = String(params.ip || '').trim();
	const explicitDomain = normalizeDomainInput(params.domain);
	const emailDomain = cleanEmail.includes('@') ? cleanEmail.split('@')[1] : '';

	const hasAnyInput = Boolean(cleanPhone || cleanEmail || cleanIp || explicitDomain);
	if (!hasAnyInput) return null;

	// Only resolve deterministically when every provided non-empty identifier belongs to the sample sheet set
	if (cleanPhone && !DEMO_PHONES[cleanPhone]) return null;
	if (cleanEmail && !DEMO_EMAILS[cleanEmail]) return null;
	if (cleanIp && !DEMO_IPS[cleanIp]) return null;
	if (explicitDomain && !DEMO_DNSBL[explicitDomain]) return null;

	const flags = params.signalFlags || {};
	const signals = {
		phone: cleanPhone && flags.phone !== false ? tryResolveSampleSingle('phone', { phoneNumber: cleanPhone }) : null,
		whatsapp: cleanPhone && flags.whatsapp !== false ? tryResolveSampleSingle('whatsapp', { phoneNumber: cleanPhone }) : null,
		email: cleanEmail && flags.email !== false ? tryResolveSampleSingle('email', { email: cleanEmail }) : null,
		callerId:
			cleanPhone && flags.callerId !== false
				? tryResolveSampleSingle('callerId', {
						phoneNumber: cleanPhone,
						leadName: params.leadName || '',
						companyName: params.companyName || '',
					})
				: null,
		blacklist:
			(explicitDomain || (emailDomain && DEMO_DNSBL[emailDomain] && !DEMO_DNSBL[emailDomain].skippedFreeWebmail)) &&
			flags.blacklist !== false
				? tryResolveSampleSingle('blacklist', { target: explicitDomain || emailDomain })
				: null,
		ipGeo:
			cleanIp && flags.ipGeo !== false
				? tryResolveSampleSingle('ipGeo', {
						ip: cleanIp,
						expectedCountry: params.countryCode || DEMO_PHONES[cleanPhone]?.countryCode || '',
					})
				: null,
	};

	const scoring = calculateLeadTrustScore({ signals });
	return {
		trustScore: scoring.trustScore,
		riskLevel: scoring.riskLevel,
		summary: scoring.summary,
		reasons: scoring.reasons,
		positives: scoring.positives,
		warnings: [],
		sampleDemo: true,
		verifiedAt: new Date().toISOString(),
		signals,
	};
}
