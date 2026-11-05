/**
 * Input normalizers, consumer/disposable email classifiers & Caller ID fuzzy matcher
 * shared by the Cloudflare Edge Worker, Excel Custom Functions, and Power Platform Gateway.
 */

export const DISPOSABLE_EMAIL_DOMAINS = new Set([
	'10minutemail.com',
	'10minutemail.net',
	'burnermail.io',
	'dispostable.com',
	'emailondeck.com',
	'fakeinbox.com',
	'getairmail.com',
	'getnada.com',
	'guerrillamail.com',
	'guerrillamailblock.com',
	'inboxkitten.com',
	'mailcatch.com',
	'maildrop.cc',
	'mailinator.com',
	'mailnesia.com',
	'mohmal.com',
	'mytemp.email',
	'sharklasers.com',
	'temp-mail.io',
	'temp-mail.org',
	'tempail.com',
	'tempmail.com',
	'tempmailo.com',
	'throwawaymail.com',
	'trashmail.com',
	'yopmail.com',
	'yopmail.fr',
	'yopmail.net',
]);

export const FREE_WEBMAIL_DOMAINS = new Set([
	'gmail.com',
	'googlemail.com',
	'yahoo.com',
	'yahoo.co.in',
	'yahoo.co.uk',
	'outlook.com',
	'hotmail.com',
	'live.com',
	'msn.com',
	'icloud.com',
	'me.com',
	'mac.com',
	'aol.com',
	'proton.me',
	'protonmail.com',
	'zoho.com',
	'zohomail.in',
	'rediffmail.com',
	'gmx.com',
	'yandex.com',
]);

export const HOSTING_PROXY_KEYWORDS = [
	'digitalocean',
	'amazonaws',
	'amazon.com',
	'amazon technologies',
	'google cloud',
	'hetzner',
	'ovh',
	'linode',
	'akamai connected',
	'vultr',
	'choopa',
	'm247',
	'datacamp',
	'contabo',
	'hostinger',
	'tzulo',
	'leaseweb',
	'quadranet',
	'vpn',
	'tor exit',
	'mullvad',
	'nordvpn',
	'expressvpn',
	'surfshark',
	'protonvpn',
	'datacenter',
	'hosting',
	'proxy',
];

export function normalizeEmailInput(input = '') {
	const raw = String(input || '')
		.trim()
		.replace(/^mailto:/i, '')
		.split('?')[0]
		.trim()
		.toLowerCase();
	if (!raw || raw.length > 254) return '';
	const match = raw.match(/^[a-z0-9._%+-]+@([a-z0-9.-]+\.[a-z]{2,24})$/);
	if (!match) return '';
	return raw;
}

export function extractEmailDomain(email = '') {
	const clean = normalizeEmailInput(email);
	if (!clean) return '';
	const at = clean.lastIndexOf('@');
	return at > 0 ? clean.slice(at + 1) : '';
}

export function isDisposableEmailDomain(email = '') {
	const domain = extractEmailDomain(email) || normalizeDomainInput(email);
	return Boolean(domain && DISPOSABLE_EMAIL_DOMAINS.has(domain));
}

export function normalizePhoneInput(input = '') {
	let raw = String(input ?? '')
		.trim()
		.replace(/^(tel:|whatsapp:\/\/send\?phone=|https?:\/\/wa\.me\/)/i, '');
	raw = raw.split('&')[0].split('?')[0].trim();
	if (!raw) return '';
	const hasLeadingPlus = raw.startsWith('+') || raw.startsWith('00');
	const digits = raw.replace(/\D+/g, '');
	if (digits.length < 7 || digits.length > 15) return '';
	if (raw.startsWith('00') && digits.length >= 9) {
		return `+${digits.slice(2)}`;
	}
	return hasLeadingPlus ? `+${digits}` : digits;
}

export function isValidIpAddress(input = '') {
	const raw = String(input || '').trim();
	if (!raw) return false;
	const ipv4 = raw.match(/^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/);
	if (ipv4) {
		return ipv4.slice(1).every((oct) => Number(oct) >= 0 && Number(oct) <= 255);
	}
	if (raw.includes(':') && /^[0-9a-fA-F:]{2,39}$/.test(raw) && (raw.match(/:/g) || []).length >= 2) {
		return true;
	}
	return false;
}

export function isPrivateOrReservedIp(ip = '') {
	const raw = String(ip || '').trim();
	if (!isValidIpAddress(raw)) return true;
	if (raw.includes(':')) {
		const lower = raw.toLowerCase();
		return lower === '::1' || lower.startsWith('fe80:') || lower.startsWith('fc') || lower.startsWith('fd');
	}
	const [a, b] = raw.split('.').map(Number);
	if (a === 10 || a === 127 || a === 0) return true;
	if (a === 172 && b >= 16 && b <= 31) return true;
	if (a === 192 && b === 168) return true;
	if (a === 169 && b === 254) return true;
	return false;
}

export function normalizeDomainInput(input = '') {
	let raw = String(input || '').trim().toLowerCase();
	if (!raw) return '';
	if (raw.includes('@')) {
		raw = extractEmailDomain(raw);
	}
	if (!raw || isValidIpAddress(raw)) return '';
	raw = raw.replace(/^[a-z]+:\/\//i, '').replace(/^www\./i, '');
	raw = raw.split('/')[0].split('?')[0].split('#')[0].split(':')[0].trim();
	if (!raw || raw.length > 253) return '';
	if (!/^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?)+$/.test(raw)) {
		return '';
	}
	return raw;
}

export function compareCallerIdWithLead(callerName = '', contactName = '', companyName = '') {
	const cleanCaller = String(callerName || '').trim();
	if (!cleanCaller) return 'unavailable';

	const stopWords = new Set([
		'inc',
		'llc',
		'ltd',
		'pvt',
		'limited',
		'corp',
		'corporation',
		'company',
		'co',
		'gmbh',
		'plc',
		'the',
		'and',
		'mr',
		'mrs',
		'ms',
		'dr',
	]);
	const tokenize = (str) =>
		String(str || '')
			.toLowerCase()
			.replace(/[^a-z0-9\s]/g, ' ')
			.split(/\s+/)
			.filter((t) => t.length >= 2 && !stopWords.has(t));

	const callerTokens = tokenize(cleanCaller);
	const targetTokens = [...tokenize(contactName), ...tokenize(companyName)];
	if (callerTokens.length === 0 || targetTokens.length === 0) {
		return 'uncompared';
	}

	const callerJoined = callerTokens.join(' ');
	const targetJoined = targetTokens.join(' ');
	if (callerJoined.length >= 4 && targetJoined.length >= 4) {
		if (callerJoined.includes(targetJoined) || targetJoined.includes(callerJoined)) {
			return 'match';
		}
	}

	const shared = callerTokens.filter((ct) =>
		targetTokens.some((tt) => ct === tt || (ct.length >= 4 && tt.startsWith(ct)) || (tt.length >= 4 && ct.startsWith(tt))),
	);
	if (shared.length >= 2 || (shared.length === 1 && shared[0].length >= 4 && (callerTokens.length <= 2 || targetTokens.length <= 2))) {
		return 'match';
	}
	if (shared.length >= 1) {
		return 'partial';
	}
	return 'mismatch';
}
