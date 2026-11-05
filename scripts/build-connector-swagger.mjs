/**
 * Builds the Microsoft Power Platform Certified Connector package:
 *  - power-platform-connector/certified-connectors/ProWebLook/apiDefinition.swagger.json (4-space indent, Swagger 2.0)
 *  - power-platform-connector/certified-connectors/ProWebLook/apiProperties.json
 *  - power-platform-connector/certified-connectors/ProWebLook/settings.json
 *  - power-platform-connector/certified-connectors/ProWebLook/Readme.md
 *  - src/generated-swagger.js (served at https://m365.proweblook.com/swagger.json)
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const CONN_DIR = path.join(ROOT, 'power-platform-connector', 'certified-connectors', 'ProWebLook');
fs.mkdirSync(CONN_DIR, { recursive: true });

function prop(type, summary, description, extra = {}) {
	return {
		type,
		'x-ms-summary': summary,
		description,
		...extra,
	};
}

function stdResponses(schemaRef, okDescription) {
	return {
		'200': {
			description: okDescription,
			schema: { $ref: `#/definitions/${schemaRef}` },
		},
		'400': {
			description: 'Invalid input parameter format or missing required identifier.',
			schema: { $ref: '#/definitions/ErrorResponse' },
		},
		'401': {
			description: 'Missing or invalid ProWebLook API key.',
			schema: { $ref: '#/definitions/ErrorResponse' },
		},
		'402': {
			description: 'Insufficient ProWebLook credit balance for the requested verification signal.',
			schema: { $ref: '#/definitions/ErrorResponse' },
		},
		default: {
			description: 'Unexpected upstream verification error.',
			schema: { $ref: '#/definitions/ErrorResponse' },
		},
	};
}

const swagger = {
	swagger: '2.0',
	info: {
		title: 'ProWebLook',
		description:
			'Verify phone numbers, carrier and VoIP line types, silent WhatsApp registration, email deliverability, Caller ID (CNAM), DNSBL domain reputation, Edge GeoIP, and composite 0-100 lead and COD fraud trust scores across Power Automate, Logic Apps, and Microsoft Copilot Studio.',
		version: '1.0.0',
		contact: {
			name: 'ProWebLook Support',
			url: 'https://excel.proweblook.com/support',
			email: 'support@proweblook.com',
		},
	},
	'x-ms-connector-metadata': [
		{
			propertyName: 'Website',
			propertyValue: 'https://proweblook.com',
		},
		{
			propertyName: 'Privacy policy',
			propertyValue: 'https://excel.proweblook.com/privacy',
		},
		{
			propertyName: 'Categories',
			propertyValue: 'Data;Security',
		},
	],
	host: 'm365.proweblook.com',
	basePath: '/api/connector/v1',
	schemes: ['https'],
	consumes: ['application/json'],
	produces: ['application/json'],
	securityDefinitions: {
		api_key: {
			type: 'apiKey',
			in: 'header',
			name: 'X-API-Key',
			description: 'Your ProWebLook API key from https://proweblook.com/user/api_key.',
		},
	},
	security: [
		{
			api_key: [],
		},
	],
	paths: {
		'/verify-lead': {
			post: {
				operationId: 'VerifyLeadMultiSignal',
				summary: 'Verify lead and calculate fraud trust score',
				description:
					'Runs parallel multi-signal verification across phone carrier, WhatsApp, email deliverability, Caller ID, corporate DNSBL, and Edge GeoIP to return a composite 0 to 100 trust score and risk level.',
				'x-ms-visibility': 'important',
				parameters: [
					{
						name: 'body',
						in: 'body',
						required: true,
						'x-ms-summary': 'Lead Verification Request',
						description: 'Lead contact identifiers and optional signal toggles to evaluate.',
						schema: { $ref: '#/definitions/VerifyLeadRequest' },
					},
				],
				responses: stdResponses(
					'VerifyLeadResponse',
					'Composite lead trust score, risk classification, and multi-signal verification details.'
				),
			},
		},
		'/validate-phone': {
			post: {
				operationId: 'ValidatePhone',
				summary: 'Validate phone number and line type',
				description:
					'Validates an international or national phone number and returns carrier name, line type (Mobile, VoIP, or Landline), E.164 formatting, and geographic region.',
				'x-ms-visibility': 'important',
				parameters: [
					{
						name: 'body',
						in: 'body',
						required: true,
						'x-ms-summary': 'Phone Validation Request',
						description: 'Phone number and optional ISO country hint to validate.',
						schema: { $ref: '#/definitions/PhoneLookupRequest' },
					},
				],
				responses: stdResponses('PhoneValidationResult', 'Validated phone carrier, line type, and formatting metadata.'),
			},
		},
		'/check-whatsapp': {
			post: {
				operationId: 'CheckWhatsApp',
				summary: 'Check WhatsApp registration status',
				description:
					'Silently checks whether a phone number is registered on WhatsApp Personal or WhatsApp Business without sending any message or notification.',
				'x-ms-visibility': 'important',
				parameters: [
					{
						name: 'body',
						in: 'body',
						required: true,
						'x-ms-summary': 'WhatsApp Check Request',
						description: 'Phone number and optional ISO country code hint for silent WhatsApp verification.',
						schema: { $ref: '#/definitions/PhoneLookupRequest' },
					},
				],
				responses: stdResponses('WhatsAppCheckResult', 'Silent WhatsApp registration and Business account status.'),
			},
		},
		'/verify-email': {
			post: {
				operationId: 'VerifyEmail',
				summary: 'Verify email deliverability and disposable status',
				description:
					'Verifies email syntax, domain MX records, mailbox existence, corporate vs. webmail type, and 10-minute disposable temp-mail domains.',
				'x-ms-visibility': 'important',
				parameters: [
					{
						name: 'body',
						in: 'body',
						required: true,
						'x-ms-summary': 'Email Verification Request',
						description: 'Email address to check for deliverability and disposable domain risk.',
						schema: { $ref: '#/definitions/EmailVerifyRequest' },
					},
				],
				responses: stdResponses('EmailVerificationResult', 'Email deliverability verdict and domain classification.'),
			},
		},
		'/lookup-caller-id': {
			post: {
				operationId: 'LookupCallerId',
				summary: 'Look up Caller ID subscriber name',
				description:
					'Retrieves the registered CNAM Caller ID subscriber name and compares it against an optional lead or company name to detect spoofed identities.',
				'x-ms-visibility': 'important',
				parameters: [
					{
						name: 'body',
						in: 'body',
						required: true,
						'x-ms-summary': 'Caller ID Lookup Request',
						description: 'Phone number and optional lead or company name for CNAM identity matching.',
						schema: { $ref: '#/definitions/CallerIdRequest' },
					},
				],
				responses: stdResponses('CallerIdResult', 'Registered Caller ID subscriber name and fuzzy match status.'),
			},
		},
		'/check-blacklist': {
			post: {
				operationId: 'CheckBlacklist',
				summary: 'Check domain or IP against DNSBL blocklists',
				description:
					'Screens a corporate domain name or public IP address against more than 50 real-time DNSBL spam, phishing, and malware blocklists.',
				parameters: [
					{
						name: 'body',
						in: 'body',
						required: true,
						'x-ms-summary': 'Blacklist Check Request',
						description: 'Domain name, email address, or IP address to screen against DNSBL feeds.',
						schema: { $ref: '#/definitions/BlacklistCheckRequest' },
					},
				],
				responses: stdResponses('BlacklistCheckResult', 'DNSBL listing count and reputation verdict.'),
			},
		},
		'/lookup-ip-geo': {
			post: {
				operationId: 'LookupIpGeo',
				summary: 'Look up IP geolocation and VPN status',
				description:
					'Resolves IPv4 or IPv6 geolocation, ISP, ASN, datacenter/VPN/proxy flags, and optional billing country mismatch.',
				parameters: [
					{
						name: 'body',
						in: 'body',
						required: true,
						'x-ms-summary': 'IP Geolocation Request',
						description: 'Public IPv4 or IPv6 address and optional expected billing country code.',
						schema: { $ref: '#/definitions/IpGeoRequest' },
					},
				],
				responses: stdResponses('IpGeoResult', 'IP geolocation, ISP/ASN metadata, and proxy or country mismatch flags.'),
			},
		},
		'/balances': {
			get: {
				operationId: 'GetAccountBalances',
				summary: 'Get ProWebLook credit balances',
				description:
					'Returns current credit balances across Phone, WhatsApp, Email, Caller ID, Blacklist, and monthly Edge GeoIP pools for the authenticated ProWebLook account.',
				'x-ms-visibility': 'important',
				parameters: [],
				responses: stdResponses('AccountBalancesResponse', 'Current ProWebLook product credit balances.'),
			},
		},
	},
	definitions: {
		ErrorResponse: {
			type: 'object',
			description: 'Standard error payload returned when validation or authentication fails.',
			properties: {
				ok: prop('boolean', 'Success Flag', 'Indicates whether the API call succeeded.'),
				code: prop('string', 'Error Code', 'Machine-readable error identifier such as invalid_input or insufficient_credit.'),
				error: prop('string', 'Error Message', 'Human-readable explanation of why the request failed.'),
			},
		},
		PhoneLookupRequest: {
			type: 'object',
			required: ['phoneNumber'],
			description: 'Parameters for validating a phone number or checking WhatsApp registration.',
			properties: {
				phoneNumber: prop(
					'string',
					'Phone Number',
					'Phone number in E.164 format (+14155552671) or national digits.'
				),
				countryCode: prop(
					'string',
					'Country Code Hint',
					'Optional 2-letter ISO 3166-1 alpha-2 country hint such as US, IN, or GB.'
				),
			},
		},
		EmailVerifyRequest: {
			type: 'object',
			required: ['email'],
			description: 'Parameters for verifying an email address.',
			properties: {
				email: prop('string', 'Email Address', 'Email address to check for deliverability and disposable status.'),
			},
		},
		CallerIdRequest: {
			type: 'object',
			required: ['phoneNumber'],
			description: 'Parameters for looking up CNAM Caller ID and matching lead identity.',
			properties: {
				phoneNumber: prop(
					'string',
					'Phone Number',
					'Phone number in E.164 format (+14155552671) or national digits.'
				),
				countryCode: prop('string', 'Country Code Hint', 'Optional 2-letter ISO country code hint such as US or IN.'),
				leadName: prop(
					'string',
					'Lead Full Name',
					'Optional customer or lead name to compare against the registered Caller ID.'
				),
				companyName: prop(
					'string',
					'Company Name',
					'Optional business or organization name to compare against Caller ID.'
				),
			},
		},
		BlacklistCheckRequest: {
			type: 'object',
			required: ['target'],
			description: 'Parameters for screening a domain or IP against DNSBL blocklists.',
			properties: {
				target: prop(
					'string',
					'Domain or IP',
					'Corporate domain name, email address, or public IPv4/IPv6 address to screen.'
				),
				targetType: prop('string', 'Target Type', 'Optional target classification.', {
					enum: ['auto', 'domain', 'ip'],
					default: 'auto',
				}),
			},
		},
		IpGeoRequest: {
			type: 'object',
			required: ['ip'],
			description: 'Parameters for resolving IP geolocation and VPN or country mismatch status.',
			properties: {
				ip: prop('string', 'IP Address', 'Public IPv4 or IPv6 address to inspect.'),
				expectedCountry: prop(
					'string',
					'Expected Country Code',
					'Optional 2-letter ISO billing or shipping country code to detect location mismatches.'
				),
			},
		},
		VerifyLeadRequest: {
			type: 'object',
			description: 'Multi-signal lead or e-commerce COD order attributes to verify in parallel.',
			properties: {
				phoneNumber: prop('string', 'Phone Number', 'Customer or lead phone number in E.164 or national format.'),
				email: prop('string', 'Email Address', 'Customer or lead email address to verify.'),
				countryCode: prop('string', 'Country Code', 'Optional 2-letter ISO country code such as US, IN, GB, or AE.'),
				leadName: prop('string', 'Lead Name', 'Optional customer full name for Caller ID identity matching.'),
				companyName: prop('string', 'Company Name', 'Optional company name for B2B Caller ID matching.'),
				ip: prop('string', 'IP Address', 'Optional customer IPv4 or IPv6 address for GeoIP and VPN detection.'),
				domain: prop('string', 'Corporate Domain', 'Optional corporate domain override for DNSBL reputation screening.'),
			},
		},
		PhoneValidationResult: {
			type: 'object',
			description: 'Carrier, line type, and E.164 formatting details for a phone number.',
			properties: {
				valid: prop('boolean', 'Is Valid', 'True when the phone number is confirmed active and assignable.'),
				status: prop('string', 'Validation Status', 'Raw telecom status such as VALID_CONFIRMED or INVALID.'),
				verdict: prop('string', 'Phone Verdict', 'Normalized verdict: valid, invalid, or unknown.'),
				phoneNumber: prop('string', 'Normalized Input', 'Normalized digits or E.164 input string.'),
				formattedNumber: prop('string', 'International Format', 'Formatted international E.164 or display number.'),
				nationalFormat: prop('string', 'National Format', 'Locally formatted national phone string.'),
				lineType: prop('string', 'Line Type', 'Detected line type such as Mobile, VoIP, or Fixed line.'),
				isVoip: prop('boolean', 'Is VoIP Line', 'True when the phone number belongs to a virtual or VoIP carrier.'),
				isLandline: prop('boolean', 'Is Landline', 'True when the phone number is a fixed landline.'),
				carrier: prop('string', 'Carrier Name', 'Operating mobile or fixed telecom network name.'),
				location: prop('string', 'Region or City', 'Geographic city or state associated with the number.'),
				countryCode: prop('string', 'ISO Country Code', 'Resolved 2-letter ISO country code.'),
				countryName: prop('string', 'Country Name', 'Full country name for the phone number.'),
				timezone: prop('string', 'Timezone', 'Primary IANA timezone identifier for the number region.'),
			},
		},
		WhatsAppCheckResult: {
			type: 'object',
			description: 'Silent WhatsApp registration and account classification result.',
			properties: {
				phoneNumber: prop('string', 'Phone Number', 'Normalized phone number checked on WhatsApp.'),
				registered: prop('boolean', 'Registered on WhatsApp', 'True when an active WhatsApp account exists.'),
				isBusiness: prop('boolean', 'Is WhatsApp Business', 'True when registered as a WhatsApp Business profile.'),
				accountType: prop('string', 'Account Type', 'Detected WhatsApp tier: personal, business, or none.'),
				verdict: prop('string', 'WhatsApp Verdict', 'Normalized status: on_whatsapp or not_on_whatsapp.'),
				countryCode: prop('string', 'Country Code', 'Resolved 2-letter ISO country code.'),
			},
		},
		EmailVerificationResult: {
			type: 'object',
			description: 'Email deliverability, MX/mailbox status, and disposable domain result.',
			properties: {
				email: prop('string', 'Email Address', 'Normalized lowercase email address.'),
				domain: prop('string', 'Email Domain', 'Extracted domain portion of the email address.'),
				deliverable: prop('boolean', 'Is Deliverable', 'True when the mailbox is verified as safe to send.'),
				verdict: prop('string', 'Email Verdict', 'Verdict: deliverable, undeliverable, disposable, risky, or unknown.'),
				isDisposable: prop('boolean', 'Is Disposable Domain', 'True when hosted by a temporary or throwaway mail provider.'),
				isFreeWebmail: prop('boolean', 'Is Free Webmail', 'True for public consumer mailboxes such as Gmail or Outlook.'),
				emailType: prop('string', 'Email Category', 'Classification: corporate, free_webmail, or disposable.'),
				reasonCode: prop('string', 'Reason Code', 'Diagnostic verification reason code.'),
				emailFormatValid: prop('boolean', 'Syntax Valid', 'True when the address conforms to RFC email syntax.'),
				mailboxExists: prop('boolean', 'Mailbox Exists', 'True when MX/SMTP verification confirms the mailbox exists.'),
			},
		},
		CallerIdResult: {
			type: 'object',
			description: 'Caller ID (CNAM) subscriber lookup and lead-name comparison result.',
			properties: {
				phoneNumber: prop('string', 'Phone Number', 'Normalized phone number queried for Caller ID.'),
				found: prop('boolean', 'Caller ID Found', 'True when a registered subscriber name was found.'),
				verdict: prop('string', 'Caller ID Verdict', 'Normalized lookup outcome: found or not_found.'),
				callerName: prop('string', 'Caller Name', 'Registered subscriber or business display name.'),
				callerType: prop('string', 'Caller Type', 'Subscriber classification such as CONSUMER or BUSINESS.'),
				nameMatch: prop(
					'string',
					'Lead Name Match',
					'Fuzzy comparison result against leadName: match, partial, mismatch, or not_checked.'
				),
				carrier: prop('string', 'Carrier Name', 'Telecom carrier reported by CNAM registry.'),
				lineType: prop('string', 'Line Type', 'Line classification reported alongside Caller ID.'),
				location: prop('string', 'Location', 'Geographic region associated with the subscriber.'),
				countryCode: prop('string', 'Country Code', '2-letter ISO country code.'),
			},
		},
		BlacklistCheckResult: {
			type: 'object',
			description: 'Domain or IP DNSBL reputation check result.',
			properties: {
				target: prop('string', 'Checked Target', 'Normalized domain name or IP address screened.'),
				targetType: prop('string', 'Target Type', 'Resolved target category: domain or ip.'),
				listed: prop('boolean', 'Is Blacklisted', 'True when listed on one or more DNSBL blocklists.'),
				listedCount: prop('integer', 'Listed Count', 'Number of DNSBL blocklists flagging the target.', { format: 'int32' }),
				checkedCount: prop('integer', 'Feeds Checked', 'Total number of DNSBL feeds queried.', { format: 'int32' }),
				skippedFreeWebmail: prop('boolean', 'Skipped Free Webmail', 'True when free consumer webmail domain was skipped.'),
				verdict: prop('string', 'Blacklist Verdict', 'Normalized reputation verdict: clean, listed, or clean_free_webmail.'),
			},
		},
		IpGeoResult: {
			type: 'object',
			description: 'Edge IP geolocation, ISP/ASN, VPN/hosting, and country mismatch result.',
			properties: {
				ip: prop('string', 'IP Address', 'Resolved IPv4 or IPv6 address.'),
				countryCode: prop('string', 'IP Country Code', '2-letter ISO country code of the IP address.'),
				countryName: prop('string', 'IP Country Name', 'Full country name of the IP address.'),
				region: prop('string', 'Region or State', 'Administrative region or state name.'),
				city: prop('string', 'City', 'Resolved city name.'),
				isp: prop('string', 'ISP or Organization', 'Internet service provider or autonomous system organization.'),
				asn: prop('string', 'ASN', 'Autonomous System Number identifier.'),
				timezone: prop('string', 'Timezone', 'IANA timezone at the IP location.'),
				isHostingOrProxy: prop('boolean', 'Is VPN or Hosting', 'True when the IP belongs to a VPN, proxy, or cloud datacenter.'),
				expectedCountry: prop('string', 'Expected Country', 'Expected 2-letter ISO country code used for comparison.'),
				countryMismatch: prop('boolean', 'Country Mismatch', 'True when IP country differs from expectedCountry.'),
				skippedPrivateIp: prop('boolean', 'Private IP Skipped', 'True when the input IP is a private RFC1918 or loopback address.'),
				verdict: prop('string', 'IP Verdict', 'Classification: residential_isp, hosting_or_vpn, country_mismatch, or private_ip_skipped.'),
			},
		},
		VerifyLeadSignals: {
			type: 'object',
			description: 'Detailed per-signal outputs returned by multi-signal lead verification.',
			properties: {
				phone: { $ref: '#/definitions/PhoneValidationResult' },
				whatsapp: { $ref: '#/definitions/WhatsAppCheckResult' },
				email: { $ref: '#/definitions/EmailVerificationResult' },
				callerId: { $ref: '#/definitions/CallerIdResult' },
				blacklist: { $ref: '#/definitions/BlacklistCheckResult' },
				ipGeo: { $ref: '#/definitions/IpGeoResult' },
			},
		},
		VerifyLeadResponse: {
			type: 'object',
			description: 'Composite 0-100 lead trust score, risk tier, and underlying 6-signal results.',
			properties: {
				trustScore: prop(
					'integer',
					'Trust Score (0-100)',
					'Composite lead trustworthiness score from 0 (highest fraud risk) to 100 (verified low risk).',
					{ format: 'int32' }
				),
				riskLevel: prop(
					'string',
					'Risk Level',
					'Actionable risk tier: trusted (75-100), review (45-74), or high_risk (0-44).'
				),
				summary: prop(
					'string',
					'Verification Summary',
					'Human-readable summary of key risk flags and positive trust signals.'
				),
				reasons: {
					type: 'array',
					'x-ms-summary': 'Risk Flags',
					description: 'List of negative fraud or deliverability indicators triggered by this lead.',
					items: { type: 'string' },
				},
				positives: {
					type: 'array',
					'x-ms-summary': 'Positive Signals',
					description: 'List of verified positive telecom, email, and identity signals.',
					items: { type: 'string' },
				},
				verifiedAt: prop('string', 'Verified Timestamp', 'ISO 8601 UTC timestamp when verification completed.'),
				signals: { $ref: '#/definitions/VerifyLeadSignals' },
			},
		},
		AccountBalancesResponse: {
			type: 'object',
			description: 'Available ProWebLook API credit balances by verification product.',
			properties: {
				accountEmail: prop('string', 'Account Email', 'Email address of the authenticated ProWebLook account.'),
				phoneCredits: prop('integer', 'Phone Credits', 'Remaining phone validation credits.', { format: 'int32' }),
				whatsappCredits: prop('integer', 'WhatsApp Credits', 'Remaining WhatsApp verification credits.', { format: 'int32' }),
				emailCredits: prop('integer', 'Email Credits', 'Remaining email verification credits.', { format: 'int32' }),
				callerIdCredits: prop('integer', 'Caller ID Credits', 'Remaining Caller ID lookup credits.', { format: 'int32' }),
				blacklistCredits: prop('integer', 'Blacklist Credits', 'Remaining DNSBL blacklist check credits.', { format: 'int32' }),
				ipGeoMonthlyFreeQuota: prop('integer', 'Monthly Edge GeoIP Quota', 'Included monthly Edge GeoIP lookups.', { format: 'int32' }),
			},
		},
	},
};

const apiProperties = {
	properties: {
		connectionParameters: {
			api_key: {
				type: 'securestring',
				uiDefinition: {
					displayName: 'API Key',
					description: 'Enter your ProWebLook API key from https://proweblook.com/user/api_key.',
					tooltip: 'Provide your ProWebLook API key',
					constraints: {
						tabIndex: 2,
						clearText: false,
						required: 'true',
					},
				},
			},
		},
		iconBrandColor: '#0F172A',
		capabilities: [],
		publisher: 'ProWebLook',
		stackOwner: 'ProWebLook',
	},
};

const settingsJson = {
	connectorId: 'ProWebLook',
	environment: '',
	apiProperties: 'apiProperties.json',
	apiDefinition: 'apiDefinition.swagger.json',
	icon: 'icon.png',
	powerAppsUrl: 'https://api.powerapps.com',
	powerAppsApiVersion: '2016-11-01',
};

const readmeMd = `# ProWebLook

ProWebLook provides real-time 6-signal phone carrier & VoIP detection, silent WhatsApp personal/Business verification, email deliverability & disposable temp-mail screening, Caller ID (CNAM) identity matching, corporate DNSBL reputation checks, Edge IP geolocation, and composite 0–100 Lead & COD Fraud Trust Scoring for Microsoft Power Automate, Azure Logic Apps, Power Apps, and Microsoft Copilot Studio.

## Publisher: ProWebLook

## Prerequisites

- An active ProWebLook account ([sign up free at proweblook.com](https://proweblook.com/auth/signup) — includes 50 free credits per product plus 20,000 monthly Edge GeoIP lookups).

## Supported Operations

### Verify lead and calculate fraud trust score (\`VerifyLeadMultiSignal\`)
Runs parallel multi-signal verification across phone carrier, silent WhatsApp, email deliverability, Caller ID (CNAM), corporate DNSBL, and Edge GeoIP to compute a composite \`0–100\` \`trustScore\` and actionable \`riskLevel\` (\`trusted\`, \`review\`, or \`high_risk\`).

### Validate phone number and line type (\`ValidatePhone\`)
Validates national or E.164 phone numbers across 220+ countries and returns operating carrier, line type (\`Mobile\`, \`VoIP\`, \`Fixed line\`), and international formatting.

### Check WhatsApp registration status (\`CheckWhatsApp\`)
Silently checks whether a phone number has an active WhatsApp Personal or WhatsApp Business account without sending a message.

### Verify email deliverability and disposable status (\`VerifyEmail\`)
Checks RFC syntax, domain MX records, mailbox existence, corporate vs. free webmail classification, and 10-minute disposable temp-mail providers.

### Look up Caller ID subscriber name (\`LookupCallerId\`)
Retrieves the registered CNAM Caller ID display name and optionally compares it against \`leadName\` or \`companyName\` (\`match\`, \`partial\`, or \`mismatch\`).

### Check domain or IP against DNSBL blocklists (\`CheckBlacklist\`)
Screens corporate domains or public IP addresses against 50+ real-time spam, phishing, and malware blocklists.

### Look up IP geolocation and VPN status (\`LookupIpGeo\`)
Resolves IPv4/IPv6 country, city, ASN/ISP, VPN/hosting flags, and billing country mismatch.

### Get ProWebLook credit balances (\`GetAccountBalances\`)
Returns live remaining credit balances across Phone, WhatsApp, Email, Caller ID, and Blacklist pools.

## Obtaining Credentials

1. Sign in or register at [https://proweblook.com/auth/signup](https://proweblook.com/auth/signup).
2. Navigate to **Dashboard → API Keys** (\`https://proweblook.com/user/api_key\`).
3. Copy your API key and paste it into the **API Key** connection prompt in Power Automate, Logic Apps, or Copilot Studio.

## Getting Started

1. In Power Automate or Copilot Studio, add a new action and search for **ProWebLook**.
2. Choose **Verify lead and calculate fraud trust score** when a new Dynamics 365 Lead, Microsoft Forms submission, HubSpot contact, or Shopify/WooCommerce order arrives.
3. Branch your flow on \`riskLevel\` (\`trusted\` → route to sales / fulfill order; \`high_risk\` → hold order or flag in Microsoft Teams).

## Known Issues and Limitations

- Caller ID (CNAM) coverage is strongest in North America, India, UK, and major European/MENA mobile networks; numbers without public registry records return \`found: false\`.
- Private or loopback RFC1918 IP addresses are automatically skipped (\`skippedPrivateIp: true\`) without consuming quota.

## Deployment Instructions

Run the following command using the Microsoft Power Platform Connectors CLI (\`paconn\`):

\`\`\`bash
paconn create --api-def apiDefinition.swagger.json --api-prop apiProperties.json --icon icon.png
\`\`\`
`;

const swaggerStr = JSON.stringify(swagger, null, 4) + '\n';
const propsStr = JSON.stringify(apiProperties, null, 4) + '\n';
const settingsStr = JSON.stringify(settingsJson, null, 4) + '\n';

fs.writeFileSync(path.join(CONN_DIR, 'apiDefinition.swagger.json'), swaggerStr);
fs.writeFileSync(path.join(CONN_DIR, 'apiProperties.json'), propsStr);
fs.writeFileSync(path.join(CONN_DIR, 'settings.json'), settingsStr);
fs.writeFileSync(path.join(CONN_DIR, 'Readme.md'), readmeMd);

fs.writeFileSync(
	path.join(ROOT, 'src', 'generated-swagger.js'),
	`// Auto-generated by scripts/build-connector-swagger.mjs — do not edit manually.\nexport const SWAGGER_JSON_TEXT = ${JSON.stringify(swaggerStr)};\n`
);

console.log('Generated Power Platform Certified Connector files + src/generated-swagger.js');
