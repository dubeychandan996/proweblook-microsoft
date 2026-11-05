/**
 * Public Landing, Support/Docs, Privacy Policy & EULA/Terms Pages
 * Served on https://excel.proweblook.com and https://m365.proweblook.com for Microsoft AppSource
 * certification & enterprise Microsoft 365 / Power Platform customers.
 */

function layout(title, subtitle, bodyHtml) {
	return `<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8" />
	<meta name="viewport" content="width=device-width, initial-scale=1.0" />
	<title>${title} — ProWebLook for Microsoft 365</title>
	<link rel="icon" type="image/png" href="/assets/icon-32.png" />
	<style>
		:root {
			--bg: #f8fafc;
			--card: #ffffff;
			--navy: #0f172a;
			--blue: #0f6cbd;
			--text: #0f172a;
			--muted: #475569;
			--border: #e2e8f0;
		}
		* { box-sizing: border-box; }
		body {
			margin: 0;
			font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
			background: var(--bg);
			color: var(--text);
			line-height: 1.6;
		}
		.hero {
			background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
			color: #ffffff;
			padding: 44px 20px 38px;
		}
		.wrap {
			max-width: 960px;
			margin: 0 auto;
		}
		.nav {
			display: flex;
			align-items: center;
			justify-content: space-between;
			gap: 16px;
			margin-bottom: 28px;
			flex-wrap: wrap;
		}
		.brand {
			display: inline-flex;
			align-items: center;
			gap: 10px;
			color: #ffffff;
			text-decoration: none;
			font-weight: 700;
			font-size: 16px;
		}
		.brand img { width: 30px; height: 30px; border-radius: 7px; }
		.nav-links { display: flex; gap: 18px; font-size: 13.5px; }
		.nav-links a { color: #cbd5e1; text-decoration: none; }
		.nav-links a:hover { color: #ffffff; }
		h1 { font-size: 30px; line-height: 1.2; margin: 0 0 10px; letter-spacing: -0.02em; }
		.lead { color: #cbd5e1; font-size: 16px; max-width: 720px; margin: 0; }
		.cta-row { display: flex; gap: 12px; margin-top: 22px; flex-wrap: wrap; }
		.btn {
			display: inline-flex;
			align-items: center;
			padding: 10px 16px;
			border-radius: 7px;
			font-weight: 600;
			font-size: 14px;
			text-decoration: none;
		}
		.btn-primary { background: #0f6cbd; color: #ffffff; }
		.btn-outline { background: rgba(255,255,255,0.1); color: #ffffff; border: 1px solid rgba(255,255,255,0.22); }
		main { padding: 32px 20px 56px; }
		.card {
			background: var(--card);
			border: 1px solid var(--border);
			border-radius: 10px;
			padding: 24px;
			margin-bottom: 20px;
			box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
		}
		h2 { margin: 0 0 12px; font-size: 20px; }
		h3 { margin: 18px 0 8px; font-size: 16px; }
		table { width: 100%; border-collapse: collapse; font-size: 13.5px; margin-top: 10px; }
		th, td { text-align: left; padding: 10px 12px; border-bottom: 1px solid var(--border); vertical-align: top; }
		th { background: #f1f5f9; font-weight: 700; }
		code {
			font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
			background: #f1f5f9;
			padding: 2px 6px;
			border-radius: 4px;
			font-size: 12.5px;
		}
		footer {
			border-top: 1px solid var(--border);
			padding: 22px 20px;
			font-size: 13px;
			color: var(--muted);
			text-align: center;
		}
		footer a { color: var(--blue); text-decoration: none; }
	</style>
</head>
<body>
	<header class="hero">
		<div class="wrap">
			<div class="nav">
				<a class="brand" href="/"><img src="/assets/icon-64.png" alt="ProWebLook" /> ProWebLook for Microsoft 365</a>
				<div class="nav-links">
					<a href="/">Overview</a>
					<a href="/support">Setup &amp; Support</a>
					<a href="/privacy">Privacy</a>
					<a href="/terms">Terms</a>
					<a href="https://proweblook.com" target="_blank" rel="noopener">ProWebLook.com ↗</a>
				</div>
			</div>
			<h1>${title}</h1>
			<p class="lead">${subtitle}</p>
			<div class="cta-row">
				<a class="btn btn-primary" href="/manifest.xml" download="proweblook-excel-manifest.xml">Download Excel Manifest (XML)</a>
				<a class="btn btn-outline" href="/taskpane.html">Launch Live Taskpane Preview</a>
				<a class="btn btn-outline" href="https://m365.proweblook.com/swagger.json">Power Platform Swagger 2.0</a>
			</div>
		</div>
	</header>
	<main>
		<div class="wrap">
			${bodyHtml}
		</div>
	</main>
	<footer>
		© ${new Date().getUTCFullYear()} ProWebLook · <a href="/support">Support</a> · <a href="/privacy">Privacy Policy</a> · <a href="/terms">Terms of Use</a> · <a href="mailto:support@proweblook.com">support@proweblook.com</a>
	</footer>
</body>
</html>`;
}

export function renderHomePage() {
	return layout(
		'6-Signal Lead, WhatsApp & COD Fraud Verification for Excel & Power Platform',
		'Enrich spreadsheets with =PROWEBLOOK.* custom formulas and automate Dynamics 365, Power Automate, Azure Logic Apps, and Microsoft Copilot Studio workflows.',
		`
		<div class="card">
			<h2>1. ProWebLook for Microsoft Excel (Add-in + Custom Functions)</h2>
			<p>Verify thousands of CRM leads, B2B prospects, and e-commerce Cash-on-Delivery (COD) orders directly inside Excel Desktop (Windows &amp; Mac) and Excel on the Web — including a 1-click 6-row Sample Sheet that runs immediately with zero signup required.</p>
			<table>
				<thead>
					<tr><th>Excel Formula</th><th>Description</th></tr>
				</thead>
				<tbody>
					<tr><td><code>=PROWEBLOOK.TRUSTSCORE(phone, email, country, name, ip)</code></td><td>Composite 0–100 Lead &amp; COD Fraud Trust Score across all 6 telecom &amp; network signals.</td></tr>
					<tr><td><code>=PROWEBLOOK.WHATSAPP(phone, country, field)</code></td><td>Silent WhatsApp personal vs. Business registration check without sending messages.</td></tr>
					<tr><td><code>=PROWEBLOOK.PHONE(phone, country, field)</code></td><td>Carrier lookup, line type (Mobile, VoIP, Landline), and E.164 formatting across 220+ countries.</td></tr>
					<tr><td><code>=PROWEBLOOK.EMAIL(email, field)</code></td><td>SMTP/MX mailbox deliverability verification + 10-minute disposable email blocking.</td></tr>
					<tr><td><code>=PROWEBLOOK.CALLERID(phone, lead_name, country)</code></td><td>Registered CNAM Caller ID subscriber lookup with fuzzy lead-name matching.</td></tr>
					<tr><td><code>=PROWEBLOOK.DNSBL(domain_or_ip, field)</code></td><td>Screens corporate domains &amp; IPs against 50+ real-time spam &amp; malware DNSBLs.</td></tr>
					<tr><td><code>=PROWEBLOOK.IPGEO(ip, country_hint)</code></td><td>Sub-millisecond Edge GeoIP, ASN/ISP, VPN/proxy detection, and billing country mismatch check.</td></tr>
				</tbody>
			</table>
		</div>

		<div class="card">
			<h2>2. Microsoft Power Automate, Logic Apps &amp; Copilot Studio Connector</h2>
			<p>Connect Dynamics 365 Sales, HubSpot, Microsoft Forms, SharePoint Lists, Teams, and Copilot Studio AI agents to ProWebLook via <code>https://m365.proweblook.com/api/connector/v1</code>.</p>
			<table>
				<thead>
					<tr><th>Operation ID</th><th>Action Summary</th></tr>
				</thead>
				<tbody>
					<tr><td><code>VerifyLeadMultiSignal</code></td><td>Run parallel 6-signal verification and compute a 0–100 Lead &amp; COD Fraud Trust Score.</td></tr>
					<tr><td><code>ValidatePhone</code></td><td>Validate phone number carrier, line type (Mobile/VoIP/Landline), and E.164 formatting.</td></tr>
					<tr><td><code>CheckWhatsApp</code></td><td>Silently check WhatsApp personal or Business registration status.</td></tr>
					<tr><td><code>VerifyEmail</code></td><td>Verify email MX/SMTP deliverability and detect disposable temp-mail domains.</td></tr>
					<tr><td><code>LookupCallerId</code></td><td>Look up Caller ID (CNAM) subscriber identity and compare against lead/company name.</td></tr>
					<tr><td><code>CheckBlacklist</code></td><td>Screen domain or IP against 50+ DNSBL spam and malware blocklists.</td></tr>
					<tr><td><code>LookupIpGeo</code></td><td>Resolve IP geolocation, ASN/ISP, VPN/hosting flags, and country mismatch.</td></tr>
					<tr><td><code>GetAccountBalances</code></td><td>Check live credit balances across Phone, WhatsApp, Email, Caller ID, and DNSBL pools.</td></tr>
				</tbody>
			</table>
		</div>
		`,
	);
}

export function renderSupportPage() {
	return layout(
		'Setup, Sideloading & Customer Support',
		'Step-by-step instructions for Microsoft Excel (Web, Windows, Mac), Power Automate, and Microsoft Copilot Studio.',
		`
		<div class="card">
			<h2>Quick Test in Excel (Zero Account Required)</h2>
			<ol>
				<li>Install <strong>ProWebLook Lead &amp; WhatsApp Verifier</strong> from Microsoft AppSource, or sideload <a href="/manifest.xml" download="proweblook-excel-manifest.xml"><code>manifest.xml</code></a> in Excel on the Web via <strong>Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in</strong>.</li>
				<li>On the Excel <strong>Home</strong> tab, click <strong>Verify Leads</strong> (or click <strong>Sample Sheet</strong>).</li>
				<li>Click <strong>Insert Sample Sheet</strong> to populate 6 realistic B2B &amp; COD test rows, then click <strong>Verify &amp; Enrich Sheet</strong> to watch color-coded Trust Scores, Carrier, WhatsApp, Caller ID, Email, DNSBL, and GeoIP columns populate automatically.</li>
				<li>Try typing <code>=PROWEBLOOK.WHATSAPP(B2, D2)</code> or <code>=PROWEBLOOK.TRUSTSCORE(B2, C2, D2, A2, E2)</code> in any cell!</li>
			</ol>
		</div>

		<div class="card">
			<h2>Connecting Your ProWebLook API Key</h2>
			<ol>
				<li>Create a free account at <a href="https://proweblook.com/auth/signup" target="_blank" rel="noopener">proweblook.com/auth/signup</a> (includes 50 free credits per product + 20,000 monthly Edge GeoIP lookups).</li>
				<li>Copy your API key from <strong>Dashboard → API Keys</strong>.</li>
				<li>In Excel: open the <strong>API Key</strong> tab in the ProWebLook taskpane, paste your key, and click <strong>Save &amp; Verify API Key</strong>.</li>
				<li>In Power Automate / Logic Apps / Copilot Studio: add the <strong>ProWebLook</strong> connector and paste your API key when prompted for the connection key.</li>
			</ol>
		</div>

		<div class="card">
			<h2>Dedicated Engineering Support</h2>
			<p>Need help with bulk Excel enrichment, Microsoft 365 Admin Center deployment, or Power Automate flows?</p>
			<ul>
				<li><strong>Email Support:</strong> <a href="mailto:support@proweblook.com">support@proweblook.com</a> (response within 1 business day)</li>
				<li><strong>API Status &amp; Docs:</strong> <a href="https://proweblook.com/api-docs" target="_blank" rel="noopener">https://proweblook.com/api-docs</a></li>
				<li><strong>Contact Portal:</strong> <a href="https://proweblook.com/contact" target="_blank" rel="noopener">https://proweblook.com/contact</a></li>
			</ul>
		</div>
		`,
	);
}

export function renderPrivacyPage() {
	return layout(
		'Privacy Policy',
		'How ProWebLook for Microsoft Excel and Microsoft Power Platform handles verification requests and protects customer data.',
		`
		<div class="card">
			<h2>1. Scope</h2>
			<p>This Privacy Policy governs the <strong>ProWebLook Lead and WhatsApp Verifier</strong> (<strong>ProWebLook Lead &amp; WhatsApp Verifier</strong>) Microsoft Excel Add-in (<code>excel.proweblook.com</code>) and the <strong>ProWebLook</strong> Microsoft Power Platform &amp; Copilot Studio Connector (<code>m365.proweblook.com</code>), operated by ProWebLook ("we", "us").</p>

			<h2>2. Data Processed</h2>
			<ul>
				<li><strong>Verification Inputs:</strong> Phone numbers, email addresses, domain names, IP addresses, optional 2-letter country hints, and optional lead/customer names explicitly passed to <code>=PROWEBLOOK.*</code> Excel formulas, the Bulk Sheet Enricher, or Power Platform Connector actions.</li>
				<li><strong>Authentication Credentials:</strong> Your ProWebLook API key is stored locally inside your browser/Excel WebView storage (for Excel) or encrypted inside Microsoft Power Platform Connection Manager (for Power Automate / Logic Apps / Copilot Studio) and transmitted exclusively over HTTPS (TLS 1.2/1.3) via the <code>X-API-Key</code> header.</li>
				<li><strong>Abuse-Prevention Quota Counter:</strong> In unauthenticated Starter Sandbox mode, a one-way SHA-256 hash of a random installation identifier and client IP is stored with a 30-day TTL solely to enforce the free trial lookup limit.</li>
			</ul>

			<h2>3. Zero Worksheet Surveillance &amp; No Data Resale</h2>
			<p>ProWebLook only reads worksheet cells when you explicitly invoke a <code>=PROWEBLOOK.*</code> formula or click <strong>Verify &amp; Enrich Sheet</strong>. We <strong>never</strong> sell, rent, broker, or use your spreadsheet leads or CRM contacts for marketing or third-party prospecting.</p>

			<h2>4. GDPR, CCPA &amp; Contact</h2>
			<p>For data-protection inquiries, DPA requests, or deletion requests, contact our privacy team at <a href="mailto:support@proweblook.com">support@proweblook.com</a>.</p>
		</div>
		`,
	);
}

export function renderTermsPage() {
	return layout(
		'Terms of Use & End-User License Agreement',
		'Commercial Marketplace Terms for ProWebLook for Microsoft Excel and Microsoft Power Platform.',
		`
		<div class="card">
			<h2>1. License Grant</h2>
			<p>ProWebLook grants you a non-exclusive, worldwide, revocable license to install and use the ProWebLook Microsoft Excel Add-in and Microsoft Power Platform Connector in connection with valid ProWebLook API services.</p>

			<h2>2. Acceptable Use</h2>
			<p>You agree to use ProWebLook verification signals solely for lawful lead qualification, fraud prevention, CRM data hygiene, and customer communications compliance, and not for unlawful stalking, harassment, or FCRA-regulated credit/employment/tenancy eligibility determinations.</p>

			<h2>3. Service Availability &amp; Credits</h2>
			<p>Sample worksheet demonstrations are free of charge. Production API lookups consume credits according to your ProWebLook account plan (<a href="https://proweblook.com/pricing" target="_blank" rel="noopener">proweblook.com/pricing</a>). Duplicate cells within the same Excel session are cached locally at zero extra credit cost.</p>

			<h2>4. Warranty Disclaimer &amp; Support</h2>
			<p>Telecom carrier, WhatsApp, Caller ID, DNSBL, and GeoIP signals are provided on an "as-is" informational basis. Support is available at <a href="mailto:support@proweblook.com">support@proweblook.com</a>.</p>
		</div>
		`,
	);
}
