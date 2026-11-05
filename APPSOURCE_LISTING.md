# Microsoft AppSource Listing Metadata — ProWebLook

## Offer 1: Office Store (Excel Add-in)
- **Offer Alias / ID:** `proweblook-excel-verifier`
- **Manifest DisplayName (exact match):** `ProWebLook Lead and WhatsApp Verifier`
- **Manifest Short Description (exact match, <=100 chars):** `Bulk verify Phone Carrier, Silent WhatsApp, Caller ID, Email, DNSBL & GeoIP in Microsoft Excel.`
- **Support URL:** `https://excel.proweblook.com/support`
- **Privacy Policy URL:** `https://excel.proweblook.com/privacy`
- **End User License Agreement (EULA) URL:** `https://excel.proweblook.com/terms`
- **Store Logo (300x300 PNG):** `dist/store-logo-300.png`
- **Manifest File:** `dist/proweblook-excel-manifest.xml`
- **Pricing:** Free Add-in (Marked: "My add-in may require purchase of a service or offers additional paid functionality")
- **Categories:** CRM, Sales & Marketing, Productivity
- **Search Keywords:** `WhatsApp Verifier`, `Phone Validation`, `Email Verification`

### Long Description
Stop wasting sales outreach, ad spend, and Cash-on-Delivery (COD) shipping fees on fake phone numbers, burner VoIP lines, and 10-minute disposable emails.

**ProWebLook Lead & WhatsApp Verifier** brings real-time 6-signal telecom, messaging, email, identity, and network fraud verification directly into **Microsoft Excel** (Windows, Mac, and Excel on the Web).

**Key Capabilities**
- **1-Click Bulk Worksheet Enrichment:** Automatically map Phone, Email, Lead Name, Country, and IP columns, verify rows in parallel, and write color-coded **Trust Scores (0–100)**, **Risk Levels** (`TRUSTED`, `REVIEW`, `HIGH RISK`), **Carrier**, **Line Type**, **WhatsApp**, **Caller ID**, **Email Verdict**, **Domain DNSBL**, and **GeoIP** columns directly to your sheet.
- **7 Native Excel Custom Formulas:**
  - `=PROWEBLOOK.TRUSTSCORE(phone, email, country, lead_name, ip)` — Composite 0–100 Lead & COD Fraud Trust Score
  - `=PROWEBLOOK.WHATSAPP(phone, country, field)` — Silent WhatsApp Personal vs. Business check without sending messages
  - `=PROWEBLOOK.PHONE(phone, country, field)` — Carrier lookup, VoIP/Mobile/Landline detection & E.164 formatting across 220+ countries
  - `=PROWEBLOOK.EMAIL(email, field)` — MX/mailbox deliverability + disposable temp-mail blocking
  - `=PROWEBLOOK.CALLERID(phone, lead_name, country)` — Registered CNAM subscriber lookup with fuzzy lead-name matching
  - `=PROWEBLOOK.DNSBL(domain_or_ip, field)` — Screens corporate domains & IPs across 50+ spam & malware blocklists
  - `=PROWEBLOOK.IPGEO(ip, country_hint)` — Sub-millisecond Edge GeoIP, ASN/ISP, VPN/proxy detection & billing country mismatch
- **Duplicate-Row Credit Saver:** Repeated phone numbers, emails, or domains in the same Excel session are deduplicated automatically at zero extra credit cost.

**Instant Zero-Signup Trial:** Click **Insert Sample Sheet** in the ribbon or taskpane to test bulk enrichment and all `=PROWEBLOOK.*` formulas immediately with zero account required. Connect a free ProWebLook API key (`proweblook.com/auth/signup`) for 50 free credits per product plus 20,000 monthly Edge GeoIP lookups.

### Notes for Microsoft Certification Reviewers
No account, login, or paid subscription is required to test every feature of this add-in:
1. On the Excel **Home** tab, click **Verify Leads** inside the **ProWebLook** group to open the taskpane.
2. Click **Insert Sample Sheet** (or click the **Sample Sheet** ribbon button) to create the 6-row `ProWebLook Sample` worksheet with realistic Trusted, Review, and High-Risk B2B/COD test leads.
3. Click **Verify & Enrich Sheet** to run full 6-signal bulk verification and observe color-coded Trust Score & Risk Level columns written into the sheet.
4. Switch to the **Quick & Formulas** tab and click **Try Sample** to test single-lead lookup, or paste `=PROWEBLOOK.TRUSTSCORE(B2, C2, D2, A2, E2)` and `=PROWEBLOOK.WHATSAPP(B2, D2)` into any cell on the sample sheet.
