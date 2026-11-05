# ProWebLook

ProWebLook provides real-time 6-signal phone carrier & VoIP detection, silent WhatsApp personal/Business verification, email deliverability & disposable temp-mail screening, Caller ID (CNAM) identity matching, corporate DNSBL reputation checks, Edge IP geolocation, and composite 0–100 Lead & COD Fraud Trust Scoring for Microsoft Power Automate, Azure Logic Apps, Power Apps, and Microsoft Copilot Studio.

## Publisher: ProWebLook

## Prerequisites

- An active ProWebLook account ([sign up free at proweblook.com](https://proweblook.com/auth/signup) — includes 50 free credits per product plus 20,000 monthly Edge GeoIP lookups).

## Supported Operations

### Verify lead and calculate fraud trust score (`VerifyLeadMultiSignal`)
Runs parallel multi-signal verification across phone carrier, silent WhatsApp, email deliverability, Caller ID (CNAM), corporate DNSBL, and Edge GeoIP to compute a composite `0–100` `trustScore` and actionable `riskLevel` (`trusted`, `review`, or `high_risk`).

### Validate phone number and line type (`ValidatePhone`)
Validates national or E.164 phone numbers across 220+ countries and returns operating carrier, line type (`Mobile`, `VoIP`, `Fixed line`), and international formatting.

### Check WhatsApp registration status (`CheckWhatsApp`)
Silently checks whether a phone number has an active WhatsApp Personal or WhatsApp Business account without sending a message.

### Verify email deliverability and disposable status (`VerifyEmail`)
Checks RFC syntax, domain MX records, mailbox existence, corporate vs. free webmail classification, and 10-minute disposable temp-mail providers.

### Look up Caller ID subscriber name (`LookupCallerId`)
Retrieves the registered CNAM Caller ID display name and optionally compares it against `leadName` or `companyName` (`match`, `partial`, or `mismatch`).

### Check domain or IP against DNSBL blocklists (`CheckBlacklist`)
Screens corporate domains or public IP addresses against 50+ real-time spam, phishing, and malware blocklists.

### Look up IP geolocation and VPN status (`LookupIpGeo`)
Resolves IPv4/IPv6 country, city, ASN/ISP, VPN/hosting flags, and billing country mismatch.

### Get ProWebLook credit balances (`GetAccountBalances`)
Returns live remaining credit balances across Phone, WhatsApp, Email, Caller ID, and Blacklist pools.

## Obtaining Credentials

1. Sign in or register at [https://proweblook.com/auth/signup](https://proweblook.com/auth/signup).
2. Navigate to **Dashboard → API Keys** (`https://proweblook.com/user/api_key`).
3. Copy your API key and paste it into the **API Key** connection prompt in Power Automate, Logic Apps, or Copilot Studio.

## Getting Started

1. In Power Automate or Copilot Studio, add a new action and search for **ProWebLook**.
2. Choose **Verify lead and calculate fraud trust score** when a new Dynamics 365 Lead, Microsoft Forms submission, HubSpot contact, or Shopify/WooCommerce order arrives.
3. Branch your flow on `riskLevel` (`trusted` → route to sales / fulfill order; `high_risk` → hold order or flag in Microsoft Teams).

## Known Issues and Limitations

- Caller ID (CNAM) coverage is strongest in North America, India, UK, and major European/MENA mobile networks; numbers without public registry records return `found: false`.
- Private or loopback RFC1918 IP addresses are automatically skipped (`skippedPrivateIp: true`) without consuming quota.

## Deployment Instructions

Run the following command using the Microsoft Power Platform Connectors CLI (`paconn`):

```bash
paconn create --api-def apiDefinition.swagger.json --api-prop apiProperties.json --icon icon.png
```
