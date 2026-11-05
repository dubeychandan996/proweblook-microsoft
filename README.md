# ProWebLook for Microsoft 365 — Excel Add-in & Power Platform Connector

Official Microsoft AppSource & Power Platform dual integration for **[ProWebLook](https://proweblook.com)**:

1. **ProWebLook Lead & WhatsApp Verifier for Microsoft Excel (`excel.proweblook.com`)**
   - Native Excel Custom Functions (`=PROWEBLOOK.TRUSTSCORE`, `=PROWEBLOOK.WHATSAPP`, `=PROWEBLOOK.PHONE`, `=PROWEBLOOK.EMAIL`, `=PROWEBLOOK.CALLERID`, `=PROWEBLOOK.DNSBL`, `=PROWEBLOOK.IPGEO`) running on `SharedRuntime 1.1` + `CustomFunctionsRuntime 1.1`.
   - Fluent 2 **Bulk Worksheet Enricher** taskpane with 1-click **Sample Sheet** (`insertSampleLeadsCommand`), automatic column-header detection, conditional color-coded Excel cell formatting (`#DFF6DD` Trusted / `#FFF4CE` Review / `#FDE7E9` High Risk), and automatic session deduplication so duplicate rows cost **0 extra credits**.
   - **Zero-Signup Sample & Starter Mode** compliant with Microsoft Commercial Marketplace Policy `1100.3`.

2. **ProWebLook Connector for Power Automate, Azure Logic Apps & Copilot Studio (`m365.proweblook.com`)**
   - Certified-connector layout (`power-platform-connector/certified-connectors/ProWebLook/`) passing Microsoft's Swagger 2.0, `apiProperties.json`, `230x230` RGB `icon.png`, and schema lint rules.
   - 8 typed operations (`VerifyLeadMultiSignal`, `ValidatePhone`, `CheckWhatsApp`, `VerifyEmail`, `LookupCallerId`, `CheckBlacklist`, `LookupIpGeo`, `GetAccountBalances`) ready for Dynamics 365 Sales, Microsoft Forms, SharePoint, Teams, and autonomous Microsoft Copilot Studio agents.

---

## Live Endpoints

| Surface | URL |
|---|---|
| Excel Add-in Landing & Docs | https://excel.proweblook.com |
| Excel Office Add-in XML Manifest | https://excel.proweblook.com/manifest.xml |
| Excel Taskpane UI | https://excel.proweblook.com/taskpane.html |
| Power Platform Gateway Base | https://m365.proweblook.com/api/connector/v1 |
| Power Platform Swagger 2.0 Spec | https://m365.proweblook.com/swagger.json |
| Support & Setup Guide | https://excel.proweblook.com/support |
| Privacy Policy | https://excel.proweblook.com/privacy |
| Terms of Use / EULA | https://excel.proweblook.com/terms |

---

## Repository Structure

```text
├── excel-addin/
│   ├── manifest.xml               # Office Add-in XML Manifest v1.1 (SharedRuntime + CustomFunctions)
│   ├── functions.json             # 7 =PROWEBLOOK.* Excel custom function definitions
│   ├── functions.js               # CustomFunctions + Ribbon command + deduplicating API client
│   ├── taskpane.html              # Fluent 2 Bulk Enricher, Quick Check & API Key taskpane
│   ├── taskpane.css               # Fluent 2 responsive styling
│   ├── taskpane.js                # Excel.run batch enrichment + column auto-mapper
│   └── assets/                    # 16, 32, 64, 80, 128 ribbon icons + 300x300 AppSource logo
├── power-platform-connector/
│   └── certified-connectors/
│       └── ProWebLook/            # Exact microsoft/PowerPlatformConnectors submission layout
│           ├── apiDefinition.swagger.json
│           ├── apiProperties.json
│           ├── settings.json
│           ├── icon.png           # 230x230 24-bit solid #0F172A PNG
│           └── Readme.md
├── src/
│   ├── worker.js                  # Cloudflare Edge Worker (excel.proweblook.com + m365.proweblook.com)
│   ├── pages.js                   # Landing, Support, Privacy & EULA HTML views
│   └── shared/                    # 6-signal engine, normalizers, Caller ID matcher & sample fixtures
├── scripts/                       # Asset generator, static bundler, Swagger builder & certification validator
└── test/
    └── microsoft-suite.test.mjs   # End-to-end automated test suite
```

---

## Validation & Testing

```bash
npm run build:assets
npm run validate:connector
npm run validate:excel
npm run validate:excel:prod
npm test
```

## License

MIT — Copyright © ProWebLook
