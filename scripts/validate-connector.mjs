/**
 * Validates power-platform-connector/certified-connectors/ProWebLook against Microsoft's
 * PowerPlatformConnectors & paconn certification rules.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CONN_DIR = path.resolve(__dirname, '..', 'power-platform-connector', 'certified-connectors', 'ProWebLook');

const errors = [];
function check(cond, msg) {
	if (!cond) errors.push(msg);
}

const ALLOWED_CATEGORIES = new Set([
	'AI',
	'Business Intelligence',
	'Business Management',
	'Collaboration',
	'Commerce',
	'Communication',
	'Content and Files',
	'Data',
	'Finance',
	'Human Resources',
	'Internet of Things',
	'IT Operations',
	'Lifestyle and Entertainment',
	'Marketing',
	'Productivity',
	'Sales and CRM',
	'Security',
	'Social Media',
	'Website',
]);

for (const reqFile of ['apiDefinition.swagger.json', 'apiProperties.json', 'settings.json', 'icon.png', 'Readme.md']) {
	check(fs.existsSync(path.join(CONN_DIR, reqFile)), `Missing required file: ${reqFile}`);
}

const swaggerRaw = fs.readFileSync(path.join(CONN_DIR, 'apiDefinition.swagger.json'), 'utf8');
check(!swaggerRaw.includes('\t'), 'apiDefinition.swagger.json must use 4 spaces, not tabs');
swaggerRaw.split('\n').forEach((line, i) => {
	check(!/\s+$/.test(line), `Trailing whitespace in apiDefinition.swagger.json at line ${i + 1}`);
});

const swagger = JSON.parse(swaggerRaw);
check(swaggerRaw === JSON.stringify(swagger, null, 4) + '\n', 'apiDefinition.swagger.json must be formatted with exact 4-space indentation');
check(swagger.swagger === '2.0', 'swagger version must be "2.0"');
check(Array.isArray(swagger.schemes) && swagger.schemes.length === 1 && swagger.schemes[0] === 'https', 'schemes must be ["https"]');
check(swagger.host === 'm365.proweblook.com', 'host must be m365.proweblook.com');
check(!/connector|\bapi\b/i.test(swagger.info?.title || ''), 'info.title must not contain "Connector" or "API"');
check(String(swagger.info?.description || '').endsWith('.'), 'info.description must end with a period');

const meta = swagger['x-ms-connector-metadata'] || [];
const metaMap = Object.fromEntries(meta.map((m) => [m.propertyName, m.propertyValue]));
check(Boolean(metaMap.Website?.startsWith('https://')), 'x-ms-connector-metadata Website must use https://');
check(Boolean(metaMap['Privacy policy']?.startsWith('https://')), 'x-ms-connector-metadata Privacy policy must use https://');
const cats = String(metaMap.Categories || '').split(';').filter(Boolean);
check(cats.length >= 1 && cats.length <= 2, 'Categories must contain 1 or 2 values separated by ;');
cats.forEach((c) => check(ALLOWED_CATEGORIES.has(c), `Invalid connector category: ${c}`));

const seenOpIds = new Set();
const seenSummaries = new Set();

for (const [routePath, methods] of Object.entries(swagger.paths || {})) {
	for (const [httpMethod, op] of Object.entries(methods)) {
		const ctx = `${httpMethod.toUpperCase()} ${routePath}`;
		check(/^[A-Z][A-Za-z0-9]+$/.test(op.operationId || ''), `${ctx}: operationId "${op.operationId}" must be PascalCase alphanumeric`);
		check(!seenOpIds.has(op.operationId), `${ctx}: duplicate operationId "${op.operationId}"`);
		seenOpIds.add(op.operationId);

		check(Boolean(op.summary) && op.summary.length <= 80 && !op.summary.endsWith('.'), `${ctx}: summary must be <=80 chars without trailing period`);
		check(!seenSummaries.has(op.summary), `${ctx}: duplicate summary "${op.summary}"`);
		seenSummaries.add(op.summary);

		check(Boolean(op.description) && op.description.endsWith('.') && op.description !== op.summary, `${ctx}: description must end with period and differ from summary`);
		check(Boolean(op.responses?.['200']) && Boolean(op.responses?.default), `${ctx}: must define both 200 and default responses`);

		for (const [code, resp] of Object.entries(op.responses || {})) {
			check(Boolean(resp.description) && resp.description.endsWith('.'), `${ctx} response ${code}: description must end with period`);
			check(Boolean(resp.schema?.$ref), `${ctx} response ${code}: must reference a typed schema definition`);
		}
	}
}

for (const [defName, def] of Object.entries(swagger.definitions || {})) {
	check(def.type === 'object', `Definition ${defName} must have type: "object"`);
	check(Boolean(def.description) && def.description.endsWith('.'), `Definition ${defName} must have period-terminated description`);
	for (const [propName, p] of Object.entries(def.properties || {})) {
		if (p.$ref) continue;
		check(Boolean(p.type), `Definition ${defName}.${propName} must specify type`);
		check(Boolean(p.description) && p.description.endsWith('.'), `Definition ${defName}.${propName} description must end with period`);
		check(p.description !== p['x-ms-summary'], `Definition ${defName}.${propName} description must differ from x-ms-summary`);
	}
}

const propsRaw = fs.readFileSync(path.join(CONN_DIR, 'apiProperties.json'), 'utf8');
const props = JSON.parse(propsRaw);
check(propsRaw === JSON.stringify(props, null, 4) + '\n', 'apiProperties.json must be formatted with 4-space indentation');
check(/^#[0-9A-Fa-f]{6}$/.test(props.properties?.iconBrandColor || '') && props.properties.iconBrandColor.toLowerCase() !== '#007ee5', 'iconBrandColor must be a custom 6-digit hex color');

const iconBuf = fs.readFileSync(path.join(CONN_DIR, 'icon.png'));
check(iconBuf.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])), 'icon.png must have valid PNG header');
const width = iconBuf.readUInt32BE(16);
const height = iconBuf.readUInt32BE(20);
const colorType = iconBuf[25];
check(width === 230 && height === 230, `icon.png must be 230x230 (got ${width}x${height})`);
check(colorType === 2, `icon.png must be non-transparent 24-bit RGB colorType=2 (got ${colorType})`);

const readme = fs.readFileSync(path.join(CONN_DIR, 'Readme.md'), 'utf8');
for (const heading of [
	'## Publisher:',
	'## Prerequisites',
	'## Supported Operations',
	'## Obtaining Credentials',
	'## Getting Started',
	'## Known Issues and Limitations',
	'## Deployment Instructions',
]) {
	check(readme.includes(heading), `Readme.md missing required section "${heading}"`);
}

if (errors.length > 0) {
	console.error('Connector Certification Validation FAILED:');
	errors.forEach((e) => console.error(' - ' + e));
	process.exit(1);
}

console.log(`Connector Certification Validation PASSED: ${seenOpIds.size} operations, ${Object.keys(swagger.definitions).length} schemas, 230x230 RGB icon.png.`);
