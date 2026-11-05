/**
 * Packages clean AppSource & Power Platform submission artifacts into dist/:
 *  - dist/proweblook-excel-manifest.xml
 *  - dist/store-logo-300.png
 *  - dist/ProWebLook-PowerPlatform-Connector-v1.0.0.zip
 */

import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const DIST = path.join(ROOT, 'dist');

fs.mkdirSync(DIST, { recursive: true });
fs.copyFileSync(path.join(ROOT, 'excel-addin', 'manifest.xml'), path.join(DIST, 'proweblook-excel-manifest.xml'));
fs.copyFileSync(path.join(ROOT, 'excel-addin', 'assets', 'store-logo-300.png'), path.join(DIST, 'store-logo-300.png'));

const zipPath = path.join(DIST, 'ProWebLook-PowerPlatform-Connector-v1.0.0.zip');
if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);

execFileSync('zip', ['-r', '-q', zipPath, 'ProWebLook'], {
	cwd: path.join(ROOT, 'power-platform-connector', 'certified-connectors'),
});

console.log('Packaged dist/ artifacts:', fs.readdirSync(DIST).join(', '));
