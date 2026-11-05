/**
 * Generates supersampled PNG icons for:
 *  1) Microsoft Excel Office Add-in (`excel-addin/assets/icon-{16,32,64,80,128}.png` + `store-logo-300.png`)
 *  2) Microsoft Power Platform Certified Connector (`power-platform-connector/certified-connectors/ProWebLook/icon.png` — 230x230 solid #0F172A RGB PNG)
 *  3) Embedded base64 map (`src/generated-icons.js`) for sub-millisecond Cloudflare Edge Worker delivery.
 */

import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

function crc32(buf) {
	let crc = 0xffffffff;
	for (let i = 0; i < buf.length; i++) {
		crc ^= buf[i];
		for (let j = 0; j < 8; j++) {
			crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
		}
	}
	return (crc ^ 0xffffffff) >>> 0;
}

function makeChunk(type, data) {
	const typeBuf = Buffer.from(type, 'ascii');
	const lenBuf = Buffer.alloc(4);
	lenBuf.writeUInt32BE(data.length, 0);
	const crcBuf = Buffer.alloc(4);
	crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
	return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
}

function encodeRgbaPng(width, height, rgba) {
	const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
	const ihdr = Buffer.alloc(13);
	ihdr.writeUInt32BE(width, 0);
	ihdr.writeUInt32BE(height, 4);
	ihdr[8] = 8;
	ihdr[9] = 6; // RGBA
	const stride = width * 4;
	const raw = Buffer.alloc((stride + 1) * height);
	for (let y = 0; y < height; y++) {
		raw[y * (stride + 1)] = 0;
		rgba.copy(raw, y * (stride + 1) + 1, y * stride, (y + 1) * stride);
	}
	const idat = zlib.deflateSync(raw, { level: 9 });
	return Buffer.concat([sig, makeChunk('IHDR', ihdr), makeChunk('IDAT', idat), makeChunk('IEND', Buffer.alloc(0))]);
}

function encodeRgb24Png(width, height, rgb) {
	const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
	const ihdr = Buffer.alloc(13);
	ihdr.writeUInt32BE(width, 0);
	ihdr.writeUInt32BE(height, 4);
	ihdr[8] = 8;
	ihdr[9] = 2; // Truecolor 24-bit RGB (no alpha channel)
	const stride = width * 3;
	const raw = Buffer.alloc((stride + 1) * height);
	for (let y = 0; y < height; y++) {
		raw[y * (stride + 1)] = 0;
		rgb.copy(raw, y * (stride + 1) + 1, y * stride, (y + 1) * stride);
	}
	const idat = zlib.deflateSync(raw, { level: 9 });
	return Buffer.concat([sig, makeChunk('IHDR', ihdr), makeChunk('IDAT', idat), makeChunk('IEND', Buffer.alloc(0))]);
}

function distToSegment(px, py, x1, y1, x2, y2) {
	const dx = x2 - x1;
	const dy = y2 - y1;
	const lenSq = dx * dx + dy * dy;
	if (lenSq === 0) return Math.hypot(px - x1, py - y1);
	const t = Math.max(0, Math.min(1, ((px - x1) * dx + (py - y1) * dy) / lenSq));
	return Math.hypot(px - (x1 + t * dx), py - (y1 + t * dy));
}

function roundedRectSdf(nx, ny, halfSize, radius) {
	const qx = Math.abs(nx) - (halfSize - radius);
	const qy = Math.abs(ny) - (halfSize - radius);
	return Math.hypot(Math.max(qx, 0), Math.max(qy, 0)) + Math.min(Math.max(qx, qy), 0) - radius;
}

function sampleShieldMark(nx, ny, baseR, baseG, baseB) {
	let r = baseR;
	let g = baseG;
	let b = baseB;

	const sy = ny + 0.03;
	const shieldTop = -0.58;
	const shieldBottom = 0.64;
	let inShield = false;
	if (sy >= shieldTop && sy <= shieldBottom) {
		const tY = (sy - shieldTop) / (shieldBottom - shieldTop);
		const maxW = tY < 0.42 ? 0.54 : 0.54 * Math.pow(Math.max(0, 1 - (tY - 0.42) / 0.58), 0.72);
		const topCurve = shieldTop + Math.abs(nx) * 0.16;
		if (sy >= topCurve && Math.abs(nx) <= maxW) {
			inShield = true;
		}
	}

	if (inShield) {
		const st = Math.max(0, Math.min(1, (nx + sy + 1.1) / 2.2));
		if (st < 0.55) {
			const k = st / 0.55;
			r = Math.round(56 * (1 - k) + 15 * k);
			g = Math.round(189 * (1 - k) + 108 * k);
			b = Math.round(248 * (1 - k) + 189 * k);
		} else {
			const k = (st - 0.55) / 0.45;
			r = Math.round(15 * (1 - k) + 16 * k);
			g = Math.round(108 * (1 - k) + 185 * k);
			b = Math.round(189 * (1 - k) + 129 * k);
		}
	}

	const d1 = distToSegment(nx, ny, -0.24, 0.03, -0.04, 0.23);
	const d2 = distToSegment(nx, ny, -0.04, 0.23, 0.29, -0.14);
	const dCheck = Math.min(d1, d2);
	const strokeRadius = 0.088;
	if (dCheck < strokeRadius + 0.025) {
		const w = Math.max(0, Math.min(1, (strokeRadius + 0.015 - dCheck) / 0.03));
		r = Math.round(r * (1 - w) + 255 * w);
		g = Math.round(g * (1 - w) + 255 * w);
		b = Math.round(b * (1 - w) + 255 * w);
	}

	return [r, g, b];
}

function renderRgbaSquircleIcon(size) {
	const ss = 4;
	const out = Buffer.alloc(size * size * 4);
	for (let y = 0; y < size; y++) {
		for (let x = 0; x < size; x++) {
			let rSum = 0;
			let gSum = 0;
			let bSum = 0;
			let aSum = 0;
			for (let sy = 0; sy < ss; sy++) {
				for (let sx = 0; sx < ss; sx++) {
					const u = (x + (sx + 0.5) / ss) / size;
					const v = (y + (sy + 0.5) / ss) / size;
					const nx = (u * 2 - 1) / 0.9;
					const ny = (v * 2 - 1) / 0.9;
					const dBox = roundedRectSdf(nx, ny, 0.96, 0.26);
					if (dBox > 0.02) continue;

					const alphaBox = Math.max(0, Math.min(1, (0.015 - dBox) / 0.03));
					const gradT = Math.max(0, Math.min(1, (nx + ny + 2) / 4));
					let r = Math.round(15 + gradT * 20);
					let g = Math.round(23 + gradT * 32);
					let b = Math.round(42 + gradT * 96);
					if (dBox > -0.055 && dBox <= 0) {
						r = Math.min(255, r + 35);
						g = Math.min(255, g + 60);
						b = Math.min(255, b + 95);
					}
					[r, g, b] = sampleShieldMark(nx, ny, r, g, b);

					rSum += r * alphaBox;
					gSum += g * alphaBox;
					bSum += b * alphaBox;
					aSum += alphaBox;
				}
			}
			const samples = ss * ss;
			const finalAlpha = aSum / samples;
			const idx = (y * size + x) * 4;
			if (finalAlpha <= 0.001) {
				out[idx] = 0;
				out[idx + 1] = 0;
				out[idx + 2] = 0;
				out[idx + 3] = 0;
			} else {
				out[idx] = Math.min(255, Math.round(rSum / aSum));
				out[idx + 1] = Math.min(255, Math.round(gSum / aSum));
				out[idx + 2] = Math.min(255, Math.round(bSum / aSum));
				out[idx + 3] = Math.min(255, Math.round(finalAlpha * 255));
			}
		}
	}
	return encodeRgbaPng(size, size, out);
}

/**
 * Renders a solid #0F172A (15, 23, 42) 24-bit RGB square icon (required for Power Platform Connector & AppSource Store Logo).
 */
function renderSolidBrandSquarePng(size) {
	const ss = 4;
	const out = Buffer.alloc(size * size * 3);
	for (let y = 0; y < size; y++) {
		for (let x = 0; x < size; x++) {
			let rSum = 0;
			let gSum = 0;
			let bSum = 0;
			for (let sy = 0; sy < ss; sy++) {
				for (let sx = 0; sx < ss; sx++) {
					const u = (x + (sx + 0.5) / ss) / size;
					const v = (y + (sy + 0.5) / ss) / size;
					const nx = (u * 2 - 1) / 0.78;
					const ny = (v * 2 - 1) / 0.78;
					// Solid #0F172A background matching apiProperties.json iconBrandColor
					const [r, g, b] = sampleShieldMark(nx, ny, 15, 23, 42);
					rSum += r;
					gSum += g;
					bSum += b;
				}
			}
			const samples = ss * ss;
			const idx = (y * size + x) * 3;
			out[idx] = Math.round(rSum / samples);
			out[idx + 1] = Math.round(gSum / samples);
			out[idx + 2] = Math.round(bSum / samples);
		}
	}
	return encodeRgb24Png(size, size, out);
}

const excelAssetsDir = path.join(ROOT, 'excel-addin', 'assets');
const connectorDir = path.join(ROOT, 'power-platform-connector', 'certified-connectors', 'ProWebLook');
fs.mkdirSync(excelAssetsDir, { recursive: true });
fs.mkdirSync(connectorDir, { recursive: true });

const iconMap = {};
for (const sz of [16, 32, 64, 80, 128]) {
	const buf = renderRgbaSquircleIcon(sz);
	const filename = `icon-${sz}.png`;
	fs.writeFileSync(path.join(excelAssetsDir, filename), buf);
	iconMap[filename] = buf.toString('base64');
}

const logo300 = renderSolidBrandSquarePng(300);
fs.writeFileSync(path.join(excelAssetsDir, 'store-logo-300.png'), logo300);
iconMap['store-logo-300.png'] = logo300.toString('base64');

const connectorIcon230 = renderSolidBrandSquarePng(230);
fs.writeFileSync(path.join(connectorDir, 'icon.png'), connectorIcon230);
iconMap['connector-icon-230.png'] = connectorIcon230.toString('base64');

const generatedJs = `// Auto-generated by scripts/generate-assets.mjs — do not edit manually.\nexport const EMBEDDED_PNG_BASE64 = ${JSON.stringify(iconMap, null, '\t')};\n`;
fs.writeFileSync(path.join(ROOT, 'src', 'generated-icons.js'), generatedJs);

console.log('Generated Excel Add-in icons (16, 32, 64, 80, 128, 300) + Power Platform 230x230 solid icon.png');
