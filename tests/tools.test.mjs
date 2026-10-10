import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { PDFDocument, degrees } from 'pdf-lib';
import { PNG } from 'pngjs';
import jsQR from 'jsqr';
import QRCode from 'qrcode';
import { createCanvas } from '@napi-rs/canvas';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { signaturePlacement, signPdf } from '../src/pdf-signing.js';

test('Arabic QR content survives generation and scanning', async () => {
  const value = 'بوابة تعليم جازان https://jazan-edu.github.io/jazan-department-portal/';
  const png = PNG.sync.read(await QRCode.toBuffer(value, { width: 1024, margin: 4, errorCorrectionLevel: 'M' }));
  const decoded = jsQR(new Uint8ClampedArray(png.data), png.width, png.height);
  assert.equal(decoded?.data, value);
});

test('signature export preserves pages and placement on rotated, cropped pages', async () => {
  const original = await PDFDocument.create();
  for (const rotation of [0, 90, 180, 270]) {
    const page = original.addPage([600, 800]); page.setRotation(degrees(rotation)); page.setCropBox(50, 70, 400, 500);
  }
  const bytes = await original.save();
  const image = new PNG({ width: 80, height: 30 });
  for (let i = 0; i < image.data.length; i += 4) { image.data[i] = 220; image.data[i + 3] = 255; }
  const dataUrl = 'data:image/png;base64,' + PNG.sync.write(image).toString('base64');
  const preview = await pdfjs.getDocument({ data: bytes.slice() }).promise;
  const marks = [];
  for (let i = 1; i <= 4; i++) {
    const page = await preview.getPage(i), viewport = page.getViewport({ scale: 1 });
    marks.push({ page: i, image: dataUrl, ...signaturePlacement(viewport, 100, 120, 80, 30) });
  }
  const result = await signPdf(bytes, marks);
  const loaded = await PDFDocument.load(result); assert.equal(loaded.getPageCount(), 4);
  const rendered = await pdfjs.getDocument({ data: result.slice() }).promise;
  for (let i = 1; i <= 4; i++) {
    const page = await rendered.getPage(i), viewport = page.getViewport({ scale: 1 });
    const canvas = createCanvas(viewport.width, viewport.height), context = canvas.getContext('2d');
    await page.render({ canvasContext: context, viewport }).promise;
    const center = context.getImageData(140, 135, 1, 1).data;
    assert.ok(center[0] > 180 && center[1] < 30, `rotation ${page.rotate}: signature center should be red`);
    const outside = context.getImageData(95, 115, 1, 1).data;
    assert.ok(outside[0] > 240 && outside[1] > 240, `rotation ${page.rotate}: outside should remain white`);
  }
  await preview.destroy(); await rendered.destroy();
});

test('portal retains old links and exposes both tools and four external systems', async () => {
  const html = await readFile('index.html', 'utf8');
  const text = html.match(/const PLATFORMS = (\[[\s\S]*?\n\]);/)[1];
  const platforms = Function('return ' + text)();
  assert.equal(platforms.length, 14);
  assert.equal(platforms[5].name, 'منصة استقبال التعديلات على منصة التميز');
  assert.equal(platforms.filter(p => p.tool).length, 2);
  assert.equal(platforms.filter(p => p.external).length, 4);
  assert.ok(platforms.filter(p => p.external).every(p => !p.url.includes('TransactionId')));
});
