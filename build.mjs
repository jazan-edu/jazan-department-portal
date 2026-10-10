import { build } from 'esbuild';
import { cp, copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
await mkdir('assets', { recursive: true });
await build({ entryPoints: ['src/tools.js'], bundle: true, minify: true, format: 'esm', target: 'es2022', outfile: 'assets/tools.js', legalComments: 'external' });
await copyFile('node_modules/pdfjs-dist/build/pdf.worker.min.mjs', 'assets/pdf.worker.mjs');
for (const dir of ['cmaps', 'standard_fonts', 'wasm']) await cp(`node_modules/pdfjs-dist/${dir}`, `assets/${dir}`, { recursive: true });
let notices = '';
for (const [name, license] of [['pdfjs-dist','LICENSE'], ['pdf-lib','LICENSE.md'], ['qrcode','license']]) {
  notices += `\n--- ${name} ---\n` + await readFile(`node_modules/${name}/${license}`, 'utf8');
}
await writeFile('assets/THIRD-PARTY-LICENSES.txt', notices);
