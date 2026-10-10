import * as pdfjs from 'pdfjs-dist/build/pdf.mjs';
import { PDFDocument } from 'pdf-lib';
import QRCode from 'qrcode';
import { signaturePlacement, signPdf } from './pdf-signing.js';

pdfjs.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.mjs', import.meta.url).href;
const $ = id => document.getElementById(id);
const status = (message, error = false) => { $('status').textContent = message; $('status').classList.toggle('error', error); };
const download = (blob, name) => {
  const url = URL.createObjectURL(blob), a = document.createElement('a');
  a.href = url; a.download = name; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
};
const mode = new URLSearchParams(location.search).get('tool') === 'qr' ? 'qr' : 'sign';
$(mode === 'qr' ? 'qrTool' : 'signTool').hidden = false;
$('toolTitle').textContent = mode === 'qr' ? 'إنشاء QR Code' : 'توقيع ملفات PDF';
document.title = $('toolTitle').textContent + ' — أدوات القسم';
status(mode === 'qr' ? 'أدخل الرابط أو النص لإنشاء الرمز.' : 'ابدأ باختيار ملف PDF وتجهيز توقيعك.');

let pdfBytes, pdfDocument, viewport, pageNumber = 1, filename = 'document.pdf';
let marks = [], signature = null, busy = false, drawing = false, hasInk = false;
let selectedMark = null;
const pad = $('signaturePad'), pen = pad.getContext('2d', { willReadFrequently: true });
function controls() {
  const ready = !!pdfDocument && !busy;
  $('pdfInput').disabled = busy; $('pageSelect').disabled = !ready;
  $('prevPage').disabled = !ready || pageNumber <= 1;
  $('nextPage').disabled = !ready || pageNumber >= pdfDocument.numPages;
  $('placeSignature').disabled = !ready || !signature;
  $('moveSignature').disabled = !ready || !marks.includes(selectedMark) || selectedMark.page !== pageNumber;
  $('undoSignature').disabled = busy || !marks.length;
  $('clearSignatures').disabled = busy || !marks.length;
  $('savePdf').disabled = busy || !marks.length;
  $('pageSurface').classList.toggle('busy', busy);
  $('signatureCount').textContent = marks.length ? `${marks.length} توقيع على الملف.` : 'لم تُضف أي توقيع بعد.';
}
function drawMarks() {
  $('signatureLayer').replaceChildren();
  for (const mark of marks.filter(m => m.page === pageNumber)) {
    const img = document.createElement('img'); img.src = mark.image; img.alt = 'توقيع مضاف؛ اسحب لتغيير مكانه';
    img.draggable = false; img.tabIndex = 0; img.setAttribute('role', 'button');
    img.classList.toggle('selected', mark === selectedMark);
    img.setAttribute('aria-pressed', String(mark === selectedMark));
    img.onclick = event => { event.stopPropagation(); selectMark(mark); };
    img.onkeydown = event => {
      if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); selectMark(mark); }
    };
    img.onpointerdown = event => {
      if (busy || event.button !== 0) return;
      event.preventDefault(); event.stopPropagation();
      selectedMark = mark; syncPosition();
      $('signatureLayer').querySelectorAll('img').forEach(el => { el.classList.toggle('selected', el === img); el.setAttribute('aria-pressed', String(el === img)); });
      controls();
      const rect = $('pageSurface').getBoundingClientRect();
      const startX = event.clientX, startY = event.clientY, left = mark.left, top = mark.top;
      img.setPointerCapture(event.pointerId);
      img.onpointermove = move => {
        if (!img.hasPointerCapture(move.pointerId)) return;
        move.preventDefault();
        moveMark(mark, left + (move.clientX - startX) / rect.width, top + (move.clientY - startY) / rect.height);
        img.style.left = mark.left * 100 + '%'; img.style.top = mark.top * 100 + '%'; syncPosition();
      };
      const finish = () => { img.onpointermove = null; status('تم تعديل مكان التوقيع. يمكنك سحبه مجددًا أو تنزيل الملف.'); };
      img.onpointerup = finish; img.onpointercancel = finish; img.onlostpointercapture = finish;
    };
    Object.assign(img.style, { left: mark.left * 100 + '%', top: mark.top * 100 + '%',
      width: mark.relativeWidth * 100 + '%', height: mark.relativeHeight * 100 + '%' });
    $('signatureLayer').append(img);
  }
  controls();
}
function syncPosition() {
  if (!selectedMark) return;
  $('positionX').value = ((selectedMark.left + selectedMark.relativeWidth / 2) * 100).toFixed(1);
  $('positionY').value = ((selectedMark.top + selectedMark.relativeHeight / 2) * 100).toFixed(1);
}
function selectMark(mark) { selectedMark = mark; syncPosition(); drawMarks(); status('التوقيع محدد. اسحبه أو عدّل حقول الموضع ثم اضغط «تعديل مكان التوقيع المحدد».'); }
function moveMark(mark, left, top) {
  mark.left = Math.max(0, Math.min(1 - mark.relativeWidth, left));
  mark.top = Math.max(0, Math.min(1 - mark.relativeHeight, top));
  Object.assign(mark, signaturePlacement(viewport, mark.left * viewport.width, mark.top * viewport.height,
    mark.relativeWidth * viewport.width, mark.relativeHeight * viewport.height));
}
async function renderPage() {
  busy = true; controls();
  try {
    const page = await pdfDocument.getPage(pageNumber);
    const base = page.getViewport({ scale: 1 });
    viewport = page.getViewport({ scale: Math.min(1.6, 1400 / Math.max(base.width, base.height)) });
    const canvas = $('pdfCanvas'); canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    $('pageSelect').value = String(pageNumber); $('pdfEmpty').hidden = true; $('pdfScroll').hidden = false;
    $('pageSurface').style.maxWidth = viewport.width + 'px'; drawMarks();
  } catch (error) { status('تعذّر عرض هذه الصفحة. جرّب ملف PDF آخر.', true); }
  finally { busy = false; controls(); }
}
$('pdfInput').addEventListener('change', async event => {
  const file = event.target.files[0]; if (!file) return;
  if (file.size > 30 * 1024 * 1024) { status('حجم الملف يتجاوز 30 ميجابايت.', true); event.target.value = ''; return; }
  busy = true; controls(); status('جاري قراءة ملف PDF…');
  try {
    const data = new Uint8Array(await file.arrayBuffer());
    await PDFDocument.load(data); // Reject encrypted/invalid files before replacing the current document.
    const next = await pdfjs.getDocument({ data: data.slice(), isEvalSupported: false,
      cMapUrl: new URL('./cmaps/', import.meta.url).href, cMapPacked: true,
      standardFontDataUrl: new URL('./standard_fonts/', import.meta.url).href,
      wasmUrl: new URL('./wasm/', import.meta.url).href }).promise;
    if (pdfDocument) await pdfDocument.destroy();
    pdfBytes = data; pdfDocument = next; pageNumber = 1; filename = file.name; marks = []; selectedMark = null;
    $('pageSelect').replaceChildren();
    for (let i = 1; i <= next.numPages; i++) $('pageSelect').add(new Option(String(i), String(i)));
    $('fileInfo').textContent = `${file.name} — ${next.numPages} صفحة`;
    await renderPage(); status('الملف جاهز. اعتمد توقيعك ثم اضغط على موضعه في المعاينة.');
  } catch (error) { status('تعذّر فتح الملف. اختر ملف PDF سليمًا وغير محمي بكلمة مرور.', true); }
  finally { event.target.value = ''; busy = false; controls(); }
});
async function goPage(number) { if (busy || !pdfDocument) return; selectedMark = null; pageNumber = number; await renderPage(); }
$('pageSelect').onchange = event => goPage(Number(event.target.value));
$('prevPage').onclick = () => goPage(pageNumber - 1);
$('nextPage').onclick = () => goPage(pageNumber + 1);

function point(event) { const rect = pad.getBoundingClientRect(); return [(event.clientX - rect.left) * pad.width / rect.width, (event.clientY - rect.top) * pad.height / rect.height]; }
pad.addEventListener('pointerdown', event => {
  if (event.button !== 0) return;
  event.preventDefault(); drawing = true; hasInk = true; pad.setPointerCapture(event.pointerId);
  const [x, y] = point(event); pen.strokeStyle = $('inkColor').value; pen.fillStyle = pen.strokeStyle;
  pen.lineWidth = 3.5; pen.lineCap = 'round'; pen.lineJoin = 'round';
  pen.beginPath(); pen.arc(x, y, 1.75, 0, Math.PI * 2); pen.fill(); pen.beginPath(); pen.moveTo(x, y);
});
pad.addEventListener('pointermove', event => { if (!drawing) return; event.preventDefault(); pen.lineTo(...point(event)); pen.stroke(); });
['pointerup', 'pointercancel', 'lostpointercapture'].forEach(type => pad.addEventListener(type, () => { drawing = false; }));
$('clearPad').onclick = () => { pen.clearRect(0, 0, pad.width, pad.height); hasInk = false; status('تم مسح الرسم. اعتمد الرسم الجديد لاستبدال التوقيع المستخدم.'); };
function acceptSignature(canvas) {
  signature = { image: canvas.toDataURL('image/png'), aspect: canvas.width / canvas.height };
  $('signaturePreview').src = signature.image; $('signatureReady').hidden = false; controls();
  status('التوقيع جاهز. اضغط على موضعه في صفحة PDF أو استخدم حقول الموضع.');
}
$('useDrawing').onclick = () => {
  if (!hasInk) { status('ارسم توقيعك أولًا داخل المساحة المخصصة.', true); return; }
  const data = pen.getImageData(0, 0, pad.width, pad.height).data;
  let minX = pad.width, minY = pad.height, maxX = 0, maxY = 0;
  for (let y = 0; y < pad.height; y++) for (let x = 0; x < pad.width; x++) if (data[(y * pad.width + x) * 4 + 3]) {
    minX = Math.min(x, minX); maxX = Math.max(x, maxX); minY = Math.min(y, minY); maxY = Math.max(y, maxY);
  }
  const cropped = document.createElement('canvas'); cropped.width = maxX - minX + 9; cropped.height = maxY - minY + 9;
  cropped.getContext('2d').drawImage(pad, minX, minY, maxX - minX + 1, maxY - minY + 1, 4, 4, maxX - minX + 1, maxY - minY + 1);
  acceptSignature(cropped);
};
$('signatureInput').onchange = async event => {
  const file = event.target.files[0]; if (!file) return;
  if (!['image/png', 'image/jpeg'].includes(file.type) || file.size > 5 * 1024 * 1024) {
    status('اختر صورة PNG أو JPG لا يتجاوز حجمها 5 ميجابايت.', true); event.target.value = ''; return;
  }
  try {
    const bitmap = await createImageBitmap(file);
    if (bitmap.width * bitmap.height > 25000000) { bitmap.close(); throw Error('Image too large'); }
    const canvas = document.createElement('canvas'); const scale = Math.min(1, 1600 / Math.max(bitmap.width, bitmap.height));
    canvas.width = Math.max(1, Math.round(bitmap.width * scale)); canvas.height = Math.max(1, Math.round(bitmap.height * scale));
    canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height); bitmap.close(); acceptSignature(canvas);
  } catch { status('تعذّر قراءة الصورة. جرّب صورة PNG أو JPG أصغر.', true); }
  event.target.value = '';
};
function signatureMode(mode) {
  $('drawMode').hidden = mode !== 'draw'; $('uploadMode').hidden = mode !== 'upload';
  $('drawTab').setAttribute('aria-pressed', String(mode === 'draw')); $('uploadTab').setAttribute('aria-pressed', String(mode === 'upload'));
}
$('drawTab').onclick = () => signatureMode('draw'); $('uploadTab').onclick = () => signatureMode('upload');
$('signatureSize').oninput = event => { $('sizeValue').textContent = event.target.value + '%'; };
function place(xFraction, yFraction) {
  if (busy || !pdfDocument || !signature || !viewport) { status('اختر ملف PDF واعتمد توقيعًا أولًا.', true); return; }
  let width = viewport.width * Number($('signatureSize').value) / 100, height = width / signature.aspect;
  if (height > viewport.height * .6) { width *= viewport.height * .6 / height; height = viewport.height * .6; }
  const x = Math.max(0, Math.min(viewport.width - width, viewport.width * xFraction - width / 2));
  const y = Math.max(0, Math.min(viewport.height - height, viewport.height * yFraction - height / 2));
  marks.push({ page: pageNumber, image: signature.image, ...signaturePlacement(viewport, x, y, width, height),
    left: x / viewport.width, top: y / viewport.height, relativeWidth: width / viewport.width, relativeHeight: height / viewport.height });
  selectedMark = marks.at(-1); syncPosition();
  drawMarks(); status(`تمت إضافة التوقيع إلى الصفحة ${pageNumber}. اسحبه لتعديل مكانه أو نزّل الملف.`);
}
$('pageSurface').onclick = event => { const rect = event.currentTarget.getBoundingClientRect(); place((event.clientX - rect.left) / rect.width, (event.clientY - rect.top) / rect.height); };
$('placeSignature').onclick = () => {
  const x = Number($('positionX').value), y = Number($('positionY').value);
  if (!$('positionX').value || !$('positionY').value || !Number.isFinite(x + y) || x < 0 || x > 100 || y < 0 || y > 100) { status('أدخل موضعًا بين 0 و100 في الحقلين.', true); return; }
  place(x / 100, y / 100);
};
$('moveSignature').onclick = () => {
  if (busy || !marks.includes(selectedMark) || selectedMark.page !== pageNumber) return;
  const x = Number($('positionX').value), y = Number($('positionY').value);
  if (!$('positionX').value || !$('positionY').value || !Number.isFinite(x + y) || x < 0 || x > 100 || y < 0 || y > 100) { status('أدخل موضعًا بين 0 و100 في الحقلين.', true); return; }
  moveMark(selectedMark, x / 100 - selectedMark.relativeWidth / 2, y / 100 - selectedMark.relativeHeight / 2);
  syncPosition(); drawMarks(); status('تم تعديل مكان التوقيع المحدد.');
};
$('undoSignature').onclick = () => { marks.pop(); if (!marks.includes(selectedMark)) selectedMark = null; drawMarks(); status('تم التراجع عن آخر توقيع.'); };
$('clearSignatures').onclick = () => { marks = []; selectedMark = null; drawMarks(); status('تم حذف التواقيع المضافة من المعاينة.'); };
$('savePdf').onclick = async () => {
  if (busy || !marks.length) return; busy = true; controls(); status('جاري إعداد الملف الموقّع…');
  try { const bytes = await signPdf(pdfBytes, marks); download(new Blob([bytes], { type: 'application/pdf' }), filename.replace(/\.pdf$/i, '') + '-signed.pdf'); status('تم تجهيز الملف الموقّع وتنزيله.'); }
  catch { status('تعذّر حفظ الملف. جرّب ملفًا آخر أو قلّل عدد التواقيع.', true); }
  finally { busy = false; controls(); }
};
controls();

let qrSvg = null;
function invalidateQr() { qrSvg = null; $('downloadQr').disabled = true; $('downloadSvg').disabled = true; $('qrCanvas').hidden = true; $('qrEmpty').hidden = false; }
$('qrText').oninput = invalidateQr; $('qrSize').onchange = invalidateQr;
$('generateQr').onclick = async () => {
  const text = $('qrText').value.trim(); invalidateQr();
  if (!text) { status('اكتب رابطًا أو نصًا لإنشاء الرمز.', true); return; }
  const options = { errorCorrectionLevel: 'M', margin: 4, width: Number($('qrSize').value), color: { dark: '#0d2724ff', light: '#ffffffff' } };
  $('generateQr').disabled = true;
  try {
    await QRCode.toCanvas($('qrCanvas'), text, options);
    const svg = await QRCode.toString(text, { ...options, type: 'svg' });
    if ($('qrText').value.trim() !== text || Number($('qrSize').value) !== options.width) return;
    qrSvg = svg; $('qrCanvas').hidden = false; $('qrEmpty').hidden = true;
    $('downloadQr').disabled = false; $('downloadSvg').disabled = false; status('الرمز جاهز. يمكنك تنزيله بصيغة PNG أو SVG.');
  } catch { status('تعذّر إنشاء الرمز. قلّل طول النص ثم حاول مجددًا.', true); }
  finally { $('generateQr').disabled = false; }
};
$('downloadQr').onclick = () => { if (qrSvg) $('qrCanvas').toBlob(blob => { if (blob) download(blob, 'qrcode.png'); }, 'image/png'); };
$('downloadSvg').onclick = () => { if (qrSvg) download(new Blob([qrSvg], { type: 'image/svg+xml;charset=utf-8' }), 'qrcode.svg'); };
