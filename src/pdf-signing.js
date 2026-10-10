import { PDFDocument, degrees } from 'pdf-lib';

// Use the same PDF.js viewport transform for preview and export (including rotation/crop).
export function signaturePlacement(viewport, x, y, width, height) {
  const [pdfX, pdfY] = viewport.convertToPdfPoint(x, y + height);
  return { x: pdfX, y: pdfY, width: width / viewport.scale,
    height: height / viewport.scale, rotation: viewport.rotation };
}

export async function signPdf(bytes, placements) {
  const doc = await PDFDocument.load(bytes);
  const images = new Map();
  for (const mark of placements) {
    if (!images.has(mark.image)) images.set(mark.image, await doc.embedPng(mark.image));
    doc.getPage(mark.page - 1).drawImage(images.get(mark.image), {
      x: mark.x, y: mark.y, width: mark.width, height: mark.height,
      rotate: degrees(mark.rotation)
    });
  }
  return doc.save();
}
