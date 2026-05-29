/**
 * word-to-pdf.ts — Convert .docx to PDF via server-side Gotenberg/LibreOffice.
 *
 * Sends the .docx file to your Cloudflare Worker, which proxies it to
 * Gotenberg running LibreOffice on Railway. The result is a perfect PDF
 * — identical to "Save as PDF" in Microsoft Word.
 */

import { PDFDocument, degrees } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

// Your deployed Cloudflare Worker URL
const CONVERT_API = 'https://pdf-convert.2710476780.workers.dev';

export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  // Send .docx to server for LibreOffice conversion
  const formData = new FormData();
  formData.append(
    'file',
    new Blob([sourceArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
    'document.docx'
  );

  let response: Response;
  try {
    response = await fetch(CONVERT_API, {
      method: 'POST',
      body: formData,
    });
  } catch (err) {
    throw new Error(
      '转换服务连接失败，请检查网络后重试'
    );
  }

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(
      `Word 转换失败（${response.status}）${text ? ': ' + text : ''}`
    );
  }

  // Load the converted PDF and copy its pages into the merged document
  const pdfBytes = new Uint8Array(await response.arrayBuffer());
  const sourcePdf = await PDFDocument.load(pdfBytes);
  const pageIndices = sourcePdf.getPageIndices();
  const copiedPages = await mergedPdf.copyPages(sourcePdf, pageIndices);

  for (let i = 0; i < copiedPages.length; i++) {
    mergedPdf.addPage(copiedPages[i]);
  }

  // Apply orientation changes if needed
  if (orientation !== 'keep-original') {
    const totalPages = mergedPdf.getPageCount();
    const startIdx = totalPages - copiedPages.length;

    for (let i = startIdx; i < totalPages; i++) {
      const page = mergedPdf.getPage(i);
      const { width, height } = page.getSize();
      const isLandscape = width > height;

      const needRotate =
        (orientation === 'all-portrait' && isLandscape) ||
        (orientation === 'all-landscape' && !isLandscape);

      if (needRotate) {
        const rot = page.getRotation().angle;
        page.setRotation(degrees((rot + 90) % 360));
        const box = page.getMediaBox();
        page.setMediaBox(box.y, box.x, box.height, box.width);
      }
    }
  }

  return copiedPages.length;
}
