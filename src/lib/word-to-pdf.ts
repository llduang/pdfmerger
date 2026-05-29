/**
 * word-to-pdf.ts — Convert .docx to PDF via server-side Gotenberg/LibreOffice.
 *
 * Sends the .docx to Cloudflare Worker → Gotenberg (Railway).
 * Includes retry logic, extended timeout, and pre-warm for Railway cold starts.
 */

import { PDFDocument, degrees } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const CONVERT_API = 'https://pdf-convert.2710476780.workers.dev';
const MAX_RETRIES = 3;
const REQUEST_TIMEOUT = 60000; // 60 seconds (Railway free tier cold start is slow)

let warmed = false;

/**
 * Pre-warm the Railway/Gotenberg service so it's ready when user clicks merge.
 * Call this as soon as a Word file is uploaded.
 */
export async function warmUpConverter(): Promise<void> {
  if (warmed) return;
  warmed = true;
  try {
    // Send a tiny POST to wake up Railway/Gotenberg
    const dummyBlob = new Blob([''], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const formData = new FormData();
    formData.append('file', dummyBlob, 'warmup.docx');

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 20000);
    await fetch(CONVERT_API, {
      method: 'POST',
      body: formData,
      signal: controller.signal,
    });
    clearTimeout(timeoutId);
  } catch {
    // Warm-up failed silently — the real request will retry anyway
    warmed = false;
  }
}

export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  const formData = new FormData();
  formData.append(
    'file',
    new Blob([sourceArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
    'document.docx'
  );

  let lastError: Error | null = null;

  for (let attempt = 1; attempt <= MAX_RETRIES + 1; attempt++) {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT);

      const response = await fetch(CONVERT_API, {
        method: 'POST',
        body: formData,
        signal: controller.signal,
      });

      clearTimeout(timeoutId);

      if (response.ok) {
        const pdfBytes = new Uint8Array(await response.arrayBuffer());
        return await embedPdfPages(mergedPdf, pdfBytes, orientation);
      }

      const text = await response.text().catch(() => '');
      lastError = new Error(
        `转换服务返回错误（HTTP ${response.status}）${text ? ': ' + text.slice(0, 200) : ''}`
      );

      // Don't retry on 4xx errors (client errors)
      if (response.status >= 400 && response.status < 500) {
        break;
      }
    } catch (err) {
      lastError = err instanceof Error ? err : new Error(String(err));

      if (lastError.message.includes('Aborted') || lastError.message.includes('abort')) {
        lastError = new Error(
          `转换服务响应超时（${REQUEST_TIMEOUT / 1000}秒），服务可能正在启动中，请稍后重试`
        );
      } else if (lastError.message.includes('Failed to fetch') || lastError.message.includes('NetworkError')) {
        lastError = new Error(
          '无法连接转换服务，请检查网络连接后重试'
        );
      }

      // Wait longer before each retry (2s, 4s, 6s)
      if (attempt <= MAX_RETRIES) {
        await new Promise((r) => setTimeout(r, 2000 * attempt));
      }
    }
  }

  throw lastError || new Error('Word 转换失败，请重试');
}

async function embedPdfPages(
  mergedPdf: PDFDocument,
  pdfBytes: Uint8Array,
  orientation: OrientationMode
): Promise<number> {
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
