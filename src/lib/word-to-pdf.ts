/**
 * word-to-pdf.ts — Convert .docx to PDF via Cloudflare Pages Function.
 *
 * Sends the .docx to /api/convert (same-origin Pages Function → Microsoft Graph API).
 * Microsoft Graph API uses the same engine as desktop Word, so the conversion
 * result is identical to "Save as PDF" in Word.
 */

import { PDFDocument, degrees } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const CONVERT_API = '/api/convert';
const MAX_RETRIES = 3;
const REQUEST_TIMEOUT = 120000; // 120 seconds

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

      const contentType = response.headers.get('Content-Type') || '';
      let errorDetail = '';
      let retryable = false;

      if (contentType.includes('application/json')) {
        try {
          const errJson = await response.json();
          errorDetail = errJson.error || errJson.message || '';
          retryable = errJson.retryable === true;
        } catch {}
      } else {
        errorDetail = (await response.text().catch(() => '')).slice(0, 200);
      }

      if (response.status === 503 || retryable) {
        lastError = new Error('转换服务暂时繁忙，正在自动重试...');
        if (attempt <= MAX_RETRIES) {
          await new Promise((r) => setTimeout(r, 2000 * attempt));
          continue;
        }
        break;
      }

      if (response.status >= 400 && response.status < 500) {
        lastError = new Error(
          `文件转换失败（${response.status}），请检查 Word 文件是否损坏`
        );
        break;
      }

      lastError = new Error(
        `转换服务返回错误（HTTP ${response.status}）${errorDetail ? ': ' + errorDetail : ''}`
      );

      if (attempt <= MAX_RETRIES) {
        await new Promise((r) => setTimeout(r, 2000 * attempt));
        continue;
      }
    } catch (err) {
      lastError = err instanceof Error ? err : new Error(String(err));

      if (lastError.message.includes('Aborted') || lastError.message.includes('abort')) {
        lastError = new Error(
          `转换服务响应超时（${REQUEST_TIMEOUT / 1000}秒），请稍后重试`
        );
      } else if (lastError.message.includes('Failed to fetch') || lastError.message.includes('NetworkError')) {
        lastError = new Error('无法连接转换服务，请检查网络连接后重试');
      }

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