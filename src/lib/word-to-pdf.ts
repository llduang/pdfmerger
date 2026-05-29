/**
 * word-to-pdf.ts — Convert .docx to PDF via Cloudflare Pages Function.
 *
 * Sends the .docx to /api/convert (same-origin Pages Function → Gotenberg).
 * Includes retry logic, extended timeout, and pre-warm for Railway cold starts.
 *
 * 使用同域 Pages Function 消除了：
 *  - CORS 跨域问题
 *  - workers.dev 域名被墙问题
 *  - 独立 Worker 部署维护成本
 */

import { PDFDocument, degrees } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

// 使用同域 Pages Function，无需跨域
const CONVERT_API = '/api/convert';
const MAX_RETRIES = 4;
const REQUEST_TIMEOUT = 90000; // 90 seconds (Railway cold start + Gotenberg conversion)

let warmed = false;
let warmupPromise: Promise<void> | null = null;

/**
 * Pre-warm the Railway/Gotenberg service so it's ready when user clicks merge.
 * Call this as soon as a Word file is uploaded.
 */
export async function warmUpConverter(): Promise<void> {
  if (warmed || warmupPromise) return;
  warmupPromise = doWarmUp();
  await warmupPromise;
}

async function doWarmUp(): Promise<void> {
  try {
    // Send a tiny POST to wake up Railway/Gotenberg
    const dummyBlob = new Blob([''], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const formData = new FormData();
    formData.append('file', dummyBlob, 'warmup.docx');

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);
    await fetch(CONVERT_API, {
      method: 'POST',
      body: formData,
      signal: controller.signal,
    });
    clearTimeout(timeoutId);
    warmed = true;
  } catch {
    // Warm-up failed silently — the real request will retry anyway
    warmed = false;
  } finally {
    warmupPromise = null;
  }
}

/**
 * Reset warm state (e.g., after conversion failure, allow re-warmup)
 */
export function resetWarmup(): void {
  warmed = false;
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

      // 处理错误响应
      const contentType = response.headers.get('Content-Type') || '';
      let errorDetail = '';

      if (contentType.includes('application/json')) {
        try {
          const errJson = await response.json();
          errorDetail = errJson.error || errJson.message || '';
        } catch {
          // ignore parse errors
        }
      } else {
        errorDetail = (await response.text().catch(() => '')).slice(0, 200);
      }

      // 503 表示服务暂时不可用（Gotenberg 正在启动），可重试
      if (response.status === 503) {
        lastError = new Error(
          '转换服务正在启动中，请稍候自动重试...'
        );
        // 重试前先重新 warmup
        resetWarmup();
        warmUpConverter();
        if (attempt <= MAX_RETRIES) {
          await new Promise((r) => setTimeout(r, 3000 * attempt));
          continue;
        }
        break;
      }

      // 400-499 客户端错误，不重试
      if (response.status >= 400 && response.status < 500) {
        lastError = new Error(
          `文件转换失败（${response.status}），请检查 Word 文件是否损坏`
        );
        break;
      }

      // 其他服务端错误，可重试
      lastError = new Error(
        `转换服务返回错误（HTTP ${response.status}）${errorDetail ? ': ' + errorDetail : ''}`
      );

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

      // 重试前先重新 warmup
      resetWarmup();
      warmUpConverter();

      // Wait longer before each retry (2s, 4s, 6s, 8s)
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
