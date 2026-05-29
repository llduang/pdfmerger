/**
 * word-to-pdf.ts — Convert .docx to PDF using client-side rendering.
 *
 * Uses docx-preview to render the Word document into HTML sections,
 * then html2canvas to capture each section as a high-quality image,
 * and finally pdf-lib to embed the images as PDF pages.
 *
 * This is the default client-side approach that works without any server.
 */

import html2canvas from 'html2canvas';
import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const RENDER_CLASS = 'wp-merge-render';

export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  // Step 1: Render .docx using docx-preview
  const { renderAsync } = await import('docx-preview');

  const container = document.createElement('div');
  container.style.cssText =
    'position:fixed;left:-9999px;top:0;width:100vw;pointer-events:none;';
  document.body.appendChild(container);

  try {
    await renderAsync(sourceArrayBuffer, container, undefined, {
      className: RENDER_CLASS,
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      ignoreLastRenderedPageBreak: true,
      experimental: false,
      trimXmlDeclaration: true,
      useBase64URL: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
    });

    // Step 2: Convert blob images to base64
    await convertBlobImagesToBase64(container);
    await waitForImages(container);
    await delay(500);

    // Step 3: Find all page sections
    let sections = container.querySelectorAll(`section.${RENDER_CLASS}`);
    if (sections.length === 0) {
      sections = container.querySelectorAll('section');
    }
    if (sections.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // Step 4: Move container into viewport for accurate rendering
    container.style.left = '0';
    container.style.top = '0';
    container.style.opacity = '1';
    container.style.zIndex = '-1';
    container.style.background = 'white';
    await delay(200);

    // Step 5: Capture each page section
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i] as HTMLElement;

      const wPx = section.offsetWidth || 794;
      const hPx = section.offsetHeight || 1123;

      try {
        const canvas = await html2canvas(section, {
          backgroundColor: '#ffffff',
          scale: 2,
          width: wPx,
          height: hPx,
          useCORS: true,
          allowTaint: true,
          logging: false,
        });

        const pngDataUrl = canvas.toDataURL('image/png');
        const base64Data = pngDataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
        const pdfImage = await mergedPdf.embedPng(imageBytes);

        // Convert CSS px to PDF points (96 dpi → 72 dpi)
        let pdfWidth = (wPx * 72) / 96;
        let pdfHeight = (hPx * 72) / 96;

        // Handle orientation swap
        if (orientation === 'all-landscape' && pdfWidth < pdfHeight) {
          [pdfWidth, pdfHeight] = [pdfHeight, pdfWidth];
        } else if (orientation === 'all-portrait' && pdfWidth > pdfHeight) {
          [pdfWidth, pdfHeight] = [pdfHeight, pdfWidth];
        }

        const page = mergedPdf.addPage([pdfWidth, pdfHeight]);
        page.drawImage(pdfImage, {
          x: 0,
          y: 0,
          width: pdfWidth,
          height: pdfHeight,
        });
      } catch (err) {
        throw new Error(
          `Word 文档第 ${i + 1} 页渲染失败: ${err instanceof Error ? err.message : String(err)}`
        );
      }
    }

    return sections.length;
  } finally {
    if (container.parentNode) {
      document.body.removeChild(container);
    }
  }
}

// ─── Helpers ───────────────────────────────────────────────────────────────

async function convertBlobImagesToBase64(
  container: HTMLElement
): Promise<void> {
  const images = container.querySelectorAll('img');
  for (let i = 0; i < images.length; i++) {
    const img = images[i] as HTMLImageElement;
    if (!img.src || !img.src.startsWith('blob:')) continue;
    try {
      const resp = await fetch(img.src);
      const blob = await resp.blob();
      img.src = await blobToBase64(blob);
    } catch {
      img.removeAttribute('src');
    }
  }
  await delay(200);
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function waitForImages(container: HTMLElement): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return Promise.resolve();

  return new Promise((resolve) => {
    let pending = images.length;
    const finish = () => {
      if (--pending <= 0) resolve();
    };
    for (let i = 0; i < images.length; i++) {
      const img = images[i] as HTMLImageElement;
      if (img.complete && img.naturalWidth > 0) {
        finish();
      } else {
        img.addEventListener('load', finish, { once: true });
        img.addEventListener('error', finish, { once: true });
      }
    }
    setTimeout(resolve, 12000);
  });
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}
