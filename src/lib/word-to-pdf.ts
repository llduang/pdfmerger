/**
 * word-to-pdf.ts — Convert .docx to PDF via continuous-flow rendering.
 *
 * Uses breakPages: false → entire document flows as one block →
 * no absolute-positioning issues between sections → no image overlap.
 * Captures entire document as one canvas, then slices into pages.
 */
import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';
import { renderWordContinuous } from './word-render';

const SCALE = 2;
const RENDER_CLASS = 'wp-pdf-render';

export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  const html2canvas = (await import('html2canvas')).default;

  const renderResult = await renderWordContinuous(sourceArrayBuffer, RENDER_CLASS);

  const pageWidthPx = Math.round(renderResult.pageWidthMm * 96 / 25.4);
  const pageHeightPx = Math.round(renderResult.pageHeightMm * 96 / 25.4);

  const captureContainer = document.createElement('div');
  captureContainer.id = 'wp-capture-container';
  captureContainer.style.cssText = [
    'position:fixed;top:0;left:0;z-index:-1;',
    'pointer-events:none;background:white;',
    `width:${pageWidthPx}px;`,
  ].join('');
  document.body.appendChild(captureContainer);

  const resetStyle = document.createElement('style');
  resetStyle.id = 'wp-capture-reset';
  resetStyle.textContent = [
    '#wp-capture-container, #wp-capture-container * {',
    '  max-width: none !important;',
    '  max-height: none !important;',
    '}',
    '#wp-capture-container img {',
    '  max-width: none !important;',
    '  width: auto !important;',
    '  height: auto !important;',
    '}',
  ].join('\n');
  document.head.appendChild(resetStyle);

  try {
    captureContainer.innerHTML = renderResult.html;

    const wrapper = captureContainer.querySelector(
      `.${RENDER_CLASS}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.cssText =
        'background:white;padding:0;margin:0;box-shadow:none;border:none;';
    }

    await waitForImages(captureContainer);

    const totalHeightPx = captureContainer.scrollHeight;
    if (totalHeightPx <= 0) {
      throw new Error('Word 文档渲染后高度为 0');
    }

    const fullCanvas = await html2canvas(captureContainer, {
      scale: SCALE,
      backgroundColor: '#ffffff',
      width: pageWidthPx,
      height: totalHeightPx,
      useCORS: true,
      allowTaint: true,
      logging: false,
    });

    const pageHeightScaled = pageHeightPx * SCALE;
    const pageWidthScaled = pageWidthPx * SCALE;
    const totalHeightScaled = fullCanvas.height;
    const numPages = Math.ceil(totalHeightScaled / pageHeightScaled);

    let pdfWidthPt = (pageWidthPx * 72) / 96;
    let pdfHeightPt = (pageHeightPx * 72) / 96;

    if (orientation === 'all-landscape' && pdfWidthPt < pdfHeightPt) {
      [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
    } else if (orientation === 'all-portrait' && pdfWidthPt > pdfHeightPt) {
      [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
    }

    for (let i = 0; i < numPages; i++) {
      const startY = i * pageHeightScaled;
      const sliceHeight = Math.min(pageHeightScaled, totalHeightScaled - startY);
      if (sliceHeight <= 0) break;

      const pageCanvas = document.createElement('canvas');
      pageCanvas.width = pageWidthScaled;
      pageCanvas.height = sliceHeight;
      const ctx = pageCanvas.getContext('2d')!;
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, pageWidthScaled, sliceHeight);
      ctx.drawImage(fullCanvas, 0, startY, pageWidthScaled, sliceHeight, 0, 0, pageWidthScaled, sliceHeight);

      const pngDataUrl = pageCanvas.toDataURL('image/png');
      const base64Data = pngDataUrl.split(',')[1];
      const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
      const image = await mergedPdf.embedPng(imageBytes);

      const page = mergedPdf.addPage([pdfWidthPt, pdfHeightPt]);
      const heightRatio = sliceHeight / pageHeightScaled;
      const drawHeightPt = pdfHeightPt * heightRatio;
      const y = (pdfHeightPt - drawHeightPt) / 2;

      page.drawImage(image, { x: 0, y, width: pdfWidthPt, height: drawHeightPt });
    }

    return numPages;
  } finally {
    const styleEl = document.getElementById('wp-capture-reset');
    if (styleEl) document.head.removeChild(styleEl);
    if (captureContainer.parentNode) {
      document.body.removeChild(captureContainer);
    }
  }
}

async function waitForImages(container: HTMLElement): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return;

  await Promise.all(
    Array.from(images).map((img) => {
      const el = img as HTMLImageElement;
      if (el.complete && el.naturalWidth > 0) return Promise.resolve();
      return new Promise<void>((resolve) => {
        const done = () => {
          el.removeEventListener('load', done);
          el.removeEventListener('error', done);
          resolve();
        };
        el.addEventListener('load', done);
        el.addEventListener('error', done);
      });
    })
  );
  await new Promise((r) => setTimeout(r, 500));
}