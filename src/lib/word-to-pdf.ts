import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const SCALE = 2; // Render at 2x for crisp PDF output
const CLASS_NAME = 'docx-word-merge';

/**
 * Render a Word document (.docx) to PDF pages using docx-preview for high-fidelity output.
 * docx-preview preserves Word's page layout, styles, tables, headers, footers, etc.
 *
 * Flow: docx-preview renderAsync → per-page html2canvas → embed PNG into pdf-lib
 */
export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  // Dynamic imports (browser-only)
  const { renderAsync } = await import('docx-preview');
  const html2canvas = (await import('html2canvas')).default;

  // 1. Create a render container
  // Place it off-screen to the left (not display:none or visibility:hidden, so browser fully renders it)
  const container = document.createElement('div');
  container.id = `${CLASS_NAME}-container`;
  container.style.cssText = `
    position: absolute;
    left: -10000px;
    top: 0;
    width: 10000px;
    overflow: visible;
    background: #ffffff;
    opacity: 1;
    visibility: visible;
  `;
  document.body.appendChild(container);

  try {
    // 2. Render the Word document
    // Pass null for styleContainer so styles are injected directly into bodyContainer
    // This avoids the issue of browsers moving <div> elements out of <head>
    await renderAsync(sourceArrayBuffer, container, null, {
      className: CLASS_NAME,
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

    // 3. Find all rendered page sections
    // docx-preview DOM structure (inWrapper=true):
    //   container > div.{className}-wrapper > section.{className}
    let pageElements = container.querySelectorAll(`section.${CLASS_NAME}`);
    if (pageElements.length === 0) {
      // Fallback: any section elements
      pageElements = container.querySelectorAll('section');
    }
    if (pageElements.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // 4. Wait for all images to fully load and browser to complete layout
    // renderAsync resolves when its internal tasks are done, but the browser
    // may still need time to decode images and finish layout/paint
    await waitForImagesLoaded(container);
    await new Promise((r) => setTimeout(r, 300));

    // 5. Render each page to canvas and embed into PDF
    for (let i = 0; i < pageElements.length; i++) {
      const pageEl = pageElements[i] as HTMLElement;

      const canvas = await html2canvas(pageEl, {
        scale: SCALE,
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        logging: false,
      });

      if (canvas.width === 0 || canvas.height === 0) continue;

      // Convert canvas to PNG
      const pngDataUrl = canvas.toDataURL('image/png');
      const base64Data = pngDataUrl.split(',')[1];
      const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
      const image = await mergedPdf.embedPng(imageBytes);

      // Calculate page size from the rendered element
      // docx-preview renders at screen resolution, convert to PDF points (96dpi → 72dpi)
      const renderedWidthPx = pageEl.offsetWidth;
      const renderedHeightPx = pageEl.offsetHeight;
      const pageWidthPt = (renderedWidthPx * 72) / 96;
      const pageHeightPt = (renderedHeightPx * 72) / 96;

      // Handle orientation
      let pdfWidth = pageWidthPt;
      let pdfHeight = pageHeightPt;
      if (orientation === 'all-landscape' && pdfWidth < pdfHeight) {
        [pdfWidth, pdfHeight] = [pdfHeight, pdfWidth];
      } else if (orientation === 'all-portrait' && pdfWidth > pdfHeight) {
        [pdfWidth, pdfHeight] = [pdfHeight, pdfWidth];
      }

      const page = mergedPdf.addPage([pdfWidth, pdfHeight]);
      page.drawImage(image, {
        x: 0,
        y: 0,
        width: pdfWidth,
        height: pdfHeight,
      });
    }

    return pageElements.length;
  } finally {
    // 6. Clean up the render container
    if (container.parentNode) {
      document.body.removeChild(container);
    }
  }
}

/**
 * Wait for all images inside a container to finish loading.
 */
function waitForImagesLoaded(container: HTMLElement): Promise<void> {
  return new Promise((resolve) => {
    const images = container.querySelectorAll('img');
    if (images.length === 0) {
      resolve();
      return;
    }

    let pending = images.length;
    let resolved = false;

    const done = () => {
      if (resolved) return;
      resolved = true;
      resolve();
    };

    for (let i = 0; i < images.length; i++) {
      const img = images[i] as HTMLImageElement;

      if (img.complete && img.naturalWidth > 0) {
        pending--;
        if (pending === 0) done();
        continue;
      }

      img.addEventListener('load', () => {
        pending--;
        if (pending === 0) done();
      }, { once: true });

      img.addEventListener('error', () => {
        // Even on error, count as done (we'll just see a broken image)
        pending--;
        if (pending === 0) done();
      }, { once: true });
    }

    // Safety timeout: don't wait forever
    setTimeout(done, 5000);
  });
}
