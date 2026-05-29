import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';
import { PAGE_SIZES } from './merge-pdf';

const SCALE = 2; // Render at 2x for crisp PDF output

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

  // 1. Create a hidden container for docx-preview to render into
  const wrapper = document.createElement('div');
  wrapper.style.cssText = 'position:fixed;top:0;left:0;z-index:-9999;overflow:visible;background:transparent;';
  document.body.appendChild(wrapper);

  try {
    // 2. Render the Word document — docx-preview creates pages as section elements
    await renderAsync(sourceArrayBuffer, wrapper, null, {
      className: 'docx-preview-render',
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
    const pageElements = wrapper.querySelectorAll('.docx-preview-render > section');
    if (pageElements.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // 4. Render each page to canvas and embed into PDF
    for (let i = 0; i < pageElements.length; i++) {
      const pageEl = pageElements[i] as HTMLElement;

      // Make sure the page element is visible for html2canvas
      pageEl.style.position = 'relative';
      pageEl.style.overflow = 'hidden';
      pageEl.style.background = '#ffffff';

      // Small delay for rendering
      await new Promise((r) => setTimeout(r, 100));

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
      // docx-preview renders at screen resolution, convert to PDF points
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
    // 5. Clean up the render container
    if (wrapper.parentNode) {
      document.body.removeChild(wrapper);
    }
  }
}
