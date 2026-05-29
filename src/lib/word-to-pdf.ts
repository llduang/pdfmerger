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

  // 1. Create containers
  // Style container: inject into head so html2canvas picks up all docx-preview styles
  const styleContainer = document.createElement('div');
  styleContainer.id = `${CLASS_NAME}-styles`;
  document.head.appendChild(styleContainer);

  // Body container: hidden off-screen area for docx-preview to render pages into
  const bodyContainer = document.createElement('div');
  bodyContainer.id = `${CLASS_NAME}-container`;
  bodyContainer.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    z-index: -9999;
    overflow: visible;
    background: transparent;
    opacity: 1;
    visibility: visible;
  `;
  document.body.appendChild(bodyContainer);

  try {
    // 2. Render the Word document
    await renderAsync(sourceArrayBuffer, bodyContainer, styleContainer, {
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
    //   bodyContainer > div.{className}-wrapper > section.{className}
    let pageElements = bodyContainer.querySelectorAll(`section.${CLASS_NAME}`);
    if (pageElements.length === 0) {
      // Fallback: any section inside the container
      pageElements = bodyContainer.querySelectorAll('section');
    }
    if (pageElements.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // 4. Wait a moment for all styles to apply and images to load
    await new Promise((r) => setTimeout(r, 500));

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
    // 6. Clean up containers
    if (bodyContainer.parentNode) {
      document.body.removeChild(bodyContainer);
    }
    if (styleContainer.parentNode) {
      document.head.removeChild(styleContainer);
    }
  }
}
