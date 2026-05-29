/**
 * word-to-pdf.ts — Convert a .docx file to PDF pages using docx-preview + html-to-image.
 *
 * Key design decisions:
 *   - Uses html-to-image (SVG foreignObject) instead of html2canvas.
 *   - All blob images are converted to base64 BEFORE capture.
 *   - Page dimensions are detected from docx-preview sections for
 *     accurate PDF page sizing.
 *   - FIX: When orientation changes, the image draw dimensions are
 *     properly adjusted to match the new page dimensions.
 */
import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';
import { renderWordToHtml } from './word-render';

const SCALE = 2;
const RENDER_CLASS = 'wp-pdf-render';

/**
 * Render a Word document to PDF pages and add them to `mergedPdf`.
 * Returns the number of pages added.
 */
export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  const { toPng } = await import('html-to-image');

  const renderResult = await renderWordToHtml(
    sourceArrayBuffer,
    RENDER_CLASS
  );

  const captureContainer = document.createElement('div');
  captureContainer.id = 'wp-capture-container';
  captureContainer.style.cssText =
    'position:fixed;top:0;left:0;width:100vw;z-index:-1;pointer-events:none;background:white;';
  document.body.appendChild(captureContainer);

  try {
    captureContainer.innerHTML = renderResult.html;

    const wrapper = captureContainer.querySelector(
      `.${RENDER_CLASS}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.cssText =
        'background:white;padding:0;margin:0;box-shadow:none;';
    }

    const captureImages = captureContainer.querySelectorAll('img');
    await Promise.all(
      Array.from(captureImages).map((img) => {
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
    await new Promise((r) => setTimeout(r, 300));

    let sections = captureContainer.querySelectorAll(
      `section.${RENDER_CLASS}`
    );
    if (sections.length === 0) {
      sections = captureContainer.querySelectorAll('section');
    }
    if (sections.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    for (let i = 0; i < sections.length; i++) {
      const section = sections[i] as HTMLElement;

      section.style.background = '#ffffff';
      section.style.boxShadow = 'none';
      section.style.margin = '0';
      section.style.overflow = 'hidden';

      try {
        const wPx = section.offsetWidth || parseFloat(section.style.width) || 794;
        const hPx = section.offsetHeight || parseFloat(section.style.height) || 1123;

        const pngDataUrl = await toPng(section, {
          pixelRatio: SCALE,
          backgroundColor: '#ffffff',
          width: wPx,
          height: hPx,
          style: {
            width: `${wPx}px`,
            height: `${hPx}px`,
            overflow: 'hidden',
            background: '#ffffff',
          },
          filter: (node) => {
            return node !== captureContainer;
          },
        });

        const base64Data = pngDataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) =>
          c.charCodeAt(0)
        );
        const image = await mergedPdf.embedPng(imageBytes);

        // Calculate PDF page dimensions (96 CSS px = 72 pt = 1 inch)
        let pdfWidthPt = (wPx * 72) / 96;
        let pdfHeightPt = (hPx * 72) / 96;

        // FIX: When orientation changes, swap page dimensions AND adjust
        // the draw dimensions accordingly. Previously only page dims were
        // swapped but the image was still drawn with the original dimensions,
        // causing stretching.
        let drawWidth: number;
        let drawHeight: number;

        if (orientation === 'all-landscape' && pdfWidthPt < pdfHeightPt) {
          // Swap to landscape: image becomes narrower relative to page
          drawWidth = pdfWidthPt;
          drawHeight = pdfHeightPt;
          [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
          // Image should fill the new landscape page, maintaining aspect ratio
          const scaleX = pdfWidthPt / drawWidth;
          const scaleY = pdfHeightPt / drawHeight;
          const scale = Math.min(scaleX, scaleY, 1);
          drawWidth *= scale;
          drawHeight *= scale;
        } else if (orientation === 'all-portrait' && pdfWidthPt > pdfHeightPt) {
          drawWidth = pdfWidthPt;
          drawHeight = pdfHeightPt;
          [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
          const scaleX = pdfWidthPt / drawWidth;
          const scaleY = pdfHeightPt / drawHeight;
          const scale = Math.min(scaleX, scaleY, 1);
          drawWidth *= scale;
          drawHeight *= scale;
        } else {
          drawWidth = pdfWidthPt;
          drawHeight = pdfHeightPt;
        }

        const page = mergedPdf.addPage([pdfWidthPt, pdfHeightPt]);

        // Center the image on the page
        const x = (pdfWidthPt - drawWidth) / 2;
        const y = (pdfHeightPt - drawHeight) / 2;

        page.drawImage(image, {
          x,
          y,
          width: drawWidth,
          height: drawHeight,
        });
      } catch (err) {
        console.error(`Failed to capture Word page ${i + 1}:`, err);
      }
    }

    return sections.length;
  } finally {
    if (captureContainer.parentNode) {
      document.body.removeChild(captureContainer);
    }
  }
}
