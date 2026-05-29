/**
 * word-to-pdf.ts — Convert a .docx file to PDF pages using docx-preview + html-to-image.
 *
 * Key design decisions (May 2026 rewrite):
 *   - Uses html-to-image (SVG foreignObject) instead of html2canvas.
 *     html-to-image relies on the browser's native SVG renderer, which
 *     handles images (including base64) and CSS more faithfully than
 *     html2canvas's JavaScript-based re-rendering.
 *   - All blob images are converted to base64 BEFORE capture (via
 *     the shared renderWordToHtml utility).
 *   - Page dimensions are detected from docx-preview sections for
 *     accurate PDF page sizing.
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

  // ── Step 1: Render with shared utility (handles image conversion) ──
  const renderResult = await renderWordToHtml(
    sourceArrayBuffer,
    RENDER_CLASS
  );

  // ── Step 2: Re-create a visible container to capture from ──
  // We need the DOM elements to be in the document for html-to-image.
  const captureContainer = document.createElement('div');
  captureContainer.id = 'wp-capture-container';
  captureContainer.style.cssText =
    'position:fixed;top:0;left:0;width:100vw;z-index:-1;pointer-events:none;background:white;';
  document.body.appendChild(captureContainer);

  try {
    // Inject the rendered HTML into the capture container
    captureContainer.innerHTML = renderResult.html;

    // Clean wrapper
    const wrapper = captureContainer.querySelector(
      `.${RENDER_CLASS}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.cssText =
        'background:white;padding:0;margin:0;box-shadow:none;';
    }

    // Wait for base64 images to load in this new context
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

    // ── Step 3: Find page sections ──
    let sections = captureContainer.querySelectorAll(
      `section.${RENDER_CLASS}`
    );
    if (sections.length === 0) {
      sections = captureContainer.querySelectorAll('section');
    }
    if (sections.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // ── Step 4: Capture each section and embed in PDF ──
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i] as HTMLElement;

      // Clean section styles for capture
      section.style.background = '#ffffff';
      section.style.boxShadow = 'none';
      section.style.margin = '0';
      section.style.overflow = 'hidden';

      try {
        // Get the section's actual pixel dimensions
        const wPx = section.offsetWidth || parseFloat(section.style.width) || 794;
        const hPx = section.offsetHeight || parseFloat(section.style.height) || 1123;

        // Capture using html-to-image (SVG foreignObject approach)
        const pngDataUrl = await toPng(section, {
          pixelRatio: SCALE,
          backgroundColor: '#ffffff',
          width: wPx,
          height: hPx,
          style: {
            // Ensure the element renders at exact dimensions during capture
            width: `${wPx}px`,
            height: `${hPx}px`,
            overflow: 'hidden',
            background: '#ffffff',
          },
          // Include computed styles for accurate rendering
          filter: (node) => {
            // Skip the capture container itself
            return node !== captureContainer;
          },
        });

        // Embed into PDF
        const base64Data = pngDataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) =>
          c.charCodeAt(0)
        );
        const image = await mergedPdf.embedPng(imageBytes);

        // Calculate PDF page dimensions (96 CSS px = 72 pt = 1 inch)
        const pageWidthPt = (wPx * 72) / 96;
        const pageHeightPt = (hPx * 72) / 96;

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
