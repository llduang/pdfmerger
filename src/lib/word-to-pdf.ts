/**
 * word-to-pdf.ts — Convert a .docx file to PDF pages using docx-preview + html2canvas.
 *
 * Key design decisions:
 *   - Uses html2canvas instead of html-to-image (SVG foreignObject).
 *     html2canvas directly renders the DOM to canvas, which correctly
 *     preserves absolute positioning and CSS transforms used by
 *     docx-preview for image placement.
 *   - All blob images are converted to base64 BEFORE capture.
 *   - Adds scoped CSS reset to neutralize Tailwind's preflight
 *     (specifically `img { max-width: 100%; height: auto }` which
 *     can break docx-preview's image sizing).
 *   - Explicitly sets `position: relative` on each section to ensure
 *     absolutely-positioned images are placed correctly.
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
  // html2canvas renders the DOM directly to canvas — much more reliable
  // than html-to-image's SVG foreignObject approach for absolute
  // positioning and CSS transforms used by docx-preview.
  const html2canvas = (await import('html2canvas')).default;

  const renderResult = await renderWordToHtml(
    sourceArrayBuffer,
    RENDER_CLASS
  );

  // Create capture container (visible for layout, behind other content)
  const captureContainer = document.createElement('div');
  captureContainer.id = 'wp-capture-container';
  captureContainer.style.cssText =
    'position:fixed;top:0;left:0;z-index:-1;pointer-events:none;background:white;';
  document.body.appendChild(captureContainer);

  // CRITICAL: Add scoped CSS reset to neutralize Tailwind v4 preflight.
  // Tailwind injects `img { max-width: 100%; height: auto }` globally,
  // which can shrink / reposition images that docx-preview sizes via
  // wrapper divs.  This scoped reset ensures images inside the capture
  // container are rendered exactly as docx-preview intended.
  const resetStyle = document.createElement('style');
  resetStyle.id = 'wp-capture-reset';
  resetStyle.textContent = [
    '#wp-capture-container img { max-width: none !important; }',
    '#wp-capture-container section { position: relative !important; }',
  ].join('\n');
  document.head.appendChild(resetStyle);

  try {
    // Insert rendered HTML (includes <style> tags from docx-preview)
    captureContainer.innerHTML = renderResult.html;

    // Reset wrapper — remove decorative styles
    const wrapper = captureContainer.querySelector(
      `.${RENDER_CLASS}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.cssText =
        'background:white;padding:0;margin:0;box-shadow:none;';
    }

    // Wait for all images to fully load before capturing
    await waitForImages(captureContainer);

    // Locate page sections
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

      // Explicitly ensure correct positioning for absolute children
      section.style.position = 'relative';
      section.style.overflow = 'hidden';
      section.style.background = '#ffffff';
      section.style.boxShadow = 'none';
      section.style.margin = '0';

      const wPx = section.offsetWidth || 794;
      const hPx = section.offsetHeight || 1123;

      // Size the capture container to exactly match the section so
      // the section's layout doesn't shift when we capture it.
      captureContainer.style.width = `${wPx}px`;
      captureContainer.style.height = `${hPx}px`;

      try {
        const canvas = await html2canvas(section, {
          scale: SCALE,
          backgroundColor: '#ffffff',
          width: wPx,
          height: hPx,
          useCORS: true,
          allowTaint: true,
          logging: false,
        });

        const pngDataUrl = canvas.toDataURL('image/png');
        const base64Data = pngDataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) =>
          c.charCodeAt(0)
        );
        const image = await mergedPdf.embedPng(imageBytes);

        // Calculate PDF page dimensions (96 CSS px = 72 pt = 1 inch)
        let pdfWidthPt = (wPx * 72) / 96;
        let pdfHeightPt = (hPx * 72) / 96;

        // Handle orientation changes — swap page dims AND adjust
        // draw dims to prevent stretching.
        let drawWidth: number;
        let drawHeight: number;

        if (orientation === 'all-landscape' && pdfWidthPt < pdfHeightPt) {
          drawWidth = pdfWidthPt;
          drawHeight = pdfHeightPt;
          [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
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
    // Clean up scoped styles and container
    const styleEl = document.getElementById('wp-capture-reset');
    if (styleEl) document.head.removeChild(styleEl);
    if (captureContainer.parentNode) {
      document.body.removeChild(captureContainer);
    }
  }
}

/**
 * Wait for all <img> elements inside `container` to finish loading.
 * Includes a short settle delay for layout reflow after images load.
 */
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
  // Extra settle time for layout reflow after images load
  await new Promise((r) => setTimeout(r, 300));
}