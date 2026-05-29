import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const SCALE = 2;
const CLASS_NAME = 'docx-word-merge';

/**
 * Render a Word document (.docx) to PDF pages.
 *
 * Pipeline:
 *   docx-preview renderAsync  (high-fidelity Word rendering)
 *   → ensure all images loaded
 *   → html-to-image toPng     (native SVG foreignObject — NOT html2canvas)
 *   → embed PNG into pdf-lib
 *
 * Why html-to-image instead of html2canvas:
 *   html2canvas reimplements CSS in JS → poor image/CSS support
 *   html-to-image uses the browser's own renderer → perfect fidelity
 */
export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  // Dynamic imports (browser-only)
  const { renderAsync } = await import('docx-preview');
  const { toPng } = await import('html-to-image');

  // 1. Create render container — visible, within viewport, behind everything
  const container = document.createElement('div');
  container.id = `${CLASS_NAME}-container`;
  container.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: auto;
    min-height: 100vh;
    z-index: -1;
    pointer-events: none;
    background: white;
  `;
  document.body.appendChild(container);

  try {
    // 2. Render Word document with docx-preview
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
    let pageElements = container.querySelectorAll(`section.${CLASS_NAME}`);
    if (pageElements.length === 0) {
      pageElements = container.querySelectorAll('section');
    }
    if (pageElements.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // 4. Ensure all images are loaded and convert any blob URLs to base64
    await ensureAllImagesLoaded(container);

    // 5. Wait for browser to complete layout/paint
    await new Promise((r) => setTimeout(r, 600));

    // 6. Override wrapper styles that could interfere with screenshot
    // docx-preview adds gray background, padding, and centering to the wrapper
    const wrapper = container.querySelector(`.${CLASS_NAME}-wrapper`) as HTMLElement;
    if (wrapper) {
      wrapper.style.background = 'white';
      wrapper.style.padding = '0';
      wrapper.style.margin = '0';
    }

    // 7. Render each page section to PNG and embed into PDF
    for (let i = 0; i < pageElements.length; i++) {
      const pageEl = pageElements[i] as HTMLElement;

      // Ensure the section has a clean white background
      pageEl.style.background = '#ffffff';
      pageEl.style.boxShadow = 'none';
      pageEl.style.margin = '0';

      try {
        // Use html-to-image which leverages browser's native SVG foreignObject rendering
        const dataUrl = await toPng(pageEl, {
          quality: 1,
          pixelRatio: SCALE,
          style: {
            // Force white background for the capture
            background: '#ffffff',
          },
          // Filter out any remaining blob URLs
          filter: (node: Node) => {
            if (node instanceof HTMLImageElement && node.src?.startsWith('blob:')) {
              return false;
            }
            return true;
          },
        });

        if (!dataUrl || dataUrl.length < 100) continue;

        // Extract base64 data from the data URL
        const base64Data = dataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
        const image = await mergedPdf.embedPng(imageBytes);

        // Calculate PDF page dimensions from rendered element
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
      } catch (err) {
        console.error(`Failed to render Word page ${i + 1}:`, err);
        // Continue with remaining pages instead of failing the whole document
      }
    }

    return pageElements.length;
  } finally {
    if (container.parentNode) {
      document.body.removeChild(container);
    }
  }
}

/**
 * Ensure all images in the container are fully loaded.
 * Converts blob URLs to base64 data URLs for reliable capture.
 */
async function ensureAllImagesLoaded(container: HTMLElement): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return;

  // Convert any blob URLs to base64 data URLs
  for (let i = 0; i < images.length; i++) {
    const img = images[i] as HTMLImageElement;
    if (!img.src || img.src === '') continue;

    if (img.src.startsWith('blob:')) {
      try {
        const response = await fetch(img.src);
        const blob = await response.blob();
        img.src = await blobToBase64(blob);
      } catch (e) {
        console.warn('Failed to convert blob image to base64:', e);
      }
    }
  }

  // Wait for all images to finish loading
  await Promise.all(
    Array.from(images).map((img) => {
      const el = img as HTMLImageElement;
      if (el.complete && el.naturalWidth > 0) return Promise.resolve();
      return new Promise<void>((resolve) => {
        const onDone = () => {
          el.removeEventListener('load', onDone);
          el.removeEventListener('error', onDone);
          resolve();
        };
        el.addEventListener('load', onDone);
        el.addEventListener('error', onDone);
      });
    })
  );

  // Extra wait for browser to decode images
  await new Promise((r) => setTimeout(r, 500));
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
