import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const SCALE = 2;
const CLASS_NAME = 'docx-merge-render';

/**
 * Render a Word document (.docx) to PDF pages using docx-preview + html2canvas.
 *
 * Key improvements:
 *   - Converts ALL blob images to base64 before capture
 *   - Waits for all images to fully render
 *   - Detects page dimensions from docx-preview sections
 *   - Cleans wrapper styles to prevent extra pages
 */
export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  const { renderAsync } = await import('docx-preview');
  const html2canvas = (await import('html2canvas')).default;

  // Create render container in viewport
  const container = document.createElement('div');
  container.id = `${CLASS_NAME}-container`;
  container.style.cssText = `
    position: fixed; top: 0; left: 0;
    width: 100vw; height: auto; min-height: 100vh;
    z-index: -1; pointer-events: none;
    background: white;
  `;
  document.body.appendChild(container);

  try {
    // Render with docx-preview
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

    // Find page sections
    let pageElements = container.querySelectorAll(`section.${CLASS_NAME}`);
    if (pageElements.length === 0) {
      pageElements = container.querySelectorAll('section');
    }
    if (pageElements.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // Convert ALL blob images to base64 (critical for html2canvas capture)
    await convertAllImagesToBase64(container);

    // Wait for all images to fully load and render
    await waitForAllImages(container);
    await new Promise((r) => setTimeout(r, 500));

    // Clean up wrapper styles that could affect capture
    const wrapper = container.querySelector(
      `.${CLASS_NAME}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.background = 'white';
      wrapper.style.padding = '0';
      wrapper.style.margin = '0';
      wrapper.style.boxShadow = 'none';
    }

    // Render each page to canvas and embed into PDF
    for (let i = 0; i < pageElements.length; i++) {
      const pageEl = pageElements[i] as HTMLElement;
      pageEl.style.background = '#ffffff';
      pageEl.style.boxShadow = 'none';
      pageEl.style.margin = '0';
      pageEl.style.overflow = 'hidden';

      try {
        // Get actual page dimensions from the section's inline styles
        const computedWidth =
          parseFloat(pageEl.style.width) || pageEl.offsetWidth;
        const computedHeight =
          parseFloat(pageEl.style.height) || pageEl.offsetHeight;

        const canvas = await html2canvas(pageEl, {
          scale: SCALE,
          useCORS: true,
          allowTaint: true,
          backgroundColor: '#ffffff',
          logging: false,
          imageTimeout: 15000,
          width: computedWidth,
          height: computedHeight,
          // Force the captured area to exactly match the section bounds
          windowWidth: computedWidth,
          windowHeight: computedHeight,
          onclone: (clonedDoc) => {
            // Find the cloned element
            const clonedEl = clonedDoc.querySelector(
              `[data-html2canvas-id="${pageEl.getAttribute('data-html2canvas-id')}"]`
            ) as HTMLElement;
            if (!clonedEl) return;

            // Ensure the cloned element has exact dimensions
            clonedEl.style.width = `${computedWidth}px`;
            clonedEl.style.height = `${computedHeight}px`;
            clonedEl.style.overflow = 'hidden';
            clonedEl.style.background = '#ffffff';

            // Fix any remaining blob images in the clone
            const imgs = clonedEl.querySelectorAll('img');
            imgs.forEach((img) => {
              const el = img as HTMLImageElement;
              if (el.src && el.src.startsWith('blob:')) {
                el.removeAttribute('src');
              }
            });
          },
        });

        if (canvas.width === 0 || canvas.height === 0) continue;

        const pngDataUrl = canvas.toDataURL('image/png');
        const base64Data = pngDataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) =>
          c.charCodeAt(0)
        );
        const image = await mergedPdf.embedPng(imageBytes);

        // Calculate PDF page dimensions from the section's actual rendered size
        const pageWidthPt = (computedWidth * 72) / 96;
        const pageHeightPt = (computedHeight * 72) / 96;

        let pdfWidth = pageWidthPt;
        let pdfHeight = pageHeightPt;
        if (orientation === 'all-landscape' && pdfWidth < pdfHeight) {
          [pdfWidth, pdfHeight] = [pdfHeight, pdfWidth];
        } else if (orientation === 'all-portrait' && pdfWidth > pdfHeight) {
          [pdfWidth, pdfHeight] = [pdfHeight, pdfWidth];
        }

        const page = mergedPdf.addPage([pdfWidth, pdfHeight]);
        page.drawImage(image, { x: 0, y: 0, width: pdfWidth, height: pdfHeight });
      } catch (err) {
        console.error(`Failed to render Word page ${i + 1}:`, err);
      }
    }

    return pageElements.length;
  } finally {
    if (container.parentNode) {
      document.body.removeChild(container);
    }
  }
}

// ─── Helpers ───────────────────────────────────────────────────────────────

/**
 * Convert every blob: image to a base64 data-URL so html2canvas can capture it.
 */
async function convertAllImagesToBase64(
  container: HTMLElement
): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return;

  for (let i = 0; i < images.length; i++) {
    const img = images[i] as HTMLImageElement;
    if (!img.src || !img.src.startsWith('blob:')) continue;
    try {
      const resp = await fetch(img.src);
      const blob = await resp.blob();
      img.src = await blobToBase64(blob);
    } catch (e) {
      console.warn('Failed to convert blob image:', e);
      img.removeAttribute('src');
    }
  }

  // Let the browser repaint with new base64 sources
  await new Promise((r) => setTimeout(r, 300));
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

/**
 * Wait for every <img> in the container to finish loading (or error).
 */
async function waitForAllImages(container: HTMLElement): Promise<void> {
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

  // Extra wait for layout to settle after images load
  await new Promise((r) => setTimeout(r, 300));
}
