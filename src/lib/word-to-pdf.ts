import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const SCALE = 2;
const CLASS_NAME = 'docx-word-merge';

/**
 * Render a Word document (.docx) to PDF pages using docx-preview + html2canvas.
 *
 * Note: This is the "automatic download" path. It uses html2canvas which has known
 * limitations with images and complex CSS. For PERFECT Word rendering, use the
 * PrintPreview component which leverages the browser's native print engine.
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

    // Wait for images to load
    await ensureAllImagesLoaded(container);
    await new Promise((r) => setTimeout(r, 800));

    // Clean up wrapper styles for cleaner screenshot
    const wrapper = container.querySelector(`.${CLASS_NAME}-wrapper`) as HTMLElement;
    if (wrapper) {
      wrapper.style.background = 'white';
      wrapper.style.padding = '0';
      wrapper.style.margin = '0';
    }

    // Render each page to canvas and embed
    for (let i = 0; i < pageElements.length; i++) {
      const pageEl = pageElements[i] as HTMLElement;
      pageEl.style.background = '#ffffff';
      pageEl.style.boxShadow = 'none';
      pageEl.style.margin = '0';

      try {
        const canvas = await html2canvas(pageEl, {
          scale: SCALE,
          useCORS: true,
          allowTaint: true,
          backgroundColor: '#ffffff',
          logging: false,
          imageTimeout: 15000,
          onclone: (clonedDoc) => {
            // Fix images in the cloned document
            const clonedEl = clonedDoc.querySelector(
              `[data-html2canvas-id="${pageEl.getAttribute('data-html2canvas-id')}"]`
            ) || pageEl;
            const imgs = (clonedEl as HTMLElement).querySelectorAll('img');
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
        const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
        const image = await mergedPdf.embedPng(imageBytes);

        const renderedWidthPx = pageEl.offsetWidth;
        const renderedHeightPx = pageEl.offsetHeight;
        const pageWidthPt = (renderedWidthPx * 72) / 96;
        const pageHeightPt = (renderedHeightPx * 72) / 96;

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

async function ensureAllImagesLoaded(container: HTMLElement): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return;

  for (let i = 0; i < images.length; i++) {
    const img = images[i] as HTMLImageElement;
    if (img.src?.startsWith('blob:')) {
      try {
        const resp = await fetch(img.src);
        const blob = await resp.blob();
        img.src = await blobToBase64(blob);
      } catch (e) {
        console.warn('Failed to convert blob image:', e);
      }
    }
  }

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

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
