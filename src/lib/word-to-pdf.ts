import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const SCALE = 2; // Render at 2x for crisp PDF output
const CLASS_NAME = 'docx-word-merge';

/**
 * Render a Word document (.docx) to PDF pages using docx-preview for high-fidelity output.
 * Flow: docx-preview renderAsync → convert images → per-page html2canvas → embed PNG into pdf-lib
 */
export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  // Dynamic imports (browser-only)
  const { renderAsync } = await import('docx-preview');
  const html2canvas = (await import('html2canvas')).default;

  // 1. Create a render container that is IN the viewport
  // Position it covering the entire viewport but behind everything else.
  // html2canvas needs elements to be within the viewport for accurate rendering.
  const container = document.createElement('div');
  container.id = `${CLASS_NAME}-container`;
  container.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    min-width: 1200px;
    height: 100vh;
    z-index: -1;
    pointer-events: none;
    overflow: auto;
    background: white;
  `;
  document.body.appendChild(container);

  // Save current scroll position to restore later
  const savedScrollX = window.scrollX;
  const savedScrollY = window.scrollY;

  try {
    // 2. Render the Word document with docx-preview
    // styleContainer=null → docx-preview injects <style> elements directly into bodyContainer
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
    // docx-preview DOM: container > div.{className}-wrapper > section.{className}
    let pageElements = container.querySelectorAll(`section.${CLASS_NAME}`);
    if (pageElements.length === 0) {
      pageElements = container.querySelectorAll('section');
    }
    if (pageElements.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // 4. Ensure all images are fully loaded and convert any blob URLs to base64
    // (some browsers might not resolve blob URLs for html2canvas)
    await ensureAllImagesLoaded(container);

    // 5. Allow browser to complete layout and paint
    await new Promise((r) => setTimeout(r, 500));

    // Scroll to top to ensure html2canvas captures from the correct position
    window.scrollTo(0, 0);

    // 6. Render each page section to canvas and embed into PDF
    for (let i = 0; i < pageElements.length; i++) {
      const pageEl = pageElements[i] as HTMLElement;

      // Scroll the page element into view within the container
      pageEl.scrollIntoView({ behavior: 'instant', block: 'start' });
      await new Promise((r) => setTimeout(r, 100));

      const canvas = await html2canvas(pageEl, {
        scale: SCALE,
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        logging: false,
        scrollX: 0,
        scrollY: 0,
        imageTimeout: 15000,
      });

      if (canvas.width === 0 || canvas.height === 0) continue;

      // Convert canvas to PNG and embed
      const pngDataUrl = canvas.toDataURL('image/png');
      const base64Data = pngDataUrl.split(',')[1];
      const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
      const image = await mergedPdf.embedPng(imageBytes);

      // Calculate PDF page dimensions from rendered element size
      // Screen: 96dpi, PDF: 72dpi → multiply by 72/96
      const renderedWidthPx = pageEl.offsetWidth;
      const renderedHeightPx = pageEl.offsetHeight;
      const pageWidthPt = (renderedWidthPx * 72) / 96;
      const pageHeightPt = (renderedHeightPx * 72) / 96;

      // Handle orientation override
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
    // 7. Restore scroll position and clean up
    window.scrollTo(savedScrollX, savedScrollY);
    if (container.parentNode) {
      document.body.removeChild(container);
    }
  }
}

/**
 * Ensure all images in the container are fully loaded.
 * Converts any remaining blob URLs to base64 data URLs for html2canvas compatibility.
 */
async function ensureAllImagesLoaded(container: HTMLElement): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return;

  // First pass: convert any blob URLs to base64 (html2canvas handles base64 better)
  for (let i = 0; i < images.length; i++) {
    const img = images[i] as HTMLImageElement;
    if (img.src && img.src.startsWith('blob:')) {
      try {
        const response = await fetch(img.src);
        const blob = await response.blob();
        img.src = await blobToBase64(blob);
      } catch (e) {
        console.warn('Failed to convert blob image to base64:', e);
      }
    }
  }

  // Second pass: wait for all images to complete loading
  const loadPromises = Array.from(images).map((img) => {
    const el = img as HTMLImageElement;
    return new Promise<void>((resolve) => {
      if (el.complete && el.naturalWidth > 0) {
        resolve();
        return;
      }
      const onDone = () => {
        el.removeEventListener('load', onDone);
        el.removeEventListener('error', onDone);
        resolve();
      };
      el.addEventListener('load', onDone, { once: false });
      el.addEventListener('error', onDone, { once: false });
    });
  });

  // Wait with a safety timeout
  await Promise.race([
    Promise.all(loadPromises),
    new Promise((r) => setTimeout(r, 8000)),
  ]);
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}
