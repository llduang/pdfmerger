import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const RENDER_CLASS = 'wp-pdf-render';

/**
 * Convert a .docx file to PDF pages.
 *
 * KEY FIX: Render with docx-preview AND capture with html2canvas
 * in the SAME container — no HTML transfer between containers.
 * Transferring innerHTML loses image positioning because
 * docx-preview uses inline styles and absolute positioning
 * that don't survive DOM serialization round-trips.
 */
export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  const { renderAsync } = await import('docx-preview');
  const { default: html2canvas } = await import('html2canvas');

  // ── Step 1: Render directly into the capture container ──
  // NO width constraint — let docx-preview use its own section sizes.
  const container = document.createElement('div');
  container.id = 'wp-capture-container';
  container.style.cssText =
    'position:fixed;top:0;left:0;z-index:-1;pointer-events:none;background:white;';
  document.body.appendChild(container);

  try {
    await renderAsync(sourceArrayBuffer, container, undefined, {
      className: RENDER_CLASS,
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

    // ── Step 2: Convert blob images to base64 ──
    const blobImages = container.querySelectorAll('img');
    for (let i = 0; i < blobImages.length; i++) {
      const img = blobImages[i] as HTMLImageElement;
      if (!img.src || !img.src.startsWith('blob:')) continue;
      try {
        const resp = await fetch(img.src);
        const blob = await resp.blob();
        const reader = new FileReader();
        const base64 = await new Promise<string>((resolve, reject) => {
          reader.onload = () => resolve(reader.result as string);
          reader.onerror = reject;
          reader.readAsDataURL(blob);
        });
        img.src = base64;
      } catch {
        img.removeAttribute('src');
      }
    }

    // ── Step 3: Wait for all images to load + layout settle ──
    await new Promise<void>((resolve) => {
      const images = container.querySelectorAll('img');
      if (images.length === 0) { resolve(); return; }
      let pending = images.length;
      const finish = () => { if (--pending <= 0) resolve(); };
      for (let i = 0; i < images.length; i++) {
        const img = images[i] as HTMLImageElement;
        if (img.complete && img.naturalWidth > 0) finish();
        else {
          img.addEventListener('load', finish, { once: true });
          img.addEventListener('error', finish, { once: true });
        }
      }
      setTimeout(resolve, 15000);
    });
    // Extra wait for layout reflow after images change src
    await new Promise((r) => setTimeout(r, 1000));

    // ── Step 4: Clean wrapper/section styles ──
    const wrapper = container.querySelector(
      `.${RENDER_CLASS}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.cssText =
        'background:white;padding:0;margin:0;box-shadow:none;border:none;';
    }

    let sections = container.querySelectorAll(
      `section.${RENDER_CLASS}`
    );
    if (sections.length === 0) {
      sections = container.querySelectorAll('section');
    }
    if (sections.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // ── Step 5: Get page dimensions from first section ──
    const first = sections[0] as HTMLElement;
    const pageWidthMm = Math.ceil(first.offsetWidth * 25.4 / 96);
    const pageHeightMm = Math.ceil(first.offsetHeight * 25.4 / 96);

    // ── Step 6: Capture each section DIRECTLY (no DOM transfer) ──
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i] as HTMLElement;

      section.style.background = '#ffffff';
      section.style.boxShadow = 'none';

      try {
        const wPx = section.offsetWidth;
        const hPx = section.offsetHeight;

        const canvas = await html2canvas(section, {
          backgroundColor: '#ffffff',
          scale: 2,
          useCORS: true,
          allowTaint: true,
          width: wPx,
          height: hPx,
          // Use the section's own width as the virtual viewport
          windowWidth: wPx,
        });

        const pngDataUrl = canvas.toDataURL('image/png');
        const base64Data = pngDataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) =>
          c.charCodeAt(0)
        );
        const image = await mergedPdf.embedPng(imageBytes);

        // Calculate PDF page size from Word document dimensions
        let pdfWidthPt = (pageWidthMm * 72) / 25.4;
        let pdfHeightPt = (pageHeightMm * 72) / 25.4;

        let drawWidth: number;
        let drawHeight: number;

        if (orientation === 'all-landscape' && pdfWidthPt < pdfHeightPt) {
          drawWidth = pdfWidthPt;
          drawHeight = pdfHeightPt;
          [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
          const scale = Math.min(pdfWidthPt / drawWidth, pdfHeightPt / drawHeight, 1);
          drawWidth *= scale;
          drawHeight *= scale;
        } else if (orientation === 'all-portrait' && pdfWidthPt > pdfHeightPt) {
          drawWidth = pdfWidthPt;
          drawHeight = pdfHeightPt;
          [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
          const scale = Math.min(pdfWidthPt / drawWidth, pdfHeightPt / drawHeight, 1);
          drawWidth *= scale;
          drawHeight *= scale;
        } else {
          drawWidth = pdfWidthPt;
          drawHeight = pdfHeightPt;
        }

        const page = mergedPdf.addPage([pdfWidthPt, pdfHeightPt]);
        const x = (pdfWidthPt - drawWidth) / 2;
        const y = (pdfHeightPt - drawHeight) / 2;
        page.drawImage(image, { x, y, width: drawWidth, height: drawHeight });
      } catch (err) {
        console.error(`Failed to capture Word page ${i + 1}:`, err);
      }
    }

    return sections.length;
  } finally {
    if (container.parentNode) {
      document.body.removeChild(container);
    }
  }
}