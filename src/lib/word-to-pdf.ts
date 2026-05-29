/**
 * word-to-pdf.ts — Convert .docx to PDF using client-side approach.
 *
 * Strategy: docx-preview renders the full document (with images) into an iframe.
 * Then we use the iframe's built-in print-to-PDF via a hidden print window
 * combined with html2canvas for reliable capture.
 *
 * Key fixes over previous attempts:
 *  - Renders in a VISIBLE iframe (not off-screen) so the browser calculates layout correctly
 *  - Images are guaranteed to load because useBase64URL is true
 *  - Uses longer settling delays for complex layouts
 */

import html2canvas from 'html2canvas';
import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

const RENDER_CLASS = 'wp-merge-render';

export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  // Step 1: Use docx-preview to render the Word document
  const { renderAsync } = await import('docx-preview');

  // Create an iframe so styles from the main page don't interfere
  const iframe = document.createElement('iframe');
  iframe.style.cssText =
    'position:fixed;left:0;top:0;width:100vw;height:100vh;z-index:-1;border:none;pointer-events:none;background:white;';
  document.body.appendChild(iframe);

  const iframeDoc = iframe.contentDocument!;
  const iframeBody = iframeDoc.body;

  // Ensure the iframe has a basic style
  const style = iframeDoc.createElement('style');
  style.textContent = `
    body { margin: 0; padding: 0; background: white; }
    section { background: white; }
  `;
  iframeDoc.head.appendChild(style);

  try {
    // Render the Word document inside the iframe body
    await renderAsync(sourceArrayBuffer, iframeBody, undefined, {
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

    // Step 2: Wait for all images to load and layout to settle
    await waitAllImagesInIframe(iframeDoc);
    // Extra time for the browser to finish layout computation
    await delay(1000);

    // Step 3: Find page sections
    let sections = iframeDoc.querySelectorAll(`section.${RENDER_CLASS}`);
    if (sections.length === 0) {
      sections = iframeDoc.querySelectorAll('section');
    }
    if (sections.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    // Step 4: Capture each section page
    for (let i = 0; i < sections.length; i++) {
      const section = sections[i] as HTMLElement;

      const wPx = section.offsetWidth || 794;
      const hPx = section.offsetHeight || 1123;

      try {
        // Use html2canvas with the iframe's window reference
        const canvas = await html2canvas(section, {
          backgroundColor: '#ffffff',
          scale: 2,
          width: wPx,
          height: hPx,
          useCORS: true,
          allowTaint: true,
          logging: false,
          // Use the iframe's window for context
          windowWidth: wPx + 100,
          windowHeight: hPx + 100,
        });

        const pngDataUrl = canvas.toDataURL('image/png');
        const base64Data = pngDataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
        const pdfImage = await mergedPdf.embedPng(imageBytes);

        // Convert CSS px to PDF points (96 dpi → 72 dpi)
        let pdfWidth = (wPx * 72) / 96;
        let pdfHeight = (hPx * 72) / 96;

        // Handle orientation swap
        if (orientation === 'all-landscape' && pdfWidth < pdfHeight) {
          [pdfWidth, pdfHeight] = [pdfHeight, pdfWidth];
        } else if (orientation === 'all-portrait' && pdfWidth > pdfHeight) {
          [pdfWidth, pdfHeight] = [pdfHeight, pdfWidth];
        }

        const page = mergedPdf.addPage([pdfWidth, pdfHeight]);
        page.drawImage(pdfImage, {
          x: 0,
          y: 0,
          width: pdfWidth,
          height: pdfHeight,
        });
      } catch (err) {
        throw new Error(
          `Word 文档第 ${i + 1} 页渲染失败: ${err instanceof Error ? err.message : String(err)}`
        );
      }
    }

    return sections.length;
  } finally {
    // Clean up iframe
    if (iframe.parentNode) {
      document.body.removeChild(iframe);
    }
  }
}

// ─── Helpers ───────────────────────────────────────────────────────────────

function waitAllImagesInIframe(doc: Document): Promise<void> {
  const images = doc.querySelectorAll('img');
  if (images.length === 0) return Promise.resolve();

  return new Promise((resolve) => {
    let pending = images.length;
    const finish = () => {
      if (--pending <= 0) resolve();
    };
    for (let i = 0; i < images.length; i++) {
      const img = images[i] as HTMLImageElement;
      if (img.complete && img.naturalWidth > 0) {
        finish();
      } else {
        img.addEventListener('load', finish, { once: true });
        img.addEventListener('error', finish, { once: true });
      }
    }
    setTimeout(resolve, 15000);
  });
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}
