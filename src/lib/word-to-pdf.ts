import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';
import { renderWordToHtml } from './word-render';

const SCALE = 2;
const RENDER_CLASS = 'wp-pdf-render';

export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  const html2canvas = (await import('html2canvas')).default;

  const renderResult = await renderWordToHtml(sourceArrayBuffer, RENDER_CLASS);

  const container = document.createElement('div');
  container.id = 'wp-capture-container';
  container.style.cssText =
    'position:fixed;top:0;left:0;z-index:-9999;pointer-events:none;background:white;';
  document.body.appendChild(container);

  const resetStyle = document.createElement('style');
  resetStyle.id = 'wp-capture-reset';
  resetStyle.textContent =
    '#wp-capture-container img { max-width: none !important; }';
  document.head.appendChild(resetStyle);

  try {
    container.innerHTML = renderResult.html;

    const wrapper = container.querySelector(
      `.${RENDER_CLASS}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.background = 'white';
      wrapper.style.boxShadow = 'none';
    }

    await waitForImages(container);

    let sections = container.querySelectorAll(`section.${RENDER_CLASS}`);
    if (sections.length === 0) sections = container.querySelectorAll('section');
    if (sections.length === 0) throw new Error('Word 文档渲染失败：未生成任何页面');

    for (let i = 0; i < sections.length; i++) {
      const section = sections[i] as HTMLElement;

      section.style.overflow = 'hidden';
      section.style.background = '#ffffff';
      section.style.boxShadow = 'none';

      const wPx = section.offsetWidth;
      const hPx = section.offsetHeight;

      if (wPx <= 0 || hPx <= 0) continue;

      try {
        const canvas = await html2canvas(section, {
          scale: SCALE,
          backgroundColor: '#ffffff',
          useCORS: true,
          allowTaint: true,
          logging: false,
        });

        const pngDataUrl = canvas.toDataURL('image/png');
        const base64Data = pngDataUrl.split(',')[1];
        const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
        const image = await mergedPdf.embedPng(imageBytes);

        let pdfWidthPt = (wPx * 72) / 96;
        let pdfHeightPt = (hPx * 72) / 96;

        if (orientation === 'all-landscape' && pdfWidthPt < pdfHeightPt) {
          [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
        } else if (orientation === 'all-portrait' && pdfWidthPt > pdfHeightPt) {
          [pdfWidthPt, pdfHeightPt] = [pdfHeightPt, pdfWidthPt];
        }

        const page = mergedPdf.addPage([pdfWidthPt, pdfHeightPt]);
        page.drawImage(image, { x: 0, y: 0, width: pdfWidthPt, height: pdfHeightPt });
      } catch (err) {
        console.error(`Failed to capture Word page ${i + 1}:`, err);
      }
    }

    return sections.length;
  } finally {
    const styleEl = document.getElementById('wp-capture-reset');
    if (styleEl) document.head.removeChild(styleEl);
    if (container.parentNode) document.body.removeChild(container);
  }
}

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
  await new Promise((r) => setTimeout(r, 300));
}