import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';
import { renderWordToHtml } from './word-render';

const SCALE = 3;
const RENDER_CLASS = 'wp-pdf-render';

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

  const pageWidthPx = Math.round(renderResult.pageWidthMm * 96 / 25.4);

  const captureContainer = document.createElement('div');
  captureContainer.id = 'wp-capture-container';
  captureContainer.style.cssText =
    `position:fixed;top:0;left:0;z-index:-1;pointer-events:none;background:white;width:${pageWidthPx}px;`;
  document.body.appendChild(captureContainer);

  try {
    captureContainer.innerHTML = renderResult.html;

    const wrapper = captureContainer.querySelector(
      `.${RENDER_CLASS}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.cssText =
        'background:white;padding:0;margin:0;box-shadow:none;border:none;';
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
    await new Promise((r) => setTimeout(r, 500));

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
      section.style.transform = 'none';

      try {
        const wPx = section.offsetWidth || pageWidthPx;
        const hPx = section.offsetHeight || Math.round(renderResult.pageHeightMm * 96 / 25.4);

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
            transform: 'none',
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

        let pdfWidthPt = (renderResult.pageWidthMm * 72) / 25.4;
        let pdfHeightPt = (renderResult.pageHeightMm * 72) / 25.4;

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
    if (captureContainer.parentNode) {
      document.body.removeChild(captureContainer);
    }
  }
}