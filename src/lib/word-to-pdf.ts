import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';
import { PAGE_SIZES } from './merge-pdf';

const PX_PER_POINT = 96 / 72;
const SCALE = 2;
const PADDING_PX = 40;

/**
 * Convert a Word document's HTML content to PDF pages and add them to the merged PDF.
 * Uses mammoth + html2canvas + canvas slicing approach.
 */
export async function addWordPages(
  mergedPdf: PDFDocument,
  htmlContent: string,
  pageSize: string,
  orientation: OrientationMode
): Promise<number> {
  // Dynamic imports for browser-only libraries
  const html2canvas = (await import('html2canvas')).default;

  // Determine page dimensions
  let targetWidth: number, targetHeight: number;
  if (pageSize === 'auto') {
    targetWidth = PAGE_SIZES.a4.width;
    targetHeight = PAGE_SIZES.a4.height;
  } else {
    const size = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
    targetWidth = size.width;
    targetHeight = size.height;
  }

  if (orientation === 'all-landscape') {
    [targetWidth, targetHeight] = [targetHeight, targetWidth];
  }

  const containerWidth = Math.round(targetWidth * PX_PER_POINT);
  const containerHeight = Math.round(targetHeight * PX_PER_POINT);

  // Create hidden render container
  const container = document.createElement('div');
  container.setAttribute('data-word-render', 'true');
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    zIndex: '-9999',
    overflow: 'hidden',
    width: containerWidth + 'px',
    padding: PADDING_PX + 'px',
    background: '#ffffff',
    color: '#000000',
    fontFamily: "'SimSun', 'Songti SC', 'Noto Serif SC', 'Times New Roman', serif",
    fontSize: '14px',
    lineHeight: '1.8',
  });

  // Style inner elements
  const styleEl = document.createElement('style');
  styleEl.textContent = `
    [data-word-render] h1 { font-size: 22px; font-weight: bold; margin: 16px 0 8px; }
    [data-word-render] h2 { font-size: 18px; font-weight: bold; margin: 14px 0 6px; }
    [data-word-render] h3 { font-size: 16px; font-weight: bold; margin: 12px 0 4px; }
    [data-word-render] p { margin: 6px 0; }
    [data-word-render] table { border-collapse: collapse; width: 100%; margin: 10px 0; }
    [data-word-render] table td, [data-word-render] table th {
      border: 1px solid #ccc;
      padding: 6px 10px;
    }
    [data-word-render] img { max-width: 100%; height: auto; }
    [data-word-render] ul, [data-word-render] ol { margin-left: 20px; }
    [data-word-render] li { margin: 4px 0; }
  `;
  container.appendChild(styleEl);
  container.insertAdjacentHTML('beforeend', htmlContent || '<p style="color:#999;">（文档内容为空）</p>');
  document.body.appendChild(container);

  // Wait for images to load
  await new Promise((resolve) => setTimeout(resolve, 500));

  try {
    // Render to canvas
    const fullCanvas = await html2canvas(container, {
      scale: SCALE,
      useCORS: true,
      allowTaint: true,
      backgroundColor: '#ffffff',
      logging: false,
    });

    if (fullCanvas.width === 0 || fullCanvas.height === 0) {
      throw new Error('Word rendering produced empty canvas');
    }

    // Calculate page slicing
    const pageContentHeight = containerHeight - PADDING_PX * 2;
    const canvasPageHeight = Math.round(pageContentHeight * SCALE);
    const canvasPageWidth = fullCanvas.width;
    const totalHeight = fullCanvas.height;

    // Slice into pages
    let pageCount = 0;
    let yOffset = 0;

    while (yOffset < totalHeight) {
      const sliceHeight = Math.min(canvasPageHeight, totalHeight - yOffset);

      const pageCanvas = document.createElement('canvas');
      pageCanvas.width = canvasPageWidth;
      pageCanvas.height = sliceHeight;
      const ctx = pageCanvas.getContext('2d')!;

      // White background
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvasPageWidth, sliceHeight);

      // Slice from full canvas
      ctx.drawImage(
        fullCanvas,
        0, yOffset, canvasPageWidth, sliceHeight,
        0, 0, canvasPageWidth, sliceHeight
      );

      // Convert to PNG and embed
      const pngDataUrl = pageCanvas.toDataURL('image/png');
      const base64Data = pngDataUrl.split(',')[1];
      const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
      const image = await mergedPdf.embedPng(imageBytes);

      const paddingPt = PADDING_PX / PX_PER_POINT;
      const pdfPageWidth = targetWidth;
      const pdfPageHeight = (sliceHeight / SCALE / PX_PER_POINT) + paddingPt * 2;

      const page = mergedPdf.addPage([pdfPageWidth, Math.max(pdfPageHeight, targetHeight)]);
      page.drawImage(image, {
        x: 0,
        y: 0,
        width: pdfPageWidth,
        height: pdfPageHeight,
      });

      pageCount++;
      yOffset += canvasPageHeight;
    }

    return pageCount;
  } finally {
    // Clean up container
    document.body.removeChild(container);
  }
}
