import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';
import { PAGE_SIZES } from './merge-pdf';

const PX_PER_POINT = 96 / 72;
const SCALE = 2;
const PADDING_PX = 40;

// Unique class name to scope Word render styles globally
const RENDER_CLS = '__word_render_container__';

// Inject global styles once (idempotent)
let stylesInjected = false;
function ensureGlobalStyles() {
  if (stylesInjected) return;
  stylesInjected = true;
  const s = document.createElement('style');
  s.id = '__word-render-styles__';
  s.textContent = `
    .${RENDER_CLS} { font-family: 'SimSun', 'Songti SC', 'Noto Serif SC', 'Times New Roman', serif; font-size: 14px; line-height: 1.8; color: #000; }
    .${RENDER_CLS} h1 { font-size: 22px; font-weight: bold; margin: 16px 0 8px; }
    .${RENDER_CLS} h2 { font-size: 18px; font-weight: bold; margin: 14px 0 6px; }
    .${RENDER_CLS} h3 { font-size: 16px; font-weight: bold; margin: 12px 0 4px; }
    .${RENDER_CLS} p { margin: 6px 0; text-indent: 0; }
    .${RENDER_CLS} table { border-collapse: collapse; width: 100%; margin: 10px 0; }
    .${RENDER_CLS} table td, .${RENDER_CLS} table th { border: 1px solid #ccc; padding: 6px 10px; }
    .${RENDER_CLS} img { max-width: 100%; height: auto; }
    .${RENDER_CLS} ul, .${RENDER_CLS} ol { margin-left: 20px; padding-left: 0; }
    .${RENDER_CLS} li { margin: 4px 0; }
    .${RENDER_CLS} strong, .${RENDER_CLS} b { font-weight: bold; }
    .${RENDER_CLS} em, .${RENDER_CLS} i { font-style: italic; }
    .${RENDER_CLS} a { color: #0066cc; text-decoration: underline; }
    .${RENDER_CLS} blockquote { border-left: 3px solid #ccc; margin: 10px 0; padding-left: 12px; color: #555; }
  `;
  document.head.appendChild(s);
}

/**
 * Convert a Word document's HTML content to PDF pages and add them to the merged PDF.
 * Uses mammoth → HTML → html2canvas → canvas slicing → PNG embed approach.
 */
export async function addWordPages(
  mergedPdf: PDFDocument,
  htmlContent: string,
  pageSize: string,
  orientation: OrientationMode
): Promise<number> {
  // Dynamic import for browser-only library
  const html2canvas = (await import('html2canvas')).default;

  // Ensure global styles are injected
  ensureGlobalStyles();

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

  // Create hidden render container — fixed position so it's in the layout but not visible
  const container = document.createElement('div');
  container.className = RENDER_CLS;
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    zIndex: '-9999',
    overflow: 'visible',  // visible so html2canvas captures full content
    width: containerWidth + 'px',
    padding: PADDING_PX + 'px',
    background: '#ffffff',
    boxSizing: 'border-box',
  });

  // Insert HTML content
  container.innerHTML = htmlContent || '<p style="color:#999;">（文档内容为空）</p>';
  document.body.appendChild(container);

  // Wait for layout + images
  await new Promise((resolve) => setTimeout(resolve, 800));

  try {
    // Render to canvas
    const fullCanvas = await html2canvas(container, {
      scale: SCALE,
      useCORS: true,
      allowTaint: true,
      backgroundColor: '#ffffff',
      logging: false,
      // Ensure the full content height is captured
      windowWidth: containerWidth + PADDING_PX * 2,
      windowHeight: container.scrollHeight + PADDING_PX * 2,
    });

    if (fullCanvas.width === 0 || fullCanvas.height === 0) {
      throw new Error('Word 渲染结果为空，文档内容可能无法解析');
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
    if (container.parentNode) {
      document.body.removeChild(container);
    }
  }
}
