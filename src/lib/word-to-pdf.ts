import { PDFDocument } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';
import { PAGE_SIZES } from './merge-pdf';

const PX_PER_POINT = 96 / 72;
const SCALE = 2;
const MARGIN_PT = 56; // ~2cm margin in PDF points
const MARGIN_PX = Math.round(MARGIN_PT * PX_PER_POINT);

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
    .${RENDER_CLS} {
      font-family: 'SimSun', 'Songti SC', 'Noto Serif SC', 'Microsoft YaHei', 'Times New Roman', serif;
      font-size: 14px;
      line-height: 1.8;
      color: #000;
      word-wrap: break-word;
      overflow-wrap: break-word;
    }
    .${RENDER_CLS} h1 {
      font-size: 24px;
      font-weight: bold;
      margin: 20px 0 10px;
      line-height: 1.4;
      color: #000;
    }
    .${RENDER_CLS} h2 {
      font-size: 20px;
      font-weight: bold;
      margin: 16px 0 8px;
      line-height: 1.4;
      color: #000;
    }
    .${RENDER_CLS} h3 {
      font-size: 17px;
      font-weight: bold;
      margin: 14px 0 6px;
      line-height: 1.5;
      color: #000;
    }
    .${RENDER_CLS} h4 {
      font-size: 15px;
      font-weight: bold;
      margin: 12px 0 4px;
      line-height: 1.5;
      color: #000;
    }
    .${RENDER_CLS} p {
      margin: 8px 0;
      text-indent: 0;
      line-height: 1.8;
    }
    .${RENDER_CLS} table {
      border-collapse: collapse;
      width: 100%;
      margin: 12px 0;
      font-size: 13px;
    }
    .${RENDER_CLS} table td,
    .${RENDER_CLS} table th {
      border: 1px solid #999;
      padding: 6px 10px;
      vertical-align: top;
    }
    .${RENDER_CLS} table th {
      background-color: #f0f0f0;
      font-weight: bold;
    }
    .${RENDER_CLS} img {
      max-width: 100%;
      height: auto;
      display: block;
      margin: 8px auto;
    }
    .${RENDER_CLS} ul,
    .${RENDER_CLS} ol {
      margin: 8px 0;
      padding-left: 24px;
    }
    .${RENDER_CLS} li {
      margin: 4px 0;
      line-height: 1.8;
    }
    .${RENDER_CLS} strong,
    .${RENDER_CLS} b {
      font-weight: bold;
    }
    .${RENDER_CLS} em,
    .${RENDER_CLS} i {
      font-style: italic;
    }
    .${RENDER_CLS} u {
      text-decoration: underline;
    }
    .${RENDER_CLS} a {
      color: #0066cc;
      text-decoration: underline;
    }
    .${RENDER_CLS} blockquote {
      border-left: 3px solid #ccc;
      margin: 10px 0;
      padding: 8px 16px;
      color: #555;
      background: #fafafa;
    }
    .${RENDER_CLS} pre {
      background: #f5f5f5;
      border: 1px solid #ddd;
      padding: 10px;
      border-radius: 4px;
      font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
      font-size: 13px;
      white-space: pre-wrap;
      word-wrap: break-word;
      margin: 10px 0;
      overflow-x: auto;
    }
    .${RENDER_CLS} code {
      font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
      font-size: 13px;
      background: #f5f5f5;
      padding: 1px 4px;
      border-radius: 3px;
    }
    .${RENDER_CLS} hr {
      border: none;
      border-top: 1px solid #ccc;
      margin: 16px 0;
    }
    .${RENDER_CLS} sup {
      font-size: 0.75em;
      vertical-align: super;
    }
    .${RENDER_CLS} sub {
      font-size: 0.75em;
      vertical-align: sub;
    }
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

  // Determine page dimensions in PDF points
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

  // The content area width (page width minus left+right margins)
  const contentWidthPt = targetWidth - MARGIN_PT * 2;
  const contentWidthPx = Math.round(contentWidthPt * PX_PER_POINT);
  const pageContentHeightPx = Math.round((targetHeight - MARGIN_PT * 2) * PX_PER_POINT);

  // Create hidden render container
  const container = document.createElement('div');
  container.className = RENDER_CLS;
  Object.assign(container.style, {
    position: 'fixed',
    top: '0',
    left: '0',
    zIndex: '-9999',
    overflow: 'visible',
    width: contentWidthPx + 'px',
    padding: '0',
    margin: '0',
    background: '#ffffff',
    boxSizing: 'border-box',
  });

  // Insert HTML content
  container.innerHTML = htmlContent || '<p style="color:#999;">（文档内容为空）</p>';
  document.body.appendChild(container);

  // Wait for layout + images to fully render
  await new Promise((resolve) => setTimeout(resolve, 1000));

  try {
    // Get the actual rendered dimensions
    const scrollHeight = container.scrollHeight;

    if (scrollHeight === 0) {
      throw new Error('Word 渲染结果为空，文档内容可能无法解析');
    }

    // Render to canvas at full content height
    const fullCanvas = await html2canvas(container, {
      scale: SCALE,
      useCORS: true,
      allowTaint: true,
      backgroundColor: '#ffffff',
      logging: false,
      // Capture the full content
      windowWidth: contentWidthPx,
      windowHeight: scrollHeight,
    });

    if (fullCanvas.width === 0 || fullCanvas.height === 0) {
      throw new Error('Word canvas 渲染结果为空');
    }

    const totalCanvasHeight = fullCanvas.height;
    const canvasWidth = fullCanvas.width;

    // Slice into pages
    const canvasPageHeight = Math.round(pageContentHeightPx * SCALE);
    let pageCount = 0;
    let yOffset = 0;

    while (yOffset < totalCanvasHeight) {
      const sliceHeight = Math.min(canvasPageHeight, totalCanvasHeight - yOffset);

      const pageCanvas = document.createElement('canvas');
      pageCanvas.width = canvasWidth;
      pageCanvas.height = sliceHeight;
      const ctx = pageCanvas.getContext('2d')!;

      // White background for the page
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvasWidth, sliceHeight);

      // Slice content from full canvas
      ctx.drawImage(
        fullCanvas,
        0, yOffset, canvasWidth, sliceHeight,
        0, 0, canvasWidth, sliceHeight
      );

      // Convert to PNG and embed into PDF
      const pngDataUrl = pageCanvas.toDataURL('image/png');
      const base64Data = pngDataUrl.split(',')[1];
      const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
      const image = await mergedPdf.embedPng(imageBytes);

      // Calculate PDF dimensions for this page
      const pdfPageWidth = targetWidth;
      const pdfPageHeight = targetHeight;

      const page = mergedPdf.addPage([pdfPageWidth, pdfPageHeight]);

      // Draw the image with margins
      const imgWidth = contentWidthPt;
      const imgHeight = (sliceHeight / SCALE / PX_PER_POINT);

      page.drawImage(image, {
        x: MARGIN_PT,
        y: pdfPageHeight - MARGIN_PT - imgHeight,
        width: imgWidth,
        height: imgHeight,
      });

      pageCount++;
      yOffset += canvasPageHeight;
    }

    return pageCount;
  } finally {
    // Clean up the render container
    if (container.parentNode) {
      document.body.removeChild(container);
    }
  }
}
