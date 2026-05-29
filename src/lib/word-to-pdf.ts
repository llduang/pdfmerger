// ⬇️ 把这个换成你第二步拿到的 Worker 地址
const CONVERT_API = 'https://pdf-convert.2710476780.workers.dev/';

import { PDFDocument, degrees } from 'pdf-lib';
import type { OrientationMode } from './merge-pdf';

export async function addWordPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  const formData = new FormData();
  formData.append(
    'file',
    new Blob([sourceArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }),
    'document.docx'
  );

  let response: Response;
  try {
    response = await fetch(CONVERT_API, { method: 'POST', body: formData });
  } catch {
    throw new Error('转换服务连接失败，请检查网络后重试');
  }

  if (!response.ok) {
    throw new Error('Word 转换失败，请稍后重试');
  }

  const pdfBytes = new Uint8Array(await response.arrayBuffer());
  const sourcePdf = await PDFDocument.load(pdfBytes);
  const copiedPages = await mergedPdf.copyPages(sourcePdf, sourcePdf.getPageIndices());

  for (const page of copiedPages) {
    mergedPdf.addPage(page);
  }

  if (orientation !== 'keep-original') {
    const total = mergedPdf.getPageCount();
    const start = total - copiedPages.length;
    for (let i = start; i < total; i++) {
      const page = mergedPdf.getPage(i);
      const { width, height } = page.getSize();
      const needRotate =
        (orientation === 'all-portrait' && width > height) ||
        (orientation === 'all-landscape' && width <= height);
      if (needRotate) {
        const rot = page.getRotation().angle;
        page.setRotation(degrees((rot + 90) % 360));
        const box = page.getMediaBox();
        page.setMediaBox(box.y, box.x, box.height, box.width);
      }
    }
  }

  return copiedPages.length;
}