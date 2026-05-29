import { PDFDocument, degrees } from 'pdf-lib';

export interface PageSize {
  width: number;
  height: number;
}

export const PAGE_SIZES: Record<string, PageSize> = {
  a4: { width: 595.28, height: 841.89 },
  letter: { width: 612, height: 792 },
  a3: { width: 841.89, height: 1190.55 },
};

export type OrientationMode = 'all-portrait' | 'all-landscape' | 'keep-original';

export interface ProgressCallback {
  (fileName: string, percent: number): void;
}

/**
 * Add PDF pages from a source PDF to the merged PDF document.
 * Handles rotation for orientation changes.
 */
export async function addPdfPages(
  mergedPdf: PDFDocument,
  sourceArrayBuffer: ArrayBuffer,
  orientation: OrientationMode
): Promise<number> {
  const sourcePdf = await PDFDocument.load(sourceArrayBuffer);
  const pages = sourcePdf.getPages();
  const totalPages = pages.length;

  const pageIndices = Array.from({ length: totalPages }, (_, i) => i);
  const copiedPages = await mergedPdf.copyPages(sourcePdf, pageIndices);

  for (let i = 0; i < copiedPages.length; i++) {
    const copiedPage = copiedPages[i];
    const originalPage = pages[i];

    let pageWidth = originalPage.getWidth();
    let pageHeight = originalPage.getHeight();
    const existingRotation = originalPage.getRotation().angle;

    let effectiveWidth = pageWidth;
    let effectiveHeight = pageHeight;

    if (existingRotation === 90 || existingRotation === 270) {
      [effectiveWidth, effectiveHeight] = [effectiveHeight, effectiveWidth];
    }

    const isLandscape = effectiveWidth > effectiveHeight;
    const addedPage = mergedPdf.addPage(copiedPage);

    let needRotate = false;
    if (orientation === 'all-portrait' && isLandscape) {
      needRotate = true;
    } else if (orientation === 'all-landscape' && !isLandscape) {
      needRotate = true;
    }

    if (needRotate) {
      const newRotation = (existingRotation + 90) % 360;
      addedPage.setRotation(degrees(newRotation));
    }
  }

  return totalPages;
}

/**
 * Add an image file as a single PDF page.
 * Handles sizing, centering, and orientation rotation.
 */
export async function addImagePage(
  mergedPdf: PDFDocument,
  file: File,
  imgWidth: number,
  imgHeight: number,
  mimeType: string,
  pageSize: string,
  orientation: OrientationMode,
  imageQuality: number
): Promise<void> {
  const isLandscape = imgWidth > imgHeight;
  let needRotate = false;

  if (orientation === 'all-portrait' && isLandscape) {
    needRotate = true;
  } else if (orientation === 'all-landscape' && !isLandscape) {
    needRotate = true;
  }

  const displayWidth = needRotate ? imgHeight : imgWidth;
  const displayHeight = needRotate ? imgWidth : imgHeight;

  let pageWidth: number, pageHeight: number, drawWidth: number, drawHeight: number;

  if (pageSize === 'auto') {
    pageWidth = displayWidth * 72 / 96;
    pageHeight = displayHeight * 72 / 96;
    drawWidth = pageWidth;
    drawHeight = pageHeight;
  } else {
    const size = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
    pageWidth = size.width;
    pageHeight = size.height;

    const scaleX = pageWidth / displayWidth;
    const scaleY = pageHeight / displayHeight;
    const scale = Math.min(scaleX, scaleY, 1);

    drawWidth = displayWidth * scale;
    drawHeight = displayHeight * scale;
  }

  const page = mergedPdf.addPage([pageWidth, pageHeight]);

  let image;
  try {
    const imageBytes = await file.arrayBuffer();
    if (mimeType === 'image/jpeg' || mimeType === 'image/jpg') {
      image = await mergedPdf.embedJpg(new Uint8Array(imageBytes));
    } else if (mimeType === 'image/png') {
      image = await mergedPdf.embedPng(new Uint8Array(imageBytes));
    } else {
      // For GIF, BMP, WebP — convert via canvas
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d')!;
      canvas.width = imgWidth;
      canvas.height = imgHeight;

      const img = new Image();
      // We need the data URL; it should be already available in ProcessedFile.preview
      // But we're passing the file, so we read it as data URL
      const dataUrl = await new Promise<string>((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target?.result as string);
        reader.readAsDataURL(file);
      });

      await new Promise<void>((resolve, reject) => {
        img.onload = () => {
          ctx.drawImage(img, 0, 0);
          resolve();
        };
        img.onerror = reject;
        img.src = dataUrl;
      });

      const pngData = canvas.toDataURL('image/png', imageQuality);
      const base64 = pngData.split(',')[1];
      const pngBytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      image = await mergedPdf.embedPng(pngBytes);
    }
  } catch (error) {
    console.error('Image embedding error:', error);
    return;
  }

  const x = (pageWidth - drawWidth) / 2;
  const y = (pageHeight - drawHeight) / 2;

  page.drawImage(image, {
    x,
    y,
    width: drawWidth,
    height: drawHeight,
  });

  if (needRotate) {
    page.setRotation(degrees(90));
  }
}
