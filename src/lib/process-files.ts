import { PDFDocument } from 'pdf-lib';

export type FileCategory = 'PDF' | 'Word' | 'Image';
export type Orientation = 'portrait' | 'landscape';

export interface ProcessedFile {
  id: string;
  file: File;
  name: string;
  type: string;
  size: number;
  category: FileCategory;
  preview: string | null;
  orientation: Orientation;
  pages: number;
  width?: number;
  height?: number;
  arrayBuffer?: ArrayBuffer;
}

export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

export function getFileCategory(fileName: string, mimeType: string): FileCategory | null {
  const isDocx = fileName.toLowerCase().endsWith('.docx');
  if (mimeType === 'application/pdf') return 'PDF';
  if (isDocx || mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') return 'Word';
  if (mimeType.startsWith('image/')) return 'Image';
  return null;
}

export async function processPdfFile(fileData: ProcessedFile): Promise<void> {
  const arrayBuffer = await fileData.file.arrayBuffer();
  const pdfDoc = await PDFDocument.load(arrayBuffer);
  const pages = pdfDoc.getPages();
  fileData.pages = pages.length;

  const firstPage = pages[0];
  const width = firstPage.getWidth();
  const height = firstPage.getHeight();

  const rotation = firstPage.getRotation().angle;
  let effectiveWidth = width;
  let effectiveHeight = height;
  if (rotation === 90 || rotation === 270) {
    [effectiveWidth, effectiveHeight] = [effectiveHeight, effectiveWidth];
  }

  fileData.orientation = effectiveWidth > effectiveHeight ? 'landscape' : 'portrait';
  fileData.width = effectiveWidth;
  fileData.height = effectiveHeight;
  fileData.arrayBuffer = arrayBuffer;
  fileData.preview = 'pdf';
}

/**
 * Process Word file — render it to detect page count and orientation.
 * Uses docx-preview for quick preview rendering (not full merge-quality).
 */
export async function processWordFile(fileData: ProcessedFile): Promise<void> {
  const arrayBuffer = await fileData.file.arrayBuffer();
  fileData.arrayBuffer = arrayBuffer;
  fileData.category = 'Word';

  try {
    const { renderAsync } = await import('docx-preview');

    const container = document.createElement('div');
    container.style.cssText =
      'position:fixed;left:-9999px;top:0;width:100vw;pointer-events:none;';
    document.body.appendChild(container);

    try {
      await renderAsync(arrayBuffer, container, null, {
        className: 'wp-detect-' + fileData.id,
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        breakPages: true,
        useBase64URL: true,
      });

      // Convert blob images to base64 for reliable measurement
      const images = container.querySelectorAll('img');
      for (let i = 0; i < images.length; i++) {
        const img = images[i] as HTMLImageElement;
        if (img.src && img.src.startsWith('blob:')) {
          try {
            const resp = await fetch(img.src);
            const blob = await resp.blob();
            img.src = await new Promise<string>((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => resolve(reader.result as string);
              reader.onerror = reject;
              reader.readAsDataURL(blob);
            });
          } catch {
            // skip broken images
          }
        }
      }

      // Wait for images to load
      await new Promise<void>((resolve) => {
        const imgs = container.querySelectorAll('img');
        if (imgs.length === 0) { resolve(); return; }
        let pending = imgs.length;
        const done = () => { if (--pending <= 0) resolve(); };
        for (let i = 0; i < imgs.length; i++) {
          const img = imgs[i] as HTMLImageElement;
          if (img.complete && img.naturalWidth > 0) { done(); }
          else {
            img.addEventListener('load', done, { once: true });
            img.addEventListener('error', done, { once: true });
          }
        }
        setTimeout(resolve, 8000);
      });

      await new Promise((r) => setTimeout(r, 300));

      const sections = container.querySelectorAll('section');
      if (sections.length > 0) {
        fileData.pages = sections.length;
        const first = sections[0] as HTMLElement;
        const wMm = Math.ceil(first.offsetWidth * 25.4 / 96);
        const hMm = Math.ceil(first.offsetHeight * 25.4 / 96);
        fileData.orientation = wMm > hMm ? 'landscape' : 'portrait';
      } else {
        fileData.pages = 1;
        fileData.orientation = 'portrait';
      }
    } finally {
      document.body.removeChild(container);
    }

    fileData.preview = 'word';
  } catch {
    // If rendering fails, use defaults
    fileData.pages = 1;
    fileData.orientation = 'portrait';
    fileData.preview = 'word';
  }
}

export async function processImageFile(fileData: ProcessedFile): Promise<void> {
  return new Promise<void>((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        fileData.width = img.width;
        fileData.height = img.height;
        fileData.orientation = img.width > img.height ? 'landscape' : 'portrait';
        fileData.preview = e.target?.result as string;
        fileData.pages = 1;
        fileData.category = 'Image';
        resolve();
      };
      img.onerror = () => {
        fileData.preview = 'image';
        fileData.orientation = 'portrait';
        fileData.pages = 1;
        fileData.category = 'Image';
        resolve();
      };
      img.src = e.target?.result as string;
    };
    reader.readAsDataURL(fileData.file);
  });
}
