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

export async function processWordFile(fileData: ProcessedFile): Promise<void> {
  const arrayBuffer = await fileData.file.arrayBuffer();
  fileData.arrayBuffer = arrayBuffer;
  fileData.preview = 'word';
  fileData.category = 'Word';

  // FIX: Detect orientation from docx-preview rendering instead of hardcoding
  try {
    const { renderAsync } = await import('docx-preview');
    const container = document.createElement('div');
    container.style.cssText =
      'position:fixed;left:-9999px;top:0;opacity:0;pointer-events:none;width:100vw;';
    document.body.appendChild(container);

    try {
      await renderAsync(arrayBuffer, container, undefined, {
        className: 'wp-detect',
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        breakPages: true,
        ignoreLastRenderedPageBreak: true,
        trimXmlDeclaration: true,
        useBase64URL: true,
        renderHeaders: false,
        renderFooters: false,
        renderFootnotes: false,
        renderEndnotes: false,
      });

      const sections = container.querySelectorAll('section.wp-detect');
      if (sections.length > 0) {
        const first = sections[0] as HTMLElement;
        const w = first.offsetWidth || 794;
        const h = first.offsetHeight || 1123;
        fileData.orientation = w > h ? 'landscape' : 'portrait';
        fileData.width = w;
        fileData.height = h;
      } else {
        fileData.orientation = 'portrait';
      }
    } finally {
      document.body.removeChild(container);
    }
  } catch {
    // Fallback to portrait if rendering fails
    fileData.orientation = 'portrait';
  }

  fileData.pages = -1; // Calculated during merge
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
