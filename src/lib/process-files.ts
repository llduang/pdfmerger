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
  preview: string | null; // data URL for images, emoji for others
  orientation: Orientation;
  pages: number;
  width?: number;
  height?: number;
  // Raw data for merge
  arrayBuffer?: ArrayBuffer;
  htmlContent?: string; // For Word files
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
  fileData.orientation = width > height ? 'landscape' : 'portrait';
  fileData.width = width;
  fileData.height = height;
  fileData.arrayBuffer = arrayBuffer;
  fileData.preview = 'pdf'; // Special marker for PDF icon
}

export async function processWordFile(fileData: ProcessedFile): Promise<void> {
  const arrayBuffer = await fileData.file.arrayBuffer();
  fileData.arrayBuffer = arrayBuffer;

  // Dynamic import to avoid SSR issues
  const mammoth = await import('mammoth');

  const result = await mammoth.convertToHtml({ arrayBuffer });
  fileData.htmlContent = result.value;
  fileData.preview = 'word'; // Special marker for Word icon
  fileData.category = 'Word';
  fileData.orientation = 'portrait';
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
