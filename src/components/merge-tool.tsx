'use client';

import { useCallback, useState } from 'react';
import { PDFDocument } from 'pdf-lib';
import { FileSymlink, Trash2, Download, Loader2, Lightbulb, Printer, Eye } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { FileUploadZone } from '@/components/file-upload-zone';
import { FileList, FileListHeader } from '@/components/file-list';
import { MergeOptions, type PageSizeOption, type OrientationOption, type QualityOption } from '@/components/merge-options';
import { MergeProgress } from '@/components/merge-progress';
import type { ProcessedFile } from '@/lib/process-files';
import {
  getFileCategory,
  processPdfFile,
  processWordFile,
  processImageFile,
} from '@/lib/process-files';
import { addPdfPages, addImagePage } from '@/lib/merge-pdf';
import { addWordPages } from '@/lib/word-to-pdf';
import { useToast } from '@/hooks/use-toast';

const VALID_TYPES = [
  'application/pdf',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'image/jpeg',
  'image/png',
  'image/gif',
  'image/bmp',
  'image/webp',
];

export type OutputMode = 'download' | 'preview';

export function MergeTool() {
  const { toast } = useToast();

  // File state
  const [files, setFiles] = useState<ProcessedFile[]>([]);
  const [isMerging, setIsMerging] = useState(false);

  // Progress state
  const [progressVisible, setProgressVisible] = useState(false);
  const [progressPercent, setProgressPercent] = useState(0);
  const [progressFileName, setProgressFileName] = useState('');

  // Options state
  const [pageSize, setPageSize] = useState<PageSizeOption>('a4');
  const [orientation, setOrientation] = useState<OrientationOption>('all-portrait');
  const [imageQuality, setImageQuality] = useState<QualityOption>('0.9');

  const hasWordFiles = files.some((f) => f.category === 'Word');

  // Handle file upload
  const handleFilesSelected = useCallback(async (fileList: FileList) => {
    const newFiles: ProcessedFile[] = [];

    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      const isDocx = file.name.toLowerCase().endsWith('.docx');

      if (!VALID_TYPES.includes(file.type) && !isDocx) {
        toast({
          title: '不支持的文件格式',
          description: `${file.name} 不是支持的格式`,
          variant: 'destructive',
        });
        continue;
      }

      const category = getFileCategory(file.name, file.type);
      if (!category) continue;

      const fileData: ProcessedFile = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
        file,
        name: file.name,
        type: isDocx ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : file.type,
        size: file.size,
        category,
        preview: null,
        orientation: 'portrait',
        pages: 0,
      };

      try {
        if (category === 'PDF') {
          await processPdfFile(fileData);
        } else if (category === 'Word') {
          await processWordFile(fileData);
        } else {
          await processImageFile(fileData);
        }
        newFiles.push(fileData);
      } catch (error) {
        console.error(`Error processing ${file.name}:`, error);
        toast({
          title: '文件处理失败',
          description: `${file.name} 处理时出错`,
          variant: 'destructive',
        });
      }
    }

    if (newFiles.length > 0) {
      setFiles((prev) => [...prev, ...newFiles]);
    }
  }, [toast]);

  // Reorder files
  const handleReorder = useCallback((fromIndex: number, toIndex: number) => {
    setFiles((prev) => {
      const next = [...prev];
      const [removed] = next.splice(fromIndex, 1);
      next.splice(toIndex, 0, removed);
      return next;
    });
  }, []);

  // Delete file
  const handleDelete = useCallback((id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  }, []);

  // Clear all
  const handleClearAll = useCallback(() => {
    setFiles([]);
  }, []);

  // Core merge logic — returns PDF bytes
  const doMerge = useCallback(async (): Promise<Uint8Array> => {
    const mergedPdf = await PDFDocument.create();

    // For Word files, we use docx-preview which determines its own page size from the document.
    // Only pass effectivePageSize for PDF/image files.
    const effectivePageSize = pageSize;

    // Estimate total work for progress
    let totalWork = 0;
    for (const fileData of files) {
      if (fileData.category === 'Word') {
        totalWork += 3; // Word is heavier per file
      } else {
        totalWork += fileData.pages || 1;
      }
    }
    let processedWork = 0;

    // Process files sequentially
    for (let i = 0; i < files.length; i++) {
      const fileData = files[i];
      setProgressFileName(fileData.name);

      if (fileData.category === 'PDF') {
        await addPdfPages(
          mergedPdf,
          fileData.arrayBuffer!,
          orientation
        );
        processedWork += fileData.pages || 1;
      } else if (fileData.category === 'Word') {
        const wordPages = await addWordPages(
          mergedPdf,
          fileData.arrayBuffer!,
          orientation
        );
        processedWork += 3;
        // Update page count for the file
        setFiles((prev) =>
          prev.map((f) => (f.id === fileData.id ? { ...f, pages: wordPages } : f))
        );
      } else if (fileData.category === 'Image') {
        await addImagePage(
          mergedPdf,
          fileData.file,
          fileData.width || 0,
          fileData.height || 0,
          fileData.type,
          effectivePageSize,
          orientation,
          parseFloat(imageQuality)
        );
        processedWork += 1;
      }

      setProgressPercent((processedWork / totalWork) * 90);
      // Give UI time to update
      await new Promise((resolve) => setTimeout(resolve, 10));
    }

    setProgressFileName('正在生成PDF文件...');
    setProgressPercent(95);

    const mergedPdfBytes = await mergedPdf.save();
    return mergedPdfBytes;
  }, [files, pageSize, orientation, imageQuality]);

  // Output the merged PDF as a Blob URL
  const outputPdf = useCallback((pdfBytes: Uint8Array, mode: OutputMode) => {
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);

    if (mode === 'download') {
      const a = document.createElement('a');
      a.href = url;
      a.download = `合并文档_${new Date().toISOString().slice(0, 10)}.pdf`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      // Delay revoke to ensure download starts
      setTimeout(() => URL.revokeObjectURL(url), 30000);
      toast({
        title: '合并完成',
        description: 'PDF文件已成功生成并开始下载',
      });
    } else {
      // Print preview — open in new tab
      window.open(url, '_blank');
      toast({
        title: '合并完成',
        description: 'PDF文件已在新标签页中打开，可直接打印预览',
      });
    }

    setProgressPercent(100);
    setProgressFileName('完成！');
    setTimeout(() => setProgressVisible(false), 1000);
  }, [toast]);

  // Merge and download
  const handleMergeAndDownload = useCallback(async () => {
    if (files.length === 0 || isMerging) return;
    setIsMerging(true);
    setProgressVisible(true);
    setProgressPercent(0);

    try {
      const pdfBytes = await doMerge();
      outputPdf(pdfBytes, 'download');
    } catch (error) {
      console.error('Merge error:', error);
      toast({
        title: '合并失败',
        description: error instanceof Error ? error.message : '合并过程中出现未知错误',
        variant: 'destructive',
      });
      setProgressVisible(false);
    } finally {
      setIsMerging(false);
    }
  }, [files, isMerging, doMerge, outputPdf, toast]);

  // Merge and preview
  const handleMergeAndPreview = useCallback(async () => {
    if (files.length === 0 || isMerging) return;
    setIsMerging(true);
    setProgressVisible(true);
    setProgressPercent(0);

    try {
      const pdfBytes = await doMerge();
      outputPdf(pdfBytes, 'preview');
    } catch (error) {
      console.error('Merge error:', error);
      toast({
        title: '合并失败',
        description: error instanceof Error ? error.message : '合并过程中出现未知错误',
        variant: 'destructive',
      });
      setProgressVisible(false);
    } finally {
      setIsMerging(false);
    }
  }, [files, isMerging, doMerge, outputPdf, toast]);

  return (
    <div className="w-full max-w-4xl mx-auto space-y-5">
      {/* Upload Zone */}
      <FileUploadZone onFilesSelected={handleFilesSelected} disabled={isMerging} />

      {/* Merge Options */}
      <MergeOptions
        pageSize={pageSize}
        orientation={orientation}
        imageQuality={imageQuality}
        hasWordFiles={hasWordFiles}
        onPageSizeChange={setPageSize}
        onOrientationChange={setOrientation}
        onImageQualityChange={setImageQuality}
      />

      {/* File List */}
      <FileListHeader fileCount={files.length} />
      <FileList
        files={files}
        onReorder={handleReorder}
        onDelete={handleDelete}
      />

      {/* Action Buttons */}
      <div className="flex flex-col sm:flex-row gap-3 pt-2">
        <Button
          variant="outline"
          onClick={handleClearAll}
          disabled={files.length === 0 || isMerging}
          className="flex items-center gap-2"
        >
          <Trash2 className="w-4 h-4" />
          清空列表
        </Button>

        {/* Print Preview Button */}
        <Button
          onClick={handleMergeAndPreview}
          disabled={files.length === 0 || isMerging}
          variant="outline"
          className="flex-1 flex items-center justify-center gap-2 border-purple-300 text-purple-700 hover:bg-purple-50 hover:text-purple-800 transition-all duration-300 hover:-translate-y-0.5 disabled:opacity-60 disabled:cursor-not-allowed disabled:hover:translate-y-0"
        >
          {isMerging ? (
            <>
              <Loader2 className="w-4 h-4 animate-spin" />
              正在合并...
            </>
          ) : (
            <>
              <Eye className="w-4 h-4" />
              打印预览
            </>
          )}
        </Button>

        {/* Download Button */}
        <Button
          onClick={handleMergeAndDownload}
          disabled={files.length === 0 || isMerging}
          className="flex-1 flex items-center justify-center gap-2 bg-gradient-to-r from-purple-600 to-purple-500 hover:from-purple-700 hover:to-purple-600 text-white shadow-lg hover:shadow-xl transition-all duration-300 hover:-translate-y-0.5 disabled:opacity-60 disabled:cursor-not-allowed disabled:hover:translate-y-0"
        >
          {isMerging ? (
            <>
              <Loader2 className="w-4 h-4 animate-spin" />
              正在合并...
            </>
          ) : (
            <>
              <Download className="w-4 h-4" />
              合并并下载
            </>
          )}
        </Button>
      </div>

      {/* Progress */}
      <MergeProgress
        visible={progressVisible}
        percent={progressPercent}
        fileName={progressFileName}
      />

      {/* Tips Section */}
      <Card className="bg-amber-50 border-amber-200/60">
        <CardContent className="p-4 md:p-5">
          <div className="flex items-start gap-2 mb-3">
            <Lightbulb className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
            <h4 className="font-semibold text-amber-700">使用提示</h4>
          </div>
          <ul className="space-y-2 text-sm text-amber-800/80 ml-7">
            <li className="flex items-start gap-2">
              <span className="text-amber-400 mt-1">•</span>
              <span><strong>拖拽排序：</strong>可以拖动文件列表中的项目来调整合并顺序</span>
            </li>
            <li className="flex items-start gap-2">
              <span className="text-amber-400 mt-1">•</span>
              <span><strong>Word文档：</strong>支持 .docx 格式，使用高保真渲染引擎，最大程度保留原始排版、样式、表格和图片</span>
            </li>
            <li className="flex items-start gap-2">
              <span className="text-amber-400 mt-1">•</span>
              <span><strong>打印预览：</strong>合并后在新标签页打开PDF，可预览确认后再打印或下载</span>
            </li>
            <li className="flex items-start gap-2">
              <span className="text-amber-400 mt-1">•</span>
              <span><strong>方向处理：</strong>选择"统一为竖向"可自动将横向页面旋转，方便打印</span>
            </li>
            <li className="flex items-start gap-2">
              <span className="text-amber-400 mt-1">•</span>
              <span><strong>自动适配：</strong>图片会自动适应目标纸张大小，保持原始比例</span>
            </li>
          </ul>
        </CardContent>
      </Card>
    </div>
  );
}
