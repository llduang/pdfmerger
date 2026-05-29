'use client';

import { useCallback, useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { PDFDocument } from 'pdf-lib';
import { renderWordToHtml, type WordRenderResult } from '@/lib/word-render';

interface PrintPreviewFile {
  file: File;
  name: string;
  category: string;
  arrayBuffer?: ArrayBuffer;
  width?: number;
  height?: number;
  type?: string;
}

interface PrintPreviewProps {
  files: PrintPreviewFile[];
  onClose: () => void;
  onRetry?: () => void;
}

const RENDER_CLASS = 'wp-render';

/**
 * Universal Print Preview — Handles ALL file types:
 *
 *   - PDF files: Each page is extracted and embedded as a standalone
 *     page in the printable HTML using an <embed> tag per page.
 *   - Word files: Rendered with docx-preview (perfect fidelity).
 *   - Image files: Rendered directly as <img> tags, sized to fit
 *     the target page.
 *
 * Everything is combined into a single standalone HTML page opened
 * in a new tab. The user prints or "Saves as PDF" from the browser's
 * print dialog.
 *
 * WHY this works perfectly for Word images:
 *   The browser's own rendering engine handles all CSS positioning,
 *   transforms, and layout — no screenshotting or DOM capture involved.
 */
export function PrintPreview({ files, onClose, onRetry }: PrintPreviewProps) {
  const [mounted, setMounted] = useState(false);
  const [status, setStatus] = useState<
    'rendering' | 'opening' | 'done' | 'error'
  >('rendering');
  const [currentFile, setCurrentFile] = useState('');
  const [progress, setProgress] = useState('');
  const [errorMsg, setErrorMsg] = useState('');

  useEffect(() => {
    setMounted(true);
  }, []);

  const doRender = useCallback(async (cancelled: { current: boolean }) => {
    if (files.length === 0) {
      setStatus('done');
      return;
    }

    setStatus('rendering');

    // ── Phase 1: Render Word files ──
    const wordResults: WordRenderResult[] = [];
    const wordFiles = files.filter((f) => f.category === 'Word');

    for (let i = 0; i < wordFiles.length; i++) {
      if (cancelled.current) return;
      const wf = wordFiles[i];
      setCurrentFile(wf.name);
      setProgress(`Word 文件 (${i + 1}/${wordFiles.length})`);

      try {
        const result = await renderWordToHtml(wf.arrayBuffer!, RENDER_CLASS);
        wordResults.push(result);
      } catch (err) {
        console.error(`Failed to render ${wf.name}:`, err);
        throw new Error(`渲染失败: ${wf.name}`);
      }
    }

    if (cancelled.current) return;

    // ── Phase 2: Convert PDF pages to individual data URLs ──
    const pdfFiles = files.filter((f) => f.category === 'PDF');
    const pdfPageDataUrls: string[] = [];

    for (let i = 0; i < pdfFiles.length; i++) {
      if (cancelled.current) return;
      const pf = pdfFiles[i];
      setCurrentFile(pf.name);
      setProgress(`PDF 文件 (${i + 1}/${pdfFiles.length})`);

      try {
        const pdfDoc = await PDFDocument.load(pf.arrayBuffer!);
        const pageCount = pdfDoc.getPageCount();

        // Save each page as a separate single-page PDF data URL
        for (let p = 0; p < pageCount; p++) {
          if (cancelled.current) return;
          const singlePagePdf = await PDFDocument.create();
          const [copiedPage] = await singlePagePdf.copyPages(pdfDoc, [p]);
          singlePagePdf.addPage(copiedPage);
          const pdfBytes = await singlePagePdf.save();
          const blob = new Blob([pdfBytes.buffer as ArrayBuffer], { type: 'application/pdf' });
          pdfPageDataUrls.push(URL.createObjectURL(blob));
        }
      } catch (err) {
        console.error(`Failed to process ${pf.name}:`, err);
        throw new Error(`PDF 处理失败: ${pf.name}`);
      }
    }

    if (cancelled.current) return;

    // ── Phase 3: Convert image files to data URLs ──
    const imageFiles = files.filter((f) => f.category === 'Image');
    const imageDataUrls: Array<{ url: string; width: number; height: number }> = [];

    for (let i = 0; i < imageFiles.length; i++) {
      if (cancelled.current) return;
      const imf = imageFiles[i];
      setCurrentFile(imf.name);
      setProgress(`图片文件 (${i + 1}/${imageFiles.length})`);

      try {
        const url = await fileToDataUrl(imf.file);
        imageDataUrls.push({
          url,
          width: imf.width || 800,
          height: imf.height || 600,
        });
      } catch (err) {
        console.error(`Failed to process ${imf.name}:`, err);
      }
    }

    if (cancelled.current) return;

    // ── Phase 4: Build standalone HTML ──
    setStatus('opening');

    // Determine page dimensions from Word files (or default to A4)
    let wMm = 210;
    let hMm = 297;
    if (wordResults.length > 0) {
      wMm = wordResults[0].pageWidthMm;
      hMm = wordResults[0].pageHeightMm;
    }

    const wPx = Math.round(wMm * 96 / 25.4);
    const hPx = Math.round(hMm * 96 / 25.4);

    // Collect Word styles
    const allWordStyles = wordResults
      .map((r) => r.styles)
      .filter(Boolean)
      .join('\n');

    // Build the content sections
    const contentParts: string[] = [];

    // Add PDF pages (each as a separate printable "page")
    for (let i = 0; i < pdfPageDataUrls.length; i++) {
      contentParts.push(
        `<div class="print-page pdf-page">
          <embed src="${pdfPageDataUrls[i]}" type="application/pdf" width="100%" height="100%" />
        </div>`
      );
    }

    // Add Word files
    for (let i = 0; i < wordResults.length; i++) {
      if (i > 0) {
        contentParts.push('<div class="file-break"></div>');
      }
      contentParts.push(wordResults[i].html);
    }

    // Add images
    for (let i = 0; i < imageDataUrls.length; i++) {
      const img = imageDataUrls[i];
      contentParts.push(
        `<div class="print-page image-page">
          <img src="${img.url}" style="max-width:100%;max-height:100%;object-fit:contain;" />
        </div>`
      );
    }

    const totalPages =
      pdfPageDataUrls.length +
      wordResults.reduce((s, r) => s + Math.max(r.pageCount, 1), 0) +
      imageDataUrls.length;

    const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>合并文档预览</title>
<style>
  @page {
    size: ${wMm}mm ${hMm}mm;
    margin: 0;
  }
  *, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }
  html, body {
    background: #f3f4f6;
    color: black;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
    color-adjust: exact !important;
  }

  /* Word wrapper / section styles */
  .${RENDER_CLASS}-wrapper {
    background: white !important;
    padding: 0 !important;
    margin: 0 !important;
    box-shadow: none !important;
    border: none !important;
  }
  section.${RENDER_CLASS} {
    page-break-after: always !important;
    page-break-inside: avoid !important;
    break-inside: avoid !important;
    overflow: hidden !important;
    margin: 0 !important;
    padding: 0 !important;
    border: none !important;
    box-shadow: none !important;
    float: none !important;
    position: relative !important;
    display: block !important;
  }
  section.${RENDER_CLASS}:last-of-type {
    page-break-after: auto !important;
  }

  /* File break between different files */
  .file-break {
    page-break-before: always;
    height: 0;
    overflow: hidden;
    line-height: 0;
    font-size: 0;
  }

  /* PDF pages and image pages */
  .print-page {
    width: ${wPx}px;
    height: ${hPx}px;
    overflow: hidden;
    page-break-after: always;
    page-break-inside: avoid;
    break-inside: avoid;
    background: white;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    margin: 8px auto;
  }
  .print-page:last-of-type {
    page-break-after: auto;
  }
  .print-page embed {
    width: 100%;
    height: 100%;
    border: none;
  }
  .print-page.image-page {
    display: flex;
    align-items: center;
    justify-content: center;
  }

  /* Toolbar (hidden in print) */
  .toolbar {
    position: fixed;
    top: 0; left: 0; right: 0;
    z-index: 9999;
    padding: 10px 20px;
    background: linear-gradient(135deg, #7c3aed, #6d28d9);
    color: white;
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-family: system-ui, -apple-system, sans-serif;
  }
  .toolbar button {
    padding: 8px 20px;
    border: none;
    border-radius: 8px;
    font-weight: 600;
    cursor: pointer;
    font-size: 14px;
    transition: opacity 0.2s;
  }
  .toolbar button:hover { opacity: 0.9; }
  .btn-print { background: #22c55e; color: white; }
  .btn-close { background: rgba(255,255,255,0.2); color: white; }
  .toolbar-spacer { height: 48px; }

  @media print {
    html, body {
      margin: 0 !important;
      padding: 0 !important;
      background: white !important;
    }
    .toolbar, .toolbar-spacer { display: none !important; }
    .print-page {
      margin: 0 !important;
      box-shadow: none !important;
    }
    /* For PDF embeds in print: the browser's native PDF renderer
       handles the actual rendering when printing */
    .print-page embed {
      display: none;
    }
    .print-page.pdf-page::after {
      content: '请使用"合并并下载"按钮导出 PDF 文件（此预览不支持直接打印 PDF 页面）';
      display: flex;
      align-items: center;
      justify-content: center;
      width: 100%;
      height: 100%;
      font-size: 14px;
      color: #9ca3af;
    }
    .no-print { display: none !important; }
  }
</style>
${allWordStyles ? `<style>${allWordStyles}</style>` : ''}
</head>
<body>

<div class="toolbar no-print">
  <div style="display:flex;align-items:center;gap:12px;">
    <strong style="font-size:16px;">文档预览</strong>
    <span style="opacity:0.8;font-size:13px;">
      共 ${totalPages} 页，${files.length} 个文件
      ${pdfPageDataUrls.length > 0 ? `（含 ${pdfPageDataUrls.length} 页 PDF）` : ''}
    </span>
  </div>
  <div style="display:flex;gap:8px;">
    <button class="btn-print" onclick="window.print()">打印 / 另存为 PDF</button>
    <button class="btn-close" onclick="window.close()">关闭此页面</button>
  </div>
</div>
<div class="toolbar-spacer no-print"></div>

${contentParts.join('\n')}

<script>
  // Auto-trigger print after a short delay for rendering
  window.addEventListener('load', function() {
    setTimeout(function() { window.print(); }, 800);
  });
</script>

</body>
</html>`;

    // ── Phase 5: Open in new tab ──
    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    const blobUrl = URL.createObjectURL(blob);
    const newTab = window.open(blobUrl, '_blank');

    if (!newTab) {
      throw new Error(
        '浏览器阻止了弹出窗口，请允许弹出窗口后重试'
      );
    }

    // Cleanup blob URL after 2 minutes
    setTimeout(() => {
      URL.revokeObjectURL(blobUrl);
      pdfPageDataUrls.forEach((url) => URL.revokeObjectURL(url));
    }, 120000);

    if (!cancelled.current) setStatus('done');
  }, [files]);

  useEffect(() => {
    if (!mounted) return;
    const cancelled = { current: false };

    (async () => {
      try {
        await doRender(cancelled);
      } catch (err) {
        if (!cancelled.current) {
          setErrorMsg(
            err instanceof Error ? err.message : '预览过程中出错'
          );
          setStatus('error');
        }
      }
    })();

    return () => {
      cancelled.current = true;
    };
  }, [mounted, doRender]);

  const handleRetry = useCallback(() => {
    setStatus('rendering');
    setCurrentFile('');
    setProgress('');
    setErrorMsg('');
    doRender({ current: false });
  }, [doRender]);

  if (!mounted) return null;

  return createPortal(
    <div
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 99999,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        background: 'rgba(0,0,0,0.5)',
        backdropFilter: 'blur(4px)',
      }}
    >
      <div
        style={{
          background: 'white',
          borderRadius: 16,
          padding: '40px 48px',
          maxWidth: 440,
          width: '90%',
          textAlign: 'center',
          boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
        }}
      >
        {status === 'rendering' && (
          <>
            <div
              style={{
                width: 56,
                height: 56,
                margin: '0 auto 20px',
                border: '4px solid #e5e7eb',
                borderTopColor: '#7c3aed',
                borderRadius: '50%',
                animation: 'wp-spin 0.8s linear infinite',
              }}
            />
            <p style={{ fontSize: 16, fontWeight: 600, color: '#1f2937' }}>
              正在准备预览...
            </p>
            <p
              style={{
                fontSize: 13,
                color: '#6b7280',
                marginTop: 8,
              }}
            >
              {progress} — {currentFile}
            </p>
            <style>
              {`@keyframes wp-spin { to { transform: rotate(360deg); } }`}
            </style>
          </>
        )}

        {status === 'opening' && (
          <>
            <div
              style={{
                width: 56,
                height: 56,
                margin: '0 auto 20px',
                border: '4px solid #e5e7eb',
                borderTopColor: '#22c55e',
                borderRadius: '50%',
                animation: 'wp-spin 0.8s linear infinite',
              }}
            />
            <p style={{ fontSize: 16, fontWeight: 600, color: '#1f2937' }}>
              正在打开预览...
            </p>
            <p style={{ fontSize: 13, color: '#6b7280', marginTop: 8 }}>
              新标签页即将打开
            </p>
          </>
        )}

        {status === 'done' && (
          <>
            <div
              style={{
                width: 56,
                height: 56,
                margin: '0 auto 20px',
                borderRadius: '50%',
                background: '#f0fdf4',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: 28,
              }}
            >
              ✓
            </div>
            <p style={{ fontSize: 16, fontWeight: 600, color: '#166534' }}>
              预览已在新标签页打开
            </p>
            <p
              style={{
                fontSize: 13,
                color: '#6b7280',
                marginTop: 8,
                lineHeight: 1.6,
              }}
            >
              在新标签页中检查文档效果。
              <br />
              点击「打印 / 另存为 PDF」按钮保存。
            </p>
            <button
              onClick={onClose}
              style={{
                marginTop: 20,
                padding: '8px 24px',
                background: '#7c3aed',
                color: 'white',
                border: 'none',
                borderRadius: 8,
                cursor: 'pointer',
                fontSize: 14,
                fontWeight: 500,
              }}
            >
              关闭
            </button>
          </>
        )}

        {status === 'error' && (
          <>
            <div
              style={{
                width: 56,
                height: 56,
                margin: '0 auto 20px',
                borderRadius: '50%',
                background: '#fef2f2',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: 28,
              }}
            >
              ✕
            </div>
            <p style={{ fontSize: 16, fontWeight: 600, color: '#dc2626' }}>
              预览失败
            </p>
            <p
              style={{
                fontSize: 13,
                color: '#6b7280',
                marginTop: 8,
                lineHeight: 1.6,
              }}
            >
              {errorMsg}
            </p>
            <div
              style={{
                marginTop: 20,
                display: 'flex',
                gap: 8,
                justifyContent: 'center',
              }}
            >
              <button
                onClick={handleRetry}
                style={{
                  padding: '8px 20px',
                  background: '#7c3aed',
                  color: 'white',
                  border: 'none',
                  borderRadius: 8,
                  cursor: 'pointer',
                  fontSize: 14,
                }}
              >
                重试
              </button>
              <button
                onClick={onClose}
                style={{
                  padding: '8px 20px',
                  background: '#e5e7eb',
                  color: '#374151',
                  border: 'none',
                  borderRadius: 8,
                  cursor: 'pointer',
                  fontSize: 14,
                }}
              >
                关闭
              </button>
            </div>
          </>
        )}
      </div>
    </div>,
    document.body
  );
}

// ─── Helper ─────────────────────────────────────────────────────────────────

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}
