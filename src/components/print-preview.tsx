'use client';

import { useCallback, useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { renderWordToHtml, type WordRenderResult } from '@/lib/word-render';

interface PrintPreviewProps {
  files: Array<{
    file: File;
    name: string;
    category: string;
    arrayBuffer?: ArrayBuffer;
  }>;
  onClose: () => void;
  onRetry?: () => void;
}

const RENDER_CLASS = 'wp-render';

/**
 * Print Preview — Renders Word docs via docx-preview and opens
 * a standalone HTML page in a new tab for printing or "Save as PDF".
 *
 * Approach:
 *   1. Render each Word file with docx-preview in hidden container
 *   2. Convert ALL images to base64 (self-contained, no blob URLs)
 *   3. Build a COMPLETE standalone HTML page with @page CSS
 *   4. Open as a Blob URL in a new browser tab
 *   5. The new tab auto-triggers window.print()
 */
export function PrintPreview({ files, onClose, onRetry }: PrintPreviewProps) {
  const [mounted, setMounted] = useState(false);
  const [status, setStatus] = useState<
    'rendering' | 'opening' | 'done' | 'error'
  >('rendering');
  const [currentFile, setCurrentFile] = useState('');
  const [renderedCount, setRenderedCount] = useState(0);
  const [errorMsg, setErrorMsg] = useState('');

  useEffect(() => {
    setMounted(true);
  }, []);

  const doRender = useCallback(async (cancelled: { current: boolean }) => {
    const wordFiles = files.filter((f) => f.category === 'Word');
    if (wordFiles.length === 0) {
      setStatus('done');
      return;
    }

    setStatus('rendering');

    // Phase 1: Render all Word files
    const results: WordRenderResult[] = [];

    for (let i = 0; i < wordFiles.length; i++) {
      if (cancelled.current) return;
      const wf = wordFiles[i];
      setCurrentFile(wf.name);

      try {
        const result = await renderWordToHtml(wf.arrayBuffer!, RENDER_CLASS);
        results.push(result);
      } catch (err) {
        console.error(`Failed to render ${wf.name}:`, err);
        throw new Error(`渲染失败: ${wf.name}`);
      }

      setRenderedCount(i + 1);
    }

    if (cancelled.current) return;

    // Phase 2: Build standalone HTML for the new tab
    setStatus('opening');

    const totalPages = results.reduce((s, r) => s + r.pageCount, 0);
    const wMm = results[0].pageWidthMm;
    const hMm = results[0].pageHeightMm;

    const allStyles = results
      .map((r) => r.styles)
      .filter(Boolean)
      .join('\n');

    const allHtml = results
      .map(
        (r, i) =>
          (i > 0
            ? '<div class="wp-file-break"></div>'
            : '') + r.html
      )
      .join('\n');

    const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>文档预览</title>
<style>
  @page {
    size: ${wMm}mm ${hMm}mm;
    margin: 0;
  }
  *, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }
  html, body {
    background: white;
    color: black;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
    color-adjust: exact !important;
  }
  .wp-file-break {
    page-break-before: always;
    height: 0;
    overflow: hidden;
    line-height: 0;
    font-size: 0;
  }
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
  @media print {
    html, body { margin: 0 !important; padding: 0 !important; }
    .wp-no-print { display: none !important; }
  }
</style>
${allStyles ? `<style>${allStyles}</style>` : ''}
</head>
<body>

<div class="wp-no-print" style="
  position: fixed; top: 0; left: 0; right: 0; z-index: 9999;
  padding: 10px 20px;
  background: linear-gradient(135deg, #7c3aed, #6d28d9);
  color: white;
  display: flex; align-items: center; justify-content: space-between;
  font-family: system-ui, -apple-system, sans-serif;
">
  <div style="display:flex;align-items:center;gap:12px;">
    <strong style="font-size:16px;">文档预览</strong>
    <span style="opacity:0.8;font-size:13px;">共 ${totalPages} 页，${wordFiles.length} 个文件</span>
  </div>
  <button onclick="window.print()" style="
    padding: 8px 20px;
    background: #22c55e;
    color: white;
    border: none;
    border-radius: 8px;
    font-weight: 600;
    cursor: pointer;
    font-size: 14px;
  ">打印 / 另存为 PDF</button>
</div>
<div class="wp-no-print" style="height:48px;"></div>

${allHtml}

<script>
  window.addEventListener('load', function() {
    setTimeout(function() { window.print(); }, 600);
  });
</script>

</body>
</html>`;

    // Phase 3: Open in new tab
    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    const blobUrl = URL.createObjectURL(blob);

    const newTab = window.open(blobUrl, '_blank');

    if (!newTab) {
      throw new Error(
        '浏览器阻止了弹出窗口，请允许弹出窗口后重试'
      );
    }

    setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);

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

  // FIX: Retry actually re-runs the render instead of just closing
  const handleRetry = useCallback(() => {
    setStatus('rendering');
    setRenderedCount(0);
    setCurrentFile('');
    setErrorMsg('');
    doRender({ current: false });
  }, [doRender]);

  const wordCount = files.filter((f) => f.category === 'Word').length;

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
              正在渲染 Word 文档...
            </p>
            <p
              style={{
                fontSize: 13,
                color: '#6b7280',
                marginTop: 8,
              }}
            >
              ({renderedCount}/{wordCount}) {currentFile}
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
              请在新标签页中检查文档效果。
              <br />
              使用打印对话框中的「另存为 PDF」来保存。
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
