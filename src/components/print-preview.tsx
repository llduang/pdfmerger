'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

interface PrintPreviewProps {
  files: Array<{
    file: File;
    name: string;
    category: string;
    arrayBuffer?: ArrayBuffer;
  }>;
  onClose: () => void;
}

/**
 * Print Preview using an isolated iframe.
 *
 * Why iframe:
 *   - Complete CSS isolation from the main app (no Tailwind, no layout interference)
 *   - docx-preview renders perfectly inside the iframe
 *   - iframe.contentWindow.print() generates pixel-perfect PDF
 *   - No complex @media print CSS needed — the iframe IS the only content
 */
export function PrintPreview({ files, onClose }: PrintPreviewProps) {
  const iframeRef = useRef<HTMLIFrameElement>(null);
  const [status, setStatus] = useState<'loading' | 'rendering' | 'ready' | 'error'>('loading');
  const [currentFile, setCurrentFile] = useState('');
  const [renderedCount, setRenderedCount] = useState(0);
  const [mounted, setMounted] = useState(false);

  // Wait for client-side mount (portal target)
  useEffect(() => { setMounted(true); }, []);

  useEffect(() => {
    if (!mounted || !iframeRef.current) return;

    let cancelled = false;

    async function renderFiles() {
      const iframe = iframeRef.current!;
      const wordFiles = files.filter((f) => f.category === 'Word');

      // Wait for iframe to load
      await new Promise<void>((resolve) => {
        iframe.onload = () => resolve();
        // Fallback if already loaded
        setTimeout(resolve, 1000);
      });

      if (cancelled) return;

      const doc = iframe.contentDocument!;
      if (!doc) {
        setStatus('error');
        return;
      }

      setStatus('rendering');

      // Clear iframe and set up base styles
      doc.open();
      doc.write(`<!DOCTYPE html>
<html><head>
<meta charset="utf-8">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'SimSun', 'Songti SC', 'Noto Serif SC', 'Microsoft YaHei', serif;
    background: white;
    color: black;
  }
  .file-separator { page-break-before: always; }
</style>
</head><body></body></html>`);
      doc.close();

      if (wordFiles.length === 0) {
        setStatus('ready');
        return;
      }

      try {
        // Import docx-preview dynamically
        const docxPreview = await import('docx-preview');
        const { renderAsync } = docxPreview;
        const CLASS_NAME = 'docx-print-render';

        // We need to render into the iframe's body using the iframe's document
        // docx-preview needs the window.document reference of the iframe
        const body = doc.body;

        for (let i = 0; i < wordFiles.length; i++) {
          if (cancelled) return;

          const wf = wordFiles[i];
          setCurrentFile(wf.name);

          // Create a page break before each file (except the first)
          if (i > 0) {
            const separator = doc.createElement('div');
            separator.className = 'file-separator';
            body.appendChild(separator);
          }

          // Create render container for this Word file
          const renderDiv = doc.createElement('div');
          body.appendChild(renderDiv);

          try {
            // Render the Word document into the iframe's body
            // docx-preview uses global `document` — we need to temporarily override it
            // Actually, docx-preview accepts bodyContainer which is an element in the iframe's document
            // The library creates elements using document.createElement which defaults to the main window's document
            // This is a known limitation — we need to use the main window but inject styles into iframe

            // Alternative approach: render in main window hidden div, then transfer HTML
            const tempContainer = document.createElement('div');
            tempContainer.style.cssText = 'position:fixed;top:0;left:0;width:100vw;z-index:-1;pointer-events:none;background:white;';
            document.body.appendChild(tempContainer);

            try {
              await renderAsync(wf.arrayBuffer!, tempContainer, null, {
                className: CLASS_NAME,
                inWrapper: true,
                ignoreWidth: false,
                ignoreHeight: false,
                ignoreFonts: false,
                breakPages: true,
                ignoreLastRenderedPageBreak: true,
                experimental: false,
                trimXmlDeclaration: true,
                useBase64URL: true,
                renderHeaders: true,
                renderFooters: true,
                renderFootnotes: true,
                renderEndnotes: true,
              });

              // Wait for images to load
              await waitForImages(tempContainer);

              // Transfer all content (including <style> elements) to the iframe
              // 1. Copy all <style> elements
              const styles = tempContainer.querySelectorAll('style');
              styles.forEach((style) => {
                const clonedStyle = doc.createElement('style');
                clonedStyle.textContent = style.textContent;
                doc.head.appendChild(clonedStyle);
              });

              // 2. Copy the rendered HTML
              renderDiv.innerHTML = tempContainer.innerHTML;

              // 3. Wait for images in the iframe to load
              await waitForImages(renderDiv, doc.defaultView!);
            } finally {
              document.body.removeChild(tempContainer);
            }
          } catch (err) {
            console.error(`Failed to render ${wf.name}:`, err);
            renderDiv.innerHTML = `<p style="color:red;padding:20px;">渲染失败: ${wf.name}</p>`;
          }

          setRenderedCount(i + 1);
        }

        if (!cancelled) {
          setStatus('ready');
        }
      } catch (err) {
        console.error('Render error:', err);
        if (!cancelled) setStatus('error');
      }
    }

    renderFiles();
    return () => { cancelled = true; };
  }, [mounted, files]);

  // Print the iframe content
  const handlePrint = useCallback(() => {
    const iframe = iframeRef.current;
    if (!iframe?.contentWindow) return;
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
  }, []);

  const wordCount = files.filter((f) => f.category === 'Word').length;
  const nonWordCount = files.length - wordCount;

  // Use portal to render at body level — critical for z-index and print isolation
  if (!mounted) return null;

  return createPortal(
    <div style={{
      position: 'fixed',
      inset: 0,
      zIndex: 99999,
      display: 'flex',
      flexDirection: 'column',
      background: '#f0f0f0',
    }}>
      {/* Header */}
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        padding: '12px 24px',
        background: 'linear-gradient(135deg, #7c3aed, #6d28d9)',
        color: 'white',
        flexShrink: 0,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <h2 style={{ fontSize: 18, fontWeight: 'bold' }}>打印预览</h2>
          {status === 'loading' && (
            <span style={{ fontSize: 14, opacity: 0.8 }}>正在初始化...</span>
          )}
          {status === 'rendering' && (
            <span style={{ fontSize: 14, opacity: 0.8 }}>
              正在渲染 ({renderedCount}/{wordCount}) {currentFile ? `— ${currentFile}` : ''}
            </span>
          )}
          {status === 'ready' && (
            <span style={{ fontSize: 14, color: '#86efac' }}>
              ✓ 渲染完成，共 {wordCount} 个Word文档
            </span>
          )}
          {status === 'error' && (
            <span style={{ fontSize: 14, color: '#fca5a5' }}>渲染出错</span>
          )}
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <button
            onClick={handlePrint}
            disabled={status !== 'ready'}
            style={{
              padding: '8px 16px',
              background: '#22c55e',
              color: 'white',
              border: 'none',
              borderRadius: 8,
              fontWeight: 600,
              cursor: status !== 'ready' ? 'not-allowed' : 'pointer',
              opacity: status !== 'ready' ? 0.5 : 1,
              fontSize: 14,
            }}
          >
            💾 另存为 PDF / 打印
          </button>
          <button
            onClick={onClose}
            style={{
              padding: '8px 12px',
              background: 'transparent',
              color: 'rgba(255,255,255,0.8)',
              border: 'none',
              fontSize: 18,
              cursor: 'pointer',
            }}
          >
            ✕
          </button>
        </div>
      </div>

      {/* Info bar for non-Word files */}
      {nonWordCount > 0 && (
        <div style={{
          padding: '10px 24px',
          background: '#fef3c7',
          borderBottom: '1px solid #fde68a',
          fontSize: 13,
          color: '#92400e',
          textAlign: 'center',
        }}>
          提示：PDF 和图片文件已通过合并引擎处理。请先使用"合并并下载"获取完整 PDF，或在打印对话框中处理。
          Word 文档将以完美排版在此预览并打印。
        </div>
      )}

      {/* Loading indicator */}
      {(status === 'loading' || status === 'rendering') && (
        <div style={{
          flex: 1,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
        }}>
          <div style={{ textAlign: 'center', color: '#6b7280' }}>
            <div style={{
              width: 48, height: 48, margin: '0 auto 16px',
              border: '4px solid #e5e7eb',
              borderTopColor: '#7c3aed',
              borderRadius: '50%',
              animation: 'spin 1s linear infinite',
            }} />
            <p>正在渲染 Word 文档，请稍候...</p>
            <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
          </div>
        </div>
      )}

      {/* The iframe that contains the rendered Word documents */}
      <iframe
        ref={iframeRef}
        title="Word 文档预览"
        style={{
          flex: 1,
          border: 'none',
          display: status === 'ready' ? 'block' : 'none',
          background: 'white',
        }}
      />

      {/* Ready state message */}
      {status === 'ready' && wordCount > 0 && (
        <div style={{
          padding: '10px 24px',
          background: '#f0fdf4',
          borderTop: '1px solid #bbf7d0',
          fontSize: 13,
          color: '#166534',
          textAlign: 'center',
        }}>
          ✓ 渲染完成！请检查上方预览效果，然后点击"另存为 PDF / 打印"按钮。
          在打印对话框中可以选择"另存为 PDF"来保存。
        </div>
      )}

      {/* Error state */}
      {status === 'error' && (
        <div style={{
          flex: 1,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
        }}>
          <div style={{ textAlign: 'center', color: '#dc2626', padding: 24 }}>
            <p style={{ fontSize: 16, fontWeight: 'bold' }}>渲染失败</p>
            <p style={{ fontSize: 14, color: '#6b7280', marginTop: 8 }}>
              请确认文件为有效的 .docx 格式，然后重试。
            </p>
            <button
              onClick={onClose}
              style={{
                marginTop: 16,
                padding: '8px 24px',
                background: '#7c3aed',
                color: 'white',
                border: 'none',
                borderRadius: 8,
                cursor: 'pointer',
                fontSize: 14,
              }}
            >
              返回
            </button>
          </div>
        </div>
      )}
    </div>,
    document.body
  );
}

/**
 * Wait for all images in a container to finish loading.
 */
function waitForImages(
  container: HTMLElement,
  win?: Window
): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return Promise.resolve();

  const targetWindow = win || window;

  return new Promise((resolve) => {
    let pending = images.length;
    let done = false;
    const finish = () => { if (!done) { done = true; resolve(); } };

    for (let i = 0; i < images.length; i++) {
      const img = images[i] as HTMLImageElement;
      if (img.complete) {
        pending--;
        if (pending === 0) finish();
      } else {
        const loadHandler = () => { pending--; if (pending === 0) finish(); };
        img.addEventListener('load', loadHandler, { once: true });
        img.addEventListener('error', loadHandler, { once: true });
      }
    }
    targetWindow.setTimeout(finish, 10000);
  });
}
