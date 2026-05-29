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

const CLASS_NAME = 'docx-print-render';

/**
 * Print Preview using an isolated iframe.
 *
 * Key fixes:
 *   1. Converts ALL blob images to base64 BEFORE transferring HTML to iframe
 *   2. Detects page dimensions from docx-preview sections
 *   3. Sets @page CSS to match exact page size with margin:0
 *   4. Adds page-break-after to each section so pages map 1:1
 *   5. Cleans wrapper padding/margin to prevent page overflow
 */
export function PrintPreview({ files, onClose }: PrintPreviewProps) {
  const iframeRef = useRef<HTMLIFrameElement>(null);
  const [status, setStatus] = useState<'loading' | 'rendering' | 'ready' | 'error'>('loading');
  const [currentFile, setCurrentFile] = useState('');
  const [renderedCount, setRenderedCount] = useState(0);
  const [mounted, setMounted] = useState(false);

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
        setTimeout(resolve, 1000);
      });

      if (cancelled) return;

      const doc = iframe.contentDocument;
      if (!doc) {
        setStatus('error');
        return;
      }

      setStatus('rendering');

      // Set up base iframe HTML with print-safe defaults
      doc.open();
      doc.write(`<!DOCTYPE html>
<html><head>
<meta charset="utf-8">
<style>
  @page {
    size: 210mm 297mm;
    margin: 0;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html, body {
    width: 100%;
    background: white;
    color: black;
  }
  .file-separator {
    page-break-before: always;
    height: 0;
    overflow: hidden;
  }
  /* Each docx-preview section = exactly one printed page */
  section.${CLASS_NAME} {
    page-break-after: always;
    page-break-inside: avoid;
    overflow: hidden !important;
    margin: 0 !important;
    float: none !important;
    position: relative !important;
  }
  section.${CLASS_NAME}:last-of-type {
    page-break-after: auto;
  }
  /* Remove wrapper padding/margin that causes extra pages */
  .${CLASS_NAME}-wrapper {
    background: white !important;
    padding: 0 !important;
    margin: 0 !important;
    box-shadow: none !important;
    display: block !important;
  }
  @media print {
    html, body {
      margin: 0 !important;
      padding: 0 !important;
    }
  }
</style>
</head><body></body></html>`);
      doc.close();

      if (wordFiles.length === 0) {
        setStatus('ready');
        return;
      }

      try {
        const { renderAsync } = await import('docx-preview');
        const body = doc.body;

        // Track detected page dimensions — will update @page after first render
        let detectedPageWidthMm = 210;
        let detectedPageHeightMm = 297;
        let pageCssUpdated = false;

        for (let i = 0; i < wordFiles.length; i++) {
          if (cancelled) return;

          const wf = wordFiles[i];
          setCurrentFile(wf.name);

          // Page break before each file (except first)
          if (i > 0) {
            const separator = doc.createElement('div');
            separator.className = 'file-separator';
            body.appendChild(separator);
          }

          // Render container for this Word file in the iframe
          const renderDiv = doc.createElement('div');
          body.appendChild(renderDiv);

          try {
            // Step 1: Render in main window's hidden container (docx-preview uses main window's document)
            const tempContainer = document.createElement('div');
            tempContainer.style.cssText =
              'position:fixed;top:0;left:0;width:100vw;z-index:-1;pointer-events:none;background:white;';
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

              // Step 2: Convert ALL blob images to base64 (critical for iframe transfer)
              await convertAllImagesToBase64(tempContainer);

              // Step 3: Wait for images to fully load
              await waitForImages(tempContainer);

              // Step 4: Detect page dimensions from first rendered section
              if (!pageCssUpdated) {
                const firstSection = tempContainer.querySelector(
                  `section.${CLASS_NAME}`
                ) as HTMLElement;
                if (firstSection) {
                  // docx-preview sets inline width/height on sections
                  const style = firstSection.style;
                  const wPx = parseFloat(style.width) || firstSection.offsetWidth;
                  const hPx = parseFloat(style.height) || firstSection.offsetHeight;
                  detectedPageWidthMm = (wPx * 25.4) / 96;
                  detectedPageHeightMm = (hPx * 25.4) / 96;

                  // Add a small buffer (0.5mm) to prevent sub-pixel overflow
                  detectedPageWidthMm = Math.ceil(detectedPageWidthMm * 2) / 2 + 0.5;
                  detectedPageHeightMm = Math.ceil(detectedPageHeightMm * 2) / 2 + 0.5;

                  // Update @page CSS in iframe
                  const pageStyle = doc.createElement('style');
                  pageStyle.id = 'docx-page-size';
                  pageStyle.textContent = `
                    @page {
                      size: ${detectedPageWidthMm}mm ${detectedPageHeightMm}mm;
                      margin: 0;
                    }
                  `;
                  doc.head.appendChild(pageStyle);
                  pageCssUpdated = true;
                }
              }

              // Step 5: Clean wrapper styles to prevent page overflow
              const wrapper = tempContainer.querySelector(
                `.${CLASS_NAME}-wrapper`
              ) as HTMLElement;
              if (wrapper) {
                wrapper.style.background = 'white';
                wrapper.style.padding = '0';
                wrapper.style.margin = '0';
                wrapper.style.boxShadow = 'none';
              }

              // Step 6: Transfer <style> elements to iframe
              const styles = tempContainer.querySelectorAll('style');
              styles.forEach((s) => {
                const cloned = doc.createElement('style');
                cloned.textContent = s.textContent;
                doc.head.appendChild(cloned);
              });

              // Step 7: Transfer rendered HTML
              renderDiv.innerHTML = tempContainer.innerHTML;

              // Step 8: Wait for images to load in iframe context
              await waitForImages(renderDiv, doc.defaultView || undefined);
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
    return () => {
      cancelled = true;
    };
  }, [mounted, files]);

  const handlePrint = useCallback(() => {
    const iframe = iframeRef.current;
    if (!iframe?.contentWindow) return;
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
  }, []);

  const wordCount = files.filter((f) => f.category === 'Word').length;
  const nonWordCount = files.length - wordCount;

  if (!mounted) return null;

  return createPortal(
    <div
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 99999,
        display: 'flex',
        flexDirection: 'column',
        background: '#f0f0f0',
      }}
    >
      {/* Header */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '12px 24px',
          background: 'linear-gradient(135deg, #7c3aed, #6d28d9)',
          color: 'white',
          flexShrink: 0,
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <h2 style={{ fontSize: 18, fontWeight: 'bold' }}>打印预览</h2>
          {status === 'loading' && (
            <span style={{ fontSize: 14, opacity: 0.8 }}>正在初始化...</span>
          )}
          {status === 'rendering' && (
            <span style={{ fontSize: 14, opacity: 0.8 }}>
              正在渲染 ({renderedCount}/{wordCount}){' '}
              {currentFile ? `— ${currentFile}` : ''}
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
            另存为 PDF / 打印
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
        <div
          style={{
            padding: '10px 24px',
            background: '#fef3c7',
            borderBottom: '1px solid #fde68a',
            fontSize: 13,
            color: '#92400e',
            textAlign: 'center',
          }}
        >
          提示：PDF 和图片文件已通过合并引擎处理。请先使用「合并并下载」获取完整
          PDF，或在打印对话框中处理。 Word 文档将以完美排版在此预览并打印。
        </div>
      )}

      {/* Loading indicator */}
      {(status === 'loading' || status === 'rendering') && (
        <div
          style={{
            flex: 1,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
          }}
        >
          <div style={{ textAlign: 'center', color: '#6b7280' }}>
            <div
              style={{
                width: 48,
                height: 48,
                margin: '0 auto 16px',
                border: '4px solid #e5e7eb',
                borderTopColor: '#7c3aed',
                borderRadius: '50%',
                animation: 'spin 1s linear infinite',
              }}
            />
            <p>正在渲染 Word 文档，请稍候...</p>
            <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
          </div>
        </div>
      )}

      {/* iframe for rendered content */}
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
        <div
          style={{
            padding: '10px 24px',
            background: '#f0fdf4',
            borderTop: '1px solid #bbf7d0',
            fontSize: 13,
            color: '#166534',
            textAlign: 'center',
          }}
        >
          ✓ 渲染完成！请检查上方预览效果，然后点击「另存为 PDF /
          打印」按钮。在打印对话框中可以选择「另存为
          PDF」来保存。
        </div>
      )}

      {/* Error state */}
      {status === 'error' && (
        <div
          style={{
            flex: 1,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
          }}
        >
          <div
            style={{ textAlign: 'center', color: '#dc2626', padding: 24 }}
          >
            <p style={{ fontSize: 16, fontWeight: 'bold' }}>渲染失败</p>
            <p
              style={{ fontSize: 14, color: '#6b7280', marginTop: 8 }}
            >
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

// ─── Helpers ───────────────────────────────────────────────────────────────

/**
 * Convert every blob: image inside `container` to a base64 data-URL so it
 * survives innerHTML transfer to the iframe.
 */
async function convertAllImagesToBase64(container: HTMLElement): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return;

  for (let i = 0; i < images.length; i++) {
    const img = images[i] as HTMLImageElement;
    if (!img.src || !img.src.startsWith('blob:')) continue;
    try {
      const resp = await fetch(img.src);
      const blob = await resp.blob();
      img.src = await blobToBase64(blob);
    } catch (e) {
      console.warn('Failed to convert blob image to base64:', e);
      img.removeAttribute('src');
    }
  }

  // Small delay to let the browser repaint with new base64 src
  await new Promise((r) => setTimeout(r, 200));
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

/**
 * Wait for all images in a container to finish loading (or error).
 */
function waitForImages(container: HTMLElement, win?: Window): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return Promise.resolve();

  const targetWindow = win || window;

  return new Promise((resolve) => {
    let pending = images.length;
    let done = false;
    const finish = () => {
      if (!done) {
        done = true;
        resolve();
      }
    };

    for (let i = 0; i < images.length; i++) {
      const img = images[i] as HTMLImageElement;
      if (img.complete && img.naturalWidth > 0) {
        pending--;
        if (pending === 0) finish();
      } else {
        const handler = () => {
          pending--;
          if (pending === 0) finish();
        };
        img.addEventListener('load', handler, { once: true });
        img.addEventListener('error', handler, { once: true });
      }
    }
    // Timeout fallback
    targetWindow.setTimeout(finish, 10000);
  });
}
