'use client';

import { useCallback, useEffect, useRef, useState } from 'react';

interface PrintPreviewProps {
  files: Array<{
    file: File;
    name: string;
    category: string;
    arrayBuffer?: ArrayBuffer;
  }>;
  onClose: () => void;
}

export function PrintPreview({ files, onClose }: PrintPreviewProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [status, setStatus] = useState<'rendering' | 'ready' | 'error'>('rendering');
  const [currentFile, setCurrentFile] = useState('');
  const [renderedCount, setRenderedCount] = useState(0);

  useEffect(() => {
    if (!containerRef.current) return;

    let cancelled = false;

    async function renderFiles() {
      const container = containerRef.current!;
      const wordFiles = files.filter((f) => f.category === 'Word');

      if (wordFiles.length === 0) {
        setStatus('ready');
        return;
      }

      try {
        const { renderAsync } = await import('docx-preview');
        const CLASS_NAME = 'docx-print-render';

        for (let i = 0; i < wordFiles.length; i++) {
          if (cancelled) return;

          const wf = wordFiles[i];
          setCurrentFile(wf.name);

          // Create a sub-container for each Word file
          const fileContainer = document.createElement('div');
          fileContainer.className = 'print-file-section';
          fileContainer.dataset.fileName = wf.name;

          // File label (only visible on screen, hidden in print)
          const label = document.createElement('div');
          label.className = 'print-file-label';
          label.textContent = `📄 ${wf.name}`;
          fileContainer.appendChild(label);

          // Separator (page break before the next file)
          const separator = document.createElement('div');
          separator.className = 'print-file-separator';
          fileContainer.appendChild(separator);

          // Render container for docx-preview
          const renderDiv = document.createElement('div');
          fileContainer.appendChild(renderDiv);
          container.appendChild(fileContainer);

          try {
            await renderAsync(wf.arrayBuffer!, renderDiv, null, {
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
            await waitForImages(renderDiv);
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
  }, [files]);

  const handlePrint = useCallback(() => {
    window.print();
  }, []);

  const handleSavePdf = useCallback(() => {
    window.print();
  }, []);

  const wordCount = files.filter((f) => f.category === 'Word').length;
  const nonWordCount = files.length - wordCount;

  return (
    <div className="print-preview-overlay">
      {/* Header bar - hidden during print */}
      <div className="print-preview-header">
        <div className="flex items-center gap-3">
          <h2 className="text-lg font-bold text-white">打印预览</h2>
          {status === 'rendering' && (
            <span className="text-sm text-white/80">
              正在渲染 ({renderedCount}/{wordCount}) {currentFile && `— ${currentFile}`}
            </span>
          )}
          {status === 'ready' && (
            <span className="text-sm text-green-300">
              ✓ 渲染完成，共 {wordCount} 个Word文档
              {nonWordCount > 0 && `（另有 ${nonWordCount} 个文件已通过PDF合并）`}
            </span>
          )}
          {status === 'error' && (
            <span className="text-sm text-red-300">渲染出错，请检查文件格式</span>
          )}
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={handleSavePdf}
            disabled={status !== 'ready'}
            className="px-4 py-2 bg-white text-purple-700 rounded-lg font-medium hover:bg-gray-100 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            💾 另存为 PDF
          </button>
          <button
            onClick={handlePrint}
            disabled={status !== 'ready'}
            className="px-4 py-2 bg-purple-500 text-white rounded-lg font-medium hover:bg-purple-600 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            🖨️ 打印
          </button>
          <button
            onClick={onClose}
            className="px-3 py-2 text-white/80 hover:text-white transition-colors"
          >
            ✕ 关闭
          </button>
        </div>
      </div>

      {/* Scrollable preview area */}
      <div className="print-preview-body">
        {files.filter((f) => f.category !== 'Word').length > 0 && (
          <div className="print-file-section">
            <div className="print-file-label">
              📎 非Word文件（PDF、图片）已通过合并引擎直接嵌入，请使用&quot;合并并下载&quot;获取完整PDF
            </div>
            <div className="print-info-box">
              <p>以下 Word 文档将以原始排版显示，请点击上方&quot;另存为 PDF&quot;或&quot;打印&quot;按钮。</p>
              <p>在打印对话框中选择<strong>&quot;另存为 PDF&quot;</strong>即可获得完美的 Word 文档 PDF。</p>
              {nonWordCount > 0 && (
                <p className="mt-2 text-sm opacity-80">
                  提示：PDF 和图片文件已通过合并引擎处理，请直接使用&quot;合并并下载&quot;按钮获取包含所有文件的完整 PDF。
                </p>
              )}
            </div>
          </div>
        )}

        <div ref={containerRef} className="print-render-container" />

        {wordCount === 0 && (
          <div className="print-info-box">
            <p>没有 Word 文件需要预览。</p>
            <p>所有文件（PDF 和图片）已通过合并引擎处理，请使用&quot;合并并下载&quot;按钮。</p>
          </div>
        )}
      </div>
    </div>
  );
}

function waitForImages(container: HTMLElement): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return Promise.resolve();

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
        img.addEventListener('load', () => { pending--; if (pending === 0) finish(); }, { once: true });
        img.addEventListener('error', () => { pending--; if (pending === 0) finish(); }, { once: true });
      }
    }
    setTimeout(finish, 10000);
  });
}
