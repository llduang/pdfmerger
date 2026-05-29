/**
 * word-render.ts — Shared Word rendering utilities.
 *
 * Renders a .docx file using docx-preview into a hidden container,
 * converts all blob images to base64 data-URLs, and returns
 * the extracted HTML + styles + detected page dimensions.
 */

const RENDER_CLASS = 'wp-render';

export interface WordRenderResult {
  /** All <style> text content concatenated */
  styles: string;
  /** innerHTML of the render container (includes wrapper + sections) */
  html: string;
  /** Detected page width in mm (rounded UP) */
  pageWidthMm: number;
  /** Detected page height in mm (rounded UP) */
  pageHeightMm: number;
  /** Number of pages (sections) detected */
  pageCount: number;
}

// ─── Public API ─────────────────────────────────────────────────────────────

/**
 * Render a .docx ArrayBuffer into clean HTML + styles.
 *
 * The rendering happens in an off-screen container in the MAIN window
 * (docx-preview requires the main window's `document`).  All blob images
 * are converted to inline base64 data-URLs before the container is
 * destroyed, so the returned HTML is completely self-contained.
 */
export async function renderWordToHtml(
  arrayBuffer: ArrayBuffer,
  className = RENDER_CLASS
): Promise<WordRenderResult> {
  const { renderAsync } = await import('docx-preview');

  // ── Step 1: Render with docx-preview in a hidden container ──
  const container = document.createElement('div');
  container.style.cssText =
    'position:fixed;left:-9999px;top:0;opacity:0;pointer-events:none;width:100vw;';
  document.body.appendChild(container);

  try {
    await renderAsync(arrayBuffer, container, undefined, {
      className,
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

    // ── Step 2: Convert every blob image to a base64 data-URL ──
    await convertAllImagesToBase64(container);

    // ── Step 3: Wait for images to fully load ──
    await waitForAllImages(container);
    // Extra settle time for layout reflow after images load
    await delay(300);

    // ── Step 4: Detect page dimensions from first section ──
    const sections = container.querySelectorAll(`section.${className}`);
    if (sections.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    const first = sections[0] as HTMLElement;
    // Use Math.ceil to round UP → @page is slightly larger than the section,
    // preventing content overflow to a spurious extra page.
    const wMm = Math.ceil(first.offsetWidth * 25.4 / 96);
    const hMm = Math.ceil(first.offsetHeight * 25.4 / 96);

    // ── Step 5: Clean wrapper styles that could cause extra spacing ──
    const wrapper = container.querySelector(
      `.${className}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.cssText =
        'background:white;padding:0;margin:0;box-shadow:none;';
    }

    // ── Step 6: Extract styles and HTML ──
    const styles = Array.from(container.querySelectorAll('style'))
      .map((s) => s.textContent || '')
      .join('\n');

    return {
      styles,
      html: container.innerHTML,
      pageWidthMm: wMm,
      pageHeightMm: hMm,
      pageCount: sections.length,
    };
  } finally {
    document.body.removeChild(container);
  }
}

// ─── Helpers ───────────────────────────────────────────────────────────────

async function convertAllImagesToBase64(
  container: HTMLElement
): Promise<void> {
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
      console.warn('[word-render] Failed to convert blob image:', e);
      img.removeAttribute('src');
    }
  }
  // Let the browser repaint
  await delay(200);
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function waitForAllImages(container: HTMLElement): Promise<void> {
  const images = container.querySelectorAll('img');
  if (images.length === 0) return Promise.resolve();

  return new Promise((resolve) => {
    let pending = images.length;
    const finish = () => {
      if (--pending <= 0) resolve();
    };
    for (let i = 0; i < images.length; i++) {
      const img = images[i] as HTMLImageElement;
      if (img.complete && img.naturalWidth > 0) {
        finish();
      } else {
        img.addEventListener('load', finish, { once: true });
        img.addEventListener('error', finish, { once: true });
      }
    }
    // Safety timeout
    setTimeout(resolve, 12000);
  });
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}
