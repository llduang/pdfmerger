const RENDER_CLASS = 'wp-render';

export interface WordRenderResult {
  styles: string;
  html: string;
  pageWidthMm: number;
  pageHeightMm: number;
  pageCount: number;
}

export async function renderWordToHtml(
  arrayBuffer: ArrayBuffer,
  className = RENDER_CLASS
): Promise<WordRenderResult> {
  const { renderAsync } = await import('docx-preview');

  const container = document.createElement('div');
  container.style.cssText =
    'position:fixed;left:-9999px;top:0;opacity:0;pointer-events:none;';
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

    await convertAllImagesToBase64(container);
    await waitForAllImages(container);
    await delay(500);

    const sections = container.querySelectorAll(`section.${className}`);
    if (sections.length === 0) {
      throw new Error('Word 文档渲染失败：未生成任何页面');
    }

    const first = sections[0] as HTMLElement;
    const wMm = Math.ceil(first.offsetWidth * 25.4 / 96);
    const hMm = Math.ceil(first.offsetHeight * 25.4 / 96);

    const wrapper = container.querySelector(
      `.${className}-wrapper`
    ) as HTMLElement;
    if (wrapper) {
      wrapper.style.cssText =
        'background:white;padding:0;margin:0;box-shadow:none;';
    }

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
    const finish = () => { if (--pending <= 0) resolve(); };
    for (let i = 0; i < images.length; i++) {
      const img = images[i] as HTMLImageElement;
      if (img.complete && img.naturalWidth > 0) { finish(); }
      else {
        img.addEventListener('load', finish, { once: true });
        img.addEventListener('error', finish, { once: true });
      }
    }
    setTimeout(resolve, 15000);
  });
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}