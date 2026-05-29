/**
 * Cloudflare Pages Function — Word 转 PDF 代理
 *
 * 部署后路径: https://<your-pages-domain>/api/convert
 * 与静态站点同域，彻底消除跨域和 workers.dev 域名访问问题。
 *
 * 接收前端发来的 .docx 文件，转发给 Gotenberg（LibreOffice），
 * 返回转换后的 PDF。
 */

export const onRequestPost: PagesFunction<Env> = async (context) => {
  const GOTENBERG_URL = context.env.GOTENBERG_URL || 'https://gotenberg-production-74d3.up.railway.app';

  try {
    const formData = await context.request.formData();
    const file = formData.get('file');

    if (!file) {
      return new Response(JSON.stringify({ error: '未收到文件' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
      });
    }

    // 转发给 Gotenberg
    const gotenbergForm = new FormData();
    gotenbergForm.append('files', file, file.name || 'document.docx');

    const gotenbergResponse = await fetch(
      `${GOTENBERG_URL}/forms/libreoffice/convert`,
      {
        method: 'POST',
        body: gotenbergForm,
      }
    );

    if (!gotenbergResponse.ok) {
      const errText = await gotenbergResponse.text().catch(() => '');
      console.error('Gotenberg error:', gotenbergResponse.status, errText);

      // 4xx 客户端错误（文件损坏等）直接返回
      if (gotenbergResponse.status >= 400 && gotenbergResponse.status < 500) {
        return new Response(
          JSON.stringify({
            error: `文件转换失败，请检查文件是否损坏（${gotenbergResponse.status}）`,
          }),
          {
            status: 502,
            headers: { 'Content-Type': 'application/json' },
          }
        );
      }

      // 5xx 服务端错误（服务正在启动等）返回特定状态码，让前端重试
      return new Response(
        JSON.stringify({
          error: `转换服务暂时不可用（${gotenbergResponse.status}）`,
          retryable: true,
        }),
        {
          status: 503,
          headers: { 'Content-Type': 'application/json' },
        }
      );
    }

    // 返回 PDF
    const pdfBuffer = await gotenbergResponse.arrayBuffer();

    return new Response(pdfBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Cache-Control': 'no-cache',
      },
    });
  } catch (err: any) {
    console.error('Pages Function error:', err);
    return new Response(
      JSON.stringify({
        error: `服务器内部错误: ${err.message || '未知错误'}`,
        retryable: true,
      }),
      {
        status: 500,
        headers: { 'Content-Type': 'application/json' },
      }
    );
  }
};
