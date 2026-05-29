/**
 * Cloudflare Worker — PDF 转换代理
 *
 * 接收前端发来的 .docx 文件，转发给 Gotenberg（LibreOffice），
 * 返回转换后的 PDF。
 *
 * 已部署地址：https://pdf-convert.2710476780.workers.dev
 * 后端 Gotenberg：https://gotenberg-production-74d3.up.railway.app
 */

const GOTENBERG_URL = 'https://gotenberg-production-74d3.up.railway.app';

export default {
  async fetch(request) {
    // CORS 预检
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        headers: corsHeaders(),
      });
    }

    if (request.method !== 'POST') {
      return jsonResponse('只允许 POST 请求', 405);
    }

    try {
      const formData = await request.formData();
      const file = formData.get('file');

      if (!file) {
        return jsonResponse('未收到文件', 400);
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
        return jsonResponse(
          `转换失败（${gotenbergResponse.status}）`,
          502
        );
      }

      // 返回 PDF
      const pdfBuffer = await gotenbergResponse.arrayBuffer();

      return new Response(pdfBuffer, {
        status: 200,
        headers: {
          'Content-Type': 'application/pdf',
          ...corsHeaders(),
        },
      });
    } catch (err) {
      console.error('Worker error:', err);
      return jsonResponse(
        `服务器内部错误: ${err.message || '未知错误'}`,
        500
      );
    }
  },
};

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}

function jsonResponse(message, status) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: {
      'Content-Type': 'application/json',
      ...corsHeaders(),
    },
  });
}
