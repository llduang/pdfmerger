/**
 * Cloudflare Pages Function — Word 转 PDF 代理（Microsoft Graph API 版）
 *
 * 流程：
 *   1. 客户端凭据流获取 Microsoft Graph access_token
 *   2. 将 .docx 上传到 OneDrive 临时目录
 *   3. 调用 Graph API 以 PDF 格式下载（微软服务端引擎转换）
 *   4. 删除 OneDrive 上的临时文件
 *   5. 返回 PDF 给前端
 *
 * 需要的环境变量（在 Cloudflare Dashboard 中配置）：
 *   MS_CLIENT_ID     — Azure AD 应用的客户端 ID
 *   MS_CLIENT_SECRET — Azure AD 应用的客户端密钥
 *   MS_TENANT_ID     — Azure AD 租户 ID
 *   MS_USER_ID       — OneDrive 所属用户的 User Principal Name 或 Object ID
 */

interface Env {
  MS_CLIENT_ID: string;
  MS_CLIENT_SECRET: string;
  MS_TENANT_ID: string;
  MS_USER_ID: string;
}

interface TokenCache {
  token: string;
  expiresAt: number;
}

let tokenCache: TokenCache | null = null;

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const TEMP_FOLDER = 'pdfmerger-temp';

async function getAccessToken(env: Env): Promise<string> {
  if (tokenCache && tokenCache.expiresAt > Date.now() + 60_000) {
    return tokenCache.token;
  }

  const tokenUrl = `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    client_secret: env.MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const resp = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  if (!resp.ok) {
    const errText = await resp.text().catch(() => '');
    console.error('Token acquisition failed:', resp.status, errText);
    throw new Error(`获取 Microsoft access_token 失败（${resp.status}）`);
  }

  const data = (await resp.json()) as {
    access_token: string;
    expires_in: number;
  };

  tokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };

  return data.access_token;
}

function tempFileName(originalName: string): string {
  const ts = Date.now();
  const rand = Math.random().toString(36).slice(2, 8);
  const baseName = originalName.replace(/\.[^.]+$/, '');
  return `${TEMP_FOLDER}/${baseName}_${ts}_${rand}.docx`;
}

async function ensureTempFolder(
  accessToken: string,
  userId: string
): Promise<void> {
  const url = `${GRAPH_BASE}/users/${userId}/drive/items/root:/${TEMP_FOLDER}`;

  const resp = await fetch(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (resp.ok) return;

  if (resp.status === 404) {
    const createUrl = `${GRAPH_BASE}/users/${userId}/drive/items/root/children`;
    const createResp = await fetch(createUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        name: TEMP_FOLDER,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'rename',
      }),
    });

    if (!createResp.ok) {
      const errText = await createResp.text().catch(() => '');
      console.error(
        'Failed to create temp folder:',
        createResp.status,
        errText
      );
      throw new Error('无法创建 OneDrive 临时文件夹');
    }
  } else {
    throw new Error('无法访问 OneDrive 临时文件夹');
  }
}

async function uploadFile(
  accessToken: string,
  userId: string,
  filePath: string,
  fileData: ArrayBuffer
): Promise<string> {
  const MAX_SIMPLE_UPLOAD = 4 * 1024 * 1024;

  if (fileData.byteLength <= MAX_SIMPLE_UPLOAD) {
    return simpleUpload(accessToken, userId, filePath, fileData);
  } else {
    return resumableUpload(accessToken, userId, filePath, fileData);
  }
}

async function simpleUpload(
  accessToken: string,
  userId: string,
  filePath: string,
  fileData: ArrayBuffer
): Promise<string> {
  const url = `${GRAPH_BASE}/users/${userId}/drive/items/root:/${encodeURIComponent(filePath)}:/content`;

  const resp = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    },
    body: fileData,
  });

  if (!resp.ok) {
    const errText = await resp.text().catch(() => '');
    console.error('Simple upload failed:', resp.status, errText);
    throw new Error(`上传文件到 OneDrive 失败（${resp.status}）`);
  }

  const item = (await resp.json()) as { id: string };
  return item.id;
}

async function resumableUpload(
  accessToken: string,
  userId: string,
  filePath: string,
  fileData: ArrayBuffer
): Promise<string> {
  const sessionUrl = `${GRAPH_BASE}/users/${userId}/drive/items/root:/${encodeURIComponent(filePath)}:/createUploadSession`;

  const sessionResp = await fetch(sessionUrl, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      item: { '@microsoft.graph.conflictBehavior': 'replace' },
    }),
  });

  if (!sessionResp.ok) {
    const errText = await sessionResp.text().catch(() => '');
    console.error('Create upload session failed:', sessionResp.status, errText);
    throw new Error(`创建上传会话失败（${sessionResp.status}）`);
  }

  const sessionData = (await sessionResp.json()) as { uploadUrl: string };
  const uploadUrl = sessionData.uploadUrl;

  const CHUNK_SIZE = 4 * 1024 * 1024;
  const totalSize = fileData.byteLength;
  let fileId = '';

  for (let offset = 0; offset < totalSize; offset += CHUNK_SIZE) {
    const end = Math.min(offset + CHUNK_SIZE, totalSize);
    const chunk = fileData.slice(offset, end);

    const chunkResp = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': String(end - offset),
        'Content-Range': `bytes ${offset}-${end - 1}/${totalSize}`,
      },
      body: chunk,
    });

    if (!chunkResp.ok) {
      const errText = await chunkResp.text().catch(() => '');
      console.error('Chunk upload failed:', chunkResp.status, errText);
      throw new Error(`分片上传失败（${chunkResp.status}）`);
    }

    if (end >= totalSize) {
      const item = (await chunkResp.json()) as { id: string };
      fileId = item.id;
    }
  }

  if (!fileId) {
    throw new Error('上传完成但未获取到文件 ID');
  }

  return fileId;
}

async function convertToPdf(
  accessToken: string,
  userId: string,
  itemId: string
): Promise<ArrayBuffer> {
  const url = `${GRAPH_BASE}/users/${userId}/drive/items/${itemId}/content?format=pdf`;

  const resp = await fetch(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!resp.ok) {
    const errText = await resp.text().catch(() => '');
    console.error('PDF conversion failed:', resp.status, errText);

    if (resp.status === 429) {
      throw new Error('Microsoft API 请求频率超限，请稍后重试');
    }

    throw new Error(
      `Word 转 PDF 失败（${resp.status}），请检查文件是否损坏`
    );
  }

  return await resp.arrayBuffer();
}

async function deleteTempFile(
  accessToken: string,
  userId: string,
  itemId: string
): Promise<void> {
  const url = `${GRAPH_BASE}/users/${userId}/drive/items/${itemId}`;
  try {
    await fetch(url, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${accessToken}` },
    });
  } catch (err) {
    console.error('Failed to delete temp file (non-critical):', err);
  }
}

export const onRequestPost: PagesFunction<Env> = async (context) => {
  const env = context.env;

  if (
    !env.MS_CLIENT_ID ||
    !env.MS_CLIENT_SECRET ||
    !env.MS_TENANT_ID ||
    !env.MS_USER_ID
  ) {
    console.error(
      'Missing required environment variables for Microsoft Graph API'
    );
    return new Response(
      JSON.stringify({
        error: '服务端配置不完整，请联系管理员配置 Microsoft Graph API 环境变量',
      }),
      {
        status: 500,
        headers: { 'Content-Type': 'application/json' },
      }
    );
  }

  let filePath = '';
  let itemId = '';

  try {
    const formData = await context.request.formData();
    const file = formData.get('file');

    if (!file) {
      return new Response(
        JSON.stringify({ error: '未收到文件' }),
        {
          status: 400,
          headers: { 'Content-Type': 'application/json' },
        }
      );
    }

    const fileData = await file.arrayBuffer();
    const originalName = (file as File).name || 'document.docx';

    const accessToken = await getAccessToken(env);
    await ensureTempFolder(accessToken, env.MS_USER_ID);

    filePath = tempFileName(originalName);
    itemId = await uploadFile(accessToken, env.MS_USER_ID, filePath, fileData);

    const pdfBuffer = await convertToPdf(accessToken, env.MS_USER_ID, itemId);

    await deleteTempFile(accessToken, env.MS_USER_ID, itemId);
    itemId = '';

    return new Response(pdfBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Cache-Control': 'no-cache',
      },
    });
  } catch (err: any) {
    console.error('Pages Function error:', err);

    if (itemId) {
      try {
        const accessToken = await getAccessToken(env);
        await deleteTempFile(accessToken, env.MS_USER_ID, itemId);
      } catch {}
    }

    const message = err.message || '未知错误';
    const retryable =
      message.includes('频率超限') ||
      message.includes('429') ||
      message.includes('503');

    return new Response(
      JSON.stringify({ error: message, retryable }),
      {
        status: retryable ? 503 : 502,
        headers: { 'Content-Type': 'application/json' },
      }
    );
  }
};