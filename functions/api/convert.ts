/**
 * Cloudflare Pages Function — Word 转 PDF 代理（Microsoft Graph API 版）
 *
 * 支持两种认证模式，自动根据环境变量选择：
 *
 * ┌─────────────────────┬──────────────────────────────────────────────┐
 * │ 模式                │ 适用场景                                     │
 * ├─────────────────────┼──────────────────────────────────────────────┤
 * │ client_credentials  │ 工作/学校账户（Azure AD 组织租户）             │
 * │ refresh_token       │ 个人微软账户（@outlook.com / @hotmail.com）   │
 * └─────────────────────┴──────────────────────────────────────────────┘
 *
 * 模式自动检测：
 *   - 设置了 MS_CLIENT_SECRET + MS_TENANT_ID → client_credentials 模式
 *   - 设置了 MS_REFRESH_TOKEN                  → refresh_token 模式
 *
 * ─── client_credentials 模式所需环境变量 ───
 *   MS_CLIENT_ID      — Azure AD 应用的客户端 ID
 *   MS_CLIENT_SECRET  — Azure AD 应用的客户端密钥
 *   MS_TENANT_ID      — Azure AD 租户 ID
 *   MS_USER_ID        — OneDrive 所属用户的 UPN 或 Object ID
 *
 * ─── refresh_token 模式所需环境变量 ───
 *   MS_CLIENT_ID      — Azure AD 应用的客户端 ID
 *   MS_REFRESH_TOKEN  — 预先获取的 refresh_token
 *
 * 转换流程：
 *   1. 获取 Microsoft Graph access_token
 *   2. 将 .docx 上传到 OneDrive 临时目录
 *   3. 调用 Graph API 以 PDF 格式下载（微软服务端引擎转换）
 *   4. 删除 OneDrive 上的临时文件
 *   5. 返回 PDF 给前端
 */

// ─── 类型定义 ────────────────────────────────────────────────────────

interface Env {
  MS_CLIENT_ID: string;
  MS_CLIENT_SECRET?: string;
  MS_TENANT_ID?: string;
  MS_USER_ID?: string;
  MS_REFRESH_TOKEN?: string;
}

type AuthMode = 'client_credentials' | 'refresh_token';

interface TokenCache {
  token: string;
  expiresAt: number; // epoch ms
}

// ─── 全局 Token 缓存 ────────────────────────────────────────────────

let tokenCache: TokenCache | null = null;

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const TEMP_FOLDER = 'pdfmerger-temp';

// ─── 模式检测 ────────────────────────────────────────────────────────

function detectAuthMode(env: Env): AuthMode | null {
  if (env.MS_CLIENT_SECRET && env.MS_TENANT_ID) return 'client_credentials';
  if (env.MS_REFRESH_TOKEN) return 'refresh_token';
  return null;
}

// ─── 获取 Access Token ────────────────────────────────────────────────

async function getAccessToken(env: Env, mode: AuthMode): Promise<string> {
  // 缓存未过期直接返回
  if (tokenCache && tokenCache.expiresAt > Date.now() + 60_000) {
    return tokenCache.token;
  }

  if (mode === 'client_credentials') {
    return getClientCredentialsToken(env);
  } else {
    return getRefreshTokenFlowToken(env);
  }
}

/**
 * 客户端凭据模式 — 适用于工作/学校账户（Azure AD 组织租户）
 */
async function getClientCredentialsToken(env: Env): Promise<string> {
  const tokenUrl = `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    client_secret: env.MS_CLIENT_SECRET!,
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
    console.error('Client credentials token failed:', resp.status, errText);
    throw new Error(`获取 access_token 失败（组织账户模式，${resp.status}）`);
  }

  const data = (await resp.json()) as { access_token: string; expires_in: number };

  tokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };

  return data.access_token;
}

/**
 * Refresh Token 模式 — 适用于个人微软账户（@outlook.com / @hotmail.com）
 */
async function getRefreshTokenFlowToken(env: Env): Promise<string> {
  const tokenUrl = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

  const body = new URLSearchParams({
    client_id: env.MS_CLIENT_ID,
    refresh_token: env.MS_REFRESH_TOKEN!,
    scope: 'https://graph.microsoft.com/Files.ReadWrite offline_access',
    grant_type: 'refresh_token',
  });

  const resp = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  if (!resp.ok) {
    const errText = await resp.text().catch(() => '');
    console.error('Refresh token exchange failed:', resp.status, errText);

    // 如果 refresh_token 过期，给出明确提示
    if (resp.status === 400 || resp.status === 401) {
      throw new Error(
        'refresh_token 已过期或无效，请重新获取。运行 node scripts/get-refresh-token.js'
      );
    }

    throw new Error(`获取 access_token 失败（个人账户模式，${resp.status}）`);
  }

  const data = (await resp.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };

  tokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };

  return data.access_token;
}

// ─── Graph API 路径构建 ────────────────────────────────────────────────
// client_credentials 模式: /users/{userId}/drive/...
// refresh_token 模式:      /me/drive/...

function driveRoot(mode: AuthMode, userId?: string): string {
  if (mode === 'refresh_token') {
    return `${GRAPH_BASE}/me/drive`;
  }
  return `${GRAPH_BASE}/users/${userId}/drive`;
}

// ─── 生成唯一临时文件名 ────────────────────────────────────────────────

function tempFileName(originalName: string): string {
  const ts = Date.now();
  const rand = Math.random().toString(36).slice(2, 8);
  const baseName = originalName.replace(/\.[^.]+$/, '');
  return `${TEMP_FOLDER}/${baseName}_${ts}_${rand}.docx`;
}

// ─── 确保临时文件夹存在 ────────────────────────────────────────────────

async function ensureTempFolder(
  accessToken: string,
  mode: AuthMode,
  userId?: string
): Promise<void> {
  const root = driveRoot(mode, userId);
  const url = `${root}/items/root:/${TEMP_FOLDER}`;

  const resp = await fetch(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (resp.ok) return;

  if (resp.status === 404) {
    const createUrl = `${root}/items/root/children`;
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
      console.error('Failed to create temp folder:', createResp.status, errText);
      throw new Error('无法创建 OneDrive 临时文件夹');
    }
  } else {
    const errText = await resp.text().catch(() => '');
    console.error('Failed to check temp folder:', resp.status, errText);
    throw new Error('无法访问 OneDrive 临时文件夹');
  }
}

// ─── 上传文件到 OneDrive ─────────────────────────────────────────────

async function uploadFile(
  accessToken: string,
  mode: AuthMode,
  userId: string | undefined,
  filePath: string,
  fileData: ArrayBuffer
): Promise<string> {
  const MAX_SIMPLE_UPLOAD = 4 * 1024 * 1024; // 4 MB

  if (fileData.byteLength <= MAX_SIMPLE_UPLOAD) {
    return simpleUpload(accessToken, mode, userId, filePath, fileData);
  } else {
    return resumableUpload(accessToken, mode, userId, filePath, fileData);
  }
}

async function simpleUpload(
  accessToken: string,
  mode: AuthMode,
  userId: string | undefined,
  filePath: string,
  fileData: ArrayBuffer
): Promise<string> {
  const root = driveRoot(mode, userId);
  const url = `${root}/items/root:/${encodeURIComponent(filePath)}:/content`;

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
  mode: AuthMode,
  userId: string | undefined,
  filePath: string,
  fileData: ArrayBuffer
): Promise<string> {
  const root = driveRoot(mode, userId);
  const sessionUrl = `${root}/items/root:/${encodeURIComponent(filePath)}:/createUploadSession`;

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

// ─── 转换为 PDF 并下载 ────────────────────────────────────────────────

async function convertToPdf(
  accessToken: string,
  mode: AuthMode,
  userId: string | undefined,
  itemId: string
): Promise<ArrayBuffer> {
  const root = driveRoot(mode, userId);
  const url = `${root}/items/${itemId}/content?format=pdf`;

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

// ─── 删除 OneDrive 临时文件 ────────────────────────────────────────────

async function deleteTempFile(
  accessToken: string,
  mode: AuthMode,
  userId: string | undefined,
  itemId: string
): Promise<void> {
  const root = driveRoot(mode, userId);
  const url = `${root}/items/${itemId}`;

  try {
    await fetch(url, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${accessToken}` },
    });
  } catch (err) {
    console.error('Failed to delete temp file (non-critical):', err);
  }
}

// ─── Pages Function 入口 ────────────────────────────────────────────────

export const onRequestPost: PagesFunction<Env> = async (context) => {
  const env = context.env;

  // 检测认证模式
  const authMode = detectAuthMode(env);

  if (!authMode) {
    return new Response(
      JSON.stringify({
        error:
          '服务端配置不完整。请设置以下环境变量之一：\n' +
          '① 组织账户：MS_CLIENT_ID + MS_CLIENT_SECRET + MS_TENANT_ID + MS_USER_ID\n' +
          '② 个人账户：MS_CLIENT_ID + MS_REFRESH_TOKEN',
      }),
      {
        status: 500,
        headers: { 'Content-Type': 'application/json' },
      }
    );
  }

  if (!env.MS_CLIENT_ID) {
    return new Response(
      JSON.stringify({ error: '缺少 MS_CLIENT_ID 环境变量' }),
      {
        status: 500,
        headers: { 'Content-Type': 'application/json' },
      }
    );
  }

  let filePath = '';
  let itemId = '';

  try {
    // 1. 解析前端请求
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

    // 2. 获取 access_token
    const accessToken = await getAccessToken(env, authMode);

    // 3. 确保临时文件夹存在
    await ensureTempFolder(accessToken, authMode, env.MS_USER_ID);

    // 4. 上传到 OneDrive
    filePath = tempFileName(originalName);
    itemId = await uploadFile(
      accessToken,
      authMode,
      env.MS_USER_ID,
      filePath,
      fileData
    );

    // 5. 转换为 PDF
    const pdfBuffer = await convertToPdf(
      accessToken,
      authMode,
      env.MS_USER_ID,
      itemId
    );

    // 6. 删除临时文件
    await deleteTempFile(accessToken, authMode, env.MS_USER_ID, itemId);
    itemId = ''; // 标记已删除

    // 7. 返回 PDF
    return new Response(pdfBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Cache-Control': 'no-cache',
      },
    });
  } catch (err: any) {
    console.error('Pages Function error:', err);

    // 尝试清理临时文件
    if (itemId) {
      try {
        const accessToken = await getAccessToken(env, authMode!);
        await deleteTempFile(accessToken, authMode!, env.MS_USER_ID, itemId);
      } catch {
        // 清理失败不影响错误返回
      }
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
