#!/usr/bin/env node

/**
 * get-refresh-token.js — 获取 Microsoft Graph API 的 refresh_token
 *
 * 使用 Device Code Flow（设备代码流），无需本地服务器，无需回调 URL。
 * 只需在浏览器中访问一个链接并输入代码即可。
 *
 * 使用方法：
 *   1. 先在 Azure 门户注册应用（见下方说明）
 *   2. 运行：node scripts/get-refresh-token.js <CLIENT_ID>
 *   3. 按照提示操作
 *   4. 将输出的 refresh_token 配置到 Cloudflare 环境变量
 *
 * ─── Azure 应用注册说明 ───
 *
 * 如果你因为个人账户无法在 Azure 门户注册应用，请按以下步骤操作：
 *
 * 方案 A：使用 Azure 门户（推荐）
 *   1. 打开 https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
 *   2. 如果提示 "interaction_required" 错误，请：
 *      - 退出个人账户
 *      - 先创建免费 Azure 账户：https://azure.microsoft.com/free/
 *      - 用个人账户注册 Azure 免费账户（这会自动创建一个 Azure AD 租户）
 *      - 然后再注册应用
 *
 * 方案 B：使用 Microsoft 365 开发者计划（免费 E5 许可证）
 *   1. 访问 https://developer.microsoft.com/en-us/microsoft-365/dev-program
 *   2. 用个人微软账户加入开发者计划
 *   3. 创建 E5 沙盒（会给你一个 xxx.onmicrosoft.com 的组织账户）
 *   4. 用这个组织账户登录 Azure 门户注册应用
 *
 * ─── 注册应用时的设置 ───
 *
 * 1. 新注册应用
 * 2. 名称：pdfmerger
 * 3. 支持的账户类型：选「任何组织目录中的账户和个人 Microsoft 账户」
 * 4. 重定向 URI：选「公共客户端/本机」，填 http://localhost
 * 5. 注册后，复制「应用程序(客户端) ID」
 * 6. 进入 API 权限 → 添加权限 → Microsoft Graph → 委托的权限 → Files.ReadWrite
 * 7. 不需要管理员同意（个人账户自己同意即可）
 *
 * ⚠️ 重要：选择「委托的权限」而不是「应用程序权限」！
 */

const CLIENT_ID = process.argv[2];

if (!CLIENT_ID) {
  console.error('用法: node get-refresh-token.js <CLIENT_ID>');
  console.error('');
  console.error('请提供 Azure AD 应用的客户端 ID');
  process.exit(1);
}

const TENANT = 'consumers'; // 个人微软账户使用 consumers
const SCOPE = 'https://graph.microsoft.com/Files.ReadWrite offline_access';

async function main() {
  // ─── Step 1: 请求设备代码 ──────────────────────────────────────────
  console.log('正在请求设备代码...\n');

  const deviceCodeUrl = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/devicecode`;

  const codeResp = await fetch(deviceCodeUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      scope: SCOPE,
    }).toString(),
  });

  if (!codeResp.ok) {
    const errText = await codeResp.text().catch(() => '');
    console.error('请求设备代码失败:', codeResp.status, errText);
    console.error('');
    console.error('可能原因：');
    console.error('  - CLIENT_ID 不正确');
    console.error('  - 应用未注册或未配置为支持个人账户');
    process.exit(1);
  }

  const codeData = await codeResp.json() as {
    user_code: string;
    device_code: string;
    verification_uri: string;
    expires_in: number;
    interval: number;
    message: string;
  };

  // ─── Step 2: 显示用户提示 ──────────────────────────────────────────
  console.log('══════════════════════════════════════════════════════');
  console.log('');
  console.log('  请在浏览器中完成以下操作：');
  console.log('');
  console.log(`  1. 访问: ${codeData.verification_uri}`);
  console.log(`  2. 输入代码: ${codeData.user_code}`);
  console.log('');
  console.log('  然后用你的微软个人账户登录并同意授权');
  console.log('');
  console.log('══════════════════════════════════════════════════════');
  console.log('');
  console.log(`代码有效期: ${codeData.expires_in / 60} 分钟`);
  console.log('等待授权中...\n');

  // ─── Step 3: 轮询等待用户完成授权 ──────────────────────────────────
  const tokenUrl = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;
  const pollInterval = codeData.interval || 5;
  const deadline = Date.now() + codeData.expires_in * 1000;

  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, pollInterval * 1000));

    const tokenResp = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        grant_type: 'urn:ietf:params:oauth2:grant-type:device_code',
        device_code: codeData.device_code,
      }).toString(),
    });

    const tokenData = await tokenResp.json() as {
      access_token?: string;
      refresh_token?: string;
      expires_in?: number;
      error?: string;
      error_description?: string;
    };

    if (tokenData.access_token && tokenData.refresh_token) {
      // ─── 成功！ ────────────────────────────────────────────────────
      console.log('✅ 授权成功！\n');
      console.log('══════════════════════════════════════════════════════');
      console.log('');
      console.log('  将以下值配置到 Cloudflare Pages 环境变量：');
      console.log('');
      console.log(`  MS_CLIENT_ID     = ${CLIENT_ID}`);
      console.log(`  MS_REFRESH_TOKEN = ${tokenData.refresh_token}`);
      console.log('');
      console.log('══════════════════════════════════════════════════════');
      console.log('');
      console.log('⚠️  注意事项：');
      console.log('  - refresh_token 请妥善保管，不要泄露');
      console.log('  - refresh_token 有效期约 90 天（不使用时过期）');
      console.log('  - 如果 refresh_token 过期，重新运行此脚本即可');
      console.log('');

      // 验证 access_token 是否能访问 OneDrive
      console.log('正在验证 OneDrive 访问权限...');
      try {
        const driveResp = await fetch('https://graph.microsoft.com/v1.0/me/drive', {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
        });

        if (driveResp.ok) {
          const driveInfo = await driveResp.json() as { driveType: string; owner?: { user?: { displayName: string } } };
          console.log(`✅ OneDrive 验证成功！类型: ${driveInfo.driveType}`);
          if (driveInfo.owner?.user?.displayName) {
            console.log(`   账户: ${driveInfo.owner.user.displayName}`);
          }
        } else {
          console.log('⚠️  OneDrive 访问失败，请确认应用已添加 Files.ReadWrite 权限');
        }
      } catch {
        console.log('⚠️  无法验证 OneDrive 访问（网络问题），但 refresh_token 应该有效');
      }

      return;
    }

    // 检查错误类型
    if (tokenData.error === 'authorization_pending') {
      // 用户还没有完成授权，继续等待
      process.stdout.write('.');
      continue;
    }

    if (tokenData.error === 'slow_down') {
      // 需要减慢轮询频率
      await new Promise((r) => setTimeout(r, pollInterval * 1000));
      continue;
    }

    if (tokenData.error === 'expired_token') {
      console.error('\n❌ 设备代码已过期，请重新运行此脚本');
      process.exit(1);
    }

    if (tokenData.error === 'access_denied') {
      console.error('\n❌ 授权被拒绝');
      process.exit(1);
    }

    // 其他错误
    console.error('\n❌ 获取令牌失败:', tokenData.error, tokenData.error_description);
    process.exit(1);
  }

  console.error('\n❌ 超时，请重新运行此脚本');
  process.exit(1);
}

main().catch((err) => {
  console.error('脚本执行出错:', err);
  process.exit(1);
});
