---
name: free-deploy-webapp
description: 如何开发可免费部署到 Cloudflare Pages 的 Web 应用。当用户想要开发一个零成本部署的网站、免费托管的 Web 工具、使用免费 AI API 的应用、或任何需要前后端但不想花钱部署的项目时，使用此 skill。即使用户没有明确说"免费"，只要涉及部署且预算为零，都应触发。
---

# Skill：让 AI 开发可免费部署的 AI 驱动 Web 应用

> 从 TextLens（图文识别工具）项目中提炼的方法论，适用于任何需要 AI 能力、前端交互、且全链路零成本部署的 Web 项目。

---

## 一、核心架构模式

### 1.1 架构选型：静态导出 + Serverless Functions

```
┌─────────────────────────────────────────────────┐
│                  用户浏览器                       │
│  Next.js 静态页面 (HTML/CSS/JS)                  │
│  - UI 交互、图片上传/压缩                         │
│  - 结果展示、格式转换、剪贴板操作                   │
└────────────────────┬────────────────────────────┘
                     │ fetch /api/xxx
                     ▼
┌─────────────────────────────────────────────────┐
│         Cloudflare Pages Functions               │
│  functions/api/xxx/index.ts                      │
│  - 接收请求，调用第三方 AI API                     │
│  - 重试机制、错误处理、模型降级                     │
│  - 环境变量读取 API Key                           │
└────────────────────┬────────────────────────────┘
                     │ HTTPS
                     ▼
┌─────────────────────────────────────────────────┐
│            第三方 AI API（免费额度）               │
│  智谱AI / Groq / Google Gemini 等                │
└─────────────────────────────────────────────────┘
```

**为什么不用 Next.js API Routes？**
- Next.js `output: "export"` 静态导出不支持 API Routes
- Cloudflare Pages Functions 完美替代，零成本，自动与静态页面同域部署
- Functions 目录：`functions/api/xxx/index.ts`，自动映射为 `/api/xxx`

**为什么不用 `@cloudflare/next-on-pages`？**
- 该适配器不支持 Next.js 16，会报 ERESOLVE 依赖冲突
- 静态导出 + Pages Functions 是更简单、更稳定的方案

### 1.2 技术栈组合

| 层级 | 技术 | 理由 |
|------|------|------|
| 框架 | Next.js 16 (`output: "export"`) | React 生态，静态导出支持好 |
| 样式 | Tailwind CSS 4 | 原子化 CSS，零运行时 |
| 组件 | shadcn/ui | 高质量可定制，按需引入 |
| 公式渲染 | KaTeX | 轻量快速，支持 MathML 输出 |
| 部署 | Cloudflare Pages | 免费额度大，全球 CDN，支持 Functions |
| AI | 智谱AI glm-4.6v-flash | 免费视觉模型，支持图片理解 |

---

## 二、免费资源清单

### 2.1 免费部署平台

| 平台 | 静态托管 | Serverless | 免费额度 | 适合场景 |
|------|---------|-----------|---------|---------|
| **Cloudflare Pages** | ✅ 无限 | ✅ 10万次/天 | 极大 | **首选**，Functions + 静态同域 |
| Vercel | ✅ | ✅ Serverless Functions | 100GB/月 | Next.js 原生支持，但免费额度较小 |
| Netlify | ✅ | ✅ Netlify Functions | 100GB/月 | 类似 CF Pages，但 Functions 限制更多 |
| GitHub Pages | ✅ | ❌ | 无限 | 纯静态项目，无后端能力 |

**推荐：Cloudflare Pages**（免费额度最大，Functions 同域部署无跨域问题）

### 2.2 免费 AI 模型

| 提供商 | 模型 | 能力 | 免费额度 | API 格式 |
|--------|------|------|---------|---------|
| **智谱AI** | glm-4.6v-flash | 文字+图片理解 | 免费 | OpenAI 兼容 |
| **Groq** | llama-3.3-70b | 纯文字 | 免费（限速） | OpenAI 兼容 |
| **Google** | gemini-2.0-flash | 文字+图片 | 免费（限速） | Google 格式 |
| **DeepSeek** | deepseek-chat | 纯文字 | 极低价 | OpenAI 兼容 |

**关键：选择 OpenAI 兼容格式的 API**，这样切换模型只需改 base_url 和 model，代码零改动。

### 2.3 免费 AI 图片生成

| 提供商 | 模型 | 免费额度 |
|--------|------|---------|
| 智谱AI | cogview-4-flash | 免费 |

### 2.4 其他免费资源

| 资源 | 服务 | 用途 |
|------|------|------|
| 域名 | Cloudflare Pages *.pages.dev | 自动分配免费子域名 |
| 代码托管 | GitHub | 代码仓库 + 自动部署触发 |
| 图标 | Lucide React | 开源图标库 |
| 字体 | Google Fonts / Noto | 免费 Web 字体 |

---

## 三、项目文件结构模板

```
project/
├── src/
│   └── app/
│       ├── layout.tsx          # 根布局（字体、全局样式）
│       ├── page.tsx            # 主页面（所有 UI 逻辑）
│       └── globals.css         # Tailwind + 自定义样式
├── functions/                  # Cloudflare Pages Functions
│   └── api/
│       └── xxx/
│           └── index.ts        # API 端点，export onRequestPost
├── components/                 # shadcn/ui 组件
│   └── ui/
│       ├── button.tsx
│       ├── card.tsx
│       └── ...
├── hooks/                      # 自定义 hooks
├── public/                     # 静态资源
├── next.config.ts              # 关键：output: "export"
├── wrangler.toml               # Cloudflare 配置
├── package.json
└── .gitignore
```

---

## 四、关键配置文件

### 4.1 next.config.ts

```typescript
import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "export",           // 核心：静态导出
  typescript: {
    ignoreBuildErrors: true,   // 加速构建
  },
  reactStrictMode: false,
  images: {
    unoptimized: true,         // 静态导出必须关闭图片优化
  },
};

export default nextConfig;
```

### 4.2 wrangler.toml

```toml
name = "your-project-name"
compatibility_date = "2024-09-23"
compatibility_flags = ["nodejs_compat"]        # 必须：Node.js 兼容
pages_build_output_dir = "./out"                # 指向 next build 输出
```

### 4.3 Cloudflare Pages Function 模板

```typescript
// functions/api/xxx/index.ts

export const onRequestPost: PagesFunction = async (context) => {
  try {
    const body = await context.request.json();

    // 1. 读取环境变量
    const apiKey = context.env.AI_API_KEY;
    const baseUrl = context.env.AI_BASE_URL || 'https://open.bigmodel.cn/api/paas/v4';
    const model = context.env.AI_MODEL || 'glm-4.6v-flash';

    if (!apiKey) {
      return new Response(
        JSON.stringify({ success: false, error: '未配置 API 密钥' }),
        { status: 500, headers: { 'Content-Type': 'application/json' } }
      );
    }

    // 2. 调用 AI API
    const response = await fetch(`${baseUrl}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        messages: [
          {
            role: 'user',
            content: [
              // 多模态内容放 user 消息，不放 system
              { type: 'text', text: '你的提示词...' },
              // { type: 'image_url', image_url: { url: base64Data } }
            ],
          },
        ],
        temperature: 0.1,
        max_tokens: 8192,
      }),
    });

    if (!response.ok) {
      // 429 限流重试逻辑
      // ...
    }

    const data = await response.json();
    const content = data.choices?.[0]?.message?.content || '';

    return new Response(
      JSON.stringify({ success: true, result: content }),
      { headers: { 'Content-Type': 'application/json' } }
    );
  } catch (error) {
    return new Response(
      JSON.stringify({ success: false, error: '服务错误' }),
      { status: 500, headers: { 'Content-Type': 'application/json' } }
    );
  }
};
```

---

## 五、踩过的坑与解决方案

### 5.1 Git 子模块问题

**现象**：`git push` 后 Cloudflare 构建报错 `formula-lens-reference` 子模块找不到

**原因**：项目目录中包含了另一个 git 仓库的目录，被自动识别为子模块

**解决**：
```bash
git rm --cached formula-lens-reference
echo "formula-lens-reference/" >> .gitignore
git add . && git commit -m "fix: remove submodule"
```

### 5.2 @cloudflare/next-on-pages 不兼容 Next.js 16

**现象**：`npm install` 报 ERESOLVE 依赖冲突

**原因**：该适配器尚未支持 Next.js 16

**解决**：放弃该适配器，改用 `output: "export"` + Cloudflare Pages Functions

### 5.3 AI API 400 错误（消息格式）

**现象**：API 返回 400，参数有误

**原因**：system 消息不能包含 multimodal 内容（图片）

**解决**：把图片和提示词都放在 `user` 消息中：
```typescript
messages: [
  // ❌ 错误：system 不能包含图片
  // { role: 'system', content: [{ type: 'image_url', ... }] }
  
  // ✅ 正确：图片放 user 消息
  {
    role: 'user',
    content: [
      { type: 'image_url', image_url: { url: imageData } },
      { type: 'text', text: '提示词...' },
    ],
  },
]
```

### 5.4 Base64 格式问题

**现象**：API 返回 400

**原因**：发送了完整的 data URL（`data:image/png;base64,xxxx`），API 只需要纯 base64

**解决**：
```typescript
function extractBase64(dataUrl: string): string {
  if (dataUrl.startsWith('data:image/')) {
    const commaIndex = dataUrl.indexOf(',');
    if (commaIndex !== -1) return dataUrl.substring(commaIndex + 1);
  }
  return dataUrl;
}
```

### 5.4 模型名称变更

**现象**：API 返回 400，模型不存在

**原因**：智谱AI 将 `glm-4v-flash` 升级为 `glm-4.6v-flash`，旧名称失效

**解决**：
- 使用环境变量配置模型名，方便随时更新
- 实现模型降级链：配置模型 → 备用模型
- 关注 API 提供商的公告

### 5.5 429 限流

**现象**：高峰期频繁 429

**解决**：指数退避重试
```typescript
if (response.status === 429) {
  const waitTime = attempt * 2000; // 2s, 4s, 6s
  await sleep(waitTime);
  continue; // 重试
}
```

---

## 六、Cloudflare Pages 部署配置

### 构建设置

| 配置项 | 值 |
|--------|-----|
| Framework preset | Next.js (Static HTML Export) |
| Build command | `npx next build` |
| Build output directory | `out` |

### 环境变量

| 变量名 | 示例值 | 说明 |
|--------|--------|------|
| `AI_API_KEY` | `your-api-key` | AI API 密钥 |
| `AI_BASE_URL` | `https://open.bigmodel.cn/api/paas/v4` | API 基础 URL |
| `AI_MODEL` | `glm-4.6v-flash` | 模型名称 |

### 重要提醒

- 环境变量设置后需要**重新部署**才会生效
- `functions/` 目录必须包含在 Git 仓库中（不是 build 产物）
- Cloudflare Pages 自动将 `functions/api/xxx/index.ts` 映射为 `/api/xxx` 端点

---

## 七、给 AI 的提示词模板

当你想让 AI 帮你开发一个类似的免费部署项目时，使用以下提示词结构：

```
请帮我开发一个 [项目名称] Web 应用，要求：

## 功能需求
[描述核心功能，例如：用户粘贴图片后识别文字和公式...]

## 技术要求
1. 使用 Next.js 16 + Tailwind CSS + shadcn/ui
2. 配置 output: "export" 静态导出
3. 后端使用 Cloudflare Pages Functions（functions/ 目录）
4. AI 能力调用 [智谱AI / Groq / Gemini] 的免费 API
5. 环境变量：AI_API_KEY、AI_BASE_URL、AI_MODEL
6. 不使用数据库，纯前端 + Serverless

## 关键约束
1. 全链路零成本：部署平台（Cloudflare Pages）、AI API（免费模型）均不收费
2. 图片等用户输入数据在前端压缩后发送
3. AI API 调用格式必须使用 OpenAI 兼容格式
4. system 消息不能包含 multimodal 内容，图片必须放 user 消息
5. 发送给 AI 的 base64 必须去掉 data URL 前缀
6. 实现 429 限流重试机制（指数退避）
7. next.config.ts 必须配置 images: { unoptimized: true }
8. wrangler.toml 必须包含 compatibility_flags = ["nodejs_compat"]

## 部署配置
- 构建命令：npx next build
- 输出目录：out
- 环境变量在 Cloudflare Pages 控制台设置

## UI 风格
[描述你想要的风格，例如：简洁现代，绿色主题...]
```

---

## 八、完整开发 → 部署流程

```
1. 让 AI 生成项目代码
       ↓
2. 本地解压到项目目录
       ↓
3. git init → git add . → git commit → git push
       ↓
4. Cloudflare Pages 连接 GitHub 仓库
       ↓
5. 设置构建命令和输出目录
       ↓
6. 设置环境变量（AI_API_KEY 等）
       ↓
7. 触发构建部署
       ↓
8. 访问 *.pages.dev 域名验证
       ↓
9. 如需修改 → 改代码 → push → 自动重新部署
```

**不需要本地安装依赖或构建**，Cloudflare 会在云端完成所有构建。

---

## 九、可复用的设计模式

### 9.1 前端图片压缩

```typescript
const compressImage = (dataUrl: string, maxWidth = 2048, quality = 0.9): Promise<string> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      if (img.width <= maxWidth && dataUrl.length < 4 * 1024 * 1024) {
        resolve(dataUrl); // 小图不压缩
        return;
      }
      const scale = maxWidth / img.width;
      const canvas = document.createElement('canvas');
      canvas.width = maxWidth;
      canvas.height = Math.round(img.height * scale);
      const ctx = canvas.getContext('2d');
      ctx?.drawImage(img, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL('image/png', quality));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
};
```

### 9.2 三种图片输入方式

```typescript
// 1. Ctrl+V 粘贴（全局监听）
useEffect(() => {
  document.addEventListener('paste', handlePaste);
  return () => document.removeEventListener('paste', handlePaste);
}, []);

// 2. 拖拽上传
<div onDrop={handleDrop} onDragOver={handleDragOver} onDragLeave={handleDragLeave} />

// 3. 点击选择文件
<input type="file" accept="image/*" onChange={handleFileSelect} className="hidden" />
```

### 9.3 多格式复制（Word HTML + 剪贴板）

```typescript
// Word 格式需要同时写入 text/html 和 text/plain
if (selectedFormat === 'word') {
  const htmlBlob = new Blob([htmlContent], { type: 'text/html' });
  const textBlob = new Blob([textContent], { type: 'text/plain' });
  const clipboardItem = new ClipboardItem({
    'text/html': htmlBlob,
    'text/plain': textBlob,
  });
  await navigator.clipboard.write([clipboardItem]);
} else {
  await navigator.clipboard.writeText(formattedOutput);
}
```

### 9.4 环境变量驱动的 AI 配置

```typescript
// 前端不需要知道 API Key，所有 AI 调用走后端
const response = await fetch('/api/recognize', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ image: rawImageData }),
});

// 后端从环境变量读取
const apiKey = context.env.AI_API_KEY;
const baseUrl = context.env.AI_BASE_URL;
const model = context.env.AI_MODEL;
```

---

## 十、项目类型适配

此 Skill 不仅适用于图文识别，还可用于：

| 项目类型 | AI 能力 | 前端交互 | 免费模型 |
|---------|--------|---------|---------|
| 图片文字识别 | 视觉理解 | 粘贴/拖拽图片 | glm-4.6v-flash |
| AI 翻译工具 | 文本生成 | 输入文本 | glm-4-flash / groq |
| 图片描述生成 | 视觉理解 | 上传图片 | glm-4.6v-flash |
| 文本摘要工具 | 文本生成 | 粘贴长文 | glm-4-flash |
| 代码解释器 | 代码生成 | 代码输入 | deepseek / groq |
| AI 对话助手 | 文本生成 | 聊天界面 | glm-4-flash |
| Markdown 编辑器 | 文本润色 | 富文本编辑 | glm-4-flash |

**核心公式：前端交互 → Cloudflare Function → 免费 AI API → 结果展示 → 多格式输出**

