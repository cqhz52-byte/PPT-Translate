# DeepSeek 文档翻译工具

这是一个可安装到手机桌面的 PWA 网页应用，用 DeepSeek API 将 PPTX/DOCX/PDF 中的文本翻译为中文、英文、日文、韩文、法文、德文、西班牙文等多种语言，并导出新的 PPTX/DOCX/PDF。

## 启动

```powershell
npm start
```

## 生成单网页版

```powershell
npm run build:standalone
```

生成的 `standalone.html` 内置页面样式、PPTX/DOCX/PDF 解析逻辑和 JSZip。它仍建议配合本项目的 `server.js` 或 `api/translate.js` 使用 DeepSeek 翻译接口，因为直接从浏览器调用 DeepSeek 可能被 CORS 拦截，并且会暴露 API Key。

## GitHub 发布

可以把本项目推到 GitHub，但要区分两种发布方式：

- GitHub Pages：只能托管前端页面。当前简化界面默认调用同域 `./api/translate`，因此静态 GitHub Pages 不能直接自动翻译，除非自行改代码接入外部代理。
- Vercel 连接 GitHub 仓库：可以同时发布前端和 `api/translate.js` 后端代理。当前默认配置可直接使用。

如果只想做静态 GitHub Pages，请不要把 DeepSeek API Key 写进代码或仓库；在页面里临时输入即可。

## Vercel 发布

推荐方式：

1. 把本目录推送到 GitHub 仓库。
2. 打开 Vercel，选择 `Add New Project`。
3. Import 这个 GitHub 仓库。
4. Framework Preset 选择 `Other`。
5. Build Command 留空或保持默认空值。
6. Output Directory 保持 `.`。
7. 部署完成后打开 Vercel 分配的域名。

DeepSeek API Key 不要写入代码仓库。默认模式是在页面里临时输入 Key，由浏览器发送给同域的 Vercel Function，再由 Function 转发给 DeepSeek。

如果你要把这个工具公开给别人使用，不建议在 Vercel 环境变量里放自己的 DeepSeek Key，否则别人也会消耗你的额度。

如果是你自己私用，并且已经开启 Vercel Deployment Protection，也可以在 Vercel 项目环境变量里设置：

```text
DEEPSEEK_API_KEY=你的 DeepSeek Key
```

设置后页面的 API Key 输入框可以留空，后端会自动使用环境变量。

## Cloudflare Pages 发布与手机号授权

本项目包含 `functions/` 目录，部署到 Cloudflare Pages 后会自动启用 Pages Functions，用于手机号授权、超级用户后台和 DeepSeek 翻译代理。

推荐方式是在 Cloudflare 里只做一次 KV 绑定，以后都在网页后台管理授权手机号：

```text
PHONE_AUTH_KV
```

部署后打开：

```text
https://你的域名/admin
```

如果是第一次打开，会显示“创建超级用户”。创建后这个初始化入口会自动关闭，之后只能用超级用户手机号和密码登录后台。

超级用户后台支持：

- 添加授权手机号
- 给授权手机号设置登录密码
- 修改授权手机号密码
- 启用或停用授权手机号
- 删除授权手机号

普通用户访问应用时，需要输入后台已授权的手机号和密码。超级用户登录 `/admin` 后也可以直接打开应用使用。

如果你想让后端统一使用自己的 DeepSeek Key，可以在 Cloudflare Pages 环境变量里设置：

```text
DEEPSEEK_API_KEY=你的 DeepSeek Key
```

如果不设置，用户仍可在页面里临时输入 DeepSeek Key。

PDF 推荐使用 Cloudflare 后端调用 LlamaParse 做结构化解析、表格识别和 OCR。要启用该能力，请在 Cloudflare Pages 环境变量中增加：

```text
LLAMAPARSE_API_KEY=你的 LlamaParse / LlamaCloud API Key
```

也可以由超级管理员登录 `/admin`，在“LlamaParse PDF 解析 Key”区域直接输入并保存。后台会把 Key 写入 `PHONE_AUTH_KV` 的 `app:llamaparse_api_key`，普通授权用户无法通过网页端读取明文 Key。

启用后，PDF 导入会先请求 `/api/pdf-parse`，由 Cloudflare Pages Functions 上传 PDF 到 LlamaParse 并取回结构化结果；如果后端未配置或解析失败，前端会自动回退到浏览器内置的 pdf.js 兼容解析。PDF 导出支持两种模式：

- 推荐：重排版译文 PDF。优先使用 LlamaParse 的段落、表格、OCR 文本重新排版，稳定性最高。
- 兼容：原 PDF 原位覆盖。继续保留原 PDF 页面背景，在原文字坐标处遮盖并写入译文；复杂背景、文字长度变化或没有坐标的结构化解析段落仍可能不适合此模式。

可选：如果你想继续用命令行管理，也可以设置 `ADMIN_TOKEN`，然后调用管理接口：

```powershell
# 查看手机号
curl -H "Authorization: Bearer <ADMIN_TOKEN>" https://你的域名/api/admin/phones

# 添加或更新手机号
curl -X POST https://你的域名/api/admin/phones `
  -H "Authorization: Bearer <ADMIN_TOKEN>" `
  -H "Content-Type: application/json" `
  -d "{\"phone\":\"13800138000\",\"password\":\"123456\"}"

# 删除手机号
curl -X DELETE "https://你的域名/api/admin/phones?phone=13800138000" `
  -H "Authorization: Bearer <ADMIN_TOKEN>"
```

兼容旧方式：仍可使用环境变量 `AUTHORIZED_PHONES` 或 `AUTHORIZED_USERS` 配置固定白名单，但这不适合经常增删用户。

注意：只校验“手机号”本身并不能证明使用者拥有这个号码，别人知道授权手机号也可能登录。更安全的做法是给每个手机号配置密码；如果需要短信验证码，需要再接入短信服务商。

启动后在电脑浏览器打开：

```text
http://127.0.0.1:5173
```

手机和电脑连接同一个 Wi-Fi 后，用手机浏览器打开服务端打印的局域网地址，例如：

```text
http://192.168.1.23:5173
```

然后在浏览器菜单里选择“添加到主屏幕”或“安装应用”。

## 使用

1. 上传或拖入 `.pptx`、`.docx` 或 `.pdf` 文件。
2. 选择翻译方向：支持中文到英文、英文到中文，也可自动识别源语言并翻译为中文、英文、日文、韩文、法文、德文、西班牙文、葡萄牙文、意大利文、俄文、阿拉伯文、泰文、越南文或印尼文。
3. PPTX 可选择版式策略和译文字号比例；默认保留原文本框自动换行，翻译后可只调字号比例并重新导出，无需重新翻译。
4. 如只有某一页或某一行排版异常，可在对应译文下方单独调整“本段字号”或勾选“本段单行”；点击“本页”可把该段设置复制到同页文本框，预览和导出都会应用这些局部设置。
5. PPTX 预览中可直接选中文本框，拖动文本框改变位置，拖动右下角改变宽高，或用浮动工具条切换单行/换行、调整本段字号，并可直接从预览窗口导出文件；在后面页面调整时会保持当前预览滚动位置。
6. PPTX 预览会显示图片边界参考框，方便移动文字时避开图片、图标和背景图；这些边界只用于预览，不会改动原图片。
7. 预览中的百分比是相对原始字号的比例，旁边会显示实际 pt 值；导出 PPTX 时会写入这个明确字号，避免 PowerPoint 因继承母版样式而把文字变小。
8. PDF 会提取可选择文本并翻译，导出时保留原 PDF 页面、图片和表格线，并在原文字坐标处覆盖写入译文；普通正文会尽量保留原 PDF 字号，表格会先尝试读取 PDF 绘图线条识别单元格，再按真实单元格宽高自动换行并在必要时压缩译文字号。遮盖原文时只覆盖原文字区域，并在译文写入后重绘识别到的表格边框，避免擦断表格线。PDF 生成时会显示加载字体、回写段落和保存文件的状态进度。扫描版 PDF 需要先 OCR，复杂 PDF 可能需要微调。
9. 填入 DeepSeek API Key。
10. 点击“自动翻译”，翻译过程中同一个按钮会变成“停止翻译”，可随时中断并保留已完成译文。
11. 可先用“预览译文”查看近似 PPT 文本框版式或文本清单；拖动“译文字号比例”会即时刷新预览，再导出文件。

DeepSeek 模型、API Key、翻译方向、PPT 版式策略和译文字号比例会保存在当前浏览器本地。电脑和手机需要各自输入一次。

## 注意

- 支持现代 Office XML 格式：`.pptx` 和 `.docx`；支持可提取文本的 `.pdf` 翻译并导出新的 PDF。
- 旧版 `.ppt` / `.doc` 是二进制格式，浏览器端无法可靠解析和回写。请先用 PowerPoint、Word 或 WPS 另存为 `.pptx` / `.docx`。
- API Key 只在当前页面和本机服务请求中使用，不会写入 PPTX 或本地文件。
- PWA 安装需要通过 `http://localhost`、局域网 HTTP 或 HTTPS 访问，直接双击 HTML 文件无法完整启用安装和离线缓存。
