# DeepSeek 文档翻译工具

这是一个可安装到手机桌面的 PWA 网页应用，用 DeepSeek API 将 PPTX/DOCX/PDF 中的中文和英文文本双向翻译，并导出新的 PPTX/DOCX/PDF。

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
2. 选择翻译方向：中文到英文、英文到中文、自动识别到英文、自动识别到中文。
3. PPTX 可选择版式策略和译文字号比例；默认保留原文本框自动换行，翻译后可只调字号比例并重新导出，无需重新翻译。
4. 如只有某一页或某一行排版异常，可在对应译文下方单独调整“本段字号”或勾选“本段单行”；点击“本页”可把该段设置复制到同页文本框，预览和导出都会应用这些局部设置。
5. PPTX 预览中可直接选中文本框，拖动文本框改变位置，拖动右下角改变宽高，或用浮动工具条切换单行/换行、调整本段字号，并可直接从预览窗口导出文件；在后面页面调整时会保持当前预览滚动位置。
6. PPTX 预览会显示图片边界参考框，方便移动文字时避开图片、图标和背景图；这些边界只用于预览，不会改动原图片。
7. 预览中的百分比是相对原始字号的比例，旁边会显示实际 pt 值；导出 PPTX 时会写入这个明确字号，避免 PowerPoint 因继承母版样式而把文字变小。
8. PDF 会提取可选择文本并翻译，导出时保留原 PDF 页面、图片和表格线，并在原文字坐标处覆盖写入译文；普通正文会尽量保留原 PDF 字号，表格内会按可用单元格宽高自动换行并在必要时压缩译文字号，尽量避免跨格。PDF 生成时会显示加载字体、回写段落和保存文件的状态进度。扫描版 PDF 需要先 OCR，复杂 PDF 可能需要微调。
9. 填入 DeepSeek API Key。
10. 点击“自动翻译”，可先用“预览译文”查看近似 PPT 文本框版式或文本清单；拖动“译文字号比例”会即时刷新预览，再导出文件。

DeepSeek 模型、API Key、翻译方向、PPT 版式策略和译文字号比例会保存在当前浏览器本地。电脑和手机需要各自输入一次。

## 注意

- 支持现代 Office XML 格式：`.pptx` 和 `.docx`；支持可提取文本的 `.pdf` 翻译并导出新的 PDF。
- 旧版 `.ppt` / `.doc` 是二进制格式，浏览器端无法可靠解析和回写。请先用 PowerPoint、Word 或 WPS 另存为 `.pptx` / `.docx`。
- API Key 只在当前页面和本机服务请求中使用，不会写入 PPTX 或本地文件。
- PWA 安装需要通过 `http://localhost`、局域网 HTTP 或 HTTPS 访问，直接双击 HTML 文件无法完整启用安装和离线缓存。
