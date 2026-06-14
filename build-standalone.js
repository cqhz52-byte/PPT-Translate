const fs = require("fs");
const path = require("path");

const root = __dirname;
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");

const css = read("styles.css");
const jszip = read("vendor/jszip.min.js");
const app = read("app.js")
  .replace(
    /if \("serviceWorker" in navigator\) \{[\s\S]*?\n\}\n\nwindow\.addEventListener/,
    "window.addEventListener"
  )
  .replace(
    /window\.addEventListener\("beforeinstallprompt",[\s\S]*?\n\}\);\n\n/,
    ""
  );

const html = `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="theme-color" content="#165a72">
    <title>DeepSeek 文档翻译工具 - 单网页版</title>
    <style>
${css}
    </style>
  </head>
  <body>
    <main class="shell">
      <header class="topbar">
        <div>
          <p class="eyebrow">Single HTML</p>
          <h1>DeepSeek 文档翻译工具</h1>
          <p class="version">单网页版 v6 · PPTX/DOCX · 字号优先 · 单行导出</p>
        </div>
        <div class="topbar-actions">
          <button class="icon-button" id="installButton" type="button" hidden title="安装应用" aria-label="安装应用">
            <span aria-hidden="true">⇩</span>
          </button>
          <button class="icon-button" id="resetButton" type="button" title="清空当前文件" aria-label="清空当前文件">
            <span aria-hidden="true">↺</span>
          </button>
        </div>
      </header>

      <section class="workspace" aria-label="文档翻译工作区">
        <aside class="sidebar">
          <label class="upload-zone" for="fileInput">
            <input id="fileInput" type="file" accept=".pptx,.ppt,.docx,.doc,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.ms-powerpoint,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/msword">
            <span class="upload-icon" aria-hidden="true">+</span>
            <span class="upload-title">选择 PPTX / DOCX 文件</span>
            <span class="upload-copy" id="fileMeta">支持 PPTX / DOCX；旧版 PPT/DOC 请先另存</span>
          </label>

          <div class="settings">
            <label>
              翻译方向
              <select id="translationDirection">
                <option value="zh-en" selected>中文 → 英文</option>
                <option value="en-zh">英文 → 中文</option>
                <option value="auto-en">自动识别 → 英文</option>
                <option value="auto-zh">自动识别 → 中文</option>
              </select>
            </label>
            <label>
              PPT 版式策略
              <select id="pptLayoutMode">
                <option value="keep-size" selected>保持字号，单行优先</option>
                <option value="compact-fit">紧凑适配，轻微缩小</option>
              </select>
            </label>
            <label>
              DeepSeek API 地址
              <input id="apiBase" type="url" value="https://api.deepseek.com" spellcheck="false">
            </label>
            <label>
              翻译代理地址
              <input id="apiProxy" type="text" value="./api/translate" spellcheck="false">
            </label>
            <label>
              模型
              <input id="modelName" type="text" value="deepseek-v4-flash" spellcheck="false">
            </label>
            <label>
              DeepSeek API Key
              <input id="apiKey" type="password" placeholder="仅在当前页面使用，不写入文件">
            </label>
          </div>

          <div class="actions">
            <button id="translateButton" class="primary" type="button" disabled>自动翻译</button>
            <button id="copyOriginalButton" type="button" disabled>原文填入译文</button>
            <button id="downloadButton" type="button" disabled>导出翻译文件</button>
          </div>

          <div class="stats" aria-live="polite">
            <div>
              <strong id="slideCount">0</strong>
              <span>页面/部件</span>
            </div>
            <div>
              <strong id="segmentCount">0</strong>
              <span>段文字</span>
            </div>
            <div>
              <strong id="translatedCount">0</strong>
              <span>段译文</span>
            </div>
          </div>
        </aside>

        <section class="panel">
          <div class="panel-head">
            <div>
              <h2>文本清单</h2>
              <p id="statusText">请先选择一个 PPTX 或 DOCX 文件。</p>
            </div>
            <div class="progress" aria-hidden="true">
              <span id="progressFill"></span>
            </div>
          </div>

          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>页码</th>
                  <th>原文</th>
                  <th>译文</th>
                </tr>
              </thead>
              <tbody id="segmentTable">
                <tr class="empty-row">
                  <td colspan="3">
                    <div class="empty-state">
                      <strong>还没有载入内容</strong>
                      <span>上传 PPTX 或 DOCX 后，可逐段翻译、校对并导出新文件。</span>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </section>
      </section>
    </main>

    <script>
${jszip}
    </script>
    <script type="module">
${app}
    </script>
  </body>
</html>
`;

fs.writeFileSync(path.join(root, "standalone.html"), html, "utf8");
console.log("Created standalone.html");
