const fs = require("fs");
const path = require("path");

const root = __dirname;
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");

const css = read("styles.css");
const jszip = read("vendor/jszip.min.js");
const sourceHtml = read("index.html");
const body = sourceHtml.match(/<body>([\s\S]*?)<\/body>/)?.[1]
  .replace(/\s*<script src="vendor\/jszip\.min\.js"><\/script>/, "")
  .replace(/\s*<script src="app\.js\?v=\d+" type="module"><\/script>/, "");
const app = read("app.js")
  .replace(
    /if \("serviceWorker" in navigator\) \{[\s\S]*?\n\}\n\n/,
    ""
  )
  .replace(
    /window\.addEventListener\("beforeinstallprompt"[\s\S]*?els\.installButton\.addEventListener\("click"[\s\S]*?\n\}\);\n\n/,
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
${body}

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
