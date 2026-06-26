const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 5173);
const HOST = process.env.HOST || "0.0.0.0";
const MAX_BODY_BYTES = 1024 * 1024;

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".webmanifest": "application/manifest+json; charset=utf-8",
  ".svg": "image/svg+xml; charset=utf-8",
  ".otf": "font/otf",
  ".ttf": "font/ttf",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".pdf": "application/pdf",
};

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === "OPTIONS") {
      sendOptions(res);
      return;
    }

    if (req.method === "POST" && req.url === "/api/translate") {
      await handleTranslate(req, res);
      return;
    }

    if (req.method === "GET" || req.method === "HEAD") {
      serveStatic(req, res);
      return;
    }

    sendJson(res, 405, { error: "Method not allowed" });
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Server error" });
  }
});

server.listen(PORT, HOST, () => {
  console.log("PPT translator running:");
  getAccessUrls(PORT).forEach((url) => console.log(`  ${url}`));
});

async function handleTranslate(req, res) {
  const body = await readJson(req);
  const apiBase = String(body.apiBase || "https://api.deepseek.com").replace(/\/$/, "");
  const apiKey = String(body.apiKey || process.env.DEEPSEEK_API_KEY || "").trim();
  const model = String(body.model || "deepseek-v4-flash").trim();
  const instruction = String(body.instruction || "").trim();
  const text = String(body.text || "");
  const task = String(body.task || "translate").trim();

  if (!apiKey) {
    sendJson(res, 400, { error: "缺少 DeepSeek API Key。" });
    return;
  }

  if (!text.trim()) {
    sendJson(res, 200, { translation: "" });
    return;
  }

  const isSummary = task === "summary";
  const response = await fetch(`${apiBase}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      temperature: 0.2,
      thinking: { type: "disabled" },
      messages: [
        {
          role: "system",
          content: isSummary ? buildSummarySystemPrompt(instruction) : buildTranslationSystemPrompt(instruction),
        },
        {
          role: "user",
          content: text,
        },
      ],
    }),
  });

  if (!response.ok) {
    const detail = await response.text();
    sendJson(res, response.status, {
      error: `DeepSeek 返回 ${response.status}：${detail.slice(0, 300)}`,
    });
    return;
  }

  const data = await response.json();
  sendJson(res, 200, {
    translation: data.choices?.[0]?.message?.content?.trim() || "",
  });
}

function buildTranslationSystemPrompt(instruction) {
  return [
    "You are a professional PowerPoint and Word document translator.",
    instruction,
    "Keep numbers, units, product names, formulas, punctuation intent, dates, and years exactly accurate.",
    "For document layout safety, keep the translation compact and close to the source length.",
    "If the source is short, fragmented, or appears to be a table cell, translate it literally and compactly.",
    "Never say you cannot translate, never ask for source text, and never return apology or explanation text.",
    "For titles, section labels, table-of-contents entries, and short phrases, use short noun phrases instead of full sentences.",
    "Do not add line breaks unless the source text clearly contains line breaks.",
    "For Chinese-to-English, prefer concise business wording over long explanatory sentences.",
    "Do not add explanations, markdown, quotation marks, or labels.",
    "Return only the translated text.",
  ].join(" ");
}

function buildSummarySystemPrompt(instruction) {
  return [
    "You are a careful literature and technical-document analyst.",
    instruction,
    "Summarize only from the provided document text.",
    "Preserve important numbers, years, patent numbers, product names, organization names, abbreviations, and units.",
    "Do not invent facts. If something is not explicit in the text, say it is not specified.",
    "Return the summary directly.",
  ].join(" ");
}

function serveStatic(req, res) {
  const url = new URL(req.url, "http://localhost");
  const decodedPath = decodeURIComponent(url.pathname);
  const requested = decodedPath === "/" ? "/index.html" : decodedPath;
  const filePath = path.resolve(ROOT, `.${requested}`);

  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.stat(filePath, (statError, stat) => {
    if (statError || !stat.isFile()) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }

    const type = MIME_TYPES[path.extname(filePath).toLowerCase()] || "application/octet-stream";
    res.writeHead(200, {
      "Content-Type": type,
      "Cache-Control": "no-cache",
    });

    if (req.method === "HEAD") {
      res.end();
      return;
    }

    fs.createReadStream(filePath).pipe(res);
  });
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    let body = "";

    req.on("data", (chunk) => {
      body += chunk;
      if (Buffer.byteLength(body) > MAX_BODY_BYTES) {
        reject(new Error("请求体过大。"));
        req.destroy();
      }
    });

    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch {
        reject(new Error("JSON 请求格式不正确。"));
      }
    });

    req.on("error", reject);
  });
}

function sendJson(res, status, payload) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
  });
  res.end(JSON.stringify(payload));
}

function sendOptions(res) {
  res.writeHead(204, {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
  });
  res.end();
}

function getAccessUrls(port) {
  const os = require("os");
  const urls = [`http://127.0.0.1:${port}`];
  const interfaces = os.networkInterfaces();

  Object.values(interfaces).forEach((items) => {
    (items || []).forEach((item) => {
      if (item.family === "IPv4" && !item.internal) {
        urls.push(`http://${item.address}:${port}`);
      }
    });
  });

  return urls;
}
