async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") {
    sendJson(res, 204, null);
    return;
  }

  if (req.method !== "POST") {
    sendJson(res, 405, { error: "Method not allowed" });
    return;
  }

  try {
    const body = await readRequestBody(req);
    const apiBase = String(body.apiBase || "https://api.deepseek.com").replace(/\/$/, "");
    const apiKey = String(body.apiKey || process.env.DEEPSEEK_API_KEY || "").trim();
    const model = String(body.model || "deepseek-v4-flash").trim();
    const instruction = String(body.instruction || "").trim();
    const text = String(body.text || "");

    if (!apiKey) {
      sendJson(res, 400, { error: "缺少 DeepSeek API Key。" });
      return;
    }

    if (!text.trim()) {
      sendJson(res, 200, { translation: "" });
      return;
    }

    if (typeof fetch !== "function") {
      throw new Error("当前 Vercel Node.js 运行时不支持 fetch，请使用 Node.js 18 或更高版本。");
    }

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
            content: [
              "You are a professional PowerPoint and Word document translator.",
              instruction,
              "Keep numbers, units, product names, formulas, punctuation intent, dates, and years exactly accurate.",
              "For document layout safety, keep the translation compact and close to the source length.",
              "For titles, section labels, table-of-contents entries, and short phrases, use short noun phrases instead of full sentences.",
              "Do not add line breaks unless the source text clearly contains line breaks.",
              "For Chinese-to-English, prefer concise business wording over long explanatory sentences.",
              "Do not add explanations, markdown, quotation marks, or labels.",
              "Return only the translated text.",
            ].join(" "),
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
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Server error" });
  }
}

module.exports = handler;
module.exports.default = handler;

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

function sendJson(res, statusCode, payload) {
  res.statusCode = statusCode;
  if (payload === null) {
    res.end();
    return;
  }

  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(payload));
}

async function readRequestBody(req) {
  if (typeof req.body === "string") return JSON.parse(req.body || "{}");
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.on !== "function") return {};

  const chunks = [];
  await new Promise((resolve, reject) => {
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("end", resolve);
    req.on("error", reject);
  });

  const text = Buffer.concat(chunks.map((chunk) => (Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)))).toString("utf8");
  return text ? JSON.parse(text) : {};
}
