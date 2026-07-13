async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") {
    sendJson(res, 204, null);
    return;
  }

  sendJson(res, 501, {
    error: "PDF 后端结构化解析仅在 Cloudflare Pages Functions 中启用；请部署到 Cloudflare 并配置 LLAMAPARSE_API_KEY。",
  });
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
