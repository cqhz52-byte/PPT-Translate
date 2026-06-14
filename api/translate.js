module.exports = async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : req.body || {};
    const apiBase = String(body.apiBase || "https://api.deepseek.com").replace(/\/$/, "");
    const apiKey = String(body.apiKey || process.env.DEEPSEEK_API_KEY || "").trim();
    const model = String(body.model || "deepseek-v4-flash").trim();
    const instruction = String(body.instruction || "").trim();
    const text = String(body.text || "");

    if (!apiKey) {
      res.status(400).json({ error: "缺少 DeepSeek API Key。" });
      return;
    }

    if (!text.trim()) {
      res.status(200).json({ translation: "" });
      return;
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
      res.status(response.status).json({
        error: `DeepSeek 返回 ${response.status}：${detail.slice(0, 300)}`,
      });
      return;
    }

    const data = await response.json();
    res.status(200).json({
      translation: data.choices?.[0]?.message?.content?.trim() || "",
    });
  } catch (error) {
    res.status(500).json({ error: error.message || "Server error" });
  }
};

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}
