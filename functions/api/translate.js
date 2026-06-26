import { json } from "../_auth.js";

export async function onRequestOptions() {
  return new Response(null, { status: 204 });
}

export async function onRequestPost({ request, env }) {
  try {
    const body = await request.json().catch(() => ({}));
    const apiBase = String(body.apiBase || "https://api.deepseek.com").replace(/\/$/, "");
    const cloudflareApiKey = String(env.DEEPSEEK_API_KEY || "").trim() || await readStoredDeepSeekKey(env);
    const apiKey = String(body.apiKey || cloudflareApiKey || "").trim();
    const model = String(body.model || "deepseek-v4-flash").trim();
    const instruction = String(body.instruction || "").trim();
    const text = String(body.text || "");
    const task = String(body.task || "translate").trim();

    if (!apiKey) {
      return json({ error: "缺少 DeepSeek API Key。" }, 400);
    }

    if (!text.trim()) {
      return json({ translation: "" });
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
      return json({
        error: `DeepSeek 返回 ${response.status}：${detail.slice(0, 300)}`,
      }, response.status);
    }

    const data = await response.json();
    return json({
      translation: data.choices?.[0]?.message?.content?.trim() || "",
    });
  } catch (error) {
    return json({ error: error.message || "Server error" }, 500);
  }
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

async function readStoredDeepSeekKey(env) {
  if (!env.PHONE_AUTH_KV) return "";
  return String(await env.PHONE_AUTH_KV.get("app:deepseek_api_key") || "").trim();
}
