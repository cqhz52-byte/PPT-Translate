import { json } from "../_auth.js";

const LLAMAPARSE_BASE = "https://api.cloud.llamaindex.ai";
const LLAMAPARSE_KV_KEY = "app:llamaparse_api_key";
const MAX_PDF_BYTES = 22 * 1024 * 1024;
const POLL_INTERVAL_MS = 1500;
const MAX_POLLS = 80;

export async function onRequestOptions() {
  return new Response(null, { status: 204 });
}

export async function onRequestPost({ request, env }) {
  try {
    const apiKey = String(env.LLAMAPARSE_API_KEY || env.LLAMA_CLOUD_API_KEY || "").trim() || await readStoredLlamaParseKey(env);
    if (!apiKey) {
      return json({ error: "Cloudflare 后端未配置 LLAMAPARSE_API_KEY。" }, 501);
    }

    const form = await request.formData();
    const file = form.get("file");
    if (!file || typeof file.arrayBuffer !== "function") {
      return json({ error: "请上传 PDF 文件。" }, 400);
    }

    if (file.size > MAX_PDF_BYTES) {
      return json({ error: "PDF 文件过大，请拆分后再上传解析。" }, 413);
    }

    const uploaded = await uploadLlamaParseFile(apiKey, file);
    const fileId = uploaded.id || uploaded.file_id || uploaded.fileId;
    if (!fileId) {
      return json({ error: "LlamaParse upload did not return file_id." }, 502);
    }

    const parseResponse = await fetch(`${LLAMAPARSE_BASE}/api/v2/parse`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({ file_id: fileId }),
    });

    if (!parseResponse.ok) {
      const detail = await parseResponse.text();
      return json({ error: `LlamaParse 返回 ${parseResponse.status}：${detail.slice(0, 300)}` }, parseResponse.status);
    }

    const parseData = await parseResponse.json();
    const jobId = parseData.id || parseData.job_id || parseData.jobId;
    if (!jobId) {
      return json({ error: "LlamaParse 没有返回解析任务 ID。" }, 502);
    }

    const status = await waitForParseResult(apiKey, jobId);
    if (!status.ok) return json(status.payload, status.status);

    return json({
      ok: true,
      source: "llamaparse",
      jobId,
      fileId,
      markdown: normalizeMarkdown(status.payload),
      pages: normalizeLlamaPages(status.payload),
      raw: {
        status: status.payload,
      },
    });
  } catch (error) {
    return json({ error: error.message || "PDF 后端解析失败。" }, 500);
  }
}

async function uploadLlamaParseFile(apiKey, file) {
  const bytes = await file.arrayBuffer();
  const upload = new FormData();
  upload.append("upload_file", new Blob([bytes], { type: file.type || "application/pdf" }), file.name || "document.pdf");

  const response = await fetch(`${LLAMAPARSE_BASE}/api/v1/beta/files`, {
    method: "POST",
    headers: { Authorization: `Bearer ${apiKey}` },
    body: upload,
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`LlamaParse file upload failed ${response.status}: ${detail.slice(0, 300)}`);
  }

  return response.json();
}

async function waitForParseResult(apiKey, jobId) {
  for (let attempt = 0; attempt < MAX_POLLS; attempt += 1) {
    const response = await fetch(`${LLAMAPARSE_BASE}/api/v2/parse/${encodeURIComponent(jobId)}?expand=markdown&expand=items`, {
      headers: { Authorization: `Bearer ${apiKey}` },
    });

    if (!response.ok) {
      const detail = await response.text();
      return { ok: false, status: response.status, payload: { error: `LlamaParse 状态查询失败 ${response.status}：${detail.slice(0, 300)}` } };
    }

    const payload = await response.json();
    const status = String(payload.status || "").toUpperCase();
    if (["SUCCESS", "COMPLETED", "DONE"].includes(status)) {
      return { ok: true, status: 200, payload };
    }
    if (["ERROR", "FAILED", "CANCELED", "CANCELLED"].includes(status)) {
      return { ok: false, status: 502, payload: { error: payload.error || payload.message || "LlamaParse 解析失败。" } };
    }

    await sleep(POLL_INTERVAL_MS);
  }

  return { ok: false, status: 504, payload: { error: "LlamaParse 解析超时，请稍后重试或使用兼容解析。" } };
}

function normalizeMarkdown(value) {
  if (!value) return "";
  if (typeof value === "string") return value.trim();
  if (Array.isArray(value.pages)) {
    return value.pages.map((page) => page.markdown || page.md || page.text || "").filter(Boolean).join("\n\n").trim();
  }
  return String(value.markdown || value.md || value.text || value.content || "").trim();
}

function normalizeLlamaPages(value) {
  if (!value) return [];
  if (Array.isArray(value.pages)) return value.pages;
  if (Array.isArray(value.items)) return value.items;
  if (Array.isArray(value.documents)) return value.documents;
  if (value.json && Array.isArray(value.json.pages)) return value.json.pages;
  return [];
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function readStoredLlamaParseKey(env) {
  if (!env.PHONE_AUTH_KV) return "";
  return String(await env.PHONE_AUTH_KV.get(LLAMAPARSE_KV_KEY) || "").trim();
}
