import { json, requireAdminSession, requireAdminToken } from "../../_auth.js";

const LLAMAPARSE_KEY = "app:llamaparse_api_key";

export async function onRequestGet({ request, env }) {
  if (!(await isAdmin(request, env))) return json({ error: "Unauthorized" }, 401);
  const hasSecret = Boolean(String(env.LLAMAPARSE_API_KEY || env.LLAMA_CLOUD_API_KEY || "").trim());
  const hasKvKey = Boolean(env.PHONE_AUTH_KV && (await env.PHONE_AUTH_KV.get(LLAMAPARSE_KEY)));
  return json({
    hasKey: hasSecret || hasKvKey,
    source: hasSecret ? "cloudflare-secret" : (hasKvKey ? "admin-kv" : ""),
  });
}

export async function onRequestPost({ request, env }) {
  if (!(await isAdmin(request, env))) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const body = await request.json().catch(() => ({}));
  const apiKey = String(body.apiKey || "").trim();
  if (!apiKey) return json({ error: "请输入 LlamaParse API Key。" }, 400);

  await env.PHONE_AUTH_KV.put(LLAMAPARSE_KEY, apiKey);
  return json({ ok: true, hasKey: true, source: "admin-kv" });
}

export async function onRequestDelete({ request, env }) {
  if (!(await isAdmin(request, env))) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);
  await env.PHONE_AUTH_KV.delete(LLAMAPARSE_KEY);
  return json({ ok: true });
}

async function isAdmin(request, env) {
  return Boolean((await requireAdminSession(request, env)) || requireAdminToken(request, env));
}
