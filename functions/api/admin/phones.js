import { json, normalizePhone, requireAdmin } from "../../_auth.js";

export async function onRequestGet({ request, env }) {
  if (!requireAdmin(request, env)) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const list = await env.PHONE_AUTH_KV.list({ prefix: "phone:" });
  const phones = [];

  for (const key of list.keys) {
    const value = await env.PHONE_AUTH_KV.get(key.name);
    const phone = key.name.replace(/^phone:/, "");
    let record = { enabled: true, hasPin: Boolean(value) };
    try {
      const parsed = JSON.parse(value || "{}");
      record = {
        enabled: parsed.enabled !== false,
        hasPin: Boolean(parsed.pin || parsed.code),
      };
    } catch {
      record = { enabled: true, hasPin: Boolean(value) };
    }
    phones.push({ phone, ...record });
  }

  return json({ phones });
}

export async function onRequestPost({ request, env }) {
  if (!requireAdmin(request, env)) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const body = await request.json().catch(() => ({}));
  const phone = normalizePhone(body.phone);
  if (!phone) return json({ error: "Invalid phone." }, 400);

  await env.PHONE_AUTH_KV.put(`phone:${phone}`, JSON.stringify({
    enabled: body.enabled !== false,
    pin: String(body.pin || body.code || "").trim(),
    updatedAt: new Date().toISOString(),
  }));

  return json({ ok: true, phone });
}

export async function onRequestDelete({ request, env }) {
  if (!requireAdmin(request, env)) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const url = new URL(request.url);
  const phone = normalizePhone(url.searchParams.get("phone"));
  if (!phone) return json({ error: "Invalid phone." }, 400);

  await env.PHONE_AUTH_KV.delete(`phone:${phone}`);
  return json({ ok: true, phone });
}
