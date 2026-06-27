import { hashSecret, json, normalizePhone, requireAdminSession, requireAdminToken } from "../../_auth.js";

const PHONE_INDEX_KEY = "phone:index";

export async function onRequestGet({ request, env }) {
  if (!(await isAdmin(request, env))) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const phoneKeys = await listPhoneKeys(env);
  const phones = [];

  for (const phone of phoneKeys) {
    const value = await env.PHONE_AUTH_KV.get(`phone:${phone}`);
    if (!value) continue;
    let record = { enabled: true, hasPin: Boolean(value) };
    try {
      const parsed = JSON.parse(value || "{}");
      record = {
        enabled: parsed.enabled !== false,
        hasPassword: Boolean(parsed.pinHash || parsed.passwordHash || parsed.pin || parsed.code),
        updatedAt: parsed.updatedAt || "",
      };
    } catch {
      record = { enabled: true, hasPassword: Boolean(value), updatedAt: "" };
    }
    phones.push({ phone, ...record });
  }

  phones.sort((a, b) => a.phone.localeCompare(b.phone));
  return json({ phones });
}

export async function onRequestPost({ request, env }) {
  if (!(await isAdmin(request, env))) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const body = await request.json().catch(() => ({}));
  const phone = normalizePhone(body.phone);
  if (!phone) return json({ error: "Invalid phone." }, 400);

  const existing = await readPhoneRecord(env, phone);
  const next = {
    enabled: body.enabled !== false,
    createdAt: existing.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  if (Object.prototype.hasOwnProperty.call(body, "password") || Object.prototype.hasOwnProperty.call(body, "pin") || Object.prototype.hasOwnProperty.call(body, "code")) {
    const password = String(body.password || body.pin || body.code || "").trim();
    const passwordConfirm = String(body.passwordConfirm || body.pinConfirm || body.codeConfirm || "").trim();
    if (password !== passwordConfirm) return json({ error: "两次输入的密码不一致。" }, 400);
    if (password) next.pinHash = await hashSecret(password);
  } else if (existing.pinHash || existing.passwordHash || existing.pin || existing.code) {
    next.pinHash = existing.pinHash || existing.passwordHash || existing.pin || existing.code;
  }

  await env.PHONE_AUTH_KV.put(`phone:${phone}`, JSON.stringify(next));
  await addPhoneToIndex(env, phone);

  return json({
    ok: true,
    phone,
    user: {
      phone,
      enabled: body.enabled !== false,
      hasPassword: Boolean(next.pinHash),
      updatedAt: next.updatedAt,
    },
  });
}

export async function onRequestDelete({ request, env }) {
  if (!(await isAdmin(request, env))) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const url = new URL(request.url);
  const phone = normalizePhone(url.searchParams.get("phone"));
  if (!phone) return json({ error: "Invalid phone." }, 400);

  await env.PHONE_AUTH_KV.delete(`phone:${phone}`);
  await removePhoneFromIndex(env, phone);
  return json({ ok: true, phone });
}

async function isAdmin(request, env) {
  return Boolean((await requireAdminSession(request, env)) || requireAdminToken(request, env));
}

async function readPhoneRecord(env, phone) {
  const value = await env.PHONE_AUTH_KV.get(`phone:${phone}`);
  if (!value) return {};
  try {
    return JSON.parse(value);
  } catch {
    return { pin: value, enabled: true };
  }
}

async function listPhoneKeys(env) {
  const indexed = await readPhoneIndex(env);
  if (indexed.length) return indexed;

  const phones = await listPhoneKeysFromKv(env);

  if (phones.length) await writePhoneIndex(env, phones);
  return phones;
}

async function readPhoneIndex(env) {
  const value = await env.PHONE_AUTH_KV.get(PHONE_INDEX_KEY);
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    if (!Array.isArray(parsed)) return [];
    return uniqueSortedPhones(parsed.map(normalizePhone).filter(Boolean));
  } catch {
    return [];
  }
}

async function writePhoneIndex(env, phones) {
  await env.PHONE_AUTH_KV.put(PHONE_INDEX_KEY, JSON.stringify(uniqueSortedPhones(phones)));
}

async function addPhoneToIndex(env, phone) {
  let phones = await readPhoneIndex(env);
  if (!phones.length) {
    phones = await listPhoneKeysFromKv(env);
  }
  phones.push(phone);
  await writePhoneIndex(env, phones);
}

async function removePhoneFromIndex(env, phone) {
  const phones = await readPhoneIndex(env);
  await writePhoneIndex(env, phones.filter((item) => item !== phone));
}

function uniqueSortedPhones(phones) {
  return [...new Set(phones)].sort((a, b) => a.localeCompare(b));
}

async function listPhoneKeysFromKv(env) {
  const list = await env.PHONE_AUTH_KV.list({ prefix: "phone:" });
  return uniqueSortedPhones(
    list.keys
      .map((key) => normalizePhone(key.name.replace(/^phone:/, "")))
      .filter(Boolean)
  );
}
