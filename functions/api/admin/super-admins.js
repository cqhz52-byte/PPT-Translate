import {
  ADMIN_PREFIX,
  MAX_SUPER_ADMINS,
  hashSecret,
  json,
  listSuperAdmins,
  normalizePhone,
  requireAdminSession,
  requireAdminToken,
} from "../../_auth.js";

export async function onRequestGet({ request, env }) {
  if (!(await isAdmin(request, env))) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  return json({
    maxAdmins: MAX_SUPER_ADMINS,
    admins: await listSuperAdmins(env),
  });
}

export async function onRequestPost({ request, env }) {
  if (!(await isAdmin(request, env))) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const body = await request.json().catch(() => ({}));
  const phone = normalizePhone(body.phone);
  if (!phone) return json({ error: "Invalid phone." }, 400);

  const existing = await readAdminRecord(env, phone);
  const admins = await listSuperAdmins(env);
  const isNewAdmin = !existing.exists;
  if (isNewAdmin && admins.length >= MAX_SUPER_ADMINS) {
    return json({ error: `最多只能设置 ${MAX_SUPER_ADMINS} 个超级管理员。` }, 400);
  }

  const password = String(body.password || "").trim();
  const passwordConfirm = String(body.passwordConfirm || "").trim();
  if (password !== passwordConfirm) return json({ error: "两次输入的密码不一致。" }, 400);
  if (isNewAdmin && password.length < 6) return json({ error: "新超级管理员密码至少 6 位。" }, 400);
  if (password && password.length < 6) return json({ error: "超级管理员密码至少 6 位。" }, 400);

  const next = {
    enabled: body.enabled !== false,
    passwordHash: password ? await hashSecret(password) : existing.record.passwordHash,
    createdAt: existing.record.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  await env.PHONE_AUTH_KV.put(`${ADMIN_PREFIX}${phone}`, JSON.stringify(next));
  return json({
    ok: true,
    admin: {
      phone,
      enabled: next.enabled,
      createdAt: next.createdAt,
      updatedAt: next.updatedAt,
    },
  });
}

export async function onRequestDelete({ request, env }) {
  const session = await requireAdminSession(request, env);
  if (!(session || requireAdminToken(request, env))) return json({ error: "Unauthorized" }, 401);
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);

  const url = new URL(request.url);
  const phone = normalizePhone(url.searchParams.get("phone"));
  if (!phone) return json({ error: "Invalid phone." }, 400);
  if (session?.phone === phone) return json({ error: "不能删除当前登录的超级管理员。" }, 400);

  const admins = await listSuperAdmins(env);
  if (admins.length <= 1) return json({ error: "至少需要保留一个超级管理员。" }, 400);

  await env.PHONE_AUTH_KV.delete(`${ADMIN_PREFIX}${phone}`);
  return json({ ok: true, phone });
}

async function isAdmin(request, env) {
  return Boolean((await requireAdminSession(request, env)) || requireAdminToken(request, env));
}

async function readAdminRecord(env, phone) {
  const value = await env.PHONE_AUTH_KV.get(`${ADMIN_PREFIX}${phone}`);
  if (!value) return { exists: false, record: {} };
  try {
    return { exists: true, record: JSON.parse(value) };
  } catch {
    return { exists: true, record: { passwordHash: value, enabled: true } };
  }
}
