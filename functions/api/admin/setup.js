import { createAdminSessionCookie, createSuperAdmin, hasSuperAdmin, json } from "../../_auth.js";

export async function onRequestPost({ request, env }) {
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);
  if (await hasSuperAdmin(env)) return json({ error: "超级用户已经存在，请直接登录。" }, 409);

  const body = await request.json().catch(() => ({}));
  if (String(body.password || "") !== String(body.passwordConfirm || "")) {
    return json({ error: "两次输入的密码不一致。" }, 400);
  }
  const result = await createSuperAdmin(body.phone, body.password, env);
  if (!result.ok) return json({ error: result.error || "创建失败。" }, 400);

  return json(
    { ok: true, phone: result.phone, role: "admin" },
    200,
    { "Set-Cookie": await createAdminSessionCookie(result.phone, env) }
  );
}

export function onRequestOptions() {
  return new Response(null, { status: 204 });
}
