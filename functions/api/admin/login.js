import { createAdminSessionCookie, hasSuperAdmin, json, verifySuperAdmin } from "../../_auth.js";

export async function onRequestPost({ request, env }) {
  if (!env.PHONE_AUTH_KV) return json({ error: "PHONE_AUTH_KV is not bound." }, 501);
  if (!(await hasSuperAdmin(env))) {
    return json({ error: "请先创建超级用户。", setupRequired: true }, 409);
  }

  const body = await request.json().catch(() => ({}));
  const result = await verifySuperAdmin(body.phone, body.password, env);
  if (!result.ok) return json({ error: result.error || "登录失败。" }, 401);

  return json(
    { ok: true, phone: result.phone, role: "admin" },
    200,
    { "Set-Cookie": await createAdminSessionCookie(result.phone, env) }
  );
}

export function onRequestOptions() {
  return new Response(null, { status: 204 });
}
