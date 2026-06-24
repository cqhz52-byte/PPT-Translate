import { createSessionCookie, json, verifyAuthorizedPhone } from "../../_auth.js";

export async function onRequestPost({ request, env }) {
  try {
    const body = await request.json().catch(() => ({}));
    const result = await verifyAuthorizedPhone(body.phone, body.pin, env);

    if (!result.ok) {
      return json({ error: result.error || "登录失败。" }, 401);
    }

    const cookie = await createSessionCookie(result.phone, env);
    return json(
      {
        ok: true,
        phone: result.phone,
      },
      200,
      { "Set-Cookie": cookie }
    );
  } catch (error) {
    return json({ error: error.message || "登录失败。" }, 500);
  }
}

export async function onRequestOptions() {
  return new Response(null, { status: 204 });
}
