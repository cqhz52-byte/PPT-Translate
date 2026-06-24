import { hasSuperAdmin, json, requireAdminSession } from "../../_auth.js";

export async function onRequestGet({ request, env }) {
  const session = await requireAdminSession(request, env);
  return json({
    authenticated: Boolean(session),
    phone: session?.phone || "",
    role: session?.role || "",
    setupRequired: env.PHONE_AUTH_KV ? !(await hasSuperAdmin(env)) : false,
    kvReady: Boolean(env.PHONE_AUTH_KV),
  });
}
