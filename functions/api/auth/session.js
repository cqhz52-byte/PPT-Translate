import { getSession, json } from "../../_auth.js";

export async function onRequestGet({ request, env }) {
  try {
    const session = await getSession(request, env);
    return json({
      authenticated: Boolean(session),
      phone: session?.phone || "",
    });
  } catch {
    return json({ authenticated: false, phone: "" });
  }
}
