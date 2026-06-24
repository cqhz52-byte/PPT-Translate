import { clearSessionCookie, json } from "../../_auth.js";

export async function onRequestPost() {
  return json({ ok: true }, 200, { "Set-Cookie": clearSessionCookie() });
}

export async function onRequestGet() {
  return json({ ok: true }, 200, { "Set-Cookie": clearSessionCookie() });
}
