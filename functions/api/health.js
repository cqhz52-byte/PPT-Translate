import { json } from "../_auth.js";

export function onRequestGet() {
  return json({
    ok: true,
    runtime: "cloudflare-pages-functions",
  });
}
