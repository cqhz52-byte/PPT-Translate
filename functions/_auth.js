const SESSION_COOKIE = "ppt_auth";
const SESSION_MAX_AGE = 60 * 60 * 24 * 30;

export async function getSession(request, env) {
  const cookie = getCookie(request, SESSION_COOKIE);
  if (!cookie) return null;

  const [payload, signature] = cookie.split(".");
  if (!payload || !signature) return null;

  const expected = await sign(payload, getSessionSecret(env));
  if (!timingSafeEqual(signature, expected)) return null;

  try {
    const data = JSON.parse(base64UrlDecode(payload));
    if (!data.phone || !data.exp || Date.now() > data.exp) return null;
    return data;
  } catch {
    return null;
  }
}

export async function createSessionCookie(phone, env) {
  const payload = base64UrlEncode(JSON.stringify({
    phone,
    exp: Date.now() + SESSION_MAX_AGE * 1000,
  }));
  const signature = await sign(payload, getSessionSecret(env));
  return [
    `${SESSION_COOKIE}=${payload}.${signature}`,
    "Path=/",
    `Max-Age=${SESSION_MAX_AGE}`,
    "HttpOnly",
    "Secure",
    "SameSite=Lax",
  ].join("; ");
}

export function clearSessionCookie() {
  return `${SESSION_COOKIE}=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Lax`;
}

export async function verifyAuthorizedPhone(phone, pin, env) {
  const normalized = normalizePhone(phone);
  if (!normalized) return { ok: false, error: "请输入手机号。" };

  if (String(env.AUTH_ALLOW_ALL || "").toLowerCase() === "true") {
    return { ok: true, phone: normalized };
  }

  const kvRecord = await readPhoneFromKv(normalized, env);
  if (kvRecord) return verifyPhoneRecord(normalized, pin, kvRecord);

  const envRecord = readPhoneFromEnv(normalized, env);
  if (envRecord) return verifyPhoneRecord(normalized, pin, envRecord);

  return { ok: false, error: "该手机号未授权使用。" };
}

export function normalizePhone(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  const hasPlus = text.startsWith("+");
  const digits = text.replace(/\D/g, "");
  if (digits.length < 7 || digits.length > 15) return "";
  return hasPlus ? `+${digits}` : digits;
}

export function requireAdmin(request, env) {
  const expected = String(env.ADMIN_TOKEN || "").trim();
  if (!expected) return false;
  const header = request.headers.get("Authorization") || "";
  return header === `Bearer ${expected}`;
}

export function json(payload, status = 200, headers = {}) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      ...headers,
    },
  });
}

export function loginPage(message = "") {
  const safeMessage = escapeHtml(message);
  return new Response(`<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>授权登录 - CuraWay 文档翻译工具</title>
    <style>
      :root { color-scheme: light; --bg: #f6f8f7; --ink: #172321; --muted: #60706c; --line: #d7e0dc; --accent: #165a72; --danger: #b42318; }
      * { box-sizing: border-box; }
      body { margin: 0; min-height: 100vh; display: grid; place-items: center; padding: 24px; background: var(--bg); color: var(--ink); font-family: "Segoe UI", "Microsoft YaHei", system-ui, sans-serif; }
      main { width: min(420px, 100%); padding: 24px; border: 1px solid var(--line); border-radius: 10px; background: #fff; box-shadow: 0 18px 48px rgba(20, 35, 32, 0.12); }
      h1 { margin: 0 0 8px; font-size: 24px; }
      p { margin: 0 0 18px; color: var(--muted); line-height: 1.55; }
      label { display: grid; gap: 6px; margin-top: 12px; color: var(--muted); font-size: 13px; font-weight: 700; }
      input { width: 100%; height: 44px; border: 1px solid var(--line); border-radius: 8px; padding: 0 12px; font: inherit; }
      button { width: 100%; height: 44px; margin-top: 18px; border: 1px solid var(--accent); border-radius: 8px; background: var(--accent); color: #fff; font: inherit; font-weight: 800; cursor: pointer; }
      .error { display: ${safeMessage ? "block" : "none"}; margin-top: 12px; color: var(--danger); font-size: 13px; }
      .hint { margin-top: 14px; font-size: 12px; }
    </style>
  </head>
  <body>
    <main>
      <h1>授权手机号登录</h1>
      <p>请输入已授权的手机号。若管理员配置了访问码，也需要一起填写。</p>
      <form id="loginForm">
        <label>手机号 <input name="phone" inputmode="tel" autocomplete="tel" required></label>
        <label>访问码/PIN <input name="pin" type="password" autocomplete="current-password" placeholder="未配置可留空"></label>
        <button type="submit">进入应用</button>
        <div class="error" id="errorText">${safeMessage}</div>
      </form>
      <p class="hint">手机号会在服务端校验；未授权用户不能调用翻译接口。</p>
    </main>
    <script>
      document.querySelector("#loginForm").addEventListener("submit", async (event) => {
        event.preventDefault();
        const form = new FormData(event.currentTarget);
        const response = await fetch("/api/auth/login", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            phone: form.get("phone"),
            pin: form.get("pin"),
          }),
        });
        const data = await response.json().catch(() => ({}));
        if (response.ok) {
          location.href = "/";
          return;
        }
        document.querySelector("#errorText").style.display = "block";
        document.querySelector("#errorText").textContent = data.error || "登录失败。";
      });
    </script>
  </body>
</html>`, {
    status: 401,
    headers: { "Content-Type": "text/html; charset=utf-8" },
  });
}

async function readPhoneFromKv(phone, env) {
  if (!env.PHONE_AUTH_KV) return null;
  const value = await env.PHONE_AUTH_KV.get(`phone:${phone}`);
  if (!value) return null;
  try {
    return JSON.parse(value);
  } catch {
    return { pin: value, enabled: true };
  }
}

function readPhoneFromEnv(phone, env) {
  const users = String(env.AUTHORIZED_USERS || "").trim();
  if (users) {
    try {
      const parsed = JSON.parse(users);
      for (const [rawPhone, record] of Object.entries(parsed)) {
        if (normalizePhone(rawPhone) === phone) {
          return typeof record === "object" ? record : { pin: String(record || ""), enabled: true };
        }
      }
    } catch {
      for (const item of users.split(",")) {
        const [rawPhone, rawPin = ""] = item.split(":");
        if (normalizePhone(rawPhone) === phone) return { pin: rawPin.trim(), enabled: true };
      }
    }
  }

  const phones = String(env.AUTHORIZED_PHONES || "").split(",").map(normalizePhone);
  if (phones.includes(phone)) return { pin: "", enabled: true };
  return null;
}

function verifyPhoneRecord(phone, pin, record) {
  if (record.enabled === false) return { ok: false, error: "该手机号已停用。" };
  const expectedPin = String(record.pin || record.code || "").trim();
  if (expectedPin && String(pin || "").trim() !== expectedPin) {
    return { ok: false, error: "访问码不正确。" };
  }
  return { ok: true, phone };
}

function getSessionSecret(env) {
  const secret = String(env.AUTH_SESSION_SECRET || "").trim();
  if (!secret) throw new Error("Missing AUTH_SESSION_SECRET");
  return secret;
}

async function sign(payload, secret) {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const signature = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(payload));
  return base64UrlEncodeBytes(new Uint8Array(signature));
}

function getCookie(request, name) {
  const cookie = request.headers.get("Cookie") || "";
  return cookie.split(";").map((item) => item.trim()).find((item) => item.startsWith(`${name}=`))?.slice(name.length + 1) || "";
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let index = 0; index < a.length; index += 1) {
    result |= a.charCodeAt(index) ^ b.charCodeAt(index);
  }
  return result === 0;
}

function base64UrlEncode(value) {
  return base64UrlEncodeBytes(new TextEncoder().encode(value));
}

function base64UrlEncodeBytes(bytes) {
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64UrlDecode(value) {
  const padded = value.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(value.length / 4) * 4, "=");
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
