import { getSession, json, loginPage } from "./_auth.js";

const PUBLIC_PATHS = [
  "/api/auth/login",
  "/api/auth/logout",
  "/api/auth/session",
  "/api/health",
  "/favicon.ico",
  "/manifest.webmanifest",
  "/sw.js",
];

export async function onRequest(context) {
  const { request, env } = context;
  const url = new URL(request.url);

  if (request.method === "OPTIONS") {
    return context.next();
  }

  if (isPublicPath(url.pathname) || isStaticAsset(url.pathname)) {
    return context.next();
  }

  let session = null;
  try {
    session = await getSession(request, env);
  } catch (error) {
    if (isApiRequest(url.pathname)) {
      return json({ error: error.message || "Authorization unavailable" }, 500);
    }
    return loginPage("授权服务未配置，请联系管理员。");
  }

  if (session) {
    return context.next();
  }

  if (isApiRequest(url.pathname)) {
    return json({ error: "请先用授权手机号登录。" }, 401);
  }

  return loginPage();
}

function isPublicPath(pathname) {
  if (PUBLIC_PATHS.includes(pathname)) return true;
  return pathname.startsWith("/api/admin/");
}

function isApiRequest(pathname) {
  return pathname.startsWith("/api/");
}

function isStaticAsset(pathname) {
  return /\.(?:css|js|mjs|png|jpg|jpeg|svg|ico|webmanifest|woff2?|ttf|map)$/i.test(pathname);
}
