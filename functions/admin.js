import { getSession, hasSuperAdmin } from "./_auth.js";

export async function onRequest({ request, env }) {
  if (!env.PHONE_AUTH_KV) return html(kvMissingPage(), 501);

  const session = await getSession(request, env);
  if (session?.role === "admin") return html(adminPage(session));

  const setupRequired = !(await hasSuperAdmin(env));
  return html(adminLoginPage(setupRequired), 401);
}

function adminLoginPage(setupRequired) {
  const mode = setupRequired ? "setup" : "login";
  const title = setupRequired ? "创建超级用户" : "超级用户登录";
  const buttonText = setupRequired ? "创建并进入后台" : "进入管理后台";
  const note = setupRequired
    ? "这是首次初始化入口。创建成功后，其他人不能再通过这里创建超级用户。"
    : "请输入超级用户手机号和密码。";

  return `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${title} - CuraWay</title>
    <style>${baseStyles()}</style>
  </head>
  <body>
    <main class="auth-card">
      <div class="brand">CuraWay</div>
      <h1>${title}</h1>
      <p>${note}</p>
      <form id="adminAuthForm" data-mode="${mode}">
        <label>手机号 <input name="phone" inputmode="tel" autocomplete="tel" required></label>
        <label>密码 <input name="password" type="password" autocomplete="current-password" required minlength="6"></label>
        ${setupRequired ? '<label>确认密码 <input name="passwordConfirm" type="password" autocomplete="new-password" required minlength="6"></label>' : ""}
        <button type="submit">${buttonText}</button>
        <div class="error" id="errorText"></div>
      </form>
      <p class="hint"><a href="/">返回应用登录</a></p>
    </main>
    <script>
      document.querySelector("#adminAuthForm").addEventListener("submit", async (event) => {
        event.preventDefault();
        const form = event.currentTarget;
        const data = new FormData(form);
        const endpoint = form.dataset.mode === "setup" ? "/api/admin/setup" : "/api/admin/login";
        if (form.dataset.mode === "setup" && data.get("password") !== data.get("passwordConfirm")) {
          const error = document.querySelector("#errorText");
          error.textContent = "两次输入的密码不一致。";
          error.style.display = "block";
          return;
        }
        const response = await fetch(endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            phone: data.get("phone"),
            password: data.get("password"),
            passwordConfirm: data.get("passwordConfirm")
          })
        });
        const payload = await response.json().catch(() => ({}));
        if (response.ok) {
          location.href = "/admin";
          return;
        }
        const error = document.querySelector("#errorText");
        error.textContent = payload.error || "操作失败。";
        error.style.display = "block";
      });
    </script>
  </body>
</html>`;
}

function adminPage(session) {
  return `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>授权手机号管理 - CuraWay</title>
    <style>${baseStyles()}</style>
  </head>
  <body>
    <main class="shell">
      <header class="topbar">
        <div>
          <div class="brand">CuraWay</div>
          <h1>授权手机号管理</h1>
          <p>超级用户：${escapeHtml(session.phone)}</p>
        </div>
        <div class="top-actions">
          <a class="ghost" href="/">打开应用</a>
          <button class="ghost" id="logoutButton" type="button">退出</button>
        </div>
      </header>

      <section class="panel">
        <h2 id="adminFormTitle">添加超级管理员</h2>
        <p class="panel-copy">最多可设置 2 个超级管理员。第二个超级管理员必须由已登录的超级管理员添加。</p>
        <form id="adminForm" class="user-form">
          <label>手机号 <input name="phone" inputmode="tel" autocomplete="tel" required></label>
          <label>密码 <input name="password" type="password" autocomplete="new-password" placeholder="新增必填，修改可留空"></label>
          <label>确认密码 <input name="passwordConfirm" type="password" autocomplete="new-password" placeholder="再次输入新密码"></label>
          <label class="check"><input name="enabled" type="checkbox" checked> 启用这个管理员</label>
          <div class="form-actions">
            <button type="submit">保存管理员</button>
            <button class="ghost" id="adminResetButton" type="button">清空</button>
          </div>
          <div class="error" id="adminErrorText"></div>
        </form>
      </section>

      <section class="panel">
        <div class="section-head">
          <h2>超级管理员</h2>
          <button class="ghost" id="adminRefreshButton" type="button">刷新</button>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>手机号</th>
                <th>状态</th>
                <th>创建时间</th>
                <th>更新时间</th>
                <th>操作</th>
              </tr>
            </thead>
            <tbody id="adminRows">
              <tr><td colspan="5">正在加载...</td></tr>
            </tbody>
          </table>
        </div>
      </section>

      <section class="panel">
        <h2 id="formTitle">添加授权用户</h2>
        <form id="userForm" class="user-form">
          <label>手机号 <input name="phone" inputmode="tel" autocomplete="tel" required></label>
          <label>密码 <input name="password" type="password" autocomplete="new-password" placeholder="留空表示不修改密码"></label>
          <label>确认密码 <input name="passwordConfirm" type="password" autocomplete="new-password" placeholder="再次输入新密码"></label>
          <label class="check"><input name="enabled" type="checkbox" checked> 启用这个用户</label>
          <div class="form-actions">
            <button type="submit">保存授权</button>
            <button class="ghost" id="resetButton" type="button">清空</button>
          </div>
          <div class="error" id="errorText"></div>
        </form>
      </section>

      <section class="panel">
        <div class="section-head">
          <h2>已授权用户</h2>
          <button class="ghost" id="refreshButton" type="button">刷新</button>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>手机号</th>
                <th>状态</th>
                <th>密码</th>
                <th>更新时间</th>
                <th>操作</th>
              </tr>
            </thead>
            <tbody id="phoneRows">
              <tr><td colspan="5">正在加载...</td></tr>
            </tbody>
          </table>
        </div>
      </section>
    </main>

    <script>
      const form = document.querySelector("#userForm");
      const rows = document.querySelector("#phoneRows");
      const errorText = document.querySelector("#errorText");
      const formTitle = document.querySelector("#formTitle");
      const adminForm = document.querySelector("#adminForm");
      const adminRows = document.querySelector("#adminRows");
      const adminErrorText = document.querySelector("#adminErrorText");
      const adminFormTitle = document.querySelector("#adminFormTitle");

      function showError(message) {
        errorText.textContent = message || "";
        errorText.style.display = message ? "block" : "none";
      }

      function showAdminError(message) {
        adminErrorText.textContent = message || "";
        adminErrorText.style.display = message ? "block" : "none";
      }

      function resetForm() {
        form.reset();
        form.elements.enabled.checked = true;
        formTitle.textContent = "添加授权用户";
        showError("");
      }

      function resetAdminForm() {
        adminForm.reset();
        adminForm.elements.enabled.checked = true;
        adminFormTitle.textContent = "添加超级管理员";
        showAdminError("");
      }

      async function api(path, options = {}) {
        const response = await fetch(path, {
          ...options,
          headers: {
            "Content-Type": "application/json",
            ...(options.headers || {})
          }
        });
        const payload = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(payload.error || "请求失败。");
        return payload;
      }

      async function loadPhones() {
        rows.innerHTML = '<tr><td colspan="5">正在加载...</td></tr>';
        try {
          const data = await api("/api/admin/phones");
          if (!data.phones.length) {
            rows.innerHTML = '<tr><td colspan="5">还没有授权用户。</td></tr>';
            return;
          }
          rows.innerHTML = data.phones.map((item) => {
            const status = item.enabled ? "启用" : "停用";
            const password = item.hasPassword ? "已设置" : "未设置";
            const updatedAt = item.updatedAt ? new Date(item.updatedAt).toLocaleString() : "-";
            return '<tr>' +
              '<td>' + item.phone + '</td>' +
              '<td><span class="badge ' + (item.enabled ? "ok" : "off") + '">' + status + '</span></td>' +
              '<td>' + password + '</td>' +
              '<td>' + updatedAt + '</td>' +
              '<td class="actions">' +
                '<button class="ghost" type="button" data-edit="' + item.phone + '" data-enabled="' + item.enabled + '">编辑</button>' +
                '<button class="danger" type="button" data-delete="' + item.phone + '">删除</button>' +
              '</td>' +
            '</tr>';
          }).join("");
        } catch (error) {
          rows.innerHTML = '<tr><td colspan="5">' + error.message + '</td></tr>';
        }
      }

      async function loadAdmins() {
        adminRows.innerHTML = '<tr><td colspan="5">正在加载...</td></tr>';
        try {
          const data = await api("/api/admin/super-admins");
          if (!data.admins.length) {
            adminRows.innerHTML = '<tr><td colspan="5">还没有超级管理员。</td></tr>';
            return;
          }
          adminRows.innerHTML = data.admins.map((item) => {
            const status = item.enabled ? "启用" : "停用";
            const createdAt = item.createdAt ? new Date(item.createdAt).toLocaleString() : "-";
            const updatedAt = item.updatedAt ? new Date(item.updatedAt).toLocaleString() : "-";
            return '<tr>' +
              '<td>' + item.phone + '</td>' +
              '<td><span class="badge ' + (item.enabled ? "ok" : "off") + '">' + status + '</span></td>' +
              '<td>' + createdAt + '</td>' +
              '<td>' + updatedAt + '</td>' +
              '<td class="actions">' +
                '<button class="ghost" type="button" data-admin-edit="' + item.phone + '" data-enabled="' + item.enabled + '">编辑</button>' +
                '<button class="danger" type="button" data-admin-delete="' + item.phone + '">删除</button>' +
              '</td>' +
            '</tr>';
          }).join("");
        } catch (error) {
          adminRows.innerHTML = '<tr><td colspan="5">' + error.message + '</td></tr>';
        }
      }

      form.addEventListener("submit", async (event) => {
        event.preventDefault();
        const data = new FormData(form);
        const body = {
          phone: data.get("phone"),
          enabled: form.elements.enabled.checked
        };
        const password = String(data.get("password") || "").trim();
        const passwordConfirm = String(data.get("passwordConfirm") || "").trim();
        if (password !== passwordConfirm) {
          showError("两次输入的密码不一致。");
          return;
        }
        if (password) body.password = password;
        if (passwordConfirm) body.passwordConfirm = passwordConfirm;
        try {
          await api("/api/admin/phones", {
            method: "POST",
            body: JSON.stringify(body)
          });
          resetForm();
          await loadPhones();
        } catch (error) {
          showError(error.message);
        }
      });

      adminForm.addEventListener("submit", async (event) => {
        event.preventDefault();
        const data = new FormData(adminForm);
        const body = {
          phone: data.get("phone"),
          enabled: adminForm.elements.enabled.checked
        };
        const password = String(data.get("password") || "").trim();
        const passwordConfirm = String(data.get("passwordConfirm") || "").trim();
        if (password !== passwordConfirm) {
          showAdminError("两次输入的密码不一致。");
          return;
        }
        if (password) body.password = password;
        if (passwordConfirm) body.passwordConfirm = passwordConfirm;
        try {
          await api("/api/admin/super-admins", {
            method: "POST",
            body: JSON.stringify(body)
          });
          resetAdminForm();
          await loadAdmins();
        } catch (error) {
          showAdminError(error.message);
        }
      });

      rows.addEventListener("click", async (event) => {
        const editPhone = event.target.dataset.edit;
        const deletePhone = event.target.dataset.delete;
        if (editPhone) {
          form.elements.phone.value = editPhone;
          form.elements.password.value = "";
          form.elements.passwordConfirm.value = "";
          form.elements.enabled.checked = event.target.dataset.enabled === "true";
          formTitle.textContent = "修改授权用户";
          form.scrollIntoView({ behavior: "smooth", block: "start" });
        }
        if (deletePhone) {
          if (!confirm("确定删除 " + deletePhone + " 的授权吗？")) return;
          try {
            await api("/api/admin/phones?phone=" + encodeURIComponent(deletePhone), { method: "DELETE" });
            await loadPhones();
          } catch (error) {
            showError(error.message);
          }
        }
      });

      adminRows.addEventListener("click", async (event) => {
        const editPhone = event.target.dataset.adminEdit;
        const deletePhone = event.target.dataset.adminDelete;
        if (editPhone) {
          adminForm.elements.phone.value = editPhone;
          adminForm.elements.password.value = "";
          adminForm.elements.passwordConfirm.value = "";
          adminForm.elements.enabled.checked = event.target.dataset.enabled === "true";
          adminFormTitle.textContent = "修改超级管理员";
          adminForm.scrollIntoView({ behavior: "smooth", block: "start" });
        }
        if (deletePhone) {
          if (!confirm("确定删除超级管理员 " + deletePhone + " 吗？")) return;
          try {
            await api("/api/admin/super-admins?phone=" + encodeURIComponent(deletePhone), { method: "DELETE" });
            await loadAdmins();
          } catch (error) {
            showAdminError(error.message);
          }
        }
      });

      document.querySelector("#resetButton").addEventListener("click", resetForm);
      document.querySelector("#refreshButton").addEventListener("click", loadPhones);
      document.querySelector("#adminResetButton").addEventListener("click", resetAdminForm);
      document.querySelector("#adminRefreshButton").addEventListener("click", loadAdmins);
      document.querySelector("#logoutButton").addEventListener("click", async () => {
        await fetch("/api/auth/logout", { method: "POST" });
        location.href = "/admin";
      });

      loadAdmins();
      loadPhones();
    </script>
  </body>
</html>`;
}

function kvMissingPage() {
  return `<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>需要配置 KV - CuraWay</title>
    <style>${baseStyles()}</style>
  </head>
  <body>
    <main class="auth-card">
      <div class="brand">CuraWay</div>
      <h1>需要绑定 KV 存储</h1>
      <p>网页内管理手机号需要一个 Cloudflare KV 命名空间，并绑定变量名 <strong>PHONE_AUTH_KV</strong>。绑定后重新部署，再打开 /admin 创建超级用户。</p>
      <p class="hint"><a href="/">返回应用</a></p>
    </main>
  </body>
</html>`;
}

function baseStyles() {
  return `
    :root { color-scheme: light; --bg: #f6f8f7; --ink: #172321; --muted: #60706c; --line: #d7e0dc; --accent: #165a72; --accent-2: #1f7a5b; --danger: #b42318; --off: #7a4d0b; }
    * { box-sizing: border-box; }
    body { margin: 0; min-height: 100vh; background: var(--bg); color: var(--ink); font-family: "Segoe UI", "Microsoft YaHei", system-ui, sans-serif; }
    a { color: var(--accent); font-weight: 800; text-decoration: none; }
    h1, h2, p { margin-top: 0; }
    h1 { margin-bottom: 8px; font-size: clamp(24px, 5vw, 34px); }
    h2 { margin-bottom: 14px; font-size: 18px; }
    p { color: var(--muted); line-height: 1.55; }
    .panel-copy { margin-bottom: 14px; font-size: 13px; }
    button, input { font: inherit; }
    button { min-height: 42px; border: 1px solid var(--accent); border-radius: 8px; padding: 0 14px; background: var(--accent); color: #fff; font-weight: 800; cursor: pointer; }
    input { width: 100%; height: 44px; border: 1px solid var(--line); border-radius: 8px; padding: 0 12px; background: #fff; color: var(--ink); }
    label { display: grid; gap: 6px; color: var(--muted); font-size: 13px; font-weight: 800; }
    .brand { margin-bottom: 10px; color: var(--accent-2); font-size: 18px; font-weight: 900; }
    .auth-card { width: min(430px, calc(100% - 32px)); margin: 10vh auto; padding: 24px; border: 1px solid var(--line); border-radius: 10px; background: #fff; box-shadow: 0 18px 48px rgba(20, 35, 32, 0.12); }
    .auth-card form { display: grid; gap: 12px; }
    .hint { margin-top: 14px; font-size: 12px; }
    .error { display: none; margin-top: 10px; color: var(--danger); font-size: 13px; font-weight: 800; }
    .shell { width: min(1120px, calc(100% - 28px)); margin: 0 auto; padding: 24px 0 40px; }
    .topbar { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; margin-bottom: 18px; }
    .top-actions, .form-actions, .section-head, .actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .panel { margin-top: 14px; padding: 18px; border: 1px solid var(--line); border-radius: 10px; background: #fff; }
    .section-head { justify-content: space-between; margin-bottom: 10px; }
    .user-form { display: grid; grid-template-columns: minmax(160px, 1fr) minmax(160px, 1fr) minmax(160px, 1fr) auto; gap: 12px; align-items: end; }
    .check { display: flex; align-items: center; gap: 8px; height: 44px; }
    .check input { width: 18px; height: 18px; }
    .form-actions { grid-column: 1 / -1; }
    .ghost { display: inline-flex; align-items: center; justify-content: center; min-height: 42px; border-color: var(--line); background: #fff; color: var(--accent); }
    .danger { border-color: #f2c2bd; background: #fff; color: var(--danger); }
    .table-wrap { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; min-width: 720px; }
    th, td { padding: 12px 10px; border-bottom: 1px solid var(--line); text-align: left; white-space: nowrap; }
    th { color: var(--muted); font-size: 12px; }
    .badge { display: inline-flex; align-items: center; min-height: 26px; border-radius: 999px; padding: 0 10px; font-size: 12px; font-weight: 900; }
    .badge.ok { background: #e6f5ed; color: var(--accent-2); }
    .badge.off { background: #fff4df; color: var(--off); }
    @media (max-width: 720px) {
      .topbar { display: grid; }
      .top-actions { width: 100%; }
      .top-actions .ghost, .top-actions button { flex: 1; }
      .user-form { grid-template-columns: 1fr; }
      .form-actions button { flex: 1; }
      .panel { padding: 14px; }
    }
  `;
}

function html(content, status = 200) {
  return new Response(content, {
    status,
    headers: { "Content-Type": "text/html; charset=utf-8" },
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
