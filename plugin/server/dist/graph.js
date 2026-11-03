// Minimal Microsoft Graph client for editing SharePoint Excel files. Zero dependencies.
// Auth: OAuth device code flow with the Microsoft Graph PowerShell public client; the refresh token is kept in
// $GRAPH_EXCEL_HOME/token.json (default ~/.graph-excel, chmod 600) so it survives extension/plugin updates.
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"; // Microsoft Graph PowerShell
const TENANT = "organizations";
const SCOPE = "https://graph.microsoft.com/Files.ReadWrite.All offline_access openid profile";
const AUTH = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0`;
// Hosts such as Claude Desktop may pass an unfilled template ("${user_config.x}") or "" — treat as unset.
const usable = (v) => (v && !v.startsWith("${") ? v : undefined);
export const HOME = usable(process.env.GRAPH_EXCEL_HOME) ?? path.join(os.homedir(), ".graph-excel");
const TOKEN_FILE = path.join(HOME, "token.json");
function loadToken() {
    try {
        return JSON.parse(fs.readFileSync(TOKEN_FILE, "utf8"));
    }
    catch {
        return null;
    }
}
function saveToken(res, previous = {}) {
    let account = previous.account;
    if (res.id_token) {
        try {
            const claims = JSON.parse(Buffer.from(res.id_token.split(".")[1], "base64url").toString());
            account = claims.preferred_username || claims.upn || account;
        }
        catch {
            // keep the previous account name
        }
    }
    const refresh_token = res.refresh_token ?? previous.refresh_token;
    if (!refresh_token)
        throw new Error("Microsoft did not return a refresh token");
    const t = {
        account,
        access_token: res.access_token,
        refresh_token,
        expires_at: Date.now() + (Number(res.expires_in) - 120) * 1000,
    };
    fs.mkdirSync(HOME, { recursive: true, mode: 0o700 });
    fs.writeFileSync(TOKEN_FILE, JSON.stringify(t, null, 2), { mode: 0o600 });
    fs.chmodSync(TOKEN_FILE, 0o600);
    return t;
}
async function post(url, form) {
    const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams(form),
    });
    return { ok: res.ok, body: await res.json().catch(() => ({})) };
}
async function refresh(t) {
    const r = await post(`${AUTH}/token`, { client_id: CLIENT_ID, grant_type: "refresh_token", refresh_token: t.refresh_token, scope: SCOPE });
    return r.ok ? saveToken(r.body, t) : null;
}
let refreshing = null;
export async function silentToken() {
    let t = loadToken();
    if (!t?.refresh_token)
        return null;
    if (!t.access_token || Date.now() >= t.expires_at) {
        const current = t;
        refreshing ??= refresh(current).finally(() => (refreshing = null));
        t = await refreshing;
    }
    return t && { token: t.access_token, account: t.account, expiresOn: new Date(t.expires_at).toISOString() };
}
export async function getToken() {
    const t = await silentToken();
    if (!t)
        throw new Error("Not signed in to Microsoft Graph: call login_status (or run `graph-excel-mcp login`)");
    return t.token;
}
// Starts a device code login. Resolves with the sign-in message right away; `done` settles when the user
// finishes signing in (token saved) or the code expires.
let pendingLogin = null;
export async function startLogin() {
    if (pendingLogin)
        return pendingLogin;
    const r = await post(`${AUTH}/devicecode`, { client_id: CLIENT_ID, scope: SCOPE });
    if (!r.ok)
        throw new Error(`Device code request failed: ${r.body.error_description || JSON.stringify(r.body)}`);
    const { device_code, message, interval = 5, expires_in = 900 } = r.body;
    const done = (async () => {
        const deadline = Date.now() + expires_in * 1000;
        let wait = interval * 1000;
        while (Date.now() < deadline) {
            await new Promise((s) => setTimeout(s, wait));
            const p = await post(`${AUTH}/token`, { client_id: CLIENT_ID, grant_type: "urn:ietf:params:oauth:grant-type:device_code", device_code });
            if (p.ok)
                return saveToken(p.body);
            if (p.body.error === "slow_down")
                wait += 5000;
            else if (p.body.error !== "authorization_pending")
                throw new Error(p.body.error_description || p.body.error);
        }
        throw new Error("Device code expired before sign-in finished");
    })().finally(() => (pendingLogin = null));
    done.catch(() => { });
    pendingLogin = { message, done };
    return pendingLogin;
}
// Terminal sign-in (`graph-excel-mcp login`).
export async function cliLogin() {
    const t = await silentToken();
    if (t)
        return void console.log(`SIGNED_IN ${t.account}`);
    const { message, done } = await startLogin();
    console.log(message);
    const saved = await done;
    console.log(`SIGNED_IN ${saved.account}`);
}
const RETRYABLE = new Set([429, 502, 503, 504]);
export class GraphError extends Error {
    status;
    constructor(method, url, status, text) {
        super(`${method} ${url} -> ${status}: ${text.slice(0, 500)}`);
        this.status = status;
    }
}
export async function graph(method, url, body, { session, retries = method === "GET" ? 3 : 0 } = {}) {
    const token = await getToken();
    const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json" };
    if (session)
        headers["workbook-session-id"] = session;
    for (let attempt = 0;; attempt++) {
        const res = await fetch(url.startsWith("http") ? url : `https://graph.microsoft.com/v1.0${url}`, {
            method,
            headers,
            body: body === undefined ? undefined : JSON.stringify(body),
        });
        const text = await res.text();
        if (res.ok)
            return text ? JSON.parse(text) : null;
        if (attempt < retries && RETRYABLE.has(res.status)) {
            const wait = Number(res.headers.get("retry-after")) * 1000 || 1000 * 2 ** attempt;
            await new Promise((r) => setTimeout(r, wait));
            continue;
        }
        throw new GraphError(method, url, res.status, text);
    }
}
// Sharing / web URL -> driveItem (driveId + id)
export async function resolveUrl(webUrl) {
    const enc = "u!" + Buffer.from(webUrl).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    return graph("GET", `/shares/${enc}/driveItem`);
}
export const wb = (item) => `/drives/${item.parentReference.driveId}/items/${item.id}/workbook`;
