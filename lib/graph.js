// Auth and HTTP for Microsoft Graph. Self-contained on purpose: this repo owns its own token cache
// so that the older CLI scripts in the workspace, which ask for a narrower scope list, cannot
// quietly reissue a token without the permissions this server depends on.

import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { CLIENT_ID, TENANT_ID, TOKEN_PATH, requireIdentity } from "./config.js";

export { CLIENT_ID, TENANT_ID };
export const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}`;
export const GRAPH = "https://graph.microsoft.com/v1.0";

// The scopes to ask for at sign-in. Asking for one the tenant has not consented to fails the whole
// request, so this is the list that is known to be grantable; what the token actually comes back
// with is what decides the tool list, not this constant.
export const SCOPES = [
  "User.Read",
  "Mail.Read",
  "Mail.Send",
  "Chat.Read",
  "ChatMessage.Send",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "ChannelMessage.Send",
  "offline_access",
];

export const TOKEN_FILE = resolve(
  TOKEN_PATH || resolve(process.env.LOCALAPPDATA ?? `${process.env.HOME}/AppData/Local`, "teams-outlook-mcp", "tokens.json")
);

export function loadTokens() {
  if (!existsSync(TOKEN_FILE)) {
    throw new Error(`No token cache at ${TOKEN_FILE}. Run "node auth.js" once to sign in.`);
  }
  const cache = JSON.parse(readFileSync(TOKEN_FILE, "utf-8"));
  return {
    accessToken: cache.access_token,
    refreshToken: cache.refresh_token,
    scope: cache.scope,
    expiresAt: new Date(cache.expires_at),
  };
}

export function saveTokens(tokenRes) {
  const expiresAt = new Date(Date.now() + (Number(tokenRes.expires_in) - 60) * 1000);
  mkdirSync(dirname(TOKEN_FILE), { recursive: true });
  writeFileSync(
    TOKEN_FILE,
    JSON.stringify(
      {
        access_token: tokenRes.access_token,
        refresh_token: tokenRes.refresh_token,
        scope: tokenRes.scope,
        expires_at: expiresAt.toISOString(),
      },
      null,
      2
    ),
    "utf-8"
  );
}

export async function refreshAccessToken(refreshToken) {
  requireIdentity();
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    client_id: CLIENT_ID,
    refresh_token: refreshToken,
    scope: SCOPES.join(" "),
  });
  const res = await fetch(`${AUTHORITY}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const text = await res.text();
  if (!res.ok) {
    // A scope consented after the cache was written is not in the refresh grant, and Azure answers
    // AADSTS65001 rather than issuing a narrower token. Signing in again is the only way through.
    const hint = text.includes("AADSTS65001") ? ' Run "node auth.js" to consent to the new scopes.' : "";
    throw new Error(`Token refresh failed: ${res.status} ${text.slice(0, 300)}${hint}`);
  }
  return JSON.parse(text);
}

export async function getAccessToken() {
  const t = loadTokens();
  if (t.expiresAt > new Date()) return t.accessToken;
  const refreshed = await refreshAccessToken(t.refreshToken);
  saveTokens(refreshed);
  return refreshed.access_token;
}

export async function graphGet(accessToken, urlOrPath) {
  const url = urlOrPath.startsWith("http") ? urlOrPath : `${GRAPH}${urlOrPath}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (!res.ok) throw new Error(`GET ${url} -> ${res.status} ${(await res.text()).slice(0, 400)}`);
  return res.json();
}

/** Bytes, for the endpoints that answer with a file rather than JSON. */
export async function graphGetBinary(accessToken, urlOrPath) {
  const url = urlOrPath.startsWith("http") ? urlOrPath : `${GRAPH}${urlOrPath}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (!res.ok) throw new Error(`GET ${url} -> ${res.status} ${(await res.text()).slice(0, 400)}`);
  return {
    buffer: Buffer.from(await res.arrayBuffer()),
    mediaType: res.headers.get("content-type") || "application/octet-stream",
  };
}

export async function graphPost(accessToken, path, bodyObj) {
  const res = await fetch(`${GRAPH}${path}`, {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify(bodyObj),
  });
  if (!res.ok) throw new Error(`POST ${path} -> ${res.status} ${(await res.text()).slice(0, 500)}`);
  // sendMail answers 202 with no body; chat and channel messages answer 201 with one.
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}
