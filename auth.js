#!/usr/bin/env node
// Interactive sign-in, run by hand: node auth.js
//
// Device code flow, because the server itself speaks stdio and has nowhere to put a browser. A
// scope consented after the cache was written never arrives through a refresh, so this is also the
// way to pick up a permission an admin has just added.

import { AUTHORITY, CLIENT_ID, SCOPES, TOKEN_FILE, saveTokens } from "./lib/graph.js";

// Scopes needing an admin's consent block the whole sign-in, not just themselves, so a run has to be
// able to leave one out and come back for it: node auth.js --without Channel.ReadBasic.All
const without = new Set();
for (let i = 2; i < process.argv.length; i++) {
  if (process.argv[i] === "--without" && process.argv[i + 1]) without.add(process.argv[++i]);
}

const requested = SCOPES.filter((s) => !without.has(s));
const scope = requested.join(" ");
if (without.size) console.log(`\n  Leaving out: ${[...without].join(", ")}`);

async function requestDeviceCode() {
  const res = await fetch(`${AUTHORITY}/oauth2/v2.0/devicecode`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CLIENT_ID, scope }),
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Device code request failed: ${res.status} ${text.slice(0, 400)}`);
  return JSON.parse(text);
}

async function poll(deviceCode, intervalSeconds, expiresInSeconds) {
  const deadline = Date.now() + expiresInSeconds * 1000;
  let waitMs = intervalSeconds * 1000;

  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, waitMs));

    const res = await fetch(`${AUTHORITY}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: CLIENT_ID,
        device_code: deviceCode,
      }),
    });
    const payload = JSON.parse(await res.text());

    if (res.ok) return payload;

    if (payload.error === "authorization_pending") continue;
    if (payload.error === "slow_down") {
      waitMs += 5000;
      continue;
    }
    throw new Error(`${payload.error}: ${payload.error_description || ""}`.slice(0, 500));
  }

  throw new Error("Device code expired before the sign-in completed.");
}

const dc = await requestDeviceCode();
console.log(`\n  ${dc.message}\n`);
console.log(`  URL:  ${dc.verification_uri}`);
console.log(`  Code: ${dc.user_code}\n`);
console.log("  Waiting for the sign-in to complete...");

const tokens = await poll(dc.device_code, Number(dc.interval) || 5, Number(dc.expires_in) || 900);
saveTokens(tokens);

const scp = JSON.parse(
  Buffer.from(tokens.access_token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf-8")
).scp;

console.log(`\n  Signed in. Token cached at ${TOKEN_FILE}`);
console.log(`  Scopes granted: ${scp}\n`);

const missing = requested.filter((s) => s !== "offline_access" && !String(scp).split(/\s+/).includes(s));
if (missing.length) {
  console.log(`  Not granted, so the matching tools will not be offered: ${missing.join(", ")}\n`);
}
