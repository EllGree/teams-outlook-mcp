import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));

/**
 * A value copied out of config.example.json and never filled in is not a value. Without this the
 * placeholder wins over the environment, because `config.x || process.env.X` sees a truthy string,
 * and the failure arrives much later and much less clearly.
 */
export function real(value) {
  if (typeof value !== "string") return "";
  const v = value.trim();
  if (!v) return "";
  if (v.startsWith("<") && v.endsWith(">")) return "";
  return v;
}

let file = {};
try {
  file = JSON.parse(readFileSync(join(here, "..", "config.json"), "utf-8").replace(/^﻿/, ""));
} catch {
  file = {};
}

function setting(key, envName, fallback = "") {
  return real(file[key]) || real(process.env[envName]) || fallback;
}

/** The Entra tenant and the app registration this server signs in against. Both are required. */
export const TENANT_ID = setting("tenantId", "TEAMS_OUTLOOK_MCP_TENANT_ID");
export const CLIENT_ID = setting("clientId", "TEAMS_OUTLOOK_MCP_CLIENT_ID");

/**
 * Appended to every message this server sends. A person reading a chat should be able to tell that
 * a program wrote it, so this has a default rather than being allowed to be empty.
 */
export const SIGNATURE = setting("signature", "TEAMS_OUTLOOK_MCP_SIGNATURE", "🤖 sent by an assistant");

export const TOKEN_PATH = setting("tokenPath", "TEAMS_OUTLOOK_MCP_TOKENS");

export const READ_ONLY = file.readOnly === true || /^(1|true|yes)$/i.test(process.env.TEAMS_OUTLOOK_MCP_READONLY || "");

export function requireIdentity() {
  if (!TENANT_ID) {
    throw new Error("Missing tenant id: set `tenantId` in config.json or TEAMS_OUTLOOK_MCP_TENANT_ID");
  }
  if (!CLIENT_ID) {
    throw new Error("Missing client id: set `clientId` in config.json or TEAMS_OUTLOOK_MCP_CLIENT_ID");
  }
}
