// What the token is allowed to do decides what this server offers. The access token carries the
// granted permissions in its `scp` claim, so the tool list is derived from the token rather than
// declared by hand: a scope that was never consented to produces no tool, and a scope granted later
// produces one on the next restart without a code change.

/** The `scp` claim of a JWT access token, as a Set. Never throws: a token we cannot read grants nothing. */
export function scopesFromToken(accessToken) {
  try {
    const payload = String(accessToken).split(".")[1];
    if (!payload) return new Set();
    const json = Buffer.from(payload.replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf-8");
    const scp = JSON.parse(json).scp;
    return new Set(String(scp || "").split(/\s+/).filter(Boolean));
  } catch {
    return new Set();
  }
}

/**
 * Split a tool list into what can be offered and what cannot, with the reason for each exclusion.
 * A tool that needs no scope is always available; a tool that writes is withheld in read-only mode
 * whatever the token says, because the point of that mode is to be surer than the consent screen.
 */
export function partitionTools(tools, granted, { readOnly = false } = {}) {
  const available = [];
  const withheld = [];
  for (const tool of tools) {
    if (readOnly && tool.writes) {
      withheld.push({ name: tool.name, reason: "read-only mode" });
      continue;
    }
    const missing = (tool.scopes || []).filter((s) => !granted.has(s));
    if (missing.length) withheld.push({ name: tool.name, reason: `missing scope: ${missing.join(", ")}` });
    else available.push(tool);
  }
  return { available, withheld };
}
