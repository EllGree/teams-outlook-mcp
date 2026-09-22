// Markdown-ish to the HTML subset Teams renders.
//
// Blank lines separate paragraphs, single newlines become <br/>, bare URLs are linked because Teams
// does not do it itself, ```fences``` become code blocks and `backticks` inline code. Nothing inside
// a fence is linkified or split into paragraphs.
//
// A code block is <codeblock><code>, not <pre>. That is the element Teams itself stores, read back
// from a message formatted by hand in the Teams composer. Inside it newlines are <br> and every
// space is &nbsp;, which is how Teams keeps indentation from collapsing.

import { SIGNATURE } from "./config.js";

const escapeHtml = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// Trailing punctuation stays outside the link.
const linkify = (s) =>
  s.replace(/(https?:\/\/[^\s<]+)/g, (match) => {
    const trailMatch = match.match(/[.,;:!?)\]]+$/);
    const trail = trailMatch ? trailMatch[0] : "";
    const url = trail ? match.slice(0, -trail.length) : match;
    return `<a href="${url}">${url}</a>${trail}`;
  });

// Runs after escaping, so whatever sits between the backticks is already safe.
const inlineCode = (s) => s.replace(/`([^`\n]+)`/g, "<code>$1</code>");

const prose = (chunk) =>
  chunk
    .split(/\n\n+/)
    .map((p) => p.trim())
    .filter((p) => p.length > 0)
    .map((p) => "<p>" + inlineCode(linkify(escapeHtml(p))).replace(/\n/g, "<br/>") + "</p>")
    .join("");

export function textToHtml(text) {
  const source = String(text ?? "").replace(/\r\n/g, "\n");
  // An opening fence may carry a language label; the closing one stands alone.
  const fence = /^```[^\n]*\n([\s\S]*?)^```[ \t]*$/gm;

  let out = "";
  let cursor = 0;

  for (const match of source.matchAll(fence)) {
    out += prose(source.slice(cursor, match.index));
    const code = escapeHtml(match[1].replace(/\n$/, "")).replace(/ /g, "&nbsp;").replace(/\n/g, "<br>");
    out += `<codeblock><code>${code}</code></codeblock>`;
    cursor = match.index + match[0].length;
  }

  return out + prose(source.slice(cursor));
}

/** Every automated send carries this, so a reader can tell at a glance that a person did not type it. */
export const SIGNATURE_TEXT = SIGNATURE;
export const SIGNATURE_HTML = `<p>${escapeHtml(SIGNATURE)}</p>`;

export const appendSignatureText = (body) => `${String(body ?? "").replace(/\s+$/, "")}\n\n${SIGNATURE_TEXT}\n`;
export const appendSignatureHtml = (html) => `${String(html ?? "").replace(/\s+$/, "")}${SIGNATURE_HTML}`;

/** Readable plain text out of a Graph HTML body, for listings where the markup is noise. */
export function htmlToText(html) {
  return String(html ?? "")
    // A mention is an <at> tag wrapping the name. Strip the tag and the @ goes with it,
    // so "Any progress @Peter?" reads as someone merely saying his name.
    .replace(/<at[^>]*>([^<]*)<\/at>/gi, "@$1")
    // A link keeps its text and loses its target otherwise, so a message that is nothing
    // but a SharePoint link comes back as the file's name with no way to reach it.
    .replace(/<a\b[^>]*\bhref="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, (_, href, text) => {
      const label = String(text).replace(/<[^>]+>/g, "").trim();
      return !label || label === href ? href : `${label} (${href})`;
    })
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
