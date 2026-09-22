import { graphGet } from "./graph.js";

/**
 * Chats, with the one thing a caller actually needs to tell them apart: a name. A one-to-one chat
 * has no topic and is known by the other person; a group chat is known by its topic, and Graph will
 * happily hand back several with none, so those are named by their members instead.
 */
export function describeChat(chat, meId) {
  const members = (chat.members || [])
    .filter((m) => !meId || m.userId !== meId)
    .map((m) => m.displayName)
    .filter(Boolean);

  const name = chat.topic || (chat.chatType === "oneOnOne" ? members[0] : members.join(", ")) || "(no name)";

  return {
    id: chat.id,
    type: chat.chatType,
    name,
    topic: chat.topic || undefined,
    members,
    last_updated: chat.lastUpdatedDateTime,
  };
}

export async function listChats(accessToken, { top = 50, meId } = {}) {
  const r = await graphGet(accessToken, `/me/chats?$top=${top}&$expand=members`);
  return (r.value || []).map((c) => describeChat(c, meId));
}

/**
 * Resolve a chat from a name fragment. Ambiguity refuses rather than guesses: picking the wrong
 * chat means sending to the wrong people, and a send cannot be taken back. A one-to-one chat wins
 * over a group only when the fragment matches both exactly as well, since "Karel" almost always
 * means his DM.
 */
export function resolveChat(chats, fragment) {
  const frag = String(fragment ?? "").trim().toLowerCase();
  if (!frag) throw new Error("Empty chat name");

  const hits = chats.filter(
    (c) =>
      String(c.name).toLowerCase().includes(frag) ||
      String(c.topic ?? "").toLowerCase().includes(frag) ||
      c.members.some((m) => String(m).toLowerCase().includes(frag))
  );

  if (hits.length === 0) throw new Error(`No chat matching "${fragment}"`);
  if (hits.length === 1) return hits[0];

  const oneOnOne = hits.filter((c) => c.type === "oneOnOne");
  if (oneOnOne.length === 1) return oneOnOne[0];

  const listed = hits.map((c) => `${c.name} (${c.type}, ${c.id})`).join("; ");
  throw new Error(`"${fragment}" matches ${hits.length} chats, so I will not guess. Pass chat_id. Matches: ${listed}`);
}

/**
 * The hosted-content URLs inside a Teams message body. An inline image is not an attachment: it
 * lives in the HTML as an absolute Graph URL that answers with the bytes for an ordinary bearer
 * request, and `message.attachments` stays empty.
 */
export function hostedImageUrls(message) {
  // Takes a message, or a body string for a caller that only has one.
  const msg = typeof message === "string" ? { body: { content: message } } : (message ?? {});

  // An image is not always in the body. A forwarded or quoted message arrives as an entry
  // in `attachments` whose `content` is a JSON string carrying the original HTML, and any
  // image it held is in there rather than in the body the sender typed.
  const haystacks = [msg.body?.content ?? ""];
  for (const att of msg.attachments ?? []) {
    if (typeof att.content !== "string") continue;
    haystacks.push(att.content);
    try {
      const parsed = JSON.parse(att.content);
      if (typeof parsed?.originalMessageContent === "string") haystacks.push(parsed.originalMessageContent);
      if (typeof parsed?.messagePreview === "string") haystacks.push(parsed.messagePreview);
    } catch {
      /* not JSON — the raw string is already in the list */
    }
  }

  const urls = [];
  for (const h of haystacks) {
    // Inside nested JSON the quotes arrive escaped, so src=\"...\" has to match too.
    for (const m of String(h).matchAll(/src=\\?"([^"\\]*hostedContents[^"\\]*)\\?"/gi)) {
      const u = m[1].replace(/&amp;/g, "&");
      if (!urls.includes(u)) urls.push(u);
    }
  }
  return urls;
}
