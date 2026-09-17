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
export function hostedImageUrls(bodyHtml) {
  const urls = [];
  const re = /<img[^>]+src="([^"]*hostedContents[^"]*)"/g;
  let m;
  while ((m = re.exec(String(bodyHtml ?? "")))) urls.push(m[1].replace(/&amp;/g, "&"));
  return urls;
}
