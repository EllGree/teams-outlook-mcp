// Real Teams @-mentions in sent messages.
//
// Text written as "@Martin Nedbal" is just characters: nobody is notified. A real mention is an
// <at id="N">Name</at> tag in the HTML body plus an entry in the message's `mentions` array that
// carries the person's directory (AAD) object id. The sender writes @[Display Name] in the text;
// this module turns each into a tag and resolves the ids.
//
// The ids come from what the token can already read, so no extra Graph permission is needed:
// first the members of the chat being written to, then the members of every chat the user is in
// (Chat.Read), then the user's relevant people (People.Read.All). A name that resolves to nobody,
// or to more than one person, refuses the send: a message that silently fails to tag someone is
// exactly the failure this exists to prevent, and a send cannot be taken back.

import { graphGet } from "./graph.js";

const MENTION = /@\[([^\]\n]+)\]/g;

/** The distinct names written as @[Name], in order of first appearance. */
export function mentionNames(text) {
  const names = [];
  for (const m of String(text ?? "").matchAll(MENTION)) {
    const name = m[1].trim();
    if (name && !names.includes(name)) names.push(name);
  }
  return names;
}

// Diacritics and case are not something a sender should have to get exactly right: "Jakubow"
// means Peter Jakubów.
export const normalizeName = (s) =>
  String(s ?? "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

// A private-use character cannot occur in ordinary text and passes untouched through escaping,
// linkifying and paragraph splitting, so it can hold a mention's place while the text becomes HTML.
const PLACEHOLDER = (i) => `${i}`;

/**
 * Swap each @[Name] for a placeholder before the text is turned into HTML, and return a function
 * that swaps the placeholders for <at> tags afterwards. Each occurrence gets its own id, as Teams
 * requires, even when the same person is mentioned twice.
 */
export function withMentionPlaceholders(text, people) {
  const occurrences = [];
  const replaced = String(text ?? "").replace(MENTION, (_, raw) => {
    const person = people.get(raw.trim());
    occurrences.push(person);
    return PLACEHOLDER(occurrences.length - 1);
  });

  const escape = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const fill = (html) =>
    html.replace(/(\d+)/g, (_, i) => `<at id="${i}">${escape(occurrences[Number(i)].displayName)}</at>`);

  const mentions = occurrences.map((p, i) => ({
    id: i,
    mentionText: p.displayName,
    mentioned: { user: { id: p.id, displayName: p.displayName, userIdentityType: "aadUser" } },
  }));

  return { text: replaced, fill, mentions };
}

/** Pick the one candidate whose name matches; refuse on none or several distinct people. */
export function pickPerson(name, candidates) {
  const want = normalizeName(name);
  const byId = new Map();
  for (const c of candidates) {
    if (c?.id && normalizeName(c.displayName) === want) byId.set(c.id, c);
  }
  if (byId.size === 1) return [...byId.values()][0];
  if (byId.size > 1) throw new Error(`@[${name}] matches ${byId.size} different people; cannot tell which to tag`);
  return null;
}

const membersOf = (chat) =>
  (chat.members || []).filter((m) => m.userId).map((m) => ({ id: m.userId, displayName: m.displayName }));

/**
 * Resolve every @[Name] in the text to { id, displayName }. Returns a Map keyed by the name as
 * written. Throws, naming the culprits, if any name cannot be resolved to exactly one person.
 */
export async function resolveMentions(token, names, { chatId } = {}) {
  const people = new Map();
  if (!names.length) return people;

  const sources = [];
  if (chatId) {
    sources.push(async () => membersOf(await graphGet(token, `/chats/${encodeURIComponent(chatId)}?$expand=members`)));
  }
  sources.push(async () => {
    const r = await graphGet(token, `/me/chats?$top=50&$expand=members`);
    return (r.value || []).flatMap(membersOf);
  });

  let pending = [...names];
  for (const load of sources) {
    if (!pending.length) break;
    const candidates = await load();
    pending = pending.filter((name) => {
      const hit = pickPerson(name, candidates);
      if (hit) people.set(name, hit);
      return !hit;
    });
  }

  // The people search ranks colleagues the user works with; an org user's person id is their
  // directory object id, which is what a mention needs.
  for (const name of [...pending]) {
    const r = await graphGet(token, `/me/people?$search=${encodeURIComponent(`"${name}"`)}&$top=10`);
    const candidates = (r.value || [])
      .filter((p) => p.personType?.class === "Person" && p.personType?.subclass === "OrganizationUser")
      .map((p) => ({ id: p.id, displayName: p.displayName }));
    const hit = pickPerson(name, candidates);
    if (hit) {
      people.set(name, hit);
      pending = pending.filter((n) => n !== name);
    }
  }

  if (pending.length) {
    throw new Error(
      `Could not find ${pending.map((n) => `@[${n}]`).join(", ")} to tag. Write the display name exactly as Teams shows it; nothing was sent.`
    );
  }
  return people;
}
