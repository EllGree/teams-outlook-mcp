import { test } from "node:test";
import assert from "node:assert/strict";
import { mentionNames, normalizeName, pickPerson, withMentionPlaceholders } from "./mentions.js";
import { textToHtml } from "./html.js";

test("mentionNames finds each @[Name] once, in order", () => {
  assert.deepEqual(mentionNames("cc @[Martin Nedbal], @[Peter Jakubów] and @[Martin Nedbal] again"), [
    "Martin Nedbal",
    "Peter Jakubów",
  ]);
  assert.deepEqual(mentionNames("no tags, @Plain or an email a@b.c"), []);
});

test("normalizeName ignores case, accents and spacing", () => {
  assert.equal(normalizeName("  Peter   Jakubów "), normalizeName("peter jakubow"));
});

test("pickPerson refuses two different people with the same name and finds none for a stranger", () => {
  const a = { id: "1", displayName: "Jan Novák" };
  const b = { id: "2", displayName: "Jan Novak" };
  assert.throws(() => pickPerson("Jan Novák", [a, b]), /2 different people/);
  assert.equal(pickPerson("Jan Novák", [a, { ...a }]).id, "1");
  assert.equal(pickPerson("Nobody", [a]), null);
});

test("mentions survive the text-to-HTML conversion as <at> tags with matching ids", () => {
  const people = new Map([
    ["Martin Nedbal", { id: "aad-m", displayName: "Martin Nedbal" }],
    ["Peter Jakubow", { id: "aad-p", displayName: "Peter Jakubów" }],
  ]);
  const { text, fill, mentions } = withMentionPlaceholders(
    "Registry is down <again>.\n\ncc @[Martin Nedbal] @[Peter Jakubow], see https://x.test/a @[Martin Nedbal]",
    people
  );
  const html = fill(textToHtml(text));

  assert.match(html, /<at id="0">Martin Nedbal<\/at> <at id="1">Peter Jakubów<\/at>/);
  assert.match(html, /<at id="2">Martin Nedbal<\/at>/);
  assert.match(html, /&lt;again&gt;/);
  assert.match(html, /<a href="https:\/\/x.test\/a">/);
  assert.equal(mentions.length, 3);
  assert.deepEqual(mentions[1], {
    id: 1,
    mentionText: "Peter Jakubów",
    mentioned: { user: { id: "aad-p", displayName: "Peter Jakubów", userIdentityType: "aadUser" } },
  });
});
