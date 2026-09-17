import { test } from "node:test";
import assert from "node:assert/strict";
import { partitionTools, scopesFromToken } from "./scopes.js";
import { describeChat, hostedImageUrls, resolveChat } from "./chats.js";

const jwt = (payload) =>
  `x.${Buffer.from(JSON.stringify(payload)).toString("base64url")}.y`;

test("scopesFromToken reads the scp claim", () => {
  const s = scopesFromToken(jwt({ scp: "Mail.Read Chat.Read" }));
  assert.deepEqual([...s].sort(), ["Chat.Read", "Mail.Read"]);
});

test("scopesFromToken grants nothing for a token it cannot read", () => {
  assert.equal(scopesFromToken("not-a-jwt").size, 0);
  assert.equal(scopesFromToken(undefined).size, 0);
  assert.equal(scopesFromToken(jwt({})).size, 0);
});

test("partitionTools withholds a tool whose scope is missing, and says which", () => {
  const tools = [
    { name: "mail_list", scopes: ["Mail.Read"] },
    { name: "channel_list", scopes: ["Channel.ReadBasic.All"] },
    { name: "graph_scopes", scopes: [] },
  ];
  const { available, withheld } = partitionTools(tools, new Set(["Mail.Read"]));

  assert.deepEqual(available.map((t) => t.name), ["mail_list", "graph_scopes"]);
  assert.deepEqual(withheld, [{ name: "channel_list", reason: "missing scope: Channel.ReadBasic.All" }]);
});

test("read-only mode withholds a write tool the token is otherwise allowed to use", () => {
  const tools = [
    { name: "mail_list", scopes: ["Mail.Read"] },
    { name: "mail_send", scopes: ["Mail.Send"], writes: true },
  ];
  const granted = new Set(["Mail.Read", "Mail.Send"]);

  assert.deepEqual(partitionTools(tools, granted).available.map((t) => t.name), ["mail_list", "mail_send"]);
  assert.deepEqual(partitionTools(tools, granted, { readOnly: true }).withheld, [
    { name: "mail_send", reason: "read-only mode" },
  ]);
});

test("describeChat names a group by its topic and a DM by the other person", () => {
  const group = describeChat({ id: "1", chatType: "group", topic: "React Devs", members: [{ displayName: "A" }] });
  assert.equal(group.name, "React Devs");

  const dm = describeChat(
    { id: "2", chatType: "oneOnOne", members: [{ displayName: "Me", userId: "me" }, { displayName: "Standa", userId: "s" }] },
    "me"
  );
  assert.equal(dm.name, "Standa");
});

test("resolveChat refuses to guess between several matches", () => {
  const chats = [
    { id: "a", type: "group", name: "Martin and friends", topic: "Martin and friends", members: [] },
    { id: "b", type: "group", name: "Martin planning", topic: "Martin planning", members: [] },
  ];
  assert.throws(() => resolveChat(chats, "Martin"), /matches 2 chats/);
  assert.throws(() => resolveChat(chats, "nobody"), /No chat matching/);
  assert.equal(resolveChat(chats, "planning").id, "b");
});

test("resolveChat prefers the single one-to-one when a name also hits groups", () => {
  const chats = [
    { id: "dm", type: "oneOnOne", name: "Karel Petrák", members: ["Karel Petrák"] },
    { id: "g", type: "group", name: "Karel and the team", topic: "Karel and the team", members: [] },
  ];
  assert.equal(resolveChat(chats, "Karel").id, "dm");
});

test("hostedImageUrls pulls inline images out of a Teams body", () => {
  const body =
    '<p>FYI:</p><p><img src="https://graph.microsoft.com/v1.0/chats/19:x/messages/1/hostedContents/abc/$value" alt="image"></p>';
  assert.deepEqual(hostedImageUrls(body), ["https://graph.microsoft.com/v1.0/chats/19:x/messages/1/hostedContents/abc/$value"]);
  assert.deepEqual(hostedImageUrls("<p>no image</p>"), []);
  assert.deepEqual(hostedImageUrls(undefined), []);
});
