#!/usr/bin/env node
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";

import { getAccessToken, graphGet, graphGetBinary, graphPost } from "./lib/graph.js";
import { partitionTools, scopesFromToken } from "./lib/scopes.js";
import { describeChat, hostedImageUrls, listChats, resolveChat } from "./lib/chats.js";
import { appendSignatureHtml, appendSignatureText, htmlToText, textToHtml } from "./lib/html.js";
import { READ_ONLY } from "./lib/config.js";


// An oversized base64 blob costs the caller its context and helps nobody, so past this a file is
// handed over as a URL instead of bytes.
const MAX_INLINE_BYTES = 4 * 1024 * 1024;

let ME = null;
async function me(token) {
  if (!ME) ME = await graphGet(token, "/me?$select=id,displayName,mail,userPrincipalName");
  return ME;
}

function imageOrText(name, buffer, mediaType, fallbackUrl) {
  if (buffer.length > MAX_INLINE_BYTES) {
    return { content: [{ type: "text", text: `${name} is ${buffer.length} bytes, too large to inline. ${fallbackUrl ?? ""}`.trim() }] };
  }
  if (mediaType.startsWith("image/")) {
    return {
      content: [
        { type: "text", text: `${name} (${mediaType}, ${buffer.length} bytes)` },
        { type: "image", data: buffer.toString("base64"), mimeType: mediaType },
      ],
    };
  }
  if (/^text\/|json|xml|yaml|csv/i.test(mediaType)) {
    return { content: [{ type: "text", text: `${name} (${mediaType}):\n\n${buffer.toString("utf-8")}` }] };
  }
  return { content: [{ type: "text", text: `${name} is ${mediaType}, which this server does not inline. ${fallbackUrl ?? ""}`.trim() }] };
}

// ---- mail ----

function summariseMessage(m) {
  return {
    id: m.id,
    received: m.receivedDateTime,
    from: m.from?.emailAddress?.address,
    from_name: m.from?.emailAddress?.name,
    subject: m.subject,
    unread: m.isRead === false,
    has_attachments: m.hasAttachments,
    preview: m.bodyPreview,
  };
}

async function toolMailList(token, { top = 15, unread_only = false, from, search, since }) {
  const params = ["$top=" + Math.min(Number(top) || 15, 50), "$select=id,receivedDateTime,from,subject,isRead,hasAttachments,bodyPreview"];
  const filters = [];
  if (unread_only) filters.push("isRead eq false");
  if (from) filters.push(`from/emailAddress/address eq '${String(from).replace(/'/g, "''")}'`);
  if (since) filters.push(`receivedDateTime ge ${new Date(since).toISOString()}`);
  // Graph refuses $filter and $search in the same request, so search wins and filters are dropped.
  if (search) params.push(`$search="${String(search).replace(/"/g, "")}"`);
  else {
    if (filters.length) params.push("$filter=" + encodeURIComponent(filters.join(" and ")));
    params.push("$orderby=receivedDateTime desc");
  }
  const r = await graphGet(token, `/me/messages?${params.join("&")}`);
  return { total: (r.value || []).length, results: (r.value || []).map(summariseMessage) };
}

async function toolMailGet(token, { message_id }) {
  if (!message_id) throw new Error("message_id is required");
  const m = await graphGet(token, `/me/messages/${encodeURIComponent(message_id)}`);
  const out = {
    ...summariseMessage(m),
    to: (m.toRecipients || []).map((r) => r.emailAddress?.address),
    cc: (m.ccRecipients || []).map((r) => r.emailAddress?.address),
    body: htmlToText(m.body?.content),
  };
  if (m.hasAttachments) {
    const a = await graphGet(token, `/me/messages/${encodeURIComponent(message_id)}/attachments?$select=id,name,contentType,size,isInline`);
    out.attachments = (a.value || []).map((it) => ({
      id: it.id,
      name: it.name,
      content_type: it.contentType,
      size: it.size,
      inline: it.isInline === true,
    }));
  }
  return out;
}

async function toolMailAttachment(token, { message_id, attachment_id, name }) {
  if (!message_id) throw new Error("message_id is required");
  if (!attachment_id && !name) throw new Error("provide attachment_id or name");
  const a = await graphGet(token, `/me/messages/${encodeURIComponent(message_id)}/attachments`);
  const list = a.value || [];
  const hit = attachment_id ? list.find((it) => it.id === attachment_id) : list.find((it) => it.name === name);
  if (!hit) throw new Error(`No such attachment. On this message: ${list.map((it) => it.name).join(", ") || "none"}`);
  if (!hit.contentBytes) throw new Error(`${hit.name} is a ${hit["@odata.type"]}, which carries no bytes to inline`);
  return imageOrText(hit.name, Buffer.from(hit.contentBytes, "base64"), hit.contentType || "application/octet-stream");
}

async function toolMailSend(token, { to, cc, subject, body }) {
  if (!to) throw new Error("to is required");
  if (!subject) throw new Error("subject is required");
  if (!body) throw new Error("body is required");
  const addresses = (v) =>
    (Array.isArray(v) ? v : String(v).split(",")).map((s) => String(s).trim()).filter(Boolean).map((address) => ({ emailAddress: { address } }));
  await graphPost(token, "/me/sendMail", {
    message: {
      subject,
      body: { contentType: "Text", content: appendSignatureText(body) },
      toRecipients: addresses(to),
      ccRecipients: cc ? addresses(cc) : [],
    },
    saveToSentItems: true,
  });
  return { sent: true, to, subject };
}

// ---- chats ----

async function toolChatList(token, { top = 50 }) {
  const meRes = await me(token);
  return { results: await listChats(token, { top, meId: meRes.id }) };
}

async function chatIdFrom(token, { chat_id, chat_name }) {
  if (chat_id) return chat_id;
  if (!chat_name) throw new Error("provide chat_id or chat_name");
  const meRes = await me(token);
  return resolveChat(await listChats(token, { meId: meRes.id }), chat_name).id;
}

async function toolChatMessages(token, { chat_id, chat_name, top = 20 }) {
  const id = await chatIdFrom(token, { chat_id, chat_name });
  const r = await graphGet(token, `/chats/${encodeURIComponent(id)}/messages?$top=${Math.min(Number(top) || 20, 50)}`);
  return {
    chat_id: id,
    results: (r.value || [])
      .map((m) => ({
        id: m.id,
        sent: m.createdDateTime,
        from: m.from?.user?.displayName || m.from?.application?.displayName,
        text: htmlToText(m.body?.content),
        images: hostedImageUrls(m.body?.content).length,
        attachments: (m.attachments || []).map((a) => a.name).filter(Boolean),
      }))
      .filter((m) => m.text || m.images || m.attachments.length),
  };
}

async function toolChatImage(token, { chat_id, chat_name, message_id, index = 0 }) {
  if (!message_id) throw new Error("message_id is required");
  const id = await chatIdFrom(token, { chat_id, chat_name });
  const m = await graphGet(token, `/chats/${encodeURIComponent(id)}/messages/${encodeURIComponent(message_id)}`);
  const urls = hostedImageUrls(m.body?.content);
  if (!urls.length) throw new Error("That message carries no inline image");
  const url = urls[Number(index) || 0];
  if (!url) throw new Error(`Message has ${urls.length} image(s); index ${index} is out of range`);
  const { buffer, mediaType } = await graphGetBinary(token, url);
  return imageOrText(`image ${Number(index) + 1} of ${urls.length}`, buffer, mediaType, url);
}

async function toolChatSend(token, { chat_id, chat_name, text }) {
  if (!text) throw new Error("text is required");
  const id = await chatIdFrom(token, { chat_id, chat_name });
  const res = await graphPost(token, `/chats/${encodeURIComponent(id)}/messages`, {
    body: { contentType: "html", content: appendSignatureHtml(textToHtml(text)) },
  });
  return { sent: true, chat_id: id, message_id: res?.id };
}

// ---- teams and channels ----

async function toolTeamList(token) {
  const r = await graphGet(token, "/me/joinedTeams?$select=id,displayName,description");
  return { results: (r.value || []).map((t) => ({ id: t.id, name: t.displayName, description: t.description })) };
}

async function toolChannelList(token, { team_id }) {
  if (!team_id) throw new Error("team_id is required");
  const r = await graphGet(token, `/teams/${encodeURIComponent(team_id)}/channels`);
  return { results: (r.value || []).map((c) => ({ id: c.id, name: c.displayName, description: c.description })) };
}

async function toolChannelMessages(token, { team_id, channel_id, top = 20 }) {
  if (!team_id || !channel_id) throw new Error("team_id and channel_id are required");
  const r = await graphGet(
    token,
    `/teams/${encodeURIComponent(team_id)}/channels/${encodeURIComponent(channel_id)}/messages?$top=${Math.min(Number(top) || 20, 50)}`
  );
  return {
    results: (r.value || []).map((m) => ({
      id: m.id,
      sent: m.createdDateTime,
      from: m.from?.user?.displayName,
      subject: m.subject || undefined,
      text: htmlToText(m.body?.content),
    })),
  };
}

async function toolChannelSend(token, { team_id, channel_id, text }) {
  if (!team_id || !channel_id) throw new Error("team_id and channel_id are required");
  if (!text) throw new Error("text is required");
  const res = await graphPost(token, `/teams/${encodeURIComponent(team_id)}/channels/${encodeURIComponent(channel_id)}/messages`, {
    body: { contentType: "html", content: appendSignatureHtml(textToHtml(text)) },
  });
  return { sent: true, message_id: res?.id };
}

// ---- tool table ----
// `scopes` is what the token must carry for the tool to be offered at all, and `writes` marks the
// ones read-only mode withholds whatever the token says.

const TOOLS = [
  {
    name: "graph_scopes",
    scopes: [],
    description: "What this token is allowed to do, and which tools are therefore unavailable.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "graph_me",
    scopes: ["User.Read"],
    description: "The signed-in user.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "mail_list",
    scopes: ["Mail.Read"],
    description: "Recent Outlook messages. Filter by unread, sender or date, or pass search (search cannot be combined with the filters).",
    inputSchema: {
      type: "object",
      properties: {
        top: { type: "number", default: 15 },
        unread_only: { type: "boolean", default: false },
        from: { type: "string", description: "Exact sender address" },
        since: { type: "string", description: "ISO date; messages received at or after it" },
        search: { type: "string", description: "Free-text search over the mailbox" },
      },
    },
  },
  {
    name: "mail_get",
    scopes: ["Mail.Read"],
    description: "One message as readable text, with its recipients and its attachment list.",
    inputSchema: { type: "object", properties: { message_id: { type: "string" } }, required: ["message_id"] },
  },
  {
    name: "mail_attachment",
    scopes: ["Mail.Read"],
    description: "Fetch a mail attachment by id or by name. Images come back viewable, text as text.",
    inputSchema: {
      type: "object",
      properties: { message_id: { type: "string" }, attachment_id: { type: "string" }, name: { type: "string" } },
      required: ["message_id"],
    },
  },
  {
    name: "mail_send",
    scopes: ["Mail.Send"],
    writes: true,
    description: "Send mail as the signed-in user. The assistant signature is appended.",
    inputSchema: {
      type: "object",
      properties: {
        to: { type: "string", description: "One address or several, comma separated" },
        cc: { type: "string" },
        subject: { type: "string" },
        body: { type: "string" },
      },
      required: ["to", "subject", "body"],
    },
  },
  {
    name: "chat_list",
    scopes: ["Chat.Read"],
    description: "Teams chats, one-to-one and group, each with a usable name: the other person, or the group topic.",
    inputSchema: { type: "object", properties: { top: { type: "number", default: 50 } } },
  },
  {
    name: "chat_messages",
    scopes: ["Chat.Read"],
    description: "Messages in a chat, addressed by chat_id or by chat_name. An ambiguous name refuses rather than guesses.",
    inputSchema: {
      type: "object",
      properties: { chat_id: { type: "string" }, chat_name: { type: "string" }, top: { type: "number", default: 20 } },
    },
  },
  {
    name: "chat_image",
    scopes: ["Chat.Read"],
    description: "The inline image in a chat message, returned as a viewable image. Use index when a message carries several.",
    inputSchema: {
      type: "object",
      properties: {
        chat_id: { type: "string" },
        chat_name: { type: "string" },
        message_id: { type: "string" },
        index: { type: "number", default: 0 },
      },
      required: ["message_id"],
    },
  },
  {
    name: "chat_send",
    scopes: ["ChatMessage.Send"],
    writes: true,
    description: "Post into a chat, addressed by chat_id or chat_name. The assistant signature is appended.",
    inputSchema: {
      type: "object",
      properties: { chat_id: { type: "string" }, chat_name: { type: "string" }, text: { type: "string" } },
      required: ["text"],
    },
  },
  {
    name: "team_list",
    scopes: ["Team.ReadBasic.All"],
    description: "Teams the user has joined.",
    inputSchema: { type: "object", properties: {} },
  },
  {
    name: "channel_list",
    scopes: ["Channel.ReadBasic.All"],
    description: "Channels in a team.",
    inputSchema: { type: "object", properties: { team_id: { type: "string" } }, required: ["team_id"] },
  },
  {
    name: "channel_messages",
    scopes: ["ChannelMessage.Read.All"],
    description: "Messages posted in a channel.",
    inputSchema: {
      type: "object",
      properties: { team_id: { type: "string" }, channel_id: { type: "string" }, top: { type: "number", default: 20 } },
      required: ["team_id", "channel_id"],
    },
  },
  {
    name: "channel_send",
    scopes: ["ChannelMessage.Send"],
    writes: true,
    description: "Post into a channel. The assistant signature is appended.",
    inputSchema: {
      type: "object",
      properties: { team_id: { type: "string" }, channel_id: { type: "string" }, text: { type: "string" } },
      required: ["team_id", "channel_id", "text"],
    },
  },
];

const HANDLERS = {
  graph_me: (token) => me(token),
  mail_list: toolMailList,
  mail_get: toolMailGet,
  mail_attachment: toolMailAttachment,
  mail_send: toolMailSend,
  chat_list: toolChatList,
  chat_messages: toolChatMessages,
  chat_image: toolChatImage,
  chat_send: toolChatSend,
  team_list: toolTeamList,
  channel_list: toolChannelList,
  channel_messages: toolChannelMessages,
  channel_send: toolChannelSend,
};

// ---- wiring ----

const token = await getAccessToken();
const granted = scopesFromToken(token);
const { available, withheld } = partitionTools(TOOLS, granted, { readOnly: READ_ONLY });

// Stderr, so it reaches the client's log without disturbing the protocol on stdout.
console.error(`teams-outlook-mcp: ${available.length} tools available, ${withheld.length} withheld`);
for (const w of withheld) console.error(`  ${w.name}: ${w.reason}`);

const listed = available.map(({ name, description, inputSchema }) => ({ name, description, inputSchema }));

const server = new Server({ name: "teams-outlook-mcp", version: "0.1.0" }, { capabilities: { tools: {} } });

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: listed }));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  if (!available.some((t) => t.name === name)) {
    const reason = withheld.find((w) => w.name === name)?.reason;
    return {
      isError: true,
      content: [{ type: "text", text: reason ? `${name} is not available: ${reason}` : `Unknown tool: ${name}` }],
    };
  }

  try {
    if (name === "graph_scopes") {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              { read_only: READ_ONLY, granted: [...granted].sort(), available: available.map((t) => t.name), withheld },
              null,
              2
            ),
          },
        ],
      };
    }

    // Refreshed per call: a long-running server outlives its access token.
    const fresh = await getAccessToken();
    const result = await HANDLERS[name](fresh, args || {});
    if (result && Array.isArray(result.content)) return result;
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  } catch (err) {
    return { isError: true, content: [{ type: "text", text: `Error: ${err.message || String(err)}` }] };
  }
});

await server.connect(new StdioServerTransport());
