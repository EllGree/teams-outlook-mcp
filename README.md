# teams-outlook-mcp

MCP server for Microsoft Teams and Outlook over the Graph API.

**The token decides what the server offers.** The tool list is derived at startup from the `scp` claim of the access token: a permission that was never consented to produces no tool, and one granted later produces one on the next restart, with no code change. `graph_scopes` reports what was granted and why anything is missing.

## Tools

| Tool | Needs |
|---|---|
| `graph_scopes` | nothing |
| `graph_me` | `User.Read` |
| `mail_list`, `mail_get`, `mail_attachment` | `Mail.Read` |
| `mail_send` | `Mail.Send` |
| `chat_list`, `chat_messages`, `chat_image` | `Chat.Read` |
| `chat_send` | `ChatMessage.Send` |
| `team_list` | `Team.ReadBasic.All` |
| `channel_list` | `Channel.ReadBasic.All` |
| `channel_messages` | `ChannelMessage.Read.All` |
| `channel_send` | `ChannelMessage.Send` |

Images come back as images rather than links, from mail attachments and from the inline images in Teams messages alike, so a screenshot someone posted can actually be read.

Chats are addressed by `chat_id` or by `chat_name`: a group chat by its topic, a one-to-one by the other person. A name matching several chats refuses rather than guesses, because a send cannot be taken back.

Every outbound message carries the assistant signature.

## Install

Needs Node 18 or newer.

```
npm install
node auth.js
```

`auth.js` is a device-code sign-in: it prints a URL and a code, waits, then caches the tokens and prints the scopes that were actually granted.

A scope needing an administrator's approval blocks the whole consent screen, not only itself, so a run can leave one out and come back for it once an admin has granted it:

```
node auth.js --without Channel.ReadBasic.All
```

Register the server with your MCP client as stdio: command `node`, argument the full path to `index.js`.

## Config

`config.json` next to `index.js`, or the matching environment variables. A value left at the example placeholder counts as absent, so a half-filled `config.json` falls through to the environment instead of sending a placeholder to Entra.

| Key | Environment | Meaning |
|---|---|---|
| `tenantId` | `TEAMS_OUTLOOK_MCP_TENANT_ID` | Your Entra tenant. Required. |
| `clientId` | `TEAMS_OUTLOOK_MCP_CLIENT_ID` | The app registration to sign in against. Required. |
| `signature` | `TEAMS_OUTLOOK_MCP_SIGNATURE` | Appended to every message sent. Defaults to a generic one rather than to nothing, so a reader can always tell a program wrote it. |
| `tokenPath` | `TEAMS_OUTLOOK_MCP_TOKENS` | Where the token cache lives. Defaults to `%LOCALAPPDATA%\teams-outlook-mcp\tokens.json`. |
| `readOnly` | `TEAMS_OUTLOOK_MCP_READONLY` | Withholds every sending tool whatever the token allows. |

`config.json` is gitignored. The token cache is deliberately this server's own: another tool refreshing the same cache with a narrower scope list would quietly cost this one its permissions.

## Notes

- A scope consented after the cache was written never arrives through a refresh. Azure answers `AADSTS65001`, and signing in again is the way through; the error says so.
- Inline images in Teams are not attachments. They live in the message body as absolute Graph `hostedContents` URLs and `message.attachments` stays empty. Mail is the other shape: inline images are attachments with `isInline` and a `contentId` that the body references as `cid:`.
- Graph refuses `$filter` and `$search` in one request, so `mail_list` drops the filters when a search is given.
- Attachments over 4 MB are described rather than inlined.

## Tests

```
npm test
```

Covers the scope gating and the chat name resolution, which are the parts where being wrong is either invisible or expensive.
